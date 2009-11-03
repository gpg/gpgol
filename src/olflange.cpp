/* olflange.cpp - Connect GpgOL to Outlook
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2004, 2005, 2007, 2008 g10 Code GmbH
 * 
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>

#ifndef INITGUID
#define INITGUID
#endif

#include <initguid.h>
#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"

#include "common.h"
#include "display.h"
#include "msgcache.h"
#include "engine.h"
#include "mapihelp.h"

#include "olflange-def.h"
#include "olflange.h"
#include "ext-commands.h"
#include "user-events.h"
#include "session-events.h"
#include "message-events.h"
#include "property-sheets.h"
#include "attached-file-events.h"
#include "item-events.h"
#include "explorers.h"
#include "inspectors.h"
#include "mailitem.h"
#include "cmdbarcontrols.h"

/* The GUID for this plugin.  */
#define CLSIDSTR_GPGOL   "{42d30988-1a3a-11da-c687-000d6080e735}"
DEFINE_GUID(CLSID_GPGOL, 0x42d30988, 0x1a3a, 0x11da, 
            0xc6, 0x87, 0x00, 0x0d, 0x60, 0x80, 0xe7, 0x35);

/* For documentation: The GUID used for our custom properties: 
   {31805ab8-3e92-11dc-879c-00061b031004}
 */


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


static bool g_initdll = FALSE;

static void install_forms (void);
static void install_sinks (LPEXCHEXTCALLBACK eecb);


static char *olversion;



/* Return a string for the context NO.  This never return NULL. */
const char *
ext_context_name (unsigned long no)
{
  switch (no)
    {
    case EECONTEXT_SESSION:           return "Session";
    case EECONTEXT_VIEWER:            return "Viewer";
    case EECONTEXT_REMOTEVIEWER:      return "RemoteViewer";
    case EECONTEXT_SEARCHVIEWER:      return "SearchViewer";
    case EECONTEXT_ADDRBOOK:          return "AddrBook";
    case EECONTEXT_SENDNOTEMESSAGE:   return "SendNoteMessage";
    case EECONTEXT_READNOTEMESSAGE:   return "ReadNoteMessage";
    case EECONTEXT_SENDPOSTMESSAGE:   return "SendPostMessage";
    case EECONTEXT_READPOSTMESSAGE:   return "ReadPostMessage";
    case EECONTEXT_READREPORTMESSAGE: return "ReadReportMessage";
    case EECONTEXT_SENDRESENDMESSAGE: return "SendResendMessage";
    case EECONTEXT_PROPERTYSHEETS:    return "PropertySheets";
    case EECONTEXT_ADVANCEDCRITERIA:  return "AdvancedCriteria";
    case EECONTEXT_TASK:              return "Task";
    default: return "?";
    }
}


EXTERN_C int
get_ol_main_version (void)
{
  return olversion? atoi (olversion): 0;
}


/* Wrapper around UlRelease with error checking. */
// static void 
// ul_release (LPVOID punk, const char *func)
// {
//   ULONG res;
  
//   if (!punk)
//     return;
//   res = UlRelease (punk);
//   if (opt.enable_debug & DBG_MEMORY)
//     log_debug ("%s:%s: UlRelease(%p) had %lu references\n", 
//                SRCNAME, func, punk, res);
// }



/* Registers this module as an Exchange extension. This basically updates
   some Registry entries. */
STDAPI 
DllRegisterServer (void)
{    
  HKEY hkey, hkey2;
  CHAR szKeyBuf[MAX_PATH+1024];
  CHAR szEntry[MAX_PATH+512];
  TCHAR szModuleFileName[MAX_PATH];
  DWORD dwTemp = 0;
  long ec;

  /* Get server location. */
  if (!GetModuleFileName(glob_hinst, szModuleFileName, MAX_PATH))
    return E_FAIL;

  lstrcpy (szKeyBuf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
  lstrcpy (szEntry, "4.0;");
  lstrcat (szEntry, szModuleFileName);
  lstrcat (szEntry, ";1"); /* Entry point ordinal. */
  /* Context information string:
     pos       context
     1 	EECONTEXT_SESSION
     2 	EECONTEXT_VIEWER
     3 	EECONTEXT_REMOTEVIEWER
     4 	EECONTEXT_SEARCHVIEWER
     5 	EECONTEXT_ADDRBOOK
     6 	EECONTEXT_SENDNOTEMESSAGE
     7 	EECONTEXT_READNOTEMESSAGE
     8 	EECONTEXT_SENDPOSTMESSAGE
     9 	EECONTEXT_READPOSTMESSAGE
     a  EECONTEXT_READREPORTMESSAGE
     b  EECONTEXT_SENDRESENDMESSAGE
     c  EECONTEXT_PROPERTYSHEETS
     d  EECONTEXT_ADVANCEDCRITERIA
     e  EECONTEXT_TASK
                   ___123456789abcde___ */                 
  lstrcat (szEntry, ";11000111111100"); 
  /* Interfaces to we want to hook into:
     pos  interface
     1    IExchExtCommands            
     2    IExchExtUserEvents          
     3    IExchExtSessionEvents       
     4    IExchExtMessageEvents       
     5    IExchExtAttachedFileEvents  
     6    IExchExtPropertySheets      
     7    IExchExtAdvancedCriteria    
     -    IExchExt              
     -    IExchExtModeless
     -    IExchExtModelessCallback
                   ___1234567___ */
  lstrcat (szEntry, ";11111101"); 
  ec = RegCreateKeyEx (HKEY_LOCAL_MACHINE, szKeyBuf, 0, NULL, 
                       REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      log_debug ("DllRegisterServer failed\n");
      return E_ACCESSDENIED;
    }
    
  dwTemp = lstrlen (szEntry) + 1;
  RegSetValueEx (hkey, "GpgOL", 0, REG_SZ, (BYTE*) szEntry, dwTemp);

  /* To avoid conflicts with the old G-DATA plugin and older versions
     of this Plugin, we remove the key used by these versions. */
  RegDeleteValue (hkey, "GPG Exchange");

  /* Set outlook update flag. */
  /* Fixme: We have not yet implemented this hint from Microsoft:

       In order for .ecf-based ECEs to be detected by Outlook on Vista,
       one needs to delete if present:

       [HKEY_CURRENT_USER\Software\Microsoft\Office\[Office Version]\Outlook]
       "Exchange Client Extension"=
             "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX"

       [Office Version] is 11.0 ( for OL 03 ), 12.0 ( OL 07 )...

     Obviously due to HKCU, that also requires to run this code at
     startup.  However, we don't use an ECF right now.  */
  strcpy (szEntry, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
  dwTemp = lstrlen (szEntry) + 1;
  RegSetValueEx (hkey, "Outlook Setup Extension",
                 0, REG_SZ, (BYTE*) szEntry, dwTemp);
  RegCloseKey (hkey);
    
  hkey = NULL;
  lstrcpy (szKeyBuf, "Software\\GNU\\GpgOL");
  RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (hkey != NULL)
    RegCloseKey (hkey);

  hkey = NULL;
  strcpy (szKeyBuf, "CLSID\\" CLSIDSTR_GPGOL );
  ec = RegCreateKeyEx (HKEY_CLASSES_ROOT, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      return E_ACCESSDENIED;
    }

  strcpy (szEntry, "GpgOL - The GnuPG Outlook Plugin");
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  strcpy (szKeyBuf, "InprocServer32");
  ec = RegCreateKeyEx (hkey, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey2, NULL);
  if (ec != ERROR_SUCCESS) 
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      RegCloseKey (hkey);
      return E_ACCESSDENIED;
    }
  strcpy (szEntry, szModuleFileName);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  strcpy (szEntry, "Neutral");
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, "ThreadingModel", 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  RegCloseKey (hkey2);
  RegCloseKey (hkey);


  log_debug ("DllRegisterServer succeeded\n");
  return S_OK;
}


/* Unregisters this module as an Exchange extension. */
STDAPI 
DllUnregisterServer (void)
{
  HKEY hkey;
  CHAR buf[MAX_PATH+1024];
  DWORD ntemp;
  long res;

  strcpy (buf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
  /* Create and open key and subkey. */
  res = RegCreateKeyEx (HKEY_LOCAL_MACHINE, buf, 0, NULL, 
			REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 
			NULL, &hkey, NULL);
  if (res != ERROR_SUCCESS) 
    {
      log_debug ("DllUnregisterServer: access denied.\n");
      return E_ACCESSDENIED;
    }
  RegDeleteValue (hkey, "GpgOL");
  
  /* Set outlook update flag.  */  
  strcpy (buf, "4.0;Outxxx.dll;7;000000000000000;0000000000;OutXXX");
  ntemp = strlen (buf) + 1;
  RegSetValueEx (hkey, "Outlook Setup Extension", 0, 
		 REG_SZ, (BYTE*) buf, ntemp);
  RegCloseKey (hkey);

  /* Delete CLSIDs. */
  strcpy (buf, "CLSID\\" CLSIDSTR_GPGOL "\\InprocServer32");
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);
  strcpy (buf, "CLSID\\" CLSIDSTR_GPGOL);
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);
  
  return S_OK;
}


static const char*
parse_version_number (const char *s, int *number)
{
  int val = 0;

  if (*s == '0' && digitp (s+1))
    return NULL;  /* Leading zeros are not allowed.  */
  for (; digitp (s); s++)
    {
      val *= 10;
      val += *s - '0';
    }
  *number = val;
  return val < 0 ? NULL : s;
}

static const char *
parse_version_string (const char *s, int *major, int *minor, int *micro)
{
  s = parse_version_number (s, major);
  if (!s || *s != '.')
    return NULL;
  s++;
  s = parse_version_number (s, minor);
  if (!s || *s != '.')
    return NULL;
  s++;
  s = parse_version_number (s, micro);
  if (!s)
    return NULL;
  return s;  /* Patchlevel.  */
}

static const char *
compare_versions (const char *my_version, const char *req_version)
{
  int my_major, my_minor, my_micro;
  int rq_major, rq_minor, rq_micro;
  const char *my_plvl, *rq_plvl;

  if (!req_version)
    return my_version;
  if (!my_version)
    return NULL;

  my_plvl = parse_version_string (my_version, &my_major, &my_minor, &my_micro);
  if (!my_plvl)
    return NULL;	/* Very strange: our own version is bogus.  */
  rq_plvl = parse_version_string(req_version,
				 &rq_major, &rq_minor, &rq_micro);
  if (!rq_plvl)
    return NULL;	/* Requested version string is invalid.  */

  if (my_major > rq_major
	|| (my_major == rq_major && my_minor > rq_minor)
      || (my_major == rq_major && my_minor == rq_minor 
	  && my_micro > rq_micro)
      || (my_major == rq_major && my_minor == rq_minor
	  && my_micro == rq_micro
	  && strcmp( my_plvl, rq_plvl ) >= 0))
    {
      return my_version;
    }
  return NULL;
}

/* Check that the the version of GpgOL is at minimum the requested one
 * and return GpgOL's version string; return NULL if that condition is
 * not met.  If a NULL is passed to this function, no check is done
 * and the version string is simply returned.  */
EXTERN_C const char * __stdcall
gpgol_check_version (const char *req_version)
{
  return compare_versions (PACKAGE_VERSION, req_version);
}




/* The entry point which Exchange/Outlook calls.  This is called for
   each context entry.  Creates a new GpgolExt object every time so
   each context will get its own GpgolExt interface. */
EXTERN_C LPEXCHEXT __stdcall
ExchEntryPoint (void)
{
  log_debug ("%s:%s: creating new GpgolExt object\n", SRCNAME, __func__);
  return new GpgolExt;
}



/* Constructor of GpgolExt

   Initializes members and creates the interface objects for the new
   context.  Does the DLL initialization if it has not been done
   before. */
GpgolExt::GpgolExt (void)
{ 
  m_lRef = 1;
  m_lContext = 0;
  m_hWndExchange = 0;

  m_pExchExtCommands           = new GpgolExtCommands (this);
  m_pExchExtUserEvents         = new GpgolUserEvents (this);
  m_pExchExtSessionEvents      = new GpgolSessionEvents (this);
  m_pExchExtMessageEvents      = new GpgolMessageEvents (this);
  m_pExchExtAttachedFileEvents = new GpgolAttachedFileEvents (this);
  m_pExchExtPropertySheets     = new GpgolPropertySheets (this);
//   m_pOutlookExtItemEvents      = new GpgolItemEvents (this);
  if (!m_pExchExtCommands
      || !m_pExchExtUserEvents
      || !m_pExchExtSessionEvents
      || !m_pExchExtMessageEvents
      || !m_pExchExtAttachedFileEvents
      || !m_pExchExtPropertySheets
      /*|| !m_pOutlookExtItemEvents*/)
    out_of_core ();

  /* For this class we need to bump the reference counter intially.
     The question is why it works at all with the other stuff.  */
//   m_pOutlookExtItemEvents->AddRef ();

  if (!g_initdll)
    {
      read_options ();
      log_debug ("%s:%s: this is GpgOL %s\n", 
                 SRCNAME, __func__, PACKAGE_VERSION);
      log_debug ("%s:%s:   using GPGME %s\n", 
                 SRCNAME, __func__, gpgme_check_version (NULL));
      engine_init ();
      g_initdll = TRUE;
      log_debug ("%s:%s: first time initialization done\n",
                 SRCNAME, __func__);

#define ANNOUNCE_NUMBER 1
      if ( ANNOUNCE_NUMBER > opt.announce_number )
        {
          /* Note: If you want to change the announcment, you need to
             increment the ANNOUNCE_NUMBER above.  The number assures
             that a user will see this message only once.  */
          MessageBox 
            (NULL,
             _("Welcome to GpgOL 1.0\n"
               "\n"
               "GpgOL adds integrated OpenPGP and S/MIME encryption "
               "and digital signing support to Outlook 2003 and 2007.\n"
               "\n"
               "Although we tested this software extensively, we can't "
               "give you any guarantee that it will work as expected. "
               "The programming interface we are using has not been properly "
               "documented by Microsoft and thus the functionality of GpgOL "
               "may cease to work with an update of your Windows system.\n"
               "\n"
               "WE STRONGLY ADVISE TO RUN ENCRYPTION TESTS BEFORE YOU START "
               "TO USE GPGOL ON ANY SENSITIVE DATA!\n"
               "\n"
               "There are some known problems, the most severe being "
               "that sending encrypted or signed mails using an Exchange "
               "based account does not work.  Using GpgOL along with "
               "other Outlook plugins may in some cases not work."
               "\n"),     
             "GpgOL", MB_ICONINFORMATION|MB_OK);
          /* Show this warning only once.  */
          opt.announce_number = ANNOUNCE_NUMBER;
          write_options ();
        }

      if ( SVN_REVISION > opt.svn_revision )
        {
          MessageBox (NULL,
                    _("You have installed a new version of GpgOL.\n"
                      "\n"
                      "Please open the option dialog and confirm that"
                      " the settings are correct for your needs.  The option"
                      " dialog can be found in the main menu at:"
                      " Extras->Options->GpgOL.\n"),
                      "GpgOL", MB_ICONINFORMATION|MB_OK);
          /* Show this warning only once.  */
          opt.svn_revision = SVN_REVISION;
          write_options ();
        }
      if ( SVN_REVISION > opt.forms_revision )
        install_forms ();
    }
}


/*  Uninitializes the DLL in the session context. */
GpgolExt::~GpgolExt (void) 
{
  log_debug ("%s:%s: cleaning up GpgolExt object; context=%s\n",
             SRCNAME, __func__, ext_context_name (m_lContext));
    
//   if (m_pOutlookExtItemEvents)
//     m_pOutlookExtItemEvents->Release ();
  
  if (m_lContext == EECONTEXT_SESSION)
    {
      if (g_initdll)
        {
          engine_deinit ();
          write_options ();
          g_initdll = FALSE;
          log_debug ("%s:%s: DLL closed down\n", SRCNAME, __func__);
	}	
    }
  

}


/* Called by Exchange to retrieve an object pointer for a named
   interface.  This is a standard COM method.  REFIID is the ID of the
   interface and PPVOBJ will get the address of the object pointer if
   this class defines the requested interface.  Return value: S_OK if
   the interface is supported, otherwise E_NOINTERFACE. */
STDMETHODIMP 
GpgolExt::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
  HRESULT hr = S_OK;
  
  *ppvObj = NULL;
  
  if ((riid == IID_IUnknown) || (riid == IID_IExchExt)) 
    {
      *ppvObj = (LPUNKNOWN) this;
    }
  else if (riid == IID_IExchExtCommands) 
    {
      *ppvObj = (LPUNKNOWN)m_pExchExtCommands;
      m_pExchExtCommands->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtUserEvents) 
    {
      *ppvObj = (LPUNKNOWN) m_pExchExtUserEvents;
      m_pExchExtUserEvents->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtSessionEvents) 
    {
      *ppvObj = (LPUNKNOWN) m_pExchExtSessionEvents;
      m_pExchExtSessionEvents->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtMessageEvents) 
    {
      *ppvObj = (LPUNKNOWN) m_pExchExtMessageEvents;
      m_pExchExtMessageEvents->SetContext (m_lContext);
    }
  else if (riid == IID_IExchExtAttachedFileEvents)
    {
      *ppvObj = (LPUNKNOWN)m_pExchExtAttachedFileEvents;
    }  
  else if (riid == IID_IExchExtPropertySheets) 
    {
      if (m_lContext != EECONTEXT_PROPERTYSHEETS)
	return E_NOINTERFACE;
      *ppvObj = (LPUNKNOWN) m_pExchExtPropertySheets;
    }
//   else if (riid == IID_IOutlookExtItemEvents)
//     {
//       *ppvObj = (LPUNKNOWN)m_pOutlookExtItemEvents;
//     }  
  else
    hr = E_NOINTERFACE;
  
  /* On success we need to bump up the reference counter for the
     requested object. */
  if (*ppvObj)
    ((LPUNKNOWN)*ppvObj)->AddRef();
  
  return hr;
}


/* Called once for each new context. Checks the Exchange extension
   version number and the context.  Returns: S_OK if the extension
   should used in the requested context, otherwise S_FALSE.  PEECB is
   a pointer to Exchange extension callback function.  LCONTEXT is the
   context code at time of being called. LFLAGS carries flags to
   indicate whether the extension should be installed modal.
*/
STDMETHODIMP 
GpgolExt::Install(LPEXCHEXTCALLBACK pEECB, ULONG lContext, ULONG lFlags)
{
  static int version_shown;
  ULONG lBuildVersion;
  ULONG lActualVersion;
  ULONG lVirtualVersion;


  /* Save the context in an instance variable. */
  m_lContext = lContext;

  log_debug ("%s:%s: context=%s flags=0x%lx\n", SRCNAME, __func__,
             ext_context_name (lContext), lFlags);
  
  /* Check version.  This install method is called by Outlook even
     before the OOM interface is available, thus we need to keep on
     checking for the olversion until we get one.  Only then we can
     display the complete version and do a final test to see whether
     this is a supported version. */
  if (!olversion)
    {
      LPDISPATCH obj = get_eecb_object (pEECB);
      if (obj)
        {
          LPDISPATCH disp = get_oom_object (obj, "Application");
          if (disp)
            {
              olversion = get_oom_string (disp, "Version");
              disp->Release ();
            }
          obj->Release ();
        }
    }
  pEECB->GetVersion (&lBuildVersion, EECBGV_GETBUILDVERSION);
  pEECB->GetVersion (&lActualVersion, EECBGV_GETACTUALVERSION);
  pEECB->GetVersion (&lVirtualVersion, EECBGV_GETVIRTUALVERSION);
  if (!version_shown)
    {
      log_debug ("%s:%s: detected Outlook build version 0x%lx (%lu.%lu)\n",
                 SRCNAME, __func__, lBuildVersion,
                 (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK) >> 16,
                 (lBuildVersion & EECBGV_BUILDVERSION_MINOR_MASK));
      log_debug ("%s:%s:                 actual version 0x%lx (%u.%u.%u.%u)\n",
                 SRCNAME, __func__, lActualVersion, 
                 (unsigned int)((lActualVersion >> 24) & 0xff),
                 (unsigned int)((lActualVersion >> 16) & 0xff),
             (unsigned int)((lActualVersion >> 8) & 0xff),
                 (unsigned int)(lActualVersion & 0xff));
      log_debug ("%s:%s:                virtual version 0x%lx (%u.%u.%u.%u)\n",
                 SRCNAME, __func__, lVirtualVersion, 
                 (unsigned int)((lVirtualVersion >> 24) & 0xff),
                 (unsigned int)((lVirtualVersion >> 16) & 0xff),
             (unsigned int)((lVirtualVersion >> 8) & 0xff),
                 (unsigned int)(lVirtualVersion & 0xff));
      if (olversion)
        {
          log_debug ("%s:%s:                    OOM version %s\n",
                     SRCNAME, __func__, olversion);
          version_shown = 1;
        }
    }
  
  if (EECBGV_BUILDVERSION_MAJOR
      != (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK))
    {
      log_debug ("%s:%s: invalid version 0x%lx\n",
                 SRCNAME, __func__, lBuildVersion);
      return S_FALSE;
    }
  
  /* The version numbers as returned by GetVersion are the same for
     OL2003 as well as for recent OL2002.  My guess is that this
     version comes from the Exchange Client Extension API and that has
     been updated in all version of OL.  Thus we also need to check
     the version number as taken from the Outlook Object Model.  */
  if ( (lBuildVersion & EECBGV_BUILDVERSION_MAJOR_MASK) < 13
       || (lBuildVersion & EECBGV_BUILDVERSION_MINOR_MASK) < 1573
       || (olversion && atoi (olversion) < 11) )
    {
      static int shown;
      HWND hwnd;
      
      if (!shown)
        {
          shown = 1;
          
          if (FAILED(pEECB->GetWindow (&hwnd)))
            hwnd = NULL;
          MessageBox (hwnd,
                      _("This version of Outlook is too old!\n\n"
                        "At least versions of Outlook 2003 older than SP2 "
                        "exhibit crashes when sending messages and messages "
                        "might get stuck in the outgoing queue.\n\n"
                        "Please update at least to SP2 before trying to send "
                        "a message"),
                      "GpgOL", MB_ICONSTOP|MB_OK);
        }
    }

  /* If the version from the OOM is available we can assume that the
     OOM is ready for use and thus we can install the event sinks. */
  if (olversion)
    {
      install_sinks (pEECB);
    }
  

  /* Check context. */
  if (   lContext == EECONTEXT_PROPERTYSHEETS
      || lContext == EECONTEXT_SENDNOTEMESSAGE
      || lContext == EECONTEXT_SENDPOSTMESSAGE
      || lContext == EECONTEXT_SENDRESENDMESSAGE
      || lContext == EECONTEXT_READNOTEMESSAGE
      || lContext == EECONTEXT_READPOSTMESSAGE
      || lContext == EECONTEXT_READREPORTMESSAGE
      || lContext == EECONTEXT_VIEWER
      || lContext == EECONTEXT_SESSION)
    {
      return S_OK;
    }
  
  log_debug ("%s:%s: can't handle this context\n", SRCNAME, __func__);
  return S_FALSE;
}


static void
install_forms (void)
{
  HRESULT hr;
  LPMAPIFORMCONTAINER formcontainer = NULL;
  static char const *forms[] = 
    {
      "gpgol",
      "gpgol-ms",
      "gpgol-cs",
      NULL,
    };
  int formidx;
  LANGID langid;
  const char *langsuffix;
  char buffer[MAX_PATH+10];
  char *datadir;
  int any_error = 0;

  langid = PRIMARYLANGID (LANGIDFROMLCID (GetThreadLocale ()));
  switch (langid)
    {
    case LANG_ENGLISH: langsuffix = "en"; break;
    case LANG_GERMAN:  langsuffix = "de"; break;
    default: 
      log_debug ("%s:%s: No forms available for primary language %d\n",
                 SRCNAME, __func__, (int)langid);
      /* Don't try again.  */
      opt.forms_revision = SVN_REVISION;
      write_options ();
      return;
    }

  MAPIOpenLocalFormContainer (&formcontainer);
  if (!formcontainer)
    {
      log_error ("%s:%s: error getting local form container\n",
                 SRCNAME, __func__);
      return;
    }

  datadir = get_data_dir ();
  if (!datadir)
    {
      log_error ("%s:%s: error getting data directory\n",
                 SRCNAME, __func__);
      return;
    }

  for (formidx=0; forms[formidx]; formidx++)
    {

      snprintf (buffer, MAX_PATH, "%s\\%s_%s.cfg",
                datadir, forms[formidx], langsuffix);
      hr = formcontainer->InstallForm (0, MAPIFORM_INSTALL_OVERWRITEONCONFLICT,
                                       buffer);
      if (hr)
        {
          any_error = 1;
          log_error ("%s:%s: installing form `%s' failed: hr=%#lx\n",
                     SRCNAME, __func__, buffer, hr);
        }
      else
        log_debug ("%s:%s: form `%s' installed\n",  SRCNAME, __func__, buffer);
    }

  xfree (datadir);

  if (!any_error)
    {
      opt.forms_revision = SVN_REVISION;
      write_options ();
    }
}



static void
install_sinks (LPEXCHEXTCALLBACK eecb)
{
  static int done;
  LPOUTLOOKEXTCALLBACK pCb;
  LPUNKNOWN rootobj;

  /* We call this function just once.  */
  if (done)
    return;
  done++;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);

  pCb = NULL;
  rootobj = NULL;
  eecb->QueryInterface (IID_IOutlookExtCallback, (LPVOID*)&pCb);
  if (pCb)
    pCb->GetObject (&rootobj);
  if (rootobj)
    {
      LPDISPATCH disp;

      disp = get_oom_object ((LPDISPATCH)rootobj, "Application.Explorers");
      if (!disp)
        log_error ("%s:%s: Explorers NOT found\n", SRCNAME, __func__);
      else
        {
          install_GpgolExplorersEvents_sink (disp);
          /* Fixme: Register the event sink object somewhere.  */
          disp->Release ();
        }
      
      /* It seems that when installing this sink the first explorer
         has already been created and thus we don't see a NewInspector
         event.  Thus we create the controls direct.  */
      disp = get_oom_object ((LPDISPATCH)rootobj,
                             "Application.ActiveExplorer");
      if (!disp)
        log_error ("%s:%s: ActiveExplorer NOT found\n", SRCNAME, __func__);
      else
        {
          add_explorer_controls ((LPOOMEXPLORER)disp);
          disp->Release ();
        }

      disp = get_oom_object ((LPDISPATCH)rootobj, "Application.Inspectors");
      if (!disp)
        log_error ("%s:%s: Inspectors NOT found\n", SRCNAME, __func__);
      else
        {
          install_GpgolInspectorsEvents_sink (disp);
          /* Fixme: Register the event sink object somewhere.  */
          disp->Release ();
        }
      
      rootobj->Release ();
    }
  
  log_debug ("%s:%s: Leave", SRCNAME, __func__);
}


/* Return the OOM object via EECB.  If it is not available return
   NULL.  */
LPDISPATCH
get_eecb_object (LPEXCHEXTCALLBACK eecb)
{
  HRESULT hr;
  LPOUTLOOKEXTCALLBACK pCb = NULL;
  LPUNKNOWN pObj = NULL;
  LPDISPATCH pDisp = NULL;
  LPDISPATCH result = NULL;
  
  hr = eecb->QueryInterface (IID_IOutlookExtCallback, (LPVOID*)&pCb);
  if (hr == S_OK && pCb)
    {
      pCb->GetObject (&pObj);
      if (pObj)
        {
          /* We better query for IDispatch.  */
          hr = pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
          if (hr == S_OK && pDisp)
            result = pDisp;
          pObj->Release ();
        }
      pCb->Release ();
    }
  return result;
}


