/* olflange.cpp - Connect GpgOL to Outlook
 * Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 * Copyright (C) 2004, 2005, 2007, 2008 g10 Code GmbH
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
/* Include every header that defines a GUID below this
   macro. Otherwise the GUID's will only be declared and
   not defined. */
#define INITGUID
#endif

#include <initguid.h>
#include "mymapi.h"
#include "mymapitags.h"

#include "common.h"
#include "mapihelp.h"

#include "olflange.h"
#include "gpgoladdin.h"

static char *olversion;

EXTERN_C int
get_ol_main_version (void)
{
  return olversion? atoi (olversion): 0;
}

/* Registers this as an addin for outlook 2010.
   This basically updates some Registry entries.
   Documentation to be found at:
   http://msdn.microsoft.com/en-us/library/bb386106%28v=vs.110%29.aspx
   */
STDAPI
DllRegisterServer (void)
{
  HKEY hkey, hkey2;
  CHAR szKeyBuf[MAX_PATH+1024];
  CHAR szEntry[MAX_PATH+512];
  TCHAR szModuleFileName[MAX_PATH];
  DWORD dwTemp = 0;
  long ec;
  HKEY root_key;

  int inst_global = is_elevated ();

  if (inst_global)
    {
      root_key = HKEY_LOCAL_MACHINE;
    }
  else
    {
      root_key = HKEY_CURRENT_USER;
    }

  /* Get server location. */
  if (!GetModuleFileName(glob_hinst, szModuleFileName, MAX_PATH))
    return E_FAIL;

  hkey = NULL;
  lstrcpy (szKeyBuf, "Software\\GNU\\GpgOL");
  RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (hkey != NULL)
    RegCloseKey (hkey);

  /* Register the CLSID in the registry */
  hkey = NULL;

  if (inst_global)
    {
      strcpy (szKeyBuf, "CLSID\\" CLSIDSTR_GPGOL);
      ec = RegCreateKeyEx (HKEY_CLASSES_ROOT, szKeyBuf, 0, NULL,
                   REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
      OutputDebugString("Created: ");
      OutputDebugString(szKeyBuf);
    }
  else
    {
      strcpy (szKeyBuf, "Software\\Classes\\CLSID\\" CLSIDSTR_GPGOL);
      ec = RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
                   REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
    }
  if (ec != ERROR_SUCCESS)
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      return E_ACCESSDENIED;
    }

  strcpy (szEntry, GPGOL_PRETTY);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* Set the Inproc server value */
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

  /* Set the threading model used */
  strcpy (szEntry, "Both");
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, "ThreadingModel", 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* Set the Prog ID */
  strcpy (szKeyBuf, "ProgID");
  ec = RegCreateKeyEx (hkey, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey2, NULL);
  if (ec != ERROR_SUCCESS)
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      RegCloseKey (hkey);
      return E_ACCESSDENIED;
    }
  strcpy (szEntry, GPGOL_PROGID);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* Make the Prog ID known. This is basically the same as above
   * but necessary so we can refer to the Prog ID as an Outlook
   * Extension
   */
  hkey = NULL;

  if (inst_global)
    {
      strcpy (szKeyBuf, GPGOL_PROGID);
      ec = RegCreateKeyEx (HKEY_CLASSES_ROOT, szKeyBuf, 0, NULL,
                      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
    }
  else
    {
      strcpy (szKeyBuf, "Software\\Classes\\" GPGOL_PROGID);
      ec = RegCreateKeyEx (HKEY_CURRENT_USER, szKeyBuf, 0, NULL,
                      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);

    }
  if (ec != ERROR_SUCCESS)
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      return E_ACCESSDENIED;
    }

  strcpy (szEntry, GPGOL_PRETTY);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* Point from the Prog ID entry to the CSLID */

  strcpy (szKeyBuf, "CLSID");
  ec = RegCreateKeyEx (hkey, szKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &hkey2, NULL);
  if (ec != ERROR_SUCCESS)
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      RegCloseKey (hkey);
      return E_ACCESSDENIED;
    }
  strcpy (szEntry, CLSIDSTR_GPGOL);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey2, NULL, 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* Register ourself as an extension for outlook >= 14 */

  strcpy (szKeyBuf, "Software\\Microsoft\\Office\\Outlook\\Addins\\" GPGOL_PROGID);
  ec = RegCreateKeyEx (root_key, szKeyBuf, 0, NULL,
                  REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hkey, NULL);
  if (ec != ERROR_SUCCESS)
    {
      fprintf (stderr, "creating key `%s' failed: ec=%#lx\n", szKeyBuf, ec);
      return E_ACCESSDENIED;
    }

  /* Load connected and load Bootload */
  dwTemp = 0x01 | 0x02;
  RegSetValueEx (hkey, "LoadBehavior", 0, REG_DWORD, (BYTE*)&dwTemp, 4);

  /* We are not commandline save */
  dwTemp = 0;
  RegSetValueEx (hkey, "CommandLineSafe", 0, REG_DWORD, (BYTE*)&dwTemp, 4);

  /* A friendly name (visible in outlook) */
  strcpy (szEntry, GPGOL_PRETTY);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, "FriendlyName", 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  /* A short description (visible in outlook) */
  strcpy (szEntry, GPGOL_DESCRIPTION);
  dwTemp = strlen (szEntry) + 1;
  RegSetValueEx (hkey, "Description", 0, REG_SZ, (BYTE*)szEntry, dwTemp);

  RegCloseKey (hkey2);
  RegCloseKey (hkey);


  log_debug ("DllRegisterServer succeeded\n");
  return S_OK;
}


/* Unregisters this module as an Exchange extension / Addin. */
STDAPI
DllUnregisterServer (void)
{
  HKEY hkey;
  CHAR buf[MAX_PATH+1024];
  DWORD ntemp;
  long res;
  HKEY root_key;

  if (is_elevated ())
    {
      root_key = HKEY_LOCAL_MACHINE;
    }
  else
    {
      root_key = HKEY_CURRENT_USER;
    }

  /* We still unregister the old client extension code */
  strcpy (buf, "Software\\Microsoft\\Exchange\\Client\\Extensions");
  /* Create and open key and subkey. */
  res = RegCreateKeyEx (root_key, buf, 0, NULL,
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
  strcpy (buf, "CLSID\\" CLSIDSTR_GPGOL "\\ProgID");
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);
  strcpy (buf, "CLSID\\" CLSIDSTR_GPGOL);
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);

  /* Delete ProgID */
  strcpy (buf, GPGOL_PROGID "\\CLSID");
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);
  strcpy (buf, GPGOL_PROGID);
  RegDeleteKey (HKEY_CLASSES_ROOT, buf);

  /* Delete Addin entry */
  strcpy (buf, "Software\\Microsoft\\Office\\Outlook\\Addins\\" GPGOL_PROGID);
  RegDeleteKey (root_key, buf);

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

void
install_forms (void)
{
  HRESULT hr;
  LPMAPIFORMCONTAINER formcontainer = NULL;
  static char const *forms[] =
    {
      "gpgol",
      "gpgol-ms",
      "gpgol-cs",
       /* The InfoPath we use for sending, to get outlook
          to do the S/MIME handling. */
      "gpgol-form-signed",
      "gpgol-form-encrypted",
      NULL,
    };
  int formidx;
  char buffer[MAX_PATH+10];
  char *datadir;
  int any_error = 0;

  MAPIOpenLocalFormContainer (&formcontainer);
  if (!formcontainer)
    {
      log_error ("%s:%s: error getting local form container\n",
                 SRCNAME, __func__);
      return;
    }
  memdbg_addRef (formcontainer);

  datadir = get_data_dir ();
  if (!datadir)
    {
      log_error ("%s:%s: error getting data directory\n",
                 SRCNAME, __func__);
      gpgol_release (formcontainer);
      return;
    }

  for (formidx=0; forms[formidx]; formidx++)
    {

      snprintf (buffer, MAX_PATH, "%s\\%s.cfg",
                datadir, forms[formidx]);
      hr = formcontainer->InstallForm (0, MAPIFORM_INSTALL_OVERWRITEONCONFLICT,
                                       buffer);
      if (hr)
        {
          any_error = 1;
          LPMAPIERROR err;
          formcontainer->GetLastError (hr, 0, &err);
          log_error ("%s:%s: installing form `%s' failed: hr=%#lx err=%s\n",
                     SRCNAME, __func__, buffer, hr, err ? err->lpszError : "null");
          MAPIFreeBuffer (err);
        }
      else
        log_debug ("%s:%s: form `%s' installed\n",  SRCNAME, __func__, buffer);
    }

  gpgol_release (formcontainer);
  xfree (datadir);

  if (!any_error)
    {
      opt.forms_revision = GPGOL_FORMS_REVISION;
      write_options ();
    }
}
