/* main.c - DLL entry point
 *	Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include <config.h>

#include <windows.h>
#include <wincrypt.h>
#include <ctype.h>

#include "mymapi.h"
#include "mymapitags.h"

#include "common.h"
#include "passcache.h"
#include "msgcache.h"
#include "mymapi.h"

/* The malloced name of the logfile and the logging stream.  If
   LOGFILE is NULL, no logging is done. */
static char *logfile;
static FILE *logfp;

/* For certain operations we need to acquire a log on the logging
   functions.  This lock is controlled by this Mutex. */
static HANDLE log_mutex;

/* The session key used to temporary encrypt attachments.  It is
   initialized at startup.  */
static char *the_session_key;

/* The session marker to identify this session.  Its value is not
  confidential.  It is initialized at startup.  */
static char *the_session_marker;

/* Local function prototypes. */
static char *get_locale_dir (void);
static void drop_locale_dir (char *locale_dir);



/* Initialization of gloabl options.  These are merely the defaults
   and will get updated later from the Registry.  That is done later
   at the time Outlook calls its entry point the first time. */
static void
init_options (void)
{
  opt.passwd_ttl = 10; /* Seconds. Use a small value, so that no
                          multiple prompts for attachment encryption
                          are issued. */
  opt.enc_format = GPG_FMT_CLASSIC;
}


/* Early initialization of this module.  This is done right at startup
   with only one thread running.  Should be called only once. Returns
   0 on success. */
static int
initialize_main (void)
{
  SECURITY_ATTRIBUTES sa;
  
  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = FALSE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  log_mutex = CreateMutex (&sa, FALSE, NULL);
  return log_mutex? 0 : -1;
}

/* Return nbytes of cryptographic strong random.  Caller needs to free
   the returned buffer.  */
static char *
get_crypt_random (size_t nbytes) 
{
  HCRYPTPROV prov;
  char *buffer;

  if (!CryptAcquireContext (&prov, NULL, NULL, PROV_RSA_FULL, 
                            (CRYPT_VERIFYCONTEXT|CRYPT_SILENT)) )
    return NULL;
  
  buffer = xmalloc (nbytes);
  if (!CryptGenRandom (prov, nbytes, buffer))
    {
      xfree (buffer);
      buffer = NULL;
    }
  CryptReleaseContext (prov, 0);
  return buffer;
}


/* Initialize the session key and the session marker.  */
static int
initialize_session_key (void)
{
  the_session_key = get_crypt_random (16+sizeof (unsigned int)+8);
  if (the_session_key)
    {
      /* We use rand() in generate_boundary so we need to seed it. */
      unsigned int tmp;

      memcpy (&tmp, the_session_key+16, sizeof (unsigned int));
      srand (tmp);

      /* And save the session marker. */
      the_session_marker = the_session_key + 16 + sizeof (unsigned int);
    }
  return !the_session_key;
}


static void
i18n_init (void)
{
  char *locale_dir;

#ifdef ENABLE_NLS
# ifdef HAVE_LC_MESSAGES
  setlocale (LC_TIME, "");
  setlocale (LC_MESSAGES, "");
# else
  setlocale (LC_ALL, "" );
# endif
#endif

  locale_dir = get_locale_dir ();
  if (locale_dir)
    {
      bindtextdomain (PACKAGE_GT, locale_dir);
      drop_locale_dir (locale_dir);
    }
  textdomain (PACKAGE_GT);
}


/* Entry point called by DLL loader. */
int WINAPI
DllMain (HINSTANCE hinst, DWORD reason, LPVOID reserved)
{
  (void)reserved;

  if (reason == DLL_PROCESS_ATTACH)
    {
      set_global_hinstance (hinst);
      /* The next call initializes subsystems of gpgme and should be
         done as early as possible.  The actual return value (the
         version string) is not used here.  It may be called at any
         time later for this. */
      gpgme_check_version (NULL);

      /* Early initializations of our subsystems. */
      if (initialize_main ())
        return FALSE;
      i18n_init ();
      if (initialize_session_key ())
        return FALSE;
      if (initialize_passcache ())
        return FALSE;
      if (initialize_msgcache ())
        return FALSE;
      if (initialize_inspectors ())
        return FALSE;
      init_options ();
    }
  else if (reason == DLL_PROCESS_DETACH)
    {
    }
  
  return TRUE;
}



/* Return the static session key we are using for temporary encrypting
   attachments.  The session key is guaranteed to be available.  */
const void *
get_128bit_session_key (void)
{
  return the_session_key;
}


const void *
get_64bit_session_marker (void)
{
  return the_session_marker;
}


/* Return a new allocated IV of size NBYTES.  Caller must free it.  On
   error NULL is returned. */
void *
create_initialization_vector (size_t nbytes)
{
  return get_crypt_random (nbytes);
}


/* Acquire the mutex for logging.  Returns 0 on success. */
static int 
lock_log (void)
{
  int code = WaitForSingleObject (log_mutex, INFINITE);
  return code != WAIT_OBJECT_0;
}

/* Release the mutex for logging. No error return is done because this
   is a fatal error anyway and we have no means for proper
   notification. */
static void
unlock_log (void)
{
  ReleaseMutex (log_mutex);
}



static void
do_log (const char *fmt, va_list a, int w32err, int err,
        const void *buf, size_t buflen)
{
  if (!logfile)
    return;

  if (lock_log ())
    return;
  
  if (!logfp)
    logfp = fopen (logfile, "a+");
  if (!logfp)
    {
      unlock_log ();
      return;
    }
  
  fprintf (logfp, "%05lu/%lu/", 
           ((unsigned long)GetTickCount () % 100000),
           (unsigned long)GetCurrentThreadId ());
  if (err == 1)
    fputs ("ERROR/", logfp);
  vfprintf (logfp, fmt, a);
  if (w32err) 
    {
      char tmpbuf[256];
      
      FormatMessage (FORMAT_MESSAGE_FROM_SYSTEM, NULL, w32err, 
                     MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), 
                     tmpbuf, sizeof (tmpbuf)-1, NULL);
      fputs (": ", logfp);
      if (*tmpbuf && tmpbuf[strlen (tmpbuf)-1] == '\n')
        tmpbuf[strlen (tmpbuf)-1] = 0;
      if (*tmpbuf && tmpbuf[strlen (tmpbuf)-1] == '\r')
        tmpbuf[strlen (tmpbuf)-1] = 0;
      fprintf (logfp, "%s (%d)", tmpbuf, w32err);
    }
  if (buf)
    {
      const unsigned char *p = (const unsigned char*)buf;

      for ( ; buflen; buflen--, p++)
        fprintf (logfp, "%02X", *p);
      putc ('\n', logfp);
    }
  else if ( *fmt && fmt[strlen (fmt) - 1] != '\n')
    putc ('\n', logfp);

  fflush (logfp);
  unlock_log ();
}


void 
log_debug (const char *fmt, ...)
{
  va_list a;
  
  va_start (a, fmt);
  do_log (fmt, a, 0, 0, NULL, 0);
  va_end (a);
}

void 
log_error (const char *fmt, ...)
{
  va_list a;
  
  va_start (a, fmt);
  do_log (fmt, a, 0, 1, NULL, 0);
  va_end (a);
}

void 
log_vdebug (const char *fmt, va_list a)
{
  do_log (fmt, a, 0, 0, NULL, 0);
}


void 
log_debug_w32 (int w32err, const char *fmt, ...)
{
  va_list a;

  if (w32err == -1)
    w32err = GetLastError ();
  
  va_start (a, fmt);
  do_log (fmt, a, w32err, 0, NULL, 0);
  va_end (a);
}

void 
log_error_w32 (int w32err, const char *fmt, ...)
{
  va_list a;

  if (w32err == -1)
    w32err = GetLastError ();
  
  va_start (a, fmt);
  do_log (fmt, a, w32err, 1, NULL, 0);
  va_end (a);
}


void 
log_hexdump (const void *buf, size_t buflen, const char *fmt, ...)
{
  va_list a;

  va_start (a, fmt);
  do_log (fmt, a, 0, 2, buf, buflen);
  va_end (a);
}


static void
do_log_window_info (HWND window, int level)
{
  char buf[1024+1];
  char name[200];
  int nname;
  char *pname;
  DWORD pid;

  if (!window)
    return;

  GetWindowThreadProcessId (window, &pid);
  if (pid != GetCurrentProcessId ())
    return;
 
  memset (buf, 0, sizeof (buf));
  GetWindowText (window, buf, sizeof (buf)-1);
  nname = GetClassName (window, name, sizeof (name)-1);
  if (nname)
    pname = name;
  else
    pname = NULL;

  if (level == -1)
    log_debug ("  parent=%p/%lu (%s) `%s'", window, (unsigned long)pid,
               pname? pname:"", buf);
  else
    log_debug ("    %*shwnd=%p/%lu (%s) `%s'", level*2, "", window,
               (unsigned long)pid, pname? pname:"", buf);
}


/* Helper to log_window_hierarchy.  */
static HWND
do_log_window_hierarchy (HWND parent, int level)
{
  HWND child;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      do_log_window_info (child, level);
      do_log_window_hierarchy (child, level+1);
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Print a debug message using the format string FMT followed by the
   window hierarchy of WINDOW.  */
void
log_window_hierarchy (HWND window, const char *fmt, ...)
{
  va_list a;

  va_start (a, fmt);
  do_log (fmt, a, 0, 0, NULL, 0);
  va_end (a);
  if (window)
    {
      do_log_window_info (window, -1);
      do_log_window_hierarchy (window, 0);
    }
}


const char *
log_srcname (const char *file)
{
  const char *s = strrchr (file, '/');
  return s? s+1:file;
}

const char *
get_log_file (void)
{
  return logfile? logfile : "";
}

void
set_log_file (const char *name)
{
  if (!lock_log ())
    {
      if (logfp)
        {
          fclose (logfp);
          logfp = NULL;
        }
      xfree (logfile);
      if (!name || *name == '\"' || !*name) 
        logfile = NULL;
      else
        logfile = xstrdup (name);
      unlock_log ();
    }
}



void
set_default_key (const char *name)
{
  if (!lock_log ())
    {
      if (!name || *name == '\"' || !*name)
        {
          xfree (opt.default_key);
          opt.default_key = NULL;
        }
      else
        {
          xfree (opt.default_key);
          opt.default_key = xstrdup (name);;
        }
      unlock_log ();
    }
}


static char *
get_locale_dir (void)
{
  char *instdir;
  char *p;
  char *dname;

  instdir = read_w32_registry_string ("HKEY_LOCAL_MACHINE", GNUPG_REGKEY,
				      "Install Directory");
  if (!instdir)
    return NULL;
  
  /* Build the key: "<instdir>/share/locale".  */
#define SLDIR "\\share\\locale"
  dname = malloc (strlen (instdir) + strlen (SLDIR) + 1);
  if (!dname)
    {
      free (instdir);
      return NULL;
    }
  p = dname;
  strcpy (p, instdir);
  p += strlen (instdir);
  strcpy (p, SLDIR);
  
  free (instdir);
  
  return dname;
}


static void
drop_locale_dir (char *locale_dir)
{
  free (locale_dir);
}


/* Read option settings from the Registry. */
void
read_options (void)
{
  static int warnings_shown;
  char *val = NULL;

  /* Set the log file first so that output from this function is
     logged too.  */
  load_extension_value ("logFile", &val);
  set_log_file (val);
  xfree (val); val = NULL;
  
  /* Parse the debug flags.  */
  load_extension_value ("enableDebug", &val);
  opt.enable_debug = 0;
  if (val)
    {
      char *p, *pend;

      trim_spaces (val);
      for (p = val; p; p = pend)
        {
          pend = strpbrk (p, ", \t\n\r\f");
          if (pend)
            {
              *pend++ = 0;
              pend += strspn (pend, ", \t\n\r\f");
            }
          if (isascii (*p) && isdigit (*p))
            opt.enable_debug |= strtoul (p, NULL, 0);
          else if (!strcmp (p, "ioworker"))
            opt.enable_debug |= DBG_IOWORKER;
          else if (!strcmp (p, "ioworker-extra"))
            opt.enable_debug |= DBG_IOWORKER_EXTRA;
          else if (!strcmp (p, "filter"))
            opt.enable_debug |= DBG_FILTER;
          else if (!strcmp (p, "filter-extra"))
            opt.enable_debug |= DBG_FILTER_EXTRA;
          else if (!strcmp (p, "memory"))
            opt.enable_debug |= DBG_MEMORY;
          else if (!strcmp (p, "commands"))
            opt.enable_debug |= DBG_COMMANDS;
          else if (!strcmp (p, "mime-parser"))
            opt.enable_debug |= DBG_MIME_PARSER;
          else if (!strcmp (p, "mime-data"))
            opt.enable_debug |= DBG_MIME_DATA;
          else
            log_debug ("invalid debug flag `%s' ignored", p);
        }
    }
  else
    {
      /* To help the user enable debugging make sure that the registry
         key exists.  Note that the other registry keys are stored
         after using the configuration dialog.  */
      store_extension_value ("enableDebug", "0");
    }
  xfree (val); val = NULL;
  if (opt.enable_debug)
    log_debug ("enabled debug flags:%s%s%s%s%s%s%s%s\n",
               (opt.enable_debug & DBG_IOWORKER)? " ioworker":"",
               (opt.enable_debug & DBG_IOWORKER_EXTRA)? " ioworker-extra":"",
               (opt.enable_debug & DBG_FILTER)? " filter":"",
               (opt.enable_debug & DBG_FILTER_EXTRA)? " filter-extra":"",
               (opt.enable_debug & DBG_MEMORY)? " memory":"",
               (opt.enable_debug & DBG_COMMANDS)? " commands":"",
               (opt.enable_debug & DBG_MIME_PARSER)? " mime-parser":"",
               (opt.enable_debug & DBG_MIME_DATA)? " mime-data":""
               );


  load_extension_value ("enableSmime", &val);
  opt.enable_smime = (!val || atoi (val));
  xfree (val); val = NULL;
  
/*   load_extension_value ("defaultProtocol", &val); */
/*   switch ((!val || *val == '0')? 0 : atol (val)) */
/*     { */
/*     case 1: opt.default_protocol = PROTOCOL_OPENPGP; break; */
/*     case 2: opt.default_protocol = PROTOCOL_SMIME; break; */
/*     case 0: */
/*     default: opt.default_protocol = PROTOCOL_UNKNOWN /\*(auto*)*\/; break; */
/*     } */
/*   xfree (val); val = NULL; */
  opt.default_protocol = PROTOCOL_UNKNOWN; /* (auto)*/

  load_extension_value ("encryptDefault", &val);
  opt.encrypt_default = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("signDefault", &val);
  opt.sign_default = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("previewDecrypt", &val);
  opt.preview_decrypt = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("enableDefaultKey", &val);
  opt.enable_default_key = val == NULL || *val != '1' ? 0 : 1;
  xfree (val); val = NULL;

  if (load_extension_value ("storePasswdTime", &val) )
    opt.passwd_ttl = 600; /* Initial default. */
  else
    opt.passwd_ttl = val == NULL || *val == '0'? 0 : atol (val);
  xfree (val); val = NULL;

  load_extension_value ("encodingFormat", &val);
  opt.enc_format = val == NULL? GPG_FMT_CLASSIC  : atol (val);
  xfree (val); val = NULL;

  load_extension_value ("defaultKey", &val);
  set_default_key (val);
  xfree (val); val = NULL;

  load_extension_value ("preferHtml", &val);
  opt.prefer_html = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("svnRevision", &val);
  opt.svn_revision = val? atol (val) : 0;
  xfree (val); val = NULL;

  load_extension_value ("formsRevision", &val);
  opt.forms_revision = val? atol (val) : 0;
  xfree (val); val = NULL;

  load_extension_value ("announceNumber", &val);
  opt.announce_number = val? atol (val) : 0;
  xfree (val); val = NULL;

  load_extension_value ("bodyAsAttachment", &val);
  opt.body_as_attachment = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  /* Note, that on purpose these flags are only Registry changeable.
     The format of the entry is a string of of "0" and "1" digits; see
     the switch below for a description. */
  memset (&opt.compat, 0, sizeof opt.compat);
  load_extension_value ("compatFlags", &val);
  if (val)
    {
      const char *s = val;
      int i, x;

      for (s=val, i=0; *s; s++, i++)
        {
          x = *s == '1';
          switch (i)
            {
            case 0: opt.compat.no_msgcache = x; break;
            case 1: opt.compat.no_pgpmime = x; break;
            case 2: opt.compat.no_oom_write = x; break;
            case 3: opt.compat.no_preview_info = x; break;
            case 4: opt.compat.old_reply_hack = x; break;
            case 5: opt.compat.auto_decrypt = x; break;
            case 6: opt.compat.no_attestation = x; break;
            case 7: opt.compat.use_mwfmo = x; break;
            }
        }
      log_debug ("Note: using compatibility flags: %s", val);
    }

  if (!warnings_shown)
    {
      char tmpbuf[512];
          
      warnings_shown = 1;
      if (val && *val)
        {
          snprintf (tmpbuf, sizeof tmpbuf,
                    _("Note: Using compatibility flags: %s"), val);
          MessageBox (NULL, tmpbuf, _("GpgOL"), MB_ICONWARNING|MB_OK);
        }
      if (logfile && !opt.enable_debug)
        {
          snprintf (tmpbuf, sizeof tmpbuf,
                    _("Note: Writing debug logs to\n\n\"%s\""), logfile);
          MessageBox (NULL, tmpbuf, _("GpgOL"), MB_ICONWARNING|MB_OK);
        }
    }
  xfree (val); val = NULL;

}


/* Write current options back to the Registry. */
int
write_options (void)
{
  struct 
  {
    const char *name;
    int  mode;
    int  value;
    char *s_val;
  } table[] = {
    {"enableSmime",              0, opt.enable_smime},
/*     {"defaultProtocol",          3, opt.default_protocol}, */
    {"encryptDefault",           0, opt.encrypt_default},
    {"signDefault",              0, opt.sign_default},
    {"previewDecrypt",           0, opt.preview_decrypt},
    {"storePasswdTime",          1, opt.passwd_ttl},
    {"encodingFormat",           1, opt.enc_format},
    {"logFile",                  2, 0, logfile},
    {"defaultKey",               2, 0, opt.default_key},
    {"enableDefaultKey",         0, opt.enable_default_key},
    {"preferHtml",               0, opt.prefer_html},
    {"svnRevision",              1, opt.svn_revision},
    {"formsRevision",            1, opt.forms_revision},
    {"announceNumber",           1, opt.announce_number},
    {"bodyAsAttachment",         0, opt.body_as_attachment},
    {NULL, 0}
  };
  char buf[32];
  int rc, i;
  const char *string;

  for (i=0; table[i].name; i++) 
    {
      switch (table[i].mode)
        {
        case 0:
          string = table[i].value? "1": "0";
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, string);
          rc = store_extension_value (table[i].name, string);
          break;
        case 1:
          sprintf (buf, "%d", table[i].value);
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, buf);
          rc = store_extension_value (table[i].name, buf);
          break;
        case 2:
          string = table[i].s_val? table[i].s_val : "";
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, string);
          rc = store_extension_value (table[i].name, string);
          break;
/*         case 3: */
/*           buf[0] = '0'; */
/*           buf[1] = 0; */
/*           switch (opt.default_protocol) */
/*             { */
/*             case PROTOCOL_UNKNOWN: buf[0] = '0'; /\* auto *\/ break; */
/*             case PROTOCOL_OPENPGP: buf[0] = '1'; break; */
/*             case PROTOCOL_SMIME:   buf[0] = '2'; break; */
/*             } */
/*           log_debug ("storing option `%s' value=`%s'\n", */
/*                      table[i].name, buf); */
/*           rc = store_extension_value (table[i].name, buf); */
/*           break;   */

        default:
          rc = -1;
          break;
        }
      if (rc)
        log_error ("error storing option `%s': rc = %d\n", table[i].name, rc);
    }
  
  return 0;
}
