/* main.c - DLL entry point
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGME Dialogs.
 *
 * GPGME Dialogs is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGME Dialogs is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGME Dialogs; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */

#include <config.h>

#include <windows.h>

#include <gpgme.h>

#include "mymapi.h"
#include "mymapitags.h"


#include "intern.h"
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
  opt.add_default_key = 1;
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


/* Entry point called by DLL loader. */
int WINAPI
DllMain (HINSTANCE hinst, DWORD reason, LPVOID reserved)
{
  if (reason == DLL_PROCESS_ATTACH )
    {
      set_global_hinstance (hinst);
      /* The next call initializes subsystems of gpgme and should be
         done as early as possible.  The actual return value is (the
         version string) is not used here.  It may be called at any
         time later for this. */
      gpgme_check_version (NULL);

      /* Early initializations of our subsystems. */
      if (initialize_main ())
        return FALSE;
      if (initialize_passcache ())
        return FALSE;
      if (initialize_msgcache ())
        return FALSE;
      init_options ();
    }

  return TRUE;
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
  
  fprintf (logfp, "%lu/", (unsigned long)GetCurrentThreadId ());
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
      fputs (tmpbuf, logfp);
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


/* Read option settings from the Registry. */
void
read_options (void)
{
  char *val = NULL;
 
  load_extension_value ("autoSignAttachments", &val);
  opt.auto_sign_attach = val == NULL || *val != '1' ? 0 : 1;
  xfree (val); val = NULL;
  
  load_extension_value ("saveDecryptedAttachments", &val);
  opt.save_decrypted_attach = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("encryptDefault", &val);
  opt.encrypt_default = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("signDefault", &val);
  opt.sign_default = val == NULL || *val != '1'? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("addDefaultKey", &val);
  opt.add_default_key = val == NULL || *val != '1' ? 0 : 1;
  xfree (val); val = NULL;

  load_extension_value ("storePasswdTime", &val);
  opt.passwd_ttl = val == NULL || *val == '0'? 0 : atol (val);
  xfree (val); val = NULL;

  load_extension_value ("encodingFormat", &val);
  opt.enc_format = val == NULL? GPG_FMT_CLASSIC  : atol (val);
  xfree (val); val = NULL;

  load_extension_value ("logFile", &val);
  set_log_file (val);
  xfree (val); val = NULL;
  
  load_extension_value ("defaultKey", &val);
  set_default_key (val);
  xfree (val); val = NULL;

  /* Note, that on prupose theses flags are only Registry changeable.
     The format of the entry is a string of the format "0xnnnnnnn". */
  load_extension_value ("compatFlags", &val);
  opt.compat_flags = val? strtoul (val, NULL, 0) : 0;
  xfree (val); val = NULL;
}


/* Write current options back to the Registry. */
int
write_options (void)
{
  struct 
  {
    const char *name;
    int  value;
  } table[] = {
    {"encryptDefault",           opt.encrypt_default},
    {"signDefault",              opt.sign_default},
    {"addDefaultKey",            opt.add_default_key},
    {"saveDecryptedAttachments", opt.save_decrypted_attach},
    {"autoSignAttachments",      opt.auto_sign_attach},
    {NULL, 0}
  };
  char buf[32];
  int i;

  for (i=0; table[i].name; i++) 
    store_extension_value (table[i].name, table[i].value? "1": "0");

  store_extension_value ("logFile", logfile? logfile: "");
  store_extension_value ("defaultKey", opt.default_key? opt.default_key:"");
  
  sprintf (buf, "%d", opt.passwd_ttl);
  store_extension_value ("storePasswdTime", buf);
  
  sprintf (buf, "%d", opt.enc_format);
  store_extension_value ("encodingFormat", buf);

  return 0;
}

