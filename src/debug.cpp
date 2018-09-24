/* debug.cpp - Debugging / Log helpers for GpgOL
 * Copyright (C) 2018 by by Intevation GmbH
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

#include "common_indep.h"

#include <gpg-error.h>

#include <string>
#include <unordered_map>

/* The malloced name of the logfile and the logging stream.  If
   LOGFILE is NULL, no logging is done. */
static char *logfile;
static FILE *logfp;

#ifdef HAVE_W32_SYSTEM

/* Acquire the mutex for logging.  Returns 0 on success. */
static int
lock_log (void)
{
  int code = WaitForSingleObject (log_mutex, 10000);
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
#endif

const char *
get_log_file (void)
{
  return logfile? logfile : "";
}

void
set_log_file (const char *name)
{
#ifdef HAVE_W32_SYSTEM
  if (!lock_log ())
    {
#endif
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
#ifdef HAVE_W32_SYSTEM
      unlock_log ();
    }
#endif
}

static void
do_log (const char *fmt, va_list a, int w32err, int err,
        const void *buf, size_t buflen)
{
  if (!logfile)
    return;

#ifdef HAVE_W32_SYSTEM
  if (!opt.enable_debug)
    return;

  if (lock_log ())
    {
      OutputDebugStringA ("GpgOL: Failed to log.");
      return;
    }
#endif

  if (!strcmp (logfile, "stdout"))
    {
      logfp = stdout;
    }
  else if (!strcmp (logfile, "stderr"))
    {
      logfp = stderr;
    }
  if (!logfp)
    logfp = fopen (logfile, "a+");
#ifdef HAVE_W32_SYSTEM
  if (!logfp)
    {
      unlock_log ();
      return;
    }

  char time_str[9];
  SYSTEMTIME utc_time;
  GetSystemTime (&utc_time);
  if (GetTimeFormatA (LOCALE_INVARIANT,
                      TIME_FORCE24HOURFORMAT | LOCALE_USE_CP_ACP,
                      &utc_time,
                      "HH:mm:ss",
                      time_str,
                      9))
    {
      fprintf (logfp, "%s/%lu/",
               time_str,
               (unsigned long)GetCurrentThreadId ());
    }
  else
    {
      fprintf (logfp, "unknown/%lu/",
               (unsigned long)GetCurrentThreadId ());
    }
#endif

  if (err == 1)
    fputs ("ERROR/", logfp);
  vfprintf (logfp, fmt, a);
#ifdef HAVE_W32_SYSTEM
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
#endif
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
#ifdef HAVE_W32_SYSTEM
  unlock_log ();
#endif
}

const char *
log_srcname (const char *file)
{
  const char *s = strrchr (file, '/');
  return s? s+1:file;
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
log_hexdump (const void *buf, size_t buflen, const char *fmt, ...)
{
  va_list a;

  va_start (a, fmt);
  do_log (fmt, a, 0, 2, buf, buflen);
  va_end (a);
}

#ifdef HAVE_W32_SYSTEM
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
#endif

GPGRT_LOCK_DEFINE (anon_str_lock);

/* Weel ok this survives unload but we don't want races
   and it makes a bit of sense to keep the strings constant. */
static std::unordered_map<std::string, std::string> str_map;

const char *anonstr (const char *data)
{
  static int64_t cnt;
  if (opt.enable_debug & DBG_DATA)
    {
      return data;
    }
  if (!data)
    {
      return "gpgol_str_null";
    }
  if (!strlen (data))
    {
      return "gpgol_str_empty";
    }

  gpgrt_lock_lock (&anon_str_lock);
  const std::string strData (data);
  auto it = str_map.find (strData);

  if (it == str_map.end ())
    {
      const auto anon = std::string ("gpgol_string_") + std::to_string (++cnt);
      str_map.insert (std::make_pair (strData, anon));
      it = str_map.find (strData);
    }

  // As the data is saved in our map we can return
  // the c_str as it won't be touched as const.

  gpgrt_lock_unlock (&anon_str_lock);

  TRACEPOINT;
  return it->second.c_str();
}
