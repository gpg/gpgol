/* watchmouse.c - Debug utility  
 * Copyright (C) 2008 g10 Code GmbH
 *
 * Watchmouse is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published
 * by the Free Software Foundation; either version 3 of the License,
 * or (at your option) any later version.
 *
 * Watchmouse is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include <stdio.h>
#include <stdlib.h>
#include <windows.h>

#define PGM "watchmouse"

/* The Instance of this process.  */
static HINSTANCE glob_hinst;

/* The current hook handle. */
static HHOOK mouse_hook;





/* Logging stuff.  */
static void
do_log (const char *fmt, va_list a, int w32err, int err,
        const void *buf, size_t buflen)
{
  FILE *logfp;

  if (!logfp)
    logfp = stderr;

/*   if (lock_log ()) */
/*     return; */
  
/*   if (!logfp) */
/*     logfp = fopen (logfile, "a+"); */
/*   if (!logfp) */
/*     { */
/*       unlock_log (); */
/*       return; */
/*     } */
  
  fprintf (logfp, PGM"(%05lu):%s ", 
           ((unsigned long)GetTickCount () % 100000),
           err == 1? " error:":"");
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
/*   unlock_log (); */
}


void 
log_info (const char *fmt, ...)
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
log_vinfo (const char *fmt, va_list a)
{
  do_log (fmt, a, 0, 0, NULL, 0);
}


void 
log_info_w32 (int w32err, const char *fmt, ...)
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






/* Here we receive mouse events.  */
static LRESULT CALLBACK
mouse_proc (int code, WPARAM wparam, LPARAM lparam)
{
  MOUSEHOOKSTRUCT mh;

      log_info ("%d received. w=%ld l=%lu", 
                code, (long)wparam, (unsigned long)lparam);
  if (code < 0)
    ;
  else if (code == HC_ACTION)
    {
      log_info ("HC_ACTION received. w=%ld l=%lu", 
                (long)wparam, (unsigned long)lparam);
    }
  else if (code == HC_NOREMOVE)
    {

    }

  return CallNextHookEx (mouse_hook, code, wparam, lparam);
}



int
main (int argc, char **argv)
{
  glob_hinst = GetModuleHandle (NULL);
  if (!glob_hinst)
    {
      log_error_w32 (-1, "GetModuleHandle failed");
      return 1;
    }
  log_info ("start");

  mouse_hook = SetWindowsHookEx (WH_MOUSE, mouse_proc, NULL, GetCurrentThreadId());
  if (!mouse_hook)
    {
      log_error_w32 (-1, "SetWindowsHookEx failed");
      return 1;
    }  

  getc (stdin);

  log_info ("stop");
  return 0;
}


/*
Local Variables:
compile-command: "i586-mingw32msvc-gcc -Wall -g -o watchmouse.exe watchmouse.c"
End:
*/
