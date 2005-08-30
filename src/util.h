/* util.h - Common functions.
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifndef UTIL_H
#define UTIL_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* To avoid that a compiler optimizes certain memset calls away, these
   macros may be used instead. */
#define wipememory2(_ptr,_set,_len) do { \
              volatile char *_vptr=(volatile char *)(_ptr); \
              size_t _vlen=(_len); \
              while(_vlen) { *_vptr=(_set); _vptr++; _vlen--; } \
                  } while(0)
#define wipememory(_ptr,_len) wipememory2(_ptr,0,_len)
#define wipestring(_ptr) do { \
              volatile char *_vptr=(volatile char *)(_ptr); \
              while(*_vptr) { *_vptr=0; _vptr++; } \
                  } while(0)



/*-- common.c --*/
void* xmalloc (size_t n);
void* xcalloc (size_t m, size_t n);
char* xstrdup (const char *s);
void  xfree (void *p);
void out_of_core (void);

char *wchar_to_utf8 (const wchar_t *string);
wchar_t *utf8_to_wchar (const char *string);
wchar_t *utf8_to_wchar2 (const char *string, size_t len);

char *trim_trailing_spaces (char *string);


/*-- main.c --*/
void log_debug (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_error (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_vdebug (const char *fmt, va_list a);
void log_debug_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_error_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_hexdump (const void *buf, size_t buflen, const char *fmt, 
                  ...)  __attribute__ ((format (printf,3,4)));
     
const char *get_log_file (void);
void set_log_file (const char *name);
void set_default_key (const char *name);
void read_options (void);
int write_options (void);


/*****  Missing functions.  ****/

#ifndef HAVE_STPCPY
static inline char *
_gpgol_stpcpy (char *a, const char *b)
{
  while (*b)
    *a++ = *b++;
  *a = 0;
  return a;
}
#define stpcpy(a,b) _gpgol_stpcpy ((a), (b))
#endif /*!HAVE_STPCPY*/



#ifdef __cplusplus
}
#endif
#endif /*UTIL_H*/
