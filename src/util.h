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

#if __GNUC__ >= 4 
# define GPGOL_GCC_A_SENTINEL(a) __attribute__ ((sentinel(a)))
#else
# define GPGOL_GCC_A_SENTINEL(a) 
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


/* i18n stuff */
#include "w32-gettext.h"
#define _(a) gettext (a)
#define N_(a) gettext_noop (a)


/*-- common.c --*/

#include "xmalloc.h"

char *wchar_to_utf8_2 (const wchar_t *string, size_t len);
wchar_t *utf8_to_wchar2 (const char *string, size_t len);
char *latin1_to_utf8 (const char *string);

char *trim_trailing_spaces (char *string);
char *read_w32_registry_string (const char *root, const char *dir,
                                const char *name);


/*-- main.c --*/
const void *get_128bit_session_key (void);
void *create_initialization_vector (size_t nbytes);

void log_debug (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_error (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_vdebug (const char *fmt, va_list a);
void log_debug_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_error_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_hexdump (const void *buf, size_t buflen, const char *fmt, 
                  ...)  __attribute__ ((format (printf,3,4)));
void log_window_hierarchy (HWND window, const char *fmt, 
                           ...) __attribute__ ((format (printf,2,3)));

const char *log_srcname (const char *s);
#define SRCNAME log_srcname (__FILE__)
     
const char *get_log_file (void);
void set_log_file (const char *name);
void set_default_key (const char *name);
void read_options (void);
int write_options (void);


/*-- Macros to replace ctype ones to avoid locale problems. --*/
#define spacep(p)   (*(p) == ' ' || *(p) == '\t')
#define digitp(p)   (*(p) >= '0' && *(p) <= '9')
#define hexdigitp(a) (digitp (a)                     \
                      || (*(a) >= 'A' && *(a) <= 'F')  \
                      || (*(a) >= 'a' && *(a) <= 'f'))
  /* Note this isn't identical to a C locale isspace() without \f and
     \v, but works for the purposes used here. */
#define ascii_isspace(a) ((a)==' ' || (a)=='\n' || (a)=='\r' || (a)=='\t')

/* The atoi macros assume that the buffer has only valid digits. */
#define atoi_1(p)   (*(p) - '0' )
#define atoi_2(p)   ((atoi_1(p) * 10) + atoi_1((p)+1))
#define atoi_4(p)   ((atoi_2(p) * 100) + atoi_2((p)+2))
#define xtoi_1(p)   (*(p) <= '9'? (*(p)- '0'): \
                     *(p) <= 'F'? (*(p)-'A'+10):(*(p)-'a'+10))
#define xtoi_2(p)   ((xtoi_1(p) * 16) + xtoi_1((p)+1))
#define xtoi_4(p)   ((xtoi_2(p) * 256) + xtoi_2((p)+2))

#define tohex(n) ((n) < 10 ? ((n) + '0') : (((n) - 10) + 'A'))

/***** Inline functions.  ****/

/* Return true if LINE consists only of white space (up to and
   including the LF). */
static inline int
trailing_ws_p (const char *line)
{
  for ( ; *line && *line != '\n'; line++)
    if (*line != ' ' && *line != '\t' && *line != '\r')
      return 0;
  return 1;
}


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
