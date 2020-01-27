#ifndef DEBUG_H
#define DEBUG_H
/* debug.h - Debugging / Log helpers for GpgOL
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif
#include "common_indep.h"

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* Bit values used for extra log file verbosity.  Value 1 is reserved
   to enable debug menu options.

   Note that the high values here are used for compatibility with
   old howtos of how to enable debug flags. Based on the old
   very split up logging categories.

   Categories are meant to be:

   DBG -> Generally useful information.
   DBG_MEMORY -> Very verbose tracing of Releases / Allocs / Refs.
   DBG_OOM -> Outlook Object Model events tracing.
   DBG_DATA -> Including potentially private data and mime parser logging.

   Common values are:
   32 -> Only memory debugging.
   544 -> OOM and Memory.
   800 -> Full debugging.
   */
#define DBG_OOM            (1<<1) // 2
#define DBG_MEMORY         (1<<2) // 4
#define DBG_TRACE          (1<<3) // 8
#define DBG_DATA           (1<<4) // 16

void log_debug (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_error (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_vdebug (const char *fmt, va_list a);
void log_debug_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_error_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_hexdump (const void *buf, size_t buflen, const char *fmt,
                  ...)  __attribute__ ((format (printf,3,4)));

const char *anonstr (const char *data);
#define log_oom(format, ...) if ((opt.enable_debug & DBG_OOM)) \
  log_debug("DBG_OOM/" format, ##__VA_ARGS__)

#define log_data(format, ...) if ((opt.enable_debug & DBG_DATA)) \
  log_debug("DBG_DATA/" format, ##__VA_ARGS__)

#define log_memory(format, ...) if ((opt.enable_debug & DBG_MEMORY)) \
  log_debug("DBG_MEM/" format, ##__VA_ARGS__)

#define log_trace(format, ...) if ((opt.enable_debug & DBG_TRACE)) \
  log_debug("TRACE/" format, ##__VA_ARGS__)

#define log_warn(format, ...) if (opt.enable_debug) \
  log_debug("WARNING/" format, ##__VA_ARGS__)

#define gpgol_release(X) \
{ \
  if (X && opt.enable_debug & DBG_MEMORY) \
    { \
      log_memory ("%s:%s:%i: Object: %p released ref: %lu \n", \
                  SRCNAME, __func__, __LINE__, X, X->Release()); \
      memdbg_released (X); \
    } \
  else if (X) \
    { \
      X->Release(); \
    } \
}


#define gpgol_lock(X) \
{ \
  if (opt.enable_debug & DBG_TRACE) \
    { \
      log_trace ("%s:%s:%i: lock %p lock", \
                  SRCNAME, __func__, __LINE__, X); \
    } \
  gpgrt_lock_lock(X); \
}


#define gpgol_unlock(X) \
{ \
  if (opt.enable_debug & DBG_TRACE) \
    { \
      log_trace ("%s:%s:%i: lock %p unlock.", \
                  SRCNAME, __func__, __LINE__, X); \
    } \
  gpgrt_lock_unlock(X); \
}

const char *log_srcname (const char *s);
#define SRCNAME log_srcname (__FILE__)

#define STRANGEPOINT log_debug ("%s:%s:%d:UNEXPECTED", \
                           SRCNAME, __func__, __LINE__);
#define TRACEPOINT log_trace ("%s:%s:%d", \
                              SRCNAME, __func__, __LINE__);
#define TSTART log_trace ("%s:%s:%d enter", SRCNAME, __func__, __LINE__);
#define TRETURN log_trace ("%s:%s:%d: return", SRCNAME, __func__, \
                           __LINE__); \
                   return
#define TBREAK log_trace ("%s:%s:%d: break", SRCNAME, __func__, \
                           __LINE__); \
                   break


const char *get_log_file (void);
void set_log_file (const char *name);

#ifdef _WIN64
#define SIZE_T_FORMAT "%I64u"
#else
# ifdef HAVE_W32_SYSTEM
#  define SIZE_T_FORMAT "%u"
# else
#  define SIZE_T_FORMAT "%lu"
# endif
#endif

#ifdef __cplusplus
}
#endif

#endif // DEBUG_H
