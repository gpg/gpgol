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
#define DBG_MEMORY         (1<<5) // 32
#define DBG_MIME_PARSER    (1<<7) // 128 Unified as DBG_DATA
#define DBG_MIME_DATA      (1<<8) // 256 Unified as DBG_DATA
#define DBG_DATA           (DBG_MIME_PARSER | DBG_MIME_DATA)
#define DBG_OOM_VAL        (1<<9) // 512 Unified as DBG_OOM
#define DBG_OOM_EXTRA      (1<<10)// 1024 Unified as DBG_OOM
#define DBG_SUPERTRACE     (1<<11)// 2048 Very verbose tracing.
#define DBG_OOM            (DBG_OOM_VAL | DBG_OOM_EXTRA)

#define debug_oom        ((opt.enable_debug & DBG_OOM) || \
                          (opt.enable_debug & DBG_OOM_EXTRA))
#define debug_oom_extra  (opt.enable_debug & DBG_OOM_EXTRA)
void log_debug (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_error (const char *fmt, ...) __attribute__ ((format (printf,1,2)));
void log_vdebug (const char *fmt, va_list a);
void log_debug_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_error_w32 (int w32err, const char *fmt,
                    ...) __attribute__ ((format (printf,2,3)));
void log_hexdump (const void *buf, size_t buflen, const char *fmt,
                  ...)  __attribute__ ((format (printf,3,4)));

#define log_oom if (opt.enable_debug & DBG_OOM) log_debug
#define log_oom if (opt.enable_debug & DBG_OOM) log_debug
#define log_mime_parser if (opt.enable_debug & DBG_DATA) log_debug
#define log_data if (opt.enable_debug & DBG_DATA) log_debug
#define log_mime_data if (opt.enable_debug & DBG_DATA) log_debug

#define gpgol_release(X) \
{ \
  if (X && opt.enable_debug & DBG_OOM_EXTRA) \
    { \
      log_debug ("%s:%s:%i: Object: %p released ref: %lu \n", \
                 SRCNAME, __func__, __LINE__, X, X->Release()); \
      memdbg_released (X); \
    } \
  else if (X) \
    { \
      X->Release(); \
    } \
}

const char *log_srcname (const char *s);
#define SRCNAME log_srcname (__FILE__)

#define TRACEPOINT log_debug ("%s:%s:%d: tracepoint\n", \
                              SRCNAME, __func__, __LINE__);
#define TSTART if (opt.enable_debug & DBG_SUPERTRACE) \
                    log_debug ("%s:%s: enter\n", SRCNAME, __func__);
#define TRETURN(X) if (opt.enable_debug & DBG_SUPERTRACE) \
                        log_debug ("%s:%s:%d: return\n", SRCNAME, __func__, \
                                   __LINE__); \
                   return X


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
