/* xmalloc.h - xmalloc prototypes
 * Copyright (C) 2006 g10 Code GmbH
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

#ifndef XMALLOC_H
#define XMALLOC_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/*-- common.c --*/
#define xmalloc(VAR1) ({void *retval; \
  retval = _xmalloc(VAR1); \
  if ((opt.enable_debug & DBG_MEMORY)) \
  { \
    memdbg_alloc (retval); \
    if ((opt.enable_debug & DBG_TRACE)) \
      memset (retval, 'X', VAR1); \
  } \
retval;})

#define xcalloc(VAR1, VAR2) ({void *retval; \
  retval = _xcalloc(VAR1, VAR2); \
  if ((opt.enable_debug & DBG_MEMORY)) \
  { \
    memdbg_alloc (retval);\
  } \
retval;})

#define xrealloc(VAR1, VAR2) ({void *retval; \
  retval = _xrealloc (VAR1, VAR2); \
  if ((opt.enable_debug & DBG_MEMORY)) \
  { \
    memdbg_alloc (retval);\
    memdbg_free ((void*)VAR1); \
  } \
retval;})

#define xfree(VAR1) \
{ \
  if (VAR1 && (opt.enable_debug & DBG_MEMORY) && !memdbg_free (VAR1)) \
    log_debug ("%s:%s:%i %p freed here", \
               log_srcname (__FILE__), __func__, __LINE__, VAR1); \
  _xfree (VAR1); \
}

#define xstrdup(VAR1) ({char *retval; \
  retval = _xstrdup (VAR1); \
  if ((opt.enable_debug & DBG_MEMORY)) \
  { \
    memdbg_alloc ((void *)retval);\
  } \
retval;})

#define xwcsdup(VAR1) ({wchar_t *retval; \
  retval = _xwcsdup (VAR1); \
  if ((opt.enable_debug & DBG_MEMORY)) \
  { \
    memdbg_alloc ((void *)retval);\
  } \
retval;})

void* _xmalloc (size_t n);
void* _xcalloc (size_t m, size_t n);
void *_xrealloc (void *a, size_t n);
char* _xstrdup (const char *s);
wchar_t * _xwcsdup (const wchar_t *s);
void  _xfree (void *p);
void out_of_core (void);

#ifdef __cplusplus
}
#endif
#endif /*XMALLOC_H*/
