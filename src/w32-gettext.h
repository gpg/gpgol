/* w32-gettext.h - A simple gettext implementation for Windows targets.
   Copyright (C) 2005 g10 Code GmbH

   This file is part of libgpg-error.
 
   libgpg-error is free software; you can redistribute it and/or
   modify it under the terms of the GNU Lesser General Public License
   as published by the Free Software Foundation; either version 2.1 of
   the License, or (at your option) any later version.
 
   libgpg-error is distributed in the hope that it will be useful, but
   WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
   Lesser General Public License for more details.
 
   You should have received a copy of the GNU Lesser General Public
   License along with libgpg-error; if not, write to the Free
   Software Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA
   02111-1307, USA.  */

#if ENABLE_NLS

#include <locale.h>
#if !defined LC_MESSAGES && !(defined __LOCALE_H || (defined _LOCALE_H && defined __sun))
# define LC_MESSAGES 1729
#endif

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* Specify that the DOMAINNAME message catalog will be found
   in DIRNAME rather than in the system locale data base.  */
char *bindtextdomain (const char *domainname, const char *dirname);

const char *gettext (const char *msgid);
const char *utf8_gettext (const char *msgid);

char *textdomain (const char *domainname);

char *dgettext (const char *domainname, const char *msgid);

/* Return the localname as used by gettext.  The return value will
   never be NULL. */
const char *gettext_localename (void);

/* A pseudo function call that serves as a marker for the automated
   extraction of messages, but does not call gettext().  The run-time
   translation is done at a different place in the code.
   The argument, String, should be a literal string.  Concatenated strings
   and other string expressions won't work.
   The macro's expansion is not parenthesized, so that it is suitable as
   initializer for static 'char[]' or 'const char[]' variables.  */
#define gettext_noop(String) String


#else	/* ENABLE_NLS */

static inline const char *gettext_localename (void) { return ""; }


#endif	/* !ENABLE_NLS */


/* Conversion function. */
char *_wchar_to_utf8 (const wchar_t *string);
wchar_t *_utf8_to_wchar (const char *string);
char *utf8_to_native (const char *string);
char *native_to_utf8 (const char *string);

#define utf8_to_wchar(VAR1) ({wchar_t *retval; \
  retval = _utf8_to_wchar (VAR1); \
  if ((opt.enable_debug & (DBG_TRACE | DBG_DATA | DBG_MEMORY))) \
  { \
    log_debug ("%s:%s:%i wchar_t alloc %p:%S", \
               SRCNAME, __func__, __LINE__, retval, retval); \
  } \
retval;})

#define wchar_to_utf8(VAR1) ({char *retval; \
  retval = _wchar_to_utf8 (VAR1); \
  if ((opt.enable_debug & (DBG_TRACE | DBG_DATA | DBG_MEMORY))) \
  { \
    log_debug ("%s:%s:%i char utf8 alloc %p:%s", \
               SRCNAME, __func__, __LINE__, retval, retval); \
  } \
retval;})

#ifdef __cplusplus
}
#include <string>
std::string wchar_to_utf8_string (const wchar_t *string);
#endif
