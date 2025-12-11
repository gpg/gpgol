/* common.h - Common declarations for GpgOL
 * Copyright (C) 2004 Timo Schulz
 * Copyright (C) 2005, 2006, 2007, 2008 g10 Code GmbH
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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

#ifndef GPGOL_COMMON_H
#define GPGOL_COMMON_H

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <gpgme.h>

#include "common_indep.h"

#include <windows.h>

/* i18n stuff */
#include "w32-gettext.h"
#define _(a) w32_gettext (a)
#define N_(a) gettext_noop (a)


/* Registry path to store plugin settings */
#define GPGOL_REGPATH "Software\\GNU\\GpgOL"

#ifdef __cplusplus
#include <string.h>
std::string readRegStr (const char *root, const char *dir, const char *name);

namespace GpgME
{
  class Data;
} // namespace GpgME

/* Read a file into a GpgME::Data object returns 0 on success */
int readFullFile (HANDLE hFile, GpgME::Data &data);

extern "C" {
#if 0
}
#endif
#endif
extern HINSTANCE glob_hinst;
extern UINT      this_dll;

/*-- common.cpp --*/
char *get_data_dir (void);
const char *get_gpg4win_dir (void);

int store_extension_value (const char *key, const char *val);
int store_extension_subkey_value (const char *subkey, const char *key,
                                  const char *val);
int load_extension_value (const char *key, char **val);

/* Get a temporary filename with and its name */
wchar_t *get_tmp_outfile (const wchar_t *name, HANDLE *outHandle);
char *get_tmp_outfile_utf8 (const char *name, HANDLE *outHandle);

wchar_t *get_pretty_attachment_name (wchar_t *path, protocol_t protocol,
                                     int signature);

/*-- verify-dialog.c --*/
int verify_dialog_box (gpgme_protocol_t protocol,
                       gpgme_verify_result_t res,
                       const char *filename);


/*-- inspectors.cpp --*/
int initialize_inspectors (void);

#if __GNUC__ >= 4
# define GPGOL_GCC_A_SENTINEL(a) __attribute__ ((sentinel(a)))
#else
# define GPGOL_GCC_A_SENTINEL(a)
#endif


/*-- common.c --*/

void fatal_error (const char *format, ...);

char *read_w32_registry_string (const char *root, const char *dir,
                                const char *name);
char *percent_escape (const char *str, const char *extra);

void fix_linebreaks (char *str, int *len);

/* Format a date from gpgme (seconds since epoch)
   with windows system locale. */
char *format_date_from_gpgme (unsigned long time);

/* Get the name of the uiserver */
char *get_uiserver_name (void);

int is_elevated (void);

/*-- main.c --*/
void read_options (void);
int write_options (void);

extern int g_ol_version_major;

void bring_to_front (HWND wid);

int gpgol_message_box (HWND parent, const char *utf8_text,
                       const char *utf8_caption, UINT type);

/* Show a bug message with the code. */
void gpgol_bug (HWND parent, int code);

void i18n_init (void);
#define ERR_CRYPT_RESOLVER_FAILED 1
#define ERR_WANTS_SEND_MIME_BODY 2
#define ERR_WANTS_SEND_INLINE_BODY 3
#define ERR_INLINE_BODY_TO_BODY 4
#define ERR_INLINE_BODY_INV_STATE 5
#define ERR_SEND_FALLBACK_FAILED 6
#define ERR_GET_BASE_MSG_FAILED 7
#define ERR_SPLIT_UNEXPECTED 8
#define ERR_SPLIT_RECIPIENTS 9
#ifdef __cplusplus
}

/* Check if we are in de_vs mode. */
bool in_de_vs_mode ();
/* Get the name of the de_vs compliance mode as configured
   in libkleopatrarc */
const char *de_vs_name (bool isCompliant = false);

/* Get a localized string of the compliance of an operation
   mode is the usual 1 encrypt, 2 sign. isCompliant if the value
   of the operation was compliant or not. */
std::string
compliance_string (bool forVerify, bool forDecrypt, bool isCompliant);

HANDLE CreateFileUtf8 (const char *utf8Name);

/* Read a registry value from HKLM or HKCU of type dword and
   return 1 if the value is larger then 0, 0 if it is zero and
   -1 if it is not set or not found. */
int
read_reg_bool (HKEY root, const char *path, const char *value);

/* Remove characters which may not be a part of a filename from
   the string @name. If @allowDirectories is set the rules
   for directories are used. */
std::string sanitizeFileName(std::string name,
                             bool allowDirectories = false);

/* Check if the filename @name can be used on Windows. If @allowDirectories
   is set the check is for directory names. */
bool validateWindowsFileName(const std::string& name,
                             bool allowDirectories = false);
#endif
#endif /*GPGOL_COMMON_H*/
