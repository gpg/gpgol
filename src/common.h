/* common.h - Common declarations for GpgOL
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005, 2006, 2007, 2008 g10 Code GmbH
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

#ifdef MIME_SEND
# define MIME_UI_DEFAULT 1
#else
# define MIME_UI_DEFAULT 0
#endif

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif
extern HINSTANCE glob_hinst;
extern UINT      this_dll;

/*-- common.c --*/
void set_global_hinstance (HINSTANCE hinst);
void center_window (HWND childwnd, HWND style);
HBITMAP get_system_check_bitmap (int checked);
char *get_save_filename (HWND root, const char *srcname);
char *get_open_filename (HWND root, const char *title);
char *utf8_to_wincp (const char *string);

const char *default_homedir (void);
char *get_data_dir (void);
char *get_gpg4win_dir (void);

/* Get a temporary filename with and its name */
wchar_t *get_tmp_outfile (wchar_t *name, HANDLE *outHandle);

wchar_t *get_pretty_attachment_name (wchar_t *path, protocol_t protocol,
                                     int signature);

/*-- recipient-dialog.c --*/
unsigned int recipient_dialog_box (gpgme_key_t **ret_rset);
unsigned int recipient_dialog_box2 (gpgme_key_t *fnd, char **unknown,
                                    gpgme_key_t **ret_rset);

/*-- passphrase-dialog.c --*/
int signer_dialog_box (gpgme_key_t *r_key, char **r_passwd, int encrypting);
gpgme_error_t passphrase_callback_box (void *opaque, const char *uid_hint,
			     const char *pass_info,
			     int prev_was_bad, int fd);
void free_decrypt_key (struct passphrase_cb_s *ctx);
const char *get_pubkey_algo_str (gpgme_pubkey_algo_t id);

/*-- config-dialog.c --*/
void config_dialog_box (HWND parent);
int store_extension_value (const char *key, const char *val);
int load_extension_value (const char *key, char **val);

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


/* i18n stuff */
#include "w32-gettext.h"
#define _(a) gettext (a)
#define N_(a) gettext_noop (a)


/*-- common.c --*/

void fatal_error (const char *format, ...);

char *wchar_to_utf8_2 (const wchar_t *string, size_t len);
wchar_t *utf8_to_wchar2 (const char *string, size_t len);
char *read_w32_registry_string (const char *root, const char *dir,
                                const char *name);
char *percent_escape (const char *str, const char *extra);

void fix_linebreaks (char *str, int *len);

/*-- main.c --*/
const void *get_128bit_session_key (void);
const void *get_64bit_session_marker (void);
void *create_initialization_vector (size_t nbytes);

void read_options (void);
int write_options (void);

extern int g_ol_version_major;

void log_window_hierarchy (HWND window, const char *fmt,
                           ...) __attribute__ ((format (printf,2,3)));

#ifdef __cplusplus
}
#endif
#endif /*GPGOL_COMMON_H*/
