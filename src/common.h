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

#include <gpgme.h>

#include "util.h"


#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* The Registry key used by GnuPg and closley related software.  */
#define GNUPG_REGKEY  "Software\\GNU\\GnuPG"


/* Identifiers for the protocol.  We use different one than those use
   by gpgme.  FIXME: We might want to define an unknown protocol to
   non-null and define such a value also in gpgme. */
typedef enum
  {
    PROTOCOL_UNKNOWN = 0,
    PROTOCOL_OPENPGP = 1000,
    PROTOCOL_SMIME   = 1001
  }
protocol_t;


/* Possible options for the recipient dialog. */
enum
  {
    OPT_FLAG_TEXT     =  2,
    OPT_FLAG_FORCE    =  4,
    OPT_FLAG_CANCEL   =  8
  };


typedef enum
  {
    GPG_FMT_NONE = 0,       /* do not encrypt attachments */
    GPG_FMT_CLASSIC = 1,    /* encrypt attachments without any encoding */
    GPG_FMT_PGP_PEF = 2     /* use the PGP partioned encoding format (PEF) */
  }
gpgol_format_t;

/* Type of a message. */
typedef enum
  {
    OPENPGP_NONE = 0,
    OPENPGP_MSG,
    OPENPGP_SIG,
    OPENPGP_CLEARSIG,
    OPENPGP_PUBKEY,   /* Note, that this type is only partly supported */
    OPENPGP_SECKEY    /* Note, that this type is only partly supported */
  }
openpgp_t;


extern HINSTANCE glob_hinst;
extern UINT      this_dll;


/* Passphrase callback structure. */
struct passphrase_cb_s
{
  gpgme_key_t signer;
  gpgme_ctx_t ctx;
  char keyid[16+1];
  char *user_id;
  char *pass;
  int opts;
  int ttl;  /* TTL of the passphrase. */
  unsigned int decrypt_cmd:1; /* 1 = show decrypt dialog, otherwise secret key
			         selection. */
  unsigned int hide_pwd:1;
  unsigned int last_was_bad:1;
};

/* Global options - initialized to default by main.c. */
#ifdef __cplusplus
extern
#endif
struct
{
  int enable_debug;	     /* Enable extra debug options.  Values
                                larger than 1 increases the debug log
                                verbosity.  */
  int enable_smime;	     /* Enable S/MIME support. */
  int passwd_ttl;            /* Time in seconds the passphrase is stored. */
  protocol_t default_protocol;/* The default protocol. */
  int encrypt_default;       /* Encrypt by default. */
  int sign_default;          /* Sign by default. */
  int enc_format;            /* Encryption format for attachments. */
  char *default_key;         /* The key we want to always encrypt to. */
  int enable_default_key;    /* Enable the use of DEFAULT_KEY. */
  int preview_decrypt;       /* Decrypt in preview window. */
  int prefer_html;           /* Prefer html in html/text alternatives. */
  int body_as_attachment;    /* Present encrypted message as attachment.  */

  /* The compatibility flags. */
  struct
  {
    unsigned int no_msgcache:1;
    unsigned int no_pgpmime:1;
    unsigned int no_oom_write:1; /* Don't write using Outlooks object model. */
    unsigned int no_preview_info:1; /* No preview info about PGP/MIME. */
    unsigned int old_reply_hack: 1; /* See gpgmsg.cpp:decrypt. */
    unsigned int auto_decrypt: 1;   /* Try to decrypt when clicked. */
    unsigned int no_attestation: 1; /* Don't create an attestation. */
    unsigned int use_mwfmo: 1;      /* Use MsgWaitForMultipleObjects.  */
  } compat;

  /* The current git commit id.  */
  unsigned int git_commit;

  /* The forms revision number of the binary.  */
  int forms_revision;

  /* The stored number of the binary which showed the last announcement.  */
  int announce_number;

  /* Disable message processing until restart.  This is required to
     implement message reverting as a perparation to remove GpgOL.  */
  int disable_gpgol;

} opt;


/* The state object used by b64_decode.  */
struct b64_state_s
{
  int idx;
  unsigned char val;
  int stop_seen;
  int invalid_encoding;
};
typedef struct b64_state_s b64_state_t;

/* Bit values used for extra log file verbosity.  Value 1 is reserved
   to enable debug menu options.  */
#define DBG_IOWORKER       (1<<1)
#define DBG_IOWORKER_EXTRA (1<<2)
#define DBG_FILTER         (1<<3)
#define DBG_FILTER_EXTRA   (1<<4)
#define DBG_MEMORY         (1<<5)
#define DBG_COMMANDS       (1<<6)
#define DBG_MIME_PARSER    (1<<7)
#define DBG_MIME_DATA      (1<<8)
#define DBG_OOM            (1<<9)
#define DBG_OOM_EXTRA      (1<<10)

/* Macros to used in conditionals to enable debug output.  */
#define debug_commands    (opt.enable_debug & DBG_COMMANDS)

/*-- common.c --*/
void set_global_hinstance (HINSTANCE hinst);
void center_window (HWND childwnd, HWND style);
HBITMAP get_system_check_bitmap (int checked);
char *get_save_filename (HWND root, const char *srcname);
char *get_open_filename (HWND root, const char *title);
char *utf8_to_wincp (const char *string);

const char *default_homedir (void);
char *get_data_dir (void);

size_t qp_decode (char *buffer, size_t length, int *r_slbrk);
void b64_init (b64_state_t *state);
size_t b64_decode (b64_state_t *state, char *buffer, size_t length);

/* Get a temporary filename with and its name */
wchar_t *get_tmp_outfile (wchar_t *name, HANDLE *outHandle);

wchar_t *get_pretty_attachment_name (wchar_t *path, protocol_t protocol,
                                     int signature);

/* The length of the boundary - the buffer needs to be allocated one
   byte larger. */
#define BOUNDARYSIZE 20
char *generate_boundary (char *buffer);

int gpgol_spawn_detached (const char *cmdline);

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


#ifdef __cplusplus
}
#endif
#endif /*GPGOL_COMMON_H*/
