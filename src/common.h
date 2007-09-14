/* common.h - Common declarations for GpgOL
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005, 2006, 2007 g10 Code GmbH
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
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
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
  int passwd_ttl;            /* Time in seconds the passphrase is stored. */
  int encrypt_default;       /* Encrypt by default. */
  int sign_default;          /* Sign by default. */
  int save_decrypted_attach; /* Save decrypted attachments. */
  int auto_sign_attach;	     /* Sign all outgoing attachments. */
  int enc_format;            /* Encryption format for attachments. */
  char *default_key;         /* The key we want to always encrypt to. */
  int enable_default_key;    /* Enable the use of DEFAULT_KEY. */
  int preview_decrypt;       /* Decrypt in preview window. */
  int prefer_html;           /* Prefer html in html/text alternatives. */

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
  } compat; 
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



/*-- common.c --*/
void set_global_hinstance (HINSTANCE hinst);
void center_window (HWND childwnd, HWND style);
char *get_save_filename (HWND root, const char *srcname);
char *utf8_to_wincp (const char *string);

HRESULT w32_shgetfolderpath (HWND a, int b, HANDLE c, DWORD d, LPSTR e);

size_t qp_decode (char *buffer, size_t length);
void b64_init (b64_state_t *state);
size_t b64_decode (b64_state_t *state, char *buffer, size_t length);

/* The length of the boundary - the buffer needs to be allocated one
   byte larger. */
#define BOUNDARYSIZE 20
char *generate_boundary (char *buffer);


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
int start_key_manager (void);
int store_extension_value (const char *key, const char *val);
int load_extension_value (const char *key, char **val);

/*-- verify-dialog.c --*/
int verify_dialog_box (gpgme_protocol_t protocol, 
                       gpgme_verify_result_t res, 
                       const char *filename);


#ifdef __cplusplus
}
#endif
#endif /*GPGOL_COMMON_H*/
