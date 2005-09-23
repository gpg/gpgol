/* intern.h
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 *
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1
 * of the License, or (at your option) any later version.
 *  
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGol; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */
#ifndef GPGOL_INTERN_H
#define GPGOL_INTERN_H

#include <gpgme.h>

#include "util.h"
 

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#ifndef STRICT
#define STRICT
#endif


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


struct decrypt_key_s 
{
  gpgme_key_t signer;
  char keyid[16+1];
  char *user_id;
  char *pass;    
  void *ctx;
  int opts;
  int ttl;  /* TTL of the passphrase. */
  unsigned int flags;
  unsigned int hide_pwd:1;
  unsigned int use_as_cb:1;
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
  char *default_key;         /* Malloced default key or NULL. */
  int add_default_key;       /* Always also encrypt to the default key. */

  /* The compatibility flags. */
  struct 
  {
    unsigned int no_msgcache:1;
    unsigned int no_pgpmime:1;
    unsigned int no_oom_write:1; /* Don't write using Outlooks object model. */
    unsigned int preview_decryption:1; /* Decrypt in preview window. */
    unsigned int old_reply_hack: 1; /* See gpgmsg.cpp:decrypt. */
    unsigned int auto_decrypt: 1;   /* Try to decrypt when clicked. */
    unsigned int no_attestation: 1; /* Don't create an attestation. */
  } compat; 
} opt;



/*-- common.c --*/
void set_global_hinstance (HINSTANCE hinst);
void center_window (HWND childwnd, HWND style);
char *get_save_filename (HWND root, const char *srcname);

HRESULT w32_shgetfolderpath (HWND a, int b, HANDLE c, DWORD d, LPSTR e);

/*-- watcher.cpp --*/
int watcher_init_hook (void);
int watcher_free_hook (void);
void watcher_set_callback_ctx (void *cb);

/*-- recipient-dialog.c --*/
int recipient_dialog_box(gpgme_key_t **ret_rset, int *ret_opts);
int recipient_dialog_box2 (gpgme_key_t *fnd, char **unknown, size_t n,
		           gpgme_key_t **ret_rset, int *ret_opts);

/*-- passphrase-dialog.c --*/
int signer_dialog_box (gpgme_key_t *r_key, char **r_passwd);
gpgme_error_t passphrase_callback_box (void *opaque, const char *uid_hint, 
			     const char *pass_info,
			     int prev_was_bad, int fd);
void free_decrypt_key (struct decrypt_key_s * ctx);
const char *get_pubkey_algo_str (gpgme_pubkey_algo_t id);

/*-- config-dialog.c --*/
void config_dialog_box (HWND parent);
int start_key_manager (void);
int store_extension_value (const char *key, const char *val);
int load_extension_value (const char *key, char **val);

/*-- verify-dialog.c --*/
int verify_dialog_box (gpgme_verify_result_t res, const char *filename);

#ifdef __cplusplus
}
#endif

#endif /*GPGOL_INTERN_H*/
