/* intern.h
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGME Dialogs.
 *
 * GPGME Dialogs is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1
 * of the License, or (at your option) any later version.
 *  
 * GPGME Dialogs is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGME Dialogs; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */
#ifndef _GPGMEDLGS_INTERN_H_
#define _GPGMEDLGS_INTERN_H_

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif


#ifndef STRICT
#define STRICT
#endif

extern HINSTANCE glob_hinst;
extern UINT      this_dll;

enum {
    OPT_FLAG_ARMOR    =  1,
    OPT_FLAG_TEXT     =  2,
    OPT_FLAG_FORCE    =  4,
    OPT_FLAG_CANCEL   =  8,
};


struct decrypt_key_s {
    gpgme_key_t signer;
    char keyid[16+1];
    char *user_id;
    char *pass;    
    void *ctx;
    int opts;
    unsigned flags;
    unsigned int hide_pwd:1;
    unsigned int use_as_cb:1;
    unsigned int last_was_bad:1;
};

struct cache_item_s {
    char keyid[16+1];
    char *pass;
    unsigned ttl;
};
typedef struct cache_item_s *cache_item_t;

/*-- common.c --*/
void set_global_hinstance (HINSTANCE hinst);
void center_window (HWND childwnd, HWND style);

void* xmalloc (size_t n);
void* xcalloc (size_t m, size_t n);
char* xstrdup (const char *s);
void  xfree (void *p);

cache_item_t cache_item_new (void);
void cache_item_free (cache_item_t itm);

HRESULT w32_shgetfolderpath (HWND a, int b, HANDLE c, DWORD d, LPSTR e);

/*-- recipient-dialog.c --*/
int recipient_dialog_box(gpgme_key_t **ret_rset, int *ret_opts);
int recipient_dialog_box2 (gpgme_key_t *fnd, char **unknown, size_t n,
		           gpgme_key_t **ret_rset, int *ret_opts);

/*-- passphrase-dialog.c --*/
int signer_dialog_box (gpgme_key_t *r_key, char **r_passwd);
int passphrase_callback_box (void *opaque, const char *uid_hint, 
			     const char *pass_info,
			     int prev_was_bad, int fd);
void free_decrypt_key (struct decrypt_key_s * ctx);

/*-- config-dialog.c --*/
void config_dialog_box (HWND parent);
int start_key_manager (void);
int store_extension_value (const char *key, const char *val);
int load_extension_value (const char *key, char **val);

/*-- verify-dialog.c --*/
int verify_dialog_box (gpgme_verify_result_t res);

/*-- keycache.c --*/
void load_keycache_objects (keycache_t ring[2]);
void init_keycache_objects (void);
void cleanup_keycache_objects (void);
void reset_gpg_seckeys (void **ctx);

#ifdef __cplusplus
}
#endif

#endif /*_GPGMEDLGS_INTERN_H_*/
