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

#ifndef _GPGME_DLGS_PRIVATE_H_
#define _GPGME_DLGS_PRIVATE_H_

#ifndef STRICT
#define STRICT
#endif

extern HINSTANCE glob_hinst;
extern UINT      this_dll;

enum {
    OPT_FLAG_ARMOR    =  1,
    OPT_FLAG_TEXT     =  2,
    OPT_FLAG_DETACHED =  4,
    OPT_FLAG_SYMETRIC =  8,
    OPT_FLAG_FORCE    = 16,
    OPT_FLAG_CANCEL   = 32,
};


struct decrypt_key_s {
    char keyid[16+1];
    char * user_id;
    char * pass;
    gpgme_key_t signer;
    int opts;
    unsigned int use_as_cb:1;
};


int recipient_dialog_box(gpgme_key_t **ret_rset, int *ret_opts);
int signer_dialog_box(gpgme_key_t *r_key, char **passwd);


const char * passphrase_callback_box(void *opaque, const char *uid_hint, 
				     const char *pass_info,
				     int prev_was_bad, int fd);
void free_decrypt_key( struct decrypt_key_s * ctx );

void cleanup_keycache_objects(void);
void reset_gpg_seckeys(void **ctx);

#endif /*_GPGME_DLGS_PRIVATE_H_*/