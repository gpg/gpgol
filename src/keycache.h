/* keycache.h
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
#ifndef _GPGMEDLGS_KEYCACHE_H_
#define _GPGMEDLGS_KEYCACHE_H_

#ifdef __cplusplus
extern "C" {
#endif

struct keycache_s {
    struct keycache_s *next;
    gpgme_key_t key;
};
typedef struct keycache_s *keycache_t;

int keycache_new (keycache_t *r_ctx);
void keycache_release (keycache_t ctx);
int keycache_add (keycache_t *ctx, gpgme_key_t key);
int keycache_init (const char *pattern, int seconly, keycache_t *r_ctx);
int keycache_size (keycache_t ctx);
void keycache_free (keycache_t ctx);

keycache_t get_gpg_keycache (int sec);
int enum_gpg_seckeys (gpgme_key_t * ret_key, void **ctx);
int enum_gpg_keys (gpgme_key_t * ret_key, void **ctx);
gpgme_key_t find_gpg_key (const char *str, int type);
gpgme_key_t find_gpg_email (const char *str);
gpgme_key_t get_gpg_key (const char *str);

#ifdef __cplusplus
}
#endif

#endif /*_GPGMEDLGS_KEYCACHE_H_*/
