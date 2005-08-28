/* engine.h - Crypto engine
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

#ifndef ENGINE_H
#define ENGINE_H 1

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#include <gpgme.h>

typedef enum 
  {
    OP_SIG_NORMAL = 0,
    OP_SIG_DETACH = 1,
    OP_SIG_CLEAR  = 2
  } 
op_sigtype_t;


int op_init (void);
void op_deinit (void);

#define op_debug_enable(file) op_set_debug_mode (5, (file))
#define op_debug_disable() op_set_debug_mode (0, NULL)
void op_set_debug_mode (int val, const char *file);

int op_encrypt (const char *inbuf, char **outbuf,
                gpgme_key_t *keys, gpgme_key_t sign_key, int ttl);
int op_encrypt_stream (LPSTREAM instream, LPSTREAM outstream,
                       gpgme_key_t *keys, gpgme_key_t sign_key, int ttl);

int op_sign (const char *inbuf, char **outbuf, int mode,
             gpgme_key_t sign_key, int ttl);
int op_sign_stream (LPSTREAM instream, LPSTREAM outstream, int mode,
                    gpgme_key_t sign_key, int ttl);

int op_decrypt (const char *inbuf, char **outbuf, int ttl);
int op_decrypt_stream (LPSTREAM instream, LPSTREAM outstream, int ttl);

int op_verify (const char *inbuf, char **outbuf);
int op_verify_detached_sig (LPSTREAM data, const char *sig,
                            const char *filename);


int op_export_keys (const char *pattern[], const char *outfile);

int op_lookup_keys (char **id, gpgme_key_t **keys, char ***unknown, size_t *n);

const char *userid_from_key (gpgme_key_t k);
const char *keyid_from_key (gpgme_key_t k);

const char* op_strerror (int err);


#ifdef __cplusplus
}
#endif


#endif /*ENGINE_H*/
