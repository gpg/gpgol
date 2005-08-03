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

#ifndef _GPGMEDLGS_ENGINE_H
#define _GPGMEDLGS_ENGINE_H 1

#ifdef __cplusplus
extern "C" {
#endif

typedef enum {
    OP_SIG_NORMAL = 0,
    OP_SIG_DETACH = 1,
    OP_SIG_CLEAR  = 2
} op_sigtype_t;


int op_init (void);
void op_deinit (void);
const char* op_strerror (int err);

#define op_debug_enable(file) op_set_debug_mode (5, (file))
#define op_debug_disable() op_set_debug_mode (0, NULL)
void op_set_debug_mode (int val, const char *file);

int op_encrypt_start (const char *inbuf, char **outbuf);
int op_encrypt (void *rset, const char *inbuf, char **outbuf);
int op_encrypt_file (void *rset, const char *infile, const char *outfile);

int op_sign_encrypt_start (const char *inbuf, char **outbuf);
int op_sign_encrypt (void *rset, void *locusr, const char *inbuf, 
		     char **outbuf);
int op_sign_encrypt_file (void *rset, const char *infile, const char *outfile);

int op_verify_start (const char *inbuf, char **outbuf);

int op_sign_start (const char *inbuf, char **outbuf);
int op_sign (void *locusr, const char *inbuf, char **outbuf);
int op_sign_file (int mode, const char *infile, const char *outfile);
int op_sign_file_ext (int mode, const char *infile, const char *outfile,
		      cache_item_t *ret_itm);
int op_sign_file_next (int (*pass_cb)(void *, const char*, const char*, int, int),
	               void *pass_cb_value,
	               int mode, const char *infile, const char *outfile);

int op_decrypt_file (const char *infile, const char *outfile);
int op_decrypt_next (int (*pass_cb)(void *, const char*, const char*, int, int),
		     void *pass_cb_value,
		     const char *inbuf, char **outbuf);
int op_decrypt_start (const char *inbuf, char **outbuf);
int op_decrypt_start_ext (const char *inbuf, char **outbuf, cache_item_t *ret_itm);

int op_lookup_keys (char **id, gpgme_key_t **keys, char ***unknown, size_t *n);

int op_export_keys (const char *pattern[], const char *outfile);


#ifdef __cplusplus
}
#endif


#endif /*_GPGMEDLGS_ENGINE_H*/
