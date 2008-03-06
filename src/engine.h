/* engine.h - Crypto engine dispatcher
 *	Copyright (C) 2007 g10 Code GmbH
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
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifndef GPGOL_ENGINE_H
#define GPGOL_ENGINE_H 1

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

typedef enum 
  {
    OP_SIG_NORMAL = 0,
    OP_SIG_DETACH = 1,
    OP_SIG_CLEAR  = 2
  } 
engine_sigtype_t;


/* The key info object.  */
struct engine_keyinfo_s;
typedef struct engine_keyinfo_s *engine_keyinfo_t;

/* The filter object.  */
struct engine_filter_s;
typedef struct engine_filter_s *engine_filter_t;




/*-- engine.c -- */
int engine_init (void);
void engine_deinit (void);

void engine_private_set_cancel (engine_filter_t filter, void *cancel_data);
void engine_private_finished (engine_filter_t filter, gpg_error_t status);

int engine_filter (engine_filter_t filter,
                   const void *indata, size_t indatalen);
int engine_create_filter (engine_filter_t *r_filter,
                          int (*outfnc) (void *, const void *, size_t),
                          void *outfncdata);
int engine_wait (engine_filter_t filter);
void engine_cancel (engine_filter_t filter);

int engine_encrypt_start (engine_filter_t filter, HWND hwnd,
                          protocol_t req_protocol, char **recipients,
                          protocol_t *r_protocol);
int engine_sign_start (engine_filter_t filter, HWND hwnd, protocol_t protocol,
                       const char *sender, protocol_t *r_protocol);

int engine_decrypt_start (engine_filter_t filter, HWND hwnd,
                          protocol_t protocol, int with_verify);
int engine_verify_start (engine_filter_t filter, HWND hwnd,
                         const char *signature, size_t sig_len,
			 protocol_t protocol);

int engine_start_keymanager (HWND hwnd);



#ifdef __cplusplus
}
#endif
#endif /*GPGOL_ENGINE_H*/
