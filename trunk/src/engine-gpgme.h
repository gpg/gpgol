/* engine-gpgme.h - GPGME based crypto engine
 *	Copyright (C) 2005, 2007 g10 Code GmbH
 *
 * This file is part of Gpgol.
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

#ifndef GPGOL_ENGINE_GPGME_H
#define GPGOL_ENGINE_GPGME_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#include <gpgme.h>

typedef engine_sigtype_t op_sigtype_t;


int op_gpgme_basic_init (void);
int op_gpgme_init (void);
void op_gpgme_deinit (void);
void engine_gpgme_cancel (void *cancel_data);


int op_gpgme_encrypt (protocol_t protocol,
                      gpgme_data_t indata, gpgme_data_t outdata, 
                      engine_filter_t filter, void *hwnd,
                      char **recipients);
int op_gpgme_sign (protocol_t protocol, 
                   gpgme_data_t indata, gpgme_data_t outdata,
                   engine_filter_t filter, void *hwnd);
int op_gpgme_decrypt (protocol_t protocol,
                      gpgme_data_t indata, gpgme_data_t outdata, 
                      engine_filter_t filter, void *hwnd,
                      int with_verify);
int op_gpgme_verify (gpgme_protocol_t protocol, 
                     gpgme_data_t data, const char *signature, size_t sig_len,
                     engine_filter_t filter, void *hwnd);




int op_export_keys (const char *pattern[], const char *outfile);

int op_lookup_keys (char **names, gpgme_key_t **keys, char ***unknown);
gpgme_key_t op_get_one_key (char *pattern);

const char *userid_from_key (gpgme_key_t k);
const char *keyid_from_key (gpgme_key_t k);

#ifdef __cplusplus
}
#endif
#endif /*GPGOL_ENGINE_GPGME_H*/
