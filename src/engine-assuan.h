/* engine-assuan.h - Assuan server based crypto engine
 *	Copyright (C) 2007, 2008 g10 Code GmbH
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

#ifndef GPGOL_ENGINE_ASSUAN_H
#define GPGOL_ENGINE_ASSUAN_H

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#include <gpgme.h>  /* We need it for gpgme_data_t.  */

#include "engine.h"

struct engine_assuan_encstate_s;

int  op_assuan_init (void);
void op_assuan_deinit (void);
void engine_assuan_cancel (void *cancel_data);

int op_assuan_encrypt (protocol_t protocol, 
                       gpgme_data_t indata, gpgme_data_t outdata,
                       engine_filter_t notify_data, void *hwnd,
                       unsigned int flags,
                       const char *sender, char **recipients,
                       protocol_t *r_used_protocol,
                       struct engine_assuan_encstate_s **r_encstate);
int op_assuan_encrypt_bottom (struct engine_assuan_encstate_s *encstate,
                              int cancel);
int op_assuan_sign (protocol_t protocol, 
                    gpgme_data_t indata, gpgme_data_t outdata,
                    engine_filter_t filter, void *hwnd,
                    const char *sender, protocol_t *r_used_protocol);
int op_assuan_decrypt (protocol_t protocol,
                       gpgme_data_t indata, gpgme_data_t outdata, 
                       engine_filter_t filter, void *hwnd,
                       int with_verify, const char *from_address);
int op_assuan_verify (gpgme_protocol_t protocol, 
                      gpgme_data_t data, const char *signature, size_t sig_len,
                      gpgme_data_t outdata,
                      engine_filter_t filter, void *hwnd,
                      const char *from_address);

int op_assuan_start_keymanager (void *hwnd);

int op_assuan_start_confdialog (void *hwnd);

int op_assuan_start_decrypt_files (void *hwnd, char **filenames);


#ifdef __cplusplus
}
#endif
#endif /*GPGOL_ENGINE_ASSUAN_H*/
