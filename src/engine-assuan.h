/* engine-assuan.h - Assuan server based crypto engine
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
 * You should have received a copy of the GNU Lesser General Public
 * License along with GpgOL; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
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



int  op_assuan_init (void);
void op_assuan_deinit (void);

int op_assuan_encrypt (protocol_t protocol, 
                       gpgme_data_t indata, gpgme_data_t outdata,
                       void *notify_data, /* FIXME: Add hwnd */
                       char **recipients);





#ifdef __cplusplus
}
#endif
#endif /*GPGOL_ENGINE_ASSUAN_H*/
