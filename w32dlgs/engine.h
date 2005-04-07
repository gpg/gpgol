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

#ifndef _ENGINE_H
#define _ENGINE_H 1

int op_init (void);
const char* op_strerror (int err);

#define op_debug_enable(file) op_set_debug_mode(5, (file))
#define op_debug_disable() op_set_debug_mode(0, NULL)
void op_set_debug_mode (int val, const char *file);

int op_encrypt_start (const char *inbuf, char **outbuf);
int op_encrypt (void *rset, const char *inbuf, char **outbuf);

int op_sign_encrypt_start (const char *inbuf, char **outbuf);
int op_sign_encrypt (void *rset, void *locusr, const char *inbuf, 
		     char **outbuf);

int op_verify (const char *inbuf, char **outbuf);

int op_sign_start (const char *inbuf, char **outbuf);
int op_sign (void *locusr, const char *inbuf, char **outbuf);

int op_decrypt (const char *inbuf, char **outbuf);

#endif /*_ENGINE_H*/
