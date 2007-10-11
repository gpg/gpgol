/* serpent.h - Definitions of the Serpent encryption algorithm.
 *	Copyright (C) 2007 g10 Code GmbH
 *
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifndef SERPENT_H
#define SERPENT_H
#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif


/* Note that there is no special rason that we use Serpent and not AES
   or even CAST5, blowfish or whatever.  Any decent block cipher will
   do; even the blocksize does not matter for our purpose.  The
   Serpent implementation was just the most convenient one top put
   into this project. */

struct symenc_context_s;
typedef struct symenc_context_s *symenc_t;


symenc_t symenc_open (const void *key, size_t keylen, 
                      const void *iv, size_t ivlen);
void symenc_close (symenc_t ctx);
void symenc_cfb_encrypt (symenc_t ctx, void *buffer_out, 
                         const void *buffer_in, size_t nbytes);
void symenc_cfb_decrypt (symenc_t ctx, void *buffer_out, 
                         const void *buffer_in, size_t nbytes);



#ifdef __cplusplus
}
#endif
#endif /*SERPENT_H*/
