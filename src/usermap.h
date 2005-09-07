/* usermap.h - Map keyid's to userid's
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 *
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGol; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */
#ifndef GPGOL_USERMAP_H
#define GPGOL_USERMAP_H

void* new_usermap (gpgme_recipient_t rset);
void  free_usermap (void *ctx);

void* HashTable_new (int n);
void  HashTable_free (void *ctx);
void  HashTable_put (void *ctx, const char *key, void *val);
void* HashTable_get (void *ctx, const char *key);
void* HashTable_get_i (void *ctx, int pos);
int   HashTable_size (void *ctx);

#endif /*GPGOL_USERMAP_H*/
