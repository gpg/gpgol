/* HashTable.cpp
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
#include <string.h>

#include "HashTable.h"

#ifdef __GNUC__
#define __INLINE__ __inline__
#else
#define __INLINE__ inline
#endif

#define CRC24_INIT 0xb704ceL
#define CRC24_POLY 0x1864cfbL
       
static long __INLINE__
CRC (unsigned char *octets, size_t len)
{ 
    long crc = CRC24_INIT;
    int i;
    while (len--) {
	crc ^= (*octets++) << 16;
	for (i = 0; i < 8; i++) {
	    crc <<= 1;
	    if (crc & 0x1000000)
		crc ^= CRC24_POLY;
	}
    }   
    return crc & 0xffffffL;    
}


HashTable::HashTable (void)
{
    pos = 0;
    n = 8;
    table = new void*[n];
    memset (table, 0, sizeof (void*)*n);
}

HashTable::HashTable (unsigned int n)
{
    pos = 0;
    this->n = n;
    table = new void*[n];
    memset (table, 0, sizeof (void*)*n);
}

HashTable::~HashTable (void)
{
    delete []table;
}


unsigned 
HashTable::size (void)
{
    return pos;
}

void 
HashTable::put (const char *key, void *val)
{
    unsigned i = CRC ((unsigned char*)key, strlen (key));
    if (table[i % n] == NULL)
	pos++;
    table[i % n] = val;
}

void*
HashTable::get (const char *key)
{
    unsigned i = CRC ((unsigned char*)key, strlen (key));
    return table[i % n];
}


void*
HashTable::get (int index)
{
    return table[index % n];
}


void
HashTable::clear (void)
{
    /* XXX: there is no way to free the object because it is not known what
	    kind of object it is. */
}



extern "C" { /* C-interface */

void*
HashTable_new (int n)
{
    return new HashTable (n);
}

void
HashTable_free (void *ctx)
{
    HashTable *c = (HashTable *)ctx;
    delete c;
}

void
HashTable_put (void *ctx, const char *key, void *val)
{
    HashTable *c = (HashTable *)ctx;
    c->put (key, val);
}

void*
HashTable_get (void *ctx, const char *key)
{
    HashTable *c = (HashTable *)ctx;
    return c->get (key);
}

void*
HashTable_get_i (void *ctx, int pos)
{
    HashTable *c = (HashTable *)ctx;
    return c->get (pos);
}

int
HashTable_size (void *ctx)
{
    HashTable *c = (HashTable *)ctx;
    return c->size ();
}

} /* end C-interface */
