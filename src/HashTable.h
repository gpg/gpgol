/* HashTable.h
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
#ifndef MAPI_HASHTABLE_H
#define MAPI_HASHTABLE_H

#define DLL_EXPORT __declspec(dllexport)

class HashTable
{
private:
    void **table;
    unsigned n;
    unsigned pos;

public:
    DLL_EXPORT HashTable ();
    DLL_EXPORT HashTable (unsigned n);
    DLL_EXPORT ~HashTable ();

public:
    DLL_EXPORT unsigned size ();
    DLL_EXPORT void put(const char *key, void* val);
    DLL_EXPORT void* get(const char *key);
    DLL_EXPORT void* get(int index);
    DLL_EXPORT void clear ();
};
#endif
