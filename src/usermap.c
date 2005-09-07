/* usermap.c - Map keyid's to userid's
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
#include <windows.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"
#include "usermap.h"

/* Creates a hash table with the keyid as the name and the user-id as 
   the value. So it is possible to map keyid->userid. */
void*
new_usermap (gpgme_recipient_t rset)
{
    gpgme_recipient_t r;
    gpgme_ctx_t ctx=NULL;
    gpgme_key_t key;
    gpgme_error_t err;
    char *p;
    int c = 0;
    void *tab = NULL;

    for (r = rset; r; r = r->next)
	c++;    
    p = xcalloc (1, c*(17+2));
    for (r = rset; r; r = r->next) {
	strcat (p, r->keyid);
	strcat (p, " ");
    }

    err = gpgme_new (&ctx);
    if (err)
	goto fail;
    err = gpgme_op_keylist_start (ctx, p, 0);
    if (err)
	goto fail;

    tab = HashTable_new (c);
    r = rset;
    while (1) {
	const char *uid;
	err = gpgme_op_keylist_next (ctx, &key);
	if (err)
	    break;
	uid = gpgme_key_get_string_attr (key, GPGME_ATTR_USERID, NULL, 0);
	if (uid != NULL) {
	    char *u = xstrdup (uid);
	    HashTable_put (tab, r->keyid, u);
	}
	gpgme_key_release (key);
	key = NULL;
	r = r->next;
    }

fail:
    if (ctx != NULL)
	gpgme_release (ctx);
    xfree (p);
    return tab;
}


/* Free all memory used by the hash table. Assume all entries are strings. */
void
free_usermap (void *ctx)
{    
    int i;

    for (i=0; i < HashTable_size (ctx); i++) {
	char *p = (char *)HashTable_get_i (ctx, i);
	xfree (p);
    }
    HashTable_free (ctx);
}
