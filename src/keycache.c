/* keycache.c
 *	Copyright (C) 2004 Timo Schulz
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
#include <windows.h>
#include <time.h>

#include "gpgme.h"
#include "keycache.h"
#include "intern.h"

static keycache_t pubring = NULL;
static keycache_t secring = NULL;
static time_t last_timest = 0;


/* Initialize global keycache objects */
void
init_keycache_objects (void)
{
    if (pubring || secring)
	cleanup_keycache_objects ();
    keycache_init (NULL, 0, &pubring);
    keycache_init (NULL, 1, &secring);
}


/* Cleanup global keycache objects */
void 
cleanup_keycache_objects (void)
{
    keycache_release (pubring);
    pubring = NULL;
    keycache_release (secring);
    secring = NULL;
}


/* Initialize global keycache from external objects */
void
load_keycache_objects (keycache_t ring[2])
{
    pubring = ring[0];
    secring = ring[1];
}

/* Create a new keycache object and return it in r_ctx. */
int 
keycache_new (keycache_t *r_ctx)
{
    keycache_t c;

    c = xcalloc (1, sizeof *c);
    *r_ctx = c;
    return 0;
}


/* Release keycache object */
void 
keycache_release (keycache_t ctx)
{
    keycache_t n;

    while (ctx) {
	n = ctx->next;
	gpgme_key_release (ctx->key);
	free (ctx);
	ctx = n;
    }
}

/* Free the keycache object but not the keys which are stored in it. */
void 
keycache_free(keycache_t ctx)
{
    keycache_t n;

    while (ctx) {
	n = ctx->next;
	free (ctx);
	ctx = n;
    }
}


/* Add a key to the keycache object. */
int 
keycache_add (keycache_t *ctx, gpgme_key_t key)
{
    keycache_t c;
    int rc;

    rc = keycache_new (&c);
    if (rc)
	return rc;
    c->key = key;
    c->next = *ctx;
    *ctx = c;

    return 0;
}


/* Initialize the keycache object with GPG keys which match the
   given pattern. */
int 
keycache_init (const char *pattern, int seconly, keycache_t *r_ctx)
{
    gpgme_ctx_t ctx;
    gpgme_key_t key;
    gpgme_error_t rc;

    rc = gpgme_new (&ctx);
    if (rc)
	return rc;

    rc = gpgme_op_keylist_start (ctx, pattern, seconly);
    if (rc)
	goto leave;
    while (!gpgme_op_keylist_next (ctx, &key))
	keycache_add (r_ctx, key);
    gpgme_op_keylist_end (ctx);

leave:
    gpgme_release (ctx);
    return rc;
}


/* Return the size of the keycache object */
int 
keycache_size (keycache_t ctx)
{
    int i;
    keycache_t n;
    for (i=0, n=ctx; n; n = n->next, i++)
	;
    return i;
}


/* Find a key specified by an email address */
gpgme_key_t
find_gpg_email (const char *str)
{
    keycache_t n;
    gpgme_user_id_t u;

    for (n=pubring; n; n = n->next) {
	for (u = n->key->uids; u; u = u->next) {
	    if (strstr (u->uid, str))
		return n->key;
	}
    }
    return NULL;
}

/* Find a key in the public keyring cache */
gpgme_key_t 
find_gpg_key (const char *str, int type)
{
    keycache_t n;
    gpgme_subkey_t s;

    if (type == 1)
	return find_gpg_email (str);

    for (n=pubring; n; n = n->next) {
	for (s=n->key->subkeys; s; s = s->next) {
	    if (!strncmp(str, s->keyid+8, 8))
		return n->key;
	}
    }
    return NULL;
}


/* this function works directly with GPG. caller has to free the key. */
gpgme_key_t
get_gpg_key (const char *str)
{
    gpgme_error_t rc;
    gpgme_ctx_t ctx;
    gpgme_key_t key;

    rc = gpgme_new (&ctx);
    if (rc)
	return NULL;
    rc = gpgme_op_keylist_start (ctx, str, 0);
    if (!rc)
	rc = gpgme_op_keylist_next (ctx, &key);
    gpgme_release (ctx);
    return key;
}


/* Enumerate all public keys. If the keycache is not initialized, load it. */
int 
enum_gpg_keys (gpgme_key_t * ret_key, void **ctx)
{
    if (!pubring) {
	keycache_release (pubring); pubring=NULL;
	keycache_init (NULL, 0, &pubring);
	*ctx = pubring;
    }
    if (!ret_key) {
	if (time (NULL) > last_timest+1750) { /* refresh after 30 minutes */
	    last_timest = time (NULL);
	    cleanup_keycache_objects ();
	    keycache_init(NULL, 0, &pubring);	    
	}
	*ctx = pubring;
	return 0;
    }
    *ret_key = ((keycache_t)(*ctx))->key;
    *ctx = ((keycache_t)(*ctx))->next;
    return *ctx != NULL? 0 : -1;
}


/* Enumerate all secret keys. */
int 
enum_gpg_seckeys (gpgme_key_t * ret_key, void **ctx)
{
    if (!secring ||*ctx == NULL) {
	keycache_release (secring); secring=NULL;
	keycache_init (NULL, 1, &secring);
	*ctx = secring;
    }
    if (!ret_key)
	return 0;
    *ret_key = ((keycache_t)(*ctx))->key;
    *ctx = ((keycache_t)(*ctx))->next;
    return *ctx != NULL? 0 : -1;
}


/* Reset the secret key enumeration. */
void 
reset_gpg_seckeys (void **ctx)
{
    *ctx = secring;
}
