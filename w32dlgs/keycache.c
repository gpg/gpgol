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
#include <malloc.h>

#include "gpgme.h"
#include "keycache.h"


//#pragma data_seg(".SHARDAT")
static keycache_t pubring = NULL;
static keycache_t secring = NULL;
static time_t last_timest = 0;
//#pragma data_seg()


void cleanup_keycache_objects(void)
{
    keycache_release(pubring);
    pubring = NULL;
    keycache_release(secring);
    secring = NULL;
}


int keycache_new (keycache_t *r_ctx)
{
    keycache_t c;

    c = calloc(1, sizeof *c);
    if (!c)
	return -1;
    *r_ctx = c;
    return 0;
}

void keycache_release (keycache_t ctx)
{
    keycache_t n;

    while (ctx) {
	n = ctx->next;
	gpgme_key_release (ctx->key);
	free (ctx);
	ctx = n;
    }
}

void keycache_free(keycache_t ctx)
{
    keycache_t n;

    while (ctx) {
	n = ctx->next;
	free (ctx);
	ctx = n;
    }
}


int keycache_add (keycache_t *ctx, gpgme_key_t key)
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


int keycache_init (const char *pattern, int seconly, keycache_t *r_ctx)
{
    gpgme_ctx_t ctx;
    gpgme_key_t key;
    gpgme_error_t rc;

    rc = gpgme_new (&ctx);
    if (rc)
	return rc;

    rc = gpgme_op_keylist_start(ctx, pattern, seconly);
    if (rc)
	goto leave;
    while (!gpgme_op_keylist_next(ctx, &key))
	keycache_add (r_ctx, key);
    gpgme_op_keylist_end(ctx);

leave:
    gpgme_release(ctx);
    return rc;
}

int keycache_size(keycache_t ctx)
{
    int i;
    keycache_t n;
    for (i=0, n=ctx; n; n = n->next, i++)
	;
    return i;
}


gpgme_key_t find_gpg_key (const char *keyid)
{
    keycache_t n;
    gpgme_subkey_t s;

    for (n=pubring; n; n = n->next) {
	for (s=n->key->subkeys; s; s = s->next) {
	    if (!strncmp(keyid, s->keyid+8, 8))
		return n->key;
	}
    }
    return NULL;
}


int enum_gpg_keys (gpgme_key_t * ret_key, void **ctx)
{
    if (!pubring) {
	keycache_init (NULL, 0, &pubring);
	*ctx = pubring;
    }
    if (!ret_key) {
	if (time(NULL) > last_timest+1750) { /* refresh after 30 minutes */
	    last_timest = time(NULL);
	    cleanup_keycache_objects();
	    keycache_init(NULL, 0, &pubring);
	    *ctx = pubring;
	}	
	return 0;
    }    
    *ret_key = ((keycache_t)(*ctx))->key;
    *ctx = ((keycache_t)(*ctx))->next;
    return *ctx != NULL? 0 : -1;
}


int enum_gpg_seckeys(gpgme_key_t * ret_key, void **ctx)
{
    if (!secring) {
	keycache_init(NULL, 1, &secring);
	*ctx = secring;
    }
    if (!ret_key)
	return 0;
    *ret_key = ((keycache_t)(*ctx))->key;
    *ctx = ((keycache_t)(*ctx))->next;
    return *ctx != NULL? 0 : -1;
}


void reset_gpg_seckeys(void **ctx)
{
    *ctx = secring;
}

