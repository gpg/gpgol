
#include <stdio.h>
#include <time.h>

#include "gpgme.h"
#include "keycache.h"

struct _gpgme_engine_info _gpgme_engine_ops_gpgsm;

int main(int argc, char **argv)
{
    keycache_t c = NULL;
    keycache_t n;

    keycache_init (NULL, 1, &c);
    for (n=c; n; n = n->next) 
	printf ("%s\n", n->key->uids->name);

    keycache_release(c);
    return 0;
}
