#include <windows.h>
#include <commctrl.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"
#include "engine.h"

struct _gpgme_engine_info  _gpgme_engine_ops_gpgsm;

int enum_gpg_seckeys (gpgme_key_t * ret_key, void **ctx);

HINSTANCE glob_hinst;

int main(int argc, char **argv)
{    
    gpgme_key_t *keys=NULL, signer = NULL;
    int opts = 0;
    int i;
    void *ctx=NULL;
    const char *s;
    char *encmsg=NULL;
    char **id = NULL;

    InitCommonControls();
    glob_hinst = GetModuleHandle(NULL);

    /*
    recipient_dialog_box(&keys, &opts);

    for (i=0; keys && keys[i] != NULL; i++)
	printf ("%s\n", keys[i]->uids->name);	
    
    free(keys);
    */
    
    /*
    signer_dialog_box(&signer, NULL);
    if (signer)
	printf ("%s\n", signer->uids->name);
    */

    op_init();

    id = calloc (3, sizeof (char*));
    id[0] = strdup ("johnfrakes");
    id[1] = strdup ("ts@g10code");
    id[2] = NULL;
    op_lookup_keys (id, (void **)&keys);
    for (i=0; i < 2; i++)
	if (keys[i] != NULL)
	    printf ("%s\n", keys[i]->uids->name);

    free (id[0]);
    free (id[1]);
    free (id);
    free (keys);
    
    /*
    enum_gpg_seckeys (NULL, &ctx);
    s = 
    "-----BEGIN PGP MESSAGE-----\r\n"
    "Version: GnuPG v1.4.2-cvs (MingW32)\r\n"
    "\r\n"
    "hQHOA2fdhuWifE69EAb/YZp0m8zkKix+R+6K9T3hql0/DfXgOTxDIJxxuPU0XxP2\r\n"
    "+sso8EA3A7QNP3kDeEK3FKJpfyT5F/7KeKWmIaaw0mpF3oRwEIXl0rxJqhgt+ipX\r\n"
    "aAXXrBWD9s8lnMdqACh7B8pA+OQXHDDhm0PveBbZcYm5MHtn/mSJY6zDsI97EMMV\r\n"
    "t9yurBc6HsxbzhjzBA9lhJl0ttpDGplLf6LSq3tlgPry1T4KHdZMh7z+NX7lcSOV\r\n"
    "x9qdWh4i7U69XLa9queGdm1+2cNvWqckGqCMr04s5NJQ0HiSqcAAqDV8TDTs3eYH\r\n"
    "AJzc52Rtl3G/5N+EAekuu4DUwDblGcMzHeiLbmDSEWfgu/bsKBjtxrXCPcl3Jwov\r\n"
    "FXKD2fiV80gCGo7cMa65RR8huzq5iri2uq8EB3X6C2AdiQhzLDH++Tdp6m0D6lDk\r\n"
    "leNlvaX5P24EoAujD54POtuGN8CaVTCvpGZcKylv8SVqOx38nPRYJyyqAqUKnxG0\r\n"
    "ZAwp0wPLrntSNgM8i3Zs0PmUY1rj4SVfcdrYd0AihXU0xOIzhDYedH+wEkyY2FPm\r\n"
    "q1ZKWoFMGlUzv8c6BK5G3Vt9BgloHpDRPT2KwgwxhpZG0kEBwo91zd/Pq5rJNTee\r\n"
    "GzbNUA+VK5sAufwiGbyUI6Az+5I+dgzFZw9CV5qgxKKxLVMYekV9Fp3xm3Jfa910\r\n"
    "Lg+hVQ==\r\n"
    "=O05/\r\n"
    "-----END PGP MESSAGE-----\r\n";
    op_decrypt_start (s, &encmsg);
    printf ("%s\n", encmsg);
    free (encmsg);
    */
    /*
    //op_set_debug_mode (5, "gpgme.dbg");
    op_sign_encrypt_start ("test", &encmsg);
    //op_sign_start("test", &encmsg);
    printf ("%s\n", encmsg);
    free (encmsg);
    */
    /*
    init_keycache_objects ();
    s =
    "-----BEGIN PGP SIGNED MESSAGE-----\r\n"
    "Hash: SHA1\r\n"
    "\r\n"
    "12345678901234567890\r\n"
    "\r\n"
    "-----BEGIN PGP SIGNATURE-----\r\n"
    "\r\n"
    "iEYEARECAAYFAkJVX78ACgkQ0pkwxcTOkYOd0gCaAr7vkyUXqbGhGcAiIDppcanM\r\n"
    "CywAnjDfzfZUoapLsXQIs0rkN9ahKU5I\r\n"
    "=MO2m\r\n"
    "-----END PGP SIGNATURE-----\r\n";
    op_verify_start (s, &encmsg);
    printf ("%s\n", encmsg);
    free (encmsg);
    */
    
    op_deinit ();
    return 0;
}
