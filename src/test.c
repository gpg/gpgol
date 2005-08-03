#include <windows.h>
#include <commctrl.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"
#include "engine.h"

int enum_gpg_seckeys (gpgme_key_t * ret_key, void **ctx);

void set_global_hinstance(HINSTANCE hinst);

HINSTANCE glob_hinst;

int main(int argc, char **argv)
{    
    gpgme_key_t *keys=NULL, *keys2=NULL;
    gpgme_key_t signer = NULL;
    gpgme_error_t err;
    int opts = 0;
    int i, n=0;
    void *ctx=NULL;
    const char *s;
    char *encmsg=NULL;
    char **id = NULL;
    char **un=NULL;

    InitCommonControls();
    //set_global_hinstance (GetModuleHandle(NULL));
    
    op_init();

#if 0
    

    recipient_dialog_box(&keys, &opts);

    for (i=0; keys && keys[i] != NULL; i++)
	printf ("%s\n", keys[i]->uids->name);

    err = op_encrypt_file (keys, "c:\\foo.txt", "c:\\foo.txt.asc");
    if (err)
	printf ("enc_file: %s\n", op_strerror (err)); 
    free(keys);
#endif
    
#if 0
    /*
    signer_dialog_box(&signer, NULL);
    if (signer)
	printf ("%s\n", signer->uids->name);
    */

    /*
    id = xcalloc (4, sizeof (char*));
    id[0] = xstrdup ("john.frakes");
    id[1] = xstrdup ("ts@g10code");
    id[2] = xstrdup ("foo@bar");
    id[3] = NULL;
    op_lookup_keys (id, (void **)&keys, &un, &n);
    for (i=0; i < 4; i++) {
	if (keys[i] != NULL)
	    printf ("%p %s\n", keys[i], keys[i]->uids->name);
    }
    */
    
    un = /*x*/calloc (2, sizeof (char*));
    un[0] = /*x*/strdup ("twoaday@freakmail.de");
    un[1] = NULL;
    keys = /*x*/calloc (2, sizeof (gpgme_key_t));
    keys[0] = find_gpg_key("9A1C182E", 0);
    keys[1] = NULL;
    n=1;
    printf ("---------------%d\n", n);
    recipient_dialog_box2 (keys, un, n, &keys2, &i);
    for (i =0; keys2 != NULL && keys2[i] != NULL; i++)
	printf ("%s\n", keys2[i]->uids->name);

    /*free (id[0]);
    free (id[1]);
    free (id[2]);
    free (id);*/
    free (keys);
    free (un[0]);
    free (un);
    free (keys2);

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
#endif

#if 0
    init_keycache_objects ();
    s =
    "-----BEGIN PGP SIGNED MESSAGE-----\r\n"
    "Hash: SHA1\r\n"
    "\r\n"
    "simply OpenPGP test file\r\n"
    "\r\n"
    "-----BEGIN PGP SIGNATURE-----\r\n"
    "Version: GnuPG v1.4.2-cvs (MingW32)\r\n"
    "\r\n"
    "iQEVAwUBQoDAPVUNGuWgkqh9AQIacgf9ENxW7zrFLQijI8ghrOf64TUUtIpxALtq\r\n"
    "AqslH9vaHtsub3ofOR9wtw9aGy0hn44TN374eOZ/c7QTnGqCmd/TGqeVbHZukYsc\r\n"
    "GVeyNJdp1zNA9oJSpj3jYwUHRcMzGAMTX2fZhkdiV38cmkfjBFx/z/lvvN2AA+La\r\n"
    "9ndfVvccmxW9RF+dcihopEpTCnbr7msfl5di9W9THKZUs8NvFRFyoXZRhifif/BG\r\n"
    "dvWvG4uGuAwXLUgSwyDnK2RlumlrHKVo29zWl/2VidzmET5UeBLzi/wh4se+whpx\r\n"
    "XKCXNBqqsUsYKThdiRxRnGAF3iTSHULSW96vloO2mUsYSJcjNB++wQ==\r\n"
    "=/+XI\r\n"
    "-----END PGP SIGNATURE-----\r\n";
    op_verify_start (s, &encmsg);
    printf ("%s\n", encmsg);
    free (encmsg);
#endif

#if 1   
    s =
    "-----BEGIN PGP MESSAGE-----\r\n"
    "Version: GnuPG v1.4.2-cvs (MingW32)\r\n"
    "\r\n"
    "hQEMA1UNGuWgkqh9AQf9EiH2b7aMsKPryhHqhn90s03BQw3Eh/KAgPe45mAYhTgL\r\n"
    "kjdqwyuLQzWLOLdt4ZpBRrW4/J61ElBACjKsWhyYtuJXvpAxtfbnfsVpDJ/6ljaT\r\n"
    "nMlHDodvd6zRpDRLnWA9EH7l+fnNLE+Ekln7xDJqSsJstSTfY+nv+FSy3lwGrSBA\r\n"
    "iuVVZheOMRLgl4ClrMzDENjWgVZiZiP/KXaea3q3PdLHvKCPXqd5ZaJ0d/CJOP3r\r\n"
    "m7E0L/HoPAIDaz4zFk/yiUwN/p+E/y+7tk+YvpHT9mNinE4vH4ychHt/SjIv/ocb\r\n"
    "qWQT9VDTR+80nxh4SVeT97ewI16igHmwm6490zWyR9JdAYY63Yh7QScDZxSLCYGD\r\n"
    "H8liUR3bw55uaMsLLwRJ/x9quhIS8ofvgxSOHBtZpEvUdWTAGM0WhkPloLp8WCZ6\r\n"
    "OnHM9sgg807GticXwhyx7fP1eNjdo+Nb+Aabgy2h\r\n"
    "=nUq/\r\n"
    "-----END PGP MESSAGE-----\r\n";
    op_decrypt_start (s, &encmsg);
    if (encmsg != NULL) {
	printf ("%s\n", encmsg);
	free (encmsg);
    }
#endif
    
    op_deinit ();
    return 0;
}
