#include <windows.h>
#include <commctrl.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"

struct _gpgme_engine_info  _gpgme_engine_ops_gpgsm;


HINSTANCE glob_hinst;

int main(int argc, char **argv)
{    
    gpgme_key_t *keys=NULL, signer = NULL;
    int opts = 0;
    int i;

    InitCommonControls();
    glob_hinst = GetModuleHandle(NULL);

    recipient_dialog_box(&keys, &opts);

    for (i=0; keys && keys[i] != NULL; i++)
	printf ("%s\n", keys[i]->uids->name);	
    
    free(keys);

    signer_dialog_box(&signer, NULL);
    if (signer)
	printf ("%s\n", signer->uids->name);

    cleanup_keycache_objects();
    return 0;
}