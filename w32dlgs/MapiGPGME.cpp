/* MapiGPGME.cpp - Mapi support with GPGME
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
#include <mapidefs.h>
#include <mapiutil.h>
#include <time.h>

#include "MapiGPGME.h"
#include "gpgme.h"
#include "engine.h"
#include "keycache.h"


MapiGPGME::MapiGPGME(LPMESSAGE mpsg)
{
     this->msg = msg;
     op_init ();
}


MapiGPGME::MapiGPGME()
{
    op_init ();
}


MapiGPGME::~MapiGPGME ()
{
    op_deinit ();
}


int 
MapiGPGME::setBody (char *body)
{
    SPropValue sProp; 
    HRESULT hr;
    
    sProp.ulPropTag = PR_BODY;
    sProp.Value.lpszA = body;
    hr = HrSetOneProp (msg, &sProp);
    if (FAILED (hr))
	return FALSE;
    return TRUE;
}


char* 
MapiGPGME::getBody (void)
{
    HRESULT hr;
    LPSPropValue lpspvFEID = NULL;
    char *body;

    hr = HrGetOneProp((LPMAPIPROP) msg, PR_BODY, &lpspvFEID);
    if (FAILED(hr))
	return NULL;

    
    body = new char[strlen (lpspvFEID->Value.lpszA)+1];
    if (!body)
	abort ();
    strcpy (body, lpspvFEID->Value.lpszA);

    MAPIFreeBuffer (lpspvFEID);
    lpspvFEID = NULL;

    return body;
}


int 
MapiGPGME::countRecipients (char **recipients)
{
    for (int i=0; recipients[i] != NULL; i++)
	;
    return i;
}


char** 
MapiGPGME::getRecipients (bool isRootMsg)
{
    HRESULT hr;
    LPMAPITABLE lpRecipientTable = NULL;
    LPSRowSet lpRecipientRows = NULL;
    char **rset = NULL;

    if (!isRootMsg)
	return NULL;
        
    static SizedSPropTagArray (1L, PropRecipientNum) = {1L, {PR_EMAIL_ADDRESS}};
    

    hr = msg->GetRecipientTable(0, &lpRecipientTable);
    if (SUCCEEDED(hr)) {
	size_t j = 0;
        hr = HrQueryAllRows (lpRecipientTable, 
		         (LPSPropTagArray) &PropRecipientNum,
			 NULL, NULL, 0L, &lpRecipientRows);
	rset = new char*[lpRecipientRows->cRows+1];
        for (j = 0L; j < lpRecipientRows->cRows; j++) {
	    const char *s = lpRecipientRows->aRow[j].lpProps[0].Value.lpszA;
	    rset[j] = new char[strlen (s)+1];
	    strcpy (rset[j], s);
	}
	rset[j] = NULL;
	if (NULL != lpRecipientTable)
	    lpRecipientTable->Release();
	if (NULL != lpRecipientRows)
	    FreeProws(lpRecipientRows);	
    }

    return rset;
}


void 
MapiGPGME::freeRecipients(char **recipients)
{
    for (int i=0; recipients[i] != NULL; i++)
	delete []recipients[i];
    delete recipients;
}


int 
MapiGPGME::encrypt (void)
{
    char *body = getBody();
    char *newBody = NULL;
    char **recipients = getRecipients (true);
    gpgme_key_t *keys=NULL;
    
    int n = op_lookup_keys (recipients, (void **)&keys);
    if (n != countRecipients (recipients))
	0; // XXX: start recipient dialog because some keys were not found

    int err = op_encrypt((void*)keys, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Encryption", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete [] body;
    free (keys);
    free (newBody);
    freeRecipients (recipients);
    return 0;
}


int 
MapiGPGME::decrypt (void)
{
    char *body = getBody();
    char *newBody = NULL;

    int err = op_decrypt_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Decryption", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete []body;
    free (newBody);
    return 0;
}


int
MapiGPGME::sign (void)
{
    char *body = getBody();
    char *newBody = NULL;

    int err = op_sign_start (body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete []body;
    free (newBody);
    return 0;
}


int
MapiGPGME::signEncrypt ()
{
    char *body = getBody();
    char *newBody = NULL;
    char **recipients = getRecipients (TRUE);
    gpgme_key_t locusr, *keys = NULL;
    
    locusr = find_gpg_key (defaultKey, 0);
    if (!locusr)
	locusr = find_gpg_key (defaultKey, 1);

    int n = op_lookup_keys (recipients, (void **)&keys);
    if (n != countRecipients (recipients))
	0; // XXX: start recipient dialog because some keys were not found

    int err = op_sign_encrypt ((void *)keys, (void*)locusr, body, &newBody);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Sign Encrypt", MB_ICONERROR|MB_OK);
    else
	setBody (newBody);

    delete []body;
    free (newBody);
    return 0;
}


int 
MapiGPGME::verify ()
{
    char *body = getBody();
    
    int err = op_verify_start (body, NULL);
    if (err)
	MessageBox (NULL, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);

    delete []body;
    return 0;
}


void MapiGPGME::setDefaultKey(const char *key)
{
    if (defaultKey) {
	delete []defaultKey;
	defaultKey = NULL;
    }
    defaultKey = new char[strlen (key)+1];
    if (!defaultKey)
	abort ();
    strcpy (defaultKey, key);
}


char* MapiGPGME::getDefaultKey (void)
{
    if (defaultKey == NULL) {
	void *ctx=NULL;
	gpgme_key_t sk=NULL;

	enum_gpg_seckeys (NULL, &ctx);
	enum_gpg_seckeys (&sk, &ctx);

	defaultKey = new char[16+1];
	if (!defaultKey)
	    abort ();
	const char *s = gpgme_key_get_string_attr (sk, GPGME_ATTR_KEYID, NULL, 0);
	strcpy (defaultKey, s);
    }

    return defaultKey;
}


void 
MapiGPGME::setMessage (LPMESSAGE msg)
{
    this->msg = msg;
}


void
MapiGPGME::setWindow(HWND hwnd)
{
    this->hwnd = hwnd;
}
