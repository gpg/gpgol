/* MapiGPGME.h - Mapi support with GPGME
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
#ifndef MAPI_GPGME_H
#define MAPI_GPGME_H

class MapiGPGME
{
private:
    HWND hwnd;
    LPMESSAGE msg;
    char* defaultKey;

public:
    MapiGPGME ();
    MapiGPGME (LPMESSAGE msg);
    ~MapiGPGME ();

private:
    int setBody (char *body);
    char *getBody ();

    int countRecipients (char **recipients);
    char **getRecipients (bool isRootMsg);
    void freeRecipients (char **recipients);

public:
    int encrypt ();
    int decrypt ();
    int sign ();
    int verify ();
    int signEncrypt ();

public:
    void setDefaultKey (const char *key);
    char* getDefaultKey ();

public:
    void setMessage(LPMESSAGE msg);
    void setWindow(HWND hwnd);
};


#endif /*MAPI_GPGME_H*/