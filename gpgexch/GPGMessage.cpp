/* GPGMessage.cpp - message class
 *	Copyright (C) 2004 g10 Code GmbH
 * 
 * This file is part of the G DATA Outlook Plugin for GnuPG.
 * 
 * This plugin is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * This plugin is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General
 * Public License along with this plugin; if not, write to the 
 * Free Software Foundation, Inc., 59 Temple Place - Suite 330, 
 * Boston, MA 02111-1307, USA.
 */

#include "stdafx.h"
#include "GPGMessage.h"

CGPGMessage::CGPGMessage (LPMESSAGE pMsg)
{
    this->pMessage = pMsg;
    this->pStreamBody = NULL;
    this->sPropVal = NULL;
}

CGPGMessage::~CGPGMessage (void)
{
}


int CGPGMessage::hasAttach (void)
{
    int nAttach = 0;

    hr = HrGetOneProp (pMessage, PR_HASATTACH, &sPropVal);
    if (SUCCEEDED (hr))
	nAttach = sPropVal->Value.b? 1 : 0;
    return nAttach;
}


int CGPGMessage::isUnset (void)
{
    HrGetOneProp (pMessage, PR_MESSAGE_FLAGS, &sPropVal);
    if (SUCCEEDED (hr) && sPropVal->Value.l & MSGFLAG_UNSENT)
	return -1;
    return 0;
}

int CGPGMessage::getBody (char ** pBody)
{
    return NULL;
}


int CGPGMessage::setBody (char * pBody)
{
    sProp.ulPropTag = PR_BODY;
    sProp.Value.lpszA = pBody;
    hr = HrSetOneProp (pMessage, &sProp);
    if (SUCCEEDED (hr))
	return 0;
    return -1;
}


int CGPGMessage::hasGPGData (void)
{
    return getGPGType () != PGP_NONE;
}

int CGPGMessage::getGPGType (void)
{
    char *pBody = NULL;
    int gpgType = PGP_NONE;

    getBody (&pBody);
    if (pBody)
    {
	if (strstr (pBody, "BEGIN PGP MESSAGE") &&
	    strstr (pBody, "END PGP MESSAGE"))
	    gpgType |= PGP_MSG;
	if (strstr (pBody, "BEGIN PGP SIGNATURE") &&
	    strstr (pBody, "END PGP SIGNATURE"))
	    gpgType |= PGP_SIG;
	delete []pBody;
    }
    return gpgType;
}
