/* Key.cpp - implementation of the key class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2003 Timo Schulz
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
#include "Key.h"
#include "resource.h"

CKey::CKey ()
{
    m_bSelected = FALSE;
    m_trust = TRUST_UNKNOWN;

}

/* SetParameter

  Sets the key parameter from a given line of the gpg key list output 
  ("pub:" or "sec:").
 
  Return value: TRUE if successful.
*/
BOOL 
CKey::SetParameter (string sParam)
{	
    string s = sParam.substr(0, 4);

    if ((s != "pub:") && (s != "sec:"))
	return FALSE;
	
    for (int i=0; i < 15; i++) 
    {
	int nPos = sParam.find (':');
	if (nPos == -1)
	    break;
	s = sParam.substr (0, nPos);
	sParam = sParam.substr (nPos+1);
	
	switch (i) 
	{
	case 1:  /* trust */
	    m_trust = TRUST_UNKNOWN;
	    if (s == "n")
		m_trust = TRUST_NONE;
	    else if (s == "m")
		m_trust = TRUST_MEDIUM;
	    else if( s == "e" )
		m_trust = TRUST_EXPIRED;
	    else if( s == "r" )
		m_trust = TRUST_REVOKED;
	    else if (s == "d")
		m_trust = TRUST_DISABLED;
	    else if ((s == "f") || (s == "u"))
		m_trust = TRUST_FULL;
	    break;

	case 4:
	    m_sIDPub = s;
	    break;

	case 6:
	    m_sValid = s;
	    break;

	case 9:
	    m_sUser = s;
	    if (s.find ('<') != -1)
		m_sAddress = s.substr (s.find('<')+1);
	    if (m_sAddress.find ('>') != -1)
		m_sAddress = m_sAddress.substr (0, m_sAddress.find('>'));
	    break;

	case 11:
	    m_caps = 0;
	    if (s.find ('E'))
		m_caps |= CAN_ENCRYPT;
	    if (s.find ('S'))
		m_caps |= CAN_SIGN;
	    if (s.find ('C'))
		m_caps |= CAN_CERTIFY;
	    if (s.find ('A'))
		m_caps |= CAN_AUTH;
	    break;

	case 14:
	    m_bSmartCard = FALSE;
	    if (s != "")
		m_bSmartCard = TRUE;
	    break;
	}
    }
    return TRUE;
}


BOOL CKey::SetUserIDParameter (string sParam)
{
    string s = sParam.substr (0, 4);

    if (s != "uid:")
	return FALSE;

    for (int i=0; i<10; i++)
    {
	int nPos = sParam.find (':');
	if (nPos == -1)
	    break;
	s = sParam.substr (0, nPos);
	sParam = sParam.substr (nPos+1);
	switch (i)
	{
	case 9:
	    m_sUserID += s + "|";
	    break;
	}
    }
    return TRUE;
}

/* SetSubParameter

  Sets the sub key parameter from a given line of the gpg key list output
  ("sub:" or "ssb:").
  
  Return value: TRUE if successful.
*/
BOOL CKey::SetSubParameter (string sParam)
{
    string s = sParam.substr(0, 4);

    if ((s != "sub:") && (s != "ssb:"))
	return FALSE;

    for (int i=0; i<12; i++) 
    {
	int nPos = sParam.find (':');
	if (nPos == -1)
	    break;
	s = sParam.substr (0, nPos);
	sParam = sParam.substr(nPos+1);
	switch (i) 
	{
	case 4:
	    if (m_sIDSub == "")
		m_sIDSub = s;
	    else
		m_sIDSubOther += s + "|";
	    break;

	case 11:
	    m_caps = 0;
	    if (s.find ('e'))
		m_caps |= CAN_ENCRYPT;
	    if (s.find ('s'))
		m_caps |= CAN_SIGN;
	    break;
	}
    }
    return TRUE;
}

/* GetTrustString

 Returns a readable decription string of the trust value.
*/
string 
CKey::GetTrustString (void)
{
    TCHAR s[64];
    UINT nId;

    switch (m_trust) {
    case TRUST_NONE:    nId = IDS_TRUST_NONE;    break;
    case TRUST_MEDIUM:  nId = IDS_TRUST_MEDIUM;  break;
    case TRUST_FULL:    nId = IDS_TRUST_FULL;    break;
    case TRUST_DISABLED:nId = IDS_TRUST_DISABLED;break;
    case TRUST_EXPIRED: nId = IDS_TRUST_EXPIRED; break;
    case TRUST_REVOKED: nId = IDS_TRUST_REVOKED; break;
    default:           nId = IDS_TRUST_UNKNOWN; break;
    }
    LoadString (_Module.m_hInstResource, nId, s, sizeof (s));
    return (string) s;
}


BOOL
CKey::IsValidKey(void)
{
    return m_trust != TRUST_EXPIRED && 
	   m_trust != TRUST_INVALID && 
	   m_trust != TRUST_REVOKED &&
	   m_trust != TRUST_DISABLED;
}


BOOL
CKey::CanDo (int flags)
{
    if (m_caps & flags)
	return TRUE;
    return FALSE;
}


string 
CKey::GetKeyID (int idx)
{
    if (idx == 0)
	return m_sIDPub;
    /* XXX: add all secondary keys and return the key-id of it */
    if (m_sIDSub != "")
	return m_sIDSub;
    if (m_sIDPub != "")
	return m_sIDPub;
    return "";
}


void CKey::Clear (void)
{
    m_sUser = "";
    m_sAddress = "";
    m_sIDPub = "";
    m_sIDSub = "";
    m_sIDSubOther = "";
    m_sUserID = "";
    m_sValid = "";
    m_bSelected = FALSE;
    m_trust = TRUST_UNKNOWN;
}
