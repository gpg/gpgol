/* Key.h - declaration of the key class
 *	Copyright (C) 2001 G Data Software AG, http://www.gdata.de
 *	Copyright (C) 2003, 2004 Timo Schulz
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

#ifndef __KEY_H_
#define __KEY_H_

typedef enum {
    CAN_NONE = 0,
    CAN_CERTIFY,
    CAN_SIGN,
    CAN_ENCRYPT,
    CAN_AUTH,
} CAPS;

typedef enum {
    TRUST_UNKNOWN = 0,
    TRUST_NONE,
    TRUST_MEDIUM,
    TRUST_FULL,
    TRUST_INVALID,
    TRUST_DISABLED,
    TRUST_EXPIRED,
    TRUST_REVOKED,
} TRUST;

/* CKey - The CKey class represents one gpg key. */
class CKey  
{
public:	
    CKey();

// attributes
public:
    string m_sUser;
    string m_sAddress;
    string m_sIDPub;
    string m_sIDSub;    
    string m_sIDSubOther;
    string m_sValid;
    string m_sUserID;
    TRUST  m_trust;
    BOOL   m_bSelected;  // used to mark keys in the key list
    BOOL   m_bSmartCard;
    int    m_caps;

public:
    string GetKeyID (int idx);
    BOOL SetParameter (string sParam);
    BOOL SetSubParameter (string sParam);
    BOOL SetUserIDParameter (string sParam);
    string GetTrustString (void);
    BOOL IsValidKey (void);
    BOOL CanDo (int flags);
    void Clear (void);
};
#endif // __KEY_H_
