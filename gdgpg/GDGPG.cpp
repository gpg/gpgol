/* GDGPG.cpp - implementation of the dll exports and the GDGPG class
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
#include "resource.h"
#include <initguid.h>
#include "GDGPG.h"

#include "GDGPG_i.c"
#include "GDGPGO.h"

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_GDGPG, CGDGPG)
END_OBJECT_MAP()

/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point
extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID /*lpReserved*/)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        _Module.Init(ObjectMap, hInstance, &LIBID_GDGPGLib);
        DisableThreadLibraryCalls(hInstance);
    }
    else if (dwReason == DLL_PROCESS_DETACH)
        _Module.Term();
    return TRUE;    // ok
}


/* Used to determine whether the DLL can be unloaded by OLE */
STDAPI DllCanUnloadNow (void)
{
    return (_Module.GetLockCount ()==0) ? S_OK : S_FALSE;
}


/* Returns a class factory to create an object of the requested type */
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _Module.GetClassObject (rclsid, riid, ppv);
}


/* DllRegisterServer - Adds entries to the system registry */
STDAPI DllRegisterServer (void)
{
    /* registers object, typelib and all interfaces in typelib */
    return _Module.RegisterServer (TRUE);
}


/* DllUnregisterServer - Removes entries from the system registry */
STDAPI DllUnregisterServer (void)
{
    return _Module.UnregisterServer (TRUE);
}


/* GetStringFromRegistry
 Gets a string value from the registry.
*/
string GetStringFromRegistry(HKEY hKey, LPCTSTR pPath, LPCTSTR pSubKey, LPCTSTR pDefault)
{
    string strResult(pDefault);
    HKEY h;

    if ((pPath == NULL) || (pSubKey == NULL))
	return strResult;


    if (RegOpenKeyEx(hKey, pPath, 0, KEY_READ, &h) == ERROR_SUCCESS)
    {
	TCHAR  buf[MAX_PATH];
	DWORD  size = sizeof(buf);
	if (RegQueryValueEx(h, pSubKey, NULL, NULL, (LPBYTE)buf, &size) == ERROR_SUCCESS)
	    strResult = buf;
	RegCloseKey(h);	
    }
    return strResult;
}


int CGDGPG::LogInfo (const char * fmt, ...)
{
    va_list arg;

    if (!UseLogging ())
	return 0;

    fprintf (m_LogFP, "gdgpg: ");
    va_start (arg, fmt);
    vfprintf (m_LogFP, fmt, arg);
    va_end (arg);
    fflush (m_LogFP);
    return 0;
}


CGDGPG::CGDGPG (void)
{
    string sGPGPath;
 
    /*
    sGPGPath = GetStringFromRegistry(HKEY_LOCAL_MACHINE, 
		"Software\\Microsoft\\Windows\\CurrentVersion", "ProgramFilesDir", ".\\");
    */
    sGPGPath = GetStringFromRegistry (HKEY_CURRENT_USER, "Software\\GNU\\GnuPG",
				      "HomeDir", "c:\\gnupg");
    if ((sGPGPath.size() > 0) && (sGPGPath[sGPGPath.size()-1] != '\\'))
	sGPGPath += '\\';
    m_sGPGExe = sGPGPath + "gpg.exe";
    m_sKeyManagerExe = sGPGPath + "winpt.exe --keymanager";
    ReadOptions();
    m_nStorePassphraseTime = 0;
    m_dwPasspharseTime = 0;
    m_sPassphraseInvalid = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
    m_sPassphraseStore = m_sPassphraseInvalid;
    m_LogFP = NULL;
    m_LogLevel = 0;
}


CGDGPG::~CGDGPG (void)
{
    if (m_LogFP != NULL)
    {
	fclose (m_LogFP);
	m_LogFP = NULL;
    }
}

/**
 * CallGPG: Call gpg.exe with specified commands.
 * Return value: TRUE if successfull.
 **/
BOOL 
CGDGPG::CallGPG(
	string sCommand,     
	string sPassphrase,  // The passpharse (if necessary).
	string &sResult,     // The result (StdOut of gpg).
	string &sError,      // The error (StdError of gpg).
	int nTimeoutSec,     // The maximum time waiting for the gpg process.
	BOOL bKillProcessOnTimeout)  // Indicates whether to kill the gpg process on timeout.
{
    SECURITY_ATTRIBUTES saAttr; 
    STARTUPINFO sInfo;
    PROCESS_INFORMATION pInfo;
    HANDLE hReadPipeIn, hWritePipeIn;
    HANDLE hReadPipeOut, hWritePipeOut;
    HANDLE hReadPipeError, hWritePipeError;
    BOOL bSuccess;
    BOOL nret;
    int nWaitSec = nTimeoutSec == 0? INFINITE : nTimeoutSec*1000;

    sResult = "";
    sError = "";
    if (sCommand.size() == 0)
    {
	LogInfo ("empty GPG command line, abort.");
	return FALSE;
    }
	
    sCommand = m_sGPGExe + " " + sCommand;

    memset (&saAttr, 0, sizeof saAttr);
    saAttr.nLength = sizeof(SECURITY_ATTRIBUTES); 
    saAttr.bInheritHandle = TRUE; 
    saAttr.lpSecurityDescriptor = NULL; 
    
    /* XXX: close pipes in case of an error */
    if (!CreatePipe(&hReadPipeIn, &hWritePipeIn, &saAttr, 1024))
    {
	LogInfo ("CreatePipe(in) failed ec=%d\n", (int)GetLastError ());
	return FALSE;
    }
    if (!CreatePipe(&hReadPipeOut, &hWritePipeOut, &saAttr, 4096))
    {
	LogInfo ("CreatePipe(out) failed ec=%d\n", (int)GetLastError ());
	return FALSE;
    }
    if (!CreatePipe(&hReadPipeError, &hWritePipeError, &saAttr, 4096))
    {
	LogInfo ("CreatePipe(err) failed ec=%d\n", (int)GetLastError ());
	return FALSE;
    }

    // create startup info for the gpg process
    memset (&sInfo, 0, sizeof(sInfo));
    sInfo.cb = sizeof(STARTUPINFO); 
    sInfo.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES; 
    sInfo.wShowWindow = SW_HIDE; 

    // connect to the pipes
    sInfo.hStdInput = hReadPipeIn; 
    sInfo.hStdError = hWritePipeError;
    sInfo.hStdOutput = hWritePipeOut;
    
    ::SetCursor(::LoadCursor(NULL, IDC_WAIT));

    LogInfo ("commandline:\n%s", sCommand.c_str ());

    bSuccess = CreateProcess( NULL, (char*) sCommand.c_str(),
			      NULL, NULL, TRUE, CREATE_DEFAULT_ERROR_MODE,
			      NULL, NULL, &sInfo, &pInfo );
    
    if (!bSuccess)
	LogInfo ("gpg procession created failed ec=%d\n", (int)GetLastError ());

    /* send passpharse (if spezified) */
    if (bSuccess && (sPassphrase.size() > 0))
    {
	string s = sPassphrase + "\r\n";
	DWORD dwWritten = 0;
	WriteFile(hWritePipeIn, s.c_str(), s.size(), &dwWritten, NULL);	
	LogInfo ("send passphrase to gpg\n");
    }
    CloseHandle(hWritePipeIn);

    // close the write side of the StdOut and StdError pipe
    CloseHandle(hWritePipeOut);
    CloseHandle(hWritePipeError);
    	
    // read result and error string from pipe
    if (bSuccess)
    {
	TCHAR cBuffer[1025];
	DWORD dwRead;

	do 
	{
	    memset(&cBuffer, 0, sizeof(cBuffer));
	    nret = ReadFile(hReadPipeOut, cBuffer, 1024, &dwRead, NULL);
	    sResult += cBuffer;
	    if (m_LogLevel > 1)
		LogInfo ("out: read %d bytes from fd=%d\n", dwRead, (int)hReadPipeOut);
	} while (dwRead > 0);	

	do 
	{
	    memset(&cBuffer, 0, sizeof(cBuffer));
	    nret = ReadFile(hReadPipeError, cBuffer, 1024, &dwRead, NULL);
	    sError += cBuffer;
	    if (m_LogLevel > 1)
		LogInfo ("err: read %d bytes from fd=%d\n", dwRead, (int)hReadPipeError);
	} while (dwRead > 0);

	// oem to ansi
	if (sResult != "")
	{
	    char* sResultAnsi = new char[sResult.size()+1];
	    strcpy(sResultAnsi, sResult.c_str());
	    OemToCharBuff(sResultAnsi, sResultAnsi, strlen(sResultAnsi));
	    sResult = sResultAnsi;
	    delete sResultAnsi;
	}
	if (sError != "")
	{
	    char* sErrorAnsi = new char[sError.size()+1];
	    strcpy(sErrorAnsi, sError.c_str());
	    OemToCharBuff(sErrorAnsi, sErrorAnsi, strlen(sErrorAnsi));
	    sError = sErrorAnsi;
	    delete sErrorAnsi;
	}	
    }

    CloseHandle(hReadPipeIn);
    CloseHandle(hReadPipeOut);
    CloseHandle(hReadPipeError);

    if (bSuccess) 
    {
	if (::WaitForSingleObject(pInfo.hProcess, nWaitSec) == WAIT_TIMEOUT) 
	{
	    bSuccess = FALSE;
	    LogInfo ("failed to wait for process %08lX (%d secs)\n", pInfo.dwProcessId, nWaitSec);
	    if (bKillProcessOnTimeout) // the hard way
	    {
		/* maybe gpg halts on an error or is waiting for a user input */
		if (TerminateProcess(pInfo.hProcess, 0) != 0)
		{
		    LogInfo ("successfully terminated the gpg process\n");
		    bSuccess = TRUE;
		}
	    }
	}
	CloseHandle(pInfo.hProcess);
	CloseHandle(pInfo.hThread);
    }

    if (sError != "")
	LogInfo ("gpg error:\n%s", sError.c_str ());
    if (m_LogLevel > 1 && sResult != "")
	LogInfo ("gpg result:\n%s", sResult.c_str ());

    ::SetCursor(::LoadCursor(NULL, IDC_ARROW));
    return bSuccess;
}


/* GenerateKeyList - Generates the key list (m_keyList).

  Return value: TRUE if successfull.
*/
BOOL 
CGDGPG::GenerateKeyList (void)
{	
    CKey * pKey = NULL;
    string sResult, sError;

    if (m_keyList.size() > 0) 
    {
	for (vector<CKey>::iterator i = m_keyList.begin();
	     i != m_keyList.end(); i++)
	    i->m_bSelected = FALSE;
	return TRUE;	
    }

    string sCommand = "--list-keys --with-colons";
    if (!CallGPG (sCommand, "", sResult, sError, 0, FALSE))
    {
	LogInfo ("GeneratedKeyList failed (gpg error).\n");
	return FALSE;
    }

    sResult += '\n';

    while (sResult.find ('\n') != -1)
    {
	string sLine = sResult.substr (0, sResult.find ('\n'));
	sResult = sResult.substr (sResult.find ('\n')+1);

	if (sLine.substr (0,4) == "pub:")
	{
	    if (pKey != NULL)
	    {
		if ((pKey->m_sUser.size () > 0) &&
		    (pKey->m_sAddress.size () > 0) &&
		    (pKey->m_sIDPub.size () > 0))
		{
		    m_keyList.push_back (*pKey);
		}	
		delete pKey;
		pKey = NULL;	
	    }
	    pKey = new CKey;
	    pKey->SetParameter (sLine);
	}
	if ((sLine.substr(0,4) == "sub:") && (pKey != NULL))
	{
	    pKey->SetSubParameter (sLine);
	}
	if ((sLine.substr (0,4) == "uid:") && (pKey != NULL))
	{
	    pKey->SetUserIDParameter (sLine);
	}
    }
    if (pKey != NULL)
    {
	if ((pKey->m_sUser.size() > 0) &&
	    (pKey->m_sAddress.size() > 0) &&
	    (pKey->m_sIDPub.size() > 0))
	{
	    m_keyList.push_back(*pKey);
	}
	delete pKey;
	pKey = NULL;
    }
    if (pKey != NULL)
	delete pKey;	
    return TRUE;
}


/* GetDefaultKey

 Sets the default key. Uses the first valid key from the
 secret keyring.

 Return value: TRUE if successfull.
*/
BOOL 
CGDGPG::GetDefaultKey (void)
{
    string sResult, sError;
	
    if (m_keyDefault.m_sUser != "")
	return TRUE;

    m_keyDefault.Clear();
    string sCommand = "--list-secret-keys --with-colons";

    if (!CallGPG (sCommand, "", sResult, sError, 0, FALSE))
    {
	LogInfo ("GetDefaultKey failed (gpg error)\n");
	return FALSE;
    }

    /* fixme: do not use the first key, but the first which is valid! */
    sResult += '\n';
    while (sResult.find('\n') != -1)
    {
	string sLine = sResult.substr(0, sResult.find('\n'));
	sResult = sResult.substr(sResult.find('\n')+1);
	if (sLine.substr(0,4) == "sec:")
	{
	    if (m_keyDefault.m_sIDPub != "")
		break;
	    m_keyDefault.SetParameter(sLine);
	}
	if (sLine.substr(0,4) == "ssb:")
	{
	    m_keyDefault.SetSubParameter(sLine);
	}	
    }
    return (m_keyDefault.m_sUser.size() > 0);
}

/////////////////////////////////////////////////////////////////////////////
// SetDefaultRecipients
//
// Uses the recipients string to set the m_bSelected flag of the recipient 
// keys.
// 
void CGDGPG::SetDefaultRecipients(
	vector<CKey> &keyList, // The key list.
	string sRecipients)    // The recipient string. The recipients are separated by a "|".
{

    sRecipients += "|";
    
    while (sRecipients.find('|') != -1)
    {
	string sRecipient = sRecipients.substr(0, sRecipients.find('|'));
	sRecipients = sRecipients.substr(sRecipients.find('|')+1);
	if (sRecipient.size() > 0)
	{
	    for (vector<CKey>::iterator i = keyList.begin(); i != keyList.end(); i++)
	    {
		if (!i->IsValidKey ())
		    continue;
		if (!i->m_bSelected)
		{		
		    if (sRecipient.find('@') != -1)
		    {
			if (stricmp(sRecipient.c_str(), i->m_sAddress.c_str()) == 0)
			    i->m_bSelected = TRUE;
			if (strstr (i->m_sUserID.c_str (), sRecipient.c_str ()))
			    i->m_bSelected = TRUE;
		    }
		    else
		    {
			if (sRecipient.find('/') != 0)
			    sRecipient = sRecipient.substr(sRecipient.rfind('/')+1);
			if (sRecipient.find('=') != 0)
			    sRecipient = sRecipient.substr(sRecipient.rfind('=')+1);
			string sAddress = i->m_sAddress;
			if (sAddress.find('@') != -1)
			    sAddress = sAddress.substr(0, sAddress.find('@'));
			if ((stricmp(sRecipient.c_str(), sAddress.c_str()) == 0) ||
			    (stricmp(sRecipient.c_str(), i->m_sUser.substr(0, sRecipient.size()).c_str()) == 0))
			    i->m_bSelected = TRUE;
			if (strstr (i->m_sUserID.c_str (), sRecipient.c_str ()))
			    i->m_bSelected = TRUE;
		    }
		}	
	    }
	}	
    }
}

/////////////////////////////////////////////////////////////////////////////
// CheckPassphrase
//
// Checks whether the passpharse is valid.
//
// Return value: TRUE if the passpharse is valid.
// 
BOOL CGDGPG::CheckPassphrase(string sPassphrase)
{
    string sCommand = "--sign --batch --yes --status-fd 1 --passphrase-fd 0 check.tmp";
    string sResult, sError;
    int bad_pass = 0;

    CallGPG(sCommand, sPassphrase, sResult, sError, 0, FALSE);
    if (sResult.find ("[GNUPG:] BAD_PASSPHRASE") == -1)
    {
	bad_pass = 1;
	LogInfo ("gpg returned 'bad passphrase'.\n");
    }
    return bad_pass;
}

/////////////////////////////////////////////////////////////////////////////
// GetDecryptInfo
//
// Gets decrypt informations of the specified file.
//
// Return value: TRUE if successful.
// 
BOOL CGDGPG::GetDecryptInfo (
	string sFilename,       // The file.
	BOOL &bIsEncrypted,     // Indicates whethet the file is encrypted (return value).
	BOOL &bIsSigned,        // Indicates whether the file is signed (return value).
	string &sDecryptKeys,   // A list with all known decrypt keys (return value).
	int &nUnknownKeys)      // The number of unknown decrpyt keys (return value).
{
    bIsEncrypted = FALSE;
    bIsSigned = FALSE;
    sDecryptKeys = "";
    nUnknownKeys = 0;
    string sCommand = "--batch --status-fd 1 \"" + sFilename + "\"";
    string sResult, sError;

    if (!CallGPG (sCommand, "", sResult, sError, 10, TRUE))
    {
	LogInfo ("GetDecryptInfo failed (gpg error).\n");
	return FALSE;
    }

    m_gpgStdErr = sError;

    if (sResult.find ("[GNUPG:] ENC_TO") != -1)
	bIsEncrypted = TRUE;

    if ((sResult.find ("[GNUPG:] GOODSIG") != -1) ||
	(sResult.find ("[GNUPG:] BADSIG")  != -1) ||
	(sResult.find ("[GNUPG:] EXPKEYSIG")  != -1) ||
	(sResult.find ("[GNUPG:] ERRSIG")  != -1))
	bIsSigned = TRUE;
    if (!bIsEncrypted)
    {
	LogInfo ("%s: file was not encrypted by gpg\n", sFilename.c_str ());
	return FALSE;
    }

    sResult += '\n';
    while (sResult.find ('\n') != -1)	
    {
	string sLine = sResult.substr (0, sResult.find('\n'));	
	sResult = sResult.substr (sResult.find('\n')+1);

	if (sLine.find ("[GNUPG:] ENC_TO") != -1)
	{
	    nUnknownKeys++;
	    sLine = sLine.substr(16);
	    if (sLine.find (' ') != -1)
	    {
		sLine = sLine.substr (0, sLine.find (' '));
		for (vector<CKey>::iterator i = m_keyList.begin(); 
		     i != m_keyList.end(); i++)
		{
		    if ((i->GetKeyID (-1) == sLine) || 
			(i->GetKeyID (-1) == sLine) ||
			(i->m_sIDSubOther.find (sLine + "|") != -1))
		    {
			if (sDecryptKeys != "")
			    sDecryptKeys += '\n';
			sDecryptKeys += i->m_sUser;
			nUnknownKeys--;
			break;	
		    }
		}	
	    }
	}	
    }
    return TRUE;
}


/* StorePassphrase - Stores the passphrase (if the store passpharse time is not zero). */
void 
CGDGPG::StorePassphrase(string sPassphrase)
{
    if (m_nStorePassphraseTime > 0)
    {
	m_sPassphraseStore = sPassphrase;
	m_dwPasspharseTime = GetTickCount();	
    }
}


/* GetStoredPassphrase - Returns the stored passphrase, if the store passpharse time not exceeds. */
string 
CGDGPG::GetStoredPassphrase (void)
{
    DWORD dwTime = GetTickCount ();
    if ((m_nStorePassphraseTime <= 0) || 
	(m_dwPasspharseTime == 0) ||
	(m_sPassphraseStore == "") || 
	(m_sPassphraseStore == m_sPassphraseInvalid))
	return "";
    if ((dwTime > m_dwPasspharseTime) && 
	(dwTime < (m_dwPasspharseTime+m_nStorePassphraseTime*1000)))
	return m_sPassphraseStore;
    return "";
}


/* ReadOptions - Reeds the options from the registry. */
void
CGDGPG::ReadOptions()
{
	
    TCHAR szKeyName[] = _T("Software\\G DATA\\GPG");
    HKEY hKey;

    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, szKeyName, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
	DWORD dwSize;
	TCHAR s[MAX_PATH+1] = "";

	dwSize	= sizeof(s)-1;
	if (RegQueryValueEx(hKey, _T("GPGExe"), NULL, NULL, (LPBYTE) &s, &dwSize) == ERROR_SUCCESS)
	    m_sGPGExe = s;
	strcpy(s, "");
	dwSize	= sizeof(s)-1;
	if (RegQueryValueEx(hKey, _T("KeyManagerExe"), NULL, NULL, (LPBYTE) &s, &dwSize) == ERROR_SUCCESS)
	    m_sKeyManagerExe = s;
	RegCloseKey(hKey);	
    }
}


/* WriteOptions - Writes the options to the registry. */
void 
CGDGPG::WriteOptions()
{
    TCHAR szKeyName[] = _T("Software\\G DATA\\GPG");

    HKEY hKey;
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, szKeyName, 0, NULL, REG_OPTION_NON_VOLATILE,
			KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)	
    {
	DWORD dwSize = m_sGPGExe.size()+1;
	RegSetValueEx(hKey, _T("GPGExe"), 0, REG_SZ, (CONST BYTE *) m_sGPGExe.c_str(), dwSize);
	dwSize = m_sKeyManagerExe.size()+1;
	RegSetValueEx(hKey, _T("KeyManagerExe"), 0, REG_SZ, (CONST BYTE *) m_sKeyManagerExe.c_str(), dwSize);
	RegCloseKey(hKey);
    }
}
