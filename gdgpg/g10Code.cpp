/* g10Code.cpp - implementation of the g10Code Interface
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
#include "GDGPG.h"
#include "g10Code.h"
#include <sys/stat.h>


Cg10Code::Cg10Code (void)
{
    this->m_alwaysTrust = 0;
    this->m_armorMode = false;
    this->m_cipherText = NULL;
    this->m_compressLevel = 0;
    this->m_expertMode = false;
    this->m_localUser = NULL;
    this->m_pgpMode = 0;
    this->m_textMode = 0;
    this->m_gpgExe = NULL;
    this->m_outPut = NULL;
    this->m_commentStr = NULL;
    this->m_passPhrase = NULL;
    this->m_recipientSet = "";
    this->m_useNoPassphrase = 0;
    this->m_encryptTo = NULL;
    this->m_keyServer = NULL;
    this->m_homeDir = NULL;
}

Cg10Code::~Cg10Code (void)
{
    if (this->m_passPhrase)
    {
	memset (this->m_passPhrase, 0, strlen (this->m_passPhrase));
	delete [] this->m_passPhrase;
	this->m_passPhrase = NULL;
    }
    if (this->m_homeDir)
    {
	delete [] this->m_homeDir;
	this->m_homeDir = NULL;
    }
    if (this->m_keyServer)
    {
	delete [] this->m_keyServer;
	this->m_keyServer = NULL;
    }
    if (this->m_encryptTo)
    {
	delete [] this->m_encryptTo;
	this->m_encryptTo = NULL;
    }
    if (this->m_outPut)
    {
	delete [] this->m_outPut;
	this->m_outPut = NULL;
    }
    if (this->m_localUser)
    {
	delete [] this->m_localUser;
	this->m_localUser = NULL;
    }
    if (this->m_plainText)
    {
	delete [] this->m_plainText;
	this->m_plainText = NULL;
    }
    if (this->m_cipherText)
    {
	delete [] this->m_cipherText;
	this->m_cipherText=NULL;
    }
    if (this->m_gpgExe)
    {
	delete [] this->m_gpgExe;
	this->m_gpgExe = NULL;
    }
    if (this->m_commentStr)
    {
	delete [] this->m_commentStr;
	this->m_commentStr = NULL;
    }
}


void Cg10Code::logInfo (const char * fmt, ...)
{
    va_list arg;

    if (!useLogging ())
	return;

    fprintf (m_logFP, "gdgpg: ");
    va_start (arg, fmt);
    vfprintf (m_logFP, fmt, arg);
    va_end (arg);
    fflush (m_logFP);
}

Cg10Code::strFromBstr (char **strVal, BSTR newVal)
{
    USES_CONVERSION;
    char *pExe = OLE2A (newVal);

    if (*strVal)
	delete [] *strVal;
    *strVal = new char [strlen (pExe)+1];
    if (!*strVal)
	return S_FALSE;
    strcpy (*strVal, pExe);
    return S_OK;
}


char* Cg10Code::data2File (const char * dataBuf, const char * fileName)
{
    char * tmpFile;
    FILE * fp;

    tmpFile = new char[200 + strlen (fileName)+1];
    if (!tmpFile)
	return NULL;
    memset (tmpFile, 0, 200+strlen (fileName)+1);
    GetTempPath (200, tmpFile);
    strcat (tmpFile, fileName);
    if (dataBuf)
    {
	fp = fopen (tmpFile, "wb");
	if (!fp)
	{
	    delete []tmpFile;
	    return NULL;
	}
	fwrite (dataBuf, 1, strlen (dataBuf), fp);
	fclose (fp);
    }
    return tmpFile;
}

int 
Cg10Code::file2Data (char ** dataBuf, const char * fileName, int delFile)
{
    struct stat st;
    FILE * fp;

    fp = fopen (fileName, "rb");
    if (!fp)
	return -1;
    fstat (fileno (fp), &st);
    if (st.st_size == 0)
    {
	fclose (fp);
	return -1;
    }
    if (*dataBuf)
	delete [] *dataBuf;
    *dataBuf = new char [st.st_size+1];
    if (!*dataBuf)
    {
	fclose (fp);
	return -1;
    }
    fread (*dataBuf, 1, st.st_size, fp);
    (*dataBuf)[st.st_size] = '\0';
    fclose (fp);
    if (delFile)
	unlink (fileName);
    return 0;
}


int
Cg10Code::parseOptions (string &sCommand, const char * outName)
{
    if (!this->m_gpgExe)
	return -1;
    sCommand = this->m_gpgExe;
    if (this->m_outPut || outName)
    {
	sCommand += " --output \"";
	sCommand += outName? outName : this->m_outPut;
	sCommand += "\" ";
    }
    else
	sCommand += " ";
    sCommand += "--status-fd=2 ";
    if (!this->m_useNoPassphrase && this->m_passPhrase)
	sCommand += "--passphrase-fd=0 ";
    if (this->m_alwaysTrust)
	sCommand += "--always-trust ";
    if (this->m_armorMode)
	sCommand += "--armor ";
    if (this->m_noVersion)
	sCommand += "--no-version ";
    if (this->m_commentStr)
    {
	sCommand += "--comment \"";
	sCommand += this->m_commentStr;
	sCommand += "\" ";
    }
    if (this->m_expertMode)
	sCommand += "--expert ";
    if (this->m_localUser)
    {
	sCommand += "--local-user ";
	sCommand += this->m_localUser;
	sCommand += " ";
    }
    if (this->m_encryptTo)
    {
	sCommand += "--encrypt-to ";
	sCommand += this->m_encryptTo;
	sCommand += " ";
    }
    if (this->m_keyServer)
    {
	sCommand += "--keyserver ";
	sCommand += this->m_keyServer;
	sCommand += " ";
    }
    if (this->m_homeDir)
    {
	sCommand += "--homedir ";
	sCommand += this->m_homeDir;
	sCommand += " ";
    }
    if (this->m_textMode)
	sCommand += "--text-mode ";
    if (this->m_forceMDC)
	sCommand += "--force-mdc ";
    if (this->m_forceV3SIG)
	sCommand += "--force-v3-sigs ";
    return 0; /* XXX */

    if (this->m_pgpMode)
    {
	sCommand += "--pgp-mode" + this->m_pgpMode;
	sCommand += " ";
    }    
    if (this->m_compressLevel)
    {
	sCommand += "--compress-level " + this->m_compressLevel;
	sCommand += " ";
    }
    return 0;
}


BOOL
Cg10Code::spawnGPG (string sCommand, string sPassphrase, 
		    string &sStdout, string &sStderr,
		    int nTimeoutSec, BOOL bKillProcessOnTimeout)
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

    sStdout = "";
    sStderr = "";
    if (sCommand.size() == 0)
    {
	logInfo ("empty GPG command line, abort.");
	return FALSE;
    }

    memset (&saAttr, 0, sizeof saAttr);
    saAttr.nLength = sizeof(SECURITY_ATTRIBUTES); 
    saAttr.bInheritHandle = TRUE; 
    saAttr.lpSecurityDescriptor = NULL; 
    
    /* XXX: close pipes in case of an error */
    if (!CreatePipe(&hReadPipeIn, &hWritePipeIn, &saAttr, 1024))
    {
	logInfo ("CreatePipe(in) failed ec=%d\n", (int)GetLastError ());
	return FALSE;
    }
    if (!CreatePipe(&hReadPipeOut, &hWritePipeOut, &saAttr, 4096))
    {
	logInfo ("CreatePipe(out) failed ec=%d\n", (int)GetLastError ());
	return FALSE;
    }
    if (!CreatePipe(&hReadPipeError, &hWritePipeError, &saAttr, 4096))
    {
	logInfo ("CreatePipe(err) failed ec=%d\n", (int)GetLastError ());
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

    logInfo ("commandline:\n%s\n", sCommand.c_str ());

    bSuccess = CreateProcess( NULL, (char*) sCommand.c_str(),
			      NULL, NULL, TRUE, CREATE_DEFAULT_ERROR_MODE,
			      NULL, NULL, &sInfo, &pInfo );
    
    if (!bSuccess)
	logInfo ("gpg procession created failed ec=%d\n", (int)GetLastError ());

    /* send passpharse (if spezified) */
    if (bSuccess && (sPassphrase.size() > 0))
    {
	string s = sPassphrase + "\r\n";
	DWORD dwWritten = 0;
	WriteFile(hWritePipeIn, s.c_str(), s.size(), &dwWritten, NULL);	
	logInfo ("send passphrase to gpg\n");
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
	    sStdout += cBuffer;
	    if (m_logLevel > 1)
		logInfo ("out: read %d bytes from fd=%d\n", dwRead, (int)hReadPipeOut);
	} while (dwRead > 0);	

	do 
	{
	    memset(&cBuffer, 0, sizeof(cBuffer));
	    nret = ReadFile(hReadPipeError, cBuffer, 1024, &dwRead, NULL);
	    sStderr += cBuffer;
	    if (m_logLevel > 1)
		logInfo ("err: read %d bytes from fd=%d\n", dwRead, (int)hReadPipeError);
	} while (dwRead > 0);

	// oem to ansi
	if (sStdout != "")
	{
	    char* sResultAnsi = new char[sStdout.size()+1];
	    strcpy(sResultAnsi, sStdout.c_str());
	    OemToCharBuff(sResultAnsi, sResultAnsi, strlen(sResultAnsi));
	    sStdout = sResultAnsi;
	    delete sResultAnsi;
	}
	if (sStderr != "")
	{
	    char* sErrorAnsi = new char[sStderr.size()+1];
	    strcpy(sErrorAnsi, sStderr.c_str());
	    OemToCharBuff(sErrorAnsi, sErrorAnsi, strlen(sErrorAnsi));
	    sStderr = sErrorAnsi;
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
	    logInfo ("failed to wait for process %08lX (%d secs)\n", pInfo.dwProcessId, nWaitSec);
	    if (bKillProcessOnTimeout) // the hard way
	    {
		/* maybe gpg halts on an error or is waiting for a user input */
		if (TerminateProcess(pInfo.hProcess, 0) != 0)
		{
		    logInfo ("successfully terminated the gpg process\n");
		    bSuccess = TRUE;
		}
	    }
	}
	CloseHandle(pInfo.hProcess);
	CloseHandle(pInfo.hThread);
    }

    if (sStderr != "")
	logInfo ("gpg stderr:\n%s", sStderr.c_str ());
    if (m_logLevel > 1 && sStdout != "")
	logInfo ("gpg stdout:\n%s", sStdout.c_str ());

    ::SetCursor(::LoadCursor(NULL, IDC_ARROW));
    return bSuccess;
}


STDMETHODIMP Cg10Code::get_Plaintext(BSTR *pVal)
{
    CComBSTR b;

    b.Empty ();
    b.Append (this->m_plainText? this->m_plainText : "");
    *pVal = b.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Plaintext(BSTR newVal)
{
    return strFromBstr (&this->m_plainText, newVal);
}

STDMETHODIMP Cg10Code::get_Ciphertext(BSTR *pVal)
{
    CComBSTR b;

    b.Empty ();
    b.Append (this->m_cipherText? this->m_cipherText : "");
    *pVal = b.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Ciphertext(BSTR newVal)
{
    return strFromBstr (&this->m_cipherText, newVal);
}

STDMETHODIMP Cg10Code::get_Armor(BOOL *pVal)
{
    *pVal = this->m_armorMode;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Armor(BOOL newVal)
{
    this->m_armorMode = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_LocalUser(BSTR *pVal)
{
    CComBSTR b;

    b.Empty ();
    b.Append (this->m_localUser? this->m_localUser : "");
    *pVal = b.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_LocalUser(BSTR newVal)
{
    return strFromBstr (&this->m_localUser, newVal);
}

STDMETHODIMP Cg10Code::get_CompressLevel(long *pVal)
{
    *pVal = this->m_compressLevel;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_CompressLevel(long newVal)
{
    if (newVal >= 0 && newVal <= 9)
	this->m_compressLevel = newVal;	
    return S_OK;
}

STDMETHODIMP Cg10Code::get_TextMode(BOOL *pVal)
{
    *pVal = this->m_textMode;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_TextMode(BOOL newVal)
{
    this->m_textMode = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_Expert(BOOL *pVal)
{
    *pVal = this->m_expertMode;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Expert(BOOL newVal)
{
    this->m_expertMode = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_PGPMode(long *pVal)
{
    *pVal = this->m_pgpMode;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_PGPMode(long newVal)
{
    if (newVal == 2 || newVal == 6 || newVal == 7
	|| newVal == 8)
	this->m_pgpMode = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_AlwaysTrust(BOOL *pVal)
{
    *pVal = this->m_alwaysTrust;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_AlwaysTrust(BOOL newVal)
{
    this->m_alwaysTrust = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::SetLogLevel(long logLevel)
{
    this->m_logLevel = logLevel;
    return S_OK;
}

STDMETHODIMP Cg10Code::SetLogFile(BSTR logFile, long *pvReturn)
{
    USES_CONVERSION;

    *pvReturn = 0;
    if (m_logFP)
    {
	fclose (m_logFP);
	m_logFP = NULL;
    }

    m_logFP = fopen (OLE2A (logFile), "a+b");
    if (m_logFP != NULL && m_logLevel < 1)
	m_logLevel = 1;
    return S_OK;
}


STDMETHODIMP Cg10Code::get_Binary(BSTR *pVal)
{
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_gpgExe? this->m_gpgExe : "");
    *pVal = a.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Binary(BSTR newVal)
{
    return strFromBstr (&this->m_gpgExe, newVal);
}

STDMETHODIMP Cg10Code::AddRecipient(BSTR name, long *pvReturn)
{
    USES_CONVERSION;
    char * rcptName = OLE2A (name);
    if (this->m_recipientSet == "")
	this->m_recipientSet = "-r ";
    else
	this->m_recipientSet += "-r ";
    this->m_recipientSet += rcptName;
    this->m_recipientSet += " ";
    *pvReturn = 0;
    return S_OK;
}


STDMETHODIMP Cg10Code::Decrypt(long *pvReturn)
{
    string sCommand = "";
    string sOut, sErr;
    char *inName, *outName;

    if (!this->m_cipherText)
    {
	*pvReturn = g10code_err_NoCiphertext;
	return S_OK;
    }
    if (!this->m_passPhrase)
    {
	*pvReturn = g10code_err_NoPassphrase;
	return S_OK;
    }
    inName = data2File (this->m_cipherText, "gpg_in_data.gpg");
    outName = data2File (NULL, "gpg_out_data.dat");
    this->m_useNoPassphrase = 0;
    if (parseOptions (sCommand, outName))
    {
	*pvReturn = g10code_err_NoBinary;
	return S_OK;
    }
    sCommand = sCommand + "--decrypt \"" + inName + "\"";
    if (spawnGPG (sCommand, this->m_passPhrase, sOut, sErr, 120, TRUE))
	file2Data (&this->m_plainText, outName, 1);
    if (sErr.find ("[GNUPG:] BAD_PASSPHRASE") != -1)
	*pvReturn = g10code_err_BadPassphrase;
    else
	*pvReturn = g10code_err_Success;
    delete []inName;
    delete []outName;
    return S_OK;
}


STDMETHODIMP Cg10Code::Export(BSTR keyNames, long *pvReturn)
{
    USES_CONVERSION;
    string sCommand = "";
    string sOut, sErr;
    char *outName;
    char *keyPatt = OLE2A (keyNames);

    this->m_useNoPassphrase = 1;
    outName = data2File (NULL, "gpg_out_key.gpg");
    if (parseOptions (sCommand, outName))
    {
	*pvReturn = g10code_err_NoBinary;
	return S_OK;
    }
    sCommand += "--no-sk-comments ";
    *pvReturn = g10code_err_Success;    
    sCommand = sCommand + "--export " + keyPatt;
    if (spawnGPG (sCommand, "", sOut, sErr, 120, TRUE))
    {
	if (file2Data (&this->m_cipherText, outName, 1))
	    *pvReturn = g10code_err_KeyNotFound;
    }
    delete []outName;
    return S_OK;
}


STDMETHODIMP Cg10Code::DecryptFile(BSTR inFile, long *pvReturn)
{
    USES_CONVERSION;
    string sCommand = "";
    string sOut, sErr;
    char * inName = OLE2A (inFile);

    if (this->m_passPhrase == NULL)
    {
	*pvReturn = g10code_err_NoPassphrase;
	return S_OK;
    }
    this->m_useNoPassphrase = 0;
    if (parseOptions (sCommand, NULL))
    {
	*pvReturn = g10code_err_NoBinary;
	return S_OK;
    }
    sCommand += "--yes ";
    sCommand = sCommand + "--decrypt \"" + inName + "\"";
    spawnGPG (sCommand, this->m_passPhrase, sOut, sErr, 120, TRUE);    
    if (sErr.find ("[GNUPG:] BAD_PASSPHRASE") != -1)
	*pvReturn = g10code_err_BadPassphrase;
    else
	*pvReturn = g10code_err_Success;
    return S_OK;
}

STDMETHODIMP Cg10Code::EncryptFile(BSTR inFile, long *pvReturn)
{
    USES_CONVERSION;
    string sCommand = "";
    string sOut, sErr;
    char * inName = OLE2A (inFile);

    if (this->m_recipientSet == "")
    {
	*pvReturn = g10code_err_NoRecipients;
	return S_OK;
    }
    this->m_useNoPassphrase = 1;
    if (parseOptions (sCommand, NULL))
    {
	*pvReturn = g10code_err_NoBinary;
	return S_OK;
    }
    if (!this->m_outPut)
	sCommand += "--no-mangle-dos-filenames ";    
    sCommand += "--yes ";
    sCommand += this->m_recipientSet.c_str ();
    sCommand = sCommand + "--encrypt \"" + inName + "\"";
    /* XXX: check return code and factor out some common code */
    spawnGPG(sCommand, "", sOut, sErr, 120, TRUE);
    if (sErr.find ("[GNUPG:] INV_RECP") != -1)
	*pvReturn = g10code_err_InvRecipients;
    else if (sErr.find ("[GNUPG:] END_ENCRYPTION") != -1)
	*pvReturn = g10code_err_Success;
    return S_OK;
}


STDMETHODIMP Cg10Code::Encrypt(long *pvReturn)
{
    string sCommand = "";
    string sOut, sErr;
    char * inName, *outName;

    if (this->m_recipientSet == "")
    {
	*pvReturn = g10code_err_NoRecipients;
	return S_OK;
    }
    if (this->m_plainText == NULL)
    {
	*pvReturn = g10code_err_NoPlaintext;
	return S_OK;
    }
    inName = data2File (this->m_plainText, "gpg_in_data.dat");
    outName = data2File (NULL, "gpg_out_data.gpg");
    this->m_useNoPassphrase = 1;
    if (parseOptions (sCommand, outName))
    {
	*pvReturn = g10code_err_NoBinary;
	return S_OK;
    }
    sCommand += this->m_recipientSet.c_str ();
    sCommand = sCommand + "--encrypt \"" + inName + "\"";

    if (spawnGPG (sCommand, "", sOut, sErr, 120, TRUE))
	file2Data (&this->m_cipherText, outName, 1);
    delete [] inName;
    delete [] outName;
    if (sErr.find ("[GNUPG:] INV_RECP") != -1)
	*pvReturn = g10code_err_InvRecipients;
    else if (sErr.find ("[GNUPG:] END_ENCRYPTION") != -1)
	*pvReturn = g10code_err_Success;
    return S_OK;
}

STDMETHODIMP Cg10Code::ClearRecipient(void)
{
    this->m_recipientSet = "";
    return S_OK;
}

STDMETHODIMP Cg10Code::get_Output(BSTR *pVal)
{
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_outPut ? this->m_outPut : "");
    *pVal = a.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Output(BSTR newVal)
{
    return strFromBstr (&this->m_outPut, newVal);
}

STDMETHODIMP Cg10Code::get_NoVersion(BOOL *pVal)
{
    *pVal = this->m_noVersion;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_NoVersion(BOOL newVal)
{
    this->m_noVersion = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_Comment(BSTR *pVal)
{
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_commentStr ? this->m_commentStr : "");
    *pVal = a.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Comment(BSTR newVal)
{
    return strFromBstr (&this->m_commentStr, newVal);
}

STDMETHODIMP Cg10Code::put_Passphrase(BSTR newVal)
{
    return strFromBstr (&this->m_passPhrase, newVal);
}


STDMETHODIMP Cg10Code::get_EncryptoTo(BSTR *pVal)
{
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_encryptTo? this->m_encryptTo : "");
    *pVal = a.m_str;
    return S_OK;
}


STDMETHODIMP Cg10Code::put_EncryptoTo(BSTR newVal)
{
    return strFromBstr (&this->m_encryptTo, newVal);
}

STDMETHODIMP Cg10Code::get_ForceMDC(BOOL *pVal)
{
    *pVal = this->m_forceMDC;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_ForceMDC(BOOL newVal)
{
    this->m_forceMDC = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_ForceV3Sig(BOOL *pVal)
{
    *pVal = this->m_forceV3SIG;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_ForceV3Sig(BOOL newVal)
{
    this->m_forceV3SIG = newVal;
    return S_OK;
}

STDMETHODIMP Cg10Code::get_Keyserver(BSTR *pVal)
{	
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_keyServer ? this->m_keyServer : "");
    *pVal = a.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_Keyserver(BSTR newVal)
{
    return strFromBstr (&this->m_keyServer, newVal);
}

STDMETHODIMP Cg10Code::get_HomeDir(BSTR *pVal)
{
    CComBSTR a;

    a.Empty ();
    a.Append (this->m_homeDir? this->m_homeDir : "");
    *pVal = a.m_str;
    return S_OK;
}

STDMETHODIMP Cg10Code::put_HomeDir(BSTR newVal)
{
    return strFromBstr (&this->m_homeDir, newVal);
}

