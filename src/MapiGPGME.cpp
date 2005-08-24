/* MapiGPGME.cpp - Mapi support with GPGME
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <time.h>
#include <assert.h>
#include <string.h>

#ifdef __MINGW32__
# include "mymapi.h"
# include "mymapitags.h"
#else /* !__MINGW32__ */
# include <atlbase.h>
# include <mapidefs.h>
# include <mapiutil.h>
# include <initguid.h>
# include <mapiguid.h>
#endif /* !__MINGW32__ */

#include "gpgme.h"
#include "keycache.h"
#include "intern.h"
#include "HashTable.h"
#include "MapiGPGME.h"
#include "engine.h"


/* Attachment information. */
#define ATT_SIGN(action) ((action) & GPG_ATTACH_SIGN)
#define ATT_ENCR(action) ((action) & GPG_ATTACH_ENCRYPT)
#define ATT_PREFIX ".pgpenc"

#define DEFAULT_ATTACHMENT_FORMAT GPG_FMT_CLASSIC


/* default extension for attachments */
#define EXT_MSG "pgp"
#define EXT_SIG "sig"



/* Given that we are only able to configure one log file, it does not
   make much sense t bind a log file to a specific instance of this
   class. We use a global variable instead which will be set by any
   instance.  The same goes for the respecive file pointer and the
   enable_logging state. */
static char *logfile;
static FILE *logfp;
static bool enable_logging;

/* For certain operations we need to acquire a log on the logging
   functions.  This lock is controlled by this Mutex. */
static HANDLE log_mutex;


static int lock_log (void);
static void unlock_log (void);


static HWND find_message_window (HWND parent, GpgMsg *msg);
static void log_key_info (MapiGPGME *g, const char *prefix,
                          gpgme_key_t *keys, gpgme_key_t locusr);





/*
   The implementation class of MapiGPGME.  
 */
class MapiGPGMEImpl : public MapiGPGME
{
public:    
  MapiGPGMEImpl () 
  {
    clearConfig ();
    clearObject ();
    this->passCache = new HashTable ();
    op_init ();
    prepareLogging ();
    log_debug ("MapiGPGME.constructor: ready\n");
  }

  ~MapiGPGMEImpl ()
  {
    unsigned int i=0;

    log_debug ("MapiGPGME.destructor: called\n");
    op_deinit ();
    xfree (defaultKey);
    log_debug ("hash entries %d\n", passCache->size ());
    for (i = 0; i < passCache->size (); i++) 
      {
	cache_item_t t = (cache_item_t)passCache->get (i);
	if (t != NULL)
          cache_item_free (t);
      }
    delete passCache; 
    passCache = NULL;
    freeAttachments ();
    cleanupTempFiles ();
  }

  void __stdcall destroy ()
  {
    delete this;
  }

  void operator delete (void *p) 
  {
    ::operator delete (p);
  }  

  
public:
  const char * __stdcall versionString (void)
  {
    return PACKAGE_VERSION;
  }
    
  int __stdcall encrypt (HWND hwnd, GpgMsg *msg);
  int __stdcall decrypt (HWND hwnd, GpgMsg *msg);
  int __stdcall sign (HWND hwnd, GpgMsg *msg);
  int __stdcall signEncrypt (HWND hwnd, GpgMsg *msg);
  int __stdcall verify (HWND hwnd, GpgMsg *msg);
  int __stdcall attachPublicKey (const char *keyid);

  int __stdcall doCmdAttach(int action);
  int __stdcall doCmdFile(int action, const char *in, const char *out);

  const char* __stdcall getLogFile (void) { return logfile; }
  void __stdcall setLogFile (const char *name)
  { 
    if (!lock_log ())
      {
        if (logfp)
          {
            fclose (logfp);
            logfp = NULL;
          }
    
        xfree (logfile);
        logfile = name? xstrdup (name) : NULL;
        unlock_log ();
      }
  }

  int __stdcall getStorePasswdTime (void)
  {
    return nstorePasswd;
  }

  void __stdcall setStorePasswdTime (int nCacheTime)
  {
    this->nstorePasswd = nCacheTime; 
  }

  bool __stdcall getEncryptDefault (void)
  {
    return doEncrypt;
  }

  void __stdcall setEncryptDefault (bool doEncrypt)
  {
    this->doEncrypt = doEncrypt; 
  }

  bool __stdcall getSignDefault (void)
  { 
    return doSign; 
  }

  void __stdcall setSignDefault (bool doSign)
  {
    this->doSign = doSign;
  }

  bool __stdcall getEncryptWithDefaultKey (void)
  {
    return encryptDefault;
  }
  
  void __stdcall setEncryptWithDefaultKey (bool encryptDefault)
  {
    this->encryptDefault = encryptDefault;
  }

  bool __stdcall getSaveDecryptedAttachments (void) 
  { 
    return saveDecryptedAtt;
  }

  void __stdcall setSaveDecryptedAttachments (bool saveDecrAtt)
  {
    this->saveDecryptedAtt = saveDecrAtt;
  }

  void __stdcall setEncodingFormat (int fmt)
  {
    encFormat = fmt; 
  }

  int __stdcall getEncodingFormat (void) 
  {
    return encFormat;
  }

  void __stdcall setSignAttachments (bool signAtt)
  {
    this->autoSignAtt = signAtt; 
  }

  bool __stdcall getSignAttachments (void)
  {
    return autoSignAtt;
  }

  void __stdcall setEnableLogging (bool val)
  {
    enable_logging = val;
  }

  bool __stdcall getEnableLogging (void)
  {
    return enable_logging;
  }

  int __stdcall readOptions (void);
  int __stdcall writeOptions (void);

  const char* __stdcall getAttachmentExtension (const char *fname);
  void __stdcall freeAttachments (void);
  

  bool __stdcall deleteAttachment (GpgMsg * msg, int pos)
  {
//     if (msg->DeleteAttach (pos, 0, NULL, 0) == S_OK)
//       return true;
//     return false;
  }

  LPATTACH __stdcall createAttachment (GpgMsg * msg, int &pos)
  {
    ULONG attnum;	
    LPATTACH newatt = NULL;
    
//     if (msg->CreateAttach (NULL, 0, &attnum, &newatt) == S_OK)
//       {
//         pos = attnum;
//         return newatt;
//       }
    return NULL;
  }

  void  __stdcall showVersion (void);

  int __stdcall startKeyManager ();
  void __stdcall startConfigDialog (HWND parent);

  void __stdcall setDefaultKey (const char *key);
  const char * __stdcall getDefaultKey (void);

  void __stdcall clearPassphrase (void) 
  {
    if (passCache != NULL)
      passCache->clear ();
  }


private:
  char	      *defaultKey; /* Malloced default key or NULL. */
  HashTable   *passCache;
  LPMAPITABLE attachTable;
  LPSRowSet   attachRows;
  void	      *recipSet;

  /* Options */
  int	  nstorePasswd;  /* Time in seconds the passphrase is stored. */
  bool    encryptDefault;
  bool    doEncrypt;
  bool    doSign;
  bool    saveDecryptedAtt; /* Save decrypted attachments. */
  bool    autoSignAtt;	    /* Sign all outgoing attachments. */
  int	  encFormat;        /* Encryption format for attachments. */

  void displayError (HWND root, const char *title);
  void prepareLogging (void)
  {
    char *val = NULL;
    
    load_extension_value ("logFile", &val);
    if (val && *val != '\"' && *val)
      {
        setLogFile (val);
        setEnableLogging (true);
      }
    xfree (val);	
  }

  void clearObject (void)
  {
    this->attachRows = NULL;
    this->attachTable = NULL;
    this->defaultKey = NULL;
    this->recipSet = NULL;
  }

  void clearConfig (void)
  {
    nstorePasswd = 0;
    doEncrypt = false;
    doSign = false;
    encryptDefault = false;
    saveDecryptedAtt = false;
    autoSignAtt = false;
    encFormat = DEFAULT_ATTACHMENT_FORMAT;
  }

  void cleanupTempFiles ();
  void freeUnknownKeys (char **unknown, int n);
  void freeKeyArray (void **key);

  bool  setAttachMethod (LPATTACH obj, int mode);
  int   getAttachMethod (LPATTACH obj);
  char* getAttachFilename (LPATTACH obj);
  char* getAttachPathname (LPATTACH obj);
  bool  setAttachFilename (LPATTACH obj, const char *name, bool islong);
  int   getMessageFlags (GpgMsg *msg);
  int   getMessageHasAttachments (GpgMsg *msg);
  bool  setXHeader (GpgMsg *msg, const char *name, const char *val);
  char* getXHeader (GpgMsg *msg, const char *name);
  bool  checkAttachmentExtension (const char *ext);
  const char* getPGPExtension (int action);
  char* generateTempname (const char *name);
  int   streamOnFile (const char *file, LPATTACH att);
  int   streamFromFile (const char *file, LPATTACH att);
  int   encryptAttachments (HWND hwnd);
  void  decryptAttachments (HWND hwnd, GpgMsg *msg);
  int   signAttachments (HWND hwnd);
  LPATTACH openAttachment (GpgMsg *msg, int pos);
  void  releaseAttachment (LPATTACH att);
  int   processAttachment (LPATTACH *att, HWND hwnd, int pos, int action);
  bool  signAttachment (const char *datfile);

};


/* Create an instance of the MapiGPGME class. */
MapiGPGME *
CreateMapiGPGME (void)
{
  return new MapiGPGMEImpl ();
}


/* Early initialization of this module.  This is done right at startup
   with only one thread running.  Should be called only once. Returns
   0 on success. */
int
initialize_mapi_gpgme (void)
{
  SECURITY_ATTRIBUTES sa;
  
  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = TRUE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  log_mutex = CreateMutex (&sa, FALSE, NULL);
  return log_mutex? 0 : -1;
}

/* Acquire the mutex for logging.  Returns 0 on success. */
static int 
lock_log (void)
{
  int code = WaitForSingleObject (log_mutex, INFINITE);
  return code != WAIT_OBJECT_0;
}

/* Release the mutex for logging. No error return is done because this
   is a fatal error anyway and we have no means for proper
   notification. */
static void
unlock_log (void)
{
  ReleaseMutex (log_mutex);
}




static void
do_log (const char *fmt, va_list a, int w32err, int err,
        const void *buf, size_t buflen)
{
  if (enable_logging == false || !logfile)
    return;

  if (!logfp)
    {
      if ( !lock_log ())
        {
          if (!logfp)
            logfp = fopen (logfile, "a+");
          unlock_log ();
        }
    }
  if (!logfp)
    return;
  fprintf (logfp, "%lu/", (unsigned long)GetCurrentThreadId ());
  if (err == 1)
    fputs ("ERROR/", logfp);
  vfprintf (logfp, fmt, a);
  if (w32err) 
    {
      char buf[256];
      
      FormatMessage (FORMAT_MESSAGE_FROM_SYSTEM, NULL, w32err, 
                     MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), 
                     buf, sizeof (buf)-1, NULL);
      fputs (": ", logfp);
      fputs (buf, logfp);
    }
  if (buf)
    {
      const unsigned char *p = (const unsigned char*)buf;

      for ( ; buflen; buflen--, p++)
        fprintf (logfp, "%02X", *p);
      putc ('\n', logfp);
    }
  else if ( *fmt && fmt[strlen (fmt) - 1] != '\n')
    putc ('\n', logfp);

  fflush (logfp);
}


void 
log_debug (const char *fmt, ...)
{
  va_list a;
  
  va_start (a, fmt);
  do_log (fmt, a, 0, 0, NULL, 0);
  va_end (a);
}

void 
log_error (const char *fmt, ...)
{
  va_list a;
  
  va_start (a, fmt);
  do_log (fmt, a, 0, 1, NULL, 0);
  va_end (a);
}

void 
log_vdebug (const char *fmt, va_list a)
{
  do_log (fmt, a, 0, 0, NULL, 0);
}


void 
log_debug_w32 (int w32err, const char *fmt, ...)
{
  va_list a;

  if (w32err == -1)
      w32err = GetLastError ();
  
  va_start (a, fmt);
  do_log (fmt, a, w32err, 0, NULL, 0);
  va_end (a);
}

void 
log_error_w32 (int w32err, const char *fmt, ...)
{
  va_list a;

  if (w32err == -1)
      w32err = GetLastError ();
  
  va_start (a, fmt);
  do_log (fmt, a, w32err, 1, NULL, 0);
  va_end (a);
}


void 
log_hexdump (const void *buf, size_t buflen, const char *fmt, ...)
{
  va_list a;

  va_start (a, fmt);
  do_log (fmt, a, 0, 2, buf, buflen);
  va_end (a);
}







void 
MapiGPGMEImpl::cleanupTempFiles (void)
{
  HANDLE hd;
  WIN32_FIND_DATA fnd;
  char path[MAX_PATH+32], tmp[MAX_PATH+32];
  
  assert (strlen (ATT_PREFIX) + 2 < 16 /* just a reasonable value*/ );
  
  *path = 0;
  GetTempPath (sizeof (path)-4, path);
  if (*path && path[strlen (path)-1] != '\\')
    strcat (path, "\\");
  strcpy (tmp, path);
  strcat (path, "*"ATT_PREFIX"*");
  hd = FindFirstFile (path, &fnd);
  if (hd == INVALID_HANDLE_VALUE)
    return;

  do
    {
      char *p = (char *)xmalloc (strlen (tmp) + strlen (fnd.cFileName) + 2);
      strcpy (stpcpy (p, tmp), fnd.cFileName);
      log_debug ("deleting temp file `%s'\n", p);
      DeleteFile (p);
      xfree (p);
    } 
  while (FindNextFile (hd, &fnd) == TRUE);
  FindClose (hd);
}


/* Update the display using the message MSG.  Return 0 on success. */
static int
update_display (HWND hwnd, GpgMsg *msg)
{
  HWND window;

  window = find_message_window (hwnd, msg);
  if (window)
    {
      log_debug ("%s:%s: window handle %p\n", __FILE__, __func__, window);
      SetWindowText (window, msg->getDisplayText ());
      log_debug ("%s:%s: window text is now `%s'",
                 __FILE__, __func__, msg->getDisplayText ());
      return 0;
    }
  else
    {
      log_debug ("%s: window handle not found for parent %p\n",
                 __func__, hwnd);
      return -1;
    }
}


static bool 
is_html_body (const char *body)
{
  char *p1, *p2;
  
  /* XXX: it is possible but unlikely that the message text
     contains the used keywords. */
  p1 = strstr (body, "<HTML>");
  p2 = strstr (body, "</HTML>");
  if (p1 && p2)
    return true;
  p1 = strstr (body, "<html>");
  p2 = strstr (body, "</html>");
  if (p1 && p2)
	return true;
  /* XXX: use case insentensive strstr version. */
  return false;
}




void 
MapiGPGMEImpl::freeKeyArray (void **key)
{
    gpgme_key_t *buf = (gpgme_key_t *)key;
    int i=0;

    if (buf == NULL)
	return;
    for (i = 0; buf[i] != NULL; i++) {
	gpgme_key_release (buf[i]);
	buf[i] = NULL;
    }
    xfree (buf);
}


/* Return the number of recipients in the array RECIPIENTS. */
static int 
count_recipients (char **recipients)
{
  int i;
  
  for (i=0; recipients[i] != NULL; i++)
    ;
  return i;
}



void
MapiGPGMEImpl::freeUnknownKeys (char **unknown, int n)
{    
    for (int i=0; i < n; i++) {
	if (unknown[i] != NULL) {
	    xfree (unknown[i]);
	    unknown[i] = NULL;
	}
    }
    if (n > 0)
	xfree (unknown);
}


/* Release an array of strings with recipient names. */
static void
free_recipient_array (char **recipients)
{
  int i;

  if (recipients)
    {
      for (i=0; recipients[i]; i++) 
	xfree (recipients[i]);	
      xfree (recipients);
    }
}



/* Create a new body from body wth suitable line endings. aller must
   release the result. */
static char *
add_html_line_endings (const char *body)
{
  size_t count;
  const char *s;
  char *p, *result;

  for (count=0, s = body; *s; s++)
    if (*s == '\n')
      count++;
  
  result = (char*)xmalloc ((s - body) + count*10 + 1);
  
  for (s=body, p = result; *s; s++ )
    if (*s == '\n')
      p = stpcpy (p, "&nbsp;<br>\n");
    else
      *p++ = *s;
  *p = 0;
  
  return result;
  
}



int 
MapiGPGMEImpl::encrypt (HWND hwnd, GpgMsg *msg)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  gpgme_key_t *keys = NULL;
  gpgme_key_t *keys2 = NULL;
  bool is_html;
  const char *plaintext;
  char *ciphertext;
  char **recipients = msg->getRecipients ();
  char **unknown = NULL;
  int opts = 0;
  int err = 0;
  size_t all = 0;
  int n;
  

  if (!msg || !*(plaintext = msg->getOrigText ())) 
    {
      free_recipient_array (recipients);
      return 0;  /* Empty message - Nothing to encrypt. */
    }

  n = op_lookup_keys (recipients, &keys, &unknown, &all);
  log_debug ("%s: found %d need %d (%p)\n", __func__, n, all, unknown);

  if (n != count_recipients (recipients))
    {
      log_debug ("recipient_dialog_box2\n");
      recipient_dialog_box2 (keys, unknown, all, &keys2, &opts);
      xfree (keys);
      keys = keys2;
      if (opts & OPT_FLAG_CANCEL) 
        {
          free_recipient_array (recipients);
          return 0;
	}
    }

  err = op_encrypt ((void*)keys, plaintext, &ciphertext);
  if (err)
    MessageBox (NULL, op_strerror (err),
                "GPG Encryption", MB_ICONERROR|MB_OK);
  else 
    {
      if (is_html) 
        {
          msg->setCipherText (add_html_line_endings (ciphertext), true);
          ciphertext = NULL;
        }
      else
        msg->setCipherText (ciphertext, false);
      xfree (ciphertext);
  }

  free_recipient_array (recipients);
  freeUnknownKeys (unknown, n);
  if (!err && msg->hasAttachments ())
    {
      log_debug ("encrypt attachments\n");
      recipSet = (void *)keys;
      encryptAttachments (hwnd);
    }
  freeKeyArray ((void **)keys);
  return err;
}


/* Decrypt the message MSG and update the window.  HWND identifies the
   current window. */
int 
MapiGPGMEImpl::decrypt (HWND hwnd, GpgMsg * msg)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  openpgp_t mtype;
  char *plaintext = NULL;
  int has_attach;
  int err;

  mtype = msg->getMessageType ();
  if (mtype == OPENPGP_CLEARSIG)
    {
      log_debug ("%s:%s: leave (passing on to verify)\n", __FILE__, __func__);
      return verify (hwnd, msg);
    }

  /* Check whether this possibly encrypted message as attachments.  We
     check right now because we need to get into the decryptio code
     even if the body is not encrypted but attachments are
     available. FIXME: I am not sure whether this is the best
     solution, we might want to skip the decryption step later and
     also test for encrypted attachments right now.*/
  has_attach = msg->hasAttachments ();

  if (mtype == OPENPGP_NONE && !has_attach ) 
    {
      MessageBox (NULL, "No valid OpenPGP data found.",
                  "GPG Decryption", MB_ICONERROR|MB_OK);
      log_debug ("MapiGPGME.decrypt: leave (no OpenPGP data)\n");
      return 0;
    }

  err = op_decrypt_start (msg->getOrigText (), &plaintext, nstorePasswd);

  if (err)
    {
      if (has_attach && gpg_err_code (err) == GPG_ERR_NO_DATA)
        ;
      else
        MessageBox (NULL, op_strerror (err),
                    "GPG Decryption", MB_ICONERROR|MB_OK);
    }
  else if (plaintext && *plaintext)
    {	
      int is_html = is_html_body (plaintext);

      msg->setPlainText (plaintext);
      plaintext = NULL;

      /* Also set PR_BODY but do not use 'SaveChanges' to make it
         permanently.  This way the user can reply with the
         plaintext but the ciphertext is still stored. */
      log_debug ("decrypt isHtml=%d\n", is_html);

      /* XXX: find a way to handle text/html message in a better way! */
      if (is_html || update_display (hwnd, msg))
        {
          const char s[] = 
            "The message text cannot be displayed.\n"
            "You have to save the decrypted message to view it.\n"
            "Then you need to re-open the message.\n\n"
            "Do you want to save the decrypted message?";
          int what;
          
          what = MessageBox (NULL, s, "GPG Decryption",
                             MB_YESNO|MB_ICONWARNING);
          if (what == IDYES) 
            {
              log_debug ("decrypt: saving plaintext message.\n");
              msg->saveChanges (1);
            }
	}
      else
        msg->saveChanges (0);
    }

  if (has_attach)
    {
      log_debug ("decrypt attachments\n");
      decryptAttachments (hwnd, msg);
    }

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
}


/* Sign the current message. Returns 0 on success. */
int
MapiGPGMEImpl::sign (HWND hwnd, GpgMsg * msg)
{
  log_debug ("%s.%s: enter\n", __FILE__, __func__);
  char *signedtext;
  int err = 0;

  /* We don't sign an empty body - a signature on a zero length string
     is pretty much useless. */
  if (!msg || !*msg->getOrigText ()) 
    {
      log_debug ("MapiGPGME.sign: leave (empty message)\n");
      return 0;
    }

  err = op_sign_start (msg->getOrigText (), &signedtext);
  if (err)
    MessageBox (NULL, op_strerror (err), "GPG Sign", MB_ICONERROR|MB_OK);
  else
    {
      msg->setSignedText (signedtext);
    }
  
  if (msg->hasAttachments () && autoSignAtt)
    signAttachments (hwnd);

  log_debug ("%s:%s: leave (rc=%d)\n", err);
  return err;
}


/* Perform a sign+encrypt operation using the data already store in
   the instance variables. Return 0 on success. */
int
MapiGPGMEImpl::signEncrypt (HWND hwnd, GpgMsg *msg)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  char **recipients = msg->getRecipients ();
  char **unknown = NULL;
  gpgme_key_t locusr=NULL, *keys = NULL, *keys2 =NULL;
  char *ciphertext = NULL;
  int is_html;

  /* Check for trivial case: We don't encrypt or sign an empty
     message. */
  if (!msg || !*msg->getOrigText ()) 
    {
      free_recipient_array (recipients);
      log_debug ("%s.%s: leave (empty message)\n", __FILE__, __func__);
      return 0;  /* Empty message - Nothing to encrypt. */
    }

  /* Pop up a dialog box to ask for the signer of the message. */
  if (signer_dialog_box (&locusr, NULL) == -1)
    {
      free_recipient_array (recipients);
      log_debug ("%s.%s: leave (dialog failed)\n", __FILE__, __func__);
      return 0;
    }

  /* Now lookup the keys of all recipients. */
  size_t all;
  int n = op_lookup_keys (recipients, &keys, &unknown, &all);
  if (n != count_recipients (recipients)) 
    {
      /* FIXME: The implementation is not correct: op_lookup_keys
         returns the number of missing keys but we compare against the
         total number of keys; thus the box would pop up even when all
         have been found. */
	recipient_dialog_box2 (keys, unknown, all, &keys2, NULL);
	xfree (keys);
	keys = keys2;
    }
  
  log_key_info (this, "signEncrypt", keys, locusr);

  is_html = is_html_body (msg->getOrigText ());

    /* FIXME: Remove the cast and use proper types. */
  int err = op_sign_encrypt ((void *)keys, (void*)locusr,
                             msg->getOrigText (), &ciphertext);
  if (err)
    MessageBox (NULL, op_strerror (err),
                "GPG Sign Encrypt", MB_ICONERROR|MB_OK);
  else 
    {
      if (is_html) 
        {
          msg->setCipherText (add_html_line_endings (ciphertext), true);
          ciphertext = NULL;
        }
      else
        msg->setCipherText (ciphertext, false);
      xfree (ciphertext);
    }

  freeUnknownKeys (unknown, n);
  if (!err && msg->hasAttachments ())
    {
      log_debug ("%s:%s: have attachments\n", __FILE__, __func__);
      recipSet = (void *)keys;
      encryptAttachments (hwnd);
    }
  freeKeyArray ((void **)keys);
  gpgme_key_release (locusr);
  free_recipient_array (recipients);
  log_debug ("%s:%s: leave (rc=%d)\n", err);
  return err;
}


int 
MapiGPGMEImpl::verify (HWND hwnd, GpgMsg *msg)
{
  log_debug ("%s:%s: enter\n", __FILE__, __func__);
  openpgp_t mtype;
  int has_attach;
  
  mtype = msg->getMessageType ();
  has_attach = msg->hasAttachments ();
  if (mtype == OPENPGP_NONE && !has_attach ) 
    {
      log_debug ("%s:%s: leave (no OpenPGP data)\n", __FILE__, __func__);
      return 0;
    }

  int err = op_verify_start (msg->getOrigText (), NULL);
  if (err)
    MessageBox (NULL, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);
  else
    update_display (hwnd, msg);

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
}




int
MapiGPGMEImpl::doCmdFile(int action, const char *in, const char *out)
{
    log_debug ( "doCmdFile action=%d in=%s out=%s\n", action, in, out);
    if (ATT_SIGN (action) && ATT_ENCR (action))
	return !op_sign_encrypt_file (recipSet, in, out);
    if (ATT_SIGN (action) && !ATT_ENCR (action))
	return !op_sign_file (OP_SIG_NORMAL, in, out, nstorePasswd);
    if (!ATT_SIGN (action) && ATT_ENCR (action))
	return !op_encrypt_file (recipSet, in, out);
    return !op_decrypt_file (in, out);    
}


int
MapiGPGMEImpl::doCmdAttach (int action)
{
//     log_debug ("doCmdAttach action=%d\n", action);
//     if (ATT_SIGN (action) && ATT_ENCR (action))
// 	return signEncrypt ();
//     if (ATT_SIGN (action) && !ATT_ENCR (action))
// 	return sign ();
//     if (!ATT_SIGN (action) && ATT_ENCR (action))
// 	return encrypt ();
// FIXME!!!
  return 0; /*decrypt ();*/
}



static const char *
userid_from_key (gpgme_key_t k)
{
  if (k && k->uids && k->uids->uid)
    return k->uids->uid;
  else
    return "?";
}

static const char *
keyid_from_key (gpgme_key_t k)
{
  
  if (k && k->subkeys && k->subkeys->keyid)
    return k->subkeys->keyid;
  else
    return "????????";
}

/* Print information on the keys in KEYS and LOCUSR. */
static void 
log_key_info (MapiGPGME *g, const char *prefix,
              gpgme_key_t *keys, gpgme_key_t locusr)
{
    if (locusr)
      log_debug ("MapiGPGME.%s: signer: 0x%s %s\n", prefix,
                   keyid_from_key (locusr), userid_from_key (locusr));
    
    else
      log_debug ("MapiGPGME.%s: no signer\n", prefix);
    
    gpgme_key_t n;
    int i;

    if (keys == NULL)
	return;
    i=0;
    for (n=keys[0]; keys[i] != NULL; i++)
      log_debug ("MapiGPGME.%s: recp.%d 0x%s %s\n", prefix,
                   i, keyid_from_key (keys[i]), userid_from_key (keys[i]));
}
	    


void 
MapiGPGMEImpl::setDefaultKey (const char *key)
{
    xfree (defaultKey);
    defaultKey = xstrdup (key);
}


const char* 
MapiGPGMEImpl::getDefaultKey (void)
{
  return defaultKey;
}



/* We need this to find the mailer window because we directly change
   the text of the window instead of the MAPI object itself. */
static HWND
find_message_window (HWND parent, GpgMsg *msg)
{
  HWND child;

  if (!parent)
    return NULL;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      char buf[1024+1];
      HWND w;

      memset (buf, 0, sizeof (buf));
      GetWindowText (child, buf, sizeof (buf)-1);
      if (msg->matchesString (buf))
        return child;
      w = find_message_window (child, msg);
      if (w)
        return w;
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


int
MapiGPGMEImpl::streamFromFile (const char *file, LPATTACH att)
{
    HRESULT hr;
    LPSTREAM to = NULL, from = NULL;
    STATSTG statInfo;

    hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 0,
	 		      MAPI_CREATE | MAPI_MODIFY, (LPUNKNOWN*) &to);
    if (FAILED (hr))
	return FALSE;

    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
			   STGM_READ, (char*)file, NULL, &from);
    if (!SUCCEEDED (hr)) {
	to->Release ();
	log_debug ( "streamFromFile %s failed.\n", file);
	return FALSE;
    }
    from->Stat (&statInfo, STATFLAG_NONAME);
    from->CopyTo (to, statInfo.cbSize, NULL, NULL);
    to->Commit (0);
    to->Release ();
    from->Release ();
    log_debug ( "streamFromFile %s succeeded\n", file);
    return TRUE;
}


int
MapiGPGMEImpl::streamOnFile (const char *file, LPATTACH att)
{
    HRESULT hr;
    LPSTREAM from = NULL, to = NULL;
    STATSTG statInfo;

    hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
			    0, 0, (LPUNKNOWN*) &from);
    if (FAILED (hr))
	return FALSE;

    hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
   		           STGM_CREATE | STGM_READWRITE, (char*) file,
			   NULL, &to);
    if (!SUCCEEDED (hr)) {
	from->Release ();
	log_debug ( "streamOnFile %s failed with %s\n", file, 
		    hr == MAPI_E_NO_ACCESS? 
		    "no access" : hr == MAPI_E_NOT_FOUND? "not found" : "unknown");
	return FALSE;
    }
    from->Stat (&statInfo, STATFLAG_NONAME);
    from->CopyTo (to, statInfo.cbSize, NULL, NULL);
    to->Commit (0);
    to->Release ();
    from->Release ();
    log_debug ( "streamOnFile %s succeeded\n", file);
    return TRUE;
}


int
MapiGPGMEImpl::getMessageFlags (GpgMsg * msg)
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int flags = 0;

    // FIXME
//     hr = HrGetOneProp (msg, PR_MESSAGE_FLAGS, &propval);
    if (FAILED (hr))
	return 0;
    flags = propval->Value.l;
    MAPIFreeBuffer (propval);
    return flags;
}


int
MapiGPGMEImpl::getMessageHasAttachments (GpgMsg * msg)
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int nattach = 0;

    // FIXME
//     hr = HrGetOneProp (msg, PR_HASATTACH, &propval);
    if (FAILED (hr))
	return 0;
    nattach = propval->Value.b? 1 : 0;
    MAPIFreeBuffer (propval);
    return nattach;   
}



bool
MapiGPGMEImpl::setAttachMethod (LPATTACH obj, int mode)
{
    SPropValue prop;
    HRESULT hr;
    prop.ulPropTag = PR_ATTACH_METHOD;
    prop.Value.ul = mode;
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? true : false;
}


int
MapiGPGMEImpl::getAttachMethod (LPATTACH obj)
{
    HRESULT hr;
    LPSPropValue propval = NULL;
    int method = 0;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_METHOD, &propval);
    if (FAILED (hr))
	return 0;
    method = propval->Value.ul;
    MAPIFreeBuffer (propval);
    return method;
}

bool
MapiGPGMEImpl::setAttachFilename (LPATTACH obj, const char *name, bool islong)
{
    HRESULT hr;
    SPropValue prop;
    prop.ulPropTag = PR_ATTACH_LONG_FILENAME;

    if (!islong)
	prop.ulPropTag = PR_ATTACH_FILENAME;
    prop.Value.lpszA = (char*) name;   
    hr = HrSetOneProp (obj, &prop);
    return FAILED (hr)? false: true;
}


char*
MapiGPGMEImpl::getAttachPathname (LPATTACH obj)
{
    LPSPropValue propval;
    HRESULT hr;
    char *path;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_PATHNAME, &propval);
    if (FAILED (hr)) {
	hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_PATHNAME, &propval);
	if (SUCCEEDED (hr)) {
	    path = xstrdup (propval[0].Value.lpszA);
	    MAPIFreeBuffer (propval);
	}
	else
	    return NULL;
    }
    else {
	path = xstrdup (propval[0].Value.lpszA);
	MAPIFreeBuffer (propval);
    }
    return path;
}


char*
MapiGPGMEImpl::getAttachFilename (LPATTACH obj)
{
    LPSPropValue propval;
    HRESULT hr;
    char *name;

    hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_LONG_FILENAME, &propval);
    if (FAILED(hr)) {
	hr = HrGetOneProp ((LPMAPIPROP)obj, PR_ATTACH_FILENAME, &propval);
	if (SUCCEEDED (hr)) {
	    name = xstrdup (propval[0].Value.lpszA);
	    MAPIFreeBuffer (propval);
	}
	else
	    return NULL;
    }
    else {
	name = xstrdup (propval[0].Value.lpszA);
	MAPIFreeBuffer (propval);
    }
    return name;
}


bool
MapiGPGMEImpl::checkAttachmentExtension (const char *ext)
{
    if (ext == NULL)
	return false;
    if (*ext == '.')
	ext++;
    log_debug ( "checkAttachmentExtension: %s\n", ext);
    if (stricmp (ext, "gpg") == 0 ||
	stricmp (ext, "pgp") == 0 ||
	stricmp (ext, "asc") == 0)
	return true;
    return false;
}


const char*
MapiGPGMEImpl::getAttachmentExtension (const char *fname)
{
    static char ext[4];
    char *p;

    p = strrchr (fname, '.');
    if (p != NULL) {
	/* XXX: what if the extension is < 3 chars */
	strncpy (ext, p, 4);
	if (checkAttachmentExtension (ext))
	    return ext;
    }
    return EXT_MSG;
}


const char*
MapiGPGMEImpl::getPGPExtension (int action)
{
    if (ATT_SIGN (action))
	return EXT_SIG;
    return EXT_MSG;
}


bool 
MapiGPGMEImpl::setXHeader (GpgMsg * msg, const char *name, const char *val)
{  
#ifndef __MINGW32__
    USES_CONVERSION;
#endif
    LPMDB lpMdb = NULL;
    HRESULT hr = 0;
    LPSPropTagArray pProps = NULL;
    SPropValue pv;
    MAPINAMEID mnid[1];	
    // {00020386-0000-0000-C000-000000000046}  ->  GUID For X-Headers	
    GUID guid = {0x00020386, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00,
		 0x00, 0x00, 0x00, 0x46} };

    memset (&mnid[0], 0, sizeof (MAPINAMEID));
    mnid[0].lpguid = &guid;
    mnid[0].ulKind = MNID_STRING;
//     mnid[0].Kind.lpwstrName = A2W (name);
// FIXME
//    hr = msg->GetIDsFromNames (1, (LPMAPINAMEID*)mnid, MAPI_CREATE, &pProps);
    if (FAILED (hr)) {
	log_debug ("set X-Header failed.\n");
	return false;
    }
    
    pv.ulPropTag = (pProps->aulPropTag[0] & 0xFFFF0000) | PT_STRING8;
    pv.Value.lpszA = (char *)val;
//FIXME     hr = HrSetOneProp(msg, &pv);	
    if (!SUCCEEDED (hr)) {
	log_debug ("set X-Header failed.\n");
	return false;
    }

    log_debug ("set X-Header succeeded.\n");
    return true;
}


char*
MapiGPGMEImpl::getXHeader (GpgMsg *msg, const char *name)
{
    /* XXX: PR_TRANSPORT_HEADERS is not available in my MSDN. */
    return NULL;
}


void
MapiGPGMEImpl::freeAttachments (void)
{
    if (attachTable != NULL) {
        attachTable->Release ();
	attachTable = NULL;
    }
    if (attachRows != NULL) {
        FreeProws (attachRows);
	attachRows = NULL;
    }
}




void
MapiGPGMEImpl::releaseAttachment (LPATTACH att)
{
    att->Release ();
}



char*
MapiGPGMEImpl::generateTempname (const char *name)
{
    char temp[MAX_PATH+2];
    char *p;

    GetTempPath (sizeof (temp)-1, temp);
    if (temp[strlen (temp)-1] != '\\')
	strcat (temp, "\\");
    p = (char *)xcalloc (1, strlen (temp) + strlen (name) + 16);
    sprintf (p, "%s%s", temp, name);
    return p;
}


bool
MapiGPGMEImpl::signAttachment (const char *datfile)
{
    char *sigfile;
    LPATTACH newatt;
    int pos=0, err=0;

    log_debug ("MapiGPGME.%s:%d: enter\n", __func__, __LINE__);

    sigfile = (char *)xcalloc (1,strlen (datfile)+5);
    strcpy (sigfile, datfile);
    strcat (sigfile, ".asc");

    newatt = createAttachment (NULL/*FIXME*/, pos);
    setAttachMethod (newatt, ATTACH_BY_VALUE);
    setAttachFilename (newatt, sigfile, false);

    err = op_sign_file (OP_SIG_DETACH, datfile, sigfile, nstorePasswd);

    if (streamFromFile (sigfile, newatt)) {
	log_debug ("signAttachment: commit changes.\n");
	newatt->SaveChanges (FORCE_SAVE);
    }
    releaseAttachment (newatt);
    xfree (sigfile);

    log_debug ("MapiGPGME.%s: leave (rc=%d)\n", __func__, err);

    return (!err)? true : false;
}

/* XXX: find a way to see if the attachment is already secured. This could be
        done by watching at the extension or checking the first lines. */
int
MapiGPGMEImpl::processAttachment (LPATTACH *attm, HWND hwnd,
                                  int pos, int action)
{    
    LPATTACH att = *attm;
    int method = getAttachMethod (att);
    BOOL success = TRUE;
    HRESULT hr;

    /* XXX: sign-only code is still not very intuitive. */

    if (action == GPG_ATTACH_NONE)
	return FALSE;
    if (action == GPG_ATTACH_DECRYPT && !saveDecryptedAtt)
	return TRUE;

    switch (method) {
    case ATTACH_EMBEDDED_MSG:
	LPMESSAGE emb;

	/* we do not support to sign these kind of attachments. */
	if (action == GPG_ATTACH_SIGN)
	    return TRUE;
	hr = att->OpenProperty (PR_ATTACH_DATA_OBJ, &IID_IMessage, 0, 
			        MAPI_MODIFY, (LPUNKNOWN*) &emb);
	if (FAILED (hr))
	    return FALSE;
//FIXME 	setWindow (hwnd);
// 	setMessage (emb);
	if (doCmdAttach (action))
	    success = FALSE;
	emb->SaveChanges (FORCE_SAVE);
	att->SaveChanges (FORCE_SAVE);
	emb->Release ();
	break;

    case ATTACH_BY_VALUE:
	char *inname;
	char *outname;
	char *tmp;

	tmp = getAttachFilename (att);
	inname =  generateTempname (tmp);
	log_debug ("enc inname: '%s'\n", inname);
	if (action != GPG_ATTACH_DECRYPT) {
	    char *tmp2 = (char *)xcalloc (1, strlen (inname) 
					     + strlen (ATT_PREFIX) + 4 + 1);
	    sprintf (tmp2, "%s"ATT_PREFIX".%s", tmp, getPGPExtension (action));
	    outname = generateTempname (tmp2);
	    xfree (tmp2);
	    log_debug ( "enc outname: '%s'\n", outname);
	}
	else {
	    if (checkAttachmentExtension (strrchr (tmp, '.')) == false) {
		log_debug ( "%s: no pgp extension found.\n", tmp);
		xfree (tmp);
		xfree (inname);
		return TRUE;
	    }
	    char *tmp2 = (char*)xcalloc (1, strlen (tmp) + 4);
	    strcpy (tmp2, tmp);
	    tmp2[strlen (tmp2) - 4] = '\0';
	    outname = generateTempname (tmp2);
	    xfree (tmp2);
	    log_debug ("dec outname: '%s'\n", outname);
	}
	success = FALSE;
	/* if we are in sign-only mode, just create a detached signature
	   for each attachment but do not alter the attachment data itself. */
	if (action != GPG_ATTACH_SIGN && streamOnFile (inname, att)) {
	    if (doCmdFile (action, inname, outname))
		success = TRUE;
	    else
		log_debug ( "doCmdFile failed\n");
	}
	if ((action == GPG_ATTACH_ENCRYPT || action == GPG_ATTACH_SIGN) 
	    && autoSignAtt)
	    signAttachment (inname);

	/*DeleteFile (inname);*/
	/* XXX: the file does not seemed to be closed. */
	xfree (inname);
	xfree (tmp);
	
	if (action != GPG_ATTACH_SIGN)
	    deleteAttachment (NULL/*FIXME*/,pos);

	if (action == GPG_ATTACH_ENCRYPT) {
	    LPATTACH newatt;
	    *attm = newatt = createAttachment (NULL/*FIXME*/,pos);
	    setAttachMethod (newatt, ATTACH_BY_VALUE);
	    setAttachFilename (newatt, outname, false);

	    if (streamFromFile (outname, newatt)) {
		log_debug ( "commit changes.\n");	    
		newatt->SaveChanges (FORCE_SAVE);
	    }
	}
	else if (success && action == GPG_ATTACH_DECRYPT) {
          //	    success = saveDecryptedAttachment (NULL, outname);
	    log_debug ("saveDecryptedAttachment ec=%d\n", success);
	}
	DeleteFile (outname);
	xfree (outname);
	releaseAttachment (att);
	break;

    case ATTACH_BY_REF_ONLY:
	break;

    case ATTACH_OLE:
	break;

    }

    return success;
}


/* Decrypt all attachemnts of message MSG.  HWND is the usual window
   handle. */
void 
MapiGPGMEImpl::decryptAttachments (HWND hwnd, GpgMsg *msg)
{
  unsigned int n;
  LPATTACH amsg;
  

  n = msg->getAttachments ();
  log_debug ("%s:%s: message has %u attachments\n",
             __FILE__, __func__, n);
  if (!n)
    return;
  
  for (int i=0; i < n; i++) 
    {
      //amsg = openAttachment (NULL/*FIXME*/,i);
      if (amsg)
        processAttachment (&amsg, hwnd, i, GPG_ATTACH_DECRYPT);
    }
  return;
}



int
MapiGPGMEImpl::signAttachments (HWND hwnd)
{
//FIXME     if (!getAttachments ()) {
//         log_debug ("MapiGPGME.signAttachments: getAttachments failed\n");
// 	return FALSE;
//     }
    
    int n = 0/*FIXME countAttachments ()*/;
    log_debug ("MapiGPGME.signAttachments: mail has %d attachments\n", n);
    if (!n) {
	freeAttachments ();
	return TRUE;
    }
    for (int i=0; i < n; i++) {
        LPATTACH amsg = openAttachment (NULL/*FIXME*/,i);
	if (!amsg)
	    continue;
	processAttachment (&amsg, hwnd, i, GPG_ATTACH_SIGN);
	releaseAttachment (amsg);
    }
    freeAttachments ();
    return 0;
}


int
MapiGPGMEImpl::encryptAttachments (HWND hwnd)
{    
    unsigned int n;

//     n = if (!getAttachments ())
// 	return FALSE;
    n = 0 /*FIXMEcountAttachments ()*/;
    log_debug ("enc: mail has %d attachments\n", n);
    if (!n) {
	freeAttachments ();
	return TRUE;
    }
    for (int i=0; i < n; i++) {
      LPATTACH amsg = openAttachment (NULL/*FIXME*/,i);
	if (amsg == NULL)
	    continue;
	processAttachment (&amsg, hwnd, i, GPG_ATTACH_ENCRYPT);
	releaseAttachment (amsg);	
    }
    freeAttachments ();
    return 0;
}



void  
MapiGPGMEImpl::showVersion (void)
{
  /* Not yet available. */
}


int
MapiGPGMEImpl::startKeyManager (void)
{
    return start_key_manager ();
}


void
MapiGPGMEImpl::startConfigDialog (HWND parent)
{
    config_dialog_box (parent);
}


int
MapiGPGMEImpl::readOptions (void)
{
    char *val=NULL;

    load_extension_value ("autoSignAttachments", &val);
    autoSignAtt = val == NULL || *val != '1' ? 0 : 1;
    xfree (val); val =NULL;

    load_extension_value ("saveDecryptedAttachments", &val);
    saveDecryptedAtt = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val =NULL;

    load_extension_value ("encryptDefault", &val);
    doEncrypt = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val=NULL;

    load_extension_value ("signDefault", &val);
    doSign = val == NULL || *val != '1'? 0 : 1;
    xfree (val); val = NULL;

    load_extension_value ("addDefaultKey", &val);
    encryptDefault = val == NULL || *val != '1' ? 0 : 1;
    xfree (val); val = NULL;

    load_extension_value ("storePasswdTime", &val);
    nstorePasswd = val == NULL || *val == '0'? 0 : atol (val);
    xfree (val); val = NULL;

    load_extension_value ("encodingFormat", &val);
    encFormat = val == NULL? GPG_FMT_CLASSIC  : atol (val);
    xfree (val); val = NULL;

    load_extension_value ("logFile", &val);
    if (val == NULL ||*val == '"' || *val == 0) {
        setLogFile (NULL);
    }
    else {
	setLogFile (val);
	setEnableLogging (true);
    }
    xfree (val); val=NULL;

    load_extension_value ("defaultKey", &val);
    if (val == NULL || *val == '"') {
	encryptDefault = 0;
        xfree (defaultKey);
	defaultKey = NULL;
    }
    else {
	setDefaultKey (val);
	encryptDefault = 1;
    }

    xfree (val); val=NULL;

    return 0;
}


void
MapiGPGMEImpl::displayError (HWND root, const char *title)
{	
    char buf[256];
    
    FormatMessage (FORMAT_MESSAGE_FROM_SYSTEM, NULL, GetLastError (), 
		   MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT), 
		   buf, sizeof (buf)-1, NULL);
    MessageBox (root, buf, title, MB_OK|MB_ICONERROR);
}


int
MapiGPGMEImpl::writeOptions (void)
{
    struct conf {
	const char *name;
	bool value;
    };
    struct conf opt[] = {
	{"encryptDefault", doEncrypt},
	{"signDefault", doSign},
	{"addDefaultKey", encryptDefault},
	{"saveDecryptedAttachments", saveDecryptedAtt},
	{"autoSignAttachments", autoSignAtt},
	{NULL, 0}
    };
    char buf[32];

    for (int i=0; opt[i].name != NULL; i++) {
	int rc = store_extension_value (opt[i].name, opt[i].value? "1": "0");
	if (rc)
	    displayError (NULL, "Save options in the registry");
	/* XXX: also show the name of the value */
    }

    if (logfile != NULL)
	store_extension_value ("logFile", logfile);
    if (defaultKey != NULL)
	store_extension_value ("defaultKey", defaultKey);
    
    sprintf (buf, "%d", nstorePasswd);
    store_extension_value ("storePasswdTime", buf);
    
    sprintf (buf, "%d", encFormat);
    store_extension_value ("encodingFormat", buf);

    return 0;
}


int 
MapiGPGMEImpl::attachPublicKey (const char *keyid)
{
    /* @untested@ */
    const char *patt[1];
    char *keyfile;
    int err, pos = 0;
    LPATTACH newatt;

    keyfile = generateTempname (keyid);
    patt[0] = xstrdup (keyid);
    err = op_export_keys (patt, keyfile);

    newatt = createAttachment (NULL/*FIXME*/,pos);
    setAttachMethod (newatt, ATTACH_BY_VALUE);
    setAttachFilename (newatt, keyfile, false);
    /* XXX: set proper RFC3156 MIME types. */

    if (streamFromFile (keyfile, newatt)) {
	log_debug ("attachPublicKey: commit changes.\n");
	newatt->SaveChanges (FORCE_SAVE);
    }
    releaseAttachment (newatt);
    xfree (keyfile);
    xfree ((void *)patt[0]);
    return err;
}
