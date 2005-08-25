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
    
  int __stdcall decrypt (HWND hwnd, GpgMsg *msg);
  int __stdcall signEncrypt (HWND hwnd, GpgMsg *msg);
  int __stdcall verify (HWND hwnd, GpgMsg *msg);
  int __stdcall attachPublicKey (const char *keyid);

  int __stdcall doCmdAttach(int action);

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

  void freeUnknownKeys (char **unknown, int n);
  void freeKeyArray (void **key);
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
  sa.bInheritHandle = FALSE;
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

  err = op_decrypt (msg->getOrigText (), &plaintext, nstorePasswd);

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
      unsigned int n;
      
      n = msg->getAttachments ();
      log_debug ("%s:%s: message has %u attachments\n", __FILE__, __func__, n);
      for (int i=0; i < n; i++) 
        msg->decryptAttachment (hwnd, i, true, nstorePasswd);
    }

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
}



/* Perform a sign+encrypt operation using the data already store in
   the instance variables. Return 0 on success. */
int
MapiGPGMEImpl::signEncrypt (HWND hwnd, GpgMsg *msg)
{
#if 0
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
      return 0;  /* Empty message - Nothing to encrypt.  FIXME: What
                    about attachments without a body - is that possible?  */
    }

  /* Pop up a dialog box to ask for the signer of the message. */
  if (signer_dialog_box (&locusr, NULL) == -1)
    {
      free_recipient_array (recipients);
      log_debug ("%s.%s: leave (dialog failed)\n", __FILE__, __func__);
      return gpg_error (GPG_ERR_CANCELED);  
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

  int err = op_encrypt (msg->getOrigText (), &ciphertext,
                        keys, locusr, nstorePasswd);
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
      unsigned int n;
      
      n = msg->getAttachments ();
      log_debug ("%s:%s: message has %u attachments\n", __FILE__, __func__, n);
      for (int i=0; !err && i < n; i++) 
        err = msg->encryptAttachment (hwnd, i, keys, locusr, nstorePasswd);
      if (err)
        MessageBox (NULL, op_strerror (err),
                    "GPG Attachment Encryption", MB_ICONERROR|MB_OK);
    }
  freeKeyArray ((void **)keys);
  gpgme_key_release (locusr);
  free_recipient_array (recipients);
  log_debug ("%s:%s: leave (rc=%d)\n", err);
  return err;
#endif
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

  int err = op_verify (msg->getOrigText (), NULL);
  if (err)
    MessageBox (NULL, op_strerror (err), "GPG Verify", MB_ICONERROR|MB_OK);
  else
    update_display (hwnd, msg);

  log_debug ("%s:%s: leave (rc=%d)\n", __FILE__, __func__, err);
  return err;
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
#if 0
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
#endif
    return -1;
    
}
