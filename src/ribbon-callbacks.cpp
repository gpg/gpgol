/* ribbon-callbacks.h - Callbacks for the ribbon extension interface
 *    Copyright (C) 2013 Intevation GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <olectl.h>
#include <stdio.h>
#include <string.h>
#include <gdiplus.h>

#include <objidl.h>

#include "ribbon-callbacks.h"
#include "gpgoladdin.h"
#include "util.h"

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"

#include "common.h"
#include "display.h"
#include "msgcache.h"
#include "engine.h"
#include "engine-assuan.h"
#include "mapihelp.h"
#include "mimemaker.h"
#include "filetype.h"
#include "gpgolstr.h"

/* Gets the context of a ribbon control. And prints some
   useful debug output */
HRESULT getContext (LPDISPATCH ctrl, LPDISPATCH *context)
{
  *context = get_oom_object (ctrl, "get_Context");
  log_debug ("%s:%s: contextObj: %s",
             SRCNAME, __func__, get_object_name (*context));
  return context ? S_OK : E_FAIL;
}

#define OP_ENCRYPT     1 /* Encrypt the data */
#define OP_SIGN        2 /* Sign the data */
#define OP_DECRYPT     1 /* Decrypt the data */
#define OP_VERIFY      2 /* Verify the data */
#define DATA_BODY      4 /* Use text body as data */
#define DATA_SELECTION 8 /* Use selection as data */

/* Read hfile in chunks of 4KB and writes them to the sink */
static int
copyFileToSink (HANDLE hFile, sink_t sink)
{
  char buf[4096];
  DWORD bytesRead = 0;
  do
    {
      if (!ReadFile (hFile, buf, sizeof buf, &bytesRead, NULL))
        {
          log_error ("%s:%s: Could not read source file.",
                     SRCNAME, __func__);
          return -1;
        }
      if (write_buffer (sink, bytesRead ? buf : NULL, bytesRead))
        {
          log_error ("%s:%s: Could not write out buffer",
                     SRCNAME, __func__);
          return -1;
        }
    }
  while (bytesRead);
  return 0;
}

static int
attachSignature (LPDISPATCH mailItem, char *subject, HANDLE hFileToSign,
                 protocol_t protocol, unsigned int session_number,
                 HWND curWindow, wchar_t *fileNameToSign, char *sender)
{
  wchar_t *sigName = NULL;
  wchar_t *sigFileName = NULL;
  HANDLE hSigFile = NULL;
  int rc = 0;
  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  engine_filter_t filter = NULL;

  memset (encsink, 0, sizeof *encsink);
  memset (sink, 0, sizeof *sink);

  /* Prepare a fresh filter */
  if ((rc = engine_create_filter (&filter, write_buffer_for_cb, sink)))
    {
      goto failure;
    }
  encsink->cb_data = filter;
  encsink->writefnc = sink_encryption_write;
  engine_set_session_number (filter, session_number);
  engine_set_session_title (filter, subject ? subject :_("GpgOL"));

  if (engine_sign_start (filter, curWindow, protocol, sender, &protocol))
    goto failure;

  sigName = get_pretty_attachment_name (fileNameToSign, protocol, 1);

  /* If we are unlucky the number of temporary file artifacts might
     differ for the signature and the encrypted file but we have
     to live with that. */
  sigFileName = get_tmp_outfile (sigName, &hSigFile);
  sink->cb_data = hSigFile;
  sink->writefnc = sink_file_write;

  if (!sigFileName)
    {
      log_error ("%s:%s: Could not get a decent attachment name",
                 SRCNAME, __func__);
      goto failure;
    }

  /* Reset the file to sign handle to the beginning of the file and
     copy it to the signature buffer */
  SetFilePointer (hFileToSign, 0, NULL, 0);
  if ((rc=copyFileToSink (hFileToSign, encsink)))
    goto failure;

  /* Lets hope the user did not select a huge file. We are hanging
     here until encryption is completed.. */
  if ((rc = engine_wait (filter)))
    goto failure;

  filter = NULL; /* Not valid anymore.  */
  encsink->cb_data = NULL; /* Not needed anymore.  */

  if (!sink->enc_counter)
    {
      log_error ("%s:%s: nothing received from engine", SRCNAME, __func__);
      goto failure;
    }

  /* Now we have an encrypted file behind encryptedFile. Let's add it */
  add_oom_attachment (mailItem, sigFileName);

failure:
  xfree (sigFileName);
  xfree (sigName);
  if (hSigFile)
    {
      CloseHandle (hSigFile);
      DeleteFileW (sigFileName);
    }
  return rc;
}

/* do_composer_action
   Encrypts / Signs text in an IInspector context.
   Depending on the flags either the
   active selection or the full body is encrypted.
   Combine OP_ENCRYPT and OP_SIGN if you want both.
*/

HRESULT
do_composer_action (LPDISPATCH ctrl, int flags)
{
  LPDISPATCH context = NULL;
  LPDISPATCH selection = NULL;
  LPDISPATCH wordEditor = NULL;
  LPDISPATCH application = NULL;
  LPDISPATCH mailItem = NULL;
  LPDISPATCH sender = NULL;
  LPDISPATCH recipients = NULL;

  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  char* senderAddr = NULL;
  char** recipientAddrs = NULL;
  LPSTREAM tmpstream = NULL;
  engine_filter_t filter = NULL;
  char* plaintext = NULL;
  int rc = 0;
  HRESULT hr;
  HWND curWindow;
  protocol_t protocol;
  unsigned int session_number;
  int i;
  STATSTG tmpStat;

  log_debug ("%s:%s: enter", SRCNAME, __func__);

  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;

  memset (encsink, 0, sizeof *encsink);
  memset (sink, 0, sizeof *sink);

  curWindow = get_oom_context_window (context);

  wordEditor = get_oom_object (context, "WordEditor");
  application = get_oom_object (wordEditor, "get_Application");
  selection = get_oom_object (application, "get_Selection");
  mailItem = get_oom_object (context, "CurrentItem");
  sender = get_oom_object (mailItem, "Session.CurrentUser");
  recipients = get_oom_object (mailItem, "Recipients");

  if (!wordEditor || !application || !selection || !mailItem ||
      !sender || !recipients)
    {
      MessageBox (NULL,
                  "Internal error in GpgOL.\n"
                  "Could not find all objects.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      log_error ("%s:%s: Could not find all objects.",
                 SRCNAME, __func__);
      goto failure;
    }

  if (flags & DATA_SELECTION)
    {
      plaintext = get_oom_string (selection, "Text");

      if (!plaintext || strlen (plaintext) <= 1)
        {
          MessageBox (NULL,
                      _("Please select text to encrypt."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }
  else if (flags & DATA_BODY)
    {
      plaintext = get_oom_string (mailItem, "Body");
      if (!plaintext || strlen (plaintext) <= 1)
        {
          MessageBox (NULL,
                      _("Textbody empty."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }

  /* Create a temporary sink to construct the encrypted data.  */
  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                         (SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE
                          | STGM_CREATE | STGM_READWRITE),
                         NULL, GpgOLStr("GPG"), &tmpstream);

  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create temp file: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      rc = -1;
      goto failure;
    }

  sink->cb_data = tmpstream;
  sink->writefnc = sink_std_write;

  /* Now lets prepare our encryption */
  session_number = engine_new_session_number ();

  /* Prepare the encryption sink */

  if (engine_create_filter (&filter, write_buffer_for_cb, sink))
    {
      goto failure;
    }

  encsink->cb_data = filter;
  encsink->writefnc = sink_encryption_write;

  engine_set_session_number (filter, session_number);
  engine_set_session_title (filter, _("GpgOL"));

  senderAddr = get_pa_string (sender, PR_SMTP_ADDRESS);

  if (flags & OP_ENCRYPT)
    {
      recipientAddrs = get_oom_recipients (recipients);

      if (!recipientAddrs || !(*recipientAddrs))
        {
          MessageBox (NULL,
                      _("Please add at least one recipent."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }

      if ((rc=engine_encrypt_prepare (filter, curWindow,
                                      PROTOCOL_UNKNOWN,
                                      (flags & OP_SIGN) ?
                                      ENGINE_FLAG_SIGN_FOLLOWS : 0,
                                      senderAddr, recipientAddrs,
                                      &protocol)))
        {
          log_error ("%s:%s: engine encrypt prepare failed : %s",
                     SRCNAME, __func__, gpg_strerror (rc));
          goto failure;
        }

      if ((rc=engine_encrypt_start (filter, 0)))
        {
          log_error ("%s:%s: engine encrypt start failed: %s",
                     SRCNAME, __func__, gpg_strerror (rc));
          goto failure;
        }
    }
  else
    {
      /* We could do some kind of clearsign / sign text as attachment here
      but it is error prone */
      if ((rc=engine_sign_opaque_start (filter, curWindow, PROTOCOL_UNKNOWN,
                                        senderAddr, &protocol)))
        {
          log_error ("%s:%s: engine sign start failed: %s",
                     SRCNAME, __func__, gpg_strerror (rc));
          goto failure;
        }
    }

  /* Write the text in the encryption sink. */
  rc = write_buffer (encsink, plaintext, strlen (plaintext));

  if (rc)
    {
      log_error ("%s:%s: writing tmpstream to encsink failed: %s",
                 SRCNAME, __func__, gpg_strerror (rc));
      goto failure;
    }
  /* Flush the encryption sink and wait for the encryption to get
     ready.  */
  if ((rc = write_buffer (encsink, NULL, 0)))
    goto failure;
  if ((rc = engine_wait (filter)))
    goto failure;
  filter = NULL; /* Not valid anymore.  */
  encsink->cb_data = NULL; /* Not needed anymore.  */

  if (!sink->enc_counter)
    {
      log_debug ("%s:%s: nothing received from engine", SRCNAME, __func__);
      goto failure;
    }

  /* Check the size of the encrypted data */
  tmpstream->Stat (&tmpStat, 0);

  if (tmpStat.cbSize.QuadPart > UINT_MAX)
    {
      log_error ("%s:%s: No one should write so large mails.",
                 SRCNAME, __func__);
      goto failure;
    }

  /* Copy the encrypted stream to the message editor.  */
  {
    LARGE_INTEGER off;
    ULONG nread;

    char buffer[(unsigned int)tmpStat.cbSize.QuadPart + 1];

    memset (buffer, 0, sizeof buffer);

    off.QuadPart = 0;
    hr = tmpstream->Seek (off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    hr = tmpstream->Read (buffer, sizeof (buffer) - 1, &nread);
    if (hr)
      {
        log_error ("%s:%s: IStream::Read failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    if (strlen (buffer) > 1)
      {
        if (flags & OP_SIGN)
          {
            /* When signing we append the signature after the body */
            unsigned int combinedSize = strlen (buffer) +
              strlen (plaintext) + 5;
            char combinedBody[combinedSize];
            memset (combinedBody, 0, combinedSize);
            snprintf (combinedBody, combinedSize, "%s\r\n\r\n%s", plaintext,
                      buffer);
            if (flags & DATA_SELECTION)
              put_oom_string (selection, "Text", combinedBody);
            else if (flags & DATA_BODY)
              put_oom_string (mailItem, "Body", combinedBody);

          }
        else if (protocol == PROTOCOL_SMIME)
          {
            unsigned int enclosedSize = strlen (buffer) + 34 + 31 + 1;
            char enclosedData[enclosedSize];
            snprintf (enclosedData, sizeof enclosedData,
                      "-----BEGIN ENCRYPTED MESSAGE-----\r\n"
                      "%s"
                      "-----END ENCRYPTED MESSAGE-----\r\n", buffer);
            if (flags & DATA_SELECTION)
              put_oom_string (selection, "Text", enclosedData);
            else if (flags & DATA_BODY)
              put_oom_string (mailItem, "Body", enclosedData);

          }
        else
          {
            if (flags & DATA_SELECTION)
              put_oom_string (selection, "Text", buffer);
            else if (flags & DATA_BODY)
              {
                put_oom_string (mailItem, "Body", buffer);
              }
          }
      }
    else
      {
        /* Just to be save not to overwrite the selection with
           an empty buffer */
        log_error ("%s:%s: unexpected problem ", SRCNAME, __func__);
        goto failure;
      }
  }

failure:
  if (rc)
    log_debug ("%s:%s: failed rc=%d (%s) <%s>", SRCNAME, __func__, rc,
               gpg_strerror (rc), gpg_strsource (rc));
  engine_cancel (filter);
  RELDISP(wordEditor);
  RELDISP(application);
  RELDISP(selection);
  RELDISP(sender);
  RELDISP(recipients);
  RELDISP(mailItem);
  RELDISP(tmpstream);
  xfree (plaintext);
  xfree (senderAddr);
  if (recipientAddrs)
    {
      for (i=0; recipientAddrs && recipientAddrs[i]; i++)
        xfree (recipientAddrs[i]);
      xfree (recipientAddrs);
    }
  log_debug ("%s:%s: leave", SRCNAME, __func__);

  return S_OK;
}

HRESULT
decryptAttachments (LPDISPATCH ctrl)
{
  LPDISPATCH context = NULL;
  LPDISPATCH attachmentSelection;
  int attachmentCount;
  HRESULT hr = 0;
  int i = 0;
  HWND curWindow;
  int err;

  hr = getContext(ctrl, &context);

  attachmentSelection = get_oom_object (context, "AttachmentSelection");
  if (!attachmentSelection)
    {
      /* We can be called from a context menu, in that case we
         directly have an AttachmentSelection context. Otherwise
         we have an Explorer context with an Attachment Selection property. */
      attachmentSelection = context;
    }

  attachmentCount = get_oom_int (attachmentSelection, "Count");

  curWindow = get_oom_context_window (context);

  {
    char *filenames[attachmentCount + 1];
    filenames[attachmentCount] = NULL;
    /* Yes the items start at 1! */
    for (i = 1; i <= attachmentCount; i++)
      {
        char buf[16];
        char *filename;
        wchar_t *wcsOutFilename;
        DISPPARAMS saveParams;
        VARIANT aVariant[1];
        LPDISPATCH attachmentObj;
        DISPID saveID;

        snprintf (buf, sizeof (buf), "Item(%i)", i);
        attachmentObj = get_oom_object (attachmentSelection, buf);
        if (!attachmentObj)
          {
            /* Should be impossible */
            filenames[i-1] = NULL;
            log_error ("%s:%s: could not find Item %i;",
                       SRCNAME, __func__, i);
            break;
          }
        filename = get_oom_string (attachmentObj, "FileName");

        saveID = lookup_oom_dispid (attachmentObj, "SaveAsFile");

        saveParams.rgvarg = aVariant;
        saveParams.rgvarg[0].vt = VT_BSTR;
        filenames[i-1] = get_save_filename (NULL, filename);
        xfree (filename);

        if (!filenames [i-1])
          continue;

        wcsOutFilename = utf8_to_wchar2 (filenames[i-1],
                                         strlen(filenames[i-1]));
        saveParams.rgvarg[0].bstrVal = SysAllocString (wcsOutFilename);
        saveParams.cArgs = 1;
        saveParams.cNamedArgs = 0;

        hr = attachmentObj->Invoke (saveID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                    DISPATCH_METHOD, &saveParams,
                                    NULL, NULL, NULL);
        SysFreeString (saveParams.rgvarg[0].bstrVal);
        RELDISP (attachmentObj);
        if (FAILED(hr))
          {
            int j;
            log_debug ("%s:%s: Saving to file failed. hr: %x",
                       SRCNAME, __func__, (unsigned int) hr);
            for (j = 0; j < i; j++)
              xfree (filenames[j]);
            RELDISP (attachmentSelection);
            return hr;
          }
      }
    RELDISP (attachmentSelection);
    err = op_assuan_start_decrypt_files (curWindow, filenames);
    for (i = 0; i < attachmentCount; i++)
      xfree (filenames[i]);
  }

  log_debug ("%s:%s: Leaving. Err: %i",
             SRCNAME, __func__, err);

  return S_OK; /* If we return an error outlook will show that our
                  callback function failed in an ugly window. */
}

/* do_reader_action
   decrypts the content of an inspector. Controled by flags
   similary to the do_composer_action.
*/

HRESULT
do_reader_action (LPDISPATCH ctrl, int flags)
{
  LPDISPATCH context = NULL;
  LPDISPATCH selection = NULL;
  LPDISPATCH wordEditor = NULL;
  LPDISPATCH mailItem = NULL;
  LPDISPATCH wordApplication = NULL;

  struct sink_s decsinkmem;
  sink_t decsink = &decsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;

  LPSTREAM tmpstream = NULL;
  engine_filter_t filter = NULL;
  HWND curWindow;
  char* encData = NULL;
  char* senderAddr = NULL;
  char* subject = NULL;
  int encDataLen = 0;
  int rc = 0;
  unsigned int session_number;
  HRESULT hr;
  STATSTG tmpStat;

  protocol_t protocol;

  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;

  memset (decsink, 0, sizeof *decsink);
  memset (sink, 0, sizeof *sink);

  curWindow = get_oom_context_window (context);

  if (!(flags & DATA_BODY))
    {
      wordEditor = get_oom_object (context, "WordEditor");
      wordApplication = get_oom_object (wordEditor, "get_Application");
      selection = get_oom_object (wordApplication, "get_Selection");
    }
  mailItem = get_oom_object (context, "CurrentItem");

  if ((!wordEditor || !wordApplication || !selection || !mailItem) &&
      !(flags & DATA_BODY))
    {
      MessageBox (NULL,
                  "Internal error in GpgOL.\n"
                    "Could not find all objects.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      log_error ("%s:%s: Could not find all objects.",
                 SRCNAME, __func__);
      goto failure;
    }

  if (!mailItem)
    {
      /* This happens when we try to decrypt the body of a mail in the
         explorer context. */
      mailItem = get_oom_object (context, "Selection.Item(1)");

      if (!mailItem)
        {
          MessageBox (NULL,
                      _("Please select a Mail."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }

  if (flags & DATA_SELECTION)
    {
      encData = get_oom_string (selection, "Text");

      if (!encData || (encDataLen = strlen (encData)) <= 1)
        {
          MessageBox (NULL,
                      _("Please select the data you wish to decrypt."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }
  else if (flags & DATA_BODY)
    {
      encData = get_oom_string (mailItem, "Body");

      if (!encData || (encDataLen = strlen (encData)) <= 1)
        {
          MessageBox (NULL,
                      _("Nothing to decrypt."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }

  fix_linebreaks (encData, &encDataLen);

  subject = get_oom_string (mailItem, "Subject");
  if (get_oom_bool (mailItem, "Sent"))
    {
      char *addrType = get_oom_string (mailItem, "SenderEmailType");
      if (addrType && strcmp("SMTP", addrType) == 0)
        {
          senderAddr = get_oom_string (mailItem, "SenderEmailAddress");
        }
      else
        {
          /* Not SMTP, fall back to try getting the property. */
          LPDISPATCH sender = get_oom_object (mailItem, "Sender");
          senderAddr = get_pa_string (sender, PR_SMTP_ADDRESS);
          RELDISP (sender);
        }
      xfree (addrType);
    }
  else
    {
      /* If the message has not been sent we might be composing
         in this case use the current address */
      LPDISPATCH sender = get_oom_object (mailItem, "Session.CurrentUser");
      senderAddr = get_pa_string (sender, PR_SMTP_ADDRESS);
      RELDISP (sender);
    }

  /* Determine the protocol based on the content */
  protocol = is_cms_data (encData, encDataLen) ? PROTOCOL_SMIME :
    PROTOCOL_OPENPGP;

  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                         (SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE
                          | STGM_CREATE | STGM_READWRITE),
                         NULL, GpgOLStr("GPG"), &tmpstream);

  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create temp file: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      rc = -1;
      goto failure;
    }

  sink->cb_data = tmpstream;
  sink->writefnc = sink_std_write;

  session_number = engine_new_session_number ();
  if (engine_create_filter (&filter, write_buffer_for_cb, sink))
    goto failure;

  decsink->cb_data = filter;
  decsink->writefnc = sink_encryption_write;

  engine_set_session_number (filter, session_number);
  engine_set_session_title (filter, subject ? subject : _("GpgOL"));

  if (flags & OP_DECRYPT)
    {
      if ((rc=engine_decrypt_start (filter, curWindow,
                                    protocol,
                                    1, NULL)))
        {
          log_error ("%s:%s: engine decrypt start failed: %s",
                     SRCNAME, __func__, gpg_strerror (rc));
          goto failure;
        }
    }
  else if (flags & OP_VERIFY)
    {
      log_debug ("Starting verify");
      if ((rc=engine_verify_start (filter, curWindow,
                                   NULL, 0, protocol, senderAddr)))
        {
          log_error ("%s:%s: engine verify start failed: %s",
                     SRCNAME, __func__, gpg_strerror (rc));
          goto failure;
        }
    }

  /* Write the text in the decryption sink. */
  rc = write_buffer (decsink, encData, encDataLen);

  /* Flush the decryption sink and wait for the decryption to get
     ready.  */
  if ((rc = write_buffer (decsink, NULL, 0)))
    goto failure;
  if ((rc = engine_wait (filter)))
    goto failure;
  filter = NULL; /* Not valid anymore.  */
  decsink->cb_data = NULL; /* Not needed anymore.  */

  if (!sink->enc_counter)
    {
      log_debug ("%s:%s: nothing received from engine", SRCNAME, __func__);
      goto failure;
    }

  /* Check the size of the decrypted data */
  tmpstream->Stat (&tmpStat, 0);

  if (tmpStat.cbSize.QuadPart > UINT_MAX)
    {
      log_error ("%s:%s: No one should write so large mails.",
                 SRCNAME, __func__);
      goto failure;
    }

  /* Copy the decrypted stream to the message editor.  */
  {
    LARGE_INTEGER off;
    ULONG nread;
    char buffer[(unsigned int)tmpStat.cbSize.QuadPart + 1];

    memset (buffer, 0, sizeof buffer);

    off.QuadPart = 0;
    hr = tmpstream->Seek (off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    hr = tmpstream->Read (buffer, sizeof (buffer) - 1, &nread);
    if (hr)
      {
        log_error ("%s:%s: IStream::Read failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    if (strlen (buffer) > 1)
      {
        /* Now replace the crypto data with the decrypted data or show it
        somehow.*/
        int err = 0;
        if (flags & DATA_SELECTION)
          {
            err = put_oom_string (selection, "Text", buffer);
          }
        else if (flags & DATA_BODY)
          {
            err = put_oom_string (mailItem, "Body", buffer);
          }

        if (err)
          {
            MessageBox (NULL, buffer,
                        flags & OP_DECRYPT ? _("Plain text") :
                        _("Signed text"),
                        MB_ICONINFORMATION|MB_OK);
          }
      }
    else
      {
        /* Just to be save not to overwrite the selection with
           an empty buffer */
        log_error ("%s:%s: unexpected problem ", SRCNAME, __func__);
        goto failure;
      }
  }

 failure:
  if (rc)
    log_debug ("%s:%s: failed rc=%d (%s) <%s>", SRCNAME, __func__, rc,
               gpg_strerror (rc), gpg_strsource (rc));
  engine_cancel (filter);
  RELDISP (mailItem);
  RELDISP (selection);
  RELDISP (wordEditor);
  RELDISP (wordApplication);
  xfree (encData);
  xfree (senderAddr);
  xfree (subject);
  if (tmpstream)
    tmpstream->Release();

  return S_OK;
}


/* getIcon
   Loads a PNG image from the resurce converts it into a Bitmap
   and Wraps it in an PictureDispatcher that is returned as result.

   Based on documentation from:
   http://www.codeproject.com/Articles/3537/Loading-JPG-PNG-resources-using-GDI
*/

HRESULT
getIcon (int id, VARIANT* result)
{
  PICTDESC pdesc;
  LPDISPATCH pPict;
  HRESULT hr;
  Gdiplus::GdiplusStartupInput gdiplusStartupInput;
  Gdiplus::Bitmap* pbitmap;
  ULONG_PTR gdiplusToken;
  HRSRC hResource;
  DWORD imageSize;
  const void* pResourceData;
  HGLOBAL hBuffer;

  memset (&pdesc, 0, sizeof pdesc);
  pdesc.cbSizeofstruct = sizeof pdesc;
  pdesc.picType = PICTYPE_BITMAP;

  /* Initialize GDI */
  gdiplusStartupInput.DebugEventCallback = NULL;
  gdiplusStartupInput.SuppressBackgroundThread = FALSE;
  gdiplusStartupInput.SuppressExternalCodecs = FALSE;
  gdiplusStartupInput.GdiplusVersion = 1;
  GdiplusStartup (&gdiplusToken, &gdiplusStartupInput, NULL);

  /* Get the image from the resource file */
  hResource = FindResource (glob_hinst, MAKEINTRESOURCE(id), RT_RCDATA);
  if (!hResource)
    {
      log_error ("%s:%s: failed to find image: %i",
                 SRCNAME, __func__, id);
      return E_FAIL;
    }

  imageSize = SizeofResource (glob_hinst, hResource);
  if (!imageSize)
    return E_FAIL;

  pResourceData = LockResource (LoadResource(glob_hinst, hResource));

  if (!pResourceData)
    {
      log_error ("%s:%s: failed to load image: %i",
                 SRCNAME, __func__, id);
      return E_FAIL;
    }

  hBuffer = GlobalAlloc (GMEM_MOVEABLE, imageSize);

  if (hBuffer)
    {
      void* pBuffer = GlobalLock (hBuffer);
      if (pBuffer)
        {
          IStream* pStream = NULL;
          CopyMemory (pBuffer, pResourceData, imageSize);

          if (CreateStreamOnHGlobal (hBuffer, FALSE, &pStream) == S_OK)
            {
              pbitmap = Gdiplus::Bitmap::FromStream (pStream);
              pStream->Release();
              if (!pbitmap || pbitmap->GetHBITMAP (0, &pdesc.bmp.hbitmap))
                {
                  log_error ("%s:%s: failed to get PNG.",
                             SRCNAME, __func__);
                }
            }
        }
      GlobalUnlock (pBuffer);
    }
  GlobalFree (hBuffer);

  Gdiplus::GdiplusShutdown (gdiplusToken);

  /* Wrap the image into an OLE object.  */
  hr = OleCreatePictureIndirect (&pdesc, IID_IPictureDisp,
                                 TRUE, (void **) &pPict);
  if (hr != S_OK || !pPict)
    {
      log_error ("%s:%s: OleCreatePictureIndirect failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return -1;
    }

  result->pdispVal = pPict;
  result->vt = VT_DISPATCH;

  return S_OK;
}

/* Adds an encrypted attachment if the flag OP_SIGN is set
   a detached signature of the encrypted file is also added. */
static HRESULT
attachEncryptedFile (LPDISPATCH ctrl, int flags)
{
  LPDISPATCH context = NULL;
  LPDISPATCH mailItem = NULL;
  LPDISPATCH sender = NULL;
  LPDISPATCH recipients = NULL;
  HRESULT hr;
  char* senderAddr = NULL;
  char** recipientAddrs = NULL;
  char* subject = NULL;

  HWND curWindow;
  char *fileToEncrypt = NULL;
  wchar_t *fileToEncryptW = NULL;
  wchar_t *encryptedFile = NULL;
  wchar_t *attachName = NULL;
  HANDLE hFile = NULL;
  HANDLE hEncFile = NULL;

  unsigned int session_number;
  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  engine_filter_t filter = NULL;
  protocol_t protocol;
  int rc = 0;
  int i = 0;

  memset (encsink, 0, sizeof *encsink);
  memset (sink, 0, sizeof *sink);

  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;

  /* First do the check for recipients as this is likely
     to fail */
  mailItem = get_oom_object (context, "CurrentItem");
  sender = get_oom_object (mailItem, "Session.CurrentUser");
  recipients = get_oom_object (mailItem, "Recipients");
  recipientAddrs = get_oom_recipients (recipients);

  if (!recipientAddrs || !(*recipientAddrs))
    {
      MessageBox (NULL,
                  _("Please add at least one recipent."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      goto failure;
    }

  /* Get a file handle to read from */
  fileToEncrypt = get_open_filename (NULL, _("Select file to encrypt"));

  if (!fileToEncrypt)
    {
      log_debug ("No file selected");
      goto failure;
    }

  fileToEncryptW = utf8_to_wchar2 (fileToEncrypt, strlen(fileToEncrypt));
  xfree (fileToEncrypt);

  hFile = CreateFileW (fileToEncryptW,
                       GENERIC_READ,
                       FILE_SHARE_READ,
                       NULL,
                       OPEN_EXISTING,
                       FILE_ATTRIBUTE_NORMAL,
                       NULL);
  if (hFile == INVALID_HANDLE_VALUE)
    {
      /* Should not happen as the Open File dialog
         should have prevented this.
         Maybe this also happens when a file is
         not readable. In that case we might want
         to switch to a localized error naming the file. */
      MessageBox (NULL,
                  "Internal error in GpgOL.\n"
                  "Could not open File.",
                  _("GpgOL"),
                  MB_ICONERROR|MB_OK);
      return S_OK;
    }

  /* Now do the encryption preperations */

  if (!mailItem || !sender || !recipients)
    {
      MessageBox (NULL,
                  "Internal error in GpgOL.\n"
                  "Could not find all objects.",
                  _("GpgOL"),
                  MB_ICONERROR|MB_OK);
      log_error ("%s:%s: Could not find all objects.",
                 SRCNAME, __func__);
      goto failure;
    }

  senderAddr = get_pa_string (sender, PR_SMTP_ADDRESS);

  curWindow = get_oom_context_window (context);

  session_number = engine_new_session_number ();

  subject = get_oom_string (mailItem, "Subject");

  /* Prepare the encryption sink */
  if ((rc = engine_create_filter (&filter, write_buffer_for_cb, sink)))
    {
      goto failure;
    }

  encsink->cb_data = filter;
  encsink->writefnc = sink_encryption_write;

  engine_set_session_number (filter, session_number);
  engine_set_session_title (filter, subject ? subject :_("GpgOL"));
  if ((rc=engine_encrypt_prepare (filter, curWindow,
                                  PROTOCOL_UNKNOWN,
                                  ENGINE_FLAG_BINARY_OUTPUT,
                                  senderAddr, recipientAddrs, &protocol)))
    {
      log_error ("%s:%s: engine encrypt prepare failed : %s",
                 SRCNAME, __func__, gpg_strerror (rc));
      goto failure;
    }

  attachName = get_pretty_attachment_name (fileToEncryptW, protocol, 0);

  if (!attachName)
    {
      log_error ("%s:%s: Could not get a decent attachment name",
                 SRCNAME, __func__);
      goto failure;
    }

  encryptedFile = get_tmp_outfile (attachName, &hEncFile);
  sink->cb_data = hEncFile;
  sink->writefnc = sink_file_write;

  if ((rc=engine_encrypt_start (filter, 0)))
    {
      log_error ("%s:%s: engine encrypt start failed: %s",
                 SRCNAME, __func__, gpg_strerror (rc));
      goto failure;
    }

  if ((rc=copyFileToSink (hFile, encsink)))
    goto failure;

  /* Lets hope the user did not select a huge file. We are hanging
   here until encryption is completed.. */
  if ((rc = engine_wait (filter)))
    goto failure;

  filter = NULL; /* Not valid anymore.  */
  encsink->cb_data = NULL; /* Not needed anymore.  */

  if (!sink->enc_counter)
    {
      log_error ("%s:%s: nothing received from engine", SRCNAME, __func__);
      goto failure;
    }

  /* Now we have an encrypted file behind encryptedFile. Let's add it */
  add_oom_attachment (mailItem, encryptedFile);

  if (flags & OP_SIGN)
    {
      attachSignature (mailItem, subject, hEncFile, protocol, session_number,
                       curWindow, encryptedFile, senderAddr);
    }

failure:
  if (filter)
    engine_cancel (filter);

  if (hEncFile)
    {
      CloseHandle (hEncFile);
      DeleteFileW (encryptedFile);
    }
  xfree (senderAddr);
  xfree (encryptedFile);
  xfree (fileToEncryptW);
  xfree (attachName);
  xfree (subject);
  RELDISP (mailItem);
  RELDISP (sender);
  RELDISP (recipients);

  if (hFile)
    CloseHandle (hFile);
  if (recipientAddrs)
    {
      for (i=0; recipientAddrs && recipientAddrs[i]; i++)
        xfree (recipientAddrs[i]);
      xfree (recipientAddrs);
    }

  return S_OK;
}

HRESULT
startCertManager (LPDISPATCH ctrl)
{
  HRESULT hr;
  LPDISPATCH context;
  HWND curWindow;

  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;

  curWindow = get_oom_context_window (context);

  engine_start_keymanager (curWindow);
  return S_OK;
}

HRESULT
decryptBody (LPDISPATCH ctrl)
{
  return do_reader_action (ctrl, OP_DECRYPT | DATA_BODY);
}

HRESULT
decryptSelection (LPDISPATCH ctrl)
{
  return do_reader_action (ctrl, OP_DECRYPT | DATA_SELECTION);
}

HRESULT
encryptBody (LPDISPATCH ctrl)
{
  return do_composer_action (ctrl, OP_ENCRYPT | DATA_BODY);
}

HRESULT
encryptSelection (LPDISPATCH ctrl)
{
  return do_composer_action (ctrl, OP_ENCRYPT | DATA_SELECTION);
}

HRESULT
addEncSignedAttachment (LPDISPATCH ctrl)
{
  return attachEncryptedFile (ctrl, OP_SIGN);
}

HRESULT
addEncAttachment (LPDISPATCH ctrl)
{
  return attachEncryptedFile (ctrl, 0);
}

HRESULT signBody (LPDISPATCH ctrl)
{
  return do_composer_action (ctrl, DATA_BODY | OP_SIGN);
}

HRESULT verifyBody (LPDISPATCH ctrl)
{
  return do_reader_action (ctrl, DATA_BODY | OP_VERIFY);
}

static void
message_flag_status (HWND window, int flags)
{
  const char * message;
  if (flags & OP_ENCRYPT && flags & OP_SIGN)
    {
      message = _("The message will be signed & encrypted.");
    }
  else if (flags & OP_ENCRYPT)
    {
      message = _("The message will be encrypted.");
    }
  else if (flags & OP_SIGN)
    {
      message = _("The message will be signed.");
    }
  else
    {
      message = _("The message will be sent plain and without a signature.");
    }
  MessageBox (NULL,
              message,
              _("GpgOL"),
              MB_ICONINFORMATION|MB_OK);
}

static HRESULT
mark_mime_action (LPDISPATCH ctrl, int flags)
{
  HRESULT hr;
  HRESULT rc = E_FAIL;
  HWND cur_window;
  LPDISPATCH context = NULL,
             mailitem = NULL;
  LPMESSAGE message = NULL;
  int oldflags,
      newflags;

  log_debug ("%s:%s: enter", SRCNAME, __func__);
  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;
  cur_window = get_oom_context_window (context);

  mailitem = get_oom_object (context, "CurrentItem");

  if (!mailitem)
    {
      log_error ("%s:%s: Failed to get mailitem.",
                 SRCNAME, __func__);
      goto done;
    }

  message = get_oom_message (mailitem);

  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
      goto done;
    }

  oldflags = get_gpgol_draft_info_flags (message);

  newflags = oldflags xor flags;

  if (set_gpgol_draft_info_flags (message, newflags))
    {
      log_error ("%s:%s: Failed to set draft flags.",
                 SRCNAME, __func__);
    }

  message_flag_status (cur_window, newflags);

  rc = S_OK;

done:
  RELDISP (context);
  RELDISP (mailitem);
  RELDISP (message);

  return rc;
}

HRESULT mime_sign (LPDISPATCH ctrl)
{
  return mark_mime_action (ctrl, OP_SIGN);
}

HRESULT mime_encrypt (LPDISPATCH ctrl)
{
  return mark_mime_action (ctrl, OP_ENCRYPT);
}