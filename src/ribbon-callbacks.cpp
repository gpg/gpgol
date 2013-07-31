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

/* Gets the context of a ribbon control. And prints some
   useful debug output */
HRESULT getContext (LPDISPATCH ctrl, LPDISPATCH *context)
{
  *context = get_oom_object (ctrl, "get_Context");
  log_debug ("%s:%s: contextObj: %s",
             SRCNAME, __func__, get_object_name (*context));
  return context ? S_OK : E_FAIL;
}

#define ENCRYPT_INSPECTOR_SELECTION  1
#define ENCRYPT_INSPECTOR_BODY       2

/* encryptInspector
   Encrypts text in an IInspector context. Depending on
   the flags either the active selection or the full body
   is encrypted.
*/

HRESULT
encryptInspector (LPDISPATCH ctrl, int flags)
{
  LPDISPATCH context = NULL;
  LPDISPATCH selection;
  LPDISPATCH wordEditor;
  LPDISPATCH application;
  LPDISPATCH mailItem;
  LPDISPATCH sender;
  LPDISPATCH recipients;

  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  char* senderAddr = NULL;
  LPSTREAM tmpstream = NULL;
  engine_filter_t filter = NULL;
  char* plaintext = NULL;
  int rc = 0;
  HRESULT hr;
  int recipientsCnt;
  HWND curWindow;
  protocol_t protocol;
  unsigned int session_number;
  int i;
  STATSTG tmpStat;

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

  if (flags & ENCRYPT_INSPECTOR_SELECTION)
    {
      plaintext = get_oom_string (selection, "Text");

      if (!plaintext || strlen (plaintext) <= 1)
        {
          /* TODO more usable if we just use all text in this case? */
          MessageBox (NULL,
                      _("Please select text to encrypt."),
                      _("GpgOL"),
                      MB_ICONINFORMATION|MB_OK);
          goto failure;
        }
    }
  else if (flags & ENCRYPT_INSPECTOR_BODY)
    {
      plaintext = get_oom_string (mailItem, "Body");
      if (!plaintext || strlen (plaintext) <= 1)
        {
          /* TODO more usable if we just use all text in this case? */
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
                         NULL, "GPG", &tmpstream);

  if (FAILED (hr))
    {
      log_error ("%s:%s: can't create temp file: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      rc = -1;
      goto failure;
    }

  sink->cb_data = tmpstream;
  sink->writefnc = sink_std_write;

  senderAddr = get_oom_string (sender, "Address");

  recipientsCnt = get_oom_int (recipients, "Count");

  if (!recipientsCnt)
    {
      MessageBox (NULL,
                  _("Please add at least one recipent."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      goto failure;
    }

  {
    /* Get the recipients */
    char *recipientAddrs[recipientsCnt + 1];
    recipientAddrs[recipientsCnt] = NULL;
    for (i = 1; i <= recipientsCnt; i++)
      {
        char buf[16];
        LPDISPATCH recipient;
        snprintf (buf, sizeof (buf), "Item(%i)", i);
        recipient = get_oom_object (recipients, buf);
        if (!recipient)
          {
            /* Should be impossible */
            recipientAddrs[i-1] = NULL;
            log_error ("%s:%s: could not find Item %i;",
                       SRCNAME, __func__, i);
            break;
          }
        recipientAddrs[i-1] = get_oom_string (recipient, "Address");
      }

    /* Not lets prepare our encryption */
    session_number = engine_new_session_number ();

    /* Prepare the encryption sink */

    if (engine_create_filter (&filter, write_buffer_for_cb, sink))
      {
        for (i = 0; i < recipientsCnt; i++)
          xfree (recipientAddrs[i]);
        goto failure;
      }

    encsink->cb_data = filter;
    encsink->writefnc = sink_encryption_write;

    engine_set_session_number (filter, session_number);
    engine_set_session_title (filter, _("GpgOL"));

    if ((rc=engine_encrypt_prepare (filter, curWindow,
                                    PROTOCOL_UNKNOWN,
                                    0 /* ENGINE_FLAG_SIGN_FOLLOWS */,
                                    senderAddr, recipientAddrs, &protocol)))
      {
        for (i = 0; i < recipientsCnt; i++)
          xfree (recipientAddrs[i]);
        log_error ("%s:%s: engine encrypt prepare failed : %s",
                   SRCNAME, __func__, gpg_strerror (rc));
        goto failure;
      }
    for (i = 0; i < recipientsCnt; i++)
      xfree (recipientAddrs[i]);

    /* lets go */

    if ((rc=engine_encrypt_start (filter, 0)))
      {
        log_error ("%s:%s: engine encrypt start failed: %s",
                   SRCNAME, __func__, gpg_strerror (rc));
        goto failure;
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
    char buffer[(unsigned int)tmpStat.cbSize.QuadPart];

    off.QuadPart = 0;
    hr = tmpstream->Seek (off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    hr = tmpstream->Read (buffer, sizeof buffer, &nread);
    if (hr)
      {
        log_error ("%s:%s: IStream::Read failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    if (strlen (buffer) > 1)
      {
        char* lastlinebreak = strrchr (buffer, '\n');
        if (lastlinebreak && (lastlinebreak - buffer) > 1)
          {
            /*XXX there is some strange data in the buffer
              after the last linebreak investigate this and
              fix it! */
            lastlinebreak[1] = '\0';
          }
        /* Now replace the selection with the encrypted text */
        if (protocol == PROTOCOL_SMIME)
          {
            unsigned int enclosedSize = strlen (buffer) + 34 + 31 + 1;
            char enclosedData[enclosedSize];
            snprintf (enclosedData, sizeof enclosedData,
                      "-----BEGIN ENCRYPTED MESSAGE-----\r\n"
                      "%s"
                      "-----END ENCRYPTED MESSAGE-----\r\n", buffer);
            if (flags & ENCRYPT_INSPECTOR_SELECTION)
              put_oom_string (selection, "Text", enclosedData);
            else if (flags & ENCRYPT_INSPECTOR_BODY)
              put_oom_string (mailItem, "Body", enclosedData);

          }
        else
          {
            if (flags & ENCRYPT_INSPECTOR_SELECTION)
              put_oom_string (selection, "Text", buffer);
            else if (flags & ENCRYPT_INSPECTOR_BODY)
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

#define DECRYPT_INSPECTOR_SELECTION  1
#define DECRYPT_INSPECTOR_BODY       2

/* decryptInspector
   decrypts the content of an inspector. Controled by flags
   similary to the encryptInspector.
*/

HRESULT
decryptInspector (LPDISPATCH ctrl, int flags)
{
  LPDISPATCH context;
  LPDISPATCH selection;
  LPDISPATCH wordEditor;
  LPDISPATCH mailItem;
  LPDISPATCH wordApplication;

  struct sink_s decsinkmem;
  sink_t decsink = &decsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;

  LPSTREAM tmpstream = NULL;
  engine_filter_t filter = NULL;
  HWND curWindow;
  char* encData = NULL;
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

  if (!(flags & DECRYPT_INSPECTOR_BODY))
    {
      wordEditor = get_oom_object (context, "WordEditor");
      wordApplication = get_oom_object (wordEditor, "get_Application");
      selection = get_oom_object (wordApplication, "get_Selection");
    }
  mailItem = get_oom_object (context, "CurrentItem");

  if ((!wordEditor || !wordApplication || !selection || !mailItem) &&
      !(flags & DECRYPT_INSPECTOR_BODY))
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

  if (flags & DECRYPT_INSPECTOR_SELECTION)
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
  else if (flags & DECRYPT_INSPECTOR_BODY)
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

  /* Determine the protocol based on the content */
  protocol = is_cms_data (encData, encDataLen) ? PROTOCOL_SMIME :
    PROTOCOL_OPENPGP;

  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                         (SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE
                          | STGM_CREATE | STGM_READWRITE),
                         NULL, "GPG", &tmpstream);

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
  engine_set_session_title (filter, _("GpgOL"));

  if ((rc=engine_decrypt_start (filter, curWindow,
                                protocol,
                                1, NULL)))
    {
      log_error ("%s:%s: engine decrypt start failed: %s",
                 SRCNAME, __func__, gpg_strerror (rc));
      goto failure;
    }

  /* Write the text in the decryption sink. */
  rc = write_buffer (decsink, encData, encDataLen);

  /* Flush the decryption sink and wait for the encryption to get
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

  /* Copy the encrypted stream to the message editor.  */
  {
    LARGE_INTEGER off;
    ULONG nread;
    char buffer[(unsigned int)tmpStat.cbSize.QuadPart];

    off.QuadPart = 0;
    hr = tmpstream->Seek (off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    hr = tmpstream->Read (buffer, sizeof buffer, &nread);
    if (hr)
      {
        log_error ("%s:%s: IStream::Read failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }
    if (strlen (buffer) > 1)
      {
        /* Now replace the crypto data with the encData or show it
        somehow.*/
        int err;
        if (flags & DECRYPT_INSPECTOR_SELECTION)
          {
            err = put_oom_string (selection, "Text", buffer);
          }
        else if (flags & DECRYPT_INSPECTOR_BODY)
          {
            err = put_oom_string (mailItem, "Body", buffer);
          }

        if (err)
          {
            MessageBox (NULL, buffer,
                        _("Plain text"),
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
}

HRESULT
decryptBody (LPDISPATCH ctrl)
{
  return decryptInspector (ctrl, DECRYPT_INSPECTOR_BODY);
}

HRESULT
decryptSelection (LPDISPATCH ctrl)
{
  return decryptInspector (ctrl, DECRYPT_INSPECTOR_SELECTION);
}

HRESULT
encryptBody (LPDISPATCH ctrl)
{
  return encryptInspector (ctrl, ENCRYPT_INSPECTOR_BODY);
}

HRESULT
encryptSelection (LPDISPATCH ctrl)
{
  return encryptInspector (ctrl, ENCRYPT_INSPECTOR_SELECTION);
}

HRESULT
addEncSignedAttachment (LPDISPATCH ctrl)
{
  /* TODO */
  return S_OK;
}
