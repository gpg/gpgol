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


HRESULT
encryptSelection (LPDISPATCH ctrl)
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
  char* text = NULL;
  int rc = 0;
  HRESULT hr;
  int recipientsCnt;
  HWND curWindow;
  LPOLEWINDOW actExplorer;
  protocol_t protocol;
  unsigned int session_number;
  int i;
  STATSTG tmpStat;

  hr = getContext (ctrl, &context);
  if (FAILED(hr))
      return hr;

  memset (encsink, 0, sizeof *encsink);
  memset (sink, 0, sizeof *sink);

  actExplorer = (LPOLEWINDOW) get_oom_object(context,
                                             "Application.ActiveExplorer");
  if (actExplorer)
    actExplorer->GetWindow (&curWindow);
  else
    {
      log_debug ("%s:%s: Could not find active window",
                 SRCNAME, __func__);
      curWindow = NULL;
    }

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
                  _("Internal error in GpgOL.\n"
                    "Could not find all objects."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      log_error ("%s:%s: Could not find all objects.",
                 SRCNAME, __func__);
      return S_OK;
    }

  text = get_oom_string (selection, "Text");

  if (!text || strlen (text) <= 1)
    {
      /* TODO more usable if we just use all text in this case? */
      MessageBox (NULL,
                  _("Please select some text for encryption."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      return S_OK;
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
                  _("Please enter the recipent of the encrypted text first."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      return S_OK;
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
      {
        char *tmp = get_oom_string (mailItem, "Subject");
        engine_set_session_title (filter, tmp);
        xfree (tmp);
      }

    if ((rc=engine_encrypt_prepare (filter, curWindow,
                                    protocol,
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
    rc = write_buffer (encsink, text, strlen (text));

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
      MessageBox (curWindow, _("GpgOL"),
                  _("Selected text too long."),
                  MB_ICONINFORMATION|MB_OK);
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
        /* Now replace the selection with the encrypted text */
        put_oom_string (selection, "Text", buffer);
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
  if (tmpstream)
    tmpstream->Release();
  xfree (text);
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
  LPOLEWINDOW actExplorer;
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

  actExplorer = (LPOLEWINDOW) get_oom_object(attachmentSelection,
                                             "Application.ActiveExplorer");
  if (actExplorer)
    actExplorer->GetWindow (&curWindow);
  else
    {
      log_debug ("%s:%s: Could not find active window",
                 SRCNAME, __func__);
      curWindow = NULL;
    }
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

        wcsOutFilename = utf8_to_wchar2 (filenames[i-1],
                                         strlen(filenames[i-1]));
        saveParams.rgvarg[0].bstrVal = SysAllocString (wcsOutFilename);
        saveParams.cArgs = 1;
        saveParams.cNamedArgs = 0;

        hr = attachmentObj->Invoke (saveID, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                    DISPATCH_METHOD, &saveParams,
                                    NULL, NULL, NULL);
        SysFreeString (saveParams.rgvarg[0].bstrVal);
        if (FAILED(hr))
          {
            int j;
            log_debug ("%s:%s: Saving to file failed. hr: %x",
                       SRCNAME, __func__, (unsigned int) hr);
            for (j = 0; j < i; j++)
              xfree (filenames[j]);
            return hr;
          }
      }
    err = op_assuan_start_decrypt_files (curWindow, filenames);
    for (i = 0; i < attachmentCount; i++)
      xfree (filenames[i]);
  }

  log_debug ("%s:%s: Leaving. Err: %i",
             SRCNAME, __func__, err);

  return S_OK; /* If we return an error outlook will show that our
                  callback function failed in an ugly window. */
}

HRESULT
decryptSelection (LPDISPATCH ctrl)
{
  LPDISPATCH context;
  LPDISPATCH selection;
  LPDISPATCH wordEditor;
  LPDISPATCH wordApplication;

  struct sink_s decsinkmem;
  sink_t decsink = &decsinkmem;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;

  LPSTREAM tmpstream = NULL;
  engine_filter_t filter = NULL;
  LPOLEWINDOW actExplorer;
  HWND curWindow;
  char* text = NULL;
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

  actExplorer = (LPOLEWINDOW) get_oom_object(context,
                                             "Application.ActiveExplorer");
  if (actExplorer)
    actExplorer->GetWindow (&curWindow);
  else
    {
      log_debug ("%s:%s: Could not find active window",
                 SRCNAME, __func__);
      curWindow = NULL;
    }

  wordEditor = get_oom_object (context, "WordEditor");
  wordApplication = get_oom_object (wordEditor, "get_Application");
  selection = get_oom_object (wordApplication, "get_Selection");

  if (!wordEditor || !wordApplication || !selection)
    {
      MessageBox (NULL,
                  _("Internal error in GpgOL.\n"
                    "Could not find all objects."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      log_error ("%s:%s: Could not find all objects.",
                 SRCNAME, __func__);
      return S_OK;
    }

  text = get_oom_string (selection, "Text");

  if (!text || strlen (text) <= 1)
    {
      /* TODO more usable if we just use all text in this case? */
      MessageBox (NULL,
                  _("Please select the data you wish to decrypt."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      return S_OK;
    }

  /* Determine the protocol based on the content */
  protocol = is_cms_data (text, strlen (text)) ? PROTOCOL_SMIME :
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
  engine_set_session_title (filter, _("Decrypt Selection"));

  if ((rc=engine_decrypt_start (filter, curWindow,
                                protocol,
                                1, NULL)))
    {
      log_error ("%s:%s: engine decrypt start failed: %s",
                 SRCNAME, __func__, gpg_strerror (rc));
      goto failure;
    }

  /* Write the text in the decryption sink. */
  rc = write_buffer (decsink, text, strlen (text));

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
      MessageBox (curWindow, _("GpgOL"),
                  _("Selected text too long."),
                  MB_ICONINFORMATION|MB_OK);
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
        /* Now replace the selection with the encrypted text */
        put_oom_string (selection, "Text", buffer);
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
  xfree (text);
  if (tmpstream)
    tmpstream->Release();

  return S_OK;
}

HRESULT
getIcon (int id, int size, VARIANT* result)
{
  PICTDESC pdesc;
  LPDISPATCH pPict;
  HRESULT hr;
  UINT fuload;

  memset (&pdesc, 0, sizeof pdesc);
  pdesc.cbSizeofstruct = sizeof pdesc;
  pdesc.picType = PICTYPE_BITMAP;

/*
   In the future we might want to use PNGs here to have
   full Alpha Channel support for the icons

   Some explanation about images and transparency in Ribbons:
   http://blogs.msdn.com/b/jensenh/archive/2006/11/27/ribbonx-image-faq.aspx

   Here is an example how this could look like with gdiplus:

   GdiplusStartupInput gdiplusStartupInput;
   ULONG_PTR gdiplusToken;
   Bitmap* pbitmap;

   GetModuleFileName(glob_hinst, szModuleFileName, MAX_PATH);

   gdiplusStartupInput.DebugEventCallback = NULL;
   gdiplusStartupInput.SuppressBackgroundThread = FALSE;
   gdiplusStartupInput.SuppressExternalCodecs = FALSE;
   gdiplusStartupInput.GdiplusVersion = 1;
   GdiplusStartup (&gdiplusToken, &gdiplusStartupInput, NULL);

   pbitmap = Bitmap::FromFile (L"c:\\foo.png", FALSE);
   if (!pbitmap || pbitmap->GetHBITMAP (0, &pdesc.bmp.hbitmap))
     {
       log_error ("%s:%s: failed to load file.",
                  SRCNAME, __func__);
     }
*/

  fuload = LR_CREATEDIBSECTION | LR_SHARED;

  pdesc.bmp.hbitmap = (HBITMAP) LoadImage (glob_hinst,
                                           MAKEINTRESOURCE (id),
                                           IMAGE_BITMAP, size, size, fuload);

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

  /*
  GdiplusShutdown (gdiplusToken);
  */

  return S_OK;
}
