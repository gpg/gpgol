/* ribbon-callbacks.h - Callbacks for the ribbon extension interface
 * Copyright (C) 2013 Intevation GmbH
 * Software engineering by Intevation GmbH
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
#include "common.h"

#include "mymapi.h"
#include "mymapitags.h"

#include "common.h"
#include "mapihelp.h"
#include "mimemaker.h"
#include "filetype.h"
#include "mail.h"

#include <gpgme++/context.h>
#include <gpgme++/data.h>

using namespace GpgME;

/* Gets the context of a ribbon control. And prints some
   useful debug output */
HRESULT getContext (LPDISPATCH ctrl, LPDISPATCH *context)
{
  *context = get_oom_object (ctrl, "get_Context");
  if (*context)
    {
      char *name = get_object_name (*context);
      log_debug ("%s:%s: contextObj: %s",
                 SRCNAME, __func__, name);
      xfree (name);
    }

  return context ? S_OK : E_FAIL;
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

  if (!result)
    {
      log_error ("getIcon called without result variant.");
      return E_POINTER;
    }

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
              gpgol_release (pStream);
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
mark_mime_action (LPDISPATCH ctrl, int flags, bool is_explorer)
{
  LPMESSAGE message = NULL;
  int oldflags,
      newflags;

  log_debug ("%s:%s: enter", SRCNAME, __func__);
  LPDISPATCH context = NULL;
  if (FAILED(getContext (ctrl, &context)))
    {
      TRACEPOINT;
      return E_FAIL;
    }

  LPDISPATCH mailitem = get_oom_object (context,
                                        is_explorer ? "ActiveInlineResponse" :
                                                      "CurrentItem");
  gpgol_release (context);

  if (!mailitem)
    {
      log_error ("%s:%s: Failed to get mailitem.",
                 SRCNAME, __func__);
      return E_FAIL;
    }

  /* Get the uid of this item. */
  char *uid = get_unique_id (mailitem, 0, nullptr);
  if (!uid)
    {
      LPMESSAGE msg = get_oom_base_message (mailitem);
      uid = mapi_get_uid (msg);
      gpgol_release (msg);
      if (!uid)
        {
          log_debug ("%s:%s: Failed to get uid for %p",
                   SRCNAME, __func__, mailitem);
        }
    }
  Mail *mail = nullptr;
  if (uid)
    {
      mail = Mail::get_mail_for_uuid (uid);
      xfree (uid);
    }

  if (mail)
    {
      mail->set_crypto_selected_manually (true);
    }
  else
    {
      log_debug ("%s:%s: Failed to get mail object.",
                 SRCNAME, __func__);
    }

  message = get_oom_base_message (mailitem);
  gpgol_release (mailitem);

  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
      return S_OK;
    }

  oldflags = get_gpgol_draft_info_flags (message);
  if (flags == 3 && oldflags != 3)
    {
      // If only one sub button is active activate
      // both now.
      newflags = 3;
    }
  else
    {
      newflags = oldflags xor flags;
    }

  if (set_gpgol_draft_info_flags (message, newflags))
    {
      log_error ("%s:%s: Failed to set draft flags.",
                 SRCNAME, __func__);
    }
  gpgol_release (message);

  /*  We need to invalidate the UI to update the toggle
      states of the subbuttons and the top button. Yeah,
      we invalidate a lot *sigh* */
  gpgoladdin_invalidate_ui ();

  if (newflags & 1)
    {
      Mail::locate_all_crypto_recipients ();
    }

  return S_OK;
}

/* Get the state of encrypt / sign toggle buttons.
  flag values: 1 get the state of the encrypt button.
               2 get the state of the sign button.
  If is_explorer is set to true we look at the inline response.
*/
HRESULT get_crypt_pressed (LPDISPATCH ctrl, int flags, VARIANT *result,
                           bool is_explorer)
{
  HRESULT hr;
  bool value;
  LPDISPATCH context = NULL,
             mailitem = NULL;
  LPMESSAGE message = NULL;

  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = (VARIANT_BOOL*) xmalloc (sizeof (VARIANT_BOOL));
  *(result->pboolVal) = VARIANT_FALSE;

  /* First the usual defensive check about our parameters */
  if (!ctrl || !result)
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      return E_FAIL;
    }

  hr = getContext (ctrl, &context);

  if (hr)
    {
      log_error ("%s:%s:%i : hresult %lx", SRCNAME, __func__, __LINE__,
                 hr);
      return E_FAIL;
    }

  mailitem = get_oom_object (context, is_explorer ? "ActiveInlineResponse" :
                                                    "CurrentItem");

  if (!mailitem)
    {
      log_error ("%s:%s: Failed to get mailitem.",
                 SRCNAME, __func__);
      goto done;
    }

  message = get_oom_base_message (mailitem);

  if (!message)
    {
      log_error ("%s:%s: No message found.",
                 SRCNAME, __func__);
      goto done;
    }

  value = (get_gpgol_draft_info_flags (message) & flags) == flags;

  *(result->pboolVal) = value ? VARIANT_TRUE : VARIANT_FALSE;

done:
  gpgol_release (context);
  gpgol_release (mailitem);
  gpgol_release (message);

  return S_OK;
}

static Mail *
get_mail_from_control (LPDISPATCH ctrl, bool *none_selected)
{
  HRESULT hr;
  LPDISPATCH context = NULL,
             mailitem = NULL;
  *none_selected = false;
  if (!ctrl)
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      return NULL;
    }
  hr = getContext (ctrl, &context);

  if (hr)
    {
      log_error ("%s:%s:%i : hresult %lx", SRCNAME, __func__, __LINE__,
                 hr);
      return NULL;
    }

  const auto ctx_name = std::string (get_object_name (context));

  if (ctx_name.empty())
    {
      log_error ("%s:%s: Failed to get context name",
                 SRCNAME, __func__);
      gpgol_release (context);
      return NULL;
    }

  if (!strcmp (ctx_name.c_str(), "_Inspector"))
    {
      mailitem = get_oom_object (context, "CurrentItem");
    }
  else if (!strcmp (ctx_name.c_str(), "_Explorer"))
    {
      if (g_ol_version_major >= 16)
        {
          // Avoid showing wrong crypto state if we don't have a reading
          // pane. In that case the parser will finish for a mail which is gone
          // and the crypto state will not get updated.
          //
          //
          // Somehow latest Outlook 2016 crashes when accessing the current view
          // of the Explorer. This is even reproducible with
          // GpgOL disabled and only with Outlook Spy active. If you select
          // the explorer of an Outlook.com resource and then access
          // the CurrentView and close the CurrentView again in Outlook Spy
          // outlook crashes.
          LPDISPATCH prevEdit = get_oom_object (context, "PreviewPane.WordEditor");
          gpgol_release (prevEdit);
          if (!prevEdit)
            {
              *none_selected = true;
              gpgol_release (mailitem);
              mailitem = nullptr;
            }
        }
      else
        {
          // Preview Pane is not available in older outlooks
          LPDISPATCH tableView = get_oom_object (context, "CurrentView");
          if (!tableView)
            {
              // Woops, should not happen.
              TRACEPOINT;
              *none_selected = true;
              gpgol_release (mailitem);
              mailitem = nullptr;
            }
          else
            {
              int hasReadingPane = get_oom_bool (tableView, "ShowReadingPane");
              gpgol_release (tableView);
              if (!hasReadingPane)
                {
                  *none_selected = true;
                  gpgol_release (mailitem);
                  mailitem = nullptr;
                }
            }
        }
      if (!*none_selected)
        {
          /* Accessing the selection item can trigger a load event
             so we only do this here if we think that there might be
             something visible / selected. To avoid triggering a load
             if there is no content shown. */
          LPDISPATCH selection = get_oom_object (context, "Selection");
          if (!selection)
            {
              log_error ("%s:%s: Failed to get selection.",
                         SRCNAME, __func__);
              gpgol_release (context);
              return NULL;
            }
          int count = get_oom_int (selection, "Count");
          if (count == 1)
            {
              // If we call this on a selection with more items
              // Outlook sends an ItemLoad event for each mail
              // in that selection.
              mailitem = get_oom_object (selection, "Item(1)");
            }
          gpgol_release (selection);

          if (!mailitem)
            {
              *none_selected = true;
            }
        }
    }
  else if (!strcmp (ctx_name.c_str(), "Selection"))
    {
      int count = get_oom_int (context, "Count");
      if (count == 1)
        {
          // If we call this on a selection with more items
          // Outlook sends an ItemLoad event for each mail
          // in that selection.
          mailitem = get_oom_object (context, "Item(1)");
        }
      if (!mailitem)
        {
          *none_selected = true;
        }
    }

  gpgol_release (context);
  if (!mailitem)
    {
      log_debug ("%s:%s: No mailitem. From %s",
                 SRCNAME, __func__, ctx_name.c_str());
      return NULL;
    }

  char *uid;
  /* Get the uid of this item. */
  uid = get_unique_id (mailitem, 0, nullptr);
  if (!uid)
    {
      LPMESSAGE msg = get_oom_base_message (mailitem);
      uid = mapi_get_uid (msg);
      gpgol_release (msg);
      if (!uid)
        {
          log_debug ("%s:%s: Failed to get uid for %p",
                   SRCNAME, __func__, mailitem);
          gpgol_release (mailitem);
          return NULL;
        }
    }

  auto ret = Mail::get_mail_for_uuid (uid);
  xfree (uid);
  if (!ret)
    {
      log_error ("%s:%s: Failed to find mail %p in map.",
                 SRCNAME, __func__, mailitem);
    }
  gpgol_release (mailitem);
  return ret;
}

/* Helper to reduce code duplication.*/
#define MY_MAIL_GETTER \
  if (!ctrl) \
    { \
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__); \
      return E_FAIL; \
    } \
  bool none_selected; \
  const auto mail = get_mail_from_control (ctrl, &none_selected); \
  (void)none_selected; \
  if (!mail) \
    { \
      log_oom ("%s:%s:%i Failed to get mail", \
               SRCNAME, __func__, __LINE__); \
    }

HRESULT get_is_details_enabled (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  if (!result)
    {
      TRACEPOINT;
      return S_OK;
    }

  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = (VARIANT_BOOL*) xmalloc (sizeof (VARIANT_BOOL));
  *(result->pboolVal) = none_selected ? VARIANT_FALSE : VARIANT_TRUE;

  return S_OK;
}

HRESULT get_sig_label (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BSTR;
  wchar_t *w_result;
  if (!mail)
    {
      log_debug ("%s:%s: No mail.",
                 SRCNAME, __func__);
      w_result = utf8_to_wchar (_("Insecure"));
      result->bstrVal = SysAllocString (w_result);
      xfree (w_result);
      return S_OK;
    }
  w_result = utf8_to_wchar (mail->get_crypto_summary ().c_str ());
  result->bstrVal = SysAllocString (w_result);
  xfree (w_result);
  return S_OK;
}

HRESULT get_sig_ttip (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BSTR;
  wchar_t *w_result;
  if (mail)
    {
      w_result = utf8_to_wchar (mail->get_crypto_one_line().c_str());
    }
  else if (!none_selected)
    {
      w_result = utf8_to_wchar (_("Insecure message"));
    }
  else
    {
      w_result = utf8_to_wchar (_("No message selected"));
    }
  result->bstrVal = SysAllocString (w_result);
  xfree (w_result);
  return S_OK;
}

HRESULT get_sig_stip (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BSTR;
  if (none_selected)
    {
      result->bstrVal = SysAllocString (L"");
      return S_OK;
    }
  if (!mail || !mail->is_crypto_mail ())
    {
      wchar_t *w_result;
      w_result = utf8_to_wchar (utf8_gettext ("You cannot be sure who sent, "
                                  "modified and read the message in transit.\n\n"
                                  "Click here to learn more."));
      result->bstrVal = SysAllocString (w_result);
      xfree (w_result);
      return S_OK;
    }
  const auto message = mail->get_crypto_details ();
  wchar_t *w_message = utf8_to_wchar (message.c_str());
  result->bstrVal = SysAllocString (w_message);
  xfree (w_message);
  return S_OK;
}

HRESULT launch_cert_details (LPDISPATCH ctrl)
{
  MY_MAIL_GETTER

  if (!mail || (!mail->is_signed () && !mail->is_encrypted ()))
    {
      ShellExecuteA(NULL, NULL, "https://emailselfdefense.fsf.org/infographic",
                    0, 0, SW_SHOWNORMAL);
      return S_OK;
    }

  if (!mail->is_signed () && mail->is_encrypted ())
    {
      /* Encrypt only, no information but show something. because
         we want the button to be active.

         Aheinecke: I don't think we should show to which keys the message
         is encrypted here. This would confuse users if they see keyids
         of unknown keys and the information can't be "true" because the
         sender could have sent the same information to other people or
         used throw keyids etc.
         */
      char * buf;
      gpgrt_asprintf (&buf, _("The message was not cryptographically signed.\n"
                      "There is no additional information available if it "
                      "was actually sent by '%s' or if someone faked the sender address."), mail->get_sender ().c_str());
      MessageBox (NULL, buf, _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      xfree (buf);
      return S_OK;
    }

  if (!mail->get_sig_fpr())
    {
      std::string buf = _("There was an error verifying the signature.\n"
                           "Full details:\n");
      buf += mail->get_verification_result_dump();
      gpgol_message_box (get_active_hwnd(), buf.c_str(), _("GpgOL"), MB_OK);
    }

  char *uiserver = get_uiserver_name ();
  bool showError = false;
  if (uiserver)
    {
      std::string path (uiserver);
      xfree (uiserver);
      if (path.find("kleopatra.exe") != std::string::npos)
        {
        size_t dpos;
        if ((dpos = path.find(" --daemon")) != std::string::npos)
            {
              path.erase(dpos, strlen(" --daemon"));
            }
          auto ctx = Context::createForEngine(SpawnEngine);
          if (!ctx)
            {
              log_error ("%s:%s: No spawn engine.",
                         SRCNAME, __func__);
            }
            std::string parentWid = std::to_string ((int) (intptr_t) get_active_hwnd ());
            const char *argv[] = {path.c_str(),
                                  "--query",
                                  mail->get_sig_fpr(),
                                  "--parent-windowid",
                                  parentWid.c_str(),
                                  NULL };
            log_debug ("%s:%s: Starting %s %s %s",
                       SRCNAME, __func__, path.c_str(), argv[1], argv[2]);
            Data d(Data::null);
            ctx->spawnAsync(path.c_str(), argv, d, d,
                            d, (GpgME::Context::SpawnFlags) (
                                GpgME::Context::SpawnAllowSetFg |
                                GpgME::Context::SpawnShowWindow));
        }
      else
        {
          showError = true;
        }
    }
  else
    {
      showError = true;
    }

  if (showError)
    {
      MessageBox (NULL,
                  _("Could not find Kleopatra.\n"
                  "Please reinstall Gpg4win with the Kleopatra component enabled."),
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
    }
  return S_OK;
}

HRESULT get_crypto_icon (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  if (mail)
    {
      return getIcon (mail->get_crypto_icon_id (), result);
    }
  return getIcon (IDI_LEVEL_0, result);
}

HRESULT get_is_crypto_mail (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = (VARIANT_BOOL*) xmalloc (sizeof (VARIANT_BOOL));
  *(result->pboolVal) = (mail && (mail->is_signed () || mail->is_encrypted ())) ?
                          VARIANT_TRUE : VARIANT_FALSE;

  return S_OK;
}

HRESULT print_decrypted (LPDISPATCH ctrl)
{
  MY_MAIL_GETTER

  if (!mail)
    {
      log_error ("%s:%s: Failed to get mail.",
                 SRCNAME, __func__);
      return S_OK;
    }
  invoke_oom_method (mail->item(), "PrintOut", NULL);
  return S_OK;
}
