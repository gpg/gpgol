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
#include "dispcache.h"
#include "addressbook.h"
#include "windowmessages.h"

#include <gpgme++/context.h>
#include <gpgme++/data.h>

#undef _
#define _(a) utf8_gettext (a)

using namespace GpgME;

/* This is so super stupid. I bet even Microsft developers laugh
   about the definition of VARIANT_BOOL. And then for COM we
   have to pass pointers to this stuff. */
static VARIANT_BOOL var_true = VARIANT_TRUE;
static VARIANT_BOOL var_false = VARIANT_FALSE;

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

/* getIconDisp
   Loads a PNG image from the resurce converts it into a Bitmap
   and Wraps it in an PictureDispatcher that is returned as result.

   Based on documentation from:
   http://www.codeproject.com/Articles/3537/Loading-JPG-PNG-resources-using-GDI
*/
static LPDISPATCH
getIconDisp (int id)
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
      return nullptr;
    }

  imageSize = SizeofResource (glob_hinst, hResource);
  if (!imageSize)
    return nullptr;

  pResourceData = LockResource (LoadResource(glob_hinst, hResource));

  if (!pResourceData)
    {
      log_error ("%s:%s: failed to load image: %i",
                 SRCNAME, __func__, id);
      return nullptr;
    }

  hBuffer = GlobalAlloc (GMEM_MOVEABLE, imageSize);

  if (hBuffer)
    {
      void* pBuffer = GlobalLock (hBuffer);
      if (pBuffer)
        {
          IStream* pStream = NULL;
          CopyMemory (pBuffer, pResourceData, imageSize);

          if (CreateStreamOnHGlobal (hBuffer, TRUE, &pStream) == S_OK)
            {
              memdbg_addRef (pStream);
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
      return nullptr;
    }

  return pPict;
}

HRESULT
getIcon (int id, VARIANT* result)
{
  if (!result)
    {
      TRACEPOINT;
      return S_OK;
    }

  auto cache = DispCache::instance ();
  result->pdispVal = cache->getDisp (id);

  if (!result->pdispVal)
    {
      result->pdispVal = getIconDisp (id);
      cache->addDisp (id, result->pdispVal);
      memdbg_addRef (result->pdispVal);
    }

  if (result->pdispVal)
    {
      result->vt = VT_DISPATCH;
      result->pdispVal->AddRef();
    }

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
      mail = Mail::getMailForUUID (uid);
      xfree (uid);
    }

  if (mail)
    {
      mail->setCryptoSelectedManually (true);
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
      Mail::locateAllCryptoRecipients_o ();
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
  result->pboolVal = &var_false;

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

  result->pboolVal = value ? &var_true: &var_false;

done:
  gpgol_release (context);
  gpgol_release (mailitem);
  gpgol_release (message);

  return S_OK;
}

static Mail *
get_mail_from_control (LPDISPATCH ctrl, bool *none_selected)
{
  TSTART;
  HRESULT hr;
  LPDISPATCH context = NULL,
             mailitem = NULL;
  *none_selected = false;
  if (!ctrl)
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      TRETURN NULL;
    }
  hr = getContext (ctrl, &context);

  if (hr)
    {
      log_error ("%s:%s:%i : hresult %lx", SRCNAME, __func__, __LINE__,
                 hr);
      TRETURN NULL;
    }
  char *name = get_object_name (context);

  std::string ctx_name;

  if (name)
    {
      ctx_name = name;
      xfree (name);
    }

  if (ctx_name.empty())
    {
      log_error ("%s:%s: Failed to get context name",
                 SRCNAME, __func__);
      gpgol_release (context);
      TRETURN NULL;
    }

  if (!strcmp (ctx_name.c_str(), "_Inspector"))
    {
      mailitem = get_oom_object (context, "CurrentItem");
    }
  else if (!strcmp (ctx_name.c_str(), "_Explorer"))
    {
      /* Avoid showing wrong crypto state if we don't have a reading
         pane. In that case the parser will finish for a mail which is gone
         and the crypto state will not get updated. */

      if (!is_preview_pane_visible (context))
        {
          *none_selected = true;
          gpgol_release (mailitem);
          mailitem = nullptr;
          log_debug ("%s:%s: Preview pane invisible", SRCNAME, __func__);
        }

#if 0
      if (g_ol_version_major >= 16)
        {
          /* Some Versions of Outlook 2016 crashed when accessing the current view
             of the Explorer. This was even reproducible with
             GpgOL disabled and only with Outlook Spy active. If you selected
             the explorer of an Outlook.com resource and then access
             the CurrentView and close the CurrentView again in Outlook Spy
             outlook crashes. See: T3484

             The crash no longer occured at least since build 10228. As
             I'm not sure which Version fixed the crash we don't do a version
             check and just use the same codepath as for Outlook 2010 and
             2013 again.

             Accessing PreviewPane.WordEditor is not a good solution here as
             it requires "Microsoft VBA for Office" (See T4056 ).
             A possible solution for that might be to check if
             "Mail.GetInspector().WordEditor()" returns NULL. In that case we
             know that we also won't get a WordEditor in the preview pane. */
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
#endif

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
              TRETURN NULL;
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
      TRETURN NULL;
    }

  char *uid;
  /* Get the uid of this item. */
  uid = get_unique_id (mailitem, 0, nullptr);
  if (!uid)
    {
      LPMESSAGE msg = get_oom_base_message (mailitem);
      if (!msg)
        {
          log_debug ("%s:%s: Failed to get message for %p",
                   SRCNAME, __func__, mailitem);
          gpgol_release (mailitem);
          TRETURN NULL;
        }
      uid = mapi_get_uid (msg);
      gpgol_release (msg);
      if (!uid)
        {
          log_debug ("%s:%s: Failed to get uid for %p",
                   SRCNAME, __func__, mailitem);
          gpgol_release (mailitem);
          TRETURN NULL;
        }
    }

  auto retV = Mail::searchMailsByUUID(uid);
  xfree (uid);
  if (retV.empty())
    {
      log_error ("%s:%s: Failed to find mail %p in map.",
                 SRCNAME, __func__, mailitem);
    }
  auto ret = retV.front();
  /* This release may have killed the Mail object we obtained earlier
   * if it was just a leftover mail in the ribbon. */
  gpgol_release (mailitem);
  if (!Mail::isValidPtr (ret))
    {
      log_err ("Mail was only valid for this context.");
      TRETURN nullptr;
    }
  TRETURN ret;
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
  result->pboolVal = none_selected ? &var_false : &var_true;

  TRACEPOINT;
  return S_OK;
}

HRESULT get_is_addr_book_enabled (LPDISPATCH ctrl, VARIANT *result)
{
  if (!ctrl)
   {
     log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
     return E_FAIL;
   }
  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = &var_false;
  LPDISPATCH context = nullptr;
  HRESULT hr = getContext (ctrl, &context);

  if (hr || !context)
    {
      log_error ("%s:%s:%i :Failed to get context hresult %lx",
                 SRCNAME, __func__, __LINE__, hr);
      return S_OK;
    }
  auto selection = get_oom_object_s (context, "Selection");
  gpgol_release (context);
  if (!selection)
    {
      log_error ("%s:%s: Failed to get selection.",
                 SRCNAME, __func__);
      return S_OK;
    }
  int count = get_oom_int (selection, "Count");
  if (count == 1)
    {
      auto contact = get_oom_object_s (selection, "Item(1)");
      const auto obj_name = get_object_name_s (contact);
      if (obj_name == "_ContactItem")
        {
          result->pboolVal = &var_true;
          log_dbg ("One contact selected");
        }

      return S_OK;
    }
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
  w_result = utf8_to_wchar (mail->getCryptoSummary ().c_str ());
  result->bstrVal = SysAllocString (w_result);
  xfree (w_result);
  TRACEPOINT;
  return S_OK;
}

HRESULT get_sig_ttip (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BSTR;
  wchar_t *w_result;
  if (mail)
    {
      w_result = utf8_to_wchar (mail->getCryptoOneLine ().c_str());
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
  TRACEPOINT;
  return S_OK;
}

HRESULT get_sig_stip (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BSTR;
  if (none_selected)
    {
      result->bstrVal = SysAllocString (L"");
      TRACEPOINT;
      return S_OK;
    }
  if (!mail || !mail->isCryptoMail ())
    {
      wchar_t *w_result;
      w_result = utf8_to_wchar (utf8_gettext ("You cannot be sure who sent, "
                                  "modified and read the message in transit.\n\n"
                                  "Click here to learn more."));
      result->bstrVal = SysAllocString (w_result);
      xfree (w_result);
      TRACEPOINT;
      return S_OK;
    }
  const auto message = mail->getCryptoDetails_o ();
  wchar_t *w_message = utf8_to_wchar (message.c_str());
  result->bstrVal = SysAllocString (w_message);
  xfree (w_message);
  TRACEPOINT;
  return S_OK;
}

HRESULT launch_cert_details (LPDISPATCH ctrl)
{
  MY_MAIL_GETTER

  if (!mail || (!mail->isSigned () && !mail->isEncrypted ()))
    {
      ShellExecuteA(NULL, NULL, "https://emailselfdefense.fsf.org/en/infographic.html",
                    0, 0, SW_SHOWNORMAL);
      return S_OK;
    }

  if (!mail->isSigned () && mail->isEncrypted ())
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
                      "was actually sent by '%s' or if someone faked the sender address."), mail->getSender_o ().c_str());
      wchar_t *w_msg = utf8_to_wchar (buf);
      wchar_t *w_title = utf8_to_wchar (_("GpgOL"));
      MessageBoxW (NULL, w_msg, w_title,
                   MB_ICONINFORMATION|MB_OK);
      xfree (buf);
      xfree (w_msg);
      xfree (w_title);
      return S_OK;
    }

  if (!mail->getSigFpr ())
    {
      std::string buf = _("There was an error verifying the signature.\n"
                           "Full details:\n");
      buf += mail->getVerificationResultDump ();
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
                                  mail->getSigFpr (),
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
      wchar_t *w_title = utf8_to_wchar (_("GpgOL"));
      wchar_t *w_msg = utf8_to_wchar (_("Could not find Kleopatra.\n"
                  "Please reinstall Gpg4win with the Kleopatra component enabled."));
      MessageBoxW (NULL, w_msg, w_title,
                   MB_ICONINFORMATION|MB_OK);
      xfree (w_title);
      xfree (w_msg);
    }
  return S_OK;
}

HRESULT get_crypto_icon (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  if (mail)
    {
      TRACEPOINT;
      return getIcon (mail->getCryptoIconID (), result);
    }
  TRACEPOINT;
  return getIcon (IDI_LEVEL_0, result);
}

HRESULT get_is_crypto_mail (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = mail && ((mail->isSigned () &&!mail->isEncrypted()) || mail->realyDecryptedSuccessfully ()) ?
                          &var_true : &var_false;

  TRACEPOINT;
  return S_OK;
}

HRESULT get_is_vd_postponed (LPDISPATCH ctrl, VARIANT *result)
{
  MY_MAIL_GETTER

  result->vt = VT_BOOL | VT_BYREF;
  result->pboolVal = mail && (mail->isVdPostponed () ) ?
                          &var_true : &var_false;

  TRACEPOINT;
  return S_OK;
}

HRESULT decrypt_permanently (LPDISPATCH ctrl)
{
  MY_MAIL_GETTER

  if (!mail)
    {
      log_error ("%s:%s: Failed to get mail.",
                 SRCNAME, __func__);
      return S_OK;
    }

  mail->decryptPermanently_o ();
  return S_OK;
}

HRESULT decrypt_manual (LPDISPATCH ctrl)
{
  MY_MAIL_GETTER

  if (!mail)
    {
      log_error ("%s:%s: Failed to get mail.",
                 SRCNAME, __func__);
      return S_OK;
    }

  mail->decryptVerify_o ();
  return S_OK;
}


HRESULT open_contact_key (LPDISPATCH ctrl)
{
  if (!ctrl)
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      return E_FAIL;
    }
  LPDISPATCH context_disp = nullptr;
  HRESULT hr = getContext (ctrl, &context_disp);
  auto context = MAKE_SHARED (context_disp);

  if (hr || !context)
    {
      log_error ("%s:%s:%i :Failed to get context hresult %lx",
                 SRCNAME, __func__, __LINE__, hr);
      return S_OK;
    }

  char *name = get_object_name (context.get ());

  std::string ctx_name;

  if (name)
    {
      ctx_name = name;
      xfree (name);
    }

  if (ctx_name.empty())
    {
      log_error ("%s:%s: Failed to get context name",
                 SRCNAME, __func__);
      return S_OK;
    }

  shared_disp_t contact = nullptr;
  if (!strcmp (ctx_name.c_str(), "_Inspector"))
    {
      contact = get_oom_object_s (context, "CurrentItem");
    }
  else if (!strcmp (ctx_name.c_str(), "_Explorer"))
    {
      auto selection = get_oom_object_s (context, "Selection");
      if (!selection)
        {
          log_error ("%s:%s: Failed to get selection.",
                     SRCNAME, __func__);
          return S_OK;
        }
      int count = get_oom_int (selection, "Count");

      if (!count)
        {
          log_err ("Nothing selected, button should have been inactive.");
          return S_OK;
        }
      else if (count != 1)
        {
          log_err ("More then one selected. Button should have been inactive.");
          return S_OK;
        }
      contact = get_oom_object_s (selection, "Item(1)");
    }


  if (!contact)
    {
      TRACEPOINT;
      return S_OK;
    }

  Addressbook::edit_key_o (contact.get ());

  return S_OK;
}

HRESULT override_file_close ()
{
  TSTART;
  auto inst = GpgolAddin::get_instance ();
  /* We need to get it first as shutdown releases the reference */
  auto app = inst->get_application ();
  app->AddRef ();
  inst->shutdown();
  log_debug ("%s:%s: Shutdown complete. Quitting.",
             SRCNAME, __func__);
  invoke_oom_method (app, "Quit", nullptr);

  TRETURN S_OK;
}

bool g_ignore_next_load = false;

HRESULT override_file_save_as (DISPPARAMS *parms)
{
  TSTART;
  if (!parms)
    {
      STRANGEPOINT;
      TRETURN S_OK;
    }

  /* Check if we are in a dedicated window which we must treat differently. */
  LPDISPATCH context = nullptr;
  if (parms->cArgs == 2 &&
      getContext (parms->rgvarg[1].pdispVal, &context) == S_OK)
    {
      char *name = get_object_name (context);

      std::string ctx_name;

      if (name)
        {
          ctx_name = name;
          xfree (name);
        }
      if (ctx_name == "_Inspector")
        {
          gpgol_release (context);
          TRETURN override_file_save_as_in_window (parms);
        }
    }
  gpgol_release (context);

  /* Do not cancel the event so that the underlying File Save As works. */
  parms->rgvarg[0].pvarVal->boolVal = VARIANT_FALSE;
  /* File->SaveAs triggers an ItemLoad event immediately after this
     callback. To avoid attaching our decrypt / save prevention code
     on that mail we set this global variable so that application-events
     knows that it should ignore the next item load. That way we do
     not interfere and it can just save what is in MAPI (the encrypted
     mail). */
  g_ignore_next_load = true;
  TRETURN S_OK;
}

HRESULT override_file_save_as_in_window (DISPPARAMS *parms)
{
  TSTART;
  if (!parms)
    {
      STRANGEPOINT;
      TRETURN S_OK;
    }

  /* Do not cancel the event so that the underlying File Save As works. */
  parms->rgvarg[0].pvarVal->boolVal = VARIANT_FALSE;

  LPDISPATCH ctrl = parms->rgvarg[1].pdispVal;

  MY_MAIL_GETTER

  if (!mail)
    {
      log_debug ("%s:%s: No mail.",
                 SRCNAME, __func__);
      TRETURN S_OK;
    }

  /* Close the mail async to allow the save as. */
  do_in_ui_thread_async (CLOSE, mail);

  TRETURN S_OK;
}
