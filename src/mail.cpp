/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
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

#include "config.h"
#include "dialogs.h"
#include "common.h"
#include "mail.h"
#include "eventsinks.h"
#include "attachment.h"
#include "mapihelp.h"
#include "mimemaker.h"
#include "message.h"
#include "revert.h"
#include "gpgoladdin.h"
#include "mymapitags.h"
#include "parsecontroller.h"
#include "cryptcontroller.h"
#include "gpgolstr.h"
#include "windowmessages.h"
#include "mlang-charset.h"
#include "wks-helper.h"
#include "keycache.h"
#include "cpphelp.h"

#include <gpgme++/configuration.h>
#include <gpgme++/tofuinfo.h>
#include <gpgme++/verificationresult.h>
#include <gpgme++/decryptionresult.h>
#include <gpgme++/key.h>
#include <gpgme++/context.h>
#include <gpgme++/keylistresult.h>
#include <gpg-error.h>

#include <map>
#include <set>
#include <vector>
#include <memory>


#undef _
#define _(a) utf8_gettext (a)

using namespace GpgME;

static std::map<LPDISPATCH, Mail*> s_mail_map;
static std::map<std::string, Mail*> s_uid_map;
static std::set<std::string> uids_searched;

static Mail *s_last_mail;

#define COPYBUFSIZE (8 * 1024)

Mail::Mail (LPDISPATCH mailitem) :
    m_mailitem(mailitem),
    m_processed(false),
    m_needs_wipe(false),
    m_needs_save(false),
    m_crypt_successful(false),
    m_is_smime(false),
    m_is_smime_checked(false),
    m_is_signed(false),
    m_is_valid(false),
    m_close_triggered(false),
    m_is_html_alternative(false),
    m_needs_encrypt(false),
    m_moss_position(0),
    m_crypto_flags(0),
    m_cached_html_body(nullptr),
    m_cached_plain_body(nullptr),
    m_cached_recipients(nullptr),
    m_type(MSGTYPE_UNKNOWN),
    m_do_inline(false),
    m_is_gsuite(false),
    m_crypt_state(NoCryptMail),
    m_window(nullptr),
    m_is_inline_response(false),
    m_is_forwarded_crypto_mail(false)
{
  if (get_mail_for_item (mailitem))
    {
      log_error ("Mail object for item: %p already exists. Bug.",
                 mailitem);
      return;
    }

  m_event_sink = install_MailItemEvents_sink (mailitem);
  if (!m_event_sink)
    {
      /* Should not happen but in that case we don't add us to the list
         and just release the Mail item. */
      log_error ("%s:%s: Failed to install MailItemEvents sink.",
                 SRCNAME, __func__);
      gpgol_release(mailitem);
      return;
    }
  s_mail_map.insert (std::pair<LPDISPATCH, Mail *> (mailitem, this));
  s_last_mail = this;
}

GPGRT_LOCK_DEFINE(dtor_lock);

Mail::~Mail()
{
  /* This should fix a race condition where the mail is
     deleted before the parser is accessed in the decrypt
     thread. The shared_ptr of the parser then ensures
     that the parser is alive even if the mail is deleted
     while parsing. */
  gpgrt_lock_lock (&dtor_lock);
  log_oom_extra ("%s:%s: dtor: Mail: %p item: %p",
                 SRCNAME, __func__, this, m_mailitem);
  std::map<LPDISPATCH, Mail *>::iterator it;

  log_oom_extra ("%s:%s: Detaching event sink",
                 SRCNAME, __func__);
  detach_MailItemEvents_sink (m_event_sink);
  gpgol_release(m_event_sink);

  log_oom_extra ("%s:%s: Erasing mail",
                 SRCNAME, __func__);
  it = s_mail_map.find(m_mailitem);
  if (it != s_mail_map.end())
    {
      s_mail_map.erase (it);
    }

  if (!m_uuid.empty())
    {
      auto it2 = s_uid_map.find(m_uuid);
      if (it2 != s_uid_map.end())
        {
          s_uid_map.erase (it2);
        }
    }

  log_oom_extra ("%s:%s: releasing mailitem",
                 SRCNAME, __func__);
  gpgol_release(m_mailitem);
  xfree (m_cached_html_body);
  xfree (m_cached_plain_body);
  release_cArray (m_cached_recipients);
  if (!m_uuid.empty())
    {
      log_oom_extra ("%s:%s: destroyed: %p uuid: %s",
                     SRCNAME, __func__, this, m_uuid.c_str());
    }
  else
    {
      log_oom_extra ("%s:%s: non crypto (or sent) mail: %p destroyed",
                     SRCNAME, __func__, this);
    }
  log_oom_extra ("%s:%s: nulling shared pointer",
                 SRCNAME, __func__);
  m_parser = nullptr;
  m_crypter = nullptr;
  gpgrt_lock_unlock (&dtor_lock);
  log_oom_extra ("%s:%s: returning",
                 SRCNAME, __func__);
}

Mail *
Mail::get_mail_for_item (LPDISPATCH mailitem)
{
  if (!mailitem)
    {
      return NULL;
    }
  std::map<LPDISPATCH, Mail *>::iterator it;
  it = s_mail_map.find(mailitem);
  if (it == s_mail_map.end())
    {
      return NULL;
    }
  return it->second;
}

Mail *
Mail::get_mail_for_uuid (const char *uuid)
{
  if (!uuid)
    {
      return NULL;
    }
  auto it = s_uid_map.find(std::string(uuid));
  if (it == s_uid_map.end())
    {
      return NULL;
    }
  return it->second;
}

bool
Mail::is_valid_ptr (const Mail *mail)
{
  auto it = s_mail_map.begin();
  while (it != s_mail_map.end())
    {
      if (it->second == mail)
        return true;
      ++it;
    }
  return false;
}

int
Mail::pre_process_message ()
{
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return 0;
    }
  log_oom_extra ("%s:%s: GetBaseMessage OK for %p.",
                 SRCNAME, __func__, m_mailitem);
  /* Change the message class here. It is important that
     we change the message class in the before read event
     regardless if it is already set to one of GpgOL's message
     classes. Changing the message class (even if we set it
     to the same value again that it already has) causes
     Outlook to reconsider what it "knows" about a message
     and reread data from the underlying base message. */
  mapi_change_message_class (message, 1, &m_type);

  if (m_type == MSGTYPE_UNKNOWN)
    {
      gpgol_release (message);
      return 0;
    }

  /* Create moss attachments here so that they are properly
     hidden when the item is read into the model. */
  m_moss_position = mapi_mark_or_create_moss_attach (message, m_type);
  if (!m_moss_position)
    {
      log_error ("%s:%s: Failed to find moss attachment.",
                 SRCNAME, __func__);
      m_type = MSGTYPE_UNKNOWN;
    }

  gpgol_release (message);
  return 0;
}

static LPDISPATCH
get_attachment (LPDISPATCH mailitem, int pos)
{
  LPDISPATCH attachment;
  LPDISPATCH attachments = get_oom_object (mailitem, "Attachments");
  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      return NULL;
    }

  std::string item_str;
  int count = get_oom_int (attachments, "Count");
  if (count < 1)
    {
      log_debug ("%s:%s: Invalid attachment count: %i.",
                 SRCNAME, __func__, count);
      gpgol_release (attachments);
      return NULL;
    }
  if (pos > 0)
    {
      item_str = std::string("Item(") + std::to_string(pos) + ")";
    }
  else
    {
      item_str = std::string("Item(") + std::to_string(count) + ")";
    }
  attachment = get_oom_object (attachments, item_str.c_str());
  gpgol_release (attachments);

  return attachment;
}

/** Helper to check that all attachments are hidden, to be
  called before crypto. */
int
Mail::check_attachments () const
{
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      return 1;
    }
  int count = get_oom_int (attachments, "Count");
  if (!count)
    {
      gpgol_release (attachments);
      return 0;
    }

  std::string message;

  if (is_encrypted () && is_signed ())
    {
      message += _("Not all attachments were encrypted or signed.\n"
                   "The unsigned / unencrypted attachments are:\n\n");
    }
  else if (is_signed ())
    {
      message += _("Not all attachments were signed.\n"
                   "The unsigned attachments are:\n\n");
    }
  else if (is_encrypted ())
    {
      message += _("Not all attachments were encrypted.\n"
                   "The unencrypted attachments are:\n\n");
    }
  else
    {
      gpgol_release (attachments);
      return 0;
    }

  bool foundOne = false;

  for (int i = 1; i <= count; i++)
    {
      std::string item_str;
      item_str = std::string("Item(") + std::to_string (i) + ")";
      LPDISPATCH oom_attach = get_oom_object (attachments, item_str.c_str ());
      if (!oom_attach)
        {
          log_error ("%s:%s: Failed to get attachment.",
                     SRCNAME, __func__);
          continue;
        }
      VARIANT var;
      VariantInit (&var);
      if (get_pa_variant (oom_attach, PR_ATTACHMENT_HIDDEN_DASL, &var) ||
          (var.vt == VT_BOOL && var.boolVal == VARIANT_FALSE))
        {
          foundOne = true;
          char *dispName = get_oom_string (oom_attach, "DisplayName");
          message += dispName ? dispName : "Unknown";
          xfree (dispName);
          message += "\n";
        }
      VariantClear (&var);
      gpgol_release (oom_attach);
    }
  gpgol_release (attachments);
  if (foundOne)
    {
      message += "\n";
      message += _("Note: The attachments may be encrypted or signed "
                    "on a file level but the GpgOL status does not apply to them.");
      wchar_t *wmsg = utf8_to_wchar (message.c_str ());
      wchar_t *wtitle = utf8_to_wchar (_("GpgOL Warning"));
      MessageBoxW (get_active_hwnd (), wmsg, wtitle,
                   MB_ICONWARNING|MB_OK);
      xfree (wmsg);
      xfree (wtitle);
    }
  return 0;
}

/** Get the cipherstream of the mailitem. */
static LPSTREAM
get_attachment_stream (LPDISPATCH mailitem, int pos)
{
  if (!pos)
    {
      log_debug ("%s:%s: Called with zero pos.",
                 SRCNAME, __func__);
      return NULL;
    }
  LPDISPATCH attachment = get_attachment (mailitem, pos);
  LPSTREAM stream = NULL;

  if (!attachment)
    {
      // For opened messages that have ms-tnef type we
      // create the moss attachment but don't find it
      // in the OOM. Try to find it through MAPI.
      HRESULT hr;
      log_debug ("%s:%s: Failed to find MOSS Attachment. "
                 "Fallback to MAPI.", SRCNAME, __func__);
      LPMESSAGE message = get_oom_message (mailitem);
      if (!message)
        {
          log_debug ("%s:%s: Failed to get MAPI Interface.",
                     SRCNAME, __func__);
          return NULL;
        }
      hr = message->OpenProperty (PR_BODY_A, &IID_IStream, 0, 0,
                                  (LPUNKNOWN*)&stream);
      if (hr)
        {
          log_debug ("%s:%s: OpenProperty failed: hr=%#lx",
                     SRCNAME, __func__, hr);
          return NULL;
        }
      return stream;
    }

  LPATTACH mapi_attachment = NULL;

  mapi_attachment = (LPATTACH) get_oom_iunknown (attachment,
                                                 "MapiObject");
  gpgol_release (attachment);
  if (!mapi_attachment)
    {
      log_debug ("%s:%s: Failed to get MapiObject of attachment: %p",
                 SRCNAME, __func__, attachment);
      return NULL;
    }
  if (FAILED (mapi_attachment->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream,
                                             0, MAPI_MODIFY, (LPUNKNOWN*) &stream)))
    {
      log_debug ("%s:%s: Failed to open stream for mapi_attachment: %p",
                 SRCNAME, __func__, mapi_attachment);
      gpgol_release (mapi_attachment);
    }
  return stream;
}

#if 0

This should work. But Outlook says no. See the comment in set_pa_variant
about this. I left the code here as an example how to work with
safearrays and how this probably should work.

static int
copy_data_property(LPDISPATCH target, std::shared_ptr<Attachment> attach)
{
  VARIANT var;
  VariantInit (&var);

  /* Get the size */
  off_t size = attach->get_data ().seek (0, SEEK_END);
  attach->get_data ().seek (0, SEEK_SET);

  if (!size)
    {
      TRACEPOINT;
      return 1;
    }

  if (!get_pa_variant (target, PR_ATTACH_DATA_BIN_DASL, &var))
    {
      log_debug("Have variant. type: %x", var.vt);
    }
  else
    {
      log_debug("failed to get variant.");
    }

  /* Set the type to an array of unsigned chars (OLE SAFEARRAY) */
  var.vt = VT_ARRAY | VT_UI1;

  /* Set up the bounds structure */
  SAFEARRAYBOUND rgsabound[1];
  rgsabound[0].cElements = static_cast<unsigned long> (size);
  rgsabound[0].lLbound = 0;

  /* Create an OLE SAFEARRAY */
  var.parray = SafeArrayCreate (VT_UI1, 1, rgsabound);
  if (var.parray == NULL)
    {
      TRACEPOINT;
      VariantClear(&var);
      return 1;
    }

  void *buffer = NULL;
  /* Get a safe pointer to the array */
  if (SafeArrayAccessData(var.parray, &buffer) != S_OK)
    {
      TRACEPOINT;
      VariantClear(&var);
      return 1;
    }

  /* Copy data to it */
  size_t nread = attach->get_data ().read (buffer, static_cast<size_t> (size));

  if (nread != static_cast<size_t> (size))
    {
      TRACEPOINT;
      VariantClear(&var);
      return 1;
    }

  /*/ Unlock the variant data */
  if (SafeArrayUnaccessData(var.parray) != S_OK)
    {
      TRACEPOINT;
      VariantClear(&var);
      return 1;
    }

  if (set_pa_variant (target, PR_ATTACH_DATA_BIN_DASL, &var))
    {
      TRACEPOINT;
      VariantClear(&var);
      return 1;
    }

  VariantClear(&var);
  return 0;
}
#endif

static int
copy_attachment_to_file (std::shared_ptr<Attachment> att, HANDLE hFile)
{
  char copybuf[COPYBUFSIZE];
  size_t nread;

  /* Security considerations: Writing the data to a temporary
     file is necessary as neither MAPI manipulation works in the
     read event to transmit the data nor Property Accessor
     works (see above). From a security standpoint there is a
     short time where the temporary files are on disk. Tempdir
     should be protected so that only the user can read it. Thus
     we have a local attack that could also take the data out
     of Outlook. FILE_SHARE_READ is necessary so that outlook
     can read the file.

     A bigger concern is that the file is manipulated
     by another software to fake the signature state. So
     we keep the write exlusive to us.

     We delete the file before closing the write file handle.
  */

  /* Make sure we start at the beginning */
  att->get_data ().seek (0, SEEK_SET);
  while ((nread = att->get_data ().read (copybuf, COPYBUFSIZE)))
    {
      DWORD nwritten;
      if (!WriteFile (hFile, copybuf, nread, &nwritten, NULL))
        {
          log_error ("%s:%s: Failed to write in tmp attachment.",
                     SRCNAME, __func__);
          return 1;
        }
      if (nread != nwritten)
        {
          log_error ("%s:%s: Write truncated.",
                     SRCNAME, __func__);
          return 1;
        }
    }
  return 0;
}

/** Sets some meta data on the last attachment added. The meta
  data is taken from the attachment object. */
static int
fixup_last_attachment (LPDISPATCH mail, std::shared_ptr<Attachment> attachment)
{
  /* Currently we only set content id */
  if (attachment->get_content_id ().empty())
    {
      log_debug ("%s:%s: Content id not found.",
                 SRCNAME, __func__);
      return 0;
    }

  LPDISPATCH attach = get_attachment (mail, -1);
  if (!attach)
    {
      log_error ("%s:%s: No attachment.",
                 SRCNAME, __func__);
      return 1;
    }
  int ret = put_pa_string (attach,
                           PR_ATTACH_CONTENT_ID_DASL,
                           attachment->get_content_id ().c_str());
  gpgol_release (attach);
  return ret;
}

/** Helper to update the attachments of a mail object in oom.
  does not modify the underlying mapi structure. */
static int
add_attachments(LPDISPATCH mail,
                std::vector<std::shared_ptr<Attachment> > attachments)
{
  int err = 0;
  for (auto att: attachments)
    {
      if (att->get_display_name().empty())
        {
          log_error ("%s:%s: Ignoring attachment without display name.",
                     SRCNAME, __func__);
          continue;
        }
      wchar_t* wchar_name = utf8_to_wchar (att->get_display_name().c_str());
      HANDLE hFile;
      wchar_t* wchar_file = get_tmp_outfile (wchar_name,
                                             &hFile);
      if (copy_attachment_to_file (att, hFile))
        {
          log_error ("%s:%s: Failed to copy attachment %s to temp file",
                     SRCNAME, __func__, att->get_display_name().c_str());
          err = 1;
        }
      if (add_oom_attachment (mail, wchar_file, wchar_name))
        {
          log_error ("%s:%s: Failed to add attachment: %s",
                     SRCNAME, __func__, att->get_display_name().c_str());
          err = 1;
        }
      if (!DeleteFileW (wchar_file))
        {
          log_error ("%s:%s: Failed to delete tmp attachment for: %s",
                     SRCNAME, __func__, att->get_display_name().c_str());
          err = 1;
        }
      xfree (wchar_file);
      xfree (wchar_name);

      err = fixup_last_attachment (mail, att);
    }
  return err;
}

GPGRT_LOCK_DEFINE(parser_lock);

static DWORD WINAPI
do_parsing (LPVOID arg)
{
  gpgrt_lock_lock (&dtor_lock);
  /* We lock with mail dtors so we can be sure the mail->parser
     call is valid. */
  Mail *mail = (Mail *)arg;
  if (!Mail::is_valid_ptr (mail))
    {
      log_debug ("%s:%s: canceling parsing for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&dtor_lock);
      return 0;
    }
  /* This takes a shared ptr of parser. So the parser is
     still valid when the mail is deleted. */
  auto parser = mail->parser();
  gpgrt_lock_unlock (&dtor_lock);

  gpgrt_lock_lock (&parser_lock);
  gpgrt_lock_lock (&invalidate_lock);
  /* We lock the parser here to avoid too many
     decryption attempts if there are
     multiple mailobjects which might have already
     been deleted (e.g. by quick switches of the mailview.)
     Let's rather be a bit slower.
     */
  log_debug ("%s:%s: preparing the parser for: %p",
             SRCNAME, __func__, arg);

  if (!Mail::is_valid_ptr (mail))
    {
      log_debug ("%s:%s: cancel for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&invalidate_lock);
      gpgrt_lock_unlock (&parser_lock);
      return 0;
    }

  if (!parser)
    {
      log_error ("%s:%s: no parser found for mail: %p",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&invalidate_lock);
      gpgrt_lock_unlock (&parser_lock);
      return -1;
    }
  parser->parse();
  do_in_ui_thread (PARSING_DONE, arg);
  gpgrt_lock_unlock (&invalidate_lock);
  gpgrt_lock_unlock (&parser_lock);
  return 0;
}

/* How encryption is done:

   There are two modes of encryption. Synchronous and Async.
   If async is used depends on the value of mail->is_inline_response.

   Synchronous crypto:

   > Send Event < | State NoCryptMail
   Needs Crypto ? (get_gpgol_draft_info_flags != 0)

   -> No:
      Pass send -> unencrypted mail.

   -> Yes:
      mail->update_oom_data
      State = Mail::NeedsFirstAfterWrite
      check_inline_response
      invoke_oom_method (m_object, "Save", NULL);

      > Write Event <
      Pass because is_crypto_mail is false (not a decrypted mail)

      > AfterWrite Event < | State NeedsFirstAfterWrite
      State = NeedsActualCrypo
      encrypt_sign_start
        collect_input_data
        -> Check if Inline PGP should be used
        do_crypt
          -> Resolve keys / do crypto

          State = NeedsUpdateInMAPI
          update_crypt_mapi
          crypter->update_mail_mapi
           if (inline) (Meaning PGP/Inline)
          <-- do nothing.
           else
            build MSOXSMIME attachment and clear body / attachments.

          State = NeedsUpdateInOOM
      <- Back to Send Event
      update_crypt_oom
        -> Cleans body or sets PGP/Inline body. (inline_body_to_body)
      State = WantsSendMIME or WantsSendInline

      -> Saftey check "has_crypted_or_empty_body"
      -> If MIME Mail do the T3656 check.

    Send.

    State order for "inline_response" (sync) Mails.
    NoCryptMail
    NeedsFirstAfterWrite
    NeedsActualCrypto
    NeedsUpdateInMAPI
    NeedsUpdateInOOM
    WantsSendMIME (or inline for PGP Inline)
    -> Send.

    State order for async Mails
    NoCryptMail
    NeedsFirstAfterWrite
    NeedsActualCrypto
    -> Cancel Send.
    Windowmessages -> Crypto Done
    NeedsUpdateInOOM
    NeedsSecondAfterWrite
    trigger Save.
    NeedsUpdateInMAPI
    WantsSendMIME
    trigger Send.
*/
static DWORD WINAPI
do_crypt (LPVOID arg)
{
  gpgrt_lock_lock (&dtor_lock);
  /* We lock with mail dtors so we can be sure the mail->parser
     call is valid. */
  Mail *mail = (Mail *)arg;
  if (!Mail::is_valid_ptr (mail))
    {
      log_debug ("%s:%s: canceling crypt for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&dtor_lock);
      return 0;
    }
  if (mail->crypt_state() != Mail::NeedsActualCrypt)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, mail->crypt_state ());
      mail->set_window_enabled (true);
      gpgrt_lock_unlock (&dtor_lock);
      return -1;
    }

  /* This takes a shared ptr of crypter. So the crypter is
     still valid when the mail is deleted. */
  auto crypter = mail->crypter();
  gpgrt_lock_unlock (&dtor_lock);

  if (!crypter)
    {
      log_error ("%s:%s: no crypter found for mail: %p",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&parser_lock);
      mail->set_window_enabled (true);
      return -1;
    }

  int rc = crypter->do_crypto();

  gpgrt_lock_lock (&dtor_lock);
  if (!Mail::is_valid_ptr (mail))
    {
      log_debug ("%s:%s: aborting crypt for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgrt_lock_unlock (&dtor_lock);
      return 0;
    }

  if (rc == -1)
    {
      mail->reset_crypter ();
      crypter = nullptr;
      gpgol_bug (mail->get_window (),
                 ERR_CRYPT_RESOLVER_FAILED);
    }

  mail->set_window_enabled (true);

  if (rc)
    {
      log_debug ("%s:%s: crypto failed for: %p with: %i",
                 SRCNAME, __func__, arg, rc);
      mail->set_crypt_state (Mail::NoCryptMail);
      mail->reset_crypter ();
      crypter = nullptr;
      gpgrt_lock_unlock (&dtor_lock);
      return rc;
    }

  if (!mail->is_inline_response ())
    {
      mail->set_crypt_state (Mail::NeedsUpdateInOOM);
      gpgrt_lock_unlock (&dtor_lock);
      // This deletes the Mail in Outlook 2010
      do_in_ui_thread (CRYPTO_DONE, arg);
      log_debug ("%s:%s: UI thread finished for %p",
                 SRCNAME, __func__, arg);
    }
  else
    {
      mail->set_crypt_state (Mail::NeedsUpdateInMAPI);
      mail->update_crypt_mapi ();
      mail->set_crypt_state (Mail::NeedsUpdateInOOM);
      gpgrt_lock_unlock (&dtor_lock);
    }
  /* This works around a bug in pinentry that it might
     bring the wrong window to front. So after encryption /
     signing we bring outlook back to front.

     See GnuPG-Bug-Id: T3732
     */
  do_in_ui_thread_async (BRING_TO_FRONT, nullptr);
  log_debug ("%s:%s: crypto thread for %p finished",
             SRCNAME, __func__, arg);
  return 0;
}

bool
Mail::is_crypto_mail() const
{
  if (m_type == MSGTYPE_UNKNOWN || m_type == MSGTYPE_GPGOL ||
      m_type == MSGTYPE_SMIME)
    {
      /* Not a message for us. */
      return false;
    }
  return true;
}

int
Mail::decrypt_verify()
{
  if (!is_crypto_mail())
    {
      log_debug ("%s:%s: Decrypt Verify for non crypto mail: %p.",
                 SRCNAME, __func__, m_mailitem);
      return 0;
    }
  if (m_needs_wipe)
    {
      log_error ("%s:%s: Decrypt verify called for msg that needs wipe: %p",
                 SRCNAME, __func__, m_mailitem);
      return 1;
    }
  set_uuid ();
  m_processed = true;


  /* Insert placeholder */
  char *placeholder_buf;
  if (m_type == MSGTYPE_GPGOL_WKS_CONFIRMATION)
    {
      gpgrt_asprintf (&placeholder_buf, opt.prefer_html ? decrypt_template_html :
                      decrypt_template,
                      "OpenPGP",
                      _("Pubkey directory confirmation"),
                      _("This is a confirmation request to publish your Pubkey in the "
                        "directory for your domain.\n\n"
                        "<p>If you did not request to publish your Pubkey in your providers "
                        "directory, simply ignore this message.</p>\n"));
    }
  else if (gpgrt_asprintf (&placeholder_buf, opt.prefer_html ? decrypt_template_html :
                      decrypt_template,
                      is_smime() ? "S/MIME" : "OpenPGP",
                      _("Encrypted message"),
                      _("Please wait while the message is being decrypted / verified...")) == -1)
    {
      log_error ("%s:%s: Failed to format placeholder.",
                 SRCNAME, __func__);
      return 1;
    }

  if (opt.prefer_html)
    {
      m_orig_body = get_oom_string (m_mailitem, "HTMLBody");
      if (put_oom_string (m_mailitem, "HTMLBody", placeholder_buf))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
        }
    }
  else
    {
      m_orig_body = get_oom_string (m_mailitem, "Body");
      if (put_oom_string (m_mailitem, "Body", placeholder_buf))
        {
          log_error ("%s:%s: Failed to modify body of item.",
                     SRCNAME, __func__);
        }
    }
  xfree (placeholder_buf);

  /* Do the actual parsing */
  auto cipherstream = get_attachment_stream (m_mailitem, m_moss_position);

  if (m_type == MSGTYPE_GPGOL_WKS_CONFIRMATION)
    {
      WKSHelper::instance ()->handle_confirmation_read (this, cipherstream);
      return 0;
    }

  if (!cipherstream)
    {
      log_debug ("%s:%s: Failed to get cipherstream.",
                 SRCNAME, __func__);
      return 1;
    }

  m_parser = std::shared_ptr <ParseController> (new ParseController (cipherstream, m_type));
  m_parser->setSender(GpgME::UserID::addrSpecFromString(get_sender().c_str()));
  log_mime_parser ("%s:%s: Parser for \"%s\" is %p",
                   SRCNAME, __func__, get_subject ().c_str(), m_parser.get());
  gpgol_release (cipherstream);

  HANDLE parser_thread = CreateThread (NULL, 0, do_parsing, (LPVOID) this, 0,
                                       NULL);

  if (!parser_thread)
    {
      log_error ("%s:%s: Failed to create decrypt / verify thread.",
                 SRCNAME, __func__);
    }
  CloseHandle (parser_thread);

  return 0;
}

void find_and_replace(std::string& source, const std::string &find,
                      const std::string &replace)
{
  for(std::string::size_type i = 0; (i = source.find(find, i)) != std::string::npos;)
    {
      source.replace(i, find.length(), replace);
      i += replace.length();
    }
}

void
Mail::update_body()
{
  if (!m_parser)
    {
      TRACEPOINT;
      return;
    }

  const auto error = m_parser->get_formatted_error ();
  if (!error.empty())
    {
      if (opt.prefer_html)
        {
          if (put_oom_string (m_mailitem, "HTMLBody",
                              error.c_str ()))
            {
              log_error ("%s:%s: Failed to modify html body of item.",
                         SRCNAME, __func__);
            }
          else
            {
              log_debug ("%s:%s: Set error html to: '%s'",
                         SRCNAME, __func__, error.c_str ());
            }

        }
      else
        {
          if (put_oom_string (m_mailitem, "Body",
                              error.c_str ()))
            {
              log_error ("%s:%s: Failed to modify html body of item.",
                         SRCNAME, __func__);
            }
          else
            {
              log_debug ("%s:%s: Set error plain to: '%s'",
                         SRCNAME, __func__, error.c_str ());
            }
        }
      return;
    }
  if (m_verify_result.error())
    {
      log_error ("%s:%s: Verification failed. Restoring Body.",
                 SRCNAME, __func__);
      if (opt.prefer_html)
        {
          if (put_oom_string (m_mailitem, "HTMLBody", m_orig_body.c_str ()))
            {
              log_error ("%s:%s: Failed to modify html body of item.",
                         SRCNAME, __func__);
            }
        }
      else
        {
          if (put_oom_string (m_mailitem, "Body", m_orig_body.c_str ()))
            {
              log_error ("%s:%s: Failed to modify html body of item.",
                         SRCNAME, __func__);
            }
        }
      return;
    }
  // No need to carry body anymore
  m_orig_body = std::string();
  auto html = m_parser->get_html_body ();
  /** Outlook does not show newlines if \r\r\n is a newline. We replace
    these as apparently some other buggy MUA sends this. */
  find_and_replace (html, "\r\r\n", "\r\n");
  if (opt.prefer_html && !html.empty())
    {
      char *converted = ansi_charset_to_utf8 (m_parser->get_html_charset().c_str(),
                                              html.c_str(), html.size());
      int ret = put_oom_string (m_mailitem, "HTMLBody", converted ? converted : "");
      xfree (converted);
      if (ret)
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
        }
      return;
    }
  auto body = m_parser->get_body ();
  find_and_replace (body, "\r\r\n", "\r\n");
  char *converted = ansi_charset_to_utf8 (m_parser->get_body_charset().c_str(),
                                          body.c_str(), body.size());
  int ret = put_oom_string (m_mailitem, "Body", converted ? converted : "");
  xfree (converted);
  if (ret)
    {
      log_error ("%s:%s: Failed to modify body of item.",
                 SRCNAME, __func__);
    }
  return;
}

void
Mail::parsing_done()
{
  TRACEPOINT;
  log_oom_extra ("Mail %p Parsing done for parser: %p",
                 this, m_parser.get());
  if (!m_parser)
    {
      /* This should not happen but it happens when outlook
         sends multiple ItemLoad events for the same Mail
         Object. In that case it could happen that one
         parser was already done while a second is now
         returning for the wrong mail (as it's looked up
         by uuid.)

         We have a check in get_uuid that the uuid was
         not in the map before (and the parser is replaced).
         So this really really should not happen. We
         handle it anyway as we crash otherwise.

         It should not happen because the parser is only
         created in decrypt_verify which is called in the
         read event. And even in there we check if the parser
         was set.
         */
      log_error ("%s:%s: No parser obj. For mail: %p",
                 SRCNAME, __func__, this);
      return;
    }
  /* Store the results. */
  m_decrypt_result = m_parser->decrypt_result ();
  m_verify_result = m_parser->verify_result ();

  m_crypto_flags = 0;
  if (!m_decrypt_result.isNull())
    {
      m_crypto_flags |= 1;
    }
  if (m_verify_result.numSignatures())
    {
      m_crypto_flags |= 2;
    }

  update_sigstate ();
  m_needs_wipe = true;

  TRACEPOINT;
  /* Set categories according to the result. */
  update_categories();

  TRACEPOINT;
  /* Update the body */
  update_body();
  TRACEPOINT;

  /* Check that there are no unsigned / unencrypted messages. */
  check_attachments ();

  /* Update attachments */
  if (add_attachments (m_mailitem, m_parser->get_attachments()))
    {
      log_error ("%s:%s: Failed to update attachments.",
                 SRCNAME, __func__);
    }

  log_debug ("%s:%s: Delayed invalidate to update sigstate.",
             SRCNAME, __func__);
  CloseHandle(CreateThread (NULL, 0, delayed_invalidate_ui, (LPVOID) this, 0,
                            NULL));
  TRACEPOINT;
  return;
}

int
Mail::encrypt_sign_start ()
{
  if (m_crypt_state != NeedsActualCrypt)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      return -1;
    }
  int flags = 0;
  if (!needs_crypto())
    {
      return 0;
    }
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return -1;
    }
  flags = get_gpgol_draft_info_flags (message);
  gpgol_release (message);

  const auto window = get_active_hwnd ();

  if (m_is_gsuite)
    {
      auto att_table = mapi_create_attach_table (message, 0);
      int n_att_usable = count_usable_attachments (att_table);
      mapi_release_attach_table (att_table);
      /* Check for attachments if we have some abort. */

      if (n_att_usable)
        {
          wchar_t *w_title = utf8_to_wchar (_(
                                              "GpgOL: Oops, G Suite Sync account detected"));
          wchar_t *msg = utf8_to_wchar (
                      _("G Suite Sync breaks outgoing crypto mails "
                        "with attachments.\nUsing crypto and attachments "
                        "with G Suite Sync is not supported.\n\n"
                        "See: https://dev.gnupg.org/T3545 for details."));
          MessageBoxW (window,
                       msg,
                       w_title,
                       MB_ICONINFORMATION|MB_OK);
          xfree (msg);
          xfree (w_title);
          return -1;
        }
    }

  m_do_inline = m_is_gsuite ? true : opt.inline_pgp;

  GpgME::Protocol proto = opt.enable_smime ? GpgME::UnknownProtocol: GpgME::OpenPGP;
  m_crypter = std::shared_ptr <CryptController> (new CryptController (this, flags & 1,
                                                                      flags & 2,
                                                                      proto));

  // Careful from here on we have to check every
  // error condition with window enabling again.
  set_window_enabled (false);
  if (m_crypter->collect_data ())
    {
      log_error ("%s:%s: Crypter for mail %p failed to collect data.",
                 SRCNAME, __func__, this);
      set_window_enabled (true);
      return -1;
    }

  if (!m_is_inline_response)
    {
      CloseHandle(CreateThread (NULL, 0, do_crypt,
                                (LPVOID) this, 0,
                                NULL));
    }
  else
    {
      do_crypt (this);
    }
  return 0;
}

int
Mail::needs_crypto ()
{
  LPMESSAGE message = get_oom_message (m_mailitem);
  bool ret;
  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
      return false;
    }
  ret = get_gpgol_draft_info_flags (message);
  gpgol_release(message);
  return ret;
}

int
Mail::wipe (bool force)
{
  if (!m_needs_wipe && !force)
    {
      return 0;
    }
  log_debug ("%s:%s: Removing plaintext from mailitem: %p.",
             SRCNAME, __func__, m_mailitem);
  if (put_oom_string (m_mailitem, "HTMLBody",
                      ""))
    {
      if (put_oom_string (m_mailitem, "Body",
                          ""))
        {
          log_debug ("%s:%s: Failed to wipe mailitem: %p.",
                     SRCNAME, __func__, m_mailitem);
          return -1;
        }
      return -1;
    }
  m_needs_wipe = false;
  return 0;
}

int
Mail::update_oom_data ()
{
  char *buf = nullptr;
  log_debug ("%s:%s", SRCNAME, __func__);

  if (!is_crypto_mail())
    {
      /* Update the body format. */
      m_is_html_alternative = get_oom_int (m_mailitem, "BodyFormat") > 1;

      /* Store the body. It was not obvious for me (aheinecke) how
         to access this through MAPI. */
      if (m_is_html_alternative)
        {
          log_debug ("%s:%s: Is html alternative mail.", SRCNAME, __func__);
          xfree (m_cached_html_body);
          m_cached_html_body = get_oom_string (m_mailitem, "HTMLBody");
        }
      xfree (m_cached_plain_body);
      m_cached_plain_body = get_oom_string (m_mailitem, "Body");

      release_cArray (m_cached_recipients);
      m_cached_recipients = get_recipients ();
    }
  /* For some reason outlook may store the recipient address
     in the send using account field. If we have SMTP we prefer
     the SenderEmailAddress string. */
  if (is_crypto_mail ())
    {
      /* This is the case where we are reading a mail and not composing.
         When composing we need to use the SendUsingAccount because if
         you send from the folder of userA but change the from to userB
         outlook will keep the SenderEmailAddress of UserA. This is all
         so horrible. */
      buf = get_sender_SenderEMailAddress (m_mailitem);

      if (!buf)
        {
          /* Try the sender Object */
          buf = get_sender_Sender (m_mailitem);
        }
    }

  if (!buf)
    {
      buf = get_sender_SendUsingAccount (m_mailitem, &m_is_gsuite);
    }
  if (!buf && !is_crypto_mail ())
    {
      /* Try the sender Object */
      buf = get_sender_Sender (m_mailitem);
    }
  if (!buf)
    {
      /* We don't have s sender object or SendUsingAccount,
         well, in that case fall back to the current user. */
      buf = get_sender_CurrentUser (m_mailitem);
    }
  if (!buf)
    {
      log_debug ("%s:%s: All fallbacks failed.",
                 SRCNAME, __func__);
      return -1;
    }
  m_sender = buf;
  xfree (buf);
  return 0;
}

std::string
Mail::get_sender ()
{
  if (m_sender.empty())
    update_oom_data();
  return m_sender;
}

std::string
Mail::get_cached_sender ()
{
  return m_sender;
}

int
Mail::close_all_mails ()
{
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  TRACEPOINT;
  std::map<LPDISPATCH, Mail *> mail_map_copy = s_mail_map;
  for (it = mail_map_copy.begin(); it != mail_map_copy.end(); ++it)
    {
      /* XXX For non racy code the is_valid_ptr check should not
         be necessary but we crashed sometimes closing a destroyed
         mail. */
      if (!is_valid_ptr (it->second))
        {
          log_debug ("%s:%s: Already deleted mail for %p",
                   SRCNAME, __func__, it->first);
          continue;
        }

      if (!it->second->is_crypto_mail())
        {
          continue;
        }
      if (close_inspector (it->second) || close (it->second))
        {
          log_error ("Failed to close mail: %p ", it->first);
          /* Should not happen */
          if (is_valid_ptr (it->second) && it->second->revert())
            {
              err++;
            }
        }
    }
  return err;
}
int
Mail::revert_all_mails ()
{
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = s_mail_map.begin(); it != s_mail_map.end(); ++it)
    {
      if (it->second->revert ())
        {
          log_error ("Failed to revert mail: %p ", it->first);
          err++;
          continue;
        }

      it->second->set_needs_save (true);
      if (!invoke_oom_method (it->first, "Save", NULL))
        {
          log_error ("Failed to save reverted mail: %p ", it->second);
          err++;
          continue;
        }
    }
  return err;
}

int
Mail::wipe_all_mails ()
{
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = s_mail_map.begin(); it != s_mail_map.end(); ++it)
    {
      if (it->second->wipe ())
        {
          log_error ("Failed to wipe mail: %p ", it->first);
          err++;
        }
    }
  return err;
}

int
Mail::revert ()
{
  int err = 0;
  if (!m_processed)
    {
      return 0;
    }

  err = gpgol_mailitem_revert (m_mailitem);
  if (err == -1)
    {
      log_error ("%s:%s: Message revert failed falling back to wipe.",
                 SRCNAME, __func__);
      return wipe ();
    }
  /* We need to reprocess the mail next time around. */
  m_processed = false;
  m_needs_wipe = false;
  return 0;
}

bool
Mail::is_smime ()
{
  msgtype_t msgtype;
  LPMESSAGE message;

  if (m_is_smime_checked)
    {
      return m_is_smime;
    }

  message = get_oom_message (m_mailitem);

  if (!message)
    {
      log_error ("%s:%s: No message?",
                 SRCNAME, __func__);
      return false;
    }

  msgtype = mapi_get_message_type (message);
  m_is_smime = msgtype == MSGTYPE_GPGOL_OPAQUE_ENCRYPTED ||
               msgtype == MSGTYPE_GPGOL_OPAQUE_SIGNED;

  /* Check if it is an smime mail. Multipart signed can
     also be true. */
  if (!m_is_smime && msgtype == MSGTYPE_GPGOL_MULTIPART_SIGNED)
    {
      char *proto;
      char *ct = mapi_get_message_content_type (message, &proto, NULL);
      if (ct && proto)
        {
          m_is_smime = (!strcmp (proto, "application/pkcs7-signature") ||
                        !strcmp (proto, "application/x-pkcs7-signature"));
        }
      else
        {
          log_error ("%s:%s: No protocol in multipart / signed mail.",
                     SRCNAME, __func__);
        }
      xfree (proto);
      xfree (ct);
    }
  gpgol_release (message);
  m_is_smime_checked  = true;
  return m_is_smime;
}

static std::string
get_string (LPDISPATCH item, const char *str)
{
  char *buf = get_oom_string (item, str);
  if (!buf)
    return std::string();
  std::string ret = buf;
  xfree (buf);
  return ret;
}

std::string
Mail::get_subject() const
{
  return get_string (m_mailitem, "Subject");
}

std::string
Mail::get_body() const
{
  return get_string (m_mailitem, "Body");
}

std::string
Mail::get_html_body() const
{
  return get_string (m_mailitem, "HTMLBody");
}

char **
Mail::get_recipients() const
{
  LPDISPATCH recipients = get_oom_object (m_mailitem, "Recipients");
  if (!recipients)
    {
      TRACEPOINT;
      return nullptr;
    }
  auto ret = get_oom_recipients (recipients);
  gpgol_release (recipients);
  return ret;
}

int
Mail::close_inspector (Mail *mail)
{
  LPDISPATCH inspector = get_oom_object (mail->item(), "GetInspector");
  HRESULT hr;
  DISPID dispid;
  if (!inspector)
    {
      log_debug ("%s:%s: No inspector.",
                 SRCNAME, __func__);
      return -1;
    }

  dispid = lookup_oom_dispid (inspector, "Close");
  if (dispid != DISPID_UNKNOWN)
    {
      VARIANT aVariant[1];
      DISPPARAMS dispparams;

      dispparams.rgvarg = aVariant;
      dispparams.rgvarg[0].vt = VT_INT;
      dispparams.rgvarg[0].intVal = 1;
      dispparams.cArgs = 1;
      dispparams.cNamedArgs = 0;

      hr = inspector->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                              DISPATCH_METHOD, &dispparams,
                              NULL, NULL, NULL);
      if (hr != S_OK)
        {
          log_debug ("%s:%s: Failed to close inspector: %#lx",
                     SRCNAME, __func__, hr);
          gpgol_release (inspector);
          return -1;
        }
    }
  gpgol_release (inspector);
  return 0;
}

/* static */
int
Mail::close (Mail *mail)
{
  VARIANT aVariant[1];
  DISPPARAMS dispparams;

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = 1;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  log_oom_extra ("%s:%s: Invoking close for: %p",
                 SRCNAME, __func__, mail->item());
  mail->set_close_triggered (true);
  int rc = invoke_oom_method_with_parms (mail->item(), "Close",
                                       NULL, &dispparams);

  log_oom_extra ("%s:%s: Returned from close",
                 SRCNAME, __func__);
  return rc;
}

void
Mail::set_close_triggered (bool value)
{
  m_close_triggered = value;
}

bool
Mail::get_close_triggered () const
{
  return m_close_triggered;
}

static const UserID
get_uid_for_sender (const Key &k, const char *sender)
{
  UserID ret;

  if (!sender)
    {
      return ret;
    }

  if (!k.numUserIDs())
    {
      log_debug ("%s:%s: Key without uids",
                 SRCNAME, __func__);
      return ret;
    }

  for (const auto uid: k.userIDs())
    {
      if (!uid.email())
        {
          log_error ("%s:%s: skipping uid without email.",
                     SRCNAME, __func__);
          continue;
        }
      auto normalized_uid = uid.addrSpec();
      auto normalized_sender = UserID::addrSpecFromString(sender);

      if (normalized_sender.empty() || normalized_uid.empty())
        {
          log_error ("%s:%s: normalizing '%s' or '%s' failed.",
                     SRCNAME, __func__, uid.email(), sender);
          continue;
        }
      if (normalized_sender == normalized_uid)
        {
          ret = uid;
        }
    }
  return ret;
}

void
Mail::update_sigstate ()
{
  std::string sender = get_sender();

  if (sender.empty())
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      return;
    }

  if (m_verify_result.isNull())
    {
      log_debug ("%s:%s: No verify result.",
                 SRCNAME, __func__);
      return;
    }

  if (m_verify_result.error())
    {
      log_debug ("%s:%s: verify error.",
                 SRCNAME, __func__);
      return;
    }

  for (const auto sig: m_verify_result.signatures())
    {
      m_is_signed = true;
      m_uid = get_uid_for_sender (sig.key(), sender.c_str());
      if (m_uid.isNull() || (sig.validity() != Signature::Validity::Marginal &&
          sig.validity() != Signature::Validity::Full &&
          sig.validity() != Signature::Validity::Ultimate))
        {
          /* For our category we only care about trusted sigs. And
          the UID needs to match.*/
          continue;
        }
      if (sig.validity() == Signature::Validity::Marginal)
        {
          const auto tofu = m_uid.tofuInfo();
          if (!tofu.isNull() &&
              (tofu.validity() != TofuInfo::Validity::BasicHistory &&
               tofu.validity() != TofuInfo::Validity::LargeHistory))
            {
              /* Marginal is only good enough without tofu.
                 Otherwise we wait for basic trust. */
              log_debug ("%s:%s: Discarding marginal signature."
                         "With too little history.",
                         SRCNAME, __func__);
              continue;
            }
        }
      log_debug ("%s:%s: Classified sender as verified uid validity: %i",
                 SRCNAME, __func__, m_uid.validity());
      m_sig = sig;
      m_is_valid = true;
      return;
    }

  log_debug ("%s:%s: No signature with enough trust. Using first",
             SRCNAME, __func__);
  m_sig = m_verify_result.signature(0);
  return;
}

bool
Mail::is_valid_sig ()
{
   return m_is_valid;
}

void
Mail::remove_categories ()
{
  const char *decCategory = _("GpgOL: Encrypted Message");
  const char *verifyCategory = _("GpgOL: Trusted Sender Address");
  remove_category (m_mailitem, decCategory);
  remove_category (m_mailitem, verifyCategory);
}

/* Now for some tasty hack: Outlook sometimes does
   not show the new categories properly but instead
   does some weird scrollbar thing. This can be
   avoided by resizing the message a bit. But somehow
   this only needs to be done once.

   Weird isn't it? But as this workaround worked let's
   do it programatically. Fun. Wan't some tomato sauce
   with this hack? */
static void
resize_active_window ()
{
  HWND wnd = get_active_hwnd ();
  static std::vector<HWND> resized_windows;
  if(std::find(resized_windows.begin(), resized_windows.end(), wnd) != resized_windows.end()) {
      /* We only need to do this once per window. XXX But sometimes we also
         need to do this once per view of the explorer. So for now this might
         break but we reduce the flicker. A better solution would be to find
         the current view and track that. */
      return;
  }

  if (!wnd)
    {
      TRACEPOINT;
      return;
    }
  RECT oldpos;
  if (!GetWindowRect (wnd, &oldpos))
    {
      TRACEPOINT;
      return;
    }

  if (!SetWindowPos (wnd, nullptr,
                     (int)oldpos.left,
                     (int)oldpos.top,
                     /* Anything smaller then 19 was ignored when the window was
                      * maximized on Windows 10 at least with a 1980*1024
                      * resolution. So I assume it's at least 1 percent.
                      * This is all hackish and ugly but should work for 90%...
                      * hopefully.
                      */
                     (int)oldpos.right - oldpos.left - 20,
                     (int)oldpos.bottom - oldpos.top, 0))
    {
      TRACEPOINT;
      return;
    }

  if (!SetWindowPos (wnd, nullptr,
                     (int)oldpos.left,
                     (int)oldpos.top,
                     (int)oldpos.right - oldpos.left,
                     (int)oldpos.bottom - oldpos.top, 0))
    {
      TRACEPOINT;
      return;
    }
  resized_windows.push_back(wnd);
}

void
Mail::update_categories ()
{
  const char *decCategory = _("GpgOL: Encrypted Message");
  const char *verifyCategory = _("GpgOL: Trusted Sender Address");
  if (is_valid_sig())
    {
      add_category (m_mailitem, verifyCategory);
    }
  else
    {
      remove_category (m_mailitem, verifyCategory);
    }

  if (!m_decrypt_result.isNull())
    {
      add_category (m_mailitem, decCategory);
    }
  else
    {
      /* As a small safeguard against fakes we remove our
         categories */
      remove_category (m_mailitem, decCategory);
    }

  resize_active_window();

  return;
}

bool
Mail::is_signed() const
{
  return m_verify_result.numSignatures() > 0;
}

bool
Mail::is_encrypted() const
{
  return !m_decrypt_result.isNull();
}

int
Mail::set_uuid()
{
  char *uuid;
  if (!m_uuid.empty())
    {
      /* This codepath is reached by decrypt again after a
         close with discard changes. The close discarded
         the uuid on the OOM object so we have to set
         it again. */
      log_debug ("%s:%s: Resetting uuid for %p to %s",
                 SRCNAME, __func__, this,
                 m_uuid.c_str());
      uuid = get_unique_id (m_mailitem, 1, m_uuid.c_str());
    }
  else
    {
      uuid = get_unique_id (m_mailitem, 1, nullptr);
      log_debug ("%s:%s: uuid for %p set to %s",
                 SRCNAME, __func__, this, uuid);
    }

  if (!uuid)
    {
      log_debug ("%s:%s: Failed to get/set uuid for %p",
                 SRCNAME, __func__, m_mailitem);
      return -1;
    }
  if (m_uuid.empty())
    {
      m_uuid = uuid;
      Mail *other = get_mail_for_uuid (uuid);
      if (other)
        {
          /* According to documentation this should not
             happen as this means that multiple ItemLoad
             events occured for the same mailobject without
             unload / destruction of the mail.

             But it happens. If you invalidate the UI
             in the selection change event Outlook loads a
             new mailobject for the mail. Might happen in
             other surprising cases. We replace in that
             case as experiments have shown that the last
             mailobject is the one that is visible.

             Still troubling state so we log this as an error.
             */
          log_error ("%s:%s: There is another mail for %p "
                     "with uuid: %s replacing it.",
                     SRCNAME, __func__, m_mailitem, uuid);
          delete other;
        }
      s_uid_map.insert (std::pair<std::string, Mail *> (m_uuid, this));
      log_debug ("%s:%s: uuid for %p is now %s",
                 SRCNAME, __func__, this,
                 m_uuid.c_str());
    }
  xfree (uuid);
  return 0;
}

/* Returns 2 if the userid is ultimately trusted.

   Returns 1 if the userid is fully trusted but has
   a signature by a key for which we have a secret
   and which is ultimately trusted. (Direct trust)

   0 otherwise */
static int
level_4_check (const UserID &uid)
{
  if (uid.isNull())
    {
      return 0;
    }
  if (uid.validity () == UserID::Validity::Ultimate)
    {
      return 2;
    }
  if (uid.validity () == UserID::Validity::Full)
    {
      for (const auto sig: uid.signatures ())
        {
          const char *sigID = sig.signerKeyID ();
          if (sig.isNull() || !sigID)
            {
              /* should not happen */
              TRACEPOINT;
              continue;
            }
          /* Direct trust information is not available
             through gnupg so we cached the keys with ultimate
             trust during parsing and now see if we find a direct
             trust path.*/
          for (const auto secKey: ParseController::get_ultimate_keys ())
            {
              /* Check that the Key id of the key matches */
              const char *secKeyID = secKey.keyID ();
              if (!secKeyID || strcmp (secKeyID, sigID))
                {
                  continue;
                }
              /* Check that the userID of the signature is the ultimately
                 trusted one. */
              const char *sig_uid_str = sig.signerUserID();
              if (!sig_uid_str)
                {
                  /* should not happen */
                  TRACEPOINT;
                  continue;
                }
              for (const auto signer_uid: secKey.userIDs ())
                {
                  if (signer_uid.validity() != UserID::Validity::Ultimate)
                    {
                      TRACEPOINT;
                      continue;
                    }
                  const char *signer_uid_str = signer_uid.id ();
                  if (!sig_uid_str)
                    {
                      /* should not happen */
                      TRACEPOINT;
                      continue;
                    }
                  if (!strcmp(sig_uid_str, signer_uid_str))
                    {
                      /* We have a match */
                      log_debug ("%s:%s: classified %s as ultimate because "
                                 "it was signed by uid %s of key %s",
                                 SRCNAME, __func__, signer_uid_str, sig_uid_str,
                                 secKeyID);
                      return 1;
                    }
                }
            }
        }
    }
  return 0;
}

std::string
Mail::get_crypto_summary ()
{
  const int level = get_signature_level ();

  bool enc = is_encrypted ();
  if (level == 4 && enc)
    {
      return _("Security Level 4");
    }
  if (level == 4)
    {
      return _("Trust Level 4");
    }
  if (level == 3 && enc)
    {
      return _("Security Level 3");
    }
  if (level == 3)
    {
      return _("Trust Level 3");
    }
  if (level == 2 && enc)
    {
      return _("Security Level 2");
    }
  if (level == 2)
    {
      return _("Trust Level 2");
    }
  if (enc)
    {
      return _("Encrypted");
    }
  if (is_signed ())
    {
      /* Even if it is signed, if it is not validly
         signed it's still completly insecure as anyone
         could have signed this. So we avoid the label
         "signed" here as this word already implies some
         security. */
      return _("Insecure");
    }
  return _("Insecure");
}

std::string
Mail::get_crypto_one_line()
{
  bool sig = is_signed ();
  bool enc = is_encrypted ();
  if (sig || enc)
    {
      if (sig && enc)
        {
          return _("Signed and encrypted message");
        }
      else if (sig)
        {
          return _("Signed message");
        }
      else if (enc)
        {
          return _("Encrypted message");
        }
    }
  return _("Insecure message");
}

std::string
Mail::get_crypto_details()
{
  std::string message;

  /* No signature with keys but error */
  if (!is_encrypted() && !is_signed () && m_verify_result.error())
    {
      message = _("You cannot be sure who sent, "
                  "modified and read the message in transit.");
      message += "\n\n";
      message += _("The message was signed but the verification failed with:");
      message += "\n";
      message += m_verify_result.error().asString();
      return message;
    }
  /* No crypo, what are we doing here? */
  if (!is_encrypted () && !is_signed ())
    {
      return _("You cannot be sure who sent, "
               "modified and read the message in transit.");
    }
  /* Handle encrypt only */
  if (is_encrypted() && !is_signed ())
    {
      if (in_de_vs_mode ())
       {
         if (m_sig.isDeVs())
           {
             message += _("The encryption was VS-NfD-compliant.");
           }
         else
           {
             message += _("The encryption was not VS-NfD-compliant.");
           }
        }
      message += "\n\n";
      message += _("You cannot be sure who sent the message because "
                   "it is not signed.");
      return message;
    }

  bool keyFound = true;
  bool isOpenPGP = m_sig.key().isNull() ? !is_smime() :
                   m_sig.key().protocol() == Protocol::OpenPGP;
  char *buf;
  bool hasConflict = false;
  int level = get_signature_level ();

  log_debug ("%s:%s: Formatting sig. Validity: %x Summary: %x Level: %i",
             SRCNAME, __func__, m_sig.validity(), m_sig.summary(),
             level);

  if (level == 4)
    {
      /* level 4 check for direct trust */
      int four_check = level_4_check (m_uid);

      if (four_check == 2 && m_sig.key().hasSecret ())
        {
          message = _("You signed this message.");
        }
      else if (four_check == 1)
        {
          message = _("The senders identity was certified by yourself.");
        }
      else if (four_check == 2)
        {
          message = _("The sender is allowed to certify identities for you.");
        }
      else
        {
          log_error ("%s:%s:%i BUG: Invalid sigstate.",
                     SRCNAME, __func__, __LINE__);
          return message;
        }
    }
  else if (level == 3 && isOpenPGP)
    {
      /* Level three is only reachable through web of trust and no
         direct signature. */
      message = _("The senders identity was certified by several trusted people.");
    }
  else if (level == 3 && !isOpenPGP)
    {
      /* Level three is the only level for trusted S/MIME keys. */
      gpgrt_asprintf (&buf, _("The senders identity is certified by the trusted issuer:\n'%s'\n"),
                      m_sig.key().issuerName());
      message = buf;
      xfree (buf);
    }
  else if (level == 2 && m_uid.tofuInfo ().isNull ())
    {
      /* Marginal trust through pgp only */
      message = _("Some trusted people "
                  "have certified the senders identity.");
    }
  else if (level == 2)
    {
      unsigned long first_contact = std::max (m_uid.tofuInfo().signFirst(),
                                              m_uid.tofuInfo().encrFirst());
      char *time = format_date_from_gpgme (first_contact);
      /* i18n note signcount is always pulral because with signcount 1 we
       * would not be in this branch. */
      gpgrt_asprintf (&buf, _("The senders address is trusted, because "
                              "you have established a communication history "
                              "with this address starting on %s.\n"
                              "You encrypted %i and verified %i messages since."),
                              time, m_uid.tofuInfo().encrCount(),
                              m_uid.tofuInfo().signCount ());
      xfree (time);
      message = buf;
      xfree (buf);
    }
  else if (level == 1)
    {
      /* This could be marginal trust through pgp, or tofu with little
         history. */
      if (m_uid.tofuInfo ().signCount() == 1)
        {
          message += _("The senders signature was verified for the first time.");
        }
      else if (m_uid.tofuInfo ().validity() == TofuInfo::Validity::LittleHistory)
        {
          unsigned long first_contact = std::max (m_uid.tofuInfo().signFirst(),
                                                  m_uid.tofuInfo().encrFirst());
          char *time = format_date_from_gpgme (first_contact);
          gpgrt_asprintf (&buf, _("The senders address is not trustworthy yet because "
                                  "you only verified %i messages and encrypted %i messages to "
                                  "it since %s."),
                                  m_uid.tofuInfo().signCount (),
                                  m_uid.tofuInfo().encrCount (), time);
          xfree (time);
          message = buf;
          xfree (buf);
        }
    }
  else
    {
      /* Now we are in level 0, this could be a technical problem, no key
         or just unkown. */
      message = is_encrypted () ? _("But the sender address is not trustworthy because:") :
                                  _("The sender address is not trustworthy because:");
      message += "\n";
      keyFound = !(m_sig.summary() & Signature::Summary::KeyMissing);

      bool general_problem = true;
      /* First the general stuff. */
      if (m_sig.summary() & Signature::Summary::Red)
        {
          message += _("The signature is invalid: \n");
        }
      else if (m_sig.summary() & Signature::Summary::SysError ||
               m_verify_result.numSignatures() < 1)
        {
          message += _("There was an error verifying the signature.\n");
        }
      else if (m_sig.summary() & Signature::Summary::SigExpired)
        {
          message += _("The signature is expired.\n");
        }
      else
        {
          message += isOpenPGP ? _("The used key") : _("The used certificate");
          message += " ";
          general_problem = false;
        }

      /* Now the key problems */
      if ((m_sig.summary() & Signature::Summary::KeyMissing))
        {
          message += _("is not available.");
        }
      else if ((m_sig.summary() & Signature::Summary::KeyRevoked))
        {
          message += _("is revoked.");
        }
      else if ((m_sig.summary() & Signature::Summary::KeyExpired))
        {
          message += _("is expired.");
        }
      else if ((m_sig.summary() & Signature::Summary::BadPolicy))
        {
          message += _("is not meant for signing.");
        }
      else if ((m_sig.summary() & Signature::Summary::CrlMissing))
        {
          message += _("could not be checked for revocation.");
        }
      else if ((m_sig.summary() & Signature::Summary::CrlTooOld))
        {
          message += _("could not be checked for revocation.");
        }
      else if ((m_sig.summary() & Signature::Summary::TofuConflict) ||
               m_uid.tofuInfo().validity() == TofuInfo::Conflict)
        {
          message += _("is not the same as the key that was used "
                       "for this address in the past.");
          hasConflict = true;
        }
      else if (m_uid.isNull())
        {
          gpgrt_asprintf (&buf, _("does not claim the address: \"%s\"."),
                          get_sender().c_str());
          message += buf;
          xfree (buf);
        }
      else if (((m_sig.validity() & Signature::Validity::Undefined) ||
               (m_sig.validity() & Signature::Validity::Unknown) ||
               (m_sig.summary() == Signature::Summary::None) ||
               (m_sig.validity() == 0))&& !general_problem)
        {
           /* Bit of a catch all for weird results. */
          if (isOpenPGP)
            {
              message += _("is not certified by any trustworthy key.");
            }
          else
            {
              message += _("is not certified by a trustworthy Certificate Authority or the Certificate Authority is unknown.");
            }
        }
      else if (m_uid.isRevoked())
        {
          message += _("The sender marked this address as revoked.");
        }
      else if ((m_sig.validity() & Signature::Validity::Never))
        {
          message += _("is marked as not trustworthy.");
        }
    }
   message += "\n\n";
   if (in_de_vs_mode ())
     {
       if (is_signed ())
         {
           if (m_sig.isDeVs ())
             {
               message += _("The signature is VS-NfD-compliant.");
             }
           else
             {
               message += _("The signature is not VS-NfD-compliant.");
             }
           message += "\n";
         }
       if (is_encrypted ())
         {
           if (m_decrypt_result.isDeVs ())
             {
               message += _("The encryption is VS-NfD-compliant.");
             }
           else
             {
               message += _("The encryption is not VS-NfD-compliant.");
             }
           message += "\n\n";
         }
       else
         {
           message += "\n";
         }
     }
   if (hasConflict)
    {
      message += _("Click here to change the key used for this address.");
    }
  else if (keyFound)
    {
      message +=  isOpenPGP ? _("Click here for details about the key.") :
                              _("Click here for details about the certificate.");
    }
  else
    {
      message +=  isOpenPGP ? _("Click here to search the key on the configured keyserver.") :
                              _("Click here to search the certificate on the configured X509 keyserver.");
    }
  return message;
}

int
Mail::get_signature_level () const
{
  if (!m_is_signed)
    {
      return 0;
    }

  if (m_uid.isNull ())
    {
      /* No m_uid matches our sender. */
      return 0;
    }

  if (m_is_valid && (m_uid.validity () == UserID::Validity::Ultimate ||
      (m_uid.validity () == UserID::Validity::Full &&
      level_4_check (m_uid))) && (!in_de_vs_mode () || m_sig.isDeVs()))
    {
      return 4;
    }
  if (m_is_valid && m_uid.validity () == UserID::Validity::Full &&
      (!in_de_vs_mode () || m_sig.isDeVs()))
    {
      return 3;
    }
  if (m_is_valid)
    {
      return 2;
    }
  if (m_sig.validity() == Signature::Validity::Marginal)
    {
      return 1;
    }
  if (m_sig.summary() & Signature::Summary::TofuConflict ||
      m_uid.tofuInfo().validity() == TofuInfo::Conflict)
    {
      return 0;
    }
  return 0;
}

int
Mail::get_crypto_icon_id () const
{
  int level = get_signature_level ();
  int offset = is_encrypted () ? ENCRYPT_ICON_OFFSET : 0;
  return IDI_LEVEL_0 + level + offset;
}

const char*
Mail::get_sig_fpr() const
{
  if (!m_is_signed || m_sig.isNull())
    {
      return nullptr;
    }
  return m_sig.fingerprint();
}

/** Try to locate the keys for all recipients */
void
Mail::locate_keys()
{
  char ** recipients = get_recipients ();
  KeyCache::instance()->startLocate (recipients);
  KeyCache::instance()->startLocate (get_sender ().c_str ());
  KeyCache::instance()->startLocateSecret (get_sender ().c_str ());
  release_cArray (recipients);
}

bool
Mail::is_html_alternative () const
{
  return m_is_html_alternative;
}

char *
Mail::take_cached_html_body ()
{
  char *ret = m_cached_html_body;
  m_cached_html_body = nullptr;
  return ret;
}

char *
Mail::take_cached_plain_body ()
{
  char *ret = m_cached_plain_body;
  m_cached_plain_body = nullptr;
  return ret;
}

int
Mail::get_crypto_flags () const
{
  return m_crypto_flags;
}

void
Mail::set_needs_encrypt (bool value)
{
  m_needs_encrypt = value;
}

bool
Mail::needs_encrypt() const
{
  return m_needs_encrypt;
}

char **
Mail::take_cached_recipients()
{
  char **ret = m_cached_recipients;
  m_cached_recipients = nullptr;
  return ret;
}

void
Mail::append_to_inline_body (const std::string &data)
{
  m_inline_body += data;
}

int
Mail::inline_body_to_body()
{
  if (!m_crypter)
    {
      log_error ("%s:%s: No crypter.",
                 SRCNAME, __func__);
      return -1;
    }

  const auto body = m_crypter->get_inline_data ();
  if (body.empty())
    {
      return 0;
    }

  int ret = put_oom_string (m_mailitem, "Body",
                            body.c_str ());
  return ret;
}

void
Mail::update_crypt_mapi()
{
  log_debug ("%s:%s: Update crypt mapi",
             SRCNAME, __func__);
  if (m_crypt_state != NeedsUpdateInMAPI)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      return;
    }
  if (!m_crypter)
    {
      if (!m_mime_data.empty())
        {
          log_debug ("%s:%s: Have override mime data creating dummy crypter",
                     SRCNAME, __func__);
          m_crypter = std::shared_ptr <CryptController> (new CryptController (this, false,
                                                                              false,
                                                                              GpgME::UnknownProtocol));
        }
      else
        {
          log_error ("%s:%s: No crypter.",
                     SRCNAME, __func__);
          m_crypt_state = NoCryptMail;
          return;
        }
    }

  if (m_crypter->update_mail_mapi ())
    {
      log_error ("%s:%s: Failed to update MAPI after crypt",
                 SRCNAME, __func__);
      m_crypt_state = NoCryptMail;
    }
  else
    {
      m_crypt_state = WantsSendMIME;
    }
  // We don't need the crypter anymore.
  reset_crypter ();
}

/** Checks in OOM if the body is either
  empty or contains the -----BEGIN tag.
  pair.first -> true if body starts with -----BEGIN
  pair.second -> true if body is empty. */
static std::pair<bool, bool>
has_crypt_or_empty_body_oom (Mail *mail)
{
  auto body = mail->get_body();
  std::pair<bool, bool> ret;
  ret.first = false;
  ret.second = false;
  ltrim (body);
  if (body.size() > 10 && !strncmp (body.c_str(), "-----BEGIN", 10))
    {
      ret.first = true;
      return ret;
    }
  if (!body.size())
    {
      ret.second = true;
    }
  return ret;
}

void
Mail::update_crypt_oom()
{
  log_debug ("%s:%s: Update crypt oom for %p",
             SRCNAME, __func__, this);
  if (m_crypt_state != NeedsUpdateInOOM)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      return;
    }

  if (should_inline_crypt ())
    {
      if (inline_body_to_body ())
        {
          log_debug ("%s:%s: Inline body to body failed %p.",
                     SRCNAME, __func__, this);
        }
    }

  const auto pair = has_crypt_or_empty_body_oom (this);
  if (pair.first)
    {
      log_debug ("%s:%s: Looks like inline body. You can pass %p.",
                 SRCNAME, __func__, this);
      m_crypt_state = WantsSendInline;
      return;
    }

  // We are in MIME land. Wipe the body.
  if (wipe (true))
    {
      log_debug ("%s:%s: Cancel send for %p.",
                 SRCNAME, __func__, this);
      wchar_t *title = utf8_to_wchar (_("GpgOL: Encryption not possible!"));
      wchar_t *msg = utf8_to_wchar (_(
                                      "Outlook returned an error when trying to send the encrypted mail.\n\n"
                                      "Please restart Outlook and try again.\n\n"
                                      "If it still fails consider using an encrypted attachment or\n"
                                      "switching to PGP/Inline in GpgOL's options."));
      MessageBoxW (get_active_hwnd(), msg, title,
                   MB_ICONERROR | MB_OK);
      xfree (msg);
      xfree (title);
      m_crypt_state = NoCryptMail;
      return;
    }
  m_crypt_state = NeedsSecondAfterWrite;
  return;
}

void
Mail::set_window_enabled (bool value)
{
  if (!value)
    {
      m_window = get_active_hwnd ();
    }
  log_debug ("%s:%s: enable window %p %i",
             SRCNAME, __func__, m_window, value);

  EnableWindow (m_window, value ? TRUE : FALSE);
}

bool
Mail::check_inline_response ()
{
/* Async sending might lead to crashes when the send invocation is done.
 * For now we treat every mail as an inline response to disable async
 * encryption. :-( For more details see: T3838 */
#ifdef DO_ASYNC_CRYPTO
  m_is_inline_response = false;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();
  if (!app)
    {
      TRACEPOINT;
      return false;
    }

  LPDISPATCH explorer = get_oom_object (app, "ActiveExplorer");

  if (!explorer)
    {
      TRACEPOINT;
      return false;
    }

  LPDISPATCH inlineResponse = get_oom_object (explorer, "ActiveInlineResponse");
  gpgol_release (explorer);

  if (!inlineResponse)
    {
      return false;
    }

  // We have inline response
  // Check if we are it. It's a bit naive but meh. Worst case
  // is that we think inline response too often and do sync
  // crypt where we could do async crypt.
  char * inlineSubject = get_oom_string (inlineResponse, "Subject");
  gpgol_release (inlineResponse);

  const auto subject = get_subject ();
  if (inlineResponse && !subject.empty() && !strcmp (subject.c_str (), inlineSubject))
    {
      log_debug ("%s:%s: Detected inline response for '%p'",
                 SRCNAME, __func__, this);
      m_is_inline_response = true;
    }
  xfree (inlineSubject);
#else
  m_is_inline_response = true;
#endif

  return m_is_inline_response;
}

// static
Mail *
Mail::get_last_mail ()
{
  if (!s_last_mail || !is_valid_ptr (s_last_mail))
    {
      s_last_mail = nullptr;
    }
  return s_last_mail;
}

// static
void
Mail::invalidate_last_mail ()
{
  s_last_mail = nullptr;
}

// static
void
Mail::locate_all_crypto_recipients()
{
  if (!opt.autoresolve)
    {
      return;
    }

  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = s_mail_map.begin(); it != s_mail_map.end(); ++it)
    {
      if (it->second->needs_crypto ())
        {
          it->second->locate_keys ();
        }
    }
}

int
Mail::remove_our_attachments ()
{
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      TRACEPOINT;
      return 0;
    }
  int count = get_oom_int (attachments, "Count");
  LPDISPATCH to_delete[count];
  int del_cnt = 0;
  for (int i = 1; i <= count; i++)
    {
      auto item_str = std::string("Item(") + std::to_string (i) + ")";
      LPDISPATCH attachment = get_oom_object (attachments, item_str.c_str());
      if (!attachment)
        {
          TRACEPOINT;
          continue;
        }

      attachtype_t att_type;
      if (get_pa_int (attachment, GPGOL_ATTACHTYPE_DASL, (int*) &att_type))
        {
          /* Not our attachment. */
          gpgol_release (attachment);
          continue;
        }

      if (att_type == ATTACHTYPE_PGPBODY || att_type == ATTACHTYPE_MOSS ||
          att_type == ATTACHTYPE_MOSSTEMPL)
        {
          /* One of ours to delete. */
          to_delete[del_cnt++] = attachment;
          /* Dont' release yet */
          continue;
        }
      gpgol_release (attachment);
    }
  gpgol_release (attachments);

  int ret = 0;

  for (int i = 0; i < del_cnt; i++)
    {
      LPDISPATCH attachment = to_delete[i];

      /* Delete the attachments that are marked to delete */
      if (invoke_oom_method (attachment, "Delete", NULL))
        {
          log_error ("%s:%s: Error: deleting attachment %i",
                     SRCNAME, __func__, i);
          ret = -1;
        }
      gpgol_release (attachment);
    }
  return ret;
}

/* We are very verbose because if we fail it might mean
   that we have leaked plaintext -> critical. */
bool
Mail::has_crypted_or_empty_body ()
{
  const auto pair = has_crypt_or_empty_body_oom (this);

  if (pair.first /* encrypted marker */)
    {
      log_debug ("%s:%s: Crypt Marker detected in OOM body. Return true %p.",
                 SRCNAME, __func__, this);
      return true;
    }

  if (!pair.second)
    {
      log_debug ("%s:%s: Unexpected content detected. Return false %p.",
                 SRCNAME, __func__, this);
      return false;
    }

  // Pair second == true (is empty) can happen on OOM error.
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message && pair.second)
    {
      if (message)
        {
          gpgol_release (message);
        }
      return true;
    }

  size_t r_nbytes = 0;
  char *mapi_body = mapi_get_body (message, &r_nbytes);
  gpgol_release (message);

  if (!mapi_body || !r_nbytes)
    {
      // Body or bytes are null. we are empty.
      xfree (mapi_body);
      log_debug ("%s:%s: MAPI error or empty message. Return true. %p.",
                 SRCNAME, __func__, this);
      return true;
    }
  if (r_nbytes > 10 && !strncmp (mapi_body, "-----BEGIN", 10))
    {
      // Body is crypt.
      log_debug ("%s:%s: MAPI Crypt marker detected. Return true. %p.",
                 SRCNAME, __func__, this);
      xfree (mapi_body);
      return true;
    }

  xfree (mapi_body);

  log_debug ("%s:%s: Found mapi body. Return false. %p.",
             SRCNAME, __func__, this);

  return false;
}
