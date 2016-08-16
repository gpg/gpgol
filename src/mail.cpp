/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 *    Copyright (C) 2015, 2016 Intevation GmbH
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
#include "common.h"
#include "mail.h"
#include "eventsinks.h"
#include "attachment.h"
#include "mapihelp.h"
#include "message.h"
#include "revert.h"
#include "gpgoladdin.h"
#include "mymapitags.h"
#include "mailparser.h"
#include "gpgolstr.h"

#include <map>
#include <vector>
#include <memory>

static std::map<LPDISPATCH, Mail*> g_mail_map;

/* TODO: Localize this once it is less bound to change.
   TODO: Use a dedicated message for failed decryption. */
#define HTML_TEMPLATE  \
"<html><head></head><body>" \
"<table border=\"0\" width=\"100%\" cellspacing=\"1\" cellpadding=\"1\" bgcolor=\"#0069cc\">" \
"<tr>" \
"<td bgcolor=\"#0080ff\">" \
"<p><span style=\"font-weight:600; background-color:#0080ff;\"><center>This message is encrypted</center><span></p></td></tr>" \
"<tr>" \
"<td bgcolor=\"#e0f0ff\">" \
"<center>" \
"<br/>You can decrypt this message with GnuPG" \
"<br/>Open this message to decrypt it." \
"<br/>Opening any attachments while this message is shown will only give you access to encrypted data. </center>" \
"<br/><br/>If you have GpgOL (The GnuPG Outlook plugin installed) this message should have been automatically decrypted." \
"<br/>Reasons that you still see this message can be: " \
"<ul>" \
"<li>Decryption failed: <ul><li> Refer to the Decrypt / Verify popup window for details.</li></ul></li>" \
"<li>Outlook tried to save the decrypted content:" \
" <ul> "\
" <li>To protect your data GpgOL encrypts a message when it is saved by Outlook.</li>" \
" <li>You will need to restart Outlook to allow GpgOL to decrypt this message again.</li>" \
" </ul>" \
"<li>GpgOL is not activated: <ul><li>Check under Options -> Add-Ins -> COM-Add-Ins to see if this is the case.</li></ul></li>" \
"</ul>"\
"</td></tr>" \
"</table></body></html>"

#define WAIT_TEMPLATE \
"<html><head></head><body>" \
"<table border=\"0\" width=\"100%\" cellspacing=\"1\" cellpadding=\"1\" bgcolor=\"#0069cc\">" \
"<tr>" \
"<td bgcolor=\"#0080ff\">" \
"<p><span style=\"font-weight:600; background-color:#0080ff;\"><center>This message is encrypted</center><span></p></td></tr>" \
"<tr>" \
"<td bgcolor=\"#e0f0ff\">" \
"<center>" \
"<br/>Please wait while the message is decrypted by GpgOL..." \
"</td></tr>" \
"</table></body></html>"

Mail::Mail (LPDISPATCH mailitem) :
    m_mailitem(mailitem),
    m_processed(false),
    m_needs_wipe(false),
    m_needs_save(false),
    m_crypt_successful(false),
    m_is_smime(false),
    m_is_smime_checked(false),
    m_sender(NULL),
    m_type(MSGTYPE_UNKNOWN)
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
  g_mail_map.insert (std::pair<LPDISPATCH, Mail *> (mailitem, this));
}

Mail::~Mail()
{
  std::map<LPDISPATCH, Mail *>::iterator it;

  detach_MailItemEvents_sink (m_event_sink);
  gpgol_release(m_event_sink);

  it = g_mail_map.find(m_mailitem);
  if (it != g_mail_map.end())
    {
      g_mail_map.erase (it);
    }

  gpgol_release(m_mailitem);
}

Mail *
Mail::get_mail_for_item (LPDISPATCH mailitem)
{
  if (!mailitem)
    {
      return NULL;
    }
  std::map<LPDISPATCH, Mail *>::iterator it;
  it = g_mail_map.find(mailitem);
  if (it == g_mail_map.end())
    {
      return NULL;
    }
  return it->second;
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
  log_oom_extra ("%s:%s: GetBaseMessage OK.",
                 SRCNAME, __func__);
  /* Change the message class here. It is important that
     we change the message class in the before read event
     regardless if it is already set to one of GpgOL's message
     classes. Changing the message class (even if we set it
     to the same value again that it already has) causes
     Outlook to reconsider what it "knows" about a message
     and reread data from the underlying base message. */
  mapi_change_message_class (message, 1);
  /* TODO: Unify this so mapi_change_message_class returns
     a useful value already. */
  m_type = mapi_get_message_type (message);

  /* Create moss attachments here so that they are properly
     hidden when the item is read into the model. */
  if (mapi_mark_or_create_moss_attach (message, m_type))
    {
      log_error ("%s:%s: Failed to find moss attachment.",
                 SRCNAME, __func__);
      m_type = MSGTYPE_UNKNOWN;
    }

  gpgol_release (message);
  return 0;
}

/** Get the cipherstream of the mailitem. */
static LPSTREAM
get_cipherstream (LPDISPATCH mailitem)
{
  LPDISPATCH attachments = get_oom_object (mailitem, "Attachments");
  LPDISPATCH attachment = NULL;
  LPATTACH mapi_attachment = NULL;
  LPSTREAM stream = NULL;

  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      return NULL;
    }

  int count = get_oom_int (attachments, "Count");
  if (count < 1)
    {
      log_debug ("%s:%s: Invalid attachment count: %i.",
                 SRCNAME, __func__, count);
      gpgol_release (attachments);
      return NULL;
    }
  if (count > 1)
    {
      log_debug ("%s:%s: More then one attachment count: %i. Continuing anway.",
                 SRCNAME, __func__, count);
    }
  /* We assume the crypto attachment is the second item. */
  attachment = get_oom_object (attachments, "Item(2)");
  gpgol_release (attachments);
  attachments = NULL;

  mapi_attachment = (LPATTACH) get_oom_iunknown (attachment,
                                                 "MapiObject");
  gpgol_release (attachment);
  if (!mapi_attachment)
    {
      log_debug ("%s:%s: Failed to get MapiObject of attachment: %p",
                 SRCNAME, __func__, attachment);
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

/** Helper to check if the body should be used for decrypt verify
  or if the mime Attachment should be used. */
static bool
use_body(msgtype_t type)
{
  switch (type)
    {
      case MSGTYPE_GPGOL_PGP_MESSAGE:
      case MSGTYPE_GPGOL_CLEAR_SIGNED:
        return true;
      default:
        return false;
    }
}

/** Helper to update the attachments of a mail object in oom.
  does not modify the underlying mapi structure. */
static bool
add_attachments(LPDISPATCH mail,
                std::vector<std::shared_ptr<Attachment> > attachments)
{
  for (auto att: attachments)
    {
      wchar_t* wchar_name = utf8_to_wchar (att->get_display_name().c_str());
      log_debug("DisplayName %s", att->get_display_name().c_str());
      HANDLE hFile;
      wchar_t* wchar_file = get_tmp_outfile (GpgOLStr("gpgol-attach-"), &hFile);
      if (add_oom_attachment (mail, wchar_file, wchar_name))
        {
          log_debug ("Failed to add attachment.");
        }
      CloseHandle (hFile);
      DeleteFileW (wchar_file);
      xfree (wchar_file);
      xfree (wchar_name);
    }
  return false;
}

int
Mail::decrypt_verify()
{
  if (m_type == MSGTYPE_UNKNOWN || m_type == MSGTYPE_GPGOL)
    {
      /* Not a message for us. */
      return 0;
    }
  if (m_needs_wipe)
    {
      log_error ("%s:%s: Decrypt verify called for msg that needs wipe: %p",
                 SRCNAME, __func__, m_mailitem);
      return 0;
    }

  m_processed = true;
  /* Do the actual parsing */
  std::unique_ptr<MailParser> parser;
  if (!use_body (m_type))
    {
      auto cipherstream = get_cipherstream (m_mailitem);

      if (!cipherstream)
        {
          /* TODO Error message? */
          log_debug ("%s:%s: Failed to get cipherstream.",
                     SRCNAME, __func__);
          return 1;
        }

      parser = std::unique_ptr<MailParser>(new MailParser (cipherstream, m_type));
      gpgol_release (cipherstream);
    }
  else
    {
      parser = std::unique_ptr<MailParser>();
    }

  const std::string err = parser->parse();
  if (!err.empty())
    {
      /* TODO Show error message. */
      log_error ("%s:%s: Failed to parse message: %s",
                 SRCNAME, __func__, err.c_str());
      return 1;
    }

  m_needs_wipe = true;
  /* Update the body */
  const auto html = parser->get_utf8_html_body();
  if (!html->empty())
    {
      if (put_oom_string (m_mailitem, "HTMLBody", html->c_str()))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
          return 1;
        }
    }
  else
    {
      const auto body = parser->get_utf8_text_body();
      if (put_oom_string (m_mailitem, "Body", body->c_str()))
        {
          log_error ("%s:%s: Failed to modify body of item.",
                     SRCNAME, __func__);
          return 1;
        }
    }

  /* Update attachments */
  if (add_attachments (m_mailitem, parser->get_attachments()))
    {
      log_error ("%s:%s: Failed to update attachments.",
                 SRCNAME, __func__);
      return 1;
    }

  /* Invalidate UI to set the correct sig status. */
  gpgoladdin_invalidate_ui ();
  return 0;
}

int
Mail::encrypt_sign ()
{
  int err = -1,
      flags = 0;
  protocol_t proto = opt.enable_smime ? PROTOCOL_UNKNOWN : PROTOCOL_OPENPGP;
  if (!needs_crypto())
    {
      return 0;
    }
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      return err;
    }
  flags = get_gpgol_draft_info_flags (message);
  if (flags == 3)
    {
      log_debug ("%s:%s: Sign / Encrypting message",
                 SRCNAME, __func__);
      err = message_sign_encrypt (message, proto,
                                  NULL, get_sender ());
    }
  else if (flags == 2)
    {
      err = message_sign (message, proto,
                          NULL, get_sender ());
    }
  else if (flags == 1)
    {
      err = message_encrypt (message, proto,
                             NULL, get_sender ());
    }
  else
    {
      log_debug ("%s:%s: Unknown flags for crypto: %i",
                 SRCNAME, __func__, flags);
    }
  log_debug ("%s:%s: Status: %i",
             SRCNAME, __func__, err);
  gpgol_release (message);
  m_crypt_successful = !err;
  return err;
}

bool
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
Mail::wipe ()
{
  if (!m_needs_wipe)
    {
      return 0;
    }
  log_debug ("%s:%s: Removing plaintext from mailitem: %p.",
             SRCNAME, __func__, m_mailitem);
  if (put_oom_string (m_mailitem, "HTMLBody",
                      HTML_TEMPLATE))
    {
      log_debug ("%s:%s: Failed to wipe mailitem: %p.",
                 SRCNAME, __func__, m_mailitem);
      return -1;
    }
  m_needs_wipe = false;
  return 0;
}

int
Mail::update_sender ()
{
  LPDISPATCH sender = NULL;
  sender = get_oom_object (m_mailitem, "SendUsingAccount");

  xfree (m_sender);
  m_sender = NULL;

  if (sender)
    {
      m_sender = get_oom_string (sender, "SmtpAddress");
      gpgol_release (sender);
      return 0;
    }
  /* Fallback to Sender object */
  sender = get_oom_object (m_mailitem, "Sender");
  if (sender)
    {
      m_sender = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
      gpgol_release (sender);
      return 0;
    }
  /* We don't have s sender object or SendUsingAccount,
     well, in that case fall back to the current user. */
  sender = get_oom_object (m_mailitem, "Session.CurrentUser");
  if (sender)
    {
      m_sender = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
      gpgol_release (sender);
      return 0;
    }

  log_error ("%s:%s: All fallbacks failed.",
             SRCNAME, __func__);
  return -1;
}

const char *
Mail::get_sender ()
{
  if (!m_sender)
    {
      update_sender();
    }
  return m_sender;
}

int
Mail::revert_all_mails ()
{
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = g_mail_map.begin(); it != g_mail_map.end(); ++it)
    {
      if (it->second->revert ())
        {
          log_error ("Failed to wipe mail: %p ", it->first);
          err++;
        }
    }
  return err;
}

int
Mail::wipe_all_mails ()
{
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = g_mail_map.begin(); it != g_mail_map.end(); ++it)
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
  int err;
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
          log_error ("Protocol in multipart signed mail.");
        }
      xfree (proto);
      xfree (ct);
    }
  gpgol_release (message);
  m_is_smime_checked  = true;
  return m_is_smime;
}

std::string
Mail::get_subject()
{
  return std::string(get_oom_string (m_mailitem, "Subject"));
}
