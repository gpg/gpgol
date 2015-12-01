/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 *    Copyright (C) 2015 Intevation GmbH
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

#include <map>

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

Mail::Mail (LPDISPATCH mailitem) :
    m_mailitem(mailitem),
    m_processed(false),
    m_needs_wipe(false),
    m_crypt_successful(false),
    m_sender(NULL)
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
      mailitem->Release ();
      return;
    }
  g_mail_map.insert (std::pair<LPDISPATCH, Mail *> (mailitem, this));
}

int
Mail::process_message ()
{
  int err;
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
  err = message_incoming_handler (message, NULL,
                                  false);
  m_processed = (err == 1) || (err == 2);

  log_debug ("%s:%s: incoming handler status: %i",
             SRCNAME, __func__, err);
  message->Release ();
  return 0;
}

Mail::~Mail()
{
  std::map<LPDISPATCH, Mail *>::iterator it;

  detach_MailItemEvents_sink (m_event_sink);
  m_event_sink->Release ();

  it = g_mail_map.find(m_mailitem);
  if (it != g_mail_map.end())
    {
      g_mail_map.erase (it);
    }

  m_mailitem->Release ();
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
Mail::insert_plaintext ()
{
  int err = 0;
  int is_html, was_protected = 0;
  char *body = NULL;

  if (!m_processed)
    {
      return 0;
    }

  if (m_needs_wipe)
    {
      log_error ("%s:%s: Insert plaintext called for msg that needs wipe: %p",
                 SRCNAME, __func__, m_mailitem);
      return 0;
    }

  /* Outlook somehow is confused about the attachment
     table of our sent mails. The securemessage interface
     gives us access to the real attach table but the attachment
     table of the message itself is broken. */
  LPMESSAGE base_message = get_oom_base_message (m_mailitem);
  if (!base_message)
    {
      log_error ("%s:%s: Failed to get base message",
                 SRCNAME, __func__);
      return 0;
    }
  err = mapi_get_gpgol_body_attachment (base_message, &body, NULL,
                                        &is_html, &was_protected);
  m_needs_wipe = was_protected;

  log_debug ("%s:%s: Setting plaintext for msg: %p",
             SRCNAME, __func__, m_mailitem);
  if (err || !body)
    {
      log_error ("%s:%s: Failed to get body attachment. Err: %i",
                 SRCNAME, __func__, err);
      put_oom_string (m_mailitem, "HTMLBody", HTML_TEMPLATE);
      err = -1;
      goto done;
    }
  if (put_oom_string (m_mailitem, is_html ? "HTMLBody" : "Body", body))
    {
      log_error ("%s:%s: Failed to modify body of item.",
                 SRCNAME, __func__);
      err = -1;
    }

  xfree (body);
  /* TODO: unprotect attachments does not work for sent mails
     as the attachment table of the mapiitem is invalid.
     We need to somehow get outlook to use the attachment table
     of the base message and and then decrypt those.
     This will probably mean removing all attachments for the
     message and adding the attachments from the base message then
     we can call unprotect_attachments as usual. */
  if (unprotect_attachments (m_mailitem))
    {
      log_error ("%s:%s: Failed to unprotect attachments.",
                 SRCNAME, __func__);
      err = -1;
    }

done:
  RELDISP (base_message);
  return err;
}

int
Mail::do_crypto ()
{
  int err = -1,
      flags = 0;
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
      err = message_sign_encrypt (message, PROTOCOL_UNKNOWN,
                                  NULL, get_sender ());
    }
  else if (flags == 2)
    {
      err = message_sign (message, PROTOCOL_UNKNOWN,
                          NULL, get_sender ());
    }
  else if (flags == 1)
    {
      err = message_encrypt (message, PROTOCOL_UNKNOWN,
                             NULL, get_sender ());
    }
  else
    {
      log_debug ("%s:%s: Unknown flags for crypto: %i",
                 SRCNAME, __func__, flags);
    }
  log_debug ("%s:%s: Status: %i",
             SRCNAME, __func__, err);
  message->Release ();
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
  message->Release ();
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
                      HTML_TEMPLATE) ||
      protect_attachments (m_mailitem))
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
  sender = get_oom_object (m_mailitem, "Session.CurrentUser");

  xfree (m_sender);

  if (!sender)
    {
      log_error ("%s:%s: Failed to get sender object.",
                 SRCNAME, __func__);
      return -1;
    }
  m_sender = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);

  if (!m_sender)
    {
      log_error ("%s:%s: Failed to get smtp address of sender.",
                 SRCNAME, __func__);
      return -1;
    }
  return 0;
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
