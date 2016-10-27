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
#include "dialogs.h"
#include "common.h"
#include "mail.h"
#include "eventsinks.h"
#include "attachment.h"
#include "mapihelp.h"
#include "message.h"
#include "revert.h"
#include "gpgoladdin.h"
#include "mymapitags.h"
#include "parsecontroller.h"
#include "gpgolstr.h"
#include "windowmessages.h"
#include "mlang-charset.h"
#include "mimemaker.h"

#include <gpgme++/tofuinfo.h>
#include <gpgme++/verificationresult.h>
#include <gpgme++/decryptionresult.h>
#include <gpgme++/key.h>

#include <map>
#include <vector>
#include <memory>

using namespace GpgME;

static std::map<LPDISPATCH, Mail*> g_mail_map;
static std::map<std::string, Mail*> g_uid_map;

#define COPYBUFSIZE (8 * 1024)

/* Our own basic trust level for tofu.
   GnuPG's can't be trusted. See comment
   in get_valid_sig why.*/
#define GPGOL_BASIC_TOFU_TRUST 10

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
    m_needs_save(false),
    m_crypt_successful(false),
    m_is_smime(false),
    m_is_smime_checked(false),
    m_is_signed(false),
    m_is_valid(false),
    m_moss_position(0),
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

  if (!m_uuid.empty())
    {
      auto it2 = g_uid_map.find(m_uuid);
      if (it2 != g_uid_map.end())
        {
          g_uid_map.erase (it2);
        }
    }

  xfree (m_sender);
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

Mail *
Mail::get_mail_for_uuid (const char *uuid)
{
  if (!uuid)
    {
      return NULL;
    }
  auto it = g_uid_map.find(std::string(uuid));
  if (it == g_uid_map.end())
    {
      return NULL;
    }
  return it->second;
}

bool
Mail::is_valid_ptr (const Mail *mail)
{
  auto it = g_mail_map.begin();
  while (it != g_mail_map.end())
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

  const auto item_str = std::string("Item(") + std::to_string(pos) + ")";
  int count = get_oom_int (attachments, "Count");
  if (count < 1)
    {
      log_debug ("%s:%s: Invalid attachment count: %i.",
                 SRCNAME, __func__, count);
      gpgol_release (attachments);
      return NULL;
    }
  attachment = get_oom_object (attachments, item_str.c_str());
  gpgol_release (attachments);

  return attachment;
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
  LPATTACH mapi_attachment = NULL;
  LPSTREAM stream = NULL;

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

int
Mail::remove_attachments()
{
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      return 0;
    }

  int count = get_oom_int (attachments, "Count");
  if (count < 1)
    {
      gpgol_release (attachments);
      return 0;
    }
  for (int i = 1; i <= count; i++)
    {
      VARIANT aVariant[1];
      DISPPARAMS dispparams;

      dispparams.rgvarg = aVariant;
      dispparams.rgvarg[0].vt = VT_INT;
      dispparams.rgvarg[0].intVal = i;
      dispparams.cArgs = 1;
      dispparams.cNamedArgs = 0;

      if (invoke_oom_method_with_parms (attachments, "Remove",
                                        NULL, &dispparams))
        {
          log_error ("%s:%s: Failed to remove attachments.",
                     SRCNAME, __func__);
          gpgol_release (attachments);
          return 1;
        }
    }
  gpgol_release (attachments);

  return 0;
}

/** Update the attachments of a mail object in oom.
  does not modify the underlying mapi structure. */
int
Mail::add_attachments(std::vector<std::shared_ptr<Attachment> > attachments)
{
  int err = 0;
  for (auto att: attachments)
    {
      wchar_t* wchar_name = utf8_to_wchar (att->get_display_name().c_str());
      HANDLE hFile;
      wchar_t* wchar_file = get_tmp_outfile (GpgOLStr (att->get_display_name().c_str()),
                                             &hFile);
      if (copy_attachment_to_file (att, hFile))
        {
          log_error ("%s:%s: Failed to copy attachment %s to temp file",
                     SRCNAME, __func__, att->get_display_name().c_str());
          err = 1;
        }
      if (add_oom_attachment (m_mailitem, wchar_file, wchar_name))
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
      if (err)
        {
          return err;
        }
    }
  return 0;
}

static DWORD WINAPI
do_parsing (LPVOID arg)
{
  log_debug ("%s:%s: starting the parser for: %p",
             SRCNAME, __func__, arg);

  Mail *mail = (Mail *)arg;
  auto parser = mail->parser();
  if (!parser)
    {
      log_error ("%s:%s: no parser found for mail: %p",
                 SRCNAME, __func__, arg);
      return -1;
    }
  parser->parse();
  do_in_ui_thread (PARSING_DONE, arg);
  return 0;
}

int
Mail::decrypt_verify()
{
  if (m_type == MSGTYPE_UNKNOWN || m_type == MSGTYPE_GPGOL ||
      m_type == MSGTYPE_SMIME)
    {
      /* Not a message for us. */
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
  if (gpgrt_asprintf (&placeholder_buf, decrypt_template,
                      is_smime() ? "S/MIME" : "OpenPGP",
                      _("Encrypted message"),
                      _("Please wait while the message is being decrypted / verified...")) == -1)
    {
      log_error ("%s:%s: Failed to format placeholder.",
                 SRCNAME, __func__);
      return 1;
    }

  if (put_oom_string (m_mailitem, "HTMLBody", placeholder_buf))
    {
      log_error ("%s:%s: Failed to modify html body of item.",
                 SRCNAME, __func__);
    }
  xfree (placeholder_buf);

  /* Do the actual parsing */
  auto cipherstream = get_attachment_stream (m_mailitem, m_moss_position);

  if (!cipherstream)
    {
      log_debug ("%s:%s: Failed to get cipherstream.",
                 SRCNAME, __func__);
      return 1;
    }

  m_parser = new ParseController (cipherstream, m_type);
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

void
Mail::update_body()
{
  const auto error = m_parser->get_formatted_error ();
  if (!error.empty())
    {
      if (put_oom_string (m_mailitem, "HTMLBody",
                          error.c_str ()))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
        }
      return;
    }
  const auto html = m_parser->get_html_body();
  if (!html.empty())
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
  const auto body = m_parser->get_body();
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
  /* Store the results. */
  m_decrypt_result = m_parser->decrypt_result ();
  m_verify_result = m_parser->verify_result ();

  update_sigstate ();
  m_needs_wipe = true;

  TRACEPOINT;
  /* Set categories according to the result. */
  update_categories();

  TRACEPOINT;
  /* Update the body */
  update_body();
  TRACEPOINT;

  /* Update attachments */
  if (add_attachments (m_parser->get_attachments()))
    {
      log_error ("%s:%s: Failed to update attachments.",
                 SRCNAME, __func__);
      return;
    }

  /* Invalidate UI to set the correct sig status. */
  delete m_parser;
  m_parser = nullptr;
  gpgoladdin_invalidate_ui ();
  TRACEPOINT;
  return;
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
  err = encrypt_sign_mail (this, (flags & 1), (flags & 2), proto,
                           NULL);
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

  /* For some reason outlook my store the recipient address
     in the send using account field. If we have SMTP we prefer
     the SenderEmailAddress string. */
  char *type = get_oom_string (m_mailitem, "SenderEmailType");
  if (type && !strcmp ("SMTP", type))
    {
      xfree (type);
      char *senderMail = get_oom_string (m_mailitem, "SenderEmailAddress");
      if (senderMail)
        {
          xfree (m_sender);
          m_sender = senderMail;
          return 0;
        }
    }
  xfree (type);

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
#if 0
  std::map<LPDISPATCH, Mail *>::iterator it;
  for (it = g_mail_map.begin(); it != g_mail_map.end(); ++it)
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
          log_error ("Failed to save reverted mail: %p ", it->first);
          err++;
          continue;
        }
    }
#endif
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
          log_error ("Protocol in multipart signed mail.");
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
  return get_oom_recipients (recipients);
}

int
Mail::close_inspector ()
{
  LPDISPATCH inspector = get_oom_object (m_mailitem, "GetInspector");
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
          return -1;
        }
    }
  return 0;
}

static const UserID
get_uid_for_sender (const Key k, const char *sender)
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
  const char *sender = get_sender();

  if (!sender)
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
      if (sig.validity() != Signature::Validity::Marginal &&
          sig.validity() != Signature::Validity::Full &&
          sig.validity() != Signature::Validity::Ultimate)
        {
          /* For our category we only care about trusted sigs. */
          continue;
        }
      const auto uid = get_uid_for_sender (sig.key(), sender);
      if (sig.validity() == Signature::Validity::Marginal)
        {
          const auto tofu = uid.tofuInfo();
          if (tofu.isNull() ||
              (tofu.validity() != TofuInfo::Validity::LittleHistory &&
               tofu.validity() != TofuInfo::Validity::BasicHistory &&
               tofu.validity() != TofuInfo::Validity::LargeHistory))
            {
              /* Marginal is not good enough without tofu.
                 We also wait for basic trust. */
              log_debug ("%s:%s: Discarding marginal signature.",
                         SRCNAME, __func__);
              continue;
            }
          /* GnuPG uses the encrypt count to determine validity.
             This does not make sense for us. E.g. Drafts may have
             been encrypted and encryption is no communication so
             it does not track communication history or consistency.
             So basically "our" tofu validity is that more then 10 messages
             have been exchanged. Which was the original code in GnuPG */
          if (!tofu.isNull() && tofu.signCount() <= GPGOL_BASIC_TOFU_TRUST) {
              log_debug ("%s:%s: Tofu signcount too small.",
                         SRCNAME, __func__);
              continue;
          }
        }
      log_debug ("%s:%s: Classified sender as verified",
                 SRCNAME, __func__);
      m_sig = sig;
      m_uid = uid;
      m_is_valid = true;
    }

  log_debug ("%s:%s: No signature with enough trust. Using first",
             SRCNAME, __func__);
  m_sig = m_verify_result.signature(0);
  return;
}

const std::pair <Signature, UserID>
Mail::get_valid_sig ()
{
  std::pair <Signature, UserID> ret;
  if (!m_is_valid)
    {
      return ret;
    }
  return std::pair<Signature, UserID> (m_sig, m_uid);
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
  const char *verifyCategory = _("GpgOL: Verified Sender");
  remove_category (m_mailitem, decCategory);
  remove_category (m_mailitem, verifyCategory);
}

void
Mail::update_categories ()
{
  const char *decCategory = _("GpgOL: Encrypted Message");
  const char *verifyCategory = _("GpgOL: Verified Sender");
  if (m_decrypt_result.numRecipients())
    {
      /* We use the number of recipients as we don't care
         if decryption was successful or not for this category */
      add_category (m_mailitem, decCategory);
    }
  else
    {
      /* As a small safeguard against fakes we remove our
         categories */
      remove_category (m_mailitem, decCategory);
    }

  if (is_valid_sig())
    {
      add_category (m_mailitem, verifyCategory);
    }
  else
    {
      remove_category (m_mailitem, verifyCategory);
    }
  return;
}

bool
Mail::is_signed()
{
  return m_verify_result.numSignatures() > 0;
}

int
Mail::set_uuid()
{
  if (!m_uuid.empty())
    {
      return 0;
    }
  char *uuid = get_unique_id (m_mailitem, 1);

  if (!uuid)
    {
      log_debug ("%s:%s: Failed to get uuid for %p",
                 SRCNAME, __func__, m_mailitem);
      return -1;
    }
  m_uuid = uuid;
  xfree (uuid);
  g_uid_map.insert (std::pair<std::string, Mail *> (m_uuid, this));
  return 0;
}

std::string
Mail::get_signature_status()
{
  std::string message;
  if (!is_signed())
    {
      message =_("This message is not signed.\n");
      message += _("You cannot be sure who wrote the message.");
      return message;
    }

  const auto pair = get_valid_sig ();
  bool keyFound = false;
  char *buf;
  bool isOpenPGP = pair.first.key().protocol() == Protocol::OpenPGP;
  if (!pair.first.isNull () && !pair.second.isNull ())
    {
      const auto sig = pair.first;
      const auto uid = pair.second;
      /* We are valid */
      keyFound = true;
      gpgrt_asprintf (&buf, _("The sender is verified because:\n\nThe used %s %s"),
                      isOpenPGP ? _("key") : _("certificate"),
                      sig.validity() == Signature::Validity::Ultimate ?
                      _("is marked as your own.") :
                      sig.validity() == Signature::Validity::Full && isOpenPGP ?
                      _("was certified by enough trusted keys.") :
                      "");
      message += buf;
      xfree (buf);
      if (sig.validity() == Signature::Validity::Full && !isOpenPGP)
        {
          gpgrt_asprintf (&buf, _("is cerified by the trusted issuer:\n'%s'\n"),
                          sig.key().issuerName());
          message += buf;
          xfree (buf);
        }
      else if (sig.validity() == Signature::Validity::Marginal)
        {
          char *time = format_date_from_gpgme (uid.tofuInfo().signFirst());
          gpgrt_asprintf (&buf, _("was consistently used for %i messages since %s."),
                          uid.tofuInfo().signCount (), time);
          xfree (time);
          message += buf;
          xfree (buf);
        }
    }
  else
    {
      if (m_verify_result.numSignatures() > 1)
        {
          log_debug ("%s:%s: More then one signature found on %p",
                     SRCNAME, __func__, m_mailitem);
        }
      /* We only handle the first signature. */
      const auto sig = m_verify_result.signature (0);
      isOpenPGP = sig.key().protocol() == Protocol::OpenPGP;
      keyFound = !(sig.summary() & Signature::Summary::KeyMissing);

      log_debug ("%s:%s: Formatting sig. Validity: %x Summary: %x",
                 SRCNAME, __func__, sig.validity(), sig.summary());

      /* There is a signature but we don't accepted it as fully valid. */
      message += _("The sender is not verified because:\n\n");

      /* First the general stuff. */
      if (sig.summary() & Signature::Summary::Red)
        {
          message += _("The signature is invalid.\n");
          message += _("You cannot be sure who wrote the message.");
        }
      else if (sig.summary() & Signature::Summary::SysError ||
               m_verify_result.numSignatures() < 1)
        {
          message += _("There was an error verifying the signature.\n");
          message += _("You cannot be sure who wrote the message.");
        }
      else if (sig.summary() & Signature::Summary::SigExpired)
        {
          message += _("The signature is expired.\n");
        }
      else
        {
          gpgrt_asprintf (&buf, _("The used %s "),
                          isOpenPGP ? _("key") : _("certificate"));
          message += buf;
          xfree (buf);
        }

      const auto uid = get_uid_for_sender (sig.key(), get_sender());
      /* Now the key problems */
      if ((sig.summary() & Signature::Summary::KeyMissing))
        {
          message += _("is not in your keyring.");
        }
      else if ((sig.summary() & Signature::Summary::KeyRevoked))
        {
          message += _("is revoked.");
        }
      else if ((sig.summary() & Signature::Summary::KeyExpired))
        {
          message += _("is expired.");
        }
      else if ((sig.summary() & Signature::Summary::BadPolicy))
        {
          message += _("is not meant for signing.");
        }
      else if ((sig.summary() & Signature::Summary::CrlMissing))
        {
          message += _("could not be checked for revocation.");
        }
      else if ((sig.summary() & Signature::Summary::CrlTooOld))
        {
          message += _("could not be checked for revocation.");
        }
      else if ((sig.summary() & Signature::Summary::TofuConflict))
        {
          message += _("conflicts with another key that was used in the past by the sender.");
        }
      else if (uid.isNull())
        {
          gpgrt_asprintf (&buf, _("does not claim the address: \"%s\"."),
                          get_sender());
          message += buf;
          xfree (buf);
        }
      else if ((sig.validity() & Signature::Validity::Marginal))
        {
          const auto tofuInfo = uid.tofuInfo();
          if (tofuInfo.isNull() || !tofuInfo.signCount())
            {
              message += _("is not certified by enough trusted keys.");
            }
          else
            {
              message += _("does not have enough history for basic trust.");
            }
        }
      else if ((sig.validity() & Signature::Validity::Undefined) ||
               (sig.validity() & Signature::Validity::Unknown) ||
               (sig.summary() == Signature::Summary::None) ||
               (sig.validity() == 0))
        {
           /* Bit of a catch all for weird results. */
           message += _("is not ceritified by any trusted key.");
        }
      else if ((sig.validity() & Signature::Validity::Never))
        {
          message += _("is explicitly marked as invalid.");
        }
    }
  message += "\n\n";
  if (keyFound)
    {
      gpgrt_asprintf (&buf, _("Click here for details about the %s."),
                  isOpenPGP ? _("key") : _("certificate."));
    }
  else
    {
      gpgrt_asprintf (&buf, _("Click here to search the %s on a public keyserver."),
                  isOpenPGP ? _("key") : _("certificate."));
    }
  message += buf;
  xfree (buf);
  return message;
}

int
Mail::get_signature_icon_id () const
{
  if (!m_is_signed)
    {
      return IDI_EMBLEM_INFORMATION_64_PNG;
    }
  if (m_is_valid)
    {
      return IDI_EMBLEM_SUCCESS_64_PNG;
    }
  const auto sig = m_verify_result.signature (0);
  if ((sig.summary() & Signature::Summary::KeyMissing))
    {
      return IDI_EMBLEM_QUESTION_64_PNG;
    }

  /* Maybe warning for unsigned and invalid sigs? */

  return IDI_EMBLEM_INFORMATION_64_PNG;
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
