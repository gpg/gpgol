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
#include "parsecontroller.h"
#include "gpgolstr.h"
#include "windowmessages.h"

#include <map>
#include <vector>
#include <memory>

static std::map<LPDISPATCH, Mail*> g_mail_map;

#define COPYBUFSIZE (8 * 1024)

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

bool
Mail::is_mail_valid (const Mail *mail)
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

/** Helper to update the attachments of a mail object in oom.
  does not modify the underlying mapi structure. */
static int
add_attachments(LPDISPATCH mail,
                std::vector<std::shared_ptr<Attachment> > attachments)
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
      CloseHandle (hFile);
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
  if (m_type == MSGTYPE_UNKNOWN || m_type == MSGTYPE_GPGOL)
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

  m_processed = true;
  /* Insert placeholder */
  char *placeholder_buf;
  if (gpgrt_asprintf (&placeholder_buf, decrypt_template,
                      is_smime() ? "S/MIME" : "OpenPGP",
                      _("Encrypted message"),
                      _("Please wait while the message is being decrypted...")) == -1)
    {
      log_error ("%s:%s: Failed to format placeholder.",
                 SRCNAME, __func__);
      return 1;
    }

  if (put_oom_string (m_mailitem, "HTMLBody", placeholder_buf))
    {
      log_error ("%s:%s: Failed to modify html body of item.",
                 SRCNAME, __func__);
      xfree (placeholder_buf);
      return 1;
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

  CreateThread (NULL, 0, do_parsing, (LPVOID) this, 0,
                NULL);
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
      if (put_oom_string (m_mailitem, "HTMLBody", html.c_str()))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
          return;
        }
    }
  else
    {
      const auto body = m_parser->get_body();
      if (put_oom_string (m_mailitem, "Body", body.c_str()))
        {
          log_error ("%s:%s: Failed to modify body of item.",
                     SRCNAME, __func__);
          return;
        }
    }
}

void
Mail::parsing_done()
{
  m_needs_wipe = true;
  /* Update the body */
  update_body();

  /* Update attachments */
  if (add_attachments (m_mailitem, m_parser->get_attachments()))
    {
      log_error ("%s:%s: Failed to update attachments.",
                 SRCNAME, __func__);
      return;
    }

  /* Invalidate UI to set the correct sig status. */
  delete m_parser;
  m_parser = nullptr;
  gpgoladdin_invalidate_ui ();
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
          log_error ("Failed to revert mail: %p ", it->first);
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
