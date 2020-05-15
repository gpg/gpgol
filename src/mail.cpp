/* @file mail.h
 * @brief High level class to work with Outlook Mailitems.
 *
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
 * Copyright (C) 2019 g10code GmbH
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

#include "categorymanager.h"
#include "dialogs.h"
#include "common.h"
#include "mail.h"
#include "eventsinks.h"
#include "attachment.h"
#include "mapihelp.h"
#include "mimemaker.h"
#include "revert.h"
#include "gpgoladdin.h"
#include "mymapitags.h"
#include "parsecontroller.h"
#include "cryptcontroller.h"
#include "windowmessages.h"
#include "mlang-charset.h"
#include "wks-helper.h"
#include "keycache.h"
#include "cpphelp.h"
#include "addressbook.h"
#include "recipient.h"

#include <gpgme++/configuration.h>
#include <gpgme++/tofuinfo.h>
#include <gpgme++/verificationresult.h>
#include <gpgme++/decryptionresult.h>
#include <gpgme++/key.h>
#include <gpgme++/context.h>
#include <gpgme++/keylistresult.h>
#include <gpg-error.h>

#include <map>
#include <unordered_map>
#include <set>
#include <vector>
#include <memory>
#include <sstream>

#undef _
#define _(a) utf8_gettext (a)

using namespace GpgME;

static std::map<LPDISPATCH, Mail*> s_mail_map;
static std::map<std::string, Mail*> s_uid_map;
static std::map<std::string, LPDISPATCH> s_folder_events_map;
static std::set<std::string> uids_searched;

GPGRT_LOCK_DEFINE (mail_map_lock);
GPGRT_LOCK_DEFINE (uid_map_lock);

static Mail *s_last_mail;
Mail *g_mail_copy_triggerer = nullptr;

#define COLOR_DARK_GREY  "#f0f0f0"
#define COLOR_LIGHT_GREY "#f8f8f8"

const char *HTML_PREVIEW_PLACEHOLDER =
"<html><head></head><body>"
"<table border=\"0\" width=\"100%%\" cellspacing=\"1\" cellpadding=\"1\" bgcolor=\"" COLOR_DARK_GREY "\">"
"<tr>"
"<td bgcolor=\"" COLOR_DARK_GREY "\">"
"<p><span style=\"font-weight:600; background-color:" COLOR_DARK_GREY ";\"><center>%s %s</center><span></p></td></tr>"
"<tr>"
"<td bgcolor=\"" COLOR_DARK_GREY "\">"
"<p><span style=\"font-weight:600; background-color:" COLOR_DARK_GREY ";\"><center>%s</center><span></p></td></tr>"
"<tr>"
"<td bgcolor=\"" COLOR_LIGHT_GREY "\">"
"%s"
"</td></tr>"
"</table></body></html>";

const char *TEXT_PREVIEW_PLACEHOLDER = "%s %s %s\n\n%s";
#define COPYBUFSIZE (8 * 1024)

Mail::Mail (LPDISPATCH mailitem) :
    m_mailitem(mailitem),
    m_currentItemRef(nullptr),
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
    m_type(MSGTYPE_UNKNOWN),
    m_do_inline(false),
    m_is_gsuite(false),
    m_crypt_state(NoCryptMail),
    m_window(nullptr),
    m_async_crypt_disabled(false),
    m_is_forwarded_crypto_mail(false),
    m_is_reply_crypto_mail(false),
    m_is_send_again(false),
    m_disable_att_remove_warning(false),
    m_manual_crypto_opts(false),
    m_first_autosecure_check(true),
    m_locate_count(0),
    m_pass_write(false),
    m_locate_in_progress(false),
    m_is_junk(false),
    m_is_draft_encrypt(false),
    m_decrypt_again(false),
    m_printing(false),
    m_recipients_set(false),
    m_is_split_copy(false),
    m_attachs_added(false)
{
  TSTART;
  if (getMailForItem (mailitem))
    {
      log_error ("Mail object for item: %p already exists. Bug.",
                 mailitem);
      TRETURN;
    }

  m_event_sink = install_MailItemEvents_sink (mailitem);
  if (!m_event_sink)
    {
      /* Should not happen but in that case we don't add us to the list
         and just release the Mail item. */
      log_error ("%s:%s: Failed to install MailItemEvents sink.",
                 SRCNAME, __func__);
      gpgol_release(mailitem);
      TRETURN;
    }
  gpgol_lock (&mail_map_lock);
  s_mail_map.insert (std::pair<LPDISPATCH, Mail *> (mailitem, this));
  gpgol_unlock (&mail_map_lock);
  s_last_mail = this;
  memdbg_ctor ("Mail");
  TRETURN;
}

GPGRT_LOCK_DEFINE(dtor_lock);

// static
void
Mail::lockDelete ()
{
  TSTART;
  gpgol_lock (&dtor_lock);
  TRETURN;
}

// static
void
Mail::unlockDelete ()
{
  TSTART;
  gpgol_unlock (&dtor_lock);
  TRETURN;
}

Mail::~Mail()
{
  TSTART;
  /* This should fix a race condition where the mail is
     deleted before the parser is accessed in the decrypt
     thread. The shared_ptr of the parser then ensures
     that the parser is alive even if the mail is deleted
     while parsing. */
  gpgol_lock (&dtor_lock);
  memdbg_dtor ("Mail");
  log_oom ("%s:%s: dtor: Mail: %p item: %p",
                 SRCNAME, __func__, this, m_mailitem);
  std::map<LPDISPATCH, Mail *>::iterator it;

  log_oom ("%s:%s: Detaching event sink",
                 SRCNAME, __func__);
  detach_MailItemEvents_sink (m_event_sink);
  gpgol_release(m_event_sink);

  log_oom ("%s:%s: Erasing mail",
                 SRCNAME, __func__);
  gpgol_lock (&mail_map_lock);
  it = s_mail_map.find(m_mailitem);
  if (it != s_mail_map.end())
    {
      s_mail_map.erase (it);
    }
  gpgol_unlock (&mail_map_lock);

  if (!m_uuid.empty())
    {
      gpgol_lock (&uid_map_lock);
      auto it2 = s_uid_map.find(m_uuid);
      if (it2 != s_uid_map.end())
        {
          s_uid_map.erase (it2);
        }
      gpgol_unlock (&uid_map_lock);
    }

  log_oom ("%s:%s: removing categories",
                 SRCNAME, __func__);
  removeCategories_o ();

  log_oom ("%s:%s: releasing mailitem",
                 SRCNAME, __func__);
  gpgol_release(m_mailitem);
  xfree (m_cached_html_body);
  xfree (m_cached_plain_body);
  if (!m_uuid.empty())
    {
      log_oom ("%s:%s: destroyed: %p uuid: %s",
                     SRCNAME, __func__, this, m_uuid.c_str());
    }
  else
    {
      log_oom ("%s:%s: non crypto (or sent) mail: %p destroyed",
                     SRCNAME, __func__, this);
    }
  log_oom ("%s:%s: nulling shared pointer",
                 SRCNAME, __func__);
  m_parser = nullptr;
  m_crypter = nullptr;

  releaseCurrentItem();
  gpgol_unlock (&dtor_lock);
  log_oom ("%s:%s: returning",
                 SRCNAME, __func__);
  TRETURN;
}

//static
Mail *
Mail::getMailForItem (LPDISPATCH mailitem)
{
  TSTART;
  if (!mailitem)
    {
      TRETURN NULL;
    }
  std::map<LPDISPATCH, Mail *>::iterator it;
  gpgol_lock (&mail_map_lock);
  it = s_mail_map.find(mailitem);
  gpgol_unlock (&mail_map_lock);
  if (it == s_mail_map.end())
    {
      TRETURN NULL;
    }
  TRETURN it->second;
}

//static
Mail *
Mail::getMailForUUID (const char *uuid)
{
  TSTART;
  if (!uuid)
    {
      TRETURN NULL;
    }
  gpgol_lock (&uid_map_lock);
  auto it = s_uid_map.find(std::string(uuid));
  gpgol_unlock (&uid_map_lock);
  if (it == s_uid_map.end())
    {
      TRETURN NULL;
    }
  TRETURN it->second;
}

//static
bool
Mail::isValidPtr (const Mail *mail)
{
  TSTART;
  gpgol_lock (&mail_map_lock);
  auto it = s_mail_map.begin();
  while (it != s_mail_map.end())
    {
      if (it->second == mail)
        {
          gpgol_unlock (&mail_map_lock);
          TRETURN true;
        }
      ++it;
    }
  gpgol_unlock (&mail_map_lock);
  TRETURN false;
}

int
Mail::preProcessMessage_m ()
{
  TSTART;
  if (m_is_split_copy)
    {
      log_dbg ("Mail was created as a copy by gpgol. Addr: %p",
               this);
      TRETURN 0;
    }
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      TRETURN 0;
    }
  log_oom ("%s:%s: GetBaseMessage OK for %p.",
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
      /* For unknown messages we still need to check for autocrypt
         headers. If the mails are crypto messages the autocrypt
         stuff is handled in the parsecontroller. */
      autocrypt_s ac;
      parseHeaders_m ();
      if (m_header_info.acInfo.exists)
        {
          log_debug ("%s:%s: Importing autocrypt header from unencrypted "
                     "mail.", SRCNAME, __func__);
          KeyCache::import_pgp_key_data (m_header_info.acInfo.data);
        }
      gpgol_release (message);
      TRETURN 0;
    }

  /* We could check PR_ACCESS here in MAPI to figure out if we can
     modify a mail or not. But this strangely does not fully tell
     us the truth. For example for a read only mail we can modify
     the body of the mail but "add oom attachments" will fail.

     Add OOM attachments has error handling and it will show the
     user that missing rights are an issue. */

  /* Create moss attachments here so that they are properly
     hidden when the item is read into the model. */
  LPMESSAGE parsed_message = get_oom_message (m_mailitem);
  m_moss_position = mapi_mark_or_create_moss_attach (message, parsed_message,
                                                     m_type);
  gpgol_release (parsed_message);
  if (!m_moss_position)
    {
      log_error ("%s:%s: Failed to find moss attachment.",
                 SRCNAME, __func__);
      m_type = MSGTYPE_UNKNOWN;
    }

  gpgol_release (message);
  TRETURN 0;
}

static LPDISPATCH
get_attachment_o (LPDISPATCH mailitem, int pos)
{
  TSTART;
  LPDISPATCH attachment;
  LPDISPATCH attachments = get_oom_object (mailitem, "Attachments");
  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      TRETURN NULL;
    }

  std::string item_str;
  int count = get_oom_int (attachments, "Count");
  if (count < 1)
    {
      log_debug ("%s:%s: Invalid attachment count: %i.",
                 SRCNAME, __func__, count);
      gpgol_release (attachments);
      TRETURN NULL;
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

  TRETURN attachment;
}

/** Helper to check that all attachments are hidden, to be
  called before crypto. */
int
Mail::checkAttachments_o (bool silent)
{
  TSTART;
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      log_debug ("%s:%s: Failed to get attachments.",
                 SRCNAME, __func__);
      TRETURN 1;
    }
  int count = count_visible_attachments (attachments);
  if (!count)
    {
      gpgol_release (attachments);
      TRETURN 0;
    }

  std::string message;

  if (isEncrypted () && isSigned ())
    {
      message += _("Not all attachments were encrypted or signed.\n"
                   "The unsigned / unencrypted attachments are:\n\n");
    }
  else if (isSigned ())
    {
      message += _("Not all attachments were signed.\n"
                   "The unsigned attachments are:\n\n");
    }
  else if (isEncrypted ())
    {
      message += _("Not all attachments were encrypted.\n"
                   "The unencrypted attachments are:\n\n");
    }
  else
    {
      gpgol_release (attachments);
      TRETURN 0;
    }

  bool foundOne = false;
  std::vector <LPDISPATCH> to_delete;
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
          to_delete.push_back (oom_attach);
        }
      else
        {
          gpgol_release (oom_attach);
        }
      VariantClear (&var);
    }

  /* In silent mode we remove the unencrypted attachments silently */
  if (foundOne)
    {
      for (auto attachment: to_delete)
        {
          if (silent)
            {
              m_disable_att_remove_warning = true;
              log_debug ("%s:%s: Deleting bad attachment",
                         SRCNAME, __func__);
              if (invoke_oom_method (attachment, "Delete", NULL))
                {
                  log_error ("%s:%s: Error deleting attachment: %i",
                             SRCNAME, __func__, __LINE__);
                }
              m_disable_att_remove_warning = false;
            }
          gpgol_release (attachment);
        }
    }
  gpgol_release (attachments);
  if (foundOne && !silent)
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
  TRETURN 0;
}

/** Get the cipherstream of the mailitem. */
static LPSTREAM
get_attachment_stream_o (LPDISPATCH mailitem, int pos)
{
  TSTART;
  if (!pos)
    {
      log_debug ("%s:%s: Called with zero pos.",
                 SRCNAME, __func__);
      TRETURN NULL;
    }
  LPDISPATCH attachment = get_attachment_o (mailitem, pos);
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
          TRETURN NULL;
        }
      hr = gpgol_openProperty (message, PR_BODY_A, &IID_IStream, 0, 0,
                               (LPUNKNOWN*)&stream);
      gpgol_release (message);
      if (hr)
        {
          log_debug ("%s:%s: OpenProperty failed: hr=%#lx",
                     SRCNAME, __func__, hr);
          TRETURN NULL;
        }
      TRETURN stream;
    }

  LPATTACH mapi_attachment = NULL;

  mapi_attachment = (LPATTACH) get_oom_iunknown (attachment,
                                                 "MapiObject");
  gpgol_release (attachment);
  if (!mapi_attachment)
    {
      log_debug ("%s:%s: Failed to get MapiObject of attachment: %p",
                 SRCNAME, __func__, attachment);
      TRETURN NULL;
    }
  if (FAILED (gpgol_openProperty (mapi_attachment, PR_ATTACH_DATA_BIN,
                                  &IID_IStream, 0, MAPI_MODIFY,
                                  (LPUNKNOWN*) &stream)))
    {
      log_debug ("%s:%s: Failed to open stream for mapi_attachment: %p",
                 SRCNAME, __func__, mapi_attachment);
      gpgol_release (mapi_attachment);
      TRETURN nullptr;
    }
  gpgol_release (mapi_attachment);
  TRETURN stream;
}

#if 0

This should work. But Outlook says no. See the comment in set_pa_variant
about this. I left the code here as an example how to work with
safearrays and how this probably should work.

static int
copy_data_property(LPDISPATCH target, std::shared_ptr<Attachment> attach)
{
  TSTART;
  VARIANT var;
  VariantInit (&var);

  /* Get the size */
  off_t size = attach->get_data ().seek (0, SEEK_END);
  attach->get_data ().seek (0, SEEK_SET);

  if (!size)
    {
      TRACEPOINT;
      TRETURN 1;
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
      TRETURN 1;
    }

  void *buffer = NULL;
  /* Get a safe pointer to the array */
  if (SafeArrayAccessData(var.parray, &buffer) != S_OK)
    {
      TRACEPOINT;
      VariantClear(&var);
      TRETURN 1;
    }

  /* Copy data to it */
  size_t nread = attach->get_data ().read (buffer, static_cast<size_t> (size));

  if (nread != static_cast<size_t> (size))
    {
      TRACEPOINT;
      VariantClear(&var);
      TRETURN 1;
    }

  /*/ Unlock the variant data */
  if (SafeArrayUnaccessData(var.parray) != S_OK)
    {
      TRACEPOINT;
      VariantClear(&var);
      TRETURN 1;
    }

  if (set_pa_variant (target, PR_ATTACH_DATA_BIN_DASL, &var))
    {
      TRACEPOINT;
      VariantClear(&var);
      TRETURN 1;
    }

  VariantClear(&var);
  TRETURN 0;
}
#endif

static int
copy_attachment_to_file (std::shared_ptr<Attachment> att, HANDLE hFile)
{
  TSTART;
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
          TRETURN 1;
        }
      if (nread != nwritten)
        {
          log_error ("%s:%s: Write truncated.",
                     SRCNAME, __func__);
          TRETURN 1;
        }
    }
  TRETURN 0;
}

/** Sets some meta data on the last attachment added. The meta
  data is taken from the attachment object. */
static int
fixup_last_attachment_o (LPDISPATCH mail,
                         std::shared_ptr<Attachment> attachment)
{
  TSTART;
  /* Currently we only set content id */
  std::string cid = attachment->get_content_id ();
  if (cid.empty())
    {
      log_debug ("%s:%s: Content id not found.",
                 SRCNAME, __func__);
      TRETURN 0;
    }

  /* We add this safeguard here was we somehow can't trigger Outlook
   * not to hide attachments with content id (see below). So we
   * have to make double sure that the attachment is actually referenced
   * in the HTML before we hide it. In doubt it is better to "break"
   * the content id reference but show the Attachment as a hidden
   * attachment can appear to be data loss. See T4526 T4203. */
  char *body = get_oom_string (mail, "HTMLBody");
  if (!body)
    {
      log_debug ("%s:%s: No HTML Body.",
                 SRCNAME, __func__);
      TRETURN 0;
    }

  if (cid.front () == '<')
    {
      cid.erase (0, 1);
    }
  if (cid.back() == '>')
    {
      cid.pop_back ();
    }

  char *cid_pos = strstr (body, cid.c_str ());
  xfree (body);
  if (!cid_pos)
    {
      log_debug ("%s:%s: Failed to find cid: '%s' in body. Not setting cid.",
                 SRCNAME, __func__, anonstr (cid.c_str ()));

      TRETURN 0;
    }

  LPDISPATCH attach = get_attachment_o (mail, -1);
  if (!attach)
    {
      log_error ("%s:%s: No attachment.",
                 SRCNAME, __func__);
      TRETURN 1;
    }
  int ret = put_pa_string (attach,
                           PR_ATTACH_CONTENT_ID_DASL,
                           cid.c_str ());

  log_debug ("%s:%s: Set attachment content id to: '%s'",
             SRCNAME, __func__, anonstr (cid.c_str()));
  if (ret)
    {
      log_error ("%s:%s: Failed.", SRCNAME, __func__);
      gpgol_release (attach);
    }
#if 0

  The following was an experiement to delete the ATTACH_FLAGS values
  so that we are not hiding attachments.

  LPATTACH mapi_attach = (LPATTACH) get_oom_iunknown (attach, "MAPIOBJECT");
  if (mapi_attach)
    {
      SPropTagArray proparray;
      HRESULT hr;

      proparray.cValues = 1;
      proparray.aulPropTag[0] = 0x37140003;
      hr = mapi_attach->DeleteProps (&proparray, NULL);
      if (hr)
        {
          log_error ("%s:%s: can't delete property attach flags: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          ret = -1;
        }
      gpgol_release (mapi_attach);
    }
  else
    {
      log_error ("%s:%s: Failed to get mapi attachment.",
                 SRCNAME, __func__);
    }
#endif
  gpgol_release (attach);
  TRETURN ret;
}

/** Helper to update the attachments of a mail object in oom.
  does not modify the underlying mapi structure. */
int
Mail::add_attachments_o (std::vector<std::shared_ptr<Attachment> > attachments)
{
  TSTART;
  if (m_attachs_added)
    {
      log_dbg ("Not adding attachments as they are already there.");
      TRETURN 0;
    }
  bool anyError = false;
  m_disable_att_remove_warning = true;

  std::string addErrStr;
  int addErrCode = 0;
  std::vector<std::string> failedNames;
  for (auto att: attachments)
    {
      int err = 0;
      const auto dispName = att->get_display_name ();
      if (dispName.empty())
        {
          log_error ("%s:%s: Ignoring attachment without display name.",
                     SRCNAME, __func__);
          continue;
        }
      wchar_t* wchar_name = utf8_to_wchar (dispName.c_str());
      if (!wchar_name)
        {
          log_error ("%s:%s: Failed to convert '%s' to wchar.",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          continue;
        }

      HANDLE hFile;
      wchar_t* wchar_file = get_tmp_outfile (wchar_name,
                                             &hFile);
      if (!wchar_file)
        {
          log_error ("%s:%s: Failed to obtain a tmp filename for: %s",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          err = 1;
        }
      if (!err && copy_attachment_to_file (att, hFile))
        {
          log_error ("%s:%s: Failed to copy attachment %s to temp file",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          err = 1;
        }
      if (!err && add_oom_attachment (m_mailitem, wchar_file, wchar_name,
                                      addErrStr, &addErrCode))
        {
          log_error ("%s:%s: Failed to add attachment: %s",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          failedNames.push_back (dispName);
          err = 1;
        }
      if (hFile && hFile != INVALID_HANDLE_VALUE)
        {
          CloseHandle (hFile);
        }
      if (wchar_file && !DeleteFileW (wchar_file))
        {
          log_error ("%s:%s: Failed to delete tmp attachment for: %s",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          err = 1;
        }
      xfree (wchar_file);
      xfree (wchar_name);

      if (!err)
        {
          log_debug ("%s:%s: Added attachment '%s'",
                     SRCNAME, __func__, anonstr (dispName.c_str()));
          err = fixup_last_attachment_o (m_mailitem, att);
        }
      if (err)
        {
          anyError = true;
        }
    }
  if (anyError)
    {
      std::string msg = _("Not all attachments can be shown.\n\n"
                          "The hidden attachments are:");
      msg += "\n";

      std::string filenames;
      join (failedNames, "\n", filenames);
      msg += filenames;
      msg += "\n\n";

      if (addErrCode == 0x80004005)
        {
          msg += _("The mail exceeds the maximum size GpgOL "
                   "can handle on this server.");
        }
      else
        {
          msg += _("Reason:");
          msg += " " + addErrStr;
        }
      gpgol_message_box (getWindow (),
                         msg.c_str (), _("GpgOL"), MB_OK);

    }
  m_disable_att_remove_warning = false;
  m_attachs_added = true;
  TRETURN anyError;
}

GPGRT_LOCK_DEFINE(parser_lock);

static DWORD WINAPI
do_parsing (LPVOID arg)
{
  TSTART;
  gpgol_lock (&dtor_lock);
  /* We lock with mail dtors so we can be sure the mail->parser
     call is valid. */
  Mail *mail = (Mail *)arg;
  if (!Mail::isValidPtr (mail))
    {
      log_debug ("%s:%s: canceling parsing for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&dtor_lock);
      TRETURN 0;
    }

  blockInv ();
  /* This takes a shared ptr of parser. So the parser is
     still valid when the mail is deleted. */
  auto parser = mail->parser ();
  gpgol_unlock (&dtor_lock);

  gpgol_lock (&parser_lock);
  /* We lock the parser here to avoid too many
     decryption attempts if there are
     multiple mailobjects which might have already
     been deleted (e.g. by quick switches of the mailview.)
     Let's rather be a bit slower.
     */
  log_debug ("%s:%s: preparing the parser for: %p",
             SRCNAME, __func__, arg);

  if (!Mail::isValidPtr (mail))
    {
      log_debug ("%s:%s: cancel for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&parser_lock);
      unblockInv();
      TRETURN 0;
    }

  if (!parser)
    {
      log_error ("%s:%s: no parser found for mail: %p",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&parser_lock);
      unblockInv();
      TRETURN -1;
    }

  bool is_smime = mail->isSMIME ();

  std::vector<GpgME::Key> senderKey;
  if (!is_smime)
    {
      senderKey = KeyCache::instance ()->getEncryptionKeys (mail->getSender (),
                                                            GpgME::OpenPGP);
    }
  bool has_already_preview = (mail->msgtype () == MSGTYPE_GPGOL_MULTIPART_SIGNED &&
                              !mail->getOriginalBody ().empty ());
  if (!has_already_preview && (senderKey.empty () || is_smime) &&
      KeyCache::instance ()->protocolIsOnline (is_smime ? GpgME::CMS :
                                                          GpgME::OpenPGP) &&
      !opt.sync_dec)
    {
      log_dbg ("Op might be online. Doing two pass verify.");
      parser->parse (true);
      const auto verify_result = parser->verify_result ();
      if (verify_result.numSignatures ())
        {
          log_dbg ("Have signature, needs second pass.");
          do_in_ui_thread (SHOW_PREVIEW, arg);
          if (!Mail::isValidPtr (mail))
            {
              log_debug ("%s:%s: canceling parsing for: %p now deleted",
                         SRCNAME, __func__, arg);
              gpgol_unlock (&parser_lock);
              unblockInv();
              TRETURN 0;
            }
          log_dbg ("Preview updated.");
          /* Sleep (10000); */
          parser->parse (false);
          log_dbg ("Second parse done.");
        }
      if (!opt.sync_dec)
        {
          do_in_ui_thread (PARSING_DONE, arg);
        }
    }
  else
    {
      parser->parse (false);
      if (!opt.sync_dec)
        {
          do_in_ui_thread (PARSING_DONE, arg);
        }
    }
  gpgol_unlock (&parser_lock);
  unblockInv();
  TRETURN 0;
}

/* How encryption is done:

   There are two modes of encryption. Synchronous and Async.
   If async is used depends on the value of mail->async_crypt_disabled.

   Synchronous crypto:

   > Send Event < | State NoCryptMail
   Needs Crypto ? (get_gpgol_draft_info_flags != 0)

   -> No:
      Pass send -> unencrypted mail.

   -> Yes:
      mail->update_oom_data
      State = Mail::NeedsFirstAfterWrite
      checkSyncCrypto_o
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
  TSTART;
  gpgol_lock (&dtor_lock);
  /* We lock with mail dtors so we can be sure the mail->parser
     call is valid. */
  Mail *mail = (Mail *)arg;
  if (!Mail::isValidPtr (mail))
    {
      log_debug ("%s:%s: canceling crypt for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&dtor_lock);
      TRETURN 0;
    }
  if (mail->cryptState () != Mail::NeedsActualCrypt)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, mail->cryptState ());
      mail->enableWindow ();
      gpgol_unlock (&dtor_lock);
      TRETURN -1;
    }

  /* This takes a shared ptr of crypter. So the crypter is
     still valid when the mail is deleted. */
  auto crypter = mail->cryper ();
  gpgol_unlock (&dtor_lock);

  if (!crypter)
    {
      log_error ("%s:%s: no crypter found for mail: %p",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&parser_lock);
      mail->enableWindow ();
      TRETURN -1;
    }

  GpgME::Error err;
  std::string diag;
  int rc = crypter->do_crypto(err, diag);

  gpgol_lock (&dtor_lock);
  if (!Mail::isValidPtr (mail))
    {
      log_debug ("%s:%s: aborting crypt for: %p already deleted",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&dtor_lock);
      TRETURN 0;
    }

  mail->enableWindow ();

  if (rc == -1 || err)
    {
      mail->resetCrypter ();
      crypter = nullptr;
      if (err)
        {
          char *buf = nullptr;
          gpgrt_asprintf (&buf, _("Crypto operation failed:\n%s"),
                          err.asString());
          std::string msg = buf;
          memdbg_alloc (buf);
          xfree (buf);
          if (!diag.empty())
            {
              msg += "\n\n";
              msg += _("Diagnostics");
              msg += ":\n";
              msg += diag;
            }
          gpgol_message_box (mail->getWindow (), msg.c_str (),
                             _("GpgOL"), MB_OK);
        }
      else
        {
          gpgol_bug (mail->getWindow (),
                     ERR_CRYPT_RESOLVER_FAILED);
        }
    }

  if (rc || err.isCanceled())
    {
      log_debug ("%s:%s: crypto failed for: %p with: %i err: %i",
                 SRCNAME, __func__, arg, rc, err.code());
      mail->setCryptState (Mail::NoCryptMail);
      mail->setIsDraftEncrypt (false);
      mail->resetCrypter ();
      crypter = nullptr;
      gpgol_unlock (&dtor_lock);

      if (rc != -3)
        {
          mail->resetRecipients ();
        }
      TRETURN rc;
    }

  if (!mail->isAsyncCryptDisabled ())
    {
      mail->setCryptState (Mail::NeedsUpdateInOOM);
      gpgol_unlock (&dtor_lock);
      // This deletes the Mail in Outlook 2010
      do_in_ui_thread (CRYPTO_DONE, arg);
      log_debug ("%s:%s: UI thread finished for %p",
                 SRCNAME, __func__, arg);
    }
  else if (mail->isDraftEncrypt ())
    {
      mail->setCryptState (Mail::NeedsUpdateInMAPI);
      mail->updateCryptMAPI_m ();
      mail->setIsDraftEncrypt (false);
      mail->setCryptState (Mail::NoCryptMail);
      log_debug ("%s:%s: Synchronous draft encrypt finished for %p",
                 SRCNAME, __func__, arg);
      gpgol_unlock (&dtor_lock);
    }
  else
    {
      mail->setCryptState (Mail::NeedsUpdateInMAPI);
      mail->updateCryptMAPI_m ();
      if (mail->cryptState () == Mail::WantsSendMIME)
        {
          // For sync crypto we need to switch this.
          mail->setCryptState (Mail::NeedsUpdateInOOM);
        }
      else
        {
          // A bug!
          log_debug ("%s:%s: Resetting crypter because of state mismatch. %p",
                     SRCNAME, __func__, arg);
          crypter = nullptr;
          mail->resetCrypter ();
        }
      gpgol_unlock (&dtor_lock);
    }
  /* This works around a bug in pinentry that it might
     bring the wrong window to front. So after encryption /
     signing we bring outlook back to front.

     See GnuPG-Bug-Id: T3732
     */
  do_in_ui_thread_async (BRING_TO_FRONT, nullptr, 250);
  log_debug ("%s:%s: crypto thread for %p finished",
             SRCNAME, __func__, arg);
  TRETURN 0;
}

bool
Mail::isCryptoMail () const
{
  TSTART;
  if (m_type == MSGTYPE_UNKNOWN || m_type == MSGTYPE_GPGOL ||
      m_type == MSGTYPE_SMIME)
    {
      /* Not a message for us. */
      TRETURN false;
    }
  TRETURN true;
}

int
Mail::decryptVerify_o ()
{
  TSTART;

  if (!isCryptoMail ())
    {
      log_debug ("%s:%s: Decrypt Verify for non crypto mail: %p.",
                 SRCNAME, __func__, m_mailitem);
      TRETURN 0;
    }
  m_decrypt_again = false;

  if (isSMIME_m ())
    {
      LPMESSAGE oom_message = get_oom_message (m_mailitem);
      if (oom_message)
        {
          char *old_class = mapi_get_old_message_class (oom_message);
          char *current_class = mapi_get_message_class (oom_message);
          if (current_class)
            {
              /* Store our own class for an eventual close */
              m_gpgol_class = current_class;
              xfree (current_class);
              current_class = nullptr;
            }
          if (old_class)
            {
              const char *new_class = old_class;
              /* Workaround that our own class might be the original */
              if (!strcmp (old_class, "IPM.Note.GpgOL.OpaqueEncrypted"))
                {
                  new_class = "IPM.Note.SMIME";
                }
              else if (!strcmp (old_class, "IPM.Note.GpgOL.MultipartSigned"))
                {
                  new_class = "IPM.Note.SMIME.MultipartSigned";
                }

              log_debug ("%s:%s:Restoring message class to %s in decverify.",
                         SRCNAME, __func__, new_class);

              put_oom_string (m_mailitem, "MessageClass", new_class);
              xfree (old_class);
              setPassWrite (true);
              /* Sync to MAPI */
              invoke_oom_method (m_mailitem, "Save", nullptr);
              setPassWrite (false);
            }
          gpgol_release (oom_message);
        }
    }

  check_html_preferred ();

  auto cipherstream = get_attachment_stream_o (m_mailitem, m_moss_position);
  if (!cipherstream)
    {
      m_is_junk = is_junk_mail (m_mailitem);
      if (m_is_junk)
        {
          log_debug ("%s:%s: Detected: %p as junk",
                     SRCNAME, __func__, m_mailitem);
          auto mngr = CategoryManager::instance ();
          m_store_id = mngr->addCategoryToMail (this,
                                   CategoryManager::getJunkMailCategory (),
                                   3 /* peach */);
          installFolderEventHandler_o ();
          TRETURN 0;
        }
      log_debug ("%s:%s: Failed to get cipherstream. Aborting handling.",
                 SRCNAME, __func__);
      m_type = MSGTYPE_UNKNOWN;
      TRETURN 1;
    }

  setUUID_o ();
  m_processed = true;
  m_pass_write = false;

  /* Insert placeholder */
  m_orig_body = get_oom_string_s (m_mailitem, "Body");
  rtrim (m_orig_body);
  if (m_orig_body.empty ())
    {
      log_dbg ("Empty body.");
      /* We only need to check the HTML body if the plain body
         is not empty as outlook would convert the html to plain. */
    }
  else if (opt.prefer_html)
    {
      m_orig_body = get_oom_string_s (m_mailitem, "HTMLBody");
    }

  char *placeholder_buf = nullptr;
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
  else if (m_type == MSGTYPE_GPGOL_MULTIPART_SIGNED &&
           !m_orig_body.empty ())
    {
      log_dbg ("Multipart signed inserting body in placeholder.");
      gpgrt_asprintf (&placeholder_buf, opt.prefer_html ? HTML_PREVIEW_PLACEHOLDER :
                      TEXT_PREVIEW_PLACEHOLDER,
                      isSMIME_m () ? "S/MIME" : "OpenPGP",
                      _("message"),
                      _("Please wait while the message is being verified..."),
                      m_orig_body.c_str ());
    }
  else
    {
      gpgrt_asprintf (&placeholder_buf, opt.prefer_html ? decrypt_template_html :
                      decrypt_template,
                      isSMIME_m () ? "S/MIME" : "OpenPGP",
                      _("message"),
                      _("Please wait while the message is being decrypted / verified..."));
    }
  if (opt.prefer_html)
    {
      if (put_oom_string (m_mailitem, "HTMLBody", placeholder_buf))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
        }
      put_oom_int (m_mailitem, "BodyFormat", 2);
    }
  else
    {
      char *tmp = get_oom_string (m_mailitem, "Body");
      if (!tmp)
        {
          TRACEPOINT;
          TRETURN 1;
        }
      m_orig_body = tmp;
      xfree (tmp);
      if (put_oom_string (m_mailitem, "Body", placeholder_buf))
        {
          log_error ("%s:%s: Failed to modify body of item.",
                     SRCNAME, __func__);
        }
      put_oom_int (m_mailitem, "BodyFormat", 1);
    }
  memdbg_alloc (placeholder_buf);
  xfree (placeholder_buf);

  if (m_type == MSGTYPE_GPGOL_WKS_CONFIRMATION)
    {
      if (m_header_info.boundary.empty())
        {
          parseHeaders_m ();
        }
      WKSHelper::instance ()->handle_confirmation_read (this, cipherstream);
      TRETURN 0;
    }

  m_parser = std::shared_ptr <ParseController> (new ParseController (cipherstream, m_type));
  m_parser->setSender(GpgME::UserID::addrSpecFromString(getSender_o ().c_str()));

  if (opt.autoimport)
    {
      /* Handle autocrypt header. As we want to have the import
         of the header in the same thread as the parser we leave it
         to the parser. */
      if (m_header_info.boundary.empty())
        {
          parseHeaders_m ();
        }
      if (m_header_info.acInfo.exists)
        {
          m_parser->setAutocryptInfo (m_header_info.acInfo);
        }
    }

  log_data ("%s:%s: Parser for \"%s\" is %p",
                   SRCNAME, __func__, getSubject_o ().c_str(), m_parser.get());
  gpgol_release (cipherstream);

  /* Printing happens in two steps. First a Mail is loaded after the
     BeforePrint event, then it is loaded a second time when the actual print
     happens. We have to catch both. */
  if (!m_printing)
    {
      m_printing = checkIfMailIsChildOfPrintMail_o ();
    }

  if (!opt.sync_dec && !m_printing)
    {
      HANDLE parser_thread = CreateThread (NULL, 0, do_parsing, (LPVOID) this, 0,
                                           NULL);

      if (!parser_thread)
        {
          log_error ("%s:%s: Failed to create decrypt / verify thread.",
                     SRCNAME, __func__);
        }
      CloseHandle (parser_thread);
      TRETURN 0;
    }
  else
    {
      /* Parse synchronously */
      do_parsing ((LPVOID) this);
      parsingDone_o ();
      TRETURN 0;
    }
}

int
Mail::parseHeaders_m ()
{
  /* Parse the headers first so that the handler for
     confirmation read can access them.
     Should be put in its own function.
     */
  auto message = MAKE_SHARED (get_oom_message (m_mailitem));
  if (!message)
    {
      /* Hmmm */
      STRANGEPOINT;
      TRETURN -1;
    }
  if (!mapi_get_header_info ((LPMESSAGE)message.get(),
                             m_header_info))
    {
      STRANGEPOINT;
      TRETURN -1;
    }
  TRETURN 0;
}

static void
set_body (LPDISPATCH item, const std::string &plain, const std::string &html)
{
  if (opt.prefer_html && !html.empty())
    {
      if (put_oom_string (item, "HTMLBody", html.c_str ()))
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
          if (!plain.empty ())
            {
              if (put_oom_string (item, "Body", plain.c_str ()))
                {
                  log_error ("%s:%s: Failed to put plaintext into body of item.",
                             SRCNAME, __func__);
                }
              put_oom_int (item, "BodyFormat", 1);
            }
          else
            {
              if (put_oom_string (item, "HTMLBody", plain.c_str ()))
                {
                  log_error ("%s:%s: Failed to put plaintext into html of item.",
                             SRCNAME, __func__);
                }
              put_oom_int (item, "BodyFormat", 2);
            }
        }
      else
        {
          put_oom_int (item, "BodyFormat", 2);
        }
    }
  else if (!plain.empty ())
    {
      if (put_oom_string (item, "Body", plain.c_str ()))
        {
          log_error ("%s:%s: Failed to put plaintext into body of item.",
                     SRCNAME, __func__);
        }
      put_oom_int (item, "BodyFormat", 1);
    }
}

void
Mail::updateBody_o (bool is_preview)
{
  TSTART;
  if (!m_parser)
    {
      TRACEPOINT;
      TRETURN;
    }

  const auto error = m_parser->get_formatted_error ();
  if (!error.empty())
    {
      set_body (m_mailitem, error, error);
      TRETURN;
    }
  if (m_verify_result.error())
    {
      log_error ("%s:%s: Verification failed. Restoring Body.",
                 SRCNAME, __func__);
      set_body (m_mailitem, m_orig_body, m_orig_body);
      TRETURN;
    }
  // No need to carry body anymore
  if (!is_preview)
    {
      m_orig_body = std::string();
    }
  auto html = m_parser->get_html_body ();
  auto body = m_parser->get_body ();
  /** Outlook does not show newlines if \r\r\n is a newline. We replace
    these as apparently some other buggy MUA sends this. */
  find_and_replace (html, "\r\r\n", "\r\n");
  if (opt.prefer_html && (!html.empty() || (is_preview && !m_block_html)))
    {
      if (!m_block_html)
        {
          auto charset = m_parser->get_html_charset();

          int codepage = 0;
          if (charset.empty())
            {
              codepage = get_oom_int (m_mailitem, "InternetCodepage");
              log_debug ("%s:%s: Did not find html charset."
                         " Using internet Codepage %i.",
                         SRCNAME, __func__, codepage);
            }

          char *converted = nullptr;

          if (!html.empty () || !is_preview)
            {
              converted = ansi_charset_to_utf8 (charset.c_str(), html.c_str(),
                                                html.size(), codepage);

            }
          if (is_preview)
            {
              char *buf;
              if (!converted)
                {
                  /* Convert plaintext to HTML for preview using outlook. */
                  charset = m_parser->get_body_charset ();
                  converted = ansi_charset_to_utf8 (charset.c_str(), body.c_str(),
                                                    body.size(), codepage);
                  put_oom_string (m_mailitem, "Body", converted);
                  xfree (converted);
                  converted = get_oom_string (m_mailitem, "HTMLBody");
                }
              if (converted)
              {
                gpgrt_asprintf (&buf, HTML_PREVIEW_PLACEHOLDER,
                                isSMIME_m () ? "S/MIME" : "OpenPGP",
                                _("message"),
                                _("Please wait while the message is being verified..."),
                                converted);
                memdbg_alloc (buf);
                xfree (converted);
                converted = buf;
              }
            }
          TRACEPOINT;
          int ret = put_oom_string (m_mailitem, "HTMLBody", converted ?
                                                            converted : "");
          xfree (converted);
          put_oom_int (m_mailitem, "BodyFormat", 2);
          TRACEPOINT;
          if (ret)
            {
              log_error ("%s:%s: Failed to modify html body of item.",
                         SRCNAME, __func__);
            }

          TRETURN;
        }
      else if (!body.empty())
        {
          /* We had a multipart/alternative mail but html should be
             blocked. So we prefer the text/plain part and warn
             once about this so that we hopefully don't get too
             many bugreports about this. */
          if (!opt.smime_html_warn_shown)
            {
              std::string caption = _("GpgOL") + std::string (": ") +
                std::string (_("HTML display disabled."));
              std::string buf = _("HTML content in unsigned S/MIME mails "
                                  "is insecure.");
              buf += "\n";
              buf += _("GpgOL will only show such mails as text.");

              buf += "\n\n";
              buf += _("This message is shown only once.");

              gpgol_message_box (getWindow (), buf.c_str(), caption.c_str(),
                                 MB_OK);
              opt.smime_html_warn_shown = true;
              write_options ();
            }
        }
    }

  if (body.empty () && m_block_html && !html.empty())
    {
#if 0
      Sadly the following code still offers to load external references
      it might also be too dangerous if Outlook somehow autoloads the
      references as soon as the Body is put into HTML


      // Fallback to show HTML as plaintext if HTML display
      // is blocked.
      log_error ("%s:%s: No text body. Putting HTML into plaintext.",
                 SRCNAME, __func__);

      char *converted = ansi_charset_to_utf8 (m_parser->get_html_charset().c_str(),
                                              html.c_str(), html.size());
      int ret = put_oom_string (m_mailitem, "HTMLBody", converted ? converted : "");
      xfree (converted);
      if (ret)
        {
          log_error ("%s:%s: Failed to modify html body of item.",
                     SRCNAME, __func__);
          body = html;
        }
      else
        {
          char *plainBody = get_oom_string (m_mailitem, "Body");

          if (!plainBody)
            {
              log_error ("%s:%s: Failed to obtain converted plain body.",
                         SRCNAME, __func__);
              body = html;
            }
          else
            {
              ret = put_oom_string (m_mailitem, "HTMLBody", plainBody);
              xfree (plainBody);
              if (ret)
                {
                  log_error ("%s:%s: Failed to put plain into html body of item.",
                             SRCNAME, __func__);
                  body = html;
                }
              else
                {
                  TRETURN;
                }
            }
        }
#endif
      body = html;
      std::string caption = _("GpgOL") + std::string (": ") +
        std::string (_("HTML display disabled."));
      std::string buf = _("HTML content in unsigned S/MIME mails "
                          "is insecure.");
      buf += "\n";
      buf += _("GpgOL will only show such mails as text.");

      buf += "\n\n";
      buf += _("Please ask the sender to sign the message or\n"
               "to send it with a plain text alternative.");

      gpgol_message_box (getWindow (), buf.c_str(), caption.c_str(),
                         MB_OK);
    }

  find_and_replace (body, "\r\r\n", "\r\n");

  const auto plain_charset = m_parser->get_body_charset();

  int codepage = 0;
  if (plain_charset.empty())
    {
      codepage = get_oom_int (m_mailitem, "InternetCodepage");
      log_debug ("%s:%s: Did not find body charset. "
                 "Using internet Codepage %i.",
                 SRCNAME, __func__, codepage);
    }

  char *converted = ansi_charset_to_utf8 (plain_charset.c_str(),
                                          body.c_str(), body.size(),
                                          codepage);
  if (is_preview)
    {
      char *buf;
      gpgrt_asprintf (&buf, TEXT_PREVIEW_PLACEHOLDER,
                      isSMIME_m () ? "S/MIME" : "OpenPGP",
                      _("message"),
                      _("Please wait while the message is being verified..."),
                      converted);
      memdbg_alloc (buf);
      xfree (converted);
      converted = buf;
    }
  TRACEPOINT;
  int ret = put_oom_string (m_mailitem, "Body", converted ? converted : "");
  put_oom_int (m_mailitem, "BodyFormat", 1);
  TRACEPOINT;
  xfree (converted);
  if (ret)
    {
      log_error ("%s:%s: Failed to modify body of item.",
                 SRCNAME, __func__);
    }
  TRETURN;
}

void
Mail::updateHeaders_o ()
{
  TSTART;
  if (!m_parser)
    {
      STRANGEPOINT;
      TRETURN;
    }

  const auto subject = m_parser->get_protected_header ("Subject");
  if (!subject.empty ())
    {
      put_oom_string (m_mailitem, "Subject", subject.c_str ());
    }

  const auto to = m_parser->get_protected_header ("To");
  if (!to.empty())
    {
      put_oom_string (m_mailitem, "To", to.c_str ());
    }

  const auto cc = m_parser->get_protected_header ("Cc");
  if (!cc.empty())
    {
      put_oom_string (m_mailitem, "CC", cc.c_str ());
    }

  /* TODO: What about Date ? */

  const auto reply_to = m_parser->get_protected_header ("Reply-To");
  const auto followup_to = m_parser->get_protected_header ("Followup-To");
  if (!reply_to.empty () || !followup_to.empty())
    {
      auto recipients = MAKE_SHARED (get_oom_object (m_mailitem, "ReplyRecipents"));
      if (recipients)
        {
          if (!reply_to.empty ())
            {
              invoke_oom_method_with_string (recipients.get (), "Add",
                                             reply_to.c_str ());
            }
          if (!followup_to.empty ())
            {
              invoke_oom_method_with_string (recipients.get (), "Add",
                                             followup_to.c_str ());
            }
        }
    }

  const auto from = m_parser->get_protected_header ("From");
  if (!from.empty ())
    {
      LPDISPATCH sender = get_oom_object (m_mailitem, "Sender");
      if (!sender)
        {
          log_debug ("%s:%s: Sender not found. From not set.", SRCNAME, __func__);
          TRETURN;
        }

      /* Declare that the address is SMTP */
      put_oom_int (sender, "AddressEntryUserType", 30);
      const auto mail_start = from.find (" <");
      if (mail_start == std::string::npos)
        {
          put_oom_string (sender, "Address", from.c_str ());
          put_oom_string (sender, "Name", "");
        }
      else
        {
          put_oom_string (sender, "Address",
                          GpgME::UserID::addrSpecFromString (from.c_str ()).c_str ());
          put_oom_string (sender, "Name", from.substr (0,
                          mail_start).c_str ());
        }
      gpgol_release (sender);
    }
}

static int parsed_count;

void
Mail::parsingDone_o (bool is_preview)
{
  TSTART;
  TRACEPOINT;
  log_oom ("Mail %p Parsing done for parser num %i: %p",
           this, parsed_count++, m_parser.get());
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
      TRETURN;
    }
  /* Store the results. */
  m_decrypt_result = m_parser->decrypt_result ();
  if (is_preview)
    {
      log_dbg ("Parser is not completely done. In Preview mode.");
    }
  else
    {
      m_verify_result = m_parser->verify_result ();
    }
  /* Handle protected headers */
  updateHeaders_o ();


  m_crypto_flags = 0;
  if (!m_decrypt_result.isNull())
    {
      m_crypto_flags |= 1;
    }
  if (m_verify_result.numSignatures())
    {
      m_crypto_flags |= 2;
    }

  TRACEPOINT;
  updateSigstate ();
  m_needs_wipe = !m_is_send_again;

  TRACEPOINT;
  /* Set categories according to the result. */
  updateCategories_o ();

  TRACEPOINT;
  m_block_html = m_parser->shouldBlockHtml ();

  if (m_block_html)
    {
      // Just to be careful.
      setBlockStatus_m ();
    }

  TRACEPOINT;
  /* Update the body */
  updateBody_o (is_preview);
  TRACEPOINT;

  /* When printing we have already shown the warning. So we
     should not show it again but silently remove any attachments
     that are not hidden before our add_attachments. This
     also fixes an issue that when printing sometimes the
     child mails which are created for preview and print already
     have the decrypted attachments. */
  checkAttachments_o (isPrint ());

  /* Update attachments */
  if (add_attachments_o (m_parser->get_attachments()))
    {
      log_error ("%s:%s: Failed to update attachments.",
                 SRCNAME, __func__);
    }

  if (m_is_send_again)
    {
      log_debug ("%s:%s: I think that this is the send again of a crypto mail.",
                 SRCNAME, __func__);

      /* We no longer want to be treated like a crypto mail. */
      m_type = MSGTYPE_UNKNOWN;
      LPMESSAGE msg = get_oom_base_message (m_mailitem);
      if (!msg)
        {
          TRACEPOINT;
        }
      else
        {
          set_gpgol_draft_info_flags (msg, m_crypto_flags);
          gpgol_release (msg);
        }
      removeOurAttachments_o ();
    }

  installFolderEventHandler_o ();

  if (!is_preview)
    {
      log_debug ("%s:%s: Delayed invalidate to update sigstate.",
                 SRCNAME, __func__);
      CloseHandle(CreateThread (NULL, 0, delayed_invalidate_ui, (LPVOID) 300, 0,
                                NULL));
    }
  TRACEPOINT;
  TRETURN;
}

int
Mail::encryptSignStart_o ()
{
  TSTART;
  if (m_crypt_state != NeedsActualCrypt)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      TRETURN -1;
    }
  int flags = 0;
  if (!needs_crypto_m ())
    {
      TRETURN 0;
    }
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message)
    {
      log_error ("%s:%s: Failed to get base message.",
                 SRCNAME, __func__);
      TRETURN -1;
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
          TRETURN -1;
        }
    }

  m_do_inline = m_is_draft_encrypt ? false :
                m_is_gsuite ? true : opt.inline_pgp;

  GpgME::Protocol proto = opt.enable_smime ? GpgME::UnknownProtocol: GpgME::OpenPGP;

  m_crypter = std::shared_ptr <CryptController> (new CryptController (this, flags & 1,
                                                                      flags & 2,
                                                                      proto));

  // Careful from here on we have to check every
  // error condition with window enabling again.
  disableWindow_o ();
  if (m_crypter->collect_data ())
    {
      log_error ("%s:%s: Crypter for mail %p failed to collect data.",
                 SRCNAME, __func__, this);
      enableWindow ();
      TRETURN -1;
    }

  if (!m_async_crypt_disabled)
    {
      CloseHandle(CreateThread (NULL, 0, do_crypt,
                                (LPVOID) this, 0,
                                NULL));
    }
  else
    {
      log_debug ("%s:%s: Starting sync crypt",
                 SRCNAME, __func__);
      do_crypt (this);
    }
  TRETURN 0;
}

int
Mail::needs_crypto_m () const
{
  TSTART;
  LPMESSAGE message = get_oom_message (m_mailitem);
  int ret;
  if (!message)
    {
      log_error ("%s:%s: Failed to get message.",
                 SRCNAME, __func__);
      TRETURN false;
    }
  ret = get_gpgol_draft_info_flags (message);
  gpgol_release(message);
  TRETURN ret;
}

int
Mail::wipe_o (bool force)
{
  TSTART;
  if (!m_needs_wipe && !force)
    {
      TRETURN 0;
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
          TRETURN -1;
        }
      TRETURN -1;
    }
  else
    {
      put_oom_string (m_mailitem, "Body", "");
    }
  m_needs_wipe = false;
  TRETURN 0;
}

int
Mail::updateOOMData_o (bool for_encryption)
{
  TSTART;
  char *buf = nullptr;
  log_debug ("%s:%s", SRCNAME, __func__);

  for_encryption |= !isCryptoMail();

  if (for_encryption)
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

      if (!m_recipients_set)
        {
          m_cached_recipients = getRecipients_o ();
        }
      else
        {
          log_dbg ("Not updating cached recipients because recipients were "
                   "set explicitly.");
        }
    }
  else
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

      /* We also want to cache sent representing email address so that
         we can use it for verification information. */
      char *buf2 = get_sender_SentRepresentingAddress (m_mailitem);

      if (buf2)
        {
          m_sent_on_behalf = buf2;
          xfree (buf2);
        }
    }

  if (!buf)
    {
      buf = get_sender_SendUsingAccount (m_mailitem, &m_is_gsuite);
    }
  if (!isCryptoMail ())
    {
      /* Try the sender Object and handle the case that someone
         changed the identity but not the account. */
      char *buf2 = get_sender_Sender (m_mailitem);
      char *tmp = buf;
      if (buf && buf2)
        {
          if (*buf == '/' && *buf2 != '/' && *buf2)
            {
              log_dbg ("Send account could not be resolved. Using sender.");
              buf = buf2;
              xfree (tmp);
            }
          else if (*buf && *buf != '/' && *buf2 && *buf2 != '/')
            {
              log_dbg ("SendUsingAccount: '%s' Sender: '%s'",
                       anonstr (buf), anonstr (buf2));
              auto key = KeyCache::instance()->getSigningKey (buf2, GpgME::OpenPGP);
              if (!key.isNull())
                {
                  log_dbg ("Found OpenPGP Key for %s using that identity",
                           anonstr (buf2));
                  buf = buf2;
                  xfree (tmp);
                }
              else if (opt.enable_smime)
                {
                  key = KeyCache::instance()->getSigningKey (buf2, GpgME::CMS);
                  log_dbg ("Found S/MIME Key for %s using that identity",
                           anonstr (buf2));
                  buf = buf2;
                  xfree (tmp);
                }
            }
        }
      else if (!buf)
        {
          buf = buf2;
        }
      else if (buf2 && buf != buf2)
        {
          log_dbg ("Keeping SendUsingAccount %s",
                   anonstr (buf));
          xfree (buf2);
        }
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
      TRETURN -1;
    }
  m_sender = buf;

  if (for_encryption)
    {
      bool sender_is_recipient = false;
      for (const auto &recp: m_cached_recipients)
        {
          if (recp.mbox () == m_sender)
            {
              sender_is_recipient = true;
              log_debug ("%s:%s: Sender already recipient.",
                         SRCNAME, __func__);
              break;
            }
        }
      if (!sender_is_recipient)
        {
          m_cached_recipients.push_back (Recipient (buf,
                                                    Recipient::olOriginator));
        }
    }

  xfree (buf);
  TRETURN 0;
}

std::string
Mail::getSender_o ()
{
  TSTART;
  if (m_sender.empty())
    updateOOMData_o ();
  TRETURN m_sender;
}

std::string
Mail::getSender () const
{
  TSTART;
  TRETURN m_sender;
}

int
Mail::closeAllMails_o ()
{
  TSTART;
  int err = 0;

  /* Detach Folder sinks */
  for (auto fit = s_folder_events_map.begin(); fit != s_folder_events_map.end(); ++fit)
    {
      detach_FolderEvents_sink (fit->second);
      gpgol_release (fit->second);
    }
  s_folder_events_map.clear();


  std::map<LPDISPATCH, Mail *>::iterator it;
  TRACEPOINT;
  gpgol_lock (&mail_map_lock);
  std::map<LPDISPATCH, Mail *> mail_map_copy = s_mail_map;
  gpgol_unlock (&mail_map_lock);
  for (it = mail_map_copy.begin(); it != mail_map_copy.end(); ++it)
    {
      /* XXX For non racy code the is_valid_ptr check should not
         be necessary but we crashed sometimes closing a destroyed
         mail. */
      if (!isValidPtr (it->second))
        {
          log_debug ("%s:%s: Already deleted mail for %p",
                   SRCNAME, __func__, it->first);
          continue;
        }

      if (!it->second->isCryptoMail ())
        {
          continue;
        }
      bool close_failed = false;
      if (closeInspector_o (it->second))
        {
          log_error ("%s:%s: Failed to close mail inspector: %p ",
                     SRCNAME, __func__, it->first);
          close_failed = true;
        }

      if (isValidPtr (it->second))
        {
          log_debug ("%s:%s: Inspector closed for %p closing object.",
                     SRCNAME, __func__, it->first);
          if (it->second->close ())
            {
              log_error ("%s:%s: Failed to close mail itself: %p ",
                         SRCNAME, __func__, it->first);
              close_failed = true;
            }
        }
      else
        {
          log_debug ("%s:%s: Mail gone after inspector close.",
                     SRCNAME, __func__);
          close_failed = false;
        }
      /* Beware: The close code removes our Plaintext from the
         Outlook Object Model and temporary MAPI. If there
         is an error we might put Plaintext into permanent
         storage and leak it to the server. So we have
         an extra safeguard below. The revert is likely
         to fail if close and closeInspector fails but
         to guard against a bug in our close code we
         try it anyway as revert will also try to remove
         the plaintext from memory and restore the original
         message. */
      if (close_failed)
        {
          if (isValidPtr (it->second) && it->second->revert_o ())
            {
              err++;
            }
        }
    }
  TRETURN err;
}
int
Mail::revertAllMails_o ()
{
  TSTART;
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  gpgol_lock (&mail_map_lock);
  auto mail_map_copy = s_mail_map;
  gpgol_unlock (&mail_map_lock);
  for (it = mail_map_copy.begin(); it != mail_map_copy.end(); ++it)
    {
      if (it->second->revert_o ())
        {
          log_error ("Failed to revert mail: %p ", it->first);
          err++;
          continue;
        }

      it->second->setNeedsSave (true);
      if (!invoke_oom_method (it->first, "Save", NULL))
        {
          log_error ("Failed to save reverted mail: %p ", it->second);
          err++;
          continue;
        }
    }
  TRETURN err;
}

int
Mail::wipeAllMails_o ()
{
  TSTART;
  int err = 0;
  std::map<LPDISPATCH, Mail *>::iterator it;
  gpgol_lock (&mail_map_lock);
  auto mail_map_copy = s_mail_map;
  gpgol_unlock (&mail_map_lock);
  for (it = mail_map_copy.begin(); it != mail_map_copy.end(); ++it)
    {
      if (it->second->wipe_o ())
        {
          log_error ("Failed to wipe mail: %p ", it->first);
          err++;
        }
    }
  TRETURN err;
}

int
Mail::revert_o ()
{
  TSTART;
  int err = 0;
  if (!m_processed)
    {
      TRETURN 0;
    }

  m_disable_att_remove_warning = true;

  err = gpgol_mailitem_revert (m_mailitem);
  if (err == -1)
    {
      log_error ("%s:%s: Message revert failed falling back to wipe.",
                 SRCNAME, __func__);
      TRETURN wipe_o ();
    }
  /* We need to reprocess the mail next time around. */
  m_processed = false;
  m_needs_wipe = false;
  m_disable_att_remove_warning = false;
  TRETURN 0;
}

bool
Mail::isSMIME () const
{
  if (!m_is_smime_checked)
    {
      log_dbg ("WARNING: SMIME check before isSMIME_m was called.");
      return false;
    }
  return m_is_smime;
}

bool
Mail::isSMIME_m ()
{
  TSTART;
  msgtype_t msgtype;
  LPMESSAGE message;

  if (m_is_smime_checked)
    {
      TRETURN m_is_smime;
    }

  message = get_oom_message (m_mailitem);

  if (!message)
    {
      log_error ("%s:%s: No message?",
                 SRCNAME, __func__);
      TRETURN false;
    }

  msgtype = mapi_get_message_type (message);
  m_is_smime = msgtype == MSGTYPE_GPGOL_OPAQUE_ENCRYPTED ||
               msgtype == MSGTYPE_GPGOL_OPAQUE_SIGNED ||
               msgtype == MSGTYPE_SMIME;

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

  log_debug ("%s:%s: Detected %s mail",
             SRCNAME, __func__,
             m_is_smime ? "S/MIME" : "not S/MIME");

  TRETURN m_is_smime;
}

static std::string
get_string_o (LPDISPATCH item, const char *str)
{
  TSTART;
  char *buf = get_oom_string (item, str);
  if (!buf)
    {
      TRETURN std::string();
    }
  std::string ret = buf;
  xfree (buf);
  TRETURN ret;
}

std::string
Mail::getSubject_o () const
{
  TSTART;
  TRETURN get_string_o (m_mailitem, "Subject");
}

std::string
Mail::getBody_o () const
{
  TSTART;
  TRETURN get_string_o (m_mailitem, "Body");
}

std::vector<Recipient>
Mail::getRecipients_o () const
{
  TSTART;
  LPDISPATCH recipients = get_oom_object (m_mailitem, "Recipients");
  if (!recipients)
    {
      TRACEPOINT;
      std::vector<std::string>();
    }
  bool err = false;
  auto ret = get_oom_recipients (recipients, &err);
  gpgol_release (recipients);

  if (err)
    {
      log_debug ("%s:%s: Failed to resolve recipients at this time.",
                 SRCNAME, __func__);

    }

  TRETURN ret;
}

int
Mail::closeInspector_o (Mail *mail)
{
  TSTART;
  LPDISPATCH inspector = get_oom_object (mail->item(), "GetInspector");
  HRESULT hr;
  DISPID dispid;
  if (!inspector)
    {
      log_debug ("%s:%s: No inspector.",
                 SRCNAME, __func__);
      TRETURN -1;
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
          TRETURN -1;
        }
    }
  gpgol_release (inspector);
  TRETURN 0;
}


int
Mail::close (bool restoreSMIMEClass)
{
  TSTART;
  wm_after_move_data_t *move_data = nullptr;
  VARIANT aVariant[1];
  DISPPARAMS dispparams;

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = 1;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  if (isSMIME_m ())
    {
      LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
      LPMESSAGE mapi_msg = get_oom_message (m_mailitem);

      if (!mapi_msg)
        {
          log_error ("%s:%s:Failed to obtain mapi message for %p on close.",
                     SRCNAME, __func__, m_mailitem);
        }
      if (m_gpgol_class.empty())
        {
          log_debug ("%s:%s: GpgOL Class empty for S/MIME Mail.",
                     SRCNAME, __func__);
        }

      /* This is strangely important. Because Outlook apparently treats
       * S/MIME Mails differently we need to change our message class
       * here again so that discard changes works properly. */
      if (mapi_msg && !m_gpgol_class.empty ())
        {
          char *current = mapi_get_message_class (mapi_msg);

          if (current && strcmp (current, m_gpgol_class.c_str ()))
            {
              log_debug ("%s:%s:Setting message class to %s on close.",
                         SRCNAME, __func__, m_gpgol_class.c_str ());
              mapi_set_mesage_class (mapi_msg, m_gpgol_class.c_str ());

              if (restoreSMIMEClass)
                {
                  /* After the close we need to change the message class back
                     again so as to not break compatibility with other clients. */
                  move_data = (wm_after_move_data_t *)
                    xmalloc (sizeof (wm_after_move_data_t));

                  size_t entryIDLen = 0;
                  char *entryID = nullptr;
                  entryID = mapi_get_binary_prop (mapi_msg, PR_ENTRYID,
                                                  &entryIDLen);

                  LPDISPATCH folder = get_oom_object (m_mailitem, "Parent");
                  if (folder)
                    {
                      char *name = get_object_name ((LPUNKNOWN) folder);
                      if (!name || strcmp (name, "MAPIFolder"))
                        {
                          log_error ("%s:%s: Failed to obtain folder on close object is %s",
                                     SRCNAME, __func__, name ? name : "(null)");
                        }
                      else
                        {
                          xfree (name);
                          move_data->target_folder = (LPMAPIFOLDER) get_oom_iunknown (
                                                      folder, "MAPIOBJECT");
                          if (!move_data->target_folder)
                            {
                              log_error ("%s:%s: Failed to obtain target folder.",
                                         SRCNAME, __func__);
                              xfree (entryID);
                              xfree (current);
                              xfree (move_data);
                              move_data = nullptr;
                            }
                          else
                            {
                              memdbg_addRef (move_data->target_folder);
                            }
                        }
                    }

                  move_data->entry_id = entryID;
                  move_data->entry_id_len = entryIDLen;
                  move_data->old_class = current;
                }
            }
        }

      if (attachments)
        {
          int count = get_oom_int (attachments, "Count");
          gpgol_release (attachments);
          if (count)
            {
              /* On exchange Outlook sometimes detects that an S/MIME mail
               * is an S/MIME Mail. When this mail is then modified by
               * us and the mail should be moved or closed Outlook will try
               * to save it. This fails and the user gets an error.
               *
               * So we save here, which should not be dangerous as we do not
               * put plaintext in mapi.
               *
               * Still better only do it if it is really
               * necessary as the changed message class can hurt.
               *
               * Tests show no plaintext leaks. The save saves the
               * message class and in that way outlook no longer
               * thinks the mails are S/MIME mails and we can
               * use our own handling. See T4525
               */
              removeCategories_o ();
              HRESULT hr = 0;
              if (mapi_msg)
                {
                  log_debug ("%s:%s: MAPI Save for: %p",
                             SRCNAME, __func__, m_mailitem);
                  mapi_msg->SaveChanges (KEEP_OPEN_READWRITE);
                }
              if (!mapi_msg || hr)
                {
                  log_error ("%s:%s: Failed to save mapi for %p hr=%#lx",
                             SRCNAME, __func__, this, hr);
                }
              /* In case the mail is still visible in a different window */
              updateCategories_o ();
            }
        }

      gpgol_release (mapi_msg);
    }

  log_oom ("%s:%s: Invoking close for: %p",
                 SRCNAME, __func__, m_mailitem);
  setCloseTriggered (true);
  int rc = invoke_oom_method_with_parms (m_mailitem, "Close",
                                         nullptr, &dispparams);

  if (move_data)
    {
      log_debug ("%s:%s:Restoring message class to %s after close.",
                 SRCNAME, __func__, move_data->old_class);
      do_in_ui_thread_async (AFTER_MOVE, move_data, 0);
    }
  setCloseTriggered (false);

  if (!rc)
    {
      /* Saveguard against oom writes when our data is in OOM. This
       * can happen when the mail was opened in a new window. But
       * also shown in the preview.
       * We get a close event and discard changes but the data is
       * still in oom because it is still visible in the opened
       * window.
       *
       * In that case we may not write! Otherwise the plaintext
       * might be leaked back to the server if the folder is synced.
       * */
      char *body = get_oom_string (m_mailitem, "Body");
      LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");

      if (body && strlen (body))
        {
          log_debug ("%s:%s: Close successful. But body found. "
                     "Mail still open.",
                     SRCNAME, __func__);
        }
      else if (count_visible_attachments (attachments))
        {
              log_debug ("%s:%s: Close successful. But attachments found. "
                         "Mail still open.",
                         SRCNAME, __func__);
        }
      else
        {
           setPassWrite (true);
           log_debug ("%s:%s: Close successful. Next write may pass.",
                      SRCNAME, __func__);
        }
      gpgol_release (attachments);
      xfree (body);
    }
  log_oom ("%s:%s: returned from close",
                 SRCNAME, __func__);
  TRETURN rc;
}

void
Mail::setCloseTriggered (bool value)
{
  TSTART;
  m_close_triggered = value;
  TRETURN;
}

bool
Mail::getCloseTriggered () const
{
  TSTART;
  TRETURN m_close_triggered;
}

static const UserID
get_uid_for_sender (const Key &k, const char *sender)
{
  TSTART;
  UserID ret;

  if (!sender)
    {
      TRETURN ret;
    }

  if (!k.numUserIDs())
    {
      log_debug ("%s:%s: Key without uids",
                 SRCNAME, __func__);
      TRETURN ret;
    }

  for (const auto uid: k.userIDs())
    {
      if (!uid.email() || !*(uid.email()))
        {
          /* This happens for S/MIME a lot */
          log_debug ("%s:%s: skipping uid without email.",
                     SRCNAME, __func__);
          continue;
        }
      auto normalized_uid = uid.addrSpec();
      auto normalized_sender = UserID::addrSpecFromString(sender);

      if (normalized_sender.empty() || normalized_uid.empty())
        {
          log_error ("%s:%s: normalizing '%s' or '%s' failed.",
                     SRCNAME, __func__, anonstr (uid.email()),
                     anonstr (sender));
          continue;
        }
      if (normalized_sender == normalized_uid)
        {
          ret = uid;
        }
    }
  TRETURN ret;
}

void
Mail::updateSigstate ()
{
  TSTART;
  std::string sender = getSender ();

  if (sender.empty())
    {
      log_error ("%s:%s:%i", SRCNAME, __func__, __LINE__);
      TRETURN;
    }

  if (m_verify_result.isNull())
    {
      log_debug ("%s:%s: No verify result.",
                 SRCNAME, __func__);
      TRETURN;
    }

  if (m_verify_result.error())
    {
      log_debug ("%s:%s: verify error.",
                 SRCNAME, __func__);
      TRETURN;
    }

  for (const auto sig: m_verify_result.signatures())
    {
      m_is_signed = true;
      const auto key = KeyCache::instance ()->getByFpr (sig.fingerprint(),
                                                        true);
      m_uid = get_uid_for_sender (key, sender.c_str());

      if (m_uid.isNull() && !m_sent_on_behalf.empty ())
        {
          m_uid = get_uid_for_sender (key, m_sent_on_behalf.c_str ());
          if (!m_uid.isNull())
            {
              log_debug ("%s:%s: Using sent on behalf '%s' instead of '%s'",
                         SRCNAME, __func__, anonstr (m_sent_on_behalf.c_str()),
                         anonstr (sender.c_str ()));
            }
        }
      /* Sigsum valid or green is somehow not set in this case.
       * Which is strange as AFAIK this worked in the past. */
      if ((sig.summary() & Signature::Summary::Valid) &&
          m_uid.origin() == GpgME::Key::OriginWKD &&
          (sig.validity() == Signature::Validity::Unknown ||
           sig.validity() == Signature::Validity::Marginal))
        {
          // WKD is a shortcut to Level 2 trust.
          log_debug ("%s:%s: Unknown or marginal from WKD -> Level 2",
                     SRCNAME, __func__);
         }
      else if (m_uid.isNull() || (sig.validity() != Signature::Validity::Marginal &&
          sig.validity() != Signature::Validity::Full &&
          sig.validity() != Signature::Validity::Ultimate))
        {
          /* For our category we only care about trusted sigs. And
          the UID needs to match.*/
          continue;
        }
      else if (sig.validity() == Signature::Validity::Marginal)
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
      log_debug ("%s:%s: Classified sender as verified uid validity: %i origin: %i",
                 SRCNAME, __func__, m_uid.validity(), m_uid.origin());
      m_sig = sig;
      m_is_valid = true;
      TRETURN;
    }

  log_debug ("%s:%s: No signature with enough trust. Using first",
             SRCNAME, __func__);
  m_sig = m_verify_result.signature(0);
  TRETURN;
}

bool
Mail::isValidSig () const
{
  TSTART;
  TRETURN m_is_valid;
}

void
Mail::removeCategories_o ()
{
  TSTART;
  if (!m_store_id.empty () && !m_verify_category.empty ())
    {
      log_oom ("%s:%s: Unreffing verify category",
                       SRCNAME, __func__);
      CategoryManager::instance ()->removeCategory (this,
                                                    m_verify_category);
    }
  if (!m_store_id.empty () && !m_decrypt_result.isNull())
    {
      log_oom ("%s:%s: Unreffing dec category",
                       SRCNAME, __func__);
      CategoryManager::instance ()->removeCategory (this,
                                CategoryManager::getEncMailCategory ());
    }
  if (m_is_junk)
    {
      log_oom ("%s:%s: Unreffing junk category",
                       SRCNAME, __func__);
      CategoryManager::instance ()->removeCategory (this,
                                CategoryManager::getJunkMailCategory ());
    }
  TRETURN;
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
  TSTART;
  HWND wnd = get_active_hwnd ();
  static std::vector<HWND> resized_windows;
  if(std::find(resized_windows.begin(), resized_windows.end(), wnd) != resized_windows.end()) {
      /* We only need to do this once per window. XXX But sometimes we also
         need to do this once per view of the explorer. So for now this might
         break but we reduce the flicker. A better solution would be to find
         the current view and track that. */
      TRETURN;
  }

  if (!wnd)
    {
      TRACEPOINT;
      TRETURN;
    }
  RECT oldpos;
  if (!GetWindowRect (wnd, &oldpos))
    {
      TRACEPOINT;
      TRETURN;
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
      TRETURN;
    }

  if (!SetWindowPos (wnd, nullptr,
                     (int)oldpos.left,
                     (int)oldpos.top,
                     (int)oldpos.right - oldpos.left,
                     (int)oldpos.bottom - oldpos.top, 0))
    {
      TRACEPOINT;
      TRETURN;
    }
  resized_windows.push_back(wnd);
  TRETURN;
}

#if 0
static std::string
pretty_id (const char *keyId)
{
  /* Three spaces, four quads and a NULL */
  char buf[20];
  buf[19] = '\0';
  if (!keyId)
    {
      return std::string ("null");
    }
  size_t len = strlen (keyId);
  if (!len)
    {
      return std::string ("empty");
    }
  if (len < 16)
    {
      return std::string (_("Invalid Key"));
    }
  const char *p = keyId + (len - 16);
  int j = 0;
  for (size_t i = 0; i < 16; i++)
    {
      if (i && i % 4 == 0)
        {
          buf[j++] = ' ';
        }
      buf[j++] = *(p + i);
    }
  return std::string (buf);
}
#endif

void
Mail::updateCategories_o ()
{
  TSTART;

  if (m_printing)
    {
      log_debug ("%s:%s: Not updating categories as we are printing.",
                 SRCNAME, __func__);
      return;
    }

  auto mngr = CategoryManager::instance ();
  if (isValidSig ())
    {
      char *buf;
      /* Resolve to the primary fingerprint */
#if 0
      const auto sigKey = KeyCache::instance ()->getByFpr (m_sig.fingerprint (),
                                                           true);
      const char *sigFpr;
      if (sigKey.isNull())
        {
          sigFpr = m_sig.fingerprint ();
        }
      else
        {
          sigFpr = sigKey.primaryFingerprint ();
        }
#endif
      /* If m_uid addrSpec would not return a result we would never
       * have gotten the UID. */
      int lvl = get_signature_level ();

      /* TRANSLATORS: The first placeholder is for tranlsation of "Level".
         The second one is for the level number. The third is for the
         translation of "trust in" and the last one is for the mail
         address used for verification. The result is used as the
         text on the green bar for signed mails. e.g.:
         "GpgOL: Level 3 trust in 'john.doe@example.org'" */
      gpgrt_asprintf (&buf, "GpgOL: %s %i %s '%s'", _("Level"), lvl,
                      _("trust in"),
                      m_uid.addrSpec ().c_str ());
      memdbg_alloc (buf);

      int color = 0;
      if (lvl == 2)
        {
          color = 7; /* Olive */
        }
      if (lvl == 3)
        {
          color = 5; /* Green */
        }
      if (lvl == 4)
        {
          color = 20; /* Dark Green */
        }
      m_store_id = mngr->addCategoryToMail (this, buf, color);
      m_verify_category = buf;
      xfree (buf);
    }
  else
    {
      remove_category (m_mailitem, "GpgOL: ", false);
    }

  if (!m_decrypt_result.isNull())
    {
      const auto id = mngr->addCategoryToMail (this,
                                 CategoryManager::getEncMailCategory (),
                                 8 /* Blue */);
      if (m_store_id.empty())
        {
          m_store_id = id;
        }
      if (m_store_id != id)
        {
          log_error ("%s:%s unexpected store mismatch "
                     "between '%s' and dec cat '%s'",
                     SRCNAME, __func__, m_store_id.c_str(), id.c_str());
        }
    }
  else
    {
      /* As a small safeguard against fakes we remove our
         categories */
      remove_category (m_mailitem,
                       CategoryManager::getEncMailCategory ().c_str (),
                       true);
    }

  resize_active_window();

  TRETURN;
}

bool
Mail::isSigned () const
{
  TSTART;
  TRETURN m_verify_result.numSignatures() > 0;
}

bool
Mail::isEncrypted () const
{
  TSTART;
  TRETURN !m_decrypt_result.isNull();
}

int
Mail::setUUID_o ()
{
  TSTART;
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
      TRETURN -1;
    }
  if (m_uuid.empty())
    {
      m_uuid = uuid;
      Mail *other = getMailForUUID (uuid);
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

      gpgol_lock (&uid_map_lock);
      s_uid_map.insert (std::pair<std::string, Mail *> (m_uuid, this));
      gpgol_unlock (&uid_map_lock);
      log_debug ("%s:%s: uuid for %p is now %s",
                 SRCNAME, __func__, this,
                 m_uuid.c_str());
    }
  xfree (uuid);
  TRETURN 0;
}

/* TRETURNs 2 if the userid is ultimately trusted.

   TRETURNs 1 if the userid is fully trusted but has
   a signature by a key for which we have a secret
   and which is ultimately trusted. (Direct trust)

   0 otherwise */
static int
level_4_check (const UserID &uid)
{
  TSTART;
  if (uid.isNull())
    {
      TRETURN 0;
    }
  if (uid.validity () == UserID::Validity::Ultimate)
    {
      TRETURN 2;
    }
  if (uid.validity () == UserID::Validity::Full)
    {
      const auto ultimate_keys = KeyCache::instance()->getUltimateKeys ();
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
          for (const auto secKey: ultimate_keys)
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
                                 SRCNAME, __func__, anonstr (signer_uid_str),
                                 anonstr (sig_uid_str),
                                 anonstr (secKeyID));
                      TRETURN 1;
                    }
                }
            }
        }
    }
  TRETURN 0;
}

std::string
Mail::getCryptoSummary () const
{
  TSTART;
  const int level = get_signature_level ();

  bool enc = isEncrypted ();
  if (level == 4 && enc)
    {
      TRETURN _("Security Level 4");
    }
  if (level == 4)
    {
      TRETURN _("Trust Level 4");
    }
  if (level == 3 && enc)
    {
      TRETURN _("Security Level 3");
    }
  if (level == 3)
    {
      TRETURN _("Trust Level 3");
    }
  if (level == 2 && enc)
    {
      TRETURN _("Security Level 2");
    }
  if (level == 2)
    {
      TRETURN _("Trust Level 2");
    }
  if (enc)
    {
      TRETURN _("Encrypted");
    }
  if (isSigned ())
    {
      /* Even if it is signed, if it is not validly
         signed it's still completly insecure as anyone
         could have signed this. So we avoid the label
         "signed" here as this word already implies some
         security. */
      TRETURN _("Insecure");
    }
  TRETURN _("Insecure");
}

std::string
Mail::getCryptoOneLine () const
{
  TSTART;
  bool sig = isSigned ();
  bool enc = isEncrypted ();
  if (sig || enc)
    {
      if (sig && enc)
        {
          TRETURN _("Signed and encrypted message");
        }
      else if (sig)
        {
          TRETURN _("Signed message");
        }
      else if (enc)
        {
          TRETURN _("Encrypted message");
        }
    }
  TRETURN _("Insecure message");
}

std::string
Mail::getCryptoDetails_o ()
{
  TSTART;
  std::string message;

  /* No signature with keys but error */
  if (!isEncrypted () && !isSigned () && m_verify_result.error())
    {
      message = _("You cannot be sure who sent, "
                  "modified and read the message in transit.");
      message += "\n\n";
      message += _("The message was signed but the verification failed with:");
      message += "\n";
      message += m_verify_result.error().asString();
      TRETURN message;
    }
  /* No crypo, what are we doing here? */
  if (!isEncrypted () && !isSigned ())
    {
      TRETURN _("You cannot be sure who sent, "
               "modified and read the message in transit.");
    }
  /* Handle encrypt only */
  if (isEncrypted () && !isSigned ())
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
      TRETURN message;
    }

  bool keyFound = true;
  const auto sigKey = KeyCache::instance ()->getByFpr (m_sig.fingerprint (),
                                                       true);
  bool isOpenPGP = sigKey.isNull() ? !isSMIME_m () :
                   sigKey.protocol() == Protocol::OpenPGP;
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

      if (four_check == 2 && sigKey.hasSecret ())
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
          TRETURN message;
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
                      sigKey.issuerName());
      memdbg_alloc (buf);
      message = buf;
      xfree (buf);
    }
  else if (level == 2 && m_uid.origin () == GpgME::Key::OriginWKD)
    {
      message = _("The mail provider of the recipient served this key.");
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
      memdbg_alloc (buf);
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
          memdbg_alloc (buf);
          xfree (time);
          message = buf;
          xfree (buf);
        }
    }
  else
    {
      /* Now we are in level 0, this could be a technical problem, no key
         or just unkown. */
      message = isEncrypted () ? _("But the sender address is not trustworthy because:") :
                                  _("The sender address is not trustworthy because:");
      message += "\n";
      keyFound = !(m_sig.summary() & Signature::Summary::KeyMissing);

      bool general_problem = true;
      /* First the general stuff. */
      if (m_sig.summary() & Signature::Summary::Red)
        {
            message += _("The signature is invalid: \n");
            if (m_sig.status().code() == GPG_ERR_BAD_SIGNATURE)
              {
                message += std::string("\n") + _("The signature does not match.");
                return message;
              }
        }
      else if (m_sig.summary() & Signature::Summary::SysError ||
               m_verify_result.numSignatures() < 1)
        {
          message += _("There was an error verifying the signature.\n");
          const auto err = m_sig.status ();
          if (err)
            {
              message += err.asString () + std::string ("\n");
            }
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
                          getSender_o ().c_str());
          memdbg_alloc (buf);
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
       if (isSigned ())
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
       if (isEncrypted ())
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
  TRETURN message;
}

int
Mail::get_signature_level () const
{
  TSTART;
  if (!m_is_signed)
    {
      TRETURN 0;
    }

  if (m_uid.isNull ())
    {
      /* No m_uid matches our sender. */
      TRETURN 0;
    }

  if (m_is_valid && (m_uid.validity () == UserID::Validity::Ultimate ||
      (m_uid.validity () == UserID::Validity::Full &&
      level_4_check (m_uid))) && (!in_de_vs_mode () || m_sig.isDeVs()))
    {
      TRETURN 4;
    }
  if (m_is_valid && m_uid.validity () == UserID::Validity::Full &&
      (!in_de_vs_mode () || m_sig.isDeVs()))
    {
      TRETURN 3;
    }
  if (m_is_valid)
    {
      TRETURN 2;
    }
  if (m_sig.validity() == Signature::Validity::Marginal)
    {
      TRETURN 1;
    }
  if (m_sig.summary() & Signature::Summary::TofuConflict ||
      m_uid.tofuInfo().validity() == TofuInfo::Conflict)
    {
      TRETURN 0;
    }
  TRETURN 0;
}

int
Mail::getCryptoIconID () const
{
  TSTART;
  int level = get_signature_level ();
  int offset = isEncrypted () ? ENCRYPT_ICON_OFFSET : 0;
  TRETURN IDI_LEVEL_0 + level + offset;
}

const char*
Mail::getSigFpr () const
{
  TSTART;
  if (!m_is_signed || m_sig.isNull())
    {
      TRETURN nullptr;
    }
  TRETURN m_sig.fingerprint();
}

/** Try to locate the keys for all recipients */
void
Mail::locateKeys_o ()
{
  TSTART;
  if (m_locate_in_progress)
    {
      /** XXX
        The strangest thing seems to happen here:
        In get_recipients the lookup for "AddressEntry" on
        an unresolved address might cause network traffic.

        So Outlook somehow "detaches" this call and keeps
        processing window messages while the call is running.

        So our do_delayed_locate might trigger a second locate.
        If we access the OOM in this call while we access the
        same object in the blocked "detached" call we crash.
        (T3931)

        After the window message is handled outlook retunrs
        in the original lookup.

        A better fix here might be a non recursive lock
        of the OOM. But I expect that if we lock the handling
        of the Windowmessage we might deadlock.
        */
      log_debug ("%s:%s: Locate for %p already in progress.",
                 SRCNAME, __func__, this);
      TRETURN;
    }
  m_locate_in_progress = true;

  Addressbook::check_o (this);

  if (opt.autoresolve)
    {
      // First update oom data to have recipients and sender updated.
      updateOOMData_o ();
      KeyCache::instance()->startLocateSecret (getSender_o ().c_str (), this);
      KeyCache::instance()->startLocate (getSender_o ().c_str (), this);
      KeyCache::instance()->startLocate (getCachedRecipientAddresses (), this);
    }

  autosecureCheck ();

  m_locate_in_progress = false;
  TRETURN;
}

bool
Mail::isHTMLAlternative () const
{
  TSTART;
  TRETURN m_is_html_alternative;
}

char *
Mail::takeCachedHTMLBody ()
{
  TSTART;
  char *ret = m_cached_html_body;
  m_cached_html_body = nullptr;
  TRETURN ret;
}

char *
Mail::takeCachedPlainBody ()
{
  TSTART;
  char *ret = m_cached_plain_body;
  m_cached_plain_body = nullptr;
  TRETURN ret;
}

int
Mail::getCryptoFlags () const
{
  TSTART;
  TRETURN m_crypto_flags;
}

void
Mail::setNeedsEncrypt (bool value)
{
  TSTART;
  m_needs_encrypt = value;
  TRETURN;
}

bool
Mail::getNeedsEncrypt () const
{
  TSTART;
  TRETURN m_needs_encrypt;
}

std::vector<Recipient>
Mail::getCachedRecipients ()
{
  TSTART;
  TRETURN m_cached_recipients;
}

void
Mail::setRecipients (const std::vector<Recipient> &recps)
{
  TSTART;
  m_recipients_set = !recps.empty ();
  m_cached_recipients = recps;
  TRETURN;
}


std::vector<std::string>
Mail::getCachedRecipientAddresses ()
{
  TSTART;
  std::vector <std::string> ret;
  for (const auto &recp: m_cached_recipients)
    {
      ret.push_back (recp.mbox());
    }
  return ret;
}

void
Mail::appendToInlineBody (const std::string &data)
{
  TSTART;
  m_inline_body += data;
  TRETURN;
}

int
Mail::inlineBodyToBody_o ()
{
  TSTART;
  if (!m_crypter)
    {
      log_error ("%s:%s: No crypter.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  const auto body = m_crypter->get_inline_data ();
  if (body.empty())
    {
      TRETURN 0;
    }

  /* For inline we always work with UTF-8 */
  put_oom_int (m_mailitem, "InternetCodepage", 65001);

  int ret = put_oom_string (m_mailitem, "Body",
                            body.c_str ());
  TRETURN ret;
}

void
Mail::updateCryptMAPI_m ()
{
  TSTART;
  log_debug ("%s:%s: Update crypt mapi",
             SRCNAME, __func__);
  if (m_crypt_state != NeedsUpdateInMAPI)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      TRETURN;
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
          TRETURN;
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

  /** If sync we need the crypter in update_crypt_oom */
  if (!isAsyncCryptDisabled ())
    {
      // We don't need the crypter anymore.
      resetCrypter ();
    }
  TRETURN;
}

/** Checks in OOM if the body is either
  empty or contains the -----BEGIN tag.
  pair.first -> true if body starts with -----BEGIN
  pair.second -> true if body is empty. */
static std::pair<bool, bool>
has_crypt_or_empty_body_oom (Mail *mail)
{
  TSTART;
  auto body = mail->getBody_o ();
  std::pair<bool, bool> ret;
  ret.first = false;
  ret.second = false;
  ltrim (body);
  if (body.size() > 10 && !strncmp (body.c_str(), "-----BEGIN", 10))
    {
      ret.first = true;
      TRETURN ret;
    }
  if (!body.size())
    {
      ret.second = true;
    }
  else
    {
      log_data ("%s:%s: Body found in %p : \"%s\"",
                       SRCNAME, __func__, mail, body.c_str ());
    }
  TRETURN ret;
}

void
Mail::updateCryptOOM_o ()
{
  TSTART;
  log_debug ("%s:%s: Update crypt oom for %p",
             SRCNAME, __func__, this);
  if (m_crypt_state != NeedsUpdateInOOM)
    {
      log_debug ("%s:%s: invalid state %i",
                 SRCNAME, __func__, m_crypt_state);
      resetCrypter ();
      resetRecipients ();
      TRETURN;
    }

  if (getDoPGPInline ())
    {
      if (inlineBodyToBody_o ())
        {
          log_error ("%s:%s: Inline body to body failed %p.",
                     SRCNAME, __func__, this);
          gpgol_bug (get_active_hwnd(), ERR_INLINE_BODY_TO_BODY);
          m_crypt_state = NoCryptMail;
          TRETURN;
        }
    }

  if (m_crypter->get_protocol () == GpgME::CMS && m_crypter->is_encrypter ())
    {
      /* We put the PIDNameContentType headers here for exchange
         because this is the only way we found to inject the
         smime-type. */
      if (put_pa_string (m_mailitem,
                         PR_PIDNameContentType_DASL,
                         "application/pkcs7-mime;smime-type=\"enveloped-data\";name=smime.p7m"))
        {
          log_debug ("%s:%s: Failed to put PIDNameContentType for %p.",
                     SRCNAME, __func__, this);
        }
    }

  /** When doing async update_crypt_mapi follows and needs
    the crypter. */
  if (isAsyncCryptDisabled ())
    {
      resetCrypter ();
      resetRecipients ();
    }

  const auto pair = has_crypt_or_empty_body_oom (this);
  if (pair.first)
    {
      log_debug ("%s:%s: Looks like inline body. You can pass %p.",
                 SRCNAME, __func__, this);
      m_crypt_state = WantsSendInline;
      TRETURN;
    }

  /* Draft encryption we do not want to wipe the oom but we have
     to modify it to trigger the wirte / second after write. */
  if (m_is_draft_encrypt)
    {
      char *subject = get_oom_string (m_mailitem, "Subject");
      put_oom_string (m_mailitem, "Subject", subject ? subject : "");
      xfree (subject);
    }

  // We are in MIME land. Wipe the body.
  if (!m_is_draft_encrypt && wipe_o (true))
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
      TRETURN;
    }
  m_crypt_state = NeedsSecondAfterWrite;
  TRETURN;
}

void
Mail::enableWindow ()
{
  TSTART;
  if (!m_window)
    {
      log_error ("%s:%s:enable window which was not disabled",
                 SRCNAME, __func__);
    }
  log_debug ("%s:%s: enable window %p",
             SRCNAME, __func__, m_window);

  EnableWindow (m_window, TRUE);
  TRETURN;
}

void
Mail::disableWindow_o ()
{
  TSTART;
  m_window = get_active_hwnd ();
  log_debug ("%s:%s: disable window %p",
             SRCNAME, __func__, m_window);

  EnableWindow (m_window, FALSE);
  TRETURN;
}

bool
Mail::isActiveInlineResponse_o ()
{
  const auto subject = getSubject_o ();
  bool ret = false;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();
  if (!app)
    {
      TRACEPOINT;
      TRETURN false;
    }

  LPDISPATCH explorer = get_oom_object (app, "ActiveExplorer");

  if (!explorer)
    {
      TRACEPOINT;
      TRETURN false;
    }

  LPDISPATCH inlineResponse = get_oom_object (explorer, "ActiveInlineResponse");
  gpgol_release (explorer);

  if (!inlineResponse)
    {
      TRETURN false;
    }

  // We have inline response
  // Check if we are it. It's a bit naive but meh. Worst case
  // is that we think inline response too often and do sync
  // crypt where we could do async crypt.
  char * inlineSubject = get_oom_string (inlineResponse, "Subject");
  gpgol_release (inlineResponse);

  if (inlineResponse && !subject.empty() && !strcmp (subject.c_str (), inlineSubject))
    {
      log_debug ("%s:%s: Detected inline response for '%p'",
                 SRCNAME, __func__, this);
      ret = true;
    }

  xfree (inlineSubject);
  TRETURN ret;
}

bool
Mail::checkSyncCrypto_o ()
{
  TSTART;
  /* Async sending is known to cause instabilities. So we keep
     a hidden option to disable it. */
  if (opt.sync_enc)
    {
      m_async_crypt_disabled = true;
      TRETURN m_async_crypt_disabled;
    }

  m_async_crypt_disabled = false;

  const auto subject = getSubject_o ();

  /* Check for an empty subject. Otherwise the question for it
     might be hidden behind our overlay. */
  if (subject.empty())
    {
      log_debug ("%s:%s: Detected empty subject. "
                 "Disabling async crypt due to T4150.",
                 SRCNAME, __func__);
      m_async_crypt_disabled = true;
      TRETURN m_async_crypt_disabled;
    }

  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (attachments)
    {
      /* This is horrible. But. For some kinds of attachments (we
         got reports about Office attachments the write in the
         send event triggered by our crypto done code fails with
         an exception. There does not appear to be a detectable
         pattern when this happens.
         As we can't be sure and do not know for which attachments
         this really happens we do not use async crypt for any
         mails with attachments. :-/
         Better be save (not crash) instead of nice (async).

         TODO: Figure this out.

         The log goes like this. We pass the send event. That triggers
         a write, which we pass. And then that fails. So it looks like
         moving to Outbox fails. Because we can save as much as we
         like before that.

         Using the IMessage::SubmitMessage MAPI interface works, but
         as it is unstable in our current implementation we do not
         want to use it.

         mailitem-events.cpp:Invoke: Passing send event for mime-encrypted message 12B7C6E0.
         application-events.cpp:Invoke: Unhandled Event: f002
         mailitem-events.cpp:Invoke: Write : 0ED4D058
         mailitem-events.cpp:Invoke: Passing write event.
         oomhelp.cpp:invoke_oom_method_with_parms_type: Method 'Send' invokation failed: 0x80020009
         oomhelp.cpp:dump_excepinfo: Exception:
         wCode: 0x1000
         wReserved: 0x0
         source: Microsoft Outlook
         desc: The operation failed.  The messaging interfaces have returned an unknown error. If the problem persists, restart Outlook.
         help: null
         helpCtx: 0x0
         deferredFill: 00000000
         scode: 0x80040119
      */

      int count = get_oom_int (attachments, "Count");
      gpgol_release (attachments);

      if (count)
        {
          m_async_crypt_disabled = true;
          log_debug ("%s:%s: Detected attachments. "
                     "Disabling async crypt due to T4131.",
                     SRCNAME, __func__);
          TRETURN m_async_crypt_disabled;
        }
   }

  if (isActiveInlineResponse_o ())
    {
      m_async_crypt_disabled = true;
    }

  TRETURN m_async_crypt_disabled;
}

// static
Mail *
Mail::getLastMail ()
{
  TSTART;
  if (!s_last_mail || !isValidPtr (s_last_mail))
    {
      s_last_mail = nullptr;
    }
  TRETURN s_last_mail;
}

// static
void
Mail::clearLastMail ()
{
  TSTART;
  s_last_mail = nullptr;
  TRETURN;
}

// static
void
Mail::locateAllCryptoRecipients_o ()
{
  TSTART;
  gpgol_lock (&mail_map_lock);
  std::map<LPDISPATCH, Mail *>::iterator it;
  auto mail_map_copy = s_mail_map;
  gpgol_unlock (&mail_map_lock);
  for (it = mail_map_copy.begin(); it != mail_map_copy.end(); ++it)
    {
      if (it->second->needs_crypto_m ())
        {
          it->second->locateKeys_o ();
        }
    }
  TRETURN;
}

int
Mail::removeAllAttachments_o ()
{
  TSTART;
  int ret = 0;
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      TRACEPOINT;
      TRETURN 0;
    }
  int count = get_oom_int (attachments, "Count");
  LPDISPATCH to_delete[count];

  /* Populate the array so that we don't get in an index mess */
  for (int i = 1; i <= count; i++)
    {
      auto item_str = std::string("Item(") + std::to_string (i) + ")";
      to_delete[i-1] = get_oom_object (attachments, item_str.c_str());
    }
  gpgol_release (attachments);

  /* Now delete all attachments */
  for (int i = 0; i < count; i++)
    {
      LPDISPATCH attachment = to_delete[i];

      if (!attachment)
        {
          log_error ("%s:%s: No such attachment %i",
                     SRCNAME, __func__, i);
          ret = -1;
        }

      /* Delete the attachments that are marked to delete */
      if (invoke_oom_method (attachment, "Delete", NULL))
        {
          log_error ("%s:%s: Deleting attachment %i",
                     SRCNAME, __func__, i);
          ret = -1;
        }
      gpgol_release (attachment);
    }
  TRETURN ret;
}

int
Mail::removeOurAttachments_o ()
{
  TSTART;
  LPDISPATCH attachments = get_oom_object (m_mailitem, "Attachments");
  if (!attachments)
    {
      TRACEPOINT;
      TRETURN 0;
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
  TRETURN ret;
}

/* We are very verbose because if we fail it might mean
   that we have leaked plaintext -> critical. */
bool
Mail::hasCryptedOrEmptyBody_o ()
{
  TSTART;
  const auto pair = has_crypt_or_empty_body_oom (this);

  if (pair.first /* encrypted marker */)
    {
      log_debug ("%s:%s: Crypt Marker detected in OOM body. Return true %p.",
                 SRCNAME, __func__, this);
      TRETURN true;
    }

  if (!pair.second)
    {
      log_debug ("%s:%s: Unexpected content detected. Return false %p.",
                 SRCNAME, __func__, this);
      TRETURN false;
    }

  // Pair second == true (is empty) can happen on OOM error.
  LPMESSAGE message = get_oom_base_message (m_mailitem);
  if (!message && pair.second)
    {
      if (message)
        {
          gpgol_release (message);
        }
      TRETURN true;
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
      TRETURN true;
    }
  if (r_nbytes > 10 && !strncmp (mapi_body, "-----BEGIN", 10))
    {
      // Body is crypt.
      log_debug ("%s:%s: MAPI Crypt marker detected. Return true. %p.",
                 SRCNAME, __func__, this);
      xfree (mapi_body);
      TRETURN true;
    }

  xfree (mapi_body);

  log_debug ("%s:%s: Found mapi body. Return false. %p.",
             SRCNAME, __func__, this);

  TRETURN false;
}

std::string
Mail::getVerificationResultDump ()
{
  TSTART;
  std::stringstream ss;
  ss << m_verify_result;
  TRETURN ss.str();
}

void
Mail::setBlockStatus_m ()
{
  TSTART;
  SPropValue prop;

  LPMESSAGE message = get_oom_base_message (m_mailitem);

  prop.ulPropTag = PR_BLOCK_STATUS;
  prop.Value.l = 1;
  HRESULT hr = message->SetProps (1, &prop, NULL);

  if (hr)
    {
      log_error ("%s:%s: can't set block value: hr=%#lx\n",
                 SRCNAME, __func__, hr);
    }

  gpgol_release (message);
  TRETURN;
}

void
Mail::setBlockHTML (bool value)
{
  TSTART;
  m_block_html = value;
  TRETURN;
}

void
Mail::incrementLocateCount ()
{
  TSTART;
  m_locate_count++;
  TRETURN;
}

void
Mail::decrementLocateCount ()
{
  TSTART;
  m_locate_count--;

  if (m_locate_count < 0)
    {
      log_error ("%s:%s: locate count mismatch.",
                 SRCNAME, __func__);
      m_locate_count = 0;
    }
  if (!m_locate_count)
    {
      autosecureCheck ();
    }
  TRETURN;
}

void
Mail::autosecureCheck ()
{
  TSTART;
  if (!opt.autosecure || !opt.autoresolve || m_manual_crypto_opts ||
      m_locate_count)
    {
      TRETURN;
    }
  bool ret = KeyCache::instance()->isMailResolvable (this);

  log_debug ("%s:%s: status %i",
             SRCNAME, __func__, ret);

  /* As we are safe to call at any time, because we need
   * to be triggered by the locator threads finishing we
   * need to actually set the draft info flags in
   * the ui thread. */
  do_in_ui_thread_async (ret ? DO_AUTO_SECURE : DONT_AUTO_SECURE,
                         this);
  TRETURN;
}

void
Mail::setDoAutosecure_m (bool value)
{
  TSTART;
  TRACEPOINT;
  LPMESSAGE msg = get_oom_base_message (m_mailitem);

  if (!msg)
    {
      TRACEPOINT;
      TRETURN;
    }
  /* We need to set a uuid so that autosecure can
     be disabled manually */
  setUUID_o ();

  int old_flags = get_gpgol_draft_info_flags (msg);
  if (old_flags && m_first_autosecure_check &&
      /* Someone with always sign and autosecure active
       * will want to get autoencryption. */
      !(old_flags == 2 && opt.sign_default))
    {
      /* They were set explicily before us. This can be
       * because they were a draft (which is bad) or
       * because they are a reply/forward to a crypto mail
       * or because there are conflicting settings. */
      log_debug ("%s:%s: Mail %p had already flags set.",
                 SRCNAME, __func__, m_mailitem);
      m_first_autosecure_check = false;
      m_manual_crypto_opts = true;
      gpgol_release (msg);
      TRETURN;
    }
  m_first_autosecure_check = false;
  set_gpgol_draft_info_flags (msg, value ? 3 : opt.sign_default ? 2 : 0);
  gpgol_release (msg);
  gpgoladdin_invalidate_ui();
  TRETURN;
}

bool
Mail::decryptedSuccessfully () const
{
  return m_decrypt_result.isNull() || !m_decrypt_result.error();
}

void
Mail::installFolderEventHandler_o()
{
  TSTART;
  TRACEPOINT;
  LPDISPATCH folder = get_oom_object (m_mailitem, "Parent");

  if (!folder)
    {
      TRACEPOINT;
      TRETURN;
    }

  char *objName = get_object_name (folder);
  if (!objName || strcmp (objName, "MAPIFolder"))
    {
      log_debug ("%s:%s: Mail %p parent is not a mapi folder.",
                 SRCNAME, __func__, m_mailitem);
      xfree (objName);
      gpgol_release (folder);
      TRETURN;
    }
  xfree (objName);

  char *path = get_oom_string (folder, "FullFolderPath");
  if (!path)
    {
      TRACEPOINT;
      path = get_oom_string (folder, "FolderPath");
    }
  if (!path)
    {
      log_error ("%s:%s: Mail %p parent has no folder path.",
                 SRCNAME, __func__, m_mailitem);
      gpgol_release (folder);
      TRETURN;
    }

  std::string strPath (path);
  xfree (path);

  if (s_folder_events_map.find (strPath) == s_folder_events_map.end())
    {
      log_debug ("%s:%s: Install folder events watcher for %s.",
                 SRCNAME, __func__, anonstr (strPath.c_str()));
      const auto sink = install_FolderEvents_sink (folder);
      s_folder_events_map.insert (std::make_pair (strPath, sink));
    }

  /* Folder already registered */
  gpgol_release (folder);
  TRETURN;
}

void
Mail::refCurrentItem()
{
  TSTART;
  if (m_currentItemRef)
    {
      log_debug ("%s:%s: Current item multi ref. Bug?",
                 SRCNAME, __func__);
      TRETURN;
    }
  /* This prevents a crash in Outlook 2013 when sending a mail as it
   * would unload too early.
   *
   * As it didn't crash when the mail was opened in Outlook Spy this
   * mimics that the mail is inspected somewhere else. */
  m_currentItemRef = get_oom_object (m_mailitem, "GetInspector.CurrentItem");
  TRETURN;
}

void
Mail::releaseCurrentItem()
{
  TSTART;
  if (!m_currentItemRef)
    {
      TRETURN;
    }
  log_oom ("%s:%s: releasing CurrentItem ref %p",
                 SRCNAME, __func__, m_currentItemRef);
  LPDISPATCH tmp = m_currentItemRef;
  m_currentItemRef = nullptr;
  /* This can cause our destruction */
  gpgol_release (tmp);
  TRETURN;
}

void
Mail::decryptPermanently_o()
{
  TSTART;
  if (!m_needs_wipe)
    {
      log_debug ("%s:%s: Mail does not yet need wipe. Called to early?",
                 SRCNAME, __func__);
      TRETURN;
    }

  if (!m_decrypt_result.isNull() && m_decrypt_result.error())
    {
      log_debug ("%s:%s: Decrypt result had error. Can't decrypt permanently.",
                 SRCNAME, __func__);
      TRETURN;
    }

  /* Remove the existing categories */
  removeCategories_o ();

  /* Drop our state variables */
  m_decrypt_result = GpgME::DecryptionResult();
  m_verify_result = GpgME::VerificationResult();
  m_needs_wipe = false;
  m_processed = false;
  m_is_smime = false;
  m_type = MSGTYPE_UNKNOWN;

  /* Remove our own attachments */
  removeOurAttachments_o ();

  updateSigstate();

  auto msg = MAKE_SHARED (get_oom_base_message (m_mailitem));
  if (!msg)
    {
      STRANGEPOINT;
      TRETURN;
    }
  mapi_delete_gpgol_tags ((LPMESSAGE)msg.get());

  mapi_set_mesage_class ((LPMESSAGE)msg.get(), "IPM.Note");

  if (invoke_oom_method (m_mailitem, "Save", NULL))
    {
      log_error ("Failed to save decrypted mail: %p ", m_mailitem);
    }

  log_debug ("%s:%s: Delayed invalidate to update sigstate after perm dec.",
             SRCNAME, __func__);
  CloseHandle(CreateThread (NULL, 0, delayed_invalidate_ui, (LPVOID) 300, 0,
                            NULL));
  TRETURN;
}

void
Mail::prepareCrypto_o ()
{
  TSTART;

  // Check inline response state to fill out asynccryptdisabled.
  checkSyncCrypto_o ();

  if (!isAsyncCryptDisabled())
    {
      /* Obtain a reference of the current item. This prevents
       * an early unload which would crash Outlook 2013
       *
       * As it didn't crash when the mail was opened in Outlook Spy this
       * mimics that the mail is inspected somewhere else. */
      refCurrentItem ();
    }

  // First contact with a mail to encrypt update
  // state and oom data.
  updateOOMData_o (true);

  setCryptState (Mail::NeedsFirstAfterWrite);

  TRETURN;
}

/* Printing happens in two steps. First a Mail is loaded after the
BeforePrint event, then it is loaded a second time when the actual print
happens. We have to catch both.
This functions looks over all mails and checks if one is currently
printing. If so we compare our EntryID's and if they match. Bingo,
we are printing, too.*/
bool
Mail::checkIfMailIsChildOfPrintMail_o ()
{
  gpgol_lock (&mail_map_lock);
  for (auto it = s_mail_map.begin(); it != s_mail_map.end(); ++it)
    {
      auto mail = it->second;
      if (mail->isPrint ())
        {
          /* This happens so rarely that we only fetch our
             entry id if we are in here. */
          char *entryID = get_oom_string (mail->item (), "EntryID");
          if (!entryID)
            {
              log_error ("%s:%s: Printing mail %p has no EntryID",
                         SRCNAME, __func__, mail);
              continue;
            }

          char *ourID = get_oom_string (m_mailitem, "EntryID");
          if (!ourID)
            {
              log_error ("%s:%s: Mail %p has no EntryID",
                         SRCNAME, __func__, this);
              xfree (entryID);
              continue;
            }
          int cmp = strcmp (ourID, entryID);
          xfree (ourID);
          xfree (entryID);

          if (cmp)
            {
              log_debug ("%s:%s: The current print is not us.",
                         SRCNAME, __func__);
              continue;
            }
          gpgrt_lock_unlock (&mail_map_lock);
          log_debug ("%s:%s: Mail %p is the actual print of %p.",
                     SRCNAME, __func__, this, mail);
          return true;
        }
    }
  gpgrt_lock_unlock (&mail_map_lock);
  return false;
}

void
Mail::setSigningKeys (const std::vector<GpgME::Key> &keys)
{
  m_resolved_signing_keys = keys;
}

std::vector<GpgME::Key>
Mail::getSigningKeys () const
{
  return m_resolved_signing_keys;
}

LPDISPATCH
Mail::copy ()
{
  TSTART;
  VARIANT result;
  VariantInit (&result);

  if (g_mail_copy_triggerer)
    {
      log_err ("BUG: Copy already in progress.");
      TRETURN nullptr;
    }
  g_mail_copy_triggerer = this;
  int err = invoke_oom_method (m_mailitem, "Copy", &result);

  if (err)
    {
      log_error ("%s:%s Failed to copy mail.",
                 SRCNAME, __func__);
      TRETURN nullptr;
    }

  if (result.vt != VT_DISPATCH || !result.pdispVal)
    {
      log_error ("%s:%s Unexpected result type %x.",
                 SRCNAME, __func__, result.vt);
      TRETURN nullptr;
    }
  return result.pdispVal;
}

static std::vector<GpgME::Key>
getKeysForProtocol (const std::vector<GpgME::Key> keys, GpgME::Protocol proto)
{
  std::vector<GpgME::Key> ret;

  std::copy_if (keys.begin(), keys.end(),
                std::back_inserter (ret),
                [proto] (const auto &k)
                {
                  return k.protocol () == proto;
                });

  return ret;
}


/* A callback of a copied mail to update our Mail objects
   after the split */
void
Mail::splitCopyMailCallback (Mail *copied_mail)
{
  TSTART;
  log_dbg ("Copy mail callback reached with mail %p", copied_mail);
  g_mail_copy_triggerer = nullptr;
  copied_mail->setSplitCopy (true);

  std::vector<Recipient> newRecipientsForCopy;
  std::vector<Recipient> newRecipientsForUs;
  GpgME::Protocol copyProtocol = GpgME::UnknownProtocol;

  /* Now split out recipients */
  bool bccFound = false;
  bool normalRecipientsFound = false;
  log_debug ("Dump before split.");
  Recipient::dump (m_cached_recipients);
  bool failureToRemoveBCC = false;
  for (const auto &recp: m_cached_recipients)
    {
      /* We want to always include the originator */
      if (recp.type () == Recipient::olOriginator)
        {
          newRecipientsForCopy.push_back (recp);
          newRecipientsForUs.push_back (recp);
        }
      else if (recp.type () == Recipient::olBCC && !bccFound)
        {
          const auto bccKeys = recp.keys ();
          if (bccKeys.empty())
            {
              log_err ("Empty keylist for recipient!");
              continue;
            }

          const auto pgpKeys = getKeysForProtocol (bccKeys, GpgME::OpenPGP);
          const auto cmsKeys = getKeysForProtocol (bccKeys, GpgME::CMS);

          if (!pgpKeys.empty () && !cmsKeys.empty())
            {
              log_dbg ("Detected BCC Recipient with both PGP and CMS keys."
                       " splitting recipient.");
              /* Copy it and keep it as BCC recipient with only PGP keys */
              auto splitRecipient = recp;
              splitRecipient.setKeys (pgpKeys);
              newRecipientsForUs.push_back (splitRecipient);

              /* For the copy use the CMS keys */
              auto splitRecipient2 = recp;
              splitRecipient2.setKeys (cmsKeys);
              newRecipientsForCopy.push_back (splitRecipient2);
              copyProtocol = GpgME::CMS;
            }
          else
            {
              newRecipientsForCopy.push_back (recp);
              /* Our keylist is of a single protocol so we can take that. */
              copyProtocol = bccKeys[0].protocol ();
              log_dbg ("Recipient '%s' is BCC. Gets its own mail with "
                       "protocol: %s", anonstr (recp.mbox().c_str ()),
                       to_cstr (copyProtocol));
            }
          int err = remove_oom_recipient (m_mailitem, recp.mbox());
          if (err)
            {
              log_dbg ("Failure to remove recipient '%s' possibly a group."
                       "overwriting with our recipients.",
                       anonstr (recp.mbox ().c_str ()));
              failureToRemoveBCC = true;
            }
          bccFound = true;
        }
      else
        {
          newRecipientsForUs.push_back (recp);
        }
    }
  if (failureToRemoveBCC)
    {
      set_oom_recipients (m_mailitem, newRecipientsForUs);
    }
  setRecipients(newRecipientsForUs);
  copied_mail->setRecipients(newRecipientsForCopy);
  set_oom_recipients (copied_mail->item(), newRecipientsForCopy);
  if (copyProtocol != GpgME::UnknownProtocol)
    {
      copied_mail->setSigningKeys (getKeysForProtocol (m_resolved_signing_keys,
                                                       copyProtocol));
    }
  else
    {
      copied_mail->setSigningKeys (m_resolved_signing_keys);
    }

  if (!bccFound)
    {
      /* TODO: Split more for S/MIME OpenPGP Mix. */
    }

  if (bccFound && normalRecipientsFound)
    {
      /* Recurse */
      log_dbg ("Both normal and hidden recipients found. "
               "Creating another copy.");
      copy ();
      TRETURN;
    }
}

void
Mail::splitAndSend_o ()
{
  TSTART;
  log_dbg ("Split and send for: %p", this);
  LPDISPATCH copied_mailitem = copy ();

  if (!copied_mailitem)
    {
      log_err ("Failed to copy mail. Aboring.");
      TRETURN;
    }
  /* This triggers the copyMailCallback in the write event */
  invoke_oom_method (copied_mailitem, "Send", nullptr);
  invoke_oom_method (m_mailitem, "Send", nullptr);
  TRETURN;
}

void
Mail::resetCrypter ()
{
  m_crypter = nullptr;
}

void
Mail::resetRecipients ()
{
  /* Clear recipient data */
  m_cached_recipients.clear ();
  m_resolved_signing_keys.clear ();
  m_recipients_set = false;
}

void
Mail::setSplitCopy (bool val)
{
  m_is_split_copy = val;
}

bool
Mail::isSplitCopy () const
{
  return m_is_split_copy;
}

int
Mail::buildProtectedHeaders_o ()
{
  TSTART;
  std::stringstream ss;
  std::vector <std::string> toRecps;
  std::vector <std::string> ccRecps;
  std::vector <std::string> bccRecps;
  ss << "protected-headers=\"v1\"\r\n";
  for (const auto &recp: m_cached_recipients)
    {
      if (recp.type() == Recipient::olCC)
        {
          ccRecps.push_back (recp.encodedDisplayName ());
        }
      else if (recp.type() == Recipient::olTo)
        {
          toRecps.push_back (recp.encodedDisplayName ());
        }
      else if (recp.type() == Recipient::olOriginator)
        {
          ss << "From: " << recp.encodedDisplayName () << "\r\n";
        }
    }
  std::string buf;
  if (toRecps.size ())
    {
      join(toRecps, ";", buf);
      ss << "To: " << buf << "\r\n";
    }
  if (ccRecps.size ())
    {
      join(toRecps, ";", buf);
      ss << "CC: " << buf << "\r\n";
    }

  if (opt.encryptSubject)
    {
      char *subject = get_oom_string (m_mailitem, "Subject");
      if (subject)
        {
          char *encodedSubject = utf8_to_rfc2047b (subject);
          if (encodedSubject)
            {
              ss << "Subject: " << encodedSubject << "\r\n";
            }
          xfree (encodedSubject);
        }
      if (!put_oom_string (m_mailitem, "Subject", "..."))
        {
          STRANGEPOINT;
          TRETURN -1;
        }
      xfree (subject);
    }

  m_protected_headers = ss.str ();
  TRETURN 0;
}

void
Mail::setProtectedHeaders (const std::string &hdr)
{
  m_protected_headers = hdr;
}

std::string
Mail::protectedHeaders () const
{
  return m_protected_headers;
}

header_info_s
Mail::headerInfo () const
{
  return m_header_info;
}

const std::string&
Mail::getOriginalBody () const
{
  return m_orig_body;
}
