/* wks-helper.cpp - Web Key Services for GpgOL
 * Copyright (C) 2018 Intevation GmbH
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

#include "wks-helper.h"

#include "common.h"
#include "cpphelp.h"
#include "oomhelp.h"
#include "windowmessages.h"
#include "mail.h"
#include "mapihelp.h"

#include <map>
#include <sstream>

#include <unistd.h>
#include <stdlib.h>

#include <gpg-error.h>
#include <gpgme++/key.h>
#include <gpgme++/data.h>
#include <gpgme++/context.h>

#define CHECK_MIN_INTERVAL (60 * 60 * 24 * 7)
#define WKS_REG_KEY "webkey"

#define DEBUG_WKS 1

#undef _
#define _(a) utf8_gettext (a)

static std::map <std::string, WKSHelper::WKSState> s_states;
static std::map <std::string, time_t> s_last_checked;
static std::map <std::string, std::pair <GpgME::Data *, Mail *> > s_confirmation_cache;

static WKSHelper* singleton = nullptr;

GPGRT_LOCK_DEFINE (wks_lock);

WKSHelper::WKSHelper()
{
  load ();
}

WKSHelper::~WKSHelper ()
{
  // Ensure that we are not destroyed while
  // worker is running.
  gpgrt_lock_lock (&wks_lock);
  gpgrt_lock_unlock (&wks_lock);
}

const WKSHelper*
WKSHelper::instance ()
{
  if (!singleton)
    {
      singleton = new WKSHelper ();
    }
  return singleton;
}

WKSHelper::WKSState
WKSHelper::get_state (const std::string &mbox) const
{
  gpgrt_lock_lock (&wks_lock);
  const auto it = s_states.find(mbox);
  const auto dataEnd = s_states.end();
  gpgrt_lock_unlock (&wks_lock);
  if (it == dataEnd)
    {
      return NotChecked;
    }
  return it->second;
}

time_t
WKSHelper::get_check_time (const std::string &mbox) const
{
  gpgrt_lock_lock (&wks_lock);
  const auto it = s_last_checked.find(mbox);
  const auto dataEnd = s_last_checked.end();
  gpgrt_lock_unlock (&wks_lock);
  if (it == dataEnd)
    {
      return 0;
    }
  return it->second;
}

std::pair <GpgME::Data *, Mail *>
WKSHelper::get_cached_confirmation (const std::string &mbox) const
{
  gpgrt_lock_lock (&wks_lock);
  const auto it = s_confirmation_cache.find(mbox);
  const auto dataEnd = s_confirmation_cache.end();

  if (it == dataEnd)
    {
      gpgrt_lock_unlock (&wks_lock);
      return std::make_pair (nullptr, nullptr);
    }
  auto ret = it->second;
  s_confirmation_cache.erase (it);
  gpgrt_lock_unlock (&wks_lock);
  return ret;
}

static std::string
get_wks_client_path ()
{
  char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      return std::string ();
    }
  const auto ret = std::string (gpg4win_dir) +
                  "\\..\\GnuPG\\bin\\gpg-wks-client.exe";
  xfree (gpg4win_dir);

  if (!access (ret.c_str (), F_OK))
    {
      return ret;
    }
  log_debug ("%s:%s: Failed to find wks-client in '%s'",
             SRCNAME, __func__, ret.c_str ());
  return std::string ();
}

static bool
check_published (const std::string &mbox)
{
  const auto wksPath = get_wks_client_path ();

  if (wksPath.empty())
    {
      return 0;
    }

  std::vector<std::string> args;

  args.push_back (wksPath);
  args.push_back (std::string ("--status-fd"));
  args.push_back (std::string ("1"));
  args.push_back (std::string ("--check"));
  args.push_back (mbox);

  // Spawn the process
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);

  if (!ctx)
    {
      TRACEPOINT;
      return false;
    }

  GpgME::Data mystdin, mystdout, mystderr;

  char **cargs = vector_to_cArray (args);

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char **> (cargs),
                                 mystdin, mystdout, mystderr,
                                 GpgME::Context::SpawnNone);
  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: WKS client spawn code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
      return false;
    }
  auto data = mystdout.toString ();
  rtrim (data);

  return data == "[GNUPG:] SUCCESS";
}

static DWORD WINAPI
do_check (LPVOID arg)
{
  const auto wksPath = get_wks_client_path ();

  if (wksPath.empty())
    {
      return 0;
    }

  std::vector<std::string> args;
  const auto mbox = std::string ((char *) arg);
  xfree (arg);

  args.push_back (wksPath);
  args.push_back (std::string ("--status-fd"));
  args.push_back (std::string ("1"));
  args.push_back (std::string ("--supported"));
  args.push_back (mbox);

  // Spawn the process
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);

  if (!ctx)
    {
      TRACEPOINT;
      return 0;
    }

  GpgME::Data mystdin, mystdout, mystderr;

  char **cargs = vector_to_cArray (args);

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char **> (cargs),
                                 mystdin, mystdout, mystderr,
                                 GpgME::Context::SpawnNone);
  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: WKS client spawn code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
      return 0;
    }

  auto data = mystdout.toString ();
  rtrim (data);

  bool success = data == "[GNUPG:] SUCCESS";
  // TODO Figure out NeedsPublish state.
  auto state = success ? WKSHelper::NeedsPublish : WKSHelper::NotSupported;
  bool isPublished = false;

  if (success)
    {
      log_debug ("%s:%s: WKS client: '%s' is supported",
                 SRCNAME, __func__, mbox.c_str ());
      isPublished = check_published (mbox);
    }

  if (isPublished)
    {
      log_debug ("%s:%s: WKS client: '%s' is published",
                 SRCNAME, __func__, mbox.c_str ());
      state = WKSHelper::IsPublished;
    }

  WKSHelper::instance()->update_state (mbox, state, false);
  WKSHelper::instance()->update_last_checked (mbox, time (0));

  return 0;
}


void
WKSHelper::start_check (const std::string &mbox, bool forced) const
{
  const auto state = get_state (mbox);

  if (!forced && (state != NotChecked && state != NotSupported))
    {
      log_debug ("%s:%s: Check aborted because its neither "
                 "not supported nor not checked.",
                 SRCNAME, __func__);
      return;
    }

  auto lastTime = get_check_time (mbox);
  auto now = time (0);

  if (!forced && (state == NotSupported && lastTime &&
                  difftime (now, lastTime) < CHECK_MIN_INTERVAL))
    {
      /* Data is new enough */
      log_debug ("%s:%s: Check aborted because last checked is too recent.",
                 SRCNAME, __func__);
      return;
    }

  if (mbox.empty())
    {
      log_debug ("%s:%s: start check called without mbox",
                 SRCNAME, __func__);
    }

  log_debug ("%s:%s: WKSHelper starting check",
             SRCNAME, __func__);
  /* Start the actual work that can be done in a background thread. */
  CloseHandle (CreateThread (nullptr, 0, do_check, xstrdup (mbox.c_str ()), 0,
                             nullptr));
  return;
}

void
WKSHelper::load () const
{
  /* Map of mbox'es to states. states are <state>;<last_checked> */
  const auto map = get_registry_subkeys (WKS_REG_KEY);

  for (const auto &pair: map)
    {
      const auto mbox = pair.first;
      const auto states = gpgol_split (pair.second, ';');

      if (states.size() != 2)
        {
          log_error ("%s:%s: Invalid state '%s' for '%s'",
                     SRCNAME, __func__, mbox.c_str (), pair.second.c_str ());
          continue;
        }

      WKSState state = (WKSState) strtol (states[0].c_str (), nullptr, 10);
      if (state == PublishInProgress)
        {
          /* Probably an error during the last publish. Let's start again. */
          update_state (mbox, NotChecked, false);
          continue;
        }

      time_t update_time = (time_t) strtol (states[1].c_str (), nullptr, 10);
      update_state (mbox, state, false);
      update_last_checked (mbox, update_time, false);
    }
}

void
WKSHelper::save () const
{
  gpgrt_lock_lock (&wks_lock);
  for (const auto &pair: s_states)
    {
      auto state = std::to_string (pair.second) + ';';

      const auto it = s_last_checked.find (pair.first);
      if (it != s_last_checked.end ())
        {
          state += std::to_string (it->second);
        }
      else
        {
          state += '0';
        }
      if (store_extension_subkey_value (WKS_REG_KEY, pair.first.c_str (),
                                        state.c_str ()))
        {
          log_error ("%s:%s: Failed to store state.",
                     SRCNAME, __func__);
        }
    }
  gpgrt_lock_unlock (&wks_lock);
}

static DWORD WINAPI
do_notify (LPVOID arg)
{
  /** Wait till a message was sent */
  std::pair<char *, int> *args = (std::pair<char *, int> *) arg;

  Sleep (args->second);
  do_in_ui_thread (WKS_NOTIFY, args->first);
  delete args;

  return 0;
}

void
WKSHelper::allow_notify (int sleepTimeMS) const
{
  gpgrt_lock_lock (&wks_lock);
  for (auto &pair: s_states)
    {
      if (pair.second == ConfirmationSeen ||
          pair.second == NeedsPublish)
        {
          auto *args = new std::pair<char *, int> (xstrdup (pair.first.c_str()),
                                                   sleepTimeMS);
          CloseHandle (CreateThread (nullptr, 0, do_notify,
                                     args, 0,
                                     nullptr));
          break;
        }
    }
  gpgrt_lock_unlock (&wks_lock);
}

void
WKSHelper::notify (const char *cBox) const
{
  std::string mbox = cBox;

  const auto state = get_state (mbox);

  if (state == NeedsPublish)
    {
      char *buf;
      gpgrt_asprintf (&buf, _("A Pubkey directory is available for the address:\n\n"
                              "\t%s\n\n"
                              "Register your Pubkey in that directory to make\n"
                              "it easy for others to send you encrypted mail.\n\n"
                              "It's secure and free!\n\n"
                              "Register automatically?"), mbox.c_str ());
      memdbg_alloc (buf);
      if (gpgol_message_box (get_active_hwnd (),
                             buf,
                             _("GpgOL: Pubkey directory available!"), MB_YESNO) == IDYES)
        {
          start_publish (mbox);
        }
      else
        {
          update_state (mbox, PublishDenied);
        }
      xfree (buf);
      return;
    }
  if (state == ConfirmationSeen)
    {
      handle_confirmation_notify (mbox);
      return;
    }

  log_debug ("%s:%s: Unhandled notify state: %i for '%s'",
             SRCNAME, __func__, state, cBox);
  return;
}

void
WKSHelper::start_publish (const std::string &mbox) const
{
  log_debug ("%s:%s: Start publish for '%s'",
             SRCNAME, __func__, mbox.c_str ());

  update_state (mbox, PublishInProgress);
  const auto key = GpgME::Key::locate (mbox.c_str ());

  if (key.isNull ())
    {
      MessageBox (get_active_hwnd (),
                  "WKS publish failed to find key for mail address.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      return;
    }

  const auto wksPath = get_wks_client_path ();

  if (wksPath.empty())
    {
      TRACEPOINT;
      return;
    }

  std::vector<std::string> args;

  args.push_back (wksPath);
  args.push_back (std::string ("--create"));
  args.push_back (std::string (key.primaryFingerprint ()));
  args.push_back (mbox);

  // Spawn the process
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      TRACEPOINT;
      return;
    }

  GpgME::Data mystdin, mystdout, mystderr;

  char **cargs = vector_to_cArray (args);

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char **> (cargs),
                                 mystdin, mystdout, mystderr,
                                 GpgME::Context::SpawnNone);
  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: WKS client spawn code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
      return;
    }
  const auto data = mystdout.toString ();

  if (data.empty ())
    {
      gpgol_message_box (get_active_hwnd (),
                         mystderr.toString().c_str (),
                         _("GpgOL: Directory request failed"),
                         MB_OK);
      return;
    }

#ifdef DEBUG_WKS
  log_debug ("%s:%s: WKS client: returned '%s'",
             SRCNAME, __func__, data.c_str ());
#endif

  if (!send_mail (data))
    {
      gpgol_message_box (get_active_hwnd (),
                         _("You might receive a confirmation challenge from\n"
                           "your provider to finish the registration."),
                         _("GpgOL: Registration request sent!"), MB_OK);
    }

  update_state (mbox, RequestSent);
  return;
}

void
WKSHelper::update_state (const std::string &mbox, WKSState state,
                         bool store) const
{
  gpgrt_lock_lock (&wks_lock);
  auto it = s_states.find(mbox);

  if (it != s_states.end())
    {
      it->second = state;
    }
  else
    {
      s_states.insert (std::make_pair (mbox, state));
    }
  gpgrt_lock_unlock (&wks_lock);

  if (store)
    {
      save ();
    }
}

void
WKSHelper::update_last_checked (const std::string &mbox, time_t time,
                                bool store) const
{
  gpgrt_lock_lock (&wks_lock);
  auto it = s_last_checked.find(mbox);
  if (it != s_last_checked.end())
    {
      it->second = time;
    }
  else
    {
      s_last_checked.insert (std::make_pair (mbox, time));
    }
  gpgrt_lock_unlock (&wks_lock);

  if (store)
    {
      save ();
    }
}

int
WKSHelper::send_mail (const std::string &mimeData) const
{
  std::istringstream ss(mimeData);

  std::string from;
  std::string to;
  std::string subject;
  std::string withoutHeaders;

  std::getline (ss, from);
  std::getline (ss, to);
  std::getline (ss, subject);

  if (from.compare (0, 6, "From: ") || to.compare (0, 4, "To: "),
      subject.compare (0, 9, "Subject: "))
    {
      log_error ("%s:%s: Invalid mime data..",
                 SRCNAME, __func__);
      return -1;
    }

  std::getline (ss, withoutHeaders, '\0');

  from.erase (0, 6);
  to.erase (0, 4);
  subject.erase (0, 9);

  rtrim (from);
  rtrim (to);
  rtrim (subject);

  LPDISPATCH mail = create_mail ();

  if (!mail)
    {
      log_error ("%s:%s: Failed to create mail for request.",
                 SRCNAME, __func__);
      return -1;
    }

  /* Now we have a problem. The created LPDISPATCH pointer has
     a different value then the one with which we saw the ItemLoad
     event. But we want to get the mail object. So,.. surpise
     a Hack! :-) */
  auto last_mail = Mail::getLastMail ();

  if (!Mail::isValidPtr (last_mail))
    {
      log_error ("%s:%s: Invalid last mail %p.",
                 SRCNAME, __func__, last_mail);
      return -1;
    }
  /* Adding to / Subject etc already leads to changes and events so
     we set up the state before this. */
  last_mail->setOverrideMIMEData (mimeData);
  last_mail->setCryptState (Mail::NeedsSecondAfterWrite);

  if (put_oom_string (mail, "Subject", subject.c_str ()))
    {
      TRACEPOINT;
      gpgol_release (mail);
      return -1;
    }

  if (put_oom_string (mail, "To", to.c_str ()))
    {
      TRACEPOINT;
      gpgol_release (mail);
      return -1;
    }

  LPDISPATCH account = get_account_for_mail (from.c_str ());
  if (account)
    {
      log_debug ("%s:%s: Found account to change for '%s'.",
                 SRCNAME, __func__, from.c_str ());
      put_oom_disp (mail, "SendUsingAccount", account);
    }
  gpgol_release (account);

  if (invoke_oom_method (mail, "Save", nullptr))
    {
      // Should not happen.
      log_error ("%s:%s: Failed to save mail.",
                 SRCNAME, __func__);
      return -1;
    }
  if (invoke_oom_method (mail, "Send", nullptr))
    {
      log_error ("%s:%s: Failed to send mail.",
                 SRCNAME, __func__);
      return -1;
    }
  log_debug ("%s:%s: Done send mail.",
             SRCNAME, __func__);
  return 0;
}

static void
copy_stream_to_data (LPSTREAM stream, GpgME::Data *data)
{
  HRESULT hr;
  char buf[4096];
  ULONG bRead;
  while ((hr = stream->Read (buf, 4096, &bRead)) == S_OK ||
         hr == S_FALSE)
    {
      if (!bRead)
        {
          // EOF
          return;
        }
      data->write (buf, (size_t) bRead);
    }
}

void
WKSHelper::handle_confirmation_notify (const std::string &mbox) const
{
  auto pair = get_cached_confirmation (mbox);
  GpgME::Data *mimeData = pair.first;
  Mail *mail = pair.second;

  if (!mail)
    {
      log_debug ("%s:%s: Confirmation notify without cached mail.",
                 SRCNAME, __func__);
    }

  if (!mimeData)
    {
      log_error ("%s:%s: Confirmation notify without cached data.",
                 SRCNAME, __func__);
      return;
    }

  /* First ask the user if he wants to confirm */
  if (gpgol_message_box (get_active_hwnd (),
                         _("Confirm registration?"),
                         _("GpgOL: Pubkey directory confirmation"), MB_YESNO) != IDYES)
    {
      log_debug ("%s:%s: User aborted confirmation.",
                 SRCNAME, __func__);
      delete mimeData;

      /* Next time we read the confirmation we ask again. */
      update_state (mbox, RequestSent);
      return;
    }

  /* Do the confirmation */
  const auto wksPath = get_wks_client_path ();

  if (wksPath.empty())
    {
      TRACEPOINT;
      return;
    }

  std::vector<std::string> args;

  args.push_back (wksPath);
  args.push_back (std::string ("--receive"));

  // Spawn the process
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      TRACEPOINT;
      return;
    }
  GpgME::Data mystdout, mystderr;

  char **cargs = vector_to_cArray (args);

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char **> (cargs),
                                 *mimeData, mystdout, mystderr,
                                 GpgME::Context::SpawnNone);
  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: WKS client spawn code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
      return;
    }
  const auto data = mystdout.toString ();

  if (data.empty ())
    {
      gpgol_message_box (get_active_hwnd (),
                         mystderr.toString().c_str (),
                         _("GpgOL: Confirmation failed"),
                         MB_OK);
      return;
    }

#ifdef DEBUG_WKS
  log_debug ("%s:%s: WKS client: returned '%s'",
             SRCNAME, __func__, data.c_str ());
#endif
  if (!send_mail (data))
   {
     gpgol_message_box (get_active_hwnd (),
                        _("Your Pubkey can soon be retrieved from your domain."),
                        _("GpgOL: Request confirmed!"), MB_OK);
   }

  if (mail && Mail::isValidPtr (mail))
    {
      invoke_oom_method (mail->item(), "Delete", nullptr);
    }

  update_state (mbox, ConfirmationSent);
}

void
WKSHelper::handle_confirmation_read (Mail *mail, LPSTREAM stream) const
{
  /* We get the handle_confirmation in the Read event. To do sending
     etc. we have to move out of that event. For this we prepare
     the data for later usage. */

  if (!mail || !stream)
    {
      TRACEPOINT;
      return;
    }

  /* Get the recipient of the confirmation mail */
  const auto recipients = mail->getRecipients_o ();

  /* We assert that we have one recipient as the mail should have been
     sent by the wks-server. */
  if (recipients.size() != 1)
    {
      log_error ("%s:%s: invalid recipients",
                 SRCNAME, __func__);
      gpgol_release (stream);
      return;
    }

  std::string mbox = recipients[0];

  /* Prepare stdin for the wks-client process */

  /* First we need to write the headers */
  LPMESSAGE message = get_oom_base_message (mail->item());
  if (!message)
    {
      log_error ("%s:%s: Failed to obtain message.",
                 SRCNAME, __func__);
      gpgol_release (stream);
      return;
    }

  const auto headers = mapi_get_header (message);
  gpgol_release (message);

  GpgME::Data *mystdin = new GpgME::Data();

  mystdin->write (headers.c_str (), headers.size ());

  /* Then the MIME data */
  copy_stream_to_data (stream, mystdin);
  gpgol_release (stream);

  /* Then lets make sure its flushy */
  mystdin->write (nullptr, 0);

  /* And reset it to start */
  mystdin->seek (0, SEEK_SET);

  gpgrt_lock_lock (&wks_lock);
  s_confirmation_cache.insert (std::make_pair (mbox, std::make_pair (mystdin, mail)));
  gpgrt_lock_unlock (&wks_lock);

  update_state (mbox, ConfirmationSeen);

  /* Send the window message for notify. */
  allow_notify (5000);
}
