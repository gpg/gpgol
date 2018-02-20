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
#include "overlay.h"
#include "mail.h"

#include <map>
#include <sstream>

#include <unistd.h>

#include <gpg-error.h>
#include <gpgme++/key.h>
#include <gpgme++/data.h>
#include <gpgme++/context.h>

#define CHECK_MIN_INTERVAL (60 * 60 * 24 * 7)

#undef _
#define _(a) utf8_gettext (a)

static std::map <std::string, WKSHelper::WKSState> s_states;
static std::map <std::string, time_t> s_last_checked;

static WKSHelper* singleton = NULL;

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
  const auto state = success ? WKSHelper::NeedsPublish : WKSHelper::NotSupported;
  if (success)
    {
      log_debug ("%s:%s: WKS client: '%s' is supported",
                 SRCNAME, __func__, mbox.c_str ());
    }

  WKSHelper::instance()->update_state (mbox, state);

  gpgrt_lock_lock (&wks_lock);
  auto tit = s_last_checked.find(mbox);
  auto now = time (0);
  if (tit != s_last_checked.end())
    {
      tit->second = now;
    }
  else
    {
      s_last_checked.insert (std::make_pair (mbox, now));
    }
  gpgrt_lock_unlock (&wks_lock);

  WKSHelper::instance()->save ();
  return 0;
}


void
WKSHelper::start_check (const std::string &mbox, bool forced) const
{
  auto lastTime = get_check_time (mbox);
  auto now = time (0);
  if (!forced && lastTime && difftime (lastTime, now) < CHECK_MIN_INTERVAL)
    {
      /* Data is new enough */
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
  CloseHandle (CreateThread (NULL, 0, do_check, strdup (mbox.c_str ()), 0,
                             NULL));
  return;
}

void
WKSHelper::load () const
{
  // TODO
}

void
WKSHelper::save () const
{
  // TODO
}

static DWORD WINAPI
do_notify (LPVOID arg)
{
  /** Wait till a message was sent */
  //Sleep (5000);
  do_in_ui_thread (WKS_NOTIFY, arg);

  return 0;
}

void
WKSHelper::allow_notify () const
{
  gpgrt_lock_lock (&wks_lock);
  for (auto &pair: s_states)
    {
      if (pair.second == NeedsPublish)
        {
          CloseHandle (CreateThread (NULL, 0, do_notify,
                                     strdup (pair.first.c_str ()), 0,
                                     NULL));
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
      wchar_t * w_title = utf8_to_wchar (_("GpgOL: Key directory available!"));
      wchar_t * w_desc = utf8_to_wchar (_("Your mail provider supports a key directory.\n\n"
                                          "Register your key in that directory to make\n"
                                          "it easier for others to send you encrypted mail.\n\n\n"
                                          "Register Key?"));
      if (MessageBoxW (get_active_hwnd (),
                       w_desc, w_title, MB_ICONINFORMATION | MB_YESNO) == IDYES)
        {
          start_publish (mbox);
        }
      else
        {
           update_state (mbox, PublishDenied);
        }

      xfree (w_desc);
      xfree (w_title);
      return;
    }
  else
    {
      log_debug ("%s:%s: Unhandled notify state: %i for '%s'",
                 SRCNAME, __func__, state, cBox);
      return;
    }
}

void
WKSHelper::start_publish (const std::string &mbox) const
{
//  Overlay (get_active_hwnd (),
//           std::string (_("Creating registration request...")));

  log_debug ("%s:%s: Start publish for '%s'",
             SRCNAME, __func__, mbox.c_str ());

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
      MessageBox (get_active_hwnd (),
                  "WKS client failed to create publishing request.",
                  _("GpgOL"),
                  MB_ICONINFORMATION|MB_OK);
      return;
    }

  log_debug ("%s:%s: WKS client: returned '%s'",
             SRCNAME, __func__, data.c_str ());

  send_mail (data);

  return;
}


void
WKSHelper::update_state (const std::string &mbox, WKSState state) const
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
}

void
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
      return;
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
      return;
    }

  if (put_oom_string (mail, "Subject", subject.c_str ()))
    {
      TRACEPOINT;
      gpgol_release (mail);
      return;
    }

  if (put_oom_string (mail, "To", to.c_str ()))
    {
      TRACEPOINT;
      gpgol_release (mail);
      return;
    }

  LPDISPATCH account = get_account_for_mail (from.c_str ());
  if (account)
    {
      log_debug ("%s:%s: Found account to change for '%s'.",
                 SRCNAME, __func__, from.c_str ());
      put_oom_disp (mail, "SendUsingAccount", account);
    }
  gpgol_release (account);

  /* Now we have a problem. The created LPDISPATCH pointer has
     a different value then the one with which we saw the ItemLoad
     event. But we want to get the mail object. So,.. surpise
     a Hack! :-) */
  auto last_mail = Mail::get_last_mail ();

  last_mail->set_override_mime_data (mimeData);
  last_mail->set_crypt_state (Mail::NeedsSecondAfterWrite);

  invoke_oom_method (mail, "Save", NULL);
  invoke_oom_method (mail, "Send", NULL);

  log_debug ("%s:%s: Publish successful",
             SRCNAME, __func__);
}
