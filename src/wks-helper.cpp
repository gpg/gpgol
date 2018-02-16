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

#include <map>

#include <unistd.h>

#include <gpg-error.h>
#include <gpgme++/key.h>
#include <gpgme++/data.h>
#include <gpgme++/context.h>

#define CHECK_MIN_INTERVAL (60 * 60 * 24 * 7)

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
  const auto state = success ? WKSHelper::Supported : WKSHelper::NotSupported;

  gpgrt_lock_lock (&wks_lock);

  auto it = s_states.find(mbox);

  // TODO figure out if it was published.
  if (success)
    {
      log_debug ("%s:%s: WKS client: '%s' is supported",
                 SRCNAME, __func__, mbox.c_str ());
    }
  if (it != s_states.end())
    {
      it->second = state;
    }
  else
    {
      s_states.insert (std::make_pair (mbox, state));
    }

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
