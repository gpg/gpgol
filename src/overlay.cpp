/*
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
#include "config.h"

#include "overlay.h"

#include "common.h"
#include "cpphelp.h"

#include <vector>

#include <gpgme++/context.h>

Overlay::Overlay (HWND wid, const std::string &text): m_wid (wid)
{
  // Disable the window early to avoid it beeing closed.
  EnableWindow (m_wid, FALSE);
  std::vector<std::string> args;

  // Collect the arguments
  char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      EnableWindow (m_wid, TRUE);
      return;
    }
  const auto overlayer = std::string (gpg4win_dir) + "\\bin\\overlayer.exe";
  xfree (gpg4win_dir);
  args.push_back (overlayer);

  args.push_back (std::string ("--hwnd"));
  args.push_back (std::to_string ((int) wid));

  args.push_back (std::string ("--overlayText"));
  args.push_back (text);
  char **cargs = vector_to_cArray (args);

  m_overlayCtx = GpgME::Context::createForEngine (GpgME::SpawnEngine);

  if (!m_overlayCtx)
    {
      // can't happen
      release_cArray (cargs);
      TRACEPOINT;
      EnableWindow (m_wid, TRUE);
      return;
    }

  GpgME::Data mystderr(GpgME::Data::null);
  GpgME::Data mystdout(GpgME::Data::null);

  GpgME::Error err = m_overlayCtx->spawnAsync (cargs[0], const_cast <const char**> (cargs),
                                               m_overlayStdin, mystdout, mystderr,
                                               (GpgME::Context::SpawnFlags) (
                                                GpgME::Context::SpawnAllowSetFg |
                                                GpgME::Context::SpawnShowWindow));
  release_cArray (cargs);

  log_debug ("%s:%s: Created overlay window over %p",
             SRCNAME, __func__, wid);
}

Overlay::~Overlay()
{
  log_debug ("%s:%s: Stopping overlay.",
             SRCNAME, __func__);
  m_overlayStdin.write ("quit\n", 5);
  m_overlayStdin.write (nullptr, 0);
  m_overlayCtx->wait ();
  EnableWindow (m_wid, TRUE);
}
