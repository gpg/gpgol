/* addin-options.cpp - Options for the Ol >= 2010 Addin
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
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

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <windows.h>
#include "dialogs.h"
#include "common.h"
#include "cpphelp.h"
#include "oomhelp.h"

#include <string>

#include <gpgme++/context.h>
#include <gpgme++/data.h>


__attribute__((__unused__)) static char const *
i18n_noops[] = {
    N_("GnuPG System"),
    N_("Enable the S/MIME support"),
    N_("Configure GpgOL"),
    N_("Automation"),
    N_("General"),

    N_("Automatically secure &messages"),
    N_("Configure GnuPG"),
    N_("Debug..."),
    N_("Version "),
    N_("&Resolve recipient keys automatically"),
    N_("&Encrypt new messages by default"),
    N_("&Sign new messages by default"),
    N_("&Send OpenPGP mails without "
       "attachments as PGP/Inline"),
    N_("S&elect crypto settings automatically "
       "for reply and forward"),
    N_("&Prefer S/MIME when automatically resolving recipients"),

    /* Tooltips */
    N_("Enable or disable any automated key handling."),
    N_("Automate trust based on communication history."),
    N_("This changes the trust model to \"tofu+pgp\" which tracks the history of key usage. Automated trust can <b>never</b> exceed level 2."),
    N_("experimental"),
    N_("Automatically toggles secure if keys with at least level 1 trust were found for all recipients."),
    N_("Toggles the encrypt option for all new mails."),
    N_("Toggles the sign option for all new mails."),
    N_("Toggles sign, encrypt options if the original mail was signed or encrypted."),
    N_("Instead of using the PGP/MIME format, "
       "which properly handles attachments and encoding, "
       "the deprecated PGP/Inline is used.\n"
       "This can be required for compatibility but should generally not "
       "be used."),
};

static bool dlg_open;

static DWORD WINAPI
open_gpgolgui (LPVOID arg)
{
  HWND wnd = (HWND) arg;

  std::vector<std::string> args;

  // Collect the arguments
  char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      return -1;
    }
  const auto gpgolgui = std::string (gpg4win_dir) + "\\bin\\gpgolgui.exe";
  args.push_back (gpgolgui);

  args.push_back (std::string ("--hwnd"));
  args.push_back (std::to_string ((int) (intptr_t) wnd));

  args.push_back (std::string("--gpgol-version"));
  args.push_back (std::string(VERSION));

  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      // can't happen
      TRACEPOINT;
      return -1;
    }

  GpgME::Data mystdin (GpgME::Data::null), mystdout, mystderr;
  dlg_open = true;

  char **cargs = vector_to_cArray (args);
  log_debug ("%s:%s: args:", SRCNAME, __func__);
  for (size_t i = 0; cargs && cargs[i]; i++)
    {
      log_debug (SIZE_T_FORMAT ": '%s'", i, cargs[i]);
    }

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char**> (cargs),
                                 mystdin, mystdout, mystderr,
                                 (GpgME::Context::SpawnFlags) (
                                  GpgME::Context::SpawnAllowSetFg |
                                  GpgME::Context::SpawnShowWindow));
  release_cArray (cargs);
  if (err)
    {
      log_error ("%s:%s: Err code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
    }
  dlg_open = false;

  log_debug ("%s:%s:finished stdout:\n'%s'",
             SRCNAME, __func__, mystdout.toString ().c_str ());
  log_debug ("%s:%s:stderr:\n'%s'",
             SRCNAME, __func__, mystderr.toString ().c_str ());
  read_options ();
  return 0;
}

void
options_dialog_box (HWND parent)
{
  if (!parent)
    parent = get_active_hwnd ();

  if (dlg_open)
    {
      log_debug ("%s:%s: Gpgolgui open. Not launching new dialog.",
                 SRCNAME, __func__);
      HWND optWindow = FindWindow (nullptr, _("Configure GpgOL"));
      if (!optWindow) {
        log_debug ("%s:%s: Gpgolgui open but could not find window.",
                 SRCNAME, __func__);
        return;
      }
      SetForegroundWindow(optWindow);

      return;
    }

  log_debug ("%s:%s: Launching gpgolgui.",
             SRCNAME, __func__);

  CloseHandle (CreateThread (NULL, 0, open_gpgolgui, (LPVOID) parent, 0,
                             NULL));
}
