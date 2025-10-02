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
    N_("(Technical)"),

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
    N_("&Prefer S/MIME"),

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
       "This can be useful for compatibility but should generally not "
       "be used."),
    N_("Prefer S/MIME over OpenPGP if both are possible."),

    /* TRANSLATORS: Part of the config dialog. */
    N_("Configuration of GnuPG System options"),
    /* TRANSLATORS: Part of the config dialog. */
    N_("Configuration of debug options"),

    /* TRANSLATORS: Part of the config dialog. */
    N_("Search and import &X509 certificates in the configured directory services"),
    /* TRANSLATORS: Part of the config dialog. Tooltip */
    N_("Searches for X509 certificates automatically and imports them. This option searches in all configured services."),
    /* TRANSLATORS: Part of the config dialog. Warning about privacy leak. */
    N_("<b>Warning:</b> The configured services will receive information about whom you send Emails!"),
    /* TRANSLATORS: Part of the config dialog. */
    N_("Also automatically toggles secure if keys with level 0 trust were found."),
    /* TRANSLATORS: Part of the config dialog. */
    N_("Also &with untrusted keys"),
    /* TRANSLATORS: Included means here both attached keys and keys from the
     * headers */
    N_("&Import any keys included in mails"),
    /* TRANSLATORS: Included means here both attached keys and keys from the
     * headers */
    N_("Import OpenPGP keys from mail attachments or from mail headers."),
    /* TRANSLATORS: Part of the config dialog. */
    N_("Encrypt &drafts of secure mails to this key:"),
    /* TRANSLATORS: Part of the config dialog. */
    N_("Encrypt drafts and autosaved mails if the secure button is toggled."),

    /* Not options but strings for the key adder */
    /* TRANSLATORS: Part of address book key configuration dialog.
       The contacts name follows. */
    N_("Settings for:"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       An example for a public key follows. */
    N_("Paste a public key export here. It should look like:"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       An example for a public certificate follows. */
    _("Paste certificates here. They should look like:"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Failed to parse any public key."),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Error"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Secret key detected."),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("You can only configure public keys in Outlook."
       " Import secret keys with Kleopatra."),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("The key is unusable for Outlook."
       " Please check Kleopatra for more information."),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Invalid key detected."),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Created:"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("User Ids:"),
    /* TRANSLATORS: Part of address book key configuration dialog. %1 is
        a placeholder for the plual for key / keys. */
    N_("You are about to configure the following OpenPGP %1 for:"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       used in a sentence as plural form. */
    N_("keys"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       used in a sentence as singular form. */
    N_("key"),
    /* TRANSLATORS: Part of address book key configuration dialog. %1 is
        a placeholder for the plual for certificate / certificates. */
    N_("You are about to configure the following S/MIME %1 for:"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       used in a sentence as plural form. */
    N_("certificates"),
    /* TRANSLATORS: Part of address book key configuration dialog.
       used in a sentence as singular form. */
    N_("certificate"),

    N_("This may take several minutes..."),
    N_("Validating S/MIME certificates"),

    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Continue?"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Confirm"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Always secure mails"),
    /* TRANSLATORS: Part of address book key configuration dialog. */
    N_("Use these keys for this contact:"),

    /* TRANSLATORS: Part of address book key configuration dialog.
       Info box for the key configuration. */
    N_("You can use this to override the keys "
       "for this contact. The keys will be imported and used "
       "regardless of their trust level."),
    N_("For S/MIME the root certificate has to be trusted."),
    N_("Place multiple keys in here to encrypt to all of them."),
    /* TRANSLATORS: Part of debugging configuration. */
    N_("Enable Logging"),
    N_("Default"),
    /* TRANSLATORS: Part of debugging configuration.  The plus should
    mean in the combo box that it is added to the above. */
    N_("+Outlook API calls"),
    /* TRANSLATORS: Part of debugging configuration.  The plus should
    mean in the combo box that it is added to the above. */
    N_("+Memory analysis"),
    /* TRANSLATORS: Part of debugging configuration.  The plus should
    mean in the combo box that it is added to the above. */
    N_("+Call tracing"),
    /* TRANSLATORS: Part of debugging configuration. */
    N_("Log File (required):"),
    /* TRANSLATORS: Part of debugging configuration.  This is a checkbox
    to select if even potentially private data should be included in the
    debug log. */
    N_("Include Mail contents (decrypted!) and meta information."),
    /* TRANSLATORS: Dialog title for the log file selection */
    N_("Select log file"),
    /* TRANSLATORS: Part of debugging configuration. */
    N_("Log level:"),
    /* TRANSLATORS: Part of debugging configuration. Warning shown
       in case the highest log level is selected. Please try to
       keep the string ~ the size of the english version as the
       warning is shown in line with the combo box to select the
       level. */
    N_("<b>Warning:</b> Decreased performance. Huge logs!"),
    /* TRANSLATORS: Config dialog category for debug options. */
    N_("Debug"),
    /* TRANSLATORS: Config dialog category for debug options. */
    N_("Configuaration of debug options"),
    /* TRANSLATORS: Config dialog debug page, can be technical. */
    N_("Potential workarounds"),
    /* TRANSLATORS: Config dialog debug page, can be technical. */
    N_("Block Outlook during encrypt / sign"),
    /* TRANSLATORS: Config dialog debug page, can be technical. */
    N_("Block Outlook during decrypt / verify"),
    /* TRANSLATORS: Config dialog debug page, can be technical. */
    N_("Do not save encrypted mails before decryption"),
    /* TRANSLATORS: Config dialog debug page, link to report bug page. */
    N_("How to report a problem?"),
    /* TRANSLATORS: Config dialog. */
    N_("&Always show security approval dialog"),
    N_("(slow)"),
    /* TRANSLATORS: Config dialog. */
    N_("Help"),
    /* TRANSLATORS: Only available in german and english please keep
       english for other langs */
    N_("handout_outlook_plugin_gnupg_en.pdf"),
    /* TRANSLATORS: Config dialog tooltip. */
    N_("Always show the security approval and certificate selection dialog. "
       "This slows down the encryption / signing process, especially with large keyrings."),
    /* TRANSLATORS: Please ignore. Only available for german. */
    N_("https://gnupg.com/vsd/report.html"),
};


static bool dlg_open;

static DWORD WINAPI
open_gpgolconfig (LPVOID arg)
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
  const auto gpgolconfig = std::string (gpg4win_dir) + "\\bin\\gpgolconfig.exe";
  args.push_back (gpgolconfig);

  args.push_back (std::string ("--hwnd"));
  args.push_back (std::to_string ((int) (intptr_t) wnd));

  args.push_back (std::string("--gpgol-version"));
  args.push_back (std::string(VERSION));

  args.push_back (std::string ("--lang"));
  args.push_back (std::string (gettext_localename ()));

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
                 SRCNAME, __func__, err.code(), err.asStdString().c_str());
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
      log_debug ("%s:%s: Gpgolconfig open. Not launching new dialog.",
                 SRCNAME, __func__);
      HWND optWindow = FindWindow (nullptr, _("Configure GpgOL"));
      if (!optWindow) {
        log_debug ("%s:%s: Gpgolconfig open but could not find window.",
                 SRCNAME, __func__);
        return;
      }
      SetForegroundWindow(optWindow);

      return;
    }

  log_debug ("%s:%s: Launching gpgolconfig.",
             SRCNAME, __func__);

  CloseHandle (CreateThread (NULL, 0, open_gpgolconfig, (LPVOID) parent, 0,
                             NULL));
}
