/* addressbook.cpp - Functions for the Addressbook
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

#include "addressbook.h"

#include "oomhelp.h"
#include "keycache.h"
#include "mail.h"
#include "cpphelp.h"
#include "windowmessages.h"
#include "recipient.h"

#include <gpgme++/context.h>
#include <gpgme++/data.h>

#include <set>
#include <sstream>

typedef struct
{
  std::string name;
  std::string pgp_data;
  std::string cms_data;
  std::string entry_id;
  HWND hwnd;
  shared_disp_t contact;
} keyadder_args_t;

static std::set <std::string> s_checked_entries;
static std::vector <std::string> s_opened_contacts_vec;
GPGRT_LOCK_DEFINE (s_opened_contacts_lock);

static Addressbook::callback_args_t
parse_output (const std::string &output)
{
  Addressbook::callback_args_t ret;

  std::istringstream ss(output);
  std::string line;

  std::string pgp_data;
  std::string cms_data;

  bool in_pgp_data = false;
  bool in_options = false;
  bool in_cms_data = false;
  while (std::getline (ss, line))
    {
      rtrim (line);
      if (in_pgp_data)
        {
          if (line == "empty")
            {
              pgp_data = "";
            }
          else if (line != "END KEYADDER PGP DATA")
            {
              pgp_data += line + std::string("\n");
            }
          else
            {
              in_pgp_data = false;
            }
        }
      else if (in_cms_data)
        {
          if (line == "empty")
            {
              in_cms_data = false;
              cms_data = "";
            }
          else if (line != "END KEYADDER CMS DATA")
            {
              cms_data += line + std::string("\n");
            }
          else
            {
              in_cms_data = false;
            }
        }
      else if (in_options)
        {
          if (line == "END KEYADDER OPTIONS")
            {
              in_options = false;
              continue;
            }
          std::istringstream lss (line);
          std::string key, value;

          std::getline (lss, key, '=');
          std::getline (lss, value, '=');

          if (key == "secure")
            {
              int val = atoi (value.c_str());
              if (val > 3 || val < 0)
                {
                  log_error ("%s:%s: Loading secure value: %s failed",
                             SRCNAME, __func__, value.c_str ());
                  continue;
                }
              ret.crypto_flags = val;
            }
          else
            {
              log_debug ("%s:%s: Unknown setting: %s",
                         SRCNAME, __func__, key.c_str ());
            }
          continue;
        }
      else
        {
          if (line == "BEGIN KEYADDER OPTIONS")
            {
              in_options = true;
            }
          else if (line == "BEGIN KEYADDER CMS DATA")
            {
              in_cms_data = true;
            }
          else if (line == "BEGIN KEYADDER PGP DATA")
            {
              in_pgp_data = true;
            }
          else
            {
              log_debug ("%s:%s: Unknown line: %s",
                         SRCNAME, __func__, line.c_str ());
            }
        }
    }

  ret.pgp_data = xstrdup (pgp_data.c_str ());
  ret.cms_data = xstrdup (cms_data.c_str ());

  return ret;
}

static DWORD WINAPI
open_keyadder (LPVOID arg)
{
  TSTART;
  auto adder_args = std::unique_ptr<keyadder_args_t> ((keyadder_args_t*) arg);

  std::vector<std::string> args;

  // Collect the arguments
  const char *gpg4win_dir = get_gpg4win_dir ();
  if (!gpg4win_dir)
    {
      TRACEPOINT;
      TRETURN -1;
    }
  const auto keyadder = std::string (gpg4win_dir) + "\\bin\\gpgolkeyadder.exe";
  args.push_back (keyadder);

  args.push_back (std::string ("--hwnd"));
  args.push_back (std::to_string ((int) (intptr_t) adder_args->hwnd));

  args.push_back (std::string ("--username"));
  args.push_back (adder_args->name);

  if (opt.enable_smime)
    {
      args.push_back (std::string ("--cms"));
    }

  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      // can't happen
      TRACEPOINT;
      TRETURN -1;
    }

  std::string input = adder_args->pgp_data;
  input += "BEGIN CMS DATA\n";
  input += adder_args->cms_data;

  GpgME::Data mystdin (input.c_str(), input.size(),
                       false);
  GpgME::Data mystdout, mystderr;

  char **cargs = vector_to_cArray (args);
  log_data ("%s:%s: launching keyadder args:", SRCNAME, __func__);
  for (size_t i = 0; cargs && cargs[i]; i++)
    {
      log_data (SIZE_T_FORMAT ": '%s'", i, cargs[i]);
    }

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char**> (cargs),
                                 mystdin, mystdout, mystderr,
                                 (GpgME::Context::SpawnFlags) (
                                  GpgME::Context::SpawnAllowSetFg |
                                  GpgME::Context::SpawnShowWindow));
  release_cArray (cargs);
  if (!adder_args->entry_id.empty ())
    {
      log_dbg ("Removing entry from list of opened contacts.");
      gpgrt_lock_lock (&s_opened_contacts_lock);
      const auto id = adder_args->entry_id;
      s_opened_contacts_vec.erase (
          std::remove_if (s_opened_contacts_vec.begin(),
                          s_opened_contacts_vec.end (),
                          [id] (const auto &val) {
                            return val == id;
                          }));
      gpgrt_lock_unlock (&s_opened_contacts_lock);
    }

  if (err)
    {
      log_error ("%s:%s: Err code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asStdString().c_str());
      TRETURN 0;
    }

  auto output = mystdout.toString ();

  rtrim(output);

  if (output.empty())
    {
      log_debug ("%s:%s: keyadder canceled.", SRCNAME, __func__);
      TRETURN 0;
    }

  Addressbook::callback_args_t cb_args = parse_output (output);
  cb_args.contact = adder_args->contact;

  do_in_ui_thread (CONFIG_KEY_DONE, (void*) &cb_args);
  xfree (cb_args.pgp_data);
  xfree (cb_args.cms_data);
  TRETURN 0;
}

void
Addressbook::update_key_o (void *callback_args)
{
  TSTART;
  if (!callback_args)
    {
      TRACEPOINT;
      TRETURN;
    }
  callback_args_t *cb_args = static_cast<callback_args_t *> (callback_args);
  LPDISPATCH contact = cb_args->contact.get();

  LPDISPATCH user_props = get_oom_object (contact, "UserProperties");
  if (!user_props)
    {
      TRACEPOINT;
      TRETURN;
    }

  LPDISPATCH pgp_key = find_or_add_text_prop (user_props, "OpenPGP Key");
  if (!pgp_key)
    {
      TRACEPOINT;
      gpgol_release (user_props);
      TRETURN;
    }
  put_oom_string (pgp_key, "Value", cb_args->pgp_data);
  gpgol_release (pgp_key);

  log_debug ("%s:%s: PGP key data updated",
             SRCNAME, __func__);

  if (opt.enable_smime)
    {
      LPDISPATCH cms_data = find_or_add_text_prop (user_props,
                                                   "GpgOL CMS Cert");
      if (!cms_data)
        {
          TRACEPOINT;
          gpgol_release (user_props);
          TRETURN;
        }
      put_oom_string (cms_data, "Value", cb_args->cms_data);
      gpgol_release (cms_data);
      log_debug ("%s:%s: CMS key data updated",
                 SRCNAME, __func__);
    }

  gpgol_release (user_props);
  s_checked_entries.clear ();
  TRETURN;
}

void
Addressbook::edit_key_o (LPDISPATCH contact)
{
  TSTART;
  if (!contact)
    {
      TRACEPOINT;
      TRETURN;
    }
  /* First check that we have not already opened an editor.
     This can happen when starting the editor takes a while
     and the user clicks multiple times because nothing
     happens. */
  const auto entry_id = get_oom_string_s (contact, "EntryID");

  if (!entry_id.empty())
    {
      gpgrt_lock_lock (&s_opened_contacts_lock);
      if (std::find (s_opened_contacts_vec.begin (),
                     s_opened_contacts_vec.end (),
                     entry_id) != s_opened_contacts_vec.end ())
        {
          log_dbg ("Contact already opened.");
          /* TODO: Find the window if it exists and bring it to front. */
          gpgrt_lock_unlock (&s_opened_contacts_lock);
          TRETURN;
        }
      gpgrt_lock_unlock (&s_opened_contacts_lock);
    }
  else
    {
      log_dbg ("Empty EntryID.");
    }

  auto user_props = MAKE_SHARED (get_oom_object (contact, "UserProperties"));
  if (!user_props)
    {
      TRACEPOINT;
      TRETURN;
    }

  auto pgp_key = MAKE_SHARED (find_or_add_text_prop (user_props.get (),
                                                     "OpenPGP Key"));

  if (!pgp_key)
    {
      TRACEPOINT;
      TRETURN;
    }

  char *key_data = get_oom_string (pgp_key.get(), "Value");
  if (!key_data)
    {
      TRACEPOINT;
      TRETURN;
    }

  char *cms_data = nullptr;
  if (opt.enable_smime)
    {
      auto cms_key = MAKE_SHARED (find_or_add_text_prop (user_props.get (),
                                                         "GpgOL CMS Cert"));
      cms_data = get_oom_string (cms_key.get(), "Value");
      if (!cms_data)
        {
          TRACEPOINT;
          TRETURN;
        }
    }

  /* Insert into our vec of opened contacts. */
  gpgrt_lock_lock (&s_opened_contacts_lock);
  s_opened_contacts_vec.push_back (entry_id);
  gpgrt_lock_unlock (&s_opened_contacts_lock);

  char *name = get_oom_string (contact, "Subject");
  if (!name)
    {
      TRACEPOINT;
      name = get_oom_string (contact, "Email1Address");
      if (!name)
        {
          name = xstrdup (/* TRANSLATORS: Placeholder for a contact without
                             a configured name */ _("Unknown contact"));
        }
    }

  keyadder_args_t *args = new keyadder_args_t;
  args->name = name;
  args->pgp_data = key_data;
  args->cms_data = cms_data ? cms_data : "";
  args->hwnd = get_active_hwnd ();
  contact->AddRef ();
  memdbg_addRef (contact);
  args->contact = MAKE_SHARED (contact);
  args->entry_id = entry_id;

  CloseHandle (CreateThread (NULL, 0, open_keyadder, (LPVOID) args, 0,
                             NULL));
  xfree (name);
  xfree (key_data);
  xfree (cms_data);

  TRETURN;
}

/* For each new recipient check the address book to look for a potentially
   configured key for this recipient and import / register
   it into the keycache.
*/
void
Addressbook::check_o (Mail *mail)
{
  TSTART;
  if (!mail)
    {
      TRACEPOINT;
      TRETURN;
    }
  LPDISPATCH mailitem = mail->item ();
  if (!mailitem)
    {
      TRACEPOINT;
      TRETURN;
    }
  auto recipients_obj = MAKE_SHARED (get_oom_object (mailitem, "Recipients"));

  if (!recipients_obj)
    {
      TRACEPOINT;
      TRETURN;
    }

  bool err = false;
  const auto recipient_entries = get_oom_recipients_with_addrEntry (recipients_obj.get(),
                                                                    &err);
  for (const auto &pair: recipient_entries)
    {
      const auto mbox = pair.first.mbox ();
      if (s_checked_entries.find (mbox) != s_checked_entries.end ())
        {
          continue;
        }

      if (!pair.second)
        {
          TRACEPOINT;
          continue;
        }

      auto contact = MAKE_SHARED (get_oom_object (pair.second.get (), "GetContact"));
      if (!contact)
        {
          log_debug ("%s:%s: failed to resolve contact for %s",
                     SRCNAME, __func__,
                     anonstr (mbox.c_str()));
          continue;
        }
      s_checked_entries.insert (mbox);

      LPDISPATCH user_props = get_oom_object (contact.get (), "UserProperties");
      if (!user_props)
        {
          TRACEPOINT;
          continue;
        }

      LPDISPATCH pgp_key = find_or_add_text_prop (user_props, "OpenPGP Key");
      LPDISPATCH cms_prop = nullptr;

      if (opt.enable_smime)
        {
          cms_prop = find_or_add_text_prop (user_props, "GpgOL CMS Cert");
        }

      gpgol_release (user_props);

      if (!pgp_key && !cms_prop)
        {
          continue;
        }

      log_debug ("%s:%s: found configured key for %s",
                 SRCNAME, __func__,
                 anonstr (mbox.c_str()));

      char *pgp_data = get_oom_string (pgp_key, "Value");
      char *cms_data = nullptr;
      if (cms_prop)
        {
          cms_data = get_oom_string (cms_prop, "Value");
        }

      if ((!pgp_data || !strlen (pgp_data))
          && (!cms_data || !strlen (cms_data)))
        {
          log_debug ("%s:%s: No key data",
                     SRCNAME, __func__);
        }
      if (pgp_data && strlen (pgp_data))
        {
          KeyCache::instance ()->importFromAddrBook (mbox, pgp_data,
                                                     mail, GpgME::OpenPGP);
        }
      if (cms_data && strlen (cms_data))
        {
          KeyCache::instance ()->importFromAddrBook (mbox, cms_data,
                                                     mail, GpgME::CMS);
        }
      xfree (pgp_data);
      xfree (cms_data);
      gpgol_release (pgp_key);
      gpgol_release (cms_prop);
    }
  TRETURN;
}
