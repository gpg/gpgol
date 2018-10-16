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

#include <gpgme++/context.h>
#include <gpgme++/data.h>

#include <set>

typedef struct
{
  std::string name;
  std::string data;
  HWND hwnd;
  shared_disp_t contact;
} keyadder_args_t;

static std::set <std::string> s_checked_entries;

static DWORD WINAPI
open_keyadder (LPVOID arg)
{
  TSTART;
  auto adder_args = std::unique_ptr<keyadder_args_t> ((keyadder_args_t*) arg);

  std::vector<std::string> args;

  // Collect the arguments
  char *gpg4win_dir = get_gpg4win_dir ();
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

  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);
  if (!ctx)
    {
      // can't happen
      TRACEPOINT;
      TRETURN -1;
    }

  GpgME::Data mystdin (adder_args->data.c_str(), adder_args->data.size(),
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
  if (err)
    {
      log_error ("%s:%s: Err code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
      TRETURN 0;
    }

  auto newKey = mystdout.toString ();

  rtrim(newKey);

  if (newKey.empty())
    {
      log_debug ("%s:%s: keyadder canceled.", SRCNAME, __func__);
      TRETURN 0;
    }
  if (newKey == "empty")
    {
      log_debug ("%s:%s: keyadder empty.", SRCNAME, __func__);
      newKey = "";
    }

  Addressbook::callback_args_t cb_args;

  /* cb args are valid in the same scope as newKey */
  cb_args.data = newKey.c_str();
  cb_args.contact = adder_args->contact;

  do_in_ui_thread (CONFIG_KEY_DONE, (void*) &cb_args);
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
      TRETURN;
    }
  put_oom_string (pgp_key, "Value", cb_args->data);

  log_debug ("%s:%s: PGP key data updated",
             SRCNAME, __func__);

  gpgol_release (pgp_key);

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

  LPDISPATCH user_props = get_oom_object (contact, "UserProperties");
  if (!user_props)
    {
      TRACEPOINT;
      TRETURN;
    }

  auto pgp_key = MAKE_SHARED (
                      find_or_add_text_prop (user_props, "OpenPGP Key"));
  gpgol_release (user_props);

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
  args->data = key_data;
  args->hwnd = get_active_hwnd ();
  contact->AddRef ();
  memdbg_addRef (contact);
  args->contact = MAKE_SHARED (contact);

  CloseHandle (CreateThread (NULL, 0, open_keyadder, (LPVOID) args, 0,
                             NULL));
  xfree (name);
  xfree (key_data);

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
  for (const auto pair: recipient_entries)
    {
      if (s_checked_entries.find (pair.first) != s_checked_entries.end ())
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
                     anonstr (pair.first.c_str()));
          continue;
        }
      s_checked_entries.insert (pair.first);

      LPDISPATCH user_props = get_oom_object (contact.get (), "UserProperties");
      if (!user_props)
        {
          TRACEPOINT;
          continue;
        }

      LPDISPATCH pgp_key = find_or_add_text_prop (user_props, "OpenPGP Key");
      gpgol_release (user_props);

      if (!pgp_key)
        {
          continue;
        }

      log_debug ("%s:%s: found configured pgp key for %s",
                 SRCNAME, __func__,
                 anonstr (pair.first.c_str()));

      char *key_data = get_oom_string (pgp_key, "Value");
      if (!key_data || !strlen (key_data))
        {
          log_debug ("%s:%s: No key data",
                     SRCNAME, __func__);
        }
      KeyCache::instance ()->importFromAddrBook (pair.first, key_data,
                                                 mail);
      xfree (key_data);
      gpgol_release (pgp_key);
    }
  TRETURN;
}
