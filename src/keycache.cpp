/* @file keycache.cpp
 * @brief Internal keycache
 *
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

#include "keycache.h"

#include "common.h"
#include "cpphelp.h"
#include "mail.h"

#include <gpg-error.h>
#include <gpgme++/context.h>
#include <gpgme++/key.h>
#include <gpgme++/data.h>
#include <gpgme++/importresult.h>
#include <gpgme++/engineinfo.h>
#include <gpgme++/defaultassuantransaction.h>

#include <windows.h>

#include <set>
#include <unordered_map>
#include <sstream>

GPGRT_LOCK_DEFINE (keycache_lock);
GPGRT_LOCK_DEFINE (fpr_map_lock);
GPGRT_LOCK_DEFINE (update_lock);
GPGRT_LOCK_DEFINE (import_lock);
static KeyCache* singleton = nullptr;

/** At some point we need to set a limit. There
  seems to be no limit on how many recipients a mail
  can have in outlook.

  We would run out of resources or block.

  50 Threads already seems a bit excessive but
  it should really cover most legit use cases.
*/

#define MAX_LOCATOR_THREADS 50
static int s_thread_cnt;

namespace
{
  class LocateArgs
    {
      public:
        LocateArgs (const std::string& mbox, Mail *mail = nullptr):
          m_mbox (mbox),
          m_mail (mail)
        {
          TSTART;
          s_thread_cnt++;
          Mail::lockDelete ();
          if (Mail::isValidPtr (m_mail))
            {
              m_mail->incrementLocateCount ();
            }
          Mail::unlockDelete ();
          TRETURN;
        };

        ~LocateArgs()
        {
          TSTART;
          s_thread_cnt--;
          Mail::lockDelete ();
          if (Mail::isValidPtr (m_mail))
            {
              m_mail->decrementLocateCount ();
            }
          Mail::unlockDelete ();
          TRETURN;
        }

        std::string m_mbox;
        Mail *m_mail;
    };
} // namespace

typedef std::pair<std::string, GpgME::Protocol> update_arg_t;

typedef std::pair<std::unique_ptr<LocateArgs>, std::string> import_arg_t;

static std::vector<GpgME::Key>
filter_chain (const std::vector<GpgME::Key> &input)
{
   std::vector<GpgME::Key> leaves;

   std::remove_copy_if(input.begin(), input.end(),
           std::back_inserter(leaves),
           [input] (const auto &k)
            {
               /* Check if a key has this fingerprint in the
                * chain ID. Meaning that there is any child of
                * this certificate. In that case remove it. */
               for (const auto &c: input)
               {
                 if (!c.chainID())
                  {
                    continue;
                  }
                 if (!k.primaryFingerprint() || !c.primaryFingerprint())
                  {
                     STRANGEPOINT;
                     continue;
                  }
                 if (!strcmp (c.chainID(), k.primaryFingerprint()))
                  {
                     log_debug ("%s:%s: Filtering %s as non leaf cert",
                                SRCNAME, __func__, k.primaryFingerprint ());
                     return true;
                  }
               }
               return false;
            });
    return leaves;
}

static DWORD WINAPI
do_update (LPVOID arg)
{
  TSTART;
  auto args = std::unique_ptr<update_arg_t> ((update_arg_t*) arg);

  log_debug ("%s:%s updating: \"%s\" with protocol %s",
             SRCNAME, __func__, anonstr (args->first.c_str ()),
             to_cstr (args->second));

  auto ctx = std::unique_ptr<GpgME::Context> (GpgME::Context::createForProtocol
                                              (args->second));

  if (!ctx)
    {
      TRACEPOINT;
      KeyCache::instance ()->onUpdateJobDone (args->first.c_str(),
                                              GpgME::Key ());
      TRETURN 0;
    }

  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Signatures |
                       GpgME::KeyListMode::Validate |
                       GpgME::KeyListMode::WithTofu);
  GpgME::Error err;
  const auto newKey = ctx->key (args->first.c_str (), err, false);
  TRACEPOINT;

  if (newKey.isNull())
    {
      log_debug ("%s:%s Failed to find key for %s",
                 SRCNAME, __func__, anonstr (args->first.c_str ()));
    }
  if (err)
    {
      log_debug ("%s:%s Failed to find key for %s err: %s",
                 SRCNAME, __func__, anonstr (args->first.c_str()),
                 err.asString ());
    }
  KeyCache::instance ()->onUpdateJobDone (args->first.c_str(),
                                          newKey);
  log_debug ("%s:%s Update job done",
             SRCNAME, __func__);
  TRETURN 0;
}

static DWORD WINAPI
do_import (LPVOID arg)
{
  TSTART;
  auto args = std::unique_ptr<import_arg_t> ((import_arg_t*) arg);

  const std::string mbox = args->first->m_mbox;

  log_debug ("%s:%s importing for: \"%s\" with data \n%s",
             SRCNAME, __func__, anonstr (mbox.c_str ()),
             anonstr (args->second.c_str ()));

  // We want to avoid unneccessary copies. The c_str will be valid
  // until args goes out of scope.
  const char *keyStr = args->second.c_str ();
  GpgME::Data data (keyStr, strlen (keyStr), /* copy */ false);

  GpgME::Protocol proto = GpgME::OpenPGP;
  auto type = data.type();
  if (type == GpgME::Data::X509Cert)
    {
      proto = GpgME::CMS;
    }
  data.rewind ();

  auto ctx = GpgME::Context::create(proto);

  if (!ctx)
    {
      TRACEPOINT;
      TRETURN 0;
    }

  if (type != GpgME::Data::PGPKey && type != GpgME::Data::X509Cert)
    {
      log_debug ("%s:%s Data for: %s is not a PGP Key or Cert ",
                 SRCNAME, __func__, anonstr (mbox.c_str ()));
      TRETURN 0;
    }
  data.rewind ();

  const auto result = ctx->importKeys (data);

  std::vector<std::string> fingerprints;
  for (const auto import: result.imports())
    {
      if (import.error())
        {
          log_debug ("%s:%s Error importing: %s",
                     SRCNAME, __func__, import.error().asString());
          continue;
        }
      const char *fpr = import.fingerprint ();
      if (!fpr)
        {
          TRACEPOINT;
          continue;
        }

      update_arg_t * update_args = new update_arg_t;
      update_args->first = std::string (fpr);
      update_args->second = proto;

      // We do it blocking to be sure that when all imports
      // are done they are also part of the keycache.
      do_update ((LPVOID) update_args);

      if (std::find(fingerprints.begin(), fingerprints.end(), fpr) ==
          fingerprints.end())
        {
          fingerprints.push_back (fpr);
        }
      log_debug ("%s:%s Imported: %s from addressbook.",
                 SRCNAME, __func__, anonstr (fpr));
    }

  KeyCache::instance ()->onAddrBookImportJobDone (mbox,
                                                  fingerprints,
                                                  proto);

  log_debug ("%s:%s Import job done for: %s",
             SRCNAME, __func__, anonstr (mbox.c_str ()));
  TRETURN 0;
}

static void
do_populate_protocol (GpgME::Protocol proto, bool secret)
{
  log_debug ("%s:%s: Starting keylisting for proto %s",
             SRCNAME, __func__, to_cstr (proto));
  auto ctx = GpgME::Context::create (proto);
  if (!ctx)
    {
      /* Maybe PGP broken and not S/MIME */
      log_error ("%s:%s: broken installation no ctx.",
                 SRCNAME, __func__);
      TRETURN;
    }

  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Validate);
  ctx->setOffline (true);
  GpgME::Error err;

   if ((err = ctx->startKeyListing ((const char*)nullptr, secret)))
    {
      log_error ("%s:%s: Failed to start keylisting err: %i: %s",
                 SRCNAME, __func__, err.code (), err.asString());
      TRETURN;
    }

  while (!err)
    {
      const auto key = ctx->nextKey(err);
      if (err || key.isNull())
        {
          TRACEPOINT;
          break;
        }
      KeyCache::instance()->onUpdateJobDone (key.primaryFingerprint(),
                                             key);

    }
  TRETURN;
}

const std::vector< std::pair<std::string, std::string> >
gpgagent_transact (std::unique_ptr <GpgME::Context> &ctx,
                   const char *command)
{
  GpgME::Error err;
  err = ctx->assuanTransact (command);
  std::unique_ptr<GpgME::AssuanTransaction> t = ctx->takeLastAssuanTransaction();
  std::unique_ptr<GpgME::DefaultAssuanTransaction> d (dynamic_cast<GpgME::DefaultAssuanTransaction*>(t.release()));

  if (d)
    {
      return d->statusLines ();
    }
  return std::vector< std::pair<std::string, std::string> > ();
}

static void
gpgsm_learn ()
{
  TSTART;
  const GpgME::EngineInfo ei = GpgME::engineInfo (GpgME::CMS);

  if (!ei.fileName ())
    {
      STRANGEPOINT;
      TRETURN;
    }

  std::vector<std::string> args;

  args.push_back (ei.fileName ());
  args.push_back ("--learn-card");

  // Spawn the process
  auto ctx = GpgME::Context::createForEngine (GpgME::SpawnEngine);

  if (!ctx)
    {
      STRANGEPOINT;
      TRETURN;
    }

  GpgME::Data mystdin, mystdout, mystderr;

  char **cargs = vector_to_cArray (args);

  GpgME::Error err = ctx->spawn (cargs[0], const_cast <const char **> (cargs),
                                 mystdin, mystdout, mystderr,
                                 GpgME::Context::SpawnNone);
  release_cArray (cargs);

  if (err)
    {
      log_debug ("%s:%s: gpgsm learn spawn code: %i asString: %s",
                 SRCNAME, __func__, err.code(), err.asString());
    }
  if ((opt.enable_debug & DBG_DATA))
    {
      log_data ("stdout:\n'%s'\nstderr:\n%s", mystdout.toString ().c_str (),
                mystderr.toString ().c_str ());
    }
  TRETURN;
}

static void
do_populate_smartcards (GpgME::Protocol proto)
{
  TSTART;
  if (proto != GpgME::CMS)
    {
      TRETURN;
    }

  GpgME::Error err;
  auto ctx = GpgME::Context::createForEngine (GpgME::AssuanEngine, &err);

  if (err)
    {
      log_dbg ("Failed to create assuan engine. %s",
               err.asString ());
      TRETURN;
    }
  const auto serials = gpgagent_transact (ctx, "scd serialno");
  if (serials.empty ())
    {
      log_dbg ("No smartcard found.");
    }

  const auto pairinfo = gpgagent_transact (ctx, "scd learn --keypairinfo");
  /* If we have a supported card output looks like this:
     S KEYPAIRINFO 84DC19DC8A563302DA14FF33D42FBBB815170E19 NKS-NKS3.4531 sa
     S KEYPAIRINFO CD9C93F2C76CCA935BC47114C5B527491BA4896D NKS-NKS3.45B1 e
     S KEYPAIRINFO 033829E72E9007213FE3E1D9FD3D286DB5D07599 NKS-NKS3.45B2 e
     S KEYPAIRINFO 0404417F832BE6266585190843B38204431EE26A NKS-SIGG.4531 sae
  */
  if (pairinfo.empty ())
    {
      log_dbg ("Did not find any smartcards.");
      TRETURN;
    }

  auto search_ctx = GpgME::Context::create (GpgME::CMS);

  bool need_to_learn = false;
  for (const auto &info: pairinfo)
    {
      if (info.first != "KEYPAIRINFO")
        {
          log_dbg ("Unexpected keypairinfo line '%s'", info.second.c_str ());
          continue;
        }

      const auto vec = gpgol_split (info.second, ' ');
      if (!vec.size ())
        {
          log_dbg ("Unexpected keypairinfo line '%s'", info.second.c_str ());
          continue;
        }
      const auto keygrip = std::string ("&") + vec[0];

      const auto key = search_ctx->key (keygrip.c_str (), err, true);
      if (err || key.isNull ())
        {
          log_debug ("Key for grip: %s not found. Searching.",
                     keygrip.c_str ());
          need_to_learn = true;
          break;
        }
    }

  if (need_to_learn)
    {
      gpgsm_learn ();
    }
  TRETURN;
}

static DWORD WINAPI
do_populate (LPVOID)
{
  TSTART;

  log_debug ("%s:%s: Populating keycache",
             SRCNAME, __func__);
  do_populate_protocol (GpgME::OpenPGP, false);
  do_populate_protocol (GpgME::OpenPGP, true);
  if (opt.enable_smime)
    {
      do_populate_protocol (GpgME::CMS, false);
      do_populate_protocol (GpgME::CMS, true);
      do_populate_smartcards (GpgME::CMS);
    }
  log_debug ("%s:%s: Keycache populated",
             SRCNAME, __func__);

  TRETURN 0;
}


class KeyCache::Private
{
public:
  Private()
  {

  }

  void setPgpKey(const std::string &mbox, const GpgME::Key &key)
  {
    TSTART;
    gpgol_lock (&keycache_lock);
    auto it = m_pgp_key_map.find (mbox);

    if (it == m_pgp_key_map.end ())
      {
        m_pgp_key_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    insertOrUpdateInFprMap (key);
    gpgol_unlock (&keycache_lock);
    TRETURN;
  }

  void setSmimeKey(const std::string &mbox, const GpgME::Key &key)
  {
    TSTART;
    gpgol_lock (&keycache_lock);
    auto it = m_smime_key_map.find (mbox);

    if (it == m_smime_key_map.end ())
      {
        m_smime_key_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    insertOrUpdateInFprMap (key);
    gpgol_unlock (&keycache_lock);
    TRETURN;
  }

  time_t getLastSubkeyCreation(const GpgME::Key &k)
  {
    TSTART;
    time_t ret = 0;

    for (const auto &sub: k.subkeys())
      {
        if (sub.isBad())
          {
            continue;
          }
        if (sub.creationTime() > ret)
          {
            ret = sub.creationTime();
          }
      }
    TRETURN ret;
  }

  GpgME::Key compareSkeys(const GpgME::Key &old, const GpgME::Key &newKey)
  {
    TSTART;

    if (newKey.isNull())
      {
        return old;
      }
    if (old.isNull())
      {
        return newKey;
      }

    if (old.primaryFingerprint() && newKey.primaryFingerprint() &&
        !strcmp (old.primaryFingerprint(), newKey.primaryFingerprint()))
      {
        // Both are the same. Take the newer one.
        return newKey;
      }

    if (old.canSign() && !newKey.canSign())
      {
        log_debug ("%s:%s Keeping old skey with "
                   "fpr %s over %s because it can sign.",
                   SRCNAME, __func__,
                   anonstr (old.primaryFingerprint()),
                   anonstr (newKey.primaryFingerprint()));
        TRETURN old;
      }

    if (!old.canSign() && newKey.canSign())
      {
        log_debug ("%s:%s Using new skey with "
                   "fpr %s over %s because it can sign.",
                   SRCNAME, __func__,
                   anonstr (newKey.primaryFingerprint()),
                   anonstr (old.primaryFingerprint()));
        TRETURN newKey;
      }

    // Both can or can't sign. Use the newest one.
    if (getLastSubkeyCreation (old) >= getLastSubkeyCreation (newKey))
      {
        log_debug ("%s:%s Keeping old skey with "
                   "fpr %s over %s because it is newer.",
                   SRCNAME, __func__,
                   anonstr (old.primaryFingerprint()),
                   anonstr (newKey.primaryFingerprint()));
        TRETURN old;
      }
    log_debug ("%s:%s Using new skey with "
               "fpr %s over %s because it is newer.",
               SRCNAME, __func__,
               anonstr (newKey.primaryFingerprint()),
               anonstr (old.primaryFingerprint()));
    TRETURN newKey;
  }

  void setPgpKeySecret(const std::string &mbox, const GpgME::Key &key,
                       bool insert = true)
  {
    TSTART;
    gpgol_lock (&keycache_lock);
    auto it = m_pgp_skey_map.find (mbox);

    if (it == m_pgp_skey_map.end ())
      {
        m_pgp_skey_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = compareSkeys (it->second, key);
      }
    if (insert)
      {
        insertOrUpdateInFprMap (key);
      }
    gpgol_unlock (&keycache_lock);
    TRETURN;
  }

  void setSmimeKeySecret(const std::string &mbox, const GpgME::Key &key,
                         bool insert = true)
  {
    TSTART;
    gpgol_lock (&keycache_lock);
    auto it = m_smime_skey_map.find (mbox);

    if (it == m_smime_skey_map.end ())
      {
        m_smime_skey_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = compareSkeys (it->second, key);
      }
    if (insert)
      {
        insertOrUpdateInFprMap (key);
      }
    gpgol_unlock (&keycache_lock);
    TRETURN;
  }

  std::vector<GpgME::Key> getOverrides (const char *addr,
                                        GpgME::Protocol proto = GpgME::OpenPGP)
  {
    TSTART;
    std::vector<GpgME::Key> ret;

    if (!addr)
      {
        TRETURN ret;
      }
    auto mbox = GpgME::UserID::addrSpecFromString (addr);

    gpgol_lock (&import_lock);
    const auto job_set = (proto == GpgME::OpenPGP ?
                       &m_pgp_import_jobs : &m_cms_import_jobs);
    int i = 0;
    while (job_set->find (mbox) != job_set->end ())
      {
        i++;
        if (i % 100 == 0)
          {
            log_debug ("%s:%s Waiting on import for \"%s\"",
                       SRCNAME, __func__, anonstr (addr));
          }
        gpgol_unlock (&import_lock);
        Sleep (10);
        gpgol_lock (&import_lock);
        if (i == 1000)
          {
            /* Just to be on the save side */
            log_error ("%s:%s Waiting on import for \"%s\" "
                       "failed! Bug!",
                       SRCNAME, __func__, anonstr (addr));
            break;
          }
      }
    gpgol_unlock (&import_lock);

    auto override_map = (proto == GpgME::OpenPGP ?
                         &m_pgp_overrides : &m_cms_overrides);
    const auto it = override_map->find (mbox);
    if (it == override_map->end ())
      {
        gpgol_unlock (&keycache_lock);
        TRETURN ret;
      }
    for (const auto fpr: it->second)
      {
        const auto key = getByFpr (fpr.c_str (), false);
        if (key.isNull())
          {
            log_debug ("%s:%s: No key for %s in the cache?!",
                       SRCNAME, __func__, anonstr (fpr.c_str()));
            continue;
          }
        ret.push_back (key);
      }
    if (proto == GpgME::CMS) {
        /* Remove root and intermediate ca's */
        ret = filter_chain (ret);
    }
    gpgol_unlock (&keycache_lock);
    TRETURN ret;
  }

  GpgME::Key getKey (const char *addr, GpgME::Protocol proto)
  {
    TSTART;
    if (!addr)
      {
        TRETURN GpgME::Key();
      }
    auto mbox = GpgME::UserID::addrSpecFromString (addr);

    if (proto == GpgME::OpenPGP)
      {
        gpgol_lock (&keycache_lock);
        const auto it = m_pgp_key_map.find (mbox);

        if (it == m_pgp_key_map.end ())
          {
            gpgol_unlock (&keycache_lock);
            TRETURN GpgME::Key();
          }
        const auto ret = it->second;
        gpgol_unlock (&keycache_lock);

        TRETURN ret;
      }
    gpgol_lock (&keycache_lock);
    const auto it = m_smime_key_map.find (mbox);

    if (it == m_smime_key_map.end ())
      {
        gpgol_unlock (&keycache_lock);
        TRETURN GpgME::Key();
      }
    const auto ret = it->second;
    gpgol_unlock (&keycache_lock);

    TRETURN ret;
  }

  GpgME::Key getSKey (const char *addr, GpgME::Protocol proto)
  {
    TSTART;
    if (!addr)
      {
        TRETURN GpgME::Key();
      }
    auto mbox = GpgME::UserID::addrSpecFromString (addr);

    if (proto == GpgME::OpenPGP)
      {
        gpgol_lock (&keycache_lock);
        const auto it = m_pgp_skey_map.find (mbox);

        if (it == m_pgp_skey_map.end ())
          {
            gpgol_unlock (&keycache_lock);
            TRETURN GpgME::Key();
          }
        const auto ret = it->second;
        gpgol_unlock (&keycache_lock);

        TRETURN ret;
      }
    gpgol_lock (&keycache_lock);
    const auto it = m_smime_skey_map.find (mbox);

    if (it == m_smime_skey_map.end ())
      {
        gpgol_unlock (&keycache_lock);
        TRETURN GpgME::Key();
      }
    const auto ret = it->second;
    gpgol_unlock (&keycache_lock);

    TRETURN ret;
  }

  GpgME::Key getSigningKey (const char *addr, GpgME::Protocol proto)
  {
    TSTART;
    const auto key = getSKey (addr, proto);
    if (key.isNull())
      {
        log_debug ("%s:%s: secret key for %s is null",
                   SRCNAME, __func__, anonstr (addr));
        TRETURN key;
      }
    if (!key.canReallySign())
      {
        log_debug ("%s:%s: Discarding key for %s because it can't sign",
                   SRCNAME, __func__, anonstr (addr));
        TRETURN GpgME::Key();
      }
    if (!key.hasSecret())
      {
        log_debug ("%s:%s: Discarding key for %s because it has no secret",
                   SRCNAME, __func__, anonstr (addr));
        TRETURN GpgME::Key();
      }
    if (in_de_vs_mode () && !key.isDeVs())
      {
        log_debug ("%s:%s: signing key for %s is not deVS",
                   SRCNAME, __func__, anonstr (addr));
        TRETURN GpgME::Key();
      }
    TRETURN key;
  }

  std::vector<GpgME::Key> getEncryptionKeys (const std::vector<std::string>
                                             &recipients,
                                             GpgME::Protocol proto)
  {
    TSTART;
    std::vector<GpgME::Key> ret;
    if (recipients.empty ())
      {
        TRACEPOINT;
        TRETURN ret;
      }
    for (const auto &recip: recipients)
      {
        const auto overrides = getOverrides (recip.c_str (), proto);

        if (!overrides.empty())
          {
            const auto filtered = (proto == GpgME::CMS ? filter_chain(overrides)
                                                       : overrides);
            ret.insert (ret.end (), filtered.begin (), filtered.end ());
            log_debug ("%s:%s: Using overrides for %s",
                       SRCNAME, __func__, anonstr (recip.c_str ()));
            continue;
          }
        const auto key = getKey (recip.c_str (), proto);
        if (key.isNull())
          {
            log_debug ("%s:%s: No key for %s in proto %s. no internal encryption",
                       SRCNAME, __func__, anonstr (recip.c_str ()),
                       to_cstr (proto));
            TRETURN std::vector<GpgME::Key>();
          }

        if (!key.canEncrypt() || key.isRevoked() ||
            key.isExpired() || key.isDisabled() || key.isInvalid())
          {
            log_data ("%s:%s: Invalid key for %s. no internal encryption",
                       SRCNAME, __func__, anonstr (recip.c_str ()));
            TRETURN std::vector<GpgME::Key>();
          }

        if (in_de_vs_mode () && !key.isDeVs ())
          {
            log_data ("%s:%s: key for %s is not deVS",
                      SRCNAME, __func__, anonstr (recip.c_str ()));
            TRETURN std::vector<GpgME::Key>();
          }

        bool validEnough = false;
        /* Here we do the check if the key is valid for this recipient */
        const auto addrSpec = GpgME::UserID::addrSpecFromString (recip.c_str ());
        for (const auto &uid: key.userIDs ())
          {
            if (addrSpec != uid.addrSpec())
              {
                // Ignore unmatching addr specs
                continue;
              }
            if (uid.validity() >= GpgME::UserID::Marginal ||
                uid.origin() == GpgME::Key::OriginWKD)
              {
                validEnough = true;
                break;
              }
            if (opt.auto_unstrusted &&
                uid.validity() == GpgME::UserID::Unknown)
              {
                log_debug ("%s:%s: Passing unknown trust key for %s because of option",
                           SRCNAME, __func__, anonstr (recip.c_str ()));
                validEnough = true;
                break;
              }
          }
        if (!validEnough)
          {
            log_debug ("%s:%s: UID for %s does not have at least marginal trust",
                       SRCNAME, __func__, anonstr (recip.c_str ()));
            TRETURN std::vector<GpgME::Key>();
          }
        // Accepting key
        ret.push_back (key);
      }
    TRETURN ret;
  }

  void insertOrUpdateInFprMap (const GpgME::Key &key)
    {
      TSTART;
      if (key.isNull() || !key.primaryFingerprint())
        {
          TRACEPOINT;
          TRETURN;
        }
      gpgol_lock (&fpr_map_lock);

      /* First ensure that we have the subkeys mapped to the primary
         fpr */
      const char *primaryFpr = key.primaryFingerprint ();

#if 0
        {
          std::stringstream ss;
          ss << key;
          log_debug ("%s:%s: Inserting key\n%s",
                     SRCNAME, __func__, ss.str().c_str ());
        }
#endif

      for (const auto &sub: key.subkeys())
        {
          const char *subFpr = sub.fingerprint();
          auto it = m_sub_fpr_map.find (subFpr);
          if (it == m_sub_fpr_map.end ())
            {
              m_sub_fpr_map.insert (std::make_pair(
                                     std::string (subFpr),
                                     std::string (primaryFpr)));
            }
        }

      auto it = m_fpr_map.find (primaryFpr);

      if (it == m_fpr_map.end ())
        {
          m_fpr_map.insert (std::make_pair (primaryFpr, key));

          gpgol_unlock (&fpr_map_lock);
          TRETURN;
        }

      for (const auto &uid: key.userIDs())
        {
          if (key.isBad() || uid.isBad())
            {
              continue;
            }
          /* Update ultimate keys map */
          if (uid.validity() == GpgME::UserID::Validity::Ultimate &&
              uid.id())
            {
              const char *fpr = key.primaryFingerprint();
              if (!fpr)
                {
                  STRANGEPOINT;
                  continue;
                }
              TRACEPOINT;
              m_ultimate_keys.erase (std::remove_if (m_ultimate_keys.begin(),
                                     m_ultimate_keys.end(),
                                     [fpr] (const GpgME::Key &ult)
                {
                  return ult.primaryFingerprint() && !strcmp (fpr, ult.primaryFingerprint());
                }), m_ultimate_keys.end());
              TRACEPOINT;
              m_ultimate_keys.push_back (key);
            }

          /* Update skey maps */
          if (key.hasSecret ())
            {
              if (key.protocol () == GpgME::OpenPGP)
                {
                  setPgpKeySecret (uid.addrSpec(), key, false);
                }
              else if (key.protocol () == GpgME::CMS)
                {
                  setSmimeKeySecret (uid.addrSpec(), key, false);
                }
              else
                {
                  STRANGEPOINT;
                }
            }
        }

      if (it->second.hasSecret () && !key.hasSecret())
        {
          log_debug ("%s:%s Lost secret info on update. Merging.",
                     SRCNAME, __func__);
          auto merged = key;
          merged.mergeWith (it->second);
          it->second = merged;
        }
      else
        {
          it->second = key;
        }
      gpgol_unlock (&fpr_map_lock);
      TRETURN;
    }

  GpgME::Key getFromMap (const char *fpr) const
  {
    TSTART;
    if (!fpr)
      {
        TRACEPOINT;
        TRETURN GpgME::Key();
      }

    gpgol_lock (&fpr_map_lock);
    std::string primaryFpr;
    const auto it = m_sub_fpr_map.find (fpr);
    if (it != m_sub_fpr_map.end ())
      {
        log_debug ("%s:%s using \"%s\" for \"%s\"",
                   SRCNAME, __func__, anonstr (it->second.c_str()),
                   anonstr (fpr));
        primaryFpr = it->second;
      }
    else
      {
        primaryFpr = fpr;
      }

    const auto keyIt = m_fpr_map.find (primaryFpr);
    if (keyIt != m_fpr_map.end ())
      {
        const auto ret = keyIt->second;
        gpgol_unlock (&fpr_map_lock);
        TRETURN ret;
      }
    gpgol_unlock (&fpr_map_lock);
    TRETURN GpgME::Key();
  }

  GpgME::Key getByFpr (const char *fpr, bool block) const
    {
      TSTART;
      if (!fpr)
        {
          TRACEPOINT;
          TRETURN GpgME::Key ();
        }

      TRACEPOINT;
      const auto ret = getFromMap (fpr);
      if (ret.isNull())
        {
          // If the key was not found we need to check if there is
          // an update running.
          if (block)
            {
              const std::string sFpr (fpr);
              int i = 0;

              gpgol_lock (&update_lock);
              while (m_update_jobs.find(sFpr) != m_update_jobs.end ())
                {
                  i++;
                  if (i % 100 == 0)
                    {
                      log_debug ("%s:%s Waiting on update for \"%s\"",
                                 SRCNAME, __func__, anonstr (fpr));
                    }
                  gpgol_unlock (&update_lock);
                  Sleep (10);
                  gpgol_lock (&update_lock);
                  if (i == 3000)
                    {
                      /* Just to be on the save side */
                      log_error ("%s:%s Waiting on update for \"%s\" "
                                 "failed! Bug!",
                                 SRCNAME, __func__, anonstr (fpr));
                      break;
                    }
                }
              gpgol_unlock (&update_lock);

              TRACEPOINT;
              const auto ret2 = getFromMap (fpr);
              if (ret2.isNull ())
                {
                  log_debug ("%s:%s Cache miss after blocking check %s.",
                             SRCNAME, __func__, anonstr (fpr));
                }
              else
                {
                  log_debug ("%s:%s Cache hit after wait for %s.",
                             SRCNAME, __func__, anonstr (fpr));
                  TRETURN ret2;
                }
            }
          log_debug ("%s:%s Cache miss for %s.",
                     SRCNAME, __func__, anonstr (fpr));
          TRETURN GpgME::Key();
        }

      log_debug ("%s:%s Cache hit for %s.",
                 SRCNAME, __func__, anonstr (fpr));
      TRETURN ret;
    }

  void update (const char *fpr, GpgME::Protocol proto)
     {
       TSTART;
       if (!fpr)
         {
           TRETURN;
         }
       const std::string sFpr (fpr);
       gpgol_lock (&update_lock);
       if (m_update_jobs.find(sFpr) != m_update_jobs.end ())
         {
           log_debug ("%s:%s Update for \"%s\" already in progress.",
                      SRCNAME, __func__, anonstr (fpr));
           gpgol_unlock (&update_lock);
         }

       m_update_jobs.insert (sFpr);
       gpgol_unlock (&update_lock);
       update_arg_t * args = new update_arg_t;
       args->first = sFpr;
       args->second = proto;
       CloseHandle (CreateThread (NULL, 0, do_update,
                                  (LPVOID) args, 0,
                                  NULL));
       TRETURN;
     }

  void onUpdateJobDone (const char *fpr, const GpgME::Key &key)
    {
      TSTART;
      if (!fpr)
        {
          TRETURN;
        }
      TRACEPOINT;
      insertOrUpdateInFprMap (key);
      gpgol_lock (&update_lock);
      const auto it = m_update_jobs.find(fpr);

      if (it == m_update_jobs.end())
        {
          gpgol_unlock (&update_lock);
          TRETURN;
        }
      m_update_jobs.erase (it);
      gpgol_unlock (&update_lock);
      TRACEPOINT;
      TRETURN;
    }

  void importFromAddrBook (const std::string &mbox, const char *data,
                           Mail *mail, GpgME::Protocol proto)
    {
      TSTART;
      if (!data || mbox.empty() || !mail)
        {
          TRACEPOINT;
          TRETURN;
        }

      std::string sdata (data);
      trim (sdata);
      if (sdata.empty())
        {
          TRETURN;
        }
      gpgol_lock (&import_lock);
      auto job_set = (proto == GpgME::OpenPGP ?
                       &m_pgp_import_jobs : &m_cms_import_jobs);
      if (job_set->find (mbox) != job_set->end ())
        {
          log_debug ("%s:%s import for \"%s\" %s already in progress.",
                     SRCNAME, __func__, anonstr (mbox.c_str ()),
                     to_cstr (proto));
          gpgol_unlock (&import_lock);
        }
      job_set->insert (mbox);
      gpgol_unlock (&import_lock);

      import_arg_t * args = new import_arg_t;
      args->first = std::unique_ptr<LocateArgs> (new LocateArgs (mbox, mail));
      args->second = sdata;
      CloseHandle (CreateThread (NULL, 0, do_import,
                                 (LPVOID) args, 0,
                                 NULL));

      TRETURN;
    }

  void onAddrBookImportJobDone (const std::string &mbox,
                                const std::vector<std::string> &result_fprs,
                                GpgME::Protocol proto)
    {
      TSTART;
      gpgol_lock (&keycache_lock);
      auto override_map = (proto == GpgME::OpenPGP ?
                           &m_pgp_overrides : &m_cms_overrides);
      auto job_set = (proto == GpgME::OpenPGP ?
                       &m_pgp_import_jobs : &m_cms_import_jobs);

      auto it = override_map->find (mbox);
      if (it != override_map->end ())
        {
          it->second = result_fprs;
        }
      else
        {
          override_map->insert (std::make_pair (mbox, result_fprs));
        }
      gpgol_unlock (&keycache_lock);
      gpgol_lock (&import_lock);
      const auto job_it = job_set->find(mbox);

      if (job_it == job_set->end())
        {
          log_error ("%s:%s import for \"%s\" %s already finished.",
                     SRCNAME, __func__, anonstr (mbox.c_str ()),
                     to_cstr (proto));
          gpgol_unlock (&import_lock);
          TRETURN;
        }
      job_set->erase (job_it);
      gpgol_unlock (&import_lock);
      TRETURN;
    }

  void populate ()
    {
      TSTART;
      gpgrt_lock_lock (&keycache_lock);
      m_ultimate_keys.clear ();
      gpgrt_lock_unlock (&keycache_lock);
      CloseHandle (CreateThread (nullptr, 0, do_populate,
                                 nullptr, 0,
                                 nullptr));
      TRETURN;
    }

  std::unordered_map<std::string, GpgME::Key> m_pgp_key_map;
  std::unordered_map<std::string, GpgME::Key> m_smime_key_map;
  std::unordered_map<std::string, GpgME::Key> m_pgp_skey_map;
  std::unordered_map<std::string, GpgME::Key> m_smime_skey_map;
  std::unordered_map<std::string, GpgME::Key> m_fpr_map;
  std::unordered_map<std::string, std::string> m_sub_fpr_map;
  std::unordered_map<std::string, std::vector<std::string> >
    m_pgp_overrides;
  std::unordered_map<std::string, std::vector<std::string> >
    m_cms_overrides;
  std::vector<GpgME::Key> m_ultimate_keys;
  std::set<std::string> m_update_jobs;
  std::set<std::string> m_pgp_import_jobs;
  std::set<std::string> m_cms_import_jobs;
};

KeyCache::KeyCache():
  d(new Private)
{
}

KeyCache *
KeyCache::instance ()
{
  if (!singleton)
    {
      singleton = new KeyCache();
    }
  return singleton;
}

GpgME::Key
KeyCache::getSigningKey (const char *addr, GpgME::Protocol proto) const
{
  return d->getSigningKey (addr, proto);
}

std::vector<GpgME::Key>
KeyCache::getEncryptionKeys (const std::vector<std::string> &recipients, GpgME::Protocol proto) const
{
  return d->getEncryptionKeys (recipients, proto);
}

std::vector<GpgME::Key>
KeyCache::getEncryptionKeys (const std::string &recipient, GpgME::Protocol proto) const
{
  std::vector<std::string> vec;
  vec.push_back (recipient);
  return d->getEncryptionKeys (vec, proto);
}

static GpgME::Key
get_most_valid_key_simple (const std::vector<GpgME::Key> &keys)
{
  GpgME::Key candidate;
  for (const auto &key: keys)
    {
      if (key.isRevoked() || key.isExpired() ||
          key.isDisabled() || key.isInvalid())
        {
          log_debug ("%s:%s: Skipping invalid S/MIME key",
                     SRCNAME, __func__);
          continue;
        }
      if (candidate.isNull() || !candidate.numUserIDs())
        {
          if (key.numUserIDs() &&
              candidate.userID(0).validity() <= key.userID(0).validity())
            {
              candidate = key;
            }
        }
    }
  return candidate;
}

static std::vector<GpgME::Key>
get_local_smime_keys (const std::string &addr)
{
  TSTART;
  std::vector<GpgME::Key> keys;
  auto ctx = std::unique_ptr<GpgME::Context> (
                                              GpgME::Context::createForProtocol (GpgME::CMS));
  if (!ctx)
    {
      TRACEPOINT;
      TRETURN keys;
    }
  // We need to validate here to fetch CRL's
  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Validate |
                       GpgME::KeyListMode::Signatures);
  GpgME::Error e = ctx->startKeyListing (addr.c_str());
  if (e)
    {
      TRACEPOINT;
      TRETURN keys;
    }

  GpgME::Error err;
  do {
      keys.push_back(ctx->nextKey(err));
  } while (!err);
  keys.pop_back();

  TRETURN keys;
}

static std::vector<GpgME::Key>
get_extern_smime_keys (const std::string &addr, bool import)
{
  TSTART;
  std::vector<GpgME::Key> keys;
  auto ctx = std::unique_ptr<GpgME::Context> (
                                              GpgME::Context::createForProtocol (GpgME::CMS));
  if (!ctx)
    {
      TRACEPOINT;
      TRETURN keys;
    }
  // We need to validate here to fetch CRL's
  ctx->setKeyListMode (GpgME::KeyListMode::Extern);
  GpgME::Error e = ctx->startKeyListing (addr.c_str());
  if (e)
    {
      TRACEPOINT;
      TRETURN keys;
    }

  GpgME::Error err;
  do
    {
      const auto key = ctx->nextKey (err);
      if (!err && !key.isNull())
        {
          keys.push_back (key);
          log_debug ("%s:%s: Found extern S/MIME key for %s with fpr: %s",
                     SRCNAME, __func__, anonstr (addr.c_str()),
                     anonstr (key.primaryFingerprint()));
        }
    } while (!err);

  if (import && keys.size ())
    {
      const GpgME::ImportResult res = ctx->importKeys(keys);
      log_debug ("%s:%s: Import result for %s: err: %s",
                 SRCNAME, __func__, anonstr (addr.c_str()),
                 res.error ().asString ());

    }

  TRETURN keys;
}

static DWORD WINAPI
do_locate (LPVOID arg)
{
  TSTART;
  if (!arg)
    {
      TRETURN 0;
    }

  auto args = std::unique_ptr<LocateArgs> ((LocateArgs *) arg);

  const auto addr = args->m_mbox;

  log_debug ("%s:%s searching key for addr: \"%s\"",
             SRCNAME, __func__, anonstr (addr.c_str()));

  const auto k = GpgME::Key::locate (addr.c_str());

  if (!k.isNull ())
    {
      log_debug ("%s:%s found key for addr: \"%s\":%s",
                 SRCNAME, __func__, anonstr (addr.c_str()),
                 anonstr (k.primaryFingerprint()));
      KeyCache::instance ()->setPgpKey (addr, k);
    }
  log_debug ("%s:%s pgp locate done",
             SRCNAME, __func__);

  if (opt.enable_smime)
    {
      GpgME::Key candidate = get_most_valid_key_simple (
                                    get_local_smime_keys (addr));
      if (!candidate.isNull())
        {
          log_debug ("%s:%s found SMIME key for addr: \"%s\":%s",
                     SRCNAME, __func__, anonstr (addr.c_str()),
                     anonstr (candidate.primaryFingerprint()));
          KeyCache::instance()->setSmimeKey (addr, candidate);
          TRETURN 0;
        }
      if (!opt.search_smime_servers || (!k.isNull() && !opt.prefer_smime))
        {
          log_debug ("%s:%s Found no S/MIME key locally and external "
                     "search is disabled.", SRCNAME, __func__);
          TRETURN 0;
        }
      /* Search for extern keys and import them */
      const auto externs = get_extern_smime_keys (addr, true);
      if (externs.empty())
        {
          TRETURN 0;
        }
      /* We found and imported external keys. We need to get them
         locally now to ensure that they are valid etc. */
      candidate = get_most_valid_key_simple (
                                    get_local_smime_keys (addr));
      if (!candidate.isNull())
        {
          log_debug ("%s:%s found ext. SMIME key for addr: \"%s\":%s",
                     SRCNAME, __func__, anonstr (addr.c_str()),
                     anonstr (candidate.primaryFingerprint()));
          KeyCache::instance()->setSmimeKey (addr, candidate);
          TRETURN 0;
        }
      else
        {
          log_debug ("%s:%s: Found no valid key in extern S/MIME certs",
                     SRCNAME, __func__);
        }
    }
  TRETURN 0;
}

static void
locate_secret (const char *addr, GpgME::Protocol proto)
{
  TSTART;
  auto ctx = std::unique_ptr<GpgME::Context> (
                      GpgME::Context::createForProtocol (proto));
  if (!ctx)
    {
      TRACEPOINT;
      TRETURN;
    }
  if (!addr)
    {
      TRACEPOINT;
      TRETURN;
    }
  const auto mbox = GpgME::UserID::addrSpecFromString (addr);

  if (mbox.empty())
    {
      log_debug ("%s:%s: Empty mbox for addr %s",
                 SRCNAME, __func__, anonstr (addr));
      TRETURN;
    }

  // We need to validate here to fetch CRL's
  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Validate);
  GpgME::Error e = ctx->startKeyListing (mbox.c_str(), true);
  if (e)
    {
      TRACEPOINT;
      TRETURN;
    }

  std::vector<GpgME::Key> keys;
  GpgME::Error err;
  do
    {
      const auto key = ctx->nextKey(err);
      if (key.isNull())
        {
          continue;
        }
      if (key.isRevoked() || key.isExpired() ||
          key.isDisabled() || key.isInvalid())
        {
          if ((opt.enable_debug & DBG_DATA))
            {
              std::stringstream ss;
              ss << key;
              log_data ("%s:%s: Skipping invalid secret key %s",
                        SRCNAME, __func__, ss.str().c_str());
            }
          continue;
        }
      if (proto == GpgME::OpenPGP)
        {
          log_debug ("%s:%s found pgp skey for addr: \"%s\":%s",
                     SRCNAME, __func__, anonstr (mbox.c_str()),
                     anonstr (key.primaryFingerprint()));
          KeyCache::instance()->setPgpKeySecret (mbox, key);
          TRETURN;
        }
      if (proto == GpgME::CMS)
        {
          log_debug ("%s:%s found cms skey for addr: \"%s\":%s",
                     SRCNAME, __func__, anonstr (mbox.c_str ()),
                     anonstr (key.primaryFingerprint()));
          KeyCache::instance()->setSmimeKeySecret (mbox, key);
          TRETURN;
        }
    } while (!err);
  TRETURN;
}

static DWORD WINAPI
do_locate_secret (LPVOID arg)
{
  TSTART;
  auto args = std::unique_ptr<LocateArgs> ((LocateArgs *) arg);

  log_debug ("%s:%s searching secret key for addr: \"%s\"",
             SRCNAME, __func__, anonstr (args->m_mbox.c_str ()));

  locate_secret (args->m_mbox.c_str(), GpgME::OpenPGP);
  if (opt.enable_smime)
    {
      locate_secret (args->m_mbox.c_str(), GpgME::CMS);
    }
  log_debug ("%s:%s locator sthread thread done",
             SRCNAME, __func__);
  TRETURN 0;
}

void
KeyCache::startLocate (const std::vector<std::string> &addrs, Mail *mail) const
{
  for (const auto &addr: addrs)
    {
      startLocate (addr.c_str(), mail);
    }
}

void
KeyCache::startLocate (const char *addr, Mail *mail) const
{
  TSTART;
  if (!addr)
    {
      TRACEPOINT;
      TRETURN;
    }
  std::string recp = GpgME::UserID::addrSpecFromString (addr);
  if (recp.empty ())
    {
      TRETURN;
    }
  gpgol_lock (&keycache_lock);
  if (d->m_pgp_key_map.find (recp) == d->m_pgp_key_map.end ())
    {
      // It's enough to look at the PGP Key map. We marked
      // searched keys there.
      d->m_pgp_key_map.insert (std::pair<std::string, GpgME::Key> (recp, GpgME::Key()));
      log_debug ("%s:%s Creating a locator thread",
                 SRCNAME, __func__);
      const auto args = new LocateArgs(recp, mail);
      HANDLE thread = CreateThread (NULL, 0, do_locate,
                                    args, 0,
                                    NULL);
      CloseHandle (thread);
    }
  gpgol_unlock (&keycache_lock);
  TRETURN;
}

void
KeyCache::startLocateSecret (const char *addr, Mail *mail) const
{
  TSTART;
  if (!addr)
    {
      TRACEPOINT;
      TRETURN;
    }
  std::string recp = GpgME::UserID::addrSpecFromString (addr);
  if (recp.empty ())
    {
      TRETURN;
    }
  gpgol_lock (&keycache_lock);
  if (d->m_pgp_skey_map.find (recp) == d->m_pgp_skey_map.end ())
    {
      // It's enough to look at the PGP Key map. We marked
      // searched keys there.
      d->m_pgp_skey_map.insert (std::pair<std::string, GpgME::Key> (recp, GpgME::Key()));
      log_debug ("%s:%s Creating a locator thread",
                 SRCNAME, __func__);
      const auto args = new LocateArgs(recp, mail);

      HANDLE thread = CreateThread (NULL, 0, do_locate_secret,
                                    (LPVOID) args, 0,
                                    NULL);
      CloseHandle (thread);
    }
  gpgol_unlock (&keycache_lock);
  TRETURN;
}


void
KeyCache::setSmimeKey(const std::string &mbox, const GpgME::Key &key)
{
  d->setSmimeKey(mbox, key);
}

void
KeyCache::setPgpKey(const std::string &mbox, const GpgME::Key &key)
{
  d->setPgpKey(mbox, key);
}

void
KeyCache::setSmimeKeySecret(const std::string &mbox, const GpgME::Key &key)
{
  d->setSmimeKeySecret(mbox, key);
}

void
KeyCache::setPgpKeySecret(const std::string &mbox, const GpgME::Key &key)
{
  d->setPgpKeySecret(mbox, key);
}

bool
KeyCache::isMailResolvable(Mail *mail)
{
  TSTART;
  /* Get the data from the mail. */
  const auto sender = mail->getSender ();
  auto recps = mail->getCachedRecipientAddresses ();

  if (sender.empty() || recps.empty())
    {
      log_debug ("%s:%s: Mail has no sender or no recipients.",
                 SRCNAME, __func__);
      TRETURN false;
    }


  GpgME::Key sigKey = getSigningKey (sender.c_str(), GpgME::OpenPGP);
  std::vector<GpgME::Key> encKeys = getEncryptionKeys (recps,
                                                       GpgME::OpenPGP);

  /* If S/MIME is prefrerred we only toggle auto encrypt for PGP if
     we both have a signing key and encryption keys. */
  if (!encKeys.empty() && (!opt.prefer_smime || !sigKey.isNull()))
    {
      TRETURN true;
    }

  if (!opt.enable_smime)
    {
      TRETURN false;
    }

  /* Check S/MIME instead here we need to include the sender
     as we can't just generate a key. */
  recps.push_back (sender);
  encKeys = getEncryptionKeys (recps, GpgME::CMS);
  sigKey = getSigningKey (sender.c_str(), GpgME::CMS);

  TRETURN !encKeys.empty() && !sigKey.isNull();
}

void
KeyCache::update (const char *fpr, GpgME::Protocol proto)
{
  d->update (fpr, proto);
}

GpgME::Key
KeyCache::getByFpr (const char *fpr, bool block) const
{
  return d->getByFpr (fpr, block);
}

void
KeyCache::onUpdateJobDone (const char *fpr, const GpgME::Key &key)
{
  return d->onUpdateJobDone (fpr, key);
}

void
KeyCache::importFromAddrBook (const std::string &mbox, const char *key_data,
                              Mail *mail, GpgME::Protocol proto) const
{
  return d->importFromAddrBook (mbox, key_data, mail, proto);
}

void
KeyCache::onAddrBookImportJobDone (const std::string &mbox,
                                   const std::vector<std::string> &result_fprs,
                                   GpgME::Protocol proto)
{
  return d->onAddrBookImportJobDone (mbox, result_fprs, proto);
}

std::vector<GpgME::Key>
KeyCache::getOverrides (const std::string &mbox, GpgME::Protocol proto)
{
  return d->getOverrides (mbox.c_str (), proto);
}

void
KeyCache::populate ()
{
  return d->populate ();
}

std::vector<GpgME::Key>
KeyCache::getUltimateKeys ()
{
  gpgrt_lock_lock (&fpr_map_lock);
  const auto ret = d->m_ultimate_keys;
  gpgrt_lock_unlock (&fpr_map_lock);
  return ret;
}

/* static */
bool
KeyCache::import_pgp_key_data (const GpgME::Data &data)
{
  TSTART;
  if (data.isNull())
    {
      STRANGEPOINT;
      TRETURN false;
    }
  auto ctx = GpgME::Context::create(GpgME::OpenPGP);

  if (!ctx)
    {
      STRANGEPOINT;
      TRETURN false;
    }

  const auto type = data.type();

  if (type != GpgME::Data::PGPKey)
    {
      log_debug ("%s:%s: Data does not look like PGP Keys",
                 SRCNAME, __func__);
      TRETURN false;
    }
  const auto keys = data.toKeys();

  if (keys.empty())
    {
      log_debug ("%s:%s: Data does not contain PGP Keys",
                 SRCNAME, __func__);
      TRETURN false;
    }

  if (opt.enable_debug & DBG_DATA)
    {
      std::stringstream ss;
      for (const auto &key: keys)
        {
          ss << key << '\n';
        }
      log_debug ("Importing keys: %s", ss.str().c_str());
    }
  const auto result = ctx->importKeys(data);

  if ((opt.enable_debug & DBG_DATA))
    {
      std::stringstream ss;
      ss << result;
      log_debug ("%s:%s: Import result: %s details:\n %s",
                 SRCNAME, __func__, result.error ().asString (),
                 ss.str().c_str());
      if (result.error())
        {
          GpgME::Data out;
          if (ctx->getAuditLog(out, GpgME::Context::DiagnosticAuditLog))
            {
              log_error ("%s:%s: Failed to get diagnostics",
                         SRCNAME, __func__);
            }
          else
            {
              log_debug ("%s:%s: Diagnostics: \n%s\n",
                         SRCNAME, __func__, out.toString().c_str());
            }
        }
    }
  else
    {
      log_debug ("%s:%s: Import result: %s",
                 SRCNAME, __func__, result.error ().asString ());
    }
  TRETURN !result.error();
}
