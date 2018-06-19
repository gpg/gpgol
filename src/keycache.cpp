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

#include <gpg-error.h>
#include <gpgme++/context.h>
#include <gpgme++/key.h>

#include <windows.h>

#include <map>

GPGRT_LOCK_DEFINE (keycache_lock);
static KeyCache* singleton = nullptr;

class KeyCache::Private
{
public:
  Private()
  {

  }

  void setPgpKey(const std::string &mbox, const GpgME::Key &key)
  {
    gpgrt_lock_lock (&keycache_lock);
    auto it = m_pgp_key_map.find (mbox);

    if (it == m_pgp_key_map.end ())
      {
        m_pgp_key_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    gpgrt_lock_unlock (&keycache_lock);
  }

  void setSmimeKey(const std::string &mbox, const GpgME::Key &key)
  {
    gpgrt_lock_lock (&keycache_lock);
    auto it = m_smime_key_map.find (mbox);

    if (it == m_smime_key_map.end ())
      {
        m_smime_key_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    gpgrt_lock_unlock (&keycache_lock);
  }

  void setPgpKeySecret(const std::string &mbox, const GpgME::Key &key)
  {
    gpgrt_lock_lock (&keycache_lock);
    auto it = m_pgp_skey_map.find (mbox);

    if (it == m_pgp_skey_map.end ())
      {
        m_pgp_skey_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    gpgrt_lock_unlock (&keycache_lock);
  }

  void setSmimeKeySecret(const std::string &mbox, const GpgME::Key &key)
  {
    gpgrt_lock_lock (&keycache_lock);
    auto it = m_smime_skey_map.find (mbox);

    if (it == m_smime_skey_map.end ())
      {
        m_smime_skey_map.insert (std::pair<std::string, GpgME::Key> (mbox, key));
      }
    else
      {
        it->second = key;
      }
    gpgrt_lock_unlock (&keycache_lock);
  }

  GpgME::Key getKey (const char *addr, GpgME::Protocol proto)
  {
    if (!addr)
      {
        return GpgME::Key();
      }
    auto mbox = GpgME::UserID::addrSpecFromString (addr);

    if (proto == GpgME::OpenPGP)
      {
        gpgrt_lock_lock (&keycache_lock);
        const auto it = m_pgp_key_map.find (mbox);

        if (it == m_pgp_key_map.end ())
          {
            gpgrt_lock_unlock (&keycache_lock);
            return GpgME::Key();
          }
        const auto ret = it->second;
        gpgrt_lock_unlock (&keycache_lock);

        return ret;
      }
    gpgrt_lock_lock (&keycache_lock);
    const auto it = m_smime_key_map.find (mbox);

    if (it == m_smime_key_map.end ())
      {
        gpgrt_lock_unlock (&keycache_lock);
        return GpgME::Key();
      }
    const auto ret = it->second;
    gpgrt_lock_unlock (&keycache_lock);

    return ret;
  }

  GpgME::Key getSKey (const char *addr, GpgME::Protocol proto)
  {
    if (!addr)
      {
        return GpgME::Key();
      }
    auto mbox = GpgME::UserID::addrSpecFromString (addr);

    if (proto == GpgME::OpenPGP)
      {
        gpgrt_lock_lock (&keycache_lock);
        const auto it = m_pgp_skey_map.find (mbox);

        if (it == m_pgp_skey_map.end ())
          {
            gpgrt_lock_unlock (&keycache_lock);
            return GpgME::Key();
          }
        const auto ret = it->second;
        gpgrt_lock_unlock (&keycache_lock);

        return ret;
      }
    gpgrt_lock_lock (&keycache_lock);
    const auto it = m_smime_skey_map.find (mbox);

    if (it == m_smime_skey_map.end ())
      {
        gpgrt_lock_unlock (&keycache_lock);
        return GpgME::Key();
      }
    const auto ret = it->second;
    gpgrt_lock_unlock (&keycache_lock);

    return ret;
  }

  GpgME::Key getSigningKey (const char *addr, GpgME::Protocol proto)
  {
    const auto key = getSKey (addr, proto);
    if (key.isNull())
      {
        log_mime_parser ("%s:%s: secret key for %s is null",
                   SRCNAME, __func__, addr);
        return key;
      }
    if (!key.canReallySign())
      {
        log_mime_parser ("%s:%s: Discarding key for %s because it can't sign",
                   SRCNAME, __func__, addr);
        return GpgME::Key();
      }
    if (!key.hasSecret())
      {
        log_mime_parser ("%s:%s: Discarding key for %s because it has no secret",
                   SRCNAME, __func__, addr);
        return GpgME::Key();
      }
    if (in_de_vs_mode () && !key.isDeVs())
      {
        log_mime_parser ("%s:%s: signing key for %s is not deVS",
                   SRCNAME, __func__, addr);
        return GpgME::Key();
      }
    return key;
  }

  std::vector<GpgME::Key> getEncryptionKeys (const std::vector<std::string> &recipients,
                                             GpgME::Protocol proto)
  {
    std::vector<GpgME::Key> ret;
    if (recipients.empty ())
      {
        TRACEPOINT;
        return ret;
      }
    for (const auto &recip: recipients)
      {
        const auto key = getKey (recip.c_str (), proto);
        if (key.isNull())
          {
            log_mime_parser ("%s:%s: No key for %s. no internal encryption",
                       SRCNAME, __func__, recip.c_str ());
            return std::vector<GpgME::Key>();
          }

        if (!key.canEncrypt() || key.isRevoked() ||
            key.isExpired() || key.isDisabled() || key.isInvalid())
          {
            log_mime_parser ("%s:%s: Invalid key for %s. no internal encryption",
                       SRCNAME, __func__, recip.c_str ());
            return std::vector<GpgME::Key>();
          }

        if (in_de_vs_mode () && !key.isDeVs ())
          {
            log_mime_parser ("%s:%s: key for %s is not deVS",
                       SRCNAME, __func__, recip.c_str ());
            return std::vector<GpgME::Key>();
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
          }
        if (!validEnough)
          {
            log_mime_parser ("%s:%s: UID for %s does not have at least marginal trust",
                             SRCNAME, __func__, recip.c_str ());
            return std::vector<GpgME::Key>();
          }
        // Accepting key
        ret.push_back (key);
      }
    return ret;
  }

  std::map<std::string, GpgME::Key> m_pgp_key_map;
  std::map<std::string, GpgME::Key> m_smime_key_map;
  std::map<std::string, GpgME::Key> m_pgp_skey_map;
  std::map<std::string, GpgME::Key> m_smime_skey_map;
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

static DWORD WINAPI
do_locate (LPVOID arg)
{
  if (!arg)
    {
      return 0;
    }
  std::string addr = (char*) arg;
  xfree (arg);

  log_mime_parser ("%s:%s searching key for addr: \"%s\"",
                   SRCNAME, __func__, addr.c_str());

  const auto k = GpgME::Key::locate (addr.c_str());

  if (!k.isNull ())
    {
      log_mime_parser ("%s:%s found key for addr: \"%s\":%s",
                       SRCNAME, __func__, addr.c_str(),
                       k.primaryFingerprint());
      KeyCache::instance ()->setPgpKey (addr, k);
    }

  if (opt.enable_smime)
    {
      auto ctx = std::unique_ptr<GpgME::Context> (
                          GpgME::Context::createForProtocol (GpgME::CMS));
      if (!ctx)
        {
          TRACEPOINT;
          return 0;
        }
      // We need to validate here to fetch CRL's
      ctx->setKeyListMode (GpgME::KeyListMode::Local |
                           GpgME::KeyListMode::Validate);
      GpgME::Error e = ctx->startKeyListing (addr.c_str());
      if (e)
        {
          TRACEPOINT;
          return 0;
        }

      std::vector<GpgME::Key> keys;
      GpgME::Error err;
      do {
          keys.push_back(ctx->nextKey(err));
      } while (!err);
      keys.pop_back();

      GpgME::Key candidate;
      for (const auto &key: keys)
        {
          if (key.isRevoked() || key.isExpired() ||
              key.isDisabled() || key.isInvalid())
            {
              log_mime_parser ("%s:%s: Skipping invalid S/MIME key",
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
      if (!candidate.isNull())
        {
          log_mime_parser ("%s:%s found SMIME key for addr: \"%s\":%s",
                           SRCNAME, __func__, addr.c_str(),
                           candidate.primaryFingerprint());
          KeyCache::instance()->setSmimeKey (addr, candidate);
        }
    }
  log_debug ("%s:%s locator thread done",
             SRCNAME, __func__);
  return 0;
}

static void
locate_secret (char *addr, GpgME::Protocol proto)
{
  auto ctx = std::unique_ptr<GpgME::Context> (
                      GpgME::Context::createForProtocol (proto));
  if (!ctx)
    {
      TRACEPOINT;
      return;
    }
  if (!addr)
    {
      TRACEPOINT;
      return;
    }
  const auto mbox = GpgME::UserID::addrSpecFromString (addr);

  // We need to validate here to fetch CRL's
  ctx->setKeyListMode (GpgME::KeyListMode::Local |
                       GpgME::KeyListMode::Validate);
  GpgME::Error e = ctx->startKeyListing (mbox.c_str(), true);
  if (e)
    {
      TRACEPOINT;
      return;
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
          log_mime_parser ("%s:%s: Skipping invalid secret key",
                           SRCNAME, __func__);
          continue;
        }
      if (proto == GpgME::OpenPGP)
        {
          log_mime_parser ("%s:%s found pgp skey for addr: \"%s\":%s",
                           SRCNAME, __func__, mbox.c_str(),
                           key.primaryFingerprint());
          KeyCache::instance()->setPgpKeySecret (mbox, key);
          return;
        }
      if (proto == GpgME::CMS)
        {
          log_mime_parser ("%s:%s found cms skey for addr: \"%s\":%s",
                           SRCNAME, __func__, mbox.c_str (),
                           key.primaryFingerprint());
          KeyCache::instance()->setSmimeKeySecret (mbox, key);
          return;
        }
    } while (!err);
  return;
}

static DWORD WINAPI
do_locate_secret (LPVOID arg)
{
  char *addr = (char*) arg;

  log_mime_parser ("%s:%s searching secret key for addr: \"%s\"",
                   SRCNAME, __func__, addr);

  locate_secret (addr, GpgME::OpenPGP);
  if (opt.enable_smime)
    {
      locate_secret (addr, GpgME::CMS);
    }
  xfree (addr);
  log_debug ("%s:%s locator sthread thread done",
             SRCNAME, __func__);
  return 0;
}

void
KeyCache::startLocate (char **recipients) const
{
  if (!recipients)
    {
      TRACEPOINT;
      return;
    }
  for (int i = 0; recipients[i]; i++)
    {
      startLocate (recipients[i]);
    }
}

void
KeyCache::startLocate (const char *addr) const
{
  if (!addr)
    {
      TRACEPOINT;
      return;
    }
  std::string recp = GpgME::UserID::addrSpecFromString (addr);
  if (recp.empty ())
    {
      return;
    }
  gpgrt_lock_lock (&keycache_lock);
  if (d->m_pgp_key_map.find (recp) == d->m_pgp_key_map.end ())
    {
      // It's enough to look at the PGP Key map. We marked
      // searched keys there.
      d->m_pgp_key_map.insert (std::pair<std::string, GpgME::Key> (recp, GpgME::Key()));
      log_debug ("%s:%s Creating a locator thread",
                 SRCNAME, __func__);
      HANDLE thread = CreateThread (NULL, 0, do_locate,
                                    (LPVOID) strdup (recp.c_str ()), 0,
                                    NULL);
      CloseHandle (thread);
    }
  gpgrt_lock_unlock (&keycache_lock);
}

void
KeyCache::startLocateSecret (const char *addr) const
{
  if (!addr)
    {
      TRACEPOINT;
      return;
    }
  std::string recp = GpgME::UserID::addrSpecFromString (addr);
  if (recp.empty ())
    {
      return;
    }
  gpgrt_lock_lock (&keycache_lock);
  if (d->m_pgp_skey_map.find (recp) == d->m_pgp_skey_map.end ())
    {
      // It's enough to look at the PGP Key map. We marked
      // searched keys there.
      d->m_pgp_skey_map.insert (std::pair<std::string, GpgME::Key> (recp, GpgME::Key()));
      log_debug ("%s:%s Creating a locator thread",
                 SRCNAME, __func__);
      HANDLE thread = CreateThread (NULL, 0, do_locate_secret,
                                    (LPVOID) strdup (recp.c_str ()), 0,
                                    NULL);
      CloseHandle (thread);
    }
  gpgrt_lock_unlock (&keycache_lock);
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
