#ifndef KEYCACHE_H
#define KEYCACHE_H

/* @file keycache.h
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

#include "config.h"

#include <memory>
#include <vector>

#include <gpgme++/global.h>

namespace GpgME
{
  class Key;
};

class Mail;

class KeyCache
{
protected:
    /** Internal ctor */
    explicit KeyCache ();

public:
    /** Get the KeyCache */
    static KeyCache* instance ();

    /* Try to find a key for signing in the internal secret key
       cache. If no proper key is found a Null key is
       returned.*/
    GpgME::Key getSigningKey (const char *addr, GpgME::Protocol proto) const;

    /* Get the keys for recipents. The keys
       are taken from the internal cache. If
       one recipient can't be resolved an empty
       list is returned.
       */
    std::vector<GpgME::Key> getEncryptionKeys (const std::vector<std::string> &recipients,
                                               GpgME::Protocol proto) const;

    /* Start a key location in a background thread filling
       the key cache.

       The mail argument is used to add / remove the
       locator thread counter.

       async
       */
    void startLocate (const std::vector<std::string> &addrs, Mail *mail) const;

    /* Look for a secret key for the addr.

       async
    */
    void startLocateSecret (const char *addr, Mail *mail) const;

    /* Start a key location in a background thread filling
       the key cache.

       async
       */
    void startLocate (const char *addr, Mail *mail) const;

    /* Check that a mail is resolvable through the keycache.
     *
     * For OpenPGP only the recipients are checked as we can
     * generate a new key for the sender.
     **/
    bool isMailResolvable (Mail *mail);

    /* Search / Update a key in the cache. This is meant to be
       called e.g. after a verify to update the key.

       async

       A known issue is that a get right after it might
       still return an outdated key but the get after that
       would return the updated one. This is acceptable as
       it only poses a minor problem with TOFU while we
       can show the correct state in the tooltip. */
    void update (const char *fpr, GpgME::Protocol proto);

    /* Get a cached key. If block is true it will block
       if the key is currently searched for.

       This function will not search a key. Call update
       to insert keys into the cache */
    GpgME::Key getByFpr (const char *fpr, bool block = true) const;

    /* Import key data from the address book for the address mbox.
       Keys imported this way take precedence over other keys for
       this mail address regardless of validity.

       The mail argument is used to add / remove the
       locator thread counter.

       async
    */
    void importFromAddrBook (const std::string &mbox,
                             const char *key_data,
                             Mail *mail) const;

    /* Get optional overrides for an address. */
    std::vector<GpgME::Key> getOverrides (const std::string &mbox);

    /* Populate the fingerprint and secret key maps */
    void populate ();

    /* Get a vector of ultimately trusted keys. */
    std::vector<GpgME::Key> getUltimateKeys ();

    // Internal for thread
    void setSmimeKey(const std::string &mbox, const GpgME::Key &key);
    void setPgpKey(const std::string &mbox, const GpgME::Key &key);
    void setSmimeKeySecret(const std::string &mbox, const GpgME::Key &key);
    void setPgpKeySecret(const std::string &mbox, const GpgME::Key &key);
    void onUpdateJobDone(const char *fpr, const GpgME::Key &key);
    void onAddrBookImportJobDone (const std::string &fpr,
                                  const std::vector<std::string> &result_fprs);

private:

    class Private;
    std::shared_ptr<Private> d;
};

#endif
