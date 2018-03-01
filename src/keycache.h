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
       list is returned. */
    std::vector<GpgME::Key> getEncryptionKeys (const char **recipients,
                                               GpgME::Protocol proto) const;

    /* Start a key location in a background thread filling
       the key cache. cArray is a null terminated array
       of address strings. */
    void startLocate (char **cArray) const;

    /* Look for a secret key for the addr. */
    void startLocateSecret (const char *addr) const;

    /* Start a key location in a background thread filling
       the key cache. */
    void startLocate (const char *addr) const;

    // Internal for thread
    void setSmimeKey(const char *mbox, const GpgME::Key &key);
    void setPgpKey(const char *mbox, const GpgME::Key &key);
    void setSmimeKeySecret(const char *mbox, const GpgME::Key &key);
    void setPgpKeySecret(const char *mbox, const GpgME::Key &key);

private:

    class Private;
    std::shared_ptr<Private> d;
};

#endif
