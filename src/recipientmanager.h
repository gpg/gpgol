/* @file recipientmanager.h
 * @brief Manage the recipients of a mail to send multiple mails to
 *        different recipients.
 *
 * Copyright (C) 2021 g10 Code GmbH
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
#ifndef RECIPIENTMANAGER_H
#define RECIPIENTMANAGER_H

#include "recipient.h"
#include <vector>

#include <gpgme++/key.h>

typedef std::vector<Recipient> RecpList;

class RecipientManager
{
public:
  /* Build a recipient manager with resolved recipients.

     The recipient manager calculates how many mails need to
     be sent to fulfil all requirements regarding split options.
  */
  RecipientManager (const RecpList &recipients,
                    const std::vector<GpgME::Key> &signing_keys);

  /* Returns the number of mails required to send */
  int getRequiredMails () const;

  /* Returns the recipients and the signing key for mail X */
  RecpList getRecipients (int x, GpgME::Key &signing_key) const;

private:
  std::vector<RecpList> m_recp_lists;
  GpgME::Key m_pgpSigKey,
             m_cmsSigKey;
};

#endif // RECIPIENTMANAGER_H
