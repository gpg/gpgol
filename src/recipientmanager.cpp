/* @file recipientmanager.cpp
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

#include "config.h"
#include "recipientmanager.h"
#include "common.h"

static std::vector<GpgME::Key>
getKeysForProtocol (const std::vector<GpgME::Key> keys, GpgME::Protocol proto)
{
  std::vector<GpgME::Key> ret;

  std::copy_if (keys.begin(), keys.end(),
                std::back_inserter (ret),
                [proto] (const auto &k)
                {
                  return k.protocol () == proto;
                });

  return ret;
}

RecipientManager::RecipientManager (const RecpList &recps,
                                    const std::vector<GpgME::Key> &signing_keys) :
  m_prot_split (false)
{
  TSTART;
  log_dbg ("Recipient manager ctor");
  Recipient::dump (recps);

  if (signing_keys.size () > 2)
    {
      log_err ("Signing with multiple keys is not supported.");
    }
  for (const auto &key: signing_keys)
    {
      if (key.protocol () == GpgME::OpenPGP)
        {
          m_pgpSigKey = key;
        }
      else if (key.protocol () == GpgME::CMS)
        {
          m_cmsSigKey = key;
        }
      else
        {
          STRANGEPOINT;
        }
    }

  if (!opt.splitBCCMails && !opt.combinedOpsEnabled)
    {
      /* Nothing to do */
      m_recp_lists.push_back (recps);
      TRETURN;
    }

  /* First collect all BCC recpients and put all others in a
     list. */
  Recipient originator;
  RecpList normalRecps;
  RecpList bccRecps;
  for (const auto &recp: recps)
    {
      if (recp.type () == Recipient::olOriginator)
        {
          originator = recp;
          normalRecps.push_back (recp);
        }
      else if (recp.type () == Recipient::olBCC && opt.splitBCCMails)
        {
          bccRecps.push_back (recp);
        }
      else
        {
          normalRecps.push_back (recp);
        }
    }

  if (originator.isNull())
    {
      log_err ("Failed to find orginator");
      TRETURN;
    }

  /* Split the normal recipients by protocol */
  m_recp_lists.push_back (normalRecps);

  /* For each BCC recipient also split by protocol */
  log_dbg ("Splitting of %i BCC recipients.", (int) bccRecps.size());

  for (const auto &recp: bccRecps)
    {
      RecpList list;
      list.push_back (originator);
      list.push_back (recp);
      m_recp_lists.push_back (list);
    }
  (void) getKeysForProtocol;
}

int
RecipientManager::getRequiredMails() const
{
  return (int) m_recp_lists.size ();
}

RecpList
RecipientManager::getRecipients (int x, GpgME::Key &signing_key) const
{
  if (x >= m_recp_lists.size () || x < 0)
    {
      log_err ("Invalid index for recpients: %i", x);
      return RecpList ();
    }
  RecpList ret = m_recp_lists[x];
  if (ret.size() && ret[0].keys().size())
    {
      signing_key = (ret[0].keys()[0].protocol () == GpgME::CMS ? m_cmsSigKey : m_pgpSigKey);
    }
  return ret;
}

bool
RecipientManager::isSplitByProtocol () const
{
  return m_prot_split;
}
