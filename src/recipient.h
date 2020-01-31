/* @file recipient.h
 * @brief Information about a recipient.
 *
 * Copyright (C) 2020, g10 code GmbH
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
#ifndef RECIPIENT_H
#define RECIPIENT_H

#include <string>
#include <vector>

namespace GpgME
{
  class Key;
} // namespace GpgME

class Recipient
{
public:
    Recipient ();
    Recipient (const Recipient &other);
    explicit Recipient (const char *addr, int type);

    enum recipientType
      {
        olOriginator = 0, /* Originator (sender) of the Item */
        olCC = 2, /* Specified in the CC property */
        olTo = 1, /* Specified in the To property */
        olBCC = 3, /* BCC */
        invalidType = -1, /* indicates that the type was not set or the
                             recipient is somehow invalid */
      };

    void setKeys (const std::vector <GpgME::Key> &keys);
    std::vector<GpgME::Key> keys () const;

    std::string mbox () const;
    recipientType type () const;
    void setType (int type);
    void setIndex (int index);
    int index() const;

    /* For debugging */
    static void dump(const std::vector<Recipient> &recps);
private:
    std::string m_mbox;
    recipientType m_type;
    std::vector<GpgME::Key> m_keys;
    int m_index;
};
#endif
