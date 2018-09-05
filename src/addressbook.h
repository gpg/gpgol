#ifndef SRC_ADDRESSBOOK_H
#define SRC_ADDRESSBOOK_H
/* addressbook.h - Functions for the Addressbook
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

#include "common.h"
#include <map>
#include <string>
#include <vector>
#include "oomhelp.h"

class Mail;

namespace Addressbook
{
typedef struct
{
  shared_disp_t contact;
  const char *data;
} callback_args_t;

/* Configure the OpenPGP Key for this contact. */
void edit_key_o (LPDISPATCH contact);

/* Check the address book for keys to import. */
void check_o (Mail *mail);

/* Update the key information for a contact. */
void update_key_o (void *callback_args);
} // namespace Addressbook

#endif // SRC_ADDRESSBOOK_H
