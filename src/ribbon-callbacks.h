/* ribbon-callbacks.h - Callbacks for the ribbon extension interface
 *    Copyright (C) 2013 Intevation GmbH
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

#ifndef RIBBON_CALLBACKS_H
#define RIBBON_CALLBACKS_H

#include "gpgoladdin.h"

/* For the Icon IDS */
#include "dialogs.h"

/* Id's of our callbacks */
#define ID_CMD_DECRYPT           2
#define ID_CMD_ENCRYPT_SELECTION 3
#define ID_CMD_DECRYPT_SELECTION 4
#define ID_CMD_CERT_MANAGER      5
#define ID_BTN_CERTMANAGER       IDI_KEY_MANAGER_64_PNG
#define ID_BTN_DECRYPT           IDI_DECRYPT_16_PNG
#define ID_BTN_DECRYPT_LARGE     IDI_DECRYPT_48_PNG
#define ID_BTN_ENCRYPT           IDI_ENCRYPT_16_PNG
#define ID_BTN_ENCRYPT_LARGE     IDI_ENCRYPT_48_PNG

HRESULT decryptAttachments (LPDISPATCH ctrl);
HRESULT encryptSelection (LPDISPATCH ctrl);
HRESULT decryptSelection (LPDISPATCH ctrl);
HRESULT getIcon (int id, VARIANT* result);
HRESULT startCertManager (LPDISPATCH ctrl);
#endif
