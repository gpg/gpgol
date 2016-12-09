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
#define ID_CMD_ENCRYPT_BODY      6
#define ID_CMD_DECRYPT_BODY      8
#define ID_CMD_ATT_ENCSIGN_FILE  9
#define ID_CMD_SIGN_BODY        10
#define ID_CMD_ATT_ENC_FILE     11
#define ID_CMD_VERIFY_BODY      12
#define ID_CMD_MIME_SIGN        13
#define ID_CMD_MIME_ENCRYPT     14
#define ID_GET_SIGN_PRESSED     15
#define ID_GET_ENCRYPT_PRESSED  16
#define ID_ON_LOAD              17
#define ID_CMD_OPEN_OPTIONS     18
#define ID_GET_IS_SIGNED        19
#define ID_CMD_MIME_SIGN_EX     21
#define ID_CMD_MIME_ENCRYPT_EX  22
#define ID_GET_SIGN_PRESSED_EX  23
#define ID_GET_ENCRYPT_PRESSED_EX 24
#define ID_GET_SIG_STIP         25
#define ID_GET_SIG_TTIP         26
#define ID_GET_SIG_LABEL        27
#define ID_LAUNCH_CERT_DETAILS  28
#define ID_BTN_SIGSTATE_LARGE   29
#define ID_GET_SIGN_ENCRYPT_PRESSED   30
#define ID_GET_SIGN_ENCRYPT_PRESSED_EX 31
#define ID_CMD_SIGN_ENCRYPT_MIME 32
#define ID_CMD_SIGN_ENCRYPT_MIME_EX 33

#define ID_BTN_CERTMANAGER       IDI_KEY_MANAGER_64_PNG
#define ID_BTN_DECRYPT           IDI_DECRYPT_16_PNG
#define ID_BTN_DECRYPT_LARGE     IDI_DECRYPT_48_PNG
#define ID_BTN_ENCRYPT           IDI_ENCRYPT_16_PNG
#define ID_BTN_ENCRYPT_LARGE     IDI_ENCRYPT_48_PNG
#define ID_BTN_ENCSIGN_LARGE     IDI_ENCSIGN_FILE_48_PNG
#define ID_BTN_SIGN_LARGE        IDI_SIGN_48_PNG
#define ID_BTN_VERIFY_LARGE      IDI_VERIFY_48_PNG

#define OP_ENCRYPT     1 /* Encrypt the data */
#define OP_SIGN        2 /* Sign the data */

HRESULT decryptAttachments (LPDISPATCH ctrl);
HRESULT encryptSelection (LPDISPATCH ctrl);
HRESULT decryptSelection (LPDISPATCH ctrl);
HRESULT decryptBody (LPDISPATCH ctrl);
HRESULT encryptBody (LPDISPATCH ctrl);
HRESULT addEncSignedAttachment (LPDISPATCH ctrl);
HRESULT addEncAttachment (LPDISPATCH ctrl);
HRESULT getIcon (int id, VARIANT* result);
HRESULT startCertManager (LPDISPATCH ctrl);
HRESULT signBody (LPDISPATCH ctrl);
HRESULT verifyBody (LPDISPATCH ctrl);

/* Get the toggle state of a crypt button. Flag value 1: encrypt, 2: sign */
HRESULT get_crypt_pressed (LPDISPATCH ctrl, int flags, VARIANT *result, bool is_explorer);
/* Mark the mail to be mime encrypted on send. Flags as above */
HRESULT mark_mime_action (LPDISPATCH ctrl, int flags, bool is_explorer);
/* Check the if the mail was signed. Returns BOOL */
HRESULT get_is_signed (LPDISPATCH ctrl, VARIANT *result);
/* Get the label for the signature. Returns BSTR */
HRESULT get_sig_label (LPDISPATCH ctrl, VARIANT *result);
/* Get the tooltip for the signature. Returns BSTR */
HRESULT get_sig_ttip (LPDISPATCH ctrl, VARIANT *result);
/* Get the supertip for the signature. Returns BSTR */
HRESULT get_sig_stip (LPDISPATCH ctrl, VARIANT *result);
/* Show a certificate details dialog. Returns nothing. */
HRESULT launch_cert_details (LPDISPATCH ctrl);
/* Callback to get the sigstate icon. */
HRESULT get_sigstate_icon (LPDISPATCH ctrl, VARIANT *result);
/* Callback to get our own control reference */
HRESULT ribbon_loaded (LPDISPATCH ctrl);
#endif
