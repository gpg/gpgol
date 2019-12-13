/* ribbon-callbacks.h - Callbacks for the ribbon extension interface
 * Copyright (C) 2013 Intevation GmbH
 * Software engineering by Intevation GmbH
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
#define ID_CMD_MIME_SIGN        13
#define ID_CMD_MIME_ENCRYPT     14
#define ID_GET_SIGN_PRESSED     15
#define ID_GET_ENCRYPT_PRESSED  16
#define ID_ON_LOAD              17
#define ID_CMD_OPEN_OPTIONS     18
#define ID_GET_IS_DETAILS_ENABLED 19
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
#define ID_GET_IS_CRYPTO_MAIL 35
#define ID_CMD_OPEN_CONTACT_KEY 36
#define ID_CMD_FILE_CLOSE 37
#define ID_CMD_DECRYPT_PERMANENTLY 38
#define ID_CMD_FILE_SAVE_AS 39
#define ID_CMD_FILE_SAVE_AS_IN_WINDOW 40

#define ID_BTN_DECRYPT           IDI_DECRYPT_16_PNG
#define ID_BTN_DECRYPT_LARGE     IDI_DECRYPT_48_PNG
#define ID_BTN_ENCRYPT           IDI_ENCRYPT_16_PNG
#define ID_BTN_ENCRYPT_LARGE     IDI_ENCRYPT_48_PNG
#define ID_BTN_ENCSIGN_LARGE     IDI_ENCSIGN_FILE_48_PNG
#define ID_BTN_SIGN_LARGE        IDI_SIGN_48_PNG
#define ID_BTN_VERIFY_LARGE      IDI_VERIFY_48_PNG

#define OP_ENCRYPT     1 /* Encrypt the data */
#define OP_SIGN        2 /* Sign the data */

HRESULT getIcon (int id, VARIANT* result);

/* Get the toggle state of a crypt button. Flag value 1: encrypt, 2: sign */
HRESULT get_crypt_pressed (LPDISPATCH ctrl, int flags, VARIANT *result, bool is_explorer);
/* Mark the mail to be mime encrypted on send. Flags as above */
HRESULT mark_mime_action (LPDISPATCH ctrl, int flags, bool is_explorer);
/* Check the if the gpgol button should be enabled */
HRESULT get_is_details_enabled (LPDISPATCH ctrl, VARIANT *result);
/* Get the label for the signature. Returns BSTR */
HRESULT get_sig_label (LPDISPATCH ctrl, VARIANT *result);
/* Get the tooltip for the signature. Returns BSTR */
HRESULT get_sig_ttip (LPDISPATCH ctrl, VARIANT *result);
/* Get the supertip for the signature. Returns BSTR */
HRESULT get_sig_stip (LPDISPATCH ctrl, VARIANT *result);
/* Show a certificate details dialog. Returns nothing. */
HRESULT launch_cert_details (LPDISPATCH ctrl);
/* Callback to get the sigstate icon. */
HRESULT get_crypto_icon (LPDISPATCH ctrl, VARIANT *result);
/* Callback to get our own control reference */
HRESULT ribbon_loaded (LPDISPATCH ctrl);
/* Is the currently selected mail a crypto mail ? */
HRESULT get_is_crypto_mail (LPDISPATCH ctrl, VARIANT *result);
/* Open key configuration for a contact */
HRESULT open_contact_key (LPDISPATCH ctrl);
/* An explorer is closed by File->Close */
HRESULT override_file_close ();
/* Decrypt permanently */
HRESULT decrypt_permanently (LPDISPATCH ctrl);
/* SaveAs from the file menu */
HRESULT override_file_save_as (DISPPARAMS *parms);
/* Like above but for mails opened in their own window as they
   behave differently. */
HRESULT override_file_save_as_in_window (DISPPARAMS *parms);
#endif
