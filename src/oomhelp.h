/* oomhelp.h - Defs for helper functions for the Outlook Object Model
 * Copyright (C) 2009 g10 Code GmbH
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
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

#ifndef OOMHELP_H
#define OOMHELP_H

#include <unknwn.h>
#include "mymapi.h"
#include "common.h"

#include <vector>
#include <string>
#include <memory>

#define MSOCONTROLBUTTON    1
#define MSOCONTROLEDIT      2
#define MSOCONTROLDROPDOWN  3
#define MSOCONTROLCOMBOBOX  4
#define MSOCONTROLPOPUP    10

enum 
  {
    msoButtonAutomatic = 0,
    msoButtonIcon = 1,
    msoButtonCaption = 2,
    msoButtonIconAndCaption = 3,
    msoButtonIconAndWrapCaption = 7,
    msoButtonIconAndCaptionBelow = 11,
    msoButtonWrapCaption = 14,
    msoButtonIconAndWrapCaptionBelow = 15 
  };

enum
  {
    msoButtonDown = -1,
    msoButtonUp = 0,
    msoButtonMixed = 2
  };


DEFINE_GUID(GUID_NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

DEFINE_GUID(IID_IConnectionPoint, 
            0xb196b286, 0xbab4, 0x101a,
            0xb6, 0x9c, 0x00, 0xaa, 0x00, 0x34, 0x1d, 0x07);
DEFINE_GUID(IID_IConnectionPointContainer, 
            0xb196b284, 0xbab4, 0x101a,
            0xb6, 0x9c, 0x00, 0xaa, 0x00, 0x34, 0x1d, 0x07);
DEFINE_GUID(IID_IPictureDisp,
            0x7bf80981, 0xbf32, 0x101a,
            0x8b, 0xbb, 0x00, 0xaa, 0x00, 0x30, 0x0c, 0xab);
DEFINE_GUID(IID_FolderEvents, 0x000630F7, 0x0000, 0x0000,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_ApplicationEvents, 0x0006304E, 0x0000, 0x0000,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_ApplicationEvents_11, 0x0006302C, 0x0000, 0x0000,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_ExplorerEvents, 0x0006300F, 0x0000, 0x0000,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_ExplorersEvents, 0x00063078, 0x0000, 0x0000,
            0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_MailItemEvents, 0x0006302B, 0x0000, 0x0000,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_MailItem, 0x00063034, 0x0000, 0x0000,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);
DEFINE_GUID(IID_IMAPISecureMessage, 0x253cc320, 0xeab6, 0x11d0,
            0x82, 0x22, 0, 0x60, 0x97, 0x93, 0x87, 0xea);

DEFINE_OLEGUID(IID_IUnknown,                  0x00000000, 0, 0);
DEFINE_OLEGUID(IID_IDispatch,                 0x00020400, 0, 0);
DEFINE_OLEGUID(IID_IOleWindow,                0x00000114, 0, 0);

#ifndef PR_SMTP_ADDRESS_DASL
#define PR_SMTP_ADDRESS_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
#endif

#ifndef PR_EMS_AB_PROXY_ADDRESSES_DASL
#define PR_EMS_AB_PROXY_ADDRESSES_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x800F101E"
#endif

#ifndef PR_ATTACHMENT_HIDDEN_DASL
#define PR_ATTACHMENT_HIDDEN_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
#endif

#ifndef PR_ADDRTYPE_DASL
#define PR_ADDRTYPE_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x3002001E"
#endif

#ifndef PR_EMAIL_ADDRESS_DASL
#define PR_EMAIL_ADDRESS_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x3003001E"
#endif

#define PR_MESSAGE_CLASS_W_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x001A001F"
#define GPGOL_ATTACHTYPE_DASL \
  "http://schemas.microsoft.com/mapi/string/" \
  "{31805AB8-3E92-11DC-879C-00061B031004}/GpgOL Attach Type/0x00000003"
#define GPGOL_UID_DASL \
  "http://schemas.microsoft.com/mapi/string/" \
  "{31805AB8-3E92-11DC-879C-00061B031004}/GpgOL UID/0x0000001F"
#define PR_ATTACH_DATA_BIN_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x37010102"
#define PR_BODY_W_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x1000001F"
#define PR_ATTACHMENT_HIDDEN_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
#define PR_ATTACH_MIME_TAG_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
#define PR_ATTACH_CONTENT_ID_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
#define PR_ATTACH_FLAGS_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x37140003"

#define PR_TAG_SENDER_SMTP_ADDRESS \
  "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
#define PR_TAG_RECEIVED_REPRESENTING_SMTP_ADDRESS \
  "http://schemas.microsoft.com/mapi/proptag/0x5D08001F"
#define PR_PIDNameContentType_DASL \
  "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/content-type/0x0000001F"
#define PR_BLOCK_STATUS_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x10960003"
#define PR_SENT_REPRESENTING_EMAIL_ADDRESS_W_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x0065001F"

#define PR_SENDER_NAME_W_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
#define PR_SENT_REPRESENTING_NAME_W_DASL \
  "http://schemas.microsoft.com/mapi/proptag/0x0042001F"

#define DISTRIBUTION_LIST_ADDRESS_ENTRY_TYPE 11

typedef std::shared_ptr<IDispatch> shared_disp_t;

/* Function to contain the gpgol_release macro */
void release_disp (LPDISPATCH obj);

#define MAKE_SHARED(X) shared_disp_t ((LPDISPATCH)X, &release_disp)

/* Return the malloced name of an COM+ object.  */
char *get_object_name (LPUNKNOWN obj);

/* Helper to lookup a dispid.  */
DISPID lookup_oom_dispid (LPDISPATCH pDisp, const char *name);

/* Return the OOM object's IDispatch interface described by FULLNAME.  */
LPDISPATCH get_oom_object (LPDISPATCH pStart, const char *fullname);

/* Set the Icon of a CommandBarControl.  */
int put_oom_icon (LPDISPATCH pDisp, int rsource_id, int size);

/* Set the boolean property NAME to VALUE.  */
int put_oom_bool (LPDISPATCH pDisp, const char *name, int value);

/* Set the property NAME to VALUE.  */
int put_oom_int (LPDISPATCH pDisp, const char *name, int value);

/* Set the property NAME to STRING.  */
int put_oom_string (LPDISPATCH pDisp, const char *name, const char *string);

/* Set the property NAME to DISP.  */
int put_oom_disp (LPDISPATCH pDisp, const char *name, LPDISPATCH value);

/* Get the boolean property NAME of the object PDISP.  */
int get_oom_bool (LPDISPATCH pDisp, const char *name);

/* Get the integer property NAME of the object PDISP.  */
int get_oom_int (LPDISPATCH pDisp, const char *name);

/* Get the string property NAME of the object PDISP.  */
char *get_oom_string (LPDISPATCH pDisp, const char *name);

/* Get an IUnknown object from property NAME of PDISP.  */
LPUNKNOWN get_oom_iunknown (LPDISPATCH pDisp, const char *name);

/* Return the control object with tag property value TAG.  */
LPDISPATCH get_oom_control_bytag (LPDISPATCH pObj, const char *tag);

/* Add a new button to an object which supports the add method.
   Returns the new object or NULL on error.  */
LPDISPATCH add_oom_button (LPDISPATCH pObj);

/* Delete a button.  */
void del_oom_button (LPDISPATCH button);

/* Get the HWND of the active window in the current context */
HWND get_oom_context_window (LPDISPATCH context);

/* Get the address of the recipients as string list.
   If r_err is not null it is set to true in case of an error. */
std::vector<std::string> get_oom_recipients (LPDISPATCH recipients,
                                             bool *r_err = nullptr);

/* Same as above but include the AddrEntry object in the result.
   Caller needs to release the AddrEntry. */
std::vector<std::pair<std::string, shared_disp_t> >
get_oom_recipients_with_addrEntry (LPDISPATCH recipients,
                                   bool *r_err = nullptr);

/* Add an attachment to a dispatcher */
int
add_oom_attachment (LPDISPATCH disp, const wchar_t* inFile,
                    const wchar_t *displayName, std::string &r_err_str,
                    int *r_err_code);

/* Look up a string with the propertyAccessor interface */
char *
get_pa_string (LPDISPATCH pDisp, const char *property);

/* Look up a long with the propertyAccessor interface.
 returns -1 on error.*/
int
get_pa_int (LPDISPATCH pDisp, const char *property, int *rInt);

/* Set a variant with the propertyAccessor interface.

   This is tested to work at least vor BSTR variants. Trying
   to set PR_ATTACH_DATA_BIN_DASL with this failed with
   hresults 0x80020005 type mismatch or 0x80020008 vad
   variable type for:
   VT_ARRAY | VT_UI1 | VT_BYREF
   VT_SAFEARRAY | VT_UI1 | VT_BYREF
   VT_BSTR | VT_BYREF
   VT_BSTR
   VT_ARRAY | VT_UI1
   VT_SAFEARRAY | VT_UI1

   No idea whats wrong there. Needs more experiments. The
   Type is only documented as "Binary". Outlookspy also
   fails with the same error when trying to modify the
   property.
*/
int
put_pa_string (LPDISPATCH pDisp, const char *dasl_id, const char *value);

int
put_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *value);

int
put_pa_int (LPDISPATCH pDisp, const char *dasl_id, int value);

/* Look up a variant with the propertyAccessor interface */
int
get_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *rVariant);

/* Look up a LONG with the propertyAccessor interface */
LONG
get_pa_long (LPDISPATCH pDisp, const char *dasl_id);

/* Queries the interface of the dispatcher for the id
   id. Returns NULL on error. The returned Object
   must be released.
   Mainly useful to check if an object is what
   it appears to be. */
LPDISPATCH
get_object_by_id (LPDISPATCH pDisp, REFIID id);

/* Obtain the MAPI Message corresponding to the
   Mailitem. Returns NULL on error.

   The returned Message needs to be released by the
   caller */
LPMESSAGE
get_oom_message (LPDISPATCH mailitem);

/* Obtain the Base MAPI Message of a MailItem.
   The parameter should be a pointer to a MailItem.
   returns NULL on error.

   The returned Message needs to be released by the
   caller.
*/
LPMESSAGE
get_oom_base_message (LPDISPATCH mailitem);

/* Get a strong reference for a mail object by calling
   Application.GetObjectReference with type strong. The
   documentation is unclear what this acutally does.
   This function is left over from experiments about
   strong references. Maybe there is a use for them.
   The reference we use in the Mail object is documented
   as a Weak reference. But changing that does not appear
   to make a difference.
*/
LPDISPATCH
get_strong_reference (LPDISPATCH mail);

/* Invoke a method of an outlook object.
   returns true on success false otherwise.

   rVariant should either point to a propery initialized
   variant (initinalized wiht VariantInit) to hold
   the return value or a pointer to NULL.
   */
int
invoke_oom_method (LPDISPATCH pDisp, const char *name, VARIANT *rVariant);

/* Invoke a method of an outlook object.
   returns true on success false otherwise.

   rVariant should either point to a propery initialized
   variant (initinalized wiht VariantInit) to hold
   the return value or a pointer to NULL.

   parms can optionally be used to provide a DISPPARAMS structure
   with parameters for the function.
   */
int
invoke_oom_method_with_parms (LPDISPATCH pDisp, const char *name,
                              VARIANT *rVariant, DISPPARAMS *params);

/* Same as invoke oom method but do the string marshalling for arg */
int invoke_oom_method_with_string (LPDISPATCH pDisp, const char *name,
                                   const char *arg,
                                   VARIANT *rVariant = nullptr);

/* Try to obtain the mapisession through the Application.
  returns NULL on error.*/
LPMAPISESSION
get_oom_mapi_session (void);

/* Ensure a category of the name name exists.

  Creates the category with the specified color if required.

  returns 0 on success. */
void
ensure_category_exists (const char *category, int color);

/* Add a category to a mail if it is not already added. */
int
add_category (LPDISPATCH mail, const char *category);

/* Remove a category from a mail if it was added. */
int
remove_category (LPDISPATCH mail, const char *category, bool exactMatch);

/* Create the category */
int
create_category (LPDISPATCH categories, const char *category, int color);

/* Delete a category from the store. */
int delete_category (LPDISPATCH store, const char *category);

/* Delete categories starting with "string" from all stores. */
void delete_all_categories_starting_with (const char *string);

/* Iterate over application stores and return the one with the ID
   @storeID */
LPDISPATCH get_store_for_id (const char *storeID);

/* Get a unique identifier for a mail object. The
   uuid is a custom property. If create is set
   a new uuid will be added if none exists and the
   value of that uuid returned.

   The optinal uuid value can be set to be used
   as uuid instead of a generated one.

   Return value has to be freed by the caller.
   */
char *
get_unique_id (LPDISPATCH mail, int create, const char* uuid);


/* Uses the Application->ActiveWindow to determine the hwnd
   through FindWindow and the caption. Does not use IOleWindow
   because that was unreliable somhow. */
HWND get_active_hwnd (void);

/* Create a new mailitem and return it */
LPDISPATCH create_mail (void);

LPDISPATCH get_account_for_mail (const char *mbox);

/* Print all active addins to log */
void log_addins (void);

/* Sender fallbacks. All return either null or a malloced address. */
char *get_sender_CurrentUser (LPDISPATCH mailitem);
char *get_sender_Sender (LPDISPATCH mailitem);
char *get_sender_SenderEMailAddress (LPDISPATCH mailitem);


/* Get the body of the active inline response */
char *get_inline_body (void);

/* Get the major version of the exchange server of the account for the
   mail address "mbox". Returns -1 if no version could be detected
   or exchange is not used.*/
int get_ex_major_version_for_addr (const char *mbox);

/* Get the language code used for Outlooks UI */
int get_ol_ui_language (void);

char *get_sender_SendUsingAccount (LPDISPATCH mailitem, bool *r_is_GSuite);

/* Get the SentRepresentingAddress */
char *get_sender_SentRepresentingAddress (LPDISPATCH mailitem);

/* memtracing query interface */
HRESULT gpgol_queryInterface (LPUNKNOWN pObj, REFIID riid, LPVOID FAR *ppvObj);

HRESULT gpgol_openProperty (LPMAPIPROP obj, ULONG ulPropTag, LPCIID lpiid,
                            ULONG ulInterfaceOptions, ULONG ulFlags,
                            LPUNKNOWN FAR * lppUnk);

/* Check if the preview pane in the explorer is visible */
bool is_preview_pane_visible (LPDISPATCH explorer);

/* Find or add a text user property with that name. */
LPDISPATCH find_or_add_text_prop (LPDISPATCH props, const char *name);

/* Find a user property and return it if found. */
LPDISPATCH find_user_prop (LPDISPATCH props, const char *name);

/* Return true if this message is in the junk folder for this account */
bool is_junk_mail (LPDISPATCH mailitem);

/* Return true if this message is in the draft folder for this account */
bool is_draft_mail (LPDISPATCH mailitem);

/* Returns info about a dispparms variable for debugging. */
void format_variant (std::istringstream &stream, VARIANT *var);
std::string format_dispparams (DISPPARAMS *p);

/* Returns the count of attachments that are not hidden. */
int count_visible_attachments (LPDISPATCH attachments);
#endif /*OOMHELP_H*/
