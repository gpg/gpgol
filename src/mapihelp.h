/* mapihelp.h - Helper functions for MAPI
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
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

#ifndef MAPIHELP_H
#define MAPIHELP_H

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <oomhelp.h>

#ifdef __cplusplus
extern "C" {
#endif
/* The filename of the attachment we create as the result of sign or
   encrypt operations. */
#define MIMEATTACHFILENAME "GpgOL_MIME_structure.txt"

/* The name of the file we use to store the original body of PGP
   encrypted messages.  Note that PGP/MIME message don't need that
   because Outlook carries them as 2 attachments.  */
#define PGPBODYFILENAME    "GpgOL_original_OpenPGP_message.txt"

void log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname);
int get_gpgololdmsgclass_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgollastdecrypted_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolmimeinfo_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolmsgclass_tag (LPMESSAGE message, ULONG *r_tag);

int mapi_do_save_changes (LPMESSAGE message, ULONG flags, int only_del_body,
                          const char *dbg_file, const char *dbg_func);
#define mapi_save_changes(a,b) \
        mapi_do_save_changes ((a),(b), 0, __FILE__, __func__)
#define mapi_delete_body_props(a,b) \
        mapi_do_save_changes ((a),(b), 1, __FILE__, __func__)


int mapi_set_header (LPMESSAGE msg, const char *name, const char *val);

int mapi_change_message_class (LPMESSAGE message, int sync_override,
                               msgtype_t *r_type);
char *mapi_get_message_class (LPMESSAGE message);
char *mapi_get_old_message_class (LPMESSAGE message);
char *mapi_get_sender (LPMESSAGE message);
msgtype_t mapi_get_message_type (LPMESSAGE message);
int mapi_to_mime (LPMESSAGE message, const char *filename);

char *mapi_get_binary_prop (LPMESSAGE message,ULONG proptype,size_t *r_nbytes);
int  mapi_get_int_prop (LPMAPIPROP object, ULONG proptype, LONG *r_value);

char *mapi_get_from_address (LPMESSAGE message);
char *mapi_get_subject (LPMESSAGE message);

LPSTREAM mapi_get_body_as_stream (LPMESSAGE message);
char *mapi_get_body (LPMESSAGE message, size_t *r_nbytes);

mapi_attach_item_t *mapi_create_attach_table (LPMESSAGE message, int fast);
void mapi_release_attach_table (mapi_attach_item_t *table);
LPSTREAM mapi_get_attach_as_stream (LPMESSAGE message, 
                                    mapi_attach_item_t *item, 
                                    LPATTACH *r_attach);
char *mapi_get_attach (LPMESSAGE message,
                       mapi_attach_item_t *item, size_t *r_nbytes);
int mapi_mark_moss_attach (LPMESSAGE message, mapi_attach_item_t *item);

int mapi_set_gpgol_msg_class (LPMESSAGE message, const char *name);

char *mapi_get_gpgol_charset (LPMESSAGE obj);
int mapi_set_gpgol_charset (LPMESSAGE obj, const char *charset);

char *mapi_get_gpgol_draft_info (LPMESSAGE msg);
int   mapi_set_gpgol_draft_info (LPMESSAGE message, const char *string);


int  mapi_set_attach_hidden (LPATTACH attach);
int  mapi_test_attach_hidden (LPATTACH attach);

char *mapi_get_mime_info (LPMESSAGE msg);

char *mapi_get_message_content_type (LPMESSAGE message, 
                                     char **r_protocol, char **r_smtype);

int   mapi_has_last_decrypted (LPMESSAGE message);

attachtype_t get_gpgolattachtype (LPATTACH obj, ULONG tag);

int get_gpgol_draft_info_flags (LPMESSAGE message);

int set_gpgol_draft_info_flags (LPMESSAGE message, int flags);

/* Mark crypto attachments as hidden. And mark the moss
 attachment for later use. Returns the attachments position
 (1 is the first attachment) or 0 in case no attachment was found. */
int mapi_mark_or_create_moss_attach (LPMESSAGE base_message,
                                     LPMESSAGE parsed_message,
                                     msgtype_t msgtype);

/* Copy the MAPI body to a PGPBODY type attachment. */
int mapi_body_to_attachment (LPMESSAGE message);

/* Get malloced uid of a message */
char * mapi_get_uid (LPMESSAGE message);

/* Remove the gpgol specific mapi tags */
void mapi_delete_gpgol_tags (LPMESSAGE message);

/* Explicitly change the message class */
void mapi_set_mesage_class (LPMESSAGE message, const char *cls);

#ifdef __cplusplus
}
#include <string>

/* Parse the headers for additional useful fields.

  r_autocrypt_info: Information about the autocrypt header.
  r_header_info: Just a generic struct for other stuff.

  A return value of false indicates an error. Check the
  individual info fields "exists" values to check if
  the header exists in the message */
bool mapi_get_header_info (LPMESSAGE message,
                           header_info_s &r_header_info);
std::string mapi_get_header (LPMESSAGE message);
#endif
#endif /*MAPIHELP_H*/
