/* mapihelp.h - Helper functions for MAPI
 *	Copyright (C) 2005, 2007, 2008 g10 Code GmbH
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

#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* The list of message types we support in GpgOL.  */
typedef enum 
  {
    MSGTYPE_UNKNOWN = 0,
    MSGTYPE_SMIME,         /* Original SMIME class. */
    MSGTYPE_GPGOL,
    MSGTYPE_GPGOL_MULTIPART_SIGNED,
    MSGTYPE_GPGOL_MULTIPART_ENCRYPTED,
    MSGTYPE_GPGOL_OPAQUE_SIGNED,
    MSGTYPE_GPGOL_OPAQUE_ENCRYPTED,
    MSGTYPE_GPGOL_CLEAR_SIGNED,
    MSGTYPE_GPGOL_PGP_MESSAGE
  }
msgtype_t;

typedef enum
  {
    ATTACHTYPE_UNKNOWN = 0,
    ATTACHTYPE_MOSS = 1,         /* The original MOSS message (ie. a
                                    S/MIME or PGP/MIME message. */
    ATTACHTYPE_FROMMOSS = 2,     /* Attachment created from MOSS.  */
    ATTACHTYPE_MOSSTEMPL = 3,    /* Attachment has been created in the
                                    course of sending a message */ 
    ATTACHTYPE_PGPBODY = 4       /* Attachment contains the original
                                    PGP message body of PGP inline
                                    encrypted messages.  */
  }
attachtype_t;

/* An object to collect information about one MAPI attachment.  */
struct mapi_attach_item_s
{
  int end_of_table;     /* True if this is the last plus one entry of
                           the table. */
  void *private_mapitable; /* Only for use by mapi_release_attach_table. */

  int mapipos;          /* The position which needs to be passed to
                           MAPI to open the attachment.  -1 means that
                           there is no valid attachment.  */
   
  int method;           /* MAPI attachment method. */
  char *filename;       /* Malloced filename of this attachment or NULL. */

  /* Malloced string with the MIME attrib or NULL.  Parameters are
     stripped off thus a compare against "type/subtype" is
     sufficient. */
  char *content_type; 

  /* If not NULL the parameters of the content_type. */
  const char *content_type_parms; 

  /* The attachment type from Property GpgOL Attach Type.  */
  attachtype_t attach_type;
};
typedef struct mapi_attach_item_s mapi_attach_item_t;

/* The filename of the attachment we create as the result of sign or
   encrypt operations.  Don't change this name as some tests rely on
   it.  */
#define MIMEATTACHFILENAME "gpgolXXX.dat"
/* The name of the file we use to store the original body of PGP
   encrypted messages.  Note that PGP/MIME message don't need that
   because Outlook carries them as 2 attachments.  */
#define PGPBODYFILENAME    "gpgolPGP.dat"

void log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname);
int get_gpgololdmsgclass_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolsigstatus_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgollastdecrypted_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolmimeinfo_tag (LPMESSAGE message, ULONG *r_tag);

int mapi_do_save_changes (LPMESSAGE message, ULONG flags, int only_del_body,
                          const char *dbg_file, const char *dbg_func);
#define mapi_save_changes(a,b) \
        mapi_do_save_changes ((a),(b), 0, __FILE__, __func__)
#define mapi_delete_body_props(a,b) \
        mapi_do_save_changes ((a),(b), 1, __FILE__, __func__)


int mapi_set_header (LPMESSAGE msg, const char *name, const char *val);

int mapi_change_message_class (LPMESSAGE message, int sync_override);
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
char *mapi_get_attach (LPMESSAGE message, int unprotect,
                       mapi_attach_item_t *item, size_t *r_nbytes);
int mapi_mark_moss_attach (LPMESSAGE message, mapi_attach_item_t *item);
int mapi_has_sig_status (LPMESSAGE msg);
int mapi_test_sig_status (LPMESSAGE msg);
char *mapi_get_sig_status (LPMESSAGE msg);

int mapi_set_sig_status (LPMESSAGE message, const char *status_string);

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
int   mapi_test_last_decrypted (LPMESSAGE message);
int   mapi_get_gpgol_body_attachment (LPMESSAGE message,
                                      char **r_body, size_t *r_nbytes,
                                      int *r_ishtml, int *r_protected);

int   mapi_delete_gpgol_body_attachment (LPMESSAGE message);

int   mapi_attachment_to_body (LPMESSAGE message, mapi_attach_item_t *item);

#ifdef __cplusplus
}
#endif
#endif /*MAPIHELP_H*/
