/* mapihelp.h - Helper functions for MAPI
 *	Copyright (C) 2005, 2007 g10 Code GmbH
 *
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
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
    ATTACHTYPE_MOSSTEMPL = 3     /* Attachment has been created in the
                                    course of sending a message */ 
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



void log_mapi_property (LPMESSAGE message, ULONG prop, const char *propname);
int get_gpgolattachtype_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolsigstatus_tag (LPMESSAGE message, ULONG *r_tag);
int get_gpgolprotectiv_tag (LPMESSAGE message, ULONG *r_tag);

int mapi_set_header (LPMESSAGE msg, const char *name, const char *val);

int mapi_change_message_class (LPMESSAGE message);
msgtype_t mapi_get_message_type (LPMESSAGE message);
int mapi_to_mime (LPMESSAGE message, const char *filename);

char *mapi_get_binary_prop (LPMESSAGE message,ULONG proptype,size_t *r_nbytes);

LPSTREAM mapi_get_body_as_stream (LPMESSAGE message);
char *mapi_get_body (LPMESSAGE message, size_t *r_nbytes);

mapi_attach_item_t *mapi_create_attach_table (LPMESSAGE message, int fast);
void mapi_release_attach_table (mapi_attach_item_t *table);
LPSTREAM mapi_get_attach_as_stream (LPMESSAGE message, 
                                    mapi_attach_item_t *item);
char *mapi_get_attach (LPMESSAGE message, 
                       mapi_attach_item_t *item, size_t *r_nbytes);
int mapi_mark_moss_attach (LPMESSAGE message, mapi_attach_item_t *item);
int mapi_has_sig_status (LPMESSAGE msg);
int mapi_test_sig_status (LPMESSAGE msg);
int mapi_set_sig_status (LPMESSAGE message, const char *status_string);

char *mapi_get_message_content_type (LPMESSAGE message, 
                                     char **r_protocol, char **r_smtype);

char *mapi_get_gpgol_body_attachment (LPMESSAGE message, size_t *r_nbytes,
                                      int *r_ishtml, int *r_protected);


#ifdef __cplusplus
}
#endif
#endif /*MAPIHELP_H*/
