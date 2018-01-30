/* mimemaker.h - Construct MIME from MAPI
 * Copyright (C) 2007, 2008 g10 Code GmbH
 * Copyright (C) 2015, 2016 by Bundesamt für Sicherheit in der Informationstechnik
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

#ifndef MIMEMAKER_H
#define MIMEMAKER_H

class Mail;
#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/* The object we use instead of IStream.  It allows us to have a
   callback method for output and thus for processing stuff
   recursively.  */
struct sink_s;
typedef struct sink_s *sink_t;
struct sink_s
{
  void *cb_data;
  sink_t extrasink;
  int (*writefnc)(sink_t sink, const void *data, size_t datalen);
  unsigned long enc_counter; /* Used by write_buffer_for_cb.  */
/*   struct { */
/*     int idx; */
/*     unsigned char inbuf[4]; */
/*     int quads; */
/*   } b64; */
};

int mime_sign (LPMESSAGE message, HWND hwnd, protocol_t protocol,
               const char *sender, Mail* mail);
int mime_encrypt (LPMESSAGE message, HWND hwnd,
                  protocol_t protocol, char **recipients,
                  const char *sender, Mail* mail);
int mime_sign_encrypt (LPMESSAGE message, HWND hwnd,
                       protocol_t protocol, char **recipients,
                       const char *sender, Mail* mail);
int sink_std_write (sink_t sink, const void *data, size_t datalen);
int sink_file_write (sink_t sink, const void *data, size_t datalen);
int sink_encryption_write (sink_t encsink, const void *data, size_t datalen);
int write_buffer_for_cb (void *opaque, const void *data, size_t datalen);
int write_buffer (sink_t sink, const void *data, size_t datalen);

/** @brief Try to restore a message from the moss attachment.
  *
  * Try to turn the moss attachment back into a Mail that other
  * MUAs could handle. Uses all the tricks available to archive
  * that. Returns 0 on success.
  */
int restore_msg_from_moss (LPMESSAGE message, LPDISPATCH moss_att,
                           msgtype_t type, char *msgcls);

int count_usable_attachments (mapi_attach_item_t *table);

int add_body_and_attachments (sink_t sink, LPMESSAGE message,
                              mapi_attach_item_t *att_table, Mail *mail,
                              const char *body, int n_att_usable);
int create_top_encryption_header (sink_t sink, protocol_t protocol, char *boundary,
                              bool is_inline = false);

/* Helper to write a boundary to the output sink.  The leading LF
   will be written as well.  */
int write_boundary (sink_t sink, const char *boundary, int lastone);

LPATTACH create_mapi_attachment (LPMESSAGE message, sink_t sink);
int close_mapi_attachment (LPATTACH *attach, sink_t sink);
int finalize_message (LPMESSAGE message, mapi_attach_item_t *att_table,
                      protocol_t protocol, int encrypt, bool is_inline = false);
void cancel_mapi_attachment (LPATTACH *attach, sink_t sink);
void create_top_signing_header (char *buffer, size_t buflen, protocol_t protocol,
                           int first, const char *boundary, const char *micalg);
int write_string (sink_t sink, const char *text);

#ifdef __cplusplus
}
#endif
#endif /*MIMEMAKER_H*/
