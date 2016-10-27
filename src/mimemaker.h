/* mimemaker.h - Construct MIME from MAPI
 *	Copyright (C) 2007, 2008 g10 Code GmbH
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
               const char *sender);
int mime_encrypt (LPMESSAGE message, HWND hwnd,
                  protocol_t protocol, char **recipients,
                  const char *sender);
int mime_sign_encrypt (LPMESSAGE message, HWND hwnd,
                       protocol_t protocol, char **recipients,
                       const char *sender);
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

#ifdef __cplusplus
}

class Mail;
int encrypt_sign_mail (Mail *, bool encrypt, bool sign, protocol_t protocol,
                       HWND hwnd);
#endif
#endif /*MIMEMAKER_H*/
