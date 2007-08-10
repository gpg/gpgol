/* message.cpp - Functions for message handling
 *	Copyright (C) 2006, 2007 g10 Code GmbH
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "mapihelp.h"
#include "mimeparser.h"
#include "message.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)



/* Verify MESSAGE and update the attachments as required.  MSGTYPE
   should be the type of the message so that the fucntion can decide
   what to do.  With FORCE set the verification is done regardlessless
   of a cached signature result. */
int
message_verify (LPMESSAGE message, msgtype_t msgtype, int force)
{
  mapi_attach_item_t *table;
  int moss_idx = -1;
  int i;
  char *inbuf;
  size_t inbuflen;
  int err;

  switch (msgtype)
    {
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      break;
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
      log_debug ("Opaque signed message are not yet supported!");
      return 0;
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
      return -1; /* Should not be called for such a message.  */
    case MSGTYPE_UNKNOWN:
    case MSGTYPE_GPGOL:
      return 0; /* Nothing to do.  */
    }
  
  /* If a verification is forced, we set the cached signature status
     first to "?" to mark that no verification has yet happened. */
  if (force)
    mapi_set_sig_status (message, "?");
  else if (mapi_has_sig_status (message))
    return 0; /* Already checked that message.  */

  table = mapi_create_attach_table (message, 0);
  if (!table)
    return -1; /* No attachment - this should not happen.  */


  for (i=0; !table[i].end_of_table; i++)
    if (table[i].attach_type == ATTACHTYPE_MOSS)
      {
        moss_idx = i;
        break;
      }
  if (moss_idx == -1 && !table[0].end_of_table && table[1].end_of_table)
    {
      /* No MOSS flag found in the table but there is only one
         attachment.  Due to the message type we know that this is
         the original MOSS message.  We mark this attachment as
         hidden, so that it won't get displayed.  We further mark it
         as our original MOSS attachment so that after parsing we have
         a mean to find it again (see above).  */ 
      moss_idx = 0;
      mapi_mark_moss_attach (message, table+0);
    }

  if (moss_idx == -1)
    {
      mapi_release_attach_table (table);
      return -1; /* No original attachment - this should not happen.  */
    }


  inbuf = mapi_get_attach (message, table+0, &inbuflen);
  if (!inbuf)
    {
      mapi_release_attach_table (table);
      return -1; /* Problem getting the attachment.  */
    }


  err = mime_verify (inbuf, inbuflen, message, 0, 
                     opt.passwd_ttl, NULL, NULL, 0);
  log_debug ("mime_verify returned %d", err);
  xfree (inbuf);
                    
  if (err)
    mapi_set_sig_status (message, gpg_strerror (err));
  else
    mapi_set_sig_status (message, "Signature was good");

  mapi_release_attach_table (table);
  return 0;
}

/* Decrypt MESSAGE, check signature and update the attachments as
   required.  MSGTYPE should be the type of the message so that the
   function can decide what to do.  With FORCE set the verification is
   done regardlessless of a cached signature result - hmmm, should we
   such a thing for an encrypted message? */
int
message_decrypt (LPMESSAGE message, msgtype_t msgtype, int force)
{
  mapi_attach_item_t *table;
  int part2_idx;
  int tblidx;
  int retval = -1;
  LPSTREAM cipherstream;
  gpg_error_t err;
  int is_opaque = 0;
  int is_smime = 0;

  switch (msgtype)
    {
    case MSGTYPE_UNKNOWN:
    case MSGTYPE_GPGOL:
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      return -1; /* Should not have been called for this.  */
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      break;
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
      is_opaque = 1;
      break;
    }
  
  table = mapi_create_attach_table (message, 0);
  if (!table)
    return -1; /* No attachment - this should not happen.  */

  if (is_opaque)
    {
      /* S/MIME opaque encrypted message: We expect 1 attachment.  As
         we don't know ether we are called the first time, we first
         try to find this attachment by looking at all attachments.
         Only if this fails we identify it by its order.  */
      part2_idx = -1;
      for (tblidx=0; !table[tblidx].end_of_table; tblidx++)
        if (table[tblidx].attach_type == ATTACHTYPE_MOSS)
          {
            if (part2_idx == -1 && table[tblidx].content_type 
                && (!strcmp (table[tblidx].content_type,
                            "application/pkcs7-mime")
                    || !strcmp (table[tblidx].content_type,
                                "application/x-pkcs7-mime")))
              part2_idx = tblidx;
          }
      if (part2_idx == -1 && tblidx >= 1)
        {
          /* We have attachments but none are marked.  Thus we assume
             that this is the first time we see this message and we
             will set the mark now if we see appropriate content
             types. */
          if (table[0].content_type               
              && (!strcmp (table[0].content_type, "application/pkcs7-mime")
                  || !strcmp (table[0].content_type,
                              "application/x-pkcs7-mime")))
            part2_idx = 0;
          if (part2_idx != -1)
            mapi_mark_moss_attach (message, table+part2_idx);
        }
      if (part2_idx == -1)
        {
          log_debug ("%s:%s: this is not an S/MIME encrypted message",
                     SRCNAME, __func__);
          goto leave;
        }
      is_smime = 1;
    }
  else 
    {
      /* Multipart/encrypted message: We expect 2 attachments.  The
         first one with the version number and the second one with the
         ciphertext.  As we don't know ether we are called the first
         time, we first try to find these attachments by looking at
         all attachments.  Only if this fails we identify them by
         their order (i.e. the first 2 attachments) and mark them as
         part1 and part2.  */
      int part1_idx;

      part1_idx = part2_idx = -1;
      for (tblidx=0; !table[tblidx].end_of_table; tblidx++)
        if (table[tblidx].attach_type == ATTACHTYPE_MOSS)
          {
            if (part1_idx == -1 && table[tblidx].content_type 
                && !strcmp (table[tblidx].content_type,
                            "application/pgp-encrypted"))
              part1_idx = tblidx;
            else if (part2_idx == -1 && table[tblidx].content_type 
                     && !strcmp (table[tblidx].content_type,
                                 "application/octet-stream"))
              part2_idx = tblidx;
          }
      if (part1_idx == -1 && part2_idx == -1 && tblidx >= 2)
        {
          /* At least 2 attachments but none are marked.  Thus we
             assume that this is the first time we see this message
             and we will set the mark now if we see appropriate
             content types. */
          if (table[0].content_type               
              && !strcmp (table[0].content_type, "application/pgp-encrypted"))
            part1_idx = 0;
          if (table[1].content_type             
              && !strcmp (table[1].content_type, "application/octet-stream"))
            part2_idx = 1;
          if (part1_idx != -1 && part2_idx != -1)
            {
              mapi_mark_moss_attach (message, table+part1_idx);
              mapi_mark_moss_attach (message, table+part2_idx);
            }
        }
      if (part1_idx == -1 || part2_idx == -1)
        {
          log_debug ("%s:%s: this is not a PGP/MIME encrypted message",
                 SRCNAME, __func__);
          goto leave;
        }
    }

  /* Get the attachment as an allocated buffer and let the mimeparser
     work on it.  */
  cipherstream = mapi_get_attach_as_stream (message, table+part2_idx);
  if (!cipherstream)
    goto leave; /* Problem getting the attachment.  */

  err = mime_decrypt (cipherstream, message, is_smime, opt.passwd_ttl,
                      NULL, NULL, 0);
  log_debug ("mime_decrypt returned %d (%s)", err, gpg_strerror (err));
  cipherstream->Release ();
  retval = 0;

 leave:
  mapi_release_attach_table (table);
  return retval;
}
