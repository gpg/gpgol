/* message.cpp - Functions for message handling
 *	Copyright (C) 2006, 2007, 2008 g10 Code GmbH
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
#include "mimemaker.h"
#include "display.h"
#include "message.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


static void 
ul_release (LPVOID punk, const char *func)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
  log_debug ("%s:%s: UlRelease(%p) had %lu references\n", 
             SRCNAME, func, punk, res);
}


/* A helper function used by OnRead and OnOpen to dispatch the
   message.  Returns true if the message has been processed.  */
bool
message_incoming_handler (LPMESSAGE message, HWND hwnd)
{
  bool retval = false;
  msgtype_t msgtype;
  int pass = 0;

 retry:
  pass++;
  msgtype = mapi_get_message_type (message);
  switch (msgtype)
    {
    case MSGTYPE_UNKNOWN: 
      /* If this message has never passed our change message class
         code it won't have an unknown msgtype _and_ no sig status
         flag.  Thus we look at the message class now and change it if
         required.  It won't get displayed correctly right away but a
         latter decrypt command or when viewed a second time all has
         been set.  Note that we should have similar code for some
         message classes in GpgolUserEvents:OnSelectionChange; but
         tehre are a couiple of problems.  */
      if (pass == 1 && !mapi_has_sig_status (message))
        {
          log_debug ("%s:%s: message class not yet checked - doing now\n",
                     SRCNAME, __func__);
          if (mapi_change_message_class (message))
            goto retry;
        }
      break;
    case MSGTYPE_SMIME:
      if (pass == 1 && opt.enable_smime)
        {
          log_debug ("%s:%s: message class not checked with smime enabled "
                     "- doing now\n", SRCNAME, __func__);
          if (mapi_change_message_class (message))
            goto retry;
        }
      break;
    case MSGTYPE_GPGOL:
      log_debug ("%s:%s: ignoring unknown message of original SMIME class\n",
                 SRCNAME, __func__);
      break;
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      log_debug ("%s:%s: processing multipart signed message\n", 
                 SRCNAME, __func__);
      retval = true;
      message_verify (message, msgtype, 0, hwnd);
      break;
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      log_debug ("%s:%s: processing multipart encrypted message\n",
                 SRCNAME, __func__);
      retval = true;
      message_decrypt (message, msgtype, 0, hwnd);
      break;
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
      log_debug ("%s:%s: processing opaque signed message\n", 
                 SRCNAME, __func__);
      retval = true;
      message_verify (message, msgtype, 0, hwnd);
      break;
    case MSGTYPE_GPGOL_CLEAR_SIGNED:
      log_debug ("%s:%s: processing clear signed pgp message\n", 
                 SRCNAME, __func__);
      retval = true;
      message_verify (message, msgtype, 0, hwnd);
      break;
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
      log_debug ("%s:%s: processing opaque encrypted message\n",
                 SRCNAME, __func__);
      retval = true;
      message_decrypt (message, msgtype, 0, hwnd);
      break;
    case MSGTYPE_GPGOL_PGP_MESSAGE:
      log_debug ("%s:%s: processing pgp message\n", SRCNAME, __func__);
      retval = true;
      message_decrypt (message, msgtype, 0, hwnd);
      break;
    }

  return retval;
}


/* Common Code used by OnReadComplete and OnOpenComplete to display a
   modified message.   Returns true if the message was encrypted.  */
bool
message_display_handler (LPEXCHEXTCALLBACK eecb, HWND hwnd)
{
  HRESULT hr;
  LPMESSAGE message = NULL;
  LPMDB mdb = NULL;
  int ishtml, wasprotected = false;
  char *body;
  
  hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
  if (SUCCEEDED (hr))
    {
      /* (old: If the message was protected we don't allow a fallback to the
         OOM display methods.)  Now: As it is difficult to find the
         actual winodw we now use the OOM display always.  */
      body = mapi_get_gpgol_body_attachment (message, NULL, 
                                             &ishtml, &wasprotected);
      if (body)
        update_display (hwnd, /*wasprotected? NULL:*/ eecb, ishtml, body);
      else
        update_display (hwnd, NULL, 0, 
                        _("[Crypto operation failed - "
                          "can't show the body of the message]"));
      xfree (body);
  
      /*  put_outlook_property (eecb, "EncryptedStatus", "MyStatus"); */
    }
  else
    log_debug_w32 (hr, "%s:%s: error getting message", SRCNAME, __func__);

  ul_release (message, __func__);
  ul_release (mdb, __func__);

  return !!wasprotected;
}


/* If the current message is an encrypted one remove the body
   properties which might have come up due to OL internal
   syncronization and a failing olDiscard feature.  */
void
message_wipe_body_cruft (LPEXCHEXTCALLBACK eecb)
{
  
  HRESULT hr;
  LPMESSAGE message = NULL;
  LPMDB mdb = NULL;
      
  log_debug ("%s:%s: enter", SRCNAME, __func__);
  hr = eecb->GetObject (&mdb, (LPMAPIPROP *)&message);
  if (SUCCEEDED (hr))
    {
      if (mapi_has_last_decrypted (message))
        {
          SPropTagArray proparray;
          int anyokay = 0;
          
          proparray.cValues = 1;
          proparray.aulPropTag[0] = PR_BODY;
          hr = message->DeleteProps (&proparray, NULL);
          if (hr)
            log_debug_w32 (hr, "%s:%s: deleting PR_BODY failed",
                           SRCNAME, __func__);
          else
            anyokay++;
          
          proparray.cValues = 1;
          proparray.aulPropTag[0] = PR_BODY_HTML;
          message->DeleteProps (&proparray, NULL);
          if (hr)
            log_debug_w32 (hr, "%s:%s: deleting PR_BODY_HTML failed", 
                           SRCNAME, __func__);
          else
            anyokay++;

          if (anyokay)
            {
              hr = message->SaveChanges (KEEP_OPEN_READWRITE);
              if (hr)
                log_error_w32 (hr, "%s:%s: SaveChanges failed",
                               SRCNAME, __func__); 
              else
                log_debug ("%s:%s: SaveChanges succeded; body cruft removed",
                           SRCNAME, __func__); 
            }
        }  
      else
        log_debug_w32 (hr, "%s:%s: error getting message", 
                       SRCNAME, __func__);
     
      ul_release (message, __func__);
      ul_release (mdb, __func__);
    }
}



/* Display some information about MESSAGE.  */
void
message_show_info (LPMESSAGE message, HWND hwnd)
{
  char *msgcls = mapi_get_message_class (message);
  char *sigstat = mapi_get_sig_status (message);
  char *mimeinfo = mapi_get_mime_info (message);
  size_t buflen;
  char *buffer;

  buflen = strlen (msgcls) + strlen (sigstat) + strlen (mimeinfo) + 200;
  buffer = (char*)xmalloc (buflen+1);
  snprintf (buffer, buflen, 
            _("Message class: %s\n"
              "Sig Status   : %s\n"
              "Structure of the message:\n"
              "%s"), 
            msgcls,
            sigstat,
            mimeinfo);
  
  MessageBox (hwnd, buffer, _("GpgOL - Message Information"),
              MB_ICONINFORMATION|MB_OK);
  xfree (buffer);
  xfree (mimeinfo);
  xfree (sigstat);
  xfree (msgcls);
}




/* Convert the clear signed message from INPUT into a PS?MIME signed
   message and return it in a new allocated buffer.  OUTPUTLEN
   received the valid length of that buffer; the buffer is guarnateed
   to be Nul terminated.  */
static char *
pgp_mime_from_clearsigned (LPSTREAM input, size_t *outputlen)
{
  HRESULT hr;
  STATSTG statinfo;
  ULONG nread;
  char *body = NULL;
  char *p, *p0, *dest, *mark;
  char boundary[BOUNDARYSIZE+1];
  char top_header[200 + 2*BOUNDARYSIZE]; 
  char sig_header[100 + BOUNDARYSIZE]; 
  char end_header[10 + BOUNDARYSIZE];
  size_t reserved_space;
  int state;

  *outputlen = 0;

  /* Note that our parser does not make use of the micalg parameter.  */
  generate_boundary (boundary);
  snprintf (top_header, sizeof top_header,
            "MIME-Version: 1.0\r\n"
            "Content-Type: multipart/signed; boundary=\"%s\";\r\n"
            "              protocol=\"application/pgp-signature\"\r\n"
            "\r\n"
            "--%s\r\n", boundary, boundary);
  snprintf (sig_header, sizeof sig_header,
            "--%s\r\n"
            "Content-Type: application/pgp-signature\r\n"
            "\r\n", boundary);
  snprintf (end_header, sizeof end_header,
            "\r\n"
            "--%s--\r\n", boundary);
  reserved_space = (strlen (top_header) + strlen (sig_header) 
                    + strlen (end_header)+ 100);

  hr = input->Stat (&statinfo, STATFLAG_NONAME);
  if (hr)
    {
      log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
      return NULL;
    }
      
  body = (char*)xmalloc (reserved_space
                         + (size_t)statinfo.cbSize.QuadPart + 2);
  hr = input->Read (body+reserved_space,
                    (size_t)statinfo.cbSize.QuadPart, &nread);
  if (hr)
    {
      log_debug ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      xfree (body);
      return NULL;
    }
  body[reserved_space + nread] = 0;
  body[reserved_space + nread+1] = 0;  /* Just in case this is
                                          accidently an wchar_t. */
  if (nread != statinfo.cbSize.QuadPart)
    {
      log_debug ("%s:%s: not enough bytes returned\n", SRCNAME, __func__);
      xfree (body);
      return NULL;
    }

  /* We do in place conversion. */
  state = 0;
  dest = NULL;
  for (p=body+reserved_space; p && *p; p = (p=strchr (p+1, '\n'))? (p+1):NULL)
    {
      if (!state && !strncmp (p, "-----BEGIN PGP SIGNED MESSAGE-----", 34)
          && trailing_ws_p (p+34) )
        {
          dest = stpcpy (body, top_header);
          state = 1;
        }
      else if (state == 1)
        {
          /* Wait for an empty line.  */
          if (trailing_ws_p (p))
            state = 2;
        }
      else if (state == 2 && strncmp (p, "-----", 5))
        {
          /* Copy signed data. */
          p0 = p;
          if (*p == '-' && p[1] == ' ')
            p +=2;  /* Remove escaping.  */
          mark = NULL;
          while (*p && *p != '\n')
            {
              if (*p == ' ' || *p == '\t' || *p == '\r')
                mark = p;
              else
                mark = NULL;
              *dest++ = *p++;
            }
          if (mark)
            *mark =0; /* Remove trailing white space.  */
          if (*p == '\n')
            {
              if (p[-1] == '\r')
                *dest++ = '\r';
              *dest++ = '\n';
            }
          if (p > p0)
            p--; /* Adjust so that the strchr (p+1, '\n') can work. */
        }
      else if (state == 2)
        {
          /* Armor line encountered.  */
          p0 = p;
          if (strncmp (p, "-----BEGIN PGP SIGNATURE-----", 29)
              || !trailing_ws_p (p+29) )
            fprintf (stderr,"Invalid clear signed message\n");
          state = 3;
          dest = stpcpy (dest, sig_header);
        
          while (*p && *p != '\n')
            *dest++ = *p++;
          if (*p == '\n')
            {
              if (p[-1] == '\r')
                *dest++ = '\r';
              *dest++ = '\n';
            }
          if (p > p0)
            p--; /* Adjust so that the strchr (p+1, '\n') can work. */
        }
      else if (state == 3 && strncmp (p, "-----", 5))
        {
          /* Copy signature.  */
          p0 = p;
          while (*p && *p != '\n')
            *dest++ = *p++;
          if (*p == '\n')
            {
              if (p[-1] == '\r')
                *dest++ = '\r';
              *dest++ = '\n';
            }
          if (p > p0)
            p--; /* Adjust so that the strchr (p+1, '\n') can work. */
        }
      else if (state == 3)
        {
          /* Armor line encountered.  */
          p0 = p;
          if (strncmp (p, "-----END PGP SIGNATURE-----", 27)
              || !trailing_ws_p (p+27) )
            fprintf (stderr,"Invalid clear signed message (no end)\n");
          while (*p && *p != '\n')
            *dest++ = *p++;
          if (*p == '\n')
            {
              if (p[-1] == '\r')
                *dest++ = '\r';
              *dest++ = '\n';
            }
          dest = stpcpy (dest, end_header);
          if (p > p0)
            p--; /* Adjust so that the strchr (p+1, '\n') can work. */
          break; /* Ah well, we can stop here.  */
        }
    }
  if (!dest)
    {
      xfree (body);
      return NULL;
    }
  *dest = 0;
  *outputlen = strlen (body);

  return body;
}



/* Verify MESSAGE and update the attachments as required.  MSGTYPE
   should be the type of the message so that the fucntion can decide
   what to do.  With FORCE set the verification is done regardlessless
   of a cached signature result. */
int
message_verify (LPMESSAGE message, msgtype_t msgtype, int force, HWND hwnd)
{
  HRESULT hr;
  mapi_attach_item_t *table = NULL;
  int moss_idx = -1;
  int i;
  char *inbuf;
  size_t inbuflen;
  protocol_t protocol = PROTOCOL_UNKNOWN;
  int err;

  switch (msgtype)
    {
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
      break;
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
      log_debug ("Opaque signed message are not yet supported!");
      return 0;
    case MSGTYPE_GPGOL_CLEAR_SIGNED:
      break;
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
    case MSGTYPE_GPGOL_PGP_MESSAGE:
      log_debug ("%s:%s: message of type %d not expected",
                 SRCNAME, __func__, msgtype);
      return -1; /* Should not be called for such a message.  */
    case MSGTYPE_UNKNOWN:
    case MSGTYPE_SMIME:
    case MSGTYPE_GPGOL:
      log_debug ("%s:%s: message of type %d ignored",
                 SRCNAME, __func__, msgtype);
      return 0; /* Nothing to do.  */
    }
  
  /* If a verification is forced, we set the cached signature status
     first to "?" to mark that no verification has yet happened. */
  if (force)
    mapi_set_sig_status (message, "?");
  else if (mapi_test_sig_status (message))
    return 0; /* Already checked that message.  */

  if (msgtype == MSGTYPE_GPGOL_CLEAR_SIGNED)
    {
      /* PGP's clear signed messages are special: All is contained in
         the body and thus there is no requirement for an
         attachment.  */
      LPSTREAM rawstream;
      
      rawstream = mapi_get_body_as_stream (message);
      if (!rawstream)
        return -1;
      
      inbuf = pgp_mime_from_clearsigned (rawstream, &inbuflen);
      rawstream->Release ();
      if (!inbuf)
        return -1;
      protocol = PROTOCOL_OPENPGP;
    }
  else
    {
      /* PGP/MIME or S/MIME stuff.  */
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
             hidden, so that it won't get displayed.  We further mark
             it as our original MOSS attachment so that after parsing
             we have a mean to find it again (see above).  */ 
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
    }

  err = mime_verify (protocol, inbuf, inbuflen, message, hwnd, 0);
  log_debug ("mime_verify returned %d", err);
  if (err)
    {
      char buf[200];
      
      snprintf (buf, sizeof buf, "Verify failed (%s)", gpg_strerror (err));
      MessageBox (NULL, buf, "GpgOL", MB_ICONINFORMATION|MB_OK);
    }
  xfree (inbuf);
                    
  if (err)
    {
      char buf[200];
      snprintf (buf, sizeof buf, "- %s", gpg_strerror (err));
      mapi_set_sig_status (message, gpg_strerror (err));
    }
  else
    mapi_set_sig_status (message, "! Good signature");

  hr = message->SaveChanges (KEEP_OPEN_READWRITE);
  if (hr)
    log_error_w32 (hr, "%s:%s: SaveChanges failed",
                   SRCNAME, __func__); 


  mapi_release_attach_table (table);
  return 0;
}

/* Decrypt MESSAGE, check signature and update the attachments as
   required.  MSGTYPE should be the type of the message so that the
   function can decide what to do.  With FORCE set the decryption is
   done regardless whether it has already been done.  */
int
message_decrypt (LPMESSAGE message, msgtype_t msgtype, int force, HWND hwnd)
{
  mapi_attach_item_t *table = NULL;
  int part2_idx;
  int tblidx;
  int retval = -1;
  LPSTREAM cipherstream;
  gpg_error_t err;
  int is_opaque = 0;
  protocol_t protocol;

  switch (msgtype)
    {
    case MSGTYPE_UNKNOWN:
    case MSGTYPE_SMIME:
    case MSGTYPE_GPGOL:
    case MSGTYPE_GPGOL_OPAQUE_SIGNED:
    case MSGTYPE_GPGOL_MULTIPART_SIGNED:
    case MSGTYPE_GPGOL_CLEAR_SIGNED:
      return -1; /* Should not have been called for this.  */
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
      break;
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
      is_opaque = 1;
      break;
    case MSGTYPE_GPGOL_PGP_MESSAGE:
      break;
    }
  
  if (!force && mapi_test_last_decrypted (message))
    return 0; /* Already decrypted this message once during this
                 session.  No need to do it again. */

  if (msgtype == MSGTYPE_GPGOL_PGP_MESSAGE)
    {
      /* PGP messages are special:  All is contained in the body and thus
         there is no requirement for an attachment.  */
      cipherstream = mapi_get_body_as_stream (message);
      if (!cipherstream)
        goto leave;
      protocol = PROTOCOL_OPENPGP;
    }
  else
    {
  
      /* PGP/MIME or S/MIME stuff.  */
      table = mapi_create_attach_table (message, 0);
      if (!table)
        goto leave; /* No attachment - this should not happen.  */

      if (is_opaque)
        {
          /* S/MIME opaque encrypted message: We expect one
             attachment.  As we don't know ether we are called the
             first time, we first try to find this attachment by
             looking at all attachments.  Only if this fails we
             identify it by its order.  */
          part2_idx = -1;
          for (tblidx=0; !table[tblidx].end_of_table; tblidx++)
            if (table[tblidx].attach_type == ATTACHTYPE_MOSSTEMPL)
              {
                /* This attachment has been generated by us in the
                   course of sendeing a new message.  The content will
                   be multipart/signed because we used this to trick
                   out OL.  We stop here and use this part for further
                   processing.  */
                part2_idx = tblidx;
                break;
              }
            else if (table[tblidx].attach_type == ATTACHTYPE_MOSS)
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
              /* We have attachments but none are marked.  Thus we
                 assume that this is the first time we see this
                 message and we will set the mark now if we see
                 appropriate content types. */
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
          protocol = PROTOCOL_SMIME;
        }
      else 
        {
          /* Multipart/encrypted message: We expect 2 attachments.
             The first one with the version number and the second one
             with the ciphertext.  As we don't know ether we are
             called the first time, we first try to find these
             attachments by looking at all attachments.  Only if this
             fails we identify them by their order (i.e. the first 2
             attachments) and mark them as part1 and part2.  */
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
                 assume that this is the first time we see this
                 message and we will set the mark now if we see
                 appropriate content types. */
              if (table[0].content_type               
                  && !strcmp (table[0].content_type,
                              "application/pgp-encrypted"))
                part1_idx = 0;
              if (table[1].content_type             
                  && !strcmp (table[1].content_type, 
                              "application/octet-stream"))
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
          protocol = PROTOCOL_OPENPGP;
        }
      
      cipherstream = mapi_get_attach_as_stream (message, table+part2_idx);
      if (!cipherstream)
        goto leave; /* Problem getting the attachment.  */
    }

  err = mime_decrypt (protocol, cipherstream, message, hwnd, 0);
  log_debug ("mime_decrypt returned %d (%s)", err, gpg_strerror (err));
  if (err)
    {
      char buf[200];
      
      snprintf (buf, sizeof buf, "Decryption failed (%s)", gpg_strerror (err));
      MessageBox (NULL, buf, "GpgOL", MB_ICONINFORMATION|MB_OK);
    }
  cipherstream->Release ();
  retval = 0;

 leave:
  mapi_release_attach_table (table);
  return retval;
}




/* Return an array of strings with the recipients of the message. On
   success a malloced array is returned containing allocated strings
   for each recipient.  The end of the array is marked by NULL.
   Caller is responsible for releasing the array.  On failure NULL is
   returned.  */
static char **
get_recipients (LPMESSAGE message)
{
  static SizedSPropTagArray (1L, PropRecipientNum) = {1L, {PR_EMAIL_ADDRESS}};
  HRESULT hr;
  LPMAPITABLE lpRecipientTable = NULL;
  LPSRowSet lpRecipientRows = NULL;
  unsigned int rowidx;
  LPSPropValue row;
  char **rset;
  int rsetidx;


  if (!message)
    return NULL;

  hr = message->GetRecipientTable (0, &lpRecipientTable);
  if (FAILED (hr)) 
    {
      log_debug_w32 (-1, "%s:%s: GetRecipientTable failed", SRCNAME, __func__);
      return NULL;
    }

  hr = HrQueryAllRows (lpRecipientTable, (LPSPropTagArray) &PropRecipientNum,
                       NULL, NULL, 0L, &lpRecipientRows);
  if (FAILED (hr)) 
    {
      log_debug_w32 (-1, "%s:%s: HrQueryAllRows failed", SRCNAME, __func__);
      if (lpRecipientTable)
        lpRecipientTable->Release();
      return NULL;
    }

  rset = (char**)xcalloc (lpRecipientRows->cRows+1, sizeof *rset);

  for (rowidx=0, rsetidx=0; rowidx < lpRecipientRows->cRows; rowidx++)
    {
      if (!lpRecipientRows->aRow[rowidx].cValues)
        continue;
      row = lpRecipientRows->aRow[rowidx].lpProps;

      switch (PROP_TYPE (row->ulPropTag))
        {
        case PT_UNICODE:
          if ((rset[rsetidx] = wchar_to_utf8 (row->Value.lpszW)))
            rsetidx++;
          else
            log_debug ("%s:%s: error converting recipient to utf8\n",
                       SRCNAME, __func__);
          break;
      
        case PT_STRING8: /* Assume ASCII. */
          rset[rsetidx++] = xstrdup (row->Value.lpszA);
          break;
          
        default:
          log_debug ("%s:%s: proptag=0x%08lx not supported\n",
                     SRCNAME, __func__, row->ulPropTag);
          break;
        }
    }

  if (lpRecipientTable)
    lpRecipientTable->Release();
  if (lpRecipientRows)
    FreeProws(lpRecipientRows);	
  
  log_debug ("%s:%s: got %d recipients:\n", SRCNAME, __func__, rsetidx);
  for (rsetidx=0; rset[rsetidx]; rsetidx++)
    log_debug ("%s:%s: \t`%s'\n", SRCNAME, __func__, rset[rsetidx]);

  return rset;
}


static void
release_recipient_array (char **recipients)
{
  int idx;

  if (recipients)
    {
      for (idx=0; recipients[idx]; idx++)
        xfree (recipients[idx]);
      xfree (recipients);
    }
}



static int
sign_encrypt (LPMESSAGE message, protocol_t protocol, HWND hwnd, int signflag)
{
  gpg_error_t err;
  char **recipients;

  recipients = get_recipients (message);
  if (!recipients || !recipients[0])
    {
      MessageBox (hwnd, _("No recipients to encrypt to are given"),
                  "GpgOL", MB_ICONERROR|MB_OK);

      err = gpg_error (GPG_ERR_GENERAL);
    }
  else
    {
      if (signflag)
        err = mime_sign_encrypt (message, hwnd, protocol, recipients);
      else
        err = mime_encrypt (message, hwnd, protocol, recipients);
      if (err)
        {
          char buf[200];
          
          snprintf (buf, sizeof buf,
                    _("Encryption failed (%s)"), gpg_strerror (err));
          MessageBox (hwnd, buf, "GpgOL", MB_ICONERROR|MB_OK);
        }
    }
  release_recipient_array (recipients);
  return err;
}


/* Sign the MESSAGE.  */
int 
message_sign (LPMESSAGE message, protocol_t protocol, HWND hwnd)
{
  gpg_error_t err;

  err = mime_sign (message, hwnd, protocol);
  if (err)
    {
      char buf[200];
      
      snprintf (buf, sizeof buf,
                _("Signing failed (%s)"), gpg_strerror (err));
      MessageBox (hwnd, buf, "GpgOL", MB_ICONERROR|MB_OK);
    }
  return err;
}



/* Encrypt the MESSAGE.  */
int 
message_encrypt (LPMESSAGE message, protocol_t protocol, HWND hwnd)
{
  return sign_encrypt (message, protocol, hwnd, 0);
}


/* Sign+Encrypt the MESSAGE.  */
int 
message_sign_encrypt (LPMESSAGE message, protocol_t protocol, HWND hwnd)
{
  return sign_encrypt (message, protocol, hwnd, 1);
}



