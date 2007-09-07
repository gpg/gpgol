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
message_verify (LPMESSAGE message, msgtype_t msgtype, int force)
{
  mapi_attach_item_t *table = NULL;
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
    case MSGTYPE_GPGOL_CLEAR_SIGNED:
      break;
    case MSGTYPE_GPGOL_MULTIPART_ENCRYPTED:
    case MSGTYPE_GPGOL_OPAQUE_ENCRYPTED:
    case MSGTYPE_GPGOL_PGP_MESSAGE:
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

  err = mime_verify (inbuf, inbuflen, message, 0, 
                     opt.passwd_ttl, NULL, NULL, 0);
  log_debug ("mime_verify returned %d", err);
  if (err)
    {
      char buf[200];
      
      snprintf (buf, sizeof buf, "Verify failed (%s)", gpg_strerror (err));
      MessageBox (NULL, buf, "GpgOL", MB_ICONINFORMATION|MB_OK);
    }
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
  mapi_attach_item_t *table = NULL;
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
  
  if (msgtype == MSGTYPE_GPGOL_PGP_MESSAGE)
    {
      /* PGP messages are special:  All is contained in the body and thus
         there is no requirement for an attachment.  */
      cipherstream = mapi_get_body_as_stream (message);
      if (!cipherstream)
        goto leave;
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
          is_smime = 1;
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
        }
      
      cipherstream = mapi_get_attach_as_stream (message, table+part2_idx);
      if (!cipherstream)
        goto leave; /* Problem getting the attachment.  */
    }

  err = mime_decrypt (cipherstream, message, is_smime, opt.passwd_ttl,
                      NULL, NULL, 0);
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




#if 0
/* Sign the current message. Returns 0 on success. */
int
message_sign (LPMESSAGE message, HWND hwnd, int want_html)
{
  HRESULT hr;
  STATSTG statinfo;
  LPSTREAM plaintext;
  size_t plaintextlen;
  char *signedtext = NULL;
  int err = 0;
  gpgme_key_t sign_key = NULL;
  SPropValue prop;
  int have_html_attach = 0;

  log_debug ("%s:%s: enter message=%p\n", SRCNAME, __func__, message);
  
  /* We don't sign an empty body - a signature on a zero length string
     is pretty much useless.  We assume that a HTML message always
     comes with a text/plain alternative. */
  plaintext = mapi_get_body_as_stream (message);
  if (!plaintext)
    plaintextlen = 0;
  else
    {
      hr = input->Stat (&statinfo, STATFLAG_NONAME);
      if (hr)
        {
          log_debug ("%s:%s: Stat failed: hr=%#lx", SRCNAME, __func__, hr);
          plaintext->Release ();
          return gpg_error (GPG_ERR_GENERAL);
        }
      plaintextlen = (size_t)statinfo.cbSize.QuadPart;
    }

  if ( !plaintextlen && !has_attachments (message)) 
    {
      log_debug ("%s:%s: leave (empty)", SRCNAME, __func__);
      plaintext->Release ();
      return 0; 
    }

  /* Pop up a dialog box to ask for the signer of the message. */
  if (signer_dialog_box (&sign_key, NULL, 0) == -1)
    {
      log_debug ("%s.%s: leave (dialog failed)\n", SRCNAME, __func__);
      plaintext->Release ();
      return gpg_error (GPG_ERR_CANCELED);  
    }

  /* Sign the plaintext */
  if (plaintextlen)
    {
      err = op_sign (plaintext, &signedtext, 
                     OP_SIG_CLEAR, sign_key, opt.passwd_ttl);
      if (err)
        {
          MessageBox (hwnd, op_strerror (err),
                      _("Signing Failure"), MB_ICONERROR|MB_OK);
          plaintext->Release ();
          return gpg_error (GPG_ERR_GENERAL);
        }
    }

  
  /* If those brain dead html mails are requested we now figure out
     whether a HTML body is actually available and move it to an
     attachment so that the code below will sign it as a regular
     attachments.  */
  if (want_html)
    {
      log_debug ("Signing HTML is not yet supported\n");
//       char *htmltext = loadBody (true);
      
//       if (htmltext && *htmltext)
//         {
//           if (!createHtmlAttachment (htmltext))
//             have_html_attach = 1;
//         }
//       xfree (htmltext);

      /* If we got a new attachment we need to release the loaded
         attachment info so that the next getAttachment call will read
         fresh info. */
//       if (have_html_attach)
//         free_attach_info ();
    }


  /* Note, there is a side-effect when we have HTML mails: The
     auto-sign-attch option is ignored.  I regard auto-sign-attach as a
     silly option anyway. */
  if ((opt.auto_sign_attach || have_html_attach) && has_attachments ())
    {
      unsigned int n;
      
      n = getAttachments ();
      log_debug ("%s:%s: message has %u attachments\n", SRCNAME, __func__, n);
      for (unsigned int i=0; i < n; i++) 
        signAttachment (hwnd, i, sign_key, opt.passwd_ttl);
      /* FIXME: we should throw an error if signing of any attachment
         failed. */
    }

  set_x_header (message, "GPGOL-VERSION", PACKAGE_VERSION);

  /* Now that we successfully processed the attachments, we can save
     the changes to the body.  */
  if (plaintextlen)
    {
      err = set_message_body (message, signedtext, 0);
      if (err)
        goto leave;

      /* In case we don't have attachments, Outlook will really insert
         the following content type into the header.  We use this to
         declare that the encrypted content of the message is utf-8
         encoded.  If we have atatchments, OUtlook has its own idea of
         the content type to use.  */
      prop.ulPropTag=PR_CONTENT_TYPE_A;
      prop.Value.lpszA="text/plain; charset=utf-8"; 
      hr = HrSetOneProp (message, &prop);
      if (hr)
        log_error ("%s:%s: can't set content type: hr=%#lx\n",
                   SRCNAME, __func__, hr);
    }
  
  hr = message->SaveChanges (KEEP_OPEN_READWRITE|FORCE_SAVE);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges(message) failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      err = gpg_error (GPG_ERR_GENERAL);
      goto leave;
    }

 leave:
  xfree (signedtext);
  gpgme_key_release (sign_key);
  xfree (plaintext);
  log_debug ("%s:%s: leave (err=%s)\n", SRCNAME, __func__, op_strerror (err));
  return err;
}

#endif
