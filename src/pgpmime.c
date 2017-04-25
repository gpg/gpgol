/* pgpmime.c - Try to handle PGP/MIME for Outlook
 * Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of Gpgol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#error not used

/*
   EXPLAIN what we are doing here.
*/
   

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include <string.h>

#define COBJMACROS
#include <windows.h>
#include <objidl.h> /* For IStream. */

#include <gpgme.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "rfc822parse.h"

#include "common.h"
#include "pgpmime.h"
#include "engine.h"


/* The maximum length of a line we ar able to porcess.  RFC822 alows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000


/* The context object we use to track information. */
struct pgpmime_context
{
  HWND hwnd;          /* A window handle to be used for message boxes etc. */
  rfc822parse_t msg;  /* The handle of the RFC822 parser. */

  int preview;        /* Do only decryption and pop up no  message bozes.  */
  
  int verify_mode;    /* True if we want to verify a PGP/MIME signature. */

  int nesting_level;  /* Current MIME nesting level. */
  int in_data;        /* We are currently in data (body or attachment). */

  gpgme_data_t signed_data;/* NULL or the data object used to collect
                              the signed data. It would bet better to
                              just hash it but there is no support in
                              gpgme for this yet. */
  gpgme_data_t sig_data;  /* NULL or data object to collect the
                             signature attachment which should be a
                             signature then.  */
  
  gpgme_data_t body;      /* NULL or a data object used to collect the
                             body part we are going to display later. */
  int collect_body;       /* True if we are collecting the body lines. */
  int collect_attachment; /* True if we are collecting an attachment. */
  int collect_signeddata; /* True if we are collecting the signed data. */
  int collect_signature;  /* True if we are collecting a signature.  */
  int start_hashing;      /* Flag used to start collecting signed data. */
  int hashing_level;      /* MIME level where we started hashing. */
  int is_qp_encoded;      /* Current part is QP encoded. */
  int is_base64_encoded;  /* Current part is base 64 encoded. */
  int is_utf8;            /* Current part has charset utf-8. */

  int part_counter;       /* Counts the number of processed parts. */
  char *filename;         /* Current filename (malloced) or NULL. */

  LPSTREAM outstream;     /* NULL or a stream to write a part to. */

  b64_state_t base64;     /* The state of the BAse-64 decoder.  */

  int line_too_long;  /* Indicates that a received line was too long. */
  int parser_error;   /* Indicates that we encountered a error from
                         the parser. */

  /* Buffer used to constructed complete files. */
  size_t linebufsize;   /* The allocated size of the buffer. */
  size_t linebufpos;    /* The actual write posituion. */  
  char linebuf[1];      /* The buffer. */
};
typedef struct pgpmime_context *pgpmime_context_t;


/* This function is a wrapper around gpgme_data_write to convert the
   data to utf-8 first.  We assume Latin-1 here. */
static int
latin1_data_write (gpgme_data_t data, const char *line, size_t len)
{
  const char *s;
  char *buffer, *p;
  size_t i, n;
  int rc;

  for (s=line, i=0, n=0 ; i < len; s++, i++ ) 
    {
      n++;
      if (*s & 0x80)
        n++;
    }
  buffer = xmalloc (n + 1);
  for (s=line, i=0, p=buffer; i < len; s++, i++ )
    {
      if (*s & 0x80)
        {
          *p++ = 0xc0 | ((*s >> 6) & 3);
          *p++ = 0x80 | (*s & 0x3f);
        }
      else
        *p++ = *s;
    }
  assert (p-buffer == n);
  rc = gpgme_data_write (data, buffer, n);
  xfree (buffer);
  return rc;
}


/* Print the message event EVENT. */
static void
debug_message_event (pgpmime_context_t ctx, rfc822parse_event_t event)
{
  const char *s;

  switch (event)
    {
    case RFC822PARSE_OPEN: s= "Open"; break;
    case RFC822PARSE_CLOSE: s= "Close"; break;
    case RFC822PARSE_CANCEL: s= "Cancel"; break;
    case RFC822PARSE_T2BODY: s= "T2Body"; break;
    case RFC822PARSE_FINISH: s= "Finish"; break;
    case RFC822PARSE_RCVD_SEEN: s= "Rcvd_Seen"; break;
    case RFC822PARSE_LEVEL_DOWN: s= "Level_Down"; break;
    case RFC822PARSE_LEVEL_UP: s= "Level_Up"; break;
    case RFC822PARSE_BOUNDARY: s= "Boundary"; break;
    case RFC822PARSE_LAST_BOUNDARY: s= "Last_Boundary"; break;
    case RFC822PARSE_BEGIN_HEADER: s= "Begin_Header"; break;
    case RFC822PARSE_PREAMBLE: s= "Preamble"; break;
    case RFC822PARSE_EPILOGUE: s= "Epilogue"; break;
    default: s= "[unknown event]"; break;
    }
  log_debug ("%s: ctx=%p, rfc822 event %s\n", SRCNAME, ctx, s);
}




/* This routine gets called by the RFC822 parser for all kind of
   events.  OPAQUE carries in our case a pgpmime context.  Should
   return 0 on success or -1 and as well as seeting errno on
   failure. */
static int
message_cb (void *opaque, rfc822parse_event_t event, rfc822parse_t msg)
{
  pgpmime_context_t ctx = opaque;
  HRESULT hr;

  debug_message_event (ctx, event);

  if (event == RFC822PARSE_BEGIN_HEADER || event == RFC822PARSE_T2BODY)
    {
      /* We need to check here whether to start collecting signed data
         because attachments might come without header lines and thus
         we won't see the BEGIN_HEADER event. */
      if (ctx->start_hashing == 1)
        {
          ctx->start_hashing = 2;
          ctx->hashing_level = ctx->nesting_level;
          ctx->collect_signeddata = 1;
          gpgme_data_new (&ctx->signed_data);
        }
    }


  if (event == RFC822PARSE_T2BODY)
    {
      rfc822parse_field_t field;
      const char *s1, *s2, *s3;
      size_t off;
      char *p;
      int is_text = 0;

      ctx->is_utf8 = 0;
      field = rfc822parse_parse_field (msg, "Content-Type", -1);
      if (field)
        {
          s1 = rfc822parse_query_media_type (field, &s2);
          if (s1)
            {
              log_debug ("%s: ctx=%p, media `%s' `%s'\n",
                         SRCNAME, ctx, s1, s2);

              if (!strcmp (s1, "multipart"))
                {
                  /* We don't care about the top level multipart layer
                     but wait until it comes to the actual parts which
                     then will get stored as attachments.

                     For now encapsulated signed or encrypted
                     containers are not processed in a special way as
                     they should.  Except for the simple verify
                     mode. */
                  if (ctx->verify_mode && !ctx->signed_data
                      && !strcmp (s2,"signed")
                      && (s3 = rfc822parse_query_parameter (field,
                                                            "protocol", 0))
                      && !strcmp (s3, "application/pgp-signature"))
                    {
                      /* Need to start the hashing after the next
                         boundary. */
                      ctx->start_hashing = 1;
                    }
                }
              else if (!strcmp (s1, "text"))
                {
                  is_text = 1;
                }
              else if (ctx->verify_mode
                       && ctx->nesting_level == 1
                       && !ctx->sig_data
                       && !strcmp (s1, "application")
                       && !strcmp (s2, "pgp-signature"))
                {
                  /* This is the second part of a PGP/MIME signature.
                     We only support here full messages thus checking
                     the nesting level is sufficient. We do this only
                     for the first signature (i.e. if sig_data has not
                     been set yet). We do this only while in verify
                     mode because we don't want to write a full MUA
                     (although this would be easier than to tame this
                     Outlook beast). */
                  if (!ctx->preview && !gpgme_data_new (&ctx->sig_data))
                    {
                      ctx->collect_signature = 1;
                    }
                }
              else /* Other type. */
                {
                  if (!ctx->preview)
                    ctx->collect_attachment = 1;
                }

            }

          s1 = rfc822parse_query_parameter (field, "charset", 0);
          if (s1 && !strcmp (s1, "utf-8"))
            ctx->is_utf8 = 1;

          rfc822parse_release_field (field);
        }
      else
        {
          /* No content-type at all indicates text/plain. */
          is_text = 1;
        }
      ctx->in_data = 1;

      log_debug ("%s: this body: nesting=%d part_counter=%d is_text=%d\n", 
                 SRCNAME, ctx->nesting_level, ctx->part_counter, is_text);

      /* Need to figure out the encoding. */
      ctx->is_qp_encoded = 0;
      ctx->is_base64_encoded = 0;
      p = rfc822parse_get_field (msg, "Content-Transfer-Encoding", -1, &off);
      if (p)
        {
          if (!stricmp (p+off, "quoted-printable"))
            ctx->is_qp_encoded = 1;
          else if (!stricmp (p+off, "base64"))
            {
              ctx->is_base64_encoded = 1;
              b64_init (&ctx->base64);
            }
          free (p);
        }

      /* If this is a text part, decide whether we treat it as our body. */
      if (is_text)
        {
          /* If this is the first text part at all we will
             start to collect it and use it later as the
             regular body.  An initialized ctx->BODY is an
             indication that this is not the first text part -
             we treat such a part like any other
             attachment. */
          if (!ctx->body)
            {
              if (!gpgme_data_new (&ctx->body))
                ctx->collect_body = 1;
            }
          else if (!ctx->preview)
            ctx->collect_attachment = 1;
        }


      if (ctx->collect_attachment)
        {
          /* Now that if we have an attachment prepare for writing it
             out. */
          p = NULL;
          field = rfc822parse_parse_field (msg, "Content-Disposition", -1);
          if (field)
            {
              s1 = rfc822parse_query_parameter (field, "filename", 0);
              if (s1)
                p = xstrdup (s1);
              rfc822parse_release_field (field);
            }
          if (!p)
            {
              p = xmalloc (50);
              snprintf (p, 49, "unnamed-%d.dat", ctx->part_counter);
            }
          if (ctx->outstream)
            {
              IStream_Release (ctx->outstream);
              ctx->outstream = NULL;
            }
        tryagain:
          xfree (ctx->filename);
          ctx->filename = ctx->preview? NULL:get_save_filename (ctx->hwnd, p);
          if (!ctx->filename)
            ctx->collect_attachment = 0; /* User does not want to save it. */
          else
            {
              hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                                     (STGM_CREATE | STGM_READWRITE),
                                     ctx->filename, NULL, &ctx->outstream); 
              if (FAILED (hr)) 
                {
                  log_error ("%s:%s: can't create file `%s': hr=%#lx\n",
                             SRCNAME, __func__, ctx->filename, hr); 
                  MessageBox (ctx->hwnd, _("Error creating file\n"
                                           "Please select another one"),
                              _("I/O-Error"), MB_ICONERROR|MB_OK);
                  goto tryagain;
                }
              log_debug ("%s:%s: writing attachment to `%s'\n",
                         SRCNAME, __func__, ctx->filename); 
            }
          xfree (p);
        }
    }
  else if (event == RFC822PARSE_LEVEL_DOWN)
    {
      ctx->nesting_level++;
    }
  else if (event == RFC822PARSE_LEVEL_UP)
    {
      if (ctx->nesting_level)
        ctx->nesting_level--;
      else 
        {
          log_error ("%s: ctx=%p, invalid structure: bad nesting level\n",
                     SRCNAME, ctx);
          ctx->parser_error = 1;
        }
    }
  else if (event == RFC822PARSE_BOUNDARY || event == RFC822PARSE_LAST_BOUNDARY)
    {
      ctx->in_data = 0;
      ctx->collect_body = 0;
      ctx->collect_attachment = 0;
      xfree (ctx->filename);
      ctx->filename = NULL;
      if (ctx->outstream)
        {
          IStream_Commit (ctx->outstream, 0);
          IStream_Release (ctx->outstream);
          ctx->outstream = NULL;
        }
      if (ctx->start_hashing == 2 && ctx->hashing_level == ctx->nesting_level)
        {
          ctx->start_hashing = 3; /* Avoid triggering it again. */
          ctx->collect_signeddata = 0;
        }
    }
  else if (event == RFC822PARSE_BEGIN_HEADER)
    {
      ctx->part_counter++;
    }

  return 0;
}


/* This handler is called by GPGME with the decrypted plaintext. */
static ssize_t
plaintext_handler (void *handle, const void *buffer, size_t size)
{
  pgpmime_context_t ctx = handle;
  const char *s;
  size_t nleft, pos, len;

  s = buffer;
  pos = ctx->linebufpos;
  nleft = size;
  for (; nleft ; nleft--, s++)
    {
      if (pos >= ctx->linebufsize)
        {
          log_error ("%s: ctx=%p, rfc822 parser failed: line too long\n",
                     SRCNAME, ctx);
          ctx->line_too_long = 1;
          return 0; /* Error. */
        }
      if (*s != '\n')
        ctx->linebuf[pos++] = *s;
      else
        { /* Got a complete line. Remove the last CR */
          if (pos && ctx->linebuf[pos-1] == '\r')
            pos--;

          if (rfc822parse_insert (ctx->msg, ctx->linebuf, pos))
            {
              log_error ("%s: ctx=%p, rfc822 parser failed: %s\n",
                         SRCNAME, ctx, strerror (errno));
              ctx->parser_error = 1;
              return 0; /* Error. */
            }


          if (ctx->collect_signeddata && ctx->signed_data)
            {
              /* Save the signed data.  Note that we need to delay
                 the CR/LF because the last line ending belongs to the
                 next boundary. */
              if (ctx->collect_signeddata == 2)
                gpgme_data_write (ctx->signed_data, "\r\n", 2);
              gpgme_data_write (ctx->signed_data, ctx->linebuf, pos);
              ctx->collect_signeddata = 2;
            }

          if (ctx->in_data && ctx->collect_body && ctx->body)
            {
              /* We are inside the body of the message.  Save it away
                 to a gpgme data object.  Note that this is only
                 used for the first text part. */
              if (ctx->collect_body == 1)  /* Need to skip the first line. */
                ctx->collect_body = 2;
              else
                {
                  if (ctx->is_qp_encoded)
                    len = qp_decode (ctx->linebuf, pos);
                  else if (ctx->is_base64_encoded)
                    len = b64_decode (&ctx->base64, ctx->linebuf, pos);
                  else
                    len = pos;
                  if (len)
                    {
                      if (ctx->is_utf8)
                        gpgme_data_write (ctx->body, ctx->linebuf, len);
                      else
                        latin1_data_write (ctx->body, ctx->linebuf, len);
                    }
                  if (!ctx->is_base64_encoded)
                    gpgme_data_write (ctx->body, "\r\n", 2);
                }
            }
          else if (ctx->in_data && ctx->collect_attachment)
            {
              /* We are inside of an attachment part.  Write it out. */
              if (ctx->collect_attachment == 1)  /* Skip the first line. */
                ctx->collect_attachment = 2;
              else if (ctx->outstream)
                {
                  HRESULT hr = 0;

                  if (ctx->is_qp_encoded)
                    len = qp_decode (ctx->linebuf, pos);
                  else if (ctx->is_base64_encoded)
                    len = b64_decode (&ctx->base64, ctx->linebuf, pos);
                  else
                    len = pos;
                  if (len)
                    hr = IStream_Write (ctx->outstream, ctx->linebuf,
                                        len, NULL);
                  if (!hr && !ctx->is_base64_encoded)
                    hr = IStream_Write (ctx->outstream, "\r\n", 2, NULL);
                  if (hr)
                    {
                      log_debug ("%s:%s: Write failed: hr=%#lx",
                                 SRCNAME, __func__, hr);
                      if (!ctx->preview)
                        MessageBox (ctx->hwnd, _("Error writing file"),
                                    _("I/O-Error"), MB_ICONERROR|MB_OK);
                      ctx->parser_error = 1;
                      return 0; /* Error. */
                    }
                }
            }
          else if (ctx->in_data && ctx->collect_signature)
            {
              /* We are inside of a signature attachment part.  */
              if (ctx->collect_signature == 1)  /* Skip the first line. */
                ctx->collect_signature = 2;
              else if (ctx->sig_data)
                {
                  if (ctx->is_qp_encoded)
                    len = qp_decode (ctx->linebuf, pos);
                  else if (ctx->is_base64_encoded)
                    len = b64_decode (&ctx->base64, ctx->linebuf, pos);
                  else
                    len = pos;
                  if (len)
                    gpgme_data_write (ctx->sig_data, ctx->linebuf, len);
                  if (!ctx->is_base64_encoded)
                    gpgme_data_write (ctx->sig_data, "\r\n", 2);
                }
            }
          
          /* Continue with next line. */
          pos = 0;
        }
    }
  ctx->linebufpos = pos;

  return size;
}


/* Decrypt the PGP/MIME INSTREAM (i.e the second part of the
   multipart/mixed) and allow saving of all attachments. On success a
   newly allocated body will be stored at BODY.  If ATTESTATION is not
   NULL a text with the result of the signature verification will get
   printed to it.  HWND is the window to be used for message box and
   such.  In PREVIEW_MODE no verification will be done, no messages
   saved and no messages boxes will pop up. */
int
pgpmime_decrypt (LPSTREAM instream, int ttl, char **body,
                 gpgme_data_t attestation, HWND hwnd, int preview_mode)
{
  gpg_error_t err;
  struct gpgme_data_cbs cbs;
  gpgme_data_t plaintext;
  pgpmime_context_t ctx;
  char *tmp;

  *body = NULL;

  memset (&cbs, 0, sizeof cbs);
  cbs.write = plaintext_handler;

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_errno (errno);
      log_error ("failed to open the RFC822 parser: %s", strerror (errno));
      goto leave;
    }

  err = gpgme_data_new_from_cbs (&plaintext, &cbs, ctx);
  if (err)
    goto leave;

  tmp = native_to_utf8 (_("[PGP/MIME message]"));
  err = op_decrypt_stream_to_gpgme (GPGME_PROTOCOL_OpenPGP,
                                    instream, plaintext, ttl, tmp,
                                    attestation, preview_mode);
  xfree (tmp);
  if (!err && (ctx->parser_error || ctx->line_too_long))
    err = gpg_error (GPG_ERR_GENERAL);

  if (!err)
    {
      if (ctx->body)
        {
          /* Return the buffer but first make sure it is a string. */
          if (gpgme_data_write (ctx->body, "", 1) == 1)
            {
              *body = gpgme_data_release_and_get_mem (ctx->body, NULL);
              ctx->body = NULL; 
            }
        }
      else
        {
          *body = native_to_utf8 (_("[PGP/MIME message "
                                    "without plain text body]"));
        }
    }

 leave:
  if (plaintext)
    gpgme_data_release (plaintext);
  if (ctx)
    {
      if (ctx->outstream)
        {
          IStream_Revert (ctx->outstream);
          IStream_Release (ctx->outstream);
        }
      rfc822parse_close (ctx->msg);
      if (ctx->body)
        gpgme_data_release (ctx->body);
      xfree (ctx->filename);
      xfree (ctx);
    }
  return err;
}



int
pgpmime_verify (const char *message, int ttl, char **body,
                gpgme_data_t attestation, HWND hwnd, int preview_mode)
{
  gpg_error_t err = 0;
  pgpmime_context_t ctx;
  const char *s;

  *body = NULL;

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;
  ctx->verify_mode = 1;

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_errno (errno);
      log_error ("failed to open the RFC822 parser: %s", strerror (errno));
      goto leave;
    }

  /* Need to pass the data line by line to the handler. */
  for (;(s = strchr (message, '\n')); message = s+1)
    {
      plaintext_handler (ctx, message, (s - message) + 1);
      if (ctx->parser_error || ctx->line_too_long)
        {
          err = gpg_error (GPG_ERR_GENERAL);
          break;
        }
    }

  /* Unless there is an error we should return the body. */
  if (!err)
    {
      if (ctx->body)
        {
          /* Return the buffer but first make sure it is a string. */
          if (gpgme_data_write (ctx->body, "", 1) == 1)
            {
              *body = gpgme_data_release_and_get_mem (ctx->body, NULL);
              ctx->body = NULL; 
            }
        }
      else
        {
          *body = native_to_utf8 (_("[PGP/MIME signed message without a "
                                    "plain text body]"));
        }
    }

  /* Now actually verify the signature. */
  if (!err && ctx->signed_data && ctx->sig_data)
    {
      char *tmp;

      gpgme_data_seek (ctx->signed_data, 0, SEEK_SET);
      gpgme_data_seek (ctx->sig_data, 0, SEEK_SET);
      tmp = native_to_utf8 (_("[PGP/MIME signature]"));
      err = op_verify_detached_sig_gpgme (GPGME_PROTOCOL_OpenPGP,
                                          ctx->signed_data, ctx->sig_data,
                                          tmp, attestation);
      xfree (tmp);
    }


 leave:
  if (ctx)
    {
      if (ctx->outstream)
        {
          IStream_Revert (ctx->outstream);
          IStream_Release (ctx->outstream);
        }
      rfc822parse_close (ctx->msg);
      if (ctx->body)
        gpgme_data_release (ctx->body);
      if (ctx->signed_data)
        gpgme_data_release (ctx->signed_data);
      if (ctx->sig_data)
        gpgme_data_release (ctx->sig_data);
      xfree (ctx->filename);
      xfree (ctx);
    }
  return err;
}
