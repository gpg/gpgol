/* pgpmime.c - Try to handle PGP/MIME for Outlook
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
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

#include <windows.h>

#include <gpgme.h>

#include "mymapi.h"
#include "mymapitags.h"

#include "rfc822parse.h"
#include "util.h"
#include "pgpmime.h"
#include "engine.h"


/* The maximum length of a line we ar able to porcess.  RFC822 alows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000

/* The context object we use to track information. */
struct pgpmime_context
{
  rfc822parse_t msg;  /* The handle of the RFC822 parser. */

  int nesting_level;  /* Current MIME nesting level. */
  int in_data;        /* We are currently in data (body or attachment). */

  
  gpgme_data_t body;  /* NULL or a data object used to collect the
                         body part we are going to display later. */
  int collect_body;   /* True if we are collecting the body lines. */
  int is_qp_encoded;  /* Current part is QP encoded. */

  int line_too_long;  /* Indicates that a received line was too long. */
  int parser_error;   /* Indicates that we encountered a error from
                         the parser. */

  /* Buffer used to constructed complete files. */
  size_t linebufsize;   /* The allocated size of the buffer. */
  size_t linebufpos;    /* The actual write posituion. */  
  char linebuf[1];      /* The buffer. */
};
typedef struct pgpmime_context *pgpmime_context_t;


/* Do in-place decoding of quoted-printable data of LENGTH in BUFFER.
   Returns the new length of the buffer. */
static size_t
qp_decode (char *buffer, size_t length)
{
  char *d, *s;

  for (s=d=buffer; length; length--)
    if (*s == '=' && length > 2 && hexdigitp (s+1) && hexdigitp (s+2))
      {
        s++;
        *(unsigned char*)d++ = xtoi_2 (s);
        s += 2;
        length -= 2;
      }
    else
      *d++ = *s++;
  
  return d - buffer;
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
  log_debug ("%s: ctx=%p, rfc822 event %s\n", __FILE__, ctx, s);
}




/* This routine gets called by the RFC822 parser for all kind of
   events.  OPAQUE carries in our case a pgpmime context.  Should
   return 0 on success or -1 and as well as seeting errno on
   failure. */
static int
message_cb (void *opaque, rfc822parse_event_t event, rfc822parse_t msg)
{
  pgpmime_context_t ctx = opaque;

  debug_message_event (ctx, event);
  if (event == RFC822PARSE_T2BODY)
    {
      rfc822parse_field_t field;
      const char *s1, *s2;
      size_t off;
      char *p;

      field = rfc822parse_parse_field (msg, "Content-Type", -1);
      if (field)
        {
          s1 = rfc822parse_query_media_type (field, &s2);
          if (s1)
            {
              log_debug ("%s: ctx=%p, media `%s' `%s'\n",
                         __FILE__, ctx, s1, s2);

              if (!strcmp (s1, "multipart"))
                {
                  if (!strcmp (s2, "signed"))
                    ;
                  else if (!strcmp (s2, "encrypted"))
                    ;
                }
              else if (!strcmp (s1, "text"))
                {
                  if (!ctx->body)
                    {
                      if (!gpgme_data_new (&ctx->body))
                        ctx->collect_body = 1;
                    }
                }
            }
          
          rfc822parse_release_field (field);
        }
      ctx->in_data = 1;

      ctx->is_qp_encoded = 0;
      p = rfc822parse_get_field (msg, "Content-Transfer-Encoding", -1, &off);
      if (p)
        {
          if (!stricmp (p+off, "quoted-printable"))
            ctx->is_qp_encoded = 1;
          free (p);
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
                     __FILE__, ctx);
          ctx->parser_error = 1;
        }
    }
  else if (event == RFC822PARSE_BOUNDARY || event == RFC822PARSE_LAST_BOUNDARY)
    {
      ctx->in_data = 0;
      ctx->collect_body = 0;
    }
  else if (event == RFC822PARSE_BEGIN_HEADER)
    {
      
    }

  return 0;
}


/* This handler is called by GPGME with the decrypted plaintext. */
static ssize_t
plaintext_handler (void *handle, const void *buffer, size_t size)
{
  pgpmime_context_t ctx = handle;
  const char *s;
  size_t n, pos;

  s = buffer;
  pos = ctx->linebufpos;
  n = size;
  for (; n ; n--, s++)
    {
      if (pos >= ctx->linebufsize)
        {
          log_error ("%s: ctx=%p, rfc822 parser failed: line too long\n",
                     __FILE__, ctx);
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
                         __FILE__, ctx, strerror (errno));
              ctx->parser_error = 1;
              return 0; /* Error. */
            }

          if (ctx->in_data && ctx->collect_body && ctx->body)
            {
              if (ctx->collect_body == 1)
                ctx->collect_body = 2;
              else
                {
                  if (ctx->is_qp_encoded)
                    pos = qp_decode (ctx->linebuf, pos);
                  gpgme_data_write (ctx->body, ctx->linebuf, pos);
                  gpgme_data_write (ctx->body, "\r\n", 2);
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
   newly allocated body will be stored at BODY. */
int
pgpmime_decrypt (LPSTREAM instream, int ttl, char **body)
{
  gpg_error_t err;
  struct gpgme_data_cbs cbs;
  gpgme_data_t plaintext;
  pgpmime_context_t ctx;

  *body = NULL;

  memset (&cbs, 0, sizeof cbs);
  cbs.write = plaintext_handler;

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;

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

  err = op_decrypt_stream_to_gpgme (instream, plaintext, ttl, NULL);
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
        *body = xstrdup ("[PGP/MIME message without plain text body]");
    }

 leave:
  rfc822parse_close (ctx->msg);
  if (plaintext)
    gpgme_data_release (plaintext);
  if (ctx && ctx->body)
    gpgme_data_release (ctx->body);
  xfree (ctx);
  return err;
}
