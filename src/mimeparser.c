/* mimeparser.c - Parse multipart MIME message
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

/*
   Fixme: Explain how the this parser works and how it fits into the
   whole picture.
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
#include "engine.h"
#include "mapihelp.h"
#include "serpent.h"
#include "mimeparser.h"

/* Define the next to get extra debug message for the MIME parser.  */
#define DEBUG_PARSER 1

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

static const char oid_mimetag[] =
    {0x2A, 0x86, 0x48, 0x86, 0xf7, 0x14, 0x03, 0x0a, 0x04};



/* The maximum length of a line we are able to process.  RFC822 allows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000



/* To keep track of the MIME message structures we use a linked list
   with each item corresponding to one part. */
struct mimestruct_item_s;
typedef struct mimestruct_item_s *mimestruct_item_t;
struct mimestruct_item_s
{
  mimestruct_item_t next;
  unsigned int level;   /* Level in the hierarchy of that part.  0
                           indicates the outer body.  */
  char *filename;       /* Malloced fileanme or NULL.  */
  char *charset;        /* Malloced charset or NULL.  */
  char content_type[1]; /* String with the content type. */
};




/* The context object we use to track information. */
struct mime_context
{
  HWND hwnd;          /* A window handle to be used for message boxes etc. */
  rfc822parse_t msg;  /* The handle of the RFC822 parser. */

  int preview;        /* Do only decryption and pop up no  message bozes.  */
  
  int protect_mode;   /* Encrypt all attachments etc. (cf. SYMENC). */
  int verify_mode;    /* True if we want to verify a signature. */

  int nesting_level;  /* Current MIME nesting level. */
  int in_data;        /* We are currently in data (body or attachment). */
  int body_seen;      /* True if we have seen a part we consider the
                         body of the message.  */

  gpgme_data_t signed_data;/* NULL or the data object used to collect
                              the signed data. It would be better to
                              just hash it but there is no support in
                              gpgme for this yet. */
  gpgme_data_t sig_data;  /* NULL or data object to collect the
                             signature attachment which should be a
                             signature then.  */
  
  int collect_attachment; /* True if we are collecting an attachment
                             or the body. */
  int collect_signeddata; /* True if we are collecting the signed data. */
  int collect_signature;  /* True if we are collecting a signature.  */
  int start_hashing;      /* Flag used to start collecting signed data. */
  int hashing_level;      /* MIME level where we started hashing. */
  int is_qp_encoded;      /* Current part is QP encoded. */
  int is_base64_encoded;  /* Current part is base 64 encoded. */
  int is_utf8;            /* Current part has charset utf-8. */
  protocol_t protocol;    /* The detected crypto protocol.  */

  int part_counter;       /* Counts the number of processed parts. */
  int any_boundary;       /* Indicates whether we have seen any
                             boundary which means that we are actually
                             working on a MIME message and not just on
                             plain rfc822 message.  */
  
  /* A linked list describing the structure of the mime message.  This
     list gets build up while parsing the message.  */
  mimestruct_item_t mimestruct;
  mimestruct_item_t *mimestruct_tail;
  mimestruct_item_t mimestruct_cur;

  LPMESSAGE mapi_message; /* The MAPI message object we are working on.  */
  LPSTREAM outstream;     /* NULL or a stream to write a part to. */
  LPATTACH mapi_attach;   /* The attachment object we are writing.  */
  symenc_t symenc;        /* NULL or the context used to protect
                             attachments. */
  int any_attachments_created;  /* True if we created a new atatchment.  */

  b64_state_t base64;     /* The state of the Base-64 decoder.  */

  int line_too_long;  /* Indicates that a received line was too long. */
  int parser_error;   /* Indicates that we encountered a error from
                         the parser. */

  /* Buffer used to constructed complete files. */
  size_t linebufsize;   /* The allocated size of the buffer. */
  size_t linebufpos;    /* The actual write posituion. */  
  char linebuf[1];      /* The buffer. */
};
typedef struct mime_context *mime_context_t;


/* This function is a wrapper around gpgme_data_write to convert the
   data to utf-8 first.  We assume Latin-1 here. */
/* static int */
/* latin1_data_write (gpgme_data_t data, const char *line, size_t len) */
/* { */
/*   const char *s; */
/*   char *buffer, *p; */
/*   size_t i, n; */
/*   int rc; */

/*   for (s=line, i=0, n=0 ; i < len; s++, i++ )  */
/*     { */
/*       n++; */
/*       if (*s & 0x80) */
/*         n++; */
/*     } */
/*   buffer = xmalloc (n + 1); */
/*   for (s=line, i=0, p=buffer; i < len; s++, i++ ) */
/*     { */
/*       if (*s & 0x80) */
/*         { */
/*           *p++ = 0xc0 | ((*s >> 6) & 3); */
/*           *p++ = 0x80 | (*s & 0x3f); */
/*         } */
/*       else */
/*         *p++ = *s; */
/*     } */
/*   assert (p-buffer == n); */
/*   rc = gpgme_data_write (data, buffer, n); */
/*   xfree (buffer); */
/*   return rc; */
/* } */


/* Print the message event EVENT. */
static void
debug_message_event (mime_context_t ctx, rfc822parse_event_t event)
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
#ifdef DEBUG_PARSER
  log_debug ("%s: ctx=%p, rfc822 event %s\n", SRCNAME, ctx, s);
#endif
}



/* Start a new atatchment.  With IS_BODY set, the attachment is
   actually the body part of the message which is treated in a special
   way. */
static int
start_attachment (mime_context_t ctx, int is_body)
{
  int retval = -1;
  HRESULT hr;
  ULONG newpos;
  SPropValue prop;
  LPATTACH newatt = NULL;
  LPSTREAM to = NULL;
  LPUNKNOWN punk;

#ifdef DEBUG_PARSER
  log_debug ("%s:%s: for ctx=%p is_body=%d", SRCNAME, __func__, ctx, is_body);
#endif

  /* Just in case something has not been finished, do it here. */
  if (ctx->outstream)
    {
      IStream_Release (ctx->outstream);
      ctx->outstream = NULL;
    }
  if (ctx->mapi_attach)
    {
      IAttach_Release (ctx->mapi_attach);
      ctx->mapi_attach = NULL;
    }
  if (ctx->symenc)
    {
      symenc_close (ctx->symenc);
      ctx->symenc = NULL;
    }
  
  /* Before we start with the first attachment we need to delete all
     attachments which might have been created already by a past
     parser run.  */
  if (!ctx->any_attachments_created)
    {
      mapi_attach_item_t *table;
      int i;
      
      table = mapi_create_attach_table (ctx->mapi_message, 1);
      if (table)
        {
          for (i=0; !table[i].end_of_table; i++)
            if (table[i].attach_type == ATTACHTYPE_FROMMOSS)
              {
                hr = IMessage_DeleteAttach (ctx->mapi_message, 
                                            table[i].mapipos,
                                            0, NULL, 0);
                if (hr)
                  log_error ("%s:%s: DeleteAttach failed: hr=%#lx\n",
                             SRCNAME, __func__, hr); 
              }
          mapi_release_attach_table (table);
        }
      ctx->any_attachments_created = 1;
    }
  
  /* Now create a new attachment.  */
  hr = IMessage_CreateAttach (ctx->mapi_message, NULL, 0, &newpos, &newatt);
  if (hr)
    {
      log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }

  prop.ulPropTag = PR_ATTACH_METHOD;
  prop.Value.ul = ATTACH_BY_VALUE;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr != S_OK)
    {
      log_error ("%s:%s: can't set attach method: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }

  /* Mark that attachment so that we know why it has been created.  */
  if (get_gpgolattachtype_tag (ctx->mapi_message, &prop.ulPropTag) )
    goto leave;
  prop.Value.l = ATTACHTYPE_FROMMOSS;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);	
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      goto leave;
    }

  /* The body attachment is special and should not be shown in the list
     of attachments.  */
  if (is_body)
    {
      prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
      prop.Value.b = TRUE;
      hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
      if (hr)
        {
          log_error ("%s:%s: can't set hidden attach flag: hr=%#lx\n",
                     SRCNAME, __func__, hr); 
          goto leave;
        }
    }
  

  /* We need to insert a short filename .  Without it, the _displayed_
     list of attachments won't get updated although the attachment has
     been created. */
  prop.ulPropTag = PR_ATTACH_FILENAME_A;
  {
    char buf[100];

    if (is_body)
      prop.Value.lpszA = is_body == 2? "gpgol000.htm":"gpgol000.txt";
    else
      {
        snprintf (buf, 100, "gpgol%03d.dat", ctx->part_counter);
        prop.Value.lpszA = buf;
      }
    hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  }
  if (hr)
    {
      log_error ("%s:%s: can't set attach filename: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }

  /* And now for the real name.  */
  if (ctx->mimestruct_cur && ctx->mimestruct_cur->filename)
    {
      prop.ulPropTag = PR_ATTACH_LONG_FILENAME_A;
      prop.Value.lpszA = ctx->mimestruct_cur->filename;
      hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
      if (hr)
        {
          log_error ("%s:%s: can't set attach long filename: hr=%#lx\n",
                     SRCNAME, __func__, hr); 
          goto leave;
        }
    }

  prop.ulPropTag = PR_ATTACH_TAG;
  prop.Value.bin.cb  = sizeof oid_mimetag;
  prop.Value.bin.lpb = (LPBYTE)oid_mimetag;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach tag: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }

  assert (ctx->mimestruct_cur && ctx->mimestruct_cur->content_type);
  prop.ulPropTag = PR_ATTACH_MIME_TAG_A;
  prop.Value.lpszA = ctx->mimestruct_cur->content_type;
  hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach mime tag: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }


  /* If we are in protect mode (i.e. working on a decrypted message,
     we need to setup the symkey context to protect (encrypt) the
     attachment in the MAPI.  */
  if (ctx->protect_mode)
    {
      char *iv;

      if (get_gpgolprotectiv_tag (ctx->mapi_message, &prop.ulPropTag) )
        goto leave;

      iv = create_initialization_vector (16);
      if (!iv)
        {
          log_error ("%s:%s: error creating initialization vector",
                     SRCNAME, __func__);
          goto leave;
        }
      prop.Value.bin.cb = 16;
      prop.Value.bin.lpb = iv;
      hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);	
      if (hr)
        {
          log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                     SRCNAME, __func__, "GpgOL Protect IV", hr); 
          goto leave;
        }

      ctx->symenc = symenc_open (get_128bit_session_key (), 16, iv, 16);
      xfree (iv);
      if (!ctx->symenc)
        {
          log_error ("%s:%s: error creating cipher context",
                     SRCNAME, __func__);
          goto leave;
        }
    }

 
  punk = (LPUNKNOWN)to;
  hr = IAttach_OpenProperty (newatt, PR_ATTACH_DATA_BIN, &IID_IStream, 0,
                             MAPI_CREATE|MAPI_MODIFY, &punk);
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't create output stream: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto leave;
    }
  to = (LPSTREAM)punk;
  
  ctx->outstream = to;
  to = NULL;
  ctx->mapi_attach = newatt;
  newatt = NULL;

  if (ctx->symenc)
    {
      char tmpbuf[16];
      /* Write an encrypted fixed 16 byte string which we need to
         check at decryption time to see whether we have actually
         encrypted it using this session key.  */
      symenc_cfb_encrypt (ctx->symenc, tmpbuf, "GpgOL attachment", 16);
      IStream_Write (ctx->outstream, tmpbuf, 16, NULL);
    }
  retval = 0; /* Success.  */

 leave:
  if (to)
    {
      IStream_Revert (to);
      IStream_Release (to);
    }
  if (newatt)
    IAttach_Release (newatt);
  return retval;
}


static int
finish_attachment (mime_context_t ctx, int cancel)
{
  HRESULT hr;
  int retval = -1;

#ifdef DEBUG_PARSER
  log_debug ("%s:%s: for ctx=%p cancel=%d", SRCNAME, __func__, ctx, cancel);
#endif

  if (ctx->outstream)
    {
      IStream_Commit (ctx->outstream, 0);
      IStream_Release (ctx->outstream);
      ctx->outstream = NULL;

      if (cancel)
        retval = 0;
      else if (ctx->mapi_attach)
        {
          hr = IAttach_SaveChanges (ctx->mapi_attach, 0);
          if (hr)
            {
              log_error ("%s:%s: SaveChanges(attachment) failed: hr=%#lx\n",
                         SRCNAME, __func__, hr); 
            }
          else
            retval = 0; /* Success.  */
        }
    }
  if (ctx->mapi_attach)
    {
      IAttach_Release (ctx->mapi_attach);
      ctx->mapi_attach = NULL;
    }
  if (ctx->symenc)
    {
      symenc_close (ctx->symenc);
      ctx->symenc = NULL;
    }
  return retval;
}


static int
finish_message (LPMESSAGE message)
{
  return 0;
}



/* Process the transition to body event. 

   This means we have received the empty line indicating the body and
   should now check the headers to see what to do about this part.  */
static int
t2body (mime_context_t ctx, rfc822parse_t msg)
{
  rfc822parse_field_t field;
  const char *ctmain, *ctsub;
  const char *s;
  size_t off;
  char *p;
  int is_text = 0;
  int is_body = 0;
  char *filename = NULL; 
  char *charset = NULL;
        

  /* Figure out the encoding.  */
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

  /* Get the filename from the header.  */
  field = rfc822parse_parse_field (msg, "Content-Disposition", -1);
  if (field)
    {
      s = rfc822parse_query_parameter (field, "filename", 0);
      if (s)
        filename = xstrdup (s);
      rfc822parse_release_field (field);
    }

  /* Process the Content-type and all its parameters.  */
  ctmain = ctsub = NULL;
  field = rfc822parse_parse_field (msg, "GnuPG-Content-Type", -1);
  if (!field)
    field = rfc822parse_parse_field (msg, "Content-Type", -1);
  if (field)
    ctmain = rfc822parse_query_media_type (field, &ctsub);
  if (!ctmain)
    {
      /* Either there is no content type field or it is faulty; in
         both cases we fall back to text/plain.  */
      ctmain = "text";
      ctsub  = "plain";
    }

#ifdef DEBUG_PARSER  
  log_debug ("%s:%s: ctx=%p, ct=`%s/%s'\n",
             SRCNAME, __func__, ctx, ctmain, ctsub);
#endif

  /* We only support UTF-8 for now.  Check here.  */
  s = rfc822parse_query_parameter (field, "charset", 0);
  if (s)
    charset = xstrdup (s);
  ctx->is_utf8 = (s && !strcmp (s, "utf-8"));

  /* Update our idea of the entire MIME structure.  */
  {
    mimestruct_item_t ms;

    ms = xmalloc (sizeof *ms + strlen (ctmain) + 1 + strlen (ctsub));
    ctx->mimestruct_cur = ms;
    *ctx->mimestruct_tail = ms;
    ctx->mimestruct_tail = &ms->next;
    ms->next = NULL;
    strcpy (stpcpy (stpcpy (ms->content_type, ctmain), "/"), ctsub);
    ms->level = ctx->nesting_level;
    ms->filename = filename;
    filename = NULL;
    ms->charset = charset;
    charset = NULL;

  }

      
  if (!strcmp (ctmain, "multipart"))
    {
      /* We don't care about the top level multipart layer but wait
         until it comes to the actual parts which then will get stored
         as attachments.

         For now encapsulated signed or encrypted containers are not
         processed in a special way as they should.  Except for the
         simple verify mode. */
      if (ctx->verify_mode && !ctx->signed_data
          && !strcmp (ctsub, "signed")
          && (s = rfc822parse_query_parameter (field, "protocol", 0)))
        {
          if (!strcmp (s, "application/pgp-signature"))
            ctx->protocol = PROTOCOL_OPENPGP;
          else if (!strcmp (s, "application/pkcs7-signature")
                   || !strcmp (s, "application/x-pkcs7-signature"))
            ctx->protocol = PROTOCOL_SMIME;
          else
            ctx->protocol = PROTOCOL_UNKNOWN;

          /* Need to start the hashing after the next boundary. */
          ctx->start_hashing = 1;
        }
    }
  else if (!strcmp (ctmain, "text"))
    {
      is_text = !strcmp (ctsub, "html")? 2:1;
    }
  else if (ctx->verify_mode && ctx->nesting_level == 1 && !ctx->sig_data
           && !strcmp (ctmain, "application")
           && ((ctx->protocol == PROTOCOL_OPENPGP   
                && !strcmp (ctsub, "pgp-signature"))
               || (ctx->protocol == PROTOCOL_SMIME   
                   && (!strcmp (ctsub, "pkcs7-signature")
                       || !strcmp (ctsub, "x-pkcs7-signature")))))
    {
      /* This is the second part of a MOSS signature.  We only support
         here full messages thus checking the nesting level is
         sufficient.  We do this only for the first signature (i.e. if
         sig_data has not been set yet).  We also do this only while
         in verify mode because we don't want to write a full MUA.  */
      if (!ctx->preview && !gpgme_data_new (&ctx->sig_data))
        ctx->collect_signature = 1;
    }
  else /* Other type. */
    {
      if (!ctx->preview)
        ctx->collect_attachment = 1;
    }
  rfc822parse_release_field (field); /* (Content-type) */
  ctx->in_data = 1;

#ifdef DEBUG_PARSER
  log_debug ("%s: this body: nesting=%d part_counter=%d is_text=%d\n",
             SRCNAME, ctx->nesting_level, ctx->part_counter, is_text);
#endif

  /* If this is a text part, decide whether we treat it as our body. */
  if (is_text)
    {
      /* If this is the first text part at all we will start to
         collect it and use it later as the regular body.  */
      if (!ctx->body_seen)
        {
          ctx->body_seen = 1;
          ctx->collect_attachment = 1;
          is_body = 1;
        }
      else if (!ctx->preview)
        ctx->collect_attachment = 1;
    }


  if (ctx->collect_attachment)
    {
      /* Now that if we have an attachment prepare a new MAPI
         attachment.  */
      if (start_attachment (ctx, is_body))
        return -1;
      assert (ctx->outstream);
    }

  return 0;
}


/* This routine gets called by the RFC822 parser for all kind of
   events.  OPAQUE carries in our case an smime context.  Should
   return 0 on success or -1 as well as setting errno on
   failure.  */
static int
message_cb (void *opaque, rfc822parse_event_t event, rfc822parse_t msg)
{
  int retval = 0;
  mime_context_t ctx = opaque;

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


  switch (event)
    {
    case RFC822PARSE_T2BODY:
      retval = t2body (ctx, msg);
      break;

    case RFC822PARSE_LEVEL_DOWN:
      ctx->nesting_level++;
      break;

    case RFC822PARSE_LEVEL_UP:
      if (ctx->nesting_level)
        ctx->nesting_level--;
      else 
        {
          log_error ("%s: ctx=%p, invalid structure: bad nesting level\n",
                     SRCNAME, ctx);
          ctx->parser_error = 1;
        }
      break;
    
    case RFC822PARSE_BOUNDARY:
    case RFC822PARSE_LAST_BOUNDARY:
      ctx->any_boundary = 1;
      ctx->in_data = 0;
      ctx->collect_attachment = 0;
      
      finish_attachment (ctx, 0);
      assert (!ctx->outstream);

      if (ctx->start_hashing == 2 && ctx->hashing_level == ctx->nesting_level)
        {
          ctx->start_hashing = 3; /* Avoid triggering it again. */
          ctx->collect_signeddata = 0;
        }
      break;
    
    case RFC822PARSE_BEGIN_HEADER:
      ctx->part_counter++;
      break;

    default:  /* Ignore all other events. */
      break;
    }

  return retval;
}



/* This handler is called by GPGME with the decrypted plaintext. */
static int
plaintext_handler (void *handle, const void *buffer, size_t size)
{
  mime_context_t ctx = handle;
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
          return -1; /* Error. */
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
              return -1; /* Error. */
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

          if (ctx->in_data && ctx->collect_attachment)
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
                    {
                      if (ctx->symenc)
                        symenc_cfb_encrypt (ctx->symenc, ctx->linebuf,
                                            ctx->linebuf, len);
                      hr = IStream_Write (ctx->outstream, ctx->linebuf,
                                          len, NULL);
                    }
                  if (!hr && !ctx->is_base64_encoded)
                    {
                      char tmp[3] = "\r\n";

                      if (ctx->symenc)
                        symenc_cfb_encrypt (ctx->symenc, tmp, tmp, 2);
                      hr = IStream_Write (ctx->outstream, tmp, 2, NULL);
                    }
                  if (hr)
                    {
                      log_debug ("%s:%s: Write failed: hr=%#lx",
                                 SRCNAME, __func__, hr);
                      if (!ctx->preview)
                        MessageBox (ctx->hwnd, _("Error writing to stream"),
                                    _("I/O-Error"), MB_ICONERROR|MB_OK);
                      ctx->parser_error = 1;
                      return -1; /* Error. */
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



static void 
show_mimestruct (mimestruct_item_t mimestruct)
{
  mimestruct_item_t ms;

  for (ms = mimestruct; ms; ms = ms->next)
    log_debug ("MIMESTRUCT: %*s%s  cs=%s  fn=%s\n",
               ms->level*2, "", ms->content_type,
               ms->charset? ms->charset : "[none]",
               ms->filename? ms->filename : "[none]");
}




int
mime_verify (protocol_t protocol, const char *message, size_t messagelen, 
             LPMESSAGE mapi_message, HWND hwnd, int preview_mode)
{
  gpg_error_t err = 0;
  mime_context_t ctx;
  const char *s;
  size_t len;
  char *signature = NULL;
  engine_filter_t filter = NULL;

  /* Note: PROTOCOL is not used here but figured out directly while
     collecting the message.  Eventually it might help use setup a
     proper verification context right at startup to avoid collecting
     all the stuff.  However there are a couple of problems with that
     - for example we don't know whether gpgsm behaves correctly by
     first reading all the data and only the reading the signature.  I
     guess it is the case but that needs to be checked first.  It is
     just a performance issue.  */
 
  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;
  ctx->verify_mode = 1;
  ctx->mapi_message = mapi_message;
  ctx->mimestruct_tail = &ctx->mimestruct;

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_syserror ();
      log_error ("%s:%s: failed to open the RFC822 parser: %s", 
                 SRCNAME, __func__, gpg_strerror (err));
      goto leave;
    }

  /* Need to pass the data line by line to the handler. */
  while ( (s = memchr (message, '\n', messagelen)) )
    {
      len = s - message + 1;
      log_debug ("passing '%.*s'\n", (int)len, message);
      plaintext_handler (ctx, message, len);
      if (ctx->parser_error || ctx->line_too_long)
        {
          err = gpg_error (GPG_ERR_GENERAL);
          break;
        }
      message += len;
      assert (messagelen >= len);
      messagelen -= len;
    }
  /* Note: the last character should be a LF, if not we ignore such an
     incomplete last line.  */
  if (ctx->sig_data && gpgme_data_write (ctx->sig_data, "", 1) == 1)
    {
      signature = gpgme_data_release_and_get_mem (ctx->sig_data, NULL);
      ctx->sig_data = NULL; 
    }

  /* Now actually verify the signature. */
  if (!err && ctx->signed_data && signature)
    {
      gpgme_data_seek (ctx->signed_data, 0, SEEK_SET);
      
      if ((err=engine_create_filter (&filter, NULL, NULL)))
        goto leave;
      if ((err=engine_verify_start (filter, signature, ctx->protocol)))
        goto leave;

      /* Filter the data.  */
      do
        {
          int nread;
          char buffer[4096];
          
          nread = gpgme_data_read (ctx->signed_data, buffer, sizeof buffer);
          if (nread < 0)
            {
              err = gpg_error_from_syserror ();
              log_error ("%s:%s: gpgme_data_read failed: %s", 
                         SRCNAME, __func__, gpg_strerror (err));
            }
          else if (nread)
            {
              TRACEPOINT();
              err = engine_filter (filter, buffer, nread);
            }
          else
            break; /* EOF */
        }
      while (!err);
      if (err)
        goto leave;
      
      /* Wait for the engine to finish.  */
      if ((err = engine_filter (filter, NULL, 0)))
        goto leave;
      if ((err = engine_wait (filter)))
        goto leave;
      filter = NULL;
    }


 leave:
  TRACEPOINT();
  gpgme_free (signature);
  engine_cancel (filter);
  if (ctx)
    {
      /* Cancel any left open attachment.  */
      finish_attachment (ctx, 1); 
      rfc822parse_close (ctx->msg);
      gpgme_data_release (ctx->signed_data);
      gpgme_data_release (ctx->sig_data);
      show_mimestruct (ctx->mimestruct);
      while (ctx->mimestruct)
        {
          mimestruct_item_t tmp = ctx->mimestruct->next;
          xfree (ctx->mimestruct->filename);
          xfree (ctx->mimestruct->charset);
          xfree (ctx->mimestruct);
          ctx->mimestruct = tmp;
        }
      symenc_close (ctx->symenc);
      xfree (ctx);
      if (!err)
        finish_message (mapi_message);
    }
  return err;
}



/* Decrypt the PGP or S/MIME message taken from INSTREAM.  HWND is the
   window to be used for message box and such.  In PREVIEW_MODE no
   verification will be done, no messages saved and no messages boxes
   will pop up. */
int
mime_decrypt (protocol_t protocol, LPSTREAM instream, LPMESSAGE mapi_message,
              HWND hwnd, int preview_mode)
{
  gpg_error_t err;
  mime_context_t ctx;
  engine_filter_t filter = NULL;

  log_debug ("%s:%s: enter (protocol=%d)", SRCNAME, __func__, protocol);

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->protect_mode = 1; 
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;
  ctx->mapi_message = mapi_message;
  ctx->mimestruct_tail = &ctx->mimestruct;

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_syserror ();
      log_error ("%s:%s: failed to open the RFC822 parser: %s", 
                 SRCNAME, __func__, gpg_strerror (err));
      goto leave;
    }

  /* Prepare the decryption.  */
/*       title = native_to_utf8 (_("[Encrypted S/MIME message]")); */
/*       title = native_to_utf8 (_("[Encrypted PGP/MIME message]")); */
  if ((err=engine_create_filter (&filter, plaintext_handler, ctx)))
    goto leave;
  if ((err=engine_decrypt_start (filter, protocol, !preview_mode)))
    goto leave;

  
  /* Filter the stream.  */
  do
    {
      HRESULT hr;
      ULONG nread;
      char buffer[4096];
      
      /* For EOF detection we assume that Read returns no error and
         thus nread will be 0.  The specs say that "Depending on the
         implementation, either S_FALSE or an error code could be
         returned when reading past the end of the stream"; thus we
         are not really sure whether our assumption is correct.  At
         another place the documentation says that the implementation
         used by ISequentialStream exhibits the same EOF behaviour has
         found on the MSDOS FAT file system.  So we seem to have good
         karma. */
      hr = IStream_Read (instream, buffer, sizeof buffer, &nread);
      if (hr)
        {
          log_error ("%s:%s: IStream::Read failed: hr=%#lx", 
                     SRCNAME, __func__, hr);
          err = gpg_error (GPG_ERR_EIO);
        }
      else if (nread)
        {
          err = engine_filter (filter, buffer, nread);
        }
      else
        break; /* EOF */
    }
  while (!err);
  if (err)
    goto leave;

  /* Wait for the engine to finish.  */
  if ((err = engine_filter (filter, NULL, 0)))
    goto leave;
  if ((err = engine_wait (filter)))
    goto leave;
  filter = NULL;

  if (ctx->parser_error || ctx->line_too_long)
    err = gpg_error (GPG_ERR_GENERAL);

 leave:
  engine_cancel (filter);
  if (ctx)
    {
      /* Cancel any left over attachment which means that the MIME
         structure was not complete.  However if we have not seen any
         boundary the message is a non-MIME one but we may have
         started the body attachment (gpgol000.txt) - this one needs
         to be finished properly.  */
      finish_attachment (ctx, ctx->any_boundary? 1: 0);
      rfc822parse_close (ctx->msg);
      if (ctx->signed_data)
        gpgme_data_release (ctx->signed_data);
      if (ctx->sig_data)
        gpgme_data_release (ctx->sig_data);
      show_mimestruct (ctx->mimestruct);
      while (ctx->mimestruct)
        {
          mimestruct_item_t tmp = ctx->mimestruct->next;
          xfree (ctx->mimestruct->filename);
          xfree (ctx->mimestruct->charset);
          xfree (ctx->mimestruct);
          ctx->mimestruct = tmp;
        }
      symenc_close (ctx->symenc);
      xfree (ctx);
      if (!err)
        finish_message (mapi_message);
    }
  return err;
}

