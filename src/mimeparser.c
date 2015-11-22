/* mimeparser.c - Parse multipart MIME message
 *	Copyright (C) 2005, 2007, 2008, 2009 g10 Code GmbH
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
#include "rfc2047parse.h"
#include "common.h"
#include "engine.h"
#include "mapihelp.h"
#include "serpent.h"
#include "mimeparser.h"
#include "parsetlv.h"


#define debug_mime_parser (opt.enable_debug & (DBG_MIME_PARSER|DBG_MIME_DATA))
#define debug_mime_data (opt.enable_debug & DBG_MIME_DATA)


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
  char *filename;       /* Malloced filename or NULL.  */
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
  int no_mail_header; /* True if we want to bypass all MIME parsing.  */

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
  int is_body;            /* The current part belongs to the body.  */
  int is_opaque_signed;   /* Flag indicating opaque signed S/MIME. */
  int may_be_opaque_signed;/* Hack, see code.  */
  protocol_t protocol;    /* The detected crypto protocol.  */

  int part_counter;       /* Counts the number of processed parts. */
  int any_boundary;       /* Indicates whether we have seen any
                             boundary which means that we are actually
                             working on a MIME message and not just on
                             plain rfc822 message.  */
  
  engine_filter_t outfilter; /* Filter as used by ciphertext_handler.  */

  /* A linked list describing the structure of the mime message.  This
     list gets build up while parsing the message.  */
  mimestruct_item_t mimestruct;
  mimestruct_item_t *mimestruct_tail;
  mimestruct_item_t mimestruct_cur;

  LPMESSAGE mapi_message; /* The MAPI message object we are working on.  */
  LPSTREAM outstream;     /* NULL or a stream to write a part to. */
  LPATTACH mapi_attach;   /* The attachment object we are writing.  */
  symenc_t symenc;        /* NULL or the context used to protect
                             an attachment. */
  struct {
    LPSTREAM outstream;   /* Saved stream used to continue a body
                             part. */
    LPATTACH mapi_attach; /* Saved attachment used to continue a body part.  */
    symenc_t symenc;      /* Saved encryption context used to continue
                             a body part.  */
  } body_saved;
  int any_attachments_created;  /* True if we created a new atatchment.  */

  b64_state_t base64;     /* The state of the Base-64 decoder.  */

  int line_too_long;  /* Indicates that a received line was too long. */
  gpg_error_t parser_error;   /* Indicates that we encountered a error from
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
  if (debug_mime_parser)
    log_debug ("%s: ctx=%p, rfc822 event %s\n", SRCNAME, ctx, s);
}


/* Returns true if the BER encoded data in BUFFER is CMS signed data.
   LENGTH gives the length of the buffer, for correct detection LENGTH
   should be at least about 24 bytes.  */
static int
is_cms_signed_data (const char *buffer, size_t length)
{
  const char *p = buffer;
  size_t n = length;
  tlvinfo_t ti;
          
  if (parse_tlv (&p, &n, &ti))
    return 0;
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_SEQUENCE
        && ti.is_cons) )
    return 0;
  if (parse_tlv (&p, &n, &ti))
    return 0;
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_OBJECT_ID
        && !ti.is_cons && ti.length) || ti.length > n)
    return 0;
  if (ti.length == 9 && !memcmp (p, "\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02", 9))
    return 1;
  return 0;
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

  if (debug_mime_parser)
    log_debug ("%s:%s: for ctx=%p is_body=%d", SRCNAME, __func__, ctx,is_body);

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

  /* The body attachment is special and should not be shown in the
     list of attachments.  If the option body-as-attachment is used
     and the message is protected we do set the hidden flag to
     false.  */
  if (is_body)
    {
      prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
      prop.Value.b = (ctx->protect_mode && opt.body_as_attachment)? FALSE:TRUE;
      hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
      if (hr)
        {
          log_error ("%s:%s: can't set hidden attach flag: hr=%#lx\n",
                     SRCNAME, __func__, hr); 
          goto leave;
        }
    }
  ctx->is_body = is_body;

  /* We need to insert a short filename .  Without it, the _displayed_
     list of attachments won't get updated although the attachment has
     been created.  If we know the content type we use an appropriate
     suffix for the filename.  This is useful so that if no filename
     is known for the attachment (to be stored in
     PR_ATTACH_LONG_FILENAME), Outlooks gets an idea about the content
     of the attachment from this made up filename.  This allows for
     example to click on the attachment and open it with an
     appropriate application.  */
  prop.ulPropTag = PR_ATTACH_FILENAME_A;
  {
    char buf[100];

    if (is_body)
      prop.Value.lpszA = is_body == 2? "gpgol000.htm":"gpgol000.txt";
    else
      {
        static struct {
          const char *suffix;
          const char *ct;
        } suffix_table[] = {
          { "doc", "application/msword" },
          { "eml", "message/rfc822" },
          { "htm", "text/html" },
          { "jpg", "image/jpeg" },
          { "pdf", "application/pdf" },
          { "png", "image/png" },
          { "pps", "application/vnd.ms-powerpoint" },
          { "ppt", "application/vnd.ms-powerpoint" },
          { "ps",  "application/postscript" },
          { NULL, NULL }
        };
        const char *suffix = "dat";  /* Default.  */
        int idx;
        
        if (ctx->mimestruct_cur && ctx->mimestruct_cur->content_type)
          {
            for (idx=0; suffix_table[idx].ct; idx++)
              if (!strcmp (ctx->mimestruct_cur->content_type,
                           suffix_table[idx].ct))
                {
                  suffix = suffix_table[idx].suffix;
                  break;
                }
          }

        snprintf (buf, sizeof buf, "gpgol%03d.%s", ctx->part_counter, suffix);
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

  /* And now for the real name.  We avoid storing the name "smime.p7m"
     because that one is used at several places in the mapi conversion
     functions.  */
  if (ctx->mimestruct_cur && ctx->mimestruct_cur->filename)
    {
      prop.ulPropTag = PR_ATTACH_LONG_FILENAME_W;
      wchar_t * utf16_str = NULL;
      if (!strcmp (ctx->mimestruct_cur->filename, "smime.p7m"))
        prop.Value.lpszW = L"x-smime.p7m";
      else
        {
          utf16_str = utf8_to_wchar (ctx->mimestruct_cur->filename);
          prop.Value.lpszW = utf16_str;
        }
      hr = HrSetOneProp ((LPMAPIPROP)newatt, &prop);
      xfree (utf16_str);
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

  /* If we have the MIME info and a charset info and that is not
     UTF-8, set our own Charset property.  */
  if (ctx->mimestruct_cur)
    {
      const char *s = ctx->mimestruct_cur->charset;
      if (s && strcmp (s, "utf-8") && strcmp (s, "UTF-8")
          && strcmp (s, "utf8") && strcmp (s, "UTF8"))
        mapi_set_gpgol_charset ((LPMESSAGE)newatt, s);
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

  if (debug_mime_parser)
    log_debug ("%s:%s: for ctx=%p cancel=%d", SRCNAME, __func__, ctx, cancel);

  if (ctx->outstream && ctx->is_body && !ctx->body_saved.outstream)
    {
      ctx->body_saved.outstream = ctx->outstream;
      ctx->outstream = NULL;
      retval = 0;
    }
  else if (ctx->outstream)
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

  if (ctx->mapi_attach && ctx->is_body && !ctx->body_saved.mapi_attach)
    {
      ctx->body_saved.mapi_attach = ctx->mapi_attach;
      ctx->mapi_attach = NULL;
    }
  else if (ctx->mapi_attach)
    {
      IAttach_Release (ctx->mapi_attach);
      ctx->mapi_attach = NULL;
    }

  if (ctx->symenc && ctx->is_body && !ctx->body_saved.symenc)
    {
      ctx->body_saved.symenc = ctx->symenc;
      ctx->symenc = NULL;
    }
  else if (ctx->symenc)
    {
      symenc_close (ctx->symenc);
      ctx->symenc = NULL;
    }

  ctx->is_body = 0;
  
  return retval;
}


/* Finish the saved body part.  This is required because we delay the
   finishing of body parts.  */
static int 
finish_saved_body (mime_context_t ctx, int cancel)
{
  HRESULT hr;
  int retval = -1;

  if (ctx->body_saved.outstream)
    {
      IStream_Commit (ctx->body_saved.outstream, 0);
      IStream_Release (ctx->body_saved.outstream);
      ctx->body_saved.outstream = NULL;

      if (cancel)
        retval = 0;
      else if (ctx->body_saved.mapi_attach)
        {
          hr = IAttach_SaveChanges (ctx->body_saved.mapi_attach, 0);
          if (hr)
            {
              log_error ("%s:%s: SaveChanges(attachment) failed: hr=%#lx\n",
                         SRCNAME, __func__, hr); 
            }
          else
            retval = 0; /* Success.  */
        }
    }

  if (ctx->body_saved.mapi_attach)
    {
      IAttach_Release (ctx->body_saved.mapi_attach);
      ctx->body_saved.mapi_attach = NULL;
    }

  if (ctx->symenc)
    {
      symenc_close (ctx->body_saved.symenc);
      ctx->body_saved.symenc = NULL;
    }

  return retval;
}



/* Create the MIME info string.  This is a LF delimited string
   with one line per MIME part.  Each line is formatted this way:
   LEVEL:ENCINFO:SIGINFO:CT:CHARSET:FILENAME
   
   LEVEL is the nesting level with 0 as the top (rfc822 header)
   ENCINFO is one of
      p   PGP/MIME encrypted
      s   S/MIME encryptyed
   SIGINFO is one of
      pX  PGP/MIME signed (also used for clearsigned)
      sX  S/MIME signed
      With X being:
        ?  unklnown status
        -  Bad signature
        ~  Good signature but with some problems
        !  Good signature
   CT ist the content type of this part
   CHARSET is the charset used for this part
   FILENAME is the file name.
*/
static char *
build_mimeinfo (mimestruct_item_t mimestruct)
{
  mimestruct_item_t ms;
  size_t buflen, n;
  char *buffer, *p;
  char numbuf[20];

  /* FIXME: We need to escape stuff so that there are no colons.  */
  for (buflen=0, ms = mimestruct; ms; ms = ms->next)
    {
      buflen += sizeof numbuf;
      buflen += strlen (ms->content_type);
      buflen += ms->charset? strlen (ms->charset) : 0;
      buflen += ms->filename? strlen (ms->filename) : 0;
      buflen += 20;
    }

  p = buffer = xmalloc (buflen+1);
  for (ms=mimestruct; ms; ms = ms->next)
    {
      snprintf (p, buflen, "%d:::%s:%s:%s:\n",
                ms->level, ms->content_type,
                ms->charset? ms->charset : "",
                ms->filename? ms->filename : "");
      n = strlen (p);
      assert (n < buflen);
      buflen -= n;
      p += n;
    }

  return buffer;
}


static int
finish_message (LPMESSAGE message, gpg_error_t err, int protect_mode, 
                mimestruct_item_t mimestruct)
{
  HRESULT hr;
  SPropValue prop;

  /* If this was an encrypted message we save the session marker in a
     special property so that we later know that we already decrypted
     that message within this session.  This is pretty useful when
     scrolling through messages and preview decryption has been
     enabled.  */
  if (protect_mode)
    {
      char sesmrk[8];

      if (get_gpgollastdecrypted_tag (message, &prop.ulPropTag) )
        return -1;
      if (err)
        memset (sesmrk, 0, 8);
      else
        memcpy (sesmrk, get_64bit_session_marker (), 8);
      prop.Value.bin.cb = 8;
      prop.Value.bin.lpb = sesmrk;
      hr = IMessage_SetProps (message, 1, &prop, NULL);
      if (hr)
        {
          log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                     SRCNAME, __func__, "GpgOL Last Decrypted", hr); 
          return -1;
        }
    }

  /* Store the MIME structure away.  */
  if (get_gpgolmimeinfo_tag (message, &prop.ulPropTag) )
    return -1;
  prop.Value.lpszA = build_mimeinfo (mimestruct);
  hr = IMessage_SetProps (message, 1, &prop, NULL);
  xfree (prop.Value.lpszA);
  if (hr)
    {
      log_error_w32 (hr, "%s:%s: error setting the mime info",
                     SRCNAME, __func__);
      return -1;
    }

  return mapi_save_changes (message, KEEP_OPEN_READWRITE|FORCE_SAVE);
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
  int not_inline_text = 0;
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
        filename = rfc2047_parse (s);
      s = rfc822parse_query_parameter (field, NULL, 1);
      if (s && strcmp (s, "inline"))
        not_inline_text = 1;
      rfc822parse_release_field (field);
    }

  /* Process the Content-type and all its parameters.  */
  ctmain = ctsub = NULL;
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

  if (debug_mime_parser)
    log_debug ("%s:%s: ctx=%p, ct=`%s/%s'\n",
               SRCNAME, __func__, ctx, ctmain, ctsub);

  s = rfc822parse_query_parameter (field, "charset", 0);
  if (s)
    charset = xstrdup (s);

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
      /* Check whether this attachment is an opaque signed S/MIME
         part.  We use a counter to later check that there is only one
         such part. */
      if (!strcmp (ctmain, "application")
          && (!strcmp (ctsub, "pkcs7-mime")
              || !strcmp (ctsub, "x-pkcs7-mime")))
        {
          const char *smtype = rfc822parse_query_parameter (field,
                                                            "smime-type", 0);
          if (smtype && !strcmp (smtype, "signed-data"))
            ctx->is_opaque_signed++;
          else
            {
              /* CryptoEx is notorious in setting wrong MIME header.
                 Mark that so we can test later if possible. */
              ctx->may_be_opaque_signed++;
            }
        }

      if (!ctx->preview)
        ctx->collect_attachment = 1;
    }
  rfc822parse_release_field (field); /* (Content-type) */
  ctx->in_data = 1;

  /* Need to start an attachment if we have seen a content disposition
     other then the inline type.  */ 
  if (is_text && not_inline_text)
    ctx->collect_attachment = 1;

  if (debug_mime_parser)
    log_debug ("%s:%s: this body: nesting=%d partno=%d is_text=%d, is_opq=%d"
               " charset=\"%s\"\n",
               SRCNAME, __func__, 
               ctx->nesting_level, ctx->part_counter, is_text, 
               ctx->is_opaque_signed,
               ctx->mimestruct_cur->charset?ctx->mimestruct_cur->charset:"");

  /* If this is a text part, decide whether we treat it as our body. */
  if (is_text && !not_inline_text)
    {
      ctx->collect_attachment = 1;

      /* If this is the first text part at all we will start to
         collect it and use it later as the regular body.  */
      if (!ctx->body_seen)
        {
          ctx->body_seen = 1;
          if (start_attachment (ctx, 1))
            return -1;
          assert (ctx->outstream);
        }
      else if (!ctx->body_saved.outstream || !ctx->body_saved.mapi_attach)
        {
          /* Oops: We expected to continue a body but the state is not
             correct.  Create a plain attachment instead.  */
          log_debug ("%s:%s: ctx=%p, no saved outstream or mapi_attach (%p,%p)",
                     SRCNAME, __func__, ctx, 
                     ctx->body_saved.outstream, ctx->body_saved.mapi_attach);
          if (start_attachment (ctx, 0))
            return -1;
          assert (ctx->outstream);
        }
      else if (ctx->outstream || ctx->mapi_attach || ctx->symenc)
        {
          /* We expected to continue a body but the last attachment
             has not been properly closed.  Create a plain attachment
             instead.  */
          log_debug ("%s:%s: ctx=%p, outstream, mapi_attach or symenc not "
                     "closed (%p,%p,%p)",
                     SRCNAME, __func__, ctx, 
                     ctx->outstream, ctx->mapi_attach, ctx->symenc);
          if (start_attachment (ctx, 0))
            return -1;
          assert (ctx->outstream);
        }
      else 
        {
          /* We already got one body and thus we can continue that
             last attachment.  */
          if (debug_mime_parser)
            log_debug ("%s:%s: continuing body part\n", SRCNAME, __func__);
          ctx->is_body = 1;
          ctx->outstream = ctx->body_saved.outstream;
          ctx->mapi_attach = ctx->body_saved.mapi_attach;
          ctx->symenc = ctx->body_saved.symenc;
          ctx->body_saved.outstream = NULL;
          ctx->body_saved.mapi_attach = NULL;
          ctx->body_saved.symenc = NULL;
        }
    }
  else if (ctx->collect_attachment)
    {
      /* Now that if we have an attachment prepare a new MAPI
         attachment.  */
      if (start_attachment (ctx, 0))
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
  if (ctx->no_mail_header)
    {
      /* Assume that this is not a regular mail but plain text. */
      if (event == RFC822PARSE_OPEN)
        return 0; /*  We need to skip the OPEN event.  */
      if (!ctx->body_seen)
        {
          if (debug_mime_parser)
            log_debug ("%s:%s: assuming this is plain text without headers\n",
                       SRCNAME, __func__);
          ctx->in_data = 1;
          ctx->collect_attachment = 2; /* 2 so we don't skip the first line. */
          ctx->body_seen = 1;
          /* Create a fake MIME structure.  */
          /* Fixme: We might want to take it from the enclosing message.  */
          {
            const char ctmain[] = "text";
            const char ctsub[] = "plain";
            mimestruct_item_t ms;
            
            ms = xmalloc (sizeof *ms + strlen (ctmain) + 1 + strlen (ctsub));
            ctx->mimestruct_cur = ms;
            *ctx->mimestruct_tail = ms;
            ctx->mimestruct_tail = &ms->next;
            ms->next = NULL;
            strcpy (stpcpy (stpcpy (ms->content_type, ctmain), "/"), ctsub);
            ms->level = 0;
            ms->filename = NULL;
            ms->charset = NULL;
          }
          if (start_attachment (ctx, 1))
            return -1;
          assert (ctx->outstream);
        }
      return 0;
    }

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
          ctx->parser_error = gpg_error (GPG_ERR_GENERAL);
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

          if (debug_mime_data)
            log_debug ("%s:%s: ctx=%p, line=`%.*s'\n",
                       SRCNAME, __func__, ctx, (int)pos, ctx->linebuf);
          if (rfc822parse_insert (ctx->msg, ctx->linebuf, pos))
            {
              log_error ("%s: ctx=%p, rfc822 parser failed: %s\n",
                         SRCNAME, ctx, strerror (errno));
              ctx->parser_error = gpg_error (GPG_ERR_GENERAL);
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
                  int slbrk = 0;

                  if (ctx->is_qp_encoded)
                    len = qp_decode (ctx->linebuf, pos, &slbrk);
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
                  if (!hr && !ctx->is_base64_encoded && !slbrk)
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
                      ctx->parser_error = gpg_error (GPG_ERR_EIO);
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
                  int slbrk = 0;

                  if (ctx->is_qp_encoded)
                    len = qp_decode (ctx->linebuf, pos, &slbrk);
                  else if (ctx->is_base64_encoded)
                    len = b64_decode (&ctx->base64, ctx->linebuf, pos);
                  else
                    len = pos;
                  if (len)
                    gpgme_data_write (ctx->sig_data, ctx->linebuf, len);
                  if (!ctx->is_base64_encoded && !slbrk)
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



/* FIXME: Needs documentation!

   MIMEHACK make the verification code ignore the first two bytes.  */
int
mime_verify (protocol_t protocol, const char *message, size_t messagelen, 
             LPMESSAGE mapi_message, HWND hwnd, int preview_mode, int mimehack)
{
  gpg_error_t err = 0;
  mime_context_t ctx;
  const char *s;
  size_t len;
  char *signature = NULL;
  size_t sig_len;
  engine_filter_t filter = NULL;

  (void)protocol;
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
      if (debug_mime_data)
        log_debug ("%s:%s: passing '%.*s'\n", 
                   SRCNAME, __func__, (int)len, message);
      plaintext_handler (ctx, message, len);
      if (ctx->parser_error)
        {
          err = ctx->parser_error;
          break;
        }
      else if (ctx->line_too_long)
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
      signature = gpgme_data_release_and_get_mem (ctx->sig_data, &sig_len);
      ctx->sig_data = NULL; 
    }

  /* Now actually verify the signature. */
  if (!err && ctx->signed_data && signature)
    {
      gpgme_data_seek (ctx->signed_data, mimehack? 2:0, SEEK_SET);
      
      if ((err=engine_create_filter (&filter, NULL, NULL)))
        goto leave;
      engine_set_session_number (filter, engine_new_session_number ());
      {
        char *tmp = mapi_get_subject (mapi_message);
        engine_set_session_title (filter, tmp);
        xfree (tmp);
      }
      {
        char *from = mapi_get_from_address (mapi_message);
        err = engine_verify_start (filter, hwnd, signature, sig_len,
                                   ctx->protocol, from);
        xfree (from);
      }
      if (err)
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
  gpgme_free (signature);
  engine_cancel (filter);
  if (ctx)
    {
      /* Cancel any left open attachment.  */
      finish_attachment (ctx, 1); 
      /* Save the body atatchment. */
      finish_saved_body (ctx, 0);
      rfc822parse_close (ctx->msg);
      gpgme_data_release (ctx->signed_data);
      gpgme_data_release (ctx->sig_data);
      finish_message (mapi_message, err, ctx->protect_mode, ctx->mimestruct);
      while (ctx->mimestruct)
        {
          mimestruct_item_t tmp = ctx->mimestruct->next;
          xfree (ctx->mimestruct->filename);
          xfree (ctx->mimestruct->charset);
          xfree (ctx->mimestruct);
          ctx->mimestruct = tmp;
        }
      symenc_close (ctx->symenc);
      symenc_close (ctx->body_saved.symenc);
      xfree (ctx);
    }
  return err;
}


/* A special version of mime_verify which works only for S/MIME opaque
   signed messages.  The message is expected to be a binary CMS
   signature either as an ISTREAM (if instream is not NULL) or
   provided in a buffer (INBUFFER and INBUFERLEN).  This function
   passes the entire message to the crypto engine and then parses the
   (cleartext) output for rendering the data.  START_PART_COUNTER
   should normally be set to 0. */
int
mime_verify_opaque (protocol_t protocol, LPSTREAM instream, 
                    const char *inbuffer, size_t inbufferlen,
                    LPMESSAGE mapi_message, HWND hwnd, int preview_mode,
                    int start_part_counter)
{
  gpg_error_t err = 0;
  mime_context_t ctx;
  engine_filter_t filter = NULL;

  log_debug ("%s:%s: enter (protocol=%d)", SRCNAME, __func__, protocol);

  if ((instream && (inbuffer || inbufferlen))
      || (!instream && !inbuffer))
    return gpg_error (GPG_ERR_INV_VALUE);

  if (protocol != PROTOCOL_SMIME)
    return gpg_error (GPG_ERR_INV_VALUE);

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;
  ctx->verify_mode = 0;
  ctx->mapi_message = mapi_message;
  ctx->mimestruct_tail = &ctx->mimestruct;
  ctx->part_counter = start_part_counter;

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_syserror ();
      log_error ("%s:%s: failed to open the RFC822 parser: %s", 
                 SRCNAME, __func__, gpg_strerror (err));
      goto leave;
    }

  if ((err=engine_create_filter (&filter, plaintext_handler, ctx)))
    goto leave;
  engine_set_session_number (filter, engine_new_session_number ());
  {
    char *tmp = mapi_get_subject (mapi_message);
    engine_set_session_title (filter, tmp);
    xfree (tmp);
  }
  {
    char *from = mapi_get_from_address (mapi_message);
    err = engine_verify_start (filter, hwnd, NULL, 0, protocol, from);
    xfree (from);
  }
  if (err)
    goto leave;

  if (instream)
    {
      /* Filter the stream.  */
      do
        {
          HRESULT hr;
          ULONG nread;
          char buffer[4096];
      
          hr = IStream_Read (instream, buffer, sizeof buffer, &nread);
          if (hr)
            {
              log_error ("%s:%s: IStream::Read failed: hr=%#lx", 
                         SRCNAME, __func__, hr);
              err = gpg_error (GPG_ERR_EIO);
            }
          else if (nread)
            {
/*               if (debug_mime_data) */
/*                 log_hexdump (buffer, nread, "%s:%s: ctx=%p, data: ", */
/*                              SRCNAME, __func__, ctx); */
              err = engine_filter (filter, buffer, nread);
            }
          else
            {
/*               if (debug_mime_data) */
/*                 log_debug ("%s:%s: ctx=%p, data: EOF\n", */
/*                            SRCNAME, __func__, ctx); */
              break; /* EOF */
            }
        }
      while (!err);
    }
  else
    {
      /* Filter the buffer.  */
/*       if (debug_mime_data) */
/*         log_hexdump (inbuffer, inbufferlen, "%s:%s: ctx=%p, data: ", */
/*                      SRCNAME, __func__, ctx); */
      err = engine_filter (filter, inbuffer, inbufferlen);
    }
  if (err)
    goto leave;

  /* Wait for the engine to finish.  */
  if ((err = engine_filter (filter, NULL, 0)))
    goto leave;
  if ((err = engine_wait (filter)))
    goto leave;
  filter = NULL;

  if (ctx->parser_error)
    err = ctx->parser_error;
  else if (ctx->line_too_long)
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
      /* Save the body attachment. */
      finish_saved_body (ctx, 0);
      rfc822parse_close (ctx->msg);
      if (ctx->signed_data)
        gpgme_data_release (ctx->signed_data);
      if (ctx->sig_data)
        gpgme_data_release (ctx->sig_data);
      finish_message (mapi_message, err, ctx->protect_mode, ctx->mimestruct);
      while (ctx->mimestruct)
        {
          mimestruct_item_t tmp = ctx->mimestruct->next;
          xfree (ctx->mimestruct->filename);
          xfree (ctx->mimestruct->charset);
          xfree (ctx->mimestruct);
          ctx->mimestruct = tmp;
        }
      symenc_close (ctx->symenc);
      symenc_close (ctx->body_saved.symenc);
      xfree (ctx);
    }
  return err;
}



/* Process the transition to body event in the decryption parser.

   This means we have received the empty line indicating the body and
   should now check the headers to see what to do about this part.  */
static int
ciphermessage_t2body (mime_context_t ctx, rfc822parse_t msg)
{
  rfc822parse_field_t field;
  const char *ctmain, *ctsub;
  size_t off;
  char *p;
  int is_text = 0;
        
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

  /* Process the Content-type and all its parameters.  */
  /* Fixme: Currently we don't make any use of it but consider all the
     content to be the encrypted data.  */
  ctmain = ctsub = NULL;
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

  if (debug_mime_parser)
    log_debug ("%s:%s: ctx=%p, ct=`%s/%s'\n",
               SRCNAME, __func__, ctx, ctmain, ctsub);

  rfc822parse_release_field (field); /* (Content-type) */
  ctx->in_data = 1;

  if (debug_mime_parser)
    log_debug ("%s:%s: this body: nesting=%d part_counter=%d is_text=%d\n",
               SRCNAME, __func__, 
               ctx->nesting_level, ctx->part_counter, is_text);


  return 0;
}

/* This routine gets called by the RFC822 decryption parser for all
   kind of events.  Should return 0 on success or -1 as well as
   setting errno on failure.  */
static int
ciphermessage_cb (void *opaque, rfc822parse_event_t event, rfc822parse_t msg)
{
  int retval = 0;
  mime_context_t decctx = opaque;

  debug_message_event (decctx, event);

  switch (event)
    {
    case RFC822PARSE_T2BODY:
      retval = ciphermessage_t2body (decctx, msg);
      break;

    case RFC822PARSE_LEVEL_DOWN:
      decctx->nesting_level++;
      break;

    case RFC822PARSE_LEVEL_UP:
      if (decctx->nesting_level)
        decctx->nesting_level--;
      else 
        {
          log_error ("%s: decctx=%p, invalid structure: bad nesting level\n",
                     SRCNAME, decctx);
          decctx->parser_error = gpg_error (GPG_ERR_GENERAL);
        }
      break;
    
    case RFC822PARSE_BOUNDARY:
    case RFC822PARSE_LAST_BOUNDARY:
      decctx->any_boundary = 1;
      decctx->in_data = 0;
      break;
    
    case RFC822PARSE_BEGIN_HEADER:
      decctx->part_counter++;
      break;

    default:  /* Ignore all other events. */
      break;
    }

  return retval;
}


/* This handler is called by us with the MIME message containing the
   ciphertext. */
static int
ciphertext_handler (void *handle, const void *buffer, size_t size)
{
  mime_context_t ctx = handle;
  const char *s;
  size_t nleft, pos, len;
  gpg_error_t err;

  s = buffer;
  pos = ctx->linebufpos;
  nleft = size;
  for (; nleft ; nleft--, s++)
    {
      if (pos >= ctx->linebufsize)
        {
          log_error ("%s:%s: ctx=%p, rfc822 parser failed: line too long\n",
                     SRCNAME, __func__, ctx);
          ctx->line_too_long = 1;
          return -1; /* Error. */
        }
      if (*s != '\n')
        ctx->linebuf[pos++] = *s;
      else
        { /* Got a complete line.  Remove the last CR.  */
          if (pos && ctx->linebuf[pos-1] == '\r')
            pos--;

          if (debug_mime_data)
            log_debug ("%s:%s: ctx=%p, line=`%.*s'\n",
                       SRCNAME, __func__, ctx, (int)pos, ctx->linebuf);
          if (rfc822parse_insert (ctx->msg, ctx->linebuf, pos))
            {
              log_error ("%s:%s: ctx=%p, rfc822 parser failed: %s\n",
                         SRCNAME, __func__, ctx, strerror (errno));
              ctx->parser_error = gpg_error (GPG_ERR_GENERAL);
              return -1; /* Error. */
            }

          if (ctx->in_data)
            {
              /* We are inside the data.  That should be the actual
                 ciphertext in the given encoding.  Pass it on to the
                 crypto engine. */
              int slbrk = 0;

              if (ctx->is_qp_encoded)
                len = qp_decode (ctx->linebuf, pos, &slbrk);
              else if (ctx->is_base64_encoded)
                len = b64_decode (&ctx->base64, ctx->linebuf, pos);
              else
                len = pos;
              if (len)
                err = engine_filter (ctx->outfilter, ctx->linebuf, len);
              else
                err = 0;
              if (!err && !ctx->is_base64_encoded && !slbrk)
                {
                  char tmp[3] = "\r\n";
                  err = engine_filter (ctx->outfilter, tmp, 2);
                }
              if (err)
                {
                  log_debug ("%s:%s: sending ciphertext to engine failed: %s",
                             SRCNAME, __func__, gpg_strerror (err));
                  ctx->parser_error = err;
                  return -1; /* Error. */
                }
            }
          
          /* Continue with next line. */
          pos = 0;
        }
    }
  ctx->linebufpos = pos;

  return size;
}



/* Decrypt the PGP or S/MIME message taken from INSTREAM.  HWND is the
   window to be used for message box and such.  In PREVIEW_MODE no
   verification will be done, no messages saved and no messages boxes
   will pop up.  If IS_RFC822 is set, the message is expected to be in
   rfc822 format.  The caller should send SIMPLE_PGP if the input
   message is a simple (non-MIME) PGP message.  If SIG_ERR is not null
   and a signature was found and verified, its status is returned
   there.  If no signature was found SIG_ERR is not changed. */
int
mime_decrypt (protocol_t protocol, LPSTREAM instream, LPMESSAGE mapi_message,
              int is_rfc822, int simple_pgp, HWND hwnd, int preview_mode,
              gpg_error_t *sig_err)
{
  gpg_error_t err;
  mime_context_t decctx, ctx;
  engine_filter_t filter = NULL;
  int opaque_signed = 0;
  int may_be_opaque_signed = 0;
  int last_part_counter = 0;
  unsigned int session_number;
  char *signature = NULL;

  log_debug ("%s:%s: enter (protocol=%d, is_rfc822=%d, simple_pgp=%d)",
             SRCNAME, __func__, protocol, is_rfc822, simple_pgp);

  if (is_rfc822)
    {
      decctx = xcalloc (1, sizeof *decctx + LINEBUFSIZE);
      decctx->linebufsize = LINEBUFSIZE;
      decctx->hwnd = hwnd;
    }
  else
    decctx = NULL;

  ctx = xcalloc (1, sizeof *ctx + LINEBUFSIZE);
  ctx->linebufsize = LINEBUFSIZE;
  ctx->protect_mode = 1; 
  ctx->hwnd = hwnd;
  ctx->preview = preview_mode;
  ctx->verify_mode = simple_pgp? 0 : 1;
  ctx->mapi_message = mapi_message;
  ctx->mimestruct_tail = &ctx->mimestruct;
  ctx->no_mail_header = simple_pgp;

  if (decctx)
    {
      decctx->msg = rfc822parse_open (ciphermessage_cb, decctx);
      if (!decctx->msg)
        {
          err = gpg_error_from_syserror ();
          log_error ("%s:%s: failed to open the RFC822 decryption parser: %s", 
                     SRCNAME, __func__, gpg_strerror (err));
          goto leave;
        }
    }

  ctx->msg = rfc822parse_open (message_cb, ctx);
  if (!ctx->msg)
    {
      err = gpg_error_from_syserror ();
      log_error ("%s:%s: failed to open the RFC822 parser: %s", 
                 SRCNAME, __func__, gpg_strerror (err));
      goto leave;
    }

  /* Prepare the decryption.  */
  if ((err=engine_create_filter (&filter, plaintext_handler, ctx)))
    goto leave;
  if (simple_pgp)
    engine_request_extra_lf (filter);
  session_number = engine_new_session_number ();
  engine_set_session_number (filter, session_number);
  {
    char *tmp = mapi_get_subject (mapi_message);
    engine_set_session_title (filter, tmp);
    xfree (tmp);
  }
  {
    char *from = preview_mode? NULL : mapi_get_from_address (mapi_message);
    err = engine_decrypt_start (filter, hwnd, protocol, !preview_mode, from);
    xfree (from);
  }
  if (err)
    goto leave;

  if (decctx)
    decctx->outfilter = filter;
  
  /* Filter the stream.  */
  do
    {
      HRESULT hr;
      ULONG nread;
      char buffer[4096];
      
      /* For EOF detection we assume that Read returns no error and
         thus NREAD will be 0.  The specs say that "Depending on the
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
          if (decctx)
            {
               ciphertext_handler (decctx, buffer, nread);
               if (decctx->parser_error)
                 {
                   err = decctx->parser_error;
                   break;
                 }
               else if (decctx->line_too_long)
                 {
                   err = gpg_error (GPG_ERR_GENERAL);
                   break;
                 }
            }
          else
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

  if (ctx->parser_error)
    err = ctx->parser_error;
  else if (ctx->line_too_long)
    err = gpg_error (GPG_ERR_GENERAL);

  /* Verify an optional inner signature.  */
  if (!err && !preview_mode 
      && ctx->sig_data && ctx->signed_data && !ctx->is_opaque_signed)
    {
      size_t sig_len;

      assert (!filter);

      if (gpgme_data_write (ctx->sig_data, "", 1) == 1)
        {
          signature = gpgme_data_release_and_get_mem (ctx->sig_data, &sig_len);
          ctx->sig_data = NULL; 
        }

      if (!err && signature)
        {
          gpgme_data_seek (ctx->signed_data, 0, SEEK_SET);
          
          if ((err=engine_create_filter (&filter, NULL, NULL)))
            goto leave;
          engine_set_session_number (filter, session_number);
          {
            char *tmp = mapi_get_subject (mapi_message);
            engine_set_session_title (filter, tmp);
            xfree (tmp);
          }
          {
            char *from = mapi_get_from_address (mapi_message);
            err = engine_verify_start (filter, hwnd, signature, sig_len,
                                       ctx->protocol, from);
            xfree (from);
          }
          if (err)
            goto leave;

          /* Filter the data.  */
          do
            {
              int nread;
              char buffer[4096];
              
              nread = gpgme_data_read (ctx->signed_data, buffer,sizeof buffer);
              if (nread < 0)
                {
                  err = gpg_error_from_syserror ();
                  log_error ("%s:%s: gpgme_data_read failed in verify: %s", 
                             SRCNAME, __func__, gpg_strerror (err));
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
          err = engine_wait (filter);
          if (sig_err)
            *sig_err = err;
          err = 0;
          filter = NULL;
        }
    }


 leave:
  engine_cancel (filter);
  xfree (signature);
  signature = NULL;
  if (ctx)
    {
      /* Cancel any left over attachment which means that the MIME
         structure was not complete.  However if we have not seen any
         boundary the message is a non-MIME one but we may have
         started the body attachment (gpgol000.txt) - this one needs
         to be finished properly.  */
      finish_attachment (ctx, ctx->any_boundary? 1: 0);
      /* Save the body attachment.  */
      finish_saved_body (ctx, 0);
      rfc822parse_close (ctx->msg);
      if (ctx->signed_data)
        gpgme_data_release (ctx->signed_data);
      if (ctx->sig_data)
        gpgme_data_release (ctx->sig_data);
      finish_message (mapi_message, err, ctx->protect_mode, ctx->mimestruct);
      while (ctx->mimestruct)
        {
          mimestruct_item_t tmp = ctx->mimestruct->next;
          xfree (ctx->mimestruct->filename);
          xfree (ctx->mimestruct->charset);
          xfree (ctx->mimestruct);
          ctx->mimestruct = tmp;
        }
      symenc_close (ctx->symenc);
      symenc_close (ctx->body_saved.symenc);
      last_part_counter = ctx->part_counter;
      opaque_signed = (ctx->is_opaque_signed == 1);
      if (!opaque_signed && ctx->may_be_opaque_signed == 1)
        may_be_opaque_signed = 1;
      xfree (ctx);
    }
  if (decctx)
    {
      rfc822parse_close (decctx->msg);
      xfree (decctx);
    }

  if (!err && (opaque_signed || may_be_opaque_signed))
    {
      /* Handle an S/MIME opaque signed part.  The decryption has
         written an attachment we are now going to verify and render
         to the body attachment.  */
      mapi_attach_item_t *table;
      char *plainbuffer = NULL;
      size_t plainbufferlen;
      int i;

      table = mapi_create_attach_table (mapi_message, 0);
      if (!table)
        {
          if (opaque_signed)
            err = gpg_error (GPG_ERR_GENERAL);
          goto leave_verify;
        }

      for (i=0; !table[i].end_of_table; i++)
        if (table[i].attach_type == ATTACHTYPE_FROMMOSS
            && table[i].content_type               
            && (!strcmp (table[i].content_type, "application/pkcs7-mime")
                || !strcmp (table[i].content_type, "application/x-pkcs7-mime"))
            )
          break;
      if (table[i].end_of_table)
        {
          if (opaque_signed)
            {
              log_debug ("%s:%s: "
                         "attachment for opaque signed S/MIME not found",
                         SRCNAME, __func__);
              err = gpg_error (GPG_ERR_GENERAL);
            }
          goto leave_verify;
        }
      plainbuffer = mapi_get_attach (mapi_message, 1, table+i,
                                     &plainbufferlen);
      if (!plainbuffer)
        {
          if (opaque_signed)
            err = gpg_error (GPG_ERR_GENERAL);
          goto leave_verify;
        }

      /* Now that we have the data, we can check whether this is an
         S/MIME signature (without proper MIME headers). */
      if (may_be_opaque_signed 
          && is_cms_signed_data (plainbuffer, plainbufferlen))
        opaque_signed = 1;

      /* And check the signature.  */
      if (opaque_signed)
        {
          err = mime_verify_opaque (PROTOCOL_SMIME, NULL, 
                                    plainbuffer, plainbufferlen,
                                    mapi_message, hwnd, 0,
                                    last_part_counter+1);
          log_debug ("%s:%s: mime_verify_opaque returned %d", 
                     SRCNAME, __func__, err);
          if (sig_err)
            *sig_err = err;
          err = 0;
        }
      
    leave_verify:
      xfree (plainbuffer);
      mapi_release_attach_table (table);
    }

  return err;
}

