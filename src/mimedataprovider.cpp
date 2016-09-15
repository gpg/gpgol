/* mimedataprover.cpp - GpgME dataprovider for mime data
 *    Copyright (C) 2016 Intevation GmbH
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
#include "config.h"
#include "common_indep.h"
#include "xmalloc.h"
#include <string.h>

#include "mimedataprovider.h"
#include "parsetlv.h"
#include "rfc822parse.h"
#include "rfc2047parse.h"
#include "attachment.h"

#ifndef HAVE_W32_SYSTEM
#define stricmp strcasecmp
#endif

/* The maximum length of a line we are able to process.  RFC822 allows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000

/* How much data is read at once in collect */
#define BUFSIZE 8192

#include <gpgme++/error.h>

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
  rfc822parse_t msg;  /* The handle of the RFC822 parser. */

  int verify_mode;    /* True if we want to verify a signature. */
  int no_mail_header; /* True if we want to bypass all MIME parsing.  */

  int nesting_level;  /* Current MIME nesting level. */
  int in_data;        /* We are currently in data (body or attachment). */
  int body_seen;      /* True if we have seen a part we consider the
                         body of the message.  */

  int collect_attachment; /* True if we are collecting an attachment */
  std::shared_ptr<Attachment> current_attachment; /* A pointer to the current
                                                     attachment */
  int collect_body;       /* True if we are collcting the body */
  int collect_html_body;  /* True if we are collcting the html body */
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

  /* A linked list describing the structure of the mime message.  This
     list gets build up while parsing the message.  */
  mimestruct_item_t mimestruct;
  mimestruct_item_t *mimestruct_tail;
  mimestruct_item_t mimestruct_cur;

  int any_attachments_created;  /* True if we created a new atatchment.  */

  b64_state_t base64;     /* The state of the Base-64 decoder.  */

  gpg_error_t parser_error;   /* Indicates that we encountered a error from
                                 the parser. */
};
typedef struct mime_context *mime_context_t;

/* Print the message event EVENT. */
static void
debug_message_event (rfc822parse_event_t event)
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
  log_mime_parser ("%s: rfc822 event %s\n", SRCNAME, s);
}

/* Returns true if the BER encoded data in BUFFER is CMS signed data.
   LENGTH gives the length of the buffer, for correct detection LENGTH
   should be at least about 24 bytes.  */
#if 0
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
#endif

/* Process the transition to body event.

   This means we have received the empty line indicating the body and
   should now check the headers to see what to do about this part.  */
static int
t2body (MimeDataProvider *provider, rfc822parse_t msg)
{
  rfc822parse_field_t field;
  mime_context_t ctx = provider->mime_context ();
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

  log_mime_parser ("%s:%s: ctx=%p, ct=`%s/%s'\n",
                   SRCNAME, __func__, ctx, ctmain, ctsub);

  s = rfc822parse_query_parameter (field, "charset", 0);
  if (s)
    charset = xstrdup (s);

  if (!filename)
    {
      /* Check for Content-Type name if Content-Disposition filename
         was not found */
      s = rfc822parse_query_parameter (field, "name", 0);
      if (s)
        filename = rfc2047_parse (s);
    }

  /* Update our idea of the entire MIME structure.  */
  {
    mimestruct_item_t ms;

    ms = (mimestruct_item_t) xmalloc (sizeof *ms + strlen (ctmain) + 1 + strlen (ctsub));
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
      if (!provider->signature()
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
  else if (ctx->nesting_level == 1 && !provider->signature()
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
      ctx->collect_signature = 1;
      log_mime_parser ("Collecting signature now");
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

      ctx->collect_attachment = 1;
    }
  rfc822parse_release_field (field); /* (Content-type) */
  ctx->in_data = 1;

  /* Need to start an attachment if we have seen a content disposition
     other then the inline type. */
  if (is_text && not_inline_text)
    ctx->collect_attachment = 1;

  log_mime_parser ("%s:%s: this body: nesting=%d partno=%d is_text=%d, is_opq=%d"
                   " charset=\"%s\"\n",
                   SRCNAME, __func__,
                   ctx->nesting_level, ctx->part_counter, is_text,
                   ctx->is_opaque_signed,
                   ctx->mimestruct_cur->charset?ctx->mimestruct_cur->charset:"");

  /* If this is a text part, decide whether we treat it as our body. */
  if (is_text && !not_inline_text)
    {
      ctx->collect_attachment = 1;
      ctx->body_seen = 1;
      if (is_text == 2)
        {
          ctx->collect_html_body = 1;
          ctx->collect_body = 0;
        }
      else
        {
          ctx->collect_body = 1;
          ctx->collect_html_body = 0;
        }
    }
  else if (ctx->collect_attachment)
    {
      /* Now that if we have an attachment prepare a new MAPI
         attachment.  */
      ctx->current_attachment = provider->create_attachment();
    }

  return 0;
}

static int
message_cb (void *opaque, rfc822parse_event_t event,
            rfc822parse_t msg)
{
  int retval = 0;

  MimeDataProvider *provider = static_cast<MimeDataProvider*> (opaque);

  mime_context_t ctx = provider->mime_context();

  debug_message_event (event);
  if (ctx->no_mail_header)
    {
      /* Assume that this is not a regular mail but plain text. */
      if (event == RFC822PARSE_OPEN)
        return 0; /*  We need to skip the OPEN event.  */
      if (!ctx->body_seen)
        {
          log_mime_parser ("%s:%s: assuming this is plain text without headers\n",
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

            ms = (mimestruct_item_t) xmalloc (sizeof *ms + strlen (ctmain) + 1 + strlen (ctsub));
            ctx->mimestruct_cur = ms;
            *ctx->mimestruct_tail = ms;
            ctx->mimestruct_tail = &ms->next;
            ms->next = NULL;
            strcpy (stpcpy (stpcpy (ms->content_type, ctmain), "/"), ctsub);
            ms->level = 0;
            ms->filename = NULL;
            ms->charset = NULL;
          }
          ctx->collect_body = 1;
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
        }
    }


  switch (event)
    {
    case RFC822PARSE_T2BODY:
      retval = t2body (provider, msg);
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
      ctx->collect_body = 0;

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

MimeDataProvider::MimeDataProvider() :
  m_signature(nullptr)
{
  m_mime_ctx = (mime_context_t) xcalloc (1, sizeof *m_mime_ctx);
  m_mime_ctx->msg = rfc822parse_open (message_cb, this);
  m_mime_ctx->mimestruct_tail = &m_mime_ctx->mimestruct;
}

#ifdef HAVE_W32_SYSTEM
MimeDataProvider::MimeDataProvider(LPSTREAM stream):
  MimeDataProvider()
{
  if (stream)
    {
      stream->AddRef ();
    }
  else
    {
      log_error ("%s:%s called without stream ", SRCNAME, __func__);
      return;
    }
  log_mime_parser ("%s:%s Collecting data.", SRCNAME, __func__);
  collect_data (stream);
  log_mime_parser ("%s:%s Data collected.", SRCNAME, __func__);
  gpgol_release (stream);
}
#endif

MimeDataProvider::MimeDataProvider(FILE *stream):
  MimeDataProvider()
{
  log_mime_parser ("%s:%s Collecting data from file.", SRCNAME, __func__);
  collect_data (stream);
  log_mime_parser ("%s:%s Data collected.", SRCNAME, __func__);
}

MimeDataProvider::~MimeDataProvider()
{
  log_debug ("%s:%s", SRCNAME, __func__);
  free (m_mime_ctx);
  if (m_signature)
    {
      delete m_signature;
    }
}

bool
MimeDataProvider::isSupported(GpgME::DataProvider::Operation op) const
{
  return op == GpgME::DataProvider::Read ||
         op == GpgME::DataProvider::Seek ||
         op == GpgME::DataProvider::Write ||
         op == GpgME::DataProvider::Release;
}

ssize_t
MimeDataProvider::read(void *buffer, size_t size)
{
  log_mime_parser ("%s:%s: Reading: " SIZE_T_FORMAT "Bytes",
                 SRCNAME, __func__, size);
  ssize_t bRead = m_crypto_data.read (buffer, size);
  if (opt.enable_debug & DBG_MIME_PARSER)
    {
      std::string buf ((char *)buffer, bRead);
      log_mime_parser ("%s:%s: Data: \"%s\"",
                     SRCNAME, __func__, buf.c_str());
    }
  return bRead;
}

/* Split some raw data into lines and handle them accordingly.
   returns the amount of bytes not taken from the input buffer.
*/
size_t
MimeDataProvider::collect_input_lines(const char *input, size_t insize)
{
  char linebuf[LINEBUFSIZE];
  const char *s = input;
  size_t pos = 0;
  size_t nleft = insize;
  size_t not_taken = nleft;
  size_t len = 0;

  /* Split the raw data into lines */
  for (; nleft; nleft--, s++)
    {
      if (pos >= LINEBUFSIZE)
        {
          log_error ("%s:%s: rfc822 parser failed: line too long\n",
                     SRCNAME, __func__);
          GpgME::Error::setSystemError (GPG_ERR_EIO);
          return not_taken;
        }
      if (*s != '\n')
        linebuf[pos++] = *s;
      else
        {
          /* Got a complete line.  Remove the last CR.  */
          not_taken -= pos + 1; /* Pos starts at 0 so + 1 for it */
          if (pos && linebuf[pos-1] == '\r')
            {
              pos--;
            }

          log_mime_parser("%s:%s: Parsing line=`%.*s'\n",
                          SRCNAME, __func__, (int)pos, linebuf);
          /* Check the next state */
          if (rfc822parse_insert (m_mime_ctx->msg,
                                  (unsigned char*) linebuf,
                                  pos))
            {
              log_error ("%s:%s: rfc822 parser failed: %s\n",
                         SRCNAME, __func__, strerror (errno));
              return not_taken;
            }

          /* If we are currently in a collecting state actually
             collect that line */
          if (m_mime_ctx->collect_signeddata)
            {
              /* Save the signed data.  Note that we need to delay
                 the CR/LF because the last line ending belongs to the
                 next boundary. */
              if (m_mime_ctx->collect_signeddata == 2)
                {
                  m_crypto_data.write ("\r\n", 2);
                }
              log_debug ("Writing signeddata: %s pos: " SIZE_T_FORMAT,
                         linebuf, pos);
              m_crypto_data.write (linebuf, pos);
              m_mime_ctx->collect_signeddata = 2;
            }
          if (m_mime_ctx->in_data && m_mime_ctx->collect_attachment)
            {
              /* We are inside of an attachment part.  Write it out. */
              if (m_mime_ctx->collect_attachment == 1)  /* Skip the first line. */
                m_mime_ctx->collect_attachment = 2;

              int slbrk = 0;
              if (m_mime_ctx->is_qp_encoded)
                len = qp_decode (linebuf, pos, &slbrk);
              else if (m_mime_ctx->is_base64_encoded)
                len = b64_decode (&m_mime_ctx->base64, linebuf, pos);
              else
                len = pos;

              if (m_mime_ctx->collect_body)
                {
                  m_body += std::string(linebuf, len);
                  if (!m_mime_ctx->is_base64_encoded && !slbrk)
                    {
                      m_body += "\r\n";
                    }
                }
              else if (m_mime_ctx->collect_html_body)
                {
                  m_html_body += std::string(linebuf, len);
                  if (!m_mime_ctx->is_base64_encoded && !slbrk)
                    {
                      m_body += "\r\n";
                    }
                }
              else if (m_mime_ctx->current_attachment && len)
                {
                  m_mime_ctx->current_attachment->write(linebuf, len);
                  if (!m_mime_ctx->is_base64_encoded && !slbrk)
                    {
                      m_mime_ctx->current_attachment->write("\r\n", 2);
                    }
                }
              else
                {
                  log_mime_parser ("%s:%s Collecting ended / failed.",
                                   SRCNAME, __func__);
                }
            }
          else if (m_mime_ctx->in_data && m_mime_ctx->collect_signature)
            {
              /* We are inside of a signature attachment part.  */
              if (m_mime_ctx->collect_signature == 1)  /* Skip the first line. */
                m_mime_ctx->collect_signature = 2;
              else
                {
                  int slbrk = 0;

                  if (m_mime_ctx->is_qp_encoded)
                    len = qp_decode (linebuf, pos, &slbrk);
                  else if (m_mime_ctx->is_base64_encoded)
                    len = b64_decode (&m_mime_ctx->base64, linebuf, pos);
                  else
                    len = pos;
                  if (!m_signature)
                    {
                      m_signature = new GpgME::Data();
                    }
                  if (len)
                    m_signature->write(linebuf, len);
                  if (!m_mime_ctx->is_base64_encoded && !slbrk)
                    m_signature->write("\r\n", 2);
                }
            }
          else if (m_mime_ctx->in_data)
            {
              /* We are inside the data.  That should be the actual
                 ciphertext in the given encoding. */
              int slbrk = 0;

              if (m_mime_ctx->is_qp_encoded)
                len = qp_decode (linebuf, pos, &slbrk);
              else if (m_mime_ctx->is_base64_encoded)
                len = b64_decode (&m_mime_ctx->base64, linebuf, pos);
              else
                len = pos;
              if (len)
                m_crypto_data.write(linebuf, len);
              if (!m_mime_ctx->is_base64_encoded && !slbrk)
                m_crypto_data.write("\r\n", 2);
            }
          /* Continue with next line. */
          pos = 0;
        }
    }
  return not_taken;
}

#ifdef HAVE_W32_SYSTEM
void
MimeDataProvider::collect_data(LPSTREAM stream)
{
  if (!stream)
    {
      return;
    }
  HRESULT hr;
  char buf[BUFSIZE];
  ULONG bRead;
  while ((hr = stream->Read (buf, BUFSIZE, &bRead)) == S_OK ||
         hr == S_FALSE)
    {
      if (!bRead)
        {
          log_mime_parser ("%s:%s: Input stream at EOF.",
                           SRCNAME, __func__);
          return;
        }
      log_mime_parser ("%s:%s: Read %lu bytes.",
                       SRCNAME, __func__, bRead);

      m_rawbuf += std::string (buf, bRead);
      size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                              m_rawbuf.size());

      if (not_taken == m_rawbuf.size())
        {
          log_error ("%s:%s: Collect failed to consume anything.\n"
                     "Buffer too small?",
                     SRCNAME, __func__);
          return;
        }
      log_mime_parser ("%s:%s: Consumed: " SIZE_T_FORMAT " bytes",
                       SRCNAME, __func__, m_rawbuf.size() - not_taken);
      m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
    }
}
#endif

void
MimeDataProvider::collect_data(FILE *stream)
{
  if (!stream)
    {
      return;
    }
  char buf[BUFSIZE];
  size_t bRead;
  while ((bRead = fread (buf, 1, BUFSIZE, stream)) > 0)
    {
      log_mime_parser ("%s:%s: Read " SIZE_T_FORMAT " bytes.",
                       SRCNAME, __func__, bRead);

      m_rawbuf += std::string (buf, bRead);
      size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                              m_rawbuf.size());

      if (not_taken == m_rawbuf.size())
        {
          log_error ("%s:%s: Collect failed to consume anything.\n"
                     "Buffer too small?",
                     SRCNAME, __func__);
          return;
        }
      log_mime_parser ("%s:%s: Consumed: " SIZE_T_FORMAT " bytes",
                       SRCNAME, __func__, m_rawbuf.size() - not_taken);
      m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
    }
}

ssize_t MimeDataProvider::write(const void *buffer, size_t bufSize)
{
    m_rawbuf += std::string ((const char*)buffer, bufSize);
    size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                            m_rawbuf.size());

    if (not_taken == m_rawbuf.size())
      {
        log_error ("%s:%s: Write failed to consume anything.\n"
                   "Buffer too small? or no newlines in text?",
                   SRCNAME, __func__);
        return bufSize;
      }
    log_mime_parser ("%s:%s: Write Consumed: " SIZE_T_FORMAT " bytes",
                     SRCNAME, __func__, m_rawbuf.size() - not_taken);
    m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
    return bufSize;
}

off_t
MimeDataProvider::seek(off_t offset, int whence)
{
  return m_crypto_data.seek (offset, whence);
}

GpgME::Data *
MimeDataProvider::signature() const
{
  return m_signature;
}

std::shared_ptr<Attachment>
MimeDataProvider::create_attachment()
{
  log_mime_parser ("%s:%s: Creating attachment.",
                   SRCNAME, __func__);

  auto attach = std::shared_ptr<Attachment> (new Attachment());
  attach->set_attach_type (ATTACHTYPE_FROMMOSS);
  m_mime_ctx->any_attachments_created = 1;

  /* And now for the real name.  We avoid storing the name "smime.p7m"
     because that one is used at several places in the mapi conversion
     functions.  */
  if (m_mime_ctx->mimestruct_cur && m_mime_ctx->mimestruct_cur->filename)
    {
      if (!strcmp (m_mime_ctx->mimestruct_cur->filename, "smime.p7m"))
        {
          attach->set_display_name ("x-smime.p7m");
        }
      else
        {
          attach->set_display_name (m_mime_ctx->mimestruct_cur->filename);
        }
    }
  m_attachments.push_back (attach);

  return attach;
  /* TODO handle encoding */
}

const std::string &MimeDataProvider::get_body()
{
  if (m_rawbuf.size())
    {
      /* If there was some data left in the rawbuf this could
         mean that some plaintext was not finished with a linefeed.
         In that case we append it to the bodies. */
      m_body += m_rawbuf;
      m_html_body += m_rawbuf;
      m_rawbuf.clear();
    }
  return m_body;
}

const std::string &MimeDataProvider::get_html_body()
{
  if (m_rawbuf.size())
    {
      /* If there was some data left in the rawbuf this could
         mean that some plaintext was not finished with a linefeed.
         In that case we append it to the bodies. */
      m_body += m_rawbuf;
      m_html_body += m_rawbuf;
      m_rawbuf.clear();
    }
  return m_html_body;
}
