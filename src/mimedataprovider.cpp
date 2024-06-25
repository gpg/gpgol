/* mimedataprover.cpp - GpgME dataprovider for mime data
 * Copyright (C) 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
#include "config.h"
#include "common_indep.h"
#include "xmalloc.h"
#include <string.h>
#include <vector>
#include <sstream>

#include "mimedataprovider.h"
#include "rfc822parse.h"
#include "rfc2047parse.h"
#include "attachment.h"
#include "cpphelp.h"

#ifndef HAVE_W32_SYSTEM
#define stricmp strcasecmp
#endif

/* How much data is read at once in collect */
#define BUFSIZE 65536

/* The maximum length of a line we are able to process.  RFC822 allows
   only for 1000 bytes; thus 2000 seemed to be a reasonable value until
   we worked with a MUA that sent whole HTML mails in a single line
   now we read more at once and allow the line to be the full size of
   our read. -1 */
#define LINEBUFSIZE (BUFSIZE - 1)

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
  char *cid;            /* Malloced content id or NULL. */
  char *charset;        /* Malloced charset or NULL.  */
  char content_type[1]; /* String with the content type. */
};

/* The context object we use to track information. */
struct mime_context
{
  rfc822parse_t msg;  /* The handle of the RFC822 parser. */

  int verify_mode;    /* True if we want to verify a signature. */

  int nesting_level;  /* Current MIME nesting level. */
  int in_data;        /* We are currently in data (body or attachment). */
  int body_seen;      /* True if we have seen a part we consider the
                         body of the message.  */

  std::shared_ptr<Attachment> current_attachment; /* A pointer to the current
                                                     attachment */
  int collect_body;       /* True if we are collcting the body */
  int collect_html_body;  /* True if we are collcting the html body */
  int collect_crypto_data; /* True if we are collecting the signed data. */
  int collect_signature;  /* True if we are collecting a signature.  */
  int pgp_marker_checked; /* Checked if the first body line is pgp marker*/
  int is_encrypted;       /* True if we are working on an encrypted mail. */
  int start_hashing;      /* Flag used to start collecting signed data. */
  int hashing_level;      /* MIME level where we started hashing. */
  int is_qp_encoded;      /* Current part is QP encoded. */
  int is_base64_encoded;  /* Current part is base 64 encoded. */
  int is_body;            /* The current part belongs to the body.  */
  protocol_t protocol;    /* The detected crypto protocol.  */

  int part_counter;       /* Counts the number of processed parts. */
  int any_boundary;       /* Indicates whether we have seen any
                             boundary which means that we are actually
                             working on a MIME message and not just on
                             plain rfc822 message.  */
  int in_protected_headers; /* Indicates if we are in a mime part that was
                               marked by a protected headers header. */

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
  log_debug ("%s: rfc822 event %s\n", SRCNAME, s);
}

/* returns true if the BER encoded data in BUFFER is CMS signed data.
   LENGTH gives the length of the buffer, for correct detection LENGTH
   should be at least about 24 bytes.  */
#if 0
static int
is_cms_signed_data (const char *buffer, size_t length)
{
  TSTART;
  const char *p = buffer;
  size_t n = length;
  tlvinfo_t ti;

  if (parse_tlv (&p, &n, &ti))
    {
      TRETURN 0;
    }
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_SEQUENCE
        && ti.is_cons) )
    TRETURN 0;
  if (parse_tlv (&p, &n, &ti))
    {
      TRETURN 0;
    }
  if (!(ti.cls == ASN1_CLASS_UNIVERSAL && ti.tag == ASN1_TAG_OBJECT_ID
        && !ti.is_cons && ti.length) || ti.length > n)
    TRETURN 0;
  if (ti.length == 9 && !memcmp (p, "\x2A\x86\x48\x86\xF7\x0D\x01\x07\x02", 9))
    {
      TRETURN 1;
    }
  TRETURN 0;
}
#endif


/* RFC 2231 and 2047 are very close in their specification so to
   reuse our established 2047 parser we ignore the part that the
   2047 parser ignored anyway (language) and then fix up the syntax
   so that it matches. */
static std::string
rfc2231_to_2047 (const std::string &in)
{
  TSTART;
  /* No we do not care about language information in our parameters...
     since 2231 always uses QP encoding we just make the beginning and
     end look like 2047 =? ?= enclosed things and remove the language
     part. */
  auto i = in.find("'");
  std::string ret;
  if (i == std::string::npos)
    {
      /* No hyphen, just assume percent encoding */
      /* Something like: title*1*=%2A%2A%2Afun%2A%2A%2A%20 */
      ret = std::string ("=?US-ASCII?Q?") + in + std::string ("?=");
      find_and_replace (ret, "%", "=");
      log_dbg ("Assuming just QP Encoded ASCII for: '%s'", anonstr (in.c_str ()));
      TRETURN ret;
    }
  /* Something like: title*0*=us-ascii'en'This%20is%20even%20more%20 */
  /* Decorate as 2047 string */
  ret = std::string ("=?") + in + std::string ("?=");
  find_and_replace (ret, "%", "=");
  /* Insert the QP marker where the language would be */
  i = ret.find("'");
  ret.insert(i, "?Q?");

  /* Remove the language part we do not support this */
  auto langstart = ret.find("'");
  auto langend = ret.find("'", langstart + 1);
  if (langstart == std::string::npos || langend == std::string::npos ||
      langstart == langend)
    {
      /* Well whatever, lets just return whatever that is */
      log_dbg ("Failed to handle '%s'", anonstr (in.c_str ()));
      TRETURN ret;
    }
  ret.erase(langstart, langend - langstart + 1);
  TRETURN ret;
}

/* Handle continued parameters according to rfc2231. */
static char *
rfc2231_query_parameter (rfc822parse_field_t field, const char *param)
{
  TSTART;
  const char *s = rfc822parse_query_parameter (field, param, 0);
  if (s)
    {
      TRETURN rfc2047_parse (s);
    }
  /* First try failed, now the "fun" of 2231 starts */
  std::string candidate = param + std::string("*");
  s = rfc822parse_query_parameter (field, candidate.c_str (), 0);
  if (s)
    {
      /* Language / Encoding but no continuation. RFC2047 parse handles this
         for us but needs to also see the asterisk.

         Example for this is:
         Content-Type: application/x-stuff;
         title*=us-ascii'en-us'This%20is%20%2A%2A%2Afun%2A%2A%2A
      */
      char *ret = rfc2047_parse (rfc2231_to_2047 (s).c_str());
      log_dbg ("Found lang parameter '%s' with value: '%s'",
               candidate.c_str (), anonstr (ret));
      TRETURN ret;
    }

  /* Now try the continued parameters. */
  std::string result;
  int i = 0;
  do
    {
      char *tmp = nullptr;
      candidate = param + std::string("*") + std::to_string(i++);
      s = rfc822parse_query_parameter (field, candidate.c_str (), 0);
      /* This is foo*0= */
      if (s)
        {
          tmp = rfc2047_parse (s);
        }
      else
        {
          /* This is continuation and language / encoding combined */
          candidate = candidate + std::string ("*");
          s = rfc822parse_query_parameter (field, candidate.c_str (), 0);
          if (!s)
            {
              break;
            }
          tmp = rfc2047_parse (rfc2231_to_2047 (s).c_str());
          log_dbg ("Found lang cont parameter '%s' with value: '%s'",
                   candidate.c_str (), anonstr (tmp));

        }
      result += tmp;
      xfree (tmp);
    } while (i < 50);

  if (i == 50)
    {
      log_warn ("Extremely deep parameter continuation. Aborting");
      TRETURN nullptr;
    }

  if (result.empty())
    {
      log_dbg ("Parameter '%s' not found.", param);
      TRETURN nullptr;
    }

  TRETURN xstrdup (result.c_str());
}

/* Process the transition to body event.

   This means we have received the empty line indicating the body and
   should now check the headers to see what to do about this part.

   This is mostly a C style function because it was based on the old
   c mimeparser.
*/
static int
t2body (MimeDataProvider *provider, rfc822parse_t msg)
{
  TSTART;
  rfc822parse_field_t field;
  mime_context_t ctx = provider->mime_context ();
  const char *ctmain, *ctsub;
  const char *s;
  size_t off;
  char *p;
  int is_text = 0;
  int is_text_attachment = 0;
  int is_protected_headers = 0;
  char *filename = NULL;
  char *cid = NULL;
  char *charset = NULL;
  bool ignore_cid = false;

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
      xfree (p);
    }

  /* Get the filename from the header.  */
  field = rfc822parse_parse_field (msg, "Content-Disposition", -1);
  if (field)
    {
      filename = rfc2231_query_parameter (field, "filename");

      s = rfc822parse_query_parameter (field, NULL, 1);
      if (s && strstr (s, "attachment"))
        {
          log_debug ("%s:%s: Found Content-Disposition attachment."
                     " Ignoring content-id to avoid hiding.",
                     SRCNAME, __func__);
          ignore_cid = true;
        }

      /* This is a bit of a taste matter how to treat inline
         attachments. Outlook does not show them inline so we
         should not put it in the body either as we have
         no way to show that it was actually an attachment.
         For something like an inline patch it is better
         to add it as an attachment instead of just putting
         it in the body.

         The handling in the old parser was:

         if (s && strcmp (s, "inline"))
           not_inline_text = 1;
       */
      if (ctx->body_seen)
        {
          /* Some MUA's like kontact e3.5 send the body as
             an inline text attachment. So if we have not
             seen the body yet we treat the first text/plain
             element as the body and not as an inline attachment. */
          is_text_attachment = 1;
        }
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

  log_debug ("%s:%s: ctx=%p, ct=`%s/%s'\n",
             SRCNAME, __func__, ctx, ctmain, ctsub);
  if (!ctx->nesting_level)
    {
      provider->set_content_type (ctmain, ctsub);
    }

  s = rfc822parse_query_parameter (field, "charset", 0);
  if (s)
    charset = xstrdup (s);

  if (!filename)
    {
      /* Check for Content-Type name if Content-Disposition filename
         was not found */
      filename = rfc2231_query_parameter (field, "name");

      if (!filename && !strcmp(ctmain, "message")&& !strcmp(ctsub,"rfc822"))
      {
        filename = xstrdup("rfc822_email.eml");
      }
    }

  /* Parse a Content Id header */
  if (!ignore_cid)
    {
      p = rfc822parse_get_field (msg, "Content-Id", -1, &off);
      if (p)
        {
           cid = xstrdup (p+off);
           xfree (p);
        }
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
    ms->cid = cid;
    filename = NULL;
    ms->charset = charset;
    charset = NULL;
  }

  if (!strcmp (ctmain, "multipart") || !strcmp (ctmain, "text"))
    {
      s = rfc822parse_query_parameter (field,
                                       "protected-headers", -1);
      if (s)
        {
          log_data ("%s:%s: Found protected headers: '%s'",
                           SRCNAME, __func__, s);
          if (!strncmp (s, "v", 1))
            {
              provider->m_protected_headers_version = atoi (s + 1);
            }
          is_protected_headers = 1;
        }
      else
        {
          is_protected_headers = 0;
        }
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
      else if (!strcmp (ctsub, "encrypted") &&
               (s = rfc822parse_query_parameter (field, "protocol", 0)))
        {
           if (!strcmp (s, "application/pgp-encrypted"))
             ctx->protocol = PROTOCOL_OPENPGP;
           /* We expect an encrypted mime part. */
           ctx->is_encrypted = 1;
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
      log_debug ("%s:%s: Collecting signature.",
                 SRCNAME, __func__);
    }
  else if (ctx->nesting_level == 1 && ctx->is_encrypted
           && !strcmp (ctmain, "application")
           && (ctx->protocol == PROTOCOL_OPENPGP
               && !strcmp (ctsub, "octet-stream")))
    {
      log_debug ("%s:%s: Collecting encrypted PGP data.",
                 SRCNAME, __func__);
      ctx->collect_crypto_data = 1;
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
          log_debug ("%s:%s: Collecting crypted S/MIME data.",
                     SRCNAME, __func__);
          ctx->collect_crypto_data = 1;
        }
    }

  /* Reset the in_data marker */
  ctx->in_data = 1;

  ctx->in_protected_headers = is_protected_headers;

  log_debug ("%s:%s: this body: nesting=%d partno=%d is_text=%d"
                   " charset=\"%s\"\n body_seen=%d is_text_attachment=%d"
                   " is_protected_headers=%d",
                   SRCNAME, __func__,
                   ctx->nesting_level, ctx->part_counter, is_text,
                   ctx->mimestruct_cur->charset?ctx->mimestruct_cur->charset:"",
                   ctx->body_seen, is_text_attachment, is_protected_headers);

  /* If this is a text part, decide whether we treat it as one
     of our bodies.
     */
  if ((is_text && !is_text_attachment))
    {
      if (is_text == 2)
        {
          ctx->body_seen = 2;
          ctx->collect_html_body = 1;
          ctx->collect_body = 0;
          log_debug ("%s:%s: Collecting HTML body.",
                     SRCNAME, __func__);
          /* We need this crutch because of one liner html
             mails which would not be collected by the line
             collector if they dont have a linefeed at the
             end. */
          provider->set_has_html_body (true);
        }
      else
        {
          log_debug ("%s:%s: Collecting text body.",
                     SRCNAME, __func__);
          ctx->body_seen = 1;
          ctx->collect_body = 1;
          ctx->collect_html_body = 0;
        }
    }
  else if (!ctx->collect_crypto_data)
    {
      bool isMultipart = (ctmain && !strcmp (ctmain, "multipart"));
      if (!ctx->nesting_level && !isMultipart)
        {
          log_dbg ("Data found that has no body. Treating it as attachment.");
        }
      else if (!ctx->nesting_level && isMultipart)
        {
          log_dbg ("Found first multipart transition");
        }
      else if (ctx->nesting_level || !isMultipart)
        {
          /* Treat it as an attachment.  */
          ctx->current_attachment = provider->create_attachment();
          ctx->collect_body = 0;
          ctx->collect_html_body = 0;
          log_debug ("%s:%s: Collecting attachment.",
                     SRCNAME, __func__);
        }
    }
  else
    {
      log_dbg ("Don't know what to collect, invalid mail?.");
    }
  rfc822parse_release_field (field); /* (Content-type) */

  TRETURN 0;
}

static int
message_cb (void *opaque, rfc822parse_event_t event,
            rfc822parse_t msg)
{
  int retval = 0;

  MimeDataProvider *provider = static_cast<MimeDataProvider*> (opaque);

  mime_context_t ctx = provider->mime_context();

  debug_message_event (event);

  if (event == RFC822PARSE_BEGIN_HEADER || event == RFC822PARSE_T2BODY)
    {
      /* We need to check here whether to start collecting signed data
         because attachments might come without header lines and thus
         we won't see the BEGIN_HEADER event. */
      if (ctx->start_hashing == 1)
        {
          ctx->start_hashing = 2;
          ctx->hashing_level = ctx->nesting_level;
          ctx->collect_crypto_data = 1;
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
      ctx->collect_body = 0;

      if (ctx->start_hashing == 2 && ctx->hashing_level == ctx->nesting_level)
        {
          ctx->start_hashing = 3; /* Avoid triggering it again. */
          ctx->collect_crypto_data = 0;
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

MimeDataProvider::MimeDataProvider(bool no_headers) :
  m_protected_headers_version(0),
  m_signature(nullptr),
  m_has_html_body(false),
  m_collect_everything(no_headers)
{
  TSTART;
  memdbg_ctor ("MimeDataProvider");
  m_mime_ctx = (mime_context_t) xcalloc (1, sizeof *m_mime_ctx);
  m_mime_ctx->msg = rfc822parse_open (message_cb, this);
  m_mime_ctx->mimestruct_tail = &m_mime_ctx->mimestruct;
  TRETURN;
}

#ifdef HAVE_W32_SYSTEM
MimeDataProvider::MimeDataProvider(LPSTREAM stream, bool no_headers):
  MimeDataProvider(no_headers)
{
  TSTART;
  if (stream)
    {
      stream->AddRef ();
      memdbg_addRef (stream);
    }
  else
    {
      log_error ("%s:%s called without stream ", SRCNAME, __func__);
      TRETURN;
    }
  log_data ("%s:%s Collecting data.", SRCNAME, __func__);
  collect_data (stream);
  log_data ("%s:%s Data collected.", SRCNAME, __func__);
  gpgol_release (stream);
  TRETURN;
}
#endif

MimeDataProvider::MimeDataProvider(FILE *stream, bool no_headers):
  MimeDataProvider(no_headers)
{
  TSTART;
  log_data ("%s:%s Collecting data from file.", SRCNAME, __func__);
  collect_data (stream);
  log_data ("%s:%s Data collected.", SRCNAME, __func__);
  TRETURN;
}

MimeDataProvider::~MimeDataProvider()
{
  TSTART;
  memdbg_dtor ("MimeDataProvider");
  log_debug ("%s:%s", SRCNAME, __func__);
  while (m_mime_ctx->mimestruct)
    {
      mimestruct_item_t tmp = m_mime_ctx->mimestruct->next;
      xfree (m_mime_ctx->mimestruct->filename);
      xfree (m_mime_ctx->mimestruct->charset);
      xfree (m_mime_ctx->mimestruct->cid);
      xfree (m_mime_ctx->mimestruct);
      m_mime_ctx->mimestruct = tmp;
    }
  rfc822parse_close (m_mime_ctx->msg);
  m_mime_ctx->current_attachment = NULL;
  xfree (m_mime_ctx);
  if (m_signature)
    {
      delete m_signature;
    }
  TRETURN;
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
  log_data ("%s:%s: Reading: " SIZE_T_FORMAT "Bytes",
                 SRCNAME, __func__, size);
  ssize_t bRead = m_crypto_data.read (buffer, size);
  if ((opt.enable_debug & DBG_DATA) && bRead)
    {
      std::string buf ((char *)buffer, bRead);

      if (!is_binary (buf))
        {
          log_data ("%s:%s: Data: \n------\n%s\n------",
                           SRCNAME, __func__, buf.c_str());
        }
      else
        {
          log_data ("%s:%s: Hex Data: \n------\n%s\n------",
                           SRCNAME, __func__,
                           string_to_hex (buf).c_str ());
        }
    }
  return bRead;
}

/* Split some raw data into lines and handle them accordingly.
   Returns the amount of bytes not taken from the input buffer.
*/
size_t
MimeDataProvider::collect_input_lines(const char *input, size_t insize)
{
  TSTART;
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
          TRETURN not_taken;
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

          log_data ("%s:%s: Parsing line=`%.*s'\n",
                    SRCNAME, __func__, (int)pos, linebuf);

#if 0 /* This is even too verbose for data debugging */
          log_dbg ("Parser state:\n"
                   "checked:      %d\n"
                   "col_body:     %d\n"
                   "crypt_data:   %d\n"
                   "hashing:      %d\n"
                   "in_data:      %d\n"
                   "signature:    %d\n"
                   "prot_headers: %d\n",
                   m_mime_ctx->pgp_marker_checked,
                   m_mime_ctx->collect_body,
                   m_mime_ctx->collect_crypto_data,
                   m_mime_ctx->start_hashing,
                   m_mime_ctx->in_data,
                   m_mime_ctx->collect_signature,
                   m_mime_ctx->in_protected_headers);
#endif
          /* Check the next state */
          if (rfc822parse_insert (m_mime_ctx->msg,
                                  (unsigned char*) linebuf,
                                  pos))
            {
              log_error ("%s:%s: rfc822 parser failed: %s\n",
                         SRCNAME, __func__, strerror (errno));
              TRETURN not_taken;
            }

          /* Check if the first line of the body is actually
             a PGP Inline message. If so treat it as crypto data. */
          if (!m_mime_ctx->pgp_marker_checked && m_mime_ctx->collect_body == 2)
            {
              m_mime_ctx->pgp_marker_checked = true;
              if (pos >= 27 && !strncmp ("-----BEGIN PGP MESSAGE-----", linebuf, 27))
                {
                  log_debug ("%s:%s: Found PGP Message in body.",
                             SRCNAME, __func__);
                  m_mime_ctx->collect_body = 0;
                  m_mime_ctx->collect_crypto_data = 1;
                  m_mime_ctx->start_hashing = 1;
                  m_collect_everything = true;
                }
            }

          /* If we are currently in a collecting state actually
             collect that line */
          if (m_mime_ctx->collect_crypto_data && m_mime_ctx->start_hashing)
            {
              /* Save the signed data.  Note that we need to delay
                 the CR/LF because the last line ending belongs to the
                 next boundary. */
              if (m_mime_ctx->collect_crypto_data == 2)
                {
                  m_crypto_data.write ("\r\n", 2);
                }
              log_data ("Writing raw crypto data: %.*s",
                               (int)pos, linebuf);
              m_crypto_data.write (linebuf, pos);
              m_mime_ctx->collect_crypto_data = 2;
            }
          if (m_mime_ctx->in_data && !m_mime_ctx->collect_signature &&
              !m_mime_ctx->collect_crypto_data)
            {
              /* We are inside of a plain part.  Write it out. */
              if (m_mime_ctx->in_data == 1)  /* Skip the first line. */
                m_mime_ctx->in_data = 2;

              int slbrk = 0;
              if (m_mime_ctx->is_qp_encoded)
                len = qp_decode (linebuf, pos, &slbrk);
              else if (m_mime_ctx->is_base64_encoded)
                len = b64_decode (&m_mime_ctx->base64, linebuf, pos);
              else
                len = pos;

              if (m_mime_ctx->collect_body)
                {
                  /* For protected headers to filter out the legacy display part
                     we have to first collect it in its own buffer and then later
                     decide if it should be hidden or not. Depending on the
                     reset of the mime structure. The legacy display part must
                     be either text/plain or text/rfc822-headers so we only
                     have to handle this case and not the HTML case below. */
                  if (m_mime_ctx->collect_body == 2)
                    {
                      std::string *target_buf =
                        (m_mime_ctx->in_protected_headers ? &m_ph_helpbuf : &m_body);
                      *target_buf += std::string(linebuf, len);
                      log_data ("Collecting as possibly protected header: %.*s",
                                (int)len, linebuf);
                      if (!m_mime_ctx->is_base64_encoded && !slbrk)
                        {
                          *target_buf += "\r\n";
                        }
                    }
                  if (m_body_charset.empty())
                    {
                      m_body_charset = m_mime_ctx->mimestruct_cur->charset ?
                                       m_mime_ctx->mimestruct_cur->charset : "";
                    }
                  m_mime_ctx->collect_body = 2;
                }
              else if (m_mime_ctx->collect_html_body)
                {
                  if (m_mime_ctx->collect_html_body == 2)
                    {
                      m_html_body += std::string(linebuf, len);
                      if (!m_mime_ctx->is_base64_encoded && !slbrk)
                        {
                          m_html_body += "\r\n";
                        }
                    }
                  if (m_html_charset.empty())
                    {
                      m_html_charset = m_mime_ctx->mimestruct_cur->charset ?
                                       m_mime_ctx->mimestruct_cur->charset : "";
                    }
                  m_mime_ctx->collect_html_body = 2;
                }
              else if (m_mime_ctx->current_attachment && len)
                {
                  m_mime_ctx->current_attachment->get_data().write(linebuf, len);
                  if (!m_mime_ctx->is_base64_encoded && !slbrk)
                    {
                      m_mime_ctx->current_attachment->get_data().write("\r\n", 2);
                    }
                }
              else
                {
                  log_debug ("%s:%s Collecting finished.",
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
          else if (m_mime_ctx->in_data && !m_mime_ctx->start_hashing)
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
              log_data ("Writing crypto data: %.*s",
                         (int)pos, linebuf);
              if (len)
                m_crypto_data.write(linebuf, len);
              if (!m_mime_ctx->is_base64_encoded && !slbrk)
                m_crypto_data.write("\r\n", 2);
            }
          /* Continue with next line. */
          pos = 0;
        }
    }
  TRETURN not_taken;
}

#ifdef HAVE_W32_SYSTEM
void
MimeDataProvider::collect_data(LPSTREAM stream)
{
  TSTART;
  if (!stream)
    {
      TRETURN;
    }
  HRESULT hr;
  char buf[BUFSIZE];
  ULONG bRead;
  bool first_read = true;
  bool is_pgp_message = false;
  size_t allRead = 0;
  while ((hr = stream->Read (buf, BUFSIZE, &bRead)) == S_OK ||
         hr == S_FALSE)
    {
      if (!bRead)
        {
          log_debug ("%s:%s: Input stream at EOF.",
                     SRCNAME, __func__);
          break;
        }
      log_debug ("%s:%s: Read %lu bytes.",
                       SRCNAME, __func__, bRead);
      allRead += bRead;
      if (first_read)
        {
          if (bRead > 12 && strncmp ("MIME-Version", buf, 12) == 0)
            {
              /* Fun! In case we have exchange or sent messages created by us
                 we get the mail attachment like it is before the MAPI to MIME
                 conversion. So it has our MIME structure. In that case
                 we have to expect MIME data even if the initial data check
                 suggests that we don't.

                 Checking if the content starts with MIME-Version appears
                 to be a robust way to check if we try to parse MIME data. */
              m_collect_everything = false;
              log_debug ("%s:%s: Found MIME-Version marker."
                         "Expecting headers even if type suggested not to.",
                         SRCNAME, __func__);

            }
          else if (bRead > 12 && !strncmp ("Content-Type:", buf, 13))
            {
              /* Similar as above but we messed with the order of the headers
                 for some s/mime mails. So also check for content type.

                 Want some cheese with that hack?
              */
              m_collect_everything = false;
              log_debug ("%s:%s: Found Content-Type header."
                         "Expecting headers even if type suggested not to.",
                         SRCNAME, __func__);

            }
          /* check for the PGP MESSAGE marker to see if we have it. */
          if (bRead && m_collect_everything)
            {
              std::string tmp (buf, bRead);
              std::size_t found = tmp.find ("-----BEGIN PGP MESSAGE-----");
              if (found != std::string::npos)
                {
                  log_debug ("%s:%s: found PGP Message marker,",
                             SRCNAME, __func__);
                  is_pgp_message = true;
                }
            }
        }
      first_read = false;

      if (m_collect_everything)
        {
          /* For S/MIME, Clearsigned, PGP MESSAGES we just pass everything
             on. Only the Multipart classes need parsing. And the output
             of course. */
          log_debug ("%s:%s: Just copying data.",
                     SRCNAME, __func__);
          m_crypto_data.write ((void*)buf, (size_t) bRead);
          continue;
        }
      m_rawbuf += std::string (buf, bRead);
      size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                              m_rawbuf.size());

      if (not_taken == m_rawbuf.size())
        {
          log_error ("%s:%s: Collect failed to consume anything.\n"
                     "Buffer too small?",
                     SRCNAME, __func__);
          break;
        }
      log_debug ("%s:%s: Consumed: " SIZE_T_FORMAT " bytes",
                 SRCNAME, __func__, m_rawbuf.size() - not_taken);
      m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
    }


  if (is_pgp_message && allRead < (1024 * 100))
    {
      /* Sometimes received PGP Messsages contain extra whitespace /
         newlines. To also accept such messages we fix up pgp inline
         messages here. We only do this for messages which are smaller
         then a hundred KByte for performance. */
      log_debug ("%s:%s: Fixing up a possible broken message.",
                 SRCNAME, __func__);
      /* Copy crypto data to string */
      std::string data = m_crypto_data.toString();
      m_crypto_data = GpgME::Data();
      std::istringstream iss (data);
      // Now parse it by line.
      std::string line;
      while (std::getline (iss, line))
        {
          trim (line);
          if (line == "-----BEGIN PGP MESSAGE-----")
            {
              /* Finish an armor header */
              line += "\n\n";
              m_crypto_data.write (line.c_str (), line.size ());
              continue;
            }
          /* Remove empty lines */
          if (line.empty())
            {
              continue;
            }
          if (line.find (':') != std::string::npos)
            {
              log_data ("%s:%s: Removing comment '%s'.",
                               SRCNAME, __func__, line.c_str ());
              continue;
            }
          line += '\n';
          m_crypto_data.write (line.c_str (), line.size ());
        }
    }
  TRETURN;
}
#endif

void
MimeDataProvider::collect_data(FILE *stream)
{
  TSTART;
  if (!stream)
    {
      TRETURN;
    }
  char buf[BUFSIZE];
  size_t bRead;
  while ((bRead = fread (buf, 1, BUFSIZE, stream)) > 0)
    {
      log_debug ("%s:%s: Read " SIZE_T_FORMAT " bytes.",
                 SRCNAME, __func__, bRead);

      if (m_collect_everything)
        {
          /* For S/MIME, Clearsigned, PGP MESSAGES we just pass everything
             on. Only the Multipart classes need parsing. And the output
             of course. */
          log_debug ("%s:%s: Making verbatim copy" SIZE_T_FORMAT " bytes.",
                     SRCNAME, __func__, bRead);
          m_crypto_data.write ((void*)buf, bRead);
          continue;
        }
      m_rawbuf += std::string (buf, bRead);
      size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                              m_rawbuf.size());

      if (not_taken == m_rawbuf.size())
        {
          log_error ("%s:%s: Collect failed to consume anything.\n"
                     "Buffer too small?",
                     SRCNAME, __func__);
          TRETURN;
        }
      log_debug ("%s:%s: Consumed: " SIZE_T_FORMAT " bytes",
                 SRCNAME, __func__, m_rawbuf.size() - not_taken);
      m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
    }
  TRETURN;
}

ssize_t MimeDataProvider::write(const void *buffer, size_t bufSize)
{
  TSTART;
  if (m_collect_everything)
    {
      /* Writing with collect everything one means that we are outputprovider.
         In this case for inline messages we want to collect everything. */
      log_debug ("%s:%s: Using complete input as body " SIZE_T_FORMAT " bytes.",
                 SRCNAME, __func__, bufSize);
      m_body += std::string ((const char *) buffer, bufSize);
      TRETURN bufSize;
    }
  m_rawbuf += std::string ((const char*)buffer, bufSize);
  size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                          m_rawbuf.size());

  if (not_taken == m_rawbuf.size())
    {
      log_error ("%s:%s: Write failed to consume anything.\n"
                 "Buffer too small? or no newlines in text?",
                 SRCNAME, __func__);
      TRETURN bufSize;
    }
  log_debug ("%s:%s: Write Consumed: " SIZE_T_FORMAT " bytes",
                   SRCNAME, __func__, m_rawbuf.size() - not_taken);
  m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
  TRETURN bufSize;
}

off_t
MimeDataProvider::seek(off_t offset, int whence)
{
  return m_crypto_data.seek (offset, whence);
}

GpgME::Data *
MimeDataProvider::signature() const
{
  TSTART;
  TRETURN m_signature;
}

std::shared_ptr<Attachment>
MimeDataProvider::create_attachment()
{
  TSTART;
  log_debug ("%s:%s: Creating attachment.",
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
          log_debug ("%s:%s: Attachment filename: %s",
                     SRCNAME, __func__, anonstr (m_mime_ctx->mimestruct_cur->filename));
          attach->set_display_name (m_mime_ctx->mimestruct_cur->filename);
        }
    }
  if (m_mime_ctx->mimestruct_cur && m_mime_ctx->mimestruct_cur->cid)
    {
      attach->set_content_id (m_mime_ctx->mimestruct_cur->cid);
      log_debug  ("%s:%s: content-id: %s",
                  SRCNAME, __func__, anonstr (m_mime_ctx->mimestruct_cur->cid));
    }
  if (m_mime_ctx->mimestruct_cur && m_mime_ctx->mimestruct_cur->content_type)
    {
      attach->set_content_type (m_mime_ctx->mimestruct_cur->content_type);
      log_debug ("%s:%s: content-type: %s",
                 SRCNAME, __func__, m_mime_ctx->mimestruct_cur->content_type);
    }
  m_attachments.push_back (attach);

  TRETURN attach;
  /* TODO handle encoding */
}

static std::string
get_header (rfc822parse_t msg, const std::string &which)
{
  TSTART;
  const char *buf = rfc822parse_get_field (msg,
                                           which.c_str (), -1,
                                           nullptr);
  if (buf)
    {
      /* String is "which: " so the + 3 is the colon, the space and one
         extra to ensure that we do not construct a std::string from null
         below. */
      if (strlen (buf) <= which.size() + 3)
        {
          log_error ("%s:%s: Invalid header value '%s'", SRCNAME, __func__,
                     buf);
          TRETURN std::string ();
        }
      std::string ret = buf + which.size () + 2;
      if (ret.size())
        {
          ret = rfc2047_parse (ret.c_str ());
        }
      TRETURN ret;
    }
  TRETURN std::string ();
}

void MimeDataProvider::finalize ()
{
  TSTART;

  if (m_rawbuf.size ())
    {
      m_rawbuf += "\r\n";
      size_t not_taken = collect_input_lines (m_rawbuf.c_str(),
                                              m_rawbuf.size());
      m_rawbuf.erase (0, m_rawbuf.size() - not_taken);
      if (m_rawbuf.size ())
        {
          log_error ("%s:%s: Collect left data in buffer.\n",
                     SRCNAME, __func__);
        }
    }

  static std::vector<std::string> user_headers = {"Subject", "From",
                                                  "To", "Cc", "Date",
                                                  "Reply-To",
                                                  "Followup-To"};
  if (m_protected_headers_version)
    {
      for (const auto &hdr: user_headers)
        {
          m_protected_headers.emplace (hdr, get_header (m_mime_ctx->msg, hdr));
        }
    }
  /* Now check the mime strucutre for that legacy-display handling of
     protected headers.

     If the mime structure is:
     multipart/mixed and the first subpart is text/plain or text/rfc822-headers
     that we hide the first text part as we parsed the headers above
     from that part. */
  if (m_protected_headers_version == 1 && m_ph_helpbuf.size () &&
      m_mime_ctx->mimestruct && m_mime_ctx->mimestruct->content_type &&
      !strcmp (m_mime_ctx->mimestruct->content_type, "multipart/mixed") &&
      m_mime_ctx->mimestruct->next && m_mime_ctx->mimestruct->next->content_type && (
      !strcmp (m_mime_ctx->mimestruct->next->content_type, "text/plain") ||
      !strcmp (m_mime_ctx->mimestruct->next->content_type, "text/rfc822-headers")))
    {
      log_debug ("%s:%s: Detected protected headers legacy part. It will be hidden.",
                 SRCNAME, __func__);
      if (m_protected_headers.empty ())
        {
          /* Do a simple parsing of the header data that is displayed. */
          log_data ("Parsing headers from the legacy part.");
          std::istringstream ss(m_ph_helpbuf);
          std::string line;
          while (std::getline (ss, line))
            {
              for (const auto &hdr: user_headers)
                {
                  const std::string needle = hdr + std::string (": ");
                  if (starts_with (line, needle.c_str ()))
                    {
                      log_data ("Found line starting with: %s", needle.c_str ());
                      find_and_replace (line, needle, std::string ());
                      m_protected_headers.emplace (hdr, line);
                    }
                }
            }
        }
      log_data ("%s:%s: PH legacy part: '%s'", SRCNAME, __func__, m_ph_helpbuf.c_str ());
    }
  else if (m_ph_helpbuf.size ())
    {
      log_debug ("%s:%s: Prepending protected headers part to buffer.",
                 SRCNAME, __func__);
      m_body = m_ph_helpbuf + m_body;
    }
  TRETURN;
}

const std::string &MimeDataProvider::get_body ()
{
  TSTART;
  TRETURN m_body;
}

const std::string &MimeDataProvider::get_html_body ()
{
  TSTART;
  TRETURN m_html_body;
}

const std::string &MimeDataProvider::get_html_charset() const
{
  TSTART;
  TRETURN m_html_charset;
}

const std::string &MimeDataProvider::get_body_charset() const
{
  TSTART;
  TRETURN m_body_charset;
}

std::string
MimeDataProvider::get_protected_header (const std::string &which) const
{
  TSTART;
  const auto it = m_protected_headers.find (which);
  if (it != m_protected_headers.end ())
    {
      TRETURN it->second;
    }
  TRETURN std::string ();
}

std::string
MimeDataProvider::get_content_type () const
{
  return m_content_type;
}

void
MimeDataProvider::set_content_type (const char *ctmain, const char *ctsub)
{
  std::string main = ctmain ? std::string (ctmain) : std::string ();
  std::string sub = ctsub ? std::string ("/") + std::string (ctsub) :
                                                std::string ();
  m_content_type = main + sub;
}
