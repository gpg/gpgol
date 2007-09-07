/* mimemaker.c - Construct MIME message out of a MAPI
 *	Copyright (C) 2007 g10 Code GmbH
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

#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include <string.h>
#include <ctype.h>

#define COBJMACROS
#include <windows.h>
#include <objidl.h> 

#include "mymapi.h"
#include "mymapitags.h"

#include "common.h"
#include "engine.h"
#include "mapihelp.h"
#include "mimemaker.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

static const char oid_mimetag[] =
    {0x2A, 0x86, 0x48, 0x86, 0xf7, 0x14, 0x03, 0x0a, 0x04};

/* The base-64 list used for base64 encoding. */
static unsigned char bintoasc[64+1] = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
                                       "abcdefghijklmnopqrstuvwxyz" 
                                       "0123456789+/"); 

/* The maximum length of a line we are able to process.  RFC822 allows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000

/* Make sure that PROTOCOL is usable or return a suitable protocol.
   On error PROTOCOL_UNKNOWN is returned.  */
static protocol_t
check_protocol (protocol_t protocol)
{
  switch (protocol)
    {
    case PROTOCOL_UNKNOWN:
      log_error ("fixme: automatic protocol selection is not yet supported");
      return PROTOCOL_UNKNOWN;
    case PROTOCOL_OPENPGP:
    case PROTOCOL_SMIME:
      return protocol;
    }

  log_error ("%s:%s: BUG", SRCNAME, __func__);
  return PROTOCOL_UNKNOWN;
}



/* Create a new MAPI attchment for MESSAGE which will be used to
   prepare the MIME message.  On sucess the stream to write the data
   to is stored at STREAM and the attchment object itself is the
   retruned.  The caller needs to call SaveChanges.  Returns NULL on
   failure in which case STREAM will be set to NULL.  */
static LPATTACH
create_mapi_attachment (LPMESSAGE message, LPSTREAM *stream)
{
  HRESULT hr;
  ULONG pos;
  SPropValue prop;
  LPATTACH att = NULL;
  LPUNKNOWN punk;

  *stream = NULL;
  hr = IMessage_CreateAttach (message, NULL, 0, &pos, &att);
  if (hr)
    {
      log_error ("%s:%s: can't create attachment: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      return NULL;
    }

  prop.ulPropTag = PR_ATTACH_METHOD;
  prop.Value.ul = ATTACH_BY_VALUE;
  hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach method: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }

  /* Mark that attachment so that we know why it has been created.  */
  if (get_gpgolattachtype_tag (message, &prop.ulPropTag) )
    goto failure;
  prop.Value.l = ATTACHTYPE_MOSSTEMPL;
  hr = HrSetOneProp ((LPMAPIPROP)att, &prop);	
  if (hr)
    {
      log_error ("%s:%s: can't set %s property: hr=%#lx\n",
                 SRCNAME, __func__, "GpgOL Attach Type", hr); 
      goto failure;
    }


  /* We better insert a short filename. */
  prop.ulPropTag = PR_ATTACH_FILENAME_A;
  prop.Value.lpszA = "gpgolXXX.dat";
  hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
  if (hr)
    {
      log_error ("%s:%s: can't set attach filename: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }


  /* Even for encrypted messages we need to set the MAPI property to
     multipart/signed.  This seems to be a part of the trigger which
     leads OL to process such a message in a special way.  */
  prop.ulPropTag = PR_ATTACH_TAG;
  prop.Value.bin.cb  = sizeof oid_mimetag;
  prop.Value.bin.lpb = (LPBYTE)oid_mimetag;
  hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
  if (!hr)
    {
      prop.ulPropTag = PR_ATTACH_MIME_TAG_A;
      prop.Value.lpszA = "multipart/signed";
      hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
    }
  if (hr)
    {
      log_error ("%s:%s: can't set attach mime tag: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }
  
  punk = NULL;
  hr = IAttach_OpenProperty (att, PR_ATTACH_DATA_BIN, &IID_IStream, 0,
                             (MAPI_CREATE|MAPI_MODIFY), &punk);
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't create output stream: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }
  *stream = (LPSTREAM)punk;
  return att;

 failure:
  IAttach_Release (att);
  return NULL;
}


/* Wrapper around IStream::Write to print an error message.  */
static int 
write_buffer (LPSTREAM stream, const void *data, size_t datalen)
{
  HRESULT hr;

  hr = IStream_Write (stream, data, datalen, NULL);
  if (hr)
    {
      log_error ("%s:%s: Write failed: hr=%#lx", SRCNAME, __func__, hr);
      return -1;
    }
  return 0;
}


/* Write the string TEXT to the IStream STREAM.  Returns 0 on sucsess,
   prints an error message and returns -1 on error.  */
static int 
write_string (LPSTREAM stream, const char *text)
{
  return write_buffer (stream, text, strlen (text));
}


/* Helper to write a boundary to the output stream.  The leading LF
   will be written as well.  */
static int
write_boundary (LPSTREAM stream, const char *boundary, int lastone)
{
  int rc = write_string (stream, "\r\n--");
  if (!rc)
    rc = write_string (stream, boundary);
  if (!rc)
    rc = write_string (stream, lastone? "--\r\n":"\r\n");
  return rc;
}


/* Write DATALEN bytes of DATA to STREAM in base64 encoding.  This
   creates a complete Base64 chunk including the trailing fillers.  */
static int
write_b64 (LPSTREAM stream, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  unsigned char inbuf[4];
  int idx, quads;
  char outbuf[4];

  idx = quads = 0;
  for (p = data; datalen; p++, datalen--)
    {
      inbuf[idx++] = *p;
      if (idx > 2)
        {
          outbuf[0] = bintoasc[(*inbuf>>2)&077];
          outbuf[1] = bintoasc[(((*inbuf<<4)&060)|((inbuf[1] >> 4)&017))&077];
          outbuf[2] = bintoasc[(((inbuf[1]<<2)&074)|((inbuf[2]>>6)&03))&077];
          outbuf[3] = bintoasc[inbuf[2]&077];
          if ((rc = write_buffer (stream, outbuf, 4)))
            return rc;
          idx = 0;
          if (++quads >= (64/4)) 
            {
              quads = 0;
              if ((rc = write_buffer (stream, "\r\n", 2)))
                return rc;
            }
        }
    }

  if (idx)
    {
      outbuf[0] = bintoasc[(*inbuf>>2)&077];
      if (idx == 1)
        {
          outbuf[1] = bintoasc[((*inbuf<<4)&060)&077];
          outbuf[2] = '=';
          outbuf[3] = '=';
        }
      else 
        { 
          outbuf[1] = bintoasc[(((*inbuf<<4)&060)|((inbuf[1]>>4)&017))&077];
          outbuf[2] = bintoasc[((inbuf[1]<<2)&074)&077];
          outbuf[3] = '=';
        }
      if ((rc = write_buffer (stream, outbuf, 4)))
        return rc;
      ++quads;
    }

  if (quads) 
    if ((rc = write_buffer (stream, "\r\n", 2)))
      return rc;

  return 0;
}

/* Write DATALEN bytes of DATA to STREAM in quoted-prinable encoding. */
static int
write_qp (LPSTREAM stream, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  char outbuf[80];  /* We only need 76 octect + 2 for the lineend. */
  int outidx;

  /* Check whether the current character is followed by a line ending.
     Note that the end of the etxt also counts as a lineending */
#define nextlf_p() ((datalen > 2 && p[1] == '\r' && p[2] == '\n') \
                    || (datalen > 1 && p[1] == '\n')              \
                    || datalen == 1 )

  /* Macro to insert a soft line break if needed.  */
# define do_softlf(n) \
          do {                                                        \
            if (outidx + (n) > 76                                     \
                || (outidx + (n) == 76 && !nextlf_p()))               \
              {                                                       \
                outbuf[outidx++] = '=';                               \
                outbuf[outidx++] = '\r';                              \
                outbuf[outidx++] = '\n';                              \
                if ((rc = write_buffer (stream, outbuf, outidx)))     \
                  return rc;                                          \
                outidx = 0;                                           \
              }                                                       \
          } while (0)
              
  outidx = 0;
  for (p = data; datalen; p++, datalen--)
    {
      if ((datalen > 1 && *p == '\r' && p[1] == '\n') || *p == '\n')
        {
          /* Line break.  */
          outbuf[outidx++] = '\r';
          outbuf[outidx++] = '\n';
          if ((rc = write_buffer (stream, outbuf, outidx)))
            return rc;
          outidx = 0;
          if (*p == '\r')
            {
              p++;
              datalen--;
            }
        }
      else if (*p == '\t' || *p == ' ')
        {
          /* Check whether tab or space is followed by a line break
             which forbids verbatim encoding.  If we are already at
             the end of the buffer we take that as a line end too. */
          if (nextlf_p())
            {
              do_softlf (3);
              outbuf[outidx++] = '=';
              outbuf[outidx++] = tohex ((*p>>4)&15);
              outbuf[outidx++] = tohex (*p&15);
            }
          else
            {
              do_softlf (1);
              outbuf[outidx++] = *p;
            }

        }
      else if (!outidx && *p == '.' && nextlf_p () )
        {
          /* We better protect a line with just a single dot.  */
          outbuf[outidx++] = '=';
          outbuf[outidx++] = tohex ((*p>>4)&15);
          outbuf[outidx++] = tohex (*p&15);
        }
      else if (*p >= '!' && *p <= '~' && *p != '=')
        {
          do_softlf (1);
          outbuf[outidx++] = *p;
        }
      else
        {
          do_softlf (3);
          outbuf[outidx++] = '=';
          outbuf[outidx++] = tohex ((*p>>4)&15);
          outbuf[outidx++] = tohex (*p&15);
        }
    }
  if (outidx)
    {
      outbuf[outidx++] = '\r';
      outbuf[outidx++] = '\n';
      if ((rc = write_buffer (stream, outbuf, outidx)))
        return rc;
    }

# undef do_softlf
# undef nextlf_p
  return 0;
}


/* Write DATALEN bytes of DATA to STREAM in plain ascii encoding. */
static int
write_plain (LPSTREAM stream, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  char outbuf[100];
  int outidx;

  outidx = 0;
  for (p = data; datalen; p++, datalen--)
    {
      if ((datalen > 1 && *p == '\r' && p[1] == '\n') || *p == '\n')
        {
          outbuf[outidx++] = '\r';
          outbuf[outidx++] = '\n';
          if ((rc = write_buffer (stream, outbuf, outidx)))
            return rc;
          outidx = 0;
          if (*p == '\r')
            {
              p++;
              datalen--;
            }
        }
      else if (!outidx && *p == '.'
               && ( (datalen > 2 && p[1] == '\r' && p[2] == '\n') 
                    || (datalen > 1 && p[1] == '\n') 
                    || datalen == 1))
        {
          /* Better protect a line with just a single dot.  We do
             this by adding a space.  */
          outbuf[outidx++] = *p;
          outbuf[outidx++] = ' ';
        }
      else if (outidx > 80)
        {
          /* We should never be called for too long lines - QP should
             have been used.  */
          log_error ("%s:%s: BUG: line longer than exepcted",
                     SRCNAME, __func__);
          return -1; 
        }
      else
        outbuf[outidx++] = *p;
    }

  if (outidx)
    {
      outbuf[outidx++] = '\r';
      outbuf[outidx++] = '\n';
      if ((rc = write_buffer (stream, outbuf, outidx)))
        return rc;
    }

  return 0;
}


/* Infer the conent type from DATA and FILENAME.  The return value is
   a static string there won't be an error return.  In case Bae 64
   encoding is required for the type true will be stored at FORCE_B64;
   however, this is only a shortcut and if that is not set, the caller
   should infer the encoding by otehr means. */
static const char *
infer_content_type (const char *data, size_t datalen, const char *filename,
                    int is_mapibody, int *force_b64)
{
  static struct {
    char b64;
    const char *suffix;
    const char *ct;
  } suffix_table[] = 
    {
      { 1, "3gp",   "video/3gpp" },
      { 1, "abw",   "application/x-abiword" },
      { 1, "ai",    "application/postscript" },
      { 1, "au",    "audio/basic" },
      { 1, "bin",   "application/octet-stream" },
      { 1, "class", "application/java-vm" },
      { 1, "cpt",   "application/mac-compactpro" },
      { 0, "css",   "text/css" },
      { 0, "csv",   "text/comma-separated-values" },
      { 1, "deb",   "application/x-debian-package" },
      { 1, "dl",    "video/dl" },
      { 1, "doc",   "application/msword" },
      { 1, "dv",    "video/dv" },
      { 1, "dvi",   "application/x-dvi" },
      { 1, "eml",   "message/rfc822" },
      { 1, "eps",   "application/postscript" },
      { 1, "fig",   "application/x-xfig" },
      { 1, "flac",  "application/x-flac" },
      { 1, "fli",   "video/fli" },
      { 1, "gif",   "image/gif" },
      { 1, "gl",    "video/gl" },
      { 1, "gnumeric", "application/x-gnumeric" },
      { 1, "hqx",   "application/mac-binhex40" },
      { 1, "hta",   "application/hta" },
      { 0, "htm",   "text/html" },
      { 0, "html",  "text/html" },
      { 0, "ics",   "text/calendar" },
      { 1, "jar",   "application/java-archive" },
      { 1, "jpeg",  "image/jpeg" },
      { 1, "jpg",   "image/jpeg" },
      { 1, "js",    "application/x-javascript" },
      { 1, "latex", "application/x-latex" },
      { 1, "lha",   "application/x-lha" },
      { 1, "lzh",   "application/x-lzh" },
      { 1, "lzx",   "application/x-lzx" },
      { 1, "m3u",   "audio/mpegurl" },
      { 1, "m4a",   "audio/mpeg" },
      { 1, "mdb",   "application/msaccess" },
      { 1, "midi",  "audio/midi" },
      { 1, "mov",   "video/quicktime" },
      { 1, "mp2",   "audio/mpeg" },
      { 1, "mp3",   "audio/mpeg" },
      { 1, "mp4",   "video/mp4" },
      { 1, "mpeg",  "video/mpeg" },
      { 1, "mpega", "audio/mpeg" },
      { 1, "mpg",   "video/mpeg" },
      { 1, "mpga",  "audio/mpeg" },
      { 1, "msi",   "application/x-msi" },
      { 1, "mxu",   "video/vnd.mpegurl" },
      { 1, "nb",    "application/mathematica" },
      { 1, "oda",   "application/oda" },
      { 1, "odb",   "application/vnd.oasis.opendocument.database" },
      { 1, "odc",   "application/vnd.oasis.opendocument.chart" },
      { 1, "odf",   "application/vnd.oasis.opendocument.formula" },
      { 1, "odg",   "application/vnd.oasis.opendocument.graphics" },
      { 1, "odi",   "application/vnd.oasis.opendocument.image" },
      { 1, "odm",   "application/vnd.oasis.opendocument.text-master" },
      { 1, "odp",   "application/vnd.oasis.opendocument.presentation" },
      { 1, "ods",   "application/vnd.oasis.opendocument.spreadsheet" },
      { 1, "odt",   "application/vnd.oasis.opendocument.text" },
      { 1, "ogg",   "application/ogg" },
      { 1, "otg",   "application/vnd.oasis.opendocument.graphics-template" },
      { 1, "oth",   "application/vnd.oasis.opendocument.text-web" },
      { 1, "otp",  "application/vnd.oasis.opendocument.presentation-template"},
      { 1, "ots",   "application/vnd.oasis.opendocument.spreadsheet-template"},
      { 1, "ott",   "application/vnd.oasis.opendocument.text-template" },
      { 1, "pdf",   "application/pdf" },
      { 1, "png",   "image/png" },
      { 1, "pps",   "application/vnd.ms-powerpoint" },
      { 1, "ppt",   "application/vnd.ms-powerpoint" },
      { 1, "prf",   "application/pics-rules" },
      { 1, "ps",    "application/postscript" },
      { 1, "qt",    "video/quicktime" },
      { 1, "rar",   "application/rar" },
      { 1, "rdf",   "application/rdf+xml" },
      { 1, "rpm",   "application/x-redhat-package-manager" },
      { 0, "rss",   "application/rss+xml" },
      { 1, "ser",   "application/java-serialized-object" },
      { 0, "sh",    "application/x-sh" },
      { 0, "shtml", "text/html" },
      { 1, "sid",   "audio/prs.sid" },
      { 0, "smil",  "application/smil" },
      { 1, "snd",   "audio/basic" },
      { 0, "svg",   "image/svg+xml" },
      { 1, "tar",   "application/x-tar" },
      { 0, "texi",  "application/x-texinfo" },
      { 0, "texinfo", "application/x-texinfo" },
      { 1, "tif",   "image/tiff" },
      { 1, "tiff",  "image/tiff" },
      { 1, "torrent", "application/x-bittorrent" },
      { 1, "tsp",   "application/dsptype" },
      { 0, "vrml",  "model/vrml" },
      { 1, "vsd",   "application/vnd.visio" },
      { 1, "wp5",   "application/wordperfect5.1" },
      { 1, "wpd",   "application/wordperfect" },
      { 0, "xhtml", "application/xhtml+xml" },
      { 1, "xlb",   "application/vnd.ms-excel" },
      { 1, "xls",   "application/vnd.ms-excel" },
      { 1, "xlt",   "application/vnd.ms-excel" },
      { 0, "xml",   "application/xml" },
      { 0, "xsl",   "application/xml" },
      { 0, "xul",   "application/vnd.mozilla.xul+xml" },
      { 1, "zip",   "application/zip" },
      { 0, NULL, NULL }
    };
  int i;
  char suffix_buffer[12+1];
  const char *suffix;

  *force_b64 = 0;
  suffix = filename? strrchr (filename, '.') : NULL;
  if (suffix && strlen (suffix) < sizeof suffix_buffer -1 )
    {
      suffix++;
      for (i=0; i < sizeof suffix_buffer - 1; i++)
        suffix_buffer[i] = tolower (*(const unsigned char*)suffix);
      suffix_buffer[i] = 0;
      for (i=0; suffix_table[i].suffix; i++)
        if (!strcmp (suffix_table[i].suffix, suffix_buffer))
          {
            if (suffix_table[i].b64)
              *force_b64 = 1;
            return suffix_table[i].ct;
          }
    }

  /* Not found via filename, look at the content.  */

  if (is_mapibody)
    {
      /* Fixme:  This is too simple. */
      if (datalen > 6  && (!memcmp (data, "<html>", 6)
                           ||!memcmp (data, "<HTML>", 6)))
        return "text/html";
      return "text/plain";
    }

  return "application/octet-stream";
}

/* Figure out the best encoding to be used for the part.  Return values are
     0: Plain ASCII.
     1: Quoted Printable
     2: Base64  */
static const int
infer_content_encoding (const void *data, size_t datalen)
{
  const unsigned char *p;
  int need_qp;
  size_t len, maxlen, highbin, lowbin, ntotal;

  ntotal = datalen;
  len = maxlen = lowbin = highbin = 0;
  need_qp = 0;
  for (p = data; datalen; p++, datalen--)
    {
      len++;
      if ((*p & 0x80))
        highbin++;
      else if ((datalen > 1 && *p == '\r' && p[1] == '\n') || *p == '\n')
        {
          len--;
          if (len > maxlen)
            maxlen = len;
          len = 0;
        }
      else if (*p == '\r')
        {
          /* CR not followed by a linefeed. */
          lowbin++;
        }
      else if (*p == '\t' || *p == ' ' || *p == '\f')
        ;
      else if (*p < ' ' || *p == 127)
        lowbin++;
      else if (len == 1 && datalen > 2
               && *p == '-' && p[1] == '-' && p[2] == ' '
               && ( (datalen > 4 && p[3] == '\r' && p[4] == '\n') 
                    || (datalen > 3 && p[3] == '\n') 
                    || datalen == 3))
        {
          /* This is a "-- \r\n" line, thus it indicates the usual
             signature line delimiter.  We need to protect the
             trailing space.  */
          need_qp = 1;
        }
      else if (len == 1 && datalen > 5 && !memcmp (p, "--=-=", 5))
        {
          /* This look pretty much like a our own boundary.
             We better protect it by forcing QP encoding.  */
          need_qp = 1;
        }
    }
  if (len > maxlen)
    maxlen = len;

  if (maxlen <= 76 && !lowbin && !highbin && !need_qp)
    return 0; /* Plain ASCII is sufficient.  */

  /* Somewhere in the Outlook documentation 20% is mentioned as
     discriminating value for Base64.  Though our counting won't be
     identical we use that value to behave closely to it. */
  if (ntotal && ((float)(lowbin+highbin))/ntotal < 0.20)
    return 1; /* Use quoted printable.  */
  
  return 2;   /* Use base64.  */
}






/* Write a MIME part to STREAM.  The BOUNDARY is written first the
   DATA is analyzed and appropriate headers are written.  If FILENAME
   is given it will be added to the part's header. IS_MAPIBODY should
   be true if teh data has been retrieved from the body property. */
static int
write_part (LPSTREAM stream, const char *data, size_t datalen,
            const char *boundary, const char *filename, int is_mapibody)
{
  int rc;
  const char *ct;
  int use_b64, use_qp, is_text;

  ct = infer_content_type (data, datalen, filename, is_mapibody, &use_b64);
  use_qp = 0;
  if (!use_b64)
    {
      switch (infer_content_encoding (data, datalen))
        {
        case 0: break;
        case 1: use_qp = 1; break;
        default: use_b64 = 1; break;
        }
    }
  is_text = !strncmp (ct, "text/", 5);

  if ((rc = write_boundary (stream, boundary, 0)))
    return rc;
  if (!(rc = write_string (stream, "Content-Type: ")))
    if (!(rc = write_string (stream, ct)))
      rc = write_string (stream, is_text? ";\r\n":"\r\n");
  if (rc)
    return rc;

  /* OL inserts a charset parameter in many cases, so we do it right
     away for all text parts.  We can assume us-ascii if no special
     encoding is required.  */
  if (is_text)
    if ((rc = write_string (stream, (!use_qp && !use_b64)?
                            "\tcharset=\"us-ascii\"\r\n":
                            "\tcharset=\"utf-8\"\r\n")))
      return rc;
    
  /* Note that we need to output even 7bit because OL inserts that
     anyway.  */
  if (!(rc = write_string (stream, "Content-Transfer-Encoding: ")))
    rc = write_string (stream, (use_b64? "base64\r\n":
                                use_qp? "quoted-printable\r\n":"7bit\r\n"));
  if (rc)
    return rc;
  
  /* Write delimiter.  */
  if ((rc = write_string (stream, "\r\n")))
    return rc;
  
  /* Write the content.  */
  if (use_b64)
    rc = write_b64 (stream, data, datalen);
  else if (use_qp)
    rc = write_qp (stream, data, datalen);
  else
    rc = write_plain (stream, data, datalen);

  return rc;
}


/* Return the number of attachments in TABLE to be put into the MIME
   message.  */
static int
count_usable_attachments (mapi_attach_item_t *table)
{
  int idx, count = 0;
  
  if (table)
    for (idx=0; !table[idx].end_of_table; idx++)
      if (table[idx].attach_type == ATTACHTYPE_UNKNOWN
          && table[idx].method == ATTACH_BY_VALUE)
        count++;
  return count;
}

/* Write old all attachments from TABLE separated by BOUNDARY to
   STREAM.  This function needs to be syncronized with
   count_usable_attachments.  */
static int
write_attachments (LPSTREAM stream, 
                   LPMESSAGE message, mapi_attach_item_t *table, 
                   const char *boundary)
{
  int idx, rc;
  char *buffer;
  size_t buflen;

  if (table)
    for (idx=0; !table[idx].end_of_table; idx++)
      if (table[idx].attach_type == ATTACHTYPE_UNKNOWN
          && table[idx].method == ATTACH_BY_VALUE)
        {
          buffer = mapi_get_attach (message, table+idx, &buflen);
          if (!buffer)
            log_debug ("Attachment at index %d not found\n", idx);
          else
            log_debug ("Attachment at index %d: length=%d\n", idx, (int)buflen);
          if (!buffer)
            return -1;
          rc = write_part (stream, buffer, buflen, boundary,
                           table[idx].filename, 0);
          xfree (buffer);
        }
  return 0;
}


/* Delete all attachments from TABLE except for the one we just created */
static int
delete_all_attachments (LPMESSAGE message, mapi_attach_item_t *table)
{
  HRESULT hr;
  int idx;

  if (table)
    for (idx=0; !table[idx].end_of_table; idx++)
      {
        if (table[idx].attach_type == ATTACHTYPE_MOSSTEMPL)
          continue;
        hr = IMessage_DeleteAttach (message, table[idx].mapipos, 0, NULL, 0);
        if (hr)
          {
            log_error ("%s:%s: DeleteAttach failed: hr=%#lx\n",
                       SRCNAME, __func__, hr); 
            return -1;
          }
      }
  return 0;
}



/* Sign the MESSAGE using PROTOCOL.  If PROTOCOL is PROTOCOL_UNKNOWN
   the engine decides what protocol to use.  On return MESSAGE is
   modified so that sending it will result in a properly MOSS (that is
   PGP or S/MIME) signed message.  On failure the function tries to
   keep the original message intact but there is no 100% guarantee for
   it. */
int 
mime_sign (LPMESSAGE message, protocol_t protocol)
{
  int rc;
  HRESULT hr;
  LPATTACH outattach;
  LPSTREAM outstream;
  char boundary[BOUNDARYSIZE+1];
  char inner_boundary[BOUNDARYSIZE+1];
  SPropValue prop;
  mapi_attach_item_t *att_table = NULL;
  char *body = NULL;
  int n_att_usable;

  protocol = check_protocol (protocol);
  if (protocol == PROTOCOL_UNKNOWN)
    return -1;

  outattach = create_mapi_attachment (message, &outstream);
  if (!outattach)
    return -1;

  /* Get the attachment info and the body.  */
  body = mapi_get_body (message, NULL);
  if (body && !*body)
    {
      xfree (body);
      body = NULL;
    }
  att_table = mapi_create_attach_table (message, 0);
  n_att_usable = count_usable_attachments (att_table);
  if (!n_att_usable && !body)
    {
      log_debug ("%s:%s: can't sign an empty message\n", SRCNAME, __func__);
      goto failure;
    }

  /* Write the top header.  */
  generate_boundary (boundary);
  rc = write_string (outstream, ("MIME-Version: 1.0\r\n"
                                 "Content-Type: multipart/signed;\r\n"
                                 "\tprotocol=\"application/"));
  if (!rc)
    rc = write_string (outstream, 
                       (protocol == PROTOCOL_OPENPGP
                        ? "pgp-signature" 
                        : "pkcs7-signature"));
  if (!rc)
    rc = write_string (outstream, ("\";\r\n\tboundary=\""));
  if (!rc)
    rc = write_string (outstream, boundary);
  if (!rc)
    rc = write_string (outstream, "\"\r\n");
  if (rc)
    goto failure;

  if ((body && n_att_usable) || n_att_usable > 1)
    {
      /* A body and at least one attachment or more than one attachment  */
      generate_boundary (inner_boundary);
      rc = write_boundary (outstream, boundary, 0);
      if (!rc)
        rc = write_string (outstream, ("Content-Type: multipart/mixed;\r\n"
                                       "\tboundary=\""));
      if (!rc)
        rc = write_string (outstream, inner_boundary);
      if (!rc)
        rc = write_string (outstream, "\"\r\n");
      if (rc)
        goto failure;
    }
  else
    {
      /* Only one part.  */
      *inner_boundary = 0;
    }

  if (body)
    rc = write_part (outstream, body, strlen (body), 
                     *inner_boundary? inner_boundary : boundary, NULL, 1);
  if (!rc && n_att_usable)
    rc = write_attachments (outstream, message, att_table,
                            *inner_boundary? inner_boundary : boundary);
  if (rc)
    goto failure;

  /* Finish the possible multipart/mixed. */
  if (*inner_boundary)
    {
      rc = write_boundary (outstream, inner_boundary, 1);
      if (rc)
        goto failure;
    }

  /* Write signature attachment.  We don't write it directly but a
     palceholder there.  This spaceholder starts with the prefix of a
     boundary so that it won't accidently occur in the actual content. */
  if ((rc = write_boundary (outstream, boundary, 0)))
    goto failure;

  if ((rc = write_string (outstream, 
                          (protocol == PROTOCOL_OPENPGP)?
                          "Content-Type: application/pgp-signature\r\n":
                          "Content-Type: application/pkcs7-signature\r\n")))
    goto failure;

  /* If we would add "Content-Transfer-Encoding: 7bit\r\n" to this
     atatchment, Outlooks does not processed with sending and even
     does not return any error.  A wild guess is that while OL adds
     this header itself, it detects that it already exists and somehow
     gets into a problem.  It is not a problem with the other parts,
     though.  Hmmm, triggered by the top levels CT protocol parameter?
     Any way, it is not required that we add it as we won't hash
     it.  */


  if ((rc = write_string (outstream, "\r\n")))
    goto failure;

  if ((rc = write_string (outstream, "--=-=@SIGNATURE@\r\n\r\n")))
    goto failure;

  /* Write the final boundary and finish the attachment.  */
  if ((rc = write_boundary (outstream, boundary, 1)))
    goto failure;

  hr = IStream_Commit (outstream, 0);
  if (hr)
    {
      log_error ("%s:%s: Commiting output stream failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      goto failure;
    }
  IStream_Release (outstream);
  outstream = NULL;
  hr = IAttach_SaveChanges (outattach, KEEP_OPEN_READWRITE);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges of the attachment failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }
  IAttach_Release (outattach);
  outattach = NULL;

  /* Set the message class.  */
  prop.ulPropTag = PR_MESSAGE_CLASS_A;
  prop.Value.lpszA = "IPM.Note.SMIME.MultipartSigned"; 
  hr = IMessage_SetProps (message, 1, &prop, NULL);
  if (hr)
    {
      log_error ("%s:%s: error setting the message class: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto failure;
    }

  /* Now delete all parts of the MAPI message except for the one
     attachment we just created.  */
  if ((rc = delete_all_attachments (message, att_table)))
    goto failure;
  {
    /* Delete the body parts.  We don't return any error because there
       might be no body part at all.  To avoid aliasing problems when
       using static initialized array (SizedSPropTagArray macro) we
       call it two times in a row.  */
    SPropTagArray proparray;

    proparray.cValues = 1;
    proparray.aulPropTag[0] = PR_BODY;
    IMessage_DeleteProps (message, &proparray, NULL);
    proparray.cValues = 1;
    proparray.aulPropTag[0] = PR_BODY_HTML;
    IMessage_DeleteProps (message, &proparray, NULL);
  }
  
  /* Save the Changes.  */
  hr = IMessage_SaveChanges (message, KEEP_OPEN_READWRITE|FORCE_SAVE);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges to the message failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      goto failure;
    }

  xfree (body);
  mapi_release_attach_table (att_table);
  return 0;

 failure:
  if (outstream)
    {
      IStream_Revert (outstream);
      IStream_Release (outstream);
    }
  if (outattach)
    {
      /* Fixme: Should we try to delete it or is tehre a Revert method? */
      IAttach_Release (outattach);
    }
  xfree (body);
  mapi_release_attach_table (att_table);
  return -1;
}


int 
mime_encrypt (LPMESSAGE message, protocol_t protocol)
{
  return -1;
}


int 
mime_sign_encrypt (LPMESSAGE message, protocol_t protocol)
{
  return -1;
}
