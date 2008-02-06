/* mimemaker.c - Construct MIME message out of a MAPI
 *	Copyright (C) 2007, 2008 g10 Code GmbH
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
#include <stdarg.h>
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

/* The object we use instead of IStream.  It allows us to have a
   callback method for output and thus for processing stuff
   recursively.  */
struct sink_s;
typedef struct sink_s *sink_t;
struct sink_s
{
  void *cb_data;
  sink_t extrasink;
  int (*writefnc)(sink_t sink, const void *data, size_t datalen);
/*   struct { */
/*     int idx; */
/*     unsigned char inbuf[4]; */
/*     int quads; */
/*   } b64; */
};


/* Object used to collect data in a memory buffer.  */
struct databuf_s
{
  size_t len;      /* Used length.  */
  size_t size;     /* Allocated length of BUF.  */
  char *buf;       /* Malloced buffer.  */
};


/*** local prototypes  ***/
static int write_multistring (sink_t sink, const char *text1,
                              ...) GPGOL_GCC_A_SENTINEL(0);





/* Standard write method used with a sink_t object.  */
static int
sink_std_write (sink_t sink, const void *data, size_t datalen)
{
  HRESULT hr;
  LPSTREAM stream = sink->cb_data;

  if (!stream)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }
  if (!data)
    return 0;  /* Flush - nothing to do here.  */

  hr = IStream_Write (stream, data, datalen, NULL);
  if (hr)
    {
      log_error ("%s:%s: Write failed: hr=%#lx", SRCNAME, __func__, hr);
      return -1;
    }
  return 0;
}


/* Make sure that PROTOCOL is usable or return a suitable protocol.
   On error PROTOCOL_UNKNOWN is returned.  */
static protocol_t
check_protocol (protocol_t protocol)
{
  switch (protocol)
    {
    case PROTOCOL_UNKNOWN:
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
   to is stored at STREAM and the attachment object itself is
   returned.  The caller needs to call SaveChanges.  Returns NULL on
   failure in which case STREAM will be set to NULL.  */
static LPATTACH
create_mapi_attachment (LPMESSAGE message, sink_t sink)
{
  HRESULT hr;
  ULONG pos;
  SPropValue prop;
  LPATTACH att = NULL;
  LPUNKNOWN punk;

  sink->cb_data = NULL;
  sink->writefnc = NULL;
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
  prop.Value.lpszA = MIMEATTACHFILENAME;
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
  sink->cb_data = (LPSTREAM)punk;
  sink->writefnc = sink_std_write;
  return att;

 failure:
  IAttach_Release (att);
  return NULL;
}


/* Write data to a sink_t.  */
static int 
write_buffer (sink_t sink, const void *data, size_t datalen)
{
  if (!sink || !sink->writefnc)
    {
      log_error ("%s:%s: sink not properliy setup", SRCNAME, __func__);
      return -1;
    }
  return sink->writefnc (sink, data, datalen);
}

/* Same as above but used for passing as callback function.  This
   fucntion does not return an error code but the number of bytes
   written.  */
static int
write_buffer_for_cb (void *opaque, const void *data, size_t datalen)
{
  sink_t sink = opaque;
  return write_buffer (sink, data, datalen) ? -1 : datalen;
}


/* Write the string TEXT to the IStream STREAM.  Returns 0 on sucsess,
   prints an error message and returns -1 on error.  */
static int 
write_string (sink_t sink, const char *text)
{
  return write_buffer (sink, text, strlen (text));
}


/* Write the string TEXT1 and all folloing arguments of type (const
   char*) to the SINK.  The list of argumens needs to be terminated
   with a NULL.  Returns 0 on sucsess, prints an error message and
   returns -1 on error.  */
static int
write_multistring (sink_t sink, const char *text1, ...)
{
  va_list arg_ptr;
  int rc;
  const char *s;
  
  va_start (arg_ptr, text1);
  s = text1;
  do
    rc = write_string (sink, s);
  while (!rc && (s=va_arg (arg_ptr, const char *)));
  va_end (arg_ptr);
  return rc;
}



/* Helper to write a boundary to the output sink.  The leading LF
   will be written as well.  */
static int
write_boundary (sink_t sink, const char *boundary, int lastone)
{
  int rc = write_string (sink, "\r\n--");
  if (!rc)
    rc = write_string (sink, boundary);
  if (!rc)
    rc = write_string (sink, lastone? "--\r\n":"\r\n");
  return rc;
}


/* Write DATALEN bytes of DATA to SINK in base64 encoding.  This
   creates a complete Base64 chunk including the trailing fillers.  */
static int
write_b64 (sink_t sink, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  unsigned char inbuf[4];
  int idx, quads;
  char outbuf[4];

  log_debug ("  writing base64 of length %d\n", (int)datalen);
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
          if ((rc = write_buffer (sink, outbuf, 4)))
            return rc;
          idx = 0;
          if (++quads >= (64/4)) 
            {
              quads = 0;
              if ((rc = write_buffer (sink, "\r\n", 2)))
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
      if ((rc = write_buffer (sink, outbuf, 4)))
        return rc;
      ++quads;
    }

  if (quads) 
    if ((rc = write_buffer (sink, "\r\n", 2)))
      return rc;

  return 0;
}

/* Write DATALEN bytes of DATA to SINK in quoted-prinable encoding. */
static int
write_qp (sink_t sink, const void *data, size_t datalen)
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
                if ((rc = write_buffer (sink, outbuf, outidx)))       \
                  return rc;                                          \
                outidx = 0;                                           \
              }                                                       \
          } while (0)
              
  log_debug ("  writing qp of length %d\n", (int)datalen);
  outidx = 0;
  for (p = data; datalen; p++, datalen--)
    {
      if ((datalen > 1 && *p == '\r' && p[1] == '\n') || *p == '\n')
        {
          /* Line break.  */
          outbuf[outidx++] = '\r';
          outbuf[outidx++] = '\n';
          if ((rc = write_buffer (sink, outbuf, outidx)))
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
      else if (!outidx && datalen >= 5 && !memcmp (p, "From ", 5))
        {
          /* Protect the 'F' so that MTAs won't prefix the "From "
             with an '>' */
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
      if ((rc = write_buffer (sink, outbuf, outidx)))
        return rc;
    }

# undef do_softlf
# undef nextlf_p
  return 0;
}


/* Write DATALEN bytes of DATA to SINK in plain ascii encoding. */
static int
write_plain (sink_t sink, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  char outbuf[100];
  int outidx;

  log_debug ("  writing ascii of length %d\n", (int)datalen);
  outidx = 0;
  for (p = data; datalen; p++, datalen--)
    {
      if ((datalen > 1 && *p == '\r' && p[1] == '\n') || *p == '\n')
        {
          outbuf[outidx++] = '\r';
          outbuf[outidx++] = '\n';
          if ((rc = write_buffer (sink, outbuf, outidx)))
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
      if ((rc = write_buffer (sink, outbuf, outidx)))
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
      else if (len == 1 && datalen >= 5 && !memcmp (p, "From ", 5))
        {
          /* The usual From hack is required so that MTAs do not
             prefix it with an '>'.  */
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






/* Write a MIME part to SINK.  First the BOUNDARY is written (unless
   it is NULL) then the DATA is analyzed and appropriate headers are
   written.  If FILENAME is given it will be added to the part's
   header.  IS_MAPIBODY should be passed as true if the data has been
   retrieved from the body property.  */
static int
write_part (sink_t sink, const char *data, size_t datalen,
            const char *boundary, const char *filename, int is_mapibody)
{
  int rc;
  const char *ct;
  int use_b64, use_qp, is_text;

  if (filename)
    {
      /* If there is a filename strip the directory part.  Take care
         that there might be slashes of backslashes.  */
      const char *s1 = strrchr (filename, '/');
      const char *s2 = strrchr (filename, '\\');
      
      if (!s1)
        s1 = s2;
      else if (s1 && s2 && s2 > s1)
        s1 = s2;

      if (s1)
        filename = s1;
      if (*filename && filename[1] == ':')
        filename += 2;
      if (!*filename)
        filename = NULL;
    }

  log_debug ("Writing part of length %d%s filename=`%s'\n",
             (int)datalen, is_mapibody? " (body)":"", 
             filename?filename:"[none]");

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

  if (boundary)
    if ((rc = write_boundary (sink, boundary, 0)))
      return rc;
  if ((rc=write_multistring (sink,
                             "Content-Type: ", ct,
                             (is_text || filename? ";\r\n" :"\r\n"),
                             NULL)))
    return rc;

  /* OL inserts a charset parameter in many cases, so we do it right
     away for all text parts.  We can assume us-ascii if no special
     encoding is required.  */
  if (is_text)
    if ((rc=write_multistring (sink,
                               "\tcharset=\"",
                               (!use_qp && !use_b64? "us-ascii" : "utf-8"),
                               filename ? "\";\r\n" : "\"\r\n",
                               NULL)))
      return rc;

  if (filename)
    if ((rc=write_multistring (sink,
                               "\tname=\"", filename, "\"\r\n",
                               NULL)))
      return rc;

  /* Note that we need to output even 7bit because OL inserts that
     anyway.  */
  if ((rc = write_multistring (sink,
                               "Content-Transfer-Encoding: ",
                               (use_b64? "base64\r\n":
                                use_qp? "quoted-printable\r\n":"7bit\r\n"),
                               NULL)))
    return rc;
  
  if (filename)
    if ((rc=write_multistring (sink,
                               "Content-Disposition: attachment;\r\n"
                               "\tfilename=\"", filename, "\"\r\n",
                               NULL)))
      return rc;

  
  /* Write delimiter.  */
  if ((rc = write_string (sink, "\r\n")))
    return rc;
  
  /* Write the content.  */
  if (use_b64)
    rc = write_b64 (sink, data, datalen);
  else if (use_qp)
    rc = write_qp (sink, data, datalen);
  else
    rc = write_plain (sink, data, datalen);

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

/* Write out all attachments from TABLE separated by BOUNDARY to SINK.
   This function needs to be syncronized with count_usable_attachments.  */
static int
write_attachments (sink_t sink, 
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
          rc = write_part (sink, buffer, buflen, boundary,
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
        if (table[idx].attach_type == ATTACHTYPE_MOSSTEMPL
            && table[idx].filename
            && !strcmp (table[idx].filename, MIMEATTACHFILENAME))
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



/* Commit changes to the attachment ATTACH and release the object.
   SINK needs to be passed as well and will also be closed.  Note that
   the address of ATTACH is expected so that the fucntion can set it
   to NULL. */
static int
close_mapi_attachment (LPATTACH *attach, sink_t sink)
{
  HRESULT hr;
  LPSTREAM stream = sink? sink->cb_data : NULL;

  if (!stream)
    {
      log_error ("%s:%s: sink not setup", SRCNAME, __func__);
      return -1;
    }
  hr = IStream_Commit (stream, 0);
  if (hr)
    {
      log_error ("%s:%s: Commiting output stream failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
  IStream_Release (stream);
  sink->cb_data = NULL;
  hr = IAttach_SaveChanges (*attach, 0);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges of the attachment failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      return -1;
    }
  IAttach_Release (*attach);
  *attach = NULL;
  return 0;
}


/* Cancel changes to the attachment ATTACH and release the object.
   SINK needs to be passed as well and will also be closed.  Note that
   the address of ATTACH is expected so that the fucntion can set it
   to NULL. */
static void
cancel_mapi_attachment (LPATTACH *attach, sink_t sink)
{
  LPSTREAM stream = sink? sink->cb_data : NULL;

  if (stream)
    {
      IStream_Revert (stream);
      IStream_Release (stream);
      sink->cb_data = NULL;
    }
  if (*attach)
    {
      /* Fixme: Should we try to delete it or is there a Revert method? */
      IAttach_Release (*attach);
      *attach = NULL;
    }
}



/* Do the final processing for a message. */
static int
finalize_message (LPMESSAGE message, mapi_attach_item_t *att_table,
                  protocol_t protocol, int encrypt)
{
  HRESULT hr;
  SPropValue prop;

  /* Set the message class.  */
  prop.ulPropTag = PR_MESSAGE_CLASS_A;
  prop.Value.lpszA = "IPM.Note.SMIME.MultipartSigned"; 
  hr = IMessage_SetProps (message, 1, &prop, NULL);
  if (hr)
    {
      log_error ("%s:%s: error setting the message class: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return -1;
    }

  /* Set a special property so that we are later able to identify
     messages signed or encrypted by us.  */
  if (mapi_set_sig_status (message, "@"))
    return -1;

  /* We also need to set the message class into our custom
     property. This override is at leas required for encrypted
     messages.  */
  if (mapi_set_gpgol_msg_class (message,
                                (encrypt? 
                                 (protocol == PROTOCOL_SMIME? 
                                  "IPM.Note.GpgOL.OpaqueEncrypted" :
                                  "IPM.Note.GpgOL.MultipartEncrypted") :
                                 "IPM.Note.GpgOL.MultipartSigned")))
    return -1;

  /* Now delete all parts of the MAPI message except for the one
     attachment we just created.  */
  if (delete_all_attachments (message, att_table))
    return -1;
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
      return -1;
    }

  return 0;
}



/* Sink write method used by mime_sign.  We write the data to the
   filter and also to the EXTRASINK but we don't pass a flush request
   to EXTRASINK. */
static int
sink_hashing_write (sink_t hashsink, const void *data, size_t datalen)
{
  int rc;
  engine_filter_t filter = hashsink->cb_data;

  if (!filter || !hashsink->extrasink)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }
  
  rc = engine_filter (filter, data, datalen);
  if (!rc && data && datalen)
    write_buffer (hashsink->extrasink, data, datalen);
  return rc;
}


/* This function is called by the filter to collect the output which
   is a detached signature.  */
static int
collect_signature (void *opaque, const void *data, size_t datalen)
{
  struct databuf_s *db = opaque;

  if (db->len + datalen >= db->size)
    {
      db->size += datalen + 1024;
      db->buf = xrealloc (db->buf, db->size);
    }
  memcpy (db->buf + db->len, data, datalen);
  db->len += datalen;

  return datalen;
}


/* Helper to create the signing header.  This includes enough space
   for later fixup of the micalg parameter.  The MIME version is only
   written if FIRST is set.  */
static void
create_top_signing_header (char *buffer, size_t buflen, protocol_t protocol,
                           int first, const char *boundary, const char *micalg)
{
  snprintf (buffer, buflen,
            "%s"
            "Content-Type: multipart/signed;\r\n"
            "\tprotocol=\"application/%s\";\r\n"
            "\tmicalg=%-15.15s;\r\n"
            "\tboundary=\"%s\"\r\n"
            "\r\n",
            first? "MIME-Version: 1.0\r\n":"",
            (protocol==PROTOCOL_OPENPGP? "pgp-signature":"pkcs7-signature"),
            micalg, boundary);
}


/* Main body of mime_sign without the the code to delete the original
   attachments.  On success the function returns the current
   attachment table at R_ATT_TABLE or sets this to NULL on error.  If
   TMPSINK is set no attachment will be created but the output
   written to that sink.  */
static int 
do_mime_sign (LPMESSAGE message, HWND hwnd, protocol_t protocol, 
              mapi_attach_item_t **r_att_table, sink_t tmpsink)
{
  int result = -1;
  int rc;
  LPATTACH attach;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  struct sink_s hashsinkmem;
  sink_t hashsink = &hashsinkmem;
  char boundary[BOUNDARYSIZE+1];
  char inner_boundary[BOUNDARYSIZE+1];
  mapi_attach_item_t *att_table = NULL;
  char *body = NULL;
  int n_att_usable;
  char top_header[BOUNDARYSIZE+200];
  engine_filter_t filter;
  struct databuf_s sigbuffer;

  *r_att_table = NULL;

  memset (sink, 0, sizeof *sink);
  memset (hashsink, 0, sizeof *hashsink);
  memset (&sigbuffer, 0, sizeof sigbuffer);

  protocol = check_protocol (protocol);
  if (protocol == PROTOCOL_UNKNOWN)
    return -1;

  if (tmpsink)
    {
      attach = NULL;
      sink = tmpsink;
    }
  else
    {
      attach = create_mapi_attachment (message, sink);
      if (!attach)
        return -1;
    }

  /* Prepare the signing.  */
  if (engine_create_filter (&filter, collect_signature, &sigbuffer))
    goto failure;
  if (engine_sign_start (filter, hwnd, protocol))
    goto failure;

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
  create_top_signing_header (top_header, sizeof top_header,
                             protocol, 1, boundary, "xxx");
  if ((rc = write_string (sink, top_header)))
    goto failure;

  /* Create the inner boundary if we have a body and at least one
     attachment or more than one attachment.  */
  if ((body && n_att_usable) || n_att_usable > 1)
    generate_boundary (inner_boundary);
  else 
    *inner_boundary = 0;

  /* Write the boundary so that it is not included in the hashing.  */
  if ((rc = write_boundary (sink, boundary, 0)))
    goto failure;

  /* Create a new sink for hashing and write/hash our content.  */
  hashsink->cb_data = filter;
  hashsink->extrasink = sink;
  hashsink->writefnc = sink_hashing_write;

  /* Note that OL2003 will add an extra line after the multipart
     header, thus we do the same to avoid running all through an
     IConverterSession first. */
  if (*inner_boundary
      && (rc=write_multistring (hashsink, 
                                "Content-Type: multipart/mixed;\r\n",
                                "\tboundary=\"", inner_boundary, "\"\r\n",
                                "\r\n",  /* <-- extra line */
                                NULL)))
        goto failure;


  if (body)
    rc = write_part (hashsink, body, strlen (body),
                     *inner_boundary? inner_boundary : NULL, NULL, 1);
  if (!rc && n_att_usable)
    rc = write_attachments (hashsink, message, att_table,
                            *inner_boundary? inner_boundary : NULL);
  if (rc)
    goto failure;

  xfree (body);
  body = NULL;

  /* Finish the possible multipart/mixed. */
  if (*inner_boundary && (rc = write_boundary (hashsink, inner_boundary, 1)))
    goto failure;

  /* Here we are ready with the hashing.  Flush the filter and wait
     for the signing process to finish.  */
  if ((rc = write_buffer (hashsink, NULL, 0)))
    goto failure;
  if ((rc = engine_wait (filter)))
    goto failure;
  filter = NULL; /* Not valid anymore.  */
  hashsink->cb_data = NULL; /* Not needed anymore.  */
  

  /* Write signature attachment.  */
  if ((rc = write_boundary (sink, boundary, 0)))
    goto failure;

  if ((rc=write_string (sink, 
                        (protocol == PROTOCOL_OPENPGP
                         ? "Content-Type: application/pgp-signature\r\n"
                         : ("Content-Transfer-Encoding: base64\r\n"
                            "Content-Type: application/pkcs7-signature\r\n")
                         ))))
    goto failure;

  /* If we would add "Content-Transfer-Encoding: 7bit\r\n" to this
     attachment, Outlooks does not proceed with sending and even does
     not return any error.  A wild guess is that while OL adds this
     header itself, it detects that it already exists and somehow gets
     into a problem.  It is not a problem with the other parts,
     though.  Hmmm, triggered by the top levels CT protocol parameter?
     Anyway, it is not required that we add it as we won't hash it.
     Note, that this only holds for OpenPGP; for S/MIME we need to set
     set CTE.  We even write it before the CT because that is the same
     as Outlook would do it for a missing CTE. */

  if ((rc = write_string (sink, "\r\n")))
    goto failure;

  /* Write the signature.  We add an extra CR,LF which should not harm
     and a terminating 0. */
  collect_signature (&sigbuffer, "\r\n", 3); 
  if ((rc = write_string (sink, sigbuffer.buf)))
    goto failure;


  /* Write the final boundary and finish the attachment.  */
  if ((rc = write_boundary (sink, boundary, 1)))
    goto failure;

  /* Fixup the micalg parameter.  */
  {
    HRESULT hr;
    LARGE_INTEGER off;
    LPSTREAM stream = sink->cb_data;

    off.QuadPart = 0;
    hr = IStream_Seek (stream, off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        goto failure;
      }

    create_top_signing_header (top_header, sizeof top_header,
                               protocol, 1, boundary,
                               protocol == PROTOCOL_SMIME? "sha1":"pgp-sha1");

    hr = IStream_Write (stream, top_header, strlen (top_header), NULL);
    if (hr)
      {
        log_error ("%s:%s: writing fixed micalg failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        goto failure;
      }

    /* Better seek again to the end. */
    off.QuadPart = 0;
    hr = IStream_Seek (stream, off, STREAM_SEEK_END, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the end failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        goto failure;
      }
  }


  if (attach)
    {
      if (close_mapi_attachment (&attach, sink))
        goto failure;
    }

  result = 0;  /* Everything is fine, fall through the cleanup now.  */

 failure:
  engine_cancel (filter);
  if (attach)
    cancel_mapi_attachment (&attach, sink);
  xfree (body);
  if (result)
    mapi_release_attach_table (att_table);
  else
    *r_att_table = att_table;
  xfree (sigbuffer.buf);
  return result;
}


/* Sign the MESSAGE using PROTOCOL.  If PROTOCOL is PROTOCOL_UNKNOWN
   the engine decides what protocol to use.  On return MESSAGE is
   modified so that sending it will result in a properly MOSS (that is
   PGP or S/MIME) signed message.  On failure the function tries to
   keep the original message intact but there is no 100% guarantee for
   it. */
int 
mime_sign (LPMESSAGE message, HWND hwnd, protocol_t protocol)
{
  int result = -1;
  mapi_attach_item_t *att_table;

  if (!do_mime_sign (message, hwnd, protocol, &att_table, 0))
    {
      if (!finalize_message (message, att_table, protocol, 0))
        result = 0;
    }

  mapi_release_attach_table (att_table);
  return result;
}



/* Sink write method used by mime_encrypt.  */
static int
sink_encryption_write (sink_t encsink, const void *data, size_t datalen)
{
  engine_filter_t filter = encsink->cb_data;

  if (!filter)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }

  return engine_filter (filter, data, datalen);
}


#if 0 /* Not used.  */
/* Sink write method used by mime_encrypt for writing Base64.  */
static int
sink_encryption_write_b64 (sink_t encsink, const void *data, size_t datalen)
{
  engine_filter_t filter = encsink->cb_data;
  int rc;
  const unsigned char *p;
  unsigned char inbuf[4];
  int idx, quads;
  char outbuf[6];
  size_t outbuflen;

  if (!filter)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }

  idx = encsink->b64.idx;
  assert (idx < 4);
  memcpy (inbuf, encsink->b64.inbuf, idx);
  quads = encsink->b64.quads;

  if (!data)  /* Flush. */
    {
      outbuflen = 0;
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
              outbuf[1] = bintoasc[(((*inbuf<<4)&060)|
                                    ((inbuf[1]>>4)&017))&077];
              outbuf[2] = bintoasc[((inbuf[1]<<2)&074)&077];
              outbuf[3] = '=';
            }
          outbuflen = 4;
          quads++;
        }
      
      if (quads)
        {
          outbuf[outbuflen++] = '\r';
          outbuf[outbuflen++] = '\n';
        }

      if (outbuflen && (rc = engine_filter (filter, outbuf, outbuflen)))
        return rc;
      /* Send the flush command to the filter.  */
      if ((rc = engine_filter (filter, data, datalen)))
        return rc;
    }
  else
    {
      for (p = data; datalen; p++, datalen--)
        {
          inbuf[idx++] = *p;
          if (idx > 2)
            {
              idx = 0;
              outbuf[0] = bintoasc[(*inbuf>>2)&077];
              outbuf[1] = bintoasc[(((*inbuf<<4)&060)
                                    |((inbuf[1] >> 4)&017))&077];
              outbuf[2] = bintoasc[(((inbuf[1]<<2)&074)
                                    |((inbuf[2]>>6)&03))&077];
              outbuf[3] = bintoasc[inbuf[2]&077];
              outbuflen = 4;
              if (++quads >= (64/4)) 
                {
                  quads = 0;
                  outbuf[4] = '\r';
                  outbuf[5] = '\n';
                  outbuflen += 2;
                }
              if ((rc = engine_filter (filter, outbuf, outbuflen)))
                return rc;
            }
        }
    }

  encsink->b64.idx = idx;
  memcpy (encsink->b64.inbuf, inbuf, idx);
  encsink->b64.quads = quads;
  
  return 0;
}
#endif /*Not used.*/


/* Helper from mime_encrypt.  BOUNDARY is a buffer of at least
   BOUNDARYSIZE+1 bytes which will be set on return from that
   function.  */
static int
create_top_encryption_header (sink_t sink, protocol_t protocol, char *boundary)
{
  int rc;

  if (protocol == PROTOCOL_SMIME)
    {
      *boundary = 0;
      rc = write_multistring (sink,
                              "MIME-Version: 1.0\r\n"
                              "Content-Type: application/pkcs7-mime;\r\n"
                              "\tsmime-type=enveloped-data;\r\n"
                              "\tname=\"smime.p7m\"\r\n"
                              "Content-Transfer-Encoding: base64\r\n"
                              "\r\n",
                              NULL);
    }
  else
    {
      generate_boundary (boundary);
      rc = write_multistring (sink,
                              "MIME-Version: 1.0\r\n"
                              "Content-Type: multipart/encrypted;\r\n"
                              "\tprotocol=\"application/pgp-encrypted\";\r\n",
                              "\tboundary=\"", boundary, "\"\r\n",
                              NULL);
      if (rc)
        return rc;

      /* Write the PGP/MIME encrypted part.  */
      rc = write_boundary (sink, boundary, 0);
      if (rc)
        return rc;
      rc = write_multistring (sink,
                              "Content-Type: application/pgp-encrypted\r\n"
                              "\r\n"
                              "Version: 1\r\n", NULL);
      if (rc)
        return rc;
      
      /* And start the second part.  */
      rc = write_boundary (sink, boundary, 0);
      if (rc)
        return rc;
      rc = write_multistring (sink,
                              "Content-Type: application/octet-stream\r\n"
                              "\r\n", NULL);
     }

  return rc;
}


/* Encrypt the MESSAGE.  */
int 
mime_encrypt (LPMESSAGE message, HWND hwnd, 
              protocol_t protocol, char **recipients)
{
  int result = -1;
  int rc;
  LPATTACH attach;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  char boundary[BOUNDARYSIZE+1];
  char inner_boundary[BOUNDARYSIZE+1];
  mapi_attach_item_t *att_table = NULL;
  char *body = NULL;
  int n_att_usable;
  engine_filter_t filter;

  memset (sink, 0, sizeof *sink);
  memset (encsink, 0, sizeof *encsink);

  attach = create_mapi_attachment (message, sink);
  if (!attach)
    return -1;

  /* Prepare the encryption.  We do this early as it is quite common
     that some recipients are not be available and thus the encryption
     will fail early. */
  if (engine_create_filter (&filter, write_buffer_for_cb, sink))
    goto failure;
  if (engine_encrypt_start (filter, hwnd, protocol, recipients, &protocol))
    goto failure;

  protocol = check_protocol (protocol);
  if (protocol == PROTOCOL_UNKNOWN)
    goto failure;

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
      log_debug ("%s:%s: can't encrypt an empty message\n", SRCNAME, __func__);
      goto failure;
    }

  /* Write the top header.  */
  rc = create_top_encryption_header (sink, protocol, boundary);
  if (rc)
    goto failure;

  /* Create a new sink for encrypting the following stuff.  */
  encsink->cb_data = filter;
  encsink->writefnc = sink_encryption_write;
  
  if ((body && n_att_usable) || n_att_usable > 1)
    {
      /* A body and at least one attachment or more than one attachment  */
      generate_boundary (inner_boundary);
      if ((rc=write_multistring (encsink, 
                                 "Content-Type: multipart/mixed;\r\n",
                                 "\tboundary=\"", inner_boundary, "\"\r\n",
                                 NULL)))
        goto failure;
    }
  else /* Only one part.  */
    *inner_boundary = 0;

  if (body)
    rc = write_part (encsink, body, strlen (body), 
                     *inner_boundary? inner_boundary : NULL, NULL, 1);
  if (!rc && n_att_usable)
    rc = write_attachments (encsink, message, att_table,
                            *inner_boundary? inner_boundary : NULL);
  if (rc)
    goto failure;

  xfree (body);
  body = NULL;

  /* Finish the possible multipart/mixed. */
  if (*inner_boundary && (rc = write_boundary (encsink, inner_boundary, 1)))
    goto failure;

  /* Flush the encryption sink and wait for the encryption to get
     ready.  */
  if ((rc = write_buffer (encsink, NULL, 0)))
    goto failure;
  if ((rc = engine_wait (filter)))
    goto failure;
  filter = NULL; /* Not valid anymore.  */
  encsink->cb_data = NULL; /* Not needed anymore.  */
  
  /* Write the final boundary (for OpenPGP) and finish the attachment.  */
  if (*boundary && (rc = write_boundary (sink, boundary, 1)))
    goto failure;
  
  if (close_mapi_attachment (&attach, sink))
    goto failure;

  if (finalize_message (message, att_table, protocol, 1))
    goto failure;

  result = 0;  /* Everything is fine, fall through the cleanup now.  */

 failure:
  engine_cancel (filter);
  cancel_mapi_attachment (&attach, sink);
  xfree (body);
  mapi_release_attach_table (att_table);
  return result;
}




/* Sign and Encrypt the MESSAGE.  */
int 
mime_sign_encrypt (LPMESSAGE message, HWND hwnd, 
                   protocol_t protocol, char **recipients)
{
  int result = -1;
  int rc = 0;
  HRESULT hr;
  LPATTACH attach;
  LPSTREAM tmpstream = NULL;
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  struct sink_s encsinkmem;
  sink_t encsink = &encsinkmem;
  struct sink_s tmpsinkmem;
  sink_t tmpsink = &tmpsinkmem;
  char boundary[BOUNDARYSIZE+1];
  mapi_attach_item_t *att_table = NULL;
  engine_filter_t filter;

  memset (sink, 0, sizeof *sink);
  memset (encsink, 0, sizeof *encsink);
  memset (tmpsink, 0, sizeof *tmpsink);

  attach = create_mapi_attachment (message, sink);
  if (!attach)
    return -1;

  /* Create a temporary sink to construct the signed data.  */ 
  hr = OpenStreamOnFile (MAPIAllocateBuffer, MAPIFreeBuffer,
                         (SOF_UNIQUEFILENAME | STGM_DELETEONRELEASE
                          | STGM_CREATE | STGM_READWRITE),
                         NULL, "GPG", &tmpstream); 
  if (FAILED (hr)) 
    {
      log_error ("%s:%s: can't create temp file: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      rc = -1;
      goto failure;
    }
  tmpsink->cb_data = tmpstream;
  tmpsink->writefnc = sink_std_write;


  /* Prepare the encryption.  We do this early as it is quite common
     that some recipients are not be available and thus the encryption
     will fail early. */
  if (engine_create_filter (&filter, write_buffer_for_cb, sink))
    goto failure;
  if ((rc=engine_encrypt_start (filter, hwnd, 
                                protocol, recipients, &protocol)))
    goto failure;

  protocol = check_protocol (protocol);
  if (protocol == PROTOCOL_UNKNOWN)
    goto failure;

  /* Now sign the message.  This creates another attachment with the
     complete MIME object of the signed message.  We can't do the
     encryption in streaming mode while running the encryption because
     we need to fix up that ugly micalg parameter after having created
     the signature.  */
  if (do_mime_sign (message, hwnd, protocol, &att_table, tmpsink))
    goto failure;

  /* Write the top header.  */
  rc = create_top_encryption_header (sink, protocol, boundary);
  if (rc)
    goto failure;

  /* Create a new sink for encrypting the temporary attachment with
     the signed message.  */
  encsink->cb_data = filter;
  encsink->writefnc = sink_encryption_write;

  /* Copy the temporary stream to the encryption sink.  */
  {
    LARGE_INTEGER off;
    ULONG nread;
    char buffer[4096];

    off.QuadPart = 0;
    hr = IStream_Seek (tmpstream, off, STREAM_SEEK_SET, NULL);
    if (hr)
      {
        log_error ("%s:%s: seeking back to the begin failed: hr=%#lx",
                   SRCNAME, __func__, hr);
        rc = gpg_error (GPG_ERR_EIO);
        goto failure;
      }

    for (;;)
      {
        hr = IStream_Read (tmpstream, buffer, sizeof buffer, &nread);
        if (hr)
          {
            log_error ("%s:%s: IStream::Read failed: hr=%#lx", 
                       SRCNAME, __func__, hr);
            rc = gpg_error (GPG_ERR_EIO);
            goto failure;
          }
        if (!nread)
          break;  /* EOF */
        rc = write_buffer (encsink, buffer, nread);
        if (rc)
          {
            log_error ("%s:%s: writing tmpstream to encsink failed: %s", 
                       SRCNAME, __func__, gpg_strerror (rc));
            goto failure;
          }
      }
  }


  /* Flush the encryption sink and wait for the encryption to get
     ready.  */
  if ((rc = write_buffer (encsink, NULL, 0)))
    goto failure;
  if ((rc = engine_wait (filter)))
    goto failure;
  filter = NULL; /* Not valid anymore.  */
  encsink->cb_data = NULL; /* Not needed anymore.  */
  
  /* Write the final boundary (for OpenPGP) and finish the attachment.  */
  if (*boundary && (rc = write_boundary (sink, boundary, 1)))
    goto failure;
  
  if (close_mapi_attachment (&attach, sink))
    goto failure;

  if (finalize_message (message, att_table, protocol, 1))
    goto failure;

  result = 0;  /* Everything is fine, fall through the cleanup now.  */

 failure:
  if (result)
    log_debug ("%s:%s: failed rc=%d (%s) <%s>", SRCNAME, __func__, rc, 
               gpg_strerror (rc), gpg_strsource (rc));
  engine_cancel (filter);
  if (tmpstream)
    IStream_Release (tmpstream);
  mapi_release_attach_table (att_table);
  return result;
}
