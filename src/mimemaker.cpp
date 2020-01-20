/* mimemaker.c - Construct MIME message out of a MAPI
 * Copyright (C) 2007, 2008 g10 Code GmbH
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
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
#include "mapihelp.h"
#include "mimemaker.h"
#include "oomhelp.h"
#include "mail.h"

#undef _
#define _(a) utf8_gettext (a)

static const unsigned char oid_mimetag[] =
    {0x2A, 0x86, 0x48, 0x86, 0xf7, 0x14, 0x03, 0x0a, 0x04};

/* The base-64 list used for base64 encoding. */
static unsigned char bintoasc[64+1] = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                       "abcdefghijklmnopqrstuvwxyz"
                                       "0123456789+/");


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
int
sink_std_write (sink_t sink, const void *data, size_t datalen)
{
  HRESULT hr;
  LPSTREAM stream = static_cast<LPSTREAM>(sink->cb_data);

  if (!stream)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }
  if (!data)
    return 0;  /* Flush - nothing to do here.  */

  hr = stream->Write(data, datalen, NULL);
  if (hr)
    {
      log_error ("%s:%s: Write failed: hr=%#lx", SRCNAME, __func__, hr);
      return -1;
    }
  return 0;
}

int
sink_string_write (sink_t sink, const void *data, size_t datalen)
{
  Mail *mail = static_cast<Mail *>(sink->cb_data);
  mail->appendToInlineBody (std::string((char*)data, datalen));
  return 0;
}

/* Write method used with a sink_t that contains a file object.  */
int
sink_file_write (sink_t sink, const void *data, size_t datalen)
{
  HANDLE hFile = sink->cb_data;
  DWORD written = 0;

  if (!hFile || hFile == INVALID_HANDLE_VALUE)
    {
      log_error ("%s:%s: sink not setup for writing", SRCNAME, __func__);
      return -1;
    }
  if (!data)
    return 0;  /* Flush - nothing to do here.  */

  if (!WriteFile (hFile, data, datalen, &written, NULL))
    {
      log_error ("%s:%s: Write failed: ", SRCNAME, __func__);
      return -1;
    }
  return 0;
}


/* Create a new MAPI attchment for MESSAGE which will be used to
   prepare the MIME message.  On sucess the stream to write the data
   to is stored at STREAM and the attachment object itself is
   returned.  The caller needs to call SaveChanges.  Returns NULL on
   failure in which case STREAM will be set to NULL.  */
LPATTACH
create_mapi_attachment (LPMESSAGE message, sink_t sink,
                        const char *overrideMimeTag)
{
  HRESULT hr;
  ULONG pos;
  SPropValue prop;
  LPATTACH att = NULL;
  LPUNKNOWN punk;

  sink->cb_data = NULL;
  sink->writefnc = NULL;
  hr = message->CreateAttach(NULL, 0, &pos, &att);
  memdbg_addRef (att);
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
  prop.Value.lpszA = xstrdup (MIMEATTACHFILENAME);
  hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
  xfree (prop.Value.lpszA);
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
      prop.Value.lpszA = overrideMimeTag ? xstrdup (overrideMimeTag) :
                         xstrdup ("multipart/signed");
      if (overrideMimeTag)
        {
          log_debug ("%s:%s: using override mimetag: %s\n",
                     SRCNAME, __func__, overrideMimeTag);
        }
      hr = HrSetOneProp ((LPMAPIPROP)att, &prop);
      xfree (prop.Value.lpszA);
    }
  if (hr)
    {
      log_error ("%s:%s: can't set attach mime tag: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      goto failure;
    }

  punk = NULL;
  hr = gpgol_openProperty (att, PR_ATTACH_DATA_BIN, &IID_IStream, 0,
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
  gpgol_release (att);
  return NULL;
}


/* Write data to a sink_t.  */
int
write_buffer (sink_t sink, const void *data, size_t datalen)
{
  if (!sink || !sink->writefnc)
    {
      log_error ("%s:%s: sink not properly setup", SRCNAME, __func__);
      return -1;
    }
  return sink->writefnc (sink, data, datalen);
}

/* Same as above but used for passing as callback function.  This
   fucntion does not return an error code but the number of bytes
   written.  */
int
write_buffer_for_cb (void *opaque, const void *data, size_t datalen)
{
  sink_t sink = (sink_t) opaque;
  sink->enc_counter += datalen;
  return write_buffer (sink, data, datalen) ? -1 : datalen;
}


/* Write the string TEXT to the IStream STREAM.  Returns 0 on sucsess,
   prints an error message and returns -1 on error.  */
int
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
int
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
int
write_b64 (sink_t sink, const void *data, size_t datalen)
{
  int rc;
  const unsigned char *p;
  unsigned char inbuf[4];
  int idx, quads;
  char outbuf[2048];
  size_t outlen;

  log_debug ("  writing base64 of length %d\n", (int)datalen);
  idx = quads = 0;
  outlen = 0;
  for (p = (const unsigned char*)data; datalen; p++, datalen--)
    {
      inbuf[idx++] = *p;
      if (idx > 2)
        {
          /* We need space for a quad and a possible CR,LF.  */
          if (outlen+4+2 >= sizeof outbuf)
            {
              if ((rc = write_buffer (sink, outbuf, outlen)))
                return rc;
              outlen = 0;
            }
          outbuf[outlen++] = bintoasc[(*inbuf>>2)&077];
          outbuf[outlen++] = bintoasc[(((*inbuf<<4)&060)
                                       |((inbuf[1] >> 4)&017))&077];
          outbuf[outlen++] = bintoasc[(((inbuf[1]<<2)&074)
                                       |((inbuf[2]>>6)&03))&077];
          outbuf[outlen++] = bintoasc[inbuf[2]&077];
          idx = 0;
          if (++quads >= (64/4))
            {
              quads = 0;
              outbuf[outlen++] = '\r';
              outbuf[outlen++] = '\n';
            }
        }
    }

  /* We need space for a quad and a final CR,LF.  */
  if (outlen+4+2 >= sizeof outbuf)
    {
      if ((rc = write_buffer (sink, outbuf, outlen)))
        return rc;
      outlen = 0;
    }
  if (idx)
    {
      outbuf[outlen++] = bintoasc[(*inbuf>>2)&077];
      if (idx == 1)
        {
          outbuf[outlen++] = bintoasc[((*inbuf<<4)&060)&077];
          outbuf[outlen++] = '=';
          outbuf[outlen++] = '=';
        }
      else
        {
          outbuf[outlen++] = bintoasc[(((*inbuf<<4)&060)
                                    |((inbuf[1]>>4)&017))&077];
          outbuf[outlen++] = bintoasc[((inbuf[1]<<2)&074)&077];
          outbuf[outlen++] = '=';
        }
      ++quads;
    }

  if (quads)
    {
      outbuf[outlen++] = '\r';
      outbuf[outlen++] = '\n';
    }

  if (outlen)
    {
      if ((rc = write_buffer (sink, outbuf, outlen)))
        return rc;
    }

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
  for (p = (const unsigned char*) data; datalen; p++, datalen--)
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
  for (p = (const unsigned char*) data; datalen; p++, datalen--)
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


/* Infer the content type from the FILENAME.  The return value is
   a static string there won't be an error return.  In case Base 64
   encoding is required for the type true will be stored at FORCE_B64;
   however, this is only a shortcut and if that is not set, the caller
   should infer the encoding by other means. */
static const char *
infer_content_type (const char * /*data*/, size_t /*datalen*/,
                    const char *filename, int is_mapibody, int *force_b64)
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
      { 1, "docx",  "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
      { 1, "dot",   "application/msword" },
      { 1, "dotx",  "application/vnd.openxmlformats-officedocument.wordprocessingml.template" },
      { 1, "docm",  "application/application/vnd.ms-word.document.macroEnabled.12" },
      { 1, "dotm",  "application/vnd.ms-word.template.macroEnabled.12" },
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
      { 1, "pot",   "application/vnd.ms-powerpoint" },
      { 1, "ppa",   "application/vnd.ms-powerpoint" },
      { 1, "pptx",  "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
      { 1, "potx",  "application/vnd.openxmlformats-officedocument.presentationml.template" },
      { 1, "ppsx",  "application/vnd.openxmlformats-officedocument.presentationml.slideshow" },
      { 1, "ppam",  "application/vnd.ms-powerpoint.addin.macroEnabled.12" },
      { 1, "pptm",  "application/vnd.ms-powerpoint.presentation.macroEnabled.12" },
      { 1, "potm",  "application/vnd.ms-powerpoint.template.macroEnabled.12" },
      { 1, "ppsm",  "application/vnd.ms-powerpoint.slideshow.macroEnabled.12" },
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
      { 1, "xlsx",  "application/vnd.ms-excel" },
      { 1, "xlt",   "application/vnd.ms-excel" },
      { 1, "xla",   "application/vnd.ms-excel" },
      { 1, "xltx",  "application/vnd.openxmlformats-officedocument.spreadsheetml.template" },
      { 1, "xlsm",  "application/vnd.ms-excel.sheet.macroEnabled.12" },
      { 1, "xltm",  "application/vnd.ms-excel.template.macroEnabled.12" },
      { 1, "xlam",  "application/vnd.ms-excel.addin.macroEnabled.12" },
      { 1, "xlsb",  "application/application/vnd.ms-excel.sheet.binary.macroEnabled.12" },
      { 0, "xml",   "application/xml" },
      { 0, "xsl",   "application/xml" },
      { 0, "xul",   "application/vnd.mozilla.xul+xml" },
      { 1, "zip",   "application/zip" },
      { 0, NULL, NULL }
    };
  int i;
  std::string suffix;

  *force_b64 = 0;
  if (filename)
    {
      const char *dot = strrchr (filename, '.');

      if (dot)
        {
          suffix = dot;
        }
    }

  /* Check for at least one char after the dot. */
  if (suffix.size() > 1)
    {
      /* Erase the dot */
      suffix.erase(0, 1);
      std::transform(suffix.begin(), suffix.end(), suffix.begin(), ::tolower);
      for (i=0; suffix_table[i].suffix; i++)
        {
          if (!strcmp (suffix_table[i].suffix, suffix.c_str()))
            {
              if (suffix_table[i].b64)
                *force_b64 = 1;
              return suffix_table[i].ct;
            }
        }
    }

  /* Not found via filename, look at the content.  */

  if (is_mapibody == 1)
    {
      return "text/plain";
    }
  else if (is_mapibody == 2)
    {
      return "text/html";
    }
  return "application/octet-stream";
}

/* Figure out the best encoding to be used for the part.  Return values are
     0: Plain ASCII.
     1: Quoted Printable
     2: Base64  */
static int
infer_content_encoding (const void *data, size_t datalen)
{
  const unsigned char *p;
  int need_qp;
  size_t len, maxlen, highbin, lowbin, ntotal;

  ntotal = datalen;
  len = maxlen = lowbin = highbin = 0;
  need_qp = 0;
  for (p = (const unsigned char*) data; datalen; p++, datalen--)
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


/* Convert an utf8 input string to RFC2047 base64 encoding which
   is the subset of RFC2047 outlook likes.
   Return value needs to be freed.
   */
static char *
utf8_to_rfc2047b (const char *input)
{
  char *ret,
       *encoded;
  int inferred_encoding = 0;
  if (!input)
    {
      return NULL;
    }
  inferred_encoding = infer_content_encoding (input, strlen (input));
  if (!inferred_encoding)
    {
      return xstrdup (input);
    }
  log_debug ("%s:%s: Encoding attachment filename. With: %s ",
             SRCNAME, __func__, inferred_encoding == 2 ? "Base64" : "QP");

  if (inferred_encoding == 2)
    {
      encoded = b64_encode (input, strlen (input));
      if (gpgrt_asprintf (&ret, "=?utf-8?B?%s?=", encoded) == -1)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          xfree (encoded);
          return NULL;
        }
    }
  else
    {
      /* There is a Bug here. If you encode 4 Byte UTF-8 outlook can't
         handle it itself. And sends out a message with ?? inserted in
         that place. This triggers an invalid signature. */
      encoded = qp_encode (input, strlen (input), NULL);
      if (gpgrt_asprintf (&ret, "=?utf-8?Q?%s?=", encoded) == -1)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          xfree (encoded);
          return NULL;
        }
    }
  xfree (encoded);
  return ret;
}

/* Write a MIME part to SINK.  First the BOUNDARY is written (unless
   it is NULL) then the DATA is analyzed and appropriate headers are
   written.  If FILENAME is given it will be added to the part's
   header.  IS_MAPIBODY should be passed as true if the data has been
   retrieved from the body property.  */
static int
write_part (sink_t sink, const char *data, size_t datalen,
            const char *boundary, const char *filename, int is_mapibody,
            const char *content_id = NULL)
{
  int rc;
  const char *ct;
  int use_b64, use_qp, is_text;
  char *encoded_filename;

  if (filename)
    {
      /* If there is a filename strip the directory part.  Take care
         that there might be slashes or backslashes.  */
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
             filename ? anonstr (filename) : "[none]");

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

  encoded_filename = utf8_to_rfc2047b (filename);
  if (encoded_filename)
    if ((rc=write_multistring (sink,
                               "\tname=\"", encoded_filename, "\"\r\n",
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

  if (content_id)
    {
      if ((rc=write_multistring (sink,
                                 "Content-ID: <", content_id, ">\r\n",
                                 NULL)))
        return rc;
    }
  else if (encoded_filename)
    if ((rc=write_multistring (sink,
                               "Content-Disposition: attachment;\r\n"
                               "\tfilename=\"", encoded_filename, "\"\r\n",
                               NULL)))
      return rc;

  xfree(encoded_filename);

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
int
count_usable_attachments (mapi_attach_item_t *table)
{
  int idx, count = 0;

  if (table)
    for (idx=0; !table[idx].end_of_table; idx++)
      if (table[idx].attach_type == ATTACHTYPE_UNKNOWN
          && (table[idx].method == ATTACH_BY_VALUE
              || table[idx].method == ATTACH_OLE
              || table[idx].method == ATTACH_EMBEDDED_MSG))
        {
          /* OLE and embedded are usable becase we plan to
             add support later. First version only handled
             them with a warning in write attachments. */
          count++;
        }
  return count;
}

/* Write out all attachments from TABLE separated by BOUNDARY to SINK.
   This function needs to be syncronized with count_usable_attachments.
   If only_related is 1 only include attachments for multipart/related they
   are excluded otherwise.
   If only_related is 2 all attachments are included regardless of
   content-id. */
static int
write_attachments (sink_t sink,
                   LPMESSAGE message, mapi_attach_item_t *table,
                   const char *boundary, int only_related)
{
  int idx, rc;
  char *buffer;
  size_t buflen;
  bool warning_shown = false;

  if (table)
    for (idx=0; !table[idx].end_of_table; idx++)
      {
        if (table[idx].attach_type == ATTACHTYPE_UNKNOWN
            && table[idx].method == ATTACH_BY_VALUE)
          {
            if (only_related == 1 && !table[idx].content_id)
              {
                continue;
              }
            else if (!only_related && table[idx].content_id)
              {
                continue;
              }
            buffer = mapi_get_attach (message, table+idx, &buflen);
            if (!buffer)
              log_debug ("Attachment at index %d not found\n", idx);
            else
              log_debug ("Attachment at index %d: length=%d\n", idx, (int)buflen);
            if (!buffer)
              return -1;
            rc = write_part (sink, buffer, buflen, boundary,
                             table[idx].filename, 0, table[idx].content_id);
            if (rc)
              {
                log_error ("Write part returned err: %i", rc);
              }
            xfree (buffer);
          }
        else if (!only_related && !warning_shown
                && table[idx].attach_type == ATTACHTYPE_UNKNOWN
                && (table[idx].method == ATTACH_OLE
                    || table[idx].method == ATTACH_EMBEDDED_MSG))
          {
            char *fmt;
            log_debug ("%s:%s: detected OLE attachment. Showing warning.",
                       SRCNAME, __func__);
            gpgrt_asprintf (&fmt, _("The attachment '%s' is an Outlook item "
                                    "which is currently unsupported in crypto mails."),
                            table[idx].filename ?
                            table[idx].filename : _("Unknown"));
            std::string msg = fmt;
            msg += "\n\n";
            xfree (fmt);

            gpgrt_asprintf (&fmt, _("Please encrypt '%s' with Kleopatra "
                                    "and attach it as a file."),
                            table[idx].filename ?
                            table[idx].filename : _("Unknown"));
            msg += fmt;
            xfree (fmt);

            msg += "\n\n";
            msg += _("Send anyway?");
            warning_shown = true;

            if (gpgol_message_box (get_active_hwnd (),
                                   msg.c_str (),
                                   _("Sorry, that's not possible, yet"),
                                   MB_APPLMODAL | MB_YESNO) == IDNO)
              {
                return -1;
              }
          }
        else
          {
            log_debug ("%s:%s: Skipping unknown attachment at idx: %d type: %d"
                       " with method: %d",
                       SRCNAME, __func__, idx, table[idx].attach_type,
                       table[idx].method);
          }
      }
  return 0;
}

/* Returns 1 if all attachments are related. 2 if there is a
   related and a mixed attachment. 0 if there are no other parts*/
static int
is_related (Mail *mail, mapi_attach_item_t *table)
{
  if (!mail || !mail->isHTMLAlternative () || !table)
    {
      return 0;
    }

  int related = 0;
  int mixed = 0;
  for (int idx = 0; !table[idx].end_of_table; idx++)
    {
      if (table[idx].content_id)
        {
          related = 1;
        }
      else
        {
          mixed = 1;
        }
    }
  return mixed + related;
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
        hr = message->DeleteAttach (table[idx].mapipos, 0, NULL, 0);
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
int
close_mapi_attachment (LPATTACH *attach, sink_t sink)
{
  HRESULT hr;
  LPSTREAM stream = sink ? (LPSTREAM) sink->cb_data : NULL;

  if (!stream)
    {
      log_error ("%s:%s: sink not setup", SRCNAME, __func__);
      return -1;
    }
  hr = stream->Commit (0);
  if (hr)
    {
      log_error ("%s:%s: Commiting output stream failed: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
  gpgol_release (stream);
  sink->cb_data = NULL;
  hr = (*attach)->SaveChanges (0);
  if (hr)
    {
      log_error ("%s:%s: SaveChanges of the attachment failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return -1;
    }
  gpgol_release ((*attach));
  *attach = NULL;
  return 0;
}


/* Cancel changes to the attachment ATTACH and release the object.
   SINK needs to be passed as well and will also be closed.  Note that
   the address of ATTACH is expected so that the fucntion can set it
   to NULL. */
void
cancel_mapi_attachment (LPATTACH *attach, sink_t sink)
{
  LPSTREAM stream = sink ? (LPSTREAM) sink->cb_data : NULL;

  if (stream)
    {
      stream->Revert();
      gpgol_release (stream);
      sink->cb_data = NULL;
    }
  if (*attach)
    {
      /* Fixme: Should we try to delete it or is there a Revert method? */
      gpgol_release ((*attach));
      *attach = NULL;
    }
}



/* Do the final processing for a message. */
int
finalize_message (LPMESSAGE message, mapi_attach_item_t *att_table,
                  protocol_t protocol, int encrypt, bool is_inline,
                  bool is_draft, int exchange_major_version)
{
  HRESULT hr = 0;
  SPropValue prop;
  SPropTagArray proparray;

  /* Set the message class.  */
  prop.ulPropTag = PR_MESSAGE_CLASS_A;
  if (protocol == PROTOCOL_SMIME)
    {
      /* When sending over exchange to the same server the recipient
         might see the message class we set here. So for S/MIME
         we keep the original. This makes the sent folder icon
         not immediately showing the GpgOL icon but gives other
         clients that do not have GpgOL installed a better chance
         to handle the mail. */
      if (encrypt && exchange_major_version >= 15)
        {
          /* This only appears to work with later exchange versions */
          prop.Value.lpszA = xstrdup ("IPM.Note.SMIME");
        }
      else
        {
          prop.Value.lpszA = xstrdup ("IPM.Note.SMIME.MultipartSigned");
        }
    }
  else if (encrypt)
    {
      prop.Value.lpszA = xstrdup ("IPM.Note.InfoPathForm.GpgOL.SMIME.MultipartSigned");
    }
  else
    {
      prop.Value.lpszA = xstrdup ("IPM.Note.InfoPathForm.GpgOLS.SMIME.MultipartSigned");
    }

  if (!is_inline)
    {
      /* For inline we stick with IPM.Note because Exchange Online would
         error out if we tried our S/MIME conversion trick with a text
         plain message */
      hr = message->SetProps(1, &prop, NULL);
    }
  xfree(prop.Value.lpszA);
  if (hr)
    {
      log_error ("%s:%s: error setting the message class: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return -1;
    }

  /* We also need to set the message class into our custom
     property. This override is at least required for encrypted
     messages.  */
  if (is_inline && mapi_set_gpgol_msg_class (message,
                                          (encrypt?
                                           (protocol == PROTOCOL_SMIME?
                                            "IPM.Note.GpgOL.OpaqueEncrypted" :
                                            "IPM.Note.GpgOL.PGPMessage") :
                                            "IPM.Note.GpgOL.ClearSigned")))
    {
      log_error ("%s:%s: error setting gpgol msgclass",
                 SRCNAME, __func__);
      return -1;
    }
  if (!is_inline && mapi_set_gpgol_msg_class (message,
                                (encrypt?
                                 (protocol == PROTOCOL_SMIME?
                                  "IPM.Note.GpgOL.OpaqueEncrypted" :
                                  "IPM.Note.GpgOL.MultipartEncrypted") :
                                 "IPM.Note.GpgOL.MultipartSigned")))
    {
      log_error ("%s:%s: error setting gpgol msgclass",
                 SRCNAME, __func__);
      return -1;
    }

  proparray.cValues = 1;
  proparray.aulPropTag[0] = PR_BODY;
  hr = message->DeleteProps (&proparray, NULL);
  if (hr)
    {
      log_debug_w32 (hr, "%s:%s: deleting PR_BODY failed",
                     SRCNAME, __func__);
    }

  proparray.cValues = 1;
  proparray.aulPropTag[0] = PR_BODY_HTML;
  hr = message->DeleteProps (&proparray, NULL);
  if (hr)
    {
      log_debug_w32 (hr, "%s:%s: deleting PR_BODY_HTML failed",
                     SRCNAME, __func__);
    }

  /* Now delete all parts of the MAPI message except for the one
     attachment we just created.  */
  if (delete_all_attachments (message, att_table))
    {
      log_error ("%s:%s: error deleting attachments",
                 SRCNAME, __func__);
      return -1;
    }

  /* Remove the draft info so that we don't leak the information on
     whether the message has been signed etc. when we send it.
     If it is a draft we are encrypting we want to keep them.

     To avoid confusion: draft_info for us means the state of
     the secure toggle button.
     */
  if (!is_draft)
    {
      mapi_set_gpgol_draft_info (message, NULL);
    }

  if (mapi_save_changes (message, KEEP_OPEN_READWRITE|FORCE_SAVE))
    {
      log_error ("%s:%s: error saving changes.",
                 SRCNAME, __func__);
      return -1;
    }
  return 0;
}


/* Helper to create the signing header.  This includes enough space
   for later fixup of the micalg parameter.  The MIME version is only
   written if FIRST is set.  */
void
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

/* Add the body, either as multipart/alternative or just as the
  simple body part. Depending on the format set in outlook. To
  avoid memory duplication it takes the plain body as parameter.

  Boundary is the potential outer boundary of a multipart/mixed
  mail. If it is null we assume the multipart/alternative is
  the only part.

  return is zero on success.
*/
static int
add_body (Mail *mail, const char *boundary, sink_t sink,
          const char *plain_body)
{
  if (!plain_body)
    {
      return 0;
    }
  bool is_alternative = false;
  if (mail)
    {
      is_alternative = mail->isHTMLAlternative ();
    }

  int rc = 0;
  if (!is_alternative || !plain_body)
    {
      if (plain_body)
        {
          rc = write_part (sink, plain_body, strlen (plain_body),
                           boundary, NULL, 1);
        }
      /* Just the plain body or no body. We are done. */
      return rc;
    }

  /* Add a new multipart / mixed element. */
  if (boundary && write_boundary (sink, boundary, 0))
    {
      TRACEPOINT;
      return 1;
    }

  /* Now for the multipart/alternative part. We never do HTML only. */
  char alt_boundary [BOUNDARYSIZE+1];
  generate_boundary (alt_boundary);

  if ((rc=write_multistring (sink,
                            "Content-Type: multipart/alternative;\r\n",
                            "\tboundary=\"", alt_boundary, "\"\r\n",
                            "\r\n",  /* <-- extra line */
                            NULL)))
    {
      TRACEPOINT;
      return rc;
    }

  /* Now the plain body part */
  if ((rc = write_part (sink, plain_body, strlen (plain_body),
                       alt_boundary, NULL, 1)))
    {
      TRACEPOINT;
      return rc;
    }

  /* Now the html body. It is somehow not accessible through PR_HTML,
     OutlookSpy also shows MAPI Unsupported (but shows the data) strange.
     We just cache it. Memory is cheap :-) */
  char *html_body = mail->takeCachedHTMLBody ();
  if (!html_body)
    {
      log_error ("%s:%s: BUG: Body but no html body in alternative mail?",
                 SRCNAME, __func__);
      return -1;
    }

  rc = write_part (sink, html_body, strlen (html_body),
                   alt_boundary, NULL, 2);
  xfree (html_body);
  if (rc)
    {
      TRACEPOINT;
      return rc;
    }
  /* Finish our multipart */
  return write_boundary (sink, alt_boundary, 1);
}

/* Add the body and attachments. Does multipart handling. */
int
add_body_and_attachments (sink_t sink, LPMESSAGE message,
                          mapi_attach_item_t *att_table, Mail *mail,
                          const char *body, int n_att_usable)
{
  int related = is_related (mail, att_table);
  int rc = 0;
  char inner_boundary[BOUNDARYSIZE+1];
  char outer_boundary[BOUNDARYSIZE+1];
  *outer_boundary = 0;
  *inner_boundary = 0;

  if (((body && n_att_usable) || n_att_usable > 1) && related == 1)
    {
      /* A body and at least one attachment or more than one attachment  */
      generate_boundary (outer_boundary);
      if ((rc=write_multistring (sink,
                                 "Content-Type: multipart/related;\r\n",
                                 "\tboundary=\"", outer_boundary, "\"\r\n",
                                 "\r\n", /* <--- Outlook adds an extra line. */
                                 NULL)))
        return rc;
    }
  else if ((body && n_att_usable) || n_att_usable > 1)
    {
      generate_boundary (outer_boundary);
      if ((rc=write_multistring (sink,
                                 "Content-Type: multipart/mixed;\r\n",
                                 "\tboundary=\"", outer_boundary, "\"\r\n",
                                 "\r\n", /* <--- Outlook adds an extra line. */
                                 NULL)))
        return rc;
    }

  /* Only one part.  */
  if (*outer_boundary && related == 2)
    {
      /* We have attachments that are related to the body and unrelated
         attachments. So we need another part. */
      if ((rc=write_boundary (sink, outer_boundary, 0)))
        {
          return rc;
        }
      generate_boundary (inner_boundary);
      if ((rc=write_multistring (sink,
                                 "Content-Type: multipart/related;\r\n",
                                 "\tboundary=\"", inner_boundary, "\"\r\n",
                                 "\r\n", /* <--- Outlook adds an extra line. */
                                 NULL)))
        {
          return rc;
        }
    }


  if ((rc=add_body (mail, *inner_boundary ? inner_boundary :
                          *outer_boundary ? outer_boundary : NULL,
                    sink, body)))
    {
      log_error ("%s:%s: Adding the body failed.",
                 SRCNAME, __func__);
      return rc;
    }
  if (!rc && n_att_usable && related)
    {
      /* Write the related attachments. */
      rc = write_attachments (sink, message, att_table,
                              *inner_boundary? inner_boundary :
                              *outer_boundary? outer_boundary : NULL, 1);
      if (rc)
        {
          return rc;
        }
      /* Close the related part if neccessary.*/
      if (*inner_boundary && (rc=write_boundary (sink, inner_boundary, 1)))
        {
          return rc;
        }
    }

  /* Now write the other attachments.

     If we are multipart related the related attachments were already
     written above. If we are not related we pass 2 to the write_attachements
     function to force that even attachments with a content id are written
     out.

     This happens for example when forwarding a plain text mail with
     attachments.
     */
  if (!rc && n_att_usable)
    {
      rc = write_attachments (sink, message, att_table,
                              *outer_boundary? outer_boundary : NULL,
                              related ? 0 : 2);
    }
  if (rc)
    {
      return rc;
    }

  /* Finish the possible multipart/mixed. */
  if (*outer_boundary && (rc = write_boundary (sink, outer_boundary, 1)))
    return rc;

  return rc;
}


/* Helper from mime_encrypt.  BOUNDARY is a buffer of at least
   BOUNDARYSIZE+1 bytes which will be set on return from that
   function.  */
int
create_top_encryption_header (sink_t sink, protocol_t protocol, char *boundary,
                              bool is_inline, int exchange_major_version)
{
  int rc;

  if (is_inline)
    {
      *boundary = 0;
      rc = 0;
      /* This would be nice and worked for Google Sync but it failed
         for Microsoft Exchange Online *sigh* so we put the body
         instead into the oom body property and stick with IPM Note.
      rc = write_multistring (sink,
                              "MIME-Version: 1.0\r\n"
                              "Content-Type: text/plain;\r\n"
                              "\tcharset=\"iso-8859-1\"\r\n"
                              "Content-Transfer-Encoding: 7BIT\r\n"
                              "\r\n",
                              NULL);
     */
    }
  else if (protocol == PROTOCOL_SMIME)
    {
      *boundary = 0;
      if (exchange_major_version >= 15)
        {
          /*
             For S/MIME encrypted mails we do not use the S/MIME conversion
             code anymore. With Exchange 2016 this no longer works. Instead
             we set an override mime tag, the extended headers in OOM in
             Mail::update_crypt_oom and let outlook convert the attachment
             to base64.

             A bit more details can be found in T3853 / T3884
             */
          rc = 0;
        }
      else
        {
          rc = write_multistring (sink,
                                  "Content-Type: application/pkcs7-mime; "
                                  "smime-type=enveloped-data;\r\n"
                                  "\tname=\"smime.p7m\"\r\n"
                                  "Content-Disposition: attachment; filename=\"smime.p7m\"\r\n"
                                  "Content-Transfer-Encoding: base64\r\n"
                                  "MIME-Version: 1.0\r\n"
                                  "\r\n",
                                  NULL);
        }
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
                              "Content-Disposition: inline;\r\n"
                              "\tfilename=\"" OPENPGP_ENC_NAME "\"\r\n"
                              "Content-Transfer-Encoding: 7Bit\r\n"
                              "\r\n", NULL);
     }

  return rc;
}


int
restore_msg_from_moss (LPMESSAGE message, LPDISPATCH moss_att,
                       msgtype_t type, char *msgcls)
{
  struct sink_s sinkmem;
  sink_t sink = &sinkmem;
  char *orig = NULL;
  int err = -1;
  char boundary[BOUNDARYSIZE+1];

  (void)msgcls;

  LPATTACH new_attach = create_mapi_attachment (message,
                                                sink);
  log_debug ("Restore message from moss called.");
  if (!new_attach)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }
  // TODO MORE
  if (type == MSGTYPE_SMIME)
    {
      create_top_encryption_header (sink, PROTOCOL_SMIME, boundary);
    }
  else if (type == MSGTYPE_GPGOL_MULTIPART_ENCRYPTED)
    {
      create_top_encryption_header (sink, PROTOCOL_OPENPGP, boundary);
    }
  else
    {
      log_error ("%s:%s: Unsupported messagetype: %i",
                 SRCNAME, __func__, type);
      goto done;
    }

  orig = get_pa_string (moss_att, PR_ATTACH_DATA_BIN_DASL);

  if (!orig)
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  if (write_string (sink, orig))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  if (*boundary && write_boundary (sink, boundary, 1))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  if (close_mapi_attachment (&new_attach, sink))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      goto done;
    }

  err = 0;
done:
  xfree (orig);
  return err;
}
