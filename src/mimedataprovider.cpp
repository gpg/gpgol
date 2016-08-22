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
#include "common.h"

#include "mimedataprovider.h"

/* The maximum length of a line we are able to process.  RFC822 allows
   only for 1000 bytes; thus 2000 seems to be a reasonable value. */
#define LINEBUFSIZE 2000

/* How much data is read from the underlying stream in a collect
   call. */
#define BUFSIZE 8192

#include <gpgme++/error.h>

static int
message_cb (void *opaque, rfc822parse_event_t event,
            rfc822parse_t msg)
{
  (void) opaque;
  (void) event;
  (void) msg;
  return 0;
}

MimeDataProvider::MimeDataProvider(LPSTREAM stream):
  m_collect(true),
  m_parser(rfc822parse_open (message_cb, this)),
  m_current_encoding(None)
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
  b64_init (&m_base64_context);
  log_mime_parser ("%s:%s Collecting data.", SRCNAME, __func__);
  collect_data (stream);
  log_mime_parser ("%s:%s Data collected.", SRCNAME, __func__);
  gpgol_release (stream);
}

MimeDataProvider::~MimeDataProvider()
{
  log_debug ("%s:%s", SRCNAME, __func__);
}

bool
MimeDataProvider::isSupported(GpgME::DataProvider::Operation op) const
{
  return op == GpgME::DataProvider::Read ||
         op == GpgME::DataProvider::Seek ||
         op == GpgME::DataProvider::Release;
}

ssize_t
MimeDataProvider::read(void *buffer, size_t size)
{
  log_mime_parser ("%s:%s: Reading: " SIZE_T_FORMAT "Bytes",
                 SRCNAME, __func__, size);
  ssize_t bRead = m_data.read (buffer, size);
  if (opt.enable_debug & DBG_MIME_PARSER)
    {
      std::string buf ((char *)buffer, bRead);
      log_mime_parser ("%s:%s: Data: \"%s\"",
                     SRCNAME, __func__, buf.c_str());
    }
  return bRead;
}

void
MimeDataProvider::decode_and_collect(char *line, size_t pos)
{
  /* We are inside the data.  That should be the actual
     ciphertext in the given encoding. Add it to our internal
     cache. */
  int slbrk = 0;
  size_t len;

  if (m_current_encoding == Quoted)
    len = qp_decode (line, pos, &slbrk);
  else if (m_current_encoding == Base64)
    len = b64_decode (&m_base64_context, line, pos);
  else
    len = pos;
  m_data.write (line, len);
  if (m_current_encoding != Encoding::Base64 && !slbrk)
    {
      m_data.write ("\r\n", 2);
    }
  return;
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

          if (rfc822parse_insert (m_parser,
                                  (unsigned char*) linebuf,
                                  pos))
            {
              log_error ("%s:%s: rfc822 parser failed: %s\n",
                         SRCNAME, __func__, strerror (errno));
              return not_taken;
            }
          /* If we are currently in a collecting state actually
             collect that line */
          if (m_collect)
            {
              decode_and_collect (linebuf, pos);
            }
          /* Continue with next line. */
          pos = 0;
        }
    }
  return not_taken;
}

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

off_t
MimeDataProvider::seek(off_t offset, int whence)
{
  return m_data.seek (offset, whence);
}
