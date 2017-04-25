/* mimedataprover.h - GpgME dataprovider for mime data
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
#ifndef MIMEDATAPROVIDER_H
#define MIMEDATAPROVIDER_H

#include "config.h"

#include <gpgme++/interfaces/dataprovider.h>
#include <gpgme++/data.h>
#include "rfc822parse.h"

#ifdef HAVE_W32_SYSTEM
#include "mapihelp.h"
#endif

#include <string>
struct mime_context;
typedef struct mime_context *mime_context_t;
class Attachment;

/** This class does simple one level mime parsing to find crypto
  data.

  Use the mimedataprovider on a body or attachment stream. It
  will do the conversion from MIME to PGP / CMS data on the fly.
  Similarly when writing it will split up the data into a body /
  html body and attachments.

  A detached signature will be made available through the
  signature function.

  When reading the raw mime data from the underlying stream is
  "collected" and parsed into crypto data which is then
  buffered in an internal gpgme data stucture.

  For historicial reasons this class both provides reading
  and writing to be able to reuse the same mimeparser code.
  Similarly using the C-Style parsing code is for historic
  reason because as this class was created to have a data
  container unrelated of the Outlook Object model (after
  creation) the mimeparser code already existed and was
  stable.
*/
class MimeDataProvider : public GpgME::DataProvider
{
public:
  /* Create an empty dataprovider, useful for writing to. */
  MimeDataProvider(bool no_headers = false);
#ifdef HAVE_W32_SYSTEM
  /* Read and parse the stream. Does not hold a reference
     to the stream but releases it after read.

     If no_headers is set to true, assume that there are no
     headers and immediately start collecting crypto data.
     Eg. When decrypting a MOSS Attachment.
     */
  MimeDataProvider(LPSTREAM stream, bool no_headers = false);
#endif
  /* Test instrumentation. */
  MimeDataProvider(FILE *stream, bool no_headers = false);
  ~MimeDataProvider();

  /* Dataprovider interface */
  bool isSupported(Operation) const;

  /** Read some data from the stream. This triggers
    the conversion code interanally to convert mime
    data into PGP/CMS Data that GpgME can work with. */
  ssize_t read(void *buffer, size_t bufSize);

  ssize_t write(const void *buffer, size_t bufSize);

  /* Seek the underlying stream. This discards the internal
     buffers as the offset is not mapped. Should not really
     be used but can be used to reset the DataProvider. */
  off_t seek(off_t offset, int whence);

  /* Noop */
  void release() {}

  /* The the data of the signature part.

     If not null then this is a pointer to the signature
     data that is valid for the lifetime of this object.
  */
  GpgME::Data *signature() const;

  /* Add an attachment to the list */
  std::shared_ptr<Attachment> create_attachment();

  mime_context_t mime_context() {return m_mime_ctx;}

  /* Checks if there is body data left in the buffer e.g. for inline messages
     that did not end with a linefeed and adds it to body / returns the body. */
  const std::string &get_body();
  /* Similar for html body */
  const std::string &get_html_body();
  const std::vector <std::shared_ptr<Attachment> > get_attachments() const
    {return m_attachments;}
  const std::string &get_html_charset() const;
  const std::string &get_body_charset() const;

  void set_has_html_body(bool value) {m_has_html_body = value;}
private:
#ifdef HAVE_W32_SYSTEM
  /* Collect the data from mapi. */
  void collect_data(LPSTREAM stream);
#endif
  /* Collect data from a file. */
  void collect_data(FILE *stream);
  /* Collect a single line. */
  size_t collect_input_lines(const char *input, size_t size);
  /* A detached signature found in the input */
  std::string m_sig_data;
  /* The data to be passed to the crypto operation */
  GpgME::Data m_crypto_data;
  /* The plaintext body. */
  std::string m_body;
  /* The plaintext html body. */
  std::string m_html_body;
  /* A detachted signature found in the mail */
  GpgME::Data *m_signature;
  /* Internal helper to read line based */
  std::string m_rawbuf;
  /* The mime context */
  mime_context_t m_mime_ctx;
  /* List of attachments. */
  std::vector<std::shared_ptr<Attachment> > m_attachments;
  /* Charset of html */
  std::string m_html_charset;
  /* Charset of body */
  std::string m_body_charset;
  /* Do we have html at all */
  bool m_has_html_body;
  /* Collect everything */
  bool m_collect_everything;
};
#endif // MIMEDATAPROVIDER_H
