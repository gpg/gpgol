/* mimedataprover.h - GpgME dataprovider for mime data
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
#ifndef MIMEDATAPROVIDER_H
#define MIMEDATAPROVIDER_H

#include <gpgme++/interfaces/dataprovider.h>
#include <gpgme++/data.h>
#include "oomhelp.h"
#include "mapihelp.h"
#include "rfc822parse.h"

#include <string>

/** This class does simple one level mime parsing to find crypto
  data.

  Use the mimedataprovider on a body or attachment stream. It
  will do the conversion from MIME to PGP / CMS data on the fly.

  The raw mime data from the underlying stream is "collected" and
  parsed into Crypto data which is then buffered in "databuf".
*/
class MimeDataProvider : public GpgME::DataProvider
{
public:
  /* Read and parse the stream. Does not hold a reference
     to the stream but releases it after read. */
  MimeDataProvider(LPSTREAM stream);
  ~MimeDataProvider();

  /* Dataprovider interface */
  bool isSupported(Operation) const;

  /** Read some data from the stream. This triggers
    the conversion code interanally to convert mime
    data into PGP/CMS Data that GpgME can work with. */
  ssize_t read(void *buffer, size_t bufSize);
  ssize_t write(const void *buffer, size_t bufSize) {
      (void)buffer; (void)bufSize; return -1;
  }
  /* Seek the underlying stream. This discards the internal
     buffers as the offset is not mapped. Should not really
     be used but can be used to reset the DataProvider. */
  off_t seek(off_t offset, int whence);
  /* Noop */
  void release() {}

  /* The the data of the signature part. */
  const GpgME::Data &get_signature_data();
private:
  /* Collect the crypto data from mime. */
  void collect_data(LPSTREAM stream);
  /* Collect a single line. */
  size_t collect_input_lines(const char *input, size_t size);
  /* Move actual data into the databuffer. */
  void decode_and_collect(char *line, size_t pos);
  enum Encoding {None, Base64, Quoted};
  std::string m_sig_data;
  GpgME::Data m_data;
  GpgME::Data m_signature;
  std::string m_rawbuf;
  bool m_collect;
  rfc822parse_t m_parser;
  Encoding m_current_encoding;
  b64_state_t m_base64_context;
};
#endif // MIMEDATAPROVIDER_H
