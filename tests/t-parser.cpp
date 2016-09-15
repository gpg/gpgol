/* t-parser.cpp - Test for gpgOL's parser.
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

#include <stdio.h>
#include "parsecontroller.h"
#include <iostream>
#include "attachment.h"
#include <gpgme.h>

struct
{
  const char *input_file;
  msgtype_t type;
  const char *expected_body_file;
  const char *expected_html_body_file;
  const char *expected_return;
  int attachment_cnt;
} test_data[] = {
  { DATADIR "/inlinepgpencrypted.mbox",
    MSGTYPE_GPGOL_PGP_MESSAGE,
    DATADIR "/inlinepgpencrypted.plain",
    NULL,
    NULL,
    0 },
  { NULL, MSGTYPE_UNKNOWN, NULL, NULL, 0 }
};


int main()
{
  int i = 0;
  putenv ((char*) "GNUPGHOME=" GPGHOMEDIR);
  gpgme_check_version (NULL);

  while (test_data[i].input_file)
    {
      auto input = fopen (test_data[i].input_file, "rb");
      if (!input)
        {
          fprintf (stderr, "Failed to open input file: %s\n",
                   test_data[i].input_file);
          exit(1);
        }
      ParseController parser (input, test_data[i].type);

      auto result = parser.parse();

      auto decResult = parser.decrypt_result();
      auto verifyResult = parser.verify_result();

      if (decResult.error() || verifyResult.error())
        {
          std::cerr << "Decrypt or verify error:\n"
                    << decResult
                    << verifyResult;
          exit(1);
        }

      if (test_data[i].expected_return)
        {
          if (result != test_data[i].expected_return)
            {
              fprintf (stderr, "Failed to get expected return. Return was: %s\n",
                       result.c_str());
              exit(1);
            }
        }
      if (test_data[i].expected_body_file)
        {
          auto expected_body = fopen (test_data[i].expected_body_file, "rb");
          if (!expected_body)
            {
              fprintf (stderr, "Failed to open input file: %s\n",
                       test_data[i].expected_body_file);
              exit(1);
            }
          char bodybuf[16000];
          fread (bodybuf, 1, 16000, expected_body);
          if (parser.get_body() != bodybuf)
            {
              fprintf (stderr, "Body was: \n\"%s\"\nExpected:\n\"%s\"\n",
                       parser.get_body().c_str(), bodybuf);
              exit(1);
            }
        }
      if (test_data[i].expected_html_body_file)
        {
          auto expected_html_body = fopen (test_data[i].expected_html_body_file, "rb");
          if (!expected_html_body)
            {
              fprintf (stderr, "Failed to open input file: %s\n",
                       test_data[i].expected_html_body_file);
              exit(1);
            }
          char bodybuf[16000];
          fread (bodybuf, 1, 16000, expected_html_body);
          if (parser.get_html_body() != bodybuf)
            {
              fprintf (stderr, "HTML was: \n\"%s\"\nExpected:\n\"%s\"\n",
                       parser.get_html_body().c_str(), bodybuf);
              exit(1);
            }
        }
      if (test_data[i].attachment_cnt)
        {
          int actual = (int)parser.get_attachments().size();
          if (actual != test_data[i].attachment_cnt)
            {
              fprintf (stderr, "Attachment count mismatch. Actual: %i Expected: %i\n",
                       actual, test_data[i].attachment_cnt);
              exit(1);
            }
        }
      fprintf (stderr, "Pass: %s\n", test_data[i].input_file);
      i++;
    }
  exit(0);
}
