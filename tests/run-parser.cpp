/* run-parser.cpp - Test for gpgOL's parser.
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

static int
show_usage (int ex)
{
  fputs ("usage: run-parser [options] FILE\n\n"
         "Options:\n"
         "  --verbose             run in verbose mode\n"
         "  --multipart-signed    multipart/signed\n"
         "  --multipart-encrypted multipart/encrypted\n"
         "  --opaque-signed       SMIME opaque signed\n"
         "  --opaque-encrypted    SMIME opaque encrypted\n"
         "  --clear-signed        clearsigned\n"
         "  --pgp-message         inline pgp message\n"
         , stderr);
  exit (ex);
}

int main(int argc, char **argv)
{
  int last_argc = -1;
  msgtype_t msgtype = MSGTYPE_UNKNOWN;
  FILE *fp_in = NULL;

  putenv ((char*) "GNUPGHOME=" GPGHOMEDIR);
  gpgme_check_version (NULL);

  if (argc)
    { argc--; argv++; }

  while (argc && last_argc != argc )
    {
      last_argc = argc;
      if (!strcmp (*argv, "--"))
        {
          argc--; argv++;
          break;
        }
      else if (!strcmp (*argv, "--help"))
        show_usage (0);
      else if (!strcmp (*argv, "--verbose"))
        {
          opt.enable_debug |= DBG_MIME_PARSER;
          opt.enable_debug |= 1;
          set_log_file ("stderr");
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--multipart-signed"))
        {
          msgtype = MSGTYPE_GPGOL_MULTIPART_SIGNED;
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--multipart-encrypted"))
        {
          msgtype = MSGTYPE_GPGOL_MULTIPART_ENCRYPTED;
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--opaque-signed"))
        {
          msgtype = MSGTYPE_GPGOL_OPAQUE_SIGNED;
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--opaque-encrypted"))
        {
          msgtype = MSGTYPE_GPGOL_OPAQUE_ENCRYPTED;
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--clear-signed"))
        {
          msgtype = MSGTYPE_GPGOL_CLEAR_SIGNED;
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--pgp-message"))
        {
          msgtype = MSGTYPE_GPGOL_PGP_MESSAGE;
          argc--; argv++;
        }
    }
  if (argc < 1 || argc > 2)
    show_usage (1);

  fp_in = fopen (argv[0], "rb");

    {
      ParseController parser(fp_in, msgtype);
      std::cout << "Parse error: " << parser.get_formatted_error ();
      std::cout << "\nDecrypt result:\n" << parser.decrypt_result()
                << "\nVerify result:\n" << parser.verify_result()
                << "\nBEGIN BODY\n" << parser.get_body() << "\nEND BODY"
                << "\nBEGIN HTML\n" << parser.get_html_body() << "\nEND HTML";
      for (auto attach: parser.get_attachments())
        {
          std::cout << "Attachment: " << attach->get_display_name();
        }
    }
  fclose (fp_in);
}
