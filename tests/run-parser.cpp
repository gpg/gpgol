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
         "  --verbose        run in verbose mode\n"
         "  --type           GPGOL_MESSAGETYPE\n"
         , stderr);
  exit (ex);
}

int main(int argc, char **argv)
{
  int last_argc = -1;
  int msgtype = 0;
  FILE *fp_in = NULL;

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
          set_log_file ("stderr");
          argc--; argv++;
        }
      else if (!strcmp (*argv, "--type"))
        {
          msgtype = atoi (*(argv + 1));
          argc--; argv++;
          argc--; argv++;
        }
    }
  if (argc < 1 || argc > 2)
    show_usage (1);

  fp_in = fopen (argv[0], "rb");

  ParseController parser(fp_in, (msgtype_t)msgtype);
  parser.parse();
  std::cout << "Decrypt result:\n" << parser.decrypt_result();
  std::cout << "Verify result:\n" << parser.verify_result();
  std::cout << "BEGIN BODY\n" << parser.get_body() << "\nEND BODY";
  std::cout << "BEGIN HTML\n" << parser.get_html_body() << "\nEND HTML";
  for (auto attach: parser.get_attachments())
    {
      std::cout << "Attachment: " << attach->get_display_name();
    }
}
