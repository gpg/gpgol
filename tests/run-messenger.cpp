/* run-messenger.cpp - Test for GpgOL's external windowmessage API.
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


#include <windows.h>

#include <stdio.h>

static int
show_usage (int ex)
{
  fputs ("usage: run-messgenger id [PAYLOAD]\n\n"
         "Options:\n"
         , stderr);
  exit (ex);
}


int main(int argc, char **argv)
{
  int last_argc = -1;
  if (argc)
    { argc--; argv++; }

  while (argc && last_argc != argc )
    {
      last_argc = argc;
      if (!strcmp (*argv, "--help"))
        {
          show_usage (0);
        }
    }

  if (argc != 1 && argc != 2)
    {
      show_usage (1);
    }

  int id = atoi (*argv);

  HWND gpgol = FindWindowA ("GpgOLResponder", "GpgOLResponder");

  if (!gpgol)
    {
      fprintf (stderr, "Failed to find GpgOL Window");
      exit (1);
    }

  if (argc == 1)
    {
      printf ("Sending message: %i\n", id);
      SendMessage (gpgol, WM_USER + 43, id, NULL);
      exit (0);
    }

  /* Send message with payload */
  char *payload = argv[1];
  COPYDATASTRUCT cds;

  cds.dwData = id;
  cds.cbData = strlen (payload) + 1;
  cds.lpData = payload;

  printf ("Sending message: %i\n with param: %s", id, payload);
  SendMessage (gpgol, WM_COPYDATA, 0, (LPARAM) &cds);
  exit (0);
}
