/* guidgen.c - Tool to create GUIDs
 *	Copyright (C) 2001 g10 Code GmbH
 *
 * This file is part of GPGME.
 *
 * GPGME is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 *
 * GPGME is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA
 */


#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <errno.h>
#include <assert.h>
#include <time.h>
#include <windows.h>



static void
create_guid ()
{
  GUID guid, *id;

  id = &guid;
  if ( CoCreateGuid (id) ) 
    {
      fprintf (stderr,"failed to create GUID\n");
      exit (1);
    }
  printf ("{%08lx-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}\n",
          id->Data1, id->Data2, id->Data3,
          id->Data4[0], id->Data4[1], id->Data4[2], id->Data4[3],
          id->Data4[4], id->Data4[5], id->Data4[6], id->Data4[7] );
}



int
main (int argc, char **argv )
{
  
  if (argc)
    {
      argv++;
      argc--;
    }

  if (!argc)
    create_guid();
  else
    {
      int n;
      
      for (n = atoi (argv[0]); n > 0; n-- )
        create_guid ();
    }
  
  return 0;
}

