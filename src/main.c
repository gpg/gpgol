/* main.c - DLL entry point
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGME Dialogs.
 *
 * GPGME Dialogs is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGME Dialogs is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGME Dialogs; if not, write to the Free Software Foundation, 
 * Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA 
 */

#include <config.h>

#include <windows.h>

#include <gpgme.h>

#include "mymapi.h"
#include "mymapitags.h"


#include "intern.h"
#include "passcache.h"
#include "msgcache.h"
#include "mymapi.h"


int WINAPI
DllMain (HINSTANCE hinst, DWORD reason, LPVOID reserved)
{
  if (reason == DLL_PROCESS_ATTACH )
    {
      set_global_hinstance (hinst);
      /* The next call initializes subsystems of gpgme and should be
         done as early as possible.  The actual return value is (the
         version string) is not used here.  It may be called at anty
         time later for this. */
      gpgme_check_version (NULL);

      /* Early initializations of our subsystems. */
      if (initialize_passcache ())
        return FALSE;
      if (initialize_msgcache ())
        return FALSE;
      if (initialize_mapi_gpgme ())
        return FALSE;
    }

  return TRUE;
}


