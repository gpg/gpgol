/* revert.cpp - Convert messages back to the orignal format
 *	Copyright (C) 2008 g10 Code GmbH
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
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "oomhelp.h"
#include "mapihelp.h"
#include "message.h"


/* Wrapper around UlRelease with error checking. */
static void 
ul_release (LPVOID punk, const char *func, int lnr)
{
  ULONG res;
  
  if (!punk)
    return;
  res = UlRelease (punk);
  if (opt.enable_debug & DBG_MEMORY)
    log_debug ("%s:%s:%d: UlRelease(%p) had %lu references\n", 
               SRCNAME, func, lnr, punk, res);
}



static int
message_revert (LPMESSAGE message)
{
  int result = 0;
  HRESULT hr;
  char *msgcls;
  char *oldmsgcls = NULL;
  mapi_attach_item_t *table;
  int tblidx;
  SPropValue prop;
  SPropTagArray proparray;
  LPATTACH att = NULL;
  ULONG tag;
  
  /* Check whether we need to care about this message.  */
  msgcls = mapi_get_message_class (message);
  log_debug ("%s:%s: message class is `%s'\n", 
             SRCNAME, __func__, msgcls? msgcls:"[none]");
  if ( !( !strncmp (msgcls, "IPM.Note.GpgOL", 14) 
          && (!msgcls[14] || msgcls[14] == '.') ) )
    {
      xfree (msgcls);
      return 0;  /* Not one of our message classes.  */
    }

  table = mapi_create_attach_table (message, 0);
  if (!table)
    {
      /* No attachments or an error building the table.  The first
         case should never happen for GpgOL messages and thus we can
         assume an error.  */
      xfree (msgcls);
      return -1; /* Error.  */
    }

  /* Copy a possible gpgolPGP.DAT attachments back to PR_BODY. */
  for (tblidx=0; !table[tblidx].end_of_table; tblidx++)
    if (table[tblidx].attach_type == ATTACHTYPE_PGPBODY
        && table[tblidx].filename 
        && !strcmp (table[tblidx].filename, PGPBODYFILENAME))
      break;
  if (!table[tblidx].end_of_table)
    {
      if (mapi_attachment_to_body (message, table+tblidx))
        {
          xfree (msgcls);
          mapi_release_attach_table (table);
          return -1; /* Error.  */
        }
      result = 2;
    }

  /* Delete all temporary attachments.  */
  for (tblidx=0; !table[tblidx].end_of_table; tblidx++)
    switch (table[tblidx].attach_type)
      {
      case ATTACHTYPE_FROMMOSS:
      case ATTACHTYPE_PGPBODY:
        hr = message->DeleteAttach (table[tblidx].mapipos, 0, NULL, 0);
        if (hr)
          {
            log_error ("%s:%s: deleting attachment %d failed: hr=%#lx",
                       SRCNAME, __func__, tblidx, hr);
            /* It is better to ignore the error than to risk a
               corrupted messages due to a non-working MAPI
               transaction mechanism. */
          }
        else
          result = 2;
        break;

      case ATTACHTYPE_MOSS:
        /* We need to remove the hidden flag. */
        hr = message->OpenAttach (table[tblidx].mapipos, NULL, 
                                  MAPI_BEST_ACCESS, &att);
        if (FAILED (hr))
          {
            log_error ("%s:%s: can't open attachment %d: hr=%#lx",
                       SRCNAME, __func__, tblidx, hr);
          }
        else
          {
            prop.ulPropTag = PR_ATTACHMENT_HIDDEN;
            prop.Value.b = FALSE;
            hr = HrSetOneProp (att, &prop);
            if (hr)
              {
                log_error ("%s:%s: can't clear hidden attach flag: hr=%#lx\n",
                           SRCNAME, __func__, hr); 
              }
            else if ( (hr = att->SaveChanges (KEEP_OPEN_READWRITE)) )
              {
                log_error ("%s:%s: SaveChanges(attachment) failed: hr=%#lx\n",
                           SRCNAME, __func__, hr); 
              }
            else
              result = 2;

            att->Release ();
            att = NULL;
          }
        break;

      default:
        break;
      }
  
  /* Remove the sig status flag.  */
  if (!get_gpgolsigstatus_tag (message, &tag) )
    {
      proparray.cValues = 1;
      proparray.aulPropTag[0] = tag;
      hr = message->DeleteProps (&proparray, NULL);
      if (hr)
        log_error ("%s:%s: deleting sig status property failed: hr=%#lx\n",
                   SRCNAME, __func__, hr);
    }

  /* Not change the message class.  */
  oldmsgcls = mapi_get_old_message_class (message);
  if (!oldmsgcls)
    {
      /* No saved message class, mangle the actual class.  */
      if (!strcmp (msgcls, "IPM.Note.GpgOL.ClearSigned")
          || !strcmp (msgcls, "IPM.Note.GpgOL.PGPMessage") )
        msgcls[8] = 0;
      else
        memcpy (msgcls+9, "SMIME", 5);
      oldmsgcls = msgcls;
      msgcls = NULL;
    }

  log_debug ("%s:%s: setting message class to `%s'%s\n",
             SRCNAME, __func__, oldmsgcls, msgcls? " (from saved class)":"");

  /* Change the message class. */
  prop.ulPropTag = PR_MESSAGE_CLASS_A;
  prop.Value.lpszA = oldmsgcls; 
  hr = message->SetProps (1, &prop, NULL);
  if (hr)
    {
      log_error ("%s:%s: can't set message class to `%s': hr=%#lx\n",
                 SRCNAME, __func__, oldmsgcls, hr);
      result = -1;
    }
  else if (result != 2)
    result = 1;

  mapi_release_attach_table (table);
  xfree (msgcls);
  xfree (oldmsgcls);
  return result;
}


/* If the message has been mangled by us (message class changed etc.)
   convert the message back to the original format.  This is for
   example useful if one wants to remove GpgOL and continue to read
   the messages with other plugins or the Outlook internal S/MIME
   code.  If this function is called for a message not processed by
   GpgOL it will do nothing.

   If DO_SAVE is set do 1, a SaveChanges with the arguments in
   SAVE_FLAGS will be called if required.  If DO_SAVE is set to 0 the
   caller needs to call SaveChanges.

   Return values are:

   -1 = error
    0 = message not touched
    1 = message class changed
    2 = message class, attachments and body changed.
*/
EXTERN_C LONG __stdcall
gpgol_message_revert (LPMESSAGE message, LONG do_save, ULONG save_flags)
{
  LONG rc;

  rc = message_revert (message);
  if (do_save && (rc == 1 || rc == 2))
    if (mapi_save_changes (message, save_flags))
      rc = -1;

  return rc;
}


/* Revert all messages in the MAPIFOLDEROBJ.  */
EXTERN_C LONG __stdcall
gpgol_folder_revert (LPDISPATCH mapifolderobj)
{
  HRESULT hr;
  LPDISPATCH disp;
  LPUNKNOWN unknown;
  LPMDB mdb;
  LPMAPIFOLDER folder;
  LONG proptype;
  char *store_entryid;
  size_t store_entryidlen;
  LPMAPISESSION session;
  LPMAPITABLE contents;
  SizedSPropTagArray (1L, proparr_entryid) = { 1L, {PR_ENTRYID} };
  LPSRowSet rows;
  ULONG mtype;
  LPMESSAGE message;

  log_debug ("%s:%s: Enter", SRCNAME, __func__);
  
  unknown = get_oom_iunknown (mapifolderobj, "MAPIOBJECT");
  if (!unknown)
    {
      log_error ("%s:%s: error getting MAPI object", SRCNAME, __func__);
      return -1;
    }
  folder = NULL;
  hr = unknown->QueryInterface (IID_IMAPIFolder, (void**)&folder);
  unknown->Release ();
  if (hr != S_OK || !folder)
    {
      log_error ("%s:%s: error getting IMAPIFolder: hr=%#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }

  mapi_get_int_prop ( folder, PR_OBJECT_TYPE, &proptype);
  if (proptype != MAPI_FOLDER)
    {
      log_error ("%s:%s: not called for a folder but an object of type %ld", 
                 SRCNAME, __func__, (long)proptype);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }

  disp = get_oom_object (mapifolderobj, "Session");
  if (!disp)
    {
      log_error ("%s:%s: session object not found", SRCNAME, __func__);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }
  unknown = get_oom_iunknown (disp, "MAPIOBJECT");
  disp->Release ();
  if (!unknown)
    {
      log_error ("%s:%s: error getting Session.MAPIOBJECT", SRCNAME, __func__);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }
  session = NULL;
  hr = unknown->QueryInterface (IID_IMAPISession, (void**)&session);
  unknown->Release ();
  if (hr != S_OK || !session)
    {
      log_error ("%s:%s: error getting IMAPISession: hr=%#lx",
                 SRCNAME, __func__, hr);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }
      
  /* Open the message store.  */
  store_entryid = mapi_get_binary_prop ((LPMESSAGE)folder, PR_STORE_ENTRYID,
                                        &store_entryidlen);
  if (!store_entryid)
    {
      log_error ("%s:%s: PR_STORE_ENTRYID missing\n",  SRCNAME, __func__);
      session->Release ();
      ul_release (folder, __func__, __LINE__);
      return -1;
    }
  mdb = NULL;
  hr = session->OpenMsgStore (0, store_entryidlen, (LPENTRYID)store_entryid,
                              NULL,  MAPI_BEST_ACCESS | MDB_NO_DIALOG, 
                              &mdb);
  xfree (store_entryid);
  if (FAILED (hr) || !mdb)
    {
      log_error ("%s:%s: OpenMsgStore failed: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      session->Release ();
      ul_release (folder, __func__, __LINE__);
      return -1;
    }

  /* Get the contents.  */
  contents = NULL;
  hr = folder->GetContentsTable ((ULONG)0, &contents);
  if (FAILED (hr) || !contents)
    {
      log_error ("%s:%s: error getting contents table: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      ul_release (mdb, __func__, __LINE__);
      session->Release ();
      ul_release (folder, __func__, __LINE__);
      return -1;
    }


  hr = contents->SetColumns ((LPSPropTagArray)&proparr_entryid, 0);
  if (FAILED (hr) )
    {
      log_error ("%s:%s: error setting contents table column: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      contents->Release ();
      ul_release (mdb, __func__, __LINE__);
      session->Release ();
      ul_release (folder, __func__, __LINE__);
      return -1;
    }

  /* Loop over all items.  We retrieve 50 rows at once.  */
  rows = NULL;
  while ( !(hr = contents->QueryRows (50, 0, &rows)) && rows && rows->cRows )
    {
      unsigned int rowidx;
      ULONG entrylen;
      LPENTRYID entry;

      log_debug ("%s:%s: Query Rows returned %d rows\n", 
                 SRCNAME, __func__, (int)rows->cRows);
      for (rowidx=0; rowidx < rows->cRows; rowidx++)
        {
          if (rows->aRow[rowidx].cValues < 1
              || rows->aRow[rowidx].lpProps[0].ulPropTag != PR_ENTRYID)
            {
              log_debug ("%s:%s: no PR_ENTRYID for this row - skipped\n", 
                         SRCNAME, __func__);
              continue;
            }

          entrylen = rows->aRow[rowidx].lpProps[0].Value.bin.cb;
          entry = (LPENTRYID)rows->aRow[rowidx].lpProps[0].Value.bin.lpb;

          hr = mdb->OpenEntry (entrylen, entry,
                               &IID_IMessage, MAPI_BEST_ACCESS, 
                               &mtype, (IUnknown**)&message);
          if (hr || !message)
            {
              log_debug ("%s:%s: failed to open entry of this row: hr=%#lx"
                         " - skipped\n",  SRCNAME, __func__, hr);
              continue;
            }

          if (mtype == MAPI_MESSAGE)
            {
              int rc = gpgol_message_revert (message, 1, 
                                             KEEP_OPEN_READWRITE|FORCE_SAVE);
              log_debug ("%s:%s: gpgol_message_revert returns %d\n", 
                         SRCNAME, __func__, rc);
            }
          else
            log_debug ("%s:%s: this row has now message object (type=%d)"
                       " - skipped\n",  SRCNAME, __func__, (int)mtype);

          message->Release ();
        }
      FreeProws (rows);
      rows = NULL;
    }
  if (hr)
    log_debug ("%s:%s: Query Rows failed: hr=%#lx\n", 
               SRCNAME, __func__, hr);
  if (rows)
    {
      FreeProws (rows);
      rows = NULL;
    }

  contents->Release ();
  ul_release (mdb, __func__, __LINE__);
  session->Release ();
  ul_release (folder, __func__, __LINE__);
  return 0;
}
