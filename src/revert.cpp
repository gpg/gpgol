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
#include "mimemaker.h"
#include "mail.h"

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

            gpgol_release (att);
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

/* Helper method for mailitem_revert to add changes on the mapi side
   and save them. */
static int finalize_mapi (LPMESSAGE message)
{
  HRESULT hr;
  SPropTagArray proparray;
  ULONG tag_id;

  if (get_gpgollastdecrypted_tag (message, &tag_id))
    {
      log_error ("%s:%s: can't getlastdecrypted tag",
                 SRCNAME, __func__);
      return -1;
    }
  proparray.cValues = 1;
  proparray.aulPropTag[0] = tag_id;
  hr = message->DeleteProps (&proparray, NULL);
  if (hr)
    {
      log_error ("%s:%s: failed to delete lastdecrypted tag",
                 SRCNAME, __func__);
      return -1;
    }

  return mapi_save_changes (message, FORCE_SAVE);
}

/* Similar to gpgol_message_revert but works on OOM and is
   used by the Ol > 2010 implementation.
   Doing this through OOM was necessary as the MAPI structures
   in the write event are not in sync with the OOM side.
   Trying to revert in the AfterWrite where MAPI is synced
   led to an additional save_changes after the wipe and
   so an additional sync.
   Updating the BODY through MAPI did not appear to work
   at all. Not sure why this is the case.
   Using the property accessor methods instead of
   MAPI properties might also not be necessary.

   Returns 0 on success, -1 on error. On error this
   function might leave plaintext in the mail.    */
EXTERN_C LONG __stdcall
gpgol_mailitem_revert (LPDISPATCH mailitem)
{
  LPDISPATCH attachments = NULL;
  LPMESSAGE message = NULL;
  char *item_str;
  char *msgcls = NULL;
  int i;
  int count = 0;
  LONG result = -1;
  msgtype_t msgtype;
  int body_restored = 0;
  LPDISPATCH *to_delete = NULL;
  int del_cnt = 0;
  LPDISPATCH to_restore = NULL;
  int mosstmpl_found = 0;
  int is_smime = 0;
  Mail *mail = NULL;

  /* Check whether we need to care about this message.  */
  msgcls = get_pa_string (mailitem, PR_MESSAGE_CLASS_W_DASL);
  log_debug ("%s:%s: message class is `%s'\n",
             SRCNAME, __func__, msgcls? msgcls:"[none]");
  if ( !( !strncmp (msgcls, "IPM.Note.GpgOL", 14)
          && (!msgcls[14] || msgcls[14] == '.') ) )
    {
      xfree (msgcls);
      log_error ("%s:%s: Message processed but not our class. Bug.",
                 SRCNAME, __func__);
      return -1;
    }

  mail = Mail::get_mail_for_item (mailitem);
  if (!mail)
    {
      xfree (msgcls);
      log_error ("%s:%s: No mail object for mailitem. Bug.",
                 SRCNAME, __func__);
      return -1;
    }
  is_smime = mail->is_smime ();

  message = get_oom_base_message (mailitem);
  attachments = get_oom_object (mailitem, "Attachments");

  if (!message)
    {
      log_error ("%s:%s: No message object.",
                 SRCNAME, __func__);
      goto done;
    }

  if (!attachments)
    {
      log_error ("%s:%s: No attachments object.",
                 SRCNAME, __func__);
      goto done;
    }
  msgtype = mapi_get_message_type (message);

  if (msgtype != MSGTYPE_GPGOL_PGP_MESSAGE &&
      msgtype != MSGTYPE_GPGOL_MULTIPART_ENCRYPTED &&
      msgtype != MSGTYPE_GPGOL_MULTIPART_SIGNED &&
      msgtype != MSGTYPE_GPGOL_OPAQUE_ENCRYPTED &&
      msgtype != MSGTYPE_GPGOL_OPAQUE_SIGNED)
    {
      log_error ("%s:%s: Revert not supported for msgtype: %i",
                 SRCNAME, __func__, msgtype);
      goto done;
    }


  count = get_oom_int (attachments, "Count");
  to_delete = (LPDISPATCH*) xmalloc (count * sizeof (LPDISPATCH));

  /* Yes the items start at 1! */
  for (i = 1; i <= count; i++)
    {
      LPDISPATCH attachment;
      attachtype_t att_type;

      if (gpgrt_asprintf (&item_str, "Item(%i)", i) == -1)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          goto done;
        }

      attachment = get_oom_object (attachments, item_str);
      xfree (item_str);
      if (!attachment)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          goto done;
        }

      if (get_pa_int (attachment, GPGOL_ATTACHTYPE_DASL, (int*) &att_type))
        {
          att_type = ATTACHTYPE_FROMMOSS;
        }

      switch (att_type)
        {
          case ATTACHTYPE_PGPBODY:
            {
              /* Restore Body */
              char *body = get_pa_string (attachment, PR_ATTACH_DATA_BIN_DASL);
              if (!body)
                {
                  log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
                  gpgol_release (attachment);
                  goto done;
                }
              log_debug ("%s:%s: Restoring pgp-body.",
                         SRCNAME, __func__);
              if (put_oom_string (mailitem, "Body", body))
                {
                  log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
                  xfree (body);
                  gpgol_release (attachment);
                  goto done;
                }
              body_restored = 1;
              xfree (body);
              to_delete[del_cnt++] = attachment;
              break;
            } /* No break we also want to delete that. */
          case ATTACHTYPE_MOSS:
            {
              char *mime_tag = get_pa_string (attachment,
                                              PR_ATTACH_MIME_TAG_DASL);
              if (!mime_tag)
                {
                  log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
                }
              else if (msgtype == MSGTYPE_GPGOL_MULTIPART_ENCRYPTED &&
                       !strcmp (mime_tag, "application/octet-stream"))
                {
                  /* This is the body attachment of a multipart encrypted
                     message. Rebuild the message. */
                  to_restore = attachment;
                  to_delete[del_cnt++] = attachment;
                }
              else if (msgtype == MSGTYPE_GPGOL_MULTIPART_SIGNED &&
                       mime_tag && !strcmp (mime_tag, "multipart/signed"))
                {
                  /* This is the MIME formatted MOSS attachment of a multipart
                     signed message. Rebuild the MIME structure from that.
                     This means treating it as a MOSSTMPL */
                  mosstmpl_found = 1;
                }
              else if (is_smime)
                {
                  /* Same here. No restoration but just rebuilding from the
                     attachment. */
                  mosstmpl_found = 1;
                }
              else
                {
                  log_oom ("%s:%s: Skipping attachment with tag: %s", SRCNAME,
                           __func__, mime_tag);
                  to_delete[del_cnt++] = attachment;
                }
              xfree (mime_tag);
              break;
            }
          case ATTACHTYPE_FROMMOSS:
          case ATTACHTYPE_FROMMOSS_DEC:
            {
              to_delete[del_cnt++] = attachment;
              break;
            }
          case ATTACHTYPE_MOSSTEMPL:
            /* This is a newly created attachment containing a MIME structure
               other clients could handle */
            {
              if (mosstmpl_found)
                {
                  log_error ("More then one mosstempl.");
                  goto done;
                }
              mosstmpl_found = 1;
              break;
            }
          default:
            to_delete[del_cnt++] = attachment;
        }
    }

  if (to_restore && !mosstmpl_found)
    {
      log_debug ("%s:%s: Restoring from MOSS.", SRCNAME, __func__);
      if (restore_msg_from_moss (message, to_restore, msgtype,
                                 msgcls))
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__,
                     __LINE__);
        }
      else
        {
          to_restore = NULL;
        }
    }
  if (to_restore || mosstmpl_found)
    {
      HRESULT hr;
      SPropValue prop;
      /* Message was either restored or the only attachment is the
         mosstmplate in which case we need to activate the
         MultipartSigned magic.*/
      prop.ulPropTag = PR_MESSAGE_CLASS_A;
      if (is_smime)
        {
#if 0
          /* FIXME this does not appear to work somehow. */
          if (opt.enable_smime)
            {
              prop.Value.lpszA =
                (char*) "IPM.Note.InfoPathForm.GpgOL.SMIME.MultipartSigned";
              hr = HrSetOneProp (message, &prop);
            }
          else
#endif
            {
              ULONG tag;
              if (msgtype == MSGTYPE_GPGOL_MULTIPART_SIGNED)
                prop.Value.lpszA = (char*) "IPM.Note.SMIME.MultipartSigned";
              else
                prop.Value.lpszA = (char*) "IPM.Note.SMIME";
              hr = HrSetOneProp (message, &prop);

              if (!get_gpgolmsgclass_tag (message, &tag))
                {
                  SPropTagArray proparray;
                  proparray.cValues = 1;
                  proparray.aulPropTag[0] = tag;
                  hr = message->DeleteProps (&proparray, NULL);
                  if (hr)
                    {
                      log_error ("%s:%s: deleteprops smime failed: hr=%#lx\n",
                                 SRCNAME, __func__, hr);

                    }
                }
            }
        }
      else if (msgtype == MSGTYPE_GPGOL_MULTIPART_SIGNED)
        {
          prop.Value.lpszA =
            (char*) "IPM.Note.InfoPathForm.GpgOLS.SMIME.MultipartSigned";
          hr = HrSetOneProp (message, &prop);
        }
      else
        {
          prop.Value.lpszA =
            (char*) "IPM.Note.InfoPathForm.GpgOL.SMIME.MultipartSigned";
          hr = HrSetOneProp (message, &prop);
        }
      if (hr)
        {
          log_error ("%s:%s: error setting the message class: hr=%#lx\n",
                     SRCNAME, __func__, hr);
          goto done;
        }

      /* Backup the real message class */
      if (!is_smime || opt.enable_smime)
        {
          if (mapi_set_gpgol_msg_class (message, msgcls))
            {
              log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
              goto done;
            }
        }
      else if (is_smime && !opt.enable_smime)
        {
          /* SMIME is disabled remove our categories. */
          mail->remove_categories ();
        }
    }

  result = 0;
done:

  /* Do the deletion body wipe even on error. */

  for (i = 0; i < del_cnt; i++)
    {
      LPDISPATCH attachment = to_delete[i];

      if (attachment == to_restore)
        {
          /* If restoring failed to restore is still set. In that case
             do not delete the MOSS attachment to avoid data loss. */
          continue;
        }
      /* Delete the attachments that are marked to delete */
      if (invoke_oom_method (attachment, "Delete", NULL))
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          result = -1;
        }
    }
  if (!body_restored && put_oom_string (mailitem, "Body", ""))
    {
      log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
      result = -1;
    }

  for (i = 0; i < del_cnt; i++)
    {
      gpgol_release (to_delete[i]);
    }

  xfree (to_delete);
  gpgol_release (attachments);
  xfree (msgcls);

  if (!result && finalize_mapi (message))
    {
      log_error ("%s:%s: Finalize failed.",
                 SRCNAME, __func__);
      result = -1;
    }

  gpgol_release (message);

  return result;
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
  gpgol_release (unknown);
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
  gpgol_release (disp);
  if (!unknown)
    {
      log_error ("%s:%s: error getting Session.MAPIOBJECT", SRCNAME, __func__);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }
  session = NULL;
  hr = unknown->QueryInterface (IID_IMAPISession, (void**)&session);
  gpgol_release (unknown);
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
      gpgol_release (session);
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
      gpgol_release (session);
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
      gpgol_release (session);
      ul_release (folder, __func__, __LINE__);
      return -1;
    }


  hr = contents->SetColumns ((LPSPropTagArray)&proparr_entryid, 0);
  if (FAILED (hr) )
    {
      log_error ("%s:%s: error setting contents table column: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      gpgol_release (contents);
      ul_release (mdb, __func__, __LINE__);
      gpgol_release (session);
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

          gpgol_release (message);
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

  gpgol_release (contents);
  ul_release (mdb, __func__, __LINE__);
  gpgol_release (session);
  ul_release (folder, __func__, __LINE__);
  return 0;
}
