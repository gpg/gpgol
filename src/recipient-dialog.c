/* recipient-dialog.c
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#include <config.h>

#include <windows.h>
#include <commctrl.h>
#include <time.h>
#include <gpgme.h>

#include "gpgol-ids.h"
#include "intern.h"
#include "util.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


struct recipient_cb_s 
{
  char **unknown_keys;  /* A string array with the names of the
                           unknown recipients. */

  char **fnd_keys;      /* A string array with the user IDs of already
                           found keys. */

  /* A bit vector used to return selected options. */
  unsigned int opts;

  /* A key array to hold all keys. */
  gpgme_key_t *keyarray;
  size_t keyarray_count;

  /* The result key Array, i.e. the selected keys.  This array is NULL
     terminated. */
  gpgme_key_t *selected_keys;
  size_t      selected_keys_count;
};

struct key_item_s 
{
  char name [150];
  char e_mail[100];
  char key_info[64];
  char keyid[32];
  char validity[32];
  char idx[20];
};


static void
initialize_rsetbox (HWND hwnd)
{
    LVCOLUMN col;

    memset (&col, 0, sizeof (col));
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    col.pszText = "Name";
    col.cx = 100;
    col.iSubItem = 0;
    ListView_InsertColumn (hwnd, 0, &col);

    col.pszText = "E-Mail";
    col.cx = 100;
    col.iSubItem = 1;
    ListView_InsertColumn( hwnd, 1, &col );

    col.pszText = "Key-Info";
    col.cx = 110;
    col.iSubItem = 2;
    ListView_InsertColumn( hwnd, 2, &col );

    col.pszText = "Key ID";
    col.cx = 70;
    col.iSubItem = 3;
    ListView_InsertColumn( hwnd, 3, &col );

    col.pszText = "Validity";
    col.cx = 70;
    col.iSubItem = 4;
    ListView_InsertColumn( hwnd, 4, &col );

    col.pszText = "Index";
    col.cx = 0;  /* Hide it. */
    col.iSubItem = 5;
    ListView_InsertColumn( hwnd, 5, &col );

/*     ListView_SetExtendedListViewStyleEx( hwnd, 0, LVS_EX_FULLROWSELECT ); */
}


static gpgme_key_t *
load_rsetbox (HWND hwnd, size_t *r_arraysize)
{
  LVITEM lvi;
  gpg_error_t err;
  gpgme_ctx_t ctx;
  gpgme_key_t key = NULL;
  gpgme_key_t *keyarray;
  size_t keyarray_size;
  size_t pos;
  char keybuf[128], *s;
  const char *trust_items[] = 
    {
      "UNKNOWN",
      "UNDEFINED",
      "NEVER",
      "MARGINAL",
      "FULL",
      "ULTIMATE"
    };
  enum {COL_NAME, COL_EMAIL, COL_KEYINF, COL_KEYID, COL_TRUST, COL_IDX};
  DWORD val;

  memset (&lvi, 0, sizeof (lvi));

  err = gpgme_new (&ctx);
  if (err)
    return NULL;

  err = gpgme_op_keylist_start (ctx, NULL, 0);
  if (err)
    {
      log_error ("failed to initiate key listing: %s\n", gpg_strerror (err));
      gpgme_release (ctx);
      return NULL;
    }

  keyarray_size = 500; 
  keyarray = xcalloc (keyarray_size, sizeof *keyarray);
  pos = 0;
                 
  while (!gpgme_op_keylist_next (ctx, &key)) 
    {
      /* We only want keys capable of encrypting. */
      if (!key->can_encrypt)
        {
          gpgme_key_release (key);
          continue;
        }
      
      /* Check that the primary key is *not* revoked, expired or invalid */
      if (key->revoked || key->expired || key->invalid || key->disabled)
        {
          gpgme_key_release (key);
          continue;
        }

      /* Ignore keys without a user ID or woithout a subkey */
      if (!key->uids || !key->subkeys )
        {
          gpgme_key_release (key);
          continue;
        }

      ListView_InsertItem (hwnd, &lvi);
      
      s = key->uids->name;
      ListView_SetItemText (hwnd, 0, COL_NAME, s);
      
      s = key->uids->email;
      ListView_SetItemText (hwnd, 0, COL_EMAIL, s);

      s = keybuf;
      *s = 0;
      s = stpcpy (s, get_pubkey_algo_str (key->subkeys->pubkey_algo));
      if (key->subkeys->next)
        {
          /* Fixme: This is not really correct because we don't know
             which encryption subkey gpg is going to select. Same
             holds true for the key length below. */
          *s++ = '/';
          s = stpcpy (s, get_pubkey_algo_str
                      (key->subkeys->next->pubkey_algo));
        }      
      
      *s++ = ' ';
      if (key->subkeys->next)
        sprintf (s, "%d", key->subkeys->next->length);
      else
        sprintf (s, "%d", key->subkeys->length);

      s = keybuf;
      ListView_SetItemText (hwnd, 0, COL_KEYINF, s);
      
      if (key->subkeys->keyid  && strlen (key->subkeys->keyid) > 8)
        ListView_SetItemText (hwnd, 0, COL_KEYID, key->subkeys->keyid+8);
      
      val = key->uids->validity;
      if (val < 0 || val > 5) 
	val = 0;
      strcpy (keybuf, trust_items[val]);
      s = keybuf;
      ListView_SetItemText (hwnd, 0, COL_TRUST, s);

      /* I'd like to use SetItemData but that one is only available as
         a member function of CListCtrl; I haved not figured out how
         the vtable is made up.  Thus we use a string with the index. */
      sprintf (keybuf, "%u", (unsigned int)pos);
      s = keybuf;
      ListView_SetItemText (hwnd, 0, COL_IDX, s);

      if (pos >= keyarray_size)
        {
          gpgme_key_t *tmparr;
          size_t i;

          keyarray_size += 500;
          tmparr = xcalloc (keyarray_size, sizeof *tmparr);
          for (i=0; i < pos; i++)
            tmparr[i] = keyarray[i];
          xfree (keyarray);
          keyarray = tmparr;
        }
      keyarray[pos++] = key;

    }

  gpgme_op_keylist_end (ctx);
  gpgme_release (ctx);

  *r_arraysize = pos;
  return keyarray;
}


/* Release the key array ARRAY as well as all COUNT keys. */
static void
release_keyarray (gpgme_key_t *array, size_t count)
{
  size_t n;

  if (array)
    {
      for (n=0; n < count; n++)
        gpgme_key_release (array[n]);
      xfree (array);
    }
}



static void
copy_item (HWND dlg, int id_from, int pos)
{
  HWND src, dst;
  LVITEM lvi;
  struct key_item_s from;
  int idx = pos;
  
  src = GetDlgItem (dlg, id_from);
  dst = GetDlgItem (dlg, id_from==IDC_ENC_RSET1 ?
                    IDC_ENC_RSET2 : IDC_ENC_RSET1);
  
  if (idx == -1)
    {
      idx = ListView_GetNextItem (src, -1, LVNI_SELECTED);
      if (idx == -1)
        return;
    }
  
  memset (&from, 0, sizeof (from));
  ListView_GetItemText (src, idx, 0, from.name, sizeof (from.name)-1);
  ListView_GetItemText (src, idx, 1, from.e_mail, sizeof (from.e_mail)-1);
  ListView_GetItemText (src, idx, 2, from.key_info, sizeof (from.key_info)-1);
  ListView_GetItemText (src, idx, 3, from.keyid, sizeof (from.keyid)-1);
  ListView_GetItemText (src, idx, 4, from.validity, sizeof (from.validity)-1);
  ListView_GetItemText (src, idx, 5, from.idx, sizeof (from.idx)-1);

  ListView_DeleteItem (src, idx);
  
  memset (&lvi, 0, sizeof (lvi));
  ListView_InsertItem (dst, &lvi);
  ListView_SetItemText (dst, 0, 0, from.name);
  ListView_SetItemText (dst, 0, 1, from.e_mail);
  ListView_SetItemText (dst, 0, 2, from.key_info);
  ListView_SetItemText (dst, 0, 3, from.keyid);
  ListView_SetItemText (dst, 0, 4, from.validity);
  ListView_SetItemText (dst, 0, 5, from.idx);
}


static int
find_item (HWND hwnd, const char *str)
{
    LVFINDINFO fnd;
    int pos;

    memset (&fnd, 0, sizeof (fnd));
    fnd.flags = LVFI_STRING;
    fnd.psz = str;
    pos = ListView_FindItem (hwnd, -1, &fnd);
    if (pos != -1)
	return pos;
    return -1;
}


static void
initialize_keybox (HWND dlg, struct recipient_cb_s *cb)
{
  size_t i;
  HWND box = GetDlgItem (dlg, IDC_ENC_NOTFOUND);
  HWND rset = GetDlgItem (dlg, IDC_ENC_RSET1);
  int n;
  
  if (cb->unknown_keys)
    {
      for (i=0; cb->unknown_keys[i]; i++)
        SendMessage (box, LB_ADDSTRING, 0,
                     (LPARAM)(const char *)cb->unknown_keys[i]);
    }

  if (cb->fnd_keys)
    {
      for (i=0; cb->fnd_keys[i]; i++) 
        {
          n = find_item (rset, cb->fnd_keys[i]);
          if (n != -1)
            copy_item (dlg, IDC_ENC_RSET1, n);
        }
    }
}


BOOL CALLBACK
recipient_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
  static struct recipient_cb_s * rset_cb;
  static int rset_state = 1;
  NMHDR * notify;
  HWND hrset;
  const char *warn;
  size_t pos;
  int i;

  switch (msg) 
    {
    case WM_INITDIALOG:
      rset_cb = (struct recipient_cb_s *)lparam;

      initialize_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET1));
      rset_cb->keyarray = load_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET1),
                                        &rset_cb->keyarray_count );

      initialize_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET2));

      if (rset_cb->unknown_keys)
        initialize_keybox (dlg, rset_cb);
      else
        {
          /* No unknown keys - hide the not required windows. */
          ShowWindow (GetDlgItem (dlg, IDC_ENC_INFO), SW_HIDE);
          ShowWindow (GetDlgItem (dlg, IDC_ENC_NOTFOUND), SW_HIDE);
	}

      center_window (dlg, NULL);
      SetForegroundWindow (dlg);
      return TRUE;

    case WM_SYSCOMMAND:
      if (wparam == SC_CLOSE)
        EndDialog (dlg, TRUE);
      break;

    case WM_NOTIFY:
      notify = (LPNMHDR)lparam;
      if (notify && notify->code == NM_DBLCLK
          && (notify->idFrom == IDC_ENC_RSET1
              || notify->idFrom == IDC_ENC_RSET2))
        copy_item (dlg, notify->idFrom, -1);
      break;

    case WM_COMMAND:
      switch (HIWORD (wparam))
        {
	case BN_CLICKED:
          if ((int)LOWORD (wparam) == IDC_ENC_OPTSYM)
            {
              rset_state ^= 1;
              EnableWindow (GetDlgItem (dlg, IDC_ENC_RSET1), rset_state);
              EnableWindow (GetDlgItem (dlg, IDC_ENC_RSET2), rset_state);
              ListView_DeleteAllItems (GetDlgItem (dlg, IDC_ENC_RSET2));
	    }
          break;
	}

      switch ( LOWORD (wparam) ) 
        {
	case IDOK:
          hrset = GetDlgItem (dlg, IDC_ENC_RSET2);
          if (ListView_GetItemCount (hrset) == 0) 
            {
              MessageBox (dlg, "Please select at least one recipient key.",
                          "Recipient Dialog", MB_ICONINFORMATION|MB_OK);
              return FALSE;
	    }

          rset_cb->selected_keys_count = ListView_GetItemCount (hrset);
          rset_cb->selected_keys = xcalloc (rset_cb->selected_keys_count + 1,
                                            sizeof *rset_cb->selected_keys);
          for (i=0, pos=0; i < rset_cb->selected_keys_count; i++) 
            {
              gpgme_key_t key;
              int idata;
              char tmpbuf[30];

              *tmpbuf = 0;
              ListView_GetItemText (hrset, i, 5, tmpbuf, sizeof tmpbuf - 1);
              idata = *tmpbuf? strtol (tmpbuf, NULL, 10) : -1;
              if (idata >= 0 && idata < rset_cb->keyarray_count)
                {
                  key = rset_cb->keyarray[idata];
                  gpgme_key_ref (key);
                  rset_cb->selected_keys[pos++] = key;

                  switch (key->uids->validity)
                    {
                    case GPGME_VALIDITY_FULL:
                    case GPGME_VALIDITY_ULTIMATE:
                      break;
                    default:
                      /* Force encryption if one key is not fully
                         trusted.  Actually this is a bit silly but
                         supposedly here to allow adding an option to
                         disable this "feature".  */
                      rset_cb->opts |= OPT_FLAG_FORCE;
                      break;
                    }
                }
              else
                log_debug ("List item not correctly initialized - ignored\n");
            }
          rset_cb->selected_keys_count = pos;
          EndDialog (dlg, TRUE);
          break;

	case IDCANCEL:
          warn = _("If you cancel this dialog, the message will be sent"
                   " in cleartext.\n\n"
                   "Do you really want to cancel?");
          i = MessageBox (dlg, warn, "Recipient Dialog",
                          MB_ICONWARNING|MB_YESNO);
          if (i != IDNO)
            {
              rset_cb->opts = OPT_FLAG_CANCEL;
              EndDialog (dlg, FALSE);
            }
          break;
	}
      break;
    }
  return FALSE;
}



/* Display a recipient dialog to select keys and return all selected
   keys in RET_RSET.  Returns the selected options which may include
   OPT_FLAG_CANCEL.  */
unsigned int 
recipient_dialog_box (gpgme_key_t **ret_rset)
{
  struct recipient_cb_s cb;
  
  *ret_rset = NULL;

  memset (&cb, 0, sizeof (cb));
  DialogBoxParam (glob_hinst, (LPCTSTR)IDD_ENC, GetDesktopWindow(),
                  recipient_dlg_proc, (LPARAM)&cb);
  if (cb.opts & OPT_FLAG_CANCEL)
    release_keyarray (cb.selected_keys, cb.selected_keys_count);
  else
    *ret_rset = cb.selected_keys;
  release_keyarray (cb.keyarray, cb.keyarray_count);
  return cb.opts;
}


/* Exactly like recipient_dialog_box with the difference, that this
   dialog is used when some recipients were not found due to automatic
   selection. In such a case, the dialog displays the found recipients
   and the listbox contains the items which were _not_ found.  FND is
   a NULL terminated array with the keys we already found, UNKNOWn is
   a string array with names of recipients for whom we don't have a
   key yet.  RET_RSET returs a NULL termintated array with all
   selected keys.  The function returns the selected options which may
   include OPT_FLAG_CANCEL.
*/
unsigned int
recipient_dialog_box2 (gpgme_key_t *fnd, char **unknown,
		       gpgme_key_t **ret_rset)
{
  struct recipient_cb_s cb;
  int i;
  size_t n;

  *ret_rset = NULL;

  memset (&cb, 0, sizeof (cb));

  for (n=0; fnd[n]; n++)
    ;
  cb.fnd_keys = xcalloc (n+1, sizeof *cb.fnd_keys);

  for (i = 0; i < n; i++) 
    {
      if (fnd[i] && fnd[i]->uids && fnd[i]->uids->uid)
        cb.fnd_keys[i] = xstrdup (fnd[i]->uids->uid);
      else
	cb.fnd_keys[i] = xstrdup (_("User-ID not found"));
    }

  cb.unknown_keys = unknown;

  DialogBoxParam (glob_hinst, (LPCTSTR)IDD_ENC, GetDesktopWindow (),
		  recipient_dlg_proc, (LPARAM)&cb);

  if (cb.opts & OPT_FLAG_CANCEL)
    release_keyarray (cb.selected_keys, cb.selected_keys_count);
  else
    *ret_rset = cb.selected_keys;

  release_keyarray (cb.keyarray, cb.keyarray_count);
  for (i = 0; i < n; i++) 
    xfree (cb.fnd_keys[i]);
  xfree (cb.fnd_keys);
  return cb.opts;
}
