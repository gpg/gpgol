/* recipient-dialog.c
 *	Copyright (C) 2004 Timo Schulz
 *	Copyright (C) 2005, 2006, 2007 g10 Code GmbH
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

#ifndef _WIN32_IE /* allow to use advanced list view modes. */
#define _WIN32_IE 0x0600
#endif

#include <windows.h>
#include <commctrl.h>
#include <time.h>
#include <gpgme.h>
#include <assert.h>

#include "common.h"
#include "gpgol-ids.h"
#include "dialogs.h"


#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)


struct recipient_cb_s 
{
  char **unknown_keys;  /* A string array with the names of the
                           unknown recipients. */

  gpgme_key_t *fnd_keys; /* email address to key mapping list. */  

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


/* Symbolic column IDs. */
enum klist_col_t 
{
  KL_COL_NAME = 0,
  KL_COL_EMAIL = 1,
  KL_COL_INFO = 2,
  KL_COL_KEYID = 3,
  KL_COL_VALID = 4,
  /* number of columns. */
  KL_COL_N = 5
};

/* Insert the columns, needed to display keys, into the list view HWND. */
static void
initialize_rsetbox (HWND hwnd)
{
  LVCOLUMN col;

  /* We cannot avoid the casting here because gettext always returns
     a constant string but the listview interface needs char*. */
  memset (&col, 0, sizeof (col));
  col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
  col.pszText = (char*)_("Name");
  col.cx = 100;
  col.iSubItem = KL_COL_NAME;
  ListView_InsertColumn (hwnd, KL_COL_NAME, &col);
  
  col.pszText = (char*)_("E-Mail");
  col.cx = 100;
  col.iSubItem = KL_COL_EMAIL;
  ListView_InsertColumn (hwnd, KL_COL_EMAIL, &col);
  
  col.pszText = (char*)_("Key-Info");
  col.cx = 100;
  col.iSubItem = KL_COL_INFO;
  ListView_InsertColumn (hwnd, KL_COL_INFO, &col);
  
  col.pszText = (char*)_("Key ID");
  col.cx = 80;
  col.iSubItem = KL_COL_KEYID;
  ListView_InsertColumn (hwnd, KL_COL_KEYID, &col);
  
  col.pszText = (char*)_("Validity");
  col.cx = 70;
  col.iSubItem = KL_COL_VALID;
  ListView_InsertColumn (hwnd, KL_COL_VALID, &col);
  
  ListView_SetExtendedListViewStyleEx (hwnd, 0, LVS_EX_FULLROWSELECT);
}


/* Load the key list view control HWND with all useable keys 
   for encryption. Return the array size in *R_ARRAYSIZE. */
static gpgme_key_t*
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
      "Unknown",
      "Undefined",
      "Never",
      "Marginal",
      "Full",
      "Ultimate"
    };
  enum {COL_NAME, COL_EMAIL, COL_KEYINF, COL_KEYID, COL_TRUST};
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

      /* Ignore keys without a user ID or without a subkey */
      if (!key->uids || !key->subkeys)
        {
          gpgme_key_release (key);
          continue;
        }

      /* Store the position in the opaque param. */
      lvi.mask = LVIF_PARAM;
      lvi.lParam = (LPARAM)pos;
      ListView_InsertItem (hwnd, &lvi);
      
      s = utf8_to_native (key->uids->name);
      ListView_SetItemText (hwnd, 0, COL_NAME, s);
      xfree (s);
      
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
      
      if (key->subkeys->keyid && strlen (key->subkeys->keyid) > 8) 
	{
	  _snprintf (keybuf, sizeof (keybuf)-1, "0x%s", key->subkeys->keyid+8);
	  ListView_SetItemText (hwnd, 0, COL_KEYID, keybuf);
	}
      
      val = key->uids->validity;
      if (val < 0 || val > 5) 
	val = 0;
      strcpy (keybuf, trust_items[val]);
      s = keybuf;
      ListView_SetItemText (hwnd, 0, COL_TRUST, s);

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

  if (!array)
    return;
  for (n=0; n < count; n++)
    gpgme_key_release (array[n]);
  xfree (array);
}


/* Default maximal text size for a column. */
#define ITEMSIZE 200

/* Return the opaque param of the item with the index IDX.
   If the function call failed, return -1. */
static LPARAM
lv_get_item_param (HWND hwnd, int idx)
{
  LVITEM lv;
  
  memset (&lv, 0, sizeof (lv));
  lv.mask = LVIF_PARAM;
  lv.iItem = idx;
  if (!ListView_GetItem (hwnd, &lv))
    return (LPARAM)-1;
  return lv.lParam;
}

/* Copy one list view item from one view to another. */
static void
copy_item (HWND dlg, int id_from, int pos)
{
  HWND src, dst;
  LVITEM lvi;
  char item[KL_COL_N][ITEMSIZE];
  int idx = pos, i;
  int lparam;
  
  src = GetDlgItem (dlg, id_from);
  dst = GetDlgItem (dlg, id_from==IDC_ENC_RSET1 ?
                    IDC_ENC_RSET2 : IDC_ENC_RSET1);
  
  if (idx == -1)
    {
      idx = ListView_GetNextItem (src, -1, LVNI_SELECTED);
      if (idx == -1)
        return;
    }
  
  for (i=0; i < KL_COL_N; i++)
    ListView_GetItemText (src, idx, i, item[i], ITEMSIZE-1);

  /* Before we delete the item, we backup the lparam which
     holds the position to copy it to the new item. */
  lparam = (int)lv_get_item_param (src, idx);
  ListView_DeleteItem (src, idx);
  
  /* Add the lparam value from the source item. */
  memset (&lvi, 0, sizeof (lvi));
  lvi.mask = LVIF_PARAM;
  lvi.lParam = lparam;
  ListView_InsertItem (dst, &lvi);
  
  for (i=0; i < KL_COL_N; i++)
    ListView_SetItemText (dst, 0, i, item[i]);
}


/* Try to find an item with STR as the text in the first column.
   Return the index of the item or -1 if no item was found. */
static int
find_item (HWND hwnd, const char *str)
{
  LVFINDINFO fnd;
  int pos;
  
  memset (&fnd, 0, sizeof (fnd));
  fnd.flags = LVFI_STRING|LVFI_PARTIAL;;
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

  if (!cb->fnd_keys)
    return;
  
  /* Copy all requested keys into the second recipient listview
     to indicate that these key were automatically picked via
     the 'From' mailing header. */
  for (i=0; cb->fnd_keys[i]; i++) 
    {
      char *uid = utf8_to_native (cb->fnd_keys[i]->uids->name);
      
      n = find_item (rset, uid);
      if (n != -1)
	copy_item (dlg, IDC_ENC_RSET1, n);
      xfree (uid);
    }  
}

/* To avoid writing a dialog template for each language we use gettext
   for the labels and hope that there is enough space in the dialog to
   fit teh longest translation.  */
static void
recipient_dlg_set_labels (HWND dlg)
{
  static struct { int itemid; const char *label; } labels[] = {
    { IDC_ENC_RSET2_T,    N_("Selected recipients:")},
    { IDC_ENC_NOTFOUND_T, N_("Recipient which were NOT found")},
    { IDCANCEL,           N_("&Cancel")},
    { 0, NULL}
  };
  int i;

  for (i=0; labels[i].itemid; i++)
    SetDlgItemText (dlg, labels[i].itemid, _(labels[i].label));
}  


BOOL CALLBACK
recipient_dlg_proc (HWND dlg, UINT msg, WPARAM wparam, LPARAM lparam)
{
  static struct recipient_cb_s *rset_cb;
  NMHDR *notify;
  HWND hrset;
  size_t pos;
  int i, j;

  switch (msg) 
    {
    case WM_INITDIALOG:
      rset_cb = (struct recipient_cb_s *)lparam;
      assert (rset_cb != NULL);
      recipient_dlg_set_labels (dlg);
      initialize_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET1));
      rset_cb->keyarray = load_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET1),
                                        &rset_cb->keyarray_count);
      initialize_rsetbox (GetDlgItem (dlg, IDC_ENC_RSET2));

      if (rset_cb->unknown_keys)
        initialize_keybox (dlg, rset_cb);
      else
        {
          /* No unknown keys; thus we do not need the unwanted key box. */
          ShowWindow (GetDlgItem (dlg, IDC_ENC_NOTFOUND_T), SW_HIDE);
          ShowWindow (GetDlgItem (dlg, IDC_ENC_NOTFOUND), SW_HIDE);
	}

      center_window (dlg, NULL);
      SetForegroundWindow (dlg);
      return TRUE;

    case WM_NOTIFY:
      notify = (LPNMHDR)lparam;
      if (notify && notify->code == NM_DBLCLK
          && (notify->idFrom == IDC_ENC_RSET1
              || notify->idFrom == IDC_ENC_RSET2))
        copy_item (dlg, notify->idFrom, -1);
      break;

    case WM_COMMAND:
      switch (LOWORD (wparam))
        {
	case IDOK:
          hrset = GetDlgItem (dlg, IDC_ENC_RSET2);
          if (ListView_GetItemCount (hrset) == 0) 
            {
              MessageBox (dlg, 
                      _("Please select at least one recipient certificate."),
                      _("Recipient Dialog"), MB_ICONINFORMATION|MB_OK);
              return TRUE;
	    }

          for (j=0; rset_cb->fnd_keys && rset_cb->fnd_keys[j]; j++)
            ;
          rset_cb->selected_keys_count = ListView_GetItemCount (hrset);
          rset_cb->selected_keys = xcalloc (rset_cb->selected_keys_count
                                            + j + 1,
                                            sizeof *rset_cb->selected_keys);
          /* Add the selected keys. */
          for (i=0, pos=0; i < rset_cb->selected_keys_count; i++) 
            {
              gpgme_key_t key;
	      int idata;
	      
	      idata = (int)lv_get_item_param (hrset, i);
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
                         disable this "feature". It is however pretty
                         much messed up: The default key should never
                         be processed here but set into the gpg.conf
                         file becuase it is always trusted.  */
                      rset_cb->opts |= OPT_FLAG_FORCE;
                      break;
                    }
                }
              else
                log_debug ("List item not correctly initialized - ignored\n");
            }
          /* Add the already found keys. */
          for (i=0; rset_cb->fnd_keys && rset_cb->fnd_keys[i]; i++)
            {
              gpgme_key_ref (rset_cb->fnd_keys[i]);
              rset_cb->selected_keys[pos++] = rset_cb->fnd_keys[i];
            }

          rset_cb->selected_keys_count = pos;
          EndDialog (dlg, TRUE);
          break;

	case IDCANCEL:
	  /* now that Outlook correctly aborts the delivery, we do not
	     need any warning message if the user cancels thi dialog. */
	  rset_cb->opts = OPT_FLAG_CANCEL;
	  EndDialog (dlg, FALSE);
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
  int resid;
  
  *ret_rset = NULL;

  memset (&cb, 0, sizeof (cb));
  resid = IDD_ENC;
  DialogBoxParam (glob_hinst, (LPCTSTR)resid, GetDesktopWindow(),
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
  int resid;
  
  *ret_rset = NULL;

  memset (&cb, 0, sizeof (cb));
  cb.fnd_keys = fnd;
  cb.unknown_keys = unknown;

  resid = IDD_ENC;
  DialogBoxParam (glob_hinst, (LPCTSTR)resid, GetDesktopWindow (),
		  recipient_dlg_proc, (LPARAM)&cb);

  if (cb.opts & OPT_FLAG_CANCEL)
    release_keyarray (cb.selected_keys, cb.selected_keys_count);
  else
    *ret_rset = cb.selected_keys;

  release_keyarray (cb.keyarray, cb.keyarray_count);
  return cb.opts;
}
