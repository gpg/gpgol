/* display.cpp - Helper functions to display messages.
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

#include <time.h>
#include <assert.h>
#include <string.h>
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "mapihelp.h"
#include "ol-ext-callback.h"
#include "display.h"


/* Check wether the string BODY is HTML formatted. */
int 
is_html_body (const char *body)
{
  char *p1, *p2;
  
  /* XXX: it is possible but unlikely that the message text
     contains the used keywords. */
  p1 = strstr (body, "<HTML>");
  p2 = strstr (body, "</HTML>");
  if (p1 && p2)
    return 1;
  p1 = strstr (body, "<html>");
  p2 = strstr (body, "</html>");
  if (p1 && p2)
	return 1;
  /* XXX: use case insentensive strstr version. */
  return 0;
}


/* Create a new body from body with suitable line endings. Caller must
   release the result. */
char *
add_html_line_endings (const char *body)
{
  size_t count;
  const char *s;
  char *p, *result;

  for (count=0, s = body; *s; s++)
    if (*s == '\n')
      count++;
  
  result = (char*)xmalloc ((s - body) + count*10 + 1);
  
  for (s=body, p = result; *s; s++ )
    if (*s == '\n')
      p = stpcpy (p, "&nbsp;<br>\n");
    else
      *p++ = *s;
  *p = 0;
  
  return result;
  
}


/* We need this to find the mailer window because we directly change
   the text of the window instead of the MAPI object itself.  To do
   this we walk all windows to find a PGP signature.  */
static HWND
find_message_window (HWND parent, int level)
{
  HWND child;

  if (!parent)
    return NULL;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      char buf[1024+1];
      HWND w;
      size_t len;
      const char *s;

      /* OL 2003 SP1 German uses this class name for the main
         inspector window.  We hope that no other windows uses this
         class name.  As a fallback we keep on testing for PGP
         strings, but this does not work for PGP/MIME or already
         decrypted messages. */
      len = GetClassName (child, buf, sizeof buf - 1);
//       if (len)
//         log_debug ("  %*sgot class `%s'", level*2, "", buf);
      if (level && len >= 10 && !strncmp (buf, "MsoCommand", 10))
        {
          /* We won't find anything below MsoCommand windows.
             Ignoring them fixes a bug where we return a RichEdit20W
             window which is actually a formatting drop down box or
             something similar.  Not sure whether the check for level
             is required, but it won't harm and might help in case an
             MsoCommand* is the top level.
             
             An example of such a message hierarchy is:
               got class `MsoCommandBarDock'
               got class `MsoCommandBarDock'
               got class `MsoCommandBarDock'
                 got class `MsoCommandBar'
                 got class `MsoCommandBar'
                   got class `RichEdit20W'  <--- We don't want that
                 got class `MsoCommandBar'
               got class `MsoCommandBarDock'
               got class `AfxWndW'
                 got class `#32770'
                   got class `Static'
                   got class `Static'
                   got class `RichEdit20WPT'
                   got class `Static'
                   got class `RichEdit20WPT'
                   got class `Static'
                   got class `RichEdit20WPT'
                   got class `Static'
                   got class `RichEdit20WPT'
                   got class `Static'
                   got class `RichEdit20WPT'
                   got class `Static'
                   got class `Static'
                   got class `AfxWndA'
                     got class `Static'
                     got class `AfxWndW'
                       got class `Static'
                       got class `RichEdit20W'  <--- We want this one
           */
          break; /* Not found at this level.  */
        }

      if (len && !strcmp (buf, "RichEdit20W"))
        {
          log_debug ("found class `%s'", "RichEdit20W");
          return child;
        }
      
      memset (buf, 0, sizeof (buf));
      GetWindowText (child, buf, sizeof (buf)-1);
      len = strlen (buf);
      if (len > 22
          && (s = strstr (buf, "-----BEGIN PGP "))
          &&  (!strncmp (s+15, "MESSAGE-----", 12)
               || !strncmp (s+15, "SIGNED MESSAGE-----", 19)))
        return child;
      w = find_message_window (child, level+1);
      if (w)
        return w;
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Update the display with TEXT using the message MSG.  Return 0 on
   success. */
int
update_display (HWND hwnd, void *exchange_cb, 
                bool is_html, const char *text)
{
  HWND window;

  /*show_window_hierarchy (hwnd, 0);*/
  window = find_message_window (hwnd, 0);
  if (window && !is_html)
    {
      const char *s;

      log_debug ("%s:%s: window handle %p\n", SRCNAME, __func__, window);
      
      /* Decide whether we need to use the Unicode version. */
      for (s=text; *s && !(*s & 0x80); s++)
        ;
      if (*s)
        {
          wchar_t *tmp = utf8_to_wchar (text);
          SetWindowTextW (window, tmp);
          xfree (tmp);
        }
      else
        SetWindowTextA (window, text);
      log_debug ("%s:%s: window text is now `%s'",
                 SRCNAME, __func__, text);
      return 0;
    }
  else if (exchange_cb && !opt.compat.no_oom_write)
    {
      log_debug ("%s:%s: updating display using OOM\n", SRCNAME, __func__);
      /* Bug in OL 2002 and 2003 - as a workaround set the body first
         to empty. */
      if (is_html)
        put_outlook_property (exchange_cb, "Body", "" );
      return put_outlook_property (exchange_cb, is_html? "HTMLBody":"Body",
                                   text);
    }
  else
    {
      log_debug ("%s:%s: window handle not found for parent %p\n",
                 SRCNAME, __func__, hwnd);
      return -1;
    }
}


/* Set the body of MESSAGE to STRING.  Returns 0 on success or an
   error code otherwise. */
int
set_message_body (LPMESSAGE message, const char *string, bool is_html)
{
  HRESULT hr;
  SPropValue prop;
  //SPropTagArray proparray;
  const char *s;
  
  assert (message);

//   if (!is_html)
//     {
//       prop.ulPropTag = PR_BODY_HTML_A;
//       prop.Value.lpszA = "";
//       hr = HrSetOneProp (message, &prop);
//     }
  
  /* Decide whether we need to use the Unicode version. */
  for (s=string; *s && !(*s & 0x80); s++)
    ;
  if (*s)
    {
      prop.ulPropTag = is_html? PR_BODY_HTML_W : PR_BODY_W;
      prop.Value.lpszW = utf8_to_wchar (string);
      hr = HrSetOneProp (message, &prop);
      xfree (prop.Value.lpszW);
    }
  else /* Only plain ASCII. */
    {
      prop.ulPropTag = is_html? PR_BODY_HTML_A : PR_BODY_A;
      prop.Value.lpszA = (CHAR*)string;
      hr = HrSetOneProp (message, &prop);
    }
  if (hr != S_OK)
    {
      log_debug ("%s:%s: HrSetOneProp failed: hr=%#lx\n",
                 SRCNAME, __func__, hr); 
      return gpg_error (GPG_ERR_GENERAL);
    }

  /* Note: we once tried to delete the RTF property here to avoid any
     syncing mess and more important to make sure that no RTF rendered
     plaintext is left over.  The side effect of this was that the
     entire PR_BODY go deleted too. */

  return 0;
}



int
open_inspector (LPEXCHEXTCALLBACK peecb, LPMESSAGE message)
{
  HRESULT hr;
  LPMAPISESSION session;
  ULONG token; 
  char *entryid, *store_entryid, *parent_entryid;
  size_t entryidlen, store_entryidlen, parent_entryidlen;
  LPMDB mdb;
  LPMAPIFOLDER mfolder;
  ULONG mtype;
  LPMESSAGE message2;
  
  hr = peecb->GetSession (&session, NULL);
  if (FAILED (hr) )
    {
      
      log_error ("%s:%s: error getting session: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      return -1;
      
    }
  
  entryid = mapi_get_binary_prop (message, PR_ENTRYID, &entryidlen);
  if (!entryid)
    {
      log_error ("%s:%s: PR_ENTRYID missing\n",  SRCNAME, __func__);
      session->Release ();
      return -1;
    }
  log_hexdump (entryid, entryidlen, "orig entryid=");
  store_entryid = mapi_get_binary_prop (message, PR_STORE_ENTRYID,
                                        &store_entryidlen);
  if (!store_entryid)
    {
      log_error ("%s:%s: PR_STORE_ENTRYID missing\n",  SRCNAME, __func__);
      session->Release ();
      xfree (entryid);
      return -1;
    }
  parent_entryid = mapi_get_binary_prop (message, PR_PARENT_ENTRYID,
                                         &parent_entryidlen);
  if (!parent_entryid)
    {
      log_error ("%s:%s: PR_PARENT_ENTRYID missing\n",  SRCNAME, __func__);
      session->Release ();
      xfree (store_entryid);
      xfree (entryid);
      return -1;
    }

  /* Open the message store by ourself.  */
  hr = session->OpenMsgStore (0, store_entryidlen, (LPENTRYID)store_entryid,
                              NULL,  MAPI_BEST_ACCESS | MDB_NO_DIALOG, 
                              &mdb);
  if (FAILED (hr))
    {
      log_error ("%s:%s: OpenMsgStore failed: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      session->Release ();
      xfree (parent_entryid);
      xfree (store_entryid);
      xfree (entryid);
      return -1;
    }

  /* Open the parent folder.  */
  hr = mdb->OpenEntry (parent_entryidlen,  (LPENTRYID)parent_entryid,
                       &IID_IMAPIFolder, MAPI_BEST_ACCESS, 
                       &mtype, (IUnknown**)&mfolder);
  if (FAILED (hr))
    {
      log_error ("%s:%s: OpenEntry failed: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      session->Release ();
      xfree (parent_entryid);
      xfree (store_entryid);
      xfree (entryid);
      mdb->Release ();
      return -1;
    }
  log_debug ("%s:%s: mdb::OpenEntry succeeded type=%lx\n", 
             SRCNAME, __func__, mtype);

  /* Open the message.  */
  hr = mdb->OpenEntry (entryidlen,  (LPENTRYID)entryid,
                       &IID_IMessage, MAPI_BEST_ACCESS, 
                       &mtype, (IUnknown**)&message2);
  if (FAILED (hr))
    {
      log_error ("%s:%s: OpenEntry[folder] failed: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      session->Release ();
      xfree (parent_entryid);
      xfree (store_entryid);
      xfree (entryid);
      mdb->Release ();
      mfolder->Release ();
      return -1;
    }
  log_debug ("%s:%s: mdb::OpenEntry[message] succeeded type=%lx\n", 
             SRCNAME, __func__, mtype);

  /* Prepare and display the form.  */
  hr = session->PrepareForm (NULL, message2, &token);
  if (FAILED (hr))
    {
      log_error ("%s:%s: PrepareForm failed: hr=%#lx\n", 
                 SRCNAME, __func__, hr);
      session->Release ();
      xfree (parent_entryid);
      xfree (store_entryid);
      xfree (entryid);
      mdb->Release ();
      mfolder->Release ();
      message2->Release ();
      return -1;
    }

  /* Message2 is now represented by TOKEN; we need to release it.  */
  message2->Release(); message2 = NULL;

  hr = session->ShowForm (0,
                          mdb, mfolder, 
                          NULL,  token,
                          NULL,  
                          0,
                          0,  0,
                          0,  "IPM.Note");
  log_debug ("%s:%s: ShowForm result: hr=%#lx\n", 
             SRCNAME, __func__, hr);
  
  session->Release ();
  xfree (parent_entryid);
  xfree (store_entryid);
  xfree (entryid);
  mdb->Release ();
  mfolder->Release ();
  return FAILED(hr)? -1:0;
}

