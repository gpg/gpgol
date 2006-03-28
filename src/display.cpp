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
#include "intern.h"
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
find_message_window (HWND parent)
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
      if (len && !strcmp (buf, "RichEdit20W"))
        {
          log_debug ("found class RichEdit20W");
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
      w = find_message_window (child);
      if (w)
        return w;
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Update the display with TEXT using the message MSG.  Return 0 on
   success. */
int
update_display (HWND hwnd, GpgMsg *msg, void *exchange_cb, 
                bool is_html, const char *text)
{
  HWND window;

  window = find_message_window (hwnd);
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
      log_debug ("updating display using OOM\n");
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
