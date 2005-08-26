/* display.cpp - Helper functions to display messages.
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of OutlGPG.
 * 
 * OutlGPG is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * OutlGPG is distributed in the hope that it will be useful,
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


/* Create a new body from body wth suitable line endings. Caller must
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
   the text of the window instead of the MAPI object itself. */
static HWND
find_message_window (HWND parent, GpgMsg *msg)
{
  HWND child;

  if (!parent)
    return NULL;

  child = GetWindow (parent, GW_CHILD);
  while (child)
    {
      char buf[1024+1];
      HWND w;

      memset (buf, 0, sizeof (buf));
      GetWindowText (child, buf, sizeof (buf)-1);
      if (msg->matchesString (buf))
        return child;
      w = find_message_window (child, msg);
      if (w)
        return w;
      child = GetNextWindow (child, GW_HWNDNEXT);	
    }

  return NULL;
}


/* Update the display using the message MSG.  Return 0 on success. */
int
update_display (HWND hwnd, GpgMsg *msg)
{
  HWND window;

  window = find_message_window (hwnd, msg);
  if (window)
    {
      log_debug ("%s:%s: window handle %p\n", __FILE__, __func__, window);
      SetWindowText (window, msg->getDisplayText ());
      log_debug ("%s:%s: window text is now `%s'",
                 __FILE__, __func__, msg->getDisplayText ());
      return 0;
    }
  else
    {
      log_debug ("%s: window handle not found for parent %p\n",
                 __func__, hwnd);
      return -1;
    }
}

