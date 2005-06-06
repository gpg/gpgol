/* common.c 
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
#include <windows.h>
#include <time.h>

#include "gpgme.h"
#include "intern.h"

HINSTANCE glob_hinst = NULL;

void
set_global_hinstance (HINSTANCE hinst)
{
    glob_hinst = hinst;
}

/* Center the given window with the desktop window as the
   parent window. */
void
center_window (HWND childwnd, HWND style) 
{     
    HWND parwnd;
    RECT rchild, rparent;    
    HDC hdc;
    int wchild, hchild, wparent, hparent;
    int wscreen, hscreen, xnew, ynew;
    int flags = SWP_NOSIZE | SWP_NOZORDER;

    parwnd = GetDesktopWindow ();
    GetWindowRect (childwnd, &rchild);     
    wchild = rchild.right - rchild.left;     
    hchild = rchild.bottom - rchild.top;

    GetWindowRect (parwnd, &rparent);     
    wparent = rparent.right - rparent.left;     
    hparent = rparent.bottom - rparent.top;      
    
    hdc = GetDC (childwnd);     
    wscreen = GetDeviceCaps (hdc, HORZRES);     
    hscreen = GetDeviceCaps (hdc, VERTRES);     
    ReleaseDC (childwnd, hdc);      
    xnew = rparent.left + ((wparent - wchild) /2);     
    if (xnew < 0)
	xnew = 0;
    else if ((xnew+wchild) > wscreen) 
	xnew = wscreen - wchild;
    ynew = rparent.top  + ((hparent - hchild) /2);
    if (ynew < 0)
	ynew = 0;
    else if ((ynew+hchild) > hscreen)
	ynew = hscreen - hchild;
    if (style == HWND_TOPMOST || style == HWND_NOTOPMOST)
	flags = SWP_NOMOVE | SWP_NOSIZE;
    SetWindowPos (childwnd, style? style : NULL, xnew, ynew, 0, 0, flags);
}


static void
out_of_core (void)
{
    MessageBox (NULL, "Out of core!", "Fatal Error", MB_OK);
}

void*
xmalloc (size_t n)
{
    void *p = malloc (n);
    if (!p)
	out_of_core ();
    return p;
}

void*
xcalloc (size_t m, size_t n)
{
    void *p = calloc (m, n);
    if (!p)
	out_of_core ();
    return p;
}

char*
xstrdup (const char *s)
{
    char *p = xmalloc (strlen (s)+1);
    strcpy (p, s);
    return p;
}

void
xfree (void *p)
{
    if (p)
	free (p);
}


cache_item_t
cache_item_new (void)
{
    return xcalloc (1, sizeof (struct cache_item_s));
}

void
cache_item_free (cache_item_t itm)
{
    if (itm == NULL)
	return;
    xfree (itm->pass);
    xfree (itm);
}
