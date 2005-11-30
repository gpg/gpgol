/* common.c 
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 *
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1 
 * of the License, or (at your option) any later version.
 *  
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with GPGol; if not, write to the Free Software Foundation, 
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
    xnew = rparent.left + ((wparent - wchild) / 2);     
    if (xnew < 0)
	xnew = 0;
    else if ((xnew+wchild) > wscreen) 
	xnew = wscreen - wchild;
    ynew = rparent.top  + ((hparent - hchild) / 2);
    if (ynew < 0)
	ynew = 0;
    else if ((ynew+hchild) > hscreen)
	ynew = hscreen - hchild;
    if (style == HWND_TOPMOST || style == HWND_NOTOPMOST)
	flags = SWP_NOMOVE | SWP_NOSIZE;
    SetWindowPos (childwnd, style? style : NULL, xnew, ynew, 0, 0, flags);
}



/* Return a filename to be used for saving an attachment. Returns a
   malloced string on success. HWND is the current Window and SRCNAME
   the filename to be used as suggestion.  On error (i.e. cancel) NULL
   is returned. */
char *
get_save_filename (HWND root, const char *srcname)
				     
{
  char filter[] = "All Files (*.*)\0*.*\0\0";
  char fname[MAX_PATH+1];
  OPENFILENAME ofn;

  memset (fname, 0, sizeof (fname));
  strncpy (fname, srcname, MAX_PATH-1);
  fname[MAX_PATH] = 0;  
  

  memset (&ofn, 0, sizeof (ofn));
  ofn.lStructSize = sizeof (ofn);
  ofn.hwndOwner = root;
  ofn.lpstrFile = fname;
  ofn.nMaxFile = MAX_PATH;
  ofn.lpstrFileTitle = NULL;
  ofn.nMaxFileTitle = 0;
  ofn.Flags |= OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
  ofn.lpstrTitle = "GPG - Save decrypted attachment";
  ofn.lpstrFilter = filter;

  if (GetSaveFileName (&ofn))
    return xstrdup (fname);
  return NULL;
}


void
out_of_core (void)
{
    MessageBox (NULL, "Out of core!", "Fatal Error", MB_OK);
    abort ();
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


/* This is a helper function to load a Windows function from either of
   one DLLs. */
HRESULT
w32_shgetfolderpath (HWND a, int b, HANDLE c, DWORD d, LPSTR e)
{
  static int initialized;
  static HRESULT (WINAPI * func)(HWND,int,HANDLE,DWORD,LPSTR);

  if (!initialized)
    {
      static char *dllnames[] = { "shell32.dll", "shfolder.dll", NULL };
      void *handle;
      int i;

      initialized = 1;

      for (i=0, handle = NULL; !handle && dllnames[i]; i++)
        {
          handle = LoadLibrary (dllnames[i]);
          if (handle)
            {
              func = (HRESULT (WINAPI *)(HWND,int,HANDLE,DWORD,LPSTR))
                     GetProcAddress (handle, "SHGetFolderPathA");
              if (!func)
                {
                  FreeLibrary (handle);
                  handle = NULL;
                }
            }
        }
    }

  if (func)
    return func (a,b,c,d,e);
  else
    return -1;
}


/* Return a malloced string encoded in UTF-8 from the wide char input
   string STRING.  Caller must xfree this value. On failure returns
   NULL; caller may use GetLastError to get the actual error number.
   The result of calling this function with STRING set to NULL is not
   defined. */
char *
wchar_to_utf8 (const wchar_t *string)
{
  int n;
  char *result;

  /* Note, that CP_UTF8 is not defined in Windows versions earlier
     than NT.*/
  n = WideCharToMultiByte (CP_UTF8, 0, string, -1, NULL, 0, NULL, NULL);
  if (n < 0)
    return NULL;

  result = xmalloc (n+1);
  n = WideCharToMultiByte (CP_UTF8, 0, string, -1, result, n, NULL, NULL);
  if (n < 0)
    {
      xfree (result);
      return NULL;
    }
  return result;
}


/* Same as above, but only convert the first LEN wchars.  */
char *
wchar_to_utf8_2 (const wchar_t *string, size_t len)
{
  int n;
  char *result;

  /* Note, that CP_UTF8 is not defined in Windows versions earlier
     than NT.*/
  n = WideCharToMultiByte (CP_UTF8, 0, string, len, NULL, 0, NULL, NULL);
  if (n < 0)
    return NULL;

  result = xmalloc (n+1);
  n = WideCharToMultiByte (CP_UTF8, 0, string, len, result, n, NULL, NULL);
  if (n < 0)
    {
      xfree (result);
      return NULL;
    }
  return result;
}

/* Return a malloced wide char string from an UTF-8 encoded input
   string STRING.  Caller must xfree this value. On failure returns
   NULL; caller may use GetLastError to get the actual error number.
   The result of calling this function with STRING set to NULL is not
   defined. */
wchar_t *
utf8_to_wchar (const char *string)
{
  int n;
  wchar_t *result;

  n = MultiByteToWideChar (CP_UTF8, 0, string, -1, NULL, 0);
  if (n < 0)
    return NULL;

  result = xmalloc ((n+1) * sizeof *result);
  n = MultiByteToWideChar (CP_UTF8, 0, string, -1, result, n);
  if (n < 0)
    {
      xfree (result);
      return NULL;
    }
  return result;
}


/* Same as above but convert only the first LEN characters.  STRING
   must be at least LEN characters long. */
wchar_t *
utf8_to_wchar2 (const char *string, size_t len)
{
  int n;
  wchar_t *result;

  n = MultiByteToWideChar (CP_UTF8, 0, string, len, NULL, 0);
  if (n < 0)
    return NULL;

  result = xmalloc ((n+1) * sizeof *result);
  n = MultiByteToWideChar (CP_UTF8, 0, string, len, result, n);
  if (n < 0)
    {
      xfree (result);
      return NULL;
    }
  result[n] = 0;
  return result;
}


/* Assume STRING is a Latin-1 encoded and convert it to utf-8.
   Returns a newly malloced UTF-8 string. */
char *
latin1_to_utf8 (const char *string)
{
  const char *s;
  char *buffer, *p;
  size_t n;

  for (s=string, n=0; *s; s++) 
    {
      n++;
      if (*s & 0x80)
        n++;
    }
  buffer = xmalloc (n + 1);
  for (s=string, p=buffer; *s; s++)
    {
      if (*s & 0x80)
        {
          *p++ = 0xc0 | ((*s >> 6) & 3);
          *p++ = 0x80 | (*s & 0x3f);
        }
      else
        *p++ = *s;
    }
  *p = 0;
  return buffer;
}


/* Strip off trailing white spaces from STRING.  Returns STRING. */
char *
trim_trailing_spaces (char *string)
{
  char *p, *mark;

  for (mark=NULL, p=string; *p; p++)
    {
      if (strchr (" \t\r\n", *p ))
        {
          if (!mark)
            mark = p;
	}
	else
          mark = NULL;
    }

  if (mark)
    *mark = 0;
  return string;
}


/* Helper for read_w32_registry_string(). */
static HKEY
get_root_key(const char *root)
{
  HKEY root_key;

  if( !root )
    root_key = HKEY_CURRENT_USER;
  else if( !strcmp( root, "HKEY_CLASSES_ROOT" ) )
    root_key = HKEY_CLASSES_ROOT;
  else if( !strcmp( root, "HKEY_CURRENT_USER" ) )
    root_key = HKEY_CURRENT_USER;
  else if( !strcmp( root, "HKEY_LOCAL_MACHINE" ) )
    root_key = HKEY_LOCAL_MACHINE;
  else if( !strcmp( root, "HKEY_USERS" ) )
    root_key = HKEY_USERS;
  else if( !strcmp( root, "HKEY_PERFORMANCE_DATA" ) )
    root_key = HKEY_PERFORMANCE_DATA;
  else if( !strcmp( root, "HKEY_CURRENT_CONFIG" ) )
    root_key = HKEY_CURRENT_CONFIG;
  else
    return NULL;
  return root_key;
}

/* Return a string from the Win32 Registry or NULL in case of error.
   Caller must release the return value.  A NULL for root is an alias
   for HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE in turn.  NOTE: The value
   is allocated with a plain malloc() - use free() and not the usual
   xfree(). */
char *
read_w32_registry_string (const char *root, const char *dir, const char *name)
{
  HKEY root_key, key_handle;
  DWORD n1, nbytes, type;
  char *result = NULL;

  if ( !(root_key = get_root_key(root) ) )
    return NULL;

  if( RegOpenKeyEx( root_key, dir, 0, KEY_READ, &key_handle ) )
    {
      if (root)
	return NULL; /* no need for a RegClose, so return direct */
      /* It seems to be common practise to fall back to HKLM. */
      if (RegOpenKeyEx (HKEY_LOCAL_MACHINE, dir, 0, KEY_READ, &key_handle) )
	return NULL; /* still no need for a RegClose, so return direct */
    }

  nbytes = 1;
  if( RegQueryValueEx( key_handle, name, 0, NULL, NULL, &nbytes ) ) {
    if (root)
      goto leave;
    /* Try to fallback to HKLM also vor a missing value.  */
    RegCloseKey (key_handle);
    if (RegOpenKeyEx (HKEY_LOCAL_MACHINE, dir, 0, KEY_READ, &key_handle) )
      return NULL; /* Nope.  */
    if (RegQueryValueEx( key_handle, name, 0, NULL, NULL, &nbytes))
      goto leave;
  }
  result = malloc( (n1=nbytes+1) );
  if( !result )
    goto leave;
  if( RegQueryValueEx( key_handle, name, 0, &type, result, &n1 ) ) {
    free(result); result = NULL;
    goto leave;
  }
  result[nbytes] = 0; /* make sure it is really a string  */
  if (type == REG_EXPAND_SZ && strchr (result, '%')) {
    char *tmp;

    n1 += 1000;
    tmp = malloc (n1+1);
    if (!tmp)
      goto leave;
    nbytes = ExpandEnvironmentStrings (result, tmp, n1);
    if (nbytes && nbytes > n1) {
      free (tmp);
      n1 = nbytes;
      tmp = malloc (n1 + 1);
      if (!tmp)
	goto leave;
      nbytes = ExpandEnvironmentStrings (result, tmp, n1);
      if (nbytes && nbytes > n1) {
	free (tmp); /* oops - truncated, better don't expand at all */
	goto leave;
      }
      tmp[nbytes] = 0;
      free (result);
      result = tmp;
    }
    else if (nbytes) { /* okay, reduce the length */
      tmp[nbytes] = 0;
      free (result);
      result = malloc (strlen (tmp)+1);
      if (!result)
	result = tmp;
      else {
	strcpy (result, tmp);
	free (tmp);
      }
    }
    else {  /* error - don't expand */
      free (tmp);
    }
  }

 leave:
  RegCloseKey( key_handle );
  return result;
}

