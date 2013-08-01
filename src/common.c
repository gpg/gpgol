/* common.c - Common routines used by GpgOL
 *	Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public License
 * as published by the Free Software Foundation; either version 2.1
 * of the License, or (at your option) any later version.
 *
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

#include <config.h>
#define OEMRESOURCE    /* Required for OBM_CHECKBOXES.  */
#include <windows.h>
#include <shlobj.h>
#ifndef CSIDL_APPDATA
#define CSIDL_APPDATA 0x001a
#endif
#ifndef CSIDL_LOCAL_APPDATA
#define CSIDL_LOCAL_APPDATA 0x001c
#endif
#ifndef CSIDL_FLAG_CREATE
#define CSIDL_FLAG_CREATE 0x8000
#endif
#include <time.h>
#include <fcntl.h>
#include <ctype.h>

#include "common.h"

HINSTANCE glob_hinst = NULL;


/* The base-64 list used for base64 encoding. */
static unsigned char bintoasc[64+1] = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                       "abcdefghijklmnopqrstuvwxyz"
                                       "0123456789+/");

/* The reverse base-64 list used for base-64 decoding. */
static unsigned char const asctobin[256] = {
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x3e, 0xff, 0xff, 0xff, 0x3f,
  0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x3b, 0x3c, 0x3d, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06,
  0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10, 0x11, 0x12,
  0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0x1a, 0x1b, 0x1c, 0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23, 0x24,
  0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b, 0x2c, 0x2d, 0x2e, 0x2f, 0x30,
  0x31, 0x32, 0x33, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
  0xff, 0xff, 0xff, 0xff
};



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


/* Return the system's bitmap of the check bar used which check boxes.
   If CHECKED is set, this check mark is returned; if it is not set,
   the one used for not-checked is returned.  May return NULL on
   error.  Taken from an example in the platform reference. 

   Not used as of now. */
HBITMAP
get_system_check_bitmap (int checked)
{
  COLORREF bg_color;
  HBRUSH bg_brush, saved_dst_brush;
  HDC src_dc, dst_dc;
  WORD xsize, ysize;
  HBITMAP result, saved_dst_bitmap, saved_src_bitmap, checkboxes;
  BITMAP bitmap;
  RECT rect;

  bg_color = GetSysColor (COLOR_MENU);
  bg_brush = CreateSolidBrush (bg_color);

  src_dc = CreateCompatibleDC (NULL);
  dst_dc = CreateCompatibleDC (src_dc);

  xsize = GetSystemMetrics (SM_CXMENUCHECK);
  ysize = GetSystemMetrics (SM_CYMENUCHECK);
  result = CreateCompatibleBitmap(src_dc, xsize, ysize);

  saved_dst_brush  = SelectObject (dst_dc, bg_brush);
  saved_dst_bitmap = SelectObject (dst_dc, result);

  PatBlt (dst_dc, 0, 0, xsize, ysize, PATCOPY);

  checkboxes = LoadBitmap (NULL, (LPTSTR)OBM_CHECKBOXES);

  saved_src_bitmap = SelectObject (src_dc, checkboxes);

  GetObject (checkboxes, sizeof (BITMAP), &bitmap);
  rect.top = 0;
  rect.bottom = (bitmap.bmHeight / 3);
  if (checked)
    {
      /* Select row 1, column 1.  */
      rect.left  = 0;
      rect.right = (bitmap.bmWidth / 4);
    }
  else
    {
      /* Select row 1, column 2. */ 
      rect.left  = (bitmap.bmWidth / 4);
      rect.right = (bitmap.bmWidth / 4) * 2;
    }

  if ( ((rect.right - rect.left) > (int)xsize)
       || ((rect.bottom - rect.top) > (int)ysize) )
    StretchBlt (dst_dc, 0, 0, xsize, ysize, src_dc, rect.left, rect.top,
                rect.right - rect.left, rect.bottom - rect.top, SRCCOPY);
  else
    BitBlt (dst_dc, 0, 0, rect.right - rect.left, rect.bottom - rect.top,
            src_dc, rect.left, rect.top, SRCCOPY);

  SelectObject (src_dc, saved_src_bitmap);
  SelectObject (dst_dc, saved_dst_brush);
  result = SelectObject (dst_dc, saved_dst_bitmap);

  DeleteObject (bg_brush);
  DeleteObject (src_dc);
  DeleteObject (dst_dc);
  return result;
}

/* Return the path to a file that should be worked with.
   Returns a malloced string (UTF-8) on success.
   HWND is the current Window.
   Title is a UTF-8 encoded string containing the
   dialog title and may be NULL.
   On error (i.e. cancel) NULL is returned. */
char *
get_open_filename (HWND root, const char *title)
{
  OPENFILENAMEW ofn;
  wchar_t fname[MAX_PATH+1];
  wchar_t *wTitle = NULL;

  if (title)
    {
      wTitle = utf8_to_wchar2 (title, strlen(title));
    }
  memset (fname, 0, sizeof (fname));

  /* Set up the ofn structure */
  memset (&ofn, 0, sizeof (ofn));
  ofn.lStructSize = sizeof (ofn);
  ofn.hwndOwner = root;
  ofn.lpstrFile = fname;
  ofn.nMaxFile = MAX_PATH;
  ofn.lpstrTitle = wTitle;
  ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

  if (GetOpenFileNameW (&ofn))
    {
      xfree (wTitle);
      return wchar_to_utf8_2 (fname, MAX_PATH);
    }
  xfree (wTitle);
  return NULL;
}


/* Return a filename to be used for saving an attachment. Returns a
   malloced string on success. HWND is the current Window and SRCNAME
   the filename to be used as suggestion.  On error (i.e. cancel) NULL
   is returned. */
char *
get_save_filename (HWND root, const char *srcname)
{
  char filter[21] = "All Files (*.*)\0*.*\0\0";
  char fname[MAX_PATH+1];
  char filterBuf[32];
  char* extSep;
  OPENFILENAME ofn;

  memset (fname, 0, sizeof (fname));
  memset (filterBuf, 0, sizeof (filterBuf));
  strncpy (fname, srcname, MAX_PATH-1);
  fname[MAX_PATH] = 0;

  if ((extSep = strrchr (srcname, '.')) && strlen (extSep) <= 4)
    {
      /* Windows removes the file extension by default so we
         need to set the first filter to the file extension.
      */
      strcpy (filterBuf, extSep);
      strcpy (filterBuf + strlen (filterBuf) + 1, extSep);
      memcpy (filterBuf + strlen (extSep) * 2 + 2, filter, 21);
    }
  else
    memcpy (filterBuf, filter, 21);


  memset (&ofn, 0, sizeof (ofn));
  ofn.lStructSize = sizeof (ofn);
  ofn.hwndOwner = root;
  ofn.lpstrFile = fname;
  ofn.nMaxFile = MAX_PATH;
  ofn.lpstrFileTitle = NULL;
  ofn.nMaxFileTitle = 0;
  ofn.Flags |= OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT;
  ofn.lpstrTitle = _("GpgOL - Save attachment");
  ofn.lpstrFilter = filterBuf;

  if (GetSaveFileName (&ofn))
    return xstrdup (fname);
  return NULL;
}


void
fatal_error (const char *format, ...)
{
  va_list arg_ptr;
  char buf[512];

  va_start (arg_ptr, format);
  vsnprintf (buf, sizeof buf -1, format, arg_ptr);
  buf[sizeof buf - 1] = 0;
  va_end (arg_ptr);
  MessageBox (NULL, buf, "Fatal Error", MB_OK);
  abort ();
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

void *
xrealloc (void *a, size_t n)
{
  void *p = realloc (a, n);
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
static HRESULT
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


/* This function is similar to strncpy().  However it won't copy more
   than N - 1 characters and makes sure that a Nul is appended. With N
   given as 0, nothing will happen.  With DEST given as NULL, memory
   will be allocated using xmalloc (i.e. if it runs out of core the
   function terminates).  Returns DEST or a pointer to the allocated
   memory.  */
char *
mem2str (char *dest, const void *src, size_t n)
{
  char *d;
  const char *s;
  
  if (n)
    {
      if (!dest)
        dest = xmalloc (n);
      d = dest;
      s = src ;
      for (n--; n && *s; n--)
        *d++ = *s++;
      *d = 0;
    }
  else if (!dest)
    {
      dest = xmalloc (1);
      *dest = 0;
    }
  
  return dest;
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


/* Strip off leading and trailing white spaces from STRING.  Returns
   STRING. */
char *
trim_spaces (char *arg_string)
{
  char *string = arg_string;
  char *p, *mark;

  /* Find first non space character. */
  for (p = string; *p && isascii (*p) && isspace (*p) ; p++ )
    ;
  /* Move characters. */
  for (mark = NULL; (*string = *p); string++, p++ )
    {
      if (isascii (*p) && isspace (*p))
        {
          if (!mark)
          mark = string;
        }
      else
        mark = NULL ;
    }
  if (mark)
    *mark = 0;
  
  return arg_string;
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


/* Get the standard home directory.  In general this function should
   not be used as it does not consider a registry value or the
   GNUPGHOME environment variable.  Please use default_homedir(). */
static const char *
standard_homedir (void)
{
  static char *dir;

  if (!dir)
    {
      char path[MAX_PATH];

      /* It might be better to use LOCAL_APPDATA because this is
         defined as "non roaming" and thus more likely to be kept
         locally.  For private keys this is desired.  However, given
         that many users copy private keys anyway forth and back,
         using a system roaming services might be better than to let
         them do it manually.  A security conscious user will anyway
         use the registry entry to have better control.  */
      if (w32_shgetfolderpath (NULL, CSIDL_APPDATA|CSIDL_FLAG_CREATE,
                               NULL, 0, path) >= 0)
        {
          char *tmp = malloc (strlen (path) + 6 + 1);

	  strcpy (tmp, path);
	  strcat (tmp, "\\gnupg");

          dir = tmp;

          /* Try to create the directory if it does not yet exists.  */
          if (access (dir, F_OK))
            CreateDirectory (dir, NULL);
        }
      else
        dir = xstrdup ("C:\\gnupg");
    }
  return dir;
}


/* Retrieve the default home directory.  */
const char *
default_homedir (void)
{
  static char *dir;

  if (!dir)
    {
      dir = getenv ("GNUPGHOME");
      if (!dir || !*dir)
        {
          char *tmp;

          tmp = read_w32_registry_string (NULL, GNUPG_REGKEY, "HomeDir");
          if (tmp && !*tmp)
            {
              free (tmp);
              tmp = NULL;
            }
          if (tmp)
            dir = tmp;
          else
            dir = xstrdup (standard_homedir ());
        }
      else
        dir = xstrdup (dir);
    }

  return dir;
}

/* Return the data dir used for forms etc.   Returns NULL on error. */
char *
get_data_dir (void)
{
  char *instdir;
  char *p;
  char *dname;

  instdir = read_w32_registry_string ("HKEY_LOCAL_MACHINE", GNUPG_REGKEY,
				      "Install Directory");
  if (!instdir)
    return NULL;
  
  /* Build the key: "<instdir>/share/gpgol".  */
#define SDDIR "\\share\\gpgol"
  dname = malloc (strlen (instdir) + strlen (SDDIR) + 1);
  if (!dname)
    {
      free (instdir);
      return NULL;
    }
  p = dname;
  strcpy (p, instdir);
  p += strlen (instdir);
  strcpy (p, SDDIR);
  
  free (instdir);
  
#undef SDDIR
  return dname;
}



/* Do in-place decoding of quoted-printable data of LENGTH in BUFFER.
   Returns the new length of the buffer and stores true at R_SLBRK if
   the line ended with a soft line break; false is stored if not.
   This fucntion asssumes that a complete line is passed in
   buffer.  */
size_t
qp_decode (char *buffer, size_t length, int *r_slbrk)
{
  char *d, *s;

  if (r_slbrk)
    *r_slbrk = 0;

  /* Fixme:  We should remove trailing white space first.  */
  for (s=d=buffer; length; length--)
    if (*s == '=')
      {
        if (length > 2 && hexdigitp (s+1) && hexdigitp (s+2))
          {
            s++;
            *(unsigned char*)d++ = xtoi_2 (s);
            s += 2;
            length -= 2;
          }
        else if (length > 2 && s[1] == '\r' && s[2] == '\n')
          {
            /* Soft line break.  */
            s += 3;
            length -= 2;
            if (r_slbrk && length == 1)
              *r_slbrk = 1;
          }
        else if (length > 1 && s[1] == '\n')
          {
            /* Soft line break with only a Unix line terminator. */
            s += 2;
            length -= 1;
            if (r_slbrk && length == 1)
              *r_slbrk = 1;
          }
        else if (length == 1)
          {
            /* Soft line break at the end of the line. */
            s += 1;
            if (r_slbrk)
              *r_slbrk = 1;
          }
        else
          *d++ = *s++;
      }
    else
      *d++ = *s++;

  return d - buffer;
}


/* Initialize the Base 64 decoder state.  */
void b64_init (b64_state_t *state)
{
  state->idx = 0;
  state->val = 0;
  state->stop_seen = 0;
  state->invalid_encoding = 0;
}


/* Do in-place decoding of base-64 data of LENGTH in BUFFER.  Returns
   the new length of the buffer. STATE is required to return errors and
   to maintain the state of the decoder.  */
size_t
b64_decode (b64_state_t *state, char *buffer, size_t length)
{
  int idx = state->idx;
  unsigned char val = state->val;
  int c;
  char *d, *s;

  if (state->stop_seen)
    return 0;

  for (s=d=buffer; length; length--, s++)
    {
      if (*s == '\n' || *s == ' ' || *s == '\r' || *s == '\t')
        continue;
      if (*s == '=')
        {
          /* Pad character: stop */
          if (idx == 1)
            *d++ = val;
          state->stop_seen = 1;
          break;
        }

      if ((c = asctobin[*(unsigned char *)s]) == 255)
        {
          if (!state->invalid_encoding)
            log_debug ("%s: invalid base64 character %02X at pos %d skipped\n",
                       __func__, *(unsigned char*)s, (int)(s-buffer));
          state->invalid_encoding = 1;
          continue;
        }

      switch (idx)
        {
        case 0:
          val = c << 2;
          break;
        case 1:
          val |= (c>>4)&3;
          *d++ = val;
          val = (c<<4)&0xf0;
          break;
        case 2:
          val |= (c>>2)&15;
          *d++ = val;
          val = (c<<6)&0xc0;
          break;
        case 3:
          val |= c&0x3f;
          *d++ = val;
          break;
        }
      idx = (idx+1) % 4;
    }


  state->idx = idx;
  state->val = val;
  return d - buffer;
}


/* Create a boundary.  Note that mimemaker.c knows about the structure
   of the boundary (i.e. that it starts with "=-=") so that it can
   protect against accidently used boundaries within the content.  */
char *
generate_boundary (char *buffer)
{
  char *p = buffer;
  int i;

#if RAND_MAX < (64*2*BOUNDARYSIZE)
#error RAND_MAX is way too small
#endif

  *p++ = '=';
  *p++ = '-';
  *p++ = '=';
  for (i=0; i < BOUNDARYSIZE-6; i++)
    *p++ = bintoasc[rand () % 64];
  *p++ = '=';
  *p++ = '-';
  *p++ = '=';
  *p = 0;

  return buffer;
}


/* Fork and exec the program given in CMDLINE with /dev/null as
   stdin, stdout and stderr.  Returns 0 on success.  */
int
gpgol_spawn_detached (const char *cmdline)
{
  int rc;
  SECURITY_ATTRIBUTES sec_attr;
  PROCESS_INFORMATION pi = { NULL, 0, 0, 0 };
  STARTUPINFO si;
  int cr_flags;
  char *cmdline_copy;

  memset (&sec_attr, 0, sizeof sec_attr);
  sec_attr.nLength = sizeof sec_attr;
  
  memset (&si, 0, sizeof si);
  si.cb = sizeof (si);
  si.dwFlags = STARTF_USESHOWWINDOW;
  si.wShowWindow = SW_SHOW;

  cr_flags = (CREATE_DEFAULT_ERROR_MODE
              | GetPriorityClass (GetCurrentProcess ())
	      | CREATE_NEW_PROCESS_GROUP
              | DETACHED_PROCESS); 

  cmdline_copy = xstrdup (cmdline);
  rc = CreateProcess (NULL,          /* No appliactionname, use CMDLINE.  */
                      cmdline_copy,  /* Command line arguments.  */
                      &sec_attr,     /* Process security attributes.  */
                      &sec_attr,     /* Thread security attributes.  */
                      FALSE,          /* Inherit handles.  */
                      cr_flags,      /* Creation flags.  */
                      NULL,          /* Environment.  */
                      NULL,          /* Use current drive/directory.  */
                      &si,           /* Startup information. */
                      &pi            /* Returns process information.  */
                      );
  xfree (cmdline_copy);
  if (!rc)
    {
      log_error_w32 (-1, "%s:%s: CreateProcess failed", SRCNAME, __func__);
      return -1;
    }

  CloseHandle (pi.hThread); 
  CloseHandle (pi.hProcess);
  return 0;
}



/* Percent-escape the string STR by replacing colons with '%3a'.  If
   EXTRA is not NULL all characters in it are also escaped. */
char *
percent_escape (const char *str, const char *extra)
{
  int i, j;
  char *ptr;

  if (!str)
    return NULL;

  for (i=j=0; str[i]; i++)
    if (str[i] == ':' || str[i] == '%' || (extra && strchr (extra, str[i])))
      j++;
  ptr = (char *) malloc (i + 2 * j + 1);
  i = 0;
  while (*str)
    {
      /* FIXME: Work around a bug in Kleo.  */
      if (*str == ':')
        {
          ptr[i++] = '%';
          ptr[i++] = '3';
          ptr[i++] = 'a';
        }
      else
        {
          if (*str == '%')
            {
              ptr[i++] = '%';
              ptr[i++] = '2';
              ptr[i++] = '5';
            }
          else if (extra && strchr (extra, *str))
            {
              ptr[i++] = '%';
              ptr[i++] = tohex_lower ((*str >> 4) & 15);
              ptr[i++] = tohex_lower (*str & 15);
            }
          else
            ptr[i++] = *str;
        }
      str++;
    }
  ptr[i] = '\0';

  return ptr;
}

/* Fix linebreaks.
   This replaces all consecutive \r or \n characters
   by a single \n.
   There can be extremly weird combinations of linebreaks
   like \r\r\n\r\r\n at the end of each line when
   getting the body of a mail message.
*/
void
fix_linebreaks (char *str, int *len)
{
  char *src;
  char *dst;

  src = str;
  dst = str;
  while (*src)
    {
      if (*src == '\r' || *src == '\n')
        {
          do
            src++;
          while (*src == '\r' || *src == '\n');
          *(dst++) = '\n';
        }
      else
        {
          *(dst++) = *(src++);
        }
    }
  *dst = '\0';
  *len = dst - str;
}
