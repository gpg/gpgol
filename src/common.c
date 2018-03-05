/* common.c - Common routines used by GpgOL
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 * 2015, 2016, 2017  Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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

#include "dialogs.h"

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
bring_to_front (HWND wid)
{
  if (wid)
    {
      if (!SetForegroundWindow (wid))
        {
          log_debug ("%s:%s: SetForegroundWindow failed", SRCNAME, __func__);
          /* Yet another fallback which will not work on some
           * versions and is not recommended by msdn */
          if (!ShowWindow (wid, SW_SHOWNORMAL))
            {
              log_debug ("%s:%s: ShowWindow failed.", SRCNAME, __func__);
            }
        }
    }
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

          tmp = read_w32_registry_string (NULL, GPG4WIN_REGKEY_3, "HomeDir");
          if (!tmp)
            {
              tmp = read_w32_registry_string (NULL, GPG4WIN_REGKEY_2, "HomeDir");
            }
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

  instdir = get_gpg4win_dir();
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

/* Get a pretty name for the file at path path. File extension
   will be set to work for the protocol as provided in protocol and
   depends on the signature setting. Set signature to 0 if the
   extension should not be a signature extension.
   Returns NULL on success.
   Caller must free result. */
wchar_t *
get_pretty_attachment_name (wchar_t *path, protocol_t protocol,
                            int signature)
{
  wchar_t* pretty;
  wchar_t* buf;

  if (!path || !wcslen (path))
    {
      log_error("%s:%s: No path given", SRCNAME, __func__);
      return NULL;
    }

  pretty = (wchar_t*) xmalloc ((MAX_PATH + 1) * sizeof (wchar_t));
  memset (pretty, 0, (MAX_PATH + 1) * sizeof (wchar_t));

  buf = wcsrchr (path, '\\') + 1;

  if (!buf || !*buf)
    {
      log_error("%s:%s: No filename found in path", SRCNAME, __func__);
      xfree (pretty);
      return NULL;
    }

  wcscpy (pretty, buf);

  buf = pretty + wcslen(pretty);
  if (signature)
    {
      if (protocol == PROTOCOL_SMIME)
        {
          *(buf++) = '.';
          *(buf++) = 'p';
          *(buf++) = '7';
          *(buf++) = 's';
        }
      else
        {
          *(buf++) = '.';
          *(buf++) = 's';
          *(buf++) = 'i';
          *(buf++) = 'g';
        }
    }
  else
    {
      if (protocol == PROTOCOL_SMIME)
        {
          *(buf++) = '.';
          *(buf++) = 'p';
          *(buf++) = '7';
          *(buf++) = 'm';
        }
      else
        {
          *(buf++) = '.';
          *(buf++) = 'g';
          *(buf++) = 'p';
          *(buf++) = 'g';
        }
    }

  return pretty;
}

/* Open a file in a temporary directory, take name as a
   suggestion and put the open Handle in outHandle.
   Returns the actually used file name in case there
   were other files with that name. */
wchar_t*
get_tmp_outfile (wchar_t *name, HANDLE *outHandle)
{
  wchar_t tmpPath[MAX_PATH];
  wchar_t *outName;
  wchar_t *fileExt = NULL;
  int tries = 1;

  if (!name || !wcslen(name))
    {
      log_error ("%s:%s: Needs a name.",
                 SRCNAME, __func__);
      return NULL;
    }

  /* We should probably use the unicode variants here
     but this would mean adding OpenStreamOnFileW to
     out mapi */

  if (!GetTempPathW (MAX_PATH, tmpPath))
    {
      log_error ("%s:%s: Could not get tmp path.",
                 SRCNAME, __func__);
      return NULL;
    }

  outName = (wchar_t*) xmalloc ((MAX_PATH + 1) * sizeof(wchar_t));
  memset (outName, 0, (MAX_PATH + 1) * sizeof (wchar_t));

  snwprintf (outName, MAX_PATH, L"%s%s", tmpPath, name);

  while ((*outHandle = CreateFileW (outName,
                                    GENERIC_WRITE | GENERIC_READ,
                                    FILE_SHARE_READ | FILE_SHARE_DELETE,
                                    NULL,
                                    CREATE_NEW,
                                    FILE_ATTRIBUTE_TEMPORARY,
                                    NULL)) == INVALID_HANDLE_VALUE)
    {
      wchar_t fnameBuf[MAX_PATH + 1];
      wchar_t origName[MAX_PATH + 1];
      memset (fnameBuf, 0, MAX_PATH + 1);
      memset (origName, 0, MAX_PATH + 1);

      snwprintf (origName, MAX_PATH, L"%s%s", tmpPath, name);
      fileExt = wcschr (wcsrchr(origName, '\\'), '.');
      if (fileExt)
        {
          wcsncpy (fnameBuf, origName, fileExt - origName);
        }
      else
        {
          wcsncpy (fnameBuf, origName, wcslen (origName));
        }
      snwprintf (outName, MAX_PATH, L"%s%i%s", fnameBuf, tries++,
                 fileExt ? fileExt : L"");
      if (tries > 100)
        {
          /* You have to know when to give up,.. */
          log_error ("%s:%s: Could not get a name out of 100 tries",
                     SRCNAME, __func__);
          xfree (outName);
          return NULL;
        }
    }

  return outName;
}

/** Get the Gpg4win Install directory.
 *
 * Looks first for the Gpg4win 3.x registry key. Then for the Gpg4win
 * 2.x registry key. And checks that the directory can be read.
 *
 * @returns NULL if no dir could be found. Otherwise a malloced string.
 */
char *
get_gpg4win_dir()
{
  const char *g4win_keys[] = {GPG4WIN_REGKEY_3,
                              GPG4WIN_REGKEY_2,
                              NULL};
  const char **key;
  for (key = g4win_keys; *key; key++)
    {
      char *tmp = read_w32_registry_string (NULL, *key, "Install Directory");
      if (!tmp)
        {
          continue;
        }
      if (!access(tmp, R_OK))
        {
          return tmp;
        }
      else
        {
          log_debug ("Failed to access: %s\n", tmp);
          xfree (tmp);
        }
    }
  return NULL;
}


static void
epoch_to_file_time (unsigned long time, LPFILETIME pft)
{
 LONGLONG ll;

 ll = Int32x32To64(time, 10000000) + 116444736000000000;
 pft->dwLowDateTime = (DWORD)ll;
 pft->dwHighDateTime = ll >> 32;
}

char *
format_date_from_gpgme (unsigned long time)
{
  wchar_t buf[256];
  FILETIME ft;
  SYSTEMTIME st;

  epoch_to_file_time (time, &ft);
  FileTimeToSystemTime(&ft, &st);
  int ret = GetDateFormatEx (NULL,
                             DATE_SHORTDATE,
                             &st,
                             NULL,
                             buf,
                             256,
                             NULL);
  if (ret == 0)
    {
      return NULL;
    }
  return wchar_to_utf8 (buf);
}

/* Return the name of the default UI server.  This name is used to
   auto start an UI server if an initial connect failed.  */
char *
get_uiserver_name (void)
{
  char *name = NULL;
  char *dir, *uiserver, *p;
  int extra_arglen = 9;

  const char * server_names[] = {"kleopatra.exe",
                                 "bin\\kleopatra.exe",
                                 "gpa.exe",
                                 "bin\\gpa.exe",
                                 NULL};
  const char **tmp = NULL;

  dir = get_gpg4win_dir ();
  if (!dir)
    {
      log_error ("Failed to find gpg4win dir");
      return NULL;
    }
  uiserver = read_w32_registry_string (NULL, GPG4WIN_REGKEY_3,
                                       "UI Server");
  if (!uiserver)
    {
      uiserver = read_w32_registry_string (NULL, GPG4WIN_REGKEY_2,
                                           "UI Server");
    }
  if (uiserver)
    {
      name = xmalloc (strlen (dir) + strlen (uiserver) + extra_arglen + 2);
      strcpy (stpcpy (stpcpy (name, dir), "\\"), uiserver);
      for (p = name; *p; p++)
        if (*p == '/')
          *p = '\\';
      xfree (uiserver);
    }
  if (name && !access (name, F_OK))
    {
      /* Set through registry and is accessible */
      xfree(dir);
      return name;
    }
  /* Fallbacks */
  for (tmp = server_names; *tmp; tmp++)
    {
      if (name)
        {
          xfree (name);
        }
      name = xmalloc (strlen (dir) + strlen (*tmp) + extra_arglen + 2);
      strcpy (stpcpy (stpcpy (name, dir), "\\"), *tmp);
      for (p = name; *p; p++)
        if (*p == '/')
          *p = '\\';
      if (!access (name, F_OK))
        {
          /* Found a viable candidate */
          if (strstr (name, "kleopatra.exe"))
            {
              strcat (name, " --daemon");
            }
          xfree (dir);
          return name;
        }
    }
  xfree (dir);
  log_error ("Failed to find a viable UIServer");
  return NULL;
}

int
has_high_integrity(HANDLE hToken)
{
  PTOKEN_MANDATORY_LABEL integrity_label = NULL;
  DWORD integrity_level = 0,
        size = 0;


  if (hToken == NULL || hToken == INVALID_HANDLE_VALUE)
    {
      log_debug ("Invalid parameters.");
      return 0;
    }

  /* Get the required size */
  if (!GetTokenInformation (hToken, TokenIntegrityLevel,
                            NULL, 0, &size))
    {
      if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
        {
          log_debug ("Failed to get required size.\n");
          return 0;
        }
    }
  integrity_label = (PTOKEN_MANDATORY_LABEL) LocalAlloc(0, size);
  if (integrity_label == NULL)
    {
      log_debug ("Failed to allocate label. \n");
      return 0;
    }

  if (!GetTokenInformation (hToken, TokenIntegrityLevel,
                            integrity_label, size, &size))
    {
      log_debug ("Failed to get integrity level.\n");
      LocalFree(integrity_label);
      return 0;
    }

  /* Get the last integrity level */
  integrity_level = *GetSidSubAuthority(integrity_label->Label.Sid,
                     (DWORD)(UCHAR)(*GetSidSubAuthorityCount(
                        integrity_label->Label.Sid) - 1));

  LocalFree (integrity_label);

  return integrity_level >= SECURITY_MANDATORY_HIGH_RID;
}

int
is_elevated()
{
  int ret = 0;
  HANDLE hToken = NULL;
  if (OpenProcessToken (GetCurrentProcess(), TOKEN_QUERY, &hToken))
    {
      DWORD elevation;
      DWORD cbSize = sizeof (DWORD);
      if (GetTokenInformation (hToken, TokenElevation, &elevation,
                               sizeof (TokenElevation), &cbSize))
        {
          ret = elevation;
        }
    }
  /* Elevation will be true and ElevationType TokenElevationTypeFull even
     if the token is a user token created by SAFER so we additionally
     check the integrity level of the token which will only be high in
     the real elevated process and medium otherwise. */

  ret = ret && has_high_integrity (hToken);

  if (hToken)
    CloseHandle (hToken);

  return ret;
}

int
gpgol_message_box (HWND parent, const char *utf8_text,
                   const char *utf8_caption, UINT type)
{
  wchar_t *w_text = utf8_to_wchar (utf8_text);
  wchar_t *w_caption = utf8_to_wchar (utf8_caption);
  int ret = 0;

  MSGBOXPARAMSW mbp;
  mbp.cbSize = sizeof (MSGBOXPARAMS);
  mbp.hwndOwner = parent;
  mbp.hInstance = glob_hinst;
  mbp.lpszText = w_text;
  mbp.lpszCaption = w_caption;
  mbp.dwStyle = type | MB_USERICON;
  mbp.dwLanguageId = MAKELANGID (LANG_NEUTRAL, SUBLANG_DEFAULT);
  mbp.lpfnMsgBoxCallback = NULL;
  mbp.dwContextHelpId = 0;
  mbp.lpszIcon = (LPCWSTR) MAKEINTRESOURCE (IDI_GPGOL_LOCK_ICON);

  ret = MessageBoxIndirectW (&mbp);

  xfree (w_text);
  xfree (w_caption);
  return ret;
}

static char*
expand_path (const char *path)
{
  DWORD len;
  char *p;

  len = ExpandEnvironmentStrings (path, NULL, 0);
  if (!len)
    {
      return NULL;
    }
  len += 1;
  p = xcalloc (1, len+1);
  if (!p)
    {
      return NULL;
    }
  len = ExpandEnvironmentStrings (path, p, len);
  if (!len)
    {
      xfree (p);
      return NULL;
    }
  return p;
}

static int
load_config_value (HKEY hk, const char *path, const char *key, char **val)
{
  HKEY h;
  DWORD size=0, type;
  int ec;

  *val = NULL;
  if (hk == NULL)
    {
      hk = HKEY_CURRENT_USER;
    }
  ec = RegOpenKeyEx (hk, path, 0, KEY_READ, &h);
  if (ec != ERROR_SUCCESS)
    {
      return -1;
    }

  ec = RegQueryValueEx(h, key, NULL, &type, NULL, &size);
  if (ec != ERROR_SUCCESS)
    {
      RegCloseKey (h);
      return -1;
    }
  if (type == REG_EXPAND_SZ)
    {
      char tmp[256];
      RegQueryValueEx (h, key, NULL, NULL, (BYTE*)tmp, &size);
      *val = expand_path (tmp);
    }
  else
    {
      *val = xcalloc(1, size+1);
      ec = RegQueryValueEx (h, key, NULL, &type, (BYTE*)*val, &size);
      if (ec != ERROR_SUCCESS)
        {
          xfree (*val);
          *val = NULL;
          RegCloseKey (h);
          return -1;
        }
    }
  RegCloseKey (h);
  return 0;
}

static int
store_config_value (HKEY hk, const char *path, const char *key, const char *val)
{
  HKEY h;
  int type;
  int ec;

  if (hk == NULL)
    {
      hk = HKEY_CURRENT_USER;
    }
  ec = RegCreateKeyEx (hk, path, 0, NULL, REG_OPTION_NON_VOLATILE,
                       KEY_ALL_ACCESS, NULL, &h, NULL);
  if (ec != ERROR_SUCCESS)
    {
      log_debug_w32 (ec, "creating/opening registry key `%s' failed", path);
      return -1;
    }
  type = strchr (val, '%')? REG_EXPAND_SZ : REG_SZ;
  ec = RegSetValueEx (h, key, 0, type, (const BYTE*)val, strlen (val));
  if (ec != ERROR_SUCCESS)
    {
      log_debug_w32 (ec, "saving registry key `%s'->`%s' failed", path, key);
      RegCloseKey(h);
      return -1;
    }
  RegCloseKey(h);
  return 0;
}

/* Store a key in the registry with the key given by @key and the
   value @value. */
int
store_extension_value (const char *key, const char *val)
{
  return store_config_value (HKEY_CURRENT_USER, GPGOL_REGPATH, key, val);
}

/* Load a key from the registry with the key given by @key. The value is
   returned in @val and needs to freed by the caller. */
int
load_extension_value (const char *key, char **val)
{
  return load_config_value (HKEY_CURRENT_USER, GPGOL_REGPATH, key, val);
}

int
store_extension_subkey_value (const char *subkey,
                              const char *key,
                              const char *val)
{
  int ret;
  char *path;
  gpgrt_asprintf (&path, "%s\\%s", GPGOL_REGPATH, subkey);
  ret = store_config_value (HKEY_CURRENT_USER, path, key, val);
  xfree (path);
  return ret;
}
