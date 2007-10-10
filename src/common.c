/* common.c - Common routines used by GpgOL
 *	Copyright (C) 2005, 2007 g10 Code GmbH
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
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#include <config.h>
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
  ofn.lpstrTitle = _("GPG - Save decrypted attachment");
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
          
          tmp = read_w32_registry_string (NULL, "Software\\GNU\\GnuPG",
                                          "HomeDir");
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


/* Do in-place decoding of quoted-printable data of LENGTH in BUFFER.
   Returns the new length of the buffer. */
size_t
qp_decode (char *buffer, size_t length)
{
  char *d, *s;

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
          }
        else if (length > 1 && s[1] == '\n')
          {
            /* Soft line break with only a Unix line terminator. */
            s += 2;
            length -= 1;
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

