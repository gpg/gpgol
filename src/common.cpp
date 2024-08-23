/* common.c - Common routines used by GpgOL
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 * 2015, 2016, 2017  Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
 * 2020 g10 Code GmbH
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
#include "cpphelp.h"

#include <string>
#include <fstream>
#include <regex>
#include <algorithm>
#include <vector>

#include <gpgme++/context.h>
#include <gpgme++/error.h>
#include <gpgme++/configuration.h>

#define COPYBUFSIZE (8 * 1024)

HINSTANCE glob_hinst = NULL;

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
  log_debug ("%s:%s: done", SRCNAME, __func__);
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
#if defined(_WIN64)
#define CROSS_ACCESS KEY_WOW64_32KEY
#else
#define CROSS_ACCESS KEY_WOW64_64KEY
#endif

/* Read a registry value from HKLM or HKCU of type dword and
   return 1 if the value is larger then 0, 0 if it is zero and
   -1 if it is not set or not found. */
int
read_reg_bool (HKEY root, const char *path, const char *value)
{
  TSTART;
  HKEY h;
  HKEY tmp = root ? root : HKEY_CURRENT_USER;

  int err = RegOpenKeyEx (tmp, path , 0, KEY_READ, &h);
  if (err != ERROR_SUCCESS)
    {
      log_debug ("%s:%s: not found %s",
                 SRCNAME, __func__, path);
      if (!root)
        {
          TRETURN read_reg_bool (HKEY_LOCAL_MACHINE, path, value);
        }
      TRETURN -1;
    }
  DWORD type;
  err = RegQueryValueEx (h, value, NULL, &type, NULL, NULL);
  if (err != ERROR_SUCCESS || type != REG_DWORD)
    {
      log_debug ("%s:%s: No type or key for %s",
                 SRCNAME, __func__, value);
      if (!root)
        {
          TRETURN read_reg_bool (HKEY_LOCAL_MACHINE, path, value);
        }
      TRETURN -1;
    }
  DWORD data;
  DWORD size = sizeof (DWORD);
  err = RegQueryValueEx (h, value, NULL, NULL, (LPBYTE)&data,
                         &size);
  if (err != ERROR_SUCCESS)
    {
      log_debug ("%s:%s: Failed to find value of %s",
                 SRCNAME, __func__, value);
      if (!root)
        {
          TRETURN read_reg_bool (HKEY_LOCAL_MACHINE, path, value);
        }
      TRETURN -1;
    }
  TRETURN !!data;
}


std::string
_readRegStr (HKEY root_key, const char *dir,
             const char *name, bool alternate)
{
#ifndef _WIN32
    (void) root_key; (void)alternate; (void)dir; (void)name;
    return std::string();
#else
    HKEY key_handle;
    DWORD n1, nbytes, type;
    std::string ret;

    DWORD flags = KEY_READ;

    if (alternate) {
        flags |= CROSS_ACCESS;
    }

    if (RegOpenKeyExA(root_key, dir, 0, flags, &key_handle)) {
        return ret;
    }

    nbytes = 1;
    if (RegQueryValueExA(key_handle, name, 0, nullptr, nullptr, &nbytes)) {
        RegCloseKey (key_handle);
        return ret;
    }
    n1 = nbytes+1;
    char result[n1];
    if (RegQueryValueExA(key_handle, name, 0, &type, (LPBYTE)result, &n1)) {
        RegCloseKey(key_handle);
        return ret;
    }
    RegCloseKey(key_handle);
    result[nbytes] = 0; /* make sure it is really a string  */
    ret = result;
    if (type == REG_EXPAND_SZ && strchr (result, '%')) {
        n1 += 1000;
        char tmp[n1 +1];

        nbytes = ExpandEnvironmentStringsA(ret.c_str(), tmp, n1);
        if (nbytes && nbytes > n1) {
            n1 = nbytes;
            char tmp2[n1 +1];
            nbytes = ExpandEnvironmentStringsA(result, tmp2, n1);
            if (nbytes && nbytes > n1) {
                /* oops - truncated, better don't expand at all */
                return ret;
            }
            tmp2[nbytes] = 0;
            ret = tmp2;
        } else if (nbytes) { /* okay, reduce the length */
            tmp[nbytes] = 0;
            ret = tmp;
        }
    }
    return ret;

#endif
}

std::string
readRegStr (const char *root, const char *dir, const char *name)
{
#ifndef _WIN32
    (void)root; (void)dir; (void)name;
    return std::string();
#else
    HKEY root_key;
    std::string ret;
    if (!(root_key = get_root_key(root))) {
        return ret;
    }
    if (root == nullptr)
      {
        /* Nullptr so we first look into HKLM if we have
           an override */
        ret = _readRegStr (HKEY_LOCAL_MACHINE, dir, name, false);
        if (ret.empty()) {
            // Try alternate as fallback
            ret = _readRegStr (HKEY_LOCAL_MACHINE, dir, name, true);
        }
        if (ret.size() && ret[ret.size() - 1] == '!')
          {
            // Using override reg value
            log_dbg ("Using override for %s", name);
            ret.pop_back();
            return ret;
          }
      }
    ret = _readRegStr (root_key, dir, name, false);
    if (ret.empty()) {
        // Try alternate as fallback
        ret = _readRegStr (root_key, dir, name, true);
    }
    if (ret.empty()) {
        // Try local machine as fallback.
        ret = _readRegStr (HKEY_LOCAL_MACHINE, dir, name, false);
        if (ret.empty()) {
            // Try alternative registry view as fallback
            ret = _readRegStr (HKEY_LOCAL_MACHINE, dir, name, true);
        }
    }
    return ret;
#endif
}

/* Return a string from the Win32 Registry or NULL in case of error.
   Caller must release the return value.  A NULL for root is an alias
   for HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE in turn.  NOTE: The value
   is allocated with a plain xmalloc () - use xfree () and not the usual
   xfree(). */
char *
read_w32_registry_string (const char *root, const char *dir, const char *name)
{
  const auto ret = readRegStr (root, dir, name);
  if (ret.empty())
    {
      return nullptr;
    }
  return xstrdup (ret.c_str ());
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
  dname = (char*) xmalloc (strlen (instdir) + strlen (SDDIR) + 1);
  if (!dname)
    {
      xfree (instdir);
      return NULL;
    }
  p = dname;
  strcpy (p, instdir);
  p += strlen (instdir);
  strcpy (p, SDDIR);

  xfree (instdir);

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
  ptr = (char *) xmalloc (i + 2 * j + 1);
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

BOOL
DeleteFileUtf8 (const char *utf8Name)
{
  SetLastError (0);
  if (!utf8Name)
    {
      STRANGEPOINT;
      return false;
    }
  wchar_t *wname = utf8_to_wchar (utf8Name);
  BOOL ret = DeleteFileW (wname);
  xfree (wname);
  return ret;
}

HANDLE
CreateFileUtf8 (const char *utf8Name)
{
  SetLastError (0);
  if (!utf8Name)
    {
      return INVALID_HANDLE_VALUE;
    }

  wchar_t *wname = utf8_to_wchar (utf8Name);
  if (!wname)
    {
      TRACEPOINT;
      return INVALID_HANDLE_VALUE;
    }

  auto ret = CreateFileW (wname,
                          GENERIC_WRITE | GENERIC_READ,
                          FILE_SHARE_READ,
                          NULL,
                          CREATE_NEW,
                          FILE_ATTRIBUTE_TEMPORARY,
                          NULL);
  xfree (wname);
  return ret;
}

int
readFullFile (HANDLE hFile, GpgME::Data &data)
{
  char buf[COPYBUFSIZE];
  DWORD bRead = 0;
  BOOL ret;
  while ((ret = ReadFile (hFile, buf, COPYBUFSIZE, &bRead, nullptr)))
    {
      if (!bRead)
        {
          // EOF
          break;
        }
      data.write (buf, bRead);
    }
  if (!ret && bRead)
    {
      log_err ("Failed to read from file");
      TRETURN -1;
    }
  TRETURN 0;
}

static std::string
getTmpPathUtf8 ()
{
  static std::string ret;
  if (!ret.empty())
    {
      return ret;
    }
  wchar_t tmpPath[MAX_PATH + 2];

  if (!GetTempPathW (MAX_PATH, tmpPath))
    {
      log_error ("%s:%s: Could not get tmp path.",
                 SRCNAME, __func__);
      return ret;
    }

  char *utf8Name = wchar_to_utf8 (tmpPath);

  if (!utf8Name)
    {
      TRACEPOINT;
      return ret;
    }
  ret = utf8Name;
  xfree (utf8Name);
  return ret;
}

/* Open a file in a temporary directory, take name as a
   suggestion and put the open Handle in outHandle.
   Returns the actually used file name in case there
   were other files with that name. */
wchar_t*
get_tmp_outfile (const wchar_t *name, HANDLE *outHandle)
{
  TSTART
  auto utf8Name = sanitizeFileName(wchar_to_utf8_string (name));
  const auto tmpPath = getTmpPathUtf8 ();

  if (utf8Name.empty() || tmpPath.empty())
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  auto outName = tmpPath + utf8Name;

  log_data("%s:%s: Attachment candidate is %s",
                  SRCNAME, __func__, outName.c_str ());

  int tries = 1;
  bool deletedFile = false;
  while ((*outHandle = CreateFileUtf8 (outName.c_str ())) == INVALID_HANDLE_VALUE)
    {

      DWORD err = GetLastError ();
      log_debug_w32 (err, "%s:%s: Failed to open candidate '%s'",
                     SRCNAME, __func__, anonstr (outName.c_str()));

      if (!deletedFile && err == ERROR_FILE_EXISTS)
        {
          /* Try to delete the file. If someone is reading it we should
             not be able to delete it. */
          if (DeleteFileUtf8 (outName.c_str ()))
            {
              log_dbg ("Deleted existing tmp file '%s'", anonstr (outName.c_str ()));
              deletedFile = true;
              continue;
            }
        }
      /* Just to make sure we do not loop endlessly if we get strange return
       * values. */
      deletedFile = false;

      char *outNameC = xstrdup ((tmpPath + utf8Name).c_str ());

      const auto lastBackslash = strrchr (outNameC, '\\');
      if (!lastBackslash)
        {
          /* This is an error because tmp name by definition contains one */
          log_error ("%s:%s: No backslash in origname '%s'",
                     SRCNAME, __func__, outNameC);
          xfree (outNameC);
          TRETURN NULL;
        }

      auto fileExt = strchr (lastBackslash, '.');
      if (fileExt)
        {
          *fileExt = '\0';
          ++fileExt;
        }
      // OutNameC is now without an extension and if
      // there is a file ext it now points to the extension.

      outName = outNameC + std::string("_") + std::to_string(tries++);

      if (fileExt)
        {
          outName += std::string(".") + fileExt;
        }
      xfree (outNameC);

      if (tries == 50)
        {
          /* Mmh fishy, maybe the name cannot be created on the file
             system. Let's switch to a generic name. */
          log_dbg ("Can't find an attachment name. "
                   "Switching over to a generic attachment name.");
          outName = tmpPath + "attachment";
          utf8Name = "attachment";
          if (fileExt)
            {
              outName += ".";
              outName += fileExt;
              utf8Name += ".";
              utf8Name += fileExt;
            }
        }
      if (tries > 100)
        {
          char *buf;
          gpgrt_asprintf (&buf,"Failed to obtain temporary filename '%s'"
                             " please check that the folder '%s' exists and the file is"
                             " writable. Look into configuration dialog for debug options.",
                             wchar_to_utf8_string (name).c_str (), tmpPath.c_str ());
          memdbg_alloc (buf);
          /* You have to know when to give up,.. */
          gpgol_message_box (NULL, buf, _("GpgOL Error"), MB_OK);
          xfree (buf);
          log_error ("%s:%s: Could not get a name out of 100 tries",
                     SRCNAME, __func__);
          TRETURN NULL;
        }
    }

  TRETURN utf8_to_wchar (outName.c_str ());
}

char *
get_tmp_outfile_utf8 (const char *name, HANDLE *outHandle)
{
  wchar_t *wname = utf8_to_wchar (name);
  wchar_t *wret = get_tmp_outfile (wname, outHandle);
  xfree (wname);
  char *ret = wchar_to_utf8 (wret);
  xfree (wret);
  return ret;
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
      name = (char*) xmalloc (strlen (dir) + strlen (uiserver) + extra_arglen + 2);
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
      name = (char *) xmalloc (strlen (dir) + strlen (*tmp) + extra_arglen + 2);
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

void
gpgol_bug (HWND parent, int code)
{
  const char *bugmsg = utf8_gettext ("Operation failed.\n\n"
                "This is usually caused by a bug in GpgOL or an error in your setup.\n"
                "Please see https://www.gpg4win.org/reporting-bugs.html "
                "or ask your Administrator for support.");
  char *with_code;
  gpgrt_asprintf (&with_code, "%s\nCode: %i", bugmsg, code);
  memdbg_alloc (with_code);
  gpgol_message_box (parent,
                     with_code,
                     _("GpgOL Error"), MB_OK);
  xfree (with_code);
  return;
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
  if (!val)
    {
      STRANGEPOINT;
      return -1;
    }
  *val = read_w32_registry_string (nullptr, GPGOL_REGPATH, key);
  log_debug ("%s:%s: LoadReg '%s' val '%s'",
             SRCNAME, __func__, key ? key : "null",
             *val ? *val : "null");
  return 0;
}

int
store_extension_subkey_value (const char *subkey,
                              const char *key,
                              const char *val)
{
  int ret;
  char *path;
  gpgrt_asprintf (&path, "%s\\%s", GPGOL_REGPATH, subkey);
  memdbg_alloc (path);
  ret = store_config_value (HKEY_CURRENT_USER, path, key, val);
  xfree (path);
  return ret;
}

bool
in_de_vs_mode()
{
/* We cache the values only once. A change requires restart.
     This is because checking this is very expensive as gpgconf
     spawns each process to query the settings. */
  static bool checked;
  static bool vs_mode;

  if (checked)
    {
      return vs_mode;
    }
  checked = true;
  GpgME::Error err;
  const auto components = GpgME::Configuration::Component::load (err);
  log_debug ("%s:%s: Checking for de-vs mode.",
             SRCNAME, __func__);
  if (err)
    {
      log_error ("%s:%s: Failed to get gpgconf components: %s",
                 SRCNAME, __func__, err.asString ());
      vs_mode = false;
      return vs_mode;
    }
  for (const auto &component: components)
    {
      if (component.name () && !strcmp (component.name (), "gpg"))
        {
          for (const auto &option: component.options ())
            {
              if (option.name () && !strcmp (option.name (), "compliance") &&
                  option.currentValue ().stringValue () &&
#ifdef HAVE_W32_SYSTEM
                  !stricmp (option.currentValue ().stringValue (), "de-vs"))
#else
                  !strcasecmp (option.currentValue ().stringValue (), "de-vs"))
#endif
                {
                  log_debug ("%s:%s: Detected de-vs mode",
                             SRCNAME, __func__);
                  vs_mode = true;
                  return vs_mode;
                }
            }
          vs_mode = false;
          return vs_mode;
        }
    }
  vs_mode = false;
  return false;
}

const char *
de_vs_name (bool isCompliant)
{
  static std::string compName;
  static std::string uncompName;
  if (!compName.empty() && isCompliant)
    {
      return compName.c_str ();
    }
  if (!uncompName.empty() && !isCompliant)
    {
      return uncompName.c_str ();
    }
  TSTART;
  /* Only look once */
  compName = utf8_gettext ("VS-NfD compliant");
  uncompName = utf8_gettext ("not VS-NfD compliant");
  /* Find the libkleopatrarc */
  char *instdir = get_gpg4win_dir();
  if (!instdir)
    {
      STRANGEPOINT;
      TRETURN isCompliant ? compName.c_str () : uncompName.c_str ();
    }
  std::string filename = instdir;
  filename += "\\share\\libkleopatrarc";

  std::ifstream file(filename.c_str ());
  if (!file.is_open ())
    {
      log_err ("Failed to open '%s'", filename.c_str ());
      TRETURN isCompliant ? compName.c_str () : uncompName.c_str ();
    }
  std::string line;
  bool in_de_vs_filter = false;
  bool in_not_de_vs_filter = false;
  const char *lname = gettext_localename ();
  /* lname is never null */
  std::string localemain = gpgol_split (lname, '_')[0];
  std::string idealname = std::string ("Name[") +
    localemain + std::string("]");
  while (std::getline(file, line))
    {
      if (starts_with (line, "["))
        {
          in_de_vs_filter = false;
          in_not_de_vs_filter = false;
          continue;
        }
      if (starts_with (line, "id=de-vs-filter"))
        {
          in_de_vs_filter = true;
          continue;
        }
      if (starts_with (line, "id=not-de-vs-filter"))
        {
          in_not_de_vs_filter = true;
          continue;
        }

      if ((in_de_vs_filter || in_not_de_vs_filter) && starts_with (line, "Name"))
        {
          const auto split = gpgol_split (line, '=');
          if (split.size() != 2)
            {
              log_err ("Invalid libkleopatrarc line: %s", line.c_str());
              continue;
            }
          if (split[0] == "Name")
            {
              if (in_de_vs_filter)
                {
                  compName = split[1];
                }
              else
                {
                  uncompName = split[1];
                }
            }
          if (split[0] == idealname)
            {
              log_dbg ("Found localized de-vs name %s", split[1].c_str ());
              if (in_de_vs_filter)
                {
                  compName = split[1];
                }
              else
                {
                  uncompName = split[1];
                }
            }
        }
    }
  file.close();
  TRETURN isCompliant ? compName.c_str () : uncompName.c_str ();
}

std::string
compliance_string (bool forVerify, bool forDecrypt, bool isCompliant)
{
  const char *comp_name = de_vs_name (isCompliant);
  if (forDecrypt && forVerify)
    {
      return asprintf_s (utf8_gettext ("The message is %s."), comp_name);
    }
  if (forVerify)
    {
      /* TRANSLATORS %s is compliance name like VS-NfD */
      return asprintf_s (utf8_gettext ("The signature is %s."), comp_name);
    }
  if (forDecrypt)
    {
      /* TRANSLATORS %s is compliance name like VS-NfD */
      return asprintf_s (utf8_gettext ("The encryption is %s."),
                            comp_name);
    }
  /* Should not happen */
  log_err ("Invalid call to compliance string.");
  STRANGEPOINT;
  return std::string();
}

char *
get_gpgme_w32_inst_dir (void)
{
  char *gpg4win_dir = get_gpg4win_dir ();
  char *tmp;
  gpgrt_asprintf (&tmp, "%s\\bin\\gpgme-w32spawn.exe", gpg4win_dir);
  memdbg_alloc (tmp);

  if (!access(tmp, R_OK))
    {
      xfree (tmp);
      gpgrt_asprintf (&tmp, "%s\\bin", gpg4win_dir);
      memdbg_alloc (tmp);
      xfree (gpg4win_dir);
      return tmp;
    }
  xfree (tmp);
  gpgrt_asprintf (&tmp, "%s\\gpgme-w32spawn.exe", gpg4win_dir);
  memdbg_alloc (tmp);

  if (!access(tmp, R_OK))
    {
      xfree (tmp);
      return gpg4win_dir;
    }
  OutputDebugString("Failed to find gpgme-w32spawn.exe!");
  return NULL;
}

#define WINDOWS_DEVICES_PATTERN "(CON|AUX|PRN|NUL|COM[1-9]|LPT[1-9])(\\..*)?"
#define SLASHES "/\\"

static const std::regex&
windowsDeviceNoSubDirPattern()
{
  static const std::regex rc ("^" WINDOWS_DEVICES_PATTERN "$", std::regex_constants::icase);
  return rc;
}

static const std::regex&
windowsDeviceSubDirPattern()
{
  static const std::regex rc ("^.*[/\\\\]" WINDOWS_DEVICES_PATTERN "$", std::regex_constants::icase);
  return rc;
}

static const char notAllowedCharsSubDir[] = ",^@={}[]~!?:&*\"|#%<>$\"'();`' ";
static const char notAllowedCharsNoSubDir[] = ",^@={}[]~!?:&*\"|#%<>$\"'();`' " SLASHES;
static const char* notAllowedSubStrings[] = {".."};

bool
validateWindowsFileName(const std::string& name, bool allowDirectories)
{
  if (name.empty ())
    {
      return false;
    }

  // Characters
  const char* notAllowedChars = allowDirectories ? notAllowedCharsSubDir
                                                 : notAllowedCharsNoSubDir;
  for (const char* c = notAllowedChars; *c; ++c)
    {
      if (name.find(*c) != std::string::npos)
        {
          return false;
        }
    }

  // Substrings
  for (const auto& subStr : notAllowedSubStrings) {
      if (name.find(subStr) != std::string::npos)
        {
          return false;
        }
  }

  // Windows devices
  bool matchesWinDevice = std::regex_match(name, windowsDeviceNoSubDirPattern());
  if (!matchesWinDevice && allowDirectories)
    {
      matchesWinDevice = std::regex_match(name, windowsDeviceSubDirPattern());
    }
  return !matchesWinDevice;
}

std::string
sanitizeFileName(std::string name, bool allowDirectories)
{
  if (name.empty())
    {
      return "attachment";
    }

  // Characters
  const char* notAllowedChars = allowDirectories ? notAllowedCharsSubDir
                                                 : notAllowedCharsNoSubDir;
  for (const char* c = notAllowedChars; *c; ++c)
    {
      name.erase(std::remove(name.begin(), name.end(), *c), name.end());
    }

  // Substrings
  for (const auto& subStr : notAllowedSubStrings)
    {
      size_t pos;
      while ((pos = name.find(subStr)) != std::string::npos)
        {
          name.erase(pos, std::strlen(subStr));
        }
    }

  // Windows devices
  if (std::regex_match(name, windowsDeviceNoSubDirPattern()) ||
      (allowDirectories && std::regex_match(name, windowsDeviceSubDirPattern())))
    {
      name = "_" + name;
    }

  return name.empty() ? "_" : name;
}
