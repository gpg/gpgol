/* main.c - DLL entry point
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
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

#include <windows.h>
#include <wincrypt.h>
#include <ctype.h>
#include <winnls.h>
#include <unistd.h>

#include "mymapi.h"
#include "mymapitags.h"

#include "common.h"
#include "mymapi.h"

/* Local function prototypes. */
static char *get_locale_dir (void);
static void drop_locale_dir (char *locale_dir);

/* The major version of Outlook we are attached to */
int g_ol_version_major;


/* For certain operations we need to acquire a log on the logging
   functions.  This lock is controlled by this Mutex. */
HANDLE log_mutex;

/* Early initialization of this module.  This is done right at startup
   with only one thread running.  Should be called only once. Returns
   0 on success. */
static int
initialize_main (void)
{
  SECURITY_ATTRIBUTES sa;

  memset (&sa, 0, sizeof sa);
  sa.bInheritHandle = FALSE;
  sa.lpSecurityDescriptor = NULL;
  sa.nLength = sizeof sa;
  log_mutex = CreateMutex (&sa, FALSE, NULL);
  return log_mutex? 0 : -1;
}

void
i18n_init (void)
{
  char *locale_dir;

#ifdef ENABLE_NLS
# ifdef HAVE_LC_MESSAGES
  setlocale (LC_TIME, "");
  setlocale (LC_MESSAGES, "");
# else
  setlocale (LC_ALL, "" );
# endif
#endif

  locale_dir = get_locale_dir ();
  if (locale_dir)
    {
      bindtextdomain (PACKAGE_GT, locale_dir);
      drop_locale_dir (locale_dir);
    }
  textdomain (PACKAGE_GT);
}

static char *
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

/* Entry point called by DLL loader. */
int WINAPI
DllMain (HINSTANCE hinst, DWORD reason, LPVOID reserved)
{
  (void)reserved;

  if (reason == DLL_PROCESS_ATTACH)
    {
      set_global_hinstance (hinst);

      gpg_err_init ();

      /* Set the installation directory for GpgME so that
         it can find tools like gpgme-w32-spawn correctly. */
      char *instdir = get_gpgme_w32_inst_dir();
      gpgme_set_global_flag ("w32-inst-dir", instdir);
      xfree (instdir);

      /* The next call initializes subsystems of gpgme and should be
         done as early as possible.  The actual return value (the
         version string) is not used here.  It may be called at any
         time later for this. */
      gpgme_check_version (NULL);

      /* Early initializations of our subsystems. */
      if (initialize_main ())
        return FALSE;
    }
  else if (reason == DLL_PROCESS_DETACH)
    {
      gpg_err_deinit (0);
    }

  return TRUE;
}

static char *
get_locale_dir (void)
{
  char *instdir;
  char *p;
  char *dname;

  instdir = get_gpg4win_dir();
  if (!instdir)
    return NULL;

  /* Build the key: "<instdir>/share/locale".  */
#define SLDIR "\\share\\locale"
  dname = xmalloc (strlen (instdir) + strlen (SLDIR) + 1);
  if (!dname)
    {
      xfree (instdir);
      return NULL;
    }
  p = dname;
  strcpy (p, instdir);
  p += strlen (instdir);
  strcpy (p, SLDIR);

  xfree (instdir);

  return dname;
}


static void
drop_locale_dir (char *locale_dir)
{
  xfree (locale_dir);
}


static int
get_conf_bool (const char *name, int defaultVal)
{
  char *val = NULL;
  int ret;
  load_extension_value (name, &val);
  ret = val == NULL ? defaultVal : *val != '1' ? 0 : 1;
  xfree (val);
  return ret;
}

/* Read option settings from the Registry. */
void
read_options (void)
{
  char *val = NULL;

  /* Set the log file first so that output from this function is
     logged too.  */
  load_extension_value ("logFile", &val);
  set_log_file (val);
  xfree (val); val = NULL;

  /* Parse the debug flags.  */
  load_extension_value ("enableDebug", &val);
  opt.enable_debug = 0;
  if (val)
    {
      char *p, *pend;

      trim_spaces (val);
      for (p = val; p; p = pend)
        {
          pend = strpbrk (p, ", \t\n\r\f");
          if (pend)
            {
              *pend++ = 0;
              pend += strspn (pend, ", \t\n\r\f");
            }
          if (isascii (*p) && isdigit (*p))
            opt.enable_debug |= strtoul (p, NULL, 0);
          else if (!strcmp (p, "ioworker"))
            opt.enable_debug |= DBG_IOWORKER;
          else if (!strcmp (p, "ioworker-extra"))
            opt.enable_debug |= DBG_IOWORKER_EXTRA;
          else if (!strcmp (p, "filter"))
            opt.enable_debug |= DBG_FILTER;
          else if (!strcmp (p, "filter-extra"))
            opt.enable_debug |= DBG_FILTER_EXTRA;
          else if (!strcmp (p, "memory"))
            opt.enable_debug |= DBG_MEMORY;
          else if (!strcmp (p, "commands"))
            opt.enable_debug |= DBG_COMMANDS;
          else if (!strcmp (p, "mime-parser"))
            opt.enable_debug |= DBG_MIME_PARSER;
          else if (!strcmp (p, "mime-data"))
            opt.enable_debug |= DBG_MIME_DATA;
          else if (!strcmp (p, "oom"))
            opt.enable_debug |= DBG_OOM;
          else if (!strcmp (p, "oom-extra"))
            opt.enable_debug |= DBG_OOM_EXTRA;
          else
            log_debug ("invalid debug flag `%s' ignored", p);
        }
    }
  else
    {
      /* To help the user enable debugging make sure that the registry
         key exists.  Note that the other registry keys are stored
         after using the configuration dialog.  */
      store_extension_value ("enableDebug", "0");
    }
  /* Yes we use free here because memtracing did not track the alloc
     as the option for debuging was not read before. */
  free (val); val = NULL;
  if (opt.enable_debug)
    log_debug ("enabled debug flags:%s%s%s%s%s%s%s%s%s%s\n",
               (opt.enable_debug & DBG_IOWORKER)? " ioworker":"",
               (opt.enable_debug & DBG_IOWORKER_EXTRA)? " ioworker-extra":"",
               (opt.enable_debug & DBG_FILTER)? " filter":"",
               (opt.enable_debug & DBG_FILTER_EXTRA)? " filter-extra":"",
               (opt.enable_debug & DBG_MEMORY)? " memory":"",
               (opt.enable_debug & DBG_COMMANDS)? " commands":"",
               (opt.enable_debug & DBG_MIME_PARSER)? " mime-parser":"",
               (opt.enable_debug & DBG_MIME_DATA)? " mime-data":"",
               (opt.enable_debug & DBG_OOM)? " oom":"",
               (opt.enable_debug & DBG_OOM_EXTRA)? " oom-extra":""
               );

  opt.enable_smime = get_conf_bool ("enableSmime", 0);
  opt.encrypt_default = get_conf_bool ("encryptDefault", 0);
  opt.sign_default = get_conf_bool ("signDefault", 0);
  opt.inline_pgp = get_conf_bool ("inlinePGP", 0);
  opt.reply_crypt = get_conf_bool ("replyCrypt", 1);
  opt.prefer_smime = get_conf_bool ("preferSmime", 0);
  opt.autoresolve = get_conf_bool ("autoresolve", 1);
  opt.automation = get_conf_bool ("automation", 1);
  opt.autosecure = get_conf_bool ("autosecure", 1);
  opt.autotrust = get_conf_bool ("autotrust", 0);
  opt.sync_enc = get_conf_bool ("syncEnc", 0);
  opt.smime_html_warn_shown = get_conf_bool ("smimeHtmlWarnShown", 0);

  if (!opt.automation)
    {
      // Disabling automation is a shorthand to disable the
      // others, too.
      opt.autosecure = 0;
      opt.autoresolve = 0;
      opt.autotrust = 0;
    }
}


/* Write current options back to the Registry. */
int
write_options (void)
{
  struct
  {
    const char *name;
    int  mode;
    int  value;
    char *s_val;
  } table[] = {
    {"smimeHtmlWarnShown",       0, opt.smime_html_warn_shown, NULL},
    {NULL, 0, 0, NULL}
  };
  char buf[32];
  int rc, i;
  const char *string;

  for (i=0; table[i].name; i++)
    {
      switch (table[i].mode)
        {
        case 0:
          string = table[i].value? "1": "0";
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, string);
          rc = store_extension_value (table[i].name, string);
          break;
        case 1:
          sprintf (buf, "%d", table[i].value);
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, buf);
          rc = store_extension_value (table[i].name, buf);
          break;
        case 2:
          string = table[i].s_val? table[i].s_val : "";
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, string);
          rc = store_extension_value (table[i].name, string);
          break;
/*         case 3: */
/*           buf[0] = '0'; */
/*           buf[1] = 0; */
/*           switch (opt.default_protocol) */
/*             { */
/*             case PROTOCOL_UNKNOWN: buf[0] = '0'; /\* auto *\/ break; */
/*             case PROTOCOL_OPENPGP: buf[0] = '1'; break; */
/*             case PROTOCOL_SMIME:   buf[0] = '2'; break; */
/*             } */
/*           log_debug ("storing option `%s' value=`%s'\n", */
/*                      table[i].name, buf); */
/*           rc = store_extension_value (table[i].name, buf); */
/*           break;   */

        case 4:
          sprintf (buf, "0x%x", table[i].value);
          log_debug ("storing option `%s' value=`%s'\n",
                     table[i].name, buf);
          rc = store_extension_value (table[i].name, buf);
          break;

        default:
          rc = -1;
          break;
        }
      if (rc)
        log_error ("error storing option `%s': rc = %d\n", table[i].name, rc);
    }

  return 0;
}
