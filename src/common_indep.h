#ifndef COMMON_INDEP_H
#define COMMON_INDEP_H
/* common_indep.h - Common, platform indepentent routines used by GpgOL
 * Copyright (C) 2005, 2007, 2008 g10 Code GmbH
 * Copyright (C) 2016 by Bundesamt für Sicherheit in der Informationstechnik
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
#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <string.h>

#include <gpgme.h>

#include "xmalloc.h"

#include "debug.h"

#include "memdbg.h"

#ifdef HAVE_W32_SYSTEM
/* Not so independenent ;-) need this for logging HANDLE */
# include <windows.h>
#endif

/* The Registry key used by Gpg4win.  */
#ifdef _WIN64
# define GPG4WIN_REGKEY_2  "Software\\Wow6432Node\\GNU\\GnuPG"
#else
# define GPG4WIN_REGKEY_2  "Software\\GNU\\GnuPG"
#endif
#ifdef _WIN64
# define GPG4WIN_REGKEY_3  "Software\\Wow6432Node\\Gpg4win"
#else
# define GPG4WIN_REGKEY_3  "Software\\Gpg4win"
#endif
/* Identifiers for the protocol.  We use different one than those use
   by gpgme.  FIXME: We might want to define an unknown protocol to
   non-null and define such a value also in gpgme. */
typedef enum
  {
    PROTOCOL_UNKNOWN = 0,
    PROTOCOL_OPENPGP = 1000,
    PROTOCOL_SMIME   = 1001
  }
protocol_t;


/* Possible options for the recipient dialog. */
enum
  {
    OPT_FLAG_TEXT     =  2,
    OPT_FLAG_FORCE    =  4,
    OPT_FLAG_CANCEL   =  8
  };


typedef enum
  {
    GPG_FMT_NONE = 0,       /* do not encrypt attachments */
    GPG_FMT_CLASSIC = 1,    /* encrypt attachments without any encoding */
    GPG_FMT_PGP_PEF = 2     /* use the PGP partioned encoding format (PEF) */
  }
gpgol_format_t;

/* Type of a message. */
typedef enum
  {
    OPENPGP_NONE = 0,
    OPENPGP_MSG,
    OPENPGP_SIG,
    OPENPGP_CLEARSIG,
    OPENPGP_PUBKEY,   /* Note, that this type is only partly supported */
    OPENPGP_SECKEY    /* Note, that this type is only partly supported */
  }
openpgp_t;

/* The list of message types we support in GpgOL.  */
typedef enum
  {
    MSGTYPE_UNKNOWN = 0,
    MSGTYPE_SMIME,         /* Original SMIME class. */
    MSGTYPE_GPGOL,
    MSGTYPE_GPGOL_MULTIPART_SIGNED,
    MSGTYPE_GPGOL_MULTIPART_ENCRYPTED,
    MSGTYPE_GPGOL_OPAQUE_SIGNED,
    MSGTYPE_GPGOL_OPAQUE_ENCRYPTED,
    MSGTYPE_GPGOL_CLEAR_SIGNED,
    MSGTYPE_GPGOL_PGP_MESSAGE,
    MSGTYPE_GPGOL_WKS_CONFIRMATION
  }
msgtype_t;

typedef enum
  {
    ATTACHTYPE_UNKNOWN = 0,
    ATTACHTYPE_MOSS = 1,         /* The original MOSS message (ie. a
                                    S/MIME or PGP/MIME message. */
    ATTACHTYPE_FROMMOSS = 2,     /* Attachment created from MOSS.  */
    ATTACHTYPE_MOSSTEMPL = 3,    /* Attachment has been created in the
                                    course of sending a message */
    ATTACHTYPE_PGPBODY = 4,      /* Attachment contains the original
                                    PGP message body of PGP inline
                                    encrypted messages.  */
    ATTACHTYPE_FROMMOSS_DEC = 5  /* A FROMMOSS attachment that has been
                                    temporarily decrypted and needs to be
                                    encrypted before it is written back
                                    into storage. */
  }
attachtype_t;

/* An object to collect information about one MAPI attachment.  */
struct mapi_attach_item_s
{
  int end_of_table;     /* True if this is the last plus one entry of
                           the table. */
  void *private_mapitable; /* Only for use by mapi_release_attach_table. */

  int mapipos;          /* The position which needs to be passed to
                           MAPI to open the attachment.  -1 means that
                           there is no valid attachment.  */

  int method;           /* MAPI attachment method. */
  char *filename;       /* Malloced filename of this attachment or NULL. */

  /* Malloced string with the MIME attrib or NULL.  Parameters are
     stripped off thus a compare against "type/subtype" is
     sufficient. */
  char *content_type;

  /* If not NULL the parameters of the content_type. */
  const char *content_type_parms;

  /* If not NULL the content_id */
  char *content_id;

  /* The attachment type from Property GpgOL Attach Type.  */
  attachtype_t attach_type;
};
typedef struct mapi_attach_item_s mapi_attach_item_t;


/* Passphrase callback structure. */
struct passphrase_cb_s
{
  gpgme_key_t signer;
  gpgme_ctx_t ctx;
  char keyid[16+1];
  char *user_id;
  char *pass;
  int opts;
  int ttl;  /* TTL of the passphrase. */
  unsigned int decrypt_cmd:1; /* 1 = show decrypt dialog, otherwise secret key
			         selection. */
  unsigned int hide_pwd:1;
  unsigned int last_was_bad:1;
};

/* Global options - initialized to default by main.c. */
#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

#ifdef __cplusplus
extern
#endif

struct
{
  int enable_debug;	     /* Enable extra debug options.  Values
                                larger than 1 increases the debug log
                                verbosity.  */
  int enable_smime;	     /* Enable S/MIME support. */
  int encrypt_default;       /* Encrypt by default. */
  int sign_default;          /* Sign by default. */
  int prefer_html;           /* Prefer html in html/text alternatives. */
  int inline_pgp;            /* Only for Addin. Use Inline PGP by default. */
  int autoresolve;           /* Autresolve keys with --locate-keys. */
  int autosecure;             /* Autmatically encrypt if locate returns enough validity. */
  int reply_crypt;           /* Only for Addin. Encrypt / Sign based on cryptostatus. */
  int automation;            /* General automation */
  int autotrust;             /* TOFU configured for GpgOL */
  int sync_enc;              /* Disabed async encryption */
  int sync_dec;              /* Disabed async decryption */
  int prefer_smime;          /* S/MIME prefered when autoresolving */
  int smime_html_warn_shown; /* Flag to save if unsigned smime warning was shown */
  int autoretrieve;           /* Use --auto-key-retrieve. */
  int search_smime_servers;  /* Search for S/MIME keys on all configured S/MIME keyservers
                                for each new unknown mail */

  /* The forms revision number of the binary.  */
  int forms_revision;

} opt;


/* The state object used by b64_decode.  */
struct b64_state_s
{
  int idx;
  unsigned char val;
  int stop_seen;
  int invalid_encoding;
};
typedef struct b64_state_s b64_state_t;

size_t qp_decode (char *buffer, size_t length, int *r_slbrk);
char *qp_encode (const char *input, size_t length, size_t* outlen);
void b64_init (b64_state_t *state);
size_t b64_decode (b64_state_t *state, char *buffer, size_t length);
char * b64_encode (const char *input, size_t length);

char *latin1_to_utf8 (const char *string);

char *mem2str (char *dest, const void *src, size_t n);

char *trim_spaces (char *string);
char *trim_trailing_spaces (char *string);

/* To avoid that a compiler optimizes certain memset calls away, these
   macros may be used instead. */
#define wipememory2(_ptr,_set,_len) do { \
              volatile char *_vptr=(volatile char *)(_ptr); \
              size_t _vlen=(_len); \
              while(_vlen) { *_vptr=(_set); _vptr++; _vlen--; } \
                  } while(0)
#define wipememory(_ptr,_len) wipememory2(_ptr,0,_len)
#define wipestring(_ptr) do { \
              volatile char *_vptr=(volatile char *)(_ptr); \
              while(*_vptr) { *_vptr=0; _vptr++; } \
                  } while(0)

void set_default_key (const char *name);

/*-- Convenience macros. -- */
#define DIM(v)		     (sizeof(v)/sizeof((v)[0]))
#define DIMof(type,member)   DIM(((type *)0)->member)

/*-- Macros to replace ctype ones to avoid locale problems. --*/
#define spacep(p)   (*(p) == ' ' || *(p) == '\t')
#define digitp(p)   (*(p) >= '0' && *(p) <= '9')
#define hexdigitp(a) (digitp (a)                     \
                      || (*(a) >= 'A' && *(a) <= 'F')  \
                      || (*(a) >= 'a' && *(a) <= 'f'))
  /* Note this isn't identical to a C locale isspace() without \f and
     \v, but works for the purposes used here. */
#define ascii_isspace(a) ((a)==' ' || (a)=='\n' || (a)=='\r' || (a)=='\t')

/* The atoi macros assume that the buffer has only valid digits. */
#define atoi_1(p)   (*(p) - '0' )
#define atoi_2(p)   ((atoi_1(p) * 10) + atoi_1((p)+1))
#define atoi_4(p)   ((atoi_2(p) * 100) + atoi_2((p)+2))
#define xtoi_1(p)   (*(p) <= '9'? (*(p)- '0'): \
                     *(p) <= 'F'? (*(p)-'A'+10):(*(p)-'a'+10))
#define xtoi_2(p)   ((xtoi_1(p) * 16) + xtoi_1((p)+1))
#define xtoi_4(p)   ((xtoi_2(p) * 256) + xtoi_2((p)+2))

#define tohex(n) ((n) < 10 ? ((n) + '0') : (((n) - 10) + 'A'))

#define tohex_lower(n) ((n) < 10 ? ((n) + '0') : (((n) - 10) + 'a'))
/***** Inline functions.  ****/

/* Return true if LINE consists only of white space (up to and
   including the LF). */
static inline int
trailing_ws_p (const char *line)
{
  for ( ; *line && *line != '\n'; line++)
    if (*line != ' ' && *line != '\t' && *line != '\r')
      return 0;
  return 1;
}

/* An strcmp variant with the compare ending at the end of B.  */
static inline int
tagcmp (const char *a, const char *b)
{
  return strncmp (a, b, strlen (b));
}

#ifdef HAVE_W32_SYSTEM
extern HANDLE log_mutex;
#endif
/*****  Missing functions.  ****/

#ifndef HAVE_STPCPY
static inline char *
_gpgol_stpcpy (char *a, const char *b)
{
  while (*b)
    *a++ = *b++;
  *a = 0;
  return a;
}
#define stpcpy(a,b) _gpgol_stpcpy ((a), (b))
#endif /*!HAVE_STPCPY*/

/* The length of the boundary - the buffer needs to be allocated one
   byte larger. */
#define BOUNDARYSIZE 20
char *generate_boundary (char *buffer);

#ifdef __cplusplus
}
#endif

#endif // COMMON_INDEP_H
