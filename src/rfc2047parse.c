/* @file rfc2047parse.c
 * @brief Parsercode for rfc2047
 *
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
 *
 * This file is part of GpgOL.
 *
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program; if not, see <http://www.gnu.org/licenses/>.
 */

/* This code is heavily based (mostly verbatim copy with glib
 *  dependencies removed) on GMime rev 496313fb
 * modified by aheinecke@intevation.de
 *
 * Copyright (C) 2000-2014 Jeffrey Stedfast
 *
 *  This library is free software; you can redistribute it and/or
 *  modify it under the terms of the GNU Lesser General Public License
 *  as published by the Free Software Foundation; either version 2.1
 *  of the License, or (at your option) any later version.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <stdbool.h>
#include "common_indep.h"
#include <ctype.h>

#ifdef HAVE_W32_SYSTEM
# include "mlang-charset.h"
#endif

#include "gmime-table-private.h"

/* mabye we need this at some point later? */
#define G_MIME_RFC2047_WORKAROUNDS 1


static unsigned char gmime_base64_rank[256] = {
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255, 62,255,255,255, 63,
    52, 53, 54, 55, 56, 57, 58, 59, 60, 61,255,255,255,  0,255,255,
    255,  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13, 14,
    15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,255,255,255,255,255,
    255, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
    41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
    255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,
};

typedef struct _rfc2047_token {
    struct _rfc2047_token *next;
    char *charset;
    const char *text;
    size_t length;
    char encoding;
    char is_8bit;
} rfc2047_token;

static rfc2047_token *
rfc2047_token_new (const char *text, size_t len)
{
  rfc2047_token *token;

  token = xmalloc (sizeof (rfc2047_token));
  memset (token, 0, sizeof (rfc2047_token));
  token->length = len;
  token->text = text;

  return token;
}

static rfc2047_token *
rfc2047_token_new_encoded_word (const char *word, size_t len)
{
  rfc2047_token *token;
  const char *payload;
  char *charset;
  const char *inptr;
  const char *tmpchar;
  char *buf, *lang;
  char encoding;
  size_t n;

  /* check that this could even be an encoded-word token */
  if (len < 7 || strncmp (word, "=?", 2) != 0 || strncmp (word + len - 2, "?=", 2) != 0)
    return NULL;

  /* skip over '=?' */
  inptr = word + 2;
  tmpchar = inptr;

  if (*tmpchar == '?' || *tmpchar == '*') {
      /* this would result in an empty charset */
      return NULL;
  }

  /* skip to the end of the charset */
  if (!(inptr = memchr (inptr, '?', len - 2)) || inptr[2] != '?')
    return NULL;

  /* copy the charset into a buffer */
  n = (size_t) (inptr - tmpchar);
  buf = malloc (n + 1);
  memcpy (buf, tmpchar, n);
  buf[n] = '\0';
  charset = buf;

  /* rfc2231 updates rfc2047 encoded words...
   * The ABNF given in RFC 2047 for encoded-words is:
   *   encoded-word := "=?" charset "?" encoding "?" encoded-text "?="
   * This specification changes this ABNF to:
   *   encoded-word := "=?" charset ["*" language] "?" encoding "?" encoded-text "?="
   */

  /* trim off the 'language' part if it's there... */
  if ((lang = strchr (charset, '*')))
    *lang = '\0';

  /* skip over the '?' */
  inptr++;

  /* make sure the first char after the encoding is another '?' */
  if (inptr[1] != '?')
    return NULL;

  switch (*inptr++) {
    case 'B': case 'b':
      encoding = 'B';
      break;
    case 'Q': case 'q':
      encoding = 'Q';
      break;
    default:
      return NULL;
  }

  /* the payload begins right after the '?' */
  payload = inptr + 1;

  /* find the end of the payload */
  inptr = word + len - 2;

  /* make sure that we don't have something like: =?iso-8859-1?Q?= */
  if (payload > inptr)
    return NULL;

  token = rfc2047_token_new (payload, inptr - payload);
  token->charset = charset;
  token->encoding = encoding;

  return token;
}

static void
rfc2047_token_free (rfc2047_token * tok)
{
  if (!tok)
    {
      return;
    }
  xfree (tok->charset);
  xfree (tok);
}

static rfc2047_token *
tokenize_rfc2047_phrase (const char *in, size_t *len)
{
  bool enable_rfc2047_workarounds = G_MIME_RFC2047_WORKAROUNDS;
  rfc2047_token list, *lwsp, *token, *tail;
  register const char *inptr = in;
  bool encoded = false;
  const char *text, *word;
  bool ascii;
  size_t n;

  tail = (rfc2047_token *) &list;
  list.next = NULL;
  lwsp = NULL;

  while (*inptr != '\0') {
      text = inptr;
      while (is_lwsp (*inptr))
        inptr++;

      if (inptr > text)
        lwsp = rfc2047_token_new (text, inptr - text);
      else
        lwsp = NULL;

      word = inptr;
      ascii = true;
      if (is_atom (*inptr)) {
          if (enable_rfc2047_workarounds) {
              /* Make an extra effort to detect and
               * separate encoded-word tokens that
               * have been merged with other
               * words. */

              if (!strncmp (inptr, "=?", 2)) {
                  inptr += 2;

                  /* skip past the charset (if one is even declared, sigh) */
                  while (*inptr && *inptr != '?') {
                      ascii = ascii && is_ascii (*inptr);
                      inptr++;
                  }

                  /* sanity check encoding type */
                  if (inptr[0] != '?' || !strchr ("BbQq", inptr[1]) || inptr[2] != '?')
                    goto non_rfc2047;

                  inptr += 3;

                  /* find the end of the rfc2047 encoded word token */
                  while (*inptr && strncmp (inptr, "?=", 2) != 0) {
                      ascii = ascii && is_ascii (*inptr);
                      inptr++;
                  }

                  if (*inptr == '\0') {
                      /* didn't find an end marker... */
                      inptr = word + 2;
                      ascii = true;

                      goto non_rfc2047;
                  }

                  inptr += 2;
              } else {
non_rfc2047:
                  /* stop if we encounter a possible rfc2047 encoded
                   * token even if it's inside another word, sigh. */
                  while (is_atom (*inptr) && strncmp (inptr, "=?", 2) != 0)
                    inptr++;
              }
          } else {
              while (is_atom (*inptr))
                inptr++;
          }

          n = (size_t) (inptr - word);
          if ((token = rfc2047_token_new_encoded_word (word, n))) {
              /* rfc2047 states that you must ignore all
               * whitespace between encoded words */
              if (!encoded && lwsp != NULL) {
                  tail->next = lwsp;
                  tail = lwsp;
              } else if (lwsp != NULL) {
                  rfc2047_token_free (lwsp);
              }

              tail->next = token;
              tail = token;

              encoded = true;
          } else {
              /* append the lwsp and atom tokens */
              if (lwsp != NULL) {
                  tail->next = lwsp;
                  tail = lwsp;
              }

              token = rfc2047_token_new (word, n);
              token->is_8bit = ascii ? 0 : 1;
              tail->next = token;
              tail = token;

              encoded = false;
          }
      } else {
          /* append the lwsp token */
          if (lwsp != NULL) {
              tail->next = lwsp;
              tail = lwsp;
          }

          ascii = true;
          while (*inptr && !is_lwsp (*inptr) && !is_atom (*inptr)) {
              ascii = ascii && is_ascii (*inptr);
              inptr++;
          }

          n = (size_t) (inptr - word);
          token = rfc2047_token_new (word, n);
          token->is_8bit = ascii ? 0 : 1;

          tail->next = token;
          tail = token;

          encoded = false;
      }
  }

  *len = (size_t) (inptr - in);

  return list.next;
}

static void
rfc2047_token_list_free (rfc2047_token * tokens)
{
  rfc2047_token * cur = tokens;
  while (cur)
    {
      rfc2047_token *next = cur->next;
      rfc2047_token_free (cur);
      cur = next;
    }
}

/* this decodes rfc2047's version of quoted-printable */
static size_t
quoted_decode (const unsigned char *in, size_t len, unsigned char *out, int *state, unsigned int *save)
{
  register const unsigned char *inptr;
  register unsigned char *outptr;
  const unsigned char *inend;
  unsigned char c, c1;
  unsigned int saved;
  int need;

  if (len == 0)
    return 0;

  inend = in + len;
  outptr = out;
  inptr = in;

  need = *state;
  saved = *save;

  if (need > 0) {
      if (isxdigit ((int) *inptr)) {
          if (need == 1) {
              c = toupper ((int) (saved & 0xff));
              c1 = toupper ((int) *inptr++);
              saved = 0;
              need = 0;

              goto decode;
          }

          saved = 0;
          need = 0;

          goto equals;
      }

      /* last encoded-word ended in a malformed quoted-printable sequence */
      *outptr++ = '=';

      if (need == 1)
        *outptr++ = (char) (saved & 0xff);

      saved = 0;
      need = 0;
  }

  while (inptr < inend) {
      c = *inptr++;
      if (c == '=') {
equals:
          if (inend - inptr >= 2) {
              if (isxdigit ((int) inptr[0]) && isxdigit ((int) inptr[1])) {
                  c = toupper (*inptr++);
                  c1 = toupper (*inptr++);
decode:
                  *outptr++ = (((c >= 'A' ? c - 'A' + 10 : c - '0') & 0x0f) << 4)
                    | ((c1 >= 'A' ? c1 - 'A' + 10 : c1 - '0') & 0x0f);
              } else {
                  /* malformed quoted-printable sequence? */
                  *outptr++ = '=';
              }
          } else {
              /* truncated payload, maybe it was split across encoded-words? */
              if (inptr < inend) {
                  if (isxdigit ((int) *inptr)) {
                      saved = *inptr;
                      need = 1;
                      break;
                  } else {
                      /* malformed quoted-printable sequence? */
                      *outptr++ = '=';
                  }
              } else {
                  saved = 0;
                  need = 2;
                  break;
              }
          }
      } else if (c == '_') {
          /* _'s are an rfc2047 shortcut for encoding spaces */
          *outptr++ = ' ';
      } else {
          *outptr++ = c;
      }
  }

  *state = need;
  *save = saved;

  return (size_t) (outptr - out);
}

/**
 * g_mime_encoding_base64_decode_step:
 * @inbuf: input buffer
 * @inlen: input buffer length
 * @outbuf: output buffer
 * @state: holds the number of bits that are stored in @save
 * @save: leftover bits that have not yet been decoded
 *
 * Decodes a chunk of base64 encoded data.
 *
 * Returns: the number of bytes decoded (which have been dumped in
 * @outbuf).
 **/
size_t
g_mime_encoding_base64_decode_step (const unsigned char *inbuf, size_t inlen, unsigned char *outbuf, int *state, unsigned int *save)
{
  register const unsigned char *inptr;
  register unsigned char *outptr;
  const unsigned char *inend;
  register unsigned int saved;
  unsigned char c;
  int npad, n, i;

  inend = inbuf + inlen;
  outptr = outbuf;
  inptr = inbuf;

  npad = (*state >> 8) & 0xff;
  n = *state & 0xff;
  saved = *save;

  /* convert 4 base64 bytes to 3 normal bytes */
  while (inptr < inend) {
      c = gmime_base64_rank[*inptr++];
      if (c != 0xff) {
          saved = (saved << 6) | c;
          n++;
          if (n == 4) {
              *outptr++ = saved >> 16;
              *outptr++ = saved >> 8;
              *outptr++ = saved;
              n = 0;

              if (npad > 0) {
                  outptr -= npad;
                  npad = 0;
              }
          }
      }
  }

  /* quickly scan back for '=' on the end somewhere */
  /* fortunately we can drop 1 output char for each trailing '=' (up to 2) */
  for (i = 2; inptr > inbuf && i; ) {
      inptr--;
      if (gmime_base64_rank[*inptr] != 0xff) {
          if (*inptr == '=' && outptr > outbuf) {
              if (n == 0) {
                  /* we've got a complete quartet so it's
                     safe to drop an output character. */
                  outptr--;
              } else if (npad < 2) {
                  /* keep a record of the number of ='s at
                     the end of the input stream, up to 2 */
                  npad++;
              }
          }

          i--;
      }
  }

  *state = (npad << 8) | n;
  *save = n ? saved : 0;

  return (outptr - outbuf);
}

static size_t
rfc2047_token_decode (rfc2047_token *token, unsigned char *outbuf, int *state, unsigned int *save)
{
  const unsigned char *inbuf = (const unsigned char *) token->text;
  size_t len = token->length;

  if (token->encoding == 'B')
    return g_mime_encoding_base64_decode_step (inbuf, len, outbuf, state, save);
  else
    return quoted_decode (inbuf, len, outbuf, state, save);
}

static char *
rfc2047_decode_tokens (rfc2047_token *tokens, size_t buflen)
{
  rfc2047_token *token, *next;
  size_t outlen, len, tmplen;
  unsigned char *outptr;
  const char *charset;
  char *outbuf;
  char *decoded;
  char encoding;
  unsigned int save;
  int state;
  char *str;

  decoded = xmalloc (buflen + 1);
  memset (decoded, 0, buflen + 1);
  tmplen = 76;
  outbuf = xmalloc (tmplen);

  token = tokens;
  while (token != NULL) {
      next = token->next;

      if (token->encoding) {
          /* In order to work around broken mailers, we need to combine
           * the raw decoded content of runs of identically encoded word
           * tokens before converting into UTF-8. */
          encoding = token->encoding;
          charset = token->charset;
          len = token->length;
          state = 0;
          save = 0;

          /* find the end of the run (and measure the buffer length we'll need) */
          while (next && next->encoding == encoding && !strcmp (next->charset, charset)) {
              len += next->length;
              next = next->next;
          }

          /* make sure our temporary output buffer is large enough... */
          if (len > tmplen)
            {
              xrealloc (outbuf, len + 1);
              tmplen = len + 1;
            }

          /* base64 / quoted-printable decode each of the tokens... */
          outptr = outbuf;
          outlen = 0;
          do {
              /* Note: by not resetting state/save each loop, we effectively
               * treat the payloads as one continuous block, thus allowing
               * us to handle cases where a hex-encoded triplet of a
               * quoted-printable encoded payload is split between 2 or more
               * encoded-word tokens. */
              len = rfc2047_token_decode (token, outptr, &state, &save);
              token = token->next;
              outptr += len;
              outlen += len;
          } while (token != next);
          outptr = outbuf;

          /* convert the raw decoded text into UTF-8 */
          if (!strcasecmp (charset, "UTF-8")) {
              strncat (decoded, (char *) outptr, outlen);
          } else {
#ifdef HAVE_W32_SYSTEM
              str = ansi_charset_to_utf8 (charset, outptr, outlen, 0);
#else
              log_debug ("%s:%s: Conversion not available on non W32 systems",
                         SRCNAME, __func__);
              str = strndup (outptr, outlen);
#endif
              if (!str)
                {
                  log_error ("%s:%s: Failed conversion from: %s for word: %s.",
                             SRCNAME, __func__, charset, outptr);
                }
              else
                {
                  strcat (decoded, str);
                  xfree (str);
                }
          }
      } else {
          strncat (decoded, token->text, token->length);
      }
      if (token && token->is_8bit)
      {
        /* We don't support this. */
        log_error ("%s:%s: Unknown 8bit encoding detected.",
                   SRCNAME, __func__);
      }

      token = next;
  }

  xfree (outbuf);

  return decoded;
}


/**
 * g_mime_utils_header_decode_phrase:
 * @phrase: header to decode
 *
 * Decodes an rfc2047 encoded 'phrase' header.
 *
 * Note: See g_mime_set_user_charsets() for details on how charset
 * conversion is handled for unencoded 8bit text and/or wrongly
 * specified rfc2047 encoded-word tokens.
 *
 * Returns: a newly allocated UTF-8 string representing the the decoded
 * header.
 **/
static char *
g_mime_utils_header_decode_phrase (const char *phrase)
{
  rfc2047_token *tokens;
  char *decoded;
  size_t len;

  tokens = tokenize_rfc2047_phrase (phrase, &len);
  decoded = rfc2047_decode_tokens (tokens, len);
  rfc2047_token_list_free (tokens);

  return decoded;
}

/* Try to parse an rfc 2047 filename for attachment handling.
   returns the parsed string. On errors the input string is just
   copied with strdup */
char *
rfc2047_parse (const char *input)
{
  char *decoded;
  if (!input)
    return strdup ("");

  log_debug ("%s:%s: Input: \"%s\"",
             SRCNAME, __func__, input);

  decoded = g_mime_utils_header_decode_phrase (input);

  log_debug ("%s:%s: Decoded: \"%s\"",
             SRCNAME, __func__, decoded);

  if (!decoded || !strlen (decoded))
    {
      xfree (decoded);
      return strdup (input);
    }
  return decoded;
}
