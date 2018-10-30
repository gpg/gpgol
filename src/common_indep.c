/* common_indep.c - Common, platform indepentent routines used by GpgOL
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

#include "common_indep.h"
#ifdef HAVE_W32_SYSTEM
#include <windows.h>
#endif

#include <wchar.h>
#include <stdlib.h>
#include <ctype.h>

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
out_of_core (void)
{
#ifdef HAVE_W32_SYSTEM
  MessageBox (NULL, "Out of core!", "Fatal Error", MB_OK);
#endif
  abort ();
}

void*
_xmalloc (size_t n)
{
  void *p = malloc (n);
  if (!p)
    out_of_core ();
  return p;
}

void*
_xcalloc (size_t m, size_t n)
{
  void *p = calloc (m, n);
  if (!p)
    out_of_core ();
  return p;
}

void *
_xrealloc (void *a, size_t n)
{
  void *p = realloc (a, n);
  if (!p)
    out_of_core ();
  return p;
}

char*
_xstrdup (const char *s)
{
  char *p = _xmalloc (strlen (s)+1);
  strcpy (p, s);
  return p;
}

wchar_t*
_xwcsdup (const wchar_t *s)
{
  wchar_t *p = wcsdup (s);
  if (!s)
    {
      out_of_core ();
    }
  return p;
}

void
_xfree (void *p)
{
  if (p)
    free (p);
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

/* Return the a quoted printable encoded version of the
   input string. If outlen is not null the size of the
   quoted printable string is returned. String will be
   malloced and zero terminated. Aborts if the output
   is more then three times the size of the input.
   This is only basic and does not handle mutliline data. */
char *
qp_encode (const char *input, size_t inlen, size_t *r_outlen)
{
  size_t max_len = inlen * 3 +1;
  char *outbuf = xmalloc (max_len);
  size_t outlen = 0;
  const unsigned char *p;

  memset (outbuf, 0, max_len);

  for (p = input; inlen; p++, inlen--)
    {
      if (*p >= '!' && *p <= '~' && *p != '=')
        {
          outbuf[outlen++] = *p;
        }
      else if (*p == ' ')
        {
          /* Outlook does it this way */
          outbuf[outlen++] = '_';
        }
      else
        {
          outbuf[outlen++] = '=';
          outbuf[outlen++] = tohex ((*p>>4)&15);
          outbuf[outlen++] = tohex (*p&15);
        }
      if (outlen == max_len -1)
        {
          log_error ("Quoted printable too long. Bug.");
          r_outlen = NULL;
          xfree (outbuf);
          return NULL;
        }
    }
  if (r_outlen)
    *r_outlen = outlen;
  return outbuf;
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


/* Base 64 encode the input. If input is null returns NULL otherwise
   a pointer to the malloced encoded string. */
char *
b64_encode (const char *input, size_t length)
{
  size_t out_len = 4 * ((length + 2) / 3);
  char *ret;
  int i, j;

  if (!length || !input)
    {
      return NULL;
    }
  ret = xmalloc (out_len);
  memset (ret, 0, out_len);

  for (i = 0, j = 0; i < length;)
    {
      unsigned int a = i < length ? (unsigned char)input[i++] : 0;
      unsigned int b = i < length ? (unsigned char)input[i++] : 0;
      unsigned int c = i < length ? (unsigned char)input[i++] : 0;

      unsigned int triple = (a << 0x10) + (b << 0x08) + c;

      ret[j++] = bintoasc[(triple >> 3 * 6) & 0x3F];
      ret[j++] = bintoasc[(triple >> 2 * 6) & 0x3F];
      ret[j++] = bintoasc[(triple >> 1 * 6) & 0x3F];
      ret[j++] = bintoasc[(triple >> 0 * 6) & 0x3F];
    }

  if (length % 3)
    {
      ret [j - 1] = '=';
    }
  if (length % 3 == 1)
    {
      ret [j - 2] = '=';
    }

  return ret;
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
