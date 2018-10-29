/* @file mlang-charset.cpp
 * @brief Convert between charsets using Mlang
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

#include "config.h"
#include "common.h"
#define INITGUID
#include <initguid.h>
DEFINE_GUID (IID_IMultiLanguage, 0x275c23e1,0x3747,0x11d0,0x9f,
                                 0xea,0x00,0xaa,0x00,0x3f,0x86,0x46);
#include <mlang.h>
#undef INITGUID

#include "dialogs.h"
#include "dispcache.h"

#include "mlang-charset.h"

static char *
iconv_to_utf8 (const char *charset, const char *input)
{
  if (!charset || !input)
    {
      STRANGEPOINT;
      return nullptr;
    }

  gpgrt_w32_iconv_t ctx = gpgrt_w32_iconv_open ("UTF-8", charset);
  if (!ctx || ctx == (gpgrt_w32_iconv_t)-1)
    {
      log_debug ("%s:%s: Failed to open iconv ctx for '%s'",
                 SRCNAME, __func__, charset);
      return nullptr;
    }

  size_t len = 0;

  for (const unsigned char *s = (const unsigned char*) input; *s; s++)
    {
      len++;
      if ((*s & 0x80))
        {
          len += 5; /* We may need up to 6 bytes for the utf8 output. */
        }
    }

  char *buffer = (char*) xmalloc (len + 1);
  size_t inlen = strlen (input) + 1; // Need to add 1 for the zero
  char *outptr = buffer;
  size_t outbytes = len;
  size_t ret = gpgrt_w32_iconv (ctx, (const char **)&input, &inlen,
                                &outptr, &outbytes);
  gpgrt_w32_iconv_close (ctx);
  if (ret == -1)
    {
      log_error ("%s:%s: Conversion failed for '%s'",
                 SRCNAME, __func__, charset);
      xfree (buffer);
      return nullptr;
    }
  return buffer;
}

char *ansi_charset_to_utf8 (const char *charset, const char *input,
                            size_t inlen, int codepage)
{
  LPMULTILANGUAGE multilang = NULL;
  MIMECSETINFO mime_info;
  HRESULT err;
  DWORD enc;
  DWORD mode = 0;
  unsigned int wlen = 0,
               uinlen = 0;
  wchar_t *buf;
  char *ret;

  if ((!charset || !strlen (charset)) && !codepage)
    {
      log_debug ("%s:%s: No charset / codepage returning plain.",
                 SRCNAME, __func__);
      return xstrdup (input);
    }

  auto cache = DispCache::instance ();
  LPDISPATCH cachedLang = cache->getDisp (DISPID_MLANG_CHARSET);

  if (!cachedLang)
    {
      CoCreateInstance(CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER,
                       IID_IMultiLanguage, (void**)&multilang);
      memdbg_addRef (multilang);
      cache->addDisp (DISPID_MLANG_CHARSET, (LPDISPATCH) multilang);
    }
  else
    {
      multilang = (LPMULTILANGUAGE) cachedLang;
    }


  if (!multilang)
    {
      log_error ("%s:%s: Failed to get multilang obj.",
                 SRCNAME, __func__);
      return NULL;
    }

  if (inlen > UINT_MAX)
    {
      log_error ("%s:%s: Inlen too long. Bug.",
                 SRCNAME, __func__);
      return NULL;
    }

  uinlen = (unsigned int) inlen;

  if (!codepage)
    {
      mime_info.uiCodePage = 0;
      mime_info.uiInternetEncoding = 0;
      BSTR w_charset = utf8_to_wchar (charset);
      err = multilang->GetCharsetInfo (w_charset, &mime_info);
      xfree (w_charset);
      if (err != S_OK)
        {
          log_debug ("%s:%s: Failed to find charset for: %s fallback to iconv",
                     SRCNAME, __func__, charset);
          /* We only use this as a fallback as the old code was older and
             known to work in most cases. */
          ret = iconv_to_utf8 (charset, input);
          if (ret)
            {
              return ret;
            }

          return xstrdup (input);
        }
      enc = (mime_info.uiInternetEncoding == 0) ? mime_info.uiCodePage :
                                                  mime_info.uiInternetEncoding;
    }
  else
    {
      enc = codepage;
    }

  /** Get the size of the result */
  err = multilang->ConvertStringToUnicode(&mode, enc, const_cast<char*>(input),
                                          &uinlen, NULL, &wlen);
  if (FAILED (err))
    {
      log_error ("%s:%s: Failed conversion.",
                 SRCNAME, __func__);
      return NULL;
  }
  buf = (wchar_t*) xmalloc(sizeof(wchar_t) * (wlen + 1));

  err = multilang->ConvertStringToUnicode(&mode, enc, const_cast<char*>(input),
                                          &uinlen, buf, &wlen);
  if (FAILED (err))
    {
      log_error ("%s:%s: Failed conversion 2.",
                 SRCNAME, __func__);
      xfree (buf);
      return NULL;
    }
  /* Doc is not clear if this is terminated. */
  buf[wlen] = L'\0';

  ret = wchar_to_utf8 (buf);
  xfree (buf);
  return ret;
}
