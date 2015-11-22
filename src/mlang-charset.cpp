/* @file mlang-charset.cpp
 * @brief Convert between charsets using Mlang
 *
 *    Copyright (C) 2015 Intevation GmbH
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

#include "util.h"
#define INITGUID
#include <initguid.h>
DEFINE_GUID (IID_IMultiLanguage, 0x275c23e1,0x3747,0x11d0,0x9f,
                                 0xea,0x00,0xaa,0x00,0x3f,0x86,0x46);
#include <mlang.h>
#undef INITGUID

#include "mlang-charset.h"

char *ansi_charset_to_utf8 (const char *charset, char *input,
                            size_t inlen)
{
  LPMULTILANGUAGE multilang = NULL;
  MIMECSETINFO mime_info;
  HRESULT err;
  DWORD enc;
  DWORD mode = 0;
  unsigned int wlen = 0;
  wchar_t *buf;
  char *ret;

  CoCreateInstance(CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER,
                   IID_IMultiLanguage, (void**)&multilang);

  if (!multilang)
    {
      log_error ("%s:%s: Failed to get multilang obj.",
                 SRCNAME, __func__);
      return NULL;
    }


  mime_info.uiCodePage = 0;
  mime_info.uiInternetEncoding = 0;
  BSTR w_charset = utf8_to_wchar (charset);
  err = multilang->GetCharsetInfo (w_charset, &mime_info);
  xfree (w_charset);
  if (err != S_OK)
    {
      log_error ("%s:%s: Failed to find charset for: %s",
                 SRCNAME, __func__, charset);
      multilang->Release ();
      return NULL;
    }
  enc = (mime_info.uiInternetEncoding == 0) ? mime_info.uiCodePage :
                                              mime_info.uiInternetEncoding;

  /** Get the size of the result */
  err = multilang->ConvertStringToUnicode(&mode, enc, input,
                                          &inlen, NULL, &wlen);
  if (FAILED (err))
    {
      log_error ("%s:%s: Failed conversion.",
                 SRCNAME, __func__);
      multilang->Release ();
      return NULL;
  }
  buf = (wchar_t*) xmalloc(sizeof(wchar_t) * (wlen + 1));

  err = multilang->ConvertStringToUnicode(&mode, enc, input, &inlen,
                                          buf, &wlen);
  multilang->Release ();
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
