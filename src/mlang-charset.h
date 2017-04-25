/* @file mlang-charset.h
 * @brief Convert between charsets using Mlang
 *
 *    Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 *    Software engineering by Intevation GmbH
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

#include "common.h"
#ifdef __cplusplus
extern "C" {
#if 0
}
#endif
#endif

/** @brief convert input to utf8.
  *
  * @param charset: ANSI name of the charset to decode.
  * @param input: The input to convert.
  * @param inlen: The size of the input.
  *
  * @returns NULL on error or an UTF-8 encoded NULL terminated string.
  */

char *ansi_charset_to_utf8 (const char *charset, const char *input,
                            size_t inlen);
#ifdef __cplusplus
}
#endif
