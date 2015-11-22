/* @file rfc2047parse.h
 * @brief Parser for filenames encoded according to rfc2047
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

/** @brief Try to parse a string according to rfc2047.
  *
  * On error the error is logged and a copy of the original
  * input string returned.
  *
  * @returns a malloced string in UTF-8 encoding or a copy
  *          of the input string.
  */
char *
rfc2047_parse (const char *input);
