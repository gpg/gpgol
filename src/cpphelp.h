#ifndef CPPHELP_H
#define CPPHELP_H
/* @file cpphelp.h
 * @brief Common cpp helper stuff
 *
 * Copyright (C) 2018 Intevation GmbH
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

#include <string>
#include <vector>

/* Stuff that should be in common but is c++ so it does not fit in there. */


/* Release a null terminated char* array */
void release_cArray (char **carray);

/* Trim whitespace from a string. */
void rtrim(std::string &s);

/* Convert a string vector to a null terminated char array */
char **vector_to_cArray (const std::vector<std::string> &vec);

#endif // CPPHELP_H
