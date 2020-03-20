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
#include <map>

#include <gpgme++/global.h>

/* Stuff that should be in common but is c++ so it does not fit in there. */


/* Release a null terminated char* array */
void release_cArray (char **carray);

/* Trim whitespace from a string. */
void rtrim (std::string &s);
void ltrim (std::string &s);
void trim (std::string &s);
void remove_whitespace (std::string &s);

/* Join a string vector */
void join(const std::vector<std::string>& v, const char *c, std::string& s);

/* Convert a string vector to a null terminated char array */
char **vector_to_cArray (const std::vector<std::string> &vec);
std::vector <std::string> cArray_to_vector (const char **cArray);

/* More string helpers */
bool starts_with(const std::string &s, const char *prefix);
bool starts_with(const std::string &s, const char prefix);

/* Check if we are in de_vs mode. */
bool in_de_vs_mode ();

#ifdef HAVE_W32_SYSTEM
/* Get a map of all subkey value pairs in a registry key */
std::map<std::string, std::string> get_registry_subkeys (const char *path);
#endif

std::vector<std::string> gpgol_split (const std::string &s, char delim);

/* Convert a string to a hex representation */
std::string string_to_hex (const std::string& input);

/* Check if a string contains a char < 32 */
bool is_binary (const std::string &input);

/* Return a string repr of the GpgME Protocol */
const char *to_cstr (const GpgME::Protocol &prot);

/* Modify source and find and replace stuff */
void find_and_replace(std::string& source, const std::string &find,
                      const std::string &replace);
#endif // CPPHELP_H
