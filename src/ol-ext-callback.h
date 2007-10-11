/* ol-ext-callback.h - Definitions for ol-ext-callback.cpp
 *	Copyright (C) 2005, 2007 g10 Code GmbH
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

#ifndef OL_EXT_CALLBACK_H
#define OL_EXT_CALLBACK_H


LPDISPATCH find_outlook_property (LPEXCHEXTCALLBACK lpeecb,
                                  const char *name, DISPID *r_dispid);
int put_outlook_property (void *pEECB, const char *key, const char *value);
int put_outlook_property_int (void *pEECB, const char *key, int value);
char *get_outlook_property (void *pEECB, const char *key);


#endif /*OL_EXT_CALLBACK_H*/
