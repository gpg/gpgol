/* display.h - Helper functions for displaying messages.
 *	Copyright (C) 2005 g10 Code GmbH
 *
 * This file is part of GPGol.
 * 
 * GPGol is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GPGol is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifndef DISPLAY_H
#define DISPLAY_H

#include "gpgmsg.hh"

int is_html_body (const char *body);

char *add_html_line_endings (const char *body);

int update_display (HWND hwnd, GpgMsg *msg, void *exchange_cb,
                    bool is_html, const char *text);

int set_message_body (LPMESSAGE message, const char *string, bool is_html);


/*-- olflange.cpp --*/
int put_outlook_property (void *pEECB, const char *key, const char *value);
int put_outlook_property_int (void *pEECB, const char *key, int value);
char *get_outlook_property (void *pEECB, const char *key);


#endif /*DISPLAY_H*/
