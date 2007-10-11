/* display.h - Helper functions for displaying messages.
 *	Copyright (C) 2005 g10 Code GmbH
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

#ifndef DISPLAY_H
#define DISPLAY_H

int is_html_body (const char *body);

char *add_html_line_endings (const char *body);

int update_display (HWND hwnd, void *exchange_cb,
                    bool is_html, const char *text);

int set_message_body (LPMESSAGE message, const char *string, bool is_html);

int open_inspector (LPEXCHEXTCALLBACK peecb, LPMESSAGE message);


#endif /*DISPLAY_H*/
