/* message.h - Declarations for message.c
 *	Copyright (C) 2007 g10 Code GmbH
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

#ifndef MESSAGE_H
#define MESSAGE_H


bool message_incoming_handler (LPMESSAGE message, HWND hwnd);
bool message_display_handler (LPEXCHEXTCALLBACK eecb, HWND hwnd);
void message_wipe_body_cruft (LPEXCHEXTCALLBACK eecb);
void message_show_info (LPMESSAGE message, HWND hwnd);


int message_verify (LPMESSAGE message, msgtype_t msgtype, int force,
                    HWND hwnd);
int message_decrypt (LPMESSAGE message, msgtype_t msgtype, int force, 
                     HWND hwnd);
int message_sign (LPMESSAGE message, protocol_t protocol, HWND hwnd);
int message_encrypt (LPMESSAGE message, protocol_t protocol, HWND hwnd);
int message_sign_encrypt (LPMESSAGE message, protocol_t protocol, HWND hwnd);


#endif /*MESSAGE_H*/