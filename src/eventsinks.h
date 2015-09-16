/* eventsinks.h - Declaraion of eventsink installation functions.
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
#ifndef EVENTSINKS_H
#define EVENTSINKS_H

#include <windows.h>

LPDISPATCH install_ApplicationEvents_sink (LPDISPATCH obj);
void detach_ApplicationEvents_sink (LPDISPATCH obj);
LPDISPATCH install_MailItemEvents_sink (LPDISPATCH obj);
void detach_MailItemEvents_sink (LPDISPATCH obj);
#endif // EVENTSINKS_H
