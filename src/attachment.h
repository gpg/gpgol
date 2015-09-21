/* attachment.h - Functions for attachment handling
 *    Copyright (C) 2005, 2007 g10 Code GmbH
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
#ifndef ATTACHMENT_H
#define ATTACHMENT_H

#include <windows.h>

/** Protect attachments so that it can be stored
  by outlook. This means to symetrically encrypt the
  data with the session key.

  This will change the messagetype back to
  ATTACHTYPE_FROMMOSS it is only supposed to be
  called on attachments with the Attachmentype
  ATTACHTYPE_FROMMOSS_DEC.

  The dispatch paramenter should be a mailitem.

  Returns 0 on success.
*/
int
protect_attachments (LPDISPATCH mailitem);

/** Remove the symetric session encryption of the attachments.

  The dispatch paramenter should be a mailitem.

  This will change the messsagetype to
  ATTACHTYPE_FROMMOSS_DEC it should only be called
  with attachments of the type ATTACHTYPE_FROMMOSS.

  Returns 0 on success. */
int
unprotect_attachments (LPDISPATCH mailitem);
#endif // ATTACHMENT_H
