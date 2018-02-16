#ifndef OVERLAY_H
#define OVERLAY_H
/* @file overlay.h
 * @brief Overlay something through WinAPI.
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
#include <windows.h>

#include <string>

#include <gpgme++/data.h>

namespace GpgME
{
  class Context;
} // namespace GpgME

class Overlay
{
public:
  /* Create an overlay over a foreign window */
  Overlay(HWND handle, const std::string &text);
  ~Overlay();

private:
  std::unique_ptr<GpgME::Context> m_overlayCtx;
  GpgME::Data m_overlayStdin;
  HWND m_wid;
};

#endif // OVERLAY_H
