/* attached-file-events.h - GpgolAttachedFileEvents definitions.
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


#ifndef ATTACHED_FILE_EVENTS_H
#define ATTACHED_FILE_EVENTS_H

/*
   GpgolAttachedFileEvents

   The GpgolAttachedFileEvents class implements the processing of
   events related to attachments.
 */
class GpgolAttachedFileEvents : public IExchExtAttachedFileEvents
{
 public:
  GpgolAttachedFileEvents (GpgolExt *pParentInterface);

 private:
  GpgolExt *m_pExchExt;
  ULONG     m_ref;
  
 public:
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID FAR *ppvObj);
  inline STDMETHODIMP_(ULONG) AddRef (void)
    {
      ++m_ref;
      return m_ref;
    }
  inline STDMETHODIMP_(ULONG) Release (void)
    {
      ULONG count = --m_ref;
      if (!count)
	delete this;
      return count;
    }
  
  
  STDMETHODIMP OnReadPattFromSzFile (LPATTACH att, LPTSTR lpszFile, 
				     ULONG ulFlags);
  STDMETHODIMP OnWritePattToSzFile (LPATTACH att, LPTSTR lpszFile, 
				    ULONG ulFlags);
  STDMETHODIMP QueryDisallowOpenPatt (LPATTACH att);
  STDMETHODIMP OnOpenPatt (LPATTACH att);
  STDMETHODIMP OnOpenSzFile (LPTSTR lpszFile, ULONG ulFlags);

};

#endif /*ATTACHED_FILE_EVENTS_H*/
