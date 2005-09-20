
#ifndef ATTACH_H
#define ATTACH_H

#include "util.h"

class CGPGExchExtAttachedFileEvents : public IExchExtAttachedFileEvents
{
private:
  CGPGExchExt *m_pExchExt;
  ULONG        m_ref;

public:
  CGPGExchExtAttachedFileEvents (CGPGExchExt *pParentInterface)
    {
      m_pExchExt = pParentInterface;
      m_ref = 0;
    }
  
  
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
  
  STDMETHODIMP QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
    {
      *ppvObj = NULL;
      if (riid == IID_IExchExtAttachedFileEvents)
	{
	  *ppvObj = (LPVOID)this;
	  AddRef ();
	  return S_OK;
	}
      
      if (riid == IID_IUnknown)
	{
	  *ppvObj = (LPVOID)m_pExchExt;
	  m_pExchExt->AddRef ();
	  return S_OK;
	}
      log_debug ("%s: request for unknown interface\n", __func__);
      return E_NOINTERFACE;
    }
  
  STDMETHODIMP OnReadPattFromSzFile (LPATTACH att, LPTSTR lpszFile, 
				     ULONG ulFlags);
  STDMETHODIMP OnWritePattToSzFile (LPATTACH att, LPTSTR lpszFile, 
				    ULONG ulFlags);
  STDMETHODIMP QueryDisallowOpenPatt (LPATTACH att);
  STDMETHODIMP OnOpenPatt (LPATTACH att);
  STDMETHODIMP OnOpenSzFile (LPTSTR lpszFile, ULONG ulFlags);
};

#endif /*ATTACH_H*/
