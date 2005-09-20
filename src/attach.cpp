#ifdef HAVE_CONFIG_H
#include <config.h>
#endif
#include <windows.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "olflange-def.h"
#include "olflange.h"
#include "attach.h"
#include "util.h"


STDMETHODIMP 
CGPGExchExtAttachedFileEvents::OnReadPattFromSzFile 
  (LPATTACH att, LPTSTR lpszFile, ULONG ulFlags)
{
  log_debug ("%s: att=%p lpszFile=%s ulFlags=%lx\n", 
	     __func__, att, lpszFile, ulFlags);
  return S_FALSE;
}
  

STDMETHODIMP 
CGPGExchExtAttachedFileEvents::OnWritePattToSzFile 
  (LPATTACH att, LPTSTR lpszFile, ULONG ulFlags)
{
  log_debug ("%s: att=%p lpszFile=%s ulFlags=%lx\n",
	     __func__, att, lpszFile, ulFlags);
  return S_FALSE;
}


STDMETHODIMP
CGPGExchExtAttachedFileEvents::QueryDisallowOpenPatt (LPATTACH att)
{
  log_debug ("%s: att=%p\n", __func__, att);
  return S_FALSE;
}


STDMETHODIMP 
CGPGExchExtAttachedFileEvents::OnOpenPatt (LPATTACH att)
{
  log_debug ("%s: att=%p\n", __func__, att);
  return S_FALSE;
}



STDMETHODIMP 
CGPGExchExtAttachedFileEvents::OnOpenSzFile (LPTSTR lpszFile, ULONG ulFlags)
{
  log_debug ("%s: lpszFile=%s ulflags=%lx\n", __func__, lpszFile, ulFlags);
  return S_FALSE;
}
