/* attached-file-events.cpp - GpgolAttachedFileEvents implementation
 *	Copyright (C) 2005, 2007 g10 Code GmbH
 * 
 * This file is part of GpgOL.
 * 
 * GpgOL is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2 of the License, or (at your option) any later version.
 * 
 * GpgOL is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this program; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA
 * 02110-1301, USA.
 */

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif
#include <windows.h>
#include <errno.h>

#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "common.h"
#include "olflange-def.h"
#include "olflange.h"
#include "serpent.h"
#include "attached-file-events.h"

#define TRACEPOINT() do { log_debug ("%s:%s:%d: tracepoint\n", \
                                     SRCNAME, __func__, __LINE__); \
                        } while (0)

#define COPYBUFFERSIZE 4096

/* Copy STREAM to a new file FILENAME while decrypting it using the
   context SYMENC.  */
static int
decrypt_and_write_file (LPSTREAM stream, const char *filename, symenc_t symenc)
{
  int rc = E_ABORT;
  HRESULT hr;
  ULONG nread;
  char *buf = NULL;
  FILE *fpout = NULL;

  fpout = fopen (filename, "wb");
  if (!fpout)
    {
      log_error ("%s:%s: fwrite failed: %s", SRCNAME, __func__,
                 strerror (errno));
      MessageBox (NULL,
                  _("Error creating file for attachment."),
                  "GpgOL", MB_ICONERROR|MB_OK);
      goto leave;
    }

  buf = (char*)xmalloc (COPYBUFFERSIZE);
  do
    {
      hr = stream->Read (buf, COPYBUFFERSIZE, &nread);
      if (hr)
        {
          log_error ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
          MessageBox (NULL,
                      _("Error reading attachment."),
                      "GpgOL", MB_ICONERROR|MB_OK);
          goto leave;
        }
      if (nread)
        symenc_cfb_decrypt (symenc, buf, buf, nread);
      if (nread && fwrite (buf, nread, 1, fpout) != 1)
        {
          log_error ("%s:%s: fwrite failed: %s", SRCNAME, __func__,
                     strerror (errno));
          MessageBox (NULL,
                      _("Error writing attachment."),
                      "GpgOL", MB_ICONERROR|MB_OK);
          goto leave;
        }
    }
  while (nread == COPYBUFFERSIZE);
  
  xfree (buf);
  if (fclose (fpout))
    {
      log_error ("%s:%s: fclose failed: %s", 
                 SRCNAME, __func__, strerror (errno));
      MessageBox (NULL,
                  _("Error writing attachment."),
                  "GpgOL", MB_ICONERROR|MB_OK);
      goto leave;
    } 

  rc = S_OK;

 leave:
  xfree (buf);
  if (fpout)
    fclose (fpout);
  if (rc != S_OK)
    remove (filename);
  return rc;
}




/* Our constructor.  */
GpgolAttachedFileEvents::GpgolAttachedFileEvents (GpgolExt *pParentInterface)
{ 
  m_pExchExt = pParentInterface;
  m_ref = 0;
}


/* The QueryInterfac.  */
STDMETHODIMP 
GpgolAttachedFileEvents::QueryInterface (REFIID riid, LPVOID FAR *ppvObj)
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
  return E_NOINTERFACE;
}
 

/* Fixme: We need to figure out what this exactly does.  There is no
   public information available exepct for the MAPI book which is out
   of print.  

   This seems to be called if one adds a new attachment to a the composer.
*/
STDMETHODIMP 
GpgolAttachedFileEvents::OnReadPattFromSzFile 
  (LPATTACH att, LPTSTR file, ULONG flags)
{
  log_debug ("%s:%s: att=%p file=`%s' flags=%lx\n", 
	     SRCNAME, __func__, att, file, flags);
  return S_FALSE;
}
  
 
/* This seems to be called if one clicks on Save in the context menu.
   And also sometimes before an Open click. */
STDMETHODIMP 
GpgolAttachedFileEvents::OnWritePattToSzFile 
  (LPATTACH att, LPTSTR file, ULONG flags)
{
  HRESULT hr;
  ULONG tag;
  char *iv;
  size_t ivlen;
  symenc_t symenc;
  LPSTREAM stream;
  char tmpbuf[16];
  ULONG nread;
  int rc;

  log_debug ("%s:%s: att=%p file=`%s' flags=%lx\n", 
	     SRCNAME, __func__, att, file, flags);
  if (!att)
    return E_FAIL;

  if (get_gpgolprotectiv_tag ((LPMESSAGE)att, &tag) )
    return E_ABORT;
  iv = mapi_get_binary_prop ((LPMESSAGE)att, tag, &ivlen);
  if (!iv)
    return S_FALSE; /* Not encrypted by us - Let OL continue as usual.  */

  symenc = symenc_open (get_128bit_session_key (), 16, iv, ivlen);
  xfree (iv);
  if (!symenc)
    {
      log_error ("%s:%s: can't open encryption context", SRCNAME, __func__);
      return E_ABORT;
    }

  hr = att->OpenProperty (PR_ATTACH_DATA_BIN, &IID_IStream, 
                          0, 0, (LPUNKNOWN*) &stream);
  if (FAILED (hr))
    {
      log_error ("%s:%s: can't open data stream of attachment: hr=%#lx",
                 SRCNAME, __func__, hr);
      symenc_close (symenc);
      return E_ABORT;
    }

  hr = stream->Read (tmpbuf, 16, &nread);
  if (hr)
    {
      log_debug ("%s:%s: Read failed: hr=%#lx", SRCNAME, __func__, hr);
      stream->Release ();
      symenc_close (symenc);
      return E_ABORT;
    }
  symenc_cfb_decrypt (symenc, tmpbuf, tmpbuf, 16);
  if (memcmp (tmpbuf, "GpgOL attachment", 16))
    {
      MessageBox (NULL,
                  _("Sorry, we are not able to decrypt this attachment.\n\n"
                    "Please use the decrypt/verify button to decrypt the\n"
                    "entire message again.  Then open this attachment."),
                  "GpgOL", MB_ICONERROR|MB_OK);
      stream->Release ();
      symenc_close (symenc);
      return E_ABORT;
    }

  rc = decrypt_and_write_file (stream, file, symenc);

  stream->Release ();
  symenc_close (symenc);
  return rc;
}


STDMETHODIMP
GpgolAttachedFileEvents::QueryDisallowOpenPatt (LPATTACH att)
{
  log_debug ("%s:%s: att=%p\n", SRCNAME, __func__, att);
  return S_FALSE;
}


STDMETHODIMP 
GpgolAttachedFileEvents::OnOpenPatt (LPATTACH att)
{
  log_debug ("%s:%s: att=%p\n", SRCNAME, __func__, att);
  return S_FALSE;
}


/* This seems to be called if one clicks on Open in the context menu.  */
STDMETHODIMP 
GpgolAttachedFileEvents::OnOpenSzFile (LPTSTR file, ULONG flags)
{
  log_debug ("%s:%s: file=`%s' flags=%lx\n", SRCNAME, __func__, file, flags);
  return S_FALSE;
}
