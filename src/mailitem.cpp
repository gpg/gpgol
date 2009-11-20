/* mailitem.cpp - Code to handle the Outlook MailItem
 *	Copyright (C) 2009 g10 Code GmbH
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <olectl.h>

#include "common.h"
#include "oomhelp.h"
#include "eventsink.h"
#include "mailitem.h"
#include "mymapi.h"
#include "mymapitags.h"
#include "myexchext.h"
#include "mapihelp.h"
#include "message.h"

/* Subclass of ItemEvents so that we can hook into the events.  */
BEGIN_EVENT_SINK(GpgolItemEvents, IOOMItemEvents)
  STDMETHOD(Read)(THIS_ );
  STDMETHOD(Write)(THIS_ PBOOL cancel);
  STDMETHOD(Open)(THIS_ PBOOL cancel);
  STDMETHOD(Close)(THIS_ PBOOL cancel);
  STDMETHOD(Send)(THIS_ PBOOL cancel);
  bool m_send_seen;
EVENT_SINK_CTOR(GpgolItemEvents)
{
  m_send_seen = false;
}
EVENT_SINK_DEFAULT_DTOR(GpgolItemEvents)
EVENT_SINK_INVOKE(GpgolItemEvents)
{
  HRESULT hr;
  (void)lcid; (void)riid; (void)result; (void)exepinfo; (void)argerr;

  switch (dispid)
    {
    case 0xf001:
      if (!(flags & DISPATCH_METHOD))
        goto badflags;
      if (parms->cArgs != 0)
        hr = DISP_E_BADPARAMCOUNT;
      else
        {
          Read ();
          hr = S_OK;
        }
      break;

    case 0xf002:
    case 0xf003:
    case 0xf004:
    case 0xf005:
      if (!(flags & DISPATCH_METHOD))
        goto badflags;
      if (!parms) 
        hr = DISP_E_PARAMNOTOPTIONAL;
      else if (parms->cArgs != 1)
        hr = DISP_E_BADPARAMCOUNT;
      else if (parms->rgvarg[0].vt != (VT_BOOL|VT_BYREF))
        hr = DISP_E_BADVARTYPE;
      else
        {
          BOOL cancel_default = !!*parms->rgvarg[0].pboolVal;
          switch (dispid)
            {
            case 0xf002: Write (&cancel_default); break;
            case 0xf003: Open (&cancel_default); break;
            case 0xf004: Close (&cancel_default); break;
            case 0xf005: Send (&cancel_default); break;
            }
          *parms->rgvarg[0].pboolVal = (cancel_default 
                                        ? VARIANT_TRUE:VARIANT_FALSE);
          hr = S_OK;
        }
      break;

    badflags:
    default:
      hr = DISP_E_MEMBERNOTFOUND;
    }
  return hr;
}
END_EVENT_SINK(GpgolItemEvents, IID_IOOMItemEvents)


/* This is the event sink for a read event.  */
STDMETHODIMP
GpgolItemEvents::Read (void)
{
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this);

  return S_OK;
}


/* This is the event sink for a write event.  OL2003 calls this method
   before an ECE OnWrite.  */
STDMETHODIMP
GpgolItemEvents::Write (PBOOL cancel_default)
{
  bool sign, encrypt, need_crypto, want_html;
  HWND hwnd = NULL;  /* Fixme.  */
 
  log_debug ("%s:%s: Called (this=%p) (send_seen=%d)",
             SRCNAME, __func__, this, m_send_seen);

  if (!m_send_seen)
    return S_OK; /* The user is only saving the message.  */
  m_send_seen = 0;

  if (opt.disable_gpgol)
    return S_OK; /* GpgOL is not active.  */
  
  // need_crypto = (!get_crypto_flags (eecb, &sign, &encrypt)
  //                && (sign || encrypt));
  need_crypto = true;
  sign = true;
  encrypt = false;

    
  /* If we are going to encrypt, check that the BodyFormat is
     something we support.  This helps avoiding surprise by sending
     out unencrypted messages. */
  if (need_crypto)
    {
      HRESULT hr;
      int rc;
      LPUNKNOWN unknown;
      LPMESSAGE message = NULL;
      int bodyfmt;
      protocol_t proto = PROTOCOL_UNKNOWN; /* Let the UI server select
                                              the protocol.  */

      bodyfmt = get_oom_int (m_object, "BodyFormat");

      if (bodyfmt == 1)
        want_html = 0;
      else if (bodyfmt == 2)
        want_html = 1;
      else
        {
          log_debug ("%s:%s: BodyFormat is %d", SRCNAME, __func__, bodyfmt);
          MessageBox (hwnd,
                      _("Sorry, we can only encrypt plain text messages and\n"
                      "no RTF messages. Please make sure that only the text\n"
                      "format has been selected."),
                      "GpgOL", MB_ICONERROR|MB_OK);

          *cancel_default = true;	
          return S_OK;
        }

      /* Unfortunately the Body has not yet been written to the MAPI
         object, although the object already exists.  Thus we have to
         take it from the OOM which requires us to rewrite parts of
         the message encryption functions.  More work ... */
      unknown = get_oom_iunknown (m_object, "MAPIOBJECT");
      if (!unknown)
        log_error ("%s:%s: error getting MAPI object", SRCNAME, __func__);
      else
        {
          hr = unknown->QueryInterface (IID_IMessage, (void**)&message);
          if (hr != S_OK || !message)
            {
              message = NULL;
              log_error ("%s:%s: error getting IMESSAGE: hr=%#lx",
                         SRCNAME, __func__, hr);
            }
          unknown->Release ();
        }

      if (!message)
        rc = -1;
      else if (encrypt && sign)
        rc = message_sign_encrypt (message, proto, hwnd);
      else if (encrypt && !sign)
        rc = message_encrypt (message, proto, hwnd);
      else if (!encrypt && sign)
        rc = message_sign (message, proto, hwnd);
      else
        {
          /* In case this is a forward message which is not to be
             signed or encrypted we need to remove a possible body
             attachment.  */
          if (mapi_delete_gpgol_body_attachment (message))
            mapi_save_changes (message, KEEP_OPEN_READWRITE);
          rc = 0;
        }
      
      if (rc)
        {
          *cancel_default = true;	
          return S_OK;
        }
    }

  return S_OK;
}


/* This is the event sink for an open event.  */
STDMETHODIMP
GpgolItemEvents::Open (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this);

  return S_OK;
}


/* This is the event sink for a close event.  */
STDMETHODIMP
GpgolItemEvents::Close (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this);

  /* Remove ourself.  */
  detach_GpgolItemEvents_sink (this);
  return S_OK;
}


/* This is the event Sink for a send event.  OL2003 calls this method
   before an ECE OnSubmit.  */
STDMETHODIMP
GpgolItemEvents::Send (PBOOL cancel_default)
{
  (void)cancel_default;
  log_debug ("%s:%s: Called (this=%p)", SRCNAME, __func__, this);
  
  m_send_seen = true;

  return S_OK;
}

