/* oomhelp.cpp - Helper functions for the Outlook Object Model
 * Copyright (C) 2009 g10 Code GmbH
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <windows.h>
#include <olectl.h>
#include <string>
#include <sstream>
#include <algorithm>
#include <rpc.h>

#include "common.h"

#include "oomhelp.h"
#include "cpphelp.h"
#include "gpgoladdin.h"
#include "categorymanager.h"
#include "recipient.h"

HRESULT
gpgol_queryInterface (LPUNKNOWN pObj, REFIID riid, LPVOID FAR *ppvObj)
{
  HRESULT ret = pObj->QueryInterface (riid, ppvObj);
  if (ret)
    {
      log_debug ("%s:%s: QueryInterface failed hr=%#lx",
                 SRCNAME, __func__, ret);
    }
  else if ((opt.enable_debug & DBG_MEMORY) && *ppvObj)
    {
      memdbg_addRef (*ppvObj);
    }
  return ret;
}

HRESULT
gpgol_openProperty (LPMAPIPROP obj, ULONG ulPropTag, LPCIID lpiid,
                    ULONG ulInterfaceOptions, ULONG ulFlags,
                    LPUNKNOWN FAR * lppUnk)
{
  HRESULT ret = obj->OpenProperty (ulPropTag, lpiid,
                                   ulInterfaceOptions, ulFlags,
                                   lppUnk);
  if (ret)
    {
      log_debug ("%s:%s: OpenProperty failed hr=%#lx %s",
                 SRCNAME, __func__, ret, mapi_err_to_string (ret));
    }
  else if ((opt.enable_debug & DBG_MEMORY) && *lppUnk)
    {
      memdbg_addRef (*lppUnk);
      log_debug ("%s:%s: OpenProperty on %p prop %lx result %p",
                 SRCNAME, __func__, obj,  ulPropTag, *lppUnk);
    }
  return ret;
}
/* Return a malloced string with the utf-8 encoded name of the object
   or NULL if not available.  */
char *
get_object_name (LPUNKNOWN obj)
{
  TSTART;
  HRESULT hr;
  LPDISPATCH disp = NULL;
  LPTYPEINFO tinfo = NULL;
  BSTR bstrname;
  char *name = NULL;

  if (!obj)
    goto leave;

  /* We can't use gpgol_queryInterface here to avoid recursion */
  hr = obj->QueryInterface (IID_IDispatch, (void **)&disp);
  if (!disp || hr != S_OK)
    goto leave;

  disp->GetTypeInfo (0, 0, &tinfo);
  if (!tinfo)
    {
      log_debug ("%s:%s: no typeinfo found for object\n",
                 SRCNAME, __func__);
      goto leave;
    }

  bstrname = NULL;
  hr = tinfo->GetDocumentation (MEMBERID_NIL, &bstrname, 0, 0, 0);
  if (hr || !bstrname)
    log_debug ("%s:%s: GetDocumentation failed: hr=%#lx\n",
               SRCNAME, __func__, hr);
  if (bstrname)
    {
      name = wchar_to_utf8 (bstrname);
      SysFreeString (bstrname);
    }

 leave:
  if (tinfo)
    tinfo->Release ();
  if (disp)
    disp->Release ();

  TRETURN name;
}

std::string
get_object_name_s (LPUNKNOWN obj)
{
  char *name = get_object_name (obj);
  std::string ret = "(null)";
  if (name)
    {
      ret = name;
    }
  xfree (name);
  return ret;
}

std::string
get_object_name_s (shared_disp_t obj)
{
  return get_object_name_s (obj.get ());
}

/* Lookup the dispid of object PDISP for member NAME.  Returns
   DISPID_UNKNOWN on error.  */
DISPID
lookup_oom_dispid (LPDISPATCH pDisp, const char *name)
{
  HRESULT hr;
  DISPID dispid;
  wchar_t *wname;

  if (!pDisp || !name)
    {
      TRETURN DISPID_UNKNOWN; /* Error: Invalid arg.  */
    }

  wname = utf8_to_wchar (name);
  if (!wname)
    {
      TRETURN DISPID_UNKNOWN;/* Error:  Out of memory.  */
    }

  hr = pDisp->GetIDsOfNames (IID_NULL, &wname, 1,
                             LOCALE_SYSTEM_DEFAULT, &dispid);
  xfree (wname);
  if (hr != S_OK || dispid == DISPID_UNKNOWN)
    log_debug ("%s:%s: error looking up dispid(%s)=%d: hr=0x%x\n",
               SRCNAME, __func__, name, (int)dispid, (unsigned int)hr);
  if (hr != S_OK)
    dispid = DISPID_UNKNOWN;

  return dispid;
}

static void
init_excepinfo (EXCEPINFO *err)
{
  if (!err)
    {
      TRETURN;
    }
  err->wCode = 0;
  err->wReserved = 0;
  err->bstrSource = nullptr;
  err->bstrDescription = nullptr;
  err->bstrHelpFile = nullptr;
  err->dwHelpContext = 0;
  err->pvReserved = nullptr;
  err->pfnDeferredFillIn = nullptr;
  err->scode = 0;
}

void
dump_excepinfo (EXCEPINFO err)
{
  log_oom ("%s:%s: Exception: \n"
             "              wCode: 0x%x\n"
             "              wReserved: 0x%x\n"
             "              source: %S\n"
             "              desc: %S\n"
             "              help: %S\n"
             "              helpCtx: 0x%x\n"
             "              deferredFill: %p\n"
             "              scode: 0x%x\n",
             SRCNAME, __func__, (unsigned int) err.wCode,
             (unsigned int) err.wReserved,
             err.bstrSource ? err.bstrSource : L"null",
             err.bstrDescription ? err.bstrDescription : L"null",
             err.bstrHelpFile ? err.bstrDescription : L"null",
             (unsigned int) err.dwHelpContext,
             err.pfnDeferredFillIn,
             (unsigned int) err.scode);
}

/* Return the OOM object's IDispatch interface described by FULLNAME.
   Returns NULL if not found.  PSTART is the object where the search
   starts.  FULLNAME is a dot delimited sequence of object names.  If
   an object name has a "(foo)" suffix this passes it as a parameter
   to the invoke function (i.e. using (DISPATCH|PROPERTYGET)).  Object
   names including the optional suffix are truncated at 127 byte.  */
LPDISPATCH
get_oom_object (LPDISPATCH pStart, const char *fullname)
{
  TSTART;
  HRESULT hr;
  LPDISPATCH pObj = pStart;
  LPDISPATCH pDisp = NULL;

  log_oom ("%s:%s: looking for %p->`%s'",
           SRCNAME, __func__, pStart, fullname);

  while (pObj)
    {
      DISPPARAMS dispparams;
      VARIANT aVariant[4];
      VARIANT vtResult;
      wchar_t *wname;
      char name[128];
      int n_parms = 0;
      BSTR parmstr = NULL;
      INT  parmint = 0;
      DISPID dispid;
      char *p, *pend;
      int dispmethod;
      unsigned int argErr = 0;
      EXCEPINFO execpinfo;

      init_excepinfo (&execpinfo);

      if (pDisp)
        {
          gpgol_release (pDisp);
          pDisp = NULL;
        }
      if (gpgol_queryInterface (pObj, IID_IDispatch, (LPVOID*)&pDisp) != S_OK)
        {
          log_error ("%s:%s Object does not support IDispatch",
                     SRCNAME, __func__);
          if (pObj != pStart)
            gpgol_release (pObj);
          TRETURN NULL;
        }
      /* Confirmed through testing that the retval needs a release */
      if (pObj != pStart)
        gpgol_release (pObj);
      pObj = NULL;
      if (!pDisp)
        {
          TRETURN NULL;  /* The object has no IDispatch interface.  */
        }
      if (!*fullname)
        {
          if ((opt.enable_debug & DBG_MEMORY))
            {
              pDisp->AddRef ();
              int ref = pDisp->Release ();
              log_oom ("%s:%s:         got %p with %i refs",
                       SRCNAME, __func__, pDisp, ref);
            }
          TRETURN pDisp; /* Ready.  */
        }

      /* Break out the next name part.  */
      {
        const char *dot;
        size_t n;

        dot = strchr (fullname, '.');
        if (dot == fullname)
          {
            gpgol_release (pDisp);
            TRETURN NULL;  /* Empty name part: error.  */
          }
        else if (dot)
          n = dot - fullname;
        else
          n = strlen (fullname);

        if (n >= sizeof name)
          n = sizeof name - 1;
/* As you can see above n is checked for bounds of name. But
   GCC warns because it just looked at n = strlen (fullname). */
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wstringop-truncation"
        strncpy (name, fullname, n);
#pragma GCC diagnostic pop
        name[n] = 0;

        if (dot)
          fullname = dot + 1;
        else
          fullname += strlen (fullname);
      }

      if (!strncmp (name, "get_", 4) && name[4])
        {
          dispmethod = DISPATCH_PROPERTYGET;
          memmove (name, name+4, strlen (name+4)+1);
        }
      else if ((p = strchr (name, '(')))
        {
          *p++ = 0;
          pend = strchr (p, ')');
          if (pend)
            *pend = 0;

          if (*p == ',' && p[1] != ',')
            {
              /* We assume this is "foo(,30007)".  I.e. the frst arg
                 is not given and the second one is an integer.  */
              parmint = (int)strtol (p+1, NULL, 10);
              n_parms = 4;
            }
          else
            {
              wname = utf8_to_wchar (p);
              if (wname)
                {
                  parmstr = SysAllocString (wname);
                  xfree (wname);
                }
              if (!parmstr)
                {
                  gpgol_release (pDisp);
                  TRETURN NULL; /* Error:  Out of memory.  */
                }
              n_parms = 1;
            }
          dispmethod = DISPATCH_METHOD|DISPATCH_PROPERTYGET;
        }
      else
        dispmethod = DISPATCH_METHOD;

      /* Lookup the dispid.  */
      dispid = lookup_oom_dispid (pDisp, name);
      if (dispid == DISPID_UNKNOWN)
        {
          if (parmstr)
            SysFreeString (parmstr);
          gpgol_release (pDisp);
          TRETURN NULL;  /* Name not found.  */
        }

      /* Invoke the method.  */
      dispparams.rgvarg = aVariant;
      dispparams.cArgs = 0;
      if (n_parms)
        {
          if (n_parms == 4)
            {
              dispparams.rgvarg[0].vt = VT_ERROR;
              dispparams.rgvarg[0].scode = DISP_E_PARAMNOTFOUND;
              dispparams.rgvarg[1].vt = VT_ERROR;
              dispparams.rgvarg[1].scode = DISP_E_PARAMNOTFOUND;
              dispparams.rgvarg[2].vt = VT_INT;
              dispparams.rgvarg[2].intVal = parmint;
              dispparams.rgvarg[3].vt = VT_ERROR;
              dispparams.rgvarg[3].scode = DISP_E_PARAMNOTFOUND;
              dispparams.cArgs = n_parms;
            }
          else if (n_parms == 1 && parmstr)
            {
              dispparams.rgvarg[0].vt = VT_BSTR;
              dispparams.rgvarg[0].bstrVal = parmstr;
              dispparams.cArgs++;
            }
        }
      dispparams.cNamedArgs = 0;
      VariantInit (&vtResult);
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          dispmethod, &dispparams,
                          &vtResult, &execpinfo, &argErr);
      if (parmstr)
        SysFreeString (parmstr);
      if (hr != S_OK || vtResult.vt != VT_DISPATCH)
        {
          log_debug ("%s:%s: failure: '%s' p=%p vt=%d hr=0x%x argErr=0x%x dispid=0x%x",
                     SRCNAME, __func__,
                     name, vtResult.pdispVal, vtResult.vt, (unsigned int)hr,
                     (unsigned int)argErr, (unsigned int)dispid);
          dump_excepinfo (execpinfo);
          VariantClear (&vtResult);
          gpgol_release (pDisp);
          TRETURN NULL;  /* Invoke failed.  */
        }

      pObj = vtResult.pdispVal;
      memdbg_addRef (pObj);
    }
  gpgol_release (pDisp);
  log_debug ("%s:%s: no object", SRCNAME, __func__);
  TRETURN NULL;
}

shared_disp_t
get_oom_object_s (shared_disp_t pStart, const char *fullname)
{
  return MAKE_SHARED (get_oom_object (pStart.get (), fullname));
}

shared_disp_t
get_oom_object_s (LPDISPATCH pStart, const char *fullname)
{
  return MAKE_SHARED (get_oom_object (pStart, fullname));
}

/* Helper for put_oom_icon.  */
static int
put_picture_or_mask (LPDISPATCH pDisp, int resource, int size, int is_mask)
{
  TSTART;
  HRESULT hr;
  PICTDESC pdesc;
  LPDISPATCH pPict;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  UINT fuload;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[2];

  /* When loading the mask we need to set the monochrome flag.  We
     better create a DIB section to avoid possible rendering
     problems.  */
  fuload = LR_CREATEDIBSECTION | LR_SHARED;
  if (is_mask)
    fuload |= LR_MONOCHROME;

  memset (&pdesc, 0, sizeof pdesc);
  pdesc.cbSizeofstruct = sizeof pdesc;
  pdesc.picType = PICTYPE_BITMAP;
  pdesc.bmp.hbitmap = (HBITMAP) LoadImage (glob_hinst,
                                           MAKEINTRESOURCE (resource),
                                           IMAGE_BITMAP, size, size, fuload);
  if (!pdesc.bmp.hbitmap)
    {
      log_error_w32 (-1, "%s:%s: LoadImage(%d) failed\n",
                     SRCNAME, __func__, resource);
      TRETURN -1;
    }

  /* Wrap the image into an OLE object.  */
  hr = OleCreatePictureIndirect (&pdesc, IID_IPictureDisp,
                                 TRUE, (void **) &pPict);
  if (hr != S_OK || !pPict)
    {
      log_error ("%s:%s: OleCreatePictureIndirect failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      TRETURN -1;
    }

  /* Store to the Picture or Mask property of the CommandBarButton.  */
  dispid = lookup_oom_dispid (pDisp, is_mask? "Mask":"Picture");

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_DISPATCH;
  dispparams.rgvarg[0].pdispVal = pPict;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT, &dispparams,
                      NULL, NULL, NULL);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting icon failed: %#lx", SRCNAME, __func__, hr);
      TRETURN -1;
    }
  TRETURN 0;
}


/* Update the icon of PDISP using the bitmap with RESOURCE ID.  The
   function adds the system pixel size to the resource id to compute
   the actual icon size.  The resource id of the mask is the N+1.  */
int
put_oom_icon (LPDISPATCH pDisp, int resource_id, int size)
{
  TSTART;
  int rc;

  /* This code is only relevant for Outlook < 2010.
    Ideally it should grab the system pixel size and use an
    icon of the appropiate size (e.g. 32 or 64px)
  */

  rc = put_picture_or_mask (pDisp, resource_id, size, 0);
  if (!rc)
    rc = put_picture_or_mask (pDisp, resource_id + 1, size, 1);

  TRETURN rc;
}


/* Set the boolean property NAME to VALUE.  */
int
put_oom_bool (LPDISPATCH pDisp, const char *name, int value)
{
  TSTART;
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    {
      TRETURN -1;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;
  dispparams.rgvarg[0].boolVal = value? VARIANT_TRUE:VARIANT_FALSE;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT, &dispparams,
                      NULL, NULL, NULL);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting '%s' failed: %#lx",
                 SRCNAME, __func__, name, hr);
      TRETURN -1;
    }
  TRETURN 0;
}


/* Set the property NAME to VALUE.  */
int
put_oom_int (LPDISPATCH pDisp, const char *name, int value)
{
  TSTART;
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    {
      TRETURN -1;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = value;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT, &dispparams,
                      NULL, NULL, NULL);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting '%s' failed: %#lx",
                 SRCNAME, __func__, name, hr);
      TRETURN -1;
    }
  TRETURN 0;
}

/* Set the property NAME to VALUE.  */
int
put_oom_array (LPDISPATCH pDisp, const char *name, unsigned char *value,
               size_t size)
{
  TSTART;
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];
  unsigned int argErr = 0;
  EXCEPINFO execpinfo;
  init_excepinfo (&execpinfo);

  VariantInit (aVariant);

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    {
      TRETURN -1;
    }
  /* Prepare the savearray */
  SAFEARRAYBOUND saBound;
  saBound.lLbound = 0;
  saBound.cElements = size;
  SAFEARRAY* psa = SafeArrayCreate(VT_UI1, 1, &saBound);

  if (!psa)
    {
      log_err ("Failed to create SafeArray");
      TRETURN -1;
    }

  hr = SafeArrayLock(psa);
  if (!SUCCEEDED (hr))
    {
      log_err ("Failed to lock array.");
    }
  memcpy (psa->pvData, value, size);
  SafeArrayUnlock(psa);

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_ARRAY;
  dispparams.rgvarg[0].parray = psa;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT, &dispparams,
                      NULL, &execpinfo, &argErr);
  SafeArrayDestroy(psa);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: error: putting %s"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__, name,
                 (unsigned int)hr,
                 (unsigned int)argErr);
      dump_excepinfo (execpinfo);
      TRETURN -1;
    }
  TRETURN 0;
}

/* Set the property NAME to STRING.  */
int
put_oom_string (LPDISPATCH pDisp, const char *name, const char *string)
{
  TSTART;
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];
  BSTR bstring;
  EXCEPINFO execpinfo;

  init_excepinfo (&execpinfo);
  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    {
      TRETURN -1;
    }

  {
    wchar_t *tmp = utf8_to_wchar (string);
    bstring = tmp? SysAllocString (tmp):NULL;
    xfree (tmp);
    if (!bstring)
      {
        log_error_w32 (-1, "%s:%s: SysAllocString failed", SRCNAME, __func__);
        TRETURN -1;
      }
  }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BSTR;
  dispparams.rgvarg[0].bstrVal = bstring;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT, &dispparams,
                      NULL, &execpinfo, NULL);
  SysFreeString (bstring);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting '%s' failed: %#lx",
                 SRCNAME, __func__, name, hr);
      dump_excepinfo (execpinfo);
      TRETURN -1;
    }
  TRETURN 0;
}

/* Set the property NAME to DISP.  */
int
put_oom_disp (LPDISPATCH pDisp, const char *name, LPDISPATCH disp)
{
  TSTART;
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];
  EXCEPINFO execpinfo;

  init_excepinfo (&execpinfo);
  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    {
      TRETURN -1;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_DISPATCH;
  dispparams.rgvarg[0].pdispVal = disp;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUTREF, &dispparams,
                      NULL, &execpinfo, NULL);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting '%s' failed: %#lx",
                 SRCNAME, __func__, name, hr);
      dump_excepinfo (execpinfo);
      TRETURN -1;
    }
  TRETURN 0;
}

/* Get the boolean property NAME of the object PDISP.  Returns False if
   not found or if it is not a boolean property.  */
int
get_oom_bool (LPDISPATCH pDisp, const char *name)
{
  TSTART;
  HRESULT hr;
  int result = 0;
  DISPID dispid;

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid != DISPID_UNKNOWN)
    {
      DISPPARAMS dispparams = {NULL, NULL, 0, 0};
      VARIANT rVariant;

      VariantInit (&rVariant);
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparams,
                          &rVariant, NULL, NULL);
      if (hr != S_OK)
        log_debug ("%s:%s: Property '%s' not found: %#lx",
                   SRCNAME, __func__, name, hr);
      else if (rVariant.vt != VT_BOOL)
        log_debug ("%s:%s: Property `%s' is not a boolean (vt=%d)",
                   SRCNAME, __func__, name, rVariant.vt);
      else
        result = !!rVariant.boolVal;
      VariantClear (&rVariant);
    }

  TRETURN result;
}


/* Get the integer property NAME of the object PDISP.  Returns 0 if
   not found or if it is not an integer property.  */
int
get_oom_int (LPDISPATCH pDisp, const char *name)
{
  TSTART;
  HRESULT hr;
  int result = 0;
  DISPID dispid;

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid != DISPID_UNKNOWN)
    {
      DISPPARAMS dispparams = {NULL, NULL, 0, 0};
      VARIANT rVariant;

      VariantInit (&rVariant);
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparams,
                          &rVariant, NULL, NULL);
      if (hr != S_OK)
        log_debug ("%s:%s: Property '%s' not found: %#lx",
                   SRCNAME, __func__, name, hr);
      else if (rVariant.vt != VT_INT && rVariant.vt != VT_I4)
        log_debug ("%s:%s: Property `%s' is not an integer (vt=%d)",
                   SRCNAME, __func__, name, rVariant.vt);
      else
        result = rVariant.intVal;
      VariantClear (&rVariant);
    }

  TRETURN result;
}

int
get_oom_int (shared_disp_t pDisp, const char *name)
{
  return get_oom_int (pDisp.get (), name);
}

int
get_oom_dirty (LPDISPATCH pDisp)
{
  TSTART;
  HRESULT hr;
  DISPID dispid = DISPID_DIRTY_RAT;

  DISPPARAMS dispparams = {NULL, NULL, 0, 0};
  VARIANT rVariant;

  VariantInit (&rVariant);
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYGET | DISPATCH_METHOD,
                      &dispparams, &rVariant, NULL, NULL);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Property dirty not found: %#lx",
                 SRCNAME, __func__, hr);
      TRETURN -1;
    }
  return !!rVariant.bVal;
}


#if 0
int
put_oom_dirty (LPDISPATCH pDisp, bool value)
{
  TSTART;

  /* NOTE: I have found no scenario where this does
     not return the exception that the property is
     write protected. But we can never know when
     we need such an arcane function so I left it in. */

  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid = DISPID_DIRTY_RAT;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];
  unsigned int argErr = 0;
  EXCEPINFO execpinfo;
  init_excepinfo (&execpinfo);

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;
  dispparams.rgvarg[0].boolVal = value? VARIANT_TRUE:VARIANT_FALSE;
  dispparams.cArgs = 1;
  dispparams.rgdispidNamedArgs = &dispid_put;
  dispparams.cNamedArgs = 1;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_PROPERTYPUT | DISPATCH_METHOD, &dispparams,
                      NULL, &execpinfo, &argErr);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: error: invoking dirty p=%p vt=%d"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 nullptr, 0, (unsigned int)hr,
                 (unsigned int)argErr);
      dump_excepinfo (execpinfo);
      TRETURN -1;
    }
  TRETURN 0;
}
#endif

/* Get the string property NAME of the object PDISP.  Returns NULL if
   not found or if it is not a string property.  */
char *
get_oom_string (LPDISPATCH pDisp, const char *name)
{
  TSTART;
  HRESULT hr;
  char *result = NULL;
  DISPID dispid;

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid != DISPID_UNKNOWN)
    {
      DISPPARAMS dispparams = {NULL, NULL, 0, 0};
      VARIANT rVariant;

      VariantInit (&rVariant);
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparams,
                          &rVariant, NULL, NULL);
      if (hr != S_OK)
        log_debug ("%s:%s: Property '%s' not found: %#lx",
                   SRCNAME, __func__, name, hr);
      else if (rVariant.vt != VT_BSTR)
        log_debug ("%s:%s: Property `%s' is not a string (vt=%d)",
                   SRCNAME, __func__, name, rVariant.vt);
      else if (rVariant.bstrVal)
        result = wchar_to_utf8 (rVariant.bstrVal);
      VariantClear (&rVariant);
    }

  TRETURN result;
}

std::string
get_oom_string_s (LPDISPATCH pDisp, const char *name)
{
  char *ret_c =  get_oom_string (pDisp, name);
  std::string ret;
  if (ret_c)
    {
       ret = ret_c;
       xfree (ret_c);
    }
  return ret;
}

std::string
get_oom_string_s (shared_disp_t pDisp, const char *name)
{
  return get_oom_string_s (pDisp.get (), name);
}
/* Get the object property NAME of the object PDISP.  Returns NULL if
   not found or if it is not an object perty.  */
LPUNKNOWN
get_oom_iunknown (LPDISPATCH pDisp, const char *name)
{
  TSTART;
  HRESULT hr;
  DISPID dispid;

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid != DISPID_UNKNOWN)
    {
      DISPPARAMS dispparams = {NULL, NULL, 0, 0};
      VARIANT rVariant;

      VariantInit (&rVariant);
      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          DISPATCH_PROPERTYGET, &dispparams,
                          &rVariant, NULL, NULL);
      if (hr != S_OK)
        log_debug ("%s:%s: Property '%s' not found: %#lx",
                   SRCNAME, __func__, name, hr);
      else if (rVariant.vt != VT_UNKNOWN)
        log_debug ("%s:%s: Property `%s' is not of class IUnknown (vt=%d)",
                   SRCNAME, __func__, name, rVariant.vt);
      else
        {
          memdbg_addRef (rVariant.punkVal);
          TRETURN rVariant.punkVal;
        }

      VariantClear (&rVariant);
    }

  TRETURN NULL;
}


/* Return the control object described by the tag property with value
   TAG. The object POBJ must support the FindControl method.  Returns
   NULL if not found.  */
LPDISPATCH
get_oom_control_bytag (LPDISPATCH pDisp, const char *tag)
{
  TSTART;
  HRESULT hr;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[4];
  VARIANT rVariant;
  BSTR bstring;
  LPDISPATCH result = NULL;

  dispid = lookup_oom_dispid (pDisp, "FindControl");
  if (dispid == DISPID_UNKNOWN)
    {
      log_debug ("%s:%s: Object %p has no FindControl method",
                 SRCNAME, __func__, pDisp);
      TRETURN NULL;
    }

  {
    wchar_t *tmp = utf8_to_wchar (tag);
    bstring = tmp? SysAllocString (tmp):NULL;
    xfree (tmp);
    if (!bstring)
      {
        log_error_w32 (-1, "%s:%s: SysAllocString failed", SRCNAME, __func__);
        TRETURN NULL;
      }
  }
  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_ERROR; /* Visible */
  dispparams.rgvarg[0].scode = DISP_E_PARAMNOTFOUND;
  dispparams.rgvarg[1].vt = VT_BSTR;  /* Tag */
  dispparams.rgvarg[1].bstrVal = bstring;
  dispparams.rgvarg[2].vt = VT_ERROR; /* Id */
  dispparams.rgvarg[2].scode = DISP_E_PARAMNOTFOUND;
  dispparams.rgvarg[3].vt = VT_ERROR;/* Type */
  dispparams.rgvarg[3].scode = DISP_E_PARAMNOTFOUND;
  dispparams.cArgs = 4;
  dispparams.cNamedArgs = 0;
  VariantInit (&rVariant);
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      &rVariant, NULL, NULL);
  SysFreeString (bstring);
  if (hr == S_OK && rVariant.vt == VT_DISPATCH && rVariant.pdispVal)
    {
      gpgol_queryInterface (rVariant.pdispVal, IID_IDispatch,
                            (LPVOID*)&result);
      gpgol_release (rVariant.pdispVal);
      if (!result)
        log_debug ("%s:%s: Object with tag `%s' has no dispatch intf.",
                   SRCNAME, __func__, tag);
    }
  else
    {
      log_debug ("%s:%s: No object with tag `%s' found: vt=%d hr=%#lx",
                 SRCNAME, __func__, tag, rVariant.vt, hr);
      VariantClear (&rVariant);
    }

  TRETURN result;
}


/* Add a new button to an object which supports the add method.
   Returns the new object or NULL on error.  */
LPDISPATCH
add_oom_button (LPDISPATCH pObj)
{
  TSTART;
  HRESULT hr;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[5];
  VARIANT rVariant;

  dispid = lookup_oom_dispid (pObj, "Add");

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;  /* Temporary */
  dispparams.rgvarg[0].boolVal = VARIANT_TRUE;
  dispparams.rgvarg[1].vt = VT_ERROR;  /* Before */
  dispparams.rgvarg[1].scode = DISP_E_PARAMNOTFOUND;
  dispparams.rgvarg[2].vt = VT_ERROR;  /* Parameter */
  dispparams.rgvarg[2].scode = DISP_E_PARAMNOTFOUND;
  dispparams.rgvarg[3].vt = VT_ERROR;  /* Id */
  dispparams.rgvarg[3].scode = DISP_E_PARAMNOTFOUND;
  dispparams.rgvarg[4].vt = VT_INT;    /* Type */
  dispparams.rgvarg[4].intVal = MSOCONTROLBUTTON;
  dispparams.cArgs = 5;
  dispparams.cNamedArgs = 0;
  VariantInit (&rVariant);
  hr = pObj->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                     DISPATCH_METHOD, &dispparams,
                     &rVariant, NULL, NULL);
  if (hr != S_OK || rVariant.vt != VT_DISPATCH || !rVariant.pdispVal)
    {
      log_error ("%s:%s: Adding Control failed: %#lx - vt=%d",
                 SRCNAME, __func__, hr, rVariant.vt);
      VariantClear (&rVariant);
      TRETURN NULL;
    }
  TRETURN rVariant.pdispVal;
}


/* Add a new button to an object which supports the add method.
   Returns the new object or NULL on error.  */
void
del_oom_button (LPDISPATCH pObj)
{
  TSTART;
  HRESULT hr;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[5];

  dispid = lookup_oom_dispid (pObj, "Delete");

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;  /* Temporary */
  dispparams.rgvarg[0].boolVal = VARIANT_FALSE;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  hr = pObj->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                     DISPATCH_METHOD, &dispparams,
                     NULL, NULL, NULL);
  if (hr != S_OK)
    log_error ("%s:%s: Deleting Control failed: %#lx",
               SRCNAME, __func__, hr);
  TRETURN;
}

/* Gets the current contexts HWND. Returns NULL on error */
HWND
get_oom_context_window (LPDISPATCH context)
{
  TSTART;
  LPOLEWINDOW actExplorer;
  HWND ret = NULL;
  actExplorer = (LPOLEWINDOW) get_oom_object(context,
                                             "Application.ActiveExplorer");
  if (actExplorer)
    actExplorer->GetWindow (&ret);
  else
    {
      log_debug ("%s:%s: Could not find active window",
                 SRCNAME, __func__);
    }
  gpgol_release (actExplorer);
  TRETURN ret;
}

int
put_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *value)
{
  TSTART;
  LPDISPATCH propertyAccessor;
  VARIANT cVariant[2];
  VARIANT rVariant;
  DISPID dispid;
  DISPPARAMS dispparams;
  HRESULT hr;
  EXCEPINFO execpinfo;
  BSTR b_property;
  wchar_t *w_property;
  unsigned int argErr = 0;

  init_excepinfo (&execpinfo);

  log_oom ("%s:%s: Looking up property: %s;",
             SRCNAME, __func__, dasl_id);

  propertyAccessor = get_oom_object (pDisp, "PropertyAccessor");
  if (!propertyAccessor)
    {
      log_error ("%s:%s: Failed to look up property accessor.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  dispid = lookup_oom_dispid (propertyAccessor, "SetProperty");

  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find SetProperty DISPID",
               SRCNAME, __func__);
    TRETURN -1;
  }

  /* Prepare the parameter */
  w_property = utf8_to_wchar (dasl_id);
  b_property = SysAllocString (w_property);
  xfree (w_property);

  /* Variant 0 carries the data. */
  VariantInit (&cVariant[0]);
  if (VariantCopy (&cVariant[0], value))
    {
      log_error ("%s:%s: Falied to copy value.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  /* Variant 1 is the DASL as found out by experiments. */
  VariantInit (&cVariant[1]);
  cVariant[1].vt = VT_BSTR;
  cVariant[1].bstrVal = b_property;
  dispparams.rgvarg = cVariant;
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  VariantInit (&rVariant);

  hr = propertyAccessor->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                 DISPATCH_METHOD, &dispparams,
                                 &rVariant, &execpinfo, &argErr);
  VariantClear (&cVariant[0]);
  VariantClear (&cVariant[1]);
  gpgol_release (propertyAccessor);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: failure: invoking SetProperty p=%p vt=%d"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 rVariant.pdispVal, rVariant.vt, (unsigned int)hr,
                 (unsigned int)argErr);
      VariantClear (&rVariant);
      dump_excepinfo (execpinfo);
      TRETURN -1;
    }
  VariantClear (&rVariant);
  TRETURN 0;
}

int
put_pa_string (LPDISPATCH pDisp, const char *dasl_id, const char *value)
{
  TSTART;
  wchar_t *w_value = utf8_to_wchar (value);
  BSTR b_value = SysAllocString(w_value);
  xfree (w_value);
  VARIANT var;
  VariantInit (&var);
  var.vt = VT_BSTR;
  var.bstrVal = b_value;
  int ret = put_pa_variant (pDisp, dasl_id, &var);
  VariantClear (&var);
  TRETURN ret;
}

int
put_pa_int (LPDISPATCH pDisp, const char *dasl_id, int value)
{
  TSTART;
  VARIANT var;
  VariantInit (&var);
  var.vt = VT_I4;
  var.lVal = value;
  int ret = put_pa_variant (pDisp, dasl_id, &var);
  VariantClear (&var);
  TRETURN ret;
}

/* Get a MAPI property through OOM using the PropertyAccessor
 * interface and the DASL Uid. Returns -1 on error.
 * Variant has to be cleared with VariantClear.
 * rVariant must be a pointer to a Variant.
 */
int get_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *rVariant)
{
  TSTART;
  LPDISPATCH propertyAccessor;
  VARIANT cVariant[1];
  DISPID dispid;
  DISPPARAMS dispparams;
  HRESULT hr;
  EXCEPINFO execpinfo;
  BSTR b_property;
  wchar_t *w_property;
  unsigned int argErr = 0;

  init_excepinfo (&execpinfo);
  log_oom ("%s:%s: Looking up property: %s;",
             SRCNAME, __func__, dasl_id);

  propertyAccessor = get_oom_object (pDisp, "PropertyAccessor");
  if (!propertyAccessor)
    {
      log_error ("%s:%s: Failed to look up property accessor.",
                 SRCNAME, __func__);
      TRETURN -1;
    }

  dispid = lookup_oom_dispid (propertyAccessor, "GetProperty");

  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find GetProperty DISPID",
               SRCNAME, __func__);
    TRETURN -1;
  }

  /* Prepare the parameter */
  w_property = utf8_to_wchar (dasl_id);
  b_property = SysAllocString (w_property);
  xfree (w_property);

  cVariant[0].vt = VT_BSTR;
  cVariant[0].bstrVal = b_property;
  dispparams.rgvarg = cVariant;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  VariantInit (rVariant);

  hr = propertyAccessor->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                                 DISPATCH_METHOD, &dispparams,
                                 rVariant, &execpinfo, &argErr);
  SysFreeString (b_property);
  gpgol_release (propertyAccessor);
  if (hr != S_OK && strcmp (GPGOL_UID_DASL, dasl_id))
    {
      /* It often happens that mails don't have a uid by us e.g. if
         they are not crypto mails or just dont have one. This is
         not an error. */
      log_debug ("%s:%s: error: invoking GetProperty p=%p vt=%d"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 rVariant->pdispVal, rVariant->vt, (unsigned int)hr,
                 (unsigned int)argErr);
      dump_excepinfo (execpinfo);
      VariantClear (rVariant);
      TRETURN -1;
    }
  TRETURN 0;
}

/* Get a property string by using the PropertyAccessor of pDisp
 * Returns NULL on error or a newly allocated result. */
char *
get_pa_string (LPDISPATCH pDisp, const char *property)
{
  TSTART;
  VARIANT rVariant;
  char *result = NULL;

  if (get_pa_variant (pDisp, property, &rVariant))
    {
      TRETURN NULL;
    }

  if (rVariant.vt == VT_BSTR && rVariant.bstrVal)
    {
      result = wchar_to_utf8 (rVariant.bstrVal);
    }
  else if (rVariant.vt & VT_ARRAY && !(rVariant.vt & VT_BYREF))
    {
      LONG uBound, lBound;
      VARTYPE vt;
      char *data;
      SafeArrayGetVartype(rVariant.parray, &vt);

      if (SafeArrayGetUBound (rVariant.parray, 1, &uBound) != S_OK ||
          SafeArrayGetLBound (rVariant.parray, 1, &lBound) != S_OK ||
          vt != VT_UI1)
        {
          log_error ("%s:%s: Error: %i", SRCNAME, __func__, __LINE__);
          VariantClear (&rVariant);
          TRETURN NULL;
        }

      result = (char *)xmalloc (uBound - lBound + 1);
      data = (char *) rVariant.parray->pvData;
      memcpy (result, data + lBound, uBound - lBound);
      result[uBound - lBound] = '\0';
    }
  else if (rVariant.vt != 0)
    {
      log_debug ("%s:%s: Property `%s' is not a string (vt=%d)",
                 SRCNAME, __func__, property, rVariant.vt);
    }

  VariantClear (&rVariant);

  TRETURN result;
}

int
get_pa_int (LPDISPATCH pDisp, const char *property, int *rInt)
{
  TSTART;
  VARIANT rVariant;

  if (get_pa_variant (pDisp, property, &rVariant))
    {
      TRETURN -1;
    }

  if (rVariant.vt != VT_I4)
    {
      log_debug ("%s:%s: Property `%s' is not a int (vt=%d)",
                 SRCNAME, __func__, property, rVariant.vt);
      TRETURN -1;
    }

  *rInt = rVariant.lVal;

  VariantClear (&rVariant);
  TRETURN 0;
}

/* Helper for exchange address lookup. */
static char *
get_recipient_addr_entry_fallbacks_ex (LPDISPATCH addr_entry)
{
  TSTART;
  /* Maybe check for type here? We are pretty sure that we are exchange */

  /* According to MSDN Message Boards the PR_EMS_AB_PROXY_ADDRESSES_DASL
     is more avilable then the SMTP Address. */
  char *ret = get_pa_string (addr_entry, PR_EMS_AB_PROXY_ADDRESSES_DASL);
  if (ret)
    {
      log_debug ("%s:%s: Found recipient through AB_PROXY: %s",
                 SRCNAME, __func__, anonstr (ret));

      char *smtpbegin = strstr(ret, "SMTP:");
      if (smtpbegin == ret)
        {
          ret += 5;
        }
      TRETURN ret;
    }
  else
    {
      log_debug ("%s:%s: Failed AB_PROXY lookup.",
                 SRCNAME, __func__);
    }

  LPDISPATCH ex_user = get_oom_object (addr_entry, "GetExchangeUser");
  if (!ex_user)
    {
      log_debug ("%s:%s: Failed to find ExchangeUser",
                 SRCNAME, __func__);
      TRETURN nullptr;
    }

  ret = get_oom_string (ex_user, "PrimarySmtpAddress");
  gpgol_release (ex_user);
  if (ret)
    {
      log_debug ("%s:%s: Found recipient through exchange user primary smtp address: %s",
                 SRCNAME, __func__, anonstr (ret));
      TRETURN ret;
    }
  TRETURN nullptr;
}

/* Helper for additional fallbacks in recipient lookup */
static char *
get_recipient_addr_fallbacks (LPDISPATCH recipient)
{
  TSTART;
  if (!recipient)
    {
      TRETURN nullptr;
    }
  LPDISPATCH addr_entry = get_oom_object (recipient, "AddressEntry");

  if (!addr_entry)
    {
      log_debug ("%s:%s: Failed to find AddressEntry",
                 SRCNAME, __func__);
      TRETURN nullptr;
    }

  char *ret = get_recipient_addr_entry_fallbacks_ex (addr_entry);

  gpgol_release (addr_entry);

  TRETURN ret;
}

/* Try to resolve a recipient group and add it to the recipients vector.

   Returns true on success.
*/
static bool
try_resolve_group (LPDISPATCH addrEntry,
                   std::vector<std::pair<Recipient, shared_disp_t> >&ret,
                   int recipient_type)
{
  TSTART;
  /* Get the name for debugging */
  std::string name;
  char *cname = get_oom_string (addrEntry, "Name");
  if (cname)
    {
      name = cname;
    }
  xfree (cname);

  int user_type = get_oom_int (addrEntry, "AddressEntryUserType");

  if (user_type != DISTRIBUTION_LIST_ADDRESS_ENTRY_TYPE)
    {
      log_data ("%s:%s: type of %s is %i",
                       SRCNAME, __func__, anonstr (name.c_str()), user_type);
      TRETURN false;
    }

  LPDISPATCH members = get_oom_object (addrEntry, "Members");
  addrEntry = nullptr;

  if (!members)
    {
      TRACEPOINT;
      TRETURN false;
    }

  int count = get_oom_int (members, "Count");

  if (!count)
    {
      TRACEPOINT;
      gpgol_release (members);
      TRETURN false;
    }

  bool foundOne = false;
  for (int i = 1; i <= count; i++)
    {
      auto item_str = std::string("Item(") + std::to_string (i) + ")";
      auto entry = MAKE_SHARED (get_oom_object (members, item_str.c_str()));
      if (!entry)
        {
          TRACEPOINT;
          continue;
        }
      std::string entryName;
      char *entry_name = get_oom_string (entry.get(), "Name");
      if (entry_name)
        {
          entryName = entry_name;
          xfree (entry_name);
        }

      int subType = get_oom_int (entry.get(), "AddressEntryUserType");
      /* Resolve recursively, yeah fun. */
      if (subType == DISTRIBUTION_LIST_ADDRESS_ENTRY_TYPE)
        {
          log_debug ("%s:%s: recursive address entry %s",
                     SRCNAME, __func__,
                     anonstr (entryName.c_str()));
          if (try_resolve_group (entry.get(), ret, recipient_type))
            {
              foundOne = true;
              continue;
            }
        }

      std::pair<Recipient, shared_disp_t> element;
      element.second = entry;

      /* Resolve directly ? */
      char *addrtype = get_pa_string (entry.get(), PR_ADDRTYPE_DASL);
      if (addrtype && !strcmp (addrtype, "SMTP"))
        {
          xfree (addrtype);
          char *resolved = get_pa_string (entry.get(), PR_EMAIL_ADDRESS_DASL);
          if (resolved)
            {
              element.first = Recipient (resolved, entryName.c_str (),
                                         recipient_type);
              ret.push_back (element);
              foundOne = true;
              continue;
            }
        }
      xfree (addrtype);

      /* Resolve through Exchange API */
      char *ex_resolved = get_recipient_addr_entry_fallbacks_ex (entry.get());
      if (ex_resolved)
        {
          element.first = Recipient (ex_resolved, entryName.c_str (),
                                     recipient_type);
          ret.push_back (element);
          foundOne = true;
          continue;
        }

      log_debug ("%s:%s: failed to resolve name %s",
                 SRCNAME, __func__,
                 anonstr (entryName.c_str()));
    }
  gpgol_release (members);
  if (!foundOne)
    {
      log_debug ("%s:%s: failed to resolve group %s",
                 SRCNAME, __func__,
                 anonstr (name.c_str()));
    }
  TRETURN foundOne;
}

/* Get the recipient mbox addresses with the addrEntry
   object corresponding to the resolved address. */
std::vector<std::pair<Recipient, shared_disp_t> >
get_oom_recipients_with_addrEntry (LPDISPATCH recipients, bool *r_err)
{
  TSTART;
  int recipientsCnt = get_oom_int (recipients, "Count");
  std::vector<std::pair<Recipient, shared_disp_t> > ret;
  int i;

  if (!recipientsCnt)
    {
      TRETURN ret;
    }

  /* Get the recipients */
  for (i = 1; i <= recipientsCnt; i++)
    {
      char buf[16];
      LPDISPATCH recipient;
      snprintf (buf, sizeof (buf), "Item(%i)", i);
      recipient = get_oom_object (recipients, buf);
      if (!recipient)
        {
          /* Should be impossible */
          log_error ("%s:%s: could not find Item %i;",
                     SRCNAME, __func__, i);
          if (r_err)
            {
              *r_err = true;
            }
          break;
        }

      int recipient_type = get_oom_int (recipient, "Type");
      std::string entryName;
      char *entry_name = get_oom_string (recipient, "Name");
      if (entry_name)
        {
          entryName = entry_name;
          xfree (entry_name);
        }


      auto addrEntry = MAKE_SHARED (get_oom_object (recipient, "AddressEntry"));
      if (addrEntry && try_resolve_group (addrEntry.get (), ret,
                                          recipient_type))
        {
          log_debug ("%s:%s: Resolved recipient group",
                     SRCNAME, __func__);
          gpgol_release (recipient);
          continue;
        }

      std::pair<Recipient, shared_disp_t> entry;
      entry.second = addrEntry;

      char *resolved = get_pa_string (recipient, PR_SMTP_ADDRESS_DASL);
      if (resolved)
        {
          entry.first = Recipient (resolved, entryName.c_str (),
                                   recipient_type);
          entry.first.setIndex (i);
          xfree (resolved);
          gpgol_release (recipient);
          ret.push_back (entry);
          continue;
        }
      /* No PR_SMTP_ADDRESS first fallback */
      resolved = get_recipient_addr_fallbacks (recipient);
      if (resolved)
        {
          entry.first = Recipient (resolved, entryName.c_str (),
                                   recipient_type);
          entry.first.setIndex (i);
          xfree (resolved);
          gpgol_release (recipient);
          ret.push_back (entry);
          continue;
        }

      char *address = get_oom_string (recipient, "Address");
      gpgol_release (recipient);
      log_debug ("%s:%s: Failed to look up Address probably "
                 "EX addr is returned",
                 SRCNAME, __func__);
      if (address)
        {
          entry.first = Recipient (resolved, recipient_type);
          entry.first.setIndex (i);
          ret.push_back (entry);
          xfree (address);
        }
      else if (r_err)
        {
          *r_err = true;
        }
    }
  TRETURN ret;
}

/* Gets the resolved smtp addresses of the recpients. */
std::vector<Recipient>
get_oom_recipients (LPDISPATCH recipients, bool *r_err)
{
  TSTART;
  std::vector<Recipient> ret;
  for (const auto &pair: get_oom_recipients_with_addrEntry (recipients, r_err))
    {
      ret.push_back (pair.first);
    }
  TRETURN ret;
}

/* Add an attachment to the outlook dispatcher disp
   that has an Attachment property.
   inFile is the path to the attachment. Name is the
   name that should be used in outlook. */
int
add_oom_attachment (LPDISPATCH disp, const wchar_t* inFileW,
                    const wchar_t* displayName, std::string &r_error_str,
                    int *r_err_code)
{
  TSTART;
  LPDISPATCH attachments = get_oom_object (disp, "Attachments");

  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT vtResult;
  VARIANT aVariant[4];
  HRESULT hr;
  BSTR inFileB = nullptr,
       dispNameB = nullptr;
  unsigned int argErr = 0;
  EXCEPINFO execpinfo;

  init_excepinfo (&execpinfo);
  dispid = lookup_oom_dispid (attachments, "Add");

  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find attachment dispatcher",
               SRCNAME, __func__);
    TRETURN -1;
  }

  if (inFileW)
    {
      inFileB = SysAllocString (inFileW);
    }
  if (displayName)
    {
      dispNameB = SysAllocString (displayName);
    }

  dispparams.rgvarg = aVariant;

  /* Contrary to the documentation the Source is the last
     parameter and not the first. Additionally DisplayName
     is documented but gets ignored by Outlook since Outlook
     2003 */
  dispparams.rgvarg[0].vt = VT_BSTR; /* DisplayName */
  dispparams.rgvarg[0].bstrVal = dispNameB;
  dispparams.rgvarg[1].vt = VT_INT;  /* Position */
  dispparams.rgvarg[1].intVal = 1;
  dispparams.rgvarg[2].vt = VT_INT;  /* Type */
  dispparams.rgvarg[2].intVal = 1;
  dispparams.rgvarg[3].vt = VT_BSTR; /* Source */
  dispparams.rgvarg[3].bstrVal = inFileB;
  dispparams.cArgs = 4;
  dispparams.cNamedArgs = 0;
  VariantInit (&vtResult);
  hr = attachments->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                            DISPATCH_METHOD, &dispparams,
                            &vtResult, &execpinfo, &argErr);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: error: invoking Add p=%p vt=%d hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 vtResult.pdispVal, vtResult.vt, (unsigned int)hr,
                 (unsigned int)argErr);
      dump_excepinfo (execpinfo);

      if (r_err_code)
        {
          *r_err_code = (int) execpinfo.scode;
        }
      if (execpinfo.bstrDescription)
        {
          char *utf8Err = wchar_to_utf8 (execpinfo.bstrDescription);
          if (utf8Err)
            {
              r_error_str = utf8Err;
            }
          xfree (utf8Err);
        }
    }

  if (inFileB)
    SysFreeString (inFileB);
  if (dispNameB)
    SysFreeString (dispNameB);
  VariantClear (&vtResult);
  gpgol_release (attachments);

  TRETURN hr == S_OK ? 0 : -1;
}

LPDISPATCH
get_object_by_id (LPDISPATCH pDisp, REFIID id)
{
  TSTART;
  LPDISPATCH disp = NULL;

  if (!pDisp)
    {
      TRETURN NULL;
    }

  if (gpgol_queryInterface(pDisp, id, (void **)&disp) != S_OK)
    {
      TRETURN NULL;
    }
  TRETURN disp;
}

LPDISPATCH
get_strong_reference (LPDISPATCH mail)
{
  TSTART;
  VARIANT var;
  VariantInit (&var);
  DISPPARAMS args;
  VARIANT argvars[2];
  VariantInit (&argvars[0]);
  VariantInit (&argvars[1]);
  argvars[1].vt = VT_DISPATCH;
  argvars[1].pdispVal = mail;
  argvars[0].vt = VT_INT;
  argvars[0].intVal = 1;
  args.cArgs = 2;
  args.cNamedArgs = 0;
  args.rgvarg = argvars;
  LPDISPATCH ret = NULL;
  if (!invoke_oom_method_with_parms (
      GpgolAddin::get_instance()->get_application(),
      "GetObjectReference", &var, &args))
    {
      ret = var.pdispVal;
      log_oom ("%s:%s: Got strong ref %p for %p",
               SRCNAME, __func__, ret, mail);
      memdbg_addRef (ret);
    }
  else
    {
      log_error ("%s:%s: Failed to get strong ref.",
                 SRCNAME, __func__);
    }
  VariantClear (&var);
  TRETURN ret;
}

LPMESSAGE
get_oom_message (LPDISPATCH mailitem)
{
  TSTART;
  LPUNKNOWN mapi_obj = get_oom_iunknown (mailitem, "MapiObject");
  if (!mapi_obj)
    {
      log_error ("%s:%s: Failed to obtain MAPI Message.",
                 SRCNAME, __func__);
      TRETURN NULL;
    }
  TRETURN (LPMESSAGE) mapi_obj;
}

static LPMESSAGE
get_oom_base_message_from_mapi (LPDISPATCH mapi_message)
{
  TSTART;
  HRESULT hr;
  LPDISPATCH secureItem = NULL;
  LPMESSAGE message = NULL;
  LPMAPISECUREMESSAGE secureMessage = NULL;

  secureItem = get_object_by_id (mapi_message,
                                 IID_IMAPISecureMessage);
  if (!secureItem)
    {
      log_error ("%s:%s: Failed to obtain SecureItem.",
                 SRCNAME, __func__);
      TRETURN NULL;
    }

  secureMessage = (LPMAPISECUREMESSAGE) secureItem;

  /* The call to GetBaseMessage is pretty much a jump
     in the dark. So it would not be surprising to get
     crashes here in the future. */
  log_oom("%s:%s: About to call GetBaseMessage.",
                SRCNAME, __func__);
  hr = secureMessage->GetBaseMessage (&message);
  memdbg_addRef (message);
  gpgol_release (secureMessage);
  if (hr != S_OK)
    {
      log_error_w32 (hr, "Failed to GetBaseMessage.");
      TRETURN NULL;
    }

  TRETURN message;
}

LPMESSAGE
get_oom_base_message (LPDISPATCH mailitem)
{
  TSTART;
  LPMESSAGE mapi_message = get_oom_message (mailitem);
  LPMESSAGE ret = NULL;
  if (!mapi_message)
    {
      log_error ("%s:%s: Failed to obtain mapi_message.",
                 SRCNAME, __func__);
      TRETURN NULL;
    }
  ret = get_oom_base_message_from_mapi ((LPDISPATCH)mapi_message);
  gpgol_release (mapi_message);
  TRETURN ret;
}

static int
invoke_oom_method_with_parms_type (LPDISPATCH pDisp, const char *name,
                                   VARIANT *rVariant, DISPPARAMS *params,
                                   int type)
{
  TSTART;
  HRESULT hr;
  DISPID dispid;

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid != DISPID_UNKNOWN)
    {
      EXCEPINFO execpinfo;
      init_excepinfo (&execpinfo);
      DISPPARAMS dispparams = {NULL, NULL, 0, 0};

      hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                          type, params ? params : &dispparams,
                          rVariant, &execpinfo, NULL);
      if (hr != S_OK)
        {
          log_debug ("%s:%s: Method '%s' invokation failed: %#lx",
                     SRCNAME, __func__, name, hr);
          dump_excepinfo (execpinfo);
          TRETURN -1;
        }
    }

  TRETURN 0;
}

int
invoke_oom_method_with_parms (LPDISPATCH pDisp, const char *name,
                              VARIANT *rVariant, DISPPARAMS *params)
{
  TSTART;
  int ret = invoke_oom_method_with_parms_type (pDisp, name, rVariant, params,
                                              DISPATCH_METHOD);
  TRETURN ret;
}

int
invoke_oom_method (LPDISPATCH pDisp, const char *name, VARIANT *rVariant)
{
  TSTART;
  int ret = invoke_oom_method_with_parms (pDisp, name, rVariant, NULL);
  TRETURN ret;
}

LPMAPISESSION
get_oom_mapi_session ()
{
  TSTART;
  LPDISPATCH application = GpgolAddin::get_instance ()->get_application ();
  LPDISPATCH oom_session = NULL;
  LPMAPISESSION session = NULL;
  LPUNKNOWN mapiobj = NULL;
  HRESULT hr;

  if (!application)
    {
      log_debug ("%s:%s: Not implemented for Ol < 14", SRCNAME, __func__);
      TRETURN NULL;
    }

  oom_session = get_oom_object (application, "Session");
  if (!oom_session)
    {
      log_error ("%s:%s: session object not found", SRCNAME, __func__);
      TRETURN NULL;
    }
  mapiobj = get_oom_iunknown (oom_session, "MAPIOBJECT");
  gpgol_release (oom_session);

  if (!mapiobj)
    {
      log_error ("%s:%s: error getting Session.MAPIOBJECT", SRCNAME, __func__);
      TRETURN NULL;
    }
  session = NULL;
  hr = gpgol_queryInterface (mapiobj, IID_IMAPISession, (void**)&session);
  gpgol_release (mapiobj);
  if (hr != S_OK || !session)
    {
      log_error ("%s:%s: error getting IMAPISession: hr=%#lx",
                 SRCNAME, __func__, hr);
      TRETURN NULL;
    }
  TRETURN session;
}

int
create_category (LPDISPATCH categories, const char *category, int color)
{
  TSTART;
  VARIANT cVariant[3];
  VARIANT rVariant;
  DISPID dispid;
  DISPPARAMS dispparams;
  HRESULT hr;
  EXCEPINFO execpinfo;
  BSTR b_name;
  wchar_t *w_name;
  unsigned int argErr = 0;

  init_excepinfo (&execpinfo);

  if (!categories || !category)
    {
      TRACEPOINT;
      TRETURN 1;
    }

  dispid = lookup_oom_dispid (categories, "Add");
  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find Add DISPID",
               SRCNAME, __func__);
    TRETURN -1;
  }

  /* Do the string dance */
  w_name = utf8_to_wchar (category);
  b_name = SysAllocString (w_name);
  xfree (w_name);

  /* Variants are in reverse order
     ShortcutKey -> 0 / Int
     Color -> 1 / Int
     Name -> 2 / Bstr */
  VariantInit (&cVariant[2]);
  cVariant[2].vt = VT_BSTR;
  cVariant[2].bstrVal = b_name;

  VariantInit (&cVariant[1]);
  cVariant[1].vt = VT_INT;
  cVariant[1].intVal = color;

  VariantInit (&cVariant[0]);
  cVariant[0].vt = VT_INT;
  cVariant[0].intVal = 0;

  dispparams.cArgs = 3;
  dispparams.cNamedArgs = 0;
  dispparams.rgvarg = cVariant;

  VariantInit (&rVariant);
  hr = categories->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                           DISPATCH_METHOD, &dispparams,
                           &rVariant, &execpinfo, &argErr);
  VariantClear (&cVariant[0]);
  VariantClear (&cVariant[1]);
  VariantClear (&cVariant[2]);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: error: invoking Add p=%p vt=%d"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 rVariant.pdispVal, rVariant.vt, (unsigned int)hr,
                 (unsigned int)argErr);
      dump_excepinfo (execpinfo);
      VariantClear (&rVariant);
      TRETURN -1;
    }
  VariantClear (&rVariant);
  log_oom ("%s:%s: Created category '%s'",
             SRCNAME, __func__, anonstr (category));
  TRETURN 0;
}

LPDISPATCH
get_store_for_id (const char *storeID)
{
  TSTART;
  LPDISPATCH application = GpgolAddin::get_instance ()->get_application ();
  if (!application || !storeID)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH stores = get_oom_object (application, "Session.Stores");
  if (!stores)
    {
      log_error ("%s:%s: No stores found.",
                 SRCNAME, __func__);
      TRETURN nullptr;
    }
  auto store_count = get_oom_int (stores, "Count");

  for (int n = 1; n <= store_count; n++)
    {
      const auto store_str = std::string("Item(") + std::to_string(n) + ")";
      LPDISPATCH store = get_oom_object (stores, store_str.c_str());

      if (!store)
        {
          TRACEPOINT;
          continue;
        }
      char *id = get_oom_string (store, "StoreID");
      if (id && !strcmp (id, storeID))
        {
          gpgol_release (stores);
          xfree (id);
          return store;
        }
      xfree (id);
      gpgol_release (store);
    }
  gpgol_release (stores);
  TRETURN nullptr;
}

void
ensure_category_exists (const char *category, int color)
{
  TSTART;
  LPDISPATCH application = GpgolAddin::get_instance ()->get_application ();
  if (!application || !category)
    {
      TRACEPOINT;
      TRETURN;
    }

  log_oom ("%s:%s: Ensure category exists called for %s, %i",
           SRCNAME, __func__,
           category, color);

  LPDISPATCH stores = get_oom_object (application, "Session.Stores");
  if (!stores)
    {
      log_error ("%s:%s: No stores found.",
                 SRCNAME, __func__);
      TRETURN;
    }
  auto store_count = get_oom_int (stores, "Count");

  for (int n = 1; n <= store_count; n++)
    {
      const auto store_str = std::string("Item(") + std::to_string(n) + ")";
      LPDISPATCH store = get_oom_object (stores, store_str.c_str());

      if (!store)
        {
          TRACEPOINT;
          continue;
        }

      LPDISPATCH categories = get_oom_object (store, "Categories");
      gpgol_release (store);
      if (!categories)
        {
          categories = get_oom_object (application, "Session.Categories");
          if (!categories)
            {
              TRACEPOINT;
              continue;
            }
        }

      auto count = get_oom_int (categories, "Count");
      bool found = false;
      for (int i = 1; i <= count && !found; i++)
        {
          const auto item_str = std::string("Item(") + std::to_string(i) + ")";
          LPDISPATCH category_obj = get_oom_object (categories, item_str.c_str());
          if (!category_obj)
            {
              TRACEPOINT;
              gpgol_release (categories);
              break;
            }
          char *name = get_oom_string (category_obj, "Name");
          if (name && !strcmp (category, name))
            {
              log_oom ("%s:%s: Found category '%s'",
                       SRCNAME, __func__, name);
              found = true;
            }
          /* We don't check the color here as the user may change that. */
          gpgol_release (category_obj);
          xfree (name);
        }

      if (!found)
        {
          if (create_category (categories, category, color))
            {
              log_oom ("%s:%s: Found category '%s'",
                       SRCNAME, __func__, category);
            }
        }
      /* Otherwise we have to create the category */
      gpgol_release (categories);
    }
  gpgol_release (stores);
  TRETURN;
}

int
add_category (LPDISPATCH mail, const char *category)
{
  TSTART;
  char *tmp = get_oom_string (mail, "Categories");
  if (!tmp)
    {
      TRACEPOINT;
      TRETURN 1;
    }

  if (strstr (tmp, category))
    {
      log_oom ("%s:%s: category '%s' already added.",
               SRCNAME, __func__, category);
      TRETURN 0;
    }

  std::string newstr (tmp);
  xfree (tmp);
  if (!newstr.empty ())
    {
      newstr += CategoryManager::getSeperator () + std::string (" ");
    }
  newstr += category;

  TRETURN put_oom_string (mail, "Categories", newstr.c_str ());
}

int
remove_category (LPDISPATCH mail, const char *category, bool exactMatch)
{
  TSTART;
  char *tmp = get_oom_string (mail, "Categories");
  if (!tmp)
    {
      TRACEPOINT;
      TRETURN 1;
    }

  std::vector<std::string> categories;
  std::istringstream f(tmp);
  std::string s;
  const std::string sep = CategoryManager::getSeperator();
  while (std::getline(f, s, *(sep.c_str())))
    {
      ltrim(s);
      categories.push_back(s);
    }
  xfree (tmp);

  const std::string categoryStr = category;

  categories.erase (std::remove_if (categories.begin(),
                                    categories.end(),
                                    [categoryStr, exactMatch] (const std::string &cat)
    {
      if (exactMatch)
        {
          return cat == categoryStr;
        }
      return cat.compare (0, categoryStr.size(), categoryStr) == 0;
    }), categories.end ());
  std::string newCategories;
  std::string newsep = sep + " ";
  join (categories, newsep.c_str (), newCategories);

  TRETURN put_oom_string (mail, "Categories", newCategories.c_str ());
}

static int
_delete_category (LPDISPATCH categories, int idx)
{
  TSTART;
  VARIANT aVariant[1];
  DISPPARAMS dispparams;

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = idx;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  TRETURN invoke_oom_method_with_parms (categories, "Remove", NULL,
                                        &dispparams);
}

int
delete_category (LPDISPATCH store, const char *category)
{
  TSTART;
  if (!store || !category)
    {
      TRETURN -1;
    }

  LPDISPATCH categories = get_oom_object (store, "Categories");
  if (!categories)
    {
      categories = get_oom_object (
                      GpgolAddin::get_instance ()->get_application (),
                      "Session.Categories");
      if (!categories)
        {
          TRACEPOINT;
          TRETURN -1;
        }
    }

  auto count = get_oom_int (categories, "Count");
  int ret = 0;
  for (int i = 1; i <= count; i++)
    {
      const auto item_str = std::string("Item(") + std::to_string(i) + ")";
      LPDISPATCH category_obj = get_oom_object (categories, item_str.c_str());
      if (!category_obj)
        {
          TRACEPOINT;
          gpgol_release (categories);
          break;
        }
      char *name = get_oom_string (category_obj, "Name");
      gpgol_release (category_obj);
      if (name && !strcmp (category, name))
        {
          if ((ret = _delete_category (categories, i)))
            {
              log_error ("%s:%s: Failed to delete category '%s'",
                         SRCNAME, __func__, anonstr (category));
            }
          else
            {
              log_debug ("%s:%s: Deleted category '%s'",
                         SRCNAME, __func__, anonstr (category));
            }
          xfree (name);
          break;
        }
      xfree (name);
    }

  gpgol_release (categories);
  TRETURN ret;
}

void
delete_all_categories_starting_with (const char *string)
{
  LPDISPATCH application = GpgolAddin::get_instance ()->get_application ();
  if (!application || !string)
    {
      TRACEPOINT;
      TRETURN;
    }

  log_oom ("%s:%s: Delete categories starting with: \"%s\"",
           SRCNAME, __func__, string);

  LPDISPATCH stores = get_oom_object (application, "Session.Stores");
  if (!stores)
    {
      log_error ("%s:%s: No stores found.",
                 SRCNAME, __func__);
      TRETURN;
    }

  auto store_count = get_oom_int (stores, "Count");

  for (int n = 1; n <= store_count; n++)
    {
      const auto store_str = std::string("Item(") + std::to_string(n) + ")";
      LPDISPATCH store = get_oom_object (stores, store_str.c_str());

      if (!store)
        {
          TRACEPOINT;
          continue;
        }

      LPDISPATCH categories = get_oom_object (store, "Categories");
      if (!categories)
        {
          categories = get_oom_object (application, "Session.Categories");
          if (!categories)
            {
              TRACEPOINT;
              gpgol_release (store);
              continue;
            }
        }

      auto count = get_oom_int (categories, "Count");
      std::vector<std::string> to_delete;
      for (int i = 1; i <= count; i++)
        {
          const auto item_str = std::string("Item(") + std::to_string(i) + ")";
          LPDISPATCH category_obj = get_oom_object (categories, item_str.c_str());
          if (!category_obj)
            {
              TRACEPOINT;
              gpgol_release (categories);
              break;
            }
          char *name = get_oom_string (category_obj, "Name");
          if (name && !strncmp (string, name, strlen (string)))
            {
              log_oom ("%s:%s: Found category for deletion '%s'",
                       SRCNAME, __func__, anonstr(name));
              to_delete.push_back (name);
            }
          /* We don't check the color here as the user may change that. */
          gpgol_release (category_obj);
          xfree (name);
        }

      /* Do this one after another to avoid messing with indexes. */
      for (const auto &str: to_delete)
        {
          delete_category (store, str.c_str ());
        }

      gpgol_release (store);
      /* Otherwise we have to create the category */
      gpgol_release (categories);
    }
  gpgol_release (stores);
  TRETURN;

}

static char *
generate_uid ()
{
  TSTART;
  UUID uuid;
  UuidCreate (&uuid);

  unsigned char *str;
  UuidToStringA (&uuid, &str);

  char *ret = xstrdup ((char*)str);
  RpcStringFreeA (&str);

  TRETURN ret;
}

char *
get_unique_id (LPDISPATCH mail, int create, const char *uuid)
{
  TSTART;
  if (!mail)
    {
      TRETURN NULL;
    }

  /* Get the User Properties. */
  if (!create)
    {
      char *uid = get_pa_string (mail, GPGOL_UID_DASL);
      if (!uid)
        {
          log_debug ("%s:%s: No uuid found in oom for '%p' looking for PR_INTERNET_MESSAGE_ID_W_DASL",
                     SRCNAME, __func__, mail);
          uid = get_pa_string(mail, PR_INTERNET_MESSAGE_ID_W_DASL);
        }
      if (!uid)
        {
          log_debug ("%s:%s: No uuid found in oom for '%p'",
                     SRCNAME, __func__, mail);
          TRETURN NULL;
        }
      else
        {
          log_debug ("%s:%s: Found uid '%s' for '%p'",
                     SRCNAME, __func__, uid, mail);
          TRETURN uid;
        }
    }
  char *newuid;
  if (!uuid)
    {
      newuid = generate_uid ();
    }
  else
    {
      newuid = xstrdup (uuid);
    }
  int ret = put_pa_string (mail, GPGOL_UID_DASL, newuid);

  if (ret)
    {
      log_debug ("%s:%s: failed to set uid '%s' for '%p'",
                 SRCNAME, __func__, newuid, mail);
      xfree (newuid);
      TRETURN NULL;
    }


  log_debug ("%s:%s: '%p' has now the uid: '%s' ",
             SRCNAME, __func__, mail, newuid);
  TRETURN newuid;
}

char *
reset_unique_id (LPDISPATCH mail)
{
  TSTART;
  char *newuid = generate_uid ();

  int ret = put_pa_string (mail, GPGOL_UID_DASL, newuid);
  if (ret)
    {
      log_debug ("%s:%s: failed to set uid '%s' for '%p'",
                 SRCNAME, __func__, newuid, mail);
      xfree (newuid);
      TRETURN NULL;
    }
  TRETURN newuid;
}

std::string
get_unique_id_s (LPDISPATCH mail, int create, const char *uuid)
{
  char *val = get_unique_id (mail, create, uuid);
  if (val)
    {
      return val;
    }
  return std::string ();
}

HWND
get_active_hwnd ()
{
  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH activeWindow = get_oom_object (app, "ActiveWindow");
  if (!activeWindow)
    {
      activeWindow = get_oom_object (app, "ActiveInspector");
      if (!activeWindow)
        {
          activeWindow = get_oom_object (app, "ActiveExplorer");
          if (!activeWindow)
            {
              TRACEPOINT;
              TRETURN nullptr;
            }
        }
    }

  /* Both explorer and inspector have this. */
  char *caption = get_oom_string (activeWindow, "Caption");
  gpgol_release (activeWindow);
  if (!caption)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }
  /* Might not be completly true for multiple explorers
     on the same folder but good enugh. */
  HWND hwnd = FindWindowExA(NULL, NULL, "rctrl_renwnd32",
                            caption);
  xfree (caption);

  TRETURN hwnd;
}

LPDISPATCH
create_mail ()
{
  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  VARIANT var;
  VariantInit (&var);
  VARIANT argvars[1];
  DISPPARAMS args;
  VariantInit (&argvars[0]);
  argvars[0].vt = VT_I2;
  argvars[0].intVal = 0;
  args.cArgs = 1;
  args.cNamedArgs = 0;
  args.rgvarg = argvars;

  LPDISPATCH ret = nullptr;

  if (invoke_oom_method_with_parms (app, "CreateItem", &var, &args))
    {
      log_error ("%s:%s: Failed to create mailitem.",
                 SRCNAME, __func__);
      TRETURN ret;
    }

  ret = var.pdispVal;
  TRETURN ret;
}

LPDISPATCH
get_account_for_mail (const char *mbox)
{
  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH accounts = get_oom_object (app, "Session.Accounts");

  if (!accounts)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  int count = get_oom_int (accounts, "Count");
  for (int i = 1; i <= count; i++)
    {
      std::string item = std::string ("Item(") + std::to_string (i) + ")";

      LPDISPATCH account = get_oom_object (accounts, item.c_str ());

      if (!account)
        {
          TRACEPOINT;
          continue;
        }
      char *smtpAddr = get_oom_string (account, "SmtpAddress");

      if (!smtpAddr)
        {
          gpgol_release (account);
          TRACEPOINT;
          continue;
        }
      if (!stricmp (mbox, smtpAddr))
        {
          gpgol_release (accounts);
          xfree (smtpAddr);
          TRETURN account;
        }
      gpgol_release (account);
      xfree (smtpAddr);
    }
  gpgol_release (accounts);

  log_error ("%s:%s: Failed to find account for '%s'.",
             SRCNAME, __func__, anonstr (mbox));

  TRETURN nullptr;
}

char *
get_sender_SendUsingAccount (LPDISPATCH mailitem, bool *r_is_GSuite)
{
  TSTART;
  LPDISPATCH sender = get_oom_object (mailitem, "SendUsingAccount");
  if (!sender)
    {
      TRETURN nullptr;
    }

  char *buf = get_oom_string (sender, "SmtpAddress");
  char *dispName = get_oom_string (sender, "DisplayName");
  gpgol_release (sender);

  /* Check for G Suite account */
  if (dispName && !strcmp ("G Suite", dispName) && r_is_GSuite)
    {
      *r_is_GSuite = true;
    }
  xfree (dispName);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s: found sender", SRCNAME, __func__);
      TRETURN buf;
    }
  xfree (buf);
  TRETURN nullptr;
}

char *
get_sender_Sender (LPDISPATCH mailitem)
{
  TSTART;
  LPDISPATCH sender = get_oom_object (mailitem, "Sender");
  if (!sender)
    {
      TRETURN nullptr;
    }
  char *buf = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
  gpgol_release (sender);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 2", SRCNAME, __func__);
      TRETURN buf;
    }
  xfree (buf);
  /* We have a sender object but not yet an smtp address likely
     exchange. Try some more propertys of the message. */
  buf = get_pa_string (mailitem, PR_TAG_SENDER_SMTP_ADDRESS);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 3", SRCNAME, __func__);
      TRETURN buf;
    }
  xfree (buf);
  buf = get_pa_string (mailitem, PR_TAG_RECEIVED_REPRESENTING_SMTP_ADDRESS);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 4", SRCNAME, __func__);
      TRETURN buf;
    }
  xfree (buf);
  TRETURN nullptr;
}

char *
get_sender_primary_send_acct (LPDISPATCH mailitem)
{
  TSTART;
  char *buf = get_pa_string (mailitem, PR_PRIMARY_SEND_ACCT_W_DASL);
  if (buf && strlen (buf))
    {
      /* The format of this is documented as implementation dependent
         for exchange this looks like
         AccountNumber\01ExchangeAddress\01SMTPAddress */
      char *last = strrchr (buf, 1);
      if (last && ++last)
        {
          char *atChar = strchr (last, '@');
          if (atChar)
            {
              log_debug ("%s:%s Sender fallback 5", SRCNAME, __func__);
              size_t len = strlen (last) + 1;
              char *ret = (char *)xmalloc (len);
              strcpy_s (ret, len, last);
              xfree (buf);
              TRETURN ret;
            }
          else
            {
              log_dbg ("Last part does not contain @ character: %s",
                       anonstr (last));
            }
        }
      log_dbg ("Failed to parse %s", anonstr (buf));
    }
  xfree (buf);
  TRETURN nullptr;
}

char *
get_sender_CurrentUser (LPDISPATCH mailitem)
{
  TSTART;
  LPDISPATCH sender = get_oom_object (mailitem,
                                      "Session.CurrentUser");
  if (!sender)
    {
      TRETURN nullptr;
    }
  char *buf = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
  gpgol_release (sender);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 6", SRCNAME, __func__);
      TRETURN buf;
    }
  xfree (buf);
  TRETURN nullptr;
}

char *
get_sender_SenderEMailAddress (LPDISPATCH mailitem)
{
  TSTART;

  char *type = get_oom_string (mailitem, "SenderEmailType");
  if (type && !strcmp ("SMTP", type))
    {
      char *senderMail = get_oom_string (mailitem, "SenderEmailAddress");
      if (senderMail)
        {
          log_debug ("%s:%s: Sender found", SRCNAME, __func__);
          xfree (type);
          TRETURN senderMail;
        }
    }
  xfree (type);
  TRETURN nullptr;
}

char *
get_sender_SentRepresentingAddress (LPDISPATCH mailitem)
{
  TSTART;
  char *buf = get_pa_string (mailitem,
                             PR_SENT_REPRESENTING_EMAIL_ADDRESS_W_DASL);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Found sent representing address \"%s\"",
                 SRCNAME, __func__, anonstr (buf));
      TRETURN buf;
    }
  xfree (buf);
  TRETURN nullptr;
}

char *
get_inline_body ()
{
  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();
  if (!app)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH explorer = get_oom_object (app, "ActiveExplorer");

  if (!explorer)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH inlineResponse = get_oom_object (explorer, "ActiveInlineResponse");
  gpgol_release (explorer);

  if (!inlineResponse)
    {
      TRETURN nullptr;
    }

  char *body = get_oom_string (inlineResponse, "Body");
  gpgol_release (inlineResponse);

  TRETURN body;
}

int
get_ex_major_version_for_addr (const char *mbox)
{
  TSTART;
  LPDISPATCH account = get_account_for_mail (mbox);
  if (!account)
    {
      TRACEPOINT;
      TRETURN -1;
    }

  char *version_str = get_oom_string (account, "ExchangeMailboxServerVersion");
  gpgol_release (account);

  if (!version_str)
    {
      TRETURN -1;
    }

  log_debug ("%s:%s: Detected exchange major version: %s",
             SRCNAME, __func__, version_str);
  long int version = strtol (version_str, nullptr, 10);
  xfree (version_str);

  TRETURN (int) version;
}

int
get_ol_ui_language ()
{
  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance()->get_application();
  if (!app)
    {
      TRACEPOINT;
      TRETURN 0;
    }

  LPDISPATCH langSettings = get_oom_object (app, "LanguageSettings");
  if (!langSettings)
    {
      TRACEPOINT;
      TRETURN 0;
    }

  VARIANT var;
  VariantInit (&var);

  VARIANT aVariant[1];
  DISPPARAMS dispparams;

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = 2;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  int ret = invoke_oom_method_with_parms_type (langSettings, "LanguageID", &var,
                                               &dispparams,
                                               DISPATCH_PROPERTYGET);
  gpgol_release (langSettings);
  if (ret)
    {
      TRACEPOINT;
      TRETURN 0;
    }
  if (var.vt != VT_INT && var.vt != VT_I4)
    {
      TRACEPOINT;
      TRETURN 0;
    }

  int result = var.intVal;

  VariantClear (&var);
  TRETURN result;
}

void
log_addins ()
{
  if (!opt.enable_debug)
    {
      return;
    }

  TSTART;
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      TRETURN;
    }

  LPDISPATCH addins = get_oom_object (app, "COMAddins");

  if (!addins)
    {
      TRACEPOINT;
      TRETURN;
    }

  std::string activeAddins;
  int count = get_oom_int (addins, "Count");
  for (int i = 1; i <= count; i++)
    {
      VARIANT aVariant[1];
      VARIANT rVariant;

      VariantInit (&rVariant);
      DISPPARAMS dispparams;

      dispparams.rgvarg = aVariant;
      dispparams.rgvarg[0].vt = VT_INT;
      dispparams.rgvarg[0].intVal = i;
      dispparams.cArgs = 1;
      dispparams.cNamedArgs = 0;

      /* We need this instead of get_oom_object item(1) as usual becase
         the item method accepts a string or an int. String would
         be the ProgID and int is just the index. So Fun. */
      if (invoke_oom_method_with_parms_type (addins, "Item", &rVariant,
                                             &dispparams,
                                             DISPATCH_METHOD |
                                             DISPATCH_PROPERTYGET))
        {
          log_error ("%s:%s: Failed to invoke item func.",
                     SRCNAME, __func__);
          continue;
        }

      if (rVariant.vt != (VT_DISPATCH))
        {
          log_error ("%s:%s: Invalid ret val",
                     SRCNAME, __func__);
          continue;
        }

      LPDISPATCH addin = rVariant.pdispVal;

      if (!addin)
        {
          TRACEPOINT;
          continue;
        }
      memdbg_addRef (addin);
      bool connected = get_oom_bool (addin, "Connect");
      if (!connected)
        {
          gpgol_release (addin);
          continue;
        }

      std::string progId = get_oom_string_s (addin, "ProgId");
      std::string description = get_oom_string_s (addin, "Description");
      gpgol_release (addin);

      if (progId.empty ())
        {
          progId = "Unknown ProgID";
        }
      if (description.empty ())
        {
          description = "No description";
        }
      activeAddins += progId + " (" + description + ")"  + "\n";
    }
  gpgol_release (addins);

  log_debug ("%s:%s:Active Addins:\n%s", SRCNAME, __func__,
             activeAddins.c_str ());
  TRETURN;
}

bool
is_preview_pane_visible (LPDISPATCH explorer)
{
  TSTART;
  if (!explorer)
    {
      TRACEPOINT;
      TRETURN false;
    }
  VARIANT var;
  VariantInit (&var);
  VARIANT argvars[1];
  DISPPARAMS args;
  VariantInit (&argvars[0]);
  argvars[0].vt = VT_INT;
  argvars[0].intVal = 3;
  args.cArgs = 1;
  args.cNamedArgs = 0;
  args.rgvarg = argvars;

  if (invoke_oom_method_with_parms (explorer, "IsPaneVisible", &var, &args))
    {
      log_error ("%s:%s: Failed to check visibilty.",
                 SRCNAME, __func__);
      TRETURN false;
    }

  if (var.vt != VT_BOOL)
    {
      TRACEPOINT;
      TRETURN false;
    }
  TRETURN !!var.boolVal;
}

static LPDISPATCH
add_user_prop (LPDISPATCH user_props, const char *name)
{
  TSTART;
  if (!user_props || !name)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  wchar_t *w_name = utf8_to_wchar (name);
  BSTR b_name = SysAllocString (w_name);
  xfree (w_name);

  /* Args:
    0: DisplayFormat int OlUserPropertyType
    1: AddToFolderFields Bool Should the filed be added to the folder.
    2: Type int OlUserPropertyType Type of the field.
    3: Name Bstr Name of the field.

    Returns the added Property.
  */
  VARIANT var;
  VariantInit (&var);
  DISPPARAMS args;
  VARIANT argvars[4];
  VariantInit (&argvars[0]);
  VariantInit (&argvars[1]);
  VariantInit (&argvars[2]);
  VariantInit (&argvars[3]);
  argvars[0].vt = VT_INT;
  argvars[0].intVal = 1; // 1 means text.
  argvars[1].vt = VT_BOOL;
  argvars[1].boolVal = VARIANT_FALSE;
  argvars[2].vt = VT_INT;
  argvars[2].intVal = 1;
  argvars[3].vt = VT_BSTR;
  argvars[3].bstrVal = b_name;
  args.cArgs = 4;
  args.cNamedArgs = 0;
  args.rgvarg = argvars;

  int res = invoke_oom_method_with_parms (user_props, "Add", &var, &args);
  VariantClear (&argvars[0]);
  VariantClear (&argvars[1]);
  VariantClear (&argvars[2]);
  VariantClear (&argvars[3]);

  if (res)
    {
      log_oom ("%s:%s: Failed to add property %s.",
               SRCNAME, __func__, name);
      TRETURN nullptr;
    }

  if (var.vt != VT_DISPATCH)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH ret = var.pdispVal;
  memdbg_addRef (ret);

  TRETURN ret;
}

LPDISPATCH
find_user_prop (LPDISPATCH user_props, const char *name)
{
  TSTART;
  if (!user_props || !name)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }
  VARIANT var;
  VariantInit (&var);

  wchar_t *w_name = utf8_to_wchar (name);
  BSTR b_name = SysAllocString (w_name);
  xfree (w_name);

  /* Name -> 1 / Bstr
     Custom 0 -> Bool True for search in custom properties. False
                 for builtin properties. */
  DISPPARAMS args;
  VARIANT argvars[2];
  VariantInit (&argvars[0]);
  VariantInit (&argvars[1]);
  argvars[1].vt = VT_BSTR;
  argvars[1].bstrVal = b_name;
  argvars[0].vt = VT_BOOL;
  argvars[0].boolVal = VARIANT_TRUE;
  args.cArgs = 2;
  args.cNamedArgs = 0;
  args.rgvarg = argvars;

  int res = invoke_oom_method_with_parms (user_props, "Find", &var, &args);
  VariantClear (&argvars[0]);
  VariantClear (&argvars[1]);
  if (res)
    {
      log_oom ("%s:%s: Failed to find property %s.",
               SRCNAME, __func__, name);
      TRETURN nullptr;
    }
  if (var.vt != VT_DISPATCH)
    {
      TRACEPOINT;
      TRETURN nullptr;
    }

  LPDISPATCH ret = var.pdispVal;
  memdbg_addRef (ret);

  TRETURN ret;
}

LPDISPATCH
find_or_add_text_prop (LPDISPATCH user_props, const char *name)
{
  TSTART;
  LPDISPATCH ret = find_user_prop (user_props, name);

  if (ret)
    {
      TRETURN ret;
    }

  ret = add_user_prop (user_props, name);

  TRETURN ret;
}

void
release_disp (LPDISPATCH obj)
{
  TSTART;
  gpgol_release (obj);
  TRETURN;
}

enum FolderID
{
  olFolderCalendar = 9,
  olFolderConflicts = 19,
  olFolderContacts = 10,
  olFolderDeletedItems = 3,
  olFolderDrafts = 16,
  olFolderInbox = 6,
  olFolderJournal = 11,
  olFolderJunk = 23,
  olFolderLocalFailures = 21,
  olFolderManagedEmail = 29,
  olFolderNotes = 12,
  olFolderOutbox = 4,
  olFolderSentMail = 5,
  olFolderServerFailures = 22,
  olFolderSuggestedContacts = 30,
  olFolderSyncIssues = 20,
  olFolderTasks = 13,
  olFolderToDo = 28,
  olPublicFoldersAllPublicFolders = 18,
  olFolderRssFeeds = 25,
};

static bool
is_mail_in_folder (LPDISPATCH mailitem, int folder)
{
  TSTART;
  if (!mailitem)
    {
      STRANGEPOINT;
      TRETURN false;
    }

  auto store = MAKE_SHARED (get_oom_object (mailitem, "Parent.Store"));

  if (!store)
    {
      log_debug ("%s:%s: Mail has no parent folder. Probably unsafed",
                 SRCNAME, __func__);
      TRETURN false;
    }

  std::string tmp = std::string("GetDefaultFolder(") + std::to_string (folder) +
                    std::string(")");
  auto target_folder = MAKE_SHARED (get_oom_object (store.get(),
                                                    tmp.c_str()));

  if (!target_folder)
    {
      STRANGEPOINT;
      TRETURN false;
    }

  auto mail_folder = MAKE_SHARED (get_oom_object (mailitem, "Parent"));

  if (!mail_folder)
    {
      STRANGEPOINT;
      TRETURN false;
    }

  char *target_id = get_oom_string (target_folder.get(), "entryID");
  if (!target_id)
    {
      STRANGEPOINT;
      TRETURN false;
    }
  char *folder_id = get_oom_string (mail_folder.get(), "entryID");
  if (!folder_id)
    {
      STRANGEPOINT;
      free (target_id);
      TRETURN false;
    }

  bool ret = !strcmp (target_id, folder_id);

  xfree (target_id);
  xfree (folder_id);
  TRETURN ret;
}

bool
is_junk_mail (LPDISPATCH mailitem)
{
  TSTART;
  TRETURN is_mail_in_folder (mailitem, FolderID::olFolderJunk);
}

bool
is_draft_mail (LPDISPATCH mailitem)
{
  TSTART;
  TRETURN is_mail_in_folder (mailitem, FolderID::olFolderDrafts);
}

void
format_variant (std::stringstream &stream, VARIANT* var)
{
  if (!var)
    {
      stream << " (null) ";
    }
  stream << "VT: " << std::hex << var->vt << " Value: ";

  VARTYPE vt = var->vt;

  if (vt == VT_BOOL)
    {
      stream << (var->boolVal == VARIANT_FALSE ? "false" : "true");
    }
  else if (vt == (VT_BOOL | VT_BYREF))
    {
      stream << (*(var->pboolVal) == VARIANT_FALSE ? "false" : "true");
    }
  else if (vt == VT_BSTR)
    {
      char *buf = wchar_to_utf8 (var->bstrVal);
      stream << "BStr: " << buf;
      xfree (buf);
    }
  else if (vt == VT_INT || vt == VT_I4)
    {
      stream << var->intVal;
    }
  else if (vt == VT_DISPATCH)
    {
      char *buf = get_object_name ((LPUNKNOWN) var->pdispVal);
      stream << "IDispatch: " << buf;
      xfree (buf);
    }
  else if (vt == (VT_VARIANT | VT_BYREF))
    {
      format_variant (stream, var->pvarVal);
    }
  else
    {
      stream << "?";
    }
  stream << std::endl;
}

std::string
format_dispparams (DISPPARAMS *p)
{
  if (!p)
    {
      return "(null)";
    }
  std::stringstream stream;
  stream << "Count: " << p->cArgs << " CNamed: " << p->cNamedArgs << std::endl;

  for (int i = 0; i < p->cArgs; i++)
    {
      format_variant (stream, p->rgvarg + i);
    }
  return stream.str ();
}

int
count_visible_attachments (LPDISPATCH attachments)
{
  int ret = 0;

  if (!attachments)
    {
      return 0;
    }

  int att_count = get_oom_int (attachments, "Count");
  for (int i = 1; i <= att_count; i++)
    {
      std::string item_str;
      item_str = std::string("Item(") + std::to_string (i) + ")";
      LPDISPATCH oom_attach = get_oom_object (attachments, item_str.c_str ());
      if (!oom_attach)
        {
          log_error ("%s:%s: Failed to get attachment.",
                     SRCNAME, __func__);
          continue;
        }
      VARIANT var;
      VariantInit (&var);
      if (get_pa_variant (oom_attach, PR_ATTACHMENT_HIDDEN_DASL, &var))
        {
          /* SECURITY: Testing has shown that for all mail types GpgOL
             handles that might contain an unsigned attachment we always
             have the MAPIOBJECT / get the hidden state. Only the transient
             MIME attachments. The ones used for the MAPI to MIME conversion
             and which are hidden by GpgOL and handled by GpgOL will have
             no MAPIOBJECT when a mail is opened from file. So this will
             remove the warning that "smime.p7m" or "gpgol_mime_structure.txt"
             are unsigned and unencrypted attachments. */
          log_dbg ("Failed to get hidden state.");
          LPUNKNOWN mapiobj = get_oom_iunknown (oom_attach, "MAPIOBJECT");
          if (!mapiobj)
            {
              const auto dispName = get_oom_string_s (oom_attach,
                                                      "DisplayName");
              log_dbg ("Attachment: %s has no mapiobject. Ignoring it.",
                       anonstr (dispName.c_str ()));
            }
          else
            {
              gpgol_release (mapiobj);
              const auto dispName = get_oom_string_s (oom_attach,
                                                      "DisplayName");
              log_dbg ("Attachment %s without hidden state but mapiobj. "
                       "Count as visible.", anonstr (dispName.c_str ()));
              ret++;
            }
        }
      else if (var.vt == VT_BOOL && var.boolVal == VARIANT_FALSE)
        {
           ret++;
        }
      gpgol_release (oom_attach);
      VariantClear (&var);
    }
  return ret;
}

int invoke_oom_method_with_int (LPDISPATCH pDisp, const char *name,
                                int arg,
                                VARIANT *rVariant)
{
  TSTART;
  DISPPARAMS parms;
  VARIANT argvars[1];
  VariantInit (&argvars[0]);
  argvars[0].vt = VT_INT;
  argvars[0].intVal = arg;
  parms.cArgs = 1;
  parms.cNamedArgs = 0;
  parms.rgvarg = argvars;

  TRETURN invoke_oom_method_with_parms (pDisp, name,
                                        rVariant, &parms);

}

int invoke_oom_method_with_string (LPDISPATCH pDisp, const char *name,
                                   const char *arg,
                                   VARIANT *rVariant)
{
  TSTART;
  if (!arg)
    {
      TRETURN 0;
    }
  wchar_t *warg = utf8_to_wchar (arg);
  if (!warg)
    {
      TRETURN 1;
    }
  VARIANT aVariant[1];
  VariantInit (&aVariant[0]);
  aVariant[0].vt = VT_BSTR;
  aVariant[0].bstrVal = SysAllocString (warg);
  xfree (warg);

  DISPPARAMS dispparams;
  dispparams.rgvarg = aVariant;
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  int ret = invoke_oom_method_with_parms (pDisp, name, rVariant, &dispparams);
  VariantClear(&aVariant[0]);
  TRETURN ret;
}

int
set_oom_recipients (LPDISPATCH item, const std::vector<Recipient> &recps)
{
  if (!item)
    {
      STRANGEPOINT;
      TRETURN -1;
    }

  auto oom_recps = MAKE_SHARED (get_oom_object (item, "Recipients"));

  if (!oom_recps)
    {
      STRANGEPOINT;
      TRETURN -1;
    }

  int count = get_oom_int (oom_recps.get (), "Count");

  for (int i = 1; i <= count; i++)
    {
      /* First clear out the current recipients. */
      int ret = invoke_oom_method_with_int (oom_recps.get (),
                                            "Remove", 1,
                                            nullptr);

      if (ret)
        {
          STRANGEPOINT;
          TRETURN ret;
        }
    }

  for (const auto &recp: recps)
    {
      if (recp.type() == Recipient::olOriginator)
        {
          /* Skip the originator, we only add it internally but
             it does not need to be in OOM. */
          continue;
        }
      VARIANT result;
      VariantInit (&result);
      int ret = invoke_oom_method_with_string (oom_recps.get (), "Add",
                                               recp.mbox ().c_str (),
                                               &result);
      if (ret)
        {
          log_err ("Failed to add recipient.");
          TRETURN ret;
        }
      if (result.vt != VT_DISPATCH || !result.pdispVal)
        {
          log_err ("No recipient result.");
          continue;
        }
      if (put_oom_int (result.pdispVal, "Type", recp.type()))
        {
          log_err ("Failed to set recipient type.");
        }
      /* This releases the recipient. */
      VariantClear (&result);
    }
  TRETURN 0;
}

int
remove_oom_recipient (LPDISPATCH item, const std::string &mbox)
{
  TSTART;
  if (!item)
    {
      STRANGEPOINT;
      TRETURN -1;
    }

  auto oom_recps = MAKE_SHARED (get_oom_object (item, "Recipients"));

  if (!oom_recps)
    {
      STRANGEPOINT;
      TRETURN -1;
    }

  bool r_err = false;
  const auto recps = get_oom_recipients (oom_recps.get (), &r_err);
  if (r_err)
    {
      log_debug ("Failure to lookup recipients via OOM");
      TRETURN -1;
    }

  for (const auto &recp: recps)
    {
      if (recp.mbox () == mbox && recp.index () != -1)
        {
          TRETURN invoke_oom_method_with_int (oom_recps.get (),
                                              "Remove", recp.index (),
                                              nullptr);
        }
    }
  TRETURN -1;
}

void
oom_dump_idispatch (LPDISPATCH obj)
{
  log_dbg ("Start infos about %p", obj);
  if (!obj)
    {
      log_dbg ("It's NULL");
      return;
    }
  log_dbg ("Name: '%s'", get_object_name_s (obj).c_str ());

  LPTYPEINFO typeinfo = nullptr;
  HRESULT hr = obj->GetTypeInfo (0, 0, &typeinfo);

  if (!typeinfo || FAILED (hr))
    {
      log_dbg ("No typeinfo.");
      return;
    }

  TYPEATTR* pta = NULL;
  hr = typeinfo->GetTypeAttr(&pta);

  if (!pta || FAILED (hr))
    {
      log_dbg ("No type attr");
      return;
    }

  /* First the IID to have it clear */
  LPOLESTR lpsz = NULL;
  hr = StringFromIID(pta->guid, &lpsz);
  if(FAILED (hr))
    {
      hr = StringFromCLSID (pta->guid, &lpsz);
    }

  if(SUCCEEDED (hr))
    {
      log_dbg ("Interface: %S", lpsz);
      CoTaskMemFree(lpsz);
    }

  FUNCDESC *pfd = nullptr;
  /* Lets see what functions we have. */
  for(int i = 0; i < pta->cFuncs; i++)
    {
      typeinfo->GetFuncDesc(i, &pfd);

      BSTR names[1];
      unsigned int dumb;

      typeinfo->GetNames(pfd->memid, names, 1, &dumb);
      if (!names[0])
        {
          typeinfo->ReleaseFuncDesc(pfd);
          continue;
        }

      log_dbg ("%i: %S id=0x%li With %d param(s)\n", i,
               (names[0]), pfd->memid, pfd->cParams);
      typeinfo->ReleaseFuncDesc(pfd);
      SysFreeString(names[0]);
    }
  typeinfo->ReleaseTypeAttr(pta);

  /* Now for some interesting object relations
     that many oom objecs have. */
  const char * relations[] = {
    "Parent",
    "GetInspector",
    "Session",
    "Sender",
    nullptr
  };

  for (int i = 0; relations [i]; i++)
    {
      LPDISPATCH rel = get_oom_object (obj, relations[i]);
      log_dbg ("%s: %s", relations [i],
               get_object_name_s (rel).c_str ());
      gpgol_release (rel);
    }

  /* Now for some interesting string values. */
  const char * stringVals[] = {
    "EntryID",
    "Subject",
    "MessageClass",
    "Body",
    nullptr
  };

  for (int i = 0; stringVals[i]; i++)
    {
      const auto str = get_oom_string_s (obj, stringVals[i]);
      log_dbg ("%s: %s", stringVals[i], str.c_str ());
    }

  log_dbg ("Object dump done");
  return;
}

int
get_oom_crypto_flags (LPDISPATCH mailitem)
{
  TSTART;
  int r_val = 0;
  int err = get_pa_int (mailitem, PR_SECURITY_FLAGS_DASL, &r_val);
  if (err)
    {
      log_dbg ("Failed to get security flags.");
      TRETURN 0;
    }
  TRETURN r_val;
}

shared_disp_t
show_folder_select ()
{
  TSTART;
  VARIANT var;
  VariantInit (&var);
  LPDISPATCH rVal;

  auto namespace_obj = get_oom_object_s (oApp (), "Session");
  if (!namespace_obj)
    {
      STRANGEPOINT;
      TRETURN nullptr;
    }

  if (!invoke_oom_method (namespace_obj.get (), "PickFolder", &var))
    {
      if (!(var.vt & VT_DISPATCH))
        {
          log_dbg ("Failed to get disp obj. No folder selected?");
          TRETURN nullptr;
        }
      rVal = var.pdispVal;
      log_oom ("%s:%s: Got folder ref %p",
               SRCNAME, __func__, rVal);
      memdbg_addRef (rVal);
      TRETURN MAKE_SHARED (rVal);
    }
  log_dbg ("No folder returned.");
  TRETURN nullptr;
}

LPDISPATCH
oApp ()
{
  return GpgolAddin::get_instance()->get_application();
}

BSTR utf8_to_bstr (const char *string)
{
  wchar_t *tmp = utf8_to_wchar (string);
  BSTR bstring = tmp ? SysAllocString (tmp) : NULL;
  xfree (tmp);
  if (!bstring)
    {
      log_error_w32 (-1, "%s:%s: SysAllocString failed", SRCNAME, __func__);
      TRETURN nullptr;
    }
  TRETURN bstring;
}

int
oom_save_as (LPDISPATCH obj, const char *path, oomSaveAsType type)
{
  if (!obj || !path)
    {
      /* invalid arguments */
      STRANGEPOINT;
      TRETURN -1;
    }
  /* Params are first path and then type as optional. With
     COM Marshalling this means that param 1 is the path
     and 0 is the type. */
  VARIANT aVariant[2];
  VariantInit(aVariant);
  VariantInit(aVariant + 1);

  DISPPARAMS dispparams;
  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = (int) type;
  dispparams.rgvarg[1].vt = VT_BSTR;
  dispparams.rgvarg[1].bstrVal = utf8_to_bstr (path);
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;

  int rc = invoke_oom_method_with_parms (obj, "SaveAs", nullptr, &dispparams);

  if (rc)
    {
      log_err ("Failed to call SaveAs");
    }
  VariantClear(aVariant);
  VariantClear(aVariant + 1);

  return rc;
}

int
oom_save_as_file (LPDISPATCH obj, const char *path)
{
  if (!obj || !path)
    {
      /* invalid arguments */
      STRANGEPOINT;
      TRETURN -1;
    }
  VARIANT aVariant[1];
  VariantInit(aVariant);

  DISPPARAMS dispparams;
  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BSTR;
  dispparams.rgvarg[0].bstrVal = utf8_to_bstr (path);
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;

  int rc = invoke_oom_method_with_parms (obj, "SaveAsFile",
                                         nullptr, &dispparams);

  if (rc)
    {
      log_err ("Failed to call SaveAsFile");
    }
  VariantClear(aVariant);

  return rc;
}

void
oom_clear_selections ()
{
  TSTART;
  auto explorers_obj = get_oom_object_s (oApp (), "Explorers");

  if (!explorers_obj)
    {
      STRANGEPOINT;
      TRETURN;
    }

  int count = get_oom_int (explorers_obj.get (), "Count");

  for (int i = 1; i <= count; i++)
    {
      auto item_str = std::string("Item(") + std::to_string (i) + ")";
      auto explorer = get_oom_object_s (explorers_obj, item_str.c_str ());
      if (!explorer)
        {
          STRANGEPOINT;
          TRETURN;
        }
      if (invoke_oom_method (explorer.get (), "ClearSelection", NULL))
        {
          log_err ("Clearing Explorers %i", i);
        }
    }
  TRETURN;
}

/* Adding/Removing a mailitem to the explorer selection results
   in a SelectionChange event
*/
void
oom_toggle_selection (LPDISPATCH mailitem, bool doAdd)
{
  TSTART;
  auto explorers_obj = get_oom_object_s (oApp(), "Explorers");

  if (!explorers_obj)
    {
      STRANGEPOINT;
      TRETURN;
    }

  int count = get_oom_int (explorers_obj.get (), "Count");

  if (count >0)
    {
      auto explorer = get_oom_object_s (explorers_obj, "Item(1)");
      if (!explorer)
        {
          STRANGEPOINT;
          TRETURN;
        }

      log_debug("%s:%s: Use explorer %p ",
                SRCNAME, __func__, explorer.get());

      if (doAdd)
        {
           VARIANT var;
            VariantInit (&var);
            DISPPARAMS params;
            VARIANT argvars[1];
            VariantInit (&argvars[0]);
            argvars[0].vt = VT_DISPATCH;
            argvars[0].pdispVal = mailitem;
            params.cArgs = 1;
            params.cNamedArgs = 0;
            params.rgvarg = argvars;
            invoke_oom_method_with_parms (explorer.get (), "AddToSelection", NULL, &params);

            log_debug("%s:%s: Add mailitem %p to selection",
                      SRCNAME, __func__, argvars[0].pdispVal);

        }
      else
        {
          auto selection_obj = get_oom_object_s (explorer, "Selection");
          if (!selection_obj)
            {
              STRANGEPOINT;
              TRETURN;
            }

          int selCount = get_oom_int (selection_obj.get (), "Count");
          auto mailEntryID = get_oom_string_s (mailitem, "EntryID");

          for (int i = 1; i <= count; i++)
            {
              auto sel_item_str = std::string("Item(") + std::to_string (i) + ")";
              auto selMail = get_oom_object_s (selection_obj, sel_item_str.c_str ());
              auto selEntryID = get_oom_string_s (selMail, "EntryID");
              if (selEntryID == mailEntryID)
              {
                VARIANT var;
                VariantInit (&var);
                DISPPARAMS params;
                VARIANT argvars[1];
                VariantInit (&argvars[0]);
                argvars[0].vt = VT_DISPATCH;
                argvars[0].pdispVal = selMail.get();
                params.cArgs = 1;
                params.cNamedArgs = 0;
                params.rgvarg = argvars;

                if (!invoke_oom_method_with_parms (explorer.get (), "RemoveFromSelection", NULL, &params))
                  {
                    log_debug("%s:%s: Remove mailitem %p from selection",
                         SRCNAME, __func__, argvars[0].pdispVal);
                    break;
                  }
              }
            }
        }
    }
  TRETURN;
}
