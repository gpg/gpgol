/* oomhelp.cpp - Helper functions for the Outlook Object Model
 * Copyright (C) 2009 g10 Code GmbH
 * Copyright (C) 2015 by Bundesamt für Sicherheit in der Informationstechnik
 * Software engineering by Intevation GmbH
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
#include <rpc.h>

#include "common.h"

#include "oomhelp.h"
#include "gpgoladdin.h"


/* Return a malloced string with the utf-8 encoded name of the object
   or NULL if not available.  */
char *
get_object_name (LPUNKNOWN obj)
{
  HRESULT hr;
  LPDISPATCH disp = NULL;
  LPTYPEINFO tinfo = NULL;
  BSTR bstrname;
  char *name = NULL;

  if (!obj)
    goto leave;

  obj->QueryInterface (IID_IDispatch, (void **)&disp);
  if (!disp)
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
    gpgol_release (tinfo);
  if (disp)
    gpgol_release (disp);

  return name;
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
    return DISPID_UNKNOWN; /* Error: Invalid arg.  */

  wname = utf8_to_wchar (name);
  if (!wname)
    return DISPID_UNKNOWN;/* Error:  Out of memory.  */

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
      return;
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
  log_debug ("%s:%s: Exception: \n"
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
      if (pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp) != S_OK)
        {
          log_error ("%s:%s Object does not support IDispatch",
                     SRCNAME, __func__);
          gpgol_release (pObj);
          return NULL;
        }
      /* Confirmed through testing that the retval needs a release */
      if (pObj != pStart)
        gpgol_release (pObj);
      pObj = NULL;
      if (!pDisp)
        return NULL;  /* The object has no IDispatch interface.  */
      if (!*fullname)
        {
          log_oom ("%s:%s:         got %p",SRCNAME, __func__, pDisp);
          return pDisp; /* Ready.  */
        }
      
      /* Break out the next name part.  */
      {
        const char *dot;
        size_t n;
        
        dot = strchr (fullname, '.');
        if (dot == fullname)
          {
            gpgol_release (pDisp);
            return NULL;  /* Empty name part: error.  */
          }
        else if (dot)
          n = dot - fullname;
        else
          n = strlen (fullname);
        
        if (n >= sizeof name)
          n = sizeof name - 1;
        strncpy (name, fullname, n);
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
                  return NULL; /* Error:  Out of memory.  */
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
          return NULL;  /* Name not found.  */
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
          return NULL;  /* Invoke failed.  */
        }

      pObj = vtResult.pdispVal;
    }
  log_debug ("%s:%s: no object", SRCNAME, __func__);
  return NULL;
}


/* Helper for put_oom_icon.  */
static int
put_picture_or_mask (LPDISPATCH pDisp, int resource, int size, int is_mask)
{
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
      return -1;
    }

  /* Wrap the image into an OLE object.  */
  hr = OleCreatePictureIndirect (&pdesc, IID_IPictureDisp, 
                                 TRUE, (void **) &pPict);
  if (hr != S_OK || !pPict)
    {
      log_error ("%s:%s: OleCreatePictureIndirect failed: hr=%#lx\n",
                 SRCNAME, __func__, hr);
      return -1;
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
      return -1;
    }
  return 0;
}


/* Update the icon of PDISP using the bitmap with RESOURCE ID.  The
   function adds the system pixel size to the resource id to compute
   the actual icon size.  The resource id of the mask is the N+1.  */
int
put_oom_icon (LPDISPATCH pDisp, int resource_id, int size)
{
  int rc;

  /* This code is only relevant for Outlook < 2010.
    Ideally it should grab the system pixel size and use an
    icon of the appropiate size (e.g. 32 or 64px)
  */

  rc = put_picture_or_mask (pDisp, resource_id, size, 0);
  if (!rc)
    rc = put_picture_or_mask (pDisp, resource_id + 1, size, 1);

  return rc;
}


/* Set the boolean property NAME to VALUE.  */
int
put_oom_bool (LPDISPATCH pDisp, const char *name, int value)
{
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    return -1;

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
      return -1;
    }
  return 0;
}


/* Set the property NAME to VALUE.  */
int
put_oom_int (LPDISPATCH pDisp, const char *name, int value)
{
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];

  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    return -1;

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
      return -1;
    }
  return 0;
}


/* Set the property NAME to STRING.  */
int
put_oom_string (LPDISPATCH pDisp, const char *name, const char *string)
{
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
    return -1;

  {
    wchar_t *tmp = utf8_to_wchar (string);
    bstring = tmp? SysAllocString (tmp):NULL;
    xfree (tmp);
    if (!bstring)
      {
        log_error_w32 (-1, "%s:%s: SysAllocString failed", SRCNAME, __func__);
        return -1;
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
      return -1;
    }
  return 0;
}

/* Set the property NAME to DISP.  */
int
put_oom_disp (LPDISPATCH pDisp, const char *name, LPDISPATCH disp)
{
  HRESULT hr;
  DISPID dispid_put = DISPID_PROPERTYPUT;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[1];
  EXCEPINFO execpinfo;

  init_excepinfo (&execpinfo);
  dispid = lookup_oom_dispid (pDisp, name);
  if (dispid == DISPID_UNKNOWN)
    return -1;

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
      return -1;
    }
  return 0;
}

/* Get the boolean property NAME of the object PDISP.  Returns False if
   not found or if it is not a boolean property.  */
int
get_oom_bool (LPDISPATCH pDisp, const char *name)
{
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

  return result;
}


/* Get the integer property NAME of the object PDISP.  Returns 0 if
   not found or if it is not an integer property.  */
int
get_oom_int (LPDISPATCH pDisp, const char *name)
{
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

  return result;
}


/* Get the string property NAME of the object PDISP.  Returns NULL if
   not found or if it is not a string property.  */
char *
get_oom_string (LPDISPATCH pDisp, const char *name)
{
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

  return result;
}


/* Get the object property NAME of the object PDISP.  Returns NULL if
   not found or if it is not an object perty.  */
LPUNKNOWN
get_oom_iunknown (LPDISPATCH pDisp, const char *name)
{
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
        return rVariant.punkVal;

      VariantClear (&rVariant);
    }

  return NULL;
}


/* Return the control object described by the tag property with value
   TAG. The object POBJ must support the FindControl method.  Returns
   NULL if not found.  */
LPDISPATCH
get_oom_control_bytag (LPDISPATCH pDisp, const char *tag)
{
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
      return NULL;
    }

  {
    wchar_t *tmp = utf8_to_wchar (tag);
    bstring = tmp? SysAllocString (tmp):NULL;
    xfree (tmp);
    if (!bstring)
      {
        log_error_w32 (-1, "%s:%s: SysAllocString failed", SRCNAME, __func__);
        return NULL;
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
      rVariant.pdispVal->QueryInterface (IID_IDispatch, (LPVOID*)&result);
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

  return result;
}


/* Add a new button to an object which supports the add method.
   Returns the new object or NULL on error.  */
LPDISPATCH
add_oom_button (LPDISPATCH pObj)
{
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
      return NULL;
    }
  return rVariant.pdispVal;
}


/* Add a new button to an object which supports the add method.
   Returns the new object or NULL on error.  */
void
del_oom_button (LPDISPATCH pObj)
{
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
}

/* Gets the current contexts HWND. Returns NULL on error */
HWND
get_oom_context_window (LPDISPATCH context)
{
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
  return ret;
}

int
put_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *value)
{
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
      return -1;
    }

  dispid = lookup_oom_dispid (propertyAccessor, "SetProperty");

  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find SetProperty DISPID",
               SRCNAME, __func__);
    return -1;
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
      return -1;
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
      log_debug ("%s:%s: error: invoking SetProperty p=%p vt=%d"
                 " hr=0x%x argErr=0x%x",
                 SRCNAME, __func__,
                 rVariant.pdispVal, rVariant.vt, (unsigned int)hr,
                 (unsigned int)argErr);
      VariantClear (&rVariant);
      dump_excepinfo (execpinfo);
      return -1;
    }
  VariantClear (&rVariant);
  return 0;
}

int
put_pa_string (LPDISPATCH pDisp, const char *dasl_id, const char *value)
{
  wchar_t *w_value = utf8_to_wchar (value);
  BSTR b_value = SysAllocString(w_value);
  VARIANT var;
  VariantInit (&var);
  var.vt = VT_BSTR;
  var.bstrVal = b_value;
  int ret = put_pa_variant (pDisp, dasl_id, &var);
  VariantClear (&var);
  return ret;
}

int
put_pa_int (LPDISPATCH pDisp, const char *dasl_id, int value)
{
  VARIANT var;
  VariantInit (&var);
  var.vt = VT_INT;
  var.intVal = value;
  int ret = put_pa_variant (pDisp, dasl_id, &var);
  VariantClear (&var);
  return ret;
}

/* Get a MAPI property through OOM using the PropertyAccessor
 * interface and the DASL Uid. Returns -1 on error.
 * Variant has to be cleared with VariantClear.
 * rVariant must be a pointer to a Variant.
 */
int get_pa_variant (LPDISPATCH pDisp, const char *dasl_id, VARIANT *rVariant)
{
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
      return -1;
    }

  dispid = lookup_oom_dispid (propertyAccessor, "GetProperty");

  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find GetProperty DISPID",
               SRCNAME, __func__);
    return -1;
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
      return -1;
    }
  return 0;
}

/* Get a property string by using the PropertyAccessor of pDisp
 * returns NULL on error or a newly allocated result. */
char *
get_pa_string (LPDISPATCH pDisp, const char *property)
{
  VARIANT rVariant;
  char *result = NULL;

  if (get_pa_variant (pDisp, property, &rVariant))
    {
      return NULL;
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
          return NULL;
        }

      result = (char *)xmalloc (uBound - lBound + 1);
      data = (char *) rVariant.parray->pvData;
      memcpy (result, data + lBound, uBound - lBound);
      result[uBound - lBound] = '\0';
    }
  else
    {
      log_debug ("%s:%s: Property `%s' is not a string (vt=%d)",
                 SRCNAME, __func__, property, rVariant.vt);
    }

  VariantClear (&rVariant);

  return result;
}

int
get_pa_int (LPDISPATCH pDisp, const char *property, int *rInt)
{
  VARIANT rVariant;

  if (get_pa_variant (pDisp, property, &rVariant))
    {
      return -1;
    }

  if (rVariant.vt != VT_I4)
    {
      log_debug ("%s:%s: Property `%s' is not a int (vt=%d)",
                 SRCNAME, __func__, property, rVariant.vt);
      return -1;
    }

  *rInt = rVariant.lVal;

  VariantClear (&rVariant);
  return 0;
}

/* Helper for additional fallbacks in recipient lookup */
static char *
get_recipient_addr_fallbacks (LPDISPATCH recipient)
{
  if (!recipient)
    {
      return nullptr;
    }
  LPDISPATCH addr_entry = get_oom_object (recipient, "AddressEntry");

  if (!addr_entry)
    {
      log_debug ("%s:%s: Failed to find AddressEntry",
                 SRCNAME, __func__);
      return nullptr;
    }

  /* Maybe check for type here? We are pretty sure that we are exchange */

  /* According to MSDN Message Boards the PR_EMS_AB_PROXY_ADDRESSES_DASL
     is more avilable then the SMTP Address. */
  char *ret = get_pa_string (addr_entry, PR_EMS_AB_PROXY_ADDRESSES_DASL);
  if (ret)
    {
      log_debug ("%s:%s: Found recipient through AB_PROXY: %s",
                 SRCNAME, __func__, ret);

      char *smtpbegin = strstr(ret, "SMTP:");
      if (smtpbegin == ret)
        {
          ret += 5;
        }
      gpgol_release (addr_entry);
      return ret;
    }
  else
    {
      log_debug ("%s:%s: Failed AB_PROXY lookup.",
                 SRCNAME, __func__);
    }

  LPDISPATCH ex_user = get_oom_object (addr_entry, "GetExchangeUser");
  gpgol_release (addr_entry);
  if (!ex_user)
    {
      log_debug ("%s:%s: Failed to find ExchangeUser",
                 SRCNAME, __func__);
      return nullptr;
    }

  ret = get_oom_string (ex_user, "PrimarySmtpAddress");
  gpgol_release (ex_user);
  if (ret)
    {
      log_debug ("%s:%s: Found recipient through exchange user primary smtp address: %s",
                 SRCNAME, __func__, ret);
      return ret;
    }

  return nullptr;
}

/* Gets the resolved smtp addresses of the recpients. */
std::vector<std::string>
get_oom_recipients (LPDISPATCH recipients, bool *r_err)
{
  int recipientsCnt = get_oom_int (recipients, "Count");
  std::vector<std::string> ret;
  int i;

  if (!recipientsCnt)
    {
      return ret;
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
      char *resolved = get_pa_string (recipient, PR_SMTP_ADDRESS_DASL);
      if (resolved)
        {
          ret.push_back (resolved);
          xfree (resolved);
          gpgol_release (recipient);
          continue;
        }
      /* No PR_SMTP_ADDRESS first fallback */
      resolved = get_recipient_addr_fallbacks (recipient);
      if (resolved)
        {
          ret.push_back (resolved);
          xfree (resolved);
          gpgol_release (recipient);
          continue;
        }

      char *address = get_oom_string (recipient, "Address");
      gpgol_release (recipient);
      log_debug ("%s:%s: Failed to look up Address probably EX addr is returned",
                 SRCNAME, __func__);
      if (address)
        {
          ret.push_back (address);
          xfree (address);
        }
      else if (r_err)
        {
          *r_err = true;
        }
    }

  return ret;
}

/* Add an attachment to the outlook dispatcher disp
   that has an Attachment property.
   inFile is the path to the attachment. Name is the
   name that should be used in outlook. */
int
add_oom_attachment (LPDISPATCH disp, const wchar_t* inFileW,
                    const wchar_t* displayName)
{
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
    return -1;
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
    }

  if (inFileB)
    SysFreeString (inFileB);
  if (dispNameB)
    SysFreeString (dispNameB);
  VariantClear (&vtResult);
  gpgol_release (attachments);

  return hr == S_OK ? 0 : -1;
}

LPDISPATCH
get_object_by_id (LPDISPATCH pDisp, REFIID id)
{
  LPDISPATCH disp = NULL;

  if (!pDisp)
    return NULL;

  if (pDisp->QueryInterface (id, (void **)&disp) != S_OK)
    return NULL;
  return disp;
}

LPDISPATCH
get_strong_reference (LPDISPATCH mail)
{
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
    }
  else
    {
      log_error ("%s:%s: Failed to get strong ref.",
                 SRCNAME, __func__);
    }
  VariantClear (&var);
  return ret;
}

LPMESSAGE
get_oom_message (LPDISPATCH mailitem)
{
  LPUNKNOWN mapi_obj = get_oom_iunknown (mailitem, "MapiObject");
  if (!mapi_obj)
    {
      log_error ("%s:%s: Failed to obtain MAPI Message.",
                 SRCNAME, __func__);
      return NULL;
    }
  return (LPMESSAGE) mapi_obj;
}

static LPMESSAGE
get_oom_base_message_from_mapi (LPDISPATCH mapi_message)
{
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
      return NULL;
    }

  secureMessage = (LPMAPISECUREMESSAGE) secureItem;

  /* The call to GetBaseMessage is pretty much a jump
     in the dark. So it would not be surprising to get
     crashes here in the future. */
  log_oom_extra("%s:%s: About to call GetBaseMessage.",
                SRCNAME, __func__);
  hr = secureMessage->GetBaseMessage (&message);
  gpgol_release (secureMessage);
  if (hr != S_OK)
    {
      log_error_w32 (hr, "Failed to GetBaseMessage.");
      return NULL;
    }

  return message;
}

LPMESSAGE
get_oom_base_message (LPDISPATCH mailitem)
{
  LPMESSAGE mapi_message = get_oom_message (mailitem);
  LPMESSAGE ret = NULL;
  if (!mapi_message)
    {
      log_error ("%s:%s: Failed to obtain mapi_message.",
                 SRCNAME, __func__);
      return NULL;
    }
  ret = get_oom_base_message_from_mapi ((LPDISPATCH)mapi_message);
  gpgol_release (mapi_message);
  return ret;
}

static int
invoke_oom_method_with_parms_type (LPDISPATCH pDisp, const char *name,
                                   VARIANT *rVariant, DISPPARAMS *params,
                                   int type)
{
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
          return -1;
        }
    }

  return 0;
}

int
invoke_oom_method_with_parms (LPDISPATCH pDisp, const char *name,
                              VARIANT *rVariant, DISPPARAMS *params)
{
  return invoke_oom_method_with_parms_type (pDisp, name, rVariant, params,
                                            DISPATCH_METHOD);
}

int
invoke_oom_method (LPDISPATCH pDisp, const char *name, VARIANT *rVariant)
{
  return invoke_oom_method_with_parms (pDisp, name, rVariant, NULL);
}

LPMAPISESSION
get_oom_mapi_session ()
{
  LPDISPATCH application = GpgolAddin::get_instance ()->get_application ();
  LPDISPATCH oom_session = NULL;
  LPMAPISESSION session = NULL;
  LPUNKNOWN mapiobj = NULL;
  HRESULT hr;

  if (!application)
    {
      log_debug ("%s:%s: Not implemented for Ol < 14", SRCNAME, __func__);
      return NULL;
    }

  oom_session = get_oom_object (application, "Session");
  if (!oom_session)
    {
      log_error ("%s:%s: session object not found", SRCNAME, __func__);
      return NULL;
    }
  mapiobj = get_oom_iunknown (oom_session, "MAPIOBJECT");
  gpgol_release (oom_session);

  if (!mapiobj)
    {
      log_error ("%s:%s: error getting Session.MAPIOBJECT", SRCNAME, __func__);
      return NULL;
    }
  session = NULL;
  hr = mapiobj->QueryInterface (IID_IMAPISession, (void**)&session);
  gpgol_release (mapiobj);
  if (hr != S_OK || !session)
    {
      log_error ("%s:%s: error getting IMAPISession: hr=%#lx",
                 SRCNAME, __func__, hr);
      return NULL;
    }
  return session;
}

static int
create_category (LPDISPATCH categories, const char *category, int color)
{
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
      return 1;
    }

  dispid = lookup_oom_dispid (categories, "Add");
  if (dispid == DISPID_UNKNOWN)
  {
    log_error ("%s:%s: could not find Add DISPID",
               SRCNAME, __func__);
    return -1;
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

  hr = categories->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                           DISPATCH_METHOD, &dispparams,
                           &rVariant, &execpinfo, &argErr);
  SysFreeString (b_name);
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
      return -1;
    }
  VariantClear (&rVariant);
  log_debug ("%s:%s: Created category '%s'",
             SRCNAME, __func__, category);
  return 0;
}

void
ensure_category_exists (LPDISPATCH application, const char *category, int color)
{
  if (!application || !category)
    {
      TRACEPOINT;
      return;
    }

  log_debug ("Ensure category exists called for %s, %i", category, color);

  LPDISPATCH stores = get_oom_object (application, "Session.Stores");
  if (!stores)
    {
      log_error ("%s:%s: No stores found.",
                 SRCNAME, __func__);
      return;
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
              break;
            }
          char *name = get_oom_string (category_obj, "Name");
          if (name && !strcmp (category, name))
            {
              log_debug ("%s:%s: Found category '%s'",
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
              log_debug ("%s:%s: Found category '%s'",
                         SRCNAME, __func__, category);
            }
        }
      /* Otherwise we have to create the category */
      gpgol_release (categories);
    }
  gpgol_release (stores);
}

int
add_category (LPDISPATCH mail, const char *category)
{
  char *tmp = get_oom_string (mail, "Categories");
  if (!tmp)
    {
      TRACEPOINT;
      return 1;
    }

  if (strstr (tmp, category))
    {
      log_debug ("%s:%s: category '%s' already added.",
                 SRCNAME, __func__, category);
      return 0;
    }

  std::string newstr (tmp);
  xfree (tmp);
  if (!newstr.empty ())
    {
      newstr += ", ";
    }
  newstr += category;

  return put_oom_string (mail, "Categories", newstr.c_str ());
}

int
remove_category (LPDISPATCH mail, const char *category)
{
  char *tmp = get_oom_string (mail, "Categories");
  if (!tmp)
    {
      TRACEPOINT;
      return 1;
    }
  std::string newstr (tmp);
  xfree (tmp);
  std::string cat (category);

  size_t pos1 = newstr.find (cat);
  size_t pos2 = newstr.find (std::string(", ") + cat);
  if (pos1 == std::string::npos && pos2 == std::string::npos)
    {
      log_debug ("%s:%s: category '%s' not found.",
                 SRCNAME, __func__, category);
      return 0;
    }

  size_t len = cat.size();
  if (pos2)
    {
      len += 2;
    }
  newstr.erase (pos2 != std::string::npos ? pos2 : pos1, len);
  log_debug ("%s:%s: removing category '%s'",
             SRCNAME, __func__, category);

  return put_oom_string (mail, "Categories", newstr.c_str ());
}

static char *
generate_uid ()
{
  UUID uuid;
  UuidCreate (&uuid);

  unsigned char *str;
  UuidToStringA (&uuid, &str);

  char *ret = strdup ((char*)str);
  RpcStringFreeA (&str);

  return ret;
}

char *
get_unique_id (LPDISPATCH mail, int create, const char *uuid)
{
  if (!mail)
    {
      return NULL;
    }

  /* Get the User Properties. */
  if (!create)
    {
      char *uid = get_pa_string (mail, GPGOL_UID_DASL);
      if (!uid)
        {
          log_debug ("%s:%s: No uuid found in oom for '%p'",
                     SRCNAME, __func__, mail);
          return NULL;
        }
      else
        {
          log_debug ("%s:%s: Found uid '%s' for '%p'",
                     SRCNAME, __func__, uid, mail);
          return uid;
        }
    }
  char *newuid;
  if (!uuid)
    {
      newuid = generate_uid ();
    }
  else
    {
      newuid = strdup (uuid);
    }
  int ret = put_pa_string (mail, GPGOL_UID_DASL, newuid);

  if (ret)
    {
      log_debug ("%s:%s: failed to set uid '%s' for '%p'",
                 SRCNAME, __func__, newuid, mail);
      xfree (newuid);
      return NULL;
    }


  log_debug ("%s:%s: '%p' has now the uid: '%s' ",
             SRCNAME, __func__, mail, newuid);
  return newuid;
}

HWND
get_active_hwnd ()
{
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      return nullptr;
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
              return nullptr;
            }
        }
    }

  /* Both explorer and inspector have this. */
  char *caption = get_oom_string (activeWindow, "Caption");
  gpgol_release (activeWindow);
  if (!caption)
    {
      TRACEPOINT;
      return nullptr;
    }
  /* Might not be completly true for multiple explorers
     on the same folder but good enugh. */
  HWND hwnd = FindWindowExA(NULL, NULL, "rctrl_renwnd32",
                            caption);
  xfree (caption);

  return hwnd;
}

LPDISPATCH
create_mail ()
{
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      return nullptr;
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
      return ret;
    }

  ret = var.pdispVal;
  return ret;
}

LPDISPATCH
get_account_for_mail (const char *mbox)
{
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      return nullptr;
   }

  LPDISPATCH accounts = get_oom_object (app, "Session.Accounts");

  if (!accounts)
    {
      TRACEPOINT;
      return nullptr;
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
          TRACEPOINT;
          continue;
        }
      if (!stricmp (mbox, smtpAddr))
        {
          gpgol_release (accounts);
          xfree (smtpAddr);
          return account;
        }
      xfree (smtpAddr);
    }
  gpgol_release (accounts);

  log_error ("%s:%s: Failed to find account for '%s'.",
             SRCNAME, __func__, mbox);

  return nullptr;
}

char *
get_sender_SendUsingAccount (LPDISPATCH mailitem, bool *r_is_GSuite)
{
  LPDISPATCH sender = get_oom_object (mailitem, "SendUsingAccount");
  if (!sender)
    {
      return nullptr;
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
      return buf;
    }
  xfree (buf);
  return nullptr;
}

char *
get_sender_Sender (LPDISPATCH mailitem)
{
  LPDISPATCH sender = get_oom_object (mailitem, "Sender");
  if (!sender)
    {
      return nullptr;
    }
  char *buf = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
  gpgol_release (sender);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 2", SRCNAME, __func__);
      return buf;
    }
  xfree (buf);
  /* We have a sender object but not yet an smtp address likely
     exchange. Try some more propertys of the message. */
  buf = get_pa_string (mailitem, PR_TAG_SENDER_SMTP_ADDRESS);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 3", SRCNAME, __func__);
      return buf;
    }
  xfree (buf);
  buf = get_pa_string (mailitem, PR_TAG_RECEIVED_REPRESENTING_SMTP_ADDRESS);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 4", SRCNAME, __func__);
      return buf;
    }
  xfree (buf);
  return nullptr;
}

char *
get_sender_CurrentUser (LPDISPATCH mailitem)
{
  LPDISPATCH sender = get_oom_object (mailitem,
                                      "Session.CurrentUser");
  if (!sender)
    {
      return nullptr;
    }
  char *buf = get_pa_string (sender, PR_SMTP_ADDRESS_DASL);
  gpgol_release (sender);
  if (buf && strlen (buf))
    {
      log_debug ("%s:%s Sender fallback 5", SRCNAME, __func__);
      return buf;
    }
  xfree (buf);
  return nullptr;
}

char *
get_sender_SenderEMailAddress (LPDISPATCH mailitem)
{

  char *type = get_oom_string (mailitem, "SenderEmailType");
  if (type && !strcmp ("SMTP", type))
    {
      char *senderMail = get_oom_string (mailitem, "SenderEmailAddress");
      if (senderMail)
        {
          log_debug ("%s:%s: Sender found", SRCNAME, __func__);
          xfree (type);
          return senderMail;
        }
    }
  xfree (type);
  return nullptr;
}

char *
get_inline_body ()
{
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();
  if (!app)
    {
      TRACEPOINT;
      return nullptr;
    }

  LPDISPATCH explorer = get_oom_object (app, "ActiveExplorer");

  if (!explorer)
    {
      TRACEPOINT;
      return nullptr;
    }

  LPDISPATCH inlineResponse = get_oom_object (explorer, "ActiveInlineResponse");
  gpgol_release (explorer);

  if (!inlineResponse)
    {
      return nullptr;
    }

  char *body = get_oom_string (inlineResponse, "Body");
  gpgol_release (inlineResponse);

  return body;
}

int
get_ex_major_version_for_addr (const char *mbox)
{
  LPDISPATCH account = get_account_for_mail (mbox);
  if (!account)
    {
      TRACEPOINT;
      return -1;
    }

  char *version_str = get_oom_string (account, "ExchangeMailboxServerVersion");
  gpgol_release (account);

  if (!version_str)
    {
      return -1;
    }
  long int version = strtol (version_str, nullptr, 10);
  xfree (version_str);

  return (int) version;
}

int
get_ol_ui_language ()
{
  LPDISPATCH app = GpgolAddin::get_instance()->get_application();
  if (!app)
    {
      TRACEPOINT;
      return 0;
    }

  LPDISPATCH langSettings = get_oom_object (app, "LanguageSettings");
  if (!langSettings)
    {
      TRACEPOINT;
      return 0;
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

  if (invoke_oom_method_with_parms_type (langSettings, "LanguageID", &var,
                                         &dispparams,
                                         DISPATCH_PROPERTYGET))
    {
      TRACEPOINT;
      return 0;
    }
  if (var.vt != VT_INT && var.vt != VT_I4)
    {
      TRACEPOINT;
      return 0;
    }

  int result = var.intVal;

  log_debug ("XXXXX %i", result);
  VariantClear (&var);
  return result;
}

void
log_addins ()
{
  LPDISPATCH app = GpgolAddin::get_instance ()->get_application ();

  if (!app)
    {
      TRACEPOINT;
      return;
   }

  LPDISPATCH addins = get_oom_object (app, "COMAddins");

  if (!addins)
    {
      TRACEPOINT;
      return;
    }

  std::string activeAddins;
  int count = get_oom_int (addins, "Count");
  for (int i = 1; i <= count; i++)
    {
      std::string item = std::string ("Item(") + std::to_string (i) + ")";

      LPDISPATCH addin = get_oom_object (addins, item.c_str ());

      if (!addin)
        {
          TRACEPOINT;
          continue;
        }
      bool connected = get_oom_bool (addin, "Connect");
      if (!connected)
        {
          gpgol_release (addin);
          continue;
        }

      char *progId = get_oom_string (addin, "ProgId");
      gpgol_release (addin);

      if (!progId)
        {
          TRACEPOINT;
          continue;
        }
      activeAddins += std::string (progId) + "\n";
      xfree (progId);
    }
  gpgol_release (addins);

  log_debug ("%s:%s:Active Addins:\n%s", SRCNAME, __func__,
             activeAddins.c_str ());
  return;
}
