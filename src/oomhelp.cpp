/* oomhelp.cpp - Helper functions for the Outlook Object Model
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

#include "myexchext.h"
#include "common.h"

#include "oomhelp.h"


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
    tinfo->Release ();
  if (disp)
    disp->Release ();

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

  log_debug ("%s:%s: looking for %p->`%s'",
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

      if (pDisp)
        {
          pDisp->Release ();
          pDisp = NULL;
        }
      pObj->QueryInterface (IID_IDispatch, (LPVOID*)&pDisp);
      if (pObj != pStart)
        pObj->Release ();
      pObj = NULL;
      if (!pDisp)
        return NULL;  /* The object has no IDispatch interface.  */
      if (!*fullname)
        {
          log_debug ("%s:%s:         got %p",SRCNAME, __func__, pDisp);
          return pDisp; /* Ready.  */
        }
      
      /* Break out the next name part.  */
      {
        const char *dot;
        size_t n;
        
        dot = strchr (fullname, '.');
        if (dot == fullname)
          {
            pDisp->Release ();
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
                  pDisp->Release ();
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
          pDisp->Release ();
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
                          &vtResult, NULL, NULL);
      if (parmstr)
        SysFreeString (parmstr);
      if (hr != S_OK || vtResult.vt != VT_DISPATCH)
        {
          log_debug ("%s:%s:       error: '%s' p=%p vt=%d hr=0x%x",
                     SRCNAME, __func__,
                     name, vtResult.pdispVal, vtResult.vt, (unsigned int)hr);
          VariantClear (&vtResult);
          if (parmstr)
            SysFreeString (parmstr);
          pDisp->Release ();
          return NULL;  /* Invoke failed.  */
        }

      pObj = vtResult.pdispVal;
    }
  log_debug ("%s:%s:       error: no object", SRCNAME, __func__);
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

  /* Fixme: For now we assume a system pixel size of 16.  What we
     actually should do is to get the current value and then find the
     best matching icon.  */
  resource_id += 16;

  rc = put_picture_or_mask (pDisp, resource_id, size, 0);
  if (!rc)
    rc = put_picture_or_mask (pDisp, resource_id+1, size, 1);
  
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
                      NULL, NULL, NULL);
  SysFreeString (bstring);
  if (hr != S_OK)
    {
      log_debug ("%s:%s: Putting '%s' failed: %#lx", 
                 SRCNAME, __func__, name, hr);
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
      rVariant.pdispVal->Release ();
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
      log_debug ("%s:%s: Adding Control failed: %#lx - vt=%d",
                 SRCNAME, __func__, hr, rVariant.vt);
      VariantClear (&rVariant);
      return NULL;
    }
  return rVariant.pdispVal;
}

