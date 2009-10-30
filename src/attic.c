/* Check whether the preview pane is visisble.  Returns:
   -1 := Don't know.
    0 := No
    1 := Yes.
 */
int
is_preview_pane_visible (LPEXCHEXTCALLBACK eecb)
{
  HRESULT hr;      
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant, rVariant;
      
  pDisp = find_outlook_property (eecb,
                                 "Application.ActiveExplorer.IsPaneVisible",
                                 &dispid);
  if (!pDisp)
    {
      log_debug ("%s:%s: ActiveExplorer.IsPaneVisible NOT found\n",
                 SRCNAME, __func__);
      return -1;
    }

  dispparams.rgvarg = &aVariant;
  dispparams.rgvarg[0].vt = VT_INT;
  dispparams.rgvarg[0].intVal = 3; /* olPreview */
  dispparams.cArgs = 1;
  dispparams.cNamedArgs = 0;
  rVariant.bstrVal = NULL;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      &rVariant, NULL, NULL);
  pDisp->Release();
  pDisp = NULL;
  if (hr == S_OK && rVariant.vt != VT_BOOL)
    {
      log_debug ("%s:%s: invoking IsPaneVisible succeeded but vt is %d",
                 SRCNAME, __func__, rVariant.vt);
      if (rVariant.vt == VT_BSTR && rVariant.bstrVal)
        SysFreeString (rVariant.bstrVal);
      return -1;
    }
  if (hr != S_OK)
    {
      log_debug ("%s:%s: invoking IsPaneVisible failed: %#lx",
                 SRCNAME, __func__, hr);
      return -1;
    }
  
  return !!rVariant.boolVal;
  
}


/* Set the preview pane to visible if visble is true or to invisible
   if visible is false.  */
void
show_preview_pane (LPEXCHEXTCALLBACK eecb, int visible)
{
  HRESULT hr;      
  LPDISPATCH pDisp;
  DISPID dispid;
  DISPPARAMS dispparams;
  VARIANT aVariant[2];
      
  pDisp = find_outlook_property (eecb,
                                 "Application.ActiveExplorer.ShowPane",
                                 &dispid);
  if (!pDisp)
    {
      log_debug ("%s:%s: ActiveExplorer.ShowPane NOT found\n",
                 SRCNAME, __func__);
      return;
    }

  dispparams.rgvarg = aVariant;
  dispparams.rgvarg[0].vt = VT_BOOL;
  dispparams.rgvarg[0].boolVal = !!visible;
  dispparams.rgvarg[1].vt = VT_INT;
  dispparams.rgvarg[1].intVal = 3; /* olPreview */
  dispparams.cArgs = 2;
  dispparams.cNamedArgs = 0;
  hr = pDisp->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT,
                      DISPATCH_METHOD, &dispparams,
                      NULL, NULL, NULL);
  pDisp->Release();
  pDisp = NULL;
  if (hr != S_OK)
    log_debug ("%s:%s: invoking ShowPane(%d) failed: %#lx",
               SRCNAME, __func__, visible, hr);
}


/***
Outlook interop – stopping user properties appearing on Outlook message print

Here’s a weird one… we were adding user properties to a message using the IUserProperties interface, but whenever we did this the property would render when the message was printed. This included whenever the message was sent to a recipient within the same Exchange organisation.

Last time I had to ask Microsoft PSS to help was 1992, and in that case they were nice enough to send me a handcrafted sample application on a 3.5″ floppy. This time however they got back to me in a day with this little code snippet:

void MarkPropNoPrint(Outlook.MailItem message, string propertyName)
{
// Late Binding in .NET
// http://support.microsoft.com/default.aspx?scid=kb;EN-US;302902
Type userPropertyType;
long dispidMember = 107;
long ulPropPrintable = 0×4;
string dispMemberName = String.Format(”[DispID={0}]“, dispidMember);
object[] dispParams;
Microsoft.Office.Interop.Outlook.UserProperty userProperty = message.UserProperties[propertyName];

if (null == userProperty) return;
userPropertyType = userProperty.GetType();

// Call IDispatch::Invoke to get the current flags
object flags = userPropertyType.InvokeMember(dispMemberName, System.Reflection.BindingFlags.GetProperty, null, userProperty, null);
long lFlags = long.Parse(flags.ToString());

// Remove the hidden property Printable flag
lFlags &= ~ulPropPrintable;

// Place the new flags property into an argument array
dispParams = new object[] {lFlags};

// Call IDispatch::Invoke to set the current flags
userPropertyType.InvokeMember(dispMemberName,
System.Reflection.BindingFlags.SetProperty, null, userProperty, dispParams);
}

Srsly… there is no way I would have worked that out for myself…

*/



/* HOWTO get the Window handle */

/* Use FindWindow. Use the Inspector.Caption property (be aware that on long */
/* captions the result will be truncated at about 255 characters). The class */
/* name to use for the Inspector window is "rctrl_renwnd32". */

/* I usually set any dialogs as modal to the Inspector window using */
/* SetWindowLong, then the dialog would end up in front of the mail window, */
/* which would be in front of the Outlook Explorer window. */


