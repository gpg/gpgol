
Overview of the GDGPG COM object
================================


To install the object run:
"regsvr32 gdgpg.dll"

To uninstall the object:
"regsvr32 /u gdgpg.dll"


Interfaces
==========

The old GDGPG interface is mainly for the GPG Exchange plug-in.
g10Code, the new interface tries to be a real GPG wrapper which
allows to do the most common GPG operations. For example to encrypt
or decrypt data or files.


How to use the object
=====================

With C++, you can use the #import directive to use the object. With
smart pointers it is pretty easy to access all provided functions of
the object.

For Visual Basic and related languages, you can use the CreateObject
syntax. Here is an example:

Dim gpg as g10Code
Dim rc as Integer
set gpg = CreateObject ("GDGPG.g10Code")

gpg.Binary = "c:\gnupg\gpg.exe"
gpg.Armor = True
gpg.AddRecipient ("twoaday@g10code.com")
gpg.Plaintext = "this is a test message"
rc = gpg.Encrypt
if rc <> 0 then
	MsgBox "encryption error"
else
	MsgBox gpg.CipherText
end if
