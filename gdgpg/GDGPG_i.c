/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Sep 30 12:18:45 2004
 */
/* Compiler settings for C:\Source_OSS\GData\Gdgpg\GDGPG.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IGDGPG = {0x83C42CA3,0x5D35,0x4120,{0x8D,0xF6,0x09,0xD8,0x35,0x16,0x65,0x94}};


const IID IID_Ig10Code = {0x4514D7BD,0x4BBE,0x47C1,{0x95,0x52,0x78,0xE1,0x89,0x62,0x06,0x36}};


const IID LIBID_GDGPGLib = {0x25EC75AD,0x5EB9,0x4E20,{0xA4,0xB4,0xE7,0x1C,0x36,0x37,0x0F,0x62}};


const CLSID CLSID_GDGPG = {0x321F09FC,0xE2FD,0x409B,{0xB8,0xD1,0x60,0xFA,0x7D,0xCD,0xA5,0x31}};


const IID DIID__Ig10CodeEvents = {0x69D0503E,0x250B,0x4994,{0xB3,0x88,0x97,0x5A,0x60,0x4E,0xAD,0xDD}};


const CLSID CLSID_g10Code = {0x658BFF7C,0x83DD,0x41F8,{0xA7,0x12,0x93,0x72,0x8A,0x34,0x9A,0x7F}};


#ifdef __cplusplus
}
#endif

