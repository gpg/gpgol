/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Mon Sep 06 12:34:18 2004
 */
/* Compiler settings for C:\Source\GData\Gdgpg\GDGPG.idl:
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


const IID LIBID_GDGPGLib = {0x25EC75AD,0x5EB9,0x4E20,{0xA4,0xB4,0xE7,0x1C,0x36,0x37,0x0F,0x62}};


const CLSID CLSID_GDGPG = {0x321F09FC,0xE2FD,0x409B,{0xB8,0xD1,0x60,0xFA,0x7D,0xCD,0xA5,0x31}};


#ifdef __cplusplus
}
#endif

