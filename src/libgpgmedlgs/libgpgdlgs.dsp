# Microsoft Developer Studio Project File - Name="libgpgmedlgs" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** NICHT BEARBEITEN **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=libgpgmedlgs - Win32 Debug
!MESSAGE Dies ist kein gültiges Makefile. Zum Erstellen dieses Projekts mit NMAKE
!MESSAGE verwenden Sie den Befehl "Makefile exportieren" und führen Sie den Befehl
!MESSAGE 
!MESSAGE NMAKE /f "libgpgdlgs.mak".
!MESSAGE 
!MESSAGE Sie können beim Ausführen von NMAKE eine Konfiguration angeben
!MESSAGE durch Definieren des Makros CFG in der Befehlszeile. Zum Beispiel:
!MESSAGE 
!MESSAGE NMAKE /f "libgpgdlgs.mak" CFG="libgpgmedlgs - Win32 Debug"
!MESSAGE 
!MESSAGE Für die Konfiguration stehen zur Auswahl:
!MESSAGE 
!MESSAGE "libgpgmedlgs - Win32 Release" (basierend auf  "Win32 (x86) Dynamic-Link Library")
!MESSAGE "libgpgmedlgs - Win32 Debug" (basierend auf  "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "libgpgmedlgs - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "libgpgmedlgs_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /I "c:\oss\w32root\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "libgpgmedlgs_EXPORTS" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x407 /d "NDEBUG"
# ADD RSC /l 0x407 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 winspool.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib mapi32.lib gdi32.lib shell32.lib comctl32.lib comdlg32.lib advapi32.lib kernel32.lib user32.lib c:\oss\w32root\lib\libgpgme.a c:\oss\w32root\lib\libgpg-error.a /nologo /dll /machine:I386 /out:"Release/libgpgmedlgs.dll"

!ELSEIF  "$(CFG)" == "libgpgmedlgs - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "libgpgmedlgs_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /I "c:\oss\w32root\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "libgpgmedlgs_EXPORTS" /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x407 /d "_DEBUG"
# ADD RSC /l 0x407 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 mapi32.lib gdi32.lib shell32.lib comctl32.lib comdlg32.lib advapi32.lib kernel32.lib user32.lib c:\oss\w32root\lib\libgpgme.a c:\oss\w32root\lib\libgpg-error.a /nologo /dll /debug /machine:I386 /out:"Debug/libgpgmedlgs.dll" /pdbtype:sept

!ENDIF 

# Begin Target

# Name "libgpgmedlgs - Win32 Release"
# Name "libgpgmedlgs - Win32 Debug"
# Begin Group "mapi"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\MapiGPGME.cpp
# End Source File
# Begin Source File

SOURCE=..\MapiGPGME.h
# End Source File
# End Group
# Begin Group "misc"

# PROP Default_Filter ""
# Begin Source File

SOURCE=..\ChangeLog
# End Source File
# Begin Source File

SOURCE=..\gpgmedlgs.rc
# End Source File
# Begin Source File

SOURCE=..\libgpgmedlgs.def
# End Source File
# Begin Source File

SOURCE="..\..\gpgme-cvs\gpgme\stpcpy.c"
# End Source File
# Begin Source File

SOURCE="..\..\gpgme-cvs\gpgme\vasprintf.c"
# End Source File
# End Group
# Begin Source File

SOURCE=..\common.c
# End Source File
# Begin Source File

SOURCE="..\config-dialog.c"
# End Source File
# Begin Source File

SOURCE="..\engine-gpgme.c"
# End Source File
# Begin Source File

SOURCE=..\engine.h
# End Source File
# Begin Source File

SOURCE=..\intern.h
# End Source File
# Begin Source File

SOURCE=..\keycache.c
# End Source File
# Begin Source File

SOURCE=..\keycache.h
# End Source File
# Begin Source File

SOURCE=..\keylist.c
# End Source File
# Begin Source File

SOURCE=..\logging.c
# End Source File
# Begin Source File

SOURCE=..\main.c
# End Source File
# Begin Source File

SOURCE="..\passphrase-dialog.c"
# End Source File
# Begin Source File

SOURCE="..\recipient-dialog.c"
# End Source File
# Begin Source File

SOURCE=..\resource.h
# End Source File
# Begin Source File

SOURCE="..\verify-dialog.c"
# End Source File
# End Target
# End Project
