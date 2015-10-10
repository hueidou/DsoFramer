# Microsoft Developer Studio Project File - Name="dsoframer" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=dsoframer - Win32 Release
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "dsoframer.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "dsoframer.mak" CFG="dsoframer - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "dsoframer - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "dsoframer - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "dsoframer - Win32 Release"

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
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /G5 /Gz /Zp4 /W3 /Zi /O1 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib advapi32.lib ole32.lib oleaut32.lib uuid.lib comdlg32.lib oledlg.lib winspool.lib shell32.lib urlmon.lib /nologo /base:"0x22000000" /version:1.2 /subsystem:windows /dll /map /debug /machine:I386 /out:"Release/dsoframer.ocx" /pdbtype:con /opt:nowin98 /merge:.rdata=.text
# SUBTRACT LINK32 /pdb:none
# Begin Custom Build - Performing registration
OutDir=.\Release
TargetPath=.\Release\dsoframer.ocx
InputPath=.\Release\dsoframer.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "dsoframer - Win32 Debug"

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
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /G6 /Zp4 /W3 /Gm /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /YX /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib advapi32.lib ole32.lib oleaut32.lib uuid.lib comdlg32.lib oledlg.lib winspool.lib shell32.lib urlmon.lib /nologo /base:"0x22000000" /version:1.2 /subsystem:windows /dll /map /debug /machine:I386 /out:"Debug/dsoframer.ocx" /pdbtype:con
# SUBTRACT LINK32 /pdb:none
# Begin Custom Build - Performing registration
OutDir=.\Debug
TargetPath=.\Debug\dsoframer.ocx
InputPath=.\Debug\dsoframer.ocx
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "dsoframer - Win32 Release"
# Name "dsoframer - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "c;cpp;rc;idl"
# Begin Source File

SOURCE=.\classfactory.cpp
# End Source File
# Begin Source File

SOURCE=.\dsofauto.cpp
# End Source File
# Begin Source File

SOURCE=.\dsofcontrol.cpp
# End Source File
# Begin Source File

SOURCE=.\dsofdocobj.cpp
# End Source File
# Begin Source File

SOURCE=.\dsofprint.cpp
# End Source File
# Begin Source File

SOURCE=.\lib\dsoframer.idl
# ADD MTL /h "lib/dsoframerlib.h" /iid "lib/dsoframerlib.c"
# End Source File
# Begin Source File

SOURCE=.\res\dsoframer.rc
# End Source File
# Begin Source File

SOURCE=.\mainentry.cpp

!IF  "$(CFG)" == "dsoframer - Win32 Release"

# ADD CPP /D "DSO_MIN_CRT_STARTUP"

!ELSEIF  "$(CFG)" == "dsoframer - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\utilities.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;def"
# Begin Source File

SOURCE=.\dsofdocobj.h
# End Source File
# Begin Source File

SOURCE=.\dsoframer.def
# End Source File
# Begin Source File

SOURCE=.\dsoframer.h
# End Source File
# Begin Source File

SOURCE=.\utilities.h
# End Source File
# Begin Source File

SOURCE=.\version.h
# End Source File
# End Group
# Begin Group "Resources"

# PROP Default_Filter "ico;cur;tlb;bmp"
# Begin Source File

SOURCE=.\res\dso.ico
# End Source File
# Begin Source File

SOURCE=.\lib\dsoframer.olb
# End Source File
# Begin Source File

SOURCE=.\res\toolbox.bmp
# End Source File
# End Group
# End Target
# End Project
