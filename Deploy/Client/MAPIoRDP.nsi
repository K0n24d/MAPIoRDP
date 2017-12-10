!include CommonPre.nsh

OutFile "..\MAPIoRDP-ClientPlugin-${PRODUCT_VERSION}.exe"

; Select language
Function .onInit
  ${If} ${RunningX64}
    StrCpy $INSTDIR "$PROGRAMFILES64\MAPIoRDP-ClientPlugin"
  ${Else}
    StrCpy $INSTDIR "$PROGRAMFILES\MAPIoRDP-ClientPlugin"
  ${EndIf}
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd
 
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""

ShowInstDetails show
ShowUnInstDetails show

Section "ProgramFiles" SEC01
  SetOverwrite try

  SetOutPath "$INSTDIR"
  File /oname=MAPIoRDPClientPlugin_x64.dll "..\..\x64\Release\MAPIoRDPClientPlugin.dll"
  File "MAPIoRDPClientPlugin_x64.reg"
  File "..\vcredist_x64.exe"
  File /oname=MAPIoRDPClientPlugin_x86.dll "..\..\Release\MAPIoRDPClientPlugin.dll"
  File "MAPIoRDPClientPlugin_x86.reg"
  File "..\vcredist_x86.exe"
  
  ${If} ${RunningX64}
    WriteRegStr HKLM "Software\Microsoft\Terminal Server Client\Default\AddIns\MAPIoRDP" "Name" "$INSTDIR\MAPIoRDPClientPlugin_x64.dll"
    Execwait '"$INSTDIR\vcredist_x64.exe" /passive'
  ${Else}
    WriteRegStr HKLM "Software\Microsoft\Terminal Server Client\Default\AddIns\MAPIoRDP" "Name" "$INSTDIR\MAPIoRDPClientPlugin_x86.dll"
    Execwait '"$INSTDIR\vcredist_x86.exe" /passive'
  ${EndIf}
  
SectionEnd

!include CommonPost.nsh
