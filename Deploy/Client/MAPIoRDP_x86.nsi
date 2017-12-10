!include CommonPre.nsh

OutFile "..\MAPIoRDP-ClientPlugin-${PRODUCT_VERSION}_x86.exe"

; Select language
Function .onInit
  StrCpy $INSTDIR "$PROGRAMFILES\MAPIoRDP-ClientPlugin"
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""

ShowInstDetails show
ShowUnInstDetails show

Section "ProgramFiles" SEC01
  SetOverwrite try

  SetOutPath "$INSTDIR"
  File /oname=MAPIoRDPClientPlugin_x86.dll "..\..\Release\MAPIoRDPClientPlugin.dll"
  File "MAPIoRDPClientPlugin_x86.reg"
  File "..\vcredist_x86.exe"
  
  WriteRegStr HKLM "Software\Microsoft\Terminal Server Client\Default\AddIns\MAPIoRDP" "Name" "$INSTDIR\MAPIoRDPClientPlugin_x86.dll"
  Execwait '"$INSTDIR\vcredist_x86.exe" /passive'
  
SectionEnd

!include CommonPost.nsh
