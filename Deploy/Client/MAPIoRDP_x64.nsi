!include CommonPre.nsh

OutFile "..\MAPIoRDP-ClientPlugin-${PRODUCT_VERSION}_x64.exe"

; Select language
Function .onInit
  StrCpy $INSTDIR "$PROGRAMFILES64\MAPIoRDP-ClientPlugin"
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
  
  WriteRegStr HKLM "Software\Microsoft\Terminal Server Client\Default\AddIns\MAPIoRDP" "Name" "$INSTDIR\MAPIoRDPClientPlugin_x64.dll"
  Execwait '"$INSTDIR\vcredist_x64.exe" /passive'
  
SectionEnd

!include CommonPost.nsh
