!include x64.nsh

; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "MAPIoRDP-ClientPlugin"
!define PRODUCT_VERSION "0.1.2"
!define PRODUCT_PUBLISHER "Christian KNAUBER"
!define PRODUCT_WEB_SITE "https://github.com/k0n24d"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\MAPIoRDP-ClientPlugin"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"
!define PRODUCT_STARTMENU_REGVAL "NSIS:StartMenuDir"

SetCompressor /SOLID lzma

!define MUI_LANGDLL_REGISTRY_ROOT "HKLM"
!define MUI_LANGDLL_REGISTRY_KEY "Software\Microsoft\Windows\CurrentVersion\App Paths\MAPIoRDP-ClientPlugin"
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"


; MUI 1.67 compatible ------
!include "MUI.nsh"

; MUI Settings
!define MUI_ABORTWARNING
;!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
;!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"
!define MUI_ICON "MAPIoRDPMailto.ico"
!define MUI_UNICON "MAPIoRDPMailto.ico"

; Welcome page
!insertmacro MUI_PAGE_WELCOME
;; License page
!insertmacro MUI_PAGE_LICENSE "..\..\LICENSE"
; Directory page
!insertmacro MUI_PAGE_DIRECTORY
; Instfiles page
!insertmacro MUI_PAGE_INSTFILES
; Finish page
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Reserve files
!insertmacro MUI_RESERVEFILE_INSTALLOPTIONS

; MUI end ------

; Language files and strings
!insertmacro MUI_LANGUAGE "French"
!insertmacro MUI_LANGUAGE "English"
LangString uninstalled       ${LANG_ENGLISH} "$(^Name) was successfully uninstalled from your computer."
LangString uninstalled       ${LANG_FRENCH}  "$(^Name) a été désinstallé avec succès de votre ordinateur."
LangString uninstalling      ${LANG_ENGLISH} "Removing previous installation."
LangString uninstalling      ${LANG_FRENCH}  "Désinstallation de la version précédente."

;Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
Name "${PRODUCT_NAME}"

; GetParent
 ; input, top of stack  (e.g. C:\Program Files\Foo)
 ; output, top of stack (replaces, with e.g. C:\Program Files)
 ; modifies no other variables.
 ;
 ; Usage:
 ;   Push "C:\Program Files\Directory\Whatever"
 ;   Call GetParent
 ;   Pop $R0
 ;   ; at this point $R0 will equal "C:\Program Files\Directory"
 
Function GetParent
 
  Exch $R0
  Push $R1
  Push $R2
  Push $R3
 
  StrCpy $R1 0
  StrLen $R2 $R0
 
  loop:
    IntOp $R1 $R1 + 1
    IntCmp $R1 $R2 get 0 get
    StrCpy $R3 $R0 1 -$R1
    StrCmp $R3 "\" get
  Goto loop
 
  get:
    StrCpy $R0 $R0 -$R1
 
    Pop $R3
    Pop $R2
    Pop $R1
    Exch $R0
 
FunctionEnd

Function GetAfterChar
  Exch $0 ; chop char
  Exch
  Exch $1 ; input string
  Push $2
  Push $3
  StrCpy $2 0
  loop:
    IntOp $2 $2 - 1
    StrCpy $3 $1 1 $2
    StrCmp $3 "" 0 +3
      StrCpy $0 ""
      Goto exit2
    StrCmp $3 $0 exit1
    Goto loop
  exit1:
    IntOp $2 $2 + 1
    StrCpy $0 $1 "" $2
  exit2:
    Pop $3
    Pop $2
    Pop $1
    Exch $0 ; output
FunctionEnd

; The "" makes the section hidden.
Section "" SecUninstallPrevious
  Call UninstallPrevious
SectionEnd

Function UninstallPrevious
  Var /GLOBAL uninstaller
  ReadRegStr $uninstaller HKLM "${PRODUCT_UNINST_KEY}" "UninstallString"
  StrCmp $uninstaller "" done

  DetailPrint $(uninstalling)

  Push $uninstaller
  Call GetParent
  Pop $R0
  Var /GLOBAL oldinstall
  StrCpy $oldinstall $R0
  
  Push $uninstaller
  Push "\"
  Call GetAfterChar
  Pop $R0
  Var /GLOBAL uninst
  StrCpy $uninst $R0

  InitPluginsDir
  CreateDirectory "$pluginsdir\unold" ; Make sure plugins do not conflict with a old uninstaller 
  CopyFiles /SILENT /FILESONLY "$uninstaller" "$pluginsdir\unold"
  ExecWait '"$pluginsdir\unold\$uninst" _?=$oldinstall /S'

  Done:
FunctionEnd
