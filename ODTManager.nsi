; Configuration "ODTManager app"
Unicode True
;CPU target
;Target amd64-unicode
Target x86-unicode

!define PRODUCT_NAME "ODTManager"
!define VERSION "1.0"
!define BUILD "0.0"
!define PRODUCT_VERSION "${VERSION} (${__DATE__})"
!define PRODUCT_PUBLISHER "dialmak"

;UAC Manifest
;RequestExecutionLevel admin
RequestExecutionLevel User
ManifestSupportedOS all
ManifestLongPathAware true
ManifestDPIAware true
CRCCheck on
SetCompress auto
SetCompressor /solid lzma
Name "${PRODUCT_NAME}"
OutFile "${PRODUCT_NAME}.exe"
InstallDir "$TEMP\${PRODUCT_NAME}"
ShowInstDetails hide
XPStyle on
SilentInstall silent
BrandingText "${PRODUCT_NAME} v${VERSION}.${BUILD}"
Icon "${PRODUCT_NAME}.ico"
Caption "${PRODUCT_NAME}"
 
VIProductVersion "${VERSION}.${BUILD}"
VIAddVersionKey  "ProductName" "${PRODUCT_NAME}"
VIAddVersionKey  "Comments" "Indtsll Office 365/2016/2019/2021"
VIAddVersionKey  "CompanyName" "${PRODUCT_PUBLISHER}"
VIAddVersionKey  "LegalCopyright" "${PRODUCT_PUBLISHER} @ 2022"
;VIAddVersionKey  "LegalTrademarks" "${PRODUCT_PUBLISHER} @ 2022"
VIAddVersionKey  "FileDescription" "Instal Office 2016/2019/365"
VIAddVersionKey  "FileVersion" "${VERSION}"
VIAddVersionKey  "ProductVersion" "${PRODUCT_VERSION}"
VIAddVersionKey  "InternalName" "${PRODUCT_NAME}"
VIAddVersionKey  "OriginalFilename" "${PRODUCT_NAME}.exe"


Function .onInit
  HideWindow
  ;System::Call 'kernel32::CreateMutex(p 0, i 0, t "${PRODUCT_NAME}") p .r1 ?e'
  ;Pop $R0
  ;StrCmp $R0 0 +3
  ;MessageBox MB_OK|MB_ICONEXCLAMATION "${PRODUCT_NAME} is already running."
  ;Abort
  ClearErrors
  Delete "$INSTDIR\AutoHotkey.exe"
  IfErrors 0 +3
  MessageBox MB_OK|MB_ICONEXCLAMATION "${PRODUCT_NAME} is already running."
  Abort
  RMDir /r "$INSTDIR"
FunctionEnd 
 
Section -Main
  HideWindow
  CreateDirectory "$INSTDIR"
  SetOutPath "$INSTDIR"
  SetOverwrite try
  File "AutoHotkey.exe"
  File "${PRODUCT_NAME}.ahk"
  File "${PRODUCT_NAME}.ico"
  ;File "setup.exe" 
SectionEnd
 
;Section -Settings
;  HideWindow
;SectionEnd
 
Section -Run
  HideWindow 
  ExecShellWait "open" "$INSTDIR\AutoHotkey.exe" "/CP65001 $\"$INSTDIR\${PRODUCT_NAME}.ahk$\" $\"$EXEDIR$\" $\"$EXEFILE$\""
SectionEnd
 
;Section -Delete
;  HideWindow
;  SetOutPath $TEMP
;  RMDir /r "$INSTDIR"
;  SetAutoClose true
;SectionEnd



