; Deli Elements for PowerPoint: per-user installer (no admin rights needed)
Unicode true
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "Sections.nsh"

!define APPNAME "Deli Elements"
!define COMPANY "The Deli Group"
!define VERSION "2.1.0"
!define ADDIN_ID "6f1d2b7e-3c4a-4d8e-9a5b-de11e1e2027a"
!define UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\DeliElements"
!define WEF_DEV "Software\Microsoft\Office\16.0\WEF\Developer"
!define PPT_OPTS "Software\Microsoft\Office\16.0\PowerPoint\Options"

Name "${APPNAME}"
OutFile "Deli-Elements-Setup.exe"
InstallDir "$LOCALAPPDATA\DeliElements"
RequestExecutionLevel user
SetCompressor /SOLID lzma
BrandingText "${COMPANY}"

VIProductVersion "${VERSION}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "CompanyName" "${COMPANY}"
VIAddVersionKey "FileDescription" "${APPNAME} for PowerPoint installer"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "ProductVersion" "${VERSION}"
VIAddVersionKey "LegalCopyright" "${COMPANY}"

!define MUI_ICON "deli.ico"
!define MUI_UNICON "deli.ico"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_RIGHT
!define MUI_HEADERIMAGE_BITMAP "header.bmp"
!define MUI_HEADERIMAGE_UNBITMAP "header.bmp"
!define MUI_WELCOMEFINISHPAGE_BITMAP "welcome.bmp"
!define MUI_UNWELCOMEFINISHPAGE_BITMAP "welcome.bmp"
!define MUI_ABORTWARNING

!define MUI_WELCOMEPAGE_TITLE "Welcome to Deli Elements"
!define MUI_WELCOMEPAGE_TEXT "This sets up PowerPoint for The Deli Group:$\r$\n$\r$\n•  the Deli tab on the ribbon, with on-brand elements, 4,800 icons, layout tools, the agenda wizard and the brand check$\r$\n•  The Deli Group template as the starting point for every new presentation$\r$\n$\r$\nPlease save your work and close PowerPoint before you continue."
!insertmacro MUI_PAGE_WELCOME
!define MUI_COMPONENTSPAGE_SMALLDESC
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TITLE "Deli Elements is installed"
!define MUI_FINISHPAGE_TEXT "Every new presentation now starts on The Deli Group template with the Deli tab and panel loaded.$\r$\n$\r$\nTip: 'New Deli presentation' is also in your Start menu."
!define MUI_FINISHPAGE_RUN
!define MUI_FINISHPAGE_RUN_TEXT "Start a new Deli presentation now"
!define MUI_FINISHPAGE_RUN_FUNCTION OpenPowerPoint
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

; ---------- helpers ----------
Function WaitForPowerPoint
  check:
    nsExec::ExecToStack 'cmd /c tasklist /FI "IMAGENAME eq POWERPNT.EXE" /NH | find /I "POWERPNT.EXE"'
    Pop $0
    ${If} $0 == 0
      MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "PowerPoint is open.$\r$\n$\r$\nPlease save your work, close PowerPoint, then click Retry." IDRETRY check
      Abort
    ${EndIf}
FunctionEnd

Function OpenPowerPoint
  ; start a new Deli deck: the template carries the add-in, so the Deli tab loads straight away
  ExecShell "" "$INSTDIR\Deli presentation.potx"
FunctionEnd

Function .onInit
  Call WaitForPowerPoint
FunctionEnd

; ---------- sections ----------
Section "Deli tab in PowerPoint (add-in)" SecAddin
  SectionIn RO
  SetOutPath "$INSTDIR"
  File "manifest.xml"
  File "deli.ico"
  File "Deli presentation.potx"
  CreateShortCut "$SMPROGRAMS\New Deli presentation.lnk" "$INSTDIR\Deli presentation.potx" "" "$INSTDIR\deli.ico" 0
  WriteRegStr HKCU "${WEF_DEV}" "${ADDIN_ID}" "$INSTDIR\manifest.xml"
  ; clear PowerPoint's add-in cache so the newest version loads
  RMDir /r "$LOCALAPPDATA\Microsoft\Office\16.0\Wef"

  WriteUninstaller "$INSTDIR\Uninstall Deli Elements.exe"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayName" "${APPNAME} for PowerPoint"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "${UNINST_KEY}" "Publisher" "${COMPANY}"
  WriteRegStr HKCU "${UNINST_KEY}" "DisplayIcon" "$INSTDIR\deli.ico"
  WriteRegStr HKCU "${UNINST_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINST_KEY}" "UninstallString" '"$INSTDIR\Uninstall Deli Elements.exe"'
  WriteRegStr HKCU "${UNINST_KEY}" "QuietUninstallString" '"$INSTDIR\Uninstall Deli Elements.exe" /S'
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINST_KEY}" "NoRepair" 1
  WriteRegDWORD HKCU "${UNINST_KEY}" "EstimatedSize" 900
SectionEnd

Section "Make The Deli Group the default template" SecTemplate
  ; default for new presentations (Ctrl+N / Default Theme)
  CreateDirectory "$APPDATA\Microsoft\Templates"
  ${If} ${FileExists} "$APPDATA\Microsoft\Templates\Blank.potx"
  ${AndIfNot} ${FileExists} "$APPDATA\Microsoft\Templates\Blank.before-deli.potx"
    CopyFiles /SILENT "$APPDATA\Microsoft\Templates\Blank.potx" "$APPDATA\Microsoft\Templates\Blank.before-deli.potx"
  ${EndIf}
  SetOutPath "$APPDATA\Microsoft\Templates"
  File "Blank.potx"

  ; full template under New > Personal
  ReadRegStr $0 HKCU "${PPT_OPTS}" "PersonalTemplates"
  ${If} $0 == ""
    StrCpy $0 "$DOCUMENTS\Custom Office Templates"
    WriteRegExpandStr HKCU "${PPT_OPTS}" "PersonalTemplates" "$0"
  ${EndIf}
  ExpandEnvStrings $0 $0
  WriteRegStr HKCU "Software\DeliElements" "PersonalTemplatesUsed" "$0"
  CreateDirectory "$0"
  SetOutPath "$0"
  File "The Deli Group.potx"
SectionEnd

LangString DESC_ADDIN ${LANG_ENGLISH} "Adds the Deli tab: elements, icons, layout tools, agenda wizard and brand check. Updates itself automatically."
LangString DESC_TPL ${LANG_ENGLISH} "New presentations start on the Deli template; the full template appears under New > Personal. Your current default is backed up."
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecAddin} $(DESC_ADDIN)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecTemplate} $(DESC_TPL)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; ---------- uninstall ----------
Function un.onInit
  check:
    nsExec::ExecToStack 'cmd /c tasklist /FI "IMAGENAME eq POWERPNT.EXE" /NH | find /I "POWERPNT.EXE"'
    Pop $0
    ${If} $0 == 0
      MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Please close PowerPoint, then click Retry." IDRETRY check
      Abort
    ${EndIf}
FunctionEnd

Section "Uninstall"
  DeleteRegValue HKCU "${WEF_DEV}" "${ADDIN_ID}"
  RMDir /r "$LOCALAPPDATA\Microsoft\Office\16.0\Wef"

  ${If} ${FileExists} "$APPDATA\Microsoft\Templates\Blank.before-deli.potx"
    Delete "$APPDATA\Microsoft\Templates\Blank.potx"
    Rename "$APPDATA\Microsoft\Templates\Blank.before-deli.potx" "$APPDATA\Microsoft\Templates\Blank.potx"
  ${Else}
    Delete "$APPDATA\Microsoft\Templates\Blank.potx"
  ${EndIf}
  ReadRegStr $0 HKCU "Software\DeliElements" "PersonalTemplatesUsed"
  ${If} $0 != ""
    Delete "$0\The Deli Group.potx"
  ${EndIf}
  DeleteRegKey HKCU "Software\DeliElements"

  Delete "$INSTDIR\manifest.xml"
  Delete "$INSTDIR\deli.ico"
  Delete "$INSTDIR\Deli presentation.potx"
  Delete "$SMPROGRAMS\New Deli presentation.lnk"
  Delete "$INSTDIR\Uninstall Deli Elements.exe"
  RMDir "$INSTDIR"
  DeleteRegKey HKCU "${UNINST_KEY}"
SectionEnd
