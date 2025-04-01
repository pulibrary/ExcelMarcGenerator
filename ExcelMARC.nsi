; NSIS Excel Add-In Installer Script
; Include
!include MUI.nsh
!include LogicLib.nsh

; General
!define filename "MARC.xlam"
!define displayname "Excel MARC Record Generator Add-In"

Name "${displayname}"
OutFile "ExcelMARCInstaller.exe"
InstallDir "$APPDATA\${displayname}"
InstallDirRegKey HKCU "Software\${displayname}" "InstallDir" ;

; Interface Settings
!define MUI_ABORTWARNING

; Installer Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

; Uninstaller Pages
!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Languages
!insertmacro MUI_LANGUAGE "English"

; Prerequisites Section

SetOverwrite On

; Installer Section
Section "-Install"
SetOutPath "$INSTDIR"

; ADD FILES HERE
File ${filename}

; Check Installed Excel Version
ReadRegStr $1 HKCR "Excel.Application\CurVer" ""

${If} $1 == 'Excel.Application.12' ; Excel 2007
StrCpy $2 "12.0"
${ElseIf} $1 == 'Excel.Application.14' ; Excel 2010
StrCpy $2 "14.0"
${ElseIf} $1 == 'Excel.Application.15' ; Excel 2013
StrCpy $2 "15.0"
${ElseIf} $1 == 'Excel.Application.16' ; Excel 2019
StrCpy $2 "16.0"
${Else}
Abort "An appropriate version of Excel is not installed. $\n${displayname} setup will be canceled."
${EndIf}

; Find available "OPEN" key
StrCpy $3 ""
loop:
ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == ""
; Available OPEN key found
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

; Write install data to registry
WriteRegStr HKCU "Software\${displayname}" "InstallDir" $INSTDIR
; Install Directory
WriteRegStr HKCU "Software\${displayname}" "ExcelCurVer" $2
; Current Excel Version

; Write key to install AddIn in Excel Addin Manager
WriteRegStr HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3" '"$INSTDIR\${filename}"'

; Write keys to uninstall
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "DisplayName" "${displayname}"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "UninstallString" '"$INSTDIR\uninstall.exe"'
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "NoModify" 1
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}" "NoRepair" 1

; Create uninstaller
WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

; Uninstaller Section
Section "Uninstall"
; ADD FILES HERE...
Delete "$INSTDIR\${filename}"
Delete "$INSTDIR\uninstall.exe"

RMDir "$INSTDIR"

; Find AddIn Manager Key and Delete
; AddIn Manager key name and location may have changed since installation depending on actions taken by user in AddIn Manager.
; Need to search for the target AddIn key and delete if found.
ReadRegStr $2 HKCU "Software\${displayname}" "ExcelCurVer"
StrCpy $3 ""

loop:
ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == '"$INSTDIR\${filename}"'
; Found Key
DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${ElseIf} $4 == ""
; Blank Key Found. Addin is no longer installed in AddIn Manager.
; Need to delete Addin Manager Reference.
DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Add-in Manager" "$INSTDIR\${filename}"
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

DeleteRegKey HKCU "Software\${displayname}"
DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${displayname}"
SectionEnd