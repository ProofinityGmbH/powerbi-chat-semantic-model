; Custom NSIS installer script for Power BI Chat
; This script registers the application as a Power BI External Tool

!include LogicLib.nsh

!macro customInstall
  ; Install Visual C++ Redistributable 2015-2022 (x64) if not already installed
  DetailPrint "==================================="
  DetailPrint "Checking Visual C++ Redistributable"
  DetailPrint "==================================="

  ; Check if VC++ 2015-2022 x64 is installed by checking registry
  ReadRegStr $2 HKLM "SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" "Version"
  ${If} $2 == ""
    ; Not installed, install it
    DetailPrint "Visual C++ Redistributable not found. Installing..."
    ExecWait '"$INSTDIR\resources\vc_redist.x64.exe" /install /quiet /norestart' $3
    ${If} $3 == 0
      DetailPrint "✓ Visual C++ Redistributable installed successfully"
    ${Else}
      DetailPrint "Warning: Visual C++ Redistributable installation returned code: $3"
      DetailPrint "The application may require manual installation of VC++ Redistributable"
    ${EndIf}
  ${Else}
    DetailPrint "✓ Visual C++ Redistributable already installed (Version: $2)"
  ${EndIf}
  DetailPrint "==================================="

  ; Get the installation directory
  Push $INSTDIR

  ; Create .pbitool.json with correct paths
  ; Since perMachine=true, app will always install to: C:\Program Files\Power BI Chat
  ; We write the path with properly escaped backslashes for JSON

  FileOpen $0 "$INSTDIR\PowerBIChat.pbitool.json" w
  FileWrite $0 '{$\r$\n'
  FileWrite $0 '  "name": "Power BI Chat",$\r$\n'
  FileWrite $0 '  "description": "AI-powered semantic model assistant that can analyze your entire data model",$\r$\n'
  FileWrite $0 '  "path": "C:\\Program Files\\Power BI Chat\\Power BI Chat.exe",$\r$\n'
  FileWrite $0 '  "arguments": "\$\"Server=%server%;Database=%database%;ApplicationName=PowerBI\$\"",$\r$\n'
  FileWrite $0 '  "version": "1.0.0",$\r$\n'
  FileWrite $0 '  "iconData": "image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAA/0lEQVR4nGNkoAB8OiH4H8bms3jPSI4ZTNSwHBuf5g6gFhi6DkCPc3LTAMWgoqKJrLiHARY8chokmEOs2hvoAriCTePZi9fXSXAAUUBKQlQT3RHY0gBNLGdgYGCAmosSWugOoJnluBwx4NmQ8eixUyipWFFJEUOR1PXdcPYzTVeChhJSf//efTibYAggG4aNT6n6AY+Cwe8A9DgklAZIVY+vJCTaEErUD/4oGPYOQK+MsBbFYSGBZFuwas16DDHkSolgCFBiOTH6CToAmw9IAYT0E5UNKXUEPjDgiXDUAegOuAHNIjQD6O1CbCFAM0dga5Ti60yQ0iwnFmA0y0cBAD/MVNZebnrjAAAAAElFTkSuQmCC"$\r$\n'
  FileWrite $0 '}$\r$\n'
  FileClose $0

  ; Register as Power BI External Tool in system-wide location
  ; Standard location: C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools

  ; Use Program Files (x86) for the Common Files location
  StrCpy $1 "$PROGRAMFILES32\Common Files\Microsoft Shared\Power BI Desktop\External Tools"

  DetailPrint "==================================="
  DetailPrint "Registering Power BI External Tool"
  DetailPrint "Target directory: $1"
  DetailPrint "==================================="

  ; Create the full directory structure
  CreateDirectory "$PROGRAMFILES32\Common Files"
  CreateDirectory "$PROGRAMFILES32\Common Files\Microsoft Shared"
  CreateDirectory "$PROGRAMFILES32\Common Files\Microsoft Shared\Power BI Desktop"
  CreateDirectory "$1"

  ; Verify directory was created
  IfFileExists "$1\*.*" DirExists 0
    DetailPrint "ERROR: Failed to create directory: $1"
    MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to create External Tools directory. Please ensure you have administrator privileges."
    Goto EndRegister
  DirExists:
    DetailPrint "✓ Directory created successfully"

  ; Copy .pbitool.json to External Tools directory
  DetailPrint "Copying PowerBIChat.pbitool.json to: $1"
  CopyFiles "$INSTDIR\PowerBIChat.pbitool.json" "$1\PowerBIChat.pbitool.json"

  ; Verify file was copied
  IfFileExists "$1\PowerBIChat.pbitool.json" FileExists 0
    DetailPrint "ERROR: Failed to copy pbitool.json file"
    MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to copy External Tool registration file. Please ensure you have administrator privileges."
    Goto EndRegister
  FileExists:
    DetailPrint "✓ File copied successfully"
    DetailPrint "==================================="
    DetailPrint "✓ Power BI Chat registered as External Tool"
    DetailPrint "Location: $1\PowerBIChat.pbitool.json"
    DetailPrint "==================================="

  EndRegister:
!macroend

!macro customUnInstall
  ; Remove from system-wide External Tools directory
  StrCpy $1 "$PROGRAMFILES32\Common Files\Microsoft Shared\Power BI Desktop\External Tools"

  DetailPrint "==================================="
  DetailPrint "Unregistering Power BI External Tool"
  DetailPrint "Removing: $1\PowerBIChat.pbitool.json"
  DetailPrint "==================================="

  Delete "$1\PowerBIChat.pbitool.json"

  IfFileExists "$1\PowerBIChat.pbitool.json" 0 RemoveSuccess
    DetailPrint "ERROR: Failed to remove pbitool.json"
    Goto EndUninstall
  RemoveSuccess:
    DetailPrint "✓ Power BI Chat unregistered from External Tools"

  EndUninstall:
!macroend
