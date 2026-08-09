; ══════════════════════════════════════════════════════════════
;  REGISTRODOC PRO — INSTALADOR DE PRUEBA / BETA (INNO SETUP)
; ══════════════════════════════════════════════════════════════
; Para entregar a una persona que probará todo el programa durante el período de
; prueba de 30 días y llenará el documento de experiencia y mejoras.
; La app funciona 30 días sin código; luego pide activación (igual que la versión final).

[Setup]
AppName=RegistroDoc Pro (Versión de Prueba)
AppVersion=5.0.3
AppPublisher=RegistroDoc Pro - MEDUCA
DefaultDirName={localappdata}\RegistroDocPro_Prueba
DefaultGroupName=RegistroDoc Pro (Prueba)
DisableProgramGroupPage=yes
OutputBaseFilename=RegistroDoc_Instalador_Prueba
SetupIconFile=assets\icon_fixed.ico
UninstallDisplayIcon={app}\RegistroDoc.exe
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\RegistroDoc.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
; Documento de experiencia para que el tester lo llene
Source: "docs\PLANTILLA_EXPERIENCIA_PRUEBA.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\RegistroDoc Pro (Prueba)"; Filename: "{app}\RegistroDoc.exe"
Name: "{userdesktop}\RegistroDoc Pro (Prueba)"; Filename: "{app}\RegistroDoc.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\RegistroDoc.exe"; Description: "{cm:LaunchProgram,RegistroDoc Pro}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
function InitializeSetup(): Boolean;
begin
  MsgBox(
    'RegistroDoc Pro — VERSIÓN DE PRUEBA (30 días)' + #13#10 + #13#10 +
    'Gracias por probar el programa. Podrás usar TODAS las funciones durante 30 días.' + #13#10 + #13#10 +
    'Al terminar, encontrarás en la carpeta del programa un documento llamado' + #13#10 +
    '"PLANTILLA_EXPERIENCIA_PRUEBA.md" para anotar tu experiencia, mejoras y errores.' + #13#10 + #13#10 +
    'Tu retroalimentación es muy valiosa. ¡Gracias!',
    mbInformation, MB_OK);
  Result := True;
end;
