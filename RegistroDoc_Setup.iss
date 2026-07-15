; ══════════════════════════════════════════════════════════════
;  REGISTRODOC PRO — CONFIGURACIÓN DE INSTALADOR WINDOWS (INNO SETUP)
; ══════════════════════════════════════════════════════════════
; Compatible con Windows 10 y 11. Cumple con normas ISO 25010 y ISO 27001.

[Setup]
AppName=RegistroDoc Pro
AppVersion=5.0.0
AppPublisher=RegistroDoc Pro - MEDUCA
DefaultDirName={localappdata}\RegistroDocPro
DefaultGroupName=RegistroDoc Pro
DisableProgramGroupPage=yes
OutputBaseFilename=RegistroDoc_Instalador
SetupIconFile=assets\icon_fixed.ico
UninstallDisplayIcon={app}\RegistroDoc.exe
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; Check: IsNotUpdate

[Files]
Source: "dist\RegistroDoc.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\RegistroDoc Pro"; Filename: "{app}\RegistroDoc.exe"
Name: "{userdesktop}\RegistroDoc Pro"; Filename: "{app}\RegistroDoc.exe"; Tasks: desktopicon; Check: IsNotUpdate

[Run]
Filename: "{app}\RegistroDoc.exe"; Description: "{cm:LaunchProgram,RegistroDoc Pro}"; Flags: nowait postinstall skipifsilent

; Registro de la aplicación para desinstalación segura
[Registry]
Root: HKCU; Subkey: "Software\RegistroDocPro"; Flags: uninsdeletekey

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
var
  PrevVersionInstalled: Boolean;

// ══════════════════════════════════════════════════════════════
//  1. VERIFICACIÓN DE SUITES OFIMÁTICAS (PRERREQUISITOS)
// ══════════════════════════════════════════════════════════════
function InitializeSetup(): Boolean;
var
  ExcelInstalado: Boolean;
  LibreOfficeInstalado: Boolean;
  ExecPath1: String;
  ExecPath2: String;
begin
  // 1. Detectar si hay una versión previa instalada (Chequeo robusto por ejecutables antiguos/nuevos y registro)
  ExecPath1 := ExpandConstant('{localappdata}\RegistroDocPro\RegistroDoc.exe');
  ExecPath2 := ExpandConstant('{localappdata}\RegistroDocPro\RegistroDocPro.exe');
  
  PrevVersionInstalled := FileExists(ExecPath1) or 
                          FileExists(ExecPath2) or 
                          RegValueExists(HKEY_CURRENT_USER, 'Software\RegistroDocPro', 'Cedula') or
                          RegKeyExists(HKEY_CURRENT_USER, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\RegistroDoc Pro_is1') or
                          RegKeyExists(HKEY_CURRENT_USER, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\RegistroDoc Pro (Con Datos)_is1');
  
  if PrevVersionInstalled then
  begin
    MsgBox(
      'Se ha detectado una versión previa de RegistroDoc Pro instalada en su equipo.' + #13#10 + #13#10 +
      'El instalador actualizará la aplicación a la nueva versión.' + #13#10 +
      'Nota: Su base de datos local y sus planillas Excel de notas existentes NO serán modificadas.', 
      mbInformation, 
      MB_OK
    );
  end;

  // 2. Verificar suites ofimáticas
  ExcelInstalado := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE') or
                    RegKeyExists(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE');
                    
  LibreOfficeInstalado := FileExists('C:\Program Files\LibreOffice\program\soffice.exe') or
                          FileExists('C:\Program Files (x86)\LibreOffice\program\soffice.exe');

  if (not ExcelInstalado) and (not LibreOfficeInstalado) then
  begin
    MsgBox(
      'Atención: No se ha detectado Microsoft Excel ni LibreOffice Calc en su sistema.' + #13#10 +
      'Se recomienda instalar una suite ofimática para habilitar la visualización e impresión ' +
      'de sus planillas y portadas oficiales.', 
      mbInformation, 
      MB_OK
    );
  end;
  
  Result := True;
end;

function IsNotUpdate(): Boolean;
begin
  Result := not PrevVersionInstalled;
end;

// ══════════════════════════════════════════════════════════════
//  2. DIÁLOGO PERSONALIZADO PARA ENTRADA DE CÉDULA
// ══════════════════════════════════════════════════════════════
function PedirCedula(var Cedula: String): Boolean;
var
  Form: TSetupForm;
  PromptLabel: TLabel;
  Edit: TEdit;
  OKButton: TButton;
  CancelButton: TButton;
begin
  Result := False;
  
  Form := TSetupForm.Create(nil);
  Form.ClientWidth := 400;
  Form.ClientHeight := 160;
  Form.Caption := 'Confirmación de Seguridad';
  Form.Position := poScreenCenter;

  PromptLabel := TLabel.Create(Form);
  PromptLabel.Parent := Form;
  PromptLabel.Left := 20;
  PromptLabel.Top := 15;
  PromptLabel.Width := 360;
  PromptLabel.Height := 45;
  PromptLabel.AutoSize := False;
  PromptLabel.WordWrap := True;
  PromptLabel.Caption := 'Para proceder con la desinstalación y proteger la privacidad de sus alumnos, ingrese la Cédula del Docente registrada:';

  Edit := TEdit.Create(Form);
  Edit.Parent := Form;
  Edit.Left := 20;
  Edit.Top := 75;
  Edit.Width := 360;

  OKButton := TButton.Create(Form);
  OKButton.Parent := Form;
  OKButton.Caption := 'Aceptar';
  OKButton.Default := True;
  OKButton.Left := 215;
  OKButton.Top := 115;
  OKButton.Width := 80;
  OKButton.ModalResult := mrOk;

  CancelButton := TButton.Create(Form);
  CancelButton.Parent := Form;
  CancelButton.Caption := 'Cancelar';
  CancelButton.Left := 300;
  CancelButton.Top := 115;
  CancelButton.Width := 80;
  CancelButton.ModalResult := mrCancel;

  if Form.ShowModal() = mrOk then
  begin
    Cedula := Edit.Text;
    Result := True;
  end;
end;

// ══════════════════════════════════════════════════════════════
//  3. DESINSTALACIÓN SEGURA — VERIFICACIÓN DEL PROPIETARIO
// ══════════════════════════════════════════════════════════════
function InitializeUninstall(): Boolean;
var
  StoredCedula: String;
  UserInput: String;
  ConfirmaBorrador: Integer;
begin
  Result := True;

  // Confirmación inicial
  ConfirmaBorrador := MsgBox(
    '¿Está seguro de que desea desinstalar RegistroDoc Pro?' + #13#10 + #13#10 +
    'ADVERTENCIA: Esta acción eliminará su base de datos local cifrada, configuraciones, ' +
    'auditorías e historial de asistencia de forma permanente.', 
    mbConfirmation, 
    MB_YESNO
  );

  if ConfirmaBorrador <> IDYES then
  begin
    Result := False;
    Exit;
  end;

  // Obtener Cédula desde el Registro de Windows
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Software\RegistroDocPro', 'Cedula', StoredCedula) then
  begin
    if StoredCedula <> '' then
    begin
      UserInput := '';
      if PedirCedula(UserInput) then
      begin
        if UserInput <> StoredCedula then
        begin
          MsgBox(
            'Error de Verificación: La Cédula ingresada no coincide con el dueño registrado. ' +
            'La desinstalación ha sido cancelada para evitar pérdida o robo de datos.', 
            mbError, 
            MB_OK
          );
          Result := False;
          Exit;
        end;
      end
      else
      begin
        // Cancelado por el usuario
        Result := False;
        Exit;
      end;
    end;
  end;
end;

// ══════════════════════════════════════════════════════════════
//  3. LIMPIEZA TOTAL — ARCHIVOS TEMPORALES Y AUDITORÍAS (NO EXPORTS)
// ══════════════════════════════════════════════════════════════
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataDir: String;
  AppDir: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Localizar el directorio AppData/Local/RegistroDoc del usuario activo (Base de datos y temporales)
    AppDataDir := ExpandConstant('{localappdata}\RegistroDoc');
    if DirExists(AppDataDir) then
    begin
      // Eliminar recursivamente la base de datos cifrada, logs de auditoría y archivos de trabajo
      DelTree(AppDataDir, True, True, True);
    end;

    // Localizar el directorio de la aplicación ({app}) que contiene config.enc, .env, etc. creados a posteriori
    AppDir := ExpandConstant('{app}');
    if DirExists(AppDir) then
    begin
      // Eliminar recursivamente el directorio de instalación
      DelTree(AppDir, True, True, True);
    end;
  end;
end;
