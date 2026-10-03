[Setup]
AppName=嘉立创库存计价器
AppVersion=1.0.1
AppPublisher=ax4987
DefaultDirName={autopf}\JlcInventory
DefaultGroupName=嘉立创库存计价器
OutputDir=C:\Users\Administrator\Desktop\嘉立创库存计价器\build
OutputBaseFilename=嘉立创库存计价器_v1.0.1
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin
SetupIconFile=C:\Users\Administrator\Desktop\嘉立创库存计价器\app.ico

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
Source: "C:\Users\Administrator\Desktop\嘉立创库存计价器\JlcInventory\JlcInventory\bin\Release\net8.0-windows\win-x64\publish\*"; Excludes: "inventory.db"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\嘉立创库存计价器"; Filename: "{app}\JlcInventory.exe"
Name: "{autodesktop}\嘉立创库存计价器"; Filename: "{app}\JlcInventory.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务:"

[Code]
var
  DataDirPage: TInputDirWizardPage;

procedure InitializeWizard;
begin
  // Create a custom page for the data directory
  DataDirPage := CreateInputDirPage(wpSelectDir,
    '选择数据存储位置', '您的库存数据库(inventory.db)将保存在哪里？',
    '请选择数据保存路径。推荐保存在非系统盘，以免重装系统丢失数据。'#13#10#13#10 +
    '点击“下一步”继续。',
    False, '');
  
  // Set default data path (e.g. AppData or Documents)
  DataDirPage.Add('数据文件夹路径:');
  DataDirPage.Values[0] := ExpandConstant('{userdocs}\JlcInventory_Data');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ConfigPath: String;
  ConfigContent: String;
  SelectedDataPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Get the selected data path
    SelectedDataPath := DataDirPage.Values[0];
    
    // Replace backslashes with double backslashes for JSON
    StringChangeEx(SelectedDataPath, '\', '\\', True);
    
    ConfigPath := ExpandConstant('{app}\config.json');
    ConfigContent := '{ "DataPath": "' + SelectedDataPath + '" }';
    
    SaveStringToFile(ConfigPath, ConfigContent, False);
  end;
end;
