unit MdlGlobal;

interface
  uses
    Windows, Messages, Dialogs, Forms, SysUtils, JvCreateProcess,
    JclFileUtils, JclRegistry;

  var
    StrAppVersion, StrAppBuild: String;
    IntSysPlatformId: integer;
    CJvCps: TJvCreateProcess;

  const
    CStrSettingPath = 'Software\Microsoft\Windows\CurrentVersion\Security';
    CStrAutoStartPath = 'Software\Microsoft\Windows\CurrentVersion\Run';
    CStrPipeServerName = 'CbAgent';
   { DAEMON }
    CStrPipeHostName = 'CbDaemon';
    CStrDaemonExe = 'CbDaemon.exe';
    CStrDaemonDesktop = 'CbDaemonShell';


  procedure MainStart;
  procedure AppClose;
  procedure AppFirstLoad;
  procedure AppSetEnv;
  procedure AppSetPersist;
  function AppInfoGetVersion: String;
  procedure StartProcess(StrAppPath,StrDesktop:String);



implementation
  uses
    MdlSettings, MdlSecurity, MdlCommunication, MdlTicker, Main,
    MdlDesktop, Config, MdlSystem;

  procedure MainStart;
  begin
    AppFirstLoad;     { Check For First Load }
    AppSetEnv;        { Set\Load Environment }
    AppSetPersist;    { Persistent Settings }
    SecurityActivate; { Activate Agent & System Security }
    NetConnect;       { Start Connecting }
  end;


  procedure AppClose;
  begin
    NetDisconnect;            { Disconnect Network }
    TickerStop;               { Stop Ticker }
    FrmMain.CPServer.Destroy; { Destroy Pipe Server }
    CJvCps.Terminate;         { Terminate CbDaemon }
    Application.Terminate;    { Self Terminating }
  end;


  procedure AppFirstLoad;
  begin
    if SettingGet('AppFirstTime') = '' then begin
      FrmOption.SetFirstTime;
      FrmOption.ShowModal;
      Application.Terminate; 
    end;
  end;


  procedure AppSetEnv;
  var
    StrDaemonPath: String;
  begin
   { GLOBAL VARIABLE }
    StrAppVersion := AppInfoGetVersion;
    IntSysPlatformId := SysInfoGetOsMajor;
   { DAEMON }
    StrDaemonPath := ExtractFilePath(Application.ExeName)+CStrDaemonExe;
    DesktopCreate(CStrDaemonDesktop);{CbDaemonShell}
    StartProcess(StrDaemonPath,CStrDaemonDesktop);
  end;


  procedure AppSetPersist;
  begin
    if SettingGet('TickGuiDisable') = '0' then TickerStart;
    if SettingGet('AppAutoStart') = '1' then begin
      RegWriteString(HKEY_CURRENT_USER,CStrAutoStartPath,'CbClient','CbClientEx.exe');
    end;
  end;


  function AppInfoGetVersion: String;
  var
    CFile: TJclFileVersionInfo;
  begin
    CFile := TJclFileVersionInfo.Create(Application.ExeName);
    Result := CFile.FileVersion;
    CFile.Free;
  end;


  procedure StartProcess(StrAppPath,StrDesktop:String);
  begin
    CJvCps := TJvCreateProcess.Create(Application);
    CJvCps.StartupInfo.Desktop := StrDesktop;
    CJvCps.ApplicationName := StrAppPath;
    CJvCps.WaitForTerminate := False;
    CJvCps.Run;
  end;
end.
