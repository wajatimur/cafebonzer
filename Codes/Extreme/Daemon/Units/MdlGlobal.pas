unit MdlGlobal;

interface
  uses
    Windows, Messages, SysUtils, StrUtils, Variants, Classes,
    Graphics, Controls, Forms, Dialogs, Registry;

  const
    CStrPipeServerName = 'CbDaemon';
    CStrPipeHostName = 'CbAgent';

  procedure NetPipeSend(Data:String);
  procedure SysShutdown(ShutdownCode:Integer; Force:Boolean);
  function SysSettingGet(SettingName:String):String;
  procedure SysSettingSet(SettingName:String; SettingValue: String);


implementation
  uses
    MdlSecurity, MdlPipes;


  procedure NetPipeSend(Data:String);
  begin
    with TPipeClient.Create(CStrPipeHostName) do
    try
      SendString(Data);
    finally
      Free;
    end;
  end;

  procedure SysShutdown(ShutdownCode:Integer; Force:Boolean);
  begin
    SecAdjustToken();
    if Force = False then
      ExitWindowsEx(ShutdownCode,0)
    else
      ExitWindowsEx(ShutdownCode or EWX_FORCE,0);
  end;

  function SysSettingGet(SettingName:String):String;
  var
    CReg: TRegistry;
  begin
    CReg := TRegistry.Create(KEY_READ);
    try
      CReg.RootKey := HKEY_LOCAL_MACHINE;
      CReg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Security',True);
      SysSettingGet := CReg.ReadString(SettingName);
    finally
      CReg.Free;
    end;
  end;

  procedure SysSettingSet(SettingName:String; SettingValue: String);
  var
    CReg: TRegistry;
  begin
    CReg := TRegistry.Create(KEY_WRITE);
    try
      CReg.RootKey := HKEY_LOCAL_MACHINE;
      CReg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Security',True);
      CReg.WriteString(SettingName,SettingValue);
    finally
      CReg.Free;
    end;
  end;

end.
