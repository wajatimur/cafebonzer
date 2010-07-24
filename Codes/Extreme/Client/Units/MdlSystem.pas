unit MdlSystem;

interface
  uses
    Windows, Classes, Forms, SysUtils, Printers, JclSysInfo;

  function SysDevPrintersGet:String;
  function SysNetGetMac:String;
  function SysInfoGetOsMajor:integer;
  function SysInfoGetName:String;
  procedure SysCommand(CmdCode,KeyData:String);
  procedure SysShellLock(Code: integer);
  procedure SysShutdown(ShutdownCode:Integer; Force:Boolean);

  

implementation
  uses
    Main, MdlCommands, MdlAgent, MdlDesktop, MdlSecurity;


  function SysDevPrintersGet:String;
  var
    CPrn: TPrinter;
    IntIdx: Integer;
    StrPrinter, StrPrinters: String;
  begin
    CPrn := TPrinter.Create;
    try
      for IntIdx := 0 to CPrn.Printers.Count-1 do begin
        StrPrinter := CPrn.Printers.Strings[IntIdx];
        StrPrinters := StrPrinters + CmdKeySet('PRINTER'+IntToStr(IntIdx),StrPrinter);
      end;
    finally
      CPrn.Free;
    end;
    Result := StrPrinters;
  end;


  function SysNetGetMac:String;
  var
    TslMac: TStringList;
    IntMacCount, IntIdx: Integer;
    StrMacs, StrMac: String;
  begin
    TslMac := TStringList.Create;
    try
      IntMacCount := GetMacAddresses('',TslMac);
      for IntIdx := 0 to IntMacCount-1 do begin
        StrMac := TslMac.Strings[IntIdx];
        StrMacs := StrMacs + CmdKeySet('NETMAC'+String(IntIdx),StrMac);
      end;
    finally
      TslMac.Free;
    end;
  end;


  function SysInfoGetOsMajor:integer;
  begin
    { Get Major Version Only }
    if GetWindowsVersion = wvWinXp then
      Result := 5
    else
      Result := 4;
  end;


  function SysInfoGetname:String;
  begin
    Result := GetLocalComputerName;
  end;

  
  procedure SysCommand(CmdCode,KeyData:String);
  var
    IntTemp:Integer;
  begin
    if CmdCode = '0010' then begin
      IntTemp := StrToInt(CmdKeyGet(KeyData,'ACTION'));
      SysShutdown(IntTemp,False);
    end;
    if CmdCode = '0020' then begin
      IntTemp := StrToInt(CmdKeyGet(KeyData,'ACTION'));
      SysShellLock(IntTemp);
    end;
  end;


  procedure SysShellLock(Code: integer);
  begin
    case Code of
      0:begin   { UNLOCK }
        DesktopSwitch(True);
      end;
      1:begin   { LOCK }
        DesktopSwitch(False);
      end;
    end;
    IntStatusLock := Code;
    AgentInfoStatus;
  end;


  procedure SysShutdown(ShutdownCode:Integer; Force:Boolean);
  begin
    SecAdjustToken();
    if Force = False then
      ExitWindowsEx(ShutdownCode,0)
    else
      ExitWindowsEx(ShutdownCode or EWX_FORCE,0);
  end;
end.
