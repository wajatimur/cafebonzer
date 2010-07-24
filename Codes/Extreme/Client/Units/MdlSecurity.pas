unit MdlSecurity;

interface
  uses
    Windows;

  procedure SecurityActivate;
  procedure SecAdjustToken;
  function SecPassCheck(StrPassword:String):Integer;

implementation
  uses
    MdlGlobal, MdlSettings, MdlSystem;


  procedure SecurityActivate;
  begin
    if SettingGet('SysAutoLock') = '1' then begin
      if SettingGet('LOGIN') = 'False' then SysShellLock(1);
    end;
    { INVISIBLE }
    { DESKTOP ICON }
    { WALLPAPER }
  end;


  procedure SecAdjustToken;
  var
    HProcess: Integer;
    HToken, IntBuffer: Cardinal;
    IntsLUID: Int64;
    DtpTokenPriv, DtpNewTokenPriv: TOKEN_PRIVILEGES;
  begin
    HProcess := GetCurrentProcess;
    OpenProcessToken(HProcess,TOKEN_ADJUST_PRIVILEGES or TOKEN_QUERY,HToken);
    LookupPrivilegeValue('','SeShutdownPrivilege',IntsLUID);

    DtpTokenPriv.PrivilegeCount := 1;
    DtpTokenPriv.Privileges[0].Luid := IntsLUID;
    DtpTokenPriv.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED;

    AdjustTokenPrivileges(HToken,False,DtpTokenPriv,SizeOf(DtpNewTokenPriv),DtpNewTokenPriv,IntBuffer);
  end;


  function SecPassCheck(StrPassword:String):Integer;
  var
    StrPassReg:String;
  begin
    { NOTE :
      1 = Granted }
    StrPassReg := SettingGet('GenAdminPass');
    if StrPassword = StrPassReg then Result := 1;
  end;
end.
