unit MdlSecurity;

interface
  uses
    Windows;

  procedure SecAdjustToken;

  
implementation

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
end.
