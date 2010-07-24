unit MdlDesktop;

interface
  uses
    Windows;
  type
    Buffer = array[0..64] of char;

  Const
    DesktopName = 'CbDaemonShell';
    
  procedure DesktopSwitch;
  function DesktopGetName:String;


implementation

  procedure DesktopSwitch;
  var
    DefDesktop: HDESK;
  begin
    DefDesktop := OpenDesktop('Default',0,False,GENERIC_ALL);
    SetThreadDesktop(DefDesktop);
    SwitchDesktop(DefDesktop);
  end;

  function DesktopGetName:String;
  var
    IntHDesktop: HDESK;
    PBuffer: ^Buffer;
    StrBuffer: String;
    CrdBuffSize: Cardinal;
  begin
    CrdBuffSize := 64;
    IntHDesktop := OpenInputDesktop(0,False,DESKTOP_READOBJECTS);
    GetUserObjectInformation(IntHDesktop, UOI_NAME, PBuffer, CrdBuffSize, CrdBuffSize);
    StrBuffer := PBuffer^;
    DesktopGetName := StrBuffer;
    CloseHandle(IntHDesktop);
  end;
end.
