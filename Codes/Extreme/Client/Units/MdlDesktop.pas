{~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 Desktop Module
 Azri Jamil - 2004
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~}
unit MdlDesktop;

interface
  uses
    Windows;
  type
    Buffer = array[0..64] of char;

  var
    HDesktopNew, HDesktopOldThread, HDesktopOldInput: Integer;

  procedure DesktopCreate(DesktopName:String);
  procedure DesktopSwitch(DefaultDesktop:Boolean);
  procedure DesktopClear;
  function DesktopGetName:String;


implementation

  procedure DesktopCreate(DesktopName:String);
  begin
    HDesktopOldThread := GetThreadDesktop(GetCurrentThreadId);
    HDesktopOldInput := OpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP);
    HDesktopNew := CreateDesktop(PChar(DesktopName),nil,nil,0,GENERIC_ALL,nil);
  end;

  procedure DesktopSwitch(DefaultDesktop:Boolean);
  var
    DefDesktop: HDESK;
  begin
    if DefaultDesktop = False then begin
      if HDesktopNew <> 0 then begin
        SetThreadDesktop(HDesktopNew);
        SwitchDesktop(HDesktopNew);
      end;
    end else begin
      DefDesktop := OpenDesktop('Default',0,False,GENERIC_ALL);
      SetThreadDesktop(DefDesktop);
      SwitchDesktop(DefDesktop);
    end;
  end;

  procedure DesktopClear;
  begin
    if HDesktopOldInput <> 0 then begin
      SwitchDesktop(HDesktopOldInput);
      HDesktopOldInput := 0;
    end;
    if HDesktopOldThread <> 0 then begin
      SetThreadDesktop(HDesktopOldThread);
      HDesktopOldThread := 0;
    end;
    if HDesktopNew <> 0 then begin
      CloseDesktop(HDesktopNew);
      HDesktopNew := 0;
    end;
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
