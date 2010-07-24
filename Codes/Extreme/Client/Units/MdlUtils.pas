unit MdlUtils;

interface

uses
  Windows, ShellApi;

  function AppBarIsAutoHide(Handle:HWND):Boolean;


  
implementation


function AppBarIsAutoHide(Handle:HWND):Boolean;
var
  DtpAbd: TAppBarData;
begin
  DtpAbd.cbSize := SizeOf(DtpAbd);
  DtpAbd.hWnd := Handle;
  Result := Boolean(SHAppBarMessage(ABM_GETSTATE,DtpAbd) and ABS_AUTOHIDE);
end;

end.
