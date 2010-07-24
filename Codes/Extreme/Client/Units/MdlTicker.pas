unit MdlTicker;

interface

uses
  Windows, SysUtils, Graphics, ShellApi, Forms, Messages;

  procedure TickerStart;
  procedure TickerStop;
  procedure TickerNormal;
  procedure TickerHover;
  procedure TickerIcon(Add:Boolean);
  procedure TickerCheck;
  procedure FindTaskbar;
  procedure GetTraySize;
  function GetTaskbarEdge:Integer;
  function GetTrayIconRow:Integer;

var
  IntIconCount: Integer;
  HwndDesktop, HwndTaskbar, HwndTray, HwndClock: HWND;
  IntTickWidth, IntTickHeight: Integer;
  IntLastWidth, IntLastHeight: Integer;
  IntLastEdge: Integer;


  
implementation

uses MdlSettings, Ticker, MdlUtils, Config;


procedure TickerStart;
var
  IntEdge,IntIconRow: Integer;
begin
  
  { Load Configuration }
  IntIconCount := StrToInt(SettingGet('TickGuiSize'));
  //FrmTicker.MainScroll.Color := TColor(SettingGet('TickGuiBackColor'));
  //FrmTicker.MainScroll.Font.Color := TColor(SettingGet('TickGuiForeColor'));
  FrmTicker.MainScroll.Font.Name := SettingGet('TickGuiFont');
  //FrmTicker.MainScroll.Font.Size := StrToInt(SettingGet('TickGuiFontSize'));

  { General Metric }
  FindTaskbar;
  GetTraySize;
  IntTickHeight := GetSystemMetrics(SM_CYCAPTION)-2;
  IntTickWidth := IntIconCount * IntTickHeight;

  { Check Edge, Draw Ticker }
  IntLastEdge := GetTaskbarEdge;
  IntEdge := GetTaskbarEdge;
  IntIconRow := GetTrayIconRow;
  if (IntEdge <> ABE_LEFT) and (IntEdge <> ABE_RIGHT) and (IntIconRow = 1) then
    TickerNormal
  else
    TickerHover;

  FrmTicker.Show;
  FrmTicker.TmrCheck.Enabled := True;
  ShowWindow(Application.Handle, SW_HIDE);
end;


procedure TickerStop;
begin
  SetParent(FrmTicker.Handle,0);
  SendMessage(HwndTaskbar,WM_SETREDRAW,0,0);
  TickerIcon(False);
  SendMessage(HwndTaskbar,WM_SETREDRAW,1,0);
  RedrawWindow(HwndTaskbar,NIL,0,RDW_INVALIDATE or RDW_ALLCHILDREN or RDW_UPDATENOW);
  FrmTicker.Close;
end;


procedure TickerNormal;
begin
  SendMessage(HwndTaskbar,WM_SETREDRAW,0,0);
  TickerIcon(False);
  TickerIcon(True);
  SendMessage(HwndTaskbar,WM_SETREDRAW,1,0);
  RedrawWindow(HwndTaskbar,NIL,0,RDW_INVALIDATE or RDW_ALLCHILDREN or RDW_UPDATENOW);

  SetParent(FrmTicker.Handle,HwndTray);
  SetWindowPos(FrmTicker.Handle,0,1,1,IntTickWidth,IntTickHeight,SWP_NOZORDER);
end;


procedure TickerHover;
var
  DtpDeskRect: TRect;
  IntEdge, IntTop, IntLeft: Integer;
begin
  SendMessage(HwndDesktop,WM_SETREDRAW,0,0);
  TickerIcon(False);
  GetClientRect(HwndDesktop,DtpDeskRect);

  IntEdge := GetTaskbarEdge;
  if IntEdge = ABE_LEFT then begin
    IntTop := DtpDeskRect.Bottom - IntTickHeight;
    IntLeft := DtpDeskRect.Right - IntTickWidth;
  end
  else if IntEdge = ABE_RIGHT then begin
    IntTop := DtpDeskRect.Bottom - IntTickHeight;
    IntLeft := 0;
  end;

  SetParent(FrmTicker.Handle,HwndDesktop);
  SetWindowPos(FrmTicker.Handle,HWND_TOPMOST,IntLeft,IntTop,IntTickWidth,IntTickHeight,0);

  SendMessage(HwndDesktop,WM_SETREDRAW,1,0);
  RedrawWindow(HwndDesktop,NIL,0,RDW_INVALIDATE or RDW_ALLCHILDREN or RDW_UPDATENOW);
end;


procedure TickerIcon(Add:Boolean);
var
  DtpNid: TNotifyIconData;
  IntIdx: Integer;
begin
  DtpNid.cbSize := SizeOf(DtpNid);
  DtpNid.Wnd := FrmTicker.Handle;
  if Add = True then begin
    DtpNid.uCallbackMessage := WM_MOUSEMOVE;
    DtpNid.uFlags := NIF_MESSAGE;
    for IntIdx := 1 to IntIconCount do begin
      DtpNid.uID := IntIdx;
      Shell_NotifyIcon(NIM_ADD,Addr(DtpNid));
    end;
  end else begin
    for IntIdx := 1 to IntIconCount do begin
      DtpNid.uID := IntIdx;
      Shell_NotifyIcon(NIM_DELETE,Addr(DtpNid));
    end;
  end;
end;


procedure TickerCheck;
var
  DtpTrayRect: TRect;
  IntCurEdge: Integer;
begin
  IntCurEdge := GetTaskbarEdge;
  case IntCurEdge of
    ABE_BOTTOM,ABE_TOP:begin
      if IntCurEdge <> IntLastEdge then
        TickerNormal
      else begin
        GetClientRect(HwndTray,DtpTrayRect);
        if GetTrayIconRow = 1 then
          if (DtpTrayRect.Bottom < IntLastHeight) or (DtpTrayRect.Right < IntLastWidth) then
            TickerNormal
        else
          if (DtpTrayRect.Bottom < IntLastHeight) or (DtpTrayRect.Right < IntLastWidth) then
            TickerHover;
      end;
    end;
    ABE_LEFT,ABE_RIGHT:begin
      if IntLastEdge <> IntCurEdge then Tickerhover;
    end;
  end;
  IntLastEdge := IntCurEdge;
  IntLastHeight := DtpTrayRect.Bottom;
  IntLastWidth := DtpTrayRect.Right;
end;


procedure FindTaskbar;
begin
  HwndDesktop := GetDesktopWindow;
  HwndTaskbar := FindWindow('Shell_TrayWnd','');

  if HwndTaskbar > 0 then begin
    HwndTray := FindWindowEx(HwndTaskbar,0,'TrayNotifyWnd','');
    if HwndTray > 0 then
      HwndClock := FindWindowEx(HwndTray,0,'TrayClockWClass','');
  end;
end;


procedure GetTraySize;
var
  DtpTrayRect: TRect;
begin
  GetClientRect(HwndTray,DtpTrayRect);
  IntLastWidth := DtpTrayRect.Right;
  IntLastHeight := DtpTrayRect.Bottom;
end;


function GetTaskbarEdge:Integer;
var
  DtpAbd: TAppBarData;
begin
  DtpAbd.cbSize := SizeOf(DtpAbd);
  DtpAbd.hWnd := HwndTaskbar;

  SHAppBarMessage(ABM_GETTASKBARPOS,DtpAbd);
  Result := DtpAbd.uEdge;
end;


function GetTrayIconRow:Integer;
var
  DtpTrayRect: TRect;
  IntIconHeight: Integer;
begin
  GetClientRect(HwndTray,DtpTrayRect);

  {if IsWindowVisible(HwndClock) then begin
    GetClientRect(HwndClock,DtpClockRect);
    MapWindowPoints(0,HwndTray,DtpClockRect,2);
    if DtpClockRect.Top <> 0 then begin
      DtpClockRect.Top := 0;
      DtpClockRect.Bottom := 0;
    end;
  end;}
  IntIconHeight := GetSystemMetrics(SM_CYCAPTION)-3;
  Result := DtpTrayRect.Bottom div IntIconHeight;
end;

end.
