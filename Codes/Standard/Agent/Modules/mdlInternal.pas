unit mdlInternal;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type

enum eninfoosver
getplatformid := 1;
getmajorversion := 2;
getminorversion := 3;
getbuild := 4;
getcsdversion := 5;
end;
lnghookllkey : longint;

function sysinfogetos(optional : as):variant;
begin
dstosver : osversioninfo;lngiovn : longint;

lngiovn := infoosversionnumber;
dstosver.dwosversioninfosize := len(dstosver);
call getversionex(dstosver)

if  lngiovn > 0  then
begin
if  lngiovn = 1  then sysinfogetos = dstosver.dwplatformid;
if  lngiovn = 2  then sysinfogetos = dstosver.dwmajorversion;
if  lngiovn = 3  then sysinfogetos = dstosver.dwminorversion;
if  lngiovn = 4  then sysinfogetos = dstosver.dwbuildnumber;
if  lngiovn = 5  then sysinfogetos = dstosver.szcsdversion;
// left unchanged ==> exit;
end; //Main if block

case dstosver.dwmajorversion of 

4 :         Select Case DstOsVer.dwMinorVersion

0 :             Select Case DstOsVer.dwBuildNumber

950 :             

1111 :             

1381 :             

end;
10 :             Select Case DstOsVer.dwBuildNumber

1998 :             

2222 :             

end;
90 :         

end;
5 :         Select Case DstOsVer.dwMinorVersion

0 :         

1 :         

end;
end;
end;


public function sysinfogetname()
strnama : string;
strnama := string(255, chr(0));
getcomputername strnama, 255
strnama := left(strnama, instr(1, strnama, chr(0)) - 1);
sysinfogetname := strnama;
end;


procedure sysinfoname(netname : string);
begin
if  netname = ''  then exit;
computername netname
end;


procedure sysdisctlaltdel(opt : boolean);
begin
lngresult : longint;
lngresult := systemparametersinfo(spi_screensaverrunning, opt, vbnull, 0);
end;


public function sysnetgetmac() as string
s_mac : string;lastat : longint;
ncb : net_control_block;
ast : astat;








ncb.ncb_command := ncbre;//netbios 3.0 specifications in the ncb_callname field.
ncb.ncb_lana_num := 0;
call netbios(ncb)





ncb.ncb_callname := '*               ';//ncb.ncb_callname field (in a 16-chr string).
ncb.ncb_command := ncbastat;








ncb.ncb_lana_num := 0;//lana number to 0 (see the comments section below).
ncb.ncb_length := len(ast);
lastat := heapalloc(getprocessheap(), heap_generate_exceptions or heap_zero_memory, ncb.ncb_length);
if  lastat = 0  then exit;
ncb.ncb_buffer := lastat;
call netbios(ncb)

copymemory ast, ncb.ncb_buffer, len(ast)
s_mac := format$(hex(ast.adapt.adapter_address(0)), '00') + _;
format$(hex(ast.adapt.adapter_address(1)), '00') + _
format$(hex(ast.adapt.adapter_address(2)), '00') + _
format$(hex(ast.adapt.adapter_address(3)), '00') + _
format$(hex(ast.adapt.adapter_address(4)), '00') + _
format$(hex(ast.adapt.adapter_address(5)), '00')

heapfree getprocessheap(), 0, lastat
sysnetgetmac := s_mac;
end;


function sysdevprintersget() as string
strprinter : string;utprinter : printer;



strprinter := cmdsubput('total', printers.count);//  total|name|default|port|drivername|... and so on
for each utprinter in printers
strprinter := strprinter + cmdsubput('name', utprinter.devicename);
strprinter := strprinter + cmdsubput('default', utprinter.trackdefault);
strprinter := strprinter + cmdsubput('port', utprinter.port);
strprinter := strprinter + cmdsubput('drivername', utprinter.drivername);
strprinter := strprinter + cmdsubput('papersize', utprinter.papersize);
strprinter := strprinter + cmdsubput('orientation', utprinter.orientation);
end; //End for For or Do => Next/Loop
sysdevprintersget := strprinter;
end;


procedure sysdevmonitoroff(opcode : string);
begin
// try-except block should come here!!! => on error goto errint
l_retcmd : longint;





l_retcmd := choose(opcode, -1, 2, 1);// 3 = suspend
l_retcmd := sendmessage(frmmain.hwnd, $h112, &hf170, l_retcmd);
// left unchanged ==> exit;

errint:
apperrorlog err, 'module command | screenoff'
end;


procedure syswindowsexit(opcode : string);
begin
lngresult : longint;




if  sysinfogetos(getplatformid) = 2  then call secadjusttoken;
case mid(opcode, of 

0 :         LngResult = ExitWindowsEx(EWX_SHUTDOWN, 0)

1 :         LngResult = ExitWindowsEx(EWX_SHUTDOWN Or EWX_FORCE, 0)

2 :         LngResult = ExitWindowsEx(EWX_REBOOT, 0)

3 :         LngResult = ExitWindowsEx(EWX_REBOOT Or EWX_FORCE, 0)

end;//3 = force reboot
end;


procedure syswindowssleep(opcode : string);
begin
// try-except block should come here!!! => on error goto errint
if  opcode = '1'  then
begin
systempowerstate 1, 1
netsend '/info.me:sleep'
else 
systempowerstate 0, 0
netsend '/info.me:wakeup'
end; //Main if block
// left unchanged ==> exit;

errint:
apperrorlog err, 'module command | tidur'
end;

procedure syswindowshook(install : boolean);
begin
if  install = true  then
begin
lnghookllkey := windowshookex(wh_keyboard_ll, addressof proclowlevelkeyboard, app.hinstance, 0);
else 
unhookwindowshookex lnghookllkey
end; //Main if block
end;

procedure syswindowsblockinput(opcode : string);
begin
// try-except block should come here!!! => on error goto errint
op : string;
op := opcode;

case op of 

1 :         BlockInput True

netsend '/info.me:block'
2 :         BlockInput False

netsend '/info.me:unblock'
end;
// left unchanged ==> exit;

errint:
apperrorlog err, 'module command | blockinput'
end;


procedure sysshellhide(opcode : string);
begin
hwndtsk := findshelltaskbar;
hwnddsk := findshellwindow;

case opcode of 

0 :         HideShowWindow hWndtsk, True

hideshowwindow hwnddsk, true
1 :         HideShowWindow hWndtsk

hideshowwindow hwnddsk
end;
end;


procedure sysshelllock(opcode : string);
begin
lngval : longint;

case opcode of 

0 :         If LngStatusLock = 0 Then Exit Sub

lngstatuslock := 0;
minallwindow false
call hidedesktop
if  lngenvplatformid = 2  then
begin
call syswindowshook(false)
else 
systemparametersinfo spi_screensaverrunning, 0, lngval, 0+
end; //Main if block
unload frmkey
1 :         If LngStatusLock = 1 Then Exit Sub

lngstatuslock := 1;
minallwindow true
call hidedesktop(true)
if  lngenvplatformid = 2  then
begin
call syswindowshook(true)
else 
systemparametersinfo spi_screensaverrunning, 1, lngval, 0+
call deskwallprotect
end; //Main if block
// left unchanged ==> frmkey.show;
end;
call agentinfostatus
end;


end.

