unit gui;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type





procedure enablegroup(grpname : string;optional : as);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error resume next
for each control in frmmain.controls
if  control.tag = grpname  then control.enabled = enable;
end;//Current For/Next
end;







function enumfontproc(byval : as;byval : as;byval : as;byval : as):long;
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
lfret : logfont;s_fntname : string;

copymemory lfret, byval lplf, lenb(lfret)
s_fntname := strconv(lfret.lffacename, vbunicode);
s_fntname := trim(s_fntname);

frmpickfont.fntlist.additem s_fntname
enumfontproc := 1;
end;





procedure drawborder(hwnd : longint);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
stle : longint;


stle := getwindowlong(hwnd, gwl_style);//{ buang 'border' asal }'
stle := stle and not ws_border;
windowlong hwnd, gwl_style, stle


stle := getwindowlong(hwnd, gwl_exstyle);//{ set 'style' baru }'
stle := stle or ws_ex_staticedge;
windowlong hwnd, gwl_exstyle, stle

end;






procedure putontop(hwnd : longint);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
i : longint;
i := windowpos(hwnd, hwnd_topmost, 0, 0, 0, 0, swp_wndflags);
end;

function gettitle(hwnd : longint):string;
begin
sbuffer : string;
getwindowtext hwnd, sbuffer, 64
gettitle := left$(sbuffer, instr(1, sbuffer, chr(0)) - 1);
end;

function getclass(hwnd : longint):string;
begin
sbuffer : string;
getclassname hwnd, sbuffer, 64
getclass := left$(sbuffer, instr(1, sbuffer, chr(0)) - 1);
end;






public function findshelltaskbar() as long//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
hwnd : longint;
// try-except block should come here!!! => on error resume next
hwnd := findwindowex(0+, 0&, g_cstrshelltaskbarwnd, vbnullstring);
if  hwnd <> 0  then
begin
findshelltaskbar := hwnd;
end; //Main if block
end;






public function findshellwindow() as long//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
hwnd : longint;
// try-except block should come here!!! => on error resume next
hwnd := findwindowex(0+, 0&, g_cstrshellviewwnd, vbnullstring);
if  hwnd <> 0  then
begin
findshellwindow := hwnd;
end; //Main if block
end;






procedure hidedesktop(optional : as);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
lngshowcmd : longint;
// try-except block should come here!!! => on error resume next
if  hide = true  then
begin
lngshowcmd := sw_hide;
else 
lngshowcmd := sw_show;
end; //Main if block
call showwindow(findshellwindow, lngshowcmd)
call showwindow(findshelltaskbar, lngshowcmd)
end;






procedure hideshowwindow(byval : as;optional : hide);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
lngshowcmd : longint;
// try-except block should come here!!! => on error resume next
if  hide = true  then
begin
lngshowcmd := sw_hide;
else 
lngshowcmd := sw_show;
end; //Main if block
call showwindow(hwnd, lngshowcmd)
end;






procedure hideallwindow(hide : boolean);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
hwndcur : longint;

hwndcur := getwindow(frmkey.hwnd, gw_hwndfirst);

do while hwndcur
if  istaskwindow(hwndcur) = true and hwndcur <> frmkey.hwnd  then
begin
if  hide = true  then
begin
showwindow hwndcur, sw_hide
else 
showwindow hwndcur, sw_show or sw_shownormal
end; //Main if block
end; //Main if block

hwndcur := getwindow(hwndcur, gw_hwndnext);
end; //End for For or Do => Next/Loop
end;






procedure minallwindow(minimize : boolean);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
hwnd := findwindow('shell_traywnd', vbnullstring);
if  minimize = true  then
begin
postmessage hwnd, wm_command, min_all, 0+
else 
postmessage hwnd, wm_command, min_all_undo, 0+
end; //Main if block
end;






procedure hideactivewin(hide : boolean);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
hwnd := getactivewindow;
if  hide  then
begin
windowpos hwnd, 0, 0, 0, 0, 0, swp_nozorder or swp_nomove or swp_nosize or swp_hidewindow
else 
windowpos hwnd, 0, 0, 0, 0, 0, swp_nozorder or swp_nomove or swp_nosize or swp_showwindow
end; //Main if block
end;






function istaskwindow(hwnd : longint):boolean;
begin
lngstyle : longint;istask : longint;
istask := ws_visible or ws_border;//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
lngstyle := getwindowlong(hwnd, gwl_style);
if  (lngstyle and istask) = istask  then istaskwindow = true;
end;






procedure formtrap(curform : form;trap : boolean);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
x : longint;y : longint;erg : longint;
newrect : rect;
deskrect : rect;
getwindowrect getdesktopwindow, deskrect

x+ := screen.twipsperpixelx;
y+ := screen.twipsperpixely;

if  trap = true  then
begin
with newrect
.left := curform.left / x+ '- 8;
.top := curform.top / y+ ' - 8;
.right := .left + (curform.width / x+) '- 14;
.bottom := .top + (curform.height / y+) '- 15;
end with
else 
with newrect
.left := 0+;
.top := 0+;
.right := deskrect.right;
.bottom := deskrect.bottom;
end with
end; //Main if block

erg+ := clipcursor(newrect);
end;


public function getwallpaper() as string
getwallpaper := getstring(hkey_current_user, 'control panel\desktop', 'wallpaper');
end;


end.

