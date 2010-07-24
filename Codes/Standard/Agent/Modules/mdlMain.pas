unit mdlMain;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type




hdesktopwnd : longint;//public variables
htaskbarwnd : longint;// taskbar window handle
htraywnd : longint;// tray window handle
hclockwnd : longint;// clock window handle
lastedge : shappbar_edges;// last checked edge where the taskbar was
lastwidth : longint;// last checked tray width
lastheight : longint;// last checked tray height


tickerwidth : longint;tickerheight : longint;
iconcount : longint;//private variables
ishidden : boolean;


// left unchanged ==> private const tt01 = 'please resize your taskbar';//private const








procedure iconsadd();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
nid : notifyicondataa;i : longint;
with nid
.cbsize := lenb(nid);
.hwnd := frmticker.hwnd;
.ucallbackmessage := wm_mousemove;
.uflags := nif_message;
end with

for i:=1 to iconcount do
begin
nid.uid := i;
shell_notifyicon nim_add, nid
end; //End for For or Do => Next/Loop
end;








procedure findtaskbar();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

htaskbarwnd := findwindow('shell_traywnd', vbnullstring);// find taskbar handle

if  htaskbarwnd  then
begin

htraywnd := findwindowex(htaskbarwnd, 0, 'traynotifywnd', vbnullstring);// find tray handle (anak kepada taskbar)
if  htraywnd  then
begin

hclockwnd := findwindowex(htraywnd, 0, 'trayclockwclass', vbnullstring);// find clock handle (anak kepada tray)
if  hclockwnd = 0  then err.raise vbobjecterror + 2;
else 
err.raise vbobjecterror + 1
end; //Main if block
else 
err.raise vbobjecterror
end; //Main if block


hdesktopwnd := getdesktopwindow;// find desktop handle
end;










procedure tickerhide(optional : as);
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
nid : notifyicondataa;

if  ishidden  then exit;
ishidden := true;// check if already hidden

// left unchanged ==> iconsremove;// remove the icons

windowpos frmticker.hwnd, 0, -2000, -2000, 0, 0, swp_nosize or swp_nozorder// move ticker outside

if  addicon = true  then
begin

with nid// add a "standard" icon
.cbsize := lenb(nid);
.hwnd := frmtray.hwnd;
.uid := -100;
.hicon := frmticker.icon.handle;
.sztip := tt01;
.ucallbackmessage := wm_mousemove;
.uflags := nif_icon or nif_message or nif_tip;
end with
shell_notifyicon nim_add, nid
end; //Main if block
end;








procedure tickerfly();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=






deskrect : rect;edge : shappbar_edges;taskpos : longint;
theight : longint;twidth : longint;ttop : longint;tright : longint;

ishidden := false;// 4 = right
sendmessage hdesktopwnd, wm_redraw, 0, byval 0+


edge := gettaskbaredge();//the edge
case edge of 

abe_top: :     Case ABE_BOTTOM: TaskPos = 2

abe_left: :     Case ABE_RIGHT: TaskPos = 4

end;


getclientrect getdesktopwindow, deskrect//ticker size other metric size

theight := getsystemmetrics(sm_cycaption);
twidth := theight * iconcount;
ttop := deskrect.bottom - (theight * 4);
tright := deskrect.right - (twidth + 2);

if  taskpos = 3 or taskpos = 1  then ttop = deskrect.bottom - theight;
if  taskpos = 4  then ttop = deskrect.bottom - theight: tright = 0;


parent frmticker.hwnd, hdesktopwnd//setting the parents and moving..
windowpos frmticker.hwnd, hwnd_topmost, tright, ttop, twidth, theight, 0+


sendmessage hdesktopwnd, wm_redraw, 1, byval 0+//redraw !
redrawwindow hdesktopwnd, byval 0+, 0&, rdw_invalidate or rdw_allchildren or rdw_updatenow
end;








procedure tickershow();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
clkr : rect;tryr : rect;
nid : notifyicondataa;


with nid//remove the "standard" icon
.cbsize := lenb(nid);
.hwnd := frmticker.hwnd;
.uid := -100;
end with

shell_notifyicon nim_delete, nid
ishidden := false;


if  iswindowvisible(hclockwnd)  then
begin

getwindowrect hclockwnd, clkr//get clock rect

clkr.right := clkr.right - clkr.left;//calculate clock width
end; //Main if block




getclientrect htraywnd, tryr// get tray client rect


parent frmticker.hwnd, htraywnd// makesure the parent is tray

if  isautohide()  then
begin

windowpos frmticker.hwnd, 0, tryr.right - clkr.right, 1, 0, 0, swp_nosize or swp_nozorder// move the ticker.
else 

windowpos frmticker.hwnd, 0, tryr.right - clkr.right - tickerwidth + 1, 1, 0, 0, swp_nosize or swp_nozorder// move the ticker.
end; //Main if block
end;








private function isautohide() as boolean//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
abd : appbardata;
with abd
.cbsize := lenb(abd);
.hwnd := htaskbarwnd;
end with
isautohide := shappbarmessage(abm_getstate, abd) and abs_autohide;
end;








procedure tickerresize();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=


sendmessage htaskbarwnd, wm_redraw, 0, byval 0+// stop the updating of taskbar



// left unchanged ==> iconsremove;// other icons
// left unchanged ==> iconsadd;


tickerheight := getsystemmetrics(sm_cycaption) - 2;// the size of icons is the same as title bar system menu icon
tickerwidth := tickerheight * iconcount;


windowpos frmticker.hwnd, 0, 0, 0, tickerwidth, tickerheight, swp_nomove or swp_nozorder// change ticker with and height


sendmessage htaskbarwnd, wm_redraw, 1, byval 0+// redraw task bar
redrawwindow htaskbarwnd, byval 0+, 0&, rdw_invalidate or rdw_allchildren or rdw_updatenow
end;








procedure iconsremove();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
nid : notifyicondataa;i : longint;
with nid
.cbsize := lenb(nid);
.hwnd := frmticker.hwnd;
end with
for i:=1 to iconcount do
begin
nid.uid := i;
shell_notifyicon nim_delete, nid
end; //End for For or Do => Next/Loop
end;








procedure traystart();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
nid : notifyicondataa;


ishidden := true;//the ticker is hidden (variable is global)
with nid
.cbsize := lenb(nid);
.hwnd := frmtray.hwnd;
.uid := 1;
.hicon := frmtray.icon.handle;
.ucallbackmessage := wm_mousemove;
.uflags := nif_icon or nif_message or nif_tip;
end with

shell_notifyicon nim_add, nid
end;





procedure trayremove();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
nid : notifyicondataa;


ishidden := false;//ticker now is visible, global juga

with nid
.cbsize := lenb(nid);
.uid := 1;
.hwnd := frmtray.hwnd;
end with

shell_notifyicon nim_delete, nid
end;





procedure tickerstart();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
edge : shappbar_edges;stmpret : string;

// left unchanged ==> findtaskbar;

stmpret := get('tickguisize');
if  stmpret = ''  then
begin
if  iconcount = 0  then iconcount = 7;
else 
iconcount := stmpret;
end; //Main if block

lastedge := gettaskbaredge();
gettraysize lastwidth, lastheight
load frmticker

frmticker.picticker.fontname := get('tickguifont', 'verdana');
frmticker.picticker.fontsize := get('tickguifontsize', 8);
frmticker.picticker.forecolor := get('tickguiforecolor', vbblack);
frmticker.picticker.backcolor := get('tickguibackcolor', $he0e0e0);

parent frmticker.hwnd, htraywnd

edge := gettaskbaredge();
if  edge <> abe_left and edge <> abe_right and trayiconrows() = 1  then
begin
// left unchanged ==> tickerresize;
// left unchanged ==> tickershow;
else 
tickerhide false
// left unchanged ==> tickerfly;
end; //Main if block

// left unchanged ==> frmticker.show;
end;





procedure tickerstop();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'

parent frmticker.hwnd, 0+// ticker termination process restore ticker parent to desktop

sendmessage htaskbarwnd, wm_redraw, 0, byval 0+// stop painting the task bar

// left unchanged ==> iconsremove;// remove the icons

sendmessage htaskbarwnd, wm_redraw, 1, byval 0+// start painting in task bar and force a repaint
redrawwindow htaskbarwnd, byval 0+, 0&, rdw_invalidate or rdw_allchildren or rdw_updatenow

unload frmticker
end;






public function trayiconrows() as long//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
cr : rect;clkr : rect;
icnsize : longint;

getclientrect htraywnd, cr


if  iswindowvisible(hclockwnd)  then
begin

getwindowrect hclockwnd, clkr// get clock rect

mapwindowpoints 0+, htraywnd, clkr, 2// map clock rect to tray coordinates

if  clkr.top <> 0  then
begin
clkr.top := 0;// ignore clock size if it isn't at the top
clkr.bottom := 0;
end; //Main if block
end; //Main if block


icnsize := getsystemmetrics(sm_cycaption) - 3;// get the icon height.

trayiconrows := (cr.bottom - (clkr.bottom - clkr.top)) \ icnsize;// calculate rows
end;





procedure gettraysize(byref : as;byref : as);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
cr : rect;
getclientrect htraywnd, cr

width := cr.right;
height := cr.bottom;
end;





public function gettaskbaredge() as shappbar_edges//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
abd : appbardata;
with abd
.cbsize := lenb(abd);
.hwnd := htaskbarwnd;
end with
shappbarmessage abm_gettaskbarpos, abd

gettaskbaredge := abd.uedge;
// left unchanged ==> exit;

errint:
apperrorlog err, 'module ticker | gettaskbaredge'
end;






procedure tickermessage(opcode : string);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
ayat : string;
frmhost.timer2.enabled := true;
blnticker := true;

frmticker.text opcode
// left unchanged ==> exit;

errint:
apperrorlog err, 'module command | tickmsg'
end;

end.

