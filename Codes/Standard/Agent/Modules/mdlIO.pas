unit mdlIO;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type

fcolprotect : new;
s_curprotectdir : string;




function fileexist(byval : as):boolean;
begin
fileexist := iif(dir$(pathname) = '', false, true);//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
end;





procedure clearrbin();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
shemptyrecyclebin frmhost.hwnd, vbnullstring, sherb_noconfirmation + sherb_nosound
end;





procedure clearhistory();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
url : new;
// left unchanged ==> url.clearhistory;
end;





procedure clearrecentdocs();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
shaddtorecentdocs 2, vbnullstring
end;





procedure deltree(pathstr : variant);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
pathret : string;

if  right(pathstr, 1) <> '\'  then pathstr = pathstr + '\';
pathret := dir(pathstr, vbdirectory);

do until pathret = ''
// left unchanged ==> doevents;
if  getattr(pathstr + pathret) = vbdirectory  then
begin
if  pathret <> '.' and pathret <> '..'  then
begin
deltree pathstr + pathret
rmdir pathstr + pathret
pathret := dir(pathstr, vbdirectory);
end; //Main if block
end; //Main if block
pathret := dir;
end; //End for For or Do => Next/Loop
filewipe pathstr
end;





procedure filewipe(pathstr : variant);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
shf : shfileopstruct;pathret : string;

if  right(pathstr, 1) <> '\'  then pathstr = pathstr + '\';
shf.hwnd := frmhost.hwnd;
shf.wfunc := fo_delete;
shf.fflags := fof_silent + fof_noconfirmation + fof_noerrorui;
pathret := dir(pathstr, vbnormal + vbarchive + vbhidden + vbsystem);

do until pathret = ''
// left unchanged ==> doevents;
shf.pfrom := pathstr + pathret & chr$(0) & chr$(0);
shfileoperation shf
pathret := dir;
end; //End for For or Do => Next/Loop
end;


function dlgfileopen(stitle : string;sinitialdir : string;hwndowner : longint;sfilter : string;optional : as;optional : as):false;
begin
tofn : openfilename;

tofn.lstructsize := len(tofn);
tofn.hinstance := app.hinstance;
tofn.hwndowner := hwndowner;
tofn.flags := lflags;
tofn.lpstrtitle := stitle;
tofn.lpstrinitialdir := sinitialdir;
tofn.lpstrfilter := sfilter;

tofn.lpstrfile := space$(256);
tofn.nmaxfile := 256;

tofn.lpstrfiletitle := space$(256);
tofn.nmaxfiletitle := 256;

if  getopenfilename(tofn)  then
begin
if  btitleonly = true  then
begin
dlgfileopen := trim$(tofn.lpstrfiletitle);

else 
dlgfileopen := trim$(tofn.lpstrfile);
end; //Main if block

dlgfileopen := left(dlgfileopen, len(dlgfileopen) - 1);//remove null terminated, feel ok now haha
end; //Main if block
end;






procedure sysdevdiskclean(opcode : string);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
o_getpath : new;
op : string;









op := opcode;//  1 = clean ok

if  instr(1, op, '0') or instr(1, op, '1')  then
begin
call deltree(o_getpath.temp)
end; //Main if block
if  instr(1, op, '0') or instr(1, op, '2')  then
begin
call clearrbin
end; //Main if block
if  instr(1, op, '0') or instr(1, op, '3')  then
begin
call clearhistory
end; //Main if block
if  instr(1, op, '0') or instr(1, op, '4')  then
begin
call clearrecentdocs
end; //Main if block

netsend '040010' + cmdsubput('clean', 1)
 o_getpath = nil;
// left unchanged ==> exit;

errint:
apperrorlog err, 'module command | cleandisk'
end;








procedure dirdisable(spath : string);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error resume next
retpath : string;lfilenum : longint;

if  right(spath, 1) <> '\'  then spath = spath + '\';
if  s_curprotectdir = spath  then exit;
s_curprotectdir := spath;
retpath := dir(spath, vbnormal + vbarchive + vbhidden + vbsystem);

do until retpath = ''
// left unchanged ==> doevents;
lfilenum := freefile;
fcolprotect.add lfilenum
open spath + retpath for random lock write as #lfilenum
retpath := dir;
end; //End for For or Do => Next/Loop
end;

procedure direnable();
begin
ffile : variant;
s_curprotectdir := '';
for each ffile in fcolprotect
close #ffile
end; //End for For or Do => Next/Loop
end;



end.

