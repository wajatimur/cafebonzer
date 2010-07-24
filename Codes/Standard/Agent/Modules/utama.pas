unit utama;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type








lmonswitch : longint;//###################################################
blnappfirsttime : boolean;
blnapptoclose : boolean;
blnconnected : boolean;

strtickmsgwelcome : string;
strappversion : string;
strappbuild : string;
lngerrortype : longint;
lngenvplatformid : longint;

strcmdsep : string;
strcmdsubsep1 : string;
strcmdsubsep2 : string;

objpol : new;
lngstatuslock : longint;
blnticker : boolean;


// left unchanged ==> public const g_cstrshellviewwnd as string = 'progman';//// names of the shell windows we'll be looking for(windows class)
// left unchanged ==> public const g_cstrshelltaskbarwnd as string = 'shell_traywnd';






procedure main();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
if  app.previnstance = true  then end;

call tingenv
call tingfirstload
call netconnect
call tingprotect
end;






procedure apppriority(optional : as);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
pid : longint;hprocess : longint;

pid := getcurrentprocessid;
hprocess := openprocess(process_dup_handle, true, pid);
if  pnormal = false  then
begin
priorityclass hprocess, realtime_priority_class
else 
priorityclass hprocess, idle_priority_class
end; //Main if block
call closehandle(hprocess)
// left unchanged ==> exit;

errint:
apperrorlog err, 'module system | apppriority'
end;





procedure apperrorlog(errobj : errobject;procname : string);
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
// try-except block should come here!!! => on error goto errint
interrnum : integer;strerrdesc : string;strerrsource : string;
errdesc : string;

interrnum := errobj.number;
strerrsource := errobj.source;
strerrdesc := errobj.description;

case lngerrortype of 

1 :         ErrDesc = "[ " & ProcName & " | "

errdesc := errdesc + interrnum & ' | ';
errdesc := errdesc + strerrdesc & ' ]';
frmticker.text errdesc
2 :         MsgBox IntErrNum & " / " & StrErrSource & vbNewLine & StrErrDesc, vbExclamation, ProcName

end;





// left unchanged ==> exit;//close #1
errint:
showmessage err.number + ' / ' & err.source & vbnewline & err.description, vbexclamation, 'error handler'
end;





procedure appexit();
begin//=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
blnapptoclose := true;

// left unchanged ==> frmhost.socket.cleanup;
call netclose
call tickerstop
call trayremove

 objpol = nil;
for each form in forms
unload form
end; //End for For or Do => Next/Loop
end;


function proclowlevelkeyboard(byval : as;byval : as;byval : as):long;
begin
featkeystroke : boolean;
p : kbdllhookstruct;

if  (ncode = hc_action)  then
begin
if  wparam = wm_keydown or wparam = wm_syskeydown or wparam = wm_keyup or wparam = wm_syskeyup  then
begin
copymemory p, byval lparam, len(p)
featkeystroke := _;
[[p.vkcode := vk_tab) and ((p.flags and llkhf_altdown) <> 0)) or _;
[[p.vkcode := vk_escape) and ((p.flags and llkhf_altdown) <> 0)) or _;
[[p.vkcode := vk_escape) and ((getkeystate(vk_control) and $h8000) <> 0)) or _;
p.vkcode := vk_lwin or p.vkcode = vk_rwin;
end; //Main if block
end; //Main if block

if  featkeystroke  then
begin
proclowlevelkeyboard := -1;
else 
proclowlevelkeyboard := callnexthookex(0, ncode, wparam, byval lparam);
end; //Main if block
end;



end.

