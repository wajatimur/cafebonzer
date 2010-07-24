unit mdlSetting;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type









// left unchanged ==> public const cstrtickmsgwelcome = '[ cafebonzer system - welcome ]';//
// left unchanged ==> public const cstrtingpath = 'software\microsoft\windows\currentversion\security';






procedure tingenv();
begin//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
call loadlanguage


strtickmsgwelcome := get('tickmsgwelcome', cstrtickmsgwelcome);//[ global variable ]
strappversion := 'cafebonzer v' + app.major & '.' & app.minor;
strappbuild := app.minor + '.' & app.revision;
lngerrortype := 2;
lngenvplatformid := sysinfogetos(getplatformid);


strcmdsep := chr(2) + chr(20);//[ command seperator ]
strcmdsubsep1 := chr(210);
strcmdsubsep2 := chr(220);
end;





procedure tingfirstload();
begin//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
cinput : new;strpassword : string;

if  command = '/up'  then
begin
strpassword := cinput.getinput('enter password');
if  seccheckpassword(strpassword) = 1  then
begin
frmmain.show vbmodal
end; //Main if block
close; //End Program command!
end; //Main if block

if  get('appfirsttime') = ''  then
begin
blnappfirsttime := true;
frmmain.show vbmodal
if  blnapptoclose = true  then end;
end; //Main if block
end;





procedure tingprotect();
begin//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

if  lngenvplatformid = 1  then
begin
registerserviceprocess getcurrentprocessid, 1// + platform dependent settings ------------------------------------------
if  get('sysdiscad', 1) = 1  then sysdisctlaltdel true;
call deskwallprotect
end; //Main if block


if  get('tickguidisable', 1) = 1  then traystart else tickerstart;
if  get('sysautolock', 1) = 1  then
begin
if  get('login', false) = false  then sysshelllock 1;
end; //Main if block// + general settings -----------------------------------------------------


end;//call deskiconprotect





procedure save(namating : string;nilai : string);
begin//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
// try-except block should come here!!! => on error goto errint
savestring hkey_local_machine, cstrtingpath, namating, nilai
// left unchanged ==> exit;

errint:
apperrorlog err, 'module ting | save'
end;





function get(namating : string;optional : as):variant;
begin//-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
// try-except block should come here!!! => on error goto errint

get := getstring(hkey_local_machine, cstrtingpath, namating);
if  get = ''  then get = default;
// left unchanged ==> exit;

errint:
apperrorlog err, 'module ting | get'
end;

end.

