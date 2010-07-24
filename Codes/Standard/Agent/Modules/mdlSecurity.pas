unit mdlSecurity;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type

function seccheckpassword(password : variant):long;
begin

strpassword : string;
strpassword := trim(get('genadminpass'));// 1 = granted

if  strpassword = trim(password)  then seccheckpassword = 1;
end;


procedure secadjusttoken();
begin
const token_adjust_privileges = $h20
const token_query = $h8
const se_privilege_enabled = $h2
hdlprocesshandle : longint;
hdltokenhandle : longint;
tmpluid : luid;
tkp : token_privileges;
tkpnewbutignored : token_privileges;
lbufferneeded : longint;

hdlprocesshandle := getcurrentprocess();
openprocesstoken hdlprocesshandle, (token_adjust_privileges or token_query), hdltokenhandle


lookupprivilegevalue '', 'seshutdownprivilege', tmpluid// get the luid for shutdown privilege.

tkp.privilegecount := 1    ;// one privilege to set
tkp.theluid := tmpluid;
tkp.attributes := se_privilege_enabled;


adjusttokenprivileges hdltokenhandle, false, tkp, len(tkpnewbutignored), tkpnewbutignored, lbufferneeded// enable the shutdown privilege in the access token of this process.
end;


procedure deskwallprotect();
begin
// try-except block should come here!!! => on error goto errint
s_curwpaper : string;s_backwpaper : string;
l_flag : longint;lret : longint;
s_backwpaper := 'c:\windows\winwall.dat';

if  get('syssecwallpaper', 1) = 1  then
begin
if  fileexist(s_backwpaper) = false  then
begin
s_curwpaper := getwallpaper;
l_flag := get('persist.wpaperf', 0);

if  s_curwpaper = ''  then
begin
if  l_flag = 0  then
begin
save 'persist.wpaperf', 2
elseif l_flag = 1 then
s_backwpaper := '';
end; //Main if block
else 
if  l_flag = 2  then
begin
s_backwpaper := '';
elseif l_flag = 0 then
filecopy s_curwpaper, s_backwpaper
save 'persist.wpaperf', 1
end; //Main if block
end; //Main if block
end; //Main if block
lret := systemparametersinfo(spi_deskwallpaper, 0+, s_backwpaper, 0);
else 
if  fileexist(s_backwpaper) = true  then
begin
save 'persist.wpaperf', 0
kill s_backwpaper
end; //Main if block
end; //Main if block
// left unchanged ==> exit;

errint:
apperrorlog err, 'deskwallprotect'
end;


procedure deskiconprotect();
begin
if  get('syssecdesktop', 0) = 1  then
begin
dirdisable 'c:\windows\desktop'
else 
// left unchanged ==> direnable;
end; //Main if block
end;

end.

