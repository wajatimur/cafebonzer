unit MdlAgent;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type


procedure agentusage(subcommand : string);
begin
subcommand := format(subcommand, '#0.00');
if  blnticker = false  then frmticker.text '  rm ' + subcommand;
end;


procedure agentlogin(subcommand : string);
begin
if  subcommand = '0'  then
begin
save 'login', false
call sysshelllock(1)
else 
save 'login', true
call sysshelllock(0)
end; //Main if block
end;


procedure agentcertified(subcommand : string);
begin
if  mid(subcommand, 1, 1) = 1  then
begin
call netsessionstart
call agentinfostatus
else 
// left unchanged ==> netclose;
// left unchanged ==> netconnect;
end; //Main if block
end;


procedure agentinfostatus();
begin




netsend '040010' + cmdsubput('lock', lngstatuslock)//   unlock  = 0
end;


procedure agentmsgreceive(opcode : string);
begin
if  lngstatuslock = 0  then
begin
frmmessaging.lblreceive.caption := opcode;
// left unchanged ==> frmmessaging.show;
end; //Main if block
end;


end.

