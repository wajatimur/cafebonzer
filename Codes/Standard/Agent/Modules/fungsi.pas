unit fungsi;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type


function cmdcount(value : variant):long;
begin
lngidxa : longint;lngcmd : longint;
for lngidxa:=1 to len(value) do
begin
if  mid(value, lngidxa, 2) = strcmdsep  then lngcmd = lngcmd + 1;
end;//Current For/Next
cmdcount := lngcmd;
end;


function cmdsubcount(value : variant):long;
begin
lngidxa : longint;lngcmd : longint;
for lngidxa:=1 to len(value) do
begin
if  mid(value, lngidxa, 1) = strcmdsubsep1  then lngcmd = lngcmd + 1;
end;//Current For/Next
cmdsubcount := lngcmd;
end;


function cmdsubput(data : variant;value : variant):string;
begin
cmdsubput := strcmdsubsep1 + data & strcmdsubsep2 & value;
end;


function cmdsubget(value : variant;subname : variant):string;
begin
strtmp : string;strtmp2 : string;
lngidxa : longint;lngcnt : longint;

lngcnt := cmdsubcount(value);
for lngidxa:=1 to lngcnt do
begin
strtmp2 := split(value, strcmdsubsep1)(lngidxa);
strtmp := split(strtmp2, strcmdsubsep2)(1);
if  lcase$(strtmp) = lcase$(subname)  then
begin
cmdsubget := split(strtmp2, strcmdsubsep2)(2);
// left unchanged ==> exit;
end; //Main if block
end; //End for For or Do => Next/Loop
end;


procedure cmdparse(cmddata : string);
begin
lngcmdcount : longint;lngidxa : longint;strcmddata : string;
strcmdmain : string;strcmdsub : string;strcmdsubdata : string;

lngcmdcount := cmdcount(cmddata);
if  lngcmdcount = 0  then exit;

for lngidxa:=1 to lngcmdcount do
begin
strcmddata := split(cmddata, strcmdsep)(lngidxa);
strcmdmain := mid(strcmddata, 1, 2);
strcmdsub := mid(strcmddata, 3, 4);
strcmdsubdata := mid(strcmddata, 7);

if  strcmdmain = '01'  then
begin
case strcmdsub of 

is :                 Call NetSend("010020")

is : 

is :                 Call AgentCertified(StrCmdsubData)

end;

elseif strcmdmain = '02' then
case strcmdsub of 

is :                 Call SysWindowsExit(StrCmdsubData)

is :                 Call SysShellLock(StrCmdsubData)

is :                 Call SysDevDiskClean(StrCmdsubData)

is :                 Call AgentLogin(StrCmdsubData)

is :                 Call AgentUsage(StrCmdsubData)

end;

elseif strcmdmain = '03' then
case strcmdsub of 

is :                 Call AgentMsgReceive(StrCmdsubData)

is :                 Call TickerSetMessage(StrCmdsubData)

end;

elseif strcmdmain = '04' then
case strcmdsub of 

is :                 Call AgentInfoStatus

is :                 Call AppExit

end;

end; //Main if block
end; //End for For or Do => Next/Loop
end;


end.

