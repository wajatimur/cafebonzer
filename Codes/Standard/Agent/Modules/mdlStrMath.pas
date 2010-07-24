unit mdlStrMath;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type





function validateip(strtocheck : variant):boolean;
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
nonenumchar := 'abcdefghijklmnopqrstuvwxyz/,<>;:'''[]{}-=\|+_()*&^%$#@!~`';
validateip := true;

for d:=1 to len(strtocheck) do
begin
for i:=1 to len(nonenumchar) do
begin
if  mid(lcase(strtocheck), d, 1) = mid(lcase(nonenumchar), i, 1)  then validateip = false: exit;
end;//Current For/Next
end;//Current For/Next
end;





function getshortstr(strtoshort : string;optional : as):8;
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
if  len(strtoshort) > ideallen  then
begin
getshortstr := left(strtoshort, ideallen) + '..';
else 
getshortstr := strtoshort;
end; //Main if block
end;

function convmem2str(memlong : longint):string;
begin
bytebuffer: array[64] of long;
lret := lstrcpy(bytebuffer(0), byval memlong);
convmem2str := strconv(bytebuffer(), vbunicode);
convmem2str := left$(convmem2str, instr(convmem2str, vbnullchar) - 1);
end;

function removenull(stringwithnull : string):string;
begin
removenull := left$(stringwithnull, instr(stringwithnull, vbnullchar) - 1);
end;

end.

