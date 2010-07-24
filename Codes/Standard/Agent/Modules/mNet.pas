unit mNet;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type






procedure netconnect();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
frmhost.connecter.enabled := true;
end;





procedure netsessionstart();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
netsend '020040' + cmdsubput('netmac', sysnetgetmac)
netsend '020030' + sysdevprintersget

blnconnected := true;
if  lngstatuslock = 1  then frmkey.staticon (connected);

frmhost.pinger := true;//frmticker.settext " [ connected to server ! ] "
// left unchanged ==> exit;

errint:
apperrorlog err, 'mdlnet | netsessionstart'
end;





procedure netclose();
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
// try-except block should come here!!! => on error goto errint
blnconnected := false;
frmhost.pinger := false;
// left unchanged ==> frmhost.socket.disconnect;


if  lngstatuslock = 1  then frmkey.staticon (discconnet);
// left unchanged ==> exit;//frmticker.settext ""

errint:
apperrorlog err, 'netclose | mdlnet'
end;





public function netping()//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
call netsend('010010')
end;






procedure netsend(data : variant);
begin//*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'

strdata : string;

if  trim(data) = ''  then exit;
strdata := strcmdsep + cstr(data);//on error goto errint

if  frmhost.socket.iswritable = true  then
begin
frmhost.socket.sendlen := len(strdata);
frmhost.socket.senddata := strdata;
end; //Main if block
// left unchanged ==> exit;

errint:
if  err.number = 24054 or err.number = 24022  then
begin
// left unchanged ==> netclose;
// left unchanged ==> netconnect;
else 
apperrorlog err, 'module net | netsend'
end; //Main if block
end;

end.

