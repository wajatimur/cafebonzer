unit MdlCommunication;

interface
  uses
    SysUtils, ComObj;
    
  var
    BlnConnected: Boolean;

  procedure NetConnect;
  procedure NetSessionStart;
  procedure NetDisconnect;
  procedure NetSend(Data: String);
  procedure NetPipeSend(Data: String);



implementation
  uses
    Main, MdlCommands, MdlSystem, MdlPipes, MdlGlobal;


  procedure NetConnect;
  begin
    FrmMain.HostConTimer.Enabled := True;
  end;


  procedure NetSessionStart;
  begin
    NetSend('020040'+SysNetGetMac);
    NetSend('020030'+SysDevPrintersGet);
    NetSend('040020'+CmdKeySet('DB','NetDefaultPass'));

    BlnConnected := True;
    FrmMain.HostPingTimer.Enabled := True;
   { Notify Status }
    NetPipeSend('STATUSONLINE');
    FrmMain.HostMainTray.IconIndex := 1;
  end;


  procedure NetDisconnect;
  begin
    BlnConnected := False;
    FrmMain.HostPingTimer.Enabled := False;
    FrmMain.HostSocket.Disconnect;
   { Notify Status }
    NetPipeSend('STATUSOFFLINE');
    FrmMain.HostMainTray.IconIndex := 0;
  end;


  procedure NetSend(Data: String);
  var
    StrData: String;
  begin
    if Trim(Data)='' then Exit;
    try
      if FrmMain.HostSocket.IsWritable = True then begin
        StrData := CStrCmdSep+Data;
        FrmMain.HostSocket.SendLen := Length(StrData);
        FrmMain.HostSocket.SendData := StrData;
      end;
    except
      on E: EOleException do begin
        if (E.ErrorCode = 24054) and (E.ErrorCode = 24022) then begin
          NetDisconnect;
          NetConnect;
        end;
      end;
    end;
  end;


  procedure NetPipeSend(Data: String);
  begin
    with TPipeClient.Create(CStrPipeHostName) do
    try
      SendString(Data);
    finally
      Free;
    end;
  end;
end.
