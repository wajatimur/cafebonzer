unit MdlAgent;

interface
  uses
    SysUtils;

    procedure AgentCertified(KeyData: String);
    procedure AgentInfoStatus;
    procedure AgentUsage(KeyData: String);
    procedure AgentLogin(KeyData: String);
    procedure AgentMessage(KeyData: String);

  var
    IntStatusLock: Integer;


  
implementation
  uses
    MdlCommands, MdlCommunication, MdlGlobal, MdlTicker, Ticker,
    MdlSettings, MdlSystem;


  procedure AgentCertified(KeyData: String);
  begin
    if CmdKeyGet(KeyData,'ACTION') = '1' then begin
      NetSessionStart;
      AgentInfoStatus;
    end
    else begin
      NetDisconnect;
      NetConnect;
    end;
  end;


  procedure AgentInfoStatus;
  begin
    NetSend('040010'+CmdKeySet('LOCK',IntToStr(IntStatusLock)));
  end;


  procedure AgentUsage(KeyData: String);
  begin
    if CmdKeyGet(KeyData,'ACTION') = '0' then
      FrmTicker.MainScroll.Caption := CmdKeyGet(KeyData,'PRICEUSE')
    else
      FrmTicker.MainScroll.Caption := CmdKeyGet(KeyData,'TIMELEFT');
  end;


  procedure AgentLogin(KeyData: String);
  begin
    if CmdKeyGet(KeyData,'ACTION') = '0' then begin
      SettingSave('LOGIN','False');
      SysShellLock(1);
    end else begin
      SettingSave('LOGIN','True');
      SysShellLock(0);
    end;
  end;


  procedure AgentMessage(KeyData: String);
  begin
    //e
  end;
end.
