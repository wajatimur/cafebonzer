unit MdlCommands;

interface
  uses
    Classes, SysUtils, StrUtils, JclStrings;

  const
    CStrCmdSep = #20;
    CStrKeySep1 = #210;
    CStrKeySep2 = #220;

  procedure CmdParse(Data: String);
  function CmdCount(Data: String):integer;
  function CmdKeySet(Key:String; Value:String):String;
  function CmdKeyGet(KeyData:String; Key:String):String;
  function CmdKeyCount(KeyData:String):integer;



implementation
  uses
    MdlCommunication, MdlAgent, MdlSystem, MdlGlobal;


  procedure CmdParse(Data: String);
  var
    CSl: TStringList;
    StrKeyData, StrCmd, StrSubCmd: String;
    IntCmdCount, IntIdx, IntTmp: Integer;
  begin
    IntCmdCount := CmdCount(Data);
    if IntCmdCount = 0 then Exit;
    CSl := TStringList.Create;

    try
      CSl.Delimiter := CStrCmdSep;
      CSl.DelimitedText := Data;
      for IntIdx := 0 to IntCmdCount-1 do begin
        StrKeyData := CSl.Strings[IntIdx];
        StrCmd := LeftStr(StrKeyData,2);
        StrSubCmd := MidStr(StrKeyData,3,4);

        if StrCmd = '01' then begin
          if StrSubCmd = '0010' then NetSend('010020');
          if StrSubCmd = '0020' then {NetPingReset};
          if StrSubCmd = '0030' then AgentCertified(StrKeyData);;
        end else
        if StrCmd = '02' then begin
          if StrSubCmd = '0010' then SysCommand('0010',StrKeyData);
          if StrSubCmd = '0020' then SysCommand('0020',StrKeyData);
          if StrSubCmd = '0030' then {SysDevDiskClean};
          if StrSubCmd = '0040' then AgentLogin(StrKeyData);
          if StrSubCmd = '0050' then AgentUsage(StrKeyData);
        end else
        if StrCmd = '03' then begin
          if StrSubCmd = '0010' then {Message Receive};
          if StrSubCmd = '0020' then {Ticker Receive};
        end else
        if StrCMd = '04' then begin
          if StrSubCmd = '0010' then AgentInfoStatus;
          if StrSubCmd = '0020' then {AgentConfiguration};
          if StrSubCmd = '0100' then AppClose;
        end;
      end;
    finally
      CSl.Free;
    end;
  end;


  function CmdCount(Data: String):integer;
  begin
    Result := StrStrCount(Data,CStrCmdSep);
  end;


  function CmdKeySet(Key:String; Value:String):String;
  begin
    Result := CStrKeySep1+Key+CStrKeySep2+Value;
  end;


  function CmdKeyGet(KeyData:String; Key:String):String;
  var
    SL: TStringList;
  begin
    SL := TStringList.Create;
    try
      SL.Delimiter := CStrKeySep1;
      SL.NameValueSeparator := CStrKeySep2;
      SL.DelimitedText := KeyData;
      Result := SL.Values[Key];
    finally
      SL.Free;
    end;
  end;


  function CmdKeyCount(KeyData:String):integer;
  begin
    Result := StrStrCount(KeyData,CStrKeySep1);
  end;
end.
