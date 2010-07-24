unit MdlSettings;

interface
  uses
    Windows, Classes, Registry;

  function SettingGet(SettingName: String):String;
  function SettingSave(SettingName: String; SettingValue: String):Boolean;



implementation
  uses
    Config, MdlGlobal, MdlSystem, Main, MdlTicker;


  function SettingGet(SettingName: String):String;
  var
    CReg: TRegistry;
  begin
    CReg := TRegistry.Create(KEY_READ);
    try
      CReg.RootKey := HKEY_LOCAL_MACHINE;
      CReg.OpenKey(CStrSettingPath,True);
      Result := CReg.ReadString(SettingName);
    finally
      CReg.Free;
    end;
  end;


  function SettingSave(SettingName: String; SettingValue: String): Boolean;
  var
    CReg: TRegistry;
  begin
    CReg := TRegistry.Create(KEY_WRITE);
    try
      CReg.RootKey := HKEY_LOCAL_MACHINE;
      CReg.OpenKey(CStrSettingPath,True);
      CReg.WriteString(SettingName,SettingValue);
      Result := True;
    finally
      CReg.Free;
    end;
  end;
end.
