program CbClientEx;

uses
  Forms,
  Main in 'Forms\Main.pas' {FrmMain},
  Config in 'Forms\Config.pas' {FrmOption},
  Ticker in 'Forms\Ticker.pas' {FrmTicker},
  MdlSettings in 'Units\MdlSettings.pas',
  MdlCommunication in 'Units\MdlCommunication.pas',
  MdlCommands in 'Units\MdlCommands.pas',
  MdlAgent in 'Units\MdlAgent.pas',
  MdlSystem in 'Units\MdlSystem.pas',
  MdlGlobal in 'Units\MdlGlobal.pas',
  MdlSecurity in 'Units\MdlSecurity.pas',
  MdlUtils in 'Units\MdlUtils.pas',
  MdlTicker in 'Units\MdlTicker.pas',
  MdlDesktop in 'Units\MdlDesktop.pas',
  MdlPipes in 'Units\MdlPipes.pas',
  About in 'Forms\About.pas' {FrmAbout};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'CbClientEx';
  Application.CreateForm(TFrmMain, FrmMain);
  Application.CreateForm(TFrmOption, FrmOption);
  Application.CreateForm(TFrmTicker, FrmTicker);
  Application.CreateForm(TFrmAbout, FrmAbout);
  Application.ShowMainForm := False;
  FrmMain.Visible := False;
  MainStart;
  Application.Run;
end.
