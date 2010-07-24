program CbDaemon;

uses
  Forms,
  Windows,
  Main in 'Forms\Main.pas' {FrmMain},
  MdlGlobal in 'Units\MdlGlobal.pas',
  MdlPipes in 'Units\MdlPipes.pas',
  MdlSecurity in 'Units\MdlSecurity.pas',
  MdlDesktop in 'Units\MdlDesktop.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'CafeBonzer Daemon';
  Application.CreateForm(TFrmMain, FrmMain);
  Application.Run;
end.
