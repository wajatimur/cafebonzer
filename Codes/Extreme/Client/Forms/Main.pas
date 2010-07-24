unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  MdlPipes, Dialogs, Sockets, cxLookAndFeelPainters, WmiAbstract, WmiComponent,
  WmiSystemEvents, ExtCtrls, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, LMDGlobalHotKey, LMDCustomComponent, LMDOneInstance, Menus,
  LMDPopupMenu, ImgList, CoolTrayIcon, TextTrayIcon, StdCtrls,
  JvExStdCtrls, JvButton, JvCtrls, JvGroupHeader, ComCtrls, cxPC,
  cxControls, OleCtrls, SocketWrenchCtrl_TLB, JvExControls, JvComponent,
  JvArrowButton, AdvPanel;

type
  TFrmMain = class(TForm)
    VisPanels: TAdvPanel;
    HostMainTray: TTextTrayIcon;
    HostImgList16: TImageList;
    VisPages: TcxPageControl;
    PageVariable: TcxTabSheet;
    VisBevel: TBevel;
    PvListView: TListView;
    HostMainPopUp: TLMDPopupMenu;
    PnuConfig: TMenuItem;
    PnuAbout: TMenuItem;
    HostSingleInstance: TLMDOneInstance;
    HostGlobalHotKey: TLMDGlobalHotKey;
    BtnMenu: TJvArrowButton;
    PnuSystem: TMenuItem;
    PnuSysShutdown: TMenuItem;
    PnuSysRestart: TMenuItem;
    PnuSysLogOff: TMenuItem;
    PnuSysRelogin: TMenuItem;
    N1: TMenuItem;
    PnuExit: TMenuItem;
    PnuSecurity: TMenuItem;
    Enable1: TMenuItem;
    Settings1: TMenuItem;
    Socket: TIdTCPClient;
    N2: TMenuItem;
    PnuActivity: TMenuItem;
    PnuActivitySend: TMenuItem;
    PnuActivityChat: TMenuItem;
    PnuEndSession: TMenuItem;
    PnuTools: TMenuItem;
    PnuToolsControl: TMenuItem;
    PnuToolsFolders: TMenuItem;
    PnuMaintenance: TMenuItem;
    ClearRecycleBin1: TMenuItem;
    ClearInternetCache1: TMenuItem;
    ClearRecents1: TMenuItem;
    ClearTempFolder1: TMenuItem;
    Desktop1: TMenuItem;
    SavePosition1: TMenuItem;
    ResetIcon1: TMenuItem;
    IconPositionConfiguration1: TMenuItem;
    HostConTimer: TTimer;
    HostSocket: TSocket;
    HostPingTimer: TTimer;
    HostSystemEvent: TWmiSystemEvents;
    PageCommunication: TcxTabSheet;
    IPCSendTxt: TLabeledEdit;
    GhdIPC: TJvGroupHeader;
    IPCSendBtn: TJvImgBtn;
    HostTimerClose: TTimer;
    PnuSysLock: TMenuItem;
    BtnDebug: TJvArrowButton;
    PnuClearHistory: TMenuItem;
    ClearAllTrack1: TMenuItem;
    N3: TMenuItem;
    PnuProgCmd: TMenuItem;
    PnuProgReg: TMenuItem;
    N4: TMenuItem;
    procedure GhKeyAction(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PnuAction(Sender: TObject);
    procedure HostAppTimer(Sender: TObject);
    procedure BtnTestClick(Sender: TObject);
    procedure HostSocketConnect(Sender: TObject);
    procedure HostSocketRead(ASender: TObject; var DataLength,IsUrgent: Smallint);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure IPCSendBtnClick(Sender: TObject);
    procedure HostSocketDisconnect(Sender: TObject);
  private
    procedure IPCProcessMsg(Sender: TObject;var ReceiveData:String);
  public
    CPServer: TPipeServer;
  end;

var
  FrmMain: TFrmMain;

  

implementation
  {$R *.dfm}
  uses
    MdlSettings, MdlCommands, MdlCommunication, MdlSystem, MdlGlobal,
    MdlUtils, MdlDesktop, About, Config;

    
  procedure TFrmMain.FormCreate(Sender: TObject);
  begin
   { IPC:Pipe - Start }
    CPServer := TPipeServer.CreatePipeServer(CStrPipeServerName,True);
    CPServer.OnReceive := FrmMain.IPCProcessMsg;
  end;


  procedure TFrmMain.GhKeyAction(Sender: TObject; var Key: Word;
    Shift: TShiftState);
  begin
    FrmMain.Show;
  end;


  procedure TFrmMain.PnuAction(Sender: TObject);
  begin
    if Sender = PnuAbout then FrmAbout.Show;
    if Sender = PnuExit then AppClose;
    if Sender = PnuConfig then FrmOption.Show;
  end;


  procedure TFrmMain.HostAppTimer(Sender: TObject);
  var
    StrPort: String;
  begin
    if Sender = HostConTImer then begin
      StrPort := SettingGet('NetServerPort');
      HostSocket.RemotePort := StrToInt(StrPort);
      HostSocket.HostAddress := SettingGet('NetServerIp');
      HostSocket.Connect;
    end;
    if Sender = HostPingTimer then  NetSend('010010');  { PING }
    if Sender = HostTimerClose then AppClose;           { Time To Kill }
  end;


  procedure TFrmMain.HostSocketConnect(Sender: TObject);
  begin
    HostConTimer.Enabled := False;                                                                { Disable Connecter Timer }
    NetSend('010030'+CmdKeySet('NAME',SysInfoGetName)+CmdKeySet('AGENTVERSION',StrAppVersion));   { Send Info to Certified }
  end;


  procedure TFrmMain.HostSocketDisconnect(Sender: TObject);
  begin
    NetDisconnect;
    NetConnect;
  end;


  procedure TFrmMain.HostSocketRead(ASender: TObject; var DataLength,
    IsUrgent: Smallint);
  var
    DataRcv: WideString;
  begin
    HostSocket.Read(DataRcv,DataLength);
    CmdParse(DataRcv);
  end;


  procedure TFrmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  begin
    CanClose := False;
    FrmMain.Hide;
  end;


  procedure TFrmMain.IPCProcessMsg(Sender: TObject;var ReceiveData:String);
  begin
    if ReceiveData = 'ACTIONUNLOCK' then begin
      SysShellLock(0);
    end;
    if ReceiveData = 'ACTIONSHUTDOWN' then begin

    end;
    if ReceiveData = 'ACTIONREBOOT' then begin

    end;
    if ReceiveData = 'ACTIONTERMINATE' then begin
      DesktopSwitch(True);            { Switch Back to Default Desktop }
      HostTimerClose.Enabled := True; { Activate App. Close Timer }
    end;
  end;

  
  procedure TFrmMain.IPCSendBtnClick(Sender: TObject);
  begin
    NetPipeSend(IPCSendTxt.Text);
  end;


  procedure TFrmMain.BtnTestClick(Sender: TObject);
    //TestVar: String;
  begin
    //TestVar := CStrCmdSep+'101000'+CmdKeySet('Test1','MyValue')+CmdKeySet('Test2','SecondValue');
    //StartProcess(ExtractFilePath(Application.ExeName)+'CbDaemon.exe','Default');
  end;
end.
