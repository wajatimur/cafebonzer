unit Main;

interface
  uses
    Windows, Messages, SysUtils, StrUtils, Variants, Classes, Graphics, Controls, Forms,
    MdlPipes, Dialogs,Registry, DdeMan, ExtCtrls, StdCtrls, JvGIF, cxGraphics, dxCore,
    dxButtons, ImgList, VrControls, VrScrollText, JvWaitingGradient,
    dxStatusBar, cxPC, cxControls, LMDCustomComponent, LMDOneInstance,
  Buttons, JvComponent, JvTransBtn;

  type
    TFrmMain = class(TForm)
      MainPanel: TPanel;
      ImgLogoWord: TImage;
      LblCompany: TLabel;
      ImgLogoSymbol: TImage;
      MainBevel: TBevel;
      MainPage: TcxPageControl;
      PageLogin: TcxTabSheet;
      PageMenu: TcxTabSheet;
      EdtPassword: TLabeledEdit;
      ImgKey: TImage;
      MainStatus: TdxStatusBar;
      MainStatusContainer1: TdxStatusBarContainerControl;
      MainScroll: TVrScrollText;
      MainImgList: TImageList;
      MainWaiting: TJvWaitingGradient;
      HostSingle: TLMDOneInstance;
    BtnExit: TJvTransparentButton;
    BtnUnlock: TJvTransparentButton;
    BtnShutdown: TJvTransparentButton;
    BtnRestart: TJvTransparentButton;
    BtnSetting: TJvTransparentButton;
      procedure EdtPasswordKeyPress(Sender: TObject; var Key: Char);
      procedure BtnActionClick(Sender: TObject);
      procedure FormCreate(Sender: TObject);
      procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    private
      CPServer: TPipeServer;
      procedure IPCProcessMsg(Sender:TObject;var ReceiveData:String);
      procedure SetIdleAnim;
    public
      procedure CreateParams(var Params: TCreateParams); override;
    end;

  var
    FrmMain: TFrmMain;
    BlnInternalClose: Boolean;
    HandleHost: THandle;


implementation
  {$R *.dfm}
  uses
    MdlGlobal, MdlDesktop, MdlSecurity;


  procedure TFrmMain.FormCreate(Sender: TObject);
  begin
    CPServer := TPipeServer.CreatePipeServer(CStrPipeServerName,True);
    CpServer.OnReceive := FrmMain.IPCProcessMsg;
    MainScroll.Caption := SysSettingGet('TickMsgWelcome');
  end;


  procedure TFrmMain.CreateParams(var Params: TCreateParams);
  begin
    inherited CreateParams(Params);
    Params.Style := (Params.Style or WS_POPUP) and not WS_CAPTION;
  end;


  procedure TFrmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  begin
    CanClose := BlnInternalClose;
  end;


  procedure TFrmMain.EdtPasswordKeyPress(Sender: TObject; var Key: Char);
  var
    StrPassword: String;
  begin
    if Key = #13 then begin
      StrPassword := SysSettingGet('GenAdminPass');
      if EdtPassword.Text = StrPassword then MainPage.ActivePage := PageMenu;
    end;
  end;


  procedure TFrmMain.BtnActionClick(Sender: TObject);
  begin
    if Sender = BtnUnlock then NetPipeSend('ACTIONUNLOCK');
    if Sender = BtnShutdown then SysShutdown(EWX_SHUTDOWN,False);
    if Sender = BtnRestart then SysShutdown(EWX_REBOOT,False);
    if Sender = BtnExit then NetPipeSend('ACTIONTERMINATE');

    if (Sender = BtnShutdown) or (Sender = BtnRestart) then SetIdleAnim
    else MainPage.ActivePage := PageLogin;
  end;


  procedure TFrmMain.IPCProcessMsg(Sender:TObject;var ReceiveData:String);
  var
    CTextPanel: TdxStatusBarTextPanelStyle;
  begin
    if ReceiveData = 'STATUSONLINE' then begin
      MainStatus.Panels[0].Text := 'Online';
      CTextPanel := TdxStatusBarTextPanelStyle(MainStatus.Panels[0].PanelStyle);
      CTextPanel.ImageIndex := 1;
    end;
    if ReceiveData = 'STATUSOFFLINE' then begin
      MainStatus.Panels[0].Text := 'Offline';
      CTextPanel := TdxStatusBarTextPanelStyle(MainStatus.Panels[0].PanelStyle);
      CTextPanel.ImageIndex := 0;
    end;
    if ReceiveData = 'STATUSWORKING' then begin
      MainScroll.Active := False;
      MainWaiting.BringToFront;
      MainWaiting.Enabled := True;
    end;
    if ReceiveData = 'STATUSWORKINGNO' then begin
      MainScroll.Active := True;
      MainWaiting.SendToBack;
      MainWaiting.Enabled := False;
    end;
  end;


  procedure TFrmMain.SetIdleAnim;
  var
    IntIdxA: Integer;
  begin
    for IntIdxA := 0 to PageMenu.ControlCount-1 do
      if PageMenu.Controls[IntIdxA].Tag = 1 then
        PageMenu.Controls[IntIdxA].Enabled := False;

    MainScroll.Active := False;
    MainWaiting.BringToFront;
    MainWaiting.Enabled := True;
  end;
end.
