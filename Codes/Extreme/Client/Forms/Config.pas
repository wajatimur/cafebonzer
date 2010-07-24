unit Config;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Registry, JvGIF, ExtCtrls, ImgList, dxsbar, JvComCtrls,
  JvComponent, JvGroupHeader, cxPC, cxControls, StdCtrls, ComCtrls, JvEdit,
  dxExEdtr, dxInspRw, dxInspct, dxCntner, LMDCustomListBox,
  LMDCustomImageListBox, LMDCustomColorListBox, LMDColorListBox,
  cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxColorComboBox, cxStyles, cxGraphics, cxCheckBox, cxSpinEdit, cxVGrid,
  cxInplaceContainer, JvExControls, JvExStdCtrls, JvValidateEdit, dxSkinsCore,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinDarkSide,
  dxSkinGlassOceans, dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky,
  dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Green, dxSkinOffice2007Pink,
  dxSkinOffice2007Silver, dxSkinPumpkin, dxSkinSilver, dxSkinSpringTime,
  dxSkinStardust, dxSkinSummer2008, dxSkinsDefaultPainters, dxSkinValentine,
  dxSkinXmas2008Blue, dxSkinscxPCPainter;

type
  TFrmOption = class(TForm)
    EdtServerIP: TJvIpAddress;
    IntPanelBtm: TPanel;
    OptPages: TcxPageControl;
    PgsGeneral: TcxTabSheet;
    PgsNetwork: TcxTabSheet;
    IntBevelMid: TBevel;
    NavSelector: TdxSideBar;
    OptNavStores: TdxSideBarStore;
    OptImgs32: TImageList;
    ItmAgnGen: TdxStoredSideItem;
    ItmAgnNet: TdxStoredSideItem;
    ItmAgnSec: TdxStoredSideItem;
    ItmAgnMon: TdxStoredSideItem;
    GhdStartup: TJvGroupHeader;
    IntImgBanner: TImage;
    CbxStartupUser: TComboBoxEx;
    LblGeneral0: TLabel;
    OptImgs16: TImageList;
    LblGeneral1: TLabel;
    CbxStartupOn: TComboBoxEx;
    GhdAuth: TJvGroupHeader;
    LbdAuthPass1: TLabeledEdit;
    LbdAuthPass2: TLabeledEdit;
    ChkAuthRetPass: TCheckBox;
    ChkAuthReqServ: TCheckBox;
    GhdServer: TJvGroupHeader;
    LblNetwork1: TLabel;
    LblNetwork2: TLabel;
    PgsSecurity: TcxTabSheet;
    GhdPolicy: TJvGroupHeader;
    ItmAgnApr: TdxStoredSideItem;
    PgsAppearence: TcxTabSheet;
    GhdTicker: TJvGroupHeader;
    IspTrayTicker: TcxVerticalGrid;
    OptIspGrpTicker: TcxCategoryRow;
    IspItmEnable: TcxEditorRow;
    IspItmDefaultText: TcxEditorRow;
    IspItmTickerSize: TcxEditorRow;
    IspItmTextColor: TcxEditorRow;
    IspItmBackColor: TcxEditorRow;
    GhdLock: TJvGroupHeader;
    EdtServerPort: TJvValidateEdit;
    OptStyleGlobal: TcxStyleRepository;
    StyHeader: TcxStyle;
    procedure ItmAgnClick(Sender: TObject; Item: TdxSideBarItem);
    procedure FormCreate(Sender: TObject);
  private
    { Private }
  public
    procedure SetFirstTime;
  end;

var
  BlnFirstTime: boolean=False;
  FrmOption: TFrmOption;

implementation
{$R *.dfm}

procedure TFrmOption.FormCreate(Sender: TObject);
begin
  CbxStartupUser.ItemIndex := 0;
  CbxStartupOn.ItemIndex := 1;
end;

procedure TFrmOption.ItmAgnClick(Sender: TObject; Item: TdxSideBarItem);
begin
  if Item.StoredItem = ItmAgnGen then OptPages.ActivePage := PgsGeneral;
  if Item.StoredItem = ItmAgnNet then OptPages.ActivePage := PgsNetwork;
  if Item.StoredItem = ItmAgnApr then OptPages.ActivePage := PgsAppearence;
  if Item.StoredItem = ItmAgnSec then OptPages.ActivePage := PgsSecurity;
end;

procedure TFrmOption.SetFirstTime;
begin
  BlnFirstTime := True;
end;



end.
