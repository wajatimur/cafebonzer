unit Ticker;

interface

uses
  Windows, ExtCtrls, Classes, Controls, Forms, VrControls, VrScrollText;

type
  TFrmTicker = class(TForm)
    MainScroll: TVrScrollText;
    TmrCheck: TTimer;
    procedure FormResize(Sender: TObject);
    procedure TmrCheckTimer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmTicker: TFrmTicker;

implementation

uses Main, MdlTicker;
{$R *.dfm}


procedure TFrmTicker.FormResize(Sender: TObject);
begin
  MainScroll.Width := FrmTicker.Width;
  Mainscroll.Height := FrmTicker.Height;
end;

procedure TFrmTicker.TmrCheckTimer(Sender: TObject);
begin
  TickerCheck;
end;

end.
