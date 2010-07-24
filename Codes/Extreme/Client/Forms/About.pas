unit About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, VrControls, VrScrollText, JvGIF, ExtCtrls;

type
  TFrmAbout = class(TForm)
    AbtImage: TImage;
    AbtScroll: TVrScrollText;
    procedure AbtImageClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure CreateParams(var Params: TCreateParams);override;
  end;

var
  FrmAbout: TFrmAbout;

implementation

{$R *.dfm}
procedure TFrmABout.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.Style := (Params.Style or WS_POPUP) and not WS_CAPTION;
end;

procedure TFrmAbout.AbtImageClick(Sender: TObject);
begin
  FrmAbout.Close;
end;

end.
