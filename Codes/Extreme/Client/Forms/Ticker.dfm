object FrmTicker: TFrmTicker
  Left = 578
  Top = 358
  BorderStyle = bsNone
  Caption = 'FrmTicker'
  ClientHeight = 33
  ClientWidth = 192
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Verdana'
  Font.Style = []
  OldCreateOrder = False
  PopupMenu = FrmMain.HostMainPopUp
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object MainScroll: TVrScrollText
    Left = 0
    Top = 0
    Width = 150
    Height = 20
    Threaded = True
    EdgeWidth = 5
    Active = True
    Caption = 'Testing'
  end
  object TmrCheck: TTimer
    Enabled = False
    Interval = 40
    OnTimer = TmrCheckTimer
    Left = 163
  end
end
