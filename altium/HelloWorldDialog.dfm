object HelloWorldForm: THelloWorldForm
  Left = 5
  Top = 4
  Caption = 'Hello World!'
  ClientHeight = 399
  ClientWidth = 857
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object bDisplay: TButton
    Left = 8
    Top = 24
    Width = 75
    Height = 25
    Caption = 'Display'
    TabOrder = 0
    OnClick = bDisplayClick
  end
  object bClose: TButton
    Left = 8
    Top = 56
    Width = 75
    Height = 25
    Caption = 'Close'
    TabOrder = 1
    OnClick = bCloseClick
  end
  object log: TMemo
    Left = 88
    Top = 24
    Width = 752
    Height = 360
    Lines.Strings = (
      'log')
    ScrollBars = ssVertical
    TabOrder = 2
  end
end
