object MergeBomForm: TMergeBomForm
  Left = 0
  Top = 0
  Caption = 'MergeBomForm'
  ClientHeight = 549
  ClientWidth = 829
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object bClose: TButton
    AlignWithMargins = True
    Left = 16
    Top = 16
    Width = 88
    Height = 32
    Caption = 'Close'
    TabOrder = 0
    OnClick = bCloseClick
  end
  object brun: TButton
    AlignWithMargins = True
    Left = 15
    Top = 56
    Width = 89
    Height = 32
    Caption = 'Run'
    TabOrder = 1
    OnClick = brunClick
  end
  object log: TMemo
    Left = 120
    Top = 16
    Width = 696
    Height = 520
    Lines.Strings = (
      'log')
    ScrollBars = ssVertical
    TabOrder = 2
  end
end
