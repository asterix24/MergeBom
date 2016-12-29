object MergeBomForm: TMergeBomForm
  Left = 0
  Top = 0
  Caption = 'MergeBomForm'
  ClientHeight = 496
  ClientWidth = 665
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object close: TButton
    AlignWithMargins = True
    Left = 16
    Top = 16
    Width = 88
    Height = 32
    Caption = 'Close'
    TabOrder = 0
    OnClick = closeClick
  end
  object run: TButton
    AlignWithMargins = True
    Left = 15
    Top = 56
    Width = 89
    Height = 32
    Caption = 'Run'
    TabOrder = 1
    OnClick = runClick
  end
  object showlog: TButton
    AlignWithMargins = True
    Left = 15
    Top = 99
    Width = 89
    Height = 32
    Caption = 'ShowLog'
    TabOrder = 2
    OnClick = showlogClick
  end
  object log: TMemo
    Left = 120
    Top = 16
    Width = 528
    Height = 464
    Lines.Strings = (
      'log')
    ScrollBars = ssVertical
    TabOrder = 3
  end
end
