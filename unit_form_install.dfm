object form_install: Tform_install
  Left = 0
  Top = 0
  BorderIcons = []
  Caption = 'install'
  ClientHeight = 308
  ClientWidth = 462
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object button_exit: TButton
    Left = 6
    Top = 272
    Width = 83
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Courier'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = button_exitClick
  end
  object button_next: TButton
    Left = 344
    Top = 275
    Width = 103
    Height = 25
    Caption = #1044#1072#1083#1077#1077
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Courier'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
  end
  object MemoInfo: TMemo
    Left = 14
    Top = 8
    Width = 433
    Height = 249
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Courier'
    Font.Style = []
    Lines.Strings = (
      'memo_info')
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 2
  end
  object MainMenu1: TMainMenu
    Left = 248
    Top = 264
    object N2: TMenuItem
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1086#1090#1095#1077#1090
      OnClick = N2Click
    end
    object N1: TMenuItem
      Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
      OnClick = N1Click
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #1041#1072#1079#1072' '#1076#1072#1085#1085#1099#1093' Access 2002-2003|*.mdb'
    Left = 184
    Top = 264
  end
  object ADOQuery1: TADOQuery
    DataSource = mr_service.DataSource1
    Parameters = <>
    Left = 112
    Top = 264
  end
  object ADODataSet1: TADODataSet
    DataSource = mr_service.DataSource1
    Parameters = <>
    Left = 304
    Top = 264
  end
end
