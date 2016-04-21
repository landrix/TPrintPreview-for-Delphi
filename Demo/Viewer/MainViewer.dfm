object MainForm: TMainForm
  Left = 177
  Top = 82
  Caption = 'Print Preview Viewer'
  ClientHeight = 493
  ClientWidth = 678
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = True
  Position = poDefault
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Splitter1: TSplitter
    Left = 0
    Top = 0
    Width = 2
    Height = 493
    MinSize = 110
  end
  object Panel: TPanel
    Left = 2
    Top = 0
    Width = 676
    Height = 493
    Align = alClient
    BevelOuter = bvNone
    ParentColor = True
    TabOrder = 0
    ExplicitWidth = 681
    object PageNavigator: TTabSet
      Left = 0
      Top = 464
      Width = 676
      Height = 29
      Align = alBottom
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -19
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      OnChange = PageNavigatorChange
      ExplicitWidth = 681
    end
  end
  object OpenDialog: TOpenDialog
    DefaultExt = 'ppv'
    Filter = 'Print Preview Files (*.ppv)|*.ppv|All Files (*.*)|*.*'
    Title = 'Load Preview'
    Left = 38
    Top = 16
  end
  object MainMenu: TMainMenu
    Left = 68
    Top = 16
    object FilePopup: TMenuItem
      Caption = 'File'
      OnClick = FilePopupClick
      object FileOpen: TMenuItem
        Caption = 'Open...'
        ShortCut = 16463
        OnClick = FileOpenClick
      end
      object FilePrint: TMenuItem
        Caption = 'Print...'
        ShortCut = 16464
        OnClick = FilePrintClick
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object FileExit: TMenuItem
        Caption = 'Exit'
        ShortCut = 32883
        OnClick = FileExitClick
      end
    end
    object ZoomPopup: TMenuItem
      Caption = 'Zoom'
      OnClick = ZoomPopupClick
      object Zoom25: TMenuItem
        Caption = '25%'
        OnClick = Zoom25Click
      end
      object Zoom50: TMenuItem
        Caption = '50%'
        OnClick = Zoom50Click
      end
      object Zoom100: TMenuItem
        Caption = '100%'
        OnClick = Zoom100Click
      end
      object Zoom150: TMenuItem
        Caption = '150%'
        OnClick = Zoom150Click
      end
      object Zoom200: TMenuItem
        Caption = '200%'
        OnClick = Zoom200Click
      end
      object N2: TMenuItem
        Caption = '-'
      end
      object ZoomPageWidth: TMenuItem
        Caption = 'Page Width'
        OnClick = ZoomPageWidthClick
      end
      object ZoomPageHeight: TMenuItem
        Caption = 'Page Height'
        OnClick = ZoomPageHeightClick
      end
      object ZoomWholePage: TMenuItem
        Caption = 'Whole Page'
        OnClick = ZoomWholePageClick
      end
    end
  end
  object PrintDialog: TPrintDialog
    Options = [poPageNums, poWarning]
    Left = 9
    Top = 16
  end
end
