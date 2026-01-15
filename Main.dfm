object MainForm: TMainForm
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu]
  Caption = 'Atualizador Sistema MMSoftwares'
  ClientHeight = 699
  ClientWidth = 1125
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = True
  Position = poMainFormCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 59
    Width = 1125
    Height = 640
    Align = alClient
    TabOrder = 0
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 1123
      Height = 597
      ActivePage = TabComando
      Align = alClient
      TabOrder = 0
      object TabComando: TTabSheet
        Caption = 'Comandos a executar'
        object AExecutarMemo: TMemo
          Left = 0
          Top = 0
          Width = 1115
          Height = 569
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Courier New'
          Font.Style = []
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
        end
      end
      object TabErros: TTabSheet
        Caption = 'Erros Ocorridos'
        ImageIndex = 1
        TabVisible = False
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object AErrorMemo: TMemo
          Left = 0
          Top = 0
          Width = 1115
          Height = 569
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Courier New'
          Font.Style = []
          ParentFont = False
          ScrollBars = ssVertical
          TabOrder = 0
        end
      end
    end
    object Panel3: TPanel
      Left = 1
      Top = 598
      Width = 1123
      Height = 41
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 1
      object lbExecutados: TLabel
        Left = 15
        Top = 6
        Width = 161
        Height = 14
        Caption = 'Comandos Executados...:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Courier New'
        Font.Style = []
        ParentFont = False
        Transparent = False
      end
      object lbErros: TLabel
        Left = 15
        Top = 22
        Width = 161
        Height = 14
        Caption = 'Erros Ocorridos.......:'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clRed
        Font.Height = -11
        Font.Name = 'Courier New'
        Font.Style = []
        ParentFont = False
        Transparent = False
      end
      object GaugeProgresso: TGauge
        Left = 460
        Top = 10
        Width = 320
        Height = 21
        ForeColor = clNavy
        Progress = 0
      end
      object GaugeScripts: TGauge
        Left = 784
        Top = 10
        Width = 320
        Height = 21
        ForeColor = clGreen
        Progress = 0
      end
    end
    object PanelProcesso: TPanel
      Left = 311
      Top = 210
      Width = 503
      Height = 180
      Color = clInfoBk
      ParentBackground = False
      TabOrder = 2
      Visible = False
      object lbProcTitulo: TLabel
        Left = 16
        Top = 12
        Width = 159
        Height = 16
        Caption = 'Processo de Atualizacao'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object lbProcEtapa: TLabel
        Left = 16
        Top = 44
        Width = 51
        Height = 13
        Caption = 'Etapa: 0/0'
      end
      object lbProcDesc: TLabel
        Left = 16
        Top = 68
        Width = 71
        Height = 13
        Caption = 'Aguardando...'
      end
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1125
    Height = 59
    Align = alTop
    TabOrder = 1
    object btnExecutarProcesso: TButton
      Left = 16
      Top = 12
      Width = 137
      Height = 25
      Caption = 'Executar Processo'
      TabOrder = 0
      Visible = False
      OnClick = btnExecutarProcessoClick
    end
    object btTestarScripts: TButton
      Left = 168
      Top = 12
      Width = 137
      Height = 25
      Caption = 'Testar Scripts'
      TabOrder = 1
      Visible = False
      OnClick = btTestarScriptsClick
    end
  end
  object FDScript: TFDScript
    SQLScripts = <>
    Connection = dtmConnec.FDConnection
    ScriptOptions.SpoolOutput = smOnAppend
    ScriptOptions.SpoolFileName = 'teste.txt'
    Params = <>
    Macros = <>
    OnSpoolPut = FDScriptSpoolPut
    Left = 128
    Top = 152
  end
  object TrayIcon: TTrayIcon
    PopupMenu = PopupMenu
    Left = 613
    Top = 164
  end
  object PopupMenu: TPopupMenu
    Left = 581
    Top = 324
    object Abrir1: TMenuItem
      Caption = '&Abrir'
      OnClick = Abrir1Click
    end
    object Fechar1: TMenuItem
      Caption = '&Fechar'
      OnClick = Fechar1Click
    end
  end
  object FTP: TIdFTP
    IPVersion = Id_IPv4
    TransferType = ftBinary
    NATKeepAlive.UseKeepAlive = False
    NATKeepAlive.IdleTimeMS = 0
    NATKeepAlive.IntervalMS = 0
    ProxySettings.ProxyType = fpcmNone
    ProxySettings.Port = 0
    Left = 301
    Top = 404
  end
  object Timer: TTimer
    Interval = 3600000
    OnTimer = TimerTimer
    Left = 96
    Top = 304
  end
end
