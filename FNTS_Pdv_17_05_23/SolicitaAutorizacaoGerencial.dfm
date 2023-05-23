inherited frmSolicitaAutorizacaoGerencial: TfrmSolicitaAutorizacaoGerencial
  Caption = 'frmSolicitaAutorizacaoGerencial'
  ClientHeight = 210
  ClientWidth = 571
  FormStyle = fsStayOnTop
  ExplicitWidth = 575
  ExplicitHeight = 214
  PixelsPerInch = 96
  TextHeight = 13
  inherited Panel1: TPanel
    Width = 571
    Color = 16384
    ExplicitWidth = 571
    inherited Label1: TLabel
      Left = 3
      Width = 486
      Height = 39
      Align = alClient
      Alignment = taCenter
      Caption = 'Solicita'#231#227'o de Autoriza'#231#227'o'
      ExplicitLeft = 3
      ExplicitWidth = 486
    end
    inherited Bevel1: TBevel
      Width = 3
      ExplicitWidth = 3
      ExplicitHeight = 39
    end
    inherited AdvMetroButton2: TAdvMetroButton
      Left = 530
      Visible = False
      ExplicitLeft = 530
    end
    inherited AdvMetroButton1: TAdvMetroButton
      Left = 489
      Visible = False
      ExplicitLeft = 489
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 39
    Width = 571
    Height = 171
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    Color = 16384
    ParentBackground = False
    TabOrder = 1
    object Panel3: TPanel
      Left = 2
      Top = 2
      Width = 567
      Height = 128
      Align = alClient
      BevelOuter = bvNone
      Color = 15728620
      ParentBackground = False
      TabOrder = 0
      object lbTempo: TcxLabel
        Left = 0
        Top = 77
        Align = alBottom
        AutoSize = False
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clRed
        Style.Font.Height = -16
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = []
        Style.IsFontAssigned = True
        Properties.Alignment.Horz = taCenter
        Properties.Alignment.Vert = taVCenter
        Height = 51
        Width = 567
        AnchorX = 284
        AnchorY = 103
      end
      object cxLabel2: TcxLabel
        Left = 0
        Top = 0
        Align = alClient
        Caption = 
          'Aguarde!'#13#10#13#10'Solicitando autoriza'#231#227'o superior para o cancelamento' +
          '.'
        ParentFont = False
        Style.Font.Charset = DEFAULT_CHARSET
        Style.Font.Color = clWindowText
        Style.Font.Height = -16
        Style.Font.Name = 'Tahoma'
        Style.Font.Style = [fsBold]
        Style.IsFontAssigned = True
        Properties.Alignment.Horz = taCenter
        Properties.Alignment.Vert = taVCenter
        AnchorX = 284
        AnchorY = 39
      end
    end
    object cxButton1: TcxButton
      Left = 2
      Top = 130
      Width = 567
      Height = 39
      Align = alBottom
      Action = Action1
      Caption = 'CANCELAR (ESC)'
      Colors.Default = clMaroon
      Colors.DefaultText = clWhite
      Colors.Normal = clMaroon
      Colors.NormalText = clWhite
      Colors.Hot = clMaroon
      Colors.HotText = clWhite
      Colors.Pressed = clMaroon
      Colors.PressedText = clWhite
      Colors.Disabled = clMaroon
      Colors.DisabledText = clWhite
      LookAndFeel.Kind = lfUltraFlat
      LookAndFeel.NativeStyle = False
      SpeedButtonOptions.Flat = True
      SpeedButtonOptions.Transparent = True
      TabOrder = 1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
  end
  object ActionList1: TActionList
    Left = 394
    Top = 121
    object Action1: TAction
      Caption = 'CANCELAR'
      ShortCut = 27
      OnExecute = Action1Execute
    end
  end
  object qrLib: TUniQuery
    Connection = frmModulo.Conexao_Servidor
    SQL.Strings = (
      'select * from liberacao_pdv where codigo = :cod')
    Left = 55
    Top = 120
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'cod'
        Value = nil
      end>
    object qrLibCODIGO: TStringField
      FieldName = 'CODIGO'
    end
    object qrLibCOD_OPERADOR: TStringField
      FieldName = 'COD_OPERADOR'
      Size = 6
    end
    object qrLibTIPO: TStringField
      FieldName = 'TIPO'
      Size = 1
    end
    object qrLibDATA: TDateField
      FieldName = 'DATA'
    end
    object qrLibHORA: TStringField
      FieldName = 'HORA'
      Size = 8
    end
    object qrLibMENSAGEM: TStringField
      FieldName = 'MENSAGEM'
      Size = 200
    end
    object qrLibSTATUS: TStringField
      FieldName = 'STATUS'
      Size = 1
    end
    object qrLibAPROVADO_DATA: TDateField
      FieldName = 'APROVADO_DATA'
    end
    object qrLibAPROVADO_HORA: TStringField
      FieldName = 'APROVADO_HORA'
      Size = 8
    end
    object qrLibAPROVADO_USUARIO: TIntegerField
      FieldName = 'APROVADO_USUARIO'
    end
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 426
    Top = 65
  end
end
