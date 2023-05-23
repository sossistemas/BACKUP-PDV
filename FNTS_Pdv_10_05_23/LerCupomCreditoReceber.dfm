object frmLerCupomCreditoReceber: TfrmLerCupomCreditoReceber
  Left = 0
  Top = 0
  ActiveControl = edCupom
  BorderStyle = bsNone
  BorderWidth = 3
  ClientHeight = 140
  ClientWidth = 302
  Color = 8404992
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 302
    Height = 140
    Align = alClient
    BevelOuter = bvNone
    Color = 16744576
    ParentBackground = False
    TabOrder = 0
    object Label1: TLabel
      Left = 58
      Top = 39
      Width = 186
      Height = 25
      Caption = 'Cupom de Cr'#233'dito'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -21
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object edCupom: TcxTextEdit
      Left = 51
      Top = 68
      ParentFont = False
      Properties.Alignment.Horz = taCenter
      Style.Font.Charset = DEFAULT_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -21
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = [fsBold]
      Style.IsFontAssigned = True
      TabOrder = 0
      OnKeyPress = edCupomKeyPress
      Width = 201
    end
  end
  object qrCupom: TUniQuery
    Connection = frmModulo.Conexao_Servidor
    SQL.Strings = (
      'select distinct'
      '  cre.*'
      'from'
      '  devolucao dev'
      '  inner join devolucao_credito cre on cre.id = dev.id_credito'
      'where dev.CUPOM_DEVOLUCAO = :cupom')
    Left = 34
    Top = 16
    ParamData = <
      item
        DataType = ftString
        Name = 'cupom'
        ParamType = ptInput
        Value = '001000000900390'
      end>
    object qrCupomID: TIntegerField
      FieldName = 'ID'
      Required = True
    end
    object qrCupomVENCIMENTO: TDateField
      FieldName = 'VENCIMENTO'
    end
    object qrCupomVALOR: TFloatField
      FieldName = 'VALOR'
    end
    object qrCupomVALOR_UTILIZADO: TFloatField
      FieldName = 'VALOR_UTILIZADO'
    end
    object qrCupomVALOR_SALDO: TFloatField
      FieldName = 'VALOR_SALDO'
    end
    object qrCupomCUPOM_UTILIZADO: TStringField
      FieldName = 'CUPOM_UTILIZADO'
      Size = 50
    end
  end
end
