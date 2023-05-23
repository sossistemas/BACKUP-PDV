object frmCliente_consulta: TfrmCliente_consulta
  Left = 3
  Top = 0
  BorderStyle = bsNone
  Caption = 'frmCliente_consulta'
  ClientHeight = 583
  ClientWidth = 790
  Color = 12615680
  Constraints.MinHeight = 532
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 790
    Height = 583
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    Color = clBlack
    ParentBackground = False
    TabOrder = 0
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 786
      Height = 579
      Align = alClient
      BevelOuter = bvNone
      BorderWidth = 2
      Color = clWhite
      ParentBackground = False
      TabOrder = 0
      object pnFundo: TPanel
        Left = 2
        Top = 2
        Width = 782
        Height = 575
        Align = alClient
        BevelOuter = bvNone
        Color = 6960384
        ParentBackground = False
        TabOrder = 0
        object lb_status: TRzLabel
          Left = 31
          Top = 550
          Width = 410
          Height = 25
          AutoSize = False
          Caption = 'Informe o C'#243'digo, CPF ou Nome do Cliente...'
          Font.Charset = ANSI_CHARSET
          Font.Color = clWhite
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          ShadowColor = clBlack
          TextStyle = tsShadow
        end
        object RzLabel1: TRzLabel
          Left = 30
          Top = 100
          Width = 79
          Height = 29
          Caption = 'Cliente:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clWhite
          Font.Height = -24
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          ShadowColor = clBlack
        end
        object lb_cliente: TRzLabel
          Left = 30
          Top = 19
          Width = 312
          Height = 45
          Caption = 'Consulta de Cliente'
          Font.Charset = ANSI_CHARSET
          Font.Color = clWhite
          Font.Height = -37
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          Transparent = True
          ShadowColor = clBlack
        end
        object grid: TNextGrid
          Left = 26
          Top = 156
          Width = 735
          Height = 388
          Touch.InteractiveGestures = [igPan, igPressAndTap]
          Touch.InteractiveGestureOptions = [igoPanSingleFingerHorizontal, igoPanSingleFingerVertical, igoPanInertia, igoPanGutter, igoParentPassthrough]
          AppearanceOptions = [aoHighlightSlideCells]
          BorderStyle = bsNone
          Caption = ''
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Courier New'
          Font.Style = []
          GridLinesStyle = lsVerticalOnly
          HideScrollBar = False
          Options = [goGrid, goHeader, goSelectFullRow]
          ParentFont = False
          TabOrder = 0
          TabStop = True
          OnEnter = gridEnter
          OnKeyPress = gridKeyPress
          object NxNumberColumn1: TNxNumberColumn
            DefaultValue = '0'
            DefaultWidth = 49
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Courier New'
            Font.Style = []
            Header.Caption = 'C'#243'digo'
            Header.Alignment = taCenter
            Header.Font.Charset = DEFAULT_CHARSET
            Header.Font.Color = clWindowText
            Header.Font.Height = -11
            Header.Font.Name = 'Tahoma'
            Header.Font.Style = []
            ParentFont = False
            Position = 0
            SortType = stNumeric
            Width = 49
            Increment = 1.000000000000000000
            Precision = 0
          end
          object NxTextColumn1: TNxTextColumn
            DefaultWidth = 334
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Courier New'
            Font.Style = []
            Header.Caption = 'Nome'
            Header.Alignment = taCenter
            Header.Font.Charset = DEFAULT_CHARSET
            Header.Font.Color = clWindowText
            Header.Font.Height = -11
            Header.Font.Name = 'Tahoma'
            Header.Font.Style = []
            ParentFont = False
            Position = 1
            SortType = stAlphabetic
            Width = 334
          end
          object NxTextColumn2: TNxTextColumn
            DefaultWidth = 134
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Courier New'
            Font.Style = []
            Header.Caption = 'CPF/CNPJ'
            Header.Alignment = taCenter
            Header.Font.Charset = DEFAULT_CHARSET
            Header.Font.Color = clWindowText
            Header.Font.Height = -11
            Header.Font.Name = 'Tahoma'
            Header.Font.Style = []
            ParentFont = False
            Position = 2
            SortType = stAlphabetic
            Width = 134
          end
          object NxTextColumn3: TNxTextColumn
            DefaultWidth = 201
            Font.Charset = ANSI_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Courier New'
            Font.Style = []
            Header.Caption = 'Endere'#231'o'
            Header.Alignment = taCenter
            Header.Font.Charset = DEFAULT_CHARSET
            Header.Font.Color = clWindowText
            Header.Font.Height = -11
            Header.Font.Name = 'Tahoma'
            Header.Font.Style = []
            ParentFont = False
            Position = 3
            SortType = stAlphabetic
            Width = 201
          end
        end
        object ed_cliente: TRzEdit
          Left = 135
          Top = 102
          Width = 626
          Height = 31
          Text = ''
          Color = clWhite
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          FrameStyle = fsNone
          FrameVisible = True
          FramingPreference = fpCustomFraming
          ParentFont = False
          TabOrder = 1
          OnEnter = ed_clienteEnter
          OnKeyPress = ed_clienteKeyPress
        end
      end
    end
  end
end
