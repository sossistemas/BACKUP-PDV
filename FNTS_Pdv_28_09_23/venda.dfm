object frmVenda: TfrmVenda
  Left = 302
  Top = 144
  BorderStyle = bsNone
  Caption = 'Venda Balc'#227'o'
  ClientHeight = 828
  ClientWidth = 1584
  Color = 1261591
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  PopupMenu = pop_principal
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnKeyUp = FormKeyUp
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pn_tampa: TFlatPanel
    Left = 4000
    Top = 6000
    Width = 1160
    Height = 621
    ParentColor = True
    Visible = False
    ColorHighLight = clBlack
    ColorShadow = clBlack
    TabOrder = 0
  end
  object fl: TFileListBoxEx
    Left = 870
    Top = 493
    Width = 121
    Height = 112
    Mask = '*.cga'
    ShowGlyphs = True
    TabOrder = 1
    Visible = False
    Version = '1.3.0.3'
  end
  object pnFundo: TPanel
    Left = 0
    Top = 0
    Width = 1584
    Height = 828
    Align = alClient
    BevelOuter = bvNone
    Color = clBlack
    ParentBackground = False
    TabOrder = 2
    object pn_principal: TPanel
      Left = 0
      Top = 0
      Width = 1584
      Height = 828
      Align = alClient
      BevelOuter = bvNone
      Color = 14086904
      ParentBackground = False
      TabOrder = 0
      object imgFundo: TImage
        Left = 0
        Top = 0
        Width = 1584
        Height = 828
        Align = alClient
        ExplicitTop = 4
        ExplicitWidth = 1276
        ExplicitHeight = 628
      end
      object ed_foco: TEdit
        Left = -500
        Top = 504
        Width = 0
        Height = 21
        TabOrder = 0
        Text = 'ed_foco'
      end
      object v_fundo: TRzPanel
        Left = 16
        Top = 0
        Width = 1553
        Height = 809
        BorderOuter = fsNone
        Color = 2434341
        TabOrder = 1
        Transparent = True
        object Panel5: TRzPanel
          Left = 29
          Top = 41
          Width = 1498
          Height = 744
          Align = alClient
          BorderOuter = fsNone
          Color = 5329233
          TabOrder = 0
          Transparent = True
          object v_barra_superior: TRzPanel
            Left = 0
            Top = 0
            Width = 1498
            Height = 57
            Align = alTop
            BorderOuter = fsNone
            Color = 8421440
            TabOrder = 0
            Transparent = True
            object Panel7: TPanel
              Left = 0
              Top = 0
              Width = 1498
              Height = 57
              Align = alClient
              BevelOuter = bvNone
              BorderWidth = 5
              Color = 8404992
              ParentBackground = False
              TabOrder = 0
              object lb_produto: TRzLabel
                Left = 5
                Top = 5
                Width = 1488
                Height = 47
                Align = alClient
                Alignment = taCenter
                Caption = 'C A I X A    L I V R E'
                Font.Charset = ANSI_CHARSET
                Font.Color = clWhite
                Font.Height = -35
                Font.Name = 'Tahoma'
                Font.Style = []
                ParentFont = False
                Transparent = True
                Layout = tlCenter
                ShadowColor = clBlack
                ExplicitWidth = 321
                ExplicitHeight = 42
              end
            end
          end
          object v_espacamento_superior: TRzPanel
            Left = 0
            Top = 57
            Width = 1498
            Height = 20
            Align = alTop
            BorderOuter = fsNone
            Color = 6447714
            TabOrder = 1
            Transparent = True
          end
          object Panel4: TRzPanel
            Left = 0
            Top = 77
            Width = 1498
            Height = 504
            Align = alClient
            BorderOuter = fsNone
            Color = 5329233
            TabOrder = 2
            Transparent = True
            object v_quadro_direito: TRzPanel
              Left = 601
              Top = 0
              Width = 897
              Height = 504
              Align = alClient
              BorderOuter = fsNone
              Color = clWhite
              TabOrder = 0
              Transparent = True
              object pnBotoesInferior: TPanel
                Left = 0
                Top = 433
                Width = 897
                Height = 71
                Align = alBottom
                BevelOuter = bvNone
                BorderWidth = 5
                Color = 8404992
                ParentBackground = False
                TabOrder = 0
                Visible = False
                object btnCancelaVenda: TAdvGlowButton
                  Left = 121
                  Top = 5
                  Width = 112
                  Height = 61
                  Align = alLeft
                  Caption = 'Cancelar Venda'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F90000009AC4944415478DABC99FB6F5BE519C71F093189494813EC0F80FD30
                    0DD05426B631DAAE491B274DDDE6E6E3737C5CE79E1CDB8993A6494ADB946449
                    53CA65FB6597A24D306063058198D46DD286C6B81746D7D116145AFB388993C6
                    3EBE9DF31EDBC7CED549BFFBC169523769D2A6B01F3E3F1CEB7DE58F9FC7CFF3
                    3E7E4DE8EFA73C068FD2425F1F31DE49C6DE364AD67988092ECAD47552426C25
                    436C279D7793C64B0F68BCD3C304E74B4C70BEABF1D284C64BC6226126BA3E62
                    82F3A4C63B0FE8BCF361C6BB48E7DDC4C4164AD5B693B1B79D747B2BA5440F31
                    CE49D9C33D842707E97A1FBA59C1C9866E4A889E7B74C1DDC378A757E3255C45
                    B536E3DA678D97A072CDD0ACCBCF8C7706182F1D67B696FB8CBA7D9476ECFB0A
                    041D6D94AC6F2326B8EF4CDA5B9FD104673227B352E8665073FBB2BAE8FA5DCA
                    DE76AFEEF0DC6E04DB2959EB298B73CD232AD7B461B195D16D826A6DD6D95E77
                    BD616F27C6B9D6103CDA9FCFF141C2403FE9363731CE35A8F11234EB5723964F
                    2E13CCEA3AC178172D3CD14B78FAD80A1F425F7F3E034769E1482F319BF364DC
                    D298F73DFA3A588CE6BFB2078F10060757F8106AFA7354F713EA9EA46C6D0FA9
                    42E3AB5FA7D46AC4F9860F676B0E12EA8FE75C16BD08DD03391E3F4A38708C0C
                    5BDBD158453DFEDF82B1CA7A24B996D7D03948383448E8EE27740F1019CE2E32
                    9C5D94F63C4E29574761DC7A0B69159CD00409ACA00AEC2766B02D65391E3543
                    373BC01C9E9B97B44A88730D484A6DF586E7E09217CD361EA259A987A66A0EDC
                    112B6F54556BD3F2269B0BAAD901B558805A625BC9760EAA4940F4B99711FEF3
                    DF1179E5CD1C7F790BCAFE3EC41F2901DB695F7D6FB10DEA4EFBF2075D2CC668
                    790326ABBBEE9D958ED06CE321A2F9F65E9AEF1C20A3AEFD9958E575A9E59AC0
                    EA3BC0A40360CDDD2B7178C01A3A1166498400288B44008CBEF1575CDE6C46C2
                    F9F8EA7BA503604D5DD06D2EB0BC54372055EB797D7E7F3FCDB7F71219622719
                    B6FD776956299BD7E72C8D508BAC303E3D8FC9F97964D2532B3116999C417A6A
                    16996B4867A6914AA69136A656DD3B959D47321845886B845E560B76358ABC84
                    38D78494B0EF4143EC224A3B3A2921B4F4C4AF4D2D2F41AB6A805A50897454C3
                    34B02633D905CCCC6431333B8F99D9794CCF64310360668D3D590053002E950A
                    B8BCAD1C297BCB92A46A6D822EB87F9BDEDB49C4ECADC4046770C529C1354133
                    3B907CEE0F489DBD003D330D3D3D053D95B96D1273F3601786107DF975C83B79
                    7CF95829826535482E4AAAD66630C19962F6D63B89D95CDFB96165896E683F28
                    429C97105475047503C170FCF68831285700B9B1034377DF07B9C8025FB11517
                    0B2B102CBF4ED2E62A2726B83AD63A6359A91D5A751B94893042E96928516D89
                    489C41896A082D3E87630CE158FE6B57D75C259C4C4361290C3BDCF06E2A846C
                    E2209B38F88A2CF992B9EA7E8198E83AB9E6595B590F56568BF0B92184B20B79
                    6FB611C2D373080D8F432EB5C1F7D8AE25C13CC9B21A24C516E8A2EB0362A2EB
                    43955B43906B02DBC121F2F6E95C2B896A8869098C4D4471E6FC250C8F2988B3
                    2494A8860B5F0EE3FCD03094A886384B623810C299CF2E62784C414CCBAD0967
                    AF2074E122E4C20AF8B6EEC9135C92DC5E81E09E6A188E563F69BC145ABBC337
                    831554217AEA9F7982A3E3617CF0E9E7904783D0120694A886FF5CF0E2CCB94B
                    50A21AB48401DFC804DEFFE43CE4D120D445C10880E0E9B3F03D5A02B9A06285
                    E0B264258265351A69BC34B5DE31C4369721FAEAA93CC1C0E5083E3E3B047F20
                    04554F41896AF8EC0B19FFFDDC0725AA41D553904783387DE60BF80321C4AF11
                    9CF8C77BF06D2A84BCA36A55C1AB92DE1D55999B13DC5286D8AF7E9F138CB1DB
                    16BCFCC6DFE07D682BE422CB0D0517C990669522EB0A1654217EE4698466B250
                    D4C46D0B8EBFF81ABCDF7D743D39C8264E27DDEEFE64CD22E125B0222BD48E3E
                    28893414DDD8B8608C21022070FC97EB0AFA4C160C9BF851D245D7EBEB0AEE14
                    A14ADD50222A945466E3825A12E1B9798CB41DCEA5784D410EC3C5D68F49135C
                    87D6FD31B4DB01CDD18A70208850667AC382E1441AE1441AC3F5EDF03D5CB86E
                    8AFDC5DC2BA48BEE07D71D26CB6BA1714D087F2923349BDDB8A0310925A2C1CF
                    37C1F743D3FA112CB1DA4913DDA4F192B6A660550398B91A91B39F23B4808D0B
                    CE6411F28D422EE1E1DB6C5E2F8273FE12EE9B64D8F7534270FF62C5B8B5EA69
                    F21142D8B86004C0C4FBFFCE45AFB0F2C6723B2C904BACAF8D94D98812C23E4A
                    F06DDF5EB75036EF41EC8F6FDEB6E0E5536FC1FBFD9FDEB049FB8BACB864AA84
                    B7DCFA6399138966DA7A68767F2F256B3D2FC42A1BD616FCCD8B080188B32446
                    C7C378FF930BF08D2C1F7567CE5DC2A79F5D5C3AEABCC31378F7F439F8468250
                    592AD703FFF426BC0F6CBE6193BE545481F122C7DB5A590769E5ED440BEE23B4
                    E0E9A3D9868377C72B1B67D51B4C366C6B39E2CF9E40E80A1051139850E2F00E
                    5FC6782886A8AA43896AF00742F0074250A21AA2AA8EF160145E7F6E4D24AE23
                    0260ECB9976FDC038B2C18329563BCB2FEFE68958722952E22BDA683F4EA0E4A
                    3674915ED72AC6B9D5A3C8B65522DEFB2C942B80624C22924C43CD4C2392CC20
                    A4A7A0240CC4D35388A7A7A0240C84F414A2A9CCD21A2561200A20F0F313AB0B
                    1673B85458855091A3276E69A030574D614B0D51B6B33747771F2D74F553CAE6
                    793E56B552929588509BBB11D29250E6E6114E666E8DF414A200469F780ADEEF
                    3DB66A6ABDBB2CEF19E51D34653E4053E64E9A327712A1F6E8220384FAE334E7
                    384C314BFD3B2B04ADCDD0050913CD5DF036EEC37063C7ADD1B41F235237E4DD
                    76C8852BC72C6F41857794B7DF1170D7D348432D8D34E620F4F42F736480AEF4
                    0E50F6600F319BF33DD5D298778BC07809BA4940E01113861EDA0AEFA642F836
                    15C27B5314E4AA775B19E49DC2B25851157C26CBC54C61E7B7E676FC8CE64A8E
                    D06C71CF128463FDCB0CF6139E1A240C1ECD5DBF59A493D7178D6E7321696FC1
                    D82E3B2E1656DCCC44B2E6CCE72BB1BC259770DF982FEC256C7D96B06590B0E5
                    D8126B5C60BAC8B0B753A2B6B539CE35CFE52E30AF9114DD1B96F41559E03359
                    306A120E8F9A45F2957234537088B0ED4942417F1EEB5F0137B411135CF727C4
                    9697AEFEA856F966E8361712A21B63E65B902CCA8D517E93706AD86C7B78AC54
                    A440A948BE9D960D0AEE6DA3449D8798CD4D93F55D94105B7FA40BAE5F33DE19
                    CD493A7392BB445CDC5EB1629EBB4E30259BB8176433BF7DCC5C4B813D7B2950
                    6AA3D152DB572078EDDF10B676D279F75D1A2F9569BCF3846E73BD93B4B778C7
                    CCF6B8774795219B384336719ADF64F5FB4DD60FE562EE797FB15594775AEF91
                    4BACE4DF2DD098B9864677DB6F4AF07F0300CBC7E9D9495EFE5F000000004945
                    4E44AE426082}
                  Transparent = True
                  TabOrder = 0
                  OnClick = CancelarCupom1Click
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
                object btnCancelaItem: TAdvGlowButton
                  Left = 5
                  Top = 5
                  Width = 116
                  Height = 61
                  Align = alLeft
                  Caption = 'Cancelar Item'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F90000009564944415478DAB4996B6C5BE51DC6FF12621293264DB02FDB00D1
                    7DD844B742B9AF6D1A9A9B9DA4F812DBE7E2A471EC243E766E6D9396B409C99C
                    A494316DD2C6C404B41B1250AA6E6C65373469C03676D13498B6412EBEC44D5B
                    484BE3738EEF766227E9B30FC776E2DA4E5C281F9E0FB6FCFAFDF9F9BFEFFBFC
                    CF6B82CB45799A18A7D5D15192190745CDBD146EEB21997550BCAD9F427C3745
                    F93E921927494CE7DD122B74CBACF0A2CC0A6F4B8CFD4389B14733BA2CF38EBF
                    CAAC704A6485C33263DF2E330E0A324E92F92E8A58FA28DADC4741733745F81E
                    928D022D1F1D223C3941D7F250B98009DB210AF13DB70659E790CC086E89B123
                    2BD1D489F5AF25C60ED1D80969DDFB32239C9319FB7199EBBA2BDAB69F622DFB
                    6F10A0B59764D67973D8DCFDB4C40A6105A610A81C89CAB8E520EF783E62EEBD
                    2DD8D2F329019B7B296CE9D1048C9D7ED1D8F189C10ADDED8068EA0CCACD4E6B
                    D4DC47B2D1B101E0B82B5FC72708632E0A724E928DC2B8C4D8F34A75E3A45442
                    36399E951907AD3E3142F8EEB1021EC2A82B5F63E3B43A3C4232279C0A18DA21
                    99EC9F015C819B6F2E0F0E1326260A7808AD2E45FB5C84B62769D9324422DBFE
                    EA6709554C01C6F64EAA7590603DAEB064B80887C6143D3E4E387C8CA25CEFF8
                    82CE5AFACB3E4DB93718BBA0B7226CEC3A8DFE09C29109C22117E1D018515418
                    A0A83040B19EC729E238B027602A5156B31392C68240851662630B247357F960
                    E58C35D91130DA10B6F75AA33D83392E4AB51FA1947D8892AD876F5AD0B68BA2
                    A9A37030EF8458C340D2DB10F9F96F205B0F42ACD02A136F06773D634D9DB8A2
                    B521B16FE0B6947D9852ED478856FA4668A57F8CA26D7D4F2FE8AD25263041AC
                    D022FAAFFF620940CC7F01A28A8758A9DF18F2138C5DD0DB10B1F49C5939E8A2
                    95BE11A228DF4F51EEE02D92C9BE5C70CEB102C4C61688554644DF7D1F4900F1
                    A5652401C4CE7D08692348B31362B50962850ED1F73EC81FEBBF08A9B119929A
                    87C4390A378CB1031176FFD6283F40146BE9A710DB351428565A9D15928A43CC
                    7D4E9960318D786209F1645A799D85ACD0E6AF2BB31362B511E26E1DA2EFFE2F
                    07174F2C219E48611140321C47D0D00E5963C926CCBAE8EC4090753E176BEE27
                    92CDDD24B3C2474553A2A91D929A47E4B5DF2B935C456E922C64CC7F0162BD19
                    E22E8D0299756EB77E9DEB991F96584262E52AD2003E7EE32DCC566A21EBAD08
                    F18515905921229BBB6F2699737CADE41A6205484D3604EEAB45F88553480248
                    603D644A819C3D0FA9DE0CF1D126886A5E296B9E73A93CB82BA7CF62EA8E7B30
                    F9880AFE061E41CE5104B21332E7D092CC3A0E6C98B19C0392CE8AC0BDD508BF
                    F08A0279B508E4DC4588155A04EEAD46F4DF1F143A975EC532802BAF9EC5E4ED
                    F7C0B3B3113E358BE93D3AF8D5AC02C939AF75F124C9BCE3D4A687EF7AC8E75F
                    C92F7772ADDCD177DF47E4CDBF153AB79C813BF54B4CDEB91DEE1D0DF0AA1978
                    6B8DF0D61A31BD4787D96B214D76C89CF0179279C73BA2B18C74C8426EAF29EE
                    6422854476092453796B4E71EE5798BAEB3EB81F51C3AB66737039C82A1DFC2A
                    26576ED1D40999137C2431F6F9B213817340D2B521705F4D092753882FA67280
                    6B70673175E7F69270EB9DBC06522289B127AF2B4F3390E2032A849EF929E2B1
                    04E2E99535C08C12A96524928BB8FCFC4B98DA723F3C3B1A3684CB8354E720E3
                    D70FC8D821357743DCAD43607B0D62B317D6CABA4E8B0062F31F637A5B25A6B6
                    DC0FDF5EF3A670F94EB208728E384926FBC7D705973DE72AF588FEF33F0ADC62
                    BA0030BEA8408627DD703FACC2F4D60AF8EAB9F221ABF4F0ABD92005CDCE7F94
                    B549CA4888447A1589F46A6ED3C417D3480388F8E630F3601DA6EFDE015F4379
                    909E1A03BC75C67314E41D67CA02DC2C21D2AB5804B00820B17215B1B8F27E0E
                    D27B4E71F21BE5417A6A0CF0A999BF93C43A8E6CFA30946999364B884BCF9CC4
                    87A34F2305289F59178B6900118F1FEE87324E6E526E4F8D017E15FB320579E7
                    D6CDE1322DD37BEF974E88D36731F9E56FE283DBBE8E4B3F3A81740652715259
                    97D972BB1FACC5F4D65D1B42BA6B0DB8A0E6CD24714E9218BB54328B1B5B2056
                    E7B75BA512C2B3AB11DE2A3D26BFF22D5CFAE10B39C8DC39B96E4DBA1F5163E6
                    815A78554C29C8F45C3DF7798A9A0F5288757EBF68BBA5B742AC630BDBAD8D12
                    42C5C053A9C1E457B7613E0B79158865CA1D4BA4B00C20168CC053A9817B6743
                    6179AB9AE0A9674ECF36351385D8FD14627ABF5474A318DA21AA7944CEFC3617
                    6DB1F81212AB9B24848A8167B70693B76FCB73323B360D407CE36D782AF6C2B3
                    5B930F586782FB513DCED7310F5F7CCC4CB4D43B44A9832314B6F49C5CD0DB4A
                    B65BA19FBC84446622A5657A5DE94A4A254416729D934BD9B12FBFA68CDDB5B7
                    A0C453953A5CD4B5FC31D9758092F65EA255E730ADF68C52CA36F88580BE3D55
                    7011C43920E96D4A2773E2552C01B8F28BDF61EA8E7BE0D999DF951485ACD428
                    6BF2D99F6105C09533BFC6E4EDDB8AC2796B8C98DCA5C105B675CBA5AE2E9AEF
                    B013055B0F5070DF010ADB0628D8D6CD078CB622F92B40D2DB1054F1081C1EC7
                    D48E06CC7C5B0D5F19D99A859CB9BF1AE7BB1E8767D75E78763616DD1C338FEA
                    30DF641B0AB4F6D16556A0CBAC40B4DC3FA2E8D028AD0EB828C2F59C58682A01
                    69EC845C6DC27C3D8F995A8372DA97135D2A06DE2A3DA6B755C253A929EAFA4C
                    A50E9E3AC39FA307FB29797498920383943C344804CB78466304EB714AB71CA5
                    0583F5AD52E7A2CC08089BBB704967C17495AE7CC84DE4AED4BBE74CCD37CDD9
                    ADE46FDD477E8B22C2906B4DC3637475648C96078748E6843F050C1DC5215901
                    1173F70D8174EF6982A7DA301DDF3FF0C5B4EB3B941E1EA6D4D0504E8463AE35
                    4DB8084F4D1026C628C8394836D84F89251E07644640E4533AE9A936C0536BF8
                    834F65FADCCA13A3841F7C4FB9E13A762CA792179841C64131732F852CDD9D01
                    63675AB9C0BC314E7AAA9BE0A932E05C037B744ECB914F65A4A5C347094F3D79
                    1D37AC2681A22D7D14B6F591CC3AB684F8AE17B30FD585905927F51B43D618E1
                    AD31C05FCFBE3EFB18BFFDBCD64CE7B51C79EB0C9F10B0B997426DBD24734E4A
                    580728C4773F14641D3F9619E1CAB5576AD98D339385CC6A0D30E2A9359E9C53
                    3355F3060BF99B5A68EE319EE634370270DDDF1031AE8F828CF31689B16B2446
                    78566285B7645670CBAC108898BBA297B496A8BBA649F2A94C3E5F9DE91D6F9D
                    F184AFCEC4FBD4A65B3D6A135D6860E9A3A6569AD5359705F8FF0100B9934224
                    A87C8EDD0000000049454E44AE426082}
                  Transparent = True
                  TabOrder = 1
                  OnClick = CancelarItem1Click
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
                object AdvGlowButton1: TAdvGlowButton
                  Left = 787
                  Top = 5
                  Width = 105
                  Height = 61
                  Align = alRight
                  Caption = 'Alterar Item'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F90000009124944415478DAAC99797055E519C67FADCB8C9D5A119DB13375A6
                    6AAD5AAB3658B7D62A022A8B802C097B240937E7DC7BB391B02624640F843D6C
                    49EE39F72624612BB288203B2461B145AC5B1154100A0A080985085888F2F48F
                    03814882E1C29979E6CE3D73CEBDBFF3BEDFFB7EDFF71C48DF4423655643DA06
                    302C882D8598B960DA9058065E1BE26C70DB106DFF01C38AC1B00218F606A2AD
                    03445B758E7C8730FDD518760586351CC30EC1B4C06D81C782F839105702DE12
                    882905C307C96B206B3357F0B418705839C4D8AD71DBC918D62EA22D35C8E553
                    A3EF4D9D33AC2F89B672F158F7915006F1A5370830761EB8AD5BF0968CC7B04E
                    E0B284CBBA12A82572EEABC70C1411EBBF8B98D2EB048C9B033115DD7015EF61
                    882F78B01F6B884FB87CC771974410570AE65501AB1A2B6B2BA457823B00A62F
                    EBB227BFB1BA9809D33713D382D4F590BDED4A1ED23734566625A4AE03C35F41
                    549188F6DD78B846D12C162EDF3A52564156D5953CA4DB97941380E45288F2CF
                    6D71AAA28A9B2E926B55A455C5C839901B68CC44C60A4799AB207B05C45A9944
                    CC6EFA470C4B98B6E85320BAE48B5E53455881E836D1F9F4F883078C28149EE2
                    7964BD0559AB20FD0217494B2069318C580E094B5E22CA6A3AADA65F4414898E
                    E3746F62B952166FD7CA8F0F6AFBFE1A1555EDD20323E68A6E93843B58489F88
                    2C16B18B2218F9B6C394B404185501C9F36078C94D44CD3CC690A6E06C316896
                    E83A511E7FA52E3F8E1CA99524D5D47DA7D6D13ED1736AF0902E4B44CE124981
                    BB4899E7B091FA26642C83B8F2F14414369DD6C822F1DA044D59FD5103D81BF9
                    0BF5B86B8AE20B57349CABDE7D4074CA15E145CE430535168B444CF902D29741
                    EA6220D107438B6F23BAB81E97DD74F43AE7ABBFB55192F4E591E3A263B2DA0E
                    2BD69E43B5FA60DF114DF0AFD6F2F73E9324652FDD225E4A17664990A9B69DCA
                    8E2F7E94241F90E8078F9DCC90666E183053AD3C019DABFF5E67CFD58B1E194A
                    0DAC9124454D5D2C1E8B16AD7AAA5D664543245F1C3B47B4CB10B1E541B61E4B
                    B8ED42126DC0EB07C33AD86C33EE3A51AF17AE932499C52B157521A5F7F44817
                    77F4D08BA3FDFAF0E0317D76A856EBDEFF429274EAF419DDD1275B74C90B1ED2
                    B04EE2F5DF02A6FF81AB5ED8255F834BAB24490BB7EE546DDD193DE19921EE0E
                    5579F52792A4310B2A75EF803C11E2D1D837374B922A77EE132F248A7ED384B7
                    2CB88231FDDDC1B413AE3A958515E8C151F31BD297B6A052DCD451259BFF2D49
                    3A5BFF83DA65948BFB0689AEA9E285446DFBE26B673C2EAE164F18CE98F29406
                    1145DB02D35F71D599C0B01B45B1A2EA637967BF2D49BAB9CB183D39C292243D
                    163F5B3C3F547418A95FF6C9D1E9B3F592A40EE9A5E2498F302DE12EBDB6BE68
                    D8956006AA70155F1D30A250741CA779EFED6D8864CF898BC41F5D22C42D7FD5
                    C792243A8C145DC68867E2D43EB35C9274FAECF7BA3D34D349B7695F03A44F18
                    F6E7106D7DF593179B7E672AEB3B5DFB8F9F9224F599B25884B8C5EBE9E2D938
                    1DF9F63B6DD97348FCC9143D3344885B994BB74A92B67CFEB5782E4E744E71FA
                    6ACB235903D1D699165DECF68B2E13F4C8E8054E08CF9FD703DE19E28524D17E
                    845AF7CB93242595AF174F7A45D734F1E23055EE3A2049CA5AB64DB43145EF4C
                    E12A6E29E4A996035E6CDAED721455E654EAFE9A93E2D564D12945FC3946AFE5
                    39F00FC7CC14CFC689B6C3F5AB7E79FA5FFDF792A4B6191522C410FD735B0A79
                    0AA27D87AFA137390B86CEF92ADFBE479254B665A703D32343B4F168FADAF775
                    5ED2AFFBE789361E715FB81E8A2FD40F3F9CD789D36775E780F1A26DA2E893D3
                    12C8E3E02ED97AD522692A8AA1D344FF19DA5FFBAD24297CF672D1C6EBB49997
                    47E9E88953AA971458F3BEDA8FB4744FEF2C557FEAA47AED27FB7473C714D13D
                    558465370FE9F209C3BF174CFF826B5E70BAFDA273BE1E1C7DA93FFEDE3B433C
                    172F9E8BD7C4B7DED5FEA327B4F2DD4FF5756D9DBE397652AB3FD8A3DABA3392
                    A4D6C634D129598466340FE9F209C3DE02A63D2AA83D87618B97F31435A75A92
                    F4D9E1E3CE9F3E1CA5714BB76AD93F778BDF0DD6EDE1F9FAADBB40BC96AA1D5F
                    1C9224B58A98E454749FDCAB439A7619B8FD8F063757DACE787C759CE6EE70FA
                    63D9D69DA2752FCD7867BBD67CB857FC25C1497BFB91E29551DA75B0E6478039
                    CD433A535DFF0B2E81551314E485E5FFCFFACFD0BE9A3A49D2E059CBF5EEEEAF
                    1455B442FC3541F4CF13DDC78AA7BCFAD7DEC34D005E84CC14615997439EC3B4
                    7F01436D70DB139B5D6EB5683C8ED743C90B1AC6E3FE6F4E38CDFA71433C1D2B
                    FE96A8F669A53A76F2B4BE3B5BDF04E0E590D9CEA23576CE3C924A80781FC416
                    DDED547290AB60C3161D721519D8D400B9EB3F47357DF93FD4698425F3B255F7
                    E1DABA66002F283453F4CE1043A63D83771690BC085297404C99456451F08083
                    0B456881124A3769E38EDD6AEE3876F2B4EE1A32D969EE4D01F64A17E193D712
                    371762CB8094F990F67718597A3B51B3CE06ED22987E11E5135D278A575275E7
                    801CF5CB5FA0F28D1FEAC091FF36826C3D64B23303FD182E2C5BF44C1591B3EF
                    27662198E540C23C48980B498BC03BBF1F91BEEB73134CBFB31FE99E2D9E728B
                    E787EAE7BDB3149258A8CC85955AFBD197FA8DBBC059F55C01982E06172413E9
                    8741B3207C3690FA8EA3B435CE6639C6F21159787D2E81E1772A317C8AE83556
                    744B132F0D771A799731BA75C0B826523B56F4C9DDC4D065307C35242E7744A6
                    ED28C3766C87D10188F0ADBF3E2BC3E744D25D2A064D16BDD32F81F4CC10A159
                    A26F6E63C0DE63773170E64D0C9E0B6FF8E18D8023D2D75C52C65AC8D800C9AB
                    C0B0373AE6D10D886403647613559B21C2B276326C492BC66E8294B58E157751
                    A46F6EACAC6D9059E5D86FD1C515D7670CF91C484F33906159222C67157DF36E
                    256D3DE4EF80EC6AC8DE7C49CD1A98A60D71A5E02D77E12A3ED7A425126C2443
                    339D340F9C309A4153202C0F46AF849C6DD76A01CF81D8F9605AF7E30904AECB
                    CC342E54F7C049A26FCE52064C0A217C2A0C9A0A61B9C1025E30D1DD36249583
                    D77E1AD39E8E611DF9E97DED1511AFC30C5878AC764414C0C0693070320C9C72
                    0300AF780D61DD86CBEE8661CFC4B0D663D8BB88B68E5E7A0D61D560D89F63D8
                    5518960FC3EE8769B7C6E3775E434416C080292D02FCFF006A413F1C30642EC3
                    0000000049454E44AE426082}
                  Transparent = True
                  TabOrder = 2
                  Visible = False
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
              end
              object cxGrid1: TcxGrid
                Left = 0
                Top = 106
                Width = 897
                Height = 327
                Align = alClient
                BevelInner = bvNone
                BevelOuter = bvNone
                BorderStyle = cxcbsNone
                TabOrder = 1
                LookAndFeel.Kind = lfUltraFlat
                LookAndFeel.NativeStyle = False
                LookAndFeel.SkinName = ''
                object cxGrid1DBTableView1: TcxGridDBTableView
                  Navigator.Buttons.CustomButtons = <>
                  DataController.DataSource = UniDataSource1
                  DataController.Summary.DefaultGroupSummaryItems = <>
                  DataController.Summary.FooterSummaryItems = <>
                  DataController.Summary.SummaryGroups = <>
                  OptionsBehavior.ImmediateEditor = False
                  OptionsCustomize.ColumnFiltering = False
                  OptionsCustomize.ColumnGrouping = False
                  OptionsCustomize.ColumnMoving = False
                  OptionsCustomize.ColumnSorting = False
                  OptionsData.Deleting = False
                  OptionsData.Editing = False
                  OptionsData.Inserting = False
                  OptionsView.ScrollBars = ssNone
                  OptionsView.ColumnAutoWidth = True
                  OptionsView.GridLines = glVertical
                  OptionsView.GroupByBox = False
                  object cxGrid1DBTableView1ITEM: TcxGridDBColumn
                    Caption = 'Item'
                    DataBinding.FieldName = 'ITEM'
                  end
                  object cxGrid1DBTableView1COD_BARRA: TcxGridDBColumn
                    Caption = 'Cod. Barra'
                    DataBinding.FieldName = 'COD_BARRA'
                  end
                  object cxGrid1DBTableView1CODIGO: TcxGridDBColumn
                    Caption = 'C'#243'digo'
                    DataBinding.FieldName = 'CODIGO'
                  end
                  object cxGrid1DBTableView1NOME: TcxGridDBColumn
                    Caption = 'Nome'
                    DataBinding.FieldName = 'NOME'
                  end
                  object cxGrid1DBTableView1UNIDADE: TcxGridDBColumn
                    DataBinding.FieldName = 'UNIDADE'
                  end
                  object cxGrid1DBTableView1QTDE: TcxGridDBColumn
                    DataBinding.FieldName = 'QTDE'
                  end
                  object cxGrid1DBTableView1VALOR_UNITARIO: TcxGridDBColumn
                    DataBinding.FieldName = 'VALOR_UNITARIO'
                  end
                  object cxGrid1DBTableView1VALOR_DESCONTO: TcxGridDBColumn
                    DataBinding.FieldName = 'VALOR_DESCONTO'
                  end
                  object cxGrid1DBTableView1VALOR_ACRESCIMO: TcxGridDBColumn
                    DataBinding.FieldName = 'VALOR_ACRESCIMO'
                  end
                  object cxGrid1DBTableView1VALOR_TOTAL: TcxGridDBColumn
                    DataBinding.FieldName = 'VALOR_TOTAL'
                  end
                end
                object cxGrid1DBBandedTableView1: TcxGridDBBandedTableView
                  Navigator.Buttons.CustomButtons = <>
                  FilterBox.CustomizeDialog = False
                  FilterBox.Visible = fvNever
                  OnFocusedRecordChanged = cxGrid1DBBandedTableView1FocusedRecordChanged
                  DataController.DataSource = UniDataSource1
                  DataController.Summary.DefaultGroupSummaryItems = <>
                  DataController.Summary.FooterSummaryItems = <>
                  DataController.Summary.SummaryGroups = <>
                  Filtering.MRUItemsList = False
                  Filtering.ColumnAddValueItems = False
                  Filtering.ColumnMRUItemsList = False
                  OptionsBehavior.ImmediateEditor = False
                  OptionsBehavior.BandHeaderHints = False
                  OptionsCustomize.ColumnFiltering = False
                  OptionsCustomize.ColumnHidingOnGrouping = False
                  OptionsCustomize.ColumnHorzSizing = False
                  OptionsCustomize.ColumnMoving = False
                  OptionsCustomize.ColumnSorting = False
                  OptionsCustomize.BandMoving = False
                  OptionsCustomize.BandSizing = False
                  OptionsCustomize.ColumnVertSizing = False
                  OptionsCustomize.NestedBands = False
                  OptionsData.CancelOnExit = False
                  OptionsData.Deleting = False
                  OptionsData.DeletingConfirmation = False
                  OptionsData.Inserting = False
                  OptionsView.NoDataToDisplayInfoText = '     '
                  OptionsView.ScrollBars = ssNone
                  OptionsView.ShowEditButtons = gsebForFocusedRecord
                  OptionsView.ColumnAutoWidth = True
                  OptionsView.GridLineColor = clWhite
                  OptionsView.GridLines = glNone
                  OptionsView.GroupByBox = False
                  OptionsView.BandHeaders = False
                  Preview.Column = colExtra
                  Preview.Visible = True
                  Styles.Content = texto
                  Styles.ContentOdd = cxStyle1
                  Styles.Header = stTitulo
                  Bands = <
                    item
                      Width = 354
                    end
                    item
                      Width = 241
                    end>
                  object cxGrid1DBBandedTableView1ITEM: TcxGridDBBandedColumn
                    Caption = 'Item'
                    DataBinding.FieldName = 'ITEM'
                    Options.Editing = False
                    Width = 40
                    Position.BandIndex = 0
                    Position.ColIndex = 0
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1COD_BARRA: TcxGridDBBandedColumn
                    Caption = 'Codigo Barra'
                    DataBinding.FieldName = 'COD_BARRA'
                    Options.Editing = False
                    Width = 146
                    Position.BandIndex = 0
                    Position.ColIndex = 1
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1NOME: TcxGridDBBandedColumn
                    Caption = 'Descri'#231#227'o'
                    DataBinding.FieldName = 'NOME'
                    Options.Editing = False
                    Width = 230
                    Position.BandIndex = 0
                    Position.ColIndex = 2
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1UNIDADE: TcxGridDBBandedColumn
                    Caption = 'UN.'
                    DataBinding.FieldName = 'UNIDADE'
                    Options.Editing = False
                    Width = 57
                    Position.BandIndex = 1
                    Position.ColIndex = 0
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object colBotDiminuir: TcxGridDBBandedColumn
                    PropertiesClassName = 'TcxButtonEditProperties'
                    Properties.Buttons = <
                      item
                        Default = True
                        Glyph.SourceDPI = 96
                        Glyph.Data = {
                          424D360400000000000036000000280000001000000010000000010020000000
                          000000000000C40E0000C40E0000000000000000000000000000000000000000
                          00020000000A170D0738542D1894814626D193502AEA924F2AE87F4525D0522C
                          1793120905300000000900000001000000000000000000000000000000030201
                          011159311B97A96239FAC58957FFD6A36DFFDDAF75FFDDAF74FFD6A46BFFC589
                          56FFA46137F53C2112730000000F000000030000000000000002020101107442
                          26B9BC7C4DFFDDAE77FFDEB076FFDEAF75FFDEAF75FFDEB074FFDDAF75FFDEAF
                          74FFDBAB72FFBD7E4EFF6F3E24B50000001000000002000000085C362095BE80
                          53FFE0B37CFFDFB076FFDEB177FFDFB279FFE0B379FFE0B27AFFE0B279FFDFB2
                          79FFDFB277FFDEB077FFC08253FF55321D920000000A190F0932B07047FADFB2
                          7DFFDFB27AFFE0B37BFFE0B57DFFE1B67EFFE2B67FFFE2B77FFFE2B77FFFE2B6
                          7EFFE0B47CFFE0B47BFFDEB079FFB3734AFB130B072F613C2795CD9B6FFFE2B7
                          80FFE5BD89FFE7C291FFE8C393FFE8C494FFE8C594FFE8C495FFE8C495FFE8C4
                          94FFE8C393FFE5BF8CFFE1B77FFFD09C6EFF5434218B935E3DD2DCB383FFE3B7
                          81FFBA8659FFA97043FFAB7245FFAC7346FFB0794AFFAB7245FFAD7547FFB078
                          4AFFB17A4BFFC29162FFE4B983FFDEB17EFF8E5B3BD0B0744CF2E3BF8FFFE4BB
                          84FFA56B3FFFF5EEE9FFFAF6F3FFFAF7F3FFFBF7F4FFFBF7F5FFFAF7F4FFFAF7
                          F3FFFAF6F2FFAB7245FFE5BD87FFE5BE8BFFAB714CEEAE764FECE9C9A0FFE5BE
                          89FFA56B3FFFE6D9D2FFE7DBD4FFE9DED7FFEAE0D9FFEAE0DAFFEBE1DBFFEBE0
                          DBFFEEE5E1FFAA7144FFE7C08CFFEACA9DFFAE764FEE9A6A49D0E9CDACFFEAC7
                          96FFB78456FFA56B3FFFA56B3FFFA56B3FFFA56B3FFFA56B3FFFA56B3FFFA56B
                          3FFFA56B3FFFB78457FFEACA99FFEBD1ADFF996A49D46E4E3697DDBB9DFFEED3
                          A9FFEECFA2FFEED2A5FFF0D6A9FFF1D7ABFFF1D8ADFFF1D8ADFFF1D8ADFFF1D6
                          AAFFF0D5A8FFEED2A5FFEFD4A7FFE0C2A2FF6246318F1C140E2BC7946CFCF5E8
                          CCFFEFD6ABFFF1D8AEFFF2DAB0FFF3DCB3FFF3DEB4FFF3DEB4FFF3DEB4FFF3DC
                          B2FFF1DBB0FFF1D8ADFFF7EACDFFC69470FA1A120D2E000000036F523C92D7B0
                          8CFFF8EFD3FFF3E0B9FFF3DFB7FFF4E1B9FFF5E3BBFFF5E2BBFFF5E2BBFFF4E1
                          B9FFF4E2BDFFFAF1D5FFD9B390FF664B368C000000060000000102020107906C
                          4EB8D9B38FFFF7EDD3FFF8EED0FFF7EBC9FFF6E8C4FFF6E8C5FFF7ECCAFFF8EE
                          D0FFF4E8CDFFD7AF8BFF88664AB30202010B0000000100000000000000010202
                          010770543F8FCFA078FCE2C4A2FFEBD7B8FFF4E9CDFFF4EACEFFECD8B9FFE3C5
                          A3FFC59973F24C392A6700000006000000010000000000000000000000000000
                          0001000000022019122C6C543E89A47E5FCCC59770F1C19570EEA47E60CD6C54
                          3F8B16110D2200000003000000010000000000000000}
                        Kind = bkGlyph
                      end>
                    Properties.ViewStyle = vsButtonsAutoWidth
                    Properties.OnButtonClick = colBotDiminuirPropertiesButtonClick
                    VisibleForEditForm = bTrue
                    Position.BandIndex = 1
                    Position.ColIndex = 1
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1QTDE: TcxGridDBBandedColumn
                    Caption = 'Qtde'
                    DataBinding.FieldName = 'QTDE'
                    Options.Editing = False
                    Width = 90
                    Position.BandIndex = 1
                    Position.ColIndex = 2
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object colBotAlmuntar: TcxGridDBBandedColumn
                    PropertiesClassName = 'TcxButtonEditProperties'
                    Properties.Buttons = <
                      item
                        Default = True
                        Glyph.SourceDPI = 96
                        Glyph.Data = {
                          424D360400000000000036000000280000001000000010000000010020000000
                          000000000000C40E0000C40E0000000000000000000000000000000000000000
                          00020000000A170D0738542D1894814626D193502AEA924F2AE87F4525D0522C
                          1793120905300000000900000001000000000000000000000000000000030201
                          011159311B97A96239FAC58957FFD6A36DFFDDAF75FFDDAF74FFD6A46BFFC589
                          56FFA46137F53C2112730000000F000000030000000000000002020101107442
                          26B9BC7C4DFFDDAE77FFDEB076FFE2B782FFE3BB87FFE3BC86FFE1B782FFDEAF
                          74FFDBAB72FFBD7E4EFF6F3E24B50000001000000002000000085C362095BE80
                          53FFE0B37CFFDFB076FFDEB177FFB78254FFAA7144FFAB7245FFBC8859FFDFB2
                          79FFDFB277FFDEB077FFC08253FF55321D920000000A190F0932B07047FADFB2
                          7DFFDFB27AFFE0B37BFFE0B57DFFA56B3FFFF5EFEAFFF8F3EEFFAB7245FFE2B6
                          7EFFE0B47CFFE0B47BFFDEB079FFB3734AFB130B072F613C2795CD9B6FFFE2B7
                          80FFE5BD89FFE7C291FFE8C393FFA56B3FFFF1E6DEFFF9F5F1FFAA7144FFE8C4
                          94FFE8C393FFE5BF8CFFE1B77FFFD09C6EFF5434218B935E3DD2DCB383FFE3B7
                          81FFBA8659FFA97043FFAB7245FFAC7346FFF5EDE6FFFAF6F3FFAD7547FFB078
                          4AFFB17A4BFFC29162FFE4B983FFDEB17EFF8E5B3BD0B0744CF2E3BF8FFFE4BB
                          84FFA56B3FFFF3EBE6FFFAF6F3FFF6EFE8FFF7F0EAFFFBF7F5FFFAF7F4FFFAF7
                          F3FFFAF6F2FFAB7245FFE5BD87FFE5BE8BFFAB714CEEAE764FECE9C9A0FFE5BE
                          89FFA56B3FFFE0D2CAFFE1D3CCFFE3D5CFFFF2EAE4FFF8F3EFFFEADFD9FFE6DA
                          D4FFE9DED9FFAA7144FFE7C08CFFEACA9DFFAE764FEE9A6A49D0E9CDACFFEAC7
                          96FFB78456FFA56B3FFFA56B3FFFA56B3FFFF1EAE5FFFAF6F3FFA56B3FFFA56B
                          3FFFA56B3FFFB78457FFEACA99FFEBD1ADFF996A49D46E4E3697DDBB9DFFEED3
                          A9FFEECFA2FFEED2A5FFF0D6A9FFA56B3FFFF0EAE7FFFDFCFBFFA56B3FFFF1D6
                          AAFFF0D5A8FFEED2A5FFEFD4A7FFE0C2A2FF6246318F1C140E2BC7946CFCF5E8
                          CCFFEFD6ABFFF1D8AEFFF2DAB0FFA56B3FFFDECFC9FFDFD1CBFFA56B3FFFF3DC
                          B2FFF1DBB0FFF1D8ADFFF7EACDFFC69470FA1A120D2E000000036F523C92D7B0
                          8CFFF8EFD3FFF3E0B9FFF3DFB7FFB98A5FFFA56B3FFFA56B3FFFBA8A5FFFF4E1
                          B9FFF4E2BDFFFAF1D5FFD9B390FF664B368C000000060000000102020107906C
                          4EB8D9B38FFFF7EDD3FFF8EED0FFF7EBC9FFF6E8C4FFF6E8C5FFF7ECCAFFF8EE
                          D0FFF4E8CDFFD7AF8BFF88664AB30202010B0000000100000000000000010202
                          010770543F8FCFA078FCE2C4A2FFEBD7B8FFF4E9CDFFF4EACEFFECD8B9FFE3C5
                          A3FFC59973F24C392A6700000006000000010000000000000000000000000000
                          0001000000022019122C6C543E89A47E5FCCC59770F1C19570EEA47E60CD6C54
                          3F8B16110D2200000003000000010000000000000000}
                        Kind = bkGlyph
                      end>
                    Properties.ViewStyle = vsButtonsOnly
                    Properties.OnButtonClick = colBotAlmuntarPropertiesButtonClick
                    Position.BandIndex = 1
                    Position.ColIndex = 3
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1VALOR_UNITARIO: TcxGridDBBandedColumn
                    Caption = 'Unit.'
                    DataBinding.FieldName = 'VALOR_UNITARIO'
                    Options.Editing = False
                    Width = 90
                    Position.BandIndex = 1
                    Position.ColIndex = 4
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1VALOR_DESCONTO: TcxGridDBBandedColumn
                    Caption = 'Vlr Desconto'
                    DataBinding.FieldName = 'VALOR_DESCONTO'
                    Options.Editing = False
                    Width = 79
                    Position.BandIndex = 0
                    Position.ColIndex = 1
                    Position.RowIndex = 1
                  end
                  object cxGrid1DBBandedTableView1VALOR_ACRESCIMO: TcxGridDBBandedColumn
                    Caption = 'Vlr Acr'#233'scimo'
                    DataBinding.FieldName = 'VALOR_ACRESCIMO'
                    Options.Editing = False
                    Width = 86
                    Position.BandIndex = 0
                    Position.ColIndex = 0
                    Position.RowIndex = 1
                  end
                  object cxGrid1DBBandedTableView1VALOR_TOTAL: TcxGridDBBandedColumn
                    Caption = 'Total'
                    DataBinding.FieldName = 'VALOR_TOTAL'
                    Options.Editing = False
                    Width = 101
                    Position.BandIndex = 1
                    Position.ColIndex = 5
                    Position.LineCount = 2
                    Position.RowIndex = 0
                  end
                  object cxGrid1DBBandedTableView1CANCELADO: TcxGridDBBandedColumn
                    DataBinding.FieldName = 'CANCELADO'
                    PropertiesClassName = 'TcxImageComboBoxProperties'
                    Properties.Images = cxImageList1
                    Properties.Items = <
                      item
                        ImageIndex = 0
                        Value = 0
                      end
                      item
                        ImageIndex = 1
                        Value = 1
                      end>
                    Options.Editing = False
                    Options.ShowEditButtons = isebNever
                    Width = 35
                    Position.BandIndex = 1
                    Position.ColIndex = 6
                    Position.LineCount = 2
                    Position.RowIndex = 0
                    IsCaptionAssigned = True
                  end
                  object colExtra: TcxGridDBBandedColumn
                    DataBinding.FieldName = 'EXTRA'
                    PropertiesClassName = 'TcxMemoProperties'
                    Options.Editing = False
                    Styles.Content = stExtra
                    Position.BandIndex = 0
                    Position.ColIndex = 3
                    Position.RowIndex = 0
                  end
                end
                object cxGrid1Level1: TcxGridLevel
                  GridView = cxGrid1DBBandedTableView1
                end
              end
              object pnBotoes: TPanel
                Left = 0
                Top = 0
                Width = 897
                Height = 106
                Align = alTop
                BevelOuter = bvNone
                BorderWidth = 5
                Color = 8404992
                ParentBackground = False
                TabOrder = 2
                Visible = False
                object AdvGlowButton2: TAdvGlowButton
                  Left = 567
                  Top = 5
                  Width = 101
                  Height = 96
                  Align = alRight
                  Caption = 'Mesas'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F900000096E4944415478DACC99F96F13671AC71FA9EA4A5D69A53DFA076CF7
                    B7ED2F5BDADD562BADD4B2057295232D6C43EC192721A5C742D9B6A894426939
                    122740B973D833764249295B5AA01CA150A0DCD5B2E5889BD89E7967ECD82697
                    67C64E9C8318DBF9EE0F26E4A694A65D2C7D25EBD56BCFC7DFE779DE79E631BD
                    FF3568B8D69C05BD7772803841A7C2DA282DFC384ABC6850D1CE1E328B11E2C5
                    4EB2081A71A2FE475ED05EE305CDC98BFA494ED0029CA0C56EA9C5E2D0CFF2A2
                    56C70BDA52CE613CC68B3A71A24166C1A0A2DA4E2AACE926DEA19159D0C82CE8
                    F4EEB104AD3D3792E5FDAF4174B7800B3EEA254E8CFC96138DE5BCA0793841C3
                    A0CCF6A1F7436B619847ACE92A276A2566C1F8FD829D5D5454DB3339802FEFEE
                    244E301E2C7046CB3841EFE484F01898BB57189CA0252C0EA3DAE28CFDCEE2D4
                    7F3CE04B75D199265B5831DB7F0CD8B8EE46788756C089FA9D01579FC108955C
                    00BD7F1A64711A64B269AB270B6A22474D766D8759D069E58914955EC4181E7A
                    EF2446E883D3A0155FA5C8E2D0EBF2AB3BC6CDAFC994C91686D91EFEEA9DA309
                    5A7D066378C82CE2B63827A8B0364E058E8EBA9FD6B9B1E285F6D3853B6F10E7
                    C408265A7A18B4F43068D951D05B874166A173755E651B7E6EC0F9556DC8B745
                    76BFF1459A65908B5EDF17A3C59FC7E88D83DDB4645FD7339CF0D38775FCC2D1
                    60B67560F1DECE82370F75D3E2CFBBE8F57D31A2E2BA382DDC1D274B4DDF032F
                    EE68D7B849ACD67B299A172BDAC03B7B1F5EB83B4EC57571A2459F2569C98124
                    15EF8A95E557B5FF1FE1D2CAAF6E47F1AECE3D4B0E2469D1E749A27C7B8CF26D
                    B1874C362DC18B23DDB3881A66ED8820675B04F36D3AF871935B83C9AE21B7C2
                    40D6D608B2B64631A7C240BE5D9B707FBE5DC3ACED0672B647C18B23BF971735
                    986C1DC8B7753D9A6F8F11591CDDC48B91E59CD0313227040D26BB8E75FB5AB0
                    F2D356CCAD3230634B1439DB22C8ABD6D317D911C18C2D51CCDA11C12BB5616C
                    3F1244D5D10016EDEA406E8581E95BA2786E7BFAC7CDB7E9786E7B7AFFF31506
                    DEDADD8E0D0702287686611223A36E8B61F062A4AAC0D14D647118C40B7A68CC
                    F964D730A7C2C0C9FF30B4FB5C38744E42D5D100567EDA8ABC6A1DB91506DEFC
                    B81DDB8E84B0F7B4826FFEEB82EA71C1EF75E1D2B70DD87F5646457D10CBF6B4
                    616E958179550696ED6943457D1007CE32F824375CAE46CCD9A8E0850A1D8535
                    91D14E7715388C078977187F18F700BD15B6CF4EAB08303714C903A5E92A5CD7
                    AEE2CB8B320E9E57F0ED1517E4C6CB50BC8D507DCD7079FC686852E1F307A178
                    9BC01A2FE3EA956B387281A1FE02C3D5ABD7C01AAF4095DC506509672E799155
                    EAC5D4750A784143A17314A4C3984505A2BEE44E80FF3EC51054257475DF40BB
                    D609590DC2EB76437237C22BFB1168D110EDEA0500F8431D907D21004067771F
                    42AD3A24D60CAFBB095E771324D68CD68E08BABA6F4055549CFCA60933D733FC
                    ED032FB2CB1570828E826190165117C8E230EAC63B5A860336332FFAFA13187C
                    29CDAD70797C480E00A901209E482139905EF7B0009229A0FF6612A981F47EB7
                    DC0C370BDCFE7CFFCD2498CC6E033EB3966146891759E50CBCA0A3D01981D91E
                    062FEAA7A9C0699C31DBBE1FB0BBB71F37E2092407005FB00D8D523312A901F4
                    F4C5D1D3174722350498480EAD270700AF1284570D213500DCE84FA0A72F3E06
                    30AB2C0D995DCEC08BB79C1475993841BBFE7D211E0E9848016AA015DF79FD88
                    2752DF0B9848011E168047092299BA33605619C3F492A17017D64474E204ADEF
                    7E02CC2E6398BE6E1052EBB9EF0007950EB7D24366BBD6763F0266972BC82895
                    225458635C30FD8022F9B90033AD329EDBA0AA6471187BCCF6FB133067BD7A9E
                    2C0E7DD9FD18E24C2B43CE7AF523E29D91472703303900A881367895F4413D19
                    0ECEFAB0398F2CA24E9CA0E91301EE39C5105024DC4C0E2099C22D9021C04168
                    200DE8B975C748A6801BF1E418C0447220ED365370E2E2C4805965ECE6CC8DFE
                    5F92458C112F46368C6EB706010F9F6708B226045B34C47A6E0000022D6134C9
                    69909EBE38DAC25128FE16C89204D9EB05F385D0D21E41ACA71F0020FB4290FD
                    D70100DDBD7184DA0CC85E2F2E5D7563F60686A7D78C84CB2891915DCE76E76E
                    F511E5DBA334BF3AFAB0C916062F8E6EB722D8795C4588B9E0975C50250F02C1
                    EB68949A71AD494520D40695C9F04BDFC1FD5D038E5FF4E0CB0B5E34BA1AE0F7
                    BAA0CA1282A15634B87D70797C08865AA1CA5EF8BD2E5C57DD38725E42A695E1
                    EFEB860196334C5F2761F6A6C093F32A4344FFDCDB4F8BF7C5A9F8A3A860AA1E
                    D9F29B450379551D78A95AC6E67D1EEC3FEDC1A56F5D08B2265C57DD68B8D680
                    FA736E38EA252CAD9591616578B68461B183A1FAB0844367DCB87CA501D75537
                    424A132E5F71E1E01937AA0E4B5824CAC82C6398563AD2BD67D778307747E8F8
                    9B077B69C9FE5EA2973F49D1AB9FA6A8B0B6FF5779151DF1E1B3178BA8C12C46
                    31E3C330FEB45CC65F5749E0B6336CDA27A3F29084576D329E2D61787C455A53
                    D7A6DD786225C394150C4FAF655850C9B0ED0B195B0FC828AA4C8773CA0A054F
                    AC6498BA8E21BB7C24E0D40FDC985FDDFCC8CB7B5AA8E8A3EB440B7747E8A58F
                    23F4DADE4E7AE51323CF6C6F1FF3D859541381D916C6B412194FAE94F0F80A05
                    535628F8CB7B69C7C649F0F48DBF94E1A95569A0292B143CB52ABD36D1FE19EB
                    24CCDDDEBADC2C86695E452BFDA3B295E8EDFA04BD5D9FA0E5C712B4F4708ACC
                    42D43E7F9CA7BBC29A084CB63032AD32A6977827BCC8BD6ADA1A2FB2CAE453FF
                    DA1FA36547FBE8AD83698D197DF0CE7E32DBDB4E8C773616D54461B285916195
                    D30DE664015A19A6AD95DDF32A7D0F2CD8E523DEA9DC162D3F86211D4F8FDEDE
                    397A932CA27E2ABFBA632C646D14265B0732AD326694483F1A6E4689848C52B9
                    69C9FEDE5FAF3A7593DE3D1EA7E5C786446BCF6184064760B7C66F759C7D6227
                    33AD32324AEF1D32D3CA9061958FE6AC67BF58792245EB2FA5E793C379261E60
                    8A3A15D444A97857B4D8640BDF1CDD500CE6E460B84757E39D94512A23A354C6
                    EC8DFE777237FB29AB5CA6B7EBE3E9D9E40F99B016EDBC3D027EC4E28C38878D
                    71C714CEDDE464A635AD991BD5FD7336A98F3DBFA5999EDFECA7CCB27B041C67
                    88FE674E34B67182DE3E12B2634CB833ADF2A8069475655A2561F626FFD4F955
                    2DF4C2561FCDD9E4A7DC4D930438EA6F888738519FC90BDA0E4ED04E14D5463D
                    265B4738AB8CC532AD722CAB8CE939EB1539BB5C399355C66CD9654A5ECE7AF6
                    9BEC728972B70428AFB2857237AB7705F8BF0100675BE3817DFA3C1900000000
                    49454E44AE426082}
                  Transparent = True
                  TabOrder = 0
                  OnClick = menu_mesaClick
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
                object AdvGlowButton3: TAdvGlowButton
                  Left = 668
                  Top = 5
                  Width = 110
                  Height = 96
                  Align = alRight
                  Caption = 'Delivery'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F9000000A324944415478DABC997B5053671AC6DF994E77A63BB3B39DEDCEEC
                    9FBBEDAC33ADDBB1B45BEB85AEA2429213445D54E492734290626B55EA6514AC
                    778170913B08C939818078E9CD4BEB686DBD94B676BBDBB5ED766B9180889742
                    92734E0249B824019EFD23880226406B3733CF9CCC9799737EE779DFEFFDBEEF
                    0DEDBE04BA5FFB3E05EDBC30442C2F91AED64169879DC40A32A5D47B88159CC4
                    09DDC4F122B182F40CC74BAF73BC58C309D20596176FB1BCE81A5687D6247DCA
                    096203C78B5B38410EE30489585E220D2F93AEAE9BB4B52ED29965D2D54AA431
                    889479CE4FFB3F1BCDB2FB1288260BB8FA500FB182F309569033595E6A627911
                    77A531DEFB7E6FCC0ECDE8B1369617B335BCFCA7947A17259BDD0F0770CD1127
                    6904F9D1E41A672ECB4B5D2C6F1F0733251945BFD6245773A6EE2752EA1C3F13
                    D0ECA4571A9C3149067BABC6F833C1C6BA6B141D3AB394ACAB954303EE6DC428
                    655F06EDFE04945C23539241DC37FCD60F0D6E3CAC58C9F222ED383F48395F60
                    1C0FEDBC8051DAF309E8CD8F07496B921A12AB6D0FCCAF87A924831D1AA3FDE3
                    8CB37EDADB88713CA4113022B606A4337B29D9643BFC4B423D485ADEDAA833F7
                    115B83514CB4E53468CB69D0B6B3A0CDA7411ABE6B5FFCC14EFCBF0113AA3A91
                    58ED38B2F15480E52E176D38EEA2F5EFB968E3FB6E4A3FDE1DC1F2BF7C5883E4
                    2234061BD6BFD395BCE90337AD7FAF9B361C7711A5367829ED8897B4B5BD8FAC
                    AAB08A6C90D9CA09816B74890D4BCBEC483488E07E0E94208315A431E376ACAA
                    EC0457D3F344DA512FA5367889D6BD3B40E9270728F5902B37B1CA1AF2A68906
                    115BDF96B1B641464CA91DF3F3AC5016D91057ED8086774EB20EDAA0ADEB43F2
                    51406BF68035DA463FA3DA8AD4435DC7D24F0ED0BAF70688128D2E4A34B81E4B
                    32887E4E085EEB9657DAC109225C1E2FBC3E1FBE69EFC5A1CB2EEC38EE409CC1
                    85A872607E29A02E91115F651DEFAED106AEC605DDDB00CB8B58B6BE082BB32E
                    417778605CA4920C36241ABAFE92687411694D6EE2044726CBDB42BEF9A2021B
                    B23E7002F003F00D5F070000B79B8EE1CBB369A83D790C9B8EDAB0BCAA1B1179
                    562CC8B76249A90D1A5EC2EAB780E48601C4ED3D83E8F81444853F09A56E17E2
                    843EA4D4B9C6859A131C55C9829B486B9289E3A5DB13856641BE1527AFB8010C
                    C0D3EB0DA86F087D43006E6C01FEFD5BA03912035E179A6CC0A9AFDDA8B8D085
                    B50D0E24997D4828FC1A31AB33A158F83C544B62A0DE580BD5BE6FA1C869C5AA
                    839D48A91D9722DD5A93FC287126F9A989E0E2ABED585E69474B671F00FF08A0
                    DB0FF476DD06BE8C001AA761F0DA2E78010083C3EE0E600840D1A10B78765638
                    94AA08A8D3B2C0ECBF02A6BC0F4C612794D94D88CABA8655951D48313B47B9C8
                    9AE425942C48E913012E2DB7E3B57A09BD7D5E0C0DFAEE031C429FFC1D06BFD1
                    009F87C1D75E07D710467EF7F9070100993B7762467804D4DB4E04C08A6530B9
                    9661B54099634154F63064AD139C2081358AD00A124F5A93DCC04EB011585860
                    43E1875D00FCE8EDF3DE0BF170983DFD40AFB3153D6E199EFE7B0E0340FB6D2B
                    16C5C46261D21630250E3007EE80D107C0EE9732A7198AECD14E7282F40925D7
                    C88D1A4368C0887C2B3EFCCE036000EE9E3180BD3E78FA01B71F707B014F4F7F
                    C0DD9E7E00C0A52FAE206CCE6C30EBABC194BA465C7B90EE3A1957D9015DAD13
                    9C49B210CB8B7742C1C5558988AF1671D33E3AFF26549F0F0050CAD763FAACBF
                    81D9F13198126750B8B190C3E19688E5C5DE50804CB10DDBDF7500437EF8BDBE
                    49030E0C0100A05BB7193355F16072AE8239F0E38480F7873BAEA2C3131290E3
                    4544E459517921907F3D93756F38FFAE5A6E609E2A1A91DA37C1147680C96B9B
                    142093DB0295DE02658EC5431AA3D8197401E745280A6D686CEA995A788701DF
                    3A750E4FBFF002A237D583297387CCBFB152E7B540A5B73848572B5F4E0A3249
                    622B6C483177C1D58F2901F6F4FB010079653CA6CF9E0FF5AEC6E1D2D2322507
                    17175CBF4E5A937C2CD879435964C7FEF71D18FBE9F70D068573F704DCF30F0E
                    2161F5AB98B32C0D4C7E1B98FC9B53068CCE6FFD9CB482B82D68881B8690F066
                    3D525FE150653E86C67F5C41FB1DEB38E03E6FA0FCB83CFD2363FFFCFA7BCC9C
                    371FAA0D4630953E30FAE6290132FA16A8F35BEB8933C9D383EDD7561F0516A7
                    F398313702D36785E385F07028635761E38E6CD41C3D8EFFFCD00267B76714AC
                    A3CB8DACDC624446AF44F8BC0550B119600AEE405D6C9F12A44A6F41CC811B09
                    A43549C4F2A2F4C073824946122F419D73158A6DA7A15A578188F87484CD53E0
                    991767E1E5480596B3AB91917500A7CF7F86FFFED00C6DEA7A3CFE873FE38FD3
                    C210363302F3E7BE04E5CA35600EDC0EAC229377D1B7B8E0C6AF492BB888131C
                    050FDA6E698C36E8CCDD60CD7D604A9C882A7480C9BF01F59ECB50679C42545A
                    015EFE7B2A66842F44D8DCB9981D1983E7672D42A47A39CA0D7528AA302142B1
                    14F3673D07E5EB3CD4E59E49D6410B985CCB91A5C52D4489062725543B7F9F64
                    B08F6CEBC7961A5DAD032C6F07936B4194BE0D4CB11858B68A653079AD50EFBC
                    0466F3612C58A2C1D3CFCCC0F63D792321DF94B10FD3A73D0545D276A84B9D60
                    F2AE8706CC6B415456339614B5BFB4A2F226D1EBEFF4D3FAE35E4AAD77F249D5
                    D6A0F530002982C96D4154F6B5E11BB68229B805A6C401B5002852F2F0D767A7
                    814DDD806FBE6FC557DF5E4382762D66CE781A8AE46CA8CBDD60F25A43022EDA
                    D7841515B73FDAF4412FA59FE8255A7374905E7B7B9074E6FEDFC457DABCC17A
                    2FE320B3AE8D2EAC65DD506DBF0845C45CCC99130E654C3C1431AB3067F65C28
                    16CE876AE7A750973A260CEF823D3F20A1BAFDC935C77EA494FA3B4469471CF4
                    CA6107AD7DA78B5E3D2AC76B8CD6A0C7CEE04E06DC54973AA17AE30822174562
                    CE8BCF61CE8B61888A5440B5F95DA8CBBA27848BCA6AC68AF28E4C8D60A79595
                    1D1477B08368EB193F6D3DE3A7CC737EDA727A9034BCD39810E274171452DF1C
                    984065DD60767F09E5863A2837D483D9FB15D465AEC03AACB78408ED3530B996
                    8B6F9C70D1B6B3BDB4F9FD80C6B53EB89A7ED2183BCF873C6407850C00A88B45
                    A82B3C505778C014072657283895BE0591FB2D4D7107DB1E597DA88DB89AD611
                    51E639DCD34781D65BC6591F6905E96262B56DEA90539422BB19CA1CCBD5F413
                    3D8FEFBAE8A3ED1F7929F3DC3DD1FECF304A775B60DA40FBAD2154EB6DA28933
                    61BDD35BA0D2379F8DCE6FFDD58EF38394FFAF407FF27E9EE00D4C41A2E45A27
                    A51E72A62619ECBE601B8A9FE2A432C70265B6054B0ADB33624B6E50745E0B6D
                    3DE30DF426A7D2614DA973D29A235DA4E1E527B535CE9A91E3E04F8454E92D50
                    E95BB0E440DB8965C5D7C362CB6ED1F2D29BA4CEB5FC34409DD939DC44770C37
                    D11D3359412EE378C93AD970ABF416A802802E26D7C22F2DBAB920D1D0412BCA
                    DB6859C94D8A2D697F188077FF86701027749196171F6305298613A40A8E17CF
                    7382D4C4F2A25DC38B2E5DADC3A5318AD2E282EB1626B7A5519DD76A8C2E688B
                    8F3970FD773105AD145B7A8B12AA7EA4D8D2EB9302FCDF000F74FB11C9776D95
                    0000000049454E44AE426082}
                  Transparent = True
                  TabOrder = 1
                  OnClick = mnDeliveryClick
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
                object AdvGlowButton4: TAdvGlowButton
                  Left = 778
                  Top = 5
                  Width = 114
                  Height = 96
                  Align = alRight
                  Caption = 'Pr'#233'-Venda'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  FocusType = ftHot
                  NotesFont.Charset = DEFAULT_CHARSET
                  NotesFont.Color = clWindowText
                  NotesFont.Height = -11
                  NotesFont.Name = 'Tahoma'
                  NotesFont.Style = []
                  ParentFont = False
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                    6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                    7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                    1F90000008D74944415478DACC99EB4F5BF719C71FA9EAA44E9A34B5FB03D649
                    7BD14DDAFA667B31A959D7CB26ADEAB4344D16DA2EA42BED7A53DB2D2460734B
                    42BBA429D8E7E20BE616DAD2827D7C6C3090723117139A7669B4B6215C7C6C02
                    24C197738E0D3884601CE7BB1760C0C124A8A46D2C7DE4A3A3E39F3E7A9EDFF3
                    9C731E536F2F682D9E93A01E7792EC9C424E73949A2A2324722AB556CD91838D
                    5293699AECBC4276567E4060E5D7044EAEB5F34AB7C0CA1704568E09AC1C1318
                    39201A947E8193EB0556CE1539F9413B27936850486422D45C3143AE8A597218
                    54729854121899BA3E49906720DDA5B717449B156CAB99272717BD57E4558D9D
                    53460456460A1BB37ABC724E1F86B0E6BC9D93CFDB39F91D9189FCD46599A516
                    4BEC36085644A9A92A4222ABDEED344D1FB5B3CA8C8D09AF93D92C364686C0C8
                    09D1A056349B66EF731AB728D8648E92B332FAA44D17F6DBF4A16F2CB65E340C
                    1B234745A3B2D7618EDC42D08334FA4F81FAFA40A25125412F1F4E2D78BBE456
                    24D9706A5D83C0C9D4DD95A4FECFB0CE877ABB9186C703EAEE4A929D57EAADE5
                    A1B47DF46D60D3876163C25D5D2716C9D38F753ED45B8325AA41278F833A6BE3
                    64E5421F7D9B52996864839EF69AAB3450B7E492F2A2BE36505F1BC8D30EEA6F
                    033519A60F351C0BE0BB166C280BC0C1443EF6B8409E0EAC7891AB2E46AEE3B3
                    D45A7F999A3F9C7DD8CAAC4FAB9D53E1E0A3707091DB031F85C8A9E9928C0CAB
                    3E0867DDF4DED6FACBE43A3E4BAEBA1851C7F138757D18A713E6F9BB1A8F0615
                    1B1386C02D875D1F44AB691E9D96EB6831CEA1D5347F5B6831CEA1C392449B79
                    015626B8A67064341C0DA0CD74E5BEAE0FE3D4713C4EE4B65EA35E47925C35B3
                    471BCBD353DB6ABA82F7CBFCC879F91876FE5D83DDD925C8DABB3576679760D7
                    1E2DF6BC588A9A632368B72C426057BB446379104DD5338D3D8E6BE4B65E236A
                    6162E4D2C7EE117472C2C685D7A455415725F0D2AB3A3CF0A7DF61DBD33BB1ED
                    E99D7868C7D6D8B66327B63DBD0BBF7AE2616C7F2E170E66162D15975703C3C9
                    B0EA4268D2CDFCB2858D11B5555C26A721AAB132E98D58E454B49916F0B73DC5
                    78EC9967D0D9771AA7BE1844CFC0992DF1E9E941749FFC02DB735EC1434FEE81
                    EEC020446616CDE6E89A461E82C31031B755CC11390C11B273CAC51BAB4AE454
                    9C30C791B5E710FEF8DC73181C99C0E42515A3BE8B5B62E2A28261E9029E7AF1
                    553CF2D71CE80F0CC1ACBD80465D084DA6946418765699751822779368507F96
                    A9EC5382CF641FC6E3CF3E8BCFCE0CE1DCE804BEF86A744B0C8E8CE3F49723D8
                    9EF30A1EDDFE2278ED088CF993E0F2A565C9C8CA9D46E495BF909D53DEFC3E05
                    39CD082A8B2E82CDF382CD4B97B4734A15899C5A9FE95EFB5D0A5A0A2FC2A0F5
                    81CDF382CBF7AE48DA59A58F4483EAB1E9EF0C4163C1AA64437910AE8A698904
                    56BEF47DA73825982E29A1A13CA892C0CAF37792E05A495EE39BBB23055392BC
                    469A2381918377A2E03251128DEAA7372B9254A3FEF29C1F6313219C1B1DDF12
                    FEF1200647CEAF34EA8D0479AD0453C1D81889BCDAB8993633707A105F0F8DE1
                    F3FF0D6F89AFCEF9F1D999A15B4690D74A3069FD03646795BC8D52DC6E4960D7
                    B325F8F9B6DFE391A7F6E2D11DCFE3912DB2B4C65EFCFAF127F0D013D960F387
                    51597429638A8D05FE0FC8C1477EB191608B711EDAC25A64BDFE3AB25ECEC5EE
                    7FEE43D6CBB95B22B546F65BFBF1DA812330148CC2AC9DCCBC070BC7B248E455
                    1258595DFF3223E3837742084FC50100F144128B1B104F24115F5CFE5E737CB3
                    EB536B9A4BFC28FF973793E0A2A9C8FF436A35C7C86188BE77E3E396C0CA387E
                    3088C153F3D8CC2799CC7C7CB3CFF099391C7D75145C9E94BEFF343E18B4BE8F
                    2D457EA226DD3439CAA67F62D3AD3EEAA7B0EA6458CB15B805055DF600BAC420
                    DC62086E3108B71844973D88131F05E0F7467179E13294E82C94E82C62576398
                    9C98418735884E21B8727D8ADEE630DA1B03607225E8F779612C58B3EFB47EB0
                    5A2F0C45DEDF9A8BBC44EDF571EA6A8C93B3325AD5581E4C7F59E265088C82BA
                    4321BCF7BA0FC5D96771F88521BCFDD2300E3D3F8482AC41BCF5E7AFE1E9B888
                    582282C9A93026A7C298BEAAE0CCE753C8DB7116F9BBCEE2E0DEA5DF94E62C71
                    F8852194E60C83D92FC154E44F8B1EA3194575C1854EE74195C4832A91BB2149
                    BD4292DAAB177E643D168ADFD8726C8C0C9157D0648AA2A67412FAFDA3E03512
                    8C053E980AFC287B53C2995341C416A73115523115523173358AA1B361B00756
                    EF0A068D6FA366BCA6B5F8A0DF3F843ACDC4FD4241881AB54122A7314A4E4394
                    9A2B6748AC8CEC6E648219A7094B9211541F9E4071F65974D80200AE637E218E
                    04120012B8BE0C90C035247065611100F069878CC3FF18BAA52097EF45F5D129
                    CDFB7A99AACB025453162072BB12E47625A8BB2D413DAE2439F96865635930E3
                    CB754AB2FC0D3F9ADFBF80F84202B3B105C8E179042ECD2338B544E0D23CC2C1
                    794CCF5CC5E242129EB6304A736E2EC8E48D82D77A7B1CF5316A6D9AA7666109
                    F2D46285813A5047D50235E802EE8DC66722AF2E477212E5FF1E01972781CBF7
                    81CB973684DF447A59AD34525972FEAEBA92F3545BE45F817A3BB0425F27A8AF
                    FB3A759E58243BAFF458CB338FDC1C06150E3E0273E1388EBD310CDD3E2FF4B9
                    12F4B9DE75E8F679C1E54969959A96568D17BC561A72955CF9714FC92275162F
                    5247717C05EA3D8955FA41FD9F833CFDCBE3379D5C9F693F662E9C5B47290D8D
                    0F7CBE04AE48FAC450ECFF81BB3849FF2D01F597804EAE61E30126BF3CC0B444
                    726CBAF062C6279E3585A3DF3FBA69392E5F029727C1523C9E6F3932417C918F
                    DA8BE2345002EABD819B8F802D516AAE8E92C8AAF73B4DD1DA4CC3CC94642A92
                    376D231A1F788D0473F179A7E5EDB107AB4A27C9726482B842E91B0A562C0FD1
                    7995DA6AAE90838DFE46E415CECE29A1F5E94E97E435CBC5B19AFA9851235599
                    8AC7FF50FB6E80AADF3D4F9587C6A9E23FE3B74170EDDF10E669B2F3CA3D7656
                    7ED2BE34BE750B9C3C22F28AEC34466235A593312ECFAB9A0AFD92A9D0EF3116
                    FA2BCD85FEDDE6C2B17BCD0512551C9CA49A23535475646C5382FF1F00B0F638
                    68A0CC0EE30000000049454E44AE426082}
                  Transparent = True
                  TabOrder = 2
                  OnClick = menu_PrevendaClick
                  Appearance.BorderColor = 14727579
                  Appearance.BorderColorHot = 10079963
                  Appearance.BorderColorDown = 4548219
                  Appearance.Color = 15653832
                  Appearance.ColorTo = 16178633
                  Appearance.ColorChecked = 7915518
                  Appearance.ColorCheckedTo = 11918331
                  Appearance.ColorDisabled = 15921906
                  Appearance.ColorDisabledTo = 15921906
                  Appearance.ColorDown = 7778289
                  Appearance.ColorDownTo = 4296947
                  Appearance.ColorHot = 15465983
                  Appearance.ColorHotTo = 11332863
                  Appearance.ColorMirror = 15586496
                  Appearance.ColorMirrorTo = 16245200
                  Appearance.ColorMirrorHot = 5888767
                  Appearance.ColorMirrorHotTo = 10807807
                  Appearance.ColorMirrorDown = 946929
                  Appearance.ColorMirrorDownTo = 5021693
                  Appearance.ColorMirrorChecked = 10480637
                  Appearance.ColorMirrorCheckedTo = 5682430
                  Appearance.ColorMirrorDisabled = 11974326
                  Appearance.ColorMirrorDisabledTo = 15921906
                  Appearance.GradientHot = ggVertical
                  Appearance.GradientMirrorHot = ggVertical
                  Appearance.GradientDown = ggVertical
                  Appearance.GradientMirrorDown = ggVertical
                  Appearance.GradientChecked = ggVertical
                end
              end
            end
            object v_quadro_esquerdo: TRzPanel
              Left = 0
              Top = 0
              Width = 569
              Height = 504
              Align = alLeft
              BorderOuter = fsNone
              Color = 12686713
              TabOrder = 1
              Transparent = True
              object v_tipo_touch: TRzPanel
                Left = 0
                Top = 0
                Width = 569
                Height = 504
                Align = alClient
                BorderOuter = fsNone
                Color = clGreen
                TabOrder = 1
                Transparent = True
                Visible = False
                object pnTouch_Grupo: TPanel
                  Left = 0
                  Top = 0
                  Width = 569
                  Height = 106
                  Align = alTop
                  BevelOuter = bvNone
                  TabOrder = 0
                  object scGrupo: TScrollBox
                    Left = 0
                    Top = 18
                    Width = 569
                    Height = 88
                    HorzScrollBar.Smooth = True
                    HorzScrollBar.Style = ssFlat
                    HorzScrollBar.Tracking = True
                    VertScrollBar.Style = ssFlat
                    Align = alClient
                    BevelInner = bvNone
                    BevelOuter = bvNone
                    BorderStyle = bsNone
                    Color = clWhite
                    ParentColor = False
                    TabOrder = 0
                    OnMouseWheelDown = scGrupoMouseWheelDown
                    OnMouseWheelUp = scGrupoMouseWheelUp
                  end
                  object Panel1: TPanel
                    Left = 0
                    Top = 0
                    Width = 569
                    Height = 18
                    Align = alTop
                    BevelOuter = bvNone
                    Caption = 'Grupos de Produtos'
                    Color = 3092275
                    Font.Charset = ANSI_CHARSET
                    Font.Color = clWhite
                    Font.Height = -11
                    Font.Name = 'Tahoma'
                    Font.Style = [fsBold]
                    ParentBackground = False
                    ParentFont = False
                    TabOrder = 1
                  end
                end
                object pnTouch_Produto: TPanel
                  Left = 0
                  Top = 124
                  Width = 569
                  Height = 380
                  Align = alClient
                  BevelOuter = bvNone
                  Color = clWhite
                  ParentBackground = False
                  TabOrder = 1
                  object scProduto: TScrollBox
                    Left = 0
                    Top = 0
                    Width = 569
                    Height = 380
                    HorzScrollBar.Smooth = True
                    HorzScrollBar.Style = ssFlat
                    HorzScrollBar.Tracking = True
                    VertScrollBar.Style = ssFlat
                    Align = alClient
                    BevelInner = bvNone
                    BevelOuter = bvNone
                    BorderStyle = bsNone
                    Color = clWhite
                    ParentColor = False
                    TabOrder = 0
                    OnMouseWheelDown = scGrupoMouseWheelDown
                    OnMouseWheelUp = scGrupoMouseWheelUp
                  end
                end
                object Panel8: TPanel
                  Left = 0
                  Top = 106
                  Width = 569
                  Height = 18
                  Align = alTop
                  BevelOuter = bvNone
                  Caption = 'Produtos do Grupo'
                  Color = 3092275
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -11
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  ParentBackground = False
                  ParentFont = False
                  TabOrder = 2
                end
              end
              object v_tipo_normal: TRzPanel
                Left = 0
                Top = 0
                Width = 569
                Height = 504
                Align = alClient
                BorderOuter = fsNone
                Color = clGreen
                TabOrder = 0
                Transparent = True
                Visible = False
                object img_logo: TAdvOfficeImage
                  Left = 0
                  Top = 0
                  Width = 569
                  Height = 304
                  Align = alClient
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000820000009308060000004289C1
                    C30000001974455874536F6674776172650041646F626520496D616765526561
                    647971C9653C0000032269545874584D4C3A636F6D2E61646F62652E786D7000
                    000000003C3F787061636B657420626567696E3D22EFBBBF222069643D225735
                    4D304D7043656869487A7265537A4E54637A6B633964223F3E203C783A786D70
                    6D65746120786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A
                    786D70746B3D2241646F626520584D5020436F726520352E332D633031312036
                    362E3134353636312C20323031322F30322F30362D31343A35363A3237202020
                    2020202020223E203C7264663A52444620786D6C6E733A7264663D2268747470
                    3A2F2F7777772E77332E6F72672F313939392F30322F32322D7264662D73796E
                    7461782D6E7323223E203C7264663A4465736372697074696F6E207264663A61
                    626F75743D222220786D6C6E733A786D703D22687474703A2F2F6E732E61646F
                    62652E636F6D2F7861702F312E302F2220786D6C6E733A786D704D4D3D226874
                    74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F6D6D2F222078
                    6D6C6E733A73745265663D22687474703A2F2F6E732E61646F62652E636F6D2F
                    7861702F312E302F73547970652F5265736F75726365526566232220786D703A
                    43726561746F72546F6F6C3D2241646F62652050686F746F73686F7020435336
                    202857696E646F7773292220786D704D4D3A496E7374616E636549443D22786D
                    702E6969643A3830313231354233323437303131453238363741454333413046
                    4233353339412220786D704D4D3A446F63756D656E7449443D22786D702E6469
                    643A383031323135423432343730313145323836374145433341304642333533
                    3941223E203C786D704D4D3A4465726976656446726F6D2073745265663A696E
                    7374616E636549443D22786D702E6969643A3830313231354231323437303131
                    453238363741454333413046423335333941222073745265663A646F63756D65
                    6E7449443D22786D702E6469643A383031323135423232343730313145323836
                    3741454333413046423335333941222F3E203C2F7264663A4465736372697074
                    696F6E3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78
                    7061636B657420656E643D2272223F3E7DFF8DB900006E834944415478DAECBD
                    079CA455953E7CDFCA39E7EEEA3C1D66A67BF2000E203908084A46928BB2A0AC
                    ECAA6BDE35FCBF755D75DD35ACB84690554414242B39E7094C9EEE9EE9DC9573
                    AE37D4F79C5B55E308435009AB3BEDAF1CBABAEA0DF79EF39CE739F7DCF30A8D
                    46831DFA39F4231C3284433F870CE1D0CF214338F4F31A86F0DF3FF8094BA6D2
                    2C93C930B546CD2451661AAD86317C4E1004A652ABF09F02D3E23D491499DEA0
                    67B55A9D994D2656AD5599C16064A5728959F03BBDAFD5EA982CCB38968AA9F0
                    7D596930BD4E87CFD6F8778AA512339BCD2C9BCD3251ACE3332A264912CEA9C5
                    D531268B12531A0AC3FF3119D7A0C80A3399718E529935F09E244B4CAD563399
                    BEA3D1B03AAE498BDF259CD3AC67ECC588836D8E7631A3567C63060CF7502AD5
                    74B8AE735C76C3656EAFC5249B1C3FED1A5B723D6B08A2462BB07219F72BE05A
                    71D92A95804B6FB07AB581EB1430660AD319D4AC5691995853985AABC23D36DF
                    2B1724A6C23930BC0CC380DB1758B584FBD360DCEA8CE98C2A26561526D615DC
                    2B1DABC1B44635AB16706F984601C7A71F15FEA5E3D3F9148C9156AFC6F1F42C
                    1B9D67567D82E9BD01F6C2774EFC83FBD2FC355B3999B856A5308BAEC627506A
                    6040843F07010572083DE6F55F36ACEFFA870D470FA9B22A03CBD784C392997A
                    B7AC0837E16FB3781598F09735567FD5865097D5ACDF5D60A3FE047B36D1CFF6
                    957A995E05A3607F9A3190776732E5CBD6AEE8BCFA8A4B0F57DDB135CD820135
                    1BEB32693345E9B3B1B4F8D1445ABC5DD40A37297263231021A152B1FA2143F8
                    5FF0A300676B928699D475E6D1A6595D303059D0B5F0E28FC2028438C5ACD1E9
                    AEBCE2DC31D3DEF9085B58AC328B5DCB62B90AB31B356C30AC3786FDBA0B9259
                    E98258AABE2D9B956E96E4C6EF64B9B10F28913E64086FF34F4DD1B025F6381B
                    B14CB0CDD2D12CA3EE605A56E1BCE7751B02D0209D2E9DB57A6570C061AAB2D9
                    ED2FB2ECAC9FED33EA58C16F6216B386D94C1AE6346959974FC33A3CDAD17451
                    1E8D26EA9F49A6C53B8A4CFE8D2235360129F6927D1E3284B78311E32529205A
                    8A91D9840CD32822370605F15DDDA8BD4E8368086A83E1DDCB06ECF65A6A9E8D
                    F40599794B854D6C8AB37C978DB9FD4666B6EB58DAA265563319848EB92C1AE6
                    B19A4CC590724122235E00A3D898C9887701211EC46B1B3870F69021BC0D3F12
                    D3B12EED146BD48B2CA259CD2A0A26CE3444349BA9E5EA6BD841A3C76ED62D37
                    D41798D7E9632693815D7D66813DBAA9C01EDB16DD9A9E350E390226BDCD6B62
                    569781A56D22B302211CFCA5657D211DEBF0E9D6A472E29A48BCFEA1644ABC2F
                    9F97EE93A4C633503FE3870CE12D4506C8CD869689CCC2C218FB4AA5C8AC8606
                    A42C60DFB316D20D7F97442EDB5EAA3EF0B3DA2D25BC031E0DB3DBAC2C932FB2
                    D5430E16B0551E7C66BBF8D148445A55CD954FCD2C1A8E32BB8C219B17E8E336
                    B2B443CF6C3C6C6861103AE6776B98CFA9F5E64BF2FB6229F17D9158FD996C56
                    7CA45E571E90A5C6139080B5B74371FC9F328403A756647A260A4616AAED6485
                    6285598C8C8992CC3CC1106B287F18C2D508E9F3A2CDB9DC271947FBC32C5F2C
                    31A3D1C06617D3C57432FDA92397B3ADBF7C5CB5B5A1C8BFAAE42BEB6BC5EA31
                    F944E92493DD70B8C563640E1884D56960566B9D1B851D4661C5CB6935B07050
                    7F3838C4E10BB1FAFB1289FAD39582F480586BDCAFD109D3944B3864086F55B8
                    10F44C1124E6C96D6395BAC806FB8C3C51264A52CB3105CEED9CFA9CFBA8E111
                    830C1BE189AB5A9DC562B1AF31457E6175BFC07EF11827942541101E1698F2B0
                    54937E564C96D697B2D59372D1D2BBCC0EBDDBE631319B4BCFAC7642091D0F1B
                    C427BA033A16F4E9C2999C145E4CD4CE8C46EB5B0A79E9895A51F92DC8E5036A
                    B570C810DEAA1F596D40D868B0B9E9694E1EBBBABB79B69232AFB59ABC64798F
                    E32C17823EE420D36A556C66666E63AD56FB2FA341C3AA1020258417ABA9D10C
                    2394815509E3308A714592EFAAE46BDFA995EAC7E69395F79AECFAF5561710C2
                    6D6076840DAB550BE9A9E548E17168C03FCCFAFEB0F1B048B27ED862A4765E32
                    51DF5AC9CBF7824BDCAE6BB0F94386F056040CD0F8523ECB53E17A8D009501C4
                    1025E6B65BAE0E063BD653AA5BA7D30009128D4432F9691849AA5EAF731271C9
                    D12A76E74613D3EB5EC24B04210BA4780ED1E605B92EDF5C48955757F2F5F7E4
                    E2A5B3D276BDC58AB06177124AE8980D8AC36640D880710D7619584FC0104EE4
                    C4F07CA4F6CE58B4FEA1524E7A4892945B402E1F636F30913864082FCB1EAA79
                    0691FE2526A1D2684E0B7674BC5FA7D7A968CDA5007E3037BFF02D49921EE3A1
                    450242C068C26E81A3027D57AF7BB91C856D2930886960C6B42C2B0F950AB56F
                    544BF553F389F2F9299B6ECCEA84418060DA1136AC24418D20987885402E832E
                    AD25D32D2D5D8CD78723D1FAD9E9B8F862352FDE0A83B815874E1D328437C310
                    3063957285EDD9330ED8B6740D0F0FFFBDC36E73E8B45A1E2AA6A7677796CBE5
                    6FABD4AA9A568309B35A59B9546223031EF6CFE726D8DEA8C06E7EDACC2C2690
                    4C35378083194516116773436EEC9098F22318C311A54CF5C24CA4485CC26227
                    94C0CB66D3E138442EF142E870F59B54BD9D8620D44670315A3B2A11153F55CC
                    88F7E058370A0DF6DC214378C343042465B562F0FBBDD78C0C0F1E4BBF9730D9
                    8B91A81C8F273E0534D8EBF37AF9CAA74AA5E233ABD791072B6CAC5B60439D55
                    66D229ECB9292BBB6B236366D32BEAD93ABE1A03F4DC0E52785FB950EBAE166B
                    E7E4E3A58B1256DDB0D56D624E92A17603B30025C820AC4635EB0FE959D8A733
                    27BAA5BE48BC7E7574B1764926296EACE4A41B713977E0C8E94386F0E71900F9
                    AF4DAD569F6D365BAEC56B19605C1D0CFA8112655AFA9E73B99C423010306432
                    997AA1282ABFFF2E08A7D24400B3A1C179C3E1FD6576D29886DDF1A2993DBE5D
                    6446BDC06A325F517FB99E65AC8CD72EBCBE2C8AF237EBC9D2F1C54CE5B2F4A2
                    EE9D66A7C1EA709BD40E0FA104858DA651045D1A167069D5D94E833D9A168F03
                    B93C3A15A97DA994116F8191FD0F8EB5132FF1F5A4B40F1942CB00F0726AB5DA
                    93FD81C0C7D7AD5FB79A6A1C262727D9CE5DBB58241261017F802D5B3AD2B374
                    64E8F685C5C8F3308AEB359AEA5DF85E1CAFEA1F1EAF39B504167559C54E1DAD
                    B075C1147321FEDFF0B49D254B5A56049F500EC8561D2860F0CAC3A46E5364E5
                    B64ABE3A502BD62ECEC50A67C72CC61EBBC768219470505E82A3848E870DB7CD
                    C07A833A4D2C6B0C47A2B58FC617EAD7E412F5C72B25E9469CE3011C8FB844F5
                    90211CE4470195C7CBADD7EB4FF4FABC7FB76AE5AA77843A822C118FB389F109
                    52080DB546232C2C461B894452999DB3AB7D7E0F0B0502EBBAC29DEB92C9D4E7
                    66E6E67E29CBF22D38CE2EBCD23C54BCCCD09AAF4A5DC5DE339A607EB7837DFB
                    213D2B571522A32C9157F18214D5C185C02494C717E4BAF2AFC564F1B452A67C
                    716A5E7398C56DF2387D269DD363024FD1310725AACC5AD61740D8F0E858AA5B
                    D2216C1C1F8DD48ECF26A4855CBAFE8B860C82D9684CE098894386D03200C479
                    9FD1683CC2ED767F68D9B2A52751DE209FCBB28D2F6C6498E078BE5058185DB6
                    B473D58A15DE3DE3E3D53DE31377A7B3194F3A931A9B9F5B70B93D2ED6190A05
                    D7AC5E756DB158BC766E76E16E87C3FEF372A5F294284A8B6A95EAA075089222
                    B0725D6017AF89B37AADCA545A0BBBE1590753A9356C21AB661A152149839166
                    790958D4A03A6E05C1BCB552A88D568BE2059948E13493C3D065F7989C300AE6
                    044A38B8516859C0A9E1AFF28085BDB039D6919D9FFA58B75FF9584D303EA5A8
                    84AB71BCADFF670D810C00DEEB01026CF0783C570E0E0EBEAB2B1C6635B1CEB6
                    6FDBC6128944A654AE6CA9562AFF212BCAAC4EA7BF49AD56798D06836CB7DABE
                    924C253756CAD5F38DC6C239E94C76756431D6CB394328C88686969CD63FD07B
                    DA6234B67B6E6EFE66288FDF55AA9571AD56933A184AD424155E6A66D234D8FB
                    D7CE31451F60373CA1625683C2F2753D8C42CBF41A89690FAA3A846D90A9DBA4
                    BAF8957C543EBD982C9F959CD1AEB17A4CBD300895D36BE23234847F0B99029B
                    DBB48B9DBADECBF4C100DBBE507D472653F9020EF3DEFF7386D0420097CE6038
                    DAED725E383030705E77B813E44E615353532C1A8D950AC5E233F55AED464996
                    6FD2EB74759DA0024F50F15C61BD5E1721256D85628189A27873A552A5D7BA42
                    A1F89E6C2E7F5C2C1E5F0634B004FC7E1608F8874301FFE713C9D4C7E7E7E76F
                    9744F94E1C7B23CE3F71B042617AA7588334348BECE2953106F1C166F20EF6EC
                    9C95E5AB7AB698D5301DAC41AB92113A1A2F358882A0166E02D9B8A95AA8ADAD
                    16EA67A517F3C70125469C7E8B23DD61659B9E9860A78F0659F7680FBB737392
                    D5AA7566B26897FE9F0A0DCDF470D5AAD71B8F0109BCB0AFAFF782704707952F
                    420A46D8FCC2623D9F2F3C2989E24D80F39B743A6D9186DA62B1902C5451AD2C
                    E515604802AD4A9A4D46562812AFE005B3CFC3289E17252954AD96DF85E3BC2B
                    994CAF9B9D9DEBF478DCCCE7F39B972F5B7651219FBF281A4F3CA3D168EEAC54
                    CA8FE23B9B715DE597192B044BA90E280752788C4576C99A12DB93F4B0672765
                    40BDC026525696296B8020F2CB7313BCD85578815E724DF6E462C5534BA9CA71
                    FB7644DFE9B3B0DEF5EB7BD9033365A635909435500171F6FF842134F300559D
                    C9647C575757F7790303FD67068341132143241665D14854C966B38F91F7D76A
                    F5DB0C7A7D82E09BFE4E5662329918264D002250C4E6438D9022D8ED76A6D168
                    799577CB233101AA457CF787501137E058C7805B9C88B071FAE26274C4E57430
                    3F50A2A72B7C784747F0F0542A3D1F8FC51FAED5C47B80324FE298732F4F2D34
                    8B6852250D0BDB0BCCD19FE40B544E7D99951533DB3C6F67F94A0328A146E890
                    0F16369282C2A014941B8BA5EA57CE38A2E7931918990C3471225CF06A70A5F1
                    F05FB9211002D4544CAB3BADAFBFEFBCE1A1E1D301D58E464386048C71142814
                    0AB562B1F479B15EBFC96C31CF924C6C1B0FA595090D18FF6F954EA5561B6966
                    6020B41E25033578097E2A9DE646F392D4B4A8341AF78BD5DAFD2E872363B359
                    BF1C8B27D2D158BCE176835B0225DC4E6727FEF3927C3E7F8EDD617B064AE43E
                    10CD0770DC170E927D6422A46759D430B9A2634B5C69E6738BCC205498DA6067
                    A562993D39EBE68663D02820A0EAE676032D26D5A201F96CE87C5695AFABDBC9
                    B2528319F46A66D4A961648DB42C2BD7FFD51A825817A9B8F494BEBEBE4B4797
                    2F3B3E180AF8E8BD5834C666E7E6583A93D983181EB45AAD5A0CFECF31D173B4
                    07A21DB709119CF06083C1C0F761A8040A0E2A3E3E0D45A9B61233981C89391C
                    0E96C964F77F978C8216ACAC30224A40C108F481408065B3F93BC11F7E0A8E70
                    423C91381388B2D4E3F6E0FB36636F4FF7B19D1D1DC76632994BE3081DD95CEE
                    2EA0C4BD5AADAE427B405E6A147510CB5C55C386DD29D61536B0854812BCA1C2
                    24C1C23602254E1F89F19C854A2330A391B1B9A26EB4AC0BBE5367373149ADE2
                    6B171023B4C2FAF94245DCF3576708B41906137A7CB8ABF3D2D5AB561DD7899F
                    5AADC2161716D8DCDC0249C14DF8FBD7483B5B2DD66FD5EAB511207EA0A152E6
                    A80085B43B4D3C4DA6C96C62B22473A340BC6DA8889D35F8C61A4C33457181E7
                    E88C88B34C70B25432C58B581C6E37370ADAB823C98A1D283348E9674CE00E4C
                    EA4340A947E1AD37552AB50DE974FA4C9BD572B2D3E5125C4E27B8846704A835
                    02E2795A3C1EDFBAB010F91DD4C96D40ABBD1AB5F66561A326ABC12504AE3A46
                    3C19C47D89D9D479B63C5007E4370BF5F58D2AAECB7176CE35D0AD31EA998948
                    2624252EF517C59AF493BFAA84124D5CB5523D2C140A5CB966CDEA937A7A7A3A
                    C10BD8CCEC0C035B67C9447A0734FD7598E4DF69B5EA490CA35F94C52AE3615D
                    6D1584261F68C7FA70670734BC863584E68E2EB58AA056D5689D4BA5C00EF8E6
                    189E4A96E1FD666E047A1D061A24B21D5ED48A6CD16AB5615ABE0691CCF21D48
                    8220C338B6417C6CABD5E43BE2B5DA28D4C6BB63D1F8D90EA7C3EF76B998CD66
                    F30D0E2E39A1BFBFFFC87822F9FEB9F9F92717172237835C3E08B97BD03120F9
                    89A8C0C2B63C8CC3C25189FFA80D470906DB07CC0E939A10420F6B570BDA1D30
                    82CF4B72A34435157FF186D05402B531E8F77F58B962E5F14343036132009281
                    0B0B8B84009330806F03A26FC767679B393D81A7911559A9130F80B79BA93CBD
                    6D0884000B4010096830B27498A3425DAC6B81087A3A5F0B10600B32E8878E6F
                    BBA31F1DFE9BB6DFD176BDF6B5C108AC7A9FAF07DF2FE26B315C277849713F11
                    C5311729E1546F884F0A42F55BC5F905A88DE485308475769B8DC2932114F42F
                    EDEEEE5C0A2572C6BE7D533BA7A7666E2D954ABFA0FA8797B3A206DFC165D1EB
                    F8F961F84EA341F3B71E9DCD9B0027A8D31643419070EA4F9B75EA71BDD3C0FC
                    76FD5FB621C03B8631607FBF72C5D869CB968D74E077617C7C822D2C46182077
                    16CCFDDB18F09BF1D108450DFA0E2686C77D8490126D256CD5019A0FD4F4A964
                    92AB015A7EDEBE7D07FF4E7F7F1F6882A0E6832BC9794C62C98D1050AFD758A9
                    54D93FF1F4C275EC4729C844AB56A709E13333F83D46C8427B3BA9804592AA1C
                    E0B9DA1084225E7B30717B61D83F8F4463C794CBE5AF802BF4988C7118848B8C
                    22B066F5CAC0E8F2E587CDCCCE5D3B39B9F75E51CCFC1407D97810B5C0AF45A3
                    D59E647358CFB1584CAC045151114114D5C2BFE023F7F05A09B5F0B2E2DCBF18
                    435064B9C764327F6CE5AA95E72E5FBECC839B56EFDDBB0F216011ECB9C82AD5
                    EA1E10B253F0FE02A65754B50645109A93A500CA6559AA6142EB544380F7AC40
                    0C41686D84248EA003C4D360D284B53600378687075B93AB0E02350E53A9553B
                    048A292FF92164A010A0F050A1F369F0411C2781D0106FB426893614AB0EBE98
                    20016F62F8DE33C1A09F30A338B977DFFD8964AADB61B7AFB6DBAD3008A7B9BF
                    B77BC99281BE5E18CCA53088A7700F37D0E4523E8A8E4F453330CB3E9D56778D
                    D9EAD207BC560629C116A4FA2FC486F1BB181179FFBA07FBCB0A0DE48D6183C1
                    F891BEBEA5EF1B1E1EF2E8F55AD5ECEC1CC1B80209378398BB677868F058B06E
                    3286797C5E82D11C20E954ADC9150911EA80FE22BD87970D9342BA51A230F012
                    CF02920A4722107C12906012251193A871F6F6F57C17837855B552F901FEFE2B
                    7C34DB5A296C65741A7C8021393B05A0000C2109CFA35A839651BEDC08E8DC1A
                    B5A6E9C96A4D97D964EE82BCDD89CF7EA652A94CE37506D4CE15308A0D209826
                    A80E0D50C971E48623DE95CDE64E80C1EC9B9D9DBD099FFF39CE33A3D569DFE3
                    74D837385C76A606AF70C466B6640CECCBB2C910D728CA5FD6323406458B1B0B
                    00CEFFB62BDCF981FE817EBF11D01E8D44D9FCC2829CC9E692889737D7AAD5AF
                    804405B45ACD4380722F881DB9D3828E4ACC380A88AC6D14E405A42E24492C35
                    93406A18825A9D4AA5240B60BB05F15618C8E1907E1F59BA74E4740A278D560E
                    18E75252A98C2A18F0AFE9EBEB59932F943E01D4B8B156ABFF0CDF9BC331ABF0
                    5E96CBE7010EBA106BA6A593B2D228B30300444BA48FA3139043AFDF6F1C4AA3
                    A1D619B4A3B80FA1582AC760A453B0AC1A48C92D38C7AFC08156E472F9CB52E9
                    EC7B61143EBBCD6A70B95CBA552BC7E01F835F9C9E9EF9085ECF602C9643DA0A
                    16B385A552692992CA7E5FA827B6A9A70F5217BFEEB2FF9D864006808908180D
                    C6CB43A1D0D5D0D94162E3C9749AEDDCB15349A7338962B97CBB58AF7F1DF26F
                    A2E9D9023C5652743AAD0DEF79709885FD37A656FF41D28788208CA5C82744AD
                    B64894416E1A801D937704B8C7B5E1659DA7E05F4A1845A3B158A11F784C8925
                    9C63727E21F278221E5FEBF3799778BDDEBEEEAEF0E7CBE5CA4730E037D7B4F5
                    1B71AEEDF89C80D013E684B65E5F68171BB411814ADD7C1E37EFE170E0B5E1EF
                    B855D37ABE545DA9EC86A1D40ED8864772650B3EB305E4F15F8018E7809C9E97
                    4CA5965A2D16B7DDEE004A38DDF942E1343A26390DAE8B4175DC03A3BA4DABD1
                    30A521BFE6F8BFED86C0A5992C878D46E379019FF7EACE7067AFA9D53863EFBE
                    BD64D90BA572F91E10AA6FC28B77B44911792C86320999560041B343D20571B8
                    178918B437A890B134DA5522D4B8A2D1C853A30F70023398BC1F5EB9DC6AB55E
                    DDD9397C06D51EE6F3F9D89EF1F1A770EEAF3B1D2E1B8E7B3B2D4CE138698341
                    FFD90C5C1E03FE4158C9052EB77BC8E7F3B93A3B3BAE0281BC0A6AE566DCC3EF
                    A02A8661101472F6D16979AD2390892608F7C1E7B75C2AF3CD346D4381119BFC
                    3EFF7A280D0908B0B3DDF8A36D449CD0A65290A9E664B158FC5EB1A8FF1E78CB
                    F920A15F3599525D30220619CDEC561BDF8F0174C8215C7E1F878F565FE73C68
                    DE460320B8EE33994CA721F6FD5D38DCB10437CA8AC502DBBD6B3765028BA562
                    E9D7B47700ECFD79927DFBE33826986E18C7C86112527ABDA11308106AC67815
                    AB829D53699901920A9EC9879B26D423BB9314C831B16BED76DB0FBAC2E1932D
                    56185DBE101D9F987802C681C16BDC4FA880BF9F8A39D0293CD9C474301A3720
                    3A86E37F2B57285C572896DE83C1BE10F2706DC0EFED0C8682E7237E9F4F8924
                    7C2E0783D5D08F4C2B54074C689BE1F334B65AD566FA6E40C22054CB22EEE945
                    3204929CB49B4A146B2DB92A606C8A16FCAD079F3FD9E3769D0174F2D27A861E
                    F7494A26168F335C9B0C127D63BD56BF9787B7D7B9E35BF3761800E4563F0CE0
                    1487DD76454728B8CA6CB1F05638E03E2C9D4AE7E07D7690A7DD88915782A8D5
                    89F51F9847204FA7F5A006198428C7CD669E670F1167A3F7553CF60A3C37DB46
                    907AAD1642DCEE83A72990656B1C0E3B2B62E0C7C7F73E5C28E67F0EAFBD8794
                    034D107918AED3481E4B7B19288D40134C3B9CF8F9054184E1FD1213F34B40F9
                    71A954F25C40F4B11E8FAB1F5C4123AA5476AFC7F30F20B0DE72A57C7FA552DD
                    0E55503C701C489D586D56422941A7D72DC764AA2181E320A8DB00FF2C97CB31
                    8BC5CCF90442891F93BA0208F06E288BB3FD7E5FC06236F170479F8DC6E2B258
                    17D5380E49E01945966E0071540E4650DF7643E00820CB3000CB29F0B68B6935
                    8E72F3D0CE80B2699648A41672F9DC6F31F62FF6F5ADFC2606D95E2816FA7189
                    BB085A691998FA0111C4E226595322526EA11E210F823CEBA0F35072A939000D
                    7E4E204A088873BCC7EBB9C8E9709E8CC914A022CAE31393B7972BD51B144906
                    9C6BF6C76CEA094531959249DC8814CE236AB0853A21165D07A99076E8C1B91F
                    C2711E2A16175660723F64B7D9AFA4BF5BADE6C1AECED097104EAE4DA452B78A
                    B5FADD98D0A7A9C6713F2A30323A03A287762D14312B944A3388EB399A60327E
                    8CCD3042D60618C0E99D1D1D670071D424432BE5128BC4622C97CDCD23566DAA
                    D7C4C0D2A5C3EB698D221E8BEF52AB55DB694C48C5084CF5BFC3109A99C07A97
                    DD613D0306707E28183A0A528837DE9A81148C27924958FF7D525DFC1F68BF7B
                    037EBF0593F7CF803F0FC2C13088DC2EDEFF08048B167A4C4E1333D39A802C37
                    FB2249528452C11A8DB683BC880C8B0C0CC7E8C4209E82C17B6F301038953282
                    60DE550C8E0E88348FCF7E44A3D525DB12935E94F82138A6F43044054FBC9021
                    E02688BC7169DA5EADDC2F015BD701ADFFA2D56AB90B867665215F9899999ED9
                    E4F7FB8FF1F9BDEE81DE9E0F821F5C6E3419EF0122DD8EF33F4689245ACD24C7
                    B53AAD87C3F8EAD94CF6E1783C86B0643FAAA7A7E7D86030787A28185847C821
                    4B74FF393E06D97C6E0764ECA3708C3BC0A19E4798B815634B8935A55CAD6C03
                    CA541BB2FC47CD93E6CD34007848D06030BE37DC197A6F28143C8E6EA85AADB1
                    B9B97986F92F63626EAF53518824FD56AFD58A02DCCE6C311BCB95D2B8C3EE38
                    0C30D9439B4D79C2049E409302B6CCA19AB27F2D2F26A5D0004278B218287872
                    009371567F7FFF393EBFEF783DD02393CEC61626F792FEAFD96C83FF8063E9D4
                    2AB55E867111F7A0EE644424484AF2C9E69EAE6A3497A0F98294402FEEC9F81C
                    2D2E5136913CD7E572F044118C0EE8AE5B45994B51121F2D572A1F9E9A9E393E
                    164FBC1BACFE2484F3CEEE70E799B57AED4C90CA2792C9F43D40BC7B80829170
                    B873793291AC61CCFA962F5FF6B5EEEEAE73F1F96EBAC772A9C817D0C0991006
                    4A8FC2C0EF85113D80CF6EAC566A74AE656E8F7B038D050C300992FA3485B03F
                    F6E74D31040C92533099CEC70D9E1F0804DE096D2E408BB37DFBA6593A9DCA17
                    4B2578867C33BCF97E787E49DDAAE923E60E9254359B8CBB7133EF00D4F5540A
                    650EC1B2ACE250478C9BFA13D086136AC18758BD1B8621E3F39D0E87EDA3FEA0
                    FF0CAFDB730C792E626E746F3C7E2348E7DD30CA47E1B1AB31807F8BE3B93179
                    8666D149831B98C3ED64A4F9A9455D734D40A5E1FBD4F8424E0356A1F0342485
                    239D5ECB28E3476DF36CB8E61632E8700DBD729304C711D38B0811B797B9E2A9
                    1E81893ED1E174BEDBE7F58C21C61F09A277642412BB301A8BEE453803B5B02B
                    471F75E455E02F469AD40C5413BEC360DCF952A57417DEBB0708F2244E390DB9
                    CC1DAD5CC9A8719E13401835E460409D395CD3D36A95EAED35040C06DDC4C55D
                    5DE1F705FDFEF5168BC5489E4308908627038A89F26F8151FC13A07CB24DE478
                    9A1644877A37420554A17DF708FC3D7D0F181DAFF0957922A6B9444C169FCDE6
                    985DAF27CD5CE3F2CC68080D0D0DFE3B3165482D48FEC44F4AE5D26DB89EE788
                    2E50C18946A3D3E07711B06F066337B0D66A22F5258491F0E3B3DFF34C838A73
                    1B2E3B6BB4A4D026E0841208054D546A2D40E13300356D1F50A251AF4B31BA4E
                    BA1F5C1F0885F2584D911F5B985FF8792C1607E933BDDF039470385DA390A0A3
                    445CF15D35F88E71767696F218B572A9B20364F37E9CFF7EF0928D20D0591EF7
                    710D40D2F62296C66A369F48A1329DCE123FDA09B34E1C48AEDF5243C0006AE0
                    0517C2FB2F87B5AF458CB6D54098228B510A0174633700362380BDABE0215DB0
                    5C3269BED0C3789C2FF3D1E752AFD1000116F791B719F47A1F20DF056F4EB777
                    1D1F704E2F60FC72B0E80B2127C91B5588971120C58FC13F6EAD82A94376D669
                    427430401E6765254F9445C5CB8A55465E5CA0E27A94A9A8DA4068C67F854211
                    2CA56DA83CE6E2BAF617A2E05AC910DBD5C9F42E2649876BE9C6C517F0B5399A
                    28BA665E69443BAADD6EAA5FD8252A720762BAC6E70B2850192A9ADC12E09FD2
                    E4B97C9E97B86773B91D8964FAEF30664FD13590C1E5F34D390905C279105FF2
                    56AB3D0EA7630318242169A92ED69F11FEC44DD29A3F570900FECE09850257FB
                    7DFE5580742765CD68675022992258BB179EFC7D4CDA23982C2F2EF244C4C735
                    7AAD6E109FDB25B416614809B45643F8BFF8FC34C55F835E070857F7342431DD
                    4E17C3E01C8893578E8C0C5DE4703896C2E3B4548984985E8D27129FC3117F06
                    665D6BB3E6A6DC5371E287705492E1DDCD8254D94269673ED9642CFAA6E6A644
                    0E9521100A917136AB8F1A020F110DA1958EE692AE95826E9E03A4D6AA056145
                    D88BE06B0BF4F9A6D77243F081305FB264C9C0D91D1D1D7D20762EDCA38AC82D
                    95BD150B30846482511670F5AA95AC430A0D4463B16F4422F187C1896EB4596D
                    3BA92B6D6B15B5DD0157C03DAE75219690C1E1BC71A0D3632A95EAAD3304BEF4
                    5A1729A1F1917057788DC96874962169A0672913086FC93C52AD89DF04297B02
                    839124E882D5C2F04B7BED0EFB1AC4D865203977224E2B4D19A6A54DA7DC73B9
                    8149721AD61DC58D7B31E85D98F84D08312E4CFCD55D5D1D17B85CAE014C8421
                    1A8DC6E3F1E477BB7BBACEC344F4812FC844081B07A47629A78F8800B4A9512E
                    A00C364D39009A64DBEFA19EB1DDBB77F3CF0F0F0F330904B5C10B519A3D13F8
                    9638FC2ECB12335BCD0C86CC3298445A5F50B752D9B4C60003D4107A61E2B9E4
                    8593ACC23CBD7F6870E8E44030D0094731D1A45196903800A45F0E936F9B9E9D
                    23325A46587C627666DE0BBEB0AA23D471584728B426994C5DB218893C124F94
                    7F84CB7BB89D1BC098AB80822753F28B10A252A94CD16EAB3FD5A9357F02113C
                    C1EBF57EBCAFB7E770C4497BB15862F3F3BC248C52B44F61D2BE8D0BFC1D062F
                    D3CCF96B9856AD23762B96CA955D60B834392B2AD59A0BA426D98CAF4A1382A9
                    238188F8EB72966AD5DA2E8BC57C0C64D56AC8BA40CFD2AE8F024ABBF03DFDDC
                    FC62329D497FBD56A9FE1837BF8843AF319BCC831AADA65FF88324125481D0EC
                    274DBFC39B8A98CC3A65F4E05056B21A5A8ACEE7737C291983C9B66EDDC243C4
                    DA756B544DD580EFD5E5B2D160AC596D169EA154B59082CE9300A10BF87D02D4
                    49903C1ADE5D2D158B477A3DDE4B962D5D7A0CC6CA0EC3D3D0B1A3D1184F14C1
                    0036622CAE834AB10703BEAF43F231B3D158C0DFBF91CE665F48A4526781D3FC
                    ADDBE55C07C5D10187BB2893CD9DB6B81879A150C8FF04F7720B0C4D76BB5DC7
                    525611DE001BAA3D4B95506FB621A801F147FA3C9E4F0E2C59720CF4B191325A
                    A402701160AF15EA11700D60E97AAAD9170EDAF058E00B2A949C8135AF4DA7D2
                    6E5C7592D7021CD06684A00F6CBF426AC06A311F0BAFF8047887004FD24D4DCF
                    2EC29BBE072EF1239C2341C40F504B1E1DE1A5E7020BD1B508CDBD084C670001
                    6D55EEF02B50AB8116AC44A1010662550932F100B93DA9CD8C9FA8AAD74BE16C
                    3A73B456A36365B982D8ACEFC2DFBB10A6B2F4E103D3B644867385BC4E51CCFD
                    E9748648E7EA77F41C7E13BCDA00A31328659ECB171BF95CAE02A77900F7F103
                    AD4EFB18B8611152F9EB7ABD41E8E9EE26A25781814CE1E0299CE227C562F166
                    FC7E642C1EBF0AF07F222D3F2F5F367242B1BBEB2838DE47F1FE934E87A38FC6
                    0FC7A524D4231AF59F1EE935AF06FFF87FDAE0BD1A31EDE3C38383A775F774EB
                    70716CE7CE9D541A5E4718D8DE15EEEC82C57A666667A3B8A892FC9244064DB2
                    C369E78C1CAFDD18882940FE3058520893BEE725B9705A81F4831C5D613219CF
                    169AEB0BDAA9A9E945E8E81F0286AFA322545A366C7B3EC56A4CCEBC00C28181
                    F0BD34B74E34A4FD0E808016A44A2A4E5299050824A85A311E2F030E37084FFC
                    0098FC051EB834A595897F402AAEF778BC0F8220FE289BCFFD17E52EF079AA69
                    D503FD8671BD97FAFCFE4B310C842C5A4CAA3689980F765FCF178A7148DE5FD5
                    25F1275AB566275596D115111421541C4668484892AD54CBB897482BEB482193
                    52D2BFAD546AF7CF9723A3C954FA0308ABE779A015972E1D5E63329B5643BA0A
                    8434E54A390E477DBAC11A6FAC21D0B2283C6C148372EDC0C0C0F9C3C343AA1A
                    6076CB8B5BD9F4D474255FC8EF032BFF01E8D03731A0DF0573BD0AACF77D8562
                    F17683D1C817590C889D6DB64ECBA22D6F9F00699A3618F5C33AAD6699224B0F
                    CBCD49D063923B41242FEBEFEDFD80D56A0DE2E6645A4584F7EF486532176070
                    76BEB4A2D78530D38AE33084860015E26F287CDD898F09E50768225B0C9B1782
                    489292D7EB39A1B2643269B5C3663763F256230C7D1064EE1C97DBAD45CC2D01
                    198AD4EA805866A15052AC369B33D411FC479BDD7A593412FB4F84B7DDB8DE8B
                    8F3CF288D3DC2EB79E623F8518E249301811EA653BC2E40D502A3F837127DBE5
                    614D03E65225440B4D747D4402C161A280A7427BC554569AABA6AD22EA2D4089
                    6BD2C9F4BFA412C90B803A97019D464D66B300676C4049D0BE88C29FF3EC0DCD
                    416ADF1C369BF58B2343831F0A76843474EC2D9BB73676EFD953CAE5B2D388C7
                    37001CAFC3C0968884010EEFEB0A872FB6DAADC7C2107871C881A59522064869
                    5D206014E3531AA7D4AF56AB5B0903F16AB45A3886F1B2404FCF3578DF0348AC
                    8E4F4C6C86C16C1D1E5A72A1D964D202C6B5C420DABB91284E834937279A978D
                    0B113A0574BB079E6181EF17F8DF1A222B2026D36AA5CD6EE3300EAE92155454
                    F1DBE844DCA7587CE9E0D0E0BB6C64B0A572767A6A6A7B3697FFC9E092812178
                    FB27884F64F3F9ADA954A6D0D1115CE7F77B7DE170C79757AC18E3E5EF4584C8
                    8585791EFB7360FFA9448AC513F12DE04FA7E37EE2B2AC1C50312530A3C5DCE6
                    30A330405BAD56E42BE7E033F3B2DC542444FE0849F9021DC6AF0EF9090E2000
                    A9F0B1FA03F8427FB8B37304C7D021F440368A0F09C29FD75CEB6088F0A99191
                    E18F2C5F3AC217771E7BEC09B66BD7AE64AD5AFD86516FF84E436F28022D70B1
                    556632591053A5478AC572DC61B7F527E2891371F1D7B70F048FE1C99FB6E4A1
                    C59872A5FA3C0CA46A30E88F9424F36730B8E7582DB6CE62A9284E4E4E6EA3CA
                    9F6A5DBCCEEBF158C1BE8FC68074695442BF4AAD7B918E212A22D3699A09A646
                    AB6109DE4758521A185833B8851F465A006231B15AE7A969B12E315C1B0DB0DA
                    E1B00B805842AA8B7B7A7BAEA09D4BE03BD18989C917E0C93FC5B5FEDA1F0828
                    50289F24FD4FD70CF49A32D96C7381406079472860D0418910EB9F1F9F6705E8
                    7B0A4F94EDA3B049E104043187494A0311482DB5388BCCB7D2D1049391625C0E
                    6F95E3517E43849899A36B2523A08799788CCD8D36A22419117AA0384C47018D
                    2E04673A8132B08440A4D4C089F2908D8FFCA9B2F1150D01073E95AA7A1F7FEA
                    69D6DFD7CB8EDC7004EBEAEAAAC662D13424A1B9542A17E99C8B91282C3FC137
                    54422E3ED1DD1DEE83879D0D54B89E17900B7FB8064FD60D03D2814069A0D573
                    60CC431E8F7B884A3880009B32D9CC2F60FDD763B0125A9EF737C15B6A930835
                    BDC0C650A558E61DD2E9885008BC18743F090471C4A0E561205A84139FA0564D
                    AA54BF570F400283DEA05F6AB1584FC2F14FA7A3C0A875D56A6DEFECCCEC2354
                    FA8681BD1F1EC78F478B62D0ED0E9A14A0153B6CFDFAF780B07283A6357F322A
                    A04612D767037FD2D1DF60D8DC400D9840FC2D9748A4249AD8E6D368344CA9CA
                    9C5493A1D05804FCFE35B82E353915DEAB21644DB55668B9D1E07A420819CBA0
                    984E0E0682E779BD9EB046AB26DE3133353DED2247A163E31AF760A467F7EF69
                    78A30C01F1E626B0D03182F34D9BB6B05C26CB962F5F1A0E05FDFFBD63D7EEF7
                    C27BFE03C6709FCB69E7848D278E72F95BBB1AEC42848777C01082B8B90855E6
                    1C50EEAD063AACB73B1D6721EE5CACD668FC24D7A84359BE58F811BCF76766B3
                    29A6B456F2B4184020440E5E3369B15A4EC4677B008D3CB60A240BE1816A12FF
                    8D66FB7C9C248773A480605E84E280D2CA5A160A45BBCD665E67F1794F0B0583
                    EFB33B6D5E92A7E4A590BAE3B8F60F82083E0EB9CA37B440F3AB8D26D30A8D4A
                    7D1C06FFDD186C5E13403908AA9520898CFBDB0782F654BD5ABF0B131886279E
                    6F3418C660083A420A2B42563A9D356322BD2A419768D6301C584C4BCDB71A06
                    18E432CAA412511425B90214D8084412703D2B802887793DEE77C330DF054BE3
                    C529C9647C53B15C79706161F15EAFDBF5DF0E7C70766E4E82313FC26B33FECC
                    ECB0E620E9E25FC02A3F0B35602146BA676282C58110ABC6C6D89A952B4E8665
                    1EBE73E7EE6F41C6FDC4EFF54ED100C612F1874BE572CC697784E291F849F080
                    1B28778F1B53C1436881E5DDF8DEA54693D11B8F25C54C269BC4DF3D98BCBBC1
                    A4BF513F6069578D09D4E1F75A5D22163DC59B5F021A41A1B51419786938FE4E
                    CF84E2AD702B7586E356E14D71FC1E007FF1150A391B08DC299D9D1D670483FE
                    7341F4F455DCCBDCCCFC3DF42C9FAEEEF04930960474F8E381909F81FFDB5C4E
                    D71198A35302FEC06981807F09E973AA129ADC0B034825EBC542716BA55A7B50
                    94C5BB01ED8F6BD55A9C5BA4DAC0DB61A85F75B95D674119F0CA2918F55140C8
                    EF6533D99F6182EFD3EA74BCE49CB283142E0545B51492D49B4A9539A1CEE5D2
                    94DC5CD3D313BE0893FF1E8FDB3540ABA270BAD8DEBD7B9F6C6E9415EF03AFDA
                    0B031BC6587AA9BA1A460924917FF7E7868557E2087338F1DDB8A1F37B7ABA49
                    0EA576ED1EDF91CDE4D62F5B3A621818E8B53BDFE1FCA7DDBBF7BC736E7EE13B
                    80AF5BA0670B880F0F7675862FC3A49C010FBF01107D3CE2D9B91E8FEB4CC072
                    209BCDB37D53D3F72512C91F0703815598848FC1180E8351B8342A55BA5D7944
                    08272A1C3E1B88D1534D09AA0FC29302909C73141B29EB50693D14AC90CFF3E5
                    0AC0F41E9BC5BA0E93701EA0FA389FD77B2EE5E669B026C6276E05A4DE5B2A96
                    7EE5F5790FC7A98EC0F9FC2E97632D42C420F8C07BE0FDC7E19A5C8430F956EA
                    37168DF31DD4B894DF2154FD238C670F0F49AA46BBFF1203699B00719D2045D2
                    90B954E11B683A3B3ADE8B639E1A8DC5EEC864733FAF94CB77D32A294D30CEB7
                    0AB0AEC98368AA9BBD1ADD6363A3DFC33934E05C083F89E761840FC2781F8633
                    3E6AB1596B4AADCE1352409ECBEC763B775290ED05852B0AF98D37044AF0412A
                    5E0F783B9FF6E42196A9D3A9F4AFE3C9D47FBEB069F35713C9C4C0C8F0101B1B
                    5B7E34066E74CF9EF12301E9DF006BFE0E06F702B0F36331C13F04149F02D8EB
                    80EE86574DDD8B9BBE1561E71E8BC5B488898B00792E07A4AE4BCAB20B042F4D
                    048AD70648B5668E9F8812241526A10E8215566BD52190A6398AB3AD82D2663B
                    53FC3798FB4AB0E94E307CB5CBE53A966236063F3D3535F3EB74267D1764C9A3
                    44E0C2E14E2A5829C15045786CC8EFF77FD7EFF38D01A6F544CC284B48F50E40
                    BB7D187405863880EF1327D909A9C7574469DDA22E8BCC6EB1723E00596FD0E9
                    342BA576136F383E24E40E18B8AEABAB73696F6FF7F9BE52F9C4582C7E77341A
                    FD492C16DB0AA5720A3C598731E27C4769281A8C5F66DFBE997BE0840FD6EAE2
                    D340A4DD442C5B092318679E1344A0D591704E0DD0B651AD569FA094FA1BB170
                    78304480332BCF178BF91DB8A865068FDBE10FF8C873BF95CDE5E6F64E49D740
                    4E41EFF73040ACF308E7FA8FEC99D8BB7ED7CEDD7700CAEA904D2E58FF15F044
                    DA8B7837A0F1E718A44700E78BB42CAB6B3EFE6F37F4730281D28FC91F42ACDD
                    4B919E649EC3E9E04BBB7CB3884A95C5CDCFC10BC2F0B880A0653CA6B62A9368
                    52377474769CEBF37A8E476C5F4E3381494EEFDAB9EB87E96CF66EBDCEF03C50
                    AD62041FA16343299C8649FF1B97D369A73D86E02AEBAA3CF51B810164A8F4EB
                    799CF7D7989487A667662F0237FA7BCA834031598090948554245C9BC361E3DC
                    818C11644D038FEEE4DBDE9A9640D9F49B6150BFCD178B17B91C8E335D6EE720
                    9CEA128CC151BD3D3DD318A2B520B32A3AF79617B7346666E79E181D5DFE39A0
                    D0165C539E8867AB0C8EBF6AC562AB3847D30FE7EBA5B3907140D1DCF7467564
                    3E581E816C210BEFBDA1902F7C957A0160F08EB758ADC743E33C881BBE36914C
                    DE572A143F0A085BB364493F5B36327CB8CB615FB96F7A564B440DA46A0EAF8F
                    41213C810988D0C20FAFB9A7F888E383E0C4E1D9FBE065CB353AF57AA5DEB81F
                    03C3CB8DA9F084343FCF3B485206A33A89EFF7E39A5C1CAE14AEB3CFE8EA0A5F
                    002F7F0742510F7D07CA2505EF76E3FAA6A0023E3DB46489C29F1DC998D36C30
                    9CEFF278CE00015B85D011246D4F8A804AE653A9B40CA3BD0BDFFB75A3213C03
                    6236918651003DCE1B1C1C6488F37C736D3B634AE7E7A5F44AEB211F2A951D48
                    15842C6E66B210D1E0C9B398D02D9148740B486E03C6F47788FD4684A51E48BE
                    1E929C38A7323737DF00C790228B8B77E2D08F916A6915DA72744CA5AAED8DAD
                    FCDC0683EE38C85F3B3D4F130E9601923CCEDE2C43682564A1F0E47B72F9FCE7
                    BCD5AA0D13E683A79E0E987C107FCE2136FD5CCD54CB71412BB66EDDAE413C64
                    1D1D21C3D8722B9B9D9B476C8DD63079F77ABDA1220D1C48167796FD1B3D7063
                    98AC6DB8A9330C3AC361D54A56ABC8022F38E0EA4025F07577B3C59292447912
                    067432067B65A55A390F04F07278D71AFCEEA318BB6FEFBEC720D57EE470DAA5
                    A03FF033788D1EE8A0803FF4DBACD60F82A89E44AB95081B56D2F569C47E5A21
                    A5D2AF52B9F23F80BFEB7189DBC12163A4340441C77A7B7BAD4E97DDEEB23B78
                    9F65DAC9443BA2DB491B4C30EBECECE44BD530E00EC844BB441AD56890718D22
                    90D30F24B812E4FA5C10BB21486C6316E78BC71353403F4F30E027E92778BD6E
                    A8045BC36AB164486D50CE82784D9D2FAB2BDCEB495EB6D74E8064C740CE9AA9
                    6EA152AE6C6C6DF67DB30CA1AD1E94D942A9747BA158BA043028582DE6F5B8C0
                    3EDCAF3F180C7CDC66B31E0989A5A1183A31B9B70032660D873B78EEC1E1B076
                    EFDA357E2394C427007313CD5D46024139BF518D8674BAB811839BC3E0AD6319
                    6681AF977EBF2388571F73A43768B53354BAEDF3FA2E073A5DA2D76B9D3499F1
                    78F2F64C36F3638109CF21444521697DF07211DED7856BB863687870B9D9680A
                    D7C5BAA6582851C127219502D81F0779EDC06459A391E80F3B3A3A1E251826C3
                    53A9255EAF28C965B1C31CAC0842A35D8F50E5E410D7473B9572F9028E9564FE
                    4080425588F7E2CFE52A5645910241BFE3D8771EFD6918B1094AC3BCB0482BB3
                    E9DF20CCFE4C50A9377A5CCECF190CC6CB703C357104F001BDD56CF96C329DE9
                    44DCFF2FCA89ECAF70361938321037A0C6A008672B804002190838D17DD4EFEB
                    4D37045C48B156ADFD3297CF5DE2713BA919C4CA8E50E0B7805B2B2E2E303333
                    070F16B2F050073CECFA9999992A38C2473B3B82EA4030A85DBF6ECD999353D3
                    63D353D3FF819BF82E2E5A69A281C277ACC32336C1403266B3B157500B43B8EF
                    98DD6AE7A95AA1D9C56400CAFC83081317519C04A1B4C513F112FEFEC37CBEF8
                    2378F2AEBEBE9E1C11DA6ABD6683B1508D24B1781BB8C819D40A8F183FAE9F65
                    D35996439C9DDA37450FD1BA7C746CF43F41B80E47D8314DCF4CB350A8A35908
                    2B345704CAC5D2105E6382C7C3D10BBCA10664E13C831E379C87AC244E400924
                    A0D252B7CBC986069798A1EA682BAC2A91887B2727F74E43297D03F3760B207C
                    164E943352EB7DBBCD0DE35483E86D037A0D154B659DC7E3EEE9E80C7D6A6131
                    7AD1DCDCDC0F6078DFC7C5E4C80894661692D2E78781DA5055166D7401624B0F
                    BD91CFF57E4543A0A5649C6C5B21977F01DEB616306E72389D4BF68C4FC88978
                    F21E30F52F0029D6F4F6747FDD66B3AD4598381DDCE1814AB5FA956C2EBF0AD6
                    2E0CF6F7F5813B7C6D7C62F2D4783CF5091C7607EF5B487B034471169F9D05EC
                    F7AAD59A35F8FDF1D65AC25AC4E0AB56AE1C3B13AAC24E648C56E82445494ECF
                    CC5CD115EEBCAFBBABABCA37A8148BFDB4F2D7DBDB733E9442378C430BB5C04B
                    BEC9A0609870DEC2F7226083B89ECF2C1D19D245A2F124E039AFE60A4576BA5D
                    6EA1150F7518D8138158978742C1E321D16C34D964B880F78BE1B59B71ACBBDA
                    E386F1391A11ECD2552BC7CEA07A489C4BBD7BCF6EB67BCF9E241CE88B1E8FF7
                    7F604465B0FF7ABB301706ED32E8F543B4990723BC157298FE3B994CA4668341
                    DFEAAE70C79280DFF7FFCDCFCF5F313BB7F0EFD46E877821C20D35E23A0AE7A1
                    DD56B41988DAFDBEA14F867B99211C989CA09DBE88A3BF424C5A4B8DA62C66B3
                    0C45700D58F40F1073650C5442ACD73F69B39A0F075C63221AF7E5F3F93DB881
                    27ABB51AD5E641EE040C6BD7AC7AD7BEA999B5FBF64D7F954AD7A00A0AE46638
                    D61678C3D198CCA3400C77C34B3FB862C5D869B4E32393CDCA5353D3205BE223
                    C3C34BDE8D58D9DB1DEE123D1E970CC2BD81F60F2C5FB6EC4C5C9705D7A3A6F2
                    38CAFF53BE021A7C0646F63D84D75FC0A3E63009470FF4F7297A834E9FC9E514
                    289022B93E148C1593D3817B3E0F1EFFBE80CF37EAF37B3588FD02710332022A
                    5D033F1B848CBEB53E3DF35D4CC4DE756B579FDFDDD5B916A1420BBEA0426894
                    E6E7172671EF43E01E511CF776AFCF9F155A0B6454754D5E0DC358A6D5689D08
                    9951DC6B956A6401FB9B1006CF98DC3775852395BED6EFF30EF4F7F70DC170AF
                    839AB8667E21F2FF30AE77D91CF6553ABD5E9DCEF022D507313DB537D5105E5A
                    9608547810F13F09CFF460B0D46BD6AC0149ACDA49A6C143A631B02F62A07A70
                    63A7085661333C230E0FBF8BB6B54F4DCFCA9814199C423738D0EF03847E7D7C
                    72EFE9F144F2930D45DE8E81DB01432861A0CEECEBED7E0F494BA04A3E168F3F
                    5B2A96AF8321DE0CD4318AA2DC4F891F98E8A5D0D11F06513BD90C9CA5CCDFBE
                    7D53C4FE65727D4994D44A43364D4F4FDF02C3FA4A674788571BC3C32B380E40
                    ABE6A43C00A46B89D622C25D9DD78602C12FF8FDDE00957CB5E4204BC713FCBF
                    ADAD5275BE78A45269D7AC5A712D29062ACAD9B16B4F06E4EF0928950DE0417A
                    7CE699BEDE1EA8EA6E259DC9555B1C8769745A5605CB97CB15E674D8D782083B
                    41B69FC4F5185A5553F31AADA65E2BD7AE4BA7B33F2FE40A57D99DF60FF8FCFE
                    EE9191E13190F05B80A85BF1DDCE764E01F7F6009C4D7ACB0CA1C5F027313877
                    80C8FC0DEDF8185CD2FFE15C2EBB7E6A6AF68BE00D8F00826FF37A3C27E0BFDF
                    9D4A55FE1D811484B6726F20E0FB5BEA6134353DF320A415A5987D94E859B36A
                    D53198BCBBF1FE6D50166B7048A83B83108DC6002AD94745B1FEDF18D8DF8D8E
                    2E23B96415986AA856AF3B695221172FA0851CAA209E999E21F81731308972B5
                    7A0F24E44F4160CFC2C47F5450ABBD557002181CEF9B08BE58A3FE490DDE0C5D
                    FB41835EB71AC6C93C5EF7082F0815255E220EF2C952E90C4FDEF87D7EDA71C4
                    960C0E308FDBCD9790A9D298924E02A0A6982F3E8321FA2797CB7137BEA3F1F9
                    3CD9BEFE3EEDE2E262550382D22C67D7B2E6B6F4E69A8BDE6018058AEA80023B
                    0C7A573FA903FCBEAFBD795768083984C17F03A9FD2150F8A35009177584821D
                    DD5DE131BEE601E342384C002136B3066BBCA986F0D2656DDC40B68A89C5645E
                    22CA49EDC2FC3C03CCAE5BB366E55DD3D333D7CD2F2CDC861896B65A2D6BD3E9
                    CC003EBFB35EAF6DAED6AA498459233CFC5E591223B8817F2A14F2F0CE101B1A
                    1A4008757F70617E1132A8CAEB051723D16F60C0FF6568C9182B576B7E90226A
                    1C712162F8392EA7C346851A541A4FE560994C46049A68F74E4D4D582DD6F361
                    00DB69E06130AB28DC83D07968E0BC60F830141F066EA9DC500C548C3238B8E4
                    D33EAF9B276A284B4939023A26587FEB5A342C1C0E43FDF451D690AA7FD8DEC9
                    BD7C7B1EC13C2686DAF4336DA7E65418C59A6AB5E682148EE25A93B42E22D6A5
                    5243106A4A2B65DEDE9185BF6921FD862B951A4FA8813B1C45EB0538E6B8D0DC
                    1ACF65A3B659529F126BE267218B7F00BE732DC6F1FD90C076CA7BE07A9E20B5
                    F5466F4A7A8584D2819C414D4BA32F60E09F1B1E5AB281D2679B5FDC56B03B6C
                    8515A3CBAFF6783D67C33BF408C01ADCF8BBEC56EBCE6C3E9F8074FA19C8E5B5
                    8889E76875BA196A583D353DCD26262659B379B58F19FAFB183C88A501C780D8
                    8E638F3E72C9F4ECEC5198D80BC02F4EA4F57B226CC4FEA96106A85FAD58AEDC
                    B977EFBE5C7777F74592A448A9746A269BA50450171513CD929F60920330AE11
                    BD4EBFD26C31BF0FB04A2B98BA9A52A1E567FE802E5A4924194A3585B43C8018
                    CC960C0C90EAE07518F1589C3DF8D0665C5F0493A4E159435A4E8E4463D4718D
                    DF03A11CEDB852F1FA774D1F877A59CE2BADFE7594FD54E31E68010BF7B154A7
                    D306308E04E9301CBD2F5FC82342CA53BCA1268C80AE854219F11292B31E8F67
                    1AA8F7595CFFE19897C37936B126DE4F7B7DFFDC4294D74116D507DBBF300DA6
                    7A2F88CA062AB4EC0C77E6366FDAF2FF16E72347AC5AB5E27CC46D232D82C07A
                    DF9348A5BF2F4AF53C62FC33B0E26B5D2EE73924145ADDCDAAE3E3934FE186AA
                    904DC7408E9A203FF902CFD0D092F3EB92B401703C469E4C4D25E66617F0B714
                    1D2CA2D368FD50237978E1C7F07D1D60F354B3C9D41F8BC7CC9228821FD42935
                    9DA7256AC0FEA0DD66FB0DFD4B296BAAF2A175FE2AE23415DCC250F97B94F124
                    84A2841849405A169E999D65FBA6A618983C478312063F180872E3A0184DE841
                    3C825AE00603640C5ED60B1A9BCF172EA773D0723895BD13DF692684D41C5974
                    7AC32854820F28B407EF51AD8209504FF26191359ADDDF28F4123AD27F63AC4D
                    308825769BFD6C93C93C4CC601A3ABE31A9EE6BDA294C69B6B08CA419A2EB55A
                    C13C0A3E162B974B7E48C24EAFCFC31617A2573CF6D8130F77F7745D0A1DBD61
                    686870BD6531F29F38CA26DCD059B4908418AD05A4D5113E2AE05405C4C24F60
                    003626E2F1AB2A95F2877D5EEFF296F1B9C0AC695713DBCB0960AE0432F83C0C
                    E237189FD99EEEAE2F10EB8EC7936178CA1EBBC35E058F0C01CADDC562D9D268
                    284738EC0E5A77606EB7CB0A34B1129A6452191E525C6E275701B46388B6B787
                    970C304AE9EAE0E56418548F19A58AEC5285C7622A4E017D67F3402CA9F520D0
                    D69E48BE9137D66CE0C560D02C10F453A24C253537D00E812CAFC28737D3A4D2
                    778BA52288A27314636B06817A1AAAC84E6B0D30C6798C6BDC6430F2A69D90B9
                    945F1801495D01D279B4D3E53ADE66B1F4ABC133C831F0DA82719A7F85EE6C6F
                    6E6868BF0708DB022FB81FB07A3165D7A8E1653693FB8D5AA3FE2988E36FA8CD
                    2CD0626528E87F3F3CFDFD345054C144BB8AC1B29F0301D3D9EDD67580D3200D
                    3406ED7B78FF1930F86BE1C5174085185A4DA8CA1393FBEEC400DEA528F27D38
                    7B5CABD518F0F90FC053C6702DA43E9E74391C694C7E1FC2C117E099411C7783
                    C36EE5491F9A2870174EEE6038DCB88933C080180697EA176865934563517084
                    08CF3B10DCB70B4C0929107AB8C7677259BEDED068C9EA763D04190691563A17
                    959685403E69238FDBE5DAE0F579FF676161E1BBC542F97A9CBB04831030D14B
                    696737EEF5392894208EA4C5F776A5532949713846F0DD3570A62381741BAC36
                    EB72034218E50CA0D0A8D9B7969E2E03847A901AA730E16D3284D6FB45C4AD07
                    11FBCFC644180177C38B06FD49989C05C8C76310D37D34E0DBB6EF805C0C0A3E
                    8F4765321AA8A53DDF5408545FC46780F0EA15B083BBA82814DFDD52ADD7AFAA
                    556B6178EF71B4F55C6FD0535DFF35308424159452F20C0302E557DF476B1088
                    B7C7C2308C805F0F794667B8E31C686FBE89858A47D3A9342F27CB64F3BC7F51
                    004AA513E48FB6AF13B453AD0175698DC148DBB982DFB7DC6DE6F4C9EB698999
                    6A0FD82BB4C66BBF4721311613C931B821D05ECACE5070A9D56CFE36C2CC3B16
                    23B17F8561CDE36F010A2BF8FC4338E847E0043A8D4AE31B593AF2EF5EB7FB9D
                    949CA365E76A85AAA1F33BE6E716EF808A91C3E18EBF017F0A51791C78C413B2
                    A45485B702115E09765A39EF078AE5D2562888C3A8851BACF78B802D2AE5EFA4
                    65DCC1C17E4AB9E6376FDEF2E5503078426F6FEFC9DDDD5D048F23B8390B3590
                    82712CA55A3FC8BDBCC83B8F08B49BF97E30E87752B1062056EB743A7CD49EAE
                    D99482AF5636A0BDEB94D45A31B6FC62C4FDF7EB753A0BAD6188F0C848244E8F
                    E1E1B50454764FCBC42B578C718F361AF4AC0068A6D8BF67D71EBED864329AB9
                    04149A9B575F36D1CDBD6EAC553FD878C53E44ED50D15205FC5F22A164547EBF
                    5F58BA74E42278FF482E57D85A2894BA53996C01F2F504ABC5F64EB7DBADF2B8
                    3D472A4C39923C1FE8B217CAE5BE4AB9F6845A256C43B8DA063D793E2D9993B1
                    168AC50570A0717E610DF6E61BC26BE497E641941E2C954B87994D5E8AC57D53
                    53D3F262347AB758976F09F8BDD700A24762B1D8DCC2E2E20740782EE9EC0CFD
                    5D07581BA0D346720D56BD04BADD0E14C81339526B7891E9B34004AA48229430
                    427B0F628077C26B2C2098677676044FC5B98E21466D34EA9D7C19B648BB834D
                    6C7A66968720F2FE0E10BFEE9E6EC4631BDFCF4025763B7747F9A65C9AF0665B
                    BB6613ABD72AEF22876834DA9C49788DCF12BAE938BAD0F6F41C78481ED74799
                    555CF32AF0A25574DEFE9EDE9A6144FFAF389E8316EB301EB172B1743F42C403
                    38CCE64AB53601A553219502E45081849F88FB37B708EA23B420F54694A5BDCE
                    0AA5573637BE20532AFFAE52AA5C223BA530A41975F1BAC16AB67E215A8CCD81
                    CC8400915F369B2D1760C26FC6C5FF1BE4571613F05DA7C3A1A7542B0C623497
                    CB5F048FFF372AF624AFC649B7E3F329204280B274B08C53A1E97A7B7ABACE45
                    CC1D02E174911B90A4C2E17815313DAA330CB64F70BC3234C6BBABD3922DA598
                    77ED1EE7655DA4BBEBFC390E32EBEBEB61206020A2338C0A425E2B91D6AC8096
                    DB3DBD5FF3F3D43C834AC64802526D428397AF1B11325C8C1EDA41390AC4783D
                    38891E4A87EDDEBD4702CA5ED9DFDFF738B84E86A432E54A8847901CC6390780
                    80A3940985CCE4610173937FB33ADCBC9C23BC4AFCA1CDA1B4071F52F20530F5
                    3049BFA5C3C36B2391C532ED1580D4BA4756942F23B6AF86FEED349B0DF346A3
                    F17159922731310351BE8EDF6158BD6AC5A76767E78FDA3735F3AF1693F949C4
                    C6241061278C232008163DE5FDC141047883496CF54E8274643910264A2B533F
                    83EEAE30D540F0656F5A0D2472188B25F867097914596ACEA140953E02CF1292
                    445C988FF0D67BAF3EB16C7FFDA4C245DDABEF0C6FEE7CAE53AF28467B31290B
                    EAF3FB9B2B95308CC86284A391C36EE3A494D61DDC6E4F0AA846F58C4CA316F8
                    776947185D6B9D3F29C6B80CE3B89C924E503A059C634BE38D5C6EFCB34203EF
                    2422D41173EF435C3B01CAC07AD8FAB56380DE7BB6EFDCFD398BD94C7B049E71
                    38EC782F7A123CE4C7D0D25380FB4741F006C080B9942A95047B57B8F3348FD7
                    BD7EEFDEA91B70D08849AB1DE0F979FCD04615DE1C3B9DE519BF426BD9971068
                    D9B2110C74272F79A7BD1533080D24ADA83917A58A1B8DE624AA040D6B6A6D99
                    B5771BF1D8FF1ADEDDDC68C61FEAD07CFEC1C19ECCD6AA942203A062165AA2B6
                    98AD084D56CABAF23C4324BAC8773D8117F04660840CF470303206AAFA82A73B
                    206BFFA954AE7C4DD0A9AB7461CD3D105483402AC7BACA0C98E0398B52E9056A
                    38AE6995E9BD258640DDCB5FF50B1A6AF054BE1344EACA6432B18AEAFD464797
                    AF3FFAE8A36ED9FAE2D6EF6632B9A7C0740FC74D9C8C49FC312DF000ED9F7038
                    ED57501F435AC88945E33B12C994100CFA47962D1DB9165C40C6A0E8C92B5AE5
                    516C02134C0343B1DAE375F1C927664E4BB23333333C354C953A94D9A312384A
                    C27096DF7C7855733285768776D63286FDFB0E5FD512DA8DECDBFB320ECCB134
                    275FC32796108932A434C1E4CDD3D333CDA657C512EF82424AA6ED412465A907
                    44B833CCEC0E3206BB1EE8F8D9C9BDFB8E4BA7D39FC1043FDDAE8CC678F96D56
                    CB2ADE29054E50ADD59E42088DAAD52AF66681C2CBF73548AFBDA8A548D24226
                    937918B0BC92B67E3FF7DC0B3428F6356B56FD233841050A81F605AEC044852C
                    56CB22754EC5A402DAD84A9960CF6CEE77BA1C0DCABE61E0B4F012AD2C3BB9F7
                    B7079A0A5DA9E2899E92464490B28F9BB76CE1DABDB5158C561539438FC5E2FB
                    FB151077E08FE7A12E061CC19AA8D068752F69153FBF2A26F0CFA99B5DDC14B9
                    D13248156F8C418F0C06896B3E83010846656F88F7FCFAA6F64D9127B7D2C4CA
                    1F10522AC8A59C06DD0FFD7773FF8359BF626CF498B9B9F97BA66766FE03D7FF
                    F55AB55E561B5503408DD5744F18CB0690EE39461BFC1485FDEF080D6D968C41
                    CA66B2B7C252CFC52486C3E1107BE1852D4FCFCDCD99D6AE59430FC324D6EBCF
                    E8F5A7184DA69B71D32E4AABDA6D765A81A34130B47B0B10F491E4A23D88CDAE
                    AC1234BC91D19E0A7A5C0E3D96279148F16C9FAAD932A629D96055D44FD1C49F
                    BA9661225FF66D1A82D08AEB8D6653ACFDDBEF9A7EAE7AD578DF68755A6D7B3F
                    35C7E8EFEFE1358ACD15C03AA5BDF9A453DFC35213BAB981D0621525980EECB7
                    F07BF36A2E42195A85BC195C339D8F32A120B20E18CF17C727264F8CC7E31F06
                    611E800385284781D0BA039FDF273081BD993FAF5A98F26A3FF55AF549C4EEED
                    26A3B1D36AB10A6EB753333D3D7B5A3E57F8E8E0D0C09583834B1C08039F5E58
                    8C9CDDD1113C0EAAC140719F0697208FE23BE967F27ECAE489BC8DBECC11898C
                    85969989F411AE538AD5D0EAB3449E2DB784B4DC6A80491D47C5BAD06E96D732
                    86D6FF5AE497D71EF27D09C24126FFF72180AE8D88255D13C93F521A64B0648C
                    FBA84530623E4DBE5813F747184280A125037C85726E7EE115E19B4297C160E4
                    C491420811E0325F70F3F2665B6B56AF3812A1E209204692AE83540F8CE1797C
                    75861CE02D350441787D27A4C9492553BFF279BD47E1C62D5EAF771D607BFDC0
                    40DFF76667E656E6F385E3FA7A7B07C646970D109CD33A3EE5FCC983694F212D
                    F7EAF456BEC842A48E4206793D0D1250847B3135CB244EC29B6C284DFECE9BCA
                    369A3ABFC9C055CDEE6CACDA9C78DE1C0BFFCACDF6B3FB7942EB5FEAA3D866FA
                    ED2C225D036531E9117D5D0847140240E238DBA71C05317B5A406AAF397032DA
                    9A18E21EB4BD8FEA1E78BB5E22A8071942FE74171833DD37F1070A0F842EF414
                    1A2AAEA5E353BD46674787158ACB4ACA0632BB81CF3C8F7314DEE845A6D73484
                    D7DFB5936F24BD1D6CFDE36AB56984060F9E749D5EA7D3AE5A39E6A266531313
                    133C96528A972699069F433C268A3C83F391F6134E54CD984E93DB4EF1D2CA5F
                    A3F998B456B64FD9EF85C0012EB3C8707973AC56E5AFD0687D96DE9230698C3F
                    7EBD992164CDA7A650C690D2DF549544CBC9402C5A23E0E18648DDF3939BF8F2
                    F802BC9B4AD6C9E3DB467330A9C937BA2A4D2EF14A4494CE4F61818C96F66DCA
                    ADF0416843C646E84724938C9F36E952A37218681CC3B98DF8127BAB4303EF67
                    F83A7F309129C4F707DC83AE21C45595DD6EF3139DA18C1E5D3891AADD7BC679
                    E7754A358F8D8E326A29434993763CA666577CE254FBD3D8BCCF2D2FE8A0B0C0
                    8DA0F5F0AA16AC375ABBA025FE502D81AF20B6DC9E36CB356D823FA98537C2E2
                    59462276145668C269E998125094572066491C65F38B2FB26834DEAC50E69DD3
                    CD2CDC158657165E33646AB921343BB833A1F10A8600496832B1666F097A5E83
                    C8151A25C4084D08912824D9682B9D41C7AF1FD7FC3C94C51E8D5ACDDEEC9F97
                    2F43FF11F284BC381289FCA8BBABF312104107DD6CAC55CE45374599C4803FC0
                    AD9EF43E6D1FA3B84B464029E252B58E81D050873326F3829266950E393EC175
                    1BD6DB86D0AEF76FE533B877532715321AA19516A696C9B419956A0B087E755A
                    3D4F3D93D4A39044B50754154D317AFB8E9D9CAB509EA2DE6CF6C93D977A32AF
                    181DA3ED736CEBF61DFB3BBE1E9C60F24E2DDCC3693BDC4BD701F6E71CA8B107
                    2E92D4055D375D13C96142520A2D740FCD25EB12C63047D9C92CEEE73690E8C4
                    9B98477A654390FE889DB534F8F56AF5C5679E7BE1DBAB57AEFCA4CFE7A566D7
                    FC419A44E0E8C6A83A87069EBA8753957105E8E00139F2032DF488BD7C1F0326
                    50DBE203F49DE653D774C4B39BDA9FF7C91538EC371A4DF41014A15565ACB48C
                    86F1C9A7A210EA2E06EEC2A866C2E7F53187ADB9858EAA9C161716794F07EA6E
                    C28DAEA1EC5F3C6AEFC2A26BA1FD87A572E935757B3B349021B4FB3BB427BFBD
                    4E418865F3B879D30D3282FEFE3EAE40A8FEB25999441CA158011A95613AB429
                    9B5061130CF44E4E90197BEB0DE18F8D44045B8978F28B4F3EF5747CC58AB17F
                    06CBF6527C6F2650EADCD2C9CA69C09D90959478A14520A8089E9421488C449B
                    05214440A9533A19037D4FD5EA8A4586A28297CB2A1EF1F7C76BFADC81CF5CA4
                    8C9D8F261F9E4FFB27E97BE4ED137BF7710388C6637CAB3B3D53910CB4D95545
                    73D0C9551A4D55F27ABC91D64CB8535008AA37EF998C93209F2AA45C0E3B7704
                    DE5054D5E429AD7E8B55289017C19B9EAB566A9BAAB5FA197D7DDDEFF5FB7C54
                    24330D299A922489B1B7C310FE941F40A30C2DFD9D8D1B37ED2B140AFF363232
                    BC9C72ED99345F616B2680789F211D7F30864ADDDCC1434F5F252246128DEA08
                    16228BAC55D9CBB385842EAD05038E0ABC394AAB633A470E7A2E3406DEEBF1F2
                    A2131BF7FC067F162449BD58B3C50D20BECC6526A10411310A19EDE292575C6B
                    E0CBCA0DF66A76D0EC7ED6CC3750328CEE8B0A54480A9201D07E46834EC33D9A
                    082215B9E4F3C54AB9547EAE56AF3D8BD33C45FB39703F7B80063AB3C9F83EEA
                    F64E5C030831DEA07DE98AF0976308ACB5AE0F4BBF676262328A9BF8CAE8E8E8
                    89948327BE401343E8C09F71DC7C22266F739F48343BB7D3C09174222FDE333E
                    C1933266B3A5D9088BDAE468558CC2AFCCC3067F2027277D6E979BD71EF0C6D6
                    80D7A9A9695E904AFB0EAAD4109B77736D721932B825D0FAF4E0703AE7ABAEB2
                    B62474A3D5E2EEA53987D67E487E5EE2132403E9DF5563CB9985C77B0D27A854
                    C69649F386160538CAB350588F831C3F2349F20438C314EDE9A4F682851CA923
                    6500C4F108228A70868C24899324B5DFAA9F37B44D3F41310667D3ECDCFC072B
                    95EA67962D1BB992605AABA52DDE19BE8C4C0FDCA29567F2647A6C6DB556E1BA
                    7A66B6CA3D960A3D9B12B399DCA1A4136B71047A7632C13F912BC2884AADCA7B
                    1BCDCCCCF1384B852BED27AAE828B7401ECB1AFBD3D6F4C00D4D6B77F1AB1B42
                    53A9B42527E31B51E4F653617946D3E97270E372F0E493AED593B9C1435CB1DC
                    E41F140292A9F46F2B95CA978086B3B8E445B351C719307584A3F31069A6EDF5
                    F4AC07A0A099EA1A10CE162AE5EAA456ABFDCB3484FD357D2AD54C2291F8F8A6
                    4DD5BD83834BBED415EEA46719F0EC5C1913AB6D9142FE185E5EB9ABE7248FB8
                    435D94F637E7A641A14126B947328B7841A95C450889F0B67652ABD4BDD1266D
                    F044521E722BB143924EDD68D55CB614010569613FF16C1C5405108210E1A46B
                    925AADF3C8F8885B90E4245E4364927776A17434F80BE50032995C2C93CE3C2B
                    2BD26830E8EF0D7774927A5AB56BCFF851922C7F950CA3DCDCF7C8C922F10442
                    39F02C9D56AD196DDF63B95C8A2074106AFCE51AC201059E0578E97FEED8B96B
                    B258287C7360A0BF933277945D242FA01F4A36F196790DFE70EDFD0450969BDE
                    6432E9399AD044D0C25253E6894D02D9CA291037A06C602299E6FABCBDA399B5
                    7209EDCFB6615ED57A3E03DF47222B7F600CEDE7484B75913F1F9A3E4FE5779C
                    7F2076F392F31689A46D6C941FA10DB63BB76F7F68D3A6177EDBD9D9F5ECC9A7
                    9E968E46227DB158E29BF8FC11207EFED5A6155F18DF33B1361A8BFF3DEE75B1
                    BD0D9FD08E261FE735EAB4DAD5A42AE81A4AA5CA2410A8DC68B0B7CF10FE18CD
                    FA5A6BE354BB00AFBC6D7E2132572A55FF6B7070601DC13F2560A8BF000D2E11
                    489AF476F1280F07947923E3805712EC73D5D05A70A24A20FE40EE56ACA63508
                    BBDDC153C264084D586FB5E96F2599DA0BCFFCD68496F2682E48349FD0D2DA12
                    4F993F8FC7C5279E94406F6F2F977E9489A41C019D8364E7FCFCFCE29EDDBB1E
                    7FF1C5CDF78DEFDEF364369B5EA0DE4CB8F606DE67577DF823A94C3A75DED4D4
                    CC97AA95EAFB3B4221E3D8D8B2730CE38625E0271FC6B53F7520DFC08F49A737
                    ACD41BF9F3186A5012E36F56DDC1EB36840D471CFE3A0CA029D7B6EDD8C921F1
                    D5622E6DAF87073C1F4BC4DF0B3EF03510B6F33D1EB740B196FA2C359FB24E8F
                    D0D13527AD95B8A122139293945B20635109CD76FEED96FE9447689F97209C8C
                    6B7F2D81AAF56C74E1C0B2B3E6F5A9F63F8FB929418D084BC43DDC6E278FF944
                    D0E8126AADCEA669BE77A1449D59D8F6EDDB76EDDEBDEB86C9F1F1E72B95F23C
                    262D8FE352DD5BB30A1706FCDCB34FF3FE93575EF5A1F94C367DF5FCC2C20E18
                    F6E73A3B3A1C4B970EAD0402DC393EB1F753ADA7BDD75B631434990CFD06DC27
                    C63381E3EC6AE64684B7CF1068825E0F12D04D8F0C0EB1C5286F3FC799785BDF
                    EF7FF8B572601F62D57C3E9FBF7CE7CE5DFB7A7B7BAE0D8542665211546246AC
                    9BD61E780651A56E3DCFB1199B9B79F6067F9FD28F7C52596BA9B9D5DA96881C
                    CF2EB6B3902D63E58F6CE24F6F51B1F61355C820E88931B4BC4C790CCA675017
                    97566F27462D6E68595994EAAC5012A1E76B2C198FB05FFFF2FB6C7E6E7A13BE
                    BF99F1670031474B52282D43909A80DAE0E49716CF5482AA0663F9463C9ED851
                    AFD5BF19EA080E767777B91012BFB76BF7F80842C3E741240BE01FFF3F7BD702
                    1D6575ADF7BC329967263393F73B044802C85BA46010792AA86871595DDE5BAA
                    BD3EAAD7ABB662ADBDD7DABAAAADF6216D5DABF5858FAE4AAF975BF10D1612DE
                    C82B21849017096112F29CCC7B3233FFCC7FF73E737EF80D41832601AFF9D73A
                    6B42C8CCFCE73FFBECFDEDBDBFBD4F892E51A7A4BC863FE0EFEAEF0F1E1F4DA0
                    F89531C2D9DD16631432B2ABB4B8643BE98174F2F27242D41A86BCC5104EFC71
                    74311BD1463E999F979B438524C43622E6516242226BA54F8124120E0265D410
                    8B3C0CA9A7B128F248A38C7516079E097110080A4E4438EBFAA9595532F54AF6
                    30AF83082DB4F874D1F7502534A57C7D88F8FBC3149E5681C16802A72F0A1151
                    8F20F6107475765097744AA224E1A004848F315EE22326F733AB8F54C1AB2FFF
                    05EEBEF73E8AA588089E3F4293B2321409AF43615B9C9991A69E3EEDB2874818
                    BA7BBA1F44CD37496FD0B18F208F0105FBE4482799460C2CCAF3FAD2113AC4DF
                    239F9EA26A84E0A9853D69125CB457F1C11EC39FD71514E4CDCCCDCE56757476
                    B2D2331D65E834F1F8017303711143BCC116C4149C7C42EA1F1F55348E0A89CC
                    428C61A62744F14CBD02E10793C1C8F21EE4AA32ED427D12895CD2E7A2669DC1
                    AE8E4E5F7D53ABED647B50298A89B0F2FA1BC064D7A3360880DE1801BFA701F6
                    EDD982F710244DD004F106153E2E0C01FEEF73A843515E1321334F0DC1407095
                    E394E399507FE88EEC9C2CF3B4A99397A399F820180808FCE051911A8D90861B
                    4DA038625E833CE32609064D8E227F2410E8FF2B932DB6166A30E538D5569A91
                    9969A47E0514E3EFEEEA89DB7DA5C00021B984D2E22AB83B407C841848BA1FCE
                    34E92253428243A15D230251623D1176600770A0B03871D7BBFBDCBE2EDC862D
                    CDCDF5B5C78EEDAB3E72E8786B6BE77F145E76C31CB5DA00A9FB8FC0DCB9D311
                    8344A0FED821D8B8E145709CAC3D855FBC09470B0E7227BA71B870F8B92044CF
                    F05E215E32979E96117715C5F8296FE415A1608450C01F42CFA9064DC7DA9C9C
                    ACC271E30A0A2921C74E8B09065C3EAFBF9AB5E91B654918B5B3A179736E1B2E
                    CC78B48965E82AADCAC9CE9A8186517BBAB383F5294E4F4D63C89D8A56422C0B
                    29B283B4E00C2741C969640A767E338B35A01741DA81848E42D554D748492FB2
                    D1E4A61220F5FB7C6E5CFCD6134D0DB54D4D4DE575B5B5BB4E9C68A473250C38
                    A881E546A7A33C79C2948579FFD8F8B6F6E30FDF8570BF0B4E3757E2AE7751AF
                    A2B7701CC24175054E1CD49F20C8C15E4CAE114A274F81929252B8F1A69B19C9
                    96EE999A7A27994CCCABA1809A3631F125743BAB108CBE69369A2650DC845540
                    3BFB5CF87FF5C651C607232E089C09A4469547BD01A66B13B5CB119C5D83B63F
                    89BC01E208102227B66F6343531DBA5AE6FCFCDC8CA27105CC6DA4C5944819E4
                    C201A79BB10EAE5A1DCB4710CAD7B2927260717F571FF550F250A5511FAAE206
                    54C9D5781FBB36FECFDBDB776EDFD6C46F8D0042218E3C1C74B878D4E5EC7A67
                    7FC586C9C9F6F4D456875F250ADE5EFC7D150E6A6AE9E0E6C0CDB540FF60D860
                    C5CA1BE0FA1B57B3C2986E34755284524E8FC3F9A6E2DC4B1007CD44DC13A57B
                    36E13C2801861B4583A6D20717E11A1141889FEE2E6421409B6DB35A17E2C22F
                    42943E4997A8676A9AF202AE780ADAE9F3F90F2188FA6754886E6D6E69C94435
                    FA547676E6245A6412148AE0C5D3C40A06E04863B0B311282A49A96744FAC423
                    A41233AFC7DB8D6E5D6D382C54521F017CB07B7272724E3635364053C3992664
                    9499B292EF2E65D2F9E8C01575387B3A22DCF61330EC940143C90C44066A8139
                    577C0B3D101BDC72EBED70A2F904E36128A4C341E37D9214D4CD05EF7BAED160
                    5D6CB559AF4AB1DB2D949720120C7929542AE7F3FB2AD09D6D8F7DDD0581CE47
                    46119F63494A2A33994D0B4C46E37CDCB13A5A309A2CB998D4F81205E1203594
                    C6BF2DA736BD88007AC8965350E994C3D1E6F5FB9E4615BA88D2D994D123FB4F
                    7D8CCCE8F211D6A08013D1BB28D28860F434AADC4A5CF403F8D03F4521388C2E
                    709BD47002FD7EF8EB1BEB11F9B729F87C45BE98FD7C8747B9BA07EE02FA6540
                    502E00824C03302D30EFCA0540CCECA5CB57B0821B625F9D759D45C22EB93887
                    792929F605E8AA96A5DAED2554D3407841EA02EBF5F85C880DB6A2F7B20D01EE
                    C660ACDF4987918E66C269580481B78B294221A0330E97E0E2CF423FB990E211
                    74303895A8D3EE27F710777D270AC4F3B8909BF07D355A0AE6A03B475050C351
                    3EBA5BFBFB9C7D6B70F17F663699EEA4C52490495E04B5ADE1954FA751EDEFC5
                    45DFAB508807C3827054212A3AA953A954E6CEF43F7EDEC103FBE1640BEBF62F
                    F9FC02B7EF4ABEC04EAE1904997608F157293E108501DCA36B565C074B965DC3
                    802995F251424B43A7C80B4202CE7B01FEFEAABCBC5C3A2271B6C56AD1D1BD90
                    A644ADC54C171D85849A228AE6EF25D48A4FE9F489EE545B1A9D5C03870E1E84
                    AD5BB78CE8C25FBF72D9F0090202B785F8006EB4596D7371B74ED668121209CC
                    F91014F522FA27A046CD3589BC42BD8B3446CD218B25E98FADA71C5E8A1ADAAD
                    29F0FAFA17990F4FEEA1C96082EFDF7D2FC5E01D1E8FFBA15EA7B32112151EC1
                    076573BB3C6DC160703B82CA5D64BBF1A1D689A2A29BD2B6EC74747170216534
                    789923C31757DAFD727B2C0C880BC4066200A605E697416E5E1E2C5EB28CE54D
                    E8BC085EFB5882F35F929D95B938D56E9D9A9C6CCD35720A1AABCCA69E0DACFC
                    4D60893116DBD0A845A3D1F00F5C7C776D7D1DBCF9DA7A96646B6A6A04F46A2E
                    AE46D00C81C54C8B8613FF2EDABC9F656464E413F3988248D1988FC5E2896442
                    769C2862143FA789B3DCBDD1546349B20466CE98096BD7FE90752B3B5275F833
                    9F4DC5234434B9F3AE7B508E3CCFE2F7ECECE9EDB54785980317B639168B3AC9
                    3D8B97AC5FB08B25CA76774446C812071967AECBA64E83B9F3CA5885736E5E01
                    D362429C1463C00F586EB7DBAE459B3F075FC7994CA644CA8950D6925AEBB034
                    79732BB510007B8A9D3D5FE91927EA741D3A9DBE8AF010D584EEDC510117EB3A
                    67D5EB1A9A868A07D6E0EECE4F4B25BE819EA9EDF2F272A8AEAA82ACEC6CB8F5
                    B6DB50104C67EA004803ECDDBBA7F5BF376CA08EAD70F8F021CE35F8EC250906
                    35AAFCC1FD0FC4BA3A3A76B1923385F28CFBF855AD191706856CC11530884E29
                    1A3F0156DF7C0B2E981EAC363BE8131359338EBEBEBEE9C996A49BF2727396DA
                    EDF602FC39855C402AD2219795B8115E9FB772CFEEDD0907F77F5A4AA1EB9B6F
                    B905269516C733A29CF0DAD2DCE27D6DFDAB221D17488DBB2EE6758E2050B5CE
                    10EDC29F11F54F41D46BA3A654540A3EA9A404766DAF80CAC30760D9B2A55090
                    97C3001DD9F88E8E0ED8B0E1ADB53BB76F6FC577BF235730832C146CF9F84396
                    13F8DE9D77D199515F7A8212A7F13C0271CECFD463E9BEFB1F64390C137A2916
                    FC37959E79DC6E1B7A31ABACD6E41B53ECB6CB10FCA5E9F43A3A6F98FD3FD540
                    A2364419F0BF8FBFFA3BE2A4FD8D8DF54B0F1EF8745D3426EAF30BF361D1D50B
                    59BD027580331ACDA801CA0BDF7B6FD346FCDA7FC1D175BE6771510461C88596
                    A2F836DAED16B479EBD0359A4DBD8A66CE9E053FB6FD049E7BEE5928DFB60D16
                    962D60FC40CA1934363442DBA9F66C7CE71F217ECEC0A7C0CA50065D2062EB89
                    EF6EFA07D31AAB6E5ACD10F6D0C2DCF19FE35E831A965F732D2C5DB60CFEF6D7
                    3760DFDE3DE7BCE7A11F3D0A7F5AF73B78ECA74F704CA1639D58882710150415
                    0AFB95A8EEBF939995B124C5664F35990C461624A2025862437B99DBBA1FEFF3
                    6F78071FA1DBD866329A3D7B76EF828A6D5B2B701ED538E6ECDEB913BAEEE882
                    A9E969100AAB5921CBE9D37DC46E598AE3D738D648F9B9019846BC28827001A9
                    4F01D5F45E54852BA2EDB167C598783B2276554141213CF9E42FE0E8D1A3A842
                    9D60494E6639FDFEB002F4C674FCFCE66C518CDE8AEF3FCA015A824C0062323F
                    9DCC8FB81935031583DCB06AF51942CB6017999FCCCC74E6A11040A30EA73417
                    03E714DE73CF7DCCCF8F57228799C0988C7AC685FCC52F7FC5B39F0A7E4CA0A7
                    50AD52AEB6582CB7E415E44D484EB2E834E8D6D0C15BD49185EEC3EDF1B6A326
                    D8849ED006FCBE83B8D841C44E029DD1B4637B39BCB6FE25C21053F0894EA4FB
                    EB625D52EAE1F2D9B399FB48C49618982035AD00BA3A9B97E39FCCC24147FBEA
                    64A075D484E1AB0882548AD68D13FEFE8E9D3B1C2525258F4D183F5E491DCCE6
                    CD9B071DDD3D8CE0919C6C86EEEE3E309933212D2D07CD444B19BE7B1CD70C06
                    3E6981FBF6C085212A69A8AD9F6C81EDE5DB60F9B5D741D98285676A11E81911
                    E99542B76A969D54CA0A61401ED0016D92194CACF620C6CAEA9820980CACAED1
                    6030AAD18418519816236ABF2D2B73C242EAD588A64925C44F9B81AE6E7F1477
                    7ED01F08ECC3F9BE85DAE27D44FE547C221010260D545B5B032FFC711D332B28
                    C4B8DB95D7A1D2B3D0ED58ED0570ACF62473A3ED686E7A9A1D8823225034611A
                    7477B518F173AEE49B23893F87D0F9125AA3220894AA1DA2696021DF04F40E7E
                    F79B5F09E846B5E517148B6BBEB706A64D9F8C93B5B18C20F115684777F5F840
                    6F4A015B4A067476B4A4E3BBC79393C247880770147CE2F21D21E529E0DD7736
                    3277F0C61B6F62E04A62120F5CFCA16447D1F2A8A80126BEBF1435D6AAB4D4D4
                    6FDBEDD6226AE4218AF17E4D1D1D5DA2DBEB41390838FAFB83EF4422C2DFD063
                    39226F61439A8832ACADAD2DF0FCEF9E937F650E28351321A684246B265C3EEF
                    7A70B4BBA1B1BE19F2176723E8F483A8D44356561E6930D436E1347C4F1ADF18
                    2A59CC23267F16A32608CF3CF5C4976126D098D2D2DCA6F8EDEF5F847FBB6B0D
                    5C317B32226D2B63F9122F21260AACE7416F6725EB5C876FC8944DD2CD7F0E4A
                    1F763E244FC0936A2368975E68CA9E53C368E7A7227E58888BFA9DCCCC4C3A30
                    4B497C054A0E9157803B3FE0F1FA7B82C1C0EE7038F4564282F6037C4FE45C12
                    8F064EA0CFFFEB679E1AECEB12158A04BDA84A44C0990305E38A19CE3950791C
                    E6CC9901A73BBDE84E66824641FD13C322D702A9FCBD212E0CA3464A18A61033
                    BBDF7EB53105826125D4D4B9202B4740D5D70AA1F454D6F27EF992D970BCAE05
                    0EEC65ECAC4E6E0BAD3CA3A7900570A2E753878C5144DD48A2D10B5D7C0DBE27
                    1D771E25BF56D9ACD69B10FDA792FF2EC47B1C93CD8F518B61BFCF571F094736
                    A12BF8366A8BD681BDA9E9F3C80C90466C3ED134981028F9FD07504BF9558929
                    A04C3082D1AC87FC7105E0F776C1968A6A6838E982D24945F0CF0F77027F0627
                    B936F0704D208C9659182E416049604AD3522F138B3D1FC282087A4B0104FA3B
                    511DF632722A751C9931BD18EEB8630DBCF4622878B4FA88640EFC3CC2E7E31A
                    213A30A9C3DC3A1480471F7B1C261697323AD950B810425440FF5E4369EF857A
                    83619535D9328B3C18BAC8E523FA3C823F17DAFDE6607FFF56B4FBEFA148966B
                    35DA78B67310A1A2D63847ABAB98D7F087E77F3B98102470C1F6C684609D35D9
                    362F885E422F6A9AE933A6825A9383E6B29771241B6BF7C1CEED74180B509FC5
                    533CAC2DE539C29FB7292E3541507215469FB33B1AEA6EB2D9E78D6B3FDD031D
                    EDED70D9B429D01F74427B5723AADD16568ABE72E50A983061C2CC175EF853E7
                    279B3F5A8F6E57030787922044B84938631A2C9664B8F7BE0760C6CCCBC1E138
                    754ED79381892F955A5D8CDEC24CA3497F8DD562B9362929C9A066A1DE306BB3
                    43BD95BD5E5F3DE28C4A747F3F8AC6A2EF23C8E8A250F8F94A4E490888E9545F
                    7F1CD63DFF1BC6721AA012951CEF50B62839AEF1228763FD6DA76D59051955D5
                    8D90949C0A45E3F321E8F7C09143BB61E35B7F46ECE3DFC1F90E1E4E7A717341
                    888CA646500C64C25C207356CD7780296E13E1E68CFC198F174C9C6FF1790498
                    FBAD59A846557014276DD6FBE0DBDF5E094545E3A9A31803586FBEF9E63B7FDF
                    F0D6CFCD26F3A1B676871AD5B2289BBC48AAFBAA858B50854E82F9572E60DC45
                    DAC9F182D3307BFC94A5A342119552956B4E324F47377611BE2EC1EF28D6B183
                    3AE3E727BADC2E70BBBCDD3EBFBF4A1022E574E8369D784BCDAF29354E5F1C3F
                    751D55BF3A81E10589884B80944E943F5259C900216983F368021D1782748E81
                    EC38E6178CBF6C49A2A930C5E512145AAD1A029E0EE86A3B8CE6314442B09E9B
                    05BF8CF012194C2B0E375764B80441C905C1209B3CE5FA57D0C405D19CD5D9E1
                    D28643281191B61EDCAB07AF2C5BE0BEFDF6EFAE282E29B6B35E83F8A0FFB975
                    EBEEBAE3C71FFFDF8D6F97E7E5E5C3BEBDBBD9879310DCF7EF0F320E616F4F37
                    6B6641E681D52E006BD44DCCE4441484D9A801CA4C06C3D526B369BEC9684820
                    3792BAABBAD8F13CD4F6C6BB3FD81FDA198D08E51121BAC360D0F5B1763C547E
                    AF49F85C41A0E743E73B1E3D52C54CC120B10C05D78A3AEEFAA57221C8E23F1B
                    E9D9A069CA51AAF4269FC785521469E101359A6C2F07875E59BA5B1CE9D8C170
                    0B82A40D6CDCF5498138FB274BAD56642B95095A44DD34410A2B1FA3D7E2E292
                    C9FFBAE6CE1FCE997339652C59C875CFEEBD2D1515154F97944EFE4BF9B64F58
                    3C203F3F1FAE4321703A7BD99751BC9E54336534C5A8508480ED2AA3D1B808DD
                    BD394693B180E2089148FC14145A7CB7CBDD1A08F657E0EFCA4531B6271A156B
                    59438B688C752419AA2090C6A93D5603EB5F7929DE1AF7FC1B423209921B48CF
                    8428EF5A59EA3BC8179E6872DD7CE1FBF9EF05999B38E201A4911004039FB095
                    EF081D9F48984F52028461AEFA7A3233B3C6DFB4FAE687CBCA16DC43D472D20E
                    DD3DBDC1AA23D52F4785D813E81E3AA9B51E7D0C3F268FB29B065CF425BA44ED
                    5283413FCB64324FA5E379A8B8858483228AE8F685BD7E7F05A27E045F8A9D91
                    48B812D57A80543B3B56872AB62F5010C853F9D31F9EA7C2952FD28C3AFE1C92
                    F9AB813F1FE95904380E9080B1C4808E0CC677186D41F82A6051B2E761BED022
                    7F55CAA28321D990881ECAF6F6B6869AA3471F2D2E29AD4355FD24EE76B3D56A
                    D14D9D32E9FE63B5F51310613FA6522A0EF1739267E97589D75A93AD0BE274B7
                    44DA6DCC4450E713C207AC5197201CF705020F20DA3FAE10E1143BBD4D50B04E
                    667A9D9E698A0B6D3A21F119A8A4EEF31C1459F0272023BC487E6754B6294232
                    F24B6434730923ED3544F9E444FEAA18101D1464AF208F90E103F6C462E2BABE
                    BEBE9A507FF0F7E148A4944E86993675CAD2BAFAFA340478EF64A4A71526275B
                    E66AB5DA7C34172A76A612BA8E3DBCE0955AD55008DB68305099D916A7CBBD85
                    763DF56A2210480BB86B47056CAFD8060B10745E597615D31E748683C7E3FEC2
                    7804E5293EFEF00338B07FDF509E4358C6739082415199900C24BD5C1202309C
                    1A417A0D0DC8EDCBFF5F1C3C55A188E1AEDBE2F1FA6E88088EA7114FACA67677
                    934A8BA73A1CEDA54A954A433B92DADE53632B62F9505F236A809D92628FB7A2
                    53AAA938366A341A76178E2B84A6C64678E4470F33179308A44EF4DF491B9C3E
                    DD0E9B3FFA80990122C2DE75F7FD90929ACACC01B5C11FF4E1A0CB495C014A9E
                    0DE159083038E145FE0C2EA9C51F4E8D30D804C5F38587CF1B9304688C4484BB
                    DBDB3BA946F0BFD2D3D3157979B91AEAC3D4D4D40C540575EC6835EBDB38E78A
                    B9B06CF952467D933E801ACCBCF2F24B0D9B3FFE98CE4062DC878117F104BDB2
                    05FFCD734FB33238D2203F7EFC89786B5F45BC89A7DC340CBDEFE419C2CB977A
                    0E5F774180F3A0DC0B9E3C6A0727AAEAA7BB7B9CD5881B9E417CA0ED38DD790C
                    55F9AB9B3F7ADF7CE8E0FE5FA20760A7C6124B165F0D469B817909445BABAFAF
                    D3BCB6FED53F84C3E13BF0A38E0F6501E4D1C9FFFCC95AF60E22D8AC7DECA720
                    84852F5B89FC959FC3D75D1086EB0AE1E6DEE8F5FA360783417534160B84E8D8
                    B760D086586215FEFFB5FBD15E9F3AD5CA5AD41148243BDEDEDEA140A1988BFF
                    FF328E251CB4A90600BAF32E8AA4298888FAC00FEE667F3865CA5478F891B50C
                    877C132EF525784F04FCBCE4D66D2F2F87375E7F8554B49907AB98EAAFACAC42
                    ADB0987110A8DED1E7172135AD083A3B1BA7E11FAFC13F7B81FBF5E28004CE17
                    46EAA4A8E1C1039FC2EDDFB979D46B102FD6A5BC685F4C7D8A58230C69C41B6E
                    B283B57154546C85D75F7B992F84622528D4F3E3004E07C7EB1C8C354D076610
                    23D8E1E881BC822960329AC9975FC6835C66FEAAE3419D2F38B1E38B7DED318D
                    30024240FE3F2F3C61358B2404227F25A4FEEACB7F96FE1C1752350B442A76D5
                    C0F4D9D78127A0829A637570CDF285E80606C1ED8980DD9E01365B9AC2EB7553
                    78B75016C851C3D95E06979CDBF68D1604CAE713877F883C7E2BA8D4A9206AC1
                    68B143D1C4E9EC6CC7BDFB6B60E9E2F9681642108A6A20232B0B6AAA63924F6F
                    E78BEEE1C2A094C538C6AE4B451006AB67F89C0BFD726554A131834E67035392
                    198A32F2A0ABBD0EAA8F35437B67082CD634B09A43D0DDDD418B4F3CC844385B
                    B22E5CAA419C31B038B44BC517CE0BD188C368B5B33636D40BA9A08048B016D8
                    57ED44AC1041845F00EF6D7C050254661DCFEAD1A2FBF990FA188C1AC9630C2C
                    0EEF7D9190F22456E4B05A11F259EC995053D3042E9713C615E542466636245B
                    F4F0E99E6DF0C9665633F3318EC3DC245059BB0BCEB27DC684E0F3E2385F9198
                    32ECF7C387C4F4A18C2665F332556AEDF78AA72E5AA5D1E582D96C8589C5E321
                    2604A1B66A37ECD9B1091585E77DEE36121FD20D9FED6A32EAD9BD4BFD1ACE34
                    F44809829ADB7889E491C5C15FA64EA72F4BCF9E38C9174C48F57A839AFE40AF
                    0F226D4475A31AF277393E08722F41DED462ECFA9A09C240DE9F44F248E5BFA3
                    FCBE45ABD5EA2311EA3A23D0C2D7E2A8972D7E70802918D3045F538D2071FF2C
                    32D360E041A128B7F952371337C4295E527E5F1280E8D8527FFD0541C517DDC8
                    B5805EE6DD44E02CE1454EF41871B2E798205C3C8F2181BF4A67DD0D4676198B
                    0FFC3F16044918A49A09E9929360C4B1C5FF6608826426CEB9FFB1251C254118
                    BBBE99D798208C5D638230768D09C2D835260863D760D7FF093000DF9C3AC9EF
                    D8ABAF0000000049454E44AE426082}
                  Version = '1.1.1.0'
                  Visible = False
                  ExplicitLeft = 163
                  ExplicitTop = 18
                  ExplicitWidth = 232
                  ExplicitHeight = 232
                end
                object img_produto: TAdvOfficeImage
                  Left = 0
                  Top = 0
                  Width = 569
                  Height = 304
                  Align = alClient
                  Picture.Data = {
                    89504E470D0A1A0A0000000D49484452000000820000009308060000004289C1
                    C30000001974455874536F6674776172650041646F626520496D616765526561
                    647971C9653C0000032269545874584D4C3A636F6D2E61646F62652E786D7000
                    000000003C3F787061636B657420626567696E3D22EFBBBF222069643D225735
                    4D304D7043656869487A7265537A4E54637A6B633964223F3E203C783A786D70
                    6D65746120786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A
                    786D70746B3D2241646F626520584D5020436F726520352E332D633031312036
                    362E3134353636312C20323031322F30322F30362D31343A35363A3237202020
                    2020202020223E203C7264663A52444620786D6C6E733A7264663D2268747470
                    3A2F2F7777772E77332E6F72672F313939392F30322F32322D7264662D73796E
                    7461782D6E7323223E203C7264663A4465736372697074696F6E207264663A61
                    626F75743D222220786D6C6E733A786D703D22687474703A2F2F6E732E61646F
                    62652E636F6D2F7861702F312E302F2220786D6C6E733A786D704D4D3D226874
                    74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F6D6D2F222078
                    6D6C6E733A73745265663D22687474703A2F2F6E732E61646F62652E636F6D2F
                    7861702F312E302F73547970652F5265736F75726365526566232220786D703A
                    43726561746F72546F6F6C3D2241646F62652050686F746F73686F7020435336
                    202857696E646F7773292220786D704D4D3A496E7374616E636549443D22786D
                    702E6969643A3830313231354233323437303131453238363741454333413046
                    4233353339412220786D704D4D3A446F63756D656E7449443D22786D702E6469
                    643A383031323135423432343730313145323836374145433341304642333533
                    3941223E203C786D704D4D3A4465726976656446726F6D2073745265663A696E
                    7374616E636549443D22786D702E6969643A3830313231354231323437303131
                    453238363741454333413046423335333941222073745265663A646F63756D65
                    6E7449443D22786D702E6469643A383031323135423232343730313145323836
                    3741454333413046423335333941222F3E203C2F7264663A4465736372697074
                    696F6E3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78
                    7061636B657420656E643D2272223F3E7DFF8DB900006E834944415478DAECBD
                    079CA455953E7CDFCA39E7EEEA3C1D66A67BF2000E203908084A46928BB2A0AC
                    ECAA6BDE35FCBF755D75DD35ACB84690554414242B39E7094C9EEE9EE9DC9573
                    AE37D4F79C5B55E308435009AB3BEDAF1CBABAEA0DF79EF39CE739F7DCF30A8D
                    46831DFA39F4231C3284433F870CE1D0CF214338F4F31A86F0DF3FF8094BA6D2
                    2C93C930B546CD2451661AAD86317C4E1004A652ABF09F02D3E23D491499DEA0
                    67B55A9D994D2656AD5599C16064A5728959F03BBDAFD5EA982CCB38968AA9F0
                    7D596930BD4E87CFD6F8778AA512339BCD2C9BCD3251ACE3332A264912CEA9C5
                    D531268B12531A0AC3FF3119D7A0C80A3399718E529935F09E244B4CAD563399
                    BEA3D1B03AAE498BDF259CD3AC67ECC588836D8E7631A3567C63060CF7502AD5
                    74B8AE735C76C3656EAFC5249B1C3FED1A5B723D6B08A2462BB07219F72BE05A
                    71D92A95804B6FB07AB581EB1430660AD319D4AC5691995853985AABC23D36DF
                    2B1724A6C23930BC0CC380DB1758B584FBD360DCEA8CE98C2A26561526D615DC
                    2B1DABC1B44635AB16706F984601C7A71F15FEA5E3D3F9148C9156AFC6F1F42C
                    1B9D67567D82E9BD01F6C2774EFC83FBD2FC355B3999B856A5308BAEC627506A
                    6040843F07010572083DE6F55F36ACEFFA870D470FA9B22A03CBD784C392997A
                    B7AC0837E16FB3781598F09735567FD5865097D5ACDF5D60A3FE047B36D1CFF6
                    957A995E05A3607F9A3190776732E5CBD6AEE8BCFA8A4B0F57DDB135CD820135
                    1BEB32693345E9B3B1B4F8D1445ABC5DD40A37297263231021A152B1FA2143F8
                    5FF0A300676B928699D475E6D1A6595D303059D0B5F0E28FC2028438C5ACD1E9
                    AEBCE2DC31D3DEF9085B58AC328B5DCB62B90AB31B356C30AC3786FDBA0B9259
                    E98258AABE2D9B956E96E4C6EF64B9B10F28913E64086FF34F4DD1B025F6381B
                    B14CB0CDD2D12CA3EE605A56E1BCE7751B02D0209D2E9DB57A6570C061AAB2D9
                    ED2FB2ECAC9FED33EA58C16F6216B386D94C1AE6346959974FC33A3CDAD17451
                    1E8D26EA9F49A6C53B8A4CFE8D2235360129F6927D1E3284B78311E32529205A
                    8A91D9840CD32822370605F15DDDA8BD4E8368086A83E1DDCB06ECF65A6A9E8D
                    F40599794B854D6C8AB37C978DB9FD4666B6EB58DAA265563319848EB92C1AE6
                    B19A4CC590724122235E00A3D898C9887701211EC46B1B3870F69021BC0D3F12
                    D3B12EED146BD48B2CA259CD2A0A26CE3444349BA9E5EA6BD841A3C76ED62D37
                    D41798D7E9632693815D7D66813DBAA9C01EDB16DD9A9E350E390226BDCD6B62
                    569781A56D22B302211CFCA5657D211DEBF0E9D6A472E29A48BCFEA1644ABC2F
                    9F97EE93A4C633503FE3870CE12D4506C8CD869689CCC2C218FB4AA5C8AC8606
                    A42C60DFB316D20D7F97442EDB5EAA3EF0B3DA2D25BC031E0DB3DBAC2C932FB2
                    D5430E16B0551E7C66BBF8D148445A55CD954FCD2C1A8E32BB8C219B17E8E336
                    B2B443CF6C3C6C6861103AE6776B98CFA9F5E64BF2FB6229F17D9158FD996C56
                    7CA45E571E90A5C6139080B5B74371FC9F328403A756647A260A4616AAED6485
                    6285598C8C8992CC3CC1106B287F18C2D508E9F3A2CDB9DC271947FBC32C5F2C
                    31A3D1C06617D3C57432FDA92397B3ADBF7C5CB5B5A1C8BFAAE42BEB6BC5EA31
                    F944E92493DD70B8C563640E1884D56960566B9D1B851D4661C5CB6935B07050
                    7F3838C4E10BB1FAFB1289FAD39582F480586BDCAFD109D3944B3864086F55B8
                    10F44C1124E6C96D6395BAC806FB8C3C51264A52CB3105CEED9CFA9CFBA8E111
                    830C1BE189AB5A9DC562B1AF31457E6175BFC07EF11827942541101E1698F2B0
                    54937E564C96D697B2D59372D1D2BBCC0EBDDBE631319B4BCFAC7642091D0F1B
                    C427BA033A16F4E9C2999C145E4CD4CE8C46EB5B0A79E9895A51F92DC8E5036A
                    B570C810DEAA1F596D40D868B0B9E9694E1EBBBABB79B69232AFB59ABC64798F
                    E32C17823EE420D36A556C66666E63AD56FB2FA341C3AA1020258417ABA9D10C
                    2394815509E3308A714592EFAAE46BDFA995EAC7E69395F79AECFAF5561710C2
                    6D6076840DAB550BE9A9E548E17168C03FCCFAFEB0F1B048B27ED862A4765E32
                    51DF5AC9CBF7824BDCAE6BB0F94386F056040CD0F8523ECB53E17A8D009501C4
                    1025E6B65BAE0E063BD653AA5BA7D30009128D4432F9691849AA5EAF731271C9
                    D12A76E74613D3EB5EC24B04210BA4780ED1E605B92EDF5C48955757F2F5F7E4
                    E2A5B3D276BDC58AB06177124AE8980D8AC36640D880710D7619584FC0104EE4
                    C4F07CA4F6CE58B4FEA1524E7A4892945B402E1F636F30913864082FCB1EAA79
                    0691FE2526A1D2684E0B7674BC5FA7D7A968CDA5007E3037BFF02D49921EE3A1
                    450242C068C26E81A3027D57AF7BB91C856D2930886960C6B42C2B0F950AB56F
                    544BF553F389F2F9299B6ECCEA84418060DA1136AC24418D20987885402E832E
                    AD25D32D2D5D8CD78723D1FAD9E9B8F862352FDE0A83B815874E1D328437C310
                    3063957285EDD9330ED8B6740D0F0FFFBDC36E73E8B45A1E2AA6A7677796CBE5
                    6FABD4AA9A568309B35A59B9546223031EF6CFE726D8DEA8C06E7EDACC2C2690
                    4C35378083194516116773436EEC9098F22318C311A54CF5C24CA4485CC26227
                    94C0CB66D3E138442EF142E870F59B54BD9D8620D44670315A3B2A11153F55CC
                    88F7E058370A0DF6DC214378C343042465B562F0FBBDD78C0C0F1E4BBF9730D9
                    8B91A81C8F273E0534D8EBF37AF9CAA74AA5E233ABD791072B6CAC5B60439D55
                    66D229ECB9292BBB6B236366D32BEAD93ABE1A03F4DC0E52785FB950EBAE166B
                    E7E4E3A58B1256DDB0D56D624E92A17603B30025C820AC4635EB0FE959D8A733
                    27BAA5BE48BC7E7574B1764926296EACE4A41B713977E0C8E94386F0E71900F9
                    AF4DAD569F6D365BAEC56B19605C1D0CFA8112655AFA9E73B99C423010306432
                    997AA1282ABFFF2E08A7D24400B3A1C179C3E1FD6576D29886DDF1A2993DBE5D
                    6446BDC06A325F517FB99E65AC8CD72EBCBE2C8AF237EBC9D2F1C54CE5B2F4A2
                    EE9D66A7C1EA709BD40E0FA104858DA651045D1A167069D5D94E833D9A168F03
                    B93C3A15A97DA994116F8191FD0F8EB5132FF1F5A4B40F1942CB00F0726AB5DA
                    93FD81C0C7D7AD5FB79A6A1C262727D9CE5DBB58241261017F802D5B3AD2B374
                    64E8F685C5C8F3308AEB359AEA5DF85E1CAFEA1F1EAF39B504167559C54E1DAD
                    B075C1147321FEDFF0B49D254B5A56049F500EC8561D2860F0CAC3A46E5364E5
                    B64ABE3A502BD62ECEC50A67C72CC61EBBC768219470505E82A3848E870DB7CD
                    C07A833A4D2C6B0C47A2B58FC617EAD7E412F5C72B25E9469CE3011C8FB844F5
                    90211CE4470195C7CBADD7EB4FF4FABC7FB76AE5AA77843A822C118FB389F109
                    52080DB546232C2C461B894452999DB3AB7D7E0F0B0502EBBAC29DEB92C9D4E7
                    66E6E67E29CBF22D38CE2EBCD23C54BCCCD09AAF4A5DC5DE339A607EB7837DFB
                    213D2B571522A32C9157F18214D5C185C02494C717E4BAF2AFC564F1B452A67C
                    716A5E7398C56DF2387D269DD363024FD1310725AACC5AD61740D8F0E858AA5B
                    D2216C1C1F8DD48ECF26A4855CBAFE8B860C82D9684CE098894386D03200C479
                    9FD1683CC2ED767F68D9B2A52751DE209FCBB28D2F6C6498E078BE5058185DB6
                    B473D58A15DE3DE3E3D53DE31377A7B3194F3A931A9B9F5B70B93D2ED6190A05
                    D7AC5E756DB158BC766E76E16E87C3FEF372A5F294284A8B6A95EAA075089222
                    B0725D6017AF89B37AADCA545A0BBBE1590753A9356C21AB661A152149839166
                    790958D4A03A6E05C1BCB552A88D568BE2059948E13493C3D065F7989C300AE6
                    044A38B8516859C0A9E1AFF28085BDB039D6919D9FFA58B75FF9584D303EA5A8
                    84AB71BCADFF670D810C00DEEB01026CF0783C570E0E0EBEAB2B1C6635B1CEB6
                    6FDBC6128944A654AE6CA9562AFF212BCAAC4EA7BF49AD56798D06836CB7DABE
                    924C253756CAD5F38DC6C239E94C76756431D6CB394328C88686969CD63FD07B
                    DA6234B67B6E6EFE66288FDF55AA9571AD56933A184AD424155E6A66D234D8FB
                    D7CE31451F60373CA1625683C2F2753D8C42CBF41A89690FAA3A846D90A9DBA4
                    BAF8957C543EBD982C9F959CD1AEB17A4CBD300895D36BE23234847F0B99029B
                    DBB48B9DBADECBF4C100DBBE507D472653F9020EF3DEFF7386D0420097CE6038
                    DAED725E383030705E77B813E44E615353532C1A8D950AC5E233F55AED464996
                    6FD2EB74759DA0024F50F15C61BD5E1721256D85628189A27873A552A5D7BA42
                    A1F89E6C2E7F5C2C1E5F0634B004FC7E1608F8874301FFE713C9D4C7E7E7E76F
                    9744F94E1C7B23CE3F71B042617AA7588334348BECE2953106F1C166F20EF6EC
                    9C95E5AB7AB698D5301DAC41AB92113A1A2F358882A0166E02D9B8A95AA8ADAD
                    16EA67A517F3C70125469C7E8B23DD61659B9E9860A78F0659F7680FBB737392
                    D5AA7566B26897FE9F0A0DCDF470D5AAD71B8F0109BCB0AFAFF782704707952F
                    420A46D8FCC2623D9F2F3C2989E24D80F39B743A6D9186DA62B1902C5451AD2C
                    E515604802AD4A9A4D46562812AFE005B3CFC3289E17252954AD96DF85E3BC2B
                    994CAF9B9D9DEBF478DCCCE7F39B972F5B7651219FBF281A4F3CA3D168EEAC54
                    CA8FE23B9B715DE597192B044BA90E280752788C4576C99A12DB93F4B0672765
                    40BDC026525696296B8020F2CB7313BCD85578815E724DF6E462C5534BA9CA71
                    FB7644DFE9B3B0DEF5EB7BD9033365A635909435500171F6FF842134F300559D
                    C9647C575757F7790303FD67068341132143241665D14854C966B38F91F7D76A
                    F5DB0C7A7D82E09BFE4E5662329918264D002250C4E6438D9022D8ED76A6D168
                    799577CB233101AA457CF787501137E058C7805B9C88B071FAE26274C4E57430
                    3F50A2A72B7C784747F0F0542A3D1F8FC51FAED5C47B80324FE298732F4F2D34
                    8B6852250D0BDB0BCCD19FE40B544E7D99951533DB3C6F67F94A0328A146E890
                    0F16369282C2A014941B8BA5EA57CE38A2E7931918990C3471225CF06A70A5F1
                    F05FB9211002D4544CAB3BADAFBFEFBCE1A1E1D301D58E464386048C71142814
                    0AB562B1F479B15EBFC96C31CF924C6C1B0FA595090D18FF6F954EA5561B6966
                    6020B41E25033578097E2A9DE646F392D4B4A8341AF78BD5DAFD2E872363B359
                    BF1C8B27D2D158BCE176835B0225DC4E6727FEF3927C3E7F8EDD617B064AE43E
                    10CD0770DC170E927D6422A46759D430B9A2634B5C69E6738BCC205498DA6067
                    A562993D39EBE68663D02820A0EAE676032D26D5A201F96CE87C5695AFABDBC9
                    B2528319F46A66D4A961648DB42C2BD7FFD51A825817A9B8F494BEBEBE4B4797
                    2F3B3E180AF8E8BD5834C666E7E6583A93D983181EB45AAD5A0CFECF31D173B4
                    07A21DB709119CF06083C1C0F761A8040A0E2A3E3E0D45A9B61233981C89391C
                    0E96C964F77F978C8216ACAC30224A40C108F481408065B3F93BC11F7E0A8E70
                    423C91381388B2D4E3F6E0FB36636F4FF7B19D1D1DC76632994BE3081DD95CEE
                    2EA0C4BD5AADAE427B405E6A147510CB5C55C386DD29D61536B0854812BCA1C2
                    24C1C23602254E1F89F19C854A2330A391B1B9A26EB4AC0BBE5367373149ADE2
                    6B171023B4C2FAF94245DCF3576708B41906137A7CB8ABF3D2D5AB561DD7899F
                    5AADC2161716D8DCDC0249C14DF8FBD7483B5B2DD66FD5EAB511207EA0A152E6
                    A80085B43B4D3C4DA6C96C62B22473A340BC6DA8889D35F8C61A4C33457181E7
                    E88C88B34C70B25432C58B581C6E37370ADAB823C98A1D283348E9674CE00E4C
                    EA4340A947E1AD37552AB50DE974FA4C9BD572B2D3E5125C4E27B8846704A835
                    02E2795A3C1EDFBAB010F91DD4C96D40ABBD1AB5F66561A326ABC12504AE3A46
                    3C19C47D89D9D479B63C5007E4370BF5F58D2AAECB7176CE35D0AD31EA998948
                    2624252EF517C59AF493BFAA84124D5CB5523D2C140A5CB966CDEA937A7A7A3A
                    C10BD8CCEC0C035B67C9447A0734FD7598E4DF69B5EA490CA35F94C52AE3615D
                    6D1584261F68C7FA70670734BC863584E68E2EB58AA056D5689D4BA5C00EF8E6
                    189E4A96E1FD666E047A1D061A24B21D5ED48A6CD16AB5615ABE0691CCF21D48
                    8220C338B6417C6CABD5E43BE2B5DA28D4C6BB63D1F8D90EA7C3EF76B998CD66
                    F30D0E2E39A1BFBFFFC87822F9FEB9F9F92717172237835C3E08B97BD03120F9
                    89A8C0C2B63C8CC3C25189FFA80D470906DB07CC0E939A10420F6B570BDA1D30
                    82CF4B72A34435157FF186D05402B531E8F77F58B962E5F14343036132009281
                    0B0B8B84009330806F03A26FC767679B393D81A7911559A9130F80B79BA93CBD
                    6D0884000B4010096830B27498A3425DAC6B81087A3A5F0B10600B32E8878E6F
                    BBA31F1DFE9BB6DFD176BDF6B5C108AC7A9FAF07DF2FE26B315C277849713F11
                    C5311729E1546F884F0A42F55BC5F905A88DE485308475769B8DC2932114F42F
                    EDEEEE5C0A2572C6BE7D533BA7A7666E2D954ABFA0FA8797B3A206DFC165D1EB
                    F8F961F84EA341F3B71E9DCD9B0027A8D31643419070EA4F9B75EA71BDD3C0FC
                    76FD5FB621C03B8631607FBF72C5D869CB968D74E077617C7C822D2C46182077
                    16CCFDDB18F09BF1D108450DFA0E2686C77D8490126D256CD5019A0FD4F4A964
                    92AB015A7EDEBE7D07FF4E7F7F1F6882A0E6832BC9794C62C98D1050AFD758A9
                    54D93FF1F4C275EC4729C844AB56A709E13333F83D46C8427B3BA9804592AA1C
                    E0B9DA1084225E7B30717B61D83F8F4463C794CBE5AF802BF4988C7118848B8C
                    22B066F5CAC0E8F2E587CDCCCE5D3B39B9F75E51CCFC1407D97810B5C0AF45A3
                    D59E647358CFB1584CAC045151114114D5C2BFE023F7F05A09B5F0B2E2DCBF18
                    435064B9C764327F6CE5AA95E72E5FBECC839B56EFDDBB0F216011ECB9C82AD5
                    EA1E10B253F0FE02A65754B50645109A93A500CA6559AA6142EB544380F7AC40
                    0C41686D84248EA003C4D360D284B53600378687075B93AB0E02350E53A9553B
                    048A292FF92164A010A0F050A1F369F0411C2781D0106FB426893614AB0EBE98
                    20016F62F8DE33C1A09F30A338B977DFFD8964AADB61B7AFB6DBAD3008A7B9BF
                    B77BC99281BE5E18CCA53088A7700F37D0E4523E8A8E4F453330CB3E9D56778D
                    D9EAD207BC560629C116A4FA2FC486F1BB181179FFBA07FBCB0A0DE48D6183C1
                    F891BEBEA5EF1B1E1EF2E8F55AD5ECEC1CC1B80209378398BB677868F058B06E
                    3286797C5E82D11C20E954ADC9150911EA80FE22BD87970D9342BA51A230F012
                    CF02920A4722107C12906012251193A871F6F6F57C17837855B552F901FEFE2B
                    7C34DB5A296C65741A7C8021393B05A0000C2109CFA35A839651BEDC08E8DC1A
                    B5A6E9C96A4D97D964EE82BCDD89CF7EA652A94CE37506D4CE15308A0D209826
                    A80E0D50C971E48623DE95CDE64E80C1EC9B9D9DBD099FFF39CE33A3D569DFE3
                    74D837385C76A606AF70C466B6640CECCBB2C910D728CA5FD6323406458B1B0B
                    00CEFFB62BDCF981FE817EBF11D01E8D44D9FCC2829CC9E692889737D7AAD5AF
                    804405B45ACD4380722F881DB9D3828E4ACC380A88AC6D14E405A42E24492C35
                    93406A18825A9D4AA5240B60BB05F15618C8E1907E1F59BA74E4740A278D560E
                    18E75252A98C2A18F0AFE9EBEB59932F943E01D4B8B156ABFF0CDF9BC331ABF0
                    5E96CBE7010EBA106BA6A593B2D228B30300444BA48FA3139043AFDF6F1C4AA3
                    A1D619B4A3B80FA1582AC760A453B0AC1A48C92D38C7AFC08156E472F9CB52E9
                    EC7B61143EBBCD6A70B95CBA552BC7E01F835F9C9E9EF9085ECF602C9643DA0A
                    16B385A552692992CA7E5FA827B6A9A70F5217BFEEB2FF9D864006808908180D
                    C6CB43A1D0D5D0D94162E3C9749AEDDCB15349A7338962B97CBB58AF7F1DF26F
                    A2E9D9023C5652743AAD0DEF79709885FD37A656FF41D28788208CA5C82744AD
                    B64894416E1A801D937704B8C7B5E1659DA7E05F4A1845A3B158A11F784C8925
                    9C63727E21F278221E5FEBF3799778BDDEBEEEAEF0E7CBE5CA4730E037D7B4F5
                    1B71AEEDF89C80D013E684B65E5F68171BB411814ADD7C1E37EFE170E0B5E1EF
                    B855D37ABE545DA9EC86A1D40ED8864772650B3EB305E4F15F8018E7809C9E97
                    4CA5965A2D16B7DDEE004A38DDF942E1343A26390DAE8B4175DC03A3BA4DABD1
                    30A521BFE6F8BFED86C0A5992C878D46E379019FF7EACE7067AFA9D53863EFBE
                    BD64D90BA572F91E10AA6FC28B77B44911792C86320999560041B343D20571B8
                    178918B437A890B134DA5522D4B8A2D1C853A30F70023398BC1F5EB9DC6AB55E
                    DDD9397C06D51EE6F3F9D89EF1F1A770EEAF3B1D2E1B8E7B3B2D4CE138698341
                    FFD90C5C1E03FE4158C9052EB77BC8E7F3B93A3B3BAE0281BC0A6AE566DCC3EF
                    A02A8661101472F6D16979AD2390892608F7C1E7B75C2AF3CD346D4381119BFC
                    3EFF7A280D0908B0B3DDF8A36D449CD0A65290A9E664B158FC5EB1A8FF1E78CB
                    F920A15F3599525D30220619CDEC561BDF8F0174C8215C7E1F878F565FE73C68
                    DE460320B8EE33994CA721F6FD5D38DCB10437CA8AC502DBBD6B3765028BA562
                    E9D7B47700ECFD79927DFBE33826986E18C7C86112527ABDA11308106AC67815
                    AB829D53699901920A9EC9879B26D423BB9314C831B16BED76DB0FBAC2E1932D
                    56185DBE101D9F987802C681C16BDC4FA880BF9F8A39D0293CD9C474301A3720
                    3A86E37F2B57285C572896DE83C1BE10F2706DC0EFED0C8682E7237E9F4F8924
                    7C2E0783D5D08F4C2B54074C689BE1F334B65AD566FA6E40C22054CB22EEE945
                    3204929CB49B4A146B2DB92A606C8A16FCAD079F3FD9E3769D0174F2D27A861E
                    F7494A26168F335C9B0C127D63BD56BF9787B7D7B9E35BF3761800E4563F0CE0
                    1487DD76454728B8CA6CB1F05638E03E2C9D4AE7E07D7690A7DD88915782A8D5
                    89F51F9847204FA7F5A006198428C7CD669E670F1167A3F7553CF60A3C37DB46
                    907AAD1642DCEE83A72990656B1C0E3B2B62E0C7C7F73E5C28E67F0EAFBD8794
                    034D107918AED3481E4B7B19288D40134C3B9CF8F9054184E1FD1213F34B40F9
                    71A954F25C40F4B11E8FAB1F5C4123AA5476AFC7F30F20B0DE72A57C7FA552DD
                    0E55503C701C489D586D56422941A7D72DC764AA2181E320A8DB00FF2C97CB31
                    8BC5CCF90442891F93BA0208F06E288BB3FD7E5FC06236F170479F8DC6E2B258
                    17D5380E49E01945966E0071540E4650DF7643E00820CB3000CB29F0B68B6935
                    8E72F3D0CE80B2699648A41672F9DC6F31F62FF6F5ADFC2606D95E2816FA7189
                    BB085A691998FA0111C4E226595322526EA11E210F823CEBA0F35072A939000D
                    7E4E204A088873BCC7EBB9C8E9709E8CC914A022CAE31393B7972BD51B144906
                    9C6BF6C76CEA094531959249DC8814CE236AB0853A21165D07A99076E8C1B91F
                    C2711E2A16175660723F64B7D9AFA4BF5BADE6C1AECED097104EAE4DA452B78A
                    B5FADD98D0A7A9C6713F2A30323A03A287762D14312B944A3388EB399A60327E
                    8CCD3042D60618C0E99D1D1D670071D424432BE5128BC4622C97CDCD23566DAA
                    D7C4C0D2A5C3EB698D221E8BEF52AB55DB694C48C5084CF5BFC3109A99C07A97
                    DD613D0306707E28183A0A528837DE9A81148C27924958FF7D525DFC1F68BF7B
                    037EBF0593F7CF803F0FC2C13088DC2EDEFF08048B167A4C4E1333D39A802C37
                    FB2249528452C11A8DB683BC880C8B0C0CC7E8C4209E82C17B6F301038953282
                    60DE550C8E0E88348FCF7E44A3D525DB12935E94F82138A6F43044054FBC9021
                    E02688BC7169DA5EADDC2F015BD701ADFFA2D56AB90B867665215F9899999ED9
                    E4F7FB8FF1F9BDEE81DE9E0F821F5C6E3419EF0122DD8EF33F4689245ACD24C7
                    B53AAD87C3F8EAD94CF6E1783C86B0643FAAA7A7E7D86030787A28185847C821
                    4B74FF393E06D97C6E0764ECA3708C3BC0A19E4798B815634B8935A55CAD6C03
                    CA541BB2FC47CD93E6CD34007848D06030BE37DC197A6F28143C8E6EA85AADB1
                    B9B97986F92F63626EAF53518824FD56AFD58A02DCCE6C311BCB95D2B8C3EE38
                    0C30D9439B4D79C2049E409302B6CCA19AB27F2D2F26A5D0004278B218287872
                    009371567F7FFF393EBFEF783DD02393CEC61626F792FEAFD96C83FF8063E9D4
                    2AB55E867111F7A0EE644424484AF2C9E69EAE6A3497A0F98294402FEEC9F81C
                    2D2E5136913CD7E572F044118C0EE8AE5B45994B51121F2D572A1F9E9A9E393E
                    164FBC1BACFE2484F3CEEE70E799B57AED4C90CA2792C9F43D40BC7B80829170
                    B873793291AC61CCFA962F5FF6B5EEEEAE73F1F96EBAC772A9C817D0C0991006
                    4A8FC2C0EF85113D80CF6EAC566A74AE656E8F7B038D050C300992FA3485B03F
                    F6E74D31040C92533099CEC70D9E1F0804DE096D2E408BB37DFBA6593A9DCA17
                    4B2578867C33BCF97E787E49DDAAE923E60E9254359B8CBB7133EF00D4F5540A
                    650EC1B2ACE250478C9BFA13D086136AC18758BD1B8621E3F39D0E87EDA3FEA0
                    FF0CAFDB730C792E626E746F3C7E2348E7DD30CA47E1B1AB31807F8BE3B93179
                    8666D149831B98C3ED64A4F9A9455D734D40A5E1FBD4F8424E0356A1F0342485
                    239D5ECB28E3476DF36CB8E61632E8700DBD729304C711D38B0811B797B9E2A9
                    1E81893ED1E174BEDBE7F58C21C61F09A277642412BB301A8BEE453803B5B02B
                    471F75E455E02F469AD40C5413BEC360DCF952A57417DEBB0708F2244E390DB9
                    CC1DAD5CC9A8719E13401835E460409D395CD3D36A95EAED35040C06DDC4C55D
                    5DE1F705FDFEF5168BC5489E4308908627038A89F26F8151FC13A07CB24DE478
                    9A1644877A37420554A17DF708FC3D7D0F181DAFF0957922A6B9444C169FCDE6
                    985DAF27CD5CE3F2CC68080D0D0DFE3B3165482D48FEC44F4AE5D26DB89EE788
                    2E50C18946A3D3E07711B06F066337B0D66A22F5258491F0E3B3DFF34C838A73
                    1B2E3B6BB4A4D026E0841208054D546A2D40E13300356D1F50A251AF4B31BA4E
                    BA1F5C1F0885F2584D911F5B985FF8792C1607E933BDDF039470385DA390A0A3
                    445CF15D35F88E71767696F218B572A9B20364F37E9CFF7EF0928D20D0591EF7
                    710D40D2F62296C66A369F48A1329DCE123FDA09B34E1C48AEDF5243C0006AE0
                    0517C2FB2F87B5AF458CB6D54098228B510A0174633700362380BDABE0215DB0
                    5C3269BED0C3789C2FF3D1E752AFD1000116F791B719F47A1F20DF056F4EB777
                    1D1F704E2F60FC72B0E80B2127C91B5588971120C58FC13F6EAD82A94376D669
                    427430401E6765254F9445C5CB8A55465E5CA0E27A94A9A8DA4068C67F854211
                    2CA56DA83CE6E2BAF617A2E05AC910DBD5C9F42E2649876BE9C6C517F0B5399A
                    28BA665E69443BAADD6EAA5FD8252A720762BAC6E70B2850192A9ADC12E09FD2
                    E4B97C9E97B86773B91D8964FAEF30664FD13590C1E5F34D390905C279105FF2
                    56AB3D0EA7630318242169A92ED69F11FEC44DD29A3F570900FECE09850257FB
                    7DFE5580742765CD68675022992258BB179EFC7D4CDA23982C2F2EF244C4C735
                    7AAD6E109FDB25B416614809B45643F8BFF8FC34C55F835E070857F7342431DD
                    4E17C3E01C8893578E8C0C5DE4703896C2E3B4548984985E8D27129FC3117F06
                    665D6BB3E6A6DC5371E287705492E1DDCD8254D94269673ED9642CFAA6E6A644
                    0E9521100A917136AB8F1A020F110DA1958EE692AE95826E9E03A4D6AA056145
                    D88BE06B0BF4F9A6D77243F081305FB264C9C0D91D1D1D7D20762EDCA38AC82D
                    95BD150B30846482511670F5AA95AC430A0D4463B16F4422F187C1896EB4596D
                    3BA92B6D6B15B5DD0157C03DAE75219690C1E1BC71A0D3632A95EAAD3304BEF4
                    5A1729A1F1917057788DC96874962169A0672913086FC93C52AD89DF04297B02
                    839124E882D5C2F04B7BED0EFB1AC4D865203977224E2B4D19A6A54DA7DC73B9
                    8149721AD61DC58D7B31E85D98F84D08312E4CFCD55D5D1D17B85CAE014C8421
                    1A8DC6E3F1E477BB7BBACEC344F4812FC844081B07A47629A78F8800B4A9512E
                    A00C364D39009A64DBEFA19EB1DDBB77F3CF0F0F0F330904B5C10B519A3D13F8
                    9638FC2ECB12335BCD0C86CC3298445A5F50B752D9B4C60003D4107A61E2B9E4
                    8593ACC23CBD7F6870E8E44030D0094731D1A45196903800A45F0E936F9B9E9D
                    23325A46587C627666DE0BBEB0AA23D471584728B426994C5DB218893C124F94
                    7F84CB7BB89D1BC098AB80822753F28B10A252A94CD16EAB3FD5A9357F02113C
                    C1EBF57EBCAFB7E770C4497BB15862F3F3BC248C52B44F61D2BE8D0BFC1D062F
                    D3CCF96B9856AD23762B96CA955D60B834392B2AD59A0BA426D98CAF4A1382A9
                    238188F8EB72966AD5DA2E8BC57C0C64D56AC8BA40CFD2AE8F024ABBF03DFDDC
                    FC62329D497FBD56A9FE1837BF8843AF319BCC831AADA65FF88324125481D0EC
                    274DBFC39B8A98CC3A65F4E05056B21A5A8ACEE7737C291983C9B66EDDC243C4
                    DA756B544DD580EFD5E5B2D160AC596D169EA154B59082CE9300A10BF87D02D4
                    49903C1ADE5D2D158B477A3DDE4B962D5D7A0CC6CA0EC3D3D0B1A3D1184F14C1
                    0036622CAE834AB10703BEAF43F231B3D158C0DFBF91CE665F48A4526781D3FC
                    ADDBE55C07C5D10187BB2893CD9DB6B81879A150C8FF04F7720B0C4D76BB5DC7
                    525611DE001BAA3D4B95506FB621A801F147FA3C9E4F0E2C59720CF4B191325A
                    A402701160AF15EA11700D60E97AAAD9170EDAF058E00B2A949C8135AF4DA7D2
                    6E5C7592D7021CD06684A00F6CBF426AC06A311F0BAFF8047887004FD24D4DCF
                    2EC29BBE072EF1239C2341C40F504B1E1DE1A5E7020BD1B508CDBD084C670001
                    6D55EEF02B50AB8116AC44A1010662550932F100B93DA9CD8C9FA8AAD74BE16C
                    3A73B456A36365B982D8ACEFC2DFBB10A6B2F4E103D3B644867385BC4E51CCFD
                    E9748648E7EA77F41C7E13BCDA00A31328659ECB171BF95CAE02A77900F7F103
                    AD4EFB18B8611152F9EB7ABD41E8E9EE26A25781814CE1E0299CE227C562F166
                    FC7E642C1EBF0AF07F222D3F2F5F367242B1BBEB2838DE47F1FE934E87A38FC6
                    0FC7A524D4231AF59F1EE935AF06FFF87FDAE0BD1A31EDE3C38383A775F774EB
                    70716CE7CE9D541A5E4718D8DE15EEEC82C57A666667A3B8A892FC9244064DB2
                    C369E78C1CAFDD18882940FE3058520893BEE725B9705A81F4831C5D613219CF
                    169AEB0BDAA9A9E945E8E81F0286AFA322545A366C7B3EC56A4CCEBC00C28181
                    F0BD34B74E34A4FD0E808016A44A2A4E5299050824A85A311E2F030E37084FFC
                    0098FC051EB834A595897F402AAEF778BC0F8220FE289BCFFD17E52EF079AA69
                    D503FD8671BD97FAFCFE4B310C842C5A4CAA3689980F765FCF178A7148DE5FD5
                    25F1275AB566275596D115111421541C4668484892AD54CBB897482BEB482193
                    52D2BFAD546AF7CF9723A3C954FA0308ABE779A015972E1D5E63329B5643BA0A
                    8434E54A390E477DBAC11A6FAC21D0B2283C6C148372EDC0C0C0F9C3C343AA1A
                    6076CB8B5BD9F4D474255FC8EF032BFF01E8D03731A0DF0573BD0AACF77D8562
                    F17683D1C817590C889D6DB64ECBA22D6F9F00699A3618F5C33AAD6699224B0F
                    CBCD49D063923B41242FEBEFEDFD80D56A0DE2E6645A4584F7EF486532176070
                    76BEB4A2D78530D38AE33084860015E26F287CDD898F09E50768225B0C9B1782
                    489292D7EB39A1B2643269B5C3663763F256230C7D1064EE1C97DBAD45CC2D01
                    198AD4EA805866A15052AC369B33D411FC479BDD7A593412FB4F84B7DDB8DE8B
                    8F3CF288D3DC2EB79E623F8518E249301811EA653BC2E40D502A3F837127DBE5
                    614D03E65225440B4D747D4402C161A280A7427BC554569AABA6AD22EA2D4089
                    6BD2C9F4BFA412C90B803A97019D464D66B300676C4049D0BE88C29FF3EC0DCD
                    416ADF1C369BF58B2343831F0A76843474EC2D9BB73676EFD953CAE5B2D388C7
                    37001CAFC3C0968884010EEFEB0A872FB6DAADC7C2107871C881A59522064869
                    5D206014E3531AA7D4AF56AB5B0903F16AB45A3886F1B2404FCF3578DF0348AC
                    8E4F4C6C86C16C1D1E5A72A1D964D202C6B5C420DABB91284E834937279A978D
                    0B113A0574BB079E6181EF17F8DF1A222B2026D36AA5CD6EE3300EAE92155454
                    F1DBE844DCA7587CE9E0D0E0BB6C64B0A572767A6A6A7B3697FFC9E092812178
                    FB27884F64F3F9ADA954A6D0D1115CE7F77B7DE170C79757AC18E3E5EF4584C8
                    8585791EFB7360FFA9448AC513F12DE04FA7E37EE2B2AC1C50312530A3C5DCE6
                    30A330405BAD56E42BE7E033F3B2DC542444FE0849F9021DC6AF0EF9090E2000
                    A9F0B1FA03F8427FB8B37304C7D021F440368A0F09C29FD75CEB6088F0A99191
                    E18F2C5F3AC217771E7BEC09B66BD7AE64AD5AFD86516FF84E436F28022D70B1
                    556632591053A5478AC572DC61B7F527E2891371F1D7B70F048FE1C99FB6E4A1
                    C59872A5FA3C0CA46A30E88F9424F36730B8E7582DB6CE62A9284E4E4E6EA3CA
                    9F6A5DBCCEEBF158C1BE8FC68074695442BF4AAD7B918E212A22D3699A09A646
                    AB6109DE4758521A185833B8851F465A006231B15AE7A969B12E315C1B0DB0DA
                    E1B00B805842AA8B7B7A7BAEA09D4BE03BD18989C917E0C93FC5B5FEDA1F0828
                    50289F24FD4FD70CF49A32D96C7381406079472860D0418910EB9F1F9F6705E8
                    7B0A4F94EDA3B049E104043187494A0311482DB5388BCCB7D2D1049391625C0E
                    6F95E3517E43849899A36B2523A08799788CCD8D36A22419117AA0384C47018D
                    2E04673A8132B08440A4D4C089F2908D8FFCA9B2F1150D01073E95AA7A1F7FEA
                    69D6DFD7CB8EDC7004EBEAEAAAC662D13424A1B9542A17E99C8B91282C3FC137
                    54422E3ED1DD1DEE83879D0D54B89E17900B7FB8064FD60D03D2814069A0D573
                    60CC431E8F7B884A3880009B32D9CC2F60FDD763B0125A9EF737C15B6A930835
                    BDC0C650A558E61DD2E9885008BC18743F090471C4A0E561205A84139FA0564D
                    AA54BF570F400283DEA05F6AB1584FC2F14FA7A3C0A875D56A6DEFECCCEC2354
                    FA8681BD1F1EC78F478B62D0ED0E9A14A0153B6CFDFAF780B07283A6357F322A
                    A04612D767037FD2D1DF60D8DC400D9840FC2D9748A4249AD8E6D368344CA9CA
                    9C5493A1D05804FCFE35B82E353915DEAB21644DB55668B9D1E07A420819CBA0
                    984E0E0682E779BD9EB046AB26DE3133353DED2247A163E31AF760A467F7EF69
                    78A30C01F1E626B0D03182F34D9BB6B05C26CB962F5F1A0E05FDFFBD63D7EEF7
                    C27BFE03C6709FCB69E7848D278E72F95BBB1AEC42848777C01082B8B90855E6
                    1C50EEAD063AACB73B1D6721EE5CACD668FC24D7A84359BE58F811BCF76766B3
                    29A6B456F2B4184020440E5E3369B15A4EC4677B008D3CB60A240BE1816A12FF
                    8D66FB7C9C248773A480605E84E280D2CA5A160A45BBCD665E67F1794F0B0583
                    EFB33B6D5E92A7E4A590BAE3B8F60F82083E0EB9CA37B440F3AB8D26D30A8D4A
                    7D1C06FFDD186C5E13403908AA9520898CFBDB0782F654BD5ABF0B131886279E
                    6F3418C660083A420A2B42563A9D356322BD2A419768D6301C584C4BCDB71A06
                    18E432CAA412511425B90214D8084412703D2B802887793DEE77C330DF054BE3
                    C529C9647C53B15C79706161F15EAFDBF5DF0E7C70766E4E82313FC26B33FECC
                    ECB0E620E9E25FC02A3F0B35602146BA676282C58110ABC6C6D89A952B4E8665
                    1EBE73E7EE6F41C6FDC4EFF54ED100C612F1874BE572CC697784E291F849F080
                    1B28778F1B53C1436881E5DDF8DEA54693D11B8F25C54C269BC4DF3D98BCBBC1
                    A4BF513F6069578D09D4E1F75A5D22163DC59B5F021A41A1B51419786938FE4E
                    CF84E2AD702B7586E356E14D71FC1E007FF1150A391B08DC299D9D1D670483FE
                    7341F4F455DCCBDCCCFC3DF42C9FAEEEF04930960474F8E381909F81FFDB5C4E
                    D71198A35302FEC06981807F09E973AA129ADC0B034825EBC542716BA55A7B50
                    94C5BB01ED8F6BD55A9C5BA4DAC0DB61A85F75B95D674119F0CA2918F55140C8
                    EF6533D99F6182EFD3EA74BCE49CB283142E0545B51492D49B4A9539A1CEE5D2
                    94DC5CD3D313BE0893FF1E8FDB3540ABA270BAD8DEBD7B9F6C6E9415EF03AFDA
                    0B031BC6587AA9BA1A460924917FF7E7868557E2087338F1DDB8A1F37B7ABA49
                    0EA576ED1EDF91CDE4D62F5B3A621818E8B53BDFE1FCA7DDBBF7BC736E7EE13B
                    80AF5BA0670B880F0F7675862FC3A49C010FBF01107D3CE2D9B91E8FEB4CC072
                    209BCDB37D53D3F72512C91F0703815598848FC1180E8351B8342A55BA5D7944
                    08272A1C3E1B88D1534D09AA0FC29302909C73141B29EB50693D14AC90CFF3E5
                    0AC0F41E9BC5BA0E93701EA0FA389FD77B2EE5E669B026C6276E05A4DE5B2A96
                    7EE5F5790FC7A98EC0F9FC2E97632D42C420F8C07BE0FDC7E19A5C8430F956EA
                    37168DF31DD4B894DF2154FD238C670F0F49AA46BBFF1203699B00719D2045D2
                    90B954E11B683A3B3ADE8B639E1A8DC5EEC864733FAF94CB77D32A294D30CEB7
                    0AB0AEC98368AA9BBD1ADD6363A3DFC33934E05C083F89E761840FC2781F8633
                    3E6AB1596B4AADCE1352409ECBEC763B775290ED05852B0AF98D37044AF0412A
                    5E0F783B9FF6E42196A9D3A9F4AFE3C9D47FBEB069F35713C9C4C0C8F0101B1B
                    5B7E34066E74CF9EF12301E9DF006BFE0E06F702B0F36331C13F04149F02D8EB
                    80EE86574DDD8B9BBE1561E71E8BC5B488898B00792E07A4AE4BCAB20B042F4D
                    048AD70648B5668E9F8812241526A10E8215566BD52190A6398AB3AD82D2663B
                    53FC3798FB4AB0E94E307CB5CBE53A966236063F3D3535F3EB74267D1764C9A3
                    44E0C2E14E2A5829C15045786CC8EFF77FD7EFF38D01A6F544CC284B48F50E40
                    BB7D187405863880EF1327D909A9C7574469DDA22E8BCC6EB1723E00596FD0E9
                    342BA576136F383E24E40E18B8AEABAB73696F6FF7F9BE52F9C4582C7E77341A
                    FD492C16DB0AA5720A3C598731E27C4769281A8C5F66DFBE997BE0840FD6EAE2
                    D340A4DD442C5B092318679E1344A0D591704E0DD0B651AD569FA094FA1BB170
                    78304480332BCF178BF91DB8A865068FDBE10FF8C873BF95CDE5E6F64E49D740
                    4E41EFF73040ACF308E7FA8FEC99D8BB7ED7CEDD7700CAEA904D2E58FF15F044
                    DA8B7837A0F1E718A44700E78BB42CAB6B3EFE6F37F4730281D28FC91F42ACDD
                    4B919E649EC3E9E04BBB7CB3884A95C5CDCFC10BC2F0B880A0653CA6B62A9368
                    52377474769CEBF37A8E476C5F4E3381494EEFDAB9EB87E96CF66EBDCEF03C50
                    AD62041FA16343299C8649FF1B97D369A73D86E02AEBAA3CF51B810164A8F4EB
                    799CF7D7989487A667662F0237FA7BCA834031598090948554245C9BC361E3DC
                    818C11644D038FEEE4DBDE9A9640D9F49B6150BFCD178B17B91C8E335D6EE720
                    9CEA128CC151BD3D3DD318A2B520B32A3AF79617B7346666E79E181D5DFE39A0
                    D0165C539E8867AB0C8EBF6AC562AB3847D30FE7EBA5B3907140D1DCF7467564
                    3E581E816C210BEFBDA1902F7C957A0160F08EB758ADC743E33C881BBE36914C
                    DE572A143F0A085BB364493F5B36327CB8CB615FB96F7A564B440DA46A0EAF8F
                    41213C810988D0C20FAFB9A7F888E383E0C4E1D9FBE065CB353AF57AA5DEB81F
                    03C3CB8DA9F084343FCF3B485206A33A89EFF7E39A5C1CAE14AEB3CFE8EA0A5F
                    002F7F0742510F7D07CA2505EF76E3FAA6A0023E3DB46489C29F1DC998D36C30
                    9CEFF278CE00015B85D011246D4F8A804AE653A9B40CA3BD0BDFFB75A3213C03
                    6236918651003DCE1B1C1C6488F37C736D3B634AE7E7A5F44AEB211F2A951D48
                    15842C6E66B210D1E0C9B398D02D9148740B486E03C6F47788FD4684A51E48BE
                    1E929C38A7323737DF00C790228B8B77E2D08F916A6915DA72744CA5AAED8DAD
                    FCDC0683EE38C85F3B3D4F130E9601923CCEDE2C43682564A1F0E47B72F9FCE7
                    BCD5AA0D13E683A79E0E987C107FCE2136FD5CCD54CB71412BB66EDDAE413C64
                    1D1D21C3D8722B9B9D9B476C8DD63079F77ABDA1220D1C48167796FD1B3D7063
                    98AC6DB8A9330C3AC361D54A56ABC8022F38E0EA4025F07577B3C59292447912
                    067432067B65A55A390F04F07278D71AFCEEA318BB6FEFBEC720D57EE470DAA5
                    A03FF033788D1EE8A0803FF4DBACD60F82A89E44AB95081B56D2F569C47E5A21
                    A5D2AF52B9F23F80BFEB7189DBC12163A4340441C77A7B7BAD4E97DDEEB23B78
                    9F65DAC9443BA2DB491B4C30EBECECE44BD530E00EC844BB441AD56890718D22
                    90D30F24B812E4FA5C10BB21486C6316E78BC71353403F4F30E027E92778BD6E
                    A8045BC36AB164486D50CE82784D9D2FAB2BDCEB495EB6D74E8064C740CE9AA9
                    6EA152AE6C6C6DF67DB30CA1AD1E94D942A9747BA158BA043028582DE6F5B8C0
                    3EDCAF3F180C7CDC66B31E0989A5A1183A31B9B70032660D873B78EEC1E1B076
                    EFDA357E2394C427007313CD5D46024139BF518D8674BAB811839BC3E0AD6319
                    6681AF977EBF2388571F73A43768B53354BAEDF3FA2E073A5DA2D76B9D3499F1
                    78F2F64C36F3638109CF21444521697DF07211DED7856BB863687870B9D9680A
                    D7C5BAA6582851C127219502D81F0779EDC06459A391E80F3B3A3A1E251826C3
                    53A9255EAF28C965B1C31CAC0842A35D8F50E5E410D7473B9572F9028E9564FE
                    4080425588F7E2CFE52A5645910241BFE3D8771EFD6918B1094AC3BCB0482BB3
                    E9DF20CCFE4C50A9377A5CCECF190CC6CB703C357104F001BDD56CF96C329DE9
                    44DCFF2FCA89ECAF70361938321037A0C6A008672B804002190838D17DD4EFEB
                    4D37045C48B156ADFD3297CF5DE2713BA919C4CA8E50E0B7805B2B2E2E303333
                    070F16B2F050073CECFA9999992A38C2473B3B82EA4030A85DBF6ECD999353D3
                    63D353D3FF819BF82E2E5A69A281C277ACC32336C1403266B3B157500B43B8EF
                    98DD6AE7A95AA1D9C56400CAFC83081317519C04A1B4C513F112FEFEC37CBEF8
                    2378F2AEBEBE9E1C11DA6ABD6683B1508D24B1781BB8C819D40A8F183FAE9F65
                    D35996439C9DDA37450FD1BA7C746CF43F41B80E47D8314DCF4CB350A8A35908
                    2B345704CAC5D2105E6382C7C3D10BBCA10664E13C831E379C87AC244E400924
                    A0D252B7CBC986069798A1EA682BAC2A91887B2727F74E43297D03F3760B207C
                    164E943352EB7DBBCD0DE35483E86D037A0D154B659DC7E3EEE9E80C7D6A6131
                    7AD1DCDCDC0F6078DFC7C5E4C80894661692D2E78781DA5055166D7401624B0F
                    BD91CFF57E4543A0A5649C6C5B21977F01DEB616306E72389D4BF68C4FC88978
                    F21E30F52F0029D6F4F6747FDD66B3AD4598381DDCE1814AB5FA956C2EBF0AD6
                    2E0CF6F7F5813B7C6D7C62F2D4783CF5091C7607EF5B487B034471169F9D05EC
                    F7AAD59A35F8FDF1D65AC25AC4E0AB56AE1C3B13AAC24E648C56E82445494ECF
                    CC5CD115EEBCAFBBABABCA37A8148BFDB4F2D7DBDB733E9442378C430BB5C04B
                    BEC9A0609870DEC2F7226083B89ECF2C1D19D245A2F124E039AFE60A4576BA5D
                    6EA1150F7518D8138158978742C1E321D16C34D964B880F78BE1B59B71ACBBDA
                    E386F1391A11ECD2552BC7CEA07A489C4BBD7BCF6EB67BCF9E241CE88B1E8FF7
                    7F604465B0FF7ABB301706ED32E8F543B4990723BC157298FE3B994CA4668341
                    DFEAAE70C79280DFF7FFCDCFCF5F313BB7F0EFD46E877821C20D35E23A0AE7A1
                    DD56B41988DAFDBEA14F867B99211C989CA09DBE88A3BF424C5A4B8DA62C66B3
                    0C45700D58F40F1073650C5442ACD73F69B39A0F075C63221AF7E5F3F93DB881
                    27ABB51AD5E641EE040C6BD7AC7AD7BEA999B5FBF64D7F954AD7A00A0AE46638
                    D61678C3D198CCA3400C77C34B3FB862C5D869B4E32393CDCA5353D3205BE223
                    C3C34BDE8D58D9DB1DEE123D1E970CC2BD81F60F2C5FB6EC4C5C9705D7A3A6F2
                    38CAFF53BE021A7C0646F63D84D75FC0A3E63009470FF4F7297A834E9FC9E514
                    289022B93E148C1593D3817B3E0F1EFFBE80CF37EAF37B3588FD02710332022A
                    5D033F1B848CBEB53E3DF35D4CC4DE756B579FDFDDD5B916A1420BBEA0426894
                    E6E7172671EF43E01E511CF776AFCF9F155A0B6454754D5E0DC358A6D5689D08
                    9951DC6B956A6401FB9B1006CF98DC3775852395BED6EFF30EF4F7F70DC170AF
                    839AB8667E21F2FF30AE77D91CF6553ABD5E9DCEF022D507313DB537D5105E5A
                    9608547810F13F09CFF460B0D46BD6AC0149ACDA49A6C143A631B02F62A07A70
                    63A7085661333C230E0FBF8BB6B54F4DCFCA9814199C423738D0EF03847E7D7C
                    72EFE9F144F2930D45DE8E81DB01432861A0CEECEBED7E0F494BA04A3E168F3F
                    5B2A96AF8321DE0CD4318AA2DC4F891F98E8A5D0D11F06513BD90C9CA5CCDFBE
                    7D53C4FE65727D4994D44A43364D4F4FDF02C3FA4A674788571BC3C32B380E40
                    ABE6A43C00A46B89D622C25D9DD78602C12FF8FDDE00957CB5E4204BC713FCBF
                    ADAD5275BE78A45269D7AC5A712D29062ACAD9B16B4F06E4EF0928950DE0417A
                    7CE699BEDE1EA8EA6E259DC9555B1C8769745A5605CB97CB15E674D8D782083B
                    41B69FC4F5185A5553F31AADA65E2BD7AE4BA7B33F2FE40A57D99DF60FF8FCFE
                    EE9191E13190F05B80A85BF1DDCE764E01F7F6009C4D7ACB0CA1C5F027313877
                    80C8FC0DEDF8185CD2FFE15C2EBB7E6A6AF68BE00D8F00826FF37A3C27E0BFDF
                    9D4A55FE1D811484B6726F20E0FB5BEA6134353DF320A415A5987D94E859B36A
                    D53198BCBBF1FE6D50166B7048A83B83108DC6002AD94745B1FEDF18D8DF8D8E
                    2E23B96415986AA856AF3B695221172FA0851CAA209E999E21F81731308972B5
                    7A0F24E44F4160CFC2C47F5450ABBD557002181CEF9B08BE58A3FE490DDE0C5D
                    FB41835EB71AC6C93C5EF7082F0815255E220EF2C952E90C4FDEF87D7EDA71C4
                    960C0E308FDBCD9790A9D298924E02A0A6982F3E8321FA2797CB7137BEA3F1F9
                    3CD9BEFE3EEDE2E262550382D22C67D7B2E6B6F4E69A8BDE6018058AEA80023B
                    0C7A573FA903FCBEAFBD795768083984C17F03A9FD2150F8A35009177584821D
                    DD5DE131BEE601E342384C002136B3066BBCA986F0D2656DDC40B68A89C5645E
                    22CA49EDC2FC3C03CCAE5BB366E55DD3D333D7CD2F2CDC861896B65A2D6BD3E9
                    CC003EBFB35EAF6DAED6AA498459233CFC5E591223B8817F2A14F2F0CE101B1A
                    1A4008757F70617E1132A8CAEB051723D16F60C0FF6568C9182B576B7E90226A
                    1C712162F8392EA7C346851A541A4FE560994C46049A68F74E4D4D582DD6F361
                    00DB69E06130AB28DC83D07968E0BC60F830141F066EA9DC500C548C3238B8E4
                    D33EAF9B276A284B4939023A26587FEB5A342C1C0E43FDF451D690AA7FD8DEC9
                    BD7C7B1EC13C2686DAF4336DA7E65418C59A6AB5E682148EE25A93B42E22D6A5
                    5243106A4A2B65DEDE9185BF6921FD862B951A4FA8813B1C45EB0538E6B8D0DC
                    1ACF65A3B659529F126BE267218B7F00BE732DC6F1FD90C076CA7BE07A9E20B5
                    F5466F4A7A8584D2819C414D4BA32F60E09F1B1E5AB281D2679B5FDC56B03B6C
                    8515A3CBAFF6783D67C33BF408C01ADCF8BBEC56EBCE6C3E9F8074FA19C8E5B5
                    8889E76875BA196A583D353DCD26262659B379B58F19FAFB183C88A501C780D8
                    8E638F3E72C9F4ECEC5198D80BC02F4EA4F57B226CC4FEA96106A85FAD58AEDC
                    B977EFBE5C7777F74592A448A9746A269BA50450171513CD929F60920330AE11
                    BD4EBFD26C31BF0FB04A2B98BA9A52A1E567FE802E5A4924194A3585B43C8018
                    CC960C0C90EAE07518F1589C3DF8D0665C5F0493A4E159435A4E8E4463D4718D
                    DF03A11CEDB852F1FA774D1F877A59CE2BADFE7594FD54E31E68010BF7B154A7
                    D306308E04E9301CBD2F5FC82342CA53BCA1268C80AE854219F11292B31E8F67
                    1AA8F7595CFFE19897C37936B126DE4F7B7DFFDC4294D74116D507DBBF300DA6
                    7A2F88CA062AB4EC0C77E6366FDAF2FF16E72347AC5AB5E27CC46D232D82C07A
                    DF9348A5BF2F4AF53C62FC33B0E26B5D2EE73924145ADDCDAAE3E3934FE186AA
                    904DC7408E9A203FF902CFD0D092F3EB92B401703C469E4C4D25E66617F0B714
                    1D2CA2D368FD50237978E1C7F07D1D60F354B3C9D41F8BC7CC9228821FD42935
                    9DA7256AC0FEA0DD66FB0DFD4B296BAAF2A175FE2AE23415DCC250F97B94F124
                    84A2841849405A169E999D65FBA6A618983C478312063F180872E3A0184DE841
                    3C825AE00603640C5ED60B1A9BCF172EA773D0723895BD13DF692684D41C5974
                    7AC32854820F28B407EF51AD8209504FF26191359ADDDF28F4123AD27F63AC4D
                    308825769BFD6C93C93C4CC601A3ABE31A9EE6BDA294C69B6B08CA419A2EB55A
                    C13C0A3E162B974B7E48C24EAFCFC31617A2573CF6D8130F77F7745D0A1DBD61
                    686870BD6531F29F38CA26DCD059B4908418AD05A4D5113E2AE05405C4C24F60
                    003626E2F1AB2A95F2877D5EEFF296F1B9C0AC695713DBCB0960AE0432F83C0C
                    E237189FD99EEEAE2F10EB8EC7936178CA1EBBC35E058F0C01CADDC562D9D268
                    284738EC0E5A77606EB7CB0A34B1129A6452191E525C6E275701B46388B6B787
                    970C304AE9EAE0E56418548F19A58AEC5285C7622A4E017D67F3402CA9F520D0
                    D69E48BE9137D66CE0C560D02C10F453A24C253537D00E812CAFC28737D3A4D2
                    778BA52288A27314636B06817A1AAAC84E6B0D30C6798C6BDC6430F2A69D90B9
                    945F1801495D01D279B4D3E53ADE66B1F4ABC133C831F0DA82719A7F85EE6C6F
                    6E6868BF0708DB022FB81FB07A3165D7A8E1653693FB8D5AA3FE2988E36FA8CD
                    2CD0626528E87F3F3CFDFD345054C144BB8AC1B29F0301D3D9EDD67580D3200D
                    3406ED7B78FF1930F86BE1C5174085185A4DA8CA1393FBEEC400DEA528F27D38
                    7B5CABD518F0F90FC053C6702DA43E9E74391C694C7E1FC2C117E099411C7783
                    C36EE5491F9A2870174EEE6038DCB88933C080180697EA176865934563517084
                    08CF3B10DCB70B4C0929107AB8C7677259BEDED068C9EA763D04190691563A17
                    959685403E69238FDBE5DAE0F579FF676161E1BBC542F97A9CBB04831030D14B
                    696737EEF5392894208EA4C5F776A5532949713846F0DD3570A62381741BAC36
                    EB72034218E50CA0D0A8D9B7969E2E03847A901AA730E16D3284D6FB45C4AD07
                    11FBCFC644180177C38B06FD49989C05C8C76310D37D34E0DBB6EF805C0C0A3E
                    8F4765321AA8A53DDF5408545FC46780F0EA15B083BBA82814DFDD52ADD7AFAA
                    556B6178EF71B4F55C6FD0535DFF35308424159452F20C0302E557DF476B1088
                    B7C7C2308C805F0F794667B8E31C686FBE89858A47D3A9342F27CB64F3BC7F51
                    004AA513E48FB6AF13B453AD0175698DC148DBB982DFB7DC6DE6F4C9EB698999
                    6A0FD82BB4C66BBF4721311613C931B821D05ECACE5070A9D56CFE36C2CC3B16
                    23B17F8561CDE36F010A2BF8FC4338E847E0043A8D4AE31B593AF2EF5EB7FB9D
                    949CA365E76A85AAA1F33BE6E716EF808A91C3E18EBF017F0A51791C78C413B2
                    A45485B702115E09765A39EF078AE5D2562888C3A8851BACF78B802D2AE5EFA4
                    65DCC1C17E4AB9E6376FDEF2E5503078426F6FEFC9DDDD5D048F23B8390B3590
                    82712CA55A3FC8BDBCC83B8F08B49BF97E30E87752B1062056EB743A7CD49EAE
                    D99482AF5636A0BDEB94D45A31B6FC62C4FDF7EB753A0BAD6188F0C848244E8F
                    E1E1B50454764FCBC42B578C718F361AF4AC0068A6D8BF67D71EBED864329AB9
                    04149A9B575F36D1CDBD6EAC553FD878C53E44ED50D15205FC5F22A164547EBF
                    5F58BA74E42278FF482E57D85A2894BA53996C01F2F504ABC5F64EB7DBADF2B8
                    3D472A4C39923C1FE8B217CAE5BE4AB9F6845A256C43B8DA063D793E2D9993B1
                    168AC50570A0717E610DF6E61BC26BE497E641941E2C954B87994D5E8AC57D53
                    53D3F262347AB758976F09F8BDD700A24762B1D8DCC2E2E20740782EE9EC0CFD
                    5D07581BA0D346720D56BD04BADD0E14C81339526B7891E9B34004AA48229430
                    427B0F628077C26B2C2098677676044FC5B98E21466D34EA9D7C19B648BB834D
                    6C7A66968720F2FE0E10BFEE9E6EC4631BDFCF4025763B7747F9A65C9AF0665B
                    BB6613ABD72AEF22876834DA9C49788DCF12BAE938BAD0F6F41C78481ED74799
                    555CF32AF0A25574DEFE9EDE9A6144FFAF389E8316EB301EB172B1743F42C403
                    38CCE64AB53601A553219502E45081849F88FB37B708EA23B420F54694A5BDCE
                    0AA5573637BE20532AFFAE52AA5C223BA530A41975F1BAC16AB67E215A8CCD81
                    CC8400915F369B2D1760C26FC6C5FF1BE4571613F05DA7C3A1A7542B0C623497
                    CB5F048FFF372AF624AFC649B7E3F329204280B274B08C53A1E97A7B7ABACE45
                    CC1D02E174911B90A4C2E17815313DAA330CB64F70BC3234C6BBABD3922DA598
                    77ED1EE7655DA4BBEBFC390E32EBEBEB61206020A2338C0A425E2B91D6AC8096
                    DB3DBD5FF3F3D43C834AC64802526D428397AF1B11325C8C1EDA41390AC4783D
                    38891E4A87EDDEBD4702CA5ED9DFDFF738B84E86A432E54A8847901CC6390780
                    80A3940985CCE4610173937FB33ADCBC9C23BC4AFCA1CDA1B4071F52F20530F5
                    3049BFA5C3C36B2391C532ED1580D4BA4756942F23B6AF86FEED349B0DF346A3
                    F17159922731310351BE8EDF6158BD6AC5A76767E78FDA3735F3AF1693F949C4
                    C6241061278C232008163DE5FDC141047883496CF54E8274643910264A2B533F
                    83EEAE30D540F0656F5A0D2472188B25F867097914596ACEA140953E02CF1292
                    445C988FF0D67BAF3EB16C7FFDA4C245DDABEF0C6FEE7CAE53AF28467B31290B
                    EAF3FB9B2B95308CC86284A391C36EE3A494D61DDC6E4F0AA846F58C4CA316F8
                    776947185D6B9D3F29C6B80CE3B89C924E503A059C634BE38D5C6EFCB34203EF
                    2422D41173EF435C3B01CAC07AD8FAB56380DE7BB6EFDCFD398BD94C7B049E71
                    38EC782F7A123CE4C7D0D25380FB4741F006C080B9942A95047B57B8F3348FD7
                    BD7EEFDEA91B70D08849AB1DE0F979FCD04615DE1C3B9DE519BF426BD9971068
                    D9B2110C74272F79A7BD1533080D24ADA83917A58A1B8DE624AA040D6B6A6D99
                    B5771BF1D8FF1ADEDDDC68C61FEAD07CFEC1C19ECCD6AA942203A062165AA2B6
                    98AD084D56CABAF23C4324BAC8773D8117F04660840CF470303206AAFA82A73B
                    206BFFA954AE7C4DD0A9AB7461CD3D105483402AC7BACA0C98E0398B52E9056A
                    38AE6995E9BD258640DDCB5FF50B1A6AF054BE1344EACA6432B18AEAFD464797
                    AF3FFAE8A36ED9FAE2D6EF6632B9A7C0740FC74D9C8C49FC312DF000ED9F7038
                    ED57501F435AC88945E33B12C994100CFA47962D1DB9165C40C6A0E8C92B5AE5
                    516C02134C0343B1DAE375F1C927664E4BB23333333C354C953A94D9A312384A
                    C27096DF7C7855733285768776D63286FDFB0E5FD512DA8DECDBFB320ECCB134
                    275FC32796108932A434C1E4CDD3D333CDA657C512EF82424AA6ED412465A907
                    44B833CCEC0E3206BB1EE8F8D9C9BDFB8E4BA7D39FC1043FDDAE8CC678F96D56
                    CB2ADE29054E50ADD59E42088DAAD52AF66681C2CBF73548AFBDA8A548D24226
                    937918B0BC92B67E3FF7DC0B3428F6356B56FD233841050A81F605AEC044852C
                    56CB22754EC5A402DAD84A9960CF6CEE77BA1C0DCABE61E0B4F012AD2C3BB9F7
                    B7079A0A5DA9E2899E92464490B28F9BB76CE1DABDB5158C561539438FC5E2FB
                    FB151077E08FE7A12E061CC19AA8D068752F69153FBF2A26F0CFA99B5DDC14B9
                    D13248156F8C418F0C06896B3E83010846656F88F7FCFAA6F64D9127B7D2C4CA
                    1F10522AC8A59C06DD0FFD7773FF8359BF626CF498B9B9F97BA66766FE03D7FF
                    F55AB55E561B5503408DD5744F18CB0690EE39461BFC1485FDEF080D6D968C41
                    CA66B2B7C252CFC52486C3E1107BE1852D4FCFCDCD99D6AE59430FC324D6EBCF
                    E8F5A7184DA69B71D32E4AABDA6D765A81A34130B47B0B10F491E4A23D88CDAE
                    AC1234BC91D19E0A7A5C0E3D96279148F16C9FAAD932A629D96055D44FD1C49F
                    BA9661225FF66D1A82D08AEB8D6653ACFDDBEF9A7EAE7AD578DF68755A6D7B3F
                    35C7E8EFEFE1358ACD15C03AA5BDF9A453DFC35213BAB981D0621525980EECB7
                    F07BF36A2E42195A85BC195C339D8F32A120B20E18CF17C727264F8CC7E31F06
                    611E800385284781D0BA039FDF273081BD993FAF5A98F26A3FF55AF549C4EEED
                    26A3B1D36AB10A6EB753333D3D7B5A3E57F8E8E0D0C09583834B1C08039F5E58
                    8C9CDDD1113C0EAAC140719F0697208FE23BE967F27ECAE489BC8DBECC11898C
                    85969989F411AE538AD5D0EAB3449E2DB784B4DC6A80491D47C5BAD06E96D732
                    86D6FF5AE497D71EF27D09C24126FFF72180AE8D88255D13C93F521A64B0648C
                    FBA84530623E4DBE5813F747184280A125037C85726E7EE115E19B4297C160E4
                    C491420811E0325F70F3F2665B6B56AF3812A1E209204692AE83540F8CE1797C
                    75861CE02D350441787D27A4C9492553BFF279BD47E1C62D5EAF771D607BFDC0
                    40DFF76667E656E6F385E3FA7A7B07C646970D109CD33A3EE5FCC983694F212D
                    F7EAF456BEC842A48E4206793D0D1250847B3135CB244EC29B6C284DFECE9BCA
                    369A3ABFC9C055CDEE6CACDA9C78DE1C0BFFCACDF6B3FB7942EB5FEAA3D866FA
                    ED2C225D036531E9117D5D0847140240E238DBA71C05317B5A406AAF397032DA
                    9A18E21EB4BD8FEA1E78BB5E22A8071942FE74171833DD37F1070A0F842EF414
                    1A2AAEA5E353BD46674787158ACB4ACA0632BB81CF3C8F7314DEE845A6D73484
                    D7DFB5936F24BD1D6CFDE36AB56984060F9E749D5EA7D3AE5A39E6A266531313
                    133C96528A972699069F433C268A3C83F391F6134E54CD984E93DB4EF1D2CA5F
                    A3F998B456B64FD9EF85C0012EB3C8707973AC56E5AFD0687D96DE9230698C3F
                    7EBD992164CDA7A650C690D2DF549544CBC9402C5A23E0E18648DDF3939BF8F2
                    F802BC9B4AD6C9E3DB467330A9C937BA2A4D2EF14A4494CE4F61818C96F66DCA
                    ADF0416843C646E84724938C9F36E952A37218681CC3B98DF8127BAB4303EF67
                    F83A7F309129C4F707DC83AE21C45595DD6EF3139DA18C1E5D3891AADD7BC679
                    E7754A358F8D8E326A29434993763CA666577CE254FBD3D8BCCF2D2FE8A0B0C0
                    8DA0F5F0AA16AC375ABBA025FE502D81AF20B6DC9E36CB356D823FA98537C2E2
                    59462276145668C269E998125094572066491C65F38B2FB26834DEAC50E69DD3
                    CD2CDC158657165E33646AB921343BB833A1F10A8600496832B1666F097A5E83
                    C8151A25C4084D08912824D9682B9D41C7AF1FD7FC3C94C51E8D5ACDDEEC9F97
                    2F43FF11F284BC381289FCA8BBABF312104107DD6CAC55CE45374599C4803FC0
                    AD9EF43E6D1FA3B84B464029E252B58E81D050873326F3829266950E393EC175
                    1BD6DB86D0AEF76FE533B877532715321AA19516A696C9B419956A0B087E755A
                    3D4F3D93D4A39044B50754154D317AFB8E9D9CAB509EA2DE6CF6C93D977A32AF
                    181DA3ED736CEBF61DFB3BBE1E9C60F24E2DDCC3693BDC4BD701F6E71CA8B107
                    2E92D4055D375D13C96142520A2D740FCD25EB12C63047D9C92CEEE73690E8C4
                    9B98477A654390FE889DB534F8F56AF5C5679E7BE1DBAB57AEFCA4CFE7A566D7
                    FC419A44E0E8C6A83A87069EBA8753957105E8E00139F2032DF488BD7C1F0326
                    50DBE203F49DE653D774C4B39BDA9FF7C91538EC371A4DF41014A15565ACB48C
                    86F1C9A7A210EA2E06EEC2A866C2E7F53187ADB9858EAA9C161716794F07EA6E
                    C28DAEA1EC5F3C6AEFC2A26BA1FD87A572E935757B3B349021B4FB3BB427BFBD
                    4E418865F3B879D30D3282FEFE3EAE40A8FEB25999441CA158011A95613AB429
                    9B5061130CF44E4E90197BEB0DE18F8D44045B8978F28B4F3EF5747CC58AB17F
                    06CBF6527C6F2650EADCD2C9CA69C09D90959478A14520A8089E9421488C449B
                    05214440A9533A19037D4FD5EA8A4586A28297CB2A1EF1F7C76BFADC81CF5CA4
                    8C9D8F261F9E4FFB27E97BE4ED137BF7710388C6637CAB3B3D53910CB4D95545
                    73D0C9551A4D55F27ABC91D64CB8535008AA37EF998C93209F2AA45C0E3B7704
                    DE5054D5E429AD7E8B55289017C19B9EAB566A9BAAB5FA197D7DDDEFF5FB7C54
                    24330D299A922489B1B7C310FE941F40A30C2DFD9D8D1B37ED2B140AFF363232
                    BC9C72ED99345F616B2680789F211D7F30864ADDDCC1434F5F252246128DEA08
                    16228BAC55D9CBB385842EAD05038E0ABC394AAB633A470E7A2E3406DEEBF1F2
                    A2131BF7FC067F162449BD58B3C50D20BECC6526A10411310A19EDE292575C6B
                    E0CBCA0DF66A76D0EC7ED6CC3750328CEE8B0A54480A9201D07E46834EC33D9A
                    082215B9E4F3C54AB9547EAE56AF3D8BD33C45FB39703F7B80063AB3C9F83EEA
                    F64E5C030831DEA07DE98AF0976308ACB5AE0F4BBF676262328A9BF8CAE8E8E8
                    89948327BE401343E8C09F71DC7C22266F739F48343BB7D3C09174222FDE333E
                    C1933266B3A5D9088BDAE468558CC2AFCCC3067F2027277D6E979BD71EF0C6D6
                    80D7A9A9695E904AFB0EAAD4109B77736D721932B825D0FAF4E0703AE7ABAEB2
                    B62474A3D5E2EEA53987D67E487E5EE2132403E9DF5563CB9985C77B0D27A854
                    C69649F386160538CAB350588F831C3F2349F20438C314EDE9A4F682851CA923
                    6500C4F108228A70868C24899324B5DFAA9F37B44D3F41310667D3ECDCFC072B
                    95EA67962D1BB992605AABA52DDE19BE8C4C0FDCA29567F2647A6C6DB556E1BA
                    7A66B6CA3D960A3D9B12B399DCA1A4136B71047A7632C13F912BC2884AADCA7B
                    1BCDCCCCF1384B852BED27AAE828B7401ECB1AFBD3D6F4C00D4D6B77F1AB1B42
                    53A9B42527E31B51E4F653617946D3E97270E372F0E493AED593B9C1435CB1DC
                    E41F140292A9F46F2B95CA978086B3B8E445B351C719307584A3F31069A6EDF5
                    F4AC07A0A099EA1A10CE162AE5EAA456ABFDCB3484FD357D2AD54C2291F8F8A6
                    4DD5BD83834BBED415EEA46719F0EC5C1913AB6D9142FE185E5EB9ABE7248FB8
                    435D94F637E7A641A14126B947328B7841A95C450889F0B67652ABD4BDD1266D
                    F044521E722BB143924EDD68D55CB614010569613FF16C1C5405108210E1A46B
                    925AADF3C8F8885B90E4245E4364927776A17434F80BE50032995C2C93CE3C2B
                    2BD26830E8EF0D7774927A5AB56BCFF851922C7F950CA3DCDCF7C8C922F10442
                    39F02C9D56AD196DDF63B95C8A2074106AFCE51AC201059E0578E97FEED8B96B
                    B258287C7360A0BF933277945D242FA01F4A36F196790DFE70EDFD0450969BDE
                    6432E9399AD044D0C25253E6894D02D9CA291037A06C602299E6FABCBDA399B5
                    7209EDCFB6615ED57A3E03DF47222B7F600CEDE7484B75913F1F9A3E4FE5779C
                    7F2076F392F31689A46D6C941FA10DB63BB76F7F68D3A6177EDBD9D9F5ECC9A7
                    9E968E46227DB158E29BF8FC11207EFED5A6155F18DF33B1361A8BFF3DEE75B1
                    BD0D9FD08E261FE735EAB4DAD5A42AE81A4AA5CA2410A8DC68B0B7CF10FE18CD
                    FA5A6BE354BB00AFBC6D7E2132572A55FF6B7070601DC13F2560A8BF000D2E11
                    489AF476F1280F07947923E3805712EC73D5D05A70A24A20FE40EE56ACA63508
                    BBDDC153C264084D586FB5E96F2599DA0BCFFCD68496F2682E48349FD0D2DA12
                    4F993F8FC7C5279E94406F6F2F977E9489A41C019D8364E7FCFCFCE29EDDBB1E
                    7FF1C5CDF78DEFDEF364369B5EA0DE4CB8F606DE67577DF823A94C3A75DED4D4
                    CC97AA95EAFB3B4221E3D8D8B2730CE38625E0271FC6B53F7520DFC08F49A737
                    ACD41BF9F3186A5012E36F56DDC1EB36840D471CFE3A0CA029D7B6EDD8C921F1
                    D5622E6DAF87073C1F4BC4DF0B3EF03510B6F33D1EB740B196FA2C359FB24E8F
                    D0D13527AD95B8A122139293945B20635109CD76FEED96FE9447689F97209C8C
                    6B7F2D81AAF56C74E1C0B2B3E6F5A9F63F8FB929418D084BC43DDC6E278FF944
                    D0E8126AADCEA669BE77A1449D59D8F6EDDB76EDDEBDEB86C9F1F1E72B95F23C
                    262D8FE352DD5BB30A1706FCDCB34FF3FE93575EF5A1F94C367DF5FCC2C20E18
                    F6E73A3B3A1C4B970EAD0402DC393EB1F753ADA7BDD75B631434990CFD06DC27
                    C63381E3EC6AE64684B7CF1068825E0F12D04D8F0C0EB1C5286F3FC799785BDF
                    EF7FF8B572601F62D57C3E9FBF7CE7CE5DFB7A7B7BAE0D8542665211546246AC
                    9BD61E780651A56E3DCFB1199B9B79F6067F9FD28F7C52596BA9B9D5DA96881C
                    CF2EB6B3902D63E58F6CE24F6F51B1F61355C820E88931B4BC4C790CCA675017
                    97566F27462D6E68595994EAAC5012A1E76B2C198FB05FFFF2FB6C7E6E7A13BE
                    BF99F1670031474B52282D43909A80DAE0E49716CF5482AA0663F9463C9ED851
                    AFD5BF19EA080E767777B91012BFB76BF7F80842C3E741240BE01FFF3F7BD702
                    1D6575ADF7BC329967263393F73B044802C85BA46010792AA86871595DDE5BAA
                    BD3EAAD7ABB662ADBDD7DABAAAADF6216D5DABF5858FAE4AAF975BF10D1612DE
                    C82B21849017096112F29CCC7B3233FFCC7FF73E737EF80D41832601AFF9D73A
                    6B42C8CCFCE73FFBECFDEDBDBFBD4F892E51A7A4BC863FE0EFEAEF0F1E1F4DA0
                    F89531C2D9DD16631432B2ABB4B8643BE98174F2F27242D41A86BCC5104EFC71
                    74311BD1463E999F979B438524C43622E6516242226BA54F8124120E0265D410
                    8B3C0CA9A7B128F248A38C7516079E097110080A4E4438EBFAA9595532F54AF6
                    30AF83082DB4F874D1F7502534A57C7D88F8FBC3149E5681C16802A72F0A1151
                    8F20F6107475765097744AA224E1A004848F315EE22326F733AB8F54C1AB2FFF
                    05EEBEF73E8AA588089E3F4293B2321409AF43615B9C9991A69E3EEDB2874818
                    BA7BBA1F44CD37496FD0B18F208F0105FBE4482799460C2CCAF3FAD2113AC4DF
                    239F9EA26A84E0A9853D69125CB457F1C11EC39FD71514E4CDCCCDCE56757476
                    B2D2331D65E834F1F8017303711143BCC116C4149C7C42EA1F1F55348E0A89CC
                    428C61A62744F14CBD02E10793C1C8F21EE4AA32ED427D12895CD2E7A2669DC1
                    AE8E4E5F7D53ABED647B50298A89B0F2FA1BC064D7A3360880DE1801BFA701F6
                    EDD982F710244DD004F106153E2E0C01FEEF73A843515E1321334F0DC1407095
                    E394E399507FE88EEC9C2CF3B4A99397A399F820180808FCE051911A8D90861B
                    4DA038625E833CE32609064D8E227F2410E8FF2B932DB6166A30E538D5569A91
                    9969A47E0514E3EFEEEA89DB7DA5C00021B984D2E22AB83B407C841848BA1FCE
                    34E92253428243A15D230251623D1176600770A0B03871D7BBFBDCBE2EDC862D
                    CDCDF5B5C78EEDAB3E72E8786B6BE77F145E76C31CB5DA00A9FB8FC0DCB9D311
                    8344A0FED821D8B8E145709CAC3D855FBC09470B0E7227BA71B870F8B92044CF
                    F05E215E32979E96117715C5F8296FE415A1608450C01F42CFA9064DC7DA9C9C
                    ACC271E30A0A2921C74E8B09065C3EAFBF9AB5E91B654918B5B3A179736E1B2E
                    CC78B48965E82AADCAC9CE9A8186517BBAB383F5294E4F4D63C89D8A56422C0B
                    29B283B4E00C2741C969640A767E338B35A01741DA81848E42D554D748492FB2
                    D1E4A61220F5FB7C6E5CFCD6134D0DB54D4D4DE575B5B5BB4E9C68A473250C38
                    A881E546A7A33C79C2948579FFD8F8B6F6E30FDF8570BF0B4E3757E2AE7751AF
                    A2B7701CC24175054E1CD49F20C8C15E4CAE114A274F81929252B8F1A69B19C9
                    96EE999A7A27994CCCABA1809A3631F125743BAB108CBE69369A2650DC845540
                    3BFB5CF87FF5C651C607232E089C09A4469547BD01A66B13B5CB119C5D83B63F
                    89BC01E208102227B66F6343531DBA5AE6FCFCDC8CA27105CC6DA4C5944819E4
                    C201A79BB10EAE5A1DCB4710CAD7B2927260717F571FF550F250A5511FAAE206
                    54C9D5781FBB36FECFDBDB776EDFD6C46F8D0042218E3C1C74B878D4E5EC7A67
                    7FC586C9C9F6F4D456875F250ADE5EFC7D150E6A6AE9E0E6C0CDB540FF60D860
                    C5CA1BE0FA1B57B3C2986E34755284524E8FC3F9A6E2DC4B1007CD44DC13A57B
                    36E13C2801861B4583A6D20717E11A1141889FEE2E6421409B6DB35A17E2C22F
                    42943E4997A8676A9AF202AE780ADAE9F3F90F2188FA6754886E6D6E69C94435
                    FA547676E6245A6412148AE0C5D3C40A06E04863B0B311282A49A96744FAC423
                    A41233AFC7DB8D6E5D6D382C54521F017CB07B7272724E3635364053C3992664
                    9499B292EF2E65D2F9E8C01575387B3A22DCF61330EC940143C90C44066A8139
                    577C0B3D101BDC72EBED70A2F904E36128A4C341E37D9214D4CD05EF7BAED160
                    5D6CB559AF4AB1DB2D949720120C7929542AE7F3FB2AD09D6D8F7DDD0581CE47
                    46119F63494A2A33994D0B4C46E37CDCB13A5A309A2CB998D4F81205E1203594
                    C6BF2DA736BD88007AC8965350E994C3D1E6F5FB9E4615BA88D2D994D123FB4F
                    7D8CCCE8F211D6A08013D1BB28D28860F434AADC4A5CF403F8D03F4521388C2E
                    709BD47002FD7EF8EB1BEB11F9B729F87C45BE98FD7C8747B9BA07EE02FA6540
                    502E00824C03302D30EFCA0540CCECA5CB57B0821B625F9D759D45C22EB93887
                    792929F605E8AA96A5DAED2554D3407841EA02EBF5F85C880DB6A2F7B20D01EE
                    C660ACDF4987918E66C269580481B78B294221A0330E97E0E2CF423FB990E211
                    74303895A8D3EE27F710777D270AC4F3B8909BF07D355A0AE6A03B475050C351
                    3EBA5BFBFB9C7D6B70F17F663699EEA4C52490495E04B5ADE1954FA751EDEFC5
                    45DFAB508807C3827054212A3AA953A954E6CEF43F7EDEC103FBE1640BEBF62F
                    F9FC02B7EF4ABEC04EAE1904997608F157293E108501DCA36B565C074B965DC3
                    802995F251424B43A7C80B4202CE7B01FEFEAABCBC5C3A2271B6C56AD1D1BD90
                    A644ADC54C171D85849A228AE6EF25D48A4FE9F489EE545B1A9D5C03870E1E84
                    AD5BB78CE8C25FBF72D9F0090202B785F8006EB4596D7371B74ED668121209CC
                    F91014F522FA27A046CD3589BC42BD8B3446CD218B25E98FADA71C5E8A1ADAAD
                    29F0FAFA17990F4FEEA1C96082EFDF7D2FC5E01D1E8FFBA15EA7B32112151EC1
                    076573BB3C6DC160703B82CA5D64BBF1A1D689A2A29BD2B6EC74747170216534
                    789923C31757DAFD727B2C0C880BC4066200A605E697416E5E1E2C5EB28CE54D
                    E8BC085EFB5882F35F929D95B938D56E9D9A9C6CCD35720A1AABCCA69E0DACFC
                    4D60893116DBD0A845A3D1F00F5C7C776D7D1DBCF9DA7A96646B6A6A04F46A2E
                    AE46D00C81C54C8B8613FF2EDABC9F656464E413F3988248D1988FC5E2896442
                    769C2862143FA789B3DCBDD1546349B20466CE98096BD7FE90752B3B5275F833
                    9F4DC5234434B9F3AE7B508E3CCFE2F7ECECE9EDB54785980317B639168B3AC9
                    3D8B97AC5FB08B25CA76774446C812071967AECBA64E83B9F3CA5885736E5E01
                    D362429C1463C00F586EB7DBAE459B3F075FC7994CA644CA8950D6925AEBB034
                    79732BB510007B8A9D3D5FE91927EA741D3A9DBE8AF010D584EEDC510117EB3A
                    67D5EB1A9A868A07D6E0EECE4F4B25BE819EA9EDF2F272A8AEAA82ACEC6CB8F5
                    B6DB50104C67EA004803ECDDBBA7F5BF376CA08EAD70F8F021CE35F8EC250906
                    35AAFCC1FD0FC4BA3A3A76B1923385F28CFBF855AD191706856CC11530884E29
                    1A3F0156DF7C0B2E981EAC363BE8131359338EBEBEBEE9C996A49BF2727396DA
                    EDF602FC39855C402AD2219795B8115E9FB772CFEEDD0907F77F5A4AA1EB9B6F
                    B905269516C733A29CF0DAD2DCE27D6DFDAB221D17488DBB2EE6758E2050B5CE
                    10EDC29F11F54F41D46BA3A654540A3EA9A404766DAF80CAC30760D9B2A55090
                    97C3001DD9F88E8E0ED8B0E1ADB53BB76F6FC577BF235730832C146CF9F84396
                    13F8DE9D77D199515F7A8212A7F13C0271CECFD463E9BEFB1F64390C137A2916
                    FC37959E79DC6E1B7A31ABACD6E41B53ECB6CB10FCA5E9F43A3A6F98FD3FD540
                    A2364419F0BF8FBFFA3BE2A4FD8D8DF54B0F1EF8745D3426EAF30BF361D1D50B
                    59BD027580331ACDA801CA0BDF7B6FD346FCDA7FC1D175BE6771510461C88596
                    A2F836DAED16B479EBD0359A4DBD8A66CE9E053FB6FD049E7BEE5928DFB60D16
                    962D60FC40CA1934363442DBA9F66C7CE71F217ECEC0A7C0CA50065D2062EB89
                    EF6EFA07D31AAB6E5ACD10F6D0C2DCF19FE35E831A965F732D2C5DB60CFEF6D7
                    3760DFDE3DE7BCE7A11F3D0A7F5AF73B78ECA74F704CA1639D58882710150415
                    0AFB95A8EEBF939995B124C5664F35990C461624A2025862437B99DBBA1FEFF3
                    6F78071FA1DBD866329A3D7B76EF828A6D5B2B701ED538E6ECDEB913BAEEE882
                    A9E969100AAB5921CBE9D37DC46E598AE3D738D648F9B9019846BC28827001A9
                    4F01D5F45E54852BA2EDB167C598783B2276554141213CF9E42FE0E8D1A3A842
                    9D60494E6639FDFEB002F4C674FCFCE66C518CDE8AEF3FCA015A824C0062323F
                    9DCC8FB81935031583DCB06AF51942CB6017999FCCCC74E6A11040A30EA73417
                    03E714DE73CF7DCCCF8F57228799C0988C7AC685FCC52F7FC5B39F0A7E4CA0A7
                    50AD52AEB6582CB7E415E44D484EB2E834E8D6D0C15BD49185EEC3EDF1B6A326
                    D8849ED006FCBE83B8D841C44E029DD1B4637B39BCB6FE25C21053F0894EA4FB
                    EB625D52EAE1F2D9B399FB48C49618982035AD00BA3A9B97E39FCCC24147FBEA
                    64A075D484E1AB0882548AD68D13FEFE8E9D3B1C2525258F4D183F5E491DCCE6
                    CD9B071DDD3D8CE0919C6C86EEEE3E309933212D2D07CD444B19BE7B1CD70C06
                    3E6981FBF6C085212A69A8AD9F6C81EDE5DB60F9B5D741D98285676A11E81911
                    E99542B76A969D54CA0A61401ED0016D92194CACF620C6CAEA9820980CACAED1
                    6030AAD18418519816236ABF2D2B73C242EAD588A64925C44F9B81AE6E7F1477
                    7ED01F08ECC3F9BE85DAE27D44FE547C221010260D545B5B032FFC711D332B28
                    C4B8DB95D7A1D2B3D0ED58ED0570ACF62473A3ED686E7A9A1D8823225034611A
                    7477B518F173AEE49B23893F87D0F9125AA3220894AA1DA2696021DF04F40E7E
                    F79B5F09E846B5E517148B6BBEB706A64D9F8C93B5B18C20F115684777F5F840
                    6F4A015B4A067476B4A4E3BBC79393C247880770147CE2F21D21E529E0DD7736
                    3277F0C61B6F62E04A62120F5CFCA16447D1F2A8A80126BEBF1435D6AAB4D4D4
                    6FDBEDD6226AE4218AF17E4D1D1D5DA2DBEB41390838FAFB83EF4422C2DFD063
                    39226F61439A8832ACADAD2DF0FCEF9E937F650E28351321A684246B265C3EEF
                    7A70B4BBA1B1BE19F2176723E8F483A8D44356561E6930D436E1347C4F1ADF18
                    2A59CC23267F16A32608CF3CF5C4976126D098D2D2DCA6F8EDEF5F847FBB6B0D
                    5C317B32226D2B63F9122F21260AACE7416F6725EB5C876FC8944DD2CD7F0E4A
                    1F763E244FC0936A2368975E68CA9E53C368E7A7227E58888BFA9DCCCC4C3A30
                    4B497C054A0E9157803B3FE0F1FA7B82C1C0EE7038F4564282F6037C4FE45C12
                    8F064EA0CFFFEB679E1AECEB12158A04BDA84A44C0990305E38A19CE3950791C
                    E6CC9901A73BBDE84E66824641FD13C322D702A9FCBD212E0CA3464A18A61033
                    BBDF7EB53105826125D4D4B9202B4740D5D70AA1F454D6F27EF992D970BCAE05
                    0EEC65ECAC4E6E0BAD3CA3A7900570A2E753878C5144DD48A2D10B5D7C0DBE27
                    1D771E25BF56D9ACD69B10FDA792FF2EC47B1C93CD8F518B61BFCF571F094736
                    A12BF8366A8BD681BDA9E9F3C80C90466C3ED134981028F9FD07504BF9558929
                    A04C3082D1AC87FC7105E0F776C1968A6A6838E982D24945F0CF0F77027F0627
                    B936F0704D208C9659182E416049604AD3522F138B3D1FC282087A4B0104FA3B
                    511DF632722A751C9931BD18EEB8630DBCF4622878B4FA88640EFC3CC2E7E31A
                    213A30A9C3DC3A1480471F7B1C261697323AD950B810425440FF5E4369EF857A
                    83619535D9328B3C18BAC8E523FA3C823F17DAFDE6607FFF56B4FBEFA148966B
                    35DA78B67310A1A2D63847ABAB98D7F087E77F3B98102470C1F6C684609D35D9
                    362F885E422F6A9AE933A6825A9383E6B29771241B6BF7C1CEED74180B509FC5
                    533CAC2DE539C29FB7292E3541507215469FB33B1AEA6EB2D9E78D6B3FDD031D
                    EDED70D9B429D01F74427B5723AADD16568ABE72E50A983061C2CC175EF853E7
                    279B3F5A8F6E57030787922044B84938631A2C9664B8F7BE0760C6CCCBC1E138
                    754ED79381892F955A5D8CDEC24CA3497F8DD562B9362929C9A066A1DE306BB3
                    43BD95BD5E5F3DE28C4A747F3F8AC6A2EF23C8E8A250F8F94A4E490888E9545F
                    7F1CD63DFF1BC6721AA012951CEF50B62839AEF1228763FD6DA76D59051955D5
                    8D90949C0A45E3F321E8F7C09143BB61E35B7F46ECE3DFC1F90E1E4E7A717341
                    888CA646500C64C25C207356CD7780296E13E1E68CFC198F174C9C6FF1790498
                    FBAD59A846557014276DD6FBE0DBDF5E094545E3A9A31803586FBEF9E63B7FDF
                    F0D6CFCD26F3A1B676871AD5B2289BBC48AAFBAA858B50854E82F9572E60DC45
                    DAC9F182D3307BFC94A5A342119552956B4E324F47377611BE2EC1EF28D6B183
                    3AE3E727BADC2E70BBBCDD3EBFBF4A1022E574E8369D784BCDAF29354E5F1C3F
                    751D55BF3A81E10589884B80944E943F5259C900216983F368021D1782748E81
                    EC38E6178CBF6C49A2A930C5E512145AAD1A029E0EE86A3B8CE6314442B09E9B
                    05BF8CF012194C2B0E375764B80441C905C1209B3CE5FA57D0C405D19CD5D9E1
                    D28643281191B61EDCAB07AF2C5BE0BEFDF6EFAE282E29B6B35E83F8A0FFB975
                    EBEEBAE3C71FFFDF8D6F97E7E5E5C3BEBDBBD9879310DCF7EF0F320E616F4F37
                    6B6641E681D52E006BD44DCCE4441484D9A801CA4C06C3D526B369BEC9684820
                    3792BAABBAD8F13CD4F6C6BB3FD81FDA198D08E51121BAC360D0F5B1763C547E
                    AF49F85C41A0E743E73B1E3D52C54CC120B10C05D78A3AEEFAA57221C8E23F1B
                    E9D9A069CA51AAF4269FC785521469E101359A6C2F07875E59BA5B1CE9D8C170
                    0B82A40D6CDCF5498138FB274BAD56642B95095A44DD34410A2B1FA3D7E2E292
                    C9FFBAE6CE1FCE997339652C59C875CFEEBD2D1515154F97944EFE4BF9B64F58
                    3C203F3F1FAE4321703A7BD99751BC9E54336534C5A8508480ED2AA3D1B808DD
                    BD394693B180E2089148FC14145A7CB7CBDD1A08F657E0EFCA4531B6271A156B
                    59438B688C752419AA2090C6A93D5603EB5F7929DE1AF7FC1B423209921B48CF
                    8428EF5A59EA3BC8179E6872DD7CE1FBF9EF05999B38E201A4911004039FB095
                    EF081D9F48984F52028461AEFA7A3233B3C6DFB4FAE687CBCA16DC43D472D20E
                    DD3DBDC1AA23D52F4785D813E81E3AA9B51E7D0C3F268FB29B065CF425BA44ED
                    5283413FCB64324FA5E379A8B8858483228AE8F685BD7E7F05A27E045F8A9D91
                    48B812D57A80543B3B56872AB62F5010C853F9D31F9EA7C2952FD28C3AFE1C92
                    F9AB813F1FE95904380E9080B1C4808E0CC677186D41F82A6051B2E761BED022
                    7F55CAA28321D990881ECAF6F6B6869AA3471F2D2E29AD4355FD24EE76B3D56A
                    D14D9D32E9FE63B5F51310613FA6522A0EF1739267E97589D75A93AD0BE274B7
                    44DA6DCC4450E713C207AC5197201CF705020F20DA3FAE10E1143BBD4D50B04E
                    667A9D9E698A0B6D3A21F119A8A4EEF31C1459F0272023BC487E6754B6294232
                    F24B6434730923ED3544F9E444FEAA18101D1464AF208F90E103F6C462E2BABE
                    BEBE9A507FF0F7E148A4944E86993675CAD2BAFAFA340478EF64A4A71526275B
                    E66AB5DA7C34172A76A612BA8E3DBCE0955AD55008DB68305099D916A7CBBD85
                    763DF56A2210480BB86B47056CAFD8060B10745E597615D31E748683C7E3FEC2
                    7804E5293EFEF00338B07FDF509E4358C6739082415199900C24BD5C1202309C
                    1A417A0D0DC8EDCBFF5F1C3C55A188E1AEDBE2F1FA6E88088EA7114FACA67677
                    934A8BA73A1CEDA54A954A433B92DADE53632B62F9505F236A809D92628FB7A2
                    53AAA938366A341A76178E2B84A6C64678E4470F33179308A44EF4DF491B9C3E
                    DD0E9B3FFA80990122C2DE75F7FD90929ACACC01B5C11FF4E1A0CB495C014A9E
                    0DE159083038E145FE0C2EA9C51F4E8D30D804C5F38587CF1B9304688C4484BB
                    DBDB3BA946F0BFD2D3D3157979B91AEAC3D4D4D40C540575EC6835EBDB38E78A
                    B9B06CF952467D933E801ACCBCF2F24B0D9B3FFE98CE4062DC878117F104BDB2
                    05FFCD734FB33238D2203F7EFC89786B5F45BC89A7DC340CBDEFE419C2CB977A
                    0E5F774180F3A0DC0B9E3C6A0727AAEAA7BB7B9CD5881B9E417CA0ED38DD790C
                    55F9AB9B3F7ADF7CE8E0FE5FA20760A7C6124B165F0D469B817909445BABAFAF
                    D3BCB6FED53F84C3E13BF0A38E0F6501E4D1C9FFFCC95AF60E22D8AC7DECA720
                    84852F5B89FC959FC3D75D1086EB0AE1E6DEE8F5FA360783417534160B84E8D8
                    B760D086586215FEFFB5FBD15E9F3AD5CA5AD41148243BDEDEDEA140A1988BFF
                    FF328E251CB4A90600BAF32E8AA4298888FAC00FEE667F3865CA5478F891B50C
                    877C132EF525784F04FCBCE4D66D2F2F87375E7F8554B49907AB98EAAFACAC42
                    ADB0987110A8DED1E7172135AD083A3B1BA7E11FAFC13F7B81FBF5E28004CE17
                    46EAA4A8E1C1039FC2EDDFB979D46B102FD6A5BC685F4C7D8A58230C69C41B6E
                    B283B57154546C85D75F7B992F84622528D4F3E3004E07C7EB1C8C354D076610
                    23D8E1E881BC822960329AC9975FC6835C66FEAAE3419D2F38B1E38B7DED318D
                    30024240FE3F2F3C61358B2404227F25A4FEEACB7F96FE1C1752350B442A76D5
                    C0F4D9D78127A0829A637570CDF285E80606C1ED8980DD9E01365B9AC2EB7553
                    78B75016C851C3D95E06979CDBF68D1604CAE713877F883C7E2BA8D4A9206AC1
                    68B143D1C4E9EC6CC7BDFB6B60E9E2F9681642108A6A20232B0B6AAA63924F6F
                    E78BEEE1C2A094C538C6AE4B451006AB67F89C0BFD726554A131834E67035392
                    198A32F2A0ABBD0EAA8F35437B67082CD634B09A43D0DDDD418B4F3CC844385B
                    B22E5CAA419C31B038B44BC517CE0BD188C368B5B33636D40BA9A08048B016D8
                    57ED44AC1041845F00EF6D7C050254661DCFEAD1A2FBF990FA188C1AC9630C2C
                    0EEF7D9190F22456E4B05A11F259EC995053D3042E9713C615E542466636245B
                    F4F0E99E6DF0C9665633F3318EC3DC245059BB0BCEB27DC684E0F3E2385F9198
                    32ECF7C387C4F4A18C2665F332556AEDF78AA72E5AA5D1E582D96C8589C5E321
                    2604A1B66A37ECD9B1091585E77DEE36121FD20D9FED6A32EAD9BD4BFD1ACE34
                    F44809829ADB7889E491C5C15FA64EA72F4BCF9E38C9174C48F57A839AFE40AF
                    0F226D4475A31AF277393E08722F41DED462ECFA9A09C240DE9F44F248E5BFA3
                    FCBE45ABD5EA2311EA3A23D0C2D7E2A8972D7E70802918D3045F538D2071FF2C
                    32D360E041A128B7F952371337C4295E527E5F1280E8D8527FFD0541C517DDC8
                    B5805EE6DD44E02CE1454EF41871B2E798205C3C8F2181BF4A67DD0D4676198B
                    0FFC3F16044918A49A09E9929360C4B1C5FF6608826426CEB9FFB1251C254118
                    BBBE99D798208C5D638230768D09C2D835260863D760D7FF093000DF9C3AC9EF
                    D8ABAF0000000049454E44AE426082}
                  Version = '1.1.1.0'
                  Visible = False
                  ExplicitLeft = 6
                  ExplicitTop = -6
                end
                object v_tipo_normal_campos: TRzPanel
                  Left = 0
                  Top = 304
                  Width = 569
                  Height = 200
                  Align = alBottom
                  BorderOuter = fsNone
                  Color = 16384
                  TabOrder = 0
                  Transparent = True
                  object Bevel3: TBevel
                    Left = 0
                    Top = 90
                    Width = 569
                    Height = 10
                    Align = alBottom
                    Shape = bsSpacer
                    ExplicitLeft = -6
                    ExplicitTop = 86
                  end
                  object v_tipo_normal_campos2: TRzPanel
                    Left = 0
                    Top = 100
                    Width = 569
                    Height = 100
                    Align = alBottom
                    BorderOuter = fsNone
                    Color = 16384
                    TabOrder = 0
                    Transparent = True
                    object Bevel2: TBevel
                      Left = 293
                      Top = 0
                      Width = 10
                      Height = 100
                      Align = alRight
                      Shape = bsSpacer
                      ExplicitLeft = 297
                      ExplicitTop = -4
                    end
                    object pnUnitario: TPanel
                      Left = 0
                      Top = 0
                      Width = 293
                      Height = 100
                      Align = alClient
                      BevelOuter = bvNone
                      BorderWidth = 10
                      Color = 8536320
                      ParentBackground = False
                      TabOrder = 0
                      object RzLabel1: TRzLabel
                        Left = 10
                        Top = 10
                        Width = 273
                        Height = 16
                        Align = alTop
                        Caption = 'Valor Unit'#225'rio'
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -13
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        ParentFont = False
                        Transparent = True
                        ExplicitWidth = 88
                      end
                      object ed_unitario: TRzNumericEdit
                        Left = 10
                        Top = 53
                        Width = 273
                        Height = 37
                        Align = alBottom
                        Alignment = taCenter
                        Color = 8536320
                        Ctl3D = True
                        DisabledColor = 8536320
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -24
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        FrameColor = 8536320
                        FrameHotColor = 8536320
                        FrameHotStyle = fsNone
                        FrameStyle = fsNone
                        FrameVisible = True
                        FramingPreference = fpCustomFraming
                        ParentCtl3D = False
                        ParentFont = False
                        ReadOnly = True
                        ReadOnlyColor = 8536320
                        TabOrder = 0
                        OnEnter = ed_unitarioEnter
                        OnExit = ed_unitarioExit
                        OnKeyPress = ed_unitarioKeyPress
                        IntegersOnly = False
                        DisplayFormat = 'R$ ###,###,##0.000'
                        FlatButtonColor = 8536320
                      end
                    end
                    object pnTotal: TPanel
                      Left = 303
                      Top = 0
                      Width = 266
                      Height = 100
                      Align = alRight
                      BevelOuter = bvNone
                      BorderWidth = 10
                      Color = 8536320
                      ParentBackground = False
                      TabOrder = 1
                      object RzLabel4: TRzLabel
                        Left = 10
                        Top = 10
                        Width = 246
                        Height = 16
                        Align = alTop
                        Caption = 'Valor Total'
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -13
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        ParentFont = False
                        Transparent = True
                        ExplicitWidth = 70
                      end
                      object ed_total_item: TRzNumericEdit
                        Left = 10
                        Top = 53
                        Width = 246
                        Height = 37
                        Align = alBottom
                        Alignment = taCenter
                        Color = 8536320
                        Ctl3D = True
                        DisabledColor = 8536320
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -24
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        FrameColor = 8536320
                        FrameHotColor = 8536320
                        FrameHotStyle = fsNone
                        FrameStyle = fsNone
                        FrameVisible = True
                        FramingPreference = fpCustomFraming
                        ParentCtl3D = False
                        ParentFont = False
                        ReadOnly = True
                        ReadOnlyColor = 8536320
                        TabOrder = 0
                        OnEnter = ed_unitarioEnter
                        OnExit = ed_unitarioExit
                        OnKeyPress = ed_total_itemKeyPress
                        IntegersOnly = False
                        DisplayFormat = 'R$ ###,###,##0.000'
                        FlatButtonColor = 8536320
                      end
                    end
                  end
                  object v_tipo_normal_campos1: TRzPanel
                    Left = 0
                    Top = 0
                    Width = 569
                    Height = 90
                    Align = alClient
                    BorderOuter = fsNone
                    Color = 16384
                    TabOrder = 1
                    Transparent = True
                    object Bevel1: TBevel
                      Left = 369
                      Top = 0
                      Width = 10
                      Height = 90
                      Align = alRight
                      Shape = bsSpacer
                      ExplicitLeft = 325
                      ExplicitHeight = 65
                    end
                    object pnBarra: TPanel
                      Left = 0
                      Top = 0
                      Width = 369
                      Height = 90
                      Align = alClient
                      BevelOuter = bvNone
                      BorderWidth = 10
                      Color = 8536320
                      ParentBackground = False
                      TabOrder = 0
                      object lbCodBarras: TRzLabel
                        Left = 10
                        Top = 10
                        Width = 349
                        Height = 16
                        Align = alTop
                        Caption = 'C'#243'digo de Barras'
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -13
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        ParentFont = False
                        Transparent = True
                        ExplicitWidth = 110
                      end
                      object ed_barra: TRzEdit
                        Left = 10
                        Top = 37
                        Width = 349
                        Height = 43
                        Text = '1111111123456'
                        Align = alBottom
                        Alignment = taCenter
                        Color = 8536320
                        Ctl3D = True
                        DisabledColor = 8536320
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -29
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        FrameColor = 8536320
                        FrameHotColor = 8536320
                        FrameHotStyle = fsNone
                        FrameStyle = fsNone
                        FrameVisible = True
                        FramingPreference = fpCustomFraming
                        ParentCtl3D = False
                        ParentFont = False
                        TabOrder = 0
                        OnEnter = ed_barraEnter
                        OnExit = ed_barraExit
                        OnKeyPress = ed_barraKeyPress
                      end
                    end
                    object pnQuantidade: TPanel
                      Left = 379
                      Top = 0
                      Width = 190
                      Height = 90
                      Align = alRight
                      BevelOuter = bvNone
                      BorderWidth = 10
                      Color = 8536320
                      ParentBackground = False
                      TabOrder = 1
                      object RzLabel3: TRzLabel
                        Left = 10
                        Top = 10
                        Width = 170
                        Height = 16
                        Align = alTop
                        Caption = 'Quantidade'
                        Font.Charset = ANSI_CHARSET
                        Font.Color = clWhite
                        Font.Height = -13
                        Font.Name = 'Tahoma'
                        Font.Style = [fsBold]
                        ParentFont = False
                        Transparent = True
                        ExplicitWidth = 74
                      end
                      object Panel9: TPanel
                        Left = 10
                        Top = 41
                        Width = 170
                        Height = 39
                        Align = alBottom
                        BevelOuter = bvNone
                        TabOrder = 0
                        object btnDimQtd: TSpeedButton
                          Left = 0
                          Top = 0
                          Width = 41
                          Height = 39
                          Align = alLeft
                          Caption = '-'
                          Font.Charset = ANSI_CHARSET
                          Font.Color = 8404992
                          Font.Height = -24
                          Font.Name = 'Tahoma'
                          Font.Style = [fsBold]
                          ParentFont = False
                          Visible = False
                          OnClick = btnDimQtdClick
                          ExplicitLeft = -6
                          ExplicitTop = -1
                        end
                        object btnAumQtd: TSpeedButton
                          Left = 132
                          Top = 0
                          Width = 38
                          Height = 39
                          Align = alRight
                          Caption = '+'
                          Font.Charset = ANSI_CHARSET
                          Font.Color = 8404992
                          Font.Height = -24
                          Font.Name = 'Tahoma'
                          Font.Style = [fsBold]
                          ParentFont = False
                          Visible = False
                          OnClick = btnAumQtdClick
                          ExplicitLeft = 208
                          ExplicitHeight = 46
                        end
                        object ed_qtde: TRzNumericEdit
                          Left = 41
                          Top = 0
                          Width = 91
                          Height = 39
                          Align = alClient
                          Alignment = taCenter
                          Color = 8536320
                          Ctl3D = True
                          DisabledColor = 8536320
                          Font.Charset = ANSI_CHARSET
                          Font.Color = clWhite
                          Font.Height = -24
                          Font.Name = 'Tahoma'
                          Font.Style = [fsBold]
                          FrameColor = 8536320
                          FrameHotColor = 8536320
                          FrameHotStyle = fsNone
                          FrameStyle = fsNone
                          FrameVisible = True
                          FramingPreference = fpCustomFraming
                          ParentCtl3D = False
                          ParentFont = False
                          ReadOnlyColor = 8536320
                          TabOrder = 0
                          OnEnter = ed_qtdeEnter
                          OnExit = ed_qtdeExit
                          OnKeyPress = ed_qtdeKeyPress
                          IntegersOnly = False
                          DisplayFormat = '###,###,##0.000'
                          FlatButtonColor = 8536320
                          ExplicitHeight = 37
                        end
                      end
                    end
                  end
                end
              end
            end
            object v_espacamento_central: TRzPanel
              Left = 569
              Top = 0
              Width = 32
              Height = 504
              Align = alLeft
              BorderOuter = fsNone
              Color = 6447714
              TabOrder = 2
              Transparent = True
            end
          end
          object v_espacamento_inferior: TRzPanel
            Left = 0
            Top = 581
            Width = 1498
            Height = 24
            Align = alBottom
            BorderOuter = fsNone
            Color = 6447714
            TabOrder = 3
            Transparent = True
          end
          object v_barra_inferior: TRzPanel
            Left = 0
            Top = 605
            Width = 1498
            Height = 139
            Align = alBottom
            BorderOuter = fsNone
            Color = 8421440
            TabOrder = 4
            Transparent = True
            object cpBarra: TdxStatusBar
              Left = 0
              Top = 118
              Width = 1498
              Height = 21
              Panels = <
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Operador'
                  Width = 55
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  Bevel = dxpbNone
                  Text = 'Sergio Luiz'
                  Width = 100
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Caixa'
                  Width = 40
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  Bevel = dxpbNone
                  Text = '001'
                  Width = 30
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Servidor'
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  Bevel = dxpbNone
                  Text = '192.168.0.100'
                  Width = 80
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Internet'
                end
                item
                  PanelStyleClassName = 'TdxStatusBarStateIndicatorPanelStyle'
                  PanelStyle.Color = 14540253
                  PanelStyle.Indicators = <
                    item
                      IndicatorType = sitGreen
                    end
                    item
                      IndicatorType = sitRed
                      Visible = False
                    end>
                  Bevel = dxpbNone
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Status NFC-e'
                  Width = 75
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clGreen
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Envio Normal'
                  Width = 120
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWindowText
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'Carga'
                  Width = 40
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  Bevel = dxpbNone
                  Width = 200
                end>
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'Tahoma'
              Font.Style = []
              Color = 14540253
            end
            object barraAtalho: TdxStatusBar
              Left = 0
              Top = 94
              Width = 1498
              Height = 24
              Panels = <
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F1 - Op'#231#245'es'
                  Width = 70
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F2 - Produtos'
                  Width = 80
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F3 - Fechar Cupom'
                  Width = 110
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F4 - Cancelar Cupom'
                  Width = 120
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F5 - Cancelar Item'
                  Width = 110
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Alignment = taCenter
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'F6 - Consulta Pre'#231'o'
                  Width = 110
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  MinWidth = 20
                  Text = 'F7 - Consulta Cliente'
                  Width = 118
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  MinWidth = 20
                  Text = 'F9 - Pr'#233'-Vendas'
                  Width = 95
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  MinWidth = 20
                  Text = 'CTRL+E - Mesas'
                  Width = 100
                end
                item
                  PanelStyleClassName = 'TdxStatusBarTextPanelStyle'
                  PanelStyle.Font.Charset = DEFAULT_CHARSET
                  PanelStyle.Font.Color = clWhite
                  PanelStyle.Font.Height = -11
                  PanelStyle.Font.Name = 'Tahoma'
                  PanelStyle.Font.Style = [fsBold]
                  PanelStyle.ParentFont = False
                  Bevel = dxpbNone
                  Text = 'CTRL+A - Consumidor'
                end>
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWhite
              Font.Height = -11
              Font.Name = 'Tahoma'
              Font.Style = []
              Color = 6960384
            end
            object Panel12: TPanel
              Left = 0
              Top = 0
              Width = 1498
              Height = 94
              Align = alClient
              BevelOuter = bvNone
              BorderWidth = 10
              Color = 8536320
              ParentBackground = False
              TabOrder = 2
              object lbDevolucao: TLabel
                Left = 6
                Top = 23
                Width = 234
                Height = 19
                Caption = 'Devolu'#231#227'o/Troca de Produto'
                Font.Charset = ANSI_CHARSET
                Font.Color = clYellow
                Font.Height = -16
                Font.Name = 'Tahoma'
                Font.Style = [fsBold]
                ParentFont = False
                Visible = False
              end
              object lbDestaque: TLabel
                Left = 6
                Top = 48
                Width = 205
                Height = 19
                Caption = 'PRODUTO EM DESTAQUE'
                Font.Charset = ANSI_CHARSET
                Font.Color = clRed
                Font.Height = -16
                Font.Name = 'Tahoma'
                Font.Style = [fsBold]
                ParentFont = False
                Visible = False
              end
              object lbEstoque: TLabel
                Left = 6
                Top = 70
                Width = 165
                Height = 19
                Caption = 'Estoque Atual: 1542'
                Font.Charset = ANSI_CHARSET
                Font.Color = clLime
                Font.Height = -16
                Font.Name = 'Tahoma'
                Font.Style = [fsBold]
                ParentFont = False
                Visible = False
              end
              object Panel13: TPanel
                Left = 1103
                Top = 10
                Width = 385
                Height = 74
                Align = alRight
                BevelOuter = bvNone
                BorderWidth = 5
                Color = 6960384
                ParentBackground = False
                TabOrder = 0
                object lb_total: TRzLabel
                  Left = 5
                  Top = 27
                  Width = 375
                  Height = 42
                  Align = alBottom
                  Alignment = taRightJustify
                  AutoSize = False
                  Caption = '0,00'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -40
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  ParentFont = False
                  Transparent = True
                  ShadowColor = clBlack
                  ExplicitTop = -8
                end
                object lbTotal: TRzLabel
                  Left = 5
                  Top = 5
                  Width = 375
                  Height = 16
                  Align = alTop
                  Caption = 'Total:'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -13
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  ParentFont = False
                  Transparent = True
                  ExplicitWidth = 37
                end
              end
              object Panel14: TPanel
                Left = 935
                Top = 10
                Width = 156
                Height = 74
                Align = alRight
                BevelOuter = bvNone
                BorderWidth = 5
                Color = 6960384
                ParentBackground = False
                TabOrder = 1
                object lb_item: TRzLabel
                  Left = 5
                  Top = 27
                  Width = 146
                  Height = 42
                  Align = alBottom
                  Alignment = taRightJustify
                  AutoSize = False
                  Caption = '001'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -40
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  ParentFont = False
                  Transparent = True
                  ShadowColor = clBlack
                  ExplicitTop = 4
                end
                object RzLabel7: TRzLabel
                  Left = 5
                  Top = 5
                  Width = 146
                  Height = 16
                  Align = alTop
                  Caption = 'Itens'
                  Font.Charset = ANSI_CHARSET
                  Font.Color = clWhite
                  Font.Height = -13
                  Font.Name = 'Tahoma'
                  Font.Style = [fsBold]
                  ParentFont = False
                  Transparent = True
                  ExplicitWidth = 34
                end
              end
              object Panel15: TPanel
                Left = 1091
                Top = 10
                Width = 12
                Height = 74
                Align = alRight
                BevelOuter = bvNone
                TabOrder = 2
              end
              object btnFechaCupom: TAdvGlowButton
                Left = 773
                Top = 10
                Width = 162
                Height = 74
                Align = alRight
                Caption = 'Fechar Cupom'
                Font.Charset = ANSI_CHARSET
                Font.Color = clWhite
                Font.Height = -11
                Font.Name = 'Tahoma'
                Font.Style = [fsBold]
                FocusType = ftHot
                NotesFont.Charset = DEFAULT_CHARSET
                NotesFont.Color = clWindowText
                NotesFont.Height = -11
                NotesFont.Name = 'Tahoma'
                NotesFont.Style = []
                ParentFont = False
                Picture.Data = {
                  89504E470D0A1A0A0000000D49484452000000280000002808060000008CFEB8
                  6D000000097048597300000EC300000EC301C76FA864000000206348524D0000
                  7A25000080830000F9FF000080E8000052080001155800003A970000176FD75A
                  1F90000008724944415478DAC4D9DB7313D71D07F0DF4C269D49673AD349FA07
                  347D6BFA929776A67DC824D399A643C0E152533B34144A42C026B8A48158C650
                  68EA1BB6A5BDE982EE360E58DA5D49088CF15596718040A04DC0B2916C631B63
                  5BBB928C0571621CF7DB07D5062339711C531EBE23CD4AABF9ECEF9CB3E7EC11
                  0502A087133C076A6F9D258953C96B9C20AF254E3217A3D3967BE46113E4334E
                  90C4A924B1CACF4556C91739C52EF16A9BC82AC322AB244556498A8C322A0B6A
                  A7C8297522ABBC2F73CA8B12A7922CA8E4616274D27487FCA624798418798C31
                  1219855A1A6728D8B5D012088068A9C006DB17E4E512CFCA7C4C23716A8FC82A
                  988B9B79F07EFE982E0AF1A1E312A70E489C52E261623FF51F9DA45347EFAE00
                  D034413E6B9C6436F6B4D730512EB1CA1D37134DC32C356E4681C82833B21033
                  9D344C3EE735C4BF1FD0679C20AF25B1C6AD8DF6B9755188ECF2710BA151B875
                  D184AC57B7788CF16F0106B1209DE7411D1D20591F2349A71E9EFBC19580A521
                  5317AC173985DA5A66A9F302D23C1468C3820483A0B6E6599278B5CE553DBEA0
                  1F3D8EB87551B899684BCB99190A7622CD43011B52B182CE3940CDB66972B1E3
                  1F3D4E54A6D4B363C1B3B62FA9CB99B2CCB9A8A301D4D1000A9E057536807CFC
                  C4E1134746F1FF069EA81A8587491C0FFA41C126CCBBC8EF4C92DF3149A7EBEE
                  D2C963932FBB98C7DFAC19C32870E9C6E0754E6C395D7797FC8E49F23B93444D
                  8E696A39364D678C534FD5978FA96E260A917B02405681C846515F3E8A06C317
                  CFB51C9BA626C73451ABEB6B0A7866C96F9B2CAFAF1E7D42B087FA62F5187CD6
                  3BF5ED9EAFA9D5F535D12926497E5DF21951ABCCB8B9E813078A9C0297761C3E
                  ED9D5F9C6293440DA6BBE415121A1733FEE471F3F7C8717884B8B1C1748FC8A3
                  8F93C4A9B7325FCD38A4AA7B902BEE433AF2D5CAA6621A52E597105935E30C25
                  B1EAA447883F4DB210FBD9E2232B06B771002E4B082EF38D958DA517AEA361C8
                  820A898DA557918D42E6635924716AC16238A9621AB5C14A580736C07EFD9DCC
                  B9967A7574EF404D4F1E6A7AF2E1E8DEB9E0B34C7144DE84E552016A2A06E165
                  A620716AA62A5A48E6627519E75A2606A9E23E6ABACA60195A0D5BCFD6F484B6
                  C0DEB31535E1B7610B6D85B9FB0D5842B9B08652C7ECBD7F812DB425E3B9F69B
                  1B61BD9A075DE175380FC570D29878041985C4A91D240BB1606AA5B25805AB60
                  EDFF03ECD776A4C5717D276A7B76417F690B741FAF47D5B92C547FBC06FCE5B5
                  E02FBC0967773E1CD7F3329F1BD90CDBA77BC01787A02B1884B3746421925120
                  F16A98445619F9E6265E1C5813DA05FDA5CD286B5D0DFD67AF61DBE13FE3D5EC
                  1D38E45D8FCACE55E02F6E424D68D7370285033DE0DE1B866E6F2F9CA5B7E033
                  2CA8648C4456995A0ED0D99D0FDBE7DB7124B016C2D5D5C8ABDA8C5737EEC01F
                  776FC31EE326F057B250DE9605F3BFB665443E0CE4DFBF055E13FE1F72043E43
                  7C0E796FD9C0B9EA95B765C1D49D859C826DF8F5CBEFA2E44C368CD75E4775E7
                  7A94B5AC0177E18D250185A2C83CB2A66CBE92F7486494B1E502F98B9B50D6BA
                  06DAAE75D09D5F87EDA59BF1D2EFF2B1BD7433F84FD7A2A27D0DD8F3B9A809E5
                  2F09F830D2597A0B270D8904C9FAD8C7CB192435A15D305DD98AF2B62C68BBD6
                  E11FFE6C5476ACC7861D6FE1E5D7F2C15D5E8B8AC06A089FFC69C9154C43968F
                  F493CCC7EA17BFCD2C0E745CDF09C7F59DA80A6E007B7915F69836E137AFBC8B
                  DFBE9E87FDC773A0BBB81A951DEB60FBFC1D38BBF3BE1350288A802B0C832F0A
                  7791C82A1F7C9F516CF9F7DBA8685F07E6D22A141872B09BCF85E1DAEF51DEF6
                  3A8E5EDD86DA9E6F1FC599808226027D715F2DC942FC85E502E790B6CF77803D
                  FF06F457D7437F753D98AE5C583FDBBEE82D6629405E1386A9F8662EC97C8C44
                  4E89650496CDA0F6F241D8132FC1399C9D9EA16C3886B2513B928363B773E11C
                  DE08E770368E8DE4A27624078EA1D477329D5BA3AE82FDC616080742E0FF3692
                  5EC1A2C8FDA3076FFE904E1B93E411129569CB2D4685473B0587781C86A603B0
                  346A613953FD8DB1366A616DD4C2D258FDEDDF6D2D85D9CF4328EE05BF6F30AD
                  7AFAC2F0714B7184C8A79D204FD5C44FDCDAF4A5BEAC57E02ABB0761771CDABC
                  31F005CACA65B70ABE200AA1A80FC2FEF03C4EBF3F0276EF0DE80F0EFECA5432
                  4474B66E9A5AEAA7C96B4E58EAABC7D2963C3E7D02121B83A0E907B3370C41D3
                  BF82E94B6B5AE6BD1E58CB6E357BA529925C5344AD27662920CED259EB573F72
                  1D199F7EF496E3661478F5A9A947BF3F02665F6FA6FEB222E1351168FF1A828D
                  1D7CDE75E2367DE41821F2EA13E4151274D27C8764733CA79E19CBF8D8F900D9
                  07DDDE5EE8F7AF3C90D97703B5FF1CD5C8550A1D2F19A5E3A5A344ADFE196AF5
                  CF505BC30CB5FB67C9CB27CCF555631957D88F13C9BCD70BBE381CF07B92D4DC
                  304567BCA950D08EF97439414DE6AFE844F568EB62DB673E431C22FB00B92240
                  4D04EC9E708F851978CA79A29FAC8E3EB23A53A14013E6D3D10CEA68FB0F359F
                  B94F12AFB6BBAAC71745CE55F2FBF64976DF0DF085E16EBFF8C58FDB03F7A9B9
                  799A9A9A1E8402E7F0209DA0CE8BA060676AFB4DD42A75EE0C4F5C2B327034A9
                  F956280A371A8AFB7ED0DA324B9F5C01757682CE9D7B90C5373079957C8609F2
                  9AE36FB9B5D1FB99563CCBED93DC0761701F84613E3458682B1D24C3FE08359E
                  9AA6AE0BDF7587D534413E6B826436F6BCD790B067DACCFC2E485E13065F1881
                  F1EF035E5349FF8BF69261B2970E91BE28BC4CA07182BCE638C97C6A13DDC326
                  7E29F32A2771EAF8A3CD3D3770987D29245F184E8134F3B34452AF095B4C8706
                  5F716847C9523940D60F87C85632B802C047FF86E0D56724565923B18A20724A
                  ABC82A3D3E435C1159256928EE4B7285E198E1407FD850DC173416F7998D0707
                  724C07069E351545C8FCE130D9AB6F93B9BC7F49C0FF0E00CCE2D35F93C0EA45
                  0000000049454E44AE426082}
                Transparent = True
                TabOrder = 3
                Visible = False
                OnClick = FecharCupom1Click
                Appearance.BorderColor = 14727579
                Appearance.BorderColorHot = 10079963
                Appearance.BorderColorDown = 4548219
                Appearance.Color = 15653832
                Appearance.ColorTo = 16178633
                Appearance.ColorChecked = 7915518
                Appearance.ColorCheckedTo = 11918331
                Appearance.ColorDisabled = 15921906
                Appearance.ColorDisabledTo = 15921906
                Appearance.ColorDown = 7778289
                Appearance.ColorDownTo = 4296947
                Appearance.ColorHot = 15465983
                Appearance.ColorHotTo = 11332863
                Appearance.ColorMirror = 15586496
                Appearance.ColorMirrorTo = 16245200
                Appearance.ColorMirrorHot = 5888767
                Appearance.ColorMirrorHotTo = 10807807
                Appearance.ColorMirrorDown = 946929
                Appearance.ColorMirrorDownTo = 5021693
                Appearance.ColorMirrorChecked = 10480637
                Appearance.ColorMirrorCheckedTo = 5682430
                Appearance.ColorMirrorDisabled = 11974326
                Appearance.ColorMirrorDisabledTo = 15921906
                Appearance.GradientHot = ggVertical
                Appearance.GradientMirrorHot = ggVertical
                Appearance.GradientDown = ggVertical
                Appearance.GradientMirrorDown = ggVertical
                Appearance.GradientChecked = ggVertical
              end
              object lbCupom: TcxLabel
                Left = 6
                Top = 3
                Caption = 'Ultimo Cupom:'
                ParentFont = False
                Style.Font.Charset = ANSI_CHARSET
                Style.Font.Color = clWhite
                Style.Font.Height = -11
                Style.Font.Name = 'Tahoma'
                Style.Font.Style = []
                Style.IsFontAssigned = True
              end
            end
          end
        end
        object v_borda_inferior: TRzPanel
          Left = 0
          Top = 785
          Width = 1553
          Height = 24
          Align = alBottom
          BorderOuter = fsNone
          Color = 3881787
          TabOrder = 1
          Transparent = True
        end
        object v_borda_esquerda: TRzPanel
          Left = 0
          Top = 41
          Width = 29
          Height = 744
          Align = alLeft
          BorderOuter = fsNone
          Color = 3881787
          TabOrder = 2
          Transparent = True
        end
        object v_borda_direita: TRzPanel
          Left = 1527
          Top = 41
          Width = 26
          Height = 744
          Align = alRight
          BorderOuter = fsNone
          Color = 3881787
          TabOrder = 3
          Transparent = True
        end
        object v_borda_superior: TRzPanel
          Left = 0
          Top = 0
          Width = 1553
          Height = 41
          Align = alTop
          BorderOuter = fsNone
          Color = 3881787
          TabOrder = 4
          Transparent = True
          object AdvMetroButton1: TAdvMetroButton
            Left = 1517
            Top = 0
            Width = 36
            Height = 41
            Align = alRight
            Appearance.PictureColor = clWhite
            Caption = ''
            Picture.Data = {
              89504E470D0A1A0A0000000D494844520000002A0000002A0806000000C5C3C9
              5B0000001974455874536F6674776172650041646F626520496D616765526561
              647971C9653C0000032269545874584D4C3A636F6D2E61646F62652E786D7000
              000000003C3F787061636B657420626567696E3D22EFBBBF222069643D225735
              4D304D7043656869487A7265537A4E54637A6B633964223F3E203C783A786D70
              6D65746120786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A
              786D70746B3D2241646F626520584D5020436F726520352E332D633031312036
              362E3134353636312C20323031322F30322F30362D31343A35363A3237202020
              2020202020223E203C7264663A52444620786D6C6E733A7264663D2268747470
              3A2F2F7777772E77332E6F72672F313939392F30322F32322D7264662D73796E
              7461782D6E7323223E203C7264663A4465736372697074696F6E207264663A61
              626F75743D222220786D6C6E733A786D703D22687474703A2F2F6E732E61646F
              62652E636F6D2F7861702F312E302F2220786D6C6E733A786D704D4D3D226874
              74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F6D6D2F222078
              6D6C6E733A73745265663D22687474703A2F2F6E732E61646F62652E636F6D2F
              7861702F312E302F73547970652F5265736F75726365526566232220786D703A
              43726561746F72546F6F6C3D2241646F62652050686F746F73686F7020435336
              202857696E646F7773292220786D704D4D3A496E7374616E636549443D22786D
              702E6969643A3445414132443945303444303131453538344545414445343436
              3535353531312220786D704D4D3A446F63756D656E7449443D22786D702E6469
              643A344541413244394630344430313145353834454541444534343635353535
              3131223E203C786D704D4D3A4465726976656446726F6D2073745265663A696E
              7374616E636549443D22786D702E6969643A3445414132443943303444303131
              453538344545414445343436353535353131222073745265663A646F63756D65
              6E7449443D22786D702E6469643A344541413244394430344430313145353834
              4545414445343436353535353131222F3E203C2F7264663A4465736372697074
              696F6E3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78
              7061636B657420656E643D2272223F3E4B533802000004B14944415478DAD499
              7968554714C65F1215778D2651A352637109C605036EE082A2455C6A5D506825
              54105C402B05A960ABE03F820A2AA2E082FE218A2888FB52975A49085151DBB8
              15375C316AAAC625C6C4EB77E01B3D0EBEF7665EDE7BE0811FF7BE3B73CF7C77
              EEDC9973E6A5044110FA1AAC4E2DEF6F023A806EA02B680B9A8306A012FC0F1E
              82CBE02AB8092A40902CA11DC170F03DE80D5A39DCF3045C007BC1617003BC77
              6D30C5F3D54BEF4D063F832EEAFA6B50069E9337A03E7BB729C8028D54FD5B60
              0BD806AEC75BE838301FF457D72E82425002AEF135BF0255A02E878688EC0CF2
              C1401E8D9D03CBC0CEA8BD2B42A3D0002C02CF824F560C6680AE0EF76BBE05D3
              C029E5EB25580A9A45BA379AE3266003A8A15311FB07E8E029D0A61D980FCA94
              E0ED202B16A10DC13AE5A8147C07D26A29D220C36E083867894DF715BA44F564
              09E81E2781369DC009B6F31E2CE74338099DA8C6E4BF0914A9C76E09DB7B0D0A
              5C84760467D5981C9A6091867EE0A11A66B99184A682DF59B986033E25494285
              59E01DDB972150279CD0CEE0162BFE0D32922832C429EA00DB9719A1A7294B55
              53AA9C8FE6EAF3166CE0B2974C93556D13780932C1789066C4196BCAB55BEC3C
              38E2D140373EE0974C0295EE1EBE8E83629E8F0119F6CAD417BC61B72FF6785D
              23C04D5008BA7C6128FDC5E134CEC3E7AF9CAA64BC0EB3C7E84C8A7C0C463A3A
              1C4691C644540ECB72F8DBD87D30D6D1EF60D617FBCD1EA33D787CC000C3C546
              821CF57B3058CDC065357F1BCBE6AB743169FF2ECFF3EC576FBEB683A091E393
              4BAF1D0E3EB76A50C1A3B6427B6E8C804C4B3BD4ECF3598F9AE0F71143351793
              B8723A38A6AEC957DAD87CADB462C6B0571CFD56333BF8A84B0B35816D95E794
              22AFA8001C0D537E1AFC08FEF3F45BC163335B688A09516398FF5E71EE0D5716
              CB7C1C685DA9563A116264EE6312C16F06A3C2940F67DA91E5E9B7318F2F6CA1
              8FD49868E8E84CEAAE053F58BE6A548FA4B17C9D63126874A5F3BCCC167A57AD
              24998E0EE78109D635F970268193D675590E173AFA95D5E81B9EDFB18596F2D8
              0674727458C47957FB98037683B9E08C2A7BCAFA2E26EDB7E3F9257B1E1D08AA
              F46AE0882C8DF7C025D0DF2AEBC55443E2CC9F3C7CFEC279B8DAAC92BA5042BA
              220A952CB18587E301202F4C9964AA833C7C49AEB65F05D0D9B650099A17A814
              765292635183F46039752C35C1B35D294FA50387A2E5DA094096EE9D6CFF29C8
              0F17E1A7F129028658B3922C742AA864FB6B22A52221060EA5AC2CBDDB274922
              25D3BDCD76AFF3438C9A2E17A820BA280E3B23D16803FE54B9FD6C3BA90C77A3
              7C582B7993D8B1048A6D0BF6A8707023A8EFB353D2925B2C3A9E1C106791BDD5
              2E49C098383B964D32D9B4DAA51CDDE36BC9AAA540E984E9E086F22DB34CFB58
              77F38CD3556ACC1AA753406B4F81999C9FF7295F32BBAC0FD79306D78DDCBA0C
              7E6523375705D812809C026719183F61C05BCD6D7709D55A70773A9F39543FEE
              F187B8DBBC826160653CB7C6257F9FCA88BEB5BA5ECE30F131838F777C3811D9
              927533ACFA5B19C75E48C41EBEE9DD5CEEAA08BD540F4532C900FE0187F88743
              6984AC202E42751297CED7DA93E2B3D973F5D8ABD273F799D45D641A5CCEA1E1
              65B5119A54FB20C0004E7B10B44195D3E70000000049454E44AE426082}
            TabOrder = 0
            Version = '1.2.0.0'
            OnClick = AdvMetroButton1Click
          end
          object AdvMetroButton2: TAdvMetroButton
            Left = 1481
            Top = 0
            Width = 36
            Height = 41
            Align = alRight
            Appearance.PictureColor = clWhite
            Caption = ''
            Picture.Data = {
              89504E470D0A1A0A0000000D494844520000002A0000002A0806000000C5C3C9
              5B0000001974455874536F6674776172650041646F626520496D616765526561
              647971C9653C0000032269545874584D4C3A636F6D2E61646F62652E786D7000
              000000003C3F787061636B657420626567696E3D22EFBBBF222069643D225735
              4D304D7043656869487A7265537A4E54637A6B633964223F3E203C783A786D70
              6D65746120786D6C6E733A783D2261646F62653A6E733A6D6574612F2220783A
              786D70746B3D2241646F626520584D5020436F726520352E332D633031312036
              362E3134353636312C20323031322F30322F30362D31343A35363A3237202020
              2020202020223E203C7264663A52444620786D6C6E733A7264663D2268747470
              3A2F2F7777772E77332E6F72672F313939392F30322F32322D7264662D73796E
              7461782D6E7323223E203C7264663A4465736372697074696F6E207264663A61
              626F75743D222220786D6C6E733A786D703D22687474703A2F2F6E732E61646F
              62652E636F6D2F7861702F312E302F2220786D6C6E733A786D704D4D3D226874
              74703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F6D6D2F222078
              6D6C6E733A73745265663D22687474703A2F2F6E732E61646F62652E636F6D2F
              7861702F312E302F73547970652F5265736F75726365526566232220786D703A
              43726561746F72546F6F6C3D2241646F62652050686F746F73686F7020435336
              202857696E646F7773292220786D704D4D3A496E7374616E636549443D22786D
              702E6969643A3442393632423335303444363131453542384232453930303337
              4433314142462220786D704D4D3A446F63756D656E7449443D22786D702E6469
              643A344239363242333630344436313145354238423245393030333744333141
              4246223E203C786D704D4D3A4465726976656446726F6D2073745265663A696E
              7374616E636549443D22786D702E6969643A3442393632423333303444363131
              453542384232453930303337443331414246222073745265663A646F63756D65
              6E7449443D22786D702E6469643A344239363242333430344436313145354238
              4232453930303337443331414246222F3E203C2F7264663A4465736372697074
              696F6E3E203C2F7264663A5244463E203C2F783A786D706D6574613E203C3F78
              7061636B657420656E643D2272223F3E1CA64786000004204944415478DAD499
              5B4855411486CF454D2D33CB2CB58B1456926668450951541241514451102205
              51D47B6FD5432F4105054190453D443DD84B5958748102C3C428EB7445BB9768
              57B3D4BC9CDDBFE0DF364D9C73F6F6EC7DCA051F7BD8CE9EF95D33CCACB58ED7
              300CCF60B0B828BF4F0139603A9806B2C1089004BAC017D00C1E8127E0396807
              46AC844E02A560252802632C7CF311DC03E7C125D004825627F4DA5C7AF1DE3A
              B0114C55DE778056D0463A4122BD3B1C6480A14AFF17E024380D1A9D16BA0AEC
              00F394770DA006D481A75CE61FA01BC4736B88C829A018CCE7D3B43B601FA88C
              E85D111A8124B01B7C357E5B2DD80AA659F85E6532D8046E28637D077B416AB8
              6F230D9C022A401F0715B1BB408E4D813AE3C00ED0AA083E0332062234191C51
              060A80A5C01FA54813D9760BC11D4D6C9A5DA17B144FD681028704EAE482EB9C
              2708F6F39FB024748DB2271FB82852DDBB759CAF03945B113A09D42B7B7291CB
              224DE68266659BE58513EA033BD9B98F1BDE1B23A1C236D0C3F9650BC485123A
              05BC60C79B203D86223D3CA22E727E39110ACDBFF9942355DACB79FBFC0415BC
              F6626972AB1D07DFC168B01AF84D71A60DE7DD2D76175C765080D746DF6BA096
              ED15205D172A77F71CB62FF3EE76CA0C1B62DB18B4C83705205F175AC4404296
              FBB64B4B6B556C3DE30689EE66EB4267F0F99E0186D366274C93F9DFB0FD9747
              27F0F90EB4B8E451AB6265555FA9BA54A166F0DBC250ED5F5A2FB3837E5D6A84
              6F06B6DD9146E9FBD97D058F25518AD9E21F927034CCDFDBF94CD53DEA1DC05E
              72D30C55974F4B273C8CCCFF071BC6E7377DE95B943D91AC08FFCBB064A52E8B
              1407A6B1DDAA7BD43C0EB2797D397DE3D8E92FB7D144B65FEB42037C66825C17
              26F7DAE82BF38F63FBA12E54EEF71E2EFD2C9B57A3D7C252FA6CE4F1721B65C9
              01C34CF50FA18F7975892D03232D0E1AE45EF74510D96771BC641E7D7E56571A
              74A19F4115DB927B2FB6B1FC3D61BC15E4016ED5168012B62FF407475AE09AAF
              A403D591726D17180A2A39FF27501C2AC2F7B318603025D81663A165A08BF31F
              0E978A78985405D859BC3B27462225D37DC9791BC14C2BE9B2A4AB9DFCE89603
              95914864822B4A6EBF5D4F2A437D28D9E8417E2476D545B1D9E09C522D390612
              ED544A46B1C4625A0D28715864915225319881660DA4482645ABB3CA406FB92C
              19510A14276C064DCAD872CA8C1F6835CF1CF490B267CD41D783B136058E066B
              419532969C2E474379D2C46A215742BF0D2CE4E62901B6A4B53778A33D630AD1
              CE033E8EA1DA4866B8C53CCCE7B2C6EF61B5F900ABCF5D4E96C6E5478532500E
              C66AB79A84891FC027DE54F114398A7DD3B5FEA7C009D6F51DAFE19BDECD6355
              4598A978289C49F5E53EA8E60F0E01BEF3B825B43F7E66702BCB5A48F159F45C
              02BDFA9959ED6306174FF9AED7EE64DEC1F283D82F0106008690EEE7FB647FE2
              0000000049454E44AE426082}
            TabOrder = 1
            Version = '1.2.0.0'
            OnClick = AdvMetroButton2Click
          end
        end
      end
    end
  end
  object Teclado: TAdvSmoothTouchKeyBoard
    Left = 361
    Top = 0
    Width = 243
    Height = 121
    Text = ''
    AutoCompletion.Font.Charset = DEFAULT_CHARSET
    AutoCompletion.Font.Color = clWhite
    AutoCompletion.Font.Height = -19
    AutoCompletion.Font.Name = 'Tahoma'
    AutoCompletion.Font.Style = []
    AutoCompletion.Color = clBlack
    Fill.ColorMirror = clNone
    Fill.ColorMirrorTo = clNone
    Fill.GradientType = gtVertical
    Fill.GradientMirrorType = gtSolid
    Fill.BorderColor = clNone
    Fill.Rounding = 0
    Fill.ShadowOffset = 0
    Fill.Glow = gmNone
    KeyboardType = ktQWERTY
    Keys = <
      item
        KeyCombination = True
        ShiftKeyCombination = True
        Caption = '`'
        ShiftCaption = '~'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 0
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '1'
        ShiftCaption = '!'
        AltGrCaption = '|'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 17
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '2'
        ShiftCaption = '@'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 33
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '3'
        ShiftCaption = '#'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 49
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '4'
        ShiftCaption = '$'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 65
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '5'
        ShiftCaption = '%'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 81
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        ShiftKeyCombination = True
        Caption = '6'
        ShiftCaption = '^'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 97
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '7'
        ShiftCaption = '&'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 113
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '8'
        ShiftCaption = '*'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 130
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '9'
        ShiftCaption = '('
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 146
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '0'
        ShiftCaption = ')'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 162
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '-'
        ShiftCaption = '_'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 178
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '='
        ShiftCaption = '+'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 194
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '\'
        ShiftCaption = '|'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 210
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Back'
        KeyValue = 8
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skBackSpace
        Color = 10526880
        X = 226
        Y = 1
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '->'
        ShiftCaption = '<-'
        KeyValue = 9
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skTab
        Color = 10526880
        X = 0
        Y = 25
        Height = 24
        Width = 24
        SubKeys = <>
      end
      item
        Caption = 'Q'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 25
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'W'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 41
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'E'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 57
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #203
          end
          item
            Caption = #202
          end
          item
            Caption = #201
          end
          item
            Caption = #200
          end
          item
            Caption = #235
          end
          item
            Caption = #234
          end
          item
            Caption = #233
          end
          item
            Caption = #232
          end>
      end
      item
        Caption = 'R'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 73
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'T'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 89
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Y'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 105
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #255
          end
          item
            Caption = #221
          end
          item
            Caption = #253
          end>
      end
      item
        Caption = 'U'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 122
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #220
          end
          item
            Caption = #219
          end
          item
            Caption = #218
          end
          item
            Caption = #217
          end
          item
            Caption = #252
          end
          item
            Caption = #251
          end
          item
            Caption = #250
          end
          item
            Caption = #249
          end>
      end
      item
        Caption = 'I'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 138
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #207
          end
          item
            Caption = #206
          end
          item
            Caption = #205
          end
          item
            Caption = #204
          end
          item
            Caption = #239
          end
          item
            Caption = #238
          end
          item
            Caption = #237
          end
          item
            Caption = #236
          end>
      end
      item
        Caption = 'O'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 154
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #214
          end
          item
            Caption = #212
          end
          item
            Caption = #211
          end
          item
            Caption = #210
          end
          item
            Caption = #246
          end
          item
            Caption = #244
          end
          item
            Caption = #243
          end
          item
            Caption = #242
          end>
      end
      item
        Caption = 'P'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 170
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '['
        ShiftCaption = '{'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 186
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = ']'
        ShiftCaption = '}'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 202
        Y = 25
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Enter'
        KeyValue = 13
        ShiftKeyValue = 0
        AltGrKeyValue = 0
        SpecialKey = skReturn
        Color = 10526880
        X = 222
        Y = 25
        Height = 48
        Width = 20
        SubKeys = <>
      end
      item
        Caption = 'Caps Lock'
        KeyValue = 20
        ShiftKeyValue = 0
        AltGrKeyValue = 0
        SpecialKey = skCaps
        Color = 10526880
        X = 0
        Y = 49
        Height = 24
        Width = 28
        SubKeys = <>
      end
      item
        Caption = 'A'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 29
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #196
          end
          item
            Caption = #194
          end
          item
            Caption = #193
          end
          item
            Caption = #192
          end
          item
            Caption = #228
          end
          item
            Caption = #226
          end
          item
            Caption = #225
          end
          item
            Caption = #224
          end>
      end
      item
        Caption = 'S'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 45
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'D'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 61
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'F'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 77
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'G'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 93
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'H'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 109
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'J'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 126
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'K'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 142
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'L'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 158
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = ';'
        ShiftCaption = ':'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 174
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = ','
        ShiftCaption = '"'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 190
        Y = 49
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Shift'
        KeyValue = 160
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skShift
        Color = 10526880
        X = 0
        Y = 72
        Height = 24
        Width = 32
        SubKeys = <>
      end
      item
        Caption = 'Z'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 33
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'X'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 49
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'C'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 65
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #231
          end>
      end
      item
        Caption = 'V'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 81
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'B'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 97
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'N'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 113
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <
          item
            Caption = #209
          end
          item
            Caption = #241
          end>
      end
      item
        Caption = 'M'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 130
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = ','
        ShiftCaption = '<'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 146
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '.'
        ShiftCaption = '>'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 162
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '/'
        ShiftCaption = '?'
        KeyValue = -1
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skNone
        X = 178
        Y = 72
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Shift'
        KeyValue = 161
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skShift
        Color = 10526880
        X = 194
        Y = 72
        Height = 24
        Width = 44
        SubKeys = <>
      end
      item
        Caption = 'Ctrl'
        KeyValue = 162
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skCtrl
        Color = 10526880
        X = 0
        Y = 96
        Height = 24
        Width = 24
        SubKeys = <>
      end
      item
        Caption = 'Win'
        KeyValue = 91
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skWin
        Color = 10526880
        X = 25
        Y = 96
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Alt'
        KeyValue = 18
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skAlt
        Color = 10526880
        X = 41
        Y = 96
        Height = 24
        Width = 24
        SubKeys = <>
      end
      item
        KeyValue = 32
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skSpaceBar
        X = 65
        Y = 96
        Height = 24
        Width = 97
        SubKeys = <>
      end
      item
        Caption = 'Alt Gr'
        KeyValue = 0
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skAltGr
        Color = 10526880
        X = 162
        Y = 96
        Height = 24
        Width = 24
        SubKeys = <>
      end
      item
        Caption = 'Win'
        KeyValue = 92
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skWin
        Color = 10526880
        X = 186
        Y = 96
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = 'Menu'
        KeyValue = 93
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skApp
        Color = 10526880
        X = 202
        Y = 96
        Height = 24
        Width = 16
        SubKeys = <>
      end
      item
        Caption = '->'
        KeyValue = 39
        ShiftKeyValue = -1
        AltGrKeyValue = -1
        SpecialKey = skRight
        Color = 10526880
        X = 218
        Y = 96
        Height = 24
        Width = 24
        SubKeys = <>
      end>
    SmallFont.Charset = DEFAULT_CHARSET
    SmallFont.Color = clWindowText
    SmallFont.Height = -9
    SmallFont.Name = 'Tahoma'
    SmallFont.Style = []
    Version = '1.8.5.1'
    Visible = False
  end
  object wspnlAlerta: TW7Panel
    Left = 840
    Top = 796
    Width = 713
    Height = 468
    Cursor = crHandPoint
    Version = '1.0.2.0'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = []
    SolidColor = clBlack
    OuterBorderColor = 1182728
    InnerBorderColor = 5721929
    InnerBorders = []
    OuterBorders = []
    Rounded = False
    Opacity = 200
    Style = wpsCustom
    ColorA = 20817
    ColorB = 2102542
    ColorC = 2960684
    ColorD = 197378
    DoubleBuffered = False
    DragMode = dmAutomatic
    ParentDoubleBuffered = False
    ParentFont = False
    TabOrder = 4
    object JvPanel1: TJvPanel
      Left = 0
      Top = 0
      Width = 713
      Height = 468
      Transparent = True
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 0
      OnClick = AlertaFecharClick
      object lblInstrucao: TLabel
        Left = 137
        Top = 385
        Width = 525
        Height = 33
        Cursor = crHandPoint
        Alignment = taCenter
        AutoSize = False
        Caption = 
          'Clique em qualquer parte da tela ou pressione Enter para continu' +
          'ar'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = 6488063
        Font.Height = -16
        Font.Name = 'Arial'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Image1: TImage
        Left = 286
        Top = 61
        Width = 223
        Height = 190
        Cursor = crHandPoint
        Picture.Data = {
          0B546478504E47496D61676589504E470D0A1A0A0000000D49484452000005DC
          000005DC08030000000FEBBC560000001974455874536F667477617265004164
          6F626520496D616765526561647971C9653C0000036869545874584D4C3A636F
          6D2E61646F62652E786D7000000000003C3F787061636B657420626567696E3D
          22EFBBBF222069643D2257354D304D7043656869487A7265537A4E54637A6B63
          3964223F3E203C783A786D706D65746120786D6C6E733A783D2261646F62653A
          6E733A6D6574612F2220783A786D70746B3D2241646F626520584D5020436F72
          6520352E332D633031312036362E3134353636312C20323031322F30322F3036
          2D31343A35363A32372020202020202020223E203C7264663A52444620786D6C
          6E733A7264663D22687474703A2F2F7777772E77332E6F72672F313939392F30
          322F32322D7264662D73796E7461782D6E7323223E203C7264663A4465736372
          697074696F6E207264663A61626F75743D222220786D6C6E733A786D704D4D3D
          22687474703A2F2F6E732E61646F62652E636F6D2F7861702F312E302F6D6D2F
          2220786D6C6E733A73745265663D22687474703A2F2F6E732E61646F62652E63
          6F6D2F7861702F312E302F73547970652F5265736F7572636552656623222078
          6D6C6E733A786D703D22687474703A2F2F6E732E61646F62652E636F6D2F7861
          702F312E302F2220786D704D4D3A4F726967696E616C446F63756D656E744944
          3D22786D702E6469643A46373746313137343037323036383131383344314537
          323342394535464133302220786D704D4D3A446F63756D656E7449443D22786D
          702E6469643A3542393938393141303538333131453739433144444242343943
          3132383442452220786D704D4D3A496E7374616E636549443D22786D702E6969
          643A343838314339453630353833313145373943314444424234394331323834
          42452220786D703A43726561746F72546F6F6C3D2241646F62652050686F746F
          73686F702043533620284D6163696E746F736829223E203C786D704D4D3A4465
          726976656446726F6D2073745265663A696E7374616E636549443D22786D702E
          6969643A46413945363832353045323036383131383344314537323342394535
          46413330222073745265663A646F63756D656E7449443D22786D702E6469643A
          4637374631313734303732303638313138334431453732334239453546413330
          222F3E203C2F7264663A4465736372697074696F6E3E203C2F7264663A524446
          3E203C2F783A786D706D6574613E203C3F787061636B657420656E643D227222
          3F3EDBB5197900000300504C5445E5A200AE5225AC4A01FFD100F5E9E4F9F3D4
          C58359BC6400DB9400F3E4C4B35400F7EDCDFFBD00DCA400C37100CB7B00A035
          00D19C83CC936AF4D200AA4A1BEEDBD2B5623AD58A00DD9A00C48262C26C00FF
          D500FFD900FFDD00B96B42CC8100D3A189A94400EDC300D4A37CD28400FFC200
          FFC500B35D33D19D75B15A2DB559009F3400BD724BF9BA00FFE400A64212C17C
          53FFCC00E6C9BCC27C5CEBD5CAF8F1EDA13601FFB800FDFAF9A64109D29200EB
          D4B2F0CB00EDB200E19D00F6EDE8DAB19DDCB5A1E3B200C98C63E9D1C5CB9175
          F6C000C57A01BE6900EEDBB9A53D00B56134E0AC00E0BDACFAF4F1D6A993F9C5
          00C88B6CBE7553FEFCDEDBB28CE9BD00CD9579B9693CF5B500F1E1D9E9A900E2
          C1B1F2B900F4BD00BC7144A43E05F3E5DDDEB993A43D0DFFE200B95D00DEB8A6
          E7CDC0FEFDFDFFCA00FCF8F6DFBAA9FADC00AA4B16F1B100FFC000E2C19DFFE0
          00D69900FFC800D5A580FBF6F4D8AC96EDAD00A23908FCF9DAE4C6A2DCB590F8
          D900CE8B00EFB500A23A03D9AE99A74515BA6D49FCBB00E6CAA6A33C04D6A982
          E1BE99AB4D20C0794EC87500EAAD00FCCA00A64107FDFCFBEDD8CFCA8D70A740
          00E4C5B6AD511DFAC800FDFADCAF562AE7A600CF997EA53F06F0DED6A13802A5
          3E03F3E6DFA13604A03501D88F00A23800E9D0AEFCE000E5C7B9A03602AF5625
          FDE200E7CCA9AE5001E8CDAAD8AC85E6B800A74408D5A690A43D0BC68567A33C
          0AEDD8B6A8450BA33A00F1E1C0C7886ACF9970A5400FEFDDBCA23906A84718A2
          3600B76640D9AE88F9F3F0AF4E00F7C200A43B00D6A88DB86000B15926EBAA00
          E1BFAFC7885FF7D500B45F2FA9480FE3C4B2FBCD00A7440CF7EFEBF0E0D7E8CE
          BEA64204A43D05A43C04B76639A64207A33B04A33A06A23903A23803FEBB00A7
          4307A64004FFB700FFCF00FFBA00FFBB00FCC900FFFEE0FEBA00A13806FDCB00
          F7B800E3C4B5F2E3DCFBF7D8A13705EFAF00F4E8C7BF774BE2C2B3FDCF00F7D7
          00F4E7E1A13700E5C8B7A33D03A33903A84505A53D04B76636A74305FFFFE1FF
          E600FFB6009F3300FFCE00FFFFFFD253468A0000010074524E53FFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
          FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0053F707250000
          8F814944415478DAECBD7B6055D599BF0F110886210977882D21982926E12A10
          10913B120151512EA282229741648411452B0A542B460BC60B20540362A15150
          B04168918A14117166B045EB20161851A945E8289D5F7F4EE4A4DF134405DD3B
          C979D75A7BAFBDF7F3FC3D534FD6E571B93E6BBF6F8D7F000040E8A8C1100000
          2077000040EE000080DC010000B903000072070040EE000080DC010000B90300
          0072070000E40E0080DC010000B903000072070000E40E0000C81D0000903B00
          0072070000E40E0000C81D0000903B000020770000E40E0000C81D0000903B00
          002077000040EE000080DC0100903B00002077000040EE000080DC010000B903
          002077000040EE000080DC010000B903000072070000E40E0080DC010000B903
          000072070000E40E0000C81D0000B903000072070000E40E0000C81D0000903B
          000020770000E40E0000C81D0000903B00002077000040EE0000C81D0000903B
          00002077000040EE000080DC010000B903002077000040EE000080DC010000B9
          03000072070040EE000080DC010000B903000072070000E40E0000C81D0000B9
          03000072070000E40E0000C81D0000903B000072070000E40E0000C81D000090
          3B00002077000040EE0000C81D0000903B00002077000040EE000080DC010090
          3B00002077000040EE000080DC010000B903000072070040EE000080DC010000
          B903000072070000E40E0080DC010000B903000072070000E40E0000C81D0000
          903B000072070000E40E0000C81D0000903B000020770000E40E0000C81D0000
          903B00002077000040EE000080DC0100903B00002077000040EE000080DC0100
          00B903002077000040EE000080DC010000B903000072070000E40E0080DC0100
          00B903000072070000E40E0000C81D0000B903000072070000E40E0000C81D00
          00903B000020770000E40E0000C81D0000903B00002077000040EE0000C81D00
          00903B00002077000040EE000080DC010000B903002077000040EE000080DC01
          0000B903000072070040EE000080DC010000B903000072070000E40E0000C81D
          0000B903000072070000E40E0000C81D0000903B000072070000E40E0000C81D
          0000903B00002077000040EE0000C81D0000903B00002077000040EE000080DC
          0100903B40889872F5EFFBB66A58A7EDA8D71EBDF9FD3FFEF18F3FFBD9CF6EBE
          79D59096C7EA0CBCB5EFDFBAF4668400B903048AEB1FB9ACFFD8472B84FE2D3F
          3B8D0B2E7822CEAF76B5BDE277D7BCD59DF102E40E60390B9AF46D38F6E6F72B
          7055FBCF7EF6C457FC29CE134F5CF0DA034B2EBD7A326307C81DC04ABAFEADC1
          D837DF3FC51FDDDDFE8413BF1A7545DFA5181E903B80550CFBF7815FFCE1FD6F
          A9FAD8EEC49BF346FE7B570613903B800D2C18B064D4AFDEFF3F37B5FFB1BA6A
          3FC5AE2B6EAAC5A0027207F095AE13EB3CFAC7F7CF4049EDA704FFF0BF0F636C
          01B903F843ADCBC6FEEAFDEF22BB91398371712E18D5602177F080DC01BCE6FA
          DF8D5AFB7EE56A971DDB4FBAFD248FAE7B9D033C207700EFE87DDBA8F79DD0A9
          F60AF65CD0B64F6F861B903B8007CCBCE9D8AF1254FBCF446A1FB7A782712D7F
          F739830EC81DC02C5DFADFFC7ED56AD7726CFF9651B7717E07E40E608CDD97ED
          5AFBFEFB9E1CDBF79CC1DBE3C6F67985F107E40E60800175DE7CFF7D63C7F671
          EE6E7FFB241774FE1BEF6700B903E865669F2FDE37A8F64A8FEDDF307FE45266
          02903B80369EDC78F3FBD573FBCFCC1CDB4F71DF7DA3FACE643A00B903E8A04B
          9D5FAD5DEBE78DCCB76E8F73F3408EEF80DC0154997CCDD8B5EF4B8EED7A6F64
          BE56FB495A5EBA808901E40E20A7FBC45D32B51B3AB67FCDA3AD76333980DC01
          644C99F35AB5D5EED9B13DCEBB712E58D7840902E40E9038337FB76AAD97C7F6
          EAAAFDA4DB2B683B894902E40E90E0A9FD366FD55EED1B99774F630897EF80DC
          0112A0FB9C55FF67E58DCC196E7FF79E7B56CDA1C536207780EA31B9EF6B6B6D
          39B657AAF60A1E9D33851903E40E50357F1B9280DAFFE8D38DCCD76E8F33FF77
          E81D903B40157419BBB652B75B7323F32DA5F36FE37206903B40257C5EE7FDB5
          F607A967A8FD9E7BF6EEDD5BB32FD12A20770017662E79D3F88DCC133A6F644E
          A9BD829C4B993F40EE000E4C9EB86AED5ABB9FB63B1FDB4FF14E17E610903BC0
          77E9326A6D106E64DE75517B9CF73A7EC234027207389DAE0DD7FA7923233BB6
          EF3DD3EDEFBDB7FEE1614C252077806FE8FB68A06F64BE527B058DFBD0AF0990
          3BC0572C1DFB7EB09EB63B1EDB4FD194AB7740EE0071BA377833704FDB9D8FED
          5FB1775D572615903B449E2E43FE1EC8A7ED2E6A7F6FD9B2659BFB32AD80DC21
          DACC1CB836B04FDB9DDCBEEC2B5AF26E06903B4499DFBF260E526DBC91F9DAED
          CB96BDD0804F5601B94354993232F04FDB9DD55E411B5A3501728768D26597A5
          4FDB956E64BE617D035E45027287E8B1A0C1AFC2F1B4DDC5EDCB96E50FE1F00E
          C81DA2C6932DC3F3B4DD51ED270FEF9731D180DC2152DC7473E08A0DDC93A8DA
          E3BC37AF37730DC81D22C3B09DC1E9A327BA91F9F66A66F3234C37207788084D
          7685FF46E62BB557F0306D9A00B94324E8FB66289FB63BAB3D4ED35ACC392077
          083D33FB87F669BBB3DBF3F3E73EC3B40372879073FD17EF87F769BBA3DAE3BC
          B7848907E40EA1E66F8F86FA69BBB3DB8B8B8BDBD2C503903B8498DFFD2AC837
          327BA56A8F93733DB30FC81D42CACC3A41EFA397B8DA4FB9BDB878EE24160020
          770825B5BE581B89A7ED4E6A8FF39B3E2C0140EE104216AE8ACCD37627B7C721
          5605E40EE1E39137D786A4D8C07B32B5C76941917740EE10326EFB5588FAE825
          7A23F30DF366B21200B9439818F97FE1EAA32738B69F64C82BAC0540EE101ABA
          EFF4F969FBCF7D0C52CF246737CB01903B848499C7D67A7C23933BA847FD8BAE
          EBF9D07FFCCF59070E9477FAB71FFDEB8913FBBB75DBB7EF273F393C6D4CDD41
          255E3C6D77246F11956600B94338D83DCABB1B9994CCFA231EFB9F03E5DFE19F
          F79F38936EFB7E327A71E6DFBDBC91F9CAED79B19AD81D903B84C2ED43D67A72
          6CDF5E77CD43FF53EE4CA7EFCAFD14077F326D7989573732156AAFA0E6E72C0A
          40EE10786AED32FFB4FD8F83863E76567965FCE8843B07DB2DCEF0E446E614B1
          45DCBB037287E0BBFD8F866F6452873E76A0BC4A0E9EA8946EED6AE4DE63FA46
          E61B1675656100728740F3C92A9337324FFCBDC7C56795578B7F3B5125FB4667
          BE6BF846E66BDA502412903B04FADCFE9AC11B9992E68FFD4F79B5F9D713D5A0
          5BBBE53F377D6CCFCB8BC562F3F85615903B04F94EC6541FBD27FE3EE3F203E5
          89E096A97EDFEFD9668FEDB193B46075007287A0B27B97A91B99B4D68999BDAA
          4CF53B09EBB464733732B15334607D00728760F2CA1033416AEED0FF2997F0DF
          27AACF4F7AFDDCD48DCCD7F465850072872032A5E55A13C7F641171F2897F12F
          2712A1DBB409866E644E31AB0B6B04903B048FC91D0DA8FD6733FEC3D5DD679D
          75CE393F38C939E79C75409CA97ECBFE9FA499B9913945639EBB037287E071C5
          5AED373249F5CF7252FA0F7E307DFA8DEDBF3C93F6D37F78CE77FFAF3B753B91
          203F596EE446E6145B26B34C00B943C068A5FDD85EF2E0E9B23E1097FA0FA77F
          57E9DF35FC0FCF3920CB54BFE1BFEB9A39B69F6420EB04903B048B4B751FDB93
          867EA3E9B8D6ABB2FA197E1766AADF7CDB54F73DC9B1BD6AB5C779869502C81D
          82449737F53E6DFFFBA90B99B3CEF961F5BDFE35379E23CC54BFE6C366DA6F64
          4EB1996B7740EE10206AADFA3F9D37323F9B51F1F6F1AC1FFC3061AF9FE28767
          0933D56FEEDE0B75DFC89CA22D8B05903B0486295F68BD9119F41FE507CE118B
          FDABCB9973C499EA29DAA5683FB69F640ECB05903B0485FE3A8FEDB9ADCB15CD
          7E86DD7F2494FB896E3BB6E9577B2C369ED61D80DC2120F4D5A8F6F7EB6B30FB
          49CE52C8544F25ABD95A6F644E319C0503C81D02419337D7EABB91F9811EB357
          C4AA4A99EAA9BB99AD7A8FED27799D2503C81D02C0B05DDA8EEDC96F7CA99173
          D432D5AFEE669AEB3DB69FFC5095DAEE80DC2100ECD4756C2FE8F9A5566E54CD
          544F05AB795AD5CEA74C80DC21105CA349EDC923BED4CD01D54CF5ABC37B5D7D
          37325F71E46A960D2077B09CDE8F6AB991C9EDF7A57ECE2A57CE544F72384BE3
          B1BD8279AC1B40EE6039C7741CDB9366B437E0F66F2EDD9532D5935FAC16EA3B
          B69FE4AF2C1C40EE603513751CDB07BDF1E59746E5AE94A99E2C06BC58A7DA63
          B1F9DD593A80DCC162763FAAAEF6927E5F1AE22C85DABFDFCB55B3F4DCC89CA2
          156B07903B58CC4EF52075D05DA6DCDE5EADF6EF77BF68CAC8D3746CAF603DDF
          A90272077BF9BDF2B1FDEFF5BF34C6F472C5DABFDF7935B35CD7B1BD82CEAC1E
          40EE602BDD77A91EDB737B9A73FB97E7C8FBA93AB3439BDAE30C60FD0072074B
          F99D6AB181D4A906DDDEFE804A3F55E78BF7525D6A8FC57268B907C81DECA4F7
          CD8A7DF47A7C69921FA8F653752AF39E95AEC9EDB1D86DAC2040EE60253BD58E
          ED6BEB1B757BFBEF76CBFE9106B99FD897AB49EDB15887AE2C2140EE60215D94
          D4FE44D29A2FBD3CB86BC9542B62D50C3D6A8FF3126B08903BD8C7E42F546E64
          9E486A6DD6ED371EF89EDC7564AA274E1C2CD4E4F658F14256112077B08E3EEF
          2B1CDB9F48EA69D6ED673E95D198A9C6CFEE99314D6C6115017207DBE8FAA8C2
          B1FD02E36EFF8183DBF564AA274EECCFD665F789AC2340EE601903FF2056FB13
          E38CBBFDC67247B464AA1AEDDE78260B09903B584593B5E21B99279EF8BB69B7
          7FEFA58CD64C55A3DD47B29200B983558C951FDBC7FDDF882F7DB870D798A9C6
          ED9EA647EEB33E61290172078BB8547A6C1F376EDC13FD7C73BBAE4CF5C4896E
          A97AECDE96B504C81DEC61E62AE18D4CDCEDE3667CE94798AA37538DDB3D598F
          DD9F61350172076B6820BD91899369DAED3F2CAF0C4D99EA89130753B4C8BD26
          6D3B00B9832D5CFF2BF1B17D5C81BF6ED796A99E38B12F4B8BDD69DB01C81D6C
          E101F1B17D5CCA549FDDAE2D533D71E227E93AE43EBE370B0A903B58C1DFC46A
          1F67FC1164956ED797A99E38315ACBD1BD052B0A903BD840F75DC21B9971E3F6
          0CF5DFEDFA32D513276A68B13B6D3B00B9830DFCEE4FC263BBF930B53A6ED798
          A99ED8AFA5CC4C1BDA76007207FFE97DB354EDE37257FAF706D248A67AA2DB04
          1D76EFC3AA02E40EBEB3537623E3C185FB39D574BBC64CF5C4BE520D729FFB0A
          CB0A903BF8CC00E9B17DDC9EFA96B85D67A6AA27547D987505C81DFC65F217C2
          63FBB871A936DCB76BCF544F34D7D1B663292B0B903BF84A1FE1B17DDCB8A4BB
          FC29046938533DD16DBB06BBB764650172073FE97AF31F656A37FE0AF207E509
          A12F533DF1939735D8FD52D6162077F09187653732E3F6EC19E44B770E2F32D5
          13D334C87DFE1416172077F08D261708D5BE2769AA3569AAF64CF5848EE2EE4B
          585D80DCC137C6CA6E64E20CB5274DD59FA9EE1BACA16DC7932C2F40EEE01337
          FD4C766CDFB3C7F04B99C4D254ED99AA96F790C3595F80DCC11FA63C2A3CB6EF
          59DBC8AA34557BA6AAE562E67C56182077F085257F92A97DCF1EC3CD97DA1F90
          C85D67A6BA6F85BADC172D60890172071FF8E43F4537327BF68CCB6D6F599AAA
          3F53BD5DC3D1FD79D6182077F08163C263FB9E7123CCBA7DBACCED5A33D5FD1A
          2A886DEACD2203E40E9EF337D9B13DEEF641F6A5A9FA33D59FD0B603903B0491
          EE397F12A97DCF9EBFDF65619AAA3F535DAE2EF7BC2EAC3340EEE031BFFB93E4
          4626CE1356A6A9FA33D5831A8AFF36659D0172076FA9F5E69F44C7F63D7B2C4D
          53F567AA9F6AB898E9CB4A03E40E9E5247766C8F63699A6A2053CD5597FBE661
          2C3540EEE02103846A7FFB6DC369EA976729C95D6BA6DA4EC3D17D206B0D903B
          78C7E421A21B99B8DBFFEF0D6BD354FD996AB6BADC8FBCC56A03E40E9E314778
          6C7FFB6D8BD35403DFA96AA8EC3E8FD506C81DBCA2EBCD82637B85DADF4E596D
          719AAA3F53DDA1E162E61AD61B2077F0882B44373215ACB1394DD59FA9764BD1
          D0B6A33B0B0E903B78825B8B8E2A6E64DE7EFBBE54BBD3543B33D506AC3840EE
          E0092D85C7F6B7FF60B8D2EF0FCBB56059EDDFF5B5587280DCC1036E921DDBDF
          7EFBDD1EB616953198A9A6ABDBBD336B0E903B9867D8A332B5BFFD76C94ACB9F
          419AC8546B68B89899C4AA03E40EC65922BA91A938B8F733EBF61B75B95DEF77
          AA5BD5E59E33996507C81D0CF3C95AD9B1FDEDB70B8290A6EACF540F6B38BADF
          C6BA03E40E863996E8D3F6AFF953CF40A4A9FA33D54275B977E8CAC203E40E46
          7944742353F10C727630D254FD99AA8EB61D2FB1F200B98349BABF26BA9189BB
          BD646A40D254FD996A2F75B9172F64ED01720783B4121EDBEFBB6F6850D25403
          DFA966D1B603903B584DAD37856ABF2FD9709A7A8E66B96BCD54476BB89899C8
          EA03E40EC6A823BA918933AE7580D254ED99EAFE0C75B9379EC9F203E40E8618
          203CB6DF775F6690D2543B33D591AC3F40EE6086054364C7F6FBEE4B0A569AAA
          3F53ADAB2EF78F3E610502720723CC11AAFDBEFBEA1B4E530F9890BBCE4CF560
          91BADDDBB20201B9830976DF2CBA9189B3BD7DC0D254FD99EA340D17338FB006
          01B98301AE101EDBEFDB13BC34F52BFE4963A69AAC2EF79AB4ED00E40EFA6922
          3CB607324DB533536DC52A04E40EDA692954FBFFFEFDB300A6A9FA33D5E5EA72
          1FDF9B6508C81D343351742353C10CC307F703E6E4AE3553DD46DB0E40EE601D
          152D3A04C7F6FBEEDB931BCC34557FA6FAA9868B9901AC4440EEA09591D53EB6
          9FA1F6FBDEBD6F8459B74F2F378ACE4C35575DEE6D68DB01C81D7472F505A21B
          9977DFFDAF41C169D1613A536DA7E1E83E87B508C81D34D2F609C18D4CDCEDEF
          FEFD8DC0A6A9FA33D56C75B9CF7D85C508C81DB4F1EFB2637B9C00A7A9FA33D5
          7D2BD4ED7E05AB11903BE862CA6B42B5BF9B12E434557FA6BA4343DB8E26AC47
          40EEA08956A21B990AD6043A4DD59EA9764B51B7FB16D6232077D043AD37F788
          8EEDEFBE9B1AEC34D5CE4CF55256242077D0421DE1B1FDDD3F340A6851197399
          6A9A86B61D53589280DC4103BF4FF469FB37F4086A511983996ABABADD97B026
          01B9833A0B76896E64E294AC0C7C9AAA3F535DAC2EF7594FB22A01B98332B789
          6E642AE867B84547B967E8CC54B7AADB7D38AB12903BA8B2FB66D9B1FDDDFB0A
          C290A6EACF540F6BC854CF675D027207455E92A9FDDD777FDE331469AAFE4CB5
          505DEE8B16B03001B983120B9F10DDC8BC7BCF3DB3C391A6EACF5475B4ED789E
          9509C81D9418253BB6DF734F92E134F507E59EA23353ED45DB0E40EEE03313AB
          796CFFAEDAEFB9676868D254ED99EAC12C75BBB7606D02720739C31E7D5B7223
          1327F9CB903C833491A98E56977B5E1756272077103352742313E7BE30A5A95F
          F1CF1ADB7664A8DB7D08AB13903B48B9FA02D9B1FD9E7B32C394A6DA99A9F665
          7D027207216D133DB69F52FB3D49534395A6EACF54EB6A68DB318C050AC81D44
          5C23BA91A9A0BEE134F5801F72D7FA9D6A91BADD07B24201B9838429AB443732
          71B6872D4DD59FA94E5397FB91B758A380DC4140ABFF951DDBEF79B775E8D254
          FD996AB2BADDE7B14601B943E27CFE9F893E6DFF9AB4F0A5A97666AAAFB34A01
          B943C2741E27B99189F3F7BB4298A6EACF5497ABCB7D3E6D3B00B943A2FC5E74
          2353712933C3F0C1FD807F72D7FA9DEA3675BB37609D027287C458B04B766CDF
          BB37379C69AAFE4CF55375B9AFFF9C950AC81D12E23291DAE36EBF678459B74F
          2FF7159D99EA0475BB7764A502728744E8FDA6E446666F9C41E169D1613A536D
          A721539DC45A05E40E09F092ECD8BE77EF1FDE086D9AAA3F53CD56977B0E6D3B
          00B943F5E9B247766CDFBB37CC69AAF64C75DF0A75BB5FC66A05E40ED5E61DA1
          DAF7A6B40F719AAA3F53DDA12EF70E5D59AE80DCA19AF415DDC854B026D469AA
          F64CB55B8ABADDD7B15E01B943F518D658766CDFBB3735DC69AA95996AF14256
          2C2077A81603FFFF049FB67FCD9F1A85B4A88CB94C354DDDEE4D59B180DCA13A
          BCF584E846264E8FB016953198A9A6ABDB7D226B16903B5483B1A21B9938252B
          439FA6EACF5417ABCB7DF34C162D2077A892D785C7F6BD7BFB196ED1516E0D3A
          BF53DDAA6EF791AC5A40EE50155356BDFBB648EDEF1544214DD59FA91E5697FB
          479FB06E01B943153410DDC8C4DDFEF39E914853F567AA85B4ED00E40EC6F9FC
          02D9B1FDBDF76647244DD59EA9EA68DBF1082B17903B544A47D9B1FDBDBD4986
          D3D41F58E576AD996A2F75B9D7ECCED205E40E953049A6F638432393A6EAFF4E
          354BDDEEAD58BB80DCC19DC939A21B9938C95F46E419A4894C75B486B61DBD59
          BD80DCC195CB84C7F6F7DE8D529AAA3D53DD9FA16EF7CEAC5E40EEE0C6EE3765
          C7F6F7F666462A4DB532531DC0FA05E40E2EAC93A9FDBDF792A6462A4DD59FA9
          3657977B9BC92C6040EEE04897FF95DCC85450DF709A7AC046B96BCD548BD4ED
          3E87150CC81D9C98FC8EECD8FEDE7BDBA396A6EACF54A7D1B603903B18A26F22
          C7F6D3D4FEDEFFD73A7269AAFE4C3559DDEE2FB18601B9C3F719D6F87F253732
          71A298A6DA98A916376115037287EF315074231327EBAE08A6A9FA33D5E5EA76
          DFC22A06E40EDF65E938D9B17DD932C3696AFB03F6CA5D67A67A709BBADD6F62
          1D037287EF3056A8F665B9ED2399A6EACF543F55977B63DA7600728733795D74
          231377FBDE1166DD3EBDDC6A7466AA13D4EDBE84950CC81D4E674A4DD9B17DD9
          B241D169D1613A536DA72EF759D7B39601B9C36934901DDB972DFBC31B914D53
          F567AAD9EA761FCE5A06E40EDFF2E40532B52F5B3623C269AAF64C75DF0A75BB
          3FC36A06E40EDFD0F11EC98D4C9C94D5114E53F567AA3BD4E5BE88B61D80DCE1
          6BCE171EDB972D5B13E934557BA6DA2D45DDEECFB39E913B43005FD13D47766C
          5FB62C35E269AA8D99EAF8DEAC68E40E7092E7856A5FF6F346112D2A632E534D
          53B77B0B56347207A8A0F75CD18D4C9C1E914F53F567AAE9CA72CFEBC29A46EE
          0071D6C98EEDCB969590A6EACF54176B68DBC19A46EE00FFF84797EA1FDB979D
          493FC32D3ACA0383CE4C75ABBADDFBB0AA913BC0E4A6A21B993805A4A92632D5
          C3EA729F3B8C758DDC21F2F411DDC8C4799B34D54CA65AA86EF78759D7C81DA2
          CE2B37CB8EEDCB96CD8E6C8B0EC399AA86B61D4F2D6565237788380FBF2B53FB
          B2A495112F2A632E53EDA56EF796AC6CE40ED166E97D921B990A8692A69ACA54
          0F66A9DBFD75D636728748D352766C5FB62CF94B9E411ACB5447ABCB7DFE1416
          3772870873A9F0D8BEEC9E9EA4A9E632D5FD19EA766FC0EA46EE105D663E9AE8
          D3F6AFC9244DB53C539DF539EB1BB943646920BA9159B6EC17495349538D66AA
          CD69DB01C81DC43C7981E446A682FAA4A986BF532D52B7FB245638728788325C
          746C8F1FDCB793A69ACE54A76968DBB180258EDC21923C2353FBB2657B5B93A6
          3A67AAFBF565AAC9EA76BF8C358EDC218A2C5824BA9159B6EC781A696A2032D5
          4DBB59E5C81D22C8F3B263FBB2FCACBB48533DC85497ABDB7D1DAB1CB943F4E8
          FD9FA2637B7E7E9EE13435302D3A4C7FA73A58BD6DC742D6397287C8D142A6F6
          FCFCDCF6A4A99E64AA9FAA1FDD9BB2CE913B448D01A21B9938CB469875FBF440
          BBBDBC93C64C7582BADDFBB2D2913B448BC94364C7F6FCFC41B4E8F02A536DA7
          2EF7CD3359EBC81D22459F848FED5FA93DBFE833D254CF32D566EA761FC85A47
          EE1025BA3696DCC854308334D5BB4C75DF0A65B91FB99AD58EDC21423C2CBA91
          C9CFFF450A69AA9799EA0EF5A3FB3C563B7287E8D0E4BF64C7F6FCFC35A4A95E
          66AADD52D4EDFE67D63B7287C8D052746CCFCF3F9E4A9A1AB84CB56677163C72
          878870934CEDF9F9831B5154C6E34C3553DDEEAD58F1C81DA2C194F9A21B9938
          3D48533DCF54D395E53EBE166B1EB9432458223BB6E71797AC264DF53C535DAC
          7E74EFCC9A47EE1005AE7F41A4F6FCE2E27EB4E8F02153DDAA6EF701AC7AE40E
          1160B8ECD85E5C5C409AEA47A67A585DEE399359F6C81D42CF33C2637BF1B69E
          A4A9BE64AA85EA769FC3BA47EE1076BA2F921DDB8B8B67D3A2C39F4C5543DB8E
          0E5D59F9C81D424E2BA1DA8B935686A3A8CCE58F052E53EDA56EF79758F9C81D
          C24DEF4D921B990A8606354D7DE8DAC5CDB30B0B26A424150DCEFB4A75C70767
          25A52467A435EBF5E0759D0290A96629CBBDB8096B1FB943A869213BB6171727
          7F19BC67901F3CB83C6D7B525EE5DA5BB1B520B3F9DD9DACCE5447AB1FDDB7B0
          F6913B84990142B517AF08589ADAE9A2E5195B13F8FEA73437ADEE75D666AAFB
          33D4ED7E13AB1FB9437899DC467423535CFC8BEC00A5A99D2ECA4E16F51F4DCA
          D02AF87FB32A536D4CDB0EE40EE1658EECD85E5C9C35352869EAE5BD32946EA8
          B7A68DE96461A6DA5CDDEE2359FFC81DC24AD7CDA2637BFCE05E3F1869EAE575
          27E4AB5B70706A8D4E21CC54675DCF0E40EE1052AEA84E1FBDEFABBDB8787B10
          D2D40F9A276B30FB5714A5EEB02C539DA6FE47B56507207708274DFE4B722313
          67596BFBD3D48B0A07C7B492D4EC55AB32D564F53FE919F600728750B265AFE8
          D85E5C9C697B9AFA50DD94987EF20BC6842B53ADB9804D80DC2184DC243BB617
          A7FFC1F234F5D5ECAC9821529A7F6C4DA6BA9CB61D80DCC18199F3456A8F6338
          4D556CD17171E18A9841B29A7D6049A67A50FDDE697C6FF6017287D031527423
          1327B7BDC569EA7566D5AEAE778D99EAA7EA7F4B0BF6017287B0F1C97AD9B1BD
          F81786D3D4E92A173285F9310FC85AFEB11599EA04F53FA50B3B01B943C8682B
          3BB617170FB2364D7D287B70CC23926A8424536D43DB0EE40EE1E211A1DA8B8B
          3EB3354D6D9E15F39009175990A93653FF3BFAB017903B8489EEAB24373215CC
          B0344DBD7642CC5BD2D31EF33D53DDA71E30CC7D85DD80DC2144B4921DDB8B8B
          53EC4C531FCA5C11F39CAC5EBE67AA3BD4FF8A87D90DC81DC243AD4DB2637B71
          F1082BD3D4DBB7C67CA1E0559F33D56EEA1F6B3DB594FD80DC213474DE2B537B
          71AAE1A232A234F5E3CC97633E213BBC6BCC54DBA9FF0D2DD90FC81DC2C200D1
          8D4C7171FAFFBE616151996B53623E92FA98BF996AA6FA9FF03A3B02B9433898
          DC46766C2FCEEB61619ADABC34E62B5BEFF637534D57FE0BE64F614F20770805
          73846ACF2B596D5D9AFA41AAF495484ECB077636DC78EEAD71CEDBD8BFCEB151
          355F90FD2F95D6F535535DACFEEFA725EC09E40E61A0EB5CC18D4C85DBF3FA59
          D7A2437025B3B965FFE7EE3DFA789903B537DC7F6E9D21893B3EF55FFCCC54D5
          C3E4594FB22B903B84809764C7F6BCF402EBD2D41D450949EC8596E7DD59AFAC
          2AAEDC7043FF9CE284FE87735FF531533DAC7E741FCEAE40EE107C9AEC951DDB
          F34A1BD996A6D64DA092CC9177CEBDF0CAB26A53FBCEFEAB12793573B78F996A
          A1BADD27B12F903B049E2DA2637BFCE03EDBB6A232D57F28F242C75B6A9725CC
          D1E74655FB00BF6DB17F99AA8612338B68DB81DC21E84C94A93D2F2F69A55D45
          653A55374AFDCD03B7DC5126A4DE7343AADBA6A9B97F996A0D75BB5FC6CE40EE
          106C663696DCC85430D4AE34F5836AF610DDF55CED322536349C5BBD7FD272DF
          32D583EA25D336ED666F207708342365C7F6BCBCE42FAD7A06F9586EB52EDA1F
          B8F05099328FDFB0AB5A82CCF62D531DAD7E745FC7DE40EE1064AE5E2F537BDE
          8A9E56A5A91F54C7EDEB1B1E2DD3C4D3630DD85D5FA6BA3F4359EE790BD91DC8
          1D024C5BD18D4C9C4CABD2D4EA9CDB376DAC57A6914B1EC8D36D77BB32D5A6EC
          0EE40EC1E5CFB2637B5ECCAE34F5E3ED55AAEABFFA6B557B052FB6ACDA90CDFC
          CA549BABDBBD2FFB03B94350E95EB3FAC7F6BC33A86F539ADAA9EA4B88078E96
          19E0DED7AABCDC68EED777AAEA99EAE661EC10E40E01A595E846267E70DF6E55
          9A9A5695A6563D5B66862B6FDD54D58BC8C53E65AAD3D48FEE03D921C81D8249
          AD4D22B5E7C5F25BDB94A656D538F437E73E5E668C7AC7AAFA9AE9229F32D564
          65B91FB99A3D82DC219074961DDB6331ABD2D41A55049B43369419E5FE2A9EBD
          67BD1AD84C751E7B04B94310F9BD50EDB1A2A916A5A9770FAEBC69DC7987CA0C
          53AF8A6791B90FF993A92E57B7FB5FD925C81D82C7E41CC98D4C0586D3D4845A
          745C9E5479DF890BCB3CE0A7B32AAF00ECD377AA83D5DB7674679F2077081CB7
          C98EEDF193687B8BD2D4824AE5D4B67699275C32BFD29FD1DC9F4CF553F5A37B
          03F6097287A0D1B583ECD81E8B194E53A727E2F64AAF1E8A371EF2C6ED65876A
          8FAAB437D3B5FE64AA1394E5BEBE163B05B943C07829D1A7ED5FEFF74116A5A9
          171DAFAC9FD02D65DE7165FFCA1C59F2505033D5CEEC14E40EC162E17B921B99
          8A34F5337BD2D4872AEB2837F7C2324FB9B5B25AEF69FE64AACDD4ED4EDB0EE4
          0EC1E21DD18D4C9C1916A5A9957DBD34FFEC328FB9FF48255FAA3EE84BA6BA6F
          85B2DC7326B359903B048889B2637B2C9662519AFA60BABB936A1E2DF39C5B2A
          793493F4812F99EA0EF5A3FB6DEC16E40EC161586399DA637923EC49532BBB94
          5955AFCC07EE9DA5E762465FA6DA2D4559EE1DBAB25F903B048691A21B998A27
          DB868BCA2492A666BB0BE9515FDC5E5676A7FBCD4CFADDBE64AAEDD48FEE2FB1
          5F903B0485ABFF4B746C8FC5B6BD614F9A7A5DA9AB8E1A1F2DF389FBDD53D55C
          7F32D54C65B917D3B603B9435098273BB6C7623D2C4A53DD3F5FDAB4A1CC377E
          ABE553269D996ABAB2DDB7B063903B0483BF0AD51E2B596D4F9A3AC6BD98E1B3
          653ED2D0BD82983F99EA62F58B9989EC19E40E41604A4DC98D4C056B2C6AD1E1
          DE58EF2A3FDD5E7668AC96B64C1A33D5ADCA726F3C935D83DC210034F885E8D8
          1E8B655894A6D67035D1CE327FA95DD3ED97153DE64BA67A58FDE83E925D83DC
          C17E6ABD203BB6C78E37B2A74547A712370FEDBAC367B9975DB2DEEDB765FA93
          A9162ACB7DD627EC1BE40ED6D3394FA4F6586CB64545657AB969E885B3CB7CE7
          2AB71F37F8725F32550D2566DAB26F903BD8CEA474C98D4CC537962B2D6AD191
          62E785FB295CAFDDB3FDA9FD5B43DDEE8FB073903BD8CD821CD9B13D16EB6751
          9AEAFA557D5B1BDC5E56AF83DB83998FFDC954B394E55E73017B07B983D55C26
          547B2CF94B7B9E4196BB357F7EA19E15723F7483DB28D6F5A79FEA68F5A37B2B
          F60E72079BD9DD4172235391A6F6B4274D2DBFD6EDC39C9F9659825BEF8E9472
          5F32D5FD19CA721FDF9BDD83DCC162D6C98EED76A5A9E56E9FD4E75C698BDC37
          B81599B9A83CA8996A0B760F72077B59F80BD1B13D16CBB2294DEDE47685FC6C
          9935B835662A2CF727536DAE6EF701EC1FE40ED6D254A6F6586CA84569AAEB07
          4C63ED717B59BD4D2EAF213F280F6AA6DA86B61DC81D6CA5AFE84626CE769BD2
          54B7926179975824F7B2F35C86B257B93F99EA34F5A3FB1C761072073B19B659
          766C8FA5DB94A6963F546AF333C86FAA10B81CDD0BCA7DCA549395E53EF715F6
          1072072B19283BB6C7629936A5A9AE5FA75E6895DCDDCA439626742FD3E9A04D
          99EAC3EC21E40E36F2D66F646A8F154DB5284D2D2F7779D637CA2EB797D5FB48
          C3BD8CCE4C75B97ADB8EA5EC22E40E16324F722353417D8B5A74C40FB3839D7F
          E5FD96C9BDEC01E7DF9991D8BFCAF465AA07072BDBBD25BB08B9837D5C233BB6
          C762B956A5A9E50F3AFFCACD8FDB26F7675D5E9526F6E76ACC543F55BF98B994
          7D84DCC136A6CCFF85E8D81ECB6B6DD6EDD313939D5B5FEC8665D6E152D8FDEE
          72BF32D509CA729F3F859D84DCC1321AC8D41E8B0DB22A4D756DC17489756E3F
          749E7A4326EB32D525EC24E40E76F1F97AC98D4C459A7A9755696AF9C7C71D7F
          E62AFB0EEE651B5C4AB025F8176BCC549BA9B7ED7892BD84DCC12A3ACA8EEDB1
          D80CABD254D72BF78D16CABD6C918E4B779D99EABE15CA761FCE5E42EE601393
          64C7F6582CA5BD55696A79F9F2403C72AFF4A9FBC5E5BE65AA3BD42F66CE6737
          2177B0878A161D12B5C7F246D895A6BABD72DF74A58D72FFA5F3A0D648F46FD6
          97A9764B5196FB22DA762077B087CB44373271520D3F833C2B61B96F75FC9DC7
          6C747BD9E3B394FB64EBCE54DBA91FDD9F673F2177B085DD1D646A8F0D7EC3AE
          3435EE39E75BE35BAD947BD9173A3E63D29BA9662ACB7D536F761472074B5827
          747BAC8765696A79F9752E57EE87CA8273E99E92F05FAD33534DA76D0720F7B0
          D0A558A6F658C96ACBD2D4F2F2C58E3FF4A93BAC74FBA1FB1D7FEDB6C4FF6C8D
          99EA6265B9E775614F2177B082A632B5C7626B6C6AD151D963995576DECAB8BD
          747F35F1BF5B63A6BA55D9EE4DD953C81D6CA0AFD4ED19D6A5A9E5E5690128E5
          FE2D57163BFEDC4F0547777D99EA61F54CB52FBB0AB983FF0CDB2C537BACB491
          4D2D3A4E911A94C2325FF1A89EB7907A33D54265B96F1EC6BE42EEE03B03656A
          379FA64A0EEEE5CEFD849EB355EEA31C7F6E5DC95FFE4F36959819C8BE42EEE0
          376F1D11BA3D69B56DCF202B4809462DF7CA6BBA674BFE728D996A0D65B91F79
          8B9D85DCC167E6C9D41E8BF5B32F4D8D93E4F85B9FB6D4ED871AEAF98A4973A6
          7A304BD9EEF3D859C81DFCE575A1DA630516A6A9719CB5B4C1D693BB73D5DF42
          D19FAE31531DAD7E31730D7B0BB9839F4C992F74FB8A9E16A6A9AE723FDB56B9
          DFAA51EE1A33D5FD19EA6D3BBAB3BB903BF8480399DA63B1D936A6A9719CFB80
          D6B355EE3FD5547FC0C24CB501BB0BB9837F3CB95EE8F6AC9536A6A9AE72AF6D
          ABDCAFD22A778D996A7365B9AFAFC5FE42EEE01BC3A53B77A89569AAEBB54CC0
          E45E28FDEB357EA7AA9EA976667F2177F08BF3A5FB76BBE134F51CB1DC938275
          2DE37CE79E26FEF3F565AAD3D42F6626B1C3903BF8C38245C25D9B6F699AEAFA
          CE3D60816AB6F8CFD798A9262BCB3D87B61DC81DFCE179E9AECDB4344D8D3321
          0C725F2EFFFBADCA546F638F2177F083DE9B847B366BAAA569AA6BF9016BE57E
          AEE3CF6D2EFFFB3566AACB95E5DEA12BBB0CB9830FB490EED9FAD6B5E8A8AA70
          D8AF6D95FB466D85C30C7CA73A58D9EEEBD865C81DBCA78BF015642CD7DA34D5
          B5E4EF8BC192FB0E9511D097A97EAA2CF7E285EC33E40E9E3344DA67A7B559B7
          4F57315B79B673973D5BE5EE5C5BE66E9511D098A94EA06D077287E0D147BA5F
          07D99BA6BA7662FAA5AD72EFEFF8732F561A02AB32D589EC34E40EDEF2CA5CE1
          6E2DBACBDE34358EF3979577DA5A15B28EE3CFFD5869083466AACDD4DB76CC64
          AF2177F09487A5BB7586C5696A9C1A21A8E75EAA36041A33D57DA5CA761FC95E
          43EEE0254B9F12EED5149BD3D4380F3AFEEAAB6C95FB58C7A7A68A63A03153DD
          A12CF78F3E61B72177F09096D2347584CD696A9C6B1D7FF64F6D95FB3B4EBF76
          ABEA20E8CB54BBA528DBBD2DBB0DB983775C2ADDA9A976B6E8F896571D7FF6B9
          B6CADDF1C952AEEA2068CC54DBA967AA8FB0DF903B78C594F9C27D3AF833ABD3
          D48A34D1F1776FB455EEAF39FDDA64F551D097A9662ACBBD266D3B903B78C592
          90A6A9AE05DD1BDA2A77C77FCBA6AA8F82C64C355FD9EEADD871C81DBCE1C959
          C25D5AD2DEEE34D5B5E6EF4E5BE53E576FC55F1399EA6265B98FEFCD9E43EEE0
          09E2161D6B6C6DD15155CDDF8EB6CA7D96DE8ABF4632D5ADB4ED40EE100CCE0F
          6D9AEA5AF377ACA56E7F5C77C55F1399EA61F54C7500BB0EB98379C42D3A4A1B
          59DBA2E3340A9C7EFA1796CABD9EEE8ABF4632D54265B9B799CCBE43EE601C71
          8B8E1E561795A9B4E6EF6B96CAFD6CED157F8D64AAE9CA769FC3BE43EE609ADE
          E385FB3369B5EDCF20DD6BFE36B654EE2F3A8EF4837A46425FA65A4359EE7369
          DB81DCC134E2161DFD0290A6C671AC75B5C952B9FFD271A4AFD533121A33D524
          65BB5FC1CE43EE6016718B8E8220A4A971EA3A564DB054EEF73B0EF5639A8642
          5FA63A5ABD6D4713F61E7207934C6E23DC9C2B7A06214D752D0B59DB4EB95FE5
          F45B8FEB1A0A7D99EAFE0C65BB6F61F321773089B845C7EC40A4A9AE65218FDA
          29F75B8D14853490A96A68DB7129BB0FB98339C42D3A925606224D752D0B6969
          87ECF38C14853491A93657967BE329EC3FE40EC610B7E8181A8C3435CEE58EBF
          FFE90075D9CBD537181A33D52C65BB2F61FF217730C5D262E1BEDC6E798B8ED3
          714C8C6FB153EE1DCD14853491A94E5396FBAC27D981C81D0CB145B82DF38392
          A65650149C564C87C61A2A0A692053ED96AC6CF7E1EC40E40E6610B7E8C80C4A
          9AEA5A16D2D26E1D434C1585B433533D9F3D88DCC104531A0BF764D6D4A0A4A9
          15E406A8A07B4DA7DFDA4CEB70E8CB54972BCB7DD1027621720703885B74D4B7
          BF45C769385E1FEC3C149C72EE75B50E87BE4CF5E06065BB3FCF2E44EEA09FEB
          A52D3A720394A6BA550E7BC0CE93FB53E6EA865999A9D2B603B98301A42D3AD2
          5B9B75FB74BD2E73AE1C36CA4AB7D7365937CCC077AA1394EDDE827D88DC4137
          CF48F7E3A020A5A971B29DFE881C2B1FCB1C355937CCCA4C35AF0B3B11B9835E
          BA4B5B7414DD15A434358E63EE373F40157F5FD53D22FA32D566CA761FC25644
          EEA017718B8E19814A53E3F472BCECB552EEF73A1E6E750F88C64CF5C35265BB
          F7612F2277D089B845476EFB40A5A971C638FE1D57DA28F71B1CFF5349FB8894
          FFAB36BBEF506FDB318CDD88DC4123D2161DE9238295A6C6B9DBF10FA967A3DC
          7FEA58A34DFF9068FC4E3545D9EE03D98DC81DF431C0D234F5CBB3F49BECD5E0
          9485DCE8F44B53F40F49F98FB41DDDDB29CBFDA9B7D88FC81D74216ED131F8B3
          80A5A915C754C7BFE49736CA7DA7E9BA610632D54C65BBCF63432277D0C59CC8
          A4A915387E4879838D727FC074DD300399EABE7C65BBBFCE8E44EEA007718B8E
          92C0A5A9AE95C37E6AA3DC4719AF1B6620535DAC2CF7F9B4ED40EEA0872BA4BB
          704D605A749C8E63E6B7D146B9AF72FAA5D946064563A6BA55D9EE0DD893C81D
          74D044DAA2233578696A05DB1D2B8745B46E98814CF5B0B2DCD77FCEAE44EEA0
          01698B8E6D8D02D4A2239095C38E785037CC40A65AA86CF78EEC4AE40EEADC24
          DD813D025654267095C33CA91B6622534D57B6FB24F625720755664A5B746C5D
          1DBC6790EE95C3165928F70D9ED40D3390A9D650967B0E6D3B903BA83252BAFF
          FA05324D8D53D7E9AFD96CA1DC9F761CF7C74C8D8BC64C3549D9EE97B133913B
          A8216ED15110CC34358EE3B9F2888572BFC5E987AE30362E1A33D5D1CA72EFD0
          95BD89DC4189B6C2CDB722A0696A9C071DFFA0DAF6C9FDB78E0D6BCD0D8CC64C
          3555D9EEEBD89BC81D5410B7E8981DD03435CE758E7FD006FBE47E9E57A565F4
          67AAEA6D3B8A17B23B913BC8595053B8F5925606344D8DF381E35FF4B47D72EF
          EFF43BB71B1C198D996A7365BB37657B227790D34ABAF38606354DADA03420C5
          651C4BCB64981C198D996A96B2DD27B23F913B4811B7E848369CA69E6354EE8E
          6F399E8B7469191B33D5CD33D9A1C81D8474166EBBE33D039BA606A9B88C87A5
          65F467AAFB0B68DB81DCC12FC42D3A32839BA606A9B88C97A5656CCC543FFA84
          3D8ADC4182B845475680D354D7E232C70ED9E6F62B635E9696D19FA9D6A56D07
          72077F10B7E8A81FC4161D5515971962DDC1BD9EA7A565BEC954BBE992FBC1C1
          CA76FF33BB14B943E274ED20DC7113029DA6BA159779D43AB9BFE86D6919FD99
          EA3465B9D7ECCE3E45EE90302F09375C7E6BB36E9F6EDA5ECEC565C65B27F73B
          BD2D2DF30DFFAD2D53DDAE6CF756EC53E40E89226ED19116EC34D5ADB84CEC0E
          DBE4FE638F4BCB7CCDBF5894A98EAFC54E45EE9020D2161D4553839DA6BA1697
          396A99DB0F9DEB756919FD996A3365BB7766A72277488C9BA29AA6BA1697B9D0
          B6937B43AF4BCBE8CF543F2C55B6FB00F62A72874410B7E8C86D1FF034D5B5B8
          CC2DB6C9BDA3E7A565F467AA3BD4DB764C66B722774800698B8EBCC0A7A9AEC5
          657E6B9BDC5B7A5E5A467FA6DA2D57D9EE73D8ADC81DAACF27D2161D8302DBA2
          A3CAE232E705A2FA409A2703A42F536D47DB0EE40E5E226DD131F8B3C0A7A9AE
          C56576DAF689EA5C1F4ACBE8CF543395EDFE12FB15B943757944BACF66043F4D
          752D2ED3D632B73F1EF3A1B48CFE4C755FBE72DB8E26EC58E40ED5A3BBB44547
          4A08D254D7E232BB2C93FB515F4ACBE8CF54172B1FDDB7B065913B540F698B8E
          F435614853DD8ACB34B64CEE17FA525AC640A6BA55D9EE37B167913B5487DEEB
          857B2C351469AA5B7199DF5826F75BFC292DA33F533DAC2CF7C653D8B5C81DAA
          81B445C7B63702DDA2A3AAE232B1DA76C9FD399F4ACBE8CF540B95EDBE845D8B
          DCA16AC42D3A7A04BEA84CE5C5657E6D97DC37FA545AC640A69AAE2AF759D7B3
          6F913B54C5E41CE106DBBA3A14CF202BF8D4F10FBCD7AED2323BFD2A2DA33F53
          ADA17C741FCEC645EE5015E2161DFDCCBAFD460FB575B1E31F78835D0FDDDBFA
          555AC640A69AA46CF767D8B9C81D2A47DCA2A3202C696AC58D439ED35F78AE5D
          D732BB9C7E63B29772D797A98E5696FB22DA762077A81C698B8E158DC292A656
          50E4F42736B4EBE4EE58DA2DD5D361D297A9A62ADBFD79F62E7287CA58286DD1
          313B34696A058E4FAF1FB0EBE4FE91D36FCCF47498F465AA1ADA76F466F72277
          A804698B8EA495A149532B98E0F4378EB2CAEDB51DE761B9B7E3A42F536DAE6C
          F716EC5EE40EEE4C94EEACA1E149532BC870ECC66C95DC7FED380FBD3C1E287D
          996A96B2DDBBB07F913BB8216ED1916C384D3DC76367A5D9DF22DBB93DF60E8F
          07CAA64CB50D1B18B9831BD2161DC77B86294D75AB3F10FB8B4D72BFCADFEA03
          DA33D5FD05CA76EFC30E46EEE0CC271F91A67E85F30DF0069BE47E9ECFD5072C
          CC54E7BEC21E46EEE088B44547D6D450A5A9AEF5079EB549EE8E1FA8967A3E52
          1A33D5BACA767F983D8CDCC1893F4BF754FD70B4E8388D8B9C3F51B5FE03D524
          EFE5AE2F533D385855EE4F2D65172377F83EE2161DDB4396A6BAD61FB0EA13D5
          5DBE571FD09EA94E533EBAB7641B2377F83ED2161DF9AD43D2A2E3F49B64C73A
          85FDADFF4035D987B1D298A96E57B6FBEBEC63E40EDFA5D678E17ECA0C5B9A5A
          4196ED9FA81E2AF6BFFA808599EA7CDA762077F82ED2161D45A14B532B2871FA
          53BFB0A8B84CBD98FFD507F467AACD94EDDE809D8CDCE14C0690A656597F60BE
          4527F7176DA83EA03D53FDB054B96DC7E7EC65E40EA7236ED191DB3E7469AA6B
          FD8159D67750EDE5CF68E9CB546FA76D077207BDDC26DD4B214C535DEB0FD8D4
          45D5B983EA833E0D97BE4C3557D9EE93D8CDC81DBE45DCA26350885A749C4E33
          DBBBA86EB4A3FA80F64CB59D7ADB8E05EC67E40EDF206DD151F45918D354D7FA
          0377DA23F73A8E3FF003BFC64B5FA69AAD6CF7CBD8CFC81DBE46DCA2634628D3
          D4388B1DFFDCABEC917B4BA7DFB7ADDC37B465AAFBF255E5BE69373B1AB9C329
          9A0AB7514A38D3D438773BFEBDE7D923F7D76CA93EA03D535DAC7C745FC78E46
          EEF015D2161D7923C299A6C679D5F10FDE69CF4377C79024D7BF01D397A976DB
          AA2AF7BC85EC69E40E15CCDC2CDC44A9214D532B70BC1C186B8DDBEF709C9002
          1F074C5FA67A58F9E8DE944D8DDCA18281C22DB4ED8D90B5E8A8B2FE408E3572
          DFE03823857E8E98BE4C354DD9EE7DD9D5C81DFEF18FABA52D3A7A84B1A8CCD7
          A4387683B046EEBF749C916C3F474C63A69AAE2AF7CDC3D8D7C81DFE314FB881
          4A5687F419E4499C0B143E6E8BDC7F6C53F501ED996A0DE5A3FB40F635720771
          8B8E3566DD7EA3AFA22A4F75FC9BCFB6BBC95E0D7FC74C5FA69AA42AF72357B3
          B3917BD411B7E8C808719AEA5A7FE069AB9BECF9567D407BA63A5AF9E83E8FAD
          8DDCA38EB445C7F146214E535DEB0F58D368CFB9DDED753E0F9ABE4C3555D9EE
          7F666F23F768536BBD70EFCC0E739AEA5A7FC09A467B8E4DF6620FF93C68FA32
          55F5B61D3569DB81DCA38DB44547D2CA30A7A971C6C4AC6EB437D7E9D70DF67B
          D03466AACD95EDDE8ADD8DDCA3CC24E9CEE917EA34D5B5FEC0314BDCFE789E65
          D507F467AA59AA721F5F8BFD8DDCA3CB02698B8E64C369EA39BE5BEA72C7BF7B
          9725723F1AB3ADFA8085996A673638728F2E9749D3D49EE14E532B58E1F48737
          B644EE4F5B577D407BA6BABF40D9EE03D8E1C83DAA885B74843D4DADC0F1A9F5
          5357DA21F71B1CA725CD8261B32953CD99CC1E47EE11659D70D364853D4DADC0
          B9DF5B3D3BEA42DE6A61F501ED996A5D65BBDFC61E47EED144DCA26368585B74
          9C8EF3ADC085769CDCFB3BFEB8E6368C9BBE4CF5E06055B977E8CA2E47EE9144
          DAA2637BE8D3D40A0A1DFFF6FBED90FB31C71FB7D88A81D397A94E533EBABFC4
          2E47EE51A4AF70C3A41B4E53A75BA1A872E75E9EB7DAFC0DD3DD768C9CBE4C75
          BBAADC8B9BB0CF917BF418266DD191198134358EF38D6F433BEEDC9DE7EE723B
          46CEA64C750B1B1DB9470F698B8EACA9114853E338979D7DC08E6F981CD39215
          968C9CC64CB599B2DD6F62A723F7A871F511E16EA91F8534D5F513553BBE6272
          FE8629C916B9EBCB543F2C55957BE399EC75E41E31A42D3A722391A696BBB5C8
          B6E32BA6A76DFD40557BA67ABBF2D17D247B1DB9478B6BA48DE55B47224DADE0
          B8634267C5574C3758FB81AAEE4CB55BAEAADC675DCF6E47EE51A2FB7CE15619
          14EE161D557EA21A3B6A6F1FA6347BC64E5FA6DA4EF9E8DE96ED8EDCA34403E1
          4629BA2B1A696A0529F6F662DA69EF07AADA33D56C65BB3FC37E47EED141DCA2
          634644D2D40A92EDEDC534D6F1A7D5B568F0F465AAFB5E566EDBD19D1D8FDC23
          4347E13649691F9134D5FD13D5736D78E89E63F107AADA33D5C5B4ED40EE505D
          A42D3AF24644264D75FD44D58A5E4C1D6CFE40557BA65AA2DCB6A3377B1EB947
          83058B849B24353A69AAEB27AA6D2D70FB1D31AB3F50D59DA91E563EBAB760D3
          23F768206DD131F88DF0B7E8A8F213551BBE62DA10B3FA0355ED996A9AB2DDBB
          B0EB917B14D8BD49B8437A44A3A8CCD75CE43808732D90FBBD967FA0AA3D534D
          57957B1BDA7620F728206DD151B23A3ACF20DD3F518DDDE1BFDCAFB2FC0355ED
          996A0DE5A37B1FF63D720F3F0BF384FB638D59B7DF689B9B9C3F518D9D6DEB37
          4C05D60DA0BE4C354955EE735F61E723F7D0236DD19111A934D5FD13D55FFA2F
          F73AD67FA0AA3B531DAD7C747F989D8FDCC38EB4454769A348A5A915389735B9
          CA7FB98F72FC61CDEC1B417D996AAA72DB8EA5EC7DE41E6EC42D3A2296A656E0
          DC45F53CFFE55ED3E20EAA863255F5B61D2DD9FCC83DDC485B7424452C4DADC0
          F9055E1DFFE5FE1BC71F36C6C221D497A93657B6FBA5EC7EE41E66DE92B6E8E8
          17B534358E731BA051BEBBBD9EF3145D6BE318EACB54B354E53E7F0AFB1FB987
          18698B8E82A8B4E8381DE7C3624DDFE57EA1F31C3D64E318DA94A92E61FF23F7
          F0F2BA705BACE819B93435CE18E7E60FBECBFD7EE76ACC568EA1BE4C757F8172
          DB8E273100720F2B53A42D3A66472F4D8D73ADF360D4F35BEEE73AFEACAD76CA
          DDA64C75380A40EE6145DAA2236B65F4D2D4380F398FC6857ECBBDBFE3CFDA6E
          A9DCF565AA7595ED3E090720F770F2F92CE19E181AA1161DA753E4381AF7FB2D
          F7B6CE253B2D1D447D99EAC1C1AA725FB4000B20F750325CB825B647314DADC0
          B990F8AD3EBBFD9073AB8E4C5B47B1FC9F74D97D9AF2D1FD792C80DCC388B445
          47BEE13475BAB556DA6E67BB0EE7561DCBAD1D467D99EA7655B96FDA8D07907B
          F810B7E8C88C649A5A81F337EFC77C76FB5F9CA7A986B5C36853A6BA0E1120F7
          F0F1BC344D9D1AC934B5824CC701C9F159EEBF769EA78BEC1D477D996A3355B9
          E7D1B603B9878EDED2161DF5239AA6C6596E65BB8E3B9DE7E9557BC7515FA6FA
          61A9AADD9BA202E41E365A0837436E54D3D472B7467B7EB7EB78CEF1471D2FB7
          196D99EAEDCA17337D7101720F175D842D3AF25A47354D2D776BB417DBE0AFDC
          1B3ADF9E592D776D996AB75C55B96F1E860D907BA890B6E81814B9161DA7E1D2
          68EF4E7FE5DED1F147A5582D777D996A3BE5A3FB406C80DCC384B44547D15D91
          4D532B706EB4F75B7FE53E24204DF60C65AAD9AA723F72353E40EEE161D85CE1
          469811DD34B502E7467B0DFD957BE3A034D93393A9EE7B59D5EEF31002720F0F
          D2161D29114E532B70BEE0EDE8ABDB1FCF0B4A933D4399EA18E58B996B300272
          0F0B6F3D254C534744384DAD20C3715886F82AF70DB1C034D93394A99628B7ED
          E88E13907B486829DC04A9514E532B706EB4D7D857B9DFEB3C570FDA3E963665
          AA0D7002720F07D2161D833F8B628B8EAABF62CA7BDC4FB95FE53C59D7593F98
          FA32D53455B9AFAF8515907B1810B7E8309DA65A7F702FEF65DF43F7431B1D7F
          527A27EB075363A69AAE6AF7CE6801B9878125C20D50D23ED2CF202BF8D47964
          EEB5EF997B567900D096A9D6A06D0772877FFCE349698B8E3566DD7E63006CE4
          F215D3557E9EDCBF08E2374CBA33D52455B9E74CC60CC83DF00C274DD5FD15D3
          C643D63D732F08C268EACB54472B1FDD6FC30CC83DE89C2F5CFCA58DA29EA6BA
          7FC5E4E743F7C78B03F90D93EE4C355555EE1DBAE206E41E6CC42D3A7A443E4D
          75FF8AE90B1F4FEE67C782F90D93E64C55BD6DC74BC801B9071B698B8EA4D591
          4F53DDBF626A6CDF33F7E6C1184F7D996A2F55B9172FC40EC83DC8885B74F423
          4D75FF8AA9F871EB9EB93F189001D597A966A9DA7D0B7A40EE4146DAA2A320E2
          45652AFF8A2976B66DCFDC03F00DD329FED99E4C75227E40EEC145DAA263454F
          D2D4CABE62FAA56DCFDC83F00D93E64C757F81AADC1BCFC410C83DB0B4112EFB
          D9A4A9957EC574956DCFDCB3CA03834599EA480C81DC834A1F699ABA9234B5D2
          AF9836DAF6CC3D253872D797A9D65595FBAC4F7004720F26AF485B740C359CA6
          1E088E898E5BF6D0DDE5997B4180E4AE2D533D3858D5EE6D9104720F260F0B97
          7C32696A155F31F9F7D0DDA59A7B5A8086545FA63A4DF962E6112C81DC83C852
          618B8E7CC369EAF4208928D7B287EE7706FA1B26DD99EA7655B9D7A46D07720F
          22D2161D99A4A9557DC5E45F45F7E79CE7AC5790C6D4A64CB5159E40EEC1E352
          E172CF9A4A9AFA2D999655746F18EC6F987467AACD54E53EBE37A640EE4143DC
          A2A3BEE183FB814069C8E52BA63BFD92FB31E7DF7371B0E4AE2D53DD574ADB0E
          E41E39A42D3A7249534FC7A52DC4737E3D73DFE5F8738E97070C6D99EAEDCA17
          33037005720F16D2161DE9AD49534FE76EE7616AE8D7C9BD83F387094193BBB6
          4CB55BAEAADCDBD0B603B9070B698B8E41B4E83883C79C87E9984F6EAFEDFC73
          26044DEEFA32D576CA47F739D802B90709698B8EA2BB4853CFA4C8719C76F9F4
          D0FD45E769CB08DCB0EACB54B355E53EF7157C81DC8383B8454773D2D4EF50E2
          DCC7C7A793FBFD2EAF578327777D99EACBAA767F186120F7E0F03C69AA2E929D
          47AAB63F723FD7F9D7D40DDEB8EACB54C728B7ED588A31907B50E83D5E98A68E
          204DFD2E85CE4375893F72DFE9FC6B16077060F565AA25B4ED40EE91A10569AA
          365CAE746FF147EE639D7FCDB5011C589B32D54B7106720F06D2161D833FA345
          C7F768EE3C56B7FA23F79ACEBFE683208EACBE4C354D55EEF3A7600DE41E0426
          4B5B74CCA0A8CCF779D079ACFAFB2377E7AF178A02E9768D996ABAAADD97A00D
          E41E04A42D3A4ADA93A67E9F8B9D076BAC2FDFA7D68B05BE5587914CB58672DB
          8E27F10672B71F718B8E35865B740453409DF21D076B952F07F7A75D0AF00754
          EEFA32D52455BB0F471CC8DD7EA42D3A5249531D71F6C67A5FE4FE63E7A92B0C
          E8D0EACB54472B67AAE7630EE46E3B4B8B658B7B5B23D254475C8A97D4F343EE
          1B43D0AAC34CA69AAA2AF7450B700772B79C2DC2C5DD8334D519E7761DB1A7FD
          907BC730B4EA3092A9AAB7ED781E772077BB91B6E8D8BA9AA232CEB8B4EBF8B1
          1F721F128A561D4632D55EAA72DFD41B7B20779B99D258B8B4FB91A6BAE0D2AE
          63A31F72DF1C8A561D6632D52C55BBB7401FC8DD66A42D3A0A282AE386CB33BB
          8E3EB8FD8E58385A75D899A9E675C11FC8DD5EAE17B6E858419AEACAB5CE4336
          C407B95F120B47AB0E2399EAFE0255BB374520C8DD5EA42D3A6693A6BAF290F3
          906DF641EEB784A55587A5996A5F0C82DC6DE519E1A24E5A499AEA8ECB65EE1D
          DECBFD56972F14023DBCFA32D5BAAA72DF3C0C8720773BE95E53B8A8871A4E53
          0F045A3E29CE83F66B6B0AFE66075BEEDA32D5838355ED3E108920773B69255C
          D2C9A4A99551604DD1DF962EEDB3823DBEFA32D569AA723FF2161641EE36226D
          D171BC27696A65A45953F4F7D1D03D73D79CA96E57B5FB3C3482DC6D44DAA223
          933455F2D07DA7E76EBFF248E89EB95B97A95E834790BB7D0C102EE72CD254D1
          43F7B187BC96FBD1103E73D79CA936536EDBD11D932077DB10B7E8A86FF8E07E
          20E8E67179E85ED3F393FBBD617CE6AE3753DD57AA6AF706A804B9DBC61CE162
          9E409A2A7BE8FE91E727F7DF86F199BBE64CF57655B9AFAF854B90BB5D485B74
          1C6F6DD6EDD343601E9787EE47BD967BC3503E73D79BA976CB55B57B476482DC
          EDE20AE1524EA34587F4A1FBB35ECBFD58289FB96BCE54DB2967AA93B00972B7
          8926C2161D7F984A9A2A7DE87E95D772CF09E73377CD996AB6AADC7326E313E4
          6E115B48533D7FE8BED1EB4BF74D217DE6AE39537D59D5EE97E113E46E0F3709
          97716E7BD254F143F7073C767BBD58489FB96BCE54C7A8CABD43578C82DC6D61
          A6B045471E69AAC243F75D1E9FDC2F0CED3377CD996A89AADDD7A114E46E0B23
          858B781069AAC243F74D1E9FDC7F1CDA67EEB665AAC50B710A72B703698B8EC1
          9F515446E1A17BAC9EB772DF18DE67EE9A33D534DA7620F790D056B884675054
          46E9A1FB85DECAFD81103F73D79CA9A6ABDA7D225641EE36206DD191429AAAF6
          D0FDC7DECA7D57889FB96BCE546B28B7ED98895790BBFF2C10B6E8485F63B845
          4778A4E3F2D07DA3152F219B87679CF565AA49AA761F895890BBFF485B74A492
          A656974C970FD53D757BED58989FB96BCE5447ABCAFDA34F300B72F71B698B8E
          6D6F90A6561797E69C436C7809197B354403AD2F534D55B57B5BD482DCFDA6B3
          70F1F6204DAD368B5D3E76B1E125E4B630B95D5FA6AADEB6E311DC82DCFD45DA
          A263EB6A8ACA549BEB5C06B1B6052F214B4225777D996A2F55B9D7A46D0772F7
          15718B8E7EA4A9099C27F39D07F1452FE5DED1A5BB79B8E4AE2F53CD52B57B2B
          F482DCFD44DAA2A380161D89E0F2F8E2062FE53EC4A56673C8E46E4FA6BABE37
          7E41EEFED1B5836CDDAE68449A9A08DB9D87F13C2FE5EE32D5CB4336D4DA32D5
          FD05AA76EF8C6090BB7FBC245CB6B3495313A2D07918EB58F012B246D8E46E51
          A63A00C32077BF90B6E8485A499A9A10CD9CC7F10B0B5E425E1BB6B1D697A9D6
          55957B1BDA762077BF90B6E8186A384D3D1036DFB83CBDD8ECA1DC6F70A9DAFC
          71E8E4AE2D533D3858D5EE73700C72F78789C2259B4C9A9A2017B98CE45FBC93
          FB79CEBF202B7C6E2FEFB45F93DDA7A9CA7D2E6D3B90BB2F485B741CEF499A9A
          208FB90CE525BEBF84CC0DA1DCF565AADB55ED7E059A41EE7E309234D5338A9C
          87F27EDF5F42A68651EEF664AAC54DF00C72F79E4F842D3AB2A692A6264C8AEF
          6F213785BEE0AF894CB599AADDB7201AE4EE3DD2161DF50D1FDC0F84D136192E
          6F210FF9FD12B27928E5AE2D53FDB054D5EE37611AE4EE358F0817EB76D25401
          993ED7853C7461040AFE9AC8546F57957BE329B806B97B4B77618B8EFCD666DD
          3E3D9CB27179323DD7EF9A90E12AF86B2053ED96AB6AF725C806B97B8BB44547
          262D3A248CF1B92EE4A18D5128F86B22536DA72AF759D7631BE4EE25BDD7CB56
          6A1169AA08B7A2BF9EF5C83E168982BF2632D56C55BB0F4737C8DD4B3A93A67A
          7B903CEE735DC89C6814FC3590A9EE7B59D5EEE7E31BE4EE1DD2161DB9ED4953
          6524F9DC237B7D340AFE9AC854C7A8CA7DD1028C83DCBD62728E709992A64AD9
          EE6F8FECA3B18814FC3591A996A8DAFD799483DCBDE236E1221D449A2AC5A5E8
          EF2E8FE4FE4B97195D1C62B9DB93A98EEF8D7390BB37485B74147D465119292E
          9F3A6EF248EEBF8D4CC15F13996A9AAADD5B201DE4EE0DD2161D33282A23A686
          CB98D6F346EEFD5D6AC0750AB3DCF565AAE98A72CFEB827590BB172C14B6E848
          214D9573ADCBA03EED8DDCC73AFFD3B786DAEDFA32D51AAA47F7216807B97B81
          B04547DE08C32D3A42ED998F5D0E7F577923F79A917B09A937534D52B57B1FBC
          83DCCD236DD1914A9A1AD8B790571E89DC4B48BD99EA68E5B61DC3300F72378D
          B445C7B63748530DBC853CE689DCCF8EE04B48BD996AAAAADD07A21EE46E1A69
          8B8E1EA4A926DE42E6785259E6CE08BE84D49BA92AB7ED78EA2DDC83DCCDF2C9
          47B2B559B29AA23226DE42BEE0C9C9FDD628BE84D49BA9F652B57B4BE483DCCD
          324FB834D790A69A790B79D40BB9EF8CE44B48BD996A96AADD5FC73EC8DD247F
          162ECC0C5A74187A0B79AF17727F27922F212DCB54E7D3B603B91B44DAA26345
          23D254436F217FEA85DC3747F325E449FE4D93DCF717A8DABD010242EEE690B6
          E8984D9A6AEA2D647F1F1BA8A64541EEF664AAEB3FC740C8DD14B5C6CB5665D2
          4AD254536F215B7A20F70BA3FA12526FA65A57D5EE1D5110723785B445473FC3
          69EA812838C6E52DE47C1F1BA82E8E84DCB565AA0707ABDA7D120E42EE6690B6
          E848264D35F71632EF0EDF1AA846E025A4DE4C751A6D3B90BB9D485B741CEF49
          9A6AF02DE425BE35508DC24B48BD99EA7655BB5F868690BB096E234DB5F12DE4
          FDE6E5BE28BA2F212DCB543BECC643C85D3FD2161D59A4A946DF429E67BE6CD8
          6F22FC12526FA6DA4CD5EEEB101172D7CF3AE1721C6AF8E07E202A8649F2AB8D
          EAD9917E09A93553FDB05451EEC50B311172D78DB445C776D2544D24BBB4513D
          645AEE7746FA25A4DE4CF576D5A37B535484DC75D354B616D30DA7A9D3A32398
          42BFDAA8DE1AED97905A33D56EB9AA769F888B90BB5EA42D3A3269D1A18BE57E
          B551ADE3F20FBE2E4272D796A9B65395FBE699D808B9EB64E666619A3A953455
          178B5DC6F8599FCA869546C9EDFA32D56CDA762077AB18285C88F54953B5719D
          CB18FFD6B4DCE7BA94E88F94DCB565AAFB5E5694FB4757E323E4AE8FAB852D3A
          72495335B2CD9FD2616E65C332A225776D99EA18D5A3FB3C8484DCF5216CD191
          D79A34552329CEA33CD6A7B26199D11A7D7D996A89AADDFF8C9190BB2EFE2A5C
          848348537592E14FE9B0AB5C66B779C4E46E4FA65AB33B4E42EE7AE83E5FB606
          8BEEA2A88C4E5CA2B8E2C7CDCABDA1CBF45E1435B96BCB54D354EDDE0A292177
          3D34102EC119A4A95A69EE4FE9B0B12E776E1F444DEEFA32D57445B98FAF8595
          90BB0E6AAD97ADC094F6A4A95AB9C865A06F302BF7475D5EB946CEEDFA32D51A
          AA47F7CE6809B9EB40D8A2236F84E1161D9173CB0779CE23BDD1A8DBEF70A93B
          31217A72D797A926A9DA7D005E42EEEA4C122EBF54D254DDB828E101A372BFC4
          657E0B2328776D99EA6855B9E74CC64CC85D9505C2161D83DF204DD58D4BAF87
          D78CCAFD06CA8619C8540B55ED3E07352177552E132EBE1EB4E8D08E8B11665D
          6952EE1B291B662053556FDBD11537217735A42D3A4A565354463B6EA5C3CEF6
          A1C75EB4CA86E9CF547BA9DAFD25E484DCD590B6E858439AAA1FB7D261B7F8D0
          636F5B7934D196A966A9B6ED68829D90BB0A0BF3EC4C53CF89A4582E7619ED73
          4DF6D83BE2F2D035A272B72753DD829E90BB0AC2161DA58D48534DE0523AAC8E
          C1664C1B281B662653DD5FA06AF79BF0137297D39734D52A5C4A870D317872BF
          DF658AB3A32A777B32D5C6B4ED40EE6286095B742491A69A21C3FB4E7BE75136
          CC54A65A57D5EE2351147297226DD1D18F34D50CD9DE77DAEB48D9305399EAC1
          2245B9CFBA1E47217719571F91ADB902D25443B8BD9FFBA539B9E75036CC58A6
          3A4DF5E83E1C49217719C2161DA53D49530D71B7CB903F67CCED875CAAC62545
          D8EDFA32D5EDAA767F064B217709D70817DC6CD254537CEC522B76A731B91F75
          99E4E428CBDD9E4C75116D3B90BB00698B8EAC95A4A9C6D8EA3CE6A38CC9FD4E
          97594E8BB4DCB565AACD54EDFE3CA242EE89236DD13194161DE648761EF3CDC6
          E47EAECB2CD78DB4DCB565AA1F96AAB6EDE88DA9907BA27C2E6CD1B19D34D520
          992EA35EDBD4957B1D1ECB18CD546F573DBAB74055C83D513ACAD65ABEE13475
          7AB49DE2D669EF695327F75D2E8F651E8BB8DCB565AAB98A72CFEB82AB907B62
          485B746492A6FAF15CE6B73C960968A6DA4EF5E8DE86B61DC83D21162C12A6A9
          534953FD782ED3DF90DCCFE6B18CE94C355BD5EE7DD015724F04698B8EFAA4A9
          A17A2E730B8F654C67AAFB5E5694FBDC57F01572AF3EBB37C9D6592E696AB89E
          CB9CC76319E399EA18D5A3FBC3080BB9571F618B8EBCD6A4A9E17A2E436519F3
          996AB71245B93FB5146321F7EA226DD131C8F0C1FD2C8CE2F173991C1ECB98CF
          540FAB1EDD5BA22CE45E5D842D3A8AEEA2A84CB89ECB5C398BC7321E64AA69AA
          767F1D6721F7EA216DD13183343564CF6536F058C6934C355D51EEF3A7602DE4
          5E1DA42D3A524853C3F65CE67E1ECB7892A9D6503DBA37405BC8BD3A085B74E4
          8DA045877FCF65E6F25826C8996A926ADB8ECFF11672AF9AB7842D3A52495343
          F75CE6011ECB7893A98EA66D0772F700618B8EC19F91A6FAF95CE65913725FC4
          63198F32D54255BB4FC25CC8BD2A5EB7344DE5E05EF9731913CD98AE74F98FB8
          2CA64177A6AADEB66301EE42EE953345D8A2A3A43D45657C7D2E63A219D3AFDD
          CA3A330DDA33D55EAA76BF0C7921F7CA91B6E858439AEA112ED9DB3B06E47E03
          8F65BCCB54B314E5BE6937F642EE95F1F92C3BD3549E417E438177CF6536F258
          264099EA3AF485DC2B63B86C5D9536224DF50AB712B1F5F4CB7D2C8F65BCCB54
          F717A8B6ED5888BF90BB3BD2161D3D48533DC3ED72F65EFD7277C95FD23F6616
          6CCC549B2230E4EE8AB44547D26AD254CFB8D665126ED5EEF6DA2EFFA4AD4C82
          914CB5AEAADDFBA230E4EEC6F3C245D58F34D543B6394F429D43BAE5FEB4CB74
          6730074632D583458A72DF3C0C872177677A0B5B741490A67A894B4BE55DDA4F
          EECFB9CC7736736026539DA67A741F88C490BB332D642B6A454FD2542F71F998
          71FD95BAE5BED365C26B3007DF39BA1FD494A9262BCAFDC8D5580CB93BD145D8
          A2633669AAA72C7799870D9ADD7EE80B977FD0C5CC81AD99EA3C3486DC9D10B6
          E8485A499AEA293B5C26E27EDD2777975BBA22A6C058A6BA5CD5EE7FC563C8FD
          FB485B740CA54587B73CE6F25F581B35BBFDA8CB844F600A8C65AA1F962ACABD
          66774C86DCBFCBB0B9B2D5944C9AEA352E05088E697E2E738BCB8C173203E632
          D5DB558FEEAD501972FF2E0FCBD652BEE134753ADEF81E2EB15B4D3A75842053
          CD5594FBFA5AB80CB99FC9D2A7646B299334D5735CFA7514DFE14DA78E4F9901
          83996A3BD5A37B676486DCCFA4A56C25654D254DF51CB702042FEA95FB2A97E2
          031F30032633D56C55BB0FC066C8FD74A42D3AEA93A67A8F5B0182ABB4BAFDF1
          628A0FF891A9EE7B5951EE3993F11972FF16698B8E09A4A97E68C4E549457FAD
          727FD1ED7B6426C06CA63A46F5E87E1B4243EEDFB244B68AD25B93A6FA418AF3
          748CD22AF71F537CC09F4CB55B89A2DC3B74C568C8FD6B9E14B6E81864F8E04E
          9AEA4CAA17FD3AFA537CC0A74CF5B0EAD1FD259486DCBF46D8A2A3E82E8ACAF8
          C2722FFA758C72F9875CC7F89BCE54D314E55EDC04B963F5AF385FB8869A93A6
          FA83DBADEC9D1E141F18CCF09BCF54D315EDBE05B9A3F593485B74E492A6FAC4
          E52E3372EE21F3C50772197EF3996A0DD58B9989C81D2A10B6E8481F418B0EBF
          702940F08007C507D2187D0F32D52445B9379E89DCE11FFFE83D9E343568B8F4
          525EA551EE1BDDEEE2187D0F32D5D1AA47F791C81DC42D3A067F469AEA1BD9E6
          0B10B47599F7BB197D2F32D54245B9CFFA04B983B445C70C8ACAF8C7629749B9
          509FDC5D3E6B2BEDC4E87B91A92AB7ED688BDCA18D6CED94B4A7A88C7FBC6ABC
          00416D977F420A83EF4DA6DA4BD5EECF20F7A8D347B874D690A65A98A8EED426
          F76729E62E4457A69AA5DAB66301728F36AF085B74A4F20CD2575C4ABA0FD126
          F75B29E61EF44CB515728F36C2161DDB1A91A6DA98A88ED7F6D0BD2379AACF99
          EAFE0245B98FEF8DDCA38CB445470FD2547F71FBCAE56C5D727F8D3C35F0996A
          0BE41E65842D3AB6AE264DB53351BD5F5731F7A7C8537DCF54EBAADABD0B728F
          2E970A174D3FD254BF7189DB361A2EE64E9EEA61A67AB04851EE6D2623F7A822
          6DD151409A6A6BA23A5693DCAF224FB520539DA67A74EF83DCA38AB045C70AD2
          546B13D5C69AE4BED365EE2F62E43DCC544F242BCA7DEE2BC83D9A485B74CC26
          4DB53751ADAD47EE5FB8FC7BFD63463E5099EAC3C83D9A085B7424AD244DF59F
          8B5D66E75E2D6EBFF205F2542B32D5E5AA6D3B9622F728226DD13194161D1627
          AA7A4ABAFFDAEDDB35C6DDDB4CF5C35245BBB744EE1144DAA2239934D5E64455
          4F4977B7E6D8CB19778F33D5DB552F662E45EED143D8A2E3784FB36E9F8E17AA
          47A6F3FC3C4A9E1AAA4C757FAEA2DCE74F41EE5143DAA2239334D5EA44B5584B
          A23AC4E5DFECE4A99E67AAED548FEE4B907BD410B6E8C8224DB584EB5C66E859
          1D79EA7AF2546B32D56CD5B61D4F22F768D145B852EA93A6DA9EA892A7862B53
          DDF7B2A2DD8723F7483159D8A2630269AAED896A4783792ADFA7FA91A98E51BD
          989984DCA384B045C7F1D6A4A9D6906DAE49F64EEAFD5A94A9762B5194FBA205
          C83D3A485B74A4193EB893A626804B1FD5E2BF18CB53A9F71BD04CF579E41E1D
          842D3AFE309534D51E2E77E96CFEB4B13C750283EE4FA69AA628F74DBD917B54
          585A4C9A1A02B63ACFD2ADC6F2D434C6DCA74C355DD1EEEB907B54D8225B21B9
          ED49536D22D5799AEA18CB537B31E63E65AA3514E59ED705B94703618B8E3CC3
          692A2D3A12C4A5A8D4226379EA758CB95F996A92A2DD9B22F74830A5B16C7D0C
          224DB58B4F5D12D53B0CE5A9590CB96F99EA61D54CB52F728F02C2161D459FD1
          A2C32E3E5EE13C53171E3293A72633E4FE65AA858A72DF3C0CB9879FEB852D3A
          665054C6365C6A4AFDD450FFD46C46DCBF4C55B96DC740E41E7E842D3A52DAF3
          0CD236D2CC7CA3FA9CCB12D8C188FB98A9F65294FB91AB917BD87946B634D2D7
          90A65A87CB7EAFA9E6F643759CFF67F33F60C4FDCC54B314ED3E0FB9879CEE35
          652B239534D53E5E35D347B52625216DCC5447AB5ECC5C83DCC34D2BD9BAD8F6
          0669AA85B83C90BB53C9EDB55DBE7C2D64BC7DCD54F767A8B6EDE88EDCC38CB4
          45470FD2541B71D9EEE729C9FD5E9735D09CF10E78A6DA00B98719618B8E92D5
          A4A936E2F219D35825B99FC7274C9666AA7515E5BEBE16720F2F0384ABA21F69
          AA955CE4F2A65949EE63F984C9D24CF56091A2DD3B23F7D0226DD151408B0E4B
          B3BA52E7093BAA22779772D0050CB7EF99EA34DA76207717E6C856C48A46A4A9
          9632C179C6EE5770FB519755D08CD1F63F534D56947BCE64E41E4EBA0A5B74CC
          264DB5954CE7196BA820F71B5C56C1ED8C760832D5DB907B38B942B61E925692
          A6DA8A4B37A6510A72EF4F17268B33D5E58A72EFD015B9879126C2161D4369D1
          612D8F39777178E14AED2521E9C26445A6FA61A9A2DD5F42EE6144D8A2239934
          D562529C27EDD7F292902E75E532196B2B32D5DB15E55EBC10B9878F9B648BE1
          784FB36E9FCEBE57C1A512EC55DA4B422E66ACEDC8547315EDBE05B9878E998D
          495343482FDDADF67EEA5238EE31C6DA8E4CB59D6AA63A11B9878D91B2959035
          9534D5662E769EB65562B93F40D530CB33D56C45B9379E89DCC385B445477DD2
          54BB71AE1D96272E0CE9F21F78698CB432FFAA47EEFB5E56B4FB48E41E2EDACA
          D6C176D254CB49D55B18D2ED13A65E8CB43599EA0E45B9CFFA04B987894764CB
          20BF3569AAE5B81493DAA8F713A6BC571969757EA4A96D478AA2DDDB22F71021
          6DD191498B0EDBB94EEF674C3B5DCA8232D061CA541F41EEE141D8A2A3883435
          A897EED2CF987268D411804C354D51EE35BB23F7B0206DD1419A1ADC4BF74B64
          5D988AB9720F42A69AAE68F756C83D2C7496AD80DCF6A4A981BD747F4E24F73B
          B9720F44A65A4351EEE37B23F770206DD161384DA54587C94BF78E12B71FDAC8
          957B3032D524DA7620F77FC85B740C224D0DF0A5FB7CD1C97D1457EEC1C8540F
          AB66AA03907B1810B6E828FA8C161D41BE74AF27A91AB69E2BF78064AA858A72
          6F3319B9079FAE1D64B33F83A23281BE74977463BAD0E5CAFD7246D9B64C55B9
          6DC71CE41E7C5E92CD7D4A7B9E4106838BF57563BA95C23281C9547B29CA7DEE
          2BC83DE8085B74E48D204D0D0A5B1D67709740EEC7282C139C4C354BD1EE5720
          F7A0236CD1914A9A1A189CAF5F9F4ABC76D8A10ED4720F4EA63A5AB56DC752E4
          1E6C26CA267EDB1BA4A981C1E53FD0EF4D58EE97B8F46BF98031B63053DD9F41
          DB8E48CB5DDAA2A307696A70B83C5DD3A5FBADB44F8D54A67A29720F32C2161D
          25AB495303847395C02109CB7DACF36AC86684EDCC54EB2ACA7DFE14E41E5C3E
          11B6E858439A1A24321D27F1C81D89BE72DFE4BC1A1E6484EDCC540F1629DA7D
          09720F2EC2161D19B4E80814639CA7F1977A7A630FEEC408EBE6BFF5D87D9A6A
          DB8E27917B5011B6E858D1883435507CBC4D4BC30E972BF70206583BFFA2E93B
          D56445BB0F47EE0145DAA26336696AC070DEE3EFE8B972AFCBF88637533D1FB9
          0713618B8EA495A4A90163B9E344FEE60E2D57EED731BE0632D56E7AECBE5C51
          EE8B1620F720524BD8A2A31F2D3A82C6B5CE33F9B48EC2325B195E8B33D50F4B
          15EDFE3C720F22C2161DC9A4A9C1C3B9ECEF7909C9FD5C6A0F043053FD5451EE
          9B7A23F7E0216CD1B1A2A759B74F675B1BA0504397ECB1D41E0860A6BA3F57D1
          EE2D907BE0989C439A1A1D9CDBAECD4AE4D2FDF1171CFF374A1F6274ADCE54DB
          29CA3DAF0B720F1AB7C9A63A8B3435887CB042B9BCCCB32EB7740CAEE5996AB6
          A2DD9B22F780216DD13194343590242B979769E8BC209633B69667AAFB5E56B4
          7B5FE41E2C842D3AB693A60613E727713909C83D87879001CD5477A8B6ED1886
          DC83C442598B8E74D2D480729DF37D6AF51BA9D6CB732E21C7D0DA9FA9A628DA
          7D20720F124D65B39C498B8EA052E238A137545BEE3F7659118C6CF833D5236F
          21F7E0206CD19135953435A8385786AC536DB977745E129F32B201C854D314ED
          3E0FB90786999B65735C9F3435B03CE838A38DABDD616FAEF3BFEEA90819884C
          355DD1EED720F7A0206CD1914B9A1AE043A073BFE45FAB95FB4D65600391A9D6
          A06D4744E4FEC947B2AF195AD3A223C0A43A4EEAADD594FB79CE6BA217E31A88
          4CB55B92A2DD1B20F760304F36BF834853838CF3E16D6C35E5FE8EF3E7A9B4C6
          0E48A67A5851EEEB6B21F720F067D9F416DD458B8E20F39063C78EF58F57CBED
          B53FA24F47B033D54245BB7744EE0140DAA263064565824D8642AFBDFB9DD744
          7346352899AA72DB8E49C8DD7E842D3A52DAF30C32D8F4729CD7FED5927B1DC7
          FFDFE38F31AA81C9547B29CA3D670172B71D618B8EBC11A4A901E731C7E26135
          ABD584692E45C3029FA96629DAFD32E46E3BC2161DA9A4A981C7B978D8D9D590
          FBD3744F0D7EA63A5A51EE1DBA2277BB99249BD8C16F90A6069EBAE2C790CE15
          21D35F654C0394A9EECF50B4FB3AE46E35D2161DA4A921E0F2E3D2764CAB1C17
          C504863452996AF142E46E33C2161D25AB4953C37A2FF354ED2A4B0F6CE05626
          14996A5DDA768458EED2161D6B4853C34073C7C9FDB1F0F3D4E39733A2C1CA54
          0F1629DA7D2272B7977576A6A91495F1860F4A9D66F798B04F076F650297A94E
          5394FBE699C8DD56842D3A4A1B91A6868302C78F54AB6A937D942F98C292A926
          2BDA7D2472B715618B8E1EA4A921C1F943965BAA90FBB9CEFFCAE70BA6E865AA
          1F7D82DCED44D8A2238934352C3C34D8B16C481572DF455D99D064AACB15EDDE
          16B95B89B445473F5A748406C7BABFE32BBF9739DB7959D460348398A9962ADA
          FD11E46E230365B359409A1A1E7608EE659CDFCA147DCC68063153FD5451EE35
          BB2377FBB8FA882C4DED69D6EDD3D9B65E9294F8BD8CF3174C698C653033D55C
          45BBB742EEF6216CD1319B34354C64277C2F7389F3BAB888B10C66A6DA4E51EE
          E37B2377DBF8AB6C2AB35692A68689EBF29C66F9864AE4DEDFB902344319D44C
          355BD1EE9D91BB65749F2F9BC9A1A4A9E16282D32CB74CB8DA6F334632A899EA
          BE9715ED3E00B9DB4503D93C6E274D0D198EF5459EAAE72AF75B1CD7453E0521
          839BA9EE50947B9BC9C8DD266AAD174D633E696AD8F8607062757F8FF1C83D6C
          996AB71445BBCF41EE36216CD191498B8ED0E1D829F93537B7D7737E63B59871
          8C70A63AB72B72B707618B8EACA9A4A9A1E36EC7A9BE30A147EE499D18C72067
          AA698A76BF02B95BC302618B8EFAA4A921C4F1A1731D9738B531716A1833D574
          C5B61D4D90BB2D5C269BC25CD2D430E258D57D7DED04E2D41554720F78A65A43
          F1E8BE05B95BC26E598B8EBCD6B4E808231F3BB66CF8A9A3DC4739AE8C0C0631
          E8996A92A2DD6F42EE76206CD1318834359C643ACDF6A3871CDCFEA2F3CA7890
          310C7AA67A5851EE8DA720771B5898279ABEA2BB68D1114E2ECEAF6EF5B08ECE
          D7750C61F033D54245BB2F41EE36206CD13183A2326125C369BEDF7168C1F494
          E3CAE8C508063F53556DDB31EB7AE4EE3F7DED4C537906E91F1755F335641DDE
          41863753EDA568F7E1C8DD7786C95A74E48D204D0D2FD52B3073B6F3074CCB19
          BF5064AA598A763F1FB9FB8DB0454721696A88A951ADA3BBF38DFB607AA78623
          531DAD28F7450B90BBBFBC256BD131F833D2D4301FFE4AAA71EBFE62B173490A
          862F1C99EAFE0C45BB3F8FDCFD651E692A7C1FE70BD7FB4F77FBA12F9C3B7351
          0F924C35046D3B4220F76B64D356D29E3435DC47F7AD4ED33EFFF48E4C37C468
          AF17F24CB5AEA2DD5B20771F91B6E858439A1A721C6B10C41A9EF60C721307F7
          B067AA078BD4E49ED705B9FB87B045472A4565A279742F7EFA9B4B99961CDCC3
          9FA94E536DDB81DC7DE373598B8E6D8D4853237A749F5FAFD252BF1CDC4396A9
          262BDABD0F72F78B8EB219EB419A1A81A3BB733B9E771EFFEAC2DDF9A54C2C9B
          8123533DBD6DC730E4EE0FC2161D5B5793A6468031CEB33FB62254FDB1CB03DA
          2CDEB8872C535DAE68F78791BB2F2C58249BAF7EB4E888042EFF49BE6B43EDFE
          6E4BA32EA316B64CB5544DEE4FBD85DCFD40D8A2A38034351A5CEBD28EA77893
          EB7FD35155267499EAA78A47F796C8DD07766F124DD68A9E66DD3E9D6D690B09
          577DA52D760833D55C45BBBF8EDCBD47D8A26336696A54B83CC167CEC90C5908
          33D5768A729F3F05B97B4D17598B8EA495A4A99121B12F144BAF63C4C298A966
          2BDABD0172F71A618B8EA1A4A9110AE52624B2349A3160A1CC54F7BDACD8B6E3
          73E4EE2DC2161DC9A4A991CA5413782B914B9A1AD24C7587E2D1BD2372F71461
          8B8E7CD2D46851FD67CEA5D7325A61CD545314ED3E09B97B89B04547262D3A22
          F65FF6DBABBB34E8BF44A61AA6B61DC195BBB04547D654D2D4887171112F65C8
          54D314ED7E1972F70E618B8EFAA4A991634CB55E55255DCE488539534D57937B
          87DDC8DD2B5E974DD104D2D408529D9770A517314EA1CE546B281EDDD721778F
          98226BD191DE9A34358A54A395662F4629DC996AB724C5B61D0B91BB37085B74
          A491A6469287AA7CEDCE0BF7D067AA87158FEE4D91BB277C3E4B343D4577D1A2
          239A5C5EC55338BA2F4520532D54B47B5FE4EE05C3ED4C5339B807D4EEB83D0A
          99AA6ADB8ECD3391BB79842D3A72DBF30C32BA76AFE46686E64BD1C8547B29DA
          7D2072378EB04547FA08B36EBF914D68331FA4BA5580A63F87F5E8C9540F66A9
          C9FDA3AB91BB699E974DCD20D2D468B37C85E3FB76DE40DA8FA64C75B4E2D17D
          1E72374C6F598B8EC19F91A6469C6BBFDFB521BF909EA941404FA6BA3F43D1EE
          7F46EE6669219B9719A4A9D07CEB996F9793EF664C028125996ACDEEC8DD24C2
          161D25A4A91097448D826F6A0067A5A1F6C0A02953ADAB68F756C8DD244364B3
          B28634154EF2F18375330BD3B29B53E0375068CA548BD4E43EBE167237471FD9
          A4A452540620C068CA54A7291EDD3B2377630C9B2B9A926D6F90A60204194D99
          6AB2A2DD072077533C2C9B911EA4A900C18E4BECC85473262377332C7D4A3421
          5B5793A602041B4D99EA7245BBCF41EE6668299B8F7EB4E800083AFFA42753DD
          A6D8B6A32B723781B0454701692A40E0D194A97EAA78747F09B91B40D8A26345
          235A7400041F3D996AB75C35B9173741EEFA59229B8CD9A4A900214053A6DA4E
          F1E8BE05B96BE749598B8EA495A4A900614053A69AAD68F79B90BB6E842D3A86
          92A60284033D99EABE156A726F3C05B9EBE57CD9442493A60284044D99EA0EC5
          A3FB48E4AE15618B8EE33D495301C282A64C35454DEEB3AE47EE3A11B6E8C8A4
          45074068B024531D8EDC35D27BBC6812B24853014284A64C354DD1EECF20777D
          085B74D4274D0508139A32D5F4D0B7ED088CDC852D3A2690A602840A4D99EA62
          C5A3FBF3C85D176D64696A6BD2548070A1A9F6EF56C5B61DBD91BB1EFA90A602
          40059A32D5C38A47F716C85D0BAFC85A74FC612A2D3A00C286A64CB550D1EE5D
          90BB0E1EB6334DE5E00EE0037A3255D5B61D6D2623777596168B063FB73DCF20
          01C287A64CB597A2DDFB207775642D3AF20CA7A937B2C9007CE19FF57CA79AA5
          26F7B9AF2077552E950DFD20D2548050A229531DAD78747F18B92B226CD151F4
          19692A4038D193A9EECF5093FB534B91BB1AC2161D33485301C28A1D996A4BE4
          AE84B045470A692A4068D194A9D655B4FBEBC85D05598B8EF435A4A9701A17EF
          A851B7595A6A4672F2F6DCDC09C9C919A969D9CB7B8DB9AE1343134C3465AA45
          6A729F3F05B9CB794636E8A9149581935C3E6679E18424D7D63BE949B919CD6A
          5CC738050D4D99EA34C5A37B03E42EA6BBAC45C7B6374853E1B1C5D9DBABF9DA
          6DF084B45EAF3262414253A69AACD8B6E373E42E45D8A2A307696AD4C55EA3B0
          24D1AAAE4919CD117C70B023531D8EDC85085B7494AC264D8D3277674F382EDC
          AB29990F720D1F0C3465AACB15ED3E09B9CB10B6E8E8478B8EE87251A66231D7
          ACC231F83D08E8C9540F6E535B2E8B162077095D64C39D419A1A552ECE4E8A69
          A028ED6EC6D27A3465AA9F2A2E96CB90BB80C9B2161D2B1AD1A223927CDC2B39
          2FA68B94E58F31A296A329539DA0B65236ED46EE89236CD1319B34359A87F6AC
          98564A0BAF6554ED464FA6DA4E71A1AC43EE09236CD191B49234357ADC9DBA22
          A69DBCE4318CACCD68CA54B31597C942E49E28C2161D43495323C7A7C93143A4
          D46074439FA9EE533C183445EE09D244D6A223993415B5A37732D544D8A1B846
          FA22F7C4D8221AE6E33D4953A3C5B546D57E52EF3B18E57067AADD52D456C8E6
          61C83D116EB2324DA54587655C9E961E334F01D12A996A650C44EE0930A5B1EC
          0394A9A4A911E2E3E583639E703C8D8791A1CE54D3D4D6C791AB917BF511B6E8
          204D8D1217A5C43C23AB17E31DE64C55F13F00E721F76A73BDAC45C776D2546E
          644C914C61E01067AA8B1557C75F917B7591B5E8C86F4D9A1A191667C53C665B
          5D463DBC99AA6245A29A53907BF510B6E8C8244D8D0A1FA4E5C5BC279992C0A1
          CD540F2BAE8D56C8BD5A74AF292BF764384DA545873DB7ED5B63BE50C4CD7B68
          33D542B5A5B1BE1672AF0EAD64C35B9FA2321161F9F1985F147ECCF087335355
          6DDBD119B95703618B8EDCF63C838C048F6508B7DF6F36AF7AE79D9663C78E7D
          E79D9CC6EBA59F3491AB863453EDA568F701C8BD6A3ACBC6D6709A7A239BC80E
          EE4EF84AA6B8E6B186BFBD77C35FCACEE08EA3CFFE7863C79CDF247C35433D82
          7066AA071523FA9CC9C8BD2A06C8867610696A24A89158E79C9A1D7FFBE21D65
          9570E5861B76E6FC22A12A80CD98855066AAA3158FEEB721F72A10B6E828BA8B
          343512D7ED09BC92B9B9CE0DF5CAAAC55FEEECBF2A918B77FAF0853153DD9FA1
          26F70E5D917BE5CC910DEC0CD2D428DCAE56FF49C3AA8D979425C4D19F0EA976
          21D20997331764AADFE325E45E295D652D3A52485323C043D5AD00D978E38632
          01476FCDA9E63F60EBC5CC460833D5E66A722F6E82DC2BE30A59339411A4A9E1
          E7F2DC6A2D86A78EDD79659994177756EFAD561285224398A9762B52B3FB16E4
          5E09C2161DA91495093FAF56AB4ED8A68647CB94A87DEBFC6A5512BB8819095F
          A63A4DF1626622727747D6A263DBFF63EFDE9FAB2812058E2759104C80808457
          EE5E4214049210501EA1100328081B4190F8E221A5C843B084E5A1F181B582A5
          1029208A4059A58185D26597E26145C0E5B202A588FCA2BFF813553118019587
          A5FCB475EE99ED1BF4EE8A327392E9E939D3D3FDFDFC01F7E6CC0CDF1DBB7BBA
          3F6536D5FCB6B76409E48BCBBE4B0456F94A4B46676AA8BB8173AA018F7EC95B
          4EDCBD481ED13182D954DAFE53DA9F4928D1D892BC5377E654AFF10071F7B05C
          EE888E9DEF329B6AFC78FBCE661F833DFD14A5FDA7BC3F9A47DD6D9C53DD1A2C
          EE55DF1277770FC85DD05D1CD1617CDB9B1D6FAF3D919350EA8D7E7B9A1D7767
          56D5B839D58663C1EA5E42DC5D7D2B7744473EB3A9A65B3BA7B987E0E47B8D09
          D5722E35BB66861591C6CDA95E083830B383B8BB2991BA98759D39A2C3F4FFE2
          6EEED3C13DCB2A136158DCDCD8CC4ECE56356E4E754EB0B8670E27EED77A4DEE
          62AE6036D574CD1D5FFCCFEB1321697D5D73DFAAB205B06973AA03CC38B643AB
          B84B1ED191B59DD954C33533C775E6A6CA44781637F3C97407EE8F6973AABD83
          C57DF628E2FE5B924774DCC16CAAE1A60F4BF900AC7E2F11AA9C2F523F805BB9
          4386CDA99EDA16ACEE3D88FB6F481ED1D186D954C33D9CFA93F092D6899035DE
          90FA58F6E9DC23C3E6548F9B706C874E71973BA263DB486653CDF655CA4590C5
          FD1269B038E58B470D67331936A75A961D2CEE7DE713F7AB8DD5723695233A22
          D721E5A9C4AF24D2E28394AB66B29954654EF5574613F7ABCCEF2F75110B994D
          355CCA4D580FBD9748939C93A9FE8E5CEE936173AAB9C1E2FEE287C4FD179247
          744C6436D56C332B52DCFDD72F26D2E6BB9EA906873673A70C9B53ED18ACEE1B
          88FB7F0C1D2875093F6236D56C67530D7EF6CF49A4D11B6B52FD17245FAA1A36
          A7BA39E0B11D6F12F77FDB2F75053B329B6AB8542B8E4FB64EA45563AAEF99DA
          70AFCC9A53AD3F12ACEE3D89FBFF933CA2A380D954B3BD70599FB637D57D508A
          67B13B77CBAC39D57301E7542711F79FC91DD151F8044774983D39966250E6E3
          B4B7BDA9EE2946666A8E72BFCC9A53AD0E16F7D50B88FB1573E52E5F7B3695B1
          7650E6F59C44042A67A5D89B94FB65D69C6AD0633BB61077217D4447D1FFB20C
          D2680FD7792F35BB3E1189EF52AC88CCE08E9935A75A1EF0D88EB788BBEC111D
          C53786DBF6C7F83712B109DEDF2EDD968848CEEBDE3BD8F129935973AA6585C1
          EA3E95B88B6FCF485DBAF1CCA69A2DC562B457129179FFA0E75FD59B7B66D69C
          EA430107666E25EE724774D43CCF6CAAD1CE7AAF453B9088D062CF955DC758EC
          6ED69C6A7D7EB0B86F5A677BDC258FE858C26CAAD9BC37719FD51865DC13FDD8
          DA9D39D596D96D79DC258FE8C86636D56C5F790E78E6B58E36EE8D9E1B1174E4
          BC6CC3E654BB078BFBC15176C75DEE888EE2FB994D359BE732C8DA3F272296E3
          7936D304EE9B6173AA35F13DB623FAB8B7923BA2A3039BCA986D8AE7BFAA7E89
          C83DEB39EC7E813B67D69CEA93C1E25E3CC4E6B8CB1DD151F129B3A9662BF0DC
          75A032FAB8274EB0C58C2D73AA6D82D5BDABC5711FCB6C2A5C6CF47A713FF381
          066D4FB45EEDF564BEC0BD634EF56A63AC8DBBE4111D3BDF6536D56CDD341E94
          B9E2253621B0654E756BB0B81FBAC7D6B83F2577C176853C9BFA25FF30A275D6
          6BA94C66A51E714F786D10398CF354F51998A95711F78663C1EA3ECED2B84B1E
          D1C16CAAE93C57A0BDA449DB1B73F670E29E2573AA1782C5FDFCE776C65DEE88
          8EBACECCA61ACE6BABDF4B8D095DDCE435D7BF91DB67D89CEA9C60755F6965DC
          D7CB1DD13182D954C3DDEEF51274BD366D4F547ACDA9B6E5FE1936A73A20E09C
          EA0C1BE3DE55EA5265319B6A3AAF63124E2434F2A8D7B7D3DC3F7D3428A97BB7
          80C7760CB72FEE924774DC11F28B3BB3A951DBE83185B52F47A7B8377EECF180
          DECE1DD4C65225713FB52D58DD3B5917F7E587A42ED40466534DD7D6E3D6DF90
          D0CA628F3FB39A3BA80F35CB218F078BFBBE51B6C55DEE888EBA91E1B6FD8FFC
          7B889CC70456558E5E716F7CC76333EA5BB885FACCA9AAD962263B865BCC4418
          77C9233A56704487E98E16C760C4FD8AE7386F4F7F8B749853ADED6257DC574A
          5DA5C2EDCCA69ACEE393C0DAEB758B7B651E5FA9EAFFEAAEE44BA66441FC9643
          4617F76FE42ED24466538D57E47EEB4B12DAF158EB7E6C2D37511F6ABE643AD5
          3158DDE75914F7057247747CC46CAAB5A332CFEA17F79C338CCBE8FFEAAEE44B
          A6E4E660713F6C51DCE58EE818C66CAAF93CD6CAACAED42FEE8935AC97B1E5D5
          BDEC48B0BA8FB526EEADF6495DA0026653CD3741EBED207FED2F1EDFD97117CD
          7B753F172CEE25D6C45DEE888EC227D854C6FC7F8A15EE37FFA28E71F79A5265
          57779DA8D98420591DACEE5D2C89FB3CB9CBD39E4D65CC37DDFDDEDFAD63DB13
          8D37B8FFB5DDB88FE6BDBA073CB663B01D71973CA2A388D9540B781C8CFDB296
          714FDCC6697BF6BCBA97078A7BD56756C45DEE888E8E37329B6A818FDCEF7E8E
          9E714FB8EF0D59C37D34F0D5BDA13050DD7BD91077C9233AC6332863C3BF43F7
          4DC3FA6BDA76AFA3B21974D78A9A0533C987826D0E39DF82B86F90BA3435CFF3
          6DAA053CB6723FA06BDC9F75FF7BB772270D7C75AFCF0F54F7EFCD8FFBB76F4B
          5D99257C9B6A038FBD076ED335EECF54B103813DAFEEC1E654A79A1FF7A9CCA6
          C25307D7BBBFA752D7B8271E74FD837772274D7C754F760F12F733434D8FFB10
          A9EB527C3F4BDCADE0BEB76A4F6DDB9E38E0FA076F63DB5F235FDDCB02CDA93E
          657ADCE57683AC6636D58E37ACBAF87C9EFAB39798518DC583A56673C8E49341
          E2DED5F0B8CBBDB857FCC06CAA1566BADFFFBFE81BF7EFDC4F79EFCEBDD48B9A
          7DDD93F56D82D4FD5BB3E35EA2E36CEA633CFB9AF0F84EA4B5BE714FB89FC754
          C0BDD4CBE3490DE6547B191DF7F552D76467291B86D9C1FD9CF9D51AB7DD6367
          4896CBE8A69DA2BA6F0D10F7BB8D8EBBDC52995DCCA65AC27D73A6593AC7BD9F
          EB9F9CCDBDD4CC5245716FA80850F75606C7FDCD62992B92CF6CAA2DDC371F38
          D0A871DC17B301413C9C5654F70B31592F93EEB8EF95B920973B339B6A0BF713
          11EED3B8ED8DEFBB3FB51BB9999A51B47D58B22C5B3EEE3DCD8DFB50A9333A0A
          984DB5C6B1B82D9649249E717F6A67723335A3EA43A620C7769C5F6E6CDCA50E
          D70BFB880E0665F47136460775FCC721D7BFF9387753378A5643063AB6E37B53
          E33E7FB586CB20994DD5C85DEE9F27576A1DF7931C921D0FAA564306590E79AF
          A971FF44E66A64BDCB6CAA35DCF7841CA875DB13B35CFFE8B6DC4DEDA85A0D99
          EC2D1DF74DA6C6BDA78667EB319BAA93CDAE8F40A6DE71BFCEF58FEECDDD3476
          4A35C8AB7B2B33E3DE4A661D6456B8DF2FB1D3AF56CA63747E6A33E7A8F289AA
          86544DA926DB4AC77DB49971EFA4E18B3B3BFDC620EEB3F47E7377FF8A2997BB
          A99F3EAAE27E4A3AEE3DCC8C7BA67E23EE2C83D44BDB38C67D99FB3EA6DC4D83
          A75493DD64E39E6964DC8768B85486D9D438C47D501CE3DE81BBA9A1D3D18FBA
          7F6662DCF7CBAC71DFCE32489B748B61DC1BEF63E730FBA6549305B2719F6160
          DC870F94B8102B580649DC757F7327EEF1714B32F257F77106C6FD5689EBB0ED
          079641322CA37BDC5F6658263E942D754FCA9EDAD1D5C0B8CB8CCA8CE7C59DB8
          3BCEA5462654A1DBB8CC4392719F6D60DC0F495C871B7971B74B777396421277
          3DC765EA55C55DFAACEC69C6C57DACC45528E2FB25CBB8AF737F50EFB8F31153
          9C28DB3D2CB95932EE738D8BFB38ED3E60E2FB25FD64B83E071FEBBD5AE604DB
          0FC4C8D264D41F328D332EEE125F30556CE7FB25CBBCEAFA201CD2FBCDFD92EB
          1FBD95BBA927655B102473F53EB0236D71FF51BBE9545EDC3534D3F541F8BBDE
          71FFC2F58FEECEDD347D5C66805CDC5F342DEEA3252EC2FD61B6FD8F3CE51ADA
          E8FE24B4D67AB9CCEBAE7FF366EEA6E9E332653572751F6A58DC07FBBF044778
          71B7CF65D747E136ADC7DCCFB8FECD2F7033F5A46EBD4CB25C2EEE630D8BBBC4
          E7A92B7871B78FFBF2B247758EFB45F7C7F728375353EDA2FE4A75B459715F2F
          710946B263987D8A5C1F859B748EFB5F5DFFE43AEEA5AED47DC7547F442AEE1B
          CC8ABBC4C9D8D9EC1866A1FCF8ED3FF037F74145EEA5B6E332CAE29ECC908AFB
          4AB3E23ED5FF1518C18BBB85DC0FA7ECAF73DCDD97B9B7E15E6AAB21E271994C
          B3E27E48AF51195EDC75E5BEFFC0F937348EFBDDEC3E10337DD48DCB64C9C4FD
          FC7C93E2FE96666B657871D795FB574CCE1FF46D7BE51EBE618A19758B2125D7
          CBB43229EE73FDFFFE025EDC6DB4D1FD10F5BFE91BF73FB83FBFD3B997FA0EBA
          D747FC1DD33C93E27EAFFFDFBF8B17772B65C56D46D57D3777670AB7525FEA16
          4396FD8F4CDCC79814F7BEBE7FFEB1EDAC71B7D204D7C7214FDFB8AF713FD89D
          3B69C5A0BBDC697B5B0C8AFBFC2ADF3F7F022FEE76725F2EE35CD436EE87582C
          63F3A0FB933271EF6150DC3FD76A21242FEE3AF35839ACEDA0FB7B0E1BFEC66F
          D03D19ED62C89506C57DA1566730B1AB8CCE36768CD7614CFD984F8DA1D3EAEA
          2E731E537F83E2FE80FFADDC4BD9C7DD52D9AE0F44D533B15AE57EF92B6EA4CE
          D46DFB9BDCAAEDA6BFE9897B894E43EEBCB8EBCDE30484C57AB63DA7D6FD8048
          EEA3D6D46D2F23774EF67073E2BE5AA321775EDC35E7F159C89A582D84CCE53E
          6AEDF18807DD5B1913F7E13AAD72E7C55D7353DC07DDF77DA765DC4F3A9CD411
          47EA3E632A1B2611F721C6C4FD4DDFBFBDF889B0DA5EFA250FB6E6DC77FDD573
          4FF7EB3DBED238CB6DD49BC219D56A89B8EF3026EE337CFFF69DA1BDB8FF9EE7
          5A771E2BDD1FD4F110A6030EABDC6349E18CEA6689B8CF3526EEBB7DFFF60EA1
          BDB8F30193F66EF77828DED72FEE6F78EC76DA96BBA8398533AAE724E23ED998
          B8EFD7673E952DC362C063E5F09DFAC5FD158F51C5BBB8899A5B1AED8C6A2763
          E2BED2F76FBF839D072CE6B118F2A07E53AA5FB8FFA573B887BA53F88D6A99AE
          07EDA525EE99BE7F7B67761E605CE61A2FEBD6F6DB3CFE50F672D79FBAE532C9
          369A6E2E9396B8CFF6BD2524EB20ADE671BC4D5EA56671BFE4FE775E3ECA2DD4
          DEE948BF519D6A4ADC976B7336361F30C583C77A19E739BDDAFE7E316B65624B
          DD96EEC9E39AEE1C968EB84FD366B10CEB20E361A6473557EBF5EABEC6E3F12D
          E70EEAAF4FA4CB650E9B12F779BE7FFA0AD641DAED238F07E31F5A6DF6EBF13F
          41357CC114033747BA5CA6BF2971F77F806A7BD641DACDEBD8E13C9D16CCCCF2
          3AFD97FB6757DC1BFCC73DD394B8FBFF866917EB20ED76D66B93EC7EFAB4FD25
          8F3FB1E34CEE5F0C28DC3AACDE7FDCF34C89FB16DF3F7D24D3A996F33A99729F
          36C7ED55BEE3309D4ADC7FB2D377E10E9A12F77B7DFFF4ED4CA75AEEE8368F47
          E3922E715FE6F5F07206532C28FC8A29996B6FDC07EBB1CC9DE9D438F1DC6AEF
          593DDA7ED1EBE38D6CEE5D3C288C7B5B7BE3EEFB1CA62CA653AD37B3A3C7C3B1
          BAB51671EFE9B00E32DE147EA29AE17F4F7353E2DED5EF2F2FE2EB544CF07A3A
          4EE8D0F6E73C5F4C6EE1D6C54399BAB85FF03FA36A4ADCFB6B71806A298F73AC
          BC50ECF578BCA4C1A0CC41AF3F8ECD7EE3A221D263544D89BBEF13543B309D8A
          7FE57B3D1E8772225F29F385D7DF768417770BE33E80B8B7D87816B9E35F333D
          CFA69CD51871DCFB398CB81377E2DE64A0DF1F5EC02277A45830E32C8BB6ED7F
          AD75582A43DC03ED3F604ADC0FEAB0B50CA332B173D731AF07A436D261F78B2F
          3AAC7127EEC4FD8AF33A1CB2C7A84CFC74F35E267C7D746D7FE6A4F74A00EE99
          95713F656FDC7DFFF0258CCAA0C9DA239E8FC8C791AD766F5CE3F947D53DCC3D
          8B8F32E26E48DC199589A3CDDECFC8836F4414F71BBCFFA6DEDC313BE3FEB5BD
          1F31E910774665626982F74332289A25332FA7F8B07A2D372C46EAA38CFB41E2
          CEA88CE58E56783F257746F2656AADF7CBD8ABDCAF1851B97118718F32EE8CCA
          C454F7148FC90DE96FFBA3DE6D7772B95BC4BDA55613774665ACD746A7BA2F4E
          B1F02B6B23372B4E54EEE77E8AB8471777F69589AD876B523C2877A675DCBDF1
          D1146D2F66897BBC2C8D34EEC69CA15A1BF9474CECF61B5FE5A99E94EB2AD339
          979AEA41E6E0D498E9938CF2232663E2EEFB0B55E5DB0FB0DB6F8C55A77A542E
          A5EDC4ECC6032977A966C3B098591469DC0F9B12F717FDFEF25CD5A3325FF22C
          C7D7DA9427549ECC49CFD0CC338352FD15157CBE1437ED92516E1C56C296BF8A
          FC914739CE6656A43C47FEBD74B43DE7EE944F6C0677296ECA228DFB6053E2BE
          29EAC33A5808196FA94F31DBF75CF86DFF735ECA3F814F536347E54A4889C33A
          F65B7B1253110B2171B5DEA99F97EB9E09B9EDCB52EF7DD78601F7D851B95846
          E298BD71A6C4FDB0DF5F7E848590F8D56B569BD40FCCC71F843A2473A999C795
          15EEF1A372B18CC401D99D4C897B89DF5F5EC14248FCCAC6ECD44FCCDF6F0A6F
          5A7571330B020A67727FE247E57C6AB2ADEFB83F654ADC07FBFEE9EF32E48E5F
          399AD5CC23F3C5FBE1A4BDF575CDFC3F3E7681BB1343F52AE35EE0BB700B4D89
          FB06DF3F7D2443EEF8B59935CD3C337FEF17C626C0CF35B78E77180B65E248E5
          E603C96491EFC2DD6A4ADCB7F8FEE9BB1872C76FBC7AACB9A726F359D569FFE0
          C1661FD5B6DC19EB87DCEBFD9FD5D1C594B84FF6FDD3DBB3CA1DBFB5795BB3CF
          4D4FA513AB39279ADF38632BF725964EAB8C7B83FFB88F3225EE637CFFF4150C
          B943A6EEB5D75D5436D87E60B643DB0DA57495BBCCF9D8C34D89FB0EDF3FBD03
          1BCB40AAEECE1935796F7DA0253B2275E39E302A23F3816A5ACEEA484BDCBBF8
          FEED4ABF62FA9267D9A6BA3BB56B6E0B9AF6EBEFDCE3D0768335288DFBD3BE03
          97694CDC3FF3FDDB6B38610F6E5EAD68D1F373F7A301BE59AD7CE9528B36A91E
          C65C6A5CA95D2B23B1CCBDA7317117E77DFFF81FF884096E6E2F6CE17FF79E90
          DC4FECE28116EE73B7AD9CBB11578BD4C67D829EFB86A527EE7951AE85643ED5
          2833B35AFA0CBD7EE00FBECB7ED3C9167F45CDC94BB1758BD22F9892F51DF5DC
          5A263D71EF1BE5417BCCA79A65CA9C963F46AF9F58DCE2C33C2AFF7AC0C70E77
          59EC39105F6AA753250ED973269B13F7A9BE7F7C35DFA7C2EBBDAB839F27E9ED
          7FDEB038A7B9B07FF752BF9E7BFCFC5F9D3385FB10DF07A82C19F16219679239
          711FE7FBC767F37D2A3C751BE6F371CA9B75C3A3EFB576DB5BACF507AFF4BBF4
          8EDF537EABD772137871FFB7E9FEE33EC49CB8FBFF4475D87616CBC0D3F14247
          C29E777AAE397160D9B27FDC77DF7DCB96F5BB7350CFFE0365FEEF6CEBCE1DE0
          C5FD17B9FE1FA1A1E6C4FD7BFFBFFE7E16CBC0DBD1394E64B25EE0FAF3E27ED5
          7CEA36DF8FD06C614EDCFD7FC5E48C60B10C5238DBBB63446DCFE7688E587BBC
          5E71DC25361FD86450DC974BFC1362B10C527A352B8AB4D79473E5E3AD9DE2B6
          4B1CA0EA941814777128BA6F54592C63AA8DD5E96FFB477771DDE3ED66D56D97
          F83ED5B9D7A4B81FF6FFFB4712773463739A5FDE2BDA721276CCDD52AF3CEE12
          0FE16493E2BED7FFEF57F51913CFB3C1BE2AB89CCED1F6A35CF1B83BADBCED12
          9F30393B4C8AFB6EFFBF7F02CBDCD1BC178AD295F6ACCD5CEDD8EBA3BCED3243
          EECE5B26C5FD3589FF047E9765EE6881F2B48CCD54743BCBA566C0DD4537FF0F
          53D57C93E2FE88C43F27357B8771C69EF1D676AB083BEDC372199131C0E3EA07
          DC9365FE57B93BFD8549719FBFCFFF1528E01B26B4CCD18250F3DE319F5DC24C
          A0FCD354C955EECE54A3E22EBAFABF023BF986092D3525BCBC937653DADE1042
          DB9319128FD416B3E22EB15C46CD6248E26ECBDB7BEF9A30D25E574DDA0D69FB
          E930DA9EDC29F1504D322BEE93252EC108E20E3F63EF6D8FA84E7B616FB6F6A5
          EDAA47659C6966C57DACC4252862F701F8B3395FE1BAF7E239DD592143DB53DB
          2CF160A56BB14CBAE27E8FCCBFAECEC41D7E07DFB7EE54F5D2FE305793F1F6E6
          644B3C5A7D8559711799118DCB1077EBDC5E107878A6B07A3A1B0D18E4F1B290
          DA2E352AD3C3B4B80F96B8083B893BA4BCD03B5BBEEC59B9C729BB5196D687D4
          F664B9CC03B6DBB4B8EF96B90A371277483ADA3D5FE2B8A6636DB6B23AC6347D
          C24A7BB25EEA44B0B1A6C55D664655C5774CC4DD6233DB7638D2F2333DB2F2B7
          5E6006D5BCE1F676A1B55DE6686CC7A95D605ADC17BC2D7332C276E28E60A61C
          DF5A5DD4CC12F8BAEC0EDD36B3C100C3ED7E15C8C47D93302DEEA2AFCC75684F
          DCA124F1B7976FCDCD9F939D55F1CB9B7C4561F69C09D5BDCB5F25EB0CC9A46D
          B7DF34CEA7A62FEEF7CA5C8739C41DCAFF3B7D4A93B55C071B5EDB1BC26C7BF2
          B854DC9F322FEE93A42EC448E20E40EA7FC5437D6D4F26EBE5369B1E625EDC47
          495D88F16C3F0040C2D2B270DB9E3C2795B47DEBCC8BBBD4674CCEB6E7893B00
          DF2332ED92619B2015F79EC2C0B8F790BA1423883B009F23328B424FBBDCD7A9
          E9DBEF37AD711F2DF72178C0D3F638AC03B06EB0BD3EFCB6277BCBC57D878971
          6F25772D9670CC1E001F692F4B43DA25D7413AB5CB4D8CBBDCA0BB93C501D900
          F44ABBD4C1D869DD1232BD71DF2F7735827DC854CAE30ED833D65E9F9EB4CBBE
          B83BE3CC8CFB277257636769A0BAF3C0037658DA2E99366D1DED87DCD319F70F
          6BA37875FF1DCF3CC0788C1E2FEE55CBCD8CBBE82AB9BFF6763E510590A2EC37
          A7F1A53DC05219E7B03034EE9D9C0816CCB0D01D307C386651597AD32EBBC6DD
          713A991AF7F59217A4E60916BA03D0A3EC4DAA65E33EC4D4B88B3CC92B52C042
          7700D78EC6445176D95D659ABC288C8DFB5EC94BB2AD336B21015CEDF19BDBD5
          272351267D426F0F73E3FEBDEC35C967B90C80A85FD97FF6B4F4E1EB93CC8DFB
          F0D9B217E50E96CB00F8E98D3DCAB037F9BA4E3663E7EF3137EE6290EC5509B0
          1C92E53280212FEC4BFBB48B36EC419641A675BBDFF4C77DA1F46559C18C2A60
          F3FBBA0E5DBF628074C4D278C25E04715F5E257B592E4B1FB857FA25FF3280D8
          BEAD5FC9FAE9FAA42ECA764AB7BD7894C9711753A52F4C512983EE8045515F7A
          739F459ABCAD5FA55CFEC5BDAB303AEEF2E332F2673231E80EC4A4E83F25BDA9
          E90DDA453DE8B7A94D269B1DF705FBA4AF4C9DECC0CC7FF701A0A5453F6BD7AE
          29E765F549ED9515C9B7BDB895D9711783E5AF4D91E4897BA565490008AE7B80
          17F7C3C2F0B8BF16E0E2C8AE98B999671240A42B651C67B4E9715F3750FEE274
          BC5FF21B551E4A00819D2E0CD0F6AA0F4D8FBBD810E0F2143E2F778E6A3D8F25
          8080EA7383BCB80F12C6C7BD4B90EB3341EED5BD0FCF2580808E076997F38DF9
          71177D835CA0118CCB0088C2AA406D3FB4CE82B84F0E72853AEE625C0640FA95
          65058AFB166141DC875605B944359D592F0320DDEA3B046A7BED8F36C45DF408
          7491B265F687FC2F9E4D0001B40D942DA7445811F721C1AE92CCC11DA50D3C9C
          00A43D19AC5A114CA74612F76053AA72DF32DDC2D30940D680806D5F3DDF92B8
          8F0E769D8ADB4BEC2FC3E30940D2AA8A80719F2C2C89FB8281C12ED4B63B58EA
          0E205DBE3E12B0ED0397DB1277F140C04B75EC46FFFBFEF284029051F651C060
          390F086BE23EEA7CC06B55E37FFBDF763CA30024DA9E1FB4ED6746D913F780AB
          21AFEC32D399AF5401842FD88E323FD92B2C8A7B97C0972BCB6FDD4B4FF39802
          F0FBDEDE3B70ACCEFF6853DC4549F0BA7FCAAB3B8090DFDBB73A717D718F2AEE
          639DB4BFBBF32113009F14B4BDEA47BBE22E7A06BF663B3BF3EA0E40EFF77667
          83B02CEEF3145CB4AC918CBA03088B82F176C73938D4B6B88BC30A2E5BE1485E
          DD0184D4F66A0591727A09EBE23E56C575ABF175AA6A296BDD01B45043BE8A46
          AD1E6E5FDC152C9869523791CF5401A877AA4845A19C49C2C2B8AF2F5671E986
          2D61871900AAAD2A54D2F695C2C6B807FF4CF567B97E3687E4BC3D00CD3BB74D
          499DAABEB533EEADF6A9A97B9B27D8D71D80429BD5B429C2D9D468E32EB628BA
          804746B21C12802A4A96405EB1699DAD715F9EA7E812564CE43455006A9C9AA3
          284C6F0F11B6C65D4C52740D9DE282D296D6FD669E5D00DECE55A8EA5227616F
          DCC54A5557D1296AE95E048F95F1F402F0505FAE2C4A5DE7DB1CF747AA945DC8
          9A890CCC00083824D3465992F64D1336C75DF472D419FF040333000278689BBA
          208D1676C77D7E578575CF6AD966048FB1F72F806B35F45698A31EC2F2B88B69
          550A2F6771C176066600C8BDB6172A8C51FF05D6C75D4C7654DAB9AB25757F9C
          E7184078AFEDCEEC6982B82B5C31F3F3C8FBF36C0F09C09FFA272B5466A87886
          20EE428C3AA4B6EE35ED5BB0C70CEB2101FCC7AA7CB515EA2588FB15F36AD55E
          57A7E84686DD01B4545977C5091A2C88BBDA3D667EF96FA2F19DD9410C40CB46
          646A1407E8EE05C4FDDFEB2157AAAEBB53B7A29945EFA56CED0E20991C90ADBA
          3EAB4709E2FE6F1F662AAFBB53B324F5B2C8C7985405ACA77AB0BDC98BDF0AE2
          FE8BCF0FAAAFBB93B5E45D26550178A7BD5A7D77F60D11C4FD6ADFD43AA1E43D
          D5DBFBEF39960920ED6A9DDF2188FBAFED7642513322C5B2F7DFF1780396AA1F
          30218CE2D42E14C4FDB7C6855377A76EFC4896CC00B85AD993D9A1E4A6768C20
          EED71A1452DD9DE23613BD8EF258CA530E58E774795D48B1F99320EE2E861F76
          4253B8A233DBFF02687A693F571D5A689E12C4DDD53D7DC3ABBBD3F1A3F6DB59
          EE0E586E5579856349DB758ABBF8B0BF13A6BA0E775CBB36B27411CF3B608753
          D3B3430C8C56E3EDBAC55D7CB6C90957458789EF5277C03EF5ABA617851A97F3
          7305714FA155A613B66D6D967C4ADD01ABC6D907941F09392C553304718FF6DD
          FDE78F9BC64FDCCEB83B6045D8573D5D1D7E530E8E15C4BDB9BAF777D223ABBA
          FDF3A5D41D3058C380E3D597D39193BC3705716F7E56B5AF93367545B9BB7E68
          2A7C292B2201B386D84F0DB8B0F548BA4AB2E94741DC5BB222B2A7935E596D56
          EC9AFED080536C2406C47F14E6EB55E79E6E9B5F97CE841CFE5010F7967DCD34
          D88946F19CDCAD194F3F746EC0AA55AB4E018889A67FB003CE3DF4E4F4EEDD3A
          6445108EBDC305716FA90D0E00C4416D2F2188BB8F3D226B796600E86FF60C41
          DC7DF97E364F0D00DD65BE2988BB4F5DF2786E00E8ADE41E41DC7D1BD5952707
          80C68A3B0941DC6516CDECE7E101A0AD177708E22E694C15CF0F003D1D6E2588
          BBB4F5AB79820068A876CB7C41DC03183A95870880760EDD2A04710FE6298666
          006866EA5041DC83AF89DCC493044023B3C70841DC1558C0AA1900FAE8FA8820
          EE8ADCCA074D00F450B57BBE20EECA7C3898470A80060EC7E4B53D2E7117E293
          433C550022367BB210C45DB17BF6F260018854C98F82B887605E26CF1680C8E4
          CD1082B887B3D9CC16D6BC038846ED86E582B887E691121E31001138DC4508E2
          1EA619EC360320ED2332938420EE618FCDF4E28C2600E954B5658120EE69306A
          2F07AC02489B1EAD8420EE69DA6E6625CF1B80B4E8D94508E29E3E6339830F40
          F8FAEF1082B8A7D737FD79EE00846AD32421887B043B12907700E1593D66BE20
          EEE41D8061695F2704718F6ED93B63EF00421890991BF7B4C73DEE428C2D29E6
          4104A052DF49F38520EE917BB3075BCE005066E5ADC2081906FC86CF3A715413
          0015CEF7785308E2AE8F750B197C071054DE9651421077DD3E5BDD7F90471380
          BCAE0B870B41DC35B4600CAFEF00E41CDCBF5E9825C3A85F33ED01B60406E0FF
          A57DCC722188BBDEA3EFF37AB02730001FF21EF85C1828C3BC9F347CC6A07D3C
          AF005A62768F79EB8420EEB1197E5F3888F77700CDD83775E10261AA0C537FD8
          F01D7B0FF1EC02F02CFBA049E696DDE4B85F31A453574E6D0270ADBCBD338C2E
          BBE9716F3274610F16D000B84A6DDF2DEB85F9322CF88D6F8D1ECCFE0400AE58
          BD77D25061850C3B7EA6F871EEBD7D19A201AC96B977CC5BC21A19F6FC54B17C
          ECEE1E9B283C60A1D93DC77DF299B04A86B0CC8221A3C7AD649406B046DECA71
          73A709FB64082B2D5F3FA9D7DEC399EC040F18FCB6DE77EAB83F8DBD47582A43
          586DE8FA6F46F71AD7A3E470DFCCBCD97CF704C4DCC1838736DDBD72F0FE4EA3
          5FEB32D4EEB8D91E770020EE0000E20E0020EE0000E20E0020EE0040DC0100C4
          1D0040DC0100C41D0040DC0100C41D00883B0080B80300883B0080B80300883B
          00107700007107001077000071070010770000710700E20E0020EE0000E20E00
          20EE0000E20E00C41D0040DC0100C41D0040DC0100C41D0040DC0180B8030088
          3B0080B80300883B0080B803007107001077000071070010770000710700E2CE
          250000E20E0020EE0000E20E0020EE0000E20E00C41D0040DC0100C41D0040DC
          0100C41D00883B0080B80300883B0080B80300883B0080B80300710700107700
          0071070010770000710700E20E0020EE0000E20E0020EE0000E20E0020EE0040
          DC0100C41D0040DC0100C41D0040DC0180B80300883B0080B80300883B0080B8
          0300883B00107700007107001077000071070010770020EE0000E20E0020EE00
          00E20E0020EE0000E20E00C41D0040DC0100C41D0040DC0100C41D00883B0080
          B80300883B0080B80300883B0080B80300710700107700007107001077000071
          0700E20E0020EE0000E20E0020EE0000E20E0020EE0040DC0100C41D0040DC01
          00C41D0040DC0180B80300883B0080B80300883B0080B80300883B0010770000
          7107001077000071070010770020EE0000E20E0020EE0000E20E0020EE0000E2
          0E00C41D0040DC0100C41D0040DC0100C41D00883B0080B80300883B0080B803
          00883B0080B803007107001077000071070010770000710700E20E0020EE0000
          E20E0020EE0000E20E0020EE0040DC0100C41D0040DC0100C41D0040DC0180B8
          0300883B0080B80300883B0080B80300883B00107700007107001077008072FF
          27C00024B91686E140A8AF0000000049454E44AE426082}
        Stretch = True
      end
      object cpnlMensagem: TCurvyPanel
        Left = 137
        Top = 269
        Width = 525
        Height = 70
        BorderColor = 4227327
        Color = 6488063
        Rounding = 20
        TabOrder = 0
        TabStop = True
        object Label23: TLabel
          Left = 7
          Top = 2
          Width = 510
          Height = 59
          Cursor = crHandPoint
          Caption = 'Produto n'#227'o encontrado!'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -43
          Font.Name = 'Segoe UI'
          Font.Style = [fsBold]
          ParentFont = False
        end
      end
    end
  end
  object pop_principal: TAdvPopupMenu
    MenuStyler = frmModulo.estilo_menu
    Version = '2.6.5.11'
    Left = 560
    Top = 243
    object Opes1: TMenuItem
      Caption = 'Op'#231#245'es'
      ShortCut = 112
      OnClick = Opes1Click
    end
    object PesquisadeProdutos1: TMenuItem
      Caption = 'Pesquisa de Produtos'
      ShortCut = 113
      OnClick = PesquisadeProdutos1Click
    end
    object FecharCupom1: TMenuItem
      Caption = 'Fechar Cupom'
      ShortCut = 114
      OnClick = FecharCupom1Click
    end
    object CancelarCupom1: TMenuItem
      Caption = 'Cancelar Cupom'
      ShortCut = 115
      OnClick = CancelarCupom1Click
    end
    object CancelarItem1: TMenuItem
      Caption = 'Cancelar Item'
      ShortCut = 116
      OnClick = CancelarItem1Click
    end
    object ConsultadePreos1: TMenuItem
      Caption = 'Consulta de Pre'#231'os'
      ShortCut = 117
      OnClick = ConsultadePreos1Click
    end
    object ConsultadeClientes1: TMenuItem
      Caption = 'Consulta de Clientes'
      ShortCut = 118
      OnClick = ConsultadeClientes1Click
    end
    object menu_Prevenda: TMenuItem
      Caption = 'Pr'#233'-vendas'
      ShortCut = 120
      OnClick = menu_PrevendaClick
    end
    object AcionaBalana1: TMenuItem
      Caption = 'Pesar Produto'
      ShortCut = 121
      OnClick = AcionaBalana1Click
    end
    object mniAlterarValorUnitario1: TMenuItem
      Caption = 'Alte&rar Valor &Unit'#225'rio'
      ShortCut = 16469
      OnClick = mniAlterarValorUnitario1Click
    end
    object mniAlterarQuantidadeItens: TMenuItem
      Caption = 'Alterar Quantidade Itens'
      ShortCut = 16465
      Visible = False
    end
    object AcionaGaveta1: TMenuItem
      Caption = 'Abrir Gaveta'
      ShortCut = 16455
      OnClick = AcionaGaveta1Click
    end
    object Cupons1: TMenuItem
      Caption = 'Cupons'
      ShortCut = 16464
      OnClick = Cupons1Click
    end
    object C1: TMenuItem
      Caption = 'Consumidor'
      ShortCut = 16449
      OnClick = C1Click
    end
    object Atendimento1: TMenuItem
      Caption = 'Atendimento'
      object Oramento1: TMenuItem
        Caption = 'Or'#231'amento'
        ShortCut = 16461
        Visible = False
        OnClick = Oramento1Click
      end
      object menu_mesa: TMenuItem
        Caption = 'Mesas'
        ShortCut = 16453
        OnClick = menu_mesaClick
      end
      object C2: TMenuItem
        Caption = 'Comanda'
        ShortCut = 16451
        OnClick = C2Click
      end
      object mnDelivery: TMenuItem
        Caption = 'Delivery'
        ShortCut = 16452
        OnClick = mnDeliveryClick
      end
    end
    object Caixa1: TMenuItem
      Caption = 'Caixa'
      object FecharCaixa1: TMenuItem
        Caption = 'Fechamento'
        ShortCut = 16454
        OnClick = FecharCaixa1Click
      end
      object Suprimento1: TMenuItem
        Caption = 'Suprimento'
        object mniRegistrarNovoSuprimento: TMenuItem
          Caption = 'Registrar Novo'
          ShortCut = 16467
          OnClick = Suprimento1Click
        end
        object mniCancelarUltimoSuprimento: TMenuItem
          Caption = 'Cancelar '#218'ltimo Suprimento'
          ShortCut = 49235
          OnClick = mniCancelarSangriaOuSuprimentoClick
        end
      end
      object Sangria1: TMenuItem
        Caption = 'Sangria'
        object mniLancarSangria: TMenuItem
          Caption = 'Registrar Nova'
          ShortCut = 16466
          OnClick = Sangria1Click
        end
        object mniCancelarUltimaSangria: TMenuItem
          Caption = 'Cancelar Ultima Sangria'
          ShortCut = 49234
          OnClick = mniCancelarSangriaOuSuprimentoClick
        end
      end
      object CaixaIndisponvel1: TMenuItem
        Caption = 'Caixa Indispon'#237'vel'
        ShortCut = 16457
        OnClick = CaixaIndisponvel1Click
      end
      object rocarOperador1: TMenuItem
        Caption = 'Trocar Operador'
        ShortCut = 16463
        OnClick = rocarOperador1Click
      end
    end
    object L1: TMenuItem
      Caption = 'Lan'#231'amentos'
      object C3: TMenuItem
        Caption = 'Cheff'
        ShortCut = 32837
        OnClick = C3Click
      end
    end
    object DevoluoTroca1: TMenuItem
      Caption = 'Devolu'#231#227'o/Troca'
      ShortCut = 16468
      OnClick = DevoluoTroca1Click
    end
    object G1: TMenuItem
      Caption = 'Gerenciar Cupons NFC-e'
      OnClick = G1Click
    end
    object mnEnviaContigencia: TMenuItem
      Caption = 'Enviar Conting'#234'ncias'
      OnClick = mnEnviaContigenciaClick
    end
    object mmSat: TMenuItem
      Caption = 'Gerenciar Cupons SAT'
      ShortCut = 49235
      OnClick = mmSatClick
    end
    object mmECF: TMenuItem
      Caption = 'Menu Fiscal ECF'
      Visible = False
      OnClick = mmECFClick
    end
    object mmiAplicarVarejo: TMenuItem
      Caption = 'Aplicar Varejo'
      ShortCut = 32854
      OnClick = mmiAplicarVarejoClick
    end
    object N2: TMenuItem
      Caption = '-'
      Enabled = False
    end
    object Retornar1: TMenuItem
      Caption = 'Retornar'
      ShortCut = 27
      OnClick = Retornar1Click
    end
    object Sair1: TMenuItem
      Caption = 'Sair'
      ShortCut = 122
      OnClick = Sair1Click
    end
  end
  object timer_carga: TTimer
    Interval = 60000
    OnTimer = timer_cargaTimer
    Left = 624
    Top = 339
  end
  object TimerTroco: TTimer
    Enabled = False
    Interval = 20000
    OnTimer = TimerTrocoTimer
    Left = 592
    Top = 339
  end
  object query: TUniQuery
    Connection = frmModulo.conexao
    Left = 192
    Top = 131
  end
  object query2: TUniQuery
    Connection = frmModulo.conexao
    Left = 560
    Top = 371
  end
  object qrForma: TUniQuery
    Connection = frmModulo.conexao
    Left = 624
    Top = 435
  end
  object query3: TUniQuery
    Connection = frmModulo.conexao
    Left = 592
    Top = 435
  end
  object qrapoio: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from C000025')
    Left = 656
    Top = 435
  end
  object qrProdNCFE: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      
        'select CODIGO, NOME,COD_BARRA,NCM,CST,ALIQNACIONAL,CFOP,CODIGO_A' +
        'NP,PERCGLNN,PGLP,VPART,PERCGLGNI,informa_codigo_barra_xml from E' +
        'STOQUE ')
    Left = 560
    Top = 435
  end
  object QRCSOSN: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select CSOSN from ESTOQUE')
    Left = 624
    Top = 403
  end
  object tmContigencia: TJvThreadTimer
    Interval = 500
    OnTimer = tmContigenciaTimer
    Left = 688
    Top = 307
  end
  object qrContingencia: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      '  *'
      'from'
      '  NFCE nf'
      'where'
      '  nf.contingencia = '#39'S'#39' and nf.enviadocontingencia = '#39'N'#39)
    CachedUpdates = True
    Left = 768
    Top = 507
    object qrContingenciaNUMERO: TIntegerField
      FieldName = 'NUMERO'
      Required = True
    end
    object qrContingenciaDATA: TDateField
      FieldName = 'DATA'
    end
    object qrContingenciaTOTAL: TFloatField
      FieldName = 'TOTAL'
    end
    object qrContingenciaCLIENTE: TStringField
      FieldName = 'CLIENTE'
      Size = 255
    end
    object qrContingenciaCHAVE: TStringField
      FieldName = 'CHAVE'
      Size = 255
    end
    object qrContingenciaXML: TStringField
      FieldName = 'XML'
      Size = 255
    end
    object qrContingenciaSITUACAO: TIntegerField
      FieldName = 'SITUACAO'
    end
    object qrContingenciaTROCO: TFloatField
      FieldName = 'TROCO'
    end
    object qrContingenciaEX: TIntegerField
      FieldName = 'EX'
    end
    object qrContingenciaHORA: TStringField
      FieldName = 'HORA'
      Size = 8
    end
    object qrContingenciaCONTINGENCIA: TStringField
      FieldName = 'CONTINGENCIA'
      Size = 1
    end
    object qrContingenciaENVIADOCONTINGENCIA: TStringField
      FieldName = 'ENVIADOCONTINGENCIA'
      Size = 1
    end
    object qrContingenciaMOTIVOCONTIGENCIA: TStringField
      FieldName = 'MOTIVOCONTIGENCIA'
      Size = 255
    end
    object qrContingenciaXMLENVIO: TBlobField
      FieldName = 'XMLENVIO'
    end
    object qrContingenciaXMLCACNELAMENTO: TBlobField
      FieldName = 'XMLCACNELAMENTO'
    end
    object qrContingenciaALTERADO: TStringField
      FieldName = 'ALTERADO'
      Size = 1
    end
    object qrContingenciaCUPOM: TStringField
      FieldName = 'CUPOM'
      Required = True
      Size = 50
    end
    object qrContingenciaXML_CANCELAMENTO: TStringField
      FieldName = 'XML_CANCELAMENTO'
      Size = 255
    end
    object qrContingenciaDATAHORACONTIMGENCIA: TDateTimeField
      FieldName = 'DATAHORACONTIMGENCIA'
    end
    object qrContingenciaSTATUS: TStringField
      FieldName = 'STATUS'
      Size = 1
    end
    object qrContingenciaERRO: TBlobField
      FieldName = 'ERRO'
    end
  end
  object UniDataSource1: TUniDataSource
    DataSet = frmModulo.qrItens
    Left = 1312
    Top = 291
  end
  object cxStyleRepository1: TcxStyleRepository
    Left = 656
    Top = 243
    PixelsPerInch = 96
    object stLinhas: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Tahoma'
      Font.Style = []
    end
    object stTitulo: TcxStyle
      AssignedValues = [svColor, svFont, svTextColor]
      Color = 16249582
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clYellow
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      TextColor = clBlack
    end
    object texto: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
    end
    object stExtra: TcxStyle
      AssignedValues = [svFont]
      Font.Charset = DEFAULT_CHARSET
      Font.Color = 8404992
      Font.Height = -9
      Font.Name = 'Tahoma'
      Font.Style = []
    end
    object cxStyle1: TcxStyle
      AssignedValues = [svColor]
      Color = 16249582
    end
  end
  object ImageList1: TImageList
    Left = 656
    Top = 307
    Bitmap = {
      494C010102000500040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000001000000001002000000000000010
      000000000000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFCEE7CE00B5DCB300B6DCB400CEE7CE00FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFF2F1EE00DBD6CF00CAC5BA00C4BEAE00C4BEAE00CBC5BB00DCD8CF00F3F2
      EF00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFF9DD1990056AE520050AF4C0050AF4C0050AF4C0050AF4C0055B152009ED2
      9C00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDAD6
      CE00C8C0B600E1DED700F3F1EF00FFFFFFFFFFFFFFFFF2F1EE00E1DED700C8C0
      B600DBD7CF00FFFFFFFFFFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFFE5F2E40065B8
      640050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF
      4C0067B96500E6F3E600FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFD1CBC200D7D2
      C900FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFD6D2C800D2CDC100FFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFF64B8620050AF
      4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF
      4C0050AF4C0067B96500FFFFFFFFFFFFFFFFFFFFFFFFDAD6CE00D7D2C900FFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFD6D2C800DBD7CF00FFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFF99CF970050AF4C0050AF
      4C0050AF4C0050AF4C0059BA5E0055B5550050AF4C0050AF4C0050AF4C0050AF
      4C0050AF4C0050AF4C009ED29C00FFFFFFFFF2F1EE00C9C2B400FFFFFFFFFFFF
      FFFF9DA3F8006F79F400FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6D75F300A1A8
      F800FFFFFFFFFFFFFFFFC8C0B600F3F2EF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFF53B0500050AF4C0050AF
      4C0050AF4C0059BB5F0089F7BF0084F0B40055B5550050AF4C0050AF4C0050AF
      4C0050AF4C0050AF4C0055B15200FFFFFFFFDBD7CE00E2DED800FFFFFFFFFFFF
      FFFF6872F5003643F4005762F400FFFFFFFFFFFFFFFF5661F4003643F4006D75
      F300FFFFFFFFFFFFFFFFE1DED700DCD8CF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000CBE5C80050AF4C0050AF4C0050AF
      4C0059BA5E0089F7BF0087F3B9008BF8C10084F0B40055B5550050AF4C0050AF
      4C0050AF4C0050AF4C0050AF4C00CEE7CE00CBC6BA00F3F2EF00FFFFFFFFFFFF
      FFFFFFFFFFFF535FF3003643F4005762F4005761F4003643F4005661F400FFFF
      FFFFFFFFFFFFFFFFFFFFF2F1EE00CBC5BB000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000B2DAB10050AF4C0050AF4C0050AF
      4C0083EFB10086F2B80056B758005BBD63008BF8C10084F0B40055B5550050AF
      4C0050AF4C0050AF4C0050AF4C00B6DCB400C2BDB000FFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFF5360F4003643F4003643F4005761F400FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFC4BEAE000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000B2DAB10050AF4C0050AF4C0050AF
      4C005ABB5F0056B7580050AF4C0050AF4C005BBD63008BF8C10084F0B40055B5
      550050AF4C0050AF4C0050AF4C00B5DCB300C2BDB000FFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFF5360F4003643F4003643F4005762F400FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFC4BEAE000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000CAE5C80050AF4C0050AF4C0050AF
      4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C005BBD63008BF8C10084F0
      B40055B5550050AF4C0050AF4C00CEE7CE00CBC4B900F3F2F000FFFFFFFFFFFF
      FFFFFFFFFFFF5360F4003643F4005360F4005360F4003643F4005762F400FFFF
      FFFFFFFFFFFFFFFFFFFFF3F1EF00CAC5BA000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFF53B0500050AF4C0050AF
      4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C005BBD63008BF8
      C1006ED4870050AF4C0056AE5200FFFFFFFFDAD6CE00E2DFD800FFFFFFFFFFFF
      FFFF6A74F5003643F4005360F400FFFFFFFFFFFFFFFF535FF3003643F4006F79
      F400FFFFFFFFFFFFFFFFE1DED700DBD6CF000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFF97CF950050AF4C0050AF
      4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0057B8
      5A0050AF4C0050AF4C009DD19900FFFFFFFFF2F1EE00C7C3B400FFFFFFFFFFFF
      FFFF9AA1F8006A74F500FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF6872F5009DA3
      F800FFFFFFFFFFFFFFFFC8C0B600F2F1EE000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFF63B7600050AF
      4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF
      4C0050AF4C0065B86400FFFFFFFFFFFFFFFFFFFFFFFFDAD5CD00D7D3CA00FFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFD7D2C900DBD7CE00FFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFFE4F1E30063B7
      600050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF4C0050AF
      4C0064B86200E5F2E400FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFD1CCC100D7D3
      CA00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFD7D2C900D1CBC200FFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFF97CF950053B0500050AF4C0050AF4C0050AF4C0050AF4C0053B0500099CF
      9700FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDAD5
      CD00C7C3B400E2DFD800F3F2F000FFFFFFFFFFFFFFFFF3F2EF00E2DED800C9C2
      B400DAD6CE00FFFFFFFFFFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FFFFFF00FFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFCAE5C800B2DAB100B2DAB100CBE5C800FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFF2F1EE00DAD6CE00CBC4B900C2BDB000C2BDB000CBC6BA00DBD7CE00F2F1
      EE00FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000100000000100010000000000800000000000000000000000
      000000000000000000000000FFFFFF0000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000}
  end
  object qrServidor: TUniQuery
    Connection = frmModulo.Conexao_Servidor
    Left = 920
    Top = 467
  end
  object frComprovanteDebito: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 42573.457939479200000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      ''
      'end.')
    OnBeforePrint = frComprovanteDebitoBeforePrint
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 336
    Top = 283
    Datasets = <
      item
        DataSet = frParcelas
        DataSetName = 'frParcelas'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 290.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object ReportTitle1: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 230.551330000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        PrintChildIfInvisible = True
        Stretched = True
        object Memo1: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 51.692950000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'COMPROVANTE DE D'#201'BITO')
          ParentFont = False
        end
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 72.149660000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: 23/03/2018 Hora: 13:02:45')
          ParentFont = False
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000001000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'NOME DA EMPRESA')
          ParentFont = False
        end
        object lbDocEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 23.677180000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CNPJ: 01.001.223/0001-11')
          ParentFont = False
        end
        object Line1: TfrxLineView
          AllowVectorExport = True
          Top = 45.354360000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 5.118120000000000000
          Top = 93.826840000000000000
          Width = 75.590600000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 5.118120000000000000
          Top = 119.401670000000000000
          Width = 64.252010000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CLIENTE')
          ParentFont = False
        end
        object Line2: TfrxLineView
          AllowVectorExport = True
          Left = 0.220470000000000000
          Top = 115.622140000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo8: TfrxMemoView
          AllowVectorExport = True
          Left = 5.118120000000000000
          Top = 138.299320000000000000
          Width = 64.252010000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'ENDERECO')
          ParentFont = False
        end
        object Memo9: TfrxMemoView
          AllowVectorExport = True
          Left = 5.118120000000000000
          Top = 157.196970000000000000
          Width = 64.252010000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CPF/CNPJ')
          ParentFont = False
        end
        object Memo10: TfrxMemoView
          AllowVectorExport = True
          Left = 4.000000000000000000
          Top = 185.094620000000000000
          Width = 264.567100000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'PEGO POR QUEM')
          ParentFont = False
        end
        object Memo11: TfrxMemoView
          AllowVectorExport = True
          Left = 71.826840000000000000
          Top = 119.401670000000000000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          AllowVectorExport = True
          Left = 71.826840000000000000
          Top = 138.299320000000000000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo13: TfrxMemoView
          AllowVectorExport = True
          Left = 71.826840000000000000
          Top = 157.196970000000000000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 82.149660000000000000
          Top = 119.401670000000000000
          Width = 188.976500000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'SERGIO LUIZ NUNES FARIA')
          ParentFont = False
          WordWrap = False
        end
        object lbEndereco: TfrxMemoView
          AllowVectorExport = True
          Left = 82.149660000000000000
          Top = 138.299320000000000000
          Width = 188.976500000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'RUA PROF. MARIA BENEDITA FERNANDES, 339')
          ParentFont = False
          WordWrap = False
        end
        object lbDocCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 82.149660000000000000
          Top = 157.196970000000000000
          Width = 188.976500000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '182.107.618-09')
          ParentFont = False
          WordWrap = False
        end
        object lbPegoPor: TfrxMemoView
          AllowVectorExport = True
          Left = 4.000000000000000000
          Top = 201.992270000000000000
          Width = 264.567100000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'O MESMO')
          ParentFont = False
        end
        object Line4: TfrxLineView
          AllowVectorExport = True
          Top = 224.889920000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo28: TfrxMemoView
          AllowVectorExport = True
          Left = 83.708720000000000000
          Top = 93.826840000000000000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 95.267780000000000000
          Top = 93.826840000000000000
          Width = 173.858380000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '0001')
          ParentFont = False
        end
        object Line3: TfrxLineView
          AllowVectorExport = True
          Top = 2.779530000000001000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object MasterData1: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 64.252010000000000000
        Top = 332.598640000000000000
        Width = 272.126160000000000000
        DataSet = frParcelas
        DataSetName = 'frParcelas'
        RowCount = 0
        object Memo18: TfrxMemoView
          AllowVectorExport = True
          Left = 4.000000000000000000
          Top = 1.559059999999931000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'PARCELA')
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 4.000000000000000000
          Top = 20.456709999999870000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'VENCIMENTO')
          ParentFont = False
        end
        object Memo20: TfrxMemoView
          AllowVectorExport = True
          Left = 4.000000000000000000
          Top = 39.354359999999700000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'VALOR')
          ParentFont = False
        end
        object Memo21: TfrxMemoView
          AllowVectorExport = True
          Left = 85.708720000000000000
          Top = 1.559059999999931000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          AllowVectorExport = True
          Left = 85.708720000000000000
          Top = 20.456709999999870000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo23: TfrxMemoView
          AllowVectorExport = True
          Left = 85.708720000000000000
          Top = 39.354359999999700000
          Width = 7.559060000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbParcela: TfrxMemoView
          AllowVectorExport = True
          Left = 95.031540000000000000
          Top = 1.559059999999931000
          Width = 173.858380000000000000
          Height = 18.897650000000000000
          DataField = 'parcela'
          DataSet = frParcelas
          DataSetName = 'frParcelas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frParcelas."parcela"]')
          ParentFont = False
        end
        object lbVencimento: TfrxMemoView
          AllowVectorExport = True
          Left = 95.031540000000000000
          Top = 20.456709999999870000
          Width = 173.858380000000000000
          Height = 18.897650000000000000
          DataField = 'Vencimento'
          DataSet = frParcelas
          DataSetName = 'frParcelas'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frParcelas."Vencimento"]')
          ParentFont = False
        end
        object lbValor: TfrxMemoView
          AllowVectorExport = True
          Left = 95.031540000000000000
          Top = 39.354359999999990000
          Width = 173.858380000000000000
          Height = 18.897650000000000000
          DataField = 'Valor'
          DataSet = frParcelas
          DataSetName = 'frParcelas'
          DisplayFormat.FormatStr = '%2.2m'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frParcelas."Valor"]')
          ParentFont = False
        end
      end
      object GroupHeader1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Top = 309.921460000000000000
        Width = 272.126160000000000000
        Condition = '1=1'
      end
      object GroupFooter1: TfrxGroupFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 132.283550000000000000
        Top = 419.527830000000000000
        Width = 272.126160000000000000
        object Line5: TfrxLineView
          AllowVectorExport = True
          Top = 3.559059999999988000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Line6: TfrxLineView
          Align = baCenter
          AllowVectorExport = True
          Left = 24.566945000000000000
          Top = 90.488249999999420000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = [ftTop]
        end
        object Memo27: TfrxMemoView
          Align = baCenter
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 90.488249999999420000
          Width = 188.976500000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'ASSINATURA DO CLIENTE')
          ParentFont = False
        end
      end
    end
  end
  object cdsParcelas: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Vencimento'
        DataType = ftDate
      end
      item
        Name = 'Valor'
        DataType = ftFloat
      end
      item
        Name = 'Parcela'
        DataType = ftString
        Size = 2
      end>
    IndexDefs = <
      item
        Name = 'cdsParcelasIndex1'
        Fields = 'Parcela'
      end>
    IndexFieldNames = 'Parcela'
    Params = <>
    StoreDefs = True
    Left = 656
    Top = 131
    object cdsParcelasVencimento: TDateField
      FieldName = 'Vencimento'
    end
    object cdsParcelasValor: TFloatField
      FieldName = 'Valor'
    end
    object cdsParcelasParcela: TStringField
      FieldName = 'Parcela'
      Size = 2
    end
  end
  object frParcelas: TfrxDBDataset
    UserName = 'frParcelas'
    CloseDataSource = False
    FieldAliases.Strings = (
      'Vencimento=Vencimento'
      'Valor=Valor'
      'Parcela=Parcela')
    DataSet = cdsParcelas
    BCDToCurrency = False
    Left = 560
    Top = 275
  end
  object RelCupomNF: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Microsoft Print to PDF'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 43465.564168055560000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      '  if <frCupomNF."SENHA"> > 0 then'
      '    cpSenha.Visible := True'
      '  else                  '
      '    cpSenha.Visible := False;'
      '  if trim(<frObs."OBS">) <> '#39#39' then begin             '
      '    cpO1.Visible := True;  '
      '    cpO2.Visible := True;  '
      '  end else begin                               '
      '    cpO1.Visible := False;'
      '    cpO2.Visible := False;'
      '  end;                    '
      '  '
      '  if <frCupomNF."PEGO_PORQUEM"> = '#39#39' then begin'
      '    cpPego.Visible := False;'
      
        '    cpTitulo.Child := cpConsumidor;                             ' +
        '                 '
      '  end;'
      '  if <frCupomNF."CPF_CONSUMIDOR"> = '#39#39' then'
      '    cpConsumidor.Visible := False;                          '
      '                                              '
      'end.')
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 152
    Top = 299
    Datasets = <
      item
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end
      item
        DataSet = frImpForma
        DataSetName = 'frImpForma'
      end
      item
        DataSet = frOBS
        DataSetName = 'frObs'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 100000.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object cpTitulo: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 188.976500000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        Child = RelCupomNF.cpPego
        PrintChildIfInvisible = True
        Stretched = True
        object Memo1: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 93.267780000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'COMPROVANTE DE VENDA')
          ParentFont = False
        end
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 112.724490000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: [frCupomNF."DATA"] Hora: [frCupomNF."HORA"]')
          ParentFont = False
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000001000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
        object lbDocEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 23.677180000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CNPJ"]')
          ParentFont = False
        end
        object Line1: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 89.708720000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 129.551330000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 164.756030000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CLIENTE')
          ParentFont = False
        end
        object Line2: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 183.653680000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo11: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 164.756030000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 164.756030000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."CLIENTE"]')
          ParentFont = False
          WordWrap = False
        end
        object Memo28: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 129.551330000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 129.551330000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NUMERO"]')
          ParentFont = False
        end
        object Line3: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000001000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo2: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 39.354360000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."ENDERECO"]')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 54.913420000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CIDADE"]/[frFilial."UF"] - [frFilial."CEP"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 70.370130000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."TELEFONE"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end
            item
            end>
        end
        object Memo6: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 147.401670000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CAIXA')
          ParentFont = False
        end
        object Memo14: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 147.401670000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo15: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 147.401670000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."COD_CAIXA"]')
          ParentFont = False
        end
      end
      object MasterData1: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 30.236240000000000000
        Top = 510.236550000000000000
        Width = 272.126160000000000000
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
        RowCount = 0
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 15.118119999999920000
          Width = 30.236240000000000000
          Height = 11.338590000000000000
          DataField = 'QTDE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."QTDE"]')
          ParentFont = False
        end
        object Memo20: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 15.118119999999920000
          Width = 22.677180000000000000
          Height = 11.338590000000000000
          DataField = 'UNIDADE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."UNIDADE"]')
          ParentFont = False
        end
        object Memo21: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 15.118119999999920000
          Width = 52.913420000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_UNITARIO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_UNITARIO"]')
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 15.118119999999920000
          Width = 60.472480000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_TOTAL'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_TOTAL"]')
          ParentFont = False
        end
        object Memo23: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 15.118119999999920000
          Width = 41.574830000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_DESCONTO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_DESCONTO"]')
          ParentFont = False
        end
        object Memo25: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 15.118119999999920000
          Width = 41.574830000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_ACRESCIMO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_ACRESCIMO"]')
          ParentFont = False
        end
        object Memo18: TfrxMemoView
          AllowVectorExport = True
          Width = 22.677180000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."ITEM"]')
          ParentFont = False
        end
        object Memo32: TfrxMemoView
          AllowVectorExport = True
          Left = 25.236240000000000000
          Width = 71.811070000000000000
          Height = 11.338590000000000000
          DataField = 'COD_BARRA'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."COD_BARRA"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo33: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Width = 170.078850000000000000
          Height = 11.338590000000000000
          DataField = 'NOME'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NOME"]')
          ParentFont = False
        end
      end
      object GroupHeader1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Top = 487.559370000000000000
        Width = 272.126160000000000000
        Condition = 'frCupomNF."COD_CUPOM"'
      end
      object MasterData3: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 18.897650000000000000
        Top = 733.228820000000000000
        Width = 272.126160000000000000
        Child = RelCupomNF.Child2
        DataSet = frImpForma
        DataSetName = 'frImpForma'
        RowCount = 0
        object Memo26: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataField = 'FORMA_DESCRICAO'
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frImpForma."FORMA_DESCRICAO"]')
          ParentFont = False
        end
        object Memo27: TfrxMemoView
          AllowVectorExport = True
          Left = 140.063080000000000000
          Top = 1.779530000000022000
          Width = 56.692950000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frImpForma."FORMA_PRESTACAO"] X')
          ParentFont = False
        end
        object Memo29: TfrxMemoView
          AllowVectorExport = True
          Left = 200.315090000000000000
          Top = 1.779530000000022000
          Width = 68.031540000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[<frImpForma."FORMA_VALOR">/<frImpForma."FORMA_PRESTACAO">]')
          ParentFont = False
        end
      end
      object cpTotal: TfrxGroupFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 104.606370000000000000
        Top = 563.149970000000000000
        Width = 272.126160000000000000
        Child = RelCupomNF.Child1
        object ltotal: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 9.252010000000041000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Total')
          ParentFont = False
        end
        object ptotal: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 9.252010000000041000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtotal: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 9.252010000000041000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TOTAL"]')
          ParentFont = False
        end
        object lpago: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 66.149659999999920000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Pago')
          ParentFont = False
        end
        object ppago: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 66.149659999999920000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vpago: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 66.149659999999920000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_PAGO"]')
          ParentFont = False
        end
        object ltroco: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 85.047309999999690000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Troco')
          ParentFont = False
        end
        object ptroco: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 85.047309999999690000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtroco: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 85.047309999999690000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TROCO"]')
          ParentFont = False
        end
        object Line5: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000136000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo40: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 47.236239999999960000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Desc.')
          ParentFont = False
        end
        object Memo41: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 47.236239999999960000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo42: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 47.236239999999960000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<frCupomNF."ITEM_DESCONTO">,MasterData1)]')
          ParentFont = False
        end
        object Memo43: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 26.456710000000040000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor  Acr'#233's.')
          ParentFont = False
        end
        object Memo44: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 26.456709999999930000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo46: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 26.456710000000040000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<frCupomNF."ITEM_ACRESCIMO">,MasterData1)]')
          ParentFont = False
        end
      end
      object Child1: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 18.897650000000000000
        Top = 691.653990000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo45: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Forma de Pagamento:')
          ParentFont = False
        end
      end
      object Child2: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 774.803650000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object cpSenha: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 3.000000000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Senha de Retirada: [frCupomNF."SENHA"]')
          ParentFont = False
        end
        object Line6: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object Header1: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 34.015770000000000000
        Top = 430.866420000000000000
        Width = 272.126160000000000000
        object Line4: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 32.133889999999890000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo8: TfrxMemoView
          AllowVectorExport = True
          Width = 26.456710000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Item')
          ParentFont = False
        end
        object Memo9: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 15.118119999999980000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Qtde')
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 15.118119999999980000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Un.')
          ParentFont = False
        end
        object Memo13: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 15.118119999999980000
          Width = 52.913420000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Unit.')
          ParentFont = False
        end
        object Memo16: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 15.118119999999980000
          Width = 60.472480000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Total')
          ParentFont = False
        end
        object Memo17: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Desc.')
          ParentFont = False
        end
        object Memo24: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Acr'#233's.')
          ParentFont = False
        end
        object Memo30: TfrxMemoView
          AllowVectorExport = True
          Left = 25.236240000000000000
          Width = 71.811070000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'C'#243'digo')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo31: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Width = 45.354360000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Produto')
          ParentFont = False
        end
      end
      object cpPego: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 37.795300000000000000
        Top = 230.551330000000000000
        Width = 272.126160000000000000
        Child = RelCupomNF.cpConsumidor
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo10: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 0.897649999999998700
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'PEGO POR QUEM')
          ParentFont = False
        end
        object lbPegoPor: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 14.015770000000000000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."PEGO_PORQUEM"]')
          ParentFont = False
        end
        object Line7: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 35.015769999999970000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object cpConsumidor: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 79.370130000000000000
        Top = 291.023810000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo34: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 0.897649999999998700
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CONSUMIDOR')
          ParentFont = False
        end
        object Memo35: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 21.574830000000020000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Nome: [frCupomNF."NOME_CONSUMIDOR"]')
          ParentFont = False
        end
        object Line8: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 74.590600000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo36: TfrxMemoView
          AllowVectorExport = True
          Top = 37.795300000000000000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CPF: [frCupomNF."CPF_CONSUMIDOR"]')
          ParentFont = False
        end
        object Memo37: TfrxMemoView
          AllowVectorExport = True
          Top = 52.913419999999970000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'End: [frCupomNF."ENDERECO_CONSUMIDOR"]')
          ParentFont = False
        end
      end
      object cpO2: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 865.512370000000000000
        Width = 272.126160000000000000
        DataSet = frOBS
        DataSetName = 'frObs'
        RowCount = 0
        Stretched = True
        object Memo38: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 3.779530000000022000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          DataField = 'obs'
          DataSet = frOBS
          DataSetName = 'frObs'
          Frame.Typ = []
          Memo.UTF8W = (
            '[frObs."obs"]')
        end
      end
      object cpO1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 820.158010000000000000
        Width = 272.126160000000000000
        Condition = '1=1'
        object Memo39: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Observa'#231#227'o:')
          ParentFont = False
        end
        object Line9: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
    end
  end
  object frCupomNF: TfrxDBDataset
    UserName = 'frCupomNF'
    CloseDataSource = False
    FieldAliases.Strings = (
      'COD_CUPOM=COD_CUPOM'
      'NUMERO=NUMERO'
      'DATA=DATA'
      'HORA=HORA'
      'CUPOM_TOTAL=CUPOM_TOTAL'
      'CUPOM_PAGO=CUPOM_PAGO'
      'CUPOM_TROCO=CUPOM_TROCO'
      'CLIENTE=CLIENTE'
      'VENDEDOR=VENDEDOR'
      'CPF_CONSUMIDOR=CPF_CONSUMIDOR'
      'NOME_CONSUMIDOR=NOME_CONSUMIDOR'
      'COD_CAIXA=COD_CAIXA'
      'PEGO_PORQUEM=PEGO_PORQUEM'
      'COUVERT=COUVERT'
      'TAXA_GARCON=TAXA_GARCON'
      'COD_ITEM=COD_ITEM'
      'ITEM=ITEM'
      'COD_PRODUTO=COD_PRODUTO'
      'NOME=NOME'
      'COD_BARRA=COD_BARRA'
      'UNIDADE=UNIDADE'
      'QTDE=QTDE'
      'ITEM_UNITARIO=ITEM_UNITARIO'
      'ITEM_SUBTOTAL=ITEM_SUBTOTAL'
      'ITEM_DESCONTO=ITEM_DESCONTO'
      'ITEM_ACRESCIMO=ITEM_ACRESCIMO'
      'ITEM_TOTAL=ITEM_TOTAL'
      'SENHA=SENHA'
      'ENDERECO_CONSUMIDOR=ENDERECO_CONSUMIDOR')
    DataSet = qrCupomNF
    BCDToCurrency = False
    Left = 792
    Top = 283
  end
  object qrCupomNF: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      ' cup.codigo cod_cupom,'
      ' cup.numero,'
      ' cup.data,'
      ' cup.hora,'
      ' cup.valor_total cupom_total,'
      ' cup.valor_pago cupom_pago,'
      ' cup.valor_troco cupom_troco,'
      ' cup.senha, '
      ' cli.nome cliente,'
      ' ven.nome vendedor,'
      ' cup.cpf_consumidor,'
      ' cup.nome_consumidor,'
      ' cup.endereco_consumidor,'
      ' cup.cod_caixa,'
      ' cup.pego_porquem,'
      ' cup.couvert,'
      ' cup.taxa_garcon,'
      ' ite.codigo cod_item,'
      ' lpad(ite.item, 3, '#39'0'#39') item,'
      ' ite.cod_produto,'
      ' case when trim(coalesce(ite.complemento, '#39#39')) = '#39#39' then'
      '   pro.nome'
      ' else'
      '  ite.complemento'
      ' end nome,'
      ' pro.cod_barra,'
      ' ite.unidade,'
      ' ite.qtde,'
      ' ite.valor_unitario item_unitario,'
      ' ite.valor_subtotal item_subtotal,'
      ' ite.valor_desconto item_desconto,'
      ' ite.valor_acrescimo item_acrescimo,'
      ' ite.valor_total item_total'
      'from'
      '  cupom cup'
      '  inner join cupom_item ite on ite.cod_cupom = cup.codigo'
      '  inner join estoque pro on pro.codigo = ite.cod_produto'
      '  left join cliente cli on cli.codigo = cup.cod_cliente'
      '  left join vendedores ven on ven.codigo = cup.cod_vendedor'
      'where'
      '  cup.codigo = :COD'
      'order by'
      '  ite.item'
      '')
    Left = 440
    Top = 467
    ParamData = <
      item
        DataType = ftString
        Name = 'cod'
        ParamType = ptInput
        Value = '099000000901166'
      end>
    object qrCupomNFCOD_CUPOM: TStringField
      FieldName = 'COD_CUPOM'
      Required = True
      Size = 50
    end
    object qrCupomNFNUMERO: TStringField
      FieldName = 'NUMERO'
      Required = True
      Size = 12
    end
    object qrCupomNFDATA: TDateField
      FieldName = 'DATA'
      Required = True
    end
    object qrCupomNFHORA: TTimeField
      FieldName = 'HORA'
      Required = True
    end
    object qrCupomNFCUPOM_TOTAL: TFloatField
      FieldName = 'CUPOM_TOTAL'
    end
    object qrCupomNFCUPOM_PAGO: TFloatField
      FieldName = 'CUPOM_PAGO'
    end
    object qrCupomNFCUPOM_TROCO: TFloatField
      FieldName = 'CUPOM_TROCO'
    end
    object qrCupomNFCLIENTE: TStringField
      FieldName = 'CLIENTE'
      ReadOnly = True
      Required = True
      Size = 80
    end
    object qrCupomNFVENDEDOR: TStringField
      FieldName = 'VENDEDOR'
      ReadOnly = True
      Size = 80
    end
    object qrCupomNFCPF_CONSUMIDOR: TStringField
      FieldName = 'CPF_CONSUMIDOR'
      Size = 18
    end
    object qrCupomNFNOME_CONSUMIDOR: TStringField
      FieldName = 'NOME_CONSUMIDOR'
      Size = 40
    end
    object qrCupomNFCOD_CAIXA: TIntegerField
      FieldName = 'COD_CAIXA'
      Required = True
    end
    object qrCupomNFPEGO_PORQUEM: TStringField
      FieldName = 'PEGO_PORQUEM'
      Size = 60
    end
    object qrCupomNFCOUVERT: TFloatField
      FieldName = 'COUVERT'
    end
    object qrCupomNFTAXA_GARCON: TFloatField
      FieldName = 'TAXA_GARCON'
    end
    object qrCupomNFCOD_ITEM: TStringField
      FieldName = 'COD_ITEM'
      ReadOnly = True
      Required = True
      Size = 50
    end
    object qrCupomNFITEM: TStringField
      FieldName = 'ITEM'
      ReadOnly = True
      Size = 3
    end
    object qrCupomNFCOD_PRODUTO: TIntegerField
      FieldName = 'COD_PRODUTO'
      ReadOnly = True
      Required = True
    end
    object qrCupomNFNOME: TStringField
      FieldName = 'NOME'
      ReadOnly = True
      Required = True
      Size = 80
    end
    object qrCupomNFCOD_BARRA: TStringField
      FieldName = 'COD_BARRA'
      ReadOnly = True
      Size = 15
    end
    object qrCupomNFUNIDADE: TStringField
      FieldName = 'UNIDADE'
      ReadOnly = True
      Required = True
      Size = 5
    end
    object qrCupomNFQTDE: TFloatField
      FieldName = 'QTDE'
      ReadOnly = True
    end
    object qrCupomNFITEM_UNITARIO: TFloatField
      FieldName = 'ITEM_UNITARIO'
      ReadOnly = True
    end
    object qrCupomNFITEM_SUBTOTAL: TFloatField
      FieldName = 'ITEM_SUBTOTAL'
      ReadOnly = True
    end
    object qrCupomNFITEM_DESCONTO: TFloatField
      FieldName = 'ITEM_DESCONTO'
      ReadOnly = True
    end
    object qrCupomNFITEM_ACRESCIMO: TFloatField
      FieldName = 'ITEM_ACRESCIMO'
      ReadOnly = True
    end
    object qrCupomNFITEM_TOTAL: TFloatField
      FieldName = 'ITEM_TOTAL'
      ReadOnly = True
    end
    object qrCupomNFSENHA: TIntegerField
      FieldName = 'SENHA'
    end
    object qrCupomNFENDERECO_CONSUMIDOR: TStringField
      FieldName = 'ENDERECO_CONSUMIDOR'
      Size = 130
    end
  end
  object frFilial: TfrxDBDataset
    UserName = 'frFilial'
    CloseDataSource = False
    FieldAliases.Strings = (
      'CODIGO=CODIGO'
      'FILIAL=FILIAL'
      'NOTAFISCAL=NOTAFISCAL'
      'ENDERECO=ENDERECO'
      'CIDADE=CIDADE'
      'UF=UF'
      'CEP=CEP'
      'TELEFONE=TELEFONE'
      'CNPJ=CNPJ'
      'IE=IE'
      'SEQNF=SEQNF'
      'FAX=FAX'
      'OBS1=OBS1'
      'OBS2=OBS2'
      'CONTRIBUINTE_IPI=CONTRIBUINTE_IPI'
      'SUBSTITUTO_TRIBUTARIO=SUBSTITUTO_TRIBUTARIO'
      'EMPRESA_ESTADUAL=EMPRESA_ESTADUAL'
      'OPTANTE_SIMPLES=OPTANTE_SIMPLES'
      'OPTANTE_SUPER_SIMPLES=OPTANTE_SUPER_SIMPLES'
      'ECF=ECF'
      'TIPO=TIPO'
      'IPI=IPI'
      'ISS=ISS'
      'NUMERO=NUMERO'
      'RESPONSAVEL=RESPONSAVEL'
      'COMPLEMENTO=COMPLEMENTO'
      'BAIRRO=BAIRRO'
      'FARMACIA_RESP_TECNICO=FARMACIA_RESP_TECNICO'
      'FARMACIA_CRF=FARMACIA_CRF'
      'FARMACIA_CPF=FARMACIA_CPF'
      'FARMCIA_DATA=FARMCIA_DATA'
      'FARMACIA_UF=FARMACIA_UF'
      'FARMACIA_SENHA=FARMACIA_SENHA'
      'FARMACIA_EMAIL=FARMACIA_EMAIL'
      'FARMACIA_LOGIN=FARMACIA_LOGIN'
      'FARMACIA_ENVIO=FARMACIA_ENVIO'
      'CONHECIMENTO=CONHECIMENTO'
      'INDUSTRIA=INDUSTRIA'
      'FARMACIA_NUMEROLICENCA=FARMACIA_NUMEROLICENCA'
      'COD_MUNICIPIO_IBGE=COD_MUNICIPIO_IBGE'
      'IBGE=IBGE'
      'PIS=PIS'
      'COFINS=COFINS'
      'EMAIL=EMAIL'
      'ATIVIDADE=ATIVIDADE'
      'CONTADOR_COD_MUNICIPIO_IBGE=CONTADOR_COD_MUNICIPIO_IBGE'
      'CONTADOR_NOME=CONTADOR_NOME'
      'CONTADOR_CPF=CONTADOR_CPF'
      'CONTADOR_CRC=CONTADOR_CRC'
      'CONTADOR_CNPJ=CONTADOR_CNPJ'
      'CONTADOR_CEP=CONTADOR_CEP'
      'CONTADOR_ENDERECO=CONTADOR_ENDERECO'
      'CONTADOR_NUMERO=CONTADOR_NUMERO'
      'CONTADOR_COMPLEMENTO=CONTADOR_COMPLEMENTO'
      'CONTADOR_BAIRRO=CONTADOR_BAIRRO'
      'CONTADOR_FONE=CONTADOR_FONE'
      'CONTADOR_FAX=CONTADOR_FAX'
      'CONTADOR_EMAIL=CONTADOR_EMAIL'
      'INSC_MUNICIPAL=INSC_MUNICIPAL'
      'DATA_ABERTURA=DATA_ABERTURA'
      'CNAE=CNAE'
      'CRT=CRT'
      'CONTADOR_CIDADE=CONTADOR_CIDADE'
      'CONTADOR_COD_MUNICIPIO=CONTADOR_COD_MUNICIPIO'
      'CONTADOR_UF=CONTADOR_UF'
      'PERMITE_CREDITO=PERMITE_CREDITO'
      'FANTASIA=FANTASIA'
      'DFIXAS=DFIXAS'
      'TIPOCALCULO=TIPOCALCULO'
      'SERIE_CTE=SERIE_CTE'
      'SEQ_CTE=SEQ_CTE'
      'COD_PAIS=COD_PAIS'
      'PAIS=PAIS')
    DataSet = frmModulo.qrfilial
    BCDToCurrency = False
    Left = 472
    Top = 323
  end
  object frImpForma: TfrxDBDataset
    UserName = 'frImpForma'
    CloseDataSource = False
    FieldAliases.Strings = (
      'COD_FORMA=COD_FORMA'
      'FORMA_DESCRICAO=FORMA_DESCRICAO'
      'FORMA_PRESTACAO=FORMA_PRESTACAO'
      'FORMA_VALOR=FORMA_VALOR')
    DataSet = qrImpForma
    BCDToCurrency = False
    Left = 656
    Top = 275
  end
  object qrImpForma: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      ' cfo.codigo cod_forma,'
      ' cfo.forma forma_descricao,'
      ' cfo.prestacao forma_prestacao,'
      ' cfo.valor forma_valor'
      'from'
      '  cupom_forma cfo'
      'where'
      '  cfo.cod_cupom = :COD'
      ''
      '')
    Left = 560
    Top = 403
    ParamData = <
      item
        DataType = ftString
        Name = 'cod'
        ParamType = ptInput
        Value = '099000000901048'
      end>
    object qrImpFormaCOD_FORMA: TStringField
      FieldName = 'COD_FORMA'
      Required = True
      Size = 50
    end
    object qrImpFormaFORMA_DESCRICAO: TStringField
      FieldName = 'FORMA_DESCRICAO'
      Required = True
      Size = 30
    end
    object qrImpFormaFORMA_PRESTACAO: TIntegerField
      FieldName = 'FORMA_PRESTACAO'
      Required = True
    end
    object qrImpFormaFORMA_VALOR: TFloatField
      FieldName = 'FORMA_VALOR'
    end
  end
  object RelCupomNFRestaurante: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 43331.805278750000000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      '  if <frCupomNF."SENHA"> > 0 then'
      '    cpSenha.Visible := True'
      '  else                  '
      '    cpSenha.Visible := False;  '
      '  if <frCupomNF."CPF_CONSUMIDOR"> = '#39#39' then'
      '    cpConsumidor.Visible := False;'
      '  if trim(<frObs."OBS">) <> '#39#39' then begin             '
      '    cpO1.Visible := True;  '
      '    cpO2.Visible := True;  '
      '  end else begin                               '
      '    cpO1.Visible := False;'
      '    cpO2.Visible := False;'
      '  end;  '
      'end.')
    OnBeforePrint = RelCupomNFRestauranteBeforePrint
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 224
    Top = 299
    Datasets = <
      item
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end
      item
        DataSet = frImpForma
        DataSetName = 'frImpForma'
      end
      item
        DataSet = frOBS
        DataSetName = 'frObs'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 100000.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object ReportTitle1: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 186.669450000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        Child = RelCupomNFRestaurante.cpConsumidor
        PrintChildIfInvisible = True
        Stretched = True
        object Memo1: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 93.267780000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'COMPROVANTE DE VENDA')
          ParentFont = False
        end
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 112.724490000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: [frCupomNF."DATA"] Hora: [frCupomNF."HORA"]')
          ParentFont = False
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000001000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
        object lbDocEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 23.677180000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CNPJ"]')
          ParentFont = False
        end
        object Line1: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 89.708720000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 129.551330000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 164.756030000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CLIENTE')
          ParentFont = False
        end
        object Memo11: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 164.756030000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 164.756030000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."CLIENTE"]')
          ParentFont = False
          WordWrap = False
        end
        object Memo28: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 129.551330000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 129.551330000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NUMERO"]')
          ParentFont = False
        end
        object Line3: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000001000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo2: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 39.354360000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."ENDERECO"]')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 54.913420000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CIDADE"]/[frFilial."UF"] - [frFilial."CEP"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 70.370130000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."TELEFONE"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end
            item
            end>
        end
        object Memo6: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 147.401670000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CAIXA')
          ParentFont = False
        end
        object Memo14: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 147.401670000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo15: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 147.401670000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."COD_CAIXA"]')
          ParentFont = False
        end
        object Line7: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 184.771800000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object MasterData1: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 34.015770000000000000
        Top = 453.543600000000000000
        Width = 272.126160000000000000
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
        RowCount = 0
        object Memo18: TfrxMemoView
          AllowVectorExport = True
          Width = 268.346630000000000000
          Height = 15.118120000000000000
          DataField = 'NOME'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NOME"]')
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 15.118119999999980000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          DataField = 'QTDE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."QTDE"]')
          ParentFont = False
        end
        object Memo20: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 15.118119999999980000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          DataField = 'UNIDADE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."UNIDADE"]')
          ParentFont = False
        end
        object Memo21: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 15.118119999999980000
          Width = 52.913420000000000000
          Height = 15.118120000000000000
          DataField = 'ITEM_UNITARIO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_UNITARIO"]')
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 15.118119999999980000
          Width = 60.472480000000000000
          Height = 15.118120000000000000
          DataField = 'ITEM_TOTAL'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_TOTAL"]')
          ParentFont = False
        end
        object Memo23: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          DataField = 'ITEM_DESCONTO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_DESCONTO"]')
          ParentFont = False
        end
        object Memo25: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          DataField = 'ITEM_ACRESCIMO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_ACRESCIMO"]')
          ParentFont = False
        end
      end
      object GroupHeader1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Top = 430.866420000000000000
        Width = 272.126160000000000000
        Condition = 'frCupomNF."COD_CUPOM"'
      end
      object MasterData3: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 18.897650000000000000
        Top = 695.433520000000000000
        Width = 272.126160000000000000
        Child = RelCupomNFRestaurante.Child3
        DataSet = frImpForma
        DataSetName = 'frImpForma'
        RowCount = 0
        object Memo26: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataField = 'FORMA_DESCRICAO'
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frImpForma."FORMA_DESCRICAO"]')
          ParentFont = False
        end
        object Memo27: TfrxMemoView
          AllowVectorExport = True
          Left = 140.063080000000000000
          Top = 1.779530000000022000
          Width = 56.692950000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frImpForma."FORMA_PRESTACAO"] X')
          ParentFont = False
        end
        object Memo29: TfrxMemoView
          AllowVectorExport = True
          Left = 200.315090000000000000
          Top = 1.779530000000022000
          Width = 68.031540000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[<frImpForma."FORMA_VALOR">/<frImpForma."FORMA_PRESTACAO">]')
          ParentFont = False
        end
      end
      object cpTotal: TfrxGroupFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 115.944960000000000000
        Top = 510.236550000000000000
        Width = 272.126160000000000000
        Child = RelCupomNFRestaurante.Child2
        object ltotal: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 59.252009999999990000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Total')
          ParentFont = False
        end
        object ptotal: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 59.252009999999990000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtotal: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 59.252009999999990000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TOTAL"]')
          ParentFont = False
        end
        object lpago: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 78.149659999999870000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Pago')
          ParentFont = False
        end
        object ppago: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 78.149659999999870000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vpago: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 78.149659999999870000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_PAGO"]')
          ParentFont = False
        end
        object ltroco: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 97.047309999999990000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Troco')
          ParentFont = False
        end
        object ptroco: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 97.047309999999990000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtroco: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 97.047309999999990000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TROCO"]')
          ParentFont = False
        end
        object Line5: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000079000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object lsub: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 3.779530000000079000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Sub Total')
          ParentFont = False
        end
        object psub: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 3.779530000000079000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vsub: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 3.779530000000079000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<frCupomNF."ITEM_TOTAL">,MasterData1)]')
          ParentFont = False
        end
        object lcouvert: TfrxMemoView
          AllowVectorExport = True
          Left = 71.811070000000000000
          Top = 21.677179999999960000
          Width = 90.708720000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Couvert Art'#237'stico')
          ParentFont = False
        end
        object pcouvert: TfrxMemoView
          AllowVectorExport = True
          Left = 164.519790000000000000
          Top = 21.677179999999960000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object cpCouvert: TfrxMemoView
          AllowVectorExport = True
          Left = 174.842610000000000000
          Top = 21.677179999999960000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."COUVERT"]')
          ParentFont = False
        end
        object lTaxa: TfrxMemoView
          AllowVectorExport = True
          Left = 71.811070000000000000
          Top = 40.795300000000050000
          Width = 90.708720000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Taxa Gar'#231'om')
          ParentFont = False
        end
        object ptaxa: TfrxMemoView
          AllowVectorExport = True
          Left = 164.519790000000000000
          Top = 40.795300000000050000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object cptaxa: TfrxMemoView
          AllowVectorExport = True
          Left = 174.842610000000000000
          Top = 40.795300000000050000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."TAXA_GARCON"]')
          ParentFont = False
        end
      end
      object Child2: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -9
        Font.Name = 'Arial'
        Font.Style = []
        Height = 22.677180000000000000
        ParentFont = False
        Top = 650.079160000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo45: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Forma de Pagamento:')
          ParentFont = False
        end
      end
      object Child3: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 737.008350000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object cpSenha: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 0.779530000000022500
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Senha de Retirada: [frCupomNF."SENHA"]')
          ParentFont = False
        end
        object Line6: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object Header1: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 41.574830000000000000
        Top = 366.614410000000000000
        Width = 272.126160000000000000
        object Line4: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 37.220469999999970000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo8: TfrxMemoView
          AllowVectorExport = True
          Top = 1.307050000000004000
          Width = 45.354360000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Produto')
          ParentFont = False
        end
        object Memo9: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 16.425169999999980000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Qtde')
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 16.425169999999980000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Un.')
          ParentFont = False
        end
        object Memo13: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 16.425169999999980000
          Width = 52.913420000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Unit.')
          ParentFont = False
        end
        object Memo16: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 16.425169999999980000
          Width = 60.472480000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Total')
          ParentFont = False
        end
        object Memo17: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 16.425169999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Desc.')
          ParentFont = False
        end
        object Memo24: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 16.425169999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Acr'#233's.')
          ParentFont = False
        end
      end
      object cpConsumidor: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 79.370130000000000000
        Top = 226.771800000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo34: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 0.897649999999998700
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CONSUMIDOR')
          ParentFont = False
        end
        object Memo35: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 21.574829999999990000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Nome: [frCupomNF."NOME_CONSUMIDOR"]')
          ParentFont = False
        end
        object Line8: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 74.590600000000030000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo36: TfrxMemoView
          AllowVectorExport = True
          Top = 37.795300000000030000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CPF: [frCupomNF."CPF_CONSUMIDOR"]')
          ParentFont = False
        end
        object Memo37: TfrxMemoView
          AllowVectorExport = True
          Top = 52.913419999999890000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'End: [frCupomNF."ENDERECO_CONSUMIDOR"]')
          ParentFont = False
        end
      end
      object cpO1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 782.362710000000000000
        Width = 272.126160000000000000
        Condition = '1=1'
        object Memo39: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Observa'#231#227'o:')
          ParentFont = False
        end
        object Line2: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object cpO2: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 827.717070000000000000
        Width = 272.126160000000000000
        DataSet = frOBS
        DataSetName = 'frObs'
        RowCount = 0
        Stretched = True
        object Memo38: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 3.779530000000022000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          DataField = 'obs'
          DataSet = frOBS
          DataSetName = 'frObs'
          Frame.Typ = []
          Memo.UTF8W = (
            '[frObs."obs"]')
        end
      end
    end
  end
  object qrGrupo: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from grupo')
    Left = 688
    Top = 339
    object qrGrupoCODIGO: TStringField
      FieldName = 'CODIGO'
      Required = True
      Size = 6
    end
    object qrGrupoGRUPO: TStringField
      FieldName = 'GRUPO'
      Size = 30
    end
    object qrGrupoFOTO: TBlobField
      FieldName = 'FOTO'
    end
  end
  object qrProduto: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      
        'select pro.codigo, pro.cod_barra, pro.nome, pro.fotobd from ESTO' +
        'QUE pro where pro.grupo = :grupo')
    Left = 592
    Top = 371
    ParamData = <
      item
        DataType = ftString
        Name = 'grupo'
        ParamType = ptInput
        Value = nil
      end>
    object qrProdutoCODIGO: TIntegerField
      FieldName = 'CODIGO'
      Required = True
    end
    object qrProdutoCOD_BARRA: TStringField
      FieldName = 'COD_BARRA'
      Size = 15
    end
    object qrProdutoNOME: TStringField
      FieldName = 'NOME'
      Required = True
      Size = 80
    end
    object qrProdutoFOTOBD: TBlobField
      FieldName = 'FOTOBD'
    end
  end
  object frComanda: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Microsoft Print to PDF'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 43371.578354930600000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      '  if <frCupomNF."SENHA"> > 0 then'
      '    cpSenha.Visible := True'
      '  else                  '
      '    cpSenha.Visible := False;'
      '  if trim(<frObs."OBS">) <> '#39#39' then begin             '
      '    cpO1.Visible := True;  '
      '    cpO2.Visible := True;  '
      '  end else begin                               '
      '    cpO1.Visible := False;'
      '    cpO2.Visible := False;'
      '  end;                    '
      '  '
      '  if <frCupomNF."PEGO_PORQUEM"> = '#39#39' then begin'
      '    cpPego.Visible := False;'
      
        '    cpTitulo.Child := cpConsumidor;                             ' +
        '                 '
      '  end;'
      '  if <frCupomNF."CPF_CONSUMIDOR"> = '#39#39' then'
      '    cpConsumidor.Visible := False;                          '
      '                                              '
      'end.')
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 64
    Top = 299
    Datasets = <
      item
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end
      item
        DataSet = frImpForma
        DataSetName = 'frImpForma'
      end
      item
        DataSet = frOBS
        DataSetName = 'frObs'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 100000.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object cpTitulo: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 188.976500000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        Child = frComanda.cpPego
        PrintChildIfInvisible = True
        Stretched = True
        object Memo1: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 93.267780000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'COMPROVANTE DE VENDA')
          ParentFont = False
        end
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 112.724490000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: [frCupomNF."DATA"] Hora: [frCupomNF."HORA"]')
          ParentFont = False
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000001000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
        object lbDocEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 23.677180000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CNPJ"]')
          ParentFont = False
        end
        object Line1: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 89.708720000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 129.551330000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 164.756030000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CLIENTE')
          ParentFont = False
        end
        object Line2: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 183.653680000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo11: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 164.756030000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 164.756030000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."CLIENTE"]')
          ParentFont = False
          WordWrap = False
        end
        object Memo28: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 129.551330000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 129.551330000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NUMERO"]')
          ParentFont = False
        end
        object Line3: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000001000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo2: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 39.354360000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."ENDERECO"]')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 54.913420000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CIDADE"]/[frFilial."UF"] - [frFilial."CEP"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 70.370130000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."TELEFONE"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end
            item
            end>
        end
        object Memo6: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 147.401670000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CAIXA')
          ParentFont = False
        end
        object Memo14: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 147.401670000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo15: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 147.401670000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."COD_CAIXA"]')
          ParentFont = False
        end
      end
      object MasterData1: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 30.236240000000000000
        Top = 510.236550000000000000
        Width = 272.126160000000000000
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
        RowCount = 0
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 15.118119999999920000
          Width = 30.236240000000000000
          Height = 11.338590000000000000
          DataField = 'QTDE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."QTDE"]')
          ParentFont = False
        end
        object Memo20: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 15.118119999999920000
          Width = 22.677180000000000000
          Height = 11.338590000000000000
          DataField = 'UNIDADE'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."UNIDADE"]')
          ParentFont = False
        end
        object Memo21: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 15.118119999999920000
          Width = 52.913420000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_UNITARIO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_UNITARIO"]')
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 15.118119999999920000
          Width = 60.472480000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_TOTAL'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_TOTAL"]')
          ParentFont = False
        end
        object Memo23: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 15.118119999999920000
          Width = 41.574830000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_DESCONTO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_DESCONTO"]')
          ParentFont = False
        end
        object Memo25: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 15.118119999999920000
          Width = 41.574830000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM_ACRESCIMO'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_ACRESCIMO"]')
          ParentFont = False
        end
        object Memo18: TfrxMemoView
          AllowVectorExport = True
          Width = 22.677180000000000000
          Height = 11.338590000000000000
          DataField = 'ITEM'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."ITEM"]')
          ParentFont = False
        end
        object Memo32: TfrxMemoView
          AllowVectorExport = True
          Left = 25.236240000000000000
          Width = 71.811070000000000000
          Height = 11.338590000000000000
          DataField = 'COD_BARRA'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."COD_BARRA"]')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo33: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Width = 170.078850000000000000
          Height = 11.338590000000000000
          DataField = 'NOME'
          DataSet = frCupomNF
          DataSetName = 'frCupomNF'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NOME"]')
          ParentFont = False
        end
      end
      object GroupHeader1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Top = 487.559370000000000000
        Width = 272.126160000000000000
        Condition = 'frCupomNF."COD_CUPOM"'
      end
      object MasterData3: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 18.897650000000000000
        Top = 718.110700000000000000
        Width = 272.126160000000000000
        Child = frComanda.Child2
        DataSet = frImpForma
        DataSetName = 'frImpForma'
        RowCount = 0
        object Memo26: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataField = 'FORMA_DESCRICAO'
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frImpForma."FORMA_DESCRICAO"]')
          ParentFont = False
        end
        object Memo27: TfrxMemoView
          AllowVectorExport = True
          Left = 140.063080000000000000
          Top = 1.779530000000022000
          Width = 56.692950000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frImpForma."FORMA_PRESTACAO"] X')
          ParentFont = False
        end
        object Memo29: TfrxMemoView
          AllowVectorExport = True
          Left = 200.315090000000000000
          Top = 1.779530000000022000
          Width = 68.031540000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[<frImpForma."FORMA_VALOR">/<frImpForma."FORMA_PRESTACAO">]')
          ParentFont = False
        end
      end
      object cpTotal: TfrxGroupFooter
        FillType = ftBrush
        Frame.Typ = []
        Height = 89.488250000000000000
        Top = 563.149970000000100000
        Width = 272.126160000000000000
        Child = frComanda.Child1
        object ltotal: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 9.252010000000041000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Total')
          ParentFont = False
        end
        object ptotal: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 9.252010000000041000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtotal: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 9.252010000000041000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TOTAL"]')
          ParentFont = False
        end
        object lpago: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 46.149659999999920000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Pago')
          ParentFont = False
        end
        object ppago: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 46.149659999999920000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vpago: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 46.149659999999920000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_PAGO"]')
          ParentFont = False
        end
        object ltroco: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 65.047309999999690000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Troco')
          ParentFont = False
        end
        object ptroco: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 65.047309999999690000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object vtroco: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 65.047309999999690000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."CUPOM_TROCO"]')
          ParentFont = False
        end
        object Line5: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000136000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo40: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Top = 27.236239999999950000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor Desc.')
          ParentFont = False
        end
        object Memo41: TfrxMemoView
          AllowVectorExport = True
          Left = 164.488250000000000000
          Top = 27.236239999999950000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object Memo42: TfrxMemoView
          AllowVectorExport = True
          Left = 174.811070000000000000
          Top = 27.236239999999950000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[SUM(<frCupomNF."ITEM_DESCONTO">,MasterData1)]')
          ParentFont = False
        end
      end
      object Child1: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 18.897650000000000000
        Top = 676.535870000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo45: TfrxMemoView
          AllowVectorExport = True
          Top = 1.779530000000022000
          Width = 136.063080000000000000
          Height = 15.118120000000000000
          DataSet = frImpForma
          DataSetName = 'frImpForma'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Forma de Pagamento:')
          ParentFont = False
        end
      end
      object Child2: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 759.685530000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object cpSenha: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 3.000000000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Senha de Retirada: [frCupomNF."SENHA"]')
          ParentFont = False
        end
        object Line6: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object Header1: TfrxHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 34.015770000000000000
        Top = 430.866420000000000000
        Width = 272.126160000000000000
        object Line4: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 32.133889999999890000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo8: TfrxMemoView
          AllowVectorExport = True
          Width = 26.456710000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Item')
          ParentFont = False
        end
        object Memo9: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Top = 15.118119999999980000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Qtde')
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          AllowVectorExport = True
          Left = 36.795300000000000000
          Top = 15.118119999999980000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Un.')
          ParentFont = False
        end
        object Memo13: TfrxMemoView
          AllowVectorExport = True
          Left = 63.031540000000000000
          Top = 15.118119999999980000
          Width = 52.913420000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Unit.')
          ParentFont = False
        end
        object Memo16: TfrxMemoView
          AllowVectorExport = True
          Left = 207.874150000000000000
          Top = 15.118119999999980000
          Width = 60.472480000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Vlr Total')
          ParentFont = False
        end
        object Memo17: TfrxMemoView
          AllowVectorExport = True
          Left = 117.944960000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Desc.')
          ParentFont = False
        end
        object Memo24: TfrxMemoView
          AllowVectorExport = True
          Left = 162.740260000000000000
          Top = 15.118119999999980000
          Width = 41.574830000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'Acr'#233's.')
          ParentFont = False
        end
        object Memo30: TfrxMemoView
          AllowVectorExport = True
          Left = 25.236240000000000000
          Width = 71.811070000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'C'#243'digo')
          ParentFont = False
          VAlign = vaCenter
        end
        object Memo31: TfrxMemoView
          AllowVectorExport = True
          Left = 98.267780000000000000
          Width = 45.354360000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Produto')
          ParentFont = False
        end
      end
      object cpPego: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 37.795300000000000000
        Top = 230.551330000000000000
        Width = 272.126160000000000000
        Child = frComanda.cpConsumidor
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo10: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 0.897649999999998700
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'PEGO POR QUEM')
          ParentFont = False
        end
        object lbPegoPor: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 14.015770000000000000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."PEGO_PORQUEM"]')
          ParentFont = False
        end
        object Line7: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 35.015769999999970000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
      object cpConsumidor: TfrxChild
        FillType = ftBrush
        Frame.Typ = []
        Height = 79.370130000000000000
        Top = 291.023810000000000000
        Width = 272.126160000000000000
        ToNRows = 0
        ToNRowsMode = rmCount
        object Memo34: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 0.897649999999998700
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CONSUMIDOR')
          ParentFont = False
        end
        object Memo35: TfrxMemoView
          AllowVectorExport = True
          Left = 2.000000000000000000
          Top = 21.574830000000020000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Nome: [frCupomNF."NOME_CONSUMIDOR"]')
          ParentFont = False
        end
        object Line8: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 74.590600000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo36: TfrxMemoView
          AllowVectorExport = True
          Top = 37.795300000000000000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CPF: [frCupomNF."CPF_CONSUMIDOR"]')
          ParentFont = False
        end
        object Memo37: TfrxMemoView
          AllowVectorExport = True
          Top = 52.913419999999970000
          Width = 268.346630000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'End: [frCupomNF."ENDERECO_CONSUMIDOR"]')
          ParentFont = False
        end
      end
      object cpO2: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 850.394250000000000000
        Width = 272.126160000000000000
        DataSet = frOBS
        DataSetName = 'frObs'
        RowCount = 0
        Stretched = True
        object Memo38: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 3.779530000000022000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          DataField = 'obs'
          DataSet = frOBS
          DataSetName = 'frObs'
          Frame.Typ = []
          Memo.UTF8W = (
            '[frObs."obs"]')
        end
      end
      object cpO1: TfrxGroupHeader
        FillType = ftBrush
        Frame.Typ = []
        Height = 22.677180000000000000
        Top = 805.039890000000000000
        Width = 272.126160000000000000
        Condition = '1=1'
        object Memo39: TfrxMemoView
          AllowVectorExport = True
          Left = 2.779530000000000000
          Width = 94.488250000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'Observa'#231#227'o:')
          ParentFont = False
        end
        object Line9: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
      end
    end
  end
  object frxComanda: TfrxDBDataset
    UserName = 'frxComanda'
    CloseDataSource = False
    FieldAliases.Strings = (
      'SENHA=SENHA'
      'DATA=DATA'
      'HORA=HORA'
      'COD_PRODUTO=COD_PRODUTO'
      'NOME=NOME'
      'QTDE=QTDE'
      'VALOR_UNITARIO=VALOR_UNITARIO'
      'VALOR_TOTAL=VALOR_TOTAL'
      'EXTRA=EXTRA')
    DataSet = qrComanda
    BCDToCurrency = False
    Left = 624
    Top = 275
  end
  object qrComanda: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      '  cup.senha,'
      '  cup.data,'
      '  cup.hora,'
      '  ite.cod_produto,'
      '  case when trim(coalesce(ite.complemento, '#39#39')) = '#39#39' then'
      '    pro.nome'
      '  else'
      '   ite.complemento'
      '  end nome,'
      '  ite.qtde,'
      '  ite.valor_unitario,'
      '  ite.valor_total,'
      '  ite.extra'
      'from'
      '  cupom cup'
      '  inner join cupom_item ite on ite.cod_cupom = cup.codigo'
      '  inner join estoque pro on pro.codigo = ite.cod_produto'
      'where cup.codigo = :cupom'
      '')
    Left = 384
    Top = 467
    ParamData = <
      item
        DataType = ftString
        Name = 'cupom'
        ParamType = ptInput
        Value = nil
      end>
    object qrComandaSENHA: TIntegerField
      FieldName = 'SENHA'
    end
    object qrComandaDATA: TDateField
      FieldName = 'DATA'
      Required = True
    end
    object qrComandaHORA: TTimeField
      FieldName = 'HORA'
      Required = True
    end
    object qrComandaCOD_PRODUTO: TIntegerField
      FieldName = 'COD_PRODUTO'
      ReadOnly = True
      Required = True
    end
    object qrComandaNOME: TStringField
      FieldName = 'NOME'
      ReadOnly = True
      Required = True
      Size = 80
    end
    object qrComandaQTDE: TFloatField
      FieldName = 'QTDE'
      ReadOnly = True
    end
    object qrComandaVALOR_UNITARIO: TFloatField
      FieldName = 'VALOR_UNITARIO'
      ReadOnly = True
    end
    object qrComandaVALOR_TOTAL: TFloatField
      FieldName = 'VALOR_TOTAL'
      ReadOnly = True
    end
    object qrComandaEXTRA: TBlobField
      FieldName = 'EXTRA'
      ReadOnly = True
    end
  end
  object qrTmp: TUniQuery
    Connection = frmModulo.conexao
    Left = 656
    Top = 371
  end
  object cdsRecebimento: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 1089
    Top = 305
    object cdsRecebimentoTipo: TStringField
      FieldName = 'Tipo'
      Size = 30
    end
    object cdsRecebimentoParcelas: TIntegerField
      FieldName = 'Parcelas'
    end
    object cdsRecebimentobandeira: TStringField
      FieldName = 'bandeira'
    end
    object cdsRecebimentovalor: TFloatField
      FieldName = 'valor'
    end
    object cdsRecebimentocliente: TStringField
      FieldName = 'cliente'
      Size = 6
    end
    object cdsRecebimentoconvenio: TStringField
      FieldName = 'convenio'
      Size = 6
    end
    object cdsRecebimentobanco: TIntegerField
      FieldName = 'banco'
    end
    object cdsRecebimentoCupomCreditoID: TIntegerField
      FieldName = 'CupomCreditoID'
    end
    object cdsRecebimentoCupomCreditoCupom: TStringField
      FieldName = 'CupomCreditoCupom'
      Size = 30
    end
    object cdsRecebimentoCupomCreditoValor: TFloatField
      FieldName = 'CupomCreditoValor'
    end
    object cdsRecebimentotef_cnpj: TStringField
      FieldName = 'tef_cnpj'
    end
    object cdsRecebimentotef_codigo_autorizacao: TStringField
      FieldName = 'tef_codigo_autorizacao'
      Size = 50
    end
    object cdsRecebimentocod_adm_sat: TStringField
      FieldName = 'cod_adm_sat'
      Size = 3
    end
    object cdsRecebimentoRec1: TStringField
      FieldName = 'Rec1'
      Size = 8196
    end
    object cdsRecebimentoRec2: TStringField
      FieldName = 'Rec2'
      Size = 8196
    end
    object cdsRecebimentoRecTEF: TBooleanField
      FieldName = 'RecTEF'
    end
  end
  object TimerExtra: TTimer
    Enabled = False
    Interval = 1500
    OnTimer = TimerExtraTimer
    Left = 848
    Top = 224
  end
  object cdsOBS: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'obs'
        DataType = ftMemo
        Size = 10
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 1112
    Top = 131
    object cdsOBSobs: TMemoField
      FieldName = 'obs'
      BlobType = ftMemo
      Size = 10
    end
  end
  object frOBS: TfrxDBDataset
    UserName = 'frObs'
    CloseDataSource = False
    FieldAliases.Strings = (
      'obs=obs')
    DataSet = cdsOBS
    BCDToCurrency = False
    Left = 384
    Top = 387
  end
  object frxDesigner1: TfrxDesigner
    DefaultScriptLanguage = 'PascalScript'
    DefaultFont.Charset = DEFAULT_CHARSET
    DefaultFont.Color = clWindowText
    DefaultFont.Height = -13
    DefaultFont.Name = 'Arial'
    DefaultFont.Style = []
    DefaultLeftMargin = 10.000000000000000000
    DefaultRightMargin = 10.000000000000000000
    DefaultTopMargin = 10.000000000000000000
    DefaultBottomMargin = 10.000000000000000000
    DefaultPaperSize = 9
    DefaultOrientation = poPortrait
    GradientEnd = 11982554
    GradientStart = clWindow
    TemplatesExt = 'fr3'
    Restrictions = []
    RTLLanguage = False
    MemoParentFont = False
    OnGetTemplateList = frxDesigner1GetTemplateList
    Left = 516
    Top = 364
  end
  object RelFicha: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 43342.990141782410000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      'end.')
    OnBeforePrint = RelFichaBeforePrint
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 152
    Top = 379
    Datasets = <
      item
        DataSet = frCupomNF
        DataSetName = 'frCupomNF'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end
      item
        DataSet = frImpForma
        DataSetName = 'frImpForma'
      end
      item
        DataSet = frOBS
        DataSetName = 'frObs'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 100000.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object cpTitulo: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 200.315090000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        PrintChildIfInvisible = True
        Stretched = True
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 42.692950000000010000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: [frCupomNF."DATA"] Hora: [frCupomNF."HORA"]')
          ParentFont = False
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000001000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
        object Line1: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 40.574830000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object lbNCupum: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 61.519790000000000000
          Width = 52.913420000000000000
          Height = 15.118120000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Line2: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          ShiftMode = smWhenOverlapped
          Top = 82.047310000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object lbPCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 56.708720000000000000
          Top = 61.519790000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCupom: TfrxMemoView
          AllowVectorExport = True
          Left = 68.031540000000000000
          Top = 61.519790000000000000
          Width = 173.858380000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCupomNF."NUMERO"]')
          ParentFont = False
          WordWrap = False
        end
        object Line3: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000001000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object lbTel: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 24.015770000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."TELEFONE"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end
            item
            end>
        end
        object lbProduto: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 103.267780000000000000
          Width = 272.126160000000000000
          Height = 30.236240000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -21
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCupomNF."NOME"]')
          ParentFont = False
        end
        object lbPreco: TfrxMemoView
          Align = baRight
          AllowVectorExport = True
          Left = 90.708720000000000000
          Top = 137.283550000000000000
          Width = 181.417440000000000000
          Height = 30.236240000000000000
          DisplayFormat.FormatStr = '%2.2m'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -21
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCupomNF."ITEM_TOTAL"]')
          ParentFont = False
        end
        object Line4: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 177.637910000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object lbValidador: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 179.858380000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          ParentFont = False
        end
      end
    end
  end
  object ACBrSATExtratoFortes1: TACBrSATExtratoFortes
    Sistema = 'SOSPDV - 1198228-2649 - SOSSolutions.com.br'
    MargemInferior = 4.000000000000000000
    MargemSuperior = 2.000000000000000000
    MargemEsquerda = 2.000000000000000000
    MargemDireita = 2.000000000000000000
    ExpandeLogoMarcaConfig.Altura = 0
    ExpandeLogoMarcaConfig.Esquerda = 0
    ExpandeLogoMarcaConfig.Topo = 0
    ExpandeLogoMarcaConfig.Largura = 0
    ExpandeLogoMarcaConfig.Dimensionar = False
    ExpandeLogoMarcaConfig.Esticar = True
    CasasDecimais.Formato = tdetInteger
    CasasDecimais.qCom = 2
    CasasDecimais.vUnCom = 2
    CasasDecimais.MaskqCom = ',0.00'
    CasasDecimais.MaskvUnCom = ',0.00'
    ACBrSAT = ACBrSAT1
    MsgAppQRCode = 
      'Consulte o QR Code pelo aplicativo  "De olho na nota", dispon'#237've' +
      'l na AppStore (Apple) e PlayStore (Android)'
    ImprimeEmUmaLinha = False
    FormularioContinuo = True
    Left = 142
    Top = 482
  end
  object ACBrSAT1: TACBrSAT
    Extrato = ACBrSATExtratoFortes1
    NomeDLL = 'c:\Sat\SAT.DLL'
    Config.infCFe_versaoDadosEnt = 0.050000000000000000
    Config.ide_numeroCaixa = 0
    Config.ide_tpAmb = taHomologacao
    Config.emit_cRegTrib = RTSimplesNacional
    Config.emit_cRegTribISSQN = RTISSMicroempresaMunicipal
    Config.emit_indRatISSQN = irSim
    Config.EhUTF8 = True
    Config.PaginaDeCodigo = 65001
    Config.XmlSignLib = xsNone
    ConfigArquivos.SalvarCFe = True
    ConfigArquivos.SalvarCFeCanc = True
    ConfigArquivos.SalvarEnvio = True
    ConfigArquivos.SepararPorMes = True
    ConfigArquivos.PastaCFeVenda = 'C:\SOS\PDV\CFesPDF\Envio'
    ConfigArquivos.PastaCFeCancelamento = 'C:\SOS\PDV\CFesPDF\Can'
    ConfigArquivos.PastaEnvio = 'C:\SOS\PDV\CFesPDF\Vd'
    ConfigArquivos.PrefixoArqCFe = 'AD'
    ConfigArquivos.PrefixoArqCFeCanc = 'ADC'
    Rede.tipoInter = infETHE
    Rede.seg = segNONE
    Rede.tipoLan = lanDHCP
    Rede.proxy = 0
    Rede.proxy_porta = 0
    OnGetcodigoDeAtivacao = ACBrSAT1GetcodigoDeAtivacao
    OnGetsignAC = ACBrSAT1GetsignAC
    Left = 230
    Top = 482
  end
  object qrCupom: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from cupom')
    Left = 209
    Top = 574
  end
  object QrCupom_Item: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from cupom_item')
    Left = 241
    Top = 574
  end
  object qrcupom_forma: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from cupom_forma')
    Left = 273
    Top = 574
  end
  object Query1: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from cupom')
    Left = 161
    Top = 582
  end
  object QrProdutoSAT: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select * from estoque')
    Left = 344
    Top = 568
  end
  object qrCombo: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select combo from estoque where codigo = :cod')
    Left = 432
    Top = 227
    ParamData = <
      item
        DataType = ftInteger
        Name = 'cod'
        ParamType = ptInput
        Value = nil
      end>
    object qrComboCOMBO: TStringField
      FieldName = 'COMBO'
      Size = 1
    end
  end
  object qrProdComb: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      '  pro.codigo,'
      '  pro.NOME produto,'
      '  pro.unidade,'
      '  comp.qtde'
      'from'
      '  ESTOQUE_COMBO comp'
      '  inner join ESTOQUE pro on pro.codigo = comp.codproduto'
      'where comp.produto = :pro')
    Left = 480
    Top = 259
    ParamData = <
      item
        DataType = ftString
        Name = 'pro'
        ParamType = ptInput
        Value = nil
      end>
    object qrProdCombCODIGO: TIntegerField
      FieldName = 'CODIGO'
      Required = True
    end
    object qrProdCombPRODUTO: TStringField
      FieldName = 'PRODUTO'
      Required = True
      Size = 80
    end
    object qrProdCombUNIDADE: TStringField
      FieldName = 'UNIDADE'
      Size = 5
    end
    object qrProdCombQTDE: TFloatField
      FieldName = 'QTDE'
      ReadOnly = True
    end
  end
  object qrLancaProdutos: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      '  *'
      'from'
      '  cupom_item_sub'
      'where 1=2')
    CachedUpdates = True
    Left = 1008
    Top = 198
    object qrLancaProdutosCOD_ITEM: TStringField
      FieldName = 'COD_ITEM'
      Required = True
      Size = 50
    end
    object qrLancaProdutosCODPRO: TStringField
      FieldName = 'CODPRO'
      Size = 10
    end
    object qrLancaProdutosQTDE: TFloatField
      FieldName = 'QTDE'
    end
  end
  object cdsLancaProdutos: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 744
    Top = 128
    object cdsLancaProdutosCOD_ITEM: TStringField
      FieldName = 'COD_ITEM'
      Size = 50
    end
    object cdsLancaProdutosCODPRO: TStringField
      FieldName = 'CODPRO'
      Size = 10
    end
    object cdsLancaProdutosQTDE: TFloatField
      FieldName = 'QTDE'
    end
  end
  object frImpressao: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    ReportOptions.CreateDate = 39189.682060544000000000
    ReportOptions.LastChange = 43425.420103796300000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      ''
      'begin'
      ''
      'end.')
    OnBeforePrint = frComprovanteDebitoBeforePrint
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 32
    Top = 531
    Datasets = <
      item
        DataSet = frCarne
        DataSetName = 'frCarne'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end>
    Variables = <>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 209.000000000000000000
      PaperHeight = 291.000000000000000000
      PaperSize = 256
      TopMargin = 5.000000000000000000
      Frame.Typ = []
      MirrorMode = []
      object MasterData1: TfrxMasterData
        FillType = ftBrush
        Frame.Typ = []
        Height = 241.330860000000000000
        Top = 18.897650000000000000
        Width = 789.921770000000000000
        DataSet = frCarne
        DataSetName = 'frCarne'
        RowCount = 0
        object Line1: TfrxLineView
          AllowVectorExport = True
          Left = 277.685220000000000000
          Top = 11.338590000000000000
          Height = 226.771800000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Line8: TfrxLineView
          AllowVectorExport = True
          Left = 3.779530000000000000
          Top = 241.330860000000000000
          Width = 789.921770000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Line4: TfrxLineView
          AllowVectorExport = True
          Left = 41.574803150000000000
          Top = 34.015770000000010000
          Width = 222.992125980000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Memo36: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 17.118120000000000000
          Width = 75.590600000000000000
          Height = 18.897650000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold, fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Recibo')
          ParentFont = False
        end
        object Memo35: TfrxMemoView
          AllowVectorExport = True
          Left = 177.637910000000000000
          Top = 21.897650000000000000
          Width = 90.708720000000000000
          Height = 15.118120000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '(Valores em Reais)')
          ParentFont = False
        end
        object Line6: TfrxLineView
          AllowVectorExport = True
          Left = 41.574803150000000000
          Top = 60.472480000000000000
          Width = 222.992125980000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Line7: TfrxLineView
          AllowVectorExport = True
          Left = 79.944960000000000000
          Top = 34.015770000000010000
          Height = 26.456710000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo1: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 34.015770000000010000
          Width = 34.015770000000010000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Parcela')
          ParentFont = False
        end
        object Memo2: TfrxMemoView
          AllowVectorExport = True
          Left = 169.756030000000000000
          Top = 34.015770000000010000
          Width = 71.811070000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Vencimento')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          AllowVectorExport = True
          Left = 46.354360000000000000
          Top = 45.354360000000000000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          DataField = 'PARCELA'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCarne."PARCELA"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          AllowVectorExport = True
          Left = 181.094620000000000000
          Top = 45.354360000000000000
          Width = 79.370130000000000000
          Height = 15.118120000000000000
          DataField = 'VENCIMENTO'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCarne."VENCIMENTO"]')
          ParentFont = False
        end
        object Line9: TfrxLineView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 86.929190000000000000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo37: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 68.031540000000010000
          Width = 83.149660000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor do Documento:')
          ParentFont = False
        end
        object Memo38: TfrxMemoView
          AllowVectorExport = True
          Left = 180.417440000000000000
          Top = 66.252010000000000000
          Width = 83.149660000000000000
          Height = 15.118120000000000000
          DataField = 'VALOR'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCarne."VALOR"]')
          ParentFont = False
        end
        object Line10: TfrxLineView
          AllowVectorExport = True
          Left = 166.976500000000000000
          Top = 34.015770000000010000
          Height = 26.456710000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo39: TfrxMemoView
          AllowVectorExport = True
          Left = 82.724463150000000000
          Top = 34.015770000000010000
          Width = 60.472480000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Emiss'#227'o')
          ParentFont = False
        end
        object Memo40: TfrxMemoView
          AllowVectorExport = True
          Left = 87.724490000000000000
          Top = 45.354360000000000000
          Width = 75.590600000000000000
          Height = 15.118120000000000000
          DataField = 'EMISSAO'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCarne."EMISSAO"]')
          ParentFont = False
        end
        object Line11: TfrxLineView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 113.385900000000000000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 94.488250000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(-) Desconto/Abatimento:')
          ParentFont = False
        end
        object Line12: TfrxLineView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 139.842610000000000000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo6: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 120.944960000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(+) Mora/Multa:')
          ParentFont = False
        end
        object Line13: TfrxLineView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 166.299320000000000000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 147.401670000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(=) Valor Cobrado:')
          ParentFont = False
        end
        object Memo8: TfrxMemoView
          AllowVectorExport = True
          Left = 181.417440000000000000
          Top = 196.535560000000000000
          Width = 83.149660000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold, fsItalic]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'VIA DO SACADO')
          ParentFont = False
        end
        object Line14: TfrxLineView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 192.756030000000000000
          Width = 222.992270000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo9: TfrxMemoView
          AllowVectorExport = True
          Left = 41.574830000000000000
          Top = 173.858380000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Data do Pagamento:')
          ParentFont = False
        end
        object Line16: TfrxLineView
          AllowVectorExport = True
          Left = 291.023810000000000000
          Top = 34.015770000000010000
          Width = 464.882190000000000000
          Color = clBlack
          Frame.Typ = []
          Diagonal = True
        end
        object Memo13: TfrxMemoView
          AllowVectorExport = True
          Left = 665.197280000000000000
          Top = 18.897650000000000000
          Width = 90.708720000000000000
          Height = 15.118120000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '(Valores em Reais)')
          ParentFont = False
        end
        object Line17: TfrxLineView
          AllowVectorExport = True
          Left = 510.236550000000000000
          Top = 88.708720000000000000
          Width = 245.669450000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo14: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 69.811070000000000000
          Width = 83.149660000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Valor do Documento:')
          ParentFont = False
        end
        object Memo15: TfrxMemoView
          AllowVectorExport = True
          Left = 667.976810000000000000
          Top = 68.031540000000010000
          Width = 83.149660000000000000
          Height = 15.118120000000000000
          DataField = 'VALOR'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          DisplayFormat.FormatStr = '%2.2n'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCarne."VALOR"]')
          ParentFont = False
        end
        object Line18: TfrxLineView
          AllowVectorExport = True
          Left = 510.236550000000000000
          Top = 115.165430000000000000
          Width = 245.669450000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo16: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 96.267780000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(-) Desconto/Abatimento:')
          ParentFont = False
        end
        object Line19: TfrxLineView
          AllowVectorExport = True
          Left = 510.236550000000000000
          Top = 141.622140000000000000
          Width = 245.669450000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo17: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 122.724490000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(+) Mora/Multa:')
          ParentFont = False
        end
        object Line20: TfrxLineView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 168.078850000000000000
          Width = 468.661720000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo18: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 149.181200000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '(=) Valor Cobrado:')
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          AllowVectorExport = True
          Left = 672.756340000000000000
          Top = 224.771800000000000000
          Width = 83.149660000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsBold, fsItalic]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            'VIA DA LOJA')
          ParentFont = False
        end
        object Line21: TfrxLineView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 194.535560000000000000
          Width = 468.661720000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo20: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 175.637910000000000000
          Width = 102.047310000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Data do Pagamento:')
          ParentFont = False
        end
        object Line22: TfrxLineView
          AllowVectorExport = True
          Left = 510.236550000000000000
          Top = 34.015770000000010000
          Height = 160.629938350000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo21: TfrxMemoView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 45.354360000000000000
          Width = 154.960730000000000000
          Height = 49.133890000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Pag'#225'vel na Loja'
            'Ap'#243's vencimento multa de 2%'
            'Juros de Mora: 0,33% ao dia')
          ParentFont = False
        end
        object Line23: TfrxLineView
          AllowVectorExport = True
          Left = 510.236550000000000000
          Top = 60.472480000000000000
          Width = 245.669450000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo41: TfrxMemoView
          AllowVectorExport = True
          Left = 514.016080000000000000
          Top = 41.574830000000000000
          Width = 83.149660000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Data do Vencimento:')
          ParentFont = False
        end
        object Memo42: TfrxMemoView
          AllowVectorExport = True
          Left = 684.094930000000000000
          Top = 41.574830000000000000
          Width = 71.811070000000000000
          Height = 15.118120000000000000
          DataField = 'VENCIMENTO'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haRight
          Memo.UTF8W = (
            '[frCarne."VENCIMENTO"]')
          ParentFont = False
        end
        object Line25: TfrxLineView
          AllowVectorExport = True
          Left = 353.748300000000000000
          Top = 169.078850000000000000
          Height = 26.456710000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo43: TfrxMemoView
          AllowVectorExport = True
          Left = 291.244280000000000000
          Top = 168.078850000000000000
          Width = 34.015770000000010000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            'Parcela')
          ParentFont = False
        end
        object Memo44: TfrxMemoView
          AllowVectorExport = True
          Left = 361.307360000000000000
          Top = 168.078850000000000000
          Width = 79.370130000000000000
          Height = 11.338590000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = [fsItalic]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data da Emiss'#227'o')
          ParentFont = False
        end
        object Memo46: TfrxMemoView
          AllowVectorExport = True
          Left = 298.023810000000000000
          Top = 177.637910000000000000
          Width = 22.677180000000000000
          Height = 15.118120000000000000
          DataField = 'PARCELA'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCarne."PARCELA"]')
          ParentFont = False
        end
        object Memo47: TfrxMemoView
          AllowVectorExport = True
          Left = 358.527830000000000000
          Top = 177.637910000000000000
          Width = 83.149660000000000000
          Height = 15.118120000000000000
          DataField = 'EMISSAO'
          DataSet = frCarne
          DataSetName = 'frCarne'
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCarne."EMISSAO"]')
          ParentFont = False
        end
        object Line26: TfrxLineView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 219.212740000000000000
          Width = 468.661720000000000000
          Color = clBlack
          Frame.Typ = []
          Frame.Width = 0.500000000000000000
          Diagonal = True
        end
        object Memo49: TfrxMemoView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 203.315090000000000000
          Width = 34.015770000000000000
          Height = 18.897650000000000000
          DisplayFormat.DecimalSeparator = ','
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -8
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            'Cliente:')
          ParentFont = False
        end
        object Memo50: TfrxMemoView
          AllowVectorExport = True
          Left = 321.260050000000000000
          Top = 200.315090000000000000
          Width = 419.527830000000000000
          Height = 18.897650000000000000
          DataSetName = 'fscontasreceber'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCarne."CODIGO"] - [frCarne."NOME"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end>
        end
        object fxemitenteFANTASIA: TfrxMemoView
          AllowVectorExport = True
          Left = 287.244280000000000000
          Top = 15.118120000000000000
          Width = 411.968770000000000000
          Height = 18.897650000000000000
          DataSetName = 'fxemitente'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -16
          Font.Name = 'Arial'
          Font.Style = [fsBold, fsItalic]
          Frame.Typ = []
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
      end
    end
  end
  object qrCarne: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      'select'
      
        '   ((Select  Count(*) From cupom_crediario  t1 Where t1.codigo <' +
        ' par.codigo and t1.cod_cupom = par.cod_cupom) + 1) as parcela,'
      '   par.data emissao,'
      '   par.vencimento,'
      '   par.valor,'
      '   cli.codigo,'
      '   cli.nome'
      'from'
      '  cupom_crediario par'
      '  inner join cliente cli on cli.codigo = par.cod_cliente'
      'where'
      '  par.cod_cupom = :cupom')
    Left = 40
    Top = 387
    ParamData = <
      item
        DataType = ftString
        Name = 'cupom'
        ParamType = ptInput
        Value = nil
      end>
    object qrCarnePARCELA: TLargeintField
      FieldName = 'PARCELA'
      ReadOnly = True
    end
    object qrCarneEMISSAO: TDateField
      FieldName = 'EMISSAO'
      Required = True
    end
    object qrCarneVENCIMENTO: TDateField
      FieldName = 'VENCIMENTO'
      Required = True
    end
    object qrCarneVALOR: TFloatField
      FieldName = 'VALOR'
    end
    object qrCarneCODIGO: TIntegerField
      FieldName = 'CODIGO'
      ReadOnly = True
      Required = True
    end
    object qrCarneNOME: TStringField
      FieldName = 'NOME'
      ReadOnly = True
      Required = True
      Size = 80
    end
  end
  object frCarne: TfrxDBDataset
    UserName = 'frCarne'
    CloseDataSource = False
    FieldAliases.Strings = (
      'PARCELA=PARCELA'
      'EMISSAO=EMISSAO'
      'VENCIMENTO=VENCIMENTO'
      'VALOR=VALOR'
      'CODIGO=CODIGO'
      'NOME=NOME')
    DataSet = qrCarne
    BCDToCurrency = False
    Left = 24
    Top = 579
  end
  object qrEstoque: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      
        'select pro.estoque, pro.csosn, pro.ncm, pro.cod_barra, pro.infor' +
        'ma_codigo_barra_xml, pro.cfop from ESTOQUE pro where pro.codigo ' +
        '= :codigo')
    Left = 784
    Top = 387
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'codigo'
        Value = nil
      end>
    object qrEstoqueESTOQUE: TFloatField
      FieldName = 'ESTOQUE'
    end
    object qrEstoqueCSOSN: TStringField
      FieldName = 'CSOSN'
      Size = 3
    end
    object qrEstoqueNCM: TStringField
      FieldName = 'NCM'
      Size = 25
    end
    object qrEstoqueCOD_BARRA: TStringField
      FieldName = 'COD_BARRA'
      Size = 15
    end
    object qrEstoqueINFORMA_CODIGO_BARRA_XML: TStringField
      FieldName = 'INFORMA_CODIGO_BARRA_XML'
      Size = 1
    end
    object qrEstoqueCFOP: TStringField
      FieldName = 'CFOP'
      Size = 4
    end
  end
  object frCredito: TfrxDBDataset
    UserName = 'frCredito'
    CloseDataSource = False
    FieldAliases.Strings = (
      'CUPOM_DEVOLUCAO=CUPOM_DEVOLUCAO'
      'DATA=DATA'
      'HORA=HORA'
      'VALOR=VALOR'
      'VALOR_UTILIZADO=VALOR_UTILIZADO'
      'NOME=NOME'
      'VENCIMENTO=VENCIMENTO')
    DataSet = qrImpCredito
    BCDToCurrency = False
    Left = 440
    Top = 539
  end
  object relCredito: TfrxReport
    Tag = 1
    Version = '6.3.7'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Microsoft Print to PDF'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 40401.475989294000000000
    ReportOptions.LastChange = 43462.481197569450000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      'end.')
    OnReportPrint = 'frxReportOnReportPrint'
    Left = 496
    Top = 539
    Datasets = <
      item
        DataSet = frCredito
        DataSetName = 'frCredito'
      end
      item
        DataSet = frFilial
        DataSetName = 'frFilial'
      end>
    Variables = <
      item
        Name = ' User'
        Value = Null
      end
      item
        Name = 'LinhasImpressas'
        Value = Null
      end>
    Style = <>
    object Data: TfrxDataPage
      Height = 1000.000000000000000000
      Width = 1000.000000000000000000
    end
    object Page1: TfrxReportPage
      PaperWidth = 72.000000000000000000
      PaperHeight = 100000.000000000000000000
      PaperSize = 256
      TopMargin = 0.500000000000000000
      Frame.Typ = []
      EndlessHeight = True
      LargeDesignHeight = True
      MirrorMode = []
      PrintIfEmpty = False
      OnBeforePrint = 'Page1OnBeforePrint'
      object cpTitulo: TfrxReportTitle
        FillType = ftBrush
        Frame.Typ = []
        Height = 317.480520000000000000
        Top = 18.897650000000000000
        Width = 272.126160000000000000
        OnBeforePrint = 'ReportTitle1OnBeforePrint'
        PrintChildIfInvisible = True
        Stretched = True
        object Memo1: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 93.267780000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'CUPOM DE CR'#201'DITO')
          ParentFont = False
        end
        object lbDataHora: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 112.724490000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Data: [frCredito."DATA"] Hora: [frCredito."HORA"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end>
        end
        object lbEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 4.779530000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."FILIAL"]')
          ParentFont = False
        end
        object lbDocEmpresa: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 23.677180000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CNPJ"]')
          ParentFont = False
        end
        object Line1: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 89.708720000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo5: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 129.551330000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CUPOM')
          ParentFont = False
        end
        object Memo7: TfrxMemoView
          AllowVectorExport = True
          Left = 1.338590000000000000
          Top = 146.756030000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          Memo.UTF8W = (
            'CLIENTE')
          ParentFont = False
        end
        object Line2: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 315.937230000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo11: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 146.756030000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object lbCliente: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 146.756030000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCredito."NOME"]')
          ParentFont = False
          WordWrap = False
        end
        object Memo28: TfrxMemoView
          AllowVectorExport = True
          Left = 68.047310000000000000
          Top = 129.551330000000000000
          Width = 7.559060000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            ':')
          ParentFont = False
        end
        object frxMemoView1: TfrxMemoView
          AllowVectorExport = True
          Left = 78.370130000000000000
          Top = 129.551330000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          Memo.UTF8W = (
            '[frCredito."CUPOM_DEVOLUCAO"]')
          ParentFont = False
        end
        object Line3: TfrxLineView
          Align = baWidth
          AllowVectorExport = True
          Top = 2.779530000000000000
          Width = 272.126160000000000000
          Color = clBlack
          Frame.Style = fsDot
          Frame.Typ = [ftTop]
        end
        object Memo2: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 39.354360000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."ENDERECO"]')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 54.913420000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."CIDADE"]/[frFilial."UF"] - [frFilial."CEP"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 70.370130000000000000
          Width = 272.126160000000000000
          Height = 15.118120000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = []
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frFilial."TELEFONE"]')
          ParentFont = False
          Formats = <
            item
            end
            item
            end
            item
            end>
        end
        object Memo6: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 192.756030000000000000
          Width = 272.126160000000000000
          Height = 26.456710000000000000
          DisplayFormat.FormatStr = '%2.2m'
          DisplayFormat.Kind = fkNumeric
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -21
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            '[frCredito."VALOR"]')
          ParentFont = False
          WordWrap = False
        end
        object Memo8: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 173.858380000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Valor do Cr'#233'dito')
          ParentFont = False
        end
        object Memo9: TfrxMemoView
          Align = baWidth
          AllowVectorExport = True
          Top = 222.992270000000000000
          Width = 272.126160000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -9
          Font.Name = 'Arial'
          Font.Style = [fsBold]
          Frame.Typ = []
          HAlign = haCenter
          Memo.UTF8W = (
            'Vencimento: [frCredito."VENCIMENTO"]')
          ParentFont = False
        end
        object BarCode1: TfrxBarCodeView
          Align = baCenter
          AllowVectorExport = True
          Left = 102.063080000000000000
          Top = 249.448980000000000000
          Width = 68.000000000000000000
          Height = 49.133890000000000000
          BarType = bcCode128
          Expression = '<frCredito."CUPOM_DEVOLUCAO">'
          Frame.Typ = []
          Rotation = 0
          TestLine = False
          Text = '900390'
          WideBarRatio = 2.000000000000000000
          Zoom = 1.000000000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Arial'
          Font.Style = []
        end
      end
    end
  end
  object qrImpCredito: TUniQuery
    Connection = frmModulo.Conexao_Servidor
    SQL.Strings = (
      'select'
      '  dev.cupom_devolucao,'
      '  dev.data,'
      '  dev.hora,'
      '  cre.valor,'
      '  cre.valor_utilizado,'
      '  cli.nome,'
      '  cre.vencimento'
      'from'
      '  devolucao dev'
      '  inner join devolucao_credito cre on cre.id = dev.id_credito'
      '  inner join c000007 cli on cli.codigo = dev.codcliente'
      'where dev.CUPOM_DEVOLUCAO = :cupom')
    Left = 536
    Top = 499
    ParamData = <
      item
        DataType = ftString
        Name = 'cupom'
        ParamType = ptInput
        Value = '001000000900390'
      end>
    object qrImpCreditoCUPOM_DEVOLUCAO: TStringField
      FieldName = 'CUPOM_DEVOLUCAO'
      Size = 50
    end
    object qrImpCreditoDATA: TDateField
      FieldName = 'DATA'
    end
    object qrImpCreditoHORA: TStringField
      FieldName = 'HORA'
      Size = 8
    end
    object qrImpCreditoVALOR: TFloatField
      FieldName = 'VALOR'
      ReadOnly = True
    end
    object qrImpCreditoVALOR_UTILIZADO: TFloatField
      FieldName = 'VALOR_UTILIZADO'
      ReadOnly = True
    end
    object qrImpCreditoNOME: TStringField
      FieldName = 'NOME'
      ReadOnly = True
      Size = 80
    end
    object qrImpCreditoVENCIMENTO: TDateField
      FieldName = 'VENCIMENTO'
      ReadOnly = True
    end
  end
  object cdsDevolucao: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 1024
    Top = 131
    object cdsDevolucaocodcupom: TStringField
      FieldName = 'codcupom'
      Size = 50
    end
    object cdsDevolucaocodcliente: TIntegerField
      FieldName = 'codcliente'
    end
    object cdsDevolucaoitem: TIntegerField
      FieldName = 'item'
    end
    object cdsDevolucaovalor: TFloatField
      FieldName = 'valor'
    end
  end
  object dsContingencia: TDataSource
    DataSet = qrContingencia
    Left = 768
    Top = 464
  end
  object cpACBrECF: TACBrECF
    QuebraLinhaRodape = False
    Porta = 'procurar'
    ReTentar = False
    TimeOut = 10000
    DescricaoGrande = True
    MsgAguarde = 'Aguardando a resposta da Impressora: %d segundos'
    MsgTrabalhando = 'Impressora est'#225' trabalhando'
    MsgPoucoPapel = -1
    MsgRelatorio = 'Imprimindo %s  %d'#170' Via '
    MsgPausaRelatorio = 'Destaque a %d'#170' via, <ENTER> proxima, %d seg.'
    PaginaDeCodigo = 850
    DecimaisPreco = 2
    MemoParams.Strings = (
      '[Cabecalho]'
      'LIN000=<center><b>Nome da Empresa</b></center>'
      'LIN001=<center>Nome da Rua , 1234  -  Bairro</center>'
      'LIN002=<center>Cidade  -  UF  -  99999-999</center>'
      
        'LIN003=<center>CNPJ: 01.234.567/0001-22    IE: 012.345.678.90</c' +
        'enter>'
      
        'LIN004=<table width=100%><tr><td align=left><code>Data</code> <c' +
        'ode>Hora</code></td><td align=right>COO: <b><code>NumCupom</code' +
        '></b></td></tr></table>'
      'LIN005=<hr>'
      ''
      '[Cabecalho_Item]'
      'LIN000=ITEM   CODIGO     DESCRICAO'
      'LIN001=QTD         x UNITARIO       Aliq     VALOR (R$)'
      'LIN002=<hr>'
      
        'MascaraItem=III CCCCCCCCCCCCCC DDDDDDDDDDDDDDDDDDDDDDDDDDDDDQQQQ' +
        'QQQQ UU x VVVVVVVVVVVVV AAAAAA TTTTTTTTTTTTT'
      ''
      '[Rodape]'
      'LIN000=<hr>'
      
        'LIN001=<table width=100%><tr><td align=left><code>Data</code> <c' +
        'ode>Hora</code></td><td align=right>Projeto ACBr: <b><code>ACBR<' +
        '/code></b></td></tr></table>'
      'LIN002=<center>Obrigado Volte Sempre</center>'
      'LIN003=<hr>'
      ''
      '[Formato]'
      'Colunas=48'
      'HTML=1'
      'HTML_Title_Size=4'
      'HTML_Font=<font size="5" face="Lucida Console">')
    ArqLOG = 'acbrlog.txt'
    ConfigBarras.MostrarCodigo = True
    ConfigBarras.LarguraLinha = 3
    ConfigBarras.Altura = 10
    ConfigBarras.Margem = 0
    InfoRodapeCupom.Imposto.ModoCompacto = False
    Left = 200
    Top = 365
  end
  object cdsCrediario: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 944
    Top = 123
    object cdsCrediarioparcela: TIntegerField
      FieldName = 'parcela'
    end
    object cdsCrediariocliente: TStringField
      FieldName = 'cliente'
      Size = 6
    end
    object cdsCrediariovencimento: TDateField
      FieldName = 'vencimento'
    end
    object cdsCrediariovalor: TFloatField
      FieldName = 'valor'
      DisplayFormat = '#,##0.00'
    end
  end
  object qrAltenaNumeroNFCeLocal: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      '')
    Left = 885
    Top = 383
  end
  object cxImageList1: TcxImageList
    SourceDPI = 96
    FormatVersion = 1
    DesignInfo = 30540101
    ImageInfo = <
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          00000000000000000002000000070000000C0000001000000012000000110000
          000E000000080000000200000000000000000000000000000000000000000000
          000100000004000101120D2A1D79184E36C6216B4BFF216B4BFF216C4BFF1A53
          3AD20F2F21840001011500000005000000010000000000000000000000000000
          0005050F0A351C5B40DC24805CFF29AC7EFF2CC592FF2DC894FF2DC693FF2AAE
          80FF258560FF1A563DD405110C3D00000007000000010000000000000003040E
          0A31206548ED299D74FF2FC896FF2EC996FF56D4ACFF68DAB5FF3BCD9DFF30C9
          96FF32CA99FF2BA479FF227050F805110C3D00000005000000000000000A1A57
          3DD02EA57CFF33CA99FF2EC896FF4CD2A8FF20835CFF00673BFF45BE96FF31CB
          99FF31CB98FF34CC9CFF31AD83FF1B5C41D300010113000000020B23185E2E8A
          66FF3BCD9EFF30CA97FF4BD3A9FF349571FF87AF9DFFB1CFC1FF238A60FF45D3
          A8FF36CF9FFF33CD9BFF3ED0A3FF319470FF0F32237F00000007184D37B63DB3
          8CFF39CD9FFF4BD5A9FF43A382FF699782FFF8F1EEFFF9F3EEFF357F5DFF56C4
          A1FF43D5A8FF3ED3A4FF3CD1A4FF41BC95FF1B5C43CD0000000B1C6446DF4BCA
          A4FF44D2A8FF4FB392FF4E826AFFF0E9E6FFC0C3B5FFEFE3DDFFCEDDD4FF1B75
          4FFF60DCB8FF48D8ACFF47D6AAFF51D4ACFF247A58F80000000E217050F266D9
          B8FF46D3A8FF0B6741FFD2D2CBFF6A8F77FF116B43FF73967EFFF1E8E3FF72A2
          8BFF46A685FF5EDFBAFF4CD9AFFF6BE2C2FF278460FF020604191E684ADC78D9
          BEFF52DAB1FF3DBA92FF096941FF2F9C76FF57DEB8FF2D9973FF73967EFFF0EA
          E7FF4F886CFF5ABB9AFF5BDEB9FF7FE2C7FF27835FF80000000C19523BAB77C8
          B0FF62E0BCFF56DDB7FF59DFBAFF5CE1BDFF5EE2BEFF5FE4C1FF288C67FF698E
          76FFE6E1DCFF176B47FF5FD8B4FF83D5BDFF1E674CC60000000909201747439C
          7BFF95ECD6FF5ADFBAFF5EE2BDFF61E4BFFF64E6C1FF67E6C5FF67E8C7FF39A1
          7EFF1F6D4AFF288B64FF98EFD9FF4DAC8CFF1036286D00000004000000041C5F
          46B578C6ADFF9AEED9FF65E5C0FF64E7C3FF69E7C6FF6BE8C8FF6CE9C9FF6BEA
          C9FF5ED6B6FF97EDD7FF86D3BBFF237759D20102010C0000000100000001030A
          0718247B5BDA70C1A8FFB5F2E3FF98F0DAFF85EDD4FF75EBCEFF88EFD6FF9CF2
          DDFFBAF4E7FF78CDB3FF2A906DEA0615102E0000000200000000000000000000
          0001030A07171E694FB844AB87FF85D2BBFFA8E6D6FFC5F4EBFFABE9D8FF89D8
          C1FF4BB692FF237F60CB05130E27000000030000000000000000000000000000
          000000000001000000030A241B411B60489D258464CF2C9D77EE258867CF1F71
          56B00E3226560000000600000002000000000000000000000000}
      end
      item
        Image.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000020000000C05031A46110852AB190C76E31D0E89FF1C0E89FF190C
          76E4120852AD06031B4D0000000E000000030000000000000000000000000000
          000301010519130A55A9211593FF2225AEFF2430C2FF2535CBFF2535CCFF2430
          C3FF2225AFFF211594FF140B58B20101051E0000000400000000000000020101
          03151C1270CD2522A6FF2D3DCCFF394BD3FF3445D1FF2939CDFF2839CDFF3344
          D0FF394AD4FF2D3CCDFF2523A8FF1C1270D20101051D00000003000000091912
          5BA72A27AAFF2F41D0FF3541C7FF2726ABFF3137BCFF384AD3FF384BD3FF3137
          BCFF2726ABFF3540C7FF2E40D0FF2927ACFF1A115EB10000000D08061C3D3129
          A2FD2C3CCCFF3842C6FF5F5DBDFFEDEDF8FF8B89CEFF3337B9FF3437B9FF8B89
          CEFFEDEDF8FF5F5DBDFF3741C6FF2B3ACDFF3028A4FF0907204A1E185F9F373B
          BCFF3042D0FF2621A5FFECE7ECFFF5EBE4FFF8F2EEFF9491D1FF9491D1FFF8F1
          EDFFF3E9E2FFECE6EBFF2621A5FF2E3FCFFF343ABEFF201A66B0312A92E03542
          CBFF3446D1FF2C2FB5FF8070ADFFEBDBD3FFF4EAE4FFF7F2EDFFF8F1EDFFF4E9
          E2FFEADAD1FF7F6FACFF2B2EB5FF3144D0FF3040CBFF312A95E53E37AEFA3648
          D0FF374AD3FF3A4ED5FF3234B4FF8A7FB9FFF6ECE7FFF5ECE6FFF4EBE5FFF6EB
          E5FF897DB8FF3233B4FF384BD3FF3547D2FF3446D1FF3E37AEFA453FB4FA4557
          D7FF3B50D5FF4C5FDAFF4343B7FF9189C7FFF7EFE9FFF6EEE9FFF6EFE8FFF7ED
          E8FF9087C5FF4242B7FF495DD8FF394CD4FF3F52D4FF443FB3FA403DA1DC5967
          DAFF5B6EDDFF4F4DBAFF8F89CAFFFBF6F4FFF7F1ECFFEDE1D9FFEDE0D9FFF7F0
          EAFFFAF5F2FF8F89CAFF4E4DB9FF576ADCFF5765D9FF403EA4E12E2D70987C85
          DDFF8798E8FF291D9BFFE5DADEFFF6EEEBFFEDDFDAFF816EA9FF816EA9FFEDDF
          D8FFF4ECE7FFE5D9DCFF291D9BFF8494E7FF7A81DDFF33317BAC111125356768
          D0FC9EACEDFF686FCEFF5646A1FFCCB6BCFF7A68A8FF4C4AB6FF4D4BB7FF7A68
          A8FFCBB5BCFF5646A1FF666DCCFF9BAAEEFF696CD0FD1212273F000000043B3B
          79977D84DFFFA5B6F1FF6D74D0FF2D219BFF5151B9FF8EA2ECFF8EA1ECFF5252
          BBFF2D219BFF6B72D0FFA2B3F0FF8086E0FF404183A700000008000000010303
          050C4E509DBC8087E2FFAEBDF3FFA3B6F1FF9DAFF0FF95A9EEFF95A8EEFF9BAD
          EFFFA2B3F0FFACBCF3FF838AE3FF4F52A0C10303051100000002000000000000
          000100000005323464797378D9F8929CEAFFA1AEEFFFB0BFF3FFB0BFF4FFA2AE
          EFFF939DE9FF7479DAF83234647D000000080000000200000000000000000000
          000000000000000000031213232D40437D935D61B5D07378DFFC7378DFFC5D61
          B5D040437D951212223000000004000000010000000000000000}
      end>
  end
  object tmCaixa: TJvThreadTimer
    OnTimer = tmCaixaTimer
    Left = 64
    Top = 155
  end
  object qrCaixa: TUniQuery
    Connection = frmModulo.conexao
    SQL.Strings = (
      
        'select caixa_situacao, caixa_data_movto, FECHAMENTO_AUTOMATICO_C' +
        'AIXA_DIA from config')
    Left = 112
    Top = 235
    object qrCaixaCAIXA_SITUACAO: TStringField
      FieldName = 'CAIXA_SITUACAO'
      Size = 10
    end
    object qrCaixaCAIXA_DATA_MOVTO: TDateField
      FieldName = 'CAIXA_DATA_MOVTO'
    end
    object qrCaixaFECHAMENTO_AUTOMATICO_CAIXA_DIA: TStringField
      FieldName = 'FECHAMENTO_AUTOMATICO_CAIXA_DIA'
      Size = 1
    end
  end
  object uQry: TUniQuery
    Connection = frmModulo.conexao
    Left = 280
    Top = 131
  end
  object uqryItens: TUniQuery
    SQLUpdate.Strings = (
      '')
    Connection = frmModulo.conexao
    SQL.Strings = (
      
        'SELECT CASE WHEN IJE.CODPRODUTO = '#39'0'#39' THEN '#39'N'#39' ELSE '#39'S'#39' END VOLU' +
        'ME'
      '      ,IJE.QTD_VOLUME'
      '      ,CIT.*'
      'FROM CUPOM_ITEM_TEMP CIT'
      'INNER JOIN ESTOQUE IJE ON'
      '    IJE.CODIGO = CIT.COD_PRODUTO'
      'WHERE CIT.CANCELADO <> 1 AND       '
      '      CIT.COD_PRODUTO = :COD_PRODUTO')
    Left = 416
    Top = 128
    ParamData = <
      item
        DataType = ftInteger
        Name = 'COD_PRODUTO'
        ParamType = ptInput
        Value = nil
      end>
    object uqryItensVOLUME: TStringField
      FieldName = 'VOLUME'
      ReadOnly = True
      FixedChar = True
      Size = 1
    end
    object uqryItensCODIGO: TStringField
      FieldName = 'CODIGO'
      ReadOnly = True
      Required = True
      Size = 50
    end
    object uqryItensCOD_CUPOM: TStringField
      FieldName = 'COD_CUPOM'
      ReadOnly = True
      Required = True
      Size = 50
    end
    object uqryItensITEM: TIntegerField
      FieldName = 'ITEM'
      ReadOnly = True
      Required = True
    end
    object uqryItensCOD_PRODUTO: TIntegerField
      FieldName = 'COD_PRODUTO'
      ReadOnly = True
      Required = True
    end
    object uqryItensUNIDADE: TStringField
      FieldName = 'UNIDADE'
      ReadOnly = True
      Required = True
      Size = 5
    end
    object uqryItensQTDE: TFloatField
      FieldName = 'QTDE'
      ReadOnly = True
    end
    object uqryItensVALOR_UNITARIO: TFloatField
      FieldName = 'VALOR_UNITARIO'
      ReadOnly = True
    end
    object uqryItensVALOR_SUBTOTAL: TFloatField
      FieldName = 'VALOR_SUBTOTAL'
      ReadOnly = True
    end
    object uqryItensVALOR_DESCONTO: TFloatField
      FieldName = 'VALOR_DESCONTO'
      ReadOnly = True
    end
    object uqryItensVALOR_ACRESCIMO: TFloatField
      FieldName = 'VALOR_ACRESCIMO'
      ReadOnly = True
    end
    object uqryItensVALOR_TOTAL: TFloatField
      FieldName = 'VALOR_TOTAL'
      ReadOnly = True
    end
    object uqryItensCST: TStringField
      FieldName = 'CST'
      ReadOnly = True
      Required = True
      Size = 3
    end
    object uqryItensALIQUOTA: TFloatField
      FieldName = 'ALIQUOTA'
      ReadOnly = True
    end
    object uqryItensCOD_TOTALIZADOR: TStringField
      FieldName = 'COD_TOTALIZADOR'
      ReadOnly = True
      Required = True
      Size = 7
    end
    object uqryItensCANCELADO: TIntegerField
      FieldName = 'CANCELADO'
      ReadOnly = True
      Required = True
    end
    object uqryItensTAMANHO: TStringField
      FieldName = 'TAMANHO'
      ReadOnly = True
      FixedChar = True
      Size = 5
    end
    object uqryItensCOR: TStringField
      FieldName = 'COR'
      ReadOnly = True
      FixedChar = True
      Size = 5
    end
    object uqryItensEXTRA: TBlobField
      FieldName = 'EXTRA'
      ReadOnly = True
    end
    object uqryItensCOMPLEMENTO: TStringField
      FieldName = 'COMPLEMENTO'
      ReadOnly = True
      Size = 50
    end
    object uqryItensAPLICADO_DESCONTO_GRUPO: TStringField
      FieldName = 'APLICADO_DESCONTO_GRUPO'
      ReadOnly = True
      Size = 1
    end
    object uqryItensQTD_APLICADO_DESCONTO: TFloatField
      FieldName = 'QTD_APLICADO_DESCONTO'
      ReadOnly = True
    end
    object uqryItensDESCONTO_CUPOM: TFloatField
      FieldName = 'DESCONTO_CUPOM'
      ReadOnly = True
    end
    object uqryItensCOD_VOLUME_PRODUTO: TIntegerField
      FieldName = 'COD_VOLUME_PRODUTO'
      ReadOnly = True
    end
    object uqryItensQTD_VOLUME: TFloatField
      FieldName = 'QTD_VOLUME'
    end
    object uqryItensPOR_VOLUME: TIntegerField
      FieldName = 'POR_VOLUME'
      ReadOnly = True
    end
  end
end
