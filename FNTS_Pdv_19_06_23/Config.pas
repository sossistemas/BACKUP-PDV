unit Config;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, ACBrNFeDANFEFRDM, ACBrNFeDANFEClass, ACBrNFeDANFEFR, ACBrBase, ACBrDFe,
  ACBrNFe, Data.DB, DBAccess, Uni, MemDS, Vcl.Dialogs, Vcl.ExtDlgs,
  AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel, cxTextEdit, cxMaskEdit,
  cxDropDownEdit, cxCalc, Vcl.Samples.Spin, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons,
  Vcl.ComCtrls, dxGDIPlusClasses, AdvEdit, AdvEdBtn, AdvDirectoryEdit, Vcl.Mask,
  Vcl.DBCtrls, PageView, AdvGlowButton, Printers, system.IniFiles, pcnConversao,
  JvBaseDlg, JvSelectDirectory, Vcl.FileCtrl, ACBrDFeSSL, blcksock, System.TypInfo,
  pcnConversaoNFe, RzEdit, RzBtnEdt, cxSpinEdit, ACBrGAV, System.Actions,
  Vcl.ActnList, ACBrSATExtratoReportClass, ACBrSATExtratoFortesFr,
  ACBrSATExtratoClass, ACBrSATExtratoESCPOS, ACBrSAT, Vcl.Menus, ACBrPosPrinter,
  Vcl.OleCtrls, SHDocVw, ACBrUtil, ACBrDFeReport, ACBrDFeDANFeReport, ACBrTEFD,
  ACBrSATClass, ACBrTEFDClass, ACBrECF, cxButtons, dxBarBuiltInMenu, cxPC,
  cxGroupBox, cxRadioGroup, cxLabel, System.Math, dxSkinsCore, dxSkinsDefaultPainters,
  vcl.wwdbedit;

const
  SELDIRHELP = 1000;
  cAssinatura = '9d4c4eef8c515e2c1269c2e4fff0719d526c5096422bf1defa20df50ba06469'+
                'a28adb25ba0447befbced7c0f805a5cc58496b7b23497af9a04f69c77f17c0c'+
                'e68161f8e4ca7e3a94c827b6c563ca6f47aea05fa90a8ce3e4327853bb2d664'+
                'ba226728fff1e2c6275ecc9b20129e1c1d2671a837aa1d265b36809501b519d'+
                'bc08129e1c1d2671a837aa1d265b36809501b519dbc08129e1c1d2671a837aa'+
                '1d265b36809501b519dbc08129e1c' ;

type
  TfrmConfig = class(TForm)
    Panel1: TPanel;
    pageview1: TPageView;
    PageSheet1: TPageSheet;
    PageSheet2: TPageSheet;
    PageSheet3: TPageSheet;
    GroupBox2: TGroupBox;
    DBCheckBox1: TDBCheckBox;
    DBCheckBox2: TDBCheckBox;
    DBCheckBox3: TDBCheckBox;
    DBCheckBox8: TDBCheckBox;
    GroupBox5: TGroupBox;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit6: TDBEdit;
    CheckHabilitaFotoProduto: TCheckBox;
    GroupBox6: TGroupBox;
    Label7: TLabel;
    CheckLogoMarca: TCheckBox;
    EditLogoMarca: TAdvEditBtn;
    OpenPictureDialog1: TOpenPictureDialog;
    GroupBox7: TGroupBox;
    Combo_Usuario: TComboBox;
    Label8: TLabel;
    EditSenhaUsuario: TEdit;
    Label10: TLabel;
    btn1: TAdvGlowButton;
    btn2: TAdvGlowButton;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    qrconfig: TUniQuery;
    dsconfig: TUniDataSource;
    PageSheet5: TPageSheet;
    grp1: TGroupBox;
    pgc1: TPageControl;
    ts1: TTabSheet;
    ts2: TTabSheet;
    grp3: TGroupBox;
    ts3: TTabSheet;
    ts4: TTabSheet;
    lbl12: TLabel;
    lbl13: TLabel;
    lbl14: TLabel;
    lbl15: TLabel;
    lbl16: TLabel;
    lbl17: TLabel;
    lbl18: TLabel;
    lbl19: TLabel;
    lbl20: TLabel;
    lbl21: TLabel;
    lbl22: TLabel;
    lbl23: TLabel;
    lbl24: TLabel;
    edtEmitCNPJ: TEdit;
    edtEmitIE: TEdit;
    edtEmitRazao: TEdit;
    edtEmitFantasia: TEdit;
    edtEmitFone: TEdit;
    edtEmitCEP: TEdit;
    edtEmitLogradouro: TEdit;
    edtEmitNumero: TEdit;
    edtEmitComp: TEdit;
    edtEmitBairro: TEdit;
    edtEmitCodCidade: TEdit;
    edtEmitCidade: TEdit;
    edtEmitUF: TEdit;
    ACBrNFe1: TACBrNFe;
    dlgOpen1: TOpenDialog;
    OpenDialog1: TOpenDialog;
    TabSheet1: TTabSheet;
    ACBRDANFENFCe: TACBrNFeDANFEFR;
    edSequencia: TEdit;
    Label21: TLabel;
    edSerie: TEdit;
    Label22: TLabel;
    edCFOP: TEdit;
    Label23: TLabel;
    DBEdit5: TDBEdit;
    Label24: TLabel;
    DBEdit7: TDBEdit;
    Label25: TLabel;
    EdFormatoOff: TRadioGroup;
    edVenedor: TRadioGroup;
    Label29: TLabel;
    Label31: TLabel;
    Label30: TLabel;
    cbFormaEmissao: TComboBox;
    cbxAtualizarXML: TCheckBox;
    cbxExibirErroSchema: TCheckBox;
    edtFormatoAlerta: TEdit;
    cbModeloDF: TComboBox;
    sbtnPathSalvar: TSpeedButton;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label42: TLabel;
    spPathSchemas: TSpeedButton;
    edtPathLogs: TEdit;
    ckSalvar: TCheckBox;
    cbxRetirarAcentos: TCheckBox;
    cbVersaoDF: TComboBox;
    edtIdToken: TEdit;
    edtToken: TEdit;
    edtPathSchemas: TEdit;
    GroupBox4: TGroupBox;
    Label13: TLabel;
    lTimeOut: TLabel;
    lSSLLib1: TLabel;
    cbxVisualizar: TCheckBox;
    cbUF: TComboBox;
    rgTipoAmb: TRadioGroup;
    cbxSalvarSOAP: TCheckBox;
    seTimeOut: TSpinEdit;
    cbSSLType: TComboBox;
    gbxRetornoEnvio: TGroupBox;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    cbxAjustarAut: TCheckBox;
    edtTentativas: TEdit;
    edtIntervalo: TEdit;
    edtAguardar: TEdit;
    gbProxy: TGroupBox;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    edtProxyHost: TEdit;
    edtProxyPorta: TEdit;
    edtProxyUser: TEdit;
    edtProxySenha: TEdit;
    TabSheet2: TTabSheet;
    cbxSalvarArqs: TCheckBox;
    cbxPastaMensal: TCheckBox;
    cbxAdicionaLiteral: TCheckBox;
    cbxEmissaoPathNFe: TCheckBox;
    cbxSalvaPathEvento: TCheckBox;
    cbxSepararPorCNPJ: TCheckBox;
    cbxSepararPorModelo: TCheckBox;
    sbPathNFe: TSpeedButton;
    Label47: TLabel;
    Label48: TLabel;
    sbPathCan: TSpeedButton;
    Label52: TLabel;
    sbPathEvento: TSpeedButton;
    edtPathNFe: TEdit;
    edtPathCan: TEdit;
    edtPathEvento: TEdit;
    grp2: TGroupBox;
    lbl1: TLabel;
    lbl2: TLabel;
    sbtnCaminhoCert: TSpeedButton;
    lbl3: TLabel;
    btnGetCert: TSpeedButton;
    lXmlSign: TLabel;
    lHttpLib: TLabel;
    lCryptLib: TLabel;
    lSSLLib: TLabel;
    edtCaminho: TEdit;
    edtSenha: TEdit;
    edtNumSerie: TEdit;
    cbXmlSignLib: TComboBox;
    cbHttpLib: TComboBox;
    cbCryptLib: TComboBox;
    cbSSLLib: TComboBox;
    edtPathInu: TEdit;
    Label50: TLabel;
    sbPathInu: TSpeedButton;
    edDescEsta: TEdit;
    Label20: TLabel;
    Label26: TLabel;
    edMargem: TSpinEdit;
    GroupBox9: TGroupBox;
    Label27: TLabel;
    Label28: TLabel;
    Label35: TLabel;
    Label39: TLabel;
    edMargDir: TcxCalcEdit;
    edMargEsq: TcxCalcEdit;
    edMargSup: TcxCalcEdit;
    edMarginf: TcxCalcEdit;
    edTipoCupom: TComboBox;
    Label18: TLabel;
    qrEstoque: TUniQuery;
    qrEstoqueCODIGO: TIntegerField;
    qrEstoqueNOME: TStringField;
    Label44: TLabel;
    edtEmitIM: TEdit;
    Button1: TButton;
    PrintDialog1: TPrintDialog;
    GroupBox11: TGroupBox;
    Label43: TLabel;
    edUtilizaSenha: TCheckBox;
    edUltimaSenha: TcxSpinEdit;
    qrSenhaPDV: TUniQuery;
    qrSenhaPDVSENHA: TLargeintField;
    qrAlteraSenhaPDV: TUniQuery;
    DBCheckBox4: TDBCheckBox;
    GroupBox13: TGroupBox;
    ComboPortaGaveta: TComboBox;
    Label5: TLabel;
    AdvGlowButton1: TAdvGlowButton;
    rdLocalGaveta: TRadioGroup;
    DBCheckBox5: TDBCheckBox;
    edContigenciaPermanente: TCheckBox;
    edMotivoContingencia: TEdit;
    Label45: TLabel;
    DBCheckBox6: TDBCheckBox;
    GroupBox8: TGroupBox;
    DBCheckBox7: TDBCheckBox;
    DBCheckBox9: TDBCheckBox;
    DBCheckBox10: TDBCheckBox;
    DBCheckBox11: TDBCheckBox;
    DBCheckBox12: TDBCheckBox;
    ckVercaixa: TCheckBox;
    CheckBox1: TCheckBox;
    cbxModeloGaveta: TComboBox;
    Label9: TLabel;
    PageSheet4: TPageSheet;
    GroupBox10: TGroupBox;
    Label11: TLabel;
    lbImpVendaSimples: TLabel;
    Label12: TLabel;
    edImprimeCupomSimples: TComboBox;
    BitBtn3: TBitBtn;
    cpExibirTela: TCheckBox;
    CheckHabGuilhotina: TCheckBox;
    ComboGuilhotina: TComboBox;
    GroupBox12: TGroupBox;
    lbImpCozinha: TLabel;
    edImpressoraEspecifica: TCheckBox;
    BitBtn4: TBitBtn;
    GroupBox14: TGroupBox;
    lbImpFicha: TLabel;
    edAtivaImpFicha: TCheckBox;
    BitBtn5: TBitBtn;
    edCodigo: TEdit;
    Label40: TLabel;
    PageSheet6: TPageSheet;
    Panel2: TPanel;
    ACBrPosPrinter1: TACBrPosPrinter;
    ACBrSAT1: TACBrSAT;
    ACBrSATExtratoESCPOS1: TACBrSATExtratoESCPOS;
    OpenDialog2: TOpenDialog;
    PrintDialog2: TPrintDialog;
    ACBrSATExtratoFortes1: TACBrSATExtratoFortes;
    SaveDialog1: TSaveDialog;
    PageControl1: TPageControl;
    tsDadosSAT: TTabSheet;
    Label54: TLabel;
    SbArqLog: TSpeedButton;
    Label55: TLabel;
    Label56: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label84: TLabel;
    edLog: TEdit;
    edNomeDLL: TEdit;
    edtCodigoAtivacao: TEdit;
    cbxUTF8: TCheckBox;
    sePagCod: TSpinEdit;
    sfeVersaoEnt: TEdit;
    cbxFormatXML: TCheckBox;
    cbxAmbiente: TComboBox;
    cbxSalvarCFe: TCheckBox;
    cbxSalvarCFeCanc: TCheckBox;
    cbxSalvarEnvio: TCheckBox;
    cbxSepararPorCNPJSat: TCheckBox;
    cbxSepararPorMES: TCheckBox;
    AdvGlowButton4: TAdvGlowButton;
    cbxModeloSat: TComboBox;
    tsDadosEmit: TTabSheet;
    Label62: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    Label65: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    edtEmitCNPJSat: TEdit;
    edtEmitIESat: TEdit;
    edtEmitIMSat: TEdit;
    cbxRegTribISSQN: TComboBox;
    cbxIndRatISSQN: TComboBox;
    cbxRegTributario: TComboBox;
    tsDadosSwHouse: TTabSheet;
    Label68: TLabel;
    Label69: TLabel;
    edtSwHCNPJ: TEdit;
    tsRede: TTabSheet;
    gbIPFix: TGroupBox;
    lSSID2: TLabel;
    lSSID3: TLabel;
    lSSID4: TLabel;
    lSSID5: TLabel;
    lSSID6: TLabel;
    edRedeIP: TEdit;
    edRedeMask: TEdit;
    edRedeGW: TEdit;
    edRedeDNS1: TEdit;
    edRedeDNS2: TEdit;
    gbPPPoE: TGroupBox;
    lSSID7: TLabel;
    lSSID8: TLabel;
    edRedeUsuario: TEdit;
    edRedeSenha: TEdit;
    rgRedeTipoInter: TRadioGroup;
    gbWiFi: TGroupBox;
    lSSID: TLabel;
    Label70: TLabel;
    lSSID1: TLabel;
    edRedeSSID: TEdit;
    cbxRedeSeg: TComboBox;
    edRedeCodigo: TEdit;
    rgRedeTipoLan: TRadioGroup;
    GroupBox15: TGroupBox;
    lSSID9: TLabel;
    lSSID10: TLabel;
    lSSID11: TLabel;
    lSSID12: TLabel;
    cbxRedeProxy: TComboBox;
    edRedeProxyIP: TEdit;
    edRedeProxyUser: TEdit;
    edRedeProxySenha: TEdit;
    edRedeProxyPorta: TSpinEdit;
    Impressao: TTabSheet;
    GroupBox16: TGroupBox;
    Label71: TLabel;
    Label72: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    Label75: TLabel;
    lImpressora: TLabel;
    seLargura: TSpinEdit;
    seMargemTopo: TSpinEdit;
    seMargemFundo: TSpinEdit;
    seMargemEsquerda: TSpinEdit;
    seMargemDireita: TSpinEdit;
    cbPreview: TCheckBox;
    bImpressora: TAdvGlowButton;
    TabSheet4: TTabSheet;
    Label83: TLabel;
    edCfopPadrao: TEdit;
    bInicializar: TAdvGlowButton;
    AdvGlowButton3: TAdvGlowButton;
    edtSwHAssinatura: TMemo;
    AdvGlowButton2: TAdvGlowButton;
    mTexto: TMemo;
    SpeedButton1: TSpeedButton;
    edPathSalvarXmlSatEnvio: TEdit;
    Label57: TLabel;
    edPathSalvarXmlSatVenda: TEdit;
    Label58: TLabel;
    SpeedButton2: TSpeedButton;
    edPathSalvarXmlSatCanc: TEdit;
    Label76: TLabel;
    SpeedButton3: TSpeedButton;
    edVersaoQR: TRadioGroup;
    edCartaoNaoInterfere: TCheckBox;
    GroupBox17: TGroupBox;
    lbImpCarne: TLabel;
    edImpCarne: TCheckBox;
    BitBtn1: TBitBtn;
    edExibeEstoqueConsulta: TCheckBox;
    edExibePrecoConsulta: TCheckBox;
    edSenhaVendaSemEstoque: TCheckBox;
    edVerificaItemAntes: TCheckBox;
    PageSheet7: TPageSheet;
    ACBrTEFD1: TACBrTEFD;
    LBlbl4: TLabel;
    grp4: TGroupBox;
    LBlbl5: TLabel;
    LBlbl6: TLabel;
    LBlbl8: TLabel;
    LBlbl7: TLabel;
    edtLoja: TEdit;
    edtTerminal: TEdit;
    edtSrv: TEdit;
    edt4: TEdit;
    cbxGP: TComboBox;
    grpTef: TGroupBox;
    LBLbImpTef: TLabel;
    LBlbl9: TLabel;
    cbxModImp: TComboBox;
    cbxPorta: TComboBox;
    edCnpj: TEdit;
    Label77: TLabel;
    edLioberado: TCheckBox;
    PageSheet8: TPageSheet;
    SpeedButton4: TSpeedButton;
    Label78: TLabel;
    Label79: TLabel;
    Label80: TLabel;
    Label81: TLabel;
    Label82: TLabel;
    Label85: TLabel;
    Label91: TLabel;
    cbxModelo_Ecf: TComboBox;
    cbxPorta_Ecf: TComboBox;
    chTentar_Ecf: TCheckBox;
    chBloqueia_Ecf: TCheckBox;
    chExibeMsg_Ecf: TCheckBox;
    chArredondaPorQtd_Ecf: TCheckBox;
    chGavetaSinalInvertido_Ecf: TCheckBox;
    mMsg_Ecf: TMemo;
    edLog_Ecf: TEdit;
    seTimeOut_Ecf: TSpinEdit;
    seIntervaloAposComando_Ecf: TSpinEdit;
    chDescricaoGrande_Ecf: TCheckBox;
    chArredondamentoItemMFD_Ecf: TCheckBox;
    chControlePorta_Ecf: TCheckBox;
    seBandWidth_Ecf: TSpinEdit;
    ACBrECF1: TACBrECF;
    sePaginaCodigo_Ecf: TSpinEdit;
    speLinBuf_Ecf: TSpinEdit;
    Label87: TLabel;
    Label88: TLabel;
    bAtivar: TcxButton;
    btnSerial: TcxButton;
    PageSheet9: TPageSheet;
    edTipoTela: TcxRadioGroup;
    v_fundo: TPanel;
    cxGroupBox1: TcxGroupBox;
    Label19: TLabel;
    Label41: TLabel;
    Label46: TLabel;
    Label49: TLabel;
    Label51: TLabel;
    edColunas: TcxSpinEdit;
    edLinhasProduto: TcxSpinEdit;
    edTecladoVirtual: TCheckBox;
    edLinhasGrupo: TcxSpinEdit;
    edAlturaGrupo: TcxSpinEdit;
    edIncremento: TcxSpinEdit;
    Label90: TLabel;
    Panel5: TPanel;
    v_borda_inferior: TPanel;
    l_borda_inferior: TLabel;
    cxButton12: TcxButton;
    cxButton13: TcxButton;
    v_borda_esquerda: TPanel;
    v_borda_direita: TPanel;
    v_barra_superior: TPanel;
    v_espacamento_superior: TPanel;
    l_espacamento_superior: TLabel;
    cxButton4: TcxButton;
    cxButton6: TcxButton;
    Panel4: TPanel;
    v_quadro_direito: TPanel;
    l_quadro_direito: TLabel;
    v_quadro_esquerdo: TPanel;
    l_quadro_esquerdo: TLabel;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    v_espacamento_central: TPanel;
    cxButton5: TcxButton;
    cxButton9: TcxButton;
    l_espacamento_central: TcxLabel;
    v_espacamento_inferior: TPanel;
    l_espacamento_inferior: TLabel;
    cxButton7: TcxButton;
    cxButton8: TcxButton;
    v_barra_inferior: TPanel;
    l_barra_inferior: TLabel;
    cxButton14: TcxButton;
    l_borda_esquerda: TcxLabel;
    cxButton15: TcxButton;
    cxButton16: TcxButton;
    l_borda_direita: TcxLabel;
    cxButton17: TcxButton;
    cxButton21: TcxButton;
    cxButton22: TcxButton;
    cxButton23: TcxButton;
    cxButton24: TcxButton;
    l_barra_superior: TLabel;
    v_borda_superior: TPanel;
    l_borda_superior: TLabel;
    cxButton18: TcxButton;
    cxButton10: TcxButton;
    edNFceLocal: TCheckBox;
    qrAltenaNumeroNFCeLocal: TUniQuery;
    edSenhaLocal: TCheckBox;
    TipoCorte: TComboBox;
    Label89: TLabel;
    edCortapapael: TCheckBox;
    edSinalInvertido: TCheckBox;
    cxGroupBox2: TcxGroupBox;
    edRecDinheiro: TComboBox;
    Label92: TLabel;
    Label94: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    edRecConvenio: TComboBox;
    edRecCartaoCredito: TComboBox;
    edRecCarteira: TComboBox;
    edRecCartaoDebito: TComboBox;
    edRecCheque: TComboBox;
    edRecECFDinheiro: TEdit;
    edRecECFConvenio: TEdit;
    edRecECFCartaoCredito: TEdit;
    edRecECFCarteira: TEdit;
    edRecECFCartaoDebito: TEdit;
    edRecECFCheque: TEdit;
    edCaminhoFiscal: TEdit;
    Label93: TLabel;
    edMD5: TEdit;
    Label86: TLabel;
    DBCheckBox13: TDBCheckBox;
    btnImpEmpresa: TcxButton;
    GroupBox3: TGroupBox;
    sangria_suprim_preview: TCheckBox;
    sangria_suprim_duasvias: TCheckBox;
    DBEdit8: TDBEdit;
    Label95: TLabel;
    qrconfigCODIGO: TIntegerField;
    qrconfigFORMA_CREDIARIO: TStringField;
    qrconfigFORMA_CHEQUE: TStringField;
    qrconfigFORMA_CARTAO: TStringField;
    qrconfigFORMA_CONVENIO: TStringField;
    qrconfigFORMA_DINHEIRO: TStringField;
    qrconfigNF_PAULISTA: TIntegerField;
    qrconfigTOTALIZADOR_NF_CREDIARIO: TStringField;
    qrconfigNOME_TOTALIZADOR_NF_CREDIARIO: TStringField;
    qrconfigTOTALIZADOR_NF_RECEBIMENTO: TStringField;
    qrconfigNOME_TOTALIZADOR_NF_RECEBIMENTO: TStringField;
    qrconfigTOTALIZADOR_SANGRIA: TStringField;
    qrconfigTOTALIZADOR_SUPRIMENTO: TStringField;
    qrconfigCOMPROVANTE_CREDIARIO: TIntegerField;
    qrconfigCOMPROVANTE_CREDIARIO_PRODUTO: TIntegerField;
    qrconfigUSA_TEF: TIntegerField;
    qrconfigUSA_TEF_TECBAN: TIntegerField;
    qrconfigUSA_TEF_HIPERCARD: TIntegerField;
    qrconfigCAIXA_SITUACAO: TStringField;
    qrconfigCAIXA_DATA_MOVTO: TDateField;
    qrconfigCAIXA_COO_INICIAL: TStringField;
    qrconfigMUDA_QTDE: TIntegerField;
    qrconfigMUDA_UNITARIO: TIntegerField;
    qrconfigMUDA_TOTAL: TIntegerField;
    qrconfigCARGA_DATA: TStringField;
    qrconfigCARGA_HORA: TStringField;
    qrconfigSENHA_DESCONTO: TIntegerField;
    qrconfigDESCONTO_MAXIMO: TFloatField;
    qrconfigPARNOM: TStringField;
    qrconfigPARVAL: TStringField;
    qrconfigPARTIP: TStringField;
    qrconfigALTERNARPRECOS: TIntegerField;
    qrconfigPRECOCLIENTE: TIntegerField;
    qrconfigCONSULTACREDIARIO: TIntegerField;
    qrconfigAVISOSONORO: TIntegerField;
    qrconfigDESCACRECPADRAO: TIntegerField;
    qrconfigCORFUNDO: TIntegerField;
    qrconfigCORLETRAS: TIntegerField;
    qrconfigNAOVENDERSEMESTOQUE: TIntegerField;
    qrconfigFECHAMENTO: TDateTimeField;
    qrconfigIMPORTA_ORCAMENTO_GERENCIAL: TStringField;
    qrconfigFORMA_CARTAO_CREDITO: TStringField;
    qrconfigFORMA_CHEQUE_APRAZO: TStringField;
    qrconfigIMPRIME_MENSAGEM_CUPOM: TStringField;
    qrconfigPESAGEM_AUTOMATICA: TStringField;
    qrconfigSOLICITA_SENHA_CANCELAMENTO: TStringField;
    qrconfigDESCONTO_MAXIMO_VENDA: TFloatField;
    qrconfigNUMCAIXA: TIntegerField;
    qrconfigMEN_CUP_SIMPLES: TStringField;
    qrconfigMEN_CUP_NFCE: TStringField;
    qrconfigMEN_CUP_SAT: TStringField;
    qrconfigMEN_CUP_COZINHA: TStringField;
    qrconfigNAO_UTILIZA_EXTRA: TStringField;
    qrconfigFECHAMENTO_AUTOMATICO_CAIXA_DIA: TStringField;
    qrconfigEMPRESA_EXTRUTURAL: TStringField;
    qrconfigEMPRESA_TIPO: TStringField;
    qrconfigEMPRESA_DESCRICAO: TStringField;
    qrconfigPRAZO_DEVOLUCAO: TIntegerField;
    GroupBox18: TGroupBox;
    chkInverterOrderImpressao: TCheckBox;
    chkImprimirListageSuprimentos: TCheckBox;
    chkImprimirListagemSangrias: TCheckBox;
    chkSepararListagens: TCheckBox;
    wwDBTrocoMaximo: TwwDBEdit;
    qrconfigTROCO_MAXIMO: TFloatField;
    Label96: TLabel;
    chk_imprimir_fech_cego: TCheckBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure EditLogoMarcaClickBtn(Sender: TObject);
    procedure CheckLogoMarcaClick(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure sbtnCaminhoCertClick(Sender: TObject);
    procedure btnGetCertClick(Sender: TObject);
    procedure cbSSLLibChange(Sender: TObject);
    procedure cbCryptLibChange(Sender: TObject);
    procedure cbHttpLibChange(Sender: TObject);
    procedure cbXmlSignLibChange(Sender: TObject);
    procedure sbtnPathSalvarClick(Sender: TObject);
    procedure spPathSchemasClick(Sender: TObject);
    procedure cbSSLTypeChange(Sender: TObject);
    procedure sbPathNFeClick(Sender: TObject);
    procedure sbPathCanClick(Sender: TObject);
    procedure sbPathInuClick(Sender: TObject);
    procedure sbPathEventoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure ACBrSAT1GetcodigoDeAtivacao(var Chave: AnsiString);
    procedure ACBrSAT1GetsignAC(var Chave: AnsiString);
    procedure ACBrSAT1GravarLog(const ALogLine: string; var Tratado: Boolean);
    procedure AdvGlowButton4Click(Sender: TObject);
    procedure cbxSalvarCFeClick(Sender: TObject);
    procedure cbxSalvarCFeCancClick(Sender: TObject);
    procedure cbxSalvarEnvioClick(Sender: TObject);
    procedure cbxSepararPorCNPJSatClick(Sender: TObject);
    procedure cbxSepararPorMESClick(Sender: TObject);
    procedure cbxUTF8Click(Sender: TObject);
    procedure sePagCodChange(Sender: TObject);
    procedure rgRedeTipoInterClick(Sender: TObject);
    procedure rgRedeTipoLanClick(Sender: TObject);
    procedure cbxRedeProxyChange(Sender: TObject);
    procedure bImpressoraClick(Sender: TObject);
    procedure bInicializarClick(Sender: TObject);
    procedure AdvGlowButton3Click(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure cbxModelo_EcfChange(Sender: TObject);
    procedure cbxPorta_EcfChange(Sender: TObject);
    procedure seTimeOut_EcfChange(Sender: TObject);
    procedure seIntervaloAposComando_EcfChange(Sender: TObject);
    procedure seBandWidth_EcfChange(Sender: TObject);
    procedure chTentar_EcfClick(Sender: TObject);
    procedure chControlePorta_EcfClick(Sender: TObject);
    procedure chBloqueia_EcfClick(Sender: TObject);
    procedure chExibeMsg_EcfClick(Sender: TObject);
    procedure chArredondaPorQtd_EcfClick(Sender: TObject);
    procedure chGavetaSinalInvertido_EcfClick(Sender: TObject);
    procedure chArredondamentoItemMFD_EcfClick(Sender: TObject);
    procedure chDescricaoGrande_EcfClick(Sender: TObject);
    procedure edLog_EcfChange(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure edOperador_ecfChange(Sender: TObject);
    procedure mMsg_EcfChange(Sender: TObject);
    procedure sePaginaCodigo_EcfChange(Sender: TObject);
    procedure speLinBuf_EcfChange(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
    procedure cxButton9Click(Sender: TObject);
    procedure cxButton7Click(Sender: TObject);
    procedure cxButton8Click(Sender: TObject);
    procedure cxButton6Click(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure cxButton10Click(Sender: TObject);
    procedure cxButton18Click(Sender: TObject);
    procedure cxButton14Click(Sender: TObject);
    procedure cxButton15Click(Sender: TObject);
    procedure cxButton13Click(Sender: TObject);
    procedure cxButton12Click(Sender: TObject);
    procedure cxButton16Click(Sender: TObject);
    procedure cxButton23Click(Sender: TObject);
    procedure cxButton24Click(Sender: TObject);
    procedure cxButton21Click(Sender: TObject);
    procedure cxButton22Click(Sender: TObject);
    procedure cxButton17Click(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure btnSerialClick(Sender: TObject);
    procedure edRecDinheiroChange(Sender: TObject);
    procedure edRecConvenioChange(Sender: TObject);
    procedure edRecCarteiraChange(Sender: TObject);
    procedure edRecCartaoCreditoChange(Sender: TObject);
    procedure edRecCartaoDebitoChange(Sender: TObject);
    procedure edRecChequeChange(Sender: TObject);
    procedure edTipoCupomChange(Sender: TObject);
    procedure btnImpEmpresaClick(Sender: TObject);
    procedure DBEdit8KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit8Change(Sender: TObject);
    procedure qrconfigAfterOpen(DataSet: TDataSet);
    procedure qrconfigPRAZO_DEVOLUCAOSetText(Sender: TField; const Text: string);
    procedure qrconfigPRAZO_DEVOLUCAOGetText(Sender: TField; var Text: string; DisplayText: Boolean);
  private
    { Private declarations }
    procedure GravaConfiguracoesLocais;
    procedure LerConfiguracoes;
    procedure LerCaixa;
    procedure AtualizaSSLLibsCombo;
    procedure PathClick(Sender: TObject);
    procedure GravarConfiguracao;
    procedure LerConfiguracao;
    procedure AjustaACBrSAT ;
    procedure LeDadosRedeSAT;
    function ValidaDados:Boolean;
    procedure GravaDados;
    procedure ConfiguraECF;


  public
    { Public declarations }
  end;

var
  frmConfig: TfrmConfig;

implementation

uses
  modulo, Funcoes, UFuncoes, principal, ConfiguraSerial;

{$R *.dfm}

procedure TfrmConfig.FormActivate(Sender: TObject);
begin
  pageview1.ActivePageIndex := 0;
end;

procedure TfrmConfig.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := cafree;
end;

procedure TfrmConfig.FormCreate(Sender: TObject);
var
  T: TSSLLib;
  I: TpcnTipoEmissao;
  J: TpcnModeloDF;
  K: TpcnVersaoDF;
  U: TSSLCryptLib;
  V: TSSLHttpLib;
  X: TSSLXmlSignLib;
  Y: TSSLType;
  S: TpcnVersaoQrCode;
  Sat_I : TACBrSATModelo;
  Sat_J : TpcnTipoAmbiente ;
  Sat_K : TpcnRegTribISSQN ;
  Sat_L : TpcnindRatISSQN ;
  Sat_M : TpcnRegTrib ;
  Sat_N: TACBrPosPrinterModelo;
  Sat_O: TACBrPosPaginaCodigo;
  TEF : TACBrTEFDTipo;
  Xpos: TACBrPosPrinterModelo;
  Kpos: integer;
  E: TACBrECFModelo;
begin
  Self.Width := 988;
  
  ///
  cbxModelo_Ecf.Items.Clear;
  For E := Low(TACBrECFModelo) to High(TACBrECFModelo) do
    cbxModelo_Ecf.Items.Add(GetEnumName(TypeInfo(TACBrECFModelo), Integer(E)));
  cbxModelo_Ecf.Items[0] := 'Procurar';
  cbxModelo_Ecf.ItemIndex := 0;
  cbxPorta_Ecf.Items.Clear;

  ACBrECF1.Device.AcharPortasSeriais(cbxPorta_Ecf.Items);
  cbxPorta_Ecf.Items.Insert(0, 'Procurar');
  cbxPorta_Ecf.Items.Add('USB');
  cbxPorta_Ecf.Items.Add('LPT1');
  cbxPorta_Ecf.Items.Add('LPT2');
  cbxPorta_Ecf.Items.Add('LPT3');
  cbxPorta_Ecf.Items.Add('/dev/ttyS0');
  cbxPorta_Ecf.Items.Add('/dev/ttyS1');
  cbxPorta_Ecf.Items.Add('/dev/ttyUSB0');
  cbxPorta_Ecf.Items.Add('/dev/ttyUSB1');
  cbxPorta_Ecf.Items.Add('c:\temp\ecf.txt');
  cbxPorta_Ecf.Items.Add('/tmp/ecf.txt');

  cbxGP.Clear;
  for TEF := Low(TACBrTEFDTipo) to High(TACBrTEFDTipo) do
    cbxGP.Items.Add( GetEnumName(TypeInfo(TACBrTEFDTipo), Integer(TEF) ) );
    cbxGP.Items[0] := 'Nenhum';

  // modelo da impressora para impressao comprovante
  cbxModImp.Items.Clear ;
  For Xpos := Low(TACBrPosPrinterModelo) to High(TACBrPosPrinterModelo) do
  cbxModImp.Items.Add( GetEnumName(TypeInfo(TACBrPosPrinterModelo), integer(Xpos) ) ) ;
  cbxModeloGaveta.Items.Clear;
  cbxModeloGaveta.Items := cbxModImp.Items;

  cbxPorta.Items.Clear;
  ACBrPosPrinter1.Device.AcharPortasSeriais( cbxPorta.Items );
  cbxPorta.Items.Add('LPT1') ;
  cbxPorta.Items.Add('LPT2') ;
  cbxPorta.Items.Add('\\localhost\Epson') ;
  cbxPorta.Items.Add('c:\temp\ecf.txt') ;
  cbxPorta.Items.Add('TCP:192.168.0.31:9100') ;
  For kpos := 0 to Printer.Printers.Count-1 do
  cbxPorta.Items.Add('RAW:'+Printer.Printers[kpos]);
  cbxPorta.Items.Add('/dev/ttyS0') ;
  cbxPorta.Items.Add('/dev/ttyS1') ;
  cbxPorta.Items.Add('/dev/ttyUSB0') ;
  cbxPorta.Items.Add('/dev/ttyUSB1') ;
  cbxPorta.Items.Add('/tmp/ecf.txt') ;
  ComboPortaGaveta.Items.Clear;
  ComboPortaGaveta.Items := cbxPorta.Items;

  cbxModeloSat.Items.Clear ;
  For Sat_I := Low(TACBrSATModelo) to High(TACBrSATModelo) do
     cbxModeloSat.Items.Add( GetEnumName(TypeInfo(TACBrSATModelo), integer(Sat_I) ) ) ;

  cbxAmbiente.Items.Clear ;
  For Sat_J := Low(TpcnTipoAmbiente) to High(TpcnTipoAmbiente) do
     cbxAmbiente.Items.Add( GetEnumName(TypeInfo(TpcnTipoAmbiente), integer(Sat_J) ) ) ;

  cbxRegTribISSQN.Items.Clear ;
  For Sat_K := Low(TpcnRegTribISSQN) to High(TpcnRegTribISSQN) do
     cbxRegTribISSQN.Items.Add( GetEnumName(TypeInfo(TpcnRegTribISSQN), integer(Sat_K) ) ) ;

  cbxIndRatISSQN.Items.Clear ;
  For Sat_L := Low(TpcnindRatISSQN) to High(TpcnindRatISSQN) do
     cbxIndRatISSQN.Items.Add( GetEnumName(TypeInfo(TpcnindRatISSQN), integer(Sat_L) ) ) ;

  cbxRegTributario.Items.Clear ;
  For Sat_M := Low(TpcnRegTrib) to High(TpcnRegTrib) do
     cbxRegTributario.Items.Add( GetEnumName(TypeInfo(TpcnRegTrib), integer(Sat_M) ) ) ;

  cbSSLLib.Items.Clear;
  for T := Low(TSSLLib) to High(TSSLLib) do
    cbSSLLib.Items.Add(GetEnumName(TypeInfo(TSSLLib), integer(T)));
  cbSSLLib.ItemIndex := 0;

  cbCryptLib.Items.Clear;
  for U := Low(TSSLCryptLib) to High(TSSLCryptLib) do
    cbCryptLib.Items.Add(GetEnumName(TypeInfo(TSSLCryptLib), integer(U)));
  cbCryptLib.ItemIndex := 0;

  cbHttpLib.Items.Clear;
  for V := Low(TSSLHttpLib) to High(TSSLHttpLib) do
    cbHttpLib.Items.Add(GetEnumName(TypeInfo(TSSLHttpLib), integer(V)));
  cbHttpLib.ItemIndex := 0;

  cbXmlSignLib.Items.Clear;
  for X := Low(TSSLXmlSignLib) to High(TSSLXmlSignLib) do
    cbXmlSignLib.Items.Add(GetEnumName(TypeInfo(TSSLXmlSignLib), integer(X)));
  cbXmlSignLib.ItemIndex := 0;

  cbSSLType.Items.Clear;
  for Y := Low(TSSLType) to High(TSSLType) do
    cbSSLType.Items.Add(GetEnumName(TypeInfo(TSSLType), integer(Y)));
  cbSSLType.ItemIndex := 0;

  cbFormaEmissao.Items.Clear;
  for I := Low(TpcnTipoEmissao) to High(TpcnTipoEmissao) do
    cbFormaEmissao.Items.Add(GetEnumName(TypeInfo(TpcnTipoEmissao), integer(I)));
  cbFormaEmissao.Items[0] := 'teNormal';
  cbFormaEmissao.ItemIndex := 0;

  cbModeloDF.Items.Clear;
  for J := Low(TpcnModeloDF) to High(TpcnModeloDF) do
    cbModeloDF.Items.Add(GetEnumName(TypeInfo(TpcnModeloDF), integer(J)));
  cbModeloDF.Items[0] := 'moNFe';
  cbModeloDF.ItemIndex := 0;

  cbVersaoDF.Items.Clear;
  for K := Low(TpcnVersaoDF) to High(TpcnVersaoDF) do
    cbVersaoDF.Items.Add(GetEnumName(TypeInfo(TpcnVersaoDF), integer(K)));
  cbVersaoDF.Items[0] := 've200';
  cbVersaoDF.ItemIndex := 0;

  edVersaoQR.Items.Clear;
  for S := Low(TpcnVersaoQrCode) to High(TpcnVersaoQrCode) do
    edVersaoQR.Items.Add(GetEnumName(TypeInfo(TpcnVersaoQrCode), integer(S)));
  edVersaoQR.ItemIndex := 0;


  ACBrNFe1.Configuracoes.WebServices.Salvar := true;

  LerConfiguracao;

end;

procedure TfrmConfig.BitBtn1Click(Sender: TObject);
begin
  if PrintDialog1.Execute then
    lbImpCarne.Caption := Printer.Printers[Printer.PrinterIndex];
end;

procedure TfrmConfig.BitBtn3Click(Sender: TObject);
begin
  if PrintDialog1.Execute then
    lbImpVendaSimples.Caption := Printer.Printers[Printer.PrinterIndex];
end;

procedure TfrmConfig.BitBtn4Click(Sender: TObject);
begin
  if PrintDialog1.Execute then
    lbImpCozinha.Caption := Printer.Printers[Printer.PrinterIndex];
end;

procedure TfrmConfig.BitBtn5Click(Sender: TObject);
begin
  if PrintDialog1.Execute then
    lbImpFicha.Caption := Printer.Printers[Printer.PrinterIndex];

end;

procedure TfrmConfig.FormShow(Sender: TObject);
var
  I: integer;
begin

  edVenedor.ItemIndex := LerParametro('INFORMAR_VENDEDOR', 'I', 0);
  qrconfig.open;
  LerConfiguracoes;
  LerConfiguracao;

  with frmModulo do begin
    if not edNFceLocal.Checked then begin
      qradic_mestre.open;
      qradic_mestre.Refresh;
      if qradic_mestre.Locate('codigo', '888888', [loCaseInsensitive]) then
        edSequencia.Text := qradic_mestre.FieldByName('sequencia').Text
      else
        edSequencia.Clear;
    end else begin
      qrAltenaNumeroNFCeLocal.SQL.Clear;
      qrAltenaNumeroNFCeLocal.SQL.Text := 'select gen_id(NUMERADOR_NFCE,0) num from rdb$database';
      qrAltenaNumeroNFCeLocal.Open;
      edSequencia.Text := IntToStr(qrAltenaNumeroNFCeLocal.FieldByName('num').AsInteger);
    end;
  end;
end;

procedure TfrmConfig.ACBrSAT1GetcodigoDeAtivacao(var Chave: AnsiString);
begin
  Chave := edtCodigoAtivacao.Text;
end;

procedure TfrmConfig.ACBrSAT1GetsignAC(var Chave: AnsiString);
begin
  Chave := edtSwHAssinatura.Lines.Text;
end;

procedure TfrmConfig.ACBrSAT1GravarLog(const ALogLine: string;
  var Tratado: Boolean);
begin
  mTexto.Lines.Add(ALogLine);
  Tratado := False;
end;

procedure TfrmConfig.AdvGlowButton1Click(Sender: TObject);
begin
  try
    ACBrPosPrinter1.Desativar ;
    ACBrPosPrinter1.Modelo := TACBrPosPrinterModelo(cbxModeloGaveta.ItemIndex );
    ACBrPosPrinter1.Porta  := ComboPortaGaveta.Text;
    ACBrPosPrinter1.CortaPapel := edCortapapael.Checked;
    if TipoCorte.Text = 'Total' then
      ACBrPosPrinter1.TipoCorte := ctTotal
    else
      ACBrPosPrinter1.TipoCorte := ctParcial;
    ACBrPosPrinter1.ConfigGaveta.SinalInvertido := edSinalInvertido.Checked;
    ACBrPosPrinter1.Ativar;
    ACBrPosPrinter1.AbrirGaveta;
    ACBrPosPrinter1.Desativar;
  except
    on e:Exception do begin
      Application.MessageBox(PWideChar('Ocorreu um erro ao tentar abrir a gaveta:'+#13+E.Message),'Atenção!',MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmConfig.AdvGlowButton2Click(Sender: TObject);
begin
  if not ValidaDados then
    Exit;
  GravaDados;

  bEdicaoRelatorio := CheckBox1.Checked;

  application.messagebox('Configurações Salvas!', 'Aviso', mb_ok + MB_ICONINFORMATION);

  LerConfiguracao;

end;

procedure TfrmConfig.AdvGlowButton3Click(Sender: TObject);
begin
  ACBrSAT1.ConsultarStatusOperacional;
  with ACBrSAT1.Status do
  begin
    mTexto.Lines.Add('NSERIE.........: '+NSERIE);
    mTexto.Lines.Add('LAN_MAC........: '+LAN_MAC);
    mTexto.Lines.Add('STATUS_LAN.....: '+StatusLanToStr(STATUS_LAN));
    mTexto.Lines.Add('NIVEL_BATERIA..: '+NivelBateriaToStr(NIVEL_BATERIA));
    mTexto.Lines.Add('MT_TOTAL.......: '+MT_TOTAL);
    mTexto.Lines.Add('MT_USADA.......: '+MT_USADA);
    mTexto.Lines.Add('DH_ATUAL.......: '+DateTimeToStr(DH_ATUAL));
    mTexto.Lines.Add('VER_SB.........: '+VER_SB);
    mTexto.Lines.Add('VER_LAYOUT.....: '+VER_LAYOUT);
    mTexto.Lines.Add('ULTIMO_CFe.....: '+ULTIMO_CFe);
    mTexto.Lines.Add('LISTA_INICIAL..: '+LISTA_INICIAL);
    mTexto.Lines.Add('LISTA_FINAL....: '+LISTA_FINAL);
    mTexto.Lines.Add('DH_CFe.........: '+DateTimeToStr(DH_CFe));
    mTexto.Lines.Add('DH_ULTIMA......: '+DateTimeToStr(DH_CFe));
    mTexto.Lines.Add('CERT_EMISSAO...: '+DateToStr(CERT_EMISSAO));
    mTexto.Lines.Add('CERT_VENCIMENTO: '+DateToStr(CERT_VENCIMENTO));
    mTexto.Lines.Add('ESTADO_OPERACAO: '+EstadoOperacaoToStr(ESTADO_OPERACAO));
  end;

  LeDadosRedeSAT;
end;

procedure TfrmConfig.bInicializarClick(Sender: TObject);
begin
  AjustaACBrSAT;

  ACBrSAT1.Inicializado := not ACBrSAT1.Inicializado ;

  if ACBrSAT1.Inicializado then
    bInicializar.Caption := 'DesInicializar'
  else
    bInicializar.Caption := 'Inicializar' ;
end;

procedure TfrmConfig.AdvGlowButton4Click(Sender: TObject);
begin
  OpenURL( ExtractFilePath( Application.ExeName ) + edLog.Text);
end;

procedure TfrmConfig.AdvMetroButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmConfig.AjustaACBrSAT;
begin
  with ACBrSAT1 do
  begin
    Modelo  := TACBrSATModelo( cbxModeloSat.ItemIndex ) ;
    ArqLOG  := edLog.Text;
    NomeDLL := edNomeDLL.Text;
    Config.ide_numeroCaixa := 1;
    Config.ide_tpAmb       := TpcnTipoAmbiente( cbxAmbiente.ItemIndex );
    Config.ide_CNPJ        := edtSwHCNPJ.Text;
    Config.emit_CNPJ       := edtEmitCNPJSat.Text;
    Config.emit_IE         := edtEmitIESat.Text;
    Config.emit_IM         := edtEmitIMSat.Text;
    Config.emit_cRegTrib      := TpcnRegTrib( cbxRegTributario.ItemIndex ) ;
    Config.emit_cRegTribISSQN := TpcnRegTribISSQN( cbxRegTribISSQN.ItemIndex ) ;
    Config.emit_indRatISSQN   := TpcnindRatISSQN( cbxIndRatISSQN.ItemIndex ) ;
    Config.PaginaDeCodigo     := sePagCod.Value;
    Config.EhUTF8             := cbxUTF8.Checked;
    Config.infCFe_versaoDadosEnt := StringToFloat( sfeVersaoEnt.Text );

    ConfigArquivos.SalvarCFe := cbxSalvarCFe.Checked;
    ConfigArquivos.SalvarCFeCanc := cbxSalvarCFeCanc.Checked;
    ConfigArquivos.SalvarEnvio := cbxSalvarEnvio.Checked;
    ConfigArquivos.SepararPorCNPJ := cbxSepararPorCNPJSat.Checked;
    ConfigArquivos.SepararPorMes := cbxSepararPorMES.Checked;
  end

end;

procedure TfrmConfig.AtualizaSSLLibsCombo;
begin
  cbSSLLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLLib);
  cbCryptLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLCryptLib);
  cbHttpLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLHttpLib);
  cbXmlSignLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLXmlSignLib);
  cbSSLType.Enabled := (ACBrNFe1.Configuracoes.Geral.SSLHttpLib in [httpWinHttp, httpOpenSSL]);
end;

procedure TfrmConfig.bImpressoraClick(Sender: TObject);
begin
  if PrintDialog1.Execute then
    lImpressora.Caption := Printer.Printers[Printer.PrinterIndex] ;
end;

procedure TfrmConfig.GravaConfiguracoesLocais;
begin
  {: FOTO PRODUTOS}
  GravaINI(sConfiguracoes, 'PDV', 'HAB_IMG', BoolToStr(CheckHabilitaFotoProduto.Checked));

  {: LOGO MARCA}
  GravaINI(sConfiguracoes, 'PDV', 'HAB_LOGO', BoolToStr(CheckLogoMarca.Checked));
  GravaINI(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', EditLogoMarca.Text);

  {: USUARIO}
  GravaINI(sConfiguracoes, 'LOGIN', 'USUARIO', Combo_Usuario.Text);
  GravaINI(sConfiguracoes, 'LOGIN', 'SENHA', Cript('C', EditSenhaUsuario.Text));
end;

procedure TfrmConfig.GravaDados;
begin
  if UpperCase(edTipoCupom.Text) = 'NFC-E' then begin
    if not edNFceLocal.Checked then begin
      with frmModulo do begin
        qradic_mestre.open;
        qradic_mestre.Refresh;
        if qradic_mestre.Locate('codigo', '888888', [loCaseInsensitive]) then begin
          if qradic_mestre.FIELDBYNAME('sequencia').AsInteger <> StrToInt(edSequencia.Text) then begin
            qradic_mestre.Edit;
            qradic_mestre.FIELDBYNAME('sequencia').AsInteger := StrToInt(edSequencia.Text);
            qradic_mestre.Post;
            Conexao_Servidor.Commit;
          end;
        end
        else begin
          qradic_mestre.Insert;
          qradic_mestre.FIELDBYNAME('CODIGO').AsString := '888888';
          qradic_mestre.FIELDBYNAME('TABELA').AsString := 'Sequencia da NFC-e';
          qradic_mestre.FIELDBYNAME('sequencia').AsInteger := StrToInt(edSequencia.Text);
          qradic_mestre.Post;
          Conexao_Servidor.Commit;
        end;
      end;
    end else begin
      qrAltenaNumeroNFCeLocal.SQL.Clear;
      qrAltenaNumeroNFCeLocal.SQL.Text := 'SET GENERATOR NUMERADOR_NFCE TO ' + IntToStr(StrToInt(edSequencia.Text));
      qrAltenaNumeroNFCeLocal.ExecSQL;
    end;
  end;

  GravarConfiguracao;

  GravaConfiguracoesLocais;

  GravarParametro('INFORMAR_VENDEDOR', 'I', edVenedor.ItemIndex);

  if qrconfig.state = dsedit then
    qrconfig.post;

end;

procedure TfrmConfig.GravarConfiguracao;
var
  IniFile: string;
  arquivoini, Ini: TIniFile;
  StreamMemo: TMemoryStream;
begin
  arquivoini := TIniFile.Create(ExtractFilePath(Application.ExeName)+'cfg\tef.ini');
  try
    arquivoini.WriteInteger('TEF','TIPO',cbxGP.ItemIndex);
    arquivoini.WriteString('TEF','CODLOJA',edtLoja.Text);
    arquivoini.WriteString('TEF','CODTRM',edtTerminal.Text);
    arquivoini.WriteString('TEF','IPSRV',edtSrv.Text);
    arquivoini.WriteString('TEF','PINPAD',edt4.Text);
    arquivoini.WriteString('TEF','PORTAIMP',cbxPorta.Text);
    arquivoini.WriteInteger('TEF','MODIMP',cbxModImp.ItemIndex);
    arquivoini.WriteBool('TEF','LIBERADO',edLioberado.Checked);
    arquivoini.WriteString('TEF','CNPJ_SOFTHOUSE',edCnpj.Text);
  finally
    arquivoini.Free;
  end;

  IniFile := sConfiguracoes;
  Ini := TIniFile.Create(IniFile);
  try
    Ini.WriteInteger('Visual', 'v_borda_superior', v_borda_superior.Height);
    Ini.WriteInteger('Visual', 'v_borda_inferior', v_borda_inferior.Height);
    Ini.WriteInteger('Visual', 'v_borda_esquerda', v_borda_esquerda.Width);
    Ini.WriteInteger('Visual', 'v_borda_direita', v_borda_direita.Width);


    Ini.WriteInteger('Visual', 'v_espacamento_superior', v_espacamento_superior.Height);
    Ini.WriteInteger('Visual', 'v_espacamento_central', v_espacamento_central.Width);
    Ini.WriteInteger('Visual', 'v_espacamento_inferior', v_espacamento_inferior.Height);

    Ini.WriteInteger('Visual', 'v_barra_superior', v_barra_superior.Height);
    Ini.WriteInteger('Visual', 'v_barra_inferior', v_barra_inferior.Height);

    Ini.WriteInteger('Visual', 'v_quadro_esquerdo', v_quadro_esquerdo.Width);


    Ini.WriteFloat('Visual', 'p_borda_superior', Roundto((v_borda_superior.Height/v_fundo.Height)/2,-2));
    Ini.WriteFloat('Visual', 'p_borda_inferior', Roundto((v_borda_inferior.Height/v_fundo.Height)/2,-2));
    Ini.WriteFloat('Visual', 'p_borda_esquerda', Roundto((v_borda_esquerda.Width/v_fundo.Width)/2,-2));
    Ini.WriteFloat('Visual', 'p_borda_direita', Roundto((v_borda_direita.Width/v_fundo.Width)/2,-2));


    Ini.WriteFloat('Visual', 'p_espacamento_superior', Roundto((v_espacamento_superior.Height/v_fundo.Height)/2,-2));
    Ini.WriteFloat('Visual', 'p_espacamento_inferior', Roundto((v_espacamento_inferior.Height/v_fundo.Height)/2,-2));
    Ini.WriteFloat('Visual', 'p_espacamento_central', Roundto((v_espacamento_central.Width/v_fundo.Width)/2,-2));

    Ini.WriteFloat('Visual', 'p_barra_superior', Roundto((v_barra_superior.Height/v_fundo.Height)/2,-2));
    Ini.WriteFloat('Visual', 'p_barra_inferior', Roundto((v_barra_inferior.Height/v_fundo.Height)/2,-2));

    Ini.WriteFloat('Visual', 'p_quadro_esquerdo', Roundto((v_quadro_esquerdo.Width/v_fundo.Width),-2));


    Ini.WriteInteger('GAVETA', 'Ativo', rdLocalGaveta.ItemIndex);
    Ini.WriteInteger('GAVETA', 'Modelo', cbxModeloGaveta.ItemIndex);
    Ini.WriteString('GAVETA', 'Porta', ComboPortaGaveta.Text);
    Ini.WriteBool('GAVETA', 'CortaPapel', edCortapapael.Checked);
    Ini.WriteBool('GAVETA', 'Invertido', edSinalInvertido.Checked);
    Ini.WriteString('GAVETA', 'TipoCorte', TipoCorte.Text);



    Ini.WriteInteger('Certificado', 'SSLLib', cbSSLLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'CryptLib', cbCryptLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'HttpLib', cbHttpLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'XmlSignLib', cbXmlSignLib.ItemIndex);
    Ini.WriteString('Certificado', 'Caminho', edtCaminho.Text);
    Ini.WriteString('Certificado', 'Senha', edtSenha.Text);
    Ini.WriteString('Certificado', 'NumSerie', edtNumSerie.Text);

    Ini.WriteBool('Geral', 'AtualizarXML', cbxAtualizarXML.Checked);
    Ini.WriteBool('Geral', 'ExibirErroSchema', cbxExibirErroSchema.Checked);
    Ini.WriteString('Geral', 'FormatoAlerta', edtFormatoAlerta.Text);
    Ini.WriteInteger('Geral', 'FormaEmissao', cbFormaEmissao.ItemIndex);
    Ini.WriteInteger('Geral', 'ModeloDF', cbModeloDF.ItemIndex);
    Ini.WriteInteger('Geral', 'VersaoDF', cbVersaoDF.ItemIndex);
    Ini.WriteString('Geral', 'IdToken', edtIdToken.Text);
    Ini.WriteString('Geral', 'Token', edtToken.Text);
    Ini.WriteBool('Geral', 'RetirarAcentos', cbxRetirarAcentos.Checked);
    Ini.WriteBool('Geral', 'Salvar', ckSalvar.Checked);
    Ini.WriteString('Geral', 'PathSalvar', edtPathLogs.Text);
    Ini.WriteString('Geral', 'PathSchemas', edtPathSchemas.Text);
    Ini.WriteString('Geral', 'CFOP_PADRAO', edCFOP.Text);
    Ini.WriteString('Geral', 'TIPO CUPOM', edTipoCupom.Text);
    Ini.WriteBool('Geral', 'VERIFICA_CAIXA', ckVercaixa.Checked);
    if edImprimeCupomSimples.ItemIndex = 0 then
      Ini.WriteString('Geral', 'IMPRESSAO CUPOM', 'N')
    else if edImprimeCupomSimples.ItemIndex = 1 then
      Ini.WriteString('Geral', 'IMPRESSAO CUPOM', 'S')
    else if edImprimeCupomSimples.ItemIndex = 2 then
      Ini.WriteString('Geral', 'IMPRESSAO CUPOM', 'P')
    else
      Ini.WriteString('Geral', 'IMPRESSAO CUPOM', 'S');

    if lbImpVendaSimples.Caption <> 'SEM IMPRESSORA DEFINIDA!' then
      Ini.WriteString('Geral', 'IMPRESSORA VENDA SIMPLES', lbImpVendaSimples.Caption);
    Ini.WriteBool('Geral', 'Exibir em Tela', cpExibirTela.Checked);
    Ini.WriteBool('Geral', 'CARTAO_CUPOM', edCartaoNaoInterfere.Checked);

    Ini.WriteBool('Geral', 'CONTROLA_NFCE_LOCAL', edNFceLocal.Checked);


    Ini.WriteBool('Geral', 'UTILIZA TOUCH', (edTipoTela.ItemIndex = 1));
    Ini.WriteInteger('Geral', 'Colunas Touch', edColunas.Value);
    Ini.WriteInteger('Geral', 'Linhas Touch', edLinhasProduto.Value);
    Ini.WriteInteger('Geral', 'Linhas Touch Grupo', edLinhasGrupo.Value);
    Ini.WriteInteger('Geral', 'Altura Touch Grupo', edAlturaGrupo.Value);
    Ini.WriteBool('Geral', 'UTILIZA TECLADO', edTecladoVirtual.Checked);

    Ini.WriteBool('Geral', 'CONTROLE DE SENHA', edUtilizaSenha.Checked);
    Ini.WriteBool('Geral', 'SENHA_LOCAL', edSenhaLocal.Checked);


    if not edSenhaLocal.Checked then begin
      PDV_OnLine := False;
      if not frmModulo.Conexao_Servidor.Connected then begin
        repeat
          try
            frmModulo.Conexao_Servidor.Connected := False;
            frmModulo.Conexao_Servidor.Connected := True;
            PDV_OnLine := True;
          except
            on E: Exception do begin
              if Application.messagebox(PWideChar('Não foi possivel conectar ao servidor!' + #13 + 'Mensagem: ' + E.Message + #13 + 'Deseja tentar outra vez?'), 'Erro', MB_YESNO + MB_ICONERROR) = IDNO then begin
                Break;
              end;
            end;
          end;
        until PDV_OnLine = True;
      end;
      if PDV_OnLine then begin
        qrAlteraSenhaPDV.SQL.Clear;
        qrAlteraSenhaPDV.SQL.Text := 'SET GENERATOR ULTIMA_SENHA TO ' + IntToStr(edUltimaSenha.Value);
        qrAlteraSenhaPDV.ExecSQL;
      end else
        GroupBox11.Enabled := False;
    end else begin
      qrAlteraSenhaPDV.SQL.Clear;
      qrAlteraSenhaPDV.Connection := frmModulo.conexao;
      qrAlteraSenhaPDV.SQL.Text := 'SET GENERATOR ULTIMA_SENHA TO ' + IntToStr(edUltimaSenha.Value);
      qrAlteraSenhaPDV.ExecSQL;
    end;

    Ini.WriteBool('Geral', 'IMPRESSORA ESPECIFICA', edImpressoraEspecifica.Checked);
    if lbImpCozinha.Caption <> 'SEM IMPRESSORA DEFINIDA!' then
      Ini.WriteString('Geral', 'IMPRESSORA COZINHA', lbImpCozinha.Caption);

    Ini.WriteBool('Geral', 'IMPRIME CARNE', edImpCarne.Checked);
    if lbImpCarne.Caption <> 'SEM IMPRESSORA DEFINIDA!' then
      Ini.WriteString('Geral', 'IMPRESSORA CARNE', lbImpCarne.Caption);


    Ini.WriteBool('Geral', 'IMPRESSAO FICHA', edAtivaImpFicha.Checked);
    if lbImpFicha.Caption <> 'SEM IMPRESSORA DEFINIDA!' then
      Ini.WriteString('Geral', 'IMPRESSORA FICHA', lbImpFicha.Caption);
    Ini.WriteString('Geral', 'SEGURANCA FICHA', edCodigo.Text);

    Ini.WriteBool('Geral', 'EXIBE ESTOQUE CONSULTA', edExibeEstoqueConsulta.Checked);
    Ini.WriteBool('Geral', 'EXIBE PRECO VENDA CONSULTA', edExibePrecoConsulta.Checked);
    Ini.WriteBool('Geral', 'SOLICITA SENHA VENDA SEM ESTOQUE', edSenhaVendaSemEstoque.Checked);

    Ini.WriteBool('Geral', 'VERIFICA PRODUTO ANTES DE VENDER', edVerificaItemAntes.Checked);



    Ini.WriteString('GUILHOTINA', 'HABILITA', BoolToStr(CheckHabGuilhotina.Checked));
    Ini.WriteString('GUILHOTINA', 'MODELO', ComboGuilhotina.Text);


    Ini.WriteString('WebService', 'UF', cbUF.Text);
    Ini.WriteInteger('WebService', 'Ambiente', rgTipoAmb.ItemIndex);
    Ini.WriteBool('WebService', 'Visualizar', cbxVisualizar.Checked);
    Ini.WriteBool('WebService', 'SalvarSOAP', cbxSalvarSOAP.Checked);
    Ini.WriteBool('WebService', 'AjustarAut', cbxAjustarAut.Checked);
    Ini.WriteString('WebService', 'Aguardar', edtAguardar.Text);
    Ini.WriteString('WebService', 'Tentativas', edtTentativas.Text);
    Ini.WriteString('WebService', 'Intervalo', edtIntervalo.Text);
    Ini.WriteInteger('WebService', 'TimeOut', seTimeOut.Value);
    Ini.WriteInteger('WebService', 'SSLType', cbSSLType.ItemIndex);
    Ini.WriteInteger('WebService', 'FormatoOFF', EdFormatoOff.ItemIndex);
    Ini.WriteBool('WebService', 'ContingenciaPermanente', edContigenciaPermanente.Checked);
    Ini.WriteString('WebService', 'MotivoContingencia', edMotivoContingencia.Text);
    Ini.WriteInteger('WebService', 'VersaoQRCode', edVersaoQR.ItemIndex);

    Ini.WriteString('Proxy', 'Host', edtProxyHost.Text);
    Ini.WriteString('Proxy', 'Porta', edtProxyPorta.Text);
    Ini.WriteString('Proxy', 'User', edtProxyUser.Text);
    Ini.WriteString('Proxy', 'Pass', edtProxySenha.Text);

    Ini.WriteBool('Arquivos', 'Salvar', cbxSalvarArqs.Checked);
    Ini.WriteBool('Arquivos', 'PastaMensal', cbxPastaMensal.Checked);
    Ini.WriteBool('Arquivos', 'AddLiteral', cbxAdicionaLiteral.Checked);
    Ini.WriteBool('Arquivos', 'EmissaoPathNFe', cbxEmissaoPathNFe.Checked);
    Ini.WriteBool('Arquivos', 'SalvarPathEvento', cbxSalvaPathEvento.Checked);
    Ini.WriteBool('Arquivos', 'SepararPorCNPJ', cbxSepararPorCNPJ.Checked);
    Ini.WriteBool('Arquivos', 'SepararPorModelo', cbxSepararPorModelo.Checked);
    Ini.WriteString('Arquivos', 'PathNFe', edtPathNFe.Text);
    Ini.WriteString('Arquivos', 'PathCan', edtPathCan.Text);
    Ini.WriteString('Arquivos', 'PathInu', edtPathInu.Text);
    Ini.WriteString('Arquivos', 'PathEvento', edtPathEvento.Text);

    Ini.WriteString('Impressao', 'Descricao via estabelecimento', edDescEsta.Text);
    Ini.WriteInteger('Impressao', 'Espessura Margem', edMargem.Value);
    Ini.WriteFloat('Impressao', 'Margem Esquerda', edMargEsq.Value);
    Ini.WriteFloat('Impressao', 'Margem Direita', edMargDir.Value);
    Ini.WriteFloat('Impressao', 'Margem Superior', edMargSup.Value);
    Ini.WriteFloat('Impressao', 'Margem Inferior', edMarginf.Value);

    Ini.WriteString('Emitente', 'CNPJ', edtEmitCNPJ.Text);
    Ini.WriteString('Emitente', 'IE', edtEmitIE.Text);
    Ini.WriteString('Emitente', 'IM', edtEmitIM.Text);
    Ini.WriteString('Emitente', 'RazaoSocial', edtEmitRazao.Text);
    Ini.WriteString('Emitente', 'Fantasia', edtEmitFantasia.Text);
    Ini.WriteString('Emitente', 'Fone', edtEmitFone.Text);
    Ini.WriteString('Emitente', 'CEP', edtEmitCEP.Text);
    Ini.WriteString('Emitente', 'Logradouro', edtEmitLogradouro.Text);
    Ini.WriteString('Emitente', 'Numero', edtEmitNumero.Text);
    Ini.WriteString('Emitente', 'Complemento', edtEmitComp.Text);
    Ini.WriteString('Emitente', 'Bairro', edtEmitBairro.Text);
    Ini.WriteString('Emitente', 'CodCidade', edtEmitCodCidade.Text);
    Ini.WriteString('Emitente', 'Cidade', edtEmitCidade.Text);
    Ini.WriteString('Emitente', 'UF', edtEmitUF.Text);
    Ini.WriteString('Emitente', 'Serie', edSerie.Text);


    INI.WriteInteger('SAT','Modelo',cbxModeloSat.ItemIndex);
    INI.WriteString('SAT','ArqLog',edLog.Text);
    INI.WriteString('SAT','NomeDLL',edNomeDLL.Text);
    INI.WriteString('SAT','CodigoAtivacao',edtCodigoAtivacao.Text);
    INI.WriteInteger('SAT','Ambiente',cbxAmbiente.ItemIndex);
    INI.WriteInteger('SAT','PaginaDeCodigo',sePagCod.Value);
    INI.WriteFloat('SAT','versaoDadosEnt', StringToFloatDef(sfeVersaoEnt.Text,cversaoDadosEnt));
    INI.WriteBool('SAT','UTF8', cbxUTF8.Checked);

    INI.WriteBool('SAT','FormatarXML', cbxFormatXML.Checked);
    INI.WriteBool('SAT','SalvarCFe', cbxSalvarCFe.Checked);
    INI.WriteBool('SAT','SalvarCFeCanc', cbxSalvarCFeCanc.Checked);
    INI.WriteBool('SAT','SalvarEnvio', cbxSalvarEnvio.Checked);
    INI.WriteBool('SAT','SepararPorCNPJ', cbxSepararPorCNPJSat.Checked);
    INI.WriteBool('SAT','SepararPorMES', cbxSepararPorMES.Checked);

    INI.WriteString('Emit','CNPJ',edtEmitCNPJSat.Text);
    INI.WriteString('Emit','IE',edtEmitIESat.Text);
    INI.WriteString('Emit','IM',edtEmitIMSat.Text);
    INI.WriteInteger('Emit','RegTributario',cbxRegTributario.ItemIndex);
    INI.WriteInteger('Emit','RegTribISSQN',cbxRegTribISSQN.ItemIndex);
    INI.WriteInteger('Emit','IndRatISSQN',cbxIndRatISSQN.ItemIndex);

    INI.WriteString('SAT','PATH_SALVAR_ENVIO',edPathSalvarXmlSatEnvio.Text);
    INI.WriteString('SAT','PATH_SALVAR_VENDA',edPathSalvarXmlSatVenda.Text);
    INI.WriteString('SAT','PATH_SALVAR_CANC',edPathSalvarXmlSatCanc.Text);

    INI.WriteString('SwH','CNPJ',edtSwHCNPJ.Text);
    INI.WriteString('SwH','Assinatura',edtSwHAssinatura.Lines.Text);

    INI.WriteInteger('Fortes','Largura',seLargura.Value);
    INI.WriteInteger('Fortes','MargemTopo',seMargemTopo.Value);
    INI.WriteInteger('Fortes','MargemFundo',seMargemFundo.Value);
    INI.WriteInteger('Fortes','MargemEsquerda',seMargemEsquerda.Value);
    INI.WriteInteger('Fortes','MargemDireita',seMargemDireita.Value);
    INI.WriteBool('Fortes','Preview',cbPreview.Checked);

    INI.WriteBool('Fortes','Sangria_Suprimento_Preview',sangria_suprim_preview.Checked);
    INI.WriteBool('Fortes','Sangria_Suprimento_DuasVias',sangria_suprim_duasvias.Checked);

    INI.WriteBool('Fortes','InverterOrderImpressao',chkInverterOrderImpressao.Checked);
    INI.WriteBool('Fortes','ImprimirListagemSuprimentos',chkImprimirListageSuprimentos.Checked);
    INI.WriteBool('Fortes','ImprimirListagemSangrias',chkImprimirListagemSangrias.Checked);
    INI.WriteBool('Fortes','SepararListagens',chkSepararListagens.Checked);
    INI.WriteBool('Fortes','ImprimirFechamentoCego',chk_imprimir_fech_cego.Checked);

    INI.WriteString('Printer','Name',lImpressora.Caption);

    INI.WriteInteger('Rede','tipoInter',rgRedeTipoInter.ItemIndex);
    INI.WriteInteger('Rede','tipoLan',rgRedeTipoLan.ItemIndex);
    INI.WriteString('Rede','SSID',edRedeSSID.Text);
    INI.WriteInteger('Rede','seg',cbxRedeSeg.ItemIndex);
    INI.WriteString('Rede','codigo',edRedeCodigo.Text);
    INI.WriteString('Rede','lanIP',edRedeIP.Text);
    INI.WriteString('Rede','lanMask',edRedeMask.Text);
    INI.WriteString('Rede','lanGW',edRedeGW.Text);
    INI.WriteString('Rede','lanDNS1',edRedeDNS1.Text);
    INI.WriteString('Rede','lanDNS2',edRedeDNS2.Text);
    INI.WriteString('Rede','usuario',edRedeUsuario.Text);
    INI.WriteString('Rede','senha',edRedeSenha.Text);
    INI.WriteInteger('Rede','proxy',cbxRedeProxy.ItemIndex);
    INI.WriteString('Rede','proxy_ip',edRedeProxyIP.Text);
    INI.WriteInteger('Rede','proxy_porta',edRedeProxyPorta.Value);
    INI.WriteString('Rede','proxy_user',edRedeProxyUser.Text);
    INI.WriteString('Rede','proxy_senha',edRedeProxySenha.Text);

    INI.WriteString('Valores','CFOP',edCfopPadrao.Text);

    INI.WriteInteger('ECF', 'Modelo', cbxModelo_Ecf.ItemIndex);
    INI.WriteString('ECF', 'Porta', cbxPorta_Ecf.Text);
    INI.WriteInteger('ECF', 'TimeOut', seTimeOut_Ecf.Value);
    INI.WriteInteger('ECF', 'IntervaloAposComando', seIntervaloAposComando_Ecf.Value);
    INI.WriteBool('ECF', 'TentarNovamente', chTentar_Ecf.Checked);
    INI.WriteBool('ECF', 'ControlePorta', chControlePorta_Ecf.Checked);
    INI.WriteBool('ECF', 'BloqueiaMouseTeclado', chBloqueia_Ecf.Checked);
    INI.WriteBool('ECF', 'ExibeMsgAguarde', chExibeMsg_Ecf.Checked);
    INI.WriteBool('ECF', 'ArredondaPorQtd', chArredondaPorQtd_Ecf.Checked);
    INI.WriteBool('ECF', 'GavetaSinalInvertido', chGavetaSinalInvertido_Ecf.Checked);
    INI.WriteBool('ECF', 'DescricaoGrande', chDescricaoGrande_Ecf.Checked);
    INI.WriteBool('ECF', 'ArredondamentoItemMFD', chArredondamentoItemMFD_Ecf.Checked);
    INI.WriteString('ECF', 'MensagemAguarde', StringReplace(mMsg_Ecf.Text, sLineBreak, '|', [rfReplaceAll]));
    INI.WriteString('ECF', 'ArqLog', edLog_Ecf.Text);
    INI.WriteString('ECF', 'SerialParams', ACBrECF1.Device.ParamsString);
    INI.WriteInteger('ECF', 'BandWidth', seBandWidth_Ecf.Value);
    INI.WriteInteger('ECF', 'Buffer', speLinBuf_Ecf.Value);
    INI.WriteInteger('ECF', 'PaginadeCodigo', sePaginaCodigo_Ecf.Value);

    INI.WriteString('ECF', 'RecECFDinheiro', edRecECFDinheiro.Text);
    INI.WriteString('ECF', 'RecECFConvenio', edRecECFConvenio.Text);
    INI.WriteString('ECF', 'RecECFCartaoCredito', edRecECFCartaoCredito.Text);
    INI.WriteString('ECF', 'RecECFCarteira', edRecECFCarteira.Text);
    INI.WriteString('ECF', 'RecECFCartaoDebito', edRecECFCartaoDebito.Text);
    INI.WriteString('ECF', 'RecECFCheque', edRecECFCheque.Text);
    INI.WriteString('ECF', 'CaminhoFiscal', edCaminhoFiscal.Text);

    INI.WriteString('ECF', 'MD5', edMD5.Text);

  finally
    Ini.Free;
  end;
end;

procedure TfrmConfig.LerConfiguracao;
var
  tmp, IniFile: string;
  Ini: TIniFile;
  Ok: Boolean;
  StreamMemo: TMemoryStream;
  arquivoini : TIniFile;
begin
  IniFile := sConfiguracoes;

  arquivoini := TIniFile.Create(ExtractFilePath(Application.ExeName)+'cfg\tef.ini');
  try
    cbxGP.ItemIndex := arquivoini.ReadInteger('TEF','TIPO',0);
    edtLoja.Text := arquivoini.ReadString('TEF','CODLOJA','');
    edtTerminal.Text := arquivoini.ReadString('TEF','CODTRM','');
    edtSrv.Text := arquivoini.ReadString('TEF','IPSRV','');
    edt4.Text := arquivoini.ReadString('TEF','PINPAD','');
    cbxPorta.Text := arquivoini.ReadString('TEF','PORTAIMP','');
    cbxModImp.ItemIndex := arquivoini.ReadInteger('TEF','MODIMP',0);
    edLioberado.Checked := arquivoini.ReadBool('TEF','LIBERADO',False);
    edCnpj.Text := arquivoini.ReadString('TEF','CNPJ_SOFTHOUSE','');
  finally
    arquivoini.Free;
  end;

  Ini := TIniFile.Create(IniFile);
  try
    v_borda_superior.Height := Ini.ReadInteger('Visual', 'v_borda_superior', v_borda_superior.Height);
    v_borda_inferior.Height := Ini.ReadInteger('Visual', 'v_borda_inferior', v_borda_inferior.Height);
    v_borda_esquerda.Width := Ini.ReadInteger('Visual', 'v_borda_esquerda', v_borda_esquerda.Width);
    v_borda_direita.Width := Ini.ReadInteger('Visual', 'v_borda_direita', v_borda_direita.Width);


    v_espacamento_superior.Height := Ini.ReadInteger('Visual', 'v_espacamento_superior', v_espacamento_superior.Height);
    v_espacamento_central.Width := Ini.ReadInteger('Visual', 'v_espacamento_central', v_espacamento_central.Width);
    v_espacamento_inferior.Height := Ini.ReadInteger('Visual', 'v_espacamento_inferior', v_espacamento_inferior.Height);

    v_barra_superior.Height := Ini.ReadInteger('Visual', 'v_barra_superior', v_barra_superior.Height);
    v_barra_inferior.Height := Ini.ReadInteger('Visual', 'v_barra_inferior', v_barra_inferior.Height);

    v_quadro_esquerdo.Width := Ini.ReadInteger('Visual', 'v_quadro_esquerdo', v_quadro_esquerdo.Width);


    rdLocalGaveta.ItemIndex := Ini.ReadInteger('GAVETA', 'Ativo', 0);
    cbxModeloGaveta.ItemIndex := Ini.ReadInteger('GAVETA', 'Modelo', 0);
    ComboPortaGaveta.Text := Ini.ReadString('GAVETA', 'Porta', '');
    edCortapapael.Checked := Ini.ReadBool('GAVETA', 'CortaPapel', True);
    edSinalInvertido.Checked := Ini.ReadBool('GAVETA', 'Invertido', False);
    TipoCorte.Text := Ini.ReadString('GAVETA', 'TipoCorte', 'Total');



    l_borda_inferior.Caption := 'Borda Inferior ('+IntToStr(v_borda_inferior.Height)+')';
    l_borda_esquerda.Caption := 'Borda Esquerda ('+IntToStr(v_borda_esquerda.Width)+')';
    l_borda_direita.Caption := 'Borda Direita ('+IntToStr(v_borda_direita.Width)+')';
    l_borda_superior.Caption := 'Borda Superior ('+IntToStr(v_borda_superior.Height)+')';

    l_quadro_esquerdo.Caption := 'Quadro Esquerdo ('+IntToStr(v_quadro_esquerdo.Width)+')';
    l_quadro_direito.Caption := 'Quadro Direito ('+IntToStr(v_quadro_direito.Width)+')';

    l_espacamento_central.Caption := 'Espaçamento Central ('+IntToStr(v_espacamento_central.Width)+')';
    l_espacamento_superior.Caption := 'Espaçamento Superior ('+IntToStr(v_espacamento_superior.Height)+')';
    l_espacamento_inferior.Caption := 'Espaçamento Inferior ('+IntToStr(v_espacamento_inferior.Height)+')';


    l_barra_inferior.Caption := 'Barra Inferior ('+IntToStr(v_barra_inferior.Height)+')';
    l_barra_superior.Caption := 'Barra Superior ('+IntToStr(v_barra_superior.Height)+')';



    cbSSLLib.ItemIndex := Ini.ReadInteger('Certificado', 'SSLLib', 0);
    cbCryptLib.ItemIndex := Ini.ReadInteger('Certificado', 'CryptLib', 0);
    cbHttpLib.ItemIndex := Ini.ReadInteger('Certificado', 'HttpLib', 0);
    cbXmlSignLib.ItemIndex := Ini.ReadInteger('Certificado', 'XmlSignLib', 0);
    edtCaminho.Text := Ini.ReadString('Certificado', 'Caminho', '');
    edtSenha.Text := Ini.ReadString('Certificado', 'Senha', '');
    edtNumSerie.Text := Ini.ReadString('Certificado', 'NumSerie', '');

    cbxAtualizarXML.Checked := Ini.ReadBool('Geral', 'AtualizarXML', False);
    cbxExibirErroSchema.Checked := Ini.ReadBool('Geral', 'ExibirErroSchema', False);
    edtFormatoAlerta.Text := Ini.ReadString('Geral', 'FormatoAlerta', '');
    cbFormaEmissao.ItemIndex := Ini.ReadInteger('Geral', 'FormaEmissao', 0);
    cbModeloDF.ItemIndex := Ini.ReadInteger('Geral', 'ModeloDF', 0);
    cbVersaoDF.ItemIndex := Ini.ReadInteger('Geral', 'VersaoDF', 0);
    edtIdToken.Text := Ini.ReadString('Geral', 'IdToken', '');
    edtToken.Text := Ini.ReadString('Geral', 'Token', '');
    cbxRetirarAcentos.Checked := Ini.ReadBool('Geral', 'RetirarAcentos', False);
    ckSalvar.Checked := Ini.ReadBool('Geral', 'Salvar', False);
    edtPathLogs.Text := Ini.ReadString('Geral', 'PathSalvar', '');
    edtPathSchemas.Text := Ini.ReadString('Geral', 'PathSchemas', '');
    edCFOP.Text := Ini.ReadString('Geral', 'CFOP_PADRAO', '');
    edTipoCupom.Text := Ini.ReadString('Geral', 'TIPO CUPOM', '');
    ckVercaixa.Checked := Ini.ReadBool('Geral', 'VERIFICA_CAIXA', True);
    edCartaoNaoInterfere.Checked := Ini.ReadBool('Geral', 'CARTAO_CUPOM', False);
    edSenhaLocal.Checked := Ini.ReadBool('Geral', 'SENHA_LOCAL', False);

    edNFceLocal.Checked := Ini.ReadBool('Geral', 'CONTROLA_NFCE_LOCAL', False);

    lbImpVendaSimples.Caption := Ini.ReadString('Geral', 'IMPRESSORA VENDA SIMPLES', 'SEM IMPRESSORA DEFINIDA!');

    tmp := Ini.ReadString('Geral', 'IMPRESSAO CUPOM', 'S');
    if tmp = 'S' then
      edImprimeCupomSimples.ItemIndex := 1
    else if tmp = 'N' then
      edImprimeCupomSimples.ItemIndex := 0
    else if tmp = 'P' then
      edImprimeCupomSimples.ItemIndex := 2
    else
      edImprimeCupomSimples.ItemIndex := 1;

    cpExibirTela.Checked := Ini.ReadBool('Geral', 'Exibir em Tela', False);
    edColunas.Value := Ini.ReadInteger('Geral', 'Colunas Touch', 5);
    edLinhasProduto.Value := Ini.ReadInteger('Geral', 'Linhas Touch', 5);
    edLinhasGrupo.Value := Ini.ReadInteger('Geral', 'Linhas Touch Grupo', 1);
    edAlturaGrupo.Value := Ini.ReadInteger('Geral', 'Altura Touch Grupo', 86);

    if Ini.ReadBool('Geral', 'UTILIZA TOUCH', False) then
      edTipoTela.ItemIndex := 1
    else
      edTipoTela.ItemIndex := 0;
    edTecladoVirtual.Checked := Ini.ReadBool('Geral', 'UTILIZA TECLADO', False);

    edUtilizaSenha.Checked := Ini.ReadBool('Geral', 'CONTROLE DE SENHA', False);

    PDV_OnLine := False;
    if not edSenhaLocal.Checked then begin
      PDV_OnLine := False;
      if not frmModulo.Conexao_Servidor.Connected then begin
          try
            frmModulo.Conexao_Servidor.Connected := False;
            frmModulo.Conexao_Servidor.Connected := True;
            PDV_OnLine := True;
          except
          end;
      end;
      if PDV_OnLine then begin
        qrSenhaPDV.Close;
        qrSenhaPDV.Open;
        edUltimaSenha.Value := qrSenhaPDVSENHA.AsInteger;
      end else
    end else begin
      qrSenhaPDV.Close;
      qrSenhaPDV.Connection := frmModulo.conexao;
      qrSenhaPDV.Open;
      edUltimaSenha.Value := qrSenhaPDVSENHA.AsInteger;
    end;

    edImpressoraEspecifica.Checked := Ini.ReadBool('Geral', 'IMPRESSORA ESPECIFICA', False);
    lbImpCozinha.Caption := Ini.ReadString('Geral', 'IMPRESSORA COZINHA', 'SEM IMPRESSORA DEFINIDA!');

    edImpCarne.Checked := Ini.ReadBool('Geral', 'IMPRIME CARNE', False);
    lbImpCarne.Caption := Ini.ReadString('Geral', 'IMPRESSORA CARNE', 'SEM IMPRESSORA DEFINIDA!');

    edAtivaImpFicha.Checked := Ini.ReadBool('Geral', 'IMPRESSAO FICHA', False);
    lbImpFicha.Caption := Ini.ReadString('Geral', 'IMPRESSORA FICHA', 'SEM IMPRESSORA DEFINIDA!');
    edCodigo.Text := Ini.ReadString('Geral', 'SEGURANCA FICHA', '');

    edExibeEstoqueConsulta.Checked := Ini.ReadBool('Geral', 'EXIBE ESTOQUE CONSULTA', True);
    edExibePrecoConsulta.Checked := Ini.ReadBool('Geral', 'EXIBE PRECO VENDA CONSULTA', True);
    edSenhaVendaSemEstoque.Checked := Ini.ReadBool('Geral', 'SOLICITA SENHA VENDA SEM ESTOQUE', False);

    edVerificaItemAntes.Checked := Ini.ReadBool('Geral', 'VERIFICA PRODUTO ANTES DE VENDER', False);


    CheckHabGuilhotina.Checked := StrToBool(Ini.ReadString('GUILHOTINA', 'HABILITA', BoolToStr(False)));
    ComboGuilhotina.Text := Ini.ReadString('GUILHOTINA', 'MODELO', '');

    cbUF.ItemIndex := cbUF.Items.IndexOf(Ini.ReadString('WebService', 'UF', 'SP'));
    rgTipoAmb.ItemIndex := Ini.ReadInteger('WebService', 'Ambiente', 0);
    cbxVisualizar.Checked := Ini.ReadBool('WebService', 'Visualizar', False);
    cbxSalvarSOAP.Checked := Ini.ReadBool('WebService', 'SalvarSOAP', False);
    cbxAjustarAut.Checked := Ini.ReadBool('WebService', 'AjustarAut', False);
    edtAguardar.Text := Ini.ReadString('WebService', 'Aguardar', '');
    edtTentativas.Text := Ini.ReadString('WebService', 'Tentativas', '');
    edtIntervalo.Text := Ini.ReadString('WebService', 'Intervalo', '');
    seTimeOut.Value := Ini.ReadInteger('WebService', 'TimeOut', 0);
    cbSSLType.ItemIndex := Ini.ReadInteger('WebService', 'SSLType', 0);
    EdFormatoOff.ItemIndex := Ini.ReadInteger('WebService', 'FormatoOFF', 0);
    edContigenciaPermanente.Checked := Ini.ReadBool('WebService', 'ContingenciaPermanente', False);
    edMotivoContingencia.Text := Ini.ReadString('WebService', 'MotivoContingencia', 'Sem conexão com a Internet.');
    edVersaoQR.ItemIndex := Ini.ReadInteger('WebService', 'VersaoQRCode', 0);

    edtProxyHost.Text := Ini.ReadString('Proxy', 'Host', '');
    edtProxyPorta.Text := Ini.ReadString('Proxy', 'Porta', '');
    edtProxyUser.Text := Ini.ReadString('Proxy', 'User', '');
    edtProxySenha.Text := Ini.ReadString('Proxy', 'Pass', '');

    cbxSalvarArqs.Checked := Ini.ReadBool('Arquivos', 'Salvar', False);
    cbxPastaMensal.Checked := Ini.ReadBool('Arquivos', 'PastaMensal', False);
    cbxAdicionaLiteral.Checked := Ini.ReadBool('Arquivos', 'AddLiteral', False);
    cbxEmissaoPathNFe.Checked := Ini.ReadBool('Arquivos', 'EmissaoPathNFe', False);
    cbxSalvaPathEvento.Checked := Ini.ReadBool('Arquivos', 'SalvarPathEvento', False);
    cbxSepararPorCNPJ.Checked := Ini.ReadBool('Arquivos', 'SepararPorCNPJ', False);
    cbxSepararPorModelo.Checked := Ini.ReadBool('Arquivos', 'SepararPorModelo', False);
    edtPathNFe.Text := Ini.ReadString('Arquivos', 'PathNFe', '');
    edtPathCan.Text := Ini.ReadString('Arquivos', 'PathCan', '');
    edtPathInu.Text := Ini.ReadString('Arquivos', 'PathInu', '');
    edtPathEvento.Text := Ini.ReadString('Arquivos', 'PathEvento', '');

    edDescEsta.Text := Ini.ReadString('Impressao', 'Descricao via estabelecimento', '');
    edMargem.Value := Ini.ReadInteger('Impressao', 'Espessura Margem', 1);
    edMargEsq.Value := Ini.ReadFloat('Impressao', 'Margem Esquerda', 0.6);
    edMargDir.Value := Ini.ReadFloat('Impressao', 'Margem Direita', 0.51);
    edMargSup.Value := Ini.ReadFloat('Impressao', 'Margem Superior', 0.8);
    edMarginf.Value := Ini.ReadFloat('Impressao', 'Margem Inferior', 0.8);

    edtEmitCNPJ.Text := Ini.ReadString('Emitente', 'CNPJ', '');
    edtEmitIE.Text := Ini.ReadString('Emitente', 'IE', '');
    edtEmitIM.Text := Ini.ReadString('Emitente', 'IM', '');
    edtEmitRazao.Text := Ini.ReadString('Emitente', 'RazaoSocial', '');
    edtEmitFantasia.Text := Ini.ReadString('Emitente', 'Fantasia', '');
    edtEmitFone.Text := Ini.ReadString('Emitente', 'Fone', '');
    edtEmitCEP.Text := Ini.ReadString('Emitente', 'CEP', '');
    edtEmitLogradouro.Text := Ini.ReadString('Emitente', 'Logradouro', '');
    edtEmitNumero.Text := Ini.ReadString('Emitente', 'Numero', '');
    edtEmitComp.Text := Ini.ReadString('Emitente', 'Complemento', '');
    edtEmitBairro.Text := Ini.ReadString('Emitente', 'Bairro', '');
    edtEmitCodCidade.Text := Ini.ReadString('Emitente', 'CodCidade', '');
    edtEmitCidade.Text := Ini.ReadString('Emitente', 'Cidade', '');
    edtEmitUF.Text := Ini.ReadString('Emitente', 'UF', '');
    edSerie.Text := Ini.ReadString('Emitente', 'Serie', '');

    cbxModeloSat.ItemIndex    := INI.ReadInteger('SAT','Modelo',0);
    edLog.Text             := INI.ReadString('SAT','ArqLog','ACBrSAT.log');
    edNomeDLL.Text         := INI.ReadString('SAT','NomeDLL','C:\SAT\SAT.DLL');
    edtCodigoAtivacao.Text := INI.ReadString('SAT','CodigoAtivacao','123456');
    cbxAmbiente.ItemIndex  := INI.ReadInteger('SAT','Ambiente',1);
    sePagCod.Value         := INI.ReadInteger('SAT','PaginaDeCodigo',0);
    sfeVersaoEnt.Text      := FloatToString( INI.ReadFloat('SAT','versaoDadosEnt', cversaoDadosEnt) );
    cbxUTF8.Checked   := INI.ReadBool('SAT','UTF8', True);
    cbxFormatXML.Checked   := INI.ReadBool('SAT','FormatarXML', True);
    cbxSalvarCFe.Checked     := INI.ReadBool('SAT','SalvarCFe', True);
    cbxSalvarCFeCanc.Checked := INI.ReadBool('SAT','SalvarCFeCanc', True);
    cbxSalvarEnvio.Checked   := INI.ReadBool('SAT','SalvarEnvio', True);
    cbxSepararPorCNPJSat.Checked:= INI.ReadBool('SAT','SepararPorCNPJ', True);
    cbxSepararPorMES.Checked := INI.ReadBool('SAT','SepararPorMES', True);
    sePagCodChange(Self);

    edtEmitCNPJSat.Text := INI.ReadString('Emit','CNPJ','');
    edtEmitIESat.Text   := INI.ReadString('Emit','IE','');
    edtEmitIMSat.Text   := INI.ReadString('Emit','IM','');
    cbxRegTributario.ItemIndex := INI.ReadInteger('Emit','RegTributario',0);
    cbxRegTribISSQN.ItemIndex  := INI.ReadInteger('Emit','RegTribISSQN',0);
    cbxIndRatISSQN.ItemIndex   := INI.ReadInteger('Emit','IndRatISSQN',0);

    edPathSalvarXmlSatEnvio.Text := INI.ReadString('SAT','PATH_SALVAR_ENVIO','');
    edPathSalvarXmlSatVenda.Text := INI.ReadString('SAT','PATH_SALVAR_VENDA','');

    edPathSalvarXmlSatCanc.Text := INI.ReadString('SAT','PATH_SALVAR_CANC','');

    edtSwHCNPJ.Text       := INI.ReadString('SwH','CNPJ','11111111111111');
    edtSwHAssinatura.Lines.Text := INI.ReadString('SwH','Assinatura',cAssinatura);

    seLargura.Value        := INI.ReadInteger('Fortes','Largura',ACBrSATExtratoFortes1.LarguraBobina);
    seMargemTopo.Value     := INI.ReadInteger('Fortes','MargemTopo',Trunc(ACBrSATExtratoFortes1.MargemSuperior));
    seMargemFundo.Value    := INI.ReadInteger('Fortes','MargemFundo',Trunc(ACBrSATExtratoFortes1.MargemInferior));
    seMargemEsquerda.Value := INI.ReadInteger('Fortes','MargemEsquerda',Trunc(ACBrSATExtratoFortes1.MargemEsquerda));
    seMargemDireita.Value  := INI.ReadInteger('Fortes','MargemDireita',Trunc(ACBrSATExtratoFortes1.MargemDireita));
    cbPreview.Checked      := INI.ReadBool('Fortes','Preview',True);

    sangria_suprim_preview.Checked := Ini.ReadBool('Fortes','Sangria_Suprimento_Preview', False);
    sangria_suprim_duasvias.Checked := Ini.ReadBool('Fortes','Sangria_Suprimento_DuasVias', False);
    chkInverterOrderImpressao.Checked := Ini.ReadBool('Fortes','InverterOrderImpressao', False);
    chkImprimirListageSuprimentos.Checked := Ini.ReadBool('Fortes','ImprimirListagemSuprimentos', False);
    chkImprimirListagemSangrias.Checked := Ini.ReadBool('Fortes','ImprimirListagemSangrias', False);
    chkSepararListagens.Checked    := Ini.ReadBool('Fortes','SepararListagens', False);
    chk_imprimir_fech_cego.Checked := Ini.ReadBool('Fortes','ImprimirFechamentoCego', False);

    lImpressora.Caption := INI.ReadString('Printer','Name','');

    rgRedeTipoInter.ItemIndex := INI.ReadInteger('Rede','tipoInter',0);
    rgRedeTipoLan.ItemIndex   := INI.ReadInteger('Rede','tipoLan',0);
    edRedeSSID.Text           := INI.ReadString('Rede','SSID','');
    cbxRedeSeg.ItemIndex      := INI.ReadInteger('Rede','seg',0);
    edRedeCodigo.Text         := INI.ReadString('Rede','codigo','');
    edRedeIP.Text             := INI.ReadString('Rede','lanIP','');
    edRedeMask.Text           := INI.ReadString('Rede','lanMask','');
    edRedeGW.Text             := INI.ReadString('Rede','lanGW','');
    edRedeDNS1.Text           := INI.ReadString('Rede','lanDNS1','');
    edRedeDNS2.Text           := INI.ReadString('Rede','lanDNS2','');
    edRedeUsuario.Text        := INI.ReadString('Rede','usuario','');
    edRedeSenha.Text          := INI.ReadString('Rede','senha','');
    cbxRedeProxy.ItemIndex    := INI.ReadInteger('Rede','proxy',0);
    edRedeProxyIP.Text        := INI.ReadString('Rede','proxy_ip','');
    edRedeProxyPorta.Value    := INI.ReadInteger('Rede','proxy_porta',0);
    edRedeProxyUser.Text      := INI.ReadString('Rede','proxy_user','');
    edRedeProxySenha.Text     := INI.ReadString('Rede','proxy_senha','');


    edCfopPadrao.Text     := INI.ReadString('Valores','CFOP','5102');


    cbxModelo_Ecf.ItemIndex := INI.ReadInteger('ECF', 'Modelo', cbxModelo_Ecf.ItemIndex);
    cbxModelo_EcfChange(nil);
    cbxPorta_Ecf.Text := INI.ReadString('ECF', 'Porta', cbxPorta_Ecf.Text);
    seTimeOut_Ecf.Value := INI.ReadInteger('ECF', 'TimeOut', seTimeOut_Ecf.Value);
    seIntervaloAposComando_Ecf.Value := INI.ReadInteger('ECF', 'IntervaloAposComando', seIntervaloAposComando_Ecf.Value);
    chTentar_Ecf.Checked := INI.ReadBool('ECF', 'TentarNovamente', chTentar_Ecf.Checked);
    chControlePorta_Ecf.Checked := INI.ReadBool('ECF', 'ControlePorta', true);
    chBloqueia_Ecf.Checked := INI.ReadBool('ECF', 'BloqueiaMouseTeclado', chBloqueia_Ecf.Checked);
    chExibeMsg_Ecf.Checked := INI.ReadBool('ECF', 'ExibeMsgAguarde', chExibeMsg_Ecf.Checked);
    chArredondaPorQtd_Ecf.Checked := INI.ReadBool('ECF', 'ArredondaPorQtd', chArredondaPorQtd_Ecf.Checked);
    chDescricaoGrande_Ecf.Checked := INI.ReadBool('ECF', 'DescricaoGrande', chDescricaoGrande_Ecf.Checked);
    chArredondamentoItemMFD_Ecf.Checked := INI.ReadBool('ECF', 'ArredondamentoItemMFD', chArredondamentoItemMFD_Ecf.Checked);
    chGavetaSinalInvertido_Ecf.Checked := INI.ReadBool('ECF', 'GavetaSinalInvertido', chGavetaSinalInvertido_Ecf.Checked);
    mMsg_Ecf.Text := StringReplace(INI.ReadString('ECF', 'MensagemAguarde', mMsg_Ecf.Text), '|', sLineBreak, [rfReplaceAll]);
    edLog_Ecf.Text := INI.ReadString('ECF', 'ArqLog', edLog.Text);
    ACBrECF1.Device.ParamsString := INI.ReadString('ECF', 'SerialParams', '');
    seBandWidth_Ecf.Value := INI.ReadInteger('ECF', 'BandWidth', seBandWidth_Ecf.Value);
    speLinBuf_Ecf.Value := INI.ReadInteger('ECF', 'Buffer', ACBrECF1.MaxLinhasBuffer);
    sePaginaCodigo_Ecf.Value := INI.ReadInteger('ECF', 'PaginadeCodigo', ACBrECF1.PaginaDeCodigo);

    edRecECFDinheiro.Text := INI.ReadString('ECF', 'RecECFDinheiro', edRecECFDinheiro.Text);
    edRecECFConvenio.Text := INI.ReadString('ECF', 'RecECFConvenio', edRecECFConvenio.Text);
    edRecECFCartaoCredito.Text := INI.ReadString('ECF', 'RecECFCartaoCredito', edRecECFCartaoCredito.Text);
    edRecECFCarteira.Text := INI.ReadString('ECF', 'RecECFCarteira', edRecECFCarteira.Text);
    edRecECFCartaoDebito.Text := INI.ReadString('ECF', 'RecECFCartaoDebito', edRecECFCartaoDebito.Text);
    edRecECFCheque.Text := INI.ReadString('ECF', 'RecECFCheque', edRecECFCheque.Text);

    edCaminhoFiscal.Text := INI.ReadString('ECF', 'CaminhoFiscal', edCaminhoFiscal.Text);
    edMD5.Text := INI.ReadString('ECF', 'MD5', edMD5.Text);

  finally
    Ini.Free;
  end;

end;

procedure TfrmConfig.LerConfiguracoes;
begin

  {: FOTO PRODUTOS}
  CheckHabilitaFotoProduto.Checked := StrToBool(LerINI(sConfiguracoes, 'PDV', 'HAB_IMG', BoolToStr(false)));

  {: LOGO MARCA}
  CheckLogoMarca.Checked := StrToBool(LerINI(sConfiguracoes, 'PDV', 'HAB_LOGO', BoolToStr(false)));
  EditLogoMarca.Text := LerINI(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', '');

  LerCaixa;

end;

procedure TfrmConfig.mMsg_EcfChange(Sender: TObject);
Var
  Msg: String;
  L: Integer;
begin
  Msg := '';
  For L := 0 to mMsg_Ecf.Lines.Count - 1 do
  begin
    Msg := Msg + mMsg_Ecf.Lines[L] + #10
  end;

  ACBrECF1.MsgAguarde := copy(Msg, 1, length(Msg) - 1);
end;

procedure TfrmConfig.PathClick(Sender: TObject);
var
  Dir: string;
begin
  if Length(TEdit(Sender).Text) <= 0 then
    Dir := ExtractFileDir(application.ExeName)
  else
    Dir := TEdit(Sender).Text;

  if SelectDirectory(Dir, [sdAllowCreate, sdPerformCreate, sdPrompt], SELDIRHELP) then
    TEdit(Sender).Text := Dir;
end;

procedure TfrmConfig.qrconfigAfterOpen(DataSet: TDataSet);
begin
  DBEdit8Change(nil);
end;

procedure TfrmConfig.qrconfigPRAZO_DEVOLUCAOGetText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  if (Sender.Value = NULL) or (Sender.Value = '0') then
    Text := '';
end;

procedure TfrmConfig.qrconfigPRAZO_DEVOLUCAOSetText(Sender: TField; const Text: string);
begin
  if Text.IsEmpty then
    Sender.Value := 0;
end;

procedure TfrmConfig.rgRedeTipoInterClick(Sender: TObject);
begin
  gbWiFi.Visible := (rgRedeTipoInter.ItemIndex = 1);
end;

procedure TfrmConfig.rgRedeTipoLanClick(Sender: TObject);
begin
  gbPPPoE.Visible := (rgRedeTipoLan.ItemIndex = 1);
  gbIPFix.Visible := (rgRedeTipoLan.ItemIndex = 2);
end;

procedure TfrmConfig.sbPathCanClick(Sender: TObject);
begin
  PathClick(edtPathCan);
end;

procedure TfrmConfig.sbPathEventoClick(Sender: TObject);
begin
  PathClick(edtPathEvento);
end;

procedure TfrmConfig.sbPathInuClick(Sender: TObject);
begin
  PathClick(edtPathInu);
end;

procedure TfrmConfig.sbPathNFeClick(Sender: TObject);
begin
  PathClick(edtPathNFe);
end;

procedure TfrmConfig.sbtnCaminhoCertClick(Sender: TObject);
begin
  OpenDialog1.Title := 'Selecione o Certificado';
  OpenDialog1.DefaultExt := '*.pfx';
  OpenDialog1.Filter := 'Arquivos PFX (*.pfx)|*.pfx|Todos os Arquivos (*.*)|*.*';
  OpenDialog1.InitialDir := ExtractFileDir(application.ExeName);
  if OpenDialog1.Execute then begin
    edtCaminho.Text := OpenDialog1.FileName;
  end;
end;

procedure TfrmConfig.sbtnPathSalvarClick(Sender: TObject);
begin
  PathClick(edtPathLogs);
end;

procedure TfrmConfig.seBandWidth_EcfChange(Sender: TObject);
begin
  ACBrECF1.Device.MaxBandwidth := seBandWidth_Ecf.Value;
end;

procedure TfrmConfig.seIntervaloAposComando_EcfChange(Sender: TObject);
begin
  ACBrECF1.IntervaloAposComando := seIntervaloAposComando_Ecf.Value;
end;

procedure TfrmConfig.sePagCodChange(Sender: TObject);
begin
  ACBrSAT1.Config.PaginaDeCodigo := sePagCod.Value;
  cbxUTF8.Checked := ACBrSAT1.Config.EhUTF8;
end;

procedure TfrmConfig.sePaginaCodigo_EcfChange(Sender: TObject);
begin
  ACBrECF1.PaginaDeCodigo := sePaginaCodigo_Ecf.Value;
end;

procedure TfrmConfig.seTimeOut_EcfChange(Sender: TObject);
begin
  ACBrECF1.TimeOut := seTimeOut_Ecf.Value;
end;

procedure TfrmConfig.SpeedButton1Click(Sender: TObject);
begin
  PathClick(edPathSalvarXmlSatEnvio);

end;

procedure TfrmConfig.SpeedButton2Click(Sender: TObject);
begin
  PathClick(edPathSalvarXmlSatVenda);

end;

procedure TfrmConfig.SpeedButton3Click(Sender: TObject);
begin
  PathClick(edPathSalvarXmlSatCanc);

end;

procedure TfrmConfig.SpeedButton4Click(Sender: TObject);
begin
  OpenURL(ExtractFilePath(Application.ExeName) + edLog_Ecf.Text);
end;

procedure TfrmConfig.speLinBuf_EcfChange(Sender: TObject);
begin
  ACBrECF1.MaxLinhasBuffer := speLinBuf_Ecf.Value;
end;

procedure TfrmConfig.spPathSchemasClick(Sender: TObject);
begin
  PathClick(edtPathSchemas);
end;

function TfrmConfig.ValidaDados: Boolean;
begin
  Result := True;
  if Trim(edTipoCupom.Text) = '' then begin
    Application.MessageBox('Informe o Tipo de Cupom Fiscal', 'Atenção!', MB_ICONINFORMATION);
    pageview1.ActivePage := PageSheet1;
    edTipoCupom.SetFocus;
    Result := False;
    Exit;
  end;
  if UpperCase(edTipoCupom.Text) = 'NFC-E' then begin
    if Trim(edtEmitCNPJ.Text) = '' then begin
      Application.MessageBox('Informe o CNPJ do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitCNPJ.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edtEmitRazao.Text) = '' then begin
      Application.MessageBox('Informe a Razão do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitRazao.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edtEmitFantasia.Text) = '' then begin
      Application.MessageBox('Informe o Nome Fantasia do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitFantasia.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edtEmitCodCidade.Text) = '' then begin
      Application.MessageBox('Informe o Código da Cidade do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitCodCidade.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edtEmitCodCidade.Text) = '' then begin
      Application.MessageBox('Informe o Código da Cidade do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitCodCidade.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edtEmitCidade.Text) = '' then begin
      Application.MessageBox('Informe o a Cidade do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitCidade.SetFocus;
      Result := False;
      Exit;
    end;

    if Trim(edtNumSerie.Text) = '' then begin
      Application.MessageBox('Informe o Número de Série do Certificado clicando no Botão na frente do campos', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts1;
      edtNumSerie.SetFocus;
      Result := False;
      Exit;
    end;

    if (rgTipoAmb.ItemIndex = 0) and (Trim(edtToken.Text) = '') then begin
      Application.MessageBox('Informe o Número CSC para emissão em Produção', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts2;
      edtToken.SetFocus;
      Result := False;
      Exit;
    end;

    if (rgTipoAmb.ItemIndex = 0) and (Trim(edtIdToken.Text) = '') then begin
      Application.MessageBox('Informe o ID CSC para emissão em Produção', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts1;
      edtIdToken.SetFocus;
      Result := False;
      Exit;
    end;

    if (ckSalvar.Checked) and (Trim(edtPathLogs.Text) = '') then begin
      Application.MessageBox('Informe o Caminho para salvar os XMLs', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts2;
      edtPathLogs.SetFocus;
      Result := False;
      Exit;
    end;

    if (ckSalvar.Checked) and not (DirectoryExists(Trim(edtPathLogs.Text))) then begin
      Application.MessageBox('Informe um Caminho valido para salvar os Envios e Respostas', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts2;
      edtPathLogs.SetFocus;
      Result := False;
      Exit;
    end;

    if not (DirectoryExists(Trim(edtPathNFe.Text))) then begin
      Application.MessageBox('Informe um Caminho valido para salvar os XMLs', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := TabSheet2;
      edtPathNFe.SetFocus;
      Result := False;
      Exit;
    end;

    if (Trim(edtPathSchemas.Text) = '') or not (DirectoryExists(Trim(edtPathSchemas.Text))) then begin
      Application.MessageBox('Informe um Caminho valido de localização do Schemas', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts2;
      edtPathSchemas.SetFocus;
      Result := False;
      Exit;
    end;

    if Trim(edtEmitUF.Text) = '' then begin
      Application.MessageBox('Informe o UF do Emitente', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edtEmitUF.SetFocus;
      Result := False;
      Exit;
    end;

    if Trim(edSequencia.Text) = '' then begin
      Application.MessageBox('Informe a Sequencia da NFC-e', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edSequencia.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edSerie.Text) = '' then begin
      Application.MessageBox('Informe a Série da NFC-e', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edSerie.SetFocus;
      Result := False;
      Exit;
    end;
    try
      StrToInt(edSequencia.Text);
    except
      Application.MessageBox('Informe uma Sequencia da NFC-e Válida', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts4;
      edSequencia.SetFocus;
      Result := False;
      Exit;
    end;
    if Trim(edCFOP.Text) = '' then begin
      Application.MessageBox('Informe o CFOP Padrão', 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet5;
      pgc1.ActivePage := ts2;
      edCFOP.SetFocus;
      Result := False;
      exit;
    end;
  end;
  if Trim(edCfopPadrao.Text) = '' then begin
    Application.MessageBox('Informe o CFOP Padrão', 'Atenção!', MB_ICONINFORMATION);
    pageview1.ActivePage := PageSheet6;
    PageControl1.ActivePage := TabSheet4;
    edCfopPadrao.SetFocus;
    Result := False;
    exit;
  end;
  if edLioberado.Checked then begin
    if cbxGP.ItemIndex = 0 then begin
      Application.MessageBox(PWideChar('TEF ativo!'+#13+'Não informado o Tipo do Tef!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet7;
      cbxGP.SetFocus;
      Result := False;
      exit;
    end;

    if (Trim(edtLoja.Text) = '') or (Trim(edtLoja.Text) = '0') then begin
      Application.MessageBox(PWideChar('TEF ativo!'+#13+'Não informado o Código da Loja!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet7;
      edtLoja.SetFocus;
      Result := False;
      exit;
    end;
    if (Trim(edtTerminal.Text) = '') or (Trim(edtTerminal.Text) = '0') then begin
      Application.MessageBox(PWideChar('TEF ativo!'+#13+'Não informado o Código do Terminal!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet7;
      edtTerminal.SetFocus;
      Result := False;
      exit;
    end;
    if (Trim(edCnpj.Text) = '') or (Trim(edCnpj.Text) = '0') then begin
      Application.MessageBox(PWideChar('TEF ativo!'+#13+'Não informado o C.N.P.J da SoftHouse!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet7;
      edCnpj.SetFocus;
      Result := False;
      exit;
    end;
    frmPrincipal.CheckCNPJ.Documento := edCnpj.Text;
    if not frmPrincipal.CheckCNPJ.Validar then begin
      Application.MessageBox(PWideChar('TEF ativo!'+#13+'C.N.P.J da SoftHouse inválido!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet7;
      edCnpj.SetFocus;
      Result := False;
      exit;
    end;
  end;
  if UpperCase(edTipoCupom.Text) = 'PAF-ECF' then begin
    if Trim(edRecECFDinheiro.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em dinheiro não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFDinheiro.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFDinheiro.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em dinheiro não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFDinheiro.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFConvenio.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em convênio não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFConvenio.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFCartaoCredito.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em cartão de crédito não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFCartaoCredito.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFCarteira.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em carteira não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFCarteira.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFCartaoDebito.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em cartão de débito não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFCartaoDebito.SetFocus;
      Result := False;
      exit;
    end;
    if Trim(edRecECFCheque.Text) = '' then begin
      Application.MessageBox(PWideChar('PAF-ECF Ativo!'+#13+'Vinculo do receimento em cheque não realizado!'), 'Atenção!', MB_ICONINFORMATION);
      pageview1.ActivePage := PageSheet8;
      edRecECFCheque.SetFocus;
      Result := False;
      exit;
    end;
  end;
end;


procedure TfrmConfig.EditLogoMarcaClickBtn(Sender: TObject);
begin
  if OpenPictureDialog1.Execute then
    EditLogoMarca.Text := OpenPictureDialog1.FileName;
end;

procedure TfrmConfig.edLog_EcfChange(Sender: TObject);
begin
  ACBrECF1.ArqLOG := edLog_Ecf.Text;
end;

procedure TfrmConfig.edOperador_ecfChange(Sender: TObject);
begin
  ACBrECF1.Operador := sNome_Operador;
end;

procedure TfrmConfig.edRecCartaoCreditoChange(Sender: TObject);
begin
  edRecECFCartaoCredito.Text := Copy(edRecCartaoCredito.Text,0,2);
end;

procedure TfrmConfig.edRecCartaoDebitoChange(Sender: TObject);
begin
  edRecECFCartaoDebito.Text := Copy(edRecCartaoDebito.Text,0,2);
end;

procedure TfrmConfig.edRecCarteiraChange(Sender: TObject);
begin
  edRecECFCarteira.Text := Copy(edRecCarteira.Text,0,2);
end;

procedure TfrmConfig.edRecChequeChange(Sender: TObject);
begin
  edRecECFCheque.Text := Copy(edRecCheque.Text,0,2);
end;

procedure TfrmConfig.edRecConvenioChange(Sender: TObject);
begin
  edRecECFConvenio.Text := Copy(edRecConvenio.Text,0,2);
end;

procedure TfrmConfig.edRecDinheiroChange(Sender: TObject);
begin
  edRecECFDinheiro.Text := Copy(edRecDinheiro.Text,0,2);
end;

procedure TfrmConfig.edTipoCupomChange(Sender: TObject);
begin
  if edTipoCupom.ItemIndex = 0 then begin

  end;
end;

procedure TfrmConfig.cbxSepararPorCNPJSatClick(Sender: TObject);
begin
  ACBrSAT1.ConfigArquivos.SepararPorCNPJ := cbxSepararPorCNPJSat.Checked;
end;

procedure TfrmConfig.CheckLogoMarcaClick(Sender: TObject);
begin
  EditLogoMarca.Enabled := CheckLogoMarca.Checked;
end;

procedure TfrmConfig.chExibeMsg_EcfClick(Sender: TObject);
begin
  ACBrECF1.ExibeMensagem := chExibeMsg_Ecf.Checked;
end;

procedure TfrmConfig.chGavetaSinalInvertido_EcfClick(Sender: TObject);
begin
  ACBrECF1.GavetaSinalInvertido := chGavetaSinalInvertido_Ecf.Checked;
end;

procedure TfrmConfig.chTentar_EcfClick(Sender: TObject);
begin
  ACBrECF1.ReTentar := chTentar_Ecf.Checked;
end;

procedure TfrmConfig.ConfiguraECF;
var
  IniFile: string;
  Ini: TIniFile;
  Msg: String;
  L: Integer;
begin

  IniFile := sConfiguracoes;
  Ini := TIniFile.Create(IniFile);
  try
    ACBrECF1.Modelo := TACBrECFModelo(cbxModelo_Ecf.ItemIndex);
    ACBrECF1.Porta := cbxPorta_Ecf.Text;
    ACBrECF1.TimeOut := seTimeOut_Ecf.Value * 1000;
    ACBrECF1.IntervaloAposComando := seIntervaloAposComando_Ecf.Value;
    ACBrECF1.ReTentar := chTentar_Ecf.Checked;
    ACBrECF1.ControlePorta := chControlePorta_Ecf.Checked;
    ACBrECF1.BloqueiaMouseTeclado := chBloqueia_Ecf.Checked;
    ACBrECF1.ExibeMensagem := chExibeMsg_Ecf.Checked;
    ACBrECF1.ArredondaPorQtd := chArredondaPorQtd_Ecf.Checked;
    ACBrECF1.DescricaoGrande := chDescricaoGrande_Ecf.Checked;
    ACBrECF1.ArredondaItemMFD := chArredondamentoItemMFD_Ecf.Checked;
    ACBrECF1.GavetaSinalInvertido := chGavetaSinalInvertido_Ecf.Checked;
    ACBrECF1.ArqLOG := edLog_Ecf.Text;
    ACBrECF1.Operador := Combo_Usuario.Text;
    ACBrECF1.Device.MaxBandwidth := seBandWidth_Ecf.Value;
    ACBrECF1.MaxLinhasBuffer := speLinBuf_Ecf.Value;
    ACBrECF1.PaginaDeCodigo := sePaginaCodigo_Ecf.Value;
    Msg := '';
    For L := 0 to mMsg_Ecf.Lines.Count - 1 do
    begin
      Msg := Msg + mMsg_Ecf.Lines[L] + #10
    end;
    ACBrECF1.MsgAguarde := copy(Msg, 1, length(Msg) - 1);
  except
    on e:Exception do begin
      Application.MessageBox(PWideChar('Ocorreu um erro ao aplicar as configurações da ECF.'+#13+e.Message),'Atenção!',MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmConfig.cxButton10Click(Sender: TObject);
begin
  v_borda_superior.Height := v_borda_superior.Height - edIncremento.Value;
  if v_borda_superior.Height < 20 then
    v_borda_superior.Height := 20;
  l_borda_superior.Caption := 'Borda Superior ('+IntToStr(v_borda_superior.Height)+')';
end;

procedure TfrmConfig.cxButton12Click(Sender: TObject);
begin
  v_borda_inferior.Height := v_borda_inferior.Height - edIncremento.Value;
  if v_borda_inferior.Height < 20 then
    v_borda_inferior.Height := 20;
  l_borda_inferior.Caption := 'Borda Inferior ('+IntToStr(v_borda_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton13Click(Sender: TObject);
begin
  v_borda_inferior.Height := v_borda_inferior.Height + edIncremento.Value;
  l_borda_inferior.Caption := 'Borda Inferior ('+IntToStr(v_borda_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton14Click(Sender: TObject);
begin
  v_borda_esquerda.Width := v_borda_esquerda.Width - edIncremento.Value;
  if v_borda_esquerda.Width < 20 then
    v_borda_esquerda.Width := 20;
  l_borda_esquerda.Caption := 'Borda Esquerda ('+IntToStr(v_borda_esquerda.Width)+')';
end;

procedure TfrmConfig.cxButton15Click(Sender: TObject);
begin
  v_borda_esquerda.Width := v_borda_esquerda.Width + edIncremento.Value;
  l_borda_esquerda.Caption := 'Borda Esquerda ('+IntToStr(v_borda_esquerda.Width)+')';
end;

procedure TfrmConfig.cxButton16Click(Sender: TObject);
begin
  v_borda_direita.Width := v_borda_direita.Width + edIncremento.Value;
  l_borda_direita.Caption := 'Borda Direita ('+IntToStr(v_borda_direita.Width)+')';
end;

procedure TfrmConfig.cxButton17Click(Sender: TObject);
begin
  v_borda_direita.Width := v_borda_direita.Width - edIncremento.Value;
  if v_borda_direita.Width < 20 then
    v_borda_direita.Width := 20;
  l_borda_direita.Caption := 'Borda Direita ('+IntToStr(v_borda_direita.Width)+')';
end;

procedure TfrmConfig.cxButton18Click(Sender: TObject);
begin
  v_borda_superior.Height := v_borda_superior.Height + edIncremento.Value;
  l_borda_superior.Caption := 'Borda Superior ('+IntToStr(v_borda_superior.Height)+')';
end;

procedure TfrmConfig.cxButton1Click(Sender: TObject);
var
  A: Integer;
begin
  if bAtivar.Caption = 'Ativar' then begin
    try
      ConfiguraECF;
      if cbxModelo_Ecf.ItemIndex = 0 then
        if not ACBrECF1.AcharECF(true, False) then
        begin
          MessageDlg('Nenhum ECF encontrado.', mtInformation, [mbOk], 0);
          exit;
        end;

      ACBrECF1.Ativar;
      ACBrECF1.LerTotaisFormaPagamento;
      edRecDinheiro.Items.Clear;
      edRecConvenio.Items.Clear;
      edRecCartaoCredito.Items.Clear;
      edRecCarteira.Items.Clear;
      edRecCartaoDebito.Items.Clear;
      edRecCheque.Items.Clear;
      edRecDinheiro.Clear;
      edRecConvenio.Clear;
      edRecCartaoCredito.Clear;
      edRecCarteira.Clear;
      edRecCartaoDebito.Clear;
      edRecCheque.Clear;

      for A := 0 to ACBrECF1.FormasPagamento.Count - 1 do begin
        if ACBrECF1.FormasPagamento[A].Descricao <> '' then begin
          edRecDinheiro.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
          edRecConvenio.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
          edRecCartaoCredito.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
          edRecCarteira.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
          edRecCartaoDebito.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
          edRecCheque.Items.Add(ACBrECF1.FormasPagamento[A].Indice + ' - ' + ACBrECF1.FormasPagamento[A].Descricao);
        end;
      end;
      bAtivar.Caption := 'Desativar';
      Application.MessageBox('Conectado com sucesso a ECF!','Atenção!',MB_ICONINFORMATION);
    except
      on e:Exception do begin
        Application.MessageBox(PWideChar('Ocorreu um erro ao tentar se conectar a ECF!'+#13+e.Message),'Atenção!',MB_ICONINFORMATION);
      end;
    end;
  end else begin
    ACBrECF1.Desativar;
    bAtivar.Caption := 'Ativar';
  end;
end;

procedure TfrmConfig.cxButton21Click(Sender: TObject);
begin
  v_barra_inferior.Height := v_barra_inferior.Height + edIncremento.Value;
  l_barra_inferior.Caption := 'Barra Inferior ('+IntToStr(v_barra_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton22Click(Sender: TObject);
begin
  v_barra_inferior.Height := v_barra_inferior.Height - edIncremento.Value;
  if v_barra_inferior.Height < 20 then
    v_barra_inferior.Height := 20;
  l_barra_inferior.Caption := 'Barra Inferior ('+IntToStr(v_barra_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton23Click(Sender: TObject);
begin
  v_barra_superior.Height := v_barra_superior.Height - edIncremento.Value;
  if v_barra_superior.Height < 20 then
    v_barra_superior.Height := 20;
  l_barra_superior.Caption := 'Barra Superior ('+IntToStr(v_barra_superior.Height)+')';
end;

procedure TfrmConfig.cxButton24Click(Sender: TObject);
begin
  v_barra_superior.Height := v_barra_superior.Height + edIncremento.Value;
  l_barra_superior.Caption := 'Barra Superior ('+IntToStr(v_barra_superior.Height)+')';
end;

procedure TfrmConfig.cxButton2Click(Sender: TObject);
begin
  v_quadro_esquerdo.Width := v_quadro_esquerdo.Width - edIncremento.Value;
  if v_quadro_esquerdo.Width < 232 then
    v_quadro_esquerdo.Width := 232;
  l_quadro_esquerdo.Caption := 'Quadro Esquerdo ('+IntToStr(v_quadro_esquerdo.Width)+')';
  l_quadro_direito.Caption := 'Quadro Direito ('+IntToStr(v_quadro_direito.Width)+')';
end;

procedure TfrmConfig.cxButton3Click(Sender: TObject);
begin
  v_quadro_esquerdo.Width := v_quadro_esquerdo.Width + edIncremento.Value;
  if v_quadro_esquerdo.Width > 500 then
    v_quadro_esquerdo.Width := 500;
  l_quadro_esquerdo.Caption := 'Quadro Esquerdo ('+IntToStr(v_quadro_esquerdo.Width)+')';
  l_quadro_direito.Caption := 'Quadro Direito ('+IntToStr(v_quadro_direito.Width)+')';
end;

procedure TfrmConfig.cxButton4Click(Sender: TObject);
begin
  v_espacamento_superior.Height := v_espacamento_superior.Height + edIncremento.Value;
  l_espacamento_superior.Caption := 'Espaçamento Superior ('+IntToStr(v_espacamento_superior.Height)+')';
end;

procedure TfrmConfig.cxButton5Click(Sender: TObject);
begin
  v_espacamento_central.Width := v_espacamento_central.Width - edIncremento.Value;
  if v_espacamento_central.Width < 20 then
    v_espacamento_central.Width := 20;
  l_espacamento_central.Caption := 'Espaçamento Central ('+IntToStr(v_espacamento_central.Width)+')';
end;

procedure TfrmConfig.cxButton6Click(Sender: TObject);
begin
  v_espacamento_superior.Height := v_espacamento_superior.Height - edIncremento.Value;
  if v_espacamento_superior.Height < 20 then
    v_espacamento_superior.Height := 20;
  l_espacamento_superior.Caption := 'Espaçamento Superior ('+IntToStr(v_espacamento_superior.Height)+')';
end;

procedure TfrmConfig.cxButton7Click(Sender: TObject);
begin
  v_espacamento_inferior.Height := v_espacamento_inferior.Height + edIncremento.Value;
  l_espacamento_inferior.Caption := 'Espaçamento Inferior ('+IntToStr(v_espacamento_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton8Click(Sender: TObject);
begin
  v_espacamento_inferior.Height := v_espacamento_inferior.Height - edIncremento.Value;
  if v_espacamento_inferior.Height < 20 then
    v_espacamento_inferior.Height := 20;
  l_espacamento_inferior.Caption := 'Espaçamento Inferior ('+IntToStr(v_espacamento_inferior.Height)+')';
end;

procedure TfrmConfig.cxButton9Click(Sender: TObject);
begin
  v_espacamento_central.Width := v_espacamento_central.Width + edIncremento.Value;
  l_espacamento_central.Caption := 'Espaçamento Central ('+IntToStr(v_espacamento_central.Width)+')';
end;

procedure TfrmConfig.cbxModelo_EcfChange(Sender: TObject);
begin
  try
    ACBrECF1.Modelo := TACBrECFModelo(cbxModelo_Ecf.ItemIndex);
  except
    cbxModelo_Ecf.ItemIndex := Integer(ACBrECF1.Modelo);
    raise;
  end;
end;

procedure TfrmConfig.cbxPorta_EcfChange(Sender: TObject);
begin
  try
    ACBrECF1.Porta := cbxPorta_Ecf.Text;
  finally
    cbxPorta_Ecf.Text := ACBrECF1.Porta;
  end;
end;

procedure TfrmConfig.cbxRedeProxyChange(Sender: TObject);
begin
  edRedeProxyIP.Enabled    := (cbxRedeProxy.ItemIndex > 0);
  edRedeProxyPorta.Enabled := edRedeProxyIP.Enabled;
  edRedeProxyUser.Enabled  := edRedeProxyIP.Enabled;
  edRedeProxySenha.Enabled := edRedeProxyIP.Enabled;
end;

procedure TfrmConfig.cbxSalvarCFeCancClick(Sender: TObject);
begin
  ACBrSAT1.ConfigArquivos.SalvarCFeCanc := cbxSalvarCFeCanc.Checked;
end;

procedure TfrmConfig.cbxSalvarCFeClick(Sender: TObject);
begin
  ACBrSAT1.ConfigArquivos.SalvarCFe := cbxSalvarCFe.Checked;
end;

procedure TfrmConfig.cbxSalvarEnvioClick(Sender: TObject);
begin
  ACBrSAT1.ConfigArquivos.SalvarEnvio := cbxSalvarEnvio.Checked;
end;

procedure TfrmConfig.cbxSepararPorMESClick(Sender: TObject);
begin
  ACBrSAT1.ConfigArquivos.SepararPorMes := cbxSepararPorMES.Checked;
end;

procedure TfrmConfig.cbxUTF8Click(Sender: TObject);
begin
  ACBrSAT1.Config.EhUTF8 := cbxUTF8.Checked;
  sePagCod.Value := ACBrSAT1.Config.PaginaDeCodigo;
end;

procedure TfrmConfig.chArredondamentoItemMFD_EcfClick(Sender: TObject);
begin
  ACBrECF1.ArredondaItemMFD := chArredondamentoItemMFD_Ecf.Checked;
end;

procedure TfrmConfig.chArredondaPorQtd_EcfClick(Sender: TObject);
begin
  ACBrECF1.ArredondaPorQtd := chArredondaPorQtd_Ecf.Checked;
end;

procedure TfrmConfig.chBloqueia_EcfClick(Sender: TObject);
begin
  ACBrECF1.BloqueiaMouseTeclado := chBloqueia_Ecf.Checked;
end;

procedure TfrmConfig.chControlePorta_EcfClick(Sender: TObject);
begin
  ACBrECF1.ControlePorta := chControlePorta_Ecf.Checked;
end;

procedure TfrmConfig.chDescricaoGrande_EcfClick(Sender: TObject);
begin
  ACBrECF1.DescricaoGrande := chDescricaoGrande_Ecf.Checked;
end;

procedure TfrmConfig.LeDadosRedeSAT;
begin
  with ACBrSAT1.Rede do
  begin
    rgRedeTipoInter.ItemIndex := Integer(tipoInter);
    edRedeSSID.Text           := SSID ;
    cbxRedeSeg.ItemIndex      := Integer(seg) ;
    edRedeCodigo.Text         := codigo ;
    rgRedeTipoLan.ItemIndex   := Integer(tipoLan);
    edRedeIP.Text             := lanIP;
    edRedeMask.Text           := lanMask;
    edRedeGW.Text             := lanGW;
    edRedeDNS1.Text           := lanDNS1;
    edRedeDNS2.Text           := lanDNS2;
    edRedeUsuario.Text        := usuario;
    edRedeSenha.Text          := senha;
    cbxRedeProxy.ItemIndex    := proxy;
    edRedeProxyIP.Text        := proxy_ip;
    edRedeProxyPorta.Value    := proxy_porta;
    edRedeProxyUser.Text      := proxy_user;
    edRedeProxySenha.Text     := proxy_senha;
  end;

end;

procedure TfrmConfig.LerCaixa;
begin

  frmmodulo.query.close;
  frmmodulo.query.sql.clear;
  frmmodulo.query.sql.add('select * from adm');
  frmmodulo.query.sql.add('where codigo < ' + QuotedStr('100'));
  frmmodulo.query.sql.add('order by codigo');
  frmmodulo.query.open;

  Combo_Usuario.Items.Clear;

  while not frmmodulo.query.Eof do begin
    if (frmmodulo.query.Fieldbyname('INFO1').AsString <> 'DELPHOS_1') and (frmmodulo.query.Fieldbyname('CODIGO').AsInteger <> 99)  then
      Combo_Usuario.Items.Add(frmmodulo.query.Fieldbyname('INFO1').AsString);
    frmmodulo.query.Next;
    Application.ProcessMessages;
  end;

  Combo_Usuario.Text := LerIni(sConfiguracoes, 'LOGIN', 'USUARIO', '');
  EditSenhaUsuario.Text := Cript('D', LerIni(sConfiguracoes, 'LOGIN', 'SENHA', ''));
end;

procedure TfrmConfig.btn1Click(Sender: TObject);
begin
  if not ValidaDados then
    Exit;
  if bAtivar.Caption <> 'Ativar' then begin
    bAtivar.Click;
  end;

  GravaDados;

  bEdicaoRelatorio := CheckBox1.Checked;

  Close;
end;

procedure TfrmConfig.btn2Click(Sender: TObject);
begin
  if qrconfig.state = dsedit then
    qrconfig.cancel;
  close;

end;

procedure TfrmConfig.btnGetCertClick(Sender: TObject);
begin
//{$IFNDEF ACBrNFeOpenSSL}
  edtNumSerie.Text := ACBrNFe1.SSL.SelecionarCertificado; //ACBrNFe1.Configuracoes.Certificados.SelecionarCertificado;
//{$ENDIF}
end;

procedure TfrmConfig.btnImpEmpresaClick(Sender: TObject);
begin
  with frmModulo do begin
    qrfilial.Close;
    qrfilial.Open;
    edtEmitCNPJ.Text := SN(qrfilialCNPJ.AsString);
    edtEmitIE.Text := SN(qrfilialIE.AsString);
    edtEmitIM.Text := qrfilialINSC_MUNICIPAL.AsString;
    edtEmitRazao.Text := qrfilialFILIAL.AsString;
    edtEmitFantasia.Text := qrfilialFANTASIA.AsString;
    edtEmitFone.Text := qrfilialTELEFONE.AsString;
    edtEmitCEP.Text := SN(qrfilialTELEFONE.AsString);
    edtEmitLogradouro.Text := qrfilialENDERECO.AsString;
    edtEmitNumero.Text := qrfilialNUMERO.AsString;
    edtEmitComp.Text := qrfilialCOMPLEMENTO.AsString;
    edtEmitBairro.Text := qrfilialBAIRRO.AsString;
    edtEmitCodCidade.Text := qrfilialCOD_MUNICIPIO_IBGE.AsString;
    edtEmitCidade.Text := qrfilialCIDADE.AsString;
    edtEmitUF.Text := qrfilialUF.AsString;
  end;
end;

procedure TfrmConfig.btnSerialClick(Sender: TObject);
Var
  frConfiguraSerial: TfrConfiguraSerial;
begin
  frConfiguraSerial := TfrConfiguraSerial.Create(Self);

  try
    frConfiguraSerial.Device.Porta := ACBrECF1.Device.Porta;
    frConfiguraSerial.cmbPortaSerial.Text := cbxPorta.Text;
    frConfiguraSerial.Device.ParamsString := ACBrECF1.Device.ParamsString;

    if frConfiguraSerial.ShowModal = mrOk then
    begin
      cbxPorta.Text := frConfiguraSerial.Device.Porta;
      ACBrECF1.Device.ParamsString := frConfiguraSerial.Device.ParamsString;
    end;
  finally
    FreeAndNil(frConfiguraSerial);
  end;
end;

procedure TfrmConfig.Button1Click(Sender: TObject);
begin
  Combo_Usuario.Text := '';
  EditSenhaUsuario.Text := '';
end;

procedure TfrmConfig.cbCryptLibChange(Sender: TObject);
begin
  try
    if cbCryptLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLCryptLib := TSSLCryptLib(cbCryptLib.ItemIndex);
  finally
    AtualizaSSLLibsCombo;
  end;
end;

procedure TfrmConfig.cbHttpLibChange(Sender: TObject);
begin
  try
    if cbHttpLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLHttpLib := TSSLHttpLib(cbHttpLib.ItemIndex);
  finally
    AtualizaSSLLibsCombo;
  end;
end;

procedure TfrmConfig.cbSSLLibChange(Sender: TObject);
begin
  try
    if cbSSLLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLLib := TSSLLib(cbSSLLib.ItemIndex);
  finally
    AtualizaSSLLibsCombo;
  end;
end;

procedure TfrmConfig.cbSSLTypeChange(Sender: TObject);
begin
  if cbSSLType.ItemIndex <> -1 then
    ACBrNFe1.SSL.SSLType := TSSLType(cbSSLType.ItemIndex);
end;

procedure TfrmConfig.cbXmlSignLibChange(Sender: TObject);
begin
  try
    if cbXmlSignLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLXmlSignLib := TSSLXmlSignLib(cbXmlSignLib.ItemIndex);
  finally
    AtualizaSSLLibsCombo;
  end;
end;

procedure TfrmConfig.DBEdit8Change(Sender: TObject);
begin
  label95.Enabled := not((Trim(DBEdit8.Text) = '') or (StrToInt(DBEdit8.Text) = 0));
end;

procedure TfrmConfig.DBEdit8KeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', #8]) then
    Key := #0;
end;

end.

