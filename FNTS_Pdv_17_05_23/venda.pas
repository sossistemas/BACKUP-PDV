unit venda;

interface

uses
  Windows, Messages, Variants, Classes, Graphics, Controls, Vcl.Forms, Dialogs,
  StdCtrls, pcnConversaoNFe, Thread.CheckConnection, Data.DB, MemDS, DBAccess,
  Uni, System.Actions, Vcl.ActnList, Vcl.XPStyleActnCtrls, Vcl.ActnMan, Vcl.Menus,
  AdvMenus, Vcl.ExtCtrls, NxColumnClasses, AdvGlowButton, TFlatComboBoxUnit,
  AdvTimePickerDropDown, RzEdit, RzPanel, NxColumns, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, Vcl.Mask, dxGDIPlusClasses,
  AdvOfficeImage, AdvReflectionLabel, RzLabel, acPNG, TFlatPanelUnit, Vcl.FileCtrl,
  pcnAuxiliar, principal, IdIPWatch, pcnConversao, ACBrUtil, ACBrBAL, Clipbrd,
  JvComponentBase, JvThreadTimer, System.SysUtils, FlCtrlEx, JPEG, Consts,
  AdvMetroButton, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  dxStatusBar, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView, cxGrid,
  cxGridBandedTableView, cxGridDBBandedTableView, cxImageComboBox, System.ImageList,
  Vcl.ImgList, uBandeira, frxClass, frxDBSet, Datasnap.DBClient, Vcl.DBCGrids,
  cxEdit, Vcl.Touch.Keyboard, Vcl.Buttons, Extra, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, AdvSmoothTouchKeyBoard, Recebimento, cxButtonEdit, Vcl.Touch.GestureMgr,
  cxContainer, cxLabel, frxDesgn, ACBrSAT, ACBrBase, ACBrSATExtratoClass,
  ACBrSATExtratoReportClass, ACBrSATExtratoFortesFr, cxMemo, MontaPizza,
  ACBrDFeReport, ACBrTEFD, ACBrTEFDClass, ACBrTEFDCliSiTef, System.DateUtils,
  ACBrPosPrinter, ACBrECF, ACBrDevice, ErroEnvioNfce, ACBrDFeUtil, dxSkinsCore, dxSkinsDefaultPainters, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog, cxImageList, CurvyControls, JvExExtCtrls, JvExtComponent, JvPanel,
  W7Classes, W7Panels;

const
  OffsetMemoryStream: Int64 = 0;

type
  // tipo de imagens do dislplay
  Tipo_Imagem = (tiNenhum, tiLivre, tiErro, tiOk, tiAlerta, tiInfo, tiFechado, tiExcluir, tiPergunta, tiPgto, tiCliente, tiPeso);
  TTipoEtiqueta = (tePeso, teTotal);

  TDirecao =
  (dAlternar
  ,dAlmentar
  ,dDiminuir
  );

  TSat = record
    Validou: Boolean;
    Numero: Integer;
    Codigo: Integer;
    Mesagem: string;
  end;

  TRecebeCupomCredito = record
    id: Integer;
    Cupom: string;
    Valor: Double;
  end;

  TLacValorAutomatico = record
    ValorDesconto,
    ValorAcressimo:Double;
  end;

  TQuantidade = record
    Vendido: Real;
    ItensVolume: Real;
  end;

  TVolume = record
    ID: Integer;
    Item: Integer;
    Quantidade: TQuantidade;
    Unidade: String;
    Desconto: Real;
    Acrescimo: Real;
    PrecoVenda: Real;
    PorVolume: Integer;
  end;

  TProduto = record
    Codigo: Integer;
    QtdVenda: Real;
    Volume: array of TVolume;
  end;

  tpPagamento = (tpDinheiro, tpCartao);


  TfrmVenda = class(TForm)
    pn_tampa: TFlatPanel;
    fl: TFileListBoxEx;
    pnFundo: TPanel;
    pn_principal: TPanel;
    imgFundo: TImage;
    ed_foco: TEdit;
    pop_principal: TAdvPopupMenu;
    Opes1: TMenuItem;
    PesquisadeProdutos1: TMenuItem;
    FecharCupom1: TMenuItem;
    CancelarCupom1: TMenuItem;
    CancelarItem1: TMenuItem;
    ConsultadePreos1: TMenuItem;
    ConsultadeClientes1: TMenuItem;
    menu_Prevenda: TMenuItem;
    AcionaBalana1: TMenuItem;
    AcionaGaveta1: TMenuItem;
    Cupons1: TMenuItem;
    Atendimento1: TMenuItem;
    Oramento1: TMenuItem;
    menu_mesa: TMenuItem;
    Caixa1: TMenuItem;
    FecharCaixa1: TMenuItem;
    Suprimento1: TMenuItem;
    Sangria1: TMenuItem;
    CaixaIndisponvel1: TMenuItem;
    rocarOperador1: TMenuItem;
    N2: TMenuItem;
    Retornar1: TMenuItem;
    Sair1: TMenuItem;
    timer_carga: TTimer;
    TimerTroco: TTimer;
    query: TUniQuery;
    query2: TUniQuery;
    qrForma: TUniQuery;
    query3: TUniQuery;
    qrapoio: TUniQuery;
    qrProdNCFE: TUniQuery;
    QRCSOSN: TUniQuery;
    tmContigencia: TJvThreadTimer;
    qrContingencia: TUniQuery;
    C1: TMenuItem;
    UniDataSource1: TUniDataSource;
    cxStyleRepository1: TcxStyleRepository;
    stLinhas: TcxStyle;
    stTitulo: TcxStyle;
    ImageList1: TImageList;
    texto: TcxStyle;
    mnDelivery: TMenuItem;
    G1: TMenuItem;
    qrServidor: TUniQuery;
    frComprovanteDebito: TfrxReport;
    cdsParcelas: TClientDataSet;
    cdsParcelasVencimento: TDateField;
    cdsParcelasValor: TFloatField;
    frParcelas: TfrxDBDataset;
    cdsParcelasParcela: TStringField;
    RelCupomNF: TfrxReport;
    frCupomNF: TfrxDBDataset;
    qrCupomNF: TUniQuery;
    frFilial: TfrxDBDataset;
    frImpForma: TfrxDBDataset;
    qrImpForma: TUniQuery;
    qrImpFormaCOD_FORMA: TStringField;
    qrImpFormaFORMA_DESCRICAO: TStringField;
    qrImpFormaFORMA_PRESTACAO: TIntegerField;
    qrImpFormaFORMA_VALOR: TFloatField;
    RelCupomNFRestaurante: TfrxReport;
    qrCupomNFCOD_CUPOM: TStringField;
    qrCupomNFNUMERO: TStringField;
    qrCupomNFDATA: TDateField;
    qrCupomNFHORA: TTimeField;
    qrCupomNFCUPOM_TOTAL: TFloatField;
    qrCupomNFCUPOM_PAGO: TFloatField;
    qrCupomNFCUPOM_TROCO: TFloatField;
    qrCupomNFCLIENTE: TStringField;
    qrCupomNFVENDEDOR: TStringField;
    qrCupomNFCPF_CONSUMIDOR: TStringField;
    qrCupomNFNOME_CONSUMIDOR: TStringField;
    qrCupomNFCOD_CAIXA: TIntegerField;
    qrCupomNFPEGO_PORQUEM: TStringField;
    qrCupomNFCOUVERT: TFloatField;
    qrCupomNFTAXA_GARCON: TFloatField;
    qrCupomNFCOD_ITEM: TStringField;
    qrCupomNFITEM: TStringField;
    qrCupomNFCOD_PRODUTO: TIntegerField;
    qrCupomNFNOME: TStringField;
    qrCupomNFCOD_BARRA: TStringField;
    qrCupomNFUNIDADE: TStringField;
    qrCupomNFQTDE: TFloatField;
    qrCupomNFITEM_UNITARIO: TFloatField;
    qrCupomNFITEM_SUBTOTAL: TFloatField;
    qrCupomNFITEM_DESCONTO: TFloatField;
    qrCupomNFITEM_ACRESCIMO: TFloatField;
    qrCupomNFITEM_TOTAL: TFloatField;
    qrGrupo: TUniQuery;
    qrGrupoCODIGO: TStringField;
    qrGrupoGRUPO: TStringField;
    qrGrupoFOTO: TBlobField;
    qrProduto: TUniQuery;
    qrProdutoCODIGO: TIntegerField;
    qrProdutoCOD_BARRA: TStringField;
    qrProdutoNOME: TStringField;
    qrProdutoFOTOBD: TBlobField;
    qrCupomNFSENHA: TIntegerField;
    frComanda: TfrxReport;
    frxComanda: TfrxDBDataset;
    qrComanda: TUniQuery;
    qrComandaSENHA: TIntegerField;
    qrComandaDATA: TDateField;
    qrComandaHORA: TTimeField;
    qrComandaCOD_PRODUTO: TIntegerField;
    qrComandaNOME: TStringField;
    qrComandaQTDE: TFloatField;
    qrComandaVALOR_UNITARIO: TFloatField;
    qrComandaVALOR_TOTAL: TFloatField;
    qrComandaEXTRA: TBlobField;
    qrTmp: TUniQuery;
    qrContingenciaNUMERO: TIntegerField;
    qrContingenciaDATA: TDateField;
    qrContingenciaTOTAL: TFloatField;
    qrContingenciaCLIENTE: TStringField;
    qrContingenciaCHAVE: TStringField;
    qrContingenciaXML: TStringField;
    qrContingenciaSITUACAO: TIntegerField;
    qrContingenciaTROCO: TFloatField;
    qrContingenciaEX: TIntegerField;
    qrContingenciaHORA: TStringField;
    qrContingenciaCONTINGENCIA: TStringField;
    qrContingenciaENVIADOCONTINGENCIA: TStringField;
    qrContingenciaMOTIVOCONTIGENCIA: TStringField;
    qrContingenciaXMLENVIO: TBlobField;
    qrContingenciaXMLCACNELAMENTO: TBlobField;
    qrContingenciaALTERADO: TStringField;
    qrContingenciaCUPOM: TStringField;
    mnEnviaContigencia: TMenuItem;
    qrCupomNFENDERECO_CONSUMIDOR: TStringField;
    Teclado: TAdvSmoothTouchKeyBoard;
    cdsRecebimento: TClientDataSet;
    cdsRecebimentoTipo: TStringField;
    cdsRecebimentoParcelas: TIntegerField;
    cdsRecebimentobandeira: TStringField;
    cdsRecebimentovalor: TFloatField;
    cdsRecebimentocliente: TStringField;
    cdsRecebimentoconvenio: TStringField;
    cdsRecebimentobanco: TIntegerField;
    TimerExtra: TTimer;
    cdsOBS: TClientDataSet;
    cdsOBSobs: TMemoField;
    frOBS: TfrxDBDataset;
    frxDesigner1: TfrxDesigner;
    RelFicha: TfrxReport;
    ACBrSATExtratoFortes1: TACBrSATExtratoFortes;
    ACBrSAT1: TACBrSAT;
    qrCupom: TUniQuery;
    QrCupom_Item: TUniQuery;
    qrcupom_forma: TUniQuery;
    Query1: TUniQuery;
    QrProdutoSAT: TUniQuery;
    mmSat: TMenuItem;
    stExtra: TcxStyle;
    qrCombo: TUniQuery;
    qrComboCOMBO: TStringField;
    qrProdComb: TUniQuery;
    qrLancaProdutos: TUniQuery;
    cdsLancaProdutos: TClientDataSet;
    cdsLancaProdutosCODPRO: TStringField;
    cdsLancaProdutosQTDE: TFloatField;
    qrLancaProdutosCOD_ITEM: TStringField;
    qrLancaProdutosCODPRO: TStringField;
    qrLancaProdutosQTDE: TFloatField;
    cdsLancaProdutosCOD_ITEM: TStringField;
    frImpressao: TfrxReport;
    qrCarne: TUniQuery;
    qrCarnePARCELA: TLargeintField;
    qrCarneEMISSAO: TDateField;
    qrCarneVENCIMENTO: TDateField;
    qrCarneVALOR: TFloatField;
    qrCarneCODIGO: TIntegerField;
    qrCarneNOME: TStringField;
    frCarne: TfrxDBDataset;
    qrEstoque: TUniQuery;
    qrEstoqueESTOQUE: TFloatField;
    qrEstoqueCSOSN: TStringField;
    qrEstoqueNCM: TStringField;
    qrEstoqueCOD_BARRA: TStringField;
    qrEstoqueINFORMA_CODIGO_BARRA_XML: TStringField;
    qrEstoqueCFOP: TStringField;
    DevoluoTroca1: TMenuItem;
    frCredito: TfrxDBDataset;
    relCredito: TfrxReport;
    qrImpCredito: TUniQuery;
    cdsDevolucao: TClientDataSet;
    cdsDevolucaocodcupom: TStringField;
    cdsDevolucaocodcliente: TIntegerField;
    cdsDevolucaoitem: TIntegerField;
    cdsDevolucaovalor: TFloatField;
    qrImpCreditoCUPOM_DEVOLUCAO: TStringField;
    qrImpCreditoDATA: TDateField;
    qrImpCreditoHORA: TStringField;
    qrImpCreditoVALOR: TFloatField;
    qrImpCreditoVALOR_UTILIZADO: TFloatField;
    qrImpCreditoNOME: TStringField;
    qrImpCreditoVENCIMENTO: TDateField;
    qrContingenciaXML_CANCELAMENTO: TStringField;
    qrContingenciaDATAHORACONTIMGENCIA: TDateTimeField;
    qrContingenciaSTATUS: TStringField;
    dsContingencia: TDataSource;
    qrContingenciaERRO: TBlobField;
    cdsRecebimentoCupomCreditoID: TIntegerField;
    cdsRecebimentoCupomCreditoCupom: TStringField;
    cdsRecebimentoCupomCreditoValor: TFloatField;
    C2: TMenuItem;
    cdsRecebimentotef_cnpj: TStringField;
    cdsRecebimentotef_codigo_autorizacao: TStringField;
    cdsRecebimentocod_adm_sat: TStringField;
    cpACBrECF: TACBrECF;
    cdsCrediario: TClientDataSet;
    cdsCrediarioparcela: TIntegerField;
    cdsCrediariocliente: TStringField;
    cdsCrediariovencimento: TDateField;
    cdsCrediariovalor: TFloatField;
    v_fundo: TRzPanel;
    Panel5: TRzPanel;
    v_barra_superior: TRzPanel;
    v_espacamento_superior: TRzPanel;
    Panel4: TRzPanel;
    v_quadro_direito: TRzPanel;
    v_quadro_esquerdo: TRzPanel;
    v_espacamento_central: TRzPanel;
    v_espacamento_inferior: TRzPanel;
    v_barra_inferior: TRzPanel;
    v_borda_inferior: TRzPanel;
    v_borda_esquerda: TRzPanel;
    v_borda_direita: TRzPanel;
    v_borda_superior: TRzPanel;
    pnBotoesInferior: TPanel;
    btnCancelaVenda: TAdvGlowButton;
    btnCancelaItem: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1DBTableView1ITEM: TcxGridDBColumn;
    cxGrid1DBTableView1COD_BARRA: TcxGridDBColumn;
    cxGrid1DBTableView1CODIGO: TcxGridDBColumn;
    cxGrid1DBTableView1NOME: TcxGridDBColumn;
    cxGrid1DBTableView1UNIDADE: TcxGridDBColumn;
    cxGrid1DBTableView1QTDE: TcxGridDBColumn;
    cxGrid1DBTableView1VALOR_UNITARIO: TcxGridDBColumn;
    cxGrid1DBTableView1VALOR_DESCONTO: TcxGridDBColumn;
    cxGrid1DBTableView1VALOR_ACRESCIMO: TcxGridDBColumn;
    cxGrid1DBTableView1VALOR_TOTAL: TcxGridDBColumn;
    cxGrid1DBBandedTableView1: TcxGridDBBandedTableView;
    cxGrid1DBBandedTableView1ITEM: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1COD_BARRA: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1NOME: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1UNIDADE: TcxGridDBBandedColumn;
    colBotDiminuir: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1QTDE: TcxGridDBBandedColumn;
    colBotAlmuntar: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1VALOR_UNITARIO: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1VALOR_DESCONTO: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1VALOR_ACRESCIMO: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1VALOR_TOTAL: TcxGridDBBandedColumn;
    cxGrid1DBBandedTableView1CANCELADO: TcxGridDBBandedColumn;
    colExtra: TcxGridDBBandedColumn;
    cxGrid1Level1: TcxGridLevel;
    pnBotoes: TPanel;
    AdvGlowButton2: TAdvGlowButton;
    AdvGlowButton3: TAdvGlowButton;
    AdvGlowButton4: TAdvGlowButton;
    cpBarra: TdxStatusBar;
    barraAtalho: TdxStatusBar;
    Panel12: TPanel;
    lbDevolucao: TLabel;
    Panel13: TPanel;
    lb_total: TRzLabel;
    lbTotal: TRzLabel;
    Panel14: TPanel;
    lb_item: TRzLabel;
    RzLabel7: TRzLabel;
    Panel15: TPanel;
    btnFechaCupom: TAdvGlowButton;
    lbCupom: TcxLabel;
    Panel7: TPanel;
    lb_produto: TRzLabel;
    v_tipo_normal: TRzPanel;
    v_tipo_normal_campos: TRzPanel;
    v_tipo_normal_campos2: TRzPanel;
    v_tipo_normal_campos1: TRzPanel;
    pnBarra: TPanel;
    lbCodBarras: TRzLabel;
    ed_barra: TRzEdit;
    pnQuantidade: TPanel;
    RzLabel3: TRzLabel;
    Panel9: TPanel;
    btnDimQtd: TSpeedButton;
    btnAumQtd: TSpeedButton;
    ed_qtde: TRzNumericEdit;
    img_logo: TAdvOfficeImage;
    img_produto: TAdvOfficeImage;
    pnUnitario: TPanel;
    RzLabel1: TRzLabel;
    ed_unitario: TRzNumericEdit;
    pnTotal: TPanel;
    RzLabel4: TRzLabel;
    ed_total_item: TRzNumericEdit;
    Bevel1: TBevel;
    Bevel2: TBevel;
    v_tipo_touch: TRzPanel;
    pnTouch_Grupo: TPanel;
    scGrupo: TScrollBox;
    Panel1: TPanel;
    pnTouch_Produto: TPanel;
    scProduto: TScrollBox;
    Panel8: TPanel;
    Bevel3: TBevel;
    AdvMetroButton1: TAdvMetroButton;
    AdvMetroButton2: TAdvMetroButton;
    qrProdCombCODIGO: TIntegerField;
    qrProdCombPRODUTO: TStringField;
    qrProdCombUNIDADE: TStringField;
    qrProdCombQTDE: TFloatField;
    qrAltenaNumeroNFCeLocal: TUniQuery;
    cxStyle1: TcxStyle;
    cxImageList1: TcxImageList;
    mmECF: TMenuItem;
    L1: TMenuItem;
    C3: TMenuItem;
    tmCaixa: TJvThreadTimer;
    qrCaixa: TUniQuery;
    lbDestaque: TLabel;
    lbEstoque: TLabel;
    qrCaixaCAIXA_SITUACAO: TStringField;
    qrCaixaCAIXA_DATA_MOVTO: TDateField;
    qrCaixaFECHAMENTO_AUTOMATICO_CAIXA_DIA: TStringField;
    mniAlterarValorUnitario1: TMenuItem;
    mniAlterarQuantidadeItens: TMenuItem;
    mmiAplicarVarejo: TMenuItem;
    uQry: TUniQuery;
    uqryItens: TUniQuery;
    uqryItensVOLUME: TStringField;
    uqryItensCODIGO: TStringField;
    uqryItensCOD_CUPOM: TStringField;
    uqryItensITEM: TIntegerField;
    uqryItensCOD_PRODUTO: TIntegerField;
    uqryItensUNIDADE: TStringField;
    uqryItensQTDE: TFloatField;
    uqryItensVALOR_UNITARIO: TFloatField;
    uqryItensVALOR_SUBTOTAL: TFloatField;
    uqryItensVALOR_DESCONTO: TFloatField;
    uqryItensVALOR_ACRESCIMO: TFloatField;
    uqryItensVALOR_TOTAL: TFloatField;
    uqryItensCST: TStringField;
    uqryItensALIQUOTA: TFloatField;
    uqryItensCOD_TOTALIZADOR: TStringField;
    uqryItensCANCELADO: TIntegerField;
    uqryItensTAMANHO: TStringField;
    uqryItensCOR: TStringField;
    uqryItensEXTRA: TBlobField;
    uqryItensCOMPLEMENTO: TStringField;
    uqryItensAPLICADO_DESCONTO_GRUPO: TStringField;
    uqryItensQTD_APLICADO_DESCONTO: TFloatField;
    uqryItensDESCONTO_CUPOM: TFloatField;
    uqryItensCOD_VOLUME_PRODUTO: TIntegerField;
    uqryItensQTD_VOLUME: TFloatField;
    uqryItensPOR_VOLUME: TIntegerField;
    mniCancelarUltimaSangria: TMenuItem;
    mniLancarSangria: TMenuItem;
    mniRegistrarNovoSuprimento: TMenuItem;
    mniCancelarUltimoSuprimento: TMenuItem;
    wspnlAlerta: TW7Panel;
    JvPanel1: TJvPanel;
    lblInstrucao: TLabel;
    Image1: TImage;
    cpnlMensagem: TCurvyPanel;
    Label23: TLabel;
    cdsRecebimentoRec1: TStringField;
    cdsRecebimentoRec2: TStringField;
    procedure Cancela_cupom_aberto();
    procedure Cancela_Item(sItem: string; Acao: string);
    procedure Registra_Item();
    procedure Limpa_controles();
    function Abre_Venda(ReabrirCupom: Boolean = False): Boolean;
    function Truncar_Valor(valor: Real): Real;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure ed_barraKeyPress(Sender: TObject; var Key: Char);
    procedure CancelarItem1Click(Sender: TObject);
    procedure ed_barraEnter(Sender: TObject);
    procedure Retornar1Click(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure CancelarCupom1Click(Sender: TObject);
    procedure ed_barraExit(Sender: TObject);
    procedure ED_DESC_ACREEnter(Sender: TObject);
    procedure FecharCupom1Click(Sender: TObject);
    procedure ed_qtdeEnter(Sender: TObject);
    procedure ed_unitarioEnter(Sender: TObject);
    procedure AcionaBalana1Click(Sender: TObject);
    procedure ConsultadeClientes1Click(Sender: TObject);
    procedure ConsultadePreos1Click(Sender: TObject);
    procedure Opes1Click(Sender: TObject);
    procedure PesquisadeProdutos1Click(Sender: TObject);
    procedure Cupons1Click(Sender: TObject);
    procedure AcionaGaveta1Click(Sender: TObject);
    procedure FecharCaixa1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Sangria1Click(Sender: TObject);
    procedure Suprimento1Click(Sender: TObject);
    procedure rocarOperador1Click(Sender: TObject);
    procedure CaixaIndisponvel1Click(Sender: TObject);
    procedure menu_PrevendaClick(Sender: TObject);
    procedure Oramento1Click(Sender: TObject);
    procedure Recebimento1Click(Sender: TObject);
    procedure menu_mesaClick(Sender: TObject);
    procedure timer_cargaTimer(Sender: TObject);
    procedure ed_unitarioKeyPress(Sender: TObject; var Key: Char);
    procedure TimerTrocoTimer(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure sp_total_pagarContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ed_qtdeKeyPress(Sender: TObject; var Key: Char);
    procedure tmContigenciaTimer(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure AdvMetroButton2Click(Sender: TObject);
    procedure mnDeliveryClick(Sender: TObject);
    procedure G1Click(Sender: TObject);
    procedure frComprovanteDebitoBeforePrint(Sender: TfrxReportComponent);
    procedure RelCupomNFRestauranteBeforePrint(Sender: TfrxReportComponent);
    procedure scGrupoMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure scGrupoMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure btnDimQtdClick(Sender: TObject);
    procedure btnAumQtdClick(Sender: TObject);
    procedure ed_qtdeExit(Sender: TObject);
    procedure ed_unitarioExit(Sender: TObject);
    procedure ed_total_itemKeyPress(Sender: TObject; var Key: Char);
    procedure mnEnviaContigenciaClick(Sender: TObject);
    procedure cxGrid1DBBandedTableView1FocusedRecordChanged(Sender: TcxCustomGridTableView; APrevFocusedRecord, AFocusedRecord: TcxCustomGridRecord; ANewItemRecordFocusingChanged: Boolean);
    procedure colBotDiminuirPropertiesButtonClick(Sender: TObject; AButtonIndex: Integer);
    procedure colBotAlmuntarPropertiesButtonClick(Sender: TObject; AButtonIndex: Integer);
    procedure TimerExtraTimer(Sender: TObject);
    procedure MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure MouseLeave(Sender: TObject);
    procedure frxDesigner1GetTemplateList(List: TStrings);
    procedure RelFichaBeforePrint(Sender: TfrxReportComponent);
    procedure ACBrSAT1GetcodigoDeAtivacao(var Chave: AnsiString);
    procedure ACBrSAT1GetsignAC(var Chave: AnsiString);
    procedure mmSatClick(Sender: TObject);
    procedure DevoluoTroca1Click(Sender: TObject);
    procedure C2Click(Sender: TObject);
    procedure finalizarCupomNFCe;
    procedure mmECFClick(Sender: TObject);
    procedure C3Click(Sender: TObject);
    procedure tmCaixaTimer(Sender: TObject);
    procedure mniAlterarValorUnitario1Click(Sender: TObject);
    procedure mmiAplicarVarejoClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure mniCancelarSangriaOuSuprimentoClick(Sender: TObject);
    ////////////
    procedure AlertaFecharClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    FLastKeyPress: Word;
    FProduto: TProduto;
    p_borda_superior, p_borda_inferior, p_borda_esquerda,p_borda_direita,p_espacamento_superior,p_espacamento_central,p_espacamento_inferior,
    p_edit_quantidade, p_edit_valor_unitario, p_barra_superior,p_barra_inferior,p_quadro_esquerdo:Double;


    Controle: TControl;
    a, b: Word;
    Carregando: Boolean;
    iImpressora: Integer;
    CheckConnect: TThreadCheckConnection;
    Touch_Colunas, Touch_Linhas, Touch_linhas_Grupo, Touch_Altura_Grupo: Integer;
    bVerificaIniternet, bFezDevolucao: Boolean;
    procedure LancaMesaComanda(MesasComandas: string);
    function ImgTipoImpressora(i: Integer): TImpressora;
    procedure MostraFotoProduto(b: Boolean);
    procedure MostraLogoMarca(b: Boolean);
    procedure CentralizarPanel(p: TPanel);

    procedure GerarNFCe(NumNFe: string);
    procedure OnInternetChange(LANStatus: TLANStatus; InternetStatus: TInternetStatus; xMsg: string = '');
    procedure MontaTelaTouch;
    procedure MontaGrupos;
    procedure MontaProdutos(Grupo: Integer);
    procedure GrupoClick(Sender: TObject);
    procedure ProdutoClick(Sender: TObject);
    function RetornaSenhaImpressao: Integer;
    procedure LongPress;
    procedure AjustaDescontoAcrescimo(Desconto, Acrescimo: Double);
    procedure AlternarValorVarejo(ADirecao: TDirecao);
    procedure AplicaDescontoPorGrupo;
    procedure ConfiguraECF;
    procedure RegistrarProduto(ACodProduto: String = '');
    procedure HabilitarCancelamentoSangriaSuprimento(AAplicarTodos: Boolean = True);
    procedure ImprimirComprovanteCrediario;
    procedure ImprimeRecebimentosCartao;
  public
    { Public declarations }
    procedure RemoveDescontoTabelaItem;
    function PrepararNFCE: Boolean;
    procedure Imprime_display(texto: string; cor_da_fonte: TColor; Imagem: Tipo_Imagem);
    procedure Imprime_display_anterior;
    procedure CorEditTotaL;
    function Localizar_Produto(referencia: string): Boolean;
    procedure EnviaContingencia;
    procedure ImprimirCupomNaoFiscal(Cupom: string);
    procedure ImprimirComanda(Cupom: string);
    function RetornaPesoBalanca: Double;
    procedure ConfiguraSAT;
    function EnviaSAT(Cupom: string): TSat;
    procedure PrepararImpressao;
    function TemDevolucao: Boolean;
    var
      vr_saldo, vr_recebido, vr_total: string;
      IndiceTransacaoTef: Integer;
      strPizza: tretPizza;
  end;

const
  SEPARADOR = '-----------------------------------------' + #10;
  Estados: array [TACBrECFEstado] of string = ('ECF Não Inicializada', 'Desconhecido', 'ECF Livre', 'ECF em Venda', 'ECF em Pagamento', 'ECF Emitindo Relatório', 'ECF Bloqueada',
    'ECF Requer Redução Z', 'ECF Requer Leitura X', 'ECF Nao Fiscal');

var
  DataHoraContingencia: TDateTime;
  ContingenciaPorErroNoEnvio: Boolean;
  ErroNoEnvioContingencia: string;
  frmtOffLine: TpcnTipoEmissao;
  frmVenda: TfrmVenda;
  SatOk, FormatXML, CancelaIMpressaoCozinha: Boolean;
  CodCliGeraCredito, Sat_Largura, Sat_MargemTopo, Sat_MargemFundo, Sat_MargemEsquerda, Sat_MargemDireita: Integer;
  Sat_FormatXML, Sat_Preview: Boolean;
  Sat_CfopPadrao, Sat_Impressora: string;
  Sat_RedeTipoInter, Sat_RedeTipoLan, Sat_RedeSeg, Sat_RedeProxy, Sat_RedeProxyPorta: integer;
  Sat_RedeSSID, Sat_RedeCodigo, Sat_RedeIP, Sat_RedeMask, Sat_RedeGW, Sat_RedeDNS1, Sat_RedeDNS2, Sat_RedeUsuario, Sat_RedeSenha, Sat_RedeProxyIP, Sat_RedeProxyUser, Sat_RedeProxySenha: string;
  SwHAssinatura, CodigoAtivacaoSAT: string;
  RecebeCupomCredito: TRecebeCupomCredito;
  vExtra: tRetExtra;
  vRecebimento: TRecebimento;

 // resolução da tela
  OldWidth: Integer;
  SenhaImpressao, OldHeight: Integer;
  TipoImprimeCupom, ImprimeCupom: string;
  CartaoNaoInterfere, UtilizaSenha: Boolean;

  TotalRecebidoTef:Double;

  LancaRecebimento:TLacValorAutomatico;

  TipoEtiqueta:TTipoEtiqueta;

  QtdVolume: real;
  CtrlVolume: Boolean;
  ReaberturaCupomAtiva: String;
  VendaPorVolume: Integer;

  // totalizadoras
  rTotal_Venda: Real;
  rTotal_Desconto: Real;
  rTotal_Acrescimo: Real;
  rTotal_Couvert: Real;
  rTotal_Comissao: Real;
  rTotal_TaxaEntrega: Real;
  iTotal_Itens: Integer;
  iItem, CodProduto: Integer;
  // STATUS das funcoes
  bVenda: Boolean; // venda aberta
  bTotalizado: Boolean; // cupom totalizado
  bPago1, bPago2, bPago3: Boolean; // lancado cada forma de pagamento
  bFinalizado: Boolean; // Cupom Finalizado (Totalizado);
  // Outras
  blimpa_edit_display: Boolean;
  TefGP: integer;
  RecebeuTEF, TEF_Ativo: Boolean;
  // Armazenar dados anteriores do display
  sTexto_anterior: string;
  sMensagem_Cupom, sObsProduto: TMemo;
  clCor_anterior: TColor;
  TiImagem_anterior: Tipo_Imagem;

  // flag de finalizacao do tef
  bTef_finaliza: Boolean;
  // totalizadores das formas que exigem tratamento especial
  rvalor_total_cheque, rvalor_total_crediario, rvalor_total_convenio: Real;

  // variaveis do posto
  sPosto_Resposta, sPosto_Comando, str: string;
  iQtde_bico: Integer;
  bPosto_ok: Boolean;
  dtContingencia: TDate;
  hrContingencia: TTime;
  // focar no campo anterior
  bVolta_foco: Boolean;

  // Dados do Vendedor
  iVendedorCodigo: Integer;
  sVendedorNome: string;

    // variaveis nfce
  vAux, vNumLote, vSincrono, vcontingencia, vgerado_nfce: string;
  vNumNFCe: Integer;
  Sincrono: Boolean;
  bc: TBitmap;
  url, cDest, tpAmb: string;
  Hash_SHA1: string;
  X: Integer;
  S: AnsiString;
  tipo_pgto: Integer;
  nfce_autorizada: Boolean;
  bContinuaNFCE: Boolean;
  cStatus: Integer;
  ChaveNFCE: string;
  NumeroNFCe, SerieNFCe: Integer;
  Ambiente:string;

implementation

uses
  modulo, Math, funcoes, cliente_consulta, preco_consulta, produto_consulta,
  protetor_tela, venda_crediario, Identifica, caixa_fechamento,
  sangria, suprimento, senha, menu_cupom, pre_venda, msg_Operador,
  Orcamento_Abrir, IniFiles, contasreceber, caixa_abertura, mesas,
  senha_supervisor, ComObj, Vendedor, UFuncoes,
  MensagemAdicional, impDelivery, frmNFCEs, Vcl.Printers, FechamentoCego,
  ACBrSATClass, CuponsSAT, PegaCupomGerencial, RecebimentoEspecial,
  EnviaContingencia, Comanda, Config, Unit4, Unit5, mensagem, uTEF, menu_fiscal,
  principalChef, SolicitaAutorizacaoGerencial, uAlterarValor;

{$R *.dfm}

const 
  cRegistrado = -1;

var
  FrmPlanoFundo: TForm;

procedure ExibirPlanoFundo;
begin
  if FrmPlanoFundo = nil then
  begin
    FrmPlanoFundo := TForm.Create(nil);
    FrmPlanoFundo.BorderStyle := bsNone;
    FrmPlanoFundo.Color := clBlack;
    FrmPlanoFundo.AlphaBlend := True;
    FrmPlanoFundo.AlphaBlendValue := 200;
    FrmPlanoFundo.BoundsRect := Screen.WorkAreaRect;
    FrmPlanoFundo.Height := Screen.WorkAreaHeight - IfThen(Screen.Height - Screen.WorkAreaHeight = 0, 2, 0);
    FrmPlanoFundo.Show;
  end;
end;

procedure FecharPlanoFundo;
begin
  if FrmPlanoFundo <> nil then
  begin
    FrmPlanoFundo.Close;
    FrmPlanoFundo.Free;
  end;
  FrmPlanoFundo := nil;
end;

// -------------------------------------------------------------------------- //

function TrocaResolucao(X, Y, a, b: Word): Boolean;
var
  lpDevMode: TDeviceMode;
begin
  if EnumDisplaySettings(nil, 0, lpDevMode) then begin
    lpDevMode.dmFields := DM_PELSWIDTH or DM_PELSHEIGHT;
    lpDevMode.dmPelsWidth := X; // altera a resolução para a que voce voce desejou
    lpDevMode.dmPelsHeight := Y; // altera a resolução para a que voce voce desejou
    Result := ChangeDisplaySettings(lpDevMode, 0) = DISP_CHANGE_SUCCESSFUL;
  end;
end;

function VoltaResolucao(a, b: Word): Boolean;
var
  lpDevMode: TDeviceMode;
begin
  if EnumDisplaySettings(nil, 0, lpDevMode) then begin
    lpDevMode.dmFields := DM_PELSWIDTH or DM_PELSHEIGHT;
    lpDevMode.dmPelsWidth := a; // ao fechar o programa recupera a resolução que vinha sendo usada
    lpDevMode.dmPelsHeight := b; // ao fechar o programa recupera a resolução que vinha sendo usada
    Result := ChangeDisplaySettings(lpDevMode, 0) = DISP_CHANGE_SUCCESSFUL;
  end;
  Result := True;
end;

procedure TfrmVenda.OnInternetChange(LANStatus: TLANStatus; InternetStatus: TInternetStatus; xMsg: string = '');
var
  ok: Boolean;
begin

  if bVerificaIniternet then begin
    if InternetStatus = tiOnline then begin
      if ContigenciaPermanente then begin
        if frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao <> frmtOffLine then begin
          TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitGreen;
          cpBarra.Panels[9].Text := 'CONTINGÊNCIA';
          tmContigencia.Enabled := False;
          frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao := frmtOffLine;
        end
      end
      else begin
        tmContigencia.Enabled := False;
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitGreen;
        cpBarra.Panels[9].Text := 'ON-LINE';
        frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao := StrToTpEmis(ok, IntToStr(rgFormaEmissao + 1));
        if not (bVenda) then
          EnviaContingencia;
      end;
    end
    else begin
      if frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao <> frmtOffLine then begin
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;
        cpBarra.Panels[9].Text := 'OFF-LINE';
        tmContigencia.Enabled := True;
        frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao := frmtOffLine;
      end;
    end;
  end;
end;

procedure TfrmVenda.EnviaContingencia;
begin
  if Carregando then
    Exit;
  if TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType = sitGreen then begin
    qrContingencia.Close;
    qrContingencia.Open;
    if not qrContingencia.IsEmpty then begin
      try
        ed_barra.Enabled := False;
        while not qrContingencia.Eof do begin
          if qrContingenciaSTATUS.AsString <> 'E' then begin
            Imprime_display('   AGUARDE...  ENVIANDO CONTINGENCIAS', clGreen, tiLivre);
            with frmModulo do begin
              LerConfiguracao;
              if EdFormatoOff = 0 then
                frmtOffLine := teContingencia
              else
                frmtOffLine := teOffLine;
              ACBRNFCe.NotasFiscais.Clear;
              try
                ACBRNFCe.NotasFiscais.LoadFromFile(qrContingenciaXML.AsString);
                ACBRNFCe.Enviar(ACBRNFCe.NotasFiscais.Items[0].NFe.Ide.nNF, False, True);
                if ACBRNFCe.NotasFiscais.Items[0].Confirmada then begin
                  if ACBRNFCe.WebServices.Enviar.cStat = 100 then begin
                    qrContingencia.Edit;
                    qrContingenciaENVIADOCONTINGENCIA.AsString := 'S';
                    qrContingenciaEX.AsInteger := 0;
                    qrContingenciaSTATUS.AsString := 'V';
                    qrContingenciaALTERADO.AsString := 'S';
                    qrContingenciaXMLENVIO.LoadFromFile(qrContingenciaXML.AsString);
                    qrContingencia.Post;
                    qrContingencia.ApplyUpdates;
                  end;
                end
                else begin
                  qrContingencia.Edit;
                  qrContingenciaSTATUS.AsString := 'E';
                  qrContingenciaERRO.AsString := ACBRNFCe.WebServices.Enviar.xMotivo;
                  qrContingencia.Post;
                  qrContingencia.ApplyUpdates;
                end;
              except
                on e: exception do begin
                  qrContingencia.Edit;
                  qrContingenciaSTATUS.AsString := 'E';
                  qrContingenciaERRO.AsString := e.Message;
                  qrContingencia.Post;
                  qrContingencia.ApplyUpdates;
                end;
              end;
            end;
          end;
          qrContingencia.Next;
        end;
      finally
        ed_barra.Enabled := True;
        if not AtivaTouch then
          ed_barra.SetFocus;
        Imprime_display('C A I X A    L I V R E', clWhite, tiLivre);
      end;
    end;
  end;
end;

function TfrmVenda.EnviaSAT(Cupom: string): TSat;
var
  TotalItem: Double;
  tGeralCup: Currency;
  ctItem: Integer;
  ST: string;
  cfOP: string;
  Origem: string;
  noPro: string;
  Ncm: string;
  TributosEST: Currency;
  TributosFED: Currency;
  itTributosEST: Currency;
  itTributosFED: Currency;
  cMsg: string;
  UMprod: string;
  edtPathCFe: string;
  ValorAcrescimo, ValorDesconto: Real;
   //Variaveis para checar o CPF e o CNPJ
  Tm: array[1..14] of integer;
  Tn: array[1..14] of integer;
  Dgt1, Dgt2: integer;
  Soma1, Soma2: integer;
begin
  try
    if not frmModulo.qrfilial.Active then
      frmModulo.qrfilial.Open;
    ACBrSAT1.CFe.IdentarXML := sat_FormatXML;
    ACBrSAT1.CFe.TamanhoIdentacao := 3;
    ConfiguraSAT;
    ACBrSAT1.InicializaCFe;
    ValorAcrescimo := 0;
    ValorDesconto := 0;
    with ACBrSAT1.CFe, frmModulo do begin

      ide.numeroCaixa := iNumCaixa;
      Dest.CNPJCPF := OnlyNumber(sConsumidor_CPF);
      Dest.xNome := sConsumidor_Nome;
      ctItem := 1;
      tGeralCup := 0;
      qrItens.First;
      while not qrItens.Eof do begin
        if qrItensCANCELADO.AsInteger = 0 then begin
          QrProdutoSAT.Close;
          QrProdutoSAT.SQL.Clear;
          QrProdutoSAT.SQL.Add('select * from estoque where codigo = ' + QuotedStr(qrItensCOD_PRODUTO.AsString));
          QrProdutoSAT.Open;

          QRCSOSN.close;
          QRCSOSN.sql.clear;
          QRCSOSN.sql.add('select pro.csosn, pro.origem, pro.aliquota icms, pro.cest, pro.codigo_servico servico ');
          QRCSOSN.sql.Add('from estoque pro');
          QRCSOSN.sql.add('where pro.codigo = :pcodigo');
          QRCSOSN.ParamByName('pcodigo').asstring := Zerar(qrItensCOD_PRODUTO.AsString, 6);
          QRCSOSN.Open;
          with Det.Add do begin
            noPro := QrProdutoSAT.FieldByname('NOME').AsString;
            UMprod := QrProdutoSAT.FieldByname('Unidade').AsString;
            Ncm := QrProdutoSAT.FieldByname('NCM').AsString;
            Origem := QRCSOSN.FieldByname('ORIGEM').AsString;
            if not (QrProdutoSAT.FieldByName('cfop').IsNull) and (Trim(QrProdutoSAT.FieldByName('cfop').AsString) <> '') then
              cfOP := QrProdutoSAT.FieldByName('cfop').AsString
            else
              cfOP := Sat_CfopPadrao;
            nItem := ctItem;
            // JONATHAN - 170/02/2020 - Codigo interno para código de barras
            // Prod.cProd := QrProdutoSAT.fieldbyname('codigo').AsString;

            Prod.cProd := QrProdutoSAT.fieldbyname('cod_barra').AsString;
            Prod.cEAN := ' ';
            Prod.xProd := noPro;
            prod.NCM := Ncm;
            Prod.uCom := UMprod;
            Prod.qCom := qrItensqtde.AsFloat;
            Prod.vUnCom := qrItensvalor_Unitario.AsFloat;
            Prod.indRegra := irArredondamento;
            Prod.vDesc := 0;
            Prod.CFOP := cfOP;
            Prod.vRatAcr := qrItensVALOR_ACRESCIMO.asfloat;
            Prod.vRatDesc := qrItensVALOR_DESCONTO.asfloat;
            ValorAcrescimo := ValorAcrescimo + Prod.vRatAcr;
            ValorDesconto := ValorDesconto + Prod.vRatDesc;

            if trim(QRCSOSN.FieldByName('servico').AsString) <> '' then begin
              TotalItem := Prod.qCom * Prod.vUnCom;
              qrServidor.Close;
              qrServidor.SQL.Clear;
              qrServidor.SQL.Add('select emp.iss, emp.cod_municipio_ibge from empresa emp');
              qrServidor.Open;
              Imposto.ICMS.CSOSN := csosn900;
              Imposto.ISSQN.cNatOp := 1;
              Imposto.ISSQN.vBC := Prod.vProd;
              Imposto.ISSQN.vAliq := qrServidor.FieldByName('iss').AsFloat;
              Imposto.ISSQN.vISSQN := RoundTo((Imposto.ISSQN.vBC * Imposto.ISSQN.vAliq) / 100, -2);
              Imposto.ISSQN.cMunFG := qrServidor.FieldByName('cod_municipio_ibge').AsInteger;
              Imposto.ISSQN.cListServ := trim(QRCSOSN.FieldByName('servico').AsString);
              Imposto.ISSQN.indIncFisc := iiSim;
              Imposto.PIS.CST := pis07;
              Imposto.COFINS.CST := cof07;
              with Total.ISSQNtot do begin
                vBC := Imposto.ISSQN.vBC;
                vISS := Imposto.ISSQN.vISSQN;
              end;
            end
            else begin
              frmModulo.qrfilial.First;
              if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'S' then begin
                if not QrProdutoSAT.IsEmpty then
                  ST := QrProdutoSAT.FieldByName('CSOSN').AsString;
              end
              else
                ST := qrItensCST.AsString;
              TotalItem := Prod.qCom * Prod.vUnCom;
              tGeralCup := tGeralCup + TotalItem;
              if Ncm <> '' then begin
                if (QrProdutoSAT.RecordCount > 0) and (QrProdutoSAT.FieldByName('ALIQNACIONAL').AsFloat > 0) then begin
                  TributosEST := TributosEST + (TotalItem * QrProdutoSAT.FieldByName('ALIQNACIONAL').AsFloat / 100);
                  itTributosEST := (TotalItem * QrProdutoSAT.FieldByName('ALIQNACIONAL').AsFloat / 100);
                end;
              end;
              if Origem = '0' then
                Imposto.vItem12741 := itTributosEST
              else
                Imposto.vItem12741 := itTributosEST + itTributosFED;
              if Origem = '0' then
                Imposto.ICMS.orig := oeNacional;
              if Origem = '1' then
                Imposto.ICMS.orig := oeEstrangeiraImportacaoDireta;
              if Origem = '2' then
                Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasil;
              if Origem = '3' then
                Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior40;
              if Origem = '4' then
                Imposto.ICMS.orig := oeNacionalProcessosBasicos;
              if Origem = '5' then
                Imposto.ICMS.orig := oeNacionalConteudoImportacaoInferiorIgual40;
              if Origem = '6' then
                Imposto.ICMS.orig := oeEstrangeiraImportacaoDiretaSemSimilar;
              if Origem = '7' then
                Imposto.ICMS.orig := oeEstrangeiraAdquiridaBrasilSemSimilar;
              if Origem = '8' then
                Imposto.ICMS.orig := oeNacionalConteudoImportacaoSuperior70;
              //CST ou CSOSN
              if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'N' then begin //Regime normal
                if ST = '000' then begin
                  Imposto.ICMS.CST := cst00;
                  Imposto.ICMS.pICMS := 18;
                end;
                if ST = '020' then
                  Imposto.ICMS.CST := cst20;
                if ST = '090' then begin
                  Imposto.ICMS.CST := cst90;
                  Imposto.ICMS.pICMS := 18;
                end;
                if ST = '040' then
                  Imposto.ICMS.CST := cst40;
                if ST = '041' then
                  Imposto.ICMS.CST := cst41;
                if ST = '050' then
                  Imposto.ICMS.CST := cst50;
                if ST = '060' then
                  Imposto.ICMS.CST := cst60;
              end
              else if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'S' then begin //Regime simples nacional
                if ST = '101' then
                  Imposto.ICMS.CSOSN := csosn101;
                if ST = '102' then
                  Imposto.ICMS.CSOSN := csosn102;
                if ST = '103' then
                  Imposto.ICMS.CSOSN := csosn103;
                if ST = '201' then
                  Imposto.ICMS.CSOSN := csosn201;
                if ST = '202' then
                  Imposto.ICMS.CSOSN := csosn202;
                if ST = '203' then
                  Imposto.ICMS.CSOSN := csosn203;
                if ST = '300' then
                  Imposto.ICMS.CSOSN := csosn300;
                if ST = '400' then
                  Imposto.ICMS.CSOSN := csosn400;
                if ST = '500' then
                  Imposto.ICMS.CSOSN := csosn500;
                if ST = '900' then
                  Imposto.ICMS.CSOSN := csosn900;
                Imposto.ICMS.pICMS := qrItensaliquota.AsFloat;
              end;
              //PIS
              if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'N' then
                Imposto.PIS.CST := pis99
              else if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'S' then
                Imposto.PIS.CST := pis49;
              //Imposto.PIS.vBC := TotalItem;
              //Imposto.PIS.pPIS := 0.0065;
              //COFINS
              if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'N' then
                Imposto.COFINS.CST := cof99
              else if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'S' then
                Imposto.COFINS.CST := cof49;
            end;
          end;
          ctItem := ctItem + 1;
        end;
        qrItens.Next;
      end;
      Total.DescAcrEntr.vDescSubtot := ValorDesconto;
      Total.DescAcrEntr.vAcresSubtot := ValorAcrescimo;

      Total.vCFeLei12741 := TributosEST + TributosFED;
      if vRecebimento.Troco > 0 then begin
        Pagto.vTroco := vRecebimento.Troco;
      end;
      cdsRecebimento.First;
      while not cdsRecebimento.Eof do begin
        with Pagto.Add do begin
          vMP := cdsRecebimentovalor.AsFloat;
          if cdsRecebimentoTipo.AsString = StRecDinheiro then
            cMP := mpDinheiro;
          if (cdsRecebimentoTipo.AsString = StRecCheque) or (cdsRecebimentoTipo.AsString = StRecCheque) then
            cMP := mpCheque;
          if (cdsRecebimentoTipo.AsString = StRecCartaoCR) then begin
            cMP := mpCartaodeCredito;
            if cdsRecebimentocod_adm_sat.AsString <> '' then
              cAdmC := cdsRecebimentocod_adm_sat.AsInteger;
          end;
          if (cdsRecebimentoTipo.AsString = StRecCartaoDE) then begin
            cMP := mpCartaodeDebito;
            if cdsRecebimentocod_adm_sat.AsString <> '' then
              cAdmC := cdsRecebimentocod_adm_sat.AsInteger;
          end;
          if (cdsRecebimentoTipo.AsString = StRecCrediario) or (cdsRecebimentoTipo.AsString = StRecConvenio) then
            cMP := mpCreditoLoja;
        end;
        cdsRecebimento.Next;
      end;
      cMsg := 'Obrigado Volte Sempre!!!;';
      if frmModulo.qrfilial.FieldByName('optante_simples').AsString = 'S' then
        cMsg := cMsg + 'Optante pelo simples nacional lei 123/2006;';
      cMsg := cMsg + 'Número do Cupom Intérno: ' + sNumero_Venda;
      InfAdic.infCpl := cMsg;
      ACBrSAT1.EnviarDadosVenda; //(mSalvaXML.Lines.Text);
      if ACBrSAT1.Resposta.codigoDeRetorno = 6000 then begin
        PrepararImpressao;
        ACBrSAT1.ImprimirExtrato;
        if (TEF_Ativo) and (RecebeuTEF) then
          finalizarCupomNFCe;
        Query1.SQL.Clear;
        Query1.SQL.Add('insert into sat (data, hora, numero, codcliente, ambiente, cancelado, chave, xml_envio, xml_cancelamento, alterado, cupom, total)');
        Query1.SQL.Add('values');
        Query1.SQL.Add('(:data, :hora, :numero, :codcliente, :ambiente, :cancelado, :chave, :xml_envio, :xml_cancelamento, :alterado, :cupom, :total)');
        Query1.ParamByName('data').AsDate := ACBrSAT1.CFe.ide.dEmi;
        Query1.ParamByName('hora').AsString := FormatDateTime('HH:MM:SS', ACBrSAT1.CFe.ide.hEmi);
        Query1.ParamByName('numero').AsInteger := ACBrSAT1.CFe.ide.nCFe;
        Query1.ParamByName('codcliente').AsString := sCli_codigo;
        Query1.ParamByName('ambiente').AsInteger := StrToInt(TpAmbToStr(ACBrSAT1.CFe.ide.tpAmb));
        Query1.ParamByName('chave').AsString := ACBrSAT1.CFe.infCFe.ID;
      //  Query1.ParamByName('xml_envio').AsString := ACBrSAT1.CFe.AsXMLString;
        Query1.ParamByName('xml_cancelamento').Clear;
        Query1.ParamByName('cancelado').AsString := 'N';
        Query1.ParamByName('alterado').AsString := 'N';
        Query1.ParamByName('cupom').AsString := sNumero_Venda;
        Query1.ParamByName('total').AsFloat := rTotal_Venda;
        Query1.ExecSQL;
        Result.Validou := True;
        Result.Numero := ACBrSAT1.CFe.ide.nCFe;
        Result.Codigo := ACBrSAT1.Resposta.codigoDeRetorno;
        Result.Mesagem := ACBrSAT1.Resposta.mensagemRetorno;
      end
      else begin
        Result.Validou := False;
        Result.Numero := 0;
        Result.Codigo := ACBrSAT1.Resposta.codigoDeRetorno;
        Result.Mesagem := ACBrSAT1.Resposta.mensagemRetorno;
      end;
    end;
  except
    on e: Exception do begin
      Result.Validou := False;
      Result.Numero := 0;
      Result.Codigo := 0;
      Result.Mesagem := e.Message;
      Application.MessageBox(PWideChar('Ocorre um erro ao tentar enviar o SAT:' + #13 + #13 + e.Message), 'Atenção!', MB_ICONINFORMATION);
    end;
  end;
end;

// -------------------------------------------------------------------------- //
// Funcao de busca de produtos

function TfrmVenda.Localizar_Produto(referencia: string): Boolean;
var
  bBusca_codigo, lNovoRegistro: Boolean; // variavel de para ativar ou nao a busca pelo campo codigo
  lPosicao, lMarcador, lCodProduto: Integer;
  lQtdVendido, lProd_qtde: Real;
  procedure AtualizaRegistro(ACodigo: Integer; AQtd: Real; AVendaPorVolume: Integer; ANovoCodigo: Integer = 0);
  begin
    CodProduto := ACodigo;
    rProd_qtde := AQtd;
    VendaPorVolume := AVendaPorVolume;
    RegistrarProduto(IfThen(ANovoCodigo = 0, ACodigo, ANovoCodigo).ToString);
  end;
  procedure IncluirRegistro(ACodigo: Integer; AQtd: Real);
  begin
    CodProduto := 0;
    rProd_qtde := AQtd;
    RegistrarProduto(ACodigo.ToString);
  end;
  begin
  try
    // limpar a variaveis de retorno
    sProd_nome := '';
    sProd_barra := '';
    sProd_unidade := '';
    sProd_CST := '';
    iProd_codigo := -1;
    rProd_aliquota := -1;
    rProd_Estoque := 0;
    rProd_preco := 0;
    if not CtrlVolume then
      rProd_qtde := 1;
    rProd_total := 0;
    sProd_Tamanho := '';
    sProd_Cor := '';

    // verificar se foi digitado a quantidade no campo de busca do produto utilizando os
    // caracteres "X" ou "*"

    if (Pos('X', AnsiUpperCase(referencia)) > 0) or (Pos('*', AnsiUpperCase(referencia)) > 0) then
    begin
      lPosicao := Pos('X', AnsiUpperCase(referencia));
      if lPosicao = 0 then
        lPosicao := Pos('*', AnsiUpperCase(referencia));
      try
        // alimenta a varivel de qtde
        rProd_qtde := StrToFloat(Copy(referencia, 1, lPosicao - 1));
        // copia apenas a parte que contem a informacao da busca do produto
        referencia := Copy(referencia, lPosicao + 1, 99);
      except
        // retornou erro, zera a qtde e deixa a referencia como estah.
        rProd_qtde := 1;
      end;
    end;

    // verificar ha possibilidade do texto ser um numero inteiro ou
    // nao ser um produto pesado em balanca --> iniciado por 2
    if Length(referencia) < 8 then
    begin
      // verificar se o campo referencia eh um numero inteiro valido
      try
        StrToInt(referencia);
        bBusca_codigo := True;
      except
        on EConvertError do
          bBusca_codigo := False;
      end;
      if bBusca_codigo then
      begin
        // comando SQL para filtrar o produto pelo codigo de barras
        query.close;
        query.sql.clear;
        query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
        query.sql.add('where e.cod_barra = ''' + referencia + '''');
        query.Open;
        if query.recordcount > 0 then
        begin
          // econtrou o registro pelo codigo de barras
          result := True;
        end
        else
        begin
          // nao achou pelo codigo de barras, procurar pelo codigo
          try
            query.close;
            query.sql.clear;
            query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
            query.sql.add('where e.codigo = ' + referencia);
            query.open;
            if query.RecordCount > 0 then
            begin
              // encontrou o registro pelo codigo
              result := True;
            end
            else
            begin
              // nao encontrou nenhum produto
              result := False;
            end;
          except
            result := False;
          end;
        end;
      end
      else
      begin
        // BUSCAR PELO CODIGO DE BARRAS
        query.close;
        query.sql.clear;
        query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
        query.sql.add('where e.cod_barra = ''' + referencia + '''');
        query.Open;

        if query.recordcount > 0 then
        begin
          if query.recordcount > 1 then
          begin
            iProd_codigo := 0;
            // encontrou varios registros
            if iPesquisa_produto = 1 then
            begin
              frmProduto_consulta := TfrmProduto_consulta.create(Self);
              frmProduto_consulta.ed_barra.text := ed_barra.text;
              frmProduto_consulta.showmodal;
              if iProd_codigo > 0 then
              begin
                query.Close;
                query.sql.Clear;
                query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
                query.sql.add('where e.codigo = ' + IntToStr(iProd_codigo));
                query.open;
                result := True;
              end
              else
                result := False;
            end
            else
              result := True;
          end
          else
            // encontrou o registro pelo nome
            result := True;
        end
        else
        begin
          // nao achou pelo codigo de barras, procurar pelo nome
          query.close;
          query.sql.clear;
          query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
          query.sql.add('where upper(e.REFERENCIA) LIKE ''' + AnsiUpperCase(referencia) + '%''');
          query.sql.add('order by e.REFERENCIA, e.COR, e.TAMANHO');
          query.open;

          if query.RecordCount > 0 then
          begin
            Result := True;
          end
          else
          begin
            // nao achou pelo codigo de barras, procurar pelo nome
            query.close;
            query.sql.clear;
            query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
            query.sql.add('where upper(e.nome) like ''' + AnsiUpperCase(referencia) + '%''');
            query.sql.add('order by e.nome');
            query.open;

            if query.RecordCount > 0 then
            begin
              if query.recordcount > 1 then
              begin
                iProd_codigo := 0;
                // encontrou varios registros
                if iPesquisa_produto = 1 then
                begin
                  frmProduto_consulta := TfrmProduto_consulta.create(Self);
                  frmProduto_consulta.ed_barra.text := ed_barra.text;
                  frmProduto_consulta.showmodal;
                  if iProd_codigo > 0 then
                  begin
                    query.Close;
                    query.sql.Clear;
                    query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
                    query.sql.add('where e.codigo = ' + IntToStr(iProd_codigo));
                    query.open;
                    result := True;
                  end
                  else
                    result := False;
                end
                else
                  result := True;
              end
              else
                // encontrou o registro pelo nome
                result := True;
            end
            else
            begin
              // nao encontrou nenhum produto
              result := False;
            end;
          end;
        end;
      end;
    end
    else
    begin
      // verificar se o produto é pesado em balanca
      if (Copy(referencia, 1, 1) = '2') and (Length(referencia) <= 13) then
      begin
        try //Exemplo de codigo na etiqueta: 2000010002550 -> Codigo = 000010 Valor = 00255 = 2,55
          query.close;
          query.sql.clear;
          query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
          query.sql.add('where e.codigo = ''' + Zerar(Copy(referencia, 2, iTamanho_codigo_balanca), 6) + '''');
          query.open;
          if query.RecordCount > 0 then
          begin
            if TipoEtiqueta = teTotal then
              rProd_total := StrToFloat(Copy(referencia, 8, 5)) / 100
            else
              rProd_qtde := StrToFloat(Copy(referencia, 8, 5)) / 100;
            result := True;
          end
          else
            result := False;
        except
          result := False;
        end;
      end
      else
      if (Copy(referencia, 1, 1) = 'P') and (Length(referencia) = 14) then
      begin
        query.close;
        query.sql.clear;
        query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
        query.sql.add('where e.cod_barra = ''' + Copy(referencia, 2, 13) + '''');
        query.Open;
        if query.recordcount > 0 then
        begin
          result := True;
          bMudouProvisorio := True;
          bMuda_unitario := True;
        end
        else
          result := False;
      end
      else
      begin
        // nao achou pelo codigo de barras, procurar pelo nome
        query.close;
        query.sql.clear;
        query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
        query.sql.add('where upper(e.REFERENCIA) LIKE ''' + AnsiUpperCase(referencia) + '%''');
        query.sql.add('order by e.REFERENCIA, e.COR, e.TAMANHO');
        query.open;

        if query.RecordCount > 0 then
        begin
          // nao encontrou nenhum produto
          result := True;
        end
        else
        begin
          // busca normal pelo codigo de barras e pelo nome do produto
          // comando SQL para filtrar o produto pelo codigo de barras
          query.close;
          query.sql.clear;
          query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
          query.sql.add('where e.cod_barra = ''' + referencia + '''');
          query.Open;

          if query.recordcount > 0 then
          begin
            if query.recordcount > 1 then
            begin
              iProd_codigo := 0;
              if iPesquisa_produto = 1 then
              begin
                frmProduto_consulta := TfrmProduto_consulta.create(Self);
                frmProduto_consulta.ed_barra.text := ed_barra.text;
                frmProduto_consulta.showmodal;
                if iProd_codigo > 0 then
                begin
                  query.Close;
                  query.sql.Clear;
                  query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
                  query.sql.add('where e.codigo = ' + IntToStr(iProd_codigo));
                  query.open;
                  result := True;
                end
                else
                  result := False;
              end
              else
                result := True;
            end
            else
              result := True;
          end
          else
          begin
            // nao achou pelo codigo de barras, procurar pelo nome
            query.close;
            query.sql.clear;
            query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
            query.sql.add('where upper(e.nome) like ''' + referencia + '%''');
            query.sql.add('order by e.nome');
            query.open;
            if query.recordcount = 0 then
              Result := False
            else if query.recordcount > 1 then
            begin
              iProd_codigo := 0;
              if iPesquisa_produto = 1 then
              begin
                frmProduto_consulta := TfrmProduto_consulta.create(Self);
                frmProduto_consulta.ed_barra.text := ed_barra.text;
                frmProduto_consulta.showmodal;
                if iProd_codigo > 0 then
                begin
                  query.Close;
                  query.sql.Clear;
                  query.sql.add('SELECT (SELECT COUNT(SSE.CODIGO) QTDE FROM ESTOQUE SSE WHERE SSE.CODPRODUTO = E.CODIGO) QTD_VOLUME, E.*  FROM ESTOQUE E');
                  query.sql.add('where codigo = ' + IntToStr(iProd_codigo));
                  query.open;
                  result := True;
                end
                else
                  result := False;
              end
              else
                result := True;
            end
            else
              result := True;
          end;
        end;
      end;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if result then
    begin
      // encontrou o produto, alimentar as variaveis de retorno
      if not CtrlVolume and
         (query.fieldbyname('usa_balanca').AsInteger <> 1) and
         (query.fieldbyname('pizza').AsString <> 'S') and
         (query.fieldbyname('QTD_VOLUME').AsInteger > 0)
      then
      begin
        lProd_qtde := rProd_qtde;
        /////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// Organiza itens na lista incluindo sempre no mesmo ja existente
        uqryItens.Close;
        uqryItens.ParamByName('COD_PRODUTO').AsInteger := query.fieldbyname('codigo').AsInteger;
        uqryItens.Open;
        ///
        uQry.Close;
        uQry.sql.Clear;
        uQry.sql.Add('SELECT * FROM ESTOQUE WHERE CODPRODUTO = ' + QuotedStr(zerar(query.fieldbyname('codigo').asstring, 6)));
        uQry.sql.add('order by QTD_VOLUME');
        uQry.Open;
        ///
        lNovoRegistro := uqryItens.RecordCount = 0;
        if not lNovoRegistro or ((uQry.RecordCount > 0) and (rProd_qtde > 1)) then
        begin
          CodProduto := query.fieldbyname('codigo').AsInteger;
          lProd_qtde := lProd_qtde + uqryItensQTDE.Value;
          QtdVolume := uqryItensQTD_VOLUME.Value;
          if QtdVolume = 0 then
            QtdVolume := 1;
          /////////////////////////////////////////////////////////////////////////////////////////////////////////
          ///  Verifica quantidades transformando em volume composto
          if uqryItensVOLUME.Value <> 'S' then
          begin
            VendaPorVolume := 0;
            with FProduto do
            begin
              if uQry.RecordCount > 0 then
              begin
                SetLength(Volume, 0);
                lMarcador := uQry.RecordCount;
                SetLength(Volume, Succ(lMarcador));
                Codigo := query.fieldbyname('codigo').AsInteger;
                QtdVenda := lProd_qtde;
                //////////////////////////////////////////////////////////////////
                ///  Filtra  todos os volumes relacionados ao produto apontado
                ///  e totaliza quantos foram lancados por unidades
                uQry.First;
                while not uQry.EOF do
                begin
                  lCodProduto := uQry.fieldbyname('CODIGO').AsInteger;
                  ///
                  uqryItens.Close;
                  uqryItens.ParamByName('COD_PRODUTO').AsInteger := lCodProduto;
                  uqryItens.Open;
                  with Volume[uQry.RecNo], Quantidade do
                  begin
                    ID := lCodProduto;
                    ItensVolume := uQry.fieldbyname('QTD_VOLUME').AsInteger;
                    Unidade := uQry.fieldbyname('UNIDADE').AsString;
                    PrecoVenda := uQry.fieldbyname('PRECO_VENDA').AsInteger;
                    ///
                    if uqryItens.RecordCount > 0 then
                    begin
                      Item := uqryItensITEM.Value;
                      Vendido := uqryItensQTDE.Value;
                      PorVolume := uqryItensPOR_VOLUME.Value;
                      QtdVenda := QtdVenda + (ItensVolume * (Vendido - PorVolume));
                    end
                    else
                    begin
                      Item := 0;
                      Vendido := 0;
                      PorVolume := 0;
                    end;
                  end;
                  uQry.Next;
                end;
                /////////////////////////////////////////////////////////////////////
                ///  Reagrupa o total unitário por volumes e lança nova atualização
                ///  ou inclusão do registro de venda
                while lMarcador > 0 do
                  with Volume[lMarcador], Quantidade do
                  begin
                    if QtdVenda >= ItensVolume then
                    begin
                      lQtdVendido := Trunc(QtdVenda / ItensVolume);
                      QtdVenda := QtdVenda - (lQtdVendido * ItensVolume);
                      Vendido := PorVolume + lQtdVendido;
                    end
                    else
                      Vendido := PorVolume;
                    Dec(lMarcador)
                  end;
                /////////////////////////////////////////////////////////////////////
                ///  Lança atualização de venda ou inclui novo registro
                if (QtdVenda > 0) and not lNovoRegistro then
                  AtualizaRegistro(Codigo, QtdVenda, 0)
                else
                begin
                  // prepara para exclusão se necessário na composição de volumes
                  uQry.Close;
                  uQry.sql.Clear;
                  uQry.sql.Add('SELECT * FROM CUPOM_ITEM_TEMP WHERE COD_PRODUTO = :COD_PRODUTO');
                  uQry.Open;

                  // primeiro verificar se existe algum volume com item e sem qtd para ser usado por outro com qtd sem item
                  lPosicao := 1;
                  for lMarcador := High(Volume) downto 1 do
                  begin
                    if lPosicao < 1 then
                      Break;
                    ///
                    with Volume[lMarcador], Quantidade do
                    if (Vendido = 0) and (Item > 0) then
                      for lPosicao := High(Volume) downto 1 do
                        with Volume[lPosicao], Quantidade do
                        if (lMarcador <> lPosicao) and (Vendido > 0) and (Item = 0) then
                        begin
                          Item := Volume[lMarcador].Item;
                          Volume[lMarcador].Item := cRegistrado;
                          AtualizaRegistro(Volume[lMarcador].ID, Vendido, PorVolume, ID);
                          Break;
                        end;
                  end;
                  ///
                  if lNovoRegistro then
                  begin
                     // senao verifico se existe algum sem item e com volume e incluo
                    /// INCLUSAO
                    for lMarcador := High(Volume) downto 1 do
                      with Volume[lMarcador], Quantidade do
                        if (Vendido > 0) and (Item = 0) then
                          IncluirRegistro(ID, Vendido);
                    // não havendo novos registros de volumes incluo os itens
                    if lMarcador < 1 then
                      IncluirRegistro(Codigo, QtdVenda);
                  end
                  else
                  begin
                    // senao verifico se existe algum sem item e com volume para usar no registro atual que foi excluido
                    /// ALTERACAO
                    for lMarcador := 1 to High(Volume) do
                      with Volume[lMarcador], Quantidade do
                        if (Vendido > 0) and (Item = 0) then
                        begin
                          Item := cRegistrado;
                          AtualizaRegistro(Codigo, Vendido, PorVolume, ID);
                          Break;
                        end;
                    // não havendo novos registros de volumes excluo os itens
                    if lMarcador > High(Volume) then
                    begin
                      uQry.Close;
                      uQry.ParamByName('COD_PRODUTO').AsInteger := Codigo;
                      uQry.Open;
                      uQry.Delete;
                    end;
                  end;
                end;
                ///
                if not lNovoRegistro then
                begin
                  lMarcador := 1;
                  while lMarcador < Length(Volume) do
                    with Volume[lMarcador], Quantidade do
                    begin
                      if Item > cRegistrado then
                      begin
                        if Item > 0 then
                        begin
                          //Registrado
                          if Vendido > 0 then
                            AtualizaRegistro(ID, Vendido, PorVolume)
                          else
                          begin
                            // Se tem registro posterior com venda sem item e uso a
                            // posicao atual para nao criar nova
                            lCodProduto := ID;
                            for lPosicao := Succ(lMarcador) to High(Volume) do
                              with Volume[lPosicao], Quantidade do
                              if (Item = 0) and (Vendido > 0) then
                              begin
                                Item := cRegistrado;
                                AtualizaRegistro(lCodProduto, Vendido, PorVolume);
                                Break;
                              end;
                            // senao excluo
                            if lPosicao > High(Volume) then
                            begin
                              uQry.Close;
                              uQry.ParamByName('COD_PRODUTO').AsInteger := ID;
                              uQry.Open;
                              uQry.Delete;
                            end;
                          end;
                        end
                        else
                        begin
                          //nao Registrado
                          if Vendido > 0 then
                          begin
                            // verifico se tem registro posterior com venda sem item e uso a posicao
                            lQtdVendido := Vendido;
                            for lPosicao := Succ(lMarcador) to High(Volume) do
                              with Volume[lPosicao], Quantidade do
                              if (Item > 0) and (Vendido = 0) then
                              begin

                                Item := cRegistrado;
                                AtualizaRegistro(ID, lQtdVendido, PorVolume);
                                Break;
                              end;
                            // senao Incluo novo
                            if lPosicao > High(Volume) then
                              IncluirRegistro(ID, Vendido);
                          end;
                        end;
                      end;
                      Inc(lMarcador)
                    end;
                end;
                Result := False;
                Exit;
              end;
            end;
          end
          else
            VendaPorVolume := uqryItensPOR_VOLUME.Value + Trunc(rProd_qtde);
          ///
          rProd_qtde := lProd_qtde;
        end
        else
        begin
          //////////////////////////////////////////////////////////
          ///  Verifica se o novo item a incluir é um volume então
          ///  define contador de Venda por Volume
          uQry.Close;
          uQry.sql.Clear;
          uQry.sql.Add('SELECT CODIGO FROM ESTOQUE WHERE CODIGO = ' + query.fieldbyname('codigo').asstring + ' AND CODPRODUTO <> ''0''');
          uQry.Open;
          VendaPorVolume := Integer(uQry.RecordCount > 0);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////////////////////////////////////
      /////////////////////////////////////////////////////////////////////////////////////////////////////////

      sProd_nome := query.fieldbyname('nome').asstring;
      sProd_barra := query.fieldbyname('cod_barra').asstring;
      sProd_unidade := query.fieldbyname('unidade').asstring;
      iProd_codigo := query.fieldbyname('codigo').asinteger;
      sProd_CST := query.fieldbyname('cst').asstring;
      rProd_aliquota := query.fieldbyname('aliquota').asfloat;
      rProd_Estoque := query.fieldbyname('estoque').asfloat;
      rProd_preco := frmModulo.PrecoDeVenda(query.fieldbyname('codigo').asinteger, rProd_qtde);
      if query.fieldbyname('DESCONTOMAXIMO').asfloat > 0 then
        rProd_desconto_maximo := query.fieldbyname('DESCONTOMAXIMO').asfloat
      else
        rProd_desconto_maximo := query.fieldbyname('desconto_maximo').asfloat;
      rProd_acrescimo_maximo := query.fieldbyname('acrecimo_maximo').asfloat;

      rProd_desconto := 0;
      rProd_acrescimo := 0;
      sProd_Tamanho := query.FieldByName('TAMANHO').AsString;
      sProd_Cor := query.FieldByName('COR').AsString;

      lbDestaque.Visible := query.fieldbyname('EMDESTAQUE').AsString = 'S';

      if ExibeEstoqueTelaInicial then
      begin
        lbEstoque.Caption := 'Estoque Atual: ' + FormatFloat('#,##0.00',query.fieldbyname('ESTOQUE').AsFloat);
        lbEstoque.Visible := True;
      end;

      if rProd_total > 0 then
      begin
        // o valor total foi achado caso o produto foi pesado
        // encontrar a qtde do produto
        rProd_qtde := rProd_total / rProd_preco;
      end;

      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
    end
    else
    begin
      // nao achou o produto
      // lancar no display do produto
      Imprime_display('PRODUTO NÃO ENCONTRADO!', clRed, tiErro);
      ed_barra.SelectAll;
    end;
  except
    Result := False;
  end;
end;

procedure TfrmVenda.LongPress;
var
  Produto: Integer;
begin
  if Controle = nil then
    Exit;
  Produto := Controle.Tag;
  if Produto > 0 then begin
    if Pos('Touch_Produto', Controle.Name) > 0 then begin
      if bNao_utiliza_extra then begin
        sObsProduto.Lines.Clear;
        RetTexto(sObsProduto, 'Cancelar (F3)', 'Confirmar (F2)', 'Informações Extras do Produto!');
        vExtra.Texto := sObsProduto.Lines.Text;
        vExtra.valor := 0;
      end
      else
        vExtra := RetornaAdcional(Zerar(IntToStr(Produto), 6));
      if Application.MessageBox('Lançar o produto no cupom?', 'Atenção!', MB_ICONQUESTION + MB_YESNO) = mrYes then begin
        if not bVenda then begin
          if not Abre_Venda then begin
            bLanca_pre_venda := False;
            bLanca_OS := False;
            bLanca_Mesa := False;
            bLanca_Abastecimento := False;
            Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
            Exit;
          end;
        end;
        iPesquisa_produto := 1;
        if frmVenda.Localizar_Produto(Zerar(IntToStr(Produto), 6)) then begin
          rProd_preco := rProd_preco + vExtra.valor;
          if (query.fieldbyname('usa_balanca').AsInteger = 1) and (bPesagemAutomatica) then begin
            ed_qtde.Value := RetornaPesoBalanca;
            rProd_qtde := ed_qtde.Value;
            if ed_qtde.Value = 0 then
              Exit;
          end
          else begin
            if ed_qtde.value > 1 then
              rProd_qtde := ed_qtde.Value
            else if ed_qtde.value <= 0 then
              ed_qtde.Value := rProd_qtde;
          end;
          ed_unitario.Value := rProd_preco;
          rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
          Registra_Item;
        end;
      end;
    end;
  end;
end;

// Cancela cupom aberto

procedure TfrmVenda.Cancela_cupom_aberto();
var
  i: Integer;
  scod_cupom: string;
begin
  Imprime_display('Aguarde! Cancelando cupom ' + sNumero_Cupom + '!', clYellow, tiAlerta);
  if TipoDeCupom = tcPafECF then
    cpACBrECF.CancelaCupom;
  if (bBusca_foto_produto) and not (AtivaTouch) then
    img_produto.Picture := nil;

  frmModulo.spCupom_Temp_delete.prepare;
  frmModulo.spCupom_Temp_delete.execute;
  bVenda := False;
  Limpa_controles;
  iTotal_Itens := 0;

  bLanca_pre_venda := False;
  bLanca_OS := False;
  bLanca_Abastecimento := False;

  if bVenda then
    gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'SIM')
  else
    gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'NAO');

end;

procedure TfrmVenda.Cancela_Item(sItem: string; Acao: string);
begin
  Application.ProcessMessages;
  query.SQL.Clear;
  if Acao = 'C' then begin
    Dec(iTotal_Itens);
    rTotal_Venda := rTotal_Venda - frmModulo.qrItensVALOR_TOTAL.AsFloat;
    query.SQL.Add('update CUPOM_ITEM_TEMP set cancelado = 1 where item = ' + IntToStr(StrToInt(sItem)));
  end
  else begin
    Dec(iTotal_Itens);
    rTotal_Venda := rTotal_Venda + frmModulo.qrItensVALOR_TOTAL.AsFloat;
    query.SQL.Add('update CUPOM_ITEM_TEMP set cancelado = 0 where item = ' + IntToStr(StrToInt(sItem)));
  end;
  query.ExecSQL;

  with frmModulo do begin
    qrItens.DisableControls;
    qrItens.Close;
    qrItens.Open;
    qrItens.Last;
    qrItens.EnableControls;
  end;

  // atualizar os labels de totais e itens
  lb_item.Caption := IntToStr(iTotal_Itens);
  lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);

  if Acao = 'C' then
    Imprime_display('ITEM ' + Zerar(frmModulo.qrItensITEM.AsString, 4) + ' CANCELADO!', clWhite, tiOk)
  else
    Imprime_display('ITEM ' + Zerar(frmModulo.qrItensITEM.AsString, 4) + ' DESCANCELADO!', clWhite, tiOk);
  Sleep(1000);
  Imprime_display('INFORME O PRODUTO...', clWhite, tiLivre);
  if not AtivaTouch then
    if ed_barra.Enabled then
      ed_barra.SETFOCUS;
end;

// -------------------------------------------------------------------------- //
// Registrar item no Grid

procedure TfrmVenda.Registra_Item;
var
  Codigo, TExtoExtra, sAliquota: string;
  i: Integer;
  Resto, Qtde: Double;
begin
  qrEstoque.Close;
  qrEstoque.Params.Items[0].Value := iProd_codigo;
  qrEstoque.Open;
  if (bSenhaVendaSemEstoque) or (bVerificaProdutoVenda) then begin
    if (bSenhaVendaSemEstoque) then begin
      if (qrEstoqueESTOQUE.AsFloat - (rProd_qtde * QtdVolume)) < 0 then begin
        Imprime_display('PERMISSÃO PARA VENDA SEM ESTOQUE!', clRed, tiLivre);

        frmsenha_supervisor := tfrmsenha_supervisor.create(self);
        frmsenha_supervisor.showmodal;
        if not bSupervisor_autenticado then
          Exit;
        FreeAndNil(frmsenha_supervisor);
      end;
    end;
    if (bVerificaProdutoVenda) and (TipoDeCupom in [tcSAT, tcNFCE]) and (frmPrincipal.TipoImpressora <> NaoFiscal) then begin
      if (trim(qrEstoqueCSOSN.AsString) = '') then begin
        Imprime_display('PRODUTO SEM CSOSN INFORMADO NO CADASTRO!', clRed, tiLivre);
        Exit;
      end;
      if (trim(qrEstoqueNCM.AsString) = '') then begin
        Imprime_display('PRODUTO SEM NCM INFORMADO NO CADASTRO!', clRed, tiLivre);
        Exit;
      end;
      if (trim(qrEstoqueCOD_BARRA.AsString) = '') and (qrEstoqueINFORMA_CODIGO_BARRA_XML.AsString = 'S') then begin
        Imprime_display('PRODUTO SEM CODIGO DE BARRA INFORMADO NO CADASTRO!', clRed, tiLivre);
        Exit;
      end;
      if (trim(qrEstoqueCFOP.AsString) = '') and (((TipoDeCupom = tcSAT) and (Sat_CfopPadrao = '')) or ((TipoDeCupom = tcNFCE) and (edtcfop = ''))) then begin
        Imprime_display('PRODUTO SEM CFOP INFORMADO NO CADASTRO!', clRed, tiLivre);
        Exit;
      end;
    end;
  end;
  if lbDevolucao.Visible then begin
    rProd_qtde := rProd_qtde *  - 1;
    rProd_total := rProd_total *  - 1;
    lbDevolucao.Visible := False;
  end;

  ///////////////////////////////////////////////////////////////////////////////////
  ///  SAIRA DAQUI E IR PARA FECHAMENTO

  if TipoDeCupom = tcPafECF then
  with frmModulo do
  begin
    if strPizza.ProdutoPrincipal.Valor > 0 then
    begin
      qrEstoque.Close;
      qrEstoque.Params.Items[0].Value := strPizza.ProdutoPrincipal.Codigo;
      qrEstoque.Open;
      qrEstoque.Locate('CODIGO', StrToInt(strPizza.ProdutoPrincipal.Codigo), [loCaseInsensitive]);
      cpACBrECF.VendeItem(strPizza.ProdutoPrincipal.Codigo
                         ,qrEstoqueNOME.AsString
                         ,qrEstoqueECF_ICMS.AsString
                         ,rProd_qtde
                         ,strPizza.ProdutoPrincipal.Valor
                         ,0
                         ,sProd_unidade
                           ,'%');
    end
    else
    begin
      qrEstoque.Close;
      qrEstoque.Params.Items[0].Value := iProd_codigo;
      qrEstoque.Open;
      qrEstoque.Locate('CODIGO', iProd_codigo, [loCaseInsensitive]);
      cpACBrECF.VendeItem(IntToStr(iProd_codigo)
                         ,qrEstoqueNOME.AsString
                         ,qrEstoqueECF_ICMS.AsString
                         ,rProd_qtde
                         ,rProd_preco
                         ,0
                         ,sProd_unidade
                           ,'%');
    end;
  end;

  /////////////////////////////////////////////////////////////////////////////////////////
  ///  Inclusão ou edição de registros de vendas
  if CodProduto > 0 then
  begin
    try
      uQry.Close;
      uQry.sql.Clear;
      uQry.sql.Add('UPDATE CUPOM_ITEM_TEMP SET COD_PRODUTO=:CODIGO, UNIDADE=:UNIDADE, QTDE=:QTDE, VALOR_UNITARIO=:VALOR_UNITARIO,');
      uQry.sql.add('VALOR_DESCONTO=:VALOR_DESCONTO, VALOR_ACRESCIMO=:VALOR_ACRESCIMO,');
      uQry.sql.add('VALOR_SUBTOTAL=:VALOR_SUBTOTAL, VALOR_TOTAL=:VALOR_TOTAL, POR_VOLUME=:POR_VOLUME');
      uQry.sql.add('WHERE CANCELADO <> 1 AND COD_PRODUTO=:COD_PRODUTO');
      uQry.ParamByName('CODIGO').AsInteger := iProd_codigo;
      uQry.ParamByName('UNIDADE').AsString := sProd_unidade;
      uQry.ParamByName('QTDE').AsFloat := rProd_qtde;
      uQry.ParamByName('VALOR_UNITARIO').AsFloat := rProd_preco;
      uQry.ParamByName('VALOR_DESCONTO').AsFloat := rProd_desconto;
      uQry.ParamByName('VALOR_ACRESCIMO').AsFloat := rProd_acrescimo;
      uQry.ParamByName('VALOR_SUBTOTAL').AsFloat := rProd_qtde * rProd_preco;
      uQry.ParamByName('VALOR_TOTAL').AsFloat := rProd_total - rProd_desconto + rProd_acrescimo;
      uQry.ParamByName('POR_VOLUME').AsInteger := VendaPorVolume;
      uQry.ParamByName('COD_PRODUTO').AsInteger := CodProduto;
      uQry.ExecSQL;
      frmModulo.qrItens.Refresh;
    finally
      CodProduto := 0;
    end;
  end
  else
  begin
    sAliquota := '';
    // verificar se o item tem acrescimo
    // lancar item no banco de dados temporario
    Codigo := codifica_item(iItem);
    frmModulo.spCupom_item_temp.Close;
    frmModulo.spCupom_item_temp.ParamByName('Codigo').AsString := Codigo;
    frmModulo.spCupom_item_temp.parambyname('cod_cupom').asstring := sNumero_Cupom;
    frmModulo.spCupom_item_temp.parambyname('item').asinteger := iItem;
    frmModulo.spCupom_item_temp.parambyname('unidade').asstring := sProd_unidade;
    frmModulo.spCupom_Item_temp.parambyname('qtde').asfloat := rProd_qtde;
    if strPizza.ProdutoPrincipal.Valor > 0 then
    begin
      frmModulo.spCupom_item_temp.ParamByName('cod_produto').asinteger := StrToInt(strPizza.ProdutoPrincipal.Codigo);
      frmModulo.spCupom_Item_temp.parambyname('valor_unitario').asfloat := strPizza.ProdutoPrincipal.Valor;
      frmModulo.spCupom_Item_temp.parambyname('valor_subtotal').asfloat := strPizza.ProdutoPrincipal.Valor;
      frmModulo.spCupom_Item_temp.parambyname('valor_desconto').asfloat := 0;
      frmModulo.spCupom_Item_temp.parambyname('valor_acrescimo').asfloat := 0;
      frmModulo.spCupom_Item_temp.parambyname('valor_total').asfloat := strPizza.ProdutoPrincipal.Valor;
      frmModulo.spCupom_Item_temp.parambyname('complemento').AsString := strPizza.Pizza;
      rProd_total := strPizza.ProdutoPrincipal.Valor;
    end
    else
    begin
      frmModulo.spCupom_item_temp.ParamByName('cod_produto').asinteger := iProd_codigo;
      frmModulo.spCupom_Item_temp.parambyname('valor_unitario').asfloat := rProd_preco; //+ rProd_desconto - rProd_acrescimo;
      frmModulo.spCupom_Item_temp.parambyname('valor_subtotal').asfloat := rProd_qtde * rProd_preco; // (rProd_preco + rProd_desconto - rProd_acrescimo);
      frmModulo.spCupom_Item_temp.parambyname('valor_desconto').asfloat := rProd_desconto;
      frmModulo.spCupom_Item_temp.parambyname('valor_acrescimo').asfloat := rProd_acrescimo;
      frmModulo.spCupom_Item_temp.parambyname('valor_total').asfloat := rProd_total - rProd_desconto + rProd_acrescimo; //  RoundTo(rProd_qtde * rProd_preco, -2);
      frmModulo.spCupom_Item_temp.parambyname('complemento').Clear;
    end;
    if Trim(sProd_Compl_Pizza) <> '' then
      frmModulo.spCupom_Item_temp.parambyname('complemento').AsString := Trim(sProd_Compl_Pizza);

    frmModulo.spCupom_item_temp.parambyname('cst').asstring := sProd_CST;
    frmModulo.spCupom_item_temp.parambyname('aliquota').asfloat := rProd_aliquota;
    frmModulo.spCupom_item_temp.parambyname('cancelado').asinteger := 0;
    frmModulo.spCupom_Item_temp.parambyname('tamanho').AsString := sProd_Tamanho;
    frmModulo.spCupom_Item_temp.parambyname('cor').AsString := sProd_Cor;
    frmModulo.spCupom_Item_temp.parambyname('por_volume').asinteger := VendaPorVolume;
    frmModulo.spCupom_item_temp.ParamByName('cod_totalizador').asstring := '';
    qrCombo.Close;
    qrCombo.Params.Items[0].Value := iProd_codigo;
    qrCombo.Open;
    TExtoExtra := '';
    if qrComboCOMBO.AsString = 'S' then
    begin
      qrProdComb.Close;
      qrProdComb.Params.Items[0].Value := Zerar(IntToStr(iProd_codigo), 6);
      qrProdComb.Open;
      if not qrProdComb.IsEmpty then
      begin
        qrProdComb.First;
        while not qrProdComb.Eof do
        begin
          TExtoExtra := TExtoExtra + qrProdCombQTDE.AsString + ' => ' + qrProdCombPRODUTO.AsString + #13;
          qrProdComb.Next;
        end;
      end;
      if vExtra.Texto <> '' then
        TExtoExtra := TExtoExtra + vExtra.Texto;
    end
    else if strPizza.Valor > 0 then
    begin
      TExtoExtra := strPizza.Sabores;
    end
    else begin
      if vExtra.Texto <> '' then
        TExtoExtra := TExtoExtra + vExtra.Texto;
    end;

    frmModulo.spCupom_item_temp.ParamByName('extra').asstring := TExtoExtra;

    frmModulo.spCupom_item_temp.Prepare;
    frmModulo.spCupom_item_temp.Execute;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
  ///  sei la vou analisar
  if (cdsLancaProdutos.Active) and (cdsLancaProdutos.RecordCount > 0) then
  begin
    cdsLancaProdutos.First;
    while not cdsLancaProdutos.Eof do
    begin
      cdsLancaProdutos.Edit;
      cdsLancaProdutosCOD_ITEM.AsString := Codigo;
      cdsLancaProdutos.Post;
      cdsLancaProdutos.Next;
    end;
  end;
  if Length(strPizza.Produtos) > 0 then
  begin
    Qtde := RoundTo(1 / Length(strPizza.Produtos), -2);
    Resto := RoundTo(1 - (Qtde * Length(strPizza.Produtos)), -2);
    for i := 0 to Length(strPizza.Produtos) - 1 do
    begin
      cdsLancaProdutos.Insert;
      cdsLancaProdutosCOD_ITEM.AsString := Codigo;
      cdsLancaProdutosCODPRO.AsString := strPizza.Produtos[i];
      if (i = Length(strPizza.Produtos) - 1) and (Resto > 0) then
        cdsLancaProdutosQTDE.AsFloat := Qtde + Resto
      else
        cdsLancaProdutosQTDE.AsFloat := Qtde;
      cdsLancaProdutos.Post;
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
  if not AtivaTouch then
    if ed_barra.Enabled then
      ed_barra.SetFocus;
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
  ///  Recalcula totais e organiza registros de vendas
  rTotal_Desconto := 0;
  rTotal_Acrescimo := 0;
  rTotal_Venda := 0;
  with uQry do
  begin
    Close;
    sql.Clear;
    sql.Add('SELECT * FROM CUPOM_ITEM_TEMP');
    Open;
    First;
    while not EOF do
    begin
      Edit;
      FieldByName('CODIGO').AsString := codifica_item(RecNo);
      FieldByName('ITEM').AsInteger := RecNo;
      Post;
      if FieldByName('CANCELADO').AsInteger = 0 then
      begin
        rProd_desconto := FieldByName('VALOR_DESCONTO').AsFloat;
        rProd_acrescimo := FieldByName('VALOR_ACRESCIMO').AsFloat;
        rTotal_Desconto := rTotal_Desconto + rProd_desconto;
        rTotal_Acrescimo := rTotal_Acrescimo + rProd_acrescimo;
        rTotal_Venda := rTotal_Venda + (FieldByName('VALOR_TOTAL').AsFloat - rProd_desconto + rProd_acrescimo);
      end;
      Next;
    end;
    iTotal_Itens := RecordCount;
    iItem := Succ(iTotal_Itens);
  end;

  // lancar no dislpay o produto
  Imprime_display(sProd_nome, clWhite, tiOk);
  // atualizar os labels de totais e itens
  lb_item.Caption := IntToStr(iTotal_Itens);
  lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
  ed_barra.text := '';
  rBal_peso := 0;
  vExtra.Texto := '';
  vExtra.valor := 0;
  ed_qtde.Value := 0;
  strPizza.Pizza := '';
  strPizza.Sabores := '';
  strPizza.Valor := 0;
  SetLength(strPizza.Produtos, 0);
  strPizza.ProdutoPrincipal.Codigo := '';
  strPizza.ProdutoPrincipal.Valor := 0;
  sProd_Compl_Pizza := '';
end;

procedure TfrmVenda.RelCupomNFRestauranteBeforePrint(Sender: TfrxReportComponent);
begin
  if Sender is TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'cptaxa' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00', rTotal_Comissao);
    if TfrxMemoView(Sender).Name = 'cpCouvert' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00', rTotal_Couvert);
  end;
end;

procedure TfrmVenda.RelFichaBeforePrint(Sender: TfrxReportComponent);
begin
  if Sender is TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'lbValidador' then
      TfrxMemoView(Sender).Text := sCodigoSegurancaFicha;
  end;

end;

procedure TfrmVenda.RemoveDescontoTabelaItem;
begin
  with frmModulo do begin
    try
      qrItens.DisableControls;
      qrItens.First;
      while not qrItens.Eof do begin
        qrItens.Edit;
        qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat - qrItensVALOR_DESCONTO.AsFloat, -2);
        qrItensVALOR_DESCONTO.AsFloat := 0;
        qrItens.Post;
        qrItens.Next;
      end;
    finally
      qrItens.Last;
      qrItens.EnableControls;
    end;
  end;

end;

// -------------------------------------------------------------------------- //
// Limpar os controles para receber uma nova venda/comando

procedure TfrmVenda.LancaMesaComanda(MesasComandas: string);
var
  qtdCouver, nItem: Integer;
  TotVen, ValorItem, TotDestaVenda, vlrAcressimoTaxa, vlrAcressimoCouver, Maior, ppValor: Double;
  ppCodigo: string;
begin
  frmModulo.query_servidor.first;
  nItem := 1;

  frmModulo.qrMesas.Close;
  frmModulo.qrMesas.MacroByName('mesas').Value := MesasComandasReceber;
  frmModulo.qrMesas.Open;

  frmModulo.qrMesas.First;
  rTotal_Comissao := 0;
  rTotal_Couvert := 0;
  qtdCouver := frmModulo.qrMesas.RecordCount;
  frmModulo.qrproduto.Close;
  frmModulo.qrproduto.Open;
  while not frmModulo.qrMesas.Eof do begin
{    frmmodulo.query_servidor.close;
    frmmodulo.query_servidor.sql.clear;
    frmmodulo.query_servidor.SQL.add('select r000002.*, C000025.produto, C000025.unidade,');
    frmmodulo.query_servidor.sql.add('c000025.cst, c000025.aliquota, c000025.codbarra');
    frmmodulo.query_servidor.sql.add('from r000002, c000025');
    frmmodulo.query_servidor.sql.add('where r000002.cod_produto = c000025.codigo');
    frmmodulo.query_servidor.sql.add('and r000002.cancelado = 0 and r000002.cod_mesa = ' + frmModulo.qrMesas.FieldByName('CODIGO').AsString);
    frmmodulo.query_servidor.sql.add('order by r000002.codigo');
    frmmodulo.query_servidor.open;}
    TotDestaVenda := 0;
    while not frmModulo.query_servidor.eof do begin
      if frmModulo.query_servidor.fieldbyname('cancelado').asfloat = 0 then begin
        cdsLancaProdutos.Close;
        cdsLancaProdutos.CreateDataSet;
        if (frmModulo.qrconfigTAXA_GARCOM_TIPO_RECEB.AsString = 'A') and ((frmModulo.qrMesasCLIENTE_AUTORIZOU.AsString = 'S') or (frmModulo.qrconfigCOBRA_COMISSAO.AsString = 'O')) then
          vlrAcressimoTaxa := (frmModulo.query_servidor.fieldbyname('total').asfloat * frmModulo.qrconfigCOMISSAO_PERCENTUAL.AsFloat) / 100
        else
          vlrAcressimoTaxa := 0;

        sProd_nome := frmModulo.query_servidor.fieldbyname('produto').asstring;
        sProd_unidade := frmModulo.query_servidor.fieldbyname('unidade').asstring;
        sProd_CST := frmModulo.query_servidor.fieldbyname('cst').asstring;
        rProd_aliquota := frmModulo.query_servidor.fieldbyname('aliquota').asfloat;
        iProd_codigo := frmModulo.query_servidor.fieldbyname('cod_produto').asinteger;
        sProd_barra := frmModulo.query_servidor.fieldbyname('codbarra').asstring;
        rProd_qtde := frmModulo.query_servidor.fieldbyname('qtde').asfloat;
        rProd_preco := frmModulo.query_servidor.fieldbyname('unitario').asfloat;
        rProd_total := frmModulo.query_servidor.fieldbyname('total').asfloat;
        sProd_Compl_Pizza := frmModulo.query_servidor.fieldbyname('complemento').AsString;
        rProd_desconto := 0;
        rProd_acrescimo := frmModulo.query_servidor.fieldbyname('acrescimo').asfloat + vlrAcressimoTaxa;
        sProd_Tamanho := '';
        sProd_Cor := '';
        vExtra.Texto := frmModulo.query_servidor.fieldbyname('extra').AsString;
        TotDestaVenda := TotDestaVenda + rProd_total;
        // registrar o item
        frmModulo.query_servidor2.Close;
        frmModulo.query_servidor2.SQL.Clear;
        frmModulo.query_servidor2.SQL.Add('select * from r000004 where codigo = ' + QuotedStr(frmModulo.query_servidor.FieldByName('codigo').AsString));
        frmModulo.query_servidor2.Open;
        frmModulo.query_servidor2.First;
        strPizza.Pizza := '';
        strPizza.Sabores := '';
        strPizza.Valor := 0;
        SetLength(strPizza.Produtos, 0);
        strPizza.ProdutoPrincipal.Codigo := '';
        strPizza.ProdutoPrincipal.Valor := 0;
        while not frmModulo.query_servidor2.Eof do begin
          cdsLancaProdutos.Insert;
          cdsLancaProdutosCODPRO.AsString := frmModulo.query_servidor2.FieldByName('CODPRO').AsString;
          cdsLancaProdutosQTDE.AsFloat := frmModulo.query_servidor2.FieldByName('QTDE').AsFloat;
          cdsLancaProdutos.Post;
          frmModulo.query_servidor2.Next;
        end;
        Registra_Item;
        rTotal_Comissao := rTotal_Comissao + vlrAcressimoTaxa;
        Inc(nItem);
      end;
      Application.ProcessMessages;
      frmModulo.query_servidor.next;
    end;
    if (frmModulo.qrconfigTAXA_GARCOM_TIPO_RECEB.AsString = 'P') and ((frmModulo.qrMesasCLIENTE_AUTORIZOU.AsString = 'S') or (frmModulo.qrconfigCOBRA_COMISSAO.AsString = 'O')) then begin
      vlrAcressimoTaxa := (TotDestaVenda * frmModulo.qrconfigCOMISSAO_PERCENTUAL.AsFloat) / 100;
      if frmModulo.qrproduto.Locate('codigo', frmModulo.qrconfigTAXA_GARCOM_PRODUTO.AsString, [loCaseInsensitive]) then begin
        sProd_nome := frmModulo.qrprodutoproduto.asstring;
        sProd_unidade := frmModulo.qrprodutounidade.asstring;
        sProd_CST := frmModulo.qrprodutocst.asstring;
        rProd_aliquota := frmModulo.qrprodutoaliquota.asfloat;
        iProd_codigo := frmModulo.qrprodutoCODIGO.asinteger;
        sProd_barra := frmModulo.qrprodutocodbarra.asstring;
        rProd_qtde := 1;
        rProd_preco := vlrAcressimoTaxa;
        rProd_total := vlrAcressimoTaxa;
        sProd_Compl_Pizza := '';
        rProd_desconto := 0;
        rProd_acrescimo := 0;
        sProd_Tamanho := '';
        sProd_Cor := '';
        Registra_Item;
        rTotal_Comissao := rTotal_Comissao + vlrAcressimoTaxa;
        Inc(nItem);
      end;
    end;
    frmModulo.qrMesas.Next;
  end;
  if (frmModulo.qrconfigCOBRA_COUVER.AsString = 'S') then begin
    if frmModulo.qrconfigCOUVER_TIPO_RECEB.AsString = 'A' then begin
      frmModulo.qrItens.Close;
      frmModulo.qrItens.Open;
      frmModulo.qrItens.First;
      vlrAcressimoCouver := 0;
      TotVen := 0;
      while not frmModulo.qrItens.Eof do begin
        ValorItem := frmModulo.qrItensVALOR_UNITARIO.AsFloat;
        TotVen := TotVen + frmModulo.qrItensQTDE.AsFloat * ValorItem;
        frmModulo.qrItens.Next;
      end;
      frmModulo.qrItens.First;
      while not frmModulo.qrItens.Eof do begin
        ValorItem := frmModulo.qrItensVALOR_UNITARIO.AsFloat;
        vlrAcressimoCouver := vlrAcressimoCouver + RoundTo(((frmModulo.qrItensQTDE.AsFloat * ValorItem) / TotVen) * (frmModulo.qrconfigVALOR_COUVER.AsFloat * qtdCouver), -2);
        frmModulo.qrItens.Edit;
        frmModulo.qrItensVALOR_ACRESCIMO.AsFloat := frmModulo.qrItensVALOR_ACRESCIMO.AsFloat + RoundTo(((frmModulo.qrItensQTDE.AsFloat * ValorItem) / TotVen) * (frmModulo.qrconfigVALOR_COUVER.AsFloat * qtdCouver), -2);
        frmModulo.qrItensVALOR_TOTAL.AsFloat := frmModulo.qrItensVALOR_TOTAL.AsFloat + RoundTo(((frmModulo.qrItensQTDE.AsFloat * ValorItem) / TotVen) * (frmModulo.qrconfigVALOR_COUVER.AsFloat * qtdCouver), -2);
        frmModulo.qrItens.Next;
      end;

      rTotal_Couvert := frmModulo.qrconfigVALOR_COUVER.AsFloat * qtdCouver;
      if vlrAcressimoCouver <> rTotal_Couvert then begin
        frmModulo.qrItens.Prior;
        frmModulo.qrItens.Edit;
        frmModulo.qrItensVALOR_ACRESCIMO.AsFloat := frmModulo.qrItensVALOR_ACRESCIMO.AsFloat + rTotal_Couvert - vlrAcressimoCouver;
        frmModulo.qrItensVALOR_TOTAL.AsFloat := frmModulo.qrItensVALOR_TOTAL.AsFloat + rTotal_Couvert - vlrAcressimoCouver;
        frmModulo.qrItens.Next;
      end;
      rTotal_Venda := rTotal_Venda + rTotal_Couvert;
    end
    else begin
      if frmModulo.qrproduto.Locate('codigo', frmModulo.qrconfigCOUVER_PRODUTO.AsString, [loCaseInsensitive]) then begin
        sProd_nome := frmModulo.qrprodutoproduto.asstring;
        sProd_unidade := frmModulo.qrprodutounidade.asstring;
        sProd_CST := frmModulo.qrprodutocst.asstring;
        rProd_aliquota := frmModulo.qrprodutoaliquota.asfloat;
        iProd_codigo := frmModulo.qrprodutoCODIGO.asinteger;
        sProd_barra := frmModulo.qrprodutocodbarra.asstring;
        rProd_qtde := qtdCouver;
        rProd_preco := frmModulo.qrconfigVALOR_COUVER.AsFloat;
        rProd_total := qtdCouver * frmModulo.qrconfigVALOR_COUVER.AsFloat;
        sProd_Compl_Pizza := '';
        rProd_desconto := 0;
        rProd_acrescimo := 0;
        sProd_Tamanho := '';
        sProd_Cor := '';
        Registra_Item;
        rTotal_Couvert := rProd_total;
        Inc(nItem);
      end;

    end;
  end;

  if rTotal_Venda <= 0 then begin
    Cancela_cupom_aberto;
  end
  else begin
    lb_item.Caption := IntToStr(iTotal_Itens);
    lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
    FecharCupom1Click(frmVenda);
  end;

end;

procedure TfrmVenda.Limpa_controles();
begin
  sCli_Nome := '';
  sCli_Endereco := '';
  sCli_Cidade := '';
  sCli_uf := '';
  scli_cep := '';
  sCli_CPF := '';
  sCli_Placa := '';
  sCli_Km := '';
  scli_vendedor := '';

  sConsumidor_CPF := '';
  sConsumidor_Nome := '';
  sConsumidor_Endereco := '';

  ed_barra.Text := '';
  ed_qtde.value := 0;
  ed_unitario.value := 0;
  ed_total_item.value := 0;
  lb_total.caption := 'R$ 0,00';
  lb_item.Caption := '0';
  Imprime_display('C A I X A    L I V R E', clWhite, tiLivre);
  MostraFotoProduto(False);
  MostraLogoMarca(True);
  if not AtivaTouch then
    if ed_barra.Enabled then
      ed_barra.setfocus;
end;

// -------------------------------------------------------------------------- //
// funcao para alimentar o display superior da tela

procedure TfrmVenda.Imprime_display(texto: string; cor_da_fonte: TColor; Imagem: Tipo_Imagem);
begin
  // armazena os dados atuais do display para enventual uso
  sTexto_anterior := lb_produto.Caption;
  clCor_anterior := lb_produto.Font.color;

  // inicia a impressao do display

  lb_produto.Caption := texto;
  lb_produto.Font.Color := cor_da_fonte;

  lb_produto.Font.Size := 28;
  Application.ProcessMessages;
   // redimensionando o tamanho do label para caber dentro do display independente do
   // tamanho do texto;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Imprime_display_anterior();
begin
  Imprime_display(sTexto_anterior, clCor_anterior, TiImagem_anterior);
end;

procedure TfrmVenda.ImprimirComanda(Cupom: string);
begin
  if not CancelaIMpressaoCozinha then begin
    qrComanda.Close;
    qrComanda.Params.Items[0].Value := Cupom;
    qrComanda.Open;
    if not (bMensagem_Cupom) or not (Men_Cup_Cozinha) then begin
      cdsOBS.Close;
      cdsOBS.CreateDataSet;
    end
    else begin
      if not cdsOBS.Active then begin
        cdsOBS.Close;
        cdsOBS.CreateDataSet;
        if (bMensagem_Cupom) and (Men_Cup_Simples) then begin
          cdsOBS.Append;
          cdsOBSobs.AsString := sMensagem_Cupom.Lines.Text;
          cdsOBS.Post;
        end;
      end;

    end;

    frComanda.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000006.fr3');
    frComanda.PrintOptions.ShowDialog := False;
    frComanda.PrintOptions.Printer := sImpressora_Cozinha;
    frComanda.PrepareReport;
    if bEdicaoRelatorio then begin
      frComanda.DesignReport;
      qrTmp.SQL.Clear;
      qrTmp.SQL.Add('update cupom set impresso = ' + QuotedStr('S') + ' where codigo = ' + QuotedStr(Cupom));
      qrTmp.ExecSQL;
    end
    else begin
      if bExibeCumponNFTela then begin
        frComanda.ShowPreparedReport;
        qrTmp.SQL.Clear;
        qrTmp.SQL.Add('update cupom set impresso = ' + QuotedStr('S') + ' where codigo = ' + QuotedStr(Cupom));
        qrTmp.ExecSQL;
      end
      else if frComanda.Print then begin
        qrTmp.SQL.Clear;
        qrTmp.SQL.Add('update cupom set impresso = ' + QuotedStr('S') + ' where codigo = ' + QuotedStr(Cupom));
        qrTmp.ExecSQL;
      end;
    end;
  end;
end;

procedure TfrmVenda.ImprimirCupomNaoFiscal(Cupom: string);
var
  i: Integer;
  Acao: Boolean;
begin
  if bAtivaImpFicha then
  begin
    RelFicha.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000007.fr3');
    if sImpressora_Ficha <> '' then
      RelFicha.PrintOptions.Printer := sImpressora_Ficha;
    qrCupomNF.Close;
    qrCupomNF.Params.Items[0].Value := Cupom;
    qrCupomNF.Open;
    qrCupomNF.First;
    while not qrCupomNF.Eof do begin
      for i := 1 to qrCupomNFQTDE.AsInteger do begin
        if bEdicaoRelatorio then
          RelFicha.DesignReport
        else
          RelFicha.ShowReport;
      end;
      qrCupomNF.Next;
    end;
  end
  else
  begin
    if (rTotal_Couvert + rTotal_Comissao) > 0 then
      RelCupomNFRestaurante.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000005.fr3')
    else
      RelCupomNF.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000004.fr3');

    if sImpressora_Venda_Simples <> '' then
      if (rTotal_Couvert + rTotal_Comissao) > 0 then
        RelCupomNFRestaurante.PrintOptions.Printer := sImpressora_Venda_Simples
      else
        RelCupomNF.PrintOptions.Printer := sImpressora_Venda_Simples;

    qrCupomNF.Close;
    qrCupomNF.Params.Items[0].Value := Cupom;
    qrCupomNF.Open;

    cdsOBS.Close;
    cdsOBS.CreateDataSet;

    if (bMensagem_Cupom) and (Men_Cup_Simples) then
    begin
      cdsOBS.Append;
      cdsOBSobs.AsString := sMensagem_Cupom.Lines.Text;
      cdsOBS.Post;
    end;

    qrImpForma.Close;
    qrImpForma.Params.Items[0].Value := Cupom;
    qrImpForma.Open;

    if bEdicaoRelatorio then
    begin
      if (rTotal_Couvert + rTotal_Comissao) > 0 then
        RelCupomNFRestaurante.DesignReport
      else
        RelCupomNF.DesignReport;
      Acao := True;
    end
    else
    begin
      if TipoImprimeCupom = 'P' then
      begin
        Acao := False;
        if Application.MessageBox('Deseja Imprimir o Cupom?', 'Atenção!', MB_ICONQUESTION + MB_YESNO) = IDYES then
        begin
          Acao := True;
          if bExibeCumponNFTela then
            if (rTotal_Couvert + rTotal_Comissao) > 0 then
              RelCupomNFRestaurante.ShowReport
            else
              RelCupomNF.ShowReport
          else if (rTotal_Couvert + rTotal_Comissao) > 0 then
          begin
            RelCupomNFRestaurante.PrepareReport(True);
            RelCupomNFRestaurante.Print
          end
          else
          begin
            RelCupomNF.PrepareReport(True);
            RelCupomNF.Print;
          end;
        end;
      end
      else
      begin
        Acao := True;
        if bExibeCumponNFTela then
          if (rTotal_Couvert + rTotal_Comissao) > 0 then
            RelCupomNFRestaurante.ShowReport
          else
            RelCupomNF.ShowReport
        else
        if (rTotal_Couvert + rTotal_Comissao) > 0 then
        begin
          RelCupomNFRestaurante.PrepareReport(True);
          RelCupomNFRestaurante.Print
        end
        else
        begin
          RelCupomNF.PrepareReport(True);
          RelCupomNF.Print;
        end;
      end;
    end;

    if Acao then
      ImprimirComprovanteCrediario;
  end;
end;

// -------------------------------------------------------------------------- //
// Funcao para iniciar a venda

function TfrmVenda.Abre_Venda(ReabrirCupom: Boolean): Boolean;
var
  ok:Boolean;
  Est: String;
begin
  lbDestaque.Visible := False;
  TimerTroco.Enabled := False;
  try
    if TipoDeCupom =  tcPafECF then begin
      try
        Est := Estados[cpACBrECF.Estado];
        cpACBrECF.TestaPodeAbrirCupom;
      except
        Application.MessageBox(PWideChar('Não foi possivel iniciar a venda:'+#13+est),'Atenção!',MB_ICONERROR);
        raise;
      end;
    end;
    lbEstoque.Visible := False;
    cdsCrediario.Close;
    cdsCrediario.CreateDataSet;
    RecebeuTEF := False;
    bFezDevolucao := False;
    cdsLancaProdutos.Close;
    cdsLancaProdutos.CreateDataSet;
    bLanca_Delivery := False;
    iID_Delivery := -1;
    TotalRecebidoTef := 0;
    rTotal_TaxaEntrega := 0;
    rTotal_Couvert := 0;
    rTotal_Comissao := 0;
    rTotal_Desconto := 0;
    rTotal_Acrescimo := 0;
    bCadastra_Crediario := False;
    rvalor_total_crediario := 0;

    if not ReabrirCupom then
    begin
      // Pega Menor numero cupom
      sNumero_Cupom := Zerar(FloatToStr(Max('')), 5);
      sNumero_contador_cupom := Zerar(sNumero_Cupom, 5);

      if Length(sNumero_Cupom) = 5 then
        sNumero_Cupom := '9' + sNumero_Cupom;

      if Length(sNumero_contador_cupom) = 5 then
        sNumero_contador_cupom := '9' + sNumero_contador_cupom;
    end
    else
    begin
      sNumero_contador_cupom := sNumero_Cupom;
    end;

    // VERIFICAR O TOTALIZADOR

    if (bLanca_pre_venda) or (bLanca_OS) or (bLanca_Abastecimento) then begin
      if bLanca_pre_venda then begin
        query.close;
        query.sql.clear;
        query.sql.add('select * from cliente where codigo = ' + IntToStr(iPre_venda_cliente));
        query.open;
      end
      else begin
        if bLanca_OS then begin
          query.close;
          query.sql.clear;
          query.sql.add('select * from cliente where codigo = ' + IntToStr(iOS_cliente));
          query.open;
        end;
      end;

      if (bLanca_OS) or (bLanca_pre_venda) then begin
        sConsumidor_CPF := query.fieldbyname('cpf').asstring;
        sConsumidor_Nome := query.fieldbyname('nome').asstring;
        sConsumidor_Endereco := query.fieldbyname('endereco').asstring + ' ' + query.fieldbyname('bairro').asstring + '  ' + query.fieldbyname('cidade').asstring + '  ' + query.fieldbyname('uf').asstring;
      end
      else begin
        sConsumidor_CPF := '';
        sConsumidor_Nome := '';
        sConsumidor_Endereco := '';
      end;
    end
    else begin
      if bIdentifica_consumidor then begin
        frmIdentifica := TfrmIdentifica.Create(Self);
        frmIdentifica.ShowModal;
        if (sConsumidor_CPF <> '') and (TipoDeCupom <> tcSimples) then
          frmPrincipal.TipoImpressora := SemImpressora;
        CorEditTotaL;
      end;
    end;

    iVendedorCodigo := 0;

    sVendedorNome := '';

    if bIdentificarVendedor = 1 then begin
      Application.CreateForm(TfrmVendedor, frmVendedor);
      try
        if frmVendedor.ShowModal = mrOk then begin
          iVendedorCodigo := StrToInt(frmModulo.Query.FieldByName('CODIGO').AsString);
          sVendedorNome := frmModulo.Query.FieldByName('NOME').AsString;
        end
        else begin
          result := False;
          Close;
          Exit;
        end
      finally
        FreeAndNil(frmVendedor);
      end;
    end;
    // buscar variaveis do ECF
    // numero do cupom
    sMsg := sNumero_Cupom;

    if Length(sMsg) <> 6 then begin
      Imprime_display(sMsg, clRed, tiErro);
      result := False;
      Exit;
    end
    else begin
      sNumero_Cupom := sMsg;
      sNumero_Venda := sMsg;
    end;

    // hora do cupom
    sMsg := TimeToStr(Now);

    try
      hHora_Cupom := StrToTime(sMsg);
    except
      Imprime_display('Hora inválida!', clRed, tiErro);
      result := False;
      Exit;
    end;

    Imprime_display('ABRINDO O CUPOM...', clWhite, tiInfo);

    if TipoDeCupom = tcPafECF then begin
      ok := False;
      cpACBrECF.AbreCupom(sConsumidor_CPF, sConsumidor_Nome, sConsumidor_Endereco);
      sNumero_Cupom := cpACBrECF.NumCupom;
    end;


    // lancar abertura do cupom na tabela temporaria
    // alimentar as variaveis de status
    bVenda := True;
    bTotalizado := False;
    bPago1 := False;
    bPago2 := False;
    bPago3 := False;
    bFinalizado := False;




    query.close;
    query.sql.clear;
    query.sql.add('select numero from cupom_temp where numero = ' + QuotedStr(sNumero_Cupom));
    query.open;

    if query.RecordCount = 0 then
    begin
      frmModulo.spCupom_Temp.Close;
      frmModulo.spCupom_Temp.ParamByName('numero').asstring := sNumero_Cupom;
      if TipoDeCupom = tcPafECF then begin
        sNumero_contador_cupom := cpACBrECF.NumCCF;
        frmModulo.spCupom_Temp.ParamByName('ccf').asstring := cpACBrECF.NumCCF;
        frmModulo.spCupom_Temp.ParamByName('ecf').asstring := cpACBrECF.NumECF;
      end else begin
        frmModulo.spCupom_Temp.ParamByName('ccf').asstring := sNumero_contador_cupom;
        frmModulo.spCupom_Temp.ParamByName('ecf').asstring := sCaixa;
      end;
      frmModulo.spCupom_Temp.ParamByName('data').AsDate := dData_Sistema;
      frmModulo.spCupom_Temp.ParamByName('hora').AsTime := hHora_Cupom;
      frmModulo.spCupom_temp.Prepare;
      frmModulo.spCupom_temp.Execute;
    end;
    ///
    if not ReabrirCupom then
    begin
      // alimentar as variaveis de totais
      rTotal_Venda := 0;
      iTotal_Itens := 0;
      SenhaImpressao := 0;
      lb_item.Caption := '000';
      lb_total.Caption := 'R$ 0,00';
      iItem := 1;
      gravaINI(sConfiguracoes, 'PDV', 'NumCupom', sNumero_Cupom);
    end
    else
    with frmModulo, qrItens do
    begin
      DisableControls;
      Close;
      Open;
      First;
      while not EOF do
      begin
        // acumluar o item nas variaveis de totalizacao
        if qrItensCANCELADO.Value = 0  then
        begin
          rTotal_Desconto := rTotal_Desconto + qrItensVALOR_DESCONTO.Value;
          rTotal_Acrescimo := rTotal_Acrescimo + qrItensVALOR_ACRESCIMO.Value;
          rTotal_Venda := rTotal_Venda + qrItensVALOR_TOTAL.Value;
        End;
        Next;
      end;
      iTotal_Itens := RecordCount;
      iItem := Succ(iTotal_Itens);

      // lancar no dislpay o produto
      Imprime_display(qrItensNOME.Value, clWhite, tiOk);
      // atualizar os labels de totais e itens
      lb_item.Caption := IntToStr(iTotal_Itens);
      lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
      ed_barra.text := '';
      rBal_peso := 0;
      vExtra.Texto := '';
      vExtra.valor := 0;
      ed_qtde.Value := 0;
      strPizza.Pizza := '';
      strPizza.Sabores := '';
      strPizza.Valor := 0;
      SetLength(strPizza.Produtos, 0);
      strPizza.ProdutoPrincipal.Codigo := '';
      strPizza.ProdutoPrincipal.Valor := 0;
      sProd_Compl_Pizza := '';
      ///
      EnableControls;
    end;

    if (bBusca_foto_produto) and not (AtivaTouch) then
      img_produto.Picture := nil;

    frmTEF.estadoSimuladoEcf := tpsVenda;
    IndiceTransacaoTef := -1;
    result := True;
  except
    result := False;
  end;
  ///
  if not ReabrirCupom then
  begin
    if bVenda then
      gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'SIM')
    else
      gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'NAO');
  end;
end;

// -------------------------------------------------------------------------- //
// Funcao para truncar o valor para nao dar diferencao no ECF

function TfrmVenda.Truncar_Valor(valor: Real): Real;
var
  txt_desconto: string;
  i: Integer;
begin
  txt_desconto := FormatFloat('#############0.00000000000', valor);
  i := Pos(',', txt_desconto);
  if i > 0 then begin
    txt_desconto := Copy(txt_desconto, 1, i - 1) + ',' + Copy(txt_desconto, i + 1, 2);
    result := StrToFloat(txt_desconto);
  end
  else
    result := valor;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // verificar se ha cupom aberto, nao permitindo a saida do sistema
  if bVenda then begin
    Imprime_display('CUPOM ABERTO!', clYellow, tiAlerta);
    Sleep(1500);
    Imprime_display('INFORME O PRODUTO...', clWhite, tiLivre);
    if not AtivaTouch then
      if ed_barra.Enabled then
        ed_barra.SetFocus;
    Abort;
  end
  else begin
    if (TipoDeCupom = tcNFCE) then begin
      CheckConnect.Terminate;
      FreeAndNil(CheckConnect);
    end;
    timer_carga.Enabled := False;
    TimerTroco.Enabled := False;
    tmContigencia.Enabled := False;
  end;
  tmCaixa.Enabled := False;
  FreeAndNil(sMensagem_Cupom);
  FreeAndNil(sObsProduto);
  if TEF_Ativo then
    FreeAndNil(frmTEF);
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FormCreate(Sender: TObject);
var
  Ini: TIniFile;
begin
  wspnlAlerta.Left := 3000; //qiqi
  wspnlAlerta.Top := 0;
  v_tipo_normal.Visible := not AtivaTouch;

  CodProduto := 0;
  v_fundo.Transparent := True;
  v_borda_superior.Transparent := True;
  v_borda_inferior.Transparent := True;
  v_borda_esquerda.Transparent := True;
  v_borda_direita.Transparent := True;
  v_espacamento_superior.Transparent := True;
  v_espacamento_central.Transparent := True;
  v_espacamento_inferior.Transparent := True;
  v_barra_superior.Transparent := True;
  v_barra_inferior.Transparent := True;
  v_quadro_esquerdo.Transparent := True;
  if AtivaTouch then
    p_edit_quantidade := 0.38
  else
    p_edit_quantidade := 0.28;
  p_edit_valor_unitario := 0.48;

  Ini := TIniFile.Create(sConfiguracoes);
  try
    p_borda_superior := Ini.ReadFloat('Visual', 'p_borda_superior', v_borda_superior.Height/v_fundo.Height);
    p_borda_inferior := Ini.ReadFloat('Visual', 'p_borda_inferior', v_borda_inferior.Height/v_fundo.Height);
    p_borda_esquerda := Ini.ReadFloat('Visual', 'p_borda_esquerda', v_borda_esquerda.Width/v_fundo.Width);
    p_borda_direita := Ini.ReadFloat('Visual', 'p_borda_direita', v_borda_direita.Width/v_fundo.Width);

    p_espacamento_superior := Ini.ReadFloat('Visual', 'p_espacamento_superior', v_espacamento_superior.Height/v_fundo.Height);
    p_espacamento_central := Ini.ReadFloat('Visual', 'p_espacamento_central', v_espacamento_central.Width/v_fundo.Width);
    p_espacamento_inferior := Ini.ReadFloat('Visual', 'p_espacamento_inferior', v_espacamento_inferior.Height/v_fundo.Height);

    p_barra_superior := Ini.ReadFloat('Visual', 'p_barra_superior', v_barra_superior.Height/v_fundo.Height);
    p_barra_inferior := Ini.ReadFloat('Visual', 'p_barra_inferior', v_barra_inferior.Height/v_fundo.Height);

    p_quadro_esquerdo := Ini.ReadFloat('Visual', 'p_quadro_esquerdo', v_quadro_esquerdo.Width/v_fundo.Width);
  finally
    Ini.Free;
  end;

  bVerificaIniternet := True;
  cdsDevolucao.Close;
  cdsDevolucao.CreateDataSet;
  sMensagem_Cupom := TMemo.Create(Self);
  sMensagem_Cupom.Visible := False;
  sMensagem_Cupom.Parent := Self;
  sMensagem_Cupom.WordWrap := False;

  sObsProduto := TMemo.Create(Self);
  sObsProduto.Visible := False;
  sObsProduto.Parent := Self;
  sObsProduto.WordWrap := False;

  vNumNFCe := 0;
  Carregando := True;
  frmModulo.LerConfiguracao;
  if (TipoDeCupom = tcSimples) then
    frmPrincipal.TipoImpressora := NaoFiscal
  else
    frmPrincipal.TipoImpressora := SemImpressora;
  CorEditTotaL;
  if EdFormatoOff = 0 then
    frmtOffLine := teContingencia
  else
    frmtOffLine := teOffLine;
  if (TipoDeCupom = tcNFCE) then
  begin
    CheckConnect := TThreadCheckConnection.GetInstance;
    with CheckConnect do
    begin
      OnStatusChange := OnInternetChange;
      Start;
    end;
  end
  else
    G1.Visible := False;
  if TipoDeCupom <> tcSAT then
    mmSat.Visible := False;

  // data e hora da ultima carga
  if sCarga_data <> '' then
    cpBarra.Panels.Items[11].Text := sCarga_data + ' às ' + scarga_hora
  else
    cpBarra.Panels.Items[11].Text := '';

  pn_tampa.visible := False;
  pn_tampa.Align := alClient;

  MostraFotoProduto(bBusca_foto_produto);
  MostraLogoMarca(bHabLogoMarca);

  TipoImprimeCupom := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSAO CUPOM', 'S');
  sImpressora_Venda_Simples := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSORA VENDA SIMPLES', '');
  bExibeCumponNFTela := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'Exibir em Tela', '0') = '1';

  bImpressaoCozinha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSORA ESPECIFICA', '0') = '1';
  sImpressora_Cozinha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSORA COZINHA', '');

  bImprimeCarne := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRIME CARNE', '0') = '1';
  sImpressora_carne := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSORA CARNE', '');

  bAtivaImpFicha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSAO FICHA', '0') = '1';
  sImpressora_Ficha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'IMPRESSORA FICHA', '');
  sCodigoSegurancaFicha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'SEGURANCA FICHA', '');

  Touch_Colunas := StrToInt(frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'Colunas Touch', '5'));
  Touch_Linhas := StrToInt(frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'Linhas Touch', '5'));
  Touch_linhas_Grupo := StrToInt(frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'Linhas Touch Grupo', '1'));
  Touch_Altura_Grupo := StrToInt(frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'Altura Touch Grupo', '86'));

  bExibeEstoqueConsulta := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'EXIBE ESTOQUE CONSULTA', '1') = '1';
  bExibePrecoConsulta := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'EXIBE PRECO VENDA CONSULTA', '1') = '1';
  bSenhaVendaSemEstoque := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'SOLICITA SENHA VENDA SEM ESTOQUE', '0') = '1';

  bVerificaProdutoVenda := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'VERIFICA PRODUTO ANTES DE VENDER', '0') = '1';

  CartaoNaoInterfere := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'CARTAO_CUPOM', '0') = '1';

  UtilizaSenha := frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'CONTROLE DE SENHA', '0') = '1';

  bIdentificarVendedor := LerParametro('INFORMAR_VENDEDOR', 'I', 0);
  iColunaCOnsultaProdutos := LerParametro('COLUNA_CONSULTA_PRODUTOS', 'I', 0);

  cpBarra.Panels[6].Visible := (TipoDeCupom = tcNFCE) or (TipoDeCupom = tcSAT) or (TipoDeCupom = tcPafECF);
  cpBarra.Panels[7].Visible := (TipoDeCupom = tcNFCE) or (TipoDeCupom = tcSAT) or (TipoDeCupom = tcPafECF);
  cpBarra.Panels[8].Visible := TipoDeCupom = tcNFCE;
  cpBarra.Panels[9].Visible := TipoDeCupom = tcNFCE;

  if frmModulo.qrconfigTIPO_ETIQUETA_BALANCA.AsString = 'P' then
    TipoEtiqueta := tePeso
  else
    TipoEtiqueta := teTotal;

  pnTotal.Visible := not AtivaTouch;
  colBotDiminuir.Visible := AtivaTouch;
  colBotAlmuntar.Visible := AtivaTouch;

  if frmModulo.Conexao_Servidor.Connected then
    cpBarra.Panels[5].Text := 'ON-LINE'
  else
    cpBarra.Panels[5].Text := 'OFF-LINE';


  if FileExists(ExtractFilePath(Application.ExeName) + 'img\fundo_venda.jpg') then
    imgFundo.Picture.LoadFromFile(ExtractFilePath(Application.ExeName) + 'img\fundo_venda.jpg');

  Carregando := False;

  Width := Screen.Width;
  Height := Screen.Height;

  mniAlterarQuantidadeItens.Enabled := bMuda_unitario;
  //mniAlterarValorUnitario.Enabled := bMuda_unitario;

end;


procedure TfrmVenda.ed_barraKeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['+', '-', '_', '='] then
    Key := #0
  else
  if (Key = #13) and (wspnlAlerta.Left > 0) then
    RegistrarProduto
  else
  if Key = #27 then
    Teclado.Visible := False;
end;

procedure TfrmVenda.RegistrarProduto(ACodProduto: String);
var
  bDados_balanca: Boolean;
  sx_barra: string;
  MemoryStream: TMemoryStream;
  Jpg: TJPEGImage;
  Bitmap: TBitmap;
  ext: string;
begin
  CtrlVolume := ACodProduto <> '';
  if not CtrlVolume then
  begin
    sx_barra := Trim(ed_barra.text);

    if sx_barra = '' then
    begin
      if not AtivaTouch then
        if ed_barra.Enabled then
          ed_barra.setfocus;
      Exit;
    end;

    if rBal_peso > 0 then
      bDados_balanca := True
    else
      bDados_balanca := False;

    if not bVenda then
    begin
      try
        // verificar o serial do ecf
        // Abrir o cupom fiscal
        // bloquear ao teclado

        if not Abre_Venda then
        begin
          bLanca_pre_venda := False;
          bLanca_OS := False;
          bLanca_Mesa := False;
          bLanca_Abastecimento := False;
          // Caso o comando de abertura de venda retornou false, abortar o processo;
          Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
          Exit;
        end;
      finally
        // desbloquear o teclado
      end;
    end;
    // fazer a busca do produto
    sProd_barra := ed_barra.text;
  end
  else
    sProd_barra := ACodProduto;

  iPesquisa_produto := 1;
  if frmVenda.Localizar_Produto(sProd_barra) then
  begin
    if (query.fieldbyname('usa_balanca').AsInteger = 1) and (bPesagemAutomatica) then
    begin
      ed_qtde.Value := RetornaPesoBalanca;
      rBal_peso := ed_qtde.Value;
      if ed_qtde.Value = 0 then
        Exit;
      bDados_balanca := True;
    end;
    if not (AtivaTouch) then begin
      if bBusca_foto_produto then begin
        if not query.FieldByName('fotobd').IsNull then
        begin
          MostraLogoMarca(False);
          MostraFotoProduto(True);
          MemoryStream := TMemoryStream.Create;
          Jpg := TJPEGImage.Create;
          TBlobField(query.FieldByName('fotobd')).savetostream(MemoryStream);
          MemoryStream.Position := OffsetMemoryStream;
          Jpg.LoadFromStream(MemoryStream);
          img_produto.Picture.Assign(Jpg);
          FreeAndNil(MemoryStream);
        end
        else begin
          img_produto.Picture := nil;
          MostraLogoMarca(True);
          MostraFotoProduto(False);
        end;
      end
      else begin
        MostraLogoMarca(True);
        MostraFotoProduto(False);
      end;
    end;
    if query.FieldByName('pizza').AsString = 'S' then
    begin
      strPizza := Pizza(query.FieldByName('nome').AsString, rProd_preco);
      if strPizza.Sabores <> '' then
      begin
        ed_unitario.Value := rProd_preco;
        rProd_total := strPizza.Valor;
        Registra_Item;
        Exit;
      end;
    end;
    // achou o produto
    // alimentar os edits da tela
    if bDados_balanca then
    begin
      rProd_qtde := rBal_peso;
      // se for pesado
      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);

      ed_qtde.value := rProd_qtde;
      ed_unitario.Value := rProd_preco;
    end
    else
    begin
      // se for normal
      if ed_qtde.Value <= 0 then
        ed_qtde.value := rProd_qtde
      else if ed_qtde.Value > 1 then
        rProd_qtde := ed_qtde.Value;
      ed_unitario.Value := rProd_preco;
      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
      ed_total_item.Value := rProd_total;
    end;
    // verificar se o sistema permite edicao dos edit's de valores
    if bMuda_qtde or bMuda_unitario or bMuda_total then
    begin
      Imprime_display(sProd_nome, clWhite, tiInfo);
      ed_qtde.ReadOnly := False;
      ed_unitario.ReadOnly := False;
      if bMuda_qtde then
        ed_qtde.ReadOnly := False
      else
        ed_qtde.ReadOnly := True;
      if bMuda_unitario then
        ed_unitario.ReadOnly := False
      else
        ed_unitario.ReadOnly := True;

      if bMuda_total then
        ed_total_item.ReadOnly := False
      else
        ed_total_item.ReadOnly := True;

      if bMuda_qtde then
      begin
        if not AtivaTouch then
          ed_qtde.setfocus;
        Exit;
      end;
      if bMuda_unitario then
      begin
        if not AtivaTouch then
          ed_unitario.setfocus;
        Exit;
      end;
      if bMuda_total then
      begin
        if not AtivaTouch then
          ed_total_item.setfocus;
        Exit;
      end;
    end
    else
    begin
      // verificar se existe algum valor zerado
      if (rProd_total <= 0) or (rProd_qtde <= 0) or (rProd_preco <= 0) then
      begin
        ed_unitario.ReadOnly := False;
        if not AtivaTouch then
          ed_unitario.SetFocus;
        Imprime_display('INFORME O PREÇO DE VENDA DO PRODUTO!', clRed, tiErro);
        Exit;
      end;
      ed_barra.text := '';
      Registra_Item;
    end;
  end
  else
  begin
    MessageBeep(16);
    wspnlAlerta.Left := 0;
    cpnlMensagem.SetFocus;
    FLastKeyPress := vk_return;
    PopupMenu := nil;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.CancelarItem1Click(Sender: TObject);
begin
  if bVenda then begin
    if iTotal_Itens > 0 then begin
      if SolicitaLiberacaoCancelarGerencial then begin
        if not AutorizaGerencial(taCancItem) then
          Exit;
      end else if bSenha_Cancel_Cupom then begin
        frmsenha_supervisor := tfrmsenha_supervisor.create(self);
        frmsenha_supervisor.CancelarItem := True;
        frmsenha_supervisor.showmodal;
        if not bSupervisor_autenticado then
          Exit;
      end;
      if TipoDeCupom = tcPafECF then begin
        if frmModulo.qrItensCANCELADO.AsInteger = 1 then begin
          Imprime_display('O ITEM SELICIONADO JA ESTA CANCELADO!', clRed, tiErro);
        end else
          Cancela_Item(Zerar(frmModulo.qrItensITEM.AsString, 3), 'C');
      end else begin
        if btnCancelaItem.Caption = 'Descancelar Item' then
          Cancela_Item(Zerar(frmModulo.qrItensITEM.AsString, 3), 'D')
        else
          Cancela_Item(Zerar(frmModulo.qrItensITEM.AsString, 3), 'C');
      end;
    end
    else begin
      Imprime_display('NÃO EXISTEM ITENS LANÇADOS!', clRed, tiErro);
      Sleep(1000);
      Imprime_display_anterior;
    end;
  end
  else begin
    Imprime_display('Não existe cupom aberto para cancelar!', clRed, tiErro);
    Sleep(1000);
    Imprime_display_anterior;
    Exit;
  end;
end;

procedure TfrmVenda.ed_barraEnter(Sender: TObject);
begin
  if (AtivaTouch) and (ExibeTecladoVirtual) then begin
    Teclado.KeyboardType := ktQWERTZ; // 'Standard';
    Teclado.Width := 736;
    Teclado.Left := RoundDiv((Screen.Width - Teclado.Width), 2);
    Teclado.Top := RoundDiv((Screen.Height - Teclado.Height), 2);
    Teclado.Visible := True;
  end;
end;

function TfrmVenda.RetornaPesoBalanca: Double;
var
  Tentativas: Integer;
  Peso: Double;
begin
  Peso := 0;
  if frmModulo.balanca.Modelo <> balNenhum then begin
    try
      frmModulo.balanca.ativar;
      frmModulo.balanca.Ativo := True;
      Tentativas := 0;
      while Peso <= 0 do begin
        Peso := frmModulo.balanca.LePeso(iBal_time);
        Imprime_display('Coloque o produto na Balança!', clYellow, tiPeso);
        Inc(Tentativas);
        if Tentativas = 5 then begin
          if Application.MessageBox('Não foi possivel obter o peso da balança em 5 tentativas, Tentar novamente?', 'Atenção!', MB_ICONQUESTION + MB_YESNO) = mrYes then begin
            Tentativas := 0;
          end
          else begin
            Peso := 0;
            Break;
          end;
        end
      end;
      frmModulo.balanca.Ativo := False;
      Application.ProcessMessages;
    except
      on E: Exception do begin
        Application.MessageBox(PWideChar('Ocorreu um erro ao tentar conectar-se a balança: ' + #13 + E.Message), 'Atenção!', MB_ICONINFORMATION);
        frmModulo.balanca.Ativo := False;
      end;
    end;
  end
  else begin
    if bPesagemAutomatica then
      Application.MessageBox('O produto informado utiliza balança e o pdv esta configurado para Pesagem Automática e não foi configurada uma balança para capturar o peso do produto!', 'Atenção!', MB_ICONINFORMATION)
    else
      Application.MessageBox('Não foi configurada uma balança para capturar o peso do produto!', 'Atenção!', MB_ICONINFORMATION)
  end;
  Result := Peso;
end;

procedure TfrmVenda.Retornar1Click(Sender: TObject);
begin
  if not AtivaTouch then
    if ed_barra.Enabled then
      ed_barra.setfocus;
end;

function TfrmVenda.RetornaSenhaImpressao: Integer;
begin
  with frmModulo do begin
    qrSenhaPDV.Close;
    if SenhaLocal then
      qrSenhaPDV.Connection := frmModulo.conexao;
    qrSenhaPDV.Open;
    Result := qrSenhaPDVSENHA.AsInteger;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Sair1Click(Sender: TObject);
begin
  Close;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.CancelarCupom1Click(Sender: TObject);
var
  str: string;
begin
  if bVenda then begin
    //Caso exista cupom aberto
    if SolicitaLiberacaoCancelarGerencial then begin
      if not AutorizaGerencial(taCancCupom) then
        Exit;
    end else if bSenha_Cancel_Cupom then begin
      frmsenha_supervisor := tfrmsenha_supervisor.create(self);
      frmsenha_supervisor.CancelarVenda := True;
      frmsenha_supervisor.showmodal;
      if not bSupervisor_autenticado then
        Exit;
    end;
    LancaRecebimento.ValorDesconto := 0;
    LancaRecebimento.ValorAcressimo := 0;
    Cancela_cupom_aberto;
    lbEstoque.Visible := False;
    ReaberturaCupomAtiva := '';
  end else begin
    Imprime_display('Não existe cupom aberto para cancelar!', clRed, tiErro);
    Sleep(1000);
    Imprime_display_anterior;
    Exit;
  end;

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ed_barraExit(Sender: TObject);
begin
  if (AtivaTouch) and (ExibeTecladoVirtual) then
    Teclado.Visible := False;
  // Desativa os comandos do Menu OPCOES
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ED_DESC_ACREEnter(Sender: TObject);
begin
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FecharCupom1Click(Sender: TObject);
var
  rValor_Temp: Real;   //qiqi
  i, icont: Integer;
  NomeArquivo, sCod_Cupom: string;
  bImprimeCupomCredito, bLanca_comprovante_crediario, bLanca_Comprovante_Prestacao: Boolean;
  sCOO_crediario, sGNF_Crediario: string;
  sCOO_Prestacao, sGNF_Prestacao, sGRG_Prestacao: string;
  sIdentificarVendedor: string;
  str: string;
  SR: TSearchRec;
  Origem, Destino: string;
  ConfDesc, Desc, Resto, Prestacao: Double;
  PTroco: Double;
  Codigo, divide: Integer;
  rSat: TSat;
  Obs: String;
  IndiceBMP: String;
  totFederal, totEstadual, totMunicipal:Double;
  ret:tRetTef;
begin
  frmModulo.qrfilial.Close;
  frmModulo.qrfilial.Open;
  frmModulo.qrconfig.Close;
  frmModulo.qrconfig.Open;
  if bVenda then
  begin
    lbEstoque.Visible := False;
    if (rTotal_Venda <= 0) and not (TemDevolucao) then
    begin
      Imprime_display('CUPOM ESTÁ ZERADO!', clRed, tiErro);
      Sleep(1000);
      Imprime_display_anterior;
      Exit;
    end;
    img_produto.Visible := False;
    if bLanca_pre_venda then
    begin
      if rPre_venda_desconto > 0 then
        rTotal_Desconto := rTotal_Desconto + rPre_venda_desconto
      else
      if rPre_venda_acrescimo > 0 then
        rTotal_Acrescimo := rTotal_Acrescimo + rPre_venda_acrescimo;
    end;
    RecebeCupomCredito.id := -1;
    RecebeCupomCredito.Cupom := '';
    RecebeCupomCredito.Valor := 0;
    cdsRecebimento.Close;
    cdsRecebimento.CreateDataSet;
    rvalor_total_convenio := 0;
    rvalor_total_cheque := 0;
    rvalor_total_crediario := 0;
    if (rTotal_Venda <= 0) then
    begin
      Application.CreateForm(TfrmRecebimentoEspecial, frmRecebimentoEspecial);
      frmRecebimentoEspecial.ShowModal;
      vRecebimento.Confirma := frmRecebimentoEspecial.Confirma;
      FreeAndNil(frmRecebimentoEspecial);
      vRecebimento.ValorPago := rTotal_Venda;
      vRecebimento.Troco := 0;
      vRecebimento.Desconto := 0;
      vRecebimento.Acrescimo := 0;
      vRecebimento.CancelaImpressaoCozinha := True;
    end
    else
    begin
      AplicaDescontoPorGrupo;
      vRecebimento := EfetuaRecebimento(rTotal_Venda, rTotal_Couvert, rTotal_Comissao, rTotal_Desconto, rTotal_Acrescimo);
    end;
    /////////////////////////////////////////////////////////////////////////////////////////
    if vRecebimento.Confirma then
    begin
      bImprimeCupomCredito := False;
      if TipoDeCupom <> tcPafECF then
      begin
        // DESCONTO
        rTotal_Desconto := rTotal_Desconto + vRecebimento.Desconto;
        ConfDesc := 0;
        with frmModulo do
        begin
          try
            qrItens.DisableControls;
            qrItens.First;
            while not qrItens.Eof do
            begin
              qrItens.Edit;
              if rTotal_Venda = 0 then
                Desc := RoundTo((qrItensVALOR_TOTAL.AsFloat / 1) * vRecebimento.Desconto, -2)
              else
                Desc := RoundTo((qrItensVALOR_TOTAL.AsFloat / rTotal_Venda) * vRecebimento.Desconto, -2);

              ConfDesc := RoundTo(ConfDesc + Desc, -2);
              qrItensVALOR_DESCONTO.AsFloat := RoundTo(qrItensVALOR_DESCONTO.AsFloat + Desc, -2);
              qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat - Desc, -2);
              qrItens.Post;
              qrItens.Next;
            end;
          finally
            if (ConfDesc <> vRecebimento.Desconto) and (ConfDesc <> 0) then
            begin
              ConfDesc := rTotal_Desconto - ConfDesc;
              qrItens.Edit;
              qrItensVALOR_DESCONTO.AsFloat := qrItensVALOR_DESCONTO.AsFloat + ConfDesc;
              qrItensVALOR_TOTAL.AsFloat := qrItensVALOR_TOTAL.AsFloat + ConfDesc;
              qrItens.Post;
            end;
            qrItens.EnableControls;
          end;
        end;
        rTotal_Venda := rTotal_Venda - vRecebimento.Desconto;
        // ACRESCIMO
        rTotal_Acrescimo := rTotal_Acrescimo + vRecebimento.Acrescimo;
        ConfDesc := 0;
        with frmModulo do
        begin
          try
            qrItens.DisableControls;
            qrItens.First;
            while not qrItens.Eof do
            begin
              qrItens.Edit;
              if rTotal_Venda = 0 then
                Desc := RoundTo((qrItensVALOR_TOTAL.AsFloat / 1) * vRecebimento.Desconto, -2)
              else
                Desc := RoundTo((qrItensVALOR_TOTAL.AsFloat / rTotal_Venda) * vRecebimento.Desconto, -2);
              ConfDesc := RoundTo(ConfDesc + Desc, -2);
              qrItensVALOR_ACRESCIMO.AsFloat := RoundTo(qrItensVALOR_ACRESCIMO.AsFloat + Desc, -2);
              qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat + Desc, -2);
              qrItens.Post;
              qrItens.Next;
            end;
          finally
            if (ConfDesc <> rTotal_Acrescimo) and (ConfDesc <> 0) then
            begin
              ConfDesc := rTotal_Acrescimo - ConfDesc;
              qrItens.Edit;
              qrItensVALOR_ACRESCIMO.AsFloat := qrItensVALOR_ACRESCIMO.AsFloat + ConfDesc;
              qrItensVALOR_TOTAL.AsFloat := qrItensVALOR_TOTAL.AsFloat + ConfDesc;
              qrItens.Post;
            end;
            qrItens.EnableControls;
          end;
        end;
        rTotal_Venda := rTotal_Venda + vRecebimento.Acrescimo;
      end
      else
      begin
        if vRecebimento.Acrescimo > 0 then
          cpACBrECF.SubtotalizaCupom(vRecebimento.Acrescimo, '')
        else
        if vRecebimento.Desconto > 0 then
          cpACBrECF.SubtotalizaCupom(-vRecebimento.Desconto, '')
        else
          cpACBrECF.SubtotalizaCupom;
        cdsRecebimento.First;
        while not cdsRecebimento.Eof do
        begin
          if cdsRecebimentoTipo.AsString = StRecDinheiro then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecDinheiro, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          if (cdsRecebimentoTipo.AsString = StRecConvenio)  then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecConvenio, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          if  (cdsRecebimentoTipo.AsString = StRecCrediario) then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecCrediario, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          if cdsRecebimentoTipo.AsString = StRecCartaoCR then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecCartaoCR, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          if cdsRecebimentoTipo.AsString = StRecCartaoDE then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecCartaoDE, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          if cdsRecebimentoTipo.AsString = StRecCheque then
          begin
            cpACBrECF.EfetuaPagamento(vEcfRecCheque, cdsRecebimentovalor.AsFloat,'',False);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
          end;
          cdsRecebimento.Next;
        end;

        with frmModulo do
        begin
          qrItens.First;
          while not qrItens.Eof do
          begin
            totFederal := totFederal + RoundTo((qrItensVALOR_TOTAL.AsFloat * qrEstoqueALIQNACIONAL.AsFloat) / 100, -2);
            totEstadual := totEstadual + RoundTo((qrItensVALOR_TOTAL.AsFloat * qrEstoqueALIQESTADUAL.AsFloat) / 100, -2);
            totMunicipal := totMunicipal + RoundTo((qrItensVALOR_TOTAL.AsFloat * qrEstoqueALIQMUNICIPAL.AsFloat) / 100, -2);
            qrItens.Next;
          end;
        end;

        Obs := 'Componentes ACBr|http://acbr.sourceforge.net';
        IndiceBMP := '0';
        cpACBrECF.InfoRodapeCupom.MD5 := sMD5;
        cpACBrECF.InfoRodapeCupom.Dav := sNumero_Cupom;
        cpACBrECF.InfoRodapeCupom.MinasLegal := (frmModulo.qrfilialUF.AsString = 'MG');
        cpACBrECF.InfoRodapeCupom.ParaibaLegal := (frmModulo.qrfilialUF.AsString = 'PB');
        if cpACBrECF.InfoRodapeCupom.NotaLegalDF.Imprimir then
        begin
          cpACBrECF.InfoRodapeCupom.NotaLegalDF.ProgramaDeCredito := true;
          cpACBrECF.InfoRodapeCupom.NotaLegalDF.ValorICMS := 0;
          cpACBrECF.InfoRodapeCupom.NotaLegalDF.ValorISS := 0;
        end;

        cpACBrECF.InfoRodapeCupom.Imposto.ValorAproximadoFederal := totFederal; // informar o valor aproximado calculado a partir dos itens
        cpACBrECF.InfoRodapeCupom.Imposto.ValorAproximadoEstadual := totEstadual; // informar o valor aproximado calculado a partir dos itens
        cpACBrECF.InfoRodapeCupom.Imposto.ValorAproximadoMunicipal := totMunicipal; // informar o valor aproximado calculado a partir dos itens
        cpACBrECF.InfoRodapeCupom.Imposto.Fonte := 'IBPT/FECOMERCIO (aWd7S8)';
        cpACBrECF.InfoRodapeCupom.Imposto.ModoCompacto := true; //Faz com que a impressão seja em apenas duas linhas;


        cdsRecebimento.First;
        cdsRecebimento.Filter := 'TIPO='+QuotedStr('TEF');
        cdsRecebimento.Filtered := True;
        if cdsRecebimento.RecordCount = 0 then
        begin
          cpACBrECF.FechaCupom('', StrToIntDef(IndiceBMP, 0));
        end
        else
        begin
          while not cdsRecebimento.Eof do
          begin
            ret := frmTEF.RecebeTEF(cdsRecebimentovalor.AsFloat);
            TotalRecebidoTef := TotalRecebidoTef + cdsRecebimentovalor.AsFloat;
            if ret.Confirmado then
            begin
              cdsRecebimento.Edit;
              cdsRecebimentoTipo.AsString := ret.TipoDeCartcao;
              cdsRecebimentoParcelas.AsInteger := ret.Parcelas;
              cdsRecebimentotef_codigo_autorizacao.AsString := ret.CodigoAutorizacao;
              cdsRecebimento.Post;
            end;
            cdsRecebimento.Next;
          end;
        end;
      end;

      CancelaIMpressaoCozinha := vRecebimento.CancelaImpressaoCozinha;

      sMensagem_Cupom.Lines.Clear;
      bFinalizado := False;
      try
        if bIdentificarVendedor = 2 then
        begin
          Application.CreateForm(TfrmVendedor, frmVendedor);
          try
            if frmVendedor.ShowModal = mrOk then
            begin
              iVendedorCodigo := StrToInt(frmModulo.Query.FieldByName('CODIGO').AsString);
              sVendedorNome := frmModulo.Query.FieldByName('NOME').AsString;
            end
            else
            begin
              Exit;
            end
          finally
            FreeAndNil(frmVendedor);
          end;
        end;
        if iPre_venda_cliente > 0 then
        begin
          sCli_codigo := StrZero(iPre_venda_cliente,6);
        end;

        if (rvalor_total_crediario = 0) and (rvalor_total_convenio = 0) and (iPre_venda_cliente = 0) then
        begin
          if sConsumidor_CPF <> '' then
          begin
            if sConsumidor_Nome = '' then
            begin
              sConsumidor_Nome := 'Consumidor Final';
              sCli_Nome := 'Consumidor Final';
            end;
          end
          else
          begin
            sConsumidor_Nome := 'Consumidor Final';
            sCli_Nome := 'Consumidor Final';
            sCli_codigo := '000001';
          end;
        end;
        Application.ProcessMessages;
        (*************** F E C H A M E N T O   D O   C U P O M ********************)
        if bMensagem_Cupom then
          RetTexto(sMensagem_Cupom);
        if bIdentificarVendedor > 0 then
          sIdentificarVendedor := 'Vendedor ' + texto_justifica(FormatFloat('000', iVendedorCodigo) + ' - ' + sVendedorNome, 37, ' ', taEsquerda) + #10
        else
          sIdentificarVendedor := '';
        if UtilizaSenha then
          SenhaImpressao := RetornaSenhaImpressao
        else
          SenhaImpressao := 0;
        //////////////////////////////////////////////////////////////////
        ///  FECHAMENTO SAT
        if (frmPrincipal.TipoImpressora = SemImpressora) and (TipoDeCupom = tcSAT) then
        begin
          rSat := EnviaSAT(sCod_Cupom);
          if not rSat.Validou then
          begin
            rTotal_Desconto := rTotal_Desconto - vRecebimento.Desconto;
            rTotal_Venda := rTotal_Venda + vRecebimento.Desconto;
            RemoveDescontoTabelaItem;
            if (TEF_Ativo) and (RecebeuTEF) then
              frmTEF.ACBrTEFD1.CancelarTransacoesPendentes;
            Application.MessageBox(PWideChar('Erro ao enviar o SAT:' + #13 + 'Código: ' + IntToStr(rSat.Codigo) + ' - ' + rSat.Mesagem), 'Atenção!', MB_ICONINFORMATION);
            Exit;
          end;
        end
        else
        //////////////////////////////////////////////////////////////////
        ///  FECHAMENTO NFCE
        if (frmPrincipal.TipoImpressora = SemImpressora) and (TipoDeCupom = tcNFCE) then
        begin
          Imprime_display('AGUARDE...  PREPARANDO NFC-E', clWhite, tiLivre);
          if not PrepararNFCE then
            Exit;
          try
            Imprime_display('AGUARDE...  GRAVANDO NFC-E NO BANCO', clWhite, tiLivre);
            if nfce_autorizada then begin
              vgerado_nfce := 'S';
              with frmModulo do begin
                NomeArquivo := ACBRNFCe.Configuracoes.Arquivos.PathNFe;
                if ACBRNFCe.Configuracoes.Arquivos.SepararPorCNPJ then
                  NomeArquivo := NomeArquivo + '\' + edtEmitCNPJ;
                if ACBRNFCe.Configuracoes.Arquivos.SepararPorModelo then
                  NomeArquivo := NomeArquivo + '\NFCe';
                if ACBRNFCe.Configuracoes.Arquivos.SepararPorMes then
                  NomeArquivo := NomeArquivo + '\' + FormatDateTime('YYYYMM', Date);
                if ACBRNFCe.Configuracoes.Arquivos.AdicionarLiteral then
                  NomeArquivo := NomeArquivo + '\NFCe';
                NomeArquivo := NomeArquivo + '\' + Copy(ChaveNFCE, 4, 47) + '-nfe.xml';

                qrNFCEInsert.Open;
                qrNFCEInsert.Insert;
                qrNFCEInsertNUMERO.asinteger := NumeroNFCe;
                qrNFCEInsertSERIE.asinteger := StrToInt(edtSerie);
                if ACBRNFCe.Configuracoes.WebServices.Ambiente  = taProducao then
                  qrNFCEInsertAMBIENTE.AsString := 'P'
                else
                  qrNFCEInsertAMBIENTE.AsString := 'H';
                qrNFCEInsertDATA.AsDateTime := Date;
                qrNFCEInsertHORA.AsString := FormatDateTime('HH:MM:SS', Time);
                qrNFCEInsertTOTAL.AsFloat := rTotal_Venda;
                qrNFCEInsertCLIENTE.AsString := sCli_Nome;
                qrNFCEInsertCHAVE.AsString := Copy(ChaveNFCE, 4, 47);
                qrNFCEInsertXML.asstring := NomeArquivo;
                qrNFCEInsertSITUACAO.AsInteger := 0;
                qrNFCEInsertTROCO.AsFloat := vRecebimento.Troco;
                if frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao = frmtOffLine then
                begin
                  vcontingencia := 'S';
                  qrNFCEInsertCONTINGENCIA.AsString := 'S';
                  qrNFCEInsertMOTIVOCONTIGENCIA.AsString := MotivoContigencia;
                  qrNFCEInsertDATAHORACONTIMGENCIA.AsDateTime := DataHoraContingencia;
                  if ContingenciaPorErroNoEnvio then
                  begin
                    qrNFCEInsertSTATUS.AsString := 'E';
                    qrNFCEInsertERRO.AsString := ErroNoEnvioContingencia;
                  end
                  else
                    qrNFCEInsertSTATUS.AsString := 'V';
                end
                else
                begin
                  qrNFCEInsertSTATUS.AsString := 'V';
                  vcontingencia := 'N';
                  qrNFCEInsertCONTINGENCIA.AsString := 'N';
                  qrNFCEInsertMOTIVOCONTIGENCIA.AsString := '';
                end;
                qrNFCEInsertENVIADOCONTINGENCIA.AsString := 'N';
                qrNFCEInsertXMLENVIO.LoadFromFile(NomeArquivo);
                qrNFCEInsertEX.AsInteger := 0;

                qrNFCEInsertALTERADO.AsString := 'N';
                qrNFCEInsertXMLCACNELAMENTO.Clear;
                qrNFCEInsertCUPOM.AsString := sNumero_Venda;
                qrNFCEInsert.Post;
                qrNFCEInsert.ApplyUpdates;
              end;
            end
            else
            begin
              Imprime_display('OCORREU UM ERRO AO ENVIAR A NFC-E', clWhite, tiLivre);
              Exit;
            end;
          except
            on E: Exception do
            begin
              rTotal_Desconto := rTotal_Desconto - vRecebimento.Desconto;
              rTotal_Venda := rTotal_Venda + vRecebimento.Desconto;
              RemoveDescontoTabelaItem;
              Imprime_display('ERRO NFCE: ' + E.Message, clWhite, tiLivre);
              if (TEF_Ativo) and (RecebeuTEF) then
                 frmTEF.ACBrTEFD1.CancelarTransacoesPendentes;
              Exit;
            end;
          end;
        end
        else
        //////////////////////////////////////////////////////////////////
        begin
          vgerado_nfce := 'N';
          vcontingencia := 'N';
        end;

        // identificacao do consumidor no cupom
        if sCli_Nome <> '' then
        begin
          bFinalizado := True;
          frmModulo.spCupom_Temp_Edit.close;
          frmModulo.spCupom_Temp_Edit.ParamByName('procedimento').asstring := 'FINALIZADO';
          frmModulo.spCupom_Temp_Edit.Prepare;
          frmModulo.spCupom_Temp_Edit.Execute;
          sPre_Venda_Numero := '';
          sDav_numero := '';
          sPosto_rodape := '';
        end
        else
        begin
          bFinalizado := True;
          frmModulo.spCupom_Temp_Edit.Close;
          frmModulo.spCupom_Temp_Edit.ParamByName('procedimento').asstring := 'FINALIZADO';
          frmModulo.spCupom_Temp_Edit.Prepare;
          frmModulo.spCupom_Temp_Edit.execute;
          // sem consumidor
        end;
        Application.ProcessMessages;
        bFinalizado := True;
        // atualizar o totalizador geral do PAF com a do ECF
        //---> FIM
        (**************************************************************************)
        //--> (C U P O M    N A O   F I S C A L )
        bLanca_comprovante_crediario := False;
        // comprovante do crediario
        if (bCadastra_Crediario) and (rvalor_total_crediario > 0) and (iComprovente_Crediario = 1) then
        begin
          Imprime_display('Aguarde! Imprimindo Comprovante...', clWhite, tiInfo);
          bLanca_comprovante_crediario := False;
          sCOO_Prestacao := sNumero_Cupom;
          sGNF_Prestacao := sGNF;
          sGRG_Prestacao := sGRG;
          bLanca_Comprovante_Prestacao := True;
        end;
        if bFinalizado then begin
          Imprime_display('Aguarde!Finalizando Venda!', clWhite, tiInfo);
          // cadastrar cheques
          if (bCadastra_Cheque) and (rvalor_total_cheque > 0) then
          begin
          end;
          // lancar convenio
          if (bCadastra_Convenio) and (rvalor_total_convenio > 0) then
          begin
          end;

          with frmModulo do
          begin
            // lancamento do cupom no banco de dados
            sNumero_Cupom := sNumero_Venda;
            lbCupom.Caption := 'Útimo Cupom: ' + sNumero_Cupom;
            sCod_Cupom := codifica_cupom;
            qrGrava_Cupom.close;
            qrGrava_Cupom.Open;
            qrGrava_Cupom.Insert;
            qrGrava_Cupomcodigo.asstring := sCod_Cupom;
            if (vNumNFCe > 0) and (TipoDeCupom = tcNFCE) then
              qrGrava_Cupomnumero_nfce.asstring := IntToStr(vNumNFCe)
            else
              qrGrava_Cupomnumero_nfce.Clear;

            qrGrava_Cupomnumero.asstring := sNumero_Cupom;
            qrGrava_Cupomccf.asstring := sNumero_contador_cupom;
            qrGrava_CupomECF.AsString := sCaixa;
            qrGrava_Cupomdata.AsDateTime := dData_Sistema;
            qrGrava_Cupomhora.AsDateTime := hHora_Cupom;
            qrGrava_Cupomqtde_item.asinteger := iTotal_Itens;
            qrGrava_Cupomvalor_produto.asfloat := rTotal_Venda + rTotal_Desconto - rTotal_Acrescimo;
            qrGrava_CupomValor_Desconto.asfloat := rTotal_Desconto;
            qrGrava_CupomValor_Acrescimo.asfloat := rTotal_Acrescimo;
            qrGrava_Cupomvalor_total.asfloat := rTotal_Venda;
            qrGrava_Cupomvalor_pago.asfloat := vRecebimento.ValorPago;
            qrGrava_Cupomvalor_troco.asfloat := vRecebimento.Troco;
            if sCli_codigo <> '' then
              qrGrava_Cupomcod_cliente.asstring := sCli_codigo
            else
              qrGrava_Cupomcod_cliente.asstring := '000001';
            qrGrava_Cupomcancelado.asinteger := 0;
            qrGrava_Cupomcpf_consumidor.asstring := sConsumidor_CPF;
            qrGrava_Cupomnome_consumidor.asstring := Copy(sConsumidor_Nome, 1, 40);
            qrGrava_Cupomcod_caixa.asinteger := iNumCaixa;
            qrGrava_Cupomecf_caixa.asstring := Copy(sCaixa, 1, 3);
            if iVendedorCodigo <= 0 then
              iVendedorCodigo := icodigo_Usuario;
            qrGrava_Cupomcod_vendedor.asinteger := iVendedorCodigo;
            qrGrava_Cupomcontingencia.asstring := vcontingencia;
            qrGrava_Cupomgerado_nfce.asstring := vgerado_nfce;
            qrGrava_CupomPEGO_PORQUEM.asstring := sCLi_PegoPor;
            qrGrava_CupomSAT.AsInteger := 1;
            if frmPrincipal.TipoImpressora = SemImpressora then
            begin
              qrGrava_CupomNAOFISC.asstring := 'N';
            end
            else
            begin
              qrGrava_CupomNAOFISC.asstring := 'S';
            end;
            qrGrava_Cupomcouvert.AsFloat := rTotal_Couvert;
            qrGrava_Cupomtaxa_garcon.AsFloat := rTotal_Comissao;
            if bImpressaoCozinha then
              qrGrava_Cupomimpresso.AsString := 'N'
            else
              qrGrava_Cupomimpresso.AsString := 'S';
            qrGrava_CupomSENHA.AsInteger := SenhaImpressao;
            qrGrava_Cupomendereco_consumidor.AsString := Copy(sConsumidor_Endereco, 1, 130);

            if (TipoDeCupom = tcSAT) and (frmPrincipal.TipoImpressora = SemImpressora) then
            begin
              qrGrava_CupomNUMERO_SAT.AsInteger := rSat.Numero;
              qrGrava_CupomGEROU_SAT.AsString := 'S';
            end
            else
            begin
              qrGrava_CupomNUMERO_SAT.AsInteger := 0;
              qrGrava_CupomGEROU_SAT.AsString := 'N';
            end;
            qrGrava_CupomEX.AsInteger := 0;
            qrGrava_Cupom.Post;

            if bLanca_pre_venda then
            begin
              // atualizar situacao da pre-venda para 2 = concluida
              frmModulo.query_servidor.close;
              frmModulo.query_servidor.sql.clear;
              frmModulo.query_servidor.sql.add('update c000074 set situacao = 2 where codigo = ''' + Zerar(IntToStr(iPre_venda_codigo), 6) + '''');
              frmModulo.query_servidor.ExecSQL;

              frmModulo.query_servidor.close;
              frmModulo.query_servidor.sql.clear;
              frmModulo.query_servidor.sql.add('update c000075 set situacao = 2 where codnota = ''' + Zerar(IntToStr(iPre_venda_codigo), 6) + '''');
              frmModulo.query_servidor.ExecSQL;

              bLanca_pre_venda := False;
            end;
            if bLanca_OS then begin
              // atualizar a situacao da os para FECHADA
              frmModulo.query_servidor.close;
              frmModulo.query_servidor.sql.clear;
              frmModulo.query_servidor.sql.add('update c000051 set situacao = ''FECHADA'' where codigo = ''' + Zerar(IntToStr(iOS_codigo), 6) + '''');
              frmModulo.query_servidor.ExecSQL;
              bLanca_OS := False;
            end;
            if bLanca_Mesa or bLanca_Comanda then
            begin
              // excluir a mesa e os itens

              frmModulo.query_servidor.close;
              frmModulo.query_servidor.sql.clear;
              frmModulo.query_servidor.sql.add('select * from r000001 where codigo in (' + MesasComandasReceber + ')');
              frmModulo.query_servidor.Open;

              frmModulo.query_servidor.First;
              while not frmModulo.query_servidor.Eof do
              begin

                frmModulo.qrFechaMesas.sql.clear;
                frmModulo.qrFechaMesas.sql.add('insert into MESAS_FECHADAS (data_abertura, hora_abertura, data_fechamento, hora_fechamento, tempo_uso, COD_FUNCIONARIO, num_mesa, VALOR_FECHAMENTO, VALOR_COMISSAO, CLIENTE_AUTORIZOU, comanda,atb)');
                frmModulo.qrFechaMesas.sql.add('values (:data_abertura, :hora_abertura, :data_fechamento, :hora_fechamento, :tempo_uso, :COD_FUNCIONARIO, :num_mesa, :VALOR_FECHAMENTO, :VALOR_COMISSAO, :CLIENTE_AUTORIZOU, :comanda,:atb)');
                frmModulo.qrFechaMesas.ParamByName('data_abertura').AsDateTime := frmModulo.query_servidor.FieldByName('DATA').AsDateTime;
                frmModulo.qrFechaMesas.ParamByName('hora_abertura').AsString := frmModulo.query_servidor.FieldByName('HORA').AsString;
                frmModulo.qrFechaMesas.ParamByName('data_fechamento').AsDate := Date;
                frmModulo.qrFechaMesas.ParamByName('hora_fechamento').AsString := FormatDateTime('HH:MM:SS', Time);
                frmModulo.qrFechaMesas.ParamByName('tempo_uso').AsString := FormatDateTime('HH:MM:SS', Time - StrToTime(frmModulo.query_servidor.FieldByName('HORA').AsString));
                frmModulo.qrFechaMesas.ParamByName('COD_FUNCIONARIO').AsString := frmModulo.query_servidor.FieldByName('COD_FUNCIONARIO').AsString;
                frmModulo.qrFechaMesas.ParamByName('num_mesa').AsInteger := frmModulo.query_servidor.FieldByName('mesa').AsInteger;
                frmModulo.qrFechaMesas.ParamByName('comanda').AsInteger := frmModulo.query_servidor.FieldByName('comanda').AsInteger;
                frmModulo.qrFechaMesas.ParamByName('VALOR_FECHAMENTO').AsFloat := rTotal_Venda;
                frmModulo.qrFechaMesas.ParamByName('VALOR_COMISSAO').AsFloat := rTotal_Comissao;
                if rTotal_Comissao > 0 then
                  frmModulo.qrFechaMesas.ParamByName('CLIENTE_AUTORIZOU').AsString := 'S'
                else
                  frmModulo.qrFechaMesas.ParamByName('CLIENTE_AUTORIZOU').AsString := 'N';
                frmModulo.qrFechaMesas.ParamByName('atb').AsString := ME_GravaATB('TB_MESAS_FECHADAS');
                frmModulo.qrFechaMesas.ExecSQL;

                frmModulo.query_servidor2.close;
                frmModulo.query_servidor2.sql.clear;
                frmModulo.query_servidor2.sql.add('delete from r000002 where cod_mesa =' + IntToStr(frmModulo.query_servidor.FieldByName('codigo').AsInteger));
                frmModulo.query_servidor2.ExecSQL;

                frmModulo.query_servidor2.close;
                frmModulo.query_servidor2.sql.clear;
                frmModulo.query_servidor2.sql.add('delete from r000001 where codigo =' + IntToStr(frmModulo.query_servidor.FieldByName('codigo').AsInteger));
                frmModulo.query_servidor2.ExecSQL;
                frmModulo.query_servidor.Next;
              end;
            end;

            if bLanca_Delivery then
            begin
              // excluir a mesa e os itens
              frmModulo.query_servidor.close;
              frmModulo.query_servidor.sql.clear;
              frmModulo.query_servidor.sql.add('update DELIVERY set RECEBIDO = ' + QuotedStr('S') + ' where id = ' + IntToStr(iID_Delivery));
              frmModulo.query_servidor.ExecSQL;
            end;

            // lancamento dos items do cupom no banco de dados
            qrItens.First;
            while not qrItens.Eof do
            begin
              // situacoes
              // 1 - Venda do Item
              // 3 - Cancelamento do Item
              // 4 - Desconto no Item
              // 5 - Acrescimo no Item

              if qrItensCANCELADO.AsInteger = 0 then
              begin
                with frmModulo do
                begin
                  qrEstoque.Close;
                  qrEstoque.ParamByName('cod').AsInteger := qrItensCOD_PRODUTO.AsInteger;
                  qrEstoque.Open;
                end;
                qrGravaItem.close;
                qrGravaItem.Open;
                qrGravaItem.Insert;
                qrGravaItemcodigo.AsString := codifica_item(qrItensITEM.AsInteger);
                qrGravaItemcod_cupom.asstring := sCod_Cupom;
                qrGravaItemitem.asinteger := qrItensITEM.AsInteger;
                qrGravaItemcod_produto.asinteger := qrItensCOD_PRODUTO.AsInteger;
                qrGravaItemunidade.asstring := qrItensUNIDADE.AsString;
                qrGravaItemqtde.asfloat := qrItensQTDE.AsFloat;
                qrGravaItemvalor_unitario.asfloat := qrItensVALOR_UNITARIO.AsFloat;
                qrGravaItemvalor_subtotal.asfloat := qrItensVALOR_SUBTOTAL.AsFloat;
                qrGravaItemvalor_desconto.asfloat := qrItensVALOR_DESCONTO.asfloat;
                qrGravaItemvalor_acrescimo.asfloat := qrItensVALOR_ACRESCIMO.asfloat;
                qrGravaItemvalor_total.asfloat := qrItensVALOR_TOTAL.asfloat;
                qrGravaItemcst.asstring := qrItensCST.asstring;
                qrGravaItemaliquota.asfloat := qrItensALIQUOTA.asfloat;
                qrGravaItemcancelado.asinteger := qrItensCANCELADO.asinteger;
                qrGravaItemtamanho.AsString := qrItensTAMANHO.AsString;
                qrGravaItemcor.AsString := qrItensCOR.AsString;
                qrGravaItemcsosn.AsString := qrEstoqueCSOSN.AsString;
                qrGravaItemclassificacao_fiscal.AsString := qrEstoqueNCM.AsString;
                qrGravaItemBASE_SUB.AsFloat := 0;
                qrGravaItemICMS_SUB.AsFloat := 0;
                qrGravaItemisentas_icms.AsFloat := 0;
                qrGravaItemoutras_icms.AsFloat := 0;
                qrGravaItemcod_totalizador.asstring := '';
                if Trim(qrEstoqueCFOP.AsString) <> '' then
                  qrGravaItemcfop.AsString := Trim(qrEstoqueCFOP.AsString)
                else
                  qrGravaItemcfop.AsString := edtcfop;

                qrGravaItemextra.AsString := qrItensEXTRA.AsString;
                qrGravaItemCOMPLEMENTO.asstring := qrItensCOMPLEMENTO.AsString;
                qrGravaItemEX.AsInteger := 0;
                qrGravaItem.Post;

                cdsLancaProdutos.Filter := 'COD_ITEM = ' + QuotedStr(qrItensCODIGO.AsString);
                cdsLancaProdutos.Filtered := True;
                cdsLancaProdutos.First;
                while not cdsLancaProdutos.Eof do
                begin
                  qrLancaProdutos.Close;
                  qrLancaProdutos.Open;
                  qrLancaProdutos.Append;
                  qrLancaProdutosCOD_ITEM.AsString := qrItensCODIGO.AsString;
                  qrLancaProdutosCODPRO.AsString := cdsLancaProdutosCODPRO.AsString;
                  qrLancaProdutosQTDE.AsFloat := cdsLancaProdutosQTDE.AsFloat;
                  qrLancaProdutos.Post;
                  qrLancaProdutos.ApplyUpdates;
                  cdsLancaProdutos.Next;
                end;
              end;
              qrItens.Next;
            end;
            // lancamento das formas de pagamento
            PTroco := 0;
            divide := 0;
            if vRecebimento.Troco > 0 then
            begin
              cdsRecebimento.First;
              while not cdsRecebimento.Eof do
              begin
                Inc(divide);
                cdsRecebimento.Next;
              end;
              PTroco := (vRecebimento.Troco / divide);
            end;
            cdsRecebimento.First;       //qiqi
            while not cdsRecebimento.Eof do
            begin
              if cdsRecebimentoTipo.AsString <> StRecCredito then
              begin
                qrGravaForma.close;
                qrGravaForma.Open;
                qrGravaForma.Insert;
                qrGravaFormacodigo.asstring := codifica_forma(cdsRecebimento.RecNo);
                qrGravaFormacod_cupom.asstring := sCod_Cupom;
                qrGravaFormaforma.asstring := cdsRecebimentoTipo.AsString;
                qrGravaFormavalor.asfloat := cdsRecebimentovalor.AsFloat;
                if ((cdsRecebimentoTipo.AsString = StRecCartaoDE) or (cdsRecebimentoTipo.AsString = StRecCartaoCR) or (cdsRecebimentoTipo.AsString = StRecCrediario) or (cdsRecebimentoTipo.AsString = StRecConvenio)) then
                  qrGravaFormatroco.asfloat := 0
                else
                  qrGravaFormatroco.asfloat := PTroco;
                if cdsRecebimentobanco.AsInteger <= 0 then
                begin
                  qrGravaFormaprestacao.asinteger := 1;
                  qrGravaFormabanco.asstring := '';
                end
                else begin
                  qrGravaFormaprestacao.asinteger := cdsRecebimentoParcelas.AsInteger;
                  qrGravaFormabanco.asstring := Zerar(IntToStr(cdsRecebimentobanco.AsInteger), 3);
                end;
                if (cdsRecebimentoTipo.AsString = StRecCrediario) then
                  qrGravaFormaprestacao.asinteger := iCrediario_prestacao;
                qrGravaFormatipo.asstring := 'CUPOM NAO FISCAL';
                if (cdsRecebimentoTipo.AsString = StRecCartaoDE) then
                  qrGravaFormatipocartao.asstring := 'D'
                else if (cdsRecebimentoTipo.AsString = StRecCartaoCR) then
                  qrGravaFormatipocartao.asstring := 'C'
                else
                  qrGravaFormatipocartao.asstring := '';
                qrGravaFormaEX.AsInteger := 0;
                if cdsRecebimentoCupomCreditoID.AsInteger > 0 then
                  qrGravaFormaID_CUPOM_CREDITO.AsInteger := cdsRecebimentoCupomCreditoID.AsInteger;
                qrGravaForma.Post;
               end;

              if (cdsRecebimentoTipo.AsString = StRecCartaoCR) or (cdsRecebimentoTipo.AsString = StRecCartaoDE) then
              begin
                Prestacao := RoundTo(cdsRecebimentovalor.AsFloat / cdsRecebimentoParcelas.AsInteger, -2);
                Resto := cdsRecebimentovalor.AsFloat - (Prestacao * cdsRecebimentoParcelas.AsInteger);
                for i := 1 to cdsRecebimentoParcelas.AsInteger do
                begin
                  qrCupom_Cartao.Open;
                  qrCupom_Cartao.Insert;
                  qrCupom_CartaoCODIGO.AsString := codifica_forma(i);
                  qrCupom_CartaoCODIGO_CUPOM.AsString := sCod_Cupom;
                  qrCupom_CartaoBANCO.AsString := Zerar(IntToStr(cdsRecebimentobanco.AsInteger), 3);
                  if (cdsRecebimentoTipo.AsString = StRecCartaoDE) then
                    qrCupom_CartaoTIPO_CARTAO.AsString := 'D'
                  else if (cdsRecebimentoTipo.AsString = StRecCartaoCR) then
                    qrCupom_CartaoTIPO_CARTAO.AsString := 'C'
                  else
                    qrCupom_CartaoTIPO_CARTAO.AsString := '';
                  qrCupom_CartaoPRESTACAO.AsInteger := i;
                  if i <> cdsRecebimentoParcelas.AsInteger then
                    qrCupom_CartaoVALOR.AsFloat := Prestacao
                  else
                    qrCupom_CartaoVALOR.AsFloat := Prestacao + Resto;
                  qrCupom_CartaoVENCIMENTO.asdatetime := IncMonth(dData_Sistema, i);
                  qrCupom_CartaoEX.AsString := 'N';
                  qrCupom_CartaoCODIGO_AUTORIZACAO.AsString := cdsRecebimentotef_codigo_autorizacao.AsString;
                  qrCupom_CartaoCNPJ.AsString := cdsRecebimentotef_cnpj.AsString;
                  qrCupom_Cartao.Post;
                  qrCupom_Cartao.ApplyUpdates;
                end;
              end;
              if cdsRecebimentoTipo.AsString = StRecCredito then
              begin
                qrSeqDevCredito.Close;
                qrSeqDevCredito.Open;
                cdsDevolucao.First;
                bImprimeCupomCredito := True;
                qrdevolucao_credito.Close;
                qrdevolucao_credito.Open;
                qrdevolucao_credito.Append;
                qrdevolucao_creditoID.AsInteger := qrSeqDevCreditoSEQ.AsInteger;
                qrdevolucao_creditoVENCIMENTO.AsDateTime := Date + qrconfigDIAS_VALIDADE_CUPOM_CREDITO.AsInteger;
                while not cdsDevolucao.Eof do
                begin
                  qrdevolucao_creditoVALOR.AsFloat := qrdevolucao_creditoVALOR.AsFloat + cdsDevolucaovalor.AsFloat;
                  cdsDevolucao.Next;
                end;
                qrdevolucao_creditoVALOR_UTILIZADO.AsFloat := 0;
                qrdevolucao_creditoVALOR_SALDO.AsFloat := qrdevolucao_creditoVALOR.AsFloat;
                qrdevolucao_credito.Post;
                qrdevolucao_credito.ApplyUpdates;
                cdsDevolucao.First;
                while not cdsDevolucao.Eof do
                begin
                  qrdevolucao.Close;
                  qrdevolucao.Open;
                  qrdevolucao.Append;
                  qrdevolucaoID_CREDITO.AsInteger := qrSeqDevCreditoSEQ.AsInteger;
                  qrdevolucaoDATA.AsDateTime := Date;
                  qrdevolucaoHORA.AsString := FormatDateTime('HH:MM:SS', Time);
                  qrdevolucaoCODCLIENTE.AsString := IntToStrZero(CodCliGeraCredito, 6);
                  qrdevolucaoCUPOM_DEVOLUCAO.AsString := sNumero_Cupom;
                  qrdevolucaoCUPOM_DEVOLVIDO.AsString := cdsDevolucaocodcupom.AsString;
                  qrdevolucaoITEM.AsString := IntToStrZero(cdsDevolucaoitem.AsInteger, 3);
                  qrdevolucaoVALOR.AsFloat := cdsDevolucaovalor.AsFloat;
                  qrdevolucao.Post;
                  qrdevolucao.ApplyUpdates;
                  qrServidor.SQL.Clear;
                  qrServidor.SQL.Add('update c000032 set devolvido = ' + QuotedStr('S') + ' where cupom_numero = ' + QuotedStr(cdsDevolucaocodcupom.AsString) + ' and CUPOM_ITEM = ' + QuotedStr(IntToStrZero(cdsDevolucaoitem.AsInteger, 3)));
                  qrServidor.ExecSQL;
                  cdsDevolucao.Next;
                end;
              end;
              cdsRecebimento.Next; //qiqi
            end;

            if (bCadastra_Convenio) and (rvalor_total_convenio > 0) then
            begin
              qrCupomConvenio.Open;
              qrCupomConvenio.Insert;
              qrCupomConvenioCODIGO.asstring := codifica_crediario(i);
              qrCupomConvenioCOD_CUPOM.asstring := sCod_Cupom;
              qrCupomConvenioDATA.AsDateTime := dData_Sistema;
              qrCupomConvenioHORA.AsString := FormatDateTime('HH:MM:SS', Time);
              qrCupomConvenioCOD_CLIENTE.asinteger := StrToInt(sCli_codigo);
              qrCupomConvenioVENCIMENTO.asdatetime := IncMonth(dData_Sistema, 1);
              qrCupomConvenioVALOR.asfloat := rvalor_total_convenio;
              qrCupomConvenioEX.AsInteger := 0;
              qrCupomConvenio.Post;
              qrCupomConvenio.ApplyUpdates;
            end;

            if (bCadastra_Crediario) and (rvalor_total_crediario > 0) then
            begin
              cdsCrediario.First;
              while not cdsCrediario.Eof do
              begin
                spCupom_Crediario.close;
                spCupom_Crediario.ParamByName('codigo').asstring := codifica_crediario(cdsCrediarioparcela.AsInteger);
                spCupom_Crediario.ParamByName('cod_cupom').asstring := sCod_Cupom;
                spCupom_Crediario.ParamByName('data').asdate := dData_Sistema;
                spCupom_Crediario.ParamByName('Hora').astime := Time;
                spCupom_Crediario.ParamByName('cod_cliente').asinteger := cdsCrediariocliente.AsInteger;
                spCupom_crediario.ParamByName('vencimento').asdatetime :=  cdsCrediariovencimento.AsDateTime;
                spCupom_Crediario.parambyname('valor').asfloat := cdsCrediariovalor.AsFloat;
                spCupom_Crediario.ParamByName('descricao').asstring := Copy(sCrediario_Nome, 1, 30);
                spCupom_Crediario.Prepare;
                spCupom_Crediario.Execute;
                cdsCrediario.Next;
              end;
            end;


            // lancamento do comprovante do crediario
            if bLanca_comprovante_crediario then
            begin
              sNumero_Cupom := Zerar(somenteNumero(sCOO_crediario), 6);
              if sNumero_Cupom <> '000000' then
              begin
                qrGravaNaoFiscal.close;
                qrGravaNaoFiscal.Open;
                qrGravaNaoFiscal.Insert;
                qrGravaNaoFiscalcodigo.asstring := codifica_cupom;
                qrGravaNaoFiscalecf.asstring := sCaixa;
                qrGravaNaoFiscaldata.asdatetime := dData_Sistema;
                qrGravaNaoFiscalindice.AsString := sTotalizador_Crediario;
                qrGravaNaoFiscaldescricao.AsString := sNome_Totalizador_Crediario;
                qrGravaNaoFiscalvalor.asfloat := rvalor_total_crediario;
                qrGravaNaoFiscalhora.AsDateTime := Time;
                qrGravaNaoFiscalCOO.asstring := sCOO_crediario;
                qrGravaNaoFiscalGNF.asstring := sGNF_Crediario;
                qrGravaNaoFiscalCDC.Clear;
                qrGravaNaoFiscalGRG.clear;
                qrGravaNaoFiscalDENOMINACAO.asstring := 'CN';
                qrGravaNaoFiscal.Post;
              end;
            end;
            // lancamento do comprovante da prestacao
            if bLanca_Comprovante_Prestacao then
            begin
              sNumero_Cupom := Zerar(somenteNumero(sCOO_Prestacao), 6);
              if sNumero_Cupom <> '000000' then
              begin
                qrGravaNaoFiscal.close;
                qrGravaNaoFiscal.Open;
                qrGravaNaoFiscal.Insert;
                qrGravaNaoFiscalcodigo.asstring := codifica_cupom;
                qrGravaNaoFiscalecf.asstring := sCaixa;
                qrGravaNaoFiscaldata.asdatetime := dData_Sistema;
                qrGravaNaoFiscalhora.AsDateTime := Time;
                qrGravaNaoFiscalindice.asstring := 'RG';
                qrGravaNaoFiscalDescricao.asstring := 'RELATÓRIO GERENCIAL';
                qrGravaNaoFiscalvalor.asfloat := 0;
                qrGravaNaoFiscalCOO.asstring := sCOO_Prestacao;
                qrGravaNaoFiscalGNF.asstring := sGNF_Prestacao;
                qrGravaNaoFiscalGRG.asstring := sGRG_Prestacao;
                qrGravaNaoFiscalCDC.Clear;
                qrGravaNaoFiscalDENOMINACAO.asstring := 'RG';
                qrGravaNaoFiscal.Post;
              end;
            end;
            // lancar os dados do comprador, mesmo sem cadastro ou a vista
            spCupom_consumidor.Close;
            spCupom_consumidor.parambyname('codigo').Asstring := sCod_Cupom;
            if sCli_codigo <> '' then
              spCupom_consumidor.parambyname('cod_cliente').Asinteger := StrToInt(sCli_codigo)
            else
              spCupom_consumidor.parambyname('cod_cliente').Asinteger := 1;
            spCupom_consumidor.parambyname('nome').asstring := Copy(sCli_Nome, 1, 80);
            spCupom_consumidor.parambyname('endereco').asstring := Copy(sCli_Endereco, 1, 80);
            spCupom_consumidor.parambyname('cidade').asstring := Copy(sCli_Cidade, 1, 50);
            spCupom_consumidor.parambyname('uf').asstring := Copy(sCli_uf, 1, 2);
            spCupom_consumidor.parambyname('cep').asstring := Copy(scli_cep, 1, 9);
            spCupom_consumidor.parambyname('cpf').asstring := Copy(sCli_CPF, 1, 18);
            spCupom_consumidor.parambyname('placa').asstring := Copy(sCli_Placa, 1, 10);
            spCupom_consumidor.parambyname('km').asstring := Copy(sCli_Km, 1, 10);
            spCupom_consumidor.parambyname('vendedor').asstring := Copy(scli_vendedor, 1, 50);
            spCupom_consumidor.Prepare;
            spCupom_consumidor.Execute;
          end;
          bVenda := False;
          ///////////////////////////////////////////////////
          ///  Finaliza processo de reabertura de cupom
          if ReaberturaCupomAtiva <> '' then
            frmCupom_Menu.agbReabrirClick(nil);
          ///////////////////////////////////////////////////
          if not (AtivaTouch) then
          begin
            if bBusca_foto_produto then
            begin
              img_produto.Picture := nil;
            end;
            if bBusca_foto_produto then
            begin
              img_produto.Visible := True;
            end;
          end;
          if ((frmPrincipal.TipoImpressora = NaoFiscal) and (TipoImprimeCupom <> 'N')) or (bAtivaImpFicha) then
            ImprimirCupomNaoFiscal(sCod_Cupom)   //qiqi
          else
            ImprimirComprovanteCrediario;
          if (bImpressaoCozinha) and not (bAtivaImpFicha) then
            ImprimirComanda(sCod_Cupom);

          if (bCadastra_Crediario) and (rvalor_total_crediario > 0) and (bImprimeCarne) then
          begin
            qrCarne.Close;
            qrCarne.Params.Items[0].Value := sCod_Cupom;
            qrCarne.Open;
            frImpressao.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000008.fr3');
            frImpressao.PrintOptions.Printer := sImpressora_carne;
            if bEdicaoRelatorio then
              frImpressao.DesignReport
            else
              frImpressao.ShowReport;
          end;
          if bImprimeCupomCredito then
          begin
            qrImpCredito.Close;
            qrImpCredito.Params.Items[0].Value := sNumero_Cupom;
            qrImpCredito.Open;
            relCredito.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000009.fr3');
            relCredito.ShowReport;
          end;

          if (not(rSat.Numero > 0)) and ((cdsRecebimentoTipo.AsString = StRecCartaoCR) or
                                        (cdsRecebimentoTipo.AsString = StRecCartaoDE)) then
             ImprimeRecebimentosCartao;
          // Excluir os arquivos temporarios
          frmModulo.spCupom_Temp_delete.prepare;
          frmModulo.spCupom_Temp_delete.execute;

          Limpa_controles;
          if vRecebimento.Troco > 0 then
            Imprime_display('Pago: R$ ' + FormatFloat('###,###,##0.00', vRecebimento.ValorPago) + '            ' + 'Troco: R$ ' + FormatFloat('###,###,##0.00', vRecebimento.Troco), clWhite, tiPgto)
          else
            Imprime_display('C A I X A    L I V R E', clWhite, tiLivre);

          TimerTroco.Enabled := True; // Apos 5 segundos

          MostraFotoProduto(False);
          MostraLogoMarca(True);

        end;
      finally
        iPre_Venda_cliente := 0;
        bVerificaIniternet := True;
//        if not (frmModulo.Conexao_Servidor.AutoCommit) and (frmModulo.Conexao_Servidor.InTransaction) then
//          frmModulo.Conexao_Servidor.Commit;

        PopupMenu := pop_principal;

        //BlockInput(false);
        if bVenda then
        begin
          gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'SIM');
        end
        else
        begin
          gravaINI(sConfiguracoes, 'PDV', 'Aberto', 'NAO');
          CorEditTotaL;
        end;
        {: ACIONAR GUILHOTINA}
        if frmPrincipal.TipoImpressora = NaoFiscal then
          AcionarGuilhotinaNaoFiscal;
        vcontingencia := 'N';
        vgerado_nfce := 'N';
        vNumNFCe := 0;
        ed_total_item.Value := 0;

        if bLanca_Mesa then
          menu_mesaClick(frmVenda)
        else if bLanca_Comanda then
          C2Click(frmVenda)
        else if bLanca_Delivery then
          mnDeliveryClick(frmVenda)
        else if not AtivaTouch then
          if ed_barra.Enabled then
            ed_barra.SETFOCUS;
        cdsDevolucao.Close;
        cdsDevolucao.CreateDataSet;
        LancaRecebimento.ValorDesconto := 0;
        LancaRecebimento.ValorAcressimo := 0;
        tmCaixa.Enabled := True;
      end;
    end
    else
    begin
      with frmModulo do
      begin
        qrItens.First;
        while not qrItens.Eof do
        begin
          if qrItensDESCONTO_CUPOM.AsFloat > 0 then
          begin
            qrItens.Edit;
            qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat+qrItensDESCONTO_CUPOM.AsFloat,-2);
            qrItensVALOR_DESCONTO.AsFloat := RoundTo(qrItensVALOR_DESCONTO.AsFloat-qrItensDESCONTO_CUPOM.AsFloat,-2);
            qrItensQTD_APLICADO_DESCONTO.AsFloat := 0;
            qrItensAPLICADO_DESCONTO_GRUPO.AsString := 'N';
            rTotal_Venda := RoundTo(rTotal_Venda + qrItensDESCONTO_CUPOM.AsFloat,-2);
            rTotal_Desconto := RoundTo(rTotal_Desconto - qrItensDESCONTO_CUPOM.AsFloat,-2);
            qrItensDESCONTO_CUPOM.AsFloat := 0;
            qrItens.Post;
          end;
          qrItens.Next;
        end;
      end;
    end;
  end
  else
  begin
    Imprime_display('Não há cupom aberto!', clYellow, tiAlerta);
    Sleep(1000);
    Imprime_display_anterior;
  end;
end;

procedure TfrmVenda.ImprimeRecebimentosCartao;
var flt:string;
    flted:boolean;
///
  Procedure ImprimeRecibo(pCampo:tField);
  begin
    if frmTef.cpPosPrinter.Ativo then
    begin
      frmTef.cpPosPrinter.Imprimir;
      frmTef.cpPosPrinter.Desativar;
    end;
    frmTef.cpPosPrinter.Ativar;
    frmTef.cpPosPrinter.BUFFER.Add('</ae>');
    frmTef.cpPosPrinter.BUFFER.Add(pCampo.AsString);
    frmTef.cpPosPrinter.BUFFER.Add('</pular_linhas>');
    frmTef.cpPosPrinter.BUFFER.Add('</corte_total>');
    frmTef.cpPosPrinter.imprimir;
    frmTef.cpPosPrinter.desativar;
  end;
///
begin
    fltEd:=cdsRecebimento.Filtered;
    flt:=cdsRecebimento.Filter;
    frmTef.cpPosPrinter.ativar;
    cdsRecebimento.Filtered:=False;
    cdsRecebimento.close;
    cdsRecebimento.open;
    try
      while not cdsRecebimento.eof do
      begin
        if cdsRecebimentoRec1.asstring<>'' then
          ImprimeRecibo(cdsRecebimentoRec1);
        if cdsRecebimentoRec1.asstring<>'' then
          ImprimeRecibo(cdsRecebimentoRec2);
        cdsRecebimento.Next;
      end;
    Finally
      cdsRecebimento.Filtered:=flted;
      cdsRecebimento.filter:=flt;
      cdsRecebimento.close;
      cdsRecebimento.open;
    end;
end;

procedure TfrmVenda.mnDeliveryClick(Sender: TObject);
var
  nitem, id: Integer;
begin
  if bVenda then begin
    Imprime_display('Venda Aberta!', clRed, tiErro);
    Sleep(1500);
    Imprime_display_anterior;
    Exit;
  end;
  lbEstoque.Visible := False;
  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  id := RetIdFecharDelivery;


  // iniciar a impressao da pre-venda
  if id > 0 then begin
    Imprime_display('Iniciando impressão do Delivery!', clWhite, tiInfo);
    try
      // verificar o serial do ec
      // Abrir o cupom fiscal
      // bloquear ao teclado
      //BlockInput(True);
      if not Abre_Venda then begin
        // Caso o comando de abertura de venda retornou false, abortar o processo;
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
      // desbloquear o teclado
      //BlockInput(False);
    end;
    rTotal_TaxaEntrega := frmModulo.qrDeliveryVALOR_TAXA_ENTREGA.AsFloat;
    if rTotal_TaxaEntrega > 0 then begin
      qrServidor.Close;
      qrServidor.SQL.Clear;
      qrServidor.SQL.Add('select PRODUTO_TAXA_ENTREGA from c000001');
      qrServidor.Open;
      if Trim(qrServidor.FieldByName('PRODUTO_TAXA_ENTREGA').AsString) = '' then begin
        Application.MessageBox('Não foi informado nas configurações do gerencial o produto para lançamento da taxa de entrega.', 'Atenção!', MB_ICONINFORMATION);
        Exit;
      end;
    end;
    bLanca_Delivery := True;
    iID_Delivery := id;
    with frmModulo do begin
      qrImpDelivery.Close;
      qrImpDelivery.ParamByName('ID').AsInteger := id;
      qrImpDelivery.Open;
      qrImpDelivery.First;
      nitem := 1;
      while not qrImpDelivery.eof do begin
        sProd_nome := qrImpDelivery.fieldbyname('produto').asstring;
        sProd_unidade := qrImpDelivery.fieldbyname('unidade').asstring;
        sProd_CST := qrImpDelivery.fieldbyname('cst').asstring;
        rProd_aliquota := qrImpDelivery.fieldbyname('aliquota').asfloat;
        iProd_codigo := qrImpDelivery.fieldbyname('codproduto').asinteger;
        sProd_barra := qrImpDelivery.fieldbyname('codbarra').asstring;
        rProd_qtde := qrImpDelivery.fieldbyname('quantidade').asfloat;
        rProd_preco := qrImpDelivery.fieldbyname('unitario').asfloat;
        rProd_total := qrImpDelivery.fieldbyname('total').asfloat;
        rProd_desconto := 0;
        rProd_acrescimo := 0;
        sProd_Tamanho := '';
        sProd_Cor := '';
        sProd_Compl_Pizza := qrImpDeliveryCOMPLEMENTO.AsString;
        vExtra.Texto := qrImpDeliveryEXTRA.AsString;
        // registrar o item
        frmModulo.query_servidor2.Close;
        frmModulo.query_servidor2.SQL.Clear;
        frmModulo.query_servidor2.SQL.Add('select * from ITEM_DELIVERY_SUB where ID_ITEM = ' + QuotedStr(qrImpDeliveryID_ITEM.AsString));
        frmModulo.query_servidor2.Open;
        frmModulo.query_servidor2.First;
        strPizza.Pizza := '';
        strPizza.Sabores := '';
        strPizza.Valor := 0;
        SetLength(strPizza.Produtos, 0);
        strPizza.ProdutoPrincipal.Codigo := '';
        strPizza.ProdutoPrincipal.Valor := 0;
        while not frmModulo.query_servidor2.Eof do begin
          cdsLancaProdutos.Insert;
          cdsLancaProdutosCODPRO.AsString := frmModulo.query_servidor2.FieldByName('CODPRO').AsString;
          cdsLancaProdutosQTDE.AsFloat := frmModulo.query_servidor2.FieldByName('QTDE').AsFloat;
          cdsLancaProdutos.Post;
          frmModulo.query_servidor2.Next;
        end;
        Registra_Item;
        Inc(nitem);
        Application.ProcessMessages;
        qrImpDelivery.next;
      end;
    end;
    if rTotal_TaxaEntrega > 0 then begin
      with frmModulo do begin
        qrTemp.Close;
        qrTemp.SQL.Clear;
        qrTemp.SQL.Add('select * from estoque where codigo = ' + QuotedStr(qrServidor.FieldByName('PRODUTO_TAXA_ENTREGA').AsString));
        qrTemp.Open;
        sProd_nome := qrTemp.FieldByName('NOME').AsString;
        sProd_unidade := qrTemp.FieldByName('UNIDADE').AsString;
        sProd_CST := qrTemp.FieldByName('CST').AsString;
        rProd_aliquota := qrTemp.FieldByName('ALIQUOTA').AsFloat;
        iProd_codigo := qrTemp.FieldByName('CODIGO').AsInteger;
        sProd_barra := qrTemp.FieldByName('COD_BARRA').AsString;
        rProd_qtde := 1;
        rProd_preco := RoundTo(rTotal_TaxaEntrega, -2);
        rProd_total := RoundTo(rTotal_TaxaEntrega, -2);
        rProd_desconto := 0;
        rProd_acrescimo := 0;
        sProd_Tamanho := '';
        sProd_Cor := '';
      end;
      // registrar o item
      Registra_Item;
    end;
    lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
    FecharCupom1Click(frmVenda);
  end;

end;

procedure TfrmVenda.mnEnviaContigenciaClick(Sender: TObject);
begin
  frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
  frmSenha_Supervisor.showmodal;

  if not bSupervisor_autenticado then
    Exit;
  if TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType = sitGreen then begin
    Application.CreateForm(TfrmEnviaContingencia, frmEnviaContingencia);
    frmEnviaContingencia.ShowModal;
  end
  else
    Application.MessageBox('Sem conexão com internet para envio das contingências!', 'Atenção!', MB_ICONERROR);
end;

procedure TfrmVenda.G1Click(Sender: TObject);
begin
  frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
  frmSenha_Supervisor.showmodal;

  if not bSupervisor_autenticado then
    Exit;

  frmNotasconsumidor := TfrmNotasconsumidor.create(Self);
  frmNotasconsumidor.showmodal;
  FreeAndNil(frmNotasconsumidor);
end;

procedure TfrmVenda.GerarNFCe(NumNFe: string);
var
  divide, i, nItem: Integer;
  total_desconto, total_acrescimo, total_produtos, total_servico, total_icms, total_base, total_trib, PTroco: Real;
  ok: Boolean;
  NfceServicoDF:Boolean;
begin
  try
    with frmModulo do begin
      total_acrescimo := 0;
      qrfilial.Open;
      with ACBRNFCe.NotasFiscais.add.NFe do begin
        Ide.cNF := GerarCodigoDFe(StrToInt(NumNFe));
        NfceServicoDF := False;
        if Trim(UpperCase(edtEmitUF)) =  'DF' then begin
          qrItens.First;
          while not qrItens.Eof do begin
            QRCSOSN.close;
            QRCSOSN.sql.clear;
            QRCSOSN.sql.add('select pro.codigo_servico servico');
            QRCSOSN.sql.Add('from estoque pro');
            QRCSOSN.sql.add('where pro.codigo = :pcodigo');
            QRCSOSN.ParamByName('pcodigo').asstring := IntToStrZero(qrItensCOD_PRODUTO.asinteger, 6);
            QRCSOSN.Open;
            if Trim(QRCSOSN.FieldByName('servico').AsString) <> '' then begin
              NfceServicoDF := True;
              Break;
            end;
            qrItens.Next;
          end;
        end;

        if NfceServicoDF then
          Ide.natOp := 'PRESTACAO DE SERVICOS'
        else
          Ide.natOp := 'VENDA AO CONSUMIDOR FINAL';
        if tipo_pgto = 1 then
          Ide.indPag := ipPrazo
        else
          Ide.indPag := ipVista;

        Ide.Modelo := 65;
        Ide.Serie := StrToInt(edtSerie);
        Ide.nNF := StrToInt(NumNFe);
        Ide.dEmi := Now;
        Ide.dSaiEnt := Now;
        Ide.hSaiEnt := Now;
        if frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao = frmtOffLine then begin
          Ide.dhCont := Now;
          DataHoraContingencia := Ide.dhCont;
          Ide.xJust := MotivoContigencia;
        end;
        Ide.tpNF := tnSaida;
        Ide.tpEmis := frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao;
        Ide.tpAmb := ACBRNFCe.Configuracoes.WebServices.Ambiente;
        Ide.cUF := UFtoCUF(edtEmitUF);
        Ide.cMunFG := StrToInt(edtEmitCodCidade);
        Ide.finNFe := fnNormal;
        Ide.tpImp := tiNFCe;
        Ide.indFinal := cfConsumidorFinal;
        Ide.indPres := pcPresencial;

        Ide.verProc := '1.0';
        Emit.CNPJCPF := edtEmitCNPJ;
        Emit.IE := edtEmitIE;
        Emit.IM := edtEmitIM;
        Emit.xNome := edtEmitRazao;
        Emit.xFant := edtEmitFantasia;

        Emit.EnderEmit.fone := edtEmitFone;
        Emit.EnderEmit.CEP := StrToInt(edtEmitCEP);
        Emit.EnderEmit.xLgr := edtEmitLogradouro;
        Emit.EnderEmit.nro := edtEmitNumero;
        Emit.EnderEmit.xCpl := edtEmitComp;
        Emit.EnderEmit.xBairro := edtEmitBairro;
        Emit.EnderEmit.cMun := StrToInt(edtEmitCodCidade);
        Emit.EnderEmit.xMun := edtEmitCidade;
        Emit.EnderEmit.UF := edtEmitUF;
        Emit.EnderEmit.cPais := 1058;
        Emit.EnderEmit.xPais := 'BRASIL';
        Emit.IEST := '';

        // regime tributario da empresa
        // (1-crtSimplesNacional, 2-crtSimplesExcessoReceita, 3-crtRegimeNormal)
        if qrfilial.fieldbyname('CRT').asinteger = 1 then
          Emit.CRT := crtSimplesNacional
        else if qrfilial.fieldbyname('CRT').asinteger = 2 then
          Emit.CRT := crtSimplesExcessoReceita
        else
          Emit.CRT := crtRegimeNormal;


        // Responsável Técnico
        if Trim(qrfilial.fieldbyname('REPTEC_CNPJ').asstring) <> '' then
          infRespTec.CNPJ := somenteNumero(qrfilial.fieldbyname('REPTEC_CNPJ').asstring);
        if Trim(qrfilial.fieldbyname('REPTEC_CONTATO').asstring) <> '' then
          infRespTec.xContato := qrfilial.fieldbyname('REPTEC_CONTATO').asstring;
        if Trim(qrfilial.fieldbyname('REPTEC_EMAIL').asstring) <> '' then
          infRespTec.email := qrfilial.fieldbyname('REPTEC_EMAIL').asstring;
        if Trim(qrfilial.fieldbyname('REPTEC_FONE').asstring) <> '' then
          infRespTec.fone := somenteNumero(qrfilial.fieldbyname('REPTEC_FONE').asstring);
        if qrfilial.fieldbyname('REPTEC_IDCSRT').AsInteger = 0 then
          infRespTec.idCSRT := qrfilial.fieldbyname('REPTEC_IDCSRT').AsInteger;
        if Trim(qrfilial.fieldbyname('REPTEC_CSRT').asstring) <> '' then
          infRespTec.fone := somenteNumero(qrfilial.fieldbyname('REPTEC_CSRT').asstring);



        // dados destinatarios

        Dest.CNPJCPF := OnlyNumber(sConsumidor_CPF);
        Dest.ISUF := '';
        Dest.xNome := sConsumidor_Nome;
        Dest.indIEDest := inNaoContribuinte;
        Dest.EnderDest.fone := '';
        if scli_cep <> '' then
          Dest.EnderDest.CEP := StrToInt(OnlyNumber(scli_cep))
        else
          Dest.EnderDest.CEP := 0;

        Dest.EnderDest.xLgr := '';
        Dest.EnderDest.nro := '';
        Dest.EnderDest.xCpl := sConsumidor_Endereco;
        Dest.EnderDest.xBairro := '';
        Dest.EnderDest.cMun := 0; // analisar funcao
        Dest.EnderDest.xMun := '';
        Dest.EnderDest.UF := sCli_uf;
        Dest.EnderDest.cPais := 1058;
        Dest.EnderDest.xPais := 'BRASIL';

        // Adicionando Produtos
        nItem := 1;
        total_produtos := 0;
        total_servico := 0;
        total_icms := 0;
        total_base := 0;
        total_trib := 0;
        total_desconto := 0;
        qrItens.First;
        while not qrItens.Eof do begin
          // prepar query
          qrProdNCFE.close;
          qrProdNCFE.sql.clear;
          qrProdNCFE.sql.add('select CODIGO, NOME,COD_BARRA,NCM,CST,ALIQNACIONAL,CFOP,CODIGO_ANP,PERCGLNN,PGLP,VPART,PERCGLGNI,informa_codigo_barra_xml, DESCRICAO_ANP from ESTOQUE ');
          qrProdNCFE.sql.add('where CODIGO = :pcodigo');
          qrProdNCFE.ParamByName('pcodigo').asinteger := qrItensCOD_PRODUTO.asinteger;
          qrProdNCFE.Open;

          if qrItensCANCELADO.AsInteger = 0 then begin
            with Det.add do begin
              Prod.nItem := nItem;
              nItem := nItem + 1;
              Prod.cProd := qrItensCOD_PRODUTO.AsString;
              Prod.uCom := qrItensUNIDADE.asstring;
              Prod.uTrib := qrItensUNIDADE.asstring;
              if qrProdNCFE.FieldByName('informa_codigo_barra_xml').AsString = 'S' then begin
                Prod.cEAN := qrProdNCFE.fieldbyname('COD_BARRA').asstring;
                Prod.cEANTrib := qrProdNCFE.fieldbyname('COD_BARRA').asstring;
              end
              else begin
                Prod.cEAN := 'SEM GTIN';
                Prod.cEANTrib := 'SEM GTIN';
              end;
              Prod.xProd := qrProdNCFE.fieldbyname('NOME').asstring;
              Prod.NCM := somenteNumero(qrProdNCFE.fieldbyname('NCM').asstring);
              Prod.EXTIPI := '';
              if Trim(qrProdNCFE.fieldbyname('CFOP').AsString) <> '' then
                Prod.CFOP := Trim(qrProdNCFE.fieldbyname('CFOP').AsString)
              else
                Prod.CFOP := edtcfop;
              Prod.qCom := qrItensQTDE.AsFloat;
              Prod.qTrib := Prod.qCom;
              Prod.vUnCom := qrItensVALOR_UNITARIO.asfloat;
              Prod.vUnTrib := Prod.vUnCom;
              Prod.vProd := RoundTo(Prod.vUnCom * Prod.qCom, -2);
              Prod.vDesc := 0;
              Prod.vOutro := qrItensVALOR_ACRESCIMO.asfloat; // RoundTo((Prod.vProd / (rTotal_Venda - rTotal_Acrescimo)) * rTotal_Acrescimo, -2);
              Prod.vDesc := qrItensVALOR_DESCONTO.asfloat; // RoundTo(Prod.vDesc + ((Prod.vProd / rTotal_Venda) * rTotal_Desconto), -2);
              if trim(qrProdNCFE.FieldByName('CODIGO_ANP').AsString) <> '' then begin
                Prod.comb.cProdANP := qrProdNCFE.FieldByName('CODIGO_ANP').AsInteger;
                Prod.comb.descANP :=  qrProdNCFE.FieldByName('DESCRICAO_ANP').AsString;
                Prod.comb.UFcons := qrfilialUF.AsString;
                Prod.comb.pGLP := qrProdNCFE.FieldByName('PGLP').AsFloat;
                Prod.comb.VPART := qrProdNCFE.FieldByName('VPART').AsFloat;
                Prod.comb.pGNi := qrProdNCFE.FieldByName('PERCGLGNI').AsFloat;
                Prod.comb.pGNn := qrProdNCFE.FieldByName('PERCGLNN').AsFloat;
              end;
              QRCSOSN.close;
              QRCSOSN.sql.clear;
              QRCSOSN.sql.add('select pro.csosn, pro.precocusto, pro.ALIQUOTA, pro.cest, pro.codigo_servico servico, s.st st_efetivo ');
              QRCSOSN.sql.Add('from estoque pro');
              QRCSOSN.sql.Add('left join csosn s on s.codigo = pro.csosn');
              QRCSOSN.sql.add('where pro.codigo = :pcodigo');
              QRCSOSN.ParamByName('pcodigo').asstring := IntToStrZero(qrItensCOD_PRODUTO.asinteger, 6);
              QRCSOSN.Open;
              if (Trim(QRCSOSN.fieldbyname('cest').asstring) <> '') and (Trim(QRCSOSN.fieldbyname('cest').asstring) <> '0') then
                Prod.CEST := QRCSOSN.fieldbyname('cest').asstring;
              Prod.vFrete := 0;
              Prod.vSeg := 0;
              total_desconto := RoundTo(total_desconto + Prod.vDesc, -2);
              total_acrescimo := RoundTo(total_acrescimo + Prod.vOutro, -2);
              with Imposto do begin
                vTotTrib := RoundTo((Prod.vProd * qrProdNCFE.fieldbyname('ALIQNACIONAL').AsFloat) / 100, -2);
                total_trib := total_trib + vTotTrib;
                if Trim(QRCSOSN.FieldByName('servico').AsString) <> '' then begin
                  ISSQN.vBC := Prod.vProd;
                  ISSQN.vAliq := qrfilialiss.AsFloat;
                  ISSQN.vISSQN := RoundTo((ISSQN.vBC * ISSQN.vAliq) / 100, -2);
                  ISSQN.cMunFG := qrfilialcod_municipio_ibge.AsInteger;
                  ISSQN.cListServ := Trim(QRCSOSN.FieldByName('servico').AsString);
                  ISSQN.indISS := iiExigivel;
                  ISSQN.cMun := qrfilialcod_municipio_ibge.AsInteger;
                  ISSQN.indIncentivo := iiSim;
                  PIS.CST := pis99;
                  COFINS.CST := cof99;
                  with Total.ISSQNtot do begin
                    vServ := vServ + ISSQN.vBC;
                    vBC := vBC + ISSQN.vBC;
                    vISS := vISS + ISSQN.vISSQN;
                    dCompet := Date;
                  end;
                  total_servico := RoundTo(total_servico + Prod.vProd, -2);
                end
                else begin
                  // simples nacional
                  total_produtos := RoundTo(total_produtos + Prod.vProd, -2);
                  if qrfilial.fieldbyname('CRT').asinteger = 1 then begin
                    with ICMS do begin
                      if (QRCSOSN.fieldbyname('csosn').asstring = '101') then
                        ICMS.CSOSN := csosn101
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '102') then
                        ICMS.CSOSN := csosn102
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '103') then
                        ICMS.CSOSN := csosn103
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '300') then
                        ICMS.CSOSN := csosn300
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '400') then
                        ICMS.CSOSN := csosn400
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '201') then
                        ICMS.CSOSN := csosn201
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '202') then
                        ICMS.CSOSN := csosn202
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '500') then
                        ICMS.CSOSN := csosn500
                      else if (QRCSOSN.fieldbyname('csosn').asstring = '900') then
                        ICMS.CSOSN := csosn900;



                      if (QRCSOSN.FieldByName('st_efetivo').AsString = 'S') and (Prod.CFOP = '5405') then begin
                        ICMS.vBCEfet := Prod.vProd;
                        ICMS.pICMSEfet := QRCSOSN.fieldbyname('ALIQUOTA').AsFloat;
                        ICMS.vICMSEfet := (Prod.vProd*QRCSOSN.fieldbyname('ALIQUOTA').AsFloat)/100;
                      end;
                      case ICMS.CSOSN of
                        csosn101:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                            ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                          end;
                        csosn102:
                          begin
                            ICMS.orig := oeNacional;
                          end;
                        csosn201:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0;
                            ICMS.pRedBCST := 00.00;
                            ICMS.vBCST := 0;
                            ICMS.pICMSST := 00.00;
                            ICMS.vICMSST := 0;
                            ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                            ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                          end;
                        csosn202:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0;
                            ICMS.pRedBCST := 00.00;
                            ICMS.vBCST := 0;
                            ICMS.pICMSST := 00.00;
                            ICMS.vICMSST := 0;
                          end;
                        csosn500:
                          begin
                            ICMS.orig := oeNacional;
//                            ICMS.vBCSTRet := 0;
//                            ICMS.vICMSSTRet := 0;
                          end;
                        csosn900:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pRedBC := 0;
                            ICMS.pICMS := RoundTo(QRCSOSN.FieldByName('ALIQUOTA').AsFloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0;
                            ICMS.pRedBCST := 00.00;
                            ICMS.vBCST := 0;
                            ICMS.pICMSST := 00.00;
                            ICMS.vICMSST := 0;
                            ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                            ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                          end;
                      end;
                      with ICMSUFDest do begin
                        vBCUFDest := 0.00;
                        pFCPUFDest := 0.00;
                        pICMSUFDest := 0.00;
                        pICMSInter := 0.00;
                        pICMSInterPart := 0.00;
                        vFCPUFDest := 0.00;
                        vICMSUFDest := 0.00;
                        vICMSUFRemet := 0.00;
                      end;
                    end;
                  end
                  else begin
                    with ICMS do begin
                      if (qrItensCST.asstring = '000') then
                        CST := cst00
                      else if (qrItensCST.asstring = '010') then
                        CST := cst10
                      else if (qrItensCST.asstring = '020') then
                        CST := cst20
                      else if (qrItensCST.asstring = '030') then
                        CST := cst30
                      else if (qrItensCST.asstring = '040') then
                        CST := cst40
                      else if (qrItensCST.asstring = '050') then
                        CST := cst50
                      else if (qrItensCST.asstring = '060') then
                        CST := cst60
                      else if (qrItensCST.asstring = '070') then
                        CST := cst70
                      else if (qrItensCST.asstring = '080') then
                        CST := cst80
                      else if (qrItensCST.asstring = '090') then
                        CST := cst90;
                      case ICMS.CST of
                        cst00:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                        cst10:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0;
                            ICMS.pRedBCST := 00.00;
                            ICMS.vBCST := 0;
                            ICMS.pICMSST := 00.00;
                            ICMS.vICMSST := 0;
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                        cst20:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.pRedBC := 0.00;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            ICMS.vICMSDeson := 0.00;
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                        cst30:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0.00;
                            ICMS.pRedBCST := 0.00;
                            ICMS.vBCST := 0;
                            ICMS.pICMSST := 0;
                            ICMS.vICMSST := 0;
                            ICMS.vICMSDeson := 0.00;
                          end;
                        cst40, cst41, cst50:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.vICMSDeson := 0.00;
                          end;
                        cst51:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.pRedBC := 0.00;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMSOp := 0.00;
                            ICMS.pDif := 0.00;
                            ICMS.vICMSDif := 0.00;
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                        cst60:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.vBCSTRet := 0.00;
                            ICMS.vICMSSTRet := 0.00;
                          end;
                        cst70:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.pRedBC := 0.00;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0.00;
                            ICMS.pRedBCST := 0.00;
                            ICMS.vBCST := 0.00;
                            ICMS.pICMSST := 0.00;
                            ICMS.vICMSST := 0.00;
                            ICMS.vICMSDeson := 0.00;
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                        cst90:
                          begin
                            ICMS.orig := oeNacional;
                            ICMS.modBC := dbiValorOperacao;
                            ICMS.vBC := RoundTo(qrItensVALOR_TOTAL.asfloat, -2);
                            ICMS.pRedBC := 0.00;
                            ICMS.pICMS := RoundTo(qrItensALIQUOTA.asfloat, -2);
                            ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                            ICMS.modBCST := dbisMargemValorAgregado;
                            ICMS.pMVAST := 0.00;
                            ICMS.pRedBCST := 0.00;
                            ICMS.vBCST := 0.00;
                            ICMS.pICMSST := 0.00;
                            ICMS.vICMSST := 0.00;
                            ICMS.vICMSDeson := 0.00;
                            total_icms := total_icms + ICMS.vICMS;
                            total_base := total_base + ICMS.vBC;
                          end;
                      end;
                    end;
                  end;
                end;
              end;
            end;
          end;
          qrItens.Next;
        end;
//        if total_acrescimo <> rTotal_Acrescimo then
//          Det[Det.Count - 1].Prod.vOutro := Det[Det.Count - 1].Prod.vOutro + (rTotal_Acrescimo - total_acrescimo);
//        if total_desconto <> rTotal_Desconto then
//          Det[Det.Count - 1].Prod.vDesc := Det[Det.Count - 1].Prod.vDesc + (rTotal_Acrescimo - total_acrescimo);

        Total.ICMSTot.vBC := total_base;
        Total.ICMSTot.vICMS := total_icms;
        Total.ICMSTot.vBCST := 0;
        Total.ICMSTot.vST := 0;
        Total.ICMSTot.vProd := total_produtos;
        Total.ICMSTot.vFrete := 0;
        Total.ICMSTot.vSeg := 0;
        Total.ICMSTot.vDesc := total_desconto;
        Total.ICMSTot.vII := 0;
        Total.ICMSTot.vIPI := 0;
        Total.ICMSTot.vPIS := 0;
        Total.ICMSTot.vCOFINS := 0;
        Total.ICMSTot.vOutro := total_acrescimo;
        Total.ICMSTot.vNF := RoundTo(total_produtos + total_servico - total_desconto + total_acrescimo, -2);

        // partilha do icms e fundo de pobreza
        Total.ICMSTot.vFCPUFDest := 0.00;
        Total.ICMSTot.vICMSUFDest := 0.00;
        Total.ICMSTot.vICMSUFRemet := 0.00;
        Total.ICMSTot.vTotTrib := total_trib;

        Transp.modFrete := mfSemFrete; // NFC-e não pode ter FRETE
        PTroco := 0;
        divide := 0;
        if vRecebimento.Troco > 0 then begin
          cdsRecebimento.First;
          while not cdsRecebimento.Eof do begin
            if (cdsRecebimentoTipo.AsString = StRecDinheiro) or (cdsRecebimentoTipo.AsString = StRecCheque) then
              Inc(divide);
            cdsRecebimento.Next;
          end;
          PTroco := (vRecebimento.Troco / divide);
        end;
        cdsRecebimento.First;
        while not cdsRecebimento.Eof do begin
          with pag.add do // PAGAMENTOS apenas para NFC-e
          begin
            if (cdsRecebimentoTipo.AsString = StRecCheque) then
              tPag := fpCheque;
            if (cdsRecebimentoTipo.AsString = StRecCrediario) then
              tPag := fpCreditoLoja;
            if (cdsRecebimentoTipo.AsString = StRecCartaoDE) then begin
              tPag := fpCartaoDebito;
              tBand := StrToBandeiraCartao(ok, cdsRecebimentobandeira.AsString);
              if (cdsRecebimentotef_codigo_autorizacao.AsString <> '') then begin
                tpIntegra := tiPagIntegrado;
//                indPag := ipVista;
                cAut := cdsRecebimentotef_codigo_autorizacao.AsString;
                CNPJ := cdsRecebimentotef_cnpj.AsString;
              end else
                tpIntegra := tiPagNaoIntegrado;
            end;
            if (cdsRecebimentoTipo.AsString = StRecCartaoCR) then begin
              tPag := fpCartaoCredito;
              tBand := StrToBandeiraCartao(ok, cdsRecebimentobandeira.AsString);
              if (cdsRecebimentotef_codigo_autorizacao.AsString <> '') then begin
                tpIntegra := tiPagIntegrado;
//                indPag := ipPrazo;
                cAut := cdsRecebimentotef_codigo_autorizacao.AsString;
                CNPJ := cdsRecebimentotef_cnpj.AsString;
              end else
                tpIntegra := tiPagNaoIntegrado;
            end;
            if (cdsRecebimentoTipo.AsString = StRecDinheiro) then
              tPag := fpDinheiro;
            if (cdsRecebimentoTipo.AsString = StRecConvenio) then
              tPag := fpCreditoLoja;
            if (cdsRecebimentoTipo.AsString = StRecCartaoDE) or (cdsRecebimentoTipo.AsString = StRecCartaoCR) or (cdsRecebimentoTipo.AsString = StRecCrediario) or (cdsRecebimentoTipo.AsString = StRecConvenio) then
              vPag := cdsRecebimentovalor.AsFloat
            else
              vPag := cdsRecebimentovalor.AsFloat - PTroco;
          end;
          cdsRecebimento.Next;
        end;
        InfAdic.infCpl := 'Atendente: ' + sNome_Operador;
        if sVendedorNome <> '' then begin
          InfAdic.infCpl := InfAdic.infCpl + ';';
          InfAdic.infCpl := InfAdic.infCpl + 'Vendedor: ' + sVendedorNome;
        end;

        if (sMensagem_Cupom.Lines.Text <> '') and (Men_Cup_NFCE) then begin
          InfAdic.infCpl := InfAdic.infCpl + ';';
          InfAdic.infCpl := InfAdic.infCpl + sMensagem_Cupom.Lines.Text;
        end;
        if UtilizaSenha then begin
          InfAdic.infCpl := InfAdic.infCpl + ';';
          InfAdic.infCpl := InfAdic.infCpl + 'Senha de Retirada: ' + IntToStr(SenhaImpressao);
        end;
        InfAdic.infCpl := InfAdic.infCpl + ';';
        InfAdic.infCpl := InfAdic.infCpl + 'Número do Cupom Intérno: ' + sNumero_Venda;
      end;
    end;
  except
    on E: Exception do
      Application.messagebox(PWideChar('Erro na geração da NFCE' + #13 + 'Erro: ' + E.Message), 'Erro', MB_OK + MB_ICONERROR);
  end;

end;

procedure TfrmVenda.GrupoClick(Sender: TObject);
var
  i, j: Integer;
begin
  for i := 0 to scGrupo.ComponentCount - 1 do begin
    if scGrupo.Components[i].Tag = TComponent(Sender).Tag then begin
      if scGrupo.Components[i].ClassType = TPanel then begin
        for j := 0 to scGrupo.Components[i].ComponentCount - 1 do begin
          if scGrupo.Components[i].Components[j].ClassType = TRzLabel then begin
            TRzLabel(TComponent(TPanel(scGrupo.Components[i]).Components[j])).Font.Color := clRed;
            TRzLabel(TComponent(TPanel(scGrupo.Components[i]).Components[j])).Font.Style := [fsBold];
          end;
        end;
      end;
    end
    else begin
      if scGrupo.Components[i].ClassType = TPanel then begin
        for j := 0 to scGrupo.Components[i].ComponentCount - 1 do begin
          if scGrupo.Components[i].Components[j].ClassType = TRzLabel then begin
            TRzLabel(TComponent(TPanel(scGrupo.Components[i]).Components[j])).Font.Color := clBlack;
            TRzLabel(TComponent(TPanel(scGrupo.Components[i]).Components[j])).Font.Style := [];
          end;
        end;
      end;
    end;
  end;
//  SetFocusedControl(TWinControl(Sender));
  MontaProdutos(TComponent(Sender).Tag);
end;

procedure TfrmVenda.HabilitarCancelamentoSangriaSuprimento(AAplicarTodos: Boolean);
begin
  if not uQry.sql.Text.Contains('SUPRIMENTO') or AAplicarTodos then
  begin
    uQry.Close;
    uQry.sql.Clear;
    uQry.sql.Add('select first 1 * from NAO_FISCAL where ECF = ' + QuotedStr(sCaixa) +
                 'AND descricao = ''SUPRIMENTO'' and data = Current_Date order by data desc, hora desc');
    uQry.Open;
    mniCancelarUltimoSuprimento.Enabled := uQry.RecordCount > 0;
  end;
  ///
  if not uQry.sql.Text.Contains('SANGRIA') or AAplicarTodos then
  begin
    uQry.Close;
    uQry.sql.Clear;
    uQry.sql.Add('select first 1 * from NAO_FISCAL where ECF = ' + QuotedStr(sCaixa) +
                 'AND descricao = ''SANGRIA'' and data = Current_Date order by data desc, hora desc');
    uQry.Open;
    mniCancelarUltimaSangria.Enabled := uQry.RecordCount > 0;
  end;
end;

procedure TfrmVenda.PrepararImpressao;
begin
  ACBrSATExtratoFortes1.LarguraBobina  := Sat_Largura;
  ACBrSATExtratoFortes1.MargemSuperior := Sat_MargemTopo;
  ACBrSATExtratoFortes1.MargemInferior := Sat_MargemFundo;
  ACBrSATExtratoFortes1.MargemEsquerda := Sat_MargemEsquerda;
  ACBrSATExtratoFortes1.MargemDireita  := Sat_MargemDireita;
  ACBrSATExtratoFortes1.MostraPreview  := Sat_Preview;
  ACBrSATExtratoFortes1.MostraStatus   := False;
  if Sat_Impressora <> '' then
    ACBrSATExtratoFortes1.Impressora := Sat_Impressora;
end;

function TfrmVenda.PrepararNFCE: Boolean;
var
  Acao: Boolean;
  i: Integer;
  Prestacao, Resto: Double;
  Validou:Boolean;
  retErroNFce:TRetornoErroNfce;
  MensagemDeErro:String;
label
  TentarNovamente;
begin
  ContingenciaPorErroNoEnvio := False;
  ErroNoEnvioContingencia := '';
  Result := True;
  nfce_autorizada := False;
  frmModulo.LerConfiguracao;
  if EdFormatoOff = 0 then
    frmtOffLine := teContingencia
  else
    frmtOffLine := teOffLine;
  if not Gera_Numero_NFCe_Local then
    vAux := frmModulo.codifica('888888')
  else begin
    frmModulo.qrNfceLocal.Close;
    frmModulo.qrNfceLocal.Open;
    vAux := frmModulo.qrNfceLocalNUMERO_NFCE.AsString;
    frmModulo.qrNfceLocal.Close;
  end;

  vNumNFCe := StrToInt(vAux);
  vSincrono := '1';
  vNumLote := vAux;
  Sincrono := True;
  with frmModulo do begin
    TentarNovamente:
    try
      ACBRNFCe.NotasFiscais.clear;

      Imprime_display('AGUARDE...  GERANDO NFC-E', clWhite, tiLivre);
      GerarNFCe(vAux);

      Imprime_display('AGUARDE...  ASSINANDO NFC-E', clWhite, tiLivre);
      ACBRNFCe.NotasFiscais.Assinar;

      if ACBRNFCe.Configuracoes.Geral.FormaEmissao <> frmtOffLine then begin
        Validou := True;
        Imprime_display('AGUARDE...  VALIDANDO NFC-E', clWhite, tiLivre);
        ACBRNFCe.NotasFiscais.Validar;

        Imprime_display('AGUARDE...  ENVIANDO NFC-E', clWhite, tiLivre);
        ACBRNFCe.Enviar(vNumLote, False, Sincrono);
        if ACBRNFCe.NotasFiscais.Items[0].Confirmada then begin
          if ACBRNFCe.WebServices.Enviar.cStat = 100 then begin
            cStatus := 100;
            ChaveNFCE := ACBRNFCe.NotasFiscais.Items[0].NFe.infNFe.Id;
            NumeroNFCe := StrToInt(vAux);
          end;
        end else begin
          Imprime_display(ACBRNFCe.WebServices.Enviar.xMotivo, clWhite, tiLivre);
        end;
      end
      else begin
        Validou := True;
        ChaveNFCE := ACBRNFCe.NotasFiscais.Items[0].NFe.infNFe.Id;
        NumeroNFCe := StrToInt(vAux);
      end;
    except
      on E: Exception do begin
        MensagemDeErro := E.Message;
        retErroNFce := ErroNFCE(E.Message);
        Validou := False;
      end;
    end;
    if not Validou then begin
      if retErroNFce = reTentarNovamente then begin
        if (Pos('12002',MensagemDeErro)>0) and (Pos('TIMEOUT',UpperCase(MensagemDeErro))>0) then begin
          Imprime_display('AGUARDE...  VERIFICANDO STATUS DA NFCE NA SEFAZ...', clWhite, tiLivre);
          Sleep(3000);
          ACBRNFCe.Consultar;
          if ACBRNFCe.WebServices.Consulta.cStat = 100 then begin
            Validou := True;
            cStatus := 100;
            ChaveNFCE := ACBRNFCe.NotasFiscais.Items[0].NFe.infNFe.Id;
            NumeroNFCe := StrToInt(vAux);
          end;
        end;
        if not Validou then
          goto TentarNovamente;
      end else if retErroNFce = reContingencia then begin
        ErroNoEnvioContingencia := MensagemDeErro;
        ContingenciaPorErroNoEnvio := True;
        bVerificaIniternet := False;
        ACBRNFCe.Configuracoes.Geral.FormaEmissao := frmtOffLine;
        ACBRNFCe.NotasFiscais.Clear;
        Imprime_display('AGUARDE...  GERANDO NFC-E - CONTINGÊNCIA', clWhite, tiLivre);
        GerarNFCe(vAux);
        Imprime_display('AGUARDE...  ASSINANDO NFC-E - CONTINGÊNCIA ', clWhite, tiLivre);
        ACBRNFCe.NotasFiscais.Assinar;
        ChaveNFCE := ACBRNFCe.NotasFiscais.Items[0].NFe.infNFe.Id;
        NumeroNFCe := StrToInt(vAux);
      end else if retErroNFce = reCancelar then begin
        rTotal_Desconto := rTotal_Desconto - vRecebimento.Desconto;
        rTotal_Venda := rTotal_Venda + vRecebimento.Desconto;
        RemoveDescontoTabelaItem;
        if not Gera_Numero_NFCe_Local then
          frmModulo.codifica('888888', 'D')
        else begin
          qrAltenaNumeroNFCeLocal.SQL.Clear;
          qrAltenaNumeroNFCeLocal.SQL.Text := 'SET GENERATOR NUMERADOR_NFCE TO ' + IntToStr(StrToInt(vAux)-1);
          qrAltenaNumeroNFCeLocal.ExecSQL;
        end;
        nfce_autorizada := False;
        if (TEF_Ativo) and (RecebeuTEF) then
          frmTEF.ACBrTEFD1.CancelarTransacoesPendentes;
        Exit;
      end;
    end;
    ACBRDANFENFCe.FastFile := 'C:\SOS\PDV\Schemas\DANFeNFCe.fr3';
    if FileExists(frmPrincipal.LerINi(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', '')) then
      ACBRDANFENFCe.Logo := frmPrincipal.LerINi(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', '');
    ACBRDANFENFCe.vTroco := vRecebimento.Troco;
    ACBRDANFENFCe.Impressora := sImpressora_Venda_Simples;
    ACBRDANFENFCe.MargemDireita := edMargDir;
    ACBRDANFENFCe.MargemEsquerda := edMargEsq;
    ACBRDANFENFCe.MargemSuperior := edMargSup;
    ACBRDANFENFCe.MargemInferior := edMarginf;
    ACBRDANFENFCe.MostraPreview := bExibeCumponNFTela;

    if TipoImprimeCupom = 'P' then begin
      Acao := False;
      if Application.MessageBox('Deseja Imprimir o Cupom?', 'Atenção!', MB_ICONQUESTION + MB_YESNO) = IDYES then
      begin
        Acao := True;
        ACBRNFCe.NotasFiscais.Imprimir;
        if (TEF_Ativo) and (RecebeuTEF) then
          finalizarCupomNFCe;
      end;
    end
    else
    begin
      Acao := True;
      ACBRNFCe.NotasFiscais.Imprimir;
      if (TEF_Ativo) and (RecebeuTEF) then
        finalizarCupomNFCe;
    end;

    if Acao then
      ImprimirComprovanteCrediario;

    ACBRNFCe.NotasFiscais.clear;

    if (cStatus = 100) or (frmModulo.ACBRNFCe.Configuracoes.Geral.FormaEmissao = frmtOffLine) then
      nfce_autorizada := True
    else
      nfce_autorizada := False;

  end;

end;

procedure TfrmVenda.ProdutoClick(Sender: TObject);
var
  i, j: Integer;
begin
  MouseLeave(Sender);
  for i := 0 to scProduto.ComponentCount - 1 do begin
    if scProduto.Components[i].Tag = TComponent(Sender).Tag then begin
      if scProduto.Components[i].ClassType = TPanel then begin
        for j := 0 to scProduto.Components[i].ComponentCount - 1 do begin
          if scProduto.Components[i].Components[j].ClassType = TRzLabel then begin
            TRzLabel(TComponent(TPanel(scProduto.Components[i]).Components[j])).Font.Color := clRed;
            TRzLabel(TComponent(TPanel(scProduto.Components[i]).Components[j])).Font.Style := [fsBold];
          end;
        end;
      end;
    end
    else begin
      if scProduto.Components[i].ClassType = TPanel then begin
        for j := 0 to scProduto.Components[i].ComponentCount - 1 do begin
          if scProduto.Components[i].Components[j].ClassType = TRzLabel then begin
            TRzLabel(TComponent(TPanel(scProduto.Components[i]).Components[j])).Font.Color := clBlack;
            TRzLabel(TComponent(TPanel(scProduto.Components[i]).Components[j])).Font.Style := [];
          end;
        end;
      end;
    end;
  end;
  if not bVenda then begin
    try
      // verificar o serial do ecf
      // Abrir o cupom fiscal
      // bloquear ao teclado

      if not Abre_Venda then begin
        bLanca_pre_venda := False;
        bLanca_OS := False;
        bLanca_Mesa := False;
        bLanca_Abastecimento := False;
        // Caso o comando de abertura de venda retornou false, abortar o processo;
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
      // desbloquear o teclado
    end;
  end;
  iPesquisa_produto := 1;
  if frmVenda.Localizar_Produto(Zerar(IntToStr(TComponent(Sender).Tag), 6)) then begin
    if (query.fieldbyname('usa_balanca').AsInteger = 1) and (bPesagemAutomatica) then begin
      ed_qtde.Value := RetornaPesoBalanca;
      rProd_qtde := ed_qtde.Value;
      if ed_qtde.Value = 0 then
        Exit;
    end
    else begin
      if ed_qtde.value > 0 then
        rProd_qtde := ed_qtde.Value
      else if ed_qtde.value <= 0 then
        ed_qtde.Value := rProd_qtde;
    end;
    qrTmp.Close;
    qrTmp.SQL.Clear;
    qrTmp.SQL.Add('select * from estoque est where est.codigo = ' + IntToStr(TComponent(Sender).Tag));
    qrTmp.Open;
    if qrTmp.FieldByName('pizza').AsString = 'S' then begin
      strPizza := Pizza(qrTmp.FieldByName('nome').AsString, rProd_preco);
      if strPizza.Sabores <> '' then begin
        ed_unitario.Value := rProd_preco;
        rProd_total := strPizza.Valor;
        Registra_Item;
      end;
    end
    else begin
      ed_unitario.Value := rProd_preco;
      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
      Registra_Item;
    end;
  end;
end;

procedure TfrmVenda.ed_qtdeKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    if rProd_qtde <> ed_qtde.value then
    begin
      if (ed_qtde.value > 0) and (ed_barra.text <> '') then
      begin
        ed_barra.text := ed_qtde.value.ToString + '*' + ed_barra.text;
        ed_qtde.value := 0;
        RegistrarProduto;
        Exit;
      end
      else
      begin
        rProd_preco := frmModulo.PrecoDeVenda(iProd_codigo, ed_qtde.value);
        rProd_qtde := ed_qtde.value;
        ed_unitario.Value := rProd_preco;
        rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
        ed_total_item.Value := rProd_total;
      end;
    end;
    if (bMuda_unitario) then begin
      Imprime_display(sProd_nome, clWhite, tiInfo);
      ed_unitario.ReadOnly := False;
      ed_unitario.setfocus;
      Exit;
    end;
    if (bMuda_total) then begin
      Imprime_display(sProd_nome, clWhite, tiInfo);
      ed_total_item.ReadOnly := False;
      ed_total_item.setfocus;
      Exit;
    end
    else
    begin
      if (rProd_total <= 0) or (rProd_qtde <= 0) or (rProd_preco <= 0) then begin
        Imprime_display('TOTAL NÃO PODE SER ZERADO OU NEGATIVO!', clRed, tiErro);
        Sleep(2000);
        Exit;
      end;
      ed_barra.text := '';
      Registra_Item;
    end;
  end
  else if Key = #27 then begin
    Teclado.Visible := False;
  end;
end;

procedure TfrmVenda.ed_qtdeEnter(Sender: TObject);
begin
  // esconder o panel de cancelamento de item, caso esteja visivel
  if AtivaTouch then
    ed_qtde.ReadOnly := False;
  if (AtivaTouch) and (ExibeTecladoVirtual) then begin
    Teclado.KeyboardType := ktNUMERIC; //'NumPad';
    Teclado.Width := 330;
    Teclado.Left := RoundDiv((Screen.Width - Teclado.Width), 2);
    Teclado.Top := RoundDiv((Screen.Height - Teclado.Height), 2);
    Teclado.Visible := True;
  end;
end;

procedure TfrmVenda.ed_qtdeExit(Sender: TObject);
begin
  if (AtivaTouch) and (ExibeTecladoVirtual) then
    Teclado.Visible := False;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ed_unitarioEnter(Sender: TObject);
begin
  // esconder o panel de cancelamento de item, caso esteja visivel
  if (AtivaTouch) and (ExibeTecladoVirtual) then begin
    Teclado.KeyboardType := ktNUMERIC; //'NumPad';
    Teclado.Width := 330;
    Teclado.Left := RoundDiv((Screen.Width - Teclado.Width), 2);
    Teclado.Top := RoundDiv((Screen.Height - Teclado.Height), 2);
    Teclado.Visible := True;
  end;
end;

procedure TfrmVenda.ed_unitarioExit(Sender: TObject);
begin
  if (AtivaTouch) and (ExibeTecladoVirtual) then
    Teclado.Visible := False;

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ed_total_itemKeyPress(Sender: TObject; var Key: Char);
var
  rd: real;
begin
  if Key = #13 then begin
    rprod_qtde := ed_total_item.value / ed_unitario.value;
    rProd_total := rProd_qtde * ed_unitario.Value;
    if rProd_total < ed_total_item.value then begin
      while formatfloat('########0.00', rProd_total) <> formatfloat('########0.00', ed_total_item.value) do begin
        rprod_qtde := rprod_qtde + 0.001;
        rProd_total := rprod_qtde * ed_unitario.Value;
      end;
    end
    else begin
      if rProd_total > ed_total_item.value then begin
        while formatfloat('########0.00', rProd_total) <> formatfloat('########0.00', ed_total_item.value) do begin
          rprod_qtde := rprod_qtde - 0.001;
          rProd_total := rprod_qtde * ed_unitario.Value;
        end;
      end;
    end;
    ed_qtde.value := rProd_qtde;
    if bMuda_unitario then begin
      if formatfloat('##########0.00', rProd_preco) <> formatfloat('##########0.00', ed_unitario.value) then begin
        // verificar seh eh desconto ou acrescimo
        if rProd_preco < ed_unitario.value then begin
          // foi acrescimo
          rProd_acrescimo := ed_unitario.value - rProd_preco;
        end
        else begin
          // foi desconto
          rProd_desconto := rProd_preco - ed_unitario.value;

          if (isenha_desconto = 1) and (rprod_desconto > 0) then begin
            if rProd_desconto_maximo > 0 then begin
              rd := (rprod_desconto / rProd_preco) * 100;
              if rd > rProd_desconto_maximo then begin
                frmsenha_supervisor := tfrmsenha_supervisor.create(self);
                frmsenha_supervisor.ValidaDesconto := True;
                frmsenha_supervisor.showmodal;
                if not bSupervisor_autenticado then
                  exit;
              end;
            end
            else begin
              frmsenha_supervisor := tfrmsenha_supervisor.create(self);
              frmsenha_supervisor.ValidaDesconto := True;
              frmsenha_supervisor.showmodal;
              if not bSupervisor_autenticado then
                exit;
            end;
          end;
        end;
      end;
    end;
    Registra_Item;
  end;
end;

procedure TfrmVenda.ACBrSAT1GetcodigoDeAtivacao(var Chave: AnsiString);
begin
  Chave := CodigoAtivacaoSAT;
end;

procedure TfrmVenda.ACBrSAT1GetsignAC(var Chave: AnsiString);
begin
  Chave := SwHAssinatura;
end;

procedure TfrmVenda.AcionaBalana1Click(Sender: TObject);
begin
  if frmModulo.balanca.Modelo <> balNenhum then begin
    try
      frmModulo.balanca.ativar;
      frmModulo.balanca.Ativo := True;
      rBal_peso := frmModulo.balanca.LePeso(iBal_time);
      Imprime_display('Peso do Produto: ' + FormatFloat('###,###,##0.00', rBal_peso), clYellow, tiPeso);
      frmModulo.balanca.Ativo := False;
      Application.ProcessMessages;
    except
      on E: Exception do begin
        Application.MessageBox(PWideChar('Ocorreu um erro ao tentar conectar-se a balança: ' + #13 + E.Message), 'Atenção!', MB_ICONINFORMATION);
        frmModulo.balanca.Ativo := False;
      end;
    end;
  end;
end;

procedure TfrmVenda.ConfiguraECF;
var
  IniFile: string;
  Ini: TIniFile;
begin
  IniFile := sConfiguracoes;
  Ini := TIniFile.Create(IniFile);
  try
    cpACBrECF.Modelo := TACBrECFModelo(INI.ReadInteger('ECF', 'Modelo', 0));
    cpACBrECF.Porta := INI.ReadString('ECF', 'Porta', '');
    cpACBrECF.TimeOut := 3;//INI.ReadInteger('ECF', 'TimeOut', 0);
    cpACBrECF.IntervaloAposComando := 100;//INI.ReadInteger('ECF', 'IntervaloAposComando', 0);
    cpACBrECF.ReTentar := INI.ReadBool('ECF', 'TentarNovamente', False);
    cpACBrECF.ControlePorta := INI.ReadBool('ECF', 'ControlePorta', true);
    cpACBrECF.BloqueiaMouseTeclado := INI.ReadBool('ECF', 'BloqueiaMouseTeclado', False);
    cpACBrECF.ExibeMensagem := INI.ReadBool('ECF', 'ExibeMsgAguarde', False);
    cpACBrECF.ArredondaPorQtd := INI.ReadBool('ECF', 'ArredondaPorQtd', False);
    cpACBrECF.DescricaoGrande := INI.ReadBool('ECF', 'DescricaoGrande', False);
    cpACBrECF.ArredondaItemMFD := INI.ReadBool('ECF', 'ArredondamentoItemMFD', False);
    cpACBrECF.GavetaSinalInvertido := INI.ReadBool('ECF', 'GavetaSinalInvertido', False);
    cpACBrECF.MsgAguarde := StringReplace(INI.ReadString('ECF', 'MensagemAguarde', ''), '|', sLineBreak, [rfReplaceAll]);
    cpACBrECF.ArqLOG := INI.ReadString('ECF', 'ArqLog', '');
    cpACBrECF.Device.ParamsString := INI.ReadString('ECF', 'SerialParams', '');
    cpACBrECF.Operador := sNome_Operador;
    cpACBrECF.Device.MaxBandwidth := INI.ReadInteger('ECF', 'BandWidth', 0);
    cpACBrECF.MaxLinhasBuffer := INI.ReadInteger('ECF', 'Buffer', 0);
    cpACBrECF.PaginaDeCodigo := INI.ReadInteger('ECF', 'PaginadeCodigo', 0);
    vEcfRecDinheiro := INI.ReadString('ECF', 'RecECFDinheiro', '');
    vEcfRecConvenio := INI.ReadString('ECF', 'RecECFConvenio', '');
    vEcfRecCrediario := INI.ReadString('ECF', 'RecECFCarteira', '');
    vEcfRecCartaoCR := INI.ReadString('ECF', 'RecECFCartaoCredito', '');
    vEcfRecCartaoDE := INI.ReadString('ECF', 'RecECFCartaoDebito', '');
    vEcfRecCheque := INI.ReadString('ECF', 'RecECFCheque', '');
    ECFCaminhoFiscal := INI.ReadString('ECF', 'CaminhoFiscal', 'C:\SOS');
    sMD5 := INI.ReadString('ECF', 'MD5', '');

  except
    on e:Exception do begin
      Application.MessageBox(PWideChar('Ocorreu um erro ao aplicar as configurações da ECF.'+#13+e.Message),'Atenção!',MB_ICONINFORMATION);
    end;
  end;
end;

procedure TfrmVenda.ConfiguraSAT;
var
  tmp, IniFile: string;
  Ini: TIniFile;
  Ok: Boolean;
  StreamMemo: TMemoryStream;
begin
  IniFile := sConfiguracoes;

  Ini := TIniFile.Create(IniFile);
  try
    with ACBrSAT1 do begin
      Modelo := TACBrSATModelo(Ini.ReadInteger('SAT', 'Modelo', 0));
      ArqLOG := Ini.ReadString('SAT', 'ArqLog', 'ACBrSAT.log');
      NomeDLL := Ini.ReadString('SAT', 'NomeDLL', 'C:\SAT\SAT.DLL');
      Config.ide_numeroCaixa := iNumCaixa;
      Config.ide_tpAmb := TpcnTipoAmbiente(Ini.ReadInteger('SAT', 'Ambiente', 1));
      Config.ide_CNPJ := Ini.ReadString('SwH', 'CNPJ', '');
      Config.emit_CNPJ := Ini.ReadString('Emit', 'CNPJ', '');
      Config.emit_IE := Ini.ReadString('Emit', 'IE', '');
      Config.emit_IM := Ini.ReadString('Emit', 'IM', '');
      Config.emit_cRegTrib := TpcnRegTrib(Ini.ReadInteger('Emit', 'RegTributario', 0));
      Config.emit_cRegTribISSQN := TpcnRegTribISSQN(Ini.ReadInteger('Emit', 'RegTribISSQN', 0));
      Config.emit_indRatISSQN := TpcnindRatISSQN(Ini.ReadInteger('Emit', 'IndRatISSQN', 0));
      Config.PaginaDeCodigo := Ini.ReadInteger('SAT', 'PaginaDeCodigo', 0);
      Config.EhUTF8 := Ini.ReadBool('SAT', 'UTF8', True);
      Config.infCFe_versaoDadosEnt := StringToFloat(FloatToString(Ini.ReadFloat('SAT', 'versaoDadosEnt', cversaoDadosEnt)));

      ConfigArquivos.SalvarCFe := Ini.ReadBool('SAT', 'SalvarCFe', True);
      ConfigArquivos.SalvarCFeCanc := Ini.ReadBool('SAT', 'SalvarCFeCanc', True);
      ConfigArquivos.SalvarEnvio := Ini.ReadBool('SAT', 'SalvarEnvio', True);
      ConfigArquivos.SepararPorCNPJ := Ini.ReadBool('SAT', 'SepararPorCNPJ', True);
      ConfigArquivos.SepararPorMes := Ini.ReadBool('SAT', 'SepararPorMES', True);
      ConfigArquivos.PastaCFeVenda := Ini.ReadString('SAT', 'PATH_SALVAR_VENDA', '');
      ConfigArquivos.PastaEnvio := Ini.ReadString('SAT', 'PATH_SALVAR_ENVIO', '');
      ConfigArquivos.PastaCFeCancelamento := Ini.ReadString('SAT', 'PATH_SALVAR_CANC', '');

    end;

    CodigoAtivacaoSAT := Ini.ReadString('SAT', 'CodigoAtivacao', '');
    SwHAssinatura := Ini.ReadString('SwH', 'Assinatura', '');
    Sat_FormatXML := Ini.ReadBool('SAT', 'FormatarXML', True);

    Sat_Largura := Ini.ReadInteger('Fortes', 'Largura', ACBrSATExtratoFortes1.LarguraBobina);
    Sat_MargemTopo := Ini.ReadInteger('Fortes', 'MargemTopo', Trunc(ACBrSATExtratoFortes1.MargemSuperior));
    Sat_MargemFundo := Ini.ReadInteger('Fortes', 'MargemFundo', Trunc(ACBrSATExtratoFortes1.MargemInferior));
    Sat_MargemEsquerda := Ini.ReadInteger('Fortes', 'MargemEsquerda', Trunc(ACBrSATExtratoFortes1.MargemEsquerda));
    Sat_MargemDireita := Ini.ReadInteger('Fortes', 'MargemDireita', Trunc(ACBrSATExtratoFortes1.MargemDireita));
    Sat_Preview := Ini.ReadBool('Fortes', 'Preview', True);

    Sat_Impressora := Ini.ReadString('Printer', 'Name', '');

    Sat_RedeTipoInter := Ini.ReadInteger('Rede', 'tipoInter', 0);
    Sat_RedeTipoLan := Ini.ReadInteger('Rede', 'tipoLan', 0);
    Sat_RedeSSID := Ini.ReadString('Rede', 'SSID', '');
    Sat_RedeSeg := Ini.ReadInteger('Rede', 'seg', 0);
    Sat_RedeCodigo := Ini.ReadString('Rede', 'codigo', '');
    Sat_RedeIP := Ini.ReadString('Rede', 'lanIP', '');
    Sat_RedeMask := Ini.ReadString('Rede', 'lanMask', '');
    Sat_RedeGW := Ini.ReadString('Rede', 'lanGW', '');
    Sat_RedeDNS1 := Ini.ReadString('Rede', 'lanDNS1', '');
    Sat_RedeDNS2 := Ini.ReadString('Rede', 'lanDNS2', '');
    Sat_RedeUsuario := Ini.ReadString('Rede', 'usuario', '');
    Sat_RedeSenha := Ini.ReadString('Rede', 'senha', '');
    Sat_RedeProxy := Ini.ReadInteger('Rede', 'proxy', 0);
    Sat_RedeProxyIP := Ini.ReadString('Rede', 'proxy_ip', '');
    Sat_RedeProxyPorta := Ini.ReadInteger('Rede', 'proxy_porta', 0);
    Sat_RedeProxyUser := Ini.ReadString('Rede', 'proxy_user', '');
    Sat_RedeProxySenha := Ini.ReadString('Rede', 'proxy_senha', '');

    Sat_CfopPadrao := Ini.ReadString('Valores', 'CFOP', '5102');

  finally
    Ini.Free;
  end;

end;

procedure TfrmVenda.ConsultadeClientes1Click(Sender: TObject);
begin
  Imprime_display('Consulta Cliente', clWhite, tiCliente);
  frmCliente_consulta := TfrmCliente_consulta.create(Self);
  frmCliente_consulta.showmodal;
  FreeAndNil(frmCliente_consulta);
  Imprime_display_anterior;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ConsultadePreos1Click(Sender: TObject);
begin
  Imprime_display('Consulta de Preço', clWhite, tiPgto);
  //abre a tela de consulta de preco
  frmPreco_consulta := TfrmPreco_consulta.create(Self);
  frmPreco_consulta.showmodal;
  Imprime_display_anterior;
end;

procedure TfrmVenda.CorEditTotaL;
begin
  if frmPrincipal.TipoImpressora = SemImpressora then begin
    barraAtalho.Color := $006A3500;
  end
  else begin
    barraAtalho.Color := $00134017;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Opes1Click(Sender: TObject);
begin
  // mostrar o panel contendo a legenda dos comandos do menu OPCOES
  pop_principal.Popup(frmVenda.left + 360, frmVenda.top + 330);
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.PesquisadeProdutos1Click(Sender: TObject);
begin
  frmProduto_consulta := TfrmProduto_consulta.create(Self);
  iPesquisa_produto := 0;
  frmProduto_consulta.showmodal;

  if iProd_codigo > 0 then
    ed_barra.Text := IntToStr(iProd_codigo);

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Cupons1Click(Sender: TObject);
begin
  if bVenda then
  begin
    Imprime_display('CUPOM ABERTO!', clYellow, tiAlerta);
    Sleep(1500);
    Imprime_display('INFORME O PRODUTO...', clWhite, tiLivre);
    if not AtivaTouch and ed_barra.Enabled then
      ed_barra.SetFocus;
    Exit;
  end;
  frmCupom_Menu := TfrmCupom_Menu.create(Self);
  frmCupom_Menu.showmodal;
end;

procedure TfrmVenda.cxGrid1DBBandedTableView1FocusedRecordChanged(Sender: TcxCustomGridTableView; APrevFocusedRecord, AFocusedRecord: TcxCustomGridRecord; ANewItemRecordFocusingChanged: Boolean);
begin
  if TipoDeCupom = tcPafECF then
  begin
    if frmModulo.qrItensCANCELADO.AsInteger = 1 then
      btnCancelaItem.Enabled := False;
  end
  else
  begin
    if frmModulo.qrItensCANCELADO.AsInteger = 1 then
      btnCancelaItem.Caption := 'Descancelar Item'
    else
      btnCancelaItem.Caption := 'Cancelar Item';
  end;
end;

procedure TfrmVenda.ImprimirComprovanteCrediario; //qiqi
var
  i: Integer;
  Prestacao, Resto: Double;
begin
  if (bCadastra_Crediario) and (rvalor_total_crediario > 0) and (iComprovente_Crediario = 1) then
  begin
    cdsParcelas.Close;
    cdsParcelas.CreateDataSet;
    Prestacao := RoundTo(rvalor_total_crediario / iCrediario_prestacao, -2);
    Resto := rvalor_total_crediario - (Prestacao * iCrediario_prestacao);

    for i := 1 to iCrediario_prestacao do
    begin
      cdsParcelas.Insert;
      cdsParcelasParcela.AsString := Zerar(IntToStr(i), 2);
      cdsParcelasVencimento.AsDateTime := IncMonth(dData_Sistema, i);
      if i <> iCrediario_prestacao then
        cdsParcelasValor.AsFloat := Prestacao
      else
        cdsParcelasValor.AsFloat := Prestacao + Resto;
      cdsParcelas.Post;
    end;

    if bEdicaoRelatorio then
      frComprovanteDebito.DesignReport
    else
    if bExibeCumponNFTela then
      frComprovanteDebito.ShowReport
    else
    begin
      frComprovanteDebito.PrepareReport(True);
      frComprovanteDebito.Print;
    end;
  end;
end;

procedure TfrmVenda.DevoluoTroca1Click(Sender: TObject);
begin
  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  if frmPrincipal.TipoImpressora = SemImpressora then begin
    Application.MessageBox('Impossivel realizar Troca/Devolução em vendas Fiscais!', 'Atenção!', MB_ICONINFORMATION);
    Exit;
  end;
  CodCliGeraCredito := 0;
  lbDevolucao.Visible := not lbDevolucao.Visible;
  Imprime_display('INFORME O PRODUTO DA DEVOLUÇÃO', clWhite, tiLivre);
  Update;
  Application.CreateForm(TfrmPegaCupomGerencial, frmPegaCupomGerencial);
  frmPegaCupomGerencial.ShowModal;
  if frmPegaCupomGerencial.Confirma then begin
    if not bVenda then begin
      try
        if not Abre_Venda then begin
          bLanca_pre_venda := False;
          bLanca_OS := False;
          bLanca_Mesa := False;
          bLanca_Abastecimento := False;
          Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
          Exit;
        end;
      finally
      end;
    end;
    sProd_unidade := frmPegaCupomGerencial.qrItensUNIDADE.AsString;
    rProd_qtde := frmPegaCupomGerencial.edQtde.Value;
    iProd_codigo := frmPegaCupomGerencial.qrItensCODPRODUTO.AsInteger;
    rProd_preco := frmPegaCupomGerencial.edValor.Value;
    rProd_desconto := 0;
    rProd_acrescimo := 0;
    rProd_total := frmPegaCupomGerencial.edTotal.Value;
    Registra_Item;
    bFezDevolucao := True;
  end;
end;

procedure TfrmVenda.AcionaGaveta1Click(Sender: TObject);
begin
  Gaveta_Abre;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FecharCaixa1Click(Sender: TObject);
begin
  frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
  frmSenha_Supervisor.FechaCaixa := True;
  frmSenha_Supervisor.showmodal;
  frmVenda.Imprime_display_anterior;

  if not bSupervisor_autenticado then
    Exit;

  try
    pn_tampa.Visible := True;
    if bFechamentoCego then begin
      frmFechamentoCego := TfrmFechamentoCego.create(Self);
      frmFechamentoCego.showmodal;
    end
    else begin
      frmCaixa_Fechamento := TfrmCaixa_Fechamento.create(Self);
      frmCaixa_Fechamento.showmodal;
    end;
  finally
    pn_tampa.visible := False;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FormShow(Sender: TObject);
var
  I: Integer;
  tmp, sCupomAbertoECF: string;
  Est: String;
  lValorUnitario: double;
begin
  HabilitarCancelamentoSangriaSuprimento;

  frmVenda.Paint;
  Application.ProcessMessages;
  qrCaixa.Close;
  qrCaixa.Open;
  tmCaixa.Enabled := True;

  vExtra.Texto := '';
  vExtra.valor := 0;

  bFezDevolucao := False;

  if TipoDeCupom = tcSAT then
  begin
    cpBarra.Panels[6].Text := 'Status S@T';
    cpBarra.Panels[6].Width := cpBarra.Panels[6].Width + 15;
    mnEnviaContigencia.Visible := False;
    frmMsg_Operador.lb_msg.caption := 'Aguarde! Inicializando o SAT...';
    frmMsg_Operador.Show;
    frmMsg_Operador.Refresh;
    Application.ProcessMessages;

    ConfiguraSAT;

    ACBrSAT1.Inicializar;
    if ACBrSAT1.Inicializado then
    begin
      frmMsg_Operador.lb_msg.caption := 'Aguarde! Consultando o SAT...';
      frmMsg_Operador.Refresh;
      ACBrSAT1.ConsultarStatusOperacional;
      if ACBrSAT1.Status.ESTADO_OPERACAO = opDESBLOQUEADO then
      begin
        SatOk := True;
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitGreen
      end
      else
      begin
        SatOk := False;
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;
      end;
    end
    else
    begin
      SatOk := False;
      TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;
    end;
  end
  else
  if TipoDeCupom = tcPafECF then
  begin
    cpBarra.Panels[6].Text := 'ECF';
    //cpBarra.Panels[6].Width := cpBarra.Panels[6].Width + 15;
    mmECF.Visible := True;
    frmMsg_Operador.lb_msg.caption := 'Aguarde! Inicializando a ECF...';
    frmMsg_Operador.Show;
    frmMsg_Operador.Refresh;
    Application.ProcessMessages;

    ConfiguraECF;
    try
      cpACBrECF.Ativar;
      if cpACBrECF.Ativo then
      begin
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitGreen;
        try
          Est := Estados[cpACBrECF.Estado];
          cpACBrECF.TestaPodeAbrirCupom;
        except
          Application.MessageBox(PWideChar('Atenção:'+#13+est),'Atenção!',MB_ICONERROR);
        end;
      end
      else
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;

    except
      on e:Exception do
      begin
        TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;
        Application.MessageBox(PWideChar('Ocorreu um erro ao tentar se conectar a ECF!'+#13+e.Message),'Atenção!',MB_ICONINFORMATION);
        Close;
        Exit;
      end;
    end;
  end
  else
  if TipoDeCupom = tcNFCE then
    frmModulo.LerConfiguracao
  else
  if TipoDeCupom = tcSimples then
  begin
    mmSat.Visible := False;
    mnEnviaContigencia.Visible := False;
    G1.Visible := False;
  end;
  frmMsg_Operador.Close;
  frmModulo.qrconfig.Close;
  frmModulo.qrconfig.Open;
  Limpa_controles;
  bMudouProvisorio := False;
  if (TipoDeCupom <> tcSimples) then
  begin
    if frmPrincipal.LerINi(sConfiguracoes, 'PDV', 'Inicia', 'F') = 'F' then
      iImpressora := 0
    else
      iImpressora := 1;

    ImgTipoImpressora(iImpressora);
    CorEditTotaL;
  end;

  timer_cargaTimer(frmVenda);

  Application.CreateForm(TfrmTEF, frmTEF);
  frmTEF.ConfiguraTEF;

  if AtivaTouch then
    MontaTelaTouch;
  Limpa_controles;

  // verificar se existem cupons abertos e fazer o tratamento //**
  sCupomAbertoECF := LerIni(sConfiguracoes, 'PDV', 'Aberto', 'NAO');
  if (sCupomAbertoECF = 'SIM') or (bcupom_aberto) then
  begin
    if TipoDeCupom = tcPafECF then
      cpACBrECF.CancelaCupom;
    if (bBusca_foto_produto) and not (AtivaTouch) then
      img_produto.Picture := nil;

    sNumero_Cupom := LerIni(sConfiguracoes, 'PDV', 'NumCupom', '');
    Abre_Venda(True);
    with frmmodulo, qrItens do
    begin
      iItem := Succ(RecordCount);
      First;
      while not EOF do
      begin
        PrecoDeVenda(qrItensCOD_PRODUTO.Value, qrItensQTDE.Value);
        Next;
      end;
    end;
  end;
end;

procedure TfrmVenda.frComprovanteDebitoBeforePrint(Sender: TfrxReportComponent);

  function FormatarCPFCNPJ(const sDoc: string): string;
  var
    Doc: string;
  begin
    Doc := somenteNumero(sDoc);
    if (Length(Doc) = 11) then begin
      Result := Copy(Doc, 1, 3) + '.' + Copy(Doc, 4, 3) + '.' + Copy(Doc, 7, 3) + '-' + Copy(Doc, 10, 2);
    end;
    if (Length(Doc) = 14) then begin
      Result := Copy(Doc, 1, 2) + '.' + Copy(Doc, 3, 3) + '.' + Copy(Doc, 6, 3) + '/' + Copy(Doc, 9, 4) + '-' + Copy(Doc, 13, 2);
    end;
  end;

begin
  if TfrxMemoView(Sender).Name = 'lbEmpresa' then
    TfrxMemoView(Sender).Text := edtEmitFantasia;
  if TfrxMemoView(Sender).Name = 'lbDocEmpresa' then
    TfrxMemoView(Sender).Text := FormatarCPFCNPJ(edtEmitCNPJ);
  if TfrxMemoView(Sender).Name = 'lbDataHora' then
    TfrxMemoView(Sender).Text := 'DATA: ' + FormatDateTime('DD/MM/YYYY', Date) + ' HORA: ' + FormatDateTime('HH:MM:SS', Time);
  if TfrxMemoView(Sender).Name = 'lbCupom' then
    TfrxMemoView(Sender).Text := Zerar(sNumero_Cupom, 6);
  if TfrxMemoView(Sender).Name = 'lbCliente' then
    TfrxMemoView(Sender).Text := sCli_codigo + '-' + sCli_Nome;
  if TfrxMemoView(Sender).Name = 'lbEndereco' then
    TfrxMemoView(Sender).Text := sCli_Endereco;
  if TfrxMemoView(Sender).Name = 'lbDocCliente' then
    TfrxMemoView(Sender).Text := FormatarCPFCNPJ(sCli_CPF);
  if TfrxMemoView(Sender).Name = 'lbPegoPor' then
    TfrxMemoView(Sender).Text := sCLi_PegoPor;
  if TfrxMemoView(Sender).Name = 'lbPegoPor' then
    TfrxMemoView(Sender).Text := sCLi_PegoPor;
end;

procedure TfrmVenda.frxDesigner1GetTemplateList(List: TStrings);
begin

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Sangria1Click(Sender: TObject);
begin
  if not bVenda then begin
    frmSangria := TfrmSangria.create(Self);
    frmSangria.showmodal;
  end;
end;

procedure TfrmVenda.scGrupoMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  with scGrupo.VertScrollBar do begin
    Position := Position + Increment;
  end;
end;

procedure TfrmVenda.scGrupoMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  with scGrupo.VertScrollBar do begin
    Position := Position - Increment;
  end;
end;

procedure TfrmVenda.btnDimQtdClick(Sender: TObject);
begin
  if ed_qtde.Value > 0 then
    ed_qtde.Value := ed_qtde.Value - 1;
end;

procedure TfrmVenda.btnAumQtdClick(Sender: TObject);
begin
  ed_qtde.Value := ed_qtde.Value + 1;
end;

procedure TfrmVenda.sp_total_pagarContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Suprimento1Click(Sender: TObject);
begin
  if not bVenda then begin
    frmSuprimento := TfrmSuprimento.create(Self);
    frmSuprimento.showmodal;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.rocarOperador1Click(Sender: TObject);
begin
  sFuncao_Senha := '001';
  frmsenha := Tfrmsenha.create(Self);
  frmsenha.showmodal;

  if busuario_autenticado then begin
    // mudar o nome do usuario na tela de venda
    cpBarra.Panels.Items[1].Text := Copy(sNome_Operador, 1, 10);

    Application.MessageBox('Troca efetuada com sucesso!', 'Aviso', MB_OK + MB_ICONINFORMATION);

  end;

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.C1Click(Sender: TObject);
begin
  frmIdentifica := TfrmIdentifica.Create(Self);
  frmIdentifica.ShowModal;
  if (TipoDeCupom <> tcSimples) then begin
    if sConsumidor_CPF <> '' then
      frmPrincipal.TipoImpressora := SemImpressora;
    CorEditTotaL;
  end;
end;

procedure TfrmVenda.C2Click(Sender: TObject);
begin
  if bVenda then begin
    Imprime_display('Venda Aberta!', clRed, tiErro);
    Sleep(1500);
    Imprime_display_anterior;
    Exit;
  end;
  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  bLanca_Comanda := False;

  FComanda := TFComanda.create(Self);
  FComanda.showmodal;
  FreeAndNil(FComanda);

  if bLanca_Comanda then begin
    Imprime_display('Iniciando impressão do Consumo!', clWhite, tiInfo);
    try
      if not Abre_Venda then begin
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
    end;
    LancaMesaComanda(MesasComandasReceber);
  end;

end;

procedure TfrmVenda.C3Click(Sender: TObject);
begin
  Application.CreateForm(TfrmPrincipalChef,frmPrincipalChef);
  frmPrincipalChef.ValidaSistema := False;
  frmPrincipalChef.ShowModal;
  FreeAndNil(frmPrincipalChef);
end;

procedure TfrmVenda.CaixaIndisponvel1Click(Sender: TObject);
begin
  if bVenda then
    Exit;

  frmProtetor_tela := TfrmProtetor_tela.Create(Self);
  frmProtetor_tela.ShowModal;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.menu_PrevendaClick(Sender: TObject);
begin
  if menu_Prevenda.Visible = False then
    Exit;

  if bVenda then begin
    Imprime_display('Venda Aberta!', clRed, tiErro);
    Sleep(1500);
    Imprime_display_anterior;
    Exit;
  end;
  lbEstoque.Visible := False;
  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  bLanca_pre_venda := False;


  LancaRecebimento.ValorDesconto := 0;
  LancaRecebimento.ValorAcressimo := 0;
  frmPre_Venda := TfrmPre_Venda.create(Self);
  frmPre_Venda.showmodal;

  // iniciar a impressao da pre-venda
  if bLanca_pre_venda then begin
    Imprime_display('Iniciando impressão da Pré-venda!', clWhite, tiInfo);
    try
      if not Abre_Venda then begin
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
    end;

    frmModulo.query_servidor.first;
    while not frmModulo.query_servidor.eof do begin
      sProd_nome := frmModulo.query_servidor.fieldbyname('produto').asstring;
      sProd_unidade := frmModulo.query_servidor.fieldbyname('unidade').asstring;
      sProd_CST := frmModulo.query_servidor.fieldbyname('cst').asstring;
      rProd_aliquota := frmModulo.query_servidor.fieldbyname('aliquota').asfloat;
      iProd_codigo := frmModulo.query_servidor.fieldbyname('codproduto').asinteger;
      sProd_barra := frmModulo.query_servidor.fieldbyname('codbarra').asstring;
      rProd_qtde := frmModulo.query_servidor.fieldbyname('qtde').asfloat;
      rProd_preco := frmModulo.query_servidor.fieldbyname('unitario').asfloat;
      rProd_total := frmModulo.query_servidor.fieldbyname('total').asfloat + frmModulo.query_servidor.fieldbyname('desconto').asfloat-frmModulo.query_servidor.fieldbyname('acrescimo').asfloat;
      rProd_desconto := frmModulo.query_servidor.fieldbyname('desconto').asfloat;
      rProd_acrescimo := frmModulo.query_servidor.fieldbyname('acrescimo').asfloat;
      sProd_Tamanho := '';
      sProd_Cor := '';

      Registra_Item;
      Application.ProcessMessages;
      frmModulo.query_servidor.next;
    end;
    FecharCupom1Click(frmVenda);
  end;
end;

procedure TfrmVenda.mmECFClick(Sender: TObject);
begin
  Application.CreateForm(TfrmMenu_Fiscal,frmMenu_Fiscal);
  frmMenu_Fiscal.ShowModal;
  FreeAndNil(frmMenu_Fiscal);
end;


procedure TfrmVenda.AlternarValorVarejo(ADirecao: TDirecao);
var
  lIdxItem: Integer;
  lValorUnitario: Double;
begin
  with frmmodulo do
    for lIdxItem := 0 to High(FValorUintario) do
      if frmmodulo.FValorUintario[lIdxItem].Produto = qrItensCOD_PRODUTO.Value then
      begin
        case ADirecao of
          dAlternar:
            if FValorUintario[lIdxItem].ValorVenda = qrItensVALOR_UNITARIO.Value then
              lValorUnitario := FValorUintario[lIdxItem].ValorDesconto[FValorUintario[lIdxItem].IndiceAtivo]
            else
              lValorUnitario := FValorUintario[lIdxItem].ValorVenda;
          dAlmentar:
            begin
              if FValorUintario[lIdxItem].IndiceAtivo < High(FValorUintario[lIdxItem].ValorDesconto)then
                Inc(FValorUintario[lIdxItem].IndiceAtivo);
              lValorUnitario := FValorUintario[lIdxItem].ValorDesconto[FValorUintario[lIdxItem].IndiceAtivo];
            end;
          dDiminuir:
            begin
              if FValorUintario[lIdxItem].IndiceAtivo > 0 then
                Dec(FValorUintario[lIdxItem].IndiceAtivo);
              lValorUnitario := FValorUintario[lIdxItem].ValorDesconto[FValorUintario[lIdxItem].IndiceAtivo];
            end;
        end;
        AlterarValorUnitario(lValorUnitario, Self);
        lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
        ed_barra.SetFocus;
        Break;
      end;
end;

procedure TfrmVenda.mmiAplicarVarejoClick(Sender: TObject);
begin
  AlternarValorVarejo(dAlternar);
end;

procedure TfrmVenda.mmSatClick(Sender: TObject);
begin
  Application.CreateForm(TfrmCuponsSAT, frmCuponsSAT);
  try
    frmCuponsSAT.ShowModal;
  finally
    FreeAndNil(frmCuponsSAT);
  end;

end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.Oramento1Click(Sender: TObject);
begin
  if Oramento1.visible = False then
    Exit;

  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  bLanca_OS := False;
  frmOrcamento_Abrir := TfrmOrcamento_Abrir.create(Self);
  frmOrcamento_Abrir.showmodal;

  // iniciar a impressao da pre-venda
  if bLanca_OS then begin
    Imprime_display('Iniciando impressão do Cupom!', clWhite, tiInfo);
    try
      // verificar o serial do ecf
      // Abrir o cupom fiscal
      // bloquear ao teclado
      //BlockInput(true);
      if not Abre_Venda then begin
        // Caso o comando de abertura de venda retornou false, abortar o processo;
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
      // desbloquear o teclado
      //BlockInput(False);
    end;

    // vender os itens
    // rodar a query do modulo filtrada com os produtos da pre-venda
    frmModulo.query_servidor.first;
    while not frmModulo.query_servidor.eof do begin
      sProd_nome := frmModulo.query_servidor.fieldbyname('produto').asstring;
      sProd_unidade := frmModulo.query_servidor.fieldbyname('unidade').asstring;
      sProd_CST := frmModulo.query_servidor.fieldbyname('cst').asstring;
      rProd_aliquota := frmModulo.query_servidor.fieldbyname('aliquota').asfloat;
      iProd_codigo := frmModulo.query_servidor.fieldbyname('codproduto').asinteger;
      sProd_barra := frmModulo.query_servidor.fieldbyname('codbarra').asstring;
      rProd_qtde := frmModulo.query_servidor.fieldbyname('qtde').asfloat;
      rProd_preco := frmModulo.query_servidor.fieldbyname('unitario').asfloat;
      rProd_total := frmModulo.query_servidor.fieldbyname('total').asfloat;
      rProd_desconto := frmModulo.query_servidor.fieldbyname('desconto').asfloat;
      rProd_acrescimo := frmModulo.query_servidor.fieldbyname('acrescimo').asfloat;
      sProd_Tamanho := '';
      sProd_Cor := '';

      // registrar o item
      Registra_Item;
      Application.ProcessMessages;
      frmModulo.query_servidor.next;
    end;

    query.CLOSE;
    query.SQL.CLEAR;
    query.SQL.ADD('UPDATE DAV SET COO_CUPOM_FISCAL = ''' + sNumero_Cupom + ''',');
    query.SQL.ADD('CCF = ''' + sNumero_contador_cupom + ''' where numero = ''' + SDAV_ATUALIZADO + '''');
    query.ExecSQL;

    frmModulo.query_servidor.close;
    frmModulo.query_servidor.sql.clear;
    frmModulo.query_servidor.sql.add('update c000074 set tipo = 10 where codigo = ''' + Zerar(IntToStr(iOS_codigo), 6) + '''');
    frmModulo.query_servidor.ExecSQL;

    FecharCupom1Click(frmVenda);
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.AdvMetroButton1Click(Sender: TObject);
begin
  Sair1.Click;
end;

procedure TfrmVenda.AdvMetroButton2Click(Sender: TObject);
begin
  Application.Minimize;
end;

procedure TfrmVenda.AjustaDescontoAcrescimo(Desconto, Acrescimo: Double);
var
  TotDes, TotAcr: Double;
begin
  if (Desconto <> 0) or (Acrescimo <> 0) then begin
    with frmModulo do begin
      TotDes := 0;
      TotAcr := 0;
      qrItemAJ.Close;
      qrItemAJ.Open;
      qrItemAJ.First;
      while not qrItemAJ.Eof do begin
        qrItemAJ.Edit;
        if Desconto > 0 then
          qrItemAJVALOR_DESCONTO.AsFloat := qrItemAJVALOR_DESCONTO.AsFloat + RoundTo(((qrItemAJVALOR_TOTAL.AsFloat / rTotal_Venda) * Desconto), -2);
        if Acrescimo > 0 then
          qrItemAJVALOR_ACRESCIMO.AsFloat := qrItemAJVALOR_ACRESCIMO.AsFloat + RoundTo(((qrItemAJVALOR_TOTAL.AsFloat / rTotal_Venda) * Acrescimo), -2);
        qrItemAJVALOR_TOTAL.AsFloat := qrItemAJVALOR_SUBTOTAL.AsFloat - qrItemAJVALOR_DESCONTO.AsFloat + qrItemAJVALOR_ACRESCIMO.AsFloat;
        TotDes := TotDes + qrItemAJVALOR_DESCONTO.AsFloat;
        TotAcr := TotAcr + qrItemAJVALOR_ACRESCIMO.AsFloat;
        qrItemAJ.Post;
        qrItemAJ.Next;
      end;
      qrItens.Close;
      qrItens.Open;
    end;
  end;
end;

procedure TfrmVenda.mniAlterarValorUnitario1Click(Sender: TObject);
var
  lCorOriginal: TColor;
  lMsgOriginal: string;
begin
  if frmModulo.qrItensVALOR_UNITARIO.Value > 0 then
  begin
    lMsgOriginal := lb_produto.Caption;
    lCorOriginal := lb_produto.Color;
    Imprime_display('PERMISSÃO PARA ALTERAR VALOR UNITÁRIO!', clRed, tiLivre);
    ///
    frmsenha_supervisor := tfrmsenha_supervisor.create(self);
    frmsenha_supervisor.ShowModal;
    if not bSupervisor_autenticado then
    begin
      frmsenha_supervisor.Free;
      ShowMessage('Senha inválida!');
      Imprime_display(lMsgOriginal, lCorOriginal, tiLivre);
      ed_barra.SetFocus;
      Exit;
    end;
    FreeAndNil(frmsenha_supervisor);
    ///
    ExibirPlanoFundo;
    frmAlterarValorUnitario := TfrmAlterarValorUnitario.Create(Self, TAcao.aValorUnitario);
    frmAlterarValorUnitario.showmodal;
    lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
    FecharPlanoFundo;
    Imprime_display(lMsgOriginal, lCorOriginal, tiLivre);
    ed_barra.SetFocus;
  end;
end;
//begin
//  if frmModulo.qrItensVALOR_UNITARIO.Value > 0 then
//  begin
//    Imprime_display('PERMISSÃO PARA ALTERAR VALOR UNITÁRIO!', clRed, tiLivre);
//    ///
//    frmsenha_supervisor := tfrmsenha_supervisor.create(self);
//    frmsenha_supervisor.ShowModal;
//    if not bSupervisor_autenticado then
//    begin
//      ShowMessage('Senha inválida!');
//      ed_barra.SetFocus;
//      Exit;
//    end;
//    FreeAndNil(frmsenha_supervisor);
//    ///
//    ExibirPlanoFundo;
//    frmAlterarValorUnitario := TfrmAlterarValorUnitario.create(Self, TAcao.aValorUnitario);
//    frmAlterarValorUnitario.showmodal;
//    lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
//    FecharPlanoFundo;
//    ed_barra.SetFocus;
//  end;
//end;

procedure TfrmVenda.mniCancelarSangriaOuSuprimentoClick(Sender: TObject);
var
  lCorOriginal: TColor;
  lValor,
  lMsgOriginal,
  lMsgDisplay: string;
begin
  lMsgOriginal := lb_produto.Caption;
  lCorOriginal := lb_produto.Color;
  if Sender = mniCancelarUltimaSangria then
    lMsgDisplay := 'Sangria'
  else
    lMsgDisplay := 'Suprimento';
  Imprime_display('PERMISSÃO PARA CANCELAR ' + lMsgDisplay.ToUpper + '!', clRed, tiLivre);
  ///
  frmsenha_supervisor := tfrmsenha_supervisor.create(self);
  frmsenha_supervisor.ShowModal;
  if not bSupervisor_autenticado then
  begin
    ShowMessage('Senha inválida!');
    Imprime_display(lMsgOriginal, lCorOriginal, tiLivre);
    ed_barra.SetFocus;
    Exit;
  end;
  FreeAndNil(frmsenha_supervisor);
  ///
  uQry.Close;
  uQry.sql.Clear;
  uQry.sql.Add('select first 1 * from NAO_FISCAL where ECF = ' + QuotedStr(sCaixa) +
               ' AND descricao = ' + QuotedStr(lMsgDisplay.ToUpper) +
               ' and data = Current_Date order by data desc, hora desc');
  uQry.Open;
  if uQry.RecordCount > 0 then
  begin
    try
      uQry.Delete;
      Application.MessageBox(PWideChar(lMsgDisplay + ' cancelada com sucesso!'), 'Atenção!', MB_ICONINFORMATION);
    except
      Application.MessageBox(PWideChar('Ocorreu um erro ao cancelar ' + lMsgDisplay + '!'#13' Operação cancelada!'), 'Atenção!', MB_ICONINFORMATION);
    end;
    TMenuItem(Sender).Enabled := uQry.RecordCount > 0;
  end;
end;

procedure TfrmVenda.AplicaDescontoPorGrupo;
var
  Grupo, Subgrupo, CodFornbecedor, CodMarca: string;
  Itens, Desconto, Indice, QTD:Double;
  IDtabela:Integer;
  Verifica, SDGrupo, SDSubgrupo, SDFornecedor, SDMarca:Boolean;
begin
  with frmModulo do begin
    qrTabelaPrecoGrupo.Close;
    qrTabelaPrecoGrupo.SQL.Clear;
    qrTabelaPrecoGrupo.SQL.Add('select * from TABELA_PRECO_GRUPO');
    qrTabelaPrecoGrupo.Open;
    if not qrTabelaPrecoGrupo.IsEmpty then begin
      qrDescontoGrupo.Close;
      qrDescontoGrupo.Open;
      qrDescontoGrupo.First;
      while not qrDescontoGrupo.Eof do begin
        Grupo := Trim(qrDescontoGrupoGRUPO.AsString);
        SubGrupo := Trim(qrDescontoGrupoCODSUBGRUPO.AsString);
        CodFornbecedor := Trim(qrDescontoGrupoCODFORNECEDOR.AsString);
        SubGrupo := Trim(qrDescontoGrupoCODMARCA.AsString);
        SDGrupo := (qrDescontoGrupoSINALIZADOR_DESCONTO.Value and 1) > 0;
        SDSubgrupo := (qrDescontoGrupoSINALIZADOR_DESCONTO.Value and 2) > 0;
        SDFornecedor := (qrDescontoGrupoSINALIZADOR_DESCONTO.Value and 4) > 0;
        SDMarca := (qrDescontoGrupoSINALIZADOR_DESCONTO.Value and 8) > 0;


        if Grupo <> '' then begin
          QTD := 0;
          while not(qrDescontoGrupo.Eof) and (Grupo = Trim(qrDescontoGrupoGRUPO.AsString)) do begin
            QTD := QTD + qrDescontoGrupoQTDE.AsFloat;
            qrDescontoGrupo.Next;
          end;

          qrTabelaPrecoGrupo.Close;
          qrTabelaPrecoGrupo.SQL.Clear;
          qrTabelaPrecoGrupo.SQL.Add('select * from TABELA_PRECO_GRUPO tab where tab.COD_GRUPO = ' + QuotedStr(Grupo) + ' order by tab.quantidade_inicial');
          qrTabelaPrecoGrupo.Open;
          if not qrTabelaPrecoGrupo.IsEmpty then begin
            qrTabelaPrecoGrupo.First;
            if qrTabelaPrecoGrupoTIPO_TABELA.AsString = 'I' then begin
              while not qrTabelaPrecoGrupo.Eof do begin
                if (QTD >= qrTabelaPrecoGrupoQUANTIDADE_INICIAL.AsFloat) and (QTD <= qrTabelaPrecoGrupoQUANTIDADE_FINAL.AsFloat) then begin
                  Indice := qrTabelaPrecoGrupoINDICE.AsFloat;
                  qrItens.Filter := 'cancelado = 0 and grupo = ' + QuotedStr(Grupo);
                  qrItens.Filtered := True;
                  qrItens.First;
                  while not qrItens.Eof do begin
                    qrItens.Edit;
                    if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'P' then
                      Desconto := RoundTo(qrItensVALOR_UNITARIO.AsFloat - Indice,-2);
                    if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'V' then
                      Desconto := Indice;
                    if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'D' then
                      Desconto := RoundTo((qrItensVALOR_UNITARIO.AsFloat * Indice)/100,-2);
                    Desconto := RoundTo(qrItensQTDE.AsFloat * Desconto,-2);
                    qrItensVALOR_DESCONTO.AsFloat := RoundTo(qrItensVALOR_DESCONTO.AsFloat + Desconto,-2);
                    qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat - Desconto,-2);
                    rTotal_Desconto := RoundTo(rTotal_Desconto + Desconto,-2);
                    rTotal_Venda := RoundTo(rTotal_Venda - Desconto,-2);
                    qrItensDESCONTO_CUPOM.AsFloat := RoundTo(qrItensDESCONTO_CUPOM.AsFloat + Desconto,-2);
                    qrItens.Post;
                    qrItens.Next;
                  end;
                  Break;
                end;
                qrTabelaPrecoGrupo.Next;
              end;
            end else if qrTabelaPrecoGrupoTIPO_TABELA.AsString = 'G' then begin
              Verifica := True;
              while Verifica do begin
                IDtabela := 0;
                qrTabelaPrecoGrupo.First;
                while not qrTabelaPrecoGrupo.Eof do begin
                  if (QTD >= qrTabelaPrecoGrupoQUANTIDADE_INICIAL.AsFloat) then
                    IDtabela := qrTabelaPrecoGrupoID.AsInteger;
                  qrTabelaPrecoGrupo.Next;
                end;
                if IDtabela > 0 then begin
                  Verifica := True;
                  qrTabelaPrecoGrupo.Locate('ID',IDtabela,[loCaseInsensitive]);
                  Itens := 0;
                  Itens := (Trunc(QTD / qrTabelaPrecoGrupoQUANTIDADE_INICIAL.AsFloat))*qrTabelaPrecoGrupoQUANTIDADE_INICIAL.AsFloat;
                  QTD := QTD - Itens;
                  Indice := qrTabelaPrecoGrupoINDICE.AsFloat;
                  qrItens.Filter := 'cancelado = 0 and grupo = ' + QuotedStr(Grupo);
                  qrItens.Filtered := True;
                  qrItens.First;
                  while not(qrItens.Eof) and (Itens > 0) do begin
                    if qrItensAPLICADO_DESCONTO_GRUPO.AsString <> 'S' then begin
                      qrItens.Edit;
                      if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'P' then
                        Desconto := RoundTo(qrItensVALOR_UNITARIO.AsFloat - Indice,-2);
                      if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'V' then
                        Desconto := Indice;
                      if qrTabelaPrecoGrupoTIPO_DESCONTO.AsString = 'D' then
                        Desconto := RoundTo((qrItensVALOR_UNITARIO.AsFloat * Indice)/100,-2);
                      if Itens >= (qrItensQTDE.AsFloat - qrItensQTD_APLICADO_DESCONTO.AsFloat) then
                        Desconto := RoundTo((qrItensQTDE.AsFloat - qrItensQTD_APLICADO_DESCONTO.AsFloat) * Desconto,-2)
                      else
                        Desconto := RoundTo(Itens * Desconto,-2);
                      qrItensVALOR_DESCONTO.AsFloat := RoundTo(qrItensVALOR_DESCONTO.AsFloat + Desconto,-2);
                      qrItensVALOR_TOTAL.AsFloat := RoundTo(qrItensVALOR_TOTAL.AsFloat - Desconto,-2);
                      rTotal_Desconto := RoundTo(rTotal_Desconto + Desconto,-2);
                      rTotal_Venda := RoundTo(rTotal_Venda - Desconto,-2);
                      qrItensDESCONTO_CUPOM.AsFloat := RoundTo(qrItensDESCONTO_CUPOM.AsFloat + Desconto,-2);
                      if Itens >= (qrItensQTDE.AsFloat - qrItensQTD_APLICADO_DESCONTO.AsFloat) then begin
                        Itens := Itens - (qrItensQTDE.AsFloat - qrItensQTD_APLICADO_DESCONTO.AsFloat);
                        qrItensAPLICADO_DESCONTO_GRUPO.AsString := 'S';
                        qrItensQTD_APLICADO_DESCONTO.AsFloat := (qrItensQTDE.AsFloat - qrItensQTD_APLICADO_DESCONTO.AsFloat);
                      end else begin
                        qrItensQTD_APLICADO_DESCONTO.AsFloat := Itens;
                        Itens := 0;
                      end;
                      qrItens.Post;
                    end;
                    qrItens.Next;
                  end;
                  qrItens.Filter := '';
                  qrItens.Filtered := False;
                end else
                  Verifica := False;
              end;
            end;
          end;
        end else
          qrDescontoGrupo.Next;
      end;
    end;
  end;
end;

procedure TfrmVenda.Recebimento1Click(Sender: TObject);
begin
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
  if not PDV_OnLine then
    Exit;


//  frmContasReceber := tfrmcontasReceber.create(self);
  Application.CreateForm(Tfrmcontasreceber, frmcontasreceber);
  frmcontasreceber.showmodal;
end;
// -------------------------------------------------------------------------- //
procedure TfrmVenda.menu_mesaClick(Sender: TObject);
var
  qtdCouver, nItem: Integer;
  TotVen, ValorItem, TotDestaVenda, vlrAcressimoTaxa, vlrAcressimoCouver, Maior, ppValor: Double;
  ppCodigo: string;
begin
  if menu_mesa.Visible = False then
    Exit;

  if bVenda then begin
    Imprime_display('Venda Aberta!', clRed, tiErro);
    Sleep(1500);
    Imprime_display_anterior;
    Exit;
  end;

  lbEstoque.Visible := False;
  PDV_OnLine := True;
  if not frmModulo.Conexao_Servidor.Connected then begin
    try
      frmModulo.Conexao_Servidor.Connected := False;
      frmModulo.Conexao_Servidor.Connected := True;
    except
      on E: Exception do begin
        PDV_OnLine := False;
        Application.messagebox('Não foi possivel conectar ao servidor!' + #13 + 'Para ter acesso a esta opção é necessário conexão com o Servidor!', 'Erro', MB_ICONERROR);
      end;
    end;
  end;
  if not PDV_OnLine then
    Exit;

  bLanca_Mesa := False;

  frmMesas := TfrmMesas.create(Self);
  frmMesas.showmodal;

  if bLanca_Mesa then begin
    Imprime_display('Iniciando impressão do Consumo!', clWhite, tiInfo);
    try
      if not Abre_Venda then begin
        Imprime_display('ERRO AO ABRIR O CUPOM!', clRed, tiErro);
        Exit;
      end;
    finally
    end;
    LancaMesaComanda(MesasComandasReceber);
  end;
end;

procedure TfrmVenda.tmCaixaTimer(Sender: TObject);
begin
  tmCaixa.Enabled := False;
  if (qrCaixa.fieldbyname('caixa_situacao').asstring = 'ABERTO') and (qrCaixa.fieldbyname('caixa_data_movto').AsDateTime <> Date) and (qrCaixa.fieldbyname('FECHAMENTO_AUTOMATICO_CAIXA_DIA').AsString = 'S') and not(bVenda) then begin
    try
      pn_tampa.Visible := True;
      if bFechamentoCego then begin
        frmFechamentoCego := TfrmFechamentoCego.create(Self);
        frmFechamentoCego.showmodal;
      end else begin
        Imprime_display('AGUARDE FECHANDO O CAIXA!', clRed, tiErro);
        frmCaixa_Fechamento := TfrmCaixa_Fechamento.create(Self);
        frmCaixa_Fechamento.Pergunta := True;
        frmCaixa_Fechamento.AdvGlowButton2.Click;
        frmCaixa_Fechamento.AdvGlowButton1.Click;
      end;
    finally
      pn_tampa.visible := False;
    end;
  end else
    tmCaixa.Enabled := True;
end;

procedure TfrmVenda.tmContigenciaTimer(Sender: TObject);
begin
  if TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType = sitRed then
    TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitYellow
  else
    TdxStatusBarStateIndicatorPanelStyle(cpBarra.Panels[7].PanelStyle).Indicators[0].IndicatorType := sitRed;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if wspnlAlerta.Left = 0 then
  begin
    if (Key = VK_RETURN) and (FLastKeyPress = VK_RETURN) then
      AlertaFecharClick(nil)
    else
      FLastKeyPress := Key;
    Exit;
  end;
  FLastKeyPress := 0;
end;

procedure TfrmVenda.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    ed_barra.SetFocus;
end;

procedure TfrmVenda.FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key in [107, 187] then // +
  begin
    AlternarValorVarejo(dAlmentar);
  end
  else
  if Key in [109, 189] then // -
  begin
    AlternarValorVarejo(dDiminuir);
  end
  else
  if (Key = 66) and (Shift = [ssCtrl]) then begin
    if (sConsumidor_CPF <> '') and (frmPrincipal.TipoImpressora = SemImpressora) then begin
      Application.MessageBox('Impossivel Realizar a ação CPF do Cliente já Informado!', 'Erro', MB_OK + MB_ICONERROR);
      Exit;
    end;
    if (bFezDevolucao) then begin
      Application.MessageBox('Impossivel Realizar a ação com Devolução lançada no Cupom!', 'Erro', MB_OK + MB_ICONERROR);
      Exit;
    end;
    if (TipoDeCupom <> tcSimples) then begin
      if frmPrincipal.TipoImpressora = SemImpressora then
        frmPrincipal.TipoImpressora := NaoFiscal
      else
        frmPrincipal.TipoImpressora := SemImpressora;
    end;
    CorEditTotaL;
  end
  else
  if (Key = 88) and (Shift = [ssCtrl]) then begin
    if (TEF_Ativo) then begin
      frmTEF.AbreGerencial;
    end else begin
      ShowMessage('TEF não esta ativo nas configurações do PDV!');
    end;
  end;
end;

procedure TfrmVenda.timer_cargaTimer(Sender: TObject);
begin
  if frmModulo.Conexao_Servidor.Connected then
    cpBarra.Panels[5].Text := 'ON-LINE'
  else
    cpBarra.Panels[5].Text := 'OFF-LINE';
  frmModulo.query.close;
  frmModulo.query.sql.clear;
  frmModulo.query.SQL.add('select carga_data, carga_hora from config');
  frmModulo.query.open;
  frmModulo.qrconfig.open;
  cpBarra.Panels.Items[11].Text := frmModulo.query.fieldbyname('carga_data').asstring + ' às ' + frmModulo.query.fieldbyname('carga_hora').asstring;
end;

// -------------------------------------------------------------------------- //

procedure TfrmVenda.ed_unitarioKeyPress(Sender: TObject; var Key: Char);
var
  rd: Real;
begin
  if Key = #13 then begin
    if ed_unitario.value <= 0 then begin
      ed_unitario.SetFocus;
      Exit;
    end;
    if (FormatFloat('##########0.00', rProd_preco) <> FormatFloat('##########0.00', ed_unitario.value)) and (rProd_preco > 0) then begin
      // verificar seh eh desconto ou acrescimo
      if rProd_preco < ed_unitario.value then begin
        // foi acrescimo
        rProd_acrescimo := (ed_unitario.value - rProd_preco);
        if rProd_acrescimo_maximo > 0 then begin
          rd := (rProd_acrescimo / rProd_preco) * 100;
          if rd > rProd_acrescimo_maximo then begin
            frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
            frmSenha_Supervisor.showmodal;
            if not bSupervisor_autenticado then
              Exit;
          end;
        end
        else begin
          frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
          frmSenha_Supervisor.showmodal;
          if not bSupervisor_autenticado then
            Exit;
        end;
        rProd_acrescimo := rProd_acrescimo * rProd_qtde;
      end
      else begin
        // foi desconto
        rProd_desconto := rProd_preco - ed_unitario.value;
        if rProd_desconto_maximo > 0 then begin
          rd := (rProd_desconto / rProd_preco) * 100;
          if rd > rProd_desconto_maximo then begin
            frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
            frmSenha_Supervisor.ValidaDesconto := True;
            frmSenha_Supervisor.showmodal;
            if not bSupervisor_autenticado then
              Exit;
          end;
        end
        else begin
          frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
          frmSenha_Supervisor.ValidaDesconto := True;
          frmSenha_Supervisor.showmodal;
          if not bSupervisor_autenticado then
            Exit;
        end;
        rProd_desconto := rProd_desconto * rProd_qtde;
      end;
      //rProd_preco := ed_unitario.value;
      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
      ed_total_item.Value := rProd_total;
    end;
    if rProd_preco = 0 then begin
      rProd_preco := ed_unitario.value;
      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
      ed_total_item.Value := rProd_total;
    end;
    if (bMuda_total) then begin
      Imprime_display(sProd_nome, clWhite, tiInfo);
      ed_total_item.ReadOnly := False;
      ed_total_item.setfocus;
      Exit;
    end;
    Registra_Item;
    if bMudouProvisorio then begin
      bMuda_unitario := False;
      bMudouProvisorio := False;
    end;
  end
  else if Key = #27 then begin
    Teclado.Visible := False;
  end;
end;

function TfrmVenda.ImgTipoImpressora(i: Integer): TImpressora;
begin
  case i of
    0:
      frmPrincipal.TipoImpressora := SemImpressora;
    1:
      frmPrincipal.TipoImpressora := NaoFiscal;
    2:
      frmPrincipal.TipoImpressora := Fiscal;
  end;
  result := frmPrincipal.TipoImpressora;
end;

procedure TfrmVenda.MontaGrupos;
var
  i: Integer;
  cpPanel: TPanel;
  cpImage: TImage;
  cpLabel: TLabel;
  Altura, Conta, TColuna, TItens, TLeft, TTop: Integer;
  MemoryStream: TMemoryStream;
  Bitmap: TJPEGImage;
  CodCrupo: Integer;
begin
  TColuna := RoundDiv(scGrupo.Width, Touch_Colunas);
  if Touch_linhas_Grupo > 1 then begin
    pnTouch_Grupo.Height := Touch_Altura_Grupo + Panel1.Height;
  end;
  qrGrupo.Close;
  qrGrupo.Open;
  i := 0;
  TLeft := 0;
  TTop := 0;
  Conta := 0;
  TItens := RoundDiv(qrGrupo.RecordCount, Touch_linhas_Grupo);
  Altura := RoundDiv(scGrupo.Height - Panel1.Height, Touch_linhas_Grupo);
  while not qrGrupo.Eof do begin
    CodCrupo := qrGrupoCODIGO.AsInteger;
    Inc(i);
    Inc(Conta);
    cpPanel := TPanel.Create(scGrupo);
    with cpPanel do begin
      Name := 'pnTouch_Grupo' + IntToStr(i);
      Parent := scGrupo;
      if Touch_linhas_Grupo > 1 then begin
        Left := TLeft;
        Top := TTop;
        Width := TColuna;
        Height := Altura;
        Align := alNone;
      end
      else begin
        Left := 0;
        Top := 0;
        Width := TColuna;
        Height := 86;
        Align := alLeft;
      end;
      Caption := '';
      TabOrder := i - 1;
      Color := clWhite;
      Tag := qrGrupoCODIGO.AsInteger;
      BorderWidth := 2;
      OnClick := GrupoClick;
      ParentBackground := False;
      if Touch_linhas_Grupo > 1 then begin
        TLeft := TLeft + TColuna;
        if Conta = TItens then begin
          TLeft := 0;
          TTop := TTop + RoundDiv(scGrupo.Height - Panel1.Height, Touch_linhas_Grupo);
          Conta := 0;
        end;
      end;
    end;
    cpLabel := TRzLabel.Create(cpPanel);
    with cpLabel do begin
      Name := 'lbTouch_Grupo' + IntToStr(i);
      Parent := cpPanel;
      Height := 16;
      Align := alBottom;
      Caption := qrGrupoGRUPO.AsString;
      Alignment := taCenter;
      TRANSPARENT := True;
      Tag := qrGrupoCODIGO.AsInteger;
      OnClick := GrupoClick;
    end;
    if not qrGrupoFOTO.IsNull then begin
      cpImage := TImage.Create(cpPanel);
      Bitmap := TJPEGImage.Create;
      MemoryStream := TMemoryStream.Create;
      qrGrupoFOTO.savetostream(MemoryStream);
      MemoryStream.Position := OffsetMemoryStream;
      Bitmap.LoadFromStream(MemoryStream);
      with cpImage do begin
        Name := 'imTouch_Grupo' + IntToStr(i);
        Parent := cpPanel;
        Align := alClient;
        Picture.Assign(Bitmap);
        Center := True;
        Stretch := False;
        Tag := qrGrupoCODIGO.AsInteger;
        OnClick := GrupoClick;
        TRANSPARENT := True;
      end;
      FreeAndNil(MemoryStream);
    end;
    qrGrupo.Next;
  end;
  if TItens <= Touch_Colunas then begin
    while (scGrupo.Height) > (Altura * Touch_linhas_Grupo) do begin
      pnTouch_Grupo.Height := pnTouch_Grupo.Height - 1;
//      scGrupo.Height := scGrupo.Height - 1;
    end;
  end;
  MontaProdutos(CodCrupo);
end;

procedure TfrmVenda.MontaProdutos(Grupo: Integer);
var
  i: Integer;
  cpPanel: TPanel;
  cpImage: TImage;
  cpLabel: TLabel;
  Conta, TColuna, TLinha, TLeft, TTop: Integer;
  MemoryStream: TMemoryStream;
  Bitmap: TJPEGImage;
begin
  for i := scProduto.ControlCount - 1 downto 0 do
    TComponent(scProduto.Controls[i]).Free;
  TColuna := RoundDiv(scProduto.Width, Touch_Colunas);
  TLinha := RoundDiv(scProduto.Height, Touch_Linhas);
  qrProduto.Close;
  qrProduto.Params.Items[0].Value := Zerar(IntToStr(Grupo), 6);
  qrProduto.Open;
  i := 0;
  TLeft := 0;
  TTop := 0;
  Conta := 0;
  while not qrProduto.Eof do begin
    Inc(i);
    Inc(Conta);
    cpPanel := TPanel.Create(scProduto);
    with cpPanel do begin
      Name := 'pnTouch_Produto' + IntToStr(i);
      Parent := scProduto;
      Left := TLeft;
      Top := TTop;
      Width := TColuna;
      Height := TLinha;
      Align := alNone;
      Caption := '';
      TabOrder := i - 1;
      Color := clWhite;
      Tag := qrProdutoCODIGO.AsInteger;
      BorderWidth := 2;
      OnClick := ProdutoClick;
      OnMouseDown := MouseDown;
      OnMouseLeave := MouseLeave;
      ParentBackground := False;
      TLeft := TLeft + TColuna;
      if Conta = Touch_Colunas then begin
        TLeft := 0;
        TTop := TTop + TLinha;
        Conta := 0;
      end;
    end;
    cpLabel := TRzLabel.Create(cpPanel);
    with cpLabel do begin
      Name := 'lbTouch_Produto' + IntToStr(i);
      Parent := cpPanel;
      Height := 16;
      Align := alBottom;
      Caption := qrProdutoNOME.AsString;
      Alignment := taCenter;
      TRANSPARENT := True;
      Tag := qrProdutoCODIGO.AsInteger;
      OnClick := ProdutoClick;
      OnMouseDown := MouseDown;
      OnMouseLeave := MouseLeave;
    end;
    if not qrProdutoFOTOBD.IsNull then begin
      cpImage := TImage.Create(cpPanel);
      Bitmap := TJPEGImage.Create;
      MemoryStream := TMemoryStream.Create;
      qrProdutoFOTOBD.savetostream(MemoryStream);
      MemoryStream.Position := OffsetMemoryStream;
      Bitmap.LoadFromStream(MemoryStream);
      with cpImage do begin
        Name := 'imTouch_Produto' + IntToStr(i);
        Parent := cpPanel;
        Align := alClient;
        Picture.Assign(Bitmap);
        Center := True;
        Stretch := False;
        Tag := qrProdutoCODIGO.AsInteger;
        OnClick := ProdutoClick;
        OnMouseDown := MouseDown;
        OnMouseLeave := MouseLeave;
        TRANSPARENT := True;
      end;
      FreeAndNil(MemoryStream);
    end;
    qrProduto.Next;
  end;

end;

procedure TfrmVenda.MontaTelaTouch;
var
  i: Integer;
  cpPanel: TPanel;
  cpImage: TImage;
  cpLabel: TLabel;
  TColuna: Integer;
  TLinha: Integer;
  MemoryStream: TMemoryStream;
  Bitmap: TJPEGImage;
begin
  cxGrid1DBBandedTableView1COD_BARRA.Visible := False;
  btnCancelaVenda.Visible := True;
  btnCancelaItem.Visible := True;
  btnFechaCupom.Visible := True;
  Repaint;
  Update;
  Refresh;
  pnBotoes.Visible := True;
  pnBotoesInferior.Visible := True;
  v_tipo_touch.Visible := True;
  v_tipo_normal.Visible := False;
  pnBarra.Parent := pnBotoes;

  pnQuantidade.Parent := pnBotoes;
  pnQuantidade.Left := 0;

  for i := scGrupo.ControlCount - 1 downto 0 do
    TComponent(scGrupo.Controls[i]).Free;
  for i := scProduto.ControlCount - 1 downto 0 do
    TComponent(scProduto.Controls[i]).Free;
  btnDimQtd.Visible := True;
  btnAumQtd.Visible := True;
  MontaGrupos;
end;

procedure TfrmVenda.MostraFotoProduto(b: Boolean);
begin
  if (bBusca_foto_produto) and not (AtivaTouch) then
    if b then begin
      img_produto.Visible := True;
    end
    else begin
      img_produto.Visible := False;
    end;
end;

procedure TfrmVenda.MostraLogoMarca(b: Boolean);
begin
  if not AtivaTouch then begin
    if bHabLogoMarca then
      img_logo.Visible := b;

    if bHabLogoMarca then
      if b then
        if FileExists(EditLogoMarca) then
          img_logo.Picture.LoadFromFile(EditLogoMarca);
  end
  else
    img_logo.Visible := False;

end;

procedure TfrmVenda.MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Controle := TControl(Sender);
  TimerExtra.Enabled := True;
end;

procedure TfrmVenda.MouseLeave(Sender: TObject);
begin
  Controle := nil;
  TimerExtra.Enabled := False;
end;

function TfrmVenda.TemDevolucao: Boolean;
begin
  with frmModulo do begin
    try
      Result := False;
      qrItens.DisableControls;
      qrItens.First;
      while not qrItens.Eof do begin
        if qrItensQTDE.AsFloat < 0 then begin
          Result := True;
          Break;
        end;
        qrItens.Next;
      end;
      qrItens.Last;
    finally
      qrItens.EnableControls;
    end;
  end;
end;

procedure TfrmVenda.TimerExtraTimer(Sender: TObject);
begin
  LongPress;
end;

procedure TfrmVenda.TimerTrocoTimer(Sender: TObject);
begin
  TimerTroco.Enabled := False;

  Imprime_display('C A I X A    L I V R E', clWhite, tiLivre);

end;

procedure TfrmVenda.CentralizarPanel(p: TPanel);
begin
  p.Top := (Height - p.Height) div 2;
  p.Left := (Width - p.Width) div 2;
end;

procedure TfrmVenda.colBotAlmuntarPropertiesButtonClick(Sender: TObject; AButtonIndex: Integer);
var
  local: Integer;
  Valor: Double;
begin
  Application.ProcessMessages;
  query.SQL.Clear;
  Inc(iTotal_Itens);
  Valor := frmModulo.qrItensVALOR_UNITARIO.AsFloat;
  rTotal_Venda := rTotal_Venda + Valor;
  with frmModulo do begin
    qrItens.Edit;
    qrItensQTDE.AsFloat := qrItensQTDE.AsFloat + 1;
    qrItensVALOR_SUBTOTAL.AsFloat := qrItensVALOR_SUBTOTAL.AsFloat + Valor;
    qrItensVALOR_TOTAL.AsFloat := qrItensVALOR_TOTAL.AsFloat + Valor;
    qrItens.Post;
  end;

  lb_item.Caption := IntToStr(iTotal_Itens);
  lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);

end;

procedure TfrmVenda.colBotDiminuirPropertiesButtonClick(Sender: TObject; AButtonIndex: Integer);
var
  local: Integer;
  Valor: Double;
begin
  if frmModulo.qrItensQTDE.AsInteger > 1 then begin
    Application.ProcessMessages;
    query.SQL.Clear;
    Dec(iTotal_Itens);
    Valor := frmModulo.qrItensVALOR_UNITARIO.AsFloat;
    rTotal_Venda := rTotal_Venda - Valor;
    with frmModulo do begin
      qrItens.Edit;
      qrItensQTDE.AsFloat := qrItensQTDE.AsFloat - 1;
      qrItensVALOR_SUBTOTAL.AsFloat := qrItensVALOR_SUBTOTAL.AsFloat - Valor;
      qrItensVALOR_TOTAL.AsFloat := qrItensVALOR_TOTAL.AsFloat - Valor;
      qrItens.Post;
    end;

    lb_item.Caption := IntToStr(iTotal_Itens);
    lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
  end;
end;

procedure TfrmVenda.FormResize(Sender: TObject);
begin
  v_fundo.Left := imgFundo.Left;
  v_fundo.Top := imgFundo.Top;
  v_fundo.Height := imgFundo.Height;
  v_fundo.Width := imgFundo.Width;
  v_borda_superior.Height := Round(p_borda_superior * v_fundo.Height);
  v_borda_inferior.Height := Round(p_borda_inferior * v_fundo.Height);
  v_borda_esquerda.Width := Round(p_borda_esquerda * v_fundo.Width);
  v_borda_direita.Width := Round(p_borda_direita * v_fundo.Width);

  v_espacamento_superior.Height := Round(p_espacamento_superior * v_fundo.Height);
  v_espacamento_central.Width := Round(p_espacamento_central * v_fundo.Width);
  v_espacamento_inferior.Height := Round(p_espacamento_inferior * v_fundo.Height);

  v_barra_superior.Height := Round(p_barra_superior * v_fundo.Height);
  v_barra_inferior.Height := Round(p_barra_inferior * v_fundo.Height);

  v_quadro_esquerdo.Width := Round(p_quadro_esquerdo * v_fundo.Width);

  if AtivaTouch then begin
    pnQuantidade.Width := Round(p_edit_quantidade * (pnBotoes.Width-AdvGlowButton4.Width-AdvGlowButton3.Width-AdvGlowButton2.Width));
  end else begin
    pnQuantidade.Width := Round(p_edit_quantidade * v_tipo_normal_campos1.Width);
    pnTotal.Width := Round(p_edit_valor_unitario * v_tipo_normal_campos2.Width);
  end;

  v_tipo_normal_campos2.Height := Round((v_tipo_normal_campos.Height-Bevel3.Height)/2);
  ///////////////
  wspnlAlerta.Visible := False;
  wspnlAlerta.Height := Self.ClientHeight;
  wspnlAlerta.Width := Self.ClientWidth;
  Image1.Top := (Self.ClientHeight - (Image1.Height + 300)) div 2;
  Image1.Left := (Self.ClientWidth - Image1.Width) div 2;
  cpnlMensagem.Top := Image1.Top + 208;
  cpnlMensagem.Left := (Self.ClientWidth - cpnlMensagem.Width) div 2;
  lblInstrucao.Top := cpnlMensagem.Top + 116;
  lblInstrucao.Left := (Self.ClientWidth - lblInstrucao.Width) div 2;
  wspnlAlerta.Visible := True;

end;

procedure TfrmVenda.finalizarCupomNFCe;
begin
  frmTEF.finalizarCupomNFCe;
end;

procedure TfrmVenda.AlertaFecharClick(Sender: TObject);
begin
  wspnlAlerta.Left := 3000;
  FLastKeyPress := 0;
  ed_barra.Text := '';
  PopupMenu := pop_principal;
  ed_barra.SetFocus;
end;

end.


