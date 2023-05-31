unit principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, MidasLib,
  JvBaseDlg, JvSelectDirectory, ACBrValidador, ACBrBase, ACBrDFe, ACBrNFe,
  System.Actions, Vcl.ActnList, Data.DB, MemDS, DBAccess, Uni, TFlatHintUnit,
  LockApplication, ResizeKit, AdvOfficeStatusBar, AdvOfficeStatusBarStylers,
  System.ImageList, Vcl.ImgList, AdvSmoothButton, Vcl.Buttons, cyBaseSpeedButton,
  cySpeedButton, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel,
  dxGDIPlusClasses, Vcl.ExtCtrls, dxTileControl, cxClasses, dxCustomTileControl,
  IniFiles, ACBrBAL, ACBrDevice, ACBrDeviceSerial, LockApplicationFunc, UniProvider,
  MySQLUniProvider, RzLabel;

type
  TImpressora = (SemImpressora, NaoFiscal, Fiscal);

  TGaveta = (SemGaveta, GavImpNaoFiscal, GavImpFiscal);

  TTipoCupom = (tcNFCE, tcSAT, tcSimples, tcPafECF);

  TStatusTeclado = (stON, stOFF);

  TTipoTeclado = (ttNumerico, ttTefone, ttAlphaNumerico);

type
  TfrmPrincipal = class(TForm)
    ImageList1: TImageList;
    AdvOfficeStatusBarOfficeStyler1: TAdvOfficeStatusBarOfficeStyler;
    ResizeKit1: TResizeKit;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    AdvMetroButton1: TAdvMetroButton;
    AdvMetroButton2: TAdvMetroButton;
    imgFundo: TImage;
    AdvSmoothExpanderPanel3: TAdvSmoothExpanderPanel;
    Label1: TLabel;
    AdvSmoothExpanderPanel4: TAdvSmoothExpanderPanel;
    FlatHint1: TFlatHint;
    IBCDataSource1: TUniDataSource;
    qrForma: TUniQuery;
    query: TUniQuery;
    query2: TUniQuery;
    query3: TUniQuery;
    ActionList1: TActionList;
    vender: TAction;
    configuracoes: TAction;
    menu_fiscal: TAction;
    balanca: TAction;
    ecf: TAction;
    sobre: TAction;
    sair: TAction;
    usuario: TAction;
    mudarsenha: TAction;
    logoff: TAction;
    ACBrNFe1: TACBrNFe;
    AdvSmoothButton4: TAdvSmoothButton;
    AdvSmoothButton6: TAdvSmoothButton;
    AdvSmoothButton7: TAdvSmoothButton;
    GerenciarNFCe: TAction;
    DIretorio: TJvSelectDirectory;
    store_protect: TLockApplication;
    checkcpf1: TACBrValidador;
    CheckCNPJ: TACBrValidador;
    pnRegistro: TPanel;
    Panel2: TPanel;
    Panel1: TPanel;
    Panel4: TPanel;
    lbEmpresa: TLabel;
    Panel5: TPanel;
    Panel6: TPanel;
    Label3: TLabel;
    Image1: TImage;
    Bevel1: TBevel;
    Panel7: TPanel;
    Panel8: TPanel;
    lbContrato: TLabel;
    Panel9: TPanel;
    Panel10: TPanel;
    Image2: TImage;
    Label4: TLabel;
    Bevel2: TBevel;
    Panel11: TPanel;
    Panel12: TPanel;
    lbVigencia: TLabel;
    Panel13: TPanel;
    Panel14: TPanel;
    Image3: TImage;
    Bevel3: TBevel;
    Label5: TLabel;
    Panel15: TPanel;
    Panel16: TPanel;
    lbProrragado: TLabel;
    Panel17: TPanel;
    Panel18: TPanel;
    Image4: TImage;
    Bevel4: TBevel;
    Label6: TLabel;
    Panel19: TPanel;
    Panel20: TPanel;
    lbProximo: TLabel;
    Panel21: TPanel;
    Panel22: TPanel;
    Image5: TImage;
    Bevel5: TBevel;
    Label2: TLabel;
    Panel23: TPanel;
    Panel24: TPanel;
    lbEMAberto: TLabel;
    Panel25: TPanel;
    Panel26: TPanel;
    Image6: TImage;
    Bevel6: TBevel;
    Label7: TLabel;
    Panel27: TPanel;
    Panel28: TPanel;
    Image7: TImage;
    Bevel7: TBevel;
    Label8: TLabel;
    Panel29: TPanel;
    lbMensagem: TLabel;
    Panel3: TPanel;
    RzLabel1: TRzLabel;
    lbTitulo: TLabel;
    Label9: TLabel;
    Image25: TImage;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Balana1Click(Sender: TObject);
    procedure Geral1Click(Sender: TObject);
    procedure AdvGlowButton175Click(Sender: TObject);
    procedure CarregaSistemaVenda;
    procedure VendasPDV1Click(Sender: TObject);
    procedure tlPhotosClick(Sender: TdxTileControlItem);
    procedure tlAgentsClick(Sender: TdxTileControlItem);
    procedure dxtlcntrltmTileItem1Click(Sender: TdxTileControlItem);
    procedure tlSystemInformationClick(Sender: TdxTileControlItem);
    procedure dxTileItem1Click(Sender: TdxTileControlItem);
    procedure V1Click(Sender: TObject);
    procedure S1Click(Sender: TObject);
    procedure C1Click(Sender: TObject);
    procedure B1Click(Sender: TObject);
    procedure bt_menu02Click(Sender: TObject);
    procedure bt_menu07Click(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure venderExecute(Sender: TObject);
    procedure configuracoesExecute(Sender: TObject);
    procedure balancaExecute(Sender: TObject);
    procedure sairExecute(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure AdvMetroButton2Click(Sender: TObject);
  private
    iGaveta: Integer;
    { Private declarations }
    function LerFormasPagto: Boolean;
  public
    { Public declarations }
    TipoImpressora: TImpressora;
    GavetaImpressora: TGaveta;
    procedure GravaIni(sArquivo, Secao, Linha, Valor: string);
    function LerIni(sArquivo, Secao, Linha: string; Def: string): string;
    function ApenasNumerosStr(pStr: string): string;
    function IsInteger(s: string): integer;
    function codifica(Tabela: string; qtde_digitos: integer): string;
    function StrZeros(Texto: string; Qtde: Integer): string;
    function LeInistr(arquivo, sessao, chave, padrao: string): string;
  end;

const
  StRecDinheiro = 'DINHEIRO';
  StRecConvenio = 'CONVENIO';
  StRecCrediario = 'CARTEIRA';
  StRecCartaoCR = 'CARTAO CREDITO';
  StRecCartaoDE = 'CARTAO DEBITO';
  StRecCheque = 'CHEQUE';
  StRecEstorno = 'ESTORNO';
  StRecCredito = 'CUPOM CREDITO';



type
 // tipo de bandeira tef
  Tbandeira_tef = (Tef_Redecard, Tef_Visanet, Tef_Amex, Tef_TecBan, Tef_Hipercard, Tef_Tribanco);

var
  frmPrincipal: TfrmPrincipal;
  TipoDeCupom: TTipoCupom;
  ExibeTecladoVirtual, AtivaTouch: Boolean;
	vEcfRecDinheiro, vEcfRecConvenio,vEcfRecCrediario,vEcfRecCartaoCR,vEcfRecCartaoDE,vEcfRecCheque:string;
	ExibeEstoqueTelaInicial:Boolean;
	SolicitaLiberacaoCancelarGerencial:Boolean;
	sModelo: string;
	bHabGuilhotina: Boolean;
	bEdicaoRelatorio:Boolean;
	bNao_utiliza_extra :Boolean;
	sProd_nome, sProd_barra, sProd_unidade, sProd_CST: string;
	iImpressora, iProd_codigo: integer;
	rProd_aliquota, rProd_Estoque, rProd_preco, rProd_qtde, rProd_total, rProd_desconto_maximo, rProd_acrescimo_maximo: real; // armazenar o desconto maximo permitido por produto
	rProd_desconto, rProd_acrescimo: real; // armazenar o desconto e acrescimo no item
	sProd_Tamanho: string;
	sProd_Cor, sProd_Compl_Pizza: string;
	sMsg: string;
	ECFCaminhoFiscal:string;
	lForma_pgto_Cheque_Avista, lForma_pgto_Cheque_Aprazo, lForma_pgto_Crediario, lForma_pgto_Cartao_Debito, lForma_pgto_Cartao_Credito, lForma_pgto_dinheiro, lForma_pgto_Convenio: string;
	sCaixa: string; // numero do caixa
	dData_Sistema: Tdatetime;
	dData_Movimento: Tdatetime;
	hHora_Cupom: TDateTime;
	sNumero_Cupom: string; // numero do coo atual
	sNumero_Venda: string; // Numero do cupom fiscal atual
	sMD5:string;
	sNumero_contador_cupom: string; // numero de contador do cupom fiscal atual
	sCOO, sGNF, sGRG, sCDC, sCER: string; // outros contadores do ecf
	(* Display do Teclado *)
	(* Classificar o tipo de venda do sistema*)
	iTipo_Venda: integer; // 1 Super 2 Prevenda 3 Posto
	iTamanho_codigo_balanca: integer; // tamanho do codigo da balaca 4 ou 5
	PrazoDevolucao: Cardinal;
	bMuda_qtde: boolean; // configuracao para permitir a edicao da qtde
	bMuda_total, bMudouProvisorio, bMuda_unitario, bMensagem_Cupom,
	Men_Cup_Simples, Men_Cup_NFCE, Men_Cup_SAT, Men_Cup_Cozinha: boolean; // configuracao para permitir a edicao do preco unitario
	bPesagemAutomatica: Boolean;
	bSenha_Cancel_Item: boolean; // exigir senha no cancelamento do item
	bSenha_Cancel_Cupom: boolean; // exigir senha no cancelamento do cupom
	iBal_time: integer; // configuracao de time out da balanca
	sBal_Resposta: string; // resposta da balanca
	rBal_peso: real; // peso lido na balanca
	sPasta_foto_produto: string; // pasta onde estao as fotos dos produtos
	bHabLogoMarca: Boolean; // Logo Marca PDV
	bBusca_foto_produto: boolean; // buscar foto do produto
	iPesquisa_produto: integer; // listar no grid do form produto_consulta os itens filtrados 0 - pesquisa feita no form de pesquisa, 1 - pesquisa feita na venda, 2 - pesquisa feita no form de consulta de preco
	Gera_Numero_NFCe_Local, PDV_OnLine:Boolean;
	SDAV_ATUALIZADO: string;
	sCli_Nome, sCli_Endereco, sCli_CPF, sCli_Cidade, sCli_Placa, sCli_Km, sCli_uf, sCli_codigo, scli_cep, scli_vendedor, sCLi_PegoPor: string;
	bVenda_Cartao, bVenda_Crediario, bVenda_Cheque, bVenda_dinheiro, bVenda_Convenio: boolean; // verificar se foi vendido nas formas de pagamento que exigemtratamento especial
	bConsulta_crediario: boolean; // verificar se a consulta na tela de clientes eh para o crediario;
	iCrediario_prestacao: integer; // armazenar a qtde de prestacoes no crediario
	sCrediario_Nome: string; // nome da forma de pagamento de creadiario
	bCadastra_Cheque, bCadastra_Convenio, bCadastra_Crediario: boolean; // verificar cadastramentos no final da venda
	sTotalizador_Crediario, sNome_Totalizador_Crediario: string; // nome do totalizador gravado no ecf para emissao de comprovante nao fiscal no final da venda em crediario.
	sTotalizador_Recebimento, sNome_Totalizador_Recebimento: string; // nome do totalizador gravado no ecf para emissao de comprovante nao fiscal no final da venda em crediario.
	sIndice_Sangria, sIndice_Suprimento: string; // numero do indice gravado na impressora da sangria e do suprimento
	iComprovente_Crediario: integer; // Imprimir comprovante nao fiscal do crediairo
	(* Outras*)
	bContinua: boolean; // Continuar ou nao um procedimento/funcao
	bSair_campo: boolean; // forçar a edicao do campo
	bSupervisor_autenticado: boolean; // verifiar se o supervisor foi logado
	(* Usuario e Senha *)
	sFuncao_Senha: string; // determinar a funcao para a tela de senha
	busuario_autenticado: boolean; // verificar se o usuario foi autenticado
	icodigo_Usuario: integer; // codigo do usuario logado
	sNome_Operador: string; // nome do operador
	iNumCaixa: integer; // Identificação do caixa
	iCFOP: string; // cfop de venda paf
	bFechamentoCego:Boolean;
	(* Identificacao do consumidor no inicio da venda --> CAt52 *)
	bIdentifica_consumidor: boolean;
	sConsumidor_CPF, sConsumidor_Nome, sConsumidor_Endereco: string;
	(* variaveis de comunicao com o servidor *)
	sServidor_host, sServidor_Base: string;
	(* Pre venda *)
	bLanca_pre_venda: boolean; // flag para fazer o fechamento da pre-venda
	sPre_Venda_Numero: string; // armazenar o numero da prevenda para ser impresso no final do cupom
	iPre_venda_codigo, iPre_Venda_Vendedor, iPre_venda_cliente: integer; // codigo do vendedor e do cliente
	rPre_venda_desconto, rPre_venda_acrescimo: real; // acrescimo e desconto no fechamento
	(* OS *)
	bLanca_OS: boolean; // flag para fazer o fechamento da OS
	sOS_Numero: string; // armazenar o numero da prevenda para ser impresso no final do cupom
	iOS_codigo, iOS_Vendedor, iOS_cliente: integer; // codigo do vendedor e do cliente
	rOS_desconto, rOS_acrescimo: real; // acrescimo e desconto no fechamento
	(* Abasteciento *)
	bLanca_Abastecimento: boolean; // flag para fazer o fechamento do abastecimento
	(* Mesas *)
	bLanca_Mesa: boolean; // flag para fazer o fechamento da Mesa
	bLanca_Comanda: boolean; // flag para fazer o fechamento da Comanda/Mesa
	iID_Delivery: integer; // codigo da Mesa
	bLanca_Delivery: boolean; // flag para fazer o fechamento do Delivery
	MesasComandasReceber:string;
	(* dav *)
	sDav_numero: string; // armazenar o numero do dav caso for ser impresso no final do cupom
	(* atualizacao de dados *)
	sCarga_data, scarga_hora: string; // data e hora da ultima atualizacao de dados
	Arquivo_ini: TIniFile; // nome do arquivo .ini
	(*flag de retorno da situacao do caixa e da venda*)
	bCaixa_aberto: boolean;
	bativa_venda: boolean;
	SenhaLocal:Boolean;
	// posto - rodape do cupom fiscal
	sPosto_rodape: string;
	iCodigo_abastecimento, ilinha_abastecimento: integer;
	// posto - porta de comunicacao com Concentrador
	iPosto_porta: integer;
	// verificar se o cupom estah aberto, caso o comando direto da ecf falhe
	bcupom_aberto: boolean;
	bVerifica_Caixa:Boolean;
	// verificar se pede senha ao dar desconto
	isenha_desconto: integer;
	rDesconto_maximo, racrescimo_maximo: real;
	bIdentificarVendedor: Integer; // Verificar se vai pedir para identificar o vendedor no fechamento da venda ou não
	sCodigoSegurancaFicha, sImpressora_Venda_Simples, sImpressora_Cozinha,
	sImpressora_Ficha, sImpressora_carne: string;
	bVerificaProdutoVenda, bExibeEstoqueConsulta, bExibePrecoConsulta, bSenhaVendaSemEstoque, bAtivaImpFicha,bImpressaoCozinha, bExibeCumponNFTela, bImprimeCarne: Boolean;
	//  Identifica qual o campo será mostrado na coluna onde ´tem o código
	// de barras na consulta do produto
	// 0 - Código de Barras, 1 - Referência, 2 - Referência+Tamanho+Cor
	iColunaCOnsultaProdutos: Integer;
	iQtdLinhasFinal: Integer; // quantidade de linhas no final da nota 40c
	GavAtiva:Boolean;
	GavModelo:Integer;
	GavPorta, GavTipoCorte:string;
	GavInvertido, GavCortaPapel:Boolean;
	LicencaOnLine:Boolean;
	IDIntalacao:Int64;
	onLineParceiro, onLineServidor, onLinePorta, onLineBanco, onLineUser, onLinePass:string;
	IDParceiro:Integer;

implementation

uses
  modulo, funcoes, venda, senha, msg_Operador, produto_consulta,
  Orcamento_Abrir, Math, caixa_abertura, config_balanca, Config,
  UtiSplash, UFuncoes, contasreceber, frmNFCEs, senha_supervisor,
  FechamentoCego, caixa_fechamento, OnLine;


{$R *.dfm}

function tfrmprincipal.codifica(tabela: string; qtde_digitos: integer): string;
begin
  frmmodulo.qrmestre.close;
  frmmodulo.qrmestre.sql.Clear;
  frmmodulo.qrmestre.sql.add('select * from c000000 where codigo = ''' + tabela + '''');
  frmmodulo.qrMestre.Open;
  if frmmodulo.qrmestre.RecNo = 1 then begin
    if frmmodulo.qrmestre.FieldByName('sequencia').asinteger < 1 then begin
      result := '0000000001';
      frmmodulo.qrMestre.Edit;
      frmmodulo.qrMestre.fieldbyname('sequencia').asinteger := 2;
      frmmodulo.qrMestre.Post;

    end
    else begin
      result := zerar(inttostr(frmmodulo.qrMestre.fieldbyname('sequencia').asinteger), qtde_digitos);
      frmmodulo.qrMestre.Edit;
      frmmodulo.qrMestre.fieldbyname('sequencia').asinteger := frmmodulo.qrMestre.fieldbyname('sequencia').asinteger + 1;
      frmmodulo.qrMestre.Post;
    end;
  end
  else begin
    Showmessage('Não foi possível concluir com a operação!' + #13 + 'Erro: Código de sequência não encontrado na tabela de códigos.');
  end;
end;

procedure TfrmPrincipal.configuracoesExecute(Sender: TObject);
begin
  frmsenha_supervisor := tfrmsenha_supervisor.create(self);
  frmsenha_supervisor.Configuracoes := True;
  frmsenha_supervisor.showmodal;

  if not bSupervisor_autenticado then
    Exit;
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;
end;

procedure TfrmPrincipal.bt_menu07Click(Sender: TObject);
begin

end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  application.Terminate;
end;

procedure TfrmPrincipal.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin

	Top := 0;
	Left := 0;

	sConfiguracoes := ExtractFilePath(Application.ExeName) + 'CFG\cfg.ini';

	try
		AcertaPadraoData;
	except
	end;

	frmModulo := TfrmModulo.create(self);

	frmModulo.conexao.Connected := false;
	frmModulo.conexao.Database := LerINi(sConfiguracoes, 'CONFIGURACOES', 'BDLOCAL', 'C:\SOS\pdv\bd\nfce_datpdv.fdb');
	frmModulo.conexao.Connected := true;
	try
		frmModulo.Conexao_Servidor.Connected := false;
		frmModulo.Conexao_Servidor.Server := LerINi(sConfiguracoes, 'CONFIGURACOES', 'SERVIDOR', 'localhost');
		frmModulo.Conexao_Servidor.Database := LerINi(sConfiguracoes, 'CONFIGURACOES', 'BDSERVIDOR', 'C:\SOS\server\bd\base.fdb');
		frmModulo.Conexao_Servidor.Connected := true;
		PDV_OnLine := True;
	except
		PDV_OnLine := false;
	end;

	FrmModulo.Conn := frmModulo.Conexao;
	if not FrmModulo.CampoExiste('CONFIG','PRAZO_DEVOLUCAO') then
		 FrmModulo.InsereCampo('CONFIG','PRAZO_DEVOLUCAO INTEGER DEFAULT 30');
	FrmModulo.Conn := frmModulo.Conexao_Servidor;
	if not FrmModulo.CampoExiste('CONFIG','PRAZO_DEVOLUCAO') then
		 FrmModulo.InsereCampo('CONFIG','PRAZO_DEVOLUCAO INTEGER DEFAULT 30');

  // abrir banco de dados de configuracoes local
	query.close;
	query.sql.text:='select * from config';
	query.open;

	isenha_desconto := query.fieldbyname('senha_desconto').AsInteger;
	iComprovente_Crediario := query.fieldbyname('COMPROVANTE_CREDIARIO').asinteger;
	sTotalizador_Crediario := query.fieldbyname('TOTALIZADOR_NF_CREDIARIO').asstring;
	sNome_Totalizador_Crediario := query.fieldbyname('Nome_TOTALIZADOR_NF_CREDIARIO').asstring;
	sTotalizador_Recebimento := query.fieldbyname('TOTALIZADOR_NF_RECEBIMENTO').asstring;
	sNome_Totalizador_Recebimento := query.fieldbyname('Nome_TOTALIZADOR_NF_RECEBIMENTO').asstring;
	sIndice_Sangria := query.fieldbyname('totalizador_sangria').asstring;
  sIndice_Suprimento := query.fieldbyname('totalizador_suprimento').asstring;
	bNao_utiliza_extra := query.fieldbyname('nao_utiliza_extra').asstring = 'S';

	// alimentar as variaveis com as formas de pagamentos pre-definidas que exigem tratamentos especiais no fechamento da venda
	if query.RecordCount > 0 then
	begin
		lForma_pgto_Cheque_Avista := query.fieldbyname('forma_cheque').asstring;
		lForma_pgto_Cheque_Aprazo := query.fieldbyname('forma_cheque_aprazo').asstring;
		lForma_pgto_Cartao_Debito := query.fieldbyname('forma_cartao').asstring;
		lForma_pgto_Cartao_Credito := query.fieldbyname('forma_cartao_credito').asstring;
		lForma_pgto_Crediario := query.fieldbyname('forma_crediario').asstring;
		lForma_pgto_Convenio := query.fieldbyname('forma_convenio').asstring;
		lForma_pgto_Dinheiro := query.fieldbyname('forma_dinheiro').asstring;
	end;

	// alimentando as variaveis globais do sistema
	PrazoDevolucao := query.fieldbyname('PRAZO_DEVOLUCAO').asinteger;
	if query.fieldbyname('muda_qtde').asinteger = 1 then
		bMuda_qtde := true
	else
		bMuda_qtde := false;

	if query.fieldbyname('muda_unitario').asinteger = 1 then
		bMuda_unitario := true
	else
		bMuda_unitario := false;

	if query.fieldbyname('muda_total').asinteger = 1 then
		bMuda_total := true
	else
		bmuda_total := false;

	bMensagem_Cupom := query.fieldbyname('imprime_mensagem_cupom').AsString = 'S';
	Men_Cup_Simples := query.fieldbyname('men_cup_simples').AsString = 'S';
	Men_Cup_NFCE := query.fieldbyname('men_cup_nfce').AsString = 'S';
	Men_Cup_SAT := query.fieldbyname('men_cup_sat').AsString = 'S';
	Men_Cup_Cozinha := query.fieldbyname('men_cup_cozinha').AsString = 'S';
	bPesagemAutomatica := query.fieldbyname('pesagem_automatica').AsString = 'S';
	bSenha_Cancel_Item := query.fieldbyname('solicita_senha_cancelamento').AsString = 'S';
	bSenha_Cancel_Cupom := query.fieldbyname('solicita_senha_cancelamento').AsString = 'S';
	rDesconto_maximo := query.fieldbyname('desconto_maximo_venda').AsFloat;
	bCadastra_cheque := true;
	bCadastra_Crediario := true;
	// alimentacao das variaveis de identificaco de consumidor para atender
	// CAT52 - Nf paulista
	if query.fieldbyname('nf_paulista').asinteger = 1 then
    bIdentifica_consumidor := true
  else
    bIdentifica_consumidor := false;
end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.FormShow(Sender: TObject);
var
  RetRegistro:TRegistro;
  arquivo_ini: TIniFile;
begin
  imgFundo.Picture.LoadFromFile('C:\SOS\PDV\img\fundo_inicial.jpg');
  RetRegistro := RegistroSistema('SIS_001','SIS_001_PDV',False,'PDV');
  if (RetRegistro.Situacao = tsBloqueado) or (RetRegistro.Situacao = tsInvalido) then begin
    if (RetRegistro.Situacao = tsBloqueado) then
      Application.MessageBox(PWideChar(RetRegistro.MensagemBloqueio),'Atenção!',MB_ICONWARNING);
    Close;
    Application.Terminate;
    Exit;
  end;
  if RetRegistro.TipoRegsitro = trOnLine then begin
    lbEmpresa.Caption := RetRegistro.Empresa;
    lbContrato.Caption := RetRegistro.Contrato;
    lbVigencia.Caption := RetRegistro.Vigencia;
    lbProrragado.Caption := RetRegistro.Prorrogacoes;
    lbProximo.Caption := RetRegistro.ProximoVencimento;
    lbEMAberto.Caption := RetRegistro.Atrazo;
    lbMensagem.Caption := RetRegistro.Mensagem;
    pnRegistro.Visible := True;
  end;

  frmModulo.qrConfigPDV.Close;
  frmModulo.qrConfigPDV.Open;
  ATB_ATIVA.ATB := frmModulo.qrConfigPDVEMPRESA_EXTRUTURAL.AsString;
  ATB_ATIVA.TIPO := frmModulo.qrConfigPDVEMPRESA_TIPO.AsString;
  ATB_ATIVA.DESCRICAO := frmModulo.qrConfigPDVEMPRESA_DESCRICAO.AsString;
  if ATB_ATIVA.ATB = '' then begin
    Application.MessageBox(PWideChar('Não foi encontrado a configuração de empresa ativa Rode o PDV_Server para atualizar a Base do PDV.'), 'Atenção!', MB_ICONERROR);
    Application.Terminate;
    Exit;
  end;
end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.sairExecute(Sender: TObject);
begin
  close;
end;
// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.Balana1Click(Sender: TObject);
begin
  frmconfig_balanca := tfrmconfig_balanca.create(self);
  frmconfig_balanca.showmodal;
end;

procedure TfrmPrincipal.balancaExecute(Sender: TObject);
begin
  frmsenha_supervisor := tfrmsenha_supervisor.create(self);
  frmsenha_supervisor.Configuracoes := True;
  frmsenha_supervisor.showmodal;

  if not bSupervisor_autenticado then
    Exit;
  frmconfig_balanca := tfrmconfig_balanca.create(self);
  frmconfig_balanca.showmodal;
end;

procedure TfrmPrincipal.bt_menu02Click(Sender: TObject);
begin

end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.Geral1Click(Sender: TObject);
begin
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;
  frmconfig.free;
end;

// -------------------------------------------------------------------------- //

procedure TfrmPrincipal.venderExecute(Sender: TObject);
begin
  try
    AdvSmoothButton4.Enabled := False;
    CarregaSistemaVenda;
  finally
    AdvSmoothButton4.Enabled := True;
  end;
end;

procedure TfrmPrincipal.AdvGlowButton175Click(Sender: TObject);
begin
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;
end;

// -------------------------------------------------------------------------- //
procedure TfrmPrincipal.AdvMetroButton1Click(Sender: TObject);
begin
  Application.Minimize;
end;

procedure TfrmPrincipal.AdvMetroButton2Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmPrincipal.CarregaSistemaVenda;
var
  sForma: string;
  i, x: integer;
  dDataHora: TSystemTime;
  Ano, Mes, Dia, H, M, S, Mil: word;
  sLinha: string;
  txt: textFile;
	TipoCupom:string;

begin
  frmmodulo.qrconfig.Close;
	frmmodulo.qrconfig.Open;

	iImpressora := 1;
	case iImpressora of
		0: TipoImpressora := SemImpressora; // Sem Impressora
		1: TipoImpressora := NaoFiscal; // Nao Fiscal
		2: TipoImpressora := Fiscal; // Fiscal
	end;

	iGaveta := StrToInt(frmPrincipal.LerINi(sConfiguracoes, 'GAVETA', 'LOCAL', '0'));
	case iGaveta of
		0: GavetaImpressora := SemGaveta; // Sem Gaveta
		1: GavetaImpressora := GavImpNaoFiscal; // Gaveta Impressora Nao Fiscal
		2: GavetaImpressora := GavImpFiscal; // Gaveta Impressora Fiscal
	end;

	try
		if strtoint(LerINi(sConfiguracoes, 'BALANCA', 'Modelo', '0')) <> 0 then
		with frmmodulo do
		begin
			balanca.Modelo           := TACBrBALModelo(StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Modelo', '0')));
			balanca.Device.HandShake := TACBrHandShake(StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Handshaking', '0')));
			balanca.Device.Parity    := TACBrSerialParity(StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Parity', ' 0')));
			balanca.Device.Stop      := TACBrSerialStop(StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Stopbits', '0')));
			balanca.Device.Data      := StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Databits', '0'));
			balanca.Device.Baud      := StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Baudrate', '0'));
			balanca.Device.Porta     := LerIni(sConfiguracoes, 'BALANCA', 'Porta', '');
			balanca.Device.TimeOut   := StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Timeout', '1000'));
			iBal_time := StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'Timeout', '1000'));
		end;
	except
		Application.MessageBox('As Configurações da balança são inválidas, verifique!', 'Atenção!', MB_ICONINFORMATION);
	end;

	AtivaTouch := LerINi(sConfiguracoes, 'Geral', 'UTILIZA TOUCH', '0') = '1';
	bVerifica_Caixa := LerINi(sConfiguracoes, 'Geral', 'VERIFICA_CAIXA', '1') = '1';
	ExibeTecladoVirtual := LerINi(sConfiguracoes, 'Geral', 'UTILIZA TECLADO', '0') = '1';
	bHabGuilhotina := StrToBool(LerIni(sConfiguracoes, 'GUILHOTINA', 'HABILITA', BoolToStr(False)));
	sModelo := LerIni(sConfiguracoes, 'GUILHOTINA', 'MODELO', 'PADRÃO');
	iTamanho_codigo_balanca := StrToInt(LerIni(sConfiguracoes, 'BALANCA', 'NDIGITOS', '5'));
	SenhaLocal := StrToBool(LerIni(sConfiguracoes, 'Geral', 'SENHA_LOCAL', BoolToStr(False)));
	bBusca_foto_produto := StrToBool(LerINI(sConfiguracoes, 'PDV', 'HAB_IMG', BoolToStr(false)));
	sPasta_foto_produto := LerINI(sConfiguracoes, 'PDV', 'CAMINHO_IMG', ExtractFilePath(Application.ExeName) + '\img\produtos\');
	bHabLogoMarca := StrToBool(LerINI(sConfiguracoes, 'PDV', 'HAB_LOGO', BoolToStr(false)));
	Gera_Numero_NFCe_Local :=  StrToBool(LerINI(sConfiguracoes, 'Geral', 'CONTROLA_NFCE_LOCAL', BoolToStr(false)));
	EditLogoMarca := LerINI(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', '');
	GavAtiva := LerINi(sConfiguracoes, 'GAVETA', 'Ativo', '0') = '1';
	GavModelo := StrToInt(LerINi(sConfiguracoes, 'GAVETA', 'Modelo', '0'));
	GavPorta := LerINi(sConfiguracoes, 'GAVETA', 'Porta', '');
	GavCortaPapel := StrToBool(LerINI(sConfiguracoes, 'GAVETA', 'CortaPapel', BoolToStr(True)));
	GavInvertido := StrToBool(LerINI(sConfiguracoes, 'GAVETA', 'Invertido', BoolToStr(false)));
	GavTipoCorte :=  LerINi(sConfiguracoes, 'GAVETA', 'TipoCorte', 'Total');

  // abrir banco de dados de configuracoes local
	query.close;
	query.sql.text:='select * from config';
  query.open;

	isenha_desconto := query.fieldbyname('senha_desconto').AsInteger;
	iComprovente_Crediario := query.fieldbyname('COMPROVANTE_CREDIARIO').asinteger;
	sTotalizador_Crediario := query.fieldbyname('TOTALIZADOR_NF_CREDIARIO').asstring;
	sNome_Totalizador_Crediario := query.fieldbyname('Nome_TOTALIZADOR_NF_CREDIARIO').asstring;
	sTotalizador_Recebimento := query.fieldbyname('TOTALIZADOR_NF_RECEBIMENTO').asstring;
	sNome_Totalizador_Recebimento := query.fieldbyname('Nome_TOTALIZADOR_NF_RECEBIMENTO').asstring;
	sIndice_Sangria := query.fieldbyname('totalizador_sangria').asstring;
	sIndice_Suprimento := query.fieldbyname('totalizador_suprimento').asstring;
	bNao_utiliza_extra := query.fieldbyname('nao_utiliza_extra').asstring = 'S';

	// alimentar as variaveis com as formas de pagamentos pre-definidas que exigem tratamentos
  // especiais no fechamento da venda
	if query.RecordCount > 0 then
	begin
    lForma_pgto_Cheque_Avista := query.fieldbyname('forma_cheque').asstring;
    lForma_pgto_Cheque_Aprazo := query.fieldbyname('forma_cheque_aprazo').asstring;
    lForma_pgto_Cartao_Debito := query.fieldbyname('forma_cartao').asstring;
    lForma_pgto_Cartao_Credito := query.fieldbyname('forma_cartao_credito').asstring;
    lForma_pgto_Crediario := query.fieldbyname('forma_crediario').asstring;
    lForma_pgto_Convenio := query.fieldbyname('forma_convenio').asstring;
    lForma_pgto_Dinheiro := query.fieldbyname('forma_dinheiro').asstring;
  end;

  LerFormasPagto;

	// alimentando as variaveis globais do sistema
	PrazoDevolucao := query.fieldbyname('PRAZO_DEVOLUCAO').asinteger;
	if query.fieldbyname('muda_qtde').asinteger = 1 then
		bMuda_qtde := true
	else
		bMuda_qtde := false;
	if query.fieldbyname('muda_unitario').asinteger = 1 then
		bMuda_unitario := true
	else
		bMuda_unitario := false;
	if query.fieldbyname('muda_total').asinteger = 1 then
		bMuda_total := true
	else
		bmuda_total := false;
	bMensagem_Cupom := query.fieldbyname('imprime_mensagem_cupom').AsString = 'S';
	Men_Cup_Simples := query.fieldbyname('men_cup_simples').AsString = 'S';
	Men_Cup_NFCE := query.fieldbyname('men_cup_nfce').AsString = 'S';
	Men_Cup_SAT := query.fieldbyname('men_cup_sat').AsString = 'S';
	Men_Cup_Cozinha := query.fieldbyname('men_cup_cozinha').AsString = 'S';
	bPesagemAutomatica := query.fieldbyname('pesagem_automatica').AsString = 'S';
	bSenha_Cancel_Item := query.fieldbyname('solicita_senha_cancelamento').AsString = 'S';
	bSenha_Cancel_Cupom := query.fieldbyname('solicita_senha_cancelamento').AsString = 'S';
	rDesconto_maximo := query.fieldbyname('desconto_maximo_venda').AsFloat;

	bCadastra_cheque := true;
	bCadastra_Crediario := true;

	// alimentacao das variaveis de identificaco de consumidor para atender
	// CAT52 - Nf paulista
	if query.fieldbyname('nf_paulista').asinteger = 1 then
		bIdentifica_consumidor := true
	else
		bIdentifica_consumidor := false;

	TipoCupom := UpperCase(frmPrincipal.LerINi(sConfiguracoes, 'Geral', 'TIPO CUPOM', 'NFCE'));
	if TipoCupom = 'SAT' then
		TipoDeCupom := tcSAT
	else
	if TipoCupom = 'NFC-E' then
		TipoDeCupom := tcNFCE
	else
	if TipoCupom = 'PAF-ECF' then
		TipoDeCupom := tcPafECF
	else
		TipoDeCupom := tcSimples;
	dData_Sistema := Date;
	dData_Movimento := dData_Sistema;

	bAtiva_Venda := true;
	iQtdLinhasFinal := IsInteger(LerIni(sConfiguracoes, 'IMPRESSORA', 'qtdlfin', '8'));

	sFuncao_Senha := '001';
	frmSenha := tfrmSenha.create(self); //qiqi

  if Length(LerIni(sConfiguracoes, 'LOGIN', 'USUARIO', '')) = 0 then
    frmSenha.showmodal
  else
    frmSenha.LogarAutomaticamente;

  if busuario_Autenticado then
    frmSenha.Free
  else
    Exit;

  iNumCaixa := icodigo_Usuario;
  iCFOP := (LerIni(sConfiguracoes, 'Emitente', 'CFOP', ''));
  if iNumCaixa = 0 then begin
    Application.MessageBox('Número do caixa nao definido, favor verificar', 'Mensagem');
    exit;
  end;

  if PDV_OnLine then
    Caixa_Geral := CaixaGeral;
	SolicitaLiberacaoCancelarGerencial := frmModulo.qrconfigLIBERA_CANC_PDV_GERENCIAL.AsString = 'S';
	try
    query.close;
    query.sql.clear;
    query.sql.add('select * from config');
    query.open;
		if busuario_Autenticado then
		begin
			if bVerifica_Caixa then
			begin
				bcaixa_aberto := True;
				if (query.fieldbyname('caixa_situacao').asstring = 'ABERTO') and(query.fieldbyname('caixa_data_movto').AsDateTime <> Date) then
				begin
					if Application.MessageBox( 'Não foi realizado o Fechamento deste caixa ainda, '+
																		 'impossível acessar a tela de vendas.'+#13+
																		 'Deseja Efetuar o Fechamento do Mesmo?','Atenção!',MB_ICONINFORMATION+MB_YESNO) = mrYes then
					begin
						frmSenha_Supervisor := TfrmSenha_Supervisor.create(Self);
						frmSenha_Supervisor.FechaCaixa := True;
						frmSenha_Supervisor.showmodal;

						if not bSupervisor_autenticado then
							Exit;

						if bFechamentoCego then
						begin
							frmFechamentoCego := TfrmFechamentoCego.create(Self);
							frmFechamentoCego.showmodal;
						end
						else
						begin
							frmCaixa_Fechamento := TfrmCaixa_Fechamento.create(Self);
							frmCaixa_Fechamento.showmodal;
						end;
            query.close;
            query.sql.clear;
            query.sql.add('select * from config');
            query.open;
						bcaixa_aberto := True;
						if (query.fieldbyname('caixa_situacao').asstring = 'ABERTO') and (query.fieldbyname('caixa_data_movto').AsDateTime <> Date) then
						begin
							Application.MessageBox('Não foi identificado o fechamento do caixa atual, impossível acessar a tela de vendas.','Atenção!',MB_ICONINFORMATION);
							Exit;
						end;
					end
					else
						Exit;
				end;
			end;
			if (frmmodulo.qrconfigCAIXA_AUTOMATICO_GERENCIAL.AsString = 'S') and (PDV_OnLine) then
			begin
				frmmodulo.qrcaixa_operador.Close;
				frmmodulo.qrcaixa_operador.Open;
				frmmodulo.qrcaixa_operador.Locate('codigo', Caixa_Geral, [loCaseInsensitive]);
				if frmmodulo.qrcaixa_operadorDATA.AsDateTime <> Date then
					frmModulo.FechaAbreCaixas;
			end
			else
			begin
				// antes de iniciar a tela venda, verificar as configuracoes necessarias
				if (bVerifica_Caixa) and (PDV_OnLine) then
				begin
					with frmModulo do
					begin
						query_servidor.Close;
						query_servidor.SQL.Clear;
						query_servidor.SQL.Add('select * from c000045 where codigo = '+QuotedStr(Caixa_Geral));
						query_servidor.Open;
						if query_servidor.IsEmpty then
						begin
							Application.MessageBox('Não foi identificado o Caixa Geral "99" no Gerencial, impossivel realizar vendas!','Atenção!',MB_ICONINFORMATION);
							Exit;
						end;
						if query_servidor.FieldByName('DATA').AsDateTime <> dData_Sistema then
						begin
							Application.MessageBox('Não foi aberto um novo Caixa Geral "99" no Gerencial, impossivel realizar vendas!','Atenção!',MB_ICONINFORMATION);
							Exit;
						end;
						if query_servidor.FieldByName('SITUACAO').AsInteger <> 1 then
						begin
							Application.MessageBox('O Caixa Geral "99" no Gerencial não esta aberto, impossivel realizar vendas!','Atenção!',MB_ICONINFORMATION);
							Exit;
						end;
					end;
				end;
			end;
      query.close;
			query.sql.text:='select * from config';
			query.open;

			if query.fieldbyname('caixa_situacao').asstring = 'FECHADO' then
			begin
				if Application.MessageBox(pwidechar('C A I X A   F E C H A D O!' + #13 + 'Deseja realizar a sua abertura?'), 'Aviso', mb_yesno + MB_ICONQUESTION) = idYes then
				begin
					bcaixa_aberto := false;
					bAtiva_Venda := true;
				end
				else
				begin
					Application.MessageBox(pwidechar('As funções de venda serão bloqueadas!'), 'Atenção', mb_ok + MB_ICONWARNING);
					bAtiva_Venda := false;
				end;
			end
			else
				bcaixa_aberto := True;

			if not bCaixa_aberto then
			begin
				frmsenha_supervisor := tfrmsenha_supervisor.create(self);
				frmsenha_supervisor.AbrirCaixa := True;
				frmsenha_supervisor.showmodal;
				if not bSupervisor_autenticado then
					Exit;
				frmCaixa_Abertura := TfrmCaixa_Abertura.create(Self);
        frmCaixa_Abertura.showmodal;
        frmModulo.spCupom_Temp_delete.prepare;
        frmModulo.spCupom_Temp_delete.execute;
        if not bCaixa_aberto then begin
          Application.messagebox(PWideChar('As funções de venda serão bloqueadas!'), 'Atenção', MB_OK + MB_ICONWARNING);
          Exit;
        end;
      end;

			ExibeEstoqueTelaInicial := frmModulo.qrconfigPDV_EXIBE_ESTOQUE_TELA_VENDA.AsString = 'S';

			if bAtiva_Venda then
			begin
				frmvenda := Tfrmvenda.CREATE(SELF);
				frmvenda.cpBarra.Panels.Items[3].Text := sCaixa;
				frmvenda.cpBarra.Panels.Items[1].Text := copy(sNome_Operador, 1, 10);
				frmvenda.SHOWMODAL;
			end;

    end
		else
			Application.Terminate;
	finally
		if frmvenda <> nil then
			FreeAndNil(frmVenda);
	end;
end;

procedure TfrmPrincipal.VendasPDV1Click(Sender: TObject);
begin
  CarregaSistemaVenda;
end;

procedure TfrmPrincipal.GravaIni(sArquivo, Secao, Linha, Valor: string);
var
  Ini: TIniFile;
  Arquivo: string;
begin

  Arquivo := sArquivo;
  Ini := TIniFile.Create(Arquivo);
  Ini.WriteString(Secao, Linha, Valor);
  Ini.Free;

end;

function TfrmPrincipal.LerIni(sArquivo, Secao, Linha, Def: string): string;
var
  Ini: TIniFile;
begin

  Ini := TIniFile.Create(sArquivo);
  Result := Ini.ReadString(Secao, Linha, Def);
  Ini.Free;

end;

function TfrmPrincipal.LeInistr(arquivo, sessao, chave, padrao: string): string;
var
	ArqIni: TIniFile;
begin
	ArqIni := TIniFile.create(arquivo);
	try
		result := ArqIni.ReadString(sessao, chave, padrao);
	finally
		ArqIni.Free;
  end;
end;

function TfrmPrincipal.LerFormasPagto: Boolean;
var
  formas: string;
  c: integer;
begin
  try
    Result := False;
    frmModulo.tbForma_pgto.Close;
    frmModulo.tbForma_pgto.Open;
    frmModulo.tbForma_pgto.First;
    while not frmModulo.tbForma_pgto.Eof do
      frmModulo.tbForma_pgto.Delete;

    c := 0;
     //Dinheiro
    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_dinheiro;
    frmModulo.tbForma_pgto.Post;

     //Cheque
    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Cheque_Avista;
    frmModulo.tbForma_pgto.Post;

    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Cheque_Aprazo;
    frmModulo.tbForma_pgto.Post;

    //Cartao
    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Cartao_Debito;
    frmModulo.tbForma_pgto.Post;

    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Cartao_Credito;
    frmModulo.tbForma_pgto.Post;

    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Crediario;
    frmModulo.tbForma_pgto.Post;

    //Convenio
    Inc(c);
    frmModulo.tbForma_pgto.Insert;
    frmModulo.tbForma_Pgtoid.Value := c;
    frmModulo.tbForma_PgtoNome.Value := lForma_pgto_Convenio;
    frmModulo.tbForma_pgto.Post;
    Result := True;
  except
    Result := False;
    ShowMessage('Erro ao carregar formas de pagamento');
  end;
end;

function TfrmPrincipal.ApenasNumerosStr(pStr: string): string;
var
  I: Integer;
begin

  Result := '';

  for I := 1 to Length(pStr) do
    if pStr[I] in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'] then
      Result := Result + pStr[I];

end;

function TfrmPrincipal.IsInteger(s: string): integer;
begin
  try
    Result := strtoint(s);

  except
    Result := 0;
  end;
end;

procedure TfrmPrincipal.tlPhotosClick(Sender: TdxTileControlItem);
begin
  CarregaSistemaVenda;
end;

procedure TfrmPrincipal.tlAgentsClick(Sender: TdxTileControlItem);
begin
  Close;
end;

procedure TfrmPrincipal.dxtlcntrltmTileItem1Click(Sender: TdxTileControlItem);
begin
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;

end;

procedure TfrmPrincipal.tlSystemInformationClick(Sender: TdxTileControlItem);
begin
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;
end;

procedure TfrmPrincipal.dxTileItem1Click(Sender: TdxTileControlItem);
begin
  frmconfig_balanca := tfrmconfig_balanca.create(self);
  frmconfig_balanca.showmodal;
end;

procedure TfrmPrincipal.V1Click(Sender: TObject);
begin
  CarregaSistemaVenda
end;

procedure TfrmPrincipal.S1Click(Sender: TObject);
begin
  Close
end;

procedure TfrmPrincipal.C1Click(Sender: TObject);
begin
  frmconfig := tfrmconfig.create(self);
  frmconfig.showmodal;
end;

procedure TfrmPrincipal.B1Click(Sender: TObject);
begin
  frmconfig_balanca := tfrmconfig_balanca.create(self);
  frmconfig_balanca.showmodal;
end;

function TfrmPrincipal.StrZeros(Texto: string; Qtde: Integer): string;
var
  i, Tam: Integer;
  Aux: string;
begin
  Aux := '';
  Tam := Length(Texto);
  for i := Tam to (Qtde - 1) do
    Aux := Aux + '0';
  StrZeros := Aux + Texto;
end;


end.

