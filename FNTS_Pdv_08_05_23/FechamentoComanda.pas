unit FechamentoComanda;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvMetroButton, Vcl.StdCtrls,
  AdvSmoothPanel, AdvSmoothExpanderPanel, Vcl.ExtCtrls, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter,
  cxData, cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData, cxCheckBox,
  cxTextEdit, cxContainer, Vcl.Menus, cxButtons, cxMaskEdit, cxSpinEdit,
  cxClasses, cxGridCustomTableView, cxGridTableView, cxGridDBTableView, MemDS,
  DBAccess, Uni, cxGridLevel, cxGridCustomView, cxGrid, frxClass, frxDBSet;

type
  TfrmFechamentoComanda = class(TForm)
    Panel1: TPanel;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    lbTitulo: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    AdvMetroButton2: TAdvMetroButton;
    AdvMetroButton3: TAdvMetroButton;
    Panel2: TPanel;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    Panel3: TPanel;
    qrFechamento: TUniQuery;
    qrFechamentoSEL: TStringField;
    qrFechamentoMESA: TIntegerField;
    qrFechamentoCOMANDA: TIntegerField;
    qrFechamentoPERGUNTOU_COMISSAO: TStringField;
    qrFechamentoCLIENTE_AUTORIZOU: TStringField;
    qrFechamentoSUB_TOTAL: TFloatField;
    qrFechamentoCOUVER: TFloatField;
    qrFechamentoCOMISSAO: TFloatField;
    qrFechamentoTOTAL: TFloatField;
    dsFechamento: TDataSource;
    cxGrid1DBTableView1SEL: TcxGridDBColumn;
    cxGrid1DBTableView1MESA: TcxGridDBColumn;
    cxGrid1DBTableView1COMANDA: TcxGridDBColumn;
    cxGrid1DBTableView1PERGUNTOU_COMISSAO: TcxGridDBColumn;
    cxGrid1DBTableView1CLIENTE_AUTORIZOU: TcxGridDBColumn;
    cxGrid1DBTableView1SUB_TOTAL: TcxGridDBColumn;
    cxGrid1DBTableView1COUVER: TcxGridDBColumn;
    cxGrid1DBTableView1COMISSAO: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    Content: TcxStyle;
    Header: TcxStyle;
    cxStyle1: TcxStyle;
    cxStyle2: TcxStyle;
    edComanda: TcxSpinEdit;
    Label1: TLabel;
    cxButton1: TcxButton;
    cxCheckBox1: TcxCheckBox;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    cxButton4: TcxButton;
    qrTmp: TUniQuery;
    Label2: TLabel;
    edMesa: TcxSpinEdit;
    cxButton5: TcxButton;
    imgFundo: TImage;
    frFilial: TfrxDBDataset;
    frComanda: TfrxReport;
    qrFilial: TUniQuery;
    qrFilialCODIGO: TStringField;
    qrFilialFILIAL: TStringField;
    qrFilialNOTAFISCAL: TIntegerField;
    qrFilialENDERECO: TStringField;
    qrFilialCIDADE: TStringField;
    qrFilialUF: TStringField;
    qrFilialCEP: TStringField;
    qrFilialTELEFONE: TStringField;
    qrFilialCNPJ: TStringField;
    qrFilialIE: TStringField;
    qrFilialSEQNF: TIntegerField;
    qrFilialFAX: TStringField;
    qrFilialOBS1: TStringField;
    qrFilialOBS2: TStringField;
    qrFilialCONTRIBUINTE_IPI: TStringField;
    qrFilialSUBSTITUTO_TRIBUTARIO: TStringField;
    qrFilialEMPRESA_ESTADUAL: TStringField;
    qrFilialOPTANTE_SIMPLES: TStringField;
    qrFilialOPTANTE_SUPER_SIMPLES: TStringField;
    qrFilialECF: TStringField;
    qrFilialTIPO: TIntegerField;
    qrFilialIPI: TFloatField;
    qrFilialISS: TFloatField;
    qrFilialNUMERO: TStringField;
    qrFilialRESPONSAVEL: TStringField;
    qrFilialCOMPLEMENTO: TStringField;
    qrFilialBAIRRO: TStringField;
    qrFilialFARMACIA_RESP_TECNICO: TStringField;
    qrFilialFARMACIA_CRF: TStringField;
    qrFilialFARMACIA_CPF: TStringField;
    qrFilialFARMCIA_DATA: TDateField;
    qrFilialFARMACIA_UF: TStringField;
    qrFilialFARMACIA_SENHA: TStringField;
    qrFilialFARMACIA_EMAIL: TStringField;
    qrFilialFARMACIA_LOGIN: TStringField;
    qrFilialFARMACIA_ENVIO: TStringField;
    qrFilialCONHECIMENTO: TIntegerField;
    qrFilialINDUSTRIA: TStringField;
    qrFilialFARMACIA_NUMEROLICENCA: TStringField;
    qrFilialCOD_MUNICIPIO_IBGE: TStringField;
    qrFilialIBGE: TStringField;
    qrFilialPIS: TFloatField;
    qrFilialCOFINS: TFloatField;
    qrFilialEMAIL: TStringField;
    qrFilialATIVIDADE: TStringField;
    qrFilialCONTADOR_COD_MUNICIPIO_IBGE: TStringField;
    qrFilialCONTADOR_NOME: TStringField;
    qrFilialCONTADOR_CPF: TStringField;
    qrFilialCONTADOR_CRC: TStringField;
    qrFilialCONTADOR_CNPJ: TStringField;
    qrFilialCONTADOR_CEP: TStringField;
    qrFilialCONTADOR_ENDERECO: TStringField;
    qrFilialCONTADOR_NUMERO: TStringField;
    qrFilialCONTADOR_COMPLEMENTO: TStringField;
    qrFilialCONTADOR_BAIRRO: TStringField;
    qrFilialCONTADOR_FONE: TStringField;
    qrFilialCONTADOR_FAX: TStringField;
    qrFilialCONTADOR_EMAIL: TStringField;
    qrFilialINSC_MUNICIPAL: TStringField;
    qrFilialDATA_ABERTURA: TDateField;
    qrFilialCNAE: TStringField;
    qrFilialCRT: TStringField;
    qrFilialCONTADOR_CIDADE: TStringField;
    qrFilialCONTADOR_COD_MUNICIPIO: TStringField;
    qrFilialCONTADOR_UF: TStringField;
    qrFilialPERMITE_CREDITO: TIntegerField;
    qrFilialFANTASIA: TStringField;
    qrFilialDFIXAS: TFloatField;
    qrFilialTIPOCALCULO: TIntegerField;
    qrFilialSERIE_CTE: TStringField;
    qrFilialSEQ_CTE: TStringField;
    qrFilialCOD_PAIS: TStringField;
    qrFilialPAIS: TStringField;
    qrFilialHOMEPAGE: TStringField;
    qrFilialAIDF: TStringField;
    qrFilialNINICIAL: TStringField;
    qrFilialNFINAL: TStringField;
    qrFilialN_NF_D: TStringField;
    qrFilialVALIDADE: TStringField;
    qrFilialRNTRC: TStringField;
    qrFilialSERIE_MDFE: TIntegerField;
    qrFilialNUMERO_MDFE: TIntegerField;
    qrFilialMAIL_SERVER: TStringField;
    qrFilialMAIL_PORTA: TStringField;
    qrFilialMAIL_USERNAME: TStringField;
    qrFilialMAIL_PASSWORD: TStringField;
    qrFilialMAIL_CABECALHO: TStringField;
    qrFilialMAIL_ASSUNTO: TStringField;
    qrFilialMAIL_MENSAGEM: TBlobField;
    qrFilialMAIL_SSL: TStringField;
    qrFilialMAIL_TLS: TStringField;
    frProdutosComanda: TfrxDBDataset;
    qrProdutosComanda: TUniQuery;
    qrProdutosComandaCODIGO: TStringField;
    qrProdutosComandaDATA: TDateField;
    qrProdutosComandaHORA: TStringField;
    qrProdutosComandaPERGUNTOU_COMISSAO: TStringField;
    qrProdutosComandaCLIENTE_AUTORIZOU: TStringField;
    qrProdutosComandaCOD_FUNCIONARIO: TStringField;
    qrProdutosComandaNOME: TStringField;
    qrProdutosComandaCOD_PRODUTO: TStringField;
    qrProdutosComandaPRODUTO: TStringField;
    qrProdutosComandaCOD_LANCAMENTO: TIntegerField;
    qrProdutosComandaQTDE: TFloatField;
    qrProdutosComandaUNITARIO: TFloatField;
    qrProdutosComandaTOTAL: TFloatField;
    qrProdutosComandaIMPRESSO: TStringField;
    qrProdutosComandaTRANSF_MESA: TIntegerField;
    qrProdutosComandaCOMPLEMENTO: TStringField;
    qrProdutosComandaEXTRA: TBlobField;
    procedure AdvMetroButton3Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
    procedure cxCheckBox1Click(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure frComandaBeforePrint(Sender: TfrxReportComponent);
  private
    { Private declarations }
    procedure AbreFechamento;
  public
    { Public declarations }
    strCodigos:String;
  end;

var
  frmFechamentoComanda: TfrmFechamentoComanda;

implementation

uses modulo, principal, funcoes;

{$R *.dfm}

procedure TfrmFechamentoComanda.AbreFechamento;
var
  Filtro:string;
begin
  try
    qrFechamento.DisableControls;
    qrFechamento.Close;
    Filtro := ' and com.codigo in ('+strCodigos+')';
    qrFechamento.MacroByName('Filtro').Value := Filtro;
    qrFechamento.ParamByName('atb').AsString := ME_FiltraATB('TB_CONFIGURACOES');
    qrFechamento.Open;
  finally
    qrFechamento.EnableControls;
  end;
end;

procedure TfrmFechamentoComanda.AdvMetroButton3Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmFechamentoComanda.cxButton1Click(Sender: TObject);
begin
  if edComanda.Value = 0 then begin
    Application.MessageBox('Informe o Número da Comanda!','Atenção!',MB_ICONINFORMATION);
    edComanda.SetFocus;
    exit;
  end;
  if qrFechamento.Locate('comanda',edComanda.Value,[loCaseInsensitive]) then begin
    Application.MessageBox('Comanda já está na listagem de Fechamento!','Atenção!',MB_ICONINFORMATION);
    edComanda.SetFocus;
    exit;
  end;
  qrTmp.Close;
  qrTmp.SQL.Clear;
  qrTmp.SQL.Add('select count(*) qtd from r000001 where comanda = ' + IntToStr(edComanda.Value));
  qrTmp.Open;
  if qrTmp.FieldByName('qtd').AsInteger = 0 then begin
    Application.MessageBox('Comanda não está aberta!','Atenção!',MB_ICONINFORMATION);
    edComanda.SetFocus;
    exit;
  end;
  strCodigos := strCodigos + ', '+IntToStr(qrTmp.FieldByName('codigo').AsInteger);
  AbreFechamento;
end;

procedure TfrmFechamentoComanda.cxButton2Click(Sender: TObject);
var
  Local:Integer;
begin
  qrTmp.Close;
  qrTmp.SQL.Clear;
  if qrFechamentoPERGUNTOU_COMISSAO.AsString = 'S' then
    if qrFechamentoCOMANDA.AsInteger > 0 then
      qrTmp.SQL.Add('update r000001 set perguntou_comissao = ' + QuotedStr('N')+' where comanda = ' + qrFechamentoCOMANDA.AsString)
    else
      qrTmp.SQL.Add('update r000001 set perguntou_comissao = ' + QuotedStr('N')+' where mesa = ' + qrFechamentoMESA.AsString)
  else
    if qrFechamentoCOMANDA.AsInteger > 0 then
      qrTmp.SQL.Add('update r000001 set perguntou_comissao = ' + QuotedStr('S')+' where comanda = ' + qrFechamentoCOMANDA.AsString)
    else
      qrTmp.SQL.Add('update r000001 set perguntou_comissao = ' + QuotedStr('S')+' where mesa = ' + qrFechamentoMESA.AsString);
  qrTmp.Execute;

  try
    Local := qrFechamentoCOMANDA.AsInteger;
    qrFechamento.DisableControls;
    qrFechamento.Close;
    qrFechamento.Open;
    qrFechamento.Locate('comanda',Local,[loCaseInsensitive]);
  finally
    qrFechamento.EnableControls;
  end;
end;

procedure TfrmFechamentoComanda.cxButton3Click(Sender: TObject);
var
  Local:Integer;
begin
  qrTmp.Close;
  qrTmp.SQL.Clear;
  if qrFechamentoPERGUNTOU_COMISSAO.AsString = 'S' then
    if qrFechamentoCOMANDA.AsInteger > 0 then
      qrTmp.SQL.Add('update r000001 set cliente_autorizou = ' + QuotedStr('N')+' where comanda = ' + qrFechamentoCOMANDA.AsString)
    else
      qrTmp.SQL.Add('update r000001 set cliente_autorizou = ' + QuotedStr('N')+' where mesa = ' + qrFechamentoMESA.AsString)
  else
    if qrFechamentoCOMANDA.AsInteger > 0 then
      qrTmp.SQL.Add('update r000001 set cliente_autorizou = ' + QuotedStr('S')+' where comanda = ' + qrFechamentoCOMANDA.AsString)
    else
      qrTmp.SQL.Add('update r000001 set cliente_autorizou = ' + QuotedStr('S')+' where mesa = ' + qrFechamentoMESA.AsString);
  qrTmp.Execute;


  try
    Local := qrFechamentoCOMANDA.AsInteger;
    qrFechamento.DisableControls;
    qrFechamento.Close;
    qrFechamento.Open;
    qrFechamento.Locate('comanda',Local,[loCaseInsensitive]);
  finally
    qrFechamento.EnableControls;
  end;
end;

procedure TfrmFechamentoComanda.cxButton4Click(Sender: TObject);
var
  Marcado:Boolean;
  Filtro:string;
begin
  try
    qrFechamento.DisableControls;
    qrFechamento.First;
    while not qrFechamento.Eof do begin
      if qrFechamentoSEL.AsString = 'S' then begin
        Marcado := True;
        Break;
      end;
      qrFechamento.Next;
    end;
    if not Marcado then begin
      Application.MessageBox('Não foi selecionado nenhum item para a impressão!','Atenção!',MB_ICONINFORMATION);
      exit;
    end;
    qrFechamento.First;
    qrFilial.Close;
    qrFilial.Open;
    frComanda.LoadFromFile(ExtractFilePath(Application.ExeName) + 'Rel\F000010.fr3');
    while not qrFechamento.Eof do begin
      if qrFechamentoSEL.AsString = 'S' then begin
        qrProdutosComanda.Close;
        if qrFechamentoCOMANDA.AsInteger = 0 then
          Filtro := ' and mesa = ' + qrFechamentoMESA.AsString
        else
          Filtro := ' and comanda = ' + qrFechamentoCOMANDA.AsString;
        qrProdutosComanda.MacroByName('Filtro').Value :=  Filtro;
        qrProdutosComanda.Open;
        if bExibeCumponNFTela then
          frComanda.ShowReport
        else begin
          frComanda.PrintOptions.Printer := sImpressora_Venda_Simples;
          frComanda.PrepareReport(True);
          frComanda.Print
        end;
      end;
      qrFechamento.Next;
    end;
  finally
    qrFechamento.First;
    qrFechamento.EnableControls;
  end;
end;

procedure TfrmFechamentoComanda.cxButton5Click(Sender: TObject);
begin
  if edMesa.Value = 0 then begin
    Application.MessageBox('Informe o Número da Mesa!','Atenção!',MB_ICONINFORMATION);
    edMesa.SetFocus;
    exit;
  end;
  if qrFechamento.Locate('mesa',edMesa.Value,[loCaseInsensitive]) then begin
    Application.MessageBox('Mesa já está na listagem de Fechamento!','Atenção!',MB_ICONINFORMATION);
    edMesa.SetFocus;
    exit;
  end;
  qrTmp.Close;
  qrTmp.SQL.Clear;
  qrTmp.SQL.Add('select count(*) qtd from r000001 where mesa = ' + IntToStr(edMesa.Value));
  qrTmp.Open;
  if qrTmp.FieldByName('qtd').AsInteger = 0 then begin
    Application.MessageBox('Mesa não está aberta!','Atenção!',MB_ICONINFORMATION);
    edMesa.SetFocus;
    exit;
  end;
  strCodigos := strCodigos + ', '+IntToStr(qrTmp.FieldByName('codigo').AsInteger);
  AbreFechamento;
end;

procedure TfrmFechamentoComanda.cxCheckBox1Click(Sender: TObject);
begin
  try
    qrFechamento.DisableControls;
    qrFechamento.First;
    while not qrFechamento.Eof do begin
      qrFechamento.Edit;
      if cxCheckBox1.Checked then
        qrFechamentoSEL.AsString := 'S'
      else
        qrFechamentoSEL.AsString := 'N';
      qrFechamento.Post;
      qrFechamento.Next;
    end;
  finally
    qrFechamento.First;
    qrFechamento.EnableControls;
  end;
end;

procedure TfrmFechamentoComanda.FormCreate(Sender: TObject);
begin
  strCodigos := '';
end;

procedure TfrmFechamentoComanda.FormShow(Sender: TObject);
begin
  AbreFechamento;
end;

procedure TfrmFechamentoComanda.frComandaBeforePrint(
  Sender: TfrxReportComponent);
begin
  if Sender.ClassType = TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'cpCouvert' then 
      TfrxMemoView(Sender).Text := 'R$ ' + FormatFloat('#,##0.00',qrFechamentoCOUVER.AsFloat);
    if TfrxMemoView(Sender).Name = 'cptaxa' then 
      TfrxMemoView(Sender).Text := 'R$ ' + FormatFloat('#,##0.00',qrFechamentoCOMISSAO.AsFloat);    
    if TfrxMemoView(Sender).Name = 'cpTotal' then 
      TfrxMemoView(Sender).Text := 'R$ ' + FormatFloat('#,##0.00',qrFechamentoTOTAL.AsFloat);    
    if TfrxMemoView(Sender).Name = 'cpmesa' then begin
      if (qrFechamentoMESA.AsInteger > 0) and (qrFechamentoCOMANDA.AsInteger > 0) then
        TfrxMemoView(Sender).Text := 'Mesa: ' + zerar(qrFechamentoMESA.AsString,3)+' - Comanda: ' + qrFechamentoCOMANDA.AsString;
      if (qrFechamentoMESA.AsInteger > 0) and (qrFechamentoCOMANDA.AsInteger = 0) then
        TfrxMemoView(Sender).Text := 'Mesa: ' + zerar(qrFechamentoMESA.AsString,3);
      if (qrFechamentoMESA.AsInteger = 0) and (qrFechamentoCOMANDA.AsInteger > 0) then
        TfrxMemoView(Sender).Text := 'Comanda: ' + qrFechamentoCOMANDA.AsString;    
    end;
      
  end;
end;

end.
