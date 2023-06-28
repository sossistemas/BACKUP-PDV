unit FechamentoCego;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvGlowButton, Vcl.StdCtrls, JvExMask,
  JvToolEdit, Vcl.Mask, RzEdit, Vcl.ExtCtrls, AdvMetroButton, AdvSmoothPanel,
  AdvSmoothExpanderPanel, ACBrBase, ACBrEnterTab, Data.DB, MemDS, DBAccess, Uni,
  frxClass, frxExportBaseDialog, frxExportPDF, frxDBSet, System.IniFiles, StrUtils,
  Datasnap.DBClient;

type
  TfrmFechamentoCego = class(TForm)
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label1: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    Panel1: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    ed_operador: TRzEdit;
    ed_data: TJvDateEdit;
    ed_ecf: TRzEdit;
    bt_cupom_encerrante: TButton;
    Panel2: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    AdvSmoothExpanderPanel2: TAdvSmoothExpanderPanel;
    AdvGlowButton1: TAdvGlowButton;
    ACBrEnterTab1: TACBrEnterTab;
    query: TUniQuery;
    queryCAIXA_SITUACAO: TStringField;
    queryNUMCAIXA: TIntegerField;
    queryCAIXA_DATA_MOVTO: TDateField;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    edCartaoCredito: TRzEdit;
    frxReport1: TfrxReport;
    frxDadosFech: TfrxDBDataset;
    frxPDFExport1: TfrxPDFExport;
    frxEmitente: TfrxDBDataset;
    qryValoresSistema: TUniQuery;
    cdsTempDados: TClientDataSet;
    cdsTempDadosCODOPERADOR: TIntegerField;
    cdsTempDadosNOMEOPERADOR: TStringField;
    cdsTempDadosDINHEIRO_SISTEMA: TCurrencyField;
    cdsTempDadosDINHEIRO_INF: TCurrencyField;
    queryFECHAMENTO: TDateTimeField;
    cdsTempDadosCONVENIO_SISTEMA: TCurrencyField;
    cdsTempDadosCONVENIO_INF: TCurrencyField;
    cdsTempDadosCARTEIRA_SISTEMA: TCurrencyField;
    cdsTempDadosCARTEIRA_INF: TCurrencyField;
    cdsTempDadosC_CRED_SISTEMA: TCurrencyField;
    cdsTempDadosC_CRED_INF: TCurrencyField;
    cdsTempDadosC_DEB_SISTEMA: TCurrencyField;
    cdsTempDadosC_DEB_INF: TCurrencyField;
    cdsTempDadosCHEQUE_SISTEMA: TCurrencyField;
    cdsTempDadosCHEQUE_INF: TCurrencyField;
    cdsTempDadosESTORNO_SISTEMA: TCurrencyField;
    cdsTempDadosESTORNO_INF: TCurrencyField;
    cdsTempDadosCUPOM_CRED_SISTEMA: TCurrencyField;
    cdsTempDadosCUPOM_CRED_INF: TCurrencyField;
    cdsTempDadosDIF_DINHEIRO: TCurrencyField;
    cdsTempDadosDIF_CONVENIO: TCurrencyField;
    cdsTempDadosDIF_CARTEIRA: TCurrencyField;
    cdsTempDadosDIF_CART_CRED: TCurrencyField;
    cdsTempDadosDIF_CART_DEB: TCurrencyField;
    cdsTempDadosDIF_CHEQUE: TCurrencyField;
    cdsTempDadosDIF_ESTORNO: TCurrencyField;
    cdsTempDadosDIF_CUPOM_CRED: TCurrencyField;
    edDinheiro: TRzEdit;
    edCheque: TRzEdit;
    edConvenio: TRzEdit;
    edCartaoDebito: TRzEdit;
    edCrediario: TRzEdit;
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cdsTempDadosCalcFields(DataSet: TDataSet);
    procedure edConvenioExit(Sender: TObject);
    procedure edConvenioKeyPress(Sender: TObject; var Key: Char);
    procedure edCrediarioExit(Sender: TObject);
    procedure edCrediarioKeyPress(Sender: TObject; var Key: Char);
    procedure edDinheiroKeyPress(Sender: TObject; var Key: Char);
    procedure edDinheiroExit(Sender: TObject);
    procedure edChequeExit(Sender: TObject);
    procedure edChequeKeyPress(Sender: TObject; var Key: Char);
    procedure edCartaoCreditoKeyPress(Sender: TObject; var Key: Char);
    procedure edCartaoCreditoExit(Sender: TObject);
    procedure edCartaoDebitoExit(Sender: TObject);
    procedure edCartaoDebitoKeyPress(Sender: TObject; var Key: Char);
    procedure frxReport1BeforePrint(Sender: TfrxReportComponent);
  private
    FHora: string;
    FDataUltFecha: TDateTime;
    FDataFechamento: TDateTime;
    FCodOperador: integer;
    FNomeOperador: string;

    function impressaoFechamento: Boolean;
    procedure Insere_cdsTempDados(aNomeCampoSis, aNomeCampoInf : string; VlrSistema, VlrInf :Currency);

    function CarregaFormaPgto(aForma:string) : string;

    property DataFechamento: TDateTime read FDataFechamento write FDataFechamento;
    property Hora: string read FHora write FHora;
    property DataUltFecha: TDateTime read FDataUltFecha write FDataUltFecha;

    property CodOperador: integer read FCodOperador write FCodOperador;
    property NomeOperador: string read FNomeOperador write FNomeOperador;

  public
    { Public declarations }
  end;

var
  frmFechamentoCego: TfrmFechamentoCego;

implementation

uses
  modulo, principal, senha_supervisor, venda;

{$R *.dfm}

procedure TfrmFechamentoCego.AdvGlowButton1Click(Sender: TObject);
begin
  if application.messagebox(pwidechar('Atenção!'+#13+
                                      'Deseja efetuar o fechamento do Caixa?'),
                                      'Atenção',mb_yesno+mb_iconwarning+MB_DEFBUTTON2) = idyes then
  begin
    with frmModulo do
    begin
      FHora           := FormatDateTime('HH:MM:SS',Time);
      FDataFechamento := Now;

      query.Close;
      query.sql.clear;
      query.sql.add('insert into fechamento_cego (Data, hora, operador, dinheiro, cheque, cartao_credito, cartao_debito, convenio, crediario, ex) values (:Data, :hora, :operador, :dinheiro, :cheque, :cartao_credito, :cartao_debito, :convenio, :crediario, :ex)');
      query.ParamByName('ex').asstring            := 'N';
      query.ParamByName('data').asdatetime        := FDataFechamento; //ed_data.Date;
      query.ParamByName('hora').AsString          := FHora;
      query.ParamByName('operador').AsInteger     := icodigo_Usuario;
      query.ParamByName('dinheiro').AsFloat       := StrToFloat(edDinheiro.Text);
      query.ParamByName('cheque').AsFloat         := StrToFloat(edCheque.Text);
      query.ParamByName('cartao_credito').AsFloat := StrToFloat(edCartaoCredito.Text);
      query.ParamByName('cartao_debito').AsFloat  := StrToFloat(edCartaoDebito.Text);
      query.ParamByName('convenio').AsFloat       := StrToFloat(edConvenio.Text);
      query.ParamByName('crediario').AsFloat      := StrToFloat(edCrediario.Text);
      query.ExecSQL;

      query.Close;
      query.sql.clear;
      query.sql.add('update config set  caixa_situacao = ''FECHADO'',');
      query.sql.add('caixa_data_movto = :data, ');
      query.sql.Add('fechamento = :datafechamento');
      query.ParamByName('datafechamento').asstring := formatdatetime('yyyy-mm-dd hh:mm:ss', now);
      query.ParamByName('data').asdatetime := ed_data.Date;
      query.ExecSQL;
      Application.MessageBox('Procedimento concluído com sucesso!','Aviso',mb_ok+MB_ICONINFORMATION);

      if impressaoFechamento then
      begin
        frxReport1.LoadFromFile(ExtractFilePath(application.ExeName) + '\rel\F000011.fr3');
        frxReport1.ShowReport;
//        frxReport1.DesignReport;
      end;

    end;
    CLOSE;
      if frmVenda <> nil then
        if FRMVENDA.Visible then
          FRMVENDA.CLOSE;
  end;

end;

procedure TfrmFechamentoCego.AdvMetroButton1Click(Sender: TObject);
begin
  Close;
end;

function TfrmFechamentoCego.CarregaFormaPgto(aForma:string): string;
var
  vQry : TUniQuery;
begin

  vQry := TUniQuery.Create(Self);

  Try
    vQry.Connection := frmModulo.conexao;

    vQry.Close;
    vQry.SQL.Clear;
    vQry.SQL.Add('select config.forma_crediario,');
    vQry.SQL.Add('       config.forma_cheque,');
    vQry.SQL.Add('       config.forma_cartao,');
    vQry.SQL.Add('       config.forma_convenio,');
    vQry.SQL.Add('       config.forma_dinheiro,');
    vQry.SQL.Add('       config.forma_cartao_credito,');
    vQry.SQL.Add('       config.forma_cheque_aprazo');
    vQry.SQL.Add('from config');
    vQry.SQL.Add('where config.codigo = 0');
    vQry.Open;

    if aForma = vQry.FieldByName('FORMA_CREDIARIO').Value then
      Result := 'CREDIARIO'
    else if aForma = vQry.FieldByName('FORMA_CHEQUE').Value then
      Result := 'CHEQUE'
    else if aForma = vQry.FieldByName('FORMA_CARTAO').Value then
      Result := 'DEBITO'
    else if aForma = vQry.FieldByName('FORMA_CONVENIO').Value then
      Result := 'CONVENIO'
    else if aForma = vQry.FieldByName('FORMA_DINHEIRO').Value then
      Result := 'DINHEIRO'
    else if aForma = vQry.FieldByName('FORMA_CARTAO_CREDITO').Value then
      Result := 'CREDITO'
    else if aForma = vQry.FieldByName('FORMA_CHEQUE_APRAZO').Value then
      Result := 'CHEQUE A PRAZO';

  Finally
    vQry.Free;
  End;

end;

procedure TfrmFechamentoCego.cdsTempDadosCalcFields(DataSet: TDataSet);
begin
  if cdsTempDados.State = dsCalcFields then
  begin
    cdsTempDadosDIF_DINHEIRO.AsCurrency   := cdsTempDadosDINHEIRO_INF.AsCurrency - cdsTempDadosDINHEIRO_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CONVENIO.AsCurrency   := cdsTempDadosCONVENIO_INF.AsCurrency - cdsTempDadosCONVENIO_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CARTEIRA.AsCurrency   := cdsTempDadosCARTEIRA_INF.AsCurrency - cdsTempDadosCARTEIRA_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CART_CRED.AsCurrency  := cdsTempDadosC_CRED_INF.AsCurrency - cdsTempDadosC_CRED_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CART_DEB.AsCurrency   := cdsTempDadosC_DEB_INF.AsCurrency - cdsTempDadosC_DEB_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CHEQUE.AsCurrency     := cdsTempDadosCHEQUE_INF.AsCurrency - cdsTempDadosCHEQUE_SISTEMA.AsCurrency;
    cdsTempDadosDIF_ESTORNO.AsCurrency    := cdsTempDadosESTORNO_INF.AsCurrency - cdsTempDadosESTORNO_SISTEMA.AsCurrency;
    cdsTempDadosDIF_CUPOM_CRED.AsCurrency := cdsTempDadosCUPOM_CRED_INF.AsCurrency - cdsTempDadosCUPOM_CRED_SISTEMA.AsCurrency;
  end;
end;

procedure TfrmFechamentoCego.edCartaoCreditoExit(Sender: TObject);
begin
  if edCartaoCredito.Text = EmptyStr then
    edCartaoCredito.Text := '0';

  edCartaoCredito.Text := FormatFloat('0.00',StrToFloat(edCartaoCredito.Text));
end;

procedure TfrmFechamentoCego.edCartaoCreditoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.edCartaoDebitoExit(Sender: TObject);
begin
  if edCartaoDebito.Text = EmptyStr then
    edCartaoDebito.Text := '0';

  edCartaoDebito.Text := FormatFloat('0.00',StrToFloat(edCartaoDebito.Text));
end;

procedure TfrmFechamentoCego.edCartaoDebitoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.edChequeExit(Sender: TObject);
begin
  if edCheque.Text = EmptyStr then
    edCheque.Text := '0';

  edCheque.Text := FormatFloat('0.00',StrToFloat(edCheque.Text));
end;

procedure TfrmFechamentoCego.edChequeKeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.edConvenioExit(Sender: TObject);
begin
  if edConvenio.Text = EmptyStr then
    edConvenio.Text := '0';

  edConvenio.Text := FormatFloat('0.00',StrToFloat(edConvenio.Text));
end;

procedure TfrmFechamentoCego.edConvenioKeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.edCrediarioExit(Sender: TObject);
begin
  if edCrediario.Text = EmptyStr then
    edCrediario.Text := '0';

  edCrediario.Text := FormatFloat('0.00',StrToFloat(edCrediario.Text));
end;

procedure TfrmFechamentoCego.edCrediarioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.edDinheiroExit(Sender: TObject);
begin
  if edDinheiro.Text = EmptyStr then
    edDinheiro.Text := '0';

  edDinheiro.Text := FormatFloat('0.00',StrToFloat(edDinheiro.Text));
end;

procedure TfrmFechamentoCego.edDinheiroKeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9',',',#13,#08,#09]) then
    Key := #0;
end;

procedure TfrmFechamentoCego.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  cdsTempDados.Close;
end;

procedure TfrmFechamentoCego.FormCreate(Sender: TObject);
begin
  cdsTempDados.CreateDataSet;
  cdsTempDados.Open;
end;

procedure TfrmFechamentoCego.FormShow(Sender: TObject);
begin
  query.Close;
  query.SQL.Clear;
  query.SQL.Add('select c.caixa_situacao,');
  query.SQL.Add('       c.numcaixa,');
  query.SQL.Add('       c.caixa_data_movto,');
  query.SQL.Add('       c.fechamento');
  query.SQL.Add('from config c');
  query.SQL.Add('where c.codigo = 0');
  query.Open;
  query.First;

  ed_data.Date     := queryCAIXA_DATA_MOVTO.AsDateTime;
  ed_operador.Text := sNome_Operador;
  ed_ecf.Text      := sCaixa;

  FDataUltFecha    := queryFECHAMENTO.Value;
end;

procedure TfrmFechamentoCego.frxReport1BeforePrint(Sender: TfrxReportComponent);
begin
  if TfrxMemoView(Sender).Name = 'mCaixa' then
    TfrxMemoView(Sender).Text := 'Caixa: ' + IntToStr(iNumCaixa);

  if TfrxMemoView(Sender).Name = 'mOperador' then
    TfrxMemoView(Sender).Text := 'Operador: ' + ed_operador.Text;
end;

function  TfrmFechamentoCego.impressaoFechamento: Boolean;
var
  impFechamento : Boolean;
  Ini: TIniFile;

  qryFechamento : TUniQuery;
  vDTMovimento : TDateTime;

  vDINHEIRO, vCONVENIO, vCARTEIRA, vCARTAOCRED, vCARTAODEB, vCHEQUE : Currency;

  i: Integer;
begin
  Result := False;

  vDINHEIRO   := 0;
  vCONVENIO   := 0;
  vCARTEIRA   := 0;
  vCARTAOCRED := 0;
  vCARTAODEB  := 0;
  vCHEQUE     := 0;

  Ini := TIniFile.Create(sConfiguracoes);

  qryFechamento := TUniQuery.Create(Self);

  try
    impFechamento := Ini.ReadBool('Fortes','ImprimirFechamentoCego', False);

    if impFechamento then
    begin
      qryFechamento.Connection := frmModulo.conexao;

      qryFechamento.Close;
      qryFechamento.SQL.Clear;
      qryFechamento.SQL.Add('select first 1 fc.data,');
      qryFechamento.SQL.Add('       fc.hora,');
      qryFechamento.SQL.Add('       fc.operador codOperador,');
      qryFechamento.SQL.Add('       adm.info1 operador,');
      qryFechamento.SQL.Add('       fc.dinheiro,');
      qryFechamento.SQL.Add('       fc.cheque,');
      qryFechamento.SQL.Add('       fc.ex,');
      qryFechamento.SQL.Add('       fc.cartao_credito,');
      qryFechamento.SQL.Add('       fc.cartao_debito,');
      qryFechamento.SQL.Add('       fc.convenio,');
      qryFechamento.SQL.Add('       fc.crediario');
      qryFechamento.SQL.Add('from fechamento_cego fc');
      qryFechamento.SQL.Add('join adm on adm.codigo = fc.operador');
      qryFechamento.SQL.Add('where fc.data = :dt');
      qryFechamento.SQL.Add('and cast(fc.hora as time) =:hr');
      qryFechamento.ParamByName('dt').AsDateTime := FDataFechamento;
      qryFechamento.ParamByName('hr').Value      := StrToTime(FHora);
      qryFechamento.Open;

      FCodOperador  := qryFechamento.FieldByName('codOperador').AsInteger;
      FNomeOperador := qryFechamento.FieldByName('operador').AsString;

      qryValoresSistema.SQL.Clear;
      qryValoresSistema.SQL.Add('with C as (');
      qryValoresSistema.SQL.Add('  select CF.FORMA as forma,');
      qryValoresSistema.SQL.Add('         sum(CF.VALOR) soma_valor');
      qryValoresSistema.SQL.Add('  from cupom C');
      qryValoresSistema.SQL.Add('  join cupom_forma CF on CF.cod_cupom = C.CODIGO');
      qryValoresSistema.SQL.Add('  where C.cancelado <> 1');
      qryValoresSistema.SQL.Add('  and C.DATA + C.hora >= :data');
      qryValoresSistema.SQL.Add('  group by 1)');
      qryValoresSistema.SQL.Add('select C.FORMA,');
      qryValoresSistema.SQL.Add('       C.soma_valor');
      qryValoresSistema.SQL.Add('From C');
      qryValoresSistema.ParamByName('data').Value := FDataUltFecha;
      qryValoresSistema.Open;

      while not qryValoresSistema.Eof do
      begin
        case AnsiIndexStr(UpperCase(qryValoresSistema.FieldByName('FORMA').Value), ['DINHEIRO','CONVENIO','CARTEIRA','CARTAO CREDITO','CARTAO DEBITO','CHEQUE']) of
        0: vDINHEIRO   := qryValoresSistema.FieldByName('soma_valor').Value;
        1: vCONVENIO   := qryValoresSistema.FieldByName('soma_valor').Value;
        2: vCARTEIRA   := qryValoresSistema.FieldByName('soma_valor').Value;
        3: vCARTAOCRED := qryValoresSistema.FieldByName('soma_valor').Value;
        4: vCARTAODEB  := qryValoresSistema.FieldByName('soma_valor').Value;
        5: vCHEQUE     := qryValoresSistema.FieldByName('soma_valor').Value;
        end;

        qryValoresSistema.Next;
      end;

      for i := 0 to 5 do
      begin
        case i of
        0: Insere_cdsTempDados('DINHEIRO_SISTEMA','DINHEIRO_INF',vDINHEIRO,qryFechamento.FieldByName('dinheiro').Value);
        1: Insere_cdsTempDados('CONVENIO_SISTEMA','CONVENIO_INF',vCONVENIO,qryFechamento.FieldByName('convenio').Value);
        2: Insere_cdsTempDados('CARTEIRA_SISTEMA','CARTEIRA_INF',vCARTEIRA,qryFechamento.FieldByName('crediario').Value);
        3: Insere_cdsTempDados('C_CRED_SISTEMA','C_CRED_INF',vCARTAOCRED,qryFechamento.FieldByName('cartao_credito').Value);
        4: Insere_cdsTempDados('C_DEB_SISTEMA','C_DEB_INF',vCARTAODEB,qryFechamento.FieldByName('cartao_debito').Value);
        5: Insere_cdsTempDados('CHEQUE_SISTEMA','CHEQUE_INF',vCHEQUE,qryFechamento.FieldByName('cheque').Value);
        end;
      end;

      cdsTempDados.First;

      Result := True;
    end;

  finally
    Ini.Free;
    qryFechamento.Free;
  end;
end;

procedure TfrmFechamentoCego.Insere_cdsTempDados(aNomeCampoSis, aNomeCampoInf : string; VlrSistema, VlrInf :Currency);
var
  i : integer;
begin
  if cdsTempDados.RecordCount = 0 then
  begin
    cdsTempDados.Append;
    cdsTempDadosCODOPERADOR.Value  := FCodOperador;
    cdsTempDadosNOMEOPERADOR.Value :=FNomeOperador;

    for i := 2 to cdsTempDados.Fields.Count -1 do
    begin
      if cdsTempDados.Fields[i].FieldName = aNomeCampoSis then
        cdsTempDados.Fields[i].Value := VlrSistema
      else if cdsTempDados.Fields[i].FieldName = aNomeCampoInf then
        cdsTempDados.Fields[i].Value := VlrInf
      else if not cdsTempDados.Fields[i].Calculated then
        cdsTempDados.Fields[i].Value := 0;
    end;

    cdsTempDados.Post;
  end
  else
  begin
    cdsTempDados.Edit;
    cdsTempDadosCODOPERADOR.Value  := FCodOperador;
    cdsTempDadosNOMEOPERADOR.Value :=FNomeOperador;

    for i := 2 to cdsTempDados.Fields.Count -1 do
    begin
      if cdsTempDados.Fields[i].FieldName = aNomeCampoSis then
        cdsTempDados.Fields[i].Value := VlrSistema
      else if cdsTempDados.Fields[i].FieldName = aNomeCampoInf then
        cdsTempDados.Fields[i].Value := VlrInf
      else if not cdsTempDados.Fields[i].Calculated then
        cdsTempDados.Fields[i].Value := cdsTempDados.Fields[i].Value;
    end;

    cdsTempDados.Post;
  end;

end;

end.
