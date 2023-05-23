unit venda_crediario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, AdvMenus, StdCtrls, Mask, RzEdit, AdvGlowButton, ExtCtrls,
  RzLabel, TFlatPanelUnit, RzSpnEdt, RzCmboBx, Data.DB, Vcl.DBCtrls, RzDBEdit,
  MemDS, DBAccess, Uni, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxButtonEdit, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxMemo, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxDropDownEdit, cxGroupBox, cxSpinEdit, Vcl.ComCtrls, dxCore, cxDateUtils,
  cxCalendar, AdvMetroButton, System.DateUtils;

type
  TfrmVenda_Crediario = class(TForm)
    pop_consumid: TAdvPopupMenu;
    MenuItem2: TMenuItem;
    Panel2: TPanel;
    DataSource1: TDataSource;
    query: TUniQuery;
    Panel1: TPanel;
    cxGroupBox2: TcxGroupBox;
    RzLabel2: TRzLabel;
    RzLabel1: TRzLabel;
    lb_cred_codigo: TRzLabel;
    lb_cred_nome: TRzLabel;
    lb_cred_endereco: TRzLabel;
    lb_cred_bairro: TRzLabel;
    lb_cred_cidade: TRzLabel;
    lb_cred_uf: TRzLabel;
    lb_cred_cep: TRzLabel;
    lb_cred_cpf: TRzLabel;
    RzLabel3: TRzLabel;
    lb_cred_situacao: TRzLabel;
    RzLabel4: TRzLabel;
    lb_cred_limite: TRzLabel;
    RzLabel5: TRzLabel;
    lb_cred_utilizado: TRzLabel;
    RzLabel6: TRzLabel;
    lb_cred_disponivel: TRzLabel;
    lb_cred_atualizado: TRzLabel;
    edCliente: TcxButtonEdit;
    edPegoPor: TcxComboBox;
    Panel4: TPanel;
    bt_cancelar_crediario: TAdvGlowButton;
    bt_gravar_crediario: TAdvGlowButton;
    edObs: TcxMemo;
    RzLabel8: TRzLabel;
    Panel7: TPanel;
    cxGroupBox3: TcxGroupBox;
    RzLabel9: TRzLabel;
    RzLabel10: TRzLabel;
    RzLabel11: TRzLabel;
    gpParcela: TcxGroupBox;
    Panel5: TPanel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    gpDados: TcxGroupBox;
    RzLabel7: TRzLabel;
    RzLabel12: TRzLabel;
    RzLabel13: TRzLabel;
    edParcelas: TcxSpinEdit;
    edDiaFixo: TcxSpinEdit;
    edIntervalo: TcxSpinEdit;
    RzLabel14: TRzLabel;
    edVencimento: TcxDateEdit;
    Bevel1: TBevel;
    lbTotal: TRzLabel;
    lbTotalParcelar: TRzLabel;
    lbValorParcela: TRzLabel;
    Panel3: TPanel;
    AdvMetroButton1: TAdvMetroButton;
    AdvMetroButton2: TAdvMetroButton;
    pn_veiculo: TFlatPanel;
    RzLabel20: TRzLabel;
    sh_consumid_km: TShape;
    RzLabel21: TRzLabel;
    sh_consumid_placa: TShape;
    RzLabel18: TRzLabel;
    sh_consumid_vendedor: TShape;
    RzLabel19: TRzLabel;
    ed_consumid_km: TRzEdit;
    ed_consumid_placa: TRzEdit;
    ed_consumid_vendedor: TRzEdit;
    dsCrediario: TDataSource;
    cxGrid1DBTableView1parcela: TcxGridDBColumn;
    cxGrid1DBTableView1vencimento: TcxGridDBColumn;
    cxGrid1DBTableView1valor: TcxGridDBColumn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure bt_gravar_crediarioClick(Sender: TObject);
    procedure bt_cancelar_crediarioClick(Sender: TObject);
    procedure bt_gravar_crediarioEnter(Sender: TObject);
    procedure bt_cancelar_crediarioEnter(Sender: TObject);
    function LocalizarCliente: Integer;
    procedure PreencherDadosCliente(pbVazio: Boolean = False);
    procedure edClienteKeyPress(Sender: TObject; var Key: Char);
    procedure edClienteEnter(Sender: TObject);
    procedure edClientePropertiesEditValueChanged(Sender: TObject);
    procedure AdvMetroButton2Click(Sender: TObject);
    procedure edPegoPorKeyPress(Sender: TObject; var Key: Char);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure edIntervaloExit(Sender: TObject);
    procedure edDiaFixoExit(Sender: TObject);
    procedure edParcelasPropertiesChange(Sender: TObject);
    procedure edParcelasKeyPress(Sender: TObject; var Key: Char);
    procedure edDiaFixoKeyPress(Sender: TObject; var Key: Char);
    procedure edIntervaloKeyPress(Sender: TObject; var Key: Char);
    procedure edVencimentoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    ValorReceber:Double;
  end;

var
  frmVenda_Crediario: TfrmVenda_Crediario;
  bcontinua_creadiario: Boolean;
  rdisponivel: real;

implementation

uses
  modulo, cliente_consulta, funcoes, venda, principal, Math, senha_supervisor;

{$R *.dfm}

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := cafree;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.FormShow(Sender: TObject);
begin
  edParcelas.Clear;
  edDiaFixo.Clear;
  edIntervalo.Clear;
  edVencimento.Clear;
  bcontinua_creadiario := false;
  edCliente.SetFocus;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.MenuItem2Click(Sender: TObject);
begin
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.bt_gravar_crediarioClick(Sender: TObject);
var
  bsenha: boolean;
  tot:Double;
begin
  if not bcontinua_creadiario then
    exit;

  with frmVenda do begin
    try
      cdsCrediario.DisableControls;
      cdsCrediario.First;
      tot := 0;
      while not cdsCrediario.Eof do begin
        tot := tot + cdsCrediariovalor.AsFloat;
        cdsCrediario.Next;
      end;
    finally
       cdsCrediario.First;
      cdsCrediario.EnableControls;
    end;
  end;
  if tot <> ValorReceber then begin
    Application.messagebox('Soma das parcelas não bate com o valor total a receber!', 'Atenção', MB_ICONINFORMATION);
    exit;
  end;
  bsenha := false;

  if lb_cred_situacao.caption = '2' then
    bsenha := true;
  if lb_cred_situacao.caption = '3' then begin
    Application.messagebox('Cliente com crédito bloqueado! Impossível prosseguir!', 'Erro', mb_ok + mb_iconerror);
    exit;
  end;

  frmmodulo.query.close;
  frmmodulo.query.sql.clear;
  frmmodulo.query.sql.add('select * from cliente');
  frmmodulo.query.sql.add('where codigo = ' + QuotedStr(lb_cred_codigo.Caption));
  frmmodulo.query.open;


  if (frmmodulo.query.fieldbyname('limite').asfloat > 0.00) and (rdisponivel < rvalor_total_crediario) then begin
    if application.messagebox('Este cliente necessita de autorização para efetuar compra no crediário!'#13 + 'Verifique o limite/situação do cadastro!' + #13 + 'Deseja prosseguir?', 'Atenção', mb_yesno + MB_ICONWARNING) = idyes then begin
      bsenha := true;
    end else
      exit;
  end;

  if bsenha then begin
    frmsenha_supervisor := tfrmsenha_supervisor.create(self);
    frmsenha_supervisor.showmodal;
    if not bSupervisor_autenticado then begin
      application.MessageBox('Não autorizado!', 'Erro', mb_ok + mb_iconerror);
      PreencherDadosCliente(True);
      edCliente.SetFocus;
      exit;
    end;
  end;

  sCli_Nome := lb_cred_nome.caption;
  sCli_Endereco := lb_cred_endereco.caption;
  sCli_Cidade := lb_cred_cidade.caption;
  sCli_uf := lb_cred_uf.caption;
  scli_cep := lb_cred_cep.caption;
  sCli_CPF := lb_cred_cpf.caption;
  sCli_codigo := lb_cred_codigo.caption;

  sCli_Placa := ed_consumid_placa.text;
  sCli_Km := ed_consumid_km.text;
  sCli_vendedor := ed_consumid_vendedor.text;
  sCLi_PegoPor := edPegoPor.Text;
  iCrediario_prestacao := edParcelas.Value;

  bContinua := true;
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.AdvMetroButton1Click(Sender: TObject);
var
  i, dia, intervalo:Integer;
  Data:TDateTime;
  VlrParcela, vlrSaldo:Double;
  vdia, vmes, vano:Word;
begin
  AdvMetroButton1.SetFocus;
  if lb_cred_codigo.Caption = '' then begin
    Application.MessageBox('Informe o cliente!','Atenção!',MB_ICONINFORMATION);
    edCliente.SetFocus;
    Exit;
  end;
  if edParcelas.Value <= 0 then begin
    Application.MessageBox('Informe a quantidade de parcelas!','Atenção!',MB_ICONINFORMATION);
    edParcelas.SetFocus;
    Exit;
  end;
  frmVenda.cdsCrediario.Close;
  frmVenda.cdsCrediario.CreateDataSet;
  dia := edDiaFixo.Value;
  intervalo := edIntervalo.Value;
  if edVencimento.Date > 0 then
    Data := edVencimento.Date
  else begin
    Data := Date;
    if intervalo > 0 then
      Data := IncDay(Data,intervalo)
    else begin
      if dia > 0 then begin
        DecodeDate(Data,vano, vmes, vdia);
        Data := EncodeDate(vano,vmes,dia);
        Data :=  IncMonth(Data);
      end else
        Data := IncMonth(Data);
    end;
  end;

  VlrParcela := RoundTo(ValorReceber/edParcelas.Value,-2);
  vlrSaldo := RoundTo(ValorReceber - (VlrParcela * edParcelas.Value),-2);
  for i:=0 to edParcelas.Value -1 do begin
    frmVenda.cdsCrediario.Append;
    frmVenda.cdsCrediarioparcela.AsInteger := i + 1;
    frmVenda.cdsCrediariocliente.AsString := lb_cred_codigo.caption;
    frmVenda.cdsCrediariovencimento.AsDateTime := Data;
    if (vlrSaldo <> 0) and (i = edParcelas.Value -1) then
      frmVenda.cdsCrediariovalor.AsFloat := VlrParcela+vlrSaldo
    else
      frmVenda.cdsCrediariovalor.AsFloat := VlrParcela;
    frmVenda.cdsCrediario.Post;
    if intervalo > 0 then
      Data := IncDay(Data,intervalo)
    else begin
      if dia > 0 then begin
        DecodeDate(Data,vano, vmes, vdia);
        Data := EncodeDate(vano,vmes,dia);
        Data :=  IncMonth(Data);
      end else
        Data := IncMonth(Data);
    end;

  end;
  bcontinua_creadiario := True;
end;

procedure TfrmVenda_Crediario.AdvMetroButton2Click(Sender: TObject);
begin
  bcontinua_creadiario := False;
  frmVenda.cdsCrediario.Close;
  frmVenda.cdsCrediario.CreateDataSet;
  edParcelas.Value := 1;
  edDiaFixo.Value := DayOfWeek(Date);
  edIntervalo.Value := 0;
  edVencimento.Date := IncMonth(Date);
  edParcelas.SetFocus;
end;

procedure TfrmVenda_Crediario.bt_cancelar_crediarioClick(Sender: TObject);
begin
  sCli_codigo := '';
  bContinua := false;
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.bt_gravar_crediarioEnter(Sender: TObject);
begin
  frmvenda.Imprime_display('ENTER confirmar o cliente!', clwhite, tiCliente);
end;

procedure TfrmVenda_Crediario.edClienteEnter(Sender: TObject);
begin
  frmvenda.Imprime_display('Informe o CODIGO/CPF ou NOME do Cliente', clwhite, tiCliente);
end;

procedure TfrmVenda_Crediario.edClienteKeyPress(Sender: TObject;
  var Key: Char);
var
  iRet: Integer;
begin
  if Key = #13 then begin
    bcontinua_creadiario := false;
    bcontinua := false;
    try

      if edCliente.Text <> '' then begin
        // GUIO: 0 - Não achou, e o cliente irá digitar os dados
        //       1 - Achou e não abrirá a pesquisa, localizou apenas um cliente
        //       2 - Achou mais de um cliente
        iRet := LocalizarCliente;
        if (iRet = 0) or (iRet = 2) then begin
          Application.CreateForm(TfrmCliente_consulta, frmCliente_consulta);
          try
            frmCliente_consulta.ed_cliente.Text := AnsiUpperCase(edCliente.Text);
            bConsulta_crediario := True;
            if frmCliente_consulta.ShowModal = mrOK then begin
              edCliente.Text := frmModulo.Query.FieldByName('CODIGO').AsString;
              LocalizarCliente;
              bContinua := True;
              edPegoPor.SetFocus;
            end else begin
            // GUIO: Limpa as labels com os dados do cliente
              PreencherDadosCliente(True);
              edCliente.SetFocus;
              bcontinua := false;
            end;
          finally
            FreeAndNil(frmCliente_Consulta);
          end;
        end;

        if iRet = 1 then begin
          edPegoPor.SetFocus;
          bContinua := True;
        end;
      end;
    finally
      bcontinua := false;
    end;
  end else begin
    if Key = #27 then
      close
  end;
end;

procedure TfrmVenda_Crediario.edClientePropertiesEditValueChanged(
  Sender: TObject);
begin
  Exit;
  gpParcela.Enabled := False;
  gpDados.Enabled := False;
  edParcelas.Clear;
  edDiaFixo.Clear;
  edIntervalo.Clear;
  edVencimento.Clear;
  frmVenda.cdsCrediario.Close;
  frmVenda.cdsCrediario.CreateDataSet;
  edPegoPor.Clear;
  PreencherDadosCliente(True);
end;

procedure TfrmVenda_Crediario.edDiaFixoExit(Sender: TObject);
begin
  if edDiaFixo.Value > 0 then
    edIntervalo.Value := 0;
end;

procedure TfrmVenda_Crediario.edDiaFixoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    edIntervalo.SetFocus;

end;

procedure TfrmVenda_Crediario.edIntervaloExit(Sender: TObject);
begin
  if edIntervalo.Value > 0 then
    edDiaFixo.Value := 0;
end;

procedure TfrmVenda_Crediario.edIntervaloKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    edVencimento.SetFocus;

end;

procedure TfrmVenda_Crediario.edParcelasKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    edDiaFixo.SetFocus;
end;

procedure TfrmVenda_Crediario.edParcelasPropertiesChange(Sender: TObject);
begin
  if edParcelas.Value > 0 then begin
    lbValorParcela.Caption := 'R$ ' + FormatFloat('#,##0.00',RoundTo(ValorReceber/edParcelas.Value,-2));
  end else
    lbValorParcela.Caption := 'R$ 0.00';
end;

procedure TfrmVenda_Crediario.edPegoPorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    edParcelas.SetFocus;
end;

procedure TfrmVenda_Crediario.edVencimentoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    AdvMetroButton1.SetFocus;

end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.bt_cancelar_crediarioEnter(Sender: TObject);
begin
  frmvenda.Imprime_display('ENTER cancelar o crediário!', clwhite, tiCliente);
end;

// -------------------------------------------------------------------------- //
function TfrmVenda_Crediario.LocalizarCliente: Integer;
begin

  with frmModulo.Query do begin
    // buscar pelo nome
    close;
    sql.clear;
    sql.add('select * from cliente');
    sql.add('where upper(nome) LIKE ' + QuotedStr(ansiuppercase(edCliente.text + '%')));
    open;
    if IsEmpty then begin
      // Busca pelo código
      close;
      SQL.clear;
      SQL.add('select * from cliente');
      SQL.add('where codigo = ' + QuotedStr(edCliente.Text));
      open;
      if IsEmpty then begin
        // buscar pelo nome
        close;
        sql.clear;
        SQL.add('select * from cliente');
        SQL.add('where cpf LIKE ' + QuotedStr(edCliente.Text + '%'));
        open;
      end;
    end;

    if not IsEmpty and (RecordCount = 1) then begin
      edParcelas.Value := 1;
      edDiaFixo.Value := DayOfWeek(Date);
      edIntervalo.Value := 0;
      edVencimento.Date := IncMonth(Date);

      gpParcela.Enabled := True;
      gpDados.Enabled := True;
      PreencherDadosCliente;
      Result := 1;  // GUIO: Não Precisa Abrir a Pesquisa
    end else if IsEmpty then
      Result := 0   // GUIO: Não Precisa Abrir a Pesquisa
    else
      Result := 2;   // GUIO: Abrirá a tela de pesquisa
  end;
end;

// -------------------------------------------------------------------------- //
procedure TfrmVenda_Crediario.PreencherDadosCliente(pbVazio: Boolean);
begin
  // GUIO: Preenche as labels com os dados do cliente
  if not pbVazio then begin
    edCliente.text := '';
    lb_cred_codigo.caption := frmmodulo.query.fieldbyname('codigo').asstring;
    lb_cred_nome.Caption := frmmodulo.query.fieldbyname('nome').asstring;
    lb_cred_endereco.caption := frmModulo.query.fieldbyname('endereco').asstring;
    lb_cred_bairro.Caption := frmModulo.query.fieldbyname('bairro').asstring;
    lb_cred_cidade.caption := frmModulo.query.fieldbyname('cidade').asstring;
    lb_cred_uf.caption := frmModulo.query.fieldbyname('uf').asstring;
    lb_cred_cep.caption := frmModulo.query.fieldbyname('cep').AsString;
    lb_cred_cpf.caption := 'CPF/CNPJ: ' + frmModulo.query.fieldbyname('cpf').asstring;
    lb_cred_situacao.caption := frmmodulo.query.fieldbyname('SITUACAO').asstring;
    lb_cred_limite.caption := formatfloat('###,###,##0.00', frmmodulo.query.fieldbyname('limite').asfloat);
    lb_cred_utilizado.caption := formatfloat('###,###,##0.00', frmmodulo.query.fieldbyname('utilizado').asfloat);

    lb_cred_disponivel.caption := formatfloat('###,###,##0.00', frmmodulo.query.fieldbyname('disponivel').asfloat);
    rdisponivel := frmmodulo.query.fieldbyname('disponivel').asfloat;
    edObs.Lines.Clear;
    edObs.Lines.Add(frmModulo.query.fieldbyname('obs').AsString);


    if frmmodulo.query.FieldByName('disponivel').asfloat <= 0 then
      lb_cred_disponivel.Font.color := clred
    else
      lb_cred_disponivel.font.color := cllime;

    if lb_cred_situacao.Caption = '1' then
      lb_cred_situacao.Font.color := clblue
    else if lb_cred_situacao.Caption = '2' then
      lb_cred_situacao.Font.color := clblack
    else if lb_cred_situacao.Caption = '3' then
      lb_cred_situacao.Font.color := clred;

    lb_cred_atualizado.Caption := frmmodulo.query.fieldbyname('atualizado').asstring;
    with frmModulo do begin
      qrPegoPor.Close;
      qrPegoPor.ParamByName('cli').AsString := Zerar(frmmodulo.query.fieldbyname('codigo').asstring, 6);
      qrPegoPor.Open;
      edPegoPor.Properties.Items.Clear;
      qrPegoPor.First;
      edPegoPor.Properties.Items.Add('O MESMO');
      while not qrPegoPor.Eof do begin
        edPegoPor.Properties.Items.Add(qrPegoPorNOME.AsString);
        qrPegoPor.Next;
      end;
      edPegoPor.ItemIndex := 0;
    end;
  end
  else begin
    edPegoPor.Properties.Items.Clear;
    edObs.Lines.Clear;
    edCliente.text := '';
    lb_cred_codigo.caption := '';
    lb_cred_nome.Caption := '';
    lb_cred_endereco.caption := '';
    lb_cred_bairro.Caption := '';
    lb_cred_cidade.caption := '';
    lb_cred_uf.caption := '';
    lb_cred_cep.caption := '';
    lb_cred_cpf.caption := '';
    lb_cred_situacao.caption := '';
    lb_cred_limite.caption := 'R$ 0,00';
    lb_cred_utilizado.caption := 'R$ 0,00';
    lb_cred_disponivel.caption := 'R$ 0,00';

    rdisponivel := 0.00;
    lb_cred_disponivel.Font.color := clWhite;
    lb_cred_disponivel.font.color := clWhite;
    lb_cred_atualizado.Caption := '';
  end
end;

end.

