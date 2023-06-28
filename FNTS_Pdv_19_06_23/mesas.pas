unit mesas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumnClasses, NxColumns, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, AdvOfficeImage, StdCtrls,
  RzLabel, jpeg, ExtCtrls, Menus, AdvMenus, DB, DBAccess, pngimage, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel, Uni,
  MemDS, frxClass, frxDBSet, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, cxDBData, cxTextEdit, cxCheckBox,
  cxGridLevel, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGridCustomView, cxGrid, cxClasses, cxButtons, Datasnap.DBClient;

type
  TfrmMesas = class(TForm)
    lb_cliente: TRzLabel;
    pop_principal: TAdvPopupMenu;
    Retornar1: TMenuItem;
    Finalizar1: TMenuItem;
    N2: TMenuItem;
    tm_Atualiza: TTimer;
    Imprimir1: TMenuItem;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label17: TLabel;
    img_cliente: TAdvOfficeImage;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    qrMesa: TUniQuery;
    qrtmp: TUniQuery;
    frFilial: TfrxDBDataset;
    frProdutos: TfrxDBDataset;
    frComanda: TfrxReport;
    qrProdutos: TUniQuery;
    qrProdutosCODIGO: TStringField;
    qrProdutosDATA: TDateField;
    qrProdutosHORA: TStringField;
    qrProdutosPERGUNTOU_COMISSAO: TStringField;
    qrProdutosCLIENTE_AUTORIZOU: TStringField;
    qrProdutosCOD_FUNCIONARIO: TStringField;
    qrProdutosNOME: TStringField;
    qrProdutosCOD_PRODUTO: TStringField;
    qrProdutosPRODUTO: TStringField;
    qrProdutosCOD_LANCAMENTO: TIntegerField;
    qrProdutosQTDE: TFloatField;
    qrProdutosUNITARIO: TFloatField;
    qrProdutosTOTAL: TFloatField;
    qrProdutosIMPRESSO: TStringField;
    qrProdutosTRANSF_MESA: TIntegerField;
    qrProdutosCOMPLEMENTO: TStringField;
    qrProdutosEXTRA: TBlobField;
    dsMesa: TDataSource;
    qrMesaMESA: TIntegerField;
    qrMesaCOMANDA: TIntegerField;
    qrMesaPERGUNTOU_COMISSAO: TStringField;
    qrMesaCLIENTE_AUTORIZOU: TStringField;
    qrMesaSUB_TOTAL: TFloatField;
    qrMesaCOUVER: TFloatField;
    qrMesaCOMISSAO: TFloatField;
    qrMesaTOTAL: TFloatField;
    cxStyleRepository1: TcxStyleRepository;
    Content: TcxStyle;
    Header: TcxStyle;
    cxStyle1: TcxStyle;
    cxStyle2: TcxStyle;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1DBTableView1MESA: TcxGridDBColumn;
    cxGrid1DBTableView1COMANDA: TcxGridDBColumn;
    cxGrid1DBTableView1PERGUNTOU_COMISSAO: TcxGridDBColumn;
    cxGrid1DBTableView1CLIENTE_AUTORIZOU: TcxGridDBColumn;
    cxGrid1DBTableView1SUB_TOTAL: TcxGridDBColumn;
    cxGrid1DBTableView1COUVER: TcxGridDBColumn;
    cxGrid1DBTableView1COMISSAO: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxGrid1Level1: TcxGridLevel;
    qrMesaCODIGO: TIntegerField;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    cdsMesas: TClientDataSet;
    cdsMesasMESA: TIntegerField;
    cdsMesasCOMANDA: TIntegerField;
    cdsMesasPERGUNTOU_COMISSAO: TStringField;
    cdsMesasCLIENTE_AUTORIZOU: TStringField;
    cdsMesasSUB_TOTAL: TFloatField;
    cdsMesasCOUVER: TFloatField;
    cdsMesasCOMISSAO: TFloatField;
    cdsMesasTOTAL: TFloatField;
    cdsMesasCODIGO: TIntegerField;
    cdsMesasSEL: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Retornar1Click(Sender: TObject);
    procedure tm_AtualizaTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Finalizar1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure frComandaBeforePrint(Sender: TfrxReportComponent);
    procedure cxButton3Click(Sender: TObject);
    procedure cxGrid1DBTableView1COMANDAGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: string);
    procedure cxGrid1DBTableView1MESAGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: string);
    procedure cxGrid1DBTableView1COUVERGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: string);
    procedure cxGrid1DBTableView1COMISSAOGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: string);
    procedure qrMesaBeforeOpen(DataSet: TDataSet);
  private
    { Private declarations }
    mTotal_Comissao,
    mTotal_Venda,
    mTotal_Couvert:Double;
  public
    { Public declarations }
  end;

var
  frmMesas: TfrmMesas;

implementation

uses modulo, funcoes, principal, msg_Operador, venda, senha_supervisor,
  UFuncoes, FechamentoComanda;

{$R *.dfm}

procedure TfrmMesas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmMesas.FormCreate(Sender: TObject);
begin
  Self.DoubleBuffered := True;
  left := 0;
  top := 0;
end;

procedure TfrmMesas.Retornar1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmMesas.tm_AtualizaTimer(Sender: TObject);
var
  Local:Integer;
begin
  try
    if not cdsMesas.Active then
      cdsMesas.CreateDataSet;
    cdsMesas.DisableControls;
    tm_Atualiza.Enabled := false;
    Local := 0;
    if not cdsMesas.IsEmpty then
      Local := cdsMesasCODIGO.AsInteger;
    qrMesa.close;
    qrMesa.open;
    qrMesa.First;
    while not qrMesa.Eof do begin
      if not cdsMesas.Locate('codigo',qrMesaCODIGO.AsInteger,[loCaseInsensitive]) then begin
        cdsMesas.Append;
        cdsMesasMESA.Value := qrMesaMESA.Value;
        cdsMesasCOMANDA.Value := qrMesaCOMANDA.Value;
        cdsMesasPERGUNTOU_COMISSAO.Value := qrMesaPERGUNTOU_COMISSAO.Value;
        cdsMesasCLIENTE_AUTORIZOU.Value := qrMesaCLIENTE_AUTORIZOU.Value;
        cdsMesasSUB_TOTAL.Value := qrMesaSUB_TOTAL.Value;
        cdsMesasCOUVER.Value := qrMesaCOUVER.Value;
        cdsMesasCOMISSAO.Value := qrMesaCOMISSAO.Value;
        cdsMesasTOTAL.Value := qrMesaTOTAL.Value;
        cdsMesasCODIGO.Value := qrMesaCODIGO.Value;
        cdsMesasSEL.AsString := 'N';
        cdsMesas.Post;
      end;
      qrMesa.Next;
    end;
    if Local > 0 then
      cdsMesas.Locate('codigo',Local,[loCaseInsensitive]);
    tm_Atualiza.Enabled := true;
  finally
    cdsMesas.EnableControls;
  end;
end;

procedure TfrmMesas.FormShow(Sender: TObject);
begin
  tm_AtualizaTimer(tm_Atualiza);
end;

procedure TfrmMesas.frComandaBeforePrint(Sender: TfrxReportComponent);
begin
  if Sender is TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'cpTotal' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00',mTotal_Venda);
    if TfrxMemoView(Sender).Name = 'cptaxa' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00',mTotal_Comissao);
    if TfrxMemoView(Sender).Name = 'cpCouvert' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00',mTotal_Couvert);
  end;
end;

procedure TfrmMesas.cxButton3Click(Sender: TObject);
var
  Sel:Boolean;
  sCodigos:string;
begin
  try
    if cdsMesas.State = dsEdit then
      cdsMesas.Post;
    tm_Atualiza.Enabled := False;
    cdsMesas.DisableControls;
    cdsMesas.First;
    Sel := False;
    sCodigos := '';
    while not cdsMesas.Eof do begin
      if cdsMesasSEL.AsString = 'S' then begin
        Sel := True;
        if sCodigos <> '' then
          sCodigos := sCodigos + ', ';
        sCodigos := sCodigos + cdsMesasCODIGO.AsString;
      end;
      cdsMesas.Next;
    end;
    if not Sel then begin
      Application.MessageBox('Não foi feito seleção para imprimir a conferência!','Atenção!',MB_ICONINFORMATION);
      exit;
    end;
    Application.CreateForm(TfrmFechamentoComanda, frmFechamentoComanda);
    frmFechamentoComanda.strCodigos := sCodigos;
    frmFechamentoComanda.ShowModal;
    FreeAndNil(frmFechamentoComanda);
  finally
    tm_Atualiza.Enabled := True;
    tm_AtualizaTimer(tm_Atualiza);
    cdsMesas.EnableControls;
  end;
end;

procedure TfrmMesas.cxGrid1DBTableView1COMANDAGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: string);
begin
  if AText = '0' then
    AText := '';
end;

procedure TfrmMesas.cxGrid1DBTableView1COMISSAOGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: string);
begin
  if AText = '0' then
    AText := '';
end;

procedure TfrmMesas.cxGrid1DBTableView1COUVERGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: string);
begin
  if AText = '0' then
    AText := '';
end;

procedure TfrmMesas.cxGrid1DBTableView1MESAGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: string);
begin
  if AText = '0' then
    AText := '';
end;

procedure TfrmMesas.Finalizar1Click(Sender: TObject);
var
  i : integer;
  pergunta, Autorizou, marcou:Boolean;
begin
  if cdsMesas.RecordCount = 0 then
    exit;
  try
    if cdsMesas.State = dsEdit then
      cdsMesas.Post;
    tm_Atualiza.Enabled := False;
    marcou := False;
    Autorizou := false;
    pergunta := False;
    cdsMesas.DisableControls;
    cdsMesas.First;
    while not cdsMesas.Eof do begin
      if cdsMesasSEL.AsString = 'S' then begin
        marcou := True;
        if frmModulo.qrconfigCOBRA_COMISSAO.AsString = 'P' then begin
          if cdsMesasperguntou_comissao.AsString = 'N' then begin
            pergunta := True;
          end;
        end;
      end;
      cdsMesas.Next;
    end;
    if not marcou then begin
      Application.MessageBox('Selecione pelo menos uma Mesa ou Comanda para receber!','Atenção!',MB_ICONINFORMATION);
      Exit;
    end;
    if pergunta then begin
      if Application.MessageBox('Cliente autorizou incluir taxa do garçon na comanda?','Atenção?',MB_ICONQUESTION+MB_YESNO) =mrYes then
        Autorizou := True;
    end;
    cdsMesas.First;
    MesasComandasReceber := '';
    while not cdsMesas.Eof do begin
      if cdsMesasSEL.AsString = 'S' then begin
        frmmodulo.query_servidor.close;
        frmmodulo.query_servidor.sql.clear;
        frmmodulo.query_servidor.SQL.add('select r000002.*, C000025.produto, C000025.unidade,');
        frmmodulo.query_servidor.sql.add('c000025.cst, c000025.aliquota, c000025.codbarra');
        frmmodulo.query_servidor.sql.add('from r000002, c000025');
        frmmodulo.query_servidor.sql.add('where r000002.cod_produto = c000025.codigo');
        frmmodulo.query_servidor.sql.add('and r000002.cod_mesa = ' + cdsMesasCODIGO.AsString);
        frmmodulo.query_servidor.sql.add('order by r000002.codigo');
        frmmodulo.query_servidor.open;
        if (frmmodulo.query_servidor.RecordCount = 0) then begin
          Application.messagebox('Não tem itens lançados para esta Mesa!','Atenção!',mb_ok+mb_ICONerror);
          tm_Atualiza.Enabled := True;
        end else begin
          if frmModulo.qrconfigCOBRA_COMISSAO.AsString = 'P' then begin
            if cdsMesasperguntou_comissao.AsString = 'N' then begin
              qrtmp.SQL.Clear;
              if Autorizou then
                qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('S') + ' where CODIGO = ' + cdsMesasCODIGO.AsString)
              else
                qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('N') + ' where CODIGO = ' + cdsMesasCODIGO.AsString);
              qrtmp.ExecSQL;
            end;
          end;
          bLanca_mesa := true;
          if MesasComandasReceber <> '' then
            MesasComandasReceber := MesasComandasReceber + ', ';
          MesasComandasReceber := MesasComandasReceber + cdsMesasCODIGO.AsString;
          close;
        end;
      end;
      cdsMesas.Next;
    end;
    frmmodulo.query_servidor.close;
    frmmodulo.query_servidor.sql.clear;
    frmmodulo.query_servidor.SQL.add('select r000002.*, C000025.produto, C000025.unidade,');
    frmmodulo.query_servidor.sql.add('c000025.cst, c000025.aliquota, c000025.codbarra');
    frmmodulo.query_servidor.sql.add('from r000002, c000025');
    frmmodulo.query_servidor.sql.add('where r000002.cod_produto = c000025.codigo');
    frmmodulo.query_servidor.sql.add('and r000002.cod_mesa in (' + MesasComandasReceber+ ')');
    frmmodulo.query_servidor.sql.add('order by r000002.codigo');
    frmmodulo.query_servidor.open;

  finally
    qrMesa.EnableControls;
  end;
end;

procedure TfrmMesas.Imprimir1Click(Sender: TObject);
var
  i : integer;
  sNumero_Dav : string;
  Rsoma : real;
begin
  if qrMesa.RecordCount = 0 then
    exit;
  qrProdutos.Close;
  qrProdutos.Params.Items[0].Value := qrMesaCODIGO.AsString;
  qrProdutos.Open;
  qrProdutos.First;
  mTotal_Venda := 0;
  mTotal_Comissao := 0;
  while not qrProdutos.Eof do begin
    mTotal_Venda := mTotal_Venda + qrProdutosTOTAL.AsFloat;
    qrProdutos.Next;
  end;
  qrProdutos.First;
  with frmModulo do begin
    if qrProdutosCLIENTE_AUTORIZOU.AsString = 'S' then begin
      mTotal_Comissao := (mTotal_Venda*qrconfigCOMISSAO_PERCENTUAL.AsFloat)/100;
    end else begin
      qrtmp.SQL.Clear;
      if qrconfigCOBRA_COMISSAO.AsString = 'O' then begin
        mTotal_Comissao := (mTotal_Venda*qrconfigCOMISSAO_PERCENTUAL.AsFloat)/100;
        qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('S') + ' where CODIGO = ' + qrMesaCODIGO.AsString)
      end else if qrconfigCOBRA_COMISSAO.AsString = 'P' then begin
        if Application.MessageBox('Cliente autorizou incluir taxa do garçon na comanda?','Atenção?',MB_ICONQUESTION+MB_YESNO) =mrYes then begin
          qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('S') + ' where CODIGO = ' + qrMesaCODIGO.AsString);
          mTotal_Comissao := (mTotal_Venda*qrconfigCOMISSAO_PERCENTUAL.AsFloat)/100
        end else begin
          qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('N') + ' where CODIGO = ' + qrMesaCODIGO.AsString);
          mTotal_Comissao := 0;
        end;
        qrtmp.ExecSQL;
      end;
    end;
    if mTotal_Comissao > 0 then
      mTotal_Venda := mTotal_venda + mTotal_Comissao;
    if qrconfigVALOR_COUVER.AsFloat > 0 then begin
      mTotal_Couvert := qrconfigVALOR_COUVER.AsFloat;
      mTotal_Venda := mTotal_venda + mTotal_Couvert;
    end;
  end;
  frComanda.PrintOptions.Printer := sImpressora_Venda_Simples;
  //frComanda.Print;
  frComanda.ShowReport;
  tm_AtualizaTimer(frmMesas);
end;



procedure TfrmMesas.qrMesaBeforeOpen(DataSet: TDataSet);
begin
  qrMesa.ParamByName('atb').Value := ME_FiltraATB('TB_CONFIGURACOES');
end;

end.
