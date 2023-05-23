unit pre_venda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumnClasses, NxColumns, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, AdvOfficeImage, StdCtrls, RzLabel, jpeg, ExtCtrls, Menus,
  AdvMenus, DB, DBAccess, pngimage, dxGDIPlusClasses, Uni, MemDS;

type
  TfrmPre_Venda = class(TForm)
    pnFundo: TPanel;
    grid_venda: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxDateColumn1: TNxDateColumn;
    NxTextColumn2: TNxTextColumn;
    NxNumberColumn1: TNxNumberColumn;
    NxTextColumn3: TNxTextColumn;
    NxNumberColumn2: TNxNumberColumn;
    NxNumberColumn4: TNxNumberColumn;
    NxNumberColumn3: TNxNumberColumn;
    NxNumberColumn5: TNxNumberColumn;
    NxCheckBoxColumn1: TNxCheckBoxColumn;
    pop_principal: TAdvPopupMenu;
    Finalizar1: TMenuItem;
    Cancelar1: TMenuItem;
    N2: TMenuItem;
    Retornar1: TMenuItem;
   // qrPre_Venda: TIBCQuery;
    tm_Atualiza: TTimer;
    Panel1: TPanel;
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    qrPre_Venda: TUniQuery;
    lb_cliente: TRzLabel;
    qrCabecalho: TUniQuery;
    qrCabecalhoCODIGO: TStringField;
    qrCabecalhoCODCAIXA: TStringField;
    qrCabecalhoCODVENDEDOR: TStringField;
    qrCabecalhoDATA: TDateField;
    qrCabecalhoCODCLIENTE: TStringField;
    qrCabecalhoOBS: TStringField;
    qrCabecalhoMEIO_DINHEIRO: TFloatField;
    qrCabecalhoMEIO_CHEQUEAV: TFloatField;
    qrCabecalhoMEIO_CHEQUEAP: TFloatField;
    qrCabecalhoMEIO_CARTAOCRED: TFloatField;
    qrCabecalhoMEIO_CARTAODEB: TFloatField;
    qrCabecalhoMEIO_CREDIARIO: TFloatField;
    qrCabecalhoSUBTOTAL: TFloatField;
    qrCabecalhoDESCONTO: TFloatField;
    qrCabecalhoACRESCIMO: TFloatField;
    qrCabecalhoTOTAL: TFloatField;
    qrCabecalhoCUPOM_FISCAL: TIntegerField;
    qrCabecalhoNUMERO_CUPOM_FISCAL: TStringField;
    qrCabecalhoRETIRADO: TStringField;
    qrCabecalhoTIPO: TIntegerField;
    qrCabecalhoMEIO_CONVENIO: TFloatField;
    qrCabecalhoDAV: TStringField;
    qrCabecalhoSITUACAO: TIntegerField;
    qrCabecalhoDAV_ANTERIOR: TStringField;
    qrCabecalhoATACADO_VAREJO: TSmallintField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Retornar1Click(Sender: TObject);
    procedure tm_AtualizaTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Finalizar1Click(Sender: TObject);
    procedure Cancelar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPre_Venda: TfrmPre_Venda;
  bConcilia: boolean;
  scod_conciliada: string;

implementation

uses
  modulo, funcoes, principal, msg_Operador, venda, senha_supervisor;

{$R *.dfm}

procedure TfrmPre_Venda.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := cafree;
end;

procedure TfrmPre_Venda.Retornar1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmPre_Venda.tm_AtualizaTimer(Sender: TObject);
var
  i: integer;
  x: integer;
  q: integer;
  b: boolean;
  Nota:string;
begin
  tm_Atualiza.Enabled := false;
  if grid_venda.RowCount > 0 then
    if grid_venda.SelectedRow >= 0 then
      Nota := grid_venda.Cell[0, i].Asstring;

  grid_venda.ClearRows;
  qrPre_Venda.CLOSE;
  qrPre_Venda.SQL.CLEAR;
  qrPre_Venda.SQL.ADD('select');
  qrPre_Venda.SQL.ADD('  c000074.*,');
  qrPre_Venda.SQL.ADD('  c000007.Nome Cliente,');
  qrPre_Venda.SQL.ADD('  c000008.Nome Vendedor');
  qrPre_Venda.SQL.ADD('from');
  qrPre_Venda.SQL.ADD('  c000074, c000007, c000008');
  qrPre_Venda.SQL.ADD('where');
  qrPre_Venda.SQL.ADD('  c000074.codcliente = c000007.codigo and');
  qrPre_Venda.SQL.ADD('  c000074.codvendedor = c000008.codigo and');
  qrPre_Venda.SQL.ADD('  c000074.tipo <> 9 and');
  qrPre_Venda.sql.add('  c000074.situacao = 1');
  qrPre_Venda.sql.add('ORDER BY c000074.CODIGO');
  qrPre_Venda.OPEN;

  qrpre_venda.First;
  while not qrpre_venda.Eof do begin
    i := grid_venda.AddRow(1);
    grid_venda.Cell[0, i].Asstring := qrPre_Venda.fieldbyname('codigo').asstring;
    grid_venda.Cell[1, i].AsDateTime := qrpre_venda.fieldbyname('data').asdatetime;
    grid_venda.Cell[2, i].Asstring := qrPre_Venda.fieldbyname('cliente').asstring;
    grid_venda.Cell[3, i].Asfloat := qrPre_Venda.fieldbyname('total').asfloat;
    grid_venda.Cell[4, i].Asstring := qrPre_Venda.fieldbyname('vendedor').asstring;
    grid_venda.Cell[5, i].Asinteger := qrPre_Venda.fieldbyname('codcliente').asinteger;
    grid_venda.Cell[6, i].Asinteger := qrPre_Venda.fieldbyname('codvendedor').asinteger;
    grid_venda.Cell[7, i].Asfloat := qrPre_Venda.fieldbyname('desconto').asfloat;
    grid_venda.Cell[8, i].Asfloat := qrPre_Venda.fieldbyname('acrescimo').asfloat;
    qrpre_venda.Next;
  end;
    for q := 0 to grid_venda.rowcount - 1 do
    begin
      if grid_venda.cell[0,q].asstring = Nota then
      begin
        grid_venda.Row[q].Selected := True;
        break;
      end;
    end;
//  grid_venda.SelectedRow := ;
  grid_venda.SetFocus;
  tm_Atualiza.Enabled := true;
end;

procedure TfrmPre_Venda.FormShow(Sender: TObject);
begin
  bConcilia := false;
  tm_AtualizaTimer(frmPre_Venda);
end;

procedure TfrmPre_Venda.Finalizar1Click(Sender: TObject);
var
  i: integer;
begin
  try
    tm_Atualiza.Enabled := False;
    if grid_venda.RowCount = 0 then
      exit;
    if grid_venda.SelectedRow < 0 then begin
      application.MessageBox('Favor selecionar uma pré-venda!', 'Atenção', MB_OK + MB_ICONWARNING);
      exit;
    end;
    i := grid_venda.SelectedRow;

    qrCabecalho.Close;
    qrCabecalho.Params.Items[0].Value := grid_venda.cell[0, i].asstring;
    qrCabecalho.Open;

    LancaRecebimento.ValorDesconto := qrCabecalhoDESCONTO.AsFloat;
    LancaRecebimento.ValorAcressimo := qrCabecalhoACRESCIMO.AsFloat;

    frmmodulo.query_servidor.Close;
    frmmodulo.query_servidor.SQL.Clear;
    frmmodulo.query_servidor.SQL.Add('select orc.*, prod.produto, prod.codbarra, prod.cst, prod.comissao, prod.aliquota');
    frmmodulo.query_servidor.sql.add('from c000075 orc,');
    frmmodulo.query_servidor.sql.add('c000025 prod where orc.codproduto = prod.codigo and numeronota = ''' + grid_venda.cell[0, i].asstring + '''');
    frmmodulo.query_servidor.SQL.Add('and coalesce(orc.situacao,1) = 1 and data = :vdat');
    frmmodulo.query_servidor.params.ParamByName('vdat').AsDateTime := grid_venda.Cell[1, i].AsDateTime;
    frmmodulo.query_servidor.Open;
    if frmmodulo.query_servidor.RecordCount = 0 then begin
      if Application.messagebox('Não itens lançados para esta pré-venda! Deseja fazer o cancelamento?', 'Atenção!', mb_yesno + MB_ICONEXCLAMATION) = idno then
        exit
      else
        Cancelar1Click(frmpre_venda);

    end
    else begin
      if not bConcilia then
        if application.messagebox('Confirma a impressão do cupom desta pré-venda?', 'Atenção', mb_yesno + MB_ICONQUESTION) = idno then
          exit;

      bLanca_pre_venda := true;
      iPre_Venda_cliente := grid_venda.cell[5, i].asinteger;
      iPre_Venda_vendedor := grid_venda.cell[6, i].asinteger;
      iPre_venda_codigo := grid_venda.Cell[0, i].asinteger;
      sPre_Venda_Numero := 'PV' + ZERAR(grid_venda.Cell[0, i].asSTRING, 10);
      close;
    end;
  finally
    tm_Atualiza.Enabled := True;
  end;
end;

procedure TfrmPre_Venda.Cancelar1Click(Sender: TObject);
var
  i: integer;
begin
  try
    tm_Atualiza.Enabled := False;
    if grid_venda.RowCount = 0 then
      exit;
    if grid_venda.SelectedRow < 0 then begin
      application.MessageBox('Favor selecionar uma pré-venda!', 'Atenção', MB_OK + MB_ICONWARNING);
      exit;
    end;

    frmsenha_supervisor := tfrmsenha_supervisor.create(self);
    frmsenha_supervisor.showmodal;
    frmvenda.Imprime_display_anterior;
    if not bSupervisor_autenticado then
      exit;
    if application.messagebox('Confirma o cancelamento desta pré-venda?', 'Atenção', mb_yesno + MB_ICONWARNING) = idno then begin
      exit;
    end;

    i := grid_venda.SelectedRow;
    frmmodulo.query_servidor.Close;
    frmmodulo.query_servidor.SQL.Clear;
    frmmodulo.query_servidor.SQL.Add('delete from c000074 where codigo = ' + QuotedStr(grid_venda.cell[0, i].asstring));
    frmmodulo.query_servidor.Execute;

    frmmodulo.query_servidor.Close;
    frmmodulo.query_servidor.SQL.Clear;
    frmmodulo.query_servidor.SQL.Add('delete from c000075 where numeronota = ' + QuotedStr(grid_venda.cell[0, i].asstring));
    frmmodulo.query_servidor.Execute;
    frmmodulo.Conexao_Servidor.Commit;
    tm_AtualizaTimer(tm_Atualiza);
    Application.MessageBox('Pré-Venda Cancelada com sucesso!', 'Atenção!', MB_ICONINFORMATION);
  finally
    tm_Atualiza.Enabled := True;
  end;
end;

end.

