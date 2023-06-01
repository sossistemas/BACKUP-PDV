unit produto_consulta;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, jpeg, ExtCtrls, StdCtrls, RzLabel, AdvOfficeImage, Mask, RzEdit,
  NxColumnClasses, NxColumns, NxScrollControl, NxCustomGridControl, NxCustomGrid,
  NxGrid, Menus, AdvMenus, DB, DBAccess, pngimage, frxClass, dxGDIPlusClasses,
  Uni, MemDS, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, cxNavigator,
  cxDBData, cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, dxSkinsCore, dxSkinsDefaultPainters, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog;

type
  TfrmProduto_consulta = class(TForm)
    pop_principal: TAdvPopupMenu;
    Estoque1: TMenuItem;
    query: TUniQuery;
    Panel11: TPanel;
    Panel12: TPanel;
    pnFundo: TPanel;
    img_cliente: TAdvOfficeImage;
    lb_status: TRzLabel;
    lb_cliente: TRzLabel;
    RzLabel1: TRzLabel;
    ed_barra: TRzEdit;
    DataSource1: TDataSource;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    queryCODIGO: TIntegerField;
    queryCOD_BARRA: TStringField;
    queryNOME: TStringField;
    queryUNIDADE: TStringField;
    queryPRECO_VENDA: TFloatField;
    queryNCM: TStringField;
    queryCSOSN: TStringField;
    queryCFOP: TStringField;
    queryESTOQUE: TFloatField;
    cxGrid1DBTableView1CODIGO: TcxGridDBColumn;
    cxGrid1DBTableView1COD_BARRA: TcxGridDBColumn;
    cxGrid1DBTableView1NOME: TcxGridDBColumn;
    cxGrid1DBTableView1UNIDADE: TcxGridDBColumn;
    colPreco: TcxGridDBColumn;
    cxGrid1DBTableView1NCM: TcxGridDBColumn;
    cxGrid1DBTableView1CSOSN: TcxGridDBColumn;
    cxGrid1DBTableView1CFOP: TcxGridDBColumn;
    colEstoque: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    ln01: TcxStyle;
    ln02: TcxStyle;
    Cab: TcxStyle;
    queryIDONLINE: TIntegerField;
    queryEMDESTAQUE: TStringField;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    procedure ed_barraEnter(Sender: TObject);
    procedure ed_barraKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure LocalizarProduto(referencia: string);
    procedure cxGrid1DBTableView1KeyPress(Sender: TObject; var Key: Char);
    procedure cxGrid1DBTableView1CellDblClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure Estoque1Click(Sender: TObject);
    procedure cxGrid1DBTableView1CustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProduto_consulta: TfrmProduto_consulta;

implementation

uses
  venda, funcoes, modulo, Math, principal, consulta_estoque, ConsultaEstoqueFilial;

{$R *.dfm}

// -------------------------------------------------------------------------- //
procedure TfrmProduto_consulta.cxGrid1DBTableView1CellDblClick(
  Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
  AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  if query.Active then begin
    if query.RecordCount > 0 then begin
      if iPesquisa_produto > 0 then begin
        iProd_codigo := queryCODIGO.asinteger;
        close;
      end;
    end;
  end;
end;

procedure TfrmProduto_consulta.cxGrid1DBTableView1CustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
begin
  if(AViewInfo.GridRecord.Values[cxGrid1DBTableView1Column1.Index] = 'S') then begin
    ACanvas.Brush.Color := clGreen;
    ACanvas.Font.Color := clWhite;
  end;
end;

procedure TfrmProduto_consulta.cxGrid1DBTableView1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
    Key := #0;
    if query.Active then begin
      if query.RecordCount > 0 then begin
        if iPesquisa_produto > 0 then begin
          iProd_codigo := queryCODIGO.asinteger;
          close;
        end;
      end;
    end;
  end else begin
    if Key = #27 then begin
      ed_barra.setfocus;
    end;
  end;
end;

procedure TfrmProduto_consulta.ed_barraEnter(Sender: TObject);
begin
end;

// -------------------------------------------------------------------------- //
procedure TfrmProduto_consulta.ed_barraKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
    Key := #0;

    if ed_barra.Text <> '' then begin
      iPesquisa_produto := 1;
      LocalizarProduto(ANSIUPPERCASE(ed_barra.text));
    end;
  end
  else begin
    if Key = #27 then
      close
  end;
end;

procedure TfrmProduto_consulta.Estoque1Click(Sender: TObject);
begin
  if frmModulo.qrconfigMULTILOJA.AsString <> 'S' then
    Exit;
  if (not query.Active) or (query.RecordCount = 0) then begin
    Application.MessageBox('Selecione o produto para consultar o estoque.','Atenção!',MB_ICONINFORMATION);
    exit;
  end;
  EstoqueFilial(queryIDONLINE.AsInteger,
    frmModulo.qrconfigDATABASE.AsString,
    frmModulo.qrconfigSERVER.AsString,
    frmModulo.qrconfigUSERNAME.AsString,
    frmModulo.qrconfigPASSWORD.AsString,
    frmModulo.qrconfigPORT.AsInteger);
end;

// -------------------------------------------------------------------------- //
procedure TfrmProduto_consulta.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  iPesquisa_produto := 0;
  Action := cafree;
end;

// -------------------------------------------------------------------------- //
procedure TfrmProduto_consulta.FormShow(Sender: TObject);
begin
  lb_status.Visible := frmModulo.qrconfigMULTILOJA.AsString = 'S';
  iProd_codigo := 0;
  if ed_barra.Text <> '' then begin
    LocalizarProduto(ed_barra.Text);
  end
  else
    ed_barra.SetFocus;
  colPreco.Visible := bExibePrecoConsulta;
  colEstoque.Visible := bExibeEstoqueConsulta;
end;

procedure TfrmProduto_consulta.LocalizarProduto(referencia: string);
var
  bBusca_codigo: boolean; // variavel de para ativar ou nao a busca pelo campo codigo
  procedure ReestruturarTSQL;
  begin
    Query.close;
    Query.sql.clear;
    Query.sql.add('SELECT CODIGO');
    Query.sql.add('      ,COD_BARRA');
    Query.sql.add('      ,NOME');
    Query.sql.add('      ,UNIDADE');
    Query.sql.add('      ,PRECO_VENDA');
    Query.sql.add('      ,NCM');
    Query.sql.add('      ,CSOSN');
    Query.sql.add('      ,CFOP');
    Query.sql.add('      ,CASE WHEN CODPRODUTO <> ''0'' THEN TRUNC(ESTOQUE / QTD_VOLUME)');
    Query.sql.add('                                   ELSE ESTOQUE');
    Query.sql.add('                                   END ESTOQUE');
    Query.sql.add('      ,IDONLINE');
    Query.sql.add('      ,EMDESTAQUE');
    Query.sql.add('FROM ESTOQUE');
  end;
begin
  bBusca_codigo := False;

  // verificar se o campo referencia eh um numero inteiro valido
  try
    strtoint(referencia);
    bBusca_codigo := true;
  except
    on EConvertError do
      bBusca_codigo := false;
  end;

  ReestruturarTSQL;

  if bBusca_codigo then
  begin
    // nao achou pelo codigo de barras, procurar pelo codigo
    Query.sql.add('where codigo = ' + referencia);
    query.sql.add('and SITUACAO = 0');
    Query.open;
  end else
  begin
    // comando SQL para filtrar o produto pelo codigo de barras
    query.sql.add('where cod_barra = ''' + referencia + '''');
    query.sql.add('and SITUACAO = 0');
    query.Open;

    if query.IsEmpty then
    begin
      // GUIO: Tenta localizar a referencia do produto
      ReestruturarTSQL;
      Query.sql.add('where upper(REFERENCIA) LIKE ''' + ANSIUPPERCASE(referencia) + '%''');
      query.sql.add('and SITUACAO = 0');
      Query.sql.add('order by REFERENCIA, COR, TAMANHO');
      Query.open;
    end;

    if query.IsEmpty then
    begin
      // nao achou pelo codigo de barras, procurar pelo nome
      ReestruturarTSQL;
      Query.sql.add('where upper(nome) like ' + QuotedStr('%' + ANSIUPPERCASE(referencia) + '%'));
      query.sql.add('and SITUACAO = 0');
      Query.sql.add('order by nome');
      Query.open;
    end;
  end;


  if Query.IsEmpty then
  begin
    Application.ProcessMessages;
    Sleep(1500);
    ed_barra.setfocus;

    ED_barra.TEXT := '';

    ed_barra.SelectAll;
  end else
    cxGrid1.SetFocus;
end;

end.

