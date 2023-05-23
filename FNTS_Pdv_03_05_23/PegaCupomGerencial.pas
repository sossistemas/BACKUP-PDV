unit PegaCupomGerencial;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, Vcl.Menus, Data.DB, MemDS,
  DBAccess, Uni, Vcl.StdCtrls, cxButtons, cxTextEdit, AdvMetroButton,
  Vcl.ExtCtrls, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView, cxGrid,
  cxMaskEdit, cxSpinEdit, cxDBEdit, cxDropDownEdit, cxCalc, cxCheckBox, acbrutil, dxSkinsCore, dxSkinsDefaultPainters, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog, cxLabel, SOSSolution.Controls;

type
  TfrmPegaCupomGerencial = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    qrItens: TUniQuery;
    qrItensDATA: TDateField;
    qrItensTIPO: TStringField;
    qrItensCODPRODUTO: TStringField;
    qrItensPRODUTO: TStringField;
    qrItensUNITARIO: TFloatField;
    qrItensQTDE: TFloatField;
    qrItensDESCONTO: TFloatField;
    qrItensTOTAL: TFloatField;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    DataSource1: TDataSource;
    cxGrid1DBTableView1DATA: TcxGridDBColumn;
    cxGrid1DBTableView1TIPO: TcxGridDBColumn;
    cxGrid1DBTableView1CODPRODUTO: TcxGridDBColumn;
    cxGrid1DBTableView1PRODUTO: TcxGridDBColumn;
    cxGrid1DBTableView1UNIDADE: TcxGridDBColumn;
    cxGrid1DBTableView1UNITARIO: TcxGridDBColumn;
    cxGrid1DBTableView1QTDE: TcxGridDBColumn;
    cxGrid1DBTableView1DESCONTO: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    Panel5: TPanel;
    Label3: TLabel;
    edCodigo: TcxDBTextEdit;
    cxDBTextEdit2: TcxDBTextEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    edQtde: TcxSpinEdit;
    edValor: TcxCalcEdit;
    edTotal: TcxCalcEdit;
    qrItensCODCLIENTE: TStringField;
    qrItensCUPOM_ITEM: TStringField;
    qrItensSITUACAO: TIntegerField;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    qrItensDEVOLVIDO: TStringField;
    qrItensUNIDADE: TStringField;
    pnlSelecaoCupom: TPanel;
    Panel7: TPanel;
    Label7: TLabel;
    cxGrid2: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxgdbcDATA: TcxGridDBColumn;
    cxgdbcCUPOM_NUMERO: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    cxgdbcQTD_DEVOLVIDO: TcxGridDBColumn;
    Panel8: TPanel;
    cxButton4: TcxButton;
    Label8: TLabel;
    uqrySelecaoCupom: TUniQuery;
    uqrySelecaoCupomDATA: TDateField;
    uqrySelecaoCupomCUPOM_NUMERO: TStringField;
    uqrySelecaoCupomQTD_DEVOLVIDO: TIntegerField;
    dsSelecaoCupom: TDataSource;
    Panel4: TPanel;
    Label2: TLabel;
    edCupom: TcxTextEdit;
    cxButton1: TcxButton;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RoundShape1: TRoundShape;
    Panel6: TPanel;
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxGrid1DBTableView1FocusedRecordChanged(
      Sender: TcxCustomGridTableView; APrevFocusedRecord,
      AFocusedRecord: TcxCustomGridRecord;
      ANewItemRecordFocusingChanged: Boolean);
    procedure edCupomKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure RadioButtonClick(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure cxGridDBTableView1CellClick(Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
      AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
    procedure cxGridDBTableView1CellDblClick(Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
      AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
  private
    { Private declarations }
  public
    { Public declarations }
    Confirma:Boolean;
  end;

var
  frmPegaCupomGerencial: TfrmPegaCupomGerencial;


implementation

uses
  modulo, venda, cliente_consulta, principal, Math;

var
  CodigoBarraProduto: String;

{$R *.dfm}

procedure TfrmPegaCupomGerencial.cxButton1Click(Sender: TObject);
const
  cInfNumCup = 'Informe o Número do Cupom!';
  cInfCodBar = 'Informe o Código de Barra do Produto';
  cCodBarInv = 'Código de Barras inválido!';
  cProNaoEnc = 'Nenhum produto compatível com o Código de Barras informado!';
  cSemVenPro = '';
  function iif(ACondition: Boolean; ATrueValue, AfalseValue: String): PWideChar;
  begin
    if ACondition then
      Result := PWideChar(ATrueValue)
    else
      Result := PWideChar(AFalseValue);
  end;
begin
  if edCupom.Text = '' then begin
    Application.MessageBox(iif(RadioButton1.Checked, cInfNumCup, cInfCodBar),'Atenção!',MB_ICONINFORMATION);
    edCupom.SetFocus;
    exit;
  end;

  if RadioButton2.Checked  then
  begin
    if Trim(edCupom.Text).Length <> 13 then
    begin
      Application.MessageBox(cCodBarInv,'Atenção!',MB_ICONINFORMATION);
      edCupom.SetFocus;
    end
    else
    with uqrySelecaoCupom do
    begin
       cxButton4.Enabled := CodigoBarraProduto <> edCupom.Text;
       CodigoBarraProduto := edCupom.Text;
       ///
       with TUniQuery.Create(Self) do
       try
         Connection := frmModulo.Conexao_Servidor;
         Close;
         SQL.Clear;
         SQL.Add('SELECT CODBARRA FROM  C000025 WHERE CODBARRA = ' + QuotedStr(Trim(CodigoBarraProduto)));
         Open;
         if RecordCount = 0 then
         begin
           Application.MessageBox(cProNaoEnc,'Atenção!',MB_ICONINFORMATION);
           edCupom.SetFocus;
           Free;
           Exit;
         end;
         Free;
       except
         Free;
       end;
       ///
       Close;
       ParamByName('DIAS').Value := IfThen(PrazoDevolucao = 0, 999, PrazoDevolucao);
       ParamByName('CODIGO').Value := Trim(CodigoBarraProduto);
       Open;
       if RecordCount = 0 then
       begin
         Application.MessageBox(cProNaoEnc,'Atenção!',MB_ICONINFORMATION);
         edCupom.SetFocus;
         Exit;
       end;
       ///
       pnlSelecaoCupom.Visible := True;
       edCupom.Text := '';
       cxButton2.Enabled := False;
    end;
  end
  else
  with frmVenda do begin
    if not cdsDevolucao.IsEmpty then begin
      if cdsDevolucaocodcupom.AsString <> edCupom.Text then begin
        Application.MessageBox('Não é permitido realização devolução de varios cupons!','Atenção!',MB_ICONINFORMATION);
        edCupom.SetFocus;
        Exit;
      end;
    end;
    try
      qrItens.DisableControls;
      qrItens.Close;
      qrItens.Params.Items[0].Value := edCupom.Text;
      qrItens.Open;
      cdsDevolucao.First;
      while not cdsDevolucao.eof do begin
        if (qrItens.Locate('CUPOM_ITEM',IntToStrZero(cdsDevolucaoitem.AsInteger,3),[loCaseInsensitive])) and (edCupom.Text = cdsDevolucaocodcupom.AsString) then begin
          qrItens.Edit;
          qrItensSITUACAO.AsInteger := 2;
          qrItens.Post;
        end;
        cdsDevolucao.Next;
      end;
    finally
      qrItens.First;
      qrItens.EnableControls;
    end;
  end;
end;

procedure TfrmPegaCupomGerencial.cxButton2Click(Sender: TObject);
begin
  if RadioButton1.Checked then
  begin
    if not(qrItens.Active) or (qrItens.IsEmpty) then begin
      Application.MessageBox('Selecione o produto a ser devolvido!','Atenção!',MB_ICONINFORMATION);
      edCupom.SetFocus;
      Exit;
    end;
    if qrItensDEVOLVIDO.AsString = 'S' then begin
      Application.MessageBox('Impossivel selecionar um produto que já foi devolvido!','Atenção!',MB_ICONINFORMATION);
      edCupom.SetFocus;
      Exit;
    end;
    with frmVenda do begin
      cdsDevolucao.Append;
      cdsDevolucaocodcupom.AsString := edCupom.Text;
      cdsDevolucaoitem.AsInteger := qrItensCUPOM_ITEM.AsInteger;
      cdsDevolucaovalor.AsFloat := qrItensTOTAL.AsFloat - qrItensDESCONTO.AsFloat;
      cdsDevolucao.Post;
    end;
    CodCliGeraCredito := qrItensCODCLIENTE.AsInteger;
    Confirma:=True;
    Close;
  end
  else
  if pnlSelecaoCupom.Visible then
  begin
    pnlSelecaoCupom.Visible := False;
    RadioButton1.Checked := True;
    if cxButton4.Enabled then
    begin
      edCupom.Text := uqrySelecaoCupomCUPOM_NUMERO.Value;
      cxButton1.Click;
    end
    else
      cxButton1.SetFocus;
  end;
end;

procedure TfrmPegaCupomGerencial.cxButton3Click(Sender: TObject);
begin
  if RadioButton1.Checked or not pnlSelecaoCupom.Visible then
  begin
    Confirma := False;
    Close;
  end
  else
  begin
    pnlSelecaoCupom.Visible := False;
    edCupom.SetFocus;
  end;
end;

procedure TfrmPegaCupomGerencial.cxButton4Click(Sender: TObject);
begin
  with uqrySelecaoCupom do
  begin
    First;
    while not Eof and (uqrySelecaoCupomQTD_DEVOLVIDO.Value <> 0) do
      Next;
    if EOF then
    begin
      Application.MessageBox('O sistema não encontrou nenhum cupom sem devolução para este produto!'#13'A devolução não poderá ser registrada.'
                            ,'Atenção!'
                            ,MB_ICONINFORMATION);
      cxButton4.Enabled := False;
    end;
    cxButton2.Click;
  end;
end;

procedure TfrmPegaCupomGerencial.cxGrid1DBTableView1FocusedRecordChanged(
  Sender: TcxCustomGridTableView; APrevFocusedRecord,
  AFocusedRecord: TcxCustomGridRecord; ANewItemRecordFocusingChanged: Boolean);
begin
  edQtde.Value := qrItensQTDE.AsFloat;
  edValor.Value := qrItensUNITARIO.AsFloat - qrItensDESCONTO.AsFloat;
  edTotal.Value := edQtde.Value * edValor.Value;
end;

procedure TfrmPegaCupomGerencial.cxGridDBTableView1CellClick(Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  CodigoBarraProduto := '';
  cxButton2.Enabled := True;
end;

procedure TfrmPegaCupomGerencial.cxGridDBTableView1CellDblClick(Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  CodigoBarraProduto := '';
  cxButton2.Enabled := True;
  cxButton2.Click;
end;

procedure TfrmPegaCupomGerencial.edCupomKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    cxButton1.Click;
end;

procedure TfrmPegaCupomGerencial.FormCreate(Sender: TObject);
begin
  pnlSelecaoCupom.Left := 0;
  Self.Width := 713;
end;

procedure TfrmPegaCupomGerencial.RadioButtonClick(Sender: TObject);
begin
  Label2.Caption := (Sender as TRadioButton).Caption;
end;

end.
