unit CuponsSAT;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, System.ImageList, Vcl.ImgList, MemDS, DBAccess, Uni, Vcl.StdCtrls,
  Vcl.ExtCtrls, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel,
  Vcl.Mask, sMaskEdit, sCustomComboEdit, sToolEdit, cxImageComboBox, Vcl.Menus,
  cxButtons, dxSkinsCore, dxSkinsDefaultPainters, dxDateRanges, cxDataControllerConditionalFormattingRulesManagerDialog,
  System.IOUtils,
  System.Types,
  System.DateUtils;

type
  TfrmCuponsSAT = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    lbCancelamento: TLabel;
    Panel3: TPanel;
    lbEnvio: TLabel;
    qryTmp: TUniQuery;
    qrSAT: TUniQuery;
    qrTemp: TUniQuery;
    ImageList1: TImageList;
    dsSAT: TDataSource;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    qrSATID: TIntegerField;
    qrSATDATA: TDateField;
    qrSATHORA: TStringField;
    qrSATNUMERO: TIntegerField;
    qrSATCODCLIENTE: TStringField;
    qrSATAMBIENTE: TIntegerField;
    qrSATCANCELADO: TStringField;
    qrSATCHAVE: TStringField;
    qrSATXML_ENVIO: TBlobField;
    qrSATXML_CANCELAMENTO: TBlobField;
    qrSATALTERADO: TStringField;
    qrSATEX: TIntegerField;
    qrSATCUPOM: TStringField;
    qrSATTOTAL: TFloatField;
    cxGrid1DBTableView1DATA: TcxGridDBColumn;
    cxGrid1DBTableView1HORA: TcxGridDBColumn;
    cxGrid1DBTableView1NUMERO: TcxGridDBColumn;
    cxGrid1DBTableView1AMBIENTE: TcxGridDBColumn;
    cxGrid1DBTableView1CANCELADO: TcxGridDBColumn;
    cxGrid1DBTableView1CUPOM: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxGrid1DBTableView1CHAVE: TcxGridDBColumn;
    Panel4: TPanel;
    AdvMetroButton1: TAdvMetroButton;
    Label1: TLabel;
    Panel5: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    edt_Numero: TEdit;
    dataini: TsDateEdit;
    datafin: TsDateEdit;
    Memo1: TMemo;
    cxbCancelar: TcxButton;
    cxbReimprimir: TcxButton;
    btnFiltrar: TcxButton;
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cxbReimprimirClick(Sender: TObject);
    procedure cxbCancelarClick(Sender: TObject);
    procedure btnFiltrarClick(Sender: TObject);
    procedure datainiKeyPress(Sender: TObject; var Key: Char);
    procedure datafinKeyPress(Sender: TObject; var Key: Char);
    procedure edt_NumeroKeyPress(Sender: TObject; var Key: Char);

  private
    function BuscarArquivoXML(const diretorio: string; const nomeArquivo: string): string;
    function BuscaArquivoReimpressao : string;
  public
    { Public declarations }
  end;

var
  frmCuponsSAT: TfrmCuponsSAT;

implementation

{$R *.dfm}

uses venda, modulo, principal;

procedure TfrmCuponsSAT.AdvMetroButton1Click(Sender: TObject);
begin
  Close;
end;

function TfrmCuponsSAT.BuscaArquivoReimpressao: string;
begin
  if frmModulo.PathSalvarXmlSatVenda <> EmptyStr then
    Result := BuscarArquivoXML(frmModulo.PathSalvarXmlSatVenda,'AD' + qrSATCHAVE.AsString + '.xml')
  else
    Result := BuscarArquivoXML(ExtractFilePath(Application.ExeName), 'AD' +  qrSATCHAVE.AsString + '.xml');
end;

function TfrmCuponsSAT.BuscarArquivoXML(const diretorio, nomeArquivo: string): string;
var
  arquivos: TStringDynArray;
  i: Integer;
begin
  arquivos := TDirectory.GetFiles(diretorio, nomeArquivo, TSearchOption.soAllDirectories);

  for i := Low(arquivos) to High(arquivos) do
  begin
    if TPath.GetExtension(arquivos[i]) = '.xml' then
    begin
      Result := arquivos[i];
      Exit;
    end;
  end;

  Result := '';

end;

procedure TfrmCuponsSAT.cxbCancelarClick(Sender: TObject);
var
  sCodigo, vArquivo, vChaveCancelamento: string;
begin
  if not(qrSAT.Active) or (qrSAT.IsEmpty) then
  begin
    Application.MessageBox('Nao existe cupom para cancelar!', 'Atenção!', MB_ICONINFORMATION);
    Exit;
  end;
  if qrSATCANCELADO.AsString = 'S' then
  begin
    Application.MessageBox('Este cupom ja foi cancelado!', 'Atenção!', MB_ICONINFORMATION);
    Exit;
  end;
  if (qrSATDATA.AsDateTime <> Date) or ((qrSATDATA.AsDateTime = Date) and (Time - StrToTime(qrSATHORA.AsString) > 30)) then
  begin
    Application.MessageBox('Esta venda ja ultrapassou o prazo legal de 30 min para cancelamento de cupons SAT, impossivel realizar o cacnelamento!',
      'Atenção!', MB_ICONINFORMATION);
    Exit;
  end;
  with frmVenda do
  begin
    ConfiguraSAT;

    vArquivo := BuscaArquivoReimpressao;

    if not (vArquivo = EmptyStr) then
    begin
      frmVenda.ACBrSAT1.CFe.LoadFromFile(vArquivo);

      ACBrSAT1.CFe2CFeCanc;

      Memo1.Lines.Clear;
      Memo1.Lines.Text   := ACBrSAT1.CFeCanc.GerarXML(True);
      vChaveCancelamento := ACBrSAT1.CFeCanc.infCFe.chCanc;

      if Memo1.Lines.Count < 1 then
        ACBrSAT1.CancelarUltimaVenda
      else
        ACBrSAT1.CancelarUltimaVenda(vChaveCancelamento, Memo1.Lines.Text);

      if ACBrSAT1.Resposta.codigoDeRetorno = 7000 then
      begin
        ACBrSAT1.ImprimirExtratoCancelamento;

        with frmModulo do
        begin
          qrCupomCancelar.Open;
          qrCupomCancelar.Insert;
          qrCupomCancelarNUMERO.AsString      := qrSATCUPOM.AsString;
          qrCupomCancelarDATA.AsDateTime      := Date;
          qrCupomCancelarHORA.AsString        := FormatDateTime('HH:MM:SS', Time);
          qrCupomCancelarUSU_CODIGO.AsInteger := icodigo_Usuario;
          qrCupomCancelarUSU_NOME.AsString    := sNome_Operador;
          qrCupomCancelarEX.AsString          := 'N';
          qrCupomCancelar.Post;
          qrCupomCancelar.ApplyUpdates;
        end;
        qrTemp.Close;
        qrTemp.SQL.Clear;
        qrTemp.SQL.Add('Select codigo from cupom where numero = ' + QuotedStr(qrSATCUPOM.AsString));
        qrTemp.Open;
        sCodigo := qrTemp.FieldByName('codigo').AsString;

        qrTemp.Close;
        qrTemp.SQL.Clear;
        qrTemp.SQL.Add('update cupom set cancelado = 1 where numero = ' + QuotedStr(qrSATCUPOM.AsString));
        qrTemp.ExecSQL;

        qrTemp.Close;
        qrTemp.SQL.Clear;
        qrTemp.SQL.Add('update CUPOM_ITEM set cancelado = 1 where cod_cupom = ' + QuotedStr(sCodigo));
        qrTemp.ExecSQL;

        qrSAT.Edit;
        qrSATCANCELADO.AsString            := 'S';
        qrSATXML_CANCELAMENTO.AsAnsiString := ACBrSAT1.CFeCanc.AsXMLString;
        qrSATALTERADO.AsString             := 'S';
        qrSATEX.AsInteger                  := 0;
        qrSAT.Post;
        qrSAT.ApplyUpdates;

        frmModulo.conexao.Commit;
        frmCuponsSAT.Close;
      end
      else
      begin
        Application.MessageBox(PWideChar('Não foi possivel cancelar este cupom:' +
          #13#13 + IntToStr(ACBrSAT1.Resposta.codigoDeRetorno) + ' - ' +
          ACBrSAT1.Resposta.mensagemRetorno), 'Atenção!', MB_ICONINFORMATION);
      end;
    end
    else
      raise Exception.Create('Arquivo .xml não encontrado para efetuar o cancelamento.');
  end;

//    qrSATXML_ENVIO.SaveToFile(ExtractFilePath(Application.ExeName) + '\cancela_sat.xml');
//    ACBrSAT1.CFe.LoadFromFile(ExtractFilePath(Application.ExeName) + '\cancela_sat.xml');
//    ACBrSAT1.CFe2CFeCanc;
//    Memo1.Lines.Clear;
//    Memo1.Lines.Text := ACBrSAT1.CFeCanc.GerarXML(True);
//    ACBrSAT1.CancelarUltimaVenda;
//    ACBrSAT1.ImprimirExtratoCancelamento;
//    if ACBrSAT1.Resposta.codigoDeRetorno = 7000 then
//    begin
//      with frmModulo do
//      begin
//        qrCupomCancelar.Open;
//        qrCupomCancelar.Insert;
//        qrCupomCancelarNUMERO.AsString := qrSATCUPOM.AsString;
//        qrCupomCancelarDATA.AsDateTime := Date;
//        qrCupomCancelarHORA.AsString := FormatDateTime('HH:MM:SS', Time);
//        qrCupomCancelarUSU_CODIGO.AsInteger := icodigo_Usuario;
//        qrCupomCancelarUSU_NOME.AsString := sNome_Operador;
//        qrCupomCancelarEX.AsString := 'N';
//        qrCupomCancelar.Post;
//        qrCupomCancelar.ApplyUpdates;
//      end;
//      qrTemp.Close;
//      qrTemp.SQL.Clear;
//      qrTemp.SQL.Add('Select codigo from cupom where numero = ' + QuotedStr(qrSATCUPOM.AsString));
//      qrTemp.Open;
//      sCodigo := qrTemp.FieldByName('codigo').AsString;
//
//      qrTemp.Close;
//      qrTemp.SQL.Clear;
//      qrTemp.SQL.Add('update cupom set cancelado = 1 where numero = ' + QuotedStr(qrSATCUPOM.AsString));
//      qrTemp.ExecSQL;
//
//      qrTemp.Close;
//      qrTemp.SQL.Clear;
//      qrTemp.SQL.Add('update CUPOM_ITEM set cancelado = 1 where cod_cupom = ' + QuotedStr(sCodigo));
//      qrTemp.ExecSQL;
//
//      qrSAT.Edit;
//      qrSATCANCELADO.AsString := 'S';
//      qrSATXML_CANCELAMENTO.AsAnsiString := ACBrSAT1.CFeCanc.AsXMLString;
//      qrSATALTERADO.AsString := 'S';
//      qrSATEX.AsInteger := 0;
//      qrSAT.Post;
//      qrSAT.ApplyUpdates;
//
//      frmModulo.conexao.Commit;
//    end
//    else
//    begin
//      Application.MessageBox(PWideChar('Não foi possivel cancelar este cupom:' +
//        #13#13 + IntToStr(ACBrSAT1.Resposta.codigoDeRetorno) + ' - ' +
//        ACBrSAT1.Resposta.mensagemRetorno), 'Atenção!', MB_ICONINFORMATION);
//    end;
//  end;
end;

procedure TfrmCuponsSAT.cxbReimprimirClick(Sender: TObject);
var
  arquivoEncontrado : string;
begin
  if not(qrSAT.Active) or (qrSAT.IsEmpty) then
  begin
    Application.MessageBox('Nao existe Cupom para reimprimir!','Atenção!',MB_ICONINFORMATION);
    Exit;
  end;

  frmVenda.ConfiguraSAT;

  arquivoEncontrado := BuscaArquivoReimpressao;

  if arquivoEncontrado <> EmptyStr then
  begin
    frmVenda.ACBrSAT1.CFe.LoadFromFile(arquivoEncontrado);

    if qrSATCANCELADO.AsString = 'S' then
      frmVenda.ACBrSAT1.ImprimirExtratoCancelamento
    else
      frmVenda.ACBrSAT1.ImprimirExtrato;
  end
  else
    raise Exception.Create('Arquivo não encontrado para reimpressão.');

//    ACBrSAT1.CFe.LoadFromFile(ExtractFilePath(Application.ExeName)+'imp_sat.xml');

end;

procedure TfrmCuponsSAT.btnFiltrarClick(Sender: TObject);
begin
  qrSAT.Close;
  qrSAT.SQL.Clear;
  qrSAT.SQL.Add('select nf.* from sat nf where 1=1 ');
  if dataini.Date > 0 then
  begin
    qrSAT.SQL.Add('and nf.data >= :pdataini');
    qrSAT.ParamByName('pdataini').AsDateTime := dataini.Date;
  end;
  if datafin.Date > 0 then
  begin
    qrSAT.SQL.Add('and nf.data <= :pdatafim');
    qrSAT.ParamByName('pdatafim').AsDateTime := datafin.Date;
  end;
  if trim(edt_Numero.Text) <> '' then
  begin
    qrSAT.SQL.Add('and nf.numero = :numero');
    qrSAT.ParamByName('numero').AsString := edt_Numero.Text;
  end;

  qrSAT.SQL.Add('order by nf.data, nf.hora');

  qrSAT.Open;

  if qrSAT.RecordCount > 0 then
    qrSAT.Last;
end;

procedure TfrmCuponsSAT.datafinKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edt_Numero.SetFocus;
end;

procedure TfrmCuponsSAT.datainiKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    datafin.SetFocus;
end;

procedure TfrmCuponsSAT.edt_NumeroKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    btnFiltrarClick(Sender);
end;

procedure TfrmCuponsSAT.FormShow(Sender: TObject);
begin
  dataini.Date := StartOfTheMonth(Date);
  datafin.Date := Now;
  btnFiltrar.Click;
end;

end.
