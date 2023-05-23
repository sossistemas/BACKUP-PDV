unit frmNFCEs;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Data.DB,
  ZAbstractRODataset,
  ZAbstractDataset, ZDataset, Vcl.Grids, Vcl.DBGrids, Vcl.ExtCtrls, NxDBColumns,
  NxColumns,XMLIntf, XMLDoc, zlib, ACBrUtil,
  NxScrollControl, NxCustomGridControl, NxCustomGrid, NxDBGrid, Vcl.ImgList,
  Wwdbigrd, Wwdbgrid, Vcl.Mask, sMaskEdit, sCustomComboEdit, sTooledit,
  pcnConversao, System.ImageList, AdvMetroButton, AdvSmoothPanel,
  AdvSmoothExpanderPanel, MemDS, DBAccess, Uni, pcnAuxiliar, Vcl.OleCtrls,
  SHDocVw, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, cxNavigator,
  cxDBData, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGridLevel, cxClasses, cxGridCustomView, cxGrid, JvExDBGrids, JvDBGrid,
  cxImageComboBox, cxLabel, Vcl.Menus, cxButtons, dxSkinsCore,
  dxSkinsDefaultPainters, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog, cxImageList;

type
  TfrmNotasconsumidor = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ds_nfce: TDataSource;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    edt_Numero: TEdit;
    dataini: TsDateEdit;
    datafin: TsDateEdit;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    qrNFCE: TUniQuery;
    Panel3: TPanel;
    Panel4: TPanel;
    lbEnvio: TLabel;
    lbCancelamento: TLabel;
    qrNFCENUMERO: TIntegerField;
    qrNFCEDATA: TDateField;
    qrNFCETOTAL: TFloatField;
    qrNFCECLIENTE: TStringField;
    qrNFCECHAVE: TStringField;
    qrNFCEXML: TStringField;
    qrNFCESITUACAO: TIntegerField;
    qrNFCETROCO: TFloatField;
    qrNFCEEX: TIntegerField;
    qrNFCEHORA: TStringField;
    qrNFCECONTINGENCIA: TStringField;
    qrNFCEENVIADOCONTINGENCIA: TStringField;
    qrNFCEMOTIVOCONTIGENCIA: TStringField;
    qrNFCEXMLENVIO: TBlobField;
    qrNFCEXMLCACNELAMENTO: TBlobField;
    qrNFCEALTERADO: TStringField;
    qrNFCECUPOM: TStringField;
    qrNFCEXML_CANCELAMENTO: TStringField;
    qrNFCEDES_SIT: TStringField;
    qrTemp: TUniQuery;
    qryTmp: TUniQuery;
    qrNFCEDATAHORACONTIMGENCIA: TDateTimeField;
    qrNFCESTATUS: TStringField;
    qrNFCESTATUS_DESCRICAO: TStringField;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1DBTableView1NUMERO: TcxGridDBColumn;
    cxGrid1DBTableView1DATA: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxGrid1DBTableView1CLIENTE: TcxGridDBColumn;
    cxGrid1DBTableView1CHAVE: TcxGridDBColumn;
    cxGrid1DBTableView1HORA: TcxGridDBColumn;
    cxGrid1DBTableView1CUPOM: TcxGridDBColumn;
    cxGrid1DBTableView1STATUS_DESCRICAO: TcxGridDBColumn;
    cxImageList1: TcxImageList;
    qrNFCESERIE: TIntegerField;
    qrNFCEAMBIENTE: TStringField;
    qrNFCEERRO: TBlobField;
    cxGrid1DBTableView1SERIE: TcxGridDBColumn;
    cxGrid1DBTableView1AMBIENTE: TcxGridDBColumn;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    cxButton4: TcxButton;
    cxButton5: TcxButton;
    procedure FormShow(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure qrNFCEAfterScroll(DataSet: TDataSet);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
  private
    { Private declarations }
    procedure LoadXML(RetWS: String; MyWebBrowser: TWebBrowser);
  public
    { Public declarations }
  end;

var
  frmNotasconsumidor: TfrmNotasconsumidor;

implementation

uses venda, modulo, principal, ResultadoWebService, senha_supervisor, justificativa,
  JustificativaInutilizacaoNFCe;


{$R *.dfm}


procedure TfrmNotasconsumidor.AdvMetroButton1Click(Sender: TObject);
begin
 Close;
end;

procedure TfrmNotasconsumidor.cxButton1Click(Sender: TObject);
begin
  qrNFCE.Close;
  qrNFCE.SQL.Clear;
  qrNFCE.SQL.Add('select nf.*, case when situacao = 0 then ' + QuotedStr('Emitido') + ' when situacao = 1 then ' + QuotedStr('Cancelado') + ' else ' + QuotedStr('Inutilizado') + ' end des_sit, ');
  qrNFCE.SQL.Add('  case');
  qrNFCE.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''S'' then');
  qrNFCE.SQL.Add('      ''CONTINGENCIA ENVIADA''');
  qrNFCE.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''N'' and nf.status <> ''E'' then');
  qrNFCE.SQL.Add('      ''CONTINGENCIA PENDENTE''');
  qrNFCE.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''N'' and nf.status = ''E'' then');
  qrNFCE.SQL.Add('      ''CONTINGENCIA COM ERRO''');
  qrNFCE.SQL.Add('    when nf.contingencia = ''N'' and nf.situacao = 0 then');
  qrNFCE.SQL.Add('      ''ENVIO NORMAL''');
  qrNFCE.SQL.Add('    when nf.contingencia = ''N'' and nf.situacao = 1 then');
  qrNFCE.SQL.Add('      ''CANCELADA''');
  qrNFCE.SQL.Add('    when nf.situacao = 2 then');
  qrNFCE.SQL.Add('      ''INUTILIZAÇÃO ENVIADA''');
  qrNFCE.SQL.Add('  end status_descricao');
  qrNFCE.SQL.Add('from NFCE nf where 1=1 ');
  if Length(edt_Numero.Text) > 0 then begin
    qrNFCE.SQL.Add('and numero = :pnumero ');
    qrNFCE.ParamByName('pnumero').AsInteger := strtoint(edt_Numero.Text);
  end;

  if dataini.Date > 0 then begin
    qrNFCE.SQL.Add('and data >= :pdataini ');
    qrNFCE.ParamByName('pdataini').AsDateTime := dataini.Date;
  end;

  if datafin.Date > 0 then begin
    qrNFCE.SQL.Add('and data <= :pdatafin ');
    qrNFCE.ParamByName('pdatafin').AsDateTime := datafin.Date;
  end;
  qrNFCE.SQL.Add('order by numero');
  qrNFCE.Open;
  qrNFCE.First;
end;

procedure TfrmNotasconsumidor.cxButton2Click(Sender: TObject);
begin
  if qrNFCE.RecordCount > 0 then
  begin
    with frmModulo do
    begin
      LerConfiguracao;
      ACBRNFCe.WebServices.Consulta.NFeChave :=
        qrNFCE.FieldByName('chave').AsString;
      if ACBRNFCe.WebServices.Consulta.Executar then begin
        Application.CreateForm(TfrmResultadoWebService,frmResultadoWebService);
        LoadXML(ACBRNFCe.WebServices.Consulta.RetornoWS, frmResultadoWebService.WBResposta);
        frmResultadoWebService.pnTexto.Caption := IntToStr(ACBRNFCe.WebServices.Consulta.cStat) + ' - ' + ACBRNFCe.WebServices.Consulta.xMotivo;
        frmResultadoWebService.ShowModal;
        FreeAndNil(frmResultadoWebService);
      end else begin
        Application.MessageBox('Ocorreu um erro ao efetuar a consulta.','Atenção!',MB_ICONINFORMATION);
      end;
    end;
  end;
end;

procedure TfrmNotasconsumidor.cxButton3Click(Sender: TObject);
var
  Chave, idLote, CNPJ, Protocolo, Justificativa, sCodigo: string;
begin
  if qrNFCE.RecordCount > 0 then begin
    if qrNFCE.FieldByName('situacao').AsInteger = 0 then begin
      if Application.MessageBox('Deseja realmente cancelar esta NFC-e, este processo é irreversível?','Atenção',MB_ICONQUESTION+MB_YESNO) <> mrYes then
        Exit;
      Application.CreateForm(TfrmSenha_Supervisor,frmSenha_Supervisor);
      frmSenha_Supervisor.ShowModal;
      FreeAndNil(frmSenha_Supervisor);
      if not bSupervisor_autenticado then begin
        Exit;
      end;
      Chave := qrNFCE.FieldByName('chave').AsString;
      idLote := '1';
      CNPJ := copy(Chave, 7, 14);
      Protocolo := '';
      Application.CreateForm(Tfrmjustificativa, frmjustificativa);
      frmjustificativa.ShowModal;
      Justificativa := frmjustificativa.edTexto.Lines.Text;
      FreeAndNil(frmjustificativa);
      with frmModulo do begin
        LerConfiguracao;
        ACBRNFCe.NotasFiscais.Clear;
        ACBRNFCe.NotasFiscais.LoadFromFile(qrNFCE.FieldByName('xml').AsString);

        idLote := '1';
        ACBRNFCe.EventoNFe.Evento.Clear;
        ACBRNFCe.EventoNFe.idLote := strtoint(idLote);
        with ACBRNFCe.EventoNFe.Evento.Add do begin
          infEvento.dhEvento := now;
          infEvento.tpEvento := teCancelamento;
          infEvento.detEvento.xJust := Justificativa;
        end;
        ACBRNFCe.EnviarEvento(strtoint(idLote));

        Application.CreateForm(TfrmResultadoWebService,frmResultadoWebService);
        LoadXML(ACBRNFCe.WebServices.EnvEvento.RetornoWS, frmResultadoWebService.WBResposta);
        frmResultadoWebService.pnTexto.Caption := IntToStr(ACBRNFCe.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.cStat) + ' - ' + ACBRNFCe.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.xMotivo;
        frmResultadoWebService.ShowModal;
        FreeAndNil(frmResultadoWebService);
        if ACBRNFCe.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.cStat = 135 then begin
          qrCupomCancelar.Open;
          qrCupomCancelar.Insert;
          qrCupomCancelarNUMERO.AsString := qrNFCECUPOM.AsString;
          qrCupomCancelarDATA.AsDateTime := Date;
          qrCupomCancelarHORA.AsString := FormatDateTime('HH:MM:SS',Time);
          qrCupomCancelarUSU_CODIGO.AsInteger := icodigo_Usuario;
          qrCupomCancelarUSU_NOME.AsString := sNome_Operador;
          qrCupomCancelarEX.AsString := 'N';
          qrCupomCancelar.Post;
          qrCupomCancelar.ApplyUpdates;

          qrTemp.Close;
          qrTemp.SQL.Clear;
          qrTemp.SQL.Add('Select codigo from cupom where numero = ' + QuotedStr(qrNFCECUPOM.AsString));
          qrTemp.Open;
          sCodigo := qrTemp.FieldByName('codigo').AsString;

          qrTemp.Close;
          qrTemp.SQL.Clear;
          qrTemp.SQL.Add('update cupom set cancelado = 1 where numero = ' + QuotedStr(qrNFCECUPOM.AsString));
          qrTemp.ExecSQL;

          qrTemp.Close;
          qrTemp.SQL.Clear;
          qrTemp.SQL.Add('update CUPOM_ITEM set cancelado = 1 where cod_cupom = ' + QuotedStr(sCodigo));
          qrTemp.ExecSQL;

          qrNFCE.Edit;
          qrNFCE.FieldByName('situacao').AsInteger := 1;
          qrNFCE.FieldByName('xml_cancelamento').AsString := ACBRNFCe.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.NomeArquivo;
          qrNFCEXMLCACNELAMENTO.LoadFromFile(ACBRNFCe.WebServices.EnvEvento.EventoRetorno.retEvento.Items[0].RetInfEvento.NomeArquivo);
          qrNFCEEX.AsInteger := 0;
          qrNFCEALTERADO.AsString := 'S';
          qrNFCE.Post;
          qrNFCE.ApplyUpdates;
          Conexao.Commit;
          cxButton1.Click;
        end;
      end;
    end else begin
      Application.MessageBox('NFC-e ja está cancelada.','Atenção!',MB_ICONINFORMATION);
      Exit;
    end;
  end;
end;

procedure TfrmNotasconsumidor.cxButton4Click(Sender: TObject);
begin
  if qrNFCE.RecordCount > 0 then begin
    if qrNFCE.FieldByName('situacao').AsInteger = 0 then begin
      with frmModulo do begin
        ACBRNFCe.NotasFiscais.Clear;
        ACBRNFCe.NotasFiscais.LoadFromFile(qrNFCE.FieldByName('xml').AsString);
        ACBRDANFENFCe.FastFile := 'C:\SOS\PDV\Schemas\DANFeNFCe.fr3';
        if FileExists(frmPrincipal.LerIni(sConfiguracoes, 'PDV', 'CAMINHO_LOGO', '')) then
          ACBRDANFENFCe.Logo := frmPrincipal.LerIni(sConfiguracoes, 'PDV','CAMINHO_LOGO', '');
        ACBRDANFENFCe.vTroco := qrNFCETROCO.AsFloat;
        ACBRNFCe.NotasFiscais.Imprimir;
      end;
    end;
  end;
end;

procedure TfrmNotasconsumidor.cxButton5Click(Sender: TObject);
var
  Modelo, Serie, Ano, NumeroInicial, NumeroFinal, Justificativa, tpAmb, Protocolo: string;
  dhRecbto: TDateTime;
begin
  with frmModulo do begin
    Application.CreateForm(TfrmJustificativaInutilizacaoNFCe, frmJustificativaInutilizacaoNFCe);
    frmJustificativaInutilizacaoNFCe.ShowModal;
    if not frmJustificativaInutilizacaoNFCe.Continua then
      Exit;
    Modelo := frmJustificativaInutilizacaoNFCe.edModelo.Text;
    Serie := frmJustificativaInutilizacaoNFCe.edSerie.Text;
    Ano := frmJustificativaInutilizacaoNFCe.edAno.Text;
    NumeroInicial := frmJustificativaInutilizacaoNFCe.edNumero.Text;
    NumeroFinal := frmJustificativaInutilizacaoNFCe.edNumero.Text;
    Justificativa := frmJustificativaInutilizacaoNFCe.edJustificativa.Lines.Text;
    FreeAndNil(frmJustificativaInutilizacaoNFCe);
    qryTmp.Close;
    qryTmp.SQL.Text := 'select * from nfce where numero = ' + QuotedStr(NumeroInicial);
    qryTmp.Open;
    if not qryTmp.IsEmpty then begin
      if (qryTmp.FieldByName('situacao').AsString = '1') then begin
        Application.MessageBox('Não é necessário Inutilizar uma numeração de uma NFC-e ja Cancelada.', 'Atenção!', MB_ICONINFORMATION);
        Exit;
      end;
      if (qryTmp.FieldByName('situacao').AsString = '0') then begin
        Application.MessageBox('Não é posivel inutilizar um NFC-e enviada.', 'Atenção!', MB_ICONINFORMATION);
        Exit;
      end;
    end;
    frmmodulo.qrfilial.Close;
    frmmodulo.qrfilial.Open;
    ACBRNFCe.WebServices.Inutiliza(frmmodulo.qrFilialCNPJ.AsString, Justificativa, StrToInt(Ano), StrToInt(Modelo), StrToInt(Serie), StrToInt(NumeroInicial), StrToInt(NumeroFinal));
    if ACBRNFCe.WebServices.Inutilizacao.cStat = 102 then begin
      qrNFCE.Append;
      qrNFCENUMERO.AsString := NumeroInicial;
      qrNFCEsituacao.AsInteger := 2;
      qrNFCEXMLENVIO.LoadFromFile(ACBRNFCe.WebServices.Inutilizacao.NomeArquivo);

      qrNFCEDATA.AsDateTime := Date;
      qrNFCETOTAL.AsFloat := 0;
      qrNFCEHORA.AsString := FormatDateTime('HH:MM:SS',Time);
      qrNFCECONTINGENCIA.AsString := 'N';
      qrNFCEENVIADOCONTINGENCIA.AsString := 'N';
      qrNFCEALTERADO.AsString := 'N';
      qrNFCESERIE.AsString := Serie;
      if ACBRNFCe.Configuracoes.WebServices.Ambiente  = taProducao then
        qrNFCEAMBIENTE.AsString := 'P'
      else
        qrNFCEAMBIENTE.AsString := 'H';
      qrNFCESTATUS_DESCRICAO.AsString := 'INUTILIZAÇÃO ENVIADA';
      qrNFCEEX.AsInteger := 0;
      qrNFCE.Post;
      qrNFCE.ApplyUpdates;
      Conexao.Commit;
      qrNFCE.Close;
      qrNFCE.Open;
      qrNFCE.Locate('NUMERO',StrToInt(NumeroInicial),[loCaseInsensitive]);
    end;
    //Gravar no Banco
    tpAmb := TpAmbToStr(ACBRNFCe.WebServices.Inutilizacao.TpAmb);
    dhRecbto := ACBRNFCe.WebServices.Inutilizacao.dhRecbto;
    Protocolo := ACBRNFCe.WebServices.Inutilizacao.Protocolo;
  end;

end;

procedure TfrmNotasconsumidor.FormShow(Sender: TObject);
begin
  dataini.Date := now;
  datafin.Date := now;
  cxButton1.Click;
end;


procedure TfrmNotasconsumidor.LoadXML(RetWS: String; MyWebBrowser: TWebBrowser);
begin
  ACBrUtil.WriteToTXT( PathWithDelim(ExtractFileDir(application.ExeName))+'temp.xml',
                        ACBrUtil.ConverteXMLtoUTF8( RetWS ), False, False);
  MyWebBrowser.Navigate(PathWithDelim(ExtractFileDir(application.ExeName))+'temp.xml');
end;

procedure TfrmNotasconsumidor.qrNFCEAfterScroll(DataSet: TDataSet);
begin
  lbEnvio.Caption := 'XML de Envio: ' + qrNFCEXML.AsString;
  lbCancelamento.Caption := 'XML de Cancelamento: ' + qrNFCEXML_CANCELAMENTO.AsString;
end;

end.


