unit EnviaContingencia;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, Data.DB, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, AdvMetroButton, Vcl.StdCtrls, AdvSmoothPanel, AdvSmoothExpanderPanel,
  Vcl.ExtCtrls, MemDS, DBAccess, Uni, AdvGlowButton, Vcl.Mask, sMaskEdit,
  sCustomComboEdit, sToolEdit, cxContainer, cxCheckBox, cxImageComboBox,
  System.ImageList, Vcl.ImgList, pcnConversao, cxMemo, System.Math, cxLabel,
  Vcl.Menus, cxButtons;

type
  TfrmEnviaContingencia = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label1: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    dsContingencia: TDataSource;
    qrContingencia: TUniQuery;
    qrContingenciaNUMERO: TIntegerField;
    qrContingenciaDATA: TDateField;
    qrContingenciaTOTAL: TFloatField;
    qrContingenciaCLIENTE: TStringField;
    qrContingenciaCHAVE: TStringField;
    qrContingenciaXML: TStringField;
    qrContingenciaSITUACAO: TIntegerField;
    qrContingenciaTROCO: TFloatField;
    qrContingenciaEX: TIntegerField;
    qrContingenciaHORA: TStringField;
    qrContingenciaCONTINGENCIA: TStringField;
    qrContingenciaENVIADOCONTINGENCIA: TStringField;
    qrContingenciaMOTIVOCONTIGENCIA: TStringField;
    qrContingenciaXMLENVIO: TBlobField;
    qrContingenciaXMLCACNELAMENTO: TBlobField;
    qrContingenciaALTERADO: TStringField;
    qrContingenciaCUPOM: TStringField;
    qrContingenciaXML_CANCELAMENTO: TStringField;
    qrContingenciaDATAHORACONTIMGENCIA: TDateTimeField;
    qrContingenciaSTATUS: TStringField;
    qrContingenciaSTATUS_DESCRICAO: TStringField;
    cxGrid1DBTableView1NUMERO: TcxGridDBColumn;
    cxGrid1DBTableView1DATA: TcxGridDBColumn;
    cxGrid1DBTableView1TOTAL: TcxGridDBColumn;
    cxGrid1DBTableView1CLIENTE: TcxGridDBColumn;
    cxGrid1DBTableView1HORA: TcxGridDBColumn;
    cxGrid1DBTableView1CUPOM: TcxGridDBColumn;
    cxGrid1DBTableView1STATUS_DESCRICAO: TcxGridDBColumn;
    Panel3: TPanel;
    Label2: TLabel;
    dataini: TsDateEdit;
    Label3: TLabel;
    datafin: TsDateEdit;
    Label4: TLabel;
    edPendencias: TcxCheckBox;
    edHoje: TcxCheckBox;
    cxImageList1: TcxImageList;
    qrProdNCFE: TUniQuery;
    qrContingenciaERRO: TBlobField;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    qrContingenciaACAO: TStringField;
    cxImageList2: TcxImageList;
    colErro: TcxGridDBColumn;
    cxGrid1DBTableView1Column2: TcxGridDBColumn;
    pnMens: TPanel;
    qrServidor: TUniQuery;
    QRCSOSN: TUniQuery;
    qrContingenciaSERIE: TIntegerField;
    qrContingenciaAMBIENTE: TStringField;
    cxGrid1DBTableView1Column3: TcxGridDBColumn;
    cxGrid1DBTableView1Column4: TcxGridDBColumn;
    cxButton2: TcxButton;
    cxButton1: TcxButton;
    cxButton3: TcxButton;
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure edHojeClick(Sender: TObject);
    procedure datainiChange(Sender: TObject);
    procedure datafinChange(Sender: TObject);
    procedure colErroGetCellHint(Sender: TcxCustomGridTableItem;
      ARecord: TcxCustomGridRecord; ACellViewInfo: TcxGridTableDataCellViewInfo;
      const [Ref] AMousePos: TPoint; var AHintText: TCaption;
      var AIsHintMultiLine: Boolean; var AHintTextRect: TRect);
    procedure AdvGlowButton3Click(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
  private
    { Private declarations }
    Procedure AjustaEnviaNFCE;
  public
    { Public declarations }
  end;

var
  frmEnviaContingencia: TfrmEnviaContingencia;

implementation

uses modulo, venda, principal, funcoes;

{$R *.dfm}

procedure TfrmEnviaContingencia.AdvGlowButton2Click(Sender: TObject);
begin
end;

//Adiconado por diego molina - 07/02/2019----------------------------------------
// adiconado para enviar apenas a nfce selecioanda
procedure TfrmEnviaContingencia.AdvGlowButton3Click(Sender: TObject);
begin
end;
//------------------------------------------------------------------------------

procedure TfrmEnviaContingencia.AdvMetroButton1Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmEnviaContingencia.AjustaEnviaNFCE;
var
  i:Integer;
  total_trib, total_servico, total_produtos, total_icms, total_base:double;
begin
  with frmModulo do begin
    LerConfiguracao;
    if EdFormatoOff = 0 then
      frmtOffLine := teContingencia
    else
      frmtOffLine := teOffLine;
    ACBRNFCe.NotasFiscais.Clear;
    try
      if (qrContingenciaCONTINGENCIA.AsString = 'S') and (qrContingenciaENVIADOCONTINGENCIA.AsString = 'N') then begin
        pnMens.Caption := 'Enviando NFC-e N° ' + qrContingenciaNUMERO.AsString + '...';
        qrContingencia.Edit;
        qrContingenciaACAO.AsString := 'E';
        qrContingencia.Post;
        Update;
        ACBRNFCe.NotasFiscais.LoadFromFile(qrContingenciaXML.AsString);
        if qrContingenciaSTATUS.AsString = 'E' then begin
          with ACBRNFCe.NotasFiscais.Items[0].NFe do begin
            Ide.cUF := UFtoCUF(edtEmitUF);
            Ide.cMunFG := StrToInt(edtEmitCodCidade);
            Emit.CNPJCPF := edtEmitCNPJ;
            Emit.IE := edtEmitIE;
            Emit.IM := edtEmitIM;
            Emit.xNome := edtEmitRazao;
            Emit.xFant := edtEmitFantasia;
            Emit.EnderEmit.fone := edtEmitFone;
            Emit.EnderEmit.CEP := StrToInt(edtEmitCEP);
            Emit.EnderEmit.xLgr := edtEmitLogradouro;
            Emit.EnderEmit.nro := edtEmitNumero;
            Emit.EnderEmit.xCpl := edtEmitComp;
            Emit.EnderEmit.xBairro := edtEmitBairro;
            Emit.EnderEmit.cMun := StrToInt(edtEmitCodCidade);
            Emit.EnderEmit.xMun := edtEmitCidade;
            Emit.EnderEmit.UF := edtEmitUF;
            Emit.EnderEmit.cPais := 1058;
            Emit.EnderEmit.xPais := 'BRASIL';
            Emit.IEST := '';
            if qrfilial.fieldbyname('CRT').asinteger = 1 then
              Emit.CRT := crtSimplesNacional
            else if qrfilial.fieldbyname('CRT').asinteger = 2 then
              Emit.CRT := crtSimplesExcessoReceita
            else
              Emit.CRT := crtRegimeNormal;
            total_trib := 0;
            total_servico := 0;
            total_produtos := 0;
            total_icms := 0;
            total_base := 0;
            for i:=0 to (ACBRNFCe.NotasFiscais.Items[0].NFe.Det.Count - 1) do begin
              with ACBRNFCe.NotasFiscais.Items[0].NFe.Det.Items[i] do begin
                qrProdNCFE.close;
                qrProdNCFE.sql.clear;
                qrProdNCFE.sql.add('select * from ESTOQUE ');
                qrProdNCFE.sql.add('where CODIGO = :pcodigo');
                qrProdNCFE.ParamByName('pcodigo').asinteger := StrToInt(Prod.cProd);
                qrProdNCFE.Open;
                if qrProdNCFE.FieldByName('informa_codigo_barra_xml').AsString = 'S' then begin
                  Prod.cEAN := qrProdNCFE.fieldbyname('COD_BARRA').asstring;
                  Prod.cEANTrib := qrProdNCFE.fieldbyname('COD_BARRA').asstring;
                end else begin
                  Prod.cEAN := 'SEM GTIN';
                  Prod.cEANTrib := 'SEM GTIN';
                end;
                Prod.xProd := qrProdNCFE.fieldbyname('NOME').asstring;
                Prod.NCM := somenteNumero(qrProdNCFE.fieldbyname('NCM').asstring);
                Prod.EXTIPI := '';
                if Trim(qrProdNCFE.fieldbyname('CFOP').AsString) <> '' then
                  Prod.CFOP := Trim(qrProdNCFE.fieldbyname('CFOP').AsString)
                else
                  Prod.CFOP := edtcfop;
                QRCSOSN.close;
                QRCSOSN.sql.clear;
                QRCSOSN.sql.add('select pro.csosn, pro.aliquota icms, pro.cest, pro.codigo_servico servico ');
                QRCSOSN.sql.Add('from estoque pro');
                QRCSOSN.sql.add('where pro.codigo = :pcodigo');
                QRCSOSN.ParamByName('pcodigo').asstring := Zerar(Prod.cProd,6);
                QRCSOSN.Open;
                if (Trim(QRCSOSN.fieldbyname('cest').asstring) <> '') and (Trim(QRCSOSN.fieldbyname('cest').asstring) <> '0') then
                  Prod.CEST := QRCSOSN.fieldbyname('cest').asstring;


                with Imposto do begin
                  vTotTrib := RoundTo((Prod.vProd * qrProdNCFE.fieldbyname('ALIQNACIONAL').AsFloat) / 100, -2);
                  total_trib := total_trib + vTotTrib;
                  if Trim(QRCSOSN.FieldByName('servico').AsString) <> '' then begin
                    qrServidor.Close;
                    qrServidor.SQL.Clear;
                    qrServidor.SQL.Add('select emp.iss, emp.cod_municipio_ibge from empresa emp');
                    qrServidor.Open;
                    ISSQN.vBC := Prod.vProd;
                    ISSQN.vAliq := qrServidor.FieldByName('iss').AsFloat;
                    ISSQN.vISSQN := RoundTo((ISSQN.vBC * ISSQN.vAliq) / 100, -2);
                    ISSQN.cMunFG := qrServidor.FieldByName('cod_municipio_ibge').AsInteger;
                    ISSQN.cListServ := Trim(QRCSOSN.FieldByName('servico').AsString);
                    ISSQN.indISS := iiExigivel;
                    ISSQN.cMun := qrServidor.FieldByName('cod_municipio_ibge').AsInteger;
                    ISSQN.indIncentivo := iiSim;
                    PIS.CST := pis07;
                    COFINS.CST := cof07;
                    with Total.ISSQNtot do begin
                      vServ := ISSQN.vBC;
                      vBC := ISSQN.vBC;
                      vISS := ISSQN.vISSQN;
                      dCompet := Date;
                    end;
                    total_servico := RoundTo(total_servico + Prod.vProd, -2);
                  end
                  else begin
                    // simples nacional
                    total_produtos := RoundTo(total_produtos + Prod.vProd, -2);
                    if qrfilial.fieldbyname('CRT').asinteger = 1 then begin
                      with ICMS do begin
                        if (QRCSOSN.fieldbyname('csosn').asstring = '101') then
                          ICMS.CSOSN := csosn101
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '102') then
                          ICMS.CSOSN := csosn102
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '103') then
                          ICMS.CSOSN := csosn103
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '300') then
                          ICMS.CSOSN := csosn300
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '400') then
                          ICMS.CSOSN := csosn400
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '201') then
                          ICMS.CSOSN := csosn201
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '202') then
                          ICMS.CSOSN := csosn202
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '500') then
                          ICMS.CSOSN := csosn500
                        else if (QRCSOSN.fieldbyname('csosn').asstring = '900') then
                          ICMS.CSOSN := csosn900;
                        case ICMS.CSOSN of
                          csosn101:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                              ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                            end;
                          csosn102:
                            begin
                              ICMS.orig := oeNacional;
                            end;
                          csosn201:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0;
                              ICMS.pRedBCST := 00.00;
                              ICMS.vBCST := 0;
                              ICMS.pICMSST := 00.00;
                              ICMS.vICMSST := 0;
                              ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                              ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                            end;
                          csosn202:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0;
                              ICMS.pRedBCST := 00.00;
                              ICMS.vBCST := 0;
                              ICMS.pICMSST := 00.00;
                              ICMS.vICMSST := 0;
                            end;
                          csosn500:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.vBCSTRet := 0;
                              ICMS.vICMSSTRet := 0;
                            end;
                          csosn900:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pRedBC := 0;
                              ICMS.pICMS := RoundTo(QRCSOSN.FieldByName('icms').AsFloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0;
                              ICMS.pRedBCST := 00.00;
                              ICMS.vBCST := 0;
                              ICMS.pICMSST := 00.00;
                              ICMS.vICMSST := 0;
                              ICMS.pCredSN := 0.0; // Colocar o percentual do Crédito
                              ICMS.vCredICMSSN := 0.0; // Colocar o valor do Crédito
                            end;
                        end;
                        with ICMSUFDest do begin
                          vBCUFDest := 0.00;
                          pFCPUFDest := 0.00;
                          pICMSUFDest := 0.00;
                          pICMSInter := 0.00;
                          pICMSInterPart := 0.00;
                          vFCPUFDest := 0.00;
                          vICMSUFDest := 0.00;
                          vICMSUFRemet := 0.00;
                        end;
                      end;
                    end
                    else begin
                      with ICMS do begin
                        if (qrProdNCFE.FieldByName('CST').asstring = '000') then
                          CST := cst00
                        else if (qrProdNCFE.FieldByName('CST').asstring = '010') then
                          CST := cst10
                        else if (qrProdNCFE.FieldByName('CST').asstring = '020') then
                          CST := cst20
                        else if (qrProdNCFE.FieldByName('CST').asstring = '030') then
                          CST := cst30
                        else if (qrProdNCFE.FieldByName('CST').asstring = '040') then
                          CST := cst40
                        else if (qrProdNCFE.FieldByName('CST').asstring = '050') then
                          CST := cst50
                        else if (qrProdNCFE.FieldByName('CST').asstring = '060') then
                          CST := cst60
                        else if (qrProdNCFE.FieldByName('CST').asstring = '070') then
                          CST := cst70
                        else if (qrProdNCFE.FieldByName('CST').asstring = '080') then
                          CST := cst80
                        else if (qrProdNCFE.FieldByName('CST').asstring = '090') then
                          CST := cst90;
                        case ICMS.CST of
                          cst00:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                          cst10:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0;
                              ICMS.pRedBCST := 00.00;
                              ICMS.vBCST := 0;
                              ICMS.pICMSST := 00.00;
                              ICMS.vICMSST := 0;
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                          cst20:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.pRedBC := 0.00;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              ICMS.vICMSDeson := 0.00;
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                          cst30:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0.00;
                              ICMS.pRedBCST := 0.00;
                              ICMS.vBCST := 0;
                              ICMS.pICMSST := 0;
                              ICMS.vICMSST := 0;
                              ICMS.vICMSDeson := 0.00;
                            end;
                          cst40, cst41, cst50:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.vICMSDeson := 0.00;
                            end;
                          cst51:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.pRedBC := 0.00;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMSOp := 0.00;
                              ICMS.pDif := 0.00;
                              ICMS.vICMSDif := 0.00;
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                          cst60:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.vBCSTRet := 0.00;
                              ICMS.vICMSSTRet := 0.00;
                            end;
                          cst70:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.pRedBC := 0.00;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0.00;
                              ICMS.pRedBCST := 0.00;
                              ICMS.vBCST := 0.00;
                              ICMS.pICMSST := 0.00;
                              ICMS.vICMSST := 0.00;
                              ICMS.vICMSDeson := 0.00;
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                          cst90:
                            begin
                              ICMS.orig := oeNacional;
                              ICMS.modBC := dbiValorOperacao;
                              ICMS.vBC := Prod.vProd;
                              ICMS.pRedBC := 0.00;
                              ICMS.pICMS := RoundTo(qrProdNCFE.FieldByName('ALIQUOTA').asfloat, -2);
                              ICMS.vICMS := RoundTo((ICMS.vBC * ICMS.pICMS) / 100, -2);
                              ICMS.modBCST := dbisMargemValorAgregado;
                              ICMS.pMVAST := 0.00;
                              ICMS.pRedBCST := 0.00;
                              ICMS.vBCST := 0.00;
                              ICMS.pICMSST := 0.00;
                              ICMS.vICMSST := 0.00;
                              ICMS.vICMSDeson := 0.00;
                              total_icms := total_icms + ICMS.vICMS;
                              total_base := total_base + ICMS.vBC;
                            end;
                        end;
                      end;
                    end;
                  end;
                end;

              end;
            end;
            Total.ICMSTot.vBC := total_base;
            Total.ICMSTot.vICMS := total_icms;
            Total.ICMSTot.vTotTrib := total_trib;
          end;
        end;

        ACBRNFCe.Enviar(ACBRNFCe.NotasFiscais.Items[0].NFe.Ide.nNF, False, True);
        if ACBRNFCe.NotasFiscais.Items[0].Confirmada then begin
          if ACBRNFCe.WebServices.Enviar.cStat = 100 then begin
            qrContingencia.Edit;
            qrContingenciaENVIADOCONTINGENCIA.AsString := 'S';
            qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA ENVIADA';
            qrContingenciaSTATUS.AsString := 'V';
            qrContingenciaEX.AsInteger := 0;
            qrContingenciaALTERADO.AsString := 'S';
            qrContingenciaXMLENVIO.LoadFromFile(qrContingenciaXML.AsString);
            qrContingenciaACAO.AsString := ' ';
            qrContingencia.Post;
            qrContingencia.ApplyUpdates;
          end;
        end else begin
          // adiconado por diego molina 12/02/2019 ---------------------------
          //verificadndo se nota fiscal ja exite e retornando valido----------
          if ACBRNFCe.Consultar then
            if ACBRNFCe.WebServices.Consulta.cStat = 100 then
              begin
                qrContingencia.Edit;
                qrContingenciaENVIADOCONTINGENCIA.AsString := 'S';
                qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA ENVIADA';
                qrContingenciaSTATUS.AsString := 'V';
                qrContingenciaEX.AsInteger := 0;
                qrContingenciaALTERADO.AsString := 'S';
                qrContingenciaXMLENVIO.LoadFromFile(qrContingenciaXML.AsString);
                qrContingenciaACAO.AsString := ' ';
                qrContingencia.Post;
                qrContingencia.ApplyUpdates;
              end
            else
              begin
                qrContingencia.Edit;
                qrContingenciaSTATUS.AsString := 'E';
                qrContingenciaERRO.AsString := ACBRNFCe.WebServices.Enviar.xMotivo;
                qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA COM ERRO';
                qrContingenciaACAO.AsString := ' ';
                qrContingencia.Post;
                qrContingencia.ApplyUpdates;
              end;
          //---------------------------------------------------------
         { qrContingencia.Edit;
          qrContingenciaSTATUS.AsString := 'E';
          qrContingenciaERRO.AsString := ACBRNFCe.WebServices.Enviar.xMotivo;
          qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA COM ERRO';
          qrContingenciaACAO.AsString := ' ';
          qrContingencia.Post;
          qrContingencia.ApplyUpdates;   }
        end;
      end;
    except
      on e:exception do begin
        // adiconado por diego molina 07/02/2019 ---------------------------
        //verificadndo se nota fiscal ja exite e retornando valido----------
        if (Pos('Duplicidade de NF-e', E.Message) > 0) or (Pos('NF-e ja esta cancelada na', E.Message) > 0) then
          begin
            if ACBRNFCe.Consultar then
              if ACBRNFCe.WebServices.Consulta.cStat = 100 then
                begin
                  qrContingencia.Edit;
                  qrContingenciaENVIADOCONTINGENCIA.AsString := 'S';
                  qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA ENVIADA';
                  qrContingenciaSTATUS.AsString := 'V';
                  qrContingenciaEX.AsInteger := 0;
                  qrContingenciaALTERADO.AsString := 'S';
                  qrContingenciaXMLENVIO.LoadFromFile(qrContingenciaXML.AsString);
                  qrContingenciaACAO.AsString := ' ';
                  qrContingencia.Post;
                  qrContingencia.ApplyUpdates;
                end
              else
                begin
                  qrContingencia.Edit;
                  qrContingenciaSTATUS.AsString := 'E';
                  qrContingenciaERRO.AsString := ACBRNFCe.WebServices.Enviar.xMotivo;
                  qrContingenciaSTATUS_DESCRICAO.AsString := 'CONTINGENCIA COM ERRO';
                  qrContingenciaACAO.AsString := ' ';
                  qrContingencia.Post;
                  qrContingencia.ApplyUpdates;
                end;
          end //------------------------------------------------------------
        else
         begin
          qrContingencia.Edit;
          qrContingenciaSTATUS.AsString := 'E';
          qrContingenciaACAO.AsString := ' ';
          qrContingenciaERRO.AsString := e.Message;
          qrContingencia.Post;
          qrContingencia.ApplyUpdates;
         end;
      end;
    end;
  end;

end;

procedure TfrmEnviaContingencia.colErroGetCellHint(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  ACellViewInfo: TcxGridTableDataCellViewInfo; const [Ref] AMousePos: TPoint;
  var AHintText: TCaption; var AIsHintMultiLine: Boolean;
  var AHintTextRect: TRect);
begin

  if qrContingenciaSTATUS.AsString = 'E' then
    AHintText := cxGrid1DBTableView1.DataController.GetDisplayText(ARecord.Index, 11)
  else
    AHintText := '';
end;

procedure TfrmEnviaContingencia.cxButton1Click(Sender: TObject);
begin
  if not(qrContingencia.Active) or (qrContingencia.IsEmpty) then begin
    Application.MessageBox('Não existem NFC-e para transmitir!','Atenção!',MB_ICONINFORMATION);
    Exit;
  end;
  try
    AjustaEnviaNFCE;
  finally
    pnMens.Caption := '';
  end;
end;

procedure TfrmEnviaContingencia.cxButton2Click(Sender: TObject);
begin
  if not(qrContingencia.Active) or (qrContingencia.IsEmpty) then begin
    Application.MessageBox('Não existem NFC-e para transmitir!','Atenção!',MB_ICONINFORMATION);
    Exit;
  end;
  try
    qrContingencia.First;
    while not qrContingencia.Eof do begin
      AjustaEnviaNFCE;
      qrContingencia.Next;
    end;
  finally
    pnMens.Caption := '';
  end;
end;

procedure TfrmEnviaContingencia.cxButton3Click(Sender: TObject);
begin
  qrContingencia.Close;
  qrContingencia.SQL.Clear;
  qrContingencia.SQL.Add('select nf.*,');
  qrContingencia.SQL.Add('  case');
  qrContingencia.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''S'' then ');
  qrContingencia.SQL.Add('      ''CONTINGENCIA ENVIADA''');
  qrContingencia.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''N'' and nf.status <> ''E'' then');
  qrContingencia.SQL.Add('      ''CONTINGENCIA PENDENTE''');
  qrContingencia.SQL.Add('    when nf.contingencia = ''S'' and nf.enviadocontingencia = ''N'' and nf.status = ''E'' then');
  qrContingencia.SQL.Add('      ''CONTINGENCIA COM ERRO''');
  qrContingencia.SQL.Add('    when nf.contingencia = ''N'' and nf.situacao = 0 then ');
  qrContingencia.SQL.Add('      ''ENVIO NORMAL'' ');
  qrContingencia.SQL.Add('    when nf.contingencia = ''N'' and nf.situacao = 1 then ');
  qrContingencia.SQL.Add('      ''CANCELADA''');
  qrContingencia.SQL.Add('  end status_descricao, '' '' acao');
  qrContingencia.SQL.Add('from NFCE nf where 1=1');
  if dataini.Date > 0 then
    qrContingencia.SQL.Add(' and nf.data >= ' + QuotedStr(FormatDateTime('dd.mm.yyyy',dataini.Date)));
  if datafin.Date > 0 then
    qrContingencia.SQL.Add(' and nf.data <= ' + QuotedStr(FormatDateTime('dd.mm.yyyy',datafin.Date)));
  if edPendencias.Checked then
    qrContingencia.SQL.Add(' and nf.contingencia = ''S'' and nf.enviadocontingencia = ''N''');
  qrContingencia.SQL.Add('order by data,numero');
  qrContingencia.Open;
end;

procedure TfrmEnviaContingencia.datafinChange(Sender: TObject);
begin
  edHoje.Checked := (dataini.Date = Date) and (datafin.Date = Date);
end;

procedure TfrmEnviaContingencia.datainiChange(Sender: TObject);
begin
  edHoje.Checked := (dataini.Date = Date) and (datafin.Date = Date);
end;

procedure TfrmEnviaContingencia.edHojeClick(Sender: TObject);
begin
  if edHoje.Checked then begin
    dataini.Date := Date;
    datafin.Date := Date;
  end;
end;

procedure TfrmEnviaContingencia.FormShow(Sender: TObject);
begin
  dataini.Date := Date;
  datafin.Date := Date;
  cxButton3.Click;
end;

end.
