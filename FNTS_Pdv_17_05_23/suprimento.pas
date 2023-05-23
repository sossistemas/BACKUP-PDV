unit suprimento;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Mask, RzEdit, Menus, AdvMenus, ComCtrls,
  XPMan, AdvGlowButton, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel,
  principal, frxClass, frxDBSet, frxDesgn, frxExportPDF, System.Actions,
  Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, cxGraphics,
  cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit,
  cxTextEdit, system.IniFiles, dxSkinsCore, dxSkinsDefaultPainters, frxExportBaseDialog;

type
  TfrmSuprimento = class(TForm)
    Panel1: TPanel;
    bt_ok1: TButton;
    bt_cancelar: TButton;
    Panel2: TPanel;
    Label1: TLabel;
    ed_valor: TRzNumericEdit;
    pop_fechamento: TAdvPopupMenu;
    Cancelar1: TMenuItem;
    list_Forma: TListBox;
    Label2: TLabel;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    bt_ok: TAdvGlowButton;
    frxDesigner1: TfrxDesigner;
    fxSangria: TfrxReport;
    frxEmitente: TfrxDBDataset;
    frxPDFExport1: TfrxPDFExport;
    lbEdicao: TLabel;
    ActionManager1: TActionManager;
    Action1: TAction;
    Shape1: TShape;
    edObs: TcxTextEdit;
    Label3: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bt_cancelarClick(Sender: TObject);
    procedure bt_ok1Click(Sender: TObject);
    procedure Cancelar1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure list_FormaKeyPress(Sender: TObject; var Key: Char);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure fxSangriaBeforePrint(Sender: TfrxReportComponent);
    procedure Action1Execute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ed_valorKeyPress(Sender: TObject; var Key: Char);
    procedure edObsKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    TipoImp: TImpressora;
    Editar:Boolean;
  public
    { Public declarations }
  end;

var
  frmSuprimento: TfrmSuprimento;

implementation

uses modulo, funcoes, UFuncoes, venda;

{$R *.dfm}

procedure TfrmSuprimento.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmPrincipal.TipoImpressora := TipoImp;
  action := cafree;
end;

procedure TfrmSuprimento.FormCreate(Sender: TObject);
begin
  Editar := False;
end;

procedure TfrmSuprimento.Action1Execute(Sender: TObject);
begin
  if Editar then
    Editar := False
  else
    Editar := True;
  lbEdicao.Visible := Editar;
end;

procedure TfrmSuprimento.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmSuprimento.bt_cancelarClick(Sender: TObject);
begin
  close;
end;

procedure TfrmSuprimento.bt_ok1Click(Sender: TObject);
var
  IniFile: string;
  Ini: TIniFile;
  preview: Boolean;
  duasVias:boolean;

begin
  IniFile := sConfiguracoes;
  Ini := TIniFile.Create(IniFile);

  if ed_valor.value <= 0 then begin
    Application.MessageBox('Favor informar um valor MAIOR que ZERO!','Erro',mb_ok+mb_iconerror);
    ed_valor.setfocus;
    exit;
  end;
  if list_forma.ItemIndex < 0 then begin
    Application.MessageBox('Favor selecionar uma forma!','Erro',mb_ok+mb_iconerror);
    list_forma.setfocus;
    exit;
  end;
  sNumero_cupom := Zerar( FloatToStr( Max('') ),5 );
  if TipoDeCupom = tcPafECF then
    frmVenda.cpACBrECF.Suprimento(ed_valor.value,'')
  else
  begin
    fxSangria.LoadFromFile(ExtractFilePath(application.ExeName) + '\rel\F000002.fr3');
    if Editar then
     fxSangria.DesignReport
    else
    begin

    fxSangria.ShowReport; {Remover quando adicionar opções abaixo}
    {
    preview := Ini.ReadBool('Fortes','Sangria_Suprimento_Preview', false);
    duasVias := Ini.ReadBool('Fortes','Sangria_Suprimento_DuasVias', true);

      if preview <> True then
      begin
        if duasVias = true then
        Begin
          fxSangria.PrintOptions.ShowDialog:=False;
          fxSangria.PrintOptions.Printer := sImpressora_Venda_Simples;
          fxSangria.Print;
          fxSangria.Print;
        end
        else
        begin
          fxSangria.Print;
        end;
      end
      else
      begin
        fxSangria.ShowReport;
      end;
}
    end;
  end;

  sGNF := Zerar( FloatToStr( Max('') ),5 );
  with frmModulo do begin
    qrGravaNaoFiscal.close;
    qrGravaNaoFiscal.Open;
    qrGravaNaoFiscal.Insert;
    qrGravaNaoFiscalcodigo.asstring := codifica_cupom;
    qrGravaNaoFiscalecf.AsString := sCaixa;
    qrGravaNaoFiscaldata.asdatetime := dData_Sistema;
    qrGravaNaoFiscalindice.asstring := sIndice_Suprimento;
    qrGravaNaoFiscaldescricao.asstring := 'SUPRIMENTO';
    qrGravaNaoFiscalvalor.asfloat := ed_valor.value;
    qrGravaNaoFiscalhora.AsDateTime := Time;
    qrGravaNaoFiscalCOO.asstring := sNumero_Cupom;
    qrGravaNaoFiscalGNF.asstring := sGNF;
    qrGravaNaoFiscalCDC.Clear;
    qrGravaNaoFiscalGRG.clear;
    qrGravaNaoFiscalDENOMINACAO.asstring := 'CN';
    qrGravaNaoFiscalcodvendedor.asstring := IntToStr(icodigo_Usuario);
    qrGravaNaoFiscalOBS.AsString := edObs.Text;
    qrGravaNaoFiscal.Post;

    qrGravaForma.close;
    qrGravaForma.Open;
    qrGravaForma.Insert;
    qrGravaFormacodigo.asstring     :=  codifica_forma(1);
    qrGravaFormacod_cupom.asstring  :=  codifica_cupom;
    qrGravaFormaforma.asstring      := list_forma.Items.Strings[list_Forma.ItemIndex];
    qrGravaFormavalor.asfloat       := ed_valor.value;
    qrGravaFormaprestacao.asinteger := 1;
    qrGravaFormatipo.asstring       := 'SUPRIMENTO';
    qrGravaFormaEX.AsInteger        := 0;
    qrGravaForma.Post;
  end;
  Close;
end;

procedure TfrmSuprimento.Cancelar1Click(Sender: TObject);
begin
   close;
end;

procedure TfrmSuprimento.edObsKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    bt_ok.SetFocus;
end;

procedure TfrmSuprimento.ed_valorKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edObs.setfocus;
end;

procedure TfrmSuprimento.FormShow(Sender: TObject);
begin
  TipoImp := frmPrincipal.TipoImpressora;
  frmPrincipal.TipoImpressora := NaoFiscal;

  list_forma.Clear;
  frmmodulo.tbForma_Pgto.First;
  while not frmModulo.tbForma_Pgto.Eof do begin
    sMsg := TrimLeft(TrimRight(frmmodulo.tbForma_Pgto.fieldbyname('nome').asstring));

    if not(sMsg = EmptyStr) and ((ansiuppercase(sMsg) <> 'VALOR RECEBIDO') and (ansiuppercase(sMsg) <> 'TROCO')) then
      list_Forma.Items.Add(sMsg);

    frmModulo.tbForma_Pgto.Next;
  end;
end;

procedure TfrmSuprimento.fxSangriaBeforePrint(Sender: TfrxReportComponent);
begin
  if Sender is TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'mCaixa' then
      TfrxMemoView(Sender).Text := 'Caixa:'+Zerar(IntToStr(iNumCaixa),3);
    if TfrxMemoView(Sender).Name = 'mValor' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00',ed_valor.Value);
    if TfrxMemoView(Sender).Name = 'mOperador' then
      TfrxMemoView(Sender).Text := 'Operador: ' + sNome_Operador;
    if TfrxMemoView(Sender).Name = 'mForma' then
      TfrxMemoView(Sender).Text := 'Forma: ' + list_forma.Items.Strings[list_Forma.ItemIndex];

    if (TfrxMemoView(Sender).Name = 'mObs') and (frmSuprimento.edObs.Text <> EmptyStr) then
    begin
      TfrxMemoView(Sender).Visible := True;
      TfrxMemoView(Sender).Text    := 'Obs.: ' + frmSuprimento.edObs.Text;
    end
    else if (TfrxMemoView(Sender).Name = 'mObs') and (frmSuprimento.edObs.Text = EmptyStr) then
    begin
      TfrxMemoView(Sender).Visible := False;
      TfrxMemoView(Sender).Text := EmptyStr;
    end;

  end;
end;

procedure TfrmSuprimento.list_FormaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then ed_valor.setfocus;
end;


end.
