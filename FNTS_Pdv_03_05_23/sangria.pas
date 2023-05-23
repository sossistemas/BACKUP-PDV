unit sangria;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Mask, RzEdit, Menus, AdvMenus, ComCtrls,
  XPMan, AdvGlowButton, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel,
  principal, frxClass, frxDBSet, frxDesgn, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, cxTextEdit, dxSkinsCore, dxSkinsDefaultPainters;

type
  TfrmSangria = class(TForm)
    Panel1: TPanel;
    bt_ok1: TButton;
    bt_cancelar1: TButton;
    Panel2: TPanel;
    Label1: TLabel;
    ed_valor: TRzNumericEdit;
    pop_fechamento: TAdvPopupMenu;
    Cancelar1: TMenuItem;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    fxSangria: TfrxReport;
    frxEmitente: TfrxDBDataset;
    frxDesigner1: TfrxDesigner;
    ActionManager1: TActionManager;
    Action1: TAction;
    lbEdicao: TLabel;
    edObs: TcxTextEdit;
    Label3: TLabel;
    AdvGlowButton1: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bt_cancelar1Click(Sender: TObject);
    procedure bt_ok1Click(Sender: TObject);
    procedure Cancelar1Click(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure fxSangriaBeforePrint(Sender: TfrxReportComponent);
    procedure FormCreate(Sender: TObject);
    procedure Action1Execute(Sender: TObject);
    procedure ed_valorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edObsKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    TipoImp: TImpressora;
    Editar:Boolean;
  public
    { Public declarations }
  end;

var
  frmSangria: TfrmSangria;

implementation

uses modulo, funcoes, UFuncoes, venda;

{$R *.dfm}

procedure TfrmSangria.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmPrincipal.TipoImpressora := TipoImp;
  action := cafree;
end;

procedure TfrmSangria.FormCreate(Sender: TObject);
begin
  Editar := False;
  Height := 213;
end;

procedure TfrmSangria.FormShow(Sender: TObject);
begin
  TipoImp := frmPrincipal.TipoImpressora;
  frmPrincipal.TipoImpressora := NaoFiscal;
end;

procedure TfrmSangria.fxSangriaBeforePrint(Sender: TfrxReportComponent);
var
  H, pini, idx, L: Integer;
  txtlinha, txtobs, palavra, Buffer: String;
begin
  if Sender is TfrxMemoView then begin
    if TfrxMemoView(Sender).Name = 'mCaixa' then
      TfrxMemoView(Sender).Text := 'Caixa:'+Zerar(IntToStr(iNumCaixa),3);
    if TfrxMemoView(Sender).Name = 'mValor' then
      TfrxMemoView(Sender).Text := FormatFloat('#,##0.00',ed_valor.Value);
    if TfrxMemoView(Sender).Name = 'mOperador' then
      TfrxMemoView(Sender).Text := 'Operador: ' + sNome_Operador;
    if TfrxMemoView(Sender).Name = 'mmObservacao' then
    begin
      if Trim(edObs.Text) = '' then
      begin
        TfrxMemoView(Sender).Visible := False;
      end
      else
      begin
        Buffer := 'Observação: ' + edObs.Text;
        idx := 1;
        L := 1;
        while Buffer[idx] <> #0 do
        begin
          if Buffer[idx] = #9 then
            Buffer[idx] := ' ';

          if (Buffer[idx] = ' ') then
          begin
            palavra := palavra + Buffer[idx];
            Canvas.Font := TfrxMemoView(Sender).Font;
            if Canvas.TextWidth(txtlinha + palavra) > 270 then
            begin
              txtobs := txtobs + txtlinha + #13#10;
              txtlinha := '                      ';
              Inc(L);
            end
            else
            begin
              txtlinha := txtlinha + palavra;
              palavra := '';
            end;
          end
          else
            palavra := palavra + Buffer[idx];
          Inc(idx);
        end;
        TfrxMemoView(Sender).Text := txtobs + txtlinha;
      end;
    end;
  end;

end;

procedure TfrmSangria.Action1Execute(Sender: TObject);
begin
  if Editar then
    Editar := False
  else
    Editar := True;
  lbEdicao.Visible := Editar;
end;

procedure TfrmSangria.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmSangria.bt_cancelar1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmSangria.bt_ok1Click(Sender: TObject);
begin
  if ed_valor.value <= 0 then begin
    Application.MessageBox('Favor informar um valor MAIOR que ZERO!','Erro',mb_ok+mb_iconerror);
    ed_valor.setfocus;
    exit;
  end;

  sMsg := 'OK';
  // extrair o numero dos contadores
  sGNF := zerar(FloatToStr(Max('')),5  );
  if Length(sGNF) = 5 then
    sNumero_cupom := Zerar( FloatToStr( Max('') ),5 );
  with frmModulo do begin
    qrGravaNaoFiscal.close;
    qrGravaNaoFiscal.Open;
    qrGravaNaoFiscal.Insert;
    qrGravaNaoFiscalcodigo.asstring := codifica_cupom;
    qrGravaNaoFiscalecf.AsString := sCaixa;
    qrGravaNaoFiscaldata.asdatetime := dData_Sistema;
    qrGravaNaoFiscalindice.asstring := sIndice_Sangria;
    qrGravaNaoFiscaldescricao.asstring := 'SANGRIA';
    qrGravaNaoFiscalvalor.asfloat := ed_valor.value;
    qrGravaNaoFiscalhora.AsDateTime := Time;
    qrGravaNaoFiscalCOO.asstring := sNumero_Cupom;
    qrGravaNaoFiscalGNF.asstring := sGNF;
    qrGravaNaoFiscalCDC.Clear;
    qrGravaNaoFiscalGRG.clear;
    qrGravaNaoFiscalDENOMINACAO.asstring := 'CN';
    qrGravaNaoFiscalcodVendedor.asstring := IntToStr(icodigo_Usuario);
    qrGravaNaoFiscalOBS.AsString := edObs.Text;
    qrGravaNaoFiscal.Post;
  end;

  if TipoDeCupom = tcPafECF then begin
    frmVenda.cpACBrECF.Sangria(ed_valor.value, '');
  end else begin
    fxSangria.LoadFromFile(ExtractFilePath(application.ExeName) + '\rel\F000001.fr3');
    if Editar then
      fxSangria.DesignReport
    else
      fxSangria.ShowReport;
  end;

  close;
end;

procedure TfrmSangria.Cancelar1Click(Sender: TObject);
begin
   close;
end;


procedure TfrmSangria.edObsKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if key  = 13 then
    if ed_valor.Value > 0 then
      AdvGlowButton2.Click
    else
    begin
      ed_valor.SetFocus;
      ed_valor.SelectAll;
    end;
end;

procedure TfrmSangria.ed_valorKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = 13) and (ed_valor.Value > 0) then
    edObs.SetFocus;
end;

end.
