unit Receber_Valor;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, Vcl.StdCtrls, cxTextEdit,
  System.Actions, Vcl.ActnList, AdvGlowButton, Vcl.ExtCtrls, Vcl.Touch.Keyboard,
  AdvTouchKeyboard, AdvSmoothTouchKeyBoard, cxMaskEdit, cxDropDownEdit, cxCalc,
  Vcl.Mask, RzEdit, Vcl.Menus, cxButtons, dxSkinsCore, dxSkinsDefaultPainters;

type
  tValorReceber = record
    Confirma:Boolean;
    Valor:Double;
  end;
  tTipoMovimento = (tmDinheiro, tmCrediario, tmConvenio, tmCartaoCredito,
      tmCartaoDebito, tmCheque, tmTEF, tmPix, tmDesconto, tmAcrescimo);
  tTipoValor = (tvRecebimento, tvDesconto, tvAcrescimo);
  TfrmReceber_Valor = class(TForm)
    Panel1: TPanel;
    AdvGlowButton1: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    Panel2: TPanel;
    ActionList1: TActionList;
    Action1: TAction;
    Action2: TAction;
    Label1: TLabel;
    pnTeclado: TPanel;
    Teclado: TAdvSmoothTouchKeyBoard;
    lbTotal: TLabel;
    edPercentual: TCheckBox;
    lbDesAcre: TLabel;
    edValor: TRzNumericEdit;
    btnCupom: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Action1Execute(Sender: TObject);
    procedure Action2Execute(Sender: TObject);
    procedure edValorChange(Sender: TObject);
    procedure edValorKeyPress(Sender: TObject; var Key: Char);
    procedure edValorEnter(Sender: TObject);
    procedure edPercentualClick(Sender: TObject);
    procedure btnCupomClick(Sender: TObject);
  private
    { Private declarations }
    vContinua:Boolean;
    RecebeMais:Boolean;
    ValorCalculado:Double;
    vTipo:tTipoValor;
    FTrocoMaximo: Currency;
  public
    { Public declarations }
    ValorMaximo:Double;

    property TrocoMaximo: Currency read FTrocoMaximo write FTrocoMaximo;
  end;
  function RetornaValorRecebimento(ValorTotal:Double;RecebimentoMaior:Boolean;Tipo:tTipoMovimento;TipoValor:tTipoValor):tValorReceber;

var
  frmReceber_Valor: TfrmReceber_Valor;

implementation

uses principal, System.Math, LerCupomCreditoReceber;

{$R *.dfm}

function RetornaValorRecebimento(ValorTotal:Double;RecebimentoMaior:Boolean;
   Tipo:tTipoMovimento;TipoValor:tTipoValor):tValorReceber;
begin
  if not Assigned(frmReceber_Valor) then
    Application.CreateForm(TfrmReceber_Valor,frmReceber_Valor);

  try
    with frmReceber_Valor do
    begin
      if Tipo <> tmDinheiro then
      begin
        btnCupom.Visible := False;
        edValor.Left := Trunc((Panel2.Width - edValor.Width)/2);
        Label1.Left := edValor.Left;
      end;
      lbTotal.Caption := 'Total R$ ' + FormatFloat('#,##0.00',ValorTotal);
      ValorMaximo := ValorTotal;
      RecebeMais := RecebimentoMaior;
      ValorCalculado := 0;
      vTipo := TipoValor;
      edValor.DisplayFormat := 'R$ #,##0.00';

      if TipoValor = tvRecebimento then
        edValor.Value := ValorTotal
      else
      begin
        edPercentual.Visible := True;
        lbDesAcre.Visible := True;
        if TipoValor = tvAcrescimo then
          lbDesAcre.Caption := 'Acréscimo R$ 0.00'
        else
          lbDesAcre.Caption := 'Desconto R$ 0.00';
      end;
      vContinua := False;
      ShowModal;
      if vContinua then
        if vTipo = tvRecebimento then
          Result.Valor := edValor.Value
        else
          Result.Valor := ValorCalculado
      else
        Result.Valor := 0;
      Result.Confirma := vContinua;
    end;
  finally
    FreeAndNil(frmReceber_Valor);
  end;
end;


procedure TfrmReceber_Valor.Action1Execute(Sender: TObject);
begin
  vContinua := False;
  Close;
end;

procedure TfrmReceber_Valor.Action2Execute(Sender: TObject);
var
  calculaTroco : Currency;
begin
  if edValor.Value <=0 then begin
    Application.MessageBox('Informe o valor recebido!','Atenção!',MB_ICONINFORMATION);
    edValor.SetFocus;
    Exit;
  end;
  if not RecebeMais then
  begin
    if RoundTo(edValor.Value,-2) > RoundTo(ValorMaximo,-2) then
    begin
      Application.MessageBox('Valor informado é superior ao saldo a receber do cupom!','Atenção!',MB_ICONINFORMATION);
      edValor.SetFocus;
      Exit;
    end;
  end
  else
  begin
    calculaTroco := (RoundTo(edValor.Value,-2) - RoundTo(ValorMaximo,-2));

    if calculaTroco > FTrocoMaximo then
    begin
       Application.MessageBox('Valor informádo gera um troco maior que o permitido!','Atenção!',MB_ICONINFORMATION);
      edValor.SetFocus;
      Exit;
    end;
  end;
  vContinua := True;
  Close;
end;

procedure TfrmReceber_Valor.btnCupomClick(Sender: TObject);
begin
  Application.CreateForm(TfrmLerCupomCreditoReceber,frmLerCupomCreditoReceber);
  frmLerCupomCreditoReceber.ShowModal;
  FreeAndNil(frmLerCupomCreditoReceber);
end;

procedure TfrmReceber_Valor.edPercentualClick(Sender: TObject);
begin
  if edPercentual.Checked then
    edValor.DisplayFormat := '% #,##0.00'
  else
    edValor.DisplayFormat := 'R$ #,##0.00';

end;

procedure TfrmReceber_Valor.edValorChange(Sender: TObject);
begin
  if vTipo = tvDesconto then begin
    if edPercentual.Checked then
      ValorCalculado := RoundTo(((ValorMaximo * edValor.Value)/100),-2)
    else
      ValorCalculado := RoundTo(edValor.Value,-2);
    lbDesAcre.Caption := 'Desconto R$ ' + FormatFloat('#,##0.00', ValorCalculado);
  end else if vTipo = tvAcrescimo then begin
    if edPercentual.Checked then
      ValorCalculado := RoundTo(((ValorMaximo * edValor.Value)/100),-2)
    else
      ValorCalculado := RoundTo(edValor.Value,-2);
    lbDesAcre.Caption := 'Acréscimo R$ ' + FormatFloat('#,##0.00', ValorCalculado);
  end;
end;

procedure TfrmReceber_Valor.edValorEnter(Sender: TObject);
begin
  edValor.SelectAll;
end;

procedure TfrmReceber_Valor.edValorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    Action2Execute(Sender);
end;

procedure TfrmReceber_Valor.FormCreate(Sender: TObject);
begin
  if not(AtivaTouch) or not(ExibeTecladoVirtual) then begin
    ClientHeight := 300;
    ClientWidth := 365;
    pnTeclado.Visible := False;
    Teclado.Update;
  end;
end;

procedure TfrmReceber_Valor.FormShow(Sender: TObject);
begin
  edValor.SelectAll;
  edValor.SetFocus;
  Update;
end;

end.
