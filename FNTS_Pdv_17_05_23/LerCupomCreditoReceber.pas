unit LerCupomCreditoReceber;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, Vcl.StdCtrls, cxTextEdit,
  Vcl.ExtCtrls, Data.DB, MemDS, DBAccess, Uni, System.Math;

type
  TfrmLerCupomCreditoReceber = class(TForm)
    Panel1: TPanel;
    edCupom: TcxTextEdit;
    Label1: TLabel;
    qrCupom: TUniQuery;
    qrCupomID: TIntegerField;
    qrCupomVENCIMENTO: TDateField;
    qrCupomVALOR: TFloatField;
    qrCupomVALOR_UTILIZADO: TFloatField;
    qrCupomVALOR_SALDO: TFloatField;
    qrCupomCUPOM_UTILIZADO: TStringField;
    procedure edCupomKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLerCupomCreditoReceber: TfrmLerCupomCreditoReceber;

implementation

{$R *.dfm}

uses Receber_Valor, venda;

procedure TfrmLerCupomCreditoReceber.edCupomKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then begin
    qrCupom.Close;
    qrCupom.Params.Items[0].Value := edCupom.Text;
    qrCupom.Open;
    if qrCupom.IsEmpty then begin
      Application.MessageBox('Cupom não Localizado!','Atenção!',MB_ICONINFORMATION);
      edCupom.SetFocus;
      exit;
    end else begin
      if qrCupomVENCIMENTO.AsDateTime > Date then begin
        Application.MessageBox('Cupom Vencido!','Atenção!',MB_ICONINFORMATION);
        edCupom.SetFocus;
        exit;
      end;
      if qrCupomCUPOM_UTILIZADO.AsString = 'S' then begin
        Application.MessageBox('Cupom já Utilizado!','Atenção!',MB_ICONINFORMATION);
        edCupom.SetFocus;
        exit;
      end;
      if RoundTo(qrCupomVALOR_SALDO.AsFloat,-2) > RoundTo(frmReceber_Valor.ValorMaximo,-2) then
        frmReceber_Valor.edValor.Value := RoundTo(frmReceber_Valor.ValorMaximo,-2)
      else
        frmReceber_Valor.edValor.Value := RoundTo(qrCupomVALOR_SALDO.AsFloat,-2);
      RecebeCupomCredito.id := qrCupomID.AsInteger;
      RecebeCupomCredito.Cupom := edCupom.Text;
      RecebeCupomCredito.Valor := RoundTo(frmReceber_Valor.edValor.Value,-2);
      Close;
      frmReceber_Valor.Action2Execute(frmReceber_Valor.AdvGlowButton1);
    end;
  end else if Key = #27 then
    Close;
end;

end.
