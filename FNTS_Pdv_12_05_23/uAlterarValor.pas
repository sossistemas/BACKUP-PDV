unit uAlterarValor;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, SOSSolution.Controls, AdvGlowButton, Vcl.StdCtrls, Vcl.Mask, RzEdit, RzLabel, Data.DB, MemDS,
  DBAccess, Uni;

type
  TAcao =
  (aValorUnitario
  ,aQuantidadeVenda);

  TfrmAlterarValorUnitario = class(TForm)
    lblNomeCampo: TRzLabel;
    edtEntrada: TRzNumericEdit;
    bgravar: TAdvGlowButton;
    bcancelar: TAdvGlowButton;
    RoundShape1: TRoundShape;
    RoundShape2: TRoundShape;
    RoundShape3: TRoundShape;
    procedure FormCreate(Sender: TObject);
    procedure edtEntradaEnter(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure bgravarClick(Sender: TObject);
  private
    FAcao: TAcao;
  public
    constructor Create(AOwner: TComponent; Acao: TAcao); overload;
  end;

  procedure AlterarValorUnitario(ANovoValorUnitario: Double; AComponent: TComponent);
var
  frmAlterarValorUnitario: TfrmAlterarValorUnitario;



implementation

uses
  principal, venda, modulo, Math;

{$R *.dfm}


procedure AlterarValorUnitario(ANovoValorUnitario: Double; AComponent: TComponent);
var
 lNovoTotal,
 lCustoVariavel: Double;
 lModoEdicaoAtivo: Boolean;
 ///
 function CreateQuery: TUniQuery;
 begin
   try
     Result := TUniQuery.Create(AComponent);
     Result.Connection := frmModulo.conexao;
     Result.Close;
     Result.SQL.Clear;
     Result.SQL.Text := 'SELECT * FROM CUPOM_ITEM_TEMP WHERE CODIGO = ' + QuotedStr(frmModulo.qrItensCODIGO.Value);
     Result.Open;
     if Result.RecordCount = 0 then
       Abort;
     Result.Edit;
   except
     Abort;
   end;
 end;
begin
  with CreateQuery, frmModulo, Application do
  begin
//    if FAcao = aValorUnitario then
//    begin
      if ANovoValorUnitario <= 0 then
      begin
        MessageBox(PWideChar('Não é permitido venda sem valor unitário!'),'Atenção!',MB_ICONINFORMATION);
        Exit;
      end
      else
      begin
        lModoEdicaoAtivo := qrItens.state = dsEdit;
        if not lModoEdicaoAtivo then
          qrItens.Edit;
        ///
        lCustoVariavel := RoundTo((qrItensVALOR_UNITARIO.Value - ANovoValorUnitario) * qrItensQTDE.Value , -2);
        qrItensVALOR_UNITARIO.Value := ANovoValorUnitario;
        lNovoTotal := RoundTo(qrItensVALOR_UNITARIO.Value * qrItensQTDE.Value, -2);
        rTotal_Venda := rTotal_Venda - (qrItensVALOR_TOTAL.Value - lNovoTotal);

        ///
//        if lCustoVariavel > 0 then
//        begin
//          rTotal_Desconto := rTotal_Desconto - (qrItensVALOR_DESCONTO.Value - lCustoVariavel);
//          qrItensVALOR_DESCONTO.Value := lCustoVariavel;
//          if qrItensVALOR_ACRESCIMO.Value > 0 then
//          begin
//            rTotal_Acrescimo := rTotal_Acrescimo - qrItensVALOR_ACRESCIMO.Value;
//            qrItensVALOR_ACRESCIMO.Value := 0;
//          end;
//        end
//        else
//        begin
//          lCustoVariavel := Abs(lCustoVariavel);
//          rTotal_Acrescimo := rTotal_Acrescimo - (qrItensVALOR_ACRESCIMO.Value - lCustoVariavel);
//          qrItensVALOR_ACRESCIMO.Value := Abs(lCustoVariavel);
//          if qrItensVALOR_DESCONTO.Value > 0 then
//          begin
//            rTotal_Desconto := rTotal_Desconto - qrItensVALOR_DESCONTO.Value;
//            qrItensVALOR_DESCONTO.Value := 0;
//          end;
//        end;
        ///
        qrItensVALOR_TOTAL.Value := lNovoTotal;
        ///
//        FieldByName('valor_desconto').AsFloat := qrItensVALOR_DESCONTO.Value;
//        FieldByName('valor_acrescimo').AsFloat := qrItensVALOR_ACRESCIMO.Value;
        ///
        if not lModoEdicaoAtivo then
          qrItens.Post;
      end;
//    else
//    if edtEntrada.Value = qrItensQTDE.Value then
//      Self.Close
//    else
//    if edtEntrada.Value <= 0 then
//    begin
//      MessageBox(PWideChar('Não é permitido venda sem quantidade especificada!'),'Atenção!',MB_ICONINFORMATION);
//      Exit;
//    end
//    else
//    begin
//      rProd_preco := PrecoDeVenda(qrItensCOD_PRODUTO.Value, qrItensQTDE.Value);
//      rProd_qtde := edtEntrada.Value;
//      if rProd_preco <> qrItensVALOR_UNITARIO.Value then
//      begin
//        FieldByName('valor_desconto').AsFloat := 0;
//        FieldByName('valor_acrescimo').AsFloat := 0;
//      end
//      else
//      begin
//        if qrItensVALOR_DESCONTO.Value > 0 then
//        begin
//          lCustoVariavel :=  qrItensVALOR_DESCONTO.Value / qrItensQTDE.Value;
//          if MessageBox(PWideChar('Este produto esta com desconto de R$' + FormatFloat(',.00', lCustoVariavel) +
//                                              ' por unidade.'#13'Manter esse desconto para a nova quantidade?')
//                                   ,'Atenção!'
//                                   ,MB_ICONQUESTION) = IDYES
//          then
//          begin
//            lCustoVariavel := RoundTo(lCustoVariavel * rProd_qtde , -2);
//            lNovoTotal := RoundTo(rProd_preco * rProd_qtde - lCustoVariavel , -2);
//            rTotal_Venda := rTotal_Venda - (qrItensVALOR_TOTAL.Value - lNovoTotal);
//          end;
//        end;
//      end;
//      rProd_total := RoundTo(rProd_qtde * rProd_preco, -2);
//      rTotal_Venda := rTotal_Venda - (qrItensVALOR_TOTAL.Value - rProd_total);
//      FieldByName('valor_unitario').AsFloat := rProd_preco;
//      FieldByName('valor_subtotal').AsFloat := rProd_total;
//    end;
//    ///
    FieldByName('valor_total').AsFloat := lNovoTotal;
    Post;
  end;
end;


procedure MakeRounded(Control: TWinControl);
var
  R: TRect;
  Rgn: HRGN;
begin
  with Control do
  begin
    R := ClientRect;
    rgn := CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, 20, 20);
    Perform(EM_GETRECT, 0, lParam(@r));
    InflateRect(r, - 5, - 5);
    Perform(EM_SETRECTNP, 0, lParam(@r));
    SetWindowRgn(Handle, rgn, True);
    Invalidate;
  end;
end;


procedure TfrmAlterarValorUnitario.bgravarClick(Sender: TObject);
begin
  AlterarValorUnitario(edtEntrada.Value, Self);
  Close;
end;

constructor TfrmAlterarValorUnitario.Create(AOwner: TComponent; Acao: TAcao);
begin
  FAcao := Acao;
  inherited Create(AOwner);
end;

procedure TfrmAlterarValorUnitario.edtEntradaEnter(Sender: TObject);
begin
  edtEntrada.DisplayFormat := '###,###,##0.000';
  edtEntrada.SelectAll;
end;

procedure TfrmAlterarValorUnitario.FormCreate(Sender: TObject);
begin
  if FAcao = aValorUnitario then
  begin
    lblNomeCampo.Caption := 'Valor Unitário';
    edtEntrada.DisplayFormat := 'R$ ###,###,##0.000';
  end
  else
  begin
    lblNomeCampo.Caption := 'Quantidade';
    edtEntrada.DisplayFormat := '###,###,##0.000';
  end;

  edtEntrada.Value := frmModulo.qrItensVALOR_UNITARIO.Value;
  MakeRounded(bgravar);
  MakeRounded(bcancelar);
  MakeRounded(Self);
end;

procedure TfrmAlterarValorUnitario.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
    bgravarClick(nil)
  else
  if Key = VK_ESCAPE then
    Close;
end;

end.
