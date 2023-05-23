unit RecebimentoEspecial;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvGlowButton, Vcl.StdCtrls,
  Vcl.ExtCtrls, ACBrUtil;

type
  TfrmRecebimentoEspecial = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    btnFechaCupom: TAdvGlowButton;
    AdvGlowButton1: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    procedure btnFechaCupomClick(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Confirma:Boolean;
  end;

var
  frmRecebimentoEspecial: TfrmRecebimentoEspecial;

implementation

{$R *.dfm}

uses venda, principal, cliente_consulta, modulo;

procedure TfrmRecebimentoEspecial.AdvGlowButton1Click(Sender: TObject);
begin
  if CodCliGeraCredito = 1 then begin
    while CodCliGeraCredito = 1 do begin
      frmCliente_consulta := TfrmCliente_consulta.create(Self);
      frmCliente_consulta.lb_cliente.Caption := 'Selecionar Cliente para Gerar Crédito!';
      bConsulta_crediario := True;
      frmCliente_consulta.showmodal;
      FreeAndNil(frmCliente_consulta);
      CodCliGeraCredito := frmModulo.Query.FieldByName('CODIGO').AsInteger;
      if CodCliGeraCredito = 1 then begin
        if Application.MessageBox(pwidechar('Cliente inválido para gerar Crédito!'+#13+'Selecionar outro cliente?'),'Atenção!',MB_ICONQUESTION+MB_YESNO) <> mrYes then begin
          Exit;
        end;
      end;
    end;
  end;
  if Application.MessageBox('Gerar um Crédito para o Cliente?','Atenção!',MB_ICONQUESTION+MB_YESNO) = mrYes then begin
    with frmVenda do begin
      cdsRecebimento.Append;
      cdsRecebimentoTipo.AsString := StRecCredito;
      cdsRecebimentovalor.AsFloat := rTotal_Venda;
      cdsRecebimentocliente.AsString := IntToStrZero(CodCliGeraCredito, 6);
      cdsRecebimento.Post;
    end;
    Confirma := True;
    Close;
    Close;
  end;
end;

procedure TfrmRecebimentoEspecial.AdvGlowButton2Click(Sender: TObject);
begin
  Confirma := False;
  Close;
end;

procedure TfrmRecebimentoEspecial.btnFechaCupomClick(Sender: TObject);
begin
  if Application.MessageBox('Devolver a Diferença para o Cliente?','Atenção!',MB_ICONQUESTION+MB_YESNO) = mrYes then begin
    with frmVenda do begin
      cdsRecebimento.Append;
      cdsRecebimentoTipo.AsString := StRecEstorno;
      cdsRecebimentovalor.AsFloat := rTotal_Venda;
      cdsRecebimento.Post;
    end;
    Confirma := True;
    Close;
  end;
end;

procedure TfrmRecebimentoEspecial.FormShow(Sender: TObject);
begin
  Confirma := False;
end;

end.
