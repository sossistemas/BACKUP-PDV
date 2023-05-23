unit Comanda;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, MemDS, DBAccess, Vcl.Mask, RzEdit, AdvMetroButton,
  AdvSmoothPanel, AdvSmoothExpanderPanel, Uni, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxSpinEdit;

type
  TFComanda = class(TForm)
    Label1: TLabel;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    QRITEM: TUniQuery;
    edComanda: TcxSpinEdit;
    qrtmp: TUniQuery;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure edComandaKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure FechaComanda;
  end;

var
  FComanda: TFComanda;

implementation

uses
  modulo, principal, funcoes;

{$R *.dfm}

procedure TFComanda.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
    Close;
end;

procedure TFComanda.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TFComanda.edComandaKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then begin
    FechaComanda;
  end;
end;

procedure TFComanda.FechaComanda;
begin

  if edComanda.Value <= 0 then begin
    application.MessageBox('Favor digitar o numero da comanda', 'Atenção', MB_OK + MB_ICONWARNING);
    exit;
  end;

  frmmodulo.query_servidor.close;
  frmmodulo.query_servidor.sql.clear;
  frmmodulo.query_servidor.SQL.add('select r000001.codigo, r000001.PERGUNTOU_COMISSAO,r000001.CLIENTE_AUTORIZOU, r000002.*, C000025.produto, C000025.unidade,');
  frmmodulo.query_servidor.sql.add('c000025.cst, c000025.aliquota, c000025.codbarra');
  frmmodulo.query_servidor.sql.add('from r000001, r000002, c000025');
  frmmodulo.query_servidor.sql.add('where r000001.codigo = r000002.cod_mesa');
  frmmodulo.query_servidor.sql.add('and r000002.cod_produto = c000025.codigo');
  frmmodulo.query_servidor.sql.add('and r000001.comanda = ' + IntToStr(edComanda.Value));
  frmmodulo.query_servidor.sql.add('order by r000002.codigo');
  frmmodulo.query_servidor.open;


  if (frmmodulo.query_servidor.RecordCount = 0) then begin
    MessageBeep(32);
    Application.messagebox('Comanda sem consumação!', 'Atenção!', mb_ok);
  end
  else begin
    if frmModulo.qrconfigCOBRA_COMISSAO.AsString = 'P' then begin
      if frmmodulo.query_servidor.FieldByName('PERGUNTOU_COMISSAO').AsString = 'N' then begin
        qrtmp.SQL.Clear;
        if Application.MessageBox('Cliente autorizou incluir taxa do garçon na comanda?','Atenção?',MB_ICONQUESTION+MB_YESNO) =mrYes then
          qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('S') + ' where CODIGO = ' + frmmodulo.query_servidor.FieldByName('codigo').AsString)
        else
          qrtmp.SQL.Add('update r000001 set PERGUNTOU_COMISSAO = ' + QuotedStr('S') + ',cliente_autorizou = ' + QuotedStr('N') + ' where CODIGO = ' + frmmodulo.query_servidor.FieldByName('codigo').AsString);
        qrtmp.ExecSQL;
      end;
    end;
    bLanca_Comanda := true;
    MesasComandasReceber := IntToStr(frmmodulo.query_servidor.FieldByName('codigo').AsInteger);
    close;
  end;
end;

end.

