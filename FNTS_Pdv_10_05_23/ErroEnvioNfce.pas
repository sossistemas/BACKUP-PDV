unit ErroEnvioNfce;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, Vcl.Menus, cxControls, cxContainer, cxEdit, cxLabel,
  Vcl.ExtCtrls, Vcl.StdCtrls, cxButtons;

type
  TRetornoErroNfce = (reTentarNovamente, reContingencia, reCancelar);
  TfrmErroEnvioNfce = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    lbMensagem: TcxLabel;
    Bevel3: TBevel;
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure FormKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
    Retorno: TRetornoErroNfce;
  public
    { Public declarations }
  end;
  function ErroNFCE(Mensagem:WideString):TRetornoErroNfce;

var
  frmErroEnvioNfce: TfrmErroEnvioNfce;

implementation

{$R *.dfm}

function ErroNFCE(Mensagem:WideString):TRetornoErroNfce;
begin
  Application.CreateForm(TfrmErroEnvioNfce, frmErroEnvioNfce);
  with frmErroEnvioNfce do begin
    Retorno := reCancelar;
    lbMensagem.Caption := Mensagem;
    ShowModal;
    Result := Retorno;
  end;
  FreeAndNil(frmErroEnvioNfce);
end;

procedure TfrmErroEnvioNfce.cxButton1Click(Sender: TObject);
begin
  Retorno:= reTentarNovamente;
  Close;
end;

procedure TfrmErroEnvioNfce.cxButton2Click(Sender: TObject);
begin
  Retorno:= reContingencia;
  Close;
end;

procedure TfrmErroEnvioNfce.cxButton3Click(Sender: TObject);
begin
  Retorno:= reCancelar;
  Close;
end;

procedure TfrmErroEnvioNfce.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F2 then
    cxButton1Click(Sender)
  else if Key = VK_F3 then
    cxButton2Click(Sender)
  else if Key = VK_F4 then
    cxButton3Click(Sender);
end;

end.


