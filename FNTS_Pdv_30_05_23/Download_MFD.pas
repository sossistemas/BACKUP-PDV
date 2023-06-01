unit Download_MFD;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Mask, XPMan, AdvGlowButton,
  AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel, JvExMask, JvToolEdit;

type
  TfrmDownload_MFD = class(TForm)
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    rb_data: TRadioButton;
    rb_coo: TRadioButton;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    ed_inicio: TEdit;
    ed_fim: TEdit;
    ed_data_ini: TJvDateEdit;
    ed_data_fim: TJvDateEdit;
    GroupBox3: TGroupBox;
    rb_espelho: TRadioButton;
    rb_arquivo: TRadioButton;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    bt_ok: TAdvGlowButton;
    bt_cancelar: TAdvGlowButton;
    mResp: TMemo;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bt_cancelar1Click(Sender: TObject);
    procedure bt_ok1Click(Sender: TObject);
    procedure rb_dataClick(Sender: TObject);
    procedure rb_cooClick(Sender: TObject);
    procedure ed_data_iniKeyPress(Sender: TObject; var Key: Char);
    procedure ed_data_fimKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure rb_dataKeyPress(Sender: TObject; var Key: Char);
    procedure rb_espelhoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDownload_MFD: TfrmDownload_MFD;

implementation

  uses funcoes, principal, msg_Operador, UFuncoes, venda;

{$R *.dfm}

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.bt_cancelar1Click(Sender: TObject);
begin
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.bt_ok1Click(Sender: TObject);
var
  Linhas: TStringList;
  I:Integer;
  NomeArquivo:string;
begin
  frmMsg_operador.lb_msg.Caption := 'Aguarde! Emitindo Leitura MF...';
  frmMsg_operador.Show;
  frmMsg_operador.Refresh;

  Linhas := TStringList.Create;
  try
    if rb_espelho.checked then begin
      if rb_data.checked then begin
        NomeArquivo := ECFCaminhoFiscal+'\Espelho_MFD.txt';
        frmVenda.cpACBrECF.EspelhoMFD_DLL(ed_data_ini.Date, ed_data_fim.Date, NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end else begin
        NomeArquivo := ECFCaminhoFiscal+'\Espelho_MFD.txt';
        frmVenda.cpACBrECF.EspelhoMFD_DLL(StrToInt(ed_inicio.Text), StrToInt(ed_fim.Text), NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end;
    end else begin
      if rb_data.checked then begin
        NomeArquivo := ECFCaminhoFiscal+'\Espelho_MFD.txt';
        frmVenda.cpACBrECF.ArquivoMFD_DLL(ed_data_ini.Date, ed_data_fim.Date, NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end else begin
        NomeArquivo := ECFCaminhoFiscal+'\Espelho_MFD.txt';
        frmVenda.cpACBrECF.ArquivoMFD_DLL(StrToInt(ed_inicio.Text), StrToInt(ed_fim.Text), NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end;
    end;
  finally
    Linhas.Free;
  end;
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.rb_dataClick(Sender: TObject);
begin
  ed_data_ini.Visible := true;
  ed_data_fim.Visible := true;
  ed_inicio.Visible   := false;
  ed_fim.Visible      := false;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.rb_cooClick(Sender: TObject);
begin
  ed_data_ini.Visible := false;
  ed_data_fim.Visible := false;
  ed_inicio.Visible   := true;
  ed_fim.Visible      := true;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.ed_data_iniKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then perform(wm_nextdlgctl,0,0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.ed_data_fimKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then bt_ok.setfocus;
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.FormShow(Sender: TObject);
begin
  if rb_espelho.Checked then
    caption := 'Espelho da MFD'
  else
    caption := 'Arquivo da MFD';

  ed_data_ini.Date := date;
  ed_data_fim.Date := date;
  ed_inicio.Text :='0';
  ed_fim.text := '0';
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.rb_dataKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    perform(wm_nextdlgctl,0,0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmDownload_MFD.rb_espelhoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    perform(wm_nextdlgctl,0,0);

end;

end.
