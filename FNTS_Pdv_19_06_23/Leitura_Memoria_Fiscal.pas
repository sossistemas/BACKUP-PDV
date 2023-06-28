unit Leitura_Memoria_Fiscal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Mask, XPMan, AdvGlowButton,
  AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel, JvExMask, JvToolEdit;

type
  TfrmLeitura_Memoria_Fiscal = class(TForm)
    GroupBox1: TGroupBox;
    rb_data: TRadioButton;
    rb_crz: TRadioButton;
    Panel1: TPanel;
    GroupBox2: TGroupBox;
    ed_inicio: TEdit;
    ed_fim: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    ed_data_ini: TJvDateEdit;
    ed_data_fim: TJvDateEdit;
    GroupBox3: TGroupBox;
    rb_ecf: TRadioButton;
    rb_arquivo: TRadioButton;
    GroupBox4: TGroupBox;
    rb_completa: TRadioButton;
    rb_simplificada: TRadioButton;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    bt_ok: TAdvGlowButton;
    bt_cancelar: TAdvGlowButton;
    mResp: TMemo;
    procedure bt_cancelar1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ed_inicioKeyPress(Sender: TObject; var Key: Char);
    procedure ed_fimKeyPress(Sender: TObject; var Key: Char);
    procedure ed_data_fimKeyPress(Sender: TObject; var Key: Char);
    procedure ed_data_iniKeyPress(Sender: TObject; var Key: Char);
    procedure rb_dataClick(Sender: TObject);
    procedure rb_crzClick(Sender: TObject);
    procedure rb_dataKeyPress(Sender: TObject; var Key: Char);
    procedure rb_crzKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure bt_ok1Click(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

function Daruma_RSA_CriarAssinatura(caminhoDoArquivo: string; sMD5: String;
  sAssinaturaEletronica: string): Integer; StdCall; External 'Daruma32.dll';

var
  frmLeitura_Memoria_Fiscal: TfrmLeitura_Memoria_Fiscal;

implementation

uses funcoes, principal, msg_Operador, UFuncoes, venda;

{$R *.dfm}

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmLeitura_Memoria_Fiscal.bt_cancelar1Click(Sender: TObject);
begin
  close;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := cafree;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.ed_inicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    perform(wm_nextdlgctl, 0, 0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.ed_fimKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    bt_ok.setfocus;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.ed_data_fimKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    bt_ok.setfocus;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.ed_data_iniKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    perform(wm_nextdlgctl, 0, 0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.rb_dataClick(Sender: TObject);
begin
  ed_data_ini.Visible := true;
  ed_data_fim.Visible := true;
  ed_inicio.Visible := false;
  ed_fim.Visible := false;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.rb_crzClick(Sender: TObject);
begin
  ed_data_ini.Visible := false;
  ed_data_fim.Visible := false;
  ed_inicio.Visible := true;
  ed_fim.Visible := true;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.rb_dataKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    perform(wm_nextdlgctl, 0, 0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.rb_crzKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    perform(wm_nextdlgctl, 0, 0);
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.FormShow(Sender: TObject);
begin
  ed_data_ini.Date := Date;
  ed_data_fim.Date := Date;
  ed_inicio.Text := '0';
  ed_fim.Text := '0';

  rb_data.setfocus;
end;

// -------------------------------------------------------------------------- //
procedure TfrmLeitura_Memoria_Fiscal.bt_ok1Click(Sender: TObject);
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
    if rb_data.checked then begin
      if rb_ecf.checked then
        frmVenda.cpACBrECF.LeituraMemoriaFiscal(ed_data_ini.Date, ed_data_fim.Date,rb_simplificada.Checked)
      else begin
        frmVenda.cpACBrECF.LeituraMemoriaFiscalSerial(ed_data_ini.Date, ed_data_fim.Date,Linhas,rb_simplificada.Checked);
        mResp.Lines.Clear;
        For I := 0 to Linhas.Count - 1 do
          mResp.Lines.Add(Linhas[I]);
        if rb_completa.checked then
          NomeArquivo := ECFCaminhoFiscal+'\txt\LMFC.txt'
        else
          NomeArquivo := ECFCaminhoFiscal+'\txt\LMFS.txt';
        mResp.Lines.SaveToFile(NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end;
    end else begin
      if rb_ecf.checked then
        frmVenda.cpACBrECF.LeituraMemoriaFiscal(StrToInt(ed_inicio.Text), StrToInt(ed_fim.Text),rb_simplificada.Checked)
      else begin
        frmVenda.cpACBrECF.LeituraMemoriaFiscalSerial(StrToInt(ed_inicio.Text), StrToInt(ed_fim.Text),Linhas,rb_simplificada.Checked);
        mResp.Lines.Clear;
        For I := 0 to Linhas.Count - 1 do
          mResp.Lines.Add(Linhas[I]);
        if rb_completa.checked then
          NomeArquivo := ECFCaminhoFiscal+'\txt\LMFC.txt'
        else
          NomeArquivo := ECFCaminhoFiscal+'\txt\LMFS.txt';
        mResp.Lines.SaveToFile(NomeArquivo);
        Application.MessageBox(PWideChar('Arquivo gerado com sucesso em:'+NomeArquivo),'Atenção!',MB_ICONINFORMATION);
      end;
    end;
  finally
    Linhas.Free;
  end;
  close;
end;

end.
