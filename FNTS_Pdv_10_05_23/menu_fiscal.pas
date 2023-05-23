unit menu_fiscal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Menus, AdvMenus, DB, DBAccess, AdvToolBar, AdvToolBarStylers, AdvGlowButton, XPMan, AdvMetroButton,
  AdvSmoothPanel, AdvSmoothExpanderPanel, Uni, MemDS;

type
  TfrmMenu_Fiscal = class(TForm)
    mf001: TAdvGlowButton;
    mf002: TAdvGlowButton;
    mf004: TAdvGlowButton;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    AdvGlowButton1: TAdvGlowButton;
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure mf001Click(Sender: TObject);
    procedure mf002Click(Sender: TObject);
    procedure mf004Click(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMenu_Fiscal: TfrmMenu_Fiscal;

implementation

uses Leitura_Memoria_Fiscal, Download_MFD, modulo, funcoes, principal,
  msg_Operador, UFuncoes, venda;

{$R *.dfm}

// -------------------------------------------------------------------------- //
procedure TfrmMenu_Fiscal.mf001Click(Sender: TObject);
begin
  frmMsg_operador.lb_msg.Caption := 'Aguarde! Emitindo Leitura X...';
  frmMsg_Operador.Show;
  frmmsg_operador.Refresh;
  frmVenda.cpACBrECF.LeituraX;
  frmMsg_Operador.Hide;
end;

procedure TfrmMenu_Fiscal.mf002Click(Sender: TObject);
begin
  frmLeitura_Memoria_Fiscal := tfrmLeitura_Memoria_Fiscal.create(self);
  frmLeitura_Memoria_Fiscal.showmodal;
end;

procedure TfrmMenu_Fiscal.mf004Click(Sender: TObject);
begin
    frmDownload_MFD := tfrmDownload_MFD.Create(self);
    frmDownload_MFD.rb_arquivo.Enabled := false;
    frmDownload_MFD.rb_espelho.Enabled := true;
    frmDownload_MFD.rb_espelho.Checked := true;
    frmDownload_MFD.ShowModal;
end;

// -------------------------------------------------------------------------- //
procedure TfrmMenu_Fiscal.AdvGlowButton1Click(Sender: TObject);
var
  OldTimeOut: Integer;
begin
  OldTimeOut:= frmVenda.cpACBrECF.TimeOut;
  try
    Application.MessageBox('Atenção este processo pode demorar em algumas impressoras!','Atenção!',MB_ICONINFORMATION);
    frmMsg_operador.lb_msg.Caption := 'Aguarde! Realizando Redução Z...';
    frmMsg_Operador.Show;
    frmmsg_operador.Refresh;
    frmVenda.cpACBrECF.TimeOut := 600; // 10 minutos
    frmVenda.cpACBrECF.ReducaoZ(now);
  finally
    frmMsg_Operador.Hide;
    frmVenda.cpACBrECF.TimeOut := OldTimeOut;
  end;
end;

procedure TfrmMenu_Fiscal.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

// -------------------------------------------------------------------------- //
end.
