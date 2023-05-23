unit caixa_abertura;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, AdvMenus,  StdCtrls, Mask, RzEdit, ExtCtrls,
  ComCtrls, NxColumnClasses, NxColumns, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, RzTabs, DB, DBAccess,
  XPMan, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel,
  Uni, MemDS, JvExMask, JvToolEdit;

type
  TfrmCaixa_Abertura = class(TForm)
    pop_fechamento: TAdvPopupMenu;
    Cancelar1: TMenuItem;
    AdvPopupMenu1: TAdvPopupMenu;
    MenuItem2: TMenuItem;
   // qrAbastecimento: TIBCQuery;
  //  qrEncerrante: TIBCQuery;
    bt_cupom_encerrante: TButton;
    XPManifest1: TXPManifest;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    ed_operador: TRzEdit;
    ed_data: TJvDateEdit;
    ed_ecf: TRzEdit;
    Panel1: TPanel;
    bt_ok: TButton;
    procedure bt_okClick(Sender: TObject);
    procedure bt_cancelarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure ed_operadorKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCaixa_Abertura: TfrmCaixa_Abertura;

implementation

uses funcoes, modulo, principal, msg_Operador,
  suprimento, venda, UFuncoes;

{$R *.dfm}


procedure TfrmCaixa_Abertura.bt_okClick(Sender: TObject);
begin
  (*Emitir leitura x*)

  sNumero_Cupom := Zerar( FloatToStr( max('')), 5);

  frmMsg_operador.Hide;

  frmmodulo.query.close;
  frmmodulo.query.sql.clear;
  frmmodulo.query.sql.add('update config set caixa_data_movto = :data, caixa_situacao = ''ABERTO'',');
  frmmodulo.query.sql.add('CAIXA_COO_INICIAL = '''+sNumero_Cupom+'''');
  frmmodulo.query.ParamByName('data').asdatetime := ed_data.Date;
  frmmodulo.query.ExecSQL;
  if TipoDeCupom <> tcPafECF then
    if Application.MessageBox('Deseja lançar algum suprimento?','Aviso',mb_yesno+MB_ICONQUESTION) = idYes then
    begin
      frmsuprimento := tfrmsuprimento.create(self);
      frmsuprimento.showmodal;
    end;
  bCaixa_aberto := true;
  close;
end;

procedure TfrmCaixa_Abertura.ed_operadorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then begin
    {Perform(wm_nextdlgctl, 0, 0);}
    bt_ok.SetFocus;
  end else if key = #27 then begin
    busuario_Autenticado := False;
    close;
  end;
end;

procedure TfrmCaixa_Abertura.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmCaixa_Abertura.bt_cancelarClick(Sender: TObject);
begin
  close;
end;

procedure TfrmCaixa_Abertura.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmCaixa_Abertura.FormShow(Sender: TObject);
var i : integer;
begin
  ed_data.Date := dData_Sistema;
  ed_operador.Text := sNome_Operador;
  ed_ecf.Text := sCaixa;
end;

end.
