unit SolicitaAutorizacaoGerencial;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, uFrmMaster, AdvMetroButton,
  Vcl.ExtCtrls, Vcl.StdCtrls, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, cxTextEdit, cxMemo, cxLabel,
  Vcl.Menus, cxButtons, System.Actions, Vcl.ActnList, Data.DB, MemDS, DBAccess,
  Uni;

type
  TipoAutorizacao = (taCancItem, taCancCupom);
  TfrmSolicitaAutorizacaoGerencial = class(TFrmMaster)
    Panel2: TPanel;
    Panel3: TPanel;
    lbTempo: TcxLabel;
    cxLabel2: TcxLabel;
    cxButton1: TcxButton;
    ActionList1: TActionList;
    Action1: TAction;
    qrLib: TUniQuery;
    qrLibCODIGO: TStringField;
    qrLibCOD_OPERADOR: TStringField;
    qrLibTIPO: TStringField;
    qrLibDATA: TDateField;
    qrLibHORA: TStringField;
    qrLibMENSAGEM: TStringField;
    qrLibSTATUS: TStringField;
    qrLibAPROVADO_DATA: TDateField;
    qrLibAPROVADO_HORA: TStringField;
    qrLibAPROVADO_USUARIO: TIntegerField;
    Timer1: TTimer;
    procedure Action1Execute(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
    Codigo:string;
    Inicio:TTime;
    Retorno:Boolean;
  public
    { Public declarations }
  end;
  function AutorizaGerencial(Tipo:TipoAutorizacao):Boolean;

var
  frmSolicitaAutorizacaoGerencial: TfrmSolicitaAutorizacaoGerencial;

implementation

uses modulo, principal, venda, funcoes;

{$R *.dfm}

function AutorizaGerencial(Tipo:TipoAutorizacao):Boolean;
begin
  Application.CreateForm(TfrmSolicitaAutorizacaoGerencial, frmSolicitaAutorizacaoGerencial);
  with frmSolicitaAutorizacaoGerencial do begin
    qrLib.Close;
    qrLib.Open;
    Codigo := FormatDateTime('DDMMYYYYHHMMSSZZZ',Now);
    qrLib.Insert;
    qrLibCODIGO.AsString := Codigo;
    qrLibCOD_OPERADOR.AsString := Zerar(IntToStr(icodigo_Usuario),6);
    if Tipo = taCancItem then
      qrLibTIPO.AsString := 'I'
    else
      qrLibTIPO.AsString := 'C';
    qrLibDATA.AsDateTime := Date;
    qrLibHORA.AsString := FormatDateTime('HH:MM:SS',Time);
    qrLibSTATUS.AsString := 'A';
    qrLib.Post;
    frmModulo.Conexao_Servidor.Commit;
    Inicio := Time;
    Timer1.Enabled := True;
    Retorno := False;
    ShowModal;
    Result := Retorno;
  end;
  FreeAndNil(frmSolicitaAutorizacaoGerencial);
end;

procedure TfrmSolicitaAutorizacaoGerencial.Action1Execute(Sender: TObject);
begin
  inherited;
  Timer1.Enabled := False;
  qrLib.Close;
  qrLib.ParamByName('cod').AsString := Codigo;
  qrLib.Open;
  qrLib.Edit;
  qrLibSTATUS.AsString := 'C';
  qrLib.Post;
  frmModulo.Conexao_Servidor.Commit;
  Retorno := False;
  close;
end;

procedure TfrmSolicitaAutorizacaoGerencial.Timer1Timer(Sender: TObject);
begin
  inherited;
  try
    Timer1.Enabled := False;
    lbTempo.Caption := 'Tempo de Espera '  + FormatDateTime('HH:MM:SS',Time - Inicio) + '...';
    qrLib.Close;
    qrLib.ParamByName('cod').AsString := Codigo;
    qrLib.Open;
    if qrLibSTATUS.AsString = 'P' then begin
      Retorno := True;
      close;
    end else if qrLibSTATUS.AsString = 'N' then begin
      Retorno := False;
      Application.MessageBox('Cancelamento não autorizado!','Atenção!',MB_ICONERROR);
      close;
    end;
  finally
    Timer1.Enabled := True;
  end;
end;

end.
