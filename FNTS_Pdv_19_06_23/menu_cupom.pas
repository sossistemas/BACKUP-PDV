unit menu_cupom;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, ExtCtrls, Buttons, NxColumnClasses,
  NxColumns, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  DB, DBAccess, Menus, AdvMenus, ComCtrls, AdvMetroButton, AdvSmoothPanel, AdvSmoothExpanderPanel, AdvGlowButton, Uni,
  JvExMask, JvToolEdit, MemDS, principal;

type
  TfrmCupom_Menu = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    ed_data: TJvDateEdit;
    Panel2: TPanel;
    Bevel2: TBevel;
    grid_cupom: TNextGrid;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxNumberColumn1: TNxNumberColumn;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxNumberColumn2: TNxNumberColumn;
    BitBtn11: TBitBtn;
    bt_fechar1: TBitBtn;
  //  query: TIBCQuery;
    pop_fechamento: TAdvPopupMenu;
    Cancelar1: TMenuItem;
    Visualizar1: TMenuItem;
    BitBtn1: TAdvGlowButton;
    bt_fechar: TAdvGlowButton;
    AdvSmoothExpanderPanel1: TAdvSmoothExpanderPanel;
    Label53: TLabel;
    AdvMetroButton1: TAdvMetroButton;
    Panel3: TPanel;
    Panel4: TPanel;
    qrcupom: TUniQuery;
    query: TUniQuery;
    AdvGlowButton1: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    N1: TMenuItem;
    ESC: TMenuItem;
    agbReabrir: TAdvGlowButton;
    NxTextColumn5: TNxTextColumn;
    procedure bt_fechar1Click(Sender: TObject);
    procedure BitBtn11Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_cupomCellFormating(Sender: TObject; ACol, ARow: Integer;
      var TextColor: TColor; var FontStyle: TFontStyles;
      CellState: TCellState);
    procedure AdvMetroButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure agbReabrirClick(Sender: TObject);
  private
    { Private declarations }
    oldTipoImpressora: TImpressora;
    procedure AbreDados;
  public
    { Public declarations }
  end;

var
  frmCupom_Menu: TfrmCupom_Menu;

implementation

uses funcoes, modulo, cupom_impressao, venda, uTEF, senha_supervisor;

{$R *.dfm}

procedure TfrmCupom_Menu.bt_fechar1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmCupom_Menu.AbreDados;
begin
  qrcupom.Close;
  qrcupom.sql.clear;
  qrcupom.sql.add('select cupom.*,');
  qrcupom.sql.add('       cupom_consumidor.nome Consumidor');
  qrcupom.sql.add('from cupom, cupom_consumidor');
  qrcupom.sql.add('where');
  qrcupom.sql.add('  cupom.codigo = cupom_consumidor.codigo');
  qrcupom.sql.add('  and cupom.data = :data');
  qrcupom.sql.add('  and cupom.ecf = '''+sCaixa+'''');
  qrcupom.sql.add('order by cupom.numero');
  qrcupom.parambyname('data').asdatetime := ed_data.Date;
  qrcupom.open;

  grid_cupom.ClearRows;

  qrcupom.first;
  while not qrcupom.eof do
  begin
    grid_cupom.AddRow(1);
    grid_cupom.Cell[0,grid_cupom.LastAddedRow].AsString := qrcupom.fieldbyname('numero').asstring;
    grid_cupom.Cell[1,grid_cupom.LastAddedRow].AsString := qrcupom.fieldbyname('hora').asstring;
    grid_cupom.Cell[2,grid_cupom.LastAddedRow].AsString := qrcupom.fieldbyname('consumidor').asstring;
    grid_cupom.Cell[3,grid_cupom.LastAddedRow].Asfloat := qrcupom.fieldbyname('valor_total').asfloat;
    grid_cupom.Cell[4,grid_cupom.LastAddedRow].AsString := qrcupom.fieldbyname('codigo').asstring;
    grid_cupom.Cell[5,grid_cupom.LastAddedRow].Asinteger:= qrcupom.fieldbyname('cancelado').asinteger;
    grid_cupom.Cell[6,grid_cupom.LastAddedRow].AsString := qrcupom.fieldbyname('NUM_REABERTURA').AsString;
    qrcupom.Next;
  end;
  agbReabrir.Enabled := qrcupom.REcordCount > 0;
  grid_cupom.SelectedRow := grid_cupom.LastAddedRow;
  grid_cupom.SetFocus;
end;

procedure TfrmCupom_Menu.AdvGlowButton1Click(Sender: TObject);
begin
  if ed_data.Date > 0 then
    AbreDados
  else
    Application.MessageBox('Informe uma data valida!','Atenção!',MB_ICONINFORMATION);
end;

procedure TfrmCupom_Menu.AdvGlowButton2Click(Sender: TObject);
begin
  ///
  frmsenha_supervisor := tfrmsenha_supervisor.create(self);
  frmsenha_supervisor.ShowModal;
  if not bSupervisor_autenticado then
  begin
    frmsenha_supervisor.Free;
    ShowMessage('Senha inválida!');
    Exit;
  end;
  FreeAndNil(frmsenha_supervisor);
  ///
  with frmModulo do begin
    qrTemp.Close;
    qrTemp.SQL.Clear;
    qrTemp.SQL.Add('select cancelado, gerado_nfce, GEROU_SAT from cupom where numero = ' + QuotedStr(grid_cupom.Cell[0,grid_cupom.selectedrow].asstring));
    qrTemp.Open;
    if qrTemp.FieldByName('cancelado').AsInteger = 1 then begin
      Application.MessageBox('Cupom já está Cancelado!','Atenção!',MB_ICONERROR);
      Exit;
    end;
    if qrTemp.FieldByName('gerado_nfce').AsString = 'S' then begin
      Application.MessageBox('Este cupom deve ser cancelado pela tela de Gerenciamento de NFC-e!','Atenção!',MB_ICONERROR);
      Exit;
    end;
    if qrTemp.FieldByName('GEROU_SAT').AsString = 'S' then begin
      Application.MessageBox('Este cupom deve ser cancelado pela tela de Gerenciamento de SAT!','Atenção!',MB_ICONERROR);
      Exit;
    end;
    qrCupomCancelar.Open;
    qrCupomCancelar.Insert;
    qrCupomCancelarNUMERO.AsString := grid_cupom.Cell[0,grid_cupom.selectedrow].asstring;
    qrCupomCancelarNUM_REABERTURA.AsString := '';
    qrCupomCancelarDATA.AsDateTime := Date;
    qrCupomCancelarHORA.AsString := FormatDateTime('HH:MM:SS',Time);
    qrCupomCancelarUSU_CODIGO.AsInteger := icodigo_Usuario;
    qrCupomCancelarUSU_NOME.AsString := sNome_Operador;
    qrCupomCancelarEX.AsString := 'N';
    qrCupomCancelar.Post;
    qrCupomCancelar.ApplyUpdates;
    qrTemp.Close;
    qrTemp.SQL.Clear;
    qrTemp.SQL.Add('update cupom set cancelado = 1 where numero = ' + QuotedStr(grid_cupom.Cell[0,grid_cupom.selectedrow].asstring));
    qrTemp.ExecSQL;
    conexao.Commit;
    AdvGlowButton1Click(AdvGlowButton1);
    Application.MessageBox('Cupom Cancelado com sucesso!','Atenção!',MB_ICONINFORMATION);
    Exit;
  end;
end;

procedure TfrmCupom_Menu.agbReabrirClick(Sender: TObject);
var
  chr: Char;
  rSat: TSat;
  NomeArquivo, sCod_Cupom: string;
  lAtualizarValorUnitario: Boolean;
  lValorUnitario, lValorTotal: Real;
begin
  if Assigned(Sender) then
  begin
    if grid_cupom.RowCount = 0 then exit;
    if grid_cupom.SelectedRow < 0 then
    begin
      application.messagebox('Favor selecionar um cupom!','Erro',mb_ok+mb_iconerror);
      exit;
    end;

    with frmModulo do
    begin
      qrTemp.Close;
      qrTemp.SQL.Clear;
      qrTemp.SQL.Add('select cancelado, gerado_nfce, GEROU_SAT from cupom where numero = ' + QuotedStr(grid_cupom.Cell[0,grid_cupom.selectedrow].asstring));
      qrTemp.Open;
      if (qrTemp.FieldByName('cancelado').AsInteger = 1) or
         (qrTemp.FieldByName('gerado_nfce').AsString = 'S') or
         (qrTemp.FieldByName('GEROU_SAT').AsString = 'S')
      then
      begin
        Application.MessageBox('Cupom não pode ser reaberto!','Atenção!',MB_ICONERROR);
        Exit;
      end;
    end;

    if iTotal_Itens = 0 then
    try
      lAtualizarValorUnitario := Application.MessageBox(PWideChar('Manter os valores de venda na reabertura?'''),'Opções!'
                                                       ,MB_YESNO + MB_ICONQUESTION) <> mrYes;

      if not bVenda then
      try
        // verificar o serial do ecf
        // Abrir o cupom fiscal
        // bloquear ao teclado
        if not frmVenda.Abre_Venda then
        begin
          bLanca_pre_venda := False;
          bLanca_OS := False;
          bLanca_Mesa := False;
          bLanca_Abastecimento := False;
          // Caso o comando de abertura de venda retornou false, abortar o processo;
          Application.MessageBox(PWideChar('Ocorreu um erro a abertura do novo cupom!'),'Atenção!',MB_ICONERROR);
          Exit;
        end;
      finally
        // desbloquear o teclado
      end;

      rTotal_Desconto := 0;
      rTotal_Acrescimo := 0;
      rTotal_Venda := 0;

      // filtrando os itens
      query.close;
      query.sql.clear;
      query.sql.add('select cupom_item.*, estoque.nome nome_produto from cupom_item, estoque');
      query.sql.add('where cod_cupom = '''+grid_cupom.Cell[4,grid_cupom.selectedrow].asstring+'''');
      query.sql.add('and cupom_item.cod_produto = estoque.codigo');
      query.sql.add('order by cupom_item.item');
      query.open;
      query.first;
      while not query.eof do
      begin
        frmModulo.spCupom_item_temp.Close;
        frmModulo.spCupom_item_temp.ParamByName('Codigo').AsString := codifica_item(query.RecNo);
        frmModulo.spCupom_item_temp.parambyname('cod_cupom').asstring := sNumero_Cupom;
        frmModulo.spCupom_item_temp.parambyname('item').asinteger := query.fieldbyname('item').asinteger;
        frmModulo.spCupom_item_temp.parambyname('unidade').asstring := query.fieldbyname('unidade').asstring;
        frmModulo.spCupom_Item_temp.parambyname('qtde').asfloat := query.fieldbyname('qtde').asfloat;
        frmModulo.spCupom_item_temp.ParamByName('cod_produto').asinteger := query.fieldbyname('cod_produto').asinteger;
        ///
        if lAtualizarValorUnitario then
        begin
          lValorUnitario := frmModulo.PrecoDeVenda(query.fieldbyname('cod_produto').asinteger
                                                  ,query.fieldbyname('qtde').asfloat);
          lValorTotal := lValorUnitario * query.fieldbyname('qtde').asfloat;
          frmModulo.spCupom_Item_temp.parambyname('valor_subtotal').asfloat := lValorTotal;
        end
        else
        begin
          lValorUnitario := query.fieldbyname('valor_unitario').asfloat;
          lValorTotal := query.fieldbyname('valor_total').asfloat;
          frmModulo.spCupom_Item_temp.parambyname('valor_subtotal').asfloat := query.fieldbyname('VALOR_SUBTOTAL').asfloat;
        end;
        ///
        frmModulo.spCupom_Item_temp.parambyname('valor_unitario').asfloat := lValorUnitario;
        frmModulo.spCupom_Item_temp.parambyname('valor_total').asfloat := lValorTotal;
        ///
        frmModulo.spCupom_Item_temp.parambyname('valor_desconto').asfloat := query.fieldbyname('valor_desconto').asfloat;
        frmModulo.spCupom_Item_temp.parambyname('valor_acrescimo').asfloat := query.fieldbyname('valor_acrescimo').asfloat;
        frmModulo.spCupom_Item_temp.parambyname('complemento').AsString := query.fieldbyname('COMPLEMENTO').asstring;
        frmModulo.spCupom_item_temp.parambyname('cst').asstring := query.fieldbyname('CST').asstring;
        frmModulo.spCupom_item_temp.parambyname('aliquota').asfloat := query.fieldbyname('ALIQUOTA').asfloat;
        frmModulo.spCupom_item_temp.parambyname('cancelado').asinteger := 0;
        frmModulo.spCupom_Item_temp.parambyname('tamanho').AsString := query.fieldbyname('TAMANHO').asstring;;
        frmModulo.spCupom_Item_temp.parambyname('cor').AsString := query.fieldbyname('COR').asstring;;
        frmModulo.spCupom_Item_temp.parambyname('por_volume').asinteger := 0;
        frmModulo.spCupom_item_temp.ParamByName('cod_totalizador').asstring := query.fieldbyname('COD_TOTALIZADOR').asstring;;
        frmModulo.spCupom_item_temp.ParamByName('extra').asstring := query.fieldbyname('EXTRA').asstring;;
        frmModulo.spCupom_item_temp.Prepare;
        frmModulo.spCupom_item_temp.Execute;
        ///
        rTotal_Desconto := rTotal_Desconto + query.fieldbyname('valor_desconto').asfloat;
        rTotal_Acrescimo := rTotal_Acrescimo + query.fieldbyname('valor_acrescimo').asfloat;
        rTotal_Venda := rTotal_Venda + lValorTotal;

        query.next;
      end;
      iTotal_Itens := query.RecordCount;
      iItem := Succ(iTotal_Itens);
      // atualizar os labels de totais e itens
      frmVenda.lb_item.Caption := IntToStr(iTotal_Itens);
      frmVenda.lb_total.Caption := 'R$ ' + FormatFloat('###,###,##0.00', rTotal_Venda);
      frmVenda.ed_barra.text := '';
      rBal_peso := 0;
      vExtra.Texto := '';
      vExtra.valor := 0;
      frmVenda.ed_qtde.Value := 0;
      frmVenda.strPizza.Pizza := '';
      frmVenda.strPizza.Sabores := '';
      frmVenda.strPizza.Valor := 0;
      SetLength(frmVenda.strPizza.Produtos, 0);
      frmVenda.strPizza.ProdutoPrincipal.Codigo := '';
      frmVenda.strPizza.ProdutoPrincipal.Valor := 0;
      sProd_Compl_Pizza := '';
      ReaberturaCupomAtiva := grid_cupom.Cell[0,grid_cupom.selectedrow].asstring;
      Close;
    except
      Application.MessageBox(PWideChar('Ocorreu um erro na reabertura do cupom!'#13'Tente novamente e caso persista contate o suporte'),'Atenção!',MB_ICONERROR);
      ReaberturaCupomAtiva := '';
    end;
  end
  else
  try
    /////////////////////////////////////////////////////////////////////////////////////////////////
    ///  Cancela cupom
    with frmModulo do
    begin
      qrCupomCancelar.Open;
      qrCupomCancelar.Insert;
      qrCupomCancelarNUMERO.AsString := ReaberturaCupomAtiva;
      qrCupomCancelarNUM_REABERTURA.AsString := sNumero_Cupom;
      qrCupomCancelarDATA.AsDateTime := Date;
      qrCupomCancelarHORA.AsString := FormatDateTime('HH:MM:SS',Time);
      qrCupomCancelarUSU_CODIGO.AsInteger := icodigo_Usuario;
      qrCupomCancelarUSU_NOME.AsString := sNome_Operador;
      qrCupomCancelarEX.AsString := 'N';
      qrCupomCancelar.Post;
      qrCupomCancelar.ApplyUpdates;
      qrTemp.Close;
      qrTemp.SQL.Clear;
      qrTemp.SQL.Add('update cupom set cancelado = 1, NUM_REABERTURA = ' + QuotedStr(sNumero_Cupom) + ' where numero = ' + QuotedStr(ReaberturaCupomAtiva));
      qrTemp.ExecSQL;
      conexao.Commit;
      ReaberturaCupomAtiva := '';
    end;
  except
    Application.MessageBox(PWideChar('Ocorreu um erro no cancelamento do cupom!'),'Atenção!',MB_ICONERROR);
  end;
end;

procedure TfrmCupom_Menu.AdvMetroButton1Click(Sender: TObject);
begin
  close;
end;

procedure TfrmCupom_Menu.BitBtn11Click(Sender: TObject);
var i : integer;
begin
  if grid_cupom.RowCount = 0 then exit;
  if grid_cupom.SelectedRow < 0 then
  begin
    application.messagebox('Favor selecionar um cupom!','Erro',mb_ok+mb_iconerror);
    exit;
  end;

  frmcupom_impressao := Tfrmcupom_impressao.create(self);
  with frmcupom_impressao do
  begin
    Cupom := grid_cupom.Cell[4,grid_cupom.selectedrow].asstring;
    grid_impressao.ClearRows;
    // fazendo o cabecario
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      texto_justifica('R E I M P R E S S A O    D E    C U P O M',48,' ',taCentralizado);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      texto_justifica('No: '+grid_cupom.Cell[0,grid_cupom.SelectedRow].asstring+
                      ' Data:'+ed_data.text+
                      ' Hora:'+grid_cupom.cell[1,grid_cupom.selectedrow].asstring,48,
                      ' ',taCentralizado);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '------------------------------------------------';
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      'Item Codigo        Descricao';
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '                    Qtde Un.  Vl.Unit    Vl.Item';
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '------------------------------------------------';

    // filtrando os itens
    query.close;
    query.sql.clear;
    query.sql.add('select cupom_item.*, estoque.nome nome_produto from cupom_item, estoque');
    query.sql.add('where cod_cupom = '''+grid_cupom.Cell[4,grid_cupom.selectedrow].asstring+'''');
    query.sql.add('and cupom_item.cod_produto = estoque.codigo and cupom_item.cancelado = 0');
    query.sql.add('order by cupom_item.item');
    query.open;
    query.first;
    while not query.eof do
    begin
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].asstring :=
        texto_justifica(query.fieldbyname('item').asstring,3,'0',taDireita)+' '+
        texto_justifica(query.fieldbyname('cod_produto').asstring,13,'0',taDireita)+' '+
        texto_justifica(query.fieldbyname('nome_produto').asstring,30,' ',taEsquerda);
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].asstring :='              '+
        texto_justifica(formatfloat('###,###,##0.000',query.fieldbyname('qtde').asfloat),10,' ',taDireita)+
        texto_justifica(query.fieldbyname('unidade').asstring,2,' ',taEsquerda)+' '+
        texto_justifica(formatfloat('###,###,##0.000',query.fieldbyname('valor_unitario').asfloat),10,' ',taDireita)+
        texto_justifica(formatfloat('###,###,##0.00',query.fieldbyname('valor_total').asfloat),11,' ',taDireita);

      if query.fieldbyname('valor_desconto').asfloat > 0 then
      begin
         grid_impressao.AddRow(1);
         grid_impressao.Cell[0,grid_impressao.LastAddedRow].asstring :='              '+
           'Desconto:  -'+texto_justifica(formatfloat('###,###,##0.00',query.fieldbyname('valor_desconto').asfloat),8,' ',taDireita);
      end;
      if query.fieldbyname('valor_acrescimo').asfloat > 0 then
      begin
         grid_impressao.AddRow(1);
         grid_impressao.Cell[0,grid_impressao.LastAddedRow].asstring :='              '+
           'Acréscimo: +'+texto_justifica(formatfloat('###,###,##0.00',query.fieldbyname('valor_acrescimo').asfloat),8,' ',taDireita);
      end;
      query.next;
    end;

    // filtrar o cupom fiscal

    qrcupom.Close;
    qrcupom.sql.clear;
    qrcupom.sql.add('select cupom.*,');
    qrcupom.sql.add('       cupom_consumidor.*');
    qrcupom.sql.add('from cupom, cupom_consumidor');
    qrcupom.sql.add('where');
    qrcupom.sql.add('  cupom.codigo = cupom_consumidor.codigo');
    qrcupom.sql.add('  and cupom.data = :data');
    qrcupom.sql.add('  and cupom.ecf = '''+sCaixa+'''');
    qrcupom.sql.add('  and cupom.codigo = '''+grid_cupom.Cell[4,grid_cupom.selectedrow].asstring+'''');
    qrcupom.sql.add('order by cupom.numero');
    qrcupom.parambyname('data').asdatetime := ed_data.Date;
    qrcupom.open;

    NaoFiscal := qrcupom.FieldByName('NAOFISC').AsString = 'S';

    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '------------------------------------------------';
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '                   Subtotal do Cupom:'+texto_justifica(formatfloat('###,###,##0.00',qrcupom.fieldbyname('valor_produto').asfloat),11,' ',taDireita);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '                            Desconto:'+texto_justifica(formatfloat('###,###,##0.00',qrcupom.fieldbyname('valor_desconto').asfloat),11,' ',taDireita);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '                           Acrescimo:'+texto_justifica(formatfloat('###,###,##0.00',qrcupom.fieldbyname('valor_acrescimo').asfloat),11,' ',taDireita);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '                      Total do Cupom:'+texto_justifica(formatfloat('###,###,##0.00',qrcupom.fieldbyname('valor_total').asfloat),11,' ',taDireita);
    grid_impressao.AddRow(1);
    grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
      '------------------------------------------------';
    // filtrar os pagamentos

    query.close;
    query.sql.clear;
    query.sql.add('select * from cupom_forma where cod_cupom = '''+grid_cupom.cell[4,grid_cupom.selectedrow].asstring+'''');
    query.open;
    query.first;
    while not query.eof do
    begin
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica(query.fieldbyname('forma').asstring,36,' ',taDireita)+':'+texto_justifica(formatfloat('###,###,##0.00',query.fieldbyname('valor').asfloat),11,' ',taDireita);
      query.Next;
    end;
    if qrcupom.fieldbyname('valor_troco').asfloat > 0 then
    begin
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        '                               Troco:'+texto_justifica(formatfloat('###,###,##0.00',qrcupom.fieldbyname('valor_troco').asfloat),11,' ',taDireita);
    end;
    if qrcupom.FieldByName('nome').asstring <> '' then
    begin
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica('Nome....:'+qrcupom.fieldbyname('nome').asstring,48,' ',taEsquerda);
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica('Endereco:'+qrcupom.fieldbyname('endereco').asstring,48,' ',taEsquerda);
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica('Cidad/UF:'+qrcupom.fieldbyname('cidade').asstring+'/'+qrcupom.fieldbyname('uf').asstring,48,' ',taEsquerda);
      grid_impressao.AddRow(1);
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica('CPF.....:'+qrcupom.fieldbyname('cpf').asstring,48,' ',taEsquerda);

      if qrcupom.fieldbyname('placa').asstring <> '' then
      grid_impressao.Cell[0,grid_impressao.LastAddedRow].AsString :=
        texto_justifica('Placa...:'+qrcupom.fieldbyname('PLACA').asstring
      +'  Km: '+qrcupom.fieldbyname('km').asstring+' Vd: '+
      qrcupom.fieldbyname('vendedor').asstring,48,' ',taEsquerda);

    end;
  end;
  hide;
  frmcupom_impressao.ShowModal;
  show;
end;

procedure TfrmCupom_Menu.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frmPrincipal.TipoImpressora := oldTipoImpressora;
  action := cafree;
end;

procedure TfrmCupom_Menu.FormCreate(Sender: TObject);
begin
  oldTipoImpressora := frmPrincipal.TipoImpressora;
  frmPrincipal.TipoImpressora := NaoFiscal;
end;

procedure TfrmCupom_Menu.FormShow(Sender: TObject);
begin
  ed_data.Date := dData_Sistema;
  AbreDados;
end;

procedure TfrmCupom_Menu.grid_cupomCellFormating(Sender: TObject; ACol,
  ARow: Integer; var TextColor: TColor; var FontStyle: TFontStyles;
  CellState: TCellState);
begin
  if grid_cupom.Cell[5,arow].AsInteger = 1 then TextColor := clred;
end;

end.
