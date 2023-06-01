program Pdv;



uses
  Winapi.Windows,
  Vcl.Forms,
  System.SysUtils,
  principal in 'principal.pas' {frmPrincipal},
  modulo in 'modulo.pas' {frmModulo: TDataModule},
  venda in 'venda.pas' {frmVenda},
  mesas in 'mesas.pas' {frmMesas},
  VendedorConsulta in 'VendedorConsulta.pas' {frmVendedorConsulta},
  preco_consulta in 'preco_consulta.pas' {frmPreco_consulta},
  protetor_tela in 'protetor_tela.pas' {frmProtetor_tela},
  venda_crediario in 'venda_crediario.pas' {frmVenda_Crediario},
  Vendedor in 'Vendedor.pas' {frmVendedor},
  msg_Operador in 'msg_Operador.pas' {frmMsg_Operador},
  Identifica in 'Identifica.pas' {frmIdentifica},
  mensagem in 'mensagem.pas' {frmMensagem},
  senha_supervisor in 'senha_supervisor.pas' {frmSenha_Supervisor},
  suprimento in 'suprimento.pas' {frmSuprimento},
  sangria in 'sangria.pas' {frmSangria},
  menu_cupom in 'menu_cupom.pas' {frmCupom_Menu},
  cupom_impressao in 'cupom_impressao.pas' {frmcupom_impressao},
  produto_consulta in 'produto_consulta.pas' {frmProduto_consulta},
  consulta_estoque in 'consulta_estoque.pas' {frmConsulta_Estoque},
  contasreceber in 's7\contasreceber.pas' {frmcontasreceber},
  contasreceber_boleto_achado in 's7\contasreceber_boleto_achado.pas' {frmcontasreceber_boleto_achado},
  xloc_cliente in 's7\xloc_cliente.pas' {frmxloc_cliente},
  contasreceber_pgto in 's7\contasreceber_pgto.pas' {frmcontasreceber_pgto},
  contasreceber_banco in 's7\contasreceber_banco.pas' {frmContasreceber_banco},
  loc_contacorrent in 's7\loc_contacorrent.pas' {frmloc_contacorrente},
  extenso1 in 's7\extenso1.pas',
  pre_venda in 'pre_venda.pas' {frmPre_Venda},
  caixa_abertura in 'caixa_abertura.pas' {frmCaixa_Abertura},
  config_balanca in 'config_balanca.pas' {frmconfig_balanca},
  Config in 'Config.pas' {frmConfig},
  cliente_consulta in 'cliente_consulta.pas' {frmCliente_consulta},
  UFuncoes in 'UFuncoes.pas',
  statusNFCE in 'statusNFCE.pas' {frmStatus},
  frmNFCEs in 'frmNFCEs.pas' {frmNotasconsumidor},
  UtiSplash in '..\FNTS_Comuns PDV\UtiSplash.pas' {frmUtiSplash},
  senha in '..\FNTS_Comuns PDV\senha.pas' {frmsenha},
  LoginWindow_U in '..\FNTS_Comuns\LoginWindow_U.pas' {frmLoginWindow},
  unSintegra in '..\FNTS_Sintegra\unSintegra.pas',
  Unit_Validacao_Sintegra in '..\FNTS_Sintegra\Validacao_Sintegra\Unit_Validacao_Sintegra.pas',
  Valida_NCM in '..\FNTS_Sintegra\Validacao_NCM\Valida_NCM.pas',
  Valida_CFOP in '..\FNTS_Sintegra\Validacao_CFOP\Valida_CFOP.pas',
  ResultadoWebService in 'ResultadoWebService.pas' {frmResultadoWebService},
  justificativa in 'justificativa.pas' {frmjustificativa},
  Thread.CheckConnection in '..\FNTS_Comuns\Thread.CheckConnection.pas',
  MensagemAdicional in 'MensagemAdicional.pas' {frmMensagemAdicional},
  uBandeira in 'uBandeira.pas' {frmBandeira},
  Parcelas in 'Parcelas.pas' {frmParcelas},
  impDelivery in 'impDelivery.pas' {frmImpDelivery},
  Extra in 'Extra.pas' {frmExtra},
  Recebimento in 'Recebimento.pas' {frmRecebimento},
  Receber_Valor in 'Receber_Valor.pas' {frmReceber_Valor},
  venda_convenio in 'venda_convenio.pas' {frmVenda_convenio},
  JustificativaInutilizacaoNFCe in '..\FNTS_Comuns\JustificativaInutilizacaoNFCe.pas' {frmJustificativaInutilizacaoNFCe},
  FechamentoCego in 'FechamentoCego.pas' {frmFechamentoCego},
  CuponsSAT in 'CuponsSAT.pas' {frmCuponsSAT},
  MontaPizza in 'MontaPizza.pas' {frmMontaPizza},
  PegaCupomGerencial in 'PegaCupomGerencial.pas' {frmPegaCupomGerencial},
  RecebimentoEspecial in 'RecebimentoEspecial.pas' {frmRecebimentoEspecial},
  LerCupomCreditoReceber in 'LerCupomCreditoReceber.pas' {frmLerCupomCreditoReceber},
  EnviaContingencia in 'EnviaContingencia.pas' {frmEnviaContingencia},
  FechamentoComanda in 'FechamentoComanda.pas' {frmFechamentoComanda},
  Comanda in 'Comanda.pas' {FComanda},
  Unit4 in 'Unit4.pas' {Form4},
  Unit5 in 'Unit5.pas' {Form5},
  uTEF in '..\FNTS_Personalizados\uTEF.pas' {frmTEF},
  CadastroCliente in '..\FNTS_Gerencial\fnts\CadastroCliente.pas' {frmCadastroCliente},
  CadMaquina in '..\FNTS_Gerencial\fnts\CadMaquina.pas' {frmCadMaquina},
  OnLine in '..\FNTS_Comuns\OnLine.pas' {frmOnLine},
  ConfiguraSerial in 'ConfiguraSerial.pas' {frConfiguraSerial},
  menu_fiscal in 'menu_fiscal.pas' {frmMenu_Fiscal},
  Leitura_Memoria_Fiscal in 'Leitura_Memoria_Fiscal.pas' {frmLeitura_Memoria_Fiscal},
  Download_MFD in 'Download_MFD.pas' {frmDownload_MFD},
  ErroEnvioNfce in 'ErroEnvioNfce.pas' {frmErroEnvioNfce},
  ModuloChef in '..\FNTS_Cheff\ModuloChef.pas' {frmModuloChef: TDataModule},
  principalChef in '..\FNTS_Cheff\principalChef.pas' {frmPrincipalChef},
  transf_Mesa in '..\FNTS_Cheff\transf_Mesa.pas' {frmTransf_Mesa},
  balanca in '..\FNTS_Cheff\balanca.pas' {frmBalanca},
  MenuMesa in '..\FNTS_Cheff\MenuMesa.pas' {frmMenuMesa},
  mesa_abertura in '..\FNTS_Cheff\mesa_abertura.pas' {frmmesa_abertura},
  mesa_consumo in '..\FNTS_Cheff\mesa_consumo.pas' {frmMesa_consumo},
  Consumo in '..\FNTS_Cheff\Consumo.pas' {frmConsumo},
  comandaChef in '..\FNTS_Cheff\comandaChef.pas' {frmComandaChef},
  FechamentoComandaChef in '..\FNTS_Cheff\FechamentoComandaChef.pas' {frmFechamentoComandaChef},
  ConfigChef in '..\FNTS_Cheff\ConfigChef.pas' {frmConfigChef},
  ConsultaEstoqueFilial in '..\FNTS_Comuns\ConsultaEstoqueFilial.pas',
  uFrmMaster in '..\FNTS_Comuns\uFrmMaster.pas' {FrmMaster},
  UCBase in '..\FNTS_Comuns\UCBase.pas',
  tabela_preco in 'tabela_preco.pas' {frm_tb_preco},
  loc_produto_cheff in '..\FNTS_Cheff\loc_produto_cheff.pas' {frmLoc_produto_cheff},
  SolicitaAutorizacaoGerencial in 'SolicitaAutorizacaoGerencial.pas' {frmSolicitaAutorizacaoGerencial},
  libeay32 in 'libeay32.pas',
  uAlterarValor in 'uAlterarValor.pas' {frmAlterarValorUnitario},
  funcoes in 'funcoes.pas',
  caixa_fechamento in 'caixa_fechamento.pas' {frmCaixa_Fechamento};

var
  Hand : THandle;

{$R *.res}

begin
  Hand := FindWindow('TApplication', 'SOS - 5.1 Frente de Caixa NFC-e');
  if Hand = 0 then begin
    Application.Initialize;
    frmUtiSplash := TfrmUtiSplash.Create(Application);
    frmUtiSplash.show;
    frmUtiSplash.Update;
    Application.Title := 'SOS - 5.1 Frente de Caixa NFC-e';
    Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(TfrmMsg_Operador, frmMsg_Operador);
  Application.CreateForm(TFrmMaster, FrmMaster);
  Application.CreateForm(TfrmSolicitaAutorizacaoGerencial, frmSolicitaAutorizacaoGerencial);
  Application.CreateForm(TfrmAlterarValorUnitario, frmAlterarValorUnitario);
  Application.CreateForm(TfrmCaixa_Fechamento, frmCaixa_Fechamento);
  Application.CreateForm(TfrmCaixa_Fechamento, frmCaixa_Fechamento);
  frmUtiSplash.Close;
    FreeAndNil(frmUtiSplash);
    application.ProcessMessages;
    Application.Run;
  end else begin
    application.MessageBox('Já existe uma sessão aberta deste aplicativo!','Atenção',mb_ok+MB_ICONERROR );
    ShowWindow(Hand,SW_NORMAL);
    Application.terminate;
  end;
end.
