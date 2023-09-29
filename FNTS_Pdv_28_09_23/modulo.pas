unit modulo;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, ClipBrd,
  Dialogs, StdCtrls, Mask, DBCtrls, DB, DBAccess, IniFiles, ACBrDFe, ACBrNFe,
  ACBrNFeDANFEFRDM, ACBrNFeDANFEClass, ACBrNFeDANFEFR, ACBrEAD, Uni, UniProvider,
  InterBaseUniProvider, Datasnap.DBClient, MemDS, VirtualTable, ACBrValidador,
  AdvMenus, AdvMenuStylers, ACBrBase, ACBrBAL, pcnConversao, ACBrDFeSSL,
  blcksock, Vcl.FileCtrl, pcnConversaoNFe, ACBrUtil, System.Math, System.TypInfo,
  ACBrGAV, ACBrDFeReport, ACBrDFeDANFeReport, ACBrPosPrinter;

type
  TvNFe = (v310, v400);

  TValorUnitario = record
   Produto: Integer;
   ValorVenda: Double;
   IndiceAtivo: Integer;
   ValorDesconto: array of Double;
  end;

  tATB = record
    ATB: string;
    ID, PAI, FILHO: Integer;
    TIPO, DESCRICAO: string
  end;

  TfrmModulo = class(TDataModule)
    Balanca: TACBrBAL;
    estilo_menu: TAdvMenuOfficeStyler;
    docValidador: TACBrValidador;
    tbForma_Pgto: TVirtualTable;
    tbForma_PgtoNome: TStringField;
    tbForma_PgtoValor_Acumulado: TFloatField;
    tbForma_PgtoValor_Ultimo_Cupom: TFloatField;
    tbForma_Pgtoid: TIntegerField;
    tbAliquota: TVirtualTable;
    tbAliquotaCodigo: TStringField;
    tbAliquotaAliquota: TFloatField;
    Dlg_arquivo: TOpenDialog;
    conexao: TUniConnection;
    Conexao_Servidor: TUniConnection;
    query_servidor2: TUniQuery;
    qrtransportador: TUniQuery;
    qrfilial: TUniQuery;
    qrlancamento_conta: TUniQuery;
    qrcaixa_operador: TUniQuery;
    qrconfig: TUniQuery;
    qrcaixa_mov: TUniQuery;
    qrcontacorrente: TUniQuery;
    qrbanco: TUniQuery;
    qrproduto: TUniQuery;
    qrfornecedor: TUniQuery;
    qrproduto_mov: TUniQuery;
    qrgrade_produto: TUniQuery;
    qrfiscal_modelo: TUniQuery;
    qrcfop: TUniQuery;
    qrgrupo: TUniQuery;
    qrmarca: TUniQuery;
    qrsubgrupo: TUniQuery;
    qrcontasreceber: TUniQuery;
    qrtotalizador: TUniQuery;
    query_servidor: TUniQuery;
    qrMestre: TUniQuery;
    qrcliente: TUniQuery;
    query: TUniQuery;
    InterBaseUniProvider1: TInterBaseUniProvider;
    spCupom_Crediario: TUniStoredProc;
    spCupom_Consumidor: TUniStoredProc;
    spCupom_Item_Temp: TUniStoredProc;
    spCupom_Temp: TUniStoredProc;
    spCupom_Temp_Edit: TUniStoredProc;
    spCupom_Temp_delete: TUniStoredProc;
    qrproduto_movCODIGO: TStringField;
    qrproduto_movCODNOTA: TStringField;
    qrproduto_movCODPRODUTO: TStringField;
    qrproduto_movUNITARIO: TFloatField;
    qrproduto_movTOTAL: TFloatField;
    qrproduto_movICMS: TFloatField;
    qrproduto_movIPI: TFloatField;
    qrproduto_movCFOP: TStringField;
    qrproduto_movDATA: TDateTimeField;
    qrproduto_movNUMERONOTA: TStringField;
    qrproduto_movCODCLIENTE: TStringField;
    qrproduto_movDESCONTO: TFloatField;
    qrproduto_movACRESCIMO: TFloatField;
    qrproduto_movMOVIMENTO: TIntegerField;
    qrproduto_movCODVENDEDOR: TStringField;
    qrproduto_movCODGRADE: TStringField;
    qrproduto_movSERIAL: TStringField;
    qrproduto_movUNIDADE: TStringField;
    qrproduto_movQTDE: TFloatField;
    qrproduto_movVALOR_ICMS: TFloatField;
    qrproduto_movICMS_REDUZIDO: TFloatField;
    qrproduto_movBASE_CALCULO: TFloatField;
    qrproduto_movVALOR_IPI: TFloatField;
    qrproduto_movSITUACAO: TIntegerField;
    qrproduto_movECF_SERIE: TStringField;
    qrproduto_movECF_CAIXA: TStringField;
    qrproduto_movCODALIQUOTA: TStringField;
    qrproduto_movCUPOM_NUMERO: TStringField;
    qrproduto_movCUPOM_MODELO: TStringField;
    qrproduto_movCUPOM_ITEM: TStringField;
    qrproduto_movALIQUOTA: TFloatField;
    qrproduto_movCST: TStringField;
    qrproduto_movLOTE_FABRICACAO: TStringField;
    qrproduto_movMOVIMENTO_ESTOQUE: TFloatField;
    qrproduto_movLANCADO: TIntegerField;
    qrproduto_movVENCIMENTO: TDateTimeField;
    qrproduto_movCODBARRA: TStringField;
    qrproduto_movMARGEM_DESCONTO: TFloatField;
    qrproduto_movCREDITO_ICMS: TFloatField;
    qrproduto_movPIS: TFloatField;
    qrproduto_movCOFINS: TFloatField;
    qrproduto_movLOJA: TStringField;
    qrproduto_movCODSUBGRUPO: TStringField;
    qrproduto_movTIPO: TStringField;
    qrproduto_movCODUSUARIO: TStringField;
    qrproduto_movORIGEM: TStringField;
    qrproduto_movDESTINO: TStringField;
    qrproduto_movPRODUTO: TStringField;
    qrproduto_movCODFILIAL: TStringField;
    qrfornecedorCODIGO: TStringField;
    qrfornecedorNOME: TStringField;
    qrfornecedorFANTASIA: TStringField;
    qrfornecedorENDERECO: TStringField;
    qrfornecedorBAIRRO: TStringField;
    qrfornecedorCIDADE: TStringField;
    qrfornecedorUF: TStringField;
    qrfornecedorCEP: TStringField;
    qrfornecedorCOMPLEMENTO: TStringField;
    qrfornecedorTELEFONE1: TStringField;
    qrfornecedorTELEFONE2: TStringField;
    qrfornecedorFAX: TStringField;
    qrfornecedorCONTATO1: TStringField;
    qrfornecedorCONTATO2: TStringField;
    qrfornecedorCELULAR1: TStringField;
    qrfornecedorCELULAR2: TStringField;
    qrfornecedorEMAIL: TStringField;
    qrfornecedorHOMEPAGE: TStringField;
    qrfornecedorCNPJ: TStringField;
    qrfornecedorIE: TStringField;
    qrfornecedorBANCO: TStringField;
    qrfornecedorAGENCIA: TStringField;
    qrfornecedorCONTA: TStringField;
    qrfornecedorOBS1: TStringField;
    qrfornecedorOBS2: TStringField;
    qrfornecedorOBS3: TStringField;
    qrfornecedorDATA: TDateTimeField;
    qrfornecedorTIPO: TIntegerField;
    qrfornecedorASSISTENCIA_TECNICA: TIntegerField;
    qrfornecedorNUMERO: TStringField;
    qrfornecedorIM: TStringField;
    qrfornecedorREP_NOME: TStringField;
    qrfornecedorREP_TELEFONE: TStringField;
    qrfornecedorREP_ENDERECO: TStringField;
    qrfornecedorREP_BAIRRO: TStringField;
    qrfornecedorREP_COMPLEMENTO: TStringField;
    qrfornecedorREP_CIDADE: TStringField;
    qrfornecedorREP_UF: TStringField;
    qrfornecedorREP_CEP: TStringField;
    qrfornecedorREP_TELEFONE1: TStringField;
    qrfornecedorREP_TELEFONE2: TStringField;
    qrfornecedorREP_TELEFONE3: TStringField;
    qrfornecedorREP_FAX: TStringField;
    qrfornecedorREP_CNPJ: TStringField;
    qrfornecedorREP_IE: TStringField;
    qrfornecedorREP_HOME_PAGE: TStringField;
    qrfornecedorREP_EMAIL: TStringField;
    qrfornecedorCNAE: TStringField;
    qrfornecedorCOD_MUNICIPIO_IBGE: TStringField;
    qrfornecedorIBGE: TStringField;
    qrtransportadorCODIGO: TStringField;
    qrtransportadorNOME: TStringField;
    qrtransportadorENDERECO: TStringField;
    qrtransportadorBAIRRO: TStringField;
    qrtransportadorCIDADE: TStringField;
    qrtransportadorUF: TStringField;
    qrtransportadorCEP: TStringField;
    qrtransportadorCPF: TStringField;
    qrtransportadorRG: TStringField;
    qrtransportadorTELEFONE: TStringField;
    qrtransportadorCELULAR: TStringField;
    qrtransportadorPLACA: TStringField;
    qrtransportadorUFPLACA: TStringField;
    qrtransportadorOBS1: TStringField;
    qrtransportadorOBS2: TStringField;
    qrtransportadorOBS3: TStringField;
    qrtransportadorDATA: TDateTimeField;
    qrtransportadorTIPO: TIntegerField;
    qrtransportadorNUMERO: TStringField;
    qrtransportadorCOD_MUNICIPIO_IBGE: TStringField;
    qrtransportadorIBGE: TStringField;
    qrtransportadorANTT: TStringField;
    qrlancamento_contaCODIGO: TStringField;
    qrlancamento_contaCODCONTACORRENTE: TStringField;
    qrlancamento_contaDATA: TDateTimeField;
    qrlancamento_contaHISTORICO: TStringField;
    qrlancamento_contaCODCONTA: TStringField;
    qrlancamento_contaVALOR: TFloatField;
    qrlancamento_contaTIPO: TStringField;
    qrlancamento_contaDOCUMENTO: TStringField;
    qrbancoNUMERO: TStringField;
    qrbancoBANCO: TStringField;
    qrbancoLOGO: TBlobField;
    qrbancoCARTAO_CREDITO: TIntegerField;
    qrbancoFINANCEIRA: TIntegerField;
    qrbancoRESSARCIMENTO: TIntegerField;
    qrbancoPRAZO: TIntegerField;
    qrbancoCOMISSAO_CREDITO: TFloatField;
    qrbancoCOMISSAO_DEBITO: TFloatField;
    qrbancoREC_DEBITO: TIntegerField;
    qrbancoREC_CREDITO: TIntegerField;
    qrbancoCONTA_PADRAO: TStringField;
    qrbancoTITULAR_CONTA_PADRAO: TStringField;
    qrprodutoCODIGO: TStringField;
    qrprodutoCODBARRA: TStringField;
    qrprodutoPRODUTO: TStringField;
    qrprodutoDATA_CADASTRO: TDateTimeField;
    qrprodutoCODGRUPO: TStringField;
    qrprodutoCODSUBGRUPO: TStringField;
    qrprodutoCODFORNECEDOR: TStringField;
    qrprodutoCODMARCA: TStringField;
    qrprodutoDATA_ULTIMACOMPRA: TDateTimeField;
    qrprodutoNOTAFISCAL: TStringField;
    qrprodutoPRECOCUSTO: TFloatField;
    qrprodutoPRECOVENDA: TFloatField;
    qrprodutoDATA_ULTIMAVENDA: TDateTimeField;
    qrprodutoESTOQUE: TFloatField;
    qrprodutoESTOQUEMINIMO: TFloatField;
    qrprodutoCODALIQUOTA: TStringField;
    qrprodutoAPLICACAO: TMemoField;
    qrprodutoLOCALICAZAO: TStringField;
    qrprodutoPESO: TFloatField;
    qrprodutoVALIDADE: TStringField;
    qrprodutoCOMISSAO: TFloatField;
    qrprodutoUSA_BALANCA: TIntegerField;
    qrprodutoUSA_SERIAL: TIntegerField;
    qrprodutoUSA_GRADE: TIntegerField;
    qrprodutoCODRECEITA: TStringField;
    qrprodutoFOTO: TStringField;
    qrprodutoDATA_ULTIMACOMPRA_ANTERIOR: TDateTimeField;
    qrprodutoNOTAFISCAL_ANTERIOR: TStringField;
    qrprodutoCODFORNECEDOR_ANTERIOR: TStringField;
    qrprodutoPRECOCUSTO_ANTERIOR: TFloatField;
    qrprodutoPRECOVENDA_ANTERIOR: TFloatField;
    qrprodutoCUSTOMEDIO: TFloatField;
    qrprodutoAUTO_APLICACAO: TStringField;
    qrprodutoAUTO_COMPLEMENTO: TStringField;
    qrprodutoDATA_REMARCACAO_CUSTO: TDateTimeField;
    qrprodutoDATA_REMARCACAO_VENDA: TDateTimeField;
    qrprodutoPRECO_PROMOCAO: TFloatField;
    qrprodutoDATA_PROMOCAO: TDateTimeField;
    qrprodutoFIM_PROMOCAO: TDateTimeField;
    qrprodutoCST: TStringField;
    qrprodutoCLASSIFICACAO_FISCAL: TStringField;
    qrprodutoNBM: TStringField;
    qrprodutoNCM: TStringField;
    qrprodutoALIQUOTA: TFloatField;
    qrprodutoIPI: TFloatField;
    qrprodutoREDUCAO: TFloatField;
    qrprodutoQTDE_EMBALAGEM: TFloatField;
    qrprodutoTIPO: TStringField;
    qrprodutoPESO_LIQUIDO: TFloatField;
    qrprodutoFARMACIA_CONTROLADO: TStringField;
    qrprodutoFARMACIA_APRESENTACAO: TIntegerField;
    qrprodutoFARMACIA_REGISTRO_MEDICAMENTO: TStringField;
    qrprodutoFARMACIA_PMC: TFloatField;
    qrprodutoULTIMA_ALTERACAO: TDateTimeField;
    qrprodutoULTIMA_CARGA: TDateTimeField;
    qrprodutoDATA_INVENTARIO: TDateTimeField;
    qrprodutoCUSTO_INVENTARIO: TFloatField;
    qrprodutoESTOQUE_INVENTARIO: TFloatField;
    qrprodutoESTOQUE_ANTERIOR: TFloatField;
    qrprodutoPRECOVENDA_NOVO: TFloatField;
    qrprodutoUSA_RENTABILIDADE: TIntegerField;
    qrprodutoQUANTIDADE_MINIMA_FAB: TFloatField;
    qrprodutoAPRESENTACAO: TStringField;
    qrprodutoSITUACAO: TIntegerField;
    qrprodutoPRECOVENDA1: TFloatField;
    qrprodutoPRECOVENDA2: TFloatField;
    qrprodutoPRECOVENDA3: TFloatField;
    qrprodutoPRECOVENDA4: TFloatField;
    qrprodutoPRECOVENDA5: TFloatField;
    qrprodutoDESCONTO_PRECOVENDA: TFloatField;
    qrprodutoDATA_INVENTARIO_ATUAL: TDateTimeField;
    qrprodutoCUSTO_INVENTARIO_ATUAL: TFloatField;
    qrprodutoESTOQUE_INVENTARIO_ATUAL: TFloatField;
    qrprodutoMARGEM_MINIMA: TFloatField;
    qrprodutoPISCOFINS: TStringField;
    qrprodutoREFERENCIA_FORNECEDOR: TStringField;
    qrprodutoCOMISSAO1: TFloatField;
    qrprodutoMARGEM_DESCONTO: TFloatField;
    qrprodutoTAMANHO: TStringField;
    qrprodutoCOR: TStringField;
    qrprodutoINCIDENCIA_PISCOFINS: TStringField;
    qrprodutoVEIC_CHASSI: TStringField;
    qrprodutoVEIC_SERIE: TStringField;
    qrprodutoVEIC_POTENCIA: TStringField;
    qrprodutoVEIC_PESO_LIQUIDO: TStringField;
    qrprodutoVEIC_PESO_BRUTO: TStringField;
    qrprodutoVEIC_TIPO_COMBUSTIVEL: TStringField;
    qrprodutoVEIC_RENAVAM: TStringField;
    qrprodutoVEIC_ANO_FABRICACAO: TIntegerField;
    qrprodutoVEIC_ANO_MODELO: TIntegerField;
    qrprodutoVEIC_TIPO: TStringField;
    qrprodutoVEIC_TIPO_PINTURA: TStringField;
    qrprodutoVEIC_COD_COR: TStringField;
    qrprodutoVEIC_COR: TStringField;
    qrprodutoVEIC_VIN: TStringField;
    qrprodutoVEIC_NUMERO_MOTOR: TStringField;
    qrprodutoVEIC_CMKG: TStringField;
    qrprodutoVEIC_CM3: TStringField;
    qrprodutoVEIC_DISTANCIA_EIXO: TStringField;
    qrprodutoVEIC_COD_MARCA: TStringField;
    qrprodutoVEIC_ESPECIE: TStringField;
    qrprodutoVEIC_CONDICAO: TStringField;
    qrprodutoLOTE_FABRICACAO: TStringField;
    qrprodutoLOTE_VALIDADE: TDateTimeField;
    qrprodutoMARGEM_AGREGADA: TFloatField;
    qrprodutoCODBARRA_NOVARTIS: TStringField;
    qrprodutoFARMACIA_DCB: TStringField;
    qrprodutoFARMACIA_ABCFARMA: TStringField;
    qrprodutoFARMACIA_APRESENTACAO_CAIXA: TStringField;
    qrprodutoFARMACIA_PRINCIPIOATIVO: TStringField;
    qrprodutoULTIMA_COMPRA: TDateTimeField;
    qrprodutoFARMACIA_DATAVIGENCIA: TDateTimeField;
    qrprodutoFARMACIA_TIPO: TStringField;
    qrprodutoUSA_COMBUSTIVEL: TStringField;
    qrprodutoREFERENCIA: TStringField;
    qrprodutoPERDA: TFloatField;
    qrprodutoCOMPOSICAO1: TStringField;
    qrprodutoCOMPOSICAO2: TStringField;
    qrprodutoIAT: TStringField;
    qrprodutoIPPT: TStringField;
    qrprodutoSITUACAO_TRIBUTARIA: TStringField;
    qrprodutoFLAG_SIS: TStringField;
    qrprodutoFLAG_ACEITO: TStringField;
    qrprodutoFLAG_EST: TStringField;
    qrprodutoCSOSN: TStringField;
    qrprodutoCODORIGINAL: TStringField;
    qrprodutoCUSTO_ATACADO: TFloatField;
    qrprodutoUNIDADE_ATACADO: TStringField;
    qrprodutoQTDE_EMBALAGEMATACADO: TFloatField;
    qrprodutoPMARGEM1: TFloatField;
    qrprodutoPMARGEM2: TFloatField;
    qrprodutoPMARGEM3: TFloatField;
    qrprodutoPMARGEM4: TFloatField;
    qrprodutoPMARGEM5: TFloatField;
    qrprodutoPMARGEMATACADO1: TFloatField;
    qrprodutoPMARGEMATACADO2: TFloatField;
    qrprodutoPMARGEMATACADO3: TFloatField;
    qrprodutoPMARGEMATACADO4: TFloatField;
    qrprodutoPMARGEMATACADO5: TFloatField;
    qrprodutoPMARGEMATACADO6: TFloatField;
    qrprodutoPRECOATACADO1: TFloatField;
    qrprodutoPRECOATACADO2: TFloatField;
    qrprodutoPRECOATACADO3: TFloatField;
    qrprodutoPRECOATACADO4: TFloatField;
    qrprodutoPRECOATACADO5: TFloatField;
    qrprodutoIND_CFOP: TStringField;
    qrprodutoCFOP_DESC: TStringField;
    qrprodutoUSA_LOTE: TIntegerField;
    qrprodutoIND_CFOP_VENDA_DENTRO: TStringField;
    qrprodutoCODCONTA: TStringField;
    qrprodutoIND_CFOP_VENDA_FORA: TStringField;
    qrprodutoIND_CFOP_DEVOLUCAO_DENTRO: TStringField;
    qrprodutoIND_CFOP_DEVOLUCAO_FORA: TStringField;
    qrprodutoIND_CFOP_GARANTIA_DENTRO: TStringField;
    qrprodutoIND_CFOP_GARANTIA_FORA: TStringField;
    qrprodutoUSA_TB_PC: TStringField;
    qrsubgrupoCODIGO: TStringField;
    qrsubgrupoCODGRUPO: TStringField;
    qrsubgrupoSUBGRUPO: TStringField;
    qrsubgrupoALTERA_PRECO_ABCFARMA: TStringField;
    qrsubgrupoCOMISSAO: TFloatField;
    qrsubgrupoDESCONTO: TFloatField;
    qrcontasreceberCODIGO: TStringField;
    qrcontasreceberCODVENDA: TStringField;
    qrcontasreceberCODVENDEDOR: TStringField;
    qrcontasreceberCODCAIXA: TStringField;
    qrcontasreceberCODCLIENTE: TStringField;
    qrcontasreceberDATA_EMISSAO: TDateTimeField;
    qrcontasreceberDATA_VENCIMENTO: TDateTimeField;
    qrcontasreceberDATA_PAGAMENTO: TDateTimeField;
    qrcontasreceberVALOR_ORIGINAL: TFloatField;
    qrcontasreceberVALOR_PAGO: TFloatField;
    qrcontasreceberVALOR_JUROS: TFloatField;
    qrcontasreceberVALOR_ATUAL: TFloatField;
    qrcontasreceberVALOR_DESCONTO: TFloatField;
    qrcontasreceberDOCUMENTO: TStringField;
    qrcontasreceberTIPO: TStringField;
    qrcontasreceberSITUACAO: TIntegerField;
    qrcontasreceberFILTRO: TIntegerField;
    qrcontasreceberNOSSONUMERO: TStringField;
    qrcontasreceberCODREGIAO: TStringField;
    qrcontasreceberCODCEDENTE: TStringField;
    qrcontasreceberP5: TFloatField;
    qrcontasreceberP3: TFloatField;
    qrcontasreceberNUMERO_CUPOM: TStringField;
    qrcontasreceberVALOR_VENDA: TFloatField;
    qrcontasreceberCOD_VENDA_ORIGINAL: TStringField;
    qrcontasreceberVALOR_ATUAL_ANTERIOR: TFloatField;
    qrcontasreceberEMAIL_ENVIADO: TIntegerField;
    qrcontasreceberHISTORICO: TStringField;
    qrconfigCODIGO: TStringField;
    qrconfigULTIMO_BACKUP: TDateTimeField;
    qrconfigPAPEL_PAREDE: TStringField;
    qrconfigAVISO_CONTAS_PAGAR: TIntegerField;
    qrconfigAVISO_CHEQUE: TIntegerField;
    qrconfigESTOQUE_NEGATIVO: TIntegerField;
    qrconfigCASAS_DECIMAIS_QTDE: TIntegerField;
    qrconfigCASAS_DECIMAIS_VALOR: TIntegerField;
    qrconfigVENDA_CADASTRO_CHEQUE: TIntegerField;
    qrconfigVENDA_CADASTRO_CARTAO: TIntegerField;
    qrconfigVENDA_TIPONOTA: TIntegerField;
    qrconfigVENDA_QTDE_VIAS_NOTA: TIntegerField;
    qrconfigVENDA_PORTA_IMPRESSORA: TStringField;
    qrconfigVENDA_TIPO_IMPRESSORA: TIntegerField;
    qrconfigVENDA_LIMITE_CLIENTE: TIntegerField;
    qrconfigOS_TIPONOTA: TIntegerField;
    qrconfigOS_QTDE_VIAS_NOTA: TIntegerField;
    qrconfigOS_PORTA_IMPRESSORA: TStringField;
    qrconfigOS_TIPO_IMPRESSORA: TIntegerField;
    qrconfigORCAMENTO_TIPONOTA: TIntegerField;
    qrconfigORCAMENTO_QTDE_VIAS_NOTA: TIntegerField;
    qrconfigORCAMENTO_PORTA_IMPRESSORA: TStringField;
    qrconfigORCAMENTO_TIPO_IMPRESSORA: TIntegerField;
    qrconfigCONTASRECEBER_CARENCIA: TIntegerField;
    qrconfigCONTASRECEBER_TAXAJUROS: TFloatField;
    qrconfigCONTASRECEBER_MULTA: TFloatField;
    qrconfigCONTASRECEBER_DESCONTO_PERCENT: TFloatField;
    qrconfigCONTASRECEBER_DESCONTO_PERMITIR: TIntegerField;
    qrconfigECF_MODELO: TStringField;
    qrconfigECF_PORTA: TStringField;
    qrconfigECF_TEF: TIntegerField;
    qrconfigECF_TIPO: TStringField;
    qrconfigBALANCAPDV_MODELO: TStringField;
    qrconfigBALANCAPDV_PORTA: TStringField;
    qrconfigIMPRESSORACHEQUE_MODELO: TStringField;
    qrconfigIMPRESSORACHEQUE_PORTA: TStringField;
    qrconfigBALANCA_MODELO: TStringField;
    qrconfigIMPRESSORABARRAS_MODELO: TStringField;
    qrconfigBALANCA_CAMINHO: TStringField;
    qrconfigCADASTRO_PRODUTO: TStringField;
    qrconfigRAMO_ATIVIDADE: TIntegerField;
    qrconfigEXTRATO_TIPONOTA: TIntegerField;
    qrconfigBOBINA_SUBIRPAPEL: TIntegerField;
    qrconfigCLASSIFICACAO: TStringField;
    qrconfigBOLETO_FORMAPGTO: TStringField;
    qrconfigPLANO_VENDA_AV: TStringField;
    qrconfigPLANO_VENDA_AP: TStringField;
    qrconfigPLANO_OS_AV: TStringField;
    qrconfigPLANO_OS_AP: TStringField;
    qrconfigPLANO_OUTRAS_ENTRADAS: TStringField;
    qrconfigPLANO_OUTRAS_SAIDAS: TStringField;
    qrconfigPLANO_RECEBTO_CREDIARIO: TStringField;
    qrconfigNF_ITENS_PAGINA: TIntegerField;
    qrconfigNF_ITENS_TRANSPORTE: TIntegerField;
    qrconfigNF_LINHA_TRANSPORTE: TIntegerField;
    qrconfigNF_SALTO_PAGINA: TIntegerField;
    qrconfigNF_TIPO: TIntegerField;
    qrconfigNF_IMPRESSORA: TStringField;
    qrconfigVENDA_FECHAMENTO_CAIXA: TIntegerField;
    qrconfigDESCONTO_AVISTA: TFloatField;
    qrconfigDESCONTO_PRODUTO: TFloatField;
    qrconfigDESCONTO_TOTALNOTA: TFloatField;
    qrconfigTIPO_VENDA: TStringField;
    qrconfigVENDA_ALTERAR_PRECO: TIntegerField;
    qrconfigNF_LINHA_SUBTOTAL: TIntegerField;
    qrconfigNF_LINHA_TOTAL: TIntegerField;
    qrconfigECF_CAIXA: TStringField;
    qrconfigECF_SERIAL: TStringField;
    qrconfigFARMACIA_ENVIO: TStringField;
    qrconfigFARMACIA_RESPOSTA: TStringField;
    qrconfigFARMACIA_ENVIO_RESPOSTA: TStringField;
    qrconfigFARMACIA_LOGIN: TStringField;
    qrconfigFARMACIA_SENHA: TStringField;
    qrconfigULTIMA_CARGA: TDateTimeField;
    qrconfigIMPRIME_CUPOM_VINCULADO: TIntegerField;
    qrconfigIMPRIME_PRODUTOS_VINCULADOS: TIntegerField;
    qrconfigIMPRIME_COMPROVANTE_DEBITO: TIntegerField;
    qrconfigIMPRIME_DUPLICATA: TIntegerField;
    qrconfigCONTASRECEBER_BLOQUEIO: TIntegerField;
    qrconfigRECIBO_TIPONOTA: TIntegerField;
    qrconfigNR_CONTROLE_ECF: TStringField;
    qrconfigBALANCA_PROGRAMA: TStringField;
    qrconfigTIPO_DUPLICATA: TStringField;
    qrconfigCONCILIA_VENDA: TIntegerField;
    qrconfigCONTROLA_ENTREGA: TIntegerField;
    qrconfigENTREGA_IMPRESSAO: TIntegerField;
    qrconfigENTREGA_VIAS: TIntegerField;
    qrconfigCONDICIONAL: TIntegerField;
    qrconfigCONDICIONAL_PRAZO: TIntegerField;
    qrconfigVS_ATIVAR: TIntegerField;
    qrconfigVS_DINHEIRO: TIntegerField;
    qrconfigVS_CHEQUEAV: TIntegerField;
    qrconfigVS_CHEQUEAP: TIntegerField;
    qrconfigVS_CARTRAOCRED: TIntegerField;
    qrconfigVS_CARTAODEB: TIntegerField;
    qrconfigVS_CREDIARIO: TIntegerField;
    qrconfigVS_CLIENTE_VISTA: TStringField;
    qrconfigVS_CAIXA: TIntegerField;
    qrconfigVS_RELATORIO: TIntegerField;
    qrconfigVS_VIAS: TIntegerField;
    qrconfigVS_VISUALIZAR: TIntegerField;
    qrconfigCONDICIONAL_MODELO: TIntegerField;
    qrconfigALTERAR_UNITARIO: TIntegerField;
    qrconfigCONTASRECEBER_VALORJUROS: TFloatField;
    qrconfigCONTASRECEBER_JUROS_DINHEIRO: TIntegerField;
    qrconfigPLANO_DUPLICATA_NFE: TStringField;
    qrconfigVAREJO_ATACADO: TIntegerField;
    qrconfigVISUALIZA_MENSAGEM: TIntegerField;
    qrconfigSERIAL_REPETIDO: TIntegerField;
    qrconfigOCULTAR_CUSTO: TIntegerField;
    qrconfigOCULTAR_PRODUTO_INATIVO: TIntegerField;
    qrconfigOCULTAR_CLIENTE_INATIVO: TIntegerField;
    qrconfigOBS_FARMA: TIntegerField;
    qrconfigOBS_FARMA_TEXTO: TStringField;
    qrconfigCUSTO_MAIOR_VENDA: TIntegerField;
    qrconfigVENDER_ABAIXO_CUSTO: TIntegerField;
    qrconfigNF_GERAR_RECEBER: TIntegerField;
    qrconfigNF_CODIGOBARRA_NFE: TIntegerField;
    qrconfigNF_GRADE_NFE: TIntegerField;
    qrconfigNF_SERIAL_NFE: TIntegerField;
    qrconfigNF_REFERENCIA_NFE: TIntegerField;
    qrconfigNF_GERAR_DUPLICATA: TIntegerField;
    qrconfigVALE_CAIXA: TIntegerField;
    qrconfigCONTAS_RECEBER_BLOQUEIO: TIntegerField;
    qrconfigCUPOM_IMPRIMIR_CODIGO: TIntegerField;
    qrconfigCLIENTE_PEDE_CPF: TIntegerField;
    qrconfigVERIFICA_COD_BARRA: TIntegerField;
    qrconfigVERIFICA_NOME_PRODUTO: TIntegerField;
    qrconfigVS_FRETE: TIntegerField;
    qrconfigINVENTARIO_AUTOMATICO: TIntegerField;
    qrconfigBAIXAR_ESTOQUE_NOTA_SAIDA: TIntegerField;
    qrconfigETIQUETA_07: TIntegerField;
    qrconfigETIQUETA_07_1: TIntegerField;
    qrconfigETIQUETA_07_2: TIntegerField;
    qrconfigEMAIL_SERVIDOR_SMTP: TStringField;
    qrconfigEMAIL_PORTA: TStringField;
    qrconfigEMAIL_USUARIO: TStringField;
    qrconfigEMAIL_SENHA: TStringField;
    qrconfigEMAIL_ASSUNTO: TStringField;
    qrconfigEMAIL_EXIGE_CONEXAO: TIntegerField;
    qrconfigEMAIL_MENSAGEM: TBlobField;
    qrconfigINESTRA: TIntegerField;
    qrconfigFARMACIA_INVENTARIO_FECHA: TIntegerField;
    qrconfigFARMACIA_ENVIO_NOVARTIS: TStringField;
    qrconfigFARMACIA_RESPOSTA_NOVARTIS: TStringField;
    qrconfigFARMACIA_INVENTARIO_CONFIRMA: TStringField;
    qrconfigFARMACIA_INVENTARIO_DATA: TDateTimeField;
    qrconfigCALCULA_ISENTAS_OUTRAS: TIntegerField;
    qrconfigUSA_RENTABILIDADE: TIntegerField;
    qrconfigUSA_PRODUTOR_RURAL: TIntegerField;
    qrconfigIND_CLIE_ENTRADA: TStringField;
    qrconfigIND_CFOP_ENTRADA: TStringField;
    qrconfigIND_CLIE_SAIDA: TStringField;
    qrconfigIND_CFOP_SAIDA: TStringField;
    qrconfigIND_CLIE_PERDA: TStringField;
    qrconfigIND_CFOP_PERDA: TStringField;
    qrconfigIND_MOSTRA_COMPOSICAO: TIntegerField;
    qrconfigCHECA_SERIAL_NOTA_SAIDA: TIntegerField;
    qrconfigCOR_COMERCIO: TStringField;
    qrconfigLOGO_COMERCIO: TStringField;
    qrconfigCOR_FARMACIA: TStringField;
    qrconfigLOGO_FARMACIA: TStringField;
    qrconfigCOR_PECAS: TStringField;
    qrconfigLOGO_PECAS: TStringField;
    qrconfigCOR_MERCADO: TStringField;
    qrconfigLOGO_MERCADO: TStringField;
    qrconfigCOR_POSTO: TStringField;
    qrconfigLOGO_POSTO: TStringField;
    qrconfigIND_CLIE_ENTRADA_DESC: TStringField;
    qrconfigIND_CFOP_ENTRADA_DESC: TStringField;
    qrconfigIND_CLIE_SAIDA_DESC: TStringField;
    qrconfigIND_CFOP_SAIDA_DESC: TStringField;
    qrconfigIND_CLIE_PERDA_DESC: TStringField;
    qrconfigIND_CFOP_PERDA_DESC: TStringField;
    qrconfigIND_CFOP_VENDA_DENTRO: TStringField;
    qrconfigIND_CFOP_VENDA_FORA: TStringField;
    qrconfigIND_CFOP_DEVOLUCAO_DENTRO: TStringField;
    qrconfigIND_CFOP_DEVOLUCAO_FORA: TStringField;
    qrconfigIND_CFOP_SERVICO_DENTRO: TStringField;
    qrconfigIND_CFOP_SERVICO_FORA: TStringField;
    qrconfigCFOP_VENDA_DENTRO_DESC: TStringField;
    qrconfigCFOP_VENDA_FORA_DESC: TStringField;
    qrconfigCFOP_DEVOLUCAO_DENTRO_DESC: TStringField;
    qrconfigCFOP_DEVOLUCAO_FORA_DESC: TStringField;
    qrconfigCFOP_SERVICO_DENTRO_DESC: TStringField;
    qrconfigCFOP_SERVICO_FORA_DESC: TStringField;
    qrconfigTIPO_NF: TIntegerField;
    qrconfigCOR_ESTOQUE_POSITIVO: TStringField;
    qrconfigCOR_ESTOQUE_NEGATIVO: TStringField;
    qrconfigCOR_ESTOQUE_ZERADO: TStringField;
    qrconfigTIPO_PESQUISA: TIntegerField;
    qrconfigECF_CONCOMITANTE: TIntegerField;
    qrconfigPLANO_PAGTO_FORN: TStringField;
    qrconfigHAB_PREVENDA: TIntegerField;
    qrconfigCODCLIENTEAV: TStringField;
    qrconfigCODCLIENTEMODELO: TStringField;
    qrconfigCODPRODUTOMODELO: TStringField;
    qrconfigCODFORNECEDORMODELO: TStringField;
    qrconfigUSA_CAIXA_DIARIO: TIntegerField;
    qrconfigIMPRIMI_ORCAMENTO_AUTO: TIntegerField;
    qrconfigFIXA_FORMA_PGTO_CLIENTE: TIntegerField;
    qrconfigVERSAOBD: TStringField;
    qrconfigUSA_CAIXA_PADRAO_USUARIO: TIntegerField;
    qrconfigTEXTO_PADRAO_VENDA: TStringField;
    qrconfigTEXTO_PADRAO_DEVOLUCAO: TStringField;
    qrconfigTEXTO_PADRAO_GARANTIA: TStringField;
    qrconfigAVISA_ESTOQUE_MINIMO: TIntegerField;
    qrcontacorrenteCODIGO: TStringField;
    qrcontacorrenteCONTA: TStringField;
    qrcontacorrenteCODBANCO: TStringField;
    qrcontacorrenteAGENCIA: TStringField;
    qrcontacorrenteSALDO: TFloatField;
    qrcontacorrenteTITULAR: TStringField;
    qrcontacorrenteNOME_AGENCIA: TStringField;
    qrcontacorrenteCONTADIG: TStringField;
    qrcontacorrenteAGENCIADIG: TStringField;
    qrcontacorrenteCODCEDENTE: TStringField;
    qrcontacorrenteFISICAJURIDICA: TStringField;
    qrcontacorrenteINSTRUCOES1: TStringField;
    qrcontacorrenteINSTRUCOES2: TStringField;
    qrcontacorrenteINSTRUCOES3: TStringField;
    qrcontacorrenteCPFCNPJ: TStringField;
    qrcontacorrenteNOSSONUM: TStringField;
    qrcontacorrenteTIPO: TIntegerField;
    qrcontacorrenteLAYOUT: TIntegerField;
    qrcontacorrenteCARTEIRA: TStringField;
    qrcfopCFOP: TStringField;
    qrcfopNATUREZA: TStringField;
    qrcfopTIPO: TIntegerField;
    qrcfopICMS: TFloatField;
    qrcfopISS: TFloatField;
    qrcfopOBS1: TStringField;
    qrcfopOBS2: TStringField;
    qrcfopOBS3: TStringField;
    qrcfopOBS4: TStringField;
    qrcfopFILTRO: TIntegerField;
    qrcfopVALOR1: TFloatField;
    qrcfopVALOR2: TFloatField;
    qrcfopVALOR3: TFloatField;
    qrcfopVALOR4: TFloatField;
    qrcfopVALOR5: TFloatField;
    qrcfopVALOR6: TFloatField;
    qrcfopVALOR7: TFloatField;
    qrcfopVALOR8: TFloatField;
    qrcfopVALOR9: TFloatField;
    qrcfopVALOR10: TFloatField;
    qrcfopST: TStringField;
    qrcfopCLASSIFICACAO_FISCAL: TStringField;
    qrcfopREDUCAO_ICMS: TFloatField;
    qrcfopMARGEM_AGREGADA: TFloatField;
    qrcfopSIMPLIFICADO: TStringField;
    qrcfopCAIXA: TIntegerField;
    qrcfopFATURAMENTO: TIntegerField;
    qrmarcaCODIGO: TStringField;
    qrmarcaNOME: TStringField;
    qrmarcaENDERECO: TStringField;
    qrmarcaBAIRRO: TStringField;
    qrmarcaCIDADE: TStringField;
    qrmarcaUF: TStringField;
    qrmarcaCEP: TStringField;
    qrmarcaTELEFONE1: TStringField;
    qrmarcaTELEFONE2: TStringField;
    qrmarcaFAX: TStringField;
    qrmarcaCONTATO1: TStringField;
    qrmarcaCONTATO2: TStringField;
    qrmarcaCELULAR1: TStringField;
    qrmarcaCELULAR2: TStringField;
    qrmarcaEMAIL: TStringField;
    qrmarcaHOMEPAGE: TStringField;
    qrmarcaCNPJ: TStringField;
    qrmarcaIE: TStringField;
    qrmarcaDATA: TDateTimeField;
    qrmarcaCOMISSAO: TFloatField;
    qrmarcaNUMERO: TStringField;
    qrgrupoCODIGO: TStringField;
    qrgrupoGRUPO: TStringField;
    qrgrupoALTERA_PRECO_ABCFARMA: TStringField;
    qrgrupoCOMISSAO: TFloatField;
    qrgrupoDESCONTO: TFloatField;
    qrclienteCODIGO: TStringField;
    qrclienteNOME: TStringField;
    qrclienteAPELIDO: TStringField;
    qrclienteENDERECO: TStringField;
    qrclienteBAIRRO: TStringField;
    qrclienteCIDADE: TStringField;
    qrclienteUF: TStringField;
    qrclienteCEP: TStringField;
    qrclienteCOMPLEMENTO: TStringField;
    qrclienteMORADIA: TIntegerField;
    qrclienteTIPO: TIntegerField;
    qrclienteSITUACAO: TIntegerField;
    qrclienteTELEFONE1: TStringField;
    qrclienteTELEFONE2: TStringField;
    qrclienteTELEFONE3: TStringField;
    qrclienteCELULAR: TStringField;
    qrclienteEMAIL: TStringField;
    qrclienteRG: TStringField;
    qrclienteCPF: TStringField;
    qrclienteFILIACAO: TStringField;
    qrclienteESTADOCIVIL: TStringField;
    qrclienteCONJUGE: TStringField;
    qrclientePROFISSAO: TStringField;
    qrclienteEMPRESA: TStringField;
    qrclienteRENDA: TFloatField;
    qrclienteLIMITE: TFloatField;
    qrclienteREF1: TStringField;
    qrclienteREF2: TStringField;
    qrclienteCODVENDEDOR: TStringField;
    qrclienteDATA_CADASTRO: TDateTimeField;
    qrclienteDATA_ULTIMACOMPRA: TDateTimeField;
    qrclienteOBS1: TStringField;
    qrclienteOBS2: TStringField;
    qrclienteOBS3: TStringField;
    qrclienteOBS4: TStringField;
    qrclienteOBS5: TStringField;
    qrclienteOBS6: TStringField;
    qrclienteNASCIMENTO: TStringField;
    qrclienteCODREGIAO: TStringField;
    qrclienteCODCONVENIO: TStringField;
    qrclienteCODUSUARIO: TStringField;
    qrclienteNUMERO: TStringField;
    qrclienteRG_ORGAO: TStringField;
    qrclienteRG_ESTADO: TStringField;
    qrclienteRG_EMISSAO: TDateTimeField;
    qrclienteSEXO: TStringField;
    qrclienteHISTORICO: TBlobField;
    qrclientePREVISAO: TDateTimeField;
    qrclienteCNAE: TStringField;
    qrclienteCOD_MUNICIPIO_IBGE: TStringField;
    qrclienteIBGE: TStringField;
    qrclienteTAMANHO_CALCA: TStringField;
    qrclienteTAMANHO_BLUSA: TStringField;
    qrclienteTAMANHO_SAPATO: TStringField;
    qrclienteCORRESP_ENDERECO: TStringField;
    qrclienteCORRESP_BAIRRO: TStringField;
    qrclienteCORRESP_CIDADE: TStringField;
    qrclienteCORRESP_UF: TStringField;
    qrclienteCORRESP_CEP: TStringField;
    qrclienteCORRESP_COMPLEMENTO: TStringField;
    qrclienteRG_PRODUTOR: TStringField;
    qrclienteRESP1_NOME: TStringField;
    qrclienteRESP1_CPF: TStringField;
    qrclienteRESP1_RG: TStringField;
    qrclienteRESP1_PROFISSAO: TStringField;
    qrclienteRESP1_ESTADO_CIVIL: TStringField;
    qrclienteRESP1_ENDERECO: TStringField;
    qrclienteRESP1_BAIRRO: TStringField;
    qrclienteRESP1_CIDADE: TStringField;
    qrclienteRESP1_UF: TStringField;
    qrclienteRESP1_CEP: TStringField;
    qrclienteRESP2_NOME: TStringField;
    qrclienteRESP2_CPF: TStringField;
    qrclienteRESP2_RG: TStringField;
    qrclienteRESP2_PROFISSAO: TStringField;
    qrclienteRESP2_ESTADO_CIVIL: TStringField;
    qrclienteRESP2_ENDERECO: TStringField;
    qrclienteRESP2_BAIRRO: TStringField;
    qrclienteRESP2_CIDADE: TStringField;
    qrclienteRESP2_UF: TStringField;
    qrclienteRESP2_CEP: TStringField;
    qrclienteFOTO: TStringField;
    qrclienteCONDPGTO: TStringField;
    qrfiscal_modeloCODIGO: TStringField;
    qrfiscal_modeloMODELO: TStringField;
    qrfiscal_modeloSIGLA: TStringField;
    qrfiscal_modeloSINTEGRA: TStringField;
    qrfiscal_modeloTIPO_REGISTRO: TStringField;
    qrfiscal_modeloSEQUENCIANF: TIntegerField;
    qrgrade_produtoCODIGO: TStringField;
    qrgrade_produtoCODNOTA: TStringField;
    qrgrade_produtoCODPRODUTO: TStringField;
    qrgrade_produtoUNITARIO: TFloatField;
    qrgrade_produtoTOTAL: TFloatField;
    qrgrade_produtoICMS: TFloatField;
    qrgrade_produtoIPI: TFloatField;
    qrgrade_produtoCFOP: TStringField;
    qrgrade_produtoDATA: TDateTimeField;
    qrgrade_produtoNUMERONOTA: TStringField;
    qrgrade_produtoCODCLIENTE: TStringField;
    qrgrade_produtoDESCONTO: TFloatField;
    qrgrade_produtoACRESCIMO: TFloatField;
    qrgrade_produtoMOVIMENTO: TIntegerField;
    qrgrade_produtoCODVENDEDOR: TStringField;
    qrgrade_produtoCODGRADE: TStringField;
    qrgrade_produtoSERIAL: TStringField;
    qrgrade_produtoUNIDADE: TStringField;
    qrgrade_produtoQTDE: TFloatField;
    qrgrade_produtoVALOR_ICMS: TFloatField;
    qrgrade_produtoICMS_REDUZIDO: TFloatField;
    qrgrade_produtoBASE_CALCULO: TFloatField;
    qrgrade_produtoVALOR_IPI: TFloatField;
    qrgrade_produtoSITUACAO: TIntegerField;
    qrgrade_produtoECF_SERIE: TStringField;
    qrgrade_produtoECF_CAIXA: TStringField;
    qrgrade_produtoCODALIQUOTA: TStringField;
    qrgrade_produtoCUPOM_NUMERO: TStringField;
    qrgrade_produtoCUPOM_MODELO: TStringField;
    qrgrade_produtoCUPOM_ITEM: TStringField;
    qrgrade_produtoALIQUOTA: TFloatField;
    qrgrade_produtoCST: TStringField;
    qrgrade_produtoLOTE_FABRICACAO: TStringField;
    qrgrade_produtoMOVIMENTO_ESTOQUE: TFloatField;
    qrgrade_produtoLANCADO: TIntegerField;
    qrgrade_produtoVENCIMENTO: TDateTimeField;
    qrgrade_produtoCODBARRA: TStringField;
    qrgrade_produtoMARGEM_DESCONTO: TFloatField;
    qrgrade_produtoCREDITO_ICMS: TFloatField;
    qrgrade_produtoPIS: TFloatField;
    qrgrade_produtoCOFINS: TFloatField;
    qrgrade_produtoLOJA: TStringField;
    qrgrade_produtoCODSUBGRUPO: TStringField;
    qrgrade_produtoTIPO: TStringField;
    qrgrade_produtoCODUSUARIO: TStringField;
    qrgrade_produtoORIGEM: TStringField;
    qrgrade_produtoDESTINO: TStringField;
    qrgrade_produtoPRODUTO: TStringField;
    qrgrade_produtoCODFILIAL: TStringField;
    ZFilial: TUniQuery;
    ZFilialCODIGO: TStringField;
    ZFilialFILIAL: TStringField;
    ZFilialNOTAFISCAL: TIntegerField;
    ZFilialENDERECO: TStringField;
    ZFilialCIDADE: TStringField;
    ZFilialUF: TStringField;
    ZFilialCEP: TStringField;
    ZFilialTELEFONE: TStringField;
    ZFilialCNPJ: TStringField;
    ZFilialIE: TStringField;
    ZFilialSEQNF: TIntegerField;
    ZFilialFAX: TStringField;
    ZFilialOBS1: TStringField;
    ZFilialOBS2: TStringField;
    ZFilialCONTRIBUINTE_IPI: TStringField;
    ZFilialSUBSTITUTO_TRIBUTARIO: TStringField;
    ZFilialEMPRESA_ESTADUAL: TStringField;
    ZFilialOPTANTE_SIMPLES: TStringField;
    ZFilialOPTANTE_SUPER_SIMPLES: TStringField;
    ZFilialECF: TStringField;
    ZFilialTIPO: TIntegerField;
    ZFilialIPI: TFloatField;
    ZFilialISS: TFloatField;
    ZFilialNUMERO: TStringField;
    ZFilialRESPONSAVEL: TStringField;
    ZFilialCOMPLEMENTO: TStringField;
    ZFilialBAIRRO: TStringField;
    ZFilialFARMACIA_RESP_TECNICO: TStringField;
    ZFilialFARMACIA_CRF: TStringField;
    ZFilialFARMACIA_CPF: TStringField;
    ZFilialFARMCIA_DATA: TDateTimeField;
    ZFilialFARMACIA_UF: TStringField;
    ZFilialFARMACIA_SENHA: TStringField;
    ZFilialFARMACIA_EMAIL: TStringField;
    ZFilialFARMACIA_LOGIN: TStringField;
    ZFilialFARMACIA_ENVIO: TStringField;
    ZFilialCONHECIMENTO: TIntegerField;
    ZFilialINDUSTRIA: TStringField;
    ZFilialFARMACIA_NUMEROLICENCA: TStringField;
    ZFilialCOD_MUNICIPIO_IBGE: TStringField;
    ZFilialIBGE: TStringField;
    ZFilialPIS: TFloatField;
    ZFilialCOFINS: TFloatField;
    ZFilialEMAIL: TStringField;
    ZFilialATIVIDADE: TStringField;
    ZFilialCONTADOR_COD_MUNICIPIO_IBGE: TStringField;
    ZFilialCONTADOR_NOME: TStringField;
    ZFilialCONTADOR_CPF: TStringField;
    ZFilialCONTADOR_CRC: TStringField;
    ZFilialCONTADOR_CNPJ: TStringField;
    ZFilialCONTADOR_CEP: TStringField;
    ZFilialCONTADOR_ENDERECO: TStringField;
    ZFilialCONTADOR_NUMERO: TStringField;
    ZFilialCONTADOR_COMPLEMENTO: TStringField;
    ZFilialCONTADOR_BAIRRO: TStringField;
    ZFilialCONTADOR_FONE: TStringField;
    ZFilialCONTADOR_FAX: TStringField;
    ZFilialCONTADOR_EMAIL: TStringField;
    ZFilialINSC_MUNICIPAL: TStringField;
    ZFilialDATA_ABERTURA: TDateTimeField;
    ZFilialCNAE: TStringField;
    ZFilialCRT: TStringField;
    ZFilialCONTADOR_CIDADE: TStringField;
    ZFilialCONTADOR_COD_MUNICIPIO: TStringField;
    ZFilialCONTADOR_UF: TStringField;
    ZFilialPERMITE_CREDITO: TIntegerField;
    ZFilialFANTASIA: TStringField;
    ZFilialDFIXAS: TFloatField;
    ZFilialTIPOCALCULO: TIntegerField;
    ZFilialSERIE_CTE: TStringField;
    ZFilialSEQ_CTE: TStringField;
    ZFilialCOD_PAIS: TStringField;
    ZFilialPAIS: TStringField;
    ACBrEAD1: TACBrEAD;
    ACBRNFCe: TACBrNFe;
    qradic_mestre: TUniQuery;
    qrItens: TUniQuery;
    qrPegoPor: TUniQuery;
    qrPegoPorCODCLIENTE: TStringField;
    qrPegoPorNOME: TStringField;
    qrMesas: TUniQuery;
    qrMesasCODIGO: TIntegerField;
    qrMesasDATA: TDateField;
    qrMesasHORA: TStringField;
    qrMesasECF: TStringField;
    qrMesasCOO: TStringField;
    qrMesasCER: TStringField;
    qrMesasCOD_FUNCIONARIO: TStringField;
    qrMesasPERGUNTOU_COMISSAO: TStringField;
    qrMesasCLIENTE_AUTORIZOU: TStringField;
    qrconfigATENDIMENTOIMPRESSAO: TStringField;
    qrconfigATENDIMENTOVIAS: TIntegerField;
    qrconfigATENDIMENTOFICHA: TStringField;
    qrconfigATENDIMENTOIMPRESSORA: TStringField;
    qrconfigATENDIMENTOIMPRESMODELO: TIntegerField;
    qrconfigPERMITE_PROD_PRECO_ZERO: TStringField;
    qrconfigACRESCIMO_PRODUTO: TFloatField;
    qrconfigNAO_SOL_SENHA_MOV_ESTOQUE: TStringField;
    qrconfigCARTA_COBRANCA_TEXTO: TBlobField;
    qrconfigCOBRA_COUVER: TStringField;
    qrconfigVALOR_COUVER: TFloatField;
    qrconfigCOBRA_COMISSAO: TStringField;
    qrconfigCOMISSAO_PERCENTUAL: TFloatField;
    qrDelivery: TUniQuery;
    qrDeliveryID: TIntegerField;
    qrDeliveryNUMERO: TIntegerField;
    qrDeliveryDATA_EMISSAO: TDateField;
    qrDeliveryHORA_EMISSAO: TStringField;
    qrDeliveryNOME: TStringField;
    qrDeliveryENTREGADOR_NOME: TStringField;
    qrDeliveryTIPO_ENTREGA: TStringField;
    qrDeliveryTAXA_DESCRICAO: TStringField;
    qrDeliveryTAXA_VALOR: TFloatField;
    qrDeliveryDATA_ENTREGA: TDateField;
    qrDeliveryHORA_ENTREGA: TStringField;
    qrDeliveryVALOR_TOTAL: TFloatField;
    qrDeliveryVALOR_TAXA_ENTREGA: TFloatField;
    qrDeliveryVALOR_LIQUIDO: TFloatField;
    qrDeliveryLEVAR_MAQUINA_CARTAO: TStringField;
    qrImpDelivery: TUniQuery;
    qrImpDeliveryID_DELIVERY: TIntegerField;
    qrImpDeliveryPRODUTO: TStringField;
    qrImpDeliveryCST: TStringField;
    qrImpDeliveryALIQUOTA: TFloatField;
    qrImpDeliveryCODPRODUTO: TStringField;
    qrImpDeliveryCODBARRA: TStringField;
    qrImpDeliveryQUANTIDADE: TFloatField;
    qrImpDeliveryUNITARIO: TFloatField;
    qrImpDeliveryTOTAL: TFloatField;
    qrFechaMesas: TUniQuery;
    qrTemp: TUniQuery;
    qrEstoque: TUniQuery;
    qrEstoqueCODIGO: TIntegerField;
    qrEstoqueCOD_BARRA: TStringField;
    qrEstoqueNOME: TStringField;
    qrEstoqueUNIDADE: TStringField;
    qrEstoquePRECO_VENDA: TFloatField;
    qrEstoquePRECO_PROMOCAO: TFloatField;
    qrEstoqueINICIO_PROMOCAO: TDateField;
    qrEstoqueFINAL_PROMOCAO: TDateField;
    qrEstoqueCST: TStringField;
    qrEstoqueALIQUOTA: TFloatField;
    qrEstoqueDESCONTO_MAXIMO: TFloatField;
    qrEstoqueST: TStringField;
    qrEstoqueESTOQUE: TFloatField;
    qrEstoqueIAT: TStringField;
    qrEstoqueIPPT: TStringField;
    qrEstoqueSITUACAO: TIntegerField;
    qrEstoqueREFERENCIA: TStringField;
    qrEstoqueTAMANHO: TStringField;
    qrEstoqueCOR: TStringField;
    qrEstoqueDESCRICAO_COR: TStringField;
    qrEstoqueDESCRICAO_MARCA: TStringField;
    qrEstoqueNCM: TStringField;
    qrEstoquePRECO_VENDA2: TFloatField;
    qrEstoquePRECO_VENDA3: TFloatField;
    qrEstoquePRECO_VENDA4: TFloatField;
    qrEstoqueALIQNACIONAL: TFloatField;
    qrEstoqueCFOP: TStringField;
    qrEstoqueFOTOBD: TBlobField;
    qrEstoqueACRECIMO_MAXIMO: TFloatField;
    qrEstoqueCSOSN: TStringField;
    qrItensCODIGO: TStringField;
    qrItensCOD_CUPOM: TStringField;
    qrItensITEM: TIntegerField;
    qrItensCOD_PRODUTO: TIntegerField;
    qrItensUNIDADE: TStringField;
    qrItensQTDE: TFloatField;
    qrItensVALOR_UNITARIO: TFloatField;
    qrItensVALOR_SUBTOTAL: TFloatField;
    qrItensVALOR_DESCONTO: TFloatField;
    qrItensVALOR_ACRESCIMO: TFloatField;
    qrItensVALOR_TOTAL: TFloatField;
    qrItensCST: TStringField;
    qrItensALIQUOTA: TFloatField;
    qrItensCOD_TOTALIZADOR: TStringField;
    qrItensCANCELADO: TIntegerField;
    qrItensTAMANHO: TStringField;
    qrItensCOR: TStringField;
    qrItensCOD_BARRA: TStringField;
    qrItensCODIGO_1: TIntegerField;
    qrItensNOME: TStringField;
    qrSenhaPDV: TUniQuery;
    qrSenhaPDVSENHA: TLargeintField;
    qrItensEXTRA: TBlobField;
    qrconfigQTD_MESAS_MOBILE: TIntegerField;
    qrconfigPRODUTO_TAXA_ENTREGA: TStringField;
    qrconfigTIP_IMP_DELIVERY: TIntegerField;
    qrconfigULTIMAVERIFICAOXML: TDateTimeField;
    qrNFCEInsert: TUniQuery;
    qrCupomCancelar: TUniQuery;
    qrCupom_Cartao: TUniQuery;
    qrCupom_CartaoCODIGO: TStringField;
    qrCupom_CartaoCODIGO_CUPOM: TStringField;
    qrCupom_CartaoBANCO: TStringField;
    qrCupom_CartaoTIPO_CARTAO: TStringField;
    qrCupom_CartaoPRESTACAO: TIntegerField;
    qrCupom_CartaoVALOR: TFloatField;
    qrCupom_CartaoVENCIMENTO: TDateField;
    qrCupom_CartaoEX: TStringField;
    qrCupomAJ: TUniQuery;
    qrCupomAJNUMERO: TStringField;
    qrCupomAJCCF: TStringField;
    qrCupomAJECF: TStringField;
    qrCupomAJDATA: TDateField;
    qrCupomAJHORA: TTimeField;
    qrCupomAJTOTALIZADO: TIntegerField;
    qrCupomAJPAGO1: TIntegerField;
    qrCupomAJPAGO2: TIntegerField;
    qrCupomAJPAGO3: TIntegerField;
    qrCupomAJFINALIZADO: TIntegerField;
    qrItemAJ: TUniQuery;
    qrItemAJCODIGO: TStringField;
    qrItemAJCOD_CUPOM: TStringField;
    qrItemAJITEM: TIntegerField;
    qrItemAJCOD_PRODUTO: TIntegerField;
    qrItemAJUNIDADE: TStringField;
    qrItemAJQTDE: TFloatField;
    qrItemAJVALOR_UNITARIO: TFloatField;
    qrItemAJVALOR_SUBTOTAL: TFloatField;
    qrItemAJVALOR_DESCONTO: TFloatField;
    qrItemAJVALOR_ACRESCIMO: TFloatField;
    qrItemAJVALOR_TOTAL: TFloatField;
    qrItemAJCST: TStringField;
    qrItemAJALIQUOTA: TFloatField;
    qrItemAJCOD_TOTALIZADOR: TStringField;
    qrItemAJCANCELADO: TIntegerField;
    qrItemAJTAMANHO: TStringField;
    qrItemAJCOR: TStringField;
    qrItemAJEXTRA: TBlobField;
    qrConvenio: TUniQuery;
    qrConvenioCODIGO: TStringField;
    qrConvenioNOME: TStringField;
    qrConvenioLIMITE: TFloatField;
    qrConvenioUTILIZADO: TFloatField;
    qrConvenioSALDO: TFloatField;
    qrConvenioATU_DATA: TDateField;
    qrConvenioATU_HORA: TStringField;
    qrCupomConvenio: TUniQuery;
    qrCupomConvenioCODIGO: TStringField;
    qrCupomConvenioCOD_CUPOM: TStringField;
    qrCupomConvenioDATA: TDateField;
    qrCupomConvenioHORA: TTimeField;
    qrCupomConvenioCOD_CLIENTE: TIntegerField;
    qrCupomConvenioVENCIMENTO: TDateField;
    qrCupomConvenioVALOR: TFloatField;
    qrCupomConvenioEX: TIntegerField;
    qrConvenioDESCONTO: TFloatField;
    qryCliPDV: TUniQuery;
    qryCliPDVCODIGO: TIntegerField;
    qryCliPDVNOME: TStringField;
    qryCliPDVCPF: TStringField;
    qryCliPDVENDERECO: TStringField;
    qryCliPDVCOMPLEMENTO: TStringField;
    qryCliPDVBAIRRO: TStringField;
    qryCliPDVCIDADE: TStringField;
    qryCliPDVUF: TStringField;
    qryCliPDVCEP: TStringField;
    qryCliPDVSITUACAO: TIntegerField;
    qryCliPDVOBS: TBlobField;
    qryCliPDVLIMITE: TFloatField;
    qryCliPDVUTILIZADO: TFloatField;
    qryCliPDVDISPONIVEL: TFloatField;
    qryCliPDVATUALIZADO: TStringField;
    qryCliPDVPRECOAPRAZO: TIntegerField;
    qryCliPDVPRECOAVISTA: TIntegerField;
    qryCliPDVCOD_CONVENIO: TStringField;
    qrConfigPDV: TUniQuery;
    qrConfigPDVCODIGO: TIntegerField;
    qrConfigPDVFORMA_CREDIARIO: TStringField;
    qrConfigPDVFORMA_CHEQUE: TStringField;
    qrConfigPDVFORMA_CARTAO: TStringField;
    qrConfigPDVFORMA_CONVENIO: TStringField;
    qrConfigPDVFORMA_DINHEIRO: TStringField;
    qrConfigPDVNF_PAULISTA: TIntegerField;
    qrConfigPDVTOTALIZADOR_NF_CREDIARIO: TStringField;
    qrConfigPDVNOME_TOTALIZADOR_NF_CREDIARIO: TStringField;
    qrConfigPDVTOTALIZADOR_NF_RECEBIMENTO: TStringField;
    qrConfigPDVNOME_TOTALIZADOR_NF_RECEBIMENTO: TStringField;
    qrConfigPDVTOTALIZADOR_SANGRIA: TStringField;
    qrConfigPDVTOTALIZADOR_SUPRIMENTO: TStringField;
    qrConfigPDVCOMPROVANTE_CREDIARIO: TIntegerField;
    qrConfigPDVCOMPROVANTE_CREDIARIO_PRODUTO: TIntegerField;
    qrConfigPDVUSA_TEF: TIntegerField;
    qrConfigPDVUSA_TEF_TECBAN: TIntegerField;
    qrConfigPDVUSA_TEF_HIPERCARD: TIntegerField;
    qrConfigPDVCAIXA_SITUACAO: TStringField;
    qrConfigPDVCAIXA_DATA_MOVTO: TDateField;
    qrConfigPDVCAIXA_COO_INICIAL: TStringField;
    qrConfigPDVMUDA_QTDE: TIntegerField;
    qrConfigPDVMUDA_UNITARIO: TIntegerField;
    qrConfigPDVMUDA_TOTAL: TIntegerField;
    qrConfigPDVCARGA_DATA: TStringField;
    qrConfigPDVCARGA_HORA: TStringField;
    qrConfigPDVSENHA_DESCONTO: TIntegerField;
    qrConfigPDVDESCONTO_MAXIMO: TFloatField;
    qrConfigPDVPARNOM: TStringField;
    qrConfigPDVPARVAL: TStringField;
    qrConfigPDVPARTIP: TStringField;
    qrConfigPDVALTERNARPRECOS: TIntegerField;
    qrConfigPDVPRECOCLIENTE: TIntegerField;
    qrConfigPDVCONSULTACREDIARIO: TIntegerField;
    qrConfigPDVAVISOSONORO: TIntegerField;
    qrConfigPDVDESCACRECPADRAO: TIntegerField;
    qrConfigPDVCORFUNDO: TIntegerField;
    qrConfigPDVCORLETRAS: TIntegerField;
    qrConfigPDVNAOVENDERSEMESTOQUE: TIntegerField;
    qrConfigPDVFECHAMENTO: TDateTimeField;
    qrConfigPDVIMPORTA_ORCAMENTO_GERENCIAL: TStringField;
    qrConfigPDVFORMA_CARTAO_CREDITO: TStringField;
    qrConfigPDVFORMA_CHEQUE_APRAZO: TStringField;
    qrConfigPDVIMPRIME_MENSAGEM_CUPOM: TStringField;
    qrConfigPDVPESAGEM_AUTOMATICA: TStringField;
    qrConfigPDVSOLICITA_SENHA_CANCELAMENTO: TStringField;
    qrConfigPDVDESCONTO_MAXIMO_VENDA: TFloatField;
    qrConfigPDVNUMCAIXA: TIntegerField;
    qrconfigV_EXIBE_BARRA_ATALHO: TStringField;
    qrconfigV_COR_FUNDO_BARRA_TITULO: TStringField;
    qrconfigV_COR_FONTE_BARRA_TITULO: TStringField;
    qrconfigV_COR_FUNDO_BARRA_RODAPE: TStringField;
    qrconfigV_COR_FONTE_BARRA_RODAPE: TStringField;
    qrconfigV_MEN_CAD_FUNDO: TStringField;
    qrconfigV_MEN_CAD_FONTE: TStringField;
    qrconfigV_MEN_EST_FUNDO: TStringField;
    qrconfigV_MEN_EST_FONTE: TStringField;
    qrconfigV_MEN_MOV_FUNDO: TStringField;
    qrconfigV_MEN_MOV_FONTE: TStringField;
    qrconfigV_MEN_FIN_FUNDO: TStringField;
    qrconfigV_MEN_FIN_FONTE: TStringField;
    qrconfigV_MEN_REL_FUNDO: TStringField;
    qrconfigV_MEN_REL_FONTE: TStringField;
    qrconfigV_IMAGEM_FUNDO: TStringField;
    qrconfigV_MEN_CON_FUNDO: TStringField;
    qrconfigV_MEN_CON_FONTE: TStringField;
    qrconfigV_ATA_BOT_COR_FUNDO: TStringField;
    qrconfigV_ATA_BOT_COR_FONTE: TStringField;
    qrconfigV_ATA_EXIBE_IMAGEM: TStringField;
    qrconfigV_ATA_EXIBE_RESUMO_CAIXA: TStringField;
    qrconfigV_ATA_CAIXA_COR_FUNDO: TStringField;
    qrconfigV_ATA_CAIXA_COR_FONTE: TStringField;
    qrconfigUTILIZA_LEITOR_CODIGO_BARRA: TStringField;
    qrconfigEMAIL_TLS: TStringField;
    qrconfigEMAIL_SSL: TStringField;
    qrconfigNAO_EXIBE_FECH_ATEND: TStringField;
    qrconfigNAO_INFORMA_CLI_ATEND: TStringField;
    qrConfigPDVMEN_CUP_SIMPLES: TStringField;
    qrConfigPDVMEN_CUP_NFCE: TStringField;
    qrConfigPDVMEN_CUP_SAT: TStringField;
    qrConfigPDVMEN_CUP_COZINHA: TStringField;
    Gaveta: TACBrGAV;
    qrItensCOMPLEMENTO: TStringField;
    qrImpDeliveryEXTRA: TBlobField;
    qrImpDeliveryCOMPLEMENTO: TStringField;
    qrImpDeliveryID_ITEM: TIntegerField;
    qrGrava_Cupom: TUniQuery;
    qrGrava_CupomCODIGO: TStringField;
    qrGrava_CupomNUMERO: TStringField;
    qrGrava_CupomCCF: TStringField;
    qrGrava_CupomECF: TStringField;
    qrGrava_CupomDATA: TDateField;
    qrGrava_CupomHORA: TTimeField;
    qrGrava_CupomQTDE_ITEM: TIntegerField;
    qrGrava_CupomVALOR_PRODUTO: TFloatField;
    qrGrava_CupomVALOR_DESCONTO: TFloatField;
    qrGrava_CupomVALOR_ACRESCIMO: TFloatField;
    qrGrava_CupomVALOR_TOTAL: TFloatField;
    qrGrava_CupomVALOR_PAGO: TFloatField;
    qrGrava_CupomVALOR_TROCO: TFloatField;
    qrGrava_CupomCOD_CLIENTE: TIntegerField;
    qrGrava_CupomCANCELADO: TIntegerField;
    qrGrava_CupomCPF_CONSUMIDOR: TStringField;
    qrGrava_CupomNOME_CONSUMIDOR: TStringField;
    qrGrava_CupomCOD_VENDEDOR: TIntegerField;
    qrGrava_CupomEX: TIntegerField;
    qrGrava_CupomCOD_CAIXA: TIntegerField;
    qrGrava_CupomECF_CAIXA: TStringField;
    qrGrava_CupomCONTINGENCIA: TStringField;
    qrGrava_CupomGERADO_NFCE: TStringField;
    qrGrava_CupomPEGO_PORQUEM: TStringField;
    qrGrava_CupomSAT: TIntegerField;
    qrGrava_CupomNAOFISC: TStringField;
    qrGrava_CupomCOUVERT: TFloatField;
    qrGrava_CupomTAXA_GARCON: TFloatField;
    qrGrava_CupomIMPRESSO: TStringField;
    qrGrava_CupomSENHA: TIntegerField;
    qrGrava_CupomNUMERO_NFCE: TIntegerField;
    qrGrava_CupomENDERECO_CONSUMIDOR: TStringField;
    qrGrava_CupomGEROU_SAT: TStringField;
    qrGrava_CupomNUMERO_SAT: TIntegerField;
    qrGravaItem: TUniQuery;
    qrGravaItemCODIGO: TStringField;
    qrGravaItemCOD_CUPOM: TStringField;
    qrGravaItemITEM: TIntegerField;
    qrGravaItemCOD_PRODUTO: TIntegerField;
    qrGravaItemUNIDADE: TStringField;
    qrGravaItemQTDE: TFloatField;
    qrGravaItemVALOR_UNITARIO: TFloatField;
    qrGravaItemVALOR_SUBTOTAL: TFloatField;
    qrGravaItemVALOR_DESCONTO: TFloatField;
    qrGravaItemVALOR_ACRESCIMO: TFloatField;
    qrGravaItemVALOR_TOTAL: TFloatField;
    qrGravaItemCST: TStringField;
    qrGravaItemALIQUOTA: TFloatField;
    qrGravaItemCOD_TOTALIZADOR: TStringField;
    qrGravaItemCANCELADO: TIntegerField;
    qrGravaItemEX: TIntegerField;
    qrGravaItemTAMANHO: TStringField;
    qrGravaItemCOR: TStringField;
    qrGravaItemCFOP: TStringField;
    qrGravaItemCSOSN: TStringField;
    qrGravaItemCLASSIFICACAO_FISCAL: TStringField;
    qrGravaItemBASE_SUB: TFloatField;
    qrGravaItemICMS_SUB: TFloatField;
    qrGravaItemISENTAS_ICMS: TFloatField;
    qrGravaItemOUTRAS_ICMS: TFloatField;
    qrGravaItemEXTRA: TBlobField;
    qrGravaItemCOMPLEMENTO: TStringField;
    qrGravaNaoFiscal: TUniQuery;
    qrGravaNaoFiscalCODIGO: TStringField;
    qrGravaNaoFiscalECF: TStringField;
    qrGravaNaoFiscalDATA: TDateField;
    qrGravaNaoFiscalINDICE: TStringField;
    qrGravaNaoFiscalDESCRICAO: TStringField;
    qrGravaNaoFiscalVALOR: TFloatField;
    qrGravaNaoFiscalCOO: TStringField;
    qrGravaNaoFiscalGNF: TStringField;
    qrGravaNaoFiscalGRG: TStringField;
    qrGravaNaoFiscalCDC: TStringField;
    qrGravaNaoFiscalDENOMINACAO: TStringField;
    qrGravaNaoFiscalHORA: TTimeField;
    qrGravaNaoFiscalCODVENDEDOR: TIntegerField;
    qrGravaForma: TUniQuery;
    qrGravaFormaCODIGO: TStringField;
    qrGravaFormaCOD_CUPOM: TStringField;
    qrGravaFormaFORMA: TStringField;
    qrGravaFormaVALOR: TFloatField;
    qrGravaFormaPRESTACAO: TIntegerField;
    qrGravaFormaTIPO: TStringField;
    qrGravaFormaEX: TIntegerField;
    qrGravaFormaTROCO: TFloatField;
    qrGravaFormaBANCO: TStringField;
    qrGravaFormaTIPOCARTAO: TStringField;
    ACBRDANFENFCe: TACBrNFeDANFCEFR;
    qrdevolucao_credito: TUniQuery;
    qrconfigNAO_UTILIZA_EXTRA: TStringField;
    qrconfigATIVA_CAD_PRODUTO_SIMPLES: TStringField;
    qrconfigATIVA_CAMPO_PEDIDO_CLIENTE: TStringField;
    qrconfigCEP_BUSCA: TStringField;
    qrconfigCEP_SERVIDOR: TIntegerField;
    qrconfigCEP_CHAVE: TStringField;
    qrconfigIND_CFOP_DEV_COMPRA_DENTRO: TStringField;
    qrconfigIND_CFOP_DEV_COMPRA_FORA: TStringField;
    qrconfigIND_CFOP_GARANTIA_DENTRO: TStringField;
    qrconfigIND_CFOP_GARANTIA_FORA: TStringField;
    qrconfigIND_CFOP_DEV_COMPRA_DENTRO_DESC: TStringField;
    qrconfigIND_CFOP_DEV_COMPRA_FORA_DESC: TStringField;
    qrconfigIND_CFOP_GARANTIA_DENTRO_DESC: TStringField;
    qrconfigIND_CFOP_GARANTIA_FORA_DESC: TStringField;
    qrconfigINICIAR_NOVA_VENDA: TStringField;
    qrconfigNOVO_RAMO_ATIVIDADE: TStringField;
    qrconfigVALIDAR_CAIXA_ATENDIMENTO: TStringField;
    qrconfigATUALIZA_ESTOQUE_PDV: TStringField;
    qrconfigBAIXA_ESTOQUE_VENDA_SIMPLES: TStringField;
    qrconfigDIAS_VAL_CUPOM_CREDITO: TIntegerField;
    qrconfigDIAS_VALIDADE_CUPOM_CREDITO: TIntegerField;
    qrdevolucao: TUniQuery;
    qrdevolucao_creditoID: TIntegerField;
    qrdevolucao_creditoVENCIMENTO: TDateField;
    qrdevolucao_creditoVALOR: TFloatField;
    qrdevolucao_creditoVALOR_UTILIZADO: TFloatField;
    qrdevolucao_creditoVALOR_SALDO: TFloatField;
    qrdevolucao_creditoCUPOM_UTILIZADO: TStringField;
    qrdevolucaoID: TIntegerField;
    qrdevolucaoID_CREDITO: TIntegerField;
    qrdevolucaoDATA: TDateField;
    qrdevolucaoHORA: TStringField;
    qrdevolucaoCODCLIENTE: TStringField;
    qrdevolucaoCUPOM_DEVOLUCAO: TStringField;
    qrdevolucaoITEM: TStringField;
    qrdevolucaoVALOR: TFloatField;
    qrSeqDevCredito: TUniQuery;
    qrSeqDevCreditoSEQ: TLargeintField;
    qrdevolucaoCUPOM_DEVOLVIDO: TStringField;
    qrGravaFormaID_CUPOM_CREDITO: TIntegerField;
    qrconfigVENDAS_SIMPLES_NAO_MOV_CAIXA: TStringField;
    qrconfigTAXA_GARCOM_TIPO_RECEB: TStringField;
    qrconfigCOUVER_TIPO_RECEB: TStringField;
    qrconfigCOUVER_PRODUTO: TStringField;
    qrconfigTAXA_GARCOM_PRODUTO: TStringField;
    qrconfigTIP_IMP_COMANDA: TIntegerField;
    qrconfigNAO_SOLICITA_EXCLUIR_ITEM_NFE: TStringField;
    qrconfigCAIXA_AUTOMATICO_GERENCIAL: TStringField;
    Qrsoma: TUniQuery;
    qrcaixa_operadorCODIGO: TStringField;
    qrcaixa_operadorCODFUNCIONARIO: TStringField;
    qrcaixa_operadorSENHA: TStringField;
    qrcaixa_operadorDATA: TDateField;
    qrcaixa_operadorSITUACAO: TIntegerField;
    qrcaixa_operadorSALDO: TFloatField;
    qrcaixa_operadorSITUACAO_ATUAL: TStringField;
    qrcaixa_operadorFECHAMENTO_CEGO: TStringField;
    qrcaixa_movCODIGO: TStringField;
    qrcaixa_movCODCAIXA: TStringField;
    qrcaixa_movCODOPERADOR: TStringField;
    qrcaixa_movDATA: TDateField;
    qrcaixa_movSAIDA: TFloatField;
    qrcaixa_movENTRADA: TFloatField;
    qrcaixa_movCODCONTA: TStringField;
    qrcaixa_movHISTORICO: TStringField;
    qrcaixa_movMOVIMENTO: TIntegerField;
    qrcaixa_movVALOR: TFloatField;
    qrcaixa_movCODNFSAIDA: TStringField;
    qrcaixa_movCODVENDA: TStringField;
    qrcaixa_movSERIAL_POS: TStringField;
    qrcaixa_movCARTAO_NOME_CONSUMIDOR: TStringField;
    qrcaixa_movCARTAO_BIN: TStringField;
    qrcaixa_movCARTAO_VALIDADE: TStringField;
    qrcaixa_movCARTAO_PARCELAS: TIntegerField;
    qrcaixa_movCARTAO_4DIGITOS: TIntegerField;
    qrcaixa_movCARTAO_N_AUTORIZACAO: TStringField;
    qrcaixa_movCARTAO_NSU: TStringField;
    qrcaixa_movCARTAO_INSTITUICAO: TStringField;
    qrcaixa_movCARTAO_ID_PAGAMENTO: TStringField;
    qrcaixa_movCHAVE: TStringField;
    qrcaixa_movNUMERO_CFE: TStringField;
    qrcaixa_movCARTAO_BANDEIRA: TStringField;
    qrcaixa_movCARTAO_RF: TStringField;
    qrcaixa_movCARTAO_MENSAGEM: TStringField;
    qrconfigPERMITIR_SENHAS_OBVIAS: TStringField;
    qrCupom_CartaoCODIGO_AUTORIZACAO: TStringField;
    qrCupom_CartaoCNPJ: TStringField;
    qrconfigB2B_XML_PRODUTO: TStringField;
    qrconfigTIPO_ETIQUETA_BALANCA: TStringField;
    qrDescontoGrupo: TUniQuery;
    qrDescontoGrupoCODIGO: TStringField;
    qrDescontoGrupoCOD_CUPOM: TStringField;
    qrDescontoGrupoITEM: TIntegerField;
    qrDescontoGrupoCOD_PRODUTO: TIntegerField;
    qrDescontoGrupoUNIDADE: TStringField;
    qrDescontoGrupoQTDE: TFloatField;
    qrDescontoGrupoVALOR_UNITARIO: TFloatField;
    qrDescontoGrupoVALOR_SUBTOTAL: TFloatField;
    qrDescontoGrupoVALOR_DESCONTO: TFloatField;
    qrDescontoGrupoVALOR_ACRESCIMO: TFloatField;
    qrDescontoGrupoVALOR_TOTAL: TFloatField;
    qrDescontoGrupoCST: TStringField;
    qrDescontoGrupoALIQUOTA: TFloatField;
    qrDescontoGrupoCOD_TOTALIZADOR: TStringField;
    qrDescontoGrupoCANCELADO: TIntegerField;
    qrDescontoGrupoTAMANHO: TStringField;
    qrDescontoGrupoCOR: TStringField;
    qrDescontoGrupoEXTRA: TBlobField;
    qrDescontoGrupoCOMPLEMENTO: TStringField;
    qrDescontoGrupoGRUPO: TStringField;
    qrTabelaPrecoGrupo: TUniQuery;
    qrTabelaPrecoGrupoID: TIntegerField;
    qrTabelaPrecoGrupoCOD_GRUPO: TStringField;
    qrTabelaPrecoGrupoQUANTIDADE_INICIAL: TFloatField;
    qrTabelaPrecoGrupoQUANTIDADE_FINAL: TFloatField;
    qrTabelaPrecoGrupoTIPO_DESCONTO: TStringField;
    qrTabelaPrecoGrupoINDICE: TFloatField;
    qrItensGRUPO: TStringField;
    qrTabelaPrecoGrupoTIPO_TABELA: TStringField;
    qrItensAPLICADO_DESCONTO_GRUPO: TStringField;
    qrItensQTD_APLICADO_DESCONTO: TFloatField;
    qrItensDESCONTO_CUPOM: TFloatField;
    qrNfceLocal: TUniQuery;
    qrNfceLocalNUMERO_NFCE: TLargeintField;
    qrfilialCODIGO: TStringField;
    qrfilialFILIAL: TStringField;
    qrfilialNOTAFISCAL: TIntegerField;
    qrfilialENDERECO: TStringField;
    qrfilialCIDADE: TStringField;
    qrfilialUF: TStringField;
    qrfilialCEP: TStringField;
    qrfilialTELEFONE: TStringField;
    qrfilialCNPJ: TStringField;
    qrfilialIE: TStringField;
    qrfilialSEQNF: TIntegerField;
    qrfilialFAX: TStringField;
    qrfilialOBS1: TStringField;
    qrfilialOBS2: TStringField;
    qrfilialCONTRIBUINTE_IPI: TStringField;
    qrfilialSUBSTITUTO_TRIBUTARIO: TStringField;
    qrfilialEMPRESA_ESTADUAL: TStringField;
    qrfilialOPTANTE_SIMPLES: TStringField;
    qrfilialOPTANTE_SUPER_SIMPLES: TStringField;
    qrfilialECF: TStringField;
    qrfilialTIPO: TIntegerField;
    qrfilialIPI: TFloatField;
    qrfilialISS: TFloatField;
    qrfilialNUMERO: TStringField;
    qrfilialRESPONSAVEL: TStringField;
    qrfilialCOMPLEMENTO: TStringField;
    qrfilialBAIRRO: TStringField;
    qrfilialFARMACIA_RESP_TECNICO: TStringField;
    qrfilialFARMACIA_CRF: TStringField;
    qrfilialFARMACIA_CPF: TStringField;
    qrfilialFARMCIA_DATA: TDateField;
    qrfilialFARMACIA_UF: TStringField;
    qrfilialFARMACIA_SENHA: TStringField;
    qrfilialFARMACIA_EMAIL: TStringField;
    qrfilialFARMACIA_LOGIN: TStringField;
    qrfilialFARMACIA_ENVIO: TStringField;
    qrfilialCONHECIMENTO: TIntegerField;
    qrfilialINDUSTRIA: TStringField;
    qrfilialFARMACIA_NUMEROLICENCA: TStringField;
    qrfilialCOD_MUNICIPIO_IBGE: TStringField;
    qrfilialIBGE: TStringField;
    qrfilialPIS: TFloatField;
    qrfilialCOFINS: TFloatField;
    qrfilialEMAIL: TStringField;
    qrfilialATIVIDADE: TStringField;
    qrfilialCONTADOR_COD_MUNICIPIO_IBGE: TStringField;
    qrfilialCONTADOR_NOME: TStringField;
    qrfilialCONTADOR_CPF: TStringField;
    qrfilialCONTADOR_CRC: TStringField;
    qrfilialCONTADOR_CNPJ: TStringField;
    qrfilialCONTADOR_CEP: TStringField;
    qrfilialCONTADOR_ENDERECO: TStringField;
    qrfilialCONTADOR_NUMERO: TStringField;
    qrfilialCONTADOR_COMPLEMENTO: TStringField;
    qrfilialCONTADOR_BAIRRO: TStringField;
    qrfilialCONTADOR_FONE: TStringField;
    qrfilialCONTADOR_FAX: TStringField;
    qrfilialCONTADOR_EMAIL: TStringField;
    qrfilialINSC_MUNICIPAL: TStringField;
    qrfilialDATA_ABERTURA: TDateField;
    qrfilialCNAE: TStringField;
    qrfilialCRT: TStringField;
    qrfilialCONTADOR_CIDADE: TStringField;
    qrfilialCONTADOR_COD_MUNICIPIO: TStringField;
    qrfilialCONTADOR_UF: TStringField;
    qrfilialPERMITE_CREDITO: TIntegerField;
    qrfilialFANTASIA: TStringField;
    qrfilialDFIXAS: TFloatField;
    qrfilialTIPOCALCULO: TIntegerField;
    qrfilialSERIE_CTE: TStringField;
    qrfilialSEQ_CTE: TStringField;
    qrfilialCOD_PAIS: TStringField;
    qrfilialPAIS: TStringField;
    qrfilialHOMEPAGE: TStringField;
    qrfilialAIDF: TStringField;
    qrfilialNINICIAL: TStringField;
    qrfilialNFINAL: TStringField;
    qrfilialN_NF_D: TStringField;
    qrfilialVALIDADE: TStringField;
    qrfilialRNTRC: TStringField;
    qrfilialSERIE_MDFE: TIntegerField;
    qrfilialNUMERO_MDFE: TIntegerField;
    qrfilialMAIL_SERVER: TStringField;
    qrfilialMAIL_PORTA: TStringField;
    qrfilialMAIL_USERNAME: TStringField;
    qrfilialMAIL_PASSWORD: TStringField;
    qrfilialMAIL_CABECALHO: TStringField;
    qrfilialMAIL_ASSUNTO: TStringField;
    qrfilialMAIL_MENSAGEM: TBlobField;
    qrfilialMAIL_SSL: TStringField;
    qrfilialMAIL_TLS: TStringField;
    qrfilialREPTEC_CNPJ: TStringField;
    qrfilialREPTEC_CONTATO: TStringField;
    qrfilialREPTEC_EMAIL: TStringField;
    qrfilialREPTEC_FONE: TStringField;
    qrfilialREPTEC_IDCSRT: TIntegerField;
    qrfilialREPTEC_CSRT: TStringField;
    qrfilialPERCENTUAL_CREDITO_ICMS: TFloatField;
    qrNFCEInsertNUMERO: TIntegerField;
    qrNFCEInsertSERIE: TIntegerField;
    qrNFCEInsertAMBIENTE: TStringField;
    qrNFCEInsertDATA: TDateField;
    qrNFCEInsertTOTAL: TFloatField;
    qrNFCEInsertCLIENTE: TStringField;
    qrNFCEInsertCHAVE: TStringField;
    qrNFCEInsertXML: TStringField;
    qrNFCEInsertSITUACAO: TIntegerField;
    qrNFCEInsertTROCO: TFloatField;
    qrNFCEInsertEX: TIntegerField;
    qrNFCEInsertHORA: TStringField;
    qrNFCEInsertCONTINGENCIA: TStringField;
    qrNFCEInsertENVIADOCONTINGENCIA: TStringField;
    qrNFCEInsertMOTIVOCONTIGENCIA: TStringField;
    qrNFCEInsertXMLENVIO: TBlobField;
    qrNFCEInsertXMLCACNELAMENTO: TBlobField;
    qrNFCEInsertALTERADO: TStringField;
    qrNFCEInsertCUPOM: TStringField;
    qrNFCEInsertXML_CANCELAMENTO: TStringField;
    qrNFCEInsertDATAHORACONTIMGENCIA: TDateTimeField;
    qrNFCEInsertSTATUS: TStringField;
    qrNFCEInsertERRO: TBlobField;
    cpGaveta: TACBrPosPrinter;
    qrEstoqueGRUPO: TStringField;
    qrEstoqueUSA_BALANCA: TIntegerField;
    qrEstoqueUSA_TB_PC: TStringField;
    qrEstoqueCODIGO_ANP: TStringField;
    qrEstoqueCOMBO: TStringField;
    qrEstoqueINFORMA_CODIGO_BARRA_XML: TStringField;
    qrEstoquePIZZA: TStringField;
    qrEstoquePERCGLNN: TFloatField;
    qrEstoquePERCGLGNI: TFloatField;
    qrEstoquePGLP: TFloatField;
    qrEstoqueVPART: TFloatField;
    qrEstoqueCEST: TStringField;
    qrEstoqueCODIGO_SERVICO: TStringField;
    qrEstoqueORIGEM: TIntegerField;
    qrEstoqueDESCRICAO_ANP: TStringField;
    qrEstoquePRECOCUSTO: TFloatField;
    qrEstoqueECF_ICMS: TStringField;
    qrEstoqueALIQIMPORTADA: TFloatField;
    qrEstoqueALIQESTADUAL: TFloatField;
    qrEstoqueALIQMUNICIPAL: TFloatField;
    qrprodutoUNIDADE: TStringField;
    qrprodutoATIVA: TStringField;
    qrprodutoCEST: TStringField;
    qrprodutoORIGEM: TIntegerField;
    qrprodutoIND_CFOP_NFCE: TStringField;
    qrprodutoFOTOBD: TBlobField;
    qrprodutoID_TIPO_SERVICO: TIntegerField;
    qrprodutoNAO_MOVIMENTA_ESTOQUE: TStringField;
    qrprodutoPESAGEM_AUOTMATICA: TStringField;
    qrprodutoINFORMA_CODIGO_BARRA_XML: TStringField;
    qrprodutoCODIGO_ANP: TStringField;
    qrprodutoCOMBO: TStringField;
    qrprodutoIND_CFOP_DEV_COMPRA_DENTRO: TStringField;
    qrprodutoIND_CFOP_DEV_COMPRA_FORA: TStringField;
    qrprodutoPIZZA: TStringField;
    qrprodutoPERCGLNN: TFloatField;
    qrprodutoPERCGLGNI: TFloatField;
    qrprodutoPGLP: TFloatField;
    qrprodutoVPART: TFloatField;
    qrprodutoECF_ICMS: TStringField;
    qrImpDeliveryUNIDADE: TStringField;
    qrconfigMULTILOJA: TStringField;
    qrconfigDATABASE: TStringField;
    qrconfigPORT: TStringField;
    qrconfigSERVER: TStringField;
    qrconfigUSERNAME: TStringField;
    qrconfigPASSWORD: TStringField;
    qrCaixaGeral: TUniQuery;
    qrCaixaGeralCODIGO: TStringField;
    qrCaixaGeralCODFUNCIONARIO: TStringField;
    qrCaixaGeralSENHA: TStringField;
    qrCaixaGeralDATA: TDateField;
    qrCaixaGeralSITUACAO: TIntegerField;
    qrCaixaGeralSALDO: TFloatField;
    qrCaixaGeralSITUACAO_ATUAL: TStringField;
    qrCaixaGeralFECHAMENTO_CEGO: TStringField;
    qrCaixaGeralIDONLINE: TIntegerField;
    qrCaixaGeralATB: TStringField;
    qrCaixaGeralCAIXA_GERAL: TStringField;
    qrconfigPDV_EXIBE_ESTOQUE_TELA_VENDA: TStringField;
    qrconfigLIBERA_CANC_PDV_GERENCIAL: TStringField;
    qrcaixa_movHORA: TStringField;
    qrcaixa_movIDONLINE: TIntegerField;
    qrcaixa_movATB: TStringField;
    qrcaixa_operadorIDONLINE: TIntegerField;
    qrcaixa_operadorATB: TStringField;
    qrcaixa_operadorCAIXA_GERAL: TStringField;
    qrConfTabelas: TUniQuery;
    qrConfTabelasTABELA: TStringField;
    qrConfTabelasDESCRICAO: TStringField;
    qrConfTabelasNIVEL_ACESSO: TStringField;
    qrConfTabelasNOME_TABELA_BANCO: TStringField;
    qrConfigPDVNAO_UTILIZA_EXTRA: TStringField;
    qrConfigPDVFECHAMENTO_AUTOMATICO_CAIXA_DIA: TStringField;
    qrConfigPDVEMPRESA_EXTRUTURAL: TStringField;
    qrConfigPDVEMPRESA_TIPO: TStringField;
    qrConfigPDVEMPRESA_DESCRICAO: TStringField;
    qrGravaNaoFiscalOBS: TStringField;
    qrAtb: TUniQuery;
    qrAtbID: TIntegerField;
    qrAtbPAI: TIntegerField;
    qrAtbFILHO: TIntegerField;
    qrAtbDESCRICAO: TStringField;
    qrAtbEXTRUTURAL: TStringField;
    qrAtbTIPO: TStringField;
    qrAtbEMPRESA_PADRAO: TStringField;
    UniQuery1: TUniQuery;
    StringField1: TStringField;
    StringField2: TStringField;
    IntegerField1: TIntegerField;
    IntegerField2: TIntegerField;
    StringField3: TStringField;
    FloatField1: TFloatField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    FloatField4: TFloatField;
    FloatField5: TFloatField;
    FloatField6: TFloatField;
    StringField4: TStringField;
    FloatField7: TFloatField;
    StringField5: TStringField;
    IntegerField3: TIntegerField;
    StringField6: TStringField;
    StringField7: TStringField;
    BlobField1: TBlobField;
    StringField8: TStringField;
    StringField9: TStringField;
    qrDescontoGrupoAPLICADO_DESCONTO_GRUPO: TStringField;
    qrDescontoGrupoQTD_APLICADO_DESCONTO: TFloatField;
    qrDescontoGrupoDESCONTO_CUPOM: TFloatField;
    qrTabelaPrecoGrupoCOD_SUBGRUPO: TStringField;
    qrTabelaPrecoGrupoAPLICAR_PARA: TStringField;
    qrTabelaPrecoGrupoCOD_FORNECEDOR_MARCA: TStringField;
    qrDescontoGrupoCODSUBGRUPO: TStringField;
    qrDescontoGrupoCODFORNECEDOR: TStringField;
    qrDescontoGrupoCODMARCA: TStringField;
    qrDescontoGrupoSINALIZADOR_DESCONTO: TIntegerField;
    qrTabelaPrecod: TUniQuery;
    qrTabelaPrecodUSA_TB_PC: TStringField;
    qrTabelaPrecodPRECO_VENDA: TFloatField;
    qrTabelaPrecodID: TIntegerField;
    qrTabelaPrecodCOD_PRODUTO: TIntegerField;
    qrTabelaPrecodTIPO_TABELA: TStringField;
    qrTabelaPrecodQUANTIDADE_INICIAL: TFloatField;
    qrTabelaPrecodQUANTIDADE_FINAL: TFloatField;
    qrTabelaPrecodVALOR_VENDA: TFloatField;
    qrTabelaPrecodVALIDADE_INICIO: TDateField;
    qrTabelaPrecodVALIDADE_FIM: TDateField;
    qrTabelaPreco: TUniQuery;
    qrTabelaPrecoUSA_TB_PC: TStringField;
    qrTabelaPrecoPRECO_VENDA: TFloatField;
    qrTabelaPrecoID: TIntegerField;
    qrTabelaPrecoCOD_PRODUTO: TIntegerField;
    qrTabelaPrecoTIPO_TABELA: TStringField;
    qrTabelaPrecoQUANTIDADE_INICIAL: TFloatField;
    qrTabelaPrecoQUANTIDADE_FINAL: TFloatField;
    qrTabelaPrecoVALOR_VENDA: TFloatField;
    qrTabelaPrecoVALIDADE_INICIO: TDateField;
    qrTabelaPrecoVALIDADE_FIM: TDateField;
    qrCupomCancelarNUMERO: TStringField;
    qrCupomCancelarDATA: TDateField;
    qrCupomCancelarHORA: TStringField;
    qrCupomCancelarUSU_CODIGO: TIntegerField;
    qrCupomCancelarUSU_NOME: TStringField;
    qrCupomCancelarEX: TStringField;
    qrCupomCancelarNUM_REABERTURA: TStringField;
    qrConfigPDVTROCO_MAXIMO: TFloatField;
    procedure BalancaLePeso(Peso: Double; Resposta: string);
    procedure DataModuleCreate(Sender: TObject);
    procedure ACBRNFCeStatusChange(Sender: TObject);
    procedure spCupom_Item_TempAfterExecute(Sender: TObject; Result: Boolean);
    procedure spCupom_Temp_deleteAfterExecute(Sender: TObject; Result: Boolean);
  private
		fConexao: tUniconnection;
    FPathSalvarXmlSatVenda: string;
	public
		mesa_comanda: string;
		FValorUintario: array of TValorUnitario;

		Property Conn:                      tUniConnection   Read fConexao                     Write fConexao;
    property PathSalvarXmlSatVenda:     string           read FPathSalvarXmlSatVenda       write FPathSalvarXmlSatVenda;

		procedure LerConfiguracao;
		function Iif(Condicao: Boolean; Verdadeiro, Falso: Variant): Variant;
		function GetSiteConsulta(UF: string; Ambiente: TpcnTipoAmbiente): string;
		function StrToHex(const S: string): string;
		function codifica(TABELA: string; Tipo: string = 'A'): string;
		function PrecoDeVenda(Produto: Integer; Quantidade: Double): Double;
		procedure FechaAbreCaixas;

		function BuscaUmDadoSql(sql: string;DefaultValue:Variant): variant;
		function BuscaUmDadoSqlAsInteger(sql: string): integer;
		function BuscaUmDadoSqlAsString(sql: string): String;
		function BuscaUmDadoSqlAsDateTime(sql: string): tDateTime;
		procedure ExecSql(pSql: string);

		Procedure InsereCampo(pTabela,pCampo:string); overload;
		Procedure InsereCampo(pTabela, pCampo, pComent: string); overload;
		Function  InsereTabela(pTabela:string;pCampos:array of string):boolean;
		Procedure InsereIndice(pTabela, pCampos: string; pNum:Integer );
		Procedure InsereGenerator(pNmGenerator:string); Overload;
		Procedure InsereGenerator(pNmGenerator:string; pNumSeq:integer); Overload;
		Procedure InsereChaveEstrangeira(pNomeFK,pTabela,pCampo,pReferencia:string);
		function  InsereView(pNomeView, vSql: string): boolean;
		function CampoExiste(pTabela, pCampo: string): boolean;
		procedure RemoveCampo(pTabela, pCampo: string);
		function TriggerExiste( pNm:string ): boolean ;
		Function GeneratorExiste( pNomeGenerator:string ):boolean;

	end;

function CaixaGeral: string;
function ME_GravaATB(Tabela: string): string;
function ME_FiltraATB(Tabela: string): string;
function DelphiAberto: Boolean;
function GetTextAfter(pTexto,pParte:string):string;
function GetTextBefore(pTexto,pParte:string):string;
Function DateToSQL( dta: tDate ): string;
function strzero( nm,tam:integer ): string ;
function qstr(pstr:string):string;
Function CopyToClipBoard(pTexto:String):boolean;



const
  sPAF_Versao_ER: string = '0103';

var
  frmModulo: TfrmModulo;
  ATB_ATIVA: tATB;
  Caixa_Geral: string;
  VersaoNFE: TvNFe;
  sConfiguracoes: string;
  emitente_uf, emitente_nome, emitente_fantasia, emitente_endereco, emitente_cidade, emitente_cep, emitente_cnpj, emitente_telefone, emitente_ie: string;
  resultado_pesquisa1, resultado_pesquisa2, resultado_pesquisa3, resultado_pesquisa4, resultado_pesquisa5, resultado_pesquisa6, resultado_pesquisa7, resultado_pesquisa8: string;
  resultado_valor1, resultado_valor2: real;
  parametro_pesquisa: string;
  StreamMemo: TMemoryStream;
  edtNumSerie: string;
  edtcfop: string;
  edtProxyHost: string;
  edtProxyPorta: string;
  edtProxyUser: string;
  edtProxySenha: string;
  EditLogoMarca: string;
  edtEmitCNPJ: string;
  edtEmitIE: string;
  edtEmitIM: string;
  edtEmitRazao: string;
  edtEmitFantasia: string;
  edtEmitFone: string;
  edtEmitCEP: string;
  edtEmailAssunto: string;
  edtSmtpPass: string;
  edtSmtpUser: string;
  edtSmtpPort: string;
  edtSmtpHost: string;
  edtEmitUF: string;
  edtEmitCidade: string;
  edtEmitCodCidade: string;
  edtEmitBairro: string;
  edtEmitLogradouro: string;
  edtEmitNumero: string;
  edtEmitComp: string;
  edtIdToken: string;
  edtSerie: string;
  edtTokenNumero: string;
  edtNatOperacao, edtPathLogs, PathXML: string;
  edDescEsta, edtSchema, MotivoContigencia: string;
  ckSalvar, ckVisualizar, ContigenciaPermanente: Boolean;
  edMargem, rgFormaEmissao, EdFormatoOff: Integer;
  edMargEsq, edMargDir, edMargSup, edMarginf: Double;

implementation

uses
  funcoes, principal, venda, tabela_preco;

{$R *.dfm}

function CaixaGeral: string;
begin
  with frmmodulo do
  begin
    qrCaixaGeral.Close;
    qrCaixaGeral.Open;
    if qrCaixaGeral.RecordCount = 0 then
    begin
      Application.MessageBox('Não foi localizado um "Caixa Geral", acesse o módulo Gerecial e Marque um Operador como Caixa Geral!', 'Atenção!', MB_ICONINFORMATION);
      Application.Terminate;
    end
    else
      Result := qrCaixaGeralCODIGO.AsString;
  end;
end;

function ME_GravaATB(Tabela: string): string;
begin
  with frmModulo do
  begin
    qrConfTabelas.Close;
    qrConfTabelas.Params.Items[0].Value := Tabela;
    qrConfTabelas.Open;
    if qrConfTabelas.RecordCount = 0 then
    begin
      Application.MessageBox(PWideChar('Não foi encontrado as configurações para a tabela "' + PWideChar(Tabela) + '"' + #13 + 'Sem estas configurações o sistema não armazena corretamente as infomrações no banco.' + #13 + 'Contacte o administardor do sistema e informe o ocorrido!'), 'Atenção!', MB_ICONWARNING);
      Result := '';
    end
    else
    begin
      if qrConfTabelasNIVEL_ACESSO.AsString = 'G' then
      begin
        if ATB_ATIVA.TIPO = 'G' then
        begin
          Result := ATB_ATIVA.ATB
        end
        else
        begin
          Result := Copy(ATB_ATIVA.ATB, 1, 3);
        end;
      end
      else if qrConfTabelasNIVEL_ACESSO.AsString = 'S' then
      begin
        if ATB_ATIVA.TIPO = 'S' then
        begin
          Result := ATB_ATIVA.ATB;
        end
        else
        begin
          Result := Copy(ATB_ATIVA.ATB, 1, 6);
        end;
      end
      else if qrConfTabelasNIVEL_ACESSO.AsString = 'E' then
      begin
        Result := ATB_ATIVA.ATB;
      end;
    end;
  end;
end;

function ME_FiltraATB(Tabela: string): string;
begin
  with frmModulo do
  begin
    qrConfTabelas.Close;
    qrConfTabelas.Params.Items[0].Value := Tabela;
    qrConfTabelas.Open;
    if qrConfTabelas.RecordCount = 0 then
    begin
      Application.MessageBox(PWideChar('Não foi encontrado as configurações para a tabela "' + PWideChar(Tabela) + '"' + #13 + 'Sem estas configurações o sistema não separa corretamente as infomrações no banco.' + #13 + 'Contacte o administardor do sistema e informe o ocorrido!'), 'Atenção!', MB_ICONWARNING);
      Result := '-1';
    end
    else
    begin
      if qrConfTabelasNIVEL_ACESSO.AsString = 'G' then
      begin
        if ATB_ATIVA.TIPO = 'G' then
        begin
          Result := ATB_ATIVA.ATB + '%';
        end
        else
        begin
          Result := Copy(ATB_ATIVA.ATB, 1, 3) + '%';
        end;
      end
      else if qrConfTabelasNIVEL_ACESSO.AsString = 'S' then
      begin
        if ATB_ATIVA.TIPO = 'S' then
        begin
          Result := ATB_ATIVA.ATB + '%';
        end
        else
        begin
          Result := Copy(ATB_ATIVA.ATB, 1, 6) + '%';
        end;
      end
      else if qrConfTabelasNIVEL_ACESSO.AsString = 'E' then
      begin
        Result := ATB_ATIVA.ATB + '%';
      end;
    end;
  end;
end;

procedure TfrmModulo.ACBRNFCeStatusChange(Sender: TObject);
begin
  Application.ProcessMessages;
end;

function TfrmModulo.codifica(TABELA: string; Tipo: string = 'A'): string;
begin
  qradic_mestre.open;
 // qradic_mestre.CommitUpdates;
  qradic_mestre.Refresh;
  if qradic_mestre.Locate('codigo', TABELA, [loCaseInsensitive]) then
  begin
    if Tipo = 'A' then
    begin
      if qradic_mestre.FIELDBYNAME('sequencia').AsInteger < 1 then
      begin
        result := '000001';
        qradic_mestre.Edit;
        qradic_mestre.FIELDBYNAME('sequencia').AsInteger := 1;
        qradic_mestre.Post;
      end
      else
      begin
        qradic_mestre.Edit;
        qradic_mestre.FIELDBYNAME('sequencia').AsInteger := qradic_mestre.FIELDBYNAME('sequencia').AsInteger + 1;
        qradic_mestre.Post;
        result := Zerar(IntToStr(frmModulo.qradic_mestre.FIELDBYNAME('sequencia').AsInteger), 6);
      end;
    end
    else
    begin
      qradic_mestre.Edit;
      qradic_mestre.FIELDBYNAME('sequencia').AsInteger := qradic_mestre.FIELDBYNAME('sequencia').AsInteger - 1;
      qradic_mestre.Post;
    end;
    Conexao_Servidor.Commit;
  end
  else
  begin
    showmessage('Não foi possível concluir com a operação!' + #13 + 'Erro: Código de sequência não encontrado na tabela de códigos.');
    Conexao_Servidor.Rollback;
  end;

end;

procedure TfrmModulo.BalancaLePeso(Peso: Double; Resposta: string);
begin
   // Buscar o Peso nas balancas PDV's
  if Peso > 0 then
  begin
     // Leitura OK
    sBal_Resposta := 'Peso Estável';
    rBal_peso := Peso;
  end
  else
  begin
     // Leitura retornou ERRo
    case Trunc(balanca.UltimoPesoLido) of
      0:
        sBal_Resposta := 'Coloque o produto sobre a Balança!';
      -1:
        sBal_Resposta := 'Peso Instavel!';
      -2:
        sBal_Resposta := 'Peso Negativo!';
      -10:
        sBal_Resposta := 'Sobrepeso!';
    end;
    rBal_peso := 0;
  end;
end;

procedure TfrmModulo.DataModuleCreate(Sender: TObject);
begin
  qrfilial.open;
  emitente_nome := qrfilialFILIAL.asstring;
  emitente_fantasia := qrfilialFANTASIA.asstring;
  emitente_endereco := qrfilialendereco.asstring + ' - ' + qrfilialbairro.asstring;
  emitente_cidade := qrfilialcidade.asstring + '/' + qrfilialuf.asstring;
  emitente_cep := qrfilialcep.asstring;
  emitente_telefone := qrfilialtelefone.asstring;
  emitente_cnpj := qrfilialcnpj.asstring;
  emitente_ie := qrfilialie.asstring;
  emitente_uf := qrfilialuf.asstring;
  qrfilial.Close;
end;


function TfrmModulo.GetSiteConsulta(UF: string; Ambiente: TpcnTipoAmbiente): string;
const
  site_uf_prod: array[0..6] of string = ('http://www.sefaznet.ac.gov.br/nfe/NFe.jsp?opc=3', 'http://sistemas.sefaz.am.gov.br/nfceweb/consultarNFCe.jsp', 'http://www.nfce.sefaz.ma.gov.br/portal/consultarNFCe.jsp', 'http://www.sefaz.mt.gov.br/nfe/portal/consultanfce', 'http://www.nfe.rn.gov.br/portal/consultarNFCe.jsp', 'https://www.sefaz.rs.gov.br/NFCE/NFCE-COM.aspx', 'http://www.nfe.se.gov.br/portal/consultarNFCe.jsp');
  site_uf_hom: array[0..6] of string = ('http://hml.sefaznet.ac.gov.br', 'http://homnfe.sefaz.am.gov.br/nfceweb/consultarNFCe.jsp', 'http://www.hom.nfce.sefaz.ma.gov.br/portal/consultarNFCe.jsp', 'http://homologacao.sefaz.mt.gov.br/nfce/consultanfce', 'http://www.hom.nfe.rn.gov.br/portal/consultarNFCe.jsp', 'https://www.sefaz.rs.gov.br/NFCE/NFCE-COM.aspx', 'http://www.hom.nfe.se.gov.br/portal/consultarNFCe.jsp');
begin
  case Ambiente of
    taProducao:
      begin
        if UF = 'AC' then
          result := site_uf_prod[0]
        else if UF = 'AM' then
          result := site_uf_prod[1]
        else if UF = 'MA' then
          result := site_uf_prod[2]
        else if UF = 'MT' then
          result := site_uf_prod[3]
        else if UF = 'RN' then
          result := site_uf_prod[4]
        else if UF = 'RS' then
          result := site_uf_prod[5]
        else if UF = 'SE' then
          result := site_uf_prod[6]
      end;
    taHomologacao:
      begin
        if UF = 'AC' then
          result := site_uf_hom[0]
        else if UF = 'AM' then
          result := site_uf_hom[1]
        else if UF = 'MA' then
          result := site_uf_hom[2]
        else if UF = 'MT' then
          result := site_uf_hom[3]
        else if UF = 'RN' then
          result := site_uf_hom[4]
        else if UF = 'RS' then
          result := site_uf_hom[5]
        else if UF = 'SE' then
          result := site_uf_hom[6]
      end;
  end;

end;

function TfrmModulo.Iif(Condicao: Boolean; Verdadeiro, Falso: Variant): Variant;
begin
  if Condicao then
    result := Verdadeiro
  else
    result := Falso;
end;

procedure TfrmModulo.LerConfiguracao;
var
  IniFile: string;
  Ini: TIniFile;
  Ok: Boolean;
begin

  IniFile := sConfiguracoes;

  Ini := TIniFile.Create(IniFile);
  try

    ACBRNFCe.Configuracoes.Certificados.ArquivoPFX := Ini.ReadString('Certificado', 'Caminho', '');
    ACBRNFCe.Configuracoes.Certificados.Senha := Ini.ReadString('Certificado', 'Senha', '');
    ACBRNFCe.Configuracoes.Certificados.NumeroSerie := Ini.ReadString('Certificado', 'NumSerie', '');

    ACBRNFCe.SSL.DescarregarCertificado;

    with ACBRNFCe.Configuracoes.Geral do
    begin
      SSLLib := TSSLLib(Ini.ReadInteger('Certificado', 'SSLLib', 0));
      SSLCryptLib := TSSLCryptLib(Ini.ReadInteger('Certificado', 'CryptLib', 0));
      SSLHttpLib := TSSLHttpLib(Ini.ReadInteger('Certificado', 'HttpLib', 0));
      SSLXmlSignLib := TSSLXmlSignLib(Ini.ReadInteger('Certificado', 'XmlSignLib', 0));
      AtualizarXMLCancelado := Ini.ReadBool('Geral', 'AtualizarXML', True);
      ExibirErroSchema := Ini.ReadBool('Geral', 'ExibirErroSchema', True);
      RetirarAcentos := Ini.ReadBool('Geral', 'RetirarAcentos', True);
      FormatoAlerta := Ini.ReadString('Geral', 'FormatoAlerta', 'TAG:%TAGNIVEL% ID:%ID%/%TAG%(%DESCRICAO%) - %MSG%.');
      ModeloDF := TpcnModeloDF(moNFCe);
      VersaoDF := TpcnVersaoDF(Ini.ReadInteger('Geral', 'VersaoDF', 0));
      IdCSC := Ini.ReadString('Geral', 'IdToken', '');
      CSC := Ini.ReadString('Geral', 'Token', '');
      Salvar := Ini.ReadBool('Geral', 'Salvar', True);
      VersaoQRCode := TpcnVersaoQrCode(Ini.ReadInteger('WebService', 'VersaoQRCode', 0));
    end;
    if ACBRNFCe.Configuracoes.Geral.VersaoDF = ve400 then
      VersaoNFE := v400
    else
      VersaoNFE := v310;

    with ACBRNFCe.Configuracoes.WebServices do
    begin
      UF := Ini.ReadString('WebService', 'UF', 'SP');
      Ambiente := StrToTpAmb(Ok, IntToStr(Ini.ReadInteger('WebService', 'Ambiente', 0) + 1));
      Visualizar := Ini.ReadBool('WebService', 'Visualizar', False);
      Salvar := Ini.ReadBool('WebService', 'SalvarSOAP', False);
      AjustaAguardaConsultaRet := Ini.ReadBool('WebService', 'AjustarAut', False);
      if NaoEstaVazio(Ini.ReadString('WebService', 'Aguardar', '0')) then
        AguardarConsultaRet := ifThen(StrToInt(Ini.ReadString('WebService', 'Aguardar', '0')) < 1000, StrToInt(Ini.ReadString('WebService', 'Aguardar', '0')) * 1000, StrToInt(Ini.ReadString('WebService', 'Aguardar', '0')));

      if NaoEstaVazio(Ini.ReadString('WebService', 'Tentativas', '5')) then
        Tentativas := StrToInt(Ini.ReadString('WebService', 'Tentativas', '5'));

      if NaoEstaVazio(Ini.ReadString('WebService', 'Intervalo', '0')) then
        IntervaloTentativas := ifThen(StrToInt(Ini.ReadString('WebService', 'Intervalo', '0')) < 1000, StrToInt(Ini.ReadString('WebService', 'Intervalo', '0')) * 1000, StrToInt(Ini.ReadString('WebService', 'Intervalo', '0')));

      TimeOut := Ini.ReadInteger('WebService', 'TimeOut', 5000);
      ProxyHost := Ini.ReadString('Proxy', 'Host', '');
      ProxyPort := Ini.ReadString('Proxy', 'Porta', '');
      ProxyUser := Ini.ReadString('Proxy', 'User', '');
      ProxyPass := Ini.ReadString('Proxy', 'Pass', '');
    end;
    MotivoContigencia := Ini.ReadString('WebService', 'MotivoContingencia', 'Sem conexão com a Internet.');
    ACBRNFCe.SSL.SSLType := TSSLType(Ini.ReadInteger('WebService', 'SSLType', 0));

    with ACBRNFCe.Configuracoes.Arquivos do
    begin
      Salvar := Ini.ReadBool('Arquivos', 'Salvar', false);
      SepararPorMes := Ini.ReadBool('Arquivos', 'PastaMensal', false);
      AdicionarLiteral := Ini.ReadBool('Arquivos', 'AddLiteral', false);
      EmissaoPathNFe := Ini.ReadBool('Arquivos', 'EmissaoPathNFe', false);
      SalvarEvento := Ini.ReadBool('Arquivos', 'SalvarPathEvento', false);
      SepararPorCNPJ := Ini.ReadBool('Arquivos', 'SepararPorCNPJ', false);
      SepararPorModelo := Ini.ReadBool('Arquivos', 'SepararPorModelo', false);
      PathSalvar := Ini.ReadString('Geral', 'PathSalvar', PathWithDelim(ExtractFilePath(Application.ExeName)) + 'Logs');
      PathSchemas := Ini.ReadString('Geral', 'PathSchemas', '');
      PathNFe := Ini.ReadString('Arquivos', 'PathNFe', '');
      PathInu := Ini.ReadString('Arquivos', 'PathInu', '');
      PathEvento := Ini.ReadString('Arquivos', 'PathEvento', '');
    end;

    EditLogoMarca := Ini.ReadString('PDV', 'CAMINHO_LOGO', '');
    if ACBRNFCe.DANFE <> nil then
      ACBRNFCe.DANFE.TipoDANFE := tiNFCe;
    if (EditLogoMarca <> '') and (FileExists(EditLogoMarca)) then
      ACBRNFCe.DANFE.Logo := EditLogoMarca;
    ACBRDANFENFCe.FastFile := 'C:\DANFeNFCe.fr3';

    edtcfop := Ini.ReadString('Geral', 'CFOP_PADRAO', '5101');

    EdFormatoOff := Ini.ReadInteger('WebService', 'FormatoOFF', 0);
    ContigenciaPermanente := Ini.ReadBool('WebService', 'ContingenciaPermanente', false);
    edtEmitCNPJ := Ini.ReadString('Emitente', 'CNPJ', '');
    edtEmitIE := Ini.ReadString('Emitente', 'IE', '');
    edtEmitIM := Ini.ReadString('Emitente', 'IM', '');
    edtEmitRazao := Ini.ReadString('Emitente', 'RazaoSocial', '');
    edtEmitFantasia := Ini.ReadString('Emitente', 'Fantasia', '');
    edtEmitFone := Ini.ReadString('Emitente', 'Fone', '');
    edtEmitCEP := Ini.ReadString('Emitente', 'CEP', '');
    edtEmitLogradouro := Ini.ReadString('Emitente', 'Logradouro', '');
    edtEmitNumero := Ini.ReadString('Emitente', 'Numero', '');
    edtEmitComp := Ini.ReadString('Emitente', 'Complemento', '');
    edtEmitBairro := Ini.ReadString('Emitente', 'Bairro', '');
    edtEmitCodCidade := Ini.ReadString('Emitente', 'CodCidade', '');
    edtEmitCidade := Ini.ReadString('Emitente', 'Cidade', '');
    edtEmitUF := Ini.ReadString('Emitente', 'UF', '');
    edtSerie := Ini.ReadString('Emitente', 'Serie', '');

    edDescEsta := Ini.ReadString('Impressao', 'Descricao via estabelecimento', 'Via do Consumidor');
    edMargem := Ini.ReadInteger('Impressao', 'Espessura Margem', 1);
    edMargEsq := Ini.ReadFloat('Impressao', 'Margem Esquerda', 0.6);
    edMargDir := Ini.ReadFloat('Impressao', 'Margem Direita', 0.51);
    edMargSup := Ini.ReadFloat('Impressao', 'Margem Superior', 0.8);
    edMarginf := Ini.ReadFloat('Impressao', 'Margem Inferior', 0.8);

    FPathSalvarXmlSatVenda     := INI.ReadString('SAT','PATH_SALVAR_VENDA','');

  finally
    Ini.Free;
  end;

end;

function TfrmModulo.PrecoDeVenda(Produto: Integer; Quantidade: Double): Double;
var
  AchouPreco: Boolean;
  lIdxItem: Integer;
begin
  AchouPreco := False;
  qrTabelaPreco.Filtered := False;
  qrTabelaPreco.Close;
  qrTabelaPreco.Params.Items[0].Value := Produto;
  qrTabelaPreco.Open;
  if not qrTabelaPreco.IsEmpty then
  begin
    Result := qrTabelaPrecoPRECO_VENDA.AsFloat;
    if qrTabelaPrecoUSA_TB_PC.AsString = 'SIM' then
    begin
      qrTabelaPreco.Filtered := True;
      qrTabelaPreco.Filter := 'TIPO_TABELA = ' + QuotedStr('P');
      if not qrTabelaPreco.IsEmpty then
      begin
        qrTabelaPreco.First;
        while not qrTabelaPreco.Eof do
        begin
          if (Date >= qrTabelaPrecoVALIDADE_INICIO.AsDateTime) and (Date <= qrTabelaPrecoVALIDADE_FIM.AsDateTime) then
          begin
            Result := qrTabelaPrecoVALOR_VENDA.AsFloat;
            AchouPreco := True;
            Break;
          end;
          qrTabelaPreco.Next;
        end;
      end;
      if not AchouPreco then
      begin
        qrTabelaPreco.Filter := 'TIPO_TABELA = ' + QuotedStr('A');
        if not qrTabelaPreco.IsEmpty then
        begin
          qrTabelaPreco.First;
          while not qrTabelaPreco.Eof do
          begin
            if (Quantidade >= qrTabelaPrecoQUANTIDADE_INICIAL.AsFloat) and (Quantidade <= qrTabelaPrecoQUANTIDADE_FINAL.AsFloat) then
            begin
              Result := qrTabelaPrecoVALOR_VENDA.AsFloat;
              AchouPreco := True;
              Break;
            end;
            qrTabelaPreco.Next;
          end;
        end;
      end;
      if not AchouPreco then
      begin
        qrTabelaPreco.Filter := 'TIPO_TABELA = ' + QuotedStr('V');
        if not qrTabelaPreco.IsEmpty then
        begin
          qrTabelaPreco.First;
          for lIdxItem := 0 to High(FValorUintario) do
            if FValorUintario[lIdxItem].Produto = Produto then
              Break;
          ///
          if lIdxItem > High(FValorUintario) then
          begin
            SetLength(FValorUintario, Succ(lIdxItem));
            FValorUintario[lIdxItem].IndiceAtivo := 0;
            FValorUintario[lIdxItem].Produto := Produto;
            FValorUintario[lIdxItem].ValorVenda := Result;
            SetLength(FValorUintario[lIdxItem].ValorDesconto, qrTabelaPreco.RecordCount);
          end;
          while not qrTabelaPreco.EOF do
          begin
            FValorUintario[lIdxItem].ValorDesconto[Pred(qrTabelaPreco.RecNo)] := qrTabelaPrecoVALOR_VENDA.AsFloat;
            qrTabelaPreco.Next;
          end;
        end;
      end;
    end;
  end
  else
  begin
    Application.MessageBox('Produto Inválido!', 'Atenção!', MB_ICONINFORMATION);
    Result := 0;
  end;
end;

procedure TfrmModulo.RemoveCampo(pTabela, pCampo: string);
begin
	if CampoExiste( pTabela, pCampo)  then
		ExecSql('ALTER TABLE '+pTabela.TOUPPER+' drop '+pCampo.TOUPPER);
end;

procedure TfrmModulo.spCupom_Item_TempAfterExecute(Sender: TObject; Result: Boolean);
begin
  qrItens.DisableControls;
  qrItens.Close;
  qrItens.Open;
  qrItens.Last;
  qrItens.EnableControls;
end;

procedure TfrmModulo.spCupom_Temp_deleteAfterExecute(Sender: TObject; Result: Boolean);
begin
  qrItens.DisableControls;
  qrItens.Close;
  qrItens.Open;
  qrItens.Last;
  qrItens.EnableControls;
end;

function TfrmModulo.StrToHex(const S: string): string;
const
  HexDigits: array[0..15] of Char = '0123456789ABCDEF';
var
  i: Integer;
  P1: PChar;
  P2: PChar;
  B: Byte;
begin

  SetLength(Result, length(S) * 2);
  P1 := @S[1];
  P2 := @Result[1];

  for i := 1 to length(S) do
  begin
    B := Byte(P1^);
    P2^ := HexDigits[B shr 4];
    Inc(P2);
    P2^ := HexDigits[B and $F];
    Inc(P1);
    Inc(P2);
  end;

end;

function TfrmModulo.TriggerExiste(pNm: string): boolean;
begin
	result := ( BuscaUmDadoSqlAsInteger( ' select count(*) from rdb$triggers a where a.rdb$trigger_name='+qStr( uppercase(trim(pNm)) ) ) )<>0 ;
end;

procedure Tfrmmodulo.FechaAbreCaixas;
var
  Saldo_Anterior: Double;
begin
  if qrcaixa_operador.State = dsBrowse then
    qrcaixa_operador.Locate('codigo', Caixa_Geral, [loCaseInsensitive]);

  Qrsoma.Close;
  Qrsoma.SQL.clear;
  Qrsoma.SQL.add('select sum(entrada) -  sum(saida) saldo ' + 'from c000044 where data = :data_caixa');
  Qrsoma.Params.ParamByName('data_caixa').asdatetime := qrcaixa_operador.FieldByName('data').asdatetime;
  Qrsoma.open;

  Saldo_Anterior := Qrsoma.FieldByName('saldo').AsFloat;

  qrcaixa_operador.First;
  while not qrcaixa_operador.Eof do
  begin
    if (qrcaixa_operador.FieldByName('codigo').asstring = Caixa_Geral) or (qrcaixa_operador.FieldByName('codigo').asstring = Zerar(IntToStr(iNumCaixa),6)) then begin
      qrcaixa_operador.Edit;
      qrcaixa_operador.FieldByName('data').asdatetime := date;
      qrcaixa_operador.FieldByName('situacao').asinteger := 1;
      qrcaixa_operador.post;
    end;
    qrcaixa_operador.Next;
  end;

  // LANCAMENTO DO SALDO ANTERIOR
  qrcaixa_operador.First;
  qrcaixa_operador.Locate('codigo', '000099', [loCaseInsensitive]);

  qrconfig.open;
  qrcaixa_mov.OPEN;
  qrcaixa_mov.insert;
  qrcaixa_mov.FieldByName('codigo').asstring := frmprincipal.codifica('000044', 6);
  qrcaixa_mov.fieldbyname('codcaixa').asstring := qrcaixa_operador.FIELDBYNAME('CODIGO').ASSTRING;
  qrcaixa_mov.fieldbyname('codoperador').asstring := Caixa_Geral;
  qrcaixa_mov.fieldbyname('data').AsDateTime := Date;
  qrcaixa_mov.fieldbyname('saida').asfloat := 0;
  qrcaixa_mov.fieldbyname('entrada').asfloat := Saldo_Anterior;
  qrcaixa_mov.FieldByName('valor').asfloat := Saldo_Anterior;
  qrcaixa_mov.fieldbyname('codconta').asstring := qrconfig.fieldbyname('PLANO_OUTRAS_ENTRADAS').asstring;
  qrcaixa_mov.fieldbyname('movimento').asINTEGER := 0;
  qrcaixa_mov.fieldbyname('historico').asstring := 'S A L D O   A N T E R I O R';
  qrcaixa_mov.fieldbyname('atb').asstring := ME_GravaATB('TB_MOVIMENTO_CAIXA');
  qrcaixa_mov.post;
  frmmodulo.Conexao.Commit;
end;

procedure Tfrmmodulo.ExecSql(pSql: string);
var qry: tUniQuery;
begin
	if not assigned(fConexao) then
		raise Exception.Create('Conn: not asssigned');
	qry:=tUniQuery.Create(self);
	qry.paramcheck:=false;
	try
	 try
		 qry.Connection:=Conexao;
		 qry.SQL.text:=pSql;
		 if delphiaberto then
				CopyToClipBoard( pSql );
		 qry.ExecSQL;
	 Except
		 on E:Exception do
		 begin
				ShowMessage( E.message+#13+pSql+#13+' *** Copiado para area de transferência ***');
				raise ;
		 end;
	 end;
	finally
	 qry.close;
	 FreeAndNil (qry);
	end;
end;

Function CopyToClipBoard(pTexto:String):boolean;
var success: boolean;
		i: integer;
begin
		 Success := False;
     i := 0;
     while not Success do
     try
        try
           Clipboard.Open;
           Clipboard.AsText := pTexto;
           Success := True;
        finally
               Clipboard.close;
        end;
     except
           on Exception do
           begin
                Inc(i);
                if i < 3 then
                   Sleep(i * 100)
                else
                begin
                     ShowMessage( 'Falha ao tentar acessar a área de transferência.' );
                     Exit;
                end;
           end;
		 end;
		 result:=success;
end;

function Tfrmmodulo.BuscaUmDadoSql(sql: string; DefaultValue: Variant): variant;
var qry: tUniQuery;
begin
 if not assigned(fConexao) then
		raise Exception.Create('Conn: not asssigned');
 qry:=tUniQuery.Create(self);
 try
	 qry.Connection:=Conn;
	 qry.SQL.text:=sql;
	 if DelphiAberto then
			CopyToClipBoard( sql );
	 qry.Open;
	 if qry.Fields[0].isnull then
			Result := DefaultValue
	 Else
			Result := qry.Fields[0].AsVariant;
 finally
	 qry.close;
	 FreeAndNil (qry);
 end;
end;

function Tfrmmodulo.BuscaUmDadoSqlAsInteger(sql: string): integer;
begin
		 Result := BuscaUmDadoSql(sql,0);
end;

function Tfrmmodulo.BuscaUmDadoSqlAsString(sql: string): String;
begin
		 Result := BuscaUmDadoSql(sql,'');
end;

function TfrmModulo.CampoExiste(pTabela, pCampo: string): boolean;
begin
	Result := BuscaUmDadoSqlasInteger( 'select count(*) from rdb$relation_fields '+
																		'where RDB$RELATION_FIELDS.rdb$relation_name = '+qstr(pTabela.TOUPPER)+' AND '+
																		'RDB$RELATION_FIELDS.RDB$FIELD_NAME = '+qStr(pCampo.TOUPPER) )<>0 ;
end;

function Tfrmmodulo.BuscaUmDadoSqlAsDateTime(sql: string): tDateTime;
begin
		 Result := BuscaUmDadoSql(sql,now);
end;

procedure Tfrmmodulo.InsereCampo(pTabela, pCampo: string);
begin
		 InsereCampo(pTabela, pCampo, '');
end;

function GetTextBefore(pTexto,pParte:string):string;
var p:integer;
begin
	p := pos(pParte,pTexto);
	if p>0 then
		 result:=Copy(pTexto,1,p-length(pParte))
	Else
		 result:='';
end;

function GetTextAfter(pTexto,pParte:string):string;
var p:integer;
begin
	p := pos(pParte,pTexto);
	if p>0 then
		 result:=Copy(pTexto,p+length(pParte),length(pTexto))
	Else
		 result:='';
end;

function qstr(pstr:string):string;
begin
	result := quotedstr(pstr);
end;

function strzero( nm,tam:integer ): string ;
var y,n: integer;
begin
     result := inttostr(nm)  ;
     y := length( result );
     if y < tam then
        for n := 1 to (tam-y) do
            result := '0'+result;
end;

Function DateToSQL( dta: tDate ): string;
var y,m,d: word;
begin
		 decodedate( dta, y,m,d);
		 result := quotedstr( strzero(y,4)+'-'+strzero(m,2)+'-'+strzero(d,2) );
end;

procedure Tfrmmodulo.InsereCampo(pTabela, pCampo, pComent: string);
var vnmCampo: string;
begin
	vnmCampo:=GetTextBefore(pCampo,' ');
	if length(vnmCampo)>30 then
		Raise Exception.Create('O nome do campo '+vnmCampo.TOUPPER+'(Tabela:'+pTabela.TOUPPER+') excede o tamanho máximo possível (30Caract).');
	if not CampoExiste( pTabela, vnmCampo)  then
	begin
		ExecSql('ALTER TABLE '+pTabela.TOUPPER+' ADD '+pCampo.TOUPPER);
		if pComent<>'' then
			ExecSql( 'COMMENT ON COLUMN '+pTabela.toUpper+'.'+vnmCampo.toUpper+' IS '+qSTR(pComent) );
	end;
end;

procedure TfrmModulo.InsereChaveEstrangeira(pNomeFK, pTabela, pCampo, pReferencia: string);
begin
	pNomeFK:=Trim(pNomeFK);
	if length(pNomeFK)>30 then
		Raise Exception.create('O nome de Chave Estrangeira '+pNomeFK+' excede o tamanho máximo possível (30Caract).');
	if buscaumdadosqlasinteger('select count(*) FROM RDB$RELATION_CONSTRAINTS where RDB$CONSTRAINT_NAME='+qStr(pNomeFK))=0 then
		ExecSql('alter table '+pTabela+' add constraint '+pNomeFK+' foreign key ('+pCampo+') references '+pReferencia);
end;

procedure TfrmModulo.InsereGenerator(pNmGenerator: string);
begin
	InsereGenerator(pNmGenerator,0);
end;

procedure TfrmModulo.InsereGenerator(pNmGenerator: string; pNumSeq: integer);
begin
	pNmGenerator:=Trim(pNmGenerator.toupper);
	if length(pNmGenerator)>30 then
		Raise Exception.create('O nome do generator '+pNmGenerator+' excede o tamanho máximo possível (30Caract).');
	if not GeneratorExiste( pNmGenerator ) then
	begin
			ExecSql('CREATE GENERATOR '+pNmGenerator);
			ExecSql('SET GENERATOR '+pNmGenerator+' TO '+pNumSeq.toString);
	end;
end;

procedure TfrmModulo.InsereIndice(pTabela, pCampos: string; pNum: Integer);
var vNmIndx: string;
begin
	vNmIndx:=Trim(pTabela.toupper)+'_INDX_'+StrZero(pNum,3);
	if length(vNmIndx)>30 then
		Raise Exception.create('O nome de indice '+vNmIndx+' excede o tamanho máximo possível (30Caract).');
	if buscaumdadosqlasinteger('select count(*) FROM RDB$INDICES where RDB$INDEX_NAME='+qStr(vNmIndx.toupper))=0 then
		ExecSql('CREATE INDEX '+vNmIndx.toupper+' ON '+pTabela+'('+pCampos.toupper+')');
end;

function Tfrmmodulo.InsereTabela(pTabela: string; pCampos: array of string): boolean;
var
	vSql,vCampo:string;
	I: integer;
begin
	if length(pTabela)>26 then
		Raise Exception.Create('O nome da tabela '+pTabela.toupper+' excede o tamanho máximo possível (26Caract).');
	result := BuscaUmDadoSqlasInteger( 'select count(*) from RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='+qstr(pTabela.toupper) )=0;
	if result then
	begin
		vSql:= 'CREATE TABLE '+pTabela.toupper+'('+
					 'ID INTEGER NOT NULL,'+
					 'DTCADASTRO DATE,'+
					 'USERCADASTROU VARCHAR(20) COLLATE WIN_PTBR,'+
					 'DTEDICAO TIMESTAMP,'+
					 'USEREDITOU VARCHAR(20) COLLATE WIN_PTBR';
		 for I := Low(pCampos) to High(pCampos) do
		 begin
				 vCampo:=Trim(GetTextBefore(pCampos[I],' '));
				 if length(vCampo)>30 then
						Raise Exception.create('O nome do campo '+vCampo.toupper+'(Tabela:'+pTabela.toupper+') excede o tamanho máximo possível (30Caract).');
				 vSql:=vSql+','+pCampos[I].toupper;
		 end;
		 vSql:=vSql+')';
		 ExecSql(vSql);
		 InsereGenerator('ID_'+pTabela.toupper);
		 ExecSql('ALTER TABLE '+pTabela.toupper+' ADD CONSTRAINT '+pTabela.toupper+'PK PRIMARY KEY (ID)');
		 ExecSql('CREATE INDEX '+pTabela.toupper+'01 ON '+pTabela.toupper+'(ID)');
	end;
end;

function TfrmModulo.InsereView(pNomeView, vSql: string): boolean;
begin
	if length(pNomeView)>26 then
		Raise Exception.create('O nome da view '+pNomeView.toupper+' excede o tamanho máximo possível (26Caract).');
	result := BuscaUmDadoSqlasInteger( 'select count(*) from RDB$VIEW_RELATIONS WHERE RDB$VIEW_NAME='+qstr(pNomeView.toupper) )=0;
	if result then
		ExecSql(vSql);
end;

function JanelaExiste(Classe, Janela: string): Boolean;
var
	aClasse, aJanela: array[0..79] of char;
begin
	if Classe = '' then
		aClasse[0] := #0
	else
		StrPCopy(aClasse, Classe);
	if Janela = '' then
		aJanela[0] := #0
	else
		StrPCopy(aJanela, Janela);
	Result := FindWindow(aClasse, aJAnela) <> 0;
end;

function DelphiAberto: Boolean;
begin
	Result := JanelaExiste('TPropertyInspector', 'Object Inspector');
end;

function TfrmModulo.GeneratorExiste(pNomeGenerator: string): boolean;
begin
	Result := BuscaUmDadoSqlasInteger( 'select Count(*) from rdb$generatorS Where rdb$generator_name='+qStr(pNomeGenerator) )<>0;
end;

end.

