unit funcoes;

// Unit contendo biblioteca de funcoes diversas
interface

uses
  IniFiles, SysUtils, Dialogs, Forms, Windows, classes, db, libeay32,
  ACBrGAV, senha_supervisor, ACBrPosPrinter;

type
  // tipos de alinhamentos
  TAlinhamento = (taEsquerda, taDireita, taCentralizado);

function retorna_codigo_aliquota(aliquota: real): string;
function formata_data(data: string): TDateTime;
function formata_hora(hora: string): string;
procedure AcertaPadraoData;
function troca_data(sdata: string): string;
function somenteNumero(sNum: string): string;
function codifica_cupom(): string;
function codifica_item(item: Integer): string;
function codifica_forma(idForma: Integer): string;
function codifica_crediario(idPrestacao: Integer): string;
function Executa_Login(usuario, senha: string; funcao: string; bValidaDesconto: Boolean = False;
 bAbrirCaixa: Boolean = False; bFechaCaixa: Boolean = False; bConfiguracoes: Boolean = False;
 bCancelarVenda: Boolean = False; bCancelarItem: Boolean = False): boolean;
function Cript(Action, Src: string): string;
function Zerar(texto: string; qtde: Integer): string;
function texto_justifica(texto: string; qtde_caracteres: Integer; caracter: string; Alinhamento: TAlinhamento): string;
function tiraacento(Str: string): string;
function Tiraacento_display(Str: string): string;
function UltimoDiaMes(Mdt: TDateTime): TDateTime;
function LerParametro(psNome: string; psTipo: string; pvDefault: Variant): Variant;
procedure GravarParametro(psNome: string; psTipo: string; pvValor: Variant);
function RetornarParametroString(psTipo: string; pvValor: Variant): string;
procedure Gaveta_Configura;
Procedure Gaveta_Abre;
function LerIni(sArquivo, Secao, Linha: string; Def: string): string;
procedure GravarIni(sArquivo, Secao, Linha: string; Valor: string);


implementation

uses
  modulo, principal, FechamentoCego, caixa_fechamento;

function LerIni(sArquivo, Secao, Linha, Def: string): string;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(sArquivo);
  Result := Ini.ReadString(Secao, Linha, Def);
  Ini.Free;
end;

procedure GravarIni(sArquivo, Secao, Linha: string; Valor: string);
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(sArquivo);
  Ini.WriteString(Secao, Linha, Valor);
  Ini.Free;
end;

function retorna_codigo_aliquota(aliquota: real): string;
begin
  frmmodulo.tbAliquota.Open;
  if frmmodulo.tbAliquota.Locate('Aliquota', aliquota, [loCaseInsensitive]) then
    Result := frmmodulo.tbAliquota.fieldbyname('codigo').asstring
  else
    Result := 'N1'; // retorna nao tributado por default

end;

function formata_data(data: string): TDateTime;
begin
  if length(data) = 6 then begin
    data := Copy(data, 1, 2) + '/' + Copy(data, 3, 2) + '/20' + Copy(data, 5, 2);

  end
  else begin
    if length(data) = 8 then begin
      data := Copy(data, 1, 2) + '/' + Copy(data, 3, 2) + '/' + Copy(data, 5, 4);
    end;
  end;
  Result := strtodate(data);
end;

function formata_hora(hora: string): string;
begin

  if length(hora) = 4 then begin
    hora := Copy(hora, 1, 2) + ':' + Copy(hora, 3, 2) + ':00';

  end
  else begin
    if length(hora) = 6 then begin
      hora := Copy(hora, 1, 2) + ':' + Copy(hora, 3, 2) + ':' + Copy(hora, 5, 4);
    end;
  end;
  Result := hora;
end;

// -------------------------------------------------------------------------- //
procedure AcertaPadraoData;
const
  arrShortDayNames: array[1..7] of string[3] = ('Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab');
  arrLongDayNames: array[1..7] of string[15] = ('Domingo', 'Segunda', 'TerÁa', 'Quarta', 'Quinta', 'Sexta', 'S·bado');
  arrShortMonthNames: array[1..12] of string[3] = ('Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez');
  arrLongMonthNames: array[1..12] of string[15] = ('Janeiro', 'Fevereiro', 'MarÁo', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro');
var
  ii: Integer;
begin
  //ShortDateFormat := 'dd/mm/yyyy';
  //DecimalSeparator := ',';
  //ThousandSeparator := '.';
  for ii := 1 to 7 do begin
    //ShortDayNames[ii] := arrShortDayNames[ii];
    //LongDayNames[ii] := arrLongDayNames[ii];
  end;
  for ii := 1 to 12 do begin
    //ShortMonthNames[ii] := arrShortMonthNames[ii];
    //LongMonthNames[ii] := arrLongMonthNames[ii];
  end;
end;

function troca_data(sdata: string): string;
begin
  sdata := somenteNumero(sdata);
  if length(sdata) = 6 then
    Result := Copy(sdata, 5, 2) + Copy(sdata, 3, 2) + Copy(sdata, 1, 2)
  else
    Result := Copy(sdata, 5, 4) + Copy(sdata, 3, 2) + Copy(sdata, 1, 2);
end;

function somenteNumero(sNum: string): string;
var
  s1, s2: string;
  i: Integer;
begin
  s1 := sNum;
  s2 := '';
  for i := 1 to length(s1) do
    if s1[i] in ['0'..'9'] then
      s2 := s2 + s1[i];
  Result := s2;
end;

function codifica_cupom(): string;
begin
  Result := Zerar(sCaixa, 3) + // numero do caixao do ecr
    Zerar(sNumero_Cupom, 12); // numero do cupom fiscal
end;

function codifica_item(item: Integer): string;
begin
  Result := Zerar(sCaixa, 3) + // numero do caixao do ecr
    Zerar(sNumero_Cupom, 12) + // numero do cupom fiscal
    Zerar(inttostr(item), 3); // numero do item
end;

function codifica_forma(idForma: Integer): string;
begin
  Result := Zerar(sCaixa, 3) + // numero do caixao do ecr
    Zerar(sNumero_Cupom, 12) + // numero do cupom fiscal
    Zerar(inttostr(idForma), 3); // numero do item
end;

function codifica_crediario(idPrestacao: Integer): string;
begin
  Result := Zerar(sCaixa, 3) + // numero do caixao do ecr
    Zerar(sNumero_Cupom, 12) + // numero do cupom fiscal
    Zerar(inttostr(idPrestacao), 3); // numero do item
end;

function Executa_Login(usuario, senha: string; funcao: string; bValidaDesconto: Boolean = False;
   bAbrirCaixa: Boolean = False; bFechaCaixa: Boolean = False; bConfiguracoes: Boolean = False;
   bCancelarVenda: Boolean = False; bCancelarItem: Boolean = False): boolean;
var
  bValidouSenha:Boolean;
begin
  Result := False;
  // usuario de gerencial para cancelamento de funcoes
  if usuario = 'DELPHOS' then begin
    frmmodulo.query.close;
    frmmodulo.query.sql.clear;
    frmmodulo.query.sql.add('select * from supervisor');
    frmmodulo.query.Open;
    bValidouSenha := False;
    while not frmmodulo.query.Eof do begin
      if copy(Cript('D', frmmodulo.query.fieldbyname('senha').asstring) ,0,7) = copy(senha,0,7) then begin
        bValidouSenha := True;
        if bValidaDesconto then begin
          if frmmodulo.query.FieldByName('aprova_desconto').AsString = 'S' then
            Result := True;
        end else if bAbrirCaixa then begin
          if frmmodulo.query.FieldByName('abrir_caixa').AsString = 'S' then
            Result := True;
        end else if bFechaCaixa then begin
          if frmmodulo.query.FieldByName('fecha_caixa').AsString = 'S' then
            Result := True;
        end else if bConfiguracoes then begin
          if frmmodulo.query.FieldByName('configuracoes').AsString = 'S' then
            Result := True;
        end else if bCancelarVenda then begin
          if frmmodulo.query.FieldByName('cancelar_venda').AsString = 'S' then
            Result := True;
        end else if bCancelarItem then begin
          if frmmodulo.query.FieldByName('cancelar_item').AsString = 'S' then
            Result := True;
        end else
          Result := True;
      end;
      frmmodulo.query.Next;
    end;
    if not Result then begin
      if not bValidouSenha then begin
//          Application.MessageBox('Acesso Negado!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bValidaDesconto then begin
          Application.MessageBox('Sem permiÁ„o de aprovaÁ„o de desconto!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bAbrirCaixa then begin
          Application.MessageBox('Sem permiÁ„o para Abrir o Caixa!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bFechaCaixa then begin
          Application.MessageBox('Sem permiÁ„o para Fechar o Caixa!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bConfiguracoes then begin
          Application.MessageBox('Sem permiÁ„o para Acessar as ConfiguraÁıes!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bCancelarVenda then begin
         Application.MessageBox('Sem permiÁ„o para Cancelar Venda!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end else if bCancelarItem then begin
          Application.MessageBox('Sem permiÁ„o para Cancelar Item da Venda!', 'AtenÁ„o!', MB_ICONINFORMATION);
      end;
    end;
  end
  else begin
    // usuario normal

    frmmodulo.query.close;
    frmmodulo.query.sql.clear;
    frmmodulo.query.sql.add('select * from adm');
    frmmodulo.query.sql.add('where codigo like ' + QuotedStr(usuario));
    frmmodulo.query.sql.add('and info2 = ''' + senha + '''');
    frmmodulo.query.Open;
    if frmmodulo.query.recordcount > 0 then begin
      if funcao = '001' then // entrada no sistema e troca de operador
      begin
        icodigo_Usuario := frmPrincipal.IsInteger(frmmodulo.query.fieldbyname('codigo').asstring);
        sCaixa := Zerar(IntToStr(icodigo_Usuario), 3);
        //sCaixa := Zerar(IntToStr(frmmodulo.qrConfigPDVNUMCAIXA.AsInteger),3);
        sNome_Operador := frmmodulo.query.fieldbyname('info1').asstring;
        bFechamentoCego := frmmodulo.query.fieldbyname('fechamento_cego').asstring = 'S';
      end;
      Result := True;
      exit;
    end;

    frmmodulo.query.close;
    frmmodulo.query.sql.clear;
    frmmodulo.query.sql.add('select * from adm');
    frmmodulo.query.sql.add('where upper(info1) like ' + QuotedStr(usuario));
    frmmodulo.query.sql.add('and info2 = ''' + senha + '''');
    frmmodulo.query.Open;
    frmmodulo.qrConfigPDV.Close;
    frmmodulo.qrConfigPDV.Open;
    if frmmodulo.query.recordcount > 0 then begin
      if (frmmodulo.qrConfigPDVCAIXA_SITUACAO.AsString <> 'FECHADO') and
         (frmmodulo.qrConfigPDVNUMCAIXA.AsInteger <> frmmodulo.query.FieldByName('codigo').AsInteger) and
         (frmmodulo.qrConfigPDVNUMCAIXA.AsInteger > 0) then begin
        frmmodulo.Query.close;
        frmmodulo.Query.sql.clear;
        frmmodulo.Query.sql.add('select * from adm');
        frmmodulo.Query.sql.add('where codigo = ' + frmmodulo.qrConfigPDVNUMCAIXA.AsString);
        frmmodulo.Query.Open;
        if frmmodulo.qrConfigPDVNUMCAIXA.AsInteger = 99 then begin
          Application.MessageBox(PWideChar('Existe um caixa aberto para o operador "'+frmmodulo.Query.FieldByName('INFO1').AsString +'"!'+#13+'Este operador usa o Caixa 99 Geral.'+#13+'De acordo com as novas regras do sistema n„o È permitido utilizar este caixa nos PDVs.'+#13+'Ser· necess·rio realizar o fechamento do caixa deste operador neste terminal para poder logar com outro operador!'),'AtenÁ„o!',MB_ICONINFORMATION);

          frmSenha_Supervisor := TfrmSenha_Supervisor.create(Application);
          frmSenha_Supervisor.FechaCaixa := True;
          frmSenha_Supervisor.showmodal;

          if not bSupervisor_autenticado then
            Exit;

          if bFechamentoCego then begin
            frmFechamentoCego := TfrmFechamentoCego.create(Application);
            frmFechamentoCego.showmodal;
          end else begin
            frmCaixa_Fechamento := TfrmCaixa_Fechamento.create(Application);
            frmCaixa_Fechamento.showmodal;
          end;

        end else begin
          Application.MessageBox(PWideChar('Existe um caixa aberto para o operador "'+frmmodulo.query.FieldByName('INFO1').AsString +'", È necess·rio fechar este caixa!'+#13+#13+'Acesse o pdv com o Operador "'+frmmodulo.query.FieldByName('INFO1').AsString +'" e faÁa o fechamento do caixa para depois poder acessar com outro operador!'),'AtenÁ„o!',MB_ICONINFORMATION);
          Result := False;
          Exit;
        end;
      end else begin
        if frmmodulo.query.FieldByName('codigo').AsInteger = 99 then begin
          Application.MessageBox(PWideChar('Este operador esta configurado no Gerencial como Caixa 99 "Caixa Geral"!'+#13+#13+'De acordo com as novas regras do sistema n„o È mais permitido acessar os pdvs com o operadores cadastramos como Caixa 99!'),'AtenÁ„o!',MB_ICONINFORMATION);
          Result := False;
          Exit;
        end else begin
          frmmodulo.qrConfigPDV.Edit;
          frmmodulo.qrConfigPDV.FieldByName('numcaixa').AsInteger := frmmodulo.query.FieldByName('codigo').AsInteger;
          frmmodulo.qrConfigPDV.post;
        end;

      end;
      if funcao = '001' then // entrada no sistema e troca de operador
      begin
        frmmodulo.query.close;
        frmmodulo.query.sql.clear;
        frmmodulo.query.sql.add('select * from adm');
        frmmodulo.query.sql.add('where upper(info1) like ' + QuotedStr(usuario));
        frmmodulo.query.sql.add('and info2 = ''' + senha + '''');
        frmmodulo.query.Open;
        icodigo_Usuario := frmPrincipal.IsInteger(frmmodulo.query.fieldbyname('codigo').asstring);
        sCaixa := Zerar(IntToStr(icodigo_Usuario), 3);
        sNome_Operador := frmmodulo.query.fieldbyname('info1').asstring;
        bFechamentoCego := frmmodulo.query.fieldbyname('fechamento_cego').asstring = 'S';
      end;
      Result := True;
    end;

  end;
end;

function Cript(Action, Src: string): string;
label
  Fim;
var
  KeyLen: Integer;
  KeyPos: Integer;
  OffSet: Integer;
  Dest, key: string;
  SrcPos: Integer;
  SrcAsc: Integer;
  TmpSrcAsc: Integer;
  Range: Integer;
begin
  if (Src = '') then begin
    Result := '';
    goto Fim;
  end;
  key := 'YUQL23K1PIUF90AERVNNMKH02NMIHJ12042 E18XM01HIBQAS150AVDOUYQA90UD1APSA12POIENC1K3210N0419RTIKJ';

  Dest := '';
  KeyLen := length(key);
  KeyPos := 0;
  SrcPos := 0;
  SrcAsc := 0;
  Range := 256;
  if (Action = UpperCase('C')) then begin
    Randomize;
    OffSet := Random(Range);
    Dest := Format('%1.2x', [OffSet]);
    for SrcPos := 1 to length(Src) do begin
      application.ProcessMessages;
      SrcAsc := (Ord(Src[SrcPos]) + OffSet) mod 255;
      if KeyPos < KeyLen then
        KeyPos := KeyPos + 1
      else
        KeyPos := 1;
      SrcAsc := SrcAsc xor Ord(key[KeyPos]);
      Dest := Dest + Format('%1.2x', [SrcAsc]);
      OffSet := SrcAsc;
    end;
  end
  else if (Action = UpperCase('D')) then begin
    OffSet := StrToInt('$' + Copy(Src, 1, 2));
    SrcPos := 3;
    repeat
      SrcAsc := StrToInt('$' + Copy(Src, SrcPos, 2));
      if (KeyPos < KeyLen) then
        KeyPos := KeyPos + 1
      else
        KeyPos := 1;
      TmpSrcAsc := SrcAsc xor Ord(key[KeyPos]);
      if TmpSrcAsc <= OffSet then
        TmpSrcAsc := 255 + TmpSrcAsc - OffSet
      else
        TmpSrcAsc := TmpSrcAsc - OffSet;
      Dest := Dest + Chr(TmpSrcAsc);
      OffSet := SrcAsc;
      SrcPos := SrcPos + 2;
    until (SrcPos >= length(Src));
  end;
  Result := Dest;
Fim:

end;

function Zerar(texto: string; qtde: Integer): string;
begin
  while length(texto) < qtde do
    texto := '0' + texto;
  Result := texto;
end;

function texto_justifica(texto: string; qtde_caracteres: Integer; caracter: string; Alinhamento: TAlinhamento): string;
begin
  texto := tiraacento(texto);
  if Alinhamento = taDireita then begin
    while length(texto) < qtde_caracteres do
      texto := caracter + texto;
    while length(texto) > qtde_caracteres do
      delete(texto, (qtde_caracteres + 1), 1);
  end;
  if Alinhamento = taEsquerda then begin
    while length(texto) < qtde_caracteres do
      texto := texto + caracter;
    while length(texto) > qtde_caracteres do
      delete(texto, (qtde_caracteres + 1), 1);
  end;
  if Alinhamento = taCentralizado then begin
    if (qtde_caracteres mod 2) <> 0 then
      qtde_caracteres := qtde_caracteres - 1;
    while length(texto) < qtde_caracteres do
      texto := caracter + texto + caracter;
    while length(texto) > qtde_caracteres do
      delete(texto, (qtde_caracteres + 1), 1);
  end;
  Result := texto;
end;

function tiraacento(Str: string): string;
var
  i: Integer;
begin
  for i := 1 to length(Str) do
    case Str[i] of
      'Ë':
        Str[i] := 'e';
      'Ï':
        Str[i] := 'i';
      '˘':
        Str[i] := 'u';
      'Ó':
        Str[i] := 'i';
      '˚':
        Str[i] := 'u';
      '·':
        Str[i] := '†'; //
      'È':
        Str[i] := 'Ç'; //
      'Ì':
        Str[i] := '°'; //
      'Û':
        Str[i] := '¢'; //
      '˙':
        Str[i] := '£'; //
      '‡':
        Str[i] := 'Ö'; //
      'Ú':
        Str[i] := 'ï'; //
      '‚':
        Str[i] := 'É'; //
      'Í':
        Str[i] := 'à'; //
      'Ù':
        Str[i] := 'ì'; //
      '‰':
        Str[i] := 'Ñ'; //
      'Î':
        Str[i] := 'â'; //
      'Ô':
        Str[i] := 'ã'; //
      'ˆ':
        Str[i] := 'î'; //
      '¸':
        Str[i] := 'Å'; //
      '„':
        Str[i] := 'Ü'; //
      'ı':
        Str[i] := 'o';
      'Ò':
        Str[i] := '§'; //
      'Á':
        Str[i] := 'á'; //
      '¡':
        Str[i] := 'A';
      '…':
        Str[i] := 'ê'; //
      'Õ':
        Str[i] := 'I';
      '”':
        Str[i] := 'O';
      '⁄':
        Str[i] := 'U';
      '¿':
        Str[i] := 'A';
      '»':
        Str[i] := 'E';
      'Ã':
        Str[i] := 'I';
      '“':
        Str[i] := 'O';
      'Ÿ':
        Str[i] := 'U';
      '¬':
        Str[i] := 'è';
      ' ':
        Str[i] := 'E';
      'Œ':
        Str[i] := 'I';
      '‘':
        Str[i] := 'O';
      '€':
        Str[i] := 'U';
      'ƒ':
        Str[i] := 'A';
      'À':
        Str[i] := 'E';
      'œ':
        Str[i] := 'I';
      '÷':
        Str[i] := 'O';
      '‹':
        Str[i] := 'ö'; //
      '√':
        Str[i] := 'A';
      '’':
        Str[i] := 'O';
      '—':
        Str[i] := '•'; //
      '«':
        Str[i] := 'Ä'; //
      '™':
        Str[i] := '¶'; //
      '∫':
        Str[i] := 'ß'; //
    end;
  Result := Str;
end;

// -------------------------------------------------------------------------- //
function Tiraacento_display(Str: string): string;
var
  i: Integer;
begin
  for i := 1 to length(Str) do
    case Str[i] of
      'Ë':
        Str[i] := 'e';
      'Ï':
        Str[i] := 'i';
      '˘':
        Str[i] := 'u';
      'Ó':
        Str[i] := 'i';
      '˚':
        Str[i] := 'u';
      '·':
        Str[i] := 'a'; //
      'È':
        Str[i] := 'e'; //
      'Ì':
        Str[i] := 'i'; //
      'Û':
        Str[i] := 'o'; //
      '˙':
        Str[i] := 'u'; //
      '‡':
        Str[i] := 'a'; //
      'Ú':
        Str[i] := 'o'; //
      '‚':
        Str[i] := 'a'; //
      'Í':
        Str[i] := 'e'; //
      'Ù':
        Str[i] := 'o'; //
      '‰':
        Str[i] := 'a'; //
      'Î':
        Str[i] := 'e'; //
      'Ô':
        Str[i] := 'i'; //
      'ˆ':
        Str[i] := 'o'; //
      '¸':
        Str[i] := 'u'; //
      '„':
        Str[i] := 'a'; //
      'ı':
        Str[i] := 'o';
      'Ò':
        Str[i] := 'n'; //
      'Á':
        Str[i] := 'c'; //
      '¡':
        Str[i] := 'A';
      '…':
        Str[i] := 'E'; //
      'Õ':
        Str[i] := 'I';
      '”':
        Str[i] := 'O';
      '⁄':
        Str[i] := 'U';
      '¿':
        Str[i] := 'A';
      '»':
        Str[i] := 'E';
      'Ã':
        Str[i] := 'I';
      '“':
        Str[i] := 'O';
      'Ÿ':
        Str[i] := 'U';
      '¬':
        Str[i] := 'A';
      ' ':
        Str[i] := 'E';
      'Œ':
        Str[i] := 'I';
      '‘':
        Str[i] := 'O';
      '€':
        Str[i] := 'U';
      'ƒ':
        Str[i] := 'A';
      'À':
        Str[i] := 'E';
      'œ':
        Str[i] := 'I';
      '÷':
        Str[i] := 'O';
      '‹':
        Str[i] := 'U'; //
      '√':
        Str[i] := 'A';
      '’':
        Str[i] := 'O';
      '—':
        Str[i] := 'N'; //
      '«':
        Str[i] := 'C'; //
      '™':
        Str[i] := 'a'; //
      '∫':
        Str[i] := 'o'; //
    end;
  Result := Str;
end;

function UltimoDiaMes(Mdt: TDateTime): TDateTime;
{ retorna o ultimo dia o mes, de uma data fornecida }
var
  ano, mes, dia: word;
  mDtTemp: TDateTime;
begin
  Decodedate(Mdt, ano, mes, dia);
  mDtTemp := (Mdt - dia) + 33;
  Decodedate(mDtTemp, ano, mes, dia);
  Result := mDtTemp - dia;
end;

// -------------------------------------------------------------------------- //
// GUIO: Faz a leitura uma informaÁ„o na tabela CONFIG
function LerParametro(psNome: string; psTipo: string; pvDefault: Variant): Variant;
var
  svalor: string;
begin
  with frmmodulo.query do begin
    try
      if Active then
        close;

      sql.Text := 'Select PARNOM, PARVAL, PARTIP FROM CONFIG WHERE PARNOM = ' + QuotedStr(psNome);

      Open;

      if IsEmpty then begin
        svalor := RetornarParametroString(psTipo, pvDefault);

        GravarParametro(psNome, psTipo, svalor);
        Result := LerParametro(psNome, psTipo, pvDefault);
      end
      else begin
        try
          if UpperCase(psTipo) = 'C' then
            Result := fieldbyname('PARVAL').asstring
          else if (UpperCase(psTipo) = 'N') then
            Result := strtofloat(fieldbyname('PARVAL').asstring)
          else if (UpperCase(psTipo) = 'D') then
            Result := strtodate(fieldbyname('PARVAL').asstring)
          else if (UpperCase(psTipo) = 'I') then
            Result := StrToInt(fieldbyname('PARVAL').asstring)
          else if (UpperCase(psTipo) = 'B') then
            Result := (fieldbyname('PARVAL').asstring = '1')
          else
            Result := fieldbyname('PARVAL').asstring;
        except
          application.messagebox('Erro de leitura de par‚metro da funÁ„o LerParametro()', 'Erro', mb_ok + mb_iconerror);

          Abort;
        end;
      end;
    finally
      close;
    end;
  end;
end;

// -------------------------------------------------------------------------- //
// GUIO: Grava uma informaÁ„o na tabela CONFIG
procedure GravarParametro(psNome: string; psTipo: string; pvValor: Variant);
var
  svalor: string;
begin
  svalor := RetornarParametroString(psTipo, pvValor);

  with frmmodulo.query do begin
    try
      if Active then
        close;

      sql.Text := 'Select PARNOM, PARVAL, PARTIP FROM CONFIG WHERE PARNOM = ' + QuotedStr(psNome);
      Open;

      if IsEmpty then begin
        close;
        sql.Text := 'INSERT INTO CONFIG (CODIGO, PARNOM, PARVAL, PARTIP) VALUES (1, ' + QuotedStr(psNome) + ', ' + QuotedStr(svalor) + ', ' + QuotedStr(psTipo) + ')';

        if not Prepared then
          Prepare;

        ExecSQL;
      end
      else begin
        close;
        sql.Text := 'UPDATE CONFIG SET ' + ' PARVAL = ' + QuotedStr(svalor) + ', ' + ' PARTIP = ' + QuotedStr(psTipo) + ' WHERE PARNOM = ' + QuotedStr(psNome);

        if not Prepared then
          Prepare;

        ExecSQL;
      end;
    finally
      close;
    end;
  end;
end;

// -------------------------------------------------------------------------- //
// GUIO: Retornar o valor passado em forma de string para poder ser armazenado
function RetornarParametroString(psTipo: string; pvValor: Variant): string;
begin
  if UpperCase(psTipo) = 'C' then
    Result := pvValor
  else if (UpperCase(psTipo) = 'N') then begin
    try
      Result := FloatToStr(pvValor);
    except
      Result := '0.00';
    end;
  end
  else if (UpperCase(psTipo) = 'D') then begin
    try
      Result := DateToStr(pvValor);
    except
      Result := DateToStr(Date);
    end
  end
  else if (UpperCase(psTipo) = 'I') then begin
    try
      Result := inttostr(pvValor);
    except
      Result := '0';
    end;
  end
  else if (UpperCase(psTipo) = 'B') then begin
    try
      if pvValor then
        Result := '1' // 0 - Falso, 1 - Verdadeiro.  D„„mmm!!!
      else
        Result := '0';

    except
      Result := '0';

    end;
  end
  else
    Result := '';

end;

procedure Gaveta_Configura;
begin
  with frmModulo do begin
    cpGaveta.Modelo := TACBrPosPrinterModelo(GavModelo);
    cpGaveta.Porta  := GavPorta;
    cpGaveta.CortaPapel := GavCortaPapel;
    if GavTipoCorte = 'Total' then
      cpGaveta.TipoCorte := ctTotal
    else
      cpGaveta.TipoCorte := ctParcial;
    cpGaveta.ConfigGaveta.SinalInvertido := GavInvertido;
  end;
end;

Procedure Gaveta_Abre;
begin
  with frmModulo do begin
    try
      cpGaveta.Desativar ;
      Gaveta_Configura;
      cpGaveta.Ativar;
      cpGaveta.AbrirGaveta;
      cpGaveta.Desativar;
    except
      on e:Exception do begin
        Application.MessageBox(PWideChar('Ocorreu um erro ao tentar abrir a gaveta:'+#13+E.Message),'AtenÁ„o!',MB_ICONINFORMATION);
      end;
    end;
  end;
end;


end.

