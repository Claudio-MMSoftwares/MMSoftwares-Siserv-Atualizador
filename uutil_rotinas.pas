unit uutil_rotinas;

interface

uses
   IniFiles, Registry, SysUtils, Forms, Clipbrd, DB,
   Graphics, ComCtrls, Windows, WinSock, ComObj, Classes, DBClient, DbGrids,
   DBXJson, System.JSON, Math, IdHash, IdHashSHA, IdGlobal,
   Variants, Dialogs, ShellAPI;

type
  TSisVersion = Record
    ID          : Integer;
    Descricao   : String;
    Sigla       : String;
    MajorVersion: Integer;
    MinorVersion: Integer;
    Release     : Integer;
    Build       : Integer;
  end;

  TDBConfig = Record
    UserDB   : String;
    PassDB   : String;
    ServerDB : String;
    PathDB   : string;
    PortaDB  : string;
    DialectDB: Integer;
    RoleDB   : string;
    CharsetDB: string;
  end;

  TProcedure1 = procedure(sAux: String) of object;

  //=================== Arquivos Ini =================
  procedure GravaConf(Secao, Chave, Valor: string; bWorkingDirectory: Boolean=False);
  function LerConf(sSecao,sChave: String; sPadrao: String=''; bWorkingDirectory: Boolean=False): String;
  //==================================================
  function iif(bCondicao: Boolean; vValorTrue, vValorFalse: Variant): Variant; overload;

  function ExtractFilePath_MM(const FileName: string): string;
  function ExtractFileName_MM(const FileName: string): string;

  function decrypt(valor: String):String;
  function encrypt(valor: String):String;

  function StringToList(Origem:String;Delimitador:String=''): TStringList;
  function ListToString(Lista:TStringList;Delimitador:String=''):String;

  // Retorno o usuário logado no windows
  function UsuarioWindows: String;
  // Retorna IP e nome da máquina
  function GetIPFromHost(var HostName, IPaddr, WSAError: String): Boolean;

  // Exemplo: GravarRegistro(HKEY_CURRENT_USER, 'Software\Tantra Software', 'Username', 'KLAUS')
  procedure GravarRegistro(RootKey: HKEY; sChave, sStr, sValor: String);
  // Exemplo: LerRegistro(HKEY_CURRENT_USER, 'Software\Tantra Software', 'Username', UsuarioWindows)
  function LerRegistro(RootKey: HKEY; sChave, sStr, sValorPadrao: String): String;

  function ExportToCSV(DataSet: TDataSet; const FileName: string): Boolean;
  function ExcelDisponivel: Boolean;

  procedure CopyToClipboard(sStr: String);



  function getVersaoArquivo(FileName: PChar): TSisVersion;

  procedure FormSempreVisivel(Form: TForm);

  function RemoverEnter(sStr: String): String;

  function RetornaSoNumero(sStr: String): string;

  function RetornaSoLetraNumero(Str:String; bFlagAllTrim: Boolean=False; bUppercase:Boolean=True):String;
  function RetornarLetraNumero(sStr: String): String;

  function AnsiToAscii(str: String): String;
  function StrZero(Variavel: String; QUANT:integer; Onde:Char): String;

  function CriarDiretorio(Dir: String): String;
  procedure CopyFile(const sourcefilename, targetfilename: String);
  procedure DeleteArq(Arquivo: String); overload;
  procedure DeleteArq(slArquivos: TStringList); overload;

  function getIP: String;
  function ValidaPath(sPath: String; sDelim: String=''):String;
  function CopyAtPos(S: String; Initial, Final: Integer): String;

  function ValidaEmail(sEmail: String): Boolean;
  function ValidarEmails(sEmail: String; bMostrarAlerta: Boolean=True; bFixSeparator: Boolean=False): Boolean;
  function EmailsValidos(sEmail: String): String;
  function RetornaCondicao(Condicao: String): String;
  function RetornaFimLike(Condicao: String): String;
  function BuscaCodIBGEUF(UF: String):String;

  function AllTrimTexto(sTexto: string): string;

  function PrimeiroNome(Nome: String): String;

  function  PegaTempDir : String;//Pega o diretorio da Pasta Temporaria
  function getCamposJsonString(json,value:String): String;
  function FormatarpCnpjCpf(const pCnpjCpf: string): string;
  function FormatarCEP(const pCEP: string): string;
  function DataParaSQL(pData: TDateTime; pNullSeZero: Boolean = False): string;
  procedure SalvarLog(const NomeArquivo, Conteudo: string);
  procedure GerarArquivoLog(const pDirSistema, pMsg, pGateway: string);

  function UFtoCUF(const UF: String): Integer;
  function CUFtoUF(CUF: Integer): String;

  function IsExcelInstalled: Boolean;
  function RemoverCaracteres(const texto: string): string;
  function ArredondaParaCima(const Valor: Extended; CasasDecimais: Integer): Extended;
  function GetStr__virgula(var Line: String): String;
  function RemoveLineBreaks(const Text: string): string;
  function ReplaceNonStandardDashes(const Input: string): string;
  function SubstituirVirgulaPorPonto(const texto: string): string;
  function ConverterPontoParaVirgula(Valor: String): String;
  function ConvertStringToFloat(const AValue: string): Extended;
  function TryIntToStr(pString: string; pValorPadrao: Integer): Integer;



{
  function StringComponentToBinaryStream(sStr: String): TStream;
  function ComponentToString(Component: TComponent): string;
}
const
  DiaDoMes: array[1..12] of Integer = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
  NL = #13#10;
  MinusculaSemAcentoF: array[1..19] of Char = ('a','o','a','e','i','o','u','a','e','i','o','u','a','e','i','o','u','c','n');
  MaiusculaSemAcentoF: array[1..19] of Char = ('A','O','A','E','I','O','U','A','E','I','O','U','A','E','I','O','U','C','N');
  MinusculaAcentuadaF: array[1..19] of Char = ('ã','õ','á','é','í','ó','ú','à','è','ì','ò','ù','â','ê','î','ô','û','ç','ñ');
  MaiusculaAcentuadaF: array[1..19] of Char = ('Ã','Õ','Á','É','Í','Ó','Ú','À','È','Ì','Ò','Ù','Â','Ê','Î','Ô','Û','Ç','Ñ');

  DFeUF: array[0..26] of String =
  ('AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA',
   'PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO');

  DFeUFCodigo: array[0..26] of Integer =
  (12,27,16,13,29,23,53,32,52,21,51,50,31,15,25,41,26,22,33,24,43,11,14,42,35,28,17);

implementation

function ExcelDisponivel: Boolean;
var
  Excel: Variant;
begin
  Result := False;
  try
    try
      Excel := CreateOleObject('Excel.Application');
      Result := True;
      Excel.Quit;
    except
      Result := False;
    end;
  finally
    Excel := Unassigned;
  end;
end;

function ExportToCSV(DataSet: TDataSet; const FileName: string): Boolean;
var
  F: TextFile;
  I: Integer;
  S: string;
begin
  Result := False;
  AssignFile(F, FileName);
  try
    Rewrite(F);

    // Cabeçalhos
    S := '';
    for I := 0 to DataSet.FieldCount - 1 do
    begin
      if DataSet.Fields[I].Visible then
        S := S + DataSet.Fields[I].FieldName + ';';
    end;
    if S <> '' then
      Delete(S, Length(S), 1);
    Writeln(F, S);

    // Dados
    DataSet.DisableControls;
    try
      DataSet.First;
      while not DataSet.Eof do
      begin
        S := '';
        for I := 0 to DataSet.FieldCount - 1 do
        begin
          if DataSet.Fields[I].Visible then
            S := S + DataSet.Fields[I].AsString + ';';
        end;
        if S <> '' then
          Delete(S, Length(S), 1);
        Writeln(F, S);
        DataSet.Next;
      end;
      Result := True;
    finally
      DataSet.EnableControls;
    end;
  finally
    CloseFile(F);
  end;
end;


function TryIntToStr(pString: string; pValorPadrao: Integer): Integer;
begin
  try
    Result := StrToInt(pString);
  except
    Result := pValorPadrao;
  end;
end;

function ConverterPontoParaVirgula(Valor: String): String;
begin
  Result := StringReplace(Valor, '.', ',', [rfReplaceAll]);
end;

function ConvertStringToFloat(const AValue: string): Extended;
begin
  // Tenta converter a string em um Float
  try
    Result := StrToFloat(AValue);
  except
    on E: EConvertError do
      begin
        // Em caso de erro, você pode tratar a exceção
        Result := 0; // ou outra lógica de erro
      end;
  end;
end;

function SubstituirVirgulaPorPonto(const texto: string): string;
var
  i: Integer;
begin
  Result := texto;

  // Primeiro, removemos todos os pontos
  for i := Length(Result) downto 1 do
  begin
    if Result[i] = '.' then
      Delete(Result, i, 1);
  end;

  // Depois, substituímos todas as vírgulas por pontos
  for i := 1 to Length(Result) do
  begin
    if Result[i] = ',' then
      Result[i] := '.';
  end;
end;


function ReplaceNonStandardDashes(const Input: string): string;
const
  EnDash = #$2013;  // Unicode en dash
  EmDash = #$2014;  // Unicode em dash
var
  I: Integer;
  ResultStr: string;
begin
  ResultStr := Input;

  // Replace en dash and em dash with a hyphen
  for I := 1 to Length(ResultStr) do
  begin
    if ResultStr[I] = EnDash then
      ResultStr[I] := '-'
    else if ResultStr[I] = EmDash then
      ResultStr[I] := '-';
  end;

  Result := ResultStr;
end;

function RemoveLineBreaks(const Text: string): string;
begin
  // Remove CRLF (Windows style)
  Result := StringReplace(Text, #13#10, '', [rfReplaceAll]);
  // Remove CR (Mac style)
  Result := StringReplace(Result, #13, '', [rfReplaceAll]);
  // Remove LF (Unix style)
  Result := StringReplace(Result, #10, '', [rfReplaceAll]);
end;

function GetStr__virgula(var Line: String): String;
var
TipoSep : String;
begin
  if copy(Line, 1, 1) = ',' then
     Result := ''
  else begin
    Result := Copy(Line, 1, Pos(',', Line)-1);
    if Result <> '' then
      if Result[1] = '"' then
        Result := Copy(Result, 2, Length(Result)-2);
  end;
  Delete(Line, 1, Pos(',',Line));
end;

function ArredondaParaCima(const Valor: Extended; CasasDecimais: Integer): Extended;
var
  Fator: Extended;
begin
  // Calcula o fator de multiplicação para arredondar o número
  Fator := Power(10, CasasDecimais);

  // Multiplica o valor pelo fator, arredonda para cima e divide novamente
  Result := Ceil(Valor * Fator) / Fator;
end;

function RemoverCaracteres(const texto: string): string;
begin
  // Substituir os caracteres '-' e ' ' (espaço em branco) por uma string vazia
  Result := StringReplace(texto, '-', '', [rfReplaceAll]);
  Result := StringReplace(Result, ' ', '', [rfReplaceAll]);
end;

function IsExcelInstalled: Boolean;
var
  Reg: TRegistry;
begin
  Result := False;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Result := Reg.OpenKeyReadOnly('Excel.Application');
  finally
    Reg.Free;
  end;
end;

function UFtoCUF(const UF: String): Integer;
var
  i: Integer;
begin
  Result := -1 ;
  for i:= Low(DFeUF) to High(DFeUF) do
  begin
    if DFeUF[I] = UF then
    begin
      Result := DFeUFCodigo[I];
      exit;
    end;
  end;
end;

function CUFtoUF(CUF: Integer): String;
var
  i: Integer;
begin
  Result := '' ;
  for i:= Low(DFeUFCodigo) to High(DFeUFCodigo) do
  begin
    if DFeUFCodigo[I] = CUF then
    begin
      Result := DFeUF[I];
      exit;
    end;
  end;
end;

procedure GerarArquivoLog(const pDirSistema,  pMsg, pGateway: string);
var
  LogFileName: string;
  LogFile: TextFile;
  lDir: string;
begin
  lDir := pDirSistema + 'AVERBSEGURO\';
  if not DirectoryExists(lDir) then
    ForceDirectories(lDir);

  LogFileName := lDir + pGateway+'.log';
  // Especifique o nome do arquivo de log e o caminho desejado
  try
    // Abre o arquivo de log em modo de escrita, criando-o se não existir
    AssignFile(LogFile, LogFileName);
    if FileExists(LogFileName) then
      Append(LogFile)
    else
      Rewrite(LogFile);
    // Escreve a mensagem no arquivo de log com data e hora
    Writeln(LogFile, FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ': ' + pMsg);
  finally
    // Fecha o arquivo
    CloseFile(LogFile);
  end;
end;

procedure SalvarLog(const NomeArquivo, Conteudo: string);
var
  ArquivoLog: TStringList;
begin
  ArquivoLog := TStringList.Create;
  try
    if FileExists(NomeArquivo) then
      ArquivoLog.LoadFromFile(NomeArquivo);
    ArquivoLog.Add(Conteudo);
    ArquivoLog.SaveToFile(NomeArquivo);
  finally
    // Libera a TStringList da memória
    ArquivoLog.Free;
  end;
 end;

function DataParaSQL(pData: TDateTime; pNullSeZero: Boolean = False): string;
begin
  if (pData > 0) then
    Result := FormatDateTime('dd.mm.yyyy', pData)
  else if not pNullSeZero then
    Result := '01.01.0001'
  else
    Result := 'null';

  if (Result <> 'null') then
    Result := Result.QuotedString;
end;

function FormatarpCnpjCpf(const pCnpjCpf: string): string;
var
  lCnpjCpf: string;
  li: Integer;
begin
  lCnpjCpf := '';
  for li := 1 to Length(pCnpjCpf) do
  begin
    if CharInSet(pCnpjCpf[li], ['0'..'9']) then
      lCnpjCpf := lCnpjCpf + pCnpjCpf[li];
  end;

  if Length(lCnpjCpf) = 11 then
  begin
    Result := Copy(lCnpjCpf, 1, 3) + '.' +
              Copy(lCnpjCpf, 4, 3) + '.' +
              Copy(lCnpjCpf, 7, 3) + '-' +
              Copy(lCnpjCpf, 10, 2);
  end
  else if Length(lCnpjCpf) = 14 then
  begin
    Result := Copy(lCnpjCpf, 1, 2) + '.' +
              Copy(lCnpjCpf, 3, 3) + '.' +
              Copy(lCnpjCpf, 6, 3) + '/' +
              Copy(lCnpjCpf, 9, 4) + '-' +
              Copy(lCnpjCpf, 13, 2);
  end
  else
    Result := '';
end;

function FormatarCEP(const pCEP: string): string;
var
  lCEP: string;
  li: Integer;
begin
  lCEP := '';

  LCEP := RetornaSoNumero(pCEP);
  if Length(lCEP) = 8 then
  begin
    Result := Copy(lCEP, 1, 5) + '-' +
              Copy(lCEP, 6, 3);
  end
  else
    Result := '00000-000';
end;

function getCamposJsonString(json,value:String): String;
var
LJSONObject: TJSONObject;

   function TrataObjeto(jObj:TJSONObject):string;
   var i:integer;
       jPar: TJSONPair;
   begin
        result := '';
        for i := 0 to jObj.Size - 1 do
        begin
             jPar := jObj.Get(i);
             if jPar.JsonValue Is TJSONObject then
                result := TrataObjeto((jPar.JsonValue As TJSONObject)) else
             if sametext(trim(jPar.JsonString.Value),value) then
             begin
                  Result := jPar.JsonValue.Value;
                  break;
             end;
             if result <> '' then
                break;
        end;
   end;
begin
   try
      LJSONObject := nil;
      LJSONObject := TJSONObject.ParseJSONValue(TEncoding.ASCII.GetBytes(json),0) as TJSONObject;
      result := TrataObjeto(LJSONObject);
   finally
      LJSONObject.Free;
   end;
end;

function PrimeiroNome(Nome: String): String;
var
  PNome: String;
begin
  PNome := '';

  if pos(' ', Nome) <> 0 then
    PNome := copy(Nome, 1, pos(' ', Nome) - 1)
  else
    PNome := Nome;

  Result := Trim(PNome);
end;
{
function ComponentToString(Component: TComponent): string;
var
  BinStream: TMemoryStream;
  StrStream: TStringStream;
begin
  BinStream := TMemoryStream.Create;
  try
    StrStream := TStringStream.Create(Result);
    try
      BinStream.WriteComponent(Component);
      BinStream.Seek(0, soFromBeginning);
      ObjectBinaryToText(BinStream, StrStream);
      StrStream.Seek(0, soFromBeginning);
      Result := StrStream.DataString;
    finally
      StrStream.Free;
    end;
  finally
    BinStream.Free
  end;
end;

function StringComponentToBinaryStream(sStr: String): TStream;
var
  BinStream: TMemoryStream;
  StrStream: TStringStream;
begin
  BinStream := TMemoryStream.Create;
  StrStream := TStringStream.Create;
  try
    StrStream.WriteString(sStr);
    StrStream.Seek(0, soFromBeginning);
    ObjectTextToBinary(StrStream, BinStream);
    BinStream.Seek(0, soFromBeginning);
    Result := BinStream;
  finally
    StrStream.Free;
//    BinStream.Free
  end;
end;
 }


function  PegaTempDir : String;//Pega o diretorio da Pasta Temporaria
var DiretorioTemp : PChar;
    TempBuffer    : Dword;
begin
  TempBuffer := 255;
  GetMem(DiretorioTemp,255);
  try
    GetTempPath(tempbuffer,diretoriotemp);
    result := DiretorioTemp;
  finally
    FreeMem(diretoriotemp);
  end;
end;

 function BuscaCodIBGEUF(UF: String):String;
const
  Estados : String = 'AC12,AL27,AM13,AP16,BA29,CE23,DF53,ES32,GO52,MA21,MG31,MS50,MT51,'+
                     'PA15,PB25,PE26,PI22,PR41,RJ33,RN24,RO11,RR14,RS43,SC42,SE28,SP35,TO17';
begin
  Result := copy(Estados, pos(UF, Estados)+2, 2);
end;

function RetornaCondicao(Condicao: String): String;
begin
    if Condicao = 'Igual' then
       Result := ' = '''
    else
     if Condicao = 'Diferente' then
        Result := ' <> '''
     else
      if Condicao = 'Maior' then
         Result := ' > '''
      else
       if Condicao = 'Menor' then
          Result := ' < '''
       else
        if Condicao = 'Contem' then
           Result := ' like ''%'
        else
         if Condicao = 'Não Contem' then
            Result := ' not like ''%'
         else
          if Condicao = 'Maior ou igual' then
             Result := ' >= '''
          else
           if Condicao = 'Menor ou igual' then
              Result := ' <= '''
           else
            if Condicao = 'Começa por' then
               Result := ' like '''
            else
             if Condicao = 'Não Começa por' then
                Result := ' not like '''
            else
                Result := ' = ''';
end;

function RetornaFimLike(Condicao: String): String;
begin
    if Condicao = 'Contem' then
       Result := '%'
    else
     if Condicao = 'Não Contem' then
        Result := '%'
     else
      if Condicao = 'Começa por' then
         Result := '%'
      else
       if Condicao = 'Não Começa por' then
          Result := '%'
       else
          Result := '';
end;

function ValidaEmail(sEmail: String): Boolean;
// Retorna True se o endereço de e-mail for válido
const
  // Caracteres válidos em um "atom"
  atom_chars = [#33..#255] - ['(', ')', '<', '>', '@', ',', ';', ':',
                              '\', '/', '"', '.', '[', ']', #127];
  // Caracteres válidos em um "quoted-string"
  quoted_string_chars = [#0..#255] - ['"', #13, '\'];

  // Caracteres válidos em um "sub-domain"
  letters = ['A'..'Z', 'a'..'z'];
  letters_digits = ['0'..'9', 'A'..'Z', 'a'..'z'];
  subdomain_chars = ['-', '0'..'9', 'A'..'Z', 'a'..'z'];
type
  States = (STATE_BEGIN, STATE_ATOM, STATE_QTEXT, STATE_QCHAR,
    STATE_QUOTE, STATE_LOCAL_PERIOD, STATE_EXPECTING_SUBDOMAIN,
    STATE_SUBDOMAIN, STATE_HYPHEN);
var
  State: States;
  i, n, subdomains: integer;
  c: char;
begin
  sEmail := StringReplace(sEmail, ' ', '', [rfReplaceAll]);

  State := STATE_BEGIN;
  n := Length(sEmail);
  i := 1;
  subdomains := 1;
  while (i <= n) do begin
    c := sEmail[i];
    case State of
    STATE_BEGIN:
      if c in atom_chars then
        State := STATE_ATOM
      else if c = '"' then
        State := STATE_QTEXT
      else
        break;
    STATE_ATOM:
      if c = '@' then
        State := STATE_EXPECTING_SUBDOMAIN
      else if c = '.' then
        State := STATE_LOCAL_PERIOD
      else if not (c in atom_chars) then
        break;
    STATE_QTEXT:
      if c = '\' then
        State := STATE_QCHAR
      else if c = '"' then
        State := STATE_QUOTE
      else if not (c in quoted_string_chars) then
        break;
    STATE_QCHAR:
      State := STATE_QTEXT;
    STATE_QUOTE:
      if c = '@' then
        State := STATE_EXPECTING_SUBDOMAIN
      else if c = '.' then
        State := STATE_LOCAL_PERIOD
      else
        break;
    STATE_LOCAL_PERIOD:
      if c in atom_chars then
        State := STATE_ATOM
      else if c = '"' then
        State := STATE_QTEXT
      else
        break;
    STATE_EXPECTING_SUBDOMAIN:
      if c in letters_digits {letters} then
        State := STATE_SUBDOMAIN
      else
        break;
    STATE_SUBDOMAIN:
      if c = '.' then begin
        inc(subdomains);
        State := STATE_EXPECTING_SUBDOMAIN
      end else if c = '-' then
        State := STATE_HYPHEN
      else if not (c in letters_digits) then
        break;
    STATE_HYPHEN:
      if c in letters_digits then
        State := STATE_SUBDOMAIN
      else if c <> '-' then
        break;
    end;
    inc(i);
  end;
  if i <= n then
    Result := False
  else
    Result := (State = STATE_SUBDOMAIN) and (subdomains >= 2);
end;


function ValidarEmails(sEmail: String; bMostrarAlerta: Boolean=True; bFixSeparator: Boolean=False): Boolean;
var
  slEmails: TStringList;
  i: Integer;
begin
  Result   := not (sEmail = '');

  if bFixSeparator then
  begin
    sEmail   := StringReplace(sEmail, ',', ';', [rfReplaceAll]);
    sEmail   := StringReplace(sEmail, ' ', '', [rfReplaceAll]);
  end;
  slEmails := StringToList(sEmail, ';');
  try
    for i := 0 to slEmails.Count - 1 do
    begin
      if not ValidaEmail(slEmails[i]) then
      begin
        if bMostrarAlerta then
          MessageBoxW(0, PWideChar(WideString('Email inválido: '+slEmails[i])), 'Aviso', MB_OK + MB_ICONWARNING + MB_TOPMOST);
        Abort;
      end;
    end;
  except
    Result := False;
  end;
end;


function EmailsValidos(sEmail: String): String;
var
  slEmails: TStringList;
  i: Integer;
begin
  Result := '';
  slEmails := StringToList(sEmail, ';');
  for i := 0 to slEmails.Count - 1 do
  begin
    if ValidaEmail(slEmails[i]) then
      Result := Result + slEmails[i]+';';
  end;
end;


function StringToList(Origem:String;Delimitador:String=''): TStringList;
var
  i:integer;
  sTmp:string;
begin
  if Delimitador='' then
    Delimitador:=';';
  Result:=TStringList.Create;
  sTmp:='';
  for i:=1 to length(Origem) do
  begin
    if (copy(Origem,i,1) <> Delimitador) then
      sTmp:=sTmp+copy(Origem,i,1);
    if (copy(Origem,i,1) = Delimitador) or (i=length(Origem))then
    begin
      if (copy(Origem,i,1) = Delimitador) and (i=length(Origem))then
      begin
        Result.Add(sTmp);
        sTmp:='';
      end
      else
      begin
        Result.Add(sTmp);
        sTmp:='';
      end;
    end;
    if (copy(Origem,i,1) = Delimitador) and (i=length(Origem))then
    begin
      Result.Add(sTmp);
      sTmp:='';
    end;
  end;
end;

function ListToString(Lista:TStringList;Delimitador:String=''):String;
var
 i:integer ;
begin
  if Delimitador='' then
    Delimitador:=';';
  Result:='';
  if Lista<> nil then
  begin
    for i:=0 to Lista.Count-1 do
    begin
      begin
        result:=result+Lista.Strings[i];
        if i<>Lista.Count-1 then
          result:=result+Delimitador;
      end;
    end;
  end;
end;

function CopyAtPos(S: String; Initial, Final: Integer): String;
var
  I: Integer;
begin
  Result := '';
  for I := Initial to Final do
    Result := Result + S[I];
end;

procedure FormSempreVisivel(Form: TForm);
begin
  SetWindowPos(Form.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
end;

function getIP: String;
//--> Declare a Winsock na clausula uses da unit
var
  WSAData: TWSAData;
  HostEnt: PHostEnt;
  Name: AnsiString;
begin
  WSAStartup(2, WSAData);
  SetLength(Name, 255);
  Gethostname(PAnsiChar(Name), 255);
  SetLength(Name, StrLen(PAnsiChar(Name)));
  HostEnt := gethostbyname(PAnsiChar(Name));
  with HostEnt^ do
  begin
    Result := Format('%d.%d.%d.%d',
    [Byte(h_addr^[0]),Byte(h_addr^[1]),
    Byte(h_addr^[2]),Byte(h_addr^[3])]);
  end;
  WSACleanup;
end;


function ValidaPath(sPath:String;sDelim:String=''):String;
begin
   Result:=sPath;
   if Length(sDelim)<=0 then
      sDelim:=PathDelim;
   if Copy(sPath,Length(sPath),1)<>sDelim then
     Result:=sPath+sDelim;
end;

//function GridToExcelFile(Grid: TDbAltGrid; ExcelFile: String; TotalRegistros: Integer;
//  pbAux: TObject; Abrir: Boolean; var bCancelarExportXLS: Boolean): Boolean;
//var
//  SavePlace: TBookmark;
//  i, j, iSheet, QuantasCol: Integer;
//  Excel, Sheet: Variant;
//  DataSet: TDataSet;
//  DataArray: Variant;
//  CorFundo: TColor;
//  VisibleColumns: array of Integer;
//begin
//  Result := False;
//  if not Assigned(Grid.DataSource) or not Grid.DataSource.DataSet.Active then Exit;
//
//  DataSet := Grid.DataSource.DataSet;
//  DataSet.DisableControls;
//  SavePlace := DataSet.GetBookmark;
//  DataSet.First;
//
//  QuantasCol := 0;
//  SetLength(VisibleColumns, Grid.Columns.Count);
//
//  // Contar colunas visíveis e armazenar índices
//  for i := 0 to Grid.Columns.Count - 1 do
//    if Grid.Columns[i].Visible then
//    begin
//      VisibleColumns[QuantasCol] := i;
//      Inc(QuantasCol);
//    end;
//
//  try
//    Excel := CreateOleObject('Excel.Application');
//    Excel.ScreenUpdating := False;
//    Excel.Calculation := -4135; // xlCalculationManual
//    Excel.Visible := False;
//    Excel.WorkBooks.Add;
//
//    iSheet := 1;
//    Sheet := Excel.WorkBooks[1].Sheets[iSheet];
//
//    // Criar array para armazenar dados
//    DataArray := VarArrayCreate([0, TotalRegistros, 0, QuantasCol - 1], varVariant);
//
//    // Adicionar cabeçalhos ao array
//    for i := 0 to QuantasCol - 1 do
//      DataArray[0, i] := Grid.Columns[VisibleColumns[i]].Title.Caption;
//
//    // Preencher array com os dados
//    j := 1;
//    while not DataSet.Eof and not bCancelarExportXLS do
//    begin
//      CorFundo := IfThen(j mod 2 = 0, $00F2F2F2, $00FBFBFB);
//      for i := 0 to QuantasCol - 1 do
//      begin
//        DataArray[j, i] := Grid.Columns[VisibleColumns[i]].Field.AsVariant;
//      end;
//      Inc(j);
//      DataSet.Next;
//    end;
//
//    // Passar array para o Excel de uma só vez (muito mais rápido)
//    Sheet.Range[Sheet.Cells[1, 1], Sheet.Cells[TotalRegistros + 1, QuantasCol]].Value := DataArray;
//
//    // Aplicar formatação separadamente para evitar operações lentas dentro do loop
//    for i := 1 to QuantasCol do
//    begin
//      Sheet.Cells[1, i].Font.Bold := True;
//      Sheet.Cells[1, i].Interior.Color := clGray;
//      Sheet.Cells[1, i].Font.Color := $0082F5FF;
//      Sheet.Cells[1, i].Borders.Color := clGray;
//    end;
//
//    // Aplicar formatação de número
//    for i := 1 to QuantasCol do
//    begin
//      case Grid.Columns[VisibleColumns[i - 1]].Field.DataType of
//        ftInteger, ftSmallInt:
//          Sheet.Columns[i].NumberFormat := '###,###,##0';
//        ftFloat, ftBCD, ftCurrency, ftFMTBcd:
//          Sheet.Columns[i].NumberFormat := '###,###,##0.00';
//      end;
//    end;
//
//    // Ajustar largura das colunas
//    Sheet.Columns.AutoFit;
//
//    // Restaurar configurações
//    Excel.Calculation := -4105; // xlCalculationAutomatic
//    Excel.ScreenUpdating := True;
//
//    // Salvar e exibir
//    Excel.WorkBooks[1].SaveAs(ExcelFile);
//    if Abrir then
//      Excel.Visible := True
//    else
//      Excel.Quit;
//
//    Result := True;
//  except
//    on E: Exception do
//    begin
//      Excel.Quit;
//      ShowMessage('Erro ao exportar para Excel: ' + E.Message);
//      Result := False;
//    end;
//  end;
//
//  DataSet.GotoBookmark(SavePlace);
//  DataSet.EnableControls;
//end;

function decrypt(valor: String):String;
var
  I   : Integer;
  Aux : String;
begin
  for I := 1 to length(valor) do
    Aux := Aux + chr(ord(valor[i])+20);
  Result := Aux;
end;

function encrypt(valor: String):String;
var
  I   : Integer;
  Aux : String;
begin
  for I := 1 to length(valor) do
    Aux := Aux + chr(ord(valor[i])-20);
  Result := Aux;
end;

procedure GravarRegistro(RootKey: HKEY; sChave, sStr, sValor: String);
var
  WinRegistro: TRegistry;
begin
  WinRegistro := TRegistry.Create;
  WinRegistro.RootKey := RootKey;

  try
    WinRegistro.OpenKey(sChave, True);
    try
      WinRegistro.WriteString(sStr, sValor);
    except
      raise;
    end;
  finally
    WinRegistro.CloseKey;
    WinRegistro.Free;
  end;
end;


function LerRegistro(RootKey: HKEY; sChave, sStr, sValorPadrao: String): String;
var
  WinRegistro: TRegistry;
begin
  Result := '';

  WinRegistro := TRegistry.Create;
  WinRegistro.RootKey := RootKey;

  try
    WinRegistro.OpenKey(sChave, True);
    try
      Result := WinRegistro.ReadString(sStr);
      if Result = '' then
        Result := sValorPadrao;
    except
      raise;
    end;
  finally
    WinRegistro.CloseKey;
    WinRegistro.Free;
  end;
end;


function UsuarioWindows: String;
var
  UserName: String;
  UserNameLen: DWord;
begin
  UserNameLen := 255;
  SetLength(UserName, UserNameLen);
  if GetUserName(PChar(UserName), UserNameLen) then
    Result := Copy(UserName, 1, UserNameLen-1)
  else
    Result := 'Desconhecido';
end;

function GetIPFromHost(var HostName, IPaddr, WSAError: String): Boolean;
type
  Name = array[0..100] of Char;
  PName = ^Name;
var
  HEnt: pHostEnt;
  HName: array[0..128] of char;//PName;
  WSAData: TWSAData;
  i: Integer;
begin
  Result := False;
  if WSAStartup($0101, WSAData) <> 0 then
  begin
    WSAError := 'Winsock não está respondendo."';
    Exit;
  end;
  IPaddr := '';
//  New(HName);
  if GetHostName(@HName, SizeOf(HName)) = 0 then
  begin
//    HostName := StrPas(HName);
    HEnt := GetHostByName(@HName);
    HostName := HEnt^.h_Name;
    for i := 0 to HEnt^.h_length - 1 do
      IPaddr := Concat(IPaddr, IntToStr(Ord(HEnt^.h_addr_list^[i])) + '.');
    SetLength(IPaddr, Length(IPaddr) - 1);
    Result := True;
  end
  else begin
  case WSAGetLastError of
    WSANOTINITIALISED:WSAError:='WSANotInitialised';
    WSAENETDOWN      :WSAError:='WSAENetDown';
    WSAEINPROGRESS   :WSAError:='WSAEInProgress';
  end;
  end;
//  Dispose(HName);
  WSACleanup;
end;

procedure CopyToClipboard(sStr: String);
begin
  with TClipboard.Create do
    AsText := Trim(sStr);
end;

function ExtractFilePath_MM(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter(iif(Pos('/',FileName)>0,'/','\') + ':', FileName);
  Result := Copy(FileName, 1, I);
end;

function ExtractFileName_MM(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter(iif(Pos('/',FileName)>0,'/','\') + ':', FileName);
  Result := Copy(FileName, I + 1, MaxInt);
end;


function iif(bCondicao: Boolean; vValorTrue, vValorFalse: Variant): Variant;
begin
  case bCondicao of
    True: Result:=vValorTrue;
    False:Result:=vValorFalse
  end;
end;


procedure GravaConf(Secao, Chave, Valor: String; bWorkingDirectory: Boolean=False);
var
  RegIni: TIniFile;
begin
  try
    if bWorkingDirectory then
      RegIni := TIniFile.Create(IncludeTrailingPathDelimiter(SysUtils.GetCurrentDir)+ChangeFileExt(ExtractFileName(ParamStr(0)), '.conf'))  // iniciar em
    else
      RegIni := TIniFile.Create(ChangeFileExt(Application.ExeName ,'.conf'));
    try
      RegIni.WriteString(Secao, Chave, Valor);
    finally
      RegIni.Free;
    end;
  except
    //erro
  end;
end;

function LerConf(sSecao, sChave: String; sPadrao: String=''; bWorkingDirectory: Boolean=False): String;
var
  RegIni: TIniFile;
begin
  try
    if bWorkingDirectory then
      RegIni := TIniFile.Create(IncludeTrailingPathDelimiter(SysUtils.GetCurrentDir)+ChangeFileExt(ExtractFileName(ParamStr(0)), '.conf'))
    else
      RegIni := TIniFile.Create(ChangeFileExt(Application.ExeName ,'.conf'));
    try
      Result := RegIni.ReadString(sSecao, sChave, sPadrao);
    finally
      RegIni.Free;
    end;
  except
    Result := sPadrao;
  end;
end;


function getVersaoArquivo(FileName: PChar): TSisVersion;
var
  ZeroValue, InfoSize: DWORD;
  Buffer: Pointer;
  Version: Pointer;
  InfoLen: UINT;
  Major, Minor, Release, Build: Integer;
begin
  if FileExists(FileName) then
  begin
    InfoSize := GetFileVersionInfoSize(FileName, ZeroValue);
    GetMem(Buffer, InfoSize);
    try
      try
        GetFileVersionInfo(FileName, 0, InfoSize, Buffer);
        VerQueryValue(Buffer, '\', Version, InfoLen);
        Major   := PVSFixedFileInfo(Version)^.dwFileVersionMS shr 16;
        Minor   := PVSFixedFileInfo(Version)^.dwFileVersionMS shl 16;
        Minor   := Minor shr 16;
        Release := PVSFixedFileInfo(Version)^.dwFileVersionLS shr 16;
        Build   := PVSFixedFileInfo(Version)^.dwFileVersionLS shl 16;
        Build   := Build shr 16;

    //    Result.Descricao   := qry.FieldByName('descricao').AsString;
    //    Result.Sigla       := qry.FieldByName('sigla').AsString;
        Result.MajorVersion:= Major;
        Result.MinorVersion:= Minor;
        Result.Release     := Release;
        Result.Build       := Build;
      except
        Result.MajorVersion:= 0;
        Result.MinorVersion:= 0;
        Result.Release     := 0;
        Result.Build       := 0;
      end;
    finally
      FreeMem(Buffer, InfoSize);
    end;
  end
  else
  begin
    Result.MajorVersion:= 0;
    Result.MinorVersion:= 0;
    Result.Release     := 0;
    Result.Build       := 0;
  end;
end;


function RemoverEnter(sStr: String): String;
var
  i: Integer;
begin
  Result := StringReplace(sStr, Chr(13), ' ', [rfReplaceAll]);
  Result := StringReplace(Result, Chr(10), ' ', [rfReplaceAll]);
  Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;

function RetornaSoNumero(sStr: String): String;
var
  i: Integer;
begin
//  Result := '';
//  for i := 1 to Length(sStr) do
//    if (sStr[i] in ['0','1','2','3','4','5','6','7','8','9']) then
//      Result := Result + sStr[i];

  Result := '' ;
  for i := 1 to Length(sStr) do
  begin
    {$IFDEF DELPHI12_UP}
    if CharInSet(sStr[i], ['0'..'9']) then
    {$ELSE}
    if (sStr[i] in ['0'..'9']) then
    {$ENDIF}
       Result := Result + sStr[i];
  end ;

end;

{ Left Trim }
function LTrim(ctext: string) : string;
var
  nloop : integer ;
  ctemp : string ;
begin
  ctemp  := '' ;
  if (length(ctext) > 0) then
    begin
      for nloop := 1 to length(ctext) do begin
        if (ord(ctext[nloop]) <> 32) then
          begin
            ctemp := ctemp + copy(ctext,nloop,length(ctext)) ;
            break ;
          end;
      end;
    end;
  result := ctemp ;
end;

{ Right Trim }
function RTrim(ctext: string) : string;
var
  nloop : integer ;
begin
  if (length(ctext) > 0) then
    begin
      for nloop := length(ctext) downto 1 do begin
        if (not (ord(ctext[nloop]) in [0,1,2,32])) then
          begin
            ctext := copy(ctext,1,nloop) ;
            break ;
          end;
      end;
    end;
  result := ctext ;
end;

function AllTrimTexto(sTexto: string): string;
begin
  while (Copy(sTexto, 1, 1) = ' ')
    and (Length(sTexto) > 0) do
    Delete(sTexto, 1, 1);

  while (Copy(sTexto, Length(sTexto), 1) = ' ')
    and (Length(sTexto) > 0) do
    Delete(sTexto, Length(sTexto), 1);

  AllTrimTexto := sTexto;
end;

function allTrim(const ctext: String; bIn: Boolean=False) : string;
begin
  if bIn then
    Result := StringReplace(ctext, ' ', '', [rfReplaceAll])
  else
    Result := lTrim(rTrim(ctext));
end;

function leftString(ctext : string; nlen: integer) : string ;
begin
  result := copy(ctext,1,nlen)
end;

function rightString(ctext : string; nlen: integer) : string ;
var
  ntemp : word ;
begin
  result := ctext ;
  if (length(ctext) > nlen)  then
    begin
      ntemp := (length(ctext)-nlen)+1 ;
      result := copy(ctext,ntemp,nlen) ;
    end;
end;

function RetornarLetraNumero(sStr: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(sStr) do
  begin
    if (sStr[i] in ['0'..'9']) or
       (sStr[i] in ['a'..'Z']) or
       (sStr[i] in ['A'..'Z']) or
       (sStr[i] in ['ã','á','à','é','í','õ','ó','ú','ç']) or
       (sStr[i] in ['Ã','Á','À','É','Í','Õ','Ó','Ú','Ç']) then
      Result := Result + sStr[i];
  end;
end;

function RetornaSoLetraNumero(Str: String; bFlagAllTrim: Boolean=False; bUppercase:Boolean=True): String;
begin
  // Com essa conversão para UTF8 talvez não precise fazer os outros StringReplace abaixo
  Result:=UTF8Encode(Str);
  Result:=StringReplace(Result,'.',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,',',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'-',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'/',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'"',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'&','E',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'''',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,':',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'§',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'|',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'*',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'+',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'@',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'#',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'$',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'%',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'=',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'>',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'<',' ',[rfReplaceAll]);
  Result:=StringReplace(Result,'   ',' ',[rfReplaceAll,rfIgnoreCase]);
  Result:=StringReplace(Result,'  ',' ',[rfReplaceAll,rfIgnoreCase]);
  if bFlagAllTrim then
    Result:=AllTrim(Result, True);
  Result:=iif(bUppercase, UpperCase(AnsiToAscii(Result)), AnsiToAscii(Result));
end;

function CriarDiretorio(Dir: String): String;
begin
  if not DirectoryExists(Dir) then
    ForceDirectories(Dir);
end;

procedure CopyFile(const sourcefilename, targetfilename: String);
{Copia um arquivo de um lugar para outro}
var
  S, T: TFileStream;
Begin
  S := TFileStream.Create( sourcefilename, fmOpenRead );
  try
    T := TFileStream.Create( targetfilename, fmOpenWrite or fmCreate );
    try
      T.CopyFrom(S, S.Size ) ;
    finally
      T.Free;
    end;
  finally
    S.Free;
  end;
end;

procedure DeleteArq(Arquivo: String);
begin
  if SysUtils.FileExists(Arquivo) then
    SysUtils.DeleteFile(Arquivo);
end;

procedure DeleteArq(slArquivos: TStringList);
var
  i: Integer;
begin
  if Assigned(slArquivos) then
    for i := 0 to slArquivos.Count - 1 do
      if FileExists(slArquivos[i]) then
        SysUtils.DeleteFile(slArquivos[i]);
end;

function AnsiToAscii(str:string): string;
var
  i:integer;
begin
  for i:=1 to Length(str) do
  begin
    case str[i] of
      'á': str[i] := 'a';
      'é': str[i] := 'e';
      'í': str[i] := 'i';
      'ó': str[i] := 'o';
      'ú': str[i] := 'u';
      'à': str[i] := 'a';
      'è': str[i] := 'e';
      'ì': str[i] := 'i';
      'ò': str[i] := 'o';
      'ù': str[i] := 'u';
      'â': str[i] := 'a';
      'ê': str[i] := 'e';
      'î': str[i] := 'i';
      'ô': str[i] := 'o';
      'û': str[i] := 'u';
      'ä': str[i] := 'a';
      'ë': str[i] := 'e';
      'ï': str[i] := 'i';
      'ö': str[i] := 'o';
      'ü': str[i] := 'u';
      'ã': str[i] := 'a';
      'õ': str[i] := 'o';
      'ñ': str[i] := 'n';
      'ç': str[i] := 'c';
      'Á': str[i] := 'A';
      'É': str[i] := 'E';
      'Í': str[i] := 'I';
      'Ó': str[i] := 'O';
      'Ú': str[i] := 'U';
      'À': str[i] := 'A';
      'È': str[i] := 'E';
      'Ì': str[i] := 'I';
      'Ò': str[i] := 'O';
      'Ù': str[i] := 'U';
      'Â': str[i] := 'A';
      'Ê': str[i] := 'E';
      'Î': str[i] := 'I';
      'Ô': str[i] := 'O';
      'Û': str[i] := 'U';
      'Ä': str[i] := 'A';
      'Ë': str[i] := 'E';
      'Ï': str[i] := 'I';
      'Ö': str[i] := 'O';
      'Ü': str[i] := 'U';
      'Ã': str[i] := 'A';
      'Õ': str[i] := 'O';
      'Ñ': str[i] := 'N';
      'Ç': str[i] := 'C';
      'º': str[i] := 'o';
      '&': str[i] := 'e';
    end;
  end;
  Result:=str;
end;

function StrZero(Variavel:string; QUANT:integer; Onde:Char):String;
var
  I,Tamanho:integer;
begin
  Tamanho:=length(Variavel);
  if quant > tamanho then
  begin
    for I:=1 to quant-tamanho do
    begin
      if onde='E' then
        Variavel:='0'+Variavel
      else
        Variavel:=Variavel+'0';
    end;
  end;
  StrZero:=Variavel;
end;

end.
