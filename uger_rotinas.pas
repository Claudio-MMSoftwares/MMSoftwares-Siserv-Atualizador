(* versões do delphi para ifdef:
{$IFDEF VER80}  - Delphi 1
{$IFDEF VER90}  - Delphi 2
{$IFDEF VER100} - Delphi 3
{$IFDEF VER120} - Delphi 4
{$IFDEF VER130} - Delphi 5
{$IFDEF VER140} - Delphi 6
{$IFDEF VER150} - Delphi 7
{$IFDEF VER160} - Delphi 8
{$IFDEF VER170} - Delphi 2005
{$IFDEF VER180} - Delphi 2006
{$IFDEF VER180} - Delphi 2007
{$IFDEF VER185} - Delphi 2007
{$IFDEF VER200} - Delphi 2009
{$IFDEF VER210} - Delphi 2010
{$IFDEF VER220} - Delphi XE
{$IFDEF VER230} - Delphi XE2
{$IFDEF VER240} - Delphi XE3
{$IFDEF VER250} - Delphi XE4
{$IFDEF VER260} - Delphi XE5
{$IFDEF VER270} - Delphi XE6


outro exemplo:
{$IF CompilerVersion > 18.5}
   //Delphi 2009 or higher
   //Unicode version of code
{$ELSE}
   //Delphi 2007 and earlier
   //NON-Unicode version of code
{$IFEND}
{$IF CompilerVersion >= 18.5}
//some code only compiled for Delphi 2007 and later
{$IFEND}
Delphi XE2  - 23
Delphi XE   - 22
Delphi 2010 - 21
Delphi 2009 - 20
Delphi 2007 - 18.5
Delphi 2006 - 18
Delphi 2005 - 17
Delphi 8    - 16
Delphi 7    - 15
Delphi 6    - 14
*)


{
  Dicas:
   - SysUtils.GetCurrentDir  -> "Iniciar em" ou "Working Directory" do atalho;
}
unit uger_rotinas;

interface

//{$I RGF.inc}

uses
  Windows, Forms, ComCtrls, DBTables, Classes, DateUtils,
//  {$ifdef VER210}
  DbAGrids, QRPDFFilt, QuickRpt, QRCtrls, //uRGFSearch, uRGFProgressBar,
//  {$endif}
  DB, SqlExpr, StdCtrls,
  IBC, Menus, XMLDoc,
  Controls, ExtCtrls, SysUtils, SHDocVw, //uutil_types,
  IdHTTP, DBGrids, Graphics, GraphUtil, TlHelp32, IdGlobal;

type
  THackDBAltGrid = class(TDBAltGrid);
  THackDBGrid = class(TDBGrid);

  TInicializaProgressBar = procedure(const iMax: Integer) of object;
  TStepIt = procedure of object;
  TStepItFmt = procedure(const sFormat: String) of object;
  TProcedure1 = procedure(sAux: String) of object;
  TFunction0 = function: String of object;
  TFunction1 = function(AURL: String): String of object;
  TFunctionPost = function(AURL: String; ASource: String): String of object;
  TProcedureGet = procedure(AURL: string; AResponseContent: TStream);

  TTypeRound = (trDef,trUp);

  TVect10 = Array[0..9] of String;
  TVect10Variant = Array[0..9] of Variant;

  TConfigur = class
  private
    FCriado            : Boolean;
    FConfigurRefresh   : TDateTime;
    FConfigurEspRefresh: TDateTime;
    FConfigEmpresa     : TDateTime;
    FConnec            : TIBCConnection;
    FqrConfigur        : TIBCQuery;
    FqrConfigurEsp     : TIBCQuery;
    FqrConfigEmpresa   : TIBCQuery;
    FNoEmpresa         : Integer;
    procedure setConnec(const Value: TIBCConnection);
    procedure setEmpresa(const Value: Integer);
  protected
    property Criado: Boolean read FCriado write FCriado default False;
  published
  public
    constructor Create(oConnec: TIBCConnection = nil);
    destructor Destroy;

    property Connec   : TIBCConnection read FConnec    write setConnec;
    property NoEmpresa: Integer        read FNoEmpresa write setEmpresa;

    function getValue(sField: string; vDefaultValue: Variant; bConfigurEsp: Boolean): Variant;
    function getConfigEmpresa(sChave, sParametro: string; sDefaultValue: String): String;
    class function AddConfigEmpresa(oConnec: TIBCConnection; iNoEmpresa: Integer; sChave, sParametro, sValor: String): Boolean;
    procedure RefreshDados;
  end;

  function iif(bCondicao: Boolean; vValorTrue, vValorFalse: Variant): Variant; overload;

  function LTrim(ctext: string) : string;
  function RTrim(ctext: string) : string;
  function allTrim(const ctext: string; bIn: Boolean=False): string;
  function rightString(ctext : string; nlen: integer) : string ;
  function leftString(ctext : string; nlen: integer) : string ;
  function PadC(const S: string; const Len: integer; const Ch: Char): string;
  function PadL(ctext : string; nlen:integer ; cchar : char) : string ;
  function PadR(ctext : string; nlen:integer ; cchar : char) : string ;
  {:@Summary Desabilitar o Botão de Fechar do Form (X)
  @Autor Klaus Ezequiel Novello - Outubro/2010
  @Desc Adicionar no FormCreate para desabilitar o botão de bechar (X)
  @Param Sender: TSysForm
  @example DesabilitarFecharBorderIcon(Sender);}
  procedure DesabilitarFecharBorderIcon(Sender: TObject);
  {:@Summary Exportar DBGrid para XLS
  @Autor Klaus Ezequiel Novello - Dezembro/2010
  @Desc Exportar DBGrid para XLS
  @Param Grid: TDbAltGrid;
  @Param ExcelFile: String;
  @Param TotalRegistros : Integer
  @example GridToExcelFile(Grid, 'c:\teste.xls', TotalRegistros); }
//  {$ifdef VER210}
//  function GridToExcelFile(Grid: TDbAltGrid;ExcelFile: String; TotalRegistros : Integer; pbAux: TObject; var bCancelarExportXLS: Boolean):Boolean;
//  {$ENDIF}
  {:@Summary Valida e obtem a versão do sistema
  @Autor Klaus Ezequiel Novello - Janeiro/2011
  @Desc Valida e obtem a versão do sistema
  @Param dbConnec: TDatabase
  @Param iIDModulo: Integer
  @Result TRGFVersion
  @example Versao := getVersaoModulo(MaindataModule.Database, 1); }
//  function getVersaoModulo(dbConnec: TDatabase; iIDModulo: Integer; FileName: PChar): TRGFVersion; overload;
//  function getVersaoModulo(Connec: TIBCConnection; iIDModulo: Integer; FileName: PChar): TRGFVersion; overload;

//  function getVersaoArquivo(FileName: PChar): TRGFVersion;

  {:@Summary Retorna Somente os Número de Uma String.
  @Author Klaus Ezequiel Novello - Janeiro/2011
  @param Texto = String a Ser verificada.
  @return Retorna uma string somente com os Números do parametro
  @Cat Manipulação de Strings}
  function RemoverEnter(sStr: String): String;

  function RetornaSoNumero(sStr: String): string;

  function RetornaSoLetraNumero(Str:String; bFlagAllTrim: Boolean=False; bUppercase:Boolean=True):String;
  function RetornarLetraNumero(sStr: String): String;

  {:@Summary Converte String para TStringList
  @Autor Klaus Ezequiel Novello
  @Desc Converte String para TStringList
  @Param Delimitador Informe o Delimitador (não obrigatório, neste caso será ; )
  @Param Origem Informe a String para converter
  @Return Retorna TStringList formatada
  @Cat Converões de Tipos}
  function StringToList(Origem:String;Delimitador:String=''): TStringList;
  {:@Summary Converte TStringList para String
  @Autor Klaus Ezequiel Novello
  @Desc Converte TStringList para String
  @Param Delimitador Informe o Delimitador (não obrigatório, neste caso será ; )
  @Param Lista Informe a lista do Tipo TStringList
  @Return Retorna String formatada
  @Cat Converões de Tipos}
  function ListToString(Lista:TStringList;Delimitador:String=''):String;
  {:@Summary Converte Componente (DFM) para String
  @Autor Klaus Ezequiel Novello
  @Desc Converte Componente (DFM) para String
  @Param Component ; )
  @Return Retorna String do DFM }
  function ComponentToString(Component: TComponent): string;
  function StringComponentToBinaryStream(sStr: String): TStream;
  {:@Summary Valida Path
  @Desc Valida o Path sempre inserindo o Delimitador no final
  @Param sPath Informe o caminho
  @Param sDelim Informe o delimitador (caso não informado pega PathDelim)
  @Cat Manipulação de Arquivos}
  function ValidaPath(sPath: String; sDelim: String=''):String;
//  {$IFDEF VER210} // Delphi 2010
//  function Encrypt(sStr: String; sKey: String=''): String;
//  function Decrypt(sStr: String; sKey: String=''): String;
//  {$ENDIF}
  function Criptografia(mStr, mChave: string): string;
  function EncryptSTR(Const InString:String; StartKey,MultKey,AddKey:Integer): String;
  function DecryptSTR(Const InString: String; StartKey, MultKey, AddKey: Integer): String;
  function getMD5HashString(value: string): string;
  {:@Summary Captura todos os arquivos de um diretório
  @Desc Captura todos os arquivos de um diretório
  @Param ASource Informe o Caminho do diretório que proverá a lista
  @Param lstDir Informe True para capturar também os diretórios
  @Param ADirList Informe um objeto tipo TStringList que receberá a lista
  @Cat Manipulação de Arquivos}
  procedure GetFileList(ASource : string; lstDir:Boolean; var lstFiles : TStringList;bIncrement:Boolean=false);
  {:@Summary Lista arquivos do diretório Recursivo (sub diretórios)
  @Desc Lista arquivos do diretório, com opção de recursividade (sub diretórios) e deleção
  @Param sDirName Informe o Caminho do diretório que proverá a lista
  @Param sFilter  Informe o filtro para os arquivos a serem listados
  @Param bRecursive Informe true para listar arquivos e diretórios recursivamente
  @Param bDelete Informe true para deletar os arquivos listados.
  @Param lstOut Informe um objeto tipo TStringList que receberá a lista com os arquivos diretórios
  @Cat Manipulação de Arquivos}
  procedure GetFilesFrom(sDirName, sFilter: String; bRecursive:Boolean; bDelete:Boolean=False; lstOut:TStringList=nil);
  {:@Summary Arredondamento de Valores
  @param AValue = Valor a ser arredondado.
  @param ADecimal = Número de casas decimais.
  @return Retorna Valor Convertido
  @Cat Manipulação de Pontos Flutuantes}
  function RoundRGF(AValue: Variant; Const ADecimal: integer; ATypeRound:TTypeRound=trDef): Double;
  function Arredonda(Connec: TIBCConnection; Valor: Double): Double;
  function Arredonda2(Valor: Double): Double;

  function DigitoABAC_EAN(S: String): Char;
  function DigitoM9(S: String): Char;
  function DigitoM11(S: String): Char;

  function ValidaEAN(CodBarras: String): String;

  function AdicionaMes(data: TDateTime; Valor: Integer): TDateTime;
  {:@Summary Adiciona Ano em uma data.
  @param Data = Data .
  @param Valor= Número de Anos.
  @return Nova Data
  @Cat Manipulação e Validação de Datas}
  function AdicionaAno(data: TDateTime; Valor: Integer): TDateTime;
  {:@Summary Subtrai Dia de uma data.
  @param Data = Data .
  @param Valor= Número de Dias.
  @return Nova Data
  @Cat Manipulação e Validação de Datas}
  function SubtraiMes(data: TDateTime; Valor: Integer): TDateTime;
  {:@Summary Subtrai Ano de uma data.
  @param Data = Data .
  @param Valor= Número de Anos.
  @return Nova Data
  @Cat Manipulação e Validação de Datas}
  function SubtraiAno(data: TDateTime; Valor: Integer): TDateTime;
  function RetornaPrimeiroDiaMes(Data : TDateTime) : TDateTime;
  function RetornaUltimoDiaMes(Data : TDateTime) : TDateTime;
  {:@Summary Acha o último dia do mês de acordo com o parâmetro passado
  @param Data = Data a qual se deseja saber qual é o último dia do mês.
  @return Retorna o último dia do mês
  @Cat Manipulação e Validação de Datas}
  function RetornaUltimoDia(Data : TDateTime) : Integer;
  {:@Summary Verifica se o ano é bissexto.
  @param Data = Data para verificação se o ano da mesma é bissexto
  @return se é bissexto retorna True
  @Cat Manipulação e Validação de Datas}
  function AnoBisexto(Data : TDateTime) : Boolean;
  function RetornarMesExtenso(iMes: Integer): String;
  function ExtractFilePath_RGF(const FileName: string): string;
  function ExtractFileName_RGF(const FileName: string): string;
  {Verifica se arquivo está em uso}
  function IsFileInUse(FileName: TFileName): Boolean;
  function GetFileDate(Arquivo: String): TDateTime;
  procedure SepararArquivosPorData(sPath, sFiltro: String; sMascara: String='yyyy-mm-dd');

  //=================== Arquivos Ini =================
  procedure GravaConf(Secao, Chave, Valor: string; bWorkingDirectory: Boolean=False);
  function LerConf(sSecao,sChave: String; sPadrao: String=''; bWorkingDirectory: Boolean=False): String;
  //==================================================
  function AnsiToAscii(str: String): String;
  function StrZero(Variavel: String; QUANT:integer; Onde:Char): String;

  function CriarDiretorio(Dir: String): String;
  procedure CopyFile(const sourcefilename, targetfilename: String);
  procedure DeleteArq(Arquivo: String); overload;
  procedure DeleteArq(slArquivos: TStringList); overload;

  // Exemplo: GravarRegistro(HKEY_CURRENT_USER, 'Software\Tantra Software', 'Username', 'KLAUS')
  procedure GravarRegistro(RootKey: HKEY; sChave, sStr, sValor: String);
  // Exemplo: LerRegistro(HKEY_CURRENT_USER, 'Software\Tantra Software', 'Username', UsuarioWindows)
  function LerRegistro(RootKey: HKEY; sChave, sStr, sValorPadrao: String): String;


//  function PutINIDB(Connec: TFDConnection; Chave, Parametro: String; Valor: Variant; User: String): Boolean; overload;
  function DelINIDB(Connec: TIBCConnection; Chave, Parametro: String; User: String): Boolean; overload;
  function getDatabase(Connec: TIBCConnection; bBancoAtualizacao: Boolean=False): string;
  function getDatabaseBDE(sAliasName: String; bBancoAtualizacao: Boolean): String;
  function getDatabaseAtualizacao(sDatabase: String): String;
  function getServer_Database(sDatabase: String): Variant;
  function getConfigur(Connec: TIBCConnection; sCampo: String; vValorPadrao: Variant; bConfigEspecial: Boolean): Variant;
  function getConfiguracaoEmpresa(Connec: TIBCConnection; iEmpresa: Integer; sChave, sParametro: String; sValorPadrao: Variant): Variant;
  function putConfiguracaoEmpresa(Connec: TIBCConnection; iEmpresa: Integer; sChave, sParametro: String; sValorPadrao: Variant): Variant;

  function BlobSaveToStream(CampoBlob: TBlobField; Stream: TStream; Count: Int64; Barra: TProgressBar): Boolean;
  function BlobLoadFromStream(CampoBlob : TBlobField; Stream: TStream; Count: Int64; Barra : TProgressBar): Boolean;

  procedure AbreForm(aClasseForm: TComponentClass; aForm: TForm);

  function ValidaEmail(sEmail: String): Boolean;
  function ValidarEmails(sEmail: String; bMostrarAlerta: Boolean=True; bFixSeparator: Boolean=False): Boolean;
  function EmailsValidos(sEmail: String): String;

  function RunDosAndReturn(DosApp: String): String;
  procedure RunDosInMemo(DosApp: String; AMemo: TMemo);
  //===================== Sessão Windows =====================
  function IsRemoteSession: Boolean;
  function GetCurrentSessionID: Integer;
  //==========================================================
  function GetComputerNameFunc: string;
  procedure FTP_Put(sHost, sUser, sPassword, sPathOrigem, sPathDestino: String; slArquivos: TStringList);
  function FTP_PutFile(sHost, sUser, sPassword, sArquivoOrigem, sPathFileDestino: String): Boolean;
  function FTP_DeleteFile(sHost, sUser, sPassword, sArquivo: String): Boolean;

  function UpperCaseRGF(sStr: String; bSemAcento: Boolean=False): String;
  function LowerCaseRGF(sStr: String; bSemAcento: Boolean=False): String;

  {:@Summary Preeche uma string com caracteres ao lado esquerdo ou direito
  @Desc Preeche uma string com caracteres ao lado esquerdo ou direito, pode-se
  determinar o limite da string de saída e também o caracter de preechimento.
  @Param Palavra Informe a String que inicial para preencher
  @Param Qtd Informar o tamanho da String de saída
  @Param Caracter Informe o caracter de preenchimento
  @Param D_ou_E Informe 'E' para Preecher a esquerda e 'D' para preenhcer a direita
  @Return Retorna a string formatada.
  @Cat Manipulação de Strings}
  function InsereStr (sTexto: String; iTam: Integer; sChar: Char; sD_ou_E_ou_C: Char): string;

  {:@Summary Centraliza uma String no tamanho informado.
  @Example CentralizaStr('XXXXXX',20) = '       XXXXXX       '
  @param sTexto = String que será centralizada.
  @param iTam = Tamanho em que a string será centralizada.
  @param sChar = Caracter de preechimento da string.
  @return Retorna uma string no tamanho informado contendo a string inicial
  centralizada preenchida com o caracter informado.
  @Cat Manipulação de Strings}
  function CentralizaStr(sTexto:string;iTam:integer;sChar:char): string;

  //FORMATAÇÃO DE MÁSCARAS
  function Formata(const Texto: string;Tipo:String): string;

  {:@Summary Formata valor com decimais
  @Desc Formata valor com decimais
  @Param Mascara Informe a mascara Ex.'###,##0.00
  @Param Vlr Informar o valor a formatar
  @Param CtrlZerado Informar True caso queira mostrar ' ' caso Vlr seja = 0
  @Param CtrZeradoMantemTamanho Informar True para manter ' ' mas com o tamanho da Mascara caso Vlr seja = 0
  @Return Retorna a string formatada.
  @Cat Manipulação de Strings}
  function MascaraVlr(Mascara:string;Vlr:double;CtrlZerado:Boolean=false;CtrZeradoMantemTamanho:Boolean=false):string;

  // Ocultar / Mostrar Barra de Títulos
  procedure ShowTitlebar(AForm: TForm; AShow: boolean);

  // Buscar informações da empresa
  function getEmpresaInfo(Connec:TDatabase; iEmpresa: Integer; sCampo: String): Variant;  Overload;
  function getEmpresaInfo(Connec:TIBCConnection; iEmpresa: Integer; sCampo: String): Variant;  OverLoad;

  function getIP: String;
  function RETORNA_VLR(Connec: TIBCConnection; tbSelect: String; cmpCriterio,cmpValores,cmpRetorno: Array of Variant): TVect10Variant;

  function DownloadFile(Source, Dest: string): Boolean;

  // Se hora for menor que informada, pega dia anterior desconsiderando sábado e domingo
  function getDateAfterHour(Data: TDateTime; HoraReducao: String='0:00:00'): TDateTime;

  // Copiar para Clipboard
  procedure CopyToClipboard(sStr: String);

  function SelectDir(hOwner: THandle; const Caption, InitialDir: string;
                     const Root: WideString; ShowStatus: Boolean; out Directory: string): Boolean;

  // Exemplo de uso:
  //  ExecAndWait('c:\windows\calc.exe','',sw_show);
  function ExecAndWait(const FileName, Params: String; const WindowState: Word): Boolean;
  function ExecutarPrograma(const FileName: String; WindowState: Word): Boolean; overload;
  function ExecutarPrograma(const sFileName: String; sFormName: String=''): Boolean; overload;

  procedure ValidarMenu(oMenu: TMainMenu);


  function getDropboxFolder: String;
  function getAppDataPath: string;
  function getIniciarEm: string;

  // Encode Arquivo Base64
  function EncodeFileBase64(const FileName: string): AnsiString;
  // Decode Arquivo Base64
  procedure DecodeFileBase64(const EncodedBase64: AnsiString; const FileName: string);

  function UpperLowerCase(x: String): String;
  // Retorna Top e Left relativo ao form - considera as bordas (Posição)
  function GetTopLeft(AForm: TForm; AControl: TControl): TPoint;
  // Retorna Top e Left relativo ao form - não considera as bordas (mais indicado)
  function RetornaTopLeft(Controle: TControl): TPoint; overload;
  function RetornaTopLeft(Controle: TControl; MainFormName: String): TPoint; overload;
  // Retorna todos os campos do registro selecionado num vetor
//  function getFieldsValue(qrLocal: TIBCQuery): TArecField;

  function GetMyDocuments: String;
  procedure SaveImageListToResFile(sFile: String; ImgList: TImageList);
  procedure ReadImageListFromResFile(sFile: String; const ImgList: TImageList);
  procedure GetImageFromResFile(sFile: String; Idx: Integer; const Img: TImage);
//  procedure CopiarParaClipBoard(sStr: String; Grid: TDBAltGrid=nil);

  // Copiar de TMainMenu para Treeview
  // Exemplo de uso:   MenuToTreeView(MainMenu1, TreeView1);
  // Colocar no OnClick do TreeView para executar o menu: TMenuItem(TreeView1.Selected.Data).Click;
  procedure MenuToTreeView(AMenu: TMainMenu; ATree: TTreeView);

  procedure AbrirHtmlWebBrowser(WebBrowser: TWebBrowser; slHTML: TStringList);
  procedure LoadStreamWebBrowser(WebBrowser: TWebBrowser; Stream: TStream);
  procedure LoadHtmlWebBrowser(WebBrowser: TWebBrowser; sHTML: String);
  function SaveHtmlWebBrowser(WebBrowser: TWebBrowser): String;
  function getStreamFromString(sStr: AnsiString): TStream;

  // Adicionar imagem numa célula do DBGrid a partir de um ImageList
  // Exemplo de uso:
  //  - no evento OnDrawColmnCell adicionar:
  //      var fixRect: TRect;
  //      {...}
  //      fixRect := Rect;
  //      fixRect := AddImageToGrid(Rect, PedidoGrid, ImageList1, 0);
  //      DefaultDrawColumnCell(fixRect, DataCol, Column, State);

//  {$IFDEF VER210} // Delphi 2010
  function AddImageToGrid(const Rect: TRect; Grid: TDBAltGrid; imgList: TImageList; imgIndex: Integer; bTransparent: Boolean=True; cTransparentColor: TColor=clNone): TRect; overload;
//  {$ENDIF}
  function AddImageToGrid(const Rect: TRect; Grid: TDBGrid; imgList: TImageList; imgIndex: Integer; bTransparent: Boolean=True; iX_Transp: Integer=0; iY_Transp: Integer=0; bShowText: Boolean=False): TRect; overload;
  procedure ImageFromListImg(Img: TImage; imgList: TImageList; imgIndex: Integer);

  procedure CapturaTela(sFileName: String; iQualidade: Integer=100);
  function getParentStructure(const Component: TWinControl): String;
  function getParentRGF(const Component: TWinControl): TWinControl; overload;
  function getParentRGF(const Component: TWinControl; sClassname: String): TWinControl; overload;
  procedure FormSempreVisivel(Form: TForm);
  // Retorno o usuário logado no windows
  function UsuarioWindows: String;
  // Retorna IP e nome da máquina
  function GetIPFromHost(var HostName, IPaddr, WSAError: String): Boolean;

//   Retorna endereço a partir do CEP
//  function getDadosEndereco(Connec: TIBCConnection; sUsuario, sCep: String; var Endereco, CodBairro, Bairro, CodCidade, Cidade, Estado: String): Boolean; overload;
//  {$IFDEF VER210} // Delphi 2010
//  procedure LimparCampos(aControl: TWinControl);
//  {$ENDIF}

  function GeraLOGArquivo(sFileName, sLog: String): String;
  procedure AddLogMemo(sMsg: String; MemoLog: TMemo);
  procedure TextToFile(sText, sFileName: String);

//  function GetComponentValue(wcComponent: TWinControl): String;
  function PathWithDelim( const APath: String ): String;
  function PathWithoutDelim( const APath: String ): String;

  function MoveCopiaDiretorios(pOperacao: Integer; pOrigem, pDestino: string): Boolean;

  // Exportar um relatório em Quick Report para PDF
//  {$IFDEF VER210} // Delphi 2010
  function ExportQRToPDF(Report: TQuickRep; sFileName: String; bOnlyPrepare: Boolean=False): Boolean;
//  {$ENDIF}
  function DiasAtraso(Connec: TIBCConnection; prDtIni, prDtFim: TDateTime): Integer;
  function isFeriado(Connec: TIBCConnection; prData: TDateTime): Boolean;
  function AumentaXDiasUteis(Connec: TIBCConnection; DataIni: TDateTime; Dias: Integer):TDateTime;

  function BuscaLiberacao(Connec: TIBCConnection; CNPJ, Opcao: String):Boolean;
  function bLocalX(Connec: TIBCConnection; CNPJ: String):Boolean;
  function QtdeInteira(Qtde: Double; Conversao: Double; TipoConv: String): String;
  function CopyAtPos(S: String; Initial, Final: Integer): String;
  function DiasUteis(Connec: TIBCConnection; prDtIni, prDtFim: TDateTime): Integer;
  function DiminuiXDiasUteis(Connec: TIBCConnection; DataIni: TDateTime; Dias: Integer):TDateTime;
  function isItemOfKit(Connec: TIBCConnection; prCodReduzido: String): Boolean;
  function getDados(Connec: TIBCConnection; prTabela, prCampoChave, prValorChave, prCampoRetorno: String): String;
  procedure getValoresDolar(var CompraPtax, CompraComercial, CompraTurismo, CompraParalelo, VendaPtax, VendaComercial,
                            VendaTurismo, VendaParalelo, Data : String);
  function getGarantiaFinal(dDataEmissao:TDateTime;sTempoGarantia:String):TDateTime;
  Procedure getQuantidadeAnos(iTempo:Integer; Var iMeses, iAnos:Integer);
  Function isAnoBissexto(iAno:Integer):Boolean;
  Function getFirstDay(iMes,iAno:Integer):TDate;
  Function getLastDay(iMes,iAno:Integer):TDate;
  function DiaUtil(dData: TDateTime): Boolean;
  function ExtractFileNameWithoutExt(sFileName: String): String;
  function ProximoDiaUtil(Connec: TIBCConnection; DataIni: TDateTime; bAntecipar: Boolean=False):TDateTime;
  function MontaAspasComVirgula(Entrada : String): String;
  function GetCheckSum(FileName: string): DWORD;
  function TirarZerosClassificacao(Classificacao: string):String;
  function getParamApp(sParam: string): string;

//  {$IFDEF VER210} // Delphi 2010
//  function Compactar(sDestZipFile: String; slFiles: TStringList): String;
//  function Descompactar(sZipFile, sPathDest: String; bDeleteZipFileAfter: Boolean=False): String;
//  {$ENDIF}
  function FileDateTime(const FileName: string; sTipo: string='Modified'): TDateTime;
  function getExcelColumnName(columnNumber: Integer): String;
  function getExcelColumnIndex(columnName: String): Integer;
  function NomeClasse(const Janela: HWND): string;
  function WinVersion: string;
  function EnumWindowsProc(Wnd: HWND; lb: TStringList): BOOL; stdcall;
  function getMonitorResolution(oForm: TForm; bWorkArea: Boolean): TPoint;
  function DateTimeConcat(dData: TDate; tHora: TTime): TDateTime;
//  function ExecMethod(OnObject: TObject; MethodName: string): String;
//  function ExecClassMethod(sClassName, sMethod: String): String;

  // Alterar intensidade da cor
  // Exemplo:
  //     cor1 := Intensidade(clMoneyGreen, -50); //fica 50% mais escura
  //     cor2 := Intensidade(clMoneyGreen, 70); //fica 70% mais clara
  function CorIntensidade(cCor: TColor; iValor: integer): TColor;

  // Reduzir a memória utilizada pelo sistema
  // http://www.agnaldocarmo.com.br/home/comando-milagroso-para-reducao-de-memoria-delphi/
  procedure TrimAppMemorySize;

  // Treeview - selecionar nodo a partir da descrição do item
  function GetNodeTreeViewByText(ATree : TTreeView; AValue:String; AVisible: Boolean): TTreeNode;

  // Ordernar Grid quando clica no título
  procedure SortTitleClick(Column: TColumn; bDesativar: Boolean=False);
  function TerminateProcessByFileName(const FileName: string): Boolean;
  Function ObterMacAddress: string;
  function GetLocalIP: string;

const
  DiaDoMes: array[1..12] of Integer = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
  NL = #13#10;
  MinusculaSemAcentoF: array[1..19] of Char = ('a','o','a','e','i','o','u','a','e','i','o','u','a','e','i','o','u','c','n');
  MaiusculaSemAcentoF: array[1..19] of Char = ('A','O','A','E','I','O','U','A','E','I','O','U','A','E','I','O','U','C','N');
  MinusculaAcentuadaF: array[1..19] of Char = ('ã','õ','á','é','í','ó','ú','à','è','ì','ò','ù','â','ê','î','ô','û','ç','ñ');
  MaiusculaAcentuadaF: array[1..19] of Char = ('Ã','Õ','Á','É','Í','Ó','Ú','À','È','Ì','Ò','Ù','Â','Ê','Î','Ô','Û','Ç','Ñ');

var
  DadosConfigur: TConfigur;

implementation

uses
//  {$IFDEF VER210} // Delphi 2010
  Tecnosoft, //ACBrNFeQRCodeBar, DCPtwofish, DCPSha1, AbZipKit, AbZipTyp, AbArcTyp,
//  {$ENDIF}
  ComObj, Variants, Math, IniFiles, jpeg,
  Registry, IdFTP,
  Winsock, SqlTimSt, UrlMon, Clipbrd, ShlObj, ActiveX,
  EncdDecd, Mask, DBCtrls, StrUtils, //RzDBEdit, CEP, uCEP, RzEdit,
  ShellAPI, FileCtrl, Messages, IdHash, IdHashMessageDigest, IdFTPCommon;



function GetLocalIP: string;
var
  WSAData: TWSAData;
  HostName: array[0..255] of AnsiChar;
  HostEnt: PHostEnt;
  Addr: PInAddr;
  IPList: TStringList;
  i: Integer;
begin
  Result := 'Desconhecido';
  IPList := TStringList.Create;
  try
    if WSAStartup($0202, WSAData) = 0 then
    try
      if GetHostName(HostName, SizeOf(HostName)) = 0 then
      begin
        HostEnt := GetHostByName(HostName);
        if HostEnt <> nil then
        begin
          Addr := PInAddr(HostEnt^.h_addr_list^);
          for i := 0 to HostEnt^.h_length - 1 do
            IPList.Add(Inet_ntoa(Addr^));
        end;
      end;
    finally
      WSACleanup;
    end;

    if IPList.Count > 0 then
      Result := IPList[0];
  finally
    IPList.Free;
  end;
end;


Function ObterMacAddress: string;
var
 Lib: Cardinal;
 Func: function(GUID: PGUID): Longint; stdcall;
 GUID1, GUID2: TGUID;
begin
 Result := '';
 Lib := LoadLibrary('rpcrt4.dll');
 if Lib <> 0 then
 begin
  @Func := GetProcAddress(Lib, 'UuidCreateSequential');
 if Assigned(Func) then
  begin
   if (Func(@GUID1) = 0) and
   (Func(@GUID2) = 0) and
   (GUID1.D4[2] = GUID2.D4[2]) and
   (GUID1.D4[3] = GUID2.D4[3]) and
   (GUID1.D4[4] = GUID2.D4[4]) and
   (GUID1.D4[5] = GUID2.D4[5]) and
   (GUID1.D4[6] = GUID2.D4[6]) and
   (GUID1.D4[7] = GUID2.D4[7]) then
     begin
      Result :=
      IntToHex(GUID1.D4[2], 2) + '-' +
      IntToHex(GUID1.D4[3], 2) + '-' +
      IntToHex(GUID1.D4[4], 2) + '-' +
      IntToHex(GUID1.D4[5], 2) + '-' +
      IntToHex(GUID1.D4[6], 2) + '-' +
      IntToHex(GUID1.D4[7], 2);
     end;
  end;
 end;
end;


function TerminateProcessByFileName(const FileName: string): Boolean;
var
  SnapShotHandle: THandle;
  ProcessEntry: TProcessEntry32;
  ProcessHandle: THandle;
begin
  Result := False;  // Inicializa o resultado como False

  SnapShotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if SnapShotHandle = INVALID_HANDLE_VALUE then Exit;

  ProcessEntry.dwSize := SizeOf(ProcessEntry);
  if not Process32First(SnapShotHandle, ProcessEntry) then
  begin
    CloseHandle(SnapShotHandle);
    Exit;
  end;

  repeat
    if CompareText(ProcessEntry.szExeFile, FileName) = 0 then
    begin
      ProcessHandle := OpenProcess(PROCESS_TERMINATE, False, ProcessEntry.th32ProcessID);
      if ProcessHandle <> 0 then
      begin
        try
          Result := TerminateProcess(ProcessHandle, 0);
        finally
          CloseHandle(ProcessHandle);
        end;
        Break;  // Sai do loop após tentar terminar o processo
      end;
    end;
  until not Process32Next(SnapShotHandle, ProcessEntry);

  CloseHandle(SnapShotHandle);
end;

function TirarZerosClassificacao(Classificacao: string):String;
var
  i, y : integer;
  Aux, Aux2 : String;
  Tira : Boolean;

begin
  Aux := Classificacao;
  i := Length(Aux)+1;

  repeat
    if (Aux[i] = '0') then begin
      y := i;
      Tira := True;
      repeat
        if Aux[y] in ['1'..'9'] then
           Tira := False;
        Dec(y);
      until (Aux[y] = '.') or (y = 1);
      if Tira then
         Aux[i] := ' ';
    end;

    if (Aux[i] = '.') then
       Aux[i] := ' ';
    dec(i);
  until Aux[i] in ['1'..'9'];
  Result := Trim(Aux);
end;

function DiaUtil(dData: TDateTime): Boolean;
begin                                  // ConfU  Tirad  Traba  Indep  NSApar Finad  Repub  Natal
  if Pos(FormatDateTime('dd/mm', dData), '01/01| 21/04| 01/05| 07/09| 12/10| 02/11| 15/11| 25/12| ')>0 then
    Result := False
  else if (DayOfWeek(dData) in [2..6]) then
    Result := True
  else
    Result := False;
end;

function ProximoDiaUtil(Connec: TIBCConnection; DataIni: TDateTime; bAntecipar: Boolean=False):TDateTime;
var
  dData : TDateTime;
  iCont : Integer;
begin
  dData := DataIni;
  iCont := 0;
  repeat
    if (not DiaUtil(dData)) or (isFeriado(Connec, dData)) then
    begin
      if bAntecipar then
        dData := dData - 1
      else
        dData := dData + 1;
    end
    else
      iCont := 25;
    inc(iCont)
  until (iCont > 20);
  Result := dData;
end;

function RoundRGF(AValue: Variant; Const ADecimal: integer; ATypeRound:TTypeRound=trDef): Double;
var
  strValue : string;
  AValueDouble : Double;
  Value,Factor,Fraction: Extended;
begin
  try
    if ATypeRound=trDef then // Padrão
    begin
//      AValueDouble:= StrToFloat(VarToStr(AValue));
//      strValue:= FloatToStrF(AValueDouble,ffFixed,18,ADecimal);
//      result:= StrToFloat(strValue);
      Result := SimpleRoundTo(aValue, Abs(aDecimal)*-1);
    end
    else if ATypeRound=trUp then // Arredonda Decimal >= 0,5 p/ cima
    begin
      Factor := IntPower(10, ADecimal);
      Value := StrToFloat(FloatToStr(AValue * Factor));
      Result := Int(Value);
      Fraction := Frac(Value);
      if Fraction >= 0.5 then
        Result := Result + 1
      else if Fraction <= -0.5 then
        Result := Result - 1;
      Result := Result / Factor;
    end;
    if Result = null then
      Result := 0;
  except
    Result:= 0.00;
  end;
end;

procedure DesabilitarFecharBorderIcon(Sender: TObject);
var
  hSysMenu: HMENU;
begin
  hSysMenu := GetSystemMenu(TForm(Sender).Handle, False);
  if hSysMenu <> 0 then
  begin
    EnableMenuItem(hSysMenu, SC_CLOSE,MF_BYCOMMAND Or MF_GRAYED);
    DrawMenuBar(TForm(Sender).Handle);
  end;
end;

function iif(bCondicao: Boolean; vValorTrue, vValorFalse: Variant): Variant;
begin
  case bCondicao of
    True: Result:=vValorTrue;
    False:Result:=vValorFalse
  end;
end;

function PadC(const S: string; const Len: integer; const Ch: Char): string;
var
  I, J: integer;
  Pad: string;
  Impar: boolean;
begin
  I := Length(S);
  if I < Len then
  begin
    J := Len - I;
    Impar := J mod 2 = 1;
    J := J div 2;
    Pad := StringOfChar(Ch, J);
    Result := Pad + S + Pad;
    if Impar then
      Result := Result + Ch;
  end
  else if I > Len then
  begin
    J := I - Len;
    Impar := J mod 2 = 1;
    J := J div 2;
    Result := S;
    Delete(Result, I-J+1, J);
    Delete(Result, 1, J);
    if Impar then
    begin
      Dec(I, J * 2);
      Delete(Result, I, 1);
    end;
  end
  else
    Result := S;
end;

function PadL(ctext : string; nlen:integer ; cchar : char) : string ;
var
  creturn : string ;
  ntemp : integer ;
  nmax  : integer ;
begin
  creturn := AllTrim(ctext) ;
  nmax    := nlen ;
  dec(nmax,length(creturn)) ;
  if nmax > 0 then
    begin
      for ntemp := 1 to nmax do
        creturn := concat(cchar,creturn) ;
      result := creturn ;
    end
  else
    result := rightString(creturn,nlen) ;
end;

function padR(ctext : string; nlen: integer ; cchar : char) : string ;
var
  creturn : string ;
  ntemp : integer ;
  nmax  : integer ;
begin
  creturn := allTrim(ctext) ;
  nmax := (nlen - length(creturn)) ;
  if (nmax > 0) then
    begin
      for ntemp := 1 to nmax do
        creturn := concat(creturn,cchar) ;
    end
  else
    creturn := copy(creturn,1,nlen) ;
  result := creturn ;
end;

//{$ifdef VER210}
{function GridToExcelFile(Grid: TDbAltGrid; ExcelFile: String; TotalRegistros: Integer; pbAux: TObject; var bCancelarExportXLS: Boolean):Boolean;
var
  bResult  : Boolean;
  SavePlace: TBookmark;
  i,eline  : Integer;
  Excel    : Variant;
  iSheet   : Integer;
  CorFundo : TColor;
  oAfterScroll, oCalcFields: TDataSetNotifyEvent;
begin
  bResult:= False;
  // If dataset is assigned and active runs Excel
  if Assigned(Grid.DataSource) then
  begin
    if Grid.DataSource.DataSet.Active then
    begin
      try
        if pbAux is TProgressBar then
        begin
          TProgressBar(pbAux).Position := 0;
          TProgressBar(pbAux).Max      := TotalRegistros;
        end
        else if pbAux is TRGFProgressBar then
        begin
          TRGFProgressBar(pbAux).PartsComplete := 0;
          TRGFProgressBar(pbAux).TotalParts    := TotalRegistros;
        end;

        Grid.DataSource.DataSet.DisableControls;
        oAfterScroll := Grid.DataSource.DataSet.AfterScroll;
//        oCalcFields  := Grid.DataSource.DataSet.OnCalcFields;
        Grid.DataSource.DataSet.AfterScroll := nil;
//        Grid.DataSource.DataSet.OnCalcFields:= nil;
        Grid.Enabled := False;

        if not DirectoryExists(ExtractFilePath(ExcelFile)) then
          CreateDir(ExtractFilePath(ExcelFile));

        try
          //Rotina para setar um painel com um ProgressBar
  //                SetaPainelMensagem(cExportandoRegistros, TotalRegistros);

          Excel:= CreateOleObject('Excel.Application');
          Excel.Visible:= False;
          Excel.WorkBooks.Add;

          //Definindo o número de worksheets
          if  (TotalRegistros > 65000) then
          begin
            if  ((TotalRegistros Mod 65000) = 0) then
              iSheet := TotalRegistros DIV 65000
            else
              iSheet := (TotalRegistros DIV 65000) + 1;
            if  (iSheet > 3) then
              //Adicionando as worksheets que faltam a partir da 3 planilha do excel
              for i:= 4 to iSheet do
                Excel.WorkBooks[1].Sheets.Add(1, Excel.WorkBooks[1].Sheets[i-1]);
          end;
          // Save grid Position
          SavePlace:= Grid.DataSource.DataSet.GetBookmark;
          Grid.DataSource.DataSet.First;
          //Sheet atual
          iSheet := 1;
          // Montando cabeçalho da planilha
          if not (Grid.DataSource.DataSet.Eof) then
          begin
            eline:= 1; // Posicionando na primeira linha da planilha(Sheet) para por o cabeçalho
            for i:=0 to (Grid.Columns.Count-1) do
            begin
              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)]                  := Grid.Columns[i].Title.Caption;
              if Grid.Columns[i].Field.DisplayWidth > 250 then
                 Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].ColumnWidth   := 250
              else
                 Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].ColumnWidth   := Grid.Columns[i].Field.DisplayWidth;
              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Font.FontStyle:= 'Negrito';
              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Interior.Color:= (ColorToRgb(Grid.Columns[i].Title.Color));
              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Font.Color    := (ColorToRgb(Grid.Columns[i].Title.Font.Color));
            end;
          end;
          while (not Grid.DataSource.DataSet.Eof) and (not bCancelarExportXLS) do //Preenchendo o restante da planilha com os dados
          begin
            Inc(eline); //Incrementa a posição da linha para preencher no excel
  //                  pbInformacao.StepBy(1);
            Application.ProcessMessages;
            //Se passar de 65000 linhas, jogar dado na outra planilha, remontando os cabeçalhos antes
            if (eline > 65000) then
            begin
              Inc(iSheet);
              eline := 1;
              for i:=0 to (Grid.Columns.Count-1) do
              begin
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)]               := Grid.Columns[i].Title.Caption;
                if Grid.Columns[i].Field.DisplayWidth > 250 then
                   Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].ColumnWidth   := 250
                else
                   Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].ColumnWidth   := Grid.Columns[i].Field.DisplayWidth;
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Font.FontStyle:= 'Negrito';
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Interior.Color:= (ColorToRgb(Grid.Columns[i].Title.Color));
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Font.Color    := (ColorToRgb(Grid.Columns[i].Title.Font.Color));
              end;
              Inc(eline);
            end;

            //Para mudar a cor de fundo da linha na planilha do excell
            if (eline mod 2) = 0 then
              CorFundo := clInfoBk
            else
              CorFundo := clAqua;

            for i:=0 to (Grid.Columns.Count-1) do
            begin
              if (Grid.Columns[i].Field.DataType in [ftDate,ftDateTime,ftTimeStamp]) and
                 (Grid.Columns[i].Field.Value < StrToDate('01/01/1900')) then
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Value := Grid.Columns[i].Field.AsString
              else if Grid.Columns[i].Field.DataType in [ftString, ftBlob, ftMemo, ftFmtMemo, ftFixedChar, ftWideString, ftFixedWideChar, ftWideMemo] then
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Value := Grid.Columns[i].Field.AsString
              else
                Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Value := Grid.Columns[i].Field.Value;
//              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Interior.Color:= (ColorToRgb(CorFundo));
//              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Font.Color    := (ColorToRgb(Grid.Columns[i].Font.Color));
//              Excel.WorkBooks[1].Sheets[iSheet].Cells[eline,(i+1)].Borders.Color := (ColorToRgb(clGray));
            end;

            if pbAux is TProgressBar then
            begin
              TProgressBar(pbAux).StepIt;
              TProgressBar(pbAux).Refresh;
            end
            else if pbAux is TRGFProgressBar then
            begin
              TRGFProgressBar(pbAux).IncPartsByOne;
            end;

            Grid.DataSource.DataSet.Next;
          end;

          //Ajustando o tamanho das colunas nas planilhas
          for i:= 1 to iSheet do
            Excel.WorkBooks[1].WorkSheets[i].Range['B1','AQ1000'].Columns.AutoFit;

          // Set saved grid position
          Grid.DataSource.DataSet.GotoBookmark(SavePlace);
          // Salvando o arquivo
          Excel.WorkBooks[1].SaveAs(ExcelFile);
          Excel.Visible := True;
//          Excel.Quit;
          bResult:= True;
  //                pnlMensagem.Visible := False;
        except
          on E: Exception do
          begin
            bResult:= False;
            Excel.Quit;
            raise Exception.Create(E.Message);
    //                pnlMensagem.Visible := False;
          end;
        end;
      finally
        Grid.DataSource.DataSet.AfterScroll := oAfterScroll;
//        Grid.DataSource.DataSet.OnCalcFields:= oCalcFields;
        Grid.DataSource.DataSet.EnableControls;
        Grid.Enabled := True;
      end;
    end;
  end;
  Result := bResult;
end;}
//{$endif}

{function getVersaoModulo(dbConnec: TDatabase; iIDModulo: Integer; FileName: PChar): TRGFVersion;
var
  qry: TQuery;
  ZeroValue, InfoSize: DWORD;
  Buffer: Pointer;
  Version: Pointer;
  InfoLen: UINT;
  Major, Minor, Release, Build: Integer;
begin
  qry := TQuery.Create(nil);
  try
    try
      qry.DatabaseName := dbConnec.DatabaseName;
      qry.SQL.Text := 'select * from rgfversion where id = :id ';
      qry.ParamByName('id').AsInteger := iIDModulo;
      qry.Open;
    except
      Application.Terminate;
    end;
    if not qry.Fields[0].IsNull then
    begin
      Result.ID          := iIDModulo;
      Result.Descricao   := qry.FieldByName('descricao').AsString;
      Result.Sigla       := qry.FieldByName('sigla').AsString;
      Result.MajorVersion:= qry.FieldByName('vmajorversion').AsInteger;
      Result.MinorVersion:= qry.FieldByName('vminorversion').AsInteger;
      Result.Release     := qry.FieldByName('vrelease').AsInteger;
      InfoSize := GetFileVersionInfoSize(FileName, ZeroValue);
      GetMem(Buffer, InfoSize);
      try
        GetFileVersionInfo(FileName, 0, InfoSize, Buffer);
        VerQueryValue(Buffer, '\', Version, InfoLen);
        Major   := PVSFixedFileInfo(Version)^.dwFileVersionMS shr 16;
        Minor   := PVSFixedFileInfo(Version)^.dwFileVersionMS shl 16;
        Minor   := Minor shr 16;
        Release := PVSFixedFileInfo(Version)^.dwFileVersionLS shr 16;
        Build   := PVSFixedFileInfo(Version)^.dwFileVersionLS shl 16;
        Build   := Build shr 16;
        Result.Build := Build;
      finally
        FreeMem(Buffer, InfoSize);
      end;
    end;

    if (Result.MajorVersion <> Major) or (Result.MinorVersion <> Minor) or (Result.Release <> Release) then
    begin
      Application.MessageBox(PWideChar('Detectada versão do sistema incompatível com a versão'+NL+
                             'atual do banco de dados. Impossível prosseguir. Contate o administrador do sistema.'+NL+
                             'Versão do Sistema: '+Format('%d.%d.%d',[Major, Minor, Release])+NL+
                             'Versão do Banco de Dados: '+Format('%d.%d.%d',[Result.MajorVersion, Result.MinorVersion, Result.Release])),
                             'Versão desatualiza', MB_OK + MB_ICONSTOP);
      Application.MessageBox(PWideChar('Detectada versão do sistema incompatível com a versão'+NL+
                             'atual do banco de dados. Impossível prosseguir. Contate o administrador do sistema.'+NL+
                             'Versão do Sistema: '+Format('%d.%d.%d',[Major, Minor, Release])+NL+
                             'Versão do Banco de Dados: '+Format('%d.%d.%d',[Result.MajorVersion, Result.MinorVersion, Result.Release])),
                             'Versão desatualiza', MB_OK + MB_ICONSTOP);
      Application.Terminate;
    end;
  finally
    FreeAndNil(qry);
  end;

end;}

{function getVersaoModulo(Connec: TIBCConnection; iIDModulo: Integer; FileName: PChar): TRGFVersion;
var
  qry: TIBCQuery;
  ZeroValue, InfoSize: DWORD;
  Buffer: Pointer;
  Version: Pointer;
  InfoLen: UINT;
  Major, Minor, Release, Build: Integer;
begin
  qry := TIBCQuery.Create(nil);
  try
    try
      qry.Connection := Connec;
      qry.Transaction:= Connec.DefaultTransaction;
      qry.SQL.Text := 'select * from rgfversion where id = :id ';
      qry.ParamByName('id').AsInteger := iIDModulo;
      qry.Open;
    except
      Application.Terminate;
    end;
    if not qry.Fields[0].IsNull then
    begin
      Result.ID          := iIDModulo;
      Result.Descricao   := qry.FieldByName('descricao').AsString;
      Result.Sigla       := qry.FieldByName('sigla').AsString;
      Result.MajorVersion:= qry.FieldByName('vmajorversion').AsInteger;
      Result.MinorVersion:= qry.FieldByName('vminorversion').AsInteger;
      Result.Release     := qry.FieldByName('vrelease').AsInteger;
      InfoSize := GetFileVersionInfoSize(FileName, ZeroValue);
      GetMem(Buffer, InfoSize);
      try
        GetFileVersionInfo(FileName, 0, InfoSize, Buffer);
        VerQueryValue(Buffer, '\', Version, InfoLen);
        Major   := PVSFixedFileInfo(Version)^.dwFileVersionMS shr 16;
        Minor   := PVSFixedFileInfo(Version)^.dwFileVersionMS shl 16;
        Minor   := Minor shr 16;
        Release := PVSFixedFileInfo(Version)^.dwFileVersionLS shr 16;
        Build   := PVSFixedFileInfo(Version)^.dwFileVersionLS shl 16;
        Build   := Build shr 16;
        Result.Build := Build;
      finally
        FreeMem(Buffer, InfoSize);
      end;
    end;

    if (Result.MajorVersion <> Major) or (Result.MinorVersion <> Minor) or (Result.Release <> Release) then
    begin
      Application.MessageBox(PWideChar('Detectada versão do sistema incompatível com a versão'+NL+
                             'atual do banco de dados. Impossível prosseguir. Contate o administrador do sistema.'+NL+
                             'Versão do Sistema: '+Format('%d.%d.%d',[Major, Minor, Release])+NL+
                             'Versão do Banco de Dados: '+Format('%d.%d.%d',[Result.MajorVersion, Result.MinorVersion, Result.Release])),
                             'Versão desatualiza', MB_OK + MB_ICONSTOP);
      Application.MessageBox(PWideChar('Detectada versão do sistema incompatível com a versão'+NL+
                             'atual do banco de dados. Impossível prosseguir. Contate o administrador do sistema.'+NL+
                             'Versão do Sistema: '+Format('%d.%d.%d',[Major, Minor, Release])+NL+
                             'Versão do Banco de Dados: '+Format('%d.%d.%d',[Result.MajorVersion, Result.MinorVersion, Result.Release])),
                             'Versão desatualiza', MB_OK + MB_ICONSTOP);
      Application.Terminate;
    end;
  finally
    FreeAndNil(qry);
  end;
end;}

{function getVersaoArquivo(FileName: PChar): TRGFVersion;
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
end;}

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

procedure GetFilesFrom(sDirName,sFilter: String;bRecursive:Boolean;bDelete:Boolean=false;lstOut:TStringList=nil);
var
  SearchRec : TSearchRec;
  lstDirs : TStringList;
  IDFile : Integer;
  i : Integer;
begin
  if not Assigned(lstOut) then
    lstOut:=TStringList.Create;
  try
    IDFile:=FindFirst(ValidaPath(sDirName)+sFilter,0,SearchRec);
    while IDFile=0 do
    begin
      lstOut.Add(WideString(ValidaPath(sDirName)+SearchRec.Name));
      if bDelete then SysUtils.DeleteFile(ValidaPath(sDirName)+SearchRec.Name);
      IDFile:=FindNext(SearchRec);
    end;
    FindClose(SearchRec);
    if not bRecursive then
    begin
      if bDelete then SysUtils.RemoveDir(ValidaPath(sDirName));
      exit;
    end;
    lstDirs:= TStringList.Create;
    try
      IDFile:= FindFirst(ValidaPath(sDirName)+'*.*',faDirectory,SearchRec);
      while IDFile=0 do
      begin
        if ((SearchRec.Attr and faDirectory) <> 0) and (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
          lstDirs.Add(ValidaPath(sDirName)+SearchRec.Name);
        IDFile:= FindNext(SearchRec);
      end;
      FindClose(SearchRec);
      for i := 0 to lstDirs.Count-1 do
        GetFilesFrom(ValidaPath(lstDirs[i]),sFilter,bRecursive,bDelete,lstOut);
    finally
      lstDirs.Free;
    end;
    if bDelete then SysUtils.RemoveDir(ValidaPath(sDirName));
  except
    on e: Exception do
    begin
      Application.MessageBox(PWideChar(widestring(e.Message)), 'Versão desatualiza', MB_OK + MB_ICONSTOP);
      Application.MessageBox(PWideChar(widestring(e.Message)), 'Versão desatualiza', MB_OK + MB_ICONSTOP);
    end;
  end;
end;

function ValidaPath(sPath:String;sDelim:String=''):String;
begin
   Result:=sPath;
   if Length(sDelim)<=0 then
      sDelim:=PathDelim;
   if Copy(sPath,Length(sPath),1)<>sDelim then
     Result:=sPath+sDelim;
end;

function InitKey: String;
begin
// a9x35fx0RGF
  Result := Chr($61)+Chr($39)+Chr($78)+Chr($33)+Chr($35)+Chr($66)+Chr($78)+
            Chr($30)+Chr($52)+Chr($47)+Chr($46);
end;

//{$IFDEF VER210} // Delphi 2010
{function Encrypt(sStr: String; sKey: String=''): String;
var
  Twofish: TDCP_twofish;
//  Key: String;
begin
  TwoFish := TDCP_twofish.Create(nil);
  with TwoFish do
  begin
    Algorithm := 'Twofish';
    MaxKeySize := 256;
    BlockSize := 128;
  end;
  if Trim(sKey) = '' then
    sKey := InitKey;
  TwoFish.InitStr(sKey, TDCP_sha1);
  Result := TwoFish.EncryptString(AnsiString(sStr));
  TwoFish.Burn;
end;}

{function Decrypt(sStr: String; sKey: String=''): String;
var
  Twofish: TDCP_twofish;
//  Key: String;
begin
  TwoFish := TDCP_twofish.Create(nil);
  with TwoFish do
  begin
    Algorithm := 'Twofish';
    MaxKeySize := 256;
    BlockSize := 128;
  end;
  if Trim(sKey) = '' then
    sKey := InitKey;
  TwoFish.InitStr(sKey, TDCP_sha1);
  Result := TwoFish.DecryptString(AnsiString(sStr));
  TwoFish.Burn;
end;}
//{$ENDIF}

//Const
//  // Chaves de encriptação
//  StKey = 8420685;
//  MtKey = 1158212;
//  AdKey = 2023849;

// ************************ Funcões Encriptação/Desencriptação **********************
// PARA ENCRIPTAR
{$R-} {$Q-}
// Habilita/Desabilita a geração de checagem de código de Faixa e
// de checagem de código exceção de overflow
function EncryptSTR(Const InString:String; StartKey,MultKey,AddKey:Integer): String;
var
  I : Byte;
begin
  Result := '';
  for I := 1 to Length(InString) do
  begin
    Result := Result + Char(Byte(InString[I]) xor (StartKey shr 8));
    StartKey := (Byte(Result[I]) + StartKey) * MultKey + AddKey;
  end;
end;

// PARA DESENCRIPTAR
function DecryptSTR(Const InString: String; StartKey, MultKey, AddKey: Integer): String;
var I : Byte;
begin
  Result := '';
  for I := 1 to Length(InString) do
  begin
    Result := Result + Char(Byte(InString[I]) xor (StartKey shr 8));
    StartKey := (Byte(InString[I]) + StartKey) * MultKey + AddKey;
  end;
end;
{$R+} {$Q+}
// ************************ Funcões Encriptação/Desencriptação **********************

function Criptografia(mStr, mChave: string): string;
var
i, TamanhoString, pos, PosLetra, TamanhoChave: Integer;
begin
  Result := mStr;
  TamanhoString := Length(mStr);
  TamanhoChave := Length(mChave);
  for i := 1 to TamanhoString do
  begin
    pos := (i mod TamanhoChave);
    if pos = 0 then
    pos := TamanhoChave;
    posLetra := ord(Result[i]) xor ord(mChave[pos]);
    if posLetra = 0 then
    posLetra := ord(Result[i]);
    Result[i] := chr(posLetra);
  end;
end;

function getMD5HashString(value: string): string;
var
  hashMessageDigest5 : TIdHashMessageDigest5;
begin
  hashMessageDigest5 := nil;
  try
    hashMessageDigest5 := TIdHashMessageDigest5.Create;
    Result := IdGlobal.IndyLowerCase ( hashMessageDigest5.HashStringAsHex ( value ) );
//    Result := hashMessageDigest5.HashStringAsHex(value);
  finally
    hashMessageDigest5.Free;
  end;
end;

procedure GetFileList(ASource: string; lstDir:Boolean;var lstFiles : TStringList; bIncrement:Boolean=false);
var
  SearchRec : TSearchRec;
  i,posi : Integer;
  lstAux:TStringList;
begin
  lstAux:=TStringList.Create;
  try
    posi := FindFirst( ASource, faAnyFile, SearchRec );
    if posi = 0 then
      while (posi = 0) do
      begin
        if not lstDir then {somente arquivos}
        begin
          if (SearchRec.Attr and faDirectory) <> faDirectory then
          begin
            lstAux.Add(SearchRec.Name);
            posi := FindNext(SearchRec);
          end
          else
            posi := FindNext( SearchRec );
        end
        else {todos arquivos e diretorios exceto (.  ..)}
        begin
          if (SearchRec.Attr and faDirectory) = faDirectory then
            if (SearchRec.Name+' ')[1] = '.' then
            begin
              posi := FindNext(SearchRec);
              Continue;
            end;
          lstAux.Add(SearchRec.Name);
          posi := FindNext(SearchRec);
        end;
      end;
    FindClose( SearchRec );
    lstAux.Sort;

    if not bIncrement then
      lstFiles:=TStringlist.Create;

    for i := 0 to lstAux.Count - 1 do lstFiles.Add(lstAux[i]);
  finally
    lstAux.free;
  end;
end;

function DigitoABAC_EAN(S: String): Char;
var
  I: Integer;
  T1, T2: Longint;
begin
  T1 := 0;
  T2 := 0;
  I := Length(S);
  while I > 0 do begin
    T1 := T1 + (Ord(S[I]) - 48);
    I := I - 1;
    if I > 0 then begin
      T2 := T2 + (Ord(S[I]) - 48);
      I := I - 1;
    end;
  end;
  T1 := (T1 * 3) + T2;
  T1 := T1 mod 10;
  if T1 = 0 then Result := '0'
  else Result := Chr((10 - T1) + 48);
end;

function ValidaEAN(CodBarras: String):String;
var
  x, Digito: String;
begin
  if (Length(CodBarras) = 8) or (Length(CodBarras) = 13) or (Length(CodBarras) = 14) then
  begin
    x := Copy(CodBarras, 1, Length(CodBarras) - 1);
    Digito := DigitoABAC_EAN(x);
    x:=  Copy(CodBarras, Length(CodBarras), 1);
    if Digito <> x then
      x := '0000000000000'
//      x := x + DigitoABAC_EAN(x);
    else
      x := CodBarras;
  end;

  Result := x;
end;

function  AdicionaMes(data: TDateTime; Valor: Integer): TDateTime;
var
  Dia, Mes, Ano: Word;
  Meses: Integer;
  Datas: TDateTime;
begin
  DecodeDate(data, Ano, Mes, Dia);
  Meses := Mes + Valor;
  repeat
    if Meses > 12 then
    begin
      Datas := AdicionaAno(EncodeDate(Ano, Mes, Dia), 1);
      DecodeDate(datas, Ano, Mes, Dia);
      Meses := Meses - 12;
    end;
  until Meses <= 12;
  Mes := Meses;
  if (Mes=2) and AnoBisexto(EncodeDate(Ano, 1, 1)) then
  begin
    if Dia > 29 then Dia := 29;
  end else
  begin
    if Dia > DiaDoMes[Mes] then Dia := DiaDoMes[Mes];
  end;
  Result := EncodeDate(Ano, Mes, Dia);
end;

function  AdicionaAno(data: TDateTime; Valor: Integer): TDateTime;
var
  Dia, Mes, Ano: Word;
begin
  DecodeDate(data, Ano, Mes, Dia);
  Ano := Ano + Valor;
  if (Mes=2) and AnoBisexto(EncodeDate(Ano, 1, 1)) then
  begin
    if Dia > 29 then Dia := 29;
  end else
  begin
    if Dia > DiaDoMes[Mes] then Dia := DiaDoMes[Mes];
  end;
  Result := EncodeDate(Ano, Mes, Dia);
end;

function  SubtraiMes(data: TDateTime; Valor: Integer): TDateTime;
var
  Dia, Mes, Ano: Word;
  Meses: Integer;
  Datas: TDateTime;
begin
  DecodeDate(data, Ano, Mes, Dia);
  Meses := Mes - Valor;
  repeat
    if Meses <= 0 then
    begin
      Datas := SubtraiAno(EncodeDate(Ano, Mes, Dia), 1);
      DecodeDate(datas, Ano, Mes, Dia);
      Meses := Meses + 12;
    end;
  until Meses > 0;
  Mes := Meses;
  if (Mes=2) and AnoBisexto(EncodeDate(Ano, 1, 1)) then
  begin
    if Dia > 29 then Dia := 29;
  end else
  begin
    if Dia > DiaDoMes[Mes] then Dia := DiaDoMes[Mes];
  end;
  Result := EncodeDate(Ano, Mes, Dia);
end;

function  SubtraiAno(data: TDateTime; Valor: Integer): TDateTime;
var
  Dia, Mes, Ano: Word;
begin
  DecodeDate(data, Ano, Mes, Dia);
  Ano := Ano - Valor;
  if (Mes=2) and AnoBisexto(EncodeDate(Ano, 1, 1)) then
  begin
    if Dia > 29 then Dia := 29;
  end else
  begin
    if Dia > DiaDoMes[Mes] then Dia := DiaDoMes[Mes];
  end;
  Result := EncodeDate(Ano, Mes, Dia);
end;

function RetornaUltimoDiaMes(Data : TDateTime) : TDateTime;
begin
  Result:=StrToDate(IntToStr(RetornaUltimoDia(Data))+'/'+IntToStr(MonthOf(Data))+'/'+IntToStr(YearOf(Data)));
end;

function RetornaPrimeiroDiaMes(Data : TDateTime) : TDateTime;
var
  iAno, iMes, iDia: Word;
begin
  DecodeDate(Data, iAno, iMes, iDia);
  Result := EncodeDate(iAno, iMes, 1);
end;

function RetornaUltimoDia(Data : TDateTime) : Integer;
var
  Dia, Mes, Ano : word;
Const
  aDias : array[1..12] of integer = (31,28,31,30,31,30,31,31,30,31,30,31);
begin
  DecodeDate(Data,Ano,Mes,Dia);
  Dia := aDias[Mes];
  if Mes = 2 then // ver ano bisexto
  begin
    Data := EncodeDate(Ano,Mes,1);
    if AnoBisexto(Data) then
      Dia := 29;
  end;
  Result := Dia
end;

function AnoBisexto(Data: TDateTime): Boolean;
var
  Dia,Mes,Ano : Word;
begin
  DecodeDate(Data,Ano,Mes,Dia);
  if Ano mod 4 <> 0 then
    Result := False
  else if Ano mod 100 <> 0 then
    Result := True
  else if Ano mod 400 <> 0 then
    Result := False
  else
    Result := True;
end;

function RetornarMesExtenso(iMes: Integer): String;
begin
  case iMEs of
    1: Result := 'Janeiro';
    2: Result := 'Fevereiro';
    3: Result := 'Março';
    4: Result := 'Abril';
    5: Result := 'Maio';
    6: Result := 'Junho';
    7: Result := 'Julho';
    8: Result := 'Agosto';
    9: Result := 'Setembro';
    10: Result := 'Outubro';
    11: Result := 'Novembro';
    12: Result := 'Dezembro';
  end;
end;

function ExtractFilePath_RGF(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter(iif(Pos('/',FileName)>0,'/','\') + ':', FileName);
  Result := Copy(FileName, 1, I);
end;

function ExtractFileName_RGF(const FileName: string): string;
var
  I: Integer;
begin
  I := LastDelimiter(iif(Pos('/',FileName)>0,'/','\') + ':', FileName);
  Result := Copy(FileName, I + 1, MaxInt);
end;

function IsFileInUse(FileName: TFileName): Boolean;
var
  HFileRes: HFILE;
begin
  Result := False;
  if not FileExists(FileName) then Exit;
  HFileRes := CreateFile(PChar(FileName),
                         GENERIC_READ or GENERIC_WRITE,
                         0,
                         nil,
                         OPEN_EXISTING,
                         FILE_ATTRIBUTE_NORMAL,
                         0);
  Result := (HFileRes = INVALID_HANDLE_VALUE);
  if not Result then
    CloseHandle(HFileRes);
end;

function GetFileDate(Arquivo: String): TDateTime;
var
  FHandle: integer;
begin
  if FileExists(Arquivo) then
  begin
    FHandle := FileOpen(Arquivo, 0);
    try
      Result := FileDateToDateTime(FileGetDate(FHandle));
    finally
      FileClose(FHandle);
    end;
  end;
end;

procedure SepararArquivosPorData(sPath, sFiltro: String; sMascara: String='yyyy-mm-dd');
var
  i: Integer;
  Arquivos: TStringList;
begin
  Arquivos := TStringList.Create;
  try
                       //'*.*'
    GetFilesFrom(sPath, sFiltro, False, False, Arquivos);

    for i := 0 to Arquivos.Count - 1 do
    begin
      if not DirectoryExists(ExtractFilePath(Arquivos[i]) + FormatDateTime(sMascara,GetFileDate(Arquivos[i]))+'\') then
        CreateDir(ExtractFilePath(Arquivos[i]) + FormatDateTime(sMascara,GetFileDate(Arquivos[i]))+'\');
      try
        if not IsFileInUse(Arquivos[i]) then
          MoveFile(PChar(Arquivos[i]), PChar(ExtractFilePath(Arquivos[i]) + FormatDateTime(sMascara,GetFileDate(Arquivos[i]))+'\' + ExtractFileName(Arquivos[i])));
      except
        on E: Exception do
        begin
          raise Exception.Create(PWideChar(WideString('Erro ao mover arquivo: '+Arquivos[i]+NL+E.Message)));
        end;
      end;
    end;
  finally
    Arquivos.Free;
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
        DeleteFile(slArquivos[i]);
end;

function getDatabase(Connec: TIBCConnection; bBancoAtualizacao: Boolean=False): string;
begin
  Result := Connec.Server+':'+Connec.Database;
  if bBancoAtualizacao then
    Result := ExtractFilePath_RGF(Result)+'atualizacao.fdb'
end;

function getDatabaseBDE(sAliasName: String; bBancoAtualizacao: Boolean): String;
var
  ParamList : TStringList;
begin
  Result := '';
  if Session.isAlias(sAliasName) then
  begin
    ParamList := TStringList.Create;
    try
      Session.GetAliasParams(sAliasName, ParamList);
      if bBancoAtualizacao then
        Result := ExtractFilePath_RGF(ParamList.Values['SERVER NAME'])+'atualizacao.fdb'
      else
        Result := ParamList.Values['SERVER NAME'];
    finally
      ParamList.Free;
    end;
  end;
end;

function getDatabaseAtualizacao(sDatabase: String): String;
begin
  Result := ExtractFilePath_RGF(sDatabase)+'atualizacao.fdb';
end;

function getServer_Database(sDatabase: String): Variant;
var
  sIP, sPath: String;
begin
  sIP := Copy(sDatabase, 1, Pos(':', sDatabase)-1);
  sPath := Copy(sDatabase, Pos(':', sDatabase)+1, Length(sDatabase)-Pos(':', sDatabase)+1);
  Result := VarArrayOf([sIP, sPath]);
end;

procedure AbreForm(aClasseForm: TComponentClass; aForm: TForm);
begin
  Application.CreateForm(aClasseForm, aForm);
  try
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

function PutINIDB(Connec: TDatabase; Chave, Parametro: String; Valor: Variant; User: String): Boolean;
begin
  with TQuery.Create(nil) do
  begin
    try
      DatabaseName := Connec.DatabaseName;

      SQL.Text := 'update or insert into CONFLOCAL (USUARIOID, CHAVE, PARAMETRO, VALOR) '+
                  '  values (:USUARIOID, :CHAVE, :PARAMETRO, :VALOR) '+
                  '  matching(USUARIOID, CHAVE, PARAMETRO) ';
      ParamByName('USUARIOID').AsString  := UpperCase(User);
      ParamByName('CHAVE').AsString      := UpperCase(Chave);
      ParamByName('PARAMETRO').AsString  := UpperCase(Parametro);
      ParamByName('VALOR').AsString      := Valor;
      try
        ExecSQL;
        Result := True;
      except
        Result := False;
      end;
    finally
      Free;
    end;
  end;
end;


function DelINIDB(Connec: TIBCConnection; Chave, Parametro: String; User: String): Boolean;
begin
  with TIBCQuery.Create(nil) do
  begin
    try
      Connection  := Connec;
      Transaction := Connec.DefaultTransaction;

      if not Connec.InTransaction then
        Connec.StartTransaction;

      try
        SQL.Text := 'delete from CONFLOCAL where USUARIOID = :USUARIOID and CHAVE = :CHAVE and PARAMETRO = :PARAMETRO ';
        ParamByName('USUARIOID').AsString  := UpperCase(User);
        ParamByName('CHAVE').AsString      := UpperCase(Chave);
        ParamByName('PARAMETRO').AsString  := UpperCase(Parametro);
        ExecSQL;
        Connec.CommitRetaining;

        Result := True;
      except
        on E: Exception do
        begin
          if Connection.InTransaction then
            Connec.RollbackRetaining;

          Result := False;
          MessageBoxW(0, PWideChar(WideString('Não foi possível deletar configuração: '+E.Message)), 'Erro', MB_OK + MB_ICONSTOP + MB_TOPMOST);

        end;
      end;
    finally
      Free;
    end;
  end;
end;

function GetINIDB(Connec: TDatabase; Chave, Parametro: String; vDefault: Variant; User: String): Variant;
begin
  with TQuery.Create(nil) do
  begin
    try
      DatabaseName := Connec.DatabaseName;
      SQL.Text := 'select * from CONFLOCAL where USUARIOID = :USUARIOID and CHAVE = :CHAVE and PARAMETRO = :PARAMETRO ';
      ParamByName('USUARIOID').AsString   := UpperCase(User);
      ParamByName('CHAVE').AsString       := UpperCase(Chave);
      ParamByName('PARAMETRO').AsString   := UpperCase(Parametro);
      Open;
      if FieldByName('VALOR').IsNull then
        Result := vDefault
      else
        Result := FieldByName('VALOR').Value;
    finally
      Free;
    end;
  end;
end;

function GetConfigur(Connec: TIBCConnection; sCampo: String; vValorPadrao: Variant; bConfigEspecial: Boolean): Variant;
begin
  if Trim(sCampo) = '' then
  begin
    MessageBoxW(0, PWideChar(WideString('Necessário informar um campo da tabela '+iif(bConfigEspecial, 'CONFIGURESP.', 'CONFIGUR.'))), 'Erro', MB_OK +MB_ICONSTOP + MB_TOPMOST);
    Abort;
  end;

  if (not Assigned(DadosConfigur)) or (not DadosConfigur.Criado) then
    DadosConfigur := TConfigur.Create(Connec);

  Result := DadosConfigur.GetValue(sCampo, vValorPadrao, bConfigEspecial);
end;

function getConfiguracaoEmpresa(Connec: TIBCConnection; iEmpresa: Integer; sChave, sParametro: String; sValorPadrao: Variant): Variant;
begin
  if Trim(sParametro) = '' then
  begin
    MessageBoxW(0, PWideChar(WideString('Necessário informar a empresa, chave e parâmetro. ')), 'Erro', MB_OK +MB_ICONSTOP + MB_TOPMOST);
    Abort;
  end;

  if (not Assigned(DadosConfigur)) or (not DadosConfigur.Criado) then
    DadosConfigur := TConfigur.Create(Connec);
  DadosConfigur.NoEmpresa := iEmpresa;
  Result := DadosConfigur.getConfigEmpresa(sChave, sParametro, sValorPadrao);
end;

function putConfiguracaoEmpresa(Connec: TIBCConnection; iEmpresa: Integer; sChave, sParametro: String; sValorPadrao: Variant): Variant;
begin
  if Trim(sParametro) = '' then
  begin
    MessageBoxW(0, PWideChar(WideString('Necessário informar a empresa, chave e parâmetro. ')), 'Erro', MB_OK +MB_ICONSTOP + MB_TOPMOST);
    Abort;
  end;

  if (not Assigned(DadosConfigur)) or (not DadosConfigur.Criado) then
    DadosConfigur := TConfigur.Create(Connec);
  DadosConfigur.NoEmpresa := iEmpresa;
  DadosConfigur.AddConfigEmpresa(Connec, iEmpresa, sChave, sParametro, sValorPadrao);
  Result := sValorPadrao;
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

procedure RunDosInMemo(DosApp: String; AMemo: TMemo);
const
  ReadBuffer = 2400;
var
  Security : TSecurityAttributes;
  ReadPipe,WritePipe : THandle;
  start : TStartUpInfo;
  ProcessInfo : TProcessInformation;
  Buffer : PAnsiChar;
  BytesRead : DWord;
  Apprunning : DWord;
begin
  With Security do
  begin
    nlength := SizeOf(TSecurityAttributes) ;
    binherithandle := true;
    lpsecuritydescriptor := nil;
  end;
  if Createpipe (ReadPipe, WritePipe,
                 @Security, 0) then
  begin
    Buffer := AllocMem(ReadBuffer + 1) ;
    FillChar(Start,Sizeof(Start),#0) ;
    start.cb := SizeOf(start) ;
    start.hStdOutput := WritePipe;
    start.hStdInput := ReadPipe;
    start.dwFlags := STARTF_USESTDHANDLES +
                         STARTF_USESHOWWINDOW;
    start.wShowWindow := SW_HIDE;
    UniqueString(DosApp);
    if CreateProcessW(nil,
           PChar(DosApp),
           @Security,
           @Security,
           true,
           NORMAL_PRIORITY_CLASS,
           nil,
           nil,
           start,
           ProcessInfo) then
    begin
      repeat
        Apprunning := WaitForSingleObject(ProcessInfo.hProcess,100);
        Application.ProcessMessages;
      until (Apprunning <> WAIT_TIMEOUT);
      repeat
        BytesRead := 0;
        ReadFile(ReadPipe,Buffer[0], ReadBuffer,BytesRead,nil);
        Buffer[BytesRead]:= #0;
        OemToAnsi(Buffer,Buffer);
        AMemo.Text := AMemo.text + String(Buffer);
      until (BytesRead < ReadBuffer);
    end;
    FreeMem(Buffer);
    CloseHandle(ProcessInfo.hProcess);
    CloseHandle(ProcessInfo.hThread);
    CloseHandle(ReadPipe);
    CloseHandle(WritePipe);
  end;
end;

function RunDosAndReturn(DosApp: String): String;
const
  ReadBuffer = 2400;
var
  Security : TSecurityAttributes;
  ReadPipe,WritePipe : THandle;
  start : TStartUpInfo;
  ProcessInfo : TProcessInformation;
  Buffer : PAnsiChar;
  BytesRead : DWord;
  Apprunning : DWord;
begin
  With Security do
  begin
    nlength := SizeOf(TSecurityAttributes) ;
    binherithandle := true;
    lpsecuritydescriptor := nil;
  end;
  if Createpipe (ReadPipe, WritePipe,
                 @Security, 0) then
  begin
    Buffer := AllocMem(ReadBuffer + 1) ;
    FillChar(Start,Sizeof(Start),#0) ;
    start.cb := SizeOf(start) ;
    start.hStdOutput := WritePipe;
    start.hStdInput := ReadPipe;
    start.dwFlags := STARTF_USESTDHANDLES +
                         STARTF_USESHOWWINDOW;
    start.wShowWindow := SW_HIDE;
    UniqueString(DosApp);
    if CreateProcessW(nil,
           PChar(DosApp),
           @Security,
           @Security,
           true,
           NORMAL_PRIORITY_CLASS,
           nil,
           nil,
           start,
           ProcessInfo) then
    begin
      repeat
        Apprunning := WaitForSingleObject(ProcessInfo.hProcess,100);
        Application.ProcessMessages;
      until (Apprunning <> WAIT_TIMEOUT);
//      repeat
      BytesRead := 0;
      ReadFile(ReadPipe,Buffer[0], ReadBuffer,BytesRead,nil);
      Buffer[BytesRead]:= #0;
      OemToAnsi(Buffer,Buffer);
      Result := String(Buffer);
//      until (BytesRead < ReadBuffer);
    end;
    FreeMem(Buffer);
    CloseHandle(ProcessInfo.hProcess);
    CloseHandle(ProcessInfo.hThread);
    CloseHandle(ReadPipe);
    CloseHandle(WritePipe);
  end;
end;

function IsRemoteSession: boolean;
const
  SM_REMOTESESSION = $1000;
begin
  Result := GetSystemMetrics(SM_REMOTESESSION) <> 0;
end;

function GetCurrentSessionID: Integer;
type
  TProcessIdToSessionId = function(dwProcessId: DWORD; pSessionId: DWORD): BOOL; stdcall;
var
  ProcessIdToSessionId: TProcessIdToSessionId;
//  hWTSapi32dll: THandle;
  Lib : THandle;
  pSessionId : DWord;
begin
  Result := 0;
  Lib := GetModuleHandle('kernel32');
  if Lib <> 0 then
  begin
    ProcessIdToSessionId := GetProcAddress(Lib, 'ProcessIdToSessionId');
    if Assigned(ProcessIdToSessionId) then
    begin
      ProcessIdToSessionId(GetCurrentProcessId(), DWORD(@pSessionId));
      Result:= pSessionId;
    end;
  end;
end;

function GetComputerNameFunc: string;
var ipbuffer : string;
      nsize : dword;
begin
  nsize := 255;
  SetLength(ipbuffer,nsize);
  if GetComputerName(pchar(ipbuffer),nsize) then
    Result := ipbuffer;
end;

procedure FTP_Put(sHost, sUser, sPassword, sPathOrigem, sPathDestino: String; slArquivos: TStringList);
var
  IdFTP: TIdFTP;
  i: Integer;
  slDir: TStringList;
  sAux, sAux2: String;
begin
  GetFilesFrom(sPathOrigem, '*.*', True, False, slArquivos);

  if slArquivos.Count > 0 then
  begin
    slDir := TStringList.Create;
    try
      for i := 0 to slArquivos.Count-1 do
      begin
        sAux := StringReplace(ExtractFilePath_RGF(slArquivos[i]), sPathOrigem, '', [rfReplaceAll, rfIgnoreCase]);
        sAux := StringReplace(sAux, '/', '\', [rfReplaceAll, rfIgnoreCase]);
//        sAux := Copy(sAux, 1, Length(sAux)-1);
        while Pos('\',sAux) > 0 do
        begin
          sAux2 := Copy(sAux, 1, Pos('\', sAux));
          sAux  := StringReplace(sAux,'\','/',[rfIgnoreCase]);
          sAux2 := '/'+StringReplace(sAux2, '\', '/', [rfReplaceAll]);
          if (sAux2 > ' ') and (slDir.IndexOf(sAux2) = -1) then
          begin
            slDir.Add(sAux2);
            slArquivos.Add(sAux2);
          end;
        end;
      end;

      IdFTP := TIdFTP.Create(nil);
      with IdFTP do
      begin
        try
          Username := sUser; //'sistema';
          Password := sPassword; //'252552';
          Host     := sHost; //'192.168.20.1';
          Passive  := True;
          Connect;

          for i := 0 to slDir.Count - 1 do
          begin
            try
              MakeDir(sPathDestino+slDir[i]);
            except
            end;
          end;
          for i := 0 to slArquivos.Count - 1 do
          begin
            if ExtractFileName_RGF(slArquivos[i]) > ' ' then
            begin
              sAux := StringReplace(ExtractFilePath_RGF(slArquivos[i]),
                                    IncludeTrailingPathDelimiter(sPathOrigem), IncludeTrailingPathDelimiter(sPathDestino), [rfReplaceAll, rfIgnoreCase]);
              sAux := StringReplace(sAux, '\', '/', [rfReplaceAll, rfIgnoreCase]);
              ChangeDir(sAux);
              Put(slArquivos[i], ExtractFileName(slArquivos[i]));
            end;
          end;
        finally
          Disconnect;
          Free;
        end;
      end;

    finally
      slDir.Free;
    end;
  end;
end;

function FTP_PutFile(sHost, sUser, sPassword, sArquivoOrigem, sPathFileDestino: String): Boolean;
var
  IdFTP: TIdFTP;
  i: Integer;
  slDir: TStringList;
  sAux, sAux2: String;
begin
  Result := False;
  if FileExists(sArquivoOrigem) then
  begin
    slDir := TStringList.Create;
    try
      sAux := ExtractFilePath_RGF(sPathFileDestino);
      sAux := StringReplace(sAux, '/', '\', [rfReplaceAll, rfIgnoreCase]);
//        sAux := Copy(sAux, 1, Length(sAux)-1);
      while Pos('\',sAux) > 0 do
      begin
        sAux2 := Copy(sAux, 1, Pos('\', sAux));
        sAux  := StringReplace(sAux,'\','/',[rfIgnoreCase]);
        sAux2 := '/'+StringReplace(sAux2, '\', '/', [rfReplaceAll]);
        if (sAux2 > ' ') and (slDir.IndexOf(sAux2) = -1) then
        begin
          slDir.Add(sAux2);
//          slArquivos.Add(sAux2);
        end;
      end;

      IdFTP := TIdFTP.Create(nil);
      with IdFTP do
      begin
        try
          Username     := sUser; //'sistema';
          Password     := sPassword; //'252552';
          Host         := sHost; //'192.168.20.1';
          TransferType := ftBinary;
// coloquei o passive pra funcionar na APG
          Passive      := True;

          Connect;

          for i := 0 to slDir.Count - 1 do
          begin
            try
              ChangeDir(slDir[i]);
            except
              MakeDir(slDir[i]);
              ChangeDir(slDir[i]);
            end;
          end;

          if ExtractFileName_RGF(sArquivoOrigem) > ' ' then
          begin
//            sAux := StringReplace(sArquivoOrigem, '\', '/', [rfReplaceAll, rfIgnoreCase]);
//            ChangeDir(sAux);
            Put(sArquivoOrigem, ExtractFileName_RGF(sPathFileDestino));

            Result := True;
          end;
        finally
          Disconnect;
          Free;
        end;
      end;
    finally
      slDir.Free;
    end;
  end;
end;

function FTP_DeleteFile(sHost, sUser, sPassword, sArquivo: String): Boolean;
var
  IdFTP: TIdFTP;
begin
  Result := False;
  IdFTP := TIdFTP.Create(nil);
  with IdFTP do
  begin
    try
      Username     := sUser; //'sistema';
      Password     := sPassword; //'252552';
      Host         := sHost; //'192.168.20.1';
      TransferType := ftBinary;
      Passive      := True;
      Connect;

      if (IdFTP.Size(sArquivo) > 0) then // remove se existir
      begin
        IdFTP.Delete(sArquivo);
        Result := True;
      end;
    finally
      Disconnect;
      Free;
    end;
  end;
end;

function MascaraVlr(Mascara:string;Vlr:double;CtrlZerado:Boolean=false;CtrZeradoMantemTamanho:Boolean=false):string;
var
  oRes : string;
begin
  oRes:=FormatFloat(Mascara,Vlr);
  oRes:=InsereStr(oRes,Length(mascara),' ','E');
  if CtrlZerado then
    if Vlr=0 then
    begin
      oRes:='';
      if CtrZeradoMantemTamanho then
        oRes:=InsereStr(oRes,Length(mascara),' ','E');
    end;
  result := oRes;
end;

function Formata(const Texto: string; Tipo: String): string;
begin
  //FormatMaskText('0.000',Texto);

  Result := '';
  if AllTrim(Texto) = '' then exit;
  //for I := 1 to Length(CEP) do
  //  if CEP[I] in ['0'..'9'] then
  //    Result := Result + CEP[I];
  //if Length(Result) <> 8 then
  //  raise Exception.Create('CEP inválido.')
  //else

  if Tipo='CEP' then
    Result := Copy(Texto, 1, 2) + '.' +Copy(Texto, 3, 3) + '-' +Copy(Texto, 6, 3)
  else if (Tipo='CPF') or ((Tipo='CPF_CNPJ') and (Length(RetornaSoNumero(Texto)) = 11)) then
  begin
     Result := RetornaSoNumero(Texto);
     Result := Copy(Result,1,3)+'.'+
               Copy(Result,4,3)+'.'+
               Copy(Result,7,3)+'-'+
               Copy(Result,10,2);
  end
  else if (Tipo='CNPJ') or ((Tipo='CPF_CNPJ') and (Length(RetornaSoNumero(Texto)) = 14)) then
  begin
    Result := RetornaSoNumero(Texto);
    Result := Copy(Result,1,2)+'.'+
              Copy(Result,3,3)+'.'+
              Copy(Result,6,3)+'/'+
              Copy(Result,9,4)+'-'+
              Copy(Result,13,2);
  end
  else if Tipo='CLAS' then
    Result := Copy(Texto,1,4)+'.'+Copy(Texto,5,2)+'.'+Copy(Texto,7,4)
  else if Tipo='PIS' then
    Result := Copy(Texto,1,3)+'.'+
              Copy(Texto,4,5)+'.'+
              Copy(Texto,9,2)+'.'+
              Copy(Texto,11,2)
  else if Tipo='CFOP' then
    Result := Copy(Texto,1,1)+'.'+Copy(Texto,2,3)
  else if Tipo='CEI' then
    Result := Copy(Texto,1,2)+'.'+
              Copy(Texto,3,3)+'.'+
              Copy(Texto,6,5)+'/'+
              Copy(Texto,11,2)
  else if Tipo='CAT' then
    Result := Copy(Texto,1,10)+'-'+Copy(Texto,11,1)+'/'+Copy(Texto,12,2)
  else if Tipo='LOC' then
    Result := Copy(Texto,1,2)+'.'+Copy(Texto,3,3)+'.'+Copy(Texto,6,3)
  else if Tipo='CNAE' then
    Result := Copy(Texto,1,4)+'-'+Copy(Texto,5,1)+'/'+Copy(Texto,6,2)
  else if Tipo='FONE' then
    Result := '('+AllTrim(Copy(Texto,1,6))+') '+AllTrim(Copy(Texto,7,20))
  else if Tipo='FONE2' then
    Result := '('+RetornaSoNumero(Copy(Texto,1,6))+')'+RetornaSoNumero(Copy(Texto,7,20))
  else if Tipo='PLACA' then
    Result := Copy(Texto,1,3)+'-'+Copy(Texto,4,7)
  else if Tipo='NCM' then
    Result := Copy(Texto,1,4)+'.'+Copy(Texto,5,2)+'.'+Copy(Texto,7,2)
  else if Tipo='NBM' then
    Result := Copy(Texto,1,4)+'.'+Copy(Texto,5,2)+'.'+Copy(Texto,7,4)
  else
    Result := Texto;
end;

function InsereStr (sTexto: String; iTam: Integer; sChar: Char; sD_ou_E_ou_C: Char): string;
var
  i: Integer;
begin
  result:=sTexto;
  if sD_ou_E_ou_C='E' then
  begin
    for i := (Length(sTexto)+1) to iTam do
        result:=sChar+result
  end
  else if sD_ou_E_ou_C='D' then
  begin
    for i := (Length(sTexto)+1) to iTam do
        result:=result+sChar;
  end
  else if sD_ou_E_ou_C='C' then
    Result:=CentralizaStr(sTexto,iTam,sChar);
end;

function CentralizaStr(sTexto:string;iTam:integer;sChar:char): string;
var i:integer;
begin
  result:=sTexto;
  if Length(sTexto) >= iTam then
    exit;
  if ((iTam-Length(sTexto)) mod 2) = 0 then
    i:=iTam-Length(sTexto)
  else
    i:=iTam-(Length(sTexto)+1);
  if i=0 then
    exit;
  result:=StringOfChar(sChar,StrToInt(FloatToStr(i/2)))+sTexto+StringOfChar(sChar,StrToInt(FloatToStr(i/2)));
  if iTam>Length(result) then
    result:=result+StringOfChar(sChar,iTam-Length(result));
end;

procedure ShowTitlebar(AForm: TForm; AShow: boolean);
var
  style: longint;
begin
  with AForm do
  begin
    if BorderStyle = bsNone then
      exit;
    style := GetWindowLong(Handle, GWL_STYLE);
    if AShow then
    begin
      if (style and WS_CAPTION) = WS_CAPTION then
        exit;
      case BorderStyle of
        bsSingle, bsSizeable:
          SetWindowLong(Handle, GWL_STYLE, style or WS_CAPTION or WS_BORDER);
        bsDialog:
          SetWindowLong(Handle, GWL_STYLE,
            style or WS_CAPTION or DS_MODALFRAME or WS_DLGFRAME);
      end;
    end
    else
    begin
      if (style and WS_CAPTION) = 0 then
        exit;
      case BorderStyle of
        bsSingle, bsSizeable:
          SetWindowLong(Handle, GWL_STYLE,
            style and (not(WS_CAPTION)) or WS_BORDER);
        bsDialog:
          SetWindowLong(Handle, GWL_STYLE,
            style and (not(WS_CAPTION)) or DS_MODALFRAME or WS_DLGFRAME);
      end;
    end;
    SetWindowPos(Handle, 0, 0, 0, 0, 0,
      SWP_NOMOVE or SWP_NOSIZE or SWP_NOZORDER or SWP_FRAMECHANGED
        or SWP_NOSENDCHANGING);
  end;
end;

function getEmpresaInfo(Connec:TDatabase; iEmpresa: Integer; sCampo: String): Variant;
begin
  with TQuery.Create(nil) do
  begin
    try
      DatabaseName := Connec.DatabaseName;
      SQL.Text := 'select empresa.* from empresa where noempresa = :noempresa ';
      ParamByName('noempresa').AsInteger := iEmpresa;
      Open;
      Result := FieldByName(sCampo).Value;
    finally
      Free;
    end;
  end;
end;

function getEmpresaInfo(Connec:TIBCConnection; iEmpresa: Integer; sCampo: String): Variant;
begin
  with TIBCQuery.Create(nil) do
  begin
    try
      Connection  := Connec;
      Transaction := Connec.DefaultTransaction;
      SQL.Text := 'select empresa.* from empresa where noempresa = :noempresa ';
      ParamByName('noempresa').AsInteger := iEmpresa;
      Open;
      Result := FieldByName(sCampo).Value;
    finally
      Free;
    end;
  end;
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

{======================================================================================================}
{ Captura os relacionamento de um Campo com uma Tabela Definida                                        }
{ Retorna : Vetor [0..9] com campos retorno da Tabela Definida                                         }
{ Exemplo:                                                                                             }
{ Var                                                                                                  }
{  VETRET:TVect10;                                                                                     }
{  begin                                                                                               }
{    VETRET:=RETORNA_VLR(conexao,'NOME_TABELA', ['CODIGO','NOME'], ['1','LOCALX'], ['CODIGO'])         }
{  ou                                                                                                  }
{    if RETORNA_VLR(conexao,'NOME_TABELA', ['CODIGO','NOME'], ['1','SYSTEM'], ['CODIGO'])[0]<>'' then    }
{       showmessage('Código já cadastrado');                                                           }
{                                                                                                      }
{  Parâmetros                                                                                          }
{         connec       = conexao = conexão atual onde econtra-se a tabela                               }
{         tbselect    = 'NOME_TABELA' = nome da tabela onde deseja executar a pesquisa                 }
{         cmpCriterio = ['CODIGO','NOME'] = nome dos campos para critério na pesquisa                  }
{         cmpValores  = ['1','SYSTEM']    = valores dos campos critérios                               }
{         cmpRetorno  = ['CODIGO'] = define quantos e quais os campos que serão retornados na função   }
{                                                                                                      }
{======================================================================================================}
function RETORNA_VLR(Connec: TIBCConnection; tbSelect: String; cmpCriterio,cmpValores,cmpRetorno: Array of Variant): TVect10Variant;
var
  i       : Integer;
  qry     : TIBCQuery;
  sRetorno: String;
begin
  qry := TIBCQuery.Create(nil);
  try
    if cmpCriterio[0]='' then
    begin
      result[0]:='';
      exit;
    end;
    // Validar Valores dos Campos Critérios (Não pode aceitar vazio)
    for i:=0 to high(cmpValores) do
    begin
      if AllTrim(cmpValores[i])='' then
      begin
        result[0]:='';
        exit;
      end;
    end;
    // Monta Campos de Retorno
    if not (high(cmpRetorno)>=0) then
      sRetorno:='*'
    else
    begin
      for i:=0 to high(cmpRetorno) do
      begin
        sRetorno := sRetorno + cmpRetorno[i];
        if i <> high(cmpRetorno) then sRetorno := sRetorno + ',';
      end;
    end;
    QRY.Connection := Connec;
    qry.SQL.Add('SELECT '+sRetorno+' FROM '+tbSelect);
    // Critérios
    for i:=0 to high(cmpCriterio) do
    begin
      if i=0 then
        qry.SQL.Add(' WHERE '+cmpCriterio[0]+'=:X0')
      else
        qry.SQL.Add(' AND '+cmpCriterio[i]+'=:X'+inttostr(i));
    end;
    // Valores
    for i:=0 to high(cmpCriterio) do
    begin
      if VarIsType(cmpValores[i],varDate) then
        QRY.ParamByName('X'+inttostr(i)).AsSQLTimeStamp:= VarToSQLTimeStamp(cmpValores[i])
      else
        QRY.ParamByName('X'+inttostr(i)).Value:= cmpValores[i];
    end;
    // Consultar
    qry.Open;
    if qry.RecordCount > 0 then
    begin
      for i:=0 to high(cmpRetorno) do
      begin
        if i>9 then break; {maior que o tamano do vetor de retorno}
        result[i]:=qry.FindField(cmpRetorno[i]).Value;
      end;
    end
    else
      result[0]:='';
  finally
    qry.Free;
  end;
end;

function DownloadFile(Source, Dest: string): Boolean;
begin
  try
    Result:= UrlDownloadToFile(nil, PChar(source),PChar(Dest), 0, nil) = 0;
  except
    Result:= False;
  end;
end;


function getDateAfterHour(Data: TDateTime; HoraReducao: String='0:00:00'): TDateTime;
begin

  if (Time < StrToTimeDef(HoraReducao,0)) then
  begin
    if DayOfWeek(Data-1) in [1,7] then
      Result := getDateAfterHour(Data-1, HoraReducao)
    else
      Result := Data - 1;
  end
  else
    Result := Data;
end;

procedure CopyToClipboard(sStr: String);
begin
  with TClipboard.Create do
    AsText := Trim(sStr);
end;

//----------- SelectDir --------------------------------
function BrowseCallbackProc(hwnd: HWND; uMsg: UINT; lParam, lpData: LPARAM):
  Integer; stdcall;
var
  Path: array[0..MAX_PATH] of Char;
begin
  case uMsg of
    BFFM_INITIALIZED:
      begin
        SendMessage(hwnd, BFFM_SETSELECTION, 1, lpData);
        SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, lpData);
      end;
    BFFM_SELCHANGED:
      begin
        if SHGetPathFromIDList(Pointer(lParam), Path) then
          SendMessage(hwnd, BFFM_SETSTATUSTEXT, 0, Integer(@Path));
      end;
  end;
  Result := 0;
end;

function SelectDir(hOwner: THandle; const Caption, InitialDir: string;
  const Root: WideString; ShowStatus: Boolean; out Directory: string): Boolean;
var
  BrowseInfo: TBrowseInfo;
  Buffer: PChar;
  RootItemIDList, ItemIDList: PItemIDList;
  ShellMalloc: IMalloc;
  IDesktopFolder: IShellFolder;
  Eaten, Flags: Cardinal;//LongInt;//LongWord;
  Windows: ^Integer;//Pointer;
  Path: string;

begin
  Result := False;
  Directory := '';
  Path := InitialDir;
  if (Length(Path) > 0) and (Path[Length(Path)] = '\') then
    Delete(Path, Length(Path), 1);

  FillChar(BrowseInfo, SizeOf(BrowseInfo), 0);
  if (ShGetMalloc(ShellMalloc) = S_OK) and (ShellMalloc <> nil) then
  begin
    Buffer := ShellMalloc.Alloc(MAX_PATH);
    try
      SHGetDesktopFolder(IDesktopFolder);
      IDesktopFolder.ParseDisplayName(hOwner, nil, PWideChar(Root), Eaten, RootItemIDList, Flags);
      with BrowseInfo do
      begin
        hwndOwner := hOwner;
        pidlRoot := RootItemIDList;
        pszDisplayName := Buffer;
        lpszTitle := PChar(Caption);
        ulFlags := BIF_RETURNONLYFSDIRS;
        if ShowStatus then
          ulFlags := ulFlags or BIF_STATUSTEXT;
        lParam := Integer(PChar(Path));
        lpfn := BrowseCallbackProc;
        iImage := 0;
      end;

      // Make the browser dialog modal.
      Windows := DisableTaskWindows(hOwner);
      try
        ItemIDList := ShBrowseForFolder(BrowseInfo);
      finally
        EnableTaskWindows(Windows);
      end;

      Result := ItemIDList <> nil;
      if Result then
      begin
        ShGetPathFromIDList(ItemIDList, Buffer);
        ShellMalloc.Free(ItemIDList);
        Directory := Buffer;
      end;
    finally
      ShellMalloc.Free(Buffer);
    end;
  end;
end;

function EncodeFileBase64(const FileName: string): AnsiString;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    try
      Stream.LoadFromFile(Filename);
      Result := EncodeBase64(Stream.Memory, Stream.Size);
    except
      on E: Exception do
      begin
        raise Exception.Create(PWideChar(WideString('Erro ao codificar arquivo '''+FileName+''': '+E.Message)));
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure DecodeFileBase64(const EncodedBase64: AnsiString; const FileName: string);
var
  Stream: TBytesStream;
begin
  try
    Stream := TBytesStream.Create(DecodeBase64(EncodedBase64));
    try
      Stream.SaveToFile(Filename);
    finally
      Stream.Free;
    end;
  except
    on E: Exception do
    begin
      raise Exception.Create(PWideChar(WideString('Erro ao criar arquivo '''+FileName+''': '+E.Message)));
    end;
  end;
end;

procedure TextToFile(sText, sFileName: String);
begin
  with TStringList.Create do
  begin
    Text := sText;
    SaveToFile(sFileName);
    Free;
  end;
end;

function ExecutarPrograma(const sFileName: String; sFormName: String=''): Boolean;
var
  hWnd: THandle;
begin
  hWnd := FindWindow(PWideChar(WideString(sFormName)), nil);
  if hWnd > 0 then
    ShowWindow(hWnd, SW_RESTORE)
  else
    WinExec(PAnsiChar(AnsiString(sFileName)), SW_SHOW);
end;

function ExecutarPrograma(const FileName: String; WindowState: Word): Boolean;
var
  SUInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CmdLine: String;
begin
  { Coloca o nome do arquivo entre aspas. Isto é necessário devido aos espaços contidos em nomes longos }
  CmdLine := '"' + Filename + '" ';
  FillChar(SUInfo, SizeOf(SUInfo), #0);
  with SUInfo do
  begin
    cb := SizeOf(SUInfo);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := WindowState;
  end;
  Result := CreateProcess(nil, PChar(CmdLine), nil, nil, False,
                          CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil,
                          PChar(ExtractFilePath(Filename)), SUInfo, ProcInfo);
  if Result then
  begin
    CloseHandle(ProcInfo.hProcess);
    CloseHandle(ProcInfo.hThread);
  end
  else
    MessageBoxW(0, 'Programa não encontrado!'#13'Contate o administrador do sistema para maiores informações.', 'Aviso', MB_OK + MB_ICONWARNING + MB_TOPMOST);
end;

// ExecAndWait('c:\windows\calc.exe','',sw_show);
function ExecAndWait(const FileName, Params: String; const WindowState: Word): Boolean;
var
  SUInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CmdLine: String;
begin
  if not FileExists(FileName) then
    MessageBoxW(0, 'Programa não encontrado!'#13'Contate o administrador do sistema para maiores informações.', 'Aviso', MB_OK + MB_ICONWARNING + MB_TOPMOST);

  { Coloca o nome do arquivo entre aspas. Isto é necessário devido aos espaços contidos em nomes longos }
  CmdLine := '"' + Filename + '" ' + Params;
  FillChar(SUInfo, SizeOf(SUInfo), #0);
  with SUInfo do
  begin
    cb := SizeOf(SUInfo);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := WindowState;
  end;
  Result := CreateProcess(nil, PChar(CmdLine), nil, nil, False,
                          CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil,
                          PChar(ExtractFilePath(Filename)), SUInfo, ProcInfo);
  { Aguarda até ser finalizado }
  if Result then
  begin
    WaitForSingleObject(ProcInfo.hProcess, INFINITE);
    { Libera os Handles }
    CloseHandle(ProcInfo.hProcess);
    CloseHandle(ProcInfo.hThread);
  end;
end;

procedure ValidarMenu(oMenu: TMainMenu);
var
  i: Integer;
begin
  for i := 0 to oMenu.Owner.ComponentCount- 1 do
  begin
    if (oMenu.Owner.Components[i] is TMenuItem){ and (oMenu.Owner.Components[i].Owner = oMenu)} then
      TMenuItem(oMenu.Owner.Components[i]).Enabled := (Assigned(TMenuItem(oMenu.Owner.Components[i]).OnClick)) or
                                                      (TMenuItem(oMenu.Owner.Components[i]).Count > 0);
  end;

end;

function UpperLowerCase(x: String): String;
var
  y:string;
  i:integer;
begin
  x := StringReplace(x, '_', ' ', [rfReplaceAll]);
  x := LowerCaseRGF(x);
  Result := '';
  for i := 1 to Length(x) do
  begin
    if (i=1) or ((i>1) and (x[i-1]=' ')) then
      Result := Result + UpperCaseRGF(x[i])
    else
      Result := Result + x[i];
  end;
end;

function GetTopLeft(AForm: TForm; AControl: TControl): TPoint;
begin
  Result := AForm.ScreenToClient(AControl.ClientToScreen(Point(AControl.Left, AControl.Top)));
end;

function RetornaTopLeft(Controle: TControl; MainFormName: String): TPoint;
var
  c: TControl;
begin
  c := Controle;
  Result.Y := c.Top;
  Result.X := c.Left;

  while ((not (c.Name = MainFormName)) and (not (c.Parent = nil))) do
  begin
    c := c.Parent;
    Inc(Result.Y, c.Top);
    Inc(Result.X, c.Left);
  end;
  if c is TForm then
  begin
    Inc(Result.Y, c.Height-c.ClientHeight);
//    Inc(Result.X, c.Width-c.ClientWidth);
  end;

end;

function RetornaTopLeft(Controle: TControl): TPoint;
var
  c: TControl;
begin
  c := Controle;
  Result.Y := c.Top;
  Result.X := c.Left;

  while (not (c.Parent is TForm)) do
  begin
    c := c.Parent;
    Inc(Result.Y, c.Top);
    Inc(Result.X, c.Left);
  end;
  if c.Parent is TForm then
  begin
    Inc(Result.Y, c.Height-c.ClientHeight);
    Inc(Result.X, c.Width-c.ClientWidth);
  end;

end;

{function getFieldsValue(qrLocal: TIBCQuery): TArecField;
var
  i: Integer;
  Aux: TArecField;
begin
  SetLength(Aux, qrLocal.FieldCount);
  for i := 0 to qrLocal.FieldCount - 1 do
  begin
    Aux[i].FieldName := qrLocal.Fields[i].FieldName;
    Aux[i].Value     := qrLocal.Fields[i].Value;
  end;
  Result := Aux;
end;}

function GetMyDocuments: string;
var
  r: Bool;
  path: array[0..Max_Path] of Char;
begin
  r := ShGetSpecialFolderPath(0, path, CSIDL_Personal, False) ;
  if not r then raise Exception.Create('Não foi possível encontrar a pasta Meus Documentos.') ;
  Result := Path;
end;

procedure SaveImageListToResFile(sFile: String; ImgList: TImageList);
begin
  WriteComponentResFile(sFile, ImgList);
end;

procedure ReadImageListFromResFile(sFile: String; const ImgList: TImageList);
begin
  ReadComponentResFile(sFile, ImgList);
end;

procedure GetImageFromResFile(sFile: String; Idx: Integer; const Img: TImage);
var
  ImageList: TImageList;
begin
  ImageList := TImageList.Create(nil) ;
  try
    ReadComponentResFile(sFile, ImageList);
//    ImageList.GetBitmap(Idx, Img.Picture.Bitmap);
    ImageList.GetBitmap(Idx, Img.Picture.Bitmap);
    Img.Picture.Bitmap.Transparent := True;
    Img.Picture.Bitmap.TransparentMode := tmFixed;
    Img.Picture.Bitmap.TransparentColor := Img.Picture.Bitmap.Canvas.Pixels[1,1];
  finally
    FreeAndNil(ImageList);
  end;
end;

procedure MenuToTreeView(AMenu: TMainMenu; ATree: TTreeView);
  procedure AddItem(AItem: TMenuItem; p: TTreeNode);
  var
    i: Integer;
    n: TTreeNode;
  begin
    if (AItem.Caption <> '-') then
      if (AItem.Visible) and (AItem.Enabled) then
      begin
        n := ATree.Items.AddChildObject(p, StripHotkey(AItem.Caption),AItem);
        for i := 0 to AItem.Count - 1 do
          AddItem(AItem.Items[i], n);
      end;
  end;

var
  i: Integer;
begin
  ATree.Items.Clear;
  ATree.Items.BeginUpdate;
  try
    for i := 0 to AMenu.Items.Count - 1 do
      AddItem(AMenu.Items[i], nil);
  finally
    ATree.Items.EndUpdate;
  end;
end;

function getStreamFromString(sStr: AnsiString): TStream;
//var
//  aStream: TStream;
begin
  Result := TMemoryStream.Create;
  Result.WriteBuffer(Pointer(sStr)^, Length(sStr));
  Result.Seek(0, soFromBeginning);
//  Result := aStream;
end;

function SaveHtmlWebBrowser(WebBrowser: TWebBrowser): String;
var
  PersistStreamInit: IPersistStreamInit;
  StreamAdapter: IStream;
  lStream: TStream;
  slHtml: TStringList;
begin
  Result := '';
  lStream := TMemoryStream.Create;
  slHtml := TStringList.Create;
  try
    if WebBrowser.Document.QueryInterface(IPersistStreamInit, PersistStreamInit) = S_OK then
    begin
      StreamAdapter := TStreamAdapter.Create(lStream);
      PersistStreamInit.Save(StreamAdapter, True);
      slHtml.LoadFromStream(lStream);
      Result := slHtml.Text;
    end;
  finally
    FreeAndNil(slHtml);
  end;

end;

procedure LoadHtmlWebBrowser(WebBrowser: TWebBrowser; sHTML: String);
var
  PersistStreamInit: IPersistStreamInit;
  lStream : TStream;
  StreamAdapter: IStream;
begin
  WebBrowser.Navigate('about:blank');
  repeat
    Application.ProcessMessages;
    Sleep(0);
  until WebBrowser.ReadyState = READYSTATE_COMPLETE;

  if WebBrowser.Document.QueryInterface(IPersistStreamInit, PersistStreamInit) = S_OK then
  begin
    lStream := getStreamFromString(sHTML);
    StreamAdapter:= TStreamAdapter.Create(lStream);
    try
      PersistStreamInit.Load(StreamAdapter);
    finally
      StreamAdapter := nil;
      FreeAndNil(lStream);
    end;
  end;
end;

procedure AbrirHtmlWebBrowser(WebBrowser: TWebBrowser; slHTML: TStringList);
var
  aStream : TMemoryStream;
  HtmlStr : AnsiString;
begin
  HtmlStr := slHTML.Text;
  WebBrowser.Navigate('about:blank');
  repeat
    Application.ProcessMessages;
    Sleep(0);
  until WebBrowser.ReadyState = READYSTATE_COMPLETE;

  if Assigned(WebBrowser.Document) then
  begin
    aStream := TMemoryStream.Create;
    try
      aStream.WriteBuffer(Pointer(HTMLStr)^, Length(HTMLStr));
      aStream.Seek(0, soFromBeginning);
//      LoadStreamWebBrowser(WebBrowser, aStream);
      (WebBrowser.Document as IPersistStreamInit).Load(TStreamAdapter.Create(aStream));
    finally
      FreeAndNil(aStream);
    end;
  end;
//    HTMLWindow2 := (WebBrowser1.Document as IHTMLDocument2).parentWindow;
end;

procedure LoadStreamWebBrowser(WebBrowser: TWebBrowser; Stream: TStream);
var
  PersistStreamInit: IPersistStreamInit;
  StreamAdapter: IStream;
  MemoryStream: TMemoryStream;
begin
  {Load empty HTML document into Webbrowser to make "Document" a valid HTML document}
  WebBrowser.Navigate('about:blank');
  {wait until finished loading}
  repeat
    Application.ProcessMessages;
    Sleep(0);
  until WebBrowser.ReadyState = READYSTATE_COMPLETE;
  {Get IPersistStreamInit - Interface}
  if WebBrowser.Document.QueryInterface(IPersistStreamInit, PersistStreamInit) = S_OK then
  begin
    {Clear document}
    if PersistStreamInit.InitNew = S_OK then
    begin
      {Make local copy of the contents of Stream if you want to use Stream directly,
                        you have to consider, that StreamAdapter will destroy it automatically}
      MemoryStream := TMemoryStream.Create;
      try
        MemoryStream.CopyFrom(Stream, 0);
        MemoryStream.Position := 0;
      except
        MemoryStream.Free;
        raise;
      end;
      {Use Stream-Adapter to get IStream Interface to our stream}
      StreamAdapter := TStreamAdapter.Create(MemoryStream, soOwned);
      {Load data from Stream into WebBrowser}
      PersistStreamInit.Load(StreamAdapter);
    end;
  end;
end;

//{$IFDEF VER210} // Delphi 2010
function AddImageToGrid(const Rect: TRect; Grid: TDBAltGrid; imgList: TImageList; imgIndex: Integer; bTransparent: Boolean=True; cTransparentColor: TColor=clNone): TRect;
var
  Bitmap   : TBitmap;
  fixRect  : TRect;
  bmpWidth : integer;
begin
  fixRect := Rect;
  with Grid do
  begin
    Bitmap := TBitmap.Create;
    try
      //grab the image from the ImageList
      imgList.GetBitmap(imgIndex,Bitmap);

      Bitmap.Transparent      := bTransparent;
      Bitmap.TransparentMode  := tmAuto;

      if cTransparentColor = clNone then
        Bitmap.TransparentColor := Bitmap.Canvas.Pixels[1,1]
      else
        Bitmap.TransparentColor := cTransparentColor;

      //Fix the bitmap dimensions
//      bmpWidth := (Rect.Bottom - Rect.Top);
//      fixRect.Right := Rect.Left + bmpWidth;
      fixRect.Right  := Rect.Left + Bitmap.Width;
      fixRect.Top    := (Rect.Bottom - Bitmap.Height) div 2;
      fixRect.Bottom := Rect.Top + Bitmap.Height;
      //draw the bitmap
      Canvas.Draw(fixRect.Left, fixRect.Top, Bitmap);
//      Canvas.StretchDraw(fixRect,Bitmap);
    finally
      Bitmap.Free;
    end;

    // reset the output rectangle,
    // add space for the graphics
    fixRect := Rect;
    fixRect.Left := fixRect.Left + bmpWidth;
    Result := fixRect;
  end;
end;
//{$ENDIF}

function AddImageToGrid(const Rect: TRect; Grid: TDBGrid; imgList: TImageList; imgIndex: Integer; bTransparent: Boolean=True; iX_Transp: Integer=0; iY_Transp: Integer=0; bShowText: Boolean=False): TRect;
var
  Bitmap   : TBitmap;
  fixRect  : TRect;
  bmpWidth : integer;
  intX, intY: Integer;
begin
  fixRect := Rect;
  with Grid do
  begin
    Canvas.FillRect(Rect);

    Bitmap := TBitmap.Create;
    try
      //grab the image from the ImageList
      imgList.GetBitmap(imgIndex,Bitmap);

      Bitmap.Transparent      := bTransparent;
      Bitmap.TransparentMode  := tmAuto;
      Bitmap.TransparentColor := Bitmap.Canvas.Pixels[iX_Transp, iY_Transp];

      if not bShowText then
      begin
        intX := ((rect.Right - rect.Left) div 2) -(Bitmap.Width div 2);
        intY := ((rect.Bottom - rect.Top) div 2) - (Bitmap.Height div 2);

        Canvas.Draw(rect.Left + intX, rect.Top + intY, Bitmap);
      end
      else
      begin
        bmpWidth := (Rect.Bottom - Rect.Top);
        fixRect.Right := Rect.Left + bmpWidth;
        //draw the bitmap
        Canvas.StretchDraw(fixRect,bitmap);
      end;

//      //Fix the bitmap dimensions
//      bmpWidth := (Rect.Bottom - Rect.Top);
//      fixRect.Right := Rect.Left + bmpWidth;
//      //draw the bitmap
//      Canvas.StretchDraw(fixRect,Bitmap);
    finally
      Bitmap.Free;
    end;

    // reset the output rectangle,
    // add space for the graphics
    fixRect := Rect;
    if bShowText then
    begin
      fixRect.Left := fixRect.Left + bmpWidth;
//      DefaultDrawColumnCell(fixRect, DataCol, Column, State);
    end;
    Result := fixRect;
  end;


//-----------------------------------------------
//fixRect := Rect;
//
//  // customizing the 'LastName' field
//   imgIndex := 2;
//    bitmap := TBitmap.Create;
//    try
//      //grab the image from the ImageList
//      //(using the "Salary" field's value)
//      ImageList1.GetBitmap(imgIndex,bitmap);
//      //Fix the bitmap dimensions
//      bmpWidth := (Rect.Bottom - Rect.Top);
//      fixRect.Right := Rect.Left + bmpWidth;
//      //draw the bitmap
//      DBGrid1.Canvas.StretchDraw(fixRect,bitmap);
//    finally
//      bitmap.Free;
//    end;
//
//    // reset the output rectangle,
//    // add space for the graphics
//    fixRect := Rect;
//    fixRect.Left := fixRect.Left + bmpWidth;
//  end;
//
//  //draw default text (fixed position)
//  DBGrid1.DefaultDrawColumnCell(
//     fixRect,
//     DataCol,
//     Column,
//     State);
//-----------------------------------------------
//var
//  bmpImage: tBitmap;
//  intX, intY: Integer;
//begin
//  if DataSource1.DataSet.RecNo > 0 then
//  begin
//    if column.FieldName = 'COUNTRY' then
//      with DBGrid1.Canvas do
//      begin
//        fillRect(rect);
//        bmpImage := tBitmap.Create;
//        try
//          if DBGrid1.DataSource.DataSet.FieldByName('COUNTRY').AsInteger = 0 then
//            bmpImage.Assign(image0.Picture.Bitmap);
//
//          if DBGrid1.DataSource.DataSet.FieldByName('COUNTRY').AsInteger = 1 then
//            bmpImage.Assign(image1.Picture.Bitmap);
//
//          intX := ((rect.Right - rect.Left) div 2) -
//            (bmpImage.Width div 2);
//
//          intY := ((rect.Bottom - rect.Top) div 2) -
//            (bmpImage.Height div 2);
//          draw(rect.Left + intX, rect.Top + intY, bmpImage);
//        finally
//          bmpimage.Free;
//        end;
//      end;
//  end;
end;

procedure ImageFromListImg(Img: TImage; imgList: TImageList; imgIndex: Integer);
begin
  try
    //grab the image from the ImageList
    imgList.GetIcon(imgIndex,Img.Picture.Icon);
//    Img.Picture.Bitmap.Transparent      := True;
//    Img.Picture.Bitmap.TransparentMode  := tmFixed;
//    Img.Picture.Bitmap.TransparentColor := Img.Picture.Bitmap.Canvas.Pixels[1,1];

  finally
  end;
end;

procedure CapturaTela(sFileName: String; iQualidade: Integer=100);
var
  dc : HDC;
  jpg: TJPEGImage;
  bmp: TBitmap;
begin
  try
    {TJPEGImage necessario declara unit JPeg na delcaração uses...}
    jpg := TJPEGImage.Create;
    bmp := TBitmap.Create;

    {Monitor e uma class i esta disponivel como
     property delcarado em  TCustomForm e guarada
     informações sobre a sobre a configuração de resoluca como largura e altura}

    bmp.Width  := Screen.Width;
    bmp.Height := Screen.Height;

    {usando a função da Api Windows getdc com    valor zero , voce recuprea o descriptor do desktop do windows}
    dc := GetDC(0);

    {usando a função BitBlt da Api windows para gravar o descriptor da tela no canvas do Bitmap capiturando a Tela}
    BitBlt(bmp.Canvas.Handle,0,0,bmp.Width,bmp.Height,dc,0,0,SRCCOPY);

    jpg.CompressionQuality := iQualidade;
    { convertendo bitmap para Jpg qualidade exelente como acima 100% e de tamanho bem menor que o bitmap}

    jpg.Assign(bmp);

    if not DirectoryExists(ExtractFileDir(sFileName)) then
      ForceDirectories(ExtractFilePath(sFileName));

    jpg.SaveToFile(ChangeFileExt(sFileName,'.jpg'));
  finally
    DeleteDC(dc);
    bmp.Free;
    jpg.Free;
  end;
end;

function getParentStructure(const Component: TWinControl): String;
var
  AuxComp: TWinControl;
begin
  Result := '[ '+Component.Name+' ] ';
  AuxComp:= Component.Parent;
  while Assigned(AuxComp) do
  begin
    Result  := AuxComp.Name +' -> '+ Result;
    AuxComp := AuxComp.Parent;
  end;
end;

function getParentRGF(const Component: TWinControl): TWinControl;
var
  AuxComp: TWinControl;
begin
  Result  := Component;
  AuxComp := Component.Parent;
  while Assigned(AuxComp) do
  begin
    Result := AuxComp;
    AuxComp := AuxComp.Parent;
  end;
end;

function getParentRGF(const Component: TWinControl; sClassname: String): TWinControl; overload;
var
  AuxComp: TWinControl;
begin
  Result := nil;
  if Component.ClassName = sClassname then
  begin
    Result  := Component;
    Exit;
  end;

  AuxComp := Component.Parent;
  while Assigned(AuxComp) do
  begin
    if AuxComp.ClassName = sClassname then
    begin
      Result  := AuxComp;
      Exit;
    end;
    AuxComp := AuxComp.Parent;
  end;
end;

procedure FormSempreVisivel(Form: TForm);
begin
  SetWindowPos(Form.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE or SWP_NOMOVE or SWP_NOSIZE);
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

function UpperCaseRGF(sStr: String; bSemAcento: Boolean=False): String;
var i, j: integer;
begin
  sStr := UpperCase(sStr);
  for i := 1 to Length(sStr) do
  begin
    for j := 1 to High(MinusculaAcentuadaF) do
    begin
      if sStr[i] = MinusculaAcentuadaF[j] then
      begin
        if bSemAcento then
          sStr[i] := MaiusculaSemAcentoF[j]
        else
          sStr[i] := MaiusculaAcentuadaF[j];
        Break;
      end
      else if sStr[i] = MaiusculaAcentuadaF[j] then
      begin
        if bSemAcento then
          sStr[i] := MaiusculaSemAcentoF[j];
        Break;
      end;
    end;
  end;
  Result := sStr;
end;

function LowerCaseRGF(sStr: String; bSemAcento: Boolean=False): String;
var i, j: integer;
begin
  sStr := LowerCase(sStr);
  for i := 1 to Length(sStr) do
  begin
    for j := 1 to High(MaiusculaAcentuadaF) do
    begin
      if sStr[i] = MaiusculaAcentuadaF[j] then
      begin
        if bSemAcento then
          sStr[i] := MinusculaSemAcentoF[j]
        else
          sStr[i] := MinusculaAcentuadaF[j];
        Break;
      end
      else if sStr[i] = MinusculaAcentuadaF[j] then
      begin
        if bSemAcento then
          sStr[i] := MinusculaSemAcentoF[j];
        Break;
      end
    end;
  end;
  Result := sStr;
end;

{function getDadosEndereco(Connec: TIBCConnection; sUsuario, sCep: String; var Endereco, CodBairro, Bairro, CodCidade, Cidade, Estado: String): Boolean;
var
  CEP: CEPServicePort;
  XMLCodigoIbge: TXMLDocument;
  Dados: ArrayOfstring;
  sAux, sCidadeAux : String;
  CodigoIbge: String;
  qrLocal: TIBCQuery;
  sProxy, sUser, sPassword: String;
begin
  Result := True;
  if sCep = '' then Exit;

  sProxy := Trim(GetINIDB(Connec, 'WEB', 'PROXY', '', sUsuario));
  if sProxy = '' then
    sProxy := Trim(GetINIDB(Connec, 'WEB', 'PROXY', '', ''));
  sUser := Trim(GetINIDB(Connec, 'WEB', 'USER', '', sUsuario));
  if sUser = '' then
    sUser := Trim(GetINIDB(Connec, 'WEB', 'USER', '', ''));
  sPassword := Trim(GetINIDB(Connec, 'WEB', 'PASSWORD', '', sUsuario));
  if sPassword = '' then
    sPassword := Trim(GetINIDB(Connec, 'WEB', 'PASSWORD', '', ''));

  CEP := GetCEPServicePort(False, '', nil, sProxy, sUser, sPassword);
  Endereco := '';
  CodBairro := '';
  Bairro := '';
  CodCidade := '';
  Cidade := '';
  Estado := '';
  CodigoIbge := '';
  SetLength(Dados,1);
  Dados[0] := CEP.obterLogradouroAuth(sCep, 'localx', 'nc123');

  if (Dados[0] = 'Usuário ou senha inválidos') then
  begin
    MessageBoxW(0, 'Usuário ou senha inválidos!'#13'Contate o administrador do sistema!', 'Aviso', MB_OK + MB_ICONWARNING + MB_TOPMOST);
    Exit;
  end
  else if (Pos(',',Dados[0]) = 0 ) then
  begin
    Result := False;
    Exit;
  end;

  sAux     := Dados[0];
  Endereco := UpperCaseRGF(Copy(sAux,1,Pos(',',sAux)-1));
  Endereco := Copy(Endereco, 1, 50);
  Delete(sAux,1,Pos(',',sAux)+1);
  Bairro   := UpperCaseRGF(Copy(sAux,1,Pos(',',sAux)-1));
  Bairro   := Copy(Bairro, 1, 30);
  Delete(sAux,1,Pos(',',sAux)+1);
  Cidade   := UpperCaseRGF(Copy(sAux,1,Pos(',',sAux)-1));
  Delete(sAux,1,Pos(',',sAux)+1);
  Estado   := UpperCaseRGF(Copy(sAux,1,Pos(',',sAux)-1));
  Delete(sAux,1,Pos(',',sAux)+1);
  CodigoIbge := UpperCaseRGF(sAux);

  if CodigoIbge = '' then
    CodigoIbge := '0';

  qrLocal := TIBCQuery.Create(nil);
  try
    try
      qrLocal.Connection  := Connec;
      qrLocal.Transaction := Connec.DefaultTransaction;

      //Código da Cidade no Cyber
//      qrLocal.Sql.Text := 'select NOCIDADE, CODIGOIBGE from CIDADE where NOME = :CIDADE and ESTADO = :ESTADO ';
//      qrLocal.ParamByName('CIDADE').AsString := Cidade;
//      qrLocal.ParamByName('ESTADO').AsString := Estado;
      qrLocal.Sql.Text := 'select NOCIDADE, CODIGOIBGE, upper((select resultado from semacentos(cidade.nome))) NOMECIDADE, ESTADO from CIDADE ';
      sCidadeAux := lowerCaseRGF(Cidade, True);
      sCidadeAux := upperCaseRGF(sCidadeAux, True);
      qrLocal.Open;
//      if qrLocal.RecordCount > 0 then
      if qrLocal.Locate('nomecidade;estado', VarArrayOf([sCidadeAux,Estado]), []) then
//
      begin
        CodCidade := qrLocal.FieldByName('NOCIDADE').AsString;
        if qrLocal.FieldByName('CODIGOIBGE').AsInteger = 0 then
        begin
          qrLocal.Close;
          qrLocal.Sql.Text := 'update CIDADE set CODIGOIBGE = :CODIGOIBGE where NOCIDADE = :NOCIDADE ';
          qrLocal.ParamByName('NOCIDADE').AsString   := CodCidade;
          qrLocal.ParamByName('CODIGOIBGE').AsString := CodigoIbge;
          qrLocal.Prepare;
          qrLocal.ExecSql;
        end;
      end
      else
      begin
        qrLocal.Close;
        qrLocal.Sql.Text := 'select MAX(NOCIDADE) from CIDADE ';
        qrLocal.Open;
        CodCidade := IntToStr(qrLocal.Fields[0].AsInteger + 1);

        qrLocal.Close;
        qrLocal.Sql.Text := 'insert into CIDADE(NOCIDADE, NOME, ESTADO, CODIGOIBGE) values(:NOCIDADE, :CIDADE, :ESTADO, :CODIGOIBGE) ';
        qrLocal.ParamByName('NOCIDADE').AsString   := CodCidade;
        qrLocal.ParamByName('CIDADE').AsString     := Cidade;
        qrLocal.ParamByName('ESTADO').AsString     := Estado;
        qrLocal.ParamByName('CODIGOIBGE').AsString := CodigoIbge;
        qrLocal.Prepare;
        qrLocal.ExecSql;
      end;

      //Código do Bairro no Cyber
      qrLocal.Close;
      qrLocal.Sql.Text := 'select NOBAIRRO from BAIRRO where NOME = :BAIRRO and NOCIDADE = :NOCIDADE ';
      qrLocal.ParamByName('BAIRRO').AsString   := Bairro;
      qrLocal.ParamByName('NOCIDADE').AsString := CodCidade;
      qrLocal.Open;
      if qrLocal.RecordCount > 0 then
        CodBairro := qrLocal.FieldByName('NOBAIRRO').AsString
      else
      begin
        qrLocal.Close;
        qrLocal.Sql.Text := 'select MAX(NOBAIRRO) from BAIRRO where NOCIDADE = :NOCIDADE ';
        qrLocal.ParamByName('NOCIDADE').AsString := CodCidade;
        qrLocal.Open;
        CodBairro := IntToStr(qrLocal.Fields[0].AsInteger + 1);

        qrLocal.Close;
        qrLocal.Sql.Text := 'insert into BAIRRO(NOCIDADE, NOBAIRRO, NOME) values(:NOCIDADE, :NOBAIRRO, :BAIRRO) ';
        qrLocal.ParamByName('NOCIDADE').AsString := CodCidade;
        qrLocal.ParamByName('NOBAIRRO').AsString := CodBairro;
        qrLocal.ParamByName('BAIRRO').AsString   := Bairro;
        qrLocal.Prepare;
        qrLocal.ExecSql;
      end;
      Connec.CommitRetaining;
    except
      Connec.RollbackRetaining;
      raise Exception.Create('Erro ao inserir nova(o) cidade/bairro. Contate o administrador do sistema.');
    end;
    qrLocal.Close;
  finally
    FreeAndNil(qrLocal);
  end;
end;                          }

//{$IFDEF VER210} // Delphi 2010
{procedure LimparCampos(aControl: TWinControl);
var
  i: Integer;
begin
  for i := 0 to aControl.ControlCount-1 do
  begin
    if aControl.Controls[i] is TEdit then
      TEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TRzEdit then
      TRzEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TNumberEdit then
      TNumberEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TRGFEdit then
      TRGFEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TMaskEdit then
      TMaskEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TComboBox then begin
      TComboBox(aControl.Controls[i]).Text := '';
      TComboBox(aControl.Controls[i]).ItemIndex := 0;
    end;
    if aControl.Controls[i] is TMemo then
      TMemo(aControl.Controls[i]).Lines.Clear;
    if aControl.Controls[i] is TDateEdit then
      TDateEdit(aControl.Controls[i]).Text := '';
    if aControl.Controls[i] is TCheckBox then
      TCheckBox(aControl.Controls[i]).State := cbUnchecked;
    if aControl.Controls[i] is TCurrencyEdit then
      TCurrencyEdit(aControl.Controls[i]).Text := '';
  end;
end;}
//{$ENDIF}

function GeraLOGArquivo(sFileName, sLog: String): String;
Var
  Arquivo : TextFile;
begin
  Result := '';
  if not DirectoryExists(ExtractFilePath_RGF(sFileName)) then
    ForceDirectories(ExtractFilePath_RGF(sFileName));

  AssignFile(Arquivo, sFileName); // Associa o arquivo
  try
    if FileExists(sFileName) then  // Verifica se existe
      Append(Arquivo) // Se existe, adiciona a informacao
    else
      ReWrite(Arquivo); // senão cria o arquivo
    if not FileExists(sFileName) then
    begin
      FileCreate(sFileName);
    end;
    Result := '['+ FormatDateTime('dd/mm/yyyy hh:nn:ss',now) + '] - ' + sLog;
    WriteLn(Arquivo,Result); // Grava a informação que eu quero com data e hora
  finally
    CloseFile(Arquivo); // Fecha o arquivo
  end;
end;

procedure AddLogMemo(sMsg: String; MemoLog: TMemo);
var
//  Arq: TextFile;
  sArquivo: String;
begin
  if sMsg = '' then Exit;

  sArquivo := ExtractFilePath(Application.ExeName) + 'log'+PathDelim+FormatDateTime('yyyy-mm-dd',Date)+'.log';
  sMsg := GeraLOGArquivo(sArquivo, sMsg);

  if MemoLog.Lines.Count = 0 then
    MemoLog.Lines.LoadFromFile(sArquivo)
  else
    MemoLog.Lines.Add(sMsg);

  MemoLog.Perform(EM_LINESCROLL,0,MemoLog.Lines.Count);
end;

{function GetComponentValue(wcComponent: TWinControl): String;
begin
  Result := '';
  try
    if Assigned(wcComponent) then
    begin
      if (wcComponent is TEdit) then
        Result := TEdit(wcComponent).Text;
      if (wcComponent is TDBEdit) then
        Result := TDBEdit(wcComponent).Text;
      if (wcComponent is TRzEdit) then
        Result := TRzEdit(wcComponent).Text;
      if (wcComponent is TRzDBEdit) then
        Result := TRzDBEdit(wcComponent).Text;
      if (wcComponent is TMaskEdit) then
        Result := TMaskEdit(wcComponent).Text;
      if (wcComponent is TButton) then
        Result := TButton(wcComponent).Caption;
      if (wcComponent is TMemo) then
        Result := TMemo(wcComponent).Lines.Text;
      if (wcComponent is TDBMemo) then
        Result := TDBMemo(wcComponent).Lines.Text;
    end;
  except
    on E: Exception do
      Result := 'Erro ao obter valor do campo ('+E.Message+')';
  end;
end;}

{-----------------------------------------------------------------------------
  Verifica se <APath> possui "PathDelim" no final. Retorna String com o Path
  já ajustado
 ---------------------------------------------------------------------------- }
function PathWithDelim( const APath: String ): String;
begin
  Result := Trim(APath) ;
  if Result <> '' then
     if RightStr(Result,1) <> PathDelim then   { Tem delimitador no final ? }
        Result := Result + PathDelim ;
end;

{-----------------------------------------------------------------------------
  Verifica se <APath> possui "PathDelim" no final. Retorna String SEM o
  DELIMITADOR de Path no final
 ---------------------------------------------------------------------------- }
function PathWithoutDelim(const APath: String): String;
Var
  Delimiters : AnsiString ;
begin
  Result := Trim(APath) ;

  Delimiters := PathDelim+'/\' ;
  while (Result <> '') and (pos(String(RightStr(Result,1)), String(Delimiters) ) > 0) do   { Tem delimitador no final ? }
     Result := copy(Result,1,Length(Result)-1)
end;

//uses ShellAPI, Windows, FileCtrl, Forms;
//FO_MOVE = $0001;
//FO_COPY = $0002;
//FO_DELETE = $0003;
//FO_RENAME = $0004;
function MoveCopiaDiretorios(pOperacao: Integer; pOrigem, pDestino: string):Boolean;
var
  recOperacao : TShFileOpStruct;
begin
  Result := False;
  if(pOrigem<>'')and(pDestino<>'')and(DirectoryExists(pOrigem))then
  begin
    pOrigem  := pOrigem+#0;
    pDestino := pDestino+#0;
    FillChar(recOperacao, Sizeof(TShFileOpStruct), 0);

    recOperacao.Wnd   := Application.Handle;
    recOperacao.wFunc := pOperacao;
    recOperacao.pFrom := PChar(pOrigem);
    recOperacao.pTo   := PChar(pDestino);
    recOperacao.fFlags:= FOF_ALLOWUNDO or FOF_SIMPLEPROGRESS or FOF_NOCONFIRMATION;

    result:= ShFileOperation(recOperacao)=0;
  end;
end;

function DigitoM9(S: String): Char;
var
  I: Integer;
  T1, T2, M1: Longint;

begin
  T1 := 0;
  I := Length(S);
  M1 := 2;
  while I > 0 do begin
    T1 := T1 + Trunc((strtofloat(S[I]) * M1));
    inc(M1);
    dec(I);
  end;
  T2 := T1 - (Trunc(T1 /11) * 11);
  T1 := 11 - T2;
  if (T1 = 0) or (T1 = 10) or (T1 = 11) then Result := '0'
  else Result := Chr(T1 + 48);
end;

function DigitoM11(S: String): Char;
var
  D, I: Integer;
  T, M: Longint;
begin
  T := 0;
  I := 1;
  M := Length(S)+1;
  while I <= Length(S) do begin
    T := T + (StrToInt(S[I]) * M);
    Dec(M);
    Inc(I);
  end;

  D := 11 - (T mod 11);
  if D = 10 then D := 0;
  if D = 11 then D := 1;
  Result := IntToStr(D)[1];
end;

//{$IFDEF VER210} // Delphi 2010
function ExportQRToPDF(Report: TQuickRep; sFileName: String; bOnlyPrepare: Boolean=False): Boolean;
var
  qf: TQRPDFDocumentFilter;
  i: Integer;
begin
  for i := 0 to Report.ComponentCount -1 do
  begin
    if (Report.Components[i] is TQRShape) and (TQRShape(Report.Components[i]).Shape = qrsRoundRect) then
    begin
      TQRShape(Report.Components[i]).Shape := qrsRectangle;
      TQRShape(Report.Components[i]).Pen.Width := 1;
    end;
  end;

  if not DirectoryExists(ExtractFilePath_RGF(sFileName)) then
    CreateDir(ExtractFilePath_RGF(sFileName));

  Report.Prepare;
  if not bOnlyPrepare then
  begin
     qf := TQRPDFDocumentFilter.Create(sFileName);
    try
      qf.CompressionOn := False;
      Report.QRPrinter.ExportToFilter(qf);
    finally
      qf.Free;
    end;
  end;
end;
//{$ENDIF}

function DiasAtraso(Connec: TIBCConnection; prDtIni, prDtFim: TDateTime): Integer;
var
  dtInicialRef, dtFinalRef: TDateTime;
begin
  Result := 0;
  dtInicialRef := prDtIni;
  dtFinalRef := prDtFim;
  while dtInicialRef <= dtFinalRef do begin
    if ((DayOfWeek(dtInicialRef) = 1) or (DayOfWeek(dtInicialRef) = 7) or isFeriado(Connec, dtInicialRef)) then
       dtInicialRef := dtInicialRef + 1
    else begin
      Result := trunc(prDtFim - dtInicialRef);
      dtInicialRef := dtFinalRef +1;
    end;
  end;
end;

function isFeriado(Connec: TIBCConnection; prData: TDateTime): Boolean;
var Ano, Mes, Dia : Word;
begin
  Result:= False;
  DecodeDate(prData, Ano, Mes, Dia);

  with TIBCQuery.Create(nil) do
  begin
    try
      Connection := Connec;
      Transaction:= Connec.DefaultTransaction;
      SQL.Text := 'select CHAVE from FERIADO where DIA = :DIA and MES = :MES';
      ParamByName('DIA').AsInteger := Dia;
      ParamByName('MES').AsInteger := Mes;
      Open;
      if RecordCount > 0 then
        Result := True;
      Close;
    finally
      Free;
    end;
  end;
end;

function AumentaXDiasUteis(Connec: TIBCConnection; DataIni: TDateTime; Dias: Integer):TDateTime;
var
  C : Integer;
  DataFinal : TDateTime;
begin
  DataFinal := DataIni;
  C := 0;
  while (C < Dias) do begin
    DataFinal := DataFinal + 1;
    if (dayofweek(DataFinal) > 1) and (dayofweek(DataFinal) < 7) and not(isFeriado(Connec, DataFinal))  then
      inc(C);
  end;
  Result := DataFinal;
end;

function BuscaLiberacao(Connec: TIBCConnection; CNPJ, Opcao: String):Boolean;
begin
  with TIBCQuery.Create(nil) do
  begin
    try
      Connection := Connec;
      Transaction:= Connec.DefaultTransaction;
      SQL.Text := 'select LIBERADO from LICENCALIB where CNPJ = :CNPJ and LIBERACAO = :LIBERACAO ';
      ParamByName('CNPJ').AsString := CNPJ;
      ParamByName('LIBERACAO').AsString := Opcao;
      Open;
      if FieldbyName('LIBERADO').AsBoolean then
         Result := True
      else
         Result := False;
      Close;
    finally
      Free;
    end;
  end;
end;

function bLocalX(Connec: TIBCConnection; CNPJ: String):Boolean;
begin
    Result := BuscaLiberacao(Connec, CNPJ, 'LOCAL X');
end;

function QtdeInteira(Qtde: Double; Conversao: Double; TipoConv: String): String;
begin
  if (Qtde - Trunc(Qtde) = 0) then
    if Conversao = 1 then
       Result := IntToStr(trunc(Qtde))
    else
      if Qtde <> 0 then
        if TipoConv = 'M' then
          Result := IntToStr(Trunc((Qtde/Conversao)+0.00001)) + ' / ' +
                    IntToStr(Trunc(Qtde - ((Trunc((Qtde/Conversao+0.00001))) * Conversao+0.00001)))
        else
          Result := IntToStr(Trunc(Qtde * Conversao+0.00001)) + ' / ' +
                    IntToStr(Trunc(Qtde - ((Trunc((Qtde*Conversao+0.00001)))/Conversao+0.00001)))
      else
        Result := '0 / 0'
  else
    if TipoConv = 'M' then
      Result := FloatToStrF((Qtde/Conversao), ffNumber, 13, 3)
    else
      Result := FloatToStrF((Qtde*Conversao), ffNumber, 13, 3);
end;

function CopyAtPos(S: String; Initial, Final: Integer): String;
var
  I: Integer;
begin
  Result := '';
  for I := Initial to Final do
    Result := Result + S[I];
end;

function DiasUteis(Connec: TIBCConnection; prDtIni, prDtFim: TDateTime): Integer;
var dtInicialRef, dtFinalRef: TDateTime;
begin
  Result := 0;
  dtInicialRef := prDtIni;
  dtFinalRef := prDtFim;
  while dtInicialRef <= dtFinalRef do
  begin
    if ((DayOfWeek(dtInicialRef) > 1) and (DayOfWeek(dtInicialRef) < 7) and not(isFeriado(Connec, dtInicialRef))) then
      Result := Result + 1;
    dtInicialRef := dtInicialRef + 1;
  end;
end;

function DiminuiXDiasUteis(Connec: TIBCConnection; DataIni: TDateTime; Dias: Integer):TDateTime;
var
  C : Integer;
  DataFinal : TDateTime;
begin
  DataFinal := DataIni;
  C := 0;
  while (C < Dias) do begin
    if (dayofweek(DataFinal) > 1) and (dayofweek(DataFinal) < 7) and not(isFeriado(Connec, DataFinal)) then
      inc(C);
    DataFinal := DataFinal - 1;
  end;
  Result := DataFinal;
end;

function isItemOfKit(Connec: TIBCConnection; prCodReduzido: String): Boolean;
begin
  with TIBCQuery.Create(nil) do
  begin
    try
      Result := False;
      Connection := Connec;
      Transaction:= Connec.DefaultTransaction;
      Sql.Text := 'select FIRST 1 KIT.CODIGOREDUZIDO '+
                  ' from KIT '+
                  ' inner join PRODUTO on PRODUTO.CODIGOREDUZIDO = KIT.CODIGOKIT '+
                  ' where KIT.CODIGOREDUZIDO = :CODIGO ' +
                  ' and PRODUTO.TIPO = ''K'' ';
      ParamByName('CODIGO').AsString := prCodReduzido;
      Open;
      if not Fields[0].IsNull then
         Result := True;
      Close;
    finally
      Free;
    end;
  end;
end;

function getDados(Connec: TIBCConnection; prTabela, prCampoChave, prValorChave, prCampoRetorno: String): String;
begin
  with TIBCQuery.Create(nil) do
  begin
    try
      Connection := Connec;
      Transaction:= Connec.DefaultTransaction;
      Sql.Text := 'select ' + prCampoRetorno +
                  ' from ' + prTabela +
                  ' where ' + prCampoChave + ' = ' + QuotedStr(prValorChave);
      Open;
      if RecordCount > 0 then
        Result := FieldByName(prCampoRetorno).AsAnsiString;
      Close;
    finally
      Free;
    end;
  end;
end;

//Norvan - Pega as Cotações do Dólar no site do Terra (Invertia)
procedure getValoresDolar(var CompraPtax, CompraComercial, CompraTurismo, CompraParalelo, VendaPtax, VendaComercial, VendaTurismo, VendaParalelo, Data : String);
var
  aux  : String;
  http : TIdHTTP;   // uses IdHTTP
begin
   try
     VendaParalelo   := '';
     CompraParalelo  := '';
     CompraComercial := '';
     VendaComercial  := '';
     CompraTurismo   := '';
     VendaTurismo    := '';
     CompraPtax      := '';
     VendaPtax       := '';
     Data           := '';

     http := TIdHTTP.Create(nil);
     try
//       http.ProxyParams.ProxyServer := '192.168.20.1';
//       http.ProxyParams.ProxyPort := 3128;
//       Aux  := http.Get('http://br.invertia.com/mercados/divisas/tiposdolar.aspx');
//       Aux  := http.Get('http://economia.terra.com.br/mercados/divisas/tiposdolar.aspx');
       Aux  := http.Get('http://economia.terra.com.br/stock/divisas.aspx');
     except
       raise;
     end;

//     VendaParalelo   := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALPR" class="masb">Dolar Paralelo', Aux)+160, 5));
//     CompraParalelo  := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALPR" class="masb">Dolar Paralelo', Aux)+115, 5));

//     CompraComercial := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALCM" class="masb">Dolar Comercial', Aux)+120, 5));
//     VendaComercial  := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALCM" class="masb">Dolar Comercial', Aux)+160, 5));
     CompraComercial := Trim(Copy(Aux, Pos('<td>DOLCM</td>', Aux)+32, 5));
     VendaComercial  := Trim(Copy(Aux, Pos('<td>DOLCM</td>', Aux)+61, 5));

//     CompraTurismo   := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALTR" class="masb">Dolar Turismo', Aux)+109, 5));
//     VendaTurismo    := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000REALTR" class="masb">Dolar Turismo', Aux)+149, 5));
     CompraTurismo   := Trim(Copy(Aux, Pos('<td>DOLTR</td>', Aux)+32, 5));
     VendaTurismo    := Trim(Copy(Aux, Pos('<td>DOLTR</td>', Aux)+61, 5));

//     CompraPtax      := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000DOLPTAX" class="masb">Dolar Ptax', Aux)+112, 5));
//     VendaPtax       := Trim(Copy(Aux, Pos('<a href="detalle.aspx?idtel=DI000DOLPTAX" class="masb">Dolar Ptax', Aux)+152, 5));
     CompraPtax      := Trim(Copy(Aux, Pos('<td>DOLPTAX</td>', Aux)+34, 5));
     VendaPtax       := Trim(Copy(Aux, Pos('<td>DOLPTAX</td>', Aux)+63, 5));

     Data           := Trim(Copy(Aux, Pos('<td colspan="5">', Aux) + 29, 10 ));
   finally
     FreeAndNil(http);
   end;
end;

{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Procedure: getGarantiaFinal                                                  }
{ Parametro: dDataEmissao (D) Data do Início da Garantia (Emissão NF)...       }
{            sTempoGarantia (S) String com o tempo de Garantia...              }
{                                                                              }
{ Objetivo.: Retornar a data final da Garantia do Produto...                   }
{                                                                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Solicitação         Procedimento                    Autor                    }
{ RMA                 Criar rotina...                 Arnaldo José  27/03/2008 }
{                                                                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
function getGarantiaFinal(dDataEmissao:TDateTime;sTempoGarantia:String):TDateTime;
Var
                     sPeriodo: String;
             wDia, wMes, wAno: Word;
  iPos, iAnos, iMeses, iTempo: Integer;
Begin                                          

  dDataEmissao := Trunc(dDataEmissao);
  Result := dDataEmissao;

  if (sTempoGarantia >' ') then Begin
    iPos := Pos(' ',sTempoGarantia);

    { Obtém o Tempo de Garantia oferecido...                                   }
    if (sTempoGarantia = '0') then begin
      iTempo := 0;
      sPeriodo := '';
    end
    Else
     if (UpperCase(Trim(sTempoGarantia)) <> 'LIFE TIME') then
      begin
       iTempo := StrToInt(Copy(sTempoGarantia, 1, iPos-1));              {        }
       sPeriodo := Copy(sTempoGarantia, iPos+1, Length(sTempoGarantia)); {        }
      end;


    if (UpperCase(Trim(sTempoGarantia)) = 'LIFE TIME') then
      Result := StrToDate('31/12/2100')
    else   { Verifica se o Período de Garantia é em Dias...                    }
    if (Pos('DIA',UpperCase(sPeriodo)) > 0) then
      Result := dDataEmissao + iTempo
    Else                              { ...Verifica se o Período é em Meses... }
    if (LowerCase(sPeriodo) = 'mês') or (LowerCase(sPeriodo) = 'meses') then begin
      { Destrincha a data...                                                   }
      DecodeDate(dDataEmissao, wAno, wMes, wDia);

      { Valida o Mês final da Garantia...                                      }
      iTempo := wMes + iTempo;
      if iTempo <= 12 then
        wMes := iTempo
      Else Begin
        { Obtém a quantidade de Meses e Anos existentes no Período...          }
        getQuantidadeAnos(iTempo, iMeses, iAnos);

        { Atribui às Variáveis o Valores apurados...                           }
        wMes := iMeses;
        wAno := wAno + iAnos;
      End;

      if ((wMes = 2) and (wdia >= 28)) then
        if Not isAnoBissexto(wAno) then
          wDia := 28
        else
          wDia := 29
      Else
      if (wDia = 31) and (wMes in [4,6,9,11]) then begin
         wDia := 1;
         wMes := wMes + 1;
      end;

      { Retorna a Data Encontrada...                                           }
      Result := EncodeDate(wAno, wMes, wDia);
    End
    Else                              { ...Verifica se o Período é em Anos...  }
    if (Pos('ANO',UpperCase(sPeriodo)) > 0) then begin
     { Destrincha a data...                                                    }
     DecodeDate(dDataEmissao, wAno, wMes, wDia);

     wAno := wAno + iTempo;

     { Retorna a Data Encontrada...                                            }
     Result := EncodeDate(wAno, wMes, wDia);
    End;
  End;
End;

{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Procedure.: getQuantidadeAnos                                                }
{ Parâmetros: iTempo (I) Tempo de Garantia...                                  }
{             iMeses (I) Quantidade de Meses a Retornar, pode receber Zero...  }
{             iAnos (I) Quantidade de Anos a Retornar, pode receber Zero...    }
{                                                                              }
{ Objetivo..: Calcular a Quantidade de Anos e Meses existentes em iTempo...    }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Ficha     Solicitação                               Autor         Data       }
{ RMA       Criar procedimento necessário             Arnaldo José  27/03/2008 }
{                                                                              }
{                                                                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
Procedure getQuantidadeAnos(iTempo:Integer; Var iMeses, iAnos:Integer);
Begin
  { Zera as Variáveis de Retorno...                                            }
  iAnos := 0;
  iMeses := 0;

  { Calcula os Anos existente no Período...                                    }
  while (iTempo > 12) do
  begin
    Inc(iAnos);
    iTempo := iTempo - 12;
  end;

  { Obtém os Meses restantes...                                                }
  iMeses := iTempo;
end;

Function isAnoBissexto(iAno:Integer):Boolean;
Begin
  if iAno mod 4 <> 0 then
    Result := False
  else
  if iAno mod 100 <> 0 then
    Result := True
  else
  if iAno mod 400 <> 0 then
    Result := False
  else
    Result := True;
end;

{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Procedure.: getFirstDay                                                      }
{ Parâmetros: iMes (I) Mês base...                                             }
{             iAno (I) Ano de base...                                          }
{                                                                              }
{ Objetivo..: Retornar a Data do 1º dia do mês...                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Ficha     Solicitação                               Autor         Data       }
{ RMA       Arnaldo José                              Arnaldo José  31/07/2008 }
{                                                                              }
{                                                                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
Function getFirstDay(iMes,iAno:Integer):TDate;
begin
  Result := StrToDate('01/'+IntToStr(iMes)+'/'+IntToStr(iAno));
end;

{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Procedure.: getLastDay                                                       }
{ Parâmetros: iMes (I) Mês base...                                             }
{             iAno (I) Ano de base...                                          }
{                                                                              }
{ Objetivo..: Retornar a Data do 1º dia do mês...                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
{ Ficha     Solicitação                               Autor         Data       }
{ RMA       Arnaldo José                              Arnaldo José  31/07/2008 }
{                                                                              }
{                                                                              }
{ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -}
Function getLastDay(iMes,iAno:Integer):TDate;
Var
  sDia: String;
begin
  { verifica qual é o último dia do Mês...                                     }
  if (iMes in [1,3,5,7,8,10,12]) then
    sDia := '31'
  Else
  if (iMes in [4,6,9,11]) then
    sDia := '30'
  Else Begin
    if isAnoBissexto(iAno) then
      sDia := '29'
    Else
      sDia := '28';
  end;

  { Monta a Data a ser retornada...                                            }
  Result := StrToDate(sDia+'/'+IntToStr(iMes)+'/'+IntToStr(iAno));
end;

function Arredonda(Connec: TIBCConnection; Valor: Double): Double;
begin
  if GetConfigur(Connec, 'QUATRODIGITOS', 'F', False) = 'F' then
    Result := roundrgf(Valor, 2)
  else
    Result := roundrgf(Valor, 4);
end;

function Arredonda2(Valor: Double): Double;
begin
  Result := round(Valor*100)/100
end;

function BlobSaveToStream(CampoBlob: TBlobField; Stream: TStream; Count: Int64; Barra: TProgressBar): Boolean;
const
  MaxBufSize = 8192; //$F000;
var
  Source: TStream;
  BufSize, N: Integer;
  Buffer: PByte;
begin
  Result := False;
  Source := CampoBlob.DataSet.CreateBlobStream(CampoBlob, bmRead);

  if Count = 0 then
  begin
    Source.Position := 0;
    Count := Source.Size;
  end;
//  Result := Count;
  if Count > MaxBufSize then
    BufSize := MaxBufSize
  else
    BufSize := Count;

  if Barra <> nil then
  begin
    Barra.Step := 1;
    Barra.Position := 0;
    Barra.Max := Count div BufSize;
    if Count mod BufSize > 0  then
      Barra.Max := Barra.Max + 1;
  end;

  GetMem(Buffer, BufSize);
  try
    N := 0;
    while Count <> 0 do
    begin
      if Count > BufSize then
        N := BufSize
      else
        N := Count;
      Source.ReadBuffer(Buffer^, N);
      Stream.WriteBuffer(Buffer^, N);
      Dec(Count, N);

      if Assigned(Barra) then
      begin
        Application.ProcessMessages;
        Barra.StepIt;
        Barra.Repaint;
        Sleep(3);
        Application.ProcessMessages;
      end
    end;
    Stream.Seek(0, soFromBeginning);
    Result := N > 0;
  finally
    FreeMem(Buffer, BufSize);
  end;
end;

function BlobLoadFromStream(CampoBlob : TBlobField; Stream: TStream; Count: Int64; Barra : TProgressBar): Boolean;
const
  MaxBufSize = 8192;//$F000;
var
  Destino : TStream;
  BufSize, N: Integer;
  Buffer: PChar;
begin
  Result := False;
  Destino := CampoBlob.DataSet.CreateBlobStream(CampoBlob, bmWrite);
  try
    if Count = 0 then
    begin
      Stream.Position := 0;
      Count := Stream.Size;
    end;

    if Count > MaxBufSize then
      BufSize := MaxBufSize
    else
      BufSize := Count;

    Barra.Position := 0;
    Barra.Max := Round(Count/BufSize);
    Barra.Step := 1;
    if Count mod BufSize > 0  then
      Barra.Max := Barra.Max + 1;

    GetMem(Buffer, BufSize);
    try
      N := 0;
      while Count <> 0 do
      begin
        if Count > BufSize then
          N := BufSize
        else
          N := Count;
        Stream.ReadBuffer(Buffer^, N);
        Destino.WriteBuffer(Buffer^, N);
        Dec(Count, N);
        Barra.StepIt;
        Application.ProcessMessages;
      end;
      Result := N > 0;
    finally
      FreeMem(Buffer, BufSize);
    end;
  finally
    Destino.Free;
  end;
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

////uses SysUtils, WinTypes, WinProcs, Messages, Dialogs, Forms,Controls, Classes, StdCtrls, Registry;
//type
//// conjunto de tipos de valores
//TKeyType = (ktString, ktBoolean, ktInteger, ktCurrency, ktDate, ktTime);
//// função exemplo
////Function ReadAppKey(const Key: String; KeyType: TKeyType; DefValue: Variant): Variant;
////Procedure WriteAppKey(const Key: String; const Value: Variant; KeyType: TKeyType);
////implementation
//function ReadAppKey(const Key: String; KeyType: TKeyType; DefValue: Variant): Variant;
//var
//  r: TRegistry;
//begin
//  // cria o objeto TRegistry
//  r := TRegistry.Create;
//  // conecta ao root diferente do padrão
//  r.RootKey := HKEY_LOCAL_MACHINE;
//  try
//    // abre a chave (no root selecionado)
//    // o segundo parâmetro True, indica que se a chave não existir, a operação de abertura poderá criá-la.
//    r.OpenKey('Software' + Application.Title, True);
//    Result := DefValue;
//    // testa se existe o valor que se deseja ler.
//    // note que, para verificar a existência de chaves, utilizados KeyExists([chave]) e para verificar a existência de conjunto de //chaves de uma chave, utilizamos ValueExists([valor])
//    if r.ValueExists(Key) then
//    begin
//      case KeyType of
//        // lê o valor da chave em formato String
//        ktString: Result := r.ReadString(Key);
//        // lê o valor da chave em formato Boolean
//        ktBoolean: Result := r.ReadBool(Key);
//        // lê o valor da chave em formato Integer
//        ktInteger: Result := r.ReadInteger(Key);
//        // lê o valor da chave em formato Currency (moeda)
//        ktCurrency: Result := r.ReadCurrency(Key);
//        // lê o valor da chave em formato TDateTime (data)
//        ktDate: Result := r.ReadDate(Key);
//        // lê o valor da chave em formato TDateTime (hora)
//        ktTime: Result := r.ReadTime(Key);
//      end;
//    end;
//  finally
//    // destroy o objeto criado
//    r.Free;
//  end;
//end;
//
//procedure WriteAppKey(const RootKey: Cardinal; const Key: String; const Value: Variant; KeyType: TKeyType);
//var
//  r: TRegistry;
//begin
//  // cria o objeto TRegistry
//  r := TRegistry.Create;
//  // conecta ao root diferente do padrão
//  r.RootKey := RootKey; //HKEY_LOCAL_MACHINE;
//  try
//    // abre a chave (no root selecionado)
//    r.OpenKey('Software' + Application.Title, True);
//    case KeyType of
//      // grava o valor da chave em formato String
//      ktString: r.WriteString(Key, Value);
//      // grava o valor da chave em formato Boolean
//      ktBoolean: r.WriteBool(Key, Value);
//      // grava o valor da chave em formato Integer
//      ktInteger: r.WriteInteger(Key, Value);
//      // grava o valor da chave em formato Currency (moeda)
//      ktCurrency: r.WriteCurrency(Key, Value);
//      // grava o valor da chave em formato TDateTime (Data)
//      ktDate: r.WriteDate(Key, Value);
//      // grava o valor da chave em formato TDateTime (Hora)
//      ktTime: r.WriteTime(Key, Value);
//    end;
//  finally
//    r.Free;
//  end;
//end;


{ TConfigur }

constructor TConfigur.Create(oConnec: TIBCConnection = nil);
begin
  inherited Create;
  FCriado          := True;
  FqrConfigur      := TIBCQuery.Create(nil);
  FqrConfigurEsp   := TIBCQuery.Create(nil);
  FqrConfigEmpresa := TIBCQuery.Create(nil);

  FNoEmpresa       := 1;

  if Assigned(oConnec) then
    Connec := oConnec;
end;

destructor TConfigur.Destroy;
begin
  FreeAndNil(FqrConfigur);
  FreeAndNil(FqrConfigurEsp);
  FreeAndNil(FqrConfigEmpresa);
  FCriado := False;
  inherited Destroy;
end;

function TConfigur.GetValue(sField: string; vDefaultValue: Variant; bConfigurEsp: Boolean): Variant;
begin
  if bConfigurEsp then
  begin
    if (not FqrConfigurEsp.Active) or (Now > FConfigurEspRefresh + (60/86400)) then // 60 segundos
    begin
      FConfigurEspRefresh := Now;
      FqrConfigurEsp.Close;
      FqrConfigurEsp.SQL.Text := 'select * from configuresp ';
      FqrConfigurEsp.Open;
    end;
    if FqrConfigurEsp.FieldByName(sField).IsNull then
      Result := vDefaultValue
    else
      Result := FqrConfigurEsp.FieldByName(sField).Value;
  end
  else
  begin
    if (not FqrConfigur.Active) or (Now > FConfigurRefresh + (60/86400)) then // 60 segundos
    begin
      FConfigurRefresh := Now;
      FqrConfigur.Close;
      FqrConfigur.SQL.Text := 'select * from configur ';
      FqrConfigur.Open;
    end;
    if FqrConfigur.FieldByName(sField).IsNull then
      Result := vDefaultValue
    else
      Result := FqrConfigur.FieldByName(sField).Value;
  end;
end;

procedure TConfigur.RefreshDados;
begin
  FConfigurRefresh   := 0;
  FConfigurEspRefresh:= 0;
  FConfigEmpresa     := 0;
end;

class function TConfigur.AddConfigEmpresa(oConnec: TIBCConnection; iNoEmpresa: Integer; sChave, sParametro, sValor: String): Boolean;
var
  qrLocal: TIBCQuery;
begin
  qrLocal := TIBCQuery.Create(nil);
  try
    qrLocal.Connection := oConnec;
    qrLocal.Transaction:= oConnec.DefaultTransaction;

    if not oConnec.InTransaction then
      oConnec.StartTransaction;
    try
      qrLocal.SQL.Text := 'update or insert into configuracao (noempresa, chave, parametro, valor) '+
                          'values (:noempresa, :chave, :parametro, :valor) '+
                          'matching(noempresa, chave, parametro) ';
      qrLocal.ParamByName('noempresa').AsInteger := iNoEmpresa;
      qrLocal.ParamByName('chave').AsString      := sChave;
      qrLocal.ParamByName('parametro').AsString  := sParametro;
      qrLocal.ParamByName('valor').AsString      := sValor;
      qrLocal.ExecSQL;

      oConnec.CommitRetaining;
    except
      on E: Exception do
      begin
        oConnec.RollbackRetaining;
        raise Exception.Create('Erro ao cadastrar configuração para a empresa: '+E.Message+#13#13+
                               'Empresa='+IntToStr(iNoEmpresa)+'   Chave="'+sChave+'"   Paramêtro="'+sParametro+'"   Valor="'+sValor+'"');
      end;
    end;

  finally
    FreeAndNil(qrLocal);
  end;

end;

function TConfigur.getConfigEmpresa(sChave, sParametro: string; sDefaultValue: String): String;
begin
  if (not FqrConfigEmpresa.Active) or (Now > FConfigEmpresa + (60/86400)) then // 60 segundos
  begin
    FConfigEmpresa := Now;
    FqrConfigEmpresa.Close;
    FqrConfigEmpresa.SQL.Text := 'select * from configuracao where noempresa = :noempresa';
    FqrConfigEmpresa.ParamByName('noempresa').AsInteger := FNoEmpresa;
    FqrConfigEmpresa.Open;
  end;
  if FqrConfigEmpresa.Locate('chave;parametro',VarArrayOf([sChave, sParametro]), []) then
    Result := FqrConfigEmpresa.FieldByName('valor').AsString
  else
  begin
    Result := sDefaultValue;
    AddConfigEmpresa(Connec, FNoEmpresa, sChave, sParametro, sDefaultValue);
    FqrConfigEmpresa.Close;
    FqrConfigEmpresa.Open;
  end;
end;

procedure TConfigur.setConnec(const Value: TIBCConnection);
begin
  FConnec := Value;

  FqrConfigur.Connection := FConnec;
  FqrConfigur.Transaction:= FConnec.DefaultTransaction;

  FqrConfigurEsp.Connection := FConnec;
  FqrConfigurEsp.Transaction:= FConnec.DefaultTransaction;
end;

procedure TConfigur.setEmpresa(const Value: Integer);
begin
  if (FNoEmpresa <> Value) and (FqrConfigEmpresa.Active) then
    FqrConfigEmpresa.Close;

  FNoEmpresa := Value;
end;

function getAppDataPath : string;
const
  SHGFP_TYPE_CURRENT = 0;
var
  path: array [0..MaxChar] of char;
begin
  SHGetFolderPath(0,CSIDL_APPDATA,0,SHGFP_TYPE_CURRENT,@path[0]);
  Result := StrPas(path);
end;

function getIniciarEm: string;
begin
  Result := SysUtils.GetCurrentDir;
end;

function getDropboxFolder: String;
var
  Stream: TBytesStream;
  dbPath, texto: string;
  Arquivo: TextFile;
begin
  AssignFile(Arquivo, getAppDataPath+'\dropbox\host.db');
  try
    Reset(Arquivo);
    ReadLn(Arquivo, texto);
    ReadLn(Arquivo, texto);
    Result := TEncoding.ASCII.GetString(DecodeBase64(Texto));
  finally
    CloseFile(Arquivo);
  end;
end;

function ExtractFileNameWithoutExt(sFileName: String): String;
begin
  if ExtractFileExt(sFileName) > ' ' then
    Result := Copy(sFileName, 1, Pos(ExtractFileExt(sFileName), sFileName)-1)
  else
    Result := sFileName;
end;

function MontaAspasComVirgula(Entrada : String): String;
begin
  Result := Entrada;
  Result := StringReplace(Result, '''', '', [rfReplaceAll]);
  Result := StringReplace(Result, ',', ''',''', [rfReplaceAll]);
  Result := '''' + Result + '''';
end;

function GetCheckSum(FileName: string): DWORD;
var
  F: file of DWORD;
  P: Pointer;
  Fsize: DWORD;
  Buffer: array [0..500] of DWORD;
begin
  FileMode := 0;
  AssignFile(F, FileName);
  Reset(F);
  Seek(F, FileSize(F) div 2);
  Fsize := FileSize(F) - 1 - FilePos(F);
  if Fsize > 500 then Fsize := 500;
  BlockRead(F, Buffer, Fsize);
  Close(F);
  P := @Buffer;
  asm
    xor eax, eax
    xor ecx, ecx
    mov edi , p
    @again:
      add eax, [edi + 4*ecx]
      inc ecx
      cmp ecx, fsize
    jl @again
    mov @result, eax
  end;
end;

function getParamApp(sParam: string): string;
var
  i, iAux: Integer;
begin
  Result := '';
  for i := 0 to ParamCount do
  begin
    if Pos(LowerCase(sParam)+'=', LowerCase(ParamStr(i))) > 0 then
    begin
      iAux   := Pos(LowerCase(sParam)+'=',LowerCase(ParamStr(i)))+Length(LowerCase(sParam)+'=');
      Result := Copy(ParamStr(i), iAux, Length(ParamStr(i))-iAux+1);
    end;
  end;
end;

//{$IFDEF VER210} // Delphi 2010
{function Compactar(sDestZipFile: String; slFiles: TStringList): String;
var
  i: Integer;
  FAbZipKit: TAbZipKit;
begin
  FAbZipKit := TAbZipKit.Create(nil);
  try
    with FAbZipKit do
    begin
      AutoSave := True;
      DOSMode  := False;
      CompressionMethodToUse := smBestMethod;
      DeflationOption        := doMaximum;
      ExtractOptions         := [eoCreateDirs];
      StoreOptions           := [soStripDrive,soStripPath,soRemoveDots,soReplace];
    end;

    Result := '';
    if slFiles.Count > 0 then
      sDestZipFile := ChangeFileExt((slFiles[0]),'.zip');
    if FileExists(sDestZipFile) then
      SysUtils.DeleteFile(sDestZipFile);
    try
      FAbZipKit.FileName := sDestZipFile;
      for i:= 0 to slFiles.Count-1 do
        FAbZipKit.AddFiles(slFiles[i], 0);
    finally
      FAbZipKit.CloseArchive;
    end;
  finally
    FreeAndNil(FAbZipKit);
  end;
  Result := sDestZipFile;
end;}

{function Descompactar(sZipFile, sPathDest: String; bDeleteZipFileAfter: Boolean=False): String;
var
  FAbZipKit: TAbZipKit;
begin
  Result    := iif(sPathDest='', ExtractFilePath(sZipFile), sPathDest);
  FAbZipKit := TAbZipKit.Create(nil);
  try
    with FAbZipKit do
    begin
      AutoSave := True;
      DOSMode  := False;
      CompressionMethodToUse := smBestMethod;
      DeflationOption        := doMaximum;
      ExtractOptions         := [eoCreateDirs];
      StoreOptions           := [soStripDrive,soStripPath,soRemoveDots,soReplace];

      FileName      := sZipFile;
      BaseDirectory := Result;
      ExtractFiles('*.*');
      CloseArchive;
      Result := BaseDirectory;
    end;
    if bDeleteZipFileAfter then
      SysUtils.DeleteFile(sZipFile);
  finally
    FreeAndNil(FAbZipKit);
  end;
end;}
//{$ENDIF}

function FileDateTime(const FileName: string; sTipo: string='Modified'): TDateTime;

  function ReportTime(const Name: string; const FileTime: TFileTime): TDateTime;
  var
    SystemTime, LocalTime: TSystemTime;
  begin
    if not FileTimeToSystemTime(FileTime, SystemTime) then
      RaiseLastOSError;
    if not SystemTimeToTzSpecificLocalTime(nil, SystemTime, LocalTime) then
      RaiseLastOSError;
    Result := SystemTimeToDateTime(LocalTime);
  end;
var
  fad: TWin32FileAttributeData;

begin
  if not GetFileAttributesEx(PChar(FileName), GetFileExInfoStandard, @fad) then
    RaiseLastOSError;
  if sTipo = 'Created' then
    ReportTime('Created', fad.ftCreationTime)
  else if sTipo = 'Acessed' then
    ReportTime('Accessed', fad.ftLastAccessTime)
  else
    Result := ReportTime('Modified', fad.ftLastWriteTime);
end;

function getExcelColumnName(columnNumber: Integer): String;
var
  dividend: Integer;
  columnName: String;
  modulo: Integer;
begin
  dividend := columnNumber;
  columnName := '';

  while (dividend > 0) do
  begin
    modulo := (dividend - 1) mod 26;
    columnName := Chr(65 + modulo) + columnName;
    dividend := Trunc((dividend - modulo) / 26);
  end;

  Result := columnName;
end;

function getExcelColumnIndex(columnName: String): Integer;
var
  i, modulo: Integer;
begin
  Result := 0;
  Modulo := 0;
  for i := 1 to Length(columnName) do
  begin
    Result := modulo + Ord(columnName[i])-64;
    if i = 1 then
      modulo := (26*Result)
  end;
end;

function NomeClasse(const Janela: HWND): string;
var
  Buffer: array[0..250] of Char;
begin
  if GetClassName(Janela, Buffer, SizeOf(Buffer)) > 0 then
    Result := string(Buffer)
  else
    Result := '';
end;

function WinVersion: string;
var
  VersionInfo: TOSVersionInfo;
begin
  Result := '';
  VersionInfo.dwOSVersionInfoSize := SizeOf(VersionInfo);
  GetVersionEx(VersionInfo);
  case VersionInfo.dwPlatformId of
    1:case VersionInfo.dwMinorVersion of
        0: Result := 'Windows 95';
        10: Result := 'Windows 98';
        90: Result := 'Windows Me';
      end;
    2:case VersionInfo.dwMajorVersion of
        3: Result:='Windows NT 3.51';
        4: Result:='Windows NT 4.0';
        5: case VersionInfo.dwMinorVersion of
             0: Result := 'Windows 2000';
             1: Result := 'Windows XP';
             2: Result := 'Windows XP x64';
           end;
        6: case VersionInfo.dwMinorVersion of
             0: Result := 'Windows Vista';
             1: Result := 'Windows 7';
             2: Result := 'Windows 8';
           end;
      end;
  end;
end;

function EnumWindowsProc(Wnd: HWND; lb: TStringList): BOOL; stdcall;
var
  Caption: Array [0..128] of Char;
begin
  Result := True;
  if IsWindowVisible(Wnd) and ((GetWindowLong(Wnd, GWL_HWNDPARENT) = 0) or
    (HWND(GetWindowLong(Wnd, GWL_HWNDPARENT)) = GetDesktopWindow)) and
    ((GetWindowLong(Wnd, GWL_EXSTYLE) and WS_EX_TOOLWINDOW) = 0) then
  begin
    SendMessage( Wnd, WM_GETTEXT, Sizeof( caption ),integer(@caption));
    lb.AddObject( caption, TObject( Wnd ));
  end;
end;

function getMonitorResolution(oForm: TForm; bWorkArea: Boolean): TPoint;
begin
  if bWorkArea then
  begin
    Result.X := oForm.Monitor.WorkareaRect.Right - oForm.Monitor.WorkareaRect.Left;
    Result.Y := oForm.Monitor.WorkareaRect.Bottom - oForm.Monitor.WorkareaRect.Top;
  end
  else
  begin
    Result.X := oForm.Monitor.Width;
    Result.Y := oForm.Monitor.Height;
  end;
end;

function DateTimeConcat(dData: TDate; tHora: TTime): TDateTime;
begin
  Result := Trunc(dData)+(tHora-Trunc(tHora));
end;

{function ExecMethod(OnObject: TObject; MethodName: string): String;
var
   Routine: TMethod;
   Exec: TExec;
begin
   Routine.Data := Pointer(OnObject);
   Routine.Code := OnObject.MethodAddress(MethodName);
   if not Assigned(Routine.Code) then
     Exit;
   Exec := TExec(Routine);
   Result := Exec;
end;}

{function ExecClassMethod(sClassName, sMethod: String): String;
var
  Form: TForm;
begin
  Form := TFormClass(FindClass(sClassName)).Create(nil);
  try
    Result := uger_rotinas.ExecMethod(Form, sMethod);
  finally
    FreeAndNil(Form);
  end;
end;}

// Alterar intensidade da cor
// Exemplo:
//     cor1 := Intensidade(clMoneyGreen, -50); //fica 50% mais escura
//     cor2 := Intensidade(clMoneyGreen, 70); //fica 70% mais clara
function CorIntensidade(cCor: TColor; iValor: integer): TColor;
var
  H, S, L: Word;
begin
  ColorRGBToHLS(cCor, H, L, S);

  if iValor > 100 then
    iValor := 100;
  if iValor < 0 then
    iValor := 0;

  iValor := trunc((255*iValor) / 100);
  Result := ColorHLSToRGB(H, iValor, S);
end;

procedure TrimAppMemorySize;
var
  MainHandle : THandle;
begin
  try
    MainHandle := OpenProcess(PROCESS_ALL_ACCESS, false, GetCurrentProcessID) ;
    SetProcessWorkingSetSize(MainHandle, $FFFFFFFF, $FFFFFFFF) ;
    CloseHandle(MainHandle) ;
  except
  end;
    Application.ProcessMessages;
end;

function GetNodeTreeViewByText(ATree : TTreeView; AValue:String; AVisible: Boolean): TTreeNode;
var
    Node: TTreeNode;
begin
  Result := nil;
  if ATree.Items.Count = 0 then Exit;
  Node := ATree.Items[0];
  while Node <> nil do
  begin
    if UpperCase(Node.Text) = UpperCase(AValue) then
    begin
      Result := Node;
      if AVisible then
        Result.MakeVisible;
      Break;
    end;
    Node := Node.GetNext;
  end;
end;

procedure SortTitleClick(Column: TColumn; bDesativar: Boolean=False);
var
  i: Integer;
begin
  if not Assigned(Column) then
    Exit;

  for i := 0 to TDBGrid(Column.Grid).Columns.Count-1 do
    TDBGrid(Column.Grid).Columns[i].Title.Font.Color := clWindowText;

  if bDesativar then
    TIBCQuery(Column.Field.DataSet).IndexFieldNames := ''
  else
  begin
    Column.Title.Font.Color := clRed;

    if TIBCQuery(Column.Field.DataSet).IndexFieldNames = Column.FieldName then
      TIBCQuery(Column.Field.DataSet).IndexFieldNames := Column.FieldName+ ' DESC'
    else
      TIBCQuery(Column.Field.DataSet).IndexFieldNames := Column.FieldName;
  end;
end;

end.

// outro exemplo - validar menu
//  While Not DMDados.TBModulos.Eof Do
//  Begin
//    MainMenu1.Items(DMDados.TBModulosMod_Nome.Text).Enabled:=DMDados.TBModulosMod_Ativo.AsBoolean;
//    Next;
//  End;
///---------
///While Not DMDados.TBModulos.Eof Do
//Begin
//if DMDados.TBModulosMod_Ativo.AsBoolean then
//  StList.Add(DMDados.TBModulosMod_Nome.AsString);
//  DMDados.TBModulos.Next;
//End;
//
//for i := 0 to ComponentCount - 1 do
//begin
//  if Components[i] is TMenuItem then
//  begin
//        s := (Components[i] as TMenuItem).Name;
//        (Components[i] as TMenuItem).Enabled := StList.IndexOf(s) >= 0;
//  end;
//end;
//-----------------------------

//procedure CopiarParaClipBoard(sStr: String; Grid: TDBAltGrid=nil);
//var
//  BM: TBookmark;
//  i: Integer;
//begin
//  inherited;
//  BM := Grid.DataSource.DataSet.GetBookmark;
//
//  Clipboard.AsText := '';
//  Grid.DataSource.DataSet.DisableControls;
//  try
//    for i := 0 to Grid.SelectedRows.Count - 1 do
//    begin
//      Grid.DataSource.DataSet.BookMark := Grid.SelectedRows[I];
//      if (i=1) or (i=Grid.SelectedRows.Count - 1) then
//        Clipboard.AsText := Clipboard.AsText + sStr
//      else
//        Clipboard.AsText := Clipboard.AsText + sStr + #13#10;
//    end;
//  finally
//    Grid.DataSource.DataSet.GotoBookmark(BM);
//    Grid.DataSource.DataSet.EnableControls;
//  end;
//
////  Clipboard.AsText := trim(qrSelecaoNOME.AsString)+' = '+FormatFloat('R$###,###,##0.00',qrSelecaoPRECOMINIMO.AsFloat);
//end;


// ClientDataset - Adicionar fields em runtime
//var
//  i:Integer;
//  Query:TQuery;
//  FieldList:TObjectList;
//  Field:TField;
//begin
//  inherited Create(AOwner);
//  Query:=TQuery.Create(Self);
//  FieldList:=TObjectList.Create;
//  try
//    while cds.FieldCount > 0 do
//    begin
//      FieldList.Add(cds.Fields[0]);
//      TField(FieldList.Items[FieldList.Count-1]).DataSet:=nil;
//    end;
//    cds.FieldDefs.Clear;
//    cds.Fields.Clear;
//    for i:=0 to FieldList.Count-1 do
//    begin
//      case TField(FieldList.Items[i]).DataType of
//        ftString:  Field:=TStringField.Create(cds);
//        ftInteger: Field:=TIntegerField.Create(cds);
//        ftFloat:   Field:=TFloatField.Create(cds);
//        ftDate:    Field:=TDateField.Create(cds);
//        ftTime:    Field:=TTimeField.Create(cds);
//        ftDateTime:Field:=TDateTimeField.Create(cds);
//        ftBlob:    Field:=TBlobField.Create(cds);
//        ftMemo:    Field:=TMemoField.Create(cds);
//        ftAutoInc: Field:=TAutoIncField.Create(cds);
//        else
//          Field:=TField.Create(cds);
//      end;
//      with Field do
//      begin
//        FieldName         := TField(FieldList.Items[i]).FieldName         ;
//        Alignment         := TField(FieldList.Items[i]).Alignment         ;
//        Calculated        := TField(FieldList.Items[i]).Calculated        ;
//        DefaultExpression := TField(FieldList.Items[i]).DefaultExpression ;
//        DisplayLabel      := TField(FieldList.Items[i]).DisplayLabel      ;
//        DisplayWidth      := TField(FieldList.Items[i]).DisplayWidth      ;
//        EditMask          := TField(FieldList.Items[i]).EditMask          ;
//        FieldKind         := TField(FieldList.Items[i]).FieldKind         ;
//        Index             := TField(FieldList.Items[i]).Index             ;
//        KeyFields         := TField(FieldList.Items[i]).KeyFields         ;
//        Lookup            := TField(FieldList.Items[i]).Lookup            ;
//        LookupCache       := TField(FieldList.Items[i]).LookupCache       ;
//        LookupDataSet     := TField(FieldList.Items[i]).LookupDataSet     ;
//        LookupKeyFields   := TField(FieldList.Items[i]).LookupKeyFields   ;
//        LookupResultField := TField(FieldList.Items[i]).LookupResultField ;
//        Origin            := TField(FieldList.Items[i]).Origin            ;
//        ParentField       := TField(FieldList.Items[i]).ParentField       ;
//        ProviderFlags     := TField(FieldList.Items[i]).ProviderFlags     ;
//        ReadOnly          := TField(FieldList.Items[i]).ReadOnly          ;
//        Required          := TField(FieldList.Items[i]).Required          ;
//        Visible           := TField(FieldList.Items[i]).Visible           ;
//      end;
//      if TField(FieldList.Items[i]).DataType in [ftInteger, ftFloat] then
//      begin
//        with TNumericField(Field) do
//        begin
//          DisplayFormat  := TNumericField(FieldList.Items[i]).DisplayFormat;
//          EditFormat     := TNumericField(FieldList.Items[i]).EditFormat   ;
//        end;
//      end else if TField(FieldList.Items[i]).DataType in [ftDateTime, ftDate, ftTime] then
//      begin
//        with TDateTimeField(Field) do
//        begin
//          DisplayFormat  := TDateTimeField(FieldList.Items[i]).DisplayFormat;
//        end;
//      end;
//      Field.DataSet:=cds;
//    end;
//    Query.DatabaseName:=DataBase;
//    Query.SQL.Add('Select nome_campo, tipo, tamanho from campos where ativo = 1');
//    Query.Open;
//    while not Query.Eof do
//    begin
//      case Query.FieldByName('tipo').AsInteger of
//        I_TIPO_CAMPO_INTEIRO:
//          begin
//            Field           := TIntegerField.Create(cds);
//          end;
//        I_TIPO_CAMPO_REAL:
//          begin
//            Field           := TFloatField.Create(cds);
//          end;
//        I_TIPO_CAMPO_LOGICO:
//          begin
//            Field           := TIntegerField.Create(cds);
//            Field.Required  := True;
//            Field.Tag       := 999;
//          end;
//        I_TIPO_CAMPO_DATA:
//          begin
//            Field           := TDateField.Create(cds);
//          end;
//        else begin
//          Field           := TStringField.Create(cds);
//          Field.Size      := Query.FieldByName('tamanho').AsInteger;
//        end;
//      end;
//      Field.FieldName   := 'user_'+NomeCampo(Query.FieldByName('nome_campo').AsString);
//      Field.DisplayLabel:= Query.FieldByName('nome_campo').AsString;
//      Field.DataSet:=cds;
//      Query.Next;
//    end;
//    cds.CreateDataSet;
//    Query.First;
//    while not Query.Eof do
//    begin
//      cds.FieldByName('user_'+NomeCampo(Query.FieldByName('nome_campo').AsString)).DisplayLabel:=Query.FieldByName('nome_campo').AsString;
//      Query.Next;
//    end;
//  finally
//    Query.Free;
//    FieldList.Free;
//  end;
//end;
//-------------------------------------------------------
//const
//  Codes64 = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/';
//
//function Encode64(S: string): string;
//var
//  i: Integer;
//  a: Integer;
//  x: Integer;
//  b: Integer;
//begin
//  Result := '';
//  a := 0;
//  b := 0;
//  for i := 1 to Length(s) do
//  begin
//    x := Ord(s[i]);
//    b := b * 256 + x;
//    a := a + 8;
//    while a >= 6 do
//    begin
//      a := a - 6;
//      x := b div (1 shl a);
//      b := b mod (1 shl a);
//      Result := Result + Codes64[x + 1];
//    end;
//  end;
//  if a > 0 then
//  begin
//    x := b shl (6 - a);
//    Result := Result + Codes64[x + 1];
//  end;
//end;
//
//function Decode64(S: string): string;
//var
//  i: Integer;
//  a: Integer;
//  x: Integer;
//  b: Integer;
//begin
//  Result := '';
//  a := 0;
//  b := 0;
//  for i := 1 to Length(s) do
//  begin
//    x := Pos(s[i], codes64) - 1;
//    if x >= 0 then
//    begin
//      b := b * 64 + x;
//      a := a + 6;
//      if a >= 8 then
//      begin
//        a := a - 8;
//        x := b shr a;
//        b := b mod (1 shl a);
//        x := x mod 256;
//        Result := Result + chr(x);
//      end;
//    end
//    else
//      Exit;
//  end;
//end;


