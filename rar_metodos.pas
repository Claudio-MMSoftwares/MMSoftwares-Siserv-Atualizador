unit rar_metodos;

interface

uses
  Winapi.Windows,
  System.SysUtils;

const
  RAR_OM_LIST     = 0;
  RAR_OM_EXTRACT  = 1;
  RAR_SKIP        = 0;
  RAR_TEST        = 1;
  RAR_EXTRACT     = 2;

type
  TRARHeaderData = record
    ArcName: array[0..1023] of AnsiChar;
    FileName: array[0..1023] of AnsiChar;
    Flags: Cardinal;
    PackSize: Cardinal;
    UnpSize: Cardinal;
    HostOS: Cardinal;
    FileCRC: Cardinal;
    FileTime: Cardinal;
    UnpVer: Cardinal;
    Method: Cardinal;
    FileAttr: Cardinal;
    CmtBuf: PAnsiChar;
    CmtBufSize: Cardinal;
    CmtSize: Cardinal;
    CmtState: Cardinal;
  end;

  TRAROpenArchiveData = record
    ArcName: PAnsiChar;
    OpenMode: Cardinal;
    OpenResult: Cardinal;
    CmtBuf: PAnsiChar;
    CmtBufSize: Cardinal;
    CmtSize: Cardinal;
    CmtState: Cardinal;
  end;

// Declarações das funções externas da DLL
function RAROpenArchive(var ArchiveData: TRAROpenArchiveData): THandle; stdcall; external 'UnRAR.dll';
function RARCloseArchive(hArcData: THandle): Integer; stdcall; external 'UnRAR.dll';
function RARReadHeader(hArcData: THandle; var HeaderData: TRARHeaderData): Integer; stdcall; external 'UnRAR.dll';
function RARProcessFile(hArcData: THandle; Operation: Integer; DestPath, DestName: PAnsiChar): Integer; stdcall; external 'UnRAR.dll';

// Declaração do método público da unit
procedure ExtrairRAR(const ARquivoRAR, ADestino: string);

implementation

procedure ExtrairRAR(const ARquivoRAR, ADestino: string);
var
  ArcData: TRAROpenArchiveData;
  Header: TRARHeaderData;
  hArc: THandle;
  Ret: Integer;
begin
  if not FileExists(ARquivoRAR) then
    raise Exception.Create('Arquivo RAR não encontrado: ' + ARquivoRAR);

  ForceDirectories(ADestino);

  FillChar(ArcData, SizeOf(ArcData), 0);
  ArcData.ArcName := PAnsiChar(AnsiString(ARquivoRAR));
  ArcData.OpenMode := RAR_OM_EXTRACT;

  hArc := RAROpenArchive(ArcData);
  if hArc = 0 then
    raise Exception.CreateFmt('Erro ao abrir o RAR (%d)', [ArcData.OpenResult]);

  FillChar(Header, SizeOf(Header), 0);

  try
    while RARReadHeader(hArc, Header) = 0 do
    begin
      Ret := RARProcessFile(hArc, RAR_EXTRACT, PAnsiChar(AnsiString(ADestino)), nil);
      if Ret <> 0 then
        raise Exception.CreateFmt('Erro ao extrair: %s (código %d)', [Header.FileName, Ret]);
    end;
  finally
    RARCloseArchive(hArc);
  end;
end;

end.

