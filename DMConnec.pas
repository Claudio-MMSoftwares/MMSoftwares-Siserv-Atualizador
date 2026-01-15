unit DMConnec;

interface

uses
  SysUtils, Classes, DB, Forms, Controls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf,
  FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async,
  FireDAC.Phys, FireDAC.Phys.MySQL, FireDAC.VCLUI.Wait, FireDAC.Comp.UI,
  FireDAC.Comp.Client, FireDAC.Phys.IBBase, FireDAC.Phys.FB, FireDAC.Phys.PG,
  FireDAC.Comp.ScriptCommands, FireDAC.Stan.Util, FireDAC.Comp.Script,
  FireDAC.Phys.IB, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, uutil_rotinas;

type
  TdtmConnec = class(TDataModule)
    FDConnection: TFDConnection;
    FDTransaction: TFDTransaction;
    FDGUIxWaitCursor: TFDGUIxWaitCursor;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDQuery: TFDQuery;
    FDConnRemoto: TFDConnection;
    FDQryRemoto: TFDQuery;
    FDTransRemoto: TFDTransaction;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
    FDateLoaded             : TDateTime;
    FUsuario                : string;
    FServerDB               : String;
    FPathDB                 : String;
    FUserDB                 : String;
    FPassDB                 : String;

    FPassDBImagem                 : String;

    FPortaDB                : String;
    FDialectDB              : Integer;
    FCharsetDB              : String;
    FRoleDB                 : String;
    FslTermosDesconsiderados: TStringList;

    function getUsuario: string;



    procedure setUserDB(const Value: String);
    procedure setServerDB(const Value: String);
    procedure setPathDB(const Value: String);
    procedure setPortaDB(const Value: String);
    procedure setDialectDB(const Value: Integer);
    procedure setRoleDB(const Value: string);
    procedure setCharsetDB(const Value: string);
    procedure setUsuario(const Value: string);
  public
    { Public declarations }
  end;

var
  dtmConnec: TdtmConnec;

implementation


{$R *.dfm}

procedure TdtmConnec.DataModuleCreate(Sender: TObject);
var
  pDatabase, pUserName, pPassword: String;
  iDatabase, iUserName, iPassword: String;
  i, iAux : Integer;
  sIP, sPath: String;

  iIP, iPath: String;
begin
end;


function TdtmConnec.getUsuario: String;
begin
  Result := FUsuario;
end;


procedure TdtmConnec.setUsuario(const Value: String);
begin
  FUsuario := Value;
end;

procedure TdtmConnec.setUserDB(const Value: String);
begin
  FUserDB := Value;
  FPassDB := '';
end;

procedure TdtmConnec.setPathDB(const Value: String);
begin
  FPathDB := Value;
end;

procedure TdtmConnec.setPortaDB(const Value: String);
begin
  FPortaDB := Value;
end;

procedure TdtmConnec.setServerDB(const Value: String);
begin
  FServerDB := Value;
end;

procedure TdtmConnec.setDialectDB(const Value: Integer);
begin
  FDialectDB := Value;
end;

procedure TdtmConnec.setCharsetDB(const Value: String);
begin
  FCharsetDB := Value;
end;

procedure TdtmConnec.setRoleDB(const Value: String);
begin
  FRoleDB := Value;
end;


end.
