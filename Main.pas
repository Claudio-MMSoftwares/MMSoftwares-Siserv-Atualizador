unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, IBX.IBScript,
  System.SysUtils, Math, Vcl.StdCtrls, FireDAC.Comp.Client, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Phys.FB, FireDAC.VCLUI.Wait, FireDAC.Comp.UI,
  FireDAC.UI.Intf, FireDAC.Stan.Intf, FireDAC.Phys, FireDAC.Phys.IBBase,
  Vcl.ComCtrls, FireDAC.DApt, FireDAC.Stan.Async, FireDAC.Comp.ScriptCommands,
  FireDAC.Stan.Util, FireDAC.Comp.Script, System.IniFiles, uutil_rotinas,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdFTP, Vcl.Menus, IdSSLOpenSSL, System.IOUtils, FileCtrl,
  rar_metodos, Vcl.Samples.Gauges;

type
  TComandosSQL = Record
    Versao     : String;
    ComandoSQL : String;
  end;

  TProcessKind = (pkAtualizacao, pkTesteScripts);

  TMainForm = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnExecutarProcesso: TButton;
    FDScript: TFDScript;
    PageControl1: TPageControl;
    TabComando: TTabSheet;
    AExecutarMemo: TMemo;
    TabErros: TTabSheet;
    AErrorMemo: TMemo;
    Panel3: TPanel;
    lbExecutados: TLabel;
    lbErros: TLabel;
    GaugeProgresso: TGauge;
    GaugeScripts: TGauge;
    PanelProcesso: TPanel;
    lbProcTitulo: TLabel;
    lbProcEtapa: TLabel;
    lbProcDesc: TLabel;
    TrayIcon: TTrayIcon;
    PopupMenu: TPopupMenu;
    Abrir1: TMenuItem;
    Fechar1: TMenuItem;
    FTP: TIdFTP;
    Timer: TTimer;
    btTestarScripts: TButton;
    procedure FormCreate(Sender: TObject);
{    procedure FDScriptSpoolPut(AEngine: TFDScript; const AMessage: string;
      AKind: TFDScriptOutputKind);}
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Fechar1Click(Sender: TObject);
    procedure Abrir1Click(Sender: TObject);
    procedure btnExecutarProcessoClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TimerTimer(Sender: TObject);
    procedure btTestarScriptsClick(Sender: TObject);
    procedure FDScriptSpoolPut(AEngine: TFDScript; const AMessage: string; AKind: TFDScriptOuputKind);
  private
    { Private declarations }
    FProcessando: Boolean;
    FEmTesteScripts: Boolean;
    FPendingExecLine: string;
    FPendingErrorLine: string;
    FPendingTabErrosVisible: Boolean;
    FPendingLabelExecutados: string;
    FPendingLabelErros: string;
    FPendingGaugeMax: Integer;
    FPendingGaugeProgress: Integer;
    FPendingGaugeScriptsMax: Integer;
    FPendingGaugeScriptsProgress: Integer;
    FLastScriptsMax: Integer;
    FPendingProcVisible: Boolean;
    FPendingProcEtapa: string;
    FPendingProcDesc: string;
    FLastScriptsProgress: Integer;
    FPendingButtonsEnabled: Boolean;
    FPendingErrorMemoText: string;
    FBackupFilePath: string;
    FBackupProgressCapBytes: Int64;
    FBackupProgressCapInt: Integer;
    procedure DoAddExecLine;
    procedure DoAddErrorLine;
    procedure DoClearExecMemo;
    procedure DoSetTabErrosVisible;
    procedure DoSetLabelExecutados;
    procedure DoSetLabelErros;
    procedure DoSetGauge;
    procedure DoSetGaugeScripts;
    procedure DoSetProcessPanelVisible;
    procedure DoSetProcessText;
    procedure DoSetBotoesEnabled;
    procedure DoGetErrorMemoText;
    procedure FinishProcessThread;
    function GaugeValueFromInt64(const Value: Int64): Integer;
    procedure AddExecLine(const Msg: string);
    procedure AddErrorLine(const Msg: string);
    procedure ClearExecMemo;
    procedure SetTabErrosVisible(const AVisible: Boolean);
    procedure SetLabelExecutados(const Msg: string);
    procedure SetLabelErros(const Msg: string);
    procedure SetGauge(const AMax, AProgress: Integer);
    procedure SetGaugeScripts(const AMax, AProgress: Integer);
    procedure SetProcessPanelVisible(const AVisible: Boolean);
    procedure SetProcessStep(const AStep, ATotal: Integer; const ADesc: string);
    procedure CenterProcessPanel;
    procedure SetBotoesEnabled(const AEnabled: Boolean);
    function GetErrorMemoText: string;
    procedure StartProcessThread(const AKind: TProcessKind; const PastaScripts: string);
    function ExecuteSQLCommands: boolean;
    function FTPConectarBaixar(pVersao: string): Boolean;
    function VerificarVersaoLiberada: string;
    procedure ProcessarAtualizacao;
    function Conectar: boolean;
    function GetComandosVersao: Boolean;
    function RenomearExe: boolean;
    function DescompactarArquivo: boolean;
    procedure ReportarResultado;
    function FecharTodosExecutaveisAbertos: Boolean;
    function TestarScriptsPasta(const Pasta: string): Boolean;
    function RunExternalProcess(const Executable, Params: string; out ExitCode: DWORD; const ProgressProc: TProc = nil): Boolean;
    procedure UpdateBackupProgress(const Desc: string; const Progress: Int64);
    procedure TryRestoreBackup(const Reason: string);
    function CopyFileWin(const SourceFileName, DestFileName: string): Boolean;

    function BuildFirebirdConnectionString: string;
    function ExtractDatabaseFilePath(const ConnectionString: string): string;
    function GetFileSizeSafe(const APath: string): Int64;
    function GetBackupMode: string;
    function BackupFirebirdDatabase: Boolean;
    procedure RestoreFirebirdDatabase;
  public    { Public declarations }
    FDConnection: TFDConnection; // Declare FDConnection aqui
  end;

var
  MainForm: TMainForm;
  DirArqBaixados: string;
  VersaoAtualizar: string;
  VersaoAtual: string;
  ComandosVersao: array of TComandosSQL;
  Origem: string;
  Destino: string;
  LocalZip: string;
  TotaScripts: integer;
  TotalErros: integer;
  NoVersaoCli: integer;
  VersaoSQL:string;

implementation

uses
  DMConnec;



type
  TProcessThread = class(TThread)
  private
    FOwner: TMainForm;
    FKind: TProcessKind;
    FPastaScripts: string;
  protected
    procedure Execute; override;
  public
    constructor Create(AOwner: TMainForm; AKind: TProcessKind; const APastaScripts: string);
  end;

constructor TProcessThread.Create(AOwner: TMainForm; AKind: TProcessKind; const APastaScripts: string);
begin
  inherited Create(True);
  FreeOnTerminate := True;
  FOwner := AOwner;
  FKind := AKind;
  FPastaScripts := APastaScripts;
end;

procedure TProcessThread.Execute;
begin
  try
    case FKind of
      pkAtualizacao:
        FOwner.ProcessarAtualizacao;
      pkTesteScripts:
        FOwner.TestarScriptsPasta(FPastaScripts);
    end;
  finally
    TThread.Synchronize(nil, FOwner.FinishProcessThread);
  end;
end;
{$R *.dfm}

{ TMainForm }

{
  1) Processo de Baixar os Pacotes (Zip) de Atualização;
  2) Processo de Descompactar o Arquivo de Pacotes;
  3) Leitura e Execução do Scripts de Banco de dados; [OK]
  3) Substituição dos Executáveis e DLL's
}

procedure TMainForm.ProcessarAtualizacao;
begin
  SetProcessPanelVisible(True);
  SetProcessStep(1, 8, 'Lendo configuracoes');
  SetGauge(100, 0);
  Origem := LerConf('ARQUIVOS', 'Origem');
  Destino := LerConf('ARQUIVOS', 'Destino');
  LocalZip := LerConf('ARQUIVOS', 'LocalZip');
  SetGauge(100, 12);
  SetProcessStep(2, 8, 'Conectando ao banco');

  if Conectar then
  begin
    SetGauge(100, 25);
    SetProcessStep(3, 8, 'Verificando versao');
    VersaoAtualizar := VerificarVersaoLiberada;
    if VersaoAtualizar <> '' then
    begin
      SetGauge(100, 37);
      SetProcessStep(4, 8, 'Fechando executaveis');
      TotalErros := 0;
      if FecharTodosExecutaveisAbertos then
      begin
        SetGauge(100, 50);
        SetProcessStep(5, 8, 'Buscando comandos SQL');
        if GetComandosVersao then
        begin
          UpdateBackupProgress('Backup do banco (0%)', 0);
          FBackupFilePath := '';
          if BackupFirebirdDatabase then
          begin
            UpdateBackupProgress('Backup do banco concluído', 100);
            SetGauge(100, 65);
            SetProcessStep(6, 8, 'Executando comandos SQL');
            if ExecuteSQLCommands then
            begin
              SetGauge(100, 75);
              SetProcessStep(7, 8, 'Baixando pacote');
              if FTPConectarBaixar(VersaoAtualizar) then
              begin
                SetGauge(100, 87);
                SetProcessStep(8, 8, 'Aplicando arquivos');
                if RenomearExe then
                  if DescompactarArquivo then
                  begin
                    SetGauge(100, 100);
                    SetProcessStep(8, 8, 'Concluido');
                  end;
              end;
            end
            else
            begin
            TryRestoreBackup('erro na execução dos scripts SQL');
            end;
          end
          else
          begin
            AddErrorLine('Não foi possível criar o backup do banco antes dos scripts.');
            TotalErros := TotalErros + 1;
          end;
        end
        else
        begin
          AddErrorLine('Não tem Comandos de Script de Atualização nesta Versão');
          TotalErros := TotalErros + 1;
        end;
      end
      else
      begin
        AddErrorLine('Não foi possível fechar todos os executáveis.');
        TotalErros := TotalErros + 1;
      end;
      if (FBackupFilePath <> '') and (TotalErros > 0) then
        TryRestoreBackup('erro no processo');
      ReportarResultado;
    end;
  end;

end;

procedure TMainForm.AddExecLine(const Msg: string);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    AExecutarMemo.Lines.Add(Msg)
  else
  begin
    FPendingExecLine := Msg;
    TThread.Synchronize(nil, DoAddExecLine);
  end;
end;

procedure TMainForm.DoAddExecLine;
begin
  AExecutarMemo.Lines.Add(FPendingExecLine);
end;

procedure TMainForm.AddErrorLine(const Msg: string);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    AErrorMemo.Lines.Add(Msg)
  else
  begin
    FPendingErrorLine := Msg;
    TThread.Synchronize(nil, DoAddErrorLine);
  end;
end;

procedure TMainForm.DoAddErrorLine;
begin
  AErrorMemo.Lines.Add(FPendingErrorLine);
end;

procedure TMainForm.ClearExecMemo;
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    AExecutarMemo.Lines.Clear
  else
    TThread.Synchronize(nil, DoClearExecMemo);
end;

procedure TMainForm.DoClearExecMemo;
begin
  AExecutarMemo.Lines.Clear;
end;

procedure TMainForm.SetTabErrosVisible(const AVisible: Boolean);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    TabErros.TabVisible := AVisible
  else
  begin
    FPendingTabErrosVisible := AVisible;
    TThread.Synchronize(nil, DoSetTabErrosVisible);
  end;
end;

procedure TMainForm.DoSetTabErrosVisible;
begin
  TabErros.TabVisible := FPendingTabErrosVisible;
end;

procedure TMainForm.SetLabelExecutados(const Msg: string);
var
  LMsg: string;
begin
  LMsg := Msg;
  if FLastScriptsMax > 0 then
    LMsg := LMsg + '  Scripts: ' + IntToStr(FLastScriptsProgress) + '/' + IntToStr(FLastScriptsMax);

  if TThread.CurrentThread.ThreadID = MainThreadID then
    lbExecutados.Caption := LMsg
  else
  begin
    FPendingLabelExecutados := LMsg;
    TThread.Synchronize(nil, DoSetLabelExecutados);
  end;
end;

procedure TMainForm.DoSetLabelExecutados;
begin
  lbExecutados.Caption := FPendingLabelExecutados;
end;

procedure TMainForm.SetLabelErros(const Msg: string);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    lbErros.Caption := Msg
  else
  begin
    FPendingLabelErros := Msg;
    TThread.Synchronize(nil, DoSetLabelErros);
  end;
end;

procedure TMainForm.DoSetLabelErros;
begin
  lbErros.Caption := FPendingLabelErros;
end;

procedure TMainForm.SetGauge(const AMax, AProgress: Integer);
var
  LMax, LProgress: Integer;
begin
  LMax := AMax;
  LProgress := AProgress;
  if LMax < 1 then
    LMax := 1;
  if LProgress < 0 then
    LProgress := 0;
  if LProgress > LMax then
    LProgress := LMax;

  if TThread.CurrentThread.ThreadID = MainThreadID then
  begin
    GaugeProgresso.MaxValue := LMax;
    GaugeProgresso.Progress := LProgress;
  end
  else
  begin
    FPendingGaugeMax := LMax;
    FPendingGaugeProgress := LProgress;
    TThread.Synchronize(nil, DoSetGauge);
  end;
end;

procedure TMainForm.SetGaugeScripts(const AMax, AProgress: Integer);
var
  LMax, LProgress: Integer;
begin
  LMax := AMax;
  LProgress := AProgress;
  if LMax < 1 then
    LMax := 1;
  if LProgress < 0 then
    LProgress := 0;
  if LProgress > LMax then
    LProgress := LMax;

  FLastScriptsMax := LMax;
  FLastScriptsProgress := LProgress;

  if TThread.CurrentThread.ThreadID = MainThreadID then
  begin
    GaugeScripts.MaxValue := LMax;
    GaugeScripts.Progress := LProgress;
  end
  else
  begin
    FPendingGaugeScriptsMax := LMax;
    FPendingGaugeScriptsProgress := LProgress;
    TThread.Synchronize(nil, DoSetGaugeScripts);
  end;

  SetLabelExecutados('Comandos Executados...: ' + IntToStr(TotaScripts));
end;

procedure TMainForm.DoSetGauge;
begin
  GaugeProgresso.MaxValue := FPendingGaugeMax;
  GaugeProgresso.Progress := FPendingGaugeProgress;
end;

procedure TMainForm.DoSetGaugeScripts;
begin
  GaugeScripts.MaxValue := FPendingGaugeScriptsMax;
  GaugeScripts.Progress := FPendingGaugeScriptsProgress;
end;

procedure TMainForm.DoSetProcessPanelVisible;
begin
  PanelProcesso.Visible := FPendingProcVisible;
  if PanelProcesso.Visible then
    CenterProcessPanel;
end;

procedure TMainForm.DoSetProcessText;
begin
  lbProcEtapa.Caption := FPendingProcEtapa;
  lbProcDesc.Caption := FPendingProcDesc;
end;

procedure TMainForm.SetProcessPanelVisible(const AVisible: Boolean);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
  begin
    PanelProcesso.Visible := AVisible;
    if PanelProcesso.Visible then
      CenterProcessPanel;
  end
  else
  begin
    FPendingProcVisible := AVisible;
    TThread.Synchronize(nil, DoSetProcessPanelVisible);
  end;
end;

procedure TMainForm.SetProcessStep(const AStep, ATotal: Integer; const ADesc: string);
var
  LTotal, LStep: Integer;
begin
  LTotal := ATotal;
  LStep := AStep;
  if LTotal < 1 then
    LTotal := 1;
  if LStep < 0 then
    LStep := 0;
  if LStep > LTotal then
    LStep := LTotal;

  if TThread.CurrentThread.ThreadID = MainThreadID then
  begin
    lbProcEtapa.Caption := 'Etapa: ' + IntToStr(LStep) + '/' + IntToStr(LTotal);
    lbProcDesc.Caption := ADesc;
  end
  else
  begin
    FPendingProcEtapa := 'Etapa: ' + IntToStr(LStep) + '/' + IntToStr(LTotal);
    FPendingProcDesc := ADesc;
    TThread.Synchronize(nil, DoSetProcessText);
  end;
end;

procedure TMainForm.CenterProcessPanel;
begin
  if PanelProcesso.Parent <> nil then
  begin
    PanelProcesso.Left := (PanelProcesso.Parent.ClientWidth - PanelProcesso.Width) div 2;
    PanelProcesso.Top := (PanelProcesso.Parent.ClientHeight - PanelProcesso.Height) div 2;
  end;
end;

function TMainForm.GaugeValueFromInt64(const Value: Int64): Integer;
begin
  if Value <= 0 then
    Result := 0
  else if Value > High(Integer) then
    Result := High(Integer)
  else
    Result := Value;
end;

procedure TMainForm.UpdateBackupProgress(const Desc: string; const Progress: Int64);
var
  GaugeMax, GaugeProgress: Integer;
begin
  SetProcessStep(6, 10, Desc);
  if FBackupProgressCapInt > 0 then
    GaugeMax := FBackupProgressCapInt
  else
    GaugeMax := 100;

  GaugeProgress := GaugeValueFromInt64(Progress);
  if GaugeProgress > GaugeMax then
    GaugeProgress := GaugeMax;
  SetGauge(GaugeMax, GaugeProgress);
end;

procedure TMainForm.TryRestoreBackup(const Reason: string);
begin
  if FBackupFilePath = '' then
    Exit;

  AddExecLine('Restaurando backup do banco após: ' + Reason);
  RestoreFirebirdDatabase;
end;

procedure TMainForm.SetBotoesEnabled(const AEnabled: Boolean);
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
  begin
    btnExecutarProcesso.Enabled := AEnabled;
    btTestarScripts.Enabled := AEnabled;
  end
  else
  begin
    FPendingButtonsEnabled := AEnabled;
    TThread.Synchronize(nil, DoSetBotoesEnabled);
  end;
end;

procedure TMainForm.DoSetBotoesEnabled;
begin
  btnExecutarProcesso.Enabled := FPendingButtonsEnabled;
  btTestarScripts.Enabled := FPendingButtonsEnabled;
end;

function TMainForm.GetErrorMemoText: string;
begin
  if TThread.CurrentThread.ThreadID = MainThreadID then
    Result := AErrorMemo.Text
  else
  begin
    FPendingErrorMemoText := '';
    TThread.Synchronize(nil, DoGetErrorMemoText);
    Result := FPendingErrorMemoText;
  end;
end;

procedure TMainForm.DoGetErrorMemoText;
begin
  FPendingErrorMemoText := AErrorMemo.Text;
end;

procedure TMainForm.StartProcessThread(const AKind: TProcessKind; const PastaScripts: string);
var
  Th: TProcessThread;
begin
  if FProcessando then
  begin
    AddExecLine('Processo em execução. Aguarde...');
    Exit;
  end;

  FProcessando := True;
  SetBotoesEnabled(False);
  Timer.Enabled := False;
  SetGauge(1, 0);
  SetGaugeScripts(1, 0);

  Th := TProcessThread.Create(Self, AKind, PastaScripts);
  Th.Resume;
end;

procedure TMainForm.FinishProcessThread;
begin
  FProcessando := False;
  SetProcessPanelVisible(False);
  SetBotoesEnabled(True);
  Timer.Enabled := True;
end;

function TMainForm.RenomearExe: boolean;
var
  BackupDestino: string;
begin
  Result := True;
  if not FileExists(Destino) then
  begin
    try
      Result := RenameFile(Origem, Destino);
    except
      on E: Exception do
      begin
        AddErrorLine('Erro ao renomear arquivo: ' + E.Message);
        TotalErros := TotalErros + 1;
        Result := False;
      end;
    end;
  end
  else
  begin
    try
      BackupDestino := Destino + '.bak';
      if FileExists(BackupDestino) then
        DeleteFile(BackupDestino);
      if RenameFile(Destino, BackupDestino) then
        Result := RenameFile(Origem, Destino)
      else
        Result := False;
    except
      on E: Exception do
      begin
        AddErrorLine('Erro ao preparar backup do executável: ' + E.Message);
        TotalErros := TotalErros + 1;
        Result := False;
      end;
    end;
  end;
end;

function TMainForm.DescompactarArquivo: boolean;
begin
  Result := False;
  try
    ExtrairRAR(LocalZip + VersaoAtualizar + '.rar', LocalZip);
    Result := true;
  except
    on E: Exception do
    begin
      AddErrorLine('Erro ao descompactar: ' + E.Message);
      TotalErros := TotalErros + 1;
    end;
  end;
end;

function TMainForm.Conectar: boolean;
begin
  Result := True;
  try
    dmconnec.dtmConnec.FDConnection.DriverName := 'FB';
    dmconnec.dtmConnec.FDConnection.Params.Clear;
    dmconnec.dtmConnec.FDConnection.Params.Add('DriverID=FB');
    dmconnec.dtmConnec.FDConnection.Params.Add('Protocol=TCPIP');
    dmconnec.dtmConnec.FDConnection.Params.Add('Database=' + LerConf('SERVIDOR', 'Database'));
    dmconnec.dtmConnec.FDConnection.Params.Add('User_Name=' + LerConf('SERVIDOR', 'User'));
    dmconnec.dtmConnec.FDConnection.Params.Add('Password=' + decrypt(LerConf('SERVIDOR', 'Password')));
    dmconnec.dtmConnec.FDConnection.Params.Add('Port=' + LerConf('SERVIDOR', 'Porta', '3050'));
    dmconnec.dtmConnec.FDConnection.Params.Add('Server=' + LerConf('SERVIDOR', 'IP'));
    dmconnec.dtmConnec.FDConnection.Params.Add('SQLDialect=' + LerConf('SERVIDOR', 'SQLDialect', '1'));
    dmconnec.dtmConnec.FDConnection.Params.Add('CharacterSet=' + LerConf('SERVIDOR', 'Charset', 'WIN1252'));
    dmconnec.dtmConnec.FDPhysFBDriverLink1.VendorLib := IncludeTrailingPathDelimiter(GetCurrentDir) + 'fbclient.dll';

    dmconnec.dtmConnec.FDConnection.Connected := True;

    dmconnec.dtmConnec.FDConnRemoto.Connected := True;

  except
    on E: Exception do
    begin
      AddErrorLine('Erro ao configurar conexão: ' + E.Message);
      TotalErros := TotalErros + 1;
      Result := False;
    end;
  end;
end;

procedure TMainForm.ReportarResultado;
begin
  dtmConnec.FDQryRemoto.Close;
  dtmConnec.FDQryRemoto.SQL.Clear;

  dtmConnec.FDQryRemoto.SQL.Text := 'UPDATE TABVERSAO_CLI SET ' + '  SUCESSO = :SUCESSO, ' + '  ERROS   = :ERROS ' + 'WHERE NOVERSAOCLI = :NOVERSAOCLI';
  if TotalErros > 0 then
  begin
    dtmConnec.FDQryRemoto.ParamByName('SUCESSO').AsString := 'F';
    dtmConnec.FDQryRemoto.ParamByName('ERROS').AsString := GetErrorMemoText;
  end
  else
  begin
    dtmConnec.FDQryRemoto.ParamByName('SUCESSO').AsString := 'T';
    dtmConnec.FDQryRemoto.ParamByName('ERROS').AsString := '';
  end;

  dtmConnec.FDQryRemoto.ParamByName('NOVERSAOCLI').AsInteger := NoVersaoCli;

  try
    if not dtmConnec.FDConnRemoto.InTransaction then
      dtmConnec.FDConnRemoto.StartTransaction;

    dtmConnec.FDQryRemoto.ExecSQL;
    dtmConnec.FDConnRemoto.Commit;
  except
    dtmConnec.FDConnRemoto.Rollback;
  end;

  dtmConnec.FDQuery.Close;
  dtmConnec.FDQuery.SQL.Clear;
  dtmConnec.FDQuery.SQL.Text := 'update tabatu t set t.desativar = "F"';
  try
    if not dtmConnec.FDConnection.InTransaction then
      dtmConnec.FDConnection.StartTransaction;
    dtmConnec.FDQuery.ExecSQL;
    dtmConnec.FDConnection.Commit;
  except
    dtmConnec.FDConnection.Rollback;
  end;

end;

procedure TMainForm.TimerTimer(Sender: TObject);
begin
  StartProcessThread(pkAtualizacao, '');
end;

function TMainForm.FecharTodosExecutaveisAbertos: Boolean;
const
  WaitSeconds = 660; // 11 minutos
var
  Remaining: Integer;
  Min, Sec: Integer;
begin
  Result := False;
  dtmConnec.FDQuery.Close;
  dtmConnec.FDQuery.SQL.Clear;
  dtmConnec.FDQuery.SQL.Text := 'update tabatu t set t.desativar = "T"';
  try
    if not dtmConnec.FDConnection.InTransaction then
      dtmConnec.FDConnection.StartTransaction;
    dtmConnec.FDQuery.ExecSQL;
    dtmConnec.FDConnection.Commit;

   { Remaining := WaitSeconds;
    while Remaining > 0 do
    begin
      if (Remaining = WaitSeconds) or (Remaining mod 5 = 0) then
      begin
        Min := Remaining div 60;
        Sec := Remaining mod 60;
        SetProcessStep(4, 8, Format('Aguardando fechamento (%2.2d:%2.2d)', [Min, Sec]));
      end;
      Sleep(1000);
      Dec(Remaining);
    end;}

    Result := True;
  except
    dtmConnec.FDConnection.Rollback;
  end;
end;

function TMainForm.VerificarVersaoLiberada: string;
var
  CNPJ: string;
begin
  Result := '';

  dtmConnec.FDQuery.Close;
  dtmConnec.FDQuery.SQL.Clear;
  dtmConnec.FDQuery.SQL.Text := 'SELECT FIRST 1 V.VERSAO, E.CGCEMP '+
                                'FROM TABVER V, TABEMP E '+
                                'WHERE E.PADRAO = ''T'' ORDER BY V.NOVER DESC';
  dtmConnec.FDQuery.Open;
  CNPJ   := dtmConnec.FDQuery.FieldByName('CGCEMP').AsString;
  VersaoAtual := dtmConnec.FDQuery.FieldByName('VERSAO').AsString;

  if Pos('R', VersaoAtual) = 0 then
    VersaoAtual := 'R'+StringReplace(VersaoAtual, '.', '', [rfReplaceAll]);

  CNPJ :=  StringReplace(StringReplace(StringReplace(CNPJ, '.', '', [rfReplaceAll]), '/', '', [rfReplaceAll]), '-', '', [rfReplaceAll]);

  dtmConnec.FDQryRemoto.Close;
  dtmConnec.FDQryRemoto.SQL.Clear;
  dtmConnec.FDQryRemoto.SQL.Text := 'SELECT FIRST 1 ' +
                                    '  TABVERSAO_CLI.NOVERSAOCLI, ' +
                                    '  TABVERSAO_CLI.CNPJ, ' +
                                    '  TABVERSAO_CLI.NOMCLI, ' +
                                    '  TABVERSAO_CLI.VERSAO, ' +
                                    '  TABVERSAO_CLI.ATUALIZAR ' +
                                    'FROM TABVERSAO_CLI ' +
                                    'WHERE TABVERSAO_CLI.CNPJ = :CNPJ ' +
                                    '  AND TABVERSAO_CLI.ATUALIZAR = "T" ' +
                                    'ORDER BY TABVERSAO_CLI.VERSAO DESC';
  dtmConnec.FDQryRemoto.ParamByName('CNPJ').AsString := CNPJ;
  dtmConnec.FDQryRemoto.Open;

  if not dtmConnec.FDQryRemoto.IsEmpty then
  begin
    NoVersaoCli := dtmConnec.FDQryRemoto.FieldByName('NOVERSAOCLI').AsInteger;
    if Pos('R', dtmConnec.FDQryRemoto.FieldByName('VERSAO').AsString) = 0 then
      Result := 'R'+StringReplace(dtmConnec.FDQryRemoto.FieldByName('VERSAO').AsString, '.', '', [rfReplaceAll])
    else
      Result := dtmConnec.FDQryRemoto.FieldByName('VERSAO').AsString;
  end;
end;

function TMainForm.RunExternalProcess(const Executable, Params: string; out ExitCode: DWORD; const ProgressProc: TProc = nil): Boolean;
var
  StartupInfo: TStartupInfo;
  ProcInfo: TProcessInformation;
  CommandLine: string;
  WaitResult: DWORD;
begin
  Result := False;
  ExitCode := DWORD(-1);
  CommandLine := Format('"%s" %s', [Executable, Params]);
  FillChar(StartupInfo, SizeOf(StartupInfo), 0);
  StartupInfo.cb := SizeOf(StartupInfo);
  if not CreateProcess(nil, PChar(CommandLine), nil, nil, False, CREATE_NO_WINDOW, nil, nil, StartupInfo, ProcInfo) then
    Exit;
  try
    repeat
      WaitResult := WaitForSingleObject(ProcInfo.hProcess, 500);
      if WaitResult = WAIT_OBJECT_0 then
      begin
        if not GetExitCodeProcess(ProcInfo.hProcess, ExitCode) then
          ExitCode := DWORD(-1);
        Break;
      end
      else if WaitResult = WAIT_FAILED then
      begin
        ExitCode := DWORD(-1);
        Break;
      end
      else if Assigned(ProgressProc) then
        ProgressProc;
    until False;
    Result := ExitCode = 0;
  finally
    CloseHandle(ProcInfo.hProcess);
    CloseHandle(ProcInfo.hThread);
  end;
end;

function TMainForm.BuildFirebirdConnectionString: string;
var
  IP, Port, Database: string;
begin
  Database := Trim(LerConf('SERVIDOR', 'Database'));
  IP := Trim(LerConf('SERVIDOR', 'IP'));
  Port := Trim(LerConf('SERVIDOR', 'Porta', '3050'));
  if IP = '' then
    Exit(Database);
  if Port = '' then
    Result := Format('%s:%s', [IP, Database])
  else
    Result := Format('%s/%s:%s', [IP, Port, Database]);
end;

function TMainForm.ExtractDatabaseFilePath(const ConnectionString: string): string;
var
  PosDelim: Integer;
begin
  Result := Trim(ConnectionString);
  if Result = '' then
    Exit;

  PosDelim := Pos(':', Result);
  if PosDelim > 0 then
  begin
    Result := Copy(Result, PosDelim + 1, MaxInt);
    Result := Trim(Result);
  end;
end;

function TMainForm.GetFileSizeSafe(const APath: string): Int64;
var
  SearchRec: TSearchRec;
begin
  Result := 0;
  if (APath = '') or not FileExists(APath) then
    Exit;

  if FindFirst(APath, faAnyFile, SearchRec) = 0 then
  try
    Result := (Int64(SearchRec.FindData.nFileSizeHigh) shl 32) or SearchRec.FindData.nFileSizeLow;
  finally
    FindClose(SearchRec);
  end;
end;

function TMainForm.CopyFileWin(const SourceFileName, DestFileName: string): Boolean;
begin
  Result := Winapi.Windows.CopyFile(PChar(SourceFileName), PChar(DestFileName), False);
end;

function TMainForm.GetBackupMode: string;
begin
  Result := LowerCase(Trim(LerConf('BACKUP', 'Modo', 'gbak')));
end;
function TMainForm.BackupFirebirdDatabase: Boolean;
var
  BackupDir, BackupFile, Params, Source, SourceFilePath, GbakPath, User, Password: string;
  ExitCode: DWORD;
  SourceSize: Int64;
  ProgressLoop: TProc;
  BackupMode: string;
begin
  Result := False;
  BackupDir := IncludeTrailingPathDelimiter(LerConf('BACKUP', 'Diretorio', IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName)) + 'Backups'));
  ForceDirectories(BackupDir);
  BackupMode := GetBackupMode;

  if BackupMode = 'copiar_gdb' then
    BackupFile := BackupDir + Format('GDBBackup_%s.gdb', [FormatDateTime('yyyymmdd_hhnnss', Now)])
  else
    BackupFile := BackupDir + Format('FBBackup_%s.fbk', [FormatDateTime('yyyymmdd_hhnnss', Now)]);

  Source := BuildFirebirdConnectionString;
  if Source = '' then
  begin
    AddErrorLine('Não foi possível localizar o banco para backup.');
    Exit;
  end;
  SourceFilePath := ExtractDatabaseFilePath(Source);
  SourceSize := GetFileSizeSafe(SourceFilePath);
  FBackupProgressCapBytes := SourceSize - ((SourceSize * 30) div 100);
  if FBackupProgressCapBytes < 1 then
    FBackupProgressCapBytes := 1;
  FBackupProgressCapInt := GaugeValueFromInt64(FBackupProgressCapBytes);
  if FBackupProgressCapInt = 0 then
    FBackupProgressCapInt := 1;
  AddExecLine('Criando backup do banco (' + BackupFile + ')...');
  UpdateBackupProgress('Backup do banco em andamento 0%', 0);
  if BackupMode = 'copiar_gdb' then
  begin
    if not CopyFileWin(SourceFilePath, BackupFile) then
    begin
      AddErrorLine('Falha ao copiar o banco para backup: ' + SysErrorMessage(GetLastError));
      Exit;
    end;
    UpdateBackupProgress('Backup do banco conclu?do', FBackupProgressCapBytes);
    FBackupProgressCapBytes := 0;
    FBackupProgressCapInt := 0;
    FBackupFilePath := BackupFile;
    Result := True;
    Exit;
  end;
  GbakPath := LerConf('BACKUP', 'GbakPath', 'gbak');
  User := LerConf('SERVIDOR', 'User', 'SYSDBA');
  Password := decrypt(LerConf('SERVIDOR', 'Password'));
  Params := Format('-b -user "%s" -password "%s" "%s" "%s"', [User, Password, Source, BackupFile]);
  ProgressLoop := procedure
  var
    BackupSize, TrackedProgress: Int64;
    PercentDone: Integer;
  begin
    BackupSize := GetFileSizeSafe(BackupFile);
    PercentDone := 0;
    if SourceSize > 0 then
      PercentDone := Min(Trunc((BackupSize * 100) / SourceSize), 99);
    TrackedProgress := BackupSize;
    if FBackupProgressCapBytes > 0 then
      TrackedProgress := Min(FBackupProgressCapBytes, TrackedProgress);
    UpdateBackupProgress(Format('Backup do banco em andamento %d%%', [PercentDone]), TrackedProgress);
  end;
  if not RunExternalProcess(GbakPath, Params, ExitCode, ProgressLoop) then
  begin
    AddErrorLine(Format('Falha ao criar backup do banco (gbak exit=%d).', [ExitCode]));
    Exit;
  end;
  UpdateBackupProgress('Backup do banco concluído', FBackupProgressCapBytes);
  FBackupProgressCapBytes := 0;
  FBackupProgressCapInt := 0;
  FBackupFilePath := BackupFile;
  Result := True;
end;
procedure TMainForm.RestoreFirebirdDatabase;
var
  Params, GbakPath, User, Password, Destination, DestinationFilePath: string;
  ExitCode: DWORD;
  TrackSize, TargetSize: Int64;
  BackupMode: string;
  ProgressLoop: TProc;
begin
  if (FBackupFilePath = '') or not FileExists(FBackupFilePath) then
  begin
    AddErrorLine('Não há backup válido para restaurar.');
    Exit;
  end;
  Destination := BuildFirebirdConnectionString;
  if Destination = '' then
  begin
    AddErrorLine('Não foi possível localizar o banco para restauração.');
    Exit;
  end;
  DestinationFilePath := ExtractDatabaseFilePath(Destination);
  BackupMode := GetBackupMode;

  if dtmConnec.FDQuery.Active then
    dtmConnec.FDQuery.Close;
  if dtmConnec.FDQryRemoto.Active then
    dtmConnec.FDQryRemoto.Close;
  if dtmConnec.FDConnection.Connected then
    dtmConnec.FDConnection.Connected := False;
  if dtmConnec.FDConnRemoto.Connected then
    dtmConnec.FDConnRemoto.Connected := False;
  TrackSize := GetFileSizeSafe(DestinationFilePath);
  TargetSize := GetFileSizeSafe(FBackupFilePath);
  if TrackSize <= 0 then
    TrackSize := TargetSize;
  FBackupProgressCapBytes := TrackSize - ((TrackSize * 30) div 100);
  if FBackupProgressCapBytes < 1 then
    FBackupProgressCapBytes := 1;
  FBackupProgressCapInt := GaugeValueFromInt64(FBackupProgressCapBytes);
  if FBackupProgressCapInt = 0 then
    FBackupProgressCapInt := 1;
  if BackupMode = 'copiar_gdb' then
  begin
    if DestinationFilePath = '' then
    begin
      AddErrorLine('N?o foi poss?vel determinar o caminho f?sico do banco para restaura??o.');
      Exit;
    end;
    AddExecLine('Restaurando banco por c?pia de arquivo (' + FBackupFilePath + ')...');
    UpdateBackupProgress('Restaurando backup do banco', 0);
    if not CopyFileWin(FBackupFilePath, DestinationFilePath) then
    begin
      AddErrorLine('Erro ao restaurar banco por c?pia: ' + SysErrorMessage(GetLastError));
      Exit;
    end;
    UpdateBackupProgress('Backup restaurado com sucesso', FBackupProgressCapBytes);
    AddExecLine('Banco restaurado com sucesso.');
    FBackupProgressCapBytes := 0;
    FBackupProgressCapInt := 0;
    FBackupFilePath := '';
    Conectar;
    Exit;
  end;


  GbakPath := LerConf('BACKUP', 'GbakPath', 'gbak');
  User := LerConf('SERVIDOR', 'User', 'SYSDBA');
  Password := decrypt(LerConf('SERVIDOR', 'Password'));
  Params := Format('-c -user "%s" -password "%s" -rep "%s" "%s"', [User, Password, FBackupFilePath, Destination]);
  AddExecLine('Restaurando banco a partir do backup (' + FBackupFilePath + ')...');
  UpdateBackupProgress('Restaurando backup do banco', 0);
  ProgressLoop := procedure
  var
    CurrentSize, TrackedProgress: Int64;
    PercentDone: Integer;
  begin
    CurrentSize := GetFileSizeSafe(DestinationFilePath);
    PercentDone := 0;
    if TargetSize > 0 then
      PercentDone := Min(Trunc((CurrentSize * 100) / TargetSize), 99);
    TrackedProgress := CurrentSize;
    if FBackupProgressCapBytes > 0 then
      TrackedProgress := Min(FBackupProgressCapBytes, TrackedProgress);
    UpdateBackupProgress(Format('Restaurando backup do banco %d%%', [PercentDone]), TrackedProgress);
  end;
  if not RunExternalProcess(GbakPath, Params, ExitCode, ProgressLoop) then
  begin
    AddErrorLine(Format('Erro ao restaurar backup do banco (gbak exit=%d).', [ExitCode]));
    Exit;
  end;
  UpdateBackupProgress('Backup restaurado com sucesso', FBackupProgressCapBytes);
  AddExecLine('Banco restaurado com sucesso.');
  FBackupProgressCapBytes := 0;
  FBackupProgressCapInt := 0;
  FBackupFilePath := '';
  Conectar;
end;

function TMainForm.GetComandosVersao: Boolean;
var
  CNPJ: string;
  x, y: Integer;
begin
  Result := True;
  dtmConnec.FDQryRemoto.Close;
  dtmConnec.FDQryRemoto.SQL.Clear;
  dtmConnec.FDQryRemoto.SQL.Text := 'SELECT TABVERSAO.COMANDOS, TABVERSAO.VERSAO FROM TABVERSAO WHERE TABVERSAO.VERSAO BETWEEN :VERSAOATUAL AND :VERSAO ORDER BY TABVERSAO.VERSAO ';
  dtmConnec.FDQryRemoto.ParamByName('VERSAOATUAL').AsString := VersaoAtual;
  dtmConnec.FDQryRemoto.ParamByName('VERSAO').AsString      := VersaoAtualizar;


  x := 0;
  SetLength(ComandosVersao, x);

  try
    dtmConnec.FDQryRemoto.Open;
    dtmConnec.FDQryRemoto.First;
    while not dtmConnec.FDQryRemoto.Eof do
    begin
      x := x + 1;
      SetLength(ComandosVersao, x);
      ComandosVersao[x-1].Versao := dtmConnec.FDQryRemoto.FieldByName('VERSAO').AsString;
      ComandosVersao[x-1].ComandoSQL := dtmConnec.FDQryRemoto.FieldByName('COMANDOS').AsString;
      dtmConnec.FDQryRemoto.Next;
    end;
  except
    on E: Exception do
    begin
      Result := False;
      AddExecLine(Format('Erro ao Buscar Comandos SQL da Versão. Erro: %s', [E.Message]));
      TotalErros := TotalErros + 1;
    end;
  end;
end;

function TMainForm.ExecuteSQLCommands: boolean;
var
  i,y: Integer;
  SQLCommand, SQLAux: string;
  Query: TFDQuery;
  ASQLCommands, LineBuffer: TStringList;
  SQLScript: TFDSQLScript;
  IniciaComSetTerm: Boolean;
  function IsImportantCommand(const SQL: string): Boolean;
  begin
    // Verifica se o comando é importante para ser executado
    Result := not SQL.StartsWith('CONNECT', True) and not SQL.StartsWith('RECONNECT', True) and not SQL.StartsWith('SET AUTODDL', True);

    if not IniciaComSetTerm then
      IniciaComSetTerm := SQL.StartsWith('SET TERM ^ ;', True);
  end;

  function IsCommandComplete(const SQL: string): Boolean;
  begin
    // Verifica se o comando está completo (termina com ; ou ^)
    if IniciaComSetTerm then
    begin
      Result := pos('SET TERM ; ^', SQL) > 0;

      IniciaComSetTerm := not Result;
    end
    else
      Result := SQL.EndsWith(';');
  end;

  function AjustaSPs(SQL: string): string;
  begin
    if IniciaComSetTerm then
    begin
      if not SQL.EndsWith('SET TERM ; ^') then
      begin
        if SQL.EndsWith('^') then
          SQL := SQL + #13#10 + 'SET TERM ; ^' + #13#10;
      end;
    end;

    if SQL.StartsWith('ALTER PROCEDURE', True) or SQL.StartsWith('ALTER TRIGGER', True) or SQL.StartsWith('CREATE TRIGGER', True) then
    begin
      if SQL.StartsWith('CREATE TRIGGER', True) then
        SQL := StringReplace(SQL, 'CREATE', 'CREATE OR ALTER', [rfReplaceAll]);

      SQL := 'SET TERM ^ ;' + #13#10 + SQL;
      IniciaComSetTerm := True;
    end;

    Result := SQL;
  end;

begin
  Result := True;
  SetTabErrosVisible(False);
  IniciaComSetTerm := False;
  ClearExecMemo;
  AddExecLine('INICIANDO EXECUÇÃO DOS COMANDOS SQL...');
  AddExecLine('');
  LineBuffer := TStringList.Create;
  ASQLCommands := TStringList.Create;
  try

    //ComandosVersao :=
//    ASQLCommands.Text := stringreplace(ComandosVersao, 'end' + #$D#$A + '^', 'end^', [rfReplaceAll]);

    TotaScripts := 0;
    if not FEmTesteScripts then SetGaugeScripts(Length(ComandosVersao), 0);

    for y := 0 to Length(ComandosVersao)-1 do
    begin
      VersaoSQL := ComandosVersao[y].Versao;
      ASQLCommands.Text := stringreplace(ComandosVersao[y].ComandoSQL, 'end' + #$D#$A + '^', 'end^', [rfReplaceAll]);

      for i := 0 to ASQLCommands.Count - 1 do
      begin
        SQLCommand := ASQLCommands[i].Trim;

        if SQLCommand.IsEmpty or SQLCommand.StartsWith('--') or not IsImportantCommand(SQLCommand) then
        begin
          Continue; // Ignorar linhas vazias, comentários ou comandos não importantes
        end;

        SQLAux := AjustaSPs(SQLCommand);
        SQLCommand := SQLAux;

        if not IsCommandComplete(SQLCommand) then
        begin
          LineBuffer.Add(SQLCommand);
          Continue;
        end;

        // Finalizar comando no ";" ou "^" e executar
        LineBuffer.Add(SQLCommand);
        SQLCommand := LineBuffer.Text;
        LineBuffer.Clear; // Limpar buffer para próximo comando

        try
          FDScript.SQLScripts.Clear;
          SQLScript := FDScript.SQLScripts.Add;

          FDScript.ScriptOptions.BreakOnError := False;
          FDScript.ScriptOptions.IgnoreError := False;
          FDScript.ScriptOptions.Verify := False;
          SQLScript.Sql.Text := SQLCommand;
          if FDScript.ExecuteAll then
          begin
            AddExecLine('');
            AddExecLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
            AddExecLine('/*                                                EXECUTANDO O COMANDO                                                      */');
            AddExecLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
            AddExecLine(SQLCommand);
            AddExecLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
            AddExecLine('/*                                                      RESULTADO                                                           */');
            AddExecLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
            AddExecLine('COMANDO EXECUTADO COM SUCESSO');
            AddExecLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
          end;
          TotaScripts := TotaScripts + 1;
        except
          on E: Exception do
          begin
            AddExecLine(Format('ERRO: %s', [E.Message]));
            TotalErros := TotalErros + 1;
          end;
        end;
      end;
      if not FEmTesteScripts then SetGaugeScripts(Length(ComandosVersao), y + 1);
    end;
  finally
    LineBuffer.Free;
    ASQLCommands.Free;
  end;
  AddExecLine('');
  AddExecLine('EXECUÇÃO CONCLUÍDA.');

  SetLabelExecutados('Comandos Executados...: ' + IntToStr(TotaScripts));
  SetLabelErros('Comandos Erros........: ' + IntToStr(TotalErros));

  SetTabErrosVisible(TotalErros > 0);
  Result := TotalErros = 0;
end;

function TMainForm.TestarScriptsPasta(const Pasta: string): Boolean;
var
  Arquivos: TStringList;
  i: Integer;
  ArquivoAtual: string;
begin
  FEmTesteScripts := True;
  Result := False;
  TotalErros := 0;

  if not TDirectory.Exists(Pasta) then
  begin
    AddErrorLine('Pasta de scripts nao encontrada: ' + Pasta);
    TotalErros := TotalErros + 1;
    Exit;
  end;

  if not Conectar then
  begin
    AddErrorLine('Nao foi possivel conectar usando o atualizador.conf.');
    TotalErros := TotalErros + 1;
    Exit;
  end;

  Arquivos := TStringList.Create;
  try
    for ArquivoAtual in TDirectory.GetFiles(Pasta, '*.txt') do
      Arquivos.Add(ArquivoAtual);
    Arquivos.Sort;

    if Arquivos.Count = 0 then
    begin
      AddErrorLine('Nenhum script encontrado na pasta: ' + Pasta);
      TotalErros := TotalErros + 1;
      Exit;
    end;

    SetGaugeScripts(Arquivos.Count, 0);

    SetLength(ComandosVersao, Arquivos.Count);
    for i := 0 to Arquivos.Count - 1 do
    begin
      ArquivoAtual := Arquivos[i];
      ComandosVersao[i].Versao := ArquivoAtual;
      ComandosVersao[i].ComandoSQL := TFile.ReadAllText(ArquivoAtual, TEncoding.ANSI);
      SetGaugeScripts(Arquivos.Count, i + 1);
      Result := ExecuteSQLCommands;
    end;
  finally
    FEmTesteScripts := False;
    Arquivos.Free;
  end;
end;

procedure TMainForm.btTestarScriptsClick(Sender: TObject);
var
  PastaScripts: string;
begin
  TotalErros := 0;
  TotaScripts := 0;
  ClearExecMemo;
  AErrorMemo.Clear;
  SetLabelExecutados('Comandos Executados...: 0');
  SetLabelErros('Comandos Erros........: 0');
  SetTabErrosVisible(False);
  SetGaugeScripts(1, 0);
  PastaScripts := LerConf('ARQUIVOS', 'PastaScripts', LocalZip);
  if PastaScripts = '' then
    PastaScripts := ExtractFilePath(Application.ExeName);

  if SelectDirectory('Selecione a pasta dos scripts', '', PastaScripts) then
    StartProcessThread(pkTesteScripts, PastaScripts);
end;

procedure TMainForm.Abrir1Click(Sender: TObject);
begin
  Show;
end;

function TMainForm.FTPConectarBaixar(pVersao: string): Boolean;
var
  SSL: TIdSSLIOHandlerSocketOpenSSL;
  RemoteSize, LocalSize: Int64;
  LocalFile: string;

  function GetSize: Int64;
  var
    FS: TFileStream;
    FileSize: Int64;
  begin
    FS := TFileStream.Create(LocalFile, fmOpenRead or fmShareDenyNone);
    try
      FileSize := FS.Size;
    finally
      FS.Free;
    end;
    Result := FileSize;
  end;

begin
  Result := False;
  SetTabErrosVisible(false);
  SSL := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
  try
    // Configura o SSL/TLS
    SSL.SSLOptions.Method := sslvTLSv1_2;
    SSL.SSLOptions.Mode := sslmUnassigned;
    FTP.IOHandler := SSL;

  // Configurações do servidor FTP
    FTP.Host := 'ftp.mmsoftwares.com.br';
    FTP.Port := 21;
    FTP.Username := 'mmsoftwares';
    FTP.Password := 'res(MmSof@)123245';

  // Configurações adicionais
    FTP.UseTLS := utUseExplicitTLS;
    FTP.Passive := True;

    try
      FTP.Connect;

      if not FTP.Connected then
      begin
        AddErrorLine('Falha ao conectar ao FTP.');
        TotalErros := TotalErros + 1;
        SetTabErrosVisible(true);
        exit;
      end;

      // Entra na pasta "atualizador"
      FTP.ChangeDir('atualizador');

      if not TDirectory.Exists(LocalZip) then
        TDirectory.CreateDirectory(LocalZip);

      LocalFile := IncludeTrailingPathDelimiter(LocalZip) + pVersao + '.rar';
      RemoteSize := FTP.Size(pVersao + '.rar');

      // Baixa o arquivo do FTP
      if RemoteSize > 0 then
      begin
        try
          FTP.Get(pVersao + '.rar', LocalZip + pVersao + '.rar', True); // True = sobrescreve se existir
        except
          on E: Exception do
          begin
            AddExecLine(Format('Erro ao Baixar o Arquivo no FTP. Erro: %s', [E.Message]));
            TotalErros := TotalErros + 1;
          end;
        end;
        // Após o OnWorkEnd ser chamado, podemos checar o tamanho
        LocalSize := GetSize;
        if LocalSize = RemoteSize then
          Result := True
        else
        begin
          AddErrorLine(Format('Erro: Tamanho incorreto (remoto=%d, local=%d)', [RemoteSize, LocalSize]));
          TotalErros := TotalErros + 1;
          SetTabErrosVisible(True);
          Result := False;
        end;
      end;

    except
      on E: Exception do
      begin
        AddErrorLine('Erro no FTP: ' + E.Message);
        SetTabErrosVisible(true);
        TotalErros := TotalErros + 1;
        Result := False;
      end;
    end;

    // Desconecta
    if FTP.Connected then
      FTP.Disconnect;

  finally
    SSL.Free;
  end;
end;

procedure TMainForm.btnExecutarProcessoClick(Sender: TObject);
begin
  TotalErros := 0;
  TotaScripts := 0;
  ClearExecMemo;
  AErrorMemo.Clear;
  SetLabelExecutados('Comandos Executados...: 0');
  SetLabelErros('Comandos Erros........: 0');
  SetTabErrosVisible(False);
  SetGaugeScripts(1, 0);
  StartProcessThread(pkAtualizacao, '');
end;

{procedure TMainForm.FDScriptSpoolPut(AEngine: TFDScript; const AMessage: string;
  AKind: TFDScriptOutputKind);
begin
  if AKind = soError then
  begin
    if (pos('already exists', AMessage) = 0) and (pos('store duplicate value', AMessage) = 0) and (pos('violation of PRIMARY or UNIQUE KEY', AMessage) = 0)   then
    begin
      AddErrorLine('');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('/*                                                EXECUTANDO O COMANDO                                                      */');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine(AEngine.SQLScripts[0].SQL.Text);
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('/*                                                      RESULTADO                                                           */');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('ERRO NO SCRIPT. MENSAGEM: ' + AMessage);
      TotalErros  := TotalErros  + 1;
    end;
  end;
end;
}
procedure TMainForm.FDScriptSpoolPut(AEngine: TFDScript; const AMessage: string; AKind: TFDScriptOuputKind);
begin
  if AKind = soError then
  begin
    if (pos('already exists', AMessage) = 0) and
       (pos('store duplicate value', AMessage) = 0) and
       (pos('violation of PRIMARY or UNIQUE KEY', AMessage) = 0) and
       (pos('Attempt to define a second PRIMARY KEY for the same table', AMessage) = 0) and
       (pos('DOMAIN RDB$32305', AMessage) = 0)  then
    begin
      AddErrorLine('');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('/*                                                EXECUTANDO O COMANDO NA VERSÃO '+VersaoSQL+'                              */');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine(AEngine.SQLScripts[0].SQL.Text);
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('/*                                                      RESULTADO                                                           */');
      AddErrorLine('/*--------------------------------------------------------------------------------------------------------------------------*/');
      AddErrorLine('ERRO NO SCRIPT. MENSAGEM: ' + AMessage);
      TotalErros := TotalErros + 1;
    end;
  end;
end;

procedure TMainForm.Fechar1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := False;
  Hide; // oculta a janela, mas app continua no tray
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  CenterProcessPanel;
  DirArqBaixados := ExtractFilePath(Application.ExeName) + 'arquivos\';
  TrayIcon.Icon := Application.Icon;  // usa o ícone do executável
  TrayIcon.Visible := True;
  if not Assigned(btTestarScripts.OnClick) then
    btTestarScripts.OnClick := btTestarScriptsClick;
end;

procedure TMainForm.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  // Detecta Ctrl + Espaço
  if (Key = VK_SPACE) and (ssCtrl in Shift) then
  begin
    btnExecutarProcesso.Visible := not btnExecutarProcesso.Visible;
    btTestarScripts.Visible := not btTestarScripts.Visible;
    Key := 0; // evita beep
  end;
end;

end.






















