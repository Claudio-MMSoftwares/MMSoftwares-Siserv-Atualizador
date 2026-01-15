program Atualizador;

uses
  Vcl.Forms,
  Main in 'Main.pas' {MainForm},
  DMConnec in 'DMConnec.pas' {dtmConnec: TDataModule},
  uutil_rotinas in 'uutil_rotinas.pas',
  rar_metodos in 'rar_metodos.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TdtmConnec, dtmConnec);
  Application.Run;
end.
