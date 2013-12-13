{$I DebugSettings.inc}
program StrtMnu2;

uses
  Forms,
  My_RunningAppUtil,
  MainU in 'MainU.pas' {frmMain},
  AppFuncU in 'AppFuncU.pas',
  Walpaper in 'Walpaper.pas';

{$R *.RES}

begin
  {$IFNDEF DEBUG}
  ReportMemoryLeaksOnShutdown := True;
  {$ENDIF}

  Application.Initialize;
  If {$IFDEF DEBUG}True{$ELSE}CheckForAppMutex('Wallpaper_app') = raNone{$ENDIF} then
  begin
    Application.CreateForm(TfrmMain, frmMain);
  {$IFNDEF DEBUG}
    Application.ShowMainForm := False;
  {$ENDIF}
    Application.Run;
  end;
end.
 