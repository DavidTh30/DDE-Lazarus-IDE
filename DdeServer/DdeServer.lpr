program DdeServer;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, usrvmain, uddeml, uddesrv, ufrmsnap;

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Title := 'DDE Server';
  Application.Scaled := True;
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
