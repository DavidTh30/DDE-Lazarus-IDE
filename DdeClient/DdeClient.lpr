program DdeClient;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}
  cthreads,
  {$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, uclimain, uddecli, uddeml, ufrmsnap;

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Title := 'DDE Client';
  Application.Scaled := True;
  Application.Initialize;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
