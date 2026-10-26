{ ============================================================================
  DDE Server - the simulated HV PLC as a DDE server (DDEML), built with
  normal Lazarus forms. Excel: =NLHV|DB_HMI!VoltAct_kV
    - service name, 10 points (Set 1) or 2 points (Set 2), decimal comma
    - item table: every topic!item with value, type, access and links
    - conversations: client program, topic, transactions, links
    - log of connects, links, pokes, executes (optional every request)
    - settings in DdeServer.ini next to the exe; --start starts at once
  ============================================================================ }
unit usrvmain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Menus, ExtCtrls, StdCtrls, Grids,
  ComCtrls, Clipbrd, IniFiles, uplcsim, uddesrv, ufrmsnap;

type

  { TMainForm }

  TMainForm = class(TForm)
    ButtonStart: TButton;
    ButtonStop: TButton;
    CheckAutoStart: TCheckBox;
    CheckDecComma: TCheckBox;
    CheckLogRequests: TCheckBox;
    ComboPoints: TComboBox;
    EditService: TEdit;
    GridConvs: TStringGrid;
    GridItems: TStringGrid;
    LabelPoints: TLabel;
    LabelService: TLabel;
    LabelState: TLabel;
    LabelStatus: TLabel;
    MainMenu1: TMainMenu;
    MemoLog: TMemo;
    MenuAbout: TMenuItem;
    MenuClearLog: TMenuItem;
    MenuCopyFormula: TMenuItem;
    MenuCopyValue: TMenuItem;
    MenuExit: TMenuItem;
    MenuFile: TMenuItem;
    MenuHelp: TMenuItem;
    MenuSep1: TMenuItem;
    MenuStart: TMenuItem;
    MenuStop: TMenuItem;
    PageControl1: TPageControl;
    PanelTop: TPanel;
    PopupItems: TPopupMenu;
    ShapeState: TShape;
    SplitterLog: TSplitter;
    TabConvs: TTabSheet;
    TabItems: TTabSheet;
    Timer1: TTimer;
    procedure ButtonStartClick(Sender: TObject);
    procedure ButtonStopClick(Sender: TObject);
    procedure CheckDecCommaChange(Sender: TObject);
    procedure CheckLogRequestsChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridItemsResize(Sender: TObject);
    procedure MenuAboutClick(Sender: TObject);
    procedure MenuClearLogClick(Sender: TObject);
    procedure MenuCopyFormulaClick(Sender: TObject);
    procedure MenuCopyValueClick(Sender: TObject);
    procedure MenuExitClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    FSim: TPlcSim;
    FServer: TDdeSimServer;
    FTicks: Integer;
    FSnapFile: string;               // --snapshot=<png> --after=<s>: save the window, exit
    FEndAt: QWord;
    procedure Log(const S: string);
    procedure DoStart;
    procedure DoStop;
    procedure ShowItems;
    procedure ShowConvs;
    procedure UpdateControls;
    function Formula(Row: Integer): string;
    procedure LoadIni;
    procedure SaveIni;
    function IniName: string;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

const
  MAX_LOG = 2000;

{ an Excel DDE name needs quotes when it is more than letters, digits and _ }
function XlName(const S: string): string;
var
  I: Integer;
begin
  for I := 1 to Length(S) do
    if not (S[I] in ['A'..'Z', 'a'..'z', '0'..'9', '_']) then
      Exit('''' + S + '''');
  Result := S;
end;

{ TMainForm }

procedure TMainForm.FormCreate(Sender: TObject);
begin
  MaskFPUExceptions;
  LoadIni;
  GridItems.Cells[0, 0] := 'Topic';
  GridItems.Cells[1, 0] := 'Item';
  GridItems.Cells[2, 0] := 'Value';
  GridItems.Cells[3, 0] := 'Type';
  GridItems.Cells[4, 0] := 'Access';
  GridItems.Cells[5, 0] := 'Links';
  GridConvs.Cells[0, 0] := 'Client';
  GridConvs.Cells[1, 0] := 'Topic';
  GridConvs.Cells[2, 0] := 'Connected';
  GridConvs.Cells[3, 0] := 'Transactions';
  GridConvs.Cells[4, 0] := 'Links';
  UpdateControls;
  if CheckAutoStart.Checked or Application.HasOption('start') then
    DoStart;
  if Application.HasOption('snapshot') then
  begin
    FSnapFile := Application.GetOptionValue('snapshot');
    FEndAt := GetTickCount64 + Round(StrToFloatDef(Application.GetOptionValue('after'), 3) * 1000);
  end;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if FSnapFile = '' then SaveIni;      // a test run keeps the user's settings
  DoStop;
end;

{ the Value column takes the width the others leave }
procedure TMainForm.GridItemsResize(Sender: TObject);
var
  I, W: Integer;
begin
  W := GridItems.ClientWidth - 2;
  for I := 0 to GridItems.ColCount - 1 do
    if I <> 2 then Dec(W, GridItems.ColWidths[I]);
  if W > 120 then GridItems.ColWidths[2] := W;
end;

procedure TMainForm.Log(const S: string);
begin
  MemoLog.Lines.BeginUpdate;
  try
    while MemoLog.Lines.Count > MAX_LOG do
      MemoLog.Lines.Delete(0);
    MemoLog.Lines.Add(FormatDateTime('hh:nn:ss.zzz  ', Now) + S);
  finally
    MemoLog.Lines.EndUpdate;
  end;
  MemoLog.SelStart := Length(MemoLog.Text);
end;

procedure TMainForm.DoStart;
var
  Err: string;
begin
  if FServer <> nil then Exit;
  if ComboPoints.ItemIndex = 1 then
    FSim := TPlcSim.Create(2)
  else
    FSim := TPlcSim.Create(10);
  FSim.Start;
  FServer := TDdeSimServer.Create(FSim);
  FServer.OnLog := @Log;
  FServer.DecimalComma := CheckDecComma.Checked;
  FServer.LogRequests := CheckLogRequests.Checked;
  if not FServer.Start(EditService.Text, Err) then
  begin
    Log('start failed: ' + Err);
    FreeAndNil(FServer);
    FreeAndNil(FSim);
  end
  else
    Log(Format('PLC simulation with %d curve points (Set %d); Excel: =%s|DB_HMI!VoltAct_kV',
      [FSim.Points, Ord(FSim.Points = 2) + 1, XlName(FServer.Service)]));
  ShowItems;
  UpdateControls;
end;

procedure TMainForm.DoStop;
begin
  if FServer = nil then Exit;
  FServer.Stop;
  FreeAndNil(FServer);
  FreeAndNil(FSim);
  GridConvs.RowCount := 1;
  UpdateControls;
end;

procedure TMainForm.ShowItems;
var
  I: Integer;
  It: TSrvItem;
  S: string;
begin
  if FServer = nil then
  begin
    GridItems.RowCount := 1;
    Exit;
  end;
  GridItems.BeginUpdate;
  try
    GridItems.RowCount := FServer.ItemCount + 1;
    for I := 0 to FServer.ItemCount - 1 do
    begin
      It := FServer.Item(I);
      GridItems.Cells[0, I + 1] := It.Topic;
      GridItems.Cells[1, I + 1] := It.Name;
      if not FServer.Value(It.Topic, It.Name, S) then S := '?';
      S := StringReplace(StringReplace(S, #13#10, '  |  ', [rfReplaceAll]), #9, '  ', [rfReplaceAll]);
      GridItems.Cells[2, I + 1] := S;
      GridItems.Cells[3, I + 1] := FServer.TypeText(It);
      if It.Writable then
        GridItems.Cells[4, I + 1] := 'read / poke'
      else
        GridItems.Cells[4, I + 1] := 'read';
      if FServer.LinksOf(It.Topic, It.Name) > 0 then
        GridItems.Cells[5, I + 1] := IntToStr(FServer.LinksOf(It.Topic, It.Name))
      else
        GridItems.Cells[5, I + 1] := '';
    end;
  finally
    GridItems.EndUpdate;
  end;
end;

procedure TMainForm.ShowConvs;
var
  I, J, N: Integer;
  C: TSrvConv;
begin
  if FServer = nil then Exit;
  GridConvs.BeginUpdate;
  try
    GridConvs.RowCount := FServer.ConvCount + 1;
    for I := 0 to FServer.ConvCount - 1 do
    begin
      C := FServer.Conv(I);
      N := 0;
      for J := 0 to FServer.LinkCount - 1 do
        if FServer.Link(J).Conv = C.Conv then Inc(N);
      GridConvs.Cells[0, I + 1] := C.Process;
      GridConvs.Cells[1, I + 1] := C.Topic;
      GridConvs.Cells[2, I + 1] := FormatDateTime('hh:nn:ss', C.Since);
      GridConvs.Cells[3, I + 1] := IntToStr(C.Transactions);
      GridConvs.Cells[4, I + 1] := IntToStr(N);
    end;
  finally
    GridConvs.EndUpdate;
  end;
end;

procedure TMainForm.UpdateControls;
var
  Run: Boolean;
begin
  Run := FServer <> nil;
  ButtonStart.Enabled := not Run;
  MenuStart.Enabled := not Run;
  ButtonStop.Enabled := Run;
  MenuStop.Enabled := Run;
  EditService.Enabled := not Run;
  ComboPoints.Enabled := not Run;
  if Run then
  begin
    ShapeState.Brush.Color := clLime;
    LabelState.Caption := 'running, service ' + FServer.Service;
  end
  else
  begin
    ShapeState.Brush.Color := clSilver;
    LabelState.Caption := 'stopped';
    LabelStatus.Caption := '';
  end;
end;

procedure TMainForm.Timer1Timer(Sender: TObject);
begin
  if (FEndAt > 0) and (GetTickCount64 >= FEndAt) then
  begin
    FEndAt := 0;
    ShowItems;
    ShowConvs;
    SaveFormSnapshot(Self, FSnapFile);
    Close;
    Exit;
  end;
  if FServer = nil then Exit;
  FServer.Tick;
  Inc(FTicks);
  if FTicks mod 5 <> 0 then Exit;
  ShowItems;
  ShowConvs;
  LabelStatus.Caption := Format('conversations %d   links %d   requests %d   advises %d   '
    + 'pokes %d   executes %d   refused %d', [FServer.ConvCount, FServer.LinkCount,
    FServer.Requests, FServer.Advises, FServer.Pokes, FServer.Executes, FServer.Refused]);
end;

function TMainForm.Formula(Row: Integer): string;
begin
  Result := '';
  if (FServer = nil) or (Row < 1) or (Row >= GridItems.RowCount) then Exit;
  Result := '=' + XlName(FServer.Service) + '|' + XlName(GridItems.Cells[0, Row]) + '!'
    + XlName(GridItems.Cells[1, Row]);
end;

procedure TMainForm.MenuCopyFormulaClick(Sender: TObject);
begin
  if Formula(GridItems.Row) <> '' then
    Clipboard.AsText := Formula(GridItems.Row);
end;

procedure TMainForm.MenuCopyValueClick(Sender: TObject);
var
  S: string;
begin
  if (FServer <> nil) and (GridItems.Row > 0) and
    FServer.Value(GridItems.Cells[0, GridItems.Row], GridItems.Cells[1, GridItems.Row], S) then
    Clipboard.AsText := S;
end;

procedure TMainForm.ButtonStartClick(Sender: TObject);
begin
  DoStart;
end;

procedure TMainForm.ButtonStopClick(Sender: TObject);
begin
  DoStop;
end;

procedure TMainForm.CheckDecCommaChange(Sender: TObject);
begin
  if FServer <> nil then FServer.DecimalComma := CheckDecComma.Checked;
end;

procedure TMainForm.CheckLogRequestsChange(Sender: TObject);
begin
  if FServer <> nil then FServer.LogRequests := CheckLogRequests.Checked;
end;

procedure TMainForm.MenuClearLogClick(Sender: TObject);
begin
  MemoLog.Clear;
end;

procedure TMainForm.MenuExitClick(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.MenuAboutClick(Sender: TObject);
begin
  MessageDlg('DDE Server',
    'Simulated HV PLC as a DDE server (DDEML), Free Pascal / Lazarus.' + LineEnding +
    LineEnding +
    'Topics: DB_HMI, Sim, Data, System' + LineEnding +
    'Excel: =NLHV|DB_HMI!VoltAct_kV' + LineEnding +
    'Execute: [Set(item,value)] [Enable] [Disable] [Ack] [Clear]', mtInformation, [mbOK], 0);
end;

function TMainForm.IniName: string;
begin
  Result := ChangeFileExt(Application.ExeName, '.ini');
end;

procedure TMainForm.LoadIni;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(IniName);
  try
    EditService.Text := Ini.ReadString('Server', 'Service', 'NLHV');
    if Ini.ReadInteger('Server', 'Points', 10) = 2 then
      ComboPoints.ItemIndex := 1
    else
      ComboPoints.ItemIndex := 0;
    CheckDecComma.Checked := Ini.ReadBool('Server', 'DecimalComma', False);
    CheckAutoStart.Checked := Ini.ReadBool('Server', 'AutoStart', False);
    CheckLogRequests.Checked := Ini.ReadBool('Server', 'LogRequests', False);
    Timer1.Interval := Ini.ReadInteger('Server', 'AdviseMs', 100);
  finally
    Ini.Free;
  end;
end;

procedure TMainForm.SaveIni;
var
  Ini: TIniFile;
begin
  try
    Ini := TIniFile.Create(IniName);
    try
      Ini.WriteString('Server', 'Service', EditService.Text);
      if ComboPoints.ItemIndex = 1 then
        Ini.WriteInteger('Server', 'Points', 2)
      else
        Ini.WriteInteger('Server', 'Points', 10);
      Ini.WriteBool('Server', 'DecimalComma', CheckDecComma.Checked);
      Ini.WriteBool('Server', 'AutoStart', CheckAutoStart.Checked);
      Ini.WriteBool('Server', 'LogRequests', CheckLogRequests.Checked);
      Ini.WriteInteger('Server', 'AdviseMs', Timer1.Interval);
    finally
      Ini.Free;
    end;
  except
    on E: Exception do
      Log('settings not saved: ' + E.Message);
  end;
end;

end.
