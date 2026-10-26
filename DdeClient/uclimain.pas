{ ============================================================================
  DDE Client - general DDE client (DDEML) built with normal Lazarus forms;
  works with DdeServer, Excel, and every program that is a DDE server
    - service and topic, a list of the running servers (Find servers)
    - item list from the server's TopicItemList, Topics from System
    - request, poke, execute; CF_TEXT or CF_UNICODETEXT; timeout
    - hot and warm links in a table (value, updates, time, state), made again
      when the server comes back
    - settings and links in DdeClient.ini next to the exe
    - command line (for tests): --service= --topic= --format=UNICODETEXT --find --connect --link=a,b
      --warm=a,b --poke=item=value --exec=cmd --request=item
      --dump=<file.txt> --after=<s> (links and log to the file, then exit)
      --snapshot=<file.png> (the window, with --after)
  ============================================================================ }
unit uclimain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Menus, ExtCtrls, StdCtrls, Grids,
  Clipbrd, IniFiles, uddeml, uddecli, ufrmsnap;

type

  { TMainForm }

  TMainForm = class(TForm)
    ButtonConnect: TButton;
    ButtonDisconnect: TButton;
    ButtonExecute: TButton;
    ButtonFind: TButton;
    ButtonHot: TButton;
    ButtonPoke: TButton;
    ButtonRequest: TButton;
    ButtonWarm: TButton;
    CheckAutoConnect: TCheckBox;
    CheckAutoReconnect: TCheckBox;
    ComboFormat: TComboBox;
    ComboItem: TComboBox;
    ComboService: TComboBox;
    ComboTopic: TComboBox;
    EditExecute: TEdit;
    EditTimeout: TEdit;
    EditValue: TEdit;
    GridLinks: TStringGrid;
    GroupTrans: TGroupBox;
    LabelExecute: TLabel;
    LabelFormat: TLabel;
    LabelItem: TLabel;
    LabelResult: TLabel;
    LabelService: TLabel;
    LabelState: TLabel;
    LabelStatus: TLabel;
    LabelTimeout: TLabel;
    LabelTopic: TLabel;
    LabelValue: TLabel;
    MainMenu1: TMainMenu;
    MemoLog: TMemo;
    MemoResult: TMemo;
    MenuAbout: TMenuItem;
    MenuClearLog: TMenuItem;
    MenuConnect: TMenuItem;
    MenuCopyValue: TMenuItem;
    MenuDisconnect: TMenuItem;
    MenuExit: TMenuItem;
    MenuFile: TMenuItem;
    MenuFind: TMenuItem;
    MenuHelp: TMenuItem;
    MenuRemoveAll: TMenuItem;
    MenuRemoveLink: TMenuItem;
    MenuRequestNow: TMenuItem;
    MenuSep1: TMenuItem;
    MenuSep2: TMenuItem;
    MenuToPoke: TMenuItem;
    PanelConn: TPanel;
    PopupLinks: TPopupMenu;
    ShapeState: TShape;
    SplitterLog: TSplitter;
    Timer1: TTimer;
    procedure ButtonConnectClick(Sender: TObject);
    procedure ButtonDisconnectClick(Sender: TObject);
    procedure ButtonExecuteClick(Sender: TObject);
    procedure ButtonFindClick(Sender: TObject);
    procedure ButtonHotClick(Sender: TObject);
    procedure ButtonPokeClick(Sender: TObject);
    procedure ButtonRequestClick(Sender: TObject);
    procedure ButtonWarmClick(Sender: TObject);
    procedure CheckAutoReconnectChange(Sender: TObject);
    procedure ComboServiceChange(Sender: TObject);
    procedure EditExecuteKeyPress(Sender: TObject; var Key: Char);
    procedure EditTimeoutChange(Sender: TObject);
    procedure EditValueKeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridLinksDblClick(Sender: TObject);
    procedure GridLinksResize(Sender: TObject);
    procedure MenuAboutClick(Sender: TObject);
    procedure MenuClearLogClick(Sender: TObject);
    procedure MenuCopyValueClick(Sender: TObject);
    procedure MenuExitClick(Sender: TObject);
    procedure MenuRemoveAllClick(Sender: TObject);
    procedure MenuRemoveLinkClick(Sender: TObject);
    procedure MenuRequestNowClick(Sender: TObject);
    procedure MenuToPokeClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    FClient: TDdeClient;
    FServers: TStringList;           // service|topic from Find servers
    FDirty: Boolean;
    FTicks: Integer;
    FDumpFile: string;
    FSnapFile: string;
    FEndAt: QWord;
    procedure Log(const S: string);
    procedure ClientData(Sender: TObject);
    function Service: string;
    function Topic: string;
    function Fmt: Cardinal;
    procedure DoConnect;
    procedure DoLink(Warm: Boolean);
    procedure DoRequest;
    procedure DoPoke;
    procedure DoExecute;
    procedure ShowLinks;
    procedure UpdateState;
    function SelectedLink: TCliLink;
    procedure AddHistory(Combo: TComboBox; const S: string);
    procedure RunCommandLine;
    procedure Dump;
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
  StateText: array[TLinkState] of string = ('active', 'no server', 'refused');

{ a value in one line: rows as ' | ', tabs as blanks }
function Shown(const S: string): string;
begin
  Result := StringReplace(StringReplace(S, #13#10, '  |  ', [rfReplaceAll]), #9, '  ',
    [rfReplaceAll]);
  if Length(Result) > 300 then
    Result := Copy(Result, 1, 300) + '...';
end;

{ TMainForm }

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FServers := TStringList.Create;
  FClient := TDdeClient.Create;
  FClient.OnLog := @Log;
  FClient.OnData := @ClientData;
  GridLinks.Cells[0, 0] := 'Link';
  GridLinks.Cells[1, 0] := 'Mode';
  GridLinks.Cells[2, 0] := 'Format';
  GridLinks.Cells[3, 0] := 'Value';
  GridLinks.Cells[4, 0] := 'Updates';
  GridLinks.Cells[5, 0] := 'Time';
  GridLinks.Cells[6, 0] := 'State';
  LoadIni;
  FClient.Timeout := StrToIntDef(EditTimeout.Text, 3000);
  FClient.AutoReconnect := CheckAutoReconnect.Checked;
  FClient.Paused := not CheckAutoConnect.Checked and not Application.HasOption('connect');
  ShowLinks;
  UpdateState;
  if Application.HasOption('service') or Application.HasOption('topic') or
    Application.HasOption('format') or Application.HasOption('find') or
    Application.HasOption('connect') or Application.HasOption('dump') or
    Application.HasOption('snapshot') then
    RunCommandLine
  else if CheckAutoConnect.Checked then
    DoConnect;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if FDumpFile = '' then SaveIni;      // a test run keeps the user's settings
  FClient.OnLog := nil;
  FClient.OnData := nil;
  FClient.Free;
  FServers.Free;
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

procedure TMainForm.ClientData(Sender: TObject);
begin
  FDirty := True;
end;

function TMainForm.Service: string;
begin
  Result := Trim(ComboService.Text);
end;

function TMainForm.Topic: string;
begin
  Result := Trim(ComboTopic.Text);
end;

function TMainForm.Fmt: Cardinal;
begin
  if ComboFormat.ItemIndex = 1 then
    Result := CF_UNICODETEXT
  else
    Result := CF_TEXT;
end;

procedure TMainForm.AddHistory(Combo: TComboBox; const S: string);
var
  I: Integer;
begin
  if Trim(S) = '' then Exit;
  I := Combo.Items.IndexOf(S);
  if I = 0 then Exit;
  if I > 0 then Combo.Items.Delete(I);
  Combo.Items.Insert(0, S);
  while Combo.Items.Count > 20 do
    Combo.Items.Delete(Combo.Items.Count - 1);
  Combo.Text := S;
end;

procedure TMainForm.DoConnect;
var
  Err, V, OldItem, OldTopic: string;
  L: TStringList;
begin
  FClient.Paused := False;
  if FClient.Connect(Service, Topic, Err) = 0 then
  begin
    Log(Format('connect %s|%s failed: %s', [Service, Topic, Err]));
    UpdateState;
    Exit;
  end;
  AddHistory(ComboService, Service);
  { the item names, if the server has a TopicItemList (or SysItems in System) }
  OldItem := ComboItem.Text;
  L := TStringList.Create;
  try
    L.Delimiter := #9;
    L.StrictDelimiter := True;
    if FClient.Request(Service, Topic, SZDDE_ITEM_ITEMLIST, CF_TEXT, V, Err) or
      (SameText(Topic, SZDDESYS_TOPIC) and
       FClient.Request(Service, Topic, SZDDESYS_ITEM_SYSITEMS, CF_TEXT, V, Err)) then
    begin
      L.DelimitedText := V;
      ComboItem.Items.Assign(L);
      ComboItem.Text := OldItem;
      Log(Format('%d items in %s|%s', [L.Count, Service, Topic]));
    end;
    { the topics from the System topic }
    if FClient.Connected(Service, SZDDESYS_TOPIC) or SameText(Topic, SZDDESYS_TOPIC) or
      (FClient.Connect(Service, SZDDESYS_TOPIC, Err) <> 0) then
      if FClient.Request(Service, SZDDESYS_TOPIC, SZDDESYS_ITEM_TOPICS, CF_TEXT, V, Err) then
      begin
        OldTopic := ComboTopic.Text;
        L.DelimitedText := V;
        ComboTopic.Items.Assign(L);
        ComboTopic.Text := OldTopic;
      end;
    { the System conversation only for this question, unless it is the topic }
    if not SameText(Topic, SZDDESYS_TOPIC) then
      FClient.Disconnect(Service, SZDDESYS_TOPIC);
  finally
    L.Free;
  end;
  UpdateState;
end;

procedure TMainForm.DoLink(Warm: Boolean);
var
  Err: string;
begin
  if Trim(ComboItem.Text) = '' then
  begin
    Log('no item');
    Exit;
  end;
  FClient.Paused := False;
  if FClient.AddLink(Service, Topic, Trim(ComboItem.Text), Warm, Fmt, Err) = nil then
    Log('link not added: ' + Err);
  ShowLinks;
  UpdateState;
end;

procedure TMainForm.DoRequest;
var
  V, Err: string;
begin
  if Trim(ComboItem.Text) = '' then Exit;
  if FClient.Request(Service, Topic, Trim(ComboItem.Text), Fmt, V, Err) then
  begin
    MemoResult.Text := StringReplace(V, #9, '    ', [rfReplaceAll]);
    LabelResult.Caption := Format('%s  (%d characters)', [Trim(ComboItem.Text), Length(V)]);
    Log(Format('request %s = %s', [LinkName(Service, Topic, Trim(ComboItem.Text)), Shown(V)]));
  end
  else
  begin
    MemoResult.Text := '';
    LabelResult.Caption := 'request failed';
    Log(Format('request %s failed: %s', [LinkName(Service, Topic, Trim(ComboItem.Text)), Err]));
  end;
  UpdateState;
end;

procedure TMainForm.DoPoke;
var
  Err, V, Ret: string;
begin
  if Trim(ComboItem.Text) = '' then Exit;
  { "\t" and "\n" for tab separated values and rows (curves) }
  V := StringReplace(StringReplace(EditValue.Text, '\t', #9, [rfReplaceAll]), '\n', #13#10,
    [rfReplaceAll]);
  if FClient.Poke(Service, Topic, Trim(ComboItem.Text), V, Fmt, Err) then
    Log(Format('poke %s = %s: accepted', [LinkName(Service, Topic, Trim(ComboItem.Text)),
      Shown(V)]))
  else
  begin
    { DdeServer and Excel tell why in System!ReturnMessage }
    Ret := '';
    if FClient.Connect(Service, SZDDESYS_TOPIC, Ret) <> 0 then
    begin
      if not FClient.Request(Service, SZDDESYS_TOPIC, SZDDESYS_ITEM_RTNMSG, CF_TEXT, Ret, Ret) then
        Ret := '';
      if not SameText(Topic, SZDDESYS_TOPIC) then
        FClient.Disconnect(Service, SZDDESYS_TOPIC);
    end;
    if Ret <> '' then Err := Err + ' - ' + Ret;
    Log(Format('poke %s = %s: %s', [LinkName(Service, Topic, Trim(ComboItem.Text)), Shown(V),
      Err]));
  end;
  UpdateState;
end;

procedure TMainForm.DoExecute;
var
  Err, Ret: string;
begin
  if Trim(EditExecute.Text) = '' then Exit;
  if FClient.Execute(Service, Topic, EditExecute.Text, Err) then
    Log(Format('execute on %s|%s: %s accepted', [Service, Topic, EditExecute.Text]))
  else
  begin
    Ret := '';
    if FClient.Connect(Service, SZDDESYS_TOPIC, Ret) <> 0 then
    begin
      if not FClient.Request(Service, SZDDESYS_TOPIC, SZDDESYS_ITEM_RTNMSG, CF_TEXT, Ret, Ret) then
        Ret := '';
      if not SameText(Topic, SZDDESYS_TOPIC) then
        FClient.Disconnect(Service, SZDDESYS_TOPIC);
    end;
    if Ret <> '' then Err := Err + ' - ' + Ret;
    Log(Format('execute on %s|%s: %s: %s', [Service, Topic, EditExecute.Text, Err]));
  end;
  UpdateState;
end;

procedure TMainForm.ShowLinks;
var
  I: Integer;
  L: TCliLink;
begin
  FDirty := False;
  GridLinks.BeginUpdate;
  try
    GridLinks.RowCount := FClient.LinkCount + 1;
    for I := 0 to FClient.LinkCount - 1 do
    begin
      L := FClient.Link(I);
      GridLinks.Cells[0, I + 1] := LinkName(L.Service, L.Topic, L.Item);
      GridLinks.Cells[1, I + 1] := BoolToStr(L.Warm, 'warm', 'hot');
      GridLinks.Cells[2, I + 1] := FormatName(L.Fmt);
      GridLinks.Cells[3, I + 1] := Shown(L.Value);
      GridLinks.Cells[4, I + 1] := IntToStr(L.Updates);
      if L.Time > 0 then
        GridLinks.Cells[5, I + 1] := FormatDateTime('hh:nn:ss.zzz', L.Time)
      else
        GridLinks.Cells[5, I + 1] := '';
      GridLinks.Cells[6, I + 1] := StateText[L.State];
    end;
  finally
    GridLinks.EndUpdate;
  end;
end;

procedure TMainForm.UpdateState;
var
  N, I: Integer;
begin
  N := FClient.ConvCount;
  ButtonDisconnect.Enabled := N > 0;
  MenuDisconnect.Enabled := N > 0;
  if FClient.Connected(Service, Topic) then
  begin
    ShapeState.Brush.Color := clLime;
    LabelState.Caption := Format('connected to %s|%s', [Service, Topic]);
  end
  else if N > 0 then
  begin
    ShapeState.Brush.Color := clYellow;
    LabelState.Caption := Format('%d conversation(s), not %s|%s', [N, Service, Topic]);
  end
  else
  begin
    ShapeState.Brush.Color := clSilver;
    LabelState.Caption := 'not connected';
    for I := 0 to FClient.LinkCount - 1 do
      if FClient.Link(I).State <> lsActive then
      begin
        ShapeState.Brush.Color := clRed;
        if FClient.Paused then
          LabelState.Caption := 'disconnected'
        else
          LabelState.Caption := 'waiting for the server';
        Break;
      end;
  end;
  LabelStatus.Caption := Format('conversations %d   links %d   transactions %d   advises %d   errors %d',
    [N, FClient.LinkCount, FClient.Transactions, FClient.Advises, FClient.Errors]);
end;

procedure TMainForm.Timer1Timer(Sender: TObject);
begin
  if FClient.Busy then Exit;
  FClient.Tick;
  if FDirty then ShowLinks;
  Inc(FTicks);
  if FTicks mod 5 = 0 then UpdateState;
  if (FEndAt > 0) and (GetTickCount64 >= FEndAt) then
  begin
    FEndAt := 0;
    ShowLinks;
    UpdateState;
    Dump;
    Close;
  end;
end;

function TMainForm.SelectedLink: TCliLink;
begin
  if (GridLinks.Row >= 1) and (GridLinks.Row <= FClient.LinkCount) then
    Result := FClient.Link(GridLinks.Row - 1)
  else
    Result := nil;
end;

procedure TMainForm.ButtonConnectClick(Sender: TObject);
begin
  DoConnect;
end;

procedure TMainForm.ButtonDisconnectClick(Sender: TObject);
begin
  FClient.Paused := True;
  FClient.DisconnectAll;
  ShowLinks;
  UpdateState;
end;

procedure TMainForm.ButtonFindClick(Sender: TObject);
var
  I: Integer;
  Svc: string;
begin
  Screen.Cursor := crHourGlass;
  try
    FClient.FindServers(FServers);
  finally
    Screen.Cursor := crDefault;
  end;
  Log(Format('%d service|topic pairs found', [FServers.Count]));
  for I := 0 to FServers.Count - 1 do
  begin
    Log('  ' + FServers[I]);
    Svc := Copy(FServers[I], 1, Pos('|', FServers[I]) - 1);
    if ComboService.Items.IndexOf(Svc) < 0 then
      ComboService.Items.Add(Svc);
  end;
  ComboServiceChange(nil);
end;

{ the topics of the chosen service, from Find servers }
procedure TMainForm.ComboServiceChange(Sender: TObject);
var
  I: Integer;
  OldTopic: string;
begin
  if FServers.Count = 0 then Exit;
  OldTopic := ComboTopic.Text;
  ComboTopic.Items.Clear;
  for I := 0 to FServers.Count - 1 do
    if SameText(Copy(FServers[I], 1, Pos('|', FServers[I]) - 1), Service) then
      ComboTopic.Items.Add(Copy(FServers[I], Pos('|', FServers[I]) + 1, MaxInt));
  ComboTopic.Text := OldTopic;
  UpdateState;
end;

procedure TMainForm.ButtonRequestClick(Sender: TObject);
begin
  DoRequest;
end;

procedure TMainForm.ButtonHotClick(Sender: TObject);
begin
  DoLink(False);
end;

procedure TMainForm.ButtonWarmClick(Sender: TObject);
begin
  DoLink(True);
end;

procedure TMainForm.ButtonPokeClick(Sender: TObject);
begin
  DoPoke;
end;

procedure TMainForm.ButtonExecuteClick(Sender: TObject);
begin
  DoExecute;
end;

procedure TMainForm.EditValueKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    DoPoke;
  end;
end;

procedure TMainForm.EditExecuteKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
  begin
    Key := #0;
    DoExecute;
  end;
end;

procedure TMainForm.CheckAutoReconnectChange(Sender: TObject);
begin
  if FClient <> nil then FClient.AutoReconnect := CheckAutoReconnect.Checked;
end;

procedure TMainForm.EditTimeoutChange(Sender: TObject);
begin
  if FClient <> nil then FClient.Timeout := StrToIntDef(EditTimeout.Text, 3000);
end;

{ double-click a link: its service, topic, item and value go to the transaction }
procedure TMainForm.GridLinksDblClick(Sender: TObject);
begin
  MenuToPokeClick(nil);
end;

{ the Value column takes the width the others leave }
procedure TMainForm.GridLinksResize(Sender: TObject);
var
  I, W: Integer;
begin
  W := GridLinks.ClientWidth - 2;
  for I := 0 to GridLinks.ColCount - 1 do
    if I <> 3 then Dec(W, GridLinks.ColWidths[I]);
  if W > 120 then GridLinks.ColWidths[3] := W;
end;

procedure TMainForm.MenuToPokeClick(Sender: TObject);
var
  L: TCliLink;
begin
  L := SelectedLink;
  if L = nil then Exit;
  ComboService.Text := L.Service;
  ComboTopic.Text := L.Topic;
  ComboItem.Text := L.Item;
  EditValue.Text := StringReplace(StringReplace(L.Value, #9, '\t', [rfReplaceAll]), #13#10,
    '\n', [rfReplaceAll]);
  if L.Fmt = CF_UNICODETEXT then ComboFormat.ItemIndex := 1 else ComboFormat.ItemIndex := 0;
  UpdateState;
  EditValue.SetFocus;
end;

procedure TMainForm.MenuRequestNowClick(Sender: TObject);
var
  L: TCliLink;
  V, Err: string;
begin
  L := SelectedLink;
  if L = nil then Exit;
  if FClient.Request(L.Service, L.Topic, L.Item, L.Fmt, V, Err) then
  begin
    L.Value := V;
    L.Time := Now;
    ShowLinks;
  end
  else
    Log(Format('request %s failed: %s', [LinkName(L.Service, L.Topic, L.Item), Err]));
end;

procedure TMainForm.MenuCopyValueClick(Sender: TObject);
begin
  if SelectedLink <> nil then
    Clipboard.AsText := SelectedLink.Value;
end;

procedure TMainForm.MenuRemoveLinkClick(Sender: TObject);
begin
  if SelectedLink = nil then Exit;
  FClient.RemoveLink(SelectedLink);
  ShowLinks;
  UpdateState;
end;

procedure TMainForm.MenuRemoveAllClick(Sender: TObject);
begin
  while FClient.LinkCount > 0 do
    FClient.RemoveLink(FClient.Link(0));
  ShowLinks;
  UpdateState;
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
  MessageDlg('DDE Client',
    'General DDE client (DDEML), Free Pascal / Lazarus.' + LineEnding + LineEnding +
    'Request, poke, execute, hot and warm links,' + LineEnding +
    'CF_TEXT and CF_UNICODETEXT.' + LineEnding +
    'Poke values: \t = tab, \n = new row.', mtInformation, [mbOK], 0);
end;

{ ---------------------------------------------------------------- command line }

procedure TMainForm.RunCommandLine;
var
  L: TStringList;
  I, P: Integer;
  S: string;
begin
  if Application.HasOption('service') then ComboService.Text := Application.GetOptionValue('service');
  if Application.HasOption('topic') then ComboTopic.Text := Application.GetOptionValue('topic');
  if Application.HasOption('format') then
    ComboFormat.ItemIndex := Ord(SameText(Application.GetOptionValue('format'), 'UNICODETEXT'));
  FDumpFile := Application.GetOptionValue('dump');
  FSnapFile := Application.GetOptionValue('snapshot');
  if Application.HasOption('find') then ButtonFindClick(nil);
  if Application.HasOption('connect') then DoConnect;
  L := TStringList.Create;
  try
    L.Delimiter := ',';
    L.StrictDelimiter := True;
    L.DelimitedText := Application.GetOptionValue('link');
    for I := 0 to L.Count - 1 do
      if Trim(L[I]) <> '' then
      begin
        ComboItem.Text := Trim(L[I]);
        DoLink(False);
      end;
    L.DelimitedText := Application.GetOptionValue('warm');
    for I := 0 to L.Count - 1 do
      if Trim(L[I]) <> '' then
      begin
        ComboItem.Text := Trim(L[I]);
        DoLink(True);
      end;
  finally
    L.Free;
  end;
  S := Application.GetOptionValue('poke');
  P := Pos('=', S);
  if P > 1 then
  begin
    ComboItem.Text := Copy(S, 1, P - 1);
    EditValue.Text := Copy(S, P + 1, MaxInt);
    DoPoke;
  end;
  if Application.HasOption('exec') then
  begin
    EditExecute.Text := Application.GetOptionValue('exec');
    DoExecute;
  end;
  if Application.HasOption('request') then
  begin
    ComboItem.Text := Application.GetOptionValue('request');
    DoRequest;
  end;
  if (FDumpFile <> '') or (FSnapFile <> '') then
    FEndAt := GetTickCount64 + Round(StrToFloatDef(Application.GetOptionValue('after'), 3) * 1000);
end;

procedure TMainForm.Dump;
var
  F: TStringList;
  R, C: Integer;
  S: string;
begin
  if FSnapFile <> '' then
    SaveFormSnapshot(Self, FSnapFile);
  if FDumpFile = '' then Exit;
  F := TStringList.Create;
  try
    F.Add('state: ' + LabelState.Caption);
    F.Add('status: ' + LabelStatus.Caption);
    F.Add('result: ' + LabelResult.Caption + ' = ' + Shown(MemoResult.Text));
    F.Add('items: ' + IntToStr(ComboItem.Items.Count) + ', topics: ' + ComboTopic.Items.CommaText);
    for R := 0 to GridLinks.RowCount - 1 do
    begin
      S := '';
      for C := 0 to GridLinks.ColCount - 1 do
        S := S + GridLinks.Cells[C, R] + ' ; ';
      F.Add(S);
    end;
    F.Add('--- log');
    F.AddStrings(MemoLog.Lines);
    F.SaveToFile(FDumpFile);
  finally
    F.Free;
  end;
end;

{ ---------------------------------------------------------------- settings }

function TMainForm.IniName: string;
begin
  Result := ChangeFileExt(Application.ExeName, '.ini');
end;

procedure TMainForm.LoadIni;
var
  Ini: TIniFile;
  I, N: Integer;
  P: TStringList;
  Err: string;
begin
  Ini := TIniFile.Create(IniName);
  P := TStringList.Create;
  try
    ComboService.Text := Ini.ReadString('Connection', 'Service', 'NLHV');
    ComboTopic.Text := Ini.ReadString('Connection', 'Topic', 'DB_HMI');
    ComboItem.Text := Ini.ReadString('Connection', 'Item', 'VoltAct_kV');
    if Ini.ReadString('Connection', 'Format', 'TEXT') = 'UNICODETEXT' then
      ComboFormat.ItemIndex := 1
    else
      ComboFormat.ItemIndex := 0;
    EditTimeout.Text := IntToStr(Ini.ReadInteger('Connection', 'TimeoutMs', 3000));
    CheckAutoReconnect.Checked := Ini.ReadBool('Connection', 'AutoReconnect', True);
    CheckAutoConnect.Checked := Ini.ReadBool('Connection', 'AutoConnect', False);
    EditExecute.Text := Ini.ReadString('Connection', 'Execute', '[Enable]');
    if Application.HasOption('dump') then Exit;   // a test run starts without saved links
    { saved links come back without a conversation; Connect (or auto connect) makes them }
    FClient.Paused := True;
    N := Ini.ReadInteger('Links', 'Count', 0);
    for I := 0 to N - 1 do
    begin
      P.Delimiter := '|';
      P.StrictDelimiter := True;
      P.DelimitedText := Ini.ReadString('Links', 'Link' + IntToStr(I), '');
      if P.Count <> 5 then Continue;
      if P[1] = 'UNICODETEXT' then
        FClient.AddLink(P[2], P[3], P[4], P[0] = 'warm', CF_UNICODETEXT, Err)
      else
        FClient.AddLink(P[2], P[3], P[4], P[0] = 'warm', CF_TEXT, Err);
    end;
  finally
    P.Free;
    Ini.Free;
  end;
end;

procedure TMainForm.SaveIni;
var
  Ini: TIniFile;
  I: Integer;
  L: TCliLink;
begin
  try
    Ini := TIniFile.Create(IniName);
    try
      Ini.WriteString('Connection', 'Service', Service);
      Ini.WriteString('Connection', 'Topic', Topic);
      Ini.WriteString('Connection', 'Item', ComboItem.Text);
      Ini.WriteString('Connection', 'Format', FormatName(Fmt));
      Ini.WriteInteger('Connection', 'TimeoutMs', StrToIntDef(EditTimeout.Text, 3000));
      Ini.WriteBool('Connection', 'AutoReconnect', CheckAutoReconnect.Checked);
      Ini.WriteBool('Connection', 'AutoConnect', CheckAutoConnect.Checked);
      Ini.WriteString('Connection', 'Execute', EditExecute.Text);
      Ini.EraseSection('Links');
      Ini.WriteInteger('Links', 'Count', FClient.LinkCount);
      for I := 0 to FClient.LinkCount - 1 do
      begin
        L := FClient.Link(I);
        Ini.WriteString('Links', 'Link' + IntToStr(I), BoolToStr(L.Warm, 'warm', 'hot') + '|'
          + FormatName(L.Fmt) + '|' + L.Service + '|' + L.Topic + '|' + L.Item);
      end;
    finally
      Ini.Free;
    end;
  except
    on E: Exception do
      Log('settings not saved: ' + E.Message);
  end;
end;

end.
