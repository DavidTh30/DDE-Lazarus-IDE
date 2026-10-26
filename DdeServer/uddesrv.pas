{ ============================================================================
  uddesrv - DDE server (DDEML) for the simulated HV PLC (uplcsim)
    service  NLHV (configurable)
    topics   DB_HMI   the tags of DB10 (VoltSP_kV, VoltAct_kV, Enable, Curve_VSP, ...)
             Sim      Counter, Sine, Toggle, Clock
             Data     free items: any client pokes any item, others link to it
             System   Topics, SysItems, Formats, Status, Help, ReturnMessage
    every topic also has TopicItemList and Formats
  Formats CF_TEXT (ANSI) and CF_UNICODETEXT. Request, hot / warm links
  (advise), poke and execute ([Set(item,value)][Enable][Disable][Ack][Clear]).
  Curves are 2 rows (Eng, Volt) of tab separated values, as Excel expects.
  DDEML calls back in the thread that called Start, here the main thread
  (the LCL message loop), so nothing here needs a lock except the PLC.
  ============================================================================ }
unit uddesrv;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, Math, uddeml, uaencode, uaserver, uplcsim;

type
  TSrvItemKind = (ikTag, ikCurve, ikCurvePart, ikData);

  TSrvItem = class
    Topic: string;
    Name: string;
    Kind: TSrvItemKind;
    TagKind: TTagKind;
    Tag: Integer;
    VoltTag: Integer;          // ikCurve: the .Volt tag (Tag is .Eng)
    Writable: Boolean;
    Data: string;              // ikData
  end;

  TSrvConv = class
    Conv: HCONV;
    Topic: string;
    Process: string;
    Since: TDateTime;
    Transactions: Integer;
  end;

  { an advise loop: one per conversation, item and format }
  TSrvLink = class
    Conv: HCONV;
    Topic: string;
    Item: string;
    Fmt: UINT;
  end;

  { the items with links: last text sent, to advise only on a change }
  TAdvised = class
    Topic: string;
    Item: string;
    Count: Integer;
    Last: string;
    HTopic, HItem: HSZ;
  end;

  TSrvLogEvent = procedure(const S: string) of object;

  { TDdeSimServer }

  TDdeSimServer = class
  private
    FInst: DWORD;
    FService: string;
    FHService: HSZ;
    FHTopics: array of HSZ;
    FSim: TPlcSim;
    FItems: TStringList;       // 'topic!item' (lower case) -> TSrvItem
    FConvs: TList;
    FLinks: TList;
    FAdvised: TStringList;     // 'topic!item' (lower case) -> TAdvised
    FReturnMsg: string;
    FDecComma: Boolean;
    FLogRequests: Boolean;
    FOnLog: TSrvLogEvent;
    FFS: TFormatSettings;
    procedure Log(const S: string);
    procedure AddItem(const ATopic, AName: string; AKind: TSrvItemKind; ATagKind: TTagKind;
      ATag, AVolt: Integer; AWritable: Boolean);
    function FindConv(C: HCONV): TSrvConv;
    function DataCount: Integer;
    function TopicName(const S: string): string;
    function NumText(X: Double; Digits: Integer): string;
    function RowText(Tag: Integer): string;
    function ParseRows(const S: string; out Rows: TStringList): Boolean;
    function ParseNum(const S: string; out X: Double): Boolean;
    function WriteItem(It: TSrvItem; const S: string; out Msg: string): Boolean;
    function Execute(const ATopic, Cmds: string; out Msg: string): Boolean;
    procedure AddLink(C: HCONV; const ATopic, AItem: string; Fmt: UINT);
    procedure RemoveLink(C: HCONV; const ATopic, AItem: string; Fmt: UINT);
    procedure RemoveConvLinks(C: HCONV);
    function Callback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ; hData: HDDEDATA;
      dwData1, dwData2: PtrUInt): HDDEDATA;
  public
    Requests, Pokes, Executes, Advises, Refused: Int64;
    constructor Create(ASim: TPlcSim);
    destructor Destroy; override;
    function Start(const AService: string; out Err: string): Boolean;
    procedure Stop;
    function Running: Boolean;
    { advise the links whose value changed; call every 50..250 ms }
    procedure Tick;
    function ItemCount: Integer;
    function Item(I: Integer): TSrvItem;
    function ConvCount: Integer;
    function Conv(I: Integer): TSrvConv;
    function LinkCount: Integer;
    function Link(I: Integer): TSrvLink;
    function LinksOf(const ATopic, AItem: string): Integer;
    { text of topic!item, False when unknown }
    function Value(const ATopic, AItem: string; out S: string): Boolean;
    function TypeText(It: TSrvItem): string;
    property Service: string read FService;
    property DecimalComma: Boolean read FDecComma write FDecComma;
    property LogRequests: Boolean read FLogRequests write FLogRequests;
    property OnLog: TSrvLogEvent read FOnLog write FOnLog;
  end;

const
  TOPIC_HMI = 'DB_HMI';
  TOPIC_SIM = 'Sim';
  TOPIC_DATA = 'Data';
  SrvTopics: array[0..3] of string = (SZDDESYS_TOPIC, TOPIC_HMI, TOPIC_SIM, TOPIC_DATA);
  SysItems: array[0..6] of string = (SZDDESYS_ITEM_TOPICS, SZDDESYS_ITEM_SYSITEMS,
    SZDDESYS_ITEM_FORMATS, SZDDESYS_ITEM_STATUS, SZDDESYS_ITEM_HELP, SZDDESYS_ITEM_RTNMSG,
    SZDDE_ITEM_ITEMLIST);
  MAX_DATA_ITEMS = 1000;

implementation

var
  GServer: TDdeSimServer = nil;

function ServerCallback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ; hData: HDDEDATA;
  dwData1, dwData2: PtrUInt): HDDEDATA; stdcall;
begin
  Result := 0;
  if GServer <> nil then
  try
    Result := GServer.Callback(wType, wFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2);
  except
    on E: Exception do
      GServer.Log('error in ' + XtypName(wType) + ': ' + E.Message);
  end;
end;

function Key(const ATopic, AItem: string): string;
begin
  Result := LowerCase(ATopic + '!' + AItem);
end;

{ a value for the log: rows as ' | ', tabs as blanks, cut at 200 characters }
function OneLine(const S: string): string;
begin
  Result := StringReplace(StringReplace(S, #13#10, ' | ', [rfReplaceAll]), #9, ' ', [rfReplaceAll]);
  if Length(Result) > 200 then
    Result := Copy(Result, 1, 200) + '...';
end;

const
  HelpText = 'DdeServer: simulated HV PLC. Topics DB_HMI, Sim, Data. '
    + 'Excel: =NLHV|DB_HMI!VoltAct_kV. Execute: [Set(item,value)] [Enable] [Disable] [Ack] [Clear].';

{ TDdeSimServer }

constructor TDdeSimServer.Create(ASim: TPlcSim);
var
  I, J: Integer;
  Name, Topic: string;
begin
  FSim := ASim;
  FItems := TStringList.Create;
  FItems.Sorted := True;
  FItems.OwnsObjects := True;
  FConvs := TList.Create;
  FLinks := TList.Create;
  FAdvised := TStringList.Create;
  FAdvised.Sorted := True;
  FAdvised.OwnsObjects := True;
  FService := 'NLHV';
  FFS := DefaultFormatSettings;
  FFS.ThousandSeparator := #0;
  for I := 0 to High(FSim.Tags) do
  begin
    Name := FSim.Tags[I].Name;
    Topic := Copy(Name, 1, Pos('.', Name) - 1);
    Name := Copy(Name, Pos('.', Name) + 1, MaxInt);
    if FSim.Tags[I].Kind = tkRealArray then
    begin
      { Curve_VSP.Eng / .Volt: the parts and the curve as 2 rows }
      AddItem(Topic, Name, ikCurvePart, tkRealArray, FSim.Tags[I].Id, -1, True);
      if not Name.EndsWith('.Eng') then Continue;
      Name := Copy(Name, 1, Length(Name) - 4);
      AddItem(Topic, Name, ikCurve, tkRealArray, FSim.Tags[I].Id, FSim.Tags[I].Id + 1, True);
      for J := 0 to High(FSim.Tags) do
        if FSim.Tags[J].Name = Topic + '.' + Name + '.Volt' then
          TSrvItem(FItems.Objects[FItems.Count - 1]).VoltTag := FSim.Tags[J].Id;
      Continue;
    end;
    AddItem(Topic, Name, ikTag, FSim.Tags[I].Kind, FSim.Tags[I].Id, -1, FSim.Tags[I].Writable);
  end;
end;

destructor TDdeSimServer.Destroy;
begin
  Stop;
  FItems.Free;
  FConvs.Free;
  FLinks.Free;
  FAdvised.Free;
  inherited Destroy;
end;

procedure TDdeSimServer.Log(const S: string);
begin
  if Assigned(FOnLog) then FOnLog(S);
end;

procedure TDdeSimServer.AddItem(const ATopic, AName: string; AKind: TSrvItemKind;
  ATagKind: TTagKind; ATag, AVolt: Integer; AWritable: Boolean);
var
  It: TSrvItem;
begin
  It := TSrvItem.Create;
  It.Topic := ATopic;
  It.Name := AName;
  It.Kind := AKind;
  It.TagKind := ATagKind;
  It.Tag := ATag;
  It.VoltTag := AVolt;
  It.Writable := AWritable;
  FItems.AddObject(Key(ATopic, AName), It);
end;

function TDdeSimServer.Start(const AService: string; out Err: string): Boolean;
var
  R: UINT;
  I: Integer;
begin
  Err := '';
  if Running then Exit(True);
  if Trim(AService) = '' then
  begin
    Err := 'no service name';
    Exit(False);
  end;
  if (GServer <> nil) and (GServer <> Self) then
  begin
    Err := 'another server runs in this program';
    Exit(False);
  end;
  GServer := Self;
  FService := Trim(AService);
  FInst := 0;
  R := DdeInitializeW(FInst, @ServerCallback, APPCLASS_STANDARD or CBF_SKIP_REGISTRATIONS
    or CBF_SKIP_UNREGISTRATIONS, 0);
  if R <> DMLERR_NO_ERROR then
  begin
    Err := 'DdeInitialize: ' + DdeErrorText(R);
    FInst := 0;
    GServer := nil;
    Exit(False);
  end;
  FHService := HszOf(FInst, FService);
  SetLength(FHTopics, Length(SrvTopics));
  for I := 0 to High(SrvTopics) do
    FHTopics[I] := HszOf(FInst, SrvTopics[I]);
  if DdeNameService(FInst, FHService, 0, DNS_REGISTER) = 0 then
  begin
    Err := 'DdeNameService: ' + DdeErrorText(DdeGetLastError(FInst));
    Stop;
    Exit(False);
  end;
  Requests := 0; Pokes := 0; Executes := 0; Advises := 0; Refused := 0;
  FReturnMsg := '';
  Log(Format('service "%s" registered, topics %s, %s, %s, %s',
    [FService, SrvTopics[0], SrvTopics[1], SrvTopics[2], SrvTopics[3]]));
  Result := True;
end;

procedure TDdeSimServer.Stop;
var
  I: Integer;
begin
  if FInst = 0 then Exit;
  DdeNameService(FInst, FHService, 0, DNS_UNREGISTER);
  for I := 0 to FAdvised.Count - 1 do
    with TAdvised(FAdvised.Objects[I]) do
    begin
      DdeFreeStringHandle(FInst, HTopic);
      DdeFreeStringHandle(FInst, HItem);
    end;
  FAdvised.Clear;
  for I := 0 to High(FHTopics) do
    DdeFreeStringHandle(FInst, FHTopics[I]);
  SetLength(FHTopics, 0);
  DdeFreeStringHandle(FInst, FHService);
  DdeUninitialize(FInst);     // ends every conversation
  FInst := 0;
  for I := 0 to FConvs.Count - 1 do
    TSrvConv(FConvs[I]).Free;
  FConvs.Clear;
  for I := 0 to FLinks.Count - 1 do
    TSrvLink(FLinks[I]).Free;
  FLinks.Clear;
  if GServer = Self then GServer := nil;
  Log('service "' + FService + '" stopped');
end;

function TDdeSimServer.Running: Boolean;
begin
  Result := FInst <> 0;
end;

function TDdeSimServer.ItemCount: Integer;
begin
  Result := FItems.Count;
end;

function TDdeSimServer.Item(I: Integer): TSrvItem;
begin
  Result := TSrvItem(FItems.Objects[I]);
end;

function TDdeSimServer.ConvCount: Integer;
begin
  Result := FConvs.Count;
end;

function TDdeSimServer.Conv(I: Integer): TSrvConv;
begin
  Result := TSrvConv(FConvs[I]);
end;

function TDdeSimServer.LinkCount: Integer;
begin
  Result := FLinks.Count;
end;

function TDdeSimServer.Link(I: Integer): TSrvLink;
begin
  Result := TSrvLink(FLinks[I]);
end;

function TDdeSimServer.LinksOf(const ATopic, AItem: string): Integer;
var
  I: Integer;
begin
  I := FAdvised.IndexOf(Key(ATopic, AItem));
  if I < 0 then
    Result := 0
  else
    Result := TAdvised(FAdvised.Objects[I]).Count;
end;

function TDdeSimServer.FindConv(C: HCONV): TSrvConv;
var
  I: Integer;
begin
  for I := 0 to FConvs.Count - 1 do
    if TSrvConv(FConvs[I]).Conv = C then
      Exit(TSrvConv(FConvs[I]));
  Result := nil;
end;

function TDdeSimServer.DataCount: Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to FItems.Count - 1 do
    if TSrvItem(FItems.Objects[I]).Kind = ikData then Inc(Result);
end;

{ the topic as this server spells it, '' when unknown }
function TDdeSimServer.TopicName(const S: string): string;
var
  I: Integer;
begin
  for I := 0 to High(SrvTopics) do
    if SameText(S, SrvTopics[I]) then
      Exit(SrvTopics[I]);
  Result := '';
end;

{ ---------------------------------------------------------------- values }

function TDdeSimServer.NumText(X: Double; Digits: Integer): string;
begin
  if IsNan(X) or IsInfinite(X) then
    Exit('#NUM!');
  if FDecComma then FFS.DecimalSeparator := ',' else FFS.DecimalSeparator := '.';
  Result := FloatToStrF(X, ffGeneral, Digits, 0, FFS);
end;

{ one array tag as tab separated values }
function TDdeSimServer.RowText(Tag: Integer): string;
var
  V: TUAVariant;
  I: Integer;
begin
  Result := '';
  if not FSim.ReadTag(Tag, V) then Exit;
  for I := 0 to High(V.D) do
  begin
    if I > 0 then Result := Result + #9;
    Result := Result + NumText(V.D[I], 7);
  end;
end;

function TDdeSimServer.TypeText(It: TSrvItem): string;
begin
  case It.Kind of
    ikCurve: Result := '2 rows x ' + IntToStr(FSim.Points);
    ikCurvePart: Result := '1 row x ' + IntToStr(FSim.Points);
    ikData: Result := 'text';
  else
    case It.TagKind of
      tkReal: Result := 'Real';
      tkInt: Result := 'Int';
      tkBool: Result := 'Bool (1/0)';
      tkUInt16: Result := 'UInt16';
      tkDouble: Result := 'Double';
      tkDateTime: Result := 'date time';
    else
      Result := '?';
    end;
  end;
end;

function TDdeSimServer.Value(const ATopic, AItem: string; out S: string): Boolean;
var
  I: Integer;
  It: TSrvItem;
  V: TUAVariant;
  Topic: string;
begin
  S := '';
  Result := True;
  Topic := TopicName(ATopic);
  if Topic = '' then Exit(False);
  if SameText(AItem, SZDDE_ITEM_ITEMLIST) then
  begin
    if Topic = SZDDESYS_TOPIC then
      for I := 0 to High(SysItems) do
        S := S + SysItems[I] + #9
    else
      for I := 0 to FItems.Count - 1 do
        if TSrvItem(FItems.Objects[I]).Topic = Topic then
          S := S + TSrvItem(FItems.Objects[I]).Name + #9;
    if S <> '' then SetLength(S, Length(S) - 1);
    Exit;
  end;
  if SameText(AItem, SZDDESYS_ITEM_FORMATS) then
  begin
    S := 'TEXT'#9'UNICODETEXT';
    Exit;
  end;
  if Topic = SZDDESYS_TOPIC then
  begin
    if SameText(AItem, SZDDESYS_ITEM_TOPICS) then
    begin
      for I := 0 to High(SrvTopics) do
      begin
        if I > 0 then S := S + #9;
        S := S + SrvTopics[I];
      end;
    end
    else if SameText(AItem, SZDDESYS_ITEM_SYSITEMS) then
    begin
      for I := 0 to High(SysItems) do
      begin
        if I > 0 then S := S + #9;
        S := S + SysItems[I];
      end;
    end
    else if SameText(AItem, SZDDESYS_ITEM_STATUS) then
      S := 'Ready'
    else if SameText(AItem, SZDDESYS_ITEM_HELP) then
      S := HelpText
    else if SameText(AItem, SZDDESYS_ITEM_RTNMSG) then
      S := FReturnMsg
    else
      Result := False;
    Exit;
  end;
  I := FItems.IndexOf(Key(Topic, AItem));
  if I < 0 then Exit(False);
  It := TSrvItem(FItems.Objects[I]);
  case It.Kind of
    ikData: S := It.Data;
    ikCurve: S := RowText(It.Tag) + #13#10 + RowText(It.VoltTag);
    ikCurvePart: S := RowText(It.Tag);
  else
    if It.TagKind = tkDateTime then
      S := FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz', Now)
    else if not FSim.ReadTag(It.Tag, V) or (Length(V.D) = 0) then
      Result := False
    else
      case It.TagKind of
        tkBool: S := IntToStr(Ord(V.D[0] <> 0));
        tkReal: S := NumText(V.D[0], 7);
        tkDouble: S := NumText(V.D[0], 10);
      else
        S := IntToStr(Round(V.D[0]));
      end;
  end;
end;

{ ---------------------------------------------------------------- writing }

function TDdeSimServer.ParseNum(const S: string; out X: Double): Boolean;
var
  F: TFormatSettings;
begin
  F := DefaultFormatSettings;
  F.DecimalSeparator := '.';
  F.ThousandSeparator := #0;
  Result := TryStrToFloat(StringReplace(Trim(S), ',', '.', []), X, F);
end;

{ rows by CR / LF, values by tab, ';', blank (and ',' unless decimal comma) }
function TDdeSimServer.ParseRows(const S: string; out Rows: TStringList): Boolean;
var
  Lines: TStringList;
  I, J: Integer;
  L: string;
begin
  Rows := TStringList.Create;
  Lines := TStringList.Create;
  try
    Lines.Text := S;
    for I := 0 to Lines.Count - 1 do
    begin
      L := Lines[I];
      for J := 1 to Length(L) do
        if (L[J] in [#9, ';']) or ((L[J] = ',') and not FDecComma) then
          L[J] := ' ';
      if Trim(L) <> '' then
        Rows.Add(Trim(L));
    end;
    Result := Rows.Count > 0;
  finally
    Lines.Free;
  end;
end;

function TDdeSimServer.WriteItem(It: TSrvItem; const S: string; out Msg: string): Boolean;

  function RowVariant(const Row: string; out V: TUAVariant): Boolean;
  var
    P: TStringList;
    D: array of Double;
    I: Integer;
  begin
    Result := False;
    P := TStringList.Create;
    try
      P.Delimiter := ' ';
      P.StrictDelimiter := False;
      P.DelimitedText := Row;
      for I := P.Count - 1 downto 0 do
        if Trim(P[I]) = '' then P.Delete(I);
      if P.Count <> FSim.Points then
      begin
        Msg := Format('%d values instead of %d', [P.Count, FSim.Points]);
        Exit;
      end;
      SetLength(D, P.Count);
      for I := 0 to P.Count - 1 do
        if not ParseNum(P[I], D[I]) then
        begin
          Msg := '"' + P[I] + '" is not a number';
          Exit;
        end;
      V := VarArray(vtFloat, D);
      Result := True;
    finally
      P.Free;
    end;
  end;

var
  T, StatusMsg: string;
  X: Double;
  Code: Cardinal;
  Rows: TStringList;
  VE, VV: TUAVariant;
begin
  Msg := '';
  Result := False;
  if It.Kind = ikData then
  begin
    It.Data := S;
    Exit(True);
  end;
  if not It.Writable then
  begin
    Msg := 'read only';
    Exit;
  end;
  T := LowerCase(Trim(S));
  Code := UA_GOOD;
  case It.Kind of
    ikCurve, ikCurvePart:
      begin
        ParseRows(S, Rows);
        try
          if (It.Kind = ikCurve) and (Rows.Count <> 2) then
          begin
            Msg := '2 rows (Eng, Volt) expected';
            Exit;
          end;
          if (It.Kind = ikCurvePart) and (Rows.Count <> 1) then
          begin
            Msg := '1 row expected';
            Exit;
          end;
          if not RowVariant(Rows[0], VE) then Exit;
          if (It.Kind = ikCurve) and not RowVariant(Rows[1], VV) then Exit;
          { both rows in one PLC cycle, else Set 2 could check a half written curve }
          FSim.Lock;
          try
            Code := FSim.WriteTag(It.Tag, VE);
            if (Code = UA_GOOD) and (It.Kind = ikCurve) then
              Code := FSim.WriteTag(It.VoltTag, VV);
          finally
            FSim.Unlock;
          end;
        finally
          Rows.Free;
        end;
      end;
  else
    case It.TagKind of
      tkBool:
        begin
          if (T = '1') or (T = 'true') or (T = 'on') or (T = 'wahr') then
            X := 1
          else if (T = '0') or (T = 'false') or (T = 'off') or (T = 'falsch') then
            X := 0
          else
          begin
            Msg := 'expected 1/0/true/false/on/off';
            Exit;
          end;
          Code := FSim.WriteTag(It.Tag, VarScalar(vtBoolean, X));
        end;
      tkReal, tkInt:
        begin
          if not ParseNum(T, X) then
          begin
            Msg := 'not a number';
            Exit;
          end;
          if It.TagKind = tkReal then
            Code := FSim.WriteTag(It.Tag, VarScalar(vtFloat, X))
          else
            Code := FSim.WriteTag(It.Tag, VarScalar(vtInt16, X));
        end;
    else
      Code := UA_BadNotWritable;
    end;
  end;
  if Code <> UA_GOOD then
  begin
    case Code of
      UA_BadTypeMismatch: StatusMsg := 'wrong type or number of values';
      UA_BadOutOfRange: StatusMsg := 'out of range';
      UA_BadNotWritable: StatusMsg := 'read only';
    else
      StatusMsg := Format('error 0x%.8x', [Code]);
    end;
    Msg := StatusMsg;
    Exit;
  end;
  Result := True;
end;

{ [Set(item,value)][Enable][Disable][Ack][Clear], also without brackets }
function TDdeSimServer.Execute(const ATopic, Cmds: string; out Msg: string): Boolean;
var
  P, Q, Done: Integer;
  Cmd, Name, Args, ItemName, Val, T, M: string;
  I: Integer;
  It: TSrvItem;
  List: TStringList;
begin
  Msg := '';
  Result := True;
  List := TStringList.Create;
  try
    { split into commands }
    T := Trim(Cmds);
    if (T <> '') and (T[1] <> '[') then
      List.Add(T)
    else
    begin
      P := 1;
      while P <= Length(T) do
      begin
        if T[P] = '[' then
        begin
          Q := P + 1;
          Done := 0;                    // ']' inside quotes or parentheses belongs to the command
          while (Q <= Length(T)) and not ((T[Q] = ']') and (Done = 0)) do
          begin
            if T[Q] = '(' then Inc(Done)
            else if (T[Q] = ')') and (Done > 0) then Dec(Done);
            Inc(Q);
          end;
          List.Add(Trim(Copy(T, P + 1, Q - P - 1)));
          P := Q + 1;
        end
        else
          Inc(P);
      end;
    end;
    if List.Count = 0 then
    begin
      Msg := 'no command';
      Exit(False);
    end;
    for I := 0 to List.Count - 1 do
    begin
      Cmd := List[I];
      P := Pos('(', Cmd);
      if P > 0 then
      begin
        Name := Trim(Copy(Cmd, 1, P - 1));
        Args := Copy(Cmd, P + 1, MaxInt);
        Q := LastDelimiter(')', Args);
        if Q > 0 then SetLength(Args, Q - 1);
      end
      else
      begin
        Name := Trim(Cmd);
        Args := '';
      end;
      Name := LowerCase(Name);
      M := '';
      if Name = 'set' then
      begin
        P := Pos(',', Args);
        if P = 0 then
          M := 'Set(item,value) expected'
        else
        begin
          ItemName := Trim(Copy(Args, 1, P - 1));
          Val := Trim(Copy(Args, P + 1, MaxInt));
          ItemName := StringReplace(ItemName, '"', '', [rfReplaceAll]);
          Val := StringReplace(Val, '"', '', [rfReplaceAll]);
          T := ATopic;
          if Pos('!', ItemName) > 0 then       // Topic!Item, from the System topic
          begin
            T := TopicName(Copy(ItemName, 1, Pos('!', ItemName) - 1));
            ItemName := Copy(ItemName, Pos('!', ItemName) + 1, MaxInt);
          end;
          Q := FItems.IndexOf(Key(T, ItemName));
          if (Q < 0) and SameText(T, TOPIC_DATA) and (ItemName <> '') and (DataCount < MAX_DATA_ITEMS) then
          begin
            AddItem(TOPIC_DATA, ItemName, ikData, tkReal, -1, -1, True);
            Q := FItems.IndexOf(Key(T, ItemName));
          end;
          if Q < 0 then
            M := 'unknown item ' + T + '!' + ItemName
          else
          begin
            It := TSrvItem(FItems.Objects[Q]);
            if WriteItem(It, Val, M) then
            begin
              Log(Format('  set %s!%s = %s', [It.Topic, It.Name, Val]));
              if It.Kind = ikData then Tick;
            end;
          end;
        end;
      end
      else if (Name = 'enable') or (Name = 'disable') then
      begin
        It := TSrvItem(FItems.Objects[FItems.IndexOf(Key(TOPIC_HMI, 'Enable'))]);
        WriteItem(It, IntToStr(Ord(Name = 'enable')), M);
      end
      else if Name = 'ack' then
      begin
        if FSim.Points <> 2 then
          M := 'Ack: only with 2 curve points (Set 2)'
        else
          for Q := 0 to FItems.Count - 1 do
          begin
            It := TSrvItem(FItems.Objects[Q]);
            if (It.Topic = TOPIC_HMI) and It.Name.StartsWith('Err_') then
              WriteItem(It, '0', M);
          end;
      end
      else if Name = 'clear' then
      begin
        for Q := FItems.Count - 1 downto 0 do
          if TSrvItem(FItems.Objects[Q]).Kind = ikData then
            if LinksOf(TOPIC_DATA, TSrvItem(FItems.Objects[Q]).Name) = 0 then
              FItems.Delete(Q)
            else
              TSrvItem(FItems.Objects[Q]).Data := '';
        Tick;
      end
      else
        M := 'unknown command "' + Cmd + '"';
      if M <> '' then
      begin
        if Msg <> '' then Msg := Msg + '; ';
        Msg := Msg + M;
        Result := False;
      end;
    end;
  finally
    List.Free;
  end;
end;

{ ---------------------------------------------------------------- links }

procedure TDdeSimServer.AddLink(C: HCONV; const ATopic, AItem: string; Fmt: UINT);
var
  L: TSrvLink;
  A: TAdvised;
  I: Integer;
begin
  L := TSrvLink.Create;
  L.Conv := C;
  L.Topic := ATopic;
  L.Item := AItem;
  L.Fmt := Fmt;
  FLinks.Add(L);
  I := FAdvised.IndexOf(Key(ATopic, AItem));
  if I < 0 then
  begin
    A := TAdvised.Create;
    A.Topic := ATopic;
    A.Item := AItem;
    A.HTopic := HszOf(FInst, ATopic);
    A.HItem := HszOf(FInst, AItem);
    Value(ATopic, AItem, A.Last);
    FAdvised.AddObject(Key(ATopic, AItem), A);
  end
  else
    A := TAdvised(FAdvised.Objects[I]);
  Inc(A.Count);
end;

procedure TDdeSimServer.RemoveLink(C: HCONV; const ATopic, AItem: string; Fmt: UINT);
var
  I, J: Integer;
  L: TSrvLink;
  A: TAdvised;
begin
  for I := FLinks.Count - 1 downto 0 do
  begin
    L := TSrvLink(FLinks[I]);
    if (L.Conv <> C) or not SameText(L.Topic, ATopic) or not SameText(L.Item, AItem)
      or ((Fmt <> 0) and (L.Fmt <> Fmt)) then
      Continue;
    J := FAdvised.IndexOf(Key(L.Topic, L.Item));
    if J >= 0 then
    begin
      A := TAdvised(FAdvised.Objects[J]);
      Dec(A.Count);
      if A.Count <= 0 then
      begin
        DdeFreeStringHandle(FInst, A.HTopic);
        DdeFreeStringHandle(FInst, A.HItem);
        FAdvised.Delete(J);
      end;
    end;
    L.Free;
    FLinks.Delete(I);
    if Fmt <> 0 then Exit;
  end;
end;

procedure TDdeSimServer.RemoveConvLinks(C: HCONV);
var
  I: Integer;
  L: TSrvLink;
begin
  for I := FLinks.Count - 1 downto 0 do
    if I < FLinks.Count then
    begin
      L := TSrvLink(FLinks[I]);
      if L.Conv = C then
        RemoveLink(C, L.Topic, L.Item, L.Fmt);
    end;
end;

procedure TDdeSimServer.Tick;
var
  I: Integer;
  A: TAdvised;
  S: string;
begin
  if FInst = 0 then Exit;
  for I := 0 to FAdvised.Count - 1 do
  begin
    A := TAdvised(FAdvised.Objects[I]);
    if not Value(A.Topic, A.Item, S) then S := '';
    if S <> A.Last then
    begin
      A.Last := S;
      DdePostAdvise(FInst, A.HTopic, A.HItem);
    end;
  end;
end;

{ ---------------------------------------------------------------- DDEML }

function TDdeSimServer.Callback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ;
  hData: HDDEDATA; dwData1, dwData2: PtrUInt): HDDEDATA;
var
  Topic, ItemName, S, Msg, Svc: string;
  C: TSrvConv;
  I, N: Integer;
  It: TSrvItem;
  Pairs: array of THszPair;
begin
  Result := 0;
  wType := wType and not (XTYPF_NODATA or XTYPF_ACKREQ);
  case wType of
    XTYP_CONNECT:
      begin
        { hsz1 = topic, hsz2 = service (DDEML already filtered the service) }
        Topic := TopicName(HszText(FInst, hsz1));
        if Topic <> '' then
          Result := 1
        else
        begin
          Inc(Refused);
          Log('connect refused: unknown topic "' + HszText(FInst, hsz1) + '"');
        end;
      end;
    XTYP_WILDCONNECT:
      begin
        { "which topics do you have?" (hsz1 = topic or 0, hsz2 = service or 0) }
        Svc := HszText(FInst, hsz2);
        if (hsz2 <> 0) and not SameText(Svc, FService) then Exit(0);
        Topic := HszText(FInst, hsz1);
        SetLength(Pairs, 0);
        for I := 0 to High(SrvTopics) do
          if (hsz1 = 0) or SameText(Topic, SrvTopics[I]) then
          begin
            N := Length(Pairs);
            SetLength(Pairs, N + 1);
            Pairs[N].hszSvc := FHService;
            Pairs[N].hszTopic := FHTopics[I];
          end;
        if Length(Pairs) = 0 then Exit(0);
        N := Length(Pairs);
        SetLength(Pairs, N + 1);
        Pairs[N].hszSvc := 0;
        Pairs[N].hszTopic := 0;
        Result := DdeCreateDataHandle(FInst, @Pairs[0], Length(Pairs) * SizeOf(THszPair), 0, 0,
          0, 0);
      end;
    XTYP_CONNECT_CONFIRM:
      begin
        C := TSrvConv.Create;
        C.Conv := hConv;
        C.Topic := TopicName(HszText(FInst, hsz1));
        C.Process := PartnerProcess(hConv);
        C.Since := Now;
        FConvs.Add(C);
        Log(Format('connect %s, topic %s', [C.Process, C.Topic]));
      end;
    XTYP_DISCONNECT:
      begin
        C := FindConv(hConv);
        RemoveConvLinks(hConv);
        if C <> nil then
        begin
          Log(Format('disconnect %s, topic %s (%d transactions)',
            [C.Process, C.Topic, C.Transactions]));
          FConvs.Remove(C);
          C.Free;
        end;
      end;
    XTYP_REQUEST, XTYP_ADVREQ:
      begin
        if (wFmt <> CF_TEXT) and (wFmt <> CF_UNICODETEXT) then Exit(0);
        Topic := HszText(FInst, hsz1);
        ItemName := HszText(FInst, hsz2);
        C := FindConv(hConv);
        if C <> nil then Inc(C.Transactions);
        if not Value(Topic, ItemName, S) then
        begin
          Inc(Refused);
          Log(Format('request refused: %s!%s unknown', [Topic, ItemName]));
          Exit(0);
        end;
        if wType = XTYP_REQUEST then
        begin
          Inc(Requests);
          if FLogRequests then
            Log(Format('request %s!%s (%s) = %s', [Topic, ItemName, FormatName(wFmt), OneLine(S)]));
        end
        else
        begin
          Inc(Advises);
          if FLogRequests then
            Log(Format('advise %s!%s = %s', [Topic, ItemName, OneLine(S)]));
        end;
        Result := TextData(FInst, S, hsz2, wFmt);
      end;
    XTYP_ADVSTART:
      begin
        if (wFmt <> CF_TEXT) and (wFmt <> CF_UNICODETEXT) then Exit(0);
        Topic := TopicName(HszText(FInst, hsz1));
        ItemName := HszText(FInst, hsz2);
        if not Value(Topic, ItemName, S) then
        begin
          { a link to a new Data item creates it }
          if (Topic = TOPIC_DATA) and (ItemName <> '') and (DataCount < MAX_DATA_ITEMS) then
            AddItem(TOPIC_DATA, ItemName, ikData, tkReal, -1, -1, True)
          else
          begin
            Inc(Refused);
            Log(Format('link refused: %s!%s unknown', [Topic, ItemName]));
            Exit(0);
          end;
        end;
        I := FItems.IndexOf(Key(Topic, ItemName));
        if I >= 0 then ItemName := TSrvItem(FItems.Objects[I]).Name;
        AddLink(hConv, Topic, ItemName, wFmt);
        C := FindConv(hConv);
        if C <> nil then Inc(C.Transactions);
        Log(Format('link %s!%s (%s)', [Topic, ItemName, FormatName(wFmt)]));
        Result := 1;
      end;
    XTYP_ADVSTOP:
      begin
        Topic := TopicName(HszText(FInst, hsz1));
        ItemName := HszText(FInst, hsz2);
        RemoveLink(hConv, Topic, ItemName, wFmt);
        Log(Format('unlink %s!%s', [Topic, ItemName]));
      end;
    XTYP_POKE:
      begin
        Result := DDE_FNOTPROCESSED;
        Inc(Pokes);
        C := FindConv(hConv);
        if C <> nil then Inc(C.Transactions);
        Topic := TopicName(HszText(FInst, hsz1));
        ItemName := HszText(FInst, hsz2);
        if (wFmt <> CF_TEXT) and (wFmt <> CF_UNICODETEXT) then
        begin
          Inc(Refused);
          FReturnMsg := 'poke ' + Topic + '!' + ItemName + ': format ' + FormatName(wFmt)
            + ' not supported, use TEXT';
          Log(FReturnMsg);
          Exit;
        end;
        S := DataText(hData, wFmt);
        I := FItems.IndexOf(Key(Topic, ItemName));
        if (I < 0) and (Topic = TOPIC_DATA) and (ItemName <> '') and (DataCount < MAX_DATA_ITEMS) then
        begin
          AddItem(TOPIC_DATA, ItemName, ikData, tkReal, -1, -1, True);
          I := FItems.IndexOf(Key(Topic, ItemName));
        end;
        if I < 0 then
        begin
          Inc(Refused);
          FReturnMsg := 'poke refused: ' + Topic + '!' + ItemName + ' unknown';
          Log(FReturnMsg);
          Exit;
        end;
        It := TSrvItem(FItems.Objects[I]);
        Svc := OneLine(S);
        if WriteItem(It, S, Msg) then
        begin
          FReturnMsg := '';
          Log(Format('poke %s!%s = %s', [It.Topic, It.Name, Svc]));
          Result := DDE_FACK;
          if It.Kind = ikData then Tick;     // the other clients see it at once
        end
        else
        begin
          Inc(Refused);
          FReturnMsg := Format('poke refused: %s!%s = %s (%s)', [It.Topic, It.Name, Svc, Msg]);
          Log(FReturnMsg);
        end;
      end;
    XTYP_EXECUTE:
      begin
        Inc(Executes);
        C := FindConv(hConv);
        if C <> nil then Inc(C.Transactions);
        Topic := TopicName(HszText(FInst, hsz1));
        { DDEML converts execute strings between ANSI and Unicode programs }
        S := DataText(hData, CF_UNICODETEXT);
        if Execute(Topic, S, Msg) then
        begin
          FReturnMsg := '';
          Log(Format('execute on %s: %s', [Topic, S]));
          Result := DDE_FACK;
        end
        else
        begin
          Inc(Refused);
          FReturnMsg := 'execute refused: ' + Msg;
          Log(Format('execute refused on %s: %s (%s)', [Topic, S, Msg]));
          Result := DDE_FNOTPROCESSED;
        end;
      end;
    XTYP_ERROR:
      Log('DDEML error: ' + DdeErrorText(LoWord(dwData1)));
  end;
end;

end.
