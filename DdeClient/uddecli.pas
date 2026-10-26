{ ============================================================================
  uddecli - general DDE client (DDEML) for any DDE server
    - conversations per service|topic, opened when needed
    - request, poke, execute (synchronous, with timeout)
    - hot links (the data comes with every change) and warm links (only a
      note of the change, the client requests the value)
    - CF_TEXT or CF_UNICODETEXT
    - list of the running servers (service|topic) through DdeConnectList
    - a server that ends or restarts: the links are made again (register
      notification or every few seconds)
  DDEML calls back in the thread that called Create (the main thread). A
  synchronous transaction runs a message loop of its own, so a timer can come
  while it waits: Tick does nothing while a transaction runs.
  ============================================================================ }
unit uddecli;

{$mode objfpc}{$H+}

interface

uses
  Windows, Classes, SysUtils, uddeml;

type
  TCliConv = class
    Service: string;
    Topic: string;
    Conv: HCONV;
  end;

  TLinkState = (lsActive, lsDisconnected, lsRefused);

  TCliLink = class
    Service: string;
    Topic: string;
    Item: string;
    Warm: Boolean;
    Fmt: UINT;
    Value: string;
    Updates: Integer;
    Time: TDateTime;
    State: TLinkState;
    Pending: Boolean;          // warm link: changed, request the value
    Conv: HCONV;
  end;

  TCliLogEvent = procedure(const S: string) of object;

  { TDdeClient }

  TDdeClient = class
  private
    FInst: DWORD;
    FConvs: TList;
    FLinks: TList;
    FBusy: Integer;
    FTimeout: DWORD;
    FAutoReconnect: Boolean;
    FLastRetry: QWord;
    FRetryNow: Boolean;
    FOnLog: TCliLogEvent;
    FOnData: TNotifyEvent;
    procedure Log(const S: string);
    function FindConv(const AService, ATopic: string): TCliConv;
    function ConvOf(C: HCONV): TCliConv;
    function Transact(L: TCliLink; wType: UINT; out Err: string): Boolean;
    function Callback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ; hData: HDDEDATA;
      dwData1, dwData2: PtrUInt): HDDEDATA;
    procedure Relink;
  public
    Transactions, Errors, Advises: Int64;
    { set after a Disconnect by the user: no reconnects until False again }
    Paused: Boolean;
    constructor Create;
    destructor Destroy; override;
    { the conversation with service|topic, made when there is none }
    function Connect(const AService, ATopic: string; out Err: string): HCONV;
    procedure Disconnect(const AService, ATopic: string);
    procedure DisconnectAll;
    function Connected(const AService, ATopic: string): Boolean;
    function ConvCount: Integer;
    function Conv(I: Integer): TCliConv;
    function Request(const AService, ATopic, AItem: string; Fmt: UINT; out Value, Err: string): Boolean;
    function Poke(const AService, ATopic, AItem, Value: string; Fmt: UINT; out Err: string): Boolean;
    function Execute(const AService, ATopic, Cmd: string; out Err: string): Boolean;
    function AddLink(const AService, ATopic, AItem: string; AWarm: Boolean; Fmt: UINT;
      out Err: string): TCliLink;
    procedure RemoveLink(L: TCliLink);
    function LinkCount: Integer;
    function Link(I: Integer): TCliLink;
    { the running servers as service|topic }
    function FindServers(List: TStrings): Integer;
    { warm links: request the changed values; reconnect lost links }
    procedure Tick;
    function Busy: Boolean;
    property Timeout: DWORD read FTimeout write FTimeout;
    property AutoReconnect: Boolean read FAutoReconnect write FAutoReconnect;
    property OnLog: TCliLogEvent read FOnLog write FOnLog;
    property OnData: TNotifyEvent read FOnData write FOnData;
  end;

function LinkName(const AService, ATopic, AItem: string): string;

implementation

var
  GClient: TDdeClient = nil;

const
  RETRY_MS = 3000;

function ClientCallback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ; hData: HDDEDATA;
  dwData1, dwData2: PtrUInt): HDDEDATA; stdcall;
begin
  Result := 0;
  if GClient <> nil then
  try
    Result := GClient.Callback(wType, wFmt, hConv, hsz1, hsz2, hData, dwData1, dwData2);
  except
    on E: Exception do
      GClient.Log('error in ' + XtypName(wType) + ': ' + E.Message);
  end;
end;

function LinkName(const AService, ATopic, AItem: string): string;
begin
  Result := AService + '|' + ATopic + '!' + AItem;
end;

{ TDdeClient }

constructor TDdeClient.Create;
var
  R: UINT;
begin
  if GClient <> nil then
    raise Exception.Create('only one DDE client per program');
  FConvs := TList.Create;
  FLinks := TList.Create;
  FTimeout := 3000;
  FAutoReconnect := True;
  GClient := Self;
  R := DdeInitializeW(FInst, @ClientCallback, APPCMD_CLIENTONLY, 0);
  if R <> DMLERR_NO_ERROR then
  begin
    GClient := nil;
    raise Exception.Create('DdeInitialize: ' + DdeErrorText(R));
  end;
end;

destructor TDdeClient.Destroy;
var
  I: Integer;
begin
  DisconnectAll;
  for I := 0 to FLinks.Count - 1 do
    TCliLink(FLinks[I]).Free;
  FLinks.Free;
  if FInst <> 0 then
    DdeUninitialize(FInst);
  FConvs.Free;
  if GClient = Self then GClient := nil;
  inherited Destroy;
end;

procedure TDdeClient.Log(const S: string);
begin
  if Assigned(FOnLog) then FOnLog(S);
end;

function TDdeClient.Busy: Boolean;
begin
  Result := FBusy > 0;
end;

function TDdeClient.FindConv(const AService, ATopic: string): TCliConv;
var
  I: Integer;
begin
  for I := 0 to FConvs.Count - 1 do
    if SameText(TCliConv(FConvs[I]).Service, AService) and
      SameText(TCliConv(FConvs[I]).Topic, ATopic) then
      Exit(TCliConv(FConvs[I]));
  Result := nil;
end;

function TDdeClient.ConvOf(C: HCONV): TCliConv;
var
  I: Integer;
begin
  for I := 0 to FConvs.Count - 1 do
    if TCliConv(FConvs[I]).Conv = C then
      Exit(TCliConv(FConvs[I]));
  Result := nil;
end;

function TDdeClient.ConvCount: Integer;
begin
  Result := FConvs.Count;
end;

function TDdeClient.Conv(I: Integer): TCliConv;
begin
  Result := TCliConv(FConvs[I]);
end;

function TDdeClient.LinkCount: Integer;
begin
  Result := FLinks.Count;
end;

function TDdeClient.Link(I: Integer): TCliLink;
begin
  Result := TCliLink(FLinks[I]);
end;

function TDdeClient.Connected(const AService, ATopic: string): Boolean;
begin
  Result := FindConv(AService, ATopic) <> nil;
end;

function TDdeClient.Connect(const AService, ATopic: string; out Err: string): HCONV;
var
  C: TCliConv;
  HS, HT: HSZ;
begin
  Err := '';
  C := FindConv(AService, ATopic);
  if C <> nil then Exit(C.Conv);
  if (Trim(AService) = '') or (Trim(ATopic) = '') then
  begin
    Err := 'service and topic are needed';
    Exit(0);
  end;
  HS := HszOf(FInst, AService);
  HT := HszOf(FInst, ATopic);
  Inc(FBusy);
  try
    Result := DdeConnect(FInst, HS, HT, nil);
  finally
    Dec(FBusy);
    DdeFreeStringHandle(FInst, HS);
    DdeFreeStringHandle(FInst, HT);
  end;
  if Result = 0 then
  begin
    Err := DdeErrorText(DdeGetLastError(FInst));
    Exit;
  end;
  C := TCliConv.Create;
  C.Service := AService;
  C.Topic := ATopic;
  C.Conv := Result;
  FConvs.Add(C);
  Log(Format('connected to %s|%s (%s)', [AService, ATopic, PartnerProcess(Result)]));
end;

procedure TDdeClient.Disconnect(const AService, ATopic: string);
var
  C: TCliConv;
  I: Integer;
begin
  C := FindConv(AService, ATopic);
  if C = nil then Exit;
  for I := 0 to FLinks.Count - 1 do
    with TCliLink(FLinks[I]) do
      if Conv = C.Conv then
      begin
        Conv := 0;
        State := lsDisconnected;
      end;
  FConvs.Remove(C);
  DdeDisconnect(C.Conv);
  Log(Format('disconnected from %s|%s', [C.Service, C.Topic]));
  C.Free;
end;

procedure TDdeClient.DisconnectAll;
begin
  while FConvs.Count > 0 do
    Disconnect(TCliConv(FConvs[0]).Service, TCliConv(FConvs[0]).Topic);
end;

{ one synchronous transaction for a link (ADVSTART / ADVSTOP / REQUEST) }
function TDdeClient.Transact(L: TCliLink; wType: UINT; out Err: string): Boolean;
var
  H: HSZ;
  R: HDDEDATA;
  Res: DWORD;
begin
  Err := '';
  if L.Conv = 0 then
  begin
    L.Conv := Connect(L.Service, L.Topic, Err);
    if L.Conv = 0 then Exit(False);
  end;
  H := HszOf(FInst, L.Item);
  Inc(FBusy);
  try
    Res := 0;
    R := DdeClientTransaction(nil, 0, L.Conv, H, L.Fmt, wType, FTimeout, @Res);
    Inc(Transactions);
    Result := R <> 0;
    if not Result then
    begin
      Inc(Errors);
      Err := DdeErrorText(DdeGetLastError(FInst));
    end
    else if wType = XTYP_REQUEST then
    begin
      L.Value := DataText(R, L.Fmt);
      DdeFreeDataHandle(R);
    end;
  finally
    Dec(FBusy);
    DdeFreeStringHandle(FInst, H);
  end;
end;

function TDdeClient.Request(const AService, ATopic, AItem: string; Fmt: UINT;
  out Value, Err: string): Boolean;
var
  C: HCONV;
  H: HSZ;
  R: HDDEDATA;
  Res: DWORD;
begin
  Value := '';
  C := Connect(AService, ATopic, Err);
  if C = 0 then Exit(False);
  H := HszOf(FInst, AItem);
  Inc(FBusy);
  try
    Res := 0;
    R := DdeClientTransaction(nil, 0, C, H, Fmt, XTYP_REQUEST, FTimeout, @Res);
    Inc(Transactions);
    Result := R <> 0;
    if Result then
    begin
      Value := DataText(R, Fmt);
      DdeFreeDataHandle(R);
    end
    else
    begin
      Inc(Errors);
      Err := DdeErrorText(DdeGetLastError(FInst));
    end;
  finally
    Dec(FBusy);
    DdeFreeStringHandle(FInst, H);
  end;
end;

function TDdeClient.Poke(const AService, ATopic, AItem, Value: string; Fmt: UINT;
  out Err: string): Boolean;
var
  C: HCONV;
  H: HSZ;
  B: RawByteString;
  Res: DWORD;
begin
  C := Connect(AService, ATopic, Err);
  if C = 0 then Exit(False);
  H := HszOf(FInst, AItem);
  B := TextBytes(Value, Fmt);
  Inc(FBusy);
  try
    Res := 0;
    Result := DdeClientTransaction(PAnsiChar(B), Length(B), C, H, Fmt, XTYP_POKE, FTimeout,
      @Res) <> 0;
    Inc(Transactions);
    if not Result then
    begin
      Inc(Errors);
      Err := DdeErrorText(DdeGetLastError(FInst));
    end;
  finally
    Dec(FBusy);
    DdeFreeStringHandle(FInst, H);
  end;
end;

function TDdeClient.Execute(const AService, ATopic, Cmd: string; out Err: string): Boolean;
var
  C: HCONV;
  B: RawByteString;
  Res: DWORD;
begin
  C := Connect(AService, ATopic, Err);
  if C = 0 then Exit(False);
  { a Unicode DDEML program sends execute strings as UTF-16, DDEML converts
    them for ANSI servers }
  B := TextBytes(Cmd, CF_UNICODETEXT);
  Inc(FBusy);
  try
    Res := 0;
    Result := DdeClientTransaction(PAnsiChar(B), Length(B), C, 0, 0, XTYP_EXECUTE, FTimeout,
      @Res) <> 0;
    Inc(Transactions);
    if not Result then
    begin
      Inc(Errors);
      Err := DdeErrorText(DdeGetLastError(FInst));
    end;
  finally
    Dec(FBusy);
  end;
end;

function TDdeClient.AddLink(const AService, ATopic, AItem: string; AWarm: Boolean; Fmt: UINT;
  out Err: string): TCliLink;
var
  I: Integer;
  L: TCliLink;
  T: UINT;
begin
  Result := nil;
  for I := 0 to FLinks.Count - 1 do
  begin
    L := TCliLink(FLinks[I]);
    if SameText(L.Service, AService) and SameText(L.Topic, ATopic) and SameText(L.Item, AItem)
      and (L.Fmt = Fmt) then
    begin
      Err := 'the link is there already';
      Exit;
    end;
  end;
  L := TCliLink.Create;
  L.Service := AService;
  L.Topic := ATopic;
  L.Item := AItem;
  L.Warm := AWarm;
  L.Fmt := Fmt;
  L.State := lsDisconnected;
  FLinks.Add(L);
  Result := L;
  T := XTYP_ADVSTART;
  if AWarm then T := T or XTYPF_NODATA;
  if Transact(L, T, Err) then
  begin
    L.State := lsActive;
    { a link sends only changes: ask for the value it has now }
    L.Pending := True;
    Log(Format('%s link %s (%s)', [BoolToStr(AWarm, 'warm', 'hot'),
      LinkName(AService, ATopic, AItem), FormatName(Fmt)]));
  end
  else
  begin
    if L.Conv <> 0 then L.State := lsRefused;
    Log(Format('link %s failed: %s', [LinkName(AService, ATopic, AItem), Err]));
  end;
end;

procedure TDdeClient.RemoveLink(L: TCliLink);
var
  Err: string;
begin
  if (L.State = lsActive) and (L.Conv <> 0) then
    Transact(L, XTYP_ADVSTOP, Err);
  FLinks.Remove(L);
  Log('link removed: ' + LinkName(L.Service, L.Topic, L.Item));
  L.Free;
end;

function TDdeClient.FindServers(List: TStrings): Integer;
var
  CL: HCONVLIST;
  C: HCONV;
  CI: TConvInfo;
  S: string;
begin
  List.Clear;
  Inc(FBusy);
  try
    CL := DdeConnectList(FInst, 0, 0, 0, nil);
    if CL = 0 then Exit(0);
    C := 0;
    repeat
      C := DdeQueryNextServer(CL, C);
      if C = 0 then Break;
      FillChar(CI, SizeOf(CI), 0);
      CI.cb := SizeOf(CI);
      if DdeQueryConvInfo(C, QID_SYNC, CI) <> 0 then
      begin
        S := HszText(FInst, CI.hszSvcPartner) + '|' + HszText(FInst, CI.hszTopic);
        if List.IndexOf(S) < 0 then
          List.Add(S);
      end;
    until False;
    DdeDisconnectList(CL);
  finally
    Dec(FBusy);
  end;
  Result := List.Count;
end;

{ links without a conversation: connect and start them again }
procedure TDdeClient.Relink;
var
  I: Integer;
  L: TCliLink;
  T: UINT;
  Err: string;
  Failed: TStringList;
begin
  Failed := TStringList.Create;
  try
    for I := 0 to FLinks.Count - 1 do
    begin
      L := TCliLink(FLinks[I]);
      if L.State = lsActive then Continue;
      if Failed.IndexOf(LowerCase(L.Service + '|' + L.Topic)) >= 0 then Continue;
      if (L.Conv = 0) and (FindConv(L.Service, L.Topic) = nil) then
      begin
        L.Conv := Connect(L.Service, L.Topic, Err);
        if L.Conv = 0 then
        begin
          Failed.Add(LowerCase(L.Service + '|' + L.Topic));   // one try per server
          Continue;
        end;
      end
      else if L.Conv = 0 then
        L.Conv := FindConv(L.Service, L.Topic).Conv;
      T := XTYP_ADVSTART;
      if L.Warm then T := T or XTYPF_NODATA;
      if Transact(L, T, Err) then
      begin
        L.State := lsActive;
        L.Pending := True;
        Log('link again: ' + LinkName(L.Service, L.Topic, L.Item));
      end
      else
        L.State := lsRefused;
    end;
  finally
    Failed.Free;
  end;
end;

procedure TDdeClient.Tick;
var
  I: Integer;
  L: TCliLink;
  Err: string;
  Lost: Boolean;
begin
  if FBusy > 0 then Exit;
  for I := 0 to FLinks.Count - 1 do
  begin
    L := TCliLink(FLinks[I]);
    if L.Pending and (L.State = lsActive) then
    begin
      L.Pending := False;
      if Transact(L, XTYP_REQUEST, Err) then
      begin
        L.Time := Now;
        if Assigned(FOnData) then FOnData(Self);
      end
      else
        Log(Format('request %s failed: %s', [LinkName(L.Service, L.Topic, L.Item), Err]));
    end;
  end;
  if not FAutoReconnect or Paused then Exit;
  Lost := False;
  for I := 0 to FLinks.Count - 1 do
    if TCliLink(FLinks[I]).State = lsDisconnected then
      Lost := True;
  if Lost and (FRetryNow or (GetTickCount64 - FLastRetry >= RETRY_MS)) then
  begin
    FRetryNow := False;
    FLastRetry := GetTickCount64;
    Relink;
  end;
end;

function TDdeClient.Callback(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ;
  hData: HDDEDATA; dwData1, dwData2: PtrUInt): HDDEDATA;
var
  I: Integer;
  L: TCliLink;
  ItemName: string;
  C: TCliConv;
begin
  Result := 0;
  case wType of
    XTYP_ADVDATA:
      begin
        ItemName := HszText(FInst, hsz2);
        for I := 0 to FLinks.Count - 1 do
        begin
          L := TCliLink(FLinks[I]);
          if (L.Conv = hConv) and (L.Fmt = wFmt) and SameText(L.Item, ItemName) then
          begin
            Inc(L.Updates);
            Inc(Advises);
            if hData = 0 then
              L.Pending := True          // warm link: only the note of a change
            else
            begin
              L.Value := DataText(hData, wFmt);
              L.Time := Now;
              L.Pending := False;
            end;
          end;
        end;
        if Assigned(FOnData) then FOnData(Self);
        Result := DDE_FACK;
      end;
    XTYP_DISCONNECT:
      begin
        C := ConvOf(hConv);
        if C <> nil then
        begin
          for I := 0 to FLinks.Count - 1 do
            with TCliLink(FLinks[I]) do
              if Conv = hConv then
              begin
                Conv := 0;
                State := lsDisconnected;
              end;
          Log(Format('the server ended the conversation %s|%s', [C.Service, C.Topic]));
          FConvs.Remove(C);
          C.Free;
          FLastRetry := GetTickCount64;
          if Assigned(FOnData) then FOnData(Self);
        end;
      end;
    XTYP_REGISTER:
      begin
        { a server started: links to it can be made again at once }
        ItemName := HszText(FInst, hsz1);
        for I := 0 to FLinks.Count - 1 do
          if SameText(TCliLink(FLinks[I]).Service, ItemName) and
            (TCliLink(FLinks[I]).State <> lsActive) then
          begin
            Log('service ' + ItemName + ' registered');
            FRetryNow := True;
            Break;
          end;
      end;
    XTYP_XACT_COMPLETE, XTYP_UNREGISTER:
      ;
    XTYP_ERROR:
      Log('DDEML error: ' + DdeErrorText(LoWord(dwData1)));
  end;
end;

end.
