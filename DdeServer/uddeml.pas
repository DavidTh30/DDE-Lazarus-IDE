{ ============================================================================
  uddeml - the Windows DDEML API (user32.dll) for 32 and 64 bit, plus helpers
  shared by DdeServer and DdeClient:
    - string handles <-> UnicodeString / UTF-8
    - CF_TEXT (ANSI code page) and CF_UNICODETEXT data <-> UTF-8 text
    - DDEML error texts, transaction names, the process behind a conversation
  The FPC Windows unit has only part of DDEML (and 32 bit types in places),
  so the functions are declared here.
  ============================================================================ }
unit uddeml;

{$mode objfpc}{$H+}
{$PACKRECORDS C}

interface

uses
  Windows, SysUtils;

type
  HCONV = THandle;
  HCONVLIST = THandle;
  HSZ = THandle;
  HDDEDATA = THandle;

  TSecurityQOS = record
    Length: DWORD;
    ImpersonationLevel: Integer;
    ContextTrackingMode: Byte;
    EffectiveOnly: Byte;
  end;

  TConvContext = record
    cb: UINT;
    wFlags: UINT;
    wCountryID: UINT;
    iCodePage: Integer;
    dwLangID: DWORD;
    dwSecurity: DWORD;
    qos: TSecurityQOS;
  end;
  PConvContext = ^TConvContext;

  TConvInfo = record
    cb: DWORD;
    hUser: PtrUInt;
    hConvPartner: HCONV;
    hszSvcPartner: HSZ;
    hszServiceReq: HSZ;
    hszTopic: HSZ;
    hszItem: HSZ;
    wFmt: UINT;
    wType: UINT;
    wStatus: UINT;
    wConvst: UINT;
    wLastError: UINT;
    hConvList: HCONVLIST;
    ConvCtxt: TConvContext;
    hwnd: HWND;
    hwndPartner: HWND;
  end;

  THszPair = record
    hszSvc: HSZ;
    hszTopic: HSZ;
  end;

  TDdeCallback = function(wType, wFmt: UINT; hConv: HCONV; hsz1, hsz2: HSZ; hData: HDDEDATA;
    dwData1, dwData2: PtrUInt): HDDEDATA; stdcall;

const
  CP_WINANSI = 1004;
  CP_WINUNICODE = 1200;

  APPCLASS_STANDARD = $00000000;
  APPCMD_CLIENTONLY = $00000010;
  APPCMD_FILTERINITS = $00000020;
  CBF_FAIL_SELFCONNECTIONS = $00001000;
  CBF_SKIP_REGISTRATIONS = $00080000;
  CBF_SKIP_UNREGISTRATIONS = $00100000;

  XCLASS_BOOL = $1000;
  XCLASS_DATA = $2000;
  XCLASS_FLAGS = $4000;
  XCLASS_NOTIFICATION = $8000;
  XTYPF_NOBLOCK = $0002;
  XTYPF_NODATA = $0004;
  XTYPF_ACKREQ = $0008;

  XTYP_ERROR = $0000 or XCLASS_NOTIFICATION or XTYPF_NOBLOCK;
  XTYP_ADVDATA = $0010 or XCLASS_FLAGS;
  XTYP_ADVREQ = $0020 or XCLASS_DATA or XTYPF_NOBLOCK;
  XTYP_ADVSTART = $0030 or XCLASS_BOOL;
  XTYP_ADVSTOP = $0040 or XCLASS_NOTIFICATION;
  XTYP_EXECUTE = $0050 or XCLASS_FLAGS;
  XTYP_CONNECT = $0060 or XCLASS_BOOL or XTYPF_NOBLOCK;
  XTYP_CONNECT_CONFIRM = $0070 or XCLASS_NOTIFICATION or XTYPF_NOBLOCK;
  XTYP_XACT_COMPLETE = $0080 or XCLASS_NOTIFICATION;
  XTYP_POKE = $0090 or XCLASS_FLAGS;
  XTYP_REGISTER = $00A0 or XCLASS_NOTIFICATION or XTYPF_NOBLOCK;
  XTYP_REQUEST = $00B0 or XCLASS_DATA;
  XTYP_DISCONNECT = $00C0 or XCLASS_NOTIFICATION or XTYPF_NOBLOCK;
  XTYP_UNREGISTER = $00D0 or XCLASS_NOTIFICATION or XTYPF_NOBLOCK;
  XTYP_WILDCONNECT = $00E0 or XCLASS_DATA or XTYPF_NOBLOCK;
  XTYP_MASK = $00F0;

  DDE_FACK = $8000;
  DDE_FBUSY = $4000;
  DDE_FNOTPROCESSED = $0000;

  DNS_REGISTER = $0001;
  DNS_UNREGISTER = $0002;
  DNS_FILTERON = $0004;
  DNS_FILTEROFF = $0008;

  TIMEOUT_ASYNC = $FFFFFFFF;
  QID_SYNC = $FFFFFFFF;

  DMLERR_NO_ERROR = 0;
  DMLERR_FIRST = $4000;
  DMLERR_LAST = $4011;

  CF_TEXT = 1;
  CF_UNICODETEXT = 13;

  { topics and items every server should know (ddeml.h SZDDESYS_xxx) }
  SZDDESYS_TOPIC = 'System';
  SZDDESYS_ITEM_TOPICS = 'Topics';
  SZDDESYS_ITEM_SYSITEMS = 'SysItems';
  SZDDESYS_ITEM_RTNMSG = 'ReturnMessage';
  SZDDESYS_ITEM_STATUS = 'Status';
  SZDDESYS_ITEM_FORMATS = 'Formats';
  SZDDESYS_ITEM_HELP = 'Help';
  SZDDE_ITEM_ITEMLIST = 'TopicItemList';

function DdeInitializeW(var idInst: DWORD; pfnCallback: TDdeCallback; afCmd: DWORD;
  ulRes: DWORD): UINT; stdcall; external 'user32' name 'DdeInitializeW';
function DdeUninitialize(idInst: DWORD): BOOL; stdcall; external 'user32' name 'DdeUninitialize';
function DdeCreateStringHandleW(idInst: DWORD; psz: PWideChar; iCodePage: Integer): HSZ; stdcall;
  external 'user32' name 'DdeCreateStringHandleW';
function DdeFreeStringHandle(idInst: DWORD; hsz: HSZ): BOOL; stdcall;
  external 'user32' name 'DdeFreeStringHandle';
function DdeQueryStringW(idInst: DWORD; hsz: HSZ; psz: PWideChar; cchMax: DWORD;
  iCodePage: Integer): DWORD; stdcall; external 'user32' name 'DdeQueryStringW';
function DdeNameService(idInst: DWORD; hsz1, hsz2: HSZ; afCmd: UINT): HDDEDATA; stdcall;
  external 'user32' name 'DdeNameService';
function DdeCreateDataHandle(idInst: DWORD; pSrc: Pointer; cb, cbOff: DWORD; hszItem: HSZ;
  wFmt, afCmd: UINT): HDDEDATA; stdcall; external 'user32' name 'DdeCreateDataHandle';
function DdeGetData(hData: HDDEDATA; pDst: Pointer; cbMax, cbOff: DWORD): DWORD; stdcall;
  external 'user32' name 'DdeGetData';
function DdeFreeDataHandle(hData: HDDEDATA): BOOL; stdcall;
  external 'user32' name 'DdeFreeDataHandle';
function DdePostAdvise(idInst: DWORD; hszTopic, hszItem: HSZ): BOOL; stdcall;
  external 'user32' name 'DdePostAdvise';
function DdeConnect(idInst: DWORD; hszService, hszTopic: HSZ; pCC: PConvContext): HCONV; stdcall;
  external 'user32' name 'DdeConnect';
function DdeDisconnect(hConv: HCONV): BOOL; stdcall; external 'user32' name 'DdeDisconnect';
function DdeConnectList(idInst: DWORD; hszService, hszTopic: HSZ; hConvList: HCONVLIST;
  pCC: PConvContext): HCONVLIST; stdcall; external 'user32' name 'DdeConnectList';
function DdeQueryNextServer(hConvList: HCONVLIST; hConvPrev: HCONV): HCONV; stdcall;
  external 'user32' name 'DdeQueryNextServer';
function DdeDisconnectList(hConvList: HCONVLIST): BOOL; stdcall;
  external 'user32' name 'DdeDisconnectList';
function DdeQueryConvInfo(hConv: HCONV; idTransaction: DWORD; var CI: TConvInfo): UINT; stdcall;
  external 'user32' name 'DdeQueryConvInfo';
function DdeClientTransaction(pData: Pointer; cbData: DWORD; hConv: HCONV; hszItem: HSZ;
  wFmt, wType: UINT; dwTimeout: DWORD; pdwResult: PDWORD): HDDEDATA; stdcall;
  external 'user32' name 'DdeClientTransaction';
function DdeGetLastError(idInst: DWORD): UINT; stdcall; external 'user32' name 'DdeGetLastError';

{ helpers }
function HszOf(Inst: DWORD; const S: string): HSZ;
function HszText(Inst: DWORD; H: HSZ): string;
function DataText(H: HDDEDATA; Fmt: UINT): string;
function TextData(Inst: DWORD; const S: string; Item: HSZ; Fmt: UINT): HDDEDATA;
function TextBytes(const S: string; Fmt: UINT): RawByteString;
function DdeErrorText(Code: UINT): string;
function XtypName(wType: UINT): string;
function FormatName(Fmt: UINT): string;
function PartnerProcess(hConv: HCONV): string;

implementation

function QueryFullProcessImageNameW(hProcess: THandle; dwFlags: DWORD; lpExeName: PWideChar;
  var lpdwSize: DWORD): BOOL; stdcall; external 'kernel32' name 'QueryFullProcessImageNameW';

const
  PROCESS_QUERY_LIMITED_INFORMATION = $1000;

function HszOf(Inst: DWORD; const S: string): HSZ;
var
  W: UnicodeString;
begin
  W := UTF8Decode(S);
  Result := DdeCreateStringHandleW(Inst, PWideChar(W), CP_WINUNICODE);
end;

function HszText(Inst: DWORD; H: HSZ): string;
var
  Buf: array[0..259] of WideChar;
  N: DWORD;
begin
  if H = 0 then Exit('');
  N := DdeQueryStringW(Inst, H, @Buf[0], Length(Buf), CP_WINUNICODE);
  Result := UTF8Encode(Copy(UnicodeString(PWideChar(@Buf[0])), 1, N));
end;

{ the bytes of a CF_TEXT / CF_UNICODETEXT item, up to the first #0 }
function DataText(H: HDDEDATA; Fmt: UINT): string;
var
  N: DWORD;
  Raw: RawByteString;
  W: UnicodeString;
  L: Integer;
begin
  Result := '';
  if H = 0 then Exit;
  N := DdeGetData(H, nil, 0, 0);
  if N = 0 then Exit;
  if Fmt = CF_UNICODETEXT then
  begin
    SetLength(W, (N + 1) div 2);
    DdeGetData(H, PWideChar(W), N, 0);
    L := 0;
    while (L < Length(W)) and (W[L + 1] <> #0) do Inc(L);
    Result := UTF8Encode(Copy(W, 1, L));
  end
  else
  begin
    SetLength(Raw, N);
    DdeGetData(H, PAnsiChar(Raw), N, 0);
    L := 0;
    while (L < Length(Raw)) and (Raw[L + 1] <> #0) do Inc(L);
    SetLength(Raw, L);
    if Fmt = CF_TEXT then
    begin
      { ANSI code page of Windows -> UTF-8 }
      SetLength(W, L);
      if L > 0 then
        SetLength(W, MultiByteToWideChar(CP_ACP, 0, PAnsiChar(Raw), L, PWideChar(W), L));
      Result := UTF8Encode(W);
    end
    else
      Result := Raw;
  end;
end;

{ UTF-8 text -> bytes of the clipboard format with the closing #0 }
function TextBytes(const S: string; Fmt: UINT): RawByteString;
var
  W: UnicodeString;
  N: Integer;
begin
  W := UTF8Decode(S);
  if Fmt = CF_UNICODETEXT then
  begin
    SetLength(Result, (Length(W) + 1) * 2);
    if Length(W) > 0 then
      Move(W[1], Result[1], Length(W) * 2);
    Result[Length(Result) - 1] := #0;
    Result[Length(Result)] := #0;
  end
  else
  begin
    N := WideCharToMultiByte(CP_ACP, 0, PWideChar(W), Length(W), nil, 0, nil, nil);
    SetLength(Result, N + 1);
    if N > 0 then
      WideCharToMultiByte(CP_ACP, 0, PWideChar(W), Length(W), PAnsiChar(Result), N, nil, nil);
    Result[N + 1] := #0;
  end;
end;

function TextData(Inst: DWORD; const S: string; Item: HSZ; Fmt: UINT): HDDEDATA;
var
  B: RawByteString;
begin
  B := TextBytes(S, Fmt);
  Result := DdeCreateDataHandle(Inst, PAnsiChar(B), Length(B), 0, Item, Fmt, 0);
end;

function DdeErrorText(Code: UINT): string;
begin
  case Code of
    DMLERR_NO_ERROR: Result := 'no error';
    $4000: Result := 'advise transaction timed out';
    $4001: Result := 'server busy';
    $4002: Result := 'data transaction timed out';
    $4003: Result := 'DDEML not initialized';
    $4004: Result := 'DDEML usage error';
    $4005: Result := 'execute transaction timed out';
    $4006: Result := 'invalid parameter';
    $4007: Result := 'low memory';
    $4008: Result := 'memory error';
    $4009: Result := 'not processed by the server';
    $400A: Result := 'no conversation (server not running or topic unknown)';
    $400B: Result := 'poke transaction timed out';
    $400C: Result := 'PostMessage failed';
    $400D: Result := 'reentrancy (transaction while another one is running)';
    $400E: Result := 'server died';
    $400F: Result := 'system error';
    $4010: Result := 'unadvise transaction timed out';
    $4011: Result := 'unknown queue id';
  else
    Result := Format('DDEML error 0x%.4x', [Code]);
  end;
end;

function XtypName(wType: UINT): string;
begin
  case wType of
    XTYP_ERROR: Result := 'ERROR';
    XTYP_ADVDATA: Result := 'ADVDATA';
    XTYP_ADVREQ: Result := 'ADVREQ';
    XTYP_ADVSTART: Result := 'ADVSTART';
    XTYP_ADVSTOP: Result := 'ADVSTOP';
    XTYP_EXECUTE: Result := 'EXECUTE';
    XTYP_CONNECT: Result := 'CONNECT';
    XTYP_CONNECT_CONFIRM: Result := 'CONNECT_CONFIRM';
    XTYP_XACT_COMPLETE: Result := 'XACT_COMPLETE';
    XTYP_POKE: Result := 'POKE';
    XTYP_REGISTER: Result := 'REGISTER';
    XTYP_REQUEST: Result := 'REQUEST';
    XTYP_DISCONNECT: Result := 'DISCONNECT';
    XTYP_UNREGISTER: Result := 'UNREGISTER';
    XTYP_WILDCONNECT: Result := 'WILDCONNECT';
  else
    Result := Format('0x%.4x', [wType]);
  end;
end;

function FormatName(Fmt: UINT): string;
var
  Buf: array[0..127] of WideChar;
  N: Integer;
begin
  case Fmt of
    CF_TEXT: Result := 'TEXT';
    CF_UNICODETEXT: Result := 'UNICODETEXT';
  else
    N := GetClipboardFormatNameW(Fmt, @Buf[0], Length(Buf));
    if N > 0 then
      Result := UTF8Encode(Copy(UnicodeString(PWideChar(@Buf[0])), 1, N))
    else
      Result := IntToStr(Fmt);
  end;
end;

{ exe name of the program at the other end of a conversation }
function PartnerProcess(hConv: HCONV): string;
var
  CI: TConvInfo;
  Pid, N: DWORD;
  P: THandle;
  Buf: array[0..MAX_PATH] of WideChar;
begin
  Result := '?';
  FillChar(CI, SizeOf(CI), 0);
  CI.cb := SizeOf(CI);
  if DdeQueryConvInfo(hConv, QID_SYNC, CI) = 0 then Exit;
  Pid := 0;
  GetWindowThreadProcessId(CI.hwndPartner, @Pid);
  if Pid = 0 then Exit;
  Result := 'pid ' + IntToStr(Pid);
  P := OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, Pid);
  if P = 0 then Exit;
  try
    N := Length(Buf);
    if QueryFullProcessImageNameW(P, 0, @Buf[0], N) then
      Result := ExtractFileName(UTF8Encode(Copy(UnicodeString(PWideChar(@Buf[0])), 1, N))) +
        ' (' + IntToStr(Pid) + ')';
  finally
    CloseHandle(P);
  end;
end;

end.
