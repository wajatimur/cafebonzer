unit MdlPipes;

interface 
uses
  Classes, Windows; 

const 
  cShutDownMsg = 'Shutdown Pipe';
  cPipeFormat = '\\%s\pipe\%s'; 

type 
  RPIPEMessage = record 
    Size: DWORD; 
    Kind: Byte; 
    Count: DWORD; 
    Data: array[0..8095] of Char; 
  end;
  TNotifyReceive = procedure(Sender: TObject; var ReceiveData: String) of object;

  TPipeServer = class(TThread)
  private
    FOnReceive: TNotifyReceive;
    FHandle: THandle;
    StrPipeServerName, StrPipeHostName, StrReceivedData: String;
  protected
    procedure DoChange; dynamic;
  public
    constructor CreatePipeServer(PipeServer: String; StartServer: Boolean);
    destructor Destroy; override;
    property OnReceive: TNotifyReceive read FOnReceive write FOnReceive;
    procedure StartUpServer;
    procedure ShutDownServer;
    procedure Execute; override;
  end;

  TPipeClient = class
  private 
    StrPipeServerName: String; 
    function ProcessMsg(RcvMessage: RPIPEMessage): RPIPEMessage;
  protected
    { Protected }
  public 
    constructor Create(PipeHost: String);
    function SendString(StringData: String): String;
  end; 



implementation 

uses 
  SysUtils; 

{ General }
procedure CalcMsgSize(var Msg: RPIPEMessage); 
begin 
  Msg.Size := SizeOf(Msg.Size) + SizeOf(Msg.Kind) + SizeOf(Msg.Count) + Msg.Count + 3;
end;


{ TPipeServer }
constructor TPipeServer.CreatePipeServer(PipeServer: String; StartServer: Boolean);
begin
  StrPipeServerName := Format(cPipeFormat, ['.', PipeServer]);
  FHandle := INVALID_HANDLE_VALUE;
  if StartServer then StartUpServer;
  Create(not StartServer);
end;

destructor TPipeServer.Destroy;
begin
  if FHandle <> INVALID_HANDLE_VALUE then ShutDownServer;
  inherited Destroy;
end;

procedure TPipeServer.StartUpServer;
begin
  { Check whether pipe does exist }
  if WaitNamedPipe(PChar(StrPipeServerName), 100 {ms}) then
    raise Exception.Create('Requested PIPE already exists.');
  { Create the pipe }
  FHandle := CreateNamedPipe(
    PChar(StrPipeServerName), PIPE_ACCESS_DUPLEX,
    PIPE_TYPE_MESSAGE or PIPE_READMODE_MESSAGE or PIPE_WAIT,
    PIPE_UNLIMITED_INSTANCES, SizeOf(RPIPEMessage), SizeOf(RPIPEMessage),
    NMPWAIT_USE_DEFAULT_WAIT, nil);
  { Check if pipe was created }
  if FHandle = INVALID_HANDLE_VALUE then
    raise Exception.Create('Could not create PIPE.');
end;

procedure TPipeServer.Execute;
var
  Written: Cardinal;
  InMsg, OutMsg: RPIPEMessage; 
begin 
  while not Terminated do 
  begin 
    if FHandle = INVALID_HANDLE_VALUE then 
      Sleep(250)
    else begin
      if ConnectNamedPipe(FHandle, nil) then 
      try 
        { Read data from pipe }
        InMsg.Size := SizeOf(InMsg); 
        ReadFile(FHandle, InMsg, InMsg.Size, InMsg.Size, nil);
        if (InMsg.Kind = 0) and (StrPas(InMsg.Data) = cShutDownMsg + StrPipeServerName) then 
        begin 
          { Process shut down }
          OutMsg.Kind := 0; 
          OutMsg.Count := 3; 
          OutMsg.Data := 'OK'#0; 
          Terminate; 
        end else begin 
          { Data send to pipe should be processed here }
          StrReceivedData := InMsg.Data;
          OutMsg := InMsg;
          InMsg.Data := '';
          DoChange;
        end;
        CalcMsgSize(OutMsg);
        WriteFile(FHandle, OutMsg, OutMsg.Size, Written, nil);
      finally
        DisconnectNamedPipe(FHandle);
      end; 
    end; 
  end; 
end; 

procedure TPipeServer.ShutDownServer;
var 
  BytesRead: Cardinal; 
  OutMsg, InMsg: RPIPEMessage; 
  ShutDownMsg: String; 
begin 
  if FHandle <> INVALID_HANDLE_VALUE then 
  begin 
    { Server still has pipe opened }
    OutMsg.Size := SizeOf(OutMsg); 
    with InMsg do
    begin 
      Kind := 0; 
      ShutDownMsg := cShutDownMsg + StrPipeServerName; 
      Count := Succ(Length(ShutDownMsg)); 
      StrPCopy(Data, ShutDownMsg); 
    end; 
    CalcMsgSize(InMsg); 
    { Send shut down message }
    CallNamedPipe(PChar(StrPipeServerName), @InMsg, InMsg.Size, @OutMsg, OutMsg.Size, BytesRead, 100);
    CloseHandle(FHandle);
    FHandle := INVALID_HANDLE_VALUE;
  end; 
end; 

procedure TPipeServer.DoChange;
begin
  if Assigned(FOnReceive) then FOnReceive(Self,StrReceivedData);
end;



{ TPipeClient } 
constructor TPipeClient.Create(PipeHost: String);
begin 
  inherited Create; 
  StrPipeServerName := Format(cPipeFormat, ['.', PipeHost])
end;

function TPipeClient.ProcessMsg(RcvMessage: RPIPEMessage): RPIPEMessage;
var
  LblReturn: LongBool;
begin
  CalcMsgSize(RcvMessage);
  Result.Size := SizeOf(Result);
  if WaitNamedPipe(PChar(StrPipeServerName), 10) then begin
    LblReturn := CallNamedPipe(PChar(StrPipeServerName), @RcvMessage, RcvMessage.Size, @Result, Result.Size, Result.Size, 500);
    if not LblReturn then raise Exception.Create('PIPE did not respond.');
  end else
    raise Exception.Create('PIPE does not exist.');
end;

function TPipeClient.SendString(StringData: String): String;
var 
  Msg: RPIPEMessage; 
begin 
  Msg.Kind := 1;
  Msg.Count := Length(StringData);
  StrPCopy(Msg.Data, StringData);
  { Send message }
  Msg := ProcessMsg(Msg); 
  { Return data send from server }
  Result := Copy(Msg.Data, 1, Msg.Count);
end; 

end. 

 