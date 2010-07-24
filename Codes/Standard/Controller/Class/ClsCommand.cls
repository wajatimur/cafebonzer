VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : clsCommand
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'===============================================================
' COMMUNICATION COMMANDS
'   FORMAT :
'
'       [MainCommand][SubCommand][VALUE][VALUE2]
'       [00][0000][VALUE][VALUE2]
'
'       Note :
'       Value more than 2, must be generate by command "CmdSubPut",
'       for single value this command is optional
'
'   EXAMPLE :
'
'       Main Command = 01
'       Sub Command = 0022
'       Value = RM 3.00
'
'       010022RM 3.00
'
'
'   LIST :
'
'       [ Agent Command ]
'       NETWORK = 01
'           0010 = PING
'           0020 = PONG
'           0030 = CERT
'       INFO = 02
'           0010 = Resource
'           0020 = App\Process
'           0030 = Printer\Printing
'           0040 = Net. Traffic
'           0050 = Hardware Enum
'           0060 = Drive Information
'
'       [ Controller Command ]
'       NETWORK = 01
'           0010 = PING
'           0020 = CERT
'       CONTROL = 02
'           0010 = Shutdown\Restart(Exit Windows Code)
'           0020 = Lock\Unlock
'           0030 = Cleaning
'           0040 = Login\Logout
'           0050 = Usage Status\Control
'       MESSAGGING = 03
'           0010 = Direct Message
'           0020 = Direct Ticker
'       AGENT = 04
'
'===============================================================
Private ObjTmpDs As New ClsDataStore
Private ObjRoot As ClsAgent

Public Enum EnNetworkPing
    Ping = 0
    Pong = 1
End Enum

Public Enum EnNetworkCert
    Denied = 0
    Granted = 1
End Enum

Public Enum EnControlExitWin
    Shutdown = 0
    ShutdownForce = 1
    Reboot = 2
    RebootForce = 3
End Enum

Public Enum EnControlLock
    TerminalUnlock = 0
    TerminalLock = 1
End Enum

Public Enum EnControlCleaning
    Clean = 0
    CleanTemp = 1
    CleanRecycleBin = 2
    CleanHistory = 3
    CleanRecentDocs = 4
End Enum

Public Enum EnControlLogIn
    LogOut = 0
    LogIn = 1
End Enum
    
Public Enum EnControlUsage
    PriceUse = 0
    TimeLeft = 1
End Enum


Public Sub ObjectInit(ObjectAgent As ClsAgent)
    Set ObjRoot = ObjectAgent
End Sub

Public Sub ObjectCommands(StrCommand As String)
Attribute ObjectCommands.VB_UserMemId = 0
    Dim ObjTmpDs As New ClsDataStore
    
    If StrCommand = "" Then Exit Sub
    
    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCommand, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Network | Ping
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub NetPing(Optional Action As EnNetworkPing, Optional PingStamp As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore
    
    If Action = Ping Then
        StrCmd = "010010" & PingStamp
    Else
        StrCmd = "010020" & PingStamp
    End If
    
    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Network | Certificate
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub NetCert(Action As EnNetworkCert, Optional Reason As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "010030" & CmdSubPut("ACTION", Action) & CmdSubPut("REASON", Reason)

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Control | Exit Windows
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub ConExitWin(Action As EnControlExitWin, Optional Reason As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "020010" & CmdSubPut("ACTION", Action) & CmdSubPut("REASON", Reason)
    
    If Action = Shutdown Then SecUserLog SL(9)
    If Action = Reboot Then SecUserLog SL(8)
    
    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Control | Locking
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub ConLock(Action As EnControlLock)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "020020" & CmdSubPut("ACTION", Action)
    
    If Action = TerminalUnlock Then SecUserLog SL(4)
    
    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Control | Cleaning
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub ConCleaning(Action As EnControlCleaning)
     Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "020030" & Action

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Control | Login
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub ConLogin(Action As EnControlLogIn, Optional Extended As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "020040" & CmdSubPut("ACTION", Action) & CmdSubPut("EXTENDED", Extended)

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Control | Usage
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub ConUsage(Action As EnControlUsage, Amount As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    If Action = PriceUse Then
        StrCmd = "020050" & CmdSubPut("ACTION", PriceUse) & CmdSubPut("PRICEUSE", Amount)
    Else
        StrCmd = "020050" & CmdSubPut("ACTION", TimeLeft) & CmdSubPut("TIMELEFT", Amount)
    End If

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Mesagging | Direct Message
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub MsgDirectMsg(Message As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "030010" & Message

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Mesagging | Direct Ticker
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub MsgDirectTick(Message As String)
    Dim StrCmd As String, ObjTmpDs As New ClsDataStore

    StrCmd = "030020" & Message

    ObjTmpDs.Name = ObjRoot.AgentSockIndex
    ObjTmpDs.Add ObjRoot.AgentSockIndex, "sockindex"
    ObjTmpDs.Add StrCmd, "data"
    StackNetData.Add ObjTmpDs
    Set ObjTmpDs = Nothing
End Sub

Public Sub SysGetConfiguration(CommandData As String)
    
End Sub
