VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSysConsole 
   Caption         =   "Cafebonzer - Terminal"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTerm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Stat1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   3570
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   30
      ScaleHeight     =   3435
      ScaleWidth      =   5550
      TabIndex        =   2
      Top             =   45
      Width           =   5610
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   2760
         ItemData        =   "FrmTerm.frx":038A
         Left            =   -15
         List            =   "FrmTerm.frx":038C
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Width           =   5565
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   30
         TabIndex        =   0
         Top             =   3135
         Width           =   5505
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   -15
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   5565
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   -30
      Top             =   4515
   End
End
Attribute VB_Name = "FrmSysConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmSysConsole
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Dim Mm As String
Dim Asiap As Boolean
Public CurSocket As Long
Public Echo As Boolean


Sub wr(Ayat As String, Optional CuciDulu As Boolean = True)
    Mm = Ayat
    Timer1.Enabled = True
    Asiap = False
    If CuciDulu = True Then
    Text1.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Text1.BackColor = vbBlack
    Text2.BackColor = vbBlack
    List1.BackColor = vbBlack
    
    Echo = True
    Asiap = True
    CbConsole = True
    If SelText <> "" Then CurSocket = SelTag
    wr "Welcome To Console"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CbConsole = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Width = Me.Width - 170
    Picture1.Height = Me.Height - (480 + Stat1.Height)
    Text1.Width = Picture1.Width - 50
    Text2.Width = Picture1.Width - 50
    Text2.Top = Picture1.Height - (Text2.Height + 20)
    List1.Width = Picture1.Width - 50
    List1.Height = Picture1.Height - (Text1.Height + Text2.Height)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Asiap = False Then Exit Sub

    If KeyAscii = 13 And Text2.Text <> "" Then
        ProcessIn
    End If
    If KeyAscii = 27 Then FrmSysConsole.Hide: CbConsole = False: Set FrmSysConsole = Nothing
End Sub


Private Sub Timer1_Timer()
    Static idx As Integer
    pj = Len(Mm)
    If idx = pj Then
        Timer1.Enabled = False
        Asiap = True
        idx = 0
        Exit Sub
    End If
    idx = idx + 1
    rd = Mid(Mm, idx, 1)
    Text1.Text = Text1.Text & " " & rd
End Sub


Private Sub ProcessIn()
    If Text1 <> "" Then List1.AddItem Text1
    If Echo = True Then wr Text2.Text Else Text1 = ""
    Fetch Text2.Text
    Text2.Text = ""
End Sub


Public Sub Fetch(arahan As String)
    If arahan = "/debug" Then FrmSysDbg.Show: Exit Sub
    If arahan = "/flush" Then FrmSysConsole.List1.Clear: Exit Sub
    If Mid(arahan, 1, 5) = "/echo" Then DisEcho Mid(arahan, 7): Exit Sub
    If Mid(arahan, 1, 6) = "/mesej" Then SendMesej Mid(arahan, 8), CurSocket: Exit Sub
    If Mid(arahan, 1, 4) = "/cur" Then CurrentHook: Exit Sub
    If Mid(arahan, 1, 5) = "/hook" Then Hook Mid(arahan, 7): Exit Sub
    If CurSocket <> 0 Then Send CurSocket, arahan: Exit Sub
    wr "Command Not Found"
End Sub


Public Sub DisEcho(Param)
    If Param = "" Then Exit Sub
    If Left(Param, 1) = "1" Then
        Echo = True
        FrmSysConsole.wr "Echo enable"
    Else
        Echo = False
        FrmSysConsole.wr "Echo disable"
    End If
End Sub


Public Sub CurrentHook()
    Dim idx As Long
    idx = AgentGetIndex(FrmSysConsole.CurSocket)
    If idx = 0 Then
        FrmSysConsole.wr "No socket currently hook !"
    Else
        FrmSysConsole.wr "Socket currently hook to " & FrmMain.ListView.ListItems(idx).Text
    End If
End Sub

Public Sub Hook(StationName)
    Dim idx As Long
    idx = AgentGetIndexB(StationName)
    If idx = 0 Then
        FrmSysConsole.wr "Station not exist ! - " & StationName
    Else
        FrmSysConsole.CurSocket = FrmMain.ListView.ListItems(idx).Tag
        FrmSysConsole.wr "Hooking socket success > " & StationName
    End If
End Sub


Public Sub SendMesej(Param As String, Sck As Long)
    Dim Nama As String, Mesej As String
    If InStr(1, Param, ":") <> 0 Then
        Nama = Mid(Param, 1, InStr(1, Param, ":") - 1)
        Mesej = Mid(Param, InStr(1, Param, ":") + 1)
        Send Sck, "//mesej:" & Nama & ":" & Mesej
    Else
        Nama = "Server"
        Mesej = Param
        Send Sck, "//mesej:Server:" & Param
    End If
    FrmSysConsole.wr Nama & ">" & Mesej
End Sub

