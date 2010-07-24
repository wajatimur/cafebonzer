VERSION 5.00
Begin VB.Form FrmDbg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internal Process Viewer"
   ClientHeight    =   4665
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   9210
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDbg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Forms"
      Height          =   3975
      Left            =   6195
      TabIndex        =   5
      Top             =   30
      Width           =   2925
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bye"
      Height          =   480
      Left            =   8040
      TabIndex        =   4
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Caption         =   "Public Var"
      Height          =   3975
      Left            =   3135
      TabIndex        =   2
      Top             =   30
      Width           =   2925
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Socket"
      Height          =   3975
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2925
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   90
      Top             =   4095
   End
End
Attribute VB_Name = "FrmDbg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
On Error GoTo ErrInt
    List1.Clear
    For u = 1 To UniAgents.SockCount
        List1.AddItem "Socket > " & UniAgents.Socks(u).Handle & " (" & UniAgents.Socks(u).Index & ")"
    Next u
    
    List2.Clear
    List2.AddItem "cbUser > " & CbUserName
    List2.AddItem "cbAkses > " & CbUserAccess
    List2.AddItem "cbDemo > " & CbDemoMode
    List2.AddItem "cbDrvStr > " & CbDrvStr
    List2.AddItem "cbMsgRcv > " & CbMsgRcv
    List2.AddItem "cbConsole > " & CbConsole
    List2.AddItem "cbLogUser > " & CbLogUser
    List2.AddItem "lSock > " & lSock
    
    List3.Clear
    For Each Form In Forms
        List3.AddItem Form.Name
    Next
Exit Sub
ErrInt:
    Timer1.Enabled = False
    FrmMain.Caption = "Internal Process Viewer - Error Detected !"
    ErrLog Err, "Debug Windows - Timer1_Timer"
End Sub
