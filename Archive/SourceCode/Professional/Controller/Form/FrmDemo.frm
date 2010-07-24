VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDemo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2745
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4020
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   75
      TabIndex        =   0
      Top             =   -30
      Width           =   3900
      Begin MSComctlLib.ProgressBar pbDay 
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   1935
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Min             =   1
         Max             =   9
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   45
         Picture         =   "FrmDemo.frx":000C
         Top             =   195
         Width           =   3000
      End
      Begin VB.Label lbDayleft 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Days Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Welcome to Cafebonzer. This is an unregistered (trial) version of CafeBonzer. You only can use it for 9 days."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         TabIndex        =   1
         Top             =   735
         Width           =   3600
      End
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   0
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Buy CafeBonzer."
      Top             =   2265
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   767
      TX              =   "Buy"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDemo.frx":0F63
      PICN            =   "FrmDemo.frx":0F7F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   1
      Left            =   1860
      TabIndex        =   5
      ToolTipText     =   "Evaluate CafeBonzer."
      Top             =   2265
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   767
      TX              =   "Try"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDemo.frx":1519
      PICN            =   "FrmDemo.frx":1535
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   2
      Left            =   2745
      TabIndex        =   6
      ToolTipText     =   "Register or Obtain a Full Version."
      Top             =   2265
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      TX              =   "Register"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmDemo.frx":1ACF
      PICN            =   "FrmDemo.frx":1AEB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "FrmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim l_DayLeft As Long
    
    PutOnTop FrmDemo.hwnd
    If SetAmbil("demoday", 10) > 9 Then pbDay.Value = 9: Exit Sub
    
    l_DayLeft = 9 - SetAmbil("demoday", 10)
    lbDayleft.Caption = l_DayLeft & " Days Left"
    pbDay.Value = SetAmbil("demoday", 10)
End Sub

Private Sub RegisterIt()
    Dim sNamaDaftar As String, sNomborDaftar As String
    If CbDrvStr = "" Then CbDrvStr = "a:"

    If ValidateDisk(CbDrvStr) = True Then
        sNamaDaftar = GetName(CbDrvStr)
        sNomborDaftar = GetKey(CbDrvStr)
        If InitReg = True Then
            SetSimpan "namadaftar", sNamaDaftar
            SetSimpan "nombordaftar", sNomborDaftar
            'DemoMode = False
            CbDemoMode = False
            SetSimpan "demo", False
            MsgBox MB(6), vbOKOnly, CbMsgWarn
        End If
    End If
End Sub

Private Sub MainBtn_Click(Index As Integer)
    Select Case Index
     Case 0
       'open website (registration page)
        If Len(Dir(App.Path & "\buy.htm", vbNormal)) = 0 Then
            Call ShellExecute(Me.hwnd, "open", "http://www.nematix.net", vbNullString, vbNullString, SW_NORMAL)
        Else
            Call ShellExecute(Me.hwnd, "open", App.Path & "\buy.htm", vbNullString, vbNullString, SW_NORMAL)
        End If
        Unload FrmDemo
        If SetAmbil("demoday", 10) > 9 Then
            Keluar False
            End
        End If
        
     Case 1
        If SetAmbil("demoday", 10) > 9 Then
            MsgBox MB(3), vbOKOnly, "CafeBonzer"
            Exit Sub
        End If
        Unload FrmDemo
        
    Case 2
        Call RegisterIt
    End Select
End Sub
