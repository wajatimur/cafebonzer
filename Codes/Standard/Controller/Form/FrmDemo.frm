VERSION 5.00
Begin VB.Form FrmSysDemo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmDemo.frx":000C
   ScaleHeight     =   4530
   ScaleWidth      =   7020
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptActivate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Activate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1845
      TabIndex        =   7
      Top             =   2355
      Width           =   2205
   End
   Begin VB.OptionButton OptEval 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evalution Version."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1845
      TabIndex        =   5
      Top             =   1470
      Value           =   -1  'True
      Width           =   2205
   End
   Begin CafeBonzer.Line3D Line3D1 
      Height          =   45
      Left            =   -15
      TabIndex        =   4
      Top             =   3900
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin CafeBonzer.XpButton MainBtn 
      Height          =   435
      Index           =   1
      Left            =   4470
      TabIndex        =   0
      ToolTipText     =   "Buy CafeBonzer."
      Top             =   4020
      Width           =   1230
      _ExtentX        =   2170
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
      MICON           =   "FrmDemo.frx":22E9
      PICN            =   "FrmDemo.frx":2305
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
      Index           =   0
      Left            =   5715
      TabIndex        =   1
      ToolTipText     =   "Register or Obtain a Full Version."
      Top             =   4020
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   767
      TX              =   "Continue"
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
      MICON           =   "FrmDemo.frx":289F
      PICN            =   "FrmDemo.frx":28BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   5475
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label LblActivate 
      BackStyle       =   0  'Transparent
      Caption         =   "Activate this version. Please provide the CafeBonzer Liscence Key to continue."
      Height          =   495
      Left            =   2130
      TabIndex        =   8
      Top             =   2700
      Width           =   4710
   End
   Begin VB.Label LblEval 
      BackStyle       =   0  'Transparent
      Caption         =   "Continue to evaluate this version of CafeBonzer. This evaluation version is full working without restriction."
      Height          =   495
      Left            =   2130
      TabIndex        =   6
      Top             =   1815
      Width           =   4680
   End
   Begin VB.Label LblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to CafeBonzer. This is an evaluation version of CafeBonzer. You only can use it for 14 days."
      Height          =   525
      Left            =   1605
      TabIndex        =   3
      Top             =   795
      Width           =   5040
   End
   Begin VB.Label lbDayleft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5895
      TabIndex        =   2
      Top             =   3585
      Width           =   1035
   End
End
Attribute VB_Name = "FrmSysDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmSysDemo
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private LngDayUse As Long

Private Sub Form_Load()
    FormOnTop FrmSysDemo.Hwnd
    LngDayUse = SetGetDb("AppDemoDay", CbDemoMaxDay)
    If LngDayUse >= CbDemoMaxDay Then
        OptEval.Enabled = False
        OptActivate.Value = True
    End If
    lbDayleft.Caption = CbDemoMaxDay - LngDayUse & " Days Left"
End Sub


Private Sub MainBtn_Click(Index As Integer)
    Select Case Index
        Case 0
        '{ Continue Evaluate\Activate }'
            If OptEval.Value = True Then
                If SetGetDb("AppDemoDay", CbDemoMaxDay) >= CbDemoMaxDay Then
                    MsgBox ST(3, 0), vbOKOnly, "CafeBonzer"
                    End
                End If
            Else
                Call SecLiscenseActivate
            End If
            Unload Me
        Case 1
        '{ Registration Page }'
           If Len(Dir(App.Path & "\buy.htm", vbNormal)) = 0 Then
               Call ShellExecute(Me.Hwnd, "open", "http://www.nematix.net", vbNullString, vbNullString, SW_NORMAL)
           Else
               Call ShellExecute(Me.Hwnd, "open", App.Path & "\buy.htm", vbNullString, vbNullString, SW_NORMAL)
           End If
    End Select
End Sub
