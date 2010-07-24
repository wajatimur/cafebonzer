VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmMesej 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server - Message"
   ClientHeight    =   3540
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   4515
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
   Icon            =   "FrmMesej.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4515
   Begin VB.TextBox tosend 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Sila tekan enter selepas menulis mesej"
      Top             =   1920
      Width           =   4395
   End
   Begin AIFCmp1.asxToolbar asx 
      Height          =   435
      Left            =   4020
      Top             =   3045
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   767
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "Ok"
      ButtonPicture1  =   "FrmMesej.frx":000C
      ButtonToolTipText1=   "Ok"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your message below.."
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   1590
      Width           =   2955
   End
   Begin VB.Label server 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   585
      TabIndex        =   2
      Top             =   105
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   45
      Picture         =   "FrmMesej.frx":035E
      Top             =   15
      Width           =   480
   End
   Begin VB.Label rcv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   0
      Top             =   525
      Width           =   4395
   End
End
Attribute VB_Name = "FrmMesej"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    FrmMesej.Hide
    CbMsgRcv = False
End Sub

Private Sub tosend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SelAgn.NetSend "//mesej:Server:" & tosend.Text
        tosend.Text = ""
    ElseIf KeyCode = vbKeyEscape Then
        FrmMesej.Hide
        CbMsgRcv = False
    End If
End Sub
