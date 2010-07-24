VERSION 5.00
Begin VB.Form FrmMesej 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agent - Message"
   ClientHeight    =   3495
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
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   420
      Left            =   3960
      Picture         =   "FrmMesej.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   495
   End
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
      Top             =   1890
      Width           =   4395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type below to send a message :"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   1575
      Width           =   2730
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
      Picture         =   "FrmMesej.frx":0596
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
Private Sub Command1_Click()
    FrmMesej.Hide
    bConMsg = False
End Sub

Private Sub Form_Load()
    PutOnTop Me.hwnd
End Sub

Private Sub tosend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then NetSend "/mesej:" & MyName & ":" & tosend.Text: tosend.Text = ""
    If KeyCode = vbKeyEscape Then FrmMesej.Hide: bConMsg = False
End Sub
