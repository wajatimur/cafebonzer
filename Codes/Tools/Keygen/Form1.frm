VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pengenerasi Nombor Daftar CC v 0.10"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1215
      Width           =   3030
   End
   Begin CbNG.Label3D Label3D2 
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   885
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   405
      Width           =   3015
   End
   Begin CbNG.Label3D Label3D1 
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DblClick()
    Form1.Hide
    Unload Form1
    End
End Sub

Private Sub Text1_Change()
    If Text1.Text = "" Then Text2.Text = "": Exit Sub
    Text2 = InitReg(Text1)
End Sub
