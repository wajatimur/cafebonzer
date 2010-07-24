VERSION 5.00
Begin VB.Form FrmInput 
   ClientHeight    =   945
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   5190
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
   Icon            =   "FrmInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   945
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin CafeBonzer.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   582
      HldrCap         =   "Input"
      HldrCapClr      =   16777215
      SysBtnMin       =   0   'False
      SysBtnMax       =   0   'False
      SysBtnClose     =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4620
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4845
      Picture         =   "FrmInput.frx":000C
      Top             =   510
      Width           =   240
   End
End
Attribute VB_Name = "FrmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Me.Hide
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then FrmInput.Hide
End Sub
