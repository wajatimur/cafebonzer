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
   Begin CbAgSystem.CasGuiTitle TitleBar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   582
      HldrCap         =   "TitleBar"
      HldrCapClr      =   0
      SysBtnMin       =   -1  'True
      SysBtnMax       =   -1  'True
      SysBtnClose     =   -1  'True
   End
   Begin VB.TextBox InputTxt 
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
      Top             =   465
      Width           =   4620
   End
   Begin VB.Image Btn 
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


Private Sub Btn_Click()
    sVsInputRet = InputTxt
    Unload Me
End Sub

Private Sub InputTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sVsInputRet = InputTxt
        Unload Me
    End If
End Sub
