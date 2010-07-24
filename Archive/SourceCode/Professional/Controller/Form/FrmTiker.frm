VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmTiker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mesej Tiker"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AIFCmp1.asxToolbar asxToolbar1 
      Height          =   390
      Left            =   3720
      Top             =   495
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   688
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
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
      ButtonCount     =   2
      CaptionOptions  =   0
      ButtonKey1      =   "Batal"
      ButtonPicture1  =   "FrmTiker.frx":0000
      ButtonToolTipText1=   "Batal"
      ButtonKey2      =   "Hantar"
      ButtonPicture2  =   "FrmTiker.frx":0352
      ButtonToolTipText2=   "Ok"
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "FrmTiker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
    Case 1
        FrmTiker.Hide
        FrmMain.Enabled = True
    Case 2
        If Text1.Text = "" Then Exit Sub
        Send SelTag, "//tiker:" & Text1.Text
        Text1.Text = ""
    End Select
End Sub
