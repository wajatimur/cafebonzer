VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B280D12A-792E-4DF1-AA2A-E84D836A12CC}#3.0#0"; "VISUAL~1.OCX"
Begin VB.Form FrmMessaging 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CafeBonzer Message Center"
   ClientHeight    =   5775
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   6615
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmMesej.frx":000C
   ScaleHeight     =   5775
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList Img16 
      Left            =   5955
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":1FD9
            Key             =   "MSGOUT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":2573
            Key             =   "MSGIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":2B0D
            Key             =   "TERMUSER"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":2EA7
            Key             =   "TERM"
         EndProperty
      EndProperty
   End
   Begin VisualSuiteX.VsGuiLine VsGuiLine1 
      Height          =   45
      Left            =   15
      TabIndex        =   2
      Top             =   720
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.TextBox MsgSend 
      ForeColor       =   &H00404040&
      Height          =   570
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4890
      Width           =   6540
   End
   Begin MSComctlLib.ListView LsvMessage 
      Height          =   4020
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   7091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Img16"
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Message"
         Object.Width           =   11465
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   5475
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11615
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmMessaging
'    Project    : CafeBonzerAG
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub Form_Load()
    PutOnTop Me.hWnd
End Sub

Private Sub MsgSend_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim StrMessage  As String
    
    If KeyCode = vbKeyReturn Then
        NetSend "030010" & MsgSend.Text
        StrMessage = "[" & SysInfoGetName & "] " & MsgSend.Text
        LsvMessage.ListItems.Add , , StrMessage, , "MSGOUT"
        MsgSend.Text = ""
    End If
    If KeyCode = vbKeyEscape Then
        FrmMessaging.Hide
    End If
End Sub
