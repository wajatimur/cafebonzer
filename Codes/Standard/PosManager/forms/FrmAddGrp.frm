VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmAddGrp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   990
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmAddGrp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fme1 
      Caption         =   "Add New Group"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   3450
      Begin VB.TextBox Txt1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1380
         TabIndex        =   0
         Top             =   285
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   330
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   3630
      ScaleHeight     =   1125
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   0
      Width           =   360
      Begin AIFCmp1.asxToolbar Asx1 
         Height          =   390
         Left            =   0
         Top             =   615
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "ok"
         ButtonPicture1  =   "FrmAddGrp.frx":000C
         ButtonToolTipText1=   "Ok/Cancel"
      End
      Begin AIFCmp1.asxToolbar AsxC 
         Height          =   390
         Left            =   0
         Top             =   285
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   688
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   1
         ButtonKey1      =   "batal"
         ButtonPicture1  =   "FrmAddGrp.frx":035E
         ButtonToolTipText1=   "Cancel"
      End
   End
End
Attribute VB_Name = "FrmAddGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    If Txt1(0) = "" Then Exit Sub
    ret = posAddGroup(Txt1(0))
    Unload Me
End Sub

Private Sub AsxC_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Unload Me
End Sub
