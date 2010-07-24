VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmUserMenu3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2460
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
   Icon            =   "FrmUserMenu3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2460
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   2385
      Begin AIFCmp1.asxToolbar Asx 
         Height          =   765
         Left            =   60
         Top             =   150
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   1349
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
         ButtonCount     =   4
         CaptionOptions  =   0
         AutoSize        =   -1  'True
         ButtonKey1      =   "bayar"
         ButtonPicture1  =   "FrmUserMenu3.frx":000C
         ButtonToolTipText1=   "Pay"
         ButtonKey2      =   "sambung"
         ButtonPicture2  =   "FrmUserMenu3.frx":0C5E
         ButtonToolTipText2=   "Continue to use this PC."
         ButtonStyle3    =   2
         ButtonKey4      =   "close"
         ButtonPicture4  =   "FrmUserMenu3.frx":18B0
         ButtonToolTipText4=   "Exit"
      End
   End
End
Attribute VB_Name = "FrmUserMenu3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonKey
    Case Is = "close"
        Unload Me
    Case Is = "bayar"
        Call UserHenti2
        Unload Me
    Case Is = "sambung"
        FrmGuna.Sambung = True
        Unload Me
        FrmGuna.Show
    End Select
End Sub
