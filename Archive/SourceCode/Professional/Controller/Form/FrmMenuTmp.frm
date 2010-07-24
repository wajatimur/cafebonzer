VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmMenu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4770
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
   Icon            =   "FrmMenuTmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   4695
      Begin AIFCmp1.asxToolbar Asx 
         Height          =   765
         Left            =   60
         Top             =   150
         Width           =   4560
         _ExtentX        =   8043
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
         ButtonCount     =   9
         CaptionOptions  =   0
         AutoSize        =   -1  'True
         ButtonCaption1  =   "Stop"
         ButtonKey1      =   "stop"
         ButtonPicture1  =   "FrmMenuTmp.frx":000C
         ButtonToolTipText1=   "Stop Using PC"
         ButtonStyle2    =   2
         ButtonKey3      =   "msgbox"
         ButtonPicture3  =   "FrmMenuTmp.frx":0C5E
         ButtonToolTipText3=   "Send Message"
         ButtonKey4      =   "msgticker"
         ButtonPicture4  =   "FrmMenuTmp.frx":18B0
         ButtonToolTipText4=   "Send Ticker Message"
         ButtonStyle5    =   2
         ButtonKey6      =   "lock"
         ButtonPicture6  =   "FrmMenuTmp.frx":2502
         ButtonToolTipText6=   "Lock PC"
         ButtonKey7      =   "unlock"
         ButtonPicture7  =   "FrmMenuTmp.frx":3154
         ButtonToolTipText7=   "Unlock PC"
         ButtonKey8      =   "Open Padlock"
         ButtonStyle8    =   2
         ButtonToolTipText8=   "Open Padlock"
         ButtonKey9      =   "close"
         ButtonPicture9  =   "FrmMenuTmp.frx":3DA6
         ButtonToolTipText9=   "Close"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Dim SckId
    SckId = FrmMain.Lv1.SelectedItem.Tag
        
    Select Case ButtonKey
    Case Is = "close"
        FrmUserMenu.Hide
    Case Is = "stop"
        HentiGuna
        FrmUserMenu.Hide
    Case Is = "msgbox"
        FrmUserMenu.Hide
        FrmMesej.Show
    Case Is = "msgticker"
        FrmUserMenu.Hide
        FrmTiker.Show
    Case Is = "lock"
        FrmMain.socket(SckId).SendData "/kunci:1"
        FrmUserMenu.Hide
    Case Is = "unlock"
        FrmMain.socket(SckId).SendData "/kunci:0"
        FrmUserMenu.Hide
    End Select
End Sub
