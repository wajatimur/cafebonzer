VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmUserMenu2 
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
   Icon            =   "FrmUserMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
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
         ButtonPicture1  =   "FrmUserMenu.frx":000C
         ButtonToolTipText1=   "Stop Using PC"
         ButtonStyle2    =   2
         ButtonKey3      =   "msgbox"
         ButtonPicture3  =   "FrmUserMenu.frx":0C5E
         ButtonToolTipText3=   "Send Message"
         ButtonKey4      =   "msgticker"
         ButtonPicture4  =   "FrmUserMenu.frx":18B0
         ButtonToolTipText4=   "Send Ticker Message"
         ButtonStyle5    =   2
         ButtonKey6      =   "lock"
         ButtonPicture6  =   "FrmUserMenu.frx":2502
         ButtonToolTipText6=   "Lock PC"
         ButtonKey7      =   "unlock"
         ButtonPicture7  =   "FrmUserMenu.frx":3154
         ButtonToolTipText7=   "Unlock PC"
         ButtonKey8      =   "Open Padlock"
         ButtonStyle8    =   2
         ButtonToolTipText8=   "Open Padlock"
         ButtonKey9      =   "close"
         ButtonPicture9  =   "FrmUserMenu.frx":3DA6
         ButtonToolTipText9=   "Close"
      End
   End
End
Attribute VB_Name = "FrmUserMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonKey
    Case Is = "close"
        Unload Me
    Case Is = "stop"
        Call UserHenti
        Unload Me
    Case Is = "msgbox"
        Unload Me
        FrmMesej.Show
    Case Is = "msgticker"
        Unload Me
        s_msg$ = MgoInpt.GetInput("Sila masukkan mesej tiker anda", BtnClose)
        If Trim(s_msg$) <> "" Then SelAgn.NetSend "tiker:" & msg
    Case Is = "lock"
        SelAgn.NetSend "//kunci:1"
        Unload Me
    Case Is = "unlock"
        If Mid(CbUserAccess, 3, 1) = 0 Then MsgBox MB(10), vbOKOnly, CbMsgWarn: Exit Sub
        SelAgn.NetSend "//kunci:0"
        Unload Me
    End Select
End Sub
