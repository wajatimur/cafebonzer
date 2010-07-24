VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmUserMdump1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   1770
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
   Icon            =   "FrmUserMdump1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   1770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   1695
      Begin AIFCmp1.asxToolbar Asx 
         Height          =   765
         Left            =   60
         Top             =   150
         Width           =   1560
         _ExtentX        =   2752
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
         ButtonCount     =   3
         CaptionOptions  =   0
         AutoSize        =   -1  'True
         ButtonCaption1  =   "Stop"
         ButtonKey1      =   "stop"
         ButtonPicture1  =   "FrmUserMdump1.frx":000C
         ButtonToolTipText1=   "Stop Using PC"
         ButtonStyle2    =   2
         ButtonKey3      =   "close"
         ButtonPicture3  =   "FrmUserMdump1.frx":0C5E
         ButtonToolTipText3=   "Close"
      End
   End
End
Attribute VB_Name = "FrmUserMdump1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Asx_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonKey
    Case Is = "close"

    Case Is = "stop"

    Case Is = "msgbox"

    Case Is = "msgticker"

    Case Is = "lock"

    Case Is = "unlock"

    End Select
End Sub
