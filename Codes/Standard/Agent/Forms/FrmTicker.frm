VERSION 5.00
Begin VB.Form FrmTicker 
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   2460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmrTicker 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1515
      Top             =   45
   End
   Begin VB.Timer TmrCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1980
      Top             =   45
   End
   Begin VB.Menu pmenu1 
      Caption         =   "<popmenu1>"
      Visible         =   0   'False
      Begin VB.Menu pmenu1_setting 
         Caption         =   "Penalaan"
      End
      Begin VB.Menu pmenu1_kunci 
         Caption         =   "Kunci PC"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu pmenu1_shutdown 
         Caption         =   "Shutdown PC"
      End
      Begin VB.Menu pmenu1_close 
         Caption         =   "Tutup Agent"
      End
   End
End
Attribute VB_Name = "FrmTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmTicker
'    Project    : CafeBonzerAG
'
'    Description: Cool Ticker aka Form In The System Tray
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Sub Form_Load()
    Call DrawBorder(Me.hWnd)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu FrmTray.PopMenu1, , x, y
End Sub

Private Sub TmrTicker_Timer()
    Call TickerDrawText
    Call TickerCheck
End Sub

