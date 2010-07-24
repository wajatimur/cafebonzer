VERSION 5.00
Begin VB.Form FrmKunci 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   405
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6855
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
   Moveable        =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "FrmKunci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Secret As String

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then MsgBox Secret
    Secret = Secret & KeyCode
    Beep
End Sub

Private Sub Form_Load()
    PutOnTop Me.hwnd
End Sub

