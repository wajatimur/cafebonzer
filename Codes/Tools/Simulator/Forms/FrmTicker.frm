VERSION 5.00
Begin VB.Form FrmTicker 
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   2460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   Begin VB.PictureBox picTicker 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   45
      ScaleHeight     =   345
      ScaleWidth      =   915
      TabIndex        =   0
      ToolTipText     =   "CafeBonzer Agent R3"
      Top             =   90
      Width           =   915
   End
   Begin VB.Timer tmrTicker 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1515
      Top             =   45
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   700
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
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cool Ticker aka Form In The System Tray
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Option Explicit
Private TickText As String


Private Sub Form_Load()
    Dim Stle As Long
    
   ' Change window styles
    Call DrawBorder(Me.hwnd)
   ' Update ticker
    Call tmrTicker_Timer
    
   ' Enable the timer if ticker not hidden
    tmrTicker.Enabled = True
    tmrCheck.Enabled = True
End Sub


Private Sub Form_Resize()
  'Resize the PictureBox
  picTicker.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


Private Sub picTicker_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Me.PopupMenu FrmTray.pmenu1, , x, y
End Sub


Private Sub tmrCheck_Timer()
On Error GoTo ErrInt
    Dim W As Long, H As Long, Edge As SHAppBar_Edges

    'This timer is used to check the tray size and position to update the ticker position
    'Get the current task bar edge
    Edge = GetTaskBarEdge()
    Select Case Edge
    Case ABE_BOTTOM, ABE_TOP
        'If the position was changed resize ticker and show it
        If LastEdge <> Edge Then
            Call TickerResize
            Call TickerShow
        Else
            'If the task bar is over the same edge check if its size was changed
            GetTraySize W, H
            If TrayIconRows() = 1 Then
                If H < LastHeight Or W < LastWidth Then
                    TickerResize
                    TickerShow
                End If
            Else
                If H < LastHeight Or W < LastWidth Then
                    TickerHide False
                    TickerFly
                  End If
            End If
        End If
    Case ABE_LEFT, ABE_RIGHT
        If LastEdge <> Edge Then
            TickerHide False
            TickerFly
        End If
    End Select
    
    'Update last* variables
    LastEdge = Edge
    LastHeight = H
    LastWidth = W
Exit Sub

ErrInt:
    ErrHand Err, "tmrCheck_Timer"
End Sub


Private Sub tmrTicker_Timer()
On Error GoTo ErrInt
    Dim txtLen As Integer
    Static Pos As Long
    
        picTicker.ToolTipText = TickText
        txtLen = picTicker.TextWidth(TickText)
        
        If TickText = "" Then TickText = ":: CafeBonzer Agent R2 ::"
        
        'If Pos > Len(Text) Then Pos = 1
        '.CurrentY = (.ScaleHeight - picTicker.TextHeight(Text)) / 2
        '.CurrentX = (.ScaleWidth - picTicker.TextWidth(Text)) / 2
        'picTicker.Print Mid$(Text, Pos); Left$(Text, Pos)
        Pos = Pos + 15
        picTicker.Cls
        picTicker.CurrentY = 0
        picTicker.CurrentX = picTicker.Width - Pos
        picTicker.Print TickText
        If -Pos < -(txtLen + picTicker.Width) Then Pos = 0
Exit Sub

ErrInt:
    ErrHand Err, "tmrTicker_Timer"
End Sub


Public Sub SetText(Ayat As String)
    TickText = Ayat
End Sub
