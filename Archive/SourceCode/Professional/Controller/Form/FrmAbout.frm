VERSION 5.00
Begin VB.Form FrmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5280
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   7
      Left            =   105
      Top             =   3645
   End
   Begin VB.Frame Frame2 
      Caption         =   "This Product Is Licensed To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1140
      Left            =   3435
      TabIndex        =   3
      Top             =   3750
      Width           =   3735
      Begin VB.Label lblnama 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   225
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   3420
      End
      Begin VB.Label lblkedai 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cc :"
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   525
         Width           =   3405
      End
      Begin VB.Label lblemail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email :"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   3420
      End
   End
   Begin CafeBonzer.Label3D Label3D2 
      Height          =   240
      Left            =   750
      TabIndex        =   2
      Top             =   3690
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   14737632
      ForeColor2      =   8388608
      Caption         =   "Copyright 2000-2002"
   End
   Begin VB.PictureBox sc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   45
      ScaleHeight     =   1875
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   30
      Width           =   5205
      Begin VB.Label lblCopy 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Nematix Technology© 1996-2002"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   15
         TabIndex        =   1
         Top             =   1680
         Width           =   3555
      End
      Begin VB.Image ImgClose 
         Height          =   480
         Left            =   4740
         Picture         =   "FrmAbout.frx":0000
         ToolTipText     =   "In Business Time Is Money"
         Top             =   1440
         Width           =   480
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   705
      Picture         =   "FrmAbout.frx":2072
      Top             =   3945
      Width           =   2295
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim xHalf As Long, yHalf As Long
    Dim Quat1 As Long, Quat2 As Long

Private idx As Long
Private idxx As Long
Private idxName As Long
    
Private Sub Form_Load()
    xHalf = sc.Width / 2
    yHalf = sc.Height / 2
    Quat1 = xHalf / 2
    Quat2 = Quat1 + xHalf
    
    Timer1.Enabled = True
    lblnama = SetAmbil("namadaftar")
    lblkedai = SetAmbil("namacc")
    lblemail = SetAmbil("emailpengguna")
    If SetAmbil("demo") = "True" Then FrmMain.Caption = FrmMain.Caption & " UNREGISTERED"
End Sub

Private Sub ImgClose_Click()
    Timer1.Enabled = False
    Set FrmAbout = Nothing
    Unload Me
End Sub

Private Sub sc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    yHalf = Y
End Sub

Private Sub Timer1_Timer()
    Dim curX As Long, curY As Long
    Dim curXX As Long, curYY As Long
    Dim nameStr(12) As String
    Dim nameStr2(12) As String
           
    nameStr(0) = "[ Azri Jamil ]"
    nameStr(1) = "wajatimur@bootbox.net"
    nameStr(2) = "Nematix Technology"
    nameStr(3) = "registered to"
    nameStr(4) = "> user"
    nameStr(5) = "> organisation"
    nameStr(6) = "> email"
    nameStr(7) = "thanks to"
    nameStr(8) = "> maui"
    nameStr(9) = "> bent/toilet"
    nameStr(10) = "> lemang/lembing"
    nameStr(11) = "> adie/comot"
    nameStr(12) = "[ end ]"
           
    nameStr2(0) = "programmer/author"
    nameStr2(1) = "email"
    nameStr2(2) = "copyright"
    nameStr2(3) = ""
    nameStr2(4) = lblnama
    nameStr2(5) = lblkedai
    nameStr2(6) = lblemail
    nameStr2(7) = ""
    nameStr2(8) = "bsd guru"
    nameStr2(9) = "script"
    nameStr2(10) = "tester"
    nameStr2(11) = "tester"
    nameStr2(12) = ""
    
    sc.Cls
    
 '>> Line 1 Counter
    If idxx >= xHalf And idxx <= xHalf + Quat1 Then
        'walking speed
        idxx = idxx + 30
    ElseIf idx >= xHalf + Quat1 Then
        'walkout speed
        idxx = idxx + 650
    Else
        'walkin speed
        idxx = idxx + 300
    End If
    
    curXX = sc.Width - idxx
    curYY = yHalf - 150
    
 'Line 2 Counter
    If idx >= xHalf And idx <= xHalf + Quat1 Then
        'walking speed
        idx = idx + 35
    ElseIf idx >= xHalf + Quat1 Then
        'walkout speed
        idx = idx + 650
    Else
        'walking speed
        idx = idx + 450
    End If
    
    curX = sc.Width - idx
    curY = yHalf - 150
    
 '>> Line 1
    sc.CurrentX = curXX + 100
    sc.CurrentY = curYY - 250
    sc.ForeColor = vbWhite
    sc.FontBold = False
    sc.Print nameStr2(idxName)
    
 '>> Line 2
    sc.CurrentX = curX - (Rnd * 90)
    sc.CurrentY = curY - (Rnd * 90)
    sc.ForeColor = &HC0C0C0
    sc.FontBold = False
    sc.Print nameStr(idxName)
    
    sc.CurrentX = curX + (Rnd * 90)
    sc.CurrentY = curY + (Rnd * 90)
    sc.ForeColor = &H808080
    sc.FontBold = False
    sc.Print nameStr(idxName)
    
    sc.CurrentX = curX
    sc.CurrentY = curY - (Rnd * 10)
    sc.ForeColor = vbWhite
    sc.FontBold = True
    sc.Print nameStr(idxName)
     
    'AntiAlias 1, 600, sc.Width, 1200, 4

    If idxx >= sc.Width + (sc.TextWidth(nameStr(idxName))) Then
        idx = 0
        idxx = 0
        idxName = idxName + 1
        If idxName = UBound(nameStr) + 1 Then idxName = 0
    End If
End Sub

Sub AntiAlias(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, step As Long)
For Y = Y1 To Y2 Step 100
    For X = X1 To X2 Step 100
        Avg = 0
        i = 0
        For yy = Y - step To Y + step
            For xx = X - step To X + step
                cp = sc.Point(xx, yy)
                Avg = Avg + (cp * cp)
                i = i + 1
            Next
        Next
        Avg = Sqr(Avg / i)
        sc.PSet (X, Y), Avg
    Next
Next

End Sub
