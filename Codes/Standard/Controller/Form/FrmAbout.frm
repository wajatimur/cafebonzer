VERSION 5.00
Begin VB.Form FrmAppAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5415
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
   Picture         =   "FrmAbout.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblEval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Evaluation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   75
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label LblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "azri@nematix.net"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   2715
      Width           =   1485
   End
   Begin VB.Label LblOrganisation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nematix Technology"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   2490
      Width           =   1740
   End
End
Attribute VB_Name = "FrmAppAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAppAbout
'    Project    : CafeBonzer
'
'    Description: Application Infomation
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Dim xHalf As Long, yHalf As Long
Dim Quat1 As Long, Quat2 As Long

Private idx As Long
Private idxx As Long
Private idxName As Long
    
Private Sub Form_Click()
    Set FrmAppAbout = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    'xHalf = sc.Width / 2
    'yHalf = sc.Height / 2
    'Quat1 = xHalf / 2
    'Quat2 = Quat1 + xHalf
    
    'Timer1.Enabled = True
    LblOrganisation = SetGetDb("GenOrgName")
    LblEmail = SetGetDb("GenOrgEmail")
    If SettingGet("RegName") = "Demo" Then LblEval.Visible = True
End Sub

'Private Sub sc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    yHalf = Y
'End Sub

Private Sub Redraw()
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
    nameStr2(6) = LblEmail
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
