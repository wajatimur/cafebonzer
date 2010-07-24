VERSION 5.00
Begin VB.UserControl TitleBar 
   Alignable       =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   5880
   Begin VisualSuite.uLine3D HolderLne 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   285
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.Image SysBtn 
      Height          =   240
      Index           =   2
      Left            =   5625
      Picture         =   "TitleBar.ctx":0000
      Top             =   15
      Width           =   240
   End
   Begin VB.Image SysBtn 
      Height          =   240
      Index           =   1
      Left            =   5400
      Picture         =   "TitleBar.ctx":0168
      Top             =   15
      Width           =   240
   End
   Begin VB.Image SysBtn 
      Height          =   240
      Index           =   0
      Left            =   5175
      Picture         =   "TitleBar.ctx":020C
      Top             =   15
      Width           =   240
   End
   Begin VB.Image HolderIcn 
      Height          =   240
      Left            =   15
      Picture         =   "TitleBar.ctx":02D2
      Top             =   15
      Width           =   240
   End
   Begin VB.Label HolderLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   30
      Width           =   735
   End
End
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private sHldrCap As String
Private oHldrCapClr As OLE_COLOR
Private bSysBtn(0 To 2) As Boolean


Private Sub UserControl_Terminate()
    sHldrCap = ""
    Erase bSysBtn
End Sub

Private Sub UserControl_InitProperties()
    sHldrCap = Extender.Name
    oHldrCapClr = vbWhite
    bSysBtn(0) = True
    bSysBtn(1) = True
    bSysBtn(2) = True
End Sub

Private Sub UserControl_Resize()
    UserControl.Parent.BorderStyle = 2
    UserControl.Parent.Caption = ""
    
    Extender.Align = 1
    Extender.Height = 330
    HolderLne.Width = Extender.Width
    Call zRedrawHolder
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    lret = SendMessage(Parent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    sHldrCap = PropBag.ReadProperty("HldrCap", Extender.Name)
    oHldrCapClr = PropBag.ReadProperty("HldrCapClr", oHldrCapClr)
    bSysBtn(0) = PropBag.ReadProperty("SysBtnMin", True)
    bSysBtn(1) = PropBag.ReadProperty("SysBtnMax", True)
    bSysBtn(2) = PropBag.ReadProperty("SysBtnClose", True)
    Call zRedrawHolder
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HldrCap", sHldrCap
    PropBag.WriteProperty "HldrCapClr", oHldrCapClr
    PropBag.WriteProperty "SysBtnMin", bSysBtn(0)
    PropBag.WriteProperty "SysBtnMax", bSysBtn(1)
    PropBag.WriteProperty "SysBtnClose", bSysBtn(2)
End Sub


Private Sub SysBtn_Click(Index As Integer)
    Select Case Index
        Case 0
            Parent.WindowState = 1
        Case 1
            If Parent.WindowState = 2 Then
                Parent.WindowState = 0
            Else
                Parent.WindowState = 2
            End If
        Case 2
            Unload Parent
    End Select
End Sub


Public Property Get Caption() As String
    Caption = sHldrCap
End Property
Public Property Let Caption(nVal As String)
    sHldrCap = nVal
    Call zRedrawHolder
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = oHldrCapClr
End Property
Public Property Let CaptionColor(nVal As OLE_COLOR)
    oHldrCapClr = nVal
    Call zRedrawHolder
End Property

Public Property Get MinButton() As Boolean
    MinButton = bSysBtn(0)
End Property
Public Property Let MinButton(nVal As Boolean)
    bSysBtn(0) = nVal
    Call zRedrawHolder
End Property

Public Property Get MaxButton() As Boolean
    MaxButton = bSysBtn(1)
End Property
Public Property Let MaxButton(nVal As Boolean)
    bSysBtn(1) = nVal
    Call zRedrawHolder
End Property

Public Property Get CloseButton() As Boolean
    CloseButton = bSysBtn(2)
End Property
Public Property Let CloseButton(nVal As Boolean)
    bSysBtn(2) = nVal
    Call zRedrawHolder
End Property


Public Sub zRedrawHolder()
    Dim l_Pos As Long
    
    HolderLbl = sHldrCap
    HolderLbl.ForeColor = oHldrCapClr
    
    For a% = UBound(bSysBtn) To 0 Step -1
        SysBtn(a%).Visible = False
        If bSysBtn(a%) = True Then
            l_Pos = l_Pos + 1
            SysBtn(a%).Visible = True
            SysBtn(a%).left = UserControl.Width - (SysBtn(a%).Width * l_Pos)
        End If
    Next a%
End Sub
