VERSION 5.00
Begin VB.UserControl PageHolder 
   Alignable       =   -1  'True
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1800
   ScaleWidth      =   6360
   ToolboxBitmap   =   "PageHolder.ctx":0000
   Begin CafeBonzer.Line3D HolderLne 
      Height          =   45
      Left            =   15
      TabIndex        =   2
      Top             =   285
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.PictureBox Holder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   15
      ScaleHeight     =   285
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      Begin VB.Label HolderLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PageHolder"
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
         Left            =   525
         TabIndex        =   1
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image HolderBtn 
         Height          =   240
         Left            =   255
         Picture         =   "PageHolder.ctx":0312
         Top             =   30
         Width           =   240
      End
      Begin VB.Image HolderIcn 
         Height          =   240
         Left            =   15
         Picture         =   "PageHolder.ctx":046E
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   5385
      Picture         =   "PageHolder.ctx":04E4
      Top             =   1425
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   3
      Left            =   5925
      Picture         =   "PageHolder.ctx":055A
      Top             =   1425
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   2
      Left            =   5655
      Picture         =   "PageHolder.ctx":0632
      Top             =   1425
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   5100
      Picture         =   "PageHolder.ctx":0709
      Top             =   1425
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "PageHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tImg As New StdPicture

Private lHldrStyle As Long
Private lHldrIcon As Long
Private sHldrTxt As String
Private oHldrTxtClr As OLE_COLOR
Private bHldrLne As Boolean

Private lPageState As Long
Private lPageHeight As Long


Public Enum eHldrStyle
    [Normal] = 0
    [Simple] = 1
    [Text Only] = 2
End Enum

Public Enum eHldrIcon
    [Default] = 0
    [Planetary 1] = 1
    [Planetary 2] = 2
End Enum

Public Event HolderButtonClick(ByVal Collapse As Boolean)
Public Event PageFlip(ByVal Collapse As Boolean)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' USERCONTROL
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_InitProperties()
    sHldrTxt = Extender.Name
    oHldrTxtClr = vbWhite
    bHldrLne = True
    lPageHeight = UserControl.Height
    Call zRedrawHolder(lHldrStyle)
End Sub
Private Sub UserControl_Resize()
    Holder.Width = UserControl.Width
    HolderLne.Width = Holder.Width
    If lPageState = 0 Then
        lPageHeight = UserControl.Height
    Else
        Extender.Height = Holder.Height
    End If
End Sub
Private Sub HolderBtn_Click()
    If lPageState = 0 Then
        lPageState = 1
        Call zCollapse
    Else
        lPageState = 0
        Call zCollapse(False)
    End If
    RaiseEvent HolderButtonClick(lPageState)
End Sub
Private Sub HolderLbl_Click()
    Call HolderBtn_Click
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' READ\WRITE PROPERTIES
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lHldrStyle = PropBag.ReadProperty("HldrStyle", 0)
    lHldrIcon = PropBag.ReadProperty("HldrIcon", 0)
    sHldrTxt = PropBag.ReadProperty("HldrTxt", Extender.Name)
    oHldrTxtClr = PropBag.ReadProperty("HldrTxtClr", vbWhite)
    bHldrLne = PropBag.ReadProperty("HldrLne", True)
    
    lPageState = PropBag.ReadProperty("PageState", 0)
    lPageHeight = PropBag.ReadProperty("PageHeight", lPageHeight)
    
    Call zRedrawHolder(lHldrStyle)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HldrStyle", lHldrStyle, 0
    PropBag.WriteProperty "HldrIcon", lHldrIcon, 0
    PropBag.WriteProperty "HldrTxt", sHldrTxt
    PropBag.WriteProperty "HldrTxtClr", oHldrTxtClr
    PropBag.WriteProperty "HldrLne", bHldrLne
    
    PropBag.WriteProperty "PageState", lPageState, 0
    PropBag.WriteProperty "PageHeight", lPageHeight
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PROPERTY SECTION
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get HolderStyle() As eHldrStyle
    HolderStyle = lHldrStyle
End Property
Public Property Let HolderStyle(nVal As eHldrStyle)
    lHldrStyle = nVal
    Call zRedrawHolder(nVal)
    Call PropertyChanged
End Property

Public Property Get HolderIcon() As eHldrIcon
    HolderIcon = lHldrIcon
End Property
Public Property Let HolderIcon(nVal As eHldrIcon)
    lHldrIcon = nVal
    Call zRedrawHolder(lHldrStyle)
    Call PropertyChanged
End Property

Public Property Get HolderText() As String
    HolderText = sHldrTxt
End Property
Public Property Let HolderText(nVal As String)
    sHldrTxt = nVal
    Call zRedrawHolder(0, True)
    Call PropertyChanged
End Property

Public Property Get HolderTextColor() As OLE_COLOR
    HolderTextColor = oHldrTxtClr
End Property
Public Property Let HolderTextColor(nVal As OLE_COLOR)
    oHldrTxtClr = nVal
    Call zRedrawHolder(0, True)
    Call PropertyChanged
End Property

Public Property Get HolderLine() As Boolean
    HolderLine = bHldrLne
End Property
Public Property Let HolderLine(nVal As Boolean)
    bHldrLne = nVal
    Call zRedrawHolder(lHldrStyle)
    Call PropertyChanged
End Property

Public Property Get PageCollapse() As Boolean
    If lPageState = 0 Then
        PageCollapse = False
    Else
        PageCollapse = True
    End If
End Property
Public Property Let PageCollapse(nVal As Boolean)
    If PageCollapse = True Then
        lPageState = 0
        Call zCollapse(False)
    Else
        lPageState = 1
        Call zCollapse
    End If
    Call PropertyChanged
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PRIVATE METHOD & FUNCTION
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub zRedrawHolder(Style, Optional TextOnly As Boolean = False)
    HolderLbl = sHldrTxt
    HolderLbl.ForeColor = oHldrTxtClr
    If TextOnly = True Then Exit Sub
    
    HolderIcn.Picture = ImgCnt(lHldrIcon + 1).Picture
    HolderLne.Visible = bHldrLne
    Select Case Style
        Case 0
            HolderIcn.Left = 15
            HolderBtn.Left = 255
            HolderLbl.Left = 525
            HolderIcn.Visible = True
            HolderBtn.Visible = True
            HolderLbl.Visible = True
        Case 1
            HolderIcn.Visible = False
            HolderBtn.Visible = True
            HolderLbl.Visible = True
            HolderBtn.Left = 15
            HolderLbl.Left = 255
        Case 2
            HolderIcn.Visible = False
            HolderBtn.Visible = False
            HolderLbl.Visible = True
            HolderLbl.Left = 15
    End Select
End Sub

Private Sub zCollapse(Optional Collapse As Boolean = True)
    If Collapse = True Then
        Extender.Top = (Extender.Top + Extender.Height) - Holder.Height
        Extender.Height = Holder.Height
        Set tImg = HolderBtn
        HolderBtn.Picture = ImgCnt(0).Picture
    Else
        Extender.Top = (Extender.Top + Holder.Height) - lPageHeight
        Extender.Height = lPageHeight
        HolderBtn.Picture = tImg
    End If
    RaiseEvent PageFlip(Collapse)
End Sub

