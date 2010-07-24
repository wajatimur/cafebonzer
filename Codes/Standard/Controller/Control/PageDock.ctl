VERSION 5.00
Begin VB.UserControl PageDock 
   Alignable       =   -1  'True
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
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
   ScaleHeight     =   4110
   ScaleWidth      =   1980
   ToolboxBitmap   =   "PageDock.ctx":0000
   Begin CafeBonzer.Line3D HolderLne 
      Height          =   4095
      Left            =   300
      TabIndex        =   1
      Top             =   0
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   7223
      horizon         =   0   'False
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
      Height          =   4110
      Left            =   15
      ScaleHeight     =   4110
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   0
      Width           =   285
      Begin VB.Image HolderBtn 
         Height          =   240
         Left            =   15
         Picture         =   "PageDock.ctx":0312
         Top             =   0
         Width           =   240
      End
      Begin VB.Image HolderIcn 
         Height          =   240
         Left            =   15
         Picture         =   "PageDock.ctx":046E
         Top             =   3855
         Width           =   240
      End
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   0
      Left            =   1410
      Picture         =   "PageDock.ctx":04E4
      Top             =   3825
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgCnt 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   1
      Left            =   1680
      Picture         =   "PageDock.ctx":0640
      Top             =   3825
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "PageDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : PageDock
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private lHldrBtnPos As Long
Private bHldrLne As Boolean

Private lPageWidth As Long
Private lPageState As Long

Public Enum eHldrBtnPos
    [Top] = 0
    [Bottom] = 1
    [Middle] = 2
End Enum

Public Event HolderButtonClick()
Public Event PageFliped(ByVal Flipped As Boolean)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' USERCONTROL
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_InitProperties()
    bHldrLne = True
    lPageState = 0
    lPageWidth = UserControl.Width
End Sub

Private Sub UserControl_Resize()
 '{ Resizer & Var }'
    Holder.Height = UserControl.Height
    HolderLne.Height = Holder.Height
    If lPageState = 0 Then
        lPageWidth = UserControl.Width
    Else
        Extender.Width = Holder.Width
    End If
    
 '{ Icon & Button Position }'
    If lHldrBtnPos = 0 Then
        HolderIcn.Top = Holder.Height - HolderIcn.Height
    ElseIf lHldrBtnPos = 1 Then
        HolderBtn.Top = Holder.Height - HolderBtn.Height
    End If
    
 '{ Smart Container Handler }'
    For Each Control In UserControl.ContainedControls
        If Control.Tag = "subcontainer" Then
            Control.Height = Extender.Height
        End If
    Next
End Sub

Private Sub HolderBtn_Click()
    RaiseEvent HolderButtonClick
    PageFlip = CBool(lPageState) Xor True
    Call zRedrawHolder
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' READ\WRITE PROPERTIES
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lHldrBtnPos = PropBag.ReadProperty("HldrBtnPos", 0)
    bHldrLne = PropBag.ReadProperty("HldrLne", True)
    lPageState = PropBag.ReadProperty("PageState", 0)
    lPageWidth = PropBag.ReadProperty("PageWidth", lPageWidth)
    Call zRedrawHolder
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HldrBtnPos", lHldrBtnPos
    PropBag.WriteProperty "HldrLne", bHldrLne
    PropBag.WriteProperty "PageState", lPageState
    PropBag.WriteProperty "PageWidth", lPageWidth
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PROPERTY SECTION
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get HolderButtonPos() As eHldrBtnPos
    HolderButtonPos = lHldrBtnPos
End Property
Public Property Let HolderButtonPos(nVal As eHldrBtnPos)
    lHldrBtnPos = nVal
    Call zRedrawHolder
    Call PropertyChanged
End Property

Public Property Get HolderLine() As Boolean
    HolderLine = bHldrLne
End Property
Public Property Let HolderLine(nVal As Boolean)
    bHldrLne = nVal
    Call zRedrawHolder
    Call PropertyChanged
End Property

Public Property Get PageFlip() As Boolean
    PageFlip = CBool(lPageState)
End Property
Public Property Let PageFlip(nVal As Boolean)
    lPageState = IIf(nVal, 1, 0)
    Call zFlip(nVal)
    Call zRedrawHolder
    Call PropertyChanged
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PRIVATE METHOD & FUNCTION
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub zRedrawHolder()
    HolderLne.Visible = bHldrLne
    Select Case lHldrBtnPos
        Case 0
            HolderBtn.Top = 0
            HolderBtn.Left = 15
            HolderIcn.Top = Holder.Height - HolderIcn.Height
            HolderIcn.Left = 15
        Case 1
            HolderBtn.Top = Holder.Height - HolderBtn.Height
            HolderBtn.Left = 15
            HolderIcn.Top = 0
            HolderIcn.Left = 15
        Case 2
            Mid% = (Holder.Height \ 2) - (HolderBtn.Height \ 2)
            HolderBtn.Top = Mid%
            HolderBtn.Left = 0
            HolderIcn.Top = -500
    End Select
    If lPageState = 1 Then HolderBtn.Picture = ImgCnt(1).Picture
End Sub

Private Sub zFlip(Optional Flip As Boolean = True)
    If Flip = True Then
        Extender.Left = (Extender.Left + Extender.Width) - Holder.Width
        Extender.Width = Holder.Width
        HolderBtn.Picture = ImgCnt(1).Picture
    Else
        Extender.Left = (Extender.Left + Extender.Width) - lPageWidth
        Extender.Width = lPageWidth
        HolderBtn.Picture = ImgCnt(0).Picture
    End If
    RaiseEvent PageFliped(Flip)
End Sub

