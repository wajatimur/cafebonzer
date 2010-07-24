VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAutoSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : ClsAutoSize
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private CControl As Control

Private LngTop As Long, LngBottom As Long
Private LngLeft As Long, LngRight As Long
Private LngTmpTop As Long, LngTmpBottom As Long
Private LngTmpLeft As Long, LngTmpRight As Long

Private BlnAnchorTop As Boolean, BlnAnchorBottom As Boolean
Private BlnAnchorLeft As Boolean, BlnAnchorRight As Boolean


Private Sub Class_Initialize()
    BlnAnchorTop = True
    BlnAnchorBottom = True
    BlnAnchorLeft = True
    BlnAnchorRight = True
End Sub

Private Sub Class_Terminate()
    Set CControl = Nothing
End Sub

Public Sub InitControl(ControlName As Control)
    Set CControl = ControlName
    
    LngTop = CControl.Top
    LngLeft = CControl.Left
    LngBottom = CControl.Container.Height - (CControl.Height + CControl.Top)
    LngRight = CControl.Container.Width - (CControl.Width + CControl.Left)
End Sub

Public Sub ResizeControl()
    On Error Resume Next
    If BlnAnchorBottom = True And BlnAnchorTop = True Then
        CControl.Height = CControl.Container.Height - (LngTop + LngBottom)
    End If
    If BlnAnchorBottom = True And BlnAnchorTop = False Then
        CControl.Top = CControl.Container.Height - (CControl.Height + LngBottom)
    End If
    
    If BlnAnchorRight = True And BlnAnchorLeft = True Then
        CControl.Width = CControl.Container.Width - (LngLeft + LngRight)
    End If
    If BlnAnchorRight = True And BlnAnchorLeft = False Then
        CControl.Left = CControl.Container.Width - (CControl.Width + LngRight)
    End If
End Sub

Public Sub NewSize()
    LngTmpTop = LngTop
    LngTmpLeft = LngLeft
    LngTmpBottom = LngBottom
    LngTmpRight = LngRight
    
    LngTop = CControl.Top
    LngLeft = CControl.Left
    LngBottom = CControl.Container.Height - (CControl.Height + CControl.Top)
    LngRight = CControl.Container.Width - (CControl.Width + CControl.Left)
End Sub

Public Sub NewSizeUndo()
    LngTop = LngTmpTop
    LngLeft = LngTmpLeft
    LngBottom = LngTmpBottom
    LngRight = LngRTmpight
End Sub


Public Property Get AnchorTop() As Boolean
    AnchorTop = BlnAnchorTop
End Property
Public Property Let AnchorTop(ByVal Value As Boolean)
    BlnAnchorTop = Value
End Property

Public Property Get AnchorBottom() As Boolean
    AnchorBottom = BlnAnchorBottom
End Property
Public Property Let AnchorBottom(ByVal Value As Boolean)
    BlnAnchorBottom = Value
End Property

Public Property Get AnchorLeft() As Boolean
    AnchorLeft = BlnAnchorLeft
End Property
Public Property Let AnchorLeft(ByVal Value As Boolean)
    BlnAnchorLeft = Value
End Property

Public Property Get AnchorRight() As Boolean
    AnchorRight = BlnAnchorRight
End Property
Public Property Let AnchorRight(ByVal Value As Boolean)
    BlnAnchorRight = Value
End Property
