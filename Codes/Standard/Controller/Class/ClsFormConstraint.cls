VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFormConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : ClsFormConstraint
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Private Const WM_GETMINMAXINFO = &H24
Private Const WM_DESTROY = &H2

Private EnuMsgRsp As EMsgResponse
Implements ISubclass

Private LngHwnd As Long
Private BlnTrackMin As Boolean
Private BlnTrackMax As Boolean
Private LngMinWidth As Long
Private LngMinHeight As Long
Private LngMaxWidth As Long
Private LngMaxHeight As Long


Public Sub InitWindow(ByVal Hwnd As Long)
    LngHwnd = Hwnd
    AttachMessage Me, LngHwnd, WM_GETMINMAXINFO
End Sub

Public Sub InitEnd()
    DetachMessage Me, LngHwnd, WM_GETMINMAXINFO
End Sub

Private Sub Class_Terminate()
    InitEnd
End Sub


Public Sub TrackMin(Enabled As Boolean, Height As Long, Width As Long)
    BlnTrackMin = Enabled
    LngMinWidth = Width
    LngMinHeight = Height
End Sub

Public Sub TrackMax(Enabled As Boolean, Height As Long, Width As Long)
    BlnTrackMax = Enabled
    LngMaxWidth = Width
    LngMaxHeight = Height
End Sub


Private Property Let ISubclass_MsgResponse(ByVal RHS As VisualSuite.EMsgResponse)
    'EnuMsgRsp = RHS
End Property

Private Property Get ISubclass_MsgResponse() As VisualSuite.EMsgResponse
    EnuMsgRsp = emrConsume
    ISubclass_MsgResponse = EnuMsgRsp
End Property



Private Function ISubclass_WindowProc(ByVal Hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim UtpMinMax As MINMAXINFO
    
    Select Case iMsg
    Case WM_GETMINMAXINFO
        CopyMemory UtpMinMax, ByVal lParam, Len(UtpMinMax)
        If BlnTrackMin = True Then
            UtpMinMax.ptMinTrackSize.X = LngMinWidth
            UtpMinMax.ptMinTrackSize.Y = LngMinHeight
        Else
            UtpMinMax.ptMaxTrackSize.X = LngMaxWidth
            UtpMinMax.ptMaxTrackSize.Y = LngMaxHeight
        End If
        CopyMemory ByVal lParam, UtpMinMax, Len(UtpMinMax)
    Case WM_DESTROY
        InitEnd
    End Select
End Function
