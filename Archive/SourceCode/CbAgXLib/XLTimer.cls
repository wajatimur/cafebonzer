VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************
'* Class       : XLTimer
'* Description : Class that implements a timer.
'* Notes       : Use this class to implement multiple timers in your applications without
'*               the need of any timer control.
'*               For advanced use, please see the Source+ Library online documentation
'*               in the Source+ Explorer (the "Documentation" tab).
'*****************************************************************************************

Option Explicit

' Error handling definitions
Private Const E_ERR_BASE = 18400 + vbObjectError
Public Enum EErrTimer
    eErrTimer_CantCreateTimer = E_ERR_BASE + 1
    eErrTimer_ComponentFailure
End Enum
Private Const S_ERR_ComponentFailure = "XLTimer component failure"
Private Const S_ERR_CantCreateTimer = "Can't create system timer"

' Private variables to hold property values
Private m_lInterval     As Long
Private m_lID           As Long
Private m_vTag          As Variant

' Event defintions
Public Event Timer()


'*****************************************************************************************
'* Property    : Interval
'* Notes       : Returns or sets the number of milliseconds between calls to a Timer event.
'*****************************************************************************************
Public Property Get Interval() As Long
    On Error GoTo hComponentFailure
    
    Interval = m_lInterval

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property

Public Property Let Interval(Value As Long)
    On Error GoTo hComponentFailure
    
    ' Can't just change interval--you must kill timer and start a new one
    If Value > 0 Then
        
        ' Don't mess with it if interval is the same
        If m_lInterval = Value Then Exit Property
        
        ' Must destroy any existing timer to change interval
        If m_lInterval Then
            
            If Not TimerDestroy(Me) Then
                On Error GoTo 0
                Err.Raise eErrTimer_ComponentFailure, "XLTimer.Interval", S_ERR_ComponentFailure
            End If
        
        End If
        
        ' Create new timer with new interval
        m_lInterval = Value
        
        If TimerCreate(Me) = False Then
            On Error GoTo 0
            Err.Raise eErrTimer_CantCreateTimer, "XLTimer.Interval", S_ERR_CantCreateTimer
        End If
    
    Else
        
        If (m_lInterval > 0) Then
            
            m_lInterval = 0
            
            If Not TimerDestroy(Me) Then
                On Error GoTo 0
                Err.Raise eErrTimer_ComponentFailure, "XLTimer.Interval", S_ERR_ComponentFailure
            End If
        
        End If
    
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : Tag
'* Notes       : Returns or sets the object associated extra data.
'*****************************************************************************************
Public Property Get Tag() As Variant
    On Error GoTo hComponentFailure
    
    If IsObject(m_vTag) Then
        Set Tag = m_vTag
    Else
        Tag = m_vTag
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property

Public Property Let Tag(Value As Variant)
    On Error GoTo hComponentFailure
    
    If IsObject(Value) Then
        Set m_vTag = Value
    Else
        m_vTag = Value
    End If

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property


'*****************************************************************************************
'* Property    : TimerID
'* Notes       : Returns or sets a long integer value containing the timer ID.
'*****************************************************************************************
Friend Property Get TimerID() As Long
    On Error GoTo hComponentFailure
    
    TimerID = m_lID

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property

Friend Property Let TimerID(Value As Long)
    On Error GoTo hComponentFailure
    
    m_lID = Value

    Exit Property

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Property

'*****************************************************************************************
'* Sub         : PulseTimer
'* Notes       : Raises a Timer event.
'*****************************************************************************************
Public Sub PulseTimer()
    On Error GoTo hComponentFailure
    
    RaiseEvent Timer

    Exit Sub

hComponentFailure:
    Err.Raise eErrTimer_ComponentFailure, App.EXEName & ".XLTimer", S_ERR_ComponentFailure
End Sub


'*****************************************************************************************
'* Sub         : Class_Terminate
'* Notes       : Perform object cleanup.
'*****************************************************************************************
Private Sub Class_Terminate()
    On Error Resume Next
    
    Interval = 0
End Sub
