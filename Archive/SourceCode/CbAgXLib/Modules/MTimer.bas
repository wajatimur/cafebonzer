Attribute VB_Name = "MTimer"
'*****************************************************************************************
'* Module      : MTimer
'* Description : Used by the CTimer class.
'* Notes       : Based on the implementations done by Bruce McKinney and Steve McMahon
'*****************************************************************************************

Option Explicit

' Private API function declarations
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

' Private constants
Private Const I_MAX_TIMERS = 100
Private Const eErrTimer_TooManyTimers = 18510 + 1  ' Too many timers
Private Const S_ERR_TooManyTimers = "Too many timers"

' Private variables for internal use
Public TimersArray(1 To I_MAX_TIMERS) As XLTimer
Private m_cTimerCount As Integer


'*****************************************************************************************
'* Function    : TimerCreate
'* Notes       : Creates the specified timer.
'*               Returns True if the timer object is created, False otherwise.
'*****************************************************************************************
Public Function TimerCreate(Timer As XLTimer) As Boolean
    On Error Resume Next
    
    Dim i As Integer
    
    Timer.TimerID = SetTimer(0&, 0&, Timer.Interval, AddressOf TimerProc)
    
    If Timer.TimerID Then
        
        TimerCreate = True
        
        For i = 1 To I_MAX_TIMERS
            
            If TimersArray(i) Is Nothing Then
                Set TimersArray(i) = Timer
                
                If (i > m_cTimerCount) Then
                    m_cTimerCount = i
                End If
                
                TimerCreate = True
                Exit Function
            End If
        
        Next
        
        On Error GoTo 0
        Err.Raise eErrTimer_TooManyTimers, "MTimer.TimerCreate", S_ERR_TooManyTimers
    
    Else
        ' TimerCreate = False
        Timer.TimerID = 0
        Timer.Interval = 0
    End If
End Function


'*****************************************************************************************
'* Function    : TimerDestroy
'* Notes       : Destroys the specified timer object.
'*               Returns True if the timer object is distroyed, False otherwise.
'*****************************************************************************************
Public Function TimerDestroy(Timer As XLTimer) As Boolean
    On Error Resume Next
    
    Dim i As Integer, f As Boolean
    
    For i = 1 To m_cTimerCount
        
        ' Find timer in array
        If Not TimersArray(i) Is Nothing Then
            
            If Timer.TimerID = TimersArray(i).TimerID Then
                f = KillTimer(0, Timer.TimerID)
                ' Remove timer and set reference to nothing
                Set TimersArray(i) = Nothing
                TimerDestroy = True
                Exit Function
            End If
        
        End If
    
    Next
End Function


'*****************************************************************************************
'* Sub         : TimerProc
'* Notes       : The main procedure for a created timer.
'*****************************************************************************************
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 1 To m_cTimerCount
        
        If Not (TimersArray(i) Is Nothing) Then
            
            If idEvent = TimersArray(i).TimerID Then
                ' Generate the event
                TimersArray(i).PulseTimer
                Exit Sub
            End If
        
        End If
    
    Next
End Sub


'*****************************************************************************************
'* Function    : StoreTimer
'* Notes       : Stores the timer in the timers array (collection).
'*****************************************************************************************
Private Function StoreTimer(Timer As XLTimer)
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 1 To m_cTimerCount
        
        If TimersArray(i) Is Nothing Then
            Set TimersArray(i) = Timer
            StoreTimer = True
            Exit Function
        End If
    
    Next
End Function
