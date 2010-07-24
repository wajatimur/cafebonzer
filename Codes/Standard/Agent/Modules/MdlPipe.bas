Attribute VB_Name = "MdlPipe"



Public CUniPipe As New ClsPipe

Public Sub PipeTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    CUniPipe.ServerRead
End Sub
