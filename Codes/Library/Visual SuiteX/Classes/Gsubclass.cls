VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsGenGSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub AttachMessage(iwp As VsGenISubclass, ByVal hwnd As Long, ByVal iMsg As Long)
    VsGenSubclass.AttachMessage iwp, hwnd, iMsg
End Sub

Sub DetachMessage(iwp As VsGenISubclass, ByVal hwnd As Long, ByVal iMsg As Long)
    VsGenSubclass.DetachMessage iwp, hwnd, iMsg
End Sub

Public Property Get CurrentMessage() As Long
Attribute CurrentMessage.VB_Description = "Returns the message currently being worked on.  Only valid in the MsgResponse and WindowProc items of your implemented interface."
   CurrentMessage = VsGenSubclass.CurrentMessage
End Property
Public Function CallOldWindowProc( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long _
   ) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
   CallOldWindowProc = VsGenSubclass.CallOldWindowProc(hwnd, iMsg, wParam, lParam)
End Function

