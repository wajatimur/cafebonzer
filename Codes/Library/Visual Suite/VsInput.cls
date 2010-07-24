VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eButStyle
    [BtnMin] = 1
    [BtnMax] = 2
    [BtnClose] = 4
End Enum

Public Function GetInput(Title, Optional ButtonStyle As eButStyle = 1) As String
    FrmInput.TitleBar.Caption = Title
    Do Until ButtonStyle = 0
        DoEvents
        Select Case ButtonStyle
            Case Is >= 4
                ButtonStyle = ButtonStyle - 4
                FrmInput.TitleBar.CloseButton = True
            Case Is >= 2
                ButtonStyle = ButtonStyle - 2
                FrmInput.TitleBar.MaxButton = True
            Case Is >= 1
                ButtonStyle = ButtonStyle - 1
                FrmInput.TitleBar.MinButton = True
        End Select
    Loop
    
    FrmInput.Show vbModal
    GetInput = StrVsInputRet
End Function

Private Sub Class_Terminate()
    StrVsInputRet = ""
End Sub
