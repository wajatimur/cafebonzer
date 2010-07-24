Attribute VB_Name = "mdlGui"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Enable Group] - Enable\Disable Group Base On GrpName Tag
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub EnableGroup(GrpName As String, Optional Enable As Boolean = False)
On Error Resume Next
    For Each Control In FrmMain.Controls
        If Control.Tag = GrpName Then Control.Enabled = Enable
    Next Control
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Get Input] - Get Input From User
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Function GetInput(Caption As String, Optional Textstr As String = "") As String
    FrmInput.Frame1.Caption = Caption
    FrmInput.Text1.Text = Textstr
    FrmInput.Show vbModal
    GetInput = FrmInput.Text1
    Set FrmInput = Nothing
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Put Object Center] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub PutObjCenter(ObjName As Object, WhichForm As Form)
    fW = WhichForm.Width
    fH = WhichForm.Height
    oW = ObjName.Width
    oH = ObjName.Height
    
    oBJNewX = (fW / 2) - (oW / 2)
    oBJNewY = (fH / 2) - (oH / 2)
    
    ObjName.Move oBJNewX, oBJNewY
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [Enum Font Procedure] - Enum All Font Callback Procedure
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function EnumFontProc(ByVal lplf As Long, ByVal lptm As Long, ByVal dwType As Long, ByVal lpData As Long) As Long
    Dim lfRet As LOGFONT, s_FntName As String
    
    CopyMemory lfRet, ByVal lplf, LenB(lfRet)
    s_FntName = StrConv(lfRet.lfFaceName, vbUnicode)
    s_FntName = Trim(s_FntName)
    
    FrmPickFont.FntList.AddItem s_FntName
    EnumFontProc = 1
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [DrawBorder] -
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub DrawBorder(hwnd As Long)
    Dim Stle As Long
  
  '{ Buang 'Border' asal }'
    Stle = GetWindowLong(hwnd, GWL_STYLE)
    Stle = Stle And Not WS_BORDER
    SetWindowLong hwnd, GWL_STYLE, Stle
    
  '{ Set 'Style' baru }'
    Stle = GetWindowLong(hwnd, GWL_EXSTYLE)
    Stle = Stle Or WS_EX_STATICEDGE
    SetWindowLong hwnd, GWL_EXSTYLE, Stle

End Sub
