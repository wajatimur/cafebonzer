Attribute VB_Name = "MdlHelper"
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function RsFilter(RsTmp As Recordset, FilterStr As String)
    If FilterStr = "" Then Exit Function
    RsTmp.Filter = FilterStr
    Set RsFilter = RsTmp.OpenRecordset
End Function

Public Sub CtlDis(Mode As Long, Optional Flag As String)
    Dim strFlag As String
    '0 = normal
    '1 = can add group only
    '2 = can add group and item
    '3 = custom
    If Mode < 3 Then
        strFlag = Choose(Mode + 1, "1111", "1000", "1010")
    Else
        strFlag = Format(Flag, "0000")
    End If
    
    For l = 0 To 3
        FrmSnmMg.menu2itemfunc(l).Enabled = Mid(strFlag, l + 1, 1)
        FrmSnmMg.BtnMenu(l).Enabled = Mid(strFlag, l + 1, 1)
    Next l
End Sub

Public Sub MoveForm(hwnd As Long)
    ReleaseCapture
    lret = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub AutoColumn(LvName As ListView, Optional ColumnNum As Long = 1)
    Dim lngDiff As Long
    
    lngDiff = 100
    For l = 1 To LvName.ColumnHeaders.Count
        If l <> ColumnNum Then lngDiff = lngDiff + LvName.ColumnHeaders(l).Width
    Next l
    LvName.ColumnHeaders(ColumnNum).Width = LvName.Width - lngDiff
End Sub
