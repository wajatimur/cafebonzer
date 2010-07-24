Attribute VB_Name = "mListview"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Listview
' Description         :
'==================================================================


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Cek jika nama telah digunakan dalam lv1
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function CekDuplicate(Nama) As Boolean
    Dim Ltm As ListItems
    Set Ltm = FrmMain.Lv1.ListItems
    CekDuplicate = False
    
    If AgentCount = 0 Then CekDuplicate = False: Exit Function
    For j = 1 To AgentCount
        If Ltm.Item(j).Text = Nama Then CekDuplicate = True: Exit Function
    Next j
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan subitem semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelSubItm(Index As Integer) As String
    If AgentCount = 0 Then Exit Function
    SelSubItm = FrmMain.Lv1.SelectedItem.SubItems(Index)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan key semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelKey() As String
    If AgentCount = 0 Then Exit Function
    SelKey = FrmMain.Lv1.SelectedItem.Key
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan index semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelIndex() As Integer
    If AgentCount = 0 Then Exit Function
    SelIndex = FrmMain.Lv1.SelectedItem.Index
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan Text semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelText() As String
    If AgentCount = 0 Then Exit Function
    SelText = FrmMain.Lv1.SelectedItem.Text
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan Tag semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelTag() As String
    If AgentCount = 0 Then Exit Function
    SelTag = FrmMain.Lv1.SelectedItem.Tag
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Tukarkan nama ke bentuk array
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Function CArrName(Nama As String) As String
    If Right(Nama, 1) <> ")" Then
        CArrName = Nama & "(1)"
        Exit Function
    Else
        j = InStrRev(Nama, "(", -1)
        If j <> 0 Then
            K = Mid(Nama, j + 1, Len(Nama) - j - 2)
            If IsNumeric(K) = True Then
                arrnum = GetArrNameNum(Nama) + 1
                CArrName = Left(Nama, j - 1) & "(" & arrnum & ")"
            Else
                CArrName = Nama & "(1)"
                Exit Function
            End If
        Else
            CArrName = Nama & "(1)"
            Exit Function
        End If
    End If
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Copy dari ListItems ke ListItems
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CopyItems(FromItems As ListItems, ToItems As ListItems, Optional CopyDataOnly As Boolean = True)
    If FromItems.Count = 0 Then Exit Sub
    
    For g = 1 To FromItems.Count
        ToItems(g).Text = FromItems(g).Text
        ToItems(g).Key = FromItems(g).Key
        ToItems(g).Tag = FromItems(g).Tag
        If CopyDataOnly = False Then
            ToItems(g).Bold = FromItems(g).Bold
            ToItems(g).Checked = FromItems(g).Checked
            ToItems(g).ForeColor = FromItems(g).ForeColor
            ToItems(g).Ghosted = FromItems(g).Ghosted
            ToItems(g).Icon = FromItems(g).Icon
            ToItems(g).SmallIcon = FromItems(g).SmallIcon
            ToItems(g).ToolTipText = FromItems(g).ToolTipText
        End If
        
        For H = 1 To FromItems(1).ListSubItems.Count
            ToItems(g).ListSubItems(H).Text = FromItems(g).Text
            ToItems(g).ListSubItems(H).Key = FromItems(g).Key
            ToItems(g).ListSubItems(H).Tag = FromItems(g).Tag
            If CopyDataOnly = False Then
                ToItems(g).ListSubItems(H).Bold = FromItems(g).Bold
                ToItems(g).ListSubItems(H).ForeColor = FromItems(g).ForeColor
                ToItems(g).ListSubItems(H).ToolTipText = FromItems(g).ToolTipText
            End If
        Next H
    Next g
End Sub
