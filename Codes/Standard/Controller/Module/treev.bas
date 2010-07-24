Attribute VB_Name = "mTreev"
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Treeview
' Description         :
'==================================================================

'dapatkan key semasa
Function TvSelKey(TrView As TreeView) As String
    For g = 1 To TrView.Nodes.Count
    If TrView.Nodes.Item(g).Selected = True Then
    TvSelKey = TrView.Nodes.Item(g).Key
    End If
    Next g
End Function
'dapatkan index semasa
Function TvSelIndex(TrView As TreeView) As Integer
    For g = 1 To TrView.Nodes.Count
    If TrView.Nodes.Item(g).Selected = True Then
    TvSelIndex = TrView.Nodes.Item(g).Index
    End If
    Next g
End Function
'dapatkan text semasa
Function TvSelText(TrView As TreeView) As String
    For g = 1 To TrView.Nodes.Count
    If TrView.Nodes.Item(g).Selected = True Then
    TvSelText = TrView.Nodes.Item(g).Text
    End If
    Next g
End Function
