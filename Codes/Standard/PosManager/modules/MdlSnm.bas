Attribute VB_Name = "MdlSnm"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GroupAdd] - Add new group
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub GroupAdd(StrGrpId As String, StrGrpName As String, StrGrpDesc As String, StrGrpSymbol As String)
On Error GoTo ErrInt
    Dim Rs As Recordset, CImgList As ImageList, CImgCmb As ImageCombo
    Dim IntIdxA As Integer
    
  ' tambah Group ke dalam db
    Set Rs = uSDB.OpenRecordset("ServiceCategory", dbOpenDynaset)
    With Rs
        .FindFirst "Name = '" & StrGrpName & "'"
        If .NoMatch = True Then
            .AddNew
            !Id = StrGrpId
            !Name = StrGrpName
            !Description = StrGrpDesc
            !Symbol = StrGrpSymbol
            .Update
        Else
            Exit Sub
        End If
    End With
    
  ' tambah ke dalam control IcGroup
    Set CImgList = FrmSnmMg.Iml
    Set CImgCmb = FrmSnmMg.IcGroup

    CImgCmb.ComboItems.Add , StrGrpId, StrGrpName, StrGrpSymbol
    If CImgCmb.ComboItems.Count = 1 Then
        CImgCmb.ComboItems(StrGrpId).Selected = True
        Call CtlDis(0)
    End If
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | GroupAdd"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GroupDel] - Delete group
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub GroupDel()
On Error GoTo ErrInt
    Dim Rs As Recordset
    Dim SelGrpId As String, StrRet As String
    
  ' jika jumlah group = 0, keluar
    If FrmSnmMg.IcGroup.ComboItems.Count = 0 Then Exit Sub
    
  ' tanya pengguna
    StrRet = MsgBox(Var(1), vbOKCancel, CbMsgWarn)
    If StrRet = vbOK Then
        Set Rs = uSDB.OpenRecordset("ServiceCategory", dbOpenTable)
      ' ambil group id
        SelGrpId = Left(FrmSnmMg.IcGroup.SelectedItem.Key, 5)
      ' seek dan delete group
        With Rs
            .Index = "Id"
            .Seek "=", SelGrpId
            If .NoMatch = True Or !Flag = 1 Then Exit Sub
            .Delete
        End With
        
      ' filter dan delete items bagi group tersebut
        Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)
        Set Rs = RsFilter(Rs, "GroupId = '" & SelGrpId & "'")
      ' jika ada item, padam kesemuanya
        If Rs.BOF = False Then
            With Rs
                .MoveFirst
                Do Until .EOF = True
                    .Delete
                    .MoveNext
                Loop
            End With
        End If
        
      ' padam group dari image combo dan bersihkan items dari Listview
        FrmSnmMg.IcGroup.ComboItems.Remove SelGrpId
        FrmSnmMg.LvItem.ListItems.Clear
        
      ' jika wujud group, pilih yang paling atas
      ' dan loadkan items-nya
        If FrmSnmMg.IcGroup.ComboItems.Count > 0 Then
            FrmSnmMg.IcGroup.ComboItems(1).Selected = True
            Call LoadItems
        Else
            FrmSnmMg.IcGroup.Text = ""
            FrmSnmMg.Refresh
        End If
    End If
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | GroupAdd"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [ItemAdd] - Add new item
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ItemAdd(GroupId As String, ItemName As String, ItemPrice As String, Optional ItemStock As Integer = 0)
On Error GoTo ErrInt
    Dim Rs As Recordset, StrItemID As String, TmpItm As ListItem
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)
    
    If IsNumeric(ItemPrice) Then Exit Sub
    
  ' retrive good item id
    StrItemID = ItemIdCheck(GroupId)
    With Rs
      ' search for duplicate items
        .FindFirst "Name = '" & ItemName & "'"
        If .NoMatch = True Then
          ' add item to database
            .AddNew
            !GroupId = GroupId
            !Id = StrItemID
            !Name = Trim(ItemName)
            !Price = Format(ItemPrice, "#0.00")
            !Stock = ItemStock
            !Symbol = "ITEM"
            .Update
        Else
          ' exit from this sub
            Exit Sub
        End If
    End With
    
  ' add current new item to listview
    If FrmSnmMg.IcGroup.SelectedItem.Key = GroupId Then
        Set TmpItm = FrmSnmMg.LvItem.ListItems.Add(, StrItemID, StrItemID, , "ITEM")
        TmpItm.SubItems(1) = Trim(ItemName)
        TmpItm.SubItems(2) = Format(ItemPrice, "#0.00")
    End If
Exit Sub

ErrInt:
    ErrLog Err, "SnmMg | ItemAdd"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [ItemSave] - Save new item
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ItemSave(ItemID As String, ItemName As String, ItemPrice As String, ItemSymbol As String)
On Error GoTo ErrInt
    Dim Rs As Recordset, StrItemID As String, TmpItm As ListItem
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)
    
    If IsNumeric(ItemPrice) = False Then Exit Sub
    
    With Rs
        .FindFirst "Id = '" & ItemID & "'"
        If .NoMatch = False Then
          ' add item to database
            .Edit
            !Name = Trim(ItemName)
            !Price = Format(ItemPrice, "#0.00")
            !Symbol = ItemSymbol
            .Update
        Else
          ' exit from this sub
            Exit Sub
        End If
    End With
    
    Set TmpItm = FrmSnmMg.LvItem.ListItems(ItemID)
    TmpItm.SubItems(1) = Trim(ItemName)
    TmpItm.SubItems(2) = Format(ItemPrice, "#0.00")
    TmpItm.SmallIcon = ItemSymbol
Exit Sub

ErrInt:
    ErrLog Err, "SnmMg | ItemAdd"
End Sub

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [ItemDel] - Delete item
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ItemDel()
On Error GoTo ErrInt
    Dim SelItm As ListItem, StrRet As String, StrItemID As String
    Dim Rs As Recordset
    
  ' initialize
    Set SelItm = FrmSnmMg.LvItem.SelectedItem
    If SelItm Is Nothing Then Exit Sub
    
  ' prompt user
    StrRet = MsgBox(Var(2) & vbNewLine & SelItm.SubItems(1), vbOKCancel, CbMsgWarn)
    If StrRet = vbCancel Then Exit Sub
    
  ' delete item from database
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)
    StrItemID = SelItm.Key
    With Rs
        .FindFirst "Id = '" & StrItemID & "'"
        If .NoMatch = False Then
            .Delete
        End If
    End With
  ' delete item from listitem
    FrmSnmMg.LvItem.ListItems.Remove StrItemID
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | ItemDel"
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [LoadGroups] - Load all groups
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function LoadGroups() As Long
On Error GoTo ErrInt
    Dim Rs As Recordset, TmpIml As ImageList, TmpIc As ImageCombo
    Dim IconName As String, LngGrpCnt As Long
    
    Set Rs = uSDB.OpenRecordset("ServiceCategory", dbOpenSnapshot)
    Rs.Sort = "Id"
    Set Rs = Rs.OpenRecordset
    
  ' checking group counts
    If Rs.BOF = True Then Exit Function
    
  ' loading groups
    Set TmpIml = FrmSnmMg.Iml
    Set TmpIc = FrmSnmMg.IcGroup
    With Rs
        .MoveFirst
        Do Until .EOF = True
            IconName = !Symbol
            If IconName = "" Then IconName = "FOLDER"
            TmpIc.ComboItems.Add , !Id, !Name, IconName
            LngGrpCnt = LngGrpCnt + 1
            .MoveNext
        Loop
        TmpIc.ComboItems(1).Selected = True
        LoadGroups = LngGrpCnt
    End With
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | LoadGroups"
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [LoadItems] - Load all items
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function LoadItems() As Long
On Error GoTo ErrInt
    Dim Rs As Recordset, StrGrpId As String, StrItemSymbol As String
    Dim LngItmCnt As Long
    
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenSnapshot)
    StrGrpId = Left(FrmSnmMg.IcGroup.SelectedItem.Key, 5)
  ' sort dan filterkan data
    Rs.Sort = "id"
    Set Rs = RsFilter(Rs, "GroupId = '" & StrGrpId & "'")
    
  ' jika tiada data, keluar
    If Rs.BOF = True Then Exit Function
    With Rs
        .MoveFirst
        Do Until .EOF = True
            StrId = !Id
            StrItemSymbol = !Symbol
            Set TmpItm = FrmSnmMg.LvItem.ListItems.Add(, StrId, StrId, , StrItemSymbol)
            TmpItm.SubItems(1) = !Name & ""
            TmpItm.SubItems(2) = !Price & ""
            LngItmCnt = LngItmCnt + 1
            .MoveNext
        Loop
    End With
    LoadItems = LngItmCnt
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | LoadItems"
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GetItems] - Get Items Information
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub ItemGet(ItemID As String)
On Error GoTo ErrInt
    Dim Rs As Recordset, StrItemSymbol As String
    Dim StrItemStock As String
    
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenDynaset)

    With Rs
        .FindFirst "Id = '" & ItemID & "'"
        If .NoMatch = False Then
            StrItemStock = !Stock
            StrItemSymbol = !Symbol
            If !Stock = -1 Then StrItemStock = Var(3)
            FrmSnmMg.HdrItemId = !Id
            FrmSnmMg.ItemInfoTxt(0) = !Name
            FrmSnmMg.ItemInfoTxt(1) = !Price
            FrmSnmMg.ItemInfoTxt(2) = StrItemStock
            FrmSnmMg.ItemInfoTxt(3) = !Consume
            FrmSnmMg.ItemInfoTxt(4) = !LastPurchaseDateTime & ""
            FrmSnmMg.HdrItemImg.Picture = FrmSnmMg.Iml.ListImages(StrItemSymbol).Picture
            FrmSnmMg.ItemSymLv.ListItems(StrItemSymbol).Selected = True
            FrmSnmMg.ItemSymLv.ListItems(StrItemSymbol).EnsureVisible
        End If
    End With
    
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | ItemGet"
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GroupIdCheck] - Check and retrive group ID
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GroupIdCheck() As String
On Error GoTo ErrInt
    Dim Rs As Recordset, lngIdx As Long, BlnIdFound As Boolean
    Dim StrTmpId As String
    
    Set Rs = uSDB.OpenRecordset("ServiceCategory", dbOpenSnapshot)
    If Rs.BOF = False Then
        With Rs
            Do Until BlnIdFound = True
                lngIdx = lngIdx + 1
                StrTmpId = StrIdGrpHead & Format(lngIdx, "00")
                .FindFirst "Id = '" & StrTmpId & "'"
                BlnIdFound = .NoMatch
                .MoveNext
            Loop
            GroupIdCheck = StrTmpId
        End With
    Else
        GroupIdCheck = "CPG01"
    End If
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | GroupIdCheck"
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [ItemIdCheck] - Check and retrive item ID
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function ItemIdCheck(GroupId) As String
On Error GoTo ErrInt
    Dim Rs As Recordset, lngIdx As Long
    Dim StrTmpId As String
    
    Set Rs = uSDB.OpenRecordset("ServiceItems", dbOpenSnapshot)
    Rs.Sort = "Id"
    Set Rs = RsFilter(Rs, "GroupId = '" & GroupId & "'")
    If Rs.BOF = False Then
        With Rs
            .MoveFirst
            Do Until .EOF = True
                lngIdx = lngIdx + 1
                StrTmpId = StrIdHead & Right(GroupId, 2) & Format(lngIdx, "#000")
                If !Id <> StrTmpId Then
                    ItemIdCheck = StrTmpId
                    Exit Function
                End If
                .MoveNext
            Loop
            lngIdx = lngIdx + 1
            StrTmpId = StrIdHead & Right(GroupId, 2) & Format(lngIdx, "#000")
            ItemIdCheck = StrTmpId
        End With
    Else
        ItemIdCheck = StrIdHead & Right(GroupId, 2) & "001"
    End If
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | ItemIdCheck"
End Function


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [LoadSymbol] - LoadSymbol
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub LoadSymbol(CListView As ListView, Optional GroupOnly As Boolean = False)
    Dim StrLoadFilter As String, IntIdxA As Integer, IntIdxB As Integer
    Dim CListImg As ListImages
    
    StrLoadFilter = "ITEM"
    If GroupOnly = True Then StrLoadFilter = "GRP"
    Set CListImg = FrmSnmMg.Iml.ListImages
    
    For IntIdxA = 1 To CListImg.Count
        If CListImg(IntIdxA).Tag = StrLoadFilter Then
            IntIdxB = IntIdxB + 1
            CListView.ListItems.Add , CListImg(IntIdxA).Key, Format(IntIdxB, "00"), IntIdxA
        End If
    Next IntIdxA
        
End Sub
