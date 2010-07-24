Attribute VB_Name = "MdlSnm"
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GroupAdd] - Add new group
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub GroupAdd()
On Error GoTo ErrInt
    Dim Rs As Recordset, strGrpId As String, strGrp As String
    Dim ObjInput As New VsInput, tmpIml As ImageList, tmpIc As ImageCombo
  
  ' minta nama group baru dari pengguna
    strGrp = ObjInput.GetInput("Group Name", BtnClose)
    If Trim(strGrp) = "" Then Exit Sub
  ' terima GroupId yang sah
    strGrpId = GroupIdCheck
    
  ' tambah Group ke dalam db
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenDynaset)
    With Rs
        .FindFirst "name = '" & strGrp & "'"
        If .NoMatch = True Then
            .AddNew
            !id = strGrpId
            !Name = strGrp
            .Update
        Else
            Exit Sub
        End If
    End With
    
  ' tambah ke dalam control IcGroup
    Set tmpIml = FrmSnmMg.Iml
    Set tmpIc = FrmSnmMg.IcGroup
    IconName = "folder"
  ' cek nama icons
    For t = 1 To tmpIml.ListImages.Count
        If LCase(tmpIml.ListImages(t).Key) = LCase(strGrp) Then IconName = LCase(strGrp)
    Next t
    tmpIc.ComboItems.Add , "g" & strGrpId, strGrp, IconName
    If tmpIc.ComboItems.Count = 1 Then
        tmpIc.ComboItems("g" & strGrpId).Selected = True
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
    Dim SelGrpId As String, strRet As String
    
  ' jika jumlah group = 0, keluar
    If FrmSnmMg.IcGroup.ComboItems.Count = 0 Then Exit Sub
  ' tanya pengguna
    strRet = MsgBox(Var(1), vbOKCancel, CbMsgWarn)
    If strRet = vbOK Then
        Set Rs = uSDB.OpenRecordset("pos-category", dbOpenTable)
      ' ambil group id
        SelGrpId = Mid(FrmSnmMg.IcGroup.SelectedItem.Key, 2)
      ' seek dan delete group
        With Rs
            .Index = "id"
            .Seek "=", SelGrpId
            If .NoMatch = True Then Exit Sub
            .Delete
        End With
        
      ' filter dan delete items bagi group tersebut
        Set Rs = uSDB.OpenRecordset("pos-items", dbOpenDynaset)
        Set Rs = RsFilter(Rs, "groupid = '" & SelGrpId & "'")
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
        FrmSnmMg.IcGroup.ComboItems.Remove "g" & SelGrpId
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
Public Sub ItemAdd(GroupId As String, ItemName As String, ItemPrice As String)
On Error GoTo ErrInt
    Dim Rs As Recordset, strItemId As String, tmpItm As ListItem
    Set Rs = uSDB.OpenRecordset("pos-items", dbOpenDynaset)
    
  ' retrive good item id
    strItemId = ItemIdCheck(GroupId)
    With Rs
      ' search for duplicate items
        .FindFirst "nama = '" & ItemName & "'"
        If .NoMatch = True Then
          ' add item to database
            .AddNew
            !GroupId = GroupId
            !id = strItemId
            !nama = Trim(ItemName)
            !Harga = Format(ItemPrice, "#0.00")
            .Update
        Else
          ' exit from this sub
            Exit Sub
        End If
    End With
    
  ' add current new item to listview
    If FrmSnmMg.IcGroup.SelectedItem.Key = "g" & GroupId Then
        Set tmpItm = FrmSnmMg.LvItem.ListItems.Add(, strItemId, strItemId, , "item")
        tmpItm.SubItems(1) = Trim(ItemName)
        tmpItm.SubItems(2) = Format(ItemPrice, "#0.00")
        'tmpItm.SubItems(3) = !stok & ""
    End If
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
    Dim SelItm As ListItem, strRet As String, strItemId As String
    Dim Rs As Recordset
    
  ' initialize
    Set SelItm = FrmSnmMg.LvItem.SelectedItem
    If SelItm Is Nothing Then Exit Sub
    
  ' prompt user
    strRet = MsgBox(Var(2) & vbNewLine & SelItm.SubItems(1), vbOKCancel, CbMsgWarn)
    If strRet = vbCancel Then Exit Sub
    
  ' delete item from database
    Set Rs = uSDB.OpenRecordset("pos-items", dbOpenDynaset)
    strItemId = SelItm.Key
    With Rs
        .FindFirst "id = '" & strItemId & "'"
        If .NoMatch = False Then
            .Delete
        End If
    End With
  ' delete item from listitem
    FrmSnmMg.LvItem.ListItems.Remove strItemId
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
    Dim Rs As Recordset, tmpIml As ImageList, tmpIc As ImageCombo
    Dim IconName As String, lngGrpCnt As Long
    
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenSnapshot)
    Rs.Sort = "id"
    Set Rs = Rs.OpenRecordset
    
  ' checking group counts
    If Rs.BOF = True Then Exit Function
    
  ' loading groups
    Set tmpIml = FrmSnmMg.Iml
    Set tmpIc = FrmSnmMg.IcGroup
    With Rs
        .MoveFirst
        Do Until .EOF = True
            IconName = "folder"
            For t = 1 To tmpIml.ListImages.Count
                If LCase(tmpIml.ListImages(t).Key) = LCase(!Name) Then IconName = LCase(!Name)
            Next t
            tmpIc.ComboItems.Add , "g" & !id, !Name, IconName
            lngGrpCnt = lngGrpCnt + 1
            .MoveNext
        Loop
        tmpIc.ComboItems(1).Selected = True
        LoadGroups = lngGrpCnt
    End With
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | LoadGroups"
End Function

'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [LoadItems] - Load all items
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Sub LoadItems()
On Error GoTo ErrInt
    Dim Rs As Recordset, strGrpId As String
    
    Set Rs = uSDB.OpenRecordset("pos-items", dbOpenSnapshot)
    strGrpId = Mid(FrmSnmMg.IcGroup.SelectedItem.Key, 2)
  ' sort dan filterkan data
    Rs.Sort = "id"
    Set Rs = RsFilter(Rs, "GroupId = '" & strGrpId & "'")
    
  ' jika tiada data, keluar
    If Rs.BOF = True Then Exit Sub
    With Rs
        .MoveFirst
        Do Until .EOF = True
            strId = !id
            Set tmpItm = FrmSnmMg.LvItem.ListItems.Add(, strId, strId, , "item")
            tmpItm.SubItems(1) = !nama & ""
            tmpItm.SubItems(2) = !Harga & ""
            tmpItm.SubItems(3) = !stok & ""
            .MoveNext
        Loop
    End With
Exit Sub

ErrInt:
    ErrLog Err, "Snm Manager | LoadItems"
End Sub


'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
' [GroupIdCheck] - Check and retrive group ID
'
'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'*'
Public Function GroupIdCheck() As String
On Error GoTo ErrInt
    Dim Rs As Recordset, lngIdx As Long
    Dim strTmpId As String
    
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenSnapshot)
    If Rs.BOF = False Then
        With Rs
            .MoveFirst
            Do Until .EOF = True
                lngIdx = lngIdx + 1
                strTmpId = Format(lngIdx, "#00")
                If !id <> strTmpId Then
                    GroupIdCheck = strTmpId
                    Exit Function
                End If
                .MoveNext
            Loop
            lngIdx = lngIdx + 1
            strTmpId = Format(lngIdx, "#00")
            GroupIdCheck = strTmpId
        End With
    Else
        GroupIdCheck = "01"
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
    Dim strTmpId As String
    
    Set Rs = uSDB.OpenRecordset("pos-items", dbOpenSnapshot)
    Rs.Sort = "id"
    Set Rs = RsFilter(Rs, "GroupId = '" & GroupId & "'")
    If Rs.BOF = False Then
        With Rs
            .MoveFirst
            Do Until .EOF = True
                lngIdx = lngIdx + 1
                strTmpId = "p" & GroupId & Format(lngIdx, "#000")
                If !id <> strTmpId Then
                    ItemIdCheck = strTmpId
                    Exit Function
                End If
                .MoveNext
            Loop
            lngIdx = lngIdx + 1
            strTmpId = "p" & GroupId & Format(lngIdx, "#000")
            ItemIdCheck = strTmpId
        End With
    Else
        ItemIdCheck = "p" & GroupId & "001"
    End If
Exit Function

ErrInt:
    ErrLog Err, "Snm Manager | ItemIdCheck"
End Function
