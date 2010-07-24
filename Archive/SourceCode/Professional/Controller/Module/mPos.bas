Attribute VB_Name = "mPos"
Public Sub LoadPosCatCB(iCboxName As ImageCombo, imglist As ImageList)
    'rekodset kosong.. terus tambah "None"
    Set Rs = uSDB.OpenRecordset("pos-category", dbOpenSnapshot)
    If Rs.BOF = True Then GoTo AddNone
    With Rs
        .MoveFirst
        Do Until .EOF = True
            iconName = "folder"
            
            'check the availablelity of the icon
            For t = 1 To imglist.ListImages.Count
                If LCase(imglist.ListImages(t).Key) = LCase(!Name) Then iconName = LCase(!Name)
            Next t
            'tambah ke dalam imagecombo
            iCboxName.ComboItems.Add , "g" & !id, !Name, iconName, iconName, 1
            .MoveNext
        Loop
    End With
    
AddNone:
    iCboxName.ComboItems.Add , VS(1), VS(1), "none", "none"
    iCboxName.ComboItems.Item(VS(1)).Selected = True
End Sub

Public Function LoadPosItmCB(iCboxName As ImageCombo, GrpID)
    Dim tInt As Integer, tWd As Integer
    Dim Rss As Recordset
    Dim CbItm As ComboItem
    Set Rss = uSDB.OpenRecordset("pos-items", dbOpenSnapshot)
    
    iCboxName.ComboItems.Clear
    Rss.Filter = "groupid = '" & GrpID & "'"
    Set Rs = Rss.OpenRecordset
    
    If Rs.BOF = True Then Exit Function
    With Rs
        .MoveFirst
        Do Until .EOF = True
            Set CbItm = iCboxName.ComboItems.Add(, !id, !Nama, , , 1)
            CbItm.Tag = !Harga
            .MoveNext
        Loop
    End With
    
    'resize the list width
    For g = 1 To iCboxName.ComboItems.Count
        tWd = FrmMain.TextWidth(iCboxName.ComboItems(g).Text)
        If tWd > tInt Then tInt = tWd
    Next g
    tInt = (tInt / 15) + 40
    ret = SendMessage(iCboxName.hwnd, CB_SETDROPPEDWIDTH, tInt, 0)
    
    'default select
    iCboxName.ComboItems(1).Selected = True
    iCboxName.Refresh
End Function
