Attribute VB_Name = "MdlServices"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlServices
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Public Sub LoadPosCatCB(iCboxName As ImageCombo, ImgList As ImageList, Optional ClearRecents As Boolean = False)
    Dim StrGrpId As String
    Set CRset = CDataS.OpenRecordset("ServiceCategory", dbOpenSnapshot)
    
    iCboxName.ImageList = ImgList
    If ClearRecents = True Then iCboxName.ComboItems.Clear
    
    'Rekodset kosong.. terus tambah "None"
    If CRset.BOF = True Then GoTo AddNone
    With CRset
        .MoveFirst
        Do Until .EOF = True
            'Tambah ke dalam imagecombo
            StrGrpId = !Id
            iCboxName.ComboItems.Add , CStr(!Id), !Name, CStr(!Symbol), CStr(!Symbol), 1
            .MoveNext
        Loop
    End With
    
AddNone:
    iCboxName.ComboItems.Add , VS(0, 1), VS(0, 1), "NONE", "NONE"
    iCboxName.ComboItems.Item(VS(0, 1)).Selected = True
End Sub


Public Function LoadPosItmCB(iCboxName As ImageCombo, GroupID, Optional ImgList As ImageList)
    Dim tInt As Integer, tWd As Integer
    Dim Rss As Recordset
    Dim CbItm As ComboItem
    Set Rss = CDataS.OpenRecordset("ServiceItems", dbOpenSnapshot)
    
    iCboxName.ComboItems.Clear
    Rss.Filter = "GroupId = '" & GroupID & "'"
    Set CRset = Rss.OpenRecordset
    
    If CRset.BOF = True Then Exit Function
    With CRset
        .MoveFirst
        Do Until .EOF = True
            Set CbItm = iCboxName.ComboItems.Add(, !Id, !Name, , , 1)
            CbItm.Tag = !Price
            .MoveNext
        Loop
    End With
    
    'resize the list width
    For g = 1 To iCboxName.ComboItems.Count
        tWd = FrmMain.TextWidth(iCboxName.ComboItems(g).Text)
        If tWd > tInt Then tInt = tWd
    Next g
    tInt = (tInt / 15) + 40
    ret = SendMessage(iCboxName.Hwnd, CB_SETDROPPEDWIDTH, tInt, 0)
    
    'default select
    iCboxName.ComboItems(1).Selected = True
    iCboxName.Refresh
End Function


Public Function LoadPosItmLV(ClistView As ListView, GroupID, Optional ImgList As ImageList) As Boolean
    Dim Rss As Recordset, TmpItem As ListItem
    Dim StrStock As String
    Set Rss = CDataS.OpenRecordset("ServiceItems", dbOpenSnapshot)
    
    LoadPosItmLV = False
    ClistView.SmallIcons = ImgList
    ClistView.ListItems.Clear
    Rss.Filter = "GroupId = '" & GroupID & "'"
    Set CRset = Rss.OpenRecordset
    
    If CRset.BOF = True Then Exit Function
    With CRset
        .MoveFirst
        Do Until .EOF = True
            StrStock = !Stock
            If !Stock = -1 Then StrStock = VS(2, 1)
            Set TmpItem = ClistView.ListItems.Add(, !Id, !Name, , CStr(!Symbol))
            TmpItem.SubItems(1) = !Price
            TmpItem.SubItems(2) = StrStock
            TmpItem.Tag = !GroupID
            .MoveNext
        Loop
    End With
    
    ClistView.ListItems(1).Selected = True
    LoadPosItmLV = True
End Function

