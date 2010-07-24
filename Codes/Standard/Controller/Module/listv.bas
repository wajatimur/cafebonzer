Attribute VB_Name = "MdlControls"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlControls
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : Controls
' Description         :
'==================================================================

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan subitem semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function CtlCheckNull(ControlName As Control, Warning As String) As Boolean
    CtlCheckNull = True
    If ControlName = "" Then
        MsgBox Warning, vbInformation, CbMsgWarn
        ControlName.SetFocus
        CtlCheckNull = False
    End If
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan subitem semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function CtlCheckMatch(ControlName1 As Control, ControlName2 As Control, Warning As String) As Boolean
    CtlCheckMatch = True
    If ControlName1.Text <> ControlName2.Text Then
        MsgBox Warning, vbInformation, CbMsgWarn
        ControlName1.SetFocus
        CtlCheckMatch = False
    End If
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Benarkan nombor sahaja
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub NumOnly(TextControl As TextBox)
    Dim Style As Long
    
    Style = GetWindowLong(TextControl.Hwnd, GWL_STYLE)
    SetWindowLong TextControl.Hwnd, GWL_STYLE, Style Or ES_NUMBER
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ComboBox Add With Trim
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Public Sub CbAddEx(Item, Cbox As ComboBox)
    For a = 0 To Cbox.ListCount - 1
        If Trim(Cbox.List(a)) = Trim(Item) Then Exit Sub
    Next a
    Cbox.AddItem Item
End Sub


Public Sub CbSelect(Item, CCbox As ComboBox)
    Dim IntIdxA As Integer
    If CCbox.ListCount = 0 Then Exit Sub
    CCbox.ListIndex = 0
    For IntIdxA = 0 To CCbox.ListCount - 1
        If Trim(CCbox.List(IntIdxA)) = Trim(Item) Then
            CCbox.ListIndex = IntIdxA
            Exit Sub
        End If
    Next IntIdxA
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan subitem semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelSubItm(Index As Integer) As String
    If UniAgents.Count = 0 Then Exit Function
    SelSubItm = FrmMain.ListView.SelectedItem.SubItems(Index)
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan Text semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelText() As String
    If UniAgents.Count = 0 Then Exit Function
    SelText = FrmMain.ListView.SelectedItem.Text
End Function

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Dapatkan Tag semasa
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function SelTag() As String
    If UniAgents.Count = 0 Then Exit Function
    SelTag = FrmMain.ListView.SelectedItem.Tag
End Function




