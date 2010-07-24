VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Lv As ListView
Attribute Lv.VB_VarHelpID = -1
Private c_Cb As ComboBox
Private StackMatrix As New Collection


Public Sub Init(Lview As ListView)
    Set Lv = Lview
End Sub

Private Sub Class_Terminate()
    Set Lv = Nothing
    Set c_Cb = Nothing
    Set StackMatrix = Nothing
End Sub

Public Function ItemAdd(Optional Key, Optional Text, Optional SmallIcon) As ListItem
'On Error GoTo ErrInt
    Dim NewItm As ListItem
    Dim NewDts As New clsDataStore
    
    Set NewItm = Lv.ListItems.Add(, Key, Text, , SmallIcon)
    NewDts.Name = Key
    StackMatrix.Add NewDts, Key
    Set ItemAdd = NewItm
    Set NewItm = Nothing
    Set NewDts = Nothing
Exit Function

ErrInt:
    AppErrorLog Err, "clsSpecial | AddItem"
End Function

Public Sub ItemRemove(Key)
    StackMatrix.Remove Key
    Lv.ListItems.Remove Key
End Sub

Public Sub ItemClear()
    Lv.ListItems.Clear
    Set StackMatrix = Nothing
End Sub


Public Sub MatrixAdd(ParentKey, Data)
    Dim DTS As clsDataStore

    Set DTS = StackMatrix(ParentKey)
    DTS.Add Data, DTS.Count + 1
    Set DTS = Nothing
End Sub

Public Sub MatrixExpand(Parent As ListItem)
On Error GoTo ErrInt
    Dim t_Rect As RECT, l_TwipsX As Long, l_TwipsY As Long
    Dim v_Col As Variant, l_mPos As Long, DTS As clsDataStore
    
    If Parent.Key = "" Then Exit Sub
    Set DTS = StackMatrix(Parent.Key)
    If DTS.Count = 0 Then Exit Sub
    
    l_mPos = 1
    l_TwipsX = Screen.TwipsPerPixelX
    l_TwipsY = Screen.TwipsPerPixelY
    Set c_Cb = FrmAgnInfo.DynaCombo
    
    SetParent c_Cb.hwnd, Lv.hwnd
    GetSubItemRect Lv.hwnd, Parent.Index - 1, l_mPos, LVIR_LABEL, t_Rect
    
   'set position
    With t_Rect
        c_Cb.Left = (.Left * l_TwipsX)
        c_Cb.Top = (.Top * l_TwipsY)
        c_Cb.Width = (.Right - .Left) * l_TwipsX
    End With
    
   'load all data to combobox
    c_Cb.Clear
    If DTS.Count = 0 Then
        c_Cb.AddItem "No Printer"
    Else
        c_Cb.AddItem "All Printer"
        For a = 1 To DTS.Count
            c_Cb.AddItem DTS(a)
        Next
    End If
    c_Cb.ListIndex = 0
    c_Cb.Visible = True
    Set DTS = Nothing
Exit Sub

ErrInt:
    AppErrorLog Err, "clsSpecial | MatrixExpand"
End Sub

Public Sub MatrixClear(Key)
On Error GoTo ErrInt
    Dim DTS As clsDataStore
    Lv.ListItems.Clear
    Set DTS = StackMatrix(Key)
    DTS.Clear
    Set DTS = Nothing
Exit Sub
ErrInt:
    AppErrorLog Err, "clsSpecial | MatrixClear"
End Sub


Private Function GetSubItemRect(ByVal hWndLV As Long, ByVal iItem As Long, ByVal iSubItem As Long, ByVal code As Long, lpRect As RECT) As Boolean
  lpRect.Top = iSubItem
  lpRect.Left = code
  GetSubItemRect = SendMessage(hWndLV, LVM_GETSUBITEMRECT, ByVal iItem, lpRect)
End Function

