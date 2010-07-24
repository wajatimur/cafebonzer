VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim iFrm As Form
Dim iLv As ListView

Private Sub Class_Initialize()
    Set iFrm = FrmMenView
    Set iLv = iFrm.Lv
End Sub

Public Sub AddItem(ItemName, ItemKey)
    iLv.ListItems.Add , ItemKey, ItemName
End Sub

Public Function ShowMenu(AnyFrm As Form, x, y) As ListItem
    iFrm.Left = AnyFrm.Left + x
    iFrm.Top = (AnyFrm.Top + x) - iFrm.Height
    iFrm.Show vbModal
    Set ShowMenu = iFrm.Lv.SelectedItem
End Function
