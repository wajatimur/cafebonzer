VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPlview 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2520
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMenView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2325
   Begin VB.PictureBox DockBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   0
      ScaleHeight     =   2550
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   0
      Width           =   285
   End
   Begin MSComctlLib.ListView Lv 
      Height          =   2520
      Left            =   285
      TabIndex        =   0
      Top             =   0
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2999
      EndProperty
   End
End
Attribute VB_Name = "FrmPlview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    CurM = "(" & Month(Date) & ")" & Year(Date)
    PutOnTop Me.hwnd
    Lv.ListItems.Clear
    INIenumSection App.Path & "\rekod\pelanggan.d"
    For d = 0 To UBound(EnumArray) - 1
        Lv.ListItems.Add , "cus" & d, EnumArray(d)
    Next d
End Sub

Private Sub Form_Deactivate()
    FrmPlview.Hide
End Sub

Private Sub Form_LostFocus()
    FrmPlview.Hide
End Sub

Private Sub Lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FrmGuna.Text1 = Item.Text
End Sub

Private Sub Lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmPlview.Hide
End Sub
