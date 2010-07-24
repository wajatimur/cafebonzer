VERSION 5.00
Begin VB.UserControl VsGuiSpinEdit 
   AutoRedraw      =   -1  'True
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   435
   ScaleWidth      =   2025
   ToolboxBitmap   =   "VsGuiEditSpin.ctx":0000
   Begin VB.TextBox SpinQty 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Text            =   "0"
      Top             =   60
      Width           =   780
   End
   Begin VB.VScrollBar SpinScroll 
      Height          =   330
      Left            =   1800
      Max             =   999
      Min             =   1
      TabIndex        =   0
      Top             =   45
      Value           =   999
      Width           =   165
   End
   Begin VB.Label SpinLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity :"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "VsGuiSpinEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : VsGuiSpinEdit
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private IntQuantity As Integer
Private BlnEnabled As Boolean


Private Sub SpinQty_Change()
    IntQuantity = SpinQty.Text
End Sub

Private Sub UserControl_InitProperties()
    IntQuantity = 0
    BlnEnabled = True
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 430
    UserControl.Width = 2020
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    IntQuantity = PropBag.ReadProperty("QUANTITY", 0)
    BlnEnabled = PropBag.ReadProperty("ENABLED", True)
    Call UpdateControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ENABLED", BlnEnabled, True
    PropBag.WriteProperty "QUANTITY", IntQuantity, 0
End Sub

Private Sub SpinScroll_Change()
    IntQuantity = 1000 - SpinScroll.Value
    SpinQty = IntQuantity
End Sub


Public Property Get EditValue() As Integer
Attribute EditValue.VB_UserMemId = 0
    EditValue = IntQuantity
End Property

Public Property Let EditValue(ByVal vNewValue As Integer)
    If IsNumeric(vNewValue) = True Then
        IntQuantity = vNewValue
        Call UpdateControl
    End If
End Property


Public Property Get Font() As Font
    Set Font = SpinLbl.Font
End Property

Public Property Let Font(ByVal vNewValue As Font)
    SpinQty.Font = vNewValue
    SpinLbl.Font = vNewValue
End Property


Public Property Get Enabled() As Boolean
    Enabled = BlnEnabled
End Property

Public Property Let Enabled(ByVal BlnNewValue As Boolean)
    BlnEnabled = BlnNewValue
    Call UpdateControl
End Property


Private Sub UpdateControl()
    SpinQty = IntQuantity
    SpinQty.Enabled = BlnEnabled
    SpinScroll.Enabled = BlnEnabled
End Sub

