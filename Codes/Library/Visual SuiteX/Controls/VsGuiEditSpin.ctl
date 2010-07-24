VERSION 5.00
Begin VB.UserControl VsGuiSpinEdit 
   BackStyle       =   0  'Transparent
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
      Enabled         =   0   'False
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Text            =   "1"
      Top             =   60
      Width           =   780
   End
   Begin VB.VScrollBar SpinScroll 
      Enabled         =   0   'False
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
Attribute VB_Exposed = True
Private IntQuantity As Integer


Private Sub SpinScroll_Change()
    IntQuantity = 1000 - SpinScroll.Value
    SpinQty = IntQuantity
End Sub


Public Property Get EditValue() As Variant
    EditValue = IntQuantity
End Property

Public Property Let EditValue(ByVal vNewValue As Integer)
    If IsNumeric(vNewValue) = True Then
        IntQuantity = vNewValue
    End If
End Property


Public Property Get Font() As Font
    Font = SpinLbl.Font
End Property

Public Property Let Font(ByVal vNewValue As Font)
    SpinQty.Font = vNewValue
    SpinLbl.Font = vNewValue
End Property
