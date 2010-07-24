VERSION 5.00
Begin VB.UserControl Label3D 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Label3D.ctx":0000
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Label3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum T_Phase
    TOPLEFT = 3
    TOPRIGHT = 2
    BOTTOMLEFT = 1
    BOTTOMRIGHT = 0
End Enum

Public Enum T_BorderStyle
    None = 0
    FixedSingle = 1
End Enum

Public Enum T_Align
    AlignLeft = 0
    AlignRight = 1
    AlignCenter = 2
End Enum

'Events declaration
Event Click() 'MappingInfo=Label1,Label1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Label1,Label1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Label1,Label1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Label1,Label1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Label1,Label1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'default variabled definition
Const m_def_Phase = 0
'veriables definition
Dim m_Phase As Byte

Private Sub UserControl_Initialize()
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    m_Phase = m_def_Phase
    Set Label1.Font = Ambient.Font
    Set Label2.Font = Ambient.Font
End Sub

Private Sub UserControl_Resize()
    If Width < 100 Then Width = 100
    If Height < 100 Then Height = 100
    Label1.Width = Width
    Label1.Height = Height
    Label2.Width = Width
    Label2.Height = Height
    Select Case m_Phase
    Case 3
        Label1.Left = 0
        Label1.Top = 0
        Label2.Left = 15
        Label2.Top = 15
    Case 2
        Label1.Left = 15
        Label1.Top = 0
        Label2.Left = 0
        Label2.Top = 15
    Case 1
        Label1.Left = 0
        Label1.Top = 15
        Label2.Left = 15
        Label2.Top = 0
    Case 0
        Label1.Left = 15
        Label1.Top = 15
        Label2.Left = 0
        Label2.Top = 0
    End Select
End Sub

'MappingInfo=Label1,Label1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Label1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Label1.Enabled() = New_Enabled
    Label2.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    Set Label2.Font = New_Font
    PropertyChanged "Font"
End Property

'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub Label2_click()
    RaiseEvent Click
End Sub

Private Sub Label2_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'MappingInfo=Label1,Label1,-1,ForeColor
Public Property Get ForeColor1() As OLE_COLOR
Attribute ForeColor1.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor1 = Label1.ForeColor
End Property

Public Property Let ForeColor1(ByVal New_ForeColor1 As OLE_COLOR)
    Label1.ForeColor() = New_ForeColor1
    PropertyChanged "ForeColor1"
End Property

'MappingInfo=Label2,Label2,-1,ForeColor
Public Property Get ForeColor2() As OLE_COLOR
Attribute ForeColor2.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor2 = Label2.ForeColor
End Property

Public Property Let ForeColor2(ByVal New_ForeColor2 As OLE_COLOR)
    Label2.ForeColor() = New_ForeColor2
    PropertyChanged "ForeColor2"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Enabled = PropBag.ReadProperty("Enabled", True)
    Label2.Enabled = PropBag.ReadProperty("Enabled", True)
    Label1.ForeColor = PropBag.ReadProperty("ForeColor1", &H0)
    Label2.ForeColor = PropBag.ReadProperty("ForeColor2", &HFFFFFF)
    Label1.Caption = PropBag.ReadProperty("Caption", "Label3D")
    Label2.Caption = PropBag.ReadProperty("Caption", "Label3D")
    Label1.Alignment = PropBag.ReadProperty("Alignment", 0)
    Label2.Alignment = PropBag.ReadProperty("Alignment", 0)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Label2.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Label2.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    m_Phase = PropBag.ReadProperty("Phase", m_def_Phase)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", Label1.Enabled, True)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor1", Label1.ForeColor, &H0)
    Call PropBag.WriteProperty("ForeColor2", Label2.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "Label3D")
    Call PropBag.WriteProperty("Alignment", Label1.Alignment, 0)
    Call PropBag.WriteProperty("Phase", m_Phase, m_def_Phase)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", Label2.BorderStyle, 0)
End Sub

'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    Label2.Caption = New_Caption
    PropertyChanged "Caption"
End Property

'MappingInfo=Label1,Label1,-1,Alignment
Public Property Get Alignment() As T_Align
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As T_Align)
    Label1.Alignment() = New_Alignment
    Label2.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get Phase() As T_Phase
    Phase = m_Phase
End Property

Public Property Let Phase(ByVal New_Phase As T_Phase)
    m_Phase = New_Phase
    UserControl_Resize
    PropertyChanged "Phase"
End Property

'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'MappingInfo=Label2,Label2,-1,BorderStyle
Public Property Get BorderStyle() As T_BorderStyle
Attribute BorderStyle.VB_Description = "Restituisce o imposta lo stile del bordo di un oggetto."
    BorderStyle = Label2.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As T_BorderStyle)
    Label2.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

