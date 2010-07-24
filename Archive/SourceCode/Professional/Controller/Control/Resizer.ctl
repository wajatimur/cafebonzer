VERSION 5.00
Begin VB.UserControl Resizer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Resizer.ctx":0000
   Begin VB.Label lblResizer 
      BackStyle       =   0  'Transparent
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4395
   End
End
Attribute VB_Name = "Resizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'********************* RESIZER CONTROL ********************
'Created by Tincani Andrea                                  26-4-1999
'Update 1.2.0
'__________________________________________________________
'Find more FREE Source Code at
'http://pages.hotbot.com/edu/tincani.andrea/index.html

'IMPORTANT: You must include in the Controls the Windows Common Controls
'because my control uses some data types defined in the MsComCtlLib!!!

'Feel free to mail at tincani.andrea@hotbot.com for any explanation, question
'or bug report about this control...

Option Explicit

Dim posy As Single
Dim posx As Single
'Default properties value
Const m_def_InvertControls = False
Const m_def_Orientation = 1
Const m_def_SeparatorSize = 45
Const m_def_ControlName1 = ""
Const m_def_ControlName2 = ""
Const m_def_MinControlSize1 = 90
Const m_def_MinControlSize2 = 90
Const m_def_Control1Visible = True
Const m_def_Control2Visible = True
Const m_def_Control1Size = 0
Const m_def_Control2Size = 0
'Properties Variables
Dim m_InvertControls As Boolean
Dim m_Orientation As OrientationConstants
Dim m_SeparatorSize As Long
Dim m_ControlName1 As String
Dim m_ControlName2 As String
Dim m_MinControlSize(0 To 1) As Long
Dim m_ControlVisible(0 To 1) As Boolean
Dim m_Control1Size As Long
Dim m_Control2Size As Long

Private Sub lblResizer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        posy = y
        posx = x
    End If
End Sub

'When the user moves the separator Bar
Private Sub lblResizer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer
    Dim MinPoint As Integer
    Static inloop As Boolean 'Static variable for non recursive calling
    
    If inloop Then Exit Sub
    inloop = True
    On Error Resume Next
    'Get the first two controls inserted into the resizer
    If m_InvertControls Then
        i = 1
        MinPoint = 1
    Else
        i = 0
        MinPoint = 0
    End If
    If Button = vbLeftButton Then
        For Each c In UserControl.ContainedControls
            If Not c Is lblResizer Then
                Set cnt(i) = c
                If m_InvertControls Then i = i - 1 Else i = i + 1
                If i = 2 Or i = -1 Then Exit For
            End If
        Next
    End If
    'Apply the new size to the two controls
    If Button = vbLeftButton And Not cnt(0) Is Nothing And Not cnt(1) Is Nothing Then
        Select Case m_Orientation
        Case ccOrientationHorizontal
            If cnt(0).Height - posy + y < m_MinControlSize(MinPoint) Then
                posy = cnt(0).Height + y - m_MinControlSize(MinPoint)
            End If
            If cnt(1).Height + posy - y < m_MinControlSize(1 - MinPoint) Then
                posy = y + m_MinControlSize(1 - MinPoint) - cnt(1).Height
            End If
            cnt(0).Move cnt(0).Left, cnt(0).Top, cnt(0).Width, cnt(0).Height - posy + y
            cnt(1).Move cnt(1).Left, cnt(1).Top - posy + y, cnt(1).Width, cnt(1).Height + posy - y
            posy = y
        Case ccOrientationVertical
            If cnt(0).Width - posx + x < m_MinControlSize(MinPoint) Then
                posx = cnt(0).Width + x - m_MinControlSize(MinPoint)
            End If
            If cnt(1).Width + posx - x < m_MinControlSize(1 - MinPoint) Then
                posx = x + m_MinControlSize(1 - MinPoint) - cnt(1).Width
            End If
            cnt(0).Move cnt(0).Left, cnt(0).Top, cnt(0).Width - posx + x, cnt(0).Height
            cnt(1).Move cnt(1).Left - posx + x, cnt(1).Top, cnt(1).Width + posx - x, cnt(1).Height
            posx = x
        End Select
    End If
    inloop = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SeparatorSize = PropBag.ReadProperty("SeparatorSize", m_def_SeparatorSize)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_InvertControls = PropBag.ReadProperty("InvertControls", m_def_InvertControls)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ControlName1 = PropBag.ReadProperty("ControlName1", m_def_ControlName1)
    m_ControlName2 = PropBag.ReadProperty("ControlName2", m_def_ControlName2)
    m_MinControlSize(0) = PropBag.ReadProperty("MinControlSize1", m_def_MinControlSize1)
    m_MinControlSize(1) = PropBag.ReadProperty("MinControlSize2", m_def_MinControlSize2)
    m_ControlVisible(0) = PropBag.ReadProperty("Control1Visible", m_def_Control1Visible)
    m_ControlVisible(1) = PropBag.ReadProperty("Control2Visible", m_def_Control2Visible)
    m_Control1Size = PropBag.ReadProperty("Control1Size", m_def_Control1Size)
    m_Control2Size = PropBag.ReadProperty("Control2Size", m_def_Control2Size)
End Sub

Private Sub UserControl_Resize()
    UserControl_Show
End Sub

'Initialize the position of the controls contained ino the resizer
Private Sub UserControl_Show()
    Dim m_FirstControlSize As Integer
    Dim cnt(0 To 1) As Control
    Dim i As Integer
    Dim c As Control
    Dim h As Long
    Dim w As Long
    Dim MinPoint As Integer
    
    lblResizer.Move 0, 0, Width, Height
    On Error Resume Next
    If m_InvertControls Then
        i = 1
        MinPoint = 1
    Else
        i = 0
        MinPoint = 0
    End If
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If cnt(0) Is Nothing Or cnt(1) Is Nothing Then Exit Sub
    If Not m_ControlVisible(0) Then
        If m_InvertControls Then i = 1 Else i = 0
        cnt(i).Visible = False
        cnt(i).Move UserControl.Width
        If TypeOf cnt(1 - i) Is MSComctlLib.ListView Then
            cnt(1 - i).Move -15, -15, UserControl.ScaleWidth + 30, UserControl.ScaleHeight + 30
        Else
            cnt(1 - i).Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
    ElseIf Not m_ControlVisible(1) Then
        If m_InvertControls Then i = 0 Else i = 1
        cnt(i).Visible = False
        cnt(i).Move UserControl.Width
        If TypeOf cnt(1 - i) Is MSComctlLib.ListView Then
            cnt(1 - i).Move -15, -15, UserControl.ScaleWidth + 30, UserControl.ScaleHeight + 30
        Else
            cnt(1 - i).Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        End If
    Else
        If Not cnt(0).Visible Or Not cnt(1).Visible Then
            cnt(0).Visible = True
            cnt(1).Visible = True
            Select Case m_Orientation
            Case ccOrientationHorizontal
                cnt(0).Height = (UserControl.Height - m_SeparatorSize) \ 2
                cnt(1).Height = (UserControl.Height - m_SeparatorSize) \ 2
            Case ccOrientationVertical
                cnt(0).Width = (UserControl.Width - m_SeparatorSize) \ 2
                cnt(1).Width = (UserControl.Width - m_SeparatorSize) \ 2
            End Select
        End If
        'Select the type of orientation
        Select Case m_Orientation
        Case ccOrientationHorizontal
            'Horizontal separator
            m_FirstControlSize = cnt(0).Height
            lblResizer.MousePointer = vbSizeNS
            If m_FirstControlSize < m_MinControlSize(MinPoint) Then
                m_FirstControlSize = m_MinControlSize(MinPoint)
            ElseIf UserControl.ScaleHeight - m_FirstControlSize - m_SeparatorSize < m_MinControlSize(1 - MinPoint) Then
                m_FirstControlSize = UserControl.ScaleHeight - m_SeparatorSize - m_MinControlSize(1 - MinPoint)
            End If
            'test if the control contained is a ListView (it has differents size values!!)
            If TypeOf cnt(0) Is MSComctlLib.ListView Then
                cnt(0).Move -15, -15, UserControl.ScaleWidth + 30, m_FirstControlSize
                h = cnt(0).Height - 30
            Else
                cnt(0).Move 0, 0, UserControl.ScaleWidth, m_FirstControlSize
                h = cnt(0).Height
            End If
            If TypeOf cnt(1) Is MSComctlLib.ListView Then
                cnt(1).Move -15, h + m_SeparatorSize - 15, UserControl.ScaleWidth + 30, UserControl.ScaleHeight - h - m_SeparatorSize
            Else
                cnt(1).Move 0, h + m_SeparatorSize, UserControl.ScaleWidth, UserControl.ScaleHeight - h - m_SeparatorSize
            End If
        Case ccOrientationVertical
            'Vertical Separator
            m_FirstControlSize = cnt(0).Width
            lblResizer.MousePointer = vbSizeWE
            If m_FirstControlSize < m_MinControlSize(MinPoint) Then
                m_FirstControlSize = m_MinControlSize(MinPoint)
            ElseIf UserControl.ScaleWidth - m_FirstControlSize - m_SeparatorSize < m_MinControlSize(1 - MinPoint) Then
                m_FirstControlSize = UserControl.ScaleWidth - m_SeparatorSize - m_MinControlSize(1 - MinPoint)
            End If
            If TypeOf cnt(0) Is MSComctlLib.ListView Then
                cnt(0).Move -15, -15, m_FirstControlSize, UserControl.ScaleHeight + 30
                w = cnt(0).Width - 30
            Else
                cnt(0).Move 0, 0, m_FirstControlSize, UserControl.ScaleHeight
                w = cnt(0).Width
            End If
            If TypeOf cnt(1) Is MSComctlLib.ListView Then
                cnt(1).Move w + m_SeparatorSize - 15, -15, UserControl.ScaleWidth - w - m_SeparatorSize, UserControl.ScaleHeight + 30
            Else
                cnt(1).Move w + m_SeparatorSize, 0, UserControl.ScaleWidth - w - m_SeparatorSize, UserControl.ScaleHeight
            End If
        End Select
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SeparatorSize", m_SeparatorSize, m_def_SeparatorSize)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("InvertControls", m_InvertControls, m_def_InvertControls)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ControlName1", m_ControlName1, m_def_ControlName1)
    Call PropBag.WriteProperty("ControlName2", m_ControlName2, m_def_ControlName2)
    Call PropBag.WriteProperty("MinControlSize1", m_MinControlSize(0), m_def_MinControlSize1)
    Call PropBag.WriteProperty("MinControlSize2", m_MinControlSize(1), m_def_MinControlSize2)
    Call PropBag.WriteProperty("Control1Visible", m_ControlVisible(0), m_def_Control1Visible)
    Call PropBag.WriteProperty("Control2Visible", m_ControlVisible(1), m_def_Control2Visible)
    Call PropBag.WriteProperty("Control1Size", m_Control1Size, m_def_Control1Size)
    Call PropBag.WriteProperty("Control2Size", m_Control2Size, m_def_Control2Size)
End Sub

'Initialize the variables value
Private Sub UserControl_InitProperties()
    m_SeparatorSize = m_def_SeparatorSize
    m_Orientation = m_def_Orientation
    m_InvertControls = m_def_InvertControls
    m_ControlName1 = m_def_ControlName1
    m_ControlName2 = m_def_ControlName2
    m_MinControlSize(0) = m_def_MinControlSize1
    m_MinControlSize(1) = m_def_MinControlSize2
    m_ControlVisible(0) = m_def_Control1Visible
    m_ControlVisible(1) = m_def_Control2Visible
    m_Control1Size = m_def_Control1Size
    m_Control2Size = m_def_Control2Size
End Sub

'MemberInfo=8,0,0,45
Public Property Get SeparatorSize() As Long
    SeparatorSize = m_SeparatorSize
End Property

Public Property Let SeparatorSize(ByVal New_SeparatorSize As Long)
    If New_SeparatorSize >= 15 Then
        m_SeparatorSize = New_SeparatorSize
        UserControl_Show
    End If
    PropertyChanged "SeparatorSize"
End Property

'MemberInfo=14,0,0,1
Public Property Get Orientation() As OrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationConstants)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    UserControl_Show
End Property

'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'MemberInfo=0,0,0,false
Public Property Get InvertControls() As Boolean
    InvertControls = m_InvertControls
End Property

Public Property Let InvertControls(ByVal New_InvertControls As Boolean)
    m_InvertControls = New_InvertControls
    PropertyChanged "InvertControls"
    UserControl_Show
End Property

'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As MSComctlLib.AppearanceConstants
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As MSComctlLib.AppearanceConstants)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Dim c As Control

    UserControl.Enabled() = New_Enabled
    For Each c In UserControl.ContainedControls
        c.Enabled = New_Enabled
    Next
    PropertyChanged "Enabled"
End Property

'MemberInfo=13,0,0,
Public Property Get ControlName1() As String
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(0) Is Nothing Then
        ControlName1 = cnt(0).Name
    End If
End Property

Public Property Let ControlName1(ByVal New_ControlName1 As String)
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(0) Is Nothing Then
        cnt(0).Name = New_ControlName1
    End If
    PropertyChanged "ControlName1"
End Property

'MemberInfo=13,0,0,
Public Property Get ControlName2() As String
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(1) Is Nothing Then
        ControlName2 = cnt(1).Name
    End If
End Property

Public Property Let ControlName2(ByVal New_ControlName2 As String)
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(1) Is Nothing Then
        cnt(1).Name = New_ControlName2
    End If
    PropertyChanged "ControlName2"
End Property

'MemberInfo=8,0,0,0
Public Property Get MinControlSize1() As Long
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    MinControlSize1 = m_MinControlSize(i)
End Property

Public Property Let MinControlSize1(ByVal New_MinControlSize1 As Long)
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    If New_MinControlSize1 >= 90 Then
        m_MinControlSize(i) = New_MinControlSize1
    End If
    PropertyChanged "MinControlSize1"
End Property

'MemberInfo=8,0,0,0
Public Property Get MinControlSize2() As Long
    Dim i As Integer

    If m_InvertControls Then i = 0 Else i = 1
    MinControlSize2 = m_MinControlSize(i)
End Property

Public Property Let MinControlSize2(ByVal New_MinControlSize2 As Long)
    Dim i As Integer

    If m_InvertControls Then i = 0 Else i = 1
    If New_MinControlSize2 >= 90 Then
        m_MinControlSize(i) = New_MinControlSize2
    End If
    PropertyChanged "MinControlSize2"
End Property

'MemberInfo=0,0,0,true
Public Property Get Control1Visible() As Boolean
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    Control1Visible = m_ControlVisible(i)
End Property

Public Property Let Control1Visible(ByVal New_Control1Visible As Boolean)
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    m_ControlVisible(i) = New_Control1Visible
    If Not New_Control1Visible And Not m_ControlVisible(1 - i) Then
        m_ControlVisible(1 - i) = True
        PropertyChanged "Control2Visible"
    End If
    UserControl_Show
    PropertyChanged "Control1Visible"
End Property

'MemberInfo=0,0,0,true
Public Property Get Control2Visible() As Boolean
    Dim i As Integer

    If m_InvertControls Then i = 0 Else i = 1
    Control2Visible = m_ControlVisible(i)
End Property

Public Property Let Control2Visible(ByVal New_Control2Visible As Boolean)
    Dim i As Integer

    If m_InvertControls Then i = 0 Else i = 1
    m_ControlVisible(i) = New_Control2Visible
    If Not New_Control2Visible And Not m_ControlVisible(1 - i) Then
        m_ControlVisible(1 - i) = True
        PropertyChanged "Control1Visible"
    End If
    UserControl_Show
    PropertyChanged "Control2Visible"
End Property

'MemberInfo=8,0,0,0
Public Property Get Control1Size() As Long
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(0) Is Nothing Then
        If m_Orientation = ccOrientationHorizontal Then
            Control1Size = cnt(0).Height
        Else
            Control1Size = cnt(0).Width
        End If
    End If
End Property

Public Property Let Control1Size(ByVal New_Control1Size As Long)
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(0) Is Nothing Then
        If m_Orientation = ccOrientationHorizontal Then
            If New_Control1Size < m_MinControlSize(0) Then
                New_Control1Size = m_MinControlSize(0)
            ElseIf New_Control1Size > Height - m_MinControlSize(0) Then
                New_Control1Size = Height - m_MinControlSize(0)
            End If
            cnt(0).Height = New_Control1Size
            cnt(1).Height = Height - New_Control1Size
        Else
            If New_Control1Size < m_MinControlSize(0) Then
                New_Control1Size = m_MinControlSize(0)
            ElseIf New_Control1Size > Width - m_MinControlSize(0) Then
                New_Control1Size = Width - m_MinControlSize(0)
            End If
            cnt(0).Width = New_Control1Size
            cnt(1).Width = Width - New_Control1Size
        End If
    End If
    UserControl_Show
    PropertyChanged "Control1Size"
End Property

'MemberInfo=8,0,0,0
Public Property Get Control2Size() As Long
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(1) Is Nothing Then
        If m_Orientation = ccOrientationHorizontal Then
            Control2Size = cnt(1).Height
        Else
            Control2Size = cnt(1).Width
        End If
    End If
End Property

Public Property Let Control2Size(ByVal New_Control2Size As Long)
    Dim c As Control
    Dim cnt(0 To 1) As Control
    Dim i As Integer

    If m_InvertControls Then i = 1 Else i = 0
    For Each c In UserControl.ContainedControls
        If Not c Is lblResizer Then
            Set cnt(i) = c
            If m_InvertControls Then i = i - 1 Else i = i + 1
            If i = 2 Or i = -1 Then Exit For
        End If
    Next
    If Not cnt(1) Is Nothing Then
        If m_Orientation = ccOrientationHorizontal Then
            If New_Control2Size < m_MinControlSize(1) Then
                New_Control2Size = m_MinControlSize(1)
            ElseIf New_Control2Size > Height - m_MinControlSize(1) Then
                New_Control2Size = Height - m_MinControlSize(1)
            End If
            cnt(1).Height = New_Control2Size
            cnt(0).Height = Height - New_Control2Size
        Else
            If New_Control2Size < m_MinControlSize(1) Then
                New_Control2Size = m_MinControlSize(1)
            ElseIf New_Control2Size > Width - m_MinControlSize(1) Then
                New_Control2Size = Width - m_MinControlSize(1)
            End If
            cnt(1).Width = New_Control2Size
            cnt(0).Width = Width - New_Control2Size
        End If
    End If
    UserControl_Show
    PropertyChanged "Control2Size"
End Property

