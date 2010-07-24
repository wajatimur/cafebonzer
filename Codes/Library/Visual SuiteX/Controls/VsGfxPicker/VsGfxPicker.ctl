VERSION 5.00
Begin VB.UserControl VsGfxPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ToolboxBitmap   =   "VsGfxPicker.ctx":0000
End
Attribute VB_Name = "VsGfxPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************
'************** Project :  ColorPicker OCX      ***********************
'************** Version :  1.0                  ***********************
'************** Author  :  Abdul Gafoor.GK      ***********************
'************** Date    :  10/October/2000      ***********************
'**********************************************************************
'
'   This is my second ActiveX control.  My first control was
'   Dropdown Calculator, which can be downloaded with source
'   code from either www.a1vbcode.com or www.vbcode.com
'
'   If you like this control, please don't forget to send
'   your comments in 'gafoorgk@yahoo.com'
'
'**********************************************************************

Option Explicit

'API function & constant declarations
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'Module specific variable declarations
Private RClr As RECT
Private RBut As RECT

Private IsInFocus As Boolean
Private IsButDown As Boolean

'Enums
Public Enum cpAppearanceConstants
    Flat
    [3D]
End Enum

'Default Property Values:
Private Const m_def_ShowToolTips = True
Private Const m_def_ShowSysColorButton = True
Private Const m_def_ShowDefault = True
Private Const m_def_ShowCustomColors = True
Private Const m_def_ShowMoreColors = True
Private Const m_def_DefaultCaption = "Default"
Private Const m_def_MoreColorsCaption = "More Colors..."
Private Const m_def_BackColor = &H8000000C
Private Const m_def_Appearance = cpAppearanceConstants.[3D]
Private Const m_def_Color = &HFFFFFF
Private Const m_def_DefaultColor = &HFFFFFF

'Property Variables:
Private m_ShowToolTips As Boolean
Private m_ShowSysColorButton    As Boolean
Private m_ShowDefault           As Boolean
Private m_ShowCustomColors      As Boolean
Private m_ShowMoreColors        As Boolean
Private m_DefaultCaption        As String
Private m_MoreColorsCaption     As String
Private m_BackColor             As OLE_COLOR
Private m_Appearance            As cpAppearanceConstants
Private m_Color                 As OLE_COLOR
Private m_DefaultColor          As OLE_COLOR

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_MemberFlags = "200"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    IsInFocus = True
    Call RedrawControl
End Sub

Private Sub UserControl_Initialize()
    ScaleMode = vbPixels
End Sub

Private Sub UserControl_LostFocus()
    IsInFocus = False
    Call RedrawControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If Button = 1 Then
        If (X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom) Then
            IsButDown = True
            Call RedrawControl
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If IsButDown Then
        If Not ((X >= RBut.Left And X <= RBut.Right) And (Y >= RBut.Top And Y <= RBut.Bottom)) Then
            IsButDown = False
            Call RedrawControl
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    
    If Button = 1 Then
        If IsButDown Then
            IsButDown = False
            Call RedrawControl
        End If
        
        If ((X >= ScaleLeft And X <= ScaleWidth) And (Y >= ScaleTop And Y <= ScaleHeight)) Then
            Call ShowPalette
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    If Height < 285 Then Height = 285
    
    Call RedrawControl
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub RedrawControl()
    Dim Rct As RECT
    Dim Brsh As Long, Clr As Long
    
    Dim lx As Long, ty As Long
    Dim rx As Long, by As Long
    
    lx = ScaleLeft: ty = ScaleTop
    rx = ScaleWidth: by = ScaleHeight
    
    Cls
    
    'Draw background
    Call SetRect(Rct, 0, 0, rx, by)
    Call OleTranslateColor(m_BackColor, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hdc, Rct, Brsh)
    If m_Appearance = [3D] Then
        Call DrawEdge(hdc, Rct, EDGE_SUNKEN, BF_RECT)
    Else
        Call DrawEdge(hdc, Rct, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT Or BF_MONO)
    End If
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
    
    'Draw button
    Dim CurFontName As String
    CurFontName = Font.Name
    Font.Name = "Marlett"
    Call OleTranslateColor(vbButtonFace, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    If m_Appearance = [3D] Then
        If IsButDown Then
            Call SetRect(RBut, rx - 15, 2, rx - 2, by - 2)
            Call FillRect(hdc, RBut, Brsh)
            Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT Or BF_FLAT)
            Call SetRect(Rct, RBut.Left + 2, RBut.Top, RBut.Right, RBut.Bottom)
            Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        Else
            Call SetRect(RBut, rx - 15, 2, rx - 2, by - 2)
            Call FillRect(hdc, RBut, Brsh)
            Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT)
            Call SetRect(Rct, RBut.Left, RBut.Top, RBut.Right, RBut.Bottom - 1)
            Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        End If
    Else
        Call SetRect(RBut, rx - 15, ty, rx, by)
        Call FillRect(hdc, RBut, Brsh)
        Call DrawEdge(hdc, RBut, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT)
        Call SetRect(Rct, RBut.Left + 1, RBut.Top, RBut.Right, RBut.Bottom - 1)
        Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    End If
    Font.Name = CurFontName
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
    
    'Draw Color
    If m_Appearance = [3D] Then
        Call SetRect(RClr, 4, 4, rx - 17, by - 4)
    Else
        Call SetRect(RClr, 3, 3, rx - 17, by - 3)
    End If
    Call OleTranslateColor(m_Color, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FillRect(hdc, RClr, Brsh)
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
    
    'Draw border to the color
    Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
    Brsh = CreateSolidBrush(Clr)
    Call FrameRect(hdc, RClr, Brsh)
    Call DeleteObject(Brsh)
    Call DeleteObject(Clr)
    
    'Draw focus
    If m_Appearance = [3D] Then
        Call SetRect(Rct, 6, 6, rx - 19, by - 6)
    Else
        Call SetRect(Rct, 5, 5, rx - 19, by - 5)
    End If
    If IsInFocus Then Call DrawFocusRect(hdc, Rct)
    
    Refresh
End Sub

Private Sub ShowPalette()
    Dim ClrCtrlPos As RECT
    
    Call GetWindowRect(hwnd, ClrCtrlPos)
    
    DefClr = m_DefaultColor
    CurClr = m_Color
    
    DefCap = m_DefaultCaption
    MorCap = m_MoreColorsCaption
    
    ShwDef = m_ShowDefault
    ShwMor = m_ShowMoreColors
    ShwCus = m_ShowCustomColors
    ShwSys = m_ShowSysColorButton

    Load GfxPickerColorPalette
    With GfxPickerColorPalette
        .Left = ClrCtrlPos.Left * Screen.TwipsPerPixelX
        .Top = ClrCtrlPos.Bottom * Screen.TwipsPerPixelY
        If (.Top + .Height) > Screen.Height Then
            .Top = ClrCtrlPos.Top * Screen.TwipsPerPixelY - .Height
        End If
        
        .Show vbModal
        
        If Not .IsCanceled Then m_Color = .SelectedColor
        Call RedrawControl
    End With
    Unload GfxPickerColorPalette
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_DefaultColor = m_def_DefaultColor
    m_Color = m_def_Color
    m_Appearance = m_def_Appearance
    m_BackColor = m_def_BackColor
    m_ShowDefault = m_def_ShowDefault
    m_ShowCustomColors = m_def_ShowCustomColors
    m_ShowMoreColors = m_def_ShowMoreColors
    m_DefaultCaption = m_def_DefaultCaption
    m_MoreColorsCaption = m_def_MoreColorsCaption
    m_ShowSysColorButton = m_def_ShowSysColorButton
    m_ShowToolTips = m_def_ShowToolTips
    
    Height = 315
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DefaultColor = PropBag.ReadProperty("DefaultColor", m_def_DefaultColor)
    m_Color = PropBag.ReadProperty("Value", m_def_Color)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ShowDefault = PropBag.ReadProperty("ShowDefault", m_def_ShowDefault)
    m_ShowCustomColors = PropBag.ReadProperty("ShowCustomColors", m_def_ShowCustomColors)
    m_ShowMoreColors = PropBag.ReadProperty("ShowMoreColors", m_def_ShowMoreColors)
    m_DefaultCaption = PropBag.ReadProperty("DefaultCaption", m_def_DefaultCaption)
    m_MoreColorsCaption = PropBag.ReadProperty("MoreColorsCaption", m_def_MoreColorsCaption)
    m_ShowSysColorButton = PropBag.ReadProperty("ShowSysColorButton", m_def_ShowSysColorButton)
    m_ShowToolTips = PropBag.ReadProperty("ShowToolTips", m_def_ShowToolTips)
    
    Call RedrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DefaultColor", m_DefaultColor, m_def_DefaultColor)
    Call PropBag.WriteProperty("Value", m_Color, m_def_Color)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ShowDefault", m_ShowDefault, m_def_ShowDefault)
    Call PropBag.WriteProperty("ShowCustomColors", m_ShowCustomColors, m_def_ShowCustomColors)
    Call PropBag.WriteProperty("ShowMoreColors", m_ShowMoreColors, m_def_ShowMoreColors)
    Call PropBag.WriteProperty("DefaultCaption", m_DefaultCaption, m_def_DefaultCaption)
    Call PropBag.WriteProperty("MoreColorsCaption", m_MoreColorsCaption, m_def_MoreColorsCaption)
    Call PropBag.WriteProperty("ShowSysColorButton", m_ShowSysColorButton, m_def_ShowSysColorButton)
    Call PropBag.WriteProperty("ShowToolTips", m_ShowToolTips, m_def_ShowToolTips)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get DefaultColor() As OLE_COLOR
Attribute DefaultColor.VB_Description = "Returns/Sets  the default color"
Attribute DefaultColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DefaultColor = m_DefaultColor
End Property

Public Property Let DefaultColor(ByVal New_DefaultColor As OLE_COLOR)
    m_DefaultColor = New_DefaultColor
    PropertyChanged "DefaultColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get Color() As OLE_COLOR
Attribute Color.VB_Description = "Returns/Sets the selected color"
Attribute Color.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Color.VB_UserMemId = 0
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    m_Color = New_Color
    PropertyChanged "Value"
    
    Call RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,cpAppearanceConstants.[3D]
Public Property Get Appearance() As cpAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As cpAppearanceConstants)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    
    Call RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000C&
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Call RedrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowDefault() As Boolean
Attribute ShowDefault.VB_Description = "Returns/Sets whether default button will be shown or not"
Attribute ShowDefault.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowDefault = m_ShowDefault
End Property

Public Property Let ShowDefault(ByVal New_ShowDefault As Boolean)
    m_ShowDefault = New_ShowDefault
    PropertyChanged "ShowDefault"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowCustomColors() As Boolean
Attribute ShowCustomColors.VB_Description = "Returns/Sets whether custom colors will be shown or not"
Attribute ShowCustomColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowCustomColors = m_ShowCustomColors
End Property

Public Property Let ShowCustomColors(ByVal New_ShowCustomColors As Boolean)
    m_ShowCustomColors = New_ShowCustomColors
    PropertyChanged "ShowCustomColors"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowMoreColors() As Boolean
Attribute ShowMoreColors.VB_Description = "Returns/Sets whether More Colors button will be shown or not"
Attribute ShowMoreColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowMoreColors = m_ShowMoreColors
End Property

Public Property Let ShowMoreColors(ByVal New_ShowMoreColors As Boolean)
    m_ShowMoreColors = New_ShowMoreColors
    PropertyChanged "ShowMoreColors"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Default
Public Property Get DefaultCaption() As String
Attribute DefaultCaption.VB_Description = "Returns/Sets the caption in default button"
Attribute DefaultCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DefaultCaption = m_DefaultCaption
End Property

Public Property Let DefaultCaption(ByVal New_DefaultCaption As String)
    m_DefaultCaption = New_DefaultCaption
    PropertyChanged "DefaultCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,More Colors...
Public Property Get MoreColorsCaption() As String
Attribute MoreColorsCaption.VB_Description = "Returns/Sets the caption in the More button"
Attribute MoreColorsCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MoreColorsCaption = m_MoreColorsCaption
End Property

Public Property Let MoreColorsCaption(ByVal New_MoreColorsCaption As String)
    m_MoreColorsCaption = New_MoreColorsCaption
    PropertyChanged "MoreColorsCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowSysColorButton() As Boolean
Attribute ShowSysColorButton.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowSysColorButton = m_ShowSysColorButton
End Property

Public Property Let ShowSysColorButton(ByVal New_ShowSysColorButton As Boolean)
    m_ShowSysColorButton = New_ShowSysColorButton
    PropertyChanged "ShowSysColorButton"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowToolTips() As Boolean
Attribute ShowToolTips.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ShowToolTips = m_ShowToolTips
End Property

Public Property Let ShowToolTips(ByVal New_ShowToolTips As Boolean)
    m_ShowToolTips = New_ShowToolTips
    PropertyChanged "ShowToolTips"
End Property

