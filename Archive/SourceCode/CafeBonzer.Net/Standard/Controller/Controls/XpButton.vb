Option Strict Off
Option Explicit On
Friend Class XpButton
	Inherits System.Windows.Forms.UserControl
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		UserControl_Initialize()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			UserControl_Terminate()
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents OverTimer As System.Windows.Forms.Timer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(XpButton))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.OverTimer = New System.Windows.Forms.Timer(components)
		Me.ClientSize = New System.Drawing.Size(320, 240)
		MyBase.Location = New System.Drawing.Point(0, 0)
		MyBase.Name = "XpButton"
		Me.OverTimer.Enabled = False
		Me.OverTimer.Interval = 3
	End Sub
#End Region 
	Public Event COLTYPEChange()
	Public Event PICPOSChange()
	Public Event PICOChange()
	Public Event MPTRChange()
	Public Event SOFTChange()
	Public Event FONTChange()
	Public Event PICNChange()
	Public Event FCOLOChange()
	Public Event TXChange()
	Public Event BCOLOChange()
	Public Event VALUEChange()
	Public Event FXChange()
	Public Event NGREYChange()
	Public Event CHECKChange()
	Public Event HANDChange()
	Public Event MCOLChange()
	Public Event UMCOLChange()
	Public Event FCOLChange()
	Public Event FOCUSRChange()
	Public Event MICONChange()
	Public Event BCOLChange()
	Public Event ENABChange()
	Private Declare Function SetPixel Lib "gdi32"  Alias "SetPixelV"(ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Integer) As Integer
	
	Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	'Private Const COLOR_HIGHLIGHT = 13
	Private Const COLOR_BTNFACE As Short = 15
	Private Const COLOR_BTNSHADOW As Short = 16
	Private Const COLOR_BTNTEXT As Short = 18
	Private Const COLOR_BTNHIGHLIGHT As Short = 20
	Private Const COLOR_BTNDKSHADOW As Short = 21
	Private Const COLOR_BTNLIGHT As Short = 22
	
	Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Integer, ByVal lHPalette As Integer, ByRef lColorRef As Integer) As Integer
	Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Integer, ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function DrawText Lib "user32"  Alias "DrawTextA"(ByVal hdc As Integer, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	Private Const DT_CALCRECT As Short = &H400s
	Private Const DT_WORDBREAK As Short = &H10s
	Private Const DT_CENTER As Boolean = &H1s Or DT_WORDBREAK Or &H4s
	
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function FillRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function FrameRect Lib "user32" (ByVal hdc As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	'Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
	'Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
	
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByRef lpPoint As POINTAPI) As Integer
	Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
	Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Integer) As Integer
	Private Const PS_SOLID As Short = 0
	
	Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
	'Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
	Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Integer, ByVal hSrcRgn1 As Integer, ByVal hSrcRgn2 As Integer, ByVal nCombineMode As Integer) As Integer
	Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Integer, ByVal hRgn As Integer, ByVal bRedraw As Integer) As Integer
	Private Const RGN_DIFF As Short = 4
	
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function InflateRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Integer, ByVal y As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function OffsetRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Integer, ByVal y As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function CopyRect Lib "user32" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Integer
	
	Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Integer, ByVal yPoint As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	
	Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Integer) As Integer
	
	Private Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
	Private Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
	
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Integer, ByVal hBitmap As Integer, ByVal nStartScan As Integer, ByVal nNumScans As Integer, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Integer) As Integer
	'UPGRADE_WARNING: Structure BITMAPINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1016"'
	Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal Scan As Integer, ByVal NumScans As Integer, ByRef Bits As Any, ByRef BitsInfo As BITMAPINFO, ByVal wUsage As Integer) As Integer
	
	Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Integer, ByVal xLeft As Integer, ByVal yTop As Integer, ByVal hIcon As Integer, ByVal cxWidth As Integer, ByVal cyWidth As Integer, ByVal istepIfAniCur As Integer, ByVal hbrFlickerFreeDraw As Integer, ByVal diFlags As Integer) As Integer
	'Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
	'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
	
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Left_Renamed As Integer
		'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Top_Renamed As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Right_Renamed As Integer
		'UPGRADE_NOTE: Bottom was upgraded to Bottom_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
		Dim Bottom_Renamed As Integer
	End Structure
	
	Private Structure POINTAPI
		Dim x As Integer
		Dim y As Integer
	End Structure
	
	Private Structure BITMAPINFOHEADER
		Dim biSize As Integer
		Dim biWidth As Integer
		Dim biHeight As Integer
		Dim biPlanes As Short
		Dim biBitCount As Short
		Dim biCompression As Integer
		Dim biSizeImage As Integer
		Dim biXPelsPerMeter As Integer
		Dim biYPelsPerMeter As Integer
		Dim biClrUsed As Integer
		Dim biClrImportant As Integer
	End Structure
	
	Private Structure RGBTRIPLE
		Dim rgbBlue As Byte
		Dim rgbGreen As Byte
		Dim rgbRed As Byte
	End Structure
	
	Private Structure BITMAPINFO
		Dim bmiHeader As BITMAPINFOHEADER
		Dim bmiColors As RGBTRIPLE
	End Structure
	
	Public Enum ColorTypes
		Use_Windows = 1
		Custom = 2
		Force_Standard = 3
		Use_Container = 4
	End Enum
	
	Public Enum PicPositions
		cbLeft = 0
		cbRight = 1
		cbTop = 2
		cbBottom = 3
		cbBackground = 4
	End Enum
	
	Public Enum fx
		cbNone = 0
		cbEmbossed = 1
		cbEngraved = 2
		cbShadowed = 3
	End Enum
	
	Private Const FXDEPTH As Integer = &H28s
	
	'events
	Public Shadows Event Click(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	<System.Runtime.InteropServices.ProgId("MouseDownEventArgs_NET.MouseDownEventArgs")> Public NotInheritable Class MouseDownEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Public Shadows Event MouseDown(ByVal Sender As System.Object, ByVal e As MouseDownEventArgs)
	<System.Runtime.InteropServices.ProgId("MouseMoveEventArgs_NET.MouseMoveEventArgs")> Public NotInheritable Class MouseMoveEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Public Shadows Event MouseMove(ByVal Sender As System.Object, ByVal e As MouseMoveEventArgs)
	<System.Runtime.InteropServices.ProgId("MouseUpEventArgs_NET.MouseUpEventArgs")> Public NotInheritable Class MouseUpEventArgs
		Inherits System.EventArgs
		Public Button As Short
		Public Shift As Short
		Public x As Single
		Public y As Single
		Public Sub New(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
			MyBase.New()
			Me.Button = Button
			Me.Shift = Shift
			Me.x = x
			Me.y = y
		End Sub
	End Class
	Public Shadows Event MouseUp(ByVal Sender As System.Object, ByVal e As MouseUpEventArgs)
	<System.Runtime.InteropServices.ProgId("KeyPressEventArgs_NET.KeyPressEventArgs")> Public NotInheritable Class KeyPressEventArgs
		Inherits System.EventArgs
		Public KeyAscii As Short
		Public Sub New(ByRef KeyAscii As Short)
			MyBase.New()
			Me.KeyAscii = KeyAscii
		End Sub
	End Class
	Public Shadows Event KeyPress(ByVal Sender As System.Object, ByVal e As KeyPressEventArgs)
	<System.Runtime.InteropServices.ProgId("KeyDownEventArgs_NET.KeyDownEventArgs")> Public NotInheritable Class KeyDownEventArgs
		Inherits System.EventArgs
		Public KeyCode As Short
		Public Shift As Short
		Public Sub New(ByRef KeyCode As Short, ByRef Shift As Short)
			MyBase.New()
			Me.KeyCode = KeyCode
			Me.Shift = Shift
		End Sub
	End Class
	Public Shadows Event KeyDown(ByVal Sender As System.Object, ByVal e As KeyDownEventArgs)
	<System.Runtime.InteropServices.ProgId("KeyUpEventArgs_NET.KeyUpEventArgs")> Public NotInheritable Class KeyUpEventArgs
		Inherits System.EventArgs
		Public KeyCode As Short
		Public Shift As Short
		Public Sub New(ByRef KeyCode As Short, ByRef Shift As Short)
			MyBase.New()
			Me.KeyCode = KeyCode
			Me.Shift = Shift
		End Sub
	End Class
	Public Shadows Event KeyUp(ByVal Sender As System.Object, ByVal e As KeyUpEventArgs)
	Public Event MouseOver(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	Public Event MouseOut(ByVal Sender As System.Object, ByVal e As System.EventArgs)
	
	'variables
	Private MyColorType As ColorTypes
	Private PicPosition As PicPositions
	Private SFX As fx 'font and picture effects
	
	Private He As Integer 'the height of the button
	Private Wi As Integer 'the width of the button
	
	Private BackC As Integer 'back color
	Private BackO As Integer 'back color when mouse is over
	Private ForeC As Integer 'fore color
	Private ForeO As Integer 'fore color when mouse is over
	Private MaskC As Integer 'mask color
	Private useMask, useGrey As Boolean
	Private useHand As Boolean
	
	Private picNormal, picHover As System.Drawing.Image
	Private pBM, pDC, oBM As Integer 'used for the treansparent button
	
	Private elTex As String 'current text
	
	Private rc2, rc, rc3 As RECT
	Private fc As POINTAPI 'text and focus rect locations
	Private picPT, picSZ As POINTAPI 'picture Position & Size
	Private rgnNorm As Integer
	
	Private LastButton, LastKeyDown As Byte
	Private isEnabled, isSoft As Boolean
	Private HasFocus, showFocusR As Boolean
	
	Private cMask, cTextO, cDarkShadow, cHighLight, cFace, cLight, cShadow, cText, cFaceO, XPFace As Integer
	
	Private lastStat As Byte
	Private TE As String
	Private isShown As Boolean 'used to avoid unnecessary repaints
	Private isOver, inLoop As Boolean
	
	'Private Locked As Boolean
	
	Private captOpt As Integer
	Private isCheckbox, cValue As Boolean
	
	Private Sub OverTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OverTimer.Tick
		If Not isMouseOver Then
			OverTimer.Enabled = False
			isOver = False
			Call Redraw(0, True)
			RaiseEvent MouseOut(Me, Nothing)
		End If
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.AccessKeyPress was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_AccessKeyPress(ByRef KeyAscii As Short)
		LastButton = 1
		Call XpButton_Click(Me, New System.EventArgs())
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.AmbientChanged was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_AmbientChanged(ByRef PropertyName As String)
		If Not MyColorType = ColorTypes.Custom Then
			Call SetColors()
			Call Redraw(lastStat, True)
		End If
	End Sub
	
	Private Sub XpButton_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
		If LastButton = 1 And isEnabled Then
			If isCheckbox Then cValue = Not cValue
			Call Redraw(0, True) 'be sure that the normal status is drawn
			MyBase.Refresh()
			RaiseEvent Click(Me, Nothing)
		End If
	End Sub
	
	Private Sub XpButton_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
		If LastButton = 1 Then
			Call XpButton_MouseDown(Me, New System.Windows.Forms.MouseEventArgs(1 * &H100000, 0, 0, 0, 0))
			SetCapture(hwnd)
		End If
	End Sub
	
	Private Sub XpButton_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		HasFocus = True
		Call Redraw(lastStat, True)
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.Hide was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_Hide()
		isShown = False
	End Sub
	
	Private Sub UserControl_Initialize()
		'this makes the control to be slow, remark this line if the "not redrawing" problem is not important for you: ie, you intercept the Load_Event (with breakpoint or messageBox) and the button does not repaint...
		isShown = True
	End Sub
	
	Private Sub XpButton_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		RaiseEvent KeyDown(Me, New KeyDownEventArgs(KeyCode, Shift))
		
		LastKeyDown = KeyCode
		Select Case KeyCode
			Case 32 'spacebar pressed
				Call Redraw(2, False)
			Case 39, 40 'right and down arrows
				System.Windows.Forms.SendKeys.Send("{Tab}")
			Case 37, 38 'left and up arrows
				System.Windows.Forms.SendKeys.Send("+{Tab}")
		End Select
	End Sub
	
	Private Sub XpButton_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		RaiseEvent KeyPress(Me, New KeyPressEventArgs(KeyAscii))
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub XpButton_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		RaiseEvent KeyUp(Me, New KeyUpEventArgs(KeyCode, Shift))
		
		If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed, and not cancelled by the user
			If isCheckbox Then cValue = Not cValue
			Call Redraw(0, False)
			MyBase.Refresh()
			RaiseEvent Click(Me, Nothing)
		End If
	End Sub
	
	Private Sub XpButton_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus
		HasFocus = False
		Call Redraw(lastStat, True)
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.InitProperties was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_InitProperties()
		isEnabled = True : showFocusR = True : useMask = True
		'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayName was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		elTex = Ambient.DisplayName
		'UPGRADE_ISSUE: AmbientProperties property Ambient.Font was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		MyBase.Font = Ambient.Font
		MyColorType = ColorTypes.Use_Windows
		Call SetColors()
		BackC = cFace : BackO = BackC
		ForeC = cText : ForeO = ForeC
		MaskC = &HC0C0C0
		Call CalcTextRects()
	End Sub
	
	Private Sub XpButton_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseDown(Me, New MouseDownEventArgs(Button, Shift, x, y))
		LastButton = Button
		If Button <> 2 Then Call Redraw(2, False)
	End Sub
	
	Private Sub XpButton_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseMove(Me, New MouseMoveEventArgs(Button, Shift, x, y))
		If Button < 2 Then
			If Not isMouseOver Then
				'we are outside the button
				Call Redraw(0, False)
			Else
				'we are inside the button
				If Button = 0 And Not isOver Then
					OverTimer.Enabled = True
					isOver = True
					Call Redraw(0, True)
					RaiseEvent MouseOver(Me, Nothing)
				ElseIf Button = 1 Then 
					isOver = True
					Call Redraw(2, False)
					isOver = False
				End If
			End If
		End If
	End Sub
	
	Private Sub XpButton_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		RaiseEvent MouseUp(Me, New MouseUpEventArgs(Button, Shift, x, y))
		If Button <> 2 Then Call Redraw(0, False)
	End Sub
	
	'########## BUTTON PROPERTIES ##########
	Public Overrides Property BackColor() As System.Drawing.Color
		Get
			Return System.Drawing.ColorTranslator.FromOle(BackC)
		End Get
		Set(ByVal Value As System.Drawing.Color)
			BackC = System.Drawing.ColorTranslator.ToOle(Value)
			If DesignMode Then BackO = System.Drawing.ColorTranslator.ToOle(Value)
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent BCOLChange()
		End Set
	End Property
	
	Public Property BackOver() As System.Drawing.Color
		Get
			BackOver = System.Drawing.ColorTranslator.FromOle(BackO)
		End Get
		Set(ByVal Value As System.Drawing.Color)
			BackO = System.Drawing.ColorTranslator.ToOle(Value)
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent BCOLOChange()
		End Set
	End Property
	
	Public Overrides Property ForeColor() As System.Drawing.Color
		Get
			Return System.Drawing.ColorTranslator.FromOle(ForeC)
		End Get
		Set(ByVal Value As System.Drawing.Color)
			ForeC = System.Drawing.ColorTranslator.ToOle(Value)
			If DesignMode Then ForeO = System.Drawing.ColorTranslator.ToOle(Value)
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent FCOLChange()
		End Set
	End Property
	
	Public Property ForeOver() As System.Drawing.Color
		Get
			ForeOver = System.Drawing.ColorTranslator.FromOle(ForeO)
		End Get
		Set(ByVal Value As System.Drawing.Color)
			ForeO = System.Drawing.ColorTranslator.ToOle(Value)
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent FCOLOChange()
		End Set
	End Property
	
	Public Property MaskColor() As System.Drawing.Color
		Get
			MaskColor = System.Drawing.ColorTranslator.FromOle(MaskC)
		End Get
		Set(ByVal Value As System.Drawing.Color)
			MaskC = System.Drawing.ColorTranslator.ToOle(Value)
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent MCOLChange()
		End Set
	End Property
	
	Public Property Caption() As String
		Get
			Caption = elTex
		End Get
		Set(ByVal Value As String)
			elTex = Value
			Call SetAccessKeys()
			Call CalcTextRects()
			Call Redraw(0, True)
			RaiseEvent TXChange()
		End Set
	End Property
	
	Public Shadows Property Enabled() As Boolean
		Get
			Return isEnabled
		End Get
		Set(ByVal Value As Boolean)
			isEnabled = Value
			Call Redraw(0, True)
			MyBase.Enabled = isEnabled
			RaiseEvent ENABChange()
		End Set
	End Property
	
	Public Overrides Property Font() As System.Drawing.Font
		Get
			Font = MyBase.Font
		End Get
		Set(ByVal Value As System.Drawing.Font)
			MyBase.Font = Value
			Call CalcTextRects()
			Call Redraw(0, True)
			RaiseEvent FONTChange()
		End Set
	End Property
	
	Public Property FontBold() As Boolean
		Get
			FontBold = MyBase.Font.Bold
		End Get
		Set(ByVal Value As Boolean)
			MyBase.Font = VB6.FontChangeBold(MyBase.Font, Value)
			Call CalcTextRects()
			Call Redraw(0, True)
		End Set
	End Property
	
	Public Property FontItalic() As Boolean
		Get
			FontItalic = MyBase.Font.Italic
		End Get
		Set(ByVal Value As Boolean)
			MyBase.Font = VB6.FontChangeItalic(MyBase.Font, Value)
			Call CalcTextRects()
			Call Redraw(0, True)
		End Set
	End Property
	
	Public Property FontUnderline() As Boolean
		Get
			FontUnderline = MyBase.Font.Underline
		End Get
		Set(ByVal Value As Boolean)
			MyBase.Font = VB6.FontChangeUnderline(MyBase.Font, Value)
			Call CalcTextRects()
			Call Redraw(0, True)
		End Set
	End Property
	
	Public Property FontSize() As Short
		Get
			FontSize = MyBase.Font.SizeInPoints
		End Get
		Set(ByVal Value As Short)
			MyBase.Font = VB6.FontChangeSize(MyBase.Font, Value)
			Call CalcTextRects()
			Call Redraw(0, True)
		End Set
	End Property
	
	Public Property FontName() As String
		Get
			FontName = MyBase.Font.Name
		End Get
		Set(ByVal Value As String)
			MyBase.Font = VB6.FontChangeName(MyBase.Font, Value)
			Call CalcTextRects()
			Call Redraw(0, True)
		End Set
	End Property
	
	'it is very common that a windows user uses custom color
	'schemes to view his/her desktop, and is also very
	'common that this color scheme has weird colors that
	'would alter the nice look of my buttons.
	'So if you want to force the button to use the windows
	'standard colors you may change this property to "Force Standard"
	
	
	Public Property ColorScheme() As ColorTypes
		Get
			ColorScheme = MyColorType
		End Get
		Set(ByVal Value As ColorTypes)
			MyColorType = Value
			Call SetColors()
			Call Redraw(0, True)
			RaiseEvent COLTYPEChange()
		End Set
	End Property
	
	
	Public Property ShowFocusRect() As Boolean
		Get
			ShowFocusRect = showFocusR
		End Get
		Set(ByVal Value As Boolean)
			showFocusR = Value
			Call Redraw(lastStat, True)
			RaiseEvent FOCUSRChange()
		End Set
	End Property
	
	
	Public Property MousePointer() As System.Windows.Forms.Cursor
		Get
			MousePointer = MyBase.Cursor
		End Get
		Set(ByVal Value As System.Windows.Forms.Cursor)
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2036"'
			MyBase.Cursor = Value
			RaiseEvent MPTRChange()
		End Set
	End Property
	
	
	Public Property MouseIcon() As System.Drawing.Image
		Get
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			MouseIcon = MyBase.MouseIcon
		End Get
		Set(ByVal Value As System.Drawing.Image)
			On Error Resume Next
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			MyBase.MouseIcon = Value
			RaiseEvent MICONChange()
		End Set
	End Property
	
	Public Property HandPointer() As Boolean
		Get
			HandPointer = useHand
		End Get
		Set(ByVal Value As Boolean)
			useHand = Value
			If useHand Then
				'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				MyBase.MouseIcon = VB6.LoadResPicture(101, 2)
				'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2036"'
				MyBase.Cursor = vbCustom
			Else
				'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				'UPGRADE_NOTE: Object UserControl.MouseIcon may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
				MyBase.MouseIcon = Nothing
				MyBase.Cursor = System.Windows.Forms.Cursors.Arrow
			End If
			RaiseEvent HANDChange()
		End Set
	End Property
	
	Public ReadOnly Property hwnd() As Integer
		Get
			hwnd = MyBase.Handle.ToInt32
		End Get
	End Property
	
	
	Public Property SoftBevel() As Boolean
		Get
			SoftBevel = isSoft
		End Get
		Set(ByVal Value As Boolean)
			isSoft = Value
			Call SetColors()
			Call Redraw(lastStat, True)
			RaiseEvent SOFTChange()
		End Set
	End Property
	
	Public Property PictureNormal() As System.Drawing.Image
		Get
			PictureNormal = picNormal
		End Get
		Set(ByVal Value As System.Drawing.Image)
			picNormal = Value
			Call CalcPicSize()
			Call CalcTextRects()
			Call Redraw(lastStat, True)
			RaiseEvent PICNChange()
		End Set
	End Property
	
	Public Property PictureOver() As System.Drawing.Image
		Get
			PictureOver = picHover
		End Get
		Set(ByVal Value As System.Drawing.Image)
			picHover = Value
			If isOver Then Call Redraw(lastStat, True) 'only redraw i we need to see this picture immediately
			RaiseEvent PICOChange()
		End Set
	End Property
	
	Public Property PicturePosition() As PicPositions
		Get
			PicturePosition = PicPosition
		End Get
		Set(ByVal Value As PicPositions)
			PicPosition = Value
			RaiseEvent PICPOSChange()
			Call CalcTextRects()
			Call Redraw(lastStat, True)
		End Set
	End Property
	
	
	Public Property UseMaskColor() As Boolean
		Get
			UseMaskColor = useMask
		End Get
		Set(ByVal Value As Boolean)
			useMask = Value
			If Not picNormal Is Nothing Then Call Redraw(lastStat, True)
			RaiseEvent UMCOLChange()
		End Set
	End Property
	
	Public Property UseGreyscale() As Boolean
		Get
			UseGreyscale = useGrey
		End Get
		Set(ByVal Value As Boolean)
			useGrey = Value
			If Not picNormal Is Nothing Then Call Redraw(lastStat, True)
			RaiseEvent NGREYChange()
		End Set
	End Property
	
	
	Public Property SpecialEffect() As fx
		Get
			SpecialEffect = SFX
		End Get
		Set(ByVal Value As fx)
			SFX = Value
			Call Redraw(lastStat, True)
			RaiseEvent FXChange()
		End Set
	End Property
	
	
	Public Property CheckBoxBehaviour() As Boolean
		Get
			CheckBoxBehaviour = isCheckbox
		End Get
		Set(ByVal Value As Boolean)
			isCheckbox = Value
			Call Redraw(lastStat, True)
			RaiseEvent CHECKChange()
		End Set
	End Property
	
	
	Public Property Value() As Boolean
		Get
			Value = cValue
		End Get
		Set(ByVal Value As Boolean)
			cValue = Value
			If isCheckbox Then Call Redraw(0, True)
			RaiseEvent VALUEChange()
		End Set
	End Property
	'########## END OF PROPERTIES ##########
	
	
	Private Sub XpButton_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If inLoop Then Exit Sub
		'get button size
		GetClientRect(MyBase.Handle.ToInt32, rc3)
		'assign these values to He and Wi
		He = rc3.Bottom_Renamed : Wi = rc3.Right_Renamed
		InflateRect(rc3, -4, -4)
		Call CalcTextRects()
		
		If rgnNorm Then DeleteObject(rgnNorm)
		Call MakeRegion()
		SetWindowRgn(MyBase.Handle.ToInt32, rgnNorm, True)
		
		If He Then Call Redraw(0, True)
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event ReadProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_ReadProperties(ByRef PropBag As PropertyBag)
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			elTex = .ReadProperty("TX", "")
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			isEnabled = .ReadProperty("ENAB", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			MyBase.Font = .ReadProperty("FONT", MyBase.Font)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			MyColorType = .ReadProperty("COLTYPE", 1)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			showFocusR = .ReadProperty("FOCUSR", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			BackO = .ReadProperty("BCOLO", BackC)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ForeO = .ReadProperty("FCOLO", ForeC)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			MaskC = .ReadProperty("MCOL", &HC0C0C0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_ISSUE: UserControl property UserControl.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2036"'
			MyBase.Cursor = .ReadProperty("MPTR", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			MyBase.MouseIcon = .ReadProperty("MICON", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			picNormal = .ReadProperty("PICN", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			picHover = .ReadProperty("PICH", Nothing)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			useMask = .ReadProperty("UMCOL", True)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			isSoft = .ReadProperty("SOFT", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			PicPosition = .ReadProperty("PICPOS", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			useGrey = .ReadProperty("NGREY", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			SFX = .ReadProperty("FX", 0)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			Me.HandPointer = .ReadProperty("HAND", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			isCheckbox = .ReadProperty("CHECK", False)
			'UPGRADE_ISSUE: PropertyBag method PropBag.ReadProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_WARNING: Couldn't resolve default property of object PropBag.ReadProperty(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			cValue = .ReadProperty("VALUE", False)
		End With
		
		MyBase.Enabled = isEnabled
		Call CalcPicSize()
		Call CalcTextRects()
		Call SetAccessKeys()
	End Sub
	
	'UPGRADE_WARNING: UserControl Event UserControl.Show was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub UserControl_Show()
		isShown = True
		Call SetColors()
		Call Redraw(0, True)
	End Sub
	
	Private Sub UserControl_Terminate()
		isShown = False
		DeleteObject(rgnNorm)
		If pDC Then
			DeleteObject(SelectObject(pDC, oBM))
			DeleteDC(pDC)
		End If
	End Sub
	
	'UPGRADE_WARNING: PropertyBag object was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6003"'
	'UPGRADE_WARNING: UserControl Event WriteProperties is not supported. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup6002"'
	Private Sub UserControl_WriteProperties(ByRef PropBag As PropertyBag)
		With PropBag
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("TX", elTex)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("ENAB", isEnabled)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("FONT", MyBase.Font)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("COLTYPE", MyColorType)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("FOCUSR", showFocusR)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("BCOL", BackC)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("BCOLO", BackO)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("FCOL", ForeC)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("FCOLO", ForeO)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("MCOL", MaskC)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("MPTR", MyBase.Cursor)
			'UPGRADE_ISSUE: UserControl property UserControl.MouseIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("MICON", MyBase.MouseIcon)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("PICN", picNormal)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("PICH", picHover)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("UMCOL", useMask)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("SOFT", isSoft)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("PICPOS", PicPosition)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("NGREY", useGrey)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("FX", SFX)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("HAND", useHand)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("CHECK", isCheckbox)
			'UPGRADE_ISSUE: PropertyBag method PropBag.WriteProperty was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			Call .WriteProperty("VALUE", cValue)
		End With
	End Sub
	
	Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
		'here is the CORE of the button, everything is drawn here
		'it's not well commented but i think that everything is
		'pretty self explanatory...
		
		If isCheckbox And cValue Then curStat = 2
		If Not Force Then 'check drawing redundancy
			If (curStat = lastStat) And (TE = elTex) Then Exit Sub
		End If
		
		If He = 0 Or Not isShown Then Exit Sub 'we don't want errors
		lastStat = curStat
		TE = elTex
		Dim XPFace2, i, tempCol As Integer
		Dim stepXP1 As Single
		
		With MyBase
			'UPGRADE_ISSUE: UserControl method UserControl.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			.Cls()
			If isOver And MyColorType = ColorTypes.Custom Then tempCol = BackC : BackC = BackO : SetColors()
			DrawRectangle(0, 0, Wi, He, cFace)
			If isEnabled Then
				If curStat = 0 Then
					'#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
					stepXP1 = 25 / He
					For i = 1 To He
						DrawLine(0, i, Wi, i, ShiftColor(XPFace, -stepXP1 * i))
					Next 
					Call DrawCaption(System.Math.Abs(CInt(isOver)))
					DrawRectangle(0, 0, Wi, He, &H733C00, True)
					mSetPixel(1, 1, &H7B4D10)
					mSetPixel(1, He - 2, &H7B4D10)
					mSetPixel(Wi - 2, 1, &H7B4D10)
					mSetPixel(Wi - 2, He - 2, &H7B4D10)
					
					If isOver Then
						DrawRectangle(1, 2, Wi - 2, He - 4, &H31B2FF, True)
						DrawLine(2, He - 2, Wi - 2, He - 2, &H96E7)
						DrawLine(2, 1, Wi - 2, 1, &HCEF3FF)
						DrawLine(1, 2, Wi - 1, 2, &H8CDBFF)
						DrawLine(2, 3, 2, He - 3, &H6BCBFF)
						DrawLine(Wi - 3, 3, Wi - 3, He - 3, &H6BCBFF)
						'UPGRADE_ISSUE: AmbientProperties property Ambient.DisplayAsDefault was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					ElseIf ((HasFocus Or Ambient.DisplayAsDefault) And showFocusR) Then 
						DrawRectangle(1, 2, Wi - 2, He - 4, &HE7AE8C, True)
						DrawLine(2, He - 2, Wi - 2, He - 2, &HEF826B)
						DrawLine(2, 1, Wi - 2, 1, &HFFE7CE)
						DrawLine(1, 2, Wi - 1, 2, &HF7D7BD)
						DrawLine(2, 3, 2, He - 3, &HF0D1B5)
						DrawLine(Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5)
					Else 'we do not draw the bevel always because the above code would repaint over it
						DrawLine(2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H30s))
						DrawLine(1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace, -&H20s))
						DrawLine(Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H24s))
						DrawLine(Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPFace, -&H18s))
						DrawLine(2, 1, Wi - 2, 1, ShiftColor(XPFace, &H10s))
						DrawLine(1, 2, Wi - 2, 2, ShiftColor(XPFace, &HAs))
						DrawLine(1, 2, 1, He - 2, ShiftColor(XPFace, -&H5s))
						DrawLine(2, 3, 2, He - 3, ShiftColor(XPFace, -&HAs))
					End If
					Call DrawPictures(0)
				ElseIf curStat = 2 Then 
					'#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
					stepXP1 = 25 / He
					XPFace2 = ShiftColor(XPFace, -32)
					For i = 1 To He
						DrawLine(0, He - i, Wi, He - i, ShiftColor(XPFace2, -stepXP1 * i))
					Next 
					Call DrawCaption(2)
					DrawRectangle(0, 0, Wi, He, &H733C00, True)
					mSetPixel(1, 1, &H7B4D10)
					mSetPixel(1, He - 2, &H7B4D10)
					mSetPixel(Wi - 2, 1, &H7B4D10)
					mSetPixel(Wi - 2, He - 2, &H7B4D10)
					
					DrawLine(2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H10s))
					DrawLine(1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace2, &HAs))
					DrawLine(Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H5s))
					DrawLine(Wi - 3, 3, Wi - 3, He - 3, XPFace)
					DrawLine(2, 1, Wi - 2, 1, ShiftColor(XPFace2, -&H20s))
					DrawLine(1, 2, Wi - 2, 2, ShiftColor(XPFace2, -&H18s))
					DrawLine(1, 2, 1, He - 2, ShiftColor(XPFace2, -&H20s))
					DrawLine(2, 2, 2, He - 2, ShiftColor(XPFace2, -&H16s))
					Call DrawPictures(1)
				End If
			Else
				'#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
				DrawRectangle(0, 0, Wi, He, ShiftColor(XPFace, -&H18s))
				Call DrawCaption(5)
				DrawRectangle(0, 0, Wi, He, ShiftColor(XPFace, -&H54s), True)
				mSetPixel(1, 1, ShiftColor(XPFace, -&H48s))
				mSetPixel(1, He - 2, ShiftColor(XPFace, -&H48s))
				mSetPixel(Wi - 2, 1, ShiftColor(XPFace, -&H48s))
				mSetPixel(Wi - 2, He - 2, ShiftColor(XPFace, -&H48s))
				Call DrawPictures(2)
			End If
		End With
		If isOver And MyColorType = ColorTypes.Custom Then BackC = tempCol : SetColors()
	End Sub
	
	Private Sub DrawRectangle(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color As Integer, Optional ByRef OnlyBorder As Boolean = False)
		'this is my custom function to draw rectangles and frames
		'it's faster and smoother than using the line method
		Dim bRECT As RECT
		Dim hBrush As Integer
		
		bRECT.Left_Renamed = x
		bRECT.Top_Renamed = y
		bRECT.Right_Renamed = x + Width
		bRECT.Bottom_Renamed = y + Height
		hBrush = CreateSolidBrush(Color)
		
		If OnlyBorder Then
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			FrameRect(MyBase.hdc, bRECT, hBrush)
		Else
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			FillRect(MyBase.hdc, bRECT, hBrush)
		End If
		
		DeleteObject(hBrush)
	End Sub
	
	Private Sub DrawLine(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal Color As Integer)
		'a fast way to draw lines
		Dim pt As POINTAPI
		Dim oldPen, hPen As Integer
		With MyBase
			hPen = CreatePen(PS_SOLID, 1, Color)
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			oldPen = SelectObject(.hdc, hPen)
			
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			MoveToEx(.hdc, X1, Y1, pt)
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			LineTo(.hdc, X2, Y2)
			
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			SelectObject(.hdc, oldPen)
			DeleteObject(hPen)
		End With
	End Sub
	
	Private Sub mSetPixel(ByVal x As Integer, ByVal y As Integer, ByVal Color As Integer)
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Call SetPixel(MyBase.hdc, x, y, Color)
	End Sub
	
	Private Sub SetColors()
		'this function sets the colors taken as a base to build
		'all the other colors and styles.
		
		If MyColorType = ColorTypes.Custom Then
			cFace = ConvertFromSystemColor(BackC)
			cFaceO = ConvertFromSystemColor(BackO)
			cText = ConvertFromSystemColor(ForeC)
			cTextO = ConvertFromSystemColor(ForeO)
			cShadow = ShiftColor(cFace, -&H40s)
			cLight = ShiftColor(cFace, &H1Fs)
			cHighLight = ShiftColor(cFace, &H2Fs) 'it should be 3F but it looks too lighter
			cDarkShadow = ShiftColor(cFace, -&HC0s)
		ElseIf MyColorType = ColorTypes.Force_Standard Then 
			cFace = &HC0C0C0
			cFaceO = cFace
			cShadow = &H808080
			cLight = &HDFDFDF
			cDarkShadow = &H0s
			cHighLight = &HFFFFFF
			cText = &H0s
			cTextO = cText
		ElseIf MyColorType = ColorTypes.Use_Container Then 
			cFace = GetBkColor(GetDC(GetParent(hwnd)))
			cFaceO = cFace
			cText = GetTextColor(GetDC(GetParent(hwnd)))
			cTextO = cText
			cShadow = ShiftColor(cFace, -&H40s)
			cLight = ShiftColor(cFace, &H1Fs)
			cHighLight = ShiftColor(cFace, &H2Fs)
			cDarkShadow = ShiftColor(cFace, -&HC0s)
		Else
			'if MyColorType is 1 or has not been set then use windows colors
			cFace = GetSysColor(COLOR_BTNFACE)
			cFaceO = cFace
			cShadow = GetSysColor(COLOR_BTNSHADOW)
			cLight = GetSysColor(COLOR_BTNLIGHT)
			cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
			cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
			cText = GetSysColor(COLOR_BTNTEXT)
			cTextO = cText
		End If
		cMask = ConvertFromSystemColor(MaskC)
		XPFace = ShiftColor(cFace, &H30s)
	End Sub
	
	Private Sub MakeRegion()
		'this function creates the regions to "cut" the UserControl
		'so it will be transparent in certain areas
		Dim rgn1, rgn2 As Integer
		
		DeleteObject(rgnNorm)
		rgnNorm = CreateRectRgn(0, 0, Wi, He)
		rgn2 = CreateRectRgn(0, 0, 0, 0)
		
		rgn1 = CreateRectRgn(0, 0, 2, 1)
		CombineRgn(rgn2, rgnNorm, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(0, He, 2, He - 1)
		CombineRgn(rgnNorm, rgn2, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
		CombineRgn(rgn2, rgnNorm, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
		CombineRgn(rgnNorm, rgn2, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(0, 1, 1, 2)
		CombineRgn(rgn2, rgnNorm, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
		CombineRgn(rgnNorm, rgn2, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
		CombineRgn(rgn2, rgnNorm, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
		CombineRgn(rgnNorm, rgn2, rgn1, RGN_DIFF)
		DeleteObject(rgn1)
		
		DeleteObject(rgn2)
	End Sub
	
	Private Sub SetAccessKeys()
		'this is a TRUE access keys parser
		'the basic rule is that if an ampersand is followed by another,
		'  a single ampersand is drawn and this is not the access key.
		'  So we continue searching for another possible access key.
		
		'   I only do a second pass because no one writes text like "Me & them & everyone"
		'   so the caption prop should be "Me && them && &everyone", this is rubbish and a
		'   search like this would only waste time
		Dim ampersandPos As Integer
		
		'we first clear the AccessKeys property, and will be filled if one is found
		'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		MyBase.AccessKeys = ""
		
		If Len(elTex) > 1 Then
			ampersandPos = InStr(1, elTex, "&", CompareMethod.Text)
			If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
				If Mid(elTex, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
					'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					MyBase.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
				Else 'do only a second pass to find another ampersand character
					ampersandPos = InStr(ampersandPos + 2, elTex, "&", CompareMethod.Text)
					If Mid(elTex, ampersandPos + 1, 1) <> "&" Then
						'UPGRADE_ISSUE: UserControl property UserControl.AccessKeys was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						MyBase.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
					End If
				End If
			End If
		End If
	End Sub
	
	Private Function ShiftColor(ByVal Color As Integer, ByVal Value As Integer) As Integer
		'this function will add or remove a certain color
		'quantity and return the result
		
		Dim Blue, Red, Green As Integer
		
		'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
		If isSoft Then Value = Value \ 2
		
		Blue = ((Color \ &H10000) Mod &H100s)
		Blue = Blue + ((Blue * Value) \ &HC0s)
		
		Green = ((Color \ &H100s) Mod &H100s) + Value
		Red = CShort(Color And &HFFs) + Value
		
		'a bit of optimization done here, values will overflow a
		' byte only in one direction... eg: if we added 32 to our
		' color, then only a > 255 overflow can occurr.
		If Value > 0 Then
			If Red > 255 Then Red = 255
			If Green > 255 Then Green = 255
			If Blue > 255 Then Blue = 255
		ElseIf Value < 0 Then 
			If Red < 0 Then Red = 0
			If Green < 0 Then Green = 0
			If Blue < 0 Then Blue = 0
		End If
		
		'more optimization by replacing the RGB function by its correspondent calculation
		ShiftColor = Red + 256 * Green + 65536 * Blue
	End Function
	
	'Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
	'    Dim Red As Long, Blue As Long, Green As Long
	'    Dim Delta As Long
	
	'    Blue = ((theColor \ &H10000) Mod &H100)
	'    Green = ((theColor \ &H100) Mod &H100)
	'    Red = (theColor And &HFF)
	'    Delta = &HFF - Base
	
	'    Blue = Base + Blue * Delta \ &HFF
	'    Green = Base + Green * Delta \ &HFF
	'    Red = Base + Red * Delta \ &HFF
	
	'    If Red > 255 Then Red = 255
	'    If Green > 255 Then Green = 255
	'    If Blue > 255 Then Blue = 255
	
	'    ShiftColorOXP = Red + 256& * Green + 65536 * Blue
	'End Function
	
	Private Sub CalcTextRects()
		'this sub will calculate the rects required to draw the text
		Select Case PicPosition
			Case 0
				rc2.Left_Renamed = 1 + picSZ.x : rc2.Right_Renamed = Wi - 2 : rc2.Top_Renamed = 1 : rc2.Bottom_Renamed = He - 2
			Case 1
				rc2.Left_Renamed = 1 : rc2.Right_Renamed = Wi - 2 - picSZ.x : rc2.Top_Renamed = 1 : rc2.Bottom_Renamed = He - 2
			Case 2
				rc2.Left_Renamed = 1 : rc2.Right_Renamed = Wi - 2 : rc2.Top_Renamed = 1 + picSZ.y : rc2.Bottom_Renamed = He - 2
			Case 3
				rc2.Left_Renamed = 1 : rc2.Right_Renamed = Wi - 2 : rc2.Top_Renamed = 1 : rc2.Bottom_Renamed = He - 2 - picSZ.y
			Case 4
				rc2.Left_Renamed = 1 : rc2.Right_Renamed = Wi - 2 : rc2.Top_Renamed = 1 : rc2.Bottom_Renamed = He - 2
		End Select
		'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		DrawText(MyBase.hdc, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK)
		CopyRect(rc, rc2) : fc.x = rc.Right_Renamed - rc.Left_Renamed : fc.y = rc.Bottom_Renamed - rc.Top_Renamed
		Select Case PicPosition
			Case 0, 2
				OffsetRect(rc, (Wi - rc.Right_Renamed) \ 2, (He - rc.Bottom_Renamed) \ 2)
			Case 1
				OffsetRect(rc, (Wi - rc.Right_Renamed - picSZ.x - 4) \ 2, (He - rc.Bottom_Renamed) \ 2)
			Case 3
				OffsetRect(rc, (Wi - rc.Right_Renamed) \ 2, (He - rc.Bottom_Renamed - picSZ.y - 4) \ 2)
			Case 4
				OffsetRect(rc, (Wi - rc.Right_Renamed) \ 2, (He - rc.Bottom_Renamed) \ 2)
		End Select
		CopyRect(rc2, rc) : OffsetRect(rc2, 1, 1)
		
		Call CalcPicPos() 'once we have the text position we are able to calculate the pic position
	End Sub
	
	Public Sub DisableRefresh()
		'this is for fast button editing, once you disable the refresh,
		' you can change every prop without triggering the drawing methods.
		' once you are done, you call Refresh.
		isShown = False
	End Sub
	
	Public Overrides Sub Refresh()
		Call SetColors()
		Call CalcTextRects()
		isShown = True
		Call Redraw(lastStat, True)
	End Sub
	
	Private Function ConvertFromSystemColor(ByVal theColor As Integer) As Integer
		Call OleTranslateColor(theColor, 0, ConvertFromSystemColor)
	End Function
	
	Private Sub DrawCaption(ByVal State As Byte)
		'this code is commonly shared through all the buttons so
		' i took it and put it toghether here for easier readability
		' of the code, and to cut-down disk size.
		
		captOpt = State
		With MyBase
			Select Case State 'in this select case, we only change the text color and draw only text that needs rc2, at the end, text that uses rc will be drawn
				Case 0 'normal caption
					TxtFX(rc)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cText)
				Case 1 'hover caption
					TxtFX(rc)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cTextO)
				Case 2 'down caption
					TxtFX(rc2)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cTextO)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					DrawText(.hdc, elTex, Len(elTex), rc2, DT_CENTER)
				Case 3 'disabled embossed caption
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cHighLight)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					DrawText(.hdc, elTex, Len(elTex), rc2, DT_CENTER)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cShadow)
				Case 4 'disabled grey caption
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, cShadow)
				Case 5 'WinXP disabled caption
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, ShiftColor(XPFace, -&H68s))
			End Select
			'we now draw the text that is common in all the captions
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			If State <> 2 Then DrawText(.hdc, elTex, Len(elTex), rc, DT_CENTER)
		End With
	End Sub
	
	Private Sub DrawPictures(ByVal State As Byte)
		If picNormal Is Nothing Then Exit Sub 'check if there is a main picture, if not then exit
		
		With MyBase
			Select Case State
				Case 0 'normal & hover
					If Not isOver Then
						Call DoFX(0, picNormal)
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						TransBlt(.hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask,  ,  , useGrey)
					Else
						If Not picHover Is Nothing Then
							Call DoFX(0, picHover)
							'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
							TransBlt(.hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picHover, cMask)
						Else
							Call DoFX(0, picNormal)
							'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
							TransBlt(.hdc, picPT.x, picPT.y, picSZ.x, picSZ.y, picNormal, cMask)
						End If
					End If
				Case 1 'down
					If picHover Is Nothing Then
						Call DoFX(1, picNormal)
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						TransBlt(.hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask)
					Else
						'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
						TransBlt(.hdc, picPT.x + System.Math.Abs(3), picPT.y + System.Math.Abs(3), picSZ.x, picSZ.y, picHover, cMask)
					End If
				Case 2 'disabled
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					TransBlt(.hdc, picPT.x + 1, picPT.y + 1, picSZ.x, picSZ.y, picNormal, cMask,  ,  , True)
			End Select
		End With
		If PicPosition = PicPositions.cbBackground Then Call DrawCaption(captOpt)
	End Sub
	
	Private Sub DoFX(ByVal offset As Integer, ByVal thePic As System.Drawing.Image)
		Dim curFace As Integer
		If SFX > fx.cbNone Then
			curFace = XPFace
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			TransBlt(MyBase.hdc, picPT.x + 1 + offset, picPT.y + 1 + offset, picSZ.x, picSZ.y, thePic, cMask, ShiftColor(curFace, System.Math.Abs(CInt(SFX = fx.cbEngraved)) * FXDEPTH + CShort(SFX <> fx.cbEngraved) * FXDEPTH))
			'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			If SFX < fx.cbShadowed Then TransBlt(MyBase.hdc, picPT.x - 1 + offset, picPT.y - 1 + offset, picSZ.x, picSZ.y, thePic, cMask, ShiftColor(curFace, System.Math.Abs(CInt(SFX <> fx.cbEngraved)) * FXDEPTH + CShort(SFX = fx.cbEngraved) * FXDEPTH))
		End If
	End Sub
	
	Private Sub TxtFX(ByRef theRect As RECT)
		Dim curFace As Integer
		Dim tempR As RECT
		If SFX > fx.cbNone Then
			With MyBase : CopyRect(tempR, theRect) : OffsetRect(tempR, 1, 1)
				
				curFace = XPFace
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				SetTextColor(.hdc, ShiftColor(curFace, System.Math.Abs(CInt(SFX = fx.cbEngraved)) * FXDEPTH + CShort(SFX <> fx.cbEngraved) * FXDEPTH))
				'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
				DrawText(.hdc, elTex, Len(elTex), tempR, DT_CENTER)
				If SFX < fx.cbShadowed Then
					OffsetRect(tempR, -2, -2)
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					SetTextColor(.hdc, ShiftColor(curFace, System.Math.Abs(CInt(SFX <> fx.cbEngraved)) * FXDEPTH + CShort(SFX = fx.cbEngraved) * FXDEPTH))
					'UPGRADE_ISSUE: UserControl property UserControl.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
					DrawText(.hdc, elTex, Len(elTex), tempR, DT_CENTER)
				End If
			End With
		End If
	End Sub
	
	Private Sub CalcPicSize()
		If Not picNormal Is Nothing Then
			'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: Picture property picNormal.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: UserControl method UserControl.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			picSZ.x = MyBase.ScaleX(picNormal.Width, 8, MyBase.ScaleMode)
			'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: Picture property picNormal.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			'UPGRADE_ISSUE: UserControl method UserControl.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			picSZ.y = MyBase.ScaleY(picNormal.Height, 8, MyBase.ScaleMode)
		Else
			picSZ.x = 0 : picSZ.y = 0
		End If
	End Sub
	
	Private Sub CalcPicPos()
		'exit if there's no picture
		If picNormal Is Nothing And picHover Is Nothing Then Exit Sub
		
		If (Trim(elTex) <> "") And (PicPosition <> 4) Then 'if there is no caption, or we have the picture as background, then we put the picture at the center of the button
			Select Case PicPosition
				Case 0 'left
					picPT.x = rc.Left_Renamed - picSZ.x - 4
					picPT.y = (He - picSZ.y) \ 2
				Case 1 'right
					picPT.x = rc.Right_Renamed + 4
					picPT.y = (He - picSZ.y) \ 2
				Case 2 'top
					picPT.x = (Wi - picSZ.x) \ 2
					picPT.y = rc.Top_Renamed - picSZ.y - 2
				Case 3 'bottom
					picPT.x = (Wi - picSZ.x) \ 2
					picPT.y = rc.Bottom_Renamed + 2
			End Select
		Else 'center the picture
			picPT.x = (Wi - picSZ.x) \ 2
			picPT.y = (He - picSZ.y) \ 2
		End If
	End Sub
	
	Private Sub TransBlt(ByVal DstDC As Integer, ByVal DstX As Integer, ByVal DstY As Integer, ByVal DstW As Integer, ByVal DstH As Integer, ByVal SrcPic As System.Drawing.Image, Optional ByVal TransColor As Integer = -1, Optional ByVal BrushColor As Integer = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False)
		If DstW = 0 Or DstH = 0 Then Exit Sub
		
		Dim i, H, B, f, newW As Integer
		Dim TmpBmp, TmpDC, TmpObj As Integer
		Dim Sr2Bmp, Sr2DC, Sr2Obj As Integer
		Dim Data1() As RGBTRIPLE
		Dim Data2() As RGBTRIPLE
		Dim info As BITMAPINFO
		Dim BrushRGB As RGBTRIPLE
		Dim gCol As Integer
		
		Dim tObj, SrcDC, ttt As Integer
		
		'UPGRADE_ISSUE: UserControl property XpButton.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		SrcDC = CreateCompatibleDC(hdc)
		
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: Picture property SrcPic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		If DstW < 0 Then DstW = MyBase.ScaleX(SrcPic.Width, 8, MyBase.ScaleMode)
		'UPGRADE_ISSUE: UserControl property UserControl.ScaleMode was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: Picture property SrcPic.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		'UPGRADE_ISSUE: UserControl method UserControl.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		If DstH < 0 Then DstH = MyBase.ScaleY(SrcPic.Height, 8, MyBase.ScaleMode)
		
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		Dim bR As RECT
		Dim hBrush As Integer
		If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
			tObj = SelectObject(SrcDC, CInt(CObj(SrcPic)))
		Else : bR.Right_Renamed = DstW : bR.Bottom_Renamed = DstH
			ttt = CreateCompatibleBitmap(DstDC, DstW, DstH) : tObj = SelectObject(SrcDC, ttt)
			hBrush = CreateSolidBrush(System.Drawing.ColorTranslator.ToOle(MaskColor)) : FillRect(SrcDC, bR, hBrush)
			DeleteObject(hBrush)
			'UPGRADE_ISSUE: Picture property SrcPic.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
			DrawIconEx(SrcDC, 0, 0, SrcPic.Handle, 0, 0, 0, 0, &H1s Or &H2s)
		End If
		
		TmpDC = CreateCompatibleDC(SrcDC)
		Sr2DC = CreateCompatibleDC(SrcDC)
		TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
		TmpObj = SelectObject(TmpDC, TmpBmp)
		Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
		ReDim Data1(DstW * DstH * 3 - 1)
		ReDim Data2(UBound(Data1))
		With info.bmiHeader
			.biSize = Len(info.bmiHeader)
			.biWidth = DstW
			.biHeight = DstH
			.biPlanes = 1
			.biBitCount = 24
		End With
		
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		BitBlt(TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy)
		'UPGRADE_ISSUE: Constant vbSrcCopy was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2070"'
		BitBlt(Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy)
		'UPGRADE_WARNING: Couldn't resolve default property of object Data1(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetDIBits(TmpDC, TmpBmp, 0, DstH, Data1(0), info, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object Data2(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		GetDIBits(Sr2DC, Sr2Bmp, 0, DstH, Data2(0), info, 0)
		
		If BrushColor > 0 Then
			BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100s
			BrushRGB.rgbGreen = (BrushColor \ &H100s) Mod &H100s
			BrushRGB.rgbRed = BrushColor And &HFFs
		End If
		
		If Not useMask Then TransColor = -1
		
		newW = DstW - 1
		
		For H = 0 To DstH - 1
			f = H * DstW
			For B = 0 To newW
				i = f + B
				If (CInt(Data2(i).rgbRed) + 256 * Data2(i).rgbGreen + 65536 * Data2(i).rgbBlue) <> TransColor Then
					With Data1(i)
						If BrushColor > -1 Then
							If MonoMask Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
								If (CInt(Data2(i).rgbRed) + Data2(i).rgbGreen + Data2(i).rgbBlue) <= 384 Then Data1(i) = BrushRGB
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
								Data1(i) = BrushRGB
							End If
						Else
							If isGreyscale Then
								gCol = CInt(Data2(i).rgbRed * 0.3) + Data2(i).rgbGreen * 0.59 + Data2(i).rgbBlue * 0.11
								.rgbRed = gCol : .rgbGreen = gCol : .rgbBlue = gCol
							Else
								'.rgbRed = (CLng(.rgbRed) + Data2(i).rgbRed * 2) \ 3
								'.rgbGreen = (CLng(.rgbGreen) + Data2(i).rgbGreen * 2) \ 3
								'.rgbBlue = (CLng(.rgbBlue) + Data2(i).rgbBlue * 2) \ 3
								'UPGRADE_WARNING: Couldn't resolve default property of object Data1(i). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
								Data1(i) = Data2(i)
							End If
						End If
					End With
				End If
			Next 
		Next 
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Data1(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		SetDIBitsToDevice(DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), info, 0)
		
		Erase Data1
		Erase Data2
		DeleteObject(SelectObject(TmpDC, TmpObj))
		DeleteObject(SelectObject(Sr2DC, Sr2Obj))
		'UPGRADE_ISSUE: Picture property SrcPic.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2064"'
		If SrcPic.Type = 3 Then DeleteObject(SelectObject(SrcDC, tObj))
		DeleteDC(TmpDC) : DeleteDC(Sr2DC)
		DeleteObject(tObj) : DeleteObject(ttt) : DeleteDC(SrcDC)
	End Sub
	
	Private Function isMouseOver() As Boolean
		Dim pt As POINTAPI
		
		GetCursorPos(pt)
		isMouseOver = (WindowFromPoint(pt.x, pt.y) = hwnd)
	End Function
End Class