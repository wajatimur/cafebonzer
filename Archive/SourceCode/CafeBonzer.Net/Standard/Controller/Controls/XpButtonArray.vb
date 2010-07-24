Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Compatibility.VB6

<ProvideProperty("Index", GetType(XpButton))> Friend Class XpButtonArray
	Inherits BaseControlArray
	Implements IExtenderProvider
	
	Public Sub New()
	End Sub
	
	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub
	
	Public Event [MouseHover] As System.EventHandler
	Public Event [SystemColorsChanged] As System.EventHandler
	
	Public Event [DoubleClick] As System.EventHandler
	Public Event [Resize] As System.EventHandler
	Public Event [Enter] As System.EventHandler
	Public Event [Leave] As System.EventHandler
	Public Event [LostFocus] As System.EventHandler
	Public Event [GotFocus] As System.EventHandler
	Public Event [Validating] As System.ComponentModel.CancelEventHandler
	Public Event [MouseOver] As System.EventHandler
	Public Event [KeyPress] As XpButton.KeyPressEventHandler
	Public Event [KeyUp] As XpButton.KeyUpEventHandler
	Public Event [MouseDown] As XpButton.MouseDownEventHandler
	Public Event [MouseUp] As XpButton.MouseUpEventHandler
	Public Event [MouseOut] As System.EventHandler
	Public Event [Click] As System.EventHandler
	Public Event [KeyDown] As XpButton.KeyDownEventHandler
	Public Event [MouseMove] As XpButton.MouseMoveEventHandler
	
	Public Function CanExtend(ByVal Target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf Target Is XpButton Then
			Return BaseCanExtend(Target)
		End If
	End Function
	
	Public Function GetIndex(ByVal o As XpButton) As Short
		Return BaseGetIndex(o)
	End Function
	
	Public Sub SetIndex(ByVal o As XpButton, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub
	
	Public Function ShouldSerializeIndex(ByVal o As XpButton) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function
	
	Public Sub ResetIndex(ByVal o As XpButton)
		BaseResetIndex(o)
	End Sub
	
	Public Default ReadOnly Property Item(ByVal Index As Short) As XpButton
		Get
			Item = CType(BaseGetItem(Index), XpButton)
		End Get
	End Property
	
	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		
		Dim ctl As XpButton
		ctl = CType(o, XpButton)
		
		If Not IsNothing(DoubleClickEvent) Then
			addHandler ctl.DoubleClick, New System.EventHandler(AddressOf HandleDoubleClick)
		End If
		
		If Not IsNothing(ResizeEvent) Then
			addHandler ctl.Resize, New System.EventHandler(AddressOf HandleResize)
		End If
		
		If Not IsNothing(EnterEvent) Then
			addHandler ctl.Enter, New System.EventHandler(AddressOf HandleEnter)
		End If
		
		If Not IsNothing(LeaveEvent) Then
			addHandler ctl.Leave, New System.EventHandler(AddressOf HandleLeave)
		End If
		
		If Not IsNothing(LostFocusEvent) Then
			addHandler ctl.LostFocus, New System.EventHandler(AddressOf HandleLostFocus)
		End If
		
		If Not IsNothing(GotFocusEvent) Then
			addHandler ctl.GotFocus, New System.EventHandler(AddressOf HandleGotFocus)
		End If
		
		If Not IsNothing(ValidatingEvent) Then
			addHandler ctl.Validating, New System.ComponentModel.CancelEventHandler(AddressOf HandleValidating)
		End If
		
		If Not IsNothing(MouseOverEvent) Then
			addHandler ctl.MouseOver, New XpButton.MouseOverEventHandler(AddressOf HandleMouseOver)
		End If
		
		If Not IsNothing(KeyPressEvent) Then
			addHandler ctl.KeyPress, New XpButton.KeyPressEventHandler(AddressOf HandleKeyPress)
		End If
		
		If Not IsNothing(KeyUpEvent) Then
			addHandler ctl.KeyUp, New XpButton.KeyUpEventHandler(AddressOf HandleKeyUp)
		End If
		
		If Not IsNothing(MouseDownEvent) Then
			addHandler ctl.MouseDown, New XpButton.MouseDownEventHandler(AddressOf HandleMouseDown)
		End If
		
		If Not IsNothing(MouseUpEvent) Then
			addHandler ctl.MouseUp, New XpButton.MouseUpEventHandler(AddressOf HandleMouseUp)
		End If
		
		If Not IsNothing(MouseOutEvent) Then
			addHandler ctl.MouseOut, New XpButton.MouseOutEventHandler(AddressOf HandleMouseOut)
		End If
		
		If Not IsNothing(ClickEvent) Then
			addHandler ctl.Click, New XpButton.ClickEventHandler(AddressOf HandleClick)
		End If
		
		If Not IsNothing(KeyDownEvent) Then
			addHandler ctl.KeyDown, New XpButton.KeyDownEventHandler(AddressOf HandleKeyDown)
		End If
		
		If Not IsNothing(MouseMoveEvent) Then
			addHandler ctl.MouseMove, New XpButton.MouseMoveEventHandler(AddressOf HandleMouseMove)
		End If
		
	End Sub
	
	Private Sub HandleDoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [DoubleClick](sender, e)
	End Sub
	
	Private Sub HandleResize(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [Resize](sender, e)
	End Sub
	
	Private Sub HandleEnter(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [Enter](sender, e)
	End Sub
	
	Private Sub HandleLeave(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [Leave](sender, e)
	End Sub
	
	Private Sub HandleLostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [LostFocus](sender, e)
	End Sub
	
	Private Sub HandleGotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [GotFocus](sender, e)
	End Sub
	
	Private Sub HandleValidating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
		RaiseEvent [Validating](sender, e)
	End Sub
	
	Private Sub HandleMouseOver(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [MouseOver](sender, e)
	End Sub
	
	Private Sub HandleKeyPress(ByVal sender As Object, ByVal e As XpButton.KeyPressEventArgs)
		RaiseEvent [KeyPress](sender, e)
	End Sub
	
	Private Sub HandleKeyUp(ByVal sender As Object, ByVal e As XpButton.KeyUpEventArgs)
		RaiseEvent [KeyUp](sender, e)
	End Sub
	
	Private Sub HandleMouseDown(ByVal sender As Object, ByVal e As XpButton.MouseDownEventArgs)
		RaiseEvent [MouseDown](sender, e)
	End Sub
	
	Private Sub HandleMouseUp(ByVal sender As Object, ByVal e As XpButton.MouseUpEventArgs)
		RaiseEvent [MouseUp](sender, e)
	End Sub
	
	Private Sub HandleMouseOut(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [MouseOut](sender, e)
	End Sub
	
	Private Sub HandleClick(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [Click](sender, e)
	End Sub
	
	Private Sub HandleKeyDown(ByVal sender As Object, ByVal e As XpButton.KeyDownEventArgs)
		RaiseEvent [KeyDown](sender, e)
	End Sub
	
	Private Sub HandleMouseMove(ByVal sender As Object, ByVal e As XpButton.MouseMoveEventArgs)
		RaiseEvent [MouseMove](sender, e)
	End Sub
	
	
	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(XpButton)
	End Function
	
End Class