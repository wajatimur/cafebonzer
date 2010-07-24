Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Compatibility.VB6

<ProvideProperty("Index", GetType(Label3D))> Friend Class Label3DArray
	Inherits BaseControlArray
	Implements IExtenderProvider
	
	Public Sub New()
	End Sub
	
	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub
	
	Public Event [MouseHover] As System.EventHandler
	Public Event [SystemColorsChanged] As System.EventHandler
	
	Public Event [Resize] As System.EventHandler
	Public Event [Enter] As System.EventHandler
	Public Event [Leave] As System.EventHandler
	Public Event [LostFocus] As System.EventHandler
	Public Event [GotFocus] As System.EventHandler
	Public Event [Validating] As System.ComponentModel.CancelEventHandler
	Public Event [MouseDown] As Label3D.MouseDownEventHandler
	Public Event [MouseUp] As Label3D.MouseUpEventHandler
	Public Event [Click] As System.EventHandler
	Public Event [DblClick] As System.EventHandler
	Public Event [MouseMove] As Label3D.MouseMoveEventHandler
	
	Public Function CanExtend(ByVal Target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf Target Is Label3D Then
			Return BaseCanExtend(Target)
		End If
	End Function
	
	Public Function GetIndex(ByVal o As Label3D) As Short
		Return BaseGetIndex(o)
	End Function
	
	Public Sub SetIndex(ByVal o As Label3D, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub
	
	Public Function ShouldSerializeIndex(ByVal o As Label3D) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function
	
	Public Sub ResetIndex(ByVal o As Label3D)
		BaseResetIndex(o)
	End Sub
	
	Public Default ReadOnly Property Item(ByVal Index As Short) As Label3D
		Get
			Item = CType(BaseGetItem(Index), Label3D)
		End Get
	End Property
	
	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		
		Dim ctl As Label3D
		ctl = CType(o, Label3D)
		
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
		
		If Not IsNothing(MouseDownEvent) Then
			addHandler ctl.MouseDown, New Label3D.MouseDownEventHandler(AddressOf HandleMouseDown)
		End If
		
		If Not IsNothing(MouseUpEvent) Then
			addHandler ctl.MouseUp, New Label3D.MouseUpEventHandler(AddressOf HandleMouseUp)
		End If
		
		If Not IsNothing(ClickEvent) Then
			addHandler ctl.Click, New Label3D.ClickEventHandler(AddressOf HandleClick)
		End If
		
		If Not IsNothing(DblClickEvent) Then
			addHandler ctl.DblClick, New Label3D.DblClickEventHandler(AddressOf HandleDblClick)
		End If
		
		If Not IsNothing(MouseMoveEvent) Then
			addHandler ctl.MouseMove, New Label3D.MouseMoveEventHandler(AddressOf HandleMouseMove)
		End If
		
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
	
	Private Sub HandleMouseDown(ByVal sender As Object, ByVal e As Label3D.MouseDownEventArgs)
		RaiseEvent [MouseDown](sender, e)
	End Sub
	
	Private Sub HandleMouseUp(ByVal sender As Object, ByVal e As Label3D.MouseUpEventArgs)
		RaiseEvent [MouseUp](sender, e)
	End Sub
	
	Private Sub HandleClick(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [Click](sender, e)
	End Sub
	
	Private Sub HandleDblClick(ByVal sender As Object, ByVal e As System.EventArgs)
		RaiseEvent [DblClick](sender, e)
	End Sub
	
	Private Sub HandleMouseMove(ByVal sender As Object, ByVal e As Label3D.MouseMoveEventArgs)
		RaiseEvent [MouseMove](sender, e)
	End Sub
	
	
	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(Label3D)
	End Function
	
End Class