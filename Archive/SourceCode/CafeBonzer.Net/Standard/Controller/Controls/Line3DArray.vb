Option Strict Off
Option Explicit On
Imports System.Windows.Forms
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Compatibility.VB6

<ProvideProperty("Index", GetType(Line3D))> Friend Class Line3DArray
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
	
	Public Function CanExtend(ByVal Target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf Target Is Line3D Then
			Return BaseCanExtend(Target)
		End If
	End Function
	
	Public Function GetIndex(ByVal o As Line3D) As Short
		Return BaseGetIndex(o)
	End Function
	
	Public Sub SetIndex(ByVal o As Line3D, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub
	
	Public Function ShouldSerializeIndex(ByVal o As Line3D) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function
	
	Public Sub ResetIndex(ByVal o As Line3D)
		BaseResetIndex(o)
	End Sub
	
	Public Default ReadOnly Property Item(ByVal Index As Short) As Line3D
		Get
			Item = CType(BaseGetItem(Index), Line3D)
		End Get
	End Property
	
	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		
		Dim ctl As Line3D
		ctl = CType(o, Line3D)
		
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
	
	
	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(Line3D)
	End Function
	
End Class