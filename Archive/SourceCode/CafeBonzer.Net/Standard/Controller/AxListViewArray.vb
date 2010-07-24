'UPGRADE_WARNING: The entire project must be compiled once before a form with an ActiveX Control Array can be displayed

Imports System.ComponentModel

<ProvideProperty("Index",GetType(AxMSComctlLib.AxListView))> Public Class AxListViewArray
	Inherits Microsoft.VisualBasic.Compatibility.VB6.BaseOcxArray
	Implements IExtenderProvider

	Public Sub New()
		MyBase.New()
	End Sub

	Public Sub New(ByVal Container As IContainer)
		MyBase.New(Container)
	End Sub

	Public Shadows Event [BeforeLabelEdit] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_BeforeLabelEditEvent)
	Public Shadows Event [AfterLabelEdit] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_AfterLabelEditEvent)
	Public Shadows Event [ColumnClick] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ColumnClickEvent)
	Public Shadows Event [ItemClick] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ItemClickEvent)
	Public Shadows Event [KeyDownEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyDownEvent)
	Public Shadows Event [KeyUpEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyUpEvent)
	Public Shadows Event [KeyPressEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyPressEvent)
	Public Shadows Event [MouseDownEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseDownEvent)
	Public Shadows Event [MouseMoveEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseMoveEvent)
	Public Shadows Event [MouseUpEvent] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseUpEvent)
	Public Shadows Event [ClickEvent] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [DblClick] (ByVal sender As System.Object, ByVal e As System.EventArgs)
	Public Shadows Event [OLEStartDrag] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEStartDragEvent)
	Public Shadows Event [OLEGiveFeedback] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEGiveFeedbackEvent)
	Public Shadows Event [OLESetData] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLESetDataEvent)
	Public Shadows Event [OLECompleteDrag] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLECompleteDragEvent)
	Public Shadows Event [OLEDragOver] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEDragOverEvent)
	Public Shadows Event [OLEDragDrop] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEDragDropEvent)
	Public Shadows Event [ItemCheck] (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ItemCheckEvent)

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function CanExtend(ByVal target As Object) As Boolean Implements IExtenderProvider.CanExtend
		If TypeOf target Is AxMSComctlLib.AxListView Then
			Return BaseCanExtend(target)
		End If
	End Function

	Public Function GetIndex(ByVal o As AxMSComctlLib.AxListView) As Short
		Return BaseGetIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub SetIndex(ByVal o As AxMSComctlLib.AxListView, ByVal Index As Short)
		BaseSetIndex(o, Index)
	End Sub

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Function ShouldSerializeIndex(ByVal o As AxMSComctlLib.AxListView) As Boolean
		Return BaseShouldSerializeIndex(o)
	End Function

	<System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)> Public Sub ResetIndex(ByVal o As AxMSComctlLib.AxListView)
		BaseResetIndex(o)
	End Sub

	Default Public ReadOnly Property Item(ByVal Index As Short) As AxMSComctlLib.AxListView
		Get
			Item = CType(BaseGetItem(Index), AxMSComctlLib.AxListView)
		End Get
	End Property

	Protected Overrides Function GetControlInstanceType() As System.Type
		Return GetType(AxMSComctlLib.AxListView)
	End Function

	Protected Overrides Sub HookUpControlEvents(ByVal o As Object)
		Dim ctl As AxMSComctlLib.AxListView = CType(o, AxMSComctlLib.AxListView)
		MyBase.HookUpControlEvents(o)
		If Not BeforeLabelEditEvent Is Nothing Then
			AddHandler ctl.BeforeLabelEdit, New AxMSComctlLib.ListViewEvents_BeforeLabelEditEventHandler(AddressOf HandleBeforeLabelEdit)
		End If
		If Not AfterLabelEditEvent Is Nothing Then
			AddHandler ctl.AfterLabelEdit, New AxMSComctlLib.ListViewEvents_AfterLabelEditEventHandler(AddressOf HandleAfterLabelEdit)
		End If
		If Not ColumnClickEvent Is Nothing Then
			AddHandler ctl.ColumnClick, New AxMSComctlLib.ListViewEvents_ColumnClickEventHandler(AddressOf HandleColumnClick)
		End If
		If Not ItemClickEvent Is Nothing Then
			AddHandler ctl.ItemClick, New AxMSComctlLib.ListViewEvents_ItemClickEventHandler(AddressOf HandleItemClick)
		End If
		If Not KeyDownEventEvent Is Nothing Then
			AddHandler ctl.KeyDownEvent, New AxMSComctlLib.ListViewEvents_KeyDownEventHandler(AddressOf HandleKeyDownEvent)
		End If
		If Not KeyUpEventEvent Is Nothing Then
			AddHandler ctl.KeyUpEvent, New AxMSComctlLib.ListViewEvents_KeyUpEventHandler(AddressOf HandleKeyUpEvent)
		End If
		If Not KeyPressEventEvent Is Nothing Then
			AddHandler ctl.KeyPressEvent, New AxMSComctlLib.ListViewEvents_KeyPressEventHandler(AddressOf HandleKeyPressEvent)
		End If
		If Not MouseDownEventEvent Is Nothing Then
			AddHandler ctl.MouseDownEvent, New AxMSComctlLib.ListViewEvents_MouseDownEventHandler(AddressOf HandleMouseDownEvent)
		End If
		If Not MouseMoveEventEvent Is Nothing Then
			AddHandler ctl.MouseMoveEvent, New AxMSComctlLib.ListViewEvents_MouseMoveEventHandler(AddressOf HandleMouseMoveEvent)
		End If
		If Not MouseUpEventEvent Is Nothing Then
			AddHandler ctl.MouseUpEvent, New AxMSComctlLib.ListViewEvents_MouseUpEventHandler(AddressOf HandleMouseUpEvent)
		End If
		If Not ClickEventEvent Is Nothing Then
			AddHandler ctl.ClickEvent, New System.EventHandler(AddressOf HandleClickEvent)
		End If
		If Not DblClickEvent Is Nothing Then
			AddHandler ctl.DblClick, New System.EventHandler(AddressOf HandleDblClick)
		End If
		If Not OLEStartDragEvent Is Nothing Then
			AddHandler ctl.OLEStartDrag, New AxMSComctlLib.ListViewEvents_OLEStartDragEventHandler(AddressOf HandleOLEStartDrag)
		End If
		If Not OLEGiveFeedbackEvent Is Nothing Then
			AddHandler ctl.OLEGiveFeedback, New AxMSComctlLib.ListViewEvents_OLEGiveFeedbackEventHandler(AddressOf HandleOLEGiveFeedback)
		End If
		If Not OLESetDataEvent Is Nothing Then
			AddHandler ctl.OLESetData, New AxMSComctlLib.ListViewEvents_OLESetDataEventHandler(AddressOf HandleOLESetData)
		End If
		If Not OLECompleteDragEvent Is Nothing Then
			AddHandler ctl.OLECompleteDrag, New AxMSComctlLib.ListViewEvents_OLECompleteDragEventHandler(AddressOf HandleOLECompleteDrag)
		End If
		If Not OLEDragOverEvent Is Nothing Then
			AddHandler ctl.OLEDragOver, New AxMSComctlLib.ListViewEvents_OLEDragOverEventHandler(AddressOf HandleOLEDragOver)
		End If
		If Not OLEDragDropEvent Is Nothing Then
			AddHandler ctl.OLEDragDrop, New AxMSComctlLib.ListViewEvents_OLEDragDropEventHandler(AddressOf HandleOLEDragDrop)
		End If
		If Not ItemCheckEvent Is Nothing Then
			AddHandler ctl.ItemCheck, New AxMSComctlLib.ListViewEvents_ItemCheckEventHandler(AddressOf HandleItemCheck)
		End If
	End Sub

	Private Sub HandleBeforeLabelEdit (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_BeforeLabelEditEvent) 
		RaiseEvent [BeforeLabelEdit] (sender, e)
	End Sub

	Private Sub HandleAfterLabelEdit (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_AfterLabelEditEvent) 
		RaiseEvent [AfterLabelEdit] (sender, e)
	End Sub

	Private Sub HandleColumnClick (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ColumnClickEvent) 
		RaiseEvent [ColumnClick] (sender, e)
	End Sub

	Private Sub HandleItemClick (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ItemClickEvent) 
		RaiseEvent [ItemClick] (sender, e)
	End Sub

	Private Sub HandleKeyDownEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyDownEvent) 
		RaiseEvent [KeyDownEvent] (sender, e)
	End Sub

	Private Sub HandleKeyUpEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyUpEvent) 
		RaiseEvent [KeyUpEvent] (sender, e)
	End Sub

	Private Sub HandleKeyPressEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_KeyPressEvent) 
		RaiseEvent [KeyPressEvent] (sender, e)
	End Sub

	Private Sub HandleMouseDownEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseDownEvent) 
		RaiseEvent [MouseDownEvent] (sender, e)
	End Sub

	Private Sub HandleMouseMoveEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseMoveEvent) 
		RaiseEvent [MouseMoveEvent] (sender, e)
	End Sub

	Private Sub HandleMouseUpEvent (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_MouseUpEvent) 
		RaiseEvent [MouseUpEvent] (sender, e)
	End Sub

	Private Sub HandleClickEvent (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [ClickEvent] (sender, e)
	End Sub

	Private Sub HandleDblClick (ByVal sender As System.Object, ByVal e As System.EventArgs) 
		RaiseEvent [DblClick] (sender, e)
	End Sub

	Private Sub HandleOLEStartDrag (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEStartDragEvent) 
		RaiseEvent [OLEStartDrag] (sender, e)
	End Sub

	Private Sub HandleOLEGiveFeedback (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEGiveFeedbackEvent) 
		RaiseEvent [OLEGiveFeedback] (sender, e)
	End Sub

	Private Sub HandleOLESetData (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLESetDataEvent) 
		RaiseEvent [OLESetData] (sender, e)
	End Sub

	Private Sub HandleOLECompleteDrag (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLECompleteDragEvent) 
		RaiseEvent [OLECompleteDrag] (sender, e)
	End Sub

	Private Sub HandleOLEDragOver (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEDragOverEvent) 
		RaiseEvent [OLEDragOver] (sender, e)
	End Sub

	Private Sub HandleOLEDragDrop (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_OLEDragDropEvent) 
		RaiseEvent [OLEDragDrop] (sender, e)
	End Sub

	Private Sub HandleItemCheck (ByVal sender As System.Object, ByVal e As AxMSComctlLib.ListViewEvents_ItemCheckEvent) 
		RaiseEvent [ItemCheck] (sender, e)
	End Sub

End Class

