VERSION 5.00
Begin VB.UserControl CasGuiLine 
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   25
   ToolboxBitmap   =   "Line3D.ctx":0000
   Begin VB.Line bLight 
      BorderColor     =   &H00FFFFFF&
      X1              =   7
      X2              =   7
      Y1              =   2
      Y2              =   159
   End
   Begin VB.Line bDark 
      BorderColor     =   &H00808080&
      X1              =   6
      X2              =   6
      Y1              =   2
      Y2              =   161
   End
End
Attribute VB_Name = "CasGuiLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum Align
    Vertical = 1
    Horizontal = 2
End Enum

Public Pbag As New PropertyBag
Private cHeight As Long
Private cWidth As Long
Private aHorizon As Boolean

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    aHorizon = PropBag.ReadProperty("horizon", True)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width > UserControl.Height Then aHorizon = True
    If aHorizon = False Then
        cWidth = 50
        cHeight = UserControl.Height
        bDark.X1 = 1
        bDark.X2 = 1
        bDark.Y1 = 0
        bDark.Y2 = cHeight
        bLight.X1 = 2
        bLight.X2 = 2
        bLight.Y1 = 0
        bLight.Y2 = cHeight
    Else
        cWidth = UserControl.Width
        cHeight = 50
        bDark.X1 = 0
        bDark.X2 = cWidth
        bDark.Y1 = 1
        bDark.Y2 = 1
        bLight.X1 = 0
        bLight.X2 = cWidth
        bLight.Y1 = 2
        bLight.Y2 = 2
    End If

    UserControl.Width = cWidth
    UserControl.Height = cHeight
End Sub


Public Property Get Alignment() As Align
    If aHorizon Then
        Alignment = Horizontal
    Else
        Alignment = Vertical
    End If
End Property

Public Property Let Alignment(ByVal value As Align)
    Select Case value
    Case 1
        If aHorizon = False Then Exit Property
        aHorizon = False
        UserControl.Height = UserControl.Width
    Case 2
        If aHorizon = True Then Exit Property
        aHorizon = True
        UserControl.Width = UserControl.Height
    End Select
    Call UserControl_Resize
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "horizon", aHorizon
End Sub
