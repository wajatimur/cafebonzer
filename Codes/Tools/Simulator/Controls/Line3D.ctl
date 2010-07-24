VERSION 5.00
Begin VB.UserControl uLine3D 
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   79
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
Attribute VB_Name = "uLine3D"
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
    aHorizon = PropBag.ReadProperty("horizon", 2)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width > UserControl.Height Then aHorizon = True
    If aHorizon = False Then
        cWidth = 50
        cHeight = UserControl.Height
        bDark.x1 = 1
        bDark.x2 = 1
        bDark.Y1 = 0
        bDark.Y2 = cHeight
        bLight.x1 = 2
        bLight.x2 = 2
        bLight.Y1 = 0
        bLight.Y2 = cHeight
    Else
        cWidth = UserControl.Width
        cHeight = 50
        bDark.x1 = 0
        bDark.x2 = cWidth
        bDark.Y1 = 1
        bDark.Y2 = 1
        bLight.x1 = 0
        bLight.x2 = cWidth
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

Public Property Let Alignment(ByVal Value As Align)
    Select Case Value
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
