Attribute VB_Name = "mdlGDI"
Option Explicit

'
' mGDI.bas for cTVBackground/cTile
'
' The BitmapToPicture routine was written by
' Bruce McKinney, from his Book "Hardcore Visual Basic" 2nd Edition
'
' http://vbaccelerator.com/
'


' Types:
'Type RECT
'    Left As Long
'    TOp As Long
'    Right As Long
'    Bottom As Long
'End Type
Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' General:
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function InvalidateRect Lib "user32" _
   (ByVal hwnd As Long, ByVal lpRect As Long, _
   ByVal bErase As Long) As Long

' GDI object functions:
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const BITSPIXEL = 12
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE = 2
    Public Const TRANSPARENT = 1
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    ' vbSrcAnd Public Const SRCAND = &H8800C6
    ' vbSrcCopy Public Const SRCCOPY = &HCC0020
    ' vbSrcErase Public Const SRCERASE = &H440328
    ' vbSrcInvert Public Const SRCINVERT = &H660046
    ' vbSrcPaint Public Const SRCPAINT = &HEE0086
    Public Const BLACKNESS = &H42&
    Public Const WHITENESS = &HFF0062
    Public Const DSna = &H220326
Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const IMAGE_BITMAP = 0
Declare Function OleTranslateColor Lib "oleaut32.dll" _
    (ByVal lOleColor As Long, ByVal lHPalette As Long, _
    lColorRef As Long) As Long

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Public Function LoadPictureTransparent( _
        pic As Object, _
        ByVal Filename As String, _
        ByRef sError As String, _
        Optional ByVal bLoadTransparent As Boolean = False _
    ) As Boolean

Dim hBmp As Long
Dim lR As Long
Dim lFlags As Long

On Error GoTo gbLoadTransErrorPic
    
   'Load the image...
   If (bLoadTransparent) Then
       lFlags = LR_LOADFROMFILE Or LR_LOADTRANSPARENT
   Else
       lFlags = LR_LOADFROMFILE Or LR_LOADMAP3DCOLORS
   End If
   hBmp = LoadImage(0, Filename, IMAGE_BITMAP, 0, 0, lFlags)
   If (hBmp <> 0) Then
       ' Ge the picture from it:
       Set pic = BitmapToPicture(hBmp)
       LoadPictureTransparent = True
   Else
       sError = "Could not read the bitmap file."
   End If
   
   Exit Function

gbLoadTransErrorPic:
    sError = Err.Description
    Exit Function

End Function


Private Function BitmapToPicture(ByVal hBmp As Long) As IPicture

    If (hBmp = 0) Then Exit Function
    
    Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid
    
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
    .cbSizeofStruct = Len(tPicConv)
    .picType = vbPicTypeBitmap
    .hImage = hBmp
    End With
    
    ' Fill in IDispatch Interface ID
    With IGuid
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
    End With
    
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    
    ' Return it:
    Set BitmapToPicture = oNewPic
    

End Function





