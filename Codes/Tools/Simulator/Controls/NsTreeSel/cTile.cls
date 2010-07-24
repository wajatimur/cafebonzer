VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' Filename: ClsTile.Cls
' Author:   Steve McMahon
' Date:     15 June 1999
'
' Tiles a picture over the specified area in a DC
'
' --------------------------------------------------------------------
' vbAccelerator - Advanced, Free Source Code:
' http://vbaccelerator.com/
' ====================================================================

Private m_lHdc As Long
Private m_lHBmp As Long
Private m_lHBmpOld As Long
Private m_lhPalOld As Long
Private m_pic As StdPicture
Private m_sFileName As String
Private m_lXOriginOffset As Long
Private m_lYOriginOffset As Long
Private m_lBitmapW As Long
Private m_lBitmapH As Long

Private Const cTileErrorBase = 5600

Public Property Get PicturehDC() As Long
    PicturehDC = m_lHdc
End Property
Public Property Get XOriginOffset() As Long
    XOriginOffset = m_lXOriginOffset
End Property
Public Property Let XOriginOffset(ByVal lPixels As Long)
    m_lXOriginOffset = lPixels
End Property
Public Property Get YOriginOffset() As Long
    YOriginOffset = m_lYOriginOffset
End Property
Public Property Let YOriginOffset(ByVal lPiYels As Long)
    m_lYOriginOffset = lPiYels
End Property
Public Property Get BitmapWidth() As Long
    BitmapWidth = m_lBitmapW
End Property
Public Property Get BitmapHeight() As Long
    BitmapHeight = m_lBitmapH
End Property
Private Sub pErr(lNumber As Long, smsg As String)
    MsgBox "Error: " & smsg & ", " & lNumber, vbExclamation
End Sub
Public Property Let Filename( _
        ByVal sFileName As String _
    )
    ' Load a picture from a file:
    If (m_sFileName <> sFileName) Then
        pClearUp
        If (pbLoadPicture(sFileName)) Then
            m_sFileName = sFileName
        End If
    End If
End Property
Public Property Get Filename() As String
    Filename = m_sFileName
End Property
Public Property Get Picture() As StdPicture
    Set Picture = m_pic
End Property
Public Property Let Picture(oPic As StdPicture)
   ' Load a picture from a StdPicture object:
   pClearUp
   If Not oPic Is Nothing Then
      If (pbEnsurePicture()) Then
          Set m_pic = oPic
          If (Err.Number = 0) Then
              pbGetBitmapIntoDC
          End If
      End If
   End If
End Property
Private Function pbEnsurePicture() As Boolean
On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture
        If (Err.Number <> 0) Then
            pErr 3, "Unable to allocate memory for picture object."
            pbEnsurePicture = False
        Else
        End If
    End If
On Error GoTo 0
    Exit Function
End Function
Private Function pbLoadPictureFromFile(sFile As String) As Boolean
On Error Resume Next
   Set m_pic = LoadPicture(sFile)
   If (Err.Number <> 0) Then
       pErr 0, "Load Picture Failed: " & Err.Description
   Else
       pbLoadPictureFromFile = True
   End If
On Error GoTo 0
    Exit Function
End Function
Private Function pbLoadPicture(sFile As String) As Boolean

    If (pbEnsurePicture()) Then
        If (pbLoadPictureFromFile(sFile)) Then
            pbLoadPicture = pbGetBitmapIntoDC()
        End If
    End If
    
End Function
Private Function pbGetBitmapIntoDC() As Boolean
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long
Dim lHDCTemp As Long, lHBmpTempOld As Long

    ' Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    lHDCTemp = CreateCompatibleDC(lHDC)
    If (m_lHdc <> 0) Then
        ' Get size of bitmap:
        GetObjectAPI m_pic.Handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        lHBmpTempOld = SelectObject(lHDCTemp, m_pic.Handle)
        
        m_lHBmp = CreateCompatibleBitmap(lHDC, m_lBitmapW, m_lBitmapH)
        m_lHBmpOld = SelectObject(m_lHdc, m_lHBmp)
        
        BitBlt m_lHdc, 0, 0, m_lBitmapW, m_lBitmapH, lHDCTemp, 0, 0, vbSrcCopy
        
        SelectObject lHDCTemp, lHBmpTempOld
        DeleteDC lHDCTemp
        
        If (m_lHBmpOld <> 0) Then
            pbGetBitmapIntoDC = True
            If (m_sFileName = "") Then
               m_sFileName = "PICTURE"
            End If
        Else
            pClearUp
            pErr 2, "Unable to select bitmap into DC"
        End If
    Else
        pErr 1, "Unable to create compatible DC"
    End If
    ReleaseDC lHwnd, lHDC
    
End Function
Public Property Get Palette() As StdPicture
    Set Palette = m_pic
End Property
Private Sub pClearUp()
    ' Clear reference to the filename:
    m_sFileName = ""
    ' If we have a DC, then clear up:
    If (m_lHdc <> 0) Then
        ' Select the bitmap out of DC:
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' The original bitmap does not have to deleted because it is owned by m_pic
        End If
        If (m_lHBmp <> 0) Then
            DeleteObject m_lHBmp
        End If
        ' Remove the DC:
        DeleteDC m_lHdc
    End If
End Sub
Public Sub TileArea( _
        ByRef hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((x + m_lXOriginOffset) Mod m_lBitmapW)
    lSrcStartY = ((y + m_lYOriginOffset) Mod m_lBitmapH)
    lSrcStartWidth = (m_lBitmapW - lSrcStartX)
    lSrcStartHeight = (m_lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, m_lHdc, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = m_lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = m_lBitmapH
    Loop
End Sub


Private Sub Class_Terminate()
    ' Ensure all GDI objects are freed:
    pClearUp
    ' Clear up the picture:
    Set m_pic = Nothing
End Sub

