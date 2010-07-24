VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTVbg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====================================================================
' Filename: cTVBackground.cls
' Author:   Ben Baird, modified by Steve McMahon
' Date:     15 June 1999
'
' Allows a background picture to be tiled into a TreeView control.
'
'
'
' The original source of this code is from Ben Baird's
' excellent web site, VB Thunder:
'     http://www.vbthunder.com/
'
' This version is modified:
' * To use SSUBTMR.DLL
' * Various GDI optimisations for minimal object creation and Blt size
'   (although they didn't do all that much for the speed)
' * Allow the background to scroll with the TreeView, makes
'   some operations quicker and prevents a redraw problem but may not
'   be the desired effect.  I'm not sure it is possible to tile the
'   background and keep it in place in the TreeView.
'
' This technique is best limited to small(-ish) TreeViews because it
' can be quite slow to draw.
'
' ====================================================================

' Don't use the VB API declare of PAINTSTRUCT-
' it misses out the full length of reserved data
' bytes, causing a GPF under NT
Private Type PAINTSTRUCT
   hdc As Long
   fErase As Long
   rcPaint As RECT
   fRestore As Long
   fIncUpdate As Long
   rgbReserved(0 To 31) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" _
   (ByVal hwnd As Long, lpPaint As Any) As Long
Private Declare Function EndPaint Lib "user32" _
   (ByVal hwnd As Long, lpPaint As Any) As Long
' Use LockWindowUpdate with care! - If you call it and there
' is an attempt to resize the control or draw something that
' was previously hidden then there is screen flicker.
Private Declare Function LockWindowUpdate Lib "user32" _
   (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" (ByVal hwnd As Long, _
   ByVal wMsg As Long, ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function GetClientRect Lib "user32" _
   (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_ERASEBKGND = &H14
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_DESTROY = &H2

Private Const TV_FIRST = &H1100&                    '// TreeView messages
Private Const TVN_FIRST = -400&                  '// treeview
Private Const TVN_LAST = -499&

Private Const TVGN_ROOT = &H0&
Private Const TVGN_NEXT = &H1&
Private Const TVGN_PREVIOUS = &H2&
Private Const TVGN_PARENT = &H3&
Private Const TVGN_CHILD = &H4&
Private Const TVGN_FIRSTVISIBLE = &H5&
Private Const TVGN_NEXTVISIBLE = &H6&
Private Const TVGN_PREVIOUSVISIBLE = &H7&
Private Const TVGN_DROPHILITE = &H8&
Private Const TVGN_CARET = &H9&
Private Const TVGN_LASTVISIBLE = &HA&

Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const WM_NOTIFY = &H4E
Private Const TVN_ITEMEXPANDINGA = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDEDA = (TVN_FIRST - 6)
Private Type NMHDR
    hwndFrom As Long   ' Window handle of control sending message
    idfrom As Long        ' Identifier of control sending message
    code  As Long          ' Specifies the notification code
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Implements ISubclass
Private m_cT As New clsTile
Private m_hWNd As Long
Private m_hWNdParent As Long
Private m_hWndOwnerForm As Long

Private Type tDC
   hdc As Long
   hBmp As Long
   hBmpOld As Long
   lWidth As Long
   lHeight As Long
End Type
Private m_tDC(1 To 3) As tDC
Private m_lXOffset As Long

Private Sub pEnsureMemDC( _
      ByVal hDCC As Long, _
      ByVal lIndex As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
   With m_tDC(lIndex)
      If lWidth > .lWidth Or lHeight > .lHeight Or .hdc = 0 Then
         pClearUpDC lIndex
         .hdc = CreateCompatibleDC(hDCC)
         If .hdc <> 0 Then
            If hDCC = 0 Then hDCC = .hdc
            .hBmp = CreateCompatibleBitmap(hDCC, lWidth, lHeight)
            If .hBmp <> 0 Then
               .hBmpOld = SelectObject(.hdc, .hBmp)
               Dim tR As RECT, hBr As Long
               tR.Right = lWidth
               tR.Bottom = lHeight
               
               If Not Tile Is Nothing Then
                  Tile.TileArea .hdc, 0, 0, lWidth, lHeight
               Else
                  hBr = CreateSolidBrush(TranslateColor(vbWindowBackground))
                  FillRect .hdc, tR, hBr
                  DeleteObject hBr
               End If
            Else
               pClearUpDC lIndex
            End If
         End If
      End If
   End With
   
End Sub
Private Sub pClearUpDC(ByVal lIndex As Long)
   With m_tDC(lIndex)
      If .hBmpOld <> 0 Then
         SelectObject .hdc, .hBmpOld
      End If
      If .hBmp <> 0 Then
         DeleteObject .hBmp
      End If
      If .hdc <> 0 Then
         DeleteDC .hdc
      End If
   End With
End Sub

Public Function TreeViewMessage( _
      ByVal hwnd As Long, _
      ByVal iMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long, _
      ByRef bUseRetVal As Boolean _
   ) As Long
Static InProc As Boolean
Dim tPS As PAINTSTRUCT
Dim hdcTreeView As Long
Dim hDCC As Long
Dim trcTV As RECT, tRTV As RECT
Dim lPtr As Long
   
   Select Case iMsg
   Case WM_PAINT
      ' WM_PAINT will recurse when we CallOldWIndowProc:
       If InProc = True Then
           Exit Function
       End If
       InProc = True
       
       ' Get the TreeView Client Area:
       GetClientRect hwnd, tRTV
   
       'Begin painting. This API must be called in
       'response to the WM_PAINT message:
       BeginPaint hwnd, tPS
       Debug.Print lPtr, tPS.hdc, tPS.fErase, tPS.rcPaint.Left, tPS.rcPaint.Right, tPS.rcPaint.TOp, tPS.rcPaint.Bottom
       'Debug.Print tPS.fErase
       
       hdcTreeView = tPS.hdc
       LSet trcTV = tPS.rcPaint
   
       'Create two memory DCs to draw on:
       pEnsureMemDC hdcTreeView, 1, tRTV.Right - tRTV.Left, tRTV.Bottom - tRTV.TOp
       pEnsureMemDC hdcTreeView, 2, tRTV.Right - tRTV.Left, tRTV.Bottom - tRTV.TOp
   
       'This actually causes the TreeView to paint
       'itself onto our memory DC!
       CallOldWindowProc hwnd, WM_PAINT, m_tDC(1).hdc, 0&
       
       'Tile the bitmap and draw the TreeView
       'over it transparently
       Dim tR As RECT, hItemFirst As Long
       hItemFirst = SendMessage(hwnd, TVM_GETNEXTITEM, 0, ByVal TVGN_ROOT)
       tR.Left = hItemFirst
       SendMessage hwnd, TVM_GETITEMRECT, 1, tR
       If tR.Left > m_lXOffset Then m_lXOffset = tR.Left
       m_cT.XOriginOffset = -tR.Left + m_lXOffset
       m_cT.YOriginOffset = -tR.TOp
       m_cT.TileArea m_tDC(2).hdc, trcTV.Left, trcTV.TOp, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp
       
       ' Create a mono DC of the size to be painted this time:
       pEnsureMemDC 0, 3, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp
       ' Set BkColor of the TreeView Mem DC to match transparent colour:
       SetBkColor m_tDC(1).hdc, TranslateColor(vbWindowBackground)
       ' Copy from TreeView Mem DC -> mono DC:
       BitBlt m_tDC(3).hdc, 0, 0, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp, m_tDC(1).hdc, trcTV.Left, trcTV.TOp, vbSrcCopy
       
       ' Now the DC(3) is a mask.
       SetTextColor m_tDC(2).hdc, vbBlack
       SetBkColor m_tDC(2).hdc, vbWhite
       ' AND the mask to the Background so we go white where the
       ' treeview is black:
       BitBlt m_tDC(2).hdc, trcTV.Left, trcTV.TOp, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp, m_tDC(3).hdc, 0, 0, vbSrcAnd
       
       ' OR the mask to the TreeView so it goes black where the background is:
       SetTextColor m_tDC(1).hdc, vbBlack
       SetBkColor m_tDC(1).hdc, vbWhite
       BitBlt m_tDC(1).hdc, trcTV.Left, trcTV.TOp, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp, m_tDC(3).hdc, 0, 0, DSna
       
       ' Now OR the treeview with the unmasked background:
       BitBlt m_tDC(2).hdc, trcTV.Left, trcTV.TOp, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp, m_tDC(1).hdc, trcTV.Left, trcTV.TOp, vbSrcPaint
       
       'Draw to the target DC
       BitBlt hdcTreeView, trcTV.Left, trcTV.TOp, trcTV.Right - trcTV.Left, trcTV.Bottom - trcTV.TOp, _
           m_tDC(2).hdc, trcTV.Left, trcTV.TOp, vbSrcCopy
         
       EndPaint hwnd, tPS
   
       TreeViewMessage = 0
       bUseRetVal = True
       
       InProc = False
   
   Case WM_ERASEBKGND
       'Return TRUE
       TreeViewMessage = 0
       bUseRetVal = True
       
   Case WM_DESTROY
      Detach
      
   End Select
   
End Function

Public Sub Attach(tvwThis As Object, ByVal hWndOwnerForm As Long)
   Detach
   m_hWNd = tvwThis.hwnd
   m_hWNdParent = GetParent(m_hWNd)
   AttachMessage Me, m_hWNd, WM_PAINT
   AttachMessage Me, m_hWNd, WM_ERASEBKGND
   AttachMessage Me, m_hWNdParent, WM_NOTIFY
   AttachMessage Me, m_hWNd, WM_DESTROY
   m_hWndOwnerForm = hWndOwnerForm
End Sub


Public Sub Detach()
   If m_hWNd <> 0 Then
      DetachMessage Me, m_hWNd, WM_PAINT
      DetachMessage Me, m_hWNd, WM_ERASEBKGND
      DetachMessage Me, m_hWNdParent, WM_NOTIFY
      DetachMessage Me, m_hWNd, WM_DESTROY
   End If
   m_hWNd = 0
   m_hWNdParent = 0
End Sub

Public Property Get Tile() As clsTile
   Set Tile = m_cT
End Property

Private Function TranslateColor(lColor As Long, Optional ByVal hPal As Long = 0) As Long
Dim lR As Long
   OleTranslateColor lColor, hPal, lR
   TranslateColor = lR
End Function

Private Sub Class_Initialize()
   Set m_cT = New clsTile
End Sub

Private Sub Class_Terminate()
Dim i As Long
   Detach
   Set m_cT = Nothing
   For i = 1 To 3
      pClearUpDC i
   Next i
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer.EMsgResponse
   If CurrentMessage = WM_PAINT Or CurrentMessage = WM_ERASEBKGND Then
      ISubclass_MsgResponse = emrConsume
   Else
      ISubclass_MsgResponse = emrPreprocess
   End If
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lR As Long
Dim bR As Boolean
   If iMsg = WM_PAINT Or iMsg = WM_ERASEBKGND Then
      ' Only do this if something to tile:
      If m_cT.Filename <> "" Then
         ' Background processing:
         lR = TreeViewMessage(hwnd, iMsg, wParam, lParam, bR)
         If bR Then
            ISubclass_WindowProc = lR
         Else
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         End If
      Else
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
   ElseIf iMsg = WM_NOTIFY Then
      ' Only do this if something to tile:
      If m_cT.Filename <> "" Then
         Dim tNHM As NMHDR
         CopyMemory tNHM, ByVal lParam, Len(tNHM)
         If tNHM.hwndFrom = m_hWNd Then
            If tNHM.code = TVN_ITEMEXPANDINGA Then
               LockWindowUpdate m_hWndOwnerForm
            ElseIf tNHM.code = TVN_ITEMEXPANDEDA Then
               InvalidateRect m_hWNd, ByVal 0&, 0
               LockWindowUpdate 0
            End If
         End If
      End If
   End If
      
End Function




