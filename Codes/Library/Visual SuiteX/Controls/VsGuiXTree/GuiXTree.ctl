VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl VsGuiXTree 
   BackColor       =   &H80000005&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1950
   ScaleWidth      =   4800
   ToolboxBitmap   =   "GuiXTree.ctx":0000
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   3975
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":00FA
            Key             =   "FOLDERCLOSED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":0696
            Key             =   "FOLDEROPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":0C32
            Key             =   "OPTIONYES3D"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":11CE
            Key             =   "OPTIONNO3D"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":176A
            Key             =   "CHECK_YES3D"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":1D06
            Key             =   "CHECK_NO3D"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":22A2
            Key             =   "CHECK_PART3D"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":283E
            Key             =   "OPTIONYES95"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":299A
            Key             =   "OPTIONNO95"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":2AF6
            Key             =   "CHECK_YES95"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":2C52
            Key             =   "CHECK_NO95"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GuiXTree.ctx":2DAE
            Key             =   "CHECK_PART95"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwOpt 
      Height          =   1665
      Left            =   285
      TabIndex        =   0
      Top             =   60
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   2937
      _Version        =   393217
      Indentation     =   452
      LabelEdit       =   1
      Style           =   5
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "VsGuiXTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ====================================================================
' Filename: XTreeOpt.ctl
' Author:   SP McMahon
' Date:     15 June 1999
'
' A Control which modifies a VB5 TreeView control to turn it into
' a Explorer Folder Options/IE Advanced Options style picker.
'
'
' --------------------------------------------------------------------
' vbAccelerator - Advanced, Free Source Code:
' http://vbaccelerator.com/
'
' ====================================================================

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Type TVITEM
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type
Private Const TVIF_STATE As Long = &H8
Private Const TVIS_CUT = &H4
Private Const TVIS_BOLD  As Long = &H10
Private Const TV_FIRST As Long = &H1100
Private Const TVS_TRACKSELECT As Long = &H200&
Private Const TVS_FULLROWSELECT As Long = &H1000
Private Const TVS_SINGLEEXPAND  As Long = &H400
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVM_GETITEM As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM As Long = (TV_FIRST + 13)
Private Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Private Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Private Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)
Private Const TVGN_ROOT               As Long = &H0
Private Const TVGN_NEXT               As Long = &H1
Private Const TVGN_PREVIOUS           As Long = &H2
Private Const TVGN_PARENT             As Long = &H3
Private Const TVGN_CHILD              As Long = &H4
Private Const TVGN_FIRSTVISIBLE       As Long = &H5
Private Const TVGN_NEXTVISIBLE        As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE    As Long = &H7
Private Const TVGN_DROPHILITE         As Long = &H8
Private Const TVGN_CARET              As Long = &H9
Private Const GWL_STYLE As Long = (-16)
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const WM_NOTIFY As Long = &H4E
Private Const H_MAX As Long = &HFFFF + 1
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Private Const TVN_FIRST = H_MAX - 400                  '// treeview
Private Const TVN_ITEMEXPANDING As Long = (TVN_FIRST - 5)
Private Const TVE_COLLAPSE As Long = &H1
Private Const TVE_EXPAND As Long = &H2
Private Const TVE_TOGGLE As Long = &H3
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type NMTREEVIEW
    hdr As NMHDR
    action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINTAPI
End Type
Public Enum OptionTreeFolderTypeCOnstants
    OptionTreeFolder = 1
    OptionTreeCheck = 2
    optiontreeFolderCustom = 3
End Enum
Public Enum OptionTreeCheckTypes
    OptionTreeCheckNone = 0
    OptionTreeCheckFull = 1
    OptionTreeCheckPartial = 2
End Enum
Public Enum OptionTreeIconSets
    OptionTreeIcons3d = 0
    OptionTreeIconsWin98 = 1
End Enum
Private Type tFolderInfo
    vKey As Variant
    vOpenKey As Variant
    vClosedKey As Variant
End Type
Private m_tFolderInfo() As tFolderInfo
Private m_iFolderCount As Integer
Private m_bTrackSelect As Boolean
Private m_bLocked As Boolean
Private m_eIconSet As OptionTreeIconSets
Private m_cTVB As GfxXTreeTile
Private Enum OptionTreeNodeClickReasons
    eoptKeyDown
    eoptMouseDown
End Enum
Private m_eNodeClickReason As OptionTreeNodeClickReasons
Public Event OptionClick(ItemNode As MSComctlLib.Node)
Public Event CheckClick(ItemNode As MSComctlLib.Node, Value As OptionTreeCheckTypes)
Public Event AfterLabelEdit(Cancel As Integer, NewString As String)
Attribute AfterLabelEdit.VB_Description = "Raised when label editing is completed by the user."
Public Event BeforeLabelEdit(Cancel As Integer)
Attribute BeforeLabelEdit.VB_Description = "Raised when a label edit is about to occur."

Public Sub Clear()
   tvwOpt.Nodes.Clear
End Sub
Public Property Set BackgroundPicture(ByRef sPic As StdPicture)
   pSetBackPic sPic
   PropertyChanged "BackgroundPicture)"
End Property
Public Property Let BackgroundPicture(ByRef sPic As StdPicture)
   pSetBackPic sPic
   PropertyChanged "BackgroundPicture)"
End Property
Public Property Get BackgroundPicture() As StdPicture
   If Not m_cTVB Is Nothing Then
      Set BackgroundPicture = m_cTVB.Tile.Picture
   End If
End Property
Private Sub pSetBackPic(ByRef sPic As StdPicture)
   If sPic Is Nothing Then
      If Not m_cTVB Is Nothing Then
         Set m_cTVB = Nothing
         InvalidateRect tvwOpt.hwnd, 0, 0
      End If
   Else
      If m_cTVB Is Nothing Then
         Set m_cTVB = New GfxXTreeTile
      End If
      m_cTVB.Tile.Picture = sPic
      If UserControl.Ambient.UserMode Then
         SendMessageLong tvwOpt.hwnd, TVM_SETBKCOLOR, 0, -1
         m_cTVB.Attach tvwOpt, UserControl.Parent.hwnd
      End If
   End If
End Sub
Public Property Get NodeCheckType(ByVal Item As Variant) As OptionTreeCheckTypes
Dim sItem As String
   sItem = tvwOpt.Nodes(Item).Image
   Select Case True
   Case InStr(sItem, "YES")
      NodeCheckType = OptionTreeCheckFull
   Case InStr(sItem, "NO")
      NodeCheckType = OptionTreeCheckNone
   Case InStr(sItem, "PART")
      NodeCheckType = OptionTreeCheckPartial
   End Select
End Property
Public Property Let NodeCheckType(ByVal Item As Variant, ByVal eType As OptionTreeCheckTypes)
Dim sImage As String
   If (eType = OptionTreeCheckPartial) Then
      ' Not allowed. Can only be set by clicking items.
   Else
      ' If the option is already set then exit else
      ' emulate a click on this node.
      sImage = tvwOpt.Nodes(Item).Image
      If (eType = OptionTreeCheckFull) And InStr(sImage, "YES") = 0 Then
         tvwOpt_NodeClick tvwOpt.Nodes(Item)
      ElseIf (eType = OptionTreeCheckNone) And InStr(sImage, "NO") = 0 Then
         tvwOpt_NodeClick tvwOpt.Nodes(Item)
      End If
   End If
End Property
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Gets/sets whether the selection rectangle is visible when the control is out of focus."
    HideSelection = tvwOpt.HideSelection
End Property
Public Property Let HideSelection(ByVal bState As Boolean)
    tvwOpt.HideSelection = bState
End Property

Public Property Get IconSet() As OptionTreeIconSets
Attribute IconSet.VB_Description = "Gets/sets whether to use Win98 or 3D style icons for the check boxes and option boxes."
    IconSet = m_eIconSet
End Property
Public Property Let IconSet(ByVal eSet As OptionTreeIconSets)
Dim sPf As String
Dim i As Long
Dim sI As String
Dim iLen As Long
    m_eIconSet = eSet
    sPf = Postfix()
    For i = 1 To tvwOpt.Nodes.Count
        With tvwOpt.Nodes(i)
            sI = Left$(.Image, 6)
            If (sI = "OPTION") Or (sI = "CHECK_") Then
                If Right$(sI, 2) <> sPf Then
                    iLen = Len(.Image) - 2
                    .Image = Left$(tvwOpt.Nodes(i).Image, iLen) & sPf
                End If
            End If
        End With
    Next i
End Property

Private Function Postfix() As String
    Select Case m_eIconSet
    Case OptionTreeIconsWin98
        Postfix = "95"
    Case OptionTreeIcons3d
        Postfix = "3D"
    End Select
End Function

Public Property Get InternalImageList() As Object
Attribute InternalImageList.VB_Description = "Gets a reference to the Image List control used in the control."
    Set InternalImageList = ilsIcons
End Property
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gets/sets whether the control allows user input or not."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal bState As Boolean)
Dim tVI As TVITEM
Dim hItem As Long
Dim lR As Long

    tVI.mask = TVIF_STATE
    hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
    Do While hItem <> 0
        With tVI
            .hItem = hItem
            .mask = TVIF_STATE
            .stateMask = TVIS_CUT
            If (bState) Then
                .State = .stateMask And Not TVIS_CUT
            Else
                .State = .stateMask Or TVIS_CUT
            End If
            lR = SendMessage(tvwOpt.hwnd, TVM_SETITEM, 0&, tVI)
        End With
        hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
    Loop
    UserControl.Enabled = bState
    PropertyChanged "Enabled"
End Property
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Gets/sets whether the control is locked.  When locked, users can view the tree but cannot change the settings."
   Locked = m_bLocked
End Property
Public Property Let Locked(ByVal bState As Boolean)
Dim tVI As TVITEM
Dim hItem As Long
Dim lR As Long
Dim lColor As Long
   
   m_bLocked = bState

   tVI.mask = TVIF_STATE
   hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
   Do While hItem <> 0
       With tVI
           .hItem = hItem
           .mask = TVIF_STATE
           .stateMask = TVIS_CUT
           If (bState) Then
              .State = .stateMask Or TVIS_CUT
           Else
               .State = .stateMask And Not TVIS_CUT
           End If
           lR = SendMessage(tvwOpt.hwnd, TVM_SETITEM, 0&, tVI)
       End With
       hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, hItem)
   Loop
   If (bState) Then
      ' Set the colour in the TreeView:
      lColor = TranslateColor(vbButtonShadow)
      SendMessageLong tvwOpt.hwnd, TVM_SETTEXTCOLOR, 0, lColor
    
      ' Request a redraw:
      InvalidateRectAsNull tvwOpt.hwnd, 0, 1
      UpdateWindow tvwOpt.hwnd
   Else
      ForeColor = UserControl.ForeColor
   End If

End Property

Public Property Get SingleClickExpand() As Boolean
Attribute SingleClickExpand.VB_Description = "Gets/sets whether nodes automatically expand when clicked and contract when left."
Dim lStyle As Long
    lStyle = GetWindowLong(tvwOpt.hwnd, GWL_STYLE)
    SingleClickExpand = ((lStyle And TVS_SINGLEEXPAND) = TVS_SINGLEEXPAND)
End Property
Public Property Let SingleClickExpand(ByVal bState As Boolean)
Dim lStyle As Long
    lStyle = GetWindowLong(tvwOpt.hwnd, GWL_STYLE)
    If (bState) Then
        lStyle = lStyle Or TVS_SINGLEEXPAND
    Else
        lStyle = lStyle And Not TVS_SINGLEEXPAND
    End If
    SetWindowLong tvwOpt.hwnd, GWL_STYLE, lStyle
    
    PropertyChanged "SingleClickExpand"
    
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Gets/sets whether the selection extends to the right of the control (true) or highlights the text only (false)."
Dim lStyle As Long
    lStyle = GetWindowLong(tvwOpt.hwnd, GWL_STYLE)
    FullRowSelect = ((lStyle And TVS_FULLROWSELECT) = TVS_FULLROWSELECT)
End Property
Public Property Let FullRowSelect(ByVal bState As Boolean)
Dim lStyle As Long
    lStyle = GetWindowLong(tvwOpt.hwnd, GWL_STYLE)
    If (bState) Then
        lStyle = lStyle Or TVS_FULLROWSELECT
    Else
        lStyle = lStyle And Not TVS_FULLROWSELECT
    End If
    SetWindowLong tvwOpt.hwnd, GWL_STYLE, lStyle
    
    PropertyChanged "FullRowSelect"
End Property
Public Property Get FolderType(ByVal vKey As Variant) As OptionTreeFolderTypeCOnstants
Dim sIcon As String
   sIcon = tvwOpt.Nodes(vKey).Image
   Select Case Left$(sIcon, 6)
   Case "FOLDER"
      If Mid$(sIcon, 7) = "@" Then
         FolderType = optiontreeFolderCustom
      Else
         FolderType = OptionTreeFolder
      End If
   Case "CHECK_"
      FolderType = OptionTreeCheck
   Case Else
      pErr 5
   End Select
   
End Property
Public Property Get Value(ByVal vKey As Variant) As OptionTreeCheckTypes
Dim sIcon As String
   sIcon = tvwOpt.Nodes(vKey).Image
   If (Left$(sIcon, 6) <> "OPTION") And (Left$(sIcon, 6) <> "CHECK_") Then
      pErr 3
   Else
      Select Case Mid$(sIcon, 7, 1)
      Case "Y"
         Value = OptionTreeCheckFull
      Case "N"
         Value = OptionTreeCheckNone
      Case "P"
         Value = OptionTreeCheckPartial
      End Select
   End If
End Property
Public Property Let Value(ByVal vKey As Variant, ByVal eOpt As OptionTreeCheckTypes)
Dim sIcon As String
Dim eCurOpt As OptionTreeCheckTypes
Dim bLocked As Boolean

   If eOpt = OptionTreeCheckPartial Then
      pErr 4
   Else
      sIcon = tvwOpt.Nodes(vKey).Image
      If (Left$(sIcon, 6) <> "OPTION") And (Left$(sIcon, 6) <> "CHECK_") Then
         pErr 3
      Else
         eCurOpt = Value(vKey)
         If (eOpt <> eCurOpt) Then
            If (m_bLocked) Then
               bLocked = True
               m_bLocked = False
            End If
            m_eNodeClickReason = eoptMouseDown
            Select Case eOpt
            Case OptionTreeCheckNone
               If eCurOpt = OptionTreeCheckPartial Then
                  tvwOpt_NodeClick tvwOpt.Nodes(vKey)
               End If
               tvwOpt_NodeClick tvwOpt.Nodes(vKey)
            Case OptionTreeCheckFull
               tvwOpt_NodeClick tvwOpt.Nodes(vKey)
            End Select
            m_eNodeClickReason = 0
            If (bLocked) Then
               m_bLocked = True
            End If
         End If
      End If
   End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the back colour of the control."
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
Dim lColor As Long
Dim iType As Integer
    
    ' Cache backcolor in the user control:
    UserControl.BackColor = oColor
    ilsIcons.BackColor = oColor
    
    ' Set the colour into the TreeView:
    lColor = TranslateColor(oColor)
    SendMessageLong tvwOpt.hwnd, TVM_SETBKCOLOR, 0, lColor
    
    ' Ensure the background to the lines is redrawn:
    iType = tvwOpt.Style
    tvwOpt.Style = 0
    tvwOpt.Style = iType
    
    PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the colour of the text in the control."
    ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal oColor As OLE_COLOR)
Dim lColor As Long
    ' Cache forecolor in the usercontrol:
    UserControl.ForeColor = oColor
    
    ' Set the colour in the TreeView:
    lColor = TranslateColor(oColor)
    SendMessageLong tvwOpt.hwnd, TVM_SETTEXTCOLOR, 0, lColor
    
    ' Request a redraw:
    InvalidateRectAsNull tvwOpt.hwnd, 0, 1
    UpdateWindow tvwOpt.hwnd
    
    PropertyChanged "ForeColor"
End Property
Public Property Get TrackSelect() As Boolean
Attribute TrackSelect.VB_Description = "Gets/sets whether the control will highlight the item that the mouse is over."
    TrackSelect = m_bTrackSelect
End Property
Public Property Let TrackSelect(ByVal bState As Boolean)
    m_bTrackSelect = bState
    TreeViewTrackSelect tvwOpt, m_bTrackSelect
    PropertyChanged "TrackSelect"
End Property

Private Sub TreeViewTrackSelect( _
        ByRef tvwThis As TreeView, _
        Optional ByVal bState = True _
    )
Dim lStyle As Long
Dim hWndTvw As Long
    hWndTvw = tvwThis.hwnd
    lStyle = GetWindowLong(hWndTvw, GWL_STYLE)
    If (bState) Then
        lStyle = lStyle Or TVS_TRACKSELECT
    Else
        lStyle = lStyle And Not TVS_TRACKSELECT
    End If
    SetWindowLong hWndTvw, GWL_STYLE, lStyle
End Sub

Public Sub AddFolderType( _
        ByRef picClosed As StdPicture, _
        Optional ByRef picOpen As StdPicture = Nothing, _
        Optional ByVal Key As Variant _
    )
Attribute AddFolderType.VB_Description = "Adds a type of folder (having your own customised icons) which you can use in the AddFolder method."
Dim vClosedKey As Variant
Dim vOpenKey As Variant

    vClosedKey = "FOLDER@" & Format$(Now, "hhnnss") & "@" & m_iFolderCount + 1 & ":CLOSED"
    ilsIcons.ListImages.Add , vClosedKey, picClosed
    vOpenKey = "FOLDER@" & Format$(Now, "hhnnss") & "@" & m_iFolderCount + 1 & ":OPEN"
    If picOpen Is Nothing Then
        ilsIcons.ListImages.Add , vOpenKey, picClosed
    Else
        ilsIcons.ListImages.Add , vOpenKey, picOpen
    End If
    If (Err.Number = 0) Then
        m_iFolderCount = m_iFolderCount + 1
        ReDim Preserve m_tFolderInfo(1 To m_iFolderCount) As tFolderInfo
        With m_tFolderInfo(m_iFolderCount)
            .vKey = Key
            .vClosedKey = vClosedKey
            .vOpenKey = vOpenKey
        End With
    End If
End Sub
Property Set OptionPicture( _
        ByVal bState As Boolean, _
        ByRef pic As StdPicture _
    )
Attribute OptionPicture.VB_Description = "Gets/sets the icon image used to draw option boxes in the control."
Dim vKey As Variant
    If (bState) Then
        vKey = "OPTIONYES" & Postfix()
    Else
        vKey = "OPTIONNO" & Postfix()
    End If
    ilsIcons.ListImages.Remove vKey
    ilsIcons.ListImages.Add , vKey, pic
End Property
Property Get OptionPicture( _
        ByVal bState As Boolean _
    ) As StdPicture
Dim vKey As Variant
    If (bState) Then
        vKey = "OPTIONYES" & Postfix()
    Else
        vKey = "OPTIONNO" & Postfix()
    End If
    Set OptionPicture = ilsIcons.ListImages(vKey).Picture
End Property
Property Set CheckPicture( _
        ByVal eType As OptionTreeCheckTypes, _
        ByRef pic As StdPicture _
    )
Attribute CheckPicture.VB_Description = "Gets/sets an image to be used to draw check boxes."
Dim vKey As Variant
    Select Case eType
    Case OptionTreeCheckPartial
        vKey = "CHECK_PART" & Postfix()
    Case OptionTreeCheckNone
        vKey = "CHECK_NO" & Postfix()
    Case OptionTreeCheckFull
        vKey = "CHECK_YES" & Postfix()
    End Select
    ilsIcons.ListImages.Remove vKey
    ilsIcons.ListImages.Add , vKey, pic
End Property
Property Get CheckPicture( _
        ByVal eType As OptionTreeCheckTypes _
    ) As StdPicture
Dim vKey As Variant
    Select Case eType
    Case OptionTreeCheckPartial
        vKey = "CHECK_PART" & Postfix()
    Case OptionTreeCheckNone
        vKey = "CHECK_NO" & Postfix()
    Case OptionTreeCheckFull
        vKey = "CHECK_YES" & Postfix()
    End Select
    Set CheckPicture = ilsIcons.ListImages(vKey).Picture
End Property
Property Get Indentation() As Double
Attribute Indentation.VB_Description = "Gets/sets the amount of indentation added for each child level."
    Indentation = tvwOpt.Indentation
End Property
Property Let Indentation(fIndentation As Double)
    If (tvwOpt.Indentation <> fIndentation) Then
        tvwOpt.Indentation = fIndentation
        PropertyChanged "Indentation"
    End If
End Property
Property Get Style() As MSComctlLib.TreeStyleConstants
Attribute Style.VB_Description = "Gets/sets whether pictures, plus/minus buttons and or treelines lines are drawn in the control."
    Style = tvwOpt.Style
End Property
Property Let Style(eStyle As MSComctlLib.TreeStyleConstants)
    If (tvwOpt.Style <> eStyle) Then
        tvwOpt.Style = eStyle
        PropertyChanged "Style"
    End If
End Property
Property Get SelectedNode() As MSComctlLib.Node
   Set SelectedNode = tvwOpt.SelectedItem
End Property
Property Get NodesCollection() As MSComctlLib.Nodes
   Set NodesCollection = tvwOpt.Nodes
End Property
Property Get Nodes(Item As Variant) As MSComctlLib.Node
Attribute Nodes.VB_Description = "Gets a reference to the controls Tree View nodes collection."
    Set Nodes = tvwOpt.Nodes(Item)
End Property
Public Function AddFolder( _
        Optional ByVal Key As Variant, _
        Optional ByRef nodParent As Node = Nothing, _
        Optional ByVal Text As String, _
        Optional ByVal FolderType As OptionTreeFolderTypeCOnstants = OptionTreeFolder, _
        Optional ByVal FolderIconsKey As Variant, _
        Optional ByVal bBold As Boolean = False _
    ) As MSComctlLib.Node
Attribute AddFolder.VB_Description = "Adds a folder to the control."
Dim vIcon As Variant
Dim iIndex As Integer
Dim nodX As Node
    If (FolderType = OptionTreeCheck) Then
        vIcon = "CHECK_YES" & Postfix()
    Else
        If (FolderType = optiontreeFolderCustom) Then
            If Not (IsMissing(FolderIconsKey)) Then
                iIndex = piFindCustomFolderIcon(FolderIconsKey)
                If (iIndex > 0) Then
                    vIcon = m_tFolderInfo(iIndex).vClosedKey
                Else
                    vIcon = "FOLDERCLOSED"
                End If
            End If
        Else
            vIcon = "FOLDERCLOSED"
        End If
    End If
    If Not (nodParent Is Nothing) Then
        Set nodX = tvwOpt.Nodes.Add(nodParent, tvwChild, Key, Text, vIcon)
    Else
        Set nodX = tvwOpt.Nodes.Add(, , Key, Text, vIcon)
    End If
    If (bBold) Then
        nodX.Selected = True
        SelectedNodeIsBold = True
    End If
    Set AddFolder = nodX
End Function
Public Property Get SelectedNodeIsBold() As Boolean
Attribute SelectedNodeIsBold.VB_Description = "Gets/sets whether the selected item should be shown in bold font."
Dim tVI As TVITEM
Dim hItem As Long
    hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    If hItem <> 0 Then
        With tVI
            .hItem = hItem
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            SendMessage tvwOpt.hwnd, TVM_GETITEM, 0&, tVI
            SelectedNodeIsBold = (tVI.State = TVIS_BOLD)
        End With
    End If
End Property
Public Property Let SelectedNodeIsBold(ByVal bState As Boolean)
Dim tVI As TVITEM
Dim hItem As Long
    hItem = SendMessageLong(tvwOpt.hwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
    If hItem <> 0 Then
        With tVI
            .hItem = hItem
            .mask = TVIF_STATE
            .stateMask = TVIS_BOLD
            SendMessage tvwOpt.hwnd, TVM_GETITEM, 0&, tVI
            If ((tVI.State = TVIS_BOLD) <> bState) Then
                If (bState) Then
                    tVI.State = tVI.State Or TVIS_BOLD
                Else
                    tVI.State = tVI.State And Not TVIS_BOLD
                End If
                SendMessage tvwOpt.hwnd, TVM_SETITEM, 0&, tVI
            End If
        End With
    End If
            
End Property
Private Function piFindCustomFolderIcon(vKey As Variant)
Dim i As Integer
    For i = 1 To m_iFolderCount
        If (m_tFolderInfo(i).vKey = vKey) Then
            piFindCustomFolderIcon = i
            Exit For
        End If
    Next i
End Function
Public Function AddCheck( _
        Optional ByVal Key As Variant, _
        Optional ByRef nodParent As Node = Nothing, _
        Optional ByVal Text As String = "", _
        Optional ByVal CheckType As OptionTreeCheckTypes = OptionTreeCheckFull, _
        Optional ByVal bBold As Boolean = False _
    ) As MSComctlLib.Node
Attribute AddCheck.VB_Description = "Adds a checked item to the control."
Dim vIcon As Variant
Dim nodChk As Node
Dim nodX As Node
    If Not (nodParent Is Nothing) Then
      If (nodParent.Children > 0) Then
         Set nodChk = nodParent.FirstSibling
      Else
         Set nodChk = Nothing
      End If
    Else
        If (tvwOpt.Nodes.Count > 0) Then
            Set nodChk = tvwOpt.Nodes(1)
        Else
            vIcon = "CHECK_"
            Select Case CheckType
            Case OptionTreeCheckFull
                vIcon = vIcon & "YES" & Postfix()
            Case OptionTreeCheckNone
                vIcon = vIcon & "NO" & Postfix()
            Case OptionTreeCheckPartial
                vIcon = vIcon & "PART" & Postfix()
            End Select
            Set nodX = tvwOpt.Nodes.Add(, , Key, Text, vIcon)
            If (bBold) Then
                nodX.Selected = True
                SelectedNodeIsBold = True
            End If
            Set AddCheck = nodX
            Exit Function
        End If
    End If
    If Not (pbOptionItemInBranch(nodChk)) Then
        vIcon = "CHECK_"
        Select Case CheckType
        Case OptionTreeCheckFull
            vIcon = vIcon & "YES" & Postfix()
        Case OptionTreeCheckNone
            vIcon = vIcon & "NO" & Postfix()
        Case OptionTreeCheckPartial
            vIcon = vIcon & "PART" & Postfix()
        End Select
        If (nodParent Is Nothing) Then
            Set nodX = tvwOpt.Nodes.Add(, , Key, Text, vIcon)
        Else
            Set nodX = tvwOpt.Nodes.Add(nodParent, tvwChild, Key, Text, vIcon)
        End If
        If (bBold) Then
            nodX.Selected = True
            SelectedNodeIsBold = True
        End If
        Set AddCheck = nodX
    End If
End Function
Public Function AddOption( _
        Optional ByVal Key As Variant, _
        Optional ByRef nodParent As Node = Nothing, _
        Optional ByVal Text As String = "", _
        Optional ByVal bBold As Boolean = False _
    ) As MSComctlLib.Node
Attribute AddOption.VB_Description = "Adds an option to the control."
Dim vIcon As Variant
Dim nodX As Node
    vIcon = "OPTION"
    If Not (nodParent Is Nothing) Then
        If (pbCheckValidForOption(nodParent)) Then
            If (nodParent.Children > 0) Then
                vIcon = vIcon & "NO" & Postfix()
            Else
                vIcon = vIcon & "YES" & Postfix()
            End If
            Set nodX = tvwOpt.Nodes.Add(nodParent, tvwChild, Key, Text, vIcon)
            If (bBold) Then
                nodX.Selected = True
                SelectedNodeIsBold = True
            End If
            Set AddOption = nodX
        End If
    Else
        If (tvwOpt.Nodes.Count > 0) Then
            If Not (pbCheckValidForOption(tvwOpt.Nodes(1))) Then
                Exit Function
            End If
            If (tvwOpt.Nodes(1).Children > 0) Then
                vIcon = vIcon & "NO"
            Else
                vIcon = vIcon & "YES"
            End If
        Else
            vIcon = vIcon & "YES"
        End If
        vIcon = vIcon & Postfix()
        Set nodX = tvwOpt.Nodes.Add(, , Key, Text, vIcon)
        If (bBold) Then
            nodX.Selected = True
            SelectedNodeIsBold = True
        End If
        Set AddOption = nodX
    End If
End Function
Public Sub ExpandAll()
Attribute ExpandAll.VB_Description = "Expands all nodes in the control."
Dim lS As Long
    LockWindowUpdate tvwOpt.hwnd
    For lS = 1 To tvwOpt.Nodes.Count
        If (tvwOpt.Nodes(lS).Children > 0) Then
            If Not (tvwOpt.Nodes(lS).Expanded) Then
                tvwOpt.Nodes(lS).Expanded = True
            End If
        End If
    Next lS
    tvwOpt.Nodes(1).EnsureVisible
    LockWindowUpdate 0
End Sub
Private Sub pErr(lErr As Long)
Dim sErr As String
   Select Case lErr
   Case 1
       sErr = "Option cannot be added to the chosen item."
   Case 2
       sErr = "Option cannot be added to a folder containing check boxes."
   Case 3
      sErr = "Folder nodes do not have a value."
   Case 4
      sErr = "You cannot set the value of a node to 'partial'."
   Case 5
      sErr = "Option nodes do not have a folder type."
   End Select
   Err.Raise lErr + vbObjectError + 1048 + 512, App.EXEName & ".XTreeOpt", sErr
End Sub
Private Function pbOptionItemInBranch( _
        ByVal nod As Node _
    )
Dim lS As Long
   If Not nod Is Nothing Then
      If (Left$(nod.Image, 6) <> "OPTION") Then
          If (nod.Children > 0) Then
              For lS = nod.Child.FirstSibling.Index To nod.Child.LastSibling.Index
                  If (Left$(nod.Image, 6) = "OPTION") Then
                      pbOptionItemInBranch = True
                      Exit For
                  End If
              Next lS
          End If
      Else
          pbOptionItemInBranch = True
      End If
   End If
End Function
Private Function pbCheckValidForOption( _
        ByVal nodParent As Node _
    )
Dim lI As Long, lS As Long
Dim lChildren As Long
Dim sI As String
    If Left$(nodParent.Image, 6) <> "CHECK_" Then
        pbCheckValidForOption = True
        lChildren = nodParent.Children
        If (lChildren > 0) Then
            lS = nodParent.Child.Index
            For lI = lS To lS + lChildren - 1
                sI = Left$(tvwOpt.Nodes(lI).Image, 6)
                If (sI = "CHECK_") Then
                    pErr 2
                    pbCheckValidForOption = False
                    Exit For
                End If
            Next lI
        End If
    Else
        pErr 1
    End If
End Function
Property Set Font(sFnt As StdFont)
Attribute Font.VB_Description = "Gets/sets the font used to draw the items in the control."
    Set UserControl.Font = sFnt
    Set tvwOpt.Font = UserControl.Font
End Property
Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Private Sub pInitialise()

    TreeViewTrackSelect tvwOpt, m_bTrackSelect
    Set tvwOpt.ImageList = ilsIcons
    
    If (UserControl.Ambient.UserMode) Then
        '
    Else
        ' In design mode, show some examples:
        Dim nodX As Node
        Set nodX = AddFolder(, , "Sample Folder", , , True)
        AddOption , nodX, "Sample Option 1"
        AddOption , nodX, "Sample Option 2"
        AddOption , nodX, "Sample Option 3"
        Set nodX = AddFolder(, , "Sample Check Tree", OptionTreeCheck, , True)
        AddCheck , nodX, "Sample Check 1"
        AddCheck , nodX, "Sample Check 2"
        tvwOpt.Nodes(1).Expanded = True
        nodX.Expanded = True
    End If
    
    ' Add default folder type
    ReDim Preserve m_tFolderInfo(1 To 1) As tFolderInfo
    m_tFolderInfo(1).vClosedKey = "FOLDERCLOSED"
    m_tFolderInfo(1).vOpenKey = "FOLDEROPEN"
    m_iFolderCount = 1
End Sub
Private Function piFindFOlderIndex( _
        Optional ByVal vClosedKey As Variant, _
        Optional ByVal vOpenKey As Variant _
    ) As Integer
Dim iItem As Integer
    If (IsMissing(vClosedKey)) Then
        For iItem = 1 To m_iFolderCount
            If (vOpenKey = m_tFolderInfo(iItem).vOpenKey) Then
                piFindFOlderIndex = iItem
            End If
        Next iItem
    Else
        For iItem = 1 To m_iFolderCount
            If (vClosedKey = m_tFolderInfo(iItem).vClosedKey) Then
                piFindFOlderIndex = iItem
            End If
        Next iItem
    End If
End Function

Private Sub tvwOpt_AfterLabelEdit(Cancel As Integer, NewString As String)
   RaiseEvent AfterLabelEdit(Cancel, NewString)
End Sub

Private Sub tvwOpt_BeforeLabelEdit(Cancel As Integer)
    RaiseEvent BeforeLabelEdit(Cancel)
End Sub

Private Sub tvwOpt_Collapse(ByVal Node As Node)
Dim iIndex As Integer
    iIndex = piFindFOlderIndex(, Node.Image)
    If (iIndex > 0) Then
        Node.Image = m_tFolderInfo(iIndex).vClosedKey
    End If
End Sub

Private Sub tvwOpt_DblClick()
    'Debug.Print "DblClick"
End Sub

Private Sub tvwOpt_Expand(ByVal Node As Node)
Dim iIndex As Integer
    iIndex = piFindFOlderIndex(Node.Image)
    If (iIndex > 0) Then
        Node.Image = m_tFolderInfo(iIndex).vOpenKey
    End If
End Sub

Private Sub tvwOpt_KeyDown(KeyCode As Integer, Shift As Integer)
    'Debug.Print "KeyDown", KeyCode
    If (KeyCode = vbKeySpace) Or (KeyCode = vbKeyReturn) Then
        If Not (tvwOpt.SelectedItem Is Nothing) Then
            m_eNodeClickReason = eoptMouseDown
            tvwOpt_NodeClick tvwOpt.SelectedItem
        End If
        KeyCode = 0
    Else
        m_eNodeClickReason = eoptKeyDown
    End If
End Sub

Private Sub tvwOpt_KeyPress(KeyAscii As Integer)
    'Debug.Print KeyAscii
End Sub

Private Sub tvwOpt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Debug.Print "MouseDown"
    m_eNodeClickReason = eoptMouseDown
End Sub

Private Sub tvwOpt_NodeClick(ByVal Node As Node)
Dim lS As Long
Dim sI As String
Dim bAllOff As Boolean
Dim bAllOn As Boolean
Dim sPf As String

   If Not (m_bLocked) Then
    'Debug.Print "NodeClick"
    If (m_eNodeClickReason = eoptMouseDown) Then
        LockWindowUpdate tvwOpt.hwnd
        sPf = Postfix()
        sI = Left$(Node.Image, 6)
        Select Case sI
        Case "OPTION"
            If (Node.Image = "OPTIONNO" & sPf) Then
                Node.Image = "OPTIONYES" & sPf
                For lS = Node.FirstSibling.Index To Node.LastSibling.Index
                    If (tvwOpt.Nodes(lS) <> Node) Then
                        If (tvwOpt.Nodes(lS).Image <> "OPTIONNO" & sPf) Then
                            tvwOpt.Nodes(lS).Image = "OPTIONNO" & sPf
                        End If
                    End If
                Next lS
                RaiseEvent OptionClick(Node)
            End If
        Case "CHECK_"
            If (Node.Image <> "CHECK_YES" & sPf) Then
                ' Set to check full and set others in
                ' the hierarchy accordingly:
                Node.Image = "CHECK_YES" & sPf
                If (Node.Children > 0) Then
                    pRecurseSetChildren Node, "CHECK_YES" & sPf
                End If
                bAllOn = pbCheckForAll(Node, "CHECK_YES" & sPf)
                If (bAllOn) Then
                    pRecurseSetParents Node, "CHECK_YES" & sPf, True
                Else
                    pRecurseSetParents Node, "CHECK_PART" & sPf
                End If
                RaiseEvent CheckClick(Node, OptionTreeCheckFull)
            Else
                Node.Image = "CHECK_NO" & sPf
                If (Node.Children > 0) Then
                    pRecurseSetChildren Node, "CHECK_NO" & sPf
                End If
                bAllOff = pbCheckForAll(Node, "CHECK_NO" & sPf)
                If (bAllOff) Then
                    pRecurseSetParents Node, "CHECK_NO" & sPf, False
                Else
                    pRecurseSetParents Node, "CHECK_PART" & sPf
                End If
                RaiseEvent CheckClick(Node, OptionTreeCheckNone)
            End If
        End Select
        LockWindowUpdate 0
    End If
   End If
End Sub
Private Function pbCheckForAll( _
        ByRef nod As Node, _
        ByVal vIcon As Variant _
    ) As Boolean
Dim lS As Long
    pbCheckForAll = True
    For lS = nod.FirstSibling.Index To nod.LastSibling.Index
        If (tvwOpt.Nodes(lS).Image <> vIcon) Then
            pbCheckForAll = False
            Exit For
        End If
    Next lS

End Function
Private Sub pRecurseSetParents( _
        ByVal Node As MSComctlLib.Node, _
        ByVal vIconKey As Variant, _
        Optional ByVal bOn As Boolean = False _
    )
Dim nodP As Node
Dim lS As Long
Dim vCheck As Variant
Dim bCheck As Boolean
Dim sI As String

    If (Node.Parent Is Nothing) Then
        ' finished
    Else
        Set nodP = Node.Parent
        sI = nodP.Image
        If Left$(sI, 6) = "CHECK_" Then
            nodP.Image = vIconKey
            If (vIconKey = "CHECK_PART" & Postfix()) Then
                pRecurseSetParents nodP, "CHECK_PART" & Postfix()
            Else
                If (bOn) Then
                    vCheck = "CHECK_YES"
                Else
                    vCheck = "CHECK_NO"
                End If
                bCheck = pbCheckForAll(nodP, vCheck & Postfix())
                If (bCheck) Then
                    pRecurseSetParents nodP, vCheck & Postfix()
                Else
                    pRecurseSetParents nodP, "CHECK_PART" & Postfix()
                End If
            End If
        End If
    End If
End Sub
Private Sub pRecurseSetChildren( _
        ByRef nod As Node, _
        ByVal vIconKey As Variant _
    )
Dim nodS As Node
Dim lS As Long
    Set nodS = nod.Child
    For lS = nodS.FirstSibling.Index To nodS.LastSibling.Index
        tvwOpt.Nodes(lS).Image = vIconKey
        If (tvwOpt.Nodes(lS).Children > 0) Then
            pRecurseSetChildren tvwOpt.Nodes(lS), vIconKey
        End If
    Next lS
End Sub

Private Sub UserControl_Initialize()
    m_eIconSet = OptionTreeIconsWin98
End Sub

Private Sub UserControl_InitProperties()
    m_bTrackSelect = True
    Set Font = UserControl.Ambient.Font
    pInitialise
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pInitialise
    Dim sFnt As New StdFont
    With sFnt
        .Name = "MS Sans Serif"
        .Size = 8
    End With
    Set Font = PropBag.ReadProperty("Font", sFnt)
    BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    Indentation = PropBag.ReadProperty("Indentation", 256.25)
    Style = PropBag.ReadProperty("Style", tvwTreelinesPlusMinusPictureText)
    TrackSelect = PropBag.ReadProperty("TrackSelect", True)
    FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    SingleClickExpand = PropBag.ReadProperty("SingleClickExpand", False)
    Enabled = PropBag.ReadProperty("Enabled", True)
    IconSet = PropBag.ReadProperty("IconSet", OptionTreeIconsWin98)
    BackgroundPicture = PropBag.ReadProperty("BackgroundPicture", Nothing)
End Sub

Private Sub UserControl_Resize()
    If (UserControl.Extender.Visible) Then
        If (UserControl.ScaleWidth > 0) And (UserControl.ScaleHeight > 0) Then
            tvwOpt.Move 0, 0, (UserControl.ScaleWidth), (UserControl.ScaleHeight)
        End If
    End If
End Sub


Private Sub UserControl_Terminate()
   Set m_cTVB = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim sFnt As New StdFont
    With sFnt
        .Name = "MS Sans Serif"
        .Size = 8
    End With
    PropBag.WriteProperty "Font", Font, sFnt
    PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
    PropBag.WriteProperty "Indentation", Indentation, 256.25
    PropBag.WriteProperty "Style", Style, tvwTreelinesPlusMinusPictureText
    PropBag.WriteProperty "TrackSelect", TrackSelect, True
    PropBag.WriteProperty "FullRowSelect", FullRowSelect, False
    PropBag.WriteProperty "SingleClickExpand", SingleClickExpand, False
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "IconSet", IconSet, OptionTreeIconsWin98
    PropBag.WriteProperty "BackgroundPicture", BackgroundPicture, Nothing
End Sub
