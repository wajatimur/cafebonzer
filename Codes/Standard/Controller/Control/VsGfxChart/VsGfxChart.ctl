VERSION 5.00
Begin VB.UserControl VsGfxChart 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5580
   ScaleWidth      =   8400
   Begin VB.PictureBox picLegend 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F5F5&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFF0F0&
      ForeColor       =   &H00FF7040&
      Height          =   5430
      Left            =   3360
      ScaleHeight     =   5430
      ScaleWidth      =   2130
      TabIndex        =   1
      Top             =   0
      Width           =   2130
      Begin VB.VScrollBar vsbContainer 
         Height          =   5445
         LargeChange     =   5
         Left            =   1905
         Max             =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F0F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5205
         Left            =   150
         ScaleHeight     =   5205
         ScaleWidth      =   1665
         TabIndex        =   2
         Top             =   0
         Width           =   1665
         Begin VB.Label lblDescription 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   135
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Shape Box 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   0  'Transparent
            Height          =   195
            Index           =   0
            Left            =   75
            Shape           =   5  'Rounded Square
            Top             =   150
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.Label lblSlider 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5430
         Left            =   15
         TabIndex        =   5
         ToolTipText     =   "Display Legend"
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectionInfo 
         Caption         =   "Selection &Information"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLegend 
         Caption         =   "&Display Legend"
      End
   End
   Begin VB.Menu mnuLegend 
      Caption         =   "&Legend"
      Begin VB.Menu mnuLegendHide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "VsGfxChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long

Const CSIDL_MYPICTURES = &H27
Const OFN_ALLOWMULTISELECT As Long = &H200
Const OFN_CREATEPROMPT As Long = &H2000
Const OFN_ENABLEHOOK As Long = &H20
Const OFN_ENABLETEMPLATE As Long = &H40
Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Const OFN_EXPLORER As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST As Long = &H1000
Const OFN_HIDEREADONLY As Long = &H4
Const OFN_LONGNAMES As Long = &H200000
Const OFN_NOCHANGEDIR As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_NOLONGNAMES As Long = &H40000
Const OFN_NONETWORKBUTTON As Long = &H20000
Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Const OFN_NOTESTFILECREATE As Long = &H10000
Const OFN_NOVALIDATE As Long = &H100
Const OFN_OVERWRITEPROMPT As Long = &H2
Const OFN_PATHMUSTEXIST As Long = &H800
Const OFN_READONLY As Long = &H1
Const OFN_SHAREAWARE As Long = &H4000
Const OFN_SHAREFALLTHROUGH As Long = 2
Const OFN_SHAREWARN As Long = 0
Const OFN_SHARENOWARN As Long = 1
Const OFN_SHOWHELP As Long = &H10
Const OFS_MAXPATHNAME As Long = 260

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS

Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Private Type OPENFILENAME
  nStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Private OFN As OPENFILENAME



Private uColumns()        As Double       'Array of column height values
                                          'used to determine hittest feature.

Private uColWidth         As Double       'The calculated width of each column.
Private uRowHeight        As Double       'The calculated height of each column.
Private uTopMargin        As Long         '--------------------------------------
Private uBottomMargin     As Long         'Margins used around the chart content.
Private uLeftMargin       As Long         '
Private uRightMargin      As Long         '--------------------------------------
Private uContentBorder    As Boolean      'Border around the chart content?
Private uSelectable       As Boolean      'Marker indicating whether user can select a column.
Private uHotTracking      As Boolean      'Marker indicating use of hot tracking.
Private uSelectedColumn   As Long         'Marker indicating the selected column.
Private uOldSelection     As Long
Private uDisplayDescript  As Boolean      'Display description when selectable
Private uChartTitle       As String       'Chart title
Private uChartSubTitle    As String       'Chart sub title
Private uDisplayXAxis     As Boolean      'Marker indicating display of x axis
Private uDisplayYAxis     As Boolean      'Marker indicating display of y axis
Private uColorBars        As Boolean      'Marker indicating use of different coloured bars
Private uIntersectMajor   As Single       'Major intersect value
Private uIntersectMinor   As Single       'Minor intersect value
Private uMaxYValue        As Double       'Default maximum y value
Private uXAxisLabel       As String       'Label to be displayed below the X-Axis
Private uYAxisLabel       As String       'Label to be displayed left of the Y-Axis
Private cItems            As Collection   'Collection of chart items

Private offsetX           As Long
Private offsetY           As Long

Private bLegendAdded      As Boolean
Private bLegendClicked    As Boolean
Private bDisplayLegend    As Boolean
Private bResize           As Boolean


Private bProcessingOver   As Boolean      'Marker to speed up mouse over effects.


Public Event ItemClick(cItem As ChartItem)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Function AddItem(cItem As ChartItem) As Boolean
    cItems.Add cItem
End Function

Public Function EditCopy() As Boolean
    Clipboard.SetData UserControl.Image
End Function

Public Property Let MarginTop(lMargin As Long)
    uTopMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
End Property
Public Property Get MarginTop() As Long
    MarginTop = uTopMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginBottom(lMargin As Long)
    uBottomMargin = lMargin * Screen.TwipsPerPixelY
    DrawChart
End Property
Public Property Get MarginBottom() As Long
    MarginBottom = uBottomMargin / Screen.TwipsPerPixelY
End Property

Public Property Let MarginLeft(lMargin As Long)
    uLeftMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
End Property
Public Property Get MarginLeft() As Long
    MarginLeft = uLeftMargin / Screen.TwipsPerPixelX
End Property

Public Property Let MarginRight(lMargin As Long)
    uRightMargin = lMargin * Screen.TwipsPerPixelX
    DrawChart
End Property
Public Property Get MarginRight() As Long
    MarginRight = uRightMargin / Screen.TwipsPerPixelX
End Property

Public Property Let ContentBorder(DisplayBorder As Boolean)
    uContentBorder = DisplayBorder
    DrawChart
End Property
Public Property Get ContentBorder() As Boolean
    ContentBorder = uContentBorder
End Property

Public Property Let Selectable(EnableSelection As Boolean)
    uSelectable = EnableSelection
    DrawChart
End Property
Public Property Get Selectable() As Boolean
    Selectable = uSelectable
End Property

Public Property Let HotTracking(UseHotTracking As Boolean)
    uHotTracking = UseHotTracking
    DrawChart
End Property
Public Property Get HotTracking() As Boolean
    HotTracking = uHotTracking
End Property

Public Property Let SelectedColumn(ColNumber As Long)
    Dim ret As Double
    Dim oItem As ChartItem
    On Error Resume Next
    
    uSelectedColumn = ColNumber
    DrawChart
    
    ret = uColumns(ColNumber)
    If Err.Number Then
        uSelectedColumn = -1
    Else
        oItem = cItems(ColNumber + 1)
        RaiseEvent ItemClick(oItem)
    End If

End Property
Public Property Get SelectedColumn() As Long
    SelectedColumn = uSelectedColumn
End Property

Public Property Let ChartTitle(sTitle As String)
    uChartTitle = sTitle
    DrawChart
End Property
Public Property Get ChartTitle() As String
    ChartTitle = uChartTitle
End Property

Public Property Let ChartSubTitle(sTitle As String)
    uChartSubTitle = sTitle
    DrawChart
End Property
Public Property Get ChartSubTitle() As String
    ChartSubTitle = uChartSubTitle
End Property

Public Property Let IntersectMajor(ISValue As Single)
    uIntersectMajor = ISValue
    DrawChart
End Property
Public Property Get IntersectMajor() As Single
    IntersectMajor = uIntersectMajor
End Property

Public Property Let IntersectMinor(ISValue As Single)
    uIntersectMinor = ISValue
    DrawChart
End Property
Public Property Get IntersectMinor() As Single
    IntersectMinor = uIntersectMinor
End Property

Public Property Let DisplayYAxis(DisplayAxis As Boolean)
    uDisplayYAxis = DisplayAxis
    DrawChart
End Property
Public Property Get DisplayYAxis() As Boolean
    DisplayYAxis = uDisplayYAxis
End Property

Public Property Let DisplayXAxis(DisplayAxis As Boolean)
    uDisplayXAxis = DisplayAxis
    DrawChart
End Property
Public Property Get DisplayXAxis() As Boolean
    DisplayXAxis = uDisplayXAxis
End Property

Public Property Let MaxY(dMax As Double)
    uMaxYValue = dMax
    DrawChart
End Property
Public Property Get MaxY() As Double
    MaxY = uMaxYValue
End Property

Public Property Let SelectionInformation(DisplayInfo As Boolean)
    uDisplayDescript = DisplayInfo
    DrawChart
End Property
Public Property Get SelectionInformation() As Boolean
    SelectionInformation = uDisplayDescript
End Property

Public Property Let AxisLabelY(sCaption As String)
    uYAxisLabel = sCaption
    DrawChart
End Property
Public Property Get AxisLabelY() As String
    AxisLabelY = uYAxisLabel
End Property

Public Property Let AxisLabelX(sCaption As String)
    uXAxisLabel = sCaption
    DrawChart
End Property
Public Property Get AxisLabelX() As String
    AxisLabelX = uXAxisLabel
End Property

Public Property Let BackColor(hColor As OLE_COLOR)
    UserControl.BackColor = hColor
    DrawChart
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let ForeColor(hColor As OLE_COLOR)
    UserControl.ForeColor = hColor
    DrawChart
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ColorBars(bUseColor As Boolean)
    uColorBars = bUseColor
    DrawChart
End Property
Public Property Get ColorBars() As Boolean
    ColorBars = uColorBars
End Property

Private Sub lblDescription_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If uSelectable Then

            uSelectedColumn = Index
            uOldSelection = uSelectedColumn
            
            lScrollvalue = vsbContainer.Value
            
            bLegendClicked = True
            
            DrawChart
            
            bLegendClicked = False
        
            vsbContainer.Value = lScrollvalue
        End If
    End If
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        offsetX = X
        offsetY = Y
        lblInfo.Drag
    Else
        PopupMenu mnuMain
    End If
End Sub

Private Sub mnuRefresh_Click()
    DrawChart
End Sub

Private Sub lblSlider_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetData UserControl.Image
End Sub

Private Sub mnuLegendHide_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend True
    DrawChart
End Sub



Private Sub mnuSaveAs_Click()
   Dim blnReturn As Long
   Dim strBuffer As String
   strBuffer = Space(255)
   blnReturn = SHGetSpecialFolderPath(0, _
      strBuffer, _
      CSIDL_MYPICTURES, _
      False)
      
   strBuffer = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
   
   
   
   Dim sFilters As String
   Dim OFN As OPENFILENAME
   Dim lret As Long
   
  'used after call
   Dim buff As String
   Dim sLname As String
   Dim sSname As String

  'create string of filters for the dialog
   sFilters = "Windows Bitmap" & vbNullChar & _
              "*.bmp" & vbNullChar & vbNullChar
  
   With OFN
      .nStructSize = Len(OFN)
      .hWndOwner = UserControl.hWnd
      .sFilter = sFilters
      .nFilterIndex = 0
      .sFile = "ActiveChart.bmp" & Space$(1024) & _
               vbNullChar & vbNullChar
      .nMaxFile = Len(.sFile)
      .sDefFileExt = "bmp" & vbNullChar & vbNullChar
      .sFileTitle = vbNullChar & Space$(512) & _
                    vbNullChar & vbNullChar
      .nMaxTitle = Len(OFN.sFileTitle)
      .sInitialDir = strBuffer & vbNullChar & vbNullChar
      .sDialogTitle = "VBnet GetSaveFileName Demo"
      .flags = OFS_FILE_SAVE_FLAGS

   End With
   
   
  'call the API
   blnReturn = GetSaveFileName(OFN)
   
   If blnReturn Then
      SavePicture UserControl.Image, OFN.sFile
   End If
End Sub

Private Sub mnuSelectionInfo_Click()
    mnuSelectionInfo.Checked = Not mnuSelectionInfo.Checked
    uDisplayDescript = mnuSelectionInfo.Checked
    DrawChart
End Sub

Private Sub mnuViewLegend_Click()
    mnuViewLegend.Checked = Not mnuViewLegend.Checked
    bDisplayLegend = mnuViewLegend.Checked
    ShowLegend Not (bDisplayLegend)
    DrawChart
End Sub


Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub picLegend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuLegend
    End If
End Sub

Private Sub UserControl_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Left = X - offsetX
    Source.Top = Y - offsetY
End Sub

Private Sub UserControl_Initialize()
    Set cItems = New Collection
End Sub

Private Sub UserControl_InitProperties()
    Dim X As Integer
    Dim oChartItem As ChartItem
    
    uTopMargin = 50 * Screen.TwipsPerPixelY
    uBottomMargin = 55 * Screen.TwipsPerPixelY
    uLeftMargin = 55 * Screen.TwipsPerPixelX
    uRightMargin = 55 * Screen.TwipsPerPixelX
    uContentBorder = True
    uSelectable = False
    uHotTracking = False
    uSelectedColumn = -1
    uOldSelection = -1
    uChartTitle = UserControl.Name
    uChartSubTitle = "YP Electronics Ltd."
    uDisplayYAxis = True
    uDisplayXAxis = True
    uColorBars = False
    uIntersectMajor = 10
    uIntersectMinor = 2
    uMaxYValue = 100
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x1 As Single
    Dim oItem As ChartItem
    
    If Button = vbLeftButton Then
        x1 = (uColWidth)
        
        On Error GoTo TrackExit
        
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (x1)) <= Y) And uSelectable Then
            If Not bProcessingOver Then
                bProcessingOver = True
                uSelectedColumn = (X - uLeftMargin) \ (x1)
                If Not uSelectedColumn = uOldSelection Then
                    Cls
                    DrawChart
                    uOldSelection = uSelectedColumn
                    oItem = cItems(uSelectedColumn + 1)
                    RaiseEvent ItemClick(oItem)
                End If
    
                bProcessingOver = False
             End If
        End If
    ElseIf Button = vbRightButton Then
        mnuSelectionInfo.Visible = False
        If uSelectable Then
            mnuSelectionInfo.Visible = True
            mnuSeperator.Visible = True
        End If
        PopupMenu mnuMain
    End If
        
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
TrackExit:
    Exit Sub
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x1 As Long
    Dim oItem As ChartItem
    x1 = (uColWidth)
    
    On Error GoTo TrackExit
    
    If uHotTracking Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (x1)) <= Y) And uSelectable Then
           If Not bProcessingOver Then
               bProcessingOver = True
               uSelectedColumn = (X - uLeftMargin) \ (x1)
               If Not uSelectedColumn = uOldSelection Then
                   Cls
                   DrawChart
                   uOldSelection = uSelectedColumn
                   oItem = cItems(uSelectedColumn + 1)
                   RaiseEvent ItemClick(oItem)
               End If
   
               bProcessingOver = False
           End If
        End If
    ElseIf Button = vbLeftButton Then
        If (Y <= UserControl.ScaleHeight - uBottomMargin) And (uColumns((X - uLeftMargin) \ (x1)) <= Y) And uSelectable Then
           If Not bProcessingOver Then
               bProcessingOver = True
               uSelectedColumn = (X - uLeftMargin) \ (x1)
               If Not uSelectedColumn = uOldSelection Then
                   Cls
                   DrawChart
                   uOldSelection = uSelectedColumn
                   oItem = cItems(uSelectedColumn + 1)
                   RaiseEvent ItemClick(oItem)
               End If
   

       
               bProcessingOver = False
           End If
        End If
    End If

TrackExit:

    Exit Sub
End Sub

Public Sub Refresh()
    DrawChart
End Sub

Public Sub Clear()
    Dim X As Integer
    
    Set cItems = Nothing
    Set cItems = New Collection
    If bLegendAdded Then
        ClearLegendItems
    End If
    DrawChart
End Sub

Public Sub DrawChart()
    Dim CurrentColor    As Integer
    Dim iCols           As Integer
    Dim X               As Integer
    Dim x1              As Double
    Dim x2              As Double
    Dim y1              As Double
    Dim y2              As Double
    Dim xTemp           As Double
    Dim yTemp           As Double
    Dim sDescription    As String
    Dim oChartItem      As ChartItem
        
    If uIntersectMajor = 0 Then uIntersectMajor = 10
    If uIntersectMinor = 0 Then uIntersectMinor = 2
    
    lblInfo.ForeColor = UserControl.ForeColor
    lblDescription(0).ForeColor = UserControl.ForeColor
    
    iCols = cItems.Count
    
    mnuSelectionInfo.Checked = uDisplayDescript
    lblInfo.Visible = False
    If uDisplayDescript And uSelectedColumn > -1 Then lblInfo.Visible = True
    
    
    'Kill existing legend
    If bDisplayLegend Then
        vsbContainer.Visible = False
        picContainer.Visible = False
    End If
    
    If Not bResize Then ClearLegendItems
    
    uRowHeight = uMaxYValue
    For X = 1 To cItems.Count
        oChartItem = cItems(X)
        If uRowHeight - CDbl(oChartItem.Value) < 0 Then uRowHeight = CDbl(oChartItem.Value)
    Next X
    
    If uRowHeight = 0 Then uRowHeight = 0.001
    
    If uMaxYValue < uRowHeight Then uMaxYValue = uRowHeight
    
    uRowHeight = ((UserControl.ScaleHeight - (uTopMargin + uBottomMargin)) / uRowHeight)
    If iCols Then uColWidth = ((UserControl.ScaleWidth - (uLeftMargin + uRightMargin)) / iCols)
    
    'UserControl.AutoRedraw = True
    UserControl.Cls

    If iCols Then ReDim uColumns(iCols - 1)

    On Error Resume Next
    'Intersect lines
    
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uChartTitle) / 2)
    UserControl.CurrentY = 0
    UserControl.FontBold = True
    UserControl.Print uChartTitle
    UserControl.FontBold = False
        
    UserControl.FontSize = UserControl.FontSize - 2
    UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uChartSubTitle) / 2)
    UserControl.Print uChartSubTitle
    UserControl.FontSize = UserControl.FontSize + 2
    
    If uDisplayYAxis Then
        For X = 0 To uMaxYValue
            x1 = uLeftMargin + (2 * Screen.TwipsPerPixelX): x2 = UserControl.ScaleWidth - uRightMargin
            y1 = (UserControl.ScaleHeight - uBottomMargin) - (X * uRowHeight)
            If (X) Mod uIntersectMajor = 0 Then
                UserControl.Line (x1, y1)-(x2, y1), vbBlue
                
                UserControl.FontSize = UserControl.FontSize - 2
                UserControl.CurrentX = uLeftMargin - UserControl.TextWidth(X) - (5 * Screen.TwipsPerPixelX)
                UserControl.CurrentY = y1 - (UserControl.TextHeight("0") / 2)
                UserControl.Print (X)
                UserControl.FontSize = UserControl.FontSize + 2
                
            ElseIf (uMaxYValue - X) Mod uIntersectMinor = 0 Then
                UserControl.Line (x1, y1)-(x2, y1), &HFFF0F0
            End If
        Next X
    End If

    On Error GoTo 0
    If uContentBorder Then UserControl.Line (uLeftMargin, uTopMargin)-(UserControl.ScaleWidth - uRightMargin, UserControl.ScaleHeight - uBottomMargin), vbBlack, B
    
    
    For X = 0 To cItems.Count - 1
        oChartItem = cItems(X + 1)
        
        x1 = (X * uColWidth) + uLeftMargin + (2 * Screen.TwipsPerPixelX)
        x2 = x1 + uColWidth - (2 * Screen.TwipsPerPixelX)
        y1 = (UserControl.ScaleHeight - uBottomMargin) - (CDbl(oChartItem.Value) * uRowHeight)
        y2 = UserControl.ScaleHeight - uBottomMargin
                
        uColumns(X) = y1
                           
        'Selected bar outline
        If X = uSelectedColumn And uSelectable Then
            UserControl.Line (x1 + 1, y1)-(x2 - 1, y2), vbBlue, BF
            UserControl.Line (x1 + 1, y1)-(x2 - 1, y2), vbRed, B
            
            sDescription = "Value: " & oChartItem.Value
            If Len(oChartItem.SelectedDescription) Then sDescription = "Description: " & oChartItem.SelectedDescription & vbCrLf & sDescription

            xTemp = UserControl.ScaleWidth - uRightMargin - UserControl.TextWidth(sDescription) - 5 * Screen.TwipsPerPixelX
            yTemp = uTopMargin + Screen.TwipsPerPixelY
                          
            'Add Legend item
            If Not bResize Then AddLegendItem oChartItem.SelectedDescription, vbBlue
                                                      
            If uDisplayDescript Then
                lblInfo.Visible = False
                lblInfo = sDescription
                lblInfo.Width = UserControl.TextWidth(sDescription) + 5 * Screen.TwipsPerPixelX
                lblInfo.Height = UserControl.TextHeight(sDescription) * 1.2
                lblInfo.Visible = True
            End If
        Else
            UserControl.Line (x1 + 1, y1)-(x2 - 1, y2), IIf(uColorBars, QBColor(CurrentColor), vbRed), BF
            'Add Legend item
            If Not bResize Then AddLegendItem oChartItem.SelectedDescription, IIf(uColorBars, QBColor(CurrentColor), vbRed)
            
            CurrentColor = CurrentColor + 1
            If CurrentColor >= 15 Then CurrentColor = 0
        End If
        
        If uDisplayXAxis Then
            UserControl.FontSize = UserControl.FontSize - 1
            
            xTemp = (((x2 - x1) / 2) + x1) / Screen.TwipsPerPixelX
            yTemp = (UserControl.ScaleHeight - uBottomMargin + UserControl.TextWidth(oChartItem.XAxisDescription) / 1.25) / Screen.TwipsPerPixelY
            
            PrintRotText UserControl.hDC, oChartItem.XAxisDescription, xTemp, yTemp, 270
            
            UserControl.Line (x1 - 1 * Screen.TwipsPerPixelX, y2)-(x1 - 1 * Screen.TwipsPerPixelX, y2 + UserControl.TextHeight(oChartItem.XAxisDescription) / 2), vbBlack
            UserControl.FontSize = UserControl.FontSize + 1
        End If
        
    Next X

    'Print the x axis label
    If Len(uXAxisLabel) Then
        UserControl.FontSize = UserControl.FontSize - 1
        UserControl.CurrentY = UserControl.ScaleHeight - UserControl.TextHeight(uXAxisLabel) * 1.5
        UserControl.CurrentX = (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(uXAxisLabel) / 2)
        UserControl.Print uXAxisLabel
        UserControl.FontSize = UserControl.FontSize + 1
    End If
    
    'Print the y axis label
    If Len(uYAxisLabel) Then
        UserControl.FontSize = UserControl.FontSize - 1
        PrintRotText UserControl.hDC, uYAxisLabel, UserControl.TextHeight(uYAxisLabel) / Screen.TwipsPerPixelX, UserControl.ScaleHeight / 2 / Screen.TwipsPerPixelY, 90
        UserControl.FontSize = UserControl.FontSize + 1
    End If

    If bDisplayLegend Then
        If uSelectable And uSelectedColumn > -1 Then
            Dim perScreen As Integer
            Dim scrollValue As Integer
                        
            perScreen = Abs((picLegend.ScaleHeight / ((Box(0).Height + (10 * Screen.TwipsPerPixelY)))) - 1)
                        
            If (uSelectedColumn + 1) > perScreen Then
                scrollValue = ((uSelectedColumn + 1) * ((Box(0).Height / Screen.TwipsPerPixelY) + 10)) - (Box(perScreen).Top / Screen.TwipsPerPixelY)
                If scrollValue > vsbContainer.Max Then scrollValue = vsbContainer.Max
                vsbContainer.Value = scrollValue
            Else
                vsbContainer.Value = 0
            End If
                        
            picContainer.Cls
            picContainer.Line ((Box(uSelectedColumn).Left - 3 * Screen.TwipsPerPixelX), (Box(uSelectedColumn).Top - 3 * Screen.TwipsPerPixelY))-(lblDescription(uSelectedColumn).Left + lblDescription(uSelectedColumn).Width + 2 * Screen.TwipsPerPixelX, Box(uSelectedColumn).Top + Box(uSelectedColumn).Height + 2 * Screen.TwipsPerPixelY), vbBlue, B
        End If
        picContainer.Visible = True
    End If
End Sub

Public Function ShowLegend(Optional bHidden As Boolean = False)
    lblSlider.Height = picLegend.ScaleHeight
    picLegend.Line (0, 0)-(picLegend.ScaleWidth - Screen.TwipsPerPixelX, picLegend.ScaleHeight - Screen.TwipsPerPixelY), &HFFE0E0, B
    
    If bHidden Then bDisplayLegend = False Else bDisplayLegend = True
    
    If bDisplayLegend Then
        uRightMargin = uRightMargin + picLegend.ScaleWidth
        picLegend.Move UserControl.ScaleWidth - picLegend.Width + Screen.TwipsPerPixelX, 0, picLegend.Width, UserControl.ScaleHeight
        lblSlider = Chr(187)
    Else
        uRightMargin = uRightMargin - picLegend.Width
        picLegend.Move UserControl.ScaleWidth - lblSlider.Width
        lblSlider = Chr(171)
    End If
End Function

Private Sub AddLegendItem(sDescription As String, Colour As OLE_COLOR)
    Dim X As Integer
    Dim ShortDescript As String
    
    ShortDescript = sDescription
    If Len(ShortDescript) > 17 Then ShortDescript = Left(ShortDescript, 15) & ".."
    
    If bLegendAdded Then
        X = Box.Count
        Load Box(X)
        Load lblDescription(X)
        
        Box(X).BackColor = Colour
        Box(X).Top = Box(X - 1).Top + Box(X - 1).Height + 10 * Screen.TwipsPerPixelY
        lblDescription(X).Top = Box(X).Top
                
        lblDescription(X) = ShortDescript
        lblDescription(X).ToolTipText = sDescription
    Else
        X = 0
        Box(X).BackColor = Colour
                
        lblDescription(X) = ShortDescript
        lblDescription(X).ToolTipText = sDescription
        bLegendAdded = True
    End If
    
    Box(X).Visible = True
    lblDescription(X).Visible = True
            
    picContainer.Height = ((Box(0).Height + (10 * Screen.TwipsPerPixelY)) * Box.Count - 1) + 10 * Screen.TwipsPerPixelY
    If picContainer.ScaleHeight > picLegend.ScaleHeight Then
        vsbContainer.Max = (picContainer.ScaleHeight / Screen.TwipsPerPixelY) - (picLegend.ScaleHeight / Screen.TwipsPerPixelY)
        If Not vsbContainer.Visible Then vsbContainer.Visible = True
    Else
        vsbContainer.Visible = False
    End If
End Sub

Private Sub ClearLegendItems()
    Dim X As Integer
    
    On Error Resume Next    'we are expecting an error for item 1
    
    If bLegendAdded Then
        bLegendAdded = False
        
        For X = 1 To Box.Count
            Unload Box(X)
            Unload lblDescription(X)
            vsbContainer.Value = 0
            Box(0).Visible = False
            lblDescription(0).Visible = False
        Next X
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        uTopMargin = .ReadProperty("uTopMargin")
        uBottomMargin = .ReadProperty("uBottomMargin")
        uLeftMargin = .ReadProperty("uLeftMargin")
        uRightMargin = .ReadProperty("uRightMargin")
        uContentBorder = .ReadProperty("uContentBorder")
        uSelectable = .ReadProperty("uSelectable", False)
        uHotTracking = .ReadProperty("uHotTracking", False)
        uSelectedColumn = .ReadProperty("uSelectedColumn", -1)
        uChartTitle = .ReadProperty("uChartTitle", UserControl.Name)
        uChartSubTitle = .ReadProperty("uChartSubTitle", uChartSubTitle)
        uDisplayYAxis = .ReadProperty("uDisplayXAxis", uDisplayXAxis)
        uDisplayXAxis = .ReadProperty("uDisplayYAxis", uDisplayYAxis)
        uColorBars = .ReadProperty("uColorBars", False)
        uIntersectMajor = .ReadProperty("uIntersectMajor", 10)
        uIntersectMinor = .ReadProperty("uIntersectMinor", 2)
        uMaxYValue = .ReadProperty("uMaxYValue", 100)
        uDisplayDescript = .ReadProperty("uDisplayDescript", False)
        uXAxisLabel = .ReadProperty("uXAxisLabel")
        uYAxisLabel = .ReadProperty("uYAxisLabel")
        UserControl.BackColor = .ReadProperty("BackColor")
        UserControl.ForeColor = .ReadProperty("ForeColor")
        uOldSelection = -1
    End With
End Sub

Private Sub UserControl_Resize()
    If bDisplayLegend Then
        picLegend.Left = UserControl.ScaleWidth - picLegend.Width
    Else
        picLegend.Left = UserControl.ScaleWidth - lblSlider.Width
    End If
    picLegend.Height = UserControl.ScaleHeight
    vsbContainer.Height = picLegend.ScaleHeight
    lblSlider.Height = picLegend.ScaleHeight

    bResize = True
    DrawChart
    bResize = False


End Sub

Private Sub UserControl_Show()
    DrawChart
End Sub

Private Sub UserControl_Terminate()
    Set cItems = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "uTopMargin", uTopMargin
        .WriteProperty "uBottomMargin", uBottomMargin
        .WriteProperty "uLeftMargin", uLeftMargin
        .WriteProperty "uRightMargin", uRightMargin
        .WriteProperty "uContentBorder", uContentBorder
        .WriteProperty "uSelectable", uSelectable
        .WriteProperty "uHotTracking", uHotTracking
        .WriteProperty "uSelectedColumn", uSelectedColumn
        .WriteProperty "uChartTitle", uChartTitle
        .WriteProperty "uChartSubTitle", uChartSubTitle
        .WriteProperty "uDisplayXAxis", uDisplayXAxis
        .WriteProperty "uDisplayYAxis", uDisplayYAxis
        .WriteProperty "uColorBars", uColorBars
        .WriteProperty "uIntersectMajor", uIntersectMajor
        .WriteProperty "uIntersectMinor", uIntersectMinor
        .WriteProperty "uMaxYValue", uMaxYValue
        .WriteProperty "uDisplayDescript", uDisplayDescript
        .WriteProperty "uXAxisLabel", uXAxisLabel
        .WriteProperty "uYAxislabel", uYAxisLabel
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "ForeColor", UserControl.ForeColor
    End With
End Sub

Private Sub vsbContainer_Change()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsbContainer_Scroll()
    picContainer.Top = -vsbContainer.Value * Screen.TwipsPerPixelY
End Sub
