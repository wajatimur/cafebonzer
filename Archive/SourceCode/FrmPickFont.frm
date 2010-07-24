VERSION 5.00
Object = "{1FB2138E-EE1B-11D6-9361-B0DA59D02E57}#1.0#0"; "NAXCLRPICK.OCX"
Begin VB.Form FrmPickFont 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPickFont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox FntType 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "FrmPickFont.frx":000C
      Left            =   2835
      List            =   "FrmPickFont.frx":001C
      TabIndex        =   7
      Top             =   870
      Width           =   1320
   End
   Begin VB.ListBox FntSize 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "FrmPickFont.frx":0044
      Left            =   2835
      List            =   "FrmPickFont.frx":0054
      TabIndex        =   6
      Top             =   75
      Width           =   1320
   End
   Begin CafeBonzerAG.uLine3D uLine3D1 
      Height          =   45
      Left            =   30
      TabIndex        =   5
      Top             =   2100
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin ClrPckr.ColorPicker FntClrPick 
      Height          =   285
      Left            =   2850
      TabIndex        =   2
      Top             =   1695
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      DefaultColor    =   0
      Value           =   0
      Appearance      =   0
      BackColor       =   12632256
   End
   Begin VB.TextBox FntSample 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   " AaBbCcDdEeFfGgHhIiJjKkLlNnOoPpQqRrSsTsUuVvWwXxYyZz"
      Top             =   2205
      Width           =   3120
   End
   Begin VB.ListBox FntList 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "FrmPickFont.frx":0065
      Left            =   60
      List            =   "FrmPickFont.frx":0067
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1950
   End
   Begin CafeBonzerAG.chameleonButton btnFntPick1 
      Height          =   390
      Left            =   3765
      TabIndex        =   3
      ToolTipText     =   "Security Settings"
      Top             =   2205
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
      BTYPE           =   7
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPickFont.frx":0069
      PICN            =   "FrmPickFont.frx":0085
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CafeBonzerAG.chameleonButton btnFntPick2 
      Height          =   390
      Left            =   3270
      TabIndex        =   4
      ToolTipText     =   "General Settings"
      Top             =   2205
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   688
      BTYPE           =   7
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmPickFont.frx":061F
      PICN            =   "FrmPickFont.frx":063B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   4305
      Y1              =   2655
      Y2              =   2655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   4270
      X2              =   4270
      Y1              =   0
      Y2              =   2670
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   15
      X2              =   4275
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   2655
      Y2              =   -15
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Colour"
      Height          =   225
      Left            =   2070
      TabIndex        =   10
      Top             =   1695
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   225
      Left            =   2040
      TabIndex        =   9
      Top             =   855
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   225
      Left            =   2025
      TabIndex        =   8
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "FrmPickFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private s_DefFaceName As String
Private l_DefFontSize As Long
Private l_DefFontColor As Long

Private Sub btnFntPick1_Click()
    SetSave "ticker.fontname", FntList.Text
    SetSave "ticker.fontsize", FntSize.Text
    SetSave "ticker.fontcolor", FntClrPick.Color
    
    FrmTicker.picTicker.FontName = FntList.Text
    FrmTicker.picTicker.FontSize = FntSize.Text
    FrmTicker.picTicker.ForeColor = FntClrPick.Color
    FrmMain.AprClrPick(0).Color = FntClrPick.Color
    FrmMain.AprFntChoose.Caption = GetShortStr(FntList.Text)

    Unload Me
End Sub


Private Sub btnFntPick2_Click()
    Unload Me
End Sub


Public Sub PopFontPicker(ObjectPlace As Object, FontName As String, FontSize As Long, FontColour As Long)
    s_DefFaceName = FontName
    l_DefFontSize = FontSize
    l_DefFontColor = FontColour
    Me.Left = FrmMain.Left + ObjectPlace.Left + 30
    Me.Top = FrmMain.Top + FrmMain.Page(0).Top + ObjectPlace.Top + (ObjectPlace.Height * 2) + 40
    Me.Show vbModal
End Sub


Private Sub FntClrPick_Click()
    Call FntList_Click
End Sub

Private Sub FntList_Click()
    On Error Resume Next
    FntSample.FontName = FntList.Text
    FntSample.FontSize = FntSize.Text
    FntSample.ForeColor = FntClrPick.Color
    
    If FntType = "Regular" Then
        FntSample.FontBold = False
        FntSample.FontItalic = False
    ElseIf FntType = "Italic" Then
        FntSample.FontBold = False
        FntSample.FontItalic = True
    ElseIf FntType = "Bold" Then
        FntSample.FontBold = True
        FntSample.FontItalic = False
    ElseIf FntType = "Bold Italic" Then
        FntSample.FontBold = True
        FntSample.FontItalic = True
    End If
End Sub

Private Sub FntSize_Click()
    Call FntList_Click
End Sub

Private Sub FntType_Click()
    Call FntList_Click
End Sub


Private Sub Form_Load()
    Dim f As Long

    EnumFonts Me.hdc, vbNullString, AddressOf EnumFontProc, 0
    
    FntSize.Selected(0) = True
    FntType.Selected(0) = True
    
    For f = 0 To FntList.ListCount - 1
        If InStr(1, FntList.List(f), s_DefFaceName) > 0 Then
            FntList.Selected(f) = True
        End If
    Next f
    
    For f = 0 To FntSize.ListCount - 1
        If FntSize.List(f) = l_DefFontSize Then
            FntSize.Selected(f) = True
        End If
    Next f
    
    FntClrPick.Color = l_DefFontColor
End Sub


Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

