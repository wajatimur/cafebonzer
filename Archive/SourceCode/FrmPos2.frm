VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "TOOLBAR2.OCX"
Begin VB.Form FrmPos2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cafebonzer - Point Of Sales"
   ClientHeight    =   5295
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   8715
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPos2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListCat 
      Height          =   4740
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   8361
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin AIFCmp1.asxToolbar AsxD 
      Height          =   480
      Left            =   4860
      Top             =   4410
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   847
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   1
      ButtonEnabled1  =   0   'False
      ButtonCaption1  =   "Add to user"
      ButtonKey1      =   "add"
      ButtonPicture1  =   "FrmPos2.frx":000C
      ButtonToolTipText1=   "Add to user"
   End
   Begin VB.ComboBox Terminal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   4
      Top             =   4515
      Width           =   2385
   End
   Begin AIFCmp1.asxToolbar asx1 
      Height          =   780
      Left            =   4860
      Top             =   2130
      Width           =   2190
      _ExtentX        =   3916
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   3
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "1"
      ButtonPicture1  =   "FrmPos2.frx":035E
      ButtonToolTipText1=   "1"
      ButtonKey2      =   "2"
      ButtonPicture2  =   "FrmPos2.frx":0FB0
      ButtonToolTipText2=   "2"
      ButtonKey3      =   "3"
      ButtonPicture3  =   "FrmPos2.frx":1C02
      ButtonToolTipText3=   "3"
   End
   Begin MSComctlLib.ListView List 
      Height          =   4335
      Left            =   2295
      TabIndex        =   0
      Top             =   90
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Perkara"
         Object.Width           =   3881
      EndProperty
   End
   Begin MSComctlLib.ListView PosItem 
      Height          =   1245
      Left            =   4755
      TabIndex        =   2
      Top             =   825
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   2196
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8421504
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Perkara"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Harga"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kuantiti"
         Object.Width           =   1587
      EndProperty
   End
   Begin AIFCmp1.asxToolbar asx2 
      Height          =   780
      Left            =   4860
      Top             =   2880
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   3
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "4"
      ButtonPicture1  =   "FrmPos2.frx":2854
      ButtonToolTipText1=   "4"
      ButtonKey2      =   "5"
      ButtonPicture2  =   "FrmPos2.frx":34A6
      ButtonToolTipText2=   "5"
      ButtonKey3      =   "6"
      ButtonPicture3  =   "FrmPos2.frx":40F8
      ButtonToolTipText3=   "6"
   End
   Begin AIFCmp1.asxToolbar asx3 
      Height          =   780
      Left            =   4860
      Top             =   3630
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   3
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "7"
      ButtonPicture1  =   "FrmPos2.frx":4D4A
      ButtonToolTipText1=   "7"
      ButtonKey2      =   "8"
      ButtonPicture2  =   "FrmPos2.frx":599C
      ButtonToolTipText2=   "8"
      ButtonKey3      =   "9"
      ButtonPicture3  =   "FrmPos2.frx":65EE
      ButtonToolTipText3=   "9"
   End
   Begin AIFCmp1.asxToolbar asr1 
      Height          =   780
      Left            =   7065
      Top             =   2115
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   2
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "0"
      ButtonPicture1  =   "FrmPos2.frx":7240
      ButtonToolTipText1=   "0"
      ButtonKey2      =   "Cancel"
      ButtonPicture2  =   "FrmPos2.frx":7E92
      ButtonToolTipText2=   "Cancel"
   End
   Begin AIFCmp1.asxToolbar asr2 
      Height          =   780
      Left            =   7065
      Top             =   2865
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   2
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "ca"
      ButtonPicture1  =   "FrmPos2.frx":8AE4
      ButtonToolTipText1=   "Cancel All"
      ButtonKey2      =   "quantity"
      ButtonPicture2  =   "FrmPos2.frx":9736
      ButtonToolTipText2=   "Quantity"
   End
   Begin AIFCmp1.asxToolbar asr3 
      Height          =   780
      Left            =   7065
      Top             =   3615
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1376
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SF Digital Readout"
         Size            =   27.75
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   2
      CaptionAlignment=   1
      AutoSize        =   -1  'True
      ButtonKey1      =   "add"
      ButtonPicture1  =   "FrmPos2.frx":A388
      ButtonToolTipText1=   "Add Item"
      ButtonKey2      =   "total"
      ButtonPicture2  =   "FrmPos2.frx":AC62
      ButtonToolTipText2=   "Total"
   End
   Begin AIFCmp1.asxToolbar AsxC 
      Height          =   495
      Left            =   8085
      Top             =   4410
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   873
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   1
      CaptionOptions  =   0
      ButtonCaption1  =   "Close"
      ButtonKey1      =   "add"
      ButtonPicture1  =   "FrmPos2.frx":B8B4
      ButtonToolTipText1=   "Close Point Of Sales"
   End
   Begin AIFCmp1.asxToolbar AsxE 
      Height          =   480
      Left            =   6330
      Top             =   4410
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   847
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ButtonCount     =   1
      ButtonEnabled1  =   0   'False
      ButtonCaption1  =   "Add to record"
      ButtonKey1      =   "add"
      ButtonPicture1  =   "FrmPos2.frx":BC06
      ButtonToolTipText1=   "Add to record"
   End
   Begin VB.Label info 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please select item..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   4950
      Width           =   8625
   End
   Begin VB.Label lcd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "SF Digital Readout"
         Size            =   30
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   690
      Left            =   4755
      TabIndex        =   1
      Top             =   90
      Width           =   3885
   End
End
Attribute VB_Name = "FrmPos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CQty As Boolean
Dim CTotal As Boolean
Dim Total As Double
Dim SelNum As Integer

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' asx button right 1
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub asr1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        SelNum = 0
        Call ActionProc
     Case 2
        If PosItem.ListItems.Count = 0 Then Exit Sub
        If PosItem.SelectedItem.Text = "" Then Exit Sub
        
        PosItem.ListItems.Remove PosItem.SelectedItem.Index
    End Select
End Sub
Private Sub asr1_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "0"
     Case 2
        info = "Cancel current item"
    End Select
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' asx button right 2
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub asr2_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        'reset variable
        CQty = False
        CTotal = False
        Total = 0
        'reset control
        List.ListItems.Clear
        PosItem.ListItems.Clear

        lcd = "0.00"
        info = "Please select item..."
     Case 2
        If PosItem.ListItems.Count = 0 Then info = "Please select item on the left list...": Exit Sub
        If CQty = True Then
            pr = CDbl(PosItem.SelectedItem.SubItems(1))
            qt = CDbl(PosItem.SelectedItem.SubItems(2))
            lcd.Font = "Endless Showroom"
            lcd = Format(pr * qt, "#0.00")
            Disable1 False
            CQty = False
            Exit Sub
        End If
        
        lcd = "QUANTITY ?"
        Disable1 True
        CQty = True
    End Select
End Sub
Private Sub asr2_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "Cancel all"
     Case 2
        info = "Enter new quantity"
    End Select
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' asx button right 3
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub asr3_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        FrmPos.Show
     Case 2
        CTotal = True
        ActionProc
    End Select
End Sub
Private Sub asr3_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "Add new item"
     Case 2
        info = "Count total"
    End Select
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' button 1 - 3
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Asx1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        SelNum = 1
     Case 2
        SelNum = 2
     Case 3
        SelNum = 3
    End Select
    Call ActionProc
End Sub
Private Sub asx1_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "1"
     Case 2
        info = "2"
     Case 3
        info = "3"
    End Select
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' button 4 - 6
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub asx2_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        SelNum = 4
     Case 2
        SelNum = 5
     Case 3
        SelNum = 6
    End Select
    Call ActionProc
End Sub
Private Sub asx2_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "4"
     Case 2
        info = "5"
     Case 3
        info = "6"
    End Select
End Sub


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' button 7 - 9
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub asx3_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        SelNum = 7
     Case 2
        SelNum = 8
     Case 3
        SelNum = 9
    End Select
Call ActionProc
End Sub
Private Sub asx3_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Select Case ButtonIndex
     Case 1
        info = "7"
     Case 2
        info = "8"
     Case 3
        info = "9"
    End Select
End Sub


Private Sub AsxC_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    Unload FrmPos2
End Sub

Private Sub AsxD_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    'jika terminal = "" keluar
    If Terminal = "" Then Exit Sub
    'periksa kewujudan terminal dalam list
    If CekDuplicate(Terminal) = False Then Exit Sub
    
    PosLog Total
    idx = GetAgentIndexB(Terminal)
    FrmMain.Lv1.ListItems(idx).SubItems(7) = Format(Total, "#0.00")
        
    CQty = False
    CTotal = False
    Total = 0
    
    List.ListItems.Clear
    PosItem.ListItems.Clear
    lcd = "0.00"
    AsxD.ButtonEnabled(1) = False
    AsxE.ButtonEnabled(1) = False
    
    Unload FrmPos2
End Sub

Private Sub AsxE_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
    PosLog Total
    
    CQty = False
    CTotal = False
    Total = 0
    
    List.ListItems.Clear
    PosItem.ListItems.Clear
    lcd = "0.00"
    AsxD.ButtonEnabled(1) = False
    AsxE.ButtonEnabled(1) = False
    
    Unload FrmPos2
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Form Entry Points
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Form_Load()
    If AgentCount > 0 Then
        If SelSubItm(5) = VS(3) Then Terminal.Text = SelText
        For t = 1 To AgentCount
            If FrmMain.Lv1.ListItems(t).SubItems(5) = VS(3) Then
                Terminal.AddItem FrmMain.Lv1.ListItems(t)
            End If
        Next t
    Else
        AsxD.Enabled = False
    End If
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub List_DblClick()
    Dim pItm As ListItem
    Dim SelItm
    
    If List.ListItems.Count = 0 Then Exit Sub
    SelItm = List.SelectedItem.Text
            
    For d = 1 To PosItem.ListItems.Count
        Set pItm = PosItem.ListItems(d)
        If SelItm = pItm.Text Then
            pItm.SubItems(2) = CInt(pItm.SubItems(2)) + 1
            Exit Sub
        End If
    Next d
    
    Set pItm = PosItem.ListItems.Add(, , List.SelectedItem.Text)
    pItm.SubItems(1) = Format(List.SelectedItem.Tag, "#0.00")
    pItm.SubItems(2) = 1
    
    pItm.Selected = True
    pItm.EnsureVisible
    
    PosItem.SetFocus
    
    lcd = "QUANTITY ?"
    info = "Please select quantity..."
    Call asr2_ButtonClick(2, "qty")
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub List_ItemClick(ByVal Item As MSComctlLib.ListItem)
    info = Crnc & " " & Format(Item.Tag, "#0.00")
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Private Sub Disable1(Disableall As Boolean)
    If Disableall = True Then
        List.Enabled = False
        Terminal.Enabled = False
        asr1.ButtonEnabled(2) = False
        asr2.ButtonEnabled(1) = False
        asr3.Enabled = False
    Else
        List.Enabled = True
        Terminal.Enabled = True
        asr1.ButtonEnabled(2) = True
        asr2.ButtonEnabled(1) = True
        asr3.Enabled = True
    End If
End Sub


Private Sub ActionProc()
    Dim Ttl As Double
    Dim pItm As ListItem
    
    If CQty = True Then
        If SelNum = 0 Then PosItem.ListItems.Remove PosItem.SelectedItem.Index
        If PosItem.SelectedItem.Text = "" Then Exit Sub
        
        PosItem.SelectedItem.SubItems(2) = SelNum
        lcd = Crnc & " " & Format(SelNum * CDbl(PosItem.SelectedItem.SubItems(1)), "#0.00")
        Disable1 False
        CQty = False
        Exit Sub
    End If
    
    If CTotal = True Then
        If PosItem.ListItems.Count = 0 Then Exit Sub
        For g = 1 To PosItem.ListItems.Count
            Set pItm = PosItem.ListItems(g)
            Ttl = Ttl + (CDbl(pItm.SubItems(1)) * CDbl(pItm.SubItems(2)))
        Next g
        Total = Ttl
        lcd = Crnc & " " & Format(Ttl, "#0.00")
        CTotal = False
        If Terminal.ListCount > 0 Then AsxD.ButtonEnabled(1) = True
        AsxE.ButtonEnabled(1) = True
    End If
End Sub

Private Sub PosLog(Sales As Double)
    Dim TmpSale  As Double
    pPath = App.Path & "\rekod\" & Year(Date) & "\" & Month(Date) & "\pos.d"
    
    TmpSale = Text2Num(INIambil(pPath, Tarikh, "jualan"))
    TmpSale = TmpSale + Sales
    INIsimpan pPath, Tarikh, "jualan", TmpSale
End Sub

