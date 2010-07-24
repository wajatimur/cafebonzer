VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAgnMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CafeBonzer Message Center"
   ClientHeight    =   5775
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   9000
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
   Icon            =   "FrmMesej.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FrmMesej.frx":6852
   ScaleHeight     =   5775
   ScaleWidth      =   9000
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList Img16 
      Left            =   8355
      Top             =   855
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":887B
            Key             =   "MSGOUT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":8E15
            Key             =   "MSGIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":93AF
            Key             =   "TERMUSER"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMesej.frx":9749
            Key             =   "TERM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LsvMessage 
      Height          =   4020
      Left            =   2385
      TabIndex        =   7
      Top             =   810
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   7091
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Img16"
      ForeColor       =   4210752
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Message"
         Object.Width           =   11465
      EndProperty
   End
   Begin VB.ComboBox CmbTo 
      Height          =   315
      ItemData        =   "FrmMesej.frx":9AE3
      Left            =   510
      List            =   "FrmMesej.frx":9AED
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   1830
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   5475
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15822
            Picture         =   "FrmMesej.frx":9B00
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LsvTerminal 
      Height          =   4245
      Left            =   30
      TabIndex        =   3
      Top             =   1215
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "Img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Terminal"
         Object.Width           =   3881
      EndProperty
   End
   Begin CafeBonzer.Line3D CmcLine 
      Height          =   45
      Left            =   15
      TabIndex        =   2
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.TextBox MsgSend 
      ForeColor       =   &H00404040&
      Height          =   570
      Left            =   2385
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4890
      Width           =   5940
   End
   Begin CafeBonzer.XpButton BtnMsgtick 
      Height          =   570
      Left            =   8385
      TabIndex        =   1
      ToolTipText     =   "Send ticker message"
      Top             =   4890
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "FrmMesej.frx":A09A
      PICN            =   "FrmMesej.frx":A0B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label CmcLabel 
      AutoSize        =   -1  'True
      Caption         =   "To :"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   870
      Width           =   345
   End
End
Attribute VB_Name = "FrmAgnMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAgnMsg
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private Sub Form_Load()
    Dim LngIdxA As Long, CAgent As ClsAgent
    Dim StrIcon As String
    
    '{ Load all terminal }'
    For LngIdxA = 1 To UniAgents.Count
        Set CAgent = UniAgents.Agents(LngIdxA)
        If CAgent.AgentStatus = VS(1, 1) Then
            StrIcon = "TERMUSER"
        Else
            StrIcon = "TERM"
        End If
        LsvTerminal.ListItems.Add , CAgent.AgentName, CAgent.AgentName, , StrIcon
    Next
    
    '{ Select default }'
    CmbTo.ListIndex = 1
End Sub

Private Sub MsgSend_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim LngIdxA As Long
    
    If KeyCode = vbKeyReturn Then
    '{ Key ENTER pressed }'
        If CmbTo.ListIndex = 0 Then
        '{ Send to all }'
            For LngIdxA = 1 To UniAgents.Count
                UniAgents.Agents(LngIdxA).NetSend "030010" & MsgSend
            Next
        Else
        '{ Send to selected terminal }'
            If LsvTerminal.ListItems.Count > 0 Then
                UniAgents.Agents(LsvTerminal.SelectedItem.Key).NetSend "030010" & MsgSend
            End If
        End If
        StatusBar.Panels(1).Text = MsgSend
        MsgSend = ""
    ElseIf KeyCode = vbKeyEscape Then
    '{ Key ESCAPE pressed }'
        FrmAgnMsg.Hide
        CbMsgRcv = False
    End If
End Sub

Private Sub BtnMsgtick_Click()
        If CmbTo.ListIndex = 0 Then
        '{ Send to all }'
            For LngIdxA = 1 To UniAgents.Count
                UniAgents.Agents(LngIdxA).NetSend "030020" & MsgSend
            Next
        Else
        '{ Send to selected terminal }'
            UniAgents.Agents(LsvTerminal.SelectedItem.Key).NetSend "030020" & MsgSend
        End If
End Sub

Public Sub AddMessage(ObjectAgent As ClsAgent, Message As String)
    Dim StrMessage As String
    
    StrMessage = "[" & ObjectAgent.AgentName & "] " & Message
    LsvMessage.ListItems.Add , , StrMessage, , "MSGIN"
End Sub

Public Sub SetToAll()
    CmbTo.ListIndex = 0
End Sub
