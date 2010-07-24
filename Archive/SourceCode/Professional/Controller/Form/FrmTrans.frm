VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTrans 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3600
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
   Icon            =   "FrmTrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView List1 
      Height          =   1350
      Left            =   60
      TabIndex        =   4
      Top             =   765
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   2381
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station"
         Object.Width           =   4233
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   3240
      ScaleHeight     =   2205
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   0
      Width           =   360
      Begin CafeBonzer.XpButton MainBtnOk 
         Height          =   345
         Left            =   0
         TabIndex        =   5
         Top             =   1830
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   609
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmTrans.frx":000C
         PICN            =   "FrmTrans.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton MainBtnKo 
         Height          =   345
         Left            =   0
         TabIndex        =   6
         Top             =   1500
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   609
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "FrmTrans.frx":05C2
         PICN            =   "FrmTrans.frx":05DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   630
   End
   Begin VB.Label TermName 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ais Krim Soda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   915
      TabIndex        =   1
      Top             =   90
      Width           =   2250
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "FrmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim tLv As ListView
    Dim tLtm As ListItem
    Set tLv = FrmMain.Lv1
    
    List1.SmallIcons = FrmMain.imglist
    TermName = tLv.SelectedItem.Text
    
    For g = 1 To tLv.ListItems.Count
        If tLv.ListItems(g).Text <> TermName Then
            If tLv.ListItems(g).SubItems(1) = VS(4) Then
                Set tLtm = List1.ListItems.Add(, , tLv.ListItems(g).Text, , "aktif1")
                tLtm.Tag = tLv.ListItems(g).Index
            End If
        End If
    Next g
End Sub

Private Sub MainBtnKo_Click()
    Unload Me
End Sub

Private Sub MainBtnOk_Click()
    If List1.ListItems.Count = 0 Then Exit Sub
    
    If List1.SelectedItem.Text = "" Then
        FrmTrans.Hide
        Unload FrmTrans
    Else
        'hentikan timer sementara
        FrmHost.Timer1.Enabled = False
        FrmHost.Pinger.Enabled = False
        
        Call SelAgn.AgnTransfer(List1.SelectedItem.Tag)
        
        FrmHost.Timer1.Enabled = True
        FrmHost.Pinger.Enabled = True
        Unload FrmTrans
    End If
End Sub
