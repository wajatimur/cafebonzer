VERSION 5.00
Begin VB.Form FrmAddWorker 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1725
   ClientLeft      =   8760
   ClientTop       =   3885
   ClientWidth     =   4245
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
   Icon            =   "FrmAddWorker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4245
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   3
      Left            =   1605
      TabIndex        =   3
      ToolTipText     =   "Enter the password for this worker."
      Top             =   1305
      Width           =   2100
   End
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   2
      Left            =   1605
      TabIndex        =   2
      ToolTipText     =   "Please enter the monthly salary."
      Top             =   900
      Width           =   2100
   End
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   1
      Left            =   1605
      TabIndex        =   1
      ToolTipText     =   "Enter the worker nick name."
      Top             =   495
      Width           =   2100
   End
   Begin VB.TextBox Txt1 
      Height          =   315
      Index           =   0
      Left            =   1605
      TabIndex        =   0
      ToolTipText     =   "Enter worker name."
      Top             =   105
      Width           =   2100
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   3885
      ScaleHeight     =   1755
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   0
      Width           =   360
      Begin CafeBonzer.XpButton MainBtn 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   1380
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
         MICON           =   "FrmAddWorker.frx":000C
         PICN            =   "FrmAddWorker.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton MainBtn 
         Height          =   345
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   1050
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
         MICON           =   "FrmAddWorker.frx":05C2
         PICN            =   "FrmAddWorker.frx":05DE
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   300
      Index           =   3
      Left            =   150
      TabIndex        =   8
      Top             =   1335
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary :"
      Height          =   300
      Index           =   2
      Left            =   135
      TabIndex        =   7
      Top             =   945
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName :"
      Height          =   300
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   540
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Worker Name :"
      Height          =   300
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   150
      Width           =   1440
   End
End
Attribute VB_Name = "FrmAddWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MainBtn_Click(Index As Integer)
    Select Case Index
    Case 0
        Dim nItm As ListItem
        If Txt1(0) = "" Then Unload FrmAddWorker
        If Txt1(1) = "" Then Unload FrmAddWorker
        If Txt1(2) = "" Then Unload FrmAddWorker
        If Txt1(3) = "" Then
            Unload FrmAddWorker
        Else
            Set nItm = FrmSet.Lv1.ListItems.Add(, , Txt1(1), , "user")
            nItm.SubItems(1) = Txt1(2)
            
            Set nItm = FrmSet.Lv1.ListItems.Add(, , Txt1(1), , "akses")
            nItm.SubItems(3) = Txt1(3)
            nItm.SubItems(4) = "000"
            
            uSDBe.DataSave "pekerja-list", "nama", Txt1(0), True, False
            uSDBe.DataSave "pekerja-list", "nick", Txt1(1), False, False
            uSDBe.DataSave "pekerja-list", "gaji", Txt1(2), False, False
            uSDBe.DataSave "pekerja-list", "akses", "000", False, False
            uSDBe.DataSave "pekerja-list", "password", Txt1(3), False, True
            Unload FrmAddWorker
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then MoveFrm Me.hwnd
End Sub
