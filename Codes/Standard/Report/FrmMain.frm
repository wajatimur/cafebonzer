VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "CafeBonzer Report System"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select Report :"
      Height          =   870
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3585
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   450
         Left            =   2475
         TabIndex        =   2
         Top             =   285
         Width           =   915
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmMain.frx":23D2
         Left            =   210
         List            =   "FrmMain.frx":23D9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2070
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Combo1 = "" Then Exit Sub
    If Combo1 = "pc-usage" Then RptPcSales.Show
    Unload Me
End Sub
