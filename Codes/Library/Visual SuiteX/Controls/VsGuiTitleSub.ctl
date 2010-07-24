VERSION 5.00
Begin VB.UserControl VsGuiTitleSub 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   LockControls    =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   5145
   Begin VisualSuiteX.VsGuiLine GtsLine 
      Height          =   45
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   79
      horizon         =   -1  'True
   End
   Begin VB.Image GtsImg 
      Height          =   240
      Left            =   45
      Picture         =   "VsGuiTitleSub.ctx":0000
      Top             =   45
      Width           =   240
   End
   Begin VB.Label GtsLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VsGuiTitleSub"
      Height          =   195
      Left            =   405
      TabIndex        =   1
      Top             =   75
      Width           =   1005
   End
End
Attribute VB_Name = "VsGuiTitleSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
