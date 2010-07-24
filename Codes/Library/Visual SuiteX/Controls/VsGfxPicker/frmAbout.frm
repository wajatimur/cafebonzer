VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Color Picker"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Copyright©: Abdul Gafoor.GK"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "EMail:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "gafoorgk@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   1920
      MouseIcon       =   "frmAbout.frx":34DA
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   1860
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1

Private Const MyEMail As String = "gafoorgk@yahoo.com"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Label2_Click(Index As Integer)
    If (Index = 2) Then
        Call ShellExecute(0&, vbNullString, "mailto:" & MyEMail, vbNullString, "C:\", SW_SHOWNORMAL)
    End If
End Sub
