VERSION 5.00
Begin VB.Form FrmAppSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5385
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
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":000C
   ScaleHeight     =   2985
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblBuild 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Build 1.7.42"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   45
      TabIndex        =   0
      Top             =   2775
      Width           =   930
   End
End
Attribute VB_Name = "FrmAppSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : FrmAppSplash
'    Project    : CafeBonzer
'
'    Description: Application Splash
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Private Sub Form_Load()
    LblBuild = "Build " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
