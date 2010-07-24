VERSION 5.00
Begin VB.Form FrmDebug 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Debug Form"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3825
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3660
   End
End
Attribute VB_Name = "FrmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    List1.Clear
End Sub
