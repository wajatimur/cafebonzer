Attribute VB_Name = "mdlSHAppBar"
Option Explicit

'*********************************************************************************************
'
' Shell application bar declarations module
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.geocities.com/SiliconValley/Foothills/9940
'
' Last Updated: 07/10/1999
'
'*********************************************************************************************

Public Enum SHAppBar_Messages
    ABM_NEW = &H0
    ABM_REMOVE = &H1
    ABM_QUERYPOS = &H2
    ABM_SETPOS = &H3
    ABM_GETSTATE = &H4
    ABM_GETTASKBARPOS = &H5
    ABM_ACTIVATE = &H6
    ABM_GETAUTOHIDEBAR = &H7
    ABM_SETAUTOHIDEBAR = &H8
    ABM_WINDOWPOSCHANGED = &H9
End Enum

Public Enum SHAppBar_Notifications
    ABN_STATECHANGE = &H0
    ABN_POSCHANGED = &H1
    ABN_FULLSCREENAPP = &H2
    ABN_WINDOWARRANGE = &H3
End Enum

Public Enum SHAppBar_States
    ABS_AUTOHIDE = &H1
    ABS_ALWAYSONTOP = &H2
End Enum

Public Enum SHAppBar_Edges
    ABE_LEFT = 0
    ABE_TOP = 1
    ABE_RIGHT = 2
    ABE_BOTTOM = 3
End Enum

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type AppBarData
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As SHAppBar_Edges
    rc As RECT
    lParam As Long
End Type

Declare Function SHAppBarMessage Lib "shell32" (ByVal dwMessage As SHAppBar_Messages, pData As AppBarData) As Long

