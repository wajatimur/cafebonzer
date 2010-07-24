VERSION 5.00
Begin VB.Form FrmMenu
   Caption         =   "FrmMenu"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MenuMain
      Caption = "Menu"
      Begin VB.Menu MenuMainSetting
         Caption = "Setting"
      End
      Begin VB.Menu MnuMainSep1
         Caption = "-"
      End
      Begin VB.Menu MnuMainLogout
         Caption = "Logout"
      End
      Begin VB.Menu MnuMainClose
         Caption = "Close"
      End
   End
   Begin VB.Menu MnuAgent
      Caption = "Station"
      Begin VB.Menu MnuAgentBroad
         Caption = "Broadcast"
         Begin VB.Menu MnuAgentBroadMsg
            Caption = "Message"
         End
         Begin VB.Menu MnuAgentBroadTick
            Caption = "Ticker"
         End
      End
      Begin VB.Menu MnuAgentControl
         Caption = "Control"
         Begin VB.Menu MnuAgentControlLock
            Caption = "Lock All"
            Index = 0
         End
         Begin VB.Menu MnuAgentControlLock
            Caption = "Lock Unused"
            Index = 2
         End
         Begin VB.Menu MnuAgentControlLock
            Caption = "Unlock All"
            Index = 1
         End
         Begin VB.Menu menu3sep2
            Caption = "-"
         End
         Begin VB.Menu MnuAgentControlExit
            Caption = "Shutdown All"
            Index = 0
         End
         Begin VB.Menu MnuAgentControlExit
            Caption = "Shutdown Unused"
            Index = 2
         End
         Begin VB.Menu MnuAgentControlExit
            Caption = "Reboot All"
            Index = 1
         End
         Begin VB.Menu MnuAgentControlExit
            Caption = "Reboot Unused"
            Index = 4
         End
      End
      Begin VB.Menu MnuAgentClean
         Caption = "Cleaning"
         Begin VB.Menu menu3clnsub
            Caption = "All"
            Index = 0
         End
         Begin VB.Menu menu3clnsub
            Caption = "-"
            Index = 1
         End
         Begin VB.Menu menu3clnsub
            Caption = "Temp Folder"
            Index = 2
         End
         Begin VB.Menu menu3clnsub
            Caption = "Recycle Bin"
            Index = 3
         End
         Begin VB.Menu menu3clnsub
            Caption = "Internet History"
            Index = 4
         End
         Begin VB.Menu menu3clnsub
            Caption = "Recent Docs"
            Index = 5
         End
      End
      Begin VB.Menu MnuAgentSep1
         Caption = "-"
      End
      Begin VB.Menu MnuAgentManager
         Caption = "Agent Manager"
      End
   End
   Begin VB.Menu MnuView
      Caption = "View"
      Begin VB.Menu MnuViewMonPrint
         Caption = "Printing Monitoring"
         Index = 0
      End
      Begin VB.Menu MnuViewMonRes
         Caption = "Resource Monitoring"
         Index = 1
      End
      Begin VB.Menu MnuViewMonApp
         Caption = "Application Monitoring"
         Index = 2
      End
      Begin VB.Menu MnuViewMonTrf
         Caption = "Traffic Monitoring"
         Index = 3
      End
   End
   Begin VB.Menu MnuTools
      Caption = "Tools"
      Begin VB.Menu MnuToolsSnm
         Caption = "S&&M Manager"
      End
      Begin VB.Menu MnuToolsStat
         Caption = "Statistic System"
      End
      Begin VB.Menu MnuToolsConsole
         Caption = "Console System"
      End
   End
   Begin VB.Menu MnuInfo
      Caption = "Info"
      Begin VB.Menu MnuInfoHelp
         Caption = "Help"
      End
      Begin VB.Menu MnuInfoSep1
         Caption = "-"
      End
      Begin VB.Menu MnuInfoAbout
         Caption = "About.."
      End
   End
   Begin VB.Menu PopMnu1
      Caption = "<popmenu1>"
      Visible = 0 'False
      Begin VB.Menu PopMnu1Flog
         Caption = "Fast Login"
      End
      Begin VB.Menu PopMnu1Flout
         Caption = "Fast Logout"
      End
      Begin VB.Menu PopMnuSep2
         Caption = "-"
      End
      Begin VB.Menu PopMnu1Cancel
         Caption = "Cancel User"
      End
      Begin VB.Menu PopMnu1Trans
         Caption = "Transfer PC"
      End
      Begin VB.Menu PopMnu1Terminal
         Caption = "Terminal"
      End
      Begin VB.Menu PopMnuSep1
         Caption = "-"
      End
      Begin VB.Menu PopMnu1Cln
         Caption = "Cleaning"
         Begin VB.Menu PopMnu1ClnSub
            Caption = "All"
            Index = 0
         End
         Begin VB.Menu PopMnu1ClnSub
            Caption = "-"
            Index = 1
         End
         Begin VB.Menu PopMnu1ClnSub
            Caption = "Temp Folder"
            Index = 2
         End
         Begin VB.Menu PopMnu1ClnSub
            Caption = "Recycle Bin"
            Index = 3
         End
         Begin VB.Menu PopMnu1ClnSub
            Caption = "Internet History"
            Index = 4
         End
         Begin VB.Menu PopMnu1ClnSub
            Caption = "Recent Docs"
            Index = 5
         End
      End
      Begin VB.Menu PopMnu1Ctl
         Caption = "Control"
         Begin VB.Menu PopMnu1CtlSub
            Caption = "Lock Computer"
            Index = 0
         End
         Begin VB.Menu PopMnu1CtlSub
            Caption = "Unlock Computer"
            Index = 1
         End
         Begin VB.Menu PopMnu1CtlSub
            Caption = "Reboot Computer"
            Index = 2
         End
         Begin VB.Menu PopMnu1CtlSub
            Caption = "Shutdown Computer"
            Index = 3
         End
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_GlobalNameSpace = False
Option Explicit
