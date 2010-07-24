VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMainTool 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10200
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin CafeBonzer.PageHolder MainPhold 
      Height          =   2970
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   5239
      HldrTxt         =   "Toolbox"
      HldrTxtClr      =   16777215
      HldrLne         =   -1  'True
      PageHeight      =   2970
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   0
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   34
         Top             =   360
         Width           =   9660
         Begin CafeBonzer.Line3D SpgInfoLine 
            Height          =   2610
            Left            =   4170
            TabIndex        =   35
            Top             =   -15
            Width           =   45
            _ExtentX        =   79
            _ExtentY        =   4604
            horizon         =   0   'False
         End
         Begin VB.Label SpgInfoLblA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connected Agent :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   4650
            TabIndex        =   47
            Top             =   495
            Width           =   1770
         End
         Begin VB.Label SpgInfoLblB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   6585
            TabIndex        =   46
            Top             =   480
            Width           =   885
         End
         Begin VB.Label SpgInfoLblA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unused Station :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   4830
            TabIndex        =   45
            Top             =   930
            Width           =   1590
         End
         Begin VB.Label SpgInfoLblB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   6585
            TabIndex        =   44
            Top             =   915
            Width           =   885
         End
         Begin VB.Image SpgInfoHdr 
            Height          =   270
            Index           =   0
            Left            =   4365
            Picture         =   "FrmMainTool.frx":0000
            Top             =   60
            Width           =   2400
         End
         Begin VB.Image SpgInfoHdr 
            Height          =   300
            Index           =   1
            Left            =   90
            Picture         =   "FrmMainTool.frx":0BCA
            Top             =   60
            Width           =   3225
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Connected At :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   43
            Top             =   540
            Width           =   1410
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   42
            Top             =   975
            Width           =   1185
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MAC Address :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   41
            Top             =   1425
            Width           =   1380
         End
         Begin VB.Label SpgInfoLblC 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Used :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   40
            Top             =   1860
            Width           =   1395
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   1860
            TabIndex        =   39
            Top             =   525
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   1860
            TabIndex        =   38
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   2
            Left            =   1860
            TabIndex        =   37
            Top             =   1410
            Width           =   2115
         End
         Begin VB.Label SpgInfoLblD 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   3
            Left            =   1860
            TabIndex        =   36
            Top             =   1845
            Width           =   2115
         End
      End
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   2
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   30
         Top             =   360
         Width           =   9660
         Begin VB.TextBox MainNote 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2430
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   60
            Width           =   7185
         End
         Begin CafeBonzer.XpButton MainNoteBtn 
            Height          =   345
            Index           =   0
            Left            =   7290
            TabIndex        =   32
            Top             =   75
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
            MICON           =   "FrmMainTool.frx":1AEB
            PICN            =   "FrmMainTool.frx":1B07
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton MainNoteBtn 
            Height          =   345
            Index           =   1
            Left            =   7290
            TabIndex        =   33
            Top             =   435
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
            MICON           =   "FrmMainTool.frx":20A1
            PICN            =   "FrmMainTool.frx":20BD
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
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   3
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   28
         Top             =   360
         Width           =   9660
         Begin VB.ListBox MainLog 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2370
            Left            =   60
            TabIndex        =   29
            Top             =   75
            Width           =   7740
         End
      End
      Begin VB.PictureBox SubPages 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2565
         Index           =   1
         Left            =   570
         ScaleHeight     =   2565
         ScaleWidth      =   9660
         TabIndex        =   6
         Top             =   360
         Width           =   9660
         Begin VB.TextBox SerTxtPriItm 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1275
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1635
            Width           =   1740
         End
         Begin VB.TextBox SerTxtTotalItm 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1275
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   2145
            Width           =   1740
         End
         Begin VB.VScrollBar SerScroll1 
            Height          =   330
            Left            =   1920
            Max             =   999
            Min             =   1
            TabIndex        =   11
            Top             =   1125
            Value           =   999
            Width           =   165
         End
         Begin VB.TextBox SerTxtBayar 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1365
         End
         Begin VB.TextBox SerTxtBaki 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   660
            Width           =   1365
         End
         Begin VB.TextBox SerTxtQty 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1275
            TabIndex        =   8
            Text            =   "1"
            Top             =   1125
            Width           =   615
         End
         Begin VB.TextBox SerTxtJumlah 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Endless Showroom"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7965
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   1365
         End
         Begin MSComctlLib.ImageCombo SerImgCb2 
            Height          =   330
            Left            =   1275
            TabIndex        =   14
            Top             =   600
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Text            =   "None"
         End
         Begin MSComctlLib.ImageCombo SerImgCb1 
            Height          =   330
            Left            =   1275
            TabIndex        =   15
            Top             =   75
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
            Text            =   "None"
            ImageList       =   "Iml"
         End
         Begin MSComctlLib.ListView SerLv1 
            Height          =   2415
            Left            =   3150
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   60
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   4260
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   12648384
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Price"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Qty."
               Object.Width           =   1058
            EndProperty
         End
         Begin CafeBonzer.XpButton SerBtn 
            Height          =   480
            Index           =   0
            Left            =   8895
            TabIndex        =   17
            ToolTipText     =   "Accept"
            Top             =   1995
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
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
            MICON           =   "FrmMainTool.frx":2657
            PICN            =   "FrmMainTool.frx":2673
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton SerBtnItm 
            Height          =   360
            Index           =   0
            Left            =   2130
            TabIndex        =   18
            Top             =   1095
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            MICON           =   "FrmMainTool.frx":46F5
            PICN            =   "FrmMainTool.frx":4711
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin CafeBonzer.XpButton SerBtnItm 
            Height          =   360
            Index           =   1
            Left            =   2565
            TabIndex        =   19
            Top             =   1095
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   635
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
            MICON           =   "FrmMainTool.frx":4CAB
            PICN            =   "FrmMainTool.frx":4CC7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Price :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   27
            Top             =   1650
            Width           =   555
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items Price :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   26
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   5
            Left            =   6585
            TabIndex        =   25
            Top             =   1275
            Width           =   1125
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balanced :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   4
            Left            =   6585
            TabIndex        =   24
            Top             =   735
            Width           =   1140
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   23
            Top             =   105
            Width           =   930
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   22
            Top             =   615
            Width           =   630
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   21
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label SerLbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   3
            Left            =   6585
            TabIndex        =   20
            Top             =   180
            Width           =   675
         End
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   0
         Left            =   30
         TabIndex        =   5
         ToolTipText     =   "Information"
         Top             =   375
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         MICON           =   "FrmMainTool.frx":5261
         PICN            =   "FrmMainTool.frx":527D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   1
         Left            =   30
         TabIndex        =   4
         ToolTipText     =   "Service & Merchandise"
         Top             =   795
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         MICON           =   "FrmMainTool.frx":5817
         PICN            =   "FrmMainTool.frx":5833
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   2
         Left            =   30
         TabIndex        =   3
         ToolTipText     =   "Note"
         Top             =   1215
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         MICON           =   "FrmMainTool.frx":5DCD
         PICN            =   "FrmMainTool.frx":5DE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.XpButton SubPagesMnu 
         Height          =   405
         Index           =   3
         Left            =   30
         TabIndex        =   2
         ToolTipText     =   "Log"
         Top             =   1635
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         MICON           =   "FrmMainTool.frx":6383
         PICN            =   "FrmMainTool.frx":639F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin CafeBonzer.Line3D MainLine 
         Height          =   2595
         Left            =   465
         TabIndex        =   1
         Top             =   345
         Width           =   45
         _ExtentX        =   79
         _ExtentY        =   4577
         horizon         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmMainTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    MainPhold.Width = Me.Width
End Sub

Private Sub MainPhold_HolderButtonClick(ByVal Collapse As Boolean)
    If Collapse = True Then
        MainPhold.Top = 0
        Me.Height = MainPhold.Height + 100
        Me.Top = FrmMaster.LngWorkSpaceY - Me.Height
    Else
        MainPhold.Top = 0
        Me.Height = MainPhold.Height
        Me.Top = FrmMaster.LngWorkSpaceY - Me.Height
    End If
End Sub
