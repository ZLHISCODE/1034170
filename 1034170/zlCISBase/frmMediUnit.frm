VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中标单位"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmMediUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUnit.frx":1CFA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUnit.frx":2294
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4921
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf中标单位选择 
      Height          =   2565
      Left            =   5400
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstCustom 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "按药品设置中标单位(&1)"
      TabPicture(0)   =   "frmMediUnit.frx":282E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgNote"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSpec"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblnote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMedi"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vsf停用(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsfUnit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk显示停用(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdStartAll(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMedi"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSave(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdRestore(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdStopAll(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdClose(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMedi"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "按中标单位设置药品(&2)"
      TabPicture(1)   =   "frmMediUnit.frx":284A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdStartAll(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSave(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRestore(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdStopAll(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdClose(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd中标单位"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt中标单位"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chk显示停用(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "vsf中标单位"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "vsf停用(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "imgNotes"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblnotes"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl中标单位"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdStartAll 
         Caption         =   "停用全清(&B)"
         Height          =   350
         Index           =   1
         Left            =   -73560
         TabIndex        =   23
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Index           =   1
         Left            =   -70080
         TabIndex        =   22
         Top             =   7740
         Width           =   1095
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "全部恢复(&R)"
         Height          =   350
         Index           =   1
         Left            =   -74880
         TabIndex        =   21
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdStopAll 
         Caption         =   "停用全选(&A)"
         Height          =   350
         Index           =   1
         Left            =   -72240
         TabIndex        =   20
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭(&C)"
         Height          =   350
         Index           =   1
         Left            =   -68880
         TabIndex        =   19
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmd中标单位 
         Caption         =   "…"
         Height          =   285
         Left            =   -68160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1185
         Width           =   285
      End
      Begin VB.TextBox txt中标单位 
         Height          =   300
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1170
         Width           =   5580
      End
      Begin VB.CheckBox chk显示停用 
         Caption         =   "显示停用"
         Height          =   255
         Index           =   1
         Left            =   -68880
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtMedi 
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1170
         Width           =   5580
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭(&C)"
         Height          =   350
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmdStopAll 
         Caption         =   "停用全选(&A)"
         Height          =   350
         Index           =   0
         Left            =   2760
         TabIndex        =   8
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "全部恢复(&R)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Index           =   0
         Left            =   4920
         TabIndex        =   6
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmdMedi 
         Caption         =   "…"
         Height          =   285
         Left            =   6840
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1185
         Width           =   285
      End
      Begin VB.CommandButton cmdStartAll 
         Caption         =   "停用全清(&B)"
         Height          =   350
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CheckBox chk显示停用 
         Caption         =   "显示停用"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfUnit 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   7095
         _cx             =   12515
         _cy             =   5741
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2866
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf停用 
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   5280
         Visible         =   0   'False
         Width           =   7095
         _cx             =   12515
         _cy             =   4048
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483644
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2974
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf中标单位 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1920
         Width           =   7095
         _cx             =   12515
         _cy             =   5741
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2A0E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf停用 
         Height          =   2295
         Index           =   1
         Left            =   -74880
         TabIndex        =   25
         Top             =   5280
         Visible         =   0   'False
         Width           =   7095
         _cx             =   12515
         _cy             =   4048
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483644
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2B20
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Image imgNotes 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMediUnit.frx":2BBA
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblnotes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    请选择中标单位后，指定其中的药品。招标药品入库时，其供应商必须属于中标单位"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74160
         TabIndex        =   27
         Top             =   540
         Width           =   5685
      End
      Begin VB.Label lbl中标单位 
         AutoSize        =   -1  'True
         Caption         =   "中标单位(&Z)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label lblMedi 
         AutoSize        =   -1  'True
         Caption         =   "药品规格(&M)"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    请选择药品后，指定其中标单位。招标药品入库时，其供应商必须属于中标单位"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   540
         Width           =   5685
      End
      Begin VB.Label lblSpec 
         AutoSize        =   -1  'True
         Caption         =   "规格：      厂牌：       单位：瓶"
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   1560
         Width           =   2970
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   120
         Picture         =   "frmMediUnit.frx":3484
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMediUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lblTag As String
Public frmTag As String
Public strPrivs As String
Private mstrTemp As String
Private mobjItem As ListItem
Private mrsTemp As New ADODB.Recordset
Private mlngId As Long      '记录选中id
Private mstr原值 As String
Private mstr原值中标单位 As String
Private mblnSave As Boolean     '记录是否保存了界面中修改的值
Private mblnSave中标单位 As Boolean
'记录状态类型
Private Enum mStates
    原始 = 0
    新增 = 1
    修改 = 2
    删除 = 3
End Enum
Private Const mcstDelColor = &HDBDBDB
Private Sub vsf_ResetSerial(Optional ByVal bln As Boolean = False)
    Dim i As Integer
    
    If bln Then
        With vsf中标单位
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
        End With
    Else
        With vsfUnit
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
        End With
    End If
End Sub

Private Sub chk显示停用_Click(Index As Integer)
    Call Resize(Index)
End Sub

Private Sub cmdClose_Click(Index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    Dim str原值中标单位 As String
    
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("停用") Then
                    strTemp = strTemp & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("操作") Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With

    With vsf中标单位
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("停用") Then
                    str原值中标单位 = str原值中标单位 & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("操作") Then
                    str原值中标单位 = str原值中标单位 & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With
    
    If (strTemp <> mstr原值 And mblnSave = False) Or (str原值中标单位 <> mstr原值中标单位 And mblnSave中标单位 = False) Then
        If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mblnSave = False
            mblnSave中标单位 = False
            Unload Me
        End If
    Else
        mblnSave = False
        mblnSave中标单位 = False
        Unload Me
    End If

End Sub

Private Sub Delete(Optional ByVal bln As Boolean = False)
    Dim i As Integer
    
    If bln Then
        With vsf中标单位
            If .Rows = 1 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("药品ID"))) = 0 Then Exit Sub
            
            Select Case Val(.TextMatrix(.Row, .ColIndex("状态")))
                Case mStates.新增
                    If .Rows - 1 = 1 Then
                        For i = 1 To .Cols - 1
                            .TextMatrix(1, i) = ""
                        Next
                    Else
                        .RemoveItem .Row
                        Call vsf_ResetSerial(True)
                    End If
            End Select
        End With
    Else
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("单位ID"))) = 0 Then Exit Sub
            
            Select Case Val(.TextMatrix(.Row, .ColIndex("状态")))
                Case mStates.新增
                    If .Rows - 1 = 1 Then
                        For i = 1 To .Cols - 1
                            .TextMatrix(1, i) = ""
                        Next
                    Else
                        .RemoveItem .Row
                        vsf_ResetSerial
                    End If
            End Select
        End With
    End If
End Sub

Private Sub cmdStartAll_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 0 Then
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("停用")) = 0
            Next
        End With
    ElseIf Index = 1 Then
        With vsf中标单位
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("停用")) = 0
            Next
        End With
    End If
End Sub

Private Sub cmdStopAll_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 0 Then
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("停用")) = 1
            Next
        End With
    ElseIf Index = 1 Then
        With vsf中标单位
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("停用")) = 1
            Next
        End With
    End If
End Sub
Private Sub cmdMedi_Click()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "尚未建立该类具体规格的药品！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "厂牌：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left + 120
        .Top = Me.txtMedi.Top + Me.txtMedi.Height + 120
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click(Index As Integer)
    Call ShowData(Index = 1)
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Dim lngUnitId As Long
    Dim lngMediId As Long
    Dim str中标序号 As String
    Dim strDelDate As String
    Dim i As Integer
    Dim strTemp As String
    Dim strContent As String
    Dim rsTemp As ADODB.Recordset
    Dim bln保存 As Boolean
    
    On Error GoTo ErrHand
     
    If vsfUnit.Rows > 1 And Val(lblMedi.Tag) > 0 Then
        mblnSave = True
        lngMediId = Val(lblMedi.Tag)
        
        With vsfUnit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("单位ID"))) > 0 And .TextMatrix(i, .ColIndex("单位")) <> "" Then
                    strTemp = strTemp & "|" & .TextMatrix(i, .ColIndex("单位"))
                    lngUnitId = Val(.TextMatrix(i, .ColIndex("单位ID")))
                    str中标序号 = .TextMatrix(i, .ColIndex("中标序号"))
                    strDelDate = .TextMatrix(i, .ColIndex("建档时间"))
                    
                    If .TextMatrix(i, .ColIndex("停用")) Like "*1" Then .TextMatrix(i, .ColIndex("操作")) = mStates.删除
                    
                    gstrSql = ""
                    Select Case Val(.TextMatrix(i, .ColIndex("操作")))
                        Case mStates.新增
                            gstrSql = "ZL_药品中标单位_INSERT(" & lngMediId & "," & lngUnitId & ", '" & str中标序号 & "')"
                        Case mStates.修改
                            gstrSql = "Zl_药品中标单位_Update(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS') , '" & str中标序号 & "')"
                        Case mStates.删除
                            gstrSql = "ZL_药品中标单位_DELETE(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                    End Select
                    
                    If gstrSql <> "" Then Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption): bln保存 = True
                
                    '同步物流平台药品信息
                    If Not gobjLogisticPlatform Is Nothing Then
                        If Val(.TextMatrix(i, .ColIndex("操作"))) = mStates.删除 Then
                            gobjLogisticPlatform.ClearDrugInfo lngMediId, lngUnitId
                        End If
                        gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngMediId
                    End If
                
                End If
            Next
        End With
    End If
    
    If vsf中标单位.Rows = 1 Or Val(lbl中标单位.Tag) = 0 Then
        If Index = 0 Then
            txtMedi.SetFocus
        Else
            txt中标单位.SetFocus
        End If
        Call ShowData
        Call ShowData(True)
        If bln保存 Then MsgBox "保存成功！", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnSave中标单位 = True
    lngUnitId = Val(lbl中标单位.Tag)
    
    With vsf中标单位
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("药品ID"))) > 0 And .TextMatrix(i, .ColIndex("药品规格")) <> "" Then
                lngMediId = Val(.TextMatrix(i, .ColIndex("药品ID")))
                str中标序号 = .TextMatrix(i, .ColIndex("中标序号"))
                strDelDate = .TextMatrix(i, .ColIndex("建档时间"))
                
                If .TextMatrix(i, .ColIndex("停用")) Like "*1" Then .TextMatrix(i, .ColIndex("操作")) = mStates.删除
                
                If Val(.TextMatrix(i, .ColIndex("操作"))) = mStates.新增 Then
                    gstrSql = "Select 1" & vbNewLine & _
                                    "From 药品中标单位 T" & vbNewLine & _
                                    "Where t.单位id =[1] And t.药品id =[2]" & vbNewLine & _
                                    "And Sysdate Between t.建档时间 And Nvl(t.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngUnitId, lngMediId)
                    If rsTemp.RecordCount > 0 Then .TextMatrix(i, .ColIndex("操作")) = mStates.原始
                End If
                
                gstrSql = ""
                Select Case Val(.TextMatrix(i, .ColIndex("操作")))
                    Case mStates.新增
                        gstrSql = "ZL_药品中标单位_INSERT(" & lngMediId & "," & lngUnitId & ", '" & str中标序号 & "')"
                    Case mStates.修改
                        gstrSql = "Zl_药品中标单位_Update(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS') , '" & str中标序号 & "')"
                    Case mStates.删除
                        gstrSql = "ZL_药品中标单位_DELETE(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                End Select
                
                If gstrSql <> "" Then Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption): bln保存 = True
            
                '同步物流平台药品信息
                If Not gobjLogisticPlatform Is Nothing Then
                    If Val(.TextMatrix(i, .ColIndex("操作"))) = mStates.删除 Then
                        gobjLogisticPlatform.ClearDrugInfo lngMediId, lngUnitId
                    End If
                    gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngMediId
                End If
            
            End If
        Next
    End With
    
    If Index = 0 Then
        txtMedi.SetFocus
    Else
        txt中标单位.SetFocus
    End If

    Call ShowData
    Call ShowData(True)
    
    If bln保存 Then MsgBox "保存成功！", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd中标单位_Click()
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(txt中标单位.hWnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top

    gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) Order By 编码 "
    
    Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
    True, dblLeft, dblTop, txt中标单位.Height, blnCancel, False, True)
    
    If rsRecord Is Nothing Then
        Exit Sub
    Else
        Me.lbl中标单位.Tag = rsRecord!ID
        Me.txt中标单位.Tag = "[" & rsRecord!编码 & "]" & rsRecord!名称
        Me.txt中标单位.Text = Me.txt中标单位.Tag
        Call ShowData(True)
    End If

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If Msf中标单位选择.Visible Then
            Msf中标单位选择.Visible = False
            Exit Sub
        End If
        If lvwItems.Visible Then
            lvwItems.Visible = False: txtMedi.SetFocus: Exit Sub
        End If
        Call cmdClose_Click(0)
    Case Else
    End Select
End Sub

Private Sub sstCustom_Click(PreviousTab As Integer)
    Call Resize(IIf(PreviousTab = 1, 0, 1))
End Sub

Private Sub txt中标单位_GotFocus()
    Me.txt中标单位.SelStart = 0: Me.txt中标单位.SelLength = 100
End Sub

Private Sub txt中标单位_KeyPress(KeyAscii As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    If InStr("~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    vRect = zlControl.GetControlRect(txt中标单位.hWnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top

    gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) and (编码 like [1] or 简码 like[1] or 名称 like [1]) Order By 编码 "
    
    Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
    True, dblLeft, dblTop, txt中标单位.Height, blnCancel, False, True, UCase(txt中标单位.Text) & "%")
    
    If blnCancel = True Then
        Me.txt中标单位.Text = Me.txt中标单位.Tag
        Exit Sub
    End If
    
    If rsRecord Is Nothing Then
        MsgBox "无该中标单位，请重新输入！", vbInformation, gstrSysName
        Me.txt中标单位.Text = Me.txt中标单位.Tag
        Exit Sub
    Else
        Me.lbl中标单位.Tag = rsRecord!ID
        Me.txt中标单位.Tag = "[" & rsRecord!编码 & "]" & rsRecord!名称
        Me.txt中标单位.Text = Me.txt中标单位.Tag
        Call ShowData(True)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Dim i As Long
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] and I.ID=[2] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                "   单位：" & IIf(IsNull(!单位), "", !单位)
            Call ShowData
        End If
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '中草药不允许设置中标单位
    On Error Resume Next
    
    Me.Tag = frmTag
    Me.lblMedi.Tag = lblTag
    frmMediUnit.Height = 6560
    sstCustom.Height = 5895
    cmdRestore(0).Top = 5340
    cmdStartAll(0).Top = 5340
    cmdStopAll(0).Top = 5340
    cmdSave(0).Top = 5340
    cmdClose(0).Top = 5340
    cmdRestore(1).Top = 5340
    cmdStartAll(1).Top = 5340
    cmdStopAll(1).Top = 5340
    cmdSave(1).Top = 5340
    cmdClose(1).Top = 5340
    
    If InStr(1, strPrivs, "中标单位") = 0 Then
        MsgBox "你不具备管理中标单位的权限！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "编码", "编码", 1000
        .Add , "规格", "规格", 1200
        .Add , "产地", IIf(Me.Tag = "7", "产地", "厂牌"), 1200
        .Add , "单位", "单位", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    With vsfUnit
        .ColComboList(.ColIndex("单位")) = "|..."
        .Editable = flexEDKbdMouse
    End With
    
    With vsf中标单位
        .ColComboList(.ColIndex("药品规格")) = "|..."
        .Editable = flexEDKbdMouse
    End With
    
    Call ShowData
    Call ShowData(True)
End Sub

Private Sub txt中标单位_LostFocus()
    Me.txt中标单位.Text = Me.txt中标单位.Tag
End Sub

Private Sub vsfUnit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    With vsfUnit
        If Col = .ColIndex("单位") Then
            gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) Order By 编码 "
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)
            
            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .TextMatrix(Row, .ColIndex("单位")) = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("单位ID")) = rsRecord!ID
    '                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = vbWhite  'mcstInsertColor
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                Else
                    MsgBox "已经有该中标单位！", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mstrTemp = UCase(Trim(Me.txtMedi.Text))
    If mstrTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, mstrTemp, "[") <> 0 And InStr(1, mstrTemp, "]") <> 0 Then mstrTemp = Mid(mstrTemp, 2, InStr(1, mstrTemp, "]") - 2)
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N" & _
            " where I.ID=N.收费细目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, mstrTemp & "%", gstrMatch & mstrTemp & "%")
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定规格的药品，请重新指定！", vbExclamation, gstrSysName
            Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "厂牌：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left + 120
        .Top = Me.txtMedi.Top + Me.txtMedi.Height + 120
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub ShowData(Optional ByVal bln As Boolean = False)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim rsTemps As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If bln Then
        '显示已停用的中标单位
        vsf停用(1).TextMatrix(1, 0) = "1"
        
        gstrSql = "Select a.Id, '[' || a.编码 || ']' || a.名称 || '(' || a.规格 || ')'  as 药品规格, b.撤档时间, b.中标序号" & vbNewLine & _
                        "From 收费项目目录 A, 药品中标单位 B, 供应商 C" & vbNewLine & _
                        "Where a.Id = b.药品id And Instr(c.类型, 1, 1) = 1 And b.单位id = c.Id And b.单位id =[1] And" & vbNewLine & _
                        "     Not (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & vbNewLine & _
                        "Order By b.建档时间"
                    
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl中标单位.Tag))
        
        With rsTemps
            vsf停用(1).Rows = 1
            vsf停用(1).Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsf停用(1).TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsf停用(1).TextMatrix(.AbsolutePosition, vsf停用(1).ColIndex("药品规格")) = !药品规格
                vsf停用(1).TextMatrix(.AbsolutePosition, vsf停用(1).ColIndex("中标序号")) = IIf(IsNull(!中标序号), "", !中标序号)
                vsf停用(1).TextMatrix(.AbsolutePosition, vsf停用(1).ColIndex("撤档时间")) = Format(!撤档时间, "YYYY-MM-DD HH:MM:SS")
                .MoveNext
            Loop
        End With
    
        '显示已初始化的中标单位
        vsf中标单位.TextMatrix(1, 0) = "1"
        
        gstrSql = "Select a.Id, '[' || a.编码 || ']' || a.名称 || '(' || a.规格 || ')'  as 药品规格, b.建档时间, b.中标序号" & vbNewLine & _
                        "From 收费项目目录 A, 药品中标单位 B, 供应商 C" & vbNewLine & _
                        "Where a.Id = b.药品id And Instr(c.类型, 1, 1) = 1 And b.单位id = c.Id And b.单位id =[1] And" & vbNewLine & _
                        "      (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                        "Order By b.建档时间"

        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl中标单位.Tag))
        
        With rsTemps
            vsf中标单位.Rows = 1
            vsf中标单位.Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsf中标单位.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("药品规格")) = !药品规格
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("中标序号")) = IIf(IsNull(!中标序号), "", !中标序号)
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("建档时间")) = Format(!建档时间, "YYYY-MM-DD HH:MM:SS")
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("状态")) = 0
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("操作")) = 0
                vsf中标单位.TextMatrix(.AbsolutePosition, vsf中标单位.ColIndex("药品ID")) = !ID
                .MoveNext
            Loop
        End With
        
        mstr原值中标单位 = ""
        With vsf中标单位
            .Cell(flexcpBackColor, 0, .ColIndex("建档时间"), .Rows - 1, .ColIndex("建档时间")) = mcstDelColor
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    If intCol = .ColIndex("停用") Then
                        mstr原值中标单位 = mstr原值中标单位 & "0|"
                    ElseIf intCol <> .ColIndex("操作") Then
                        mstr原值中标单位 = mstr原值中标单位 & .TextMatrix(intRow, intCol) & "|"
                    End If
                Next
            Next
        End With
    Else
        '显示已停用的中标单位
        vsf停用(0).TextMatrix(1, 0) = "1"
        gstrSql = "Select C.ID,'['||C.编码||']'||C.名称 单位,B.撤档时间,B.中标序号 From 药品规格 A,药品中标单位 B,供应商 C" & _
                " Where A.药品ID=B.药品ID And instr(C.类型,1,1)=1 And B.单位ID=C.ID And A.药品ID=[1] " & _
                " And Not (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
                " Order by B.建档时间"
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lblMedi.Tag))
        
        With rsTemps
            vsf停用(0).Rows = 1
            vsf停用(0).Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsf停用(0).TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsf停用(0).TextMatrix(.AbsolutePosition, vsf停用(0).ColIndex("中标单位")) = !单位
                vsf停用(0).TextMatrix(.AbsolutePosition, vsf停用(0).ColIndex("中标序号")) = IIf(IsNull(!中标序号), "", !中标序号)
                vsf停用(0).TextMatrix(.AbsolutePosition, vsf停用(0).ColIndex("撤档时间")) = Format(!撤档时间, "YYYY-MM-DD HH:MM:SS")
                .MoveNext
            Loop
        End With
        
        '显示已初始化的中标单位
        vsfUnit.TextMatrix(1, 0) = "1"
        
        gstrSql = "Select C.ID,'['||C.编码||']'||C.名称 单位,B.建档时间,B.中标序号 From 药品规格 A,药品中标单位 B,供应商 C" & _
                " Where A.药品ID=B.药品ID And instr(C.类型,1,1)=1 And B.单位ID=C.ID And A.药品ID=[1] " & _
                " And (B.撤档时间 is null or B.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                " Order by B.建档时间"
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lblMedi.Tag))
        
        With rsTemps
            vsfUnit.Rows = 1
            vsfUnit.Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsfUnit.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("单位")) = !单位
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("中标序号")) = IIf(IsNull(!中标序号), "", !中标序号)
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("建档时间")) = Format(!建档时间, "YYYY-MM-DD HH:MM:SS")
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("状态")) = 0
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("操作")) = 0
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("单位ID")) = !ID
                .MoveNext
            Loop
        End With
        
        mstr原值 = ""
        With vsfUnit
            .Cell(flexcpBackColor, 0, .ColIndex("建档时间"), .Rows - 1, .ColIndex("建档时间")) = mcstDelColor
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    If intCol = .ColIndex("停用") Then
                        mstr原值 = mstr原值 & "0|"
                    ElseIf intCol <> .ColIndex("操作") Then
                        mstr原值 = mstr原值 & .TextMatrix(intRow, intCol) & "|"
                    End If
                Next
            Next
        End With
    End If
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("停用") Then
                    strTemp = strTemp & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("操作") Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With
    
    If strTemp <> mstr原值 Then
        If mblnSave = False Then
            If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                lvwItems.Visible = False
                Exit Sub
            End If
        End If
    End If
    
    With Me.lvwItems
        If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.txtMedi.Text = Me.txtMedi.Tag
            lblSpec.Caption = "规格：" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1) & _
                                "      厂牌：" & lvwItems.SelectedItem.SubItems(3) & _
                                "     单位：" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
            Call ShowData
        End If
        Me.txtMedi.SetFocus
        Call OS.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msf中标单位选择_LostFocus()
    With Msf中标单位选择
        .ZOrder 1
        .Visible = False
    End With
End Sub

Private Sub vsfUnit_EnterCell()
    Dim rs修改中标序号 As ADODB.Recordset
    
    On Error GoTo ErrHand
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = .ColIndex("单位") Then
            .ColComboList(.ColIndex("单位")) = ""

            gstrSql = "Select 1" & vbNewLine & _
                            "From 药品中标单位 T" & vbNewLine & _
                            "Where t.单位id =[1] And t.药品id =[2]" & vbNewLine & _
                            "And Sysdate Between t.建档时间 And Nvl(t.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs修改中标序号 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, .TextMatrix(.Row, .ColIndex("单位ID")), Val(lblMedi.Tag))
            If rs修改中标序号.RecordCount > 0 Then .Editable = flexEDNone
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfUnit_GotFocus()
    lvwItems.Visible = False
End Sub

Private Sub vsfUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfUnit
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("中标序号") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("中标序号") Then
                .Row = .Row + 1
                .Col = .ColIndex("单位")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("单位")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete
        End If
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsfUnit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = vsfUnit.ColIndex("单位") Then
        vsfUnit.ColComboList(Col) = "|..."
    End If
End Sub

Private Sub vsfUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsfUnit
        If .Col = .ColIndex("单位") Then
            .ColComboList(.ColIndex("单位")) = "|..."
        Else
            .ColComboList(.ColIndex("单位")) = ""
        End If
    End With
End Sub

Private Sub vsfUnit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsfUnit
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsfUnit_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfUnit
        If Col = .ColIndex("中标序号") Then
            .EditMaxLength = 50
        Else
            .EditMaxLength = 50
        End If
    End With
End Sub

Private Sub vsfUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rs修改中标序号 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    With vsfUnit
        If Col = .ColIndex("单位") And .EditText = "" Then Exit Sub
        If .Rows < 2 Then Exit Sub
        If Col = .ColIndex("单位") And InStr(1, .EditText, "[") = 0 Then
            gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) and (编码 like [1] or 简码 like[1] or 名称 like [1]) Order By 编码 "
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%")

            If blnCancel = True Then
                Cancel = True
                Exit Sub
            End If
            
            If rsRecord Is Nothing Then
                MsgBox "无该中标单位！", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .EditText = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(Row, .ColIndex("单位")) = .EditText
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("单位ID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                Else
                    MsgBox "已经有该中标单位！", vbInformation, gstrSysName
                    Cancel = True
                End If
            End If
 
         ElseIf Col = .ColIndex("中标序号") And .TextMatrix(.Row, .ColIndex("中标序号")) <> .EditText Then
            gstrSql = "Select 1" & vbNewLine & _
                            "From 药品中标单位 T" & vbNewLine & _
                            "Where t.单位id =[1] And t.药品id =[2]" & vbNewLine & _
                            "And Sysdate Between t.建档时间 And Nvl(t.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs修改中标序号 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, .TextMatrix(Row, .ColIndex("单位ID")), Val(lblMedi.Tag))
            If rs修改中标序号.RecordCount > 0 Then .TextMatrix(Row, .ColIndex("操作")) = mStates.修改
            
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDub(Optional ByVal bln As Boolean = False) As Boolean
    '检查是否存在该中标单位或是否存在药品
    Dim i As Integer
    
    If bln Then
        With vsf中标单位
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("药品id")) <> "" Then
                    If .TextMatrix(i, .ColIndex("药品id")) = mlngId Then
                        CheckDub = True
                        Exit Function
                    End If
                Else
                    CheckDub = False
                End If
            Next
        End With
        CheckDub = False
    Else
        With vsfUnit
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("单位id")) <> "" Then
                    If .TextMatrix(i, .ColIndex("单位id")) = mlngId Then
                        CheckDub = True
                        Exit Function
                    End If
                Else
                    CheckDub = False
                End If
            Next
        End With
        CheckDub = False
    End If
End Function

Private Sub vsf中标单位_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsf中标单位.hWnd) '获取位置
    dblLeft = vRect.Left + vsf中标单位.CellLeft
    dblTop = vRect.Top + vsf中标单位.CellTop + vsf中标单位.CellHeight + 3300
    With vsf中标单位
        If Col = .ColIndex("药品规格") Then
            gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地" & _
                    " from 收费项目目录 I" & _
                    " where I.类别 in('5','6')" & _
                    "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
                    
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "药品规格", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)

            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub(True) = False Then
                    .TextMatrix(Row, 0) = Row
                    .TextMatrix(Row, .ColIndex("药品规格")) = "[" & rsRecord!编码 & "]" & rsRecord!名称 & "(" & rsRecord!规格 & ")"
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("药品ID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                Else
                    MsgBox "已经有该药品！", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub vsf中标单位_EnterCell()
    Dim rs修改中标序号 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With vsf中标单位
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = .ColIndex("药品规格") Then
            .ColComboList(.ColIndex("药品规格")) = ""

            gstrSql = "Select 1" & vbNewLine & _
                            "From 药品中标单位 T" & vbNewLine & _
                            "Where t.单位id =[1] And t.药品id =[2]" & vbNewLine & _
                            "And Sysdate Between t.建档时间 And Nvl(t.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs修改中标序号 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl中标单位.Tag), .TextMatrix(.Row, .ColIndex("药品ID")))
            If rs修改中标序号.RecordCount > 0 Then .Editable = flexEDNone
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf中标单位_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsf中标单位
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("中标序号") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("中标序号") Then
                .Row = .Row + 1
                .Col = .ColIndex("药品规格")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("药品规格")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete(True)
        End If
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsf中标单位_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = vsf中标单位.ColIndex("药品规格") Then
        vsf中标单位.ColComboList(Col) = "|..."
    End If
End Sub

Private Sub vsf中标单位_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsf中标单位
        If .Col = .ColIndex("药品规格") Then
            .ColComboList(.ColIndex("药品规格")) = "|..."
        Else
            .ColComboList(.ColIndex("药品规格")) = ""
        End If
    End With
End Sub

Private Sub vsf中标单位_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsf中标单位
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsf中标单位_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsf中标单位
        If Col = .ColIndex("中标序号") Then
            .EditMaxLength = 50
        Else
            .EditMaxLength = 50
        End If
    End With
End Sub

Private Sub vsf中标单位_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rs修改中标序号 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsf中标单位.hWnd) '获取位置
    dblLeft = vRect.Left + vsf中标单位.CellLeft
    dblTop = vRect.Top + vsf中标单位.CellTop + vsf中标单位.CellHeight + 3300
    With vsf中标单位
        If Col = .ColIndex("药品规格") And .EditText = "" Then Exit Sub
        If .Rows < 2 Then Exit Sub
        If Col = .ColIndex("药品规格") And InStr(1, .EditText, "[") = 0 Then
        gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地" & _
                " from 收费项目目录 I,收费项目别名 N" & _
                " where I.ID=N.收费细目ID and I.类别 in('5','6') " & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "药品规格", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%", gstrMatch & UCase(.EditText) & "%")
            
            If blnCancel = True Then
                Cancel = True
                Exit Sub
            End If
  
            If rsRecord Is Nothing Then
                MsgBox "没有找到该药品！", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub(True) = False Then
                    .TextMatrix(Row, 0) = Row
                    .EditText = "[" & rsRecord!编码 & "]" & rsRecord!名称 & "(" & rsRecord!规格 & ")"
                    .TextMatrix(Row, .ColIndex("药品规格")) = .EditText
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("药品ID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                Else
                    MsgBox "已经有该药品！", vbInformation, gstrSysName
                    Cancel = True
                End If
            End If
            
         ElseIf Col = .ColIndex("中标序号") And .TextMatrix(.Row, .ColIndex("中标序号")) <> .EditText Then
            gstrSql = "Select 1" & vbNewLine & _
                            "From 药品中标单位 T" & vbNewLine & _
                            "Where t.单位id =[1] And t.药品id =[2]" & vbNewLine & _
                            "And Sysdate Between t.建档时间 And Nvl(t.撤档时间, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs修改中标序号 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl中标单位.Tag), .TextMatrix(Row, .ColIndex("药品ID")))
            If rs修改中标序号.RecordCount > 0 Then .TextMatrix(Row, .ColIndex("操作")) = mStates.修改
            
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Resize(ByVal Index As Integer)
    On Error Resume Next
 
    If chk显示停用(Index).Value = 1 Then
        vsf停用(Index).Visible = True
        frmMediUnit.Height = 9000
        sstCustom.Height = 8295
        cmdRestore(Index).Top = 7740
        cmdStartAll(Index).Top = 7740
        cmdStopAll(Index).Top = 7740
        cmdSave(Index).Top = 7740
        cmdClose(Index).Top = 7740
        vsf停用(Index).Top = 5280
        vsf停用(Index).Left = 120
    Else
        vsf停用(Index).Visible = False
        frmMediUnit.Height = 6560
        sstCustom.Height = 5895
        cmdRestore(Index).Top = 5340
        cmdStartAll(Index).Top = 5340
        cmdStopAll(Index).Top = 5340
        cmdSave(Index).Top = 5340
        cmdClose(Index).Top = 5340
    End If
End Sub


