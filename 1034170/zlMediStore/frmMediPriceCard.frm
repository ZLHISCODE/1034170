VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMediPriceCard 
   Caption         =   "药品调价单"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
   Icon            =   "frmMediPriceCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   48
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   1080
      TabIndex        =   44
      Top             =   9505
      Width           =   1965
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3135
      Left            =   3600
      ScaleHeight     =   3075
      ScaleWidth      =   4755
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2400
         Picture         =   "frmMediPriceCard.frx":6852
         TabIndex        =   41
         Top             =   2640
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3480
         Picture         =   "frmMediPriceCard.frx":699C
         TabIndex        =   40
         Top             =   2640
         Width           =   1100
      End
      Begin VB.Frame fra辅助选项 
         Caption         =   "辅助选项（成本价调价相关）"
         Height          =   2535
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4695
         Begin VB.CheckBox chk加成率 
            Caption         =   "指定加成率"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk供应商 
            Caption         =   "指定供应商"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk应付记录 
            Caption         =   "产生成本价调价带来的应付款修正记录"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt加成率 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   32
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   31
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   30
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            GridColor       =   32768
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblComment加成率 
            Caption         =   "（指定加成率，则统一默认按该加成率计算成本价；不指定，则默认显示实际加成率）"
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment供应商 
            AutoSize        =   -1  'True
            Caption         =   "（指定供应商，则只调整该供应商的库存药品成本价）"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   37
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   10575
      TabIndex        =   23
      Top             =   8640
      Width           =   10575
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   26
         Top             =   120
         Width           =   5565
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "调价说明"
         Height          =   180
         Left            =   3360
         TabIndex        =   27
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "调价人"
         Height          =   180
         Left            =   360
         TabIndex        =   25
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清空(&D)"
      Height          =   350
      Left            =   6960
      Picture         =   "frmMediPriceCard.frx":6AE6
      TabIndex        =   15
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   12600
      Picture         =   "frmMediPriceCard.frx":6C30
      TabIndex        =   14
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   13920
      Picture         =   "frmMediPriceCard.frx":6D7A
      TabIndex        =   13
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印库存变动表(&P)…"
      Height          =   350
      Left            =   10200
      Picture         =   "frmMediPriceCard.frx":6EC4
      TabIndex        =   12
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "批量选择项目(&I)"
      Height          =   350
      Left            =   8400
      Picture         =   "frmMediPriceCard.frx":700E
      TabIndex        =   11
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   16335
      Begin VB.CheckBox chkAutoPay 
         Caption         =   "自动计算应付款变动记录"
         Height          =   210
         Left            =   8160
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkCostBatch 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   2160
         TabIndex        =   42
         Top             =   480
         Width           =   2370
      End
      Begin VB.CheckBox chkAotuCost 
         Caption         =   "调售价时自动按加成率调整成本价"
         Height          =   210
         Left            =   4680
         TabIndex        =   20
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox Chk定价 
         Caption         =   "时价药品改为定价"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1770
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "…"
         Height          =   300
         Left            =   3360
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   0
         Width           =   2415
      End
      Begin VB.CheckBox chk按批次 
         Caption         =   "成本价按库房批次调整"
         Height          =   210
         Left            =   10560
         TabIndex        =   8
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk自动计算应付款变动 
         Caption         =   "自动计算应付款变动"
         Height          =   210
         Left            =   12840
         TabIndex        =   7
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "立即执行"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   6
         Top             =   8
         Width           =   1095
      End
      Begin VB.OptionButton opt时间 
         Caption         =   "指定日期执行"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   5
         Top             =   8
         Width           =   1455
      End
      Begin VB.ComboBox cbo售价计算方式 
         Height          =   300
         Left            =   13080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpRunDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   9
         Top             =   0
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   125698051
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lbl调价方式 
         AutoSize        =   -1  'True
         Caption         =   "售价计算方式"
         Height          =   180
         Left            =   11520
         TabIndex        =   22
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         Caption         =   "调价方式"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl执行时间 
         Caption         =   "执行时间"
         Height          =   180
         Left            =   4200
         TabIndex        =   10
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   300
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   46
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
      Appearance      =   1
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   47
      Top             =   4680
      Width           =   3495
      _cx             =   6165
      _cy             =   1720
      Appearance      =   1
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   2295
      Left            =   480
      TabIndex        =   49
      Top             =   2040
      Width           =   11055
      _cx             =   19500
      _cy             =   4048
      Appearance      =   1
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   50
      Top             =   10020
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPriceCard.frx":7158
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20082
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "查找"
      Height          =   255
      Left            =   480
      TabIndex        =   45
      Top             =   9528
      Width           =   495
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "调价流水号"
      Height          =   180
      Left            =   12120
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblDrugName 
      AutoSize        =   -1  'True
      Caption         =   "药品调价单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmMediPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'各种全局变量
Private Const mlngRowHeight As Long = 300 '表格中各行行高
Private mintUnit As Integer     '用来记录启用的是什么单位
Private mint调价 As Integer     '0-调售价;1-调成本价;2-调售价及成本价
Private mlng供应商ID As Long  '用来记录供应商id
Private mdbl加成率 As Double
Private mbln应付记录 As Boolean '记录是否产生应付记录
Private marrSql() As Variant     '纪录按delete键删除药品的存储过程的数组

Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
'颜色方案
Private Const mconlngColor As Long = &HFFFFFF        '不能修改列颜色为白色
Private Const mconlngCanColColor As Long = &HE7CFBA    '能修改列颜色为淡蓝色

Private mbln时价药品按批次调价 As Boolean '时价药品按照批次调价
Private mbln现价提示 As Boolean         '限价药品提示 true-提示 false-不提示
Private mdbl分段加成率 As Double    '用来记录分段加成率
Private mdbl成本价 As Double            '记录修改之前的成本价
Private mrs分段加成 As ADODB.Recordset  '记录分段加成率集合
Private mstrNo As String            '调价单No
Private mintModal As Integer        '本次是什么状态 0-新增 1-修改 2-查阅
Private mintMethod As Integer   '调价方式 0-调售价;1-调成本价;2-调售价及成本价
Private mstr调价汇总号 As String
Private mblnLoad As Boolean     '是否加载完成
Private mrsReturn As ADODB.Recordset '批量选择返回的数据集
Private mblnOK As Boolean
Private mrsFindName As ADODB.Recordset '查询的数据集
Private mBlnClick As Boolean
Private mblnUpdateAdd As Boolean    '修改情况下的新增卫材
Private mlngOldDrugID As Long '检查原始行是否有药品
Private mdblOldPrice As Double   '原售价
Private mblnBatchItem As Boolean   '记录是否点击了批量选择按钮
Private mstrPrivs As String     '操作员权限
Private Const MStrCaption As String = "药品调价单"

Private Enum menuPriceCol
    药品id = 0
    原价id = 1
    品名 = 2
    规格 = 3
    是否变价
    厂牌
    单位
    包装系数
    加成率
    差价让利比
    是否有库存
    收入项目ID
    原成本价
    现成本价
    原零售价
    现零售价
    原采购限价
    现采购限价
    原指导售价
    现指导售价
    总列数
End Enum
Private Enum menuStoreCol
    药品id = 0
    库房 = 1
    库房id = 2
    供应商
    供应商id
    药品
    规格
    批号
    效期
    产地
    批次
    变价
    数量
    单位
    包装系数
    原零售价
    现零售价
    调整金额
    加成率
    原采购价
    现采购价
    差价差
    总列数
End Enum

Private Enum menuPayCol
    药品id = 0
    品名 = 1
    发票号 = 2
    发票日期
    发票金额
    总列数
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str调价汇总号 As String, ByVal intMethod As Integer)
    mintModal = intModal
    mstr调价汇总号 = str调价汇总号
    mintMethod = intMethod

    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "仅调售价" Then
            intTemp = 0
            lbl调价方式.Visible = False
            cbo售价计算方式.Visible = False
        ElseIf .Text = "仅调成本价" Then
            intTemp = 1
            lbl调价方式.Visible = False
            cbo售价计算方式.Visible = False
        Else
            intTemp = 2
            lbl调价方式.Visible = True
            cbo售价计算方式.Visible = True
        End If
    End With


    If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
        If vsfPrice.TextMatrix(1, menuPriceCol.药品id) <> "" Then
            If MsgBox("调价方式改变将清空列表中数据，是否继续？", vbYesNo, gstrSysName) = vbNo Then
                cboPriceMethod.ListIndex = mint调价
                Exit Sub
            Else
                vsfPrice.rows = 2
                For intCol = 0 To vsfPrice.Cols - 1
                    vsfPrice.TextMatrix(1, intCol) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End If
    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
            lblMethod.Tag = 0
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
            lblMethod.Tag = 1
            opt时间(0).Value = True
            opt时间(0).Enabled = False
            opt时间(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
            lblMethod.Tag = 2
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
        If .Text = "仅调售价" Then
            cmdPriceMethod.Visible = False
            picOtherSelect.Visible = cmdPriceMethod.Visible
        Else
            cmdPriceMethod.Visible = True
        End If
    End With
    vsfStore.Cols = menuStoreCol.总列数
    vsfPay.Cols = menuPayCol.总列数
    vsfPrice.Cols = menuPriceCol.总列数
    Call setColEdit
    Call setColHiddenVsf
End Sub

Private Sub cboPriceMethod_DropDown()
    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
        End If
    End With
End Sub

Private Sub cbo售价计算方式_Click()
    On Error GoTo errHandle
    Set mrs分段加成 = Nothing
    If cbo售价计算方式.Text = "售价按分段加成计算" Then
        gstrSQL = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明, 类型 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zlDataBase.OpenSQLRecord(gstrSQL, "药品加成方案")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkAotuCost_Click()
    If chkAotuCost.Value = 1 Then
        cbo售价计算方式.Visible = False
        cbo售价计算方式.ListIndex = 0
        lbl调价方式.Visible = False
    Else
        cbo售价计算方式.Visible = True
        lbl调价方式.Visible = True
    End If
End Sub


Private Sub Chk供应商_Click()
    If chk供应商.Value = 1 Then
        cmd供应商.Enabled = True
        txt供应商.Enabled = True
        chk应付记录.Enabled = True
    Else
        cmd供应商.Enabled = False
        txt供应商.Enabled = False
        chk应付记录.Enabled = False
        chk应付记录.Value = 0
    End If
End Sub

Private Sub chk加成率_Click()
    If chk加成率.Value = 1 Then
        txt加成率.Enabled = True
    Else
        txt加成率.Enabled = False
    End If
End Sub

Private Sub cmdCanc_Click()
    Call ReleaseSelectorRS '卸载数据集
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim intCol As Integer

    If MsgBox("你确定要清空所有数据？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        vsfPrice.rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.rows = 1
        vsfPay.rows = 1
    End If
End Sub

Private Sub cmdFilterCan_Click()
    picOtherSelect.Visible = False
End Sub

Private Sub cmdFilterOk_Click()
    Dim i As Integer

    If chk供应商.Value = 1 Then
        If Val(Split(txt供应商.Tag, "|")(0)) = 0 Then
            MsgBox "请选择供应商。", vbInformation, gstrSysName
            txt供应商.SetFocus
            Exit Sub
        End If
    End If
    With vsfPrice
        If Val(.TextMatrix(1, menuPriceCol.药品id)) <> 0 Then
            If MsgBox("将清空表格中的数据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End With

    mlng供应商ID = IIf(chk供应商.Value = 1, Val(Split(txt供应商.Tag, "|")(0)), 0)
    mdbl加成率 = IIf(chk加成率.Value = 1, Val(Trim(txt加成率.Text)), 0)
    mbln应付记录 = (chk应付记录.Enabled And chk应付记录.Value = 1)
    picOtherSelect.Visible = False
    If mbln应付记录 = True Then
        TabCtlDetails.Item(1).Visible = True
    Else
        TabCtlDetails.Item(1).Visible = False
    End If

    With cboPriceMethod
        If .Text = "仅调售价" Then
            mint调价 = 0
            lblMethod.Tag = 0
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "仅调成本价" Then
            mint调价 = 1
            lblMethod.Tag = 1
            opt时间(0).Value = True
            opt时间(0).Enabled = False
            opt时间(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "售价成本价一起调价" Then
            mint调价 = 2
            lblMethod.Tag = 2
            opt时间(0).Value = False
            opt时间(1).Value = True
            opt时间(0).Enabled = True
            opt时间(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mbln应付记录 = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
    End With

End Sub

Private Sub CmdHelp_Click()

End Sub

Private Sub cmdItem_Click()
    Dim intRow As Integer

    frmBatchSelect.ShowME Me, mrsReturn, mblnOK

    On Error GoTo errHandle
    If mblnOK = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.rows - 1, menuPriceCol.药品id) = "" Then
            intRow = .rows - 1
        Else
            .rows = .rows + 1
            intRow = .rows - 1
        End If
    End With
    mblnBatchItem = True

    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '清除未执行价格
    Dim intRow As Integer

    On Error GoTo errHandle
    With vsfPrice
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, menuPriceCol.药品id) <> "" Then
                gstrSQL = "Zl_删除未执行价格_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.药品id)) & "," & 0 & ")"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With
    
    '删除delete删除的数据
    For intRow = 0 To UBound(marrSql)
        Call zlDataBase.ExecuteProcedure(CStr(marrSql(intRow)), Me.Caption)
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToday As Date
    Dim lngAdjId As Long
    Dim LngCurID As Long
    Dim strID As String
    Dim intCount As Integer
    Dim dbl包装 As Double
    Dim strTmp As String
    Dim lngCurrBatch As Long
    Dim str批次价格 As String
    Dim blnPrint As Boolean '是否打印调价通知单
    Dim blnOne As Boolean   '检查是否是第一行
    Dim n As Integer
    Dim intProc As Integer
    Dim blnIgnore As Boolean
    Dim blnPrice As Boolean '记录是否售价调价了
    Dim blnCost As Boolean  '记录是否成本价调价了
    Dim intUpdateModel As Integer '调价模式 0-售价调价 1-成本价调价 2-成本价售价一起调价
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim ArrayID
    Dim Array批次价格
    Dim strUpdate As String

    Dim lng库房ID As Long
    Dim lng供应商ID As Long
    Dim lng药品ID As Long
    Dim lng批次  As Long
    Dim str批号 As String
    Dim str效期 As String
    Dim str产地 As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim Str发票号 As String
    Dim str发票日期 As String
    Dim dbl发票金额 As Double
    Dim strInfo As String
    Dim strMsg As String '记录提示信息
    Dim intCount2 As Integer '用来计数
    Dim lngDouID As Long

    If vsfPrice.rows > 1 Then   '只有有数据的情况下才能保存
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.药品id)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub

    On Error GoTo ErrHand
    dtToday = zlDataBase.Currentdate()

    gstrSQL = "select 收费价目_ID.nextval from dual"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取收费价目序号")
    lngAdjId = rsTemp.Fields(0).Value

    gcnOracle.BeginTrans
    If mintModal = 1 Then '修改 在修改模式下先删除原来的调价信息，然后插入新的调价信息
        Call deleteNotExecutePirce
    End If

    '检查是否存在未执行的价格
    If checkNotExecutePrice(, strInfo) = True Then
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Sub
    End If
    '获取调价NO
    mstrNo = zlDataBase.GetNextNo(9)
    '获取调价汇总NO
    gstrSQL = "select nextno(135) as 流水号 from dual"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "调价流水号")
    If rsTemp.RecordCount = 0 Then
        MsgBox "调价流水号未能初始化成功，请与管理员联系！", vbInformation, gstrSysName
        Exit Sub
    End If
    txtNO.Text = rsTemp!流水号

    With Me.vsfPrice
        '售价调价
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            If mint调价 <> 1 Then

                LngCurID = zlDataBase.GetNextId("收费价目")
                strID = strID & IIf(strID = "", "", ",") & LngCurID

                dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

                If .TextMatrix(intCount, menuPriceCol.是否变价) = "1" And mbln时价药品按批次调价 And mint调价 <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(intCount, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                            If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.批次) & ",") = 0 Then
                                lngCurrBatch = vsfStore.TextMatrix(n, menuStoreCol.批次)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.批次) & "," & vsfStore.TextMatrix(n, menuStoreCol.现零售价) / dbl包装
                            End If
                        End If
                    Next
                    str批次价格 = str批次价格 & strTmp
                End If
                str批次价格 = str批次价格 & ";"

                If CLng(.TextMatrix(intCount, menuPriceCol.原价id)) <> 0 Then
                    '设置上一次的价格记录终止执行
                    gstrSQL = "zl_收费价目_stop(" & .TextMatrix(intCount, menuPriceCol.药品id) & ","
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

                    '产生价格记录
                    gstrSQL = "zl_收费价目_Insert(" & LngCurID & "," & IIf(.TextMatrix(intCount, menuPriceCol.原价id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.原价id))) & _
                              "," & .TextMatrix(intCount, menuPriceCol.药品id) & "," & Val(.TextMatrix(intCount, menuPriceCol.收入项目ID)) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.原零售价)) / dbl包装, gtype_UserDrugDigits.Digit_零售价) & "," & _
                              Round(IIf(Val(.TextMatrix(intCount, menuPriceCol.现零售价)) = Val(.TextMatrix(intCount, menuPriceCol.原零售价)), Val(.TextMatrix(intCount, menuPriceCol.现零售价)) + 1, Val(.TextMatrix(intCount, menuPriceCol.现零售价))) / dbl包装, gtype_UserDrugDigits.Digit_零售价) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If opt时间(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ",0,'" & mstrNo & "'," & intCount & ",Null," & txtNO & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnPrice = True
                    blnPrint = True
                End If
            End If
        Next
    End With

    '成本价调价处理
    If mint调价 = 1 Or mint调价 = 2 Then
        If vsfStore.rows > 1 Then
            If vsfStore.TextMatrix(1, menuStoreCol.药品id) <> "" Then
'                lngDouID = 0
'                For n = 1 To vsfStore.rows - 1
'                    If vsfStore.TextMatrix(n, menuStoreCol.药品id) = "" Then Exit For
'
'                    '检查未审核单据
'                    If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) <> lngDouID Then
'                        lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
'                        strMsg = vsfStore.TextMatrix(n, menuStoreCol.药品) & ","
'                        intCount2 = intCount2 + 1
'                        If intCount2 > 3 Then Exit For '只判断3个
'                    End If
'                Next
'
'                If strMsg <> "" Then
'                    If MsgBox(strMsg & "存在未审核单据，调整成本价可能会造成差价误差。" & _
'                        vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        gcnOracle.RollbackTrans
'                        Exit Sub
'                    End If
'                End If

                For n = 1 To vsfStore.rows - 1
                    For i = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(i, 0) = "" Then Exit For
                        If Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) = Val(vsfPay.TextMatrix(i, menuPayCol.药品id)) Then
                            lng库房ID = Val(vsfStore.TextMatrix(n, menuStoreCol.库房id))
                            lng供应商ID = Val(vsfStore.TextMatrix(n, menuStoreCol.供应商id))
                            lng药品ID = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
                            lng批次 = Val(vsfStore.TextMatrix(n, menuStoreCol.批次))
                            str批号 = vsfStore.TextMatrix(n, menuStoreCol.批号)
                            str效期 = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.效期)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.效期))
                            str产地 = vsfStore.TextMatrix(n, menuStoreCol.产地)
                            dblOldCost = GetFormat(Val(vsfStore.TextMatrix(n, menuStoreCol.原采购价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数)), gtype_UserDrugDigits.Digit_成本价)
                            dblNewCost = GetFormat(Val(vsfStore.TextMatrix(n, menuStoreCol.现采购价)) / Val(vsfStore.TextMatrix(n, menuStoreCol.包装系数)), gtype_UserDrugDigits.Digit_成本价)
                            Str发票号 = vsfPay.TextMatrix(i, menuPayCol.发票号)
                            str发票日期 = Format(vsfPay.TextMatrix(i, menuPayCol.发票日期), "yyyy-mm-dd")
                            dbl发票金额 = Val(vsfPay.TextMatrix(i, menuPayCol.发票金额))

                            gstrSQL = "Zl_成本价调价信息_Insert(" & IIf(lng供应商ID = 0, "Null", lng供应商ID) & "," & lng库房ID & "," & lng药品ID & "," & lng批次 & ",'" & str批号 & "'" & _
                                    "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str产地 & "',Null," & dblOldCost & ", " & dblNewCost & "," & _
                                    IIf(Str发票号 <> "", "'" & Str发票号 & "'", "NULL") & "," & IIf(str发票日期 = "", "Null", "to_date('" & Format(str发票日期, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl发票金额 & "," & IIf(mbln应付记录 = True, 1, 0) & "," & txtNO.Text & ")"
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                            blnCost = True
                        End If
                    Next
                Next
            End If
        End If
    End If

    '无库存时调整成本价
    If mint调价 = 1 Or mint调价 = 2 Then
        With Me.vsfPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                If .TextMatrix(intCount, menuPriceCol.是否有库存) = "0" And Val(.TextMatrix(intCount, menuPriceCol.原成本价)) <> Val(.TextMatrix(intCount, menuPriceCol.现成本价)) Then
                    dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

                    lng药品ID = Val(.TextMatrix(intCount, menuPriceCol.药品id))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.原成本价)) / dbl包装, gtype_UserDrugDigits.Digit_成本价))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现成本价)) / dbl包装, gtype_UserDrugDigits.Digit_成本价))

                    gstrSQL = "Zl_成本价调价信息_Insert(Null,Null," & lng药品ID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, " & txtNO.Text & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnCost = True
                End If
            Next
        End With
    End If

    '立即执行
    If mint调价 = 1 Then
        '单独成本价调价时
        If opt时间(0).Value = True Then
            With Me.vsfPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                    gstrSQL = "zl_药品收发记录_Adjust(0,0,Null," & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                Next
            End With
        End If
    Else
        '调售价
        ArrayID = Split(strID, ",")
        Array批次价格 = Split(str批次价格, ";")
        For intCount = 0 To UBound(ArrayID)
            If opt时间(0).Value = True Or vsfPrice.TextMatrix(intCount + 1, menuPriceCol.原价id) = "" Then
                gstrSQL = "zl_药品收发记录_Adjust(" & ArrayID(intCount) & "," & Me.Chk定价.Value & ",'" & Array批次价格(intCount) & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End If

    '调整指导价格
    With Me.vsfPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            dbl包装 = Val(.TextMatrix(intCount, menuPriceCol.包装系数))

            '更新指导零售价
            If Val(.TextMatrix(intCount, menuPriceCol.原指导售价)) < Val(.TextMatrix(intCount, menuPriceCol.现零售价)) And Val(.TextMatrix(intCount, menuPriceCol.原指导售价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现指导售价)) / dbl包装, gtype_UserDrugDigits.Digit_零售价))

                gstrSQL = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ",'指导零售价=" & strUpdate & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If

            '更新采购限价
            If Val(.TextMatrix(intCount, menuPriceCol.原采购限价)) < Val(.TextMatrix(intCount, menuPriceCol.现成本价)) And Val(.TextMatrix(intCount, menuPriceCol.原采购限价)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.现采购限价)) / dbl包装, gtype_UserDrugDigits.Digit_成本价))

                gstrSQL = "zl_药品目录_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.药品id)) & ",'指导批发价=" & strUpdate & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    '产生调价汇总记录
    If blnPrice = True And blnCost = True Then
        intUpdateModel = 2
    ElseIf blnPrice = True And blnCost = False Then
        intUpdateModel = 0
    ElseIf blnPrice = False And blnCost = True Then
        intUpdateModel = 1
    End If

    gstrSQL = "Zl_调价汇总记录_Insert(" & txtNO.Text & "," & intUpdateModel & ","
    If opt时间(0).Value = True Then
        gstrSQL = gstrSQL & "sysdate" & ","
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    End If
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",0,'" & UserInfo.用户姓名 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("你需要打印调价通知单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1333", Me, "NO=" & txtNO.Text, "包装单位=" & mintUnit, 2)
        End If
    End If

    '清空列表中数据
    With vsfPrice
        .rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.rows = 1
    vsfPay.rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""

    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lng药品ID As Long) As Boolean
    '检查药品是否存在未审核单据
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 1 From 药品收发记录 Where 药品id = [1] And Rownum = 1 And 审核日期 Is Null"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查药品是否存在未审核单据", lng药品ID)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0, Optional ByRef strInfo As String) As Boolean
    '功能 ：检查是否存在未执行的价格
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    Err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        '循环判断所有药品
        For IntCheck = 1 To vsfPrice.rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.药品id))
            If LngmediIDThis <> 0 Then
                If mint调价 = 0 Or mint调价 = 2 Then
                    '判断是否有未执行的历史价格
                    gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]"
                    Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    strInfo = "药品" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.品名) & "存在未执行价格，未执行药品不能调价！"
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint调价 = 1 Or mint调价 = 2 Then
                    '检查是否还有未执行的成本价调价计划
                    gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
                    Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        strInfo = "药品" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.品名) & "存在未执行成本价，未执行药品不能调价！"
                        checkNotExecutePrice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint调价 = 0 Or mint调价 = 2 Then
            '判断是否有未执行的历史价格
            gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 执行日期 > Sysdate And 收费细目ID=[1]"
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            strInfo = "还存在未执行的售价调价记录，未执行药品不能调价！"
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint调价 = 1 Or mint调价 = 2 Then
            '检查是否还有未执行的成本价调价计划
            gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
                strInfo = "还存在未执行的成本价调价，未执行药品不能调价！"
                checkNotExecutePrice = True
                Exit Function
            End If
        End If
    End If


    checkNotExecutePrice = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.vsfPrice.SetFocus

End Function

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim strTmp As String
    Dim bln无库存 As Boolean
    Dim dbl包装 As Double
    Dim bln有无库存 As Boolean
    Dim lngDouID As Long
    Dim strMsg As String '记录提示信息
    Dim intCount2 As Integer '用来计数
    
    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.现零售价))) Then
                    MsgBox "第" & IntCheck & "行的药品售价中含有非法字符！", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.现零售价
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

                '检查价格是否为空
                If .TextMatrix(IntCheck, menuPriceCol.现零售价) = "" Or .TextMatrix(IntCheck, menuPriceCol.原零售价) = "" Or .TextMatrix(IntCheck, menuPriceCol.现成本价) = "" Or .TextMatrix(IntCheck, menuPriceCol.原成本价) = "" Then
                    MsgBox "第" & IntCheck & "行的药品有价格为空，不能执行调价！", vbInformation, gstrSysName
                    .Row = IntCheck
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                For n = 1 To vsfStore.rows - 1
                    If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                        If vsfStore.TextMatrix(n, menuStoreCol.现零售价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.原零售价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.现采购价) = "" Or vsfStore.TextMatrix(n, menuStoreCol.原采购价) = "" Then
                            MsgBox "第" & IntCheck & "行的药品有价格为空，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                Next
                
                '检查售价是否相同
                If mint调价 = 0 Or mint调价 = 2 Then
                    strTmp = ""
                    bln有无库存 = False
                    dbl包装 = Val(.TextMatrix(IntCheck, menuPriceCol.包装系数))
                    If .TextMatrix(IntCheck, menuPriceCol.是否变价) = "1" Then
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                                bln有无库存 = True
                                If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.批次) & ",") = 0 And vsfStore.TextMatrix(n, menuStoreCol.现零售价) <> vsfStore.TextMatrix(n, menuStoreCol.原零售价) Then
                                    strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.批次) & "," & vsfStore.TextMatrix(n, menuStoreCol.现零售价) / dbl包装
                                End If
                            End If
                        Next
                        If strTmp = "" And bln有无库存 = True Then
                            MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.现零售价
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                        If bln有无库存 = False And .TextMatrix(IntCheck, menuPriceCol.现零售价) = .TextMatrix(IntCheck, menuPriceCol.原零售价) Then
                            MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.现零售价
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                    If .TextMatrix(IntCheck, menuPriceCol.是否变价) <> "1" And .TextMatrix(IntCheck, menuPriceCol.现零售价) = .TextMatrix(IntCheck, menuPriceCol.原零售价) Then
                        MsgBox "第" & IntCheck & "行的药品现零售价与原零售价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现零售价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '检查成本价是否相同
                If mint调价 = 1 Or mint调价 = 2 Then
                    bln有无库存 = False
                    strTmp = ""
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(IntCheck, menuPriceCol.药品id)) = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) Then
                            bln有无库存 = True
                            If vsfStore.TextMatrix(n, menuStoreCol.现采购价) <> vsfStore.TextMatrix(n, menuStoreCol.原采购价) Then
                                strTmp = "调过成本价"
                            End If
                        End If
                    Next
                    If bln有无库存 = True And strTmp = "" Then
                        MsgBox "第" & IntCheck & "行的药品现采购价与原采购价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现成本价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                    If bln有无库存 = False And .TextMatrix(IntCheck, menuPriceCol.现成本价) = .TextMatrix(IntCheck, menuPriceCol.原成本价) Then
                        MsgBox "第" & IntCheck & "行的药品现成本价与原成本价相同，不能执行调价！", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.现成本价
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If

                If .TextMatrix(IntCheck, menuPriceCol.是否变价) = "1" And opt时间(0).Value <> True And mint调价 <> 1 Then
                    MsgBox "第" & IntCheck & "行为时价药品，必须设置为立即执行！", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.现零售价
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

            End If
        Next
    End With

    '检查未审核单据
    If vsfStore.rows > 1 And (mint调价 = 1 Or mint调价 = 2) Then
        If vsfStore.TextMatrix(1, menuStoreCol.药品id) <> "" Then
            lngDouID = 0
            For n = 1 To vsfStore.rows - 1
                If vsfStore.TextMatrix(n, menuStoreCol.药品id) = "" Then Exit For
    
                If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.药品id)) <> lngDouID Then
                    lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.药品id))
                    strMsg = strMsg & vsfStore.TextMatrix(n, menuStoreCol.药品) & ","
                    intCount2 = intCount2 + 1
                    If intCount2 > 3 Then Exit For '只判断3个
                End If
            Next
    
            If strMsg <> "" Then
                If MsgBox(strMsg & "存在未审核单据，调整成本价可能会造成差价误差。" & _
                    vbCrLf & Space(4) & "建议先处理未审核单据。是否还继续调价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt供应商.Tag = "" Then
        Me.txt供应商.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If vsfStore.rows = 1 Then Exit Sub
    If Trim(Me.vsfStore.TextMatrix(1, menuStoreCol.库房)) = "" Then Exit Sub

    objPrint.Title.Text = "调价库存变动表"

    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(opt时间(0).Value = True, zlDataBase.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(zlDataBase.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfStore.Object
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
End Sub

Private Sub Cmd供应商_Click()
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 编码,名称,简码,id" & _
        " From 供应商" & _
        " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By 编码 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取供应商信息")
    If rsTemp.EOF Then
        MsgBox "请初始化供应商（字典管理）！", vbInformation, gstrSysName
        Exit Sub
    End If

    With Me.mshProvider
        .Left = chk供应商.Left
        .Top = txt供应商.Top + txt供应商.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mBlnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.品名
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picOtherSelect.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    Dim intUnitTemp As Integer

    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '获取设置的单位
    mintUnit = Val(zlDataBase.GetPara("药品单位", glngSys, 1333, "1"))
    mstrPrivs = GetPrivFunc(glngSys, 1333)
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 1 '住院
            intUnitTemp = 3
        Case 2 '门诊
            intUnitTemp = 2
        Case 3 '售价
            intUnitTemp = 1
    End Select
    '获取各级单位精度
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    '初始化时间为当前时间+1天
    StrToday = Format(zlDataBase.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    
    If mintModal = 0 Then '新增的时候最小时间设置为当前时间+1天
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    mbln时价药品按批次调价 = Val(zlDataBase.GetPara("时价药品按批次调价", glngSys, 1333, 0))
    mbln现价提示 = Val(zlDataBase.GetPara("限价提示", glngSys, 1333, 1))
    
    marrSql = Array()
    
    txtValuer.Text = UserInfo.用户姓名  'gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr调价汇总号)
    If mintModal = 0 Then
        lblNO.Visible = False
        txtNO.Visible = False
    End If

    Call initComboBox '初始化下拉控件
    If mintModal = 1 Then '修改
        If (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0) Or (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '查阅
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call InitTabControl
    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    If mbln应付记录 = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    If mintModal <> 0 Then
        Call initGrid
    End If

    If mintModal = 2 Then '查阅
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        opt时间(0).Enabled = False
        opt时间(1).Enabled = False
        dtpRunDate.Enabled = False
        cbo售价计算方式.Enabled = False
        Chk定价.Enabled = False
        chkCostBatch.Enabled = False
        chkAotuCost.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False
        cmdClear.Visible = False
        cmdItem.Visible = False
        cmdOk.Visible = False
        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
End Sub

Private Sub initComboBox()
    With cbo售价计算方式
        .AddItem "售价与成本价不关联计算"
        .AddItem "售价按固定比例计算"
        .AddItem "售价按分段加成计算"
        .ListIndex = 0
    End With

    With cboPriceMethod
        If mintModal <> 2 Then  '非查阅
            If InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") = 0 Then
                .AddItem "仅调成本价"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 Then
                .AddItem "仅调售价"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";成本价调价;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";售价调价;") > 0 Then
                .AddItem "仅调售价"
                .AddItem "仅调成本价"
                .AddItem "售价成本价一起调价"
                .ListIndex = 0
                lblMethod.Tag = 0
            End If
        Else
            .AddItem "仅调售价"
            .AddItem "仅调成本价"
            .AddItem "售价成本价一起调价"
            .ListIndex = 0
            lblMethod.Tag = 0
        End If
    End With
End Sub

Private Sub InitTabControl()
    '初始化TabControl控件
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    With TabCtlDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "库存变动表", vsfStore.hWnd, 0
        .InsertItem 1, "应付款变动表", vsfPay.hWnd, 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsfPrice.Height - vsfPrice.Top - 20
        .Top = picSplit.Height + picSplit.Top + 20
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If

    With fraCondition
        .Width = Me.ScaleWidth
    End With
    txtNO.Left = Me.ScaleWidth - txtNO.Width
    lblNO.Left = txtNO.Left - lblNO.Width - 200
    lblDrugName.Left = Me.ScaleWidth / 2 - lblDrugName.Width / 2
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - cmdClear.Height - 300 - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
    lblFind.Top = picInfo.Top + picInfo.Height + 180
    lblFind.Left = picInfo.Left + 380
    txtFind.Top = lblFind.Top - 50
    txtFind.Left = 985
    cmdClear.Top = txtFind.Top
    cmdItem.Top = txtFind.Top
    cmdPrint.Top = txtFind.Top
    cmdOk.Top = txtFind.Top
    cmdCanc.Top = txtFind.Top
    cmdCanc.Left = Me.ScaleWidth - cmdCanc.Width - 300
    cmdOk.Left = cmdCanc.Left - cmdOk.Width - 200
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 500
    cmdItem.Left = cmdPrint.Left - cmdPrint.Width - 20
    cmdClear.Left = cmdItem.Left - cmdItem.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
    Call SaveWinState(Me, App.ProductName, MStrCaption)
    mblnLoad = False
    mbln应付记录 = False
    mlng供应商ID = 0
    mblnUpdateAdd = False
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt供应商.Text = .TextMatrix(.Row, 1)
        Me.txt供应商.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With

    Me.txt供应商.SetFocus
End Sub

Private Sub opt时间_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub InitVsfGridFlex()
    With vsfPrice

        .Cols = menuPriceCol.总列数
        .rows = 2
        .RowHeight(1) = mlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .Editable = flexEDNone
'        .GridLineWidth = 2
'        .GridLines = flexGridInset
'        .GridColor = &H80000011
'        .GridColorFixed = &H80000011
'        .ForeColorFixed = &H80000012
'        .BackColorSel = &HF4F4EA

        .TextMatrix(0, menuPriceCol.药品id) = "药品ID"
        .TextMatrix(0, menuPriceCol.原价id) = "原价id"
        .TextMatrix(0, menuPriceCol.品名) = "品名"
        .TextMatrix(0, menuPriceCol.规格) = "规格"
        .TextMatrix(0, menuPriceCol.是否变价) = "是否变价"
        .TextMatrix(0, menuPriceCol.厂牌) = "厂牌"
        .TextMatrix(0, menuPriceCol.单位) = "单位"
        .TextMatrix(0, menuPriceCol.包装系数) = "包装系数"
        .TextMatrix(0, menuPriceCol.加成率) = "加成率"
        .TextMatrix(0, menuPriceCol.差价让利比) = "差价让利比"
        .TextMatrix(0, menuPriceCol.是否有库存) = "是否有库存"
        .TextMatrix(0, menuPriceCol.收入项目ID) = "收入项目id"
        .TextMatrix(0, menuPriceCol.原成本价) = "原成本价"
        .TextMatrix(0, menuPriceCol.现成本价) = "现成本价"
        .TextMatrix(0, menuPriceCol.原零售价) = "原零售价"
        .TextMatrix(0, menuPriceCol.现零售价) = "现零售价"
        .TextMatrix(0, menuPriceCol.原采购限价) = "原采购限价"
        .TextMatrix(0, menuPriceCol.现采购限价) = "现采购限价"
        .TextMatrix(0, menuPriceCol.原指导售价) = "原指导售价"
        .TextMatrix(0, menuPriceCol.现指导售价) = "现指导售价"

        '设置列宽
        .ColWidth(menuPriceCol.药品id) = 0
        .ColWidth(menuPriceCol.原价id) = 0
        .ColWidth(menuPriceCol.品名) = 3000
        .ColWidth(menuPriceCol.规格) = 1500
        .ColWidth(menuPriceCol.是否变价) = 0
        .ColWidth(menuPriceCol.厂牌) = 2000
        .ColWidth(menuPriceCol.单位) = 800
        .ColWidth(menuPriceCol.包装系数) = 0
        .ColWidth(menuPriceCol.加成率) = 0
        .ColWidth(menuPriceCol.差价让利比) = 0
        .ColWidth(menuPriceCol.是否有库存) = 0
        .ColWidth(menuPriceCol.收入项目ID) = 0
        .ColWidth(menuPriceCol.原成本价) = 1000
        .ColWidth(menuPriceCol.现成本价) = 1000
        .ColWidth(menuPriceCol.原零售价) = 1000
        .ColWidth(menuPriceCol.现零售价) = 1000
        .ColWidth(menuPriceCol.原采购限价) = 0
        .ColWidth(menuPriceCol.现采购限价) = 0
        .ColWidth(menuPriceCol.原指导售价) = 0
        .ColWidth(menuPriceCol.现指导售价) = 0
        '设置对齐方式
        .ColAlignment(menuPriceCol.品名) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.厂牌) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.原成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现成本价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.现零售价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原采购限价) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.原指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
        .ColComboList(menuPriceCol.品名) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.总列数
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '设置列名
        .TextMatrix(0, menuStoreCol.药品id) = "药品id"
        .TextMatrix(0, menuStoreCol.库房) = "库房"
        .TextMatrix(0, menuStoreCol.库房id) = "库房id"
        .TextMatrix(0, menuStoreCol.供应商) = "供应商"
        .TextMatrix(0, menuStoreCol.供应商id) = "供应商id"
        .TextMatrix(0, menuStoreCol.药品) = "药品"
        .TextMatrix(0, menuStoreCol.规格) = "规格"
        .TextMatrix(0, menuStoreCol.单位) = "单位"
        .TextMatrix(0, menuStoreCol.批号) = "批号"
        .TextMatrix(0, menuStoreCol.效期) = "效期"
        .TextMatrix(0, menuStoreCol.产地) = "产地"
        .TextMatrix(0, menuStoreCol.数量) = "数量"
        .TextMatrix(0, menuStoreCol.包装系数) = "包装系数"
        .TextMatrix(0, menuStoreCol.批次) = "批次"
        .TextMatrix(0, menuStoreCol.变价) = "变价"
        .TextMatrix(0, menuStoreCol.原零售价) = "原零售价"
        .TextMatrix(0, menuStoreCol.现零售价) = "现零售价"
        .TextMatrix(0, menuStoreCol.调整金额) = "调整金额"
        .TextMatrix(0, menuStoreCol.加成率) = "加成率"
        .TextMatrix(0, menuStoreCol.原采购价) = "原采购价"
        .TextMatrix(0, menuStoreCol.现采购价) = "现采购价"
        .TextMatrix(0, menuStoreCol.差价差) = "差价差"
        '设置列宽
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.库房) = 1500
        .ColWidth(menuStoreCol.库房id) = 0
        .ColWidth(menuStoreCol.供应商) = 2000
        .ColWidth(menuStoreCol.供应商id) = 0
        .ColWidth(menuStoreCol.药品) = 3000
        .ColWidth(menuStoreCol.规格) = 1500
        .ColWidth(menuStoreCol.单位) = 800
        .ColWidth(menuStoreCol.批号) = 1500
        .ColWidth(menuStoreCol.效期) = 2000
        .ColWidth(menuStoreCol.产地) = 1500
        .ColWidth(menuStoreCol.数量) = 1500
        .ColWidth(menuStoreCol.包装系数) = 0
        .ColWidth(menuStoreCol.批次) = 0
        .ColWidth(menuStoreCol.变价) = 0
        .ColWidth(menuStoreCol.原零售价) = 1000
        .ColWidth(menuStoreCol.现零售价) = 1000
        .ColWidth(menuStoreCol.调整金额) = 1000
        .ColWidth(menuStoreCol.加成率) = 1000
        .ColWidth(menuStoreCol.原采购价) = 1000
        .ColWidth(menuStoreCol.现采购价) = 1000
        .ColWidth(menuStoreCol.差价差) = 1000
        '对齐方式
        .ColAlignment(menuStoreCol.库房) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.供应商) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.药品) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.规格) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.单位) = flexAlignCenterCenter
        .ColAlignment(menuStoreCol.批号) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.效期) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.产地) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.数量) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.原零售价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.现零售价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.调整金额) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.加成率) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.原采购价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.现采购价) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.差价差) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.总列数
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
'        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMoveRows '拖动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.药品id) = "药品id"
        .TextMatrix(0, menuPayCol.品名) = "品名"
        .TextMatrix(0, menuPayCol.发票号) = "发票号"
        .TextMatrix(0, menuPayCol.发票日期) = "发票日期"
        .TextMatrix(0, menuPayCol.发票金额) = "发票金额"
        '设置列宽
        .ColWidth(menuPayCol.药品id) = 0
        .ColWidth(menuPayCol.品名) = 2000
        .ColWidth(menuPayCol.发票号) = 1500
        .ColWidth(menuPayCol.发票日期) = 2000
        .ColWidth(menuPayCol.发票金额) = 1500
        '对齐方式
        .ColAlignment(menuPayCol.品名) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票号) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票日期) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.发票金额) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '列头居中对齐
    End With
End Sub

Private Sub initGrid()
    '如果是修改或者查阅则提取相应的记录并填充到表格中
    Dim rsTemp As ADODB.Recordset
    Dim intRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db包装系数 As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs产地 As ADODB.Recordset

    On Error GoTo errHandle
    '调价方式 0-调售价;1-调成本价;2-调售价及成本价
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                nvl(s.加成率,0) / 100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 厂牌, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, s.成本价 As 原成本价, s.成本价 As 新成本价, p.原价, p.现价," & vbNewLine & _
            "                p.收入项目id, p.调价人, p.调价说明, s.差价让利比, To_Char(a.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id From 药品库存 where 性质=1) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where a.调价号 = p.调价汇总号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And i.Id = p.收费细目id And" & vbNewLine & _
            "      p.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And a.调价号 = [1] " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "                nvl(s.加成率,0) / 100As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 厂牌, i.计算单位 As 单位," & vbNewLine & _
            "                s.门诊单位, s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.现价 as 原价, p.现价, p.收入项目id," & vbNewLine & _
            "                a.填制人 As 调价人, a.说明 As 调价说明, s.差价让利比, To_Char(m.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id," & vbNewLine & _
            "                Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号,药品id,min(执行日期) as 执行日期 From 成本价调价信息 Where 调价汇总号 = [1] Group By 调价汇总号,药品id) M, (Select 药品id From 药品库存 where 性质=1) K, 调价汇总记录 A, 收费项目别名 B, 药品规格 S, 收费项目目录 I, 收费价目 P" & vbNewLine & _
            "Where m.调价汇总号(+) = a.调价号 And b.收费细目id(+) = s.药品id And s.药品id = i.Id And i.Id = k.药品id(+) And m.药品id = i.Id And" & vbNewLine & _
            "      i.Id = p.收费细目id And Sysdate Between p.执行日期 And p.终止日期 And m.调价汇总号 = [1] And a.分类 = 0 And b.性质(+) = 3 And" & vbNewLine & _
            "      a.调价号 = [1] " & IIf(mintModal = 2, "", " And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By 药品id"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.原价id, i.是否变价, Nvl(s.指导批发价, 0) As 指导批价, Nvl(s.扣率, 0) As 扣率, Nvl(s.指导零售价, 0) As 指导售价," & vbNewLine & _
            "       nvl(s.加成率,0) / 100 As 加成率, i.编码, b.名称 As 商品名, i.名称 As 通用名, i.规格, i.产地 As 厂牌, i.计算单位 As 单位, s.门诊单位," & vbNewLine & _
            "       s.门诊包装, s.住院单位, s.住院包装, s.药库单位, Nvl(s.药库包装, 1) 药库包装, m.原成本价, m.新成本价, p.原价, p.现价, p.收入项目id, p.调价人, p.调价说明, s.差价让利比," & vbNewLine & _
            "       To_Char(p.执行日期, 'YYYY-MM-DD HH24:MI:SS') As 执行日期, i.Id 药品id, Decode(k.药品id, Null, 0, 1) 是否有库存" & vbNewLine & _
            "From (Select 药品id,Min(原成本价) As 原成本价, Min(新成本价) As 新成本价, min(产地) as 产地,调价汇总号 From 成本价调价信息 Where 调价汇总号 = [1] Group By 药品id,调价汇总号) M, 收费价目 P, 调价汇总记录 A, (Select 药品id From 药品库存 where 性质=1) K, 收费项目别名 B, 药品规格 S, 收费项目目录 I" & vbNewLine & _
            "Where m.调价汇总号 = a.调价号 and m.药品id=i.id And p.调价汇总号 = a.调价号 And p.收费细目id = k.药品id(+) And p.收费细目id = b.收费细目id(+) And p.收费细目id = s.药品id And" & vbNewLine & _
            "      s.药品id = i.Id And a.调价号 =[1] And b.性质(+) = 3 " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.撤档时间 Is Null Or i.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))") & "Order By 药品id "
    End If
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr调价汇总号)
    If rsTemp.RecordCount = 0 Then
        MsgBox "该调价记录已经被删除了！", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!药品id <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db包装系数 = rsTemp!药库包装
                        strUnit = rsTemp!药库单位
                    Case 1
                        db包装系数 = rsTemp!住院包装
                        strUnit = rsTemp!住院单位
                    Case 2
                        db包装系数 = rsTemp!门诊包装
                        strUnit = rsTemp!门诊单位
                    Case 3
                        db包装系数 = 1
                        strUnit = rsTemp!单位
                End Select

                lngDrugID = rsTemp!药品id
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.原价id) = IIf(IsNull(rsTemp!原价id), "", rsTemp!原价id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.药品id) = rsTemp!药品id

                If gint药品名称显示 = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.品名) = "[" & rsTemp!编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.品名) = "[" & rsTemp!编码 & "]" & rsTemp!通用名
                End If
                .TextMatrix(.rows - 1, menuPriceCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, menuPriceCol.是否变价) = rsTemp!是否变价
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(产地) as 厂牌 from 成本价调价信息 where 调价汇总号=[1] and 药品id=[2]"
'                    Set rs产地 = zldatabase.OpenSQLRecord(gstrSQL, "产地查询", mstr调价汇总号, rsTemp!药品ID)
'                    If rs产地.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.厂牌) = IIf(IsNull(rs产地!厂牌), "", rs产地!厂牌)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.厂牌) = IIf(IsNull(rsTemp!厂牌), "", rsTemp!厂牌)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.单位) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.包装系数) = db包装系数

                .TextMatrix(.rows - 1, menuPriceCol.加成率) = rsTemp!加成率
                .TextMatrix(.rows - 1, menuPriceCol.差价让利比) = Nvl(rsTemp!差价让利比, 100)
                .TextMatrix(.rows - 1, menuPriceCol.是否有库存) = rsTemp!是否有库存
                .TextMatrix(.rows - 1, menuPriceCol.收入项目ID) = IIf(IsNull(rsTemp!收入项目ID), "", rsTemp!收入项目ID)
                .TextMatrix(.rows - 1, menuPriceCol.原成本价) = GetFormat(Nvl(rsTemp!原成本价, 0) * db包装系数, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.现成本价) = GetFormat(rsTemp!新成本价 * db包装系数, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.原零售价) = GetFormat(IIf(IsNull(rsTemp!原价), rsTemp!现价, rsTemp!原价) * db包装系数, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.现零售价) = GetFormat(rsTemp!现价 * db包装系数, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.原采购限价) = GetFormat(rsTemp!指导批价 * db包装系数, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.现采购限价) = GetFormat(rsTemp!指导批价 * db包装系数, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.原指导售价) = GetFormat(rsTemp!指导售价 * db包装系数, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.现指导售价) = GetFormat(rsTemp!指导售价 * db包装系数, mintPriceDigit)

                txtValuer.Text = IIf(IsNull(rsTemp!调价人), "", rsTemp!调价人)
                txtSummary.Text = IIf(IsNull(rsTemp!调价说明), "", rsTemp!调价说明)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!执行日期)
                End If
                If IsNull(rsTemp!执行日期) Then
                    StrToday = Format(zlDataBase.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!执行日期, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.药品id)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long

    '查找药品
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If

        For lngRow = 1 To vsfPrice.rows - 1
            lngFindRow = vsfPrice.FindRow(str药名, lngRow, CLng(menuPriceCol.品名), True, True)
            If lngFindRow > 0 Then
                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
                vsfPrice.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '查询到数据后就移动下下一条并退出本次查询
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + y <= 800 Then Exit Sub
    If TabCtlDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 5
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - y
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) >= 100 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) > 100 Then
        MsgBox "说明太长！", vbInformation, gstrSysName
        txtSummary.SelStart = 0
        txtSummary.SelLength = LenB(StrConv(txtSummary.Text, vbFromUnicode))
        Cancel = True
    End If
End Sub

Private Sub txt供应商_GotFocus()
    Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    strTmp = UCase(Trim(Me.txt供应商.Text))

    If strTmp = "" Then
        Me.txt供应商.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt供应商.Tag, "|")(1) Then
        Exit Sub
    End If

    gstrSQL = "Select 编码,名称,简码,id" & _
            " From 供应商" & _
            " where (编码 Like [1] " & _
            "       Or 名称 Like [2] " & _
            "       Or 简码 Like [2])" & _
            " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

    With rsTemp
        If .EOF Then
            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
            Me.txt供应商.Text = Split(Me.txt供应商.Tag, "|")(1)
            Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
            Exit Sub
        End If

        If .RecordCount = 1 Then
            Me.txt供应商.Text = Trim(rsTemp!名称): Me.txt供应商.Tag = rsTemp!id & "|" & rsTemp!名称
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk供应商.Left
                .Top = Me.txt供应商.Top + Me.txt供应商.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get分段加成售价(ByVal lng药品ID As Long, ByVal lng比例系数 As Long, ByVal dbl采购价 As Double, ByRef dbl售价 As Double)
'功能：通过成本价按分段加成方式计算售价
'参数：成本价,售价
    Dim dbl差价额 As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle

    mdbl分段加成率 = 0
    dbl差价额 = 0
    
    gstrSQL = "select 类别 from  收费项目目录 a where a.id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取得药品材质分类", lng药品ID)
    If rsTemp!类别 = 7 Then
        mrs分段加成.Filter = "类型=1"
    Else
        mrs分段加成.Filter = "类型=0"
    End If
    
    If mrs分段加成.RecordCount <> 0 Then
        mrs分段加成.MoveFirst
        Do While Not mrs分段加成.EOF
            With mrs分段加成
                If dbl采购价 > !最低价 And dbl采购价 <= !最高价 Then
                    mdbl分段加成率 = IIf(IsNull(!加成率), 0, !加成率) / 100
                    dbl差价额 = IIf(IsNull(!差价额), 0, !差价额)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs分段加成.MoveNext
        Loop
    End If
    
    If blnData = False Then
        MsgBox "没有设置金额段为：" & dbl采购价 & "  的分段加成数据，请在药品目录管理（分段加成率）中设置！", vbInformation, gstrSysName
        dbl售价 = 0
        Exit Sub
    End If
    
    dbl售价 = dbl采购价 * (1 + mdbl分段加成率) + dbl差价额
    
    Set rsTemp = Nothing
    gstrSQL = "Select 指导零售价 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    If rsTemp!指导零售价 * lng比例系数 < dbl售价 Then
        dbl售价 = rsTemp!指导零售价 * lng比例系数
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt供应商_Validate(Cancel As Boolean)
    If Me.txt供应商.Text = "" Then
        Me.txt供应商.Tag = "|"
    ElseIf Me.txt供应商.Text <> Split(Me.txt供应商.Tag, "|")(1) Then
        txt供应商_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub vsfPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPay
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPay
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfPay_DblClick()
    With vsfPay
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPay
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsfPay_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfPay
        If KeyCode = vbKeyReturn Then
            If .Col = menuPayCol.品名 Then
                .Col = menuPayCol.发票号
            ElseIf .Col = menuPayCol.发票号 Then
                .Col = menuPayCol.发票日期
            ElseIf .Col = menuPayCol.发票日期 Then
                .Col = menuPayCol.发票金额
            ElseIf .Col = menuPayCol.发票金额 And .Row <> .rows - 1 Then
                .Col = menuPayCol.品名
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.发票金额 Then
                strkey = .EditText
                intDigit = mintMoneyDigit
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            ElseIf Col = menuPayCol.发票号 Then
                If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String

    With vsfPay
        If Col = menuPayCol.发票日期 Then
            strkey = .EditText
            If strkey <> "" Then
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "对不起，发票日期必须为日期型,格式(20000101或者2000-01-01)！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(Row, menuPayCol.发票日期) = .EditText
                End If
                
                If Not IsDate(strkey) Then
                    MsgBox "对不起，发票日期必须为日期型(20000101或者2000-01-01)！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfprice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfPrice
        If Col = menuPriceCol.现成本价 Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.原成本价)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        ElseIf Col = menuPriceCol.现零售价 Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.原零售价)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.药品id)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '功能：行的显示与隐藏
    '参数：药品id
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.药品id)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'            .Editable = flexEDNone
'        Else
'            .Editable = flexEDKbdMouse
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double

    mBlnClick = True
    vRect = GetControlRect(vsfPrice.hWnd) '获取位置
    dblLeft = vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight


    On Error GoTo errHandle
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "", 0, , , , , , , , , True)
    End If
    Set mrsReturn = frmSelector.ShowME(Me, 0, 1, , dblLeft, dblTop, , , , , , , , , False, mstrPrivs)

    If mrsReturn.RecordCount = 0 Then Exit Sub
    mblnUpdateAdd = True
    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '用来获取药品信息
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim lngRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '是否是时价
    Dim strUnit As String
    Dim db包装系数 As Double
    Dim strInfo As String

    On Error GoTo errHandle

    mlngOldDrugID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.药品id))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!药品id

            '检查是否存在为执行的价格
            If checkNotExecutePrice(lngDrugID, strInfo) = True Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case mintUnit
                Case 0
                    db包装系数 = rsReturn!药库包装
                    strUnit = rsReturn!药库单位
                Case 1
                    db包装系数 = rsReturn!住院包装
                    strUnit = rsReturn!住院单位
                Case 2
                    db包装系数 = rsReturn!门诊包装
                    strUnit = rsReturn!门诊单位
                Case 3
                    db包装系数 = 1
                    strUnit = rsReturn!售价单位
            End Select

            .TextMatrix(Row, menuPriceCol.药品id) = lngDrugID

            If gint药品名称显示 = 1 Then
                .TextMatrix(Row, menuPriceCol.品名) = "[" & rsReturn!药品编码 & "]" & IIf(IsNull(rsReturn!商品名), rsReturn!通用名, rsReturn!商品名)
            Else
                .TextMatrix(Row, menuPriceCol.品名) = "[" & rsReturn!药品编码 & "]" & rsReturn!通用名
            End If

            .TextMatrix(Row, menuPriceCol.规格) = IIf(IsNull(rsReturn!规格), "", rsReturn!规格)
            .TextMatrix(Row, menuPriceCol.是否变价) = rsReturn!时价
            intCurrentPrice = rsReturn!时价
            .TextMatrix(Row, menuPriceCol.厂牌) = IIf(IsNull(rsReturn!产地), "", rsReturn!产地)
            .TextMatrix(Row, menuPriceCol.单位) = strUnit
            .TextMatrix(Row, menuPriceCol.包装系数) = db包装系数
            gstrSQL = "select 药品id from 药品库存 where 药品id=[1] and 性质=1 "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查库存", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.是否有库存) = 0
            Else
                .TextMatrix(Row, menuPriceCol.是否有库存) = 1
            End If

            If intCurrentPrice = 0 Then '定价药品
                '表示定价药品调价，成本价取平均价格，售价取收费价目现价
                gstrSQL = "Select b.Id, Decode(Nvl(k.库存数量, 0), 0, a.成本价, (k.库存金额 - k.库存差价) / k.库存数量) As 成本价, a.指导批发价, a.指导零售价, b.现价, a.差价让利比," & vbNewLine & _
                            "       nvl(a.加成率,0) / 100 As 加成率, b.收入项目id" & vbNewLine & _
                            "From 药品规格 A, 收费价目 B," & vbNewLine & _
                            "     (Select Sum(实际金额) 库存金额, Sum(实际差价) As 库存差价, Sum(实际数量) 库存数量" & vbNewLine & _
                            "       From 药品库存" & vbNewLine & _
                            "       Where 性质 = 1 And 药品id = [1] ) K" & vbNewLine & _
                            "Where a.药品id = b.收费细目id And a.药品id = [1] And Sysdate Between 执行日期 And 终止日期"
            Else '时价药品
                '表示时价药品调价，取库存金额/库存数量做为其价格
                gstrSQL = "select P.id,Decode(Nvl(K.库存数量,0),0,P.现价,K.库存金额/Nvl(K.库存数量,1)) 现价,nvl(j.加成率,0) / 100 as 加成率,decode(nvl(k.库存数量,0),0,j.成本价,(k.库存金额-k.库存差价)/k.库存数量) as 成本价,j.指导批发价,j.指导零售价,j.差价让利比,p.收入项目id,P.执行日期,P.收入项目id,I.名称 as 收入名称" & _
                        " from 收费价目 P,收入项目 I,药品规格 J," & _
                        "   (Select Sum(实际金额) 库存金额,Sum(实际差价) as 库存差价,Sum(实际数量) 库存数量" & _
                        "    From 药品库存 Where 性质=1 and 药品ID=[1] ) K" & _
                        " where P.收入项目id=I.id and p.收费细目id=j.药品id and P.收费细目id=[1] " & _
                        "       and (P.终止日期 is null or SYSDATE BETWEEN P.执行日期 AND P.终止日期)"
            End If
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "查询药品", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                MsgBox "该药品不存在，请重新建立该药品卡片！", vbInformation, gstrSysName
                Exit Sub
            End If
            .TextMatrix(Row, menuPriceCol.原价id) = rsTemp!id
            .TextMatrix(Row, menuPriceCol.收入项目ID) = IIf(IsNull(rsTemp!收入项目ID), 0, rsTemp!收入项目ID)
            .TextMatrix(Row, menuPriceCol.加成率) = GetFormat(IIf(IsNull(rsTemp!加成率), 0, rsTemp!加成率), 5)
            .TextMatrix(Row, menuPriceCol.差价让利比) = IIf(IsNull(rsTemp!差价让利比), 100, rsTemp!差价让利比)
            .TextMatrix(Row, menuPriceCol.原成本价) = GetFormat(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * db包装系数, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.现成本价) = GetFormat(IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * db包装系数, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.原零售价) = GetFormat(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * db包装系数, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.现零售价) = GetFormat(IIf(IsNull(rsTemp!现价), 0, rsTemp!现价) * db包装系数, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.原采购限价) = GetFormat(IIf(IsNull(rsTemp!指导批发价), 0, rsTemp!指导批发价) * db包装系数, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.现采购限价) = .TextMatrix(Row, menuPriceCol.原采购限价)
            .TextMatrix(Row, menuPriceCol.原指导售价) = GetFormat(IIf(IsNull(rsTemp!指导零售价), 0, rsTemp!指导零售价) * db包装系数, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.现指导售价) = .TextMatrix(Row, menuPriceCol.原指导售价)

            Call GetDrugStore(lngDrugID, Row)
            If Row = .rows - 1 Then '最后一行才新增行
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
                Row = Row + 1
            End If
        End With
'        If mint调价 = 0 And mbln时价药品按批次调价 = True Then '售价调价
'            Call GetDrugStore(lngDrugID, db包装系数)
'        ElseIf mint调价 <> 0 Then

'        End If
'        Call SetRowHidden(lngDrugID)

        rsReturn.MoveNext
    Next
    Call setColEdit

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl加成率 As Double
    Dim lngCurRow As Long     '当前行
    Dim i As Long
    Dim dbl发票金额 As Double
    Dim str药品名称 As String
    Dim str发票 As String
    Dim str发票日期 As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl包装换算 As Double
    Dim bln相同药品 As Boolean
    Dim lng药品ID As Long
    Dim str单位 As String


    '功能：为库存列表填充数据
    '参数：药品id

    On Error GoTo errHandle
    '先检查是否有重复的数据，如果有就先清除掉重复的数据
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.药品id)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.库房id,s.药品id, d.名称 As 库房, '[' || m.编码 || ']' || m.名称 As 药品, m.规格, m.产地, m.计算单位 售价单位," & vbNewLine & _
                        "       p.药库单位, s.上次批号 As 批号, nvl(s.实际数量,0) As 数量,s.批次, Nvl(m.是否变价, 0) 变价, m.Id," & vbNewLine & _
                        "       Decode(Nvl(m.是否变价, 0), 0, e.现价, Decode(Nvl(s.零售价, 0),0,Decode(Nvl(s.实际数量, 0),0,e.现价, s.实际金额/s.实际数量),s.零售价)) As 时价售价," & vbNewLine & _
                        "       p.指导差价率 As 差价率,nvl(p.加成率,0) as 加成率 ,Decode(Nvl(s.平均成本价, 0), 0, p.成本价, s.平均成本价) As 成本价, s.上次供应商id, n.名称 As 供应商, s.效期, s.上次产地 As 产地" & vbNewLine & _
                        "From 药品库存 S, 部门表 D, 收费项目目录 M, 药品规格 P, 供应商 N, 收费价目 E" & vbNewLine & _
                        "Where d.Id = s.库房id And s.药品id = m.Id And m.Id = p.药品id And Nvl(s.上次供应商id, 0) = n.Id(+) And m.Id = e.收费细目id And" & vbNewLine & _
                        "      s.性质 = 1 And s.药品id = [1] And Sysdate Between e.执行日期 And e.终止日期 " & vbNewLine & _
                        "Order By 库房, s.上次批号"

        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng供应商ID > 0 Then
            rsTemp.Filter = "上次供应商ID=" & mlng供应商ID
        End If
    Else '修改，查阅
        If mintModal = 2 Then   '查阅
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.药品id, b.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & vbNewLine & _
                            "                b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额, b.产地, b.批次, b.批号, e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位," & vbNewLine & _
                            "                nvl(a.填写数量,0) As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率 ,b.效期" & vbNewLine & _
                            "From 药品收发记录 A,成本价调价信息 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & vbNewLine & _
                            "Where a.id=b.收发id And a.库房id = c.Id And b.供药单位id = d.Id(+) And" & vbNewLine & _
                            "      a.药品id = e.Id And e.Id = f.药品id And b.调价汇总号 = [1] and a.单据 = 5"
                Else
                    gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.药库单位,nvl(a.实际数量,0) as 数量,f.指导差价率 as 差价率,nvl(f.加成率,0) as 加成率 ,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,药品规格 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And Decode(b.库房id, Null, 1, a.库房id) = Decode(b.库房id, Null, 1, b.库房id) " & _
                            " and Decode(b.库房id, Null, 1, Nvl(a.批次, 0)) = Decode(b.库房id, Null, 1, Nvl(b.批次, 0)) " & _
                            " and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.药品id and a.性质=1 "
                End If
            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "select (sysdate-执行日期 ) as 是否执行 from 调价汇总记录 where 调价号=[1]"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "是否执行", txtNO.Text)
                If rsTemp!是否执行 > 0 Then
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.供药单位id As 上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格," & vbNewLine & _
                            "                d.名称 As 供应商, f.成本价 As 新成本价, f.成本价 As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.产地, a.批次, a.批号, e.是否变价 As 变价," & vbNewLine & _
                            "                e.计算单位 As 售价单位, f.药库单位, nvl(a.填写数量,0) As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率 ,a.效期" & vbNewLine & _
                            "From 药品收发记录 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & vbNewLine & _
                            "Where a.价格id = b.Id And a.库房id = c.Id And a.供药单位id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And" & vbNewLine & _
                            "      b.调价汇总号 = [1] and a.单据=13"
                Else
                    gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率 ,a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1  And" & _
                                  " b.调价汇总号 = [1]"
                End If
            End If
        Else '修改
            If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                gstrSQL = "Select Distinct a.库房id,c.名称 as 库房, b.药品id,a.上次供应商id, '[' || e.编码 || ']' ||e.名称 as 药品,e.规格,d.名称 as 供应商, b.新成本价, b.原成本价, b.发票号, b.发票日期, b.发票金额" & _
                            " ,a.上次产地 as 产地,a.批次,a.上次批号 as 批号,e.是否变价 as 变价,e.计算单位 as 售价单位,f.药库单位,nvl(a.实际数量,0) as 数量,f.指导差价率 as 差价率,nvl(f.加成率,0) as 加成率 ,a.效期" & _
                            " From 药品库存 A,部门表 C,供应商 D,收费项目目录 E,药品规格 F," & _
                                 " (Select Distinct 药品id, 库房id, 批次, 批号, 效期, 产地, 原成本价, 新成本价, 发票号, 发票日期, 发票金额, 应付款变动, 执行日期" & _
                                   " From 成本价调价信息" & _
                                   " Where 调价汇总号 = [1]) B" & _
                            " Where a.药品id = b.药品id And Decode(b.库房id, Null, 1, a.库房id) = Decode(b.库房id, Null, 1, b.库房id) " & _
                            " and Decode(b.库房id, Null, 1, Nvl(a.批次, 0)) = Decode(b.库房id, Null, 1, Nvl(b.批次, 0)) " & _
                            " and a.库房id=c.id and a.上次供应商id=d.id(+) and a.药品id=e.id and e.id=f.药品id and a.性质=1  "
            ElseIf cboPriceMethod.Text = "仅调售价" Then
                gstrSQL = "Select Distinct a.库房id, c.名称 As 库房, b.收费细目id As 药品id, a.上次供应商id, '[' || e.编码 || ']' || e.名称 As 药品, e.规格, d.名称 As 供应商," & _
                                            " nvl(a.平均成本价,f.成本价) As 新成本价, nvl(a.平均成本价,f.成本价) As 原成本价, '' 发票号, '' 发票日期, '' 发票金额, a.上次产地 As 产地, a.批次, a.上次批号 As 批号," & _
                                            " e.是否变价 As 变价, e.计算单位 As 售价单位, f.药库单位, nvl(a.实际数量,0) As 数量, f.指导差价率 As 差价率, nvl(f.加成率,0) as 加成率 ,a.效期" & _
                            " From 药品库存 A, 收费价目 B, 部门表 C, 供应商 D, 收费项目目录 E, 药品规格 F" & _
                            " Where a.药品id = b.收费细目id And a.库房id = c.Id And a.上次供应商id = d.Id(+) And a.药品id = e.Id And e.Id = f.药品id And a.性质 = 1  " & _
                                  "And b.调价汇总号 = [1]"
            End If
        End If
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl包装换算 = 0
            dbl发票金额 = 0
            dblOldPrice = 0
            dblNewPrice = 0
            str单位 = ""
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!药品id = vsfPrice.TextMatrix(i, menuPriceCol.药品id) Then
                    dbl包装换算 = vsfPrice.TextMatrix(i, menuPriceCol.包装系数)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.原零售价))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.现零售价))
                    str单位 = vsfPrice.TextMatrix(i, menuPriceCol.单位)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '从空白行开始插入数据
            .TextMatrix(.rows - 1, menuStoreCol.药品id) = rsTemp!药品id
            .TextMatrix(.rows - 1, menuStoreCol.库房) = rsTemp!库房
            .TextMatrix(.rows - 1, menuStoreCol.库房id) = rsTemp!库房id
            .TextMatrix(.rows - 1, menuStoreCol.供应商) = Nvl(rsTemp!供应商, "")
            .TextMatrix(.rows - 1, menuStoreCol.供应商id) = IIf(mlng供应商ID > 0, mlng供应商ID, Nvl(rsTemp!上次供应商ID))
            .TextMatrix(.rows - 1, menuStoreCol.药品) = rsTemp!药品
            str药品名称 = rsTemp!药品

            .TextMatrix(.rows - 1, menuStoreCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            .TextMatrix(.rows - 1, menuStoreCol.单位) = str单位
            .TextMatrix(.rows - 1, menuStoreCol.批号) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
            .TextMatrix(.rows - 1, menuStoreCol.效期) = Format(IIf(IsNull(rsTemp!效期), "", rsTemp!效期), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            .TextMatrix(.rows - 1, menuStoreCol.数量) = GetFormat(rsTemp!数量 / dbl包装换算, mintNumberDigit)
            .TextMatrix(.rows - 1, menuStoreCol.包装系数) = dbl包装换算
            .TextMatrix(.rows - 1, menuStoreCol.批次) = Nvl(rsTemp!批次, 0)
            .TextMatrix(.rows - 1, menuStoreCol.变价) = rsTemp!变价


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!成本价), 0, rsTemp!成本价) * dbl包装换算
                
                If mdbl加成率 > 0 Then
                    dbl加成率 = Round(mdbl加成率 / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl加成率 = Round(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl加成率 = Nvl(rsTemp!加成率, 0) / 100
                End If
                If 1 + dbl加成率 = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!时价售价 * dbl包装换算 / (1 + dbl加成率)
                End If
                If dbl加成率 = -1 Then dbl加成率 = 0

                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit)
                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(IIf(rsTemp!变价 = 1, rsTemp!时价售价 * dbl包装换算, dblOldPrice), mintPriceDigit)
                .TextMatrix(.rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.加成率) = GetFormat(GetFormat(dbl加成率, 5) * 100, 5)
                .TextMatrix(.rows - 1, menuStoreCol.原采购价) = GetFormat(dblOldCost, mintCostDigit)
                .TextMatrix(.rows - 1, menuStoreCol.现采购价) = GetFormat(dblNewCost, mintCostDigit)
                .TextMatrix(.rows - 1, menuStoreCol.差价差) = Format((Val(.TextMatrix(.rows - 1, menuStoreCol.现采购价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原采购价))) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                dbl发票金额 = dbl发票金额 + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                
                '为应付记录表赋值
                If mint调价 = 1 Or mint调价 = 2 Then
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = GetFormat(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.品名) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.品名) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                    End If
                End If
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "仅调售价" Or cboPriceMethod.Text = "售价成本价一起调价") Then   '查阅
                    gstrSQL = "Select a.成本价 As 原价, a.零售价 As 现价" & vbNewLine & _
                        "From 药品收发记录 A, 收费价目 B" & vbNewLine & _
                        "Where a.价格id = b.Id And b.调价汇总号 = [1] And a.库房id = [2] And a.药品id = [3] And Nvl(a.批次, 0) = [4]"
                    Set rsPirce = zlDataBase.OpenSQLRecord(gstrSQL, "获取售价", txtNO.Text, rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(Val(rsPirce!原价) * dbl包装换算, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(Val(rsPirce!现价) * dbl包装换算, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(dblOldPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(dblNewPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    End If
                    If cboPriceMethod.Text = "仅调售价" Then
                        gstrSQL = "Select 成本价" & vbNewLine & _
                                    "      From (Select 平均成本价 As 成本价" & vbNewLine & _
                                    "             From 药品库存" & vbNewLine & _
                                    "             Where 性质=1 And 库房id = [1] And 药品id = [2] And nvl(批次,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select 成本价 From 药品规格 Where 药品id = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDataBase.OpenSQLRecord(gstrSQL, "获取成本价", rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                        .TextMatrix(.rows - 1, menuStoreCol.原采购价) = GetFormat(rsCost!成本价 * dbl包装换算, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.现采购价) = GetFormat(rsCost!成本价 * dbl包装换算, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.差价差) = Format(0, mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原采购价) = GetFormat(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.现采购价) = GetFormat(rsTemp!新成本价 * dbl包装换算, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.差价差) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                    End If
                Else '修改或者成本价调价
                    '定价直接从收费价目取现价，时价优先从库存取，如果没有则从收费价目取
                    If Nvl(rsTemp!变价, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.零售价, Decode(Nvl(s.实际数量, 0), 0, 0, Nvl(s.实际金额, 0) / s.实际数量)) 时价售价" & vbNewLine & _
                        "From 药品库存 S" & vbNewLine & _
                        "Where s.性质=1 And s.库房id = [1] And s.药品id = [2] And nvl(s.批次,0) = [3]"
                        
                        Set rsPirce = zlDataBase.OpenSQLRecord(gstrSQL, "获取售价", rsTemp!库房id, rsTemp!药品id, Nvl(rsTemp!批次, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!时价售价 > 0 Then
                                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(rsPirce!时价售价 * dbl包装换算, mintPriceDigit)
                                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(rsPirce!时价售价 * dbl包装换算, mintPriceDigit)
                            Else
                                .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(dblOldPrice, mintPriceDigit)
                                .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(dblNewPrice, mintPriceDigit)
                            End If
                        Else
                            .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(dblOldPrice, mintPriceDigit)
                            .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(dblNewPrice, mintPriceDigit)
                        End If
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.原零售价) = GetFormat(dblOldPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.现零售价) = GetFormat(dblNewPrice, mintPriceDigit)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.调整金额) = Format(rsTemp!数量 / dbl包装换算 * (Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) - Val(.TextMatrix(.rows - 1, menuStoreCol.原零售价))), mstrMoneyFormat)
                    .TextMatrix(.rows - 1, menuStoreCol.原采购价) = GetFormat(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.现采购价) = GetFormat(rsTemp!新成本价 * dbl包装换算, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.差价差) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                End If
                 
                If cboPriceMethod.Text = "仅调成本价" Or cboPriceMethod.Text = "售价成本价一起调价" Then
                    If rsTemp!新成本价 = 0 Then
                        dbl加成率 = 0
                    Else
                        dbl加成率 = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.现零售价)) / (rsTemp!新成本价 * dbl包装换算) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.加成率) = GetFormat(GetFormat(dbl加成率, 5) * 100, 5)
                    .TextMatrix(.rows - 1, menuStoreCol.原采购价) = GetFormat(Nvl(rsTemp!原成本价, 0) * dbl包装换算, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.现采购价) = GetFormat(rsTemp!新成本价 * dbl包装换算, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.差价差) = Format((rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量)), mstrMoneyFormat)
                    dbl发票金额 = dbl发票金额 + (rsTemp!新成本价 * dbl包装换算 - Nvl(rsTemp!原成本价, 0) * dbl包装换算) * Val(.TextMatrix(.rows - 1, menuStoreCol.数量))
                    str发票 = IIf(IsNull(rsTemp!发票号), "", rsTemp!发票号)
                    str发票日期 = IIf(IsNull(rsTemp!发票日期), "", rsTemp!发票日期)
                    
                    '为付款记录列表赋值
                    If vsfPay.rows > 1 Then
                        bln相同药品 = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.药品id) = rsTemp!药品id Then
                                bln相同药品 = True
                                Exit For
                            End If
                        Next
                        If bln相同药品 = True Then
                            vsfPay.TextMatrix(i, menuPayCol.发票金额) = GetFormat(Val(vsfPay.TextMatrix(i, menuPayCol.发票金额)) + dbl发票金额, mintMoneyDigit)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.品名) = str药品名称
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.药品id) = rsTemp!药品id
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.品名) = str药品名称
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票号) = str发票
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票日期) = Format(str发票日期, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    '修改和查阅时重算规格列表平均成本价，售价
    'mintModal 0-新增 1-修改 2-查阅
'    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lng药品ID <> .TextMatrix(i, menuStoreCol.药品id) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.药品id)))
                    lng药品ID = Val(.TextMatrix(i, menuStoreCol.药品id))
                End If
            Next
        End With
'    End If

    If mint调价 = 1 Or mint调价 = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If

        If .Col = menuPriceCol.现零售价 Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.现零售价))
        ElseIf .Col = menuPriceCol.现成本价 Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.现成本价))
        End If
    End With
    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id)) = 0 Then Exit Sub

        If .rows > 1 Then
            For i = 1 To .rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id)) = Val(.TextMatrix(i, menuStoreCol.药品id)) Then
                    .Select i, 0, i, .Cols - 1
                    .TopRow = i
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String

    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.现零售价 Then '成本价调价
                If .Col = menuPriceCol.品名 And cboPriceMethod.Text = "仅调成本价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.品名 And cboPriceMethod.Text = "仅调售价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "仅调成本价" Then
                    If .Row = .rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .ColComboList(menuPriceCol.品名) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                    End If
                ElseIf .Col = menuPriceCol.品名 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现成本价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现成本价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    .Col = menuPriceCol.现零售价
'                    .EditCell
                ElseIf .Col = menuPriceCol.现零售价 And cboPriceMethod.Text = "售价成本价一起调价" Then
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                        .ColComboList(menuPriceCol.品名) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.品名
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 And .Row = .rows - 1 Then
                    .ColComboList(menuPriceCol.品名) = ""
                    .rows = .rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.品名
                    .RowHeight(.rows - 1) = mlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.药品id)) <> 0 Then
                    .ColComboList(menuPriceCol.品名) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.品名
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.药品id))
            If lngDrugID = 0 Then Exit Sub
            If MsgBox("是否继续删除这个药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '修改模式时删除一条价格表中数据，则清楚未执行价格
            If mintModal = 1 Then
                gstrSQL = "Zl_删除未执行价格_Delete(" & lngDrugID & "," & 0 & ")"
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = gstrSQL
            End If
            
            If .rows > 2 Then
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.药品id)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim mrsReturn As Recordset
    Dim rsTemp As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strkey As String
    Dim lngDrugID As Long
    Dim intCurrentPirce As Integer '是否是时价

    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    mBlnClick = True
    vRect = GetControlRect(vsfPrice.hWnd) '获取位置
    dblLeft = vRect.Left + vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight

    With vsfPrice
        strkey = .EditText
        Select Case Col
        Case menuPriceCol.品名
            If grsMaster.State = adStateClosed Then
                Call SetSelectorRS(1, "", 0, , , , , , , , , True)
            End If
            Set mrsReturn = frmSelector.ShowME(Me, 1, 1, strkey, dblLeft, dblTop, , , , , , , , , False, mstrPrivs)
            If mrsReturn.RecordCount = 0 Then Exit Sub
            mblnUpdateAdd = True
            Call GetDrugPirce(mrsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '检查是否有重复的药品
    'lngDrugId 药品id
    '返回值 true-存在重复值 false-不存在重复值
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strName As String
    Dim intCount As Integer
    Dim intLength As Integer

    If rsTemp.RecordCount = 0 Then Exit Function
    rsTemp.MoveFirst
    With vsfPrice
        For i = 0 To rsTemp.RecordCount - 1
            For j = 1 To .rows - 1
                If Val(.TextMatrix(j, menuPriceCol.药品id)) = rsTemp!药品id Then
                    strTemp = strTemp & " 药品id <> " & rsTemp!药品id & " and "
                    intCount = intCount + 1
                    If intCount < 5 Then
                        strName = strName & rsTemp!通用名 & " "
                    End If
                End If
            Next
            rsTemp.MoveNext
        Next
    End With

    If strTemp <> "" Then
        intLength = LenB(StrConv(strTemp, vbFromUnicode)) '得到字符串长度
        Do Until Mid(strTemp, intLength, 3) = "and" '从后向前查找倒数第一个"and"
           intLength = intLength - 1
        Loop
        strTemp = Left(strTemp, intLength - 1) '倒数第一个"and"之前的字符串

        rsTemp.Filter = strTemp
        MsgBox strName & "等" & intCount & "种药品在列表中已经存在，已存在药品不再添加！", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.品名 Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If .Col = menuPriceCol.现成本价 Then
            mdbl成本价 = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.现成本价 Or Col = menuPriceCol.现零售价 Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.现成本价
                    intDigit = mintCostDigit
                Case menuPriceCol.现零售价
                    intDigit = mintPriceDigit
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strkey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.品名 Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.品名 Then
        vsfPrice.ColComboList(menuPriceCol.品名) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '功能：设置列是否可以修改
    '不能修改的列颜色为灰色，能修改的列颜色为白色
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .rows - 1, menuPriceCol.品名) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .rows - 1, menuPriceCol.现零售价) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .rows - 1, menuPriceCol.品名) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .rows - 1, menuPriceCol.现成本价) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.品名, .rows - 1, menuPriceCol.品名) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现成本价, .rows - 1, menuPriceCol.现成本价) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.现零售价, .rows - 1, menuPriceCol.现零售价) = mconlngCanColColor
        End If

    End With

    With vsfStore
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "仅调售价" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.现零售价, .rows - 1, menuStoreCol.现零售价) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
'            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .rows - 1, menuStoreCol.加成率) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现采购价, .rows - 1, menuStoreCol.现采购价) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.加成率, .rows - 1, menuStoreCol.加成率) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现采购价, .rows - 1, menuStoreCol.现采购价) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.现零售价, .rows - 1, menuStoreCol.现零售价) = mconlngCanColColor
        End If
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.现零售价, intRow, menuStoreCol.现零售价) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票号, .rows - 1, menuPayCol.发票号) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票日期, .rows - 1, menuPayCol.发票日期) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.发票金额, .rows - 1, menuPayCol.发票金额) = mconlngCanColColor
    End With
End Sub


Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.品名 And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.品名) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim dblSalePrice As Double
    Dim intRow As Integer
    Dim dbl加成率 As Double

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.药品id))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.现成本价
                If Val(.EditText) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "成本价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .EditText = GetFormat(.EditText, mintPriceDigit)
                If mbln现价提示 = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原采购限价)) Then
                        If MsgBox("现成本价高于采购价限价" & Val(.TextMatrix(.Row, menuPriceCol.原采购限价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                .TextMatrix(.Row, menuPriceCol.现采购限价) = GetFormat(.EditText, mintCostDigit)

                If cbo售价计算方式.Text = "售价按分段加成计算" And .TextMatrix(.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                    Call get分段加成售价(lngDrugID, Val(.TextMatrix(.Row, menuPriceCol.包装系数)), Val(.EditText), dblSalePrice)
                    If dblSalePrice = 0 Then
                        .EditText = mdbl成本价
                        .TextMatrix(vsfPrice.Row, menuPriceCol.现成本价) = GetFormat(.EditText, mintCostDigit)
                        Exit Sub
                    End If
                    dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, menuPriceCol.差价让利比)) / 100)
                    .TextMatrix(.Row, menuPriceCol.现零售价) = GetFormat(dblSalePrice, mintPriceDigit)
                    
                    '调了售价应该同步更新库存列表价格信息
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现零售价) = GetFormat(dblSalePrice, mintPriceDigit)
                                vsfStore.TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) <> 0 Then
                                    dbl加成率 = GetFormat(GetFormat(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) - 1), 5) * 100, 5)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = dbl加成率
                            End If
                        Next
                    End If
                ElseIf cbo售价计算方式 = "售价按固定比例计算" And .TextMatrix(.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                    dblSalePrice = Val(.EditText) * (1 + Val(.TextMatrix(.Row, menuPriceCol.加成率)))
                    If dblSalePrice > Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) Then dblSalePrice = Val(.TextMatrix(.Row, menuPriceCol.原指导售价))
                    .TextMatrix(.Row, menuPriceCol.现零售价) = GetFormat(dblSalePrice, mintPriceDigit)
                    
                    '调了售价应该同步更新库存列表价格信息
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现零售价) = GetFormat(dblSalePrice, mintPriceDigit)
                                vsfStore.TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) <> 0 Then
                                    dbl加成率 = GetFormat(GetFormat(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现零售价))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) - 1), 5) * 100, 5)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = dbl加成率
                            End If
                        Next
                    End If
                End If

                Call CaculateCost(lngDrugID, .EditText) '重新计算成本价
            Case menuPriceCol.现零售价
                If Val(.EditText) < 0 Then
                    MsgBox "售价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If

                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

                .EditText = GetFormat(.EditText, mintPriceDigit)
'                If mdblOldPrice = .EditText Then '未做修改直接退出
'                    Exit Sub
'                End If

                If mbln现价提示 = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.原指导售价)) Then
                        If MsgBox("现零售价高于指导售价" & Val(.TextMatrix(.Row, menuPriceCol.原指导售价)) & "。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                .TextMatrix(.Row, menuPriceCol.现指导售价) = GetFormat(.EditText, mintPriceDigit)
                If chkAotuCost.Value = 1 Then '修改售价后自动计算成本价
                    .TextMatrix(.Row, menuPriceCol.现成本价) = GetFormat(.EditText / (1 + Val(.TextMatrix(.Row, menuPriceCol.加成率))), mintCostDigit)
                    
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(.Row, menuPriceCol.药品id) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.现采购价) = GetFormat(.TextMatrix(.Row, menuPriceCol.现成本价), mintCostDigit)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) <> 0 Then
                                    dbl加成率 = GetFormat((.EditText / Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) - 1), 5)
                                Else
                                    dbl加成率 = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.加成率) = GetFormat(dbl加成率 * 100, 5)
                                vsfStore.TextMatrix(intRow, menuStoreCol.差价差) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原采购价))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                            End If
                        Next
                    End If
                End If

                Call ChangeDrugStore(Row, lngDrugID, .EditText)
        End Select
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugID As Long, ByVal dblNewPrice As Double)
    '功能：通过修改价格表中的零售价修改库存列表中相对应的零售价
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl包装 As Double
    Dim n As Integer
    Dim dbl发票金额 As Double
    Dim dbl加成率 As Double

    If intRow = 0 Or mint调价 = 1 Then Exit Sub

    dbl包装 = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.包装系数))

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.药品id)) = lngDrugID Then
                    dblNum = Val(.TextMatrix(n, menuStoreCol.数量))
                    dblOldPrice = Val(vsfStore.TextMatrix(n, menuStoreCol.原零售价))

                    .TextMatrix(n, menuStoreCol.现零售价) = GetFormat(dblNewPrice, mintPriceDigit)
                    .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    
                    If Val(.TextMatrix(n, menuStoreCol.现采购价)) <> 0 Then
                        dbl加成率 = GetFormat(((Val(.TextMatrix(n, menuStoreCol.现零售价))) / Val(.TextMatrix(n, menuStoreCol.现采购价)) - 1), 5)
                    Else
                        dbl加成率 = 0
                    End If
                    .TextMatrix(n, menuStoreCol.加成率) = GetFormat(dbl加成率 * 100, 5)
                
                    If mint调价 = 2 And chkAotuCost.Value = 1 Then
                        dblOldCost = .TextMatrix(n, menuStoreCol.原采购价)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.加成率)) / 100, 7))
                        .TextMatrix(n, menuStoreCol.现采购价) = GetFormat(dblNewCost, mintCostDigit)
                        .TextMatrix(n, menuStoreCol.差价差) = Format((.TextMatrix(n, menuStoreCol.现采购价) - dblOldCost) * dblNum, mstrMoneyFormat)
                    End If
                    dbl发票金额 = dbl发票金额 + Val(.TextMatrix(n, menuStoreCol.差价差))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        With vsfPay
            For n = 1 To .rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, menuPayCol.药品id)) = lngDrugID Then
                        .TextMatrix(n, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                    End If
                End If
            Next
        End With
    End If

    If mint调价 = 2 Then
        CaluateAverCost lngDrugID
    End If
End Sub

Private Sub CaluateAverCost(ByVal lng药品ID As Long)
    '计算平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.现采购价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品ID Then
                        .TextMatrix(i, menuPriceCol.现成本价) = GetFormat(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverOldCost(ByVal lng药品ID As Long)
    '计算原始平均成本价
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品ID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.原采购价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品ID Then
                        .TextMatrix(i, menuPriceCol.原成本价) = GetFormat(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateCost(ByVal lng药品ID As Long, ByVal dbl现成本价 As Double)
    '功能：通过修改价格表中的成本价修改库存列表中相对应的成本价

    Dim n As Integer
    Dim dbl发票金额 As Double

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.药品id)) = lng药品ID Then
                    .TextMatrix(n, menuStoreCol.现采购价) = GetFormat(dbl现成本价, mintCostDigit)
                    If (cbo售价计算方式.Text = "售价按分段加成计算" Or cbo售价计算方式.Text = "售价按固定比例计算") And vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.是否变价) = "1" And mint调价 = 2 Then
                        .TextMatrix(n, menuStoreCol.现零售价) = vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.现零售价)
                    End If
                    If dbl现成本价 <> 0 Then
                        .TextMatrix(n, menuStoreCol.加成率) = GetFormat(GetFormat((Val(.TextMatrix(n, menuStoreCol.现零售价)) / dbl现成本价 - 1), 5) * 100, 5)
                    End If
                    If cbo售价计算方式 = "售价按分段加成计算" Then
                        .TextMatrix(n, menuStoreCol.加成率) = GetFormat(GetFormat(mdbl分段加成率, 5) * 100, 5)
                    End If
                    .TextMatrix(n, menuStoreCol.差价差) = Format((dbl现成本价 - Val(.TextMatrix(n, menuStoreCol.原采购价))) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat)

                    dbl发票金额 = dbl发票金额 + (dbl现成本价 - .TextMatrix(n, menuStoreCol.原采购价)) * Val(.TextMatrix(n, menuStoreCol.数量))
                    .TextMatrix(n, menuStoreCol.调整金额) = (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))) * Val(.TextMatrix(n, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        For n = 1 To vsfPay.rows - 1
            If vsfPay.TextMatrix(1, 0) <> "" Then
                If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = lng药品ID Then
                    vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub


Private Sub vsfStore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfStore
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfStore
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub setColHiddenVsf()
    '不同模式下面，列显示不一样
    With vsfStore
        If cboPriceMethod.Text = "仅调售价" Then
            .ColHidden(menuStoreCol.批次) = True
            .ColHidden(menuStoreCol.变价) = True
            .ColHidden(menuStoreCol.加成率) = True
            .ColHidden(menuStoreCol.原采购价) = True
            .ColHidden(menuStoreCol.现采购价) = True
            .ColHidden(menuStoreCol.差价差) = True
            .ColHidden(menuStoreCol.原零售价) = False
            .ColHidden(menuStoreCol.现零售价) = False
        ElseIf cboPriceMethod.Text = "仅调成本价" Then
            .ColHidden(menuStoreCol.原零售价) = True
            .ColHidden(menuStoreCol.现零售价) = True
            .ColHidden(menuStoreCol.调整金额) = True
            .ColHidden(menuStoreCol.加成率) = False
            .ColHidden(menuStoreCol.原采购价) = False
            .ColHidden(menuStoreCol.现采购价) = False
            .ColHidden(menuStoreCol.差价差) = False
        ElseIf cboPriceMethod.Text = "售价成本价一起调价" Then
            .ColHidden(menuStoreCol.原零售价) = False
            .ColHidden(menuStoreCol.现零售价) = False
            .ColHidden(menuStoreCol.调整金额) = False
            .ColHidden(menuStoreCol.加成率) = False
            .ColHidden(menuStoreCol.原采购价) = False
            .ColHidden(menuStoreCol.现采购价) = False
            .ColHidden(menuStoreCol.差价差) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.药品id)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.药品id)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfStore_EnterCell()
    With vsfStore
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuStoreCol.加成率 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.加成率))
        ElseIf .Col = menuStoreCol.现采购价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.现采购价))
        ElseIf .Col = menuStoreCol.现零售价 Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.现零售价))
        End If
    End With
End Sub

Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfStore
        If KeyCode = vbKeyReturn Then
            If .Col < vsfStore.Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = menuStoreCol.规格
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfStore_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfStore
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.现采购价 Or Col = menuStoreCol.现零售价 Or Col = menuStoreCol.加成率 Then
                strkey = .EditText
                Select Case Col
                    Case menuStoreCol.现采购价
                        intDigit = mintCostDigit
                    Case menuStoreCol.现零售价
                        intDigit = mintPriceDigit
                    Case menuStoreCol.加成率
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfStore_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl发票金额 As Double
    Dim Dbl数量 As Double
    Dim Dbl金额 As Double
    Dim dbl现采购价 As Double
    Dim dblTempNum As Double
    Dim dbl成本金额 As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.现零售价
                If Not IsNumeric(.EditText) Then
                    MsgBox "请输入新的售价。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .EditText = GetFormat(.EditText, mintPriceDigit)
                End If

                If .EditText > 9999999 Then
                    MsgBox "零售价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If chkAotuCost.Value = 1 Then '修改售价后自动计算成本价
                    .TextMatrix(intRow, menuStoreCol.现采购价) = GetFormat(.EditText / (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit)
                    .TextMatrix(intRow, menuStoreCol.差价差) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.原采购价))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                End If
                
                .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.现零售价) = GetFormat(Val(.EditText), mintPriceDigit)
'                .TextMatrix(intRow, menuStoreCol.现采购价) = GetFormat(Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit)
'                .TextMatrix(intRow, menuStoreCol.差价差) = Format((Val(.TextMatrix(intRow, menuStoreCol.现采购价)) - Val(.TextMatrix(intRow, menuStoreCol.原采购价))) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                If chkAotuCost.Value <> 1 Then
                    If Val(.TextMatrix(intRow, menuStoreCol.现采购价)) <> 0 Then
                        .TextMatrix(intRow, menuStoreCol.加成率) = GetFormat(GetFormat((Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / Val(.TextMatrix(intRow, menuStoreCol.现采购价)) - 1), 5) * 100, 5)
                    Else
                        .TextMatrix(intRow, menuStoreCol.加成率) = GetFormat(0, 5)
                    End If
                End If
                
                For n = 1 To .rows - 1
                    If .TextMatrix(intRow, menuStoreCol.药品id) = .TextMatrix(n, menuStoreCol.药品id) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次)) Then
                            .TextMatrix(n, menuStoreCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                            .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                            If chkAotuCost.Value <> 1 Then
                                If Val(.TextMatrix(n, menuStoreCol.现采购价)) <> 0 Then
                                    .TextMatrix(n, menuStoreCol.加成率) = GetFormat(GetFormat((Val(.TextMatrix(n, menuStoreCol.现零售价)) / Val(.TextMatrix(n, menuStoreCol.现采购价)) - 1), 5) * 100, 5)
                                Else
                                    .TextMatrix(n, menuStoreCol.加成率) = GetFormat(0, 5)
                                End If
                            End If
                        End If
                        Dbl数量 = Dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                        dbl成本金额 = dbl成本金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现采购价))
                    End If
                Next
                For n = 1 To vsfPrice.rows - 1
                    If .TextMatrix(intRow, menuStoreCol.药品id) = vsfPrice.TextMatrix(n, menuPriceCol.药品id) Then
                        If Dbl数量 <> 0 Then
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = GetFormat(dbl成本金额 / Dbl数量, mintPriceDigit)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = GetFormat(Dbl金额 / Dbl数量, mintPriceDigit)
                        Else
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = vsfStore.TextMatrix(intRow, menuStoreCol.现采购价)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.现零售价) = vsfStore.TextMatrix(intRow, menuStoreCol.现零售价)
                        End If
                    End If
                Next

                If mint调价 > 0 Then
                    For n = 1 To .rows - 1
                        If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.药品id)) = Val(.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                dbl发票金额 = dbl发票金额 + (Val(.TextMatrix(n, menuStoreCol.现采购价)) - Val(.TextMatrix(n, menuStoreCol.原采购价))) * Val(.TextMatrix(n, menuStoreCol.数量))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.发票金额) = GetFormat(dbl发票金额, mintMoneyDigit)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.加成率
                If Val(.EditText) < 0 Then Exit Sub
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                
                .EditText = GetFormat(.EditText, 5)
                .TextMatrix(intRow, menuStoreCol.加成率) = GetFormat(Val(.EditText), 5)
                .TextMatrix(intRow, menuStoreCol.现零售价) = GetFormat(Val(.TextMatrix(intRow, menuStoreCol.现采购价)) * (1 + Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100), mintCostDigit)
                .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                For n = 1 To .rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.药品id) = .TextMatrix(n, menuStoreCol.药品id) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 0 Or mbln时价药品按批次调价 = False Then
                            .TextMatrix(n, menuStoreCol.加成率) = GetFormat(Val(.EditText), 5)
                            .TextMatrix(n, menuStoreCol.现零售价) = GetFormat(Val(.TextMatrix(n, menuStoreCol.现采购价)) * (1 + GetFormat(Val(.EditText), 5) / 100), mintCostDigit)
                            .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                        End If
                        Dbl数量 = Dbl数量 + .TextMatrix(n, menuStoreCol.数量)
                        Dbl金额 = Dbl金额 + .TextMatrix(n, menuStoreCol.数量) * Val(.TextMatrix(n, menuStoreCol.现零售价))
                    End If
                Next
                If Dbl数量 <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = GetFormat(Dbl金额 / Dbl数量, mintPriceDigit)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.现零售价) = .TextMatrix(intRow, menuStoreCol.现零售价)
                End If
            Case menuStoreCol.现采购价
                If Val(.EditText) > Val(.TextMatrix(.Row, menuStoreCol.现零售价)) Then
                    MsgBox "注意，新成本价大于了新售价！", vbExclamation, gstrSysName
                End If

                If Val(.EditText) < 0 Then
                    MsgBox "成本价不能为负数！", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If .EditText > 9999999 Then
                    MsgBox "采购价过大，请重新输入！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                
                .EditText = GetFormat(.EditText, mintCostDigit)
                .TextMatrix(intRow, menuStoreCol.现采购价) = GetFormat(Val(.EditText), mintCostDigit)
'                If Val(.EditText) <> 0 Then
'                    .TextMatrix(intRow, menuStoreCol.加成率) = GetFormat((Val(.TextMatrix(intRow, menuStoreCol.现零售价)) / Val(.EditText) - 1) * 100, 5)
'                End If
                .TextMatrix(intRow, menuStoreCol.差价差) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原采购价)) * Val(.TextMatrix(intRow, menuStoreCol.数量)), mstrMoneyFormat)
                
                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.现零售价) = GetFormat(GetFormat(Val(.EditText), mintCostDigit) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.加成率)) / 100)), mintPriceDigit)
                    .TextMatrix(intRow, menuStoreCol.调整金额) = Format(Val(.TextMatrix(intRow, menuStoreCol.数量)) * (Val(.TextMatrix(intRow, menuStoreCol.现零售价)) - Val(.TextMatrix(intRow, menuStoreCol.原零售价))), mstrMoneyFormat)
                End If
                
                dbl发票金额 = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.原采购价)) * Val(.TextMatrix(intRow, menuStoreCol.数量))

                For n = 1 To .rows - 1
                    If .TextMatrix(n, menuStoreCol.药品id) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.药品id)) = Val(.TextMatrix(intRow, menuStoreCol.药品id)) And n <> intRow Then
                            If chkCostBatch.Value = 0 Or (Val(.TextMatrix(intRow, menuStoreCol.批次)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.批次)) = Val(.TextMatrix(n, menuStoreCol.批次))) Then
                                dbl现采购价 = Val(.EditText)
                                .TextMatrix(n, menuStoreCol.现采购价) = GetFormat(dbl现采购价, mintCostDigit)
'                                If dbl现采购价 <> 0 Then
'                                    .TextMatrix(n, menuStoreCol.加成率) = GetFormat((Val(.TextMatrix(n, menuStoreCol.现零售价)) / dbl现采购价 - 1) * 100, 5)
'                                End If
                                .TextMatrix(n, menuStoreCol.差价差) = Format((dbl现采购价 - .TextMatrix(n, menuStoreCol.原采购价)) * Val(.TextMatrix(n, menuStoreCol.数量)), mstrMoneyFormat)
                                
                                If Val(.TextMatrix(intRow, menuStoreCol.变价)) = 1 And mbln时价药品按批次调价 = True And mint调价 <> 1 Then
                                    .TextMatrix(n, menuStoreCol.现零售价) = GetFormat(GetFormat(dbl现采购价, mintCostDigit) * (1 + (Val(.TextMatrix(n, menuStoreCol.加成率)) / 100)), mintPriceDigit)
                                    .TextMatrix(n, menuStoreCol.调整金额) = Format(Val(.TextMatrix(n, menuStoreCol.数量)) * (Val(.TextMatrix(n, menuStoreCol.现零售价)) - Val(.TextMatrix(n, menuStoreCol.原零售价))), mstrMoneyFormat)
                                End If
                            Else
                                dbl现采购价 = Val(.TextMatrix(n, menuStoreCol.现采购价))
                            End If
                            dbl发票金额 = dbl发票金额 + (dbl现采购价 - .TextMatrix(n, menuStoreCol.原采购价)) * Val(.TextMatrix(n, menuStoreCol.数量))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.药品id)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.药品id)) Then
                                vsfPay.TextMatrix(n, menuPayCol.发票金额) = Format(dbl发票金额, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If

                If chkCostBatch.Value = 0 Then
                    For n = 1 To vsfPrice.rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.药品id)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.药品id)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.现成本价) = .TextMatrix(intRow, menuStoreCol.现采购价)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.药品id))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.药品id)))  '价格变动，计算平均售价
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lng药品ID As Long)
    '自动计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.现零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品ID Then
                        .TextMatrix(i, menuPriceCol.现零售价) = GetFormat(dblSumPrice / dblSumNumber, mintPriceDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lng药品ID As Long)
    '自动原始计算平均售价
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.药品id) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.药品id)) = lng药品ID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.原零售价)) * Val(.TextMatrix(i, menuStoreCol.数量))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.数量))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.药品id) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.药品id)) = lng药品ID Then
                        .TextMatrix(i, menuPriceCol.原零售价) = GetFormat(dblSumPrice / dblSumNumber, mintPriceDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub




