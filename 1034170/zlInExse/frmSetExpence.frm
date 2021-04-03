VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab stab 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "记帐参数"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt转出"
      Tab(0).Control(1)=   "txtOutDay0"
      Tab(0).Control(2)=   "fraDoctor"
      Tab(0).Control(3)=   "cboSendMateria"
      Tab(0).Control(4)=   "fra记帐药品单位"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "lst收费类别"
      Tab(0).Control(7)=   "fraPrint"
      Tab(0).Control(8)=   "UDOutDay(0)"
      Tab(0).Control(9)=   "chk转出"
      Tab(0).Control(10)=   "fra药房"
      Tab(0).Control(11)=   "fra分离"
      Tab(0).Control(12)=   "lblOutDate(0)"
      Tab(0).Control(13)=   "lbl发药"
      Tab(0).Control(14)=   "Label1"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "结帐参数(&1)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkRefundStyle"
      Tab(1).Control(1)=   "fraMzDepositDefaultUse"
      Tab(1).Control(2)=   "fra缴款控制"
      Tab(1).Control(3)=   "fra输血检查"
      Tab(1).Control(4)=   "chk病历接收"
      Tab(1).Control(5)=   "chk结帐不清"
      Tab(1).Control(6)=   "chk(10)"
      Tab(1).Control(7)=   "cbo代收款项"
      Tab(1).Control(8)=   "chk(16)"
      Tab(1).Control(9)=   "chk(15)"
      Tab(1).Control(10)=   "chk(14)"
      Tab(1).Control(11)=   "fraFeeDate"
      Tab(1).Control(12)=   "chk(12)"
      Tab(1).Control(13)=   "UDOutDay(1)"
      Tab(1).Control(14)=   "txtOutDay1"
      Tab(1).Control(15)=   "chk(11)"
      Tab(1).Control(16)=   "Label2"
      Tab(1).Control(17)=   "lblOutDate(1)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "结算方式顺序设置(&2)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "结帐票据控制(&3)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lbl退款收据"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblListPrint"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblUnit"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblOutUse"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblInUse"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "fra票据格式"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmd退款收据"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "chk(13)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdListPrintSet"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cbo费用明细"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "cmdPrintSetup"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "fraTitle"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cbo退款收据"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "cbo使用类别"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "cmdBillMZ"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "cboInvoiceKindMZ"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cboInvoiceKindZY"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "cmdBillZY"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "其他票据控制(&4)"
      TabPicture(4)   =   "frmSetExpence.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdRed"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraRed"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraDepositPrint"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "fraDeposit"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cmdBillZY 
         Caption         =   "结帐票据设置(&P)"
         Height          =   350
         Left            =   5085
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2505
         Width           =   1560
      End
      Begin VB.ComboBox cboInvoiceKindZY 
         Height          =   300
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2535
         Width           =   3270
      End
      Begin VB.ComboBox cboInvoiceKindMZ 
         Height          =   300
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   2160
         Width           =   3270
      End
      Begin VB.CommandButton cmdBillMZ 
         Caption         =   "结帐票据设置(&P)"
         Height          =   350
         Left            =   5085
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1560
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "结帐红票设置(&P)"
         Height          =   350
         Left            =   -74355
         TabIndex        =   104
         Top             =   5640
         Width           =   1560
      End
      Begin VB.Frame fraRed 
         Caption         =   "结帐红票格式"
         Height          =   1995
         Left            =   -74880
         TabIndex        =   102
         Top             =   3540
         Width           =   6870
         Begin VSFlex8Ctl.VSFlexGrid vsRedFormat 
            Height          =   1605
            Left            =   60
            TabIndex        =   103
            Top             =   285
            Width           =   6705
            _cx             =   11827
            _cy             =   2831
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":0098
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
            ExplorerBar     =   2
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
      End
      Begin VB.Frame fraDepositPrint 
         Caption         =   "预交票据打印方式"
         Height          =   765
         Left            =   -74880
         TabIndex        =   97
         Top             =   2640
         Width           =   6855
         Begin VB.CommandButton cmdDeposit 
            Caption         =   "预交票据打印设置"
            Height          =   350
            Left            =   4710
            TabIndex        =   101
            Top             =   282
            Width           =   1860
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "选择是否打印"
            Height          =   255
            Index           =   2
            Left            =   2985
            TabIndex        =   100
            Top             =   330
            Width           =   1395
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "自动打印"
            Height          =   255
            Index           =   1
            Left            =   1695
            TabIndex        =   99
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "不打印"
            Height          =   255
            Index           =   0
            Left            =   495
            TabIndex        =   98
            Top             =   330
            Width           =   1110
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "本地共用预交票据"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   95
         Top             =   555
         Width           =   6855
         Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
            Height          =   1560
            Left            =   75
            TabIndex        =   96
            Top             =   255
            Width           =   6705
            _cx             =   11827
            _cy             =   2752
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":012A
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
            ExplorerBar     =   2
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
      End
      Begin VB.CheckBox chkRefundStyle 
         Caption         =   "结帐退款缺省按预交缴款结算方式退款"
         Height          =   255
         Left            =   -74700
         TabIndex        =   94
         Top             =   3255
         Width           =   3525
      End
      Begin VB.Frame fraMzDepositDefaultUse 
         Caption         =   "门诊预交缺省使用方式"
         Height          =   900
         Left            =   -74748
         TabIndex        =   90
         Top             =   4980
         Width           =   6636
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "使用剩余所有预交款"
            Height          =   564
            Index           =   2
            Left            =   4476
            TabIndex        =   93
            Top             =   288
            Value           =   -1  'True
            Width           =   2028
         End
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "按结帐金额使用预交"
            Height          =   564
            Index           =   1
            Left            =   2136
            TabIndex        =   92
            Top             =   288
            Width           =   2256
         End
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "不使用预交款"
            Height          =   300
            Index           =   0
            Left            =   252
            TabIndex        =   91
            Top             =   420
            Width           =   1524
         End
      End
      Begin VB.Frame fra缴款控制 
         Caption         =   "结帐缴款控制"
         Height          =   780
         Left            =   -74760
         TabIndex        =   87
         Top             =   4020
         Width           =   6645
         Begin VB.OptionButton opt缴款 
            Caption         =   "存在收取现金时,必须输入缴款"
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   89
            Top             =   315
            Width           =   2835
         End
         Begin VB.OptionButton opt缴款 
            Caption         =   "不进行缴款控制"
            Height          =   315
            Index           =   0
            Left            =   885
            TabIndex        =   88
            Top             =   315
            Value           =   -1  'True
            Width           =   1770
         End
      End
      Begin VB.ComboBox cbo使用类别 
         Height          =   300
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4995
         Width           =   2205
      End
      Begin VB.ComboBox cbo退款收据 
         Height          =   300
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   4650
         Width           =   2040
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用收费票据"
         Height          =   1620
         Left            =   135
         TabIndex        =   83
         Top             =   480
         Width           =   6855
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1290
            Left            =   75
            TabIndex        =   84
            Top             =   255
            Width           =   6705
            _cx             =   11827
            _cy             =   2275
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":0208
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
            ExplorerBar     =   2
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
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "回单票据打印设置(&4)"
         Height          =   350
         Left            =   4695
         TabIndex        =   80
         Top             =   5700
         Width           =   1860
      End
      Begin VB.ComboBox cbo费用明细 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   5340
         Width           =   1980
      End
      Begin VB.CommandButton cmdListPrintSet 
         Caption         =   "打印费用明细设置"
         Height          =   315
         Left            =   2475
         TabIndex        =   78
         Top             =   5715
         Width           =   1635
      End
      Begin VB.CheckBox chk 
         Caption         =   "合约单位结帐每位病人分别打印票据"
         Height          =   225
         Index           =   13
         Left            =   165
         TabIndex        =   68
         Top             =   4695
         Width           =   3300
      End
      Begin VB.CommandButton cmd退款收据 
         Caption         =   "退款收据设置(&S)"
         Height          =   350
         Left            =   240
         TabIndex        =   71
         Top             =   5700
         Width           =   1620
      End
      Begin VB.Frame fra输血检查 
         Caption         =   "结帐时输血费检查"
         Height          =   1110
         Left            =   -70995
         TabIndex        =   72
         Top             =   2730
         Width           =   2880
         Begin VB.OptionButton opt输血 
            Caption         =   "检查并提示"
            Height          =   210
            Index           =   1
            Left            =   390
            TabIndex        =   76
            Top             =   705
            Width           =   1305
         End
         Begin VB.OptionButton opt输血 
            Caption         =   "不检查"
            Height          =   210
            Index           =   0
            Left            =   405
            TabIndex        =   74
            Top             =   435
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Frame fra 
         Height          =   5430
         Left            =   -74490
         TabIndex        =   64
         Top             =   540
         Width           =   5880
         Begin VB.CommandButton cmdUp 
            Caption         =   "↑"
            Height          =   510
            Left            =   5280
            TabIndex        =   66
            Top             =   1140
            Width           =   375
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "↓"
            Height          =   510
            Left            =   5280
            TabIndex        =   65
            Top             =   1740
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBalanceSort 
            Height          =   5055
            Left            =   135
            TabIndex        =   67
            Top             =   240
            Width           =   4995
            _cx             =   8811
            _cy             =   8916
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
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   2
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":02E6
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
            ExplorerBar     =   8
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
      End
      Begin VB.CheckBox chk病历接收 
         Caption         =   "结帐检查病历接收情况"
         Height          =   255
         Left            =   -74700
         TabIndex        =   39
         Top             =   2955
         Width           =   2190
      End
      Begin VB.CheckBox chk结帐不清 
         Caption         =   "结帐后不清除界面信息"
         Height          =   225
         Left            =   -74700
         TabIndex        =   37
         Top             =   2655
         Width           =   2175
      End
      Begin VB.TextBox txt转出 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -70695
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "3"
         Top             =   3090
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "中途结帐缺省退预交款"
         Height          =   195
         Index           =   10
         Left            =   -74700
         TabIndex        =   33
         Top             =   1155
         Width           =   2160
      End
      Begin VB.ComboBox cbo代收款项 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -72750
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3600
         Width           =   1515
      End
      Begin VB.TextBox txtOutDay0 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73965
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "设置为 0 表示只能选择在院病人"
         Top             =   3090
         Width           =   450
      End
      Begin VB.CheckBox chk 
         Caption         =   "有多次住院费用的病人自动弹出结帐设置"
         Height          =   195
         Index           =   16
         Left            =   -74700
         TabIndex        =   38
         Top             =   2355
         Width           =   3720
      End
      Begin VB.CheckBox chk 
         Caption         =   "仅使用指定住院次数的预交款"
         Height          =   195
         Index           =   15
         Left            =   -74700
         TabIndex        =   36
         Top             =   2070
         Width           =   2760
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "显示开单人"
         Height          =   1170
         Left            =   -71880
         TabIndex        =   60
         Top             =   480
         Width           =   1755
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按简码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   19
            Top             =   435
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按编码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   20
            Top             =   735
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboSendMateria 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -73965
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3450
         Width           =   2010
      End
      Begin VB.CheckBox chk 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Index           =   14
         Left            =   -70980
         TabIndex        =   40
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   1095
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.Frame fraFeeDate 
         Caption         =   "结帐费用期间设置"
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -70980
         TabIndex        =   56
         Top             =   1440
         Width           =   2865
         Begin VB.OptionButton optTime 
            Caption         =   "按登记时间"
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton optTime 
            Caption         =   "按发生时间"
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   42
            Top             =   720
            Width           =   1320
         End
      End
      Begin VB.Frame fra记帐药品单位 
         Caption         =   " 药品单位 "
         Height          =   1140
         Left            =   -71880
         TabIndex        =   27
         Top             =   1845
         Width           =   1785
         Begin VB.OptionButton opt记帐药品单位 
            Caption         =   "住院单位"
            Height          =   180
            Index           =   1
            Left            =   195
            TabIndex        =   22
            Top             =   705
            Width           =   1020
         End
         Begin VB.OptionButton opt记帐药品单位 
            Caption         =   "售价单位"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   21
            Top             =   405
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "对病人的零费用进行结帐"
         Height          =   195
         Index           =   12
         Left            =   -74700
         TabIndex        =   34
         Top             =   1755
         Width           =   2280
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   1
         Top             =   345
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "欠费时允许保存为划价单"
            Height          =   195
            Index           =   6
            Left            =   300
            TabIndex        =   8
            Top             =   2115
            Width           =   2400
         End
         Begin VB.CheckBox chk 
            Caption         =   "开单人定开单科室"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   5
            Top             =   1275
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "中药可以输入付数"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   2
            Top             =   435
            Value           =   1  'Checked
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "开单人中包含护士"
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   4
            Top             =   990
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "变价允许输入数次"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   3
            Top             =   720
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "门诊留观病人记帐"
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   6
            Top             =   1560
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "住院留观病人记帐"
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   7
            Top             =   1830
            Width           =   1740
         End
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   1
         Left            =   -73410
         TabIndex        =   46
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay1"
         BuddyDispid     =   196648
         OrigLeft        =   1486
         OrigTop         =   3375
         OrigRight       =   1726
         OrigBottom      =   3645
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOutDay1 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73860
         MaxLength       =   3
         TabIndex        =   32
         Text            =   "0"
         ToolTipText     =   "设置为 0 表示只能选择在院病人"
         Top             =   645
         Width           =   450
      End
      Begin VB.ListBox lst收费类别 
         Height          =   2160
         Left            =   -69930
         Style           =   1  'Checkbox
         TabIndex        =   23
         ToolTipText     =   "请复选允许使用的收费类别"
         Top             =   690
         Width           =   1875
      End
      Begin VB.Frame fraPrint 
         Caption         =   " 打印单据"
         Height          =   1515
         Left            =   -69870
         TabIndex        =   45
         Top             =   3945
         Width           =   1845
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "审核"
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   26
            Top             =   1080
            Width           =   660
         End
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "划价"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   25
            Top             =   720
            Width           =   660
         End
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "记帐"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   24
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "病人出院结帐后自动出院"
         Height          =   195
         Index           =   11
         Left            =   -74700
         TabIndex        =   35
         Top             =   1455
         Width           =   2280
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   0
         Left            =   -73515
         TabIndex        =   61
         Top             =   3090
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay0"
         BuddyDispid     =   196639
         OrigLeft        =   1486
         OrigTop         =   2760
         OrigRight       =   1726
         OrigBottom      =   3030
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chk转出 
         Caption         =   "显示最近   天的转出病人"
         Height          =   195
         Left            =   -71685
         TabIndex        =   10
         Top             =   3120
         Width           =   2370
      End
      Begin VB.Frame fra药房 
         Caption         =   " 药房与发料部门设置 "
         Height          =   1515
         Left            =   -74760
         TabIndex        =   28
         Top             =   3945
         Width           =   4725
         Begin VB.ComboBox cbo卫材 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox chk 
            Caption         =   "显示其它药房库存"
            Height          =   195
            Index           =   8
            Left            =   2400
            TabIndex        =   18
            Top             =   1150
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "显示其它药库库存"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   17
            Top             =   1150
            Width           =   1850
         End
         Begin VB.ComboBox cbo中药 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo西药 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   1305
         End
         Begin VB.ComboBox cbo成药 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发料部门"
            Height          =   180
            Left            =   2100
            TabIndex        =   57
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中草药"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西成药"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中成药"
            Height          =   180
            Left            =   2280
            TabIndex        =   48
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame fra分离 
         Caption         =   " 进行库存检查的药房 "
         ForeColor       =   &H00C00000&
         Height          =   1515
         Left            =   -74760
         TabIndex        =   29
         Top             =   3960
         Visible         =   0   'False
         Width           =   4740
         Begin VB.ListBox lst西药房 
            Height          =   480
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   30
            Top             =   480
            Width           =   1350
         End
         Begin VB.ListBox lst成药房 
            Height          =   480
            Left            =   1485
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   480
            Width           =   1350
         End
         Begin VB.ListBox lst中药房 
            Height          =   480
            Left            =   2880
            Style           =   1  'Checkbox
            TabIndex        =   44
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西药房"
            Height          =   180
            Left            =   90
            TabIndex        =   55
            Top             =   250
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "成药房"
            Height          =   180
            Left            =   1485
            TabIndex        =   54
            Top             =   250
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中药房"
            Height          =   180
            Left            =   2880
            TabIndex        =   53
            Top             =   250
            Width           =   540
         End
      End
      Begin VB.Frame fra票据格式 
         Caption         =   "收费票据格式"
         Height          =   1725
         Left            =   135
         TabIndex        =   81
         Top             =   2880
         Width           =   6870
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1365
            Left            =   60
            TabIndex        =   82
            Top             =   285
            Width           =   6705
            _cx             =   11827
            _cy             =   2408
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":03E0
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
            ExplorerBar     =   2
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
      End
      Begin VB.Label lblInUse 
         AutoSize        =   -1  'True
         Caption         =   "住院结帐票据使用"
         Height          =   180
         Left            =   180
         TabIndex        =   110
         Top             =   2595
         Width           =   1440
      End
      Begin VB.Label lblOutUse 
         AutoSize        =   -1  'True
         Caption         =   "门诊结帐票据使用"
         Height          =   180
         Left            =   180
         TabIndex        =   109
         Top             =   2220
         Width           =   1440
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "合约单位结帐使用                         的票据"
         Height          =   180
         Left            =   165
         TabIndex        =   85
         Top             =   5055
         Width           =   4230
      End
      Begin VB.Label lblListPrint 
         Caption         =   "结帐后打印费用明细"
         Height          =   225
         Left            =   165
         TabIndex        =   77
         Top             =   5385
         Width           =   1665
      End
      Begin VB.Label lbl退款收据 
         AutoSize        =   -1  'True
         Caption         =   "病人退款收据打印"
         Height          =   180
         Left            =   3495
         TabIndex        =   69
         Top             =   4710
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "出院结帐检查代收款项"
         Height          =   180
         Left            =   -74670
         TabIndex        =   63
         Top             =   3660
         Width           =   1800
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许选择         天内出院的病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   62
         Top             =   3135
         Width           =   2790
      End
      Begin VB.Label lbl发药 
         Caption         =   "记帐之后"
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   3495
         Width           =   735
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许选择         天内出院的病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   -74655
         TabIndex        =   52
         Top             =   705
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入类别:"
         Height          =   180
         Left            =   -69960
         TabIndex        =   51
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   1320
      TabIndex        =   59
      Top             =   6405
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4980
      TabIndex        =   73
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6180
      TabIndex        =   75
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   47
      Top             =   6405
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytInFun As Byte '0=记帐,1=结帐
Public mbytUseType As Byte '0:普通记帐,1-科室分散记帐,2-医技科室记帐
Public mstrPrivs As String
Public mlngModul As Long
Public mblnOnlyDrugStock As Boolean  '仅显示药房设置
Private Enum chkBPS
    C0记帐 = 0
    C1划价 = 1
    C2审核 = 2
End Enum
Private Enum chks
    C00中药输付数 = 0
    C01变价输次数 = 1
    C02开单人含护士 = 2
    C03开单人定科室 = 3
    C04门诊留观记帐 = 4
    C05住院留观记帐 = 5
    C06欠费存划价单 = 6
    C07其它药库库存 = 7
    C08其它药房库存 = 8
    C09医保结帐不打 = 9
    C10中途结帐退预交 = 10
    C11结帐自动出院 = 11
    C12零费用可结帐 = 12
    C13合约单位按病人打印 = 13
    C14LED欢迎信息 = 14
    C15仅用指定预交款 = 15
    C16多次住院弹出结帐设置 = 16
End Enum
Private Enum InvoiceKind
    C1收费收据 = 1
    C3结帐收据 = 3
    C4多种收据 = 10
End Enum
Private Const CModule As Long = 1150    '住院记帐操作
Private Sub zlOnlyDrugStrock()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:仅显示药房的相关设置
    '编制:刘兴洪
    '日期:2010-01-25 15:24:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    Err = 0: On Error GoTo ErrHand:
    If mblnOnlyDrugStock And mbytInFun = 0 Then
        For Each ctl In Me.Controls
           Select Case UCase(TypeName(ctl))
           Case UCase("ImageList")
           Case UCase("sstab")
                ctl.Visible = True
           Case Else
                If ctl Is fra分离 Or ctl Is fra药房 Or ctl.Container Is fra药房 Or ctl.Container Is fra分离 Or ctl Is cmdOK Or ctl Is cmdCancel Then
                    ctl.Visible = True
                Else
                     ctl.Visible = False
                End If
           End Select
        Next
        fra药房.Top = Frame1.Top + 200
        fra分离.Top = fra药房.Top
        
        Me.Height = 3525: Me.Width = 5470
        cmdCancel.Top = ScaleHeight - cmdCancel.Height - 100
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 100
        cmdOK.Top = cmdCancel.Top
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
        
        stab.Height = cmdOK.Top - stab.Top - 100
        stab.Width = ScaleWidth - stab.Left * 2
        stab.TabCaption(0) = "药房设置"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'问题:27380
Private Sub chk转出_Click()
    txt转出.Enabled = chk转出.Value = 1
    If txt转出.Visible And txt转出.Enabled Then txt转出.SetFocus
End Sub
Private Sub chk转出_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdDeposit_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdListPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me)
End Sub

Private Sub cmdPrintSetup_Click()
     Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me)
End Sub
Private Sub cmd退款收据_Click()
    '刘兴洪 问题:27776 日期:2010-02-04 16:44:39
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me)
End Sub

Private Sub txt转出_GotFocus()
   zlControl.TxtSelAll txt转出
End Sub

Private Sub txt转出_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cboInvoiceKindZY_Click()
    Dim bytKind As Byte
    If Visible Then '启动时强制调用
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
    End If
End Sub

Private Sub cboInvoiceKindMZ_Click()
    Dim bytKind As Byte
    If Visible Then '启动时强制调用
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
    End If
End Sub

Private Sub cmdBillZY_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdBillMZ_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdRed_Click()
    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6"), Me)
End Sub

Private Sub cmdCancel_Click()
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1137)
End Sub

Private Sub cmdHelp_Click()
    Select Case stab.Tab
        Case 0
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence1"
        Case 1
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence2"
    End Select
End Sub

Private Sub cmdOK_Click()
    Dim strValue As String, i As Long, lngShareID As Long
    Dim blnHavePrivs As Boolean, strTemp As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    If mbytInFun = 0 And cbo西药.Visible Then
        If cbo西药.ListIndex = -1 And cbo西药.ListCount > 0 And cbo西药.Enabled Then
            MsgBox "请选择西药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo西药.SetFocus: Exit Sub
        End If
        If cbo成药.ListIndex = -1 And cbo成药.ListCount > 0 And cbo成药.Enabled Then
            MsgBox "请选择成药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo成药.SetFocus: Exit Sub
        End If
        If cbo中药.ListIndex = -1 And cbo中药.ListCount > 0 And cbo中药.Enabled Then
            MsgBox "请选择中药房.", vbInformation, gstrSysName
            stab.Tab = 0: cbo中药.SetFocus: Exit Sub
        End If
        If cbo卫材.ListIndex = -1 And cbo卫材.ListCount > 0 And cbo卫材.Enabled Then
            MsgBox "请选择卫材发料部门.", vbInformation, gstrSysName
            stab.Tab = 0: cbo卫材.SetFocus: Exit Sub
        End If
    End If
    '保存参数注册信息
    '当不使用门诊留观记帐时,检查如果不显示门诊科室是否有其它可用记帐科室
    If mbytInFun = 0 And (mbytUseType = 0 Or mbytUseType = 1) And chk(chks.C04门诊留观记帐).Value = 0 Then
        If Not CheckUnits Then
            MsgBox "当不使用门诊留观记帐时,你没有可以记帐的科室,参数无法被设置！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    If mbytInFun = 0 Then
    
        '药房
        zlDatabase.SetPara "缺省中药房", IIf(cbo中药.ListIndex = 0, "0", cbo中药.ItemData(cbo中药.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "缺省西药房", IIf(cbo西药.ListIndex = 0, "0", cbo西药.ItemData(cbo西药.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "缺省成药房", IIf(cbo成药.ListIndex = 0, "0", cbo成药.ItemData(cbo成药.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "缺省发料部门", IIf(cbo卫材.ListIndex = 0, "0", cbo卫材.ItemData(cbo卫材.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "显示其它药房库存", chk(chks.C08其它药房库存).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "显示其它药库库存", chk(chks.C07其它药库库存).Value, glngSys, CModule, blnHavePrivs
        '分离发药时的选择
        '--------------------------------------------------------------------------
        strValue = ""
        For i = 0 To lst西药房.ListCount - 1
            If lst西药房.Selected(i) Then
                strValue = strValue & "," & lst西药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "西药房选择", Mid(strValue, 2), glngSys, CModule
        strValue = ""
        For i = 0 To lst成药房.ListCount - 1
            If lst成药房.Selected(i) Then
                strValue = strValue & "," & lst成药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "成药房选择", Mid(strValue, 2), glngSys, CModule
        strValue = ""
        For i = 0 To lst中药房.ListCount - 1
            If lst中药房.Selected(i) Then
                strValue = strValue & "," & lst中药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "中药房选择", Mid(strValue, 2), glngSys, CModule, blnHavePrivs
        '--------------------------------------------------------------------------
        If mblnOnlyDrugStock Then GoTo GoOver:
        
        
        zlDatabase.SetPara "记帐打印", chkBillPrint(chkBPS.C0记帐).Value, glngSys, mlngModul, blnHavePrivs  '不是1150的参数
        
        '1150的参数
        '--------------------------------------------------------------------------------
        '收费类别
        For i = lst收费类别.ListCount - 1 To 0 Step -1
            If lst收费类别.Selected(i) Then strValue = strValue & "'" & Chr(lst收费类别.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "收费类别", strValue, glngSys, CModule, blnHavePrivs
    
           
        '留观病人记帐
        zlDatabase.SetPara "门诊留观病人记帐", chk(chks.C04门诊留观记帐).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "住院留观病人记帐", chk(chks.C05住院留观记帐).Value, glngSys, CModule, blnHavePrivs
        
        zlDatabase.SetPara "出院病人天数", Val(txtOutDay0.Text), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "开单人显示方式", IIf(optDoctorKind(0).Value, 1, 2), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "科室医生", IIf(chk(chks.C03开单人定科室).Value = 1, 0, 1), glngSys, CModule, blnHavePrivs
        
        zlDatabase.SetPara "允许保存为划价单", chk(chks.C06欠费存划价单).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "中药付数", chk(chks.C00中药输付数).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "变价数次", chk(chks.C01变价输次数).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "显示护士", chk(chks.C02开单人含护士).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "记帐药品单位", IIf(opt记帐药品单位(0).Value, 0, 1), glngSys, CModule, blnHavePrivs
        If mbytUseType = 0 Then
            zlDatabase.SetPara "划价打印", chkBillPrint(chkBPS.C1划价).Value, glngSys, CModule, blnHavePrivs
            zlDatabase.SetPara "审核打印", chkBillPrint(chkBPS.C2审核).Value, glngSys, CModule, blnHavePrivs
            zlDatabase.SetPara "记帐后发药", cboSendMateria.ListIndex, glngSys, CModule, blnHavePrivs
        ElseIf mbytUseType = 1 Then
            '刘兴洪 问题:27380 日期:2010-01-22 14:45:32
            zlDatabase.SetPara "最近转出天数", IIf(chk转出.Value = 1, "1", "0") & "|" & Val(txt转出.Text), glngSys, mlngModul, blnHavePrivs
        End If
    Else
        '本地共用结帐票据
        zlDatabase.SetPara "住院结帐票据类型", cboInvoiceKindZY.ListIndex, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "门诊结帐票据类型", cboInvoiceKindMZ.ListIndex, glngSys, mlngModul, blnHavePrivs
        Call SaveInvoice
        
'        lngShareID = 0
'        For i = 1 To lvwBill.ListItems.Count
'            If lvwBill.ListItems(i).Checked Then lngShareID = Val(Mid(lvwBill.ListItems(i).Key, 2))
'        Next
'        zlDatabase.SetPara "共用结帐票据批次", lngShareID, glngSys, mlngModul, blnHavePrivs
        
        'LED设备
        zlDatabase.SetPara "LED显示欢迎信息", chk(chks.C14LED欢迎信息).Value, glngSys, mlngModul, blnHavePrivs
                
        zlDatabase.SetPara "结帐检查代收款项", cbo代收款项.ListIndex, glngSys, mlngModul, blnHavePrivs
        
        zlDatabase.SetPara "出院病人天数", Val(txtOutDay1.Text), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "仅用指定预交款", chk(chks.C15仅用指定预交款).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "多次住院弹出结帐设置", chk(chks.C16多次住院弹出结帐设置).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "在院病人结帐后自动出院", chk(chks.C11结帐自动出院).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "中途结帐退预交", chk(chks.C10中途结帐退预交).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "处理零费用", chk(chks.C12零费用可结帐).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "结帐费用时间", IIf(optTime(1).Value, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "合约单位按病人打印", chk(chks.C13合约单位按病人打印).Value, glngSys, mlngModul, blnHavePrivs
        '刘兴洪 问题:27776 日期:2010-02-04 16:44:39
        zlDatabase.SetPara "退款收据打印", cbo退款收据.ItemData(cbo退款收据.ListIndex), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "结帐后不清除信息", chk结帐不清.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "结帐检查病历接收", chk病历接收.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "结帐退款缺省方式", chkRefundStyle.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "结帐时输血费检查", IIf(opt输血(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs '34260
        zlDatabase.SetPara "结帐明细打印", cbo费用明细.ItemData(cbo费用明细.ListIndex), glngSys, mlngModul, blnHavePrivs
        
        '65352
        zlDatabase.SetPara "门诊预交缺省使用方式", IIf(optMzDeposit(2).Value, 2, IIf(optMzDeposit(1).Value, 1, 0)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "预交票据打印方式", IIf(optBalanceDepositPrint(2).Value, 2, IIf(optBalanceDepositPrint(1).Value, 1, 0)), glngSys, mlngModul, blnHavePrivs
        
        '保存预交票据
        strValue = ""
        With vsDeposit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                    strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("使用类别")))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModul, blnHavePrivs
        
        '43153
        zlDatabase.SetPara "结帐缴款输入控制", IIf(opt缴款(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
        '32322
        With vsBalanceSort
            strTemp = ""
            For i = 1 To .Rows - 1
                strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("结算类别")))
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            zlDatabase.SetPara "结算方式显示顺序", strTemp, glngSys, mlngModul, blnHavePrivs  '30036
        End With
    
    End If
GoOver:
    If mblnOnlyDrugStock Then
        Call zlInit药房
    Else
        Call InitLocPar(mlngModul)
    End If
    gblnOK = True
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If stab.TabVisible(0) Then
        If chk(chks.C00中药输付数).Visible And chk(chks.C00中药输付数).Enabled Then chk(chks.C00中药输付数).SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub


Private Sub Load药房()
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    Set rsTmp = GetDepartments("'中药房','西药房','成药房','发料部门'", "2,3")
        
    cbo中药.AddItem "人工选择"
    cbo西药.AddItem "人工选择"
    cbo成药.AddItem "人工选择"
    cbo卫材.AddItem "人工选择"
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "工作性质='中药房'"
        Do While Not rsTmp.EOF
            cbo中药.AddItem rsTmp!名称
            cbo中药.ItemData(cbo中药.ListCount - 1) = rsTmp!ID
            
            lst中药房.AddItem rsTmp!名称
            lst中药房.ItemData(lst中药房.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "工作性质='西药房'"
        Do While Not rsTmp.EOF
            cbo西药.AddItem rsTmp!名称
            cbo西药.ItemData(cbo西药.ListCount - 1) = rsTmp!ID
            
            lst西药房.AddItem rsTmp!名称
            lst西药房.ItemData(lst西药房.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "工作性质='成药房'"
        Do While Not rsTmp.EOF
            cbo成药.AddItem rsTmp!名称
            cbo成药.ItemData(cbo成药.ListCount - 1) = rsTmp!ID
            
            lst成药房.AddItem rsTmp!名称
            lst成药房.ItemData(lst成药房.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "工作性质='发料部门'"
        Do While Not rsTmp.EOF
            cbo卫材.AddItem rsTmp!名称
            cbo卫材.ItemData(cbo卫材.ListCount - 1) = rsTmp!ID
                            
            rsTmp.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSql As String
    Dim i As Long, strValue As String, blnParSet As Boolean, blnBillOptSet As Boolean
    Dim strDefault As String
    Dim varData As Variant
    Dim bytKind As Byte
    
    gblnOK = False
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0

    If mbytInFun = 0 Then
        blnBillOptSet = InStr(1, GetInsidePrivs(Enum_Inside_Program.p记帐操作), "记帐选项设置") > 0
        '不是1150的参数
        '--------------------------------------------------------------------------------------
    
        '单据打印
        chkBillPrint(chkBPS.C0记帐).Value = IIf(zlDatabase.GetPara("记帐打印", glngSys, mlngModul, , Array(chkBillPrint(chkBPS.C0记帐)), blnParSet) = "1", 1, 0)
        
        
        '1150的参数
        '------------------------------------------------------------------
        '收费类别(挂号除外)
        strSql = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            lst收费类别.AddItem rsTmp!类别
            lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
            rsTmp.MoveNext
        Loop
        strValue = zlDatabase.GetPara("收费类别", glngSys, CModule, , Array(lst收费类别), blnBillOptSet)
        If strValue = "" Then
            For i = 0 To lst收费类别.ListCount - 1
                lst收费类别.Selected(i) = True
            Next
        Else
            For i = 0 To lst收费类别.ListCount - 1
                If InStr(strValue, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
            Next
        End If
        If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
        
        '留观病人记帐
        chk(chks.C04门诊留观记帐).Value = IIf(zlDatabase.GetPara("门诊留观病人记帐", glngSys, CModule, , Array(chk(chks.C04门诊留观记帐)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C05住院留观记帐).Value = IIf(zlDatabase.GetPara("住院留观病人记帐", glngSys, CModule, , Array(chk(chks.C05住院留观记帐)), blnBillOptSet) = "1", 1, 0)
                      
        txtOutDay0.Text = Val(zlDatabase.GetPara("出院病人天数", glngSys, CModule, 0, Array(txtOutDay0, lblOutDate(0), UDOutDay(0)), blnBillOptSet))
        If Val(zlDatabase.GetPara("开单人显示方式", glngSys, CModule, 0, Array(optDoctorKind(0), optDoctorKind(1)), blnBillOptSet)) = 1 Then
            optDoctorKind(0).Value = True
        Else
            optDoctorKind(1).Value = True
        End If
        
        
        chk(chks.C00中药输付数).Value = IIf(zlDatabase.GetPara("中药付数", glngSys, CModule, , Array(chk(chks.C00中药输付数)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C01变价输次数).Value = IIf(zlDatabase.GetPara("变价数次", glngSys, CModule, , Array(chk(chks.C01变价输次数)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C02开单人含护士).Value = IIf(zlDatabase.GetPara("显示护士", glngSys, CModule, , Array(chk(chks.C02开单人含护士)), blnBillOptSet) = "1", 1, 0)
        
        chk(chks.C03开单人定科室).Value = IIf(zlDatabase.GetPara("科室医生", glngSys, CModule, , Array(chk(chks.C03开单人定科室)), blnBillOptSet) = "1", 0, 1)
        chk(chks.C06欠费存划价单).Value = IIf(zlDatabase.GetPara("允许保存为划价单", glngSys, CModule, , Array(chk(chks.C06欠费存划价单)), blnBillOptSet) = "1", 1, 0)
        
                
        i = Val(zlDatabase.GetPara("记帐药品单位", glngSys, CModule, 0, Array(opt记帐药品单位(0), opt记帐药品单位(1)), blnBillOptSet))
        opt记帐药品单位(IIf(i = 0, 0, 1)).Value = True
        
       
        '--------------------------
        Call Load药房
        
        strValue = zlDatabase.GetPara("缺省中药房", glngSys, CModule, , Array(cbo中药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo中药, strValue, True)
        If cbo中药.ListIndex = -1 And Val(strValue) = 0 Then cbo中药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省西药房", glngSys, CModule, , Array(cbo西药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo西药, strValue, True)
        If cbo西药.ListIndex = -1 And Val(strValue) = 0 Then cbo西药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省成药房", glngSys, CModule, , Array(cbo成药), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo成药, strValue, True)
        If cbo成药.ListIndex = -1 And Val(strValue) = 0 Then cbo成药.ListIndex = 0
        
        strValue = zlDatabase.GetPara("缺省发料部门", glngSys, CModule, , Array(cbo卫材), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo卫材, strValue, True)
        If cbo卫材.ListIndex = -1 And Val(strValue) = 0 Then cbo卫材.ListIndex = 0
        
        chk(chks.C08其它药房库存).Value = IIf(zlDatabase.GetPara("显示其它药房库存", glngSys, CModule, , Array(chk(chks.C08其它药房库存)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C07其它药库库存).Value = IIf(zlDatabase.GetPara("显示其它药库库存", glngSys, CModule, , Array(chk(chks.C07其它药库库存)), blnBillOptSet) = "1", 1, 0)
        
        '分离发药时的选择
        '------------------------------------------------------------------
        strValue = zlDatabase.GetPara("西药房选择", glngSys, CModule, , Array(lst西药房), blnBillOptSet)
        For i = 0 To lst西药房.ListCount - 1
            If InStr("," & strValue & ",", "," & lst西药房.ItemData(i) & ",") > 0 Then
                lst西药房.Selected(i) = True
            End If
        Next
        strValue = zlDatabase.GetPara("成药房选择", glngSys, CModule, , Array(lst成药房), blnBillOptSet)
        For i = 0 To lst成药房.ListCount - 1
            If InStr("," & strValue & ",", "," & lst成药房.ItemData(i) & ",") > 0 Then
                lst成药房.Selected(i) = True
            End If
        Next
        strValue = zlDatabase.GetPara("中药房选择", glngSys, CModule, , Array(lst中药房), blnBillOptSet)
        For i = 0 To lst中药房.ListCount - 1
            If InStr("," & strValue & ",", "," & lst中药房.ItemData(i) & ",") > 0 Then
                lst中药房.Selected(i) = True
            End If
        Next
        If lst西药房.ListCount > 0 Then lst西药房.ListIndex = 0
        If lst成药房.ListCount > 0 Then lst成药房.ListIndex = 0
        If lst中药房.ListCount > 0 Then lst中药房.ListIndex = 0
        '------------------------------------------------------------------
        chk转出.Visible = False: txt转出.Visible = False
        If mbytUseType = 0 Then
            chkBillPrint(chkBPS.C1划价).Value = IIf(zlDatabase.GetPara("划价打印", glngSys, CModule, , Array(chkBillPrint(chkBPS.C1划价)), blnBillOptSet) = "1", 1, 0)
            chkBillPrint(chkBPS.C2审核).Value = IIf(zlDatabase.GetPara("审核打印", glngSys, CModule, , Array(chkBillPrint(chkBPS.C2审核)), blnBillOptSet) = "1", 1, 0)
            
            cboSendMateria.AddItem "不发药"
            cboSendMateria.AddItem "自动发药"
            cboSendMateria.AddItem "提示发药"
            i = Val(zlDatabase.GetPara("记帐后发药", glngSys, CModule, 0, Array(cboSendMateria), blnBillOptSet))
            If i > cboSendMateria.ListCount Then i = 0
            cboSendMateria.ListIndex = i
        ElseIf mbytUseType = 1 Then
            '刘兴洪 问题:27380 日期:2010-01-22 14:45:32
            chk转出.Visible = True: txt转出.Visible = True
            Dim str转出 As String
            'CModule
            str转出 = zlDatabase.GetPara("最近转出天数", glngSys, mlngModul, "0|3", Array(chk转出, txt转出), InStr(1, mstrPrivs, ";参数设置;") > 0)
            txt转出.Text = Val(Split(str转出 & "|", "|")(1))
            chk转出.Value = IIf(Val(Split(str转出 & "|", "|")(0)) = 1, 1, 0)
        End If
        
    ElseIf mbytInFun = 1 Then
        '刘兴洪 问题:27776 日期:2010-02-04 16:44:39
        i = Val(zlDatabase.GetPara("退款收据打印", glngSys, mlngModul, , Array(lbl退款收据, cbo退款收据), blnParSet))
        With cbo退款收据
            .AddItem "0-不打印": .ItemData(.NewIndex) = 0: If i = 0 Then .ListIndex = .NewIndex
            .AddItem "1-提示打印": .ItemData(.NewIndex) = 1: .ItemData(.NewIndex) = 1: If i = 1 Then .ListIndex = .NewIndex
            .AddItem "2-打印,但不提示": .ItemData(.NewIndex) = 2: .ItemData(.NewIndex) = 2: If i = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
        End With
        '问题:35511
        i = Val(zlDatabase.GetPara("结帐明细打印", glngSys, mlngModul, , Array(lblListPrint, cbo费用明细), blnParSet))
        With cbo费用明细
            .AddItem "0-不打印": .ItemData(.NewIndex) = 0: If i = 0 Then .ListIndex = .NewIndex
            .AddItem "1-提示打印": .ItemData(.NewIndex) = 1: .ItemData(.NewIndex) = 1: If i = 1 Then .ListIndex = .NewIndex
            .AddItem "2-打印,但不提示": .ItemData(.NewIndex) = 2: .ItemData(.NewIndex) = 2: If i = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
        End With
        chk结帐不清.Value = IIf(Val(zlDatabase.GetPara("结帐后不清除信息", glngSys, mlngModul, , Array(chk结帐不清), blnParSet)) = 1, 1, 0)
        chk病历接收.Value = IIf(Val(zlDatabase.GetPara("结帐检查病历接收", glngSys, mlngModul, , Array(chk病历接收), blnParSet)) = 1, 1, 0) '30036
        chkRefundStyle.Value = IIf(Val(zlDatabase.GetPara("结帐退款缺省方式", glngSys, mlngModul, , Array(chkRefundStyle), blnParSet)) = 1, 1, 0)
       If Val(zlDatabase.GetPara("结帐时输血费检查", glngSys, mlngModul, , Array(opt输血(0), opt输血(1), fra输血检查), blnParSet)) = 1 Then '34260
            opt输血(1).Value = True
       Else
            opt输血(0).Value = True
       End If
       '43153
       If Val(zlDatabase.GetPara("结帐缴款输入控制", glngSys, mlngModul, , Array(opt缴款(0), opt缴款(1), fra缴款控制), blnParSet)) = 1 Then  '34260
            opt缴款(1).Value = True
       Else
            opt缴款(0).Value = True
       End If

        cboInvoiceKindZY.AddItem "住院医疗费收据"
        cboInvoiceKindZY.AddItem "门诊医疗费收据"
        i = Val(zlDatabase.GetPara("住院结帐票据类型", glngSys, mlngModul, 0, Array(cboInvoiceKindZY), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindZY.ListIndex = i
        
        cboInvoiceKindMZ.AddItem "住院医疗费收据"
        cboInvoiceKindMZ.AddItem "门诊医疗费收据"
        i = Val(zlDatabase.GetPara("门诊结帐票据类型", glngSys, mlngModul, 0, Array(cboInvoiceKindMZ), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindMZ.ListIndex = i
        
        If InStr(1, mstrPrivs, ";门诊费用结帐;") = 0 Then '不允许对门诊费用结帐时,只能使用住院医疗费收据
            cboInvoiceKindZY.ListIndex = 0
            cboInvoiceKindZY.Enabled = False
            cboInvoiceKindMZ.Enabled = False
        End If
        
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3结帐收据
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1收费收据
        Else
            bytKind = InvoiceKind.C4多种收据
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3结帐收据, InvoiceKind.C1收费收据))
        '问题:35142
        'Call SetFactBillFormat '设置普通和医保病人结帐发票格式
        
        'LED设备
        chk(chks.C14LED欢迎信息).Value = IIf(zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, "1", Array(chk(chks.C14LED欢迎信息)), blnParSet) = "1", 1, 0)
        
        cbo代收款项.AddItem "0-禁止"
        cbo代收款项.AddItem "1-提醒"
        cbo代收款项.ListIndex = IIf(zlDatabase.GetPara("结帐检查代收款项", glngSys, mlngModul, , Array(cbo代收款项), blnParSet) = "1", 1, 0)
        
        txtOutDay1.Text = Val(zlDatabase.GetPara("出院病人天数", glngSys, mlngModul, 0, Array(txtOutDay1, lblOutDate(1), UDOutDay(1)), blnParSet))
        chk(chks.C13合约单位按病人打印).Value = IIf(zlDatabase.GetPara("合约单位按病人打印", glngSys, mlngModul, , Array(chk(chks.C13合约单位按病人打印)), blnParSet) = "1", 1, 0)
        chk(chks.C15仅用指定预交款).Value = IIf(zlDatabase.GetPara("仅用指定预交款", glngSys, mlngModul, , Array(chk(chks.C15仅用指定预交款)), blnParSet) = "1", 1, 0)
        chk(chks.C16多次住院弹出结帐设置).Value = IIf(zlDatabase.GetPara("多次住院弹出结帐设置", glngSys, mlngModul, , Array(chk(chks.C16多次住院弹出结帐设置)), blnParSet) = "1", 1, 0)
        chk(chks.C10中途结帐退预交).Value = IIf(zlDatabase.GetPara("中途结帐退预交", glngSys, mlngModul, , Array(chk(chks.C10中途结帐退预交)), blnParSet) = "1", 1, 0)
        chk(chks.C11结帐自动出院).Value = IIf(zlDatabase.GetPara("在院病人结帐后自动出院", glngSys, mlngModul, , Array(chk(chks.C11结帐自动出院)), blnParSet) = "1", 1, 0)
        chk(chks.C12零费用可结帐).Value = IIf(zlDatabase.GetPara("处理零费用", glngSys, mlngModul, , Array(chk(chks.C12零费用可结帐)), blnParSet) = "1", 1, 0)
                
        i = Val(zlDatabase.GetPara("结帐费用时间", glngSys, mlngModul, 0, Array(optTime(0), optTime(1)), blnParSet))
        If i <> 0 Then i = 1
        optTime(i).Value = True
        
        '65352
        i = Val(zlDatabase.GetPara("门诊预交缺省使用方式", glngSys, mlngModul, 2, Array(optMzDeposit(0), optMzDeposit(1), optMzDeposit(2), fraMzDepositDefaultUse), blnParSet))
        If i < 0 Or i > 2 Then i = 2
        optMzDeposit(i).Value = True
        
        i = Val(zlDatabase.GetPara("预交票据打印方式", glngSys, mlngModul, 2, Array(optBalanceDepositPrint(0), optBalanceDepositPrint(1), optBalanceDepositPrint(2), fraDepositPrint), blnParSet))
        If i < 0 Or i > 2 Then i = 2
        optBalanceDepositPrint(i).Value = True
        
        
        '32322
        strDefault = "非医保结算-有金额;非医保结算-无金额;医保结算-有金额且允许修改;医保结算-无金额且允许修改;医保结算-有金额且不允许修改;医保结算-无金额且不允许修改"
        strValue = Trim(zlDatabase.GetPara("结算方式显示顺序", glngSys, mlngModul, strDefault, Array(vsBalanceSort, cmdUp, cmdDown), blnParSet))
        varData = Split(strValue, ";")
        With vsBalanceSort
            .Clear 1
            .Rows = 2
            For i = 0 To UBound(varData)
                .TextMatrix(i + 1, .ColIndex("序号")) = i + 1
                 .TextMatrix(i + 1, .ColIndex("结算类别")) = varData(i)
                 If i < UBound(varData) Then .Rows = .Rows + 1
            Next
        End With
    End If
    If mbytInFun = 0 Then
        cboSendMateria.Visible = (mbytInFun = 0 And mbytUseType = 0)
        lbl发药.Visible = (mbytInFun = 0 And mbytUseType = 0)
    
        If gbln分离发药 Then
            fra药房.Visible = False
            fra分离.Visible = True
        End If
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
        stab.TabVisible(3) = False
        stab.TabVisible(4) = False
        If mbytUseType <> 0 Then
            chkBillPrint(1).Visible = False
            chkBillPrint(2).Visible = False
        End If
        
        '问题:27380
        txt转出.Visible = mbytUseType = 1 '科室分散记帐
        chk转出.Visible = mbytUseType = 1 '科室分散记帐

    ElseIf mbytInFun = 1 Then
        If InStr(1, mstrPrivs, ";门诊费用结帐;") = 0 Then chk(chks.C13合约单位按病人打印).Visible = False
        stab.TabVisible(0) = False
    End If
    Call zlOnlyDrugStrock
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'
'Private Sub SetShareInvoice(ByVal bytKind As Byte)
'    Dim rstmp As New ADODB.Recordset, strSQL As String
'    Dim i As Long, lngShareID As Long
'    Dim objItem As ListItem
'
'    '读取可用公用结帐领用
'    Set rstmp = GetShareInvoiceGroupID(bytKind)
'    lngShareID = Val(zlDatabase.GetPara("共用结帐票据批次", glngSys, mlngModul, 0, Array(lvwBill), InStr(1, mstrPrivs, ";参数设置;") > 0))
'    lvwBill.ListItems.Clear
'    For i = 1 To rstmp.RecordCount
'        Set objItem = lvwBill.ListItems.Add(, "_" & rstmp!ID, rstmp!领用人, , 1)
'        objItem.SubItems(1) = Format(rstmp!登记时间, "yyyy-MM-dd")
'        objItem.SubItems(2) = rstmp!开始号码 & "," & rstmp!终止号码
'        objItem.SubItems(3) = rstmp!剩余数量
'        If rstmp!ID = lngShareID Then
'            objItem.Checked = True
'            objItem.Selected = True
'            lngShareID = 0
'        End If
'        rstmp.MoveNext
'    Next
'    If lngShareID <> 0 Then zlDatabase.SetPara "共用结帐票据批次", 0, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mbytUseType = 0
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub
'
'Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Long
'    For i = 1 To lvwBill.ListItems.Count
'        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
'    Next
'    Item.Selected = True
'End Sub

Private Sub txtOutDay0_GotFocus()
    SelAll txtOutDay0
End Sub

Private Sub txtOutDay0_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutDay1_GotFocus()
    SelAll txtOutDay1
End Sub

Private Sub txtOutDay1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function CheckUnits() As Boolean
'功能：检查按参数设置之后,是否有可用记帐临床科室
'说明：当不使用门诊留观记帐之后,将不显示门诊临床科室
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lng病区ID As Long
    Dim strSql As String
    
    On Error GoTo errH
    
    '有权则显示门诊观察室对应的临床科室,住院留观与住院相同
    If InStr(mstrPrivs, ";门诊留观记帐;") And (chk(chks.C04门诊留观记帐).Value = 1) Then
        strSql = "1,2,3"
    Else
        strSql = "2,3"
    End If
    If InStr(";" & mstrPrivs, ";所有病区;") > 0 Then
        strSql = _
             " Select Distinct A.ID,A.编码,A.名称" & _
             " From 部门表 A,部门性质说明 B" & _
             " Where B.部门ID = A.ID And B.服务对象 IN(" & strSql & ") And B.工作性质 IN('临床','手术')" & _
             " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
             " Order by A.编码"
    Else
        '求有权限的科室：本身所在科室+所属病区包含的科室
        '#当操作员属于门诊观察室时，即使没有门诊留观记帐的权限,也显示对应的门诊临床科室,但无法记帐
        strSql = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And B.服务对象 IN(" & strSql & ") And B.工作性质 IN('临床','手术')" & _
            " Order by A.编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    CheckUnits = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub vsBalanceSort_AfterMoveRow(ByVal Row As Long, Position As Long)
    '调整顺序
    '32322
    Call RefreshNO
    Call SetDownAndUpEnable
End Sub
Private Sub RefreshNO()
    Dim lngRow As Long
    With vsBalanceSort
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("序号")) = lngRow
        Next
    End With
End Sub
Private Sub cmdDown_Click()
    With vsBalanceSort
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
    End With
    Call RefreshNO
End Sub
Private Sub cmdUp_Click()
    With vsBalanceSort
        If .Row <= 1 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
    End With
    Call RefreshNO
End Sub
Private Sub SetDownAndUpEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置上下控件的Enable属性
    '编制:刘兴洪
    '日期:2010-09-26 11:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    With vsBalanceSort
        cmdUp.Enabled = vsBalanceSort.Enabled And .Row > 1
        cmdDown.Enabled = vsBalanceSort.Enabled And (.Row < .Rows - 1)
    End With
ErrHand:
End Sub
Private Sub vsBalanceSort_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetDownAndUpEnable
End Sub
'
'Private Sub SetFactBillFormat()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:设置发票格式
'    '编制:刘兴洪
'    '日期:2010-12-31 19:29:48
'    '问题:35142
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strRptName As String, rstmp As ADODB.Recordset, i As Long, blnParSet As Boolean, strSQL As String
'    blnParSet = zlCheckPrivs(mstrPrivs, ";参数设置;")
'    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
'    cboFactNormal.Clear: cboFactMediCare.Clear
'
'    cboFactNormal.AddItem "使用本地缺省格式"
'    cboFactMediCare.AddItem "使用本地缺省格式"
'    '    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
'    strSQL = "" & _
'    "   Select B.说明,B.序号 From zlReports A,zlRptFmts B" & _
'    "    Where A.ID=B.报表ID And A.编号=[1] " & _
'    "   Order by b.序号"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
'    For i = 1 To rstmp.RecordCount
'        cboFactNormal.AddItem rstmp!说明
'        cboFactNormal.ItemData(cboFactNormal.NewIndex) = rstmp!序号
'        cboFactMediCare.AddItem rstmp!说明
'        cboFactMediCare.ItemData(cboFactMediCare.NewIndex) = rstmp!序号
'        rstmp.MoveNext
'    Next
'    cboFactNormal.ListIndex = 0: cboFactMediCare.ListIndex = 0
'    i = Val(zlDatabase.GetPara("普通发票格式", glngSys, mlngModul, , Array(lblFactNormal, cboFactNormal), blnParSet))
'    Call zlControl.CboLocate(cboFactNormal, i, True)
'    i = Val(zlDatabase.GetPara("医保发票格式", glngSys, mlngModul, , Array(lblFactMediCare, cboFactMediCare), blnParSet))
'    Call zlControl.CboLocate(cboFactMediCare, i, True)
'End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(Row, .ColIndex("使用类别"))) = Trim(.TextMatrix(i, .ColIndex("使用类别"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
 
Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("选择")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub vsBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("门诊结帐票据格式"), .ColIndex("结帐后打印方式"), .ColIndex("合约单位结帐"), .ColIndex("住院结帐票据格式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsRedFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
End Sub
Private Sub vsRedFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
End Sub

Private Sub vsRedFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRedFormat
        Select Case Col
        Case .ColIndex("票据格式"), .ColIndex("作废后打印方式"), .ColIndex("合约单位结帐")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存发票相关票据
    '编制:刘兴洪
    '日期:2011-04-28 18:16:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    Dim strPrintMode As String, str合约结帐 As String
    Dim strMZValue As String, strZYValue As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("使用类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用结帐票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    
    '保存收费格式
    strValue = "": strPrintMode = "": str合约结帐 = "普通病人"
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strMZValue = strMZValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("门诊结帐票据格式")))
                strZYValue = strZYValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("住院结帐票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("结帐后打印方式")), 1))
            End If
        Next
        str合约结帐 = Trim(cbo使用类别.Text)
        If strMZValue <> "" Then strMZValue = Mid(strMZValue, 2)
        If strZYValue <> "" Then strZYValue = Mid(strZYValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "门诊结帐发票格式", strMZValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "住院结帐发票格式", strZYValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "病人结帐打印", strPrintMode, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "合约单位结帐打印", str合约结帐, glngSys, mlngModul, blnHavePrivs
    End With
    '保存红票格式
    strValue = "": strPrintMode = ""
    With vsRedFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("作废后打印方式")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "作废发票格式", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "作废发票打印方式", strPrintMode, glngSys, mlngModul, blnHavePrivs
    End With
End Sub


Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-28 18:24:16
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
     If mbytInFun <> 0 Then isValied = True: Exit Function
     
    isValied = False
    On Error GoTo errHandle
    '检查每种使用种式只能一个选择
    With vsBill
        str类别 = "-"
        For i = 1 To vsBill.Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("使用类别"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("使用类别")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("使用类别"))) = Trim(.TextMatrix(j, .ColIndex("使用类别"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    使用类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    If cbo使用类别.ListIndex < 0 Then
        MsgBox "注意:" & vbCrLf & "    你未选择合约单位结帐时所使用的何种票据!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitShareInvoice(ByVal intKind As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant
    Dim VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer, intType2 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSql As String
    Dim strRptName As String, blnHavePrivs As Boolean
    Dim strPrintMode As String, varDataMZ As Variant
    Dim str合约单位结帐 As String, strShareInvoiceMZ As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    On Error GoTo errHandle
    
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "结帐票据格式", False, False
    zl_vsGrid_Para_Restore mlngModul, vsRedFormat, Me.Name, "结帐红票格式", False, False
    strShareInvoice = zlDatabase.GetPara("共用结帐票据批次", glngSys, mlngModul, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And Not blnHavePrivs Then .Editable = flexEDNone
    End With
    
    
    '格式:领用ID1,使用类别1|领用IDn,使用类别n|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(intKind)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!使用类别, " ")
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '住院票据格式处理
    strSql = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("住院结帐票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    strRptName = IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '门诊票据格式处理
    strSql = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    With vsBillFormat
        .ColComboList(.ColIndex("门诊结帐票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("住院结帐发票格式", glngSys, mlngModul, , , True, intType)
    strShareInvoiceMZ = zlDatabase.GetPara("门诊结帐发票格式", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("病人结帐打印", glngSys, mlngModul, , , True, intType1)
    str合约单位结帐 = zlDatabase.GetPara("合约单位结帐打印", glngSys, mlngModul, "普通病人", Array(cbo使用类别, lblUnit), blnHavePrivs)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat
         .ColData(.ColIndex("住院结帐票据格式")) = "0"
         .ColData(.ColIndex("结帐后打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("住院结帐票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("结帐后打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        If (Val(.ColData(.ColIndex("住院结帐票据格式"))) = 1 And _
            Val(.ColData(.ColIndex("结帐后打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        .ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    With vsBillFormat
         .ColData(.ColIndex("门诊结帐票据格式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("门诊结帐票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("门诊结帐票据格式"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("结帐后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    varData = Split(strShareInvoice, "|")
    varDataMZ = Split(strShareInvoiceMZ, "|")
    VarType = Split(strPrintMode, "|")
    strSql = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With vsBillFormat
        .Clear 1: cbo使用类别.Clear
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("结帐后打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("住院结帐票据格式")) = "0"
            .TextMatrix(lngRow, .ColIndex("门诊结帐票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("住院结帐票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varDataMZ)
                varTemp = Split(varDataMZ(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("门诊结帐票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("结帐后打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            cbo使用类别.AddItem Nvl(rsTemp!名称)
            If Nvl(rsTemp!名称) = str合约单位结帐 Then
                cbo使用类别.ListIndex = cbo使用类别.NewIndex
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("结帐后打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("结帐后打印方式"), .Rows - 1, .ColIndex("结帐后打印方式")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("住院结帐票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("住院结帐票据格式"), .Rows - 1, .ColIndex("住院结帐票据格式")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("门诊结帐票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("门诊结帐票据格式"), .Rows - 1, .ColIndex("门诊结帐票据格式")) = vbBlue
        End If
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6")
    '票据格式处理
    strSql = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号=[1]" & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    With vsRedFormat
        .Clear 1
        .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
    End With
    
    '读取参数值
    strShareInvoice = zlDatabase.GetPara("作废发票格式", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("作废发票打印方式", glngSys, mlngModul, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsRedFormat
         .ColData(.ColIndex("票据格式")) = "0"
         .ColData(.ColIndex("作废后打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("作废后打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        If (Val(.ColData(.ColIndex("票据格式"))) = 1 And _
            Val(.ColData(.ColIndex("作废后打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        .ColComboList(.ColIndex("作废后打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    strSql = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With vsRedFormat
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("作废后打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("作废后打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("作废后打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("作废后打印方式"), .Rows - 1, .ColIndex("作废后打印方式")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitDepositInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSql As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    On Error GoTo errHandle
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsDeposit, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModul, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Select Case intType
    Case 1, 3, 5, 15
        vsDeposit.ForeColor = vbBlue: vsDeposit.ForeColorFixed = vbBlue
        fraDepositPrint.ForeColor = vbBlue
    Case Else
        vsDeposit.ForeColor = &H80000008: vsDeposit.ForeColorFixed = &H80000008
        fraDepositPrint.ForeColor = &H80000008
    End Select
    With vsDeposit
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsDeposit
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            '58071
            Select Case Val(Nvl(rsTemp!使用类别, ""))
            Case 0 '不区分门诊和住院票据
                .TextMatrix(lngRow, .ColIndex("使用类别")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("使用类别")) = 0
            Case 1  '门诊票据
                .TextMatrix(lngRow, .ColIndex("使用类别")) = "预交门诊票据"
                .Cell(flexcpData, lngRow, .ColIndex("使用类别")) = 1
            Case Else   '住院票据
                .TextMatrix(lngRow, .ColIndex("使用类别")) = "预交住院票据"
                .Cell(flexcpData, lngRow, .ColIndex("使用类别")) = 2
            End Select
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("使用类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


