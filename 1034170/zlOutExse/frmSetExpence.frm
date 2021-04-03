VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6975
      TabIndex        =   79
      Top             =   6585
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6975
      TabIndex        =   78
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6975
      TabIndex        =   77
      Top             =   345
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   7110
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   12541
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   564
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "单据控制(&1)"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk禁止取消挂号单"
      Tab(0).Control(1)=   "chk抗生素"
      Tab(0).Control(2)=   "fraSetMoneyMode"
      Tab(0).Control(3)=   "fra单位"
      Tab(0).Control(4)=   "chk医保结算光标缺省定位"
      Tab(0).Control(5)=   "chkInsurePartFee"
      Tab(0).Control(6)=   "chk皮试"
      Tab(0).Control(7)=   "fra退费缺省选择方式"
      Tab(0).Control(8)=   "fraDrugNotFee"
      Tab(0).Control(9)=   "chkPayKey"
      Tab(0).Control(10)=   "chk划价立即缴款"
      Tab(0).Control(11)=   "fra库存显示"
      Tab(0).Control(12)=   "chkAddedItem"
      Tab(0).Control(13)=   "txtAddedItem"
      Tab(0).Control(14)=   "cmdAddedItem"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkPrePayPriority"
      Tab(0).Control(16)=   "chkTime"
      Tab(0).Control(17)=   "chk护士"
      Tab(0).Control(18)=   "chk累计"
      Tab(0).Control(19)=   "txtDay"
      Tab(0).Control(20)=   "txtMax"
      Tab(0).Control(21)=   "chkPay"
      Tab(0).Control(22)=   "fra类别"
      Tab(0).Control(23)=   "cbo费别"
      Tab(0).Control(24)=   "cbo结算方式"
      Tab(0).Control(25)=   "udDay"
      Tab(0).Control(26)=   "fraPrintBill"
      Tab(0).Control(27)=   "chk住院按门诊收费"
      Tab(0).Control(28)=   "fra分离"
      Tab(0).Control(29)=   "lblDay"
      Tab(0).Control(30)=   "lblMax"
      Tab(0).Control(31)=   "lbl费别"
      Tab(0).Control(32)=   "lbl结算方式"
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "输入输出(&2)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chk收费执行科室"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkSeekName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra划价通知单打印"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkSeekBill"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdPrintSetup(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraInputItem"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkLed"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chk必须输开单人"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chk不缺省开单人"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "opt分类(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "opt分类(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkMulti"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fra科室与医生"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fra病人"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtSeekDays"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fraLine"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkLedDispDetail"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkLedWelcome"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "fraDoctor"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "fraShortLine"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtNameDays"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkOnlyUnitPatient"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdDeviceSetup"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "chkAutoSplitBill"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chk缺省科室优先"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkUnPopPriceBill"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cboAutoSplitBill"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "fraRegPrompt"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "chkMustRegevent"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "opt分类(2)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "fra缴款控制"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txt收费执行科室"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmd收费执行科室"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "票据控制(&3)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPrintSetup(7)"
      Tab(2).Control(1)=   "picDelBillFormat"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "picBillFormat"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "tbBillSet"
      Tab(2).Control(4)=   "fraFeeExe"
      Tab(2).Control(5)=   "cmdPrintSetup(6)"
      Tab(2).Control(6)=   "fraRefundReceipt"
      Tab(2).Control(7)=   "cmdPrintSetup(5)"
      Tab(2).Control(8)=   "fraFeeList"
      Tab(2).Control(9)=   "cmdPrintSetup(4)"
      Tab(2).Control(10)=   "upd票据张数"
      Tab(2).Control(11)=   "txt票据张数"
      Tab(2).Control(12)=   "chkRegistInvoice"
      Tab(2).Control(13)=   "cmdPrintSetup(2)"
      Tab(2).Control(14)=   "cmdPrintSetup(1)"
      Tab(2).Control(15)=   "cmdPrintSetup(0)"
      Tab(2).Control(16)=   "fraTitle"
      Tab(2).Control(17)=   "chk票据张数"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "票号分配(&4)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picBill(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "药房设置(&5)"
      TabPicture(4)   =   "frmSetExpence.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cbo卫材"
      Tab(4).Control(1)=   "vsfDrugStore"
      Tab(4).Control(2)=   "lbl发料部门"
      Tab(4).ControlCount=   3
      Begin VB.CheckBox chk禁止取消挂号单 
         Caption         =   "禁止取消挂号划价单"
         Height          =   180
         Left            =   -74805
         TabIndex        =   34
         Top             =   5430
         Width           =   1920
      End
      Begin VB.CheckBox chk抗生素 
         Caption         =   "允许录入特殊使用的抗生素"
         Height          =   180
         Left            =   -72450
         TabIndex        =   35
         Top             =   5430
         Width           =   2490
      End
      Begin VB.Frame fraSetMoneyMode 
         Caption         =   "门诊收费刷卡缺省金额操作（三方卡）"
         Height          =   780
         Left            =   -74805
         TabIndex        =   43
         Top             =   6480
         Width           =   6075
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "缺省刷卡金额且金额允许更改"
            Height          =   210
            Index           =   1
            Left            =   210
            TabIndex        =   45
            Top             =   510
            Width           =   2670
         End
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "缺省刷卡金额且金额不允许更改"
            Height          =   210
            Index           =   2
            Left            =   3150
            TabIndex        =   46
            Top             =   510
            Width           =   2820
         End
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "不缺省刷卡金额"
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   44
            Top             =   255
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin VB.Frame fra单位 
         Caption         =   " 药品单位 "
         Height          =   630
         Left            =   -74805
         TabIndex        =   21
         Top             =   2520
         Width           =   4455
         Begin VB.OptionButton opt单位 
            Caption         =   "门诊单位"
            Height          =   180
            Index           =   1
            Left            =   2880
            TabIndex        =   23
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton opt单位 
            Caption         =   "售价单位"
            Height          =   180
            Index           =   0
            Left            =   1590
            TabIndex        =   22
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.Label lbl单位 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收费时按"
            Height          =   180
            Left            =   600
            TabIndex        =   86
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CheckBox chk医保结算光标缺省定位 
         Caption         =   "医保结算光标缺省定位到“医保结算”按钮"
         Height          =   180
         Left            =   -72450
         TabIndex        =   31
         Top             =   4980
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "退费票据打印设置(&2)"
         Height          =   350
         Index           =   7
         Left            =   -70725
         TabIndex        =   193
         Top             =   4020
         Width           =   1950
      End
      Begin VB.PictureBox picDelBillFormat 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   -72780
         ScaleHeight     =   1260
         ScaleWidth      =   6015
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   2340
         Width           =   6015
         Begin VSFlex8Ctl.VSFlexGrid vsDelBillFormat 
            Height          =   1230
            Left            =   30
            TabIndex        =   189
            Top             =   30
            Width           =   5865
            _cx             =   10345
            _cy             =   2170
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
      Begin VB.PictureBox picBillFormat 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   -74850
         ScaleHeight     =   1260
         ScaleWidth      =   6015
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   2100
         Width           =   6015
         Begin VB.CheckBox chkOnePatiPrint 
            Caption         =   "按病人补打票据不根据结算次数补打发票"
            Height          =   180
            Left            =   30
            TabIndex        =   192
            Top             =   30
            Width           =   3540
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1020
            Left            =   30
            TabIndex        =   191
            Top             =   240
            Width           =   5865
            _cx             =   10345
            _cy             =   1799
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
            FormatString    =   $"frmSetExpence.frx":012E
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
      Begin XtremeSuiteControls.TabControl tbBillSet 
         Height          =   1650
         Left            =   -74760
         TabIndex        =   187
         Top             =   1980
         Width           =   5985
         _Version        =   589884
         _ExtentX        =   10557
         _ExtentY        =   2910
         _StockProps     =   64
      End
      Begin VB.CommandButton cmd收费执行科室 
         Caption         =   "…"
         Height          =   280
         Left            =   5850
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   5830
         Width           =   280
      End
      Begin VB.TextBox txt收费执行科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   5820
         Width           =   3735
      End
      Begin VB.CheckBox chkInsurePartFee 
         Caption         =   "多单据分单据结算时，只对医保结算成功的单据收费"
         Height          =   195
         Left            =   -74805
         TabIndex        =   20
         Top             =   2430
         Width           =   4470
      End
      Begin VB.CheckBox chk皮试 
         Caption         =   "提取划价单收费时检查皮试结果"
         Height          =   195
         Left            =   -74805
         TabIndex        =   15
         Top             =   1700
         Width           =   2820
      End
      Begin VB.Frame fraFeeExe 
         Caption         =   "收费执行单"
         Height          =   585
         Left            =   -74760
         TabIndex        =   180
         Top             =   5505
         Width           =   3870
         Begin VB.OptionButton optExe 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   181
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optExe 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   183
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optExe 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   182
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "执行清单打印设置(&7)"
         Height          =   350
         Index           =   6
         Left            =   -70725
         TabIndex        =   175
         Top             =   5775
         Width           =   1950
      End
      Begin VB.Frame fra退费缺省选择方式 
         Caption         =   "退费缺省选择方式"
         Height          =   570
         Left            =   -74805
         TabIndex        =   40
         Top             =   5880
         Width           =   6075
         Begin VB.OptionButton opt退费缺省选择方式 
            Caption         =   "缺省全选择退费项目"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   3840
            TabIndex        =   42
            Top             =   300
            Width           =   2010
         End
         Begin VB.OptionButton opt退费缺省选择方式 
            Caption         =   "缺省按单据号或发票号选择退费项目"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   41
            Top             =   300
            Value           =   -1  'True
            Width           =   3195
         End
      End
      Begin VB.Frame fraRefundReceipt 
         Caption         =   "退费回单控制"
         Height          =   585
         Left            =   -74760
         TabIndex        =   176
         Top             =   4905
         Width           =   3870
         Begin VB.OptionButton optRefund 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   179
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optRefund 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   178
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optRefund 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   177
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "退费回单打印设置(&6)"
         Height          =   350
         Index           =   5
         Left            =   -70725
         TabIndex        =   174
         Top             =   5430
         Width           =   1950
      End
      Begin VB.Frame fraDrugNotFee 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   -74805
         TabIndex        =   170
         Top             =   5625
         Width           =   4125
         Begin VB.OptionButton optDrug 
            Caption         =   "提醒"
            Height          =   180
            Index           =   2
            Left            =   3450
            TabIndex        =   39
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optDrug 
            Caption         =   "禁止"
            Height          =   180
            Index           =   1
            Left            =   2670
            TabIndex        =   38
            Top             =   15
            Width           =   855
         End
         Begin VB.OptionButton optDrug 
            Caption         =   "不检查"
            Height          =   180
            Index           =   0
            Left            =   1770
            TabIndex        =   37
            Top             =   15
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label lblDrugNotFee 
            AutoSize        =   -1  'True
            Caption         =   "药品摆药后退费方式"
            Height          =   180
            Left            =   15
            TabIndex        =   36
            Top             =   15
            Width           =   1620
         End
      End
      Begin VB.CheckBox chkPayKey 
         Caption         =   "使用小键盘的加减(+-)来切换支付方式"
         Height          =   180
         Left            =   -72450
         TabIndex        =   33
         Top             =   5205
         Width           =   3375
      End
      Begin VB.ComboBox cbo卫材 
         Height          =   300
         Left            =   -73575
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   5280
         Width           =   2355
      End
      Begin VB.Frame fra缴款控制 
         Caption         =   "缴款金额输入控制"
         Height          =   975
         Left            =   300
         TabIndex        =   114
         Top             =   4830
         Width           =   5865
         Begin VB.OptionButton opt缴款 
            Caption         =   "收费时按单病人累计"
            Height          =   285
            Index           =   3
            Left            =   2985
            TabIndex        =   130
            Top             =   555
            Width           =   2715
         End
         Begin VB.OptionButton opt缴款 
            Caption         =   $"frmSetExpence.frx":01FC
            Height          =   285
            Index           =   2
            Left            =   2985
            TabIndex        =   117
            Top             =   270
            Width           =   2655
         End
         Begin VB.OptionButton opt缴款 
            Caption         =   "收费时按多病人累计"
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   116
            Top             =   555
            Width           =   2715
         End
         Begin VB.OptionButton opt缴款 
            Caption         =   $"frmSetExpence.frx":021A
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   115
            Top             =   270
            Value           =   -1  'True
            Width           =   3780
         End
      End
      Begin VB.CheckBox chk划价立即缴款 
         Caption         =   "提取划价单后立即缴款"
         Height          =   180
         Left            =   -74805
         TabIndex        =   32
         Top             =   5220
         Width           =   2160
      End
      Begin VB.Frame fraFeeList 
         Caption         =   "收费后费用清单"
         Height          =   585
         Left            =   -74760
         TabIndex        =   126
         Top             =   4290
         Width           =   3870
         Begin VB.OptionButton optPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   128
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   127
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   129
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "医保回单打印设置(&5)"
         Height          =   350
         Index           =   4
         Left            =   -70725
         TabIndex        =   124
         Top             =   5070
         Width           =   1950
      End
      Begin VB.OptionButton opt分类 
         Caption         =   "按单据分类汇总显示"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   63
         Top             =   3750
         Width           =   2280
      End
      Begin VB.Frame fra库存显示 
         Caption         =   "库存显示"
         Height          =   1155
         Left            =   -74820
         TabIndex        =   118
         Top             =   3240
         Width           =   4455
         Begin VB.OptionButton opt库存 
            Caption         =   "仅显示有无"
            Height          =   180
            Index           =   1
            Left            =   2835
            TabIndex        =   123
            Top             =   810
            Width           =   1215
         End
         Begin VB.OptionButton opt库存 
            Caption         =   "显示库存数"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   121
            Top             =   810
            Width           =   1290
         End
         Begin VB.CheckBox chk药房 
            Caption         =   "显示其它药房库存"
            Height          =   195
            Left            =   150
            TabIndex        =   120
            Top             =   375
            Width           =   1770
         End
         Begin VB.CheckBox chk药库 
            Caption         =   "显示其它药库库存"
            Height          =   195
            Left            =   2250
            TabIndex        =   119
            Top             =   390
            Width           =   1770
         End
         Begin VB.Line lnSplit 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   15
            X2              =   4425
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line lnSplit 
            BorderColor     =   &H80000000&
            Index           =   1
            X1              =   15
            X2              =   4440
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Label lbl库存显示方式 
            AutoSize        =   -1  'True
            Caption         =   "库存显示方式"
            Height          =   180
            Left            =   300
            TabIndex        =   122
            Top             =   810
            Width           =   1080
         End
      End
      Begin MSComCtl2.UpDown upd票据张数 
         Height          =   300
         Left            =   -73275
         TabIndex        =   103
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt票据张数"
         BuddyDispid     =   196656
         OrigLeft        =   1500
         OrigTop         =   3285
         OrigRight       =   1755
         OrigBottom      =   3570
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt票据张数 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -73740
         TabIndex        =   102
         Text            =   "10"
         Top             =   3960
         Width           =   465
      End
      Begin VB.CheckBox chkMustRegevent 
         Caption         =   "收费时检查病人挂号科室"
         Height          =   195
         Left            =   315
         TabIndex        =   68
         ToolTipText     =   "要求必须挂了该科室的号才能保存开单科室为该科室的费用单据"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Frame fraRegPrompt 
         Caption         =   "未挂号病人收费"
         Height          =   990
         Left            =   4560
         TabIndex        =   110
         Top             =   465
         Visible         =   0   'False
         Width           =   1620
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "禁止"
            Height          =   180
            Index           =   2
            Left            =   210
            TabIndex        =   113
            Top             =   720
            Width           =   1020
         End
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "允许"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   112
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "提醒"
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   111
            Top             =   490
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboAutoSplitBill 
         Height          =   300
         Left            =   1880
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   4800
         Width           =   1170
      End
      Begin VB.CheckBox chkUnPopPriceBill 
         Caption         =   "不弹出划价单选择窗口"
         Height          =   195
         Left            =   3975
         TabIndex        =   67
         Top             =   4530
         Width           =   2160
      End
      Begin VB.CheckBox chk缺省科室优先 
         Caption         =   "缺省科室优先"
         Height          =   195
         Left            =   2160
         TabIndex        =   109
         Top             =   2760
         Width           =   1620
      End
      Begin VB.CheckBox chkAutoSplitBill 
         Caption         =   "收费明细自动按              组合单据"
         Height          =   195
         Left            =   315
         TabIndex        =   69
         Top             =   4830
         Width           =   3960
      End
      Begin VB.CheckBox chkAddedItem 
         Caption         =   "未挂号时自动加收收费项目"
         Height          =   195
         Left            =   -74805
         TabIndex        =   17
         Top             =   2200
         Width           =   2460
      End
      Begin VB.TextBox txtAddedItem 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   -72250
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2162
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddedItem 
         Caption         =   "…"
         Height          =   280
         Left            =   -70680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2157
         Width           =   280
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   4605
         TabIndex        =   108
         Top             =   4830
         Width           =   1500
      End
      Begin VB.CheckBox chkPrePayPriority 
         Caption         =   "优先使用预交款缴费"
         Height          =   195
         Left            =   -74805
         TabIndex        =   16
         Top             =   1950
         Width           =   2340
      End
      Begin VB.CheckBox chkOnlyUnitPatient 
         Caption         =   "只查找合约单位病人"
         Height          =   195
         Left            =   2160
         TabIndex        =   107
         Top             =   3000
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2925
         MaxLength       =   3
         TabIndex        =   105
         Text            =   "0"
         ToolTipText     =   "0表示查找时不限制时间"
         Top             =   2520
         Width           =   285
      End
      Begin VB.Frame fraShortLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2925
         TabIndex        =   104
         Top             =   2700
         Width           =   285
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "显示开单人"
         Height          =   885
         Left            =   4560
         TabIndex        =   98
         Top             =   1530
         Width           =   1620
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按编码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   100
            Top             =   540
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "按简码"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   99
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   3975
         TabIndex        =   93
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   4050
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkLedDispDetail 
         Caption         =   "LED显示收费明细"
         Height          =   225
         Left            =   3975
         TabIndex        =   92
         ToolTipText     =   "收费窗口,输入收费项目后是否显示信息"
         Top             =   3810
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkRegistInvoice 
         Caption         =   "挂号时使用与收费相同的票据"
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   3705
         Width           =   2640
      End
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1695
         TabIndex        =   88
         Top             =   4500
         Width           =   405
      End
      Begin VB.TextBox txtSeekDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   1695
         MaxLength       =   4
         TabIndex        =   66
         Text            =   "1"
         Top             =   4305
         Width           =   435
      End
      Begin VB.Frame fra病人 
         Caption         =   "病人来源"
         Height          =   990
         Left            =   285
         TabIndex        =   87
         Top             =   465
         Width           =   1380
         Begin VB.OptionButton opt病人 
            Caption         =   "门诊病人"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   47
            Top             =   330
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt病人 
            Caption         =   "住院病人"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   48
            Top             =   650
            Width           =   1020
         End
      End
      Begin VB.Frame fra科室与医生 
         Caption         =   "科室与医生"
         Height          =   990
         Left            =   1800
         TabIndex        =   83
         Top             =   465
         Width           =   2685
         Begin VB.OptionButton optSelf 
            Caption         =   "科室和医生互相独立输入"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   720
            Width           =   2280
         End
         Begin VB.OptionButton optDoctor 
            Caption         =   "通过输入医生来确定科室"
            Height          =   180
            Left            =   180
            TabIndex        =   50
            Top             =   490
            Width           =   2280
         End
         Begin VB.OptionButton optUnit 
            Caption         =   "通过输入科室来确定医生"
            Height          =   180
            Left            =   180
            TabIndex        =   49
            Top             =   270
            Value           =   -1  'True
            Width           =   2280
         End
      End
      Begin VB.CheckBox chkMulti 
         Caption         =   "收费时允许同时输入多张单据"
         Height          =   195
         Left            =   315
         TabIndex        =   64
         Top             =   4035
         Width           =   3000
      End
      Begin VB.OptionButton opt分类 
         Caption         =   "以收入项目显示分类合计"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   62
         Top             =   3530
         Width           =   2280
      End
      Begin VB.OptionButton opt分类 
         Caption         =   "以收据费目显示分类合计"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   61
         Top             =   3300
         Value           =   -1  'True
         Width           =   2280
      End
      Begin VB.CheckBox chk不缺省开单人 
         Caption         =   "不使用缺省开单人"
         Height          =   195
         Left            =   315
         TabIndex        =   59
         Top             =   2760
         Width           =   1740
      End
      Begin VB.CheckBox chk必须输开单人 
         Caption         =   "必须要输入开单人"
         Height          =   195
         Left            =   315
         TabIndex        =   60
         Top             =   3000
         Width           =   1740
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收费清单打印设置(&4)"
         Height          =   350
         Index           =   2
         Left            =   -70725
         TabIndex        =   76
         Top             =   4725
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收据证明打印设置(&3)"
         Height          =   350
         Index           =   1
         Left            =   -70725
         TabIndex        =   75
         Top             =   4380
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "收费票据打印设置(&1)"
         Height          =   350
         Index           =   0
         Left            =   -70725
         TabIndex        =   74
         Top             =   3675
         Width           =   1950
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "变价输入数次"
         Height          =   195
         Left            =   -72030
         TabIndex        =   5
         Top             =   810
         Width           =   1380
      End
      Begin VB.CheckBox chkLed 
         Caption         =   "人工控制LED报价"
         Height          =   225
         Left            =   3975
         TabIndex        =   72
         Top             =   3540
         Width           =   1650
      End
      Begin VB.CheckBox chk护士 
         Caption         =   "开单人含护士"
         Height          =   195
         Left            =   -72030
         TabIndex        =   6
         Top             =   1080
         Width           =   1380
      End
      Begin VB.CheckBox chk累计 
         Caption         =   "显示收款累计"
         Height          =   195
         Left            =   -72030
         TabIndex        =   7
         Top             =   1350
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用收费票据"
         Height          =   1500
         Left            =   -74775
         TabIndex        =   85
         Top             =   450
         Width           =   6000
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1155
            Left            =   75
            TabIndex        =   125
            Top             =   255
            Width           =   5790
            _cx             =   10213
            _cy             =   2037
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
            FormatString    =   $"frmSetExpence.frx":0238
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
      Begin VB.TextBox txtDay 
         ForeColor       =   &H80000012&
         Height          =   270
         Left            =   -73665
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "0"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtMax 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -73665
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   525
         Width           =   1335
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "中药输入付数"
         Height          =   195
         Left            =   -72030
         TabIndex        =   4
         Top             =   555
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.Frame fra类别 
         Caption         =   "可用收费类别"
         Height          =   4020
         Left            =   -70275
         TabIndex        =   28
         Top             =   480
         Width           =   1485
         Begin VB.ListBox lst收费类别 
            ForeColor       =   &H00C00000&
            Height          =   3630
            Left            =   105
            Style           =   1  'Checkbox
            TabIndex        =   29
            ToolTipText     =   "请复选允许使用的收费类别"
            Top             =   255
            Width           =   1275
         End
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   915
         Width           =   1350
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1305
         Width           =   1350
      End
      Begin VB.Frame fraInputItem 
         Caption         =   "收费或划价时要输入的项目"
         Height          =   885
         Left            =   285
         TabIndex        =   80
         Top             =   1530
         Width           =   4170
         Begin VB.CheckBox chk医疗付款 
            Caption         =   "医疗付款方式"
            Height          =   210
            Left            =   2520
            TabIndex        =   57
            Top             =   540
            Value           =   1  'Checked
            Width           =   1380
         End
         Begin VB.CheckBox chk性别 
            Caption         =   "性别"
            Height          =   210
            Left            =   165
            TabIndex        =   52
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk年龄 
            Caption         =   "年龄"
            Height          =   210
            Left            =   2520
            TabIndex        =   55
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk费别 
            Caption         =   "费别"
            Height          =   210
            Left            =   3240
            TabIndex        =   56
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk是否加班 
            Caption         =   "是否加班"
            Height          =   210
            Left            =   1350
            TabIndex        =   53
            Top             =   270
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk开单日期 
            Caption         =   "开单日期"
            Height          =   210
            Left            =   165
            TabIndex        =   54
            Top             =   540
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk开单人 
            Caption         =   "开单人"
            Height          =   210
            Left            =   1350
            TabIndex        =   58
            Top             =   540
            Value           =   1  'Checked
            Width           =   840
         End
      End
      Begin MSComCtl2.UpDown udDay 
         Height          =   270
         Left            =   -72570
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1665
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDay"
         BuddyDispid     =   196694
         OrigLeft        =   3045
         OrigTop         =   615
         OrigRight       =   3285
         OrigBottom      =   885
         Max             =   32767
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "划价通知单打印设置(&1)"
         Height          =   350
         Index           =   3
         Left            =   3840
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkSeekBill 
         Caption         =   "自动搜寻病人      天内的划价单据"
         Height          =   195
         Left            =   315
         TabIndex        =   65
         Top             =   4305
         Width           =   3180
      End
      Begin VB.Frame fra划价通知单打印 
         Caption         =   "划价通知单打印"
         Height          =   1230
         Left            =   3840
         TabIndex        =   94
         Top             =   3240
         Visible         =   0   'False
         Width           =   2325
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   97
            Top             =   900
            Width           =   1500
         End
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   96
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Value           =   -1  'True
            Width           =   1260
         End
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
         Height          =   195
         Left            =   315
         TabIndex        =   106
         Top             =   2535
         Width           =   4260
      End
      Begin VB.CheckBox chk票据张数 
         Caption         =   "票据剩余         张时开始提醒收费员"
         Height          =   285
         Left            =   -74760
         TabIndex        =   101
         Top             =   3960
         Width           =   3450
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4695
         Left            =   -74775
         TabIndex        =   131
         Top             =   540
         Width           =   5655
         _cx             =   9975
         _cy             =   8281
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetExpence.frx":0316
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
      Begin VB.Frame fraPrintBill 
         Caption         =   "打印单据"
         Height          =   1185
         Left            =   -72000
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1605
         Begin VB.CheckBox chk 
            Caption         =   "记帐时"
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   9
            Top             =   300
            Width           =   840
         End
         Begin VB.CheckBox chk 
            Caption         =   "划价时"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   10
            Top             =   555
            Width           =   840
         End
         Begin VB.CheckBox chk 
            Caption         =   "审核时"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   11
            Top             =   810
            Width           =   840
         End
      End
      Begin VB.CheckBox chk收费执行科室 
         Caption         =   "本机收费执行科室"
         Height          =   210
         Left            =   330
         TabIndex        =   186
         Top             =   5880
         Width           =   1770
      End
      Begin VB.PictureBox picBill 
         BorderStyle     =   0  'None
         Height          =   4635
         Index           =   2
         Left            =   -74850
         ScaleHeight     =   4635
         ScaleWidth      =   5985
         TabIndex        =   134
         Top             =   615
         Width           =   5985
         Begin VB.PictureBox picRuleBack 
            BorderStyle     =   0  'None
            Height          =   2790
            Index           =   0
            Left            =   -15
            ScaleHeight     =   2790
            ScaleWidth      =   6255
            TabIndex        =   143
            Top             =   495
            Visible         =   0   'False
            Width           =   6255
            Begin VB.CheckBox chk体检按单据分别打印 
               Caption         =   "体检病人每张单据分别打印(该参数同时影响工本费数量计算)"
               Height          =   195
               Left            =   630
               TabIndex        =   145
               Top             =   300
               Width           =   5160
            End
            Begin VB.CheckBox chkAutoAddBookFee 
               Caption         =   "门诊收费时自动加收工本费"
               Height          =   195
               Left            =   345
               TabIndex        =   147
               Top             =   825
               Width           =   2460
            End
            Begin VB.CheckBox chkOlnyOneBill 
               Caption         =   "收费每次打印只用一张票据(该参数同时影响工本费数量计算)"
               Height          =   195
               Left            =   345
               TabIndex        =   146
               Top             =   555
               Width           =   5160
            End
            Begin VB.Frame fraActuallyPrint 
               Height          =   1695
               Left            =   150
               TabIndex        =   148
               Top             =   825
               Width           =   5850
               Begin VB.CheckBox chkErrorItemNotBill 
                  Caption         =   "误差项不使用票据"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   149
                  Top             =   885
                  Width           =   1740
               End
               Begin VB.OptionButton optBillMode 
                  Caption         =   "打印收据费目"
                  Height          =   255
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   153
                  Top             =   1245
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton optBillMode 
                  Caption         =   "打印收费项目"
                  Height          =   255
                  Index           =   1
                  Left            =   4200
                  TabIndex        =   154
                  Top             =   1245
                  Width           =   1455
               End
               Begin VB.CheckBox chkExcuteDept 
                  Caption         =   "按执行科室分别打印"
                  Height          =   195
                  Left            =   200
                  TabIndex        =   152
                  Top             =   1275
                  Width           =   1980
               End
               Begin MSComCtl2.UpDown updRows 
                  Height          =   300
                  Left            =   4320
                  TabIndex        =   151
                  TabStop         =   0   'False
                  Top             =   825
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   3
                  BuddyControl    =   "txtRowsUD"
                  BuddyDispid     =   196729
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.TextBox txtRowsUD 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   4005
                  Locked          =   -1  'True
                  TabIndex        =   150
                  Text            =   "3"
                  Top             =   832
                  Width           =   330
               End
               Begin VB.Label lblRows 
                  AutoSize        =   -1  'True
                  Caption         =   "收费收据行次"
                  Height          =   180
                  Left            =   2850
                  TabIndex        =   160
                  Top             =   892
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Caption         =   "工本费数量由票据张数决定,票据张数按以下规则计算.但实际打印张数由票据数据源及票据设计决定,如果两者不一致,工本费数量将不准确."
                  Height          =   495
                  Index           =   25
                  Left            =   120
                  TabIndex        =   155
                  Top             =   360
                  Width           =   5655
               End
            End
            Begin VB.CheckBox chkBillNO 
               Caption         =   "门诊收费每张单据分别打印(该参数同时影响工本费数量计算)"
               Height          =   195
               Left            =   345
               TabIndex        =   144
               Top             =   75
               Width           =   5160
            End
         End
         Begin VB.ComboBox cboBillRole 
            Height          =   300
            ItemData        =   "frmSetExpence.frx":03A4
            Left            =   1125
            List            =   "frmSetExpence.frx":03A6
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   105
            Width           =   3015
         End
         Begin VB.PictureBox picRuleBack 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   2
            Left            =   30
            ScaleHeight     =   1035
            ScaleWidth      =   6330
            TabIndex        =   156
            Top             =   435
            Visible         =   0   'False
            Width           =   6330
            Begin VB.Label lblCustomInfor 
               Caption         =   $"frmSetExpence.frx":03A8
               Height          =   570
               Left            =   165
               TabIndex        =   157
               Top             =   375
               Width           =   5670
            End
         End
         Begin VB.PictureBox picRuleBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3825
            Index           =   1
            Left            =   15
            ScaleHeight     =   3825
            ScaleWidth      =   6150
            TabIndex        =   135
            Top             =   435
            Visible         =   0   'False
            Width           =   6150
            Begin VB.Frame fraRuleSystem 
               Height          =   3525
               Left            =   0
               TabIndex        =   136
               Top             =   105
               Width           =   5955
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "按执行科室分组汇总"
                  Height          =   240
                  Index           =   2
                  Left            =   2985
                  TabIndex        =   173
                  Top             =   2250
                  Width           =   2025
               End
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "首页打印汇总"
                  Height          =   240
                  Index           =   1
                  Left            =   1425
                  TabIndex        =   172
                  Top             =   2250
                  Width           =   1440
               End
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "不汇总"
                  Height          =   240
                  Index           =   0
                  Left            =   330
                  TabIndex        =   171
                  Top             =   2250
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   2
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   167
                  Text            =   "3"
                  Top             =   1875
                  Width           =   330
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   1
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   164
                  Text            =   "3"
                  Top             =   1530
                  Width           =   330
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   0
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   161
                  Text            =   "3"
                  Top             =   1155
                  Width           =   345
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "4.按收费细目分页"
                  Height          =   180
                  Index           =   3
                  Left            =   285
                  TabIndex        =   140
                  Top             =   1920
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "3.按收据费目分页"
                  Height          =   180
                  Index           =   2
                  Left            =   270
                  TabIndex        =   139
                  Top             =   1575
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "2.按执行科室分页"
                  Height          =   180
                  Index           =   1
                  Left            =   270
                  TabIndex        =   138
                  Top             =   1215
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "1.按单据分页"
                  Height          =   225
                  Index           =   0
                  Left            =   270
                  TabIndex        =   137
                  Top             =   915
                  Width           =   1635
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   0
                  Left            =   2775
                  TabIndex        =   162
                  TabStop         =   0   'False
                  Top             =   1155
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   1
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(0)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   0
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   165
                  TabStop         =   0   'False
                  Top             =   1530
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   4
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(1)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   1
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   2
                  Left            =   2760
                  TabIndex        =   168
                  TabStop         =   0   'False
                  Top             =   1875
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   20
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(2)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   2
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "：每　　   个收费细目分一页"
                  Height          =   180
                  Index           =   2
                  Left            =   2025
                  TabIndex        =   169
                  Top             =   1935
                  Width           =   2430
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "：每　　   个收据费目分一页"
                  Height          =   180
                  Index           =   1
                  Left            =   2025
                  TabIndex        =   166
                  Top             =   1590
                  Width           =   2430
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "：每　　   个执行科室分一页"
                  Height          =   180
                  Index           =   0
                  Left            =   2025
                  TabIndex        =   163
                  Top             =   1215
                  Width           =   2430
               End
               Begin VB.Label lblInfor 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   540
                  Left            =   90
                  TabIndex        =   142
                  Top             =   2880
                  Width           =   5760
               End
               Begin VB.Label lblRuleSystem 
                  Caption         =   "工本费数量由票据张数决定,票据张数按以下规则计算.但实际打印张数由收费划价单决定,如果手工录入费用单据,工本费的数量计算将不准确."
                  Height          =   585
                  Left            =   180
                  TabIndex        =   141
                  Top             =   330
                  Width           =   5730
               End
            End
         End
         Begin VB.Label lblBillRole 
            AutoSize        =   -1  'True
            Caption         =   "票据分配规则"
            Height          =   180
            Left            =   15
            TabIndex        =   159
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.CheckBox chk住院按门诊收费 
         Caption         =   "住院病人按门诊收费"
         Height          =   180
         Left            =   -74805
         TabIndex        =   30
         Top             =   4980
         Width           =   2025
      End
      Begin VB.Frame fra分离 
         Caption         =   " 从以下药房检查库存 "
         Height          =   1230
         Left            =   -74805
         TabIndex        =   24
         Top             =   3750
         Visible         =   0   'False
         Width           =   4440
         Begin VB.ListBox lst中药房 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   2955
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   465
            Width           =   1350
         End
         Begin VB.ListBox lst成药房 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   1560
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   465
            Width           =   1350
         End
         Begin VB.ListBox lst西药房 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   165
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   465
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中药房"
            Height          =   180
            Left            =   2955
            TabIndex        =   91
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "成药房"
            Height          =   180
            Left            =   1560
            TabIndex        =   90
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西药房"
            Height          =   180
            Left            =   165
            TabIndex        =   89
            Top             =   255
            Width           =   540
         End
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取消划价超过                 天未处理的划价单"
         Height          =   180
         Left            =   -74805
         TabIndex        =   12
         Top             =   1680
         Width           =   4050
      End
      Begin VB.Label lbl发料部门 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省发料部门"
         Height          =   180
         Left            =   -74775
         TabIndex        =   133
         Top             =   5340
         Width           =   1080
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据最大金额"
         Height          =   180
         Left            =   -74820
         TabIndex        =   84
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省病人费别"
         Height          =   180
         Left            =   -74820
         TabIndex        =   82
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lbl结算方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省结算方式"
         Height          =   180
         Left            =   -74805
         TabIndex        =   81
         Top             =   1365
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mbytInFun As Byte '0=收费,1=划价,2=门诊记帐
Public mstrPrivs As String
Public mlngModul As Long
Public mblnSetDrugStore As Boolean
Private mblnNotClick As Boolean

Private Sub cboBillRole_Click()
     '56963
      Call SetBillNoRule
End Sub
Private Function GetPrintListHaveData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据打印明细是否有数据
    '返回:有数据返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-05-17 14:24:40
    '说明:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 1 From 票据打印明细 where Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetPrintListHaveData = rsTemp.RecordCount >= 1
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowRuleInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示票号的分配规则
    '编制:刘兴洪
    '日期:2013-03-26 14:14:08
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, i As Integer
    Dim strName As String
    
    On Error GoTo errHandle
    strInfor = ""
    If chkBillRule(0).Value = 1 Then
            strInfor = strInfor & "+ NO"
    End If
    For i = 1 To 3
        If chkBillRule(i).Value = 1 Then
            strName = Switch(i = 1, "执行科室", i = 2, "收据费目", True, "收据细目")
            strInfor = strInfor & "+" & strName & "(" & txtBillRuleNum(i - 1).Text & ")"
        End If
    Next
    If strInfor <> "" Then strInfor = Mid(strInfor, 2)
    lblInfor.Caption = strInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chkAddedItem_Click()
    Dim i As Long
    If chkAddedItem.Value = 1 Then
        If txtAddedItem.Text = "" And Me.Visible Then cmdAddedItem_Click
    
        '其它类别可用
        For i = 0 To lst收费类别.ListCount - 1
            If lst收费类别.ItemData(i) = Asc("Z") Then
                If lst收费类别.Selected(i) = False Then lst收费类别.Selected(i) = True
            End If
        Next
    Else
        txtAddedItem.Text = ""
    End If
End Sub

Private Sub chkAutoSplitBill_Click()
    cboAutoSplitBill.Enabled = chkAutoSplitBill.Value = 1 And cboAutoSplitBill.Tag = "1"
End Sub

Private Sub chkBillNO_Click()
    chk体检按单据分别打印.Enabled = (chkBillNO.Value = vbChecked)
End Sub

Private Sub chkBillRule_Click(Index As Integer)
    '56963
    If Index <> 0 And chkBillRule(Index).Value = 1 Then
        If Val(txtBillRuleNum(Index - 1).Text) = 0 Then
            updBillRuleNum(Index - 1).Value = Val(txtBillRuleNum(Index - 1).Tag)    '恢复缺省值
        End If
    End If
    Call SetBillRuleEnable
    Call ShowRuleInfor
    If Not optRuleTotal(2).Visible Then
         If optRuleTotal(2).Value Then optRuleTotal(0).Value = True
    End If
End Sub
Private Sub SetBillRuleEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据票据分配规则,设置相应控件的Enabled属性
    '编制:刘兴洪
    '日期:2013-03-26 17:55:47
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnEnable As Boolean
    On Error GoTo errHandle
    '汇总条件(0-不汇总;1-首页汇总(按第1页汇总),2-分组汇总(选择明细时有效))
    '1.分组汇总:有费用明细时,同时存在勾选执行科室或者收据费目或按单据才会存在分组汇总项,才会存在汇总项
    blnEnable = chkBillRule(3).Enabled And chkBillRule(3).Value = 1 And (chkBillRule(2).Value = 1 Or chkBillRule(1).Value = 1 Or chkBillRule(0).Value = 1)
    optRuleTotal(2).Visible = blnEnable
    optRuleTotal(2).Enabled = blnEnable
    '2.首页汇总:都允许设置成汇总额
    optRuleTotal(1).Enabled = chkBillRule(3).Enabled
    optRuleTotal(0).Enabled = chkBillRule(3).Enabled
    
    '设置分组汇总相题
    If chkBillRule(0).Value = 1 Then
        optRuleTotal(2).Caption = "按单据号分组汇总"
    ElseIf chkBillRule(1).Value = 1 Then
        optRuleTotal(2).Caption = "按执行科室分组汇总"
    ElseIf chkBillRule(3).Value = 1 Then
        optRuleTotal(2).Caption = "按收据费目分组汇总"
    ElseIf chkBillRule(3).Value = 1 Then
        optRuleTotal(2).Caption = "按分组条件汇总"
    End If
    For intIndex = 1 To 3
        txtBillRuleNum(intIndex - 1).Enabled = chkBillRule(intIndex).Value = 1 And chkBillRule(intIndex).Enabled
        updBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
        lblBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Sub

Private Sub chkOnePatiPrint_Click()
  With vsBillFormat
        .ColHidden(.ColIndex("按病人补打票据格式")) = chkOnePatiPrint.Value <> 1
    End With

End Sub

Private Sub chk票据张数_Click()
    txt票据张数.Enabled = chk票据张数.Enabled And chk票据张数.Value = 1
    upd票据张数.Enabled = txt票据张数.Enabled
End Sub

Private Sub chk收费执行科室_Click()
    If mblnNotClick Then Exit Sub
    If chk收费执行科室.Value = vbChecked Then
        cmd收费执行科室.Enabled = True
        Call cmd收费执行科室_Click
    Else
        txt收费执行科室.Text = ""
        txt收费执行科室.Tag = ""
        cmd收费执行科室.Enabled = False
    End If
End Sub

Private Sub cmdAddedItem_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID, 编码, 名称, 计算单位, 说明" & vbNewLine & _
            "From 收费项目目录" & vbNewLine & _
            "Where 类别 = 'Z' And Nvl(是否变价, 0) = 0 And 服务对象 In(1,3)" & vbNewLine & _
            "Order By 编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "(定价)收费项目")
    If Not rsTmp Is Nothing Then
        txtAddedItem.Text = rsTmp!名称
        txtAddedItem.Tag = rsTmp!ID
        If chkAddedItem.Value = 0 Then chkAddedItem.Value = 1
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDeviceSetup_Click()
    Dim lngModule As Long
    Select Case mbytInFun
    Case 0
        lngModule = 1121
    Case 1
        lngModule = 1120
    Case 2
        lngModule = 1122
    End Select
    Call zlCommFun.DeviceSetup(Me, 100, lngModule)
End Sub
Private Sub cbo费别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo费别.ListIndex = -1
End Sub

Private Sub chkMulti_Click()
    If chkMulti.Value = 0 Then
        chkSeekBill.Value = 0
        chkSeekBill.Enabled = False
        chkUnPopPriceBill.Value = 0
        chkUnPopPriceBill.Enabled = False
        
        chkAutoSplitBill.Value = 0
        chkAutoSplitBill.Enabled = False
    Else
        chkSeekBill.Enabled = True And chkSeekBill.Tag = "1"
        chkUnPopPriceBill.Enabled = chkSeekBill.Value = 1 And chkUnPopPriceBill.Tag = "1"
        chkAutoSplitBill.Enabled = True And chkAutoSplitBill.Tag = "1"
    End If
    cboAutoSplitBill.Enabled = chkAutoSplitBill.Enabled And cboAutoSplitBill.Tag = "1"
End Sub

Private Sub chkSeekBill_Click()
    txtSeekDays.Enabled = chkSeekBill.Value = 1 And txtSeekDays.Tag = "1"
    If Visible And txtSeekDays.Enabled And txtSeekDays.Visible Then
        txtSeekDays.SetFocus
    End If
    chkUnPopPriceBill.Enabled = chkSeekBill.Value = 1 And chkUnPopPriceBill.Tag = "1"
    If chkSeekBill.Value = 0 Then chkUnPopPriceBill.Value = 0
End Sub

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
    chkOnlyUnitPatient.Enabled = chkSeekName.Value = 1 And chkOnlyUnitPatient.Tag = "1"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
End Sub
Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存发票相关票据
    '编制:刘兴洪
    '日期:2011-04-28 18:16:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String, strOnePatiPrintValue As String
    Dim i As Long
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
    zlDatabase.SetPara "共用收费票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    '保存收费格式
    
    Dim strPrintMode As String
    '保存收费格式
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("收费票据格式")))
                strOnePatiPrintValue = strOnePatiPrintValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("按病人补打票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("收费打印方式")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strOnePatiPrintValue <> "" Then strOnePatiPrintValue = Mid(strOnePatiPrintValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "收费发票格式", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "按病人补打发票格式", strOnePatiPrintValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "结帐发票格式", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "收费发票打印方式", strPrintMode, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "按病人补打发票不区分结算次数", IIf(chkOnePatiPrint.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    End With
    
    '保存退费格式
    strValue = "": strPrintMode = ""
    With vsDelBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("退费票据格式")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("退费打印方式")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "退费发票格式", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "退费发票打印方式", strPrintMode, glngSys, mlngModul, blnHavePrivs
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
    If cboBillRole.ListIndex = 1 Then
        If chkBillRule(0).Value = 0 And chkBillRule(1).Value = 0 And chkBillRule(2).Value = 0 And chkBillRule(3).Value = 0 Then
            MsgBox "注意:" & vbCrLf & "    票据号分配规则按『" & cboBillRole.Text & "』的必须设置一种分配规则,请检查!", vbInformation + vbOKOnly
            stab.Tab = 3
            If chkBillRule(0).Enabled And chkBillRule(0).Visible Then chkBillRule(0).SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim strValue As String, i As Long
    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    
    'a.数据检查
    '--------------------------------------------------------------
    'b.本机注册表存储的模块参数
    '------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If isValied = False Then Exit Sub
     
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    If Not mblnSetDrugStore Then
        For i = lst收费类别.ListCount - 1 To 0 Step -1
            If lst收费类别.Selected(i) Then strValue = strValue & "'" & Chr(lst收费类别.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "收费类别", strValue, glngSys, mlngModul, blnHavePrivs
        
        If mbytInFun <> 2 Then
            zlDatabase.SetPara "缺省费别", cbo费别.Text, glngSys, mlngModul, blnHavePrivs
        End If
        If mbytInFun = 0 Then
            Call SaveInvoice
            zlDatabase.SetPara "缺省结算方式", cbo结算方式.Text, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "挂号共用收费票据", chkRegistInvoice.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "手工报价", chkLed.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED显示收费明细", chkLedDispDetail.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "医保结算光标缺省定位", chk医保结算光标缺省定位.Value, glngSys, mlngModul, blnHavePrivs
        End If
        
        On Error Resume Next
        zlDatabase.SetPara "中药付数", chkPay.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "变价数次", chkTime.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "显示护士", chk护士.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "药品单位", IIf(opt单位(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    End If
    
    If gbln分离发药 Then
        strValue = ""
        For i = 0 To lst西药房.ListCount - 1
            If lst西药房.Selected(i) Then
                strValue = strValue & "," & lst西药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "西药房选择", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
        strValue = ""
        For i = 0 To lst成药房.ListCount - 1
            If lst成药房.Selected(i) Then
                strValue = strValue & "," & lst成药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "成药房选择", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
        strValue = ""
        For i = 0 To lst中药房.ListCount - 1
            If lst中药房.Selected(i) Then
                strValue = strValue & "," & lst中药房.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "中药房选择", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
    Else
        With vsfDrugStore
            For i = 1 To vsfDrugStore.Rows - 1
                If (mbytInFun = 0 Or mbytInFun = 1) And .TextMatrix(i, .ColIndex("窗口")) <> "自动分配" And .TextMatrix(i, .ColIndex("窗口")) <> "" Then
                    Select Case .TextMatrix(i, 0)
                        Case "西药房"
                            str西药房窗口 = str西药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                        Case "中药房"
                            str中药房窗口 = str中药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                        Case "成药房"
                            str成药房窗口 = str成药房窗口 & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("窗口"))
                    End Select
                End If
                
                If Abs(Val(.TextMatrix(i, .ColIndex("缺省")))) = 1 Then
                    Select Case .TextMatrix(i, .ColIndex("类别"))
                        Case "西药房"
                            lng缺省西药房 = .RowData(i)
                        Case "中药房"
                            lng缺省中药房 = .RowData(i)
                        Case "成药房"
                            lng缺省成药房 = .RowData(i)
                    End Select
                End If
            Next
        End With
        If cbo卫材.ListIndex <> -1 Then
            lng缺省发料部门 = cbo卫材.ItemData(cbo卫材.ListIndex)
        End If
        
        
        If mbytInFun = 0 Or mbytInFun = 1 Then
            str西药房窗口 = Mid(str西药房窗口, 2)
            str中药房窗口 = Mid(str中药房窗口, 2)
            str成药房窗口 = Mid(str成药房窗口, 2)
            zlDatabase.SetPara "西药房窗口", str西药房窗口, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "中药房窗口", str中药房窗口, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "成药房窗口", str成药房窗口, glngSys, mlngModul, blnHavePrivs
        End If
        
        zlDatabase.SetPara "缺省西药房", lng缺省西药房, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "缺省中药房", lng缺省中药房, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "缺省成药房", lng缺省成药房, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "缺省发料部门", lng缺省发料部门, glngSys, mlngModul, blnHavePrivs
                    
                    
        zlDatabase.SetPara "显示其它药房库存", chk药房.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "显示其它药库库存", chk药库.Value, glngSys, mlngModul, blnHavePrivs
        If mbytInFun <> 0 Then
            zlDatabase.SetPara "库存显示方式", IIf(opt库存(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
        End If
    End If
    
        
    If Not mblnSetDrugStore Then
        zlDatabase.SetPara "科室医生", IIf(optDoctor.Value, 0, IIf(optUnit.Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "开单人显示方式", IIf(optDoctorKind(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        
        zlDatabase.SetPara "姓名模糊查找", chkSeekName.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "姓名查找天数", Val(txtNameDays.Text), glngSys, mlngModul, blnHavePrivs
        '92727
        zlDatabase.SetPara "允许录入特殊使用的抗生素", IIf(chk抗生素.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            
        If mbytInFun = 0 Or mbytInFun = 1 Then
            zlDatabase.SetPara "病人来源", IIf(opt病人(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "最大金额", txtMax.Text, glngSys, mlngModul, blnHavePrivs
            
            zlDatabase.SetPara "性别", chk性别.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "年龄", chk年龄.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "费别", chk费别.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "医疗付款", chk医疗付款.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "加班", chk是否加班.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "开单日期", chk开单日期.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "开单人", chk开单人.Value, glngSys, mlngModul, blnHavePrivs
                    
            zlDatabase.SetPara "必须要输入开单人", chk必须输开单人.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "不使用缺省开单人", chk不缺省开单人.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "缺省科室优先", chk缺省科室优先.Value, glngSys, mlngModul, blnHavePrivs
            
            zlDatabase.SetPara "分类合计方式", IIf(opt分类(0).Value, 0, IIf(opt分类(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs '34179
            
            '刘兴洪 问题:27663 日期:2010-01-27 11:17:48
            zlDatabase.SetPara "住院病人按门诊收费", IIf(chk住院按门诊收费.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '刘兴洪 问题:39253
            zlDatabase.SetPara "提取划价后立即缴款", IIf(chk划价立即缴款.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '47457
            zlDatabase.SetPara "使用加减切换支付方式", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '47400
            zlDatabase.SetPara "药品摆药退费方式", IIf(optDrug(0).Value, 0, IIf(optDrug(1).Value, "1", "2")), glngSys, mlngModul, blnHavePrivs
            '87489
            zlDatabase.SetPara "退费缺省选择方式", IIf(opt退费缺省选择方式(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
            '86853
            zlDatabase.SetPara "刷卡缺省金额操作", IIf(optSetMoneyMode(0).Value, 0, IIf(optSetMoneyMode(1).Value, 1, 2)), glngSys, 1151, blnHavePrivs
            
            If mbytInFun = 0 Then
                zlDatabase.SetPara "显示累计", chk累计.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "检查皮试结果", chk皮试.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "优先使用预交款", chkPrePayPriority.Value, glngSys, mlngModul, blnHavePrivs
                '120836
                zlDatabase.SetPara "禁止取消挂号划价单", IIf(chk禁止取消挂号单.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
                '刘兴洪:22343:51670
                zlDatabase.SetPara "收费缴款输入控制", IIf(opt缴款(0).Value = True, 0, IIf(opt缴款(1).Value = True, 1, IIf(opt缴款(2).Value = True, 2, 3))), glngSys, mlngModul, blnHavePrivs
                '91665
                zlDatabase.SetPara "只对医保结算成功单据收费", chkInsurePartFee.Value, glngSys, mlngModul, blnHavePrivs
                '96357
                zlDatabase.SetPara "本机收费执行科室", txt收费执行科室.Tag, glngSys, mlngModul, blnHavePrivs
                
                If chkAddedItem.Value = 1 And Val(txtAddedItem.Tag) <> 0 Then
                    zlDatabase.SetPara "自动加收挂号费", txtAddedItem.Tag & ";" & txtAddedItem.Text, glngSys, mlngModul, blnHavePrivs
                Else
                    zlDatabase.SetPara "自动加收挂号费", "", glngSys, mlngModul, blnHavePrivs
                End If
'                zlDatabase.SetPara "显示误差费用", chkShowError.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "多单据收费", chkMulti.Value, glngSys, mlngModul, blnHavePrivs
                
                zlDatabase.SetPara "搜寻划价单据", chkSeekBill.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "搜寻单据天数", Val(txtSeekDays.Text), glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "不弹出划价单选择", chkUnPopPriceBill.Value, glngSys, mlngModul, blnHavePrivs
                    
                zlDatabase.SetPara "检查病人挂号科室", chkMustRegevent.Value, glngSys, mlngModul, blnHavePrivs
                For i = 0 To optRegPrompt.UBound
                    If optRegPrompt(i).Value Then
                        zlDatabase.SetPara "未挂号病人收费", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                zlDatabase.SetPara "自动组合单据", IIf(chkAutoSplitBill.Value = 1, cboAutoSplitBill.ListIndex + 1, 0), glngSys, mlngModul
               For i = 0 To optPrint.UBound
                    If optPrint(i).Value Then
                        zlDatabase.SetPara "收费清单打印方式", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                For i = 0 To optRefund.UBound
                    If optRefund(i).Value Then
                        zlDatabase.SetPara "退费回单打印方式", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                '62982:李南春,2015/08/25,收费执行单
                For i = 0 To optExe.UBound
                    If optExe(i).Value Then
                        zlDatabase.SetPara "收费执行单打印方式", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                
                '刘兴洪 问题:26948 日期:2009-12-28 16:54:11
                zlDatabase.SetPara "票据剩余X张时开始提醒收费员", IIf(chk票据张数.Value = 1, "1", "0") & "|" & Val(txt票据张数.Text), glngSys, mlngModul, blnHavePrivs
            
            Else
                zlDatabase.SetPara "取消划价单", Val(txtDay.Text), glngSys, mlngModul, blnHavePrivs
                For i = 0 To optPrintRequisition.UBound
                    If optPrintRequisition(i).Value Then
                        zlDatabase.SetPara "划价通知单打印方式", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
            End If
        ElseIf mbytInFun = 2 Then
            zlDatabase.SetPara "只查找合约单位病人", chkOnlyUnitPatient.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "记帐打印", chk(0).Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "划价打印", chk(1).Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "审核打印", chk(2).Value, glngSys, mlngModul, blnHavePrivs
        End If
    End If
    Call SaveBillRulePara '56963
    Call InitLocPar(Choose(mbytInFun + 1, 1121, 1120, 1122))     '主要是要重读存到本机注册表的参数,存在数据库的参数在保存时已重读
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '门诊医疗费收费
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_1", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me)
                End If
            End If
        Case 1 '门诊诊断证明
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me)
            End If
        Case 2 '门诊收费清单
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me)
            End If
        Case 3 '划价通知单
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me)
        Case 4 '医保回单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me)
        Case 5  '退费回单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me)
        '62982:李南春,2015/08/25,收费执行单
        Case 6  '收费执行单设置
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me)
        Case 7  '门诊医疗费收费(红票)
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_7", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_7", Me)
                End If
            End If
    End Select
End Sub

Private Sub SetStockCheck()
'功能:设置分离发药模式下检查指定药房的库存
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim str西药ID As String, str成药ID As String, str中药ID As String
    
    On Error GoTo errH
    
    For i = 0 To lst西药房.ListCount - 1
        If lst西药房.Selected(i) Then
            str西药ID = str西药ID & "," & lst西药房.ItemData(i)
        End If
    Next
    For i = 0 To lst成药房.ListCount - 1
        If lst成药房.Selected(i) Then
            str成药ID = str成药ID & "," & lst成药房.ItemData(i)
        End If
    Next
    For i = 0 To lst中药房.ListCount - 1
        If lst中药房.Selected(i) Then
            str中药ID = str中药ID & "," & lst中药房.ItemData(i)
        End If
    Next
    If str西药ID <> "" Then str西药ID = str西药ID & ","
    If str成药ID <> "" Then str成药ID = str成药ID & ","
    If str中药ID <> "" Then str中药ID = str中药ID & ","
    lst西药房.Clear: lst成药房.Clear: lst中药房.Clear
    
    Set rsTmp = GetDepartments("'西药房','成药房','中药房'", IIf(opt病人(0).Value, 1, 2) & ",3")
    If Not rsTmp.EOF Then
        rsTmp.Filter = "工作性质='西药房'"
        Do While Not rsTmp.EOF
            lst西药房.AddItem rsTmp!名称
            lst西药房.ItemData(lst西药房.ListCount - 1) = rsTmp!ID
            If InStr(str西药ID, "," & rsTmp!ID & ",") > 0 Then lst西药房.Selected(lst西药房.NewIndex) = True
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "工作性质='成药房'"
        Do While Not rsTmp.EOF
            lst成药房.AddItem rsTmp!名称
            lst成药房.ItemData(lst成药房.ListCount - 1) = rsTmp!ID
            If InStr(str成药ID, "," & rsTmp!ID & ",") > 0 Then lst成药房.Selected(lst成药房.NewIndex) = True
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "工作性质='中药房'"
        Do While Not rsTmp.EOF
            lst中药房.AddItem rsTmp!名称
            lst中药房.ItemData(lst中药房.ListCount - 1) = rsTmp!ID
            If InStr(str中药ID, "," & rsTmp!ID & ",") > 0 Then lst中药房.Selected(lst中药房.NewIndex) = True
            rsTmp.MoveNext
        Loop
    End If
    
    If lst西药房.ListCount > 0 Then lst西药房.ListIndex = 0
    If lst成药房.ListCount > 0 Then lst成药房.ListIndex = 0
    If lst中药房.ListCount > 0 Then lst中药房.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    With vsfDrugStore
        strTmp = "'西药房','中药房','成药房','发料部门'"
        
        If stab.TabVisible(1) = True Then
            lngType = IIf(opt病人(0).Value, 1, 2)
        Else
            lngType = gint病人来源
        End If
        Set rsTmp = GetDepartments(strTmp, lngType & ",3")
        .Rows = 1
        If mbytInFun = 2 Then .ColHidden(3) = True '门诊记帐不设窗口
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "工作性质<>'发料部门'"
            .Rows = rsTmp.RecordCount + 1
            .MergeCells = flexMergeFixedOnly
            .MergeCol(0) = True
            
            strTmp = "'西药房','中药房','成药房'"
            arrTmp = Split(strTmp, ",")
            lngRow = 1
            For j = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质=" & arrTmp(j)
                If rsTmp.RecordCount > 0 Then
                    For i = 1 To rsTmp.RecordCount
                        .TextMatrix(lngRow, 0) = Replace(arrTmp(j), "'", "")
                        .TextMatrix(lngRow, 1) = 0
                        .TextMatrix(lngRow, 2) = rsTmp!名称
                        If mbytInFun <> 2 Then .TextMatrix(lngRow, 3) = "自动分配"
                        .RowData(lngRow) = Val(rsTmp!ID)
                        lngRow = lngRow + 1
                        rsTmp.MoveNext
                    Next
                    
                    If lngRow < .Rows - 1 Then  '划分隔线
                        .Select lngRow, .FixedCols, lngRow, .COLS - 1
                        .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            Next
            
            cbo卫材.AddItem "人工选择"
            rsTmp.Filter = "工作性质='发料部门'"
            For j = 1 To rsTmp.RecordCount
                cbo卫材.AddItem rsTmp!名称
                cbo卫材.ItemData(cbo卫材.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            cbo卫材.ListIndex = 0
        End If
    
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd收费执行科室_Click()
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    '96357
    strSQL = "Select Distinct A.ID, A.编码, A.名称, A.简码" & vbNewLine & _
            " From 部门表 A, 部门性质说明 B" & vbNewLine & _
            " Where B.部门ID=A.ID And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & vbNewLine & _
            "       And B.工作性质 In('中药房', '西药房', '成药房', '发料部门')" & vbNewLine & _
            "       And B.服务对象 In (1, 2, 3)" & vbNewLine & _
            "       And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.编码"
    vRect = GetControlRect(txt收费执行科室.hWnd)
    Set rsDept = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "本机收费执行科室", True, "", "", False, False, False, _
        vRect.Left, vRect.Top, txt收费执行科室.Height, blnCancel, False, True, "MultiCheckReturn=1")
    If blnCancel Then Exit Sub
    If rsDept Is Nothing Then Exit Sub
    
    txt收费执行科室.Text = ""
    txt收费执行科室.Tag = ""
    Do While Not rsDept.EOF
        txt收费执行科室.Text = txt收费执行科室.Text & ";" & Nvl(rsDept!名称)
        strTemp = strTemp & "," & Nvl(rsDept!ID)
        rsDept.MoveNext
    Loop
    If txt收费执行科室.Text <> "" Then txt收费执行科室.Text = Mid(txt收费执行科室.Text, 2)
    If strTemp <> "" Then txt收费执行科室.Tag = Mid(strTemp, 2)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
        
    Select Case mbytInFun
    Case 0
    Case 1
        Me.Height = 5325: stab.Height = 4695
        cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 100
        cbo卫材.Top = stab.Height + stab.Top - cbo卫材.Height - 200
        lbl发料部门.Top = cbo卫材.Top + (cbo卫材.Height - lbl发料部门.Height) \ 2
        vsfDrugStore.Height = cbo卫材.Top - vsfDrugStore.Top - 50
    Case 2
        Me.Height = 6025 + IIf(chk住院按门诊收费.Visible, chk住院按门诊收费.Height + 20, 0)
        stab.Height = 6055 + IIf(chk住院按门诊收费.Visible, chk住院按门诊收费.Height + 20, 0)
        Me.cmdHelp.Top = 5095
        cbo卫材.Top = stab.Height + stab.Top - cbo卫材.Height - 200
        lbl发料部门.Top = cbo卫材.Top + (cbo卫材.Height - lbl发料部门.Height) \ 2
        vsfDrugStore.Height = cbo卫材.Top - vsfDrugStore.Top - 50
    End Select
End Sub
Private Sub MoveCtrol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的位置
    '编制:刘兴洪
    '日期:2011-09-12 13:46:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ''0=收费,1=划价,2=门诊记帐

    Select Case mbytInFun
    Case 2   '门诊记帐
        '显示为最顶部的一行
        chkPay.Top = chkPay.Top
        chkPay.Left = fra单位.Left
        chkTime.Left = chkPay.Left + chkPay.Width + 50
        chk护士.Left = chkTime.Left + chkTime.Width + 50
        chkTime.Top = chkPay.Top
        chk护士.Top = chkPay.Top
        '自动显示,包括Frame中的控件,虽然不在当前卡页上
        fra科室与医生.Left = fra单位.Left
        fra科室与医生.Top = fra单位.Top - fra科室与医生.Height - 100
        optUnit.Top = optUnit.Top + 30
        optDoctor.Top = optUnit.Top + optUnit.Height + 80
        optSelf.Top = optDoctor.Top + optDoctor.Height + 80
        
        fraPrintBill.Height = fraPrintBill.Height - 50
        fra科室与医生.Height = fraPrintBill.Height
        
        fra单位.Top = fra单位.Top + 100
        fraPrintBill.Left = fra单位.Left + fra单位.Width - fraPrintBill.Width
        fraPrintBill.Top = fra科室与医生.Top ' fra单位.Top - fraPrintBill.Height - 100
        
        chkSeekName.Top = chkPay.Top + chkPay.Height + 80
        chkSeekName.Left = chkPay.Left
        txtNameDays.Top = chkSeekName.Top
        txtNameDays.Left = chkSeekName.Left + chkSeekName.Width * 14 / 23 + 50
        fraShortLine.Top = txtNameDays.Top + txtNameDays.Height
        fraShortLine.Left = txtNameDays.Left
        
        chkOnlyUnitPatient.Top = chkSeekName.Top + chkSeekName.Height + 50
        chkOnlyUnitPatient.Left = chkPay.Left
        
       ' lst收费类别.Height = lst收费类别.Height - 80
        cmdDeviceSetup.Left = fra类别.Left
        cmdDeviceSetup.Visible = True
        cmdDeviceSetup.Top = cmdDeviceSetup.Top - 150
        'fraDoctor.Top = cmdDeviceSetup.Top - fraDoctor.Height - 100
        'fraDoctor.Left = fra类别.Left
        
        fra库存显示.Top = fra库存显示.Top + 100
        fraDoctor.Left = fra库存显示.Left
        fraDoctor.Top = fra库存显示.Top + fra库存显示.Height + 50
        fraDoctor.Width = fra库存显示.Width
        optDoctorKind(0).Caption = "按简码+姓名显示"
        optDoctorKind(1).Caption = "按编码+姓名显示"
        
        optDoctorKind(0).Width = Me.TextWidth(optDoctorKind(0).Caption) + 400
        optDoctorKind(1).Width = Me.TextWidth(optDoctorKind(1).Caption) + 400
        
        
        optDoctorKind(0).Top = optDoctorKind(0).Top + 40
        optDoctorKind(0).Left = fraDoctor.Left + (fraDoctor.Width - optDoctorKind(0).Width - optDoctorKind(1).Width - 1000) \ 2
        
        optDoctorKind(1).Left = optDoctorKind(0).Left + optDoctorKind(0).Width + 50
        optDoctorKind(1).Top = optDoctorKind(0).Top
        
        fraDoctor.Height = fraDoctor.Height - 250
        chk抗生素.Top = fraDoctor.Top + fraDoctor.Height + 50: chk抗生素.Left = chk划价立即缴款.Left
        'fra类别.Height = fraDoctor.Top - fra类别.Top - 100
        fra类别.Height = fra类别.Height - 80
        lst收费类别.Height = fra类别.Height - 300
    Case 1 '划价单
        txtDay.Top = cbo结算方式.Top
        udDay.Top = txtDay.Top
        lblDay.Top = txtDay.Top + (txtDay.Height - lblDay.Height) \ 2
        fra单位.Top = txtDay.Top + txtDay.Height + 200
        fra库存显示.Top = fra单位.Top + fra单位.Height + 200
        chk抗生素.Top = fra库存显示.Top + fra库存显示.Height + 50: chk抗生素.Left = chk划价立即缴款.Left
        
        fraDoctor.Top = fraRegPrompt.Top
        fraDoctor.Height = fraRegPrompt.Height
        fraDoctor.Left = fraRegPrompt.Left
        fraDoctor.Width = fraRegPrompt.Width
        optDoctorKind(0).Top = optDoctorKind(0).Top + 80
        optDoctorKind(1).Top = optDoctorKind(0).Top + optDoctorKind(0).Height + 50
        fraInputItem.Width = fraRegPrompt.Left + fraRegPrompt.Width - fraInputItem.Left
        chk性别.Left = chk性别.Left + 100
        chk开单日期.Left = chk性别.Left
        
        chk是否加班.Left = chk开单日期.Left + chk开单日期.Width + 800
        chk开单人.Left = chk是否加班.Left
        
        chk年龄.Left = chk是否加班.Left + chk是否加班.Width + 800
        chk医疗付款.Left = chk年龄.Left
        chk费别.Left = chk医疗付款.Left + chk医疗付款.Width - chk费别.Width
        fra划价通知单打印.Top = cmdPrintSetup(3).Top
        cmdPrintSetup(3).Top = fra划价通知单打印.Top + fra划价通知单打印.Height + 100
        cmdDeviceSetup.Left = fraInputItem.Left
        cmdDeviceSetup.Top = cmdPrintSetup(3).Top
           
    Case Else   '收费
        fra类别.Top = fra类别.Top
        txtMax.Top = fra类别.Top
        lblMax.Top = txtMax.Top + (txtMax.Height - lblMax.Height) \ 2
        chkPay.Top = lblMax.Top
        chkTime.Top = chkPay.Top + chkPay.Height + IIf(gbln分离发药, 50, 100)
        chk护士.Top = chkTime.Top + chkTime.Height + IIf(gbln分离发药, 50, 100)
        chk累计.Top = chk护士.Top + chk护士.Height + IIf(gbln分离发药, 50, 100)
        
        cbo费别.Top = txtMax.Top + txtMax.Height + IIf(gbln分离发药, 50, 100)
        lbl费别.Top = cbo费别.Top + (cbo费别.Height - lbl费别.Height) \ 2
        
        
        cbo结算方式.Top = cbo费别.Top + cbo费别.Height + IIf(gbln分离发药, 50, 100)
        lbl结算方式.Top = cbo结算方式.Top + (cbo结算方式.Height - lbl结算方式.Height) \ 2
        
        chk皮试.Top = cbo结算方式.Height + cbo结算方式.Top + IIf(gbln分离发药, 50, 100)
        chkPrePayPriority.Top = chk皮试.Top + chk皮试.Height + IIf(gbln分离发药, 50, 100)
        
        txtAddedItem.Top = chkPrePayPriority.Top + chkPrePayPriority.Height + IIf(gbln分离发药, 0, 100)
        cmdAddedItem.Top = txtAddedItem.Top
        chkAddedItem.Top = txtAddedItem.Top + (txtAddedItem.Height - chkAddedItem.Height) \ 2
        
        chkInsurePartFee.Top = txtAddedItem.Top + txtAddedItem.Height + IIf(gbln分离发药, 50, 100)
        
        fra单位.Top = chkInsurePartFee.Top + chkInsurePartFee.Height + IIf(gbln分离发药, 0, 100)
        fra单位.Height = fra单位.Height + IIf(gbln分离发药, 0, 100)
        opt单位(0).Top = opt单位(0).Top + IIf(gbln分离发药, 0, 50)
        opt单位(1).Top = opt单位(0).Top
        lbl单位.Top = opt单位(0).Top
        fra库存显示.Height = 700
        fra库存显示.Top = fra单位.Top + fra单位.Height + IIf(gbln分离发药, 50, 100)
        fra分离.Top = fra库存显示.Top
        
        If Not gbln分离发药 Then
            chk住院按门诊收费.Top = fra库存显示.Top + fra库存显示.Height + 50
        Else
            chk住院按门诊收费.Top = fra分离.Top + fra分离.Height + 20
        End If
        chk医保结算光标缺省定位.Top = chk住院按门诊收费.Top
        chk划价立即缴款.Top = chk住院按门诊收费.Top + chk住院按门诊收费.Height + 50
        chkPayKey.Top = chk划价立即缴款.Top
        chk禁止取消挂号单.Top = chk划价立即缴款.Top + chk划价立即缴款.Height + 50
        chk抗生素.Top = chk禁止取消挂号单.Top: chk抗生素.Left = chkPayKey.Left
        fraDrugNotFee.Top = chk抗生素.Top + chk抗生素.Height + 50
        fra退费缺省选择方式.Top = fraDrugNotFee.Top + fraDrugNotFee.Height + 50
        fraSetMoneyMode.Top = fra退费缺省选择方式.Top + fra退费缺省选择方式.Height + 50
        
        '第二页控制
        chk必须输开单人.Top = chk缺省科室优先.Top
        chk必须输开单人.Left = chkLed.Left
        opt分类(0).Top = chkOnlyUnitPatient.Top
        opt分类(1).Top = opt分类(0).Top + opt分类(0).Height + 20
        opt分类(2).Top = opt分类(1).Top + opt分类(1).Height + 20
        
        chkLed.Top = chk必须输开单人.Top + chk必须输开单人.Height + 20
        chkLedDispDetail.Top = chkLed.Top + chkLed.Height + 20
        chkLedWelcome.Top = chkLedDispDetail.Top + chkLedDispDetail.Height + 20
        chkUnPopPriceBill.Top = chkLedWelcome.Top + chkLedWelcome.Height + 20
        chkMustRegevent.Top = chkUnPopPriceBill.Top + chkUnPopPriceBill.Height + 20
        chkMustRegevent.Left = chkUnPopPriceBill.Left
        cmdDeviceSetup.Top = chkMustRegevent.Top + chkMustRegevent.Height + 100
        
'        chkShowError.Top = opt分类(2).Top + opt分类(2).Height + 20
'        chkMulti.Top = chkShowError.Top + chkShowError.Height + 20
        chkMulti.Top = opt分类(2).Top + opt分类(2).Height + 20
        chkSeekBill.Top = chkMulti.Top + chkMulti.Height + 20
        txtSeekDays.Top = chkSeekBill.Top
        fraLine.Top = txtSeekDays.Top + txtSeekDays.Height
        
        cboAutoSplitBill.Top = chkSeekBill.Top + chkSeekBill.Height + 20
        chkAutoSplitBill.Top = cboAutoSplitBill.Top + (cboAutoSplitBill.Height - chkAutoSplitBill.Height) \ 2
        fra缴款控制.Top = IIf(cboAutoSplitBill.Top + cboAutoSplitBill.Height + 20 > cmdDeviceSetup.Top + cmdDeviceSetup.Height + 20, cboAutoSplitBill.Top + cboAutoSplitBill.Height + 20, cmdDeviceSetup.Top + cmdDeviceSetup.Height + 20)
        
        txt收费执行科室.Top = fra缴款控制.Top + fra缴款控制.Height + 100
        chk收费执行科室.Top = txt收费执行科室.Top + (txt收费执行科室.Height - chk收费执行科室.Height) / 2
        cmd收费执行科室.Top = txt收费执行科室.Top
        
        fra类别.Height = fra类别.Height - 50
        lst收费类别.Height = lst收费类别.Height + 100
    End Select
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的隐藏属性
    '编制:刘兴洪
    '日期:2011-09-12 14:55:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln分离发药 Then
        stab.TabVisible(4) = False: fra库存显示.Visible = False
        fra分离.Visible = True
    End If
    '刘兴洪 问题:27663 日期:2010-01-27 13:29:19
    chk住院按门诊收费.Visible = mbytInFun = 0
    chk划价立即缴款.Visible = mbytInFun = 0
    '47457
    chkPayKey.Visible = mbytInFun = 0
    '87489
    fra退费缺省选择方式.Visible = mbytInFun = 0
    fraSetMoneyMode.Visible = mbytInFun = 0
End Sub

Private Sub Load药房ParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载药房相关参数值
    '编制:刘兴洪
    '日期:2011-12-07 15:05:10
    '问题:43775
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean
    Dim i As Long, k As Long, j As Long, intType As Integer
    Dim arrTmp  As Variant, arrWindow As Variant
    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0
    If gbln分离发药 = True Then
        strTmp = zlDatabase.GetPara("西药房选择", glngSys, mlngModul, , Array(lst西药房), blnParSet)
        For i = 0 To lst西药房.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst西药房.ItemData(i) & ",") > 0 Then
                lst西药房.Selected(i) = True
            End If
        Next
        strTmp = zlDatabase.GetPara("成药房选择", glngSys, mlngModul, , Array(lst成药房), blnParSet)
        For i = 0 To lst成药房.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst成药房.ItemData(i) & ",") > 0 Then
                lst成药房.Selected(i) = True
            End If
        Next
        strTmp = zlDatabase.GetPara("中药房选择", glngSys, mlngModul, , Array(lst中药房), blnParSet)
        For i = 0 To lst中药房.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst中药房.ItemData(i) & ",") > 0 Then
                lst中药房.Selected(i) = True
            End If
        Next
        If lst西药房.ListCount > 0 Then lst西药房.ListIndex = 0
        If lst成药房.ListCount > 0 Then lst成药房.ListIndex = 0
        If lst中药房.ListCount > 0 Then lst中药房.ListIndex = 0
        Exit Sub
    End If
    
    With vsfDrugStore
        arrTmp = Split("缺省西药房,缺省中药房,缺省成药房", ",")
        .Cell(flexcpData, 0, 0, .Rows - 1, .COLS - 1) = "0" '存储是否允许编译.:0-不锁定,1-锁定
        
        For j = 0 To UBound(arrTmp)
            '刘兴洪:由于可能参数权限发生变更,因此,不能统一进行设置,需要设置某一部分:
            '问题:25132,intType-'返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
            strTmp = zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, "0", , blnParSet, intType)
            If Val(strTmp) > 0 Then
                Select Case arrTmp(j)
                    Case "缺省西药房"
                        lng缺省西药房 = Val(strTmp)
                    Case "缺省中药房"
                        lng缺省中药房 = Val(strTmp)
                    Case "缺省成药房"
                        lng缺省成药房 = Val(strTmp)
                End Select
                Call SetDrugStockEdit(Replace(arrTmp(j), "缺省", ""), intType, .ColIndex("缺省"), Val(strTmp))
            Else
                Call SetDrugStockEdit(Replace(arrTmp(j), "缺省", ""), intType, .ColIndex("缺省"), "")
            End If
        Next
        
        strTmp = zlDatabase.GetPara("缺省发料部门", glngSys, mlngModul, "0", Array(cbo卫材), blnParSet)
        zlControl.CboLocate cbo卫材, strTmp, True
        
        If mbytInFun <> 2 Then
                arrTmp = Split("西药房窗口,中药房窗口,成药房窗口", ",")
                For j = 0 To UBound(arrTmp)
                    strTmp = Trim(zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, , , blnParSet, intType))
                    If strTmp <> "" Then
                        '处理旧的数据,窗口参数中没有存储药房ID
                        If InStr(strTmp, ":") = 0 Then
                            Select Case arrTmp(j)
                                Case "西药房窗口"
                                    strTmp = lng缺省西药房 & ":" & strTmp
                                Case "中药房窗口"
                                    strTmp = lng缺省中药房 & ":" & strTmp
                                Case "成药房窗口"
                                    strTmp = lng缺省成药房 & ":" & strTmp
                            End Select
                        End If
                        arrWindow = Split(strTmp, ",")
                        strTmp = Replace(arrTmp(j), "窗口", "")
                        For k = 0 To UBound(arrWindow)
                            Call SetDrugStockEdit(Replace(arrTmp(j), "窗口", ""), intType, .ColIndex("窗口"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
                        Next
                    Else
                        Call SetDrugStockEdit(Replace(arrTmp(j), "窗口", ""), intType, .ColIndex("窗口"), "")
                    End If
                Next
            End If
        End With
        chk药房.Value = IIf(zlDatabase.GetPara("显示其它药房库存", glngSys, mlngModul, , Array(chk药房), blnParSet) = "1", 1, 0)
        chk药库.Value = IIf(zlDatabase.GetPara("显示其它药库库存", glngSys, mlngModul, , Array(chk药库), blnParSet) = "1", 1, 0)
        If mbytInFun <> 0 Then
            If Val(Val(zlDatabase.GetPara("库存显示方式", glngSys, mlngModul, , Array(opt库存(0), opt库存(1)), blnParSet))) = 0 Then
                opt库存(0).Value = True
            Else
                opt库存(1).Value = True
            End If
            If opt库存(0).Enabled = False Then opt库存(0).Tag = "1"
        End If
     '   Call chk药房_Click
End Sub
Private Sub LoadParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数值
    '编制:刘兴洪
    '日期:2011-09-12 15:03:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean, k As Long, rsTmp As ADODB.Recordset
    Dim i As Long, arrTmp As Variant, j As Long, intType As Integer, arrWindow As Variant
        
    
    blnParSet = InStr(1, mstrPrivs, ";参数设置;") > 0

    strTmp = zlDatabase.GetPara("收费类别", glngSys, mlngModul, , Array(lst收费类别), blnParSet)
    If strTmp = "" Then
        For i = 0 To lst收费类别.ListCount - 1
            lst收费类别.Selected(i) = True
        Next
    Else
        For i = 0 To lst收费类别.ListCount - 1
            If InStr(strTmp, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
        Next
    End If
    If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
    If mbytInFun <> 2 Then
        strTmp = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , Array(cbo费别), blnParSet)
        zlControl.CboLocate cbo费别, strTmp
    End If
    chkPay.Value = IIf(zlDatabase.GetPara("中药付数", glngSys, mlngModul, , Array(chkPay), blnParSet) = "1", 1, 0)
    chkTime.Value = IIf(zlDatabase.GetPara("变价数次", glngSys, mlngModul, , Array(chkTime), blnParSet) = "1", 1, 0)
    chk护士.Value = IIf(zlDatabase.GetPara("显示护士", glngSys, mlngModul, , Array(chk护士), blnParSet) = "1", 1, 0)
    i = IIf(zlDatabase.GetPara("药品单位", glngSys, mlngModul, , Array(opt单位(0), opt单位(1)), blnParSet) = "0", 0, 1)
    opt单位(i).Value = True
    If mbytInFun = 0 Or mbytInFun = 1 Then
        i = IIf(zlDatabase.GetPara("病人来源", glngSys, mlngModul, , Array(opt病人(0), opt病人(1)), blnParSet) = "1", 0, 1)
        opt病人(i).Value = True
    End If
    chk抗生素.Value = IIf(Val(zlDatabase.GetPara("允许录入特殊使用的抗生素", glngSys, mlngModul, "0", Array(chk抗生素), blnParSet)) = 1, 1, 0)
    
    Call opt病人_Click(IIf(opt病人(0).Value, 0, 1)) '加载药品库房和卫材发料部门
    Call Load药房ParaValue
    Select Case mbytInFun
    Case 2 '记帐
    Case 1 '划价
    Case Else
        opt库存(0).Visible = False: opt库存(1).Visible = False: lbl库存显示方式.Visible = False
        lnSplit(0).Visible = False: lnSplit(1).Visible = False
        chkRegistInvoice.Value = IIf(zlDatabase.GetPara("挂号共用收费票据", glngSys, mlngModul, 0, Array(chkRegistInvoice), blnParSet) = "1", 1, 0)
        chkLed.Value = IIf(zlDatabase.GetPara("手工报价", glngSys, mlngModul, 0, Array(chkLed), blnParSet) = "1", 1, 0)
        chkLedDispDetail.Value = IIf(zlDatabase.GetPara("LED显示收费明细", glngSys, mlngModul, 1, Array(chkLedDispDetail), blnParSet) = "1", 1, 0)
        chkLedWelcome.Value = IIf(zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), blnParSet) = "1", 1, 0)
        Set rsTmp = Get结算方式("收费", "1,2,7")
        For i = 1 To rsTmp.RecordCount
            cbo结算方式.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo结算方式.ListIndex = cbo结算方式.NewIndex
            rsTmp.MoveNext
        Next
        '问题:54923
        strTmp = zlDatabase.GetPara("缺省结算方式", glngSys, mlngModul, , Array(cbo结算方式), blnParSet)
        For i = 0 To cbo结算方式.ListCount - 1
            If cbo结算方式.List(i) = strTmp Then cbo结算方式.ListIndex = i: Exit For
        Next
        '加载发票相关
        Call InitShareInvoice
         '39253
        chk划价立即缴款.Value = IIf(Val(zlDatabase.GetPara("提取划价后立即缴款", glngSys, mlngModul, "0", Array(chk划价立即缴款), blnParSet)) = 1, 1, 0)
        chk住院按门诊收费.Value = IIf(Val(zlDatabase.GetPara("住院病人按门诊收费", glngSys, mlngModul, "0", Array(chk住院按门诊收费), blnParSet)) = 1, 1, 0)
        '120836
        chk禁止取消挂号单.Value = IIf(Val(zlDatabase.GetPara("禁止取消挂号划价单", glngSys, mlngModul, "0", Array(chk禁止取消挂号单), blnParSet)) = 1, 1, 0)
        '47457
        chkPayKey.Value = IIf(Val(zlDatabase.GetPara("使用加减切换支付方式", glngSys, mlngModul, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
        '47400
        strTmp = zlDatabase.GetPara("药品摆药退费方式", glngSys, mlngModul, , Array(optDrug(0), optDrug(1), optDrug(2)), blnParSet)
        For i = 0 To 2
            If Val(strTmp) = i Then
                optDrug(i).Value = True: Exit For
            End If
        Next
        '87489
        strTmp = zlDatabase.GetPara("退费缺省选择方式", glngSys, mlngModul, "0", Array(opt退费缺省选择方式(0), opt退费缺省选择方式(1)), blnParSet)
        For i = 0 To 1
            If Val(strTmp) = i Then opt退费缺省选择方式(i).Value = True: Exit For
        Next
        chk医保结算光标缺省定位.Value = IIf(zlDatabase.GetPara("医保结算光标缺省定位", glngSys, mlngModul, "0", Array(chk医保结算光标缺省定位), blnParSet) = "1", 1, 0)
        '86853
        i = Val(zlDatabase.GetPara("刷卡缺省金额操作", glngSys, 1151, "0", Array(optSetMoneyMode(0), optSetMoneyMode(1), optSetMoneyMode(2)), blnParSet))
        If i < 0 Or i > optSetMoneyMode.UBound Then i = 0
        optSetMoneyMode(i).Value = True
        
        '56963:票号分配规则
        chkAutoAddBookFee.Value = IIf(Val(zlDatabase.GetPara("收据加收工本费", glngSys, mlngModul, "0", Array(chkAutoAddBookFee), blnParSet)) = 1, 1, 0)
        chkErrorItemNotBill.Value = IIf(Val(zlDatabase.GetPara("误差项不使用票据", glngSys, mlngModul, "0", Array(chkErrorItemNotBill), blnParSet)) = 1, 1, 0)

         
         '56963:2.根据预定规则分配票号
         strTmp = Trim(zlDatabase.GetPara("票据分配规则", glngSys, mlngModul, "0||0;0;0;0;0;0", _
         Array(cboBillRole, lblBillRole, chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
         lblBillRuleNum(0), updBillRuleNum(0), txtBillRuleNum(0), lblBillRuleNum(1), updBillRuleNum(1), lblBillRuleNum(2), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2)), blnParSet))
         arrTmp = Split(strTmp & "||", "||")
         
         '避免被更改
         optRuleTotal(0).Tag = IIf(optRuleTotal(0).Enabled, 1, 0)
        With cboBillRole
            .Clear
            .AddItem "1-根据实际打印分配票号"
            If Val(arrTmp(0)) = 0 Then .ListIndex = .NewIndex
            .AddItem "2-根据预定规则分配票号"
            If Val(arrTmp(0)) = 1 Then .ListIndex = .NewIndex
            .AddItem "3-根据自定义规则分配票号"
            If Val(arrTmp(0)) = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
            .Tag = .ListIndex   '记录修改前的选择
            '56963:存在打印数据时,不允许更改票号分配规则
            .Enabled = .Enabled And Not GetPrintListHaveData
        End With
        '2.根据预定规则分配票号
        arrTmp = Split(arrTmp(1) & ";;;", ";")
        '按单据分
        i = Val(arrTmp(0))
        chkBillRule(0).Value = IIf(i = 1, 1, 0)
        '按执行科室分
        i = Val(arrTmp(1))
        chkBillRule(1).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(0).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(0).Text = updBillRuleNum(0).Value
        txtBillRuleNum(0).Tag = IIf(updBillRuleNum(0).Value = 0, 1, updBillRuleNum(0).Value)
        
        '按收据费目
        i = Val(arrTmp(2))
        chkBillRule(2).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(1).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(1).Text = updBillRuleNum(1).Value
        txtBillRuleNum(1).Tag = IIf(updBillRuleNum(1).Value = 0, 1, updBillRuleNum(1).Value)
        '按收费细目(先处理收费细目，不然会触发Click事件，将首页汇总执行为空了
        i = Val(arrTmp(3))
        chkBillRule(3).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(2).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(2).Text = updBillRuleNum(2).Value
        txtBillRuleNum(2).Tag = IIf(updBillRuleNum(2).Value = 0, 20, updBillRuleNum(2).Value)
        
        '分组汇总
        i = Val(arrTmp(4)): i = IIf(i > 3 Or i < 0, 0, i)
        optRuleTotal(i).Value = True
         
        '1.根据实际打印分配票号
        chkBillNO.Value = IIf(Val(zlDatabase.GetPara("多张单据收费分别打印", glngSys, mlngModul, "0", Array(chkBillNO), blnParSet)) = 1, 1, 0)
        chk体检按单据分别打印.Value = IIf(Val(zlDatabase.GetPara("体检病人分单据打印", glngSys, mlngModul, "0", Array(chk体检按单据分别打印), blnParSet)) = 1, 1, 0)
        chk体检按单据分别打印.Enabled = (chkBillNO.Value = vbChecked)
        chkOlnyOneBill.Value = IIf(Val(zlDatabase.GetPara("收费每次只用一张票据", glngSys, mlngModul, "0", Array(chkOlnyOneBill), blnParSet)) = 1, 1, 0)
        i = Val(zlDatabase.GetPara("收费收据总行次", glngSys, mlngModul, "3", Array(lblRows, updRows, txtRowsUD), blnParSet))
        updRows.Value = IIf(i < 0 Or i > 100, 3, i)
        i = Val(zlDatabase.GetPara("收费票据生成方式", glngSys, mlngModul, "0", Array(optBillMode(0), optBillMode(1), chkExcuteDept), blnParSet))
        chkExcuteDept.Value = IIf(i >= 10, 1, 0)
        optBillMode(i Mod 10).Value = True
        '3-根据自定义规则分配票号
        Call SetBillRuleEnable
        Call ShowRuleInfor
    End Select
End Sub
Private Sub SaveBillRulePara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存票据分配规则的相关参数
    '编制:刘兴洪
    '日期:2013-03-26 16:32:46
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strTemp As String
    Dim intBillRull As Integer
    
    On Error GoTo errHandle
    If mbytInFun <> 0 Then Exit Sub
        
    If Val(cboBillRole.Tag) <> cboBillRole.ListIndex And cboBillRole.ListIndex <> 0 And Val(cboBillRole.Tag) <= 0 Then
       '如果当前切换成新模式,需要将票据打印格式记录下来,以便在重打或部分退费时按切换前的票据格式打印
       Call zlDatabase.ExecuteProcedure("Zl_Update_Bill_Printformat(" & glngSys & ")", Me.Caption)
    End If
    '只适合收费
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    intBillRull = IIf(cboBillRole.ListIndex < 0, 0, cboBillRole.ListIndex)
    strTemp = intBillRull & "||"
    '分单据
    strTemp = strTemp & IIf(chkBillRule(0).Value = 1, 1, 0)
    '执行科室
    strTemp = strTemp & ";" & IIf(chkBillRule(1).Value = 1, Val(txtBillRuleNum(0).Text), 0)
    '收据费目
    strTemp = strTemp & ";" & IIf(chkBillRule(2).Value = 1, Val(txtBillRuleNum(1).Text), 0)
    '收费细目
    strTemp = strTemp & ";" & IIf(chkBillRule(3).Value = 1, Val(txtBillRuleNum(2).Text), 0)
    '汇总条件
    strTemp = strTemp & ";" & IIf(optRuleTotal(0).Value, 0, IIf(optRuleTotal(1).Value, 1, 2))
    
    zlDatabase.SetPara "票据分配规则", strTemp, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "收据加收工本费", IIf(chkAutoAddBookFee.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "误差项不使用票据", IIf(chkErrorItemNotBill.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    If intBillRull = 0 Then
        '根据实际打印分配票据
        zlDatabase.SetPara "收费每次只用一张票据", IIf(chkOlnyOneBill.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "多张单据收费分别打印", IIf(chkBillNO.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "体检病人分单据打印", IIf(chk体检按单据分别打印.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        strTemp = CStr(IIf(optBillMode(1).Value, 1, 0) + Val(chkExcuteDept.Value) * 10)
        zlDatabase.SetPara "收费票据生成方式", strTemp, glngSys, mlngModul, blnHavePrivs
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Private Sub SetBillNoRule()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置票据分配规则的位置
    '编制:刘兴洪
    '日期:2013-03-26 15:43:12
    '问题:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intIndex As Integer
    On Error GoTo errHandle
    intIndex = cboBillRole.ListIndex
    If intIndex < 0 Or intIndex > 2 Then intIndex = 0
    For i = 0 To 2
        picRuleBack(i).Visible = intIndex = i
    Next
    If intIndex = 0 Then
        '根据实际打印分配票号
        '设置容器
        Set chkAutoAddBookFee.Container = picRuleBack(0)
        Set chkErrorItemNotBill.Container = fraActuallyPrint
        '工本费
       chkAutoAddBookFee.Top = fraActuallyPrint.Top
       chkAutoAddBookFee.Left = chkOlnyOneBill.Left
       '误差项
       chkErrorItemNotBill.Left = chkExcuteDept.Left
       chkErrorItemNotBill.Top = txtRowsUD.Top + (txtRowsUD.Height - chkErrorItemNotBill.Height) \ 2
    End If

    If intIndex = 1 Then
        '根据预定规则分配票号
        '设置容器
        Set chkAutoAddBookFee.Container = picRuleBack(1)
        Set chkErrorItemNotBill.Container = fraRuleSystem
        '工本费
       chkAutoAddBookFee.Top = fraRuleSystem.Top
       chkAutoAddBookFee.Left = fraRuleSystem.Left + 100
       '误差项
       chkErrorItemNotBill.Left = optRuleTotal(0).Left
       chkErrorItemNotBill.Top = optRuleTotal(0).Top + optRuleTotal(0).Height + 50
       lblInfor.Top = chkErrorItemNotBill.Top + chkErrorItemNotBill.Height + 50
       
    End If
    If intIndex = 2 Then
        '根据用户自定义规则处理
        '设置容器
        Set chkAutoAddBookFee.Container = picRuleBack(2)
        Set chkErrorItemNotBill.Container = picRuleBack(2)
        '工本费
       chkAutoAddBookFee.Top = lblCustomInfor.Top - chkAutoAddBookFee.Height - 50
       chkAutoAddBookFee.Left = lblCustomInfor.Left
       '误差项
       chkErrorItemNotBill.Left = chkAutoAddBookFee.Left + chkAutoAddBookFee.Width + 100
       chkErrorItemNotBill.Top = chkAutoAddBookFee.Top
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Sub

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-04-28 15:09:10
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    Dim strOnePatiPrintShareInvoice As String, intOnePatiPrintType As Integer, varData1 As Variant
    
    On Error GoTo errHandle
    
    
      
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    chkOnePatiPrint.Value = IIf(Val(zlDatabase.GetPara("按病人补打发票不区分结算次数", glngSys, mlngModul, "0", Array(chkOnePatiPrint), blnHavePrivs)) = 1, 1, 0)
    
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "收费票据格式", False, False
    zl_vsGrid_Para_Restore mlngModul, vsDelBillFormat, Me.Name, "退费票据格式", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用收费票据批次", glngSys, mlngModul, , , True, intType)
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
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,使用类别1|领用IDn,使用类别n|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(1)
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
    
    
    With vsBillFormat
        .ColHidden(.ColIndex("按病人补打票据格式")) = chkOnePatiPrint.Value <> 1
    End With
    
    '票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='ZL" & glngSys \ 100 & "_BILL_1121_1'  " & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("收费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("按病人补打票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("收费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With
    

    '读取参数值
    strShareInvoice = zlDatabase.GetPara("收费发票格式", glngSys, mlngModul, , , True, intType)
    strOnePatiPrintShareInvoice = zlDatabase.GetPara("按病人补打发票格式", glngSys, mlngModul, , , True, intOnePatiPrintType)
    strPrintMode = zlDatabase.GetPara("收费发票打印方式", glngSys, mlngModul, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsBillFormat
         .ColData(.ColIndex("收费票据格式")) = "0"
         .ColData(.ColIndex("按病人补打票据格式")) = "0"
         .ColData(.ColIndex("收费打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("收费票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        
        Select Case intOnePatiPrintType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("按病人补打票据格式")) = IIf(intOnePatiPrintType = 5, 0, 1)
        End Select
        
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("收费打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("收费票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("按病人补打票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("收费打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    varData1 = Split(strOnePatiPrintShareInvoice, "|")
    strSQL = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("收费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("收费票据格式")) = "0"
            .TextMatrix(lngRow, .ColIndex("按病人补打票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("收费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varData1)
                varTemp = Split(varData1(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("按病人补打票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("收费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("收费打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("收费打印方式"), .Rows - 1, .ColIndex("收费打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("收费票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("收费票据格式"), .Rows - 1, .ColIndex("收费票据格式")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("按病人补打票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("按病人补打票据格式"), .Rows - 1, .ColIndex("按病人补打票据格式")) = vbBlue
        End If
    End With
    
    '退费票据格式处理
    strSQL = "" & _
    "   Select '使用本地缺省格式' as 说明,0 as 序号  From Dual Union ALL " & _
    "   Select B.说明,B.序号  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.报表ID And A.编号='ZL" & glngSys \ 100 & "_BILL_1121_7'  " & _
    "   Order by  序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .ColComboList(.ColIndex("退费票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
        .ColComboList(.ColIndex("退费打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
    End With

    '读取参数值
    strShareInvoice = zlDatabase.GetPara("退费发票格式", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("退费发票打印方式", glngSys, mlngModul, , , True, intType1)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    With vsDelBillFormat
        .ColData(.ColIndex("退费票据格式")) = "0"
        .ColData(.ColIndex("退费打印方式")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("退费票据格式")) = IIf(intType = 5, 0, 1)
        End Select
        
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("退费打印方式")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("退费票据格式"))) = 1 Or _
            Val(.ColData(.ColIndex("退费打印方式"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsDelBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    strSQL = "" & _
    "   Select 编码 ,名称" & _
    "   From  票据使用类别" & _
    "   order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("使用类别")) = Nvl(rsTemp!名称)
            .TextMatrix(lngRow, .ColIndex("退费打印方式")) = "0-不打印票据"
            .TextMatrix(lngRow, .ColIndex("退费票据格式")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("退费票据格式")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!名称)) Then
                    .TextMatrix(lngRow, .ColIndex("退费打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("退费打印方式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("退费打印方式"), .Rows - 1, .ColIndex("退费打印方式")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("退费票据格式"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("退费票据格式"), .Rows - 1, .ColIndex("退费票据格式")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String, objItem As ListItem, blnParSet As Boolean
    Dim strTmp As String, i As Integer, j As Long, k As Long, arrTmp As Variant, arrWindow As Variant, intType As Integer, blnSeted As Boolean '被设置了缺省值

    Dim str西药房窗口 As String, str中药房窗口 As String, str成药房窗口 As String
    Dim lng缺省西药房 As Long, lng缺省中药房 As Long, lng缺省成药房 As Long, lng缺省发料部门 As Long
    
    gblnOK = False
    On Error GoTo errH
    Call InitTabControl
    blnParSet = InStr(1, mstrPrivs, "参数设置") > 0
    
    'a.初始数据
    '----------------------------------------------------------------------------------------
    '收费类别(挂号除外):按序号排序
    strSQL = "Select 编码,名称 as 类别 from 收费项目类别 Where 编码<>'1' Order by 序号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst收费类别.AddItem rsTmp!类别
        lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    If mbytInFun <> 2 Then
        strSQL = _
            " Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(1,3)" & _
            " Order by 编码"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo费别.ListIndex = cbo费别.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
    If mbytInFun = 0 Then
        lbl单位.Caption = "收费时按："
    ElseIf mbytInFun = 1 Then
        lbl单位.Caption = "划价时按："
    ElseIf mbytInFun = 2 Then
        lbl单位.Caption = "记帐时按："
    End If
    
    'b.本机注册表存储的模块参数
    '----------------------------------------------------------------------------------------
    Call LoadParaValue
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    '--------------------------
    strTmp = zlDatabase.GetPara("科室医生", glngSys, mlngModul, , Array(optUnit, optDoctor, optSelf), blnParSet)
    If strTmp = "1" Then
        optUnit.Value = True
    ElseIf strTmp = "0" Then
        optDoctor.Value = True
    Else
        optSelf.Value = True
    End If
    
    i = IIf(zlDatabase.GetPara("开单人显示方式", glngSys, mlngModul, "1", Array(optDoctorKind(0), optDoctorKind(1)), blnParSet) = "1", 0, 1)
    optDoctorKind(i).Value = True
    
    txtNameDays.Enabled = True
    txtNameDays.Text = zlDatabase.GetPara("姓名查找天数", glngSys, mlngModul, , Array(txtNameDays), blnParSet)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul, , Array(chkSeekName), blnParSet) = "1", 1, 0)
    '调用chkSeekName_click
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        txtMax.Text = Format(zlDatabase.GetPara("最大金额", glngSys, mlngModul, "0", Array(txtMax), blnParSet), "0.00")
        
        chk性别.Value = IIf(zlDatabase.GetPara("性别", glngSys, mlngModul, , Array(chk性别), blnParSet) = "1", 1, 0)
        chk年龄.Value = IIf(zlDatabase.GetPara("年龄", glngSys, mlngModul, , Array(chk年龄), blnParSet) = "1", 1, 0)
        chk费别.Value = IIf(zlDatabase.GetPara("费别", glngSys, mlngModul, , Array(chk费别), blnParSet) = "1", 1, 0)
        chk医疗付款.Value = IIf(zlDatabase.GetPara("医疗付款", glngSys, mlngModul, , Array(chk医疗付款), blnParSet) = "1", 1, 0)
        chk是否加班.Value = IIf(zlDatabase.GetPara("加班", glngSys, mlngModul, , Array(chk是否加班), blnParSet) = "1", 1, 0)
        chk开单日期.Value = IIf(zlDatabase.GetPara("开单日期", glngSys, mlngModul, , Array(chk开单日期), blnParSet) = "1", 1, 0)
        chk开单人.Value = IIf(zlDatabase.GetPara("开单人", glngSys, mlngModul, , Array(chk开单人), blnParSet) = "1", 1, 0)
                
        chk不缺省开单人.Value = IIf(zlDatabase.GetPara("不使用缺省开单人", glngSys, mlngModul, , Array(chk不缺省开单人), blnParSet) = "1", 1, 0)
        chk必须输开单人.Value = IIf(zlDatabase.GetPara("必须要输入开单人", glngSys, mlngModul, , Array(chk必须输开单人), blnParSet) = "1", 1, 0)
        chk缺省科室优先.Value = IIf(zlDatabase.GetPara("缺省科室优先", glngSys, mlngModul, , Array(chk缺省科室优先), blnParSet) = "1", 1, 0)
        chk缺省科室优先.Left = chk不缺省开单人.Left
        Call optUnit_Click
        
        i = Val(zlDatabase.GetPara("分类合计方式", glngSys, mlngModul, , Array(opt分类(0), opt分类(1), opt分类(2)), blnParSet))  '34179
        If i > 2 Or i < 0 Then i = 0
        opt分类(i).Value = True
        
        If mbytInFun = 0 Then
            chk累计.Value = IIf(zlDatabase.GetPara("显示累计", glngSys, mlngModul, , Array(chk累计), blnParSet) = "1", 1, 0)
            chk皮试.Value = IIf(zlDatabase.GetPara("检查皮试结果", glngSys, mlngModul, , Array(chk皮试), blnParSet) = "1", 1, 0)
            chkPrePayPriority.Value = IIf(zlDatabase.GetPara("优先使用预交款", glngSys, mlngModul, , Array(chkPrePayPriority), blnParSet) = "1", 1, 0)
            '91665
            chkInsurePartFee.Value = IIf(zlDatabase.GetPara("只对医保结算成功单据收费", glngSys, mlngModul, , Array(chkInsurePartFee), blnParSet) = "1", 1, 0)
            '96357
            strTmp = zlDatabase.GetPara("本机收费执行科室", glngSys, mlngModul, , Array(chk收费执行科室, txt收费执行科室, cmd收费执行科室), blnParSet)
            mblnNotClick = True
            chk收费执行科室.Value = IIf(strTmp <> "", vbChecked, vbUnchecked)
            mblnNotClick = False
            cmd收费执行科室.Enabled = chk收费执行科室.Value = vbChecked
            txt收费执行科室.Text = GetDeptNameStr(strTmp)
            txt收费执行科室.Tag = strTmp
            '刘兴洪:22343:51670
             i = Val(zlDatabase.GetPara("收费缴款输入控制", glngSys, mlngModul, , Array(opt缴款(0), opt缴款(1), opt缴款(2), opt缴款(3)), blnParSet))
             If i <= opt缴款.UBound And i >= opt缴款.LBound Then opt缴款(i).Value = True
            strTmp = zlDatabase.GetPara("自动加收挂号费", glngSys, mlngModul, , Array(chkAddedItem, txtAddedItem, cmdAddedItem), blnParSet)
            If InStr(1, strTmp, ";") > 0 Then
                chkAddedItem.Value = 1  '会调用click事件,须先加载收费类别
                txtAddedItem.Tag = Split(strTmp, ";")(0)
                txtAddedItem.Text = Split(strTmp, ";")(1)
            End If
'            chkShowError.Value = IIf(zlDatabase.GetPara("显示误差费用", glngSys, mlngModul, , Array(chkShowError), blnParSet) = "1", 1, 0)
            chkMulti.Value = IIf(zlDatabase.GetPara("多单据收费", glngSys, mlngModul, , Array(chkMulti), blnParSet) = "1", 1, 0)
            
            chkSeekBill.Enabled = True
            chkSeekBill.Value = IIf(zlDatabase.GetPara("搜寻划价单据", glngSys, mlngModul, , Array(chkSeekBill), blnParSet) = "1", 1, 0)
            chkSeekBill.Tag = IIf(chkSeekBill.Enabled, "1", "0")
            txtSeekDays.Enabled = True
            txtSeekDays.Text = zlDatabase.GetPara("搜寻单据天数", glngSys, mlngModul, , Array(txtSeekDays), blnParSet)
            txtSeekDays.Tag = IIf(txtSeekDays.Enabled, "1", "0")
            chkUnPopPriceBill.Enabled = True
            chkUnPopPriceBill.Value = IIf(zlDatabase.GetPara("不弹出划价单选择", glngSys, mlngModul, , Array(chkUnPopPriceBill), blnParSet) = "1", 1, 0)
            chkUnPopPriceBill.Tag = IIf(chkUnPopPriceBill.Enabled, "1", "0")
            
            chkMustRegevent.Value = IIf(zlDatabase.GetPara("检查病人挂号科室", glngSys, mlngModul, , Array(chkMustRegevent), blnParSet) = "1", 1, 0)
            i = Val(zlDatabase.GetPara("未挂号病人收费", glngSys, mlngModul, , Array(optRegPrompt(0), optRegPrompt(1), optRegPrompt(2)), blnParSet))
            If i <= optRegPrompt.UBound Then optRegPrompt(i).Value = True
            
            chkAutoSplitBill.Enabled = True
            cboAutoSplitBill.Enabled = True
            i = Val(zlDatabase.GetPara("自动组合单据", glngSys, mlngModul, , Array(chkAutoSplitBill, cboAutoSplitBill), blnParSet))
            chkAutoSplitBill.Tag = IIf(chkAutoSplitBill.Enabled, "1", "0")
            cboAutoSplitBill.Tag = IIf(cboAutoSplitBill.Enabled, "1", "0")
            chkAutoSplitBill.Value = IIf(i > 0, 1, 0)
            cboAutoSplitBill.AddItem "收费类别"
            cboAutoSplitBill.AddItem "执行科室"
            cboAutoSplitBill.ListIndex = IIf(i = 1 Or i = 2, i - 1, 0)
            If chkAutoSplitBill.Value = 0 Then cboAutoSplitBill.Enabled = False
            Call chkMulti_Click
                
            i = Val(zlDatabase.GetPara("收费清单打印方式", glngSys, mlngModul, , Array(optPrint(0), optPrint(1), optPrint(2)), blnParSet))
            If i <= optPrint.UBound Then optPrint(i).Value = True
            i = Val(zlDatabase.GetPara("退费回单打印方式", glngSys, mlngModul, , Array(optRefund(0), optRefund(1), optRefund(2)), blnParSet))
            If i <= optRefund.UBound Then optRefund(i).Value = True
            '62982:李南春,2015/08/25,收费执行单
            i = Val(zlDatabase.GetPara("收费执行单打印方式", glngSys, mlngModul, , Array(optExe(0), optExe(1), optExe(2)), blnParSet))
            If i <= optExe.UBound Then optExe(i).Value = True
            
            '刘兴洪 问题:26948 日期:2009-12-28 16:54:11
            strTmp = zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, mlngModul, "0|10", Array(txt票据张数, upd票据张数, chk票据张数), blnParSet)
            
            upd票据张数.Value = Val(Split(strTmp & "|", "|")(1))
            txt票据张数.Text = upd票据张数.Value
            chk票据张数.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
            txt票据张数.Enabled = chk票据张数.Enabled And chk票据张数.Value = 1
            upd票据张数.Enabled = txt票据张数.Enabled
        Else
            txtDay.Text = zlDatabase.GetPara("取消划价单", glngSys, mlngModul, , Array(txtDay, lblDay, udDay), blnParSet)
            
            i = Val(zlDatabase.GetPara("划价通知单打印方式", glngSys, mlngModul, , Array(optPrintRequisition(0), optPrintRequisition(1), optPrintRequisition(2)), blnParSet))
            If i <= optPrintRequisition.UBound Then optPrintRequisition(i).Value = True
        End If
    ElseIf mbytInFun = 2 Then
        chkOnlyUnitPatient.Enabled = True
        chkOnlyUnitPatient.Value = IIf(zlDatabase.GetPara("只查找合约单位病人", glngSys, mlngModul, , Array(chkOnlyUnitPatient), blnParSet) = "1", 1, 0)
        chkOnlyUnitPatient.Tag = IIf(chkOnlyUnitPatient.Enabled, "1", "0")
        chkOnlyUnitPatient.Enabled = chkSeekName.Value = 1 And chkOnlyUnitPatient.Tag = "1"
        
        Call chkSeekName_Click
        
        chk(0).Value = IIf(zlDatabase.GetPara("记帐打印", glngSys, mlngModul, , Array(chk(0)), blnParSet) = "1", 1, 0)
        chk(1).Value = IIf(zlDatabase.GetPara("划价打印", glngSys, mlngModul, , Array(chk(1)), blnParSet) = "1", 1, 0)
        chk(2).Value = IIf(zlDatabase.GetPara("审核打印", glngSys, mlngModul, , Array(chk(2)), blnParSet) = "1", 1, 0)
    End If
    
    '刘兴洪 问题:26948 日期:2009-12-28 16:54:11
    If mbytInFun <> 0 Then
        txt票据张数.Visible = False: upd票据张数.Visible = False: chk票据张数.Visible = False
    End If
    
    Call SetCtrlVisible
    
    'd.权限控制
    '----------------------------------------------------------------------------------------
    chkLed.Visible = mbytInFun = 0
    If InStr(mstrPrivs, "LED与语音") = 0 Then
        chkLed.Visible = False
        chkLed.Value = Unchecked
    End If
    chkLedDispDetail.Visible = mbytInFun = 0
    chkLedWelcome.Visible = mbytInFun = 0
    chk医保结算光标缺省定位.Visible = mbytInFun = 0
    
    chk皮试.Visible = mbytInFun = 0
    chkPrePayPriority.Visible = mbytInFun = 0
    lbl结算方式.Visible = mbytInFun = 0
    cbo结算方式.Visible = mbytInFun = 0
    
    chkAddedItem.Visible = mbytInFun = 0
    txtAddedItem.Visible = mbytInFun = 0
    cmdAddedItem.Visible = mbytInFun = 0
    
    chkInsurePartFee.Visible = mbytInFun = 0
    chk禁止取消挂号单.Visible = mbytInFun = 0
    
    '根据病人搜寻划价单据
    chkSeekBill.Visible = mbytInFun = 0
    txtSeekDays.Visible = mbytInFun = 0
    fraLine.Visible = mbytInFun = 0
    chkUnPopPriceBill.Visible = mbytInFun = 0
    
    '支持多单据收费,显示误差费用
    fraRegPrompt.Visible = mbytInFun = 0
    chkMustRegevent.Visible = mbytInFun = 0
    chkMulti.Visible = mbytInFun = 0
    chkAutoSplitBill.Visible = mbytInFun = 0
    cboAutoSplitBill.Visible = mbytInFun = 0
'    chkShowError.Visible = mbytInFun = 0
    
    '刘兴洪 问题号:22343
    fra缴款控制.Visible = mbytInFun = 0
    chk累计.Visible = mbytInFun = 0
    fra划价通知单打印.Visible = mbytInFun = 1
    cmdPrintSetup(3).Visible = mbytInFun = 1
    lblDay.Visible = mbytInFun = 1
    txtDay.Visible = mbytInFun = 1
    udDay.Visible = mbytInFun = 1
    
    chkOnlyUnitPatient.Visible = mbytInFun = 2
    fraPrintBill.Visible = mbytInFun = 2
        
    lbl费别.Visible = mbytInFun <> 2
    cbo费别.Visible = mbytInFun <> 2
    lblMax.Visible = mbytInFun <> 2
    txtMax.Visible = mbytInFun <> 2
    
    stab.TabVisible(1) = mbytInFun <> 2
    stab.TabVisible(2) = mbytInFun = 0
    stab.TabVisible(3) = mbytInFun = 0  '56963
    
    txt收费执行科室.Visible = mbytInFun = 0
    chk收费执行科室.Visible = mbytInFun = 0
    cmd收费执行科室.Visible = mbytInFun = 0
    
    'f.位置调整
    '-------------------------------------------------------------
    Call MoveCtrol
    '药店设置
    If glngSys Like "8??" Then
        fra病人.Visible = False
                
        lbl费别.Caption = "缺省会员等级"
                
        chk药房.Caption = "显示其它药店库存"
        fra类别.Visible = False '固定输入药品类别
        fra科室与医生.Visible = False '固定独立输入
        
        chk护士.Visible = False
        chk护士.Value = 0
    End If
    
    If mblnSetDrugStore Then
        '56963
        stab.TabCaption(4) = "药房设置"
        stab.TabVisible(0) = False
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
    Else
        If mbytInFun = 1 Then
            stab.TabCaption(4) = "药房设置(&3)"
        ElseIf mbytInFun = 2 Then
            stab.TabCaption(4) = "药房设置(&2)"
        End If
    End If
    
    If stab.TabVisible(0) Then stab.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDeptNameStr(ByVal strIDs As String) As String
    '将部门ID字符串装换成名称字符串
    '入参：
    '   strIDs 部门ID，格式：ID1,ID2,ID3,...
    '返回：
    '   部门名称s，格式：部门名称1;部门名称2;部门名称3;...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If strIDs = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10) */a.编码, a.名称" & vbNewLine & _
            " From 部门表 A, Table(f_Str2list([1], ',')) B" & vbNewLine & _
            " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "根据科室ID获取科室名称", strIDs)
    If rsTemp Is Nothing Then Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & ";" & Nvl(rsTemp!名称)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetDeptNameStr = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnSetDrugStore = False
    If mbytInFun = 0 Then
        zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共享票据批次列", False, False
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "收费票据格式", False, False
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "退费票据格式", False, False
    End If
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub
Private Sub optDoctor_Click()
    Call optUnit_Click
End Sub

Private Sub optSelf_Click()
    Call optUnit_Click
End Sub

Private Sub optUnit_Click()
    chk不缺省开单人.Visible = optUnit.Value
    chk缺省科室优先.Visible = optDoctor.Value
End Sub

Private Sub opt病人_Click(Index As Integer)
    If gbln分离发药 Then
        Call SetStockCheck
    Else
        Call SetDrugStore
    End If
    
    If opt病人(0).Value Then
        opt单位(1).Caption = "门诊单位"
    Else
        opt单位(1).Caption = "住院单位"
    End If
End Sub

Private Sub stab_Click(PreviousTab As Integer)
    Select Case stab.Tab
        Case 0
            If txtMax.Visible Then
                If txtMax.Enabled And txtMax.Visible Then txtMax.SetFocus
            Else
                If cbo费别.Enabled And cbo费别.Visible Then cbo费别.SetFocus
            End If
        Case 1
            If opt病人(0).Enabled And opt病人(0).Visible And opt病人(0).Value Then
                opt病人(0).SetFocus
            ElseIf opt病人(1).Enabled And opt病人(1).Visible And opt病人(1).Value Then
                opt病人(1).SetFocus
            End If
        Case 2
            If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
        Case 3
            ' 56963
            If cboBillRole.Visible And cboBillRole.Enabled Then cboBillRole.SetFocus
        Case 4
            ' 56963
            If vsfDrugStore.Visible And vsfDrugStore.Enabled Then vsfDrugStore.SetFocus
    End Select
End Sub

Private Sub txtBillRuleNum_Change(Index As Integer)
    '56963
    Call ShowRuleInfor
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDay_LostFocus()
    If Not IsNumeric(txtDay.Text) Then txtDay.Text = 0
End Sub

Private Sub txtMax_GotFocus()
    SelAll txtMax
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMax_LostFocus()
    If IsNumeric(txtMax.Text) Then
        txtMax.Text = Format(txtMax.Text, "0.00")
    Else
        txtMax.Text = "0.00"
    End If
End Sub

Private Sub txtSeekDays_GotFocus()
    Call SelAll(txtSeekDays)
End Sub

Private Sub txtSeekDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSeekDays_Validate(Cancel As Boolean)
    If Val(txtSeekDays.Text) < 1 Then
        txtSeekDays.Text = 1
    End If
End Sub

Private Sub txtNameDays_GotFocus()
    Call SelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
End Sub
Private Sub updBillRuleNum_Change(Index As Integer)
    If updBillRuleNum(Index).Value = 0 Then
        chkBillRule(Index + 1).Value = 0
    End If
End Sub

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
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "收费票据格式", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "收费票据格式", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("收费票据格式"), .ColIndex("按病人补打票据格式"), .ColIndex("收费打印方式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsDelBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "退费票据格式", False, False
End Sub
Private Sub vsDelBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "退费票据格式", False, False
End Sub

Private Sub vsDelBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDelBillFormat
        Select Case Col
        Case .ColIndex("退费票据格式"), .ColIndex("退费打印方式")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsfDrugStore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("缺省")
           Call SetDrugStockDeFault(Row)
        Case Else
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("缺省"), .ColIndex("窗口")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    Dim strTmp As String, i As Long
    With vsfDrugStore
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("缺省")) = 1 Then Exit Sub
        
        .TextMatrix(.Row, .Col) = IIf(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row)
    End With
End Sub
Private Sub SetDrugStockDeFault(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的缺省值
    '入参:lngRow-指定行
    '编制:刘兴洪
    '日期:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lng缺省 As Long, strType As String
    With vsfDrugStore
        lng缺省 = Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省"))))
        If lng缺省 = 1 Then
            strType = .TextMatrix(lngRow, .ColIndex("类别"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = strType And i <> lngRow Then
                    .TextMatrix(i, .ColIndex("缺省")) = 0
                End If
            Next
        End If
    End With
End Sub
Private Sub SetDrugStockEdit(ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置药房的编辑属性
    '入参:strType-类别
    '     intType-返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    '     lngEditCol-控制的编辑列
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-02 14:53:10
    '问题:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '设置了缺省值了,随后不能再设置缺省值
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-锁定,0-不锁定
    
    '刘兴洪:由于可能参数权限发生变更,因此,不能统一进行设置,需要设置某一部分:
    With vsfDrugStore
        blnSetDefault = False: blnAllowEdit = InStr(1, mstrPrivs, ";参数设置;") > 0
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
            lngEditForColor = IIf(blnAllowEdit, vbBlue, &H8000000C)  '授权限控制
            bytLockEdit = IIf(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
            lngEditForColor = vbBlue    '公共模块,但不授权限控制
        Else
            lngEditForColor = &H80000008    '正常编辑
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = strType Then
                If lngEditCol = .ColIndex("缺省") Then
                    '设置药房
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, .ColIndex("缺省")) = IIf(Val(strMachValue) > 0, 1, 0)
                        blnSetDefault = True
                    End If
                     .Cell(flexcpForeColor, i, .ColIndex("缺省")) = lngEditForColor
                     .Cell(flexcpForeColor, i, .ColIndex("药房")) = lngEditForColor:
                Else
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, lngEditCol) = strDefaultValue
                    End If
                    '设置窗口
                     .Cell(flexcpForeColor, i, .ColIndex("窗口")) = lngEditForColor
                End If
                .Cell(flexcpData, i, lngEditCol) = bytLockEdit
            End If
        Next
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("窗口") Then
                Set rsTmp = Read发药窗口(.RowData(.Row))
                strList = "自动分配|" & .BuildComboList(rsTmp, "名称")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
              '  .Editable = flexEDNone
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub InitTabControl()
    With tbBillSet
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
'        .PaintManager.StaticFrame = True
'        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 0, "收费票据格式", picBillFormat.hWnd, 0
        .InsertItem 1, "退费票据格式", picDelBillFormat.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub
