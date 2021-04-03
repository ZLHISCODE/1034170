VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLocalPara 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "本机参数设置"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "frmLocalPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab Tabs1 
      Height          =   6585
      Left            =   120
      TabIndex        =   95
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   617
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmLocalPara.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSortMode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblColor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstDept"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraDefaultSet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraClearMZInfor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboSortMode"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pic提前颜色"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "输入输出(&1)"
      TabPicture(1)   =   "frmLocalPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblGuardian"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkSeekName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraTitle"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraInput"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDeviceSetup"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkPrintFree"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkTotal"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkDoctor"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "fraLine2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtNameDays"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fraInvoice"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkPrintCase"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fraDeposit"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkAddressAssnInput"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtMustGuardianInfo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fraSlip"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fra退号回单打印方式"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "医疗卡(&2)"
      TabPicture(2)   =   "frmLocalPara.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDefaultPayCard"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkAutoAddName"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkCardMoney"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkNewCardNoPop"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkRePrint"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboType"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraCards"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkScanIDVisa"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkAlwaysSendCard"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "ChkMustBill"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "预约挂号(&3)"
      TabPicture(3)   =   "frmLocalPara.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblAvailabilityTimes"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblBreakAnAppointmentNums"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Line1(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Line1(2)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblCancelBespeak"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Line1(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Line1(4)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblBespeakMinTime"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblBespeakDefaultDays"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Line1(5)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Line1(1)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Line1(7)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label1"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "chkDeptNums"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txtAvailabilityTimes"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "chkDeptBespeakOneNum"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Frame3"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Frame5"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "txtBreakAnAppointmentNums"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "chkBespeakFee"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "chkMzh"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "chkBackNoToVerfy"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "chkBreakAnAppointmentToRegist"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "txtCancelBespeak"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "txtBespeakMinTime"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "txtBespeakDefaultDays"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "fraBespeak"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "fraReceiveMode"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "txtDeptNums"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "txtDeptBespeakOneNum"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "cboDefaultStyle"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "cbo预约有效时间"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).ControlCount=   34
      Begin VB.Frame fra退号回单打印方式 
         Caption         =   "退号回单打印方式"
         Height          =   645
         Left            =   -74760
         TabIndex        =   127
         Top             =   4470
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   4
            Left            =   5160
            TabIndex        =   60
            Top             =   210
            Width           =   990
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "选择是否打印"
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   59
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "自动打印"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   58
            Top             =   300
            Width           =   1275
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "不打印"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.PictureBox pic提前颜色 
         BackColor       =   &H00000000&
         Height          =   270
         Left            =   -68415
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   126
         Top             =   5820
         Width           =   270
      End
      Begin VB.ComboBox cbo预约有效时间 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   3270
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   4110
         Width           =   780
      End
      Begin VB.CheckBox ChkMustBill 
         Caption         =   "严格控制下卡费为零也走票号"
         Height          =   195
         Left            =   -74625
         TabIndex        =   68
         Top             =   2520
         Width           =   3345
      End
      Begin VB.ComboBox cboDefaultStyle 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   3405
         Width           =   2160
      End
      Begin VB.ComboBox cboSortMode 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -69480
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   5445
         Width           =   1350
      End
      Begin VB.TextBox txtDeptBespeakOneNum 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   74
         Text            =   "0"
         Top             =   1200
         Width           =   660
      End
      Begin VB.TextBox txtDeptNums 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   72
         Text            =   "0"
         Top             =   885
         Width           =   660
      End
      Begin VB.Frame fraReceiveMode 
         Caption         =   "预约接收模式"
         Height          =   615
         Left            =   480
         TabIndex        =   123
         Top             =   5835
         Width           =   6210
         Begin VB.OptionButton optReceiveMode 
            Caption         =   "仅预约接收"
            Height          =   255
            Index           =   1
            Left            =   2865
            TabIndex        =   91
            Top             =   255
            Width           =   3165
         End
         Begin VB.OptionButton optReceiveMode 
            Caption         =   "预约接收就诊"
            Height          =   255
            Index           =   0
            Left            =   675
            TabIndex        =   90
            Top             =   255
            Width           =   2130
         End
      End
      Begin VB.CheckBox chkAlwaysSendCard 
         Caption         =   "非严格控制卡时始终为发卡"
         Height          =   195
         Left            =   -74625
         TabIndex        =   67
         Top             =   2190
         Width           =   3345
      End
      Begin VB.Frame fraSlip 
         Caption         =   "挂号凭条"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74760
         TabIndex        =   122
         Top             =   3180
         Width           =   6480
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   51
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   50
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   49
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   3
            Left            =   5160
            TabIndex        =   52
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.TextBox txtMustGuardianInfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "0"
         Top             =   2100
         Width           =   660
      End
      Begin VB.Frame fraBespeak 
         Caption         =   "预约挂号单"
         Height          =   615
         Left            =   480
         TabIndex        =   120
         Top             =   5085
         Width           =   6210
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1755
            TabIndex        =   87
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   86
            Top             =   315
            Width           =   900
         End
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2850
            TabIndex        =   88
            Top             =   300
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   2
            Left            =   5100
            TabIndex        =   89
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.CheckBox chkScanIDVisa 
         Caption         =   "扫描身份证签约"
         Height          =   195
         Left            =   -74625
         TabIndex        =   66
         Top             =   1875
         Width           =   3345
      End
      Begin VB.CheckBox chkAddressAssnInput 
         Caption         =   "家庭地址联想输入"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         ToolTipText     =   "挂号时家庭地址输入时是否联想"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "病人条码打印"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74760
         TabIndex        =   119
         Top             =   3825
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   0
            Left            =   5160
            TabIndex        =   56
            Top             =   180
            Width           =   990
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   53
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   54
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   55
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.TextBox txtBespeakDefaultDays 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   81
         Text            =   "0"
         Top             =   3060
         Width           =   540
      End
      Begin VB.TextBox txtBespeakMinTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   80
         Text            =   "0"
         Top             =   2640
         Width           =   540
      End
      Begin VB.TextBox txtCancelBespeak 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   77
         Text            =   "0"
         Top             =   1875
         Width           =   660
      End
      Begin VB.CheckBox chkBreakAnAppointmentToRegist 
         Caption         =   "预约失约用于挂号"
         Height          =   210
         Left            =   3960
         TabIndex        =   78
         Top             =   1890
         Width           =   2055
      End
      Begin VB.CheckBox chkBackNoToVerfy 
         Caption         =   "退号审核:N天内取消预约需要通过审核"
         Height          =   210
         Left            =   720
         TabIndex        =   79
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Frame fraCards 
         Caption         =   "本地共用医疗卡"
         Height          =   2655
         Left            =   -74775
         TabIndex        =   114
         Top             =   2895
         Width           =   6660
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2220
            Left            =   195
            TabIndex        =   69
            Top             =   300
            Width           =   6405
            _cx             =   11298
            _cy             =   3916
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
            FormatString    =   $"frmLocalPara.frx":007C
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
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73548
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   5640
         Width           =   2580
      End
      Begin VB.CheckBox chkMzh 
         Caption         =   "预约时不生成门诊号"
         Height          =   285
         Left            =   3960
         TabIndex        =   76
         Top             =   1515
         Width           =   2655
      End
      Begin VB.Frame fraClearMZInfor 
         Caption         =   "退号清除门诊号信息(挂号有效天数内的病人)"
         Height          =   645
         Left            =   -74865
         TabIndex        =   113
         Top             =   5805
         Width           =   4455
         Begin VB.OptionButton optClearInfor 
            Caption         =   "提示清除"
            Height          =   180
            Index           =   2
            Left            =   2880
            TabIndex        =   24
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton optClearInfor 
            Caption         =   "自动清除"
            Height          =   180
            Index           =   1
            Left            =   1410
            TabIndex        =   23
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton optClearInfor 
            Caption         =   "不清除"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   22
            Top             =   300
            Width           =   1110
         End
      End
      Begin VB.CheckBox chkBespeakFee 
         Caption         =   "挂号费用以预约接收时间为准!"
         Height          =   210
         Left            =   720
         TabIndex        =   75
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtBreakAnAppointmentNums 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1665
         MaxLength       =   4
         TabIndex        =   85
         Text            =   "0"
         Top             =   4650
         Width           =   660
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   1425
         TabIndex        =   111
         Top             =   3750
         Width           =   4845
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   1290
         TabIndex        =   109
         Top             =   540
         Width           =   5130
      End
      Begin VB.CheckBox chkDeptBespeakOneNum 
         Caption         =   "病人同一科室限约        个号"
         Height          =   210
         Left            =   720
         TabIndex        =   73
         Top             =   1215
         Width           =   3735
      End
      Begin VB.TextBox txtAvailabilityTimes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4035
         MaxLength       =   4
         TabIndex        =   84
         Text            =   "0"
         Top             =   4170
         Width           =   660
      End
      Begin VB.CheckBox chkPrintCase 
         Caption         =   "挂号后打印病历标签"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         ToolTipText     =   "购买病历后打印病历标签"
         Top             =   1550
         Width           =   1980
      End
      Begin VB.CheckBox chkRePrint 
         Caption         =   "退号不退卡时重打票据"
         Height          =   195
         Left            =   -74625
         TabIndex        =   65
         Top             =   1545
         Width           =   2400
      End
      Begin VB.Frame fraInvoice 
         Caption         =   "挂号票据"
         Height          =   615
         Left            =   -74760
         TabIndex        =   106
         Top             =   2520
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   1
            Left            =   5175
            TabIndex        =   48
            Top             =   180
            Width           =   990
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   47
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   45
            Top             =   285
            Width           =   900
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   46
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkNewCardNoPop 
         Caption         =   "发卡不弹出病人信息登记窗口"
         Height          =   195
         Left            =   -74640
         TabIndex        =   62
         Top             =   600
         Width           =   3345
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   -72870
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "0表示查找时不限制时间"
         Top             =   480
         Width           =   285
      End
      Begin VB.Frame fraLine2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   -72840
         TabIndex        =   105
         Top             =   675
         Width           =   285
      End
      Begin VB.CheckBox chkDoctor 
         Caption         =   "挂未设置医生的号别时必须输入医生"
         Height          =   195
         Left            =   -74760
         TabIndex        =   30
         ToolTipText     =   "如果所选号别没有医生,则允许选择费别所定科室的医生,否则不允许选择."
         Top             =   750
         Width           =   3180
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "输入缴款金额之后才结束本次挂号收费"
         Height          =   195
         Left            =   -74760
         TabIndex        =   31
         Top             =   1020
         Width           =   3360
      End
      Begin VB.CheckBox chkCardMoney 
         Caption         =   "卡费与挂号费一起收(否则卡费存为划价单)"
         Height          =   195
         Left            =   -74640
         TabIndex        =   64
         Top             =   1200
         Width           =   3960
      End
      Begin VB.CheckBox chkAutoAddName 
         Caption         =   "发卡新病人自动产生临时姓名"
         Height          =   195
         Left            =   -74640
         TabIndex        =   63
         Top             =   900
         Width           =   3345
      End
      Begin VB.CheckBox chkPrintFree 
         Caption         =   "挂号费用为零时也打印票据"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         ToolTipText     =   "号别要求建病案的时此项必输"
         Top             =   1260
         Width           =   2460
      End
      Begin VB.Frame fraDefaultSet 
         Caption         =   "缺省值"
         Height          =   1005
         Left            =   -74850
         TabIndex        =   17
         Top             =   4695
         Width           =   4440
         Begin VB.ComboBox cboDefaultSex 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   615
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultPayMode 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultFeeType 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultBalance 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   645
            Width           =   1350
         End
         Begin VB.Label lblDefaultSex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "性别"
            Height          =   180
            Left            =   2460
            TabIndex        =   103
            Top             =   675
            Width           =   360
         End
         Begin VB.Label lblDefaultPayMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "付款方式"
            Height          =   180
            Left            =   180
            TabIndex        =   102
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblDefaultBalance 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算方式"
            Height          =   180
            Left            =   180
            TabIndex        =   101
            Top             =   705
            Width           =   720
         End
         Begin VB.Label lblDefaultFeeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费别"
            Height          =   180
            Left            =   2460
            TabIndex        =   100
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   330
         Left            =   -69810
         TabIndex        =   44
         Top             =   2190
         Width           =   1425
      End
      Begin VB.ListBox lstDept 
         ForeColor       =   &H80000012&
         Height          =   4680
         Left            =   -70320
         Style           =   1  'Checkbox
         TabIndex        =   25
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全消,如果一个都未选则表示不限制科室"
         Top             =   660
         Width           =   2175
      End
      Begin VB.Frame fraInput 
         Caption         =   "要求输入项"
         Height          =   1635
         Left            =   -70800
         TabIndex        =   98
         Top             =   480
         Width           =   2715
         Begin VB.CheckBox chkAllowPhoneInput 
            Caption         =   "联系电话"
            Height          =   195
            Left            =   1560
            TabIndex        =   43
            Top             =   1200
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowPatientInput 
            Caption         =   "病人"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "号别要求建病案的时此项必输"
            Top             =   315
            Width           =   660
         End
         Begin VB.CheckBox chkAllowSexInput 
            Caption         =   "性别"
            Height          =   195
            Left            =   1560
            TabIndex        =   37
            ToolTipText     =   "号别要求建病案的时此项必输"
            Top             =   315
            Width           =   660
         End
         Begin VB.CheckBox chkAllowAgeInput 
            Caption         =   "年龄"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            ToolTipText     =   "号别要求建病案的时此项必输"
            Top             =   610
            Width           =   660
         End
         Begin VB.CheckBox chkAllowFeeTypeInput 
            Caption         =   "费别"
            Height          =   195
            Left            =   1560
            TabIndex        =   39
            Top             =   610
            Width           =   660
         End
         Begin VB.CheckBox chkAllowBalanceInput 
            Caption         =   "结算方式"
            Height          =   195
            Left            =   1560
            TabIndex        =   41
            Top             =   915
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowPayModeInput 
            Caption         =   "付款方式"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   905
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowAddressInput 
            Caption         =   "家庭地址"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1200
            Width           =   1020
         End
      End
      Begin VB.Frame fraTitle 
         Caption         =   "共用挂号票据"
         Height          =   1155
         Left            =   -74760
         TabIndex        =   97
         Top             =   5160
         Width           =   6675
         Begin MSComctlLib.ListView lvwBill 
            Height          =   1455
            Left            =   150
            TabIndex        =   61
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483630
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "领用人"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "领用日期"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "号码范围"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "剩余"
               Object.Width           =   1499
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "功能参数"
         Height          =   4560
         Left            =   -74880
         TabIndex        =   96
         Top             =   480
         Width           =   4395
         Begin VB.CheckBox chkDeptRegistOneEmer 
            Caption         =   "病人同科挂号限制用于急诊"
            Height          =   210
            Left            =   360
            TabIndex        =   14
            Top             =   3420
            Width           =   3735
         End
         Begin VB.TextBox txtDeptRegistOneNum 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "0"
            Top             =   3135
            Width           =   660
         End
         Begin VB.TextBox txtRegistDeptNums 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2055
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "0"
            Top             =   3675
            Width           =   660
         End
         Begin VB.CheckBox chkRegistDeptNums 
            Caption         =   "同一病人最多能挂号        个科室"
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   3690
            Width           =   3705
         End
         Begin VB.CheckBox chkDeptRegistOneNum 
            Caption         =   "病人同一科室限挂        个号"
            Height          =   210
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Width           =   3735
         End
         Begin VB.CheckBox chkNOValidityCheck 
            Caption         =   "门诊号有效性检查"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2835
            Width           =   2325
         End
         Begin VB.CheckBox chkDefaultMedBook 
            Caption         =   "挂号时默认勾选购买病历选项"
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   2550
            Width           =   2700
         End
         Begin VB.CheckBox chkReuseCanceledNO 
            Caption         =   "已退序号允许挂号"
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   2550
         End
         Begin VB.CheckBox chkTimeRangeRegist 
            Caption         =   "分时段号别严格按时段挂号"
            Height          =   300
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   $"frmLocalPara.frx":015D
            Top             =   2010
            Width           =   2550
         End
         Begin VB.CheckBox chkRandSelectNum 
            Caption         =   "随机序号选择"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1500
            Width           =   2325
         End
         Begin VB.CheckBox chkRigistHeadSort 
            Caption         =   "挂号安排表点击列头排序"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1755
            Width           =   2325
         End
         Begin VB.CheckBox chkAllowZyRigist 
            Caption         =   "允许住院病人挂号"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1215
            Width           =   1845
         End
         Begin VB.CheckBox chkPrePayPriority 
            Caption         =   "优先使用预交款缴费"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   975
            Width           =   2340
         End
         Begin VB.CheckBox chkPrice 
            Caption         =   "建档病人挂号存为划价单    (此模式下不能进行医保验卡)"
            Height          =   435
            Left            =   120
            TabIndex        =   3
            Top             =   510
            Width           =   2700
         End
         Begin VB.TextBox txtInterval 
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
            Left            =   825
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "5"
            Top             =   30
            Width           =   285
         End
         Begin VB.Frame fraLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   820
            TabIndex        =   104
            Top             =   225
            Width           =   285
         End
         Begin VB.CheckBox chkAutoRefresh 
            Caption         =   "每隔     分钟自动刷新挂号安排表"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   30
            Width           =   3480
         End
         Begin VB.CheckBox chkAutoGet 
            Caption         =   "挂号必须建病案时自动产生病人门诊号"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Width           =   3360
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   1815
            X2              =   2535
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2025
            X2              =   2745
            Y1              =   3915
            Y2              =   3915
         End
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "输入姓名后模糊查找    天内的病人"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   3300
      End
      Begin VB.CheckBox chkDeptNums 
         Caption         =   "同一病人最多能预约        个科室"
         Height          =   210
         Left            =   720
         TabIndex        =   71
         Top             =   900
         Width           =   3705
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新版提前挂号安排颜色"
         Height          =   180
         Left            =   -70320
         TabIndex        =   27
         Top             =   5850
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "缺省预约方式"
         Height          =   180
         Left            =   720
         TabIndex        =   125
         Top             =   3465
         Width           =   1080
      End
      Begin VB.Label lblSortMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "排序方式"
         Height          =   180
         Left            =   -70320
         TabIndex        =   124
         Top             =   5505
         Width           =   720
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2430
         X2              =   3150
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2640
         X2              =   3360
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   -74760
         X2              =   -74040
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label lblGuardian 
         AutoSize        =   -1  'True
         Caption         =   "岁以下必须录入监护人"
         Height          =   180
         Left            =   -74040
         TabIndex        =   121
         Top             =   2100
         Width           =   1800
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1800
         X2              =   2520
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label lblBespeakDefaultDays 
         AutoSize        =   -1  'True
         Caption         =   "预约缺省天数        天"
         Height          =   180
         Left            =   720
         TabIndex        =   118
         Top             =   3045
         Width           =   1980
      End
      Begin VB.Label lblBespeakMinTime 
         AutoSize        =   -1  'True
         Caption         =   "预约限制时间        分钟：指预约时间距离现在时刻的最小间隔"
         Height          =   180
         Left            =   705
         TabIndex        =   117
         Top             =   2655
         Width           =   5220
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1800
         X2              =   2520
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1320
         X2              =   2040
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Label lblCancelBespeak 
         AutoSize        =   -1  'True
         Caption         =   "预约号         天内不能取消预约"
         Height          =   180
         Left            =   735
         TabIndex        =   116
         Top             =   1890
         Width           =   2790
      End
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "缺省发卡类型"
         Height          =   210
         Left            =   -74700
         TabIndex        =   115
         Top             =   5700
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4080
         X2              =   4800
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1695
         X2              =   2415
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label lblBreakAnAppointmentNums 
         AutoSize        =   -1  'True
         Caption         =   "病人预约失约         次自动进入黑名单"
         Height          =   180
         Left            =   540
         TabIndex        =   112
         Top             =   4650
         Width           =   3330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "黑名单控制"
         Height          =   180
         Left            =   540
         TabIndex        =   110
         Top             =   3750
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "预约限制"
         Height          =   180
         Left            =   525
         TabIndex        =   108
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblAvailabilityTimes 
         AutoSize        =   -1  'True
         Caption         =   "预约有效时间：预约单在预约时间                 分钟未接收的为失约！"
         Height          =   180
         Left            =   540
         TabIndex        =   107
         Top             =   4170
         Width           =   6030
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号科室"
         Height          =   180
         Left            =   -70335
         TabIndex        =   99
         ToolTipText     =   "设定本机可挂哪些科室的号"
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   330
      Left            =   120
      TabIndex        =   92
      Top             =   6825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   6060
      TabIndex        =   94
      Top             =   6825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   330
      Left            =   4680
      TabIndex        =   93
      Top             =   6825
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLocalPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long
Private mstrColor As String

'Private mblnOK As Boolean
Private Sub cboType_Click()
  chkScanIDVisa.Enabled = Not (cboType.Text = "二代身份证")
  If cboType.Text = "二代身份证" Then
        chkScanIDVisa.Value = 1
  End If
End Sub
Private Sub chkAutoAddName_Click()
    If chkAutoAddName.Value = 1 Then
        chkAllowPatientInput.Value = 0
        chkAllowSexInput.Value = 0
        chkAllowAgeInput.Value = 0
        chkAllowAddressInput.Value = 0
        chkAllowPayModeInput.Value = 0
        chkNewCardNoPop.Value = 0
    End If
End Sub

Private Sub chkAutoRefresh_Click()
    txtInterval.Enabled = chkAutoRefresh.Value = 1
    If txtInterval.Enabled And txtInterval.Visible Then
        txtInterval.SetFocus
    End If
End Sub

Private Sub chkCardMoney_Click()
    If chkCardMoney.Value = 0 Then
        chkNewCardNoPop.Value = 0
        
        chkRePrint.Value = 0
        chkRePrint.Enabled = False
    Else
        chkRePrint.Enabled = True
    End If
End Sub

Private Sub chkDeptBespeakOneNum_Click()
    If chkDeptBespeakOneNum.Value = 1 Then
        txtDeptBespeakOneNum.Enabled = True
    Else
        txtDeptBespeakOneNum.Enabled = False
    End If
End Sub

Private Sub chkDeptNums_Click()
    If chkDeptNums.Value = 1 Then
        txtDeptNums.Enabled = True
    Else
        txtDeptNums.Enabled = False
    End If
End Sub

Private Sub chkDeptRegistOneNum_Click()
    If chkDeptRegistOneNum.Value = 1 Then
        txtDeptRegistOneNum.Enabled = True
        chkDeptRegistOneEmer.Enabled = True
    Else
        txtDeptRegistOneNum.Enabled = False
        chkDeptRegistOneEmer.Enabled = False
    End If
End Sub

Private Sub chkNewCardNoPop_Click()
    If chkNewCardNoPop.Value = 1 Then
        chkAutoAddName.Value = 0
        chkCardMoney.Value = 1  '不弹窗口时,卡费不能先存为划价单,因为此时未输姓名不能建档
    End If
End Sub


Private Sub chkRegistDeptNums_Click()
    If chkRegistDeptNums.Value = 1 Then
        txtRegistDeptNums.Enabled = True
    Else
        txtRegistDeptNums.Enabled = False
    End If
End Sub

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub chkAllowPatientInput_Click()
    If chkAllowPatientInput.Value = 0 Then
        chkAllowSexInput.Value = 0
        chkAllowAgeInput.Value = 0
        chkAllowAddressInput.Value = 0
        chkAllowPayModeInput.Value = 0
        chkAllowSexInput.Enabled = False
        chkAllowAgeInput.Enabled = False
        chkAllowAddressInput.Enabled = False
        chkAllowPayModeInput.Enabled = False
    Else
        chkAllowSexInput.Enabled = True And chkAllowSexInput.Tag = "1"
        chkAllowAgeInput.Enabled = True And chkAllowAgeInput.Tag = "1"
        chkAllowAddressInput.Enabled = True And chkAllowAddressInput.Tag = "1"
        chkAllowPayModeInput.Enabled = True And chkAllowPayModeInput.Tag = "1"
    End If
End Sub

 

Private Sub chkDeptBespeakOneNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub



Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    
    On Error GoTo Hd
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '数据库存储的模块参数
    '-------------------------------------------------------------------------------------------
    zlDatabase.SetPara "自动刷新间隔", IIf(chkAutoRefresh.Value = 1, Val(txtInterval.Text), 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "自动门诊号", chkAutoGet.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "存为划价单", chkPrice.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "优先使用预交款", chkPrePayPriority.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省付款方式", cboDefaultPayMode.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省费别", cboDefaultFeeType.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省结算方式", cboDefaultBalance.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省性别", cboDefaultSex.Text, glngSys, mlngModul, blnHavePrivs
    '问题号:53408
    zlDatabase.SetPara "扫描身份证签约", chkScanIDVisa.Value, glngSys, mlngModul, blnHavePrivs
    '69506
    zlDatabase.SetPara "已退序号允许挂号", chkReuseCanceledNO.Value, glngSys, mlngModul, blnHavePrivs
    '53045
    zlDatabase.SetPara "默认购买病历", chkDefaultMedBook.Value, glngSys, mlngModul, blnHavePrivs
    '问题:35176
    zlDatabase.SetPara "退号清除门诊信息", IIf(optClearInfor(0).Value, 0, IIf(optClearInfor(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    '问题:31182
    If chkDeptBespeakOneNum.Value = 1 Then
        zlDatabase.SetPara "病人同科限约N个号", Val(txtDeptBespeakOneNum.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "病人同科限约N个号", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkDeptRegistOneNum.Value = 1 Then
        zlDatabase.SetPara "病人同科限挂N个号", Val(txtDeptRegistOneNum.Text) & "|" & IIf(chkDeptRegistOneEmer.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "病人同科限挂N个号", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkDeptNums.Value = 1 Then
        zlDatabase.SetPara "病人预约科室数", Val(txtDeptNums.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "病人预约科室数", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkRegistDeptNums.Value = 1 Then
        zlDatabase.SetPara "病人挂号科室限制", Val(txtRegistDeptNums.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "病人挂号科室限制", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    zlDatabase.SetPara "预约有效时间", IIf(cbo预约有效时间.ListIndex = 0, 1, -1) * Val(txtAvailabilityTimes.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约限制时间", Val(txtBespeakDefaultDays) & "|" & Val(txtBespeakMinTime.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约失约次数", Val(txtBreakAnAppointmentNums.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约接收确定挂号费", chkBespeakFee.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "失约用于挂号", chkBreakAnAppointmentToRegist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约不生成门诊号", chkMzh.Value, glngSys, mlngModul, blnHavePrivs '36028
    '71651:刘尔旋,2014-03-31,新增参数 门诊号有效性检查
    zlDatabase.SetPara "门诊号有效性检查", chkNOValidityCheck.Value, glngSys, mlngModul, blnHavePrivs
     '问题 43847
    zlDatabase.SetPara "允许列头排序", chkRigistHeadSort.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省排序方式", IIf(cboSortMode.ListIndex = -1, 0, cboSortMode.ListIndex), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "N天内不能取消预约号", Val(txtCancelBespeak.Text), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "N岁以下必须录入监护人", Val(txtMustGuardianInfo.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "退号审核", chkBackNoToVerfy.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "随机序号选择", chkRandSelectNum.Value, glngSys, mlngModul, blnHavePrivs
     '62467
    zlDatabase.SetPara "严格按时段挂号", chkTimeRangeRegist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缺省预约方式", NeedName(cboDefaultStyle.Text), glngSys, mlngModul, blnHavePrivs
    strTmp = ""
    If lstDept.ListCount <> lstDept.SelCount Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) = True Then
                strTmp = strTmp & "," & lstDept.ItemData(i)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End If
    zlDatabase.SetPara "挂号科室", strTmp, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "姓名模糊查找", chkSeekName.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "姓名查找天数", Val(txtNameDays.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入医生", chkDoctor.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "缴款挂号结束", chkTotal.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "预约接收模式", IIf(optReceiveMode(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "零费用打印", chkPrintFree.Value, glngSys, mlngModul, blnHavePrivs
    For i = 0 To optPrintFact.UBound
        If optPrintFact(i).Value Then
            zlDatabase.SetPara "挂号发票打印方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    For i = 0 To Me.optPrepayPrint.UBound
        If optPrepayPrint(i).Value Then
            zlDatabase.SetPara "病人条码打印方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    '56274
    For i = 0 To Me.optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "预约挂号单打印方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    '68408
    For i = 0 To Me.optSlipPrint.UBound
        If optSlipPrint(i).Value Then
            zlDatabase.SetPara "挂号凭条打印方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    
    For i = 0 To optReceipt.UBound
        If optReceipt(i).Value Then
            zlDatabase.SetPara "退号回单打印方式", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    
    zlDatabase.SetPara "打印病历标签", chkPrintCase.Value, glngSys, mlngModul, blnHavePrivs
    
    '共用挂号票据批次
    strTmp = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTmp = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "共用挂号票据批次", strTmp, glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "输入姓名", chkAllowPatientInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入性别", chkAllowSexInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入年龄", chkAllowAgeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入家庭地址", chkAllowAddressInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入付款方式", chkAllowPayModeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入费别", chkAllowFeeTypeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入结算方式", chkAllowBalanceInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "输入联系电话", chkAllowPhoneInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "允许住院病人挂号", chkAllowZyRigist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "发卡不弹窗口", chkNewCardNoPop.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "自动产生姓名", chkAutoAddName.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "收取卡费", chkCardMoney.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "退费重打", IIf(chkRePrint.Enabled, chkRePrint.Value, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "家庭地址输入方式", IIf(chkAddressAssnInput.Enabled, chkAddressAssnInput.Value, 1), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "非严格控制时始终发卡", IIf(chkAlwaysSendCard.Enabled, chkAlwaysSendCard.Value, 0), glngSys, mlngModul, blnHavePrivs
    '92468:李南春,2016/1/25,严格控制下卡费为0也走票号
    zlDatabase.SetPara "零卡费走票号", IIf(ChkMustBill.Enabled, ChkMustBill.Value, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "提前挂号颜色", mstrColor, glngSys, mlngModul, blnHavePrivs
    
    Call SaveInvoice
    Call InitLocPar(mlngModul)
    gblnOk = True
    Unload Me
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Function LoadFactList(bytKind As Byte) As Boolean
'功能：读取可用公用挂号票据或就诊卡领用
'参数:bytKind=4-挂号票据,5-就诊卡
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim ObjItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("共用挂号票据批次", glngSys, mlngModul, 0, Array(lvwBill), InStr(mstrPrivs, "参数设置") > 0)
    Set rsTmp = GetShareInvoiceGroupID(bytKind)
    
    For i = 1 To rsTmp.RecordCount
        Set ObjItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!领用人)
        ObjItem.SubItems(1) = Format(rsTmp!登记时间, "yyyy-MM-dd")
        ObjItem.SubItems(2) = rsTmp!开始号码 & "," & rsTmp!终止号码
        ObjItem.SubItems(3) = rsTmp!剩余数量
        If rsTmp!ID = lngTmp Then
            ObjItem.Checked = True
            ObjItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara IIf(bytKind = 4, "共用挂号票据批次", "共用就诊卡批次"), "0", glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPrintSet_Click(Index As Integer)
    On Error GoTo Hd
    Select Case Index
    '病人条码打印
    Case 0:
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me)
    Case 1:
        '挂号收费打印
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 2:
        '预约挂号打印   '56274
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    Case 3:
        '68408,刘尔旋,2013-12-11,挂号凭条打印
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
    Case 4:
        '退号回单打印
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_3", Me)
    Case Else:
    End Select
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = False
            Next
        End If
    End If
End Sub

Private Sub Load支付方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:冉俊明
    '日期:2014-07-02
    '问题号:74552
    '说明:挂号管理中设置默认结算方式时候可以选择结算方式性质为"7-一卡通结算"的结算方式
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String

    strSQL = _
        " Select B.编码,B.名称,Nvl(B.缺省标志,0) as 缺省,Nvl(B.性质,1) as 性质,Nvl(B.应付款,0) as 应付款" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式" & _
        "   And(B.性质<>7 Or B.性质=7 And Exists(Select 1 From 一卡通目录 C Where C.结算方式=B.名称 And C.启用=1))" & _
        "   and B.性质<>8 And Instr(',1,2,7,',','||B.性质||',')>0" & _
        " Order by 性质,lpad(编码,3,' ')"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "挂号")
    
    '获取三方卡的结算方式
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    
    varData = Split(strPayType, ";")
    With cboDefaultBalance
        .Clear
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True: Exit For
                End If
            Next
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
                .ItemData(.NewIndex) = 1
                If Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称) = gstr结算方式 Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!缺省)) = 1 Then .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        '加载结算方式性质为“7-一卡通结算”的医疗卡类别
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Integer
    Dim str科室ID       As String
    Dim strTmp          As String
    Dim blnParSet       As Boolean
    Dim intIndex        As Integer
    Dim lngValue        As Long
    
    gblnOk = False
    
    blnParSet = InStr(mstrPrivs, "参数设置") > 0
    On Error GoTo errH
    'a.初始数据
    '----------------------------------------------------------------------------------------
    strSQL = "Select Distinct B.编码 ||'-'|| B.名称 as 名称,B.ID From 挂号安排 A,部门表 B Where A.科室ID=B.ID Order by 名称"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    zlControl.CboAddData lstDept, rsTmp, True
    
    strSQL = "Select '医疗付款方式' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 医疗付款方式" & _
            " Union All " & _
            " Select '性别' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别" & _
            " Union All " & _
            " Select '费别' 分类,编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别" & _
            " Where 属性=1 And Nvl(仅限初诊,0)=0 And Nvl(服务对象,3) IN(1,3)" & _
            " Order by 分类,编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '缺省医疗付款方式
    rsTmp.Filter = "分类='医疗付款方式'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultPayMode.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultPayMode.ListIndex = cboDefaultPayMode.NewIndex
        rsTmp.MoveNext
    Next
     '缺省费别    '不是仅限初诊身份唯一性项目(包含了缺省费别),不管有效期间及科室
    rsTmp.Filter = "分类='费别'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultFeeType.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultFeeType.ListIndex = cboDefaultFeeType.NewIndex
        rsTmp.MoveNext
    Next
    '缺省性别
    rsTmp.Filter = "分类='性别'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultSex.AddItem rsTmp!名称
        If rsTmp!缺省 = 1 Then cboDefaultSex.ListIndex = cboDefaultSex.NewIndex
        rsTmp.MoveNext
    Next
    cboDefaultSex.AddItem "无"
    '缺省结算方式
    Call Load支付方式

    cboSortMode.Clear
    cboSortMode.AddItem "0.号别"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 0
    cboSortMode.ListIndex = 0
    cboSortMode.AddItem "1.科室-项目"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 1
    cboSortMode.AddItem "2.科室"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 2
    
    With cbo预约有效时间
        .Clear
        .AddItem "提前"
        .AddItem "延后"
    End With
                
    
    'c.数据库存储的模块参数
    '----------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("自动刷新间隔", glngSys, mlngModul, , Array(chkAutoRefresh, txtInterval), blnParSet)
    chkAutoRefresh.Value = IIf(Val(strTmp) > 0, 1, 0)
    If chkAutoRefresh.Value = 1 Then txtInterval.Text = strTmp
    chkAutoGet.Value = IIf(zlDatabase.GetPara("自动门诊号", glngSys, mlngModul, , Array(chkAutoGet), blnParSet) = "1", 1, 0)
    chkPrice.Value = IIf(zlDatabase.GetPara("存为划价单", glngSys, mlngModul, , Array(chkPrice), blnParSet) = "1", 1, 0)
    chkPrePayPriority.Value = IIf(zlDatabase.GetPara("优先使用预交款", glngSys, mlngModul, , Array(chkPrePayPriority), blnParSet) = "1", 1, 0)
    chkRigistHeadSort.Value = IIf(zlDatabase.GetPara("允许列头排序", glngSys, mlngModul, , Array(chkRigistHeadSort), blnParSet) = "1", 1, 0)
    chkRandSelectNum.Value = IIf(zlDatabase.GetPara("随机序号选择", glngSys, mlngModul, , Array(chkRandSelectNum), blnParSet) = "1", 1, 0)
    chkBreakAnAppointmentToRegist.Value = IIf(zlDatabase.GetPara("失约用于挂号", glngSys, mlngModul, 0, Array(chkBreakAnAppointmentToRegist), blnParSet) = "1", 1, 0)
    chkBackNoToVerfy.Value = IIf(zlDatabase.GetPara("退号审核", glngSys, mlngModul, 0, Array(chkBackNoToVerfy), blnParSet) = "1", 1, 0)
    chkAddressAssnInput.Value = IIf(zlDatabase.GetPara("家庭地址输入方式", glngSys, mlngModul, 1, Array(chkAddressAssnInput), blnParSet) = "1", 1, 0)
    chkAlwaysSendCard.Value = IIf(zlDatabase.GetPara("非严格控制时始终发卡", glngSys, mlngModul, 1, Array(chkAlwaysSendCard), blnParSet) = "1", 1, 0)
    '69506，刘尔旋,2014-01-14,新增参数"已退序号允许挂号"
    chkReuseCanceledNO.Value = IIf(zlDatabase.GetPara("已退序号允许挂号", glngSys, mlngModul, 1, Array(chkReuseCanceledNO), blnParSet) = "1", 1, 0)
    chkNOValidityCheck.Value = IIf(zlDatabase.GetPara("门诊号有效性检查", glngSys, mlngModul, 1, Array(chkNOValidityCheck), blnParSet) = "1", 1, 0)
    '53045，刘尔旋,2014-02-13,默认勾选购买病历选项
    chkDefaultMedBook.Value = IIf(zlDatabase.GetPara("默认购买病历", glngSys, mlngModul, 0, Array(chkDefaultMedBook), blnParSet) = "1", 1, 0)
    strTmp = zlDatabase.GetPara("缺省排序方式", glngSys, glngModul, , Array(cboSortMode), blnParSet)
    cboSortMode.ListIndex = Val(strTmp)
    strTmp = zlDatabase.GetPara("缺省付款方式", glngSys, mlngModul, , Array(cboDefaultPayMode), blnParSet)
    zlControl.CboLocate cboDefaultPayMode, strTmp
    strTmp = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , Array(cboDefaultFeeType), blnParSet)
    zlControl.CboLocate cboDefaultFeeType, strTmp
    strTmp = zlDatabase.GetPara("缺省性别", glngSys, mlngModul, , Array(cboDefaultSex), blnParSet)
    zlControl.CboLocate cboDefaultSex, strTmp
    If cboDefaultSex.ListIndex = -1 Or strTmp = "无" Then cboDefaultSex.ListIndex = cboDefaultSex.ListCount - 1
    strTmp = zlDatabase.GetPara("缺省结算方式", glngSys, mlngModul, , Array(cboDefaultBalance), blnParSet)
    zlControl.CboLocate cboDefaultBalance, strTmp
    '问题号:53408
    chkScanIDVisa.Value = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, mlngModul, 0, Array(chkScanIDVisa), blnParSet) = "1", 1, 0)
    '问题:35176
    strTmp = zlDatabase.GetPara("退号清除门诊信息", glngSys, mlngModul, , Array(fraClearMZInfor, optClearInfor(0), optClearInfor(1), optClearInfor(2)), blnParSet)
    If Val(strTmp) = 0 Then
        optClearInfor(0).Value = True
    ElseIf Val(strTmp) = 1 Then
        optClearInfor(1).Value = True
    Else
        optClearInfor(2).Value = True
    End If
    
    strTmp = zlDatabase.GetPara("预约接收模式", glngSys, mlngModul, , Array(fraReceiveMode, optReceiveMode(0), optReceiveMode(1)), blnParSet)
    If Val(strTmp) = 0 Then
        optReceiveMode(0).Value = True
    Else
        optReceiveMode(1).Value = True
    End If
    
    mstrColor = zlDatabase.GetPara("提前挂号颜色", glngSys, 1111, "", , blnParSet)
    If mstrColor = "" Then mstrColor = &H0&
    pic提前颜色.BackColor = mstrColor
    
    strTmp = zlDatabase.GetPara("缺省预约方式", glngSys, 1115, "", Array(cboDefaultStyle), True)
    strSQL = "Select 编码,名称,缺省标志 From 预约方式 Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboDefaultStyle.Clear
    Do While Not rsTmp.EOF
        cboDefaultStyle.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If strTmp = Nvl(rsTmp!名称) Then intIndex = cboDefaultStyle.NewIndex
        If Val(Nvl(rsTmp!缺省标志)) = 1 Then cboDefaultStyle.ListIndex = cboDefaultStyle.NewIndex
        rsTmp.MoveNext
    Loop
    If cboDefaultStyle.ListCount <> 0 And intIndex <> 0 Then cboDefaultStyle.ListIndex = intIndex
    
    '62467
    chkTimeRangeRegist.Value = IIf(zlDatabase.GetPara("严格按时段挂号", glngSys, mlngModul, 0, Array(chkTimeRangeRegist), blnParSet) = "1", 1, 0)
    
    '问题:31182
    txtDeptBespeakOneNum.Text = Val(zlDatabase.GetPara("病人同科限约N个号", glngSys, mlngModul, 0, Array(chkDeptBespeakOneNum, txtDeptBespeakOneNum), blnParSet))
    If Val(txtDeptBespeakOneNum.Text) = 0 Then
        chkDeptBespeakOneNum.Value = 0
        txtDeptBespeakOneNum.Text = ""
        txtDeptBespeakOneNum.Enabled = False
    Else
        chkDeptBespeakOneNum.Value = 1
        txtDeptBespeakOneNum.Enabled = True
    End If
    
    txtDeptRegistOneNum.Text = Val(Split(zlDatabase.GetPara("病人同科限挂N个号", glngSys, mlngModul, 0, Array(chkDeptRegistOneNum, txtDeptRegistOneNum), blnParSet) & "|", "|")(0))
    If Val(txtDeptRegistOneNum.Text) = 0 Then
        chkDeptRegistOneNum.Value = 0
        chkDeptRegistOneEmer.Value = 0
        chkDeptRegistOneEmer.Enabled = False
        txtDeptRegistOneNum.Text = ""
        txtDeptRegistOneNum.Enabled = False
    Else
        chkDeptRegistOneNum.Value = 1
        chkDeptRegistOneEmer.Value = Val(Split(zlDatabase.GetPara("病人同科限挂N个号", glngSys, mlngModul, 0, Array(chkDeptRegistOneNum, txtDeptRegistOneNum), blnParSet) & "|", "|")(1))
        chkDeptRegistOneEmer.Enabled = True
        txtDeptRegistOneNum.Enabled = True
    End If
    
    txtDeptNums.Text = Val(zlDatabase.GetPara("病人预约科室数", glngSys, mlngModul, 0, Array(txtDeptNums, chkDeptNums), blnParSet))
    If Val(txtDeptNums.Text) = 0 Then
        chkDeptNums.Value = 0
        txtDeptNums.Text = ""
        txtDeptNums.Enabled = False
    Else
        chkDeptNums.Value = 1
        txtDeptNums.Enabled = True
    End If
    txtRegistDeptNums.Text = Val(zlDatabase.GetPara("病人挂号科室限制", glngSys, mlngModul, 0, Array(chkRegistDeptNums, txtRegistDeptNums), blnParSet))
    If Val(txtRegistDeptNums.Text) = 0 Then
        chkRegistDeptNums.Value = 0
        txtRegistDeptNums.Text = ""
        txtRegistDeptNums.Enabled = False
    Else
        chkRegistDeptNums.Value = 1
        txtRegistDeptNums.Enabled = True
    End If
    
    lngValue = Val(zlDatabase.GetPara("预约有效时间", glngSys, mlngModul, 0, Array(txtAvailabilityTimes, lblAvailabilityTimes), blnParSet))
    
    If lngValue >= 0 Then
        cbo预约有效时间.ListIndex = 0
    Else
        cbo预约有效时间.ListIndex = 1
    End If
    txtAvailabilityTimes.Text = Abs(lngValue)
    
    txtBreakAnAppointmentNums.Text = Val(zlDatabase.GetPara("预约失约次数", glngSys, mlngModul, 0, Array(txtBreakAnAppointmentNums, lblBreakAnAppointmentNums), blnParSet))
    chkBespeakFee.Value = IIf(zlDatabase.GetPara("预约接收确定挂号费", glngSys, mlngModul, 0, Array(chkBespeakFee), blnParSet) = "1", 1, 0)
    chkBreakAnAppointmentToRegist.Value = IIf(zlDatabase.GetPara("失约用于挂号", glngSys, mlngModul, 0, Array(chkBreakAnAppointmentToRegist), blnParSet) = "1", 1, 0)
    txtCancelBespeak.Text = Val(zlDatabase.GetPara("N天内不能取消预约号", glngSys, mlngModul, 0, Array(txtCancelBespeak, lblCancelBespeak), blnParSet))
    Call txtCancelBespeak_Change
    txtMustGuardianInfo.Text = Val(zlDatabase.GetPara("N岁以下必须录入监护人", glngSys, mlngModul, 0, Array(txtMustGuardianInfo, lblGuardian), blnParSet))
    
    strTmp = zlDatabase.GetPara("预约限制时间", glngSys, mlngModul, "1|60", Array(txtBespeakMinTime, lblBespeakMinTime, lblBespeakDefaultDays, txtBespeakDefaultDays), blnParSet)
    txtBespeakMinTime.Text = Val(Split(strTmp, "|")(1))
    txtBespeakDefaultDays.Text = Val(Split(strTmp, "|")(0))
    '读取可用的挂号科室
    str科室ID = zlDatabase.GetPara("挂号科室", glngSys, mlngModul, , Array(lstDept), blnParSet)
    If str科室ID = "" Then
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = True
        Next
    Else
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = InStr(1, "," & str科室ID & ",", "," & lstDept.ItemData(i) & ",") > 0
        Next
    End If
    If lstDept.ListCount > 0 Then lstDept.TopIndex = 0: lstDept.ListIndex = 0
    
    
    txtNameDays.Text = zlDatabase.GetPara("姓名查找天数", glngSys, mlngModul, , Array(txtNameDays), blnParSet)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul, , Array(chkSeekName), blnParSet) = "1", 1, 0)
    
    chkDoctor.Value = IIf(zlDatabase.GetPara("输入医生", glngSys, mlngModul, , Array(chkDoctor), blnParSet) = "1", 1, 0)
    chkTotal.Value = IIf(zlDatabase.GetPara("缴款挂号结束", glngSys, mlngModul, , Array(chkTotal), blnParSet) = "1", 1, 0)
    chkPrintFree.Value = IIf(zlDatabase.GetPara("零费用打印", glngSys, mlngModul, , Array(chkPrintFree), blnParSet) = "1", 1, 0)
    
    i = Val(zlDatabase.GetPara("挂号发票打印方式", glngSys, mlngModul, 1, Array(optPrintFact(0), optPrintFact(1), optPrintFact(2)), blnParSet))
    If i <= optPrintFact.UBound Then optPrintFact(i).Value = True
    i = Val(zlDatabase.GetPara("病人条码打印方式", glngSys, mlngModul, 1, Array(optPrepayPrint(0), optPrepayPrint(1), optPrepayPrint(2), cmdPrintSet(0)), blnParSet))
    If i <= optPrepayPrint.UBound Then optPrepayPrint(i).Value = True
    
    '问题:56274
    i = Val(zlDatabase.GetPara("预约挂号单打印方式", glngSys, mlngModul, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2), cmdPrintSet(2)), blnParSet))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
    '68408
    i = Val(zlDatabase.GetPara("挂号凭条打印方式", glngSys, mlngModul, 1, Array(optSlipPrint(0), optSlipPrint(1), optSlipPrint(2), cmdPrintSet(3)), blnParSet))
    If i <= optSlipPrint.UBound Then optSlipPrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("退号回单打印方式", glngSys, mlngModul, 1, Array(optReceipt(0), optReceipt(1), optReceipt(2), cmdPrintSet(4)), blnParSet))
    If i <= optSlipPrint.UBound Then optReceipt(i).Value = True
    
    i = Val(zlDatabase.GetPara("打印病历标签", glngSys, mlngModul, 0, Array(chkPrintCase), blnParSet))
    chkPrintCase.Value = i
    chkAllowPatientInput.Value = IIf(zlDatabase.GetPara("输入姓名", glngSys, mlngModul, , Array(chkAllowPatientInput), blnParSet) = "1", 1, 0)
    chkAllowSexInput.Enabled = True
    chkAllowSexInput.Value = IIf(zlDatabase.GetPara("输入性别", glngSys, mlngModul, , Array(chkAllowSexInput), blnParSet) = "1", 1, 0)
    chkAllowSexInput.Tag = IIf(chkAllowSexInput.Enabled, "1", "0")
    chkAllowAgeInput.Enabled = True
    chkAllowAgeInput.Value = IIf(zlDatabase.GetPara("输入年龄", glngSys, mlngModul, , Array(chkAllowAgeInput), blnParSet) = "1", 1, 0)
    chkAllowAgeInput.Tag = IIf(chkAllowAgeInput.Enabled, "1", "0")
    chkAllowAddressInput.Enabled = True
    chkAllowAddressInput.Value = IIf(zlDatabase.GetPara("输入家庭地址", glngSys, mlngModul, , Array(chkAllowAddressInput), blnParSet) = "1", 1, 0)
    chkAllowAddressInput.Tag = IIf(chkAllowAddressInput.Enabled, "1", "0")
    chkAllowPayModeInput.Enabled = True
    chkAllowPayModeInput.Value = IIf(zlDatabase.GetPara("输入付款方式", glngSys, mlngModul, , Array(chkAllowPayModeInput), blnParSet) = "1", 1, 0)
    chkAllowPayModeInput.Tag = IIf(chkAllowPayModeInput.Enabled, "1", "0")
    
    chkAllowFeeTypeInput.Value = IIf(zlDatabase.GetPara("输入费别", glngSys, mlngModul, , Array(chkAllowFeeTypeInput), blnParSet) = "1", 1, 0)
    '31724
    chkAllowZyRigist.Value = IIf(zlDatabase.GetPara("允许住院病人挂号", glngSys, mlngModul, , Array(chkAllowZyRigist), blnParSet) = "1", 1, 0)
    chkAllowBalanceInput.Value = IIf(zlDatabase.GetPara("输入结算方式", glngSys, mlngModul, , Array(chkAllowBalanceInput), blnParSet) = "1", 1, 0)
    chkAllowPhoneInput.Value = IIf(zlDatabase.GetPara("输入联系电话", glngSys, mlngModul, , Array(chkAllowPhoneInput), blnParSet) = "1", 1, 0)
    Call chkAllowPatientInput_Click
    
    '读取可用公用挂号票据领用
    Call LoadFactList(4)
    
            
    chkNewCardNoPop.Value = IIf(zlDatabase.GetPara("发卡不弹窗口", glngSys, mlngModul, , Array(chkNewCardNoPop), blnParSet) = "1", 1, 0)
    chkAutoAddName.Value = IIf(zlDatabase.GetPara("自动产生姓名", glngSys, mlngModul, , Array(chkAutoAddName), blnParSet) = "1", 1, 0)
    
    
    chkRePrint.Value = IIf(zlDatabase.GetPara("退费重打", glngSys, mlngModul, , Array(chkRePrint), blnParSet) = "1", 1, 0)
    
    chkCardMoney.Value = IIf(zlDatabase.GetPara("收取卡费", glngSys, mlngModul, , Array(chkCardMoney), blnParSet) = "1", 1, 0)
    Call chkCardMoney_Click
    '36028
    chkMzh.Value = IIf(zlDatabase.GetPara("预约不生成门诊号", glngSys, mlngModul, , Array(chkMzh), blnParSet) = "1", 1, 0)
    '92468:李南春,2016/1/25,严格控制下卡费为0也走票号
    ChkMustBill.Value = IIf(zlDatabase.GetPara("零卡费走票号", glngSys, mlngModul, 0, Array(ChkMustBill), blnParSet) = "1", 1, 0)
    
    
    '挂号列表过滤条件,是否已发生时间为准
'    ‘chkRegList.Value = IIf(zlDatabase.GetPara("按发生时间显示记录", glngSys, mlngModul, , Array(chkRegList), blnParSet) = "1", 1, 0)

    '读取公用的就诊卡领用
     Call InitShareInvoice
    If Tabs1.TabVisible(0) Then Tabs1.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub

Private Sub pic提前颜色_Click()
    dlgColor.ShowColor
    mstrColor = dlgColor.Color
    pic提前颜色.BackColor = mstrColor
End Sub

Private Sub txtCancelBespeak_Change()
    '问题号:56407
    If Val(txtCancelBespeak.Text) = 0 Then
        chkBackNoToVerfy.Caption = "当天内取消预约需要通过审核"
    Else
        chkBackNoToVerfy.Caption = "在" & txtCancelBespeak.Text & "天内取消预约需要通过审核"
    End If
End Sub

Private Sub txtInterval_GotFocus()
    Call SelAll(txtInterval)
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtInterval_Validate(Cancel As Boolean)
    If Val(txtInterval.Text) < 1 Then
        txtInterval.Text = 1
    ElseIf Val(txtInterval.Text) > 99 Then
        txtInterval.Text = 99
    End If
End Sub

Private Sub txtMustGuardianInfo_GotFocus()
    zlControl.TxtSelAll txtMustGuardianInfo
End Sub

Private Sub txtMustGuardianInfo_KeyPress(KeyAscii As Integer)
     If KeyAscii = Asc("-") Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txtMustGuardianInfo, KeyAscii, m数字式
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

Private Sub txtCancelBespeak_GotFocus()
    zlControl.TxtSelAll txtCancelBespeak
End Sub

Private Sub txtDeptNums_GotFocus()
    zlControl.TxtSelAll txtDeptNums
End Sub

Private Sub txtCancelBespeak_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCancelBespeak_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCancelBespeak, KeyAscii, m数字式
End Sub

Private Sub txtDeptNums_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtDeptNums_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtDeptNums, KeyAscii, m数字式
End Sub

Private Sub txtBreakAnAppointmentNums_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAvailabilityTimes_GotFocus()
    zlControl.TxtSelAll txtAvailabilityTimes
End Sub

Private Sub txtAvailabilityTimes_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAvailabilityTimes_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtAvailabilityTimes, KeyAscii, m数字式
End Sub
Private Sub txtBreakAnAppointmentNums_GotFocus()
    zlControl.TxtSelAll txtBreakAnAppointmentNums
End Sub

Private Sub txtBreakAnAppointmentNums_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBreakAnAppointmentNums, KeyAscii, m数字式
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置共享发票
    '编制:刘兴洪
    '日期:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '共享票据批次,格式:批次,批次
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    Dim lngTemp As Long, i As Long, strSQL As String, rs医疗卡类别 As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim str缺省医疗卡 As String, lng缺省医疗卡 As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True, intType))
    gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  Where nvl(是否启用,0)=1"
    On Error GoTo Hd
    Set rs医疗卡类别 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rs医疗卡类别.Filter = "名称='就诊卡' and 是否固定=1"
    If rs医疗卡类别.EOF = False Then
        str缺省医疗卡 = rs医疗卡类别!名称: lng缺省医疗卡 = Val(rs医疗卡类别!ID)
    End If
    With rs医疗卡类别
        cboType.Clear
        rs医疗卡类别.Filter = 0
        If rs医疗卡类别.RecordCount <> 0 Then rs医疗卡类别.MoveFirst
        Do While Not .EOF
            cboType.AddItem Nvl(!名称)
            cboType.ItemData(cboType.NewIndex) = Nvl(!ID)
            If Nvl(!名称) = "就诊卡" And cboType.ListIndex < 0 Then cboType.ListIndex = cboType.NewIndex
            If lngCardTypeID = Val(Nvl(!ID)) Then
                cboType.ListIndex = cboType.NewIndex
            End If
            .MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
    strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, , , True)
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
    
    '格式:领用ID1,医疗卡类别ID1|领用IDn,医疗卡类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!使用类别ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = str缺省医疗卡
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = lng缺省医疗卡
            Else
                rs医疗卡类别.Filter = "ID=" & Val(Nvl(rsTemp!使用类别ID))
                If Not rs医疗卡类别.EOF Then
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rs医疗卡类别!名称)
                Else
                    .TextMatrix(lngRow, .ColIndex("医疗卡类别")) = Nvl(rsTemp!使用类别)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("医疗卡类别")) = Val(Nvl(rsTemp!使用类别ID))
            End If
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("医疗卡类别"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long, lng卡类别ID As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '保存共享票据
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("医疗卡类别")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用医疗卡批次", strValue, glngSys, mlngModul, blnHavePrivs
    If cboType.ListIndex >= 0 Then
        lng卡类别ID = cboType.ItemData(cboType.ListIndex)
    End If
    Call zlDatabase.SetPara("缺省医疗卡类别", lng卡类别ID, glngSys, mlngModul, blnHavePrivs)
End Sub

Private Sub txtBespeakMinTime_GotFocus()
    zlControl.TxtSelAll txtBespeakMinTime
End Sub
Private Sub txtBespeakMinTime_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBespeakMinTime_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBespeakMinTime, KeyAscii, m数字式
End Sub
 
Private Sub txtBespeakDefaultDays_GotFocus()
    zlControl.TxtSelAll txtBespeakDefaultDays
End Sub
Private Sub txtBespeakDefaultDays_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBespeakDefaultDays_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBespeakDefaultDays, KeyAscii, m数字式
End Sub

'Public Sub ShowParSet(ByVal frmMain As Object, ByRef blnCancel As Boolean)
'    '显示参数设置
'    mblnOK = False
'    If frmMain Is Nothing Then
'        Me.Show 1
'    Else
'        Me.Show 1, frmMain
'    End If
'
'    blnCancel = Not mblnOK
'End Sub
