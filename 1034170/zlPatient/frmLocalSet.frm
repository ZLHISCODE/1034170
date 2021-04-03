VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sTab 
      Height          =   5925
      Left            =   90
      TabIndex        =   63
      Top             =   60
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10451
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "输入控制(&1)"
      TabPicture(0)   =   "frmLocalSet.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInput"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkAutoRefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "医疗卡票据控制(&2)"
      TabPicture(1)   =   "frmLocalSet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDefaultPayCard"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "img16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkMustCard"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDeviceSetup(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkCardFeeCharge"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraTitle"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkBruhCardBackCard"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkBrushCardVerfy"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cboType"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkScanIDPatiVisa"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "预交款设置(&3)"
      TabPicture(2)   =   "frmLocalSet.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblEdit"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "vs代收"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cboDefaultBalance"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fra退款设置"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkNotClearPatiInfor"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "chkNotInDeptNotJk"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkAdvance"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkSeekName"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkVeryfyInDeposit"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "预交票据控制(&4)"
      TabPicture(3)   =   "frmLocalSet.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fra红票格式"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "fraPrepay"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "chkAllowDept"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "chkHave"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdDeviceSetup(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "chkLedWelcome"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "chkCheckBillNum"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txt票据张数"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "upd票据张数"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "fra票据格式"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).ControlCount=   10
      Begin VB.Frame fra票据格式 
         Caption         =   "预交票据格式"
         Height          =   1305
         Left            =   90
         TabIndex        =   53
         Top             =   2205
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   6435
            _cx             =   11351
            _cy             =   1773
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
            FormatString    =   $"frmLocalSet.frx":0070
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
      Begin VB.CheckBox chkVeryfyInDeposit 
         Caption         =   "退住院预交刷卡验证"
         Height          =   300
         Left            =   -71235
         TabIndex        =   65
         Top             =   5445
         Width           =   2340
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "是否允许通过姓名模糊查找病人"
         Height          =   180
         Left            =   -74760
         TabIndex        =   64
         ToolTipText     =   "缴预交时输入姓名是否模糊查找病人"
         Top             =   5505
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chkAdvance 
         Caption         =   "允许出院病人缴住院预交"
         Height          =   300
         Left            =   -71235
         TabIndex        =   50
         Top             =   5145
         Width           =   2340
      End
      Begin VB.CheckBox chkNotInDeptNotJk 
         Caption         =   "在院病人未入科不准收预交"
         Height          =   300
         Left            =   -74745
         TabIndex        =   49
         Top             =   5145
         Width           =   2475
      End
      Begin VB.CheckBox chkAutoRefresh 
         Caption         =   "切换病人类型选项卡时，自动刷新病人数据"
         Height          =   180
         Left            =   -73590
         TabIndex        =   31
         Top             =   4170
         Width           =   3840
      End
      Begin MSComCtl2.UpDown upd票据张数 
         Height          =   300
         Left            =   1755
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   4935
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txt票据张数"
         BuddyDispid     =   196615
         OrigLeft        =   1500
         OrigTop         =   3285
         OrigRight       =   1755
         OrigBottom      =   3570
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt票据张数 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1275
         TabIndex        =   57
         Text            =   "10"
         Top             =   4935
         Width           =   480
      End
      Begin VB.CheckBox chkCheckBillNum 
         Caption         =   "票据剩余         张时开始提醒收费员"
         Height          =   285
         Left            =   255
         TabIndex        =   56
         Top             =   4950
         Width           =   3450
      End
      Begin VB.CheckBox chkScanIDPatiVisa 
         Caption         =   "扫描身份证签约"
         Height          =   180
         Left            =   -74640
         TabIndex        =   38
         Top             =   5025
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkNotClearPatiInfor 
         Caption         =   "缴预交后不清除界面信息"
         Height          =   300
         Left            =   -71235
         TabIndex        =   48
         Top             =   4785
         Width           =   2340
      End
      Begin VB.Frame fra退款设置 
         Caption         =   "退款设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74865
         TabIndex        =   43
         Top             =   3930
         Width           =   6510
         Begin VB.OptionButton optCheck 
            Caption         =   "余额不足时禁止退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3600
            TabIndex        =   45
            Top             =   315
            Value           =   -1  'True
            Width           =   2220
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "余额不足时提醒是否退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   495
            TabIndex        =   44
            Top             =   315
            Width           =   2625
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73455
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   5355
         Width           =   2580
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   4785
         TabIndex        =   62
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   5580
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkBrushCardVerfy 
         Caption         =   "退卡获取单据号后刷卡验证退卡"
         Height          =   180
         Left            =   -71235
         TabIndex        =   35
         Top             =   4455
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CheckBox chkBruhCardBackCard 
         Caption         =   "发卡按“退”刷卡退卡"
         Height          =   240
         Left            =   -71250
         TabIndex        =   37
         Top             =   4725
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Index           =   1
         Left            =   4980
         TabIndex        =   61
         Top             =   4995
         Width           =   1500
      End
      Begin VB.CheckBox chkHave 
         Caption         =   "只显示有剩余的历史缴款"
         Height          =   195
         Left            =   255
         TabIndex        =   59
         Top             =   5310
         Width           =   3120
      End
      Begin VB.CheckBox chkAllowDept 
         Caption         =   "允许更改病人的缴款科室"
         Height          =   195
         Left            =   255
         TabIndex        =   60
         Top             =   5595
         Value           =   1  'Checked
         Width           =   2280
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用票据"
         Height          =   1740
         Left            =   90
         TabIndex        =   51
         Top             =   435
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1455
            Left            =   60
            TabIndex        =   52
            Top             =   225
            Width           =   6480
            _cx             =   11430
            _cy             =   2566
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
            FormatString    =   $"frmLocalSet.frx":00FE
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
      Begin VB.ComboBox cboDefaultBalance 
         Height          =   300
         Left            =   -73635
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   4770
         Width           =   1875
      End
      Begin VB.Frame fraInput 
         Caption         =   "输入光标经过项目"
         Height          =   3270
         Left            =   -74265
         TabIndex        =   3
         Top             =   750
         Width           =   5190
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人身份证号"
            Height          =   195
            Index           =   26
            Left            =   3450
            TabIndex        =   24
            Top             =   930
            Value           =   1  'Checked
            Width           =   1560
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "籍贯"
            Height          =   195
            Index           =   25
            Left            =   3450
            TabIndex        =   30
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "户口地址邮编"
            Height          =   195
            Index           =   24
            Left            =   1785
            TabIndex        =   20
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "户口地址"
            Height          =   195
            Index           =   23
            Left            =   1785
            TabIndex        =   19
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "区域"
            Height          =   195
            Index           =   22
            Left            =   1785
            TabIndex        =   21
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "其他证件"
            Height          =   195
            Index           =   21
            Left            =   285
            TabIndex        =   11
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位开户行"
            Height          =   195
            Index           =   20
            Left            =   3450
            TabIndex        =   28
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位邮编"
            Height          =   195
            Index           =   19
            Left            =   3450
            TabIndex        =   27
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位电话"
            Height          =   195
            Index           =   18
            Left            =   3450
            TabIndex        =   26
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "工作单位"
            Height          =   195
            Index           =   17
            Left            =   3450
            TabIndex        =   25
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人电话"
            Height          =   195
            Index           =   16
            Left            =   3450
            TabIndex        =   23
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人地址"
            Height          =   195
            Index           =   15
            Left            =   3450
            TabIndex        =   22
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人关系"
            Height          =   195
            Index           =   14
            Left            =   1785
            TabIndex        =   18
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人姓名"
            Height          =   195
            Index           =   13
            Left            =   1785
            TabIndex        =   17
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "家庭电话"
            Height          =   195
            Index           =   12
            Left            =   1785
            TabIndex        =   16
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "家庭地址邮编"
            Height          =   195
            Index           =   11
            Left            =   1785
            TabIndex        =   15
            Top             =   930
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "现住址"
            Height          =   195
            Index           =   10
            Left            =   1785
            TabIndex        =   14
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "出生地点"
            Height          =   195
            Index           =   9
            Left            =   1785
            TabIndex        =   13
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "身份证号"
            Height          =   195
            Index           =   8
            Left            =   285
            TabIndex        =   12
            Top             =   2820
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "出生日期"
            Height          =   195
            Index           =   7
            Left            =   285
            TabIndex        =   10
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "身份"
            Height          =   195
            Index           =   6
            Left            =   285
            TabIndex        =   9
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "职业"
            Height          =   195
            Index           =   5
            Left            =   285
            TabIndex        =   8
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "婚姻状况"
            Height          =   195
            Index           =   4
            Left            =   285
            TabIndex        =   7
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "学历"
            Height          =   195
            Index           =   3
            Left            =   285
            TabIndex        =   6
            Top             =   930
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "民族"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   5
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "国籍"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   4
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位帐号"
            Height          =   195
            Index           =   0
            Left            =   3450
            TabIndex        =   29
            Top             =   2475
            Value           =   1  'Checked
            Width           =   1200
         End
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用..."
         Height          =   3825
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   6435
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   3405
            Left            =   60
            TabIndex        =   33
            Top             =   300
            Width           =   6300
            _cx             =   11112
            _cy             =   6006
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
            FormatString    =   $"frmLocalSet.frx":01DB
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
      Begin VB.CheckBox chkCardFeeCharge 
         Caption         =   "就诊卡费用以记账方式收取"
         Height          =   180
         Left            =   -74640
         TabIndex        =   36
         Top             =   4755
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Index           =   0
         Left            =   -70290
         TabIndex        =   41
         Top             =   5325
         Width           =   1500
      End
      Begin VB.CheckBox chkMustCard 
         Caption         =   "建档同时必须发卡"
         Height          =   255
         Left            =   -74640
         TabIndex        =   34
         Top             =   4425
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -71145
         Top             =   1155
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   1
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLocalSet.frx":02BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vs代收 
         Height          =   3315
         Left            =   -74865
         TabIndex        =   42
         Top             =   465
         Width           =   6525
         _cx             =   11509
         _cy             =   5847
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLocalSet.frx":039E
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   3
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
      Begin VB.Frame fra红票格式 
         Caption         =   "预交红票格式"
         Height          =   1305
         Left            =   90
         TabIndex        =   66
         Top             =   3570
         Width           =   6615
         Begin VSFlex8Ctl.VSFlexGrid vsRedBillFormat 
            Height          =   1005
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   6435
            _cx             =   11351
            _cy             =   1773
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
            FormatString    =   $"frmLocalSet.frx":03FF
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
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "缺省发卡类型"
         Height          =   210
         Left            =   -74625
         TabIndex        =   39
         Top             =   5400
         Width           =   1290
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "缺省结算方式"
         Height          =   180
         Left            =   -74775
         TabIndex        =   46
         Top             =   4830
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7065
      TabIndex        =   2
      Top             =   4410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7050
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7050
      TabIndex        =   0
      Top             =   360
      Width           =   1100
   End
End
Attribute VB_Name = "frmLocalSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlngModul As Long, mstrPrivs As String, mbln担保 As Boolean
Private mstrClass As String, mstrDeposit As String
Private mblnOK As Boolean
Public Function zlSetPara(ByVal frmMain As Object, ByVal strPrivs As String, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置
    '入参:mlngModul-1101-病人信息管理,1102-就诊卡管理,1103-预交款管理
    '出参:
    '返回:保存,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 14:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mstrPrivs = strPrivs: mlngModul = lngModule
    mbln担保 = InStr(mstrPrivs, ";担保信息;") > 0 And mlngModul = 1101
    Me.Show 1, frmMain
    zlSetPara = True
End Function
Private Sub cboType_Click()
    chkScanIDPatiVisa.Enabled = Not (cboType.Text = "二代身份证")
    If cboType.Text = "二代身份证" Then
        chkScanIDPatiVisa.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDeviceSetup_Click(Index As Integer)
    Call zlCommFun.DeviceSetup(Me, 100, mlngModul)
End Sub

Private Sub cmdHelp_Click()
    Select Case mlngModul
        Case 1101 '病人信息
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet1"
        Case 1102 '就诊卡
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet2"
        Case 1103 '预交款
            ShowHelp App.ProductName, Me.hWnd, "frmLocalSet3"
    End Select
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据有效性检查
    '返回:检查合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-06 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str类别 As String
    IsValied = False
    
    On Error GoTo errHandle
    If mlngModul <> 1103 Then
        '检查每种使用种式只能一个选择
        With vsBill
            str类别 = "-"
            For i = 1 To vsBill.Rows - 1
                If str类别 <> Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) Then
                   str类别 = Trim(.TextMatrix(i, .ColIndex("医疗卡类别")))
                   lngSelCount = 0
                    For j = 1 To vsBill.Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("医疗卡类别"))) = Trim(.TextMatrix(j, .ColIndex("医疗卡类别"))) Then
                            If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                                lngSelCount = lngSelCount + 1
                            End If
                        End If
                    Next
                    If lngSelCount > 1 Then
                        MsgBox "注意:" & vbCrLf & "    医疗卡类别为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    If mlngModul = 1102 Then IsValied = True: Exit Function
  '检查每种使用预交只能一个选择
    With vsPrepay
        str类别 = "-"
        For i = 1 To .Rows - 1
            If str类别 <> Trim(.TextMatrix(i, .ColIndex("预交类型"))) Then
               str类别 = Trim(.TextMatrix(i, .ColIndex("预交类型")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("预交类型"))) = Trim(.TextMatrix(j, .ColIndex("预交类型"))) Then
                        If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "注意:" & vbCrLf & "    预交类型为『" & str类别 & "』的只能选择一种票据,请检查!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    If mlngModul <> 1103 Then
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
    End If
    If mlngModul = 1102 Then Exit Sub
    
    
    '保存预交票据
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("预交类型")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "共用预交票据批次", strValue, glngSys, mlngModul, blnHavePrivs
    
    '61808:刘鹏飞,2013-05-21,只有预交款管理本参数才有效
    '78751:李南春,2015/08/24,增加预交票据打印格式
    If mlngModul = 1103 Or mlngModul = 1101 Then
        '问题号:50656
        Dim strPrintMode As String
        '保存收费格式
        strValue = "": strPrintMode = ""
        With vsBillFormat
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
                End If
            Next
            If strValue <> "" Then strValue = Mid(strValue, 2)
            If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
            zlDatabase.SetPara "预交发票格式", strValue, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "预交发票打印方式", strPrintMode, glngSys, mlngModul, blnHavePrivs
        End With
        '红票格式
        strValue = "": strPrintMode = ""
        With vsRedBillFormat
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("使用类别"))) <> "" Then
                    strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(.TextMatrix(i, .ColIndex("票据格式")))
                    strPrintMode = strPrintMode & "|" & Trim(.Cell(flexcpData, i, .ColIndex("使用类别"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("预交打印方式")), 1))
                End If
            Next
            If strValue <> "" Then strValue = Mid(strValue, 2)
            If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
            zlDatabase.SetPara "退款发票格式", strValue, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "预交退款打印方式", strPrintMode, glngSys, mlngModul, blnHavePrivs
        End With
    End If
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
    Dim strBillFormat As String
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    On Error GoTo errHandle
    '恢复列宽度
    If mlngModul <> 1103 Then
            lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , , True, intType))
            
            gstrSQL = "Select ID,编码,名称, nvl(是否固定,0) as 是否固定  from 医疗卡类别  Where nvl(是否启用,0)=1"
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
                    If Nvl(!名称) = "就诊卡" Then cboType.ListIndex = cboType.NewIndex
                    If lngCardTypeID = Val(Nvl(!ID)) Then
                        cboType.ListIndex = cboType.NewIndex
                    End If
                    .MoveNext
                Loop
            End With
            
            zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
            strShareInvoice = zlDatabase.GetPara("共用医疗卡批次", glngSys, mlngModul, , , True, intType)
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
                    '105985:李南春,2017/4/10,以医疗卡名称区分票据
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
    End If
    If mlngModul = 1102 Then Exit Sub
    '共用预交票据批次
    '恢复列宽度
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
    
    strShareInvoice = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModul, , , True, intType)
    '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";参数设置;") = 0 Then .Editable = flexEDNone
    End With
    
    '格式:领用ID1,预交类别ID1|领用IDn,预交类别IDn|...
    varData = Split(strShareInvoice, "|")
    '1.设置共享票据
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
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
                .TextMatrix(lngRow, .ColIndex("预交类型")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 0
            Case 1  '门诊票据
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 1
            Case Else   '住院票据
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
                .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = 2
            End Select
            
            .TextMatrix(lngRow, .ColIndex("领用人")) = Nvl(rsTemp!领用人)
            .TextMatrix(lngRow, .ColIndex("领用日期")) = Format(rsTemp!登记时间, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("号码范围")) = rsTemp!开始号码 & "," & rsTemp!终止号码
            .TextMatrix(lngRow, .ColIndex("剩余")) = Format(Val(Nvl(rsTemp!剩余数量)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("预交类型"))) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    '78751:李南春,2015/08/24,增加预交票据打印格式
    If mlngModul = 1103 Or mlngModul = 1101 Then
        '票据格式处理
        Dim strReport As String
        
        zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "预交发票打印方式", False, False
        strReport = "ZL" & glngSys \ 100 & "_BILL_1103"
        Set rsTemp = zlReadBillFormat(strReport)
        With vsBillFormat
            .Clear 1
            .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
            .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
        End With
        
        '读取参数值
        strBillFormat = zlDatabase.GetPara("预交发票格式", glngSys, mlngModul, , , True, intType)
        strPrintMode = zlDatabase.GetPara("预交发票打印方式", glngSys, mlngModul, , , True, intType1)
        '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        With vsBillFormat
            .TextMatrix(1, 0) = "门诊预交"
            .Cell(flexcpData, 1, 0) = 1
            .TextMatrix(2, 0) = "住院预交"
            .Cell(flexcpData, 2, 0) = 2
            .ColData(.ColIndex("票据格式")) = "0"
            .ColData(.ColIndex("预交打印方式")) = "0"
            .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
            Select Case intType
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("票据格式")) = IIf(intType = 5, 0, 1)
            End Select
            Select Case intType1
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("预交打印方式")) = IIf(intType1 = 5, 0, 1)
            End Select
            
            If (Val(.ColData(.ColIndex("票据格式"))) = 1 Or _
                Val(.ColData(.ColIndex("预交打印方式"))) = 1) Then
                .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
            Else
                .Editable = flexEDKbdMouse
            End If
        End With
        
        vsBillFormat.Tag = ""
        varData = Split(strBillFormat, "|")
        VarType = Split(strPrintMode, "|")
        
        With vsBillFormat
            .Clear 1
            .Rows = 3
            For lngRow = 1 To .Cols - 1
                .TextMatrix(lngRow, .ColIndex("预交打印方式")) = "0-不打印票据"
                .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & "," & ",", ",")
                    If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                    End If
                Next
                For i = 0 To UBound(VarType)
                    varTemp1 = Split(VarType(i) & "," & ",", ",")
                    If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("预交打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                        Exit For
                    End If
                Next
            Next
            If Val(.ColData(.ColIndex("预交打印方式"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("预交打印方式"), .Rows - 1, .ColIndex("预交打印方式")) = vbBlue
            End If
            
            If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
            End If
        End With
        '红票
        zl_vsGrid_Para_Restore mlngModul, vsRedBillFormat, Me.Name, "预交退款打印方式", False, False
        strReport = "ZL" & glngSys \ 100 & "_BILL_1103_1"
        Set rsTemp = zlReadBillFormat(strReport)
        With vsRedBillFormat
            .Clear 1
            .ColComboList(.ColIndex("票据格式")) = .BuildComboList(rsTemp, "序号,*说明", "序号")
            .ColComboList(.ColIndex("预交打印方式")) = "0-不打印票据|1-自动打印票据|2-选择是否打印票据"
        End With
        
        '读取参数值
        strBillFormat = zlDatabase.GetPara("退款发票格式", glngSys, mlngModul, , , True, intType)
        strPrintMode = zlDatabase.GetPara("预交退款打印方式", glngSys, mlngModul, , , True, intType1)
        '1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        With vsRedBillFormat
            .TextMatrix(1, 0) = "门诊预交"
            .Cell(flexcpData, 1, 0) = 1
            .TextMatrix(2, 0) = "住院预交"
            .Cell(flexcpData, 2, 0) = 2
            .ColData(.ColIndex("票据格式")) = "0"
            .ColData(.ColIndex("预交打印方式")) = "0"
            .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
            Select Case intType
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("票据格式")) = IIf(intType = 5, 0, 1)
            End Select
            Select Case intType1
            Case 1, 3, 5, 15
                 .ColData(.ColIndex("预交打印方式")) = IIf(intType1 = 5, 0, 1)
            End Select
            
            If (Val(.ColData(.ColIndex("票据格式"))) = 1 Or _
                Val(.ColData(.ColIndex("预交打印方式"))) = 1) Then
                .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
            Else
                .Editable = flexEDKbdMouse
            End If
        End With
        
        vsRedBillFormat.Tag = ""
        varData = Split(strBillFormat, "|")
        VarType = Split(strPrintMode, "|")
        
        With vsRedBillFormat
            .Clear 1
            .Rows = 3
            For lngRow = 1 To .Cols - 1
                .TextMatrix(lngRow, .ColIndex("预交打印方式")) = "0-不打印票据"
                .TextMatrix(lngRow, .ColIndex("票据格式")) = "0"
                For i = 0 To UBound(varData)
                    varTemp = Split(varData(i) & "," & ",", ",")
                    If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("票据格式")) = Val(varTemp(1)): Exit For
                    End If
                Next
                For i = 0 To UBound(VarType)
                    varTemp1 = Split(VarType(i) & "," & ",", ",")
                    If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                        .TextMatrix(lngRow, .ColIndex("预交打印方式")) = Decode(Val(varTemp1(1)), 0, "0-不打印票据", 1, "1-自动打印票据", "2-选择是否打印票据")
                        Exit For
                    End If
                Next
            Next
            If Val(.ColData(.ColIndex("预交打印方式"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("预交打印方式"), .Rows - 1, .ColIndex("预交打印方式")) = vbBlue
            End If
            
            If Val(.ColData(.ColIndex("票据格式"))) = 1 Then
                .Cell(flexcpForeColor, 0, .ColIndex("票据格式"), .Rows - 1, .ColIndex("票据格式")) = vbBlue
            End If
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, strTmp As String
    
    '本地共用就诊卡
    If IsValied = False Then Exit Sub
    Call SaveInvoice
    
    zlDatabase.SetPara "卡费记帐", chkCardFeeCharge.Value, glngSys, glngModul, IIf(chkCardFeeCharge.Enabled = True, True, False)
    Select Case mlngModul
    Case 1101 '病人信息
        zlDatabase.SetPara "建档同时必须发卡", chkMustCard.Value, glngSys, mlngModul, IIf(chkMustCard.Enabled = True, True, False)
        '问题27390  光标经过项目
        For i = 0 To chkItem.UBound
            zlDatabase.SetPara chkItem(i).Caption, chkItem(i).Value, glngSys, mlngModul, IIf(chkItem(i).Enabled = True, True, False)
        Next
        '76824，李南春，2014/8/19，医疗卡类别处理
        If cboType.ListIndex >= 0 Then
            zlDatabase.SetPara "缺省医疗卡类别", cboType.ItemData(cboType.ListIndex), glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        Else
            zlDatabase.SetPara "缺省医疗卡类别", 0, glngSys, mlngModul, IIf(cboType.Enabled = True, True, False)
        End If
        '54701:刘鹏飞,2012-09-19
        zlDatabase.SetPara "自动刷新数据", chkAutoRefresh.Value, glngSys, mlngModul, IIf(chkAutoRefresh.Enabled = True, True, False)
    Case 1102   '就诊卡
        '问题28130、27929
        If chkBruhCardBackCard.Value And chkBrushCardVerfy.Value Then
            strTmp = "3"
        ElseIf chkBruhCardBackCard.Value Then
            strTmp = "1"
        ElseIf chkBrushCardVerfy.Value Then
            strTmp = "2"
        Else
            strTmp = "0"
        End If
        Call zlDatabase.SetPara("退卡刷卡", strTmp, glngSys, mlngModul, IIf(chkBruhCardBackCard.Enabled = True, True, False))
    Case 1103
        zlDatabase.SetPara "允许更改缴款科室", chkAllowDept.Value, glngSys, glngModul, IIf(chkAllowDept.Enabled = True, True, False)
        zlDatabase.SetPara "仅显有余款的缴款单", chkHave.Value, glngSys, glngModul, IIf(chkHave.Enabled = True, True, False)
        zlDatabase.SetPara "退款禁止方式", IIf(optCheck(1).Value, 1, 0), glngSys, glngModul, IIf(optCheck(0).Enabled = True, True, False)
        '问题号:51628 修改人:刘兴洪,修改时间:2012-12-11 11:56:43
        zlDatabase.SetPara "病人未入科不准收预交", IIf(chkNotInDeptNotJk.Value, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        
        zlDatabase.SetPara "允许出院病人缴预交", IIf(chkAdvance, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "姓名模糊查找", chkSeekName.Value, glngSys, glngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        '63113:刘尔旋,2013-10-29,增加参数,住院退预交不需验证
        zlDatabase.SetPara "住院退预交验证", IIf(chkVeryfyInDeposit, "1", "0"), glngSys, glngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
        zlDatabase.SetPara "票据剩余X张时开始提醒收费员", IIf(chkCheckBillNum.Value = 1, "1", "0") & "|" & Val(txt票据张数.Text), glngSys, mlngModul, IIf(chkCheckBillNum.Enabled = True, True, False)
    End Select
    If mlngModul = 1101 Then '病人信息
    ElseIf mlngModul = 1102 Then '
    ElseIf mlngModul = 1103 Then '预交款
        
        With vs代收
            strTmp = ""
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("固定金额"))) <> 0 And Trim(.TextMatrix(i, .ColIndex("代收款项"))) <> "" Then
                    strTmp = strTmp & "|" & Trim(.TextMatrix(i, .ColIndex("代收款项"))) & ":" & Val(.TextMatrix(i, .ColIndex("固定金额")))
                End If
            Next
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        End With
        
        zlDatabase.SetPara "缺省预交结算方式", Trim(cboDefaultBalance.Text), glngSys, glngModul, IIf(cboDefaultBalance.Enabled = True, True, False)
        zlDatabase.SetPara "代收款设置", strTmp, glngSys, glngModul, IIf(vs代收.Editable = flexEDKbdMouse, True, False)
        zlDatabase.SetPara "缴预交后不清除信息", IIf(chkNotClearPatiInfor.Value = 1, 1, 0), glngSys, glngModul, IIf(chkNotClearPatiInfor.Enabled, True, False)
    End If
    'LED设备
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)
    '问题号:53408
    '扫描身份证签约
    zlDatabase.SetPara "扫描身份证签约", IIf(chkScanIDPatiVisa.Value = 1, 1, 0), glngSys, glngModul, InStr(1, mstrPrivs, ";参数设置;")
    
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

 Private Sub Load代收款()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载代收款
    '编制:刘兴洪
    '日期:2011-07-19 15:13:59
    '问题:  34705
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, varData As Variant, varTemp As Variant, j As Long, strTmp As String
    
     str结算方式 = zlDatabase.GetPara("缺省预交结算方式", glngSys, glngModul, , Array(cboDefaultBalance), InStr(mstrPrivs, ";参数设置;") > 0)
     
     On Error GoTo errHandle
    '结算方式
    strSQL = _
    " Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
    " From 结算方式应用 A,结算方式 B" & _
    " Where A.应用场合='预交款' And B.名称=A.结算方式 And Nvl(B.性质,1) In(1,2,3,5,8)" & _
    " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cboDefaultBalance
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp!名称)
            If .ListIndex < 0 And Val(Nvl(rsTmp!缺省)) = 1 Then .ListIndex = .NewIndex
            If str结算方式 = Nvl(rsTmp!名称) Then .ListIndex = .NewIndex
            rsTmp.MoveNext
        Loop
    End With
    '结算方式:金额|结算方式:金额....
    strTmp = zlDatabase.GetPara("代收款设置", glngSys, glngModul, , Array(vs代收), InStr(mstrPrivs, ";参数设置;") > 0)
    varData = Split(strTmp, "|")
    vs代收.Tag = "1"
    If vs代收.Enabled = False Then vs代收.Tag = "0"
    If vs代收.Tag = "1" Then vs代收.Editable = flexEDKbdMouse
    vs代收.Enabled = True
    rsTmp.Filter = "性质=5" '加入代收款
    With vs代收
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        i = 1
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        Do While rsTmp.EOF = False
            .TextMatrix(i, .ColIndex("代收款项")) = Nvl(rsTmp!名称)
            For j = 0 To UBound(varData)
                varTemp = Split(varData(j) & ":", ":")
                If Nvl(rsTmp!名称) = varTemp(0) Then
                    .TextMatrix(i, .ColIndex("固定金额")) = Format(Val(varTemp(1)), "###0.00;-###0.00;;")
                    Exit For
                End If
            Next
            i = i + 1
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
 End Sub

Private Sub Form_Load()
    Dim i As Long, lngCardTypeID As Long
    Dim strPrintMode As String '问题号:50656
    Dim strArr打印方式() As String '问题号:50656
    Dim strTmp As String
    gblnOK = False
    Me.sTab.TabVisible(2) = False   '34705
    sTab.TabVisible(0) = mlngModul = 1101
    sTab.TabVisible(2) = mlngModul = 1103    '34705
    sTab.TabVisible(1) = mlngModul <> 1103    '34705
    If mlngModul = 1103 Then Call Load代收款
    Call InitShareInvoice   '加载共用批票据信息
    
    'LED设备
    chkLedWelcome.Value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, ";参数设置;") > 0)
    chkCardFeeCharge.Value = IIf(zlDatabase.GetPara("卡费记帐", glngSys, glngModul, , Array(chkCardFeeCharge), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    '问题号:53408
    chkScanIDPatiVisa.Value = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul, , Array(chkScanIDPatiVisa), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    
    Select Case mlngModul
    Case 1101 ''病人信息
        chkMustCard.Value = IIf(zlDatabase.GetPara("建档同时必须发卡", glngSys, glngModul, , Array(chkMustCard), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        '问题27390 光标经过项目
        For i = 0 To chkItem.UBound
            chkItem(i).Value = zlDatabase.GetPara(chkItem(i).Caption, glngSys, mlngModul, 1, Array(chkItem(i)), InStr(mstrPrivs, ";参数设置;") > 0)
        Next
        lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, , Array(cboType), InStr(mstrPrivs, ";参数设置;") > 0))
        For i = 0 To cboType.ListCount - 1
            If cboType.ItemData(i) = lngCardTypeID Then cboType.ListIndex = i: Exit For
        Next
        
        '54701:刘鹏飞,2012-09-19
        chkAutoRefresh.Value = zlDatabase.GetPara("自动刷新数据", glngSys, mlngModul, 1, Array(chkAutoRefresh), InStr(mstrPrivs, ";参数设置;") > 0)
    
    Case 1102   '就诊卡
        '问题28130
        Select Case Val(zlDatabase.GetPara("退卡刷卡", glngSys, mlngModul, "0", Array(chkBruhCardBackCard, chkBrushCardVerfy), InStr(mstrPrivs, ";参数设置;") > 0))
        Case 0: chkBruhCardBackCard.Value = 0: chkBrushCardVerfy.Value = 0
        Case 1: chkBruhCardBackCard.Value = 1
        Case 2: chkBrushCardVerfy.Value = 1
        Case 3: chkBruhCardBackCard.Value = 1: chkBrushCardVerfy.Value = 1
        End Select
        chkBruhCardBackCard.Visible = True: chkBrushCardVerfy.Visible = True
    Case 1103  '预交款
        chkHave.Value = IIf(zlDatabase.GetPara("仅显有余款的缴款单", glngSys, glngModul, , Array(chkHave), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        chkAllowDept.Value = IIf(zlDatabase.GetPara("允许更改缴款科室", glngSys, glngModul, , Array(chkAllowDept), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        chkAdvance.Value = IIf(zlDatabase.GetPara("允许出院病人缴预交", glngSys, glngModul, , Array(chkAdvance), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        chkSeekName.Value = IIf(zlDatabase.GetPara("姓名模糊查找", glngSys, glngModul, , Array(chkSeekName), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        '63113:刘尔旋,2013-10-29,增加参数,住院退预交验证
        If gbln消费验证 = False Then chkVeryfyInDeposit.Visible = False
        chkVeryfyInDeposit.Value = Val(zlDatabase.GetPara("住院退预交验证", glngSys, glngModul, "0", Array(chkVeryfyInDeposit), InStr(mstrPrivs, ";参数设置;") > 0))

        
        If zlDatabase.GetPara("退款禁止方式", glngSys, glngModul, , Array(optCheck(0), optCheck(1), fra退款设置), InStr(mstrPrivs, ";参数设置;") > 0) = "1" Then
            optCheck(1).Value = True
        Else
            optCheck(0).Value = True
        End If
        '问题:43061
        chkNotClearPatiInfor.Value = IIf(zlDatabase.GetPara("缴预交后不清除信息", glngSys, glngModul, , Array(chkNotClearPatiInfor), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        '问题号:51628 修改人:刘兴洪,修改时间:2012-12-11 11:56:43
        chkNotInDeptNotJk.Value = IIf(zlDatabase.GetPara("病人未入科不准收预交", glngSys, mlngModul, , Array(chkNotInDeptNotJk), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
        '问题:50656
        '78410,冉俊明,2014-10-8,按类型及权限设置控件字体颜色和可编辑状态
        '37372
        strTmp = zlDatabase.GetPara("票据剩余X张时开始提醒收费员", glngSys, mlngModul, "0|10", Array(txt票据张数, upd票据张数, chkCheckBillNum), InStr(mstrPrivs, ";参数设置;") > 0)
        upd票据张数.Value = Val(Split(strTmp & "|", "|")(1))
        txt票据张数.Text = upd票据张数.Value
        chkCheckBillNum.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
        txt票据张数.Enabled = chkCheckBillNum.Enabled And chkCheckBillNum.Value = 1
        upd票据张数.Enabled = txt票据张数.Enabled
    End Select
    
    '61808:刘鹏飞,2013-05-21
    '78751:李南春,2015/08/24,增加预交票据打印格式
    If Not (mlngModul = 1103 Or mlngModul = 1101) Then
        fra票据格式.Visible = False
        fra票据格式.Enabled = False
        fraPrepay.Height = fra票据格式.Top + fra票据格式.Height - fraPrepay.Top
        vsPrepay.Height = fraPrepay.Height - vsPrepay.Top - 90
    End If
    chkCheckBillNum.Visible = mlngModul = 1103
    txt票据张数.Visible = mlngModul = 1103
    upd票据张数.Visible = mlngModul = 1103
    
    '问题28130、27929
    chkHave.Visible = mlngModul = 1103
    'chkAllowOut.Visible = mlngModul = 1103
    chkAllowDept.Visible = mlngModul = 1103
    chkLedWelcome.Visible = mlngModul = 1103
    chkMustCard.Visible = mlngModul = 1101
    chkCardFeeCharge.Visible = mlngModul = 1102
    Exit Sub
errH:
    If ErrCenter() = 1 Then
         Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln担保 = False
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用医疗票据列表", False, False
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

 
'问题27390
Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 0 And chkItem(1).Enabled And chkItem(1).Visible Then chkItem(1).SetFocus
End Sub

Private Sub vs代收_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        With vs代收
            Select Case Col
            Case .ColIndex("固定金额")
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, .Col)), "###0.00;-###0.00;;")
            Case .ColIndex("选择")
            Case Else
            End Select
        End With
End Sub

Private Sub vs代收_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs代收
        Select Case Col
        Case .ColIndex("固定金额")
            Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vs代收_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs代收
        If .Col >= .ColIndex("固定金额") And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
End Sub

Private Sub vs代收_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '编辑处理
    Dim intCol As Integer, strKey As String, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vs代收
        Select Case Col
        Case .ColIndex("固定金额")
                If Row < .Rows - 1 Then
                    .Col = Col: .Row = .Row + 1
                End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs代收_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs代收_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vs代收
        Select Case .Col
            Case .ColIndex("固定金额")
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    If KeyAscii = vbKeyBack Then Exit Sub
                    If KeyAscii = vbKeyReturn Then Exit Sub
                    If KeyAscii = Asc(".") Then
                        If InStr(1, .EditText, ".") = 0 Then
                            Exit Sub
                        End If
                    End If
                    KeyAscii = 0
                End If
            Case Else
        End Select
    End With
End Sub

Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "共用预交票据列表", False, False
End Sub
 
Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "共用预交票据列表", False, False
End Sub
Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("医疗卡类别"))) = Trim(.Cell(flexcpData, i, .ColIndex("医疗卡类别"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
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
Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("选择")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("预交类型"))) = Trim(.Cell(flexcpData, i, .ColIndex("预交类型"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
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
Private Sub chkCheckBillNum_Click()
    txt票据张数.Enabled = chkCheckBillNum.Enabled And chkCheckBillNum.Value = 1
    upd票据张数.Enabled = txt票据张数.Enabled
End Sub
