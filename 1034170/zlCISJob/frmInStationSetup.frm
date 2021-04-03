VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "frmInStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7740
      TabIndex        =   47
      Top             =   8445
      Width           =   7740
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   49
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5280
         TabIndex        =   48
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   8040
         Y1              =   15
         Y2              =   0
      End
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   14949
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frmInStationSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAdvice"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEPR"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra首页整理"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraMedRec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDeptView"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra药房"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "首页附加项目"
      TabPicture(1)   =   "frmInStationSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGroup"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra药房 
         Caption         =   "住院医嘱编辑"
         Height          =   615
         Left            =   120
         TabIndex        =   90
         Top             =   7320
         Width           =   7575
         Begin VB.CheckBox chk缺省药房 
            Caption         =   "住院医嘱下达强制缺省药房"
            Height          =   240
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   2580
         End
      End
      Begin VB.CheckBox chkDeptView 
         Caption         =   "拥有全院病人权限操作者不显示无床位的病区科室"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   8040
         Width           =   4335
      End
      Begin VB.Frame fraMedRec 
         Caption         =   "病案审查反馈设置"
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   6690
         Width           =   7560
         Begin VB.TextBox txtMedRec 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   1040
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   1025
            TabIndex        =   33
            Top             =   420
            Width           =   300
         End
         Begin VB.Label lblMedRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "显示    天内的病案审查反馈数"
            Height          =   180
            Left            =   645
            TabIndex        =   35
            Top             =   255
            Width           =   2520
         End
      End
      Begin VB.Frame fra首页整理 
         Caption         =   " 首页整理设置 "
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7560
         Begin VB.CheckBox chkZLZD 
            Caption         =   "出院肿瘤诊断填写分化程度和最高诊断依据"
            Height          =   255
            Left            =   2280
            TabIndex        =   89
            Top             =   4200
            Width           =   4095
         End
         Begin VB.CheckBox Chk病理 
            Caption         =   "按ICD-10录入时，病理诊断只允许录入M打头的肿瘤形态学编码"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1560
            Width           =   5295
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   72
            Top             =   1875
            Width           =   3615
            Begin VB.OptionButton optICD附码 
               Caption         =   "必须填写"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   75
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optICD附码 
               Caption         =   "提示是否填写"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   74
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton optICD附码 
               Caption         =   "不检查"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   73
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.CheckBox chkSeparEdit 
            Caption         =   "医生和护士分别填写病案首页"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   3420
            Width           =   4095
         End
         Begin VB.CheckBox chk中医 
            Caption         =   $"frmInStationSetup.frx":0044
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2475
            Width           =   4095
         End
         Begin VB.CheckBox chk使用手术结束时间 
            Caption         =   "使用手术结束时间"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   2535
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   25
            Top             =   2200
            Width           =   3735
            Begin VB.OptionButton opt区域 
               Caption         =   "必须填写"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   28
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton opt区域 
               Caption         =   "提示是否填写"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   27
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt区域 
               Caption         =   "不检查"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   26
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   11
            Top             =   970
            Width           =   3495
            Begin VB.OptionButton opt损伤中毒 
               Caption         =   "必须填写"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   14
               ToolTipText     =   "如果出院诊断不是S、T类，则禁止填写损伤中毒。"
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton opt损伤中毒 
               Caption         =   "提示是否填写"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   13
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt损伤中毒 
               Caption         =   "不检查"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   12
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   7
            Top             =   1300
            Width           =   3495
            Begin VB.OptionButton opt病理诊断 
               Caption         =   "不检查"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton opt病理诊断 
               Caption         =   "提示是否填写"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   9
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt病理诊断 
               Caption         =   "必须填写"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   8
               ToolTipText     =   "如果出院诊断不是C00到D48时，则禁止填写病理诊断。"
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkGet附码 
            Caption         =   "诊断自动提取附码"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label lblICD附码 
            Caption         =   "主要出院诊断编码为C00到D48时,ICD附码："
            Height          =   210
            Left            =   240
            TabIndex        =   76
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Label lblSeparEdit 
            Caption         =   "不良事件的项目、输液反应、引发药物、临床表现、住院期间身体约束、离院时透析(血透、腹透)尿素氮值等信息在启用该参数时只能由护士填写"
            Height          =   360
            Left            =   480
            TabIndex        =   51
            Top             =   3720
            Width           =   6615
         End
         Begin VB.Label lbl中医 
            Caption         =   $"frmInStationSetup.frx":0070
            Height          =   585
            Left            =   480
            TabIndex        =   38
            Top             =   2790
            Width           =   6735
         End
         Begin VB.Label lbl首页标准 
            Caption         =   "病案首页标准"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lbl区域 
            Caption         =   "保存时区域项："
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2200
            Width           =   1335
         End
         Begin VB.Label lbl损伤中毒 
            Caption         =   "主要出院诊断编码为S、T类时,损伤中毒诊断："
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   970
            Width           =   3735
         End
         Begin VB.Label lbl病理诊断 
            Caption         =   "主要出院诊断编码为C00到D48时,病理诊断："
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   3375
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6800
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   7395
         Begin VB.CommandButton cmdAdd 
            Caption         =   "增加(&A)"
            Height          =   350
            Left            =   3120
            TabIndex        =   4
            Top             =   30
            Width           =   1100
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "修改(&M)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   4320
            TabIndex        =   3
            Top             =   30
            Width           =   1100
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "删除(&D)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   5520
            TabIndex        =   2
            Top             =   30
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   6300
            Left            =   315
            TabIndex        =   5
            Top             =   480
            Width           =   6840
            _cx             =   12065
            _cy             =   11112
            Appearance      =   3
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
            BackColorSel    =   16574424
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
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
      End
      Begin VB.Frame fraEPR 
         Caption         =   "提醒设置"
         Height          =   1545
         Left            =   120
         TabIndex        =   39
         Top             =   5040
         Width           =   7560
         Begin VB.CheckBox chkWarn 
            Caption         =   "输血反应"
            Height          =   195
            Index           =   26
            Left            =   6480
            TabIndex        =   88
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "用血审核"
            Height          =   195
            Index           =   25
            Left            =   5400
            TabIndex        =   83
            Top             =   1185
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "校对疑问"
            Height          =   195
            Index           =   24
            Left            =   4320
            TabIndex        =   78
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "备血完成"
            Height          =   195
            Index           =   23
            Left            =   3255
            TabIndex        =   55
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "病历质控"
            Height          =   195
            Index           =   22
            Left            =   2235
            TabIndex        =   71
            Top             =   1185
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "处方审查"
            Height          =   195
            Index           =   20
            Left            =   6480
            TabIndex        =   70
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkSoundYS 
            Caption         =   "启用语音提示"
            Height          =   195
            Left            =   4335
            TabIndex        =   69
            Top             =   345
            Width           =   1470
         End
         Begin VB.CommandButton cmdSoundYSSet 
            Caption         =   "语音设置(&S)"
            Height          =   350
            Left            =   5775
            TabIndex        =   68
            Top             =   270
            Width           =   1410
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "传染病"
            Height          =   195
            Index           =   21
            Left            =   1200
            TabIndex        =   65
            Top             =   1185
            Width           =   885
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "医嘱审核"
            Height          =   195
            Index           =   19
            Left            =   5400
            TabIndex        =   63
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "报告撤消"
            Height          =   195
            Index           =   18
            Left            =   4320
            TabIndex        =   56
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "危急值"
            Height          =   195
            Index           =   17
            Left            =   3255
            TabIndex        =   54
            Top             =   885
            Width           =   885
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "医嘱安排"
            Height          =   195
            Index           =   16
            Left            =   2235
            TabIndex        =   53
            Top             =   885
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "病历审阅"
            Height          =   195
            Index           =   15
            Left            =   1200
            TabIndex        =   52
            Top             =   885
            Width           =   1065
         End
         Begin VB.TextBox txtNotifyEPRDay 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   600
            MaxLength       =   2
            TabIndex        =   43
            Text            =   "1"
            Top             =   600
            Width           =   300
         End
         Begin VB.Frame fraNotifyEPRDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   585
            TabIndex        =   42
            Top             =   780
            Width           =   300
         End
         Begin VB.Frame fraNotifyEPR 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   585
            TabIndex        =   41
            Top             =   510
            Width           =   300
         End
         Begin VB.TextBox txtNotifyEPR 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   600
            MaxLength       =   3
            TabIndex        =   40
            Text            =   "10"
            Top             =   330
            Width           =   300
         End
         Begin VB.CheckBox chkNotifyEPR 
            Caption         =   "每    分钟自动刷新提醒区域中的内容"
            Height          =   195
            Left            =   105
            TabIndex        =   44
            Top             =   345
            Width           =   3900
         End
         Begin VB.Label lblNotifyEPRDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "将    天内完成的内容显示在提醒区域"
            Height          =   180
            Left            =   375
            TabIndex        =   46
            Top             =   615
            Width           =   3060
         End
         Begin VB.Label lblArea 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提醒内容:"
            Height          =   180
            Left            =   360
            TabIndex        =   45
            Top             =   885
            Width           =   810
         End
      End
      Begin VB.Frame fraAdvice 
         Caption         =   "提醒设置 "
         Height          =   1590
         Left            =   120
         TabIndex        =   17
         Top             =   5040
         Width           =   7560
         Begin VB.CheckBox chkWarn 
            Caption         =   "血袋回收"
            Height          =   195
            Index           =   12
            Left            =   6240
            TabIndex        =   87
            Top             =   1245
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "备血完成"
            Height          =   195
            Index           =   11
            Left            =   5160
            TabIndex        =   86
            Top             =   1245
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "RIS预约准备"
            Height          =   195
            Index           =   8
            Left            =   2640
            TabIndex        =   79
            Top             =   1245
            Width           =   1335
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "取血通知"
            Height          =   195
            Index           =   9
            Left            =   3960
            TabIndex        =   82
            Top             =   1245
            Width           =   1095
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "销帐申请"
            Height          =   195
            Index           =   6
            Left            =   600
            TabIndex        =   81
            Top             =   1230
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "RIS预约"
            Height          =   195
            Index           =   7
            Left            =   1680
            TabIndex        =   80
            Top             =   1245
            Width           =   1005
         End
         Begin VB.CommandButton cmdSoundHSSet 
            Caption         =   "语音设置(&S)"
            Height          =   350
            Left            =   6000
            TabIndex        =   67
            Top             =   240
            Width           =   1410
         End
         Begin VB.CheckBox chkSoundHS 
            Caption         =   "启用语音提示"
            Height          =   195
            Left            =   4440
            TabIndex        =   66
            Top             =   360
            Width           =   1470
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "输液拒绝"
            Height          =   195
            Index           =   5
            Left            =   5520
            TabIndex        =   62
            Top             =   915
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "危急值"
            Height          =   195
            Index           =   4
            Left            =   4560
            TabIndex        =   61
            Top             =   915
            Width           =   870
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "安排"
            Height          =   195
            Index           =   3
            Left            =   3840
            TabIndex        =   60
            Top             =   915
            Width           =   675
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "新废"
            Height          =   195
            Index           =   2
            Left            =   3060
            TabIndex        =   59
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "新停"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   58
            Top             =   915
            Width           =   675
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "新开"
            Height          =   195
            Index           =   0
            Left            =   1500
            TabIndex        =   57
            Top             =   915
            Width           =   675
         End
         Begin VB.TextBox txtNotifyAdviceDay 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   795
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "1"
            Top             =   585
            Width           =   300
         End
         Begin VB.Frame fraNotifyAdviceDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   780
            TabIndex        =   20
            Top             =   765
            Width           =   300
         End
         Begin VB.Frame fraNotifyAdvice 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   780
            TabIndex        =   19
            Top             =   495
            Width           =   300
         End
         Begin VB.TextBox txtNotifyAdvice 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   795
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "10"
            Top             =   315
            Width           =   300
         End
         Begin VB.CheckBox chkNotifyAdvice 
            Caption         =   "每    分钟自动刷新医嘱提醒区域中的内容"
            Height          =   195
            Left            =   300
            TabIndex        =   22
            Top             =   330
            Width           =   3900
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "标本拒收（暂未启用）"
            Enabled         =   0   'False
            Height          =   195
            Index           =   10
            Left            =   4560
            TabIndex        =   85
            Top             =   600
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label lblNotifyAdviceDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "将    天内处理的医嘱病人显示在提醒区域"
            Height          =   180
            Left            =   570
            TabIndex        =   24
            Top             =   600
            Width           =   3420
         End
         Begin VB.Label lbl提醒内容 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提醒内容:"
            Height          =   180
            Left            =   555
            TabIndex        =   23
            Top             =   915
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmInStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbln护士站 As Boolean
Public mstrPrivs As String
Private mlngModual As Long
Private lng损伤中毒 As Long
Private lng病理诊断 As Long
Private lngICD附码 As Long
Private lng区域 As Long

Private Enum Enum_chkWarn
    '护士站提醒参数
    chkN新开 = 0
    chkN新停 = 1
    chkN新废 = 2
    chkN安排 = 3
    chkN危急值 = 4
    chkN输液拒绝 = 5
    chkN销帐申请 = 6
    chkNRIS预约 = 7
    chkNRIS预约准备 = 8
    chk取血通知 = 9
    chk标本拒收 = 10        '暂未启用，保留此处一是为了版本之间的兼容性，一是为了后续加入此功能时不造成混乱。
    chk备血完成 = 11
    chk血袋回收 = 12
    '医生站提醒参数
    chkD病历审阅 = 15
    chkD医嘱安排 = 16
    chkD危急值 = 17
    chkD报告撤消 = 18
    chkD医嘱审核 = 19
    chkD处方审查 = 20
    chkD传染病 = 21
    chkD病历质控 = 22
    chkD备血完成 = 23
    chkD校对疑问 = 24
    chkD用血审核 = 25
    chkD输血反应 = 26
End Enum

Public Sub ShowMe()
    '由新版住院护士工作站调用，显示标注按钮
    Me.Show vbModal
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex = 0 Or cboType.ListIndex = 3 Then
        chk使用手术结束时间.Visible = True
    Else
        chk使用手术结束时间.Visible = False
    End If
End Sub

Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    If vsfMain.Row > 0 Then
        If CheckVsf = False Then vsfMain.SetFocus: Exit Sub
        frmInMedSetup.ShowMe vsfMain.TextMatrix(vsfMain.Row, 0), vsfMain.TextMatrix(vsfMain.Row, 1), vsfMain.TextMatrix(vsfMain.Row, 2), "修改", Me
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    
    If vsfMain.Row > 0 Then
        If CheckVsf = False Then vsfMain.SetFocus: Exit Sub
        If MsgBox("确认要删除[" & vsfMain.TextMatrix(vsfMain.Row, 1) & "]吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "zl_病案项目_edit(null,null,null,'" & vsfMain.TextMatrix(vsfMain.Row, 0) & "',2)"
        On Error GoTo errHandle
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfMain.RemoveItem vsfMain.Row
        If vsfMain.Rows = 1 Then
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
    frmInMedSetup.ShowMe "", "", "", "新增", Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If mbln护士站 Then
        If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
            If txtNotifyAdvice.Text = "" Then
                MsgBox "请设置医嘱提醒的自动刷新间隔。", vbInformation, gstrSysName
            Else
                MsgBox "医嘱提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
            End If
            txtNotifyAdvice.SetFocus: Exit Sub
        End If
        If Val(txtNotifyAdviceDay.Text) = 0 Then
            If txtNotifyAdviceDay.Text = "" Then
                MsgBox "请设置要提醒的医嘱天数。", vbInformation, gstrSysName
            Else
                MsgBox "要提醒的医嘱天数至少应为1天。", vbInformation, gstrSysName
            End If
            txtNotifyAdviceDay.SetFocus: Exit Sub
        End If
    Else
        If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
            If txtNotifyEPR.Text = "" Then
                MsgBox "请设置病历审阅提醒的自动刷新间隔。", vbInformation, gstrSysName
            Else
                MsgBox "病历审阅提醒的自动刷新间隔至少应为1分钟。", vbInformation, gstrSysName
            End If
            txtNotifyEPR.SetFocus: Exit Sub
        End If
        
        If Val(txtNotifyEPRDay.Text) = 0 Then
            If txtNotifyEPRDay.Text = "" Then
                MsgBox "请设置要提醒审阅的病历完成天数。", vbInformation, gstrSysName
            Else
                MsgBox "要提醒审阅的病历完成天数至少应为1天。", vbInformation, gstrSysName
            End If
            txtNotifyEPRDay.SetFocus: Exit Sub
        End If
    End If
    
    If txtMedRec.Text = "" Then
        MsgBox "请设置病理审查反馈提醒的天数。", vbInformation, gstrSysName
        txtMedRec.SetFocus: Exit Sub
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    Call zlDatabase.SetPara("病案审查反馈天数", txtMedRec.Text, glngSys, mlngModual, blnSetup)
        
    '自动刷新医嘱提醒
    If mbln护士站 Then
        Call zlDatabase.SetPara("自动刷新医嘱间隔", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, p住院护士站, blnSetup)
        Call zlDatabase.SetPara("自动刷新医嘱天数", Val(txtNotifyAdviceDay.Text), glngSys, p住院护士站, blnSetup)
        strTmp = ""
        For i = chkN新开 To chk血袋回收
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("自动刷新医嘱类型", strTmp, glngSys, p住院护士站, blnSetup)
        Call zlDatabase.SetPara("启用语音提示", chkSoundHS.Value, glngSys, p住院护士站, blnSetup)
    Else
        Call zlDatabase.SetPara("自动刷新病历审阅间隔", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("自动刷新病历审阅天数", Val(txtNotifyEPRDay.Text), glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("损伤中毒检查", lng损伤中毒, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("病理诊断检查", lng病理诊断, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("ICD附码检查", lngICD附码, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("区域检查", lng区域, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("病理诊断只允许录入肿瘤形态学编码", Chk病理.Value, glngSys, p住院医生站, blnSetup)
        
        strTmp = ""
        For i = chkD病历审阅 To chkD输血反应
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("自动刷新内容", strTmp, glngSys, p住院医生站, blnSetup)
        
        Call zlDatabase.SetPara("病案首页标准", cboType.ListIndex, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("使用手术结束时间", chk使用手术结束时间, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("中医科室不使用西医病案首页项目", chk中医.Value, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("医生和护士分别填写病案首页", chkSeparEdit.Value, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("不显示无床位的病区科室", chkDeptView.Value, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("启用语音提示", chkSoundYS.Value, glngSys, p住院医生站, blnSetup)
        Call zlDatabase.SetPara("诊断自动提取附码", chkGet附码.Value, glngSys, 0, blnSetup)
        Call zlDatabase.SetPara("出院肿瘤诊断填写分化程度和最高诊断依据", chkZLZD.Value, glngSys, 0, blnSetup)
    End If
    
    Call zlDatabase.SetPara("住院医嘱下达强制缺省药房", chk缺省药房.Value, glngSys, p住院医嘱下达, blnSetup)
    
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundHSSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub cmdSoundYSSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    Dim strNotify As String
    Dim varTmp As Variant
    
    gblnOK = False
    lng病理诊断 = 0
    lng损伤中毒 = 0
    lngICD附码 = 0
    mlngModual = IIf(mbln护士站, p住院护士站, p住院医生站)
    If mbln护士站 Then
        fraAdvice.Visible = True
        fraEPR.Visible = False
        sstInfo.TabVisible(1) = False
        fra首页整理.Visible = False
        chkDeptView.Visible = False
        fraAdvice.Top = fra首页整理.Top
        fraEPR.Top = fra首页整理.Top
        fraMedRec.Top = fraAdvice.Top + fraAdvice.Height + 50
        fra药房.Top = fraMedRec.Top + fraMedRec.Height + 50
        i = fra首页整理.Height + chkDeptView.Height + 100
    Else
        fraAdvice.Visible = False
        fraEPR.Visible = True
        chkDeptView.Visible = True
        fraMedRec.Top = fraEPR.Top + fraEPR.Height + 50
        chkDeptView.Top = fra药房.Top + fra药房.Height + 50
        i = fraAdvice.Height - fraEPR.Height
    End If
    Me.Height = Me.Height - i
    sstInfo.Height = sstInfo.Height - i
    chkWarn(chk取血通知).Visible = gbln血库系统
    chkWarn(chk备血完成).Visible = gbln血库系统
    chkWarn(chk血袋回收).Visible = gbln血库系统
    
    
    '住院医嘱下达强制缺省药房
    chk缺省药房.Value = Val(zlDatabase.GetPara("住院医嘱下达强制缺省药房", glngSys, p住院医嘱下达, "1", Array(chk缺省药房), intType))
    
    '自动刷新医嘱提醒
    If mbln护士站 Then
        strPar = zlDatabase.GetPara("自动刷新医嘱间隔", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "参数设置") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
        End If
        '前面事件中会自动可用，因此后面强制设置
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
            txtNotifyAdvice.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("自动刷新医嘱天数", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "参数设置") > 0)
        txtNotifyAdviceDay.Text = Val(strPar)
        
        strPar = zlDatabase.GetPara("自动刷新医嘱类型", glngSys, mlngModual, "000000000000", Array(lbl提醒内容, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "参数设置") > 0)
        For i = 1 To Len(strPar)
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        Next
        txtMedRec.Text = zlDatabase.GetPara("病案审查反馈天数", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "参数设置") > 0)
        
        chkSoundHS.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, mlngModual, , Array(chkSoundHS, cmdSoundHSSet), InStr(mstrPrivs, "参数设置") > 0, intType))
        
    Else
        strPar = zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, mlngModual, , Array(chkNotifyEPR), InStr(mstrPrivs, "参数设置") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
        End If
        '前面事件中会自动可用，因此后面强制设置
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "参数设置") = 0 Then
            txtNotifyEPR.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, mlngModual, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), InStr(mstrPrivs, "参数设置") > 0)
        txtNotifyEPRDay.Text = Val(strPar)
        opt损伤中毒(Val(zlDatabase.GetPara("损伤中毒检查", glngSys, p住院医生站, 0, Array(opt损伤中毒(0), opt损伤中毒(1), opt损伤中毒(2), lbl损伤中毒), InStr(mstrPrivs, "参数设置") > 0) & "")).Value = True
        opt病理诊断(Val(zlDatabase.GetPara("病理诊断检查", glngSys, p住院医生站, 0, Array(opt病理诊断(0), opt病理诊断(1), opt病理诊断(2), lbl病理诊断), InStr(mstrPrivs, "参数设置") > 0) & "")).Value = True
        optICD附码(Val(zlDatabase.GetPara("ICD附码检查", glngSys, p住院医生站, 0, Array(optICD附码(0), optICD附码(1), optICD附码(2), lblICD附码), InStr(mstrPrivs, "参数设置") > 0) & "")).Value = True
        opt区域(Val(zlDatabase.GetPara("区域检查", glngSys, p住院医生站, 1, Array(opt区域(0), lbl区域, opt区域(1), opt区域(2)), InStr(mstrPrivs, "参数设置") > 0) & "")).Value = True
        With vsfMain
            vsfMain.Rows = 1
            vsfMain.Cols = 3
            .TextMatrix(0, 0) = "编码"
            .TextMatrix(0, 1) = "名称"
            .TextMatrix(0, 2) = "内容"
            .ColWidth(0) = 1400
            .ColWidth(1) = 2500
            .ColWidth(2) = 2500
            .Cell(flexcpAlignment, 0, 0, 0, 2) = 4
        End With
        
        varTmp = Array(chkWarn(chkD病历审阅), chkWarn(chkD医嘱安排), chkWarn(chkD危急值), chkWarn(chkD报告撤消), chkWarn(chkD医嘱审核), chkWarn(chkD处方审查), chkWarn(chkD传染病), chkWarn(chkD病历质控), chkWarn(chkD备血完成), chkWarn(chkD校对疑问), chkWarn(chkD用血审核), chkWarn(chkD输血反应), lblArea)
        strNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p住院医生站, , varTmp, InStr(mstrPrivs, "参数设置") > 0)
            
        chkWarn(chkD病历审阅).Value = Val(Mid(strNotify, 1, 1))
        chkWarn(chkD医嘱安排).Value = Val(Mid(strNotify, 2, 1))
        chkWarn(chkD危急值).Value = Val(Mid(strNotify, 3, 1))
        chkWarn(chkD报告撤消).Value = Val(Mid(strNotify, 4, 1))
        chkWarn(chkD医嘱审核).Value = Val(Mid(strNotify, 5, 1))
        chkWarn(chkD处方审查).Value = Val(Mid(strNotify, 6, 1))
        chkWarn(chkD传染病).Value = Val(Mid(strNotify, 7, 1))
        chkWarn(chkD病历质控).Value = Val(Mid(strNotify, 8, 1))
        chkWarn(chkD备血完成).Value = Val(Mid(strNotify, 9, 1))
        chkWarn(chkD备血完成).Visible = gbln血库系统
        chkWarn(chkD校对疑问).Value = Val(Mid(strNotify, 10, 1))
        chkWarn(chkD用血审核).Value = Val(Mid(strNotify, 11, 1))
        chkWarn(chkD用血审核).Visible = gbln血库系统
        chkWarn(chkD输血反应).Value = Val(Mid(strNotify, 12, 1))
        chkWarn(chkD输血反应).Visible = gbln血库系统
        
        Call Get病案项目
        cboType.Clear
        cboType.AddItem "0-卫生部标准"
        cboType.AddItem "1-四川省标准"
        cboType.AddItem "2-云南省标准"
        cboType.AddItem "3-湖南省标准"
        Call zlControl.CboSetIndex(cboType.hwnd, Val(zlDatabase.GetPara("病案首页标准", glngSys, p住院医生站, "0", Array(cboType, lbl首页标准), InStr(mstrPrivs, "参数设置") > 0)))
        Call cboType_Click
        If InStr(mstrPrivs, "参数设置") = 0 Then
        
            chkWarn(chkD病历审阅).Enabled = False
            chkWarn(chkD医嘱安排).Enabled = False
            chkWarn(chkD危急值).Enabled = False
            chkWarn(chkD报告撤消).Enabled = False
            chkWarn(chkD医嘱审核).Enabled = False
            chkWarn(chkD处方审查).Enabled = False
            chkWarn(chkD传染病).Enabled = False
            chkWarn(chkD病历质控).Enabled = False
            chkWarn(chkD备血完成).Enabled = False
            chkWarn(chkD校对疑问).Enabled = False
            chkWarn(chkD用血审核).Enabled = False
            chkWarn(chkD输血反应).Enabled = False
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
        Chk病理.Value = Val(zlDatabase.GetPara("病理诊断只允许录入肿瘤形态学编码", glngSys, mlngModual, 0, Array(Chk病理), InStr(mstrPrivs, "参数设置") > 0))
        txtMedRec.Text = zlDatabase.GetPara("病案审查反馈天数", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "参数设置") > 0)
        chk使用手术结束时间.Value = Val(zlDatabase.GetPara("使用手术结束时间", glngSys, mlngModual, 0, Array(chk使用手术结束时间), InStr(mstrPrivs, "参数设置") > 0))
        chk中医.Value = Val(zlDatabase.GetPara("中医科室不使用西医病案首页项目", glngSys, mlngModual, 0, Array(chk中医), InStr(mstrPrivs, "参数设置") > 0))
        chkSeparEdit.Value = Val(zlDatabase.GetPara("医生和护士分别填写病案首页", glngSys, mlngModual, 0, Array(chkSeparEdit), InStr(mstrPrivs, "参数设置") > 0))
        chkDeptView.Value = Val(zlDatabase.GetPara("不显示无床位的病区科室", glngSys, mlngModual, 0, Array(chkDeptView), InStr(mstrPrivs, "参数设置") > 0))
        chkSoundYS.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, mlngModual, , Array(chkSoundYS, cmdSoundYSSet), InStr(mstrPrivs, "参数设置") > 0))
        chkGet附码.Value = Val(zlDatabase.GetPara("诊断自动提取附码", glngSys, 0, 0, Array(chkGet附码), InStr(mstrPrivs, "参数设置") > 0))
        chkZLZD.Value = Val(zlDatabase.GetPara("出院肿瘤诊断填写分化程度和最高诊断依据", glngSys, 0, 0, Array(chkZLZD), InStr(mstrPrivs, "参数设置") > 0))
    End If
End Sub

Private Sub Get病案项目()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    
    strSQL = "select 编码,名称,内容 from 病案项目 order by 编码"
On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then cmdModify.Enabled = False: cmdDelete.Enabled = False: Exit Sub
    lngRow = 1
    While Not rsTemp.EOF
        With vsfMain
            .Rows = lngRow + 1
            .TextMatrix(lngRow, 0) = rsTemp!编码 & ""
            .TextMatrix(lngRow, 1) = rsTemp!名称 & ""
            .TextMatrix(lngRow, 2) = rsTemp!内容 & ""
        End With
        lngRow = lngRow + 1
        rsTemp.MoveNext
    Wend
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln护士站 = False
End Sub

Private Sub optICD附码_Click(Index As Integer)
    lngICD附码 = Index
End Sub

Private Sub opt病理诊断_Click(Index As Integer)
    lng病理诊断 = Index
End Sub

Private Sub opt区域_Click(Index As Integer)
    lng区域 = Index
End Sub

Private Sub opt损伤中毒_Click(Index As Integer)
    lng损伤中毒 = Index
End Sub

Private Sub sstInfo_Click(PreviousTab As Integer)
    If sstInfo.Tab = 1 Then
        vsfMain.SetFocus
        If vsfMain.Rows > 1 Then vsfMain.Row = 1
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub txtMedRec_GotFocus()
    Call zlControl.TxtSelAll(txtMedRec)
End Sub

Private Sub txtMedRec_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsfMain_DblClick()
    Call cmdModify_Click
End Sub

Private Sub vsfMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMain.Row > 0 Then
            Call cmdDelete_Click
        End If
    End If
End Sub

Private Function CheckVsf() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(信息名) as 数量 from 病案主页从表 where 信息名='" & vsfMain.TextMatrix(vsfMain.Row, 1) & "'"
    
    err = 0: On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp!数量 > 0 Then
        MsgBox "该项目已经使用,不能进行修改或删除!"
        CheckVsf = False
        Exit Function
    End If
    CheckVsf = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
