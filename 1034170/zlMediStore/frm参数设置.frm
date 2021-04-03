VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm参数设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frm参数设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabMain 
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frm参数设置.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl天数"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl查询天数"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra同步处理库存"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra库房选择"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra成本价"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkALLPlanPoint"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra移库流程控制"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra盘点时间范围"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra上次采购信息"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra采购计划"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt查询天数"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "资质校验(&1)"
      TabPicture(1)   =   "frm参数设置.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCheck"
      Tab(1).Control(1)=   "vsfCheck"
      Tab(1).Control(2)=   "lblComment"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txt查询天数 
         Height          =   300
         Left            =   3240
         TabIndex        =   66
         Text            =   "7"
         Top             =   6000
         Width           =   540
      End
      Begin VB.Frame fra采购计划 
         Caption         =   "其他控制"
         Height          =   1485
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Width           =   7485
         Begin VB.ComboBox cbo供应商选择 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   2655
         End
         Begin VB.ComboBox cbo供应商范围 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label lbl供应商选择 
            AutoSize        =   -1  'True
            Caption         =   "供应商默认选择"
            Height          =   180
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lbl供应商范围 
            AutoSize        =   -1  'True
            Caption         =   "供应商选择范围"
            Height          =   180
            Left            =   360
            TabIndex        =   8
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "    药品采购计划编辑界面中药品供应商的默认处理，以及手工选择供应商时的可选范围。"
            Height          =   855
            Left            =   4680
            TabIndex        =   7
            Top             =   360
            Width           =   2085
         End
      End
      Begin VB.Frame fra上次采购信息 
         Caption         =   "上次采购信息来源方式"
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   240
         TabIndex        =   17
         Top             =   5280
         Visible         =   0   'False
         Width           =   7485
         Begin VB.OptionButton opt取成本价方式 
            Caption         =   "优先从上一次入库业务中取成本价等信息"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   6615
         End
         Begin VB.OptionButton opt取成本价方式 
            Caption         =   "优先从当前库房的库存最近批次中取成本价等信息"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   6615
         End
      End
      Begin VB.Frame fra盘点时间范围 
         Caption         =   "盘点时间范围设置"
         Height          =   735
         Left            =   240
         TabIndex        =   58
         Top             =   5280
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txt盘点时间 
            Height          =   300
            Left            =   1560
            TabIndex        =   60
            Top             =   240
            Width           =   705
         End
         Begin MSComCtl2.UpDown UpD盘点时间 
            Height          =   300
            Left            =   2266
            TabIndex        =   59
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txt盘点时间"
            BuddyDispid     =   196620
            OrigLeft        =   1800
            OrigTop         =   360
            OrigRight       =   2055
            OrigBottom      =   735
            Max             =   90
            Enabled         =   -1  'True
         End
         Begin VB.Label lblday 
            Caption         =   "天"
            Height          =   195
            Left            =   2880
            TabIndex        =   61
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Frame fra移库流程控制 
         Caption         =   "移库流程控制"
         Height          =   1485
         Left            =   240
         TabIndex        =   10
         Top             =   5280
         Width           =   7485
         Begin VB.CheckBox chk移库流程控制 
            Caption         =   "移库时需要备药、发送、接收这一过程。"
            Height          =   180
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Value           =   1  'Checked
            Width           =   6945
         End
         Begin VB.CheckBox chkRequestStrike 
            Caption         =   "移库冲销时，移入库房需要先申请冲销"
            Height          =   180
            Left            =   180
            TabIndex        =   11
            Top             =   1080
            Width           =   5895
         End
         Begin VB.Label Label1 
            Caption         =   "注意：如果不勾选，那么在填写移库单后，增加一个审核操作，审核后自动完成备药、发送、接收这一过程。审核前可以修改单据。"
            Height          =   375
            Left            =   450
            TabIndex        =   13
            Top             =   540
            Width           =   6945
         End
      End
      Begin VB.CheckBox chkALLPlanPoint 
         Caption         =   "全院计划不管站点"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   6000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fra成本价 
         Caption         =   "成本价来源方式"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   240
         TabIndex        =   54
         Top             =   5280
         Width           =   7665
         Begin VB.OptionButton opt成本来源 
            Caption         =   $"frm参数设置.frx":0044
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   56
            Top             =   277
            Width           =   3735
         End
         Begin VB.OptionButton opt成本来源 
            Caption         =   "根据原料药品的成本价计算"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraCheck 
         Caption         =   "选择校验方式"
         Height          =   615
         Left            =   -74760
         TabIndex        =   47
         Top             =   5160
         Width           =   7350
         Begin VB.OptionButton optCheck 
            Caption         =   "校验未通过时禁止保存"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   49
            Top             =   280
            Width           =   2175
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "校验未通过时提醒"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   48
            Top             =   280
            Width           =   2175
         End
      End
      Begin VB.Frame fra库房选择 
         Caption         =   "库房选择"
         Height          =   1665
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   3675
         Begin VB.CheckBox chkStock 
            Caption         =   "允许选择库房"
            Height          =   375
            Left            =   210
            TabIndex        =   44
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label Label4 
            Caption         =   "    如果选择库房，则在单据中有'所有库房'权限人就可以选择不同库房；否则，不能选择库房。"
            Height          =   615
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "药品单位"
         Enabled         =   0   'False
         Height          =   1665
         Left            =   3960
         TabIndex        =   37
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label lblUnitComment 
            Caption         =   "    请选择一种药品单位，在单据输入中，所有药品将用这种单位。"
            Height          =   405
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   3315
         End
         Begin VB.Label lbl盘点表 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "大包装"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   41
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl盘点单 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "小包装"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   40
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "控制"
         Height          =   2835
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         Width           =   3675
         Begin VB.Frame fra计算库存数量时方式 
            Caption         =   "计算库存数量时方式"
            Height          =   735
            Left            =   120
            TabIndex        =   72
            Top             =   1680
            Visible         =   0   'False
            Width           =   3450
            Begin VB.OptionButton Opt采用实际数量 
               Caption         =   "采用实际数量"
               Height          =   180
               Left            =   120
               TabIndex        =   74
               Top             =   375
               Width           =   1380
            End
            Begin VB.OptionButton Opt采用可用数量 
               Caption         =   "采用可用数量"
               Height          =   180
               Left            =   1680
               TabIndex        =   73
               Top             =   375
               Width           =   1440
            End
         End
         Begin VB.CheckBox chk分段加成入库 
            Caption         =   "时价药品按分段加成入库"
            Height          =   255
            Left            =   165
            TabIndex        =   71
            Top             =   2040
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk取上次售价 
            Caption         =   "时价药品入库时取上次售价"
            Height          =   255
            Left            =   165
            TabIndex        =   70
            Top             =   1800
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk加成入库 
            Caption         =   "时价药品以加成率入库"
            Height          =   255
            Left            =   165
            TabIndex        =   69
            Top             =   1560
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk招标药品 
            Caption         =   "招标药品可选择非中标单位入库"
            Height          =   255
            Left            =   165
            TabIndex        =   30
            Top             =   1335
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk服务对象 
            Caption         =   "忽略药品服务对象"
            Height          =   255
            Left            =   165
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox chk时价调价 
            Caption         =   "时价药品按批次调价"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   63
            Top             =   960
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CheckBox chk限价提示 
            Caption         =   "新成本价、新售价超过限价时提示"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   1320
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.CheckBox chkStopDrug 
            Caption         =   "盘点停用药品"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   57
            Top             =   1100
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "审核后打印"
            Height          =   255
            Left            =   2010
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFixPrice 
            Caption         =   "定价采购"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   34
            Top             =   495
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox Chk允许修改批发价 
            Caption         =   "修改采购限价"
            Height          =   255
            Left            =   2010
            TabIndex        =   33
            Top             =   513
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmd打印设置 
            Caption         =   "打印设置(&P)"
            Height          =   315
            Left            =   270
            TabIndex        =   32
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CheckBox chk外购NO 
            Caption         =   "修改单据号"
            Height          =   255
            Left            =   165
            TabIndex        =   29
            Top             =   780
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "存盘后打印"
            Height          =   255
            Left            =   165
            TabIndex        =   36
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox Chk存储库房 
            Caption         =   "允许盘点没有设置存储库房的药品"
            Height          =   255
            Left            =   165
            TabIndex        =   46
            Top             =   540
            Visible         =   0   'False
            Width           =   3360
         End
         Begin VB.CheckBox chkSendPrint 
            Caption         =   "发送后打印"
            Height          =   255
            Left            =   165
            TabIndex        =   52
            Top             =   495
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Frame Frame价格显示 
            Caption         =   "价格显示方式"
            Height          =   735
            Left            =   120
            TabIndex        =   25
            Top             =   705
            Visible         =   0   'False
            Width           =   3450
            Begin VB.OptionButton Opt混合 
               Caption         =   "成本价和售价"
               Height          =   180
               Left            =   1950
               TabIndex        =   27
               Top             =   375
               Width           =   1400
            End
            Begin VB.OptionButton Opt成本价 
               Caption         =   "成本价"
               Height          =   180
               Left            =   45
               TabIndex        =   28
               Top             =   375
               Width           =   900
            End
            Begin VB.OptionButton Opt售价 
               Caption         =   "售价"
               Height          =   180
               Left            =   1100
               TabIndex        =   26
               Top             =   375
               Width           =   720
            End
         End
         Begin VB.CheckBox chk留存领用 
            Caption         =   "按月留存领用"
            Height          =   255
            Left            =   165
            TabIndex        =   64
            Top             =   1680
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk退货发票金额 
            Caption         =   "退货时发票金额以零售金额为准"
            Height          =   255
            Left            =   165
            TabIndex        =   31
            Top             =   1059
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "排序方式"
         Height          =   2835
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   3675
         Begin VB.ComboBox Cbo列名 
            Height          =   300
            ItemData        =   "frm参数设置.frx":0066
            Left            =   120
            List            =   "frm参数设置.frx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox Cbo方向 
            Height          =   300
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "    本参数的设置，将影响所有编辑窗体中单据的显示内容的排序方式。缺省：按用户输入的顺序显示各单据的内容"
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   180
            TabIndex        =   23
            Top             =   1080
            Width           =   3345
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
         Height          =   4125
         Left            =   -74760
         TabIndex        =   50
         Top             =   960
         Width           =   7095
         _cx             =   12515
         _cy             =   7276
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   13
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm参数设置.frx":006A
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
         VirtualData     =   0   'False
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
      Begin VB.Frame fra同步处理库存 
         Caption         =   "同步处理库存控制"
         ForeColor       =   &H00800000&
         Height          =   1485
         Left            =   240
         TabIndex        =   14
         Top             =   5280
         Visible         =   0   'False
         Width           =   7485
         Begin VB.CheckBox chk处理库存 
            Caption         =   "药品质量管理审核时同步减少库存"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "说明：如果勾选此选项，相当于在审核后自动完成其他出库操作；要实现该功能，必须确保已先设置了其他出库的入出类别。"
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Width           =   7020
         End
      End
      Begin VB.Label lbl查询天数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查询天数"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   67
         Top             =   6060
         Width           =   720
      End
      Begin VB.Label lbl天数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "天"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   68
         Top             =   6060
         Width           =   180
      End
      Begin VB.Label lblComment 
         Caption         =   "    说明：药品外购入库编辑单据时是否校供应商的信息是否完整，及资质是否过期。请选择需要进行校验的项目，并双击“校验”列打勾。"
         Height          =   540
         Left            =   -74760
         TabIndex        =   51
         Top             =   480
         Width           =   7140
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6930
      TabIndex        =   2
      Top             =   7560
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5730
      TabIndex        =   1
      Top             =   7560
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   1100
   End
End
Attribute VB_Name = "frm参数设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Dim mstrPrivs As String
Dim mlngModul As Long
Dim mblnSetPara As Boolean      '是否具有参数设置权限
Private mint盘点时间 As Integer  '用来记录设置的盘点时间范围

Private Sub Cbo列名_Click()
    If Cbo方向.ListCount < 1 Then Exit Sub
    Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
    If Not Cbo方向.Enabled Then Cbo方向.ListIndex = 0
End Sub

Private Sub chkRequestStrike_Click()
    '当变为不需要申请时，要检查是否有未审核的冲销申请单，如果有则不能改变
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If chkRequestStrike.Value = 0 Then
        If MsgBox("即将检查是否存在未审核的冲销申请单，可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '该功能是10.20版本新增，增加一个条件填制日期范围，避免全表扫描
            gstrSQL = "Select 1 From 药品收发记录 Where 单据 = 6 And Mod(记录状态, 3) = 2 And 审核日期 Is Null " & _
                " And 填制日期 Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Rownum = 1"
            
            DoEvents
            zlCommFun.ShowFlash "正在查找数据,请稍候...", Me
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否有未审核的冲销申请单")
            
            DoEvents
            zlCommFun.StopFlash
            
            If rsTemp.RecordCount > 0 Then
                MsgBox "存在未审核的冲销申请单，不能改变此参数！", vbInformation, gstrSysName
                chkRequestStrike.Value = 1
            End If
        Else
            chkRequestStrike.Value = 1
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk分段加成入库_Click()
    If chk分段加成入库.Value = 1 Then
        chk加成入库.Value = 0
        chk取上次售价.Value = 0
    End If
End Sub

Private Sub chk加成入库_Click()
    If chk加成入库.Value = 1 Then
        chk取上次售价.Value = 0
        chk分段加成入库.Value = 0
    End If
End Sub

Private Sub chk留存领用_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
    
    If chk留存领用.Value = 0 Then
        gstrSQL = "Select 期间 From 药品留存 Where Length(期间) > 4"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTemp.RecordCount > 0 Then
            MsgBox "按月留存模式下已经产生数据，不能修改！", vbInformation, gstrSysName
            chk留存领用.Value = 1
        End If
    End If
    Exit Sub
errh:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chk取上次售价_Click()
    If chk取上次售价.Value = 1 Then
        chk加成入库.Value = 0
        chk分段加成入库.Value = 0
    End If
End Sub

Private Sub chk移库流程控制_Click()
    If chk移库流程控制.Value = 1 Then
        chkSendPrint.Visible = True
    Else
        chkSendPrint.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    
    If ISValid = False Then Exit Sub
    
    Select Case mlngModul
        Case 1300   '药品外购入库管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            
            zlDatabase.SetPara "定价采购", chkFixPrice.Value, glngSys, mlngModul
            zlDatabase.SetPara "修改外购单据号", chk外购NO.Value, glngSys, mlngModul
            zlDatabase.SetPara "修改采购限价", Chk允许修改批发价.Value, glngSys, mlngModul
            zlDatabase.SetPara "招标药品可选择非中标单位入库", chk招标药品.Value, glngSys, mlngModul
            zlDatabase.SetPara "退货发票金额", chk退货发票金额.Value, glngSys, mlngModul
            zlDatabase.SetPara "取上次采购价方式", IIf(opt取成本价方式(0).Value, "0", "1"), glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
            
            zlDatabase.SetPara "时价药品以加价率入库", chk加成入库.Value, glngSys, mlngModul
            zlDatabase.SetPara "时价药品入库时取上次售价", chk取上次售价.Value, glngSys, mlngModul
            zlDatabase.SetPara "时价药品入库采用分段加成", chk分段加成入库.Value, glngSys, mlngModul
            
            Save资质校验
        Case 1301   '药品自制入库管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "药品自制入库成本价计算方式", IIf(opt成本来源(0).Value = True, "0", "1"), glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1302   '药品其他入库管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
            
            zlDatabase.SetPara "时价药品以加价率入库", chk加成入库.Value, glngSys, mlngModul
            zlDatabase.SetPara "时价药品入库时取上次售价", chk取上次售价.Value, glngSys, mlngModul
            zlDatabase.SetPara "时价药品入库采用分段加成", chk分段加成入库.Value, glngSys, mlngModul
        Case 1303   '药品库存差价调整管理
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1304   '药品移库管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "发送打印", IIf(chkSendPrint.Value = 1, "1", "0"), glngSys, mlngModul
            
            zlDatabase.SetPara "移库流程", chk移库流程控制.Value, glngSys, mlngModul
            zlDatabase.SetPara "冲销申请", chkRequestStrike.Value, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1305   '药品领用管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "按月留存领用", IIf(chk留存领用.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1306   '药品其他出库管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1307   '药品盘点管理
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "小包装单位", CboUnit1.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
                       
            zlDatabase.SetPara "存储库房", IIf(Chk存储库房.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "忽略药品服务对象", IIf(chk服务对象.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "盘已停用的药品", IIf(chkStopDrug.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "盘点时间范围设置", txt盘点时间.Text, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1330   '药品计划管理
            zlDatabase.SetPara "是否选择库房", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "价格显示方式", IIf(Opt成本价.Value = True, "0", IIf(Opt售价.Value = True, "1", "2")), glngSys, mlngModul
            zlDatabase.SetPara "存盘打印", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "审核打印", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "供应商默认选择", cbo供应商选择.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "供应商选择范围", cbo供应商范围.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "全院计划不管站点", IIf(chkALLPlanPoint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
            zlDatabase.SetPara "计算库存数量时方式", IIf(Opt采用实际数量.Value = True, "0", "1"), glngSys, mlngModul
        Case 1331   '药品质量管理
            zlDatabase.SetPara "审核时减少库存", chk处理库存.Value, glngSys, mlngModul
            zlDatabase.SetPara "查询天数", Val(txt查询天数.Text), glngSys, mlngModul
        Case 1333 '药品调价管理
            zlDatabase.SetPara "排序", CStr(Cbo列名.ListIndex) & CStr(Cbo方向.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "时价药品按批次调价", chk时价调价.Value, glngSys, mlngModul
            zlDatabase.SetPara "限价提示", chk限价提示.Value, glngSys, mlngModul
            zlDatabase.SetPara "药品单位", cboUnit.ListIndex, glngSys, mlngModul
    End Select
           
    Unload Me
End Sub

Private Function ISValid() As Boolean
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    '资质校验
    If tabMain.TabVisible(1) = True Then
        blnAllUnCheck = True
        With vsfCheck
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("校验")) <> "" Then
                    blnAllUnCheck = False
                    Exit For
                End If
            Next
        End With
        
        '如果选择了校验项目，则必须选择校验方式
        If blnAllUnCheck = False And optCheck(0).Value = 0 And optCheck(1).Value = 0 Then
            MsgBox "请选择资质校验方式！", vbExclamation, gstrSysName
            tabMain.Tab = 1
            If vsfCheck.Enabled Then vsfCheck.SetFocus
            Exit Function
        End If
    End If
    
    If Val(txt查询天数.Text) > 7 Then
        If MsgBox("查询时间大于7天可能会导致查询很慢，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt查询天数.SetFocus
            zlControl.TxtSelAll txt查询天数
            Exit Function
        End If
    End If
    If Val(txt查询天数.Text) = 0 Then
        MsgBox "查询时间必须大于0，请重新输入！", vbInformation, gstrSysName
        txt查询天数.SetFocus
        zlControl.TxtSelAll txt查询天数
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub Save资质校验()
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    If mlngModul <> 1300 Then Exit Sub
    
    blnAllUnCheck = True
    
    '保存资质校验项目和方式，格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    With vsfCheck
        For i = 1 To .rows - 1
            strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("类别")) & "," & .TextMatrix(i, .ColIndex("校验项目")) & "," & _
                IIf(.TextMatrix(i, .ColIndex("校验")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("校验")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If blnAllUnCheck = True Then
        strCheck = "0|" & strCheck
    ElseIf optCheck(0).Value = True Then
        strCheck = "2|" & strCheck
    Else
        strCheck = "1|" & strCheck
    End If
        
    Call zlDatabase.SetPara("资质校验", strCheck, glngSys, mlngModul)
End Sub
Public Sub 设置参数(frmParent As Object, ByVal strPrivs As String, ByVal lngModual As Long, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mstrPrivs = strPrivs
    mlngModul = lngModual
    Dim str单据打印 As String
    Dim int查询天数 As Integer
    
    '通用（私有模块）
    Dim int是否选择库房 As Integer
    Dim str排序 As String
    Dim int存盘打印 As Integer
    Dim int审核打印 As Integer
        
    '用于主要流通模块（私有模块）
    Dim int药品单位 As Integer
    Dim int成本价来源 As Integer
        
    '用于盘点（私有模块）
    Dim int小包装单位 As Integer
        
    '用于药品计划（私有模块）
    Dim int价格显示方式 As Integer
    Dim int供应商选择 As Integer
    Dim int供应商范围 As Integer
    Dim intPlanPoint As Integer
    Dim int计算库存数量时方式 As Integer
    
    '用于外购入库（公共模块）
    Dim int定价采购 As Integer
    Dim int修改外购单据号 As Integer
    Dim int修改批发价 As Integer
    Dim int招标药品 As Integer
    Dim int退货发票金额 As Integer
    Dim int取上次采购价方式 As Integer
    Dim int加成率入库 As Integer    '其他出库也用这个变量
    Dim int取上次售价 As Integer    '其他出库也用这个变量
    Dim int分段加成入库 As Integer  '其他出库也用这个变量
    
    '用于移库（公共模块）
    Dim int移库流程 As Integer
    Dim int冲销申请 As Integer
    
    '(私有)
    Dim int发送打印 As Integer
    
    '用于盘点（公共模块）
    Dim int存储库房 As Integer
    Dim int检查可用数量 As Integer
    Dim int服务对象 As Integer
    Dim int盘点停用 As Integer
    
    '用于质量管理（公共模块）
    Dim int处理库存 As Integer
    
    '用于领用
    Dim int留存领用 As Integer
    
    On Error Resume Next
    
    mblnSetPara = IsHavePrivs(mstrPrivs, "参数设置")
    
    '取参数值
    Select Case mlngModul
        Case 1300   '药品外购入库管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            
            int定价采购 = Val(zlDatabase.GetPara("定价采购", glngSys, mlngModul, 0, Array(chkFixPrice), mblnSetPara))
            int修改外购单据号 = Val(zlDatabase.GetPara("修改外购单据号", glngSys, mlngModul, 0, Array(chk外购NO), mblnSetPara))
            int修改批发价 = Val(zlDatabase.GetPara("修改采购限价", glngSys, mlngModul, 0, Array(Chk允许修改批发价), mblnSetPara))
            int招标药品 = Val(zlDatabase.GetPara("招标药品可选择非中标单位入库", glngSys, mlngModul, 0, Array(chk招标药品), mblnSetPara))
            int退货发票金额 = Val(zlDatabase.GetPara("退货发票金额", glngSys, mlngModul, 1, Array(chk退货发票金额), mblnSetPara))
            int取上次采购价方式 = Val(zlDatabase.GetPara("取上次采购价方式", glngSys, mlngModul, 0, Array(fra上次采购信息, opt取成本价方式(0), opt取成本价方式(1)), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
            
            int加成率入库 = Val(zlDatabase.GetPara("时价药品以加价率入库", glngSys, mlngModul, 1, Array(chk加成入库), mblnSetPara))
            int取上次售价 = Val(zlDatabase.GetPara("时价药品入库时取上次售价", glngSys, mlngModul, 0, Array(chk取上次售价), mblnSetPara))
            int分段加成入库 = Val(zlDatabase.GetPara("时价药品入库采用分段加成", glngSys, mlngModul, 0, Array(chk分段加成入库), mblnSetPara))
            
            '参数规则检查
            If int加成率入库 = 1 Then
                int取上次售价 = 0
                int分段加成入库 = 0
            ElseIf int取上次售价 = 1 Then
                int加成率入库 = 0
                int分段加成入库 = 0
            ElseIf int分段加成入库 = 1 Then
                int加成率入库 = 0
                int取上次售价 = 0
            End If
        Case 1301   '药品自制入库管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int成本价来源 = Val(zlDatabase.GetPara("药品自制入库成本价计算方式", glngSys, mlngModul, 0, Array(fra成本价, opt成本来源(0), opt成本来源(1)), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1302   '药品其他入库管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
            
            int加成率入库 = Val(zlDatabase.GetPara("时价药品以加价率入库", glngSys, mlngModul, 1, Array(chk加成入库), mblnSetPara))
            int取上次售价 = Val(zlDatabase.GetPara("时价药品入库时取上次售价", glngSys, mlngModul, 0, Array(chk取上次售价), mblnSetPara))
            int分段加成入库 = Val(zlDatabase.GetPara("时价药品入库采用分段加成", glngSys, mlngModul, 0, Array(chk分段加成入库), mblnSetPara))
            
            '参数规则检查
            If int加成率入库 = 1 Then
                int取上次售价 = 0
                int分段加成入库 = 0
            ElseIf int取上次售价 = 1 Then
                int加成率入库 = 0
                int分段加成入库 = 0
            ElseIf int分段加成入库 = 1 Then
                int加成率入库 = 0
                int取上次售价 = 0
            End If
        Case 1303   '药品库存差价调整管理
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1304   '药品移库管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int发送打印 = Val(zlDatabase.GetPara("发送打印", glngSys, mlngModul, 0, Array(chkSendPrint), mblnSetPara))
            
            int移库流程 = Val(zlDatabase.GetPara("移库流程", glngSys, mlngModul, 1, Array(chk移库流程控制, Label1), mblnSetPara))
            int冲销申请 = Val(zlDatabase.GetPara("冲销申请", glngSys, mlngModul, 0, Array(chkRequestStrike), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1305   '药品领用管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int留存领用 = Val(zlDatabase.GetPara("按月留存领用", glngSys, mlngModul, 0, Array(chk留存领用), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1306   '药品其他出库管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1307   '药品盘点管理
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int小包装单位 = Val(zlDatabase.GetPara("小包装单位", glngSys, mlngModul, 0, Array(lbl盘点单, CboUnit1), mblnSetPara))
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
                        
            int存储库房 = Val(zlDatabase.GetPara("存储库房", glngSys, mlngModul, 0, Array(Chk存储库房), mblnSetPara))
            int服务对象 = Val(zlDatabase.GetPara("忽略药品服务对象", glngSys, mlngModul, 0, Array(chk服务对象), mblnSetPara))
            int盘点停用 = Val(zlDatabase.GetPara("盘已停用的药品", glngSys, mlngModul, 0, Array(chkStopDrug), mblnSetPara))
            mint盘点时间 = Val(zlDatabase.GetPara("盘点时间范围设置", glngSys, mlngModul, 30))
            txt盘点时间.Text = mint盘点时间
            UpD盘点时间.Value = mint盘点时间
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1330   '药品计划管理
            int是否选择库房 = Val(zlDatabase.GetPara("是否选择库房", glngSys, mlngModul, 0, Array(fra库房选择, chkStock, Label4), mblnSetPara))
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            int价格显示方式 = Val(zlDatabase.GetPara("价格显示方式", glngSys, mlngModul, 1, Array(Frame价格显示, Opt成本价, Opt售价, Opt混合), mblnSetPara))
            int存盘打印 = Val(zlDatabase.GetPara("存盘打印", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int审核打印 = Val(zlDatabase.GetPara("审核打印", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int供应商选择 = Val(zlDatabase.GetPara("供应商默认选择", glngSys, mlngModul, 0, Array(cbo供应商选择), mblnSetPara))
            int供应商范围 = Val(zlDatabase.GetPara("供应商选择范围", glngSys, mlngModul, 0, Array(cbo供应商范围), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            intPlanPoint = Val(zlDatabase.GetPara("全院计划不管站点", glngSys, mlngModul, 0, Array(chkALLPlanPoint), mblnSetPara))
            chkALLPlanPoint.Value = intPlanPoint
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
            int计算库存数量时方式 = Val(zlDatabase.GetPara("计算库存数量时方式", glngSys, mlngModul, 0, Array(fra计算库存数量时方式, Opt采用实际数量, Opt采用可用数量), mblnSetPara))
        Case 1331  '药品质量管理
            int处理库存 = Val(zlDatabase.GetPara("审核时减少库存", glngSys, mlngModul))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
        Case 1333 '药品调价管理
            str排序 = zlDatabase.GetPara("排序", glngSys, mlngModul, "00", Array(Frame5, Cbo列名, Cbo方向, Label5), mblnSetPara)
            chk时价调价.Value = Val(zlDatabase.GetPara("时价药品按批次调价", glngSys, 1333, 0, Array(Frame3, chk时价调价), mblnSetPara))
            chk限价提示.Value = Val(zlDatabase.GetPara("限价提示", glngSys, 1333, 1, Array(Frame3, chk限价提示), mblnSetPara))
            int药品单位 = Val(zlDatabase.GetPara("药品单位", glngSys, mlngModul, 0, Array(lbl盘点表, cboUnit), mblnSetPara))
            int查询天数 = Val(zlDatabase.GetPara("查询天数", glngSys, mlngModul, 7))
    End Select
    
    txt查询天数.Text = int查询天数
    If strFunction = "药品计划管理" Then
        str单据打印 = "采购计划打印"
    Else
        str单据打印 = "单据打印"
    End If
    
    '装入缺省数据
    With Cbo列名
        .Clear
        .AddItem "输入顺序"
        .ItemData(.NewIndex) = 0
        .AddItem "编码"
        .ItemData(.NewIndex) = 1
        .AddItem "药品名称"
        .ItemData(.NewIndex) = 2
        
        If InStr("药品盘点管理/药品移库管理/药品领用管理/药品其他出库管理", strFunction) > 0 Then
            .AddItem "库房货位"
            .ItemData(.NewIndex) = 3
        End If
     
        .ListIndex = 0
    End With
    With Cbo方向
        .Clear
        .AddItem "升序"
        .ItemData(.NewIndex) = 0
        .AddItem "降序"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    '取排序字段及方向，如果为缺省，则置cbo方向.Enabled=False
    Cbo列名.ListIndex = Mid(str排序, 1, 1)
    Cbo方向.ListIndex = Right(str排序, 1)
    Cbo方向.Enabled = Not (Cbo列名.ListIndex = 0)
    
    If int存盘打印 = 0 Then
        chkSavePrint.Value = 0
    Else
        chkSavePrint.Value = 1
    End If
    
    If int审核打印 = 0 Then
        chkVerifyPrint.Value = 0
    Else
        chkVerifyPrint.Value = 1
    End If


    If int是否选择库房 = 0 Then
        chkStock.Value = 0
    Else
        chkStock.Value = 1
    End If
    
    If int存储库房 = 0 Then
        Chk存储库房.Value = 0
    Else
        Chk存储库房.Value = 1
    End If
    If int成本价来源 = 0 Then
        opt成本来源(0).Value = 1
        opt成本来源(1).Value = 0
    Else
        opt成本来源(0).Value = 0
        opt成本来源(1).Value = 1
    End If
    
    If int留存领用 = 0 Then
        chk留存领用.Value = 0
    Else
        chk留存领用.Value = 1
    End If
    
    chk服务对象.Value = IIf(int服务对象 = 1, 1, 0)
    chkStopDrug.Value = IIf(int盘点停用 = 1, 1, 0)
    
    fra移库流程控制.Visible = False
    fra上次采购信息.Visible = False
    fra采购计划.Visible = False
    chkALLPlanPoint.Visible = False
    fra计算库存数量时方式.Visible = False
    
    If mstrFunction = "药品盘点管理" Then
        If glngSys \ 100 = 8 Then
            With CboUnit1
                .AddItem "采购单位"
                .AddItem "售价单位"
            End With
        Else
            With CboUnit1
                .AddItem "和大包装相同"
                .AddItem "药库单位"
                .AddItem "门诊单位"
                .AddItem "住院单位"
                .AddItem "售价单位"
            End With
        End If
        CboUnit1.ListIndex = int小包装单位
        lblUnitComment.Caption = "    请选择盘点时的大小包装，盘点单及盘点表编辑时按所选包装进行盘点。"
    Else
        CboUnit1.Visible = False
        lbl盘点表.Visible = False
        lbl盘点单.Visible = False
        cboUnit.Left = lbl盘点表.Left
        cboUnit.Width = Frame2.Width - cboUnit.Left - 250
        lblUnitComment.Top = lbl盘点单.Top
    End If
    
    With cboUnit
        .Clear
        If glngSys \ 100 = 8 Then
            .AddItem "缺省（当前库房对应的单位）"
            .AddItem "采购单位"
            .AddItem "售价单位"
        Else
            If mlngModul <> 1333 Then   '调价不需要库房
                .AddItem "缺省（当前库房对应的单位）"
            End If
            .AddItem "药库单位"
            .AddItem "门诊单位"
            .AddItem "住院单位"
            .AddItem "售价单位"
        End If
        .ListIndex = int药品单位
    End With
    
    If strFunction = "药品外购入库管理" Then
        chkFixPrice.Visible = True
        chk外购NO.Visible = True
        Chk允许修改批发价.Visible = True
        chk退货发票金额.Visible = True
        chk招标药品.Visible = True
        chkFixPrice.Value = int定价采购
        chk外购NO.Value = int修改外购单据号
        Chk允许修改批发价.Value = int修改批发价
        chk退货发票金额.Value = int退货发票金额
        chk招标药品.Value = int招标药品
        
        fra上次采购信息.Visible = True
        If int取上次采购价方式 = 1 Then
            opt取成本价方式(1).Value = True
        Else
            opt取成本价方式(0).Value = True
        End If
        
        chk加成入库.Visible = True
        chk取上次售价.Visible = True
        chk分段加成入库.Visible = True
        
        chk加成入库.Value = int加成率入库
        chk取上次售价.Value = int取上次售价
        chk分段加成入库.Value = int分段加成入库
        
        lbl查询天数.Move fra上次采购信息.Left, fra上次采购信息.Top + fra上次采购信息.Height + 200
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    If strFunction = "药品其他入库管理" Then
        chk加成入库.Visible = True
        chk取上次售价.Visible = True
        chk分段加成入库.Visible = True
        
        chk加成入库.Value = int加成率入库
        chk取上次售价.Value = int取上次售价
        chk分段加成入库.Value = int分段加成入库
    End If
    
'    Frame2.Enabled = (strFunction = "药品申领管理" Or strFunction = "药品移库管理" Or strFunction = "药品领用管理")
    If strFunction <> "药品质量管理" Then
        Frame2.Enabled = True
    End If
    
    chkStopDrug.Visible = False
    If strFunction = "药品盘点管理" Then
        Frame2.Enabled = True
        cboUnit.Enabled = False
        Chk存储库房.Visible = True
        chk服务对象.Visible = True
        chkStopDrug.Visible = True
    End If
    
    fra库房选择.Enabled = (InStr(1, "药品盘点管理,库存差价调整管理", strFunction) = 0)
    If fra库房选择.Enabled = False Then
        chkStock.Enabled = False
    End If
    
    If strFunction = "药品移库管理" Then
        chk移库流程控制.Value = int移库流程
        chkRequestStrike.Value = int冲销申请
        fra移库流程控制.Visible = True
        
        chkSendPrint.Value = IIf(int发送打印 = 1, 1, 0)
        chkSendPrint.Visible = (chk移库流程控制.Value = 1)
        
        lbl查询天数.Move fra移库流程控制.Left, fra移库流程控制.Top + fra移库流程控制.Height + 150
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    If strFunction = "药品计划管理" Then
        cbo供应商选择.Clear
        cbo供应商选择.AddItem "1-取上次入库供应商"
        cbo供应商选择.AddItem "2-取合同单位"
        cbo供应商选择.ListIndex = IIf(int供应商选择 < 0 Or int供应商选择 > 1, 0, int供应商选择)
        
        cbo供应商范围.Clear
        cbo供应商范围.AddItem "1-所有供应商"
        cbo供应商范围.AddItem "2-中标单位"
        cbo供应商范围.ListIndex = IIf(int供应商范围 < 0 Or int供应商范围 > 1, 0, int供应商范围)
        
        fra采购计划.Visible = True
        chkALLPlanPoint.Visible = True
        fra计算库存数量时方式.Visible = True
        chkALLPlanPoint.Top = fra采购计划.Top + fra采购计划.Height + 113
        chkALLPlanPoint.Left = fra采购计划.Left
        
        fra计算库存数量时方式.Top = Frame价格显示.Top + Frame价格显示.Height + 150
        fra计算库存数量时方式.Left = Frame价格显示.Left
        
        lbl查询天数.Move chkALLPlanPoint.Left + chkALLPlanPoint.Width + 150, fra采购计划.Top + fra采购计划.Height + 150
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    If Frame3.Enabled = True Then
        If strFunction = "药品计划管理" Then
            Frame价格显示.Visible = True
            If int价格显示方式 = 0 Then
                Opt成本价.Value = True
            ElseIf int价格显示方式 = 1 Then
                Opt售价.Value = True
            Else
                Opt混合.Value = True
            End If
            
            If int计算库存数量时方式 = 0 Then
                Opt采用实际数量.Value = True
            Else
                Opt采用可用数量.Value = True
            End If
        End If
    End If
    
    If mlngModul = 1331 Then    '药品质量管理
        chk处理库存.Value = int处理库存
    End If
    
    '界面控制
    If strFunction <> "药品移库管理" And strFunction <> "药品外购入库管理" And strFunction <> "药品计划管理" Then
        fra移库流程控制.Visible = False
        
        tabMain.Height = tabMain.Height - fra移库流程控制.Height
        
        cmdHelp.Top = cmdHelp.Top - fra移库流程控制.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fra移库流程控制.Height
    End If
    
    If strFunction = "药品质量管理" Then
        fra库房选择.Visible = False
        Frame2.Visible = False
        Frame5.Visible = False
        Frame3.Visible = False
        fra移库流程控制.Visible = False
        fra上次采购信息.Visible = False
        
        fra同步处理库存.Visible = True
        fra同步处理库存.Top = 580
        
        lbl查询天数.Move fra同步处理库存.Left, fra同步处理库存.Top + fra同步处理库存.Height + 200
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
        
        tabMain.Height = tabMain.Height - fra同步处理库存.Height
        
        cmdHelp.Top = cmdHelp.Top - fra同步处理库存.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fra同步处理库存.Height
    End If
    If strFunction = "药品盘点管理" Then
        tabMain.Height = fra盘点时间范围.Height + fra盘点时间范围.Top + 200
        Me.Height = fra盘点时间范围.Top + fra盘点时间范围.Height + 1300
        fra盘点时间范围.Visible = True
        cmdHelp.Top = tabMain.Height + 250
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        lbl查询天数.Move Frame3.Left, Frame3.Top + Frame3.Height + 500
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    '资质校验页面
    tabMain.TabVisible(1) = strFunction = "药品外购入库管理"
    If tabMain.TabVisible(1) = True Then
        With vsfCheck
            .MergeCol(0) = True
            .MergeCells = flexMergeRestrictColumns
        End With
        fraCheck.Top = tabMain.Height - fraCheck.Height - 100
        vsfCheck.Height = fraCheck.Top - vsfCheck.Top - 100
        
        Load资质校验
    End If
    
    If strFunction = "药品自制入库管理" Then
        fra成本价.Visible = True
        tabMain.Height = tabMain.Height + fra成本价.Height + 100
        frm参数设置.Height = tabMain.Height + cmdOK.Height + 800
        cmdHelp.Top = frm参数设置.Height - 900
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        lbl查询天数.Move fra成本价.Left, fra成本价.Top + fra成本价.Height + 200
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    Else
        fra成本价.Visible = False
    End If
    
    If strFunction = "药品调价管理" Then
        fra库房选择.Visible = False
        Frame5.Visible = False
        Frame2.Left = fra库房选择.Left
        Frame2.Top = fra库房选择.Top
        
        Frame3.Top = Frame2.Top
        
        Frame2.Height = Frame3.Height
        fra同步处理库存.Visible = False
        chkSavePrint.Visible = False
        chkVerifyPrint.Visible = False
        Frame5.Enabled = False
        chk时价调价.Visible = True
        chk限价提示.Visible = True
        chk时价调价.Move chkSavePrint.Left, chkSavePrint.Top
        chk限价提示.Move chk时价调价.Left, chk时价调价.Top + chk时价调价.Height + 100
        
        tabMain.Height = Frame2.Height + cmdOK.Height + 500
        cmdHelp.Top = tabMain.Height + tabMain.Top + 100
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = tabMain.Height + tabMain.Top + 1000
        
    End If
    If strFunction = "药品领用管理" Then
        chk留存领用.Visible = True
        chk留存领用.Left = 165
        chk留存领用.Top = chkSavePrint.Top + chkSavePrint.Height + 50
        lbl查询天数.Move Frame5.Left, Frame5.Top + Frame5.Height + 200
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    If mlngModul = 1302 Or mlngModul = 1303 Or mlngModul = 1306 Then
        '1302 :其他入库;1303:库存差价; 1306：其他出库
        lbl查询天数.Move Frame5.Left, Frame5.Top + Frame5.Height + 200
        txt查询天数.Move lbl查询天数.Left + lbl查询天数.Width + 100, lbl查询天数.Top - 50
        lbl天数.Move txt查询天数.Left + txt查询天数.Width + 50, lbl查询天数.Top
    End If
    
    frm参数设置.Show vbModal, frmParent
End Sub

Private Sub Load资质校验()
    Dim i As Integer
    Dim n As Integer
    Dim strCheck As String
    Dim intCheckType As Integer
    Dim arrColumn
    
    On Error Resume Next
    
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    strCheck = zlDatabase.GetPara("资质校验", glngSys, mlngModul, "", Array(vsfCheck, fraCheck), mblnSetPara)
    
    If strCheck <> "" Then
        If InStr(1, strCheck, "|") > 0 Then
            '校验方式：0-不检查；1－提醒；2－禁止
            intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
            If intCheckType = 2 Then
                optCheck(0).Value = True
            ElseIf intCheckType = 1 Then
                optCheck(1).Value = True
            End If
            
            strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)
             
            If strCheck <> "" Then
                strCheck = strCheck & ";"
                arrColumn = Split(strCheck, ";")
                For n = 0 To UBound(arrColumn)
                    If arrColumn(n) <> "" Then
                        With vsfCheck
                            For i = 1 To .rows - 1
                                If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("类别")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("校验项目")) Then
                                    If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                        .TextMatrix(i, .ColIndex("校验")) = "√"
                                    End If
                                End If
                            Next
                        End With
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "药品外购入库管理"
        strBill = "ZL1_BILL_1300"
    Case "药品其他入库管理"
        strBill = "ZL1_BILL_1302"
    Case "药品自制入库管理"
        strBill = "ZL1_BILL_1301"
    Case "库存差价调整管理"
        strBill = "ZL1_BILL_1303"
    Case "药品移库管理"
        strBill = "ZL1_BILL_1304"
    Case "药品领用管理"
        strBill = "ZL1_BILL_1305"
    Case "药品其他出库管理"
        strBill = "ZL1_BILL_1306"
    Case "药品盘点管理"
        strBill = "ZL1_BILL_1307"
    Case "药品计划管理"
        strBill = "zl1_bill_1330"
    Case "药品调价管理"
        strBill = "ZL1_BILL_1333"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.cmd打印设置.Caption = "票据《" & Mid(mstrFunction, 1, Len(mstrFunction) - 2) & "单》打印设置"
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If vsfCheck.Enabled = True Then vsfCheck.SetFocus
    End If
End Sub

Private Sub txt查询天数_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txt查询天数_Validate(Cancel As Boolean)
    If Val(txt查询天数.Text) = 0 Then
        MsgBox "查询时间必须大于0，请重新输入！", vbInformation, gstrSysName
        Cancel = False
        txt查询天数.SetFocus
        zlControl.TxtSelAll txt查询天数
    End If
End Sub

Private Sub txt盘点时间_Change()
    txt盘点时间.Text = IIf(txt盘点时间.Text = "", "0", txt盘点时间.Text) '防止文本为空
    UpD盘点时间.Value = IIf(Val(txt盘点时间.Text) > 90, Val(Mid(txt盘点时间.Text, 1, Len(txt盘点时间.Text) - 1)), Val(txt盘点时间.Text))
End Sub

Private Sub txt盘点时间_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt盘点时间_Validate(Cancel As Boolean)
    If Val(txt盘点时间.Text) > 90 Then
        MsgBox "盘点时间范围不能大于3个月！", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UpD盘点时间_Change()
    txt盘点时间.Text = UpD盘点时间.Value
    txt盘点时间.SelStart = Len(txt盘点时间.Text) '定位到文本末尾
End Sub

Private Sub vsfCheck_DblClick()
    With vsfCheck
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("校验") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "√" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            .TextMatrix(.Row, .Col) = "√"
        End If
    End With
End Sub


