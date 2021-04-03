VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPIVAParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "静脉输液配置中心参数设置"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11715
   Icon            =   "frmPIVAParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPRI 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   3120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdYes 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmdNO 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6DDC
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView lvwPRI 
         Height          =   1305
         Left            =   120
         TabIndex        =   47
         ToolTipText     =   "双击或按回车键确认"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.PictureBox pic剂型 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   2535
      TabIndex        =   40
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmd剂型Cancel 
         Height          =   360
         Left            =   1560
         Picture         =   "frmPIVAParaSet.frx":6F26
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1560
         Width           =   810
      End
      Begin VB.CommandButton cmd剂型Ok 
         Height          =   360
         Left            =   720
         Picture         =   "frmPIVAParaSet.frx":7070
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1560
         Width           =   810
      End
      Begin MSComctlLib.ListView Lvw药品剂型 
         Height          =   1305
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "双击或按回车键确认"
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2302
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "基础设置(&0)"
      TabPicture(0)   =   "frmPIVAParaSet.frx":D8C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra打印控制"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra操作控制"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "工作批次(&1)"
      TabPicture(1)   =   "frmPIVAParaSet.frx":D8DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAdd"
      Tab(1).Control(1)=   "cmdDel"
      Tab(1).Control(2)=   "vsfBatch"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "辅助控制(&2)"
      TabPicture(2)   =   "frmPIVAParaSet.frx":D8FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabPrice"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdNext"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdLast"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "picprice"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fra医嘱类型"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fra输液量"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fra输液单控制"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblprice"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "给药途径(&3)"
      TabPicture(3)   =   "frmPIVAParaSet.frx":D916
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chk给药途径"
      Tab(3).Control(1)=   "Lvw给药途径"
      Tab(3).Control(2)=   "lbl给药途径"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "来源病区(&4)"
      TabPicture(4)   =   "frmPIVAParaSet.frx":D932
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chk来源科室"
      Tab(4).Control(1)=   "lvw来源科室"
      Tab(4).Control(2)=   "lbl来源科室"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "优先级设置(&5)"
      TabPicture(5)   =   "frmPIVAParaSet.frx":D94E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblpritip"
      Tab(5).Control(1)=   "vsfPri"
      Tab(5).Control(2)=   "vsfDept"
      Tab(5).Control(3)=   "cmdAddPri"
      Tab(5).Control(4)=   "cmdDelPri"
      Tab(5).Control(5)=   "cmdIN"
      Tab(5).Control(6)=   "chkAll"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "容量设置(&6)"
      TabPicture(6)   =   "frmPIVAParaSet.frx":D96A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblvoltip"
      Tab(6).Control(1)=   "vsfVolume"
      Tab(6).Control(2)=   "cmdVolDel"
      Tab(6).Control(3)=   "cmdVolAdd"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "常用药品设置(&7)"
      TabPicture(7)   =   "frmPIVAParaSet.frx":D986
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblMedi"
      Tab(7).Control(1)=   "vsfPrint"
      Tab(7).Control(2)=   "chkByMedi"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "不配置药品(&8)"
      TabPicture(8)   =   "frmPIVAParaSet.frx":D9A2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblNoMedi"
      Tab(8).Control(1)=   "vsfNoMedi"
      Tab(8).ControlCount=   2
      Begin TabDlg.SSTab tabPrice 
         Height          =   1935
         Left            =   -68760
         TabIndex        =   109
         Top             =   3600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3413
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "配药类型"
         TabPicture(0)   =   "frmPIVAParaSet.frx":D9BE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "VSFPrice"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "给药途径(只支持静脉营养类型)"
         TabPicture(1)   =   "frmPIVAParaSet.frx":D9DA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "VSFPrice_给药途径"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid VSFPrice 
            Height          =   1245
            Left            =   360
            TabIndex        =   110
            Top             =   480
            Width           =   3960
            _cx             =   6985
            _cy             =   2196
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   10329501
            GridColorFixed  =   10329501
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
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
            FormatString    =   $"frmPIVAParaSet.frx":D9F6
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
            AccessibleDescription=   "200"
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFPrice_给药途径 
            Height          =   1245
            Left            =   -74520
            TabIndex        =   111
            Top             =   480
            Width           =   3600
            _cx             =   6350
            _cy             =   2196
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   10329501
            GridColorFixed  =   10329501
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
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
            FormatString    =   $"frmPIVAParaSet.frx":DAA2
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
            AccessibleDescription=   "200"
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "向下(&N)"
         Height          =   350
         Left            =   -67080
         TabIndex        =   108
         Top             =   5640
         Width           =   1100
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "向上(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -68760
         TabIndex        =   107
         Top             =   5640
         Width           =   1100
      End
      Begin VB.CheckBox chkByMedi 
         Caption         =   "是否根据设置的常用药品进行药品过滤操作"
         Height          =   255
         Left            =   -74880
         TabIndex        =   101
         Top             =   360
         Width           =   3855
      End
      Begin VB.PictureBox picprice 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -68760
         Picture         =   "frmPIVAParaSet.frx":DB4B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   90
         Top             =   3360
         Width           =   240
      End
      Begin VB.Frame fra操作控制 
         Caption         =   "  操作控制 "
         Height          =   1695
         Left            =   120
         TabIndex        =   76
         Top             =   1320
         Width           =   11295
         Begin VB.CheckBox chkPeople 
            Caption         =   "打印瓶签时填写各个环节的实际操作员"
            Height          =   255
            Left            =   7320
            TabIndex        =   106
            Top             =   600
            Width           =   3495
         End
         Begin VB.CheckBox chkPacket 
            Caption         =   "打包药品在发送环节收取配置费"
            Height          =   255
            Left            =   7320
            TabIndex        =   105
            Top             =   300
            Width           =   2895
         End
         Begin VB.CheckBox chkBeach 
            Caption         =   "当天发送的医嘱产生的输液单全部到备用批次"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   3975
         End
         Begin VB.CheckBox chkOutPai 
            Caption         =   "出院病人不收配置费"
            Height          =   255
            Left            =   4320
            TabIndex        =   99
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkMedi 
            Caption         =   "特殊药品按药品类型指定批次"
            Height          =   255
            Left            =   4320
            TabIndex        =   93
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox chkSort 
            Caption         =   "输液单按批次，药品规则排序"
            Height          =   255
            Left            =   4320
            TabIndex        =   92
            Top             =   840
            Width           =   2775
         End
         Begin VB.CheckBox chkMoney 
            Caption         =   "配置费按病人收取"
            Height          =   255
            Left            =   4320
            TabIndex        =   89
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chksend 
            Caption         =   "条码扫描一次自动发送"
            Height          =   255
            Left            =   4320
            TabIndex        =   88
            Top             =   300
            Width           =   2895
         End
         Begin VB.CheckBox chkPackage 
            Caption         =   "配液输液单配药后允许销帐申请"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   840
            Width           =   3975
         End
         Begin VB.CheckBox chk打包设置 
            Caption         =   "允许调整打包状态（排药印签、配药环节）"
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   560
            Width           =   3855
         End
         Begin VB.CheckBox chk批次设置 
            Caption         =   "允许手工调整批次（排药印签环节）"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   300
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "  配置中心库房选择 "
         Height          =   615
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   11295
         Begin VB.CheckBox chkCheck 
            Caption         =   "审核该药房的所有医嘱"
            Height          =   255
            Left            =   4680
            TabIndex        =   98
            Top             =   240
            Width           =   3855
         End
         Begin VB.ComboBox CboStore 
            ForeColor       =   &H80000012&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   240
            Width           =   2280
         End
         Begin VB.Label lblStore 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "配置中心"
            Height          =   180
            Left            =   360
            TabIndex        =   73
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " 卡片控制 "
         Height          =   735
         Left            =   -68760
         TabIndex        =   67
         Top             =   1800
         Width           =   4575
         Begin VB.ComboBox cbo数量 
            Height          =   300
            ItemData        =   "frmPIVAParaSet.frx":1439D
            Left            =   1200
            List            =   "frmPIVAParaSet.frx":143AA
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl数量1 
            Caption         =   "单行显示"
            Height          =   195
            Left            =   360
            TabIndex        =   70
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lbl数量2 
            Caption         =   "张卡片"
            Height          =   195
            Left            =   2160
            TabIndex        =   69
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "应用于所有科室的优先级规则"
         Height          =   250
         Left            =   -74880
         TabIndex        =   55
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CommandButton cmdIN 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   54
         Top             =   1800
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   51
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdVolDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   50
         Top             =   1560
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelPri 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   49
         Top             =   2400
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPri 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -64680
         TabIndex        =   48
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CheckBox chk来源科室 
         Caption         =   "启用来源病区控制"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox chk给药途径 
         Caption         =   "启用输液给药途径控制"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame fra医嘱类型 
         Caption         =   "  医嘱类型选择  "
         Height          =   615
         Left            =   -68760
         TabIndex        =   22
         Top             =   2640
         Width           =   4575
         Begin VB.CheckBox chk医嘱类型 
            Caption         =   "临嘱"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chk医嘱类型 
            Caption         =   "长嘱"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fra输液量 
         Caption         =   "  配药控制 "
         Height          =   4215
         Left            =   -74880
         TabIndex        =   20
         Top             =   1800
         Width           =   5895
         Begin VB.CheckBox chkAutoMode 
            Caption         =   "自动排批时输液单的批次只往后面批次变动"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   3840
            Width           =   5535
         End
         Begin VB.CheckBox chkChangeDrug 
            Caption         =   "不允许置换药房到输液配置中心"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   3260
            Width           =   5535
         End
         Begin VB.CheckBox chkAutoBatch 
            Caption         =   "启用自动排批（启用自动排批后，将不再保持上次批次）"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2680
            Width           =   5535
         End
         Begin VB.CheckBox chkLastBatch 
            Caption         =   "保持上次批次"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkTpn 
            Caption         =   "配置中心不接收的静脉营养医嘱在病区配置"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   820
            Width           =   4335
         End
         Begin VB.CheckBox chkSpecial 
            Caption         =   "自备药、不取药、离院带药允许发送到配置中心"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1400
            Width           =   4215
         End
         Begin VB.CheckBox chkBag 
            Caption         =   "单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包"
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   1980
            Width           =   5655
         End
      End
      Begin VB.Frame fra输液单控制 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   10695
         Begin MSComCtl2.UpDown updDeff 
            Height          =   270
            Left            =   3600
            TabIndex        =   65
            Top             =   795
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDeff 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3240
            TabIndex        =   64
            Text            =   "0"
            Top             =   795
            Width           =   375
         End
         Begin VB.CheckBox chkOpen 
            Caption         =   "启用接收时间段控制"
            Height          =   180
            Left            =   360
            TabIndex        =   59
            Top             =   0
            Width           =   1935
         End
         Begin VB.PictureBox Picture5 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4680
            Picture         =   "frmPIVAParaSet.frx":143B7
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
            Top             =   795
            Width           =   240
         End
         Begin VB.CheckBox chk当日医嘱 
            Caption         =   "接收当日及以前的医嘱"
            Height          =   180
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   2175
         End
         Begin VB.PictureBox picHelpIcon 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4680
            Picture         =   "frmPIVAParaSet.frx":1AC09
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   25
            Top             =   240
            Width           =   240
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   315
            Left            =   960
            TabIndex        =   61
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   80871426
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   315
            Left            =   3240
            TabIndex        =   63
            Top             =   330
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
            Format          =   80871426
            CurrentDate     =   36985
         End
         Begin VB.Label lblDeff 
            Caption         =   "小时差"
            Height          =   255
            Left            =   2595
            TabIndex        =   66
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblEnd 
            Caption         =   "结束时间"
            Height          =   255
            Left            =   2400
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblBegin 
            Caption         =   "开始时间"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl当日医嘱 
            AutoSize        =   -1  'True
            Caption         =   "勾选时配置中心将接收满足时间差条件的当日执行的医嘱。"
            Height          =   180
            Left            =   5040
            TabIndex        =   39
            Top             =   795
            Width           =   4680
         End
         Begin VB.Label lbl时间控制 
            AutoSize        =   -1  'True
            Caption         =   "医嘱发送不在该时间段输液医嘱将不再产生输液单。"
            Height          =   180
            Left            =   5040
            TabIndex        =   21
            Top             =   240
            Width           =   4140
         End
      End
      Begin VB.Frame fra打印控制 
         Caption         =   "  打印控制 "
         Height          =   2840
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   11295
         Begin VB.ComboBox cboSum 
            Height          =   300
            Left            =   6120
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   885
            Width           =   2415
         End
         Begin VB.ComboBox cboNum 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   2460
            Width           =   2415
         End
         Begin VB.ComboBox cbo标签打印 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   540
            Width           =   2415
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "配药后"
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   34
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkPrintLabelStep 
            Caption         =   "摆药后"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cbo标签打印 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   180
            Width           =   2415
         End
         Begin VB.CommandButton cmd打印设置 
            Caption         =   "打印设置(&P)"
            Height          =   345
            Left            =   3960
            TabIndex        =   18
            Top             =   2055
            Width           =   1155
         End
         Begin VB.ComboBox cbo票据设置 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2085
            Width           =   2415
         End
         Begin VB.CheckBox chkManPrint 
            Caption         =   "允许手工控制打印瓶签（可进行补打）"
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   915
            Width           =   3375
         End
         Begin VB.ComboBox cbo发送单 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1695
            Width           =   2415
         End
         Begin VB.ComboBox cbo摆药单 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "仅限于摆药或配药后打印"
            Height          =   180
            Left            =   3960
            TabIndex        =   87
            Top             =   2520
            Width           =   1980
         End
         Begin VB.Label lblSumPrint 
            AutoSize        =   -1  'True
            Caption         =   "打印标签后"
            Height          =   180
            Left            =   5040
            TabIndex        =   86
            Top             =   952
            Width           =   900
         End
         Begin VB.Label lblSum 
            AutoSize        =   -1  'True
            Caption         =   "汇总报表"
            Height          =   180
            Left            =   8640
            TabIndex        =   85
            Top             =   952
            Width           =   720
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "瓶签打印份数"
            Height          =   180
            Left            =   180
            TabIndex        =   74
            Top             =   2520
            Width           =   1080
         End
         Begin VB.Label lblPrintLabel 
            AutoSize        =   -1  'True
            Caption         =   "瓶签打印方式"
            Height          =   180
            Left            =   180
            TabIndex        =   36
            Top             =   405
            Width           =   1080
         End
         Begin VB.Label lbl票据 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "票据和报表"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   17
            Top             =   2145
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "瓶签手工打印"
            Height          =   180
            Left            =   180
            TabIndex        =   15
            Top             =   952
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "汇总发送清单"
            Height          =   180
            Left            =   3960
            TabIndex        =   13
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "摆药汇总清单"
            Height          =   180
            Left            =   3960
            TabIndex        =   12
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "发送确认后"
            Height          =   180
            Left            =   360
            TabIndex        =   10
            Top             =   1755
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "摆药确认后"
            Height          =   180
            Left            =   360
            TabIndex        =   8
            Top             =   1380
            Width           =   900
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新增(&A)"
         Height          =   350
         Left            =   -64800
         TabIndex        =   5
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   -64800
         TabIndex        =   4
         Top             =   1560
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   9960
         _cx             =   17568
         _cy             =   8864
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2145B
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
      Begin MSComctlLib.ListView Lvw给药途径 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8387
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvw来源科室 
         Height          =   4755
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   8387
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   4785
         Left            =   -74880
         TabIndex        =   52
         Top             =   1080
         Width           =   2400
         _cx             =   4233
         _cy             =   8440
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":215EF
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPri 
         Height          =   4785
         Left            =   -72360
         TabIndex        =   53
         Top             =   1080
         Width           =   7560
         _cx             =   13335
         _cy             =   8440
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":21685
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
      Begin VSFlex8Ctl.VSFlexGrid vsfVolume 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   58
         Top             =   840
         Width           =   10080
         _cx             =   17780
         _cy             =   8864
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2173E
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
      Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
         Height          =   5025
         Left            =   -74880
         TabIndex        =   80
         Top             =   960
         Width           =   10080
         _cx             =   17780
         _cy             =   8864
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":217E2
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoMedi 
         Height          =   5145
         Left            =   -74880
         TabIndex        =   82
         Top             =   840
         Width           =   10080
         _cx             =   17780
         _cy             =   9075
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAParaSet.frx":2184B
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
      Begin VB.Label lblprice 
         AutoSize        =   -1  'True
         Caption         =   "配置收费设置"
         Height          =   180
         Left            =   -68400
         TabIndex        =   91
         Top             =   3360
         Width           =   1080
      End
      Begin VB.Label lblNoMedi 
         Caption         =   "设置配置中心不进行配置的药品"
         Height          =   255
         Left            =   -74880
         TabIndex        =   83
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblMedi 
         Caption         =   "设置常用药品，在输液单界面可以按药品进行过滤和排序"
         Height          =   255
         Left            =   -74880
         TabIndex        =   81
         Top             =   720
         Width           =   7935
      End
      Begin VB.Label lblvoltip 
         AutoSize        =   -1  'True
         Caption         =   "设置某个科室单个病人某个批次可以配药的容量"
         Height          =   180
         Left            =   -74880
         TabIndex        =   57
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lblpritip 
         AutoSize        =   -1  'True
         Caption         =   "可以设置同个批次中同组药品的优先级"
         Height          =   180
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label lbl来源科室 
         AutoSize        =   -1  'True
         Caption         =   "启用时可选择病区。输液医嘱发送时如果病人的所在病区没有选择，则不会产生输液单据。"
         Height          =   180
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   7200
      End
      Begin VB.Label lbl给药途径 
         AutoSize        =   -1  'True
         Caption         =   "启用时可选择下列输液类的给药途径。输液医嘱发送时如果医嘱的给药途径没有选择，则不会产生输液单据。"
         Height          =   180
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   8640
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "设置配置中心工作批次(0批次作为特殊批次存在，不按给药时间范围划定)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   5850
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10320
      TabIndex        =   1
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9000
      TabIndex        =   0
      Top             =   6360
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvwSel 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":218B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":21BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":21EE8
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAParaSet.frx":2223A
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   5880
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPIVAParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String                              '权限串
Public mlng库房id As Long
Private mblnSetPara As Boolean
Private mRsDept As Recordset
Private mRsPC As Recordset
Private mRsWay As Recordset
Private mRsType As Recordset
Private mRsPrice As Recordset
Private mintRow As Integer
Private mintCol As Integer
Private mintPri As Integer
Private mblnPrice As Boolean
Private mblnEdit As Boolean     '是否编辑优先级
Private mrs收费项目 As Recordset
Private Sub LoadStore()
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.id,B.名称 From 部门性质说明 A,部门表 B" & _
    " Where A.部门ID=B.ID And A.工作性质='配制中心' And B.Id In (Select 部门id From 部门人员 Where 人员id = [1])"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取配置中心的部门", glngUserId)
    
    With Me.CboStore
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadParams()
    Dim int摆药单 As Integer
    Dim int发药单 As Integer
    Dim int批次设置 As Integer
    Dim int上次批次 As Integer
    Dim int打包设置 As Integer
    Dim strAutoPrint As String
    Dim intManPrint As Integer
    Dim str截止时间 As String
    Dim int医嘱类型 As Integer
    Dim dbl输液量 As Double
    Dim str大输液药品剂型 As String
    Dim str输液给药途径 As String
    Dim str来源科室 As String
    Dim rsData As ADODB.Recordset
    Dim str当日医嘱 As String
    Dim IntCount As Integer
    Dim intOpen As Integer
    Dim lng部门ID As Long
    Dim IntLocate As Integer
    Dim dateNow As Date
    Dim intNum As Integer
    Dim int配药后打包 As Integer
    Dim i As Integer
    Dim int汇总 As Integer
    Dim intTPN As Integer
    Dim intSpecial As Integer
    
    On Error GoTo errHandle
    '基础
    int摆药单 = Val(zlDatabase.GetPara("摆药后打印", glngSys, 1345, 0, Array(Label3, cbo摆药单, Label5), mblnSetPara))
    int发药单 = Val(zlDatabase.GetPara("发送后打印", glngSys, 1345, 0, Array(Label4, cbo发送单, Label6), mblnSetPara))
    int批次设置 = Val(zlDatabase.GetPara("批次设置", glngSys, 1345, 0, Array(chk批次设置), mblnSetPara))
    int打包设置 = Val(zlDatabase.GetPara("打包设置", glngSys, 1345, 0, Array(chk打包设置), mblnSetPara))
    strAutoPrint = zlDatabase.GetPara("瓶签自动打印", glngSys, 1345, "00|00", Array(lblPrintLabel, chkPrintLabelStep(0), chkPrintLabelStep(1)), mblnSetPara)
    intManPrint = Val(zlDatabase.GetPara("瓶签手工打印", glngSys, 1345, "0", Array(Label8, chkManPrint), mblnSetPara))
    IntCount = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\输液卡片", "卡片数量", 3))
    int配药后打包 = Val(zlDatabase.GetPara("配液输液单配药后允许销帐申请", glngSys, 1345, 0, Array(chkPackage), mblnSetPara))
    int汇总 = Val(zlDatabase.GetPara("打印标签后是否打印汇总报表", glngSys, 1345, 0, Array(lblSumPrint, cboSum, lblSum), mblnSetPara))
    
    '辅助控制
    str截止时间 = zlDatabase.GetPara("工作截止时间", glngSys, 1345, "", Array(lblBegin, dtpBegin, lblEnd, dtpEnd), mblnSetPara)
    str当日医嘱 = zlDatabase.GetPara("不接收当日及以前医嘱", glngSys, 1345, 0, Array(chk当日医嘱, txtDeff, updDeff, lblDeff), mblnSetPara)
    int医嘱类型 = Val(zlDatabase.GetPara("医嘱类型", glngSys, 1345, 1, Array(chk医嘱类型(0), chk医嘱类型(1)), mblnSetPara))
'    dbl输液量 = Val(zlDatabase.GetPara("同批次输液总量", glngSys, 1345, "", Array(chk输液量控制, txt输液总量, lbl总量控制说明), mblnSetPara))
'    str大输液药品剂型 = zldatabase.GetPara("大输液药品剂型", glngSys, 1345, "", Array(chk药品剂型, txt药品剂型, cmd药品剂型), mblnSetPara)
    int上次批次 = Val(zlDatabase.GetPara("保持上次批次", glngSys, 1345, 0, Array(chkLastBatch), mblnSetPara))
    intOpen = Val(zlDatabase.GetPara("启用接收时间控制", glngSys, 1345, 0, Array(chkOpen), mblnSetPara))
    lng部门ID = Val(zlDatabase.GetPara("配置中心", glngSys, 1345, 0, Array(CboStore, lblStore), mblnSetPara))
    intNum = Val(zlDatabase.GetPara("瓶签打印份数", glngSys, 1345, 1, Array(lblNum, cboNum), mblnSetPara))
    intTPN = Val(zlDatabase.GetPara("配置中心不接收的静脉营养医嘱在病区配置", glngSys, 1345, 0, Array(chkTpn), mblnSetPara))
    intSpecial = Val(zlDatabase.GetPara("特殊性质药品允许发送到配制中心", glngSys, 1345, 0, Array(chkSpecial), mblnSetPara))
    Me.chksend.Value = Val(zlDatabase.GetPara("扫两次瓶签号自动发送", glngSys, 1345, 0, Array(chksend), mblnSetPara))
    Me.chkMoney.Value = Val(zlDatabase.GetPara("配置费按病人收取", glngSys, 1345, 0, Array(chkMoney), mblnSetPara))
    Me.chkBag.Value = Val(zlDatabase.GetPara("单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包", glngSys, 1345, 0, Array(chkBag), mblnSetPara))
    Me.chkSort.Value = Val(zlDatabase.GetPara("按批次，药品排序", glngSys, 1345, 0, Array(chkSort), mblnSetPara))
    Me.chkMedi.Value = Val(zlDatabase.GetPara("特殊药品按药品类型指定批次", glngSys, 1345, 0, Array(chkMedi), mblnSetPara))
    Me.chkCheck.Value = Val(zlDatabase.GetPara("审核该药房的所有数据", glngSys, 1345, 0, Array(chkCheck), mblnSetPara))
    Me.chkOutPai.Value = Val(zlDatabase.GetPara("出院病人不收配置费", glngSys, 1345, 0, Array(chkOutPai), mblnSetPara))
    Me.chkAutoBatch.Value = Val(zlDatabase.GetPara("启动自动排批", glngSys, 1345, 0, Array(chkAutoBatch), mblnSetPara))
    Me.chkByMedi.Value = Val(zlDatabase.GetPara("是否按设置的常用药品进行药品过滤操作", glngSys, 1345, 0, Array(chkByMedi), mblnSetPara))
    Me.chkChangeDrug.Value = Val(zlDatabase.GetPara("不允许置换药房到输液配置中心", glngSys, 1345, 0, Array(chkChangeDrug), mblnSetPara))
    Me.chkAutoMode.Value = Val(zlDatabase.GetPara("自动排批时输液单的批次只往后面批次变动", glngSys, 1345, 0, Array(chkAutoMode), mblnSetPara))
    Me.chkBeach.Value = Val(zlDatabase.GetPara("当天发送的医嘱产生的输液单全部到备用批次", glngSys, 1345, 0, Array(chkBeach), mblnSetPara))
    Me.chkPeople.Value = Val(zlDatabase.GetPara("打印瓶签时填写各个环节的实际操作员", glngSys, 1345, 0, Array(chkPeople), mblnSetPara))
    Me.chkPacket.Value = Val(zlDatabase.GetPara("打包药品在发送环节收取配置费", glngSys, 1345, 0, Array(chkPacket), mblnSetPara))
    
    '给药途径
    str输液给药途径 = zlDatabase.GetPara("输液给药途径", glngSys, 1345, "", Array(chk给药途径, Lvw给药途径), mblnSetPara)
    
    '来源科室
    str来源科室 = zlDatabase.GetPara("来源病区", glngSys, 1345, "", Array(chk来源科室, lvw来源科室), mblnSetPara)
    
    If lng部门ID <> 0 Then                                  '定位药房
        '不存在该药房则提示
        For IntLocate = 0 To Me.CboStore.ListCount - 1
            If Me.CboStore.ItemData(IntLocate) = lng部门ID Then
                Me.CboStore.ListIndex = IntLocate
                Exit For
            End If
        Next
        If IntLocate > (CboStore.ListCount - 1) Then
            MsgBox "请重新设置配置中心（原来设置的配置中心已失效）！", vbInformation, gstrSysName
            If CboStore.ListCount >= 1 Then CboStore.ListIndex = 0
        End If
    Else
        MsgBox "请设置配置中心！", vbInformation, gstrSysName
    End If
    
    Me.chkOpen.Value = intOpen
    
    If InStr(1, str截止时间, "|") > 0 Then
        Me.dtpBegin.Value = Mid(str截止时间, 1, InStr(1, str截止时间, "|") - 1)
        Me.dtpEnd.Value = Mid(str截止时间, InStr(1, str截止时间, "|") + 1)
    End If
    
    Me.chk当日医嘱.Value = Mid(str当日医嘱, 1, 1)
    If InStr(1, str当日医嘱, "|") > 1 Then
        Me.txtDeff.Text = Mid(str当日医嘱, 3)
    Else
        Me.txtDeff.Text = 0
    End If
    
    ''基础设置
    If int摆药单 >= 0 And int摆药单 <= cbo摆药单.ListCount - 1 Then
        cbo摆药单.ListIndex = int摆药单
    End If
    
    If int汇总 >= 0 And int汇总 <= cboSum.ListCount - 1 Then
        cboSum.ListIndex = int汇总
    End If
    
    If int发药单 >= 0 And int发药单 <= cbo摆药单.ListCount - 1 Then
        cbo发送单.ListIndex = int发药单
    End If
    
    If int批次设置 >= 0 And int批次设置 <= 1 Then
        chk批次设置.Value = int批次设置
    End If
        
    If int打包设置 >= 0 And int打包设置 <= 1 Then
        chk打包设置.Value = int打包设置
    End If
    
    If int配药后打包 >= 0 And int配药后打包 <= 1 Then
        chkPackage.Value = int配药后打包
    End If
    
    If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
        strAutoPrint = "00|00"
    End If
    
    If Mid(strAutoPrint, 1, 1) = 1 Then
        chkPrintLabelStep(0).Value = 1
        If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
            cbo标签打印(0).ListIndex = 1
        Else
            cbo标签打印(0).ListIndex = 0
        End If
    End If
    
    If Mid(strAutoPrint, 4, 1) = 1 Then
        chkPrintLabelStep(1).Value = 1
        If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
            cbo标签打印(1).ListIndex = 1
        Else
            cbo标签打印(1).ListIndex = 0
        End If
    End If
    
    cbo标签打印(0).Enabled = chkPrintLabelStep(0).Enabled And (chkPrintLabelStep(0).Value = 1)
    cbo标签打印(1).Enabled = chkPrintLabelStep(1).Enabled And (chkPrintLabelStep(1).Value = 1)
    
    vsfVolume.Enabled = mblnSetPara
    vsfPrint.Enabled = mblnSetPara
    vsfNoMedi.Enabled = mblnSetPara
    vsfPri.Enabled = mblnSetPara
    cmdAddPri.Enabled = mblnSetPara
    cmdIN.Enabled = mblnSetPara
    cmdDelPri.Enabled = mblnSetPara
    cmdVolAdd.Enabled = mblnSetPara
    cmdVolDel.Enabled = mblnSetPara
    
    If intManPrint < 0 Or intManPrint > 1 Then
        chkManPrint.Value = 0
    Else
        chkManPrint.Value = intManPrint
    End If
    
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
    
    '卡片张数
    Me.cbo数量.Text = IIf(IntCount = 0, 3, IntCount)
    
    Me.cboNum.Text = IIf(intNum = 0, 3, intNum)
    
    ''辅助控制
'    chk截止时间.Value = IIf(str截止时间 = "", 0, 1)
'    dtpTime.Enabled = (chk截止时间.Value = 1)
'    If str截止时间 <> "" Then
'        If IsDate(str截止时间) = True Then
'            dtpTime.Value = str截止时间
'        End If
'    End If
    
'    txt输液总量.Text = ""
'    chk输液量控制.Value = IIf(dbl输液量 = 0, 0, 1)
'    txt输液总量.Enabled = (chk输液量控制.Value = 1)
    If dbl输液量 > 0 Then
'        txt输液总量.Text = dbl输液量
    End If
    
    If int医嘱类型 = 0 Then
        chk医嘱类型(0).Value = 1
        chk医嘱类型(1).Value = 1
    ElseIf int医嘱类型 = 1 Then
        chk医嘱类型(0).Value = 1
        chk医嘱类型(1).Value = 0
    ElseIf int医嘱类型 = 2 Then
        chk医嘱类型(0).Value = 0
        chk医嘱类型(1).Value = 1
    End If
    
'    chk药品剂型.Value = IIf(str大输液药品剂型 = "", 0, 1)
'    txt药品剂型.Text = str大输液药品剂型
'    txt药品剂型.Enabled = (chk药品剂型.Value = 1)
'    cmd药品剂型.Enabled = (chk药品剂型.Value = 1)
    
    chkLastBatch.Value = IIf(int上次批次 = 0, 0, 1)
    
    '静脉营养药物处置方式
    If intTPN >= 0 And intTPN <= 1 Then
        chkTpn.Value = intTPN
    End If
    
    '特殊药品处理
    If intSpecial >= 0 And intSpecial <= 1 Then
        chkSpecial.Value = intSpecial
    End If
    
    ''给药途径
    gstrSQL = "Select ID, 名称 as 用法 ,标本部位 As 分类 From 诊疗项目目录 Where 类别='E' And 操作类型='2'And (服务对象=2 Or 服务对象=3) And 执行分类 = 1 " & _
            " And (撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or 撤档时间 Is Null) Order by 编码 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取给药途径")
    
    With Lvw给药途径
        .ListItems.Clear
        Do While Not rsData.EOF
            .ListItems.Add , "_" & rsData!Id, rsData!用法, 1, 1
            If InStr(1, "," & str输液给药途径 & ",", "," & rsData!Id & ",") > 0 Then
                .ListItems(.ListItems.count).Checked = True
            End If
            rsData.MoveNext
        Loop
    End With
    
    If str输液给药途径 <> "" Then
        chk给药途径.Value = 1
    End If
    
    Lvw给药途径.Enabled = chk给药途径.Enabled And (chk给药途径.Value = 1)
    Lvw给药途径.BackColor = IIf(Lvw给药途径.Enabled, &H80000005, &H8000000F)
    
    ''来源科室
    gstrSQL = "Select 编码 || '-' || 名称 科室, Id " & _
            " From 部门表 " & _
            " Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And Id In (Select 部门id From 部门性质说明 Where 工作性质 = '护理' And 服务对象 In (2,3)) And " & _
            " (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By 编码 || '-' || 名称 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load来源科室")

    With rsData
        lvw来源科室.ListItems.Clear
        Do While Not .EOF
            lvw来源科室.ListItems.Add , "_" & !Id, !科室, 1, 1
            If str来源科室 <> "" Then
                If InStr("," & str来源科室 & ",", "," & CStr(!Id) & ",") > 0 Then
                    lvw来源科室.ListItems("_" & !Id).Checked = True
                End If
            End If
            .MoveNext
        Loop
    End With
    
    If str来源科室 <> "" Then
        chk来源科室.Value = 1
    End If
    lvw来源科室.Enabled = chk来源科室.Enabled And (chk来源科室.Value = 1)
    lvw来源科室.BackColor = IIf(lvw来源科室.Enabled, &H80000005, &H8000000F)
    
    '常用药品打印设置
    gstrSQL = "select 药品id,名称 from 输液优先打印药品"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load药品")
    
    Me.vsfPrint.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("药品id")) = rsData!药品ID
        Me.vsfPrint.TextMatrix(i, vsfPrint.ColIndex("药品名称与编码")) = rsData!名称
       
       rsData.MoveNext
    Next
    
    
    '输液不配置药品
    gstrSQL = "select 药品id,名称 from 输液不配置药品"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Load药品")
    
    Me.vsfNoMedi.rows = rsData.RecordCount + 2
    For i = 1 To rsData.RecordCount
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("药品id")) = rsData!药品ID
        Me.vsfNoMedi.TextMatrix(i, vsfNoMedi.ColIndex("药品名称与编码")) = rsData!名称
       
       rsData.MoveNext
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub CboStore_Click()
    Call LoadBatchSet
    Call loadVolume
End Sub

Private Sub chkAll_Click()
    If mblnEdit Then
        If MsgBox("请保存设置的优先级，切换科室后所作的优先级设置将失效，是否切换？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
           If Me.chkAll.Value = 0 Then
                Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
                Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
                Me.vsfDept.Visible = True
                Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
            Else
                Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
                Me.vsfPri.Left = Me.vsfDept.Left
                Me.vsfDept.Visible = False
                
                Call LoadVsfPRI(0)
            End If
            mblnEdit = False
            
        End If
    Else
        If Me.chkAll.Value = 0 Then
            Me.vsfPri.Left = Me.vsfDept.Width + Me.vsfDept.Left + 100
            Me.vsfPri.Width = Me.vsfPri.Width - Me.vsfDept.Width - 100
            Me.vsfDept.Visible = True
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
        Else
            Me.vsfPri.Width = Me.vsfPri.Width + Me.vsfDept.Width + 100
            Me.vsfPri.Left = Me.vsfDept.Left
            Me.vsfDept.Visible = False
            
            Call LoadVsfPRI(0)
        End If
    End If
End Sub

Private Sub chkManPrint_Click()
    If chkManPrint.Value = 1 Then
        cboSum.Enabled = True
    Else
        cboSum.Enabled = False
    End If
End Sub

Private Sub chkOpen_Click()
    Me.dtpBegin.Enabled = (Me.chkOpen.Value = 1)
    Me.dtpEnd.Enabled = (Me.chkOpen.Value = 1)
    Me.chk当日医嘱.Enabled = (Me.chkOpen.Value = 1)
    Me.updDeff.Enabled = (Me.chkOpen.Value = 1)
End Sub

Private Sub chkPrintLabelStep_Click(Index As Integer)
    cbo标签打印(Index).Enabled = (chkPrintLabelStep(Index).Value = 1)
End Sub


Private Sub chk给药途径_Click()
    Lvw给药途径.Enabled = (chk给药途径.Value = 1)
    Lvw给药途径.BackColor = IIf(Lvw给药途径.Enabled, &H80000005, &H8000000F)
End Sub

Private Sub chk来源科室_Click()
    lvw来源科室.Enabled = (chk来源科室.Value = 1)
    lvw来源科室.BackColor = IIf(lvw来源科室.Enabled, &H80000005, &H8000000F)
End Sub

'Private Sub chk输液量控制_Click()
'    If chk输液量控制.Value = 1 Then
'        txt输液总量.Enabled = True
'    Else
'        txt输液总量.Enabled = False
'    End If
'End Sub


'Private Sub chk药品剂型_Click()
'    txt药品剂型.Enabled = (chk药品剂型.Value = 1)
'    cmd药品剂型.Enabled = (chk药品剂型.Value = 1)
'
'    If chk药品剂型.Value = 0 And pic剂型.Visible = True Then
'        Call cmd剂型Cancel_Click
'    End If
'End Sub

Private Sub chk医嘱类型_Click(Index As Integer)
    If chk医嘱类型(0).Value = 0 And chk医嘱类型(1).Value = 0 Then
        chk医嘱类型(Index).Value = 1
    End If
End Sub

Private Sub cmdAdd_Click()
    With vsfBatch
        If .rows > 2 Then
            If Trim(.TextMatrix(.rows - 1, .ColIndex("配置时间开始"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("配置时间结束"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("给药时间开始"))) = "" Or _
                Trim(.TextMatrix(.rows - 1, .ColIndex("给药时间结束"))) = "" Then
                Exit Sub
            End If
        End If
        
        .rows = .rows + 1
        
        If .Row >= 2 Then
            .TextMatrix(.rows - 1, .ColIndex("批次")) = Mid(.TextMatrix(.rows - 2, .ColIndex("批次")), 1, Len(.TextMatrix(.rows - 2, .ColIndex("批次"))) - 1) + 1 & "#"
        Else
            .TextMatrix(.rows - 1, .ColIndex("批次")) = "0#"
        End If
        .TextMatrix(.rows - 1, .ColIndex("启用")) = "√"
    End With
End Sub

Private Sub cmdAddPri_Click()
    If Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("配药类型")) <> "" And Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, Me.vsfPri.ColIndex("频次")) <> "" Then
        Me.vsfPri.rows = Me.vsfPri.rows + 1
        Me.vsfPri.RowHeight(Me.vsfPri.rows - 1) = 250
        Me.vsfPri.TextMatrix(Me.vsfPri.rows - 1, vsfPri.ColIndex("序号")) = Me.vsfPri.rows - 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngRow As Long
    
    With vsfBatch
        If .Row > 1 Then
            lngRow = .Row
            If MsgBox("是否删除批次(" & .TextMatrix(.Row, .ColIndex("批次")) & ")？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            .Redraw = flexRDNone
            
            .RemoveItem .Row
            
            '重设批次号
            For lngRow = lngRow To .rows - 1
                .TextMatrix(lngRow, .ColIndex("批次")) = Mid(.TextMatrix(lngRow, .ColIndex("批次")), 1, Len(.TextMatrix(lngRow, .ColIndex("批次"))) - 1) - 1 & "#"
            Next
            
            .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub cmdDelPri_Click()
    Dim i As Integer
    Dim intRow As Integer
    
    If Me.vsfPri.Row = 0 Then Exit Sub
    intRow = Me.vsfPri.Row
    Me.vsfPri.RemoveItem Me.vsfPri.Row
    
    '调整序号
    For i = intRow To Me.vsfPri.rows - 1
        Me.vsfPri.TextMatrix(i, Me.vsfPri.ColIndex("序号")) = i
    Next
    
    mblnEdit = True
End Sub

Private Sub cmdIN_Click()
    Dim IntCount As Integer
    Dim lngRow As Long
    
    If mblnSetPara Then
         '保存优先级设置
        With vsfPri
            IntCount = 1
            
            If .rows = 1 Then
                gstrSQL = "Zl_输液药品优先级_Save("
                '科室id
                gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))) & "'"
                gstrSQL = gstrSQL & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存优先级")
            End If
            
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("配药类型")) <> "" And .TextMatrix(lngRow, .ColIndex("频次")) <> "" Then
                    
                    gstrSQL = "Zl_输液药品优先级_Save("
                    '科室id
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, 0, vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))) & "',"
                    '科室名称
                    gstrSQL = gstrSQL & "'" & IIf(chkAll.Value = 1, "所有科室", vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室名称"))) & "',"
                    '配药类型
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("配药类型")) & "',"
                    '频次
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("频次")) & "',"
                    '有效
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("有效"))) & ","
                    '优先级
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("序号")))
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存优先级")
                    IntCount = IntCount + 1
                End If
            Next
        End With
    End If
    
    mblnEdit = False
End Sub

Private Sub cmdLast_Click()
    Dim intRow As Integer
    Dim str配药类型 As String
    Dim str收费项目 As String
    Dim lng项目id As Long
    
    With VSFPrice
        intRow = .Row
        If intRow < 2 Then Exit Sub
        lng项目id = .TextMatrix(.Row - 1, .ColIndex("项目id"))
        str配药类型 = .TextMatrix(.Row - 1, .ColIndex("配药类型"))
        str收费项目 = .TextMatrix(.Row - 1, .ColIndex("收费项目"))
        .TextMatrix(.Row - 1, .ColIndex("项目id")) = .TextMatrix(.Row, .ColIndex("项目id"))
        .TextMatrix(.Row - 1, .ColIndex("配药类型")) = .TextMatrix(.Row, .ColIndex("配药类型"))
        .TextMatrix(.Row - 1, .ColIndex("收费项目")) = .TextMatrix(.Row, .ColIndex("收费项目"))
        
        
        .TextMatrix(.Row, .ColIndex("项目id")) = lng项目id
        .TextMatrix(.Row, .ColIndex("配药类型")) = str配药类型
        .TextMatrix(.Row, .ColIndex("收费项目")) = str收费项目
        
        .Row = intRow - 1
    End With
End Sub

Private Sub cmdNext_Click()
    Dim intRow As Integer
    Dim str配药类型 As String
    Dim str收费项目 As String
    Dim lng项目id As Long
    
    With VSFPrice
        intRow = .Row
        If intRow = .rows - 1 Then Exit Sub
        lng项目id = .TextMatrix(.Row + 1, .ColIndex("项目id"))
        str配药类型 = .TextMatrix(.Row + 1, .ColIndex("配药类型"))
        str收费项目 = .TextMatrix(.Row + 1, .ColIndex("收费项目"))
        .TextMatrix(.Row + 1, .ColIndex("项目id")) = .TextMatrix(.Row, .ColIndex("项目id"))
        .TextMatrix(.Row + 1, .ColIndex("配药类型")) = .TextMatrix(.Row, .ColIndex("配药类型"))
        .TextMatrix(.Row + 1, .ColIndex("收费项目")) = .TextMatrix(.Row, .ColIndex("收费项目"))
        
        
        .TextMatrix(.Row, .ColIndex("项目id")) = lng项目id
        .TextMatrix(.Row, .ColIndex("配药类型")) = str配药类型
        .TextMatrix(.Row, .ColIndex("收费项目")) = str收费项目
        
        .Row = intRow + 1
    End With
End Sub

Private Sub cmdNo_Click()
    picPRI.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub cmdOk_Click()
    Dim strInput As String
    Dim lngRow As Long
    Dim int医嘱类型 As Integer
    Dim str给药途径 As String
    Dim str来源科室 As String
    Dim strPrintLabel As String
    Dim IntCount As Integer
    Dim i As Integer
    Dim n As Integer
    
    On Error GoTo errHandle
    
    If chk医嘱类型(0).Value = 1 And chk医嘱类型(1).Value = 1 Then
        int医嘱类型 = 0
    ElseIf chk医嘱类型(0).Value = 1 Then
        int医嘱类型 = 1
    ElseIf chk医嘱类型(1).Value = 1 Then
        int医嘱类型 = 2
    End If
    
    '来源科室
    With Me.Lvw给药途径
        For lngRow = 1 To .ListItems.count
            If .ListItems(lngRow).Checked Then
                If str给药途径 = "" Then
                    str给药途径 = Mid(.ListItems(lngRow).Key, 2)
                Else
                    str给药途径 = str给药途径 & "," & Mid(.ListItems(lngRow).Key, 2)
                End If
            End If
        Next
    End With
    
    '来源科室
    With Me.lvw来源科室
        For lngRow = 1 To .ListItems.count
            If .ListItems(lngRow).Checked Then
                
                str来源科室 = str来源科室 & Mid(.ListItems(lngRow).Key, 2) & ","
            End If
        Next
    End With
    
    '瓶签打印方式
    If chkPrintLabelStep(0).Value = 0 Then
        strPrintLabel = "00"
    Else
        strPrintLabel = "1" & cbo标签打印(0).ListIndex
    End If
    strPrintLabel = strPrintLabel & "|"
    If chkPrintLabelStep(1).Value = 0 Then
        strPrintLabel = strPrintLabel & "00"
    Else
        strPrintLabel = strPrintLabel & "1" & cbo标签打印(1).ListIndex
    End If
    
    '保存公共及私有参数
    '基础设置
    zlDatabase.SetPara "摆药后打印", cbo摆药单.ListIndex, glngSys, 1345
    zlDatabase.SetPara "发送后打印", cbo发送单.ListIndex, glngSys, 1345
    zlDatabase.SetPara "批次设置", chk批次设置.Value, glngSys, 1345
    zlDatabase.SetPara "打包设置", chk打包设置.Value, glngSys, 1345
    zlDatabase.SetPara "瓶签自动打印", strPrintLabel, glngSys, 1345
    zlDatabase.SetPara "瓶签手工打印", chkManPrint.Value, glngSys, 1345
    zlDatabase.SetPara "配液输液单配药后允许销帐申请", chkPackage.Value, glngSys, 1345
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\输液卡片", "卡片数量", Me.cbo数量.Text
    zlDatabase.SetPara "打印标签后是否打印汇总报表", cboSum.ListIndex, glngSys, 1345
    zlDatabase.SetPara "扫两次瓶签号自动发送", chksend.Value, glngSys, 1345
    zlDatabase.SetPara "配置费按病人收取", chkMoney.Value, glngSys, 1345
    zlDatabase.SetPara "单个药品，不予配置药品及根据给药时间没有配药批次的输液单默认为0批次并打包", chkBag.Value, glngSys, 1345
    zlDatabase.SetPara "按批次，药品排序", chkSort.Value, glngSys, 1345
    zlDatabase.SetPara "特殊药品按药品类型指定批次", chkMedi.Value, glngSys, 1345
    zlDatabase.SetPara "审核该药房的所有数据", chkCheck.Value, glngSys, 1345
     
    '辅助控制
    zlDatabase.SetPara "工作截止时间", Format(dtpBegin.Value, "hh:mm:ss") & "|" & Format(Me.dtpEnd.Value, "hh:mm:ss"), glngSys, 1345
    zlDatabase.SetPara "不接收当日及以前医嘱", chk当日医嘱.Value & "|" & Me.txtDeff.Text, glngSys, 1345
    zlDatabase.SetPara "医嘱类型", int医嘱类型, glngSys, 1345
'    zlDatabase.SetPara "同批次输液总量", IIf(chk输液量控制.Value = 1, Val(txt输液总量.Text), ""), glngSys, 1345
'    zldatabase.SetPara "大输液药品剂型", IIf(chk药品剂型.Value = 1, txt药品剂型.Text, ""), glngSys, 1345
    zlDatabase.SetPara "保持上次批次", chkLastBatch.Value, glngSys, 1345
    zlDatabase.SetPara "启用接收时间控制", chkOpen.Value, glngSys, 1345
    zlDatabase.SetPara "配置中心", Me.CboStore.ItemData(Me.CboStore.ListIndex), glngSys, 1345
    zlDatabase.SetPara "瓶签打印份数", Me.cboNum.Text, glngSys, 1345
    zlDatabase.SetPara "配置中心不接收的静脉营养医嘱在病区配置", chkTpn.Value, glngSys, 1345
    zlDatabase.SetPara "特殊性质药品允许发送到配制中心", chkSpecial.Value, glngSys, 1345
    zlDatabase.SetPara "出院病人不收配置费", chkOutPai.Value, glngSys, 1345
    zlDatabase.SetPara "启动自动排批", chkAutoBatch.Value, glngSys, 1345
    zlDatabase.SetPara "是否按设置的常用药品进行药品过滤操作", chkByMedi.Value, glngSys, 1345
    zlDatabase.SetPara "不允许置换药房到输液配置中心", chkChangeDrug.Value, glngSys, 1345
    zlDatabase.SetPara "自动排批时输液单的批次只往后面批次变动", chkAutoMode.Value, glngSys, 1345
    zlDatabase.SetPara "当天发送的医嘱产生的输液单全部到备用批次", chkBeach.Value, glngSys, 1345
    zlDatabase.SetPara "打印瓶签时填写各个环节的实际操作员", chkPeople.Value, glngSys, 1345
    zlDatabase.SetPara "打包药品在发送环节收取配置费", chkPacket.Value, glngSys, 1345
    


    '给药途径
    zlDatabase.SetPara "输液给药途径", IIf(chk给药途径.Value = 1, str给药途径, ""), glngSys, 1345
    
    '来源科室
    zlDatabase.SetPara "来源病区", IIf(chk来源科室.Value = 1, str来源科室, ""), glngSys, 1345
    
    If IsHavePrivs(mstrPrivs, "设置工作批次") Then
        With vsfBatch
            For lngRow = 2 To .rows - 1
                If IsDate(.TextMatrix(lngRow, .ColIndex("配置时间开始"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("配置时间结束"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("给药时间开始"))) And _
                    IsDate(.TextMatrix(lngRow, .ColIndex("给药时间结束"))) Then
                    
                    strInput = IIf(strInput = "", "", strInput & "|") & _
                        Mid(.TextMatrix(lngRow, .ColIndex("批次")), 1, Len(.TextMatrix(lngRow, .ColIndex("批次"))) - 1) & "," & _
                        .TextMatrix(lngRow, .ColIndex("配置时间开始")) & "-" & .TextMatrix(lngRow, .ColIndex("配置时间结束")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("给药时间开始")) & "-" & .TextMatrix(lngRow, .ColIndex("给药时间结束")) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("打包")) = "", 0, 1) & "," & _
                        IIf(.TextMatrix(lngRow, .ColIndex("启用")) = "", 0, 1) & "," & _
                        .Cell(flexcpBackColor, lngRow, .ColIndex("颜色")) & "," & _
                        .TextMatrix(lngRow, .ColIndex("药品类型"))
                End If
            Next
        End With
        
        '如果strInput为空表示删除整个工作批次
        gstrSQL = "Zl_配药工作批次_Save("
        '批次信息
        gstrSQL = gstrSQL & "'" & strInput & "',"
        gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存配药工作批次")
    End If
     
    If mblnSetPara Then
        '保存容量设置
        With Me.vsfVolume
            For lngRow = 0 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("科室名称")) <> "" And .TextMatrix(lngRow, .ColIndex("容量")) <> "" Then
                    
                    gstrSQL = "Zl_科室容量设置_Save("
                    '科室id
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("科室id")) & "',"
                    '科室名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("科室名称")) & "',"
                    '批次
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("配药批次")) & "',"
                    '容量
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("容量"))) & ","
                    '优先级
                    gstrSQL = gstrSQL & lngRow & ","
                    '配置中心ID
                    gstrSQL = gstrSQL & Me.CboStore.ItemData(Me.CboStore.ListIndex)
                    gstrSQL = gstrSQL & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存容量")
                End If
            Next
        End With
        
        '保存常用药品
        With Me.vsfPrint
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("药品id")) <> "" And .TextMatrix(i, .ColIndex("药品名称与编码")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_输液优先打印药品_打印设置("
                    '药品id
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                    '药品名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("药品名称与编码")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存常用药品")
                End If
            Next
        End With
        
        '保存不接受药品
        With Me.vsfNoMedi
            For i = 1 To .rows - 1
                If (.TextMatrix(i, .ColIndex("药品id")) <> "" And .TextMatrix(i, .ColIndex("药品名称与编码")) <> "") Or i = 1 Then
                    gstrSQL = "Zl_输液不配置药品_设置("
                    '药品id
                    gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("药品id"))) & ","
                    '药品名称
                    gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("药品名称与编码")) & "',"
                    gstrSQL = gstrSQL & i & ")"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存不接受药品")
                End If
            Next
        End With
    End If
    
    With Me.VSFPrice
        For i = 1 To .rows - 1
            If (.TextMatrix(i, .ColIndex("优先级")) <> "" And .TextMatrix(i, .ColIndex("收费项目")) <> "" And .TextMatrix(i, .ColIndex("项目id")) <> "" And .TextMatrix(i, .ColIndex("配药类型")) <> "") Or i = 1 Then
                gstrSQL = "Zl_配置收费方案_设置("
                '序号
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("优先级"))) & ","
                '配药类型
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("配药类型")) & "',"
                '项目id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("项目id"))) & ","
                '收费项目
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("收费项目")) & "',"
                '诊疗id
                gstrSQL = gstrSQL & "NULL" & ","
                '是否第一次重置
                gstrSQL = gstrSQL & i & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存配置费")
            End If
        Next
    End With
    
    n = i - 1
    
    With Me.VSFPrice_给药途径
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("收费项目")) <> "" And .TextMatrix(i, .ColIndex("项目id")) <> "" And .TextMatrix(i, .ColIndex("给药途径")) <> "" And .TextMatrix(i, .ColIndex("诊疗id")) <> "" Then
                gstrSQL = "Zl_配置收费方案_设置("
                '序号
                gstrSQL = gstrSQL & i + n & ","
                '配药类型
                gstrSQL = gstrSQL & "NULL" & ","
                '项目id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("项目id"))) & ","
                '收费项目
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("收费项目")) & "',"
                '诊疗id
                gstrSQL = gstrSQL & Val(.TextMatrix(i, .ColIndex("诊疗id"))) & ","
                '是否第一次重置
                gstrSQL = gstrSQL & i + n & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, "保存配置费")
            End If
        Next
    End With
    
    frmPIVAMain.mblnParamsRefresh = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPRI(ByVal str科室id As String)
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 科室id,科室名称,配药类型,频次,有效,优先级 from 输液药品优先级 where (科室id=[1] or 科室id='0') order by 优先级"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取优先级数据", str科室id)
    
    i = 1
    rsTemp.Filter = "科室id='" & str科室id & "'"
    If rsTemp.EOF Then rsTemp.Filter = ""
    With Me.vsfPri
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("序号")) = 1
        Else
            .rows = rsTemp.RecordCount + 1
        End If
       
        Do While Not rsTemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("序号")) = rsTemp!优先级
            .TextMatrix(i, .ColIndex("配药类型")) = rsTemp!配药类型
            .TextMatrix(i, .ColIndex("频次")) = rsTemp!频次
            .TextMatrix(i, .ColIndex("有效")) = rsTemp!有效
            i = i + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPrice()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 序号,配药类型,项目id,收费项目 from 配置收费方案 where nvl(诊疗id,0) = 0 order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
            .TextMatrix(1, .ColIndex("优先级")) = 1
        Else
            .rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            If NVL(rsTemp!项目id) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("优先级")) = i
                .TextMatrix(i, .ColIndex("配药类型")) = rsTemp!配药类型
                .TextMatrix(i, .ColIndex("项目id")) = rsTemp!项目id
                .TextMatrix(i, .ColIndex("收费项目")) = rsTemp!收费项目
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadVsfPrice_给药途径()
    Dim rsTemp As Recordset
    Dim rsData As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 诊疗id,项目id,收费项目 from 配置收费方案 where nvl(诊疗id,0) <> 0 order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "LoadVsfPrice")
    
    With Me.VSFPrice_给药途径
        .RowHeight(0) = 250
        
        If rsTemp.RecordCount = 0 Then
            .rows = 1
            .rows = 2
        Else
            .rows = rsTemp.RecordCount + 1
        End If
        
        i = 1
        Do While Not rsTemp.EOF
            '查询诊疗项目名称
            gstrSQL = "select 名称 from 诊疗项目目录 where id = [1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "查询诊疗项目名称", rsTemp!诊疗id)
            
            If NVL(rsTemp!项目id) <> 0 Then
                .RowHeight(i) = 250
                .TextMatrix(i, .ColIndex("诊疗id")) = rsTemp!诊疗id
                .TextMatrix(i, .ColIndex("给药途径")) = rsData!名称
                .TextMatrix(i, .ColIndex("项目id")) = rsTemp!项目id
                .TextMatrix(i, .ColIndex("收费项目")) = rsTemp!收费项目
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdVolAdd_Click()
    If Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("科室名称")) <> "" And Me.vsfVolume.TextMatrix(Me.vsfVolume.rows - 1, Me.vsfVolume.ColIndex("容量")) <> "" Then
        Me.vsfVolume.rows = Me.vsfVolume.rows + 1
        Me.vsfVolume.RowHeight(Me.vsfVolume.rows - 1) = 250
    End If
End Sub

Private Sub cmdVolDel_Click()
    If Me.vsfVolume.Row = 0 Then Exit Sub
    Me.vsfVolume.RemoveItem Me.vsfVolume.Row
End Sub

Private Sub cmdYes_Click()
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(1, strIDS)
    
    If mintPri = 1 Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    ElseIf mintPri = 2 Then
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    ElseIf mintPri = 3 Then
        With Me.VSFPrice
            If mintCol = .ColIndex("配药类型") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "该配药类型已经添加，请重新选择！", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
        End With
    ElseIf mintPri = 4 Then
        Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice.ColIndex("收费项目") Then
            VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("项目id")) = strIDS
        End If
    ElseIf mintPri = 5 Then
        With Me.VSFPrice_给药途径
            If mintCol = .ColIndex("给药途径") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "该给药途径已经添加，请重新选择！", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
            .TextMatrix(mintRow, .ColIndex("诊疗id")) = strIDS
        End With
    ElseIf mintPri = 6 Then
        Me.VSFPrice_给药途径.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice_给药途径.ColIndex("收费项目") Then
            VSFPrice_给药途径.TextMatrix(mintRow, VSFPrice_给药途径.ColIndex("项目id")) = strIDS
        End If
    End If
    
End Sub

Private Sub cmd打印设置_Click()
    Dim strBill As String
    Select Case cbo票据设置.ListIndex
    Case 0
        '输液瓶标签
        strBill = "ZL1_BILL_1345_1"
    Case 1
        '摆药药品汇总清单
        strBill = "ZL1_INSIDE_1345_1"
    Case 2
        '发送药品汇总清单
        strBill = "ZL1_INSIDE_1345_2"
    Case 3
        '退药销帐清单
        strBill = "ZL1_BILL_1345_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

'Private Sub cmd剂型Cancel_Click()
'    pic剂型.Visible = False
'    CmdOK.Enabled = True
'    CmdCancel.Enabled = True
'End Sub
'
'Private Sub cmd剂型ok_Click()
'    ReturnSelected剂型 0
'End Sub
'
'Private Sub cmd药品剂型_Click()
'    Dim lngRow As Long
'    Dim str剂型 As String
'
'    On Error Resume Next
'
'    With pic剂型
'        .Visible = True
'
'        .Height = fra药品剂型.Top - 50
'        .Top = sstMain.Top + fra药品剂型.Top + txt药品剂型.Top - .Height - 50
'        .Left = sstMain.Left + fra药品剂型.Left
'        .Width = fra药品剂型.Width - 15
'
'        txt药品剂型.Text = Trim(txt药品剂型.Text)
'        If txt药品剂型.Text <> "" Then
'            With Me.Lvw药品剂型
'                For lngRow = 1 To .ListItems.count
'                    .ListItems(lngRow).Checked = False
'                    str剂型 = Mid(.ListItems(lngRow).Text, InStr(1, .ListItems(lngRow).Text, "-") + 1)
'                    If InStr(1, "," & txt药品剂型.Text & ",", "," & str剂型 & ",") > 0 Then
'                        .ListItems(lngRow).Checked = True
'                    End If
'                Next
'            End With
'        End If
'
'        .SetFocus
'        .ZOrder 0
'
'        CmdOK.Enabled = False
'        CmdCancel.Enabled = False
'    End With
'End Sub

'Private Sub Command1_Click()
'    ReturnSelected剂型 0
'End Sub

Private Sub Command2_Click()
    pic剂型.Visible = False
    CmdOK.Enabled = True
    CmdCancel.Enabled = True
End Sub

Private Sub Form_Load()
    mblnSetPara = IsHavePrivs(mstrPrivs, "参数设置")
    
    With cbo标签打印(0)
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .ListIndex = 0
    End With
    
    With cbo标签打印(1)
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .ListIndex = 0
    End With
    
    With cbo摆药单
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cbo发送单
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cboSum
        .Clear
        .AddItem "0-提示是否打印"
        .AddItem "1-自动打印"
        .AddItem "2-不打印"
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-输液瓶标签"
        .AddItem "2-摆药药品汇总清单"
        .AddItem "3-发送药品汇总清单"
        .AddItem "4-退药销帐清单"

        .ListIndex = 0
    End With
    
    With VSFPrice
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
    With VSFPrice_给药途径
        .Left = 0
        .Top = tabPrice.TabHeight
        .Width = tabPrice.Width
        .Height = tabPrice.Height - tabPrice.TabHeight
    End With
    
'    With cboTPN
'        .Clear
'        .AddItem "0-配置中心不接收，通过部门发药业务处置"
'        .AddItem "1-配置中心始终接收并打包"
'        .AddItem "2-配置中心始终接收并配置"
'
'        .ListIndex = 0
'    End With
    
    With cboNum
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        
        .ListIndex = 0
    End With
        
    With vsfBatch
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("批次")) = True
        .MergeCol(.ColIndex("颜色")) = True
        .MergeCol(.ColIndex("配置时间开始")) = True
        .MergeCol(.ColIndex("配置时间结束")) = True
        .MergeCol(.ColIndex("给药时间开始")) = True
        .MergeCol(.ColIndex("给药时间结束")) = True
        .MergeCol(.ColIndex("打包")) = True
        .MergeCol(.ColIndex("启用")) = True
        .MergeCol(.ColIndex("药品类型")) = True
        .MergeCells = flexMergeFixedOnly
    End With
    
    Call LoadStore
    Call Load药品剂型
    
    '提取参数
    Call LoadBatchSet
    Call LoadParams
    Call LoadPRI
    Call LoadVsfPrice
    Call LoadVsfPrice_给药途径
    
    Call loadVolume
    Call LoadDept
    
    Call chkAll_Click
    
    Call chkOpen_Click
End Sub
Private Sub LoadBatchSet()
    '提取配药中心工作批次
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 批次,颜色, 配药时间, 给药时间, 打包, 启用,药品类型 From 配药工作批次 where 配置中心ID=[1] Order By 批次"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取配药中心工作批次", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    With vsfBatch
        .rows = 2
        .ColComboList(.ColIndex("药品类型")) = "肿瘤药|营养药|抗生素"
        Do While Not rsTmp.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("批次")) = rsTmp!批次 & "#"
            .TextMatrix(.rows - 1, .ColIndex("配置时间开始")) = Mid(rsTmp!配药时间, 1, InStr(rsTmp!配药时间, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("配置时间结束")) = Mid(rsTmp!配药时间, InStr(rsTmp!配药时间, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("给药时间开始")) = Mid(rsTmp!给药时间, 1, InStr(rsTmp!给药时间, "-") - 1)
            .TextMatrix(.rows - 1, .ColIndex("给药时间结束")) = Mid(rsTmp!给药时间, InStr(rsTmp!给药时间, "-") + 1)
            .TextMatrix(.rows - 1, .ColIndex("打包")) = IIf(rsTmp!打包 = 0, "", "√")
            .TextMatrix(.rows - 1, .ColIndex("启用")) = IIf(rsTmp!启用 = 0, IIf(rsTmp!批次 = 0, "√", ""), "√")
            .TextMatrix(.rows - 1, .ColIndex("药品类型")) = NVL(rsTmp!药品类型)
            
            If .TextMatrix(.rows - 1, .ColIndex("启用")) = "" Then
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &HE0E0E0
            Else
                .Cell(flexcpBackColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H80000005
            End If
            
            .Cell(flexcpBackColor, .rows - 1, .ColIndex("颜色"), .rows - 1, .ColIndex("颜色")) = IIf(rsTmp!批次 = 0, &H80000005, rsTmp!颜色)
            rsTmp.MoveNext
        Loop
        
        vsfBatch.Enabled = IsHavePrivs(mstrPrivs, "设置工作批次")
        If vsfBatch.Enabled = False Then
            Label2.Caption = Label2.Caption & "(无权限进行修改)"
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnEdit = False
End Sub

Private Sub lvwPRI_DblClick()
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer
    
    strReturn = ReturnSelectedPri(0, strIDS)
    
    If mintPri = 1 Then
        With Me.vsfPri
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    ElseIf mintPri = 2 Then
        With Me.vsfVolume
            .TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = .ColIndex("科室名称") Then
                .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
            End If
        End With
    ElseIf mintPri = 3 Then
        With Me.VSFPrice
            If mintCol = .ColIndex("配药类型") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "该配药类型已经添加，请重新选择！", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
        End With
    ElseIf mintPri = 4 Then
        Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice.ColIndex("收费项目") Then
            VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("项目id")) = strIDS
        End If
    ElseIf mintPri = 5 Then
        With Me.VSFPrice_给药途径
            If mintCol = .ColIndex("给药途径") Then
                For i = 1 To .rows - 1
                    If strReturn = .TextMatrix(i, mintCol) Then
                        MsgBox "该给药途径已经添加，请重新选择！", vbInformation + vbOKOnly
                        Exit Sub
                    End If
                Next
            End If
            
            .TextMatrix(mintRow, mintCol) = strReturn
            .TextMatrix(mintRow, .ColIndex("诊疗id")) = strIDS
        End With
    ElseIf mintPri = 6 Then
        Me.VSFPrice_给药途径.TextMatrix(mintRow, mintCol) = strReturn
        If mintCol = VSFPrice_给药途径.ColIndex("收费项目") Then
            VSFPrice_给药途径.TextMatrix(mintRow, VSFPrice_给药途径.ColIndex("项目id")) = strIDS
        End If
    End If
End Sub

Private Sub lvwPRI_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With lvwPRI
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        
        Item.Selected = True
        If Mid(Item.Text, 1, 2) = "所有" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub lvwPRI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strIDS As String
    Dim strReturn As String
    Dim i As Integer

    If KeyCode = vbKeyReturn Then
        strReturn = ReturnSelectedPri(1, strIDS)
        
        If mintPri = 1 Then
            With Me.vsfPri
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("科室名称") Then
                    .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
                End If
            End With
        ElseIf mintPri = 2 Then
            With Me.vsfVolume
                .TextMatrix(mintRow, mintCol) = strReturn
                If mintCol = .ColIndex("科室名称") Then
                    .TextMatrix(mintRow, .ColIndex("科室id")) = strIDS
                End If
            End With
        ElseIf mintPri = 3 Then
            With Me.VSFPrice
                If mintCol = .ColIndex("配药类型") Then
                    For i = 1 To .rows - 1
                        If strReturn = .TextMatrix(i, mintCol) Then
                            MsgBox "该配药类型已经添加，请重新选择！", vbInformation + vbOKOnly
                            Exit Sub
                        End If
                    Next
                End If
                
                .TextMatrix(mintRow, mintCol) = strReturn
            End With
        ElseIf mintPri = 4 Then
            Me.VSFPrice.TextMatrix(mintRow, mintCol) = strReturn
            If mintCol = VSFPrice.ColIndex("收费项目") Then
                VSFPrice.TextMatrix(mintRow, VSFPrice.ColIndex("项目id")) = strIDS
            End If
        
        End If
    End If
End Sub

'Private Sub Lvw药品剂型_DblClick()
'    ReturnSelected剂型 0
'End Sub
'
'
'Private Sub ReturnSelected剂型(ByVal intType As Integer)
'    'intType:0-双击剂型列表时；1-剂型列表中按回车时
'    Dim n As Integer
'
'    With Lvw药品剂型
'        If .SelectedItem Is Nothing Then Exit Sub
'        Me.txt药品剂型.Text = ""
'
'        '如果选择了全选，则不用取所有给药途径了
'        If .ListItems(1).Checked Then
'            Me.txt药品剂型.Text = "所有药品剂型"
'            pic剂型.Visible = False
'            Exit Sub
'        End If
'
'        For n = 1 To .ListItems.count
'            If .ListItems(n).Checked Then
'                Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.ListItems(n).Text, InStr(1, .ListItems(n).Text, "-") + 1))
'            End If
'        Next
'
'        If intType = 0 Then
'            '如果当前双击的给药途径未被选上，将当前双击的给药途径也加入到编辑框中
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                Me.txt药品剂型.Text = IIf(Me.txt药品剂型.Text = "", Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1), Me.txt药品剂型.Text & "," & Mid(.SelectedItem.Text, InStr(1, .SelectedItem.Text, "-") + 1))
'            End If
'
'            If .ListItems(1).Checked Then
'                 Me.txt药品剂型.Text = "所有药品剂型"
'                pic剂型.Visible = False
'                Exit Sub
'            End If
'        End If
'
'        pic剂型.Visible = False
'
'        CmdOK.Enabled = True
'        CmdCancel.Enabled = True
'    End With
'End Sub

Private Function ReturnSelectedPri(ByVal intType As Integer, ByRef strIDS As String) As String
    'intType:0-双击列表时；1-列表中按回车时
    Dim n As Integer
    Dim strReturn As String
    
    With lvwPRI
        If .SelectedItem Is Nothing Then Exit Function
        
        strReturn = .SelectedItem.Text
        strIDS = Mid(.SelectedItem.Key, 2)
        
'        '如果选择了全选，则不用取所有选项了
'        If .ListItems(1).Checked Then
'            strReturn = .ListItems(1).Text
'            ReturnSelectedPri = strReturn
'            picPRI.Visible = False
'            Exit Function
'        End If
'
'        For n = 1 To .ListItems.Count
'            If .ListItems(n).Checked Then
'                strReturn = IIf(strReturn = "", .ListItems(n).Text, strReturn & "," & .ListItems(n).Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'        Next
'
'        If intType = 0 Then
'            '如果当前双击的选项未被选上，将当前双击的选项也加入到编辑框中
'            If .SelectedItem.Checked = False Then
'                .SelectedItem.Checked = True
'                strReturn = IIf(strReturn = "", .SelectedItem.Text, strReturn & "," & .SelectedItem.Text)
'                strIDS = IIf(strIDS = "", Mid(.ListItems(n).Key, 2), strIDS & "," & Mid(.ListItems(n).Key, 2))
'            End If
'
'            If .ListItems(1).Checked Then
'                strReturn = .ListItems(1).Text
'                ReturnSelectedPri = strReturn
'                Exit Function
'            End If
'        End If
        
        picPRI.Visible = False
        
        CmdOK.Enabled = True
        CmdCancel.Enabled = True
        ReturnSelectedPri = strReturn
        mblnEdit = True
    End With
End Function

Private Sub Lvw药品剂型_ItemCheck(ByVal Item As MSComctlLib.listItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With Lvw药品剂型
        For n = 1 To .ListItems.count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "所有药品剂型" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

'Private Sub Lvw药品剂型_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        ReturnSelected剂型 1
'    End If
'End Sub
'
'
'Private Sub Lvw药品剂型_LostFocus()
''    Lvw药品剂型.Visible = False
'End Sub


Private Sub picPRI_Resize()
    On Error Resume Next
    
    With lvwPRI
        .Top = 0
        .Left = 0
        .Width = picPRI.Width
        .Height = picPRI.Height - 200 - cmdNO.Height
    End With
    
    With cmdNO
        .Top = picPRI.Height - .Height - 50
        .Left = picPRI.Width - .Width - 50
    End With
    
    With cmdYes
        .Top = cmdNO.Top
        .Left = cmdNO.Left - .Width - 100
    End With
End Sub

Private Sub pic剂型_Resize()
    On Error Resume Next
    
    With Lvw药品剂型
        .Top = 0
        .Left = 0
        .Width = pic剂型.Width
        .Height = pic剂型.Height - cmd剂型Ok.Height - 100
    End With
    
    With cmd剂型Cancel
        .Top = pic剂型.Height - .Height - 50
        .Left = pic剂型.Width - .Width - 50
    End With
    
    With cmd剂型Ok
        .Top = cmd剂型Cancel.Top
        .Left = cmd剂型Cancel.Left - .Width - 100
    End With
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    Dim i As Integer
    
    If PreviousTab = 2 And pic剂型.Visible = True Then
'        Call cmd剂型Cancel_Click
    ElseIf PreviousTab = 5 Then
        Me.vsfVolume.Row = Me.vsfVolume.rows - 1
        Me.vsfVolume.Col = Me.vsfVolume.ColIndex("科室名称")
    End If
End Sub

Private Sub txt输液总量_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub Load药品剂型()
    Dim rsData As ADODB.Recordset

    Set rsData = DeptSendWork_Get剂型(mlng库房id)
    
    With Lvw药品剂型
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.count + 1, "所有药品剂型", 1, 1
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.count + 1, rsData!剂型, 1, 1
            rsData.MoveNext
        Loop
    End With
End Sub


Private Sub LoadPRI()

    Set mRsDept = DeptSendWork_Get科室名称
    
    Set mRsType = DeptSendWork_Get配药类型
    
    Set mRsPC = DeptSendWork_Get频次
    
    Set mRsPrice = DeptSendWork_Get收费项目
        
    Set mRsWay = DeptSendWork_给药途径
End Sub


Private Sub updDeff_DownClick()
    If Me.txtDeff.Text <> "0" Then
        Me.txtDeff.Text = Me.txtDeff.Text - 1
    End If
End Sub

Private Sub updDeff_UpClick()
    If Me.txtDeff.Text <> "24" Then
        Me.txtDeff.Text = Me.txtDeff.Text + 1
    End If
End Sub

Private Sub vsfBatch_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBatch
        Select Case Col
            Case .ColIndex("配置时间开始"), .ColIndex("配置时间结束"), .ColIndex("给药时间开始"), .ColIndex("给药时间结束")
                If .TextMatrix(Row, Col) = "" Then Exit Sub
                
                If IsDate(.TextMatrix(Row, Col)) = False Then
                    MsgBox "请录入时间格式，比如12:59或者9:20等。", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
                
                If Col = .ColIndex("配置时间开始") And .TextMatrix(Row, .ColIndex("配置时间结束")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("配置时间开始"))) >= CDate(.TextMatrix(Row, .ColIndex("配置时间结束"))) Then
                        MsgBox "开始时间必须小于结束时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("配置时间结束") And .TextMatrix(Row, .ColIndex("配置时间开始")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("配置时间结束"))) <= CDate(.TextMatrix(Row, .ColIndex("配置时间开始"))) Then
                        MsgBox "结束时间必须大于开始时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("给药时间开始") And .TextMatrix(Row, .ColIndex("给药时间结束")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("给药时间开始"))) >= CDate(.TextMatrix(Row, .ColIndex("给药时间结束"))) Then
                        MsgBox "开始时间必须小于结束时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
                
                If Col = .ColIndex("给药时间结束") And .TextMatrix(Row, .ColIndex("给药时间开始")) <> "" Then
                    If CDate(.TextMatrix(Row, .ColIndex("给药时间结束"))) <= CDate(.TextMatrix(Row, .ColIndex("给药时间开始"))) Then
                        MsgBox "结束时间必须大于开始时间，请重新设置。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = ""
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfBatch_DblClick()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        If (.Col <> .ColIndex("打包") And .Col <> .ColIndex("启用")) And .Col <> .ColIndex("颜色") Then Exit Sub
        If (.MouseRow <> .Row Or .MouseCol <> .Col) And .Col <> .ColIndex("颜色") Then Exit Sub
        
        If .Col <> .ColIndex("颜色") Then
            If .TextMatrix(.Row, .Col) = "√" Then
                If .TextMatrix(.Row, .ColIndex("批次")) = "0#" And .Col = .ColIndex("启用") Then
                    MsgBox "0批次作为特殊批次，无法设置为【不启用】状态！"
                Else
                    .TextMatrix(.Row, .Col) = ""
                End If
            Else
                .TextMatrix(.Row, .Col) = "√"
            End If
            
            If .Col = .ColIndex("启用") Then
                If .TextMatrix(.Row, .Col) = "" Then
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HE0E0E0
                Else
                    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &H80000005
                End If
            End If
        
        Else
            On Error GoTo errHandle
            cmdialog.CancelError = True
            cmdialog.ShowColor
            .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = cmdialog.Color
            
errHandle:
        End If
    End With
End Sub


Private Sub vsfBatch_EnterCell()
    With vsfBatch
        If .Row < 2 Then Exit Sub
        .Editable = flexEDNone
        
        If .Col = .ColIndex("配置时间开始") Or .Col = .ColIndex("配置时间结束") Or .Col = .ColIndex("给药时间开始") Or .Col = .ColIndex("给药时间结束") Or .Col = .ColIndex("药品类型") Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfBatch_KeyPress(KeyAscii As Integer)
    With vsfBatch
        If KeyAscii = 13 Then
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row < .rows - 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("配置时间开始")
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfBatch
        Select Case Col
            Case .ColIndex("配置时间开始"), .ColIndex("配置时间结束"), .ColIndex("给药时间开始"), .ColIndex("给药时间结束")
                If InStr("1234567890:" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(":") Then
                    If InStr(.EditText, ":") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row <> 1 Then Cancel = True
End Sub

Private Sub vsfDept_EnterCell()
    If mblnEdit Then
        If MsgBox("请保存设置的优先级，切换科室后所作的优先级设置将失效，是否切换？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
            mblnEdit = False
            
        End If
    Else
        If Me.vsfDept.Row > 1 Then
            Call LoadVsfPRI(Val(Me.vsfDept.TextMatrix(vsfDept.Row, vsfDept.ColIndex("科室id"))))
        End If
    End If
    
End Sub

Private Sub vsfDept_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    
    With Me.vsfDept
        If KeyAscii <> 13 Or .TextMatrix(1, .ColIndex("科室名称")) = "" Or .Row <> 1 Then Exit Sub
        
        For intRow = 2 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("简码")) = UCase(.TextMatrix(1, .ColIndex("科室名称"))) Then
                .Row = intRow
                Exit Sub
            End If
        Next
    End With
End Sub

Private Sub vsfNoMedi_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strkey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = GetControlRect(vsfNoMedi.hWnd)
        dblLeft = vRect.Left + vsfNoMedi.CellLeft
        dblTop = vRect.Top + vsfNoMedi.CellTop + vsfNoMedi.CellHeight + 3200
        
        With vsfNoMedi
            If Col = .ColIndex("药品名称与编码") Then
                strkey = Trim(.EditText)
                If strkey = "" Then Exit Sub
                
                If IsNumeric(strkey) Then
                    '纯数字
                    StrCode = " d.编码 like [1] "
                ElseIf zlCommFun.IsCharAlpha(strkey) Then
                    '纯字母
                    StrCode = " n.简码 Like [1] "
                ElseIf zlCommFun.IsCharChinese(strkey) Then
                    '纯汉字
                    StrCode = " d.名称 like [1] "
                Else
                    StrCode = " (n.简码 Like [1] Or d.编码 Like [1] Or n.名称 Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')' As 通用名" & vbNewLine & _
                    " From 药品规格 T, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    " Where t.药品id = d.Id And t.药品id = n.收费细目id And D.类别 In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.撤档时间 Is Null Or To_Char(d.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品名称与编码", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("药品ID"))) Then
                            MsgBox rsRecord!通用名 & "已经录入，请重新选择！", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("药品名称与编码")) = rsRecord!通用名
                    .EditText = rsRecord!通用名
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfNoMedi_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfPRI_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    mintPri = 1
    mintRow = Row
    mintCol = Col
    With Me.picPRI
        .Visible = True
    
        .Height = vsfPri.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfPri.Left
        .Width = vsfPri.Width
    End With
            
    Select Case Col
        Case vsfPri.ColIndex("科室名称")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "所有科室", 1, 1
                mRsDept.MoveFirst
                Do While Not mRsDept.EOF
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
                    mRsDept.MoveNext
                Loop
                .ListItems.Add , "_00", "其他科室", 1, 1
            End With
        Case vsfPri.ColIndex("配药类型")
            With Me.lvwPRI
                .ListItems.Clear
                mRsType.MoveFirst
                Do While Not mRsType.EOF
                    .ListItems.Add , "_" & mRsType!编码, mRsType!名称, 1, 1
                    mRsType.MoveNext
                Loop
                 .ListItems.Add , "_00", "其他类型", 1, 1
            End With
        Case vsfPri.ColIndex("频次")
            With Me.lvwPRI
                .ListItems.Clear
                .ListItems.Add , "_" & 0, "所有频次", 1, 1
                mRsPC.MoveFirst
                Do While Not mRsPC.EOF
                    .ListItems.Add , "_" & mRsPC!编码, mRsPC!名称 & "(" & mRsPC!英文名称 & ")", 1, 1
                    mRsPC.MoveNext
                Loop
                .ListItems.Add , "_00", "其他频次", 1, 1
            End With
    End Select
End Sub

Private Sub VSFPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_给药途径_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        Cancel = True
    End If
End Sub

Private Sub VSFPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice.Height
        .Top = sstMain.Top + tabPrice.Top + VSFPrice.Top
        .Left = sstMain.Left + tabPrice.Left + VSFPrice.Left
        .Width = VSFPrice.Width
    End With
    
    mintRow = Row
    mintCol = Col
    
    If Col = VSFPrice.ColIndex("配药类型") Then
        mintPri = 3
        With Me.lvwPRI
            .ListItems.Clear
            mRsType.MoveFirst
            Do While Not mRsType.EOF
                .ListItems.Add , "_" & mRsType!编码, mRsType!名称, 1, 1
                mRsType.MoveNext
            Loop
             .ListItems.Add , "_00", "其他类型", 1, 1
        End With
    ElseIf Col = VSFPrice.ColIndex("收费项目") Then
        mintPri = 4
        With Me.lvwPRI
            .ListItems.Clear
            mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!Id, mRsPrice!名称, 1, 1
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub VSFPrice_给药途径_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.picPRI
        .Visible = True
    
        .Height = VSFPrice_给药途径.Height
        .Top = sstMain.Top + tabPrice.Top + VSFPrice_给药途径.Top
        .Left = sstMain.Left + tabPrice.Left + VSFPrice_给药途径.Left
        .Width = VSFPrice_给药途径.Width
    End With
    
    mintRow = Row
    mintCol = Col
    
    If Col = VSFPrice_给药途径.ColIndex("给药途径") Then
        mintPri = 5
        With Me.lvwPRI
            .ListItems.Clear
            If mRsWay.RecordCount > 0 Then mRsWay.MoveFirst
            Do While Not mRsWay.EOF
                .ListItems.Add , "_" & mRsWay!Id, mRsWay!名称, 1, 1
                mRsWay.MoveNext
            Loop
        End With
    ElseIf Col = VSFPrice_给药途径.ColIndex("收费项目") Then
        mintPri = 6
        With Me.lvwPRI
            .ListItems.Clear
            If mRsPrice.RecordCount > 0 Then mRsPrice.MoveFirst
            Do While Not mRsPrice.EOF
                .ListItems.Add , "_" & mRsPrice!Id, mRsPrice!名称, 1, 1
                mRsPrice.MoveNext
            Loop
        End With
    End If
    
End Sub

Private Sub VSFPrice_EnterCell()
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    If Me.VSFPrice.Row < 2 Then
        cmdLast.Enabled = False
    ElseIf Me.VSFPrice.Row = Me.VSFPrice.rows - 1 Then
        cmdNext.Enabled = False
    End If
    
    VSFPrice.Editable = flexEDNone
    
    If VSFPrice.ColSel <> VSFPrice.ColIndex("优先级") Then
        VSFPrice.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub VSFPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice.Row = VSFPrice.rows - 1 Then
        With Me.VSFPrice
            If .TextMatrix(.Row, .ColIndex("配药类型")) <> "" And .TextMatrix(.Row, .ColIndex("收费项目")) <> "" Then
                .rows = .rows + 1
                .Row = .rows - 1
                .Col = .ColIndex("配药类型")
                .TextMatrix(.Row, .ColIndex("优先级")) = .Row
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice.Row
        If VSFPrice.rows = 2 Then
           VSFPrice.rows = 1
           VSFPrice.rows = 2
        Else
            Me.VSFPrice.RemoveItem VSFPrice.Row
        End If
        
        '调整序号
        For i = intRow To Me.VSFPrice.rows - 1
            Me.VSFPrice.TextMatrix(i, Me.VSFPrice.ColIndex("优先级")) = i
        Next
    End If
    
End Sub

Private Sub VSFPrice_给药途径_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim i As Integer
    
    If VSFPrice_给药途径.Row = 0 Then Exit Sub
    If KeyCode = 13 And VSFPrice_给药途径.Row = VSFPrice_给药途径.rows - 1 Then
        Me.VSFPrice_给药途径.Editable = flexEDNone
        With Me.VSFPrice_给药途径
            If .TextMatrix(.Row, .ColIndex("给药途径")) <> "" And .TextMatrix(.Row, .ColIndex("收费项目")) <> "" Then
                .rows = .rows + 1
                .Row = .rows - 1
                .Col = .ColIndex("给药途径")
            End If
        End With
    ElseIf KeyCode = 46 Then
        intRow = VSFPrice_给药途径.Row
        If VSFPrice_给药途径.rows = 2 Then
           VSFPrice_给药途径.rows = 1
           VSFPrice_给药途径.rows = 2
        Else
            Me.VSFPrice_给药途径.RemoveItem VSFPrice_给药途径.Row
        End If
    End If
    Me.VSFPrice_给药途径.Editable = flexEDKbd
    
End Sub


Private Sub vsfPrint_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Me.vsfPrint.rows = 2 Then
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("药品id")) = ""
            Me.vsfPrint.TextMatrix(vsfPrint.Row, vsfPrint.ColIndex("药品名称与编码")) = ""
        Else
            Me.vsfPrint.RemoveItem vsfPrint.Row
        End If
        
    End If
End Sub

Private Sub vsfPrint_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim i As Integer
    Dim strkey As String
    Dim StrCode As String
    
    If KeyCode = 13 Then
        vRect = GetControlRect(vsfPrint.hWnd)
        dblLeft = vRect.Left + vsfPrint.CellLeft
        dblTop = vRect.Top + vsfPrint.CellTop + vsfPrint.CellHeight + 3200
        
        With vsfPrint
            If Col = .ColIndex("药品名称与编码") Then
                strkey = Trim(.EditText)
                If strkey = "" Then Exit Sub
                
                If IsNumeric(strkey) Then
                    '纯数字
                    StrCode = " d.编码 like [1] "
                ElseIf zlCommFun.IsCharAlpha(strkey) Then
                    '纯字母
                    StrCode = " n.简码 Like [1] "
                ElseIf zlCommFun.IsCharChinese(strkey) Then
                    '纯汉字
                    StrCode = " d.名称 like [1] "
                Else
                    StrCode = " (n.简码 Like [1] Or d.编码 Like [1] Or n.名称 Like [1]) "
                End If
                                
                gstrSQL = "Select Distinct d.Id ,'【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')' As 通用名" & vbNewLine & _
                    " From 药品规格 T, 收费项目目录 D, 收费项目别名 N" & vbNewLine & _
                    " Where t.药品id = d.Id And t.药品id = n.收费细目id And D.类别 In ('5', '6') And" & StrCode & vbNewLine & _
                    " And (d.撤档时间 Is Null Or To_Char(d.撤档时间, 'yyyy-MM-dd') = '3000-01-01')" & vbNewLine & _
                    " Order By '【' || d.编码 || '】' || d.名称 || '(' || d.规格 || ')'"
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品名称与编码", False, "", "", False, False, _
                True, dblLeft, dblTop, .Height, blnCancel, False, True, IIf(gstrMatchMethod = 0, "", "%") & UCase(.EditText) & "%")
    
                If rsRecord Is Nothing Then
                    .EditText = ""
                    Exit Sub
                Else
                    For i = 1 To .rows - 1
                        If rsRecord!Id = Val(.TextMatrix(i, .ColIndex("药品ID"))) Then
                            MsgBox rsRecord!通用名 & "已经录入，请重新选择！", vbInformation + vbOKOnly, gstrSysName
                            .EditText = ""
                            Exit Sub
                        End If
                    Next
                    
                    .TextMatrix(.Row, .ColIndex("药品ID")) = rsRecord!Id
                    .TextMatrix(.Row, .ColIndex("药品名称与编码")) = rsRecord!通用名
                    .EditText = rsRecord!通用名
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPrint_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfVolume_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = Me.vsfVolume.ColIndex("容量") Then
        If Not IsNumeric(vsfVolume.TextMatrix(Row, Col)) Then
            MsgBox "容量请录入数字！", vbInformation + vbOKOnly, gstrSysName
            vsfVolume.Col = vsfVolume.ColIndex("容量")
        End If
    End If
End Sub

Private Sub vsfVolume_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str批次 As String
    Dim i As Integer
    
    If Col <> vsfVolume.ColIndex("配药批次") Then Exit Sub
    With Me.vsfBatch
        If .rows > 2 Then
            For i = 2 To .rows - 1
                If .TextMatrix(i, .ColIndex("批次")) <> "" And .TextMatrix(i, .ColIndex("启用")) <> "" Then
                    str批次 = IIf(str批次 = "", "", str批次 & "|") & .TextMatrix(i, .ColIndex("批次"))
                End If
            Next
        End If
        If str批次 <> "" Then Me.vsfVolume.ColComboList(vsfVolume.ColIndex("配药批次")) = str批次
    End With
End Sub

Private Sub vsfVolume_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    mblnPri = False
'    mintRow = Row
'
'    mintCol = Col
''    With Me.picPRI
'        .Visible = True
'        .Height = vsfVolume.Height
'        .Top = sstMain.Top + vsfPri.Top
'        .Left = sstMain.Left + vsfVolume.Left
'        .Width = vsfVolume.Width
'    End With
'
'    With vsfVolume
'        If Col = .ColIndex("科室名称") Then
'            With Me.lvwPRI
'                .ListItems.Clear
'                .ListItems.Add , "_" & 0, "所有科室", 1, 1
'                mRsDept.MoveFirst
'                Do While Not mRsDept.EOF
'                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
'                    mRsDept.MoveNext
'                Loop
'                .ListItems.Add , "_00", "其他科室", 1, 1
'            End With
'        End If
'    End With

    mintPri = 2
    mintRow = vsfVolume.Row
    mintCol = vsfVolume.Col

    With Me.lvwPRI
        .ListItems.Clear
        .ListItems.Add , "_" & 0, "所有科室", 1, 1
        mRsDept.MoveFirst
        Do While Not mRsDept.EOF
            If vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) <> "" Then
                If mRsDept!简码 = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or mRsDept!五笔简码 = UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) Or vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!名称 Then
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col) = mRsDept!名称
                    vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.ColIndex("科室id")) = mRsDept!Id
                    Exit Sub

                ElseIf InStr(1, mRsDept!五笔简码, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!简码, UCase(vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col))) > 0 Or InStr(1, mRsDept!名称, vsfVolume.TextMatrix(vsfVolume.Row, vsfVolume.Col)) > 0 Then
                    .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
                End If
            Else
                .ListItems.Add , "_" & mRsDept!Id, mRsDept!名称, 1, 1
            End If
            mRsDept.MoveNext
        Loop
        
        If .ListItems.count = 1 Then
            .ListItems.Clear
            MsgBox "你输入的简码没有与之匹配的科室，请重新录入！"
            Exit Sub
        End If
        
        .ListItems.Add , "_00", "其他科室", 1, 1
    End With
    

    With Me.picPRI
        .Visible = True
        .Height = vsfVolume.Height
        .Top = sstMain.Top + vsfPri.Top
        .Left = sstMain.Left + vsfVolume.Left
        .Width = vsfVolume.Width
    End With
End Sub

Private Sub loadVolume()
    Dim rsTemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    gstrSQL = "select 科室id,科室名称,容量,配药批次 from 科室容量设置 where 配置中心ID=[1] order by 科室id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室容量数据", Me.CboStore.ItemData(Me.CboStore.ListIndex))
    
    i = 1
    With Me.vsfVolume
        .RowHeight(0) = 250
        .rows = 1
        .rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        Do While Not rsTemp.EOF
            .RowHeight(i) = 250
            .TextMatrix(i, .ColIndex("科室id")) = rsTemp!科室ID
            .TextMatrix(i, .ColIndex("科室名称")) = rsTemp!科室名称
            .TextMatrix(i, .ColIndex("配药批次")) = NVL(rsTemp!配药批次)
            .TextMatrix(i, .ColIndex("容量")) = rsTemp!容量
            i = i + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfVolume_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'
'    With Me.vsfVolume
'        If .Row = .rows - 1 Then
'            If .Col = .Cols - 1 Then
'                Exit Sub
'            Else
'                .Col = .Col + 1
'            End If
'        Else
'            If .Col = .Cols - 1 Then
'                .Row = .Row + 1
'                .Col = .ColIndex("科室名称")
'            Else
'                .Col = .Col + 1
'            End If
'        End If
'    End With
End Sub

Private Sub vsfVolume_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With Me.vsfVolume
        If .Row = .rows - 1 Then
            If .Col = .Cols - 1 Then
                Exit Sub
            Else
                .Col = .Col + 1
            End If
        Else
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .ColIndex("科室名称")
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfVolume_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    With vsfVolume
        If Col = .ColIndex("容量") Then
            If InStr("1234567890-." & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With

End Sub

Private Sub LoadDept()
    Dim i As Integer
    
    i = 1
    vsfDept.rows = mRsDept.RecordCount + 2
    Do While Not mRsDept.EOF
        With Me.vsfDept
            i = i + 1
            .TextMatrix(i, .ColIndex("序号")) = i - 1
            .TextMatrix(i, .ColIndex("科室id")) = mRsDept!Id
            .TextMatrix(i, .ColIndex("科室名称")) = mRsDept!名称
            .TextMatrix(i, .ColIndex("简码")) = mRsDept!简码
        End With
        mRsDept.MoveNext
    Loop
End Sub

Private Sub vsfPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfPrint.ColIndex("药品名称与编码") Then Cancel = True
End Sub

Private Sub vsfNoMedi_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfNoMedi.ColIndex("药品名称与编码") Then Cancel = True
End Sub

Private Sub vsfNoMedi_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = 46 Then
        If Me.vsfNoMedi.rows = 2 Then
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("药品id")) = ""
            Me.vsfNoMedi.TextMatrix(vsfNoMedi.Row, vsfNoMedi.ColIndex("药品名称与编码")) = ""
        Else
            Me.vsfNoMedi.RemoveItem vsfNoMedi.Row
        End If
    End If
End Sub
