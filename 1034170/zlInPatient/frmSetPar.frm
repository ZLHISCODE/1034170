VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetPar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frmSetPar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7170
      TabIndex        =   2
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7200
      TabIndex        =   0
      Top             =   450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin TabDlg.SSTab sTab 
      Height          =   5550
      Left            =   135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   9790
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "输入控制(&1)"
      TabPicture(0)   =   "frmSetPar.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFee"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDeviceSetup"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "optDept(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "optDept(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk记帐"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chk担保"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkSeekName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNameDays"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkLedWelcome"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDiagDays"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chk计算"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cbo预交结算"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboFee"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "预交票据控制(&2)"
      TabPicture(1)   =   "frmSetPar.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrepay"
      Tab(1).Control(1)=   "fraWristlet"
      Tab(1).Control(2)=   "fraPatientPage"
      Tab(1).Control(3)=   "fraDeposit"
      Tab(1).Control(4)=   "img16"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "医疗卡票据控制(&3)"
      TabPicture(2)   =   "frmSetPar.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chk扫描身份证签约"
      Tab(2).Control(1)=   "fraTitle"
      Tab(2).Control(2)=   "cboType"
      Tab(2).Control(3)=   "lbl缺省发卡"
      Tab(2).ControlCount=   4
      Begin VB.ComboBox cboFee 
         Height          =   300
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   4695
         Width           =   2580
      End
      Begin VB.ComboBox cbo预交结算 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   5070
         Width           =   2580
      End
      Begin VB.CheckBox chk扫描身份证签约 
         Caption         =   "扫描身份证签约"
         Height          =   180
         Left            =   -74715
         TabIndex        =   66
         Top             =   4335
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.CheckBox chk计算 
         Caption         =   "入院时自动计算一次费用"
         Height          =   180
         Left            =   2985
         MaskColor       =   &H00000000&
         TabIndex        =   63
         Top             =   4455
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2040
         TabIndex        =   59
         Top             =   4350
         Width           =   285
      End
      Begin VB.TextBox txtDiagDays 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   57
         Text            =   "3"
         Top             =   4154
         Width           =   285
      End
      Begin VB.Frame Frame1 
         Caption         =   "输入光标经过项目"
         Height          =   3030
         Left            =   150
         TabIndex        =   34
         Top             =   420
         Width           =   6405
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人身份证号"
            Height          =   195
            Index           =   26
            Left            =   3450
            TabIndex        =   55
            Top             =   930
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位开户行"
            Height          =   195
            Index           =   20
            Left            =   3450
            TabIndex        =   62
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位邮编"
            Height          =   195
            Index           =   19
            Left            =   3450
            TabIndex        =   60
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位电话"
            Height          =   195
            Index           =   18
            Left            =   3450
            TabIndex        =   58
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "工作单位"
            Height          =   195
            Index           =   17
            Left            =   3450
            TabIndex        =   56
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人电话"
            Height          =   195
            Index           =   16
            Left            =   3450
            TabIndex        =   54
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人地址"
            Height          =   195
            Index           =   15
            Left            =   3450
            TabIndex        =   53
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人关系"
            Height          =   195
            Index           =   14
            Left            =   1785
            TabIndex        =   49
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "联系人姓名"
            Height          =   195
            Index           =   13
            Left            =   1785
            TabIndex        =   48
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "家庭电话"
            Height          =   195
            Index           =   12
            Left            =   1785
            TabIndex        =   47
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "家庭地址邮编"
            Height          =   195
            Index           =   11
            Left            =   1785
            TabIndex        =   46
            Top             =   930
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "现住址"
            Height          =   195
            Index           =   10
            Left            =   1785
            TabIndex        =   45
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "出生地点"
            Height          =   195
            Index           =   9
            Left            =   1785
            TabIndex        =   44
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "身份证号"
            Height          =   195
            Index           =   8
            Left            =   285
            TabIndex        =   43
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "出生日期"
            Height          =   195
            Index           =   7
            Left            =   285
            TabIndex        =   41
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "身份"
            Height          =   195
            Index           =   6
            Left            =   285
            TabIndex        =   40
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "职业"
            Height          =   195
            Index           =   5
            Left            =   285
            TabIndex        =   39
            Top             =   1530
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "婚姻状况"
            Height          =   195
            Index           =   4
            Left            =   285
            TabIndex        =   38
            Top             =   1230
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "学历"
            Height          =   195
            Index           =   3
            Left            =   285
            TabIndex        =   37
            Top             =   930
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "民族"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   36
            Top             =   615
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "国籍"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   35
            Top             =   315
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "单位帐号"
            Height          =   195
            Index           =   0
            Left            =   3450
            TabIndex        =   64
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "其他证件"
            Height          =   195
            Index           =   21
            Left            =   285
            TabIndex        =   42
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "户口地址"
            Height          =   195
            Index           =   22
            Left            =   1785
            TabIndex        =   50
            Top             =   2145
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "户口地址邮编"
            Height          =   195
            Index           =   23
            Left            =   1785
            TabIndex        =   51
            Top             =   2445
            Value           =   1  'Checked
            Width           =   1440
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "区域"
            Height          =   195
            Index           =   24
            Left            =   1785
            TabIndex        =   52
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "籍贯"
            Height          =   195
            Index           =   25
            Left            =   3450
            TabIndex        =   65
            Top             =   2745
            Value           =   1  'Checked
            Width           =   1200
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED显示欢迎信息"
         Height          =   225
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   "收费窗口输入病人后,是否显示欢迎信息并发声"
         Top             =   3600
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "本地共用预交票据"
         Height          =   2535
         Left            =   -74895
         TabIndex        =   31
         Top             =   420
         Width           =   6510
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   2145
            Left            =   60
            TabIndex        =   32
            Top             =   270
            Width           =   6285
            _cx             =   11086
            _cy             =   3784
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
            FormatString    =   $"frmSetPar.frx":0060
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
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用医疗卡"
         Height          =   3570
         Left            =   -74865
         TabIndex        =   28
         Top             =   555
         Width           =   6390
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   3150
            Left            =   60
            TabIndex        =   29
            Top             =   300
            Width           =   6150
            _cx             =   10848
            _cy             =   5556
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
            FormatString    =   $"frmSetPar.frx":013F
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
         Left            =   -73575
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4635
         Width           =   2580
      End
      Begin VB.Frame fraWristlet 
         Caption         =   "病人腕带"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   22
         Top             =   4365
         Width           =   6465
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2685
            TabIndex        =   26
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   25
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   24
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   2
            Left            =   5370
            TabIndex        =   23
            Top             =   160
            Width           =   990
         End
      End
      Begin VB.Frame fraPatientPage 
         Caption         =   "病案首页"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   17
         Top             =   3690
         Width           =   6465
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   20
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2655
            TabIndex        =   19
            Top             =   285
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   1
            Left            =   5370
            TabIndex        =   18
            Top             =   160
            Width           =   990
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "预交款票据"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74865
         TabIndex        =   12
         Top             =   3000
         Width           =   6480
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "选择是否打印"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   16
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "自动打印"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   15
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "不打印"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   14
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "打印设置"
            Height          =   345
            Index           =   0
            Left            =   5355
            TabIndex        =   13
            Top             =   180
            Width           =   990
         End
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
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   3090
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "0表示查找时不限制时间"
         Top             =   3884
         Width           =   285
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3040
         TabIndex        =   5
         Top             =   4080
         Width           =   285
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "允许通过输入姓名来模糊查找    天内的病人信息"
         Height          =   195
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   3877
         Width           =   4620
      End
      Begin VB.CheckBox chk担保 
         Caption         =   "输入病人担保信息"
         Height          =   195
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   3600
         Width           =   1740
      End
      Begin VB.CheckBox chk记帐 
         Caption         =   "医疗卡费用以记账方式收取"
         Height          =   180
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   4455
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.OptionButton optDept 
         Caption         =   "先选科室"
         Height          =   255
         Index           =   0
         Left            =   405
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   4725
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optDept 
         Caption         =   "先选病区"
         Height          =   255
         Index           =   1
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   4740
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   4830
         TabIndex        =   6
         Top             =   5070
         Width           =   1500
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   -70560
         Top             =   1320
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
               Picture         =   "frmSetPar.frx":0221
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "缺省费别"
         Height          =   180
         Left            =   2985
         TabIndex        =   70
         Top             =   4770
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "缺省缴款方式"
         Height          =   225
         Left            =   390
         TabIndex        =   68
         Top             =   5145
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "预约接收时提取病人    天内的诊断信息"
         Height          =   180
         Left            =   405
         TabIndex        =   61
         Top             =   4154
         Width           =   3855
      End
      Begin VB.Label lbl缺省发卡 
         Caption         =   "缺省发卡类型"
         Height          =   225
         Left            =   -74730
         TabIndex        =   30
         Top             =   4695
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmSetPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String
Public mlngModul As Long

 

Private Sub cboType_Click()
    chk扫描身份证签约.Enabled = Not (cboType.Text = "二代身份证")
    If cboType.Text = "二代身份证" Then
        chk扫描身份证签约.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1131)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    '光标经过项目
    For i = 0 To chkItem.UBound
        zlDatabase.SetPara chkItem(i).Caption, chkItem(i).Value, glngSys, mlngModul, IIf(chkItem(i).Enabled = True, True, False)
    Next
    
    Call SaveInvoice
    
    zlDatabase.SetPara "担保信息", chk担保.Value, glngSys, mlngModul, IIf(chk担保.Enabled = True, True, False)
    zlDatabase.SetPara "姓名模糊查找", chkSeekName.Value, glngSys, mlngModul, IIf(chkSeekName.Enabled = True, True, False)
    zlDatabase.SetPara "姓名查找天数", Val(txtNameDays.Text), glngSys, mlngModul, IIf(txtNameDays.Enabled = True, True, False)
    zlDatabase.SetPara "卡费记帐", chk记帐.Value, glngSys, mlngModul, IIf(chk记帐.Enabled = True, True, False)
    zlDatabase.SetPara "先选病区", IIf(optDept(1).Value, 1, 0), glngSys, mlngModul, IIf(optDept(1).Enabled = True, True, False)
    zlDatabase.SetPara "诊断查找天数", Val(txtDiagDays.Text), glngSys, mlngModul, True
    '36454,刘鹏飞,2012-09-06
    zlDatabase.SetPara "费用计算时机", chk计算.Value, glngSys, mlngModul, IIf(chk计算.Enabled = True, True, False)

    'LED设备
    zlDatabase.SetPara "LED显示欢迎信息", chkLedWelcome.Value, glngSys, mlngModul, IIf(chkLedWelcome.Enabled = True, True, False)
    '预交款票据打印
    For i = 0 To optPrepayPrint.UBound
        If optPrepayPrint(i).Value Then
            zlDatabase.SetPara "预交款票据打印", i, glngSys, mlngModul, IIf(optPrepayPrint(i).Enabled = True, True, False)
        End If
    Next
    
    '病案首页打印方式
    For i = 0 To optFpagePrint.UBound
        If optFpagePrint(i).Value Then
            zlDatabase.SetPara "病案首页打印", i, glngSys, mlngModul, IIf(optFpagePrint(i).Enabled = True, True, False)
        End If
    Next
    
    '病人腕带打印方式
    For i = 0 To optWristletPrint.UBound
        If optWristletPrint(i).Value Then
            zlDatabase.SetPara "病人腕带打印", i, glngSys, mlngModul, IIf(optWristletPrint(i).Enabled = True, True, False)
        End If
    Next
    '问题号:53408
    zlDatabase.SetPara "扫描身份证签约", IIf(chk扫描身份证签约.Value = 1, 1, 0), glngSys, glngModul, InStr(mstrPrivs, "参数设置") > 0
    
    Call InitLocPar(mlngModul)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1131_1", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset, objItem As ListItem
    Dim blnBill As Boolean
    
    gblnOK = False
    On Error GoTo errH
    
    '光标经过项目
    For i = 0 To chkItem.UBound
        chkItem(i).Value = zlDatabase.GetPara(chkItem(i).Caption, glngSys, mlngModul, 1, Array(chkItem(i)), InStr(mstrPrivs, "参数设置") > 0)
    Next
    Call InitShareInvoice
    
    chk担保.Value = IIf(zlDatabase.GetPara("担保信息", glngSys, mlngModul, , Array(chk担保), InStr(mstrPrivs, "担保信息") > 0) = "1", 1, 0)
    chkSeekName.Value = IIf(zlDatabase.GetPara("姓名模糊查找", glngSys, mlngModul, , Array(chkSeekName), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    txtNameDays.Text = Val(zlDatabase.GetPara("姓名查找天数", glngSys, mlngModul, , Array(txtNameDays), InStr(mstrPrivs, "参数设置") > 0))
    txtDiagDays.Text = Val(zlDatabase.GetPara("诊断查找天数", glngSys, mlngModul, "3", Array(txtDiagDays, Label1), InStr(mstrPrivs, "参数设置") > 0))
     '问题号:53408
    chk扫描身份证签约.Value = IIf(zlDatabase.GetPara("扫描身份证签约", glngSys, glngModul, , Array(chk扫描身份证签约), InStr(mstrPrivs, ";参数设置;") > 0) = "1", 1, 0)
    
    'LED设备
    chkLedWelcome.Value = zlDatabase.GetPara("LED显示欢迎信息", glngSys, mlngModul, 1, Array(chkLedWelcome), InStr(mstrPrivs, "参数设置") > 0)
        
    i = Val(zlDatabase.GetPara("先选病区", glngSys, mlngModul, , Array(optDept(0), optDept(1)), InStr(mstrPrivs, "参数设置") > 0))
    optDept(1).Value = (i = 1)
    optDept(0).Value = Not optDept(1).Value
    
    chk记帐.Value = IIf(zlDatabase.GetPara("卡费记帐", glngSys, mlngModul, , Array(chk记帐), InStr(mstrPrivs, "参数设置") > 0) = "1", 1, 0)
    
    i = Val(zlDatabase.GetPara("预交款票据打印", glngSys, mlngModul, , Array(fraDeposit), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optPrepayPrint.UBound Then optPrepayPrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("病案首页打印", glngSys, mlngModul, , Array(fraPatientPage), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optFpagePrint.UBound Then optFpagePrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("病人腕带打印", glngSys, mlngModul, , Array(fraWristlet), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optWristletPrint.UBound Then optWristletPrint(i).Value = True
    
    '36454,刘鹏飞,2012-09-06
    chk计算.Value = Val(zlDatabase.GetPara("费用计算时机", glngSys, mlngModul, "1", Array(chk计算), InStr(mstrPrivs, "参数设置") > 0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
  

Private Sub sTab_Click(PreviousTab As Integer)
    If sTab.Tab = 0 Then
        chkItem(1).SetFocus
    ElseIf sTab.Tab = 1 Then
        If vsPrepay.Enabled And vsPrepay.Visible Then vsPrepay.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
     End If
End Sub

Private Sub txtDiagDays_GotFocus()
    Call SelAll(txtDiagDays)
End Sub

Private Sub txtDiagDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDiagDays_Validate(Cancel As Boolean)
    If Val(txtDiagDays.Text) <= 0 Then
        txtDiagDays.Text = 0
    ElseIf Val(txtDiagDays.Text) > 999 Then
        txtDiagDays.Text = 999
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
Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1
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
    Dim str缺省费别 As String
    
    On Error GoTo ErrHand
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    '恢复列宽度
    lngCardTypeID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, , Array(cboType), blnHavePrivs, intType))
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
            .MoveNext
        Loop
    End With
    '问题号:58776
    For i = 0 To cboType.ListCount - 1
        If Val(cboType.ItemData(i)) = lngCardTypeID Then
             cboType.ListIndex = i
        End If
    Next
    
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
        rsTemp.Filter = " 使用类别<>1   "   '不包含预交门诊票据
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!使用类别, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "门诊和住院共用"
            ElseIf Val(Nvl(rsTemp!使用类别, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交门诊票据"
            Else
                .TextMatrix(lngRow, .ColIndex("预交类型")) = "预交住院票据"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("预交类型")) = Val(Nvl(rsTemp!使用类别))
            
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
    '加载缺省缴款方式(预交款)
    Load缴款方式
    '加载费别
    strSQL = "Select A.编码,A.名称,A.简码,Nvl(A.缺省标志,0) as 缺省 From 费别 A,Table(Cast(f_Num2List([1]) As zlTools.t_Numlist)) B " & _
             " Where (A.服务对象 = B.Column_Value or A.服务对象 is null) And A.属性=1 And Nvl(A.仅限初诊,0)=0 And  " & _
             "        Sysdate Between NVL(A.有效开始,Sysdate-1) and NVL(A.有效结束,Sysdate+1) Order by A.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "1,2,3")
    cboFee.Clear
    Do While Not rsTemp.EOF
        cboFee.AddItem rsTemp!名称
        If rsTemp!缺省 = 1 Then cboFee.ListIndex = cboFee.NewIndex
    rsTemp.MoveNext
    Loop
    If cboFee.ListCount > 0 And cboFee.ListIndex < 0 Then cboFee.ListIndex = 0
    str缺省费别 = zlDatabase.GetPara("缺省费别", glngSys, mlngModul, , blnHavePrivs)
    If str缺省费别 <> "" Then
        For i = 0 To cboFee.ListCount - 1
            If cboFee.List(i) = str缺省费别 Then
                cboFee.ListIndex = i
            End If
        Next
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存相关票据
    '编制:刘兴洪
    '日期:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    Dim lng卡类别ID As Long
    If cboType.ListIndex >= 0 Then
        lng卡类别ID = cboType.ItemData(cboType.ListIndex)
    End If
    zlDatabase.SetPara "缺省医疗卡类别", lng卡类别ID, glngSys, mlngModul, blnHavePrivs
        
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
    
    zlDatabase.SetPara "缺省缴款方式", Trim(cbo预交结算.Text), glngSys, mlngModul, blnHavePrivs
    '69489
    zlDatabase.SetPara "缺省费别", Trim(cboFee.Text), glngSys, mlngModul, blnHavePrivs
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

Public Sub Load缴款方式()
    Dim strTemp As String, str缺省预交款方式 As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim objSquareCard As Object
    Dim varData As Variant, varTemp As Variant
    Dim strPayType As String
    Dim j As Long, i As Long
    Dim blnFind As Boolean, blnHavePrivs As Boolean
    
    strTemp = "1,2,5,7,8" & IIf(InStr(mstrPrivs, ";保险病人登记;") > 0, ",3", "")

    
    strSQL = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合 ='预交款'  And B.名称=A.结算方式  " & _
        "           And Nvl(B.性质,1) In(" & strTemp & ")" & _
        " Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    strPayType = objSquareCard.zlGetAvailabilityCardType: varData = Split(strPayType, ";")
    With cbo预交结算
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            
            If Not blnFind Then
                .AddItem Nvl(rsTemp!名称)
                If rsTemp!缺省 = 1 Then .ListIndex = .NewIndex:  .Tag = .NewIndex
                .ItemData(.NewIndex) = Val(Nvl(rsTemp!性质))
                j = j + 1
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                j = j + 1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
        str缺省预交款方式 = zlDatabase.GetPara("缺省缴款方式", glngSys, mlngModul, , blnHavePrivs)
        If str缺省预交款方式 <> "" Then
            For i = 0 To cbo预交结算.ListCount
                If cbo预交结算.List(i) = str缺省预交款方式 Then
                    cbo预交结算.ListIndex = i
                End If
            Next
        End If
    End With
End Sub
