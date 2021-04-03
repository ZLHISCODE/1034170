VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOutAdviceEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊医嘱编辑"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmOutAdviceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1035
      Left            =   10845
      TabIndex        =   74
      Top             =   225
      Width           =   150
      _Version        =   589884
      _ExtentX        =   265
      _ExtentY        =   1826
      _StockProps     =   64
   End
   Begin VB.PictureBox picSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   0
      ScaleHeight     =   8340
      ScaleWidth      =   11085
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   0
      Width           =   11085
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9720
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   69
         Top             =   2160
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pic疑问 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   0
         ScaleHeight     =   270
         ScaleWidth      =   11070
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   5130
         Visible         =   0   'False
         Width           =   11070
         Begin VB.Label lbl疑问 
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   495
            TabIndex        =   61
            Top             =   45
            Width           =   1725
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   150
            Picture         =   "frmOutAdviceEdit.frx":058A
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame fra诊断 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   0
         TabIndex        =   56
         Top             =   960
         Width           =   11060
         Begin VB.CommandButton cmdLastDiag 
            Caption         =   "上次诊断"
            Height          =   300
            Left            =   50
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   320
            Width           =   900
         End
         Begin VB.OptionButton opt诊断 
            Caption         =   "按诊断标准"
            Height          =   180
            Index           =   0
            Left            =   9840
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   100
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton opt诊断 
            Caption         =   "按疾病编码"
            Height          =   180
            Index           =   1
            Left            =   9820
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   380
            Width           =   1200
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiag 
            Height          =   630
            Left            =   975
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   8760
            _cx             =   15452
            _cy             =   1111
            Appearance      =   2
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
            BackColorSel    =   13684944
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOutAdviceEdit.frx":0B14
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   115
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
         Begin VB.Label lbl诊断 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人诊断"
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   60
            Width           =   720
         End
      End
      Begin VB.Frame fraPati 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         TabIndex        =   37
         Top             =   510
         Width           =   10995
         Begin VB.CommandButton cmdAlley 
            Caption         =   "过敏史/病生状态"
            Height          =   350
            Left            =   9240
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.ComboBox cbo婴儿 
            Height          =   300
            ItemData        =   "frmOutAdviceEdit.frx":0CDD
            Left            =   9555
            List            =   "frmOutAdviceEdit.frx":0CF3
            Style           =   2  'Dropdown List
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   75
            Width           =   1395
         End
         Begin VB.Label lbl婴儿 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "婴儿"
            Height          =   180
            Left            =   9135
            TabIndex        =   30
            Top             =   135
            Width           =   360
         End
         Begin VB.Label lblPati 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名: 性别: 年龄: 门诊号: 费别: 医疗付款方式:"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   210
            TabIndex        =   38
            Top             =   135
            Width           =   4050
         End
      End
      Begin MSComCtl2.MonthView dtpDate 
         Height          =   2220
         Left            =   1725
         TabIndex        =   1
         Top             =   2505
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   179830785
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3645
         Left            =   60
         TabIndex        =   0
         Top             =   1650
         Width           =   10995
         _cx             =   19394
         _cy             =   6429
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   18
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOutAdviceEdit.frx":0D42
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         Begin MSComctlLib.ImageList img16 
            Left            =   1965
            Top             =   450
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":0E2A
                  Key             =   "警示"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":13C4
                  Key             =   "诊断"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":195E
                  Key             =   "诊断_当前"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutAdviceEdit.frx":1EF8
                  Key             =   "诊断_关联"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraAdvice 
         Height          =   2685
         Left            =   45
         TabIndex        =   39
         Top             =   5340
         Width           =   11040
         Begin VB.ComboBox cboDruPur 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   2220
            Width           =   1815
         End
         Begin VB.CommandButton cmdComExcReason 
            Height          =   300
            Left            =   10260
            Picture         =   "frmOutAdviceEdit.frx":2492
            Style           =   1  'Graphical
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "将当前说明设置为常用说明"
            Top             =   1860
            Width           =   315
         End
         Begin VB.CommandButton cmdExcReason 
            Caption         =   "…"
            Height          =   265
            Left            =   9930
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1875
            Width           =   285
         End
         Begin VB.TextBox txt超量说明 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   6255
            MaxLength       =   500
            TabIndex        =   18
            Top             =   1875
            Width           =   3945
         End
         Begin VB.ComboBox cbo分零 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1200
            Visible         =   0   'False
            Width           =   4260
         End
         Begin VB.CommandButton cmd医生嘱托 
            Caption         =   "…"
            Height          =   265
            Left            =   9960
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   573
            Width           =   285
         End
         Begin VB.TextBox cbo医生嘱托 
            Height          =   300
            Left            =   6255
            MaxLength       =   100
            TabIndex        =   24
            Top             =   555
            Width           =   4000
         End
         Begin VB.CommandButton cmd收藏用药理由 
            Height          =   300
            Left            =   10275
            Picture         =   "frmOutAdviceEdit.frx":2A1C
            Style           =   1  'Graphical
            TabIndex        =   64
            TabStop         =   0   'False
            ToolTipText     =   "将当前理由设置为常用理由。"
            Top             =   2250
            Width           =   315
         End
         Begin VB.CommandButton cmdReason 
            Caption         =   "…"
            Height          =   265
            Left            =   9930
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2250
            Width           =   285
         End
         Begin VB.PictureBox picHelp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5200
            Picture         =   "frmOutAdviceEdit.frx":2FA6
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   62
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chk免试 
            Caption         =   "免试"
            Height          =   180
            Left            =   4560
            TabIndex        =   5
            Top             =   255
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.TextBox txt用药理由 
            Height          =   300
            Left            =   4860
            MaxLength       =   1000
            TabIndex        =   21
            Top             =   2250
            Width           =   5385
         End
         Begin VB.CheckBox chkZeroBilling 
            Caption         =   "发送为不收费的记帐单(&F)"
            Height          =   225
            Left            =   8280
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   233
            Width           =   2370
         End
         Begin VB.ComboBox cbo滴速 
            Height          =   300
            Left            =   6255
            TabIndex        =   22
            Top             =   195
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CommandButton cmd安排时间 
            Height          =   240
            Left            =   2490
            Picture         =   "frmOutAdviceEdit.frx":97F8
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "选择日期(F4)"
            Top             =   1545
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txt安排时间 
            Height          =   300
            Left            =   960
            TabIndex        =   9
            Top             =   1515
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmd常用嘱托 
            Height          =   300
            Left            =   10275
            Picture         =   "frmOutAdviceEdit.frx":98EE
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "将当前嘱托设置为常用嘱托(Ctrl+D)"
            Top             =   555
            Width           =   315
         End
         Begin VB.ComboBox cbo附加执行 
            Height          =   300
            Left            =   6255
            TabIndex        =   29
            Text            =   "cbo附加执行"
            Top             =   1515
            Width           =   1860
         End
         Begin VB.TextBox txt天数 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2385
            MaxLength       =   3
            TabIndex        =   16
            Top             =   1875
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.CommandButton cmd频率 
            Height          =   240
            Left            =   4920
            Picture         =   "frmOutAdviceEdit.frx":9E78
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(F4)"
            Top             =   1545
            Width           =   270
         End
         Begin VB.TextBox txt频率 
            Height          =   300
            Left            =   3540
            TabIndex        =   13
            Top             =   1515
            Width           =   1665
         End
         Begin VB.TextBox txt单量 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   17
            Top             =   1875
            Width           =   1365
         End
         Begin VB.TextBox txt总量 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   960
            MaxLength       =   10
            TabIndex        =   15
            Top             =   1875
            Width           =   1530
         End
         Begin VB.CommandButton cmd用法 
            Height          =   240
            Left            =   2445
            Picture         =   "frmOutAdviceEdit.frx":9F6E
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(F4)"
            Top             =   1545
            Width           =   270
         End
         Begin VB.TextBox txt用法 
            Height          =   300
            Left            =   960
            TabIndex        =   11
            Top             =   1515
            Width           =   1815
         End
         Begin VB.CommandButton cmd开始时间 
            Height          =   240
            Left            =   2460
            Picture         =   "frmOutAdviceEdit.frx":A064
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "选择日期(F4)"
            Top             =   225
            Width           =   255
         End
         Begin VB.CheckBox chk紧急 
            Caption         =   "紧急医嘱(&E)"
            Height          =   225
            Left            =   3000
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   233
            Width           =   1290
         End
         Begin VB.CommandButton cmdExt 
            Height          =   285
            Left            =   4920
            Picture         =   "frmOutAdviceEdit.frx":A15A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "编辑(F4)"
            Top             =   552
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   285
            Left            =   4920
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   870
            Width           =   285
         End
         Begin VB.ComboBox cbo执行性质 
            Height          =   300
            ItemData        =   "frmOutAdviceEdit.frx":A250
            Left            =   9015
            List            =   "frmOutAdviceEdit.frx":A25D
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1200
            Width           =   1590
         End
         Begin VB.ComboBox cbo执行科室 
            Height          =   300
            Left            =   6255
            TabIndex        =   27
            Text            =   "cbo执行科室"
            Top             =   1200
            Width           =   1860
         End
         Begin VB.TextBox txt医嘱内容 
            Height          =   900
            Left            =   960
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "按 ~ 键显示成套方案选择器,Ctrl+F1调用医保信息"
            Top             =   552
            Width           =   3945
         End
         Begin VB.TextBox txt开始时间 
            Height          =   300
            Left            =   960
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   195
            Width           =   1815
         End
         Begin VB.ComboBox cbo执行时间 
            Height          =   300
            Left            =   6255
            TabIndex        =   26
            Top             =   877
            Width           =   4350
         End
         Begin VB.Label lbl超量说明 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "超量说明"
            Height          =   180
            Left            =   5490
            TabIndex        =   20
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label lbl分零 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "分零方式"
            Height          =   180
            Left            =   180
            TabIndex        =   68
            Top             =   1260
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl医嘱内容 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医嘱内容"
            Height          =   180
            Left            =   180
            TabIndex        =   67
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl用药目的 
            AutoSize        =   -1  'True
            Caption         =   "用药目的"
            Height          =   180
            Left            =   165
            TabIndex        =   59
            Top             =   2265
            Width           =   720
         End
         Begin VB.Label lbl用药理由 
            AutoSize        =   -1  'True
            Caption         =   "用药理由"
            Height          =   180
            Left            =   4080
            TabIndex        =   58
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label lbl滴速单位 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "滴/分钟"
            Height          =   180
            Left            =   7335
            TabIndex        =   55
            Top             =   255
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lbl滴速 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "↓滴速"
            Height          =   180
            Left            =   5670
            TabIndex        =   54
            Top             =   255
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lbl安排时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "安排时间"
            Height          =   180
            Left            =   165
            TabIndex        =   53
            Top             =   1575
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl附加执行 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "附加执行"
            Height          =   180
            Left            =   5490
            TabIndex        =   52
            Top             =   1575
            Width           =   720
         End
         Begin VB.Label lbl天数 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "用    天"
            Height          =   180
            Left            =   2205
            TabIndex        =   51
            Top             =   1935
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl频率 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "频率"
            Height          =   180
            Left            =   3135
            TabIndex        =   46
            Top             =   1575
            Width           =   360
         End
         Begin VB.Label lbl单量单位 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "单位"
            Height          =   180
            Left            =   4905
            TabIndex        =   42
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl单量 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "↓单量"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2955
            TabIndex        =   41
            ToolTipText     =   "常用剂量比例(Ctrl+~)"
            Top             =   1935
            Width           =   540
         End
         Begin VB.Label lbl总量单位 
            BackStyle       =   0  'Transparent
            Caption         =   "单位"
            Height          =   180
            Left            =   2505
            TabIndex        =   44
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl总量 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总量"
            Height          =   180
            Left            =   540
            TabIndex        =   43
            Top             =   1935
            Width           =   360
         End
         Begin VB.Label lbl医生嘱托 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医生嘱托"
            Height          =   180
            Left            =   5490
            TabIndex        =   50
            Top             =   615
            Width           =   720
         End
         Begin VB.Label lbl执行性质 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行性质"
            Height          =   180
            Left            =   8250
            TabIndex        =   49
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label lbl执行科室 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行科室"
            Height          =   180
            Left            =   5490
            TabIndex        =   48
            Top             =   1260
            Width           =   720
         End
         Begin VB.Label lbl用法 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "用法"
            Height          =   180
            Left            =   540
            TabIndex        =   45
            Top             =   1575
            Width           =   360
         End
         Begin VB.Label lbl开始时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始时间"
            Height          =   180
            Left            =   180
            TabIndex        =   40
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lbl执行时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行时间"
            Height          =   180
            Left            =   5490
            TabIndex        =   47
            Top             =   937
            Width           =   720
         End
      End
      Begin MSComctlLib.StatusBar stbThis 
         Height          =   360
         Left            =   0
         TabIndex        =   57
         Top             =   8025
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   9
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Bevel           =   2
               Object.Width           =   2355
               MinWidth        =   882
               Picture         =   "frmOutAdviceEdit.frx":A27F
               Text            =   "中联软件"
               TextSave        =   "中联软件"
               Key             =   "ZLFLAG"
               Object.ToolTipText     =   "欢迎使用中联有限公司软件"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   12912
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   318
               MinWidth        =   2
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   970
               MinWidth        =   970
               Picture         =   "frmOutAdviceEdit.frx":AB13
               Key             =   "KB"
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   617
               MinWidth        =   617
               Picture         =   "frmOutAdviceEdit.frx":B879
               Key             =   "PY"
               Object.ToolTipText     =   "拼音(F7)"
            EndProperty
            BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   617
               MinWidth        =   617
               Picture         =   "frmOutAdviceEdit.frx":BEB3
               Key             =   "WB"
               Object.ToolTipText     =   "五笔(F7)"
            EndProperty
            BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Bevel           =   0
               Object.Width           =   953
               MinWidth        =   25
               Text            =   "计价"
               TextSave        =   "计价"
               Key             =   "Price"
               Object.ToolTipText     =   "显示诊疗计价面板(F8)"
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
      Begin VB.Image imgButtonDel 
         Height          =   240
         Left            =   1680
         Picture         =   "frmOutAdviceEdit.frx":C4ED
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgButtonNew 
         Height          =   240
         Left            =   1005
         Picture         =   "frmOutAdviceEdit.frx":12D3F
         Top             =   105
         Visible         =   0   'False
         Width           =   240
      End
      Begin XtremeCommandBars.CommandBars cbsMain 
         Left            =   75
         Top             =   75
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
End
Attribute VB_Name = "frmOutAdviceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event FormUnload(Cancel As Integer)
Public Event EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean) '编辑门诊诊断
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断Id As String, ByRef blnNo As Boolean) '根据诊断检查是否书写传染病报告卡

Public mblnOK As Boolean
'入口参数
Private mint场合 As Integer '调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Private mblnModal As Boolean
Private mfrmParent As Object
Private mMainPrivs As String
Private mlng病人ID As Long
Private mstr挂号单 As String '病人挂号单据号
Private mlng挂号ID As Long
Private mlng合同单位ID As Long
Private mlng前提ID As Long '医技工作站下医嘱时用
Private mstr前提IDs As String
Private mlng医技科室ID As Long
Private mint婴儿 As Integer '修改时用
Private mlng医嘱ID As Long '修改时用
Private mblnCancle As Boolean   '保存时验证

'程序变量
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset '医嘱内容定义
Private mrsDrugScale As ADODB.Recordset '常用剂量比例
Private mrsPrice As ADODB.Recordset '医嘱对应的收费项目信息集

Private WithEvents mfrmSend As frmOutAdviceSend
Attribute mfrmSend.VB_VarHelpID = -1
Private WithEvents mfrmShortCut As frmClinicShortCut
Attribute mfrmShortCut.VB_VarHelpID = -1
Private WithEvents mfrmPrice As frmAdvicePrice
Attribute mfrmPrice.VB_VarHelpID = -1
Private mobjKeyBoard As Object '屏幕键盘对象动态创建
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库房的出库检查方式
Private mstrDelIDs As String '记录需要被删除的医嘱ID
Private mstrAduitDelIDs As String '处方审查已审查的假删除
Private mstr性别 As String '用于项目输入限制判断
Private mint年龄 As Integer '病人的整数年龄
Private mDat出生日期 As Date '病人出生日期
Private mdbl门诊号 As Double   '病人门诊号 PASS =3-太元通
Private mstr姓名 As String
Private mstr身份证号 As String
Private mstr费别 As String
Private mdat挂号时间 As Date '用于相关判断
Private mlng病人科室id As Long '病人(挂号)科室ID
Private mint险类 As Integer '当前病人险类
Private mstr付款码 As String '当前病人医疗付款方式编码
Private mbln中医 As Boolean
Private mblnReturn As Boolean
Private mblnIsInHelp As Boolean
Private mrs诊断 As ADODB.Recordset
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln天数反算 As Boolean
Private msngPre天数 As Single
Private mbln复诊 As Boolean  '当前病人是否复诊
Private mstr自定义申请单IDs As String 'ID1,名称1|ID2,名称2···

Private mbln扩展页签 As Boolean '是否支持扩展页签
Private mblnNoRefresh As Boolean '不刷新界面
Private mcolSubForm As Collection '扩展页签对象集合
Private mlngID序列 As Long '假医嘱ID，用负数表示
Private mstrDel输血 As String '记录需要被删除的输血医嘱ID，仅记录主医嘱ID
Private mlng危急值ID As Long

'出口参数
Private mobjPassMap As Object  ' PASS  窗口对象映射
Private mblnPass As Boolean
'本地参数
Private mint简码 As Integer
Private mstrLike As String
Private mbln自动皮试 As Boolean
Private mbln天数 As Boolean
Private msng天数 As Single
Private mbln单量 As Boolean
Private mblnStaKB As Boolean '是否自动启用屏幕键盘
Private mbln提醒对码 As Boolean
Private mblnAutoClose As Boolean
Private mblnAddAgent As Boolean '是否要求登记毒麻药品代办人信息
Private mblnNewLIS As Boolean
Private mstrPurMed As String  '抗菌药物缺省用药目的 "1"-预防，"2"-治疗，"0"-下达时确定
Private mbytSize As Byte
Private mrsDiag  As ADODB.Recordset '西医诊断记录集
Private mblnFreeInput As Boolean

'事件状态控制变量
Private mblnRowMerge As Boolean
Private mblnNoSave As Boolean
Private mblnRunFirst As Boolean
Private mblnRowChange As Boolean
Private mblnDoCheck As Boolean
Private mbytPatiType As Byte  '病人类型，1 普通，2 急诊
'门诊疗程天数，急诊3天，普通7天
Private Const conEmergency = 3
Private Const conOrdinary = 7

'工具栏命令
Private Const conMenu_New = 100
Private Const conMenu_Insert = 101
Private Const conMenu_Delete = 102
Private Const conMenu_Merge = 104
Private Const conMenu_Copy = 105
Private Const conMenu_Scheme = 106
Private Const conMenu_Save = 107
Private Const conMenu_Sign = 108
Private Const conMenu_Reference = 109
Private Const conMenu_Help = 110
Private Const conMenu_Exit = 111
Private Const conMenu_Agent = 112
Private Const conMenu_Send = 205
Private Const conMenu_DrugScale = 300
Private Const conMenu_AdvicePay = 3006

'执行时间示例
Private Const COL_按周执行 = _
    "每周三次 1/8-3/8-5/8 或 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "表示在每周星期一的8:00,星期三的8:00,星期五的8:00这几个时间执行"
Private Const COL_按天执行 = _
    "每天三次 8-12-16 或 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "表示在每天8:00,12:00,16:00这几个时间执行" & vbCrLf & _
    "两天一次 1/8 或 1/8:00" & vbCrLf & _
        vbTab & "表示在每两天中的第1天8:00这个时间执行"
Private Const COL_按时执行 = _
    "每小时两次 1:20-1:40" & vbCrLf & _
        vbTab & "表示在每小时内的20和40分钟这两个时间执行" & vbCrLf & _
    "两小时一次 2:30 或 1:30 或 1:00" & vbCrLf & _
        vbTab & "表示在每两小时内的第2的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时的30分钟这个时间执行" & vbCrLf & _
        vbTab & "　或在每两小时内的第1的个小时这个时间执行"

'固定列
Private Const COL_F标志 = 0 '自由录入，紧急，补录,抗菌用药审核状态
'可见列索引
Private Const COL_警示 = 1 'Pass:以字符串类型处理,空表示没有审查结果
Private Const COL_诊断 = 2
Private Const COL_开始时间 = 3
Private Const col_医嘱内容 = 4
Private Const COL_总量 = 5
Private Const COL_总量单位 = 6
Private Const COL_单量 = 7
Private Const COL_单量单位 = 8
Private Const COL_天数 = 9
Private Const COL_频率 = 10
Private Const COL_用法 = 11
Private Const COL_医生嘱托 = 12 'Data用于存放摘要(医保)
Private Const COL_执行时间 = 13
Private Const COL_开嘱医生 = 14
Private Const COL_超量说明 = 15  '药品超量说明
Private Const COL_基本药物 = 16

'隐藏列索引
Private Const COL_EDIT = 17 '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号,它的Data值=新下的成套方案ID
Private Const COL_相关ID = COL_EDIT + 1
Private Const COL_婴儿 = COL_EDIT + 2
Private Const COL_序号 = COL_EDIT + 3  'Pass:Data值用于记录是否更改了审核结果
Private Const COL_状态 = COL_EDIT + 4  '病人医嘱记录.状态，Data值记录该医嘱是否已进行了医保管控检查
Private Const COL_类别 = COL_EDIT + 5  '门诊没有自由录入医嘱(*)，AdviceSet成套项目中Nvl(类别,'*')只是为了预防万一
Private Const COL_诊疗项目ID = COL_EDIT + 6
Private Const COL_名称 = COL_EDIT + 7
Private Const COL_标本部位 = COL_EDIT + 8
    Private Const COL_手术时间 = COL_EDIT + 8
    Private Const COL_输血时间 = COL_EDIT + 8
Private Const COL_检查方法 = COL_EDIT + 9 '当 类别=K 输血医嘱时该字段允许存入"1"表示备血医嘱
    Private Const COL_中药形态 = COL_EDIT + 9 '0=散装，1=中药饮片，2=免煎剂
Private Const COL_执行标记 = COL_EDIT + 10
Private Const COL_收费细目ID = COL_EDIT + 11
Private Const COL_频率次数 = COL_EDIT + 12
Private Const COL_频率间隔 = COL_EDIT + 13
Private Const COL_间隔单位 = COL_EDIT + 14
Private Const COL_计价性质 = COL_EDIT + 15
Private Const COL_执行科室ID = COL_EDIT + 16
Private Const COL_执行性质 = COL_EDIT + 17 '病人医嘱记录.执行性质=诊疗项目目录.执行科室
Private Const COL_开嘱科室ID = COL_EDIT + 18
Private Const COL_开嘱时间 = COL_EDIT + 19
Private Const COL_标志 = COL_EDIT + 20     '0-普通,1-紧急，2-补录

Private Const COL_计算方式 = COL_标志 + 1 '诊疗项目目录.计算方式
Private Const COL_频率性质 = COL_标志 + 2 '诊疗项目目录.执行频率
Private Const COL_操作类型 = COL_标志 + 3 '诊疗项目目录.操作类型
Private Const COL_执行分类 = COL_标志 + 4 '诊疗项目目录.执行分类
Private Const COL_库存 = COL_标志 + 5 '按门诊包装存放的可用库存
Private Const COL_可否分零 = COL_标志 + 6 '卫材用于存放是否跟踪在用
    Private Const COL_跟踪在用 = COL_标志 + 6
Private Const COL_剂量系数 = COL_标志 + 7
Private Const COL_门诊单位 = COL_标志 + 8
Private Const COL_门诊包装 = COL_标志 + 9
Private Const COL_处方限量 = COL_标志 + 10 '非药诊疗项目为录入限量
Private Const COL_处方职务 = COL_标志 + 11
Private Const COL_毒理分类 = COL_标志 + 12
Private Const COL_药品剂型 = COL_标志 + 13
Private Const COL_单价 = COL_标志 + 14
Private Const COL_签名否 = COL_标志 + 15
Private Const COL_附项 = COL_标志 + 16
Private Const COL_零费记帐 = COL_标志 + 17
Private Const COL_抗菌等级 = COL_标志 + 18 '抗菌药物等级:0-非抗菌药,1-非限制级,2-限制级,3-特殊使用级
Private Const COL_用药目的 = COL_标志 + 19
Private Const COL_用药理由 = COL_标志 + 20
Private Const COL_审核状态 = COL_标志 + 21 '抗菌药物审核状态：Null-无需审核，1-待审核，2-审核通过，3-审核未通过
Private Const COL_免试 = COL_标志 + 22    '是否免试   1-免试，0-不免试
Private Const COL_是否超量 = COL_标志 + 23   '药品是否超量
Private Const COL_是否超期 = COL_标志 + 24
Private Const COL_配方ID = COL_标志 + 25
Private Const COL_临床自管药 = COL_标志 + 26
Private Const COL_高危药品 = COL_标志 + 27
Private Const COL_组合项目ID = COL_标志 + 28
Private Const COL_是否停用 = COL_标志 + 29 '=1标识已停用，=0或NULL标识未停用
Private Const COL_是否溶媒 = COL_标志 + 30 '=1溶媒
Private Const COL_申请序号 = COL_标志 + 31 '
Private Const COL_单独应用 = COL_标志 + 32
Private Const COL_处方审查状态 = COL_标志 + 33
Private Const COL_处方审查结果 = COL_标志 + 34
Private Const COL_处方序号 = COL_标志 + 35    '合理用药监测用
Private Const COL_危急值ID = COL_标志 + 36

Private Const M_LNG_DIAGCOUNT = 10 '最大诊断数

Private Type AGENT_INFO
    代办人姓名      As String
    代办人身份证号  As String
    本次就诊已录入  As Boolean
End Type
Private AgentInfo As AGENT_INFO

Private Enum COL_ENUM_诊断
    col标志 = 0
    col中医 = 1
    COL西医 = 2
    col编码 = 3
    col诊断 = 4
    col中医证候 = 5
    col发病时间 = 6
    col疑诊 = 7
    col增加 = 8
    col诊断ID = 9
    col疾病ID = 10
    col证候ID = 11
    colICD码 = 12
    col医嘱ID = 13
    COLDEL = 14
    col诊断编码 = 15
    col疾病编码 = 16
    col疾病类别 = 17
    col疾病附码 = 18
    col证候编码 = 19
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal int场合 As Integer, ByVal MainPrivs As String, ByVal lng病人ID As Long, ByVal str挂号单 As String, _
    Optional ByVal lng前提ID As Long, Optional ByVal int婴儿 As Integer, Optional ByVal lng医嘱ID As Long, Optional ByVal blnModal As Boolean, _
     Optional ByVal lng界面科室ID As Long, Optional ByVal str前提IDs As String, Optional ByRef objMip As Object, Optional ByVal lng危急值ID As Long) As Boolean
    
    Set mfrmParent = frmParent
    mint场合 = int场合
    mblnModal = blnModal
    mMainPrivs = MainPrivs
    mlng病人ID = lng病人ID
    mstr挂号单 = str挂号单
    mlng前提ID = lng前提ID
    mstr前提IDs = str前提IDs
    mlng医技科室ID = IIF(mlng前提ID <> 0, lng界面科室ID, 0)
    mint婴儿 = int婴儿
    mlng医嘱ID = lng医嘱ID
    mlng危急值ID = lng危急值ID
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Private Sub InitCommandBar()
'功能：初始化工具栏
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    Dim i As Long
    
    Dim blnTwo As Boolean, lng合同单位ID As Long, bln单位记帐 As Boolean
    Dim strInsidePrivs As String
    
    strInsidePrivs = GetInsidePrivs(p门诊医嘱下达)
    lng合同单位ID = Get合同单位ID
    bln单位记帐 = Val(zlDatabase.GetPara("单位记帐", glngSys, p门诊医嘱下达)) <> 0
    
    blnTwo = Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) <> 2 Or _
             Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 2 And _
             (InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Or _
            InStr(GetInsidePrivs(p门诊医嘱下达), "发送为收费单") = 0 Or _
            bln单位记帐 And lng合同单位ID = 0)
            
    If bln单位记帐 And lng合同单位ID = 0 Or _
        InStr(GetInsidePrivs(p门诊医嘱下达), "零费记帐") = 0 Or _
        InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Or _
        Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 0 Then
        '是否显示零费记帐
        chkZeroBilling.Visible = False
    End If
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = frmIcons.imgMain.Icons
    
    '菜单
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "药嘱审查", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        '插件扩展功能
        Call CreatePlugInOK(p门诊医嘱下达, mint场合)
        If Not gobjPlugIn Is Nothing Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能")
            objPopup.BeginGroup = True
        End If
                If Not gobjDrugExplain Is Nothing Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "查看药品说明书")
            objControl.IconId = 3205
        End If
    End With
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_New, "增加")
        Set objControl = .Add(xtpControlButton, conMenu_Insert, "插入")
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "删除")
        
        If mint场合 = 0 Then '只有门诊医生工作站调用时才有这几个按钮
            intTmp = Val(Mid(gstrOutUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrOutUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrOutUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
            

                        Get自定义申请单 1, mstr自定义申请单IDs
            If mstr自定义申请单IDs <> "" Then
                For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
                Next
            End If
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then
                If InStr(strTmp, ",") = 0 Then
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    Set objControl = .Add(xtpControlButton, lngID, strName)
                        objControl.IconId = conMenu_Edit_PacsApply
                        objControl.ToolTipText = strName
                        objControl.BeginGroup = True
                                            If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Else
                    varArr = Split(strTmp, ",")
                    For i = 0 To UBound(varArr)
                        strTmp = varArr(i)
                        strName = Split(strTmp, ":")(0)
                        lngID = Val(Split(strTmp, ":")(1))
                        
                        If i = 0 Then
                            Set objPopup = .Add(xtpControlSplitButtonPopup, lngID, strName)
                                objPopup.IconId = conMenu_Edit_PacsApply
                                objPopup.BeginGroup = True
                                objPopup.ToolTipText = strName
                                With objPopup.CommandBar.Controls
                                    Set objControl = .Add(xtpControlButton, lngID * 10# + 1, strName)
                                End With
                        Else
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                        End If
                                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    Next
                End If
            End If
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Merge, "一并给药"): objControl.BeginGroup = True
        If InStr(strInsidePrivs, ";复制医嘱;") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Copy, "复制医嘱")
        End If
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Scheme, "成套方案")
        objPopup.ToolTipText = "保存为成套方案"
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Scheme * 10# + 1, "保存为成套方案"
            .Add xtpControlButton, conMenu_Scheme * 10# + 2, "显示成套方案选择器"
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_Agent, "代办人")
        Set objControl = .Add(xtpControlButton, conMenu_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Sign, "签名")
        
        If blnTwo Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Send, "发送")
            objPopup.ToolTipText = "自动完成发送(F3)"
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Send * 10# + 1, "自动完成发送"
                .Add xtpControlButton, conMenu_Send * 10# + 2, "医嘱发送处理"
            End With
        Else
            Set objControl = .Add(xtpControlButton, conMenu_Send, "发送")
        End If
        If InStr(GetInsidePrivs(p门诊医嘱下达), ";诊间无卡支付;") > 0 And Not mblnAutoClose Then
            Set objControl = .Add(xtpControlButton, conMenu_AdvicePay, "诊间支付"): objControl.BeginGroup = True
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Reference, "参考"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help, "帮助")
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.ID = conMenu_Help Or objControl.ID = conMenu_Exit Or objControl.ID = conMenu_Reference Then
            objControl.Style = xtpButtonIcon
        Else
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
   
    
    '热键绑定:注意不能和系统的文本编辑热键冲突，以及Form_keydown中的冲突
    With cbsMain.KeyBindings
        .Add 0, vbKeyF2, conMenu_Save
        If blnTwo Then
            .Add 0, vbKeyF3, conMenu_Send * 10# + 1
            .Add FCONTROL, vbKeyG, conMenu_Send * 10# + 2
        Else
            .Add 0, vbKeyF3, conMenu_Send
        End If
        .Add 0, vbKeyF6, conMenu_Reference
        .Add 0, vbKeyF9, conMenu_Merge
        .Add 0, vbKeyF1, conMenu_Help
        .Add FCONTROL, vbKeyA, conMenu_New
        .Add FCONTROL, vbKeyI, conMenu_Insert
        .Add FCONTROL, vbKeyK, conMenu_Merge
        .Add FCONTROL, vbKeyY, conMenu_Copy
        .Add FCONTROL, vbKeyT, conMenu_Scheme
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add FALT, vbKeyX, conMenu_Exit
    End With
End Sub

Private Sub InitAdviceTable()
'功能：初始化表格内容，用在窗体个性化设置恢复之前
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant

    strHead = _
        ",240,4;,270,4;开始时间,1530,1;医嘱内容,3500,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;" & _
        "天数,450,1;频率,1200,1;用法,1200,1;医生嘱托,1000,1;执行时间;开嘱医生,850,1;超量说明,1000,1;基本药物,850,1;" & _
        "EDIT;相关ID;婴儿;序号;医嘱状态;诊疗类别;诊疗项目ID;名称;标本部位;检查方法;执行标记;收费细目ID;" & _
        "频率次数;频率间隔;间隔单位;计价性质;执行科室ID;执行性质;开嘱科室ID;开嘱时间;标志;计算方式;" & _
        "频率性质;操作类型;执行分类;库存;可否分零;剂量系数;门诊单位;门诊包装;处方限量;处方职务;" & _
        "毒理分类;药品剂型;单价;签名否;附项;零费记帐;抗菌等级;用药目的;用药理由;审核状态;免试;" & _
        "是否超量;是否超期;配方ID;临床自管药;高危药品;组合项目ID;是否停用;是否溶媒;申请序号;单独应用;处方审查状态;处方审查结果;处方序号;危急值ID"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColHidden(COL_警示) = True 'Pass
        '.FrozenCols = COL_医嘱内容 + 1 - .FixedCols
        .ColWidth(0) = 14 * Screen.TwipsPerPixelX
        
        '列头图标
        Set .Cell(flexcpPicture, 0, COL_警示) = img16.ListImages("警示").Picture
        Set .Cell(flexcpPicture, 0, COL_诊断) = img16.ListImages("诊断").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 4
    End With
End Sub

Private Sub Set用法Input(rsInput As ADODB.Recordset, ByVal int类型 As Integer)
'功能：输入给药途径或中药用法后调用
'参数：rsInput=输入或选择的返回记录
'      int类型=2-给药途径,4-中药用法
'说明：如果可选频率,则配合给药途径处理可用执行时间方案的变化
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnValid As Boolean, sng天数 As Single
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim vMsg As VbMsgBoxResult, strMsg As String
    Dim bln用法用量 As Boolean, bln用法频率 As Boolean
    Dim strIDs1 As String, strIDs2 As String, str医嘱内容 As String
    
    On Error GoTo errH
    cmd用法.Tag = rsInput!ID
    txt用法.Text = rsInput!名称
    txt用法.Tag = "1"
    
    With vsAdvice
        '非输液类清除滴速
        If int类型 = 2 Then
            If Nvl(rsInput!执行分类ID, 0) <> 1 And cbo滴速.Text <> "" Then
                cbo滴速.Text = ""
                cbo滴速.Tag = "1"
            End If
        End If
        
        '重新获取可用的缺省时间方案
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            Call Get时间方案(cbo执行时间, Get频率范围(.Row), .TextMatrix(.Row, COL_频率), rsInput!ID)
            If cbo执行时间.ListCount > 0 Then
                cbo执行时间.ListIndex = 0
                cbo执行时间.Tag = "1"
            Else
                '判断当前执行时间是否合法
                If cbo执行时间.Text <> "" Then
                    blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位))
                    If Not blnValid Then '如果不合法,则另取,否则保持
                        cbo执行时间.Text = ""
                        cbo执行时间.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '根据诊疗用法用量作缺省设置
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            strSQL = "Select 频次,小儿剂量,成人剂量,医生嘱托,疗程" & _
                " From 诊疗用法用量 Where 性质>0 And 项目ID=[1] And 用法ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_诊疗项目ID)), Val(rsInput!ID))
            If rsTmp.EOF Then
                '药品没有设置用法用量，则找给药途径的缺省频率（如果只设置了一个可用频率的话）,之前输入医嘱项目时设置的缺省频率是按编码排序取的第一个
                strSQL = "Select 频次 From 诊疗用法用量 Where 性质>0 And 项目ID=[1] Order by 频次"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!ID))
                bln用法频率 = rsTmp.RecordCount > 0
            Else
                bln用法用量 = True
            End If
            
            If bln用法用量 Or bln用法频率 Then
                If Not IsNull(rsTmp!频次) Then
                    Call Get频率信息_编码(rsTmp!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    txt频率.Text = str频率
                    cmd频率.Tag = str频率
                    txt频率.Tag = "1"
                End If
                
                '根据新的频率重新设置执行时间
                If cbo执行时间.Enabled Then
                    Call Get时间方案(cbo执行时间, Get频率范围(.Row), str频率, rsInput!ID)
                    If cbo执行时间.ListCount > 0 Then
                        cbo执行时间.ListIndex = 0
                        cbo执行时间.Tag = "1"
                    Else
                        '判断当前执行时间是否合法
                        If cbo执行时间.Text <> "" Then
                            blnValid = ExeTimeValid(cbo执行时间.Text, int频率次数, int频率间隔, str间隔单位)
                            If Not blnValid Then '如果不合法,则另取,否则保持
                                cbo执行时间.Text = ""
                                cbo执行时间.Tag = "1"
                            End If
                        End If
                    End If
                End If
                
                If bln用法用量 Then
                    '药品单量
                    If mint年龄 > 12 Then
                        If Nvl(rsTmp!成人剂量, 0) <> 0 Then
                            txt单量.Text = FormatEx(rsTmp!成人剂量, 5)
                            txt单量.Tag = "1"
                        End If
                    Else
                        If Nvl(rsTmp!小儿剂量, 0) <> 0 Then
                            txt单量.Text = FormatEx(rsTmp!小儿剂量, 5)
                            txt单量.Tag = "1"
                        ElseIf Nvl(rsTmp!成人剂量, 0) <> 0 Then
                            txt单量.Text = FormatEx(rsTmp!成人剂量 * (mint年龄 + 2) * 5 / 100, 5)
                            txt单量.Tag = "1"
                        End If
                    End If
                    
                    '取缺省的天数
                    sng天数 = msng天数
                    If mbln天数 Then
                        If str间隔单位 = "周" Then
                            sng天数 = IIF(7 > sng天数, 7, sng天数)
                        ElseIf str间隔单位 = "天" Then
                            sng天数 = IIF(int频率间隔 > sng天数, int频率间隔, sng天数)
                        ElseIf str间隔单位 = "小时" Then
                            sng天数 = IIF(int频率间隔 \ 24 > sng天数, int频率间隔 \ 24, sng天数)
                        ElseIf str间隔单位 = "分钟" Then
                            If sng天数 = 0 Then sng天数 = 1
                        End If
                        If sng天数 = 0 Then sng天数 = 1
                    End If
                    If Nvl(rsTmp!疗程, 1) > sng天数 Then
                        sng天数 = Nvl(rsTmp!疗程, 1)
                    End If
                    If Val(.TextMatrix(.Row, COL_天数)) > sng天数 Then
                        sng天数 = Val(.TextMatrix(.Row, COL_天数))
                    End If
                    If Val(.TextMatrix(.Row, COL_天数)) <> sng天数 Then
                        txt天数.Text = sng天数
                        txt天数.Tag = "1"
                    End If
                    
                    '药品临嘱总量:门诊包装
                    If str频率 <> "" And Val(txt单量.Text) <> 0 _
                        And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                        And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                        
                        txt总量.Text = FormatEx(Calc缺省药品总量( _
                            Val(txt单量.Text), sng天数, _
                            int频率次数, int频率间隔, str间隔单位, _
                            .TextMatrix(.Row, COL_执行时间), _
                            Val(.TextMatrix(.Row, COL_剂量系数)), _
                            Val(.TextMatrix(.Row, COL_门诊包装)), _
                            Val(.TextMatrix(.Row, COL_可否分零))), 5)
                        If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                            txt总量.Text = IntEx(Val(txt总量.Text))
                        ElseIf Val(.TextMatrix(.Row, COL_可否分零)) <> 0 Then
                            txt总量.Text = IntEx(Val(txt总量.Text))
                        End If
                        txt总量.Tag = "1"
                    End If
                    
                    '医生嘱托
                    If Not IsNull(rsTmp!医生嘱托) Then
                        cbo医生嘱托.Text = rsTmp!医生嘱托
                        cbo医生嘱托.Tag = "1"
                    End If
                End If
            End If
        End If
    End With
    
    '处理当前医嘱给药途径/煎法的变化
    Call AdviceChange
    
    '对保险对码进行检查
    Call GetInsureStr(strIDs1, strIDs2, str医嘱内容, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, 1, strIDs1, strIDs2, str医嘱内容)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set频率Input(rsInput As ADODB.Recordset, ByVal int范围 As Integer)
'功能：输入执行频率后调用
'参数：rsInput=输入或选择的返回记录
'      int范围=1-西医;2-中医;-1-一次性;-2-持续性
'说明：配合用法处理可用执行时间方案的变化
    Dim lng用法ID As Long, blnValid As Boolean
    Dim sng天数 As Single, dbl总量 As Double
    Dim lngRow As Long
    
    cmd频率.Tag = rsInput!名称
    txt频率.Text = rsInput!名称
    txt频率.Tag = "1"
        
    With vsAdvice
        '先设置执行时间的可用性:分钟频率切换
        lngRow = GetBaseRow(.Row)
        If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Or InStr(",5,6,7,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If Not cbo执行时间.Enabled Then SetItemEditable , , , , 1
        Else
            If cbo执行时间.Enabled Then SetItemEditable , , , , -1
        End If
        
        If cbo执行时间.Enabled Then '"可选频率"或药品时
            If rsInput!间隔单位 = "分钟" Then
                If cbo执行时间.Text <> "" Then cbo执行时间.Tag = "1"
                cbo执行时间.Text = ""
            Else
                '处理可用执行时间方案的变化
                If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    '查找给药途径对应的行
                    lng用法ID = .FindRow(CLng(.TextMatrix(.Row, COL_相关ID)), .Row + 1)
                    If lng用法ID <> -1 Then '未找到给药途径的情况,应该不可能
                        lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                    Else
                        lng用法ID = 0
                    End If
                ElseIf RowIn配方行(.Row) Then
                    '得到对应的中药用法ID
                    lng用法ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                End If
                
                Call Get时间方案(cbo执行时间, int范围, txt频率.Text, lng用法ID)
                '取新的频率的默认执行时间
                If cbo执行时间.ListCount > 0 Then
                    cbo执行时间.ListIndex = 0
                    cbo执行时间.Tag = "1"
                Else
                    '判断当前执行时间是否合法
                    If cbo执行时间.Text <> "" Then
                        blnValid = ExeTimeValid(cbo执行时间.Text, rsInput!频率次数, rsInput!频率间隔, rsInput!间隔单位)
                        If Not blnValid Then '如果不合法,则另取,否则保持
                            cbo执行时间.Text = ""
                            cbo执行时间.Tag = "1"
                        End If
                    End If
                End If
            End If
            
            '重新计算总量
            If mbln天数 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                sng天数 = Val(txt天数.Text)
                If sng天数 = 0 Then sng天数 = 1
                
                If txt频率.Text <> "" And Val(txt单量.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                    
                    txt总量.Text = FormatEx(Calc缺省药品总量( _
                        Val(txt单量.Text), sng天数, rsInput!频率次数, _
                        rsInput!频率间隔, rsInput!间隔单位, cbo执行时间.Text, _
                        Val(.TextMatrix(.Row, COL_剂量系数)), _
                        Val(.TextMatrix(.Row, COL_门诊包装)), _
                        Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                        txt总量.Text = IntEx(Val(txt总量.Text))
                    ElseIf Val(.TextMatrix(.Row, COL_可否分零)) <> 0 Then
                        txt总量.Text = IntEx(Val(txt总量.Text))
                    End If
                    txt总量.Tag = "1"
                End If
            End If
        End If
        If rsInput!间隔单位 = "分钟" Then
            If cbo执行时间.Enabled Then SetItemEditable , , , , -1
        End If
    End With
        
    '处理当前医嘱执行频率的变化
    Call AdviceChange
End Sub

Private Function GetBaseRow(ByVal lngRow As Long) As Long
'功能：由当前可见行获取主项目的行
    If RowIn配方行(lngRow) Then
        '获取中药配方第一味中药行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    ElseIf RowIn检验行(lngRow) Then
        '获取一并采样的第一个项目行
        GetBaseRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
    Else
        GetBaseRow = lngRow
    End If
End Function

Private Sub cbo滴速_Change()
    cbo滴速.Tag = "1"
End Sub

Private Sub cbo滴速_Click()
    cbo滴速.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo滴速_GotFocus()
    zlControl.TxtSelAll cbo滴速
End Sub

Private Sub cbo滴速_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo滴速_Validate(False)
    ElseIf InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbo滴速_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo滴速.Text) > 10 Then
        MsgBox "滴速输入内容过长，请检查输入是否正确。", vbInformation, gstrSysName
        Call cbo滴速_GotFocus: Cancel = True: Exit Sub
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo附加执行_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo附加执行.ListIndex = -1 Then Exit Sub
    
    If cbo附加执行.ItemData(cbo附加执行.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 IN(1,3)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
        vRect = GetControlRect(cbo附加执行.hWnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl附加执行.Caption, , , , , , True, vRect.Left, vRect.Top, txt用法.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = SeekCboIndex(cbo附加执行, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call zlControl.CboSetIndex(cbo附加执行.hWnd, intIdx)
        End If
    Else
        cbo附加执行.Tag = "1"
        lngRow = vsAdvice.Row
        
        '更新更改了的执行科室医嘱内容
       Call AdviceChange
    End If
End Sub

Private Sub cbo附加执行_GotFocus()
    Call zlControl.TxtSelAll(cbo附加执行)
End Sub

Private Sub cbo附加执行_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo附加执行.ListIndex = -1 Then
            Call cbo附加执行_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo附加执行_Validate(False)
        End If
    End If
End Sub

Private Sub cbo附加执行_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If cbo附加执行.ListIndex <> -1 Then Exit Sub '已选中
    If cbo附加执行.Text = "" Then '无输入
        With vsAdvice
            '原液皮试
            If .TextMatrix(.Row, COL_类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "1" And .TextMatrix(.Row, COL_执行分类) = "5" Then
                cbo附加执行.Tag = "1"
                Call AdviceChange
                Exit Sub
            Else
                If cbo附加执行.ListCount > 0 Then Cancel = True
                Exit Sub
            End If
        End With
    End If
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo附加执行.ListCount > 0 Then
        If cbo附加执行.ItemData(cbo附加执行.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    strInput = UCase(NeedName(cbo附加执行.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(1,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then cbo附加执行.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo附加执行.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo附加执行.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl附加执行.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo附加执行, rsTmp!ID)
            If intIdx <> -1 Then
                cbo附加执行.ListIndex = intIdx
            Else
                cbo附加执行.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo附加执行.ListCount - 1
                cbo附加执行.ItemData(cbo附加执行.NewIndex) = rsTmp!ID
                cbo附加执行.ListIndex = cbo附加执行.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo医生嘱托_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim strTmp As String, arrTmp As Variant
    Dim objControl As CommandBarControl
    Dim i As Long
    
    If CommandBar Is Nothing Then Exit Sub
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
    Case conMenu_Tool_PlugIn
        Call CreatePlugInOK(p门诊医嘱下达, mint场合)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            If mint场合 = 0 Then '医生站调用
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p门诊医嘱下达, 0)
            ElseIf mint场合 = 1 Then '护士站调用
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p门诊医嘱下达, 1)
            ElseIf mint场合 = 2 Then '医技站调用
                strTmp = gobjPlugIn.GetFuncNames(glngSys, p门诊医嘱下达, 2)
            End If
            Call zlPlugInErrH(err, "GetFuncNames")
            err.Clear: On Error GoTo 0
        End If
        If strTmp <> "" Then
            With CommandBar.Controls
                If .Count = 0 Then
                    strTmp = Replace(strTmp, "Auto:", "")
                    arrTmp = Split(strTmp, ",")
                    For i = 0 To UBound(arrTmp)
                        Set objControl = .Add(xtpControlButton, conMenu_Tool_PlugIn_Item + i + 1, CStr(arrTmp(i)))
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIF(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = arrTmp(i)
                    Next
                End If
            End With
        End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim lngColW As Long, i As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraPati.Left = lngLeft
    fraPati.Top = lngTop
    fraPati.Width = lngRight - lngLeft
    fraPati.Height = cmdAlley.Top + cmdAlley.Height + 60
    
    fra诊断.Left = lngLeft
    fra诊断.Top = fraPati.Top + fraPati.Height
    fra诊断.Width = lngRight - lngLeft
    
    opt诊断(1).Left = fra诊断.Width - opt诊断(1).Width - 10 * Screen.TwipsPerPixelX
    opt诊断(0).Left = opt诊断(1).Left
    vsDiag.Width = opt诊断(0).Left - vsDiag.Left - 100
    For i = 0 To vsDiag.Cols - 1
        If Not vsDiag.ColHidden(i) And i <> col诊断 Then
            lngColW = lngColW + vsDiag.ColWidth(i)
        End If
    Next
    vsDiag.ColWidth(col诊断) = vsDiag.Width - lngColW - 2 * Screen.TwipsPerPixelX
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = fraPati.Top + fraPati.Height + fra诊断.Height
    vsAdvice.Height = lngBottom - lngTop - fraPati.Height - fra诊断.Height - (fraAdvice.Height - 80) - IIF(pic疑问.Visible, pic疑问.Height, 0)
    vsAdvice.Width = lngRight - lngLeft
    
    pic疑问.Left = 0
    pic疑问.Top = vsAdvice.Top + vsAdvice.Height
    pic疑问.Width = vsAdvice.Width
    lbl疑问.Width = pic疑问.ScaleWidth - lbl疑问.Left - 45
    
    fraAdvice.Left = lngLeft
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 6 * Screen.TwipsPerPixelX + IIF(pic疑问.Visible, pic疑问.Height, 0)
    fraAdvice.Width = lngRight - lngLeft
    
    stbThis.Top = lngBottom - 10
    stbThis.Width = picSub.ScaleWidth
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim vRect As RECT, vPos As PointAPI
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    
    'PASS二级菜单的可见性与可用性在独立部件中设置，此时需要屏蔽，如不屏蔽，部件中设置值将会被覆盖
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_Delete
            With vsAdvice
                blnEnabled = True
                If .RowData(.Row) <> 0 Then
                    If Not fraAdvice.Enabled And Val(.TextMatrix(.Row, COL_审核状态)) <> 2 Then blnEnabled = False
                    If Not fraAdvice.Enabled And Val(.TextMatrix(.Row, COL_审核状态)) = 2 And .TextMatrix(.Row, COL_类别) = "K" And gbln血库系统 Then blnEnabled = False
                    If .TextMatrix(.Row, COL_类别) = "K" And Val(.TextMatrix(.Row, COL_检查方法)) = 1 And Val(.TextMatrix(.Row, COL_审核状态)) = 1 Then blnEnabled = False
                    If Val(.TextMatrix(.Row, COL_状态)) <> 1 Then blnEnabled = False
                    '已签名医嘱不可删除
                    If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then blnEnabled = False
                End If
                Control.Enabled = blnEnabled
            End With
        Case conMenu_Merge
            Control.Checked = mblnRowMerge
        Case conMenu_Scheme, conMenu_Scheme * 10# + 1, conMenu_Scheme * 10# + 2
                        If mblnModal Then
                Control.Visible = False
            Else
            If InStr(GetInsidePrivs(p门诊医嘱下达), "保存成套方案") = 0 Then
                Control.Visible = False
            End If
            If Control.ID = conMenu_Scheme * 10# + 2 Then
                Control.Caption = IIF(mfrmShortCut.Visible, "隐藏", "显示") & "成套方案选择器"
            End If
                        End If
        Case conMenu_Agent
            If Not mblnAddAgent Then
                Control.Visible = False
            Else
                strTmp = GetTsPrivs(p门诊医嘱下达)
                Control.Visible = InStr(strTmp, "下达麻醉药嘱") > 0 Or InStr(strTmp, "下达毒性药嘱") > 0 Or InStr(strTmp, "下达精神药嘱") > 0
            End If
        Case conMenu_Save
            Control.Enabled = mblnNoSave
        Case conMenu_Sign
            If mint场合 = 1 Or InStr(UserInfo.性质, "医生") = 0 Or gobjESign Is Nothing _
                Or InStr(GetInsidePrivs(p门诊医嘱下达), ";医嘱下达;") = 0 Then
                Control.Visible = False
            ElseIf mint场合 = 0 And Control.Category <> "已判断" Then
                If CheckSign(0, 0, mlng医技科室ID, mlng病人科室id, 1, False, gobjESign) = False Then
                    Control.Visible = False '不同场合没有设置要使用签名
                End If
                Control.Category = "已判断"
            ElseIf mint场合 = 2 And Control.Category <> "已判断" Then
                If CheckSign(3, 0, mlng医技科室ID, mlng病人科室id, 1, False, gobjESign) = False Then
                    Control.Visible = False '不同场合没有设置要使用签名
                End If
                Control.Category = "已判断"
            End If
        Case conMenu_Send, conMenu_Send * 10# + 1, conMenu_Send * 10# + 2
            If InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱发送") = 0 Then
                Control.Visible = False
            ElseIf InStr(GetInsidePrivs(p门诊医嘱下达), "发送为收费单") = 0 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Then
                Control.Visible = False
            End If
            If Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 0 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为收费单") = 0 Or _
               Val(zlDatabase.GetPara("发送单据类型", glngSys, p门诊医嘱下达)) = 1 And InStr(GetInsidePrivs(p门诊医嘱下达), "发送为记帐单") = 0 Then
                Control.Visible = False
            End If
                Case conMenu_Edit_ViewDrugExplain '查看药品说明书
                Control.Enabled = vsAdvice.RowData(vsAdvice.Row) <> 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0
        Case conMenu_Reference
            If GetInsidePrivs(p药品诊疗参考) = "" Then
                Control.Visible = False
            End If
        Case Else
            '单量显示状态
            blnEnabled = False
            If txt单量.Enabled Then
                If InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 Then
                    GetCursorPos vPos
                    GetWindowRect fraAdvice.hWnd, vRect
                    If Between(vPos.X * Screen.TwipsPerPixelX, vRect.Left * Screen.TwipsPerPixelX + lbl单量.Left, vRect.Left * Screen.TwipsPerPixelX + lbl单量.Left + lbl单量.Width) Then
                        If Between(vPos.Y * Screen.TwipsPerPixelY, vRect.Top * Screen.TwipsPerPixelY + lbl单量.Top, vRect.Top * Screen.TwipsPerPixelY + lbl单量.Top + lbl单量.Height) Then
                            blnEnabled = True
                        End If
                    End If
                End If
            End If
            Call Set单量Face(blnEnabled)
    End Select
End Sub

Private Sub chk免试_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk免试.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub chk免试_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdAlley_Click()
'功能：对病人过敏史/病生状态进行管理
    'Pass
    If mblnPass Then
        Call gobjPass.zlPassCmdAlleyManage(mobjPassMap)
    End If
End Sub

Private Sub cmdLastDiag_Click()
'功能：读取上次就诊的诊断信息，追加到现有诊断后
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, strTmp As String
    
    If MsgBox("是否读取上次诊断信息添加到现有诊断之后？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        On Error GoTo errH
        strSQL = "Select A.ID,A.记录来源,A.诊断类型,A.疾病ID,A.诊断ID,A.证候ID,A.诊断描述,A.是否疑诊,c.编码 as ICD码,D.编码 as 诊断编码,E.编码 as 证候编码,A.发病时间" & vbNewLine & _
            "From 病人诊断记录 A, 病人挂号记录 B,疾病编码目录 C, 疾病诊断目录 D,疾病编码目录 E" & vbNewLine & _
            "Where a.病人id = b.病人id And a.主页id = b.Id And A.疾病ID=C.ID(+)  And a.诊断id = D.Id(+) And  a.证候ID=E.ID(+) And a.病人id = [1] And (诊断类型 = 1 Or 诊断类型 = 11) And" & vbNewLine & _
            "   (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And b.登记时间 =" & vbNewLine & _
            "      (Select Max(登记时间)" & vbNewLine & _
            "       From 病人挂号记录 C" & vbNewLine & _
            "       Where c.病人id = [1] And ID <> [2] And Exists (Select 1 From 病人诊断记录 D Where d.病人id = [1] And 主页id = c.Id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
        If rsTmp.RecordCount = 0 Then
            MsgBox "该病人在本次就诊之前未填写过门诊诊断或者填写过的诊断已停用。", vbInformation
        Else
            With vsDiag
                Do While Not rsTmp.EOF
                    i = -1
                    If Val(rsTmp!诊断id & "") <> 0 Then i = .FindRow(Val(rsTmp!诊断id & ""), , col诊断ID)
                    If i = -1 And Val(rsTmp!疾病id & "") <> 0 Then i = .FindRow(Val(rsTmp!疾病id & ""), , col疾病ID)
                    '如果当前已经填写，则只更新发病时间
                    If i <> -1 Then
                        If .TextMatrix(i, col发病时间) = "" Then
                            .TextMatrix(i, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                        End If
                    Else
                        If Not (.TextMatrix(1, col诊断) = "" And .Rows = 2) Then .AddItem ""
                        i = .Rows - 1
                        Call SetDiagType(i, rsTmp!诊断类型)
                        
                        If IsNull(rsTmp!诊断描述) Then
                            .TextMatrix(i, col编码) = ""
                            .TextMatrix(i, col诊断) = ""
                        Else
                            If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                                '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                                If Val(rsTmp!疾病id & "") <> 0 Then
                                    .TextMatrix(i, col编码) = Nvl(rsTmp!ICD码)
                                ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                    .TextMatrix(i, col编码) = Nvl(rsTmp!诊断编码)
                                Else
                                    .TextMatrix(i, col编码) = ""
                                End If
                                .TextMatrix(i, col诊断) = rsTmp!诊断描述
                            Else
                                .TextMatrix(i, col编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                                .TextMatrix(i, col诊断) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                            End If
                        End If

                        .Cell(flexcpData, i, col疑诊) = Val(Nvl(rsTmp!是否疑诊, 0))
                        .Cell(flexcpForeColor, i, col疑诊) = IIF(Nvl(rsTmp!是否疑诊, 0) = 1, vbRed, .GridColor)
                        
                        .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断id, 0)
                        .Cell(flexcpData, i, col诊断ID) = Nvl(rsTmp!ID, 0)
                        .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                        .TextMatrix(i, col证候ID) = Nvl(rsTmp!证候id, 0)
                        .TextMatrix(i, colICD码) = Nvl(rsTmp!ICD码)
                        .TextMatrix(i, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                        '取证候名称
                        If InStr(.TextMatrix(i, col诊断), "(") > 0 And InStr(.TextMatrix(i, col诊断), ")") > 0 Then
                            strTmp = Mid(.TextMatrix(i, col诊断), InStrRev(.TextMatrix(i, col诊断), "(") + 1)
                            strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                            '先取证候
                            .TextMatrix(i, col中医证候) = strTmp
                            '去掉诊断描述的证候
                            .TextMatrix(i, col诊断) = Mid(.TextMatrix(i, col诊断), 1, InStrRev(.TextMatrix(i, col诊断), "(") - 1)
                        Else
                           .TextMatrix(i, col中医证候) = ""
                        End If
                        '自由录入诊断的诊断描述，需要去掉证候，因此此句代码后移
                        If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                            .Cell(flexcpData, i, col诊断) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                        Else
                            .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                mblnNoSave = True: lbl诊断.Tag = "1"
                Call SetDiagHeight
            End With
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdReason_Click()
    Call ReasonSelect("", 1)
    Call AdviceChange
End Sub

Private Sub cmdExcReason_Click()
    Call ReasonSelect("", 3)
    Call AdviceChange
End Sub

Private Function ReasonSelect(ByVal strFind As String, ByVal intType As Integer) As Boolean
'常用嘱托和抗菌用药理由超量说明选择器
'intType  1-抗菌用药理由，2-常用嘱托，3-超量说明
    Dim blnCancle As Boolean
    Dim strRetrun As String
    Dim lngLeft As Long, lngTop As Long
    Dim strName As String
    
    If intType = 1 Then
        lngLeft = txt用药理由.Left
        lngTop = txt用药理由.Top
        strName = "抗菌用药理由。"
    ElseIf intType = 2 Then
        lngLeft = cbo医生嘱托.Left
        lngTop = cbo医生嘱托.Top
        strName = "常用嘱托。"
    ElseIf intType = 3 Then
        lngLeft = txt超量说明.Left
        lngTop = txt超量说明.Top
        strName = "超量说明。"
    End If
    
    lngLeft = lngLeft + fraAdvice.Left + Me.Left
    lngTop = lngTop + fraAdvice.Top + Me.Top - 2600
    
    strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, intType)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "没有找到可用的" & strName, vbInformation, Me.Caption
            End If
        Else
            If intType = 1 Then
                txt用药理由.Text = strRetrun
            ElseIf intType = 2 Then
                cbo医生嘱托.Text = strRetrun
            ElseIf intType = 3 Then
                txt超量说明.Text = strRetrun
            End If
        End If
    End If
    ReasonSelect = blnCancle
End Function

Private Sub ReasonSave(ByVal intType As Integer)
'功能：抗菌用药理由和超量说明保存
'参数：intType  0-抗菌用药理由，1-超量说明
    Dim strSQL As String, rsTmp As Recordset
    Dim strTmp As String
    Dim strPar As String
    
    If txt用药理由.Text = "" And intType = 0 Then MsgBox "请输入您需要收藏的用药理由。", vbInformation, Me.Caption: txt用药理由.SetFocus: Exit Sub
    If txt超量说明.Text = "" And intType = 1 Then MsgBox "请输入您需要收藏的超量说明。", vbInformation, Me.Caption: txt超量说明.SetFocus: Exit Sub
    
    If intType = 0 Then
        strPar = txt用药理由.Text
        strTmp = "用药理由"
    ElseIf intType = 1 Then
        strPar = txt超量说明.Text
        strTmp = "超量说明"
    End If
    
    On Error GoTo errH
    strSQL = "Select 1 From 医嘱常用原因 Where 名称=[1]"
    If intType = 1 Then strSQL = strSQL & " And 性质=1 And 人员=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPar, UserInfo.姓名)
    '如果已经有了，提示用户是否继续。
    If rsTmp.RecordCount > 0 Then
        MsgBox "已经存在相同的" & strTmp & "。", vbInformation, Me.Caption
        Exit Sub
    End If
    strSQL = "zl_医嘱常用原因_Update(0,Null,'" & strPar & "',Null" & _
        IIF(intType = 1, ",1,'" & UserInfo.姓名 & "'", "") & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    MsgBox strTmp & "收藏成功。", vbInformation, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd常用嘱托_Click()
    Dim strSQL As String, i As Integer
    Dim rsTmp As Recordset
    
    If Trim(cbo医生嘱托.Text) = "" Then
        MsgBox "请输入嘱托内容。", vbInformation, gstrSysName
        If cbo医生嘱托.Enabled Then cbo医生嘱托.SetFocus
        Exit Sub
    End If
    On Error GoTo errH
    strSQL = "Select 1 From 常用嘱托 Where 名称=[1] And (人员=[2] Or 人员 is null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(cbo医生嘱托.Text), UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        MsgBox "该嘱托内容已经在常用嘱托中。", vbInformation, gstrSysName
        If cbo医生嘱托.Enabled Then cbo医生嘱托.SetFocus
        Exit Sub
    End If
    
    strSQL = zlCommFun.zlGetSymbol(cbo医生嘱托.Text, CByte(mint简码))
    strSQL = "zl_常用嘱托_Insert('" & Replace(cbo医生嘱托.Text, "'", "''") & "','" & strSQL & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cbo医生嘱托.hWnd, CB_ADDSTRING, 0, cbo医生嘱托.Text
    MsgBox "已设置为常用嘱托。", vbInformation, gstrSysName
    If cbo医生嘱托.Enabled Then cbo医生嘱托.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd频率_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
    Dim strSeek As String, lng诊疗项目ID As Long, lngFind As Long
    
    On Error GoTo errH
    
    With vsAdvice
        int范围 = Get频率范围(.Row)
        
        If txt频率.Text <> "" Then
            strSQL = "Select 编码 From 诊疗频率项目 Where 名称=[1] And 适用范围=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt频率.Text, int范围)
            If Not rsTmp.EOF Then strSeek = rsTmp!编码
        End If
        
        '可选择频率的常用频率
        lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
        If RowIn检验行(.Row) Then
            lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_相关ID)
            If lngFind <> -1 Then
                lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
            End If
        ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            lngFind = .FindRow(CLng(.TextMatrix(.Row, COL_相关ID)), .Row + 1)
            If lngFind <> -1 Then
                lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
            End If
        End If
        strSQL = ""
        If int范围 = 1 Then
            strSQL = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[2] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
        End If
        strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
            " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位" & _
            " From 诊疗频率项目 A Where A.适用范围=[1]" & strSQL & _
            " Order by A.编码"
        vRect = GetControlRect(txt频率.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗频率", False, strSeek, "", False, False, True, _
            vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, int范围, lng诊疗项目ID)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的诊疗频率项目，请先到医嘱频率管理中设置。", vbInformation, gstrSysName
            End If
            txt频率.Text = .TextMatrix(.Row, COL_频率)
            Call zlControl.TxtSelAll(txt频率)
            txt频率.SetFocus: Exit Sub
        End If
        Call Set频率Input(rsTmp, int范围)
        txt频率.SetFocus
        Call SeekNextControl
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd安排时间_Click()
    If IsDate(txt安排时间.Text) Then
        dtpDate.value = CDate(txt安排时间.Text)
    ElseIf IsDate(txt开始时间.Text) Then
        dtpDate.value = CDate(txt开始时间.Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "安排时间"
    dtpDate.Left = txt安排时间.Left + fraAdvice.Left
    dtpDate.Top = txt安排时间.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub cmd收藏用药理由_Click()
    Call ReasonSave(0)
End Sub

Private Sub cmdComExcReason_Click()
    Call ReasonSave(1)
End Sub

Private Sub cmd医生嘱托_Click()
    If ReasonSelect("", 2) Then Exit Sub
    cbo医生嘱托.Tag = "1"
    Call AdviceChange
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnNoSave Then
        If MsgBox("当前医嘱内容编辑后尚未保存，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    If Not mfrmShortCut Is Nothing Then mfrmShortCut.SaveShowState '系统自动卸载该子窗体
End Sub

Private Sub Set单量Face(ByVal blnOver As Boolean)
    If blnOver Then
        If lbl单量.BorderStyle = 0 Then
            lbl单量.BorderStyle = 1
            lbl单量.BackStyle = 1
        End If
    Else
        If lbl单量.BorderStyle = 1 Then
            lbl单量.BorderStyle = 0
            lbl单量.BackStyle = 0
        End If
    End If
End Sub

Private Sub lbl单量_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSQL As String
    Dim str单位 As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    If Not (InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 And txt单量.Enabled) Then Exit Sub
    
    If mrsDrugScale Is Nothing Then
        strSQL = "Select 名称,比例 From 常用剂量比例 Where 名称 is Not NULL And 比例 is Not NULL Order by 编码"
        Set mrsDrugScale = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If mrsDrugScale.EOF Then
            MsgBox "没有设置常用剂量比例，请先到字典管理工具进行设置。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_收费细目ID) <> 0 Then
        Set rsTmp = Get收费项目记录(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_收费细目ID)))
        str单位 = rsTmp!计算单位 & ""
    End If
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        mrsDrugScale.MoveFirst
        Do While Not mrsDrugScale.EOF
            Set objControl = .Add(xtpControlButton, conMenu_DrugScale * 100# + .Count + 1, mrsDrugScale!名称 & "[" & str单位 & "]")
            objControl.Parameter = mrsDrugScale!比例
            mrsDrugScale.MoveNext
        Loop
    End With
    GetWindowRect fraAdvice.hWnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lbl单量.Left + lbl单量.Width, vRect.Top * Screen.TwipsPerPixelY + lbl单量.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lbl滴速_Click()
    Call Load输液滴速(cbo滴速, lbl滴速单位, True)
    cbo滴速.Tag = "1"
    Call AdviceChange
End Sub

Private Sub mfrmPrice_PanelHide()
    Call stbThis_PanelClick(stbThis.Panels("Price"))
End Sub

Private Sub mfrmSend_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 挂号单, Succeed)
End Sub

Private Sub mfrmShortCut_ItemClick(ByVal 类型 As Integer, ByVal 分类ID As Long)
    If cmdSel.Enabled And cmdSel.Visible Then
        Call ClinicSelecter(类型, 分类ID)
    End If
End Sub

Private Sub picHelp_Click()
    Dim strTip As String
    
    On Error Resume Next
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "周" Then
        strTip = COL_按周执行
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "天" Then
        strTip = COL_按天执行
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "小时" Then
        strTip = COL_按时执行
    End If
    MsgBox strTip, vbInformation, Me.Caption
    cbo执行时间.SetFocus
    mblnIsInHelp = False
End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    On Error Resume Next
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "周" Then
        strTip = COL_按周执行
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "天" Then
        strTip = COL_按天执行
    ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "小时" Then
        strTip = COL_按时执行
    End If
    
    zlCommFun.ShowTipInfo picHelp.hWnd, strTip, True
    
    If X >= 0 And X <= picHelp.Width And Y >= 0 And Y <= picHelp.Height Then
        mblnIsInHelp = True
        SetCapture picHelp.hWnd
    Else
        mblnIsInHelp = False
        ReleaseCapture
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Price" Then
        If Panel.Bevel <> sbrNoBevel Then
            Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Panel.Tag = IIF(Panel.Bevel = sbrInset, "1", "")
            Call ShowPrice(vsAdvice.Row)
        End If
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("简码方式", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint简码 = IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
    ElseIf Panel.Key = "KB" Then
        On Error Resume Next
        If mobjKeyBoard Is Nothing Then Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        Call mobjKeyBoard.StartUp
        Call mobjKeyBoard.SetPos
        err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub txt超量说明_GotFocus()
    Call zlControl.TxtSelAll(txt超量说明)
End Sub

Private Sub txt超量说明_Change()
    txt超量说明.Tag = "1"
End Sub

Private Sub txt超量说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt超量说明.Text <> "" Then
            If ReasonSelect(txt超量说明.Text, 3) Then Exit Sub
        End If
        If SeekNextControl Then Call txt超量说明_Validate(False)
    End If
End Sub

Private Sub txt超量说明_Validate(Cancel As Boolean)
    Call AdviceChange
End Sub

Private Sub txt单量_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = 192 Then '~
        Call lbl单量_MouseDown(1, 0, 0, 0)
    End If
End Sub

Private Sub txt单量_LostFocus()
    mblnReturn = False
End Sub

Private Sub txt频率_GotFocus()
    Call zlControl.TxtSelAll(txt频率)
End Sub

Private Sub txt频率_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int范围 As Integer, vRect As RECT
    Dim lng诊疗项目ID As Long, lngFind As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmd频率.Tag <> "" And txt频率.Text = .TextMatrix(.Row, COL_频率) And txt频率.Text <> "" Then
                Call SeekNextControl
            ElseIf txt频率.Text = "" Then
                If cmd频率.Enabled And cmd频率.Visible Then cmd频率_Click
            Else
                int范围 = Get频率范围(.Row)
                
                '可选择频率的常用频率
                lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                If RowIn检验行(.Row) Then
                    lngFind = .FindRow(CStr(.RowData(.Row)), .FixedRows, COL_相关ID)
                    If lngFind <> -1 Then
                        lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
                    End If
                ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(.Row, COL_相关ID)), .Row + 1)
                    If lngFind <> -1 Then
                        lng诊疗项目ID = Val(.TextMatrix(lngFind, COL_诊疗项目ID))
                    End If
                End If
                strSQL = ""
                If int范围 = 1 Then
                    strSQL = " And (Exists(Select 1 From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次=A.编码 And A.适用范围=1)" & _
                        " Or (Select Count(*) From 诊疗用法用量 Where 项目ID=[4] And 用法ID is NULL And 频次 Is Not NULL)<=1)"
                End If
                strSQL = "Select Rownum as ID,A.编码,A.名称,A.简码," & _
                    " A.英文名称,A.频率次数,A.频率间隔,A.间隔单位" & _
                    " From 诊疗频率项目 A Where A.适用范围=[3]" & strSQL & _
                    " And (A.编码 Like [1] Or Upper(A.名称) Like [2]" & _
                    " Or Upper(A.简码) Like [2] Or Upper(A.英文名称) Like [2])" & _
                    " Order by A.编码"
                vRect = GetControlRect(txt频率.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "诊疗频率", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt频率.Height, blnCancel, False, True, UCase(txt频率.Text) & "%", mstrLike & UCase(txt频率.Text) & "%", int范围, lng诊疗项目ID)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的诊疗频率项目。", vbInformation, gstrSysName
                    End If
                    txt频率.Text = .TextMatrix(.Row, COL_频率)
                    Call zlControl.TxtSelAll(txt频率)
                    txt频率.SetFocus: Exit Sub
                End If
                Call Set频率Input(rsTmp, int范围)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub txt安排时间_Change()
    txt安排时间.Tag = "1"
End Sub

Private Sub txt安排时间_GotFocus()
    If txt安排时间.Text = "" Then txt安排时间.Text = txt开始时间.Text
    zlControl.TxtSelAll txt安排时间
End Sub

Private Sub txt安排时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt安排时间.Text <> "" Then
            txt安排时间.Text = GetFullDate(txt安排时间.Text)
            If SeekNextControl Then Call txt安排时间_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt安排时间_Validate(Cancel As Boolean)
    If txt安排时间.Locked Then Exit Sub
        
    If Not IsDate(txt安排时间.Text) Then
        If txt安排时间.Text <> "" Then
            Cancel = True
            txt安排时间_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)) Then
                '恢复人为的清除缺省为开始时间
                txt安排时间.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)
            End If
        End If
    Else
        '检查时间合法性
        If Not Check安排时间(txt安排时间.Text, vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间), vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) Then
            Cancel = True
            txt安排时间_GotFocus
            Exit Sub
        End If
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt天数_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_天数)) <> Val(txt天数.Text) Then
                txt天数.Tag = "1"
            End If
        Else
            txt天数.Tag = "1"
        End If
    End With
End Sub

Private Sub txt天数_GotFocus()
    Call zlControl.TxtSelAll(txt天数)
End Sub

Private Sub txt天数_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txt天数.Text) > 0 Then
            Call txt天数_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        Else
            If Val(txt天数.Text) = 0 Then txt天数.Text = ""
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt天数_LostFocus()
    mblnReturn = False
End Sub

Private Sub txt天数_Validate(Cancel As Boolean)
    Dim sng天数 As Single, i As Long
    Dim strSame As String, strMsg As String
    Dim dbl总量 As Double
    Dim strTmpTag As String
    Dim bln计算总量 As Boolean
    
    With vsAdvice
        If Val(txt天数.Text) = 0 Then txt天数.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Val(txt天数.Text) <= 0 Then
            Cancel = True: txt天数_GotFocus: Exit Sub
        End If
        
        '天数至少需要一个频率同期的天数
        If Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 Then
            If .TextMatrix(.Row, COL_间隔单位) = "周" Then
                sng天数 = 7
            ElseIf .TextMatrix(.Row, COL_间隔单位) = "天" Then
                sng天数 = Val(.TextMatrix(.Row, COL_频率间隔))
            ElseIf .TextMatrix(.Row, COL_间隔单位) = "小时" Then
                sng天数 = Val(.TextMatrix(.Row, COL_频率间隔)) \ 24
            ElseIf .TextMatrix(.Row, COL_间隔单位) = "分钟" Then
                sng天数 = Val(.TextMatrix(.Row, COL_频率间隔)) \ (24 * 60)
            End If
            If Val(txt天数.Text) < sng天数 Then
                If MsgBox("按""" & .TextMatrix(.Row, COL_频率) & """执行时，至少需要 " & sng天数 & " 天的用药，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Cancel = True: txt天数_GotFocus: Exit Sub
                End If
            End If
        End If
        
        dbl总量 = Val(txt总量.Text)
        If mbln天数反算 Then
            '如果总量为0或者说天数变了则计算下总量
            If dbl总量 = 0 Or Val(txt天数.Text) <> msngPre天数 Then
                bln计算总量 = True
            End If
        Else
            bln计算总量 = True
        End If
        
        If bln计算总量 Then
            txt总量.Text = ReGet药品总量(dbl总量, Val(txt单量.Text), Val(txt天数.Text), .Row) '隐式调用Change事件

            Call txt总量_Validate(Cancel)
            If Cancel Then
                txt总量.Text = dbl总量
                Exit Sub
            End If
        End If
                         
        Call CheckDrugOutOfRange(.Row, Val(txt天数.Text))
        
        msng天数 = Val(txt天数.Text)

    End With
    msngPre天数 = Val(txt天数.Text)
    Call AdviceChange
    
    '成套方案批量处理
    With vsAdvice
        If CStr(.Cell(flexcpData, .Row, COL_EDIT)) <> "" Then
            strSame = CStr(.Cell(flexcpData, .Row, COL_EDIT))
            If InStr(strSame, ",") > 0 Then
                strMsg = "该次复制的所有的药品都按这个天数执行吗？"
            Else
                strMsg = "该成套方案的所有药品都按这个天数执行吗？"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                For i = .FixedRows To .Rows - 1
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        If Not (Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) _
                            Or .RowData(i) = Val(.TextMatrix(.Row, COL_相关ID)) Or i = .Row) _
                                And CStr(.Cell(flexcpData, i, COL_EDIT)) = strSame Then
                                
                            If .TextMatrix(i, COL_频率) <> "" And Val(.TextMatrix(i, COL_频率次数)) <> 0 And Val(.TextMatrix(i, COL_频率间隔)) <> 0 _
                                And Val(.TextMatrix(i, COL_单量)) <> 0 And Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_门诊包装)) <> 0 Then
                                
                                .TextMatrix(i, COL_天数) = txt天数.Text
                                .TextMatrix(i, COL_总量) = ReGet药品总量(Val(.TextMatrix(i, COL_总量)), Val(.TextMatrix(i, COL_单量)), Val(txt天数.Text), i)
                                If Val(.TextMatrix(i, COL_相关ID)) = Val(.RowData(i + 1)) Then
                                    .TextMatrix(i + 1, COL_天数) = txt天数.Text
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txt用法_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long, strWhere As String
        
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd用法.Tag) <> 0 And txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法)) And txt用法.Text <> "" Then
                Call SeekNextControl
            ElseIf txt用法.Text = "" Then
                If cmd用法.Enabled And cmd用法.Visible Then cmd用法_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                    int类型 = 2 '给药途径
                ElseIf RowIn检验行(vsAdvice.Row) Then
                    int类型 = 6 '采集方法
                ElseIf .TextMatrix(.Row, COL_类别) = "K" Then
                    If gbln血库系统 = True Then
                        If Val(.TextMatrix(.Row, COL_检查方法)) = 0 Then
                            int类型 = 9 '采集输血途径
                        Else
                            int类型 = 8 '输血途径
                            strWhere = " And nvl(A.执行分类,0)=1 "
                        End If
                    Else
                        int类型 = 8 '输血途径
                    End If
                Else
                    int类型 = 4 '中药用法
                End If
                If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
                    strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[4] And 性质>0)" & _
                        " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                            " Where A.用法ID=B.ID And B.服务对象 IN(1,3) And A.项目ID=[4] And A.性质>0)<=1)"
                End If
                
                '优化
                strLike = mstrLike
                If Len(txt用法.Text) < 2 Then strLike = ""
                
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.执行分类 as 执行分类ID" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B" & _
                    " Where A.ID=B.诊疗项目ID" & _
                    " And A.类别='E' And A.操作类型=[3] And A.服务对象 IN(1,3)" & strWhere & strSQL & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2])" & _
                    " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[6])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
                    Decode(mint简码, 0, " And B.码类 IN([5],3)", 1, " And B.码类 IN([5],3)", "") & _
                    " Order by A.编码"
                vRect = GetControlRect(txt用法.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl用法.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, UCase(txt用法.Text) & "%", _
                    strLike & UCase(txt用法.Text) & "%", CStr(int类型), Val(.TextMatrix(.Row, COL_诊疗项目ID)), mint简码 + 1, mlng病人科室id)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "未找到匹配的" & lbl用法.Caption & "。", vbInformation, gstrSysName
                    End If
                    txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                    Call zlControl.TxtSelAll(txt用法)
                    txt用法.SetFocus: Exit Sub
                End If
                
                '对一并给药的其它药品的可用给药途径进行检查
                If int类型 = 2 Then
                    Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), 1) Then
                                .Refresh
                                MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, col_医嘱内容) & """。", vbInformation, gstrSysName
                                .Refresh
                                txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                                Call zlControl.TxtSelAll(txt用法)
                                txt用法.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                If Val(cmd用法.Tag) <> Val(rsTmp!ID & "") Then .TextMatrix(.Row, COL_单价) = ""
                Call Set用法Input(rsTmp, int类型)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub cmd用法_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int类型 As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim strSeek As String
    Dim strWhere As String
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            int类型 = 2 '给药途径
            lngBegin = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))), .Row + 1)
        ElseIf RowIn检验行(vsAdvice.Row) Then
            int类型 = 6 '采集方法
            lngBegin = .Row
        ElseIf .TextMatrix(.Row, COL_类别) = "K" Then
            If gbln血库系统 = True Then
                If Val(.TextMatrix(.Row, COL_检查方法)) = 0 Then
                    int类型 = 9 '采集输血途径
                Else
                    int类型 = 8 '输血途径
                    strWhere = " And nvl(A.执行分类,0)=1 "
                End If
            Else
                int类型 = 8 '输血途径
            End If
            lngBegin = .FindRow(CStr(.RowData(.Row)), .Row + 1, COL_相关ID)
        Else
            int类型 = 4 '中药用法
            lngBegin = .Row
        End If
        If txt用法.Text <> "" And lngBegin <> -1 Then
            strSeek = GetItemField("诊疗项目目录", Val(.TextMatrix(lngBegin, COL_诊疗项目ID)), "编码")
        End If
        
        If int类型 = 2 Then '只取有效范围的给药途径(无设置或仅一个时可任选)
            strSQL = " And (A.ID IN(Select 用法ID From 诊疗用法用量 Where 项目ID=[2] And 性质>0)" & _
                " Or (Select Count(A.用法ID) From 诊疗用法用量 A,诊疗项目目录 B" & _
                    " Where A.用法ID=B.ID And B.服务对象 IN(1,3) And A.项目ID=[2] And A.性质>0)<=1)"
        End If
        strSQL = "Select Distinct A.ID,A.编码,A.名称,C.名称 as 分类,A.执行分类 as 执行分类ID" & _
            " From 诊疗项目目录 A,诊疗分类目录 C" & _
            " Where A.分类ID=C.ID(+) And A.类别='E' And A.操作类型=[1] And A.服务对象 IN(1,3)" & strWhere & strSQL & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[3])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))" & _
            " Order by A.编码"
        vRect = GetControlRect(txt用法.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl用法.Caption, False, strSeek, "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, CStr(int类型), Val(.TextMatrix(.Row, COL_诊疗项目ID)), mlng病人科室id)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "没有可用的" & lbl用法.Caption & "，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
            Call zlControl.TxtSelAll(txt用法)
            txt用法.SetFocus: Exit Sub
        End If
        
        '对一并给药的其它药品的可用给药途径进行检查
        If int类型 = 2 Then
            Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check适用用法(rsTmp!ID, Val(.TextMatrix(i, COL_诊疗项目ID)), 1) Then
                        .Refresh
                        MsgBox """" & rsTmp!名称 & """不适用于与当前药品一并给药的""" & .TextMatrix(i, col_医嘱内容) & """。", vbInformation, gstrSysName
                        .Refresh
                        txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
                        Call zlControl.TxtSelAll(txt用法)
                        txt用法.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set用法Input(rsTmp, int类型)
        txt用法.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt用法_GotFocus()
    Call zlControl.TxtSelAll(txt用法)
End Sub

Private Sub txt用法_LostFocus()
    'PASS
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            Call gobjPass.zlPassCloseDrugHint(mobjPassMap)
        End If
    End If
End Sub

Private Sub txt用法_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If Val(cmd用法.Tag) <> 0 And txt用法.Text <> IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法)) Then
            txt用法.Text = IIF(cbo滴速.Text <> "", Replace(.TextMatrix(.Row, COL_用法), cbo滴速.Text & lbl滴速单位.Caption, ""), .TextMatrix(.Row, COL_用法))
        End If
    End With
End Sub

Private Sub txt频率_Validate(Cancel As Boolean)
    With vsAdvice
        '恢复人为的清除
        If cmd频率.Tag <> "" And txt频率.Text <> .TextMatrix(.Row, COL_频率) Then
            txt频率.Text = .TextMatrix(.Row, COL_频率)
        End If
    End With
End Sub

Private Sub cbo婴儿_Click()
    If Not Visible Then Exit Sub
    If cbo婴儿.ListIndex = -1 Then Exit Sub
    
    If cbo婴儿.ListIndex = Val(cbo婴儿.Tag) Then Exit Sub
    cbo婴儿.Tag = cbo婴儿.ListIndex
    
    Call ShowAdvice
    'PASS 婴儿发生改变
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            mobjPassMap.PassPati.int婴儿 = cbo婴儿.ListIndex
        End If
    End If
    vsAdvice.SetFocus
End Sub

Private Sub cbo执行科室_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
    Dim lng配制中心 As Long, str药房IDs As String
    Dim lngBegin As Long, lngEnd As Long, blnNode As Boolean, bln入院 As Boolean, bln留观 As Boolean
        
    If cbo执行科室.ListIndex = -1 Then Exit Sub
    
    If cbo执行科室.ItemData(cbo执行科室.ListIndex) = -1 Then
        
        blnNode = True
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_类别) = "Z" Then   '入院或留观
            bln入院 = vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "2"
            bln留观 = vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "1" '可以为门诊或住院留观
            blnNode = Not (bln入院 Or bln留观)
        End If
    
        '他科执行，弹出选择执行科室
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
            " From 部门表 A,部门性质说明 B" & _
            " Where A.ID=B.部门ID And B.服务对象 IN(" & IIF(bln入院, "2", IIF(bln留观, "1,2", "1")) & ",3)" & _
            IIF(blnNode, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
            IIF(bln入院 Or bln留观, " And B.工作性质='临床'", "") & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            " Order by A.编码"
        vRect = GetControlRect(cbo执行科室.hWnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl执行科室.Caption, , , , , , True, vRect.Left, vRect.Top, cbo执行科室.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "没有科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
            '恢复成现有的科室(不引发Click)
            intIdx = SeekCboIndex(cbo执行科室, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
            Call zlControl.CboSetIndex(cbo执行科室.hWnd, intIdx)
        End If
    Else
        lngRow = vsAdvice.Row
        
        '检查一并给药的配制中心
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And RowIn一并给药(lngRow) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
                
                '当前行由普通药房或其他配制中心改为配制中心
                If Have部门性质(cbo执行科室.ItemData(cbo执行科室.ListIndex), "配制中心") Then
                    lng配制中心 = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                End If
                '当前行由配置中心或改为普通药房
                If lng配制中心 = 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                            '自备药不管它
                            If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                If Have部门性质(Val(.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                                    lng配制中心 = Val(.TextMatrix(i, COL_执行科室ID)): Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                '这两种情况所有药品都执行科室相同，检查存储设定
                If lng配制中心 <> 0 Then
                    For i = lngBegin To lngEnd
                        If i <> lngRow And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                            '自备药不管它
                            If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                str药房IDs = Get可用药房IDs(.TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), Val(.TextMatrix(i, COL_收费细目ID)), mlng病人科室id, 1)
                                If InStr("," & str药房IDs & ",", "," & cbo执行科室.ItemData(cbo执行科室.ListIndex) & ",") = 0 Then
                                    MsgBox "一并给药的药品中，""" & .TextMatrix(i, col_医嘱内容) & """在""" & NeedName(cbo执行科室.Text) & """中没有存储。", vbInformation, gstrSysName
                                    '恢复成现有的科室(不引发Click)
                                    intIdx = SeekCboIndex(cbo执行科室, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行科室ID)))
                                    Call zlControl.CboSetIndex(cbo执行科室.hWnd, intIdx)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End With
        
        cbo执行科室.Tag = "1"
        
        '更新更改了的执行科室医嘱内容
        Call AdviceChange
        
        '重新获取库存并显示：以门诊单位，中药配方不显示
        With vsAdvice
            If (.TextMatrix(lngRow, COL_类别) = "4" And Val(.TextMatrix(lngRow, COL_跟踪在用)) = 1 _
                Or InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0) And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                Call GetDrugStock(lngRow)
                If InStr(GetInsidePrivs(p门诊医嘱下达), "显示药品库存") = 0 Then
                    stbThis.Panels(3).Text = IIF(Val(.TextMatrix(lngRow, COL_库存)) > 0, "有库存", "无库存")
                Else
                    stbThis.Panels(3).Text = "库存: " & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & .TextMatrix(lngRow, COL_门诊单位)
                End If
            ElseIf RowIn配方行(lngRow) Then
                Call GetDrugStock(lngRow)
            End If
        End With
    End If
End Sub

Private Sub cbo执行科室_GotFocus()
    Call zlControl.TxtSelAll(cbo执行科室)
End Sub

Private Sub cbo执行科室_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行科室.ListIndex = -1 Then
            Call cbo执行科室_Validate(blnCancel)
            cbo执行科室.SetFocus
        Else
            If SeekNextControl Then Call cbo执行科室_Validate(False)
        End If
    End If
End Sub

Private Sub cbo执行科室_Validate(Cancel As Boolean)
'功能：根据输入的内容,自动匹配执行科室
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, strInput As String, strIDs As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim blnNode As Boolean, bln门诊留观 As Boolean, bln入院 As Boolean
    
    If cbo执行科室.ListIndex <> -1 Then Exit Sub '已选中
    If cbo执行科室.Text = "" Then '无输入
        If cbo执行科室.ListCount > 0 Then Cancel = True
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '是否可以任意或选择科室
    blnLimit = True
    If cbo执行科室.ListCount > 0 Then
        If cbo执行科室.ItemData(cbo执行科室.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    blnNode = True
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_类别) = "Z" Then
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "1" Then
            blnNode = False
            bln门诊留观 = True
        ElseIf vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "2" Then
            blnNode = False
            bln入院 = True
        End If
    End If
    
    strInput = UCase(NeedName(cbo执行科室.Text))
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(" & IIF(bln门诊留观, "1,2", IIF(bln入院, "2", "1")) & ",3)" & _
        IIF(blnNode, " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)", "") & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And (A.编码 Like [1] Or A.名称 Like [2] Or Upper(A.简码) Like [2])" & _
        " Order by A.编码"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then strIDs = strIDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        
        If strIDs <> "" Then
            strIDs = Mid(strIDs, 2)
            If InStr(strIDs, ",") = 0 Then
                intIdx = SeekCboIndex(cbo执行科室, CLng(strIDs))
                If intIdx <> -1 Then cbo执行科室.ListIndex = intIdx
            Else
                strSQL = "Select /*+ rule*/ A.ID,A.编码,A.名称,A.简码 From 部门表 A,Table(f_num2list([1])) B Where A.ID = B.Column_Value"
        
                vRect = GetControlRect(cbo执行科室.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl执行科室.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, strIDs)
                If Not rsTmp Is Nothing Then
                    intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
                    If intIdx <> -1 Then cbo执行科室.ListIndex = intIdx
                End If
            End If
        End If
        If cbo执行科室.ListIndex = -1 Then
            MsgBox "未到对应的科室。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo执行科室.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl执行科室.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt用法.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo执行科室, rsTmp!ID)
            If intIdx <> -1 Then
                cbo执行科室.ListIndex = intIdx
            Else
                cbo执行科室.AddItem rsTmp!编码 & "-" & rsTmp!名称, cbo执行科室.ListCount - 1
                cbo执行科室.ItemData(cbo执行科室.NewIndex) = rsTmp!ID
                cbo执行科室.ListIndex = cbo执行科室.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "未找到对应的科室。", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo执行时间_Change()
    cbo执行时间.Tag = "1"
End Sub

Private Sub cbo执行时间_LostFocus()
    If Not mblnIsInHelp Then picHelp.Visible = False
    mblnIsInHelp = False
End Sub

Private Sub cbo执行时间_Click()
    'cbo执行时间_Change
    '更新数据
    cbo执行时间.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cbo执行时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cbo执行时间_Validate(False)
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cbo执行时间_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cbo执行时间.Text <> "" Then
            '检查长度
            If Len(cbo执行时间.Text) > 50 Then
                MsgBox "输入内容不能超过 50 个字符。", vbInformation, gstrSysName
                Call cbo执行时间_GotFocus
                Cancel = True: Exit Sub
            End If
            '检查合法性
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cbo执行时间.Text, Val(.TextMatrix(lngRow, COL_频率次数)), Val(.TextMatrix(lngRow, COL_频率间隔)), .TextMatrix(lngRow, COL_间隔单位))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                        strTmp = COL_按周执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                        strTmp = COL_按天执行
                    ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                        strTmp = COL_按时执行
                    End If
                    MsgBox "输入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cbo执行时间_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_Click()
    cbo执行性质.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo执行性质_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo执行性质.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo执行性质.hWnd, KeyAscii)
        If lngIdx = -1 And cbo执行性质.ListCount > 0 Then lngIdx = 0
        cbo执行性质.ListIndex = lngIdx
    End If
End Sub

Private Sub chk紧急_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk紧急.Tag = "1"
    '更新数据
    Call AdviceChange
    
    If txt用药理由.Enabled And Trim(txt用药理由.Text) = "" Then
        txt用药理由.SetFocus
    End If
End Sub

Private Sub chkZeroBilling_click()
    If Not mblnDoCheck Then Exit Sub
    
    chkZeroBilling.Tag = "1"
    '更新数据
    Call AdviceChange
End Sub

Private Sub chk紧急_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdExt_Click()
'功能：修改现有医嘱的扩充内容
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strAppend As String
    Dim lngRow As Long, lngFirstRow As Long
    Dim lng诊疗项目ID As Long, lng用法ID As Long
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim strTmp As String, lngDiag As Long
    Dim lng配方ID As Long
    Dim t_Pati As TYPE_PatiInfoEx
    Dim lng项目id As Long, intType As Integer
    Dim blnOK As Boolean
    Dim str手术部位 As String
    Dim strIDs1 As String, strIDs2 As String, str医嘱内容 As String
    Dim lngAppType As Long '申请单应用
    Dim objAppPages()  As clsApplicationData
    Dim rsCard As ADODB.Recordset
    Dim lngNo As Long
    Dim lngTmp As Long
    Dim str摘要 As String '医保摘要 GetItemInfo
        Dim strSQL As String, rsTmp As Recordset
    
    lngRow = vsAdvice.Row
    '读取申请附项内容：不管新录医嘱，在录入、调成套、复制时已读取
    If vsAdvice.TextMatrix(lngRow, COL_附项) = "" And Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
        If Not RowIn配方行(lngRow) Then
            vsAdvice.TextMatrix(lngRow, COL_附项) = Get病人医嘱附件(vsAdvice.RowData(vsAdvice.Row))
        End If
    End If
    strAppend = vsAdvice.TextMatrix(lngRow, COL_附项)
    
    lngNo = Val(vsAdvice.TextMatrix(lngRow, COL_申请序号) & "")
    intType = -1
    lngAppType = -1
        If lngNo <> 0 Then
        strSQL = "Select 文件ID From 医嘱申请单文件 Where 医嘱ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) = 0, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))))
        If rsTmp.RecordCount > 0 Then Call FuncApplyCustom(1, Val(rsTmp!文件ID), lngNo): Exit Sub
    End If
    If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
        If lngNo <> 0 Then
            Call GetData检查申请(lngRow, objAppPages())
            lngAppType = 0
        Else
            strExtData = Get检查部位方法(lngRow)
            If strExtData = "" Then
                MsgBox "该检查医嘱是系统升级以前下达的，与现有方式不兼容。请重新下达该检查医嘱。", vbInformation, gstrSysName
                Exit Sub
            End If
            intType = 0
        End If
    ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
        strExtData = Get手术附加IDs(lngRow)
        intType = 1
    ElseIf RowIn配方行(lngRow) Then
        strExtData = Get中药配方IDs(lngRow)
        intType = 2
    ElseIf RowIn检验行(lngRow) Then
        If lngNo <> 0 Then
            lngAppType = 3
            Call GetData检验申请(lngRow, rsCard)
        Else
            strExtData = Get检验组合IDs(lngRow)
            intType = 4
        End If
    ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "E" Or vsAdvice.TextMatrix(lngRow, COL_类别) = "K" Or vsAdvice.TextMatrix(lngRow, COL_类别) = "Z" Then
        If CanUseApply(vsAdvice.TextMatrix(lngRow, COL_类别)) Then
            Call GetData输血申请(lngRow, rsCard)
            lngAppType = 1
        Else
            intType = 5
            If vsAdvice.TextMatrix(lngRow, COL_类别) = "K" And Val(vsAdvice.TextMatrix(lngRow, COL_申请序号) & "") <> 0 Then
                Call frmBloodApply.ShowMe(Me, mlng病人ID, 0, 1, 3, Val(vsAdvice.RowData(lngRow)), mlng病人科室id, , Val(vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID)), , , mrsDefine, mclsMipModule, 1, mstr挂号单)
                Exit Sub
            End If
        End If
    Else
        Exit Sub '兼容以前的检验项目
    End If
    
    If intType = 4 Then
        lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
        lng项目id = Val(vsAdvice.TextMatrix(lngFirstRow, COL_诊疗项目ID))
    Else
        lng项目id = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
    End If
    With t_Pati
        .bln医保 = InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> ""
        .int险类 = mint险类
        .int婴儿 = mint婴儿
        .lng病人ID = mlng病人ID
        .lng病人科室ID = mlng病人科室id
        .str挂号单 = mstr挂号单
        .str性别 = mstr性别
    End With

    On Error Resume Next
    '改造接口：以前int场合传未传，现在传0，bytUseType以前未传，现在传0
    If intType = 2 Then
        blnOK = frmAdviceFormula.ShowMe(Me, gclsInsure, txt医嘱内容.hWnd, t_Pati, 0, 0, 1, 1, 1, _
                    lng项目id, strExtData, str摘要)
    ElseIf intType <> -1 Then
        blnOK = frmAdviceEditEx.ShowMe(Me, txt医嘱内容.hWnd, t_Pati, 0, intType, 0, 1, 1, 1, mblnNewLIS, False, _
                    lng项目id, strExtData, strAppend, , GetAdviceDiagnosis, str手术部位)
    End If
    '申请单
    If lngAppType = 0 Then
        blnOK = ApplyNew检查申请(1, "", objAppPages())
    ElseIf lngAppType = 1 Then
        blnOK = ApplyNew输血申请(1, "", rsCard)
    ElseIf lngAppType = 3 Then
        blnOK = ApplyNew检验申请(1, "", rsCard)
    End If
    On Error GoTo 0
    
    '重新设置相关内容
    If blnOK Then
        '获取以前的诊断关联行
        lngDiag = AdviceHaveDiag(lngRow)
        '更新开嘱时间
        vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        vsAdvice.TextMatrix(lngRow, COL_单价) = "" '清除重新计算
        
        If vsAdvice.TextMatrix(lngRow, COL_类别) = "D" Then
            If lngAppType = 0 Then
                Call Delete检查手术输血(lngRow, True, lngTmp)
                lngRow = lngTmp
                Call AdviceSet检查申请(lngRow, objAppPages())
                strAppend = vsAdvice.TextMatrix(lngRow, COL_附项)
            Else
                '检查组合
                Call AdviceSet检查组合(lngRow, strExtData)
                vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            End If
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_类别) = "F" Then
            '一组手术
            Call AdviceSet手术组合(lngRow, strExtData)
            vsAdvice.Cell(flexcpData, lngRow, COL_标本部位) = str手术部位
            vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
        ElseIf lngAppType = 1 Then
            Call Delete检查手术输血(lngRow)
            Call DeleteRow(lngRow, True)
            Call AdviceSet输血申请(lngRow, rsCard)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
            strAppend = ""
        ElseIf lngAppType = 3 Then
            Call Delete检验申请单(lngRow, True, lngTmp)
            lngRow = lngTmp
            Call AdviceSet检验申请(lngRow, rsCard)
            strAppend = vsAdvice.TextMatrix(lngRow, COL_附项)
            txt医嘱内容.Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容)
        ElseIf RowIn检验行(lngRow) Then
            '检验组合
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "开始时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱医生", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "开嘱时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "标志", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '采集方法的执行科室可能与检验项目不同
            If Val(vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID)) <> 0 Then
                rsCurr!执行科室ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_总量)) <> 0 Then
                rsCurr!总量 = Val(vsAdvice.TextMatrix(lngRow, COL_总量))
            End If
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngRow, COL_执行时间)
            rsCurr!频率 = vsAdvice.TextMatrix(lngRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngRow, COL_间隔单位)
            rsCurr!开始时间 = vsAdvice.Cell(flexcpData, lngRow, COL_开始时间)
            rsCurr!开嘱医生 = vsAdvice.TextMatrix(lngRow, COL_开嘱医生)
            rsCurr!开嘱科室id = Val(vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID))
            rsCurr!开嘱时间 = vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!标志 = vsAdvice.TextMatrix(lngRow, COL_标志)
            '修改了检验组合内容,采集方法行应标记为修改
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '完全重新设置该检验组合
            '------------------------
            '删除检验项目行:删除之后重新定位的当前行
            lngRow = Delete检验组合(lngRow)
            '清除当前行(采集方法行)
            Call DeleteRow(lngRow, True, False)
            '重新产生:产生之后重新定位的当前行
            lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData, rsCurr)
        ElseIf RowIn配方行(lngRow) Then
            '中药配方
            lngFirstRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_相关ID)
            lng诊疗项目ID = Val(vsAdvice.TextMatrix(lngFirstRow, COL_诊疗项目ID))
            lng用法ID = Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
            
            '先获取当前已经设置好值
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "医嘱ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "执行性质", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "频率", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "频率次数", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "频率间隔", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "间隔单位", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "总量", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "执行时间", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "开始时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱医生", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "开嘱科室ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "开嘱时间", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "医生嘱托", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "标志", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!执行性质 = NeedName(cbo执行性质.Text) '正常,自备药,离院带药
            
            '取配方界面选择的药房
            rsCurr!执行科室ID = Val(Split(strExtData, "|")(4))
            
            rsCurr!频率 = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
            rsCurr!频率次数 = Val(vsAdvice.TextMatrix(lngFirstRow, COL_频率次数))
            rsCurr!频率间隔 = Val(vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔))
            rsCurr!间隔单位 = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
            
            '取配方界面选择的付数
            rsCurr!总量 = Val(Split(strExtData, "|")(3))
            
            rsCurr!执行时间 = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
            rsCurr!开始时间 = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
            rsCurr!开嘱医生 = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
            rsCurr!开嘱科室id = Val(vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID))
            rsCurr!开嘱时间 = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
            rsCurr!医生嘱托 = vsAdvice.TextMatrix(lngRow, COL_医生嘱托)
            rsCurr!标志 = vsAdvice.TextMatrix(lngRow, COL_标志)
            '修改了配方内容,用法行应标记为修改
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!医嘱ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '完全重新设置该中药配方行
            '------------------------
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '如果当前用法的配方ID不为空，则传入配方ID
            lng配方ID = Val(vsAdvice.TextMatrix(lngRow, COL_配方ID))
            '清除当前行(中药用法行)
            Call DeleteRow(lngRow, True, False)
            '产生配方:产生之后重新定位的当前行
            lngRow = AdviceSet中药配方(lng诊疗项目ID, lngRow, lng用法ID, strExtData, rsCurr, str摘要, lng配方ID)
        End If
        
        '更新附件内容:以当前可见行为准
        If strAppend <> "" Then
            vsAdvice.TextMatrix(lngRow, COL_附项) = strAppend
            vsAdvice.Cell(flexcpData, lngRow, COL_附项) = 1 '表明需要重新写入(新增或修改)
            Call ReplaceAdviceAppend(lngRow) '缺省替换其他医嘱的申请附项
        End If
        
        '强行显示当前医嘱卡片
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
                
        Call CalcAdviceMoney '显示新开医嘱金额
        
        If InStr(",0,3,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
            vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
            Call ReSetColor(lngRow)
        End If
        
        '修改后关联诊断的标记处理
        If lngDiag <> -1 Then
            Call SetDiagFlag(vsAdvice.Row, 1, lngDiag)
        End If
        
        mblnNoSave = True '标记为未保存
    End If
    
    Call vsAdvice.AutoSize(col_医嘱内容)
    
    '对保险对码进行检查
    Call GetInsureStr(strIDs1, strIDs2, str医嘱内容, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, 1, strIDs1, strIDs2, str医嘱内容)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    
    txt医嘱内容.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReplaceAdviceAppend(ByVal lngRow As Long)
'功能：根据指定行的申请附项输入情况，对其他新录医嘱的申请附项进行缺省替换
'参数：lngRow=才新输入或修改的可见医嘱行
    Dim strAppend As String, i As Long
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_附项) = "" Then Exit Sub
        
        For i = .FixedRows To .Rows - 1
            '只针对新录入的医嘱，修改的医嘱修改时已检查
            '".Cell(flexcpData, i, COL_附项) = 1"的可能还没有在输入过程中检查，只是自动替换了
            If .RowData(i) <> 0 And Not .RowHidden(i) And i <> lngRow And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                If .TextMatrix(i, COL_附项) <> "" Then
                    strAppend = ReplaceAppend(.TextMatrix(i, COL_附项), .TextMatrix(lngRow, COL_附项))
                    If .TextMatrix(i, COL_附项) <> strAppend Then
                        .TextMatrix(i, COL_附项) = strAppend
                        .Cell(flexcpData, i, COL_附项) = 1
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub ClinicSelecter(Optional ByVal 类型 As Integer, Optional ByVal lng分类id As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If 类型 = 8 And lng分类id <> 0 Then
        '直接读取选择的成套项目
        On Error GoTo errH
        strSQL = "Select A.类别 As 类别ID,A.ID as 诊疗项目ID,Null as 收费细目ID,B.名称 As 类别,A.编码,A.名称,A.计算单位,A.标本部位,NULL as 项目特性" & _
            " From 诊疗项目目录 A,诊疗项目类别 B Where A.类别=B.编码 And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng分类id)
    Else
        '打开选择器，可能指定了初始分类目录
        Set rsTmp = frmClinicSelect.ShowSelect(Me, IIF(mlng前提ID <> 0, 2, 0), 0, mlng病人科室id, 1, mstr性别, , , 1, lng分类id, mint险类)
        If rsTmp Is Nothing Then '取消或无数据
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
    End If
    
    '根据选择项目设置缺省医嘱信息
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '显示已缺省设置的值
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        
        Call CalcAdviceMoney '显示新开医嘱金额
        
        '医保管控实时监测
        If mint险类 <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
            '总量不可输入：缺省并固定总量的医嘱，以及长嘱
            '成套医嘱不在这里检查
            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) And Not txt总量.Enabled Then
                If MakePriceRecord(vsAdvice.Row) Then
                    If Not gclsInsure.CheckItem(mint险类, 0, 0, mrsPrice) Then
                        Call AdviceCurRowClear: Exit Sub
                    End If
                End If
                '标记为已经作了检查
                vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_状态) = 1
            End If
        End If
                 
        txt医嘱内容.SetFocus: Call SeekNextControl '必须先定位
    Else
        '恢复原值(AdviceInput函数中可能处理了一下)
        txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
        txt医嘱内容.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click()
    Call ClinicSelecter
End Sub

Private Sub cmd开始时间_Click()
    If IsDate(txt开始时间.Text) Then
        dtpDate.value = CDate(txt开始时间.Text)
    Else
        dtpDate.value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "开始时间"
    dtpDate.Left = txt开始时间.Left + fraAdvice.Left
    dtpDate.Top = txt开始时间.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "开始时间" Then
        '取值
        If IsDate(txt开始时间.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt开始时间.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check开始时间(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt开始时间.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt开始时间_Validate(False) '更新数据
        txt开始时间.SetFocus
    ElseIf dtpDate.Tag = "安排时间" Then
        '取值
        If IsDate(txt安排时间.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt安排时间.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '判断时间合法性
        If Not Check安排时间(strDate, txt开始时间.Text, vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt安排时间.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt安排时间_Validate(False) '更新数据
        txt安排时间.SetFocus
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpDate_DateClick(dtpDate.value)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    dtpDate.Visible = False
    dtpDate.Tag = ""
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If vsDiag.Rows = 2 And vsDiag.TextMatrix(1, col诊断) = "" Then
            If vsDiag.Enabled Then
                Call vsDiag_AfterRowColChange(vsDiag.Row, vsDiag.Col, vsDiag.Row, vsDiag.Col)
                vsDiag.SetFocus
            End If
        Else
            If txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
        End If
        '第一次进入时将滚动条移到当前行，因为Load中不能移动滚动条,调用两次才生效。
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str摘要 As String
    Dim str中药IDs As String
    Dim lngBaseRow As Long, i As Long
    Dim lng收费细目ID As Long, str诊疗项目ID As String
    
    If Shift = vbAltMask Then
        If Between(Chr(KeyCode), "1", "9") And Not mfrmShortCut Is Nothing Then
            Call mfrmShortCut.ShowShortCut(Val(Chr(KeyCode)))
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        '保存常用嘱托
        If cmd常用嘱托.Enabled And cmd常用嘱托.Visible Then
            Call cmd常用嘱托_Click
        End If
    ElseIf KeyCode = vbKeyF1 And Shift = vbCtrlMask Then
        '调用医保提示
        With vsAdvice
            If .RowData(.Row) <> 0 Then
                lng收费细目ID = Val(.TextMatrix(.Row, COL_收费细目ID))
                str诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                If RowIn配方行(.Row) Then
                    '获取中药配方第一味中药行
                    lngBaseRow = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                    For i = lngBaseRow To .Row
                        If i = lngBaseRow Then lng收费细目ID = Val(.TextMatrix(i, COL_收费细目ID))
                        If .TextMatrix(i, COL_类别) = "7" Then
                            str中药IDs = str中药IDs & "," & .TextMatrix(i, COL_诊疗项目ID)
                        End If
                    Next
                    str中药IDs = Mid(str中药IDs, 2)
                    If UBound(Split(str中药IDs, ",")) <> 0 Then
                        lng收费细目ID = 0
                    End If
                    str诊疗项目ID = str中药IDs
                End If
                '医保病人输入内容时的提示：非医保病人也要调(Or And mint险类 <> 0)
                str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, lng收费细目ID, CStr(.Cell(flexcpData, .Row, COL_医生嘱托)), 0, "", str诊疗项目ID)
                .Cell(flexcpData, .Row, COL_医生嘱托) = str摘要
            End If
        End With
    Else
        Select Case KeyCode
            Case vbKeyEscape
                If dtpDate.Visible Then
                    dtpDate.Visible = False
                    dtpDate.Tag = ""
                End If
            Case vbKeyF4, vbKeyUp, vbKeyDown
                If Me.ActiveControl Is txt开始时间 Then
                    If cmd开始时间.Visible And cmd开始时间.Enabled Then cmd开始时间_Click
                ElseIf Me.ActiveControl Is txt安排时间 Then
                    If cmd安排时间.Enabled And cmd安排时间.Visible Then cmd安排时间_Click
                ElseIf Me.ActiveControl Is txt医嘱内容 Then
                    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
                ElseIf Me.ActiveControl Is txt用法 Then
                    If cmd用法.Visible And cmd用法.Enabled Then cmd用法_Click
                ElseIf Me.ActiveControl Is txt频率 Then
                    If cmd频率.Visible And cmd频率.Enabled Then cmd频率_Click
                End If
            Case vbKeyF7 '切换输入法
                If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
                    If stbThis.Panels("WB").Bevel = sbrRaised Then
                        Call stbThis_PanelClick(stbThis.Panels("WB"))
                    Else
                        Call stbThis_PanelClick(stbThis.Panels("PY"))
                    End If
                End If
            Case vbKeyF8 '切换显示计价项目
                If stbThis.Panels("Price").Visible Then
                    Call stbThis_PanelClick(stbThis.Panels("Price"))
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        If Not mfrmShortCut Is Nothing Then Call mfrmShortCut.ShowMe(Me, mint场合, 1, 0, mlng病人科室id)
    ElseIf KeyAscii = vbKeySpace Then
        If Me.ActiveControl Is txt开始时间 And txt开始时间.SelLength = Len(txt开始时间.Text) Then
            KeyAscii = 0
            If cmd开始时间.Visible And cmd开始时间.Enabled Then cmd开始时间_Click
        ElseIf Me.ActiveControl Is txt安排时间 And txt安排时间.SelLength = Len(txt安排时间.Text) Then
            KeyAscii = 0
            If cmd安排时间.Enabled And cmd安排时间.Visible Then cmd安排时间_Click
        ElseIf Me.ActiveControl Is txt医嘱内容 Then
            KeyAscii = 0
            If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
        ElseIf Me.ActiveControl Is txt用法 Then
            KeyAscii = 0
            If cmd用法.Visible And cmd用法.Enabled Then cmd用法_Click
        ElseIf Me.ActiveControl Is txt频率 Then
            KeyAscii = 0
            If cmd频率.Visible And cmd频率.Enabled Then cmd频率_Click
        ElseIf Me.ActiveControl Is cbo滴速 Then
            KeyAscii = 0
            If cbo滴速.Visible And cbo滴速.Enabled Then zlCommFun.PressKey (vbKeyF4)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngTmp As Long
    Dim strErr As String
    
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim objForm As Object
    Dim strNames As String
    Dim i As Long
    
    mbln扩展页签 = False
    
    '外挂提供的卡片
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    If Not gobjPlugIn Is Nothing Then
        Set mcolSubForm = New Collection
        On Error Resume Next
        strTmp = gobjPlugIn.GetFormCaption(glngSys, p门诊医嘱下达)
        Call zlPlugInErrH(err, "GetFormCaption")
        If strTmp <> "" Then
            arrTmp = Split(strTmp, ",")
            For i = 0 To UBound(arrTmp)
                strTmp = arrTmp(i)
                Set objForm = gobjPlugIn.GetForm(glngSys, p门诊医嘱下达, strTmp)
                Call zlPlugInErrH(err, "GetForm")
                If Not objForm Is Nothing Then
                    mcolSubForm.Add objForm, "_" & strTmp
                    strNames = strNames & "," & strTmp
                End If
                Set objForm = Nothing
            Next
        End If
        err.Clear: On Error GoTo 0
    End If

    If strNames <> "" Then
        mbln扩展页签 = True
        strNames = Mid(strNames, 2)
    End If
    
    If mbln扩展页签 Then
        arrTmp = Split(strNames, ",")
        With tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .ClientFrame = xtpTabFrameSingleLine
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
            
            .InsertItem(0, "医嘱编辑", picSub.hWnd, 0).Tag = "医嘱编辑"
            For i = 0 To UBound(arrTmp)
                strTmp = arrTmp(i)
                .InsertItem(i + 1, strTmp, mcolSubForm("_" & strTmp).hWnd, 0).Tag = strTmp
            Next
        End With
    Else
        tbcSub.Visible = False
    End If
    
    Call InitObjLis(p门诊医生站)
    If gobjLIS Is Nothing Then
        mblnNewLIS = False
    Else
        On Error Resume Next
        mblnNewLIS = gobjLIS.GetApplicationFormShowType
        err.Clear: On Error GoTo 0
    End If
    Call InitCommandBar
    Call InitAdviceTable
    If mint场合 = 0 Then
        '字体设置
        mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")
    ElseIf mint场合 = 2 Then
        mbytSize = zlDatabase.GetPara("字体", glngSys, p医技工作站, "0")
    End If
    Call SetFontSize(mbytSize)
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(0) = 14 * Screen.TwipsPerPixelX

    Call zlControl.CboSetHeight(cbo滴速, Me.Height)
    Call zlControl.CboSetHeight(cbo执行科室, Me.Height)
    Call zlControl.CboSetWidth(cbo执行科室.hWnd, cbo执行科室.Width * 1.3)
    fraPati.BackColor = Me.BackColor
    fra诊断.BackColor = Me.BackColor

    mblnOK = False
    mblnNoSave = False
    mblnRowMerge = False
    mblnRunFirst = True
    mblnRowChange = True
    mblnDoCheck = True
    mstrDelIDs = ""
    mstrAduitDelIDs = ""
    mstrDel输血 = ""
    mlngID序列 = 0
    mblnAddAgent = zlDatabase.GetPara("要求登记代办人", glngSys, p门诊医嘱下达, "1") = "1"
    mblnFreeInput = Val(zlDatabase.GetPara("诊断手术名称自由调整", glngSys, 0)) = 1
    '输入匹配
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    '简码匹配方式：0-拼音,1-五笔
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    '启用屏幕键盘
    mblnStaKB = Val(zlDatabase.GetPara("启用屏幕键盘", glngSys, p门诊医生站)) <> 0
    If Not mblnStaKB Then
        stbThis.Panels("KB").Visible = False
    End If
    stbThis.Panels("KB").ToolTipText = "点击启用屏幕键盘"
    Select Case mint简码
    Case 0
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrRaised
    Case 1
        stbThis.Panels("PY").Bevel = sbrRaised
        stbThis.Panels("WB").Bevel = sbrInset
    Case Else
        stbThis.Panels("PY").Bevel = sbrInset
        stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not gbln简码匹配方式切换 Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If

    'PASS接口初始化
    Call zlPASSMap
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then        'Pass
            '病人过敏史/病生状态可用检测
            Call gobjPass.zlPassCmdAlleyEnable(mobjPassMap)
        End If
    End If
    '诊断输入来源
    opt诊断(Val(zlDatabase.GetPara("门诊诊断输入", glngSys, p门诊医生站, 0, Array(opt诊断(0), opt诊断(1)), InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱选项设置") > 0))).value = True
    If gint诊断来源 > 1 Then
        opt诊断(0).Enabled = False
        opt诊断(1).Enabled = False
        If gint诊断来源 = 2 Then
            opt诊断(0).value = True
        ElseIf gint诊断来源 = 3 Then
            opt诊断(1).value = True
        End If
    End If
    opt诊断(0).TabStop = False
    opt诊断(1).TabStop = False

    '计价面板状态
    If mblnModal Then
        stbThis.Panels("Price").Visible = False
    Else
        Set mfrmPrice = New frmAdvicePrice
        stbThis.Panels("Price").Tag = zlDatabase.GetPara("显示医嘱计价面板", glngSys, p门诊医嘱下达)
    End If

    '必须录入药品单量
    mbln单量 = Val(zlDatabase.GetPara("必须录入药品单量", glngSys, p门诊医嘱下达)) <> 0

    '执行天数
    mbln天数 = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, p门诊医嘱下达)) <> 0
    vsAdvice.ColHidden(COL_天数) = Not mbln天数
    
    mbln天数反算 = (gbln反算天数 And mbln天数)

    '抗菌药物缺省用药目的
    mstrPurMed = zlDatabase.GetPara("抗菌药物缺省用药目的", glngSys, p门诊医嘱下达, "0")

    '自动处理皮试
    mbln自动皮试 = Val(zlDatabase.GetPara("自动处理皮试", glngSys, p门诊医嘱下达)) <> 0 And mlng前提ID = 0

    '自动关闭窗体
    mblnAutoClose = Val(zlDatabase.GetPara("发送完成后关闭医嘱窗体", glngSys, p门诊医嘱下达)) <> 0

    '药品、卫材出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    With cboDruPur
        .Clear
        .AddItem " "
        .AddItem "预防"
        .AddItem "治疗"
    End With
    
    '常用嘱托
    Call ReadEnjoin
    '医嘱内容定义
    If CreateScript(mobjVBA, mobjScript) Then
        Set mrsDefine = InitAdviceDefine
    End If
    '--------------------------------------------------
    '读取病人信息
    Call GetPatiInfo
    Call SetBabyVisible(mlng病人科室id)
    '读取代办人信息
    Call GetAgentInfo

    '修改时强行定位婴儿
    If mlng医嘱ID = 0 Then    '新增
        cbo婴儿.ListIndex = 0    '缺省新增病人的医嘱
    Else    '修改
        cbo婴儿.ListIndex = mint婴儿
    End If
    cbo婴儿.Tag = cbo婴儿.ListIndex

    '读取并显示病人医嘱
    Call ReLoadAdvice(mlng医嘱ID)

    '处理快捷输入窗体
        If mblnModal = False Then
        Set mfrmShortCut = New frmClinicShortCut
        mfrmShortCut.ShowMe Me, mint场合, 1, 0, mlng病人科室id, True    '根据上次上否显示
        End If
    vsDiag.AllowUserResizing = flexResizeNone

    If mblnStaKB Then
        On Error Resume Next
        Set mobjKeyBoard = Nothing
        Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        err.Clear: On Error GoTo 0
        If Not mobjKeyBoard Is Nothing Then
            Call mobjKeyBoard.StartUp
        Else
            stbThis.Panels("KB").Visible = False
            MsgBox "屏幕键盘部件未能正确安装，不能使用！", vbInformation, gstrSysName
        End If
    End If
    stbThis.Visible = True
End Sub

Private Function TheStockCheck(ByVal lng库房ID As Long, ByVal str类别 As String) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    
    On Error Resume Next
    If InStr(",5,6,7,", str类别) > 0 Then
        intStyle = mcolStock1("_" & lng库房ID)
    ElseIf str类别 = "4" Then
        intStyle = mcolStock2("_" & lng库房ID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Sub ReLoadAdvice(Optional ByVal lng医嘱ID As Long)
'功能：重新读取并显示病人的当前医嘱清单
'参数：lng医嘱ID=用于定位
    Dim lngRow As Long
    
    If LoadAdvice Then
        '显示关联医嘱标识
        Call ShowDiagFlag(vsDiag.Row)
    
        '显示可见的医嘱
        Call ShowAdvice
        
        If lng医嘱ID = 0 Then
            If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
                cbsMain.FindControl(, conMenu_New, True, True).Execute
            End If
        Else
            '修改的医嘱ID应该是显示行
            lngRow = vsAdvice.FindRow(lng医嘱ID)
            If lngRow <> -1 Then
                If Not vsAdvice.RowHidden(lngRow) Then
                    mblnRowChange = False
                    vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngRow
                    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
                    mblnRowChange = True
                End If
            End If
        End If
        '进入时屏蔽了ShowAdvice中的调用,强行进入
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Function ReadEnjoin() As Boolean
'功能：读取并加入常用滴速
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPre As String
        
    On Error GoTo errH
    
    '常用滴速
    strPre = cbo滴速.Text '加入后保持原有值
    Call Load输液滴速(cbo滴速, lbl滴速单位, False)
    cbo滴速.Text = strPre
    
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    If dtpDate.Visible Then
        dtpDate.Visible = False
        dtpDate.Tag = ""
    End If
    
    If mbln扩展页签 Then
        tbcSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Else
        picSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Call cbsMain_Resize
    
    'Pass
    cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 2 * Screen.TwipsPerPixelX
    cbo婴儿.Left = Me.ScaleWidth - IIF(cmdAlley.Visible, cmdAlley.Width + 30, 0) - cbo婴儿.Width - 2 * Screen.TwipsPerPixelX
    lbl婴儿.Left = cbo婴儿.Left - lbl婴儿.Width - 2 * Screen.TwipsPerPixelX
    
    If cmdAlley.Visible Or lbl婴儿.Visible Then
        lblPati.Width = IIF(lbl婴儿.Visible, lbl婴儿.Left, cmdAlley.Left) - lblPati.Left - 6 * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mdat挂号时间 = Empty
    msng天数 = 0
    
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    Set mrsDrugScale = Nothing
    Set mfrmSend = Nothing
    Set mrsPrice = Nothing
    Set mobjKeyBoard = Nothing
    Set mclsMipModule = Nothing
    '计价面板状态
    If Not mfrmPrice Is Nothing Then
        Unload mfrmPrice
        Set mfrmPrice = Nothing
        Call zlDatabase.SetPara("显示医嘱计价面板", IIF(Val(stbThis.Panels("Price").Tag) <> 0, 1, 0), glngSys, p门诊医嘱下达, InStr(GetInsidePrivs(p门诊医嘱下达), ";医嘱选项设置;") > 0)
    End If
    
    If mbln扩展页签 Then
        For i = 1 To mcolSubForm.Count
            Unload mcolSubForm(i)
        Next
        Set mcolSubForm = Nothing
    End If
    
    If mblnPass Then
        Call gobjPass.zlPassClearLight(mobjPassMap)
        Set mobjPassMap = Nothing
    End If
     
    Call zlDatabase.SetPara("门诊诊断输入", IIF(opt诊断(0).value, 0, 1), glngSys, p门诊医生站, InStr(GetInsidePrivs(p门诊医生站), ";参数设置;") > 0)
    Call SaveWinState(Me, App.ProductName)
    mblnNoSave = False
    RaiseEvent FormUnload(Cancel)
    Set mrs诊断 = Nothing
    mlng危急值ID = 0
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'功能：判断两行是否可以一并给药
'参数：lngRow1=前面一条已经输入的药品行
'      lngRow2=当前行(已输入或未输入)
'返回：如果不可以，则strMsg返回提示信息
    Dim lngFind As Long, lngRxCount As Long
    Dim lng配制中心 As Long, str药房IDs As String
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '必须全部为成药且类别相同
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
            
            '不能包含已发送的医嘱
            If Val(.TextMatrix(lngRow1, COL_状态)) <> 1 Then
                strMsg = "要设置为一并给药的药品包含已经发送的医嘱。"
                Exit Function
            End If
            '不能包含已签名的医嘱
            If Val(.TextMatrix(lngRow1, COL_签名否)) = 1 Then
                strMsg = "要设置为一并给药的药品包含已经签名的医嘱。"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_类别)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_类别)) = 0 Then
                strMsg = "一并给药的药品必须都为西成药或都为中成药。"
                Exit Function
            End If
            
            '不能包含已发送的医嘱
            If Val(.TextMatrix(lngRow1, COL_状态)) <> 1 Or Val(.TextMatrix(lngRow2, COL_状态)) <> 1 Then
                strMsg = "要设置为一并给药的药品包含已经发送的医嘱。"
                Exit Function
            End If
            '不能包含已签名的医嘱
            If Val(.TextMatrix(lngRow1, COL_签名否)) = 1 Or Val(.TextMatrix(lngRow2, COL_签名否)) = 1 Then
                strMsg = "要设置为一并给药的药品包含已经签名的医嘱。"
                Exit Function
            End If
            
            '一并给药(前面药品)的给药途径是否适用于当前药品
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_相关ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check适用用法(Val(.TextMatrix(lngFind, COL_诊疗项目ID)), Val(.TextMatrix(lngRow2, COL_诊疗项目ID)), 1) Then
                    strMsg = """" & .TextMatrix(lngRow2, col_医嘱内容) & """不能使用""" & .TextMatrix(lngFind, col_医嘱内容) & """给药途径，" & _
                    vbCrLf & "不能与""" & .TextMatrix(lngRow1, col_医嘱内容) & """设置为一并给药。"
                    Exit Function
                End If
            End If
            
            '检查如果有配制中心，是否都可以存储，自备药不管它
            If Not (Val(.TextMatrix(lngRow1, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow1, COL_执行性质)) = 5) Then
                If Have部门性质(Val(.TextMatrix(lngRow1, COL_执行科室ID)), "配制中心") Then
                    lng配制中心 = Val(.TextMatrix(lngRow1, COL_执行科室ID))
                End If
            End If
            If lng配制中心 = 0 Then
                If Not (Val(.TextMatrix(lngRow2, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow2, COL_执行性质)) = 5) Then
                    If Have部门性质(Val(.TextMatrix(lngRow2, COL_执行科室ID)), "配制中心") Then
                        lng配制中心 = Val(.TextMatrix(lngRow2, COL_执行科室ID))
                    End If
                End If
            End If
            If lng配制中心 <> 0 Then
                If Not (Val(.TextMatrix(lngRow1, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow1, COL_执行性质)) = 5) Then
                    str药房IDs = Get可用药房IDs(.TextMatrix(lngRow1, COL_类别), Val(.TextMatrix(lngRow1, COL_诊疗项目ID)), Val(.TextMatrix(lngRow1, COL_收费细目ID)), mlng病人科室id, 1)
                    If InStr("," & str药房IDs & ",", "," & lng配制中心 & ",") = 0 Then
                        strMsg = "药品""" & .TextMatrix(lngRow1, col_医嘱内容) & """在配制中心""" & Get部门名称(lng配制中心) & """没有存储。"
                        Exit Function
                    End If
                End If
                If Not (Val(.TextMatrix(lngRow2, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow2, COL_执行性质)) = 5) Then
                    str药房IDs = Get可用药房IDs(.TextMatrix(lngRow2, COL_类别), Val(.TextMatrix(lngRow2, COL_诊疗项目ID)), Val(.TextMatrix(lngRow2, COL_收费细目ID)), mlng病人科室id, 1)
                    If InStr("," & str药房IDs & ",", "," & lng配制中心 & ",") = 0 Then
                        strMsg = "药品""" & .TextMatrix(lngRow2, col_医嘱内容) & """在配制中心""" & Get部门名称(lng配制中心) & """没有存储。"
                        Exit Function
                    End If
                End If
            End If
        End If
        
        '检查处方药品种数限制
        If gintRXCount > 0 Then
            lngFind = .FindRow(.TextMatrix(lngRow1, COL_相关ID), , COL_相关ID)
            lngRxCount = GetMergeCount(vsAdvice, lngFind, COL_相关ID, COL_收费细目ID)
            If lngRxCount >= gintRXCount Then
                strMsg = "一并给药的药品种数 " & lngRxCount & " 种已达到或超过药品处方最多允许的种数 " & gintRXCount & " 种。"
                Exit Function
            End If
        End If
    End With
    RowCanMerge = True
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng医嘱ID As Long, lng相关ID As Long
    Dim str类别 As String, lngTmp As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim lng诊疗项目ID As Long, i As Long, j As Long
    Dim lngDiag As Long, lngSeek As Long
    Dim intLoop As Integer
    Dim blnTag As Boolean
    
    Dim lng病人ID As Long, str挂号单 As String, blnMoved As Boolean
    
    Call AdviceChange '强制更新医嘱内容
    
    With vsAdvice
        Select Case Control.ID
            Case conMenu_New
                If .RowData(.Row) = 0 Then
'                    If .Row <> .Rows - 1 Then
'                        MsgBox "当前行无内容，请先在当前行录入有效医嘱或删除当前行。", vbInformation, gstrSysName
'                    Else
'                        MsgBox "当前行无内容，请先在当前行录入有效医嘱。", vbInformation, gstrSysName
'                    End If
'                    Exit Sub
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '先删除中间间隔的空行
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                Call .ShowCell(.Row, .Col)
                If Visible And txt医嘱内容.Enabled Then txt医嘱内容.SetFocus
            Case conMenu_Insert
                If .RowData(.Row) = 0 Then
                    MsgBox "当前行无内容，请先在当前行录入有效医嘱。", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '插入后成自动成为一并给药:插入在一并给药的中间才行
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                        
                        '不能在已发送的一并给药中插入
                        If Val(.TextMatrix(.Row, COL_状态)) <> 1 Then
                            MsgBox "该组一并给药的医嘱已经发送，不能再插入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '不能在已签名的一并给药中插入
                        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                            MsgBox "该组一并给药的医嘱已经签名，不能再插入。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        lng相关ID = Val(.TextMatrix(lngPreRow, COL_相关ID))
                    End If
                End If
                
                '先删除中间间隔的空行
                mblnRowChange = False
                lng医嘱ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lng医嘱ID)
                mblnRowChange = True
                            
                '当前行之前插入新行
                '--------------------------------------------------------------
                If RowIn配方行(.Row) Or RowIn检验行(.Row) Then
                    '中药配方及检验组合行是前面的行隐藏
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                txt医嘱内容.SetFocus '先定位避免光标晃
                        Case conMenu_Edit_ViewDrugExplain '查看药品说明书
                Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_收费细目ID)), Me)
            Case conMenu_Merge '一并给药
                If Not Control.Checked Then '想按下
                    lngBegin = GetPreRow(.Row)
                    '前面没有行
                    If lngBegin = -1 Then
                        MsgBox "前面没有可以一并给药的医嘱行。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '两行不符合条件
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '当前行尚未输入内容的情况
                        If DateDiff("n", CDate(.Cell(flexcpData, lngBegin, COL_开始时间)), zlDatabase.Currentdate) <= gint门诊新开医嘱间隔 Then
                            txt开始时间.Text = .Cell(flexcpData, lngBegin, COL_开始时间)
                        End If
                        mblnRowMerge = True: cbsMain.RecalcLayout '*允许按下
                        txt医嘱内容.SetFocus: Exit Sub
                    Else
                        '要把当前行与前面行一起一并给药
                        Call MergeRow(lngBegin, .Row, False)
                        Call ReSetColor(.Row) '一并之后再一并设置
                    End If
                Else '想弹起
                    If .RowData(.Row) = 0 Then
                        '当前行尚未输入内容的情况
                        If Not RowIn一并给药(.Row) Then
                            mblnRowMerge = False '*允许弹起
                            cbsMain.RecalcLayout
                        End If
                        Exit Sub
                    Else
                        '当前行是一并给药中的行
                        Call Get一并给药范围(Val(.TextMatrix(.Row, COL_相关ID)), lngBegin, lngEnd)
                                                
                        '先判断可否取消一并给药
                        '不能包含已发送的医嘱
                        If Val(.TextMatrix(.Row, COL_状态)) <> 1 Then
                            MsgBox "当前医嘱已经发送。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '不能包含已签名的医嘱
                        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                            MsgBox "当前医嘱已经签名。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                                                
                        '先提示
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '整个一并给药取消为单独给药
                            If MsgBox("要将该组一并给药的药品全部取消为单独给药吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Exit Sub
                            End If
                        End If
                        
                        '删除中间的空行
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        '记录当前一并时的诊断关联以便恢复
                        lngDiag = AdviceHaveDiag(.Row)
                        lngSeek = .RowData(lngEnd)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '从一并给药中分离该行
                            Call ReSetColor(.Row) '在取消之前一并设置
                            Call SplitRow(.Row)
                        Else
                            '取消一并给药
                            Call ReSetColor(.Row) '在取消之前一并设置
                            lngTmp = .RowData(.Row) '记录用于恢复行定位
                            Call AdviceSet单独给药(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                        
                        '根据记录的诊断关联进行恢复
                        If lngDiag <> -1 Then
                            lngSeek = .FindRow(lngSeek, lngBegin)
                            For i = lngBegin To lngSeek
                                If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 _
                                    And Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) _
                                    And Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                                    Call SetDiagFlag(i, 1, lngDiag)
                                End If
                            Next
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case conMenu_Delete
                If .RowSel <> .Row Then
                    MsgBox "一次只能删除一条医嘱，请选择要删除的医嘱行。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    '已发送的医嘱不能删除
                    If Val(.TextMatrix(.Row, COL_状态)) <> 1 Then
                        MsgBox "该条医嘱已经发送，不能删除。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '已签名的医嘱不能删除
                    If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                        MsgBox "该条医嘱已经签名，不能删除。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                Call AdviceDelete(.Row) '删除当前行
                Call CalcAdviceMoney '显示新开医嘱金额
                
                vsAdvice.SetFocus
            Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10# + 1 '检查申请
                Call AdviceInput申请单(1)
            Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10# + 1 '检验申请
                Call AdviceInput申请单(2)
            Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10# + 1 '输血申请
                Call AdviceInput申请单(3)
                        Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101# '自定义申请单
                FuncApplyCustom 0, Control.Parameter
            Case conMenu_AdvicePay
                Call FuncClinicPay(Me, mlng病人ID, mstr挂号单)
            Case conMenu_Reference
                If Val(.TextMatrix(.Row, COL_诊疗项目ID)) <> 0 Then
                    If RowIn配方行(.Row) Or RowIn检验行(.Row) Then
                        i = .FindRow(CStr(.RowData(.Row)), , COL_相关ID)
                        If i <> -1 Then
                            lng诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                        End If
                    Else
                        lng诊疗项目ID = Val(.TextMatrix(.Row, COL_诊疗项目ID))
                    End If
                End If
                '诊疗措施参考
                Call frmClinicHelp.ShowMe(IIF(mblnModal, 1, 0), mfrmParent, lng诊疗项目ID)
            Case conMenu_Copy
                lng病人ID = mlng病人ID: str挂号单 = mstr挂号单: blnMoved = False
                strMsg = frmAdviceCopy.ShowMe(Me, mMainPrivs, lng病人ID, str挂号单, blnMoved, False, mlng前提ID, , mlng病人科室id, , , mstr性别)
                If strMsg <> "" Then
                    cbsMain.FindControl(, conMenu_New, True, True).Execute
                    Call AdviceSet复制医嘱(lng病人ID, str挂号单, strMsg, blnMoved)
                End If
            Case conMenu_Scheme, conMenu_Scheme * 10# + 1
                Call frmAdviceScheme.ShowMe(IIF(mlng前提ID <> 0, 2, 0), 1, mlng病人ID, 0, mstr挂号单, cbo婴儿.ListIndex, Me)
            Case conMenu_Scheme * 10# + 2
                Call mfrmShortCut.ShowMe(Me, mint场合, 1, 0, mlng病人科室id)
            Case conMenu_Agent
                
                For intLoop = 1 To vsAdvice.Rows - 1
                    If (Val(vsAdvice.TextMatrix(intLoop, COL_状态)) = 1 And InStr(",麻醉药,毒性药,精神I类,", "," & Trim(vsAdvice.TextMatrix(intLoop, COL_毒理分类)) & ",") > 0) Then
                        blnTag = True: Exit For
                    End If
                Next
                If Not blnTag Then MsgBox "病人未执行的医嘱中不存在麻醉药、毒性药、精神I类药品医嘱，不需要填写代办人信息！", vbInformation, gstrSysName: Exit Sub
                Call GetPatiInfo
                Call frmAgentInfo.ShowMe(Me, mlng病人ID, mlng挂号ID, mstr姓名, mstr身份证号, AgentInfo.代办人姓名, AgentInfo.代办人身份证号)
                Call GetAgentInfo
                Screen.MousePointer = 0
            Case conMenu_Save
                If vsDiag.EditText <> "" Then
                    mblnCancle = False
                    Me.SetFocus
                    If mblnCancle = True Then mblnCancle = False: Exit Sub
                End If
                If Not CheckAdvice Then Exit Sub '检查中处理了光标定位
                If Not SaveAdvice Then .SetFocus: Exit Sub
            Case conMenu_Send, conMenu_Send * 10# + 1, conMenu_Send * 10# + 2
                '发送之前自动保存
                If mblnNoSave Then
                    If Not CheckAdvice Then Exit Sub
                    If Not SaveAdvice Then .SetFocus: Exit Sub
                End If
                If mfrmSend Is Nothing Then Set mfrmSend = New frmOutAdviceSend
                If mfrmSend.ShowMe(Me, mMainPrivs, mlng病人ID, mstr挂号单, mstr前提IDs, _
                    Control.ID = conMenu_Send And Control.Type = xtpControlSplitButtonPopup Or Control.ID = conMenu_Send * 10# + 1, mlng医技科室ID, mint场合, mclsMipModule) Then
                    '病人医嘱发送完成后自动关闭窗体
                    If mblnAutoClose Then
                        If Not ExistNoSendAdvice(mlng病人ID, mstr挂号单) Then
                            mblnOK = True: Unload Me: Exit Sub
                        End If
                    End If
                    
                    '重新读取显示医嘱
                    Call ReLoadAdvice
                    mblnOK = True '强行
                    If txt医嘱内容.Enabled Then
                        txt医嘱内容.SetFocus
                    Else
                        .SetFocus
                    End If
                End If
            Case conMenu_Sign
                Call AdviceSign
            Case conMenu_Help
                ShowHelp App.ProductName, Me.hWnd, Me.Name
            Case conMenu_Exit
                Unload Me
            Case conMenu_DrugScale * 100# + 1 To conMenu_DrugScale * 100# + 99
                With vsAdvice
                    txt单量.Text = FormatEx(Val(.TextMatrix(.Row, COL_剂量系数)) * Val(Control.Parameter), 5)
                    Call zlControl.TxtSelAll(txt单量)
                End With
            Case conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  'PASS合理用药
                If mblnPass Then
                    Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#, mblnNoSave)
                End If
            Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
                Call ExePlugIn(Control.Parameter)
        End Select
    End With
End Sub

Private Function ExistNoSendAdvice(ByVal lng病人ID As Long, ByVal str挂号单 As String) As Boolean
'功能：检查当前病人的医嘱是否已经发送完成。
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '叮嘱、免试皮试、护理等级不发送
    strSQL = "Select 1 From 病人医嘱记录 A,诊疗项目目录 B" & _
        " Where Nvl(A.皮试结果,'无')<>'免试' And Not (A.诊疗类别='H' And B.操作类型='1')" & _
        " And Nvl(A.执行性质,0)<>0 And A.开始执行时间 is Not NULL And A.病人来源<>3" & _
        " And A.诊疗项目ID=B.ID And A.医嘱状态=1 And A.病人ID=[1] And A.挂号单=[2] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistNoSendAdvice", lng病人ID, str挂号单)
    If Not rsTmp.EOF Then ExistNoSendAdvice = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Get一并给药范围(ByVal lng相关ID As Long, lngBegin As Long, lngEnd As Long)
'功能：根据相关的给药途径医嘱ID,确定一并给药的一组药品的起止行号
'说明：中间可能包含有空行
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng相关ID), , COL_相关ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txt单量_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_单量)) <> Val(txt单量.Text) Then
                txt单量.Tag = "1"
            End If
        Else
            txt单量.Tag = "1"
        End If
    End With
    
End Sub

Private Sub txt单量_GotFocus()
    zlControl.TxtSelAll txt单量
End Sub

Private Sub txt单量_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean

    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt单量.Text) And Val(txt单量.Text) > 0 _
            Or txt单量.Text = "" And (Not mbln单量 Or InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) = 0) Then
            Call txt单量_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt单量_Validate(Cancel As Boolean)
    Dim strMsg As String, dbl次数 As Double, sng天数 As Single
    Dim lngFind As Long, blnDo As Boolean
    Dim sng总量 As Single
    
    With vsAdvice
        If Val(txt单量.Text) = 0 Then txt单量.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Not IsNumeric(txt单量.Text) Then
            If txt单量.Text <> "" Then
                Cancel = True: txt单量_GotFocus: Exit Sub
            ElseIf txt单量.Text = "" And mbln单量 And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 Then
                Cancel = True: txt单量_GotFocus: Exit Sub
            End If
        ElseIf CDbl(txt单量.Text) <= 0 Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        ElseIf CDbl(txt单量.Text) > LONG_MAX Then
            Cancel = True: txt单量_GotFocus: Exit Sub
        ElseIf txt单量.Text <> "" Then
            '单量合法性检查
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_收费细目ID)) <> 0 Then
                blnDo = Not txt天数.Visible '给药方法为叮嘱，或输入天数时不检查
                If blnDo Then
                    lngFind = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))), .Row + 1)
                    If lngFind <> -1 Then blnDo = blnDo And Val(.TextMatrix(lngFind, COL_执行性质)) <> 0
                End If
                If blnDo Then
                    dbl次数 = IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, 1, Val(.TextMatrix(.Row, COL_总量))) * _
                        Val(.TextMatrix(.Row, COL_门诊包装)) * Val(.TextMatrix(.Row, COL_剂量系数)) / Val(txt单量.Text)
                    If dbl次数 > 200 Then
                        If MsgBox("该药品按每次 " & FormatEx(txt单量.Text, 5) & .TextMatrix(.Row, COL_单量单位) & " 使用，" & _
                            IIF(Val(.TextMatrix(.Row, COL_总量)) = 0, "每", Val(.TextMatrix(.Row, COL_总量))) & _
                            .TextMatrix(.Row, COL_门诊单位) & "可以使用 " & FormatEx(dbl次数, 5) & " 次。" & _
                            vbCrLf & vbCrLf & "你确认单量输入正确吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                            Cancel = True: txt单量_GotFocus: Exit Sub
                        End If
                    End If
                End If
            End If
            
            txt单量.Text = FormatEx(txt单量.Text, 5)
            
            '重新计算药品总量(先输入单量时)
            If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                If mbln天数 Then
                    If mbln天数反算 Then
                        If Val(txt总量.Text) <> 0 Then
                            '开始反算天数
                            '不输天数时，根据总量、单量、频率计算天数，检查是否超期
                            If Val(txt总量.Text) <> 0 And Val(txt单量.Text) <> 0 <> 0 _
                                And .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率次数)) <> 0 And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 _
                                And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                                
                                sng天数 = Calc缺省药品天数(Val(txt总量.Text), Val(txt单量.Text), _
                                    Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位), _
                                    Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_门诊包装)), _
                                    Val(.TextMatrix(.Row, COL_可否分零)))
                                Call CheckDrugOutOfRange(.Row, sng天数)
                            End If
                        End If
                        If sng天数 = 0 Then sng天数 = 1
                        msngPre天数 = sng天数
                        txt天数.Text = sng天数
                    Else
                        sng天数 = Val(.TextMatrix(.Row, COL_天数))
                        If sng天数 = 0 Then sng天数 = 1
                        sng总量 = Val(txt总量.Text)
                        
                        txt总量.Text = ReGet药品总量(sng总量, Val(txt单量.Text), sng天数, .Row)
                        '隐式调用了Change事件
                                               
                        Call txt总量_Validate(Cancel)
                        If Cancel Then
                            txt总量.Text = sng总量
                            Exit Sub
                        End If
                    End If
                Else
                    '不输天数时，根据总量、单量、频率计算天数，检查是否超期
                    If Val(txt总量.Text) <> 0 And Val(txt单量.Text) <> 0 <> 0 _
                        And .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率次数)) <> 0 And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 _
                        And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                        
                        sng天数 = Calc缺省药品天数(Val(txt总量.Text), Val(txt单量.Text), _
                            Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位), _
                            Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_门诊包装)), _
                            Val(.TextMatrix(.Row, COL_可否分零)))
                            
                        Call CheckDrugOutOfRange(.Row, sng天数)
                    End If
                End If
            End If
            
            '更新数据
            Call AdviceChange
        End If
    End With
End Sub

Private Sub txt开始时间_Change()
    txt开始时间.Tag = "1"
End Sub

Private Sub txt开始时间_GotFocus()
    If txt开始时间.Text = "" Then txt开始时间.Text = GetDefaultTime(vsAdvice.Row)
    zlControl.TxtSelAll txt开始时间
End Sub

Private Sub txt开始时间_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt开始时间.Text <> "" Then
            txt开始时间.Text = GetFullDate(txt开始时间.Text)
            If SeekNextControl Then Call txt开始时间_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt开始时间_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt开始时间.Locked Then
        glngTXTProc = GetWindowLong(txt开始时间.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt开始时间.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt开始时间_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt开始时间.Locked Then
        Call SetWindowLong(txt开始时间.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt开始时间_Validate(Cancel As Boolean)
    If txt开始时间.Locked Then Exit Sub
        
    If Not IsDate(txt开始时间.Text) Then
        If txt开始时间.Text <> "" Then
            Cancel = True
            txt开始时间_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)) Then
                '恢复人为的清除
                txt开始时间.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间)
            End If
        End If
    Else
        '检查时间合法性
        If Not Check开始时间(txt开始时间.Text) Then
            Cancel = True
            txt开始时间_GotFocus
            Exit Sub
        End If
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cbo医生嘱托_Change()
    cbo医生嘱托.Tag = "1"
End Sub

Private Sub cbo医生嘱托_GotFocus()
    zlControl.TxtSelAll cbo医生嘱托
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo医生嘱托_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo医生嘱托.Text <> "" Then
            If ReasonSelect(cbo医生嘱托.Text, 2) Then Exit Sub
        End If
        If SeekNextControl Then Call cbo医生嘱托_Validate(False)
    End If
End Sub

Private Sub cbo医生嘱托_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cbo医生嘱托.Text) > 100 Then
        MsgBox "输入内容不过超过 50 个汉字或 100 个字符。", vbInformation, gstrSysName
        cbo医生嘱托_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub txt医嘱内容_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txt医嘱内容_GotFocus()
    If txt开始时间.Text = "" Then txt开始时间_GotFocus
    Call zlControl.TxtSelAll(txt医嘱内容)
End Sub

Private Sub txt医嘱内容_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intTmp As Integer
    
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txt医嘱内容)
    End If
    If KeyCode = vbKeySpace And txt医嘱内容.Text = "" And gblnOut必用 Then
        intTmp = ApplySelect
        If intTmp <> 0 Then Call AdviceInput申请单(intTmp)
    End If
End Sub

Private Sub txt医嘱内容_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim str输入 As String
    Dim blnBarcode As Boolean '是否置空卫材的 批次 字段

    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt医嘱内容.Text = "" Then Exit Sub
        If txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        Set rsTmp = frmClinicSelect.ShowSelect(Me, IIF(mlng前提ID <> 0, 2, 0), 0, mlng病人科室id, 1, mstr性别, txt医嘱内容.Text, txt医嘱内容, 1, , mint险类)
        If rsTmp Is Nothing Then '取消或无数据
            '恢复原值
            'txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
        '新项目的录入
        '成套项目中如果包含成药,则不能按规格下医嘱
        
        If Val(rsTmp!类别ID & "") = 4 And rsTmp!批次 & "" <> "" Then
            str输入 = txt医嘱内容.Text
            '使用条码匹配的规则：1、全数字或者数字+字母；长度10位数以上；
            If (Not zlCommFun.IsCharChinese(str输入)) And Len(str输入) >= 10 Then
                If InStr("," & rsTmp!编码, "," & str输入) > 0 Then
                    '则匹配的是编码，置空批次
                    blnBarcode = True
                End If
            End If
            If IsNumeric(str输入) Then
                '1X.输入全是数字时只匹配编码
                If Mid(gstrMatchMode, 1, 1) = "1" Then
                    If Len(str输入) >= 10 Then
                        If InStr("," & rsTmp!编码, "," & str输入) > 0 Then
                            '则匹配的是编码，置空批次
                            blnBarcode = True
                        End If
                    End If
                End If
            ElseIf zlCommFun.IsCharAlpha(str输入) Then
                'X1.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then
                    If Len(str输入) >= 10 Then
                        If InStr("," & rsTmp!简码, "," & str输入) > 0 Then
                            '则匹配的是编码，置空批次
                            blnBarcode = True
                        End If
                    End If
                End If
            End If
        End If
   
        If blnBarcode Then
            Set rsTmpOther = zlDatabase.CopyNewRec(rsTmp)
            rsTmpOther!批次 = Null
            Set rsTmp = Nothing
            Set rsTmp = rsTmpOther
        End If
        
        '根据选择项目设置缺省医嘱信息
        Me.Refresh
        If AdviceInput(rsTmp, vsAdvice.Row) Then
            '显示已缺省设置的值
            Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
                        
            Call CalcAdviceMoney '显示新开医嘱金额
            
            '医保管控实时监测
            If mint险类 <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
                '总量不可输入：缺省并固定总量的医嘱，以及长嘱
                '成套医嘱不在这里检查
                If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) And Not txt总量.Enabled Then
                    If MakePriceRecord(vsAdvice.Row) Then
                        If Not gclsInsure.CheckItem(mint险类, 0, 0, mrsPrice) Then
                            Call AdviceCurRowClear: Exit Sub
                        End If
                    End If
                    '标记为已经作了检查
                    vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_状态) = 1
                End If
            End If
            
            Call SeekNextControl
        Else
            '恢复原值
            'txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱内容)
            zlControl.TxtSelAll txt医嘱内容
            txt医嘱内容.SetFocus: Exit Sub
        End If
    End If
End Sub

Private Sub AdviceCurRowClear()
'功能：清除当前医嘱行的内容，清除后可望而保留当前输入行的新输入状态
    Dim str开始时间 As String
    Dim lngPre As Long
    
    LockWindowUpdate Me.hWnd
    
    '记录之前基本的输入内容
    Call GetRowScope(vsAdvice.Row, lngPre, 0)
    str开始时间 = txt开始时间.Text
    
    '删除行
    Call AdviceDelete(vsAdvice.Row)
    
    '在原位置插入新行
    mblnRowChange = False
    vsAdvice.AddItem "", lngPre
    vsAdvice.Row = lngPre: vsAdvice.Col = col_医嘱内容
    mblnRowChange = True
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    txt开始时间.Text = str开始时间
    txt医嘱内容.SetFocus
    
    LockWindowUpdate 0
End Sub

Private Sub cbo执行时间_GotFocus()
    zlControl.TxtSelAll cbo执行时间
    If vsAdvice.Row < 1 Then Exit Sub
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "周" Or vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "天" Or vsAdvice.TextMatrix(vsAdvice.Row, COL_间隔单位) = "小时" Then
        picHelp.Visible = True
    End If
End Sub

Private Sub txt医嘱内容_Validate(Cancel As Boolean)
    '恢复人为的改变
    If txt医嘱内容.Text <> vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) Then
        txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
    End If
End Sub

Private Sub txt总量_Change()
    With vsAdvice
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_总量)) <> Val(txt总量.Text) Then
                txt总量.Tag = "1"
            End If
        Else
            txt总量.Tag = "1"
        End If
    End With
End Sub

Private Sub txt总量_GotFocus()
    zlControl.TxtSelAll txt总量
End Sub

Private Sub txt总量_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt总量.Text) Or (txt总量.Text = "" And vsAdvice.TextMatrix(vsAdvice.Row, COL_类别) = "K") Then
            Call txt总量_Validate(blnCancel)
            If Not blnCancel Then mblnReturn = True: Call SeekNextControl
        End If
    Else
        If RowIn配方行(vsAdvice.Row) Then
            strMask = "0123456789" '中药配方只能输入整数
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_类别)) > 0 Then
            If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") > 0 _
                And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_可否分零)) = 0 Then
                strMask = "0123456789."
            Else
                strMask = "0123456789"
            End If
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt总量_LostFocus()
    mblnReturn = False
End Sub

Private Sub txt总量_Validate(Cancel As Boolean)
    Dim bln配方行 As Boolean
    Dim strMsg, strTmp As String
    Dim dbl总量 As Double, sng天数 As Single
    Dim blnOutTotal As Boolean '药品超量
    Dim blnTmp As Boolean
    Dim blnTag As Boolean
 
    With vsAdvice
        If Val(txt总量.Text) = 0 Then txt总量.Text = ""
        If mblnReturn Then mblnReturn = False: Exit Sub
        If Not IsNumeric(txt总量.Text) Then
            If txt总量.Text <> "" Then
                Cancel = True: txt总量_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
                '恢复人为的清除：输血临嘱允许不输入总量
                If .TextMatrix(.Row, COL_类别) <> "K" Then
                    If IsNumeric(.TextMatrix(.Row, COL_总量)) Then
                        txt总量.Text = .TextMatrix(.Row, COL_总量)
                    End If
                End If
            End If
        ElseIf CDbl(txt总量.Text) <= 0 Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        ElseIf CDbl(txt总量.Text) > LONG_MAX Then
            Cancel = True: txt总量_GotFocus: Exit Sub
        Else
            txt总量.Text = FormatEx(txt总量.Text, 5)
        End If
        
        bln配方行 = RowIn配方行(.Row)
        If IsNumeric(txt总量.Text) Then
            If bln配方行 Then
                txt总量.Text = CInt(txt总量.Text)
            ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
                If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                    txt总量.Text = IntEx(Val(txt总量.Text))
                ElseIf Val(.TextMatrix(.Row, COL_可否分零)) <> 0 Then
                    txt总量.Text = IntEx(Val(txt总量.Text))
                End If
            ElseIf Val(.TextMatrix(.Row, COL_计算方式)) = 3 Then
                '计次项目总量限制为整数。计次项目不输入单量,因此单量不管
                'txt总量.Text = IntEx(Val(txt总量.Text))
            End If
        End If
        
        '检查总量够否
        If InStr(",4,5,6,", .TextMatrix(.Row, COL_类别)) > 0 Then
            If .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率次数)) <> 0 And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 _
                And Val(txt单量.Text) <> 0 And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                
                sng天数 = Val(txt天数.Text)
                If sng天数 = 0 Then sng天数 = 1
                
                dbl总量 = FormatEx(Calc缺省药品总量( _
                    Val(txt单量.Text), sng天数, _
                    Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), _
                    .TextMatrix(.Row, COL_间隔单位), .TextMatrix(.Row, COL_执行时间), _
                    Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_门诊包装)), _
                    Val(.TextMatrix(.Row, COL_可否分零))), 5)
                    
                If Val(txt总量.Text) < dbl总量 Then
                    If MsgBox(.TextMatrix(.Row, COL_名称) & "按每次 " & _
                        txt单量.Text & .TextMatrix(.Row, COL_单量单位) & "," & _
                        .TextMatrix(.Row, COL_频率) & IIF(mbln天数 And .TextMatrix(.Row, COL_类别) <> "4", ",用药 " & sng天数 & " 天", "") & _
                        "执行时,至少需要 " & FormatEx(dbl总量, 5) & .TextMatrix(.Row, COL_总量单位) & ",要继续吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        Cancel = True: txt总量_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '检查处方限量
        .TextMatrix(.Row, COL_是否超量) = ""
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_处方限量)) <> 0 Then
           dbl总量 = Val(txt总量.Text) * Val(.TextMatrix(.Row, COL_门诊包装)) * Val(.TextMatrix(.Row, COL_剂量系数))
           If dbl总量 > Val(.TextMatrix(.Row, COL_处方限量)) And Val(.TextMatrix(.Row, COL_处方限量)) > 0 Then
               .TextMatrix(.Row, COL_是否超量) = "1"
           End If
    
        ElseIf bln配方行 Then
            txt总量.Tag = "1" '中药配方药品行被隐藏，通过调用AdviceChange间接设文本框可用性
            blnTmp = CheckCHLimited(.Row, Val(txt总量.Text), blnOutTotal, vsAdvice, COL_相关ID, COL_诊疗项目ID, COL_类别, COL_单量)
            If blnOutTotal Then .TextMatrix(.Row, COL_是否超量) = "1"
            
            '同时检查草药的用药是否超期
            If Val(txt总量.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                .TextMatrix(.Row, COL_是否超期) = "1"
            End If
            
        ElseIf InStr(",5,6,7,", .TextMatrix(.Row, COL_类别)) = 0 And Val(.TextMatrix(.Row, COL_处方限量)) > 0 Then
            If Val(txt总量.Text) > Val(.TextMatrix(.Row, COL_处方限量)) And Val(.TextMatrix(.Row, COL_处方限量)) > 0 Then
                .TextMatrix(.Row, COL_是否超量) = "1"
            End If
        End If
                
        '最大金额提示
        If gcurMaxMoney > 0 Then
            If .TextMatrix(.Row, COL_单价) = "" Then .TextMatrix(.Row, COL_单价) = GetItemPrice(.Row) '用""不是零
            If Val(.TextMatrix(.Row, COL_单价)) * Val(txt总量.Text) > gcurMaxMoney Then
                If MsgBox("当前医嘱 " & txt总量.Text & lbl总量单位.Caption & " 的金额达到了：" & Format(Val(.TextMatrix(.Row, COL_单价)) * Val(txt总量.Text), "0.00") & "，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt总量_GotFocus: Exit Sub
                End If
            End If
        End If
        
        '反算天数（没有输入天数才反算，中药不会显示天数）
        If mbln天数反算 And txt总量.Tag <> "" Then
            If .TextMatrix(.Row, COL_频率) <> "" And Val(.TextMatrix(.Row, COL_频率次数)) <> 0 And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 _
                And Val(txt单量.Text) <> 0 And Val(.TextMatrix(.Row, COL_频率次数)) <> 0 And Val(.TextMatrix(.Row, COL_频率间隔)) <> 0 _
                And Val(.TextMatrix(.Row, COL_剂量系数)) <> 0 And Val(.TextMatrix(.Row, COL_门诊包装)) <> 0 Then
                
                sng天数 = Calc缺省药品天数(Val(txt总量.Text), Val(txt单量.Text), _
                    Val(.TextMatrix(.Row, COL_频率次数)), Val(.TextMatrix(.Row, COL_频率间隔)), .TextMatrix(.Row, COL_间隔单位), _
                    Val(.TextMatrix(.Row, COL_剂量系数)), Val(.TextMatrix(.Row, COL_门诊包装)), _
                    Val(.TextMatrix(.Row, COL_可否分零)))
                
                Call CheckDrugOutOfRange(.Row, sng天数)
                If sng天数 = 0 Then sng天数 = 1
                msngPre天数 = sng天数
                If sng天数 <> Val(txt天数.Text) Then
                    txt天数.Text = sng天数
                    txt天数.Tag = "1"
                    msng天数 = sng天数
                End If
            End If
        End If
        
        '更新数据
        blnTag = (txt总量.Tag <> "")
        Call AdviceChange
        '医保管控实时监测：首次输入(经过)或者更改时检查
        If mint险类 <> 0 And (.Cell(flexcpData, .Row, COL_状态) = 0 Or blnTag) Then
            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                If MakePriceRecord(.Row) Then
                    If Not gclsInsure.CheckItem(mint险类, 0, 0, mrsPrice) Then
                        Cancel = True: txt总量_GotFocus: Exit Sub
                    End If
                End If
                '标记为已经作了检查
                .Cell(flexcpData, .Row, COL_状态) = 1
            End If
        End If

        Call CalcAdviceMoney '显示新开医嘱金额
        
        '药品库存检查:只提醒,修改了才提醒
        If InStr(",5,6,", .TextMatrix(.Row, COL_类别)) > 0 Or bln配方行 _
            Or .TextMatrix(.Row, COL_类别) = "4" And Val(.TextMatrix(.Row, COL_跟踪在用)) = 1 Then
            strMsg = CheckStock(.Row)
            If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
        End If
        
    End With
End Sub
 
Private Sub ClearAdviceCard()
'功能：清除医嘱显示卡片相关的内容
'参数：bln开始时间=是否清除开始时间
    Call SetCardEditable(True)
    
    txt开始时间.Text = ""
    txt安排时间.Text = ""
    txt医嘱内容.Text = ""
    cbo医生嘱托.Text = ""
    cbo执行科室.Clear
    cbo附加执行.Clear
    chk紧急.value = 0
    chk紧急.Visible = False '输医嘱内容后才可用
    cbo滴速.Text = ""
    txt超量说明.Text = ""
    txt用药理由.Text = ""
    
    mblnDoCheck = False
    chk紧急.value = 0
    chkZeroBilling.value = 0
    mblnDoCheck = True
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1)
    Call SetStartTime(True)
    
    stbThis.Panels(3).Text = ""
    stbThis.Panels(4).Text = ""
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'功能：用颜色标识当前医嘱是否可以编辑
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox;OptionButton", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
    cmdSel.Enabled = fraAdvice.Enabled
    cmd常用嘱托.Enabled = fraAdvice.Enabled
    cmd医生嘱托.Enabled = fraAdvice.Enabled
End Sub

Private Function Get频率范围(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn配方行(lngRow) Then
            Get频率范围 = 2 '中医
        Else
            If RowIn检验行(lngRow) Then '以检验项目行为准
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Or Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                Get频率范围 = 1 '成药或可选频率的项目使用西医频率项目
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 1 Then
                Get频率范围 = -1 '一次性
            ElseIf Val(.TextMatrix(lngRow, COL_频率性质)) = 2 Then
                Get频率范围 = -2 '持续性
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'功能：当前行为隐藏行时，定位到它所属的可见行
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_类别)) > 0 And Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_相关ID))))
        ElseIf .TextMatrix(.Row, COL_类别) = "E" And Val(.TextMatrix(.Row, COL_相关ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub FuncApplyCustom(ByVal intType As Integer, ByVal lng文件ID As Long, Optional ByVal lng申请序号 As Long, Optional ByVal lng项目id As Long)
'功能：自定义申请单
    Dim objApplyCustom As New frmApplyCustom
    Dim lngOut医嘱ID As Long

    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    If objApplyCustom.ShowMe(Me, 1, intType, mlng病人ID, mstr挂号单, 1, lng文件ID, lng申请序号, mlng病人科室id, mlng病人科室id, , mrsDefine, , , 1, mclsMipModule, mlng前提ID, , mint险类, lngOut医嘱ID, lng项目id) Then
         '重新读取显示医嘱
        Call ReLoadAdvice(lngOut医嘱ID)
        mblnOK = True '强行
        If txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        Else
            vsAdvice.SetFocus
        End If
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能：当行改变时，更新卡片内容
    Dim rsItem As New ADODB.Recordset
    Dim rsBlood As New ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    Dim lng用法ID As Long, blnEditable As Boolean
    Dim lng药品ID As Long, lngBaseRow As Long '中药配方的第一味组成药行
    Dim dblPrice As Double, strTmp As String, i As Long
    Dim bln显示单量 As Boolean, bln显示抗菌药物 As Boolean
    Dim blnEditableTmp As Boolean
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_开始时间)
    End If
    
    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub
    'Pass
    If mblnPass And Me.Visible Then
        If NewRow <> OldRow Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassSetDrug(mobjPassMap)
            End If
        End If
    End If
    
    lngRow = NewRow
     '当前行是空行时，如果前一行是一并给药行，则缺省按下“一并”按钮
    If vsAdvice.RowData(lngRow) = 0 Then
        i = GetPreRow(lngRow)
        If i = -1 Then
            mblnRowMerge = False
        Else
            mblnRowMerge = RowIn一并给药(i)
        End If
    Else
        mblnRowMerge = RowIn一并给药(lngRow)
    End If
    cbsMain.RecalcLayout '即时刷新
        
    '显示或隐藏审核疑问说明
    Call ShowOrHideQuestion
        
    Me.Refresh
    LockWindowUpdate Me.hWnd
    
    chk免试.Visible = False
        
    On Error GoTo errH
    
    With vsAdvice
        If Val(.RowData(lngRow)) = 0 Then
            '无效行清除卡片内容
            Call ClearAdviceCard
            
            '缺省开始时间
            Call txt开始时间_GotFocus
        Else
            '卡片编辑
            blnEditable = True
            '已发送的医嘱不能修改
            If Val(.TextMatrix(lngRow, COL_状态)) <> 1 Then blnEditable = False
            '已签名的医嘱不可修改
            If Val(.TextMatrix(lngRow, COL_签名否)) = 1 Then blnEditable = False
            
            If .TextMatrix(lngRow, COL_类别) = "K" And gbln血库系统 Then
                '血库环节
                If Not (Val(.TextMatrix(lngRow, COL_检查方法)) = 1 And Val(.TextMatrix(lngRow, COL_审核状态)) = 2) Then
                    '旧程 5－血库配血中
                    If Val(.TextMatrix(lngRow, COL_审核状态)) = 5 Or Val(.TextMatrix(lngRow, COL_审核状态)) = 2 Then blnEditable = False
                End If
            Else
                '审核通过的不允许修改
                If Val(.TextMatrix(lngRow, COL_审核状态)) = 2 Then blnEditable = False
            End If
            
            '是输血医嘱时，新开后直接进入4这个状态，这时候是允许编辑的，否则禁止
            If Val(.TextMatrix(lngRow, COL_审核状态)) = 4 And .TextMatrix(lngRow, COL_类别) = "K" Then
                If CanEditBloodAdvice(Val(.RowData(lngRow)), Val(.TextMatrix(lngRow, COL_审核状态)), Val(.TextMatrix(lngRow, COL_标志)) = 1, Val(.TextMatrix(lngRow, COL_检查方法)) = 1, False) = False Then blnEditable = False
            End If
            
            '显示或隐藏免试标记
            If Val(.TextMatrix(lngRow, COL_操作类型)) = 1 And .TextMatrix(lngRow, COL_类别) = "E" Then chk免试.Visible = True
            mblnDoCheck = False
            chk免试.value = Val(.TextMatrix(lngRow, COL_免试))
            mblnDoCheck = True
            Call SetCardEditable(blnEditable)
            
            '获取诊疗项目基本信息
            '---------------------
            If InStr(",4,5,6,7,", Val(.TextMatrix(lngRow, COL_类别))) > 0 Then
                lng药品ID = Val(.TextMatrix(lngRow, COL_收费细目ID))
            End If
            
            
            If RowIn配方行(lngRow) Then
                txt总量.MaxLength = 3
                '获取中药配方第一味中药行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                lng药品ID = Val(.TextMatrix(lngBaseRow, COL_收费细目ID))
            ElseIf RowIn检验行(lngRow) Then
                '获取一并采样的第一个项目行
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
                txt总量.MaxLength = txt单量.MaxLength
            Else
                lngBaseRow = lngRow
                txt总量.MaxLength = txt单量.MaxLength
            End If
            Set rsItem = Get诊疗项目记录(Val(.TextMatrix(lngBaseRow, COL_诊疗项目ID)))
            
            '扩展按钮可用状态(检查组合,检验组合,手术,中药配方)
            cmdExt.Enabled = InStr(",7,C,F,D,", rsItem!类别) > 0
            If rsItem!类别 = "E" Or rsItem!类别 = "K" Or rsItem!类别 = "Z" Then
                If rsItem!类别 = "K" And Val(.TextMatrix(lngRow, COL_申请序号) & "") <> 0 Then
                    cmdExt.Enabled = True
                Else
                    cmdExt.Enabled = CheckApplication(Val(.TextMatrix(lngBaseRow, COL_诊疗项目ID)), 1)
                End If
            End If
            
            '显示当前医嘱卡片内容
            '--------------------------------------------------------------------------------------------
            '开始时间：只有新增医嘱时可以修改开始时间
            txt开始时间.Text = .Cell(flexcpData, lngRow, COL_开始时间)
            Call SetStartTime(.TextMatrix(lngRow, COL_EDIT) = "1")
            
            '医嘱内容
            txt医嘱内容.Text = .TextMatrix(lngRow, col_医嘱内容)
            
            
            '超量说明
            txt超量说明.Text = .TextMatrix(lngRow, COL_超量说明)
            SetItemEditable , , , , , , , , , , , , IIF(.TextMatrix(lngRow, COL_是否超量) = "1" Or .TextMatrix(lngRow, COL_是否超期) = "1", 1, -1)
            
            If txt超量说明.Text <> "" Then
                cmdExcReason.Enabled = blnEditable
                cmdComExcReason.Enabled = blnEditable
            End If
                        
            '单量：临嘱,成药或可选择频率的计时,计量项目可以录入
            '----------------------
            If rsItem!类别 = "7" Then '中药配方(中草药)虽然有单量,但不在这里填写
                SetItemEditable -1
            ElseIf (Nvl(rsItem!执行频率, 0) = 0 And InStr(",1,2,", Nvl(rsItem!计算方式, 0)) > 0) _
                    Or InStr(",5,6,", rsItem!类别) > 0 Then
                SetItemEditable 1
                bln显示单量 = True
                txt单量.Text = .TextMatrix(lngRow, COL_单量)
                lbl单量单位.Caption = .TextMatrix(lngRow, COL_单量单位)
            Else
                SetItemEditable -1
            End If
            
            '天数：西药，中成药临嘱才使用，用于计算总量
            '一般：临嘱的药品(非中药)或可选择频率的计时,计量项目可以使用天数来自动计算总量
            blnEditableTmp = False
            If InStr(",5,6,", rsItem!类别) > 0 Then
                If mbln天数 Then blnEditableTmp = True
            End If
            If blnEditableTmp Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt天数.Text = Val(.TextMatrix(lngRow, COL_天数))
            If Val(txt天数.Text) = 0 Then txt天数.Text = ""
            
            '总量
            '--------------------
            If rsItem!类别 = "7" Then
                '中药配方(中草药)填写为付数
                SetItemEditable , 1
                lbl总量单位.Caption = "付"
                txt总量.Text = .TextMatrix(lngRow, COL_总量) '付数
                If Val(txt总量.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                    SetItemEditable , , , , , , , , , , , , 1
                End If
                bln显示单量 = True
                
                '非散装形态，只允许在配方界面输付数
                If Val(.TextMatrix(lngRow, COL_中药形态)) <> 0 Then
                    txt总量.Enabled = False
                    txt总量.BackColor = Me.BackColor
                End If
            Else
                '临嘱都需要填写总量:临嘱发送以总量为准
                If rsItem!类别 = "Z" And Nvl(rsItem!操作类型) <> "0" Then
                    SetItemEditable , -1 '特殊医嘱不允许修改总量(固定为1次)
                ElseIf Nvl(rsItem!执行频率, 0) = 1 And Nvl(rsItem!计算方式, 0) = 3 Then
                    SetItemEditable , -1 '一次性计次项目不输入总量
                Else
                    SetItemEditable , 1
                    bln显示单量 = True
                End If
                lbl总量单位.Caption = .TextMatrix(lngRow, COL_总量单位)
                txt总量.Text = .TextMatrix(lngRow, COL_总量)
            End If
            
            '给药途径和中药用法
            '--------------
            If InStr(",5,6,", rsItem!类别) > 0 Then
                SetItemEditable , , 1
                lbl用法.Caption = "给药途径"
                '查找给药途径对应的行:查找的Rowdata(Variant)数据要转为Long型,才能精确匹配
                lng用法ID = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            ElseIf rsItem!类别 = "K" Then
                '输血医嘱：要兼容以前没有输血途径的情况
                lng用法ID = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                If lng用法ID <> -1 Then
                    SetItemEditable , , 1
                     If Val(.TextMatrix(lngRow, COL_检查方法)) = 0 And gbln血库系统 = True Then
                        lbl用法.Caption = "采集方法"
                    Else
                        lbl用法.Caption = "输血途径"
                    End If
                    
                    lng用法ID = Val(.TextMatrix(lng用法ID, COL_诊疗项目ID))
                    cmd用法.Tag = lng用法ID
                    txt用法.Text = Get项目名称(lng用法ID)
                Else
                    SetItemEditable , , -1
                End If
            ElseIf rsItem!类别 = "7" Then
                SetItemEditable , , 1
                lbl用法.Caption = "中药用法"
                
                '中药配方显示行就是中药用法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            ElseIf RowIn检验行(lngRow) Then '不用类别判断,兼容以前的检验
                '检验组合
                SetItemEditable , , 1
                lbl用法.Caption = "采集方法"
                
                '检验组合显示行就是采集方法行
                lng用法ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                cmd用法.Tag = lng用法ID
                txt用法.Text = Get项目名称(lng用法ID)
            Else
                SetItemEditable , , -1
            End If
            
            '手术时间/输血时间：只有手术/输血可用(隐藏在用法位置)
            If rsItem!类别 = "F" Or rsItem!类别 = "K" Then
                SetItemEditable , , , , , , , , 1
                If IsDate(.TextMatrix(lngRow, COL_手术时间)) Then
                    txt安排时间.Text = .TextMatrix(lngRow, COL_手术时间)
                Else
                    txt安排时间.Text = .Cell(flexcpData, lngRow, COL_开始时间)
                End If
                Call Set安排时间(rsItem!类别)
            Else
                SetItemEditable , , , , , , , , -1
            End If
            
            '频率(临嘱输入用于指导使用)
            If InStr("F,G,H,I", rsItem!类别) > 0 Or rsItem!类别 = "Z" And InStr(",1,2,3,4,5,6,7,8,9,10,11,12,14,", "," & rsItem!操作类型 & ",") > 0 Then
                SetItemEditable , , , -1
            Else
                SetItemEditable , , , 1
            End If
            cmd频率.Tag = .TextMatrix(lngRow, COL_频率)
            txt频率.Text = .TextMatrix(lngRow, COL_频率)
                    
            '执行时间："可选频率"或药品。
            If (Nvl(rsItem!执行频率, 0) = 0 Or InStr(",5,6,7,", rsItem!类别) > 0) And .TextMatrix(lngRow, COL_间隔单位) <> "分钟" Then
                SetItemEditable , , , , 1
                Call Get时间方案(cbo执行时间, Get频率范围(lngRow), .TextMatrix(lngRow, COL_频率), lng用法ID)
                cbo执行时间.Text = .TextMatrix(lngRow, COL_执行时间)
            Else
                SetItemEditable , , , , -1
            End If
                    
            '滴速：输液类给药途径的药品可以输入
            If InStr(",5,6,", rsItem!类别) > 0 Then
                i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                If Val(.TextMatrix(i, COL_执行分类)) = 1 Then
                    SetItemEditable , , , , , , , , , 1
                    If InStr(.TextMatrix(i, COL_医生嘱托), "滴/分钟") > 0 Then
                        lbl滴速单位.Caption = "滴/分钟"
                    ElseIf InStr(.TextMatrix(i, COL_医生嘱托), "毫升/小时") > 0 Then
                        lbl滴速单位.Caption = "毫升/小时"
                    End If
                    Call Load输液滴速(cbo滴速, lbl滴速单位, False)
                    cbo滴速.Text = Replace(.TextMatrix(i, COL_医生嘱托), lbl滴速单位.Caption, "")
                Else
                    SetItemEditable , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , -1
            End If
            
            
            '用药目的和用药理由
            If Val(.TextMatrix(lngRow, COL_抗菌等级)) = 0 Then
                SetItemEditable , , , , , , , , , , -1, -1
            Else
                SetItemEditable , , , , , , , , , , 1, 1
                bln显示抗菌药物 = True
                
                If .TextMatrix(lngRow, COL_用药目的) = "1" Then
                    Call zlControl.CboSetIndex(cboDruPur.hWnd, 1)
                ElseIf .TextMatrix(lngRow, COL_用药目的) = "2" Then
                    Call zlControl.CboSetIndex(cboDruPur.hWnd, 2)
                End If
                
                txt用药理由.Text = .TextMatrix(lngRow, COL_用药理由)
            End If
            cboDruPur.Enabled = blnEditable
            cmdReason.Enabled = blnEditable
            cmd收藏用药理由.Enabled = blnEditable
            
            '输血医嘱显示用药目的，将用药目的显示为输血原因
            If (gbln输血分级管理 Or gbln血库系统) And .TextMatrix(lngRow, COL_类别) = "K" Then
                bln显示抗菌药物 = True
                SetItemEditable , , , , , , , , , , , , , 1
                txt用药理由.Width = cmdExcReason.Left + cmdExcReason.Width - txt用药理由.Left
                If .TextMatrix(lngRow, COL_标志) = "1" Then
                    SetItemEditable , , , , , , , , , , , 1
                    txt用药理由.Text = .TextMatrix(lngRow, COL_用药理由)
                Else
                    SetItemEditable , , , , , , , , , , , -1
                End If
            Else
                SetItemEditable , , , , , , , , , , , , , -1
                txt用药理由.Width = cmdExcReason.Left + cmdExcReason.Width - txt用药理由.Left
                cmd收藏用药理由.Left = txt用药理由.Left + txt用药理由.Width + 30
                cmdReason.Left = txt用药理由.Left + txt用药理由.Width - cmdReason.Width
            End If
            
            '医生嘱托
            cbo医生嘱托.Text = .TextMatrix(lngRow, COL_医生嘱托)
                    
            '执行性质
            If InStr(",5,6,7,", rsItem!类别) > 0 Then
                '如果是自管药则固定选择自备药
                If Val(.TextMatrix(lngRow, COL_临床自管药)) = 1 And InStr(",5,6,", rsItem!类别) > 0 Then
                    strTmp = "自备药"
                Else
                    If rsItem!类别 = "7" Then
                        '对于中药配方,根据诊疗项目管理中限制及本程序处理,不可能用法和煎法一个为院外执行,一个不为
                        If Val(.TextMatrix(lngBaseRow, COL_执行性质)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 Then
                            strTmp = "自备药"
                        ElseIf Val(.TextMatrix(lngBaseRow, COL_执行性质)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                            strTmp = "离院带药"
                        Else
                            strTmp = "正常"
                        End If
                    Else
                        i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                        If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                            strTmp = "自备药"
                        ElseIf Val(.TextMatrix(lngRow, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            strTmp = "离院带药"
                        Else
                            strTmp = "正常"
                        End If
                    End If
                End If
                Call SetCbo执行性质(gbln抗菌药物使用自备药 Or Not gblnKSSStrict Or Val(.TextMatrix(lngRow, COL_抗菌等级)) = 0, Val(.TextMatrix(lngRow, COL_临床自管药)) = 1 And rsItem!类别 & "" <> "7")
                SetItemEditable , , , , , , 1
                Call SeekIndex(cbo执行性质, strTmp)
            Else
                SetItemEditable , , , , , , -1
            End If
            
            lbl执行科室.Caption = "执行科室"
            '执行科室:留观或住院医嘱用临床科室
            If rsItem!类别 = "Z" And InStr(",1,2,", Nvl(rsItem!操作类型, 0)) > 0 Then
                SetItemEditable , , , , , 1
                If Nvl(rsItem!操作类型, 0) = 1 Then
                    lbl执行科室.Caption = "留观科室"
                    '留观:包含门诊或住院临床科室,由服务对象决定是门诊留观或住院留观
                    Call Get临床科室(3, , Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室, True, False, True)
                ElseIf Nvl(rsItem!操作类型, 0) = 2 Then
                    lbl执行科室.Caption = "住院科室"
                    '住院:包含住院临床科室
                    Call Get临床科室(2, , Val(.TextMatrix(lngRow, COL_执行科室ID)), cbo执行科室, True, False, True, 1)
                End If
                If Val(.TextMatrix(lngRow, COL_执行科室ID)) <> 0 And cbo执行科室.ListIndex = -1 Then .TextMatrix(lngRow, COL_执行科室ID) = 0
            Else
                '是药品则以药品行为准显示,检验组合以检验项目为准显示
                i = lngRow
                If rsItem!类别 = "7" Then
                    i = lngBaseRow
                ElseIf RowIn检验行(lngRow) Then '不用类别判断,兼容以前的检验
                    i = lngBaseRow
                End If
                
                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                    '非叮嘱和院外执行时才显示和可以选择(包括药品)
                    SetItemEditable , , , , , 1
                    Call Get诊疗执行科室(mlng病人ID, 0, cbo执行科室, rsItem!类别, rsItem!ID, lng药品ID, Nvl(rsItem!执行科室, 0), _
                        mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), Val(.TextMatrix(i, COL_执行科室ID)), 1, 1, , blnEditable)
                        
                     
                    '非散装形态，只允许在配方界面选药房
                    If rsItem!类别 = "7" Then
                        If Val(.TextMatrix(lngRow, COL_中药形态)) <> 0 Then
                            cbo执行科室.Enabled = False
                            cbo执行科室.BackColor = Me.BackColor
                        End If
                    End If
                ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                    SetItemEditable , , , , , -1
                    If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                        cbo执行科室.AddItem "<无执行叮嘱>"
                    Else
                        cbo执行科室.AddItem "-"
                    End If
                    Call zlControl.CboSetIndex(cbo执行科室.hWnd, 0)
                End If
                If Val(.TextMatrix(i, COL_执行科室ID)) <> 0 And cbo执行科室.ListIndex = -1 Then .TextMatrix(i, COL_执行科室ID) = 0
                If InStr("5,6,7", rsItem!类别) > 0 Then lbl执行科室.Caption = "发药药房"
            End If

            If cbo执行科室.ListIndex = -1 And cbo执行科室.ListCount = 1 Then
                If Val(.TextMatrix(i, COL_状态)) < 3 Then
                    cbo执行科室.ListIndex = 0
                Else
                    Call zlControl.CboSetIndex(cbo执行科室.hWnd, 0)
                End If
            End If
            
            If cbo执行科室.ListCount = 1 Then
                If cbo执行科室.List(cbo执行科室.ListIndex) <> "[其它...]" Then
                    cbo执行科室.Enabled = False
                Else
                    cbo执行科室.Enabled = True
                End If
            Else
                cbo执行科室.Enabled = True
            End If
            
            '附加执行:指给药途径,中药用法,手术麻醉,采集方式的执行科室，原液皮试项目
            If Should附加执行(lngRow, i, strTmp) Then
                If .TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" Then
                    '对于原液皮试加载药房
                    lng药品ID = Get原液皮试药品(Val(.TextMatrix(i, COL_诊疗项目ID)))
                    If lng药品ID <> 0 Then
                        SetItemEditable , , , , , , , 1
                        Call Get诊疗执行科室(mlng病人ID, 0, cbo附加执行, "5", 0, lng药品ID, 0, mlng病人科室id, 0, Val(.TextMatrix(i, COL_用药理由)), 1, 1, , blnEditable)
                        '如果没有选药房则不加载任何默认值
                        If Val(.TextMatrix(i, COL_用药理由)) = 0 Then cbo附加执行.ListIndex = -1
                    Else
                        cbo附加执行.Clear
                        SetItemEditable , , , , , , , -1
                    End If
                Else
                    '执行科室:非药嘱药品及跟踪卫材的专门取
                    SetItemEditable , , , , , , , 1
                    Call Get诊疗执行科室(mlng病人ID, 0, cbo附加执行, .TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), lng药品ID, _
                        Val(.TextMatrix(i, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(i, COL_开嘱科室ID)), Val(.TextMatrix(i, COL_执行科室ID)), 1, 1, , blnEditable)
                        
                    If Val(.TextMatrix(i, COL_执行科室ID)) <> 0 And cbo附加执行.ListIndex = -1 Then .TextMatrix(i, COL_执行科室ID) = 0
                    If cbo附加执行.ListIndex = -1 And cbo附加执行.ListCount = 1 Then cbo附加执行.ListIndex = 0
                End If
            Else
                SetItemEditable , , , , , , , -1
                If i <> -1 Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                        If Val(.TextMatrix(i, COL_执行性质)) = 0 Then
                            cbo附加执行.AddItem "<无执行叮嘱>"
                        ElseIf Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            cbo附加执行.AddItem "-"
                        End If
                        Call zlControl.CboSetIndex(cbo附加执行.hWnd, 0)
                    End If
                End If
            End If
            lbl附加执行.Caption = strTmp
            
            '紧急标志
            chk紧急.Visible = True
            mblnDoCheck = False
            chk紧急.value = Val(.TextMatrix(lngRow, COL_标志))
            chkZeroBilling.value = Val(.TextMatrix(lngRow, COL_零费记帐))
            mblnDoCheck = True
                        
            
            '显示药品库存：以门诊单位，中药配方不显示
            '----------------------------------------
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 _
                And (InStr(",5,6,", rsItem!类别) > 0 Or rsItem!类别 = "4" And Val(.TextMatrix(lngRow, COL_跟踪在用)) = 1) Then
                If .TextMatrix(lngRow, COL_库存) = "" And Val(.TextMatrix(lngRow, COL_执行科室ID)) <> 0 Then Call GetDrugStock(lngRow)
                If .TextMatrix(lngRow, COL_库存) <> "" And Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then .TextMatrix(lngRow, COL_库存) = ""
                If .TextMatrix(lngRow, COL_库存) <> "" Then
                    If InStr(GetInsidePrivs(p门诊医嘱下达), "显示药品库存") = 0 Then
                        stbThis.Panels(3).Text = IIF(Val(.TextMatrix(lngRow, COL_库存)) > 0, "有库存", "无库存")
                    Else
                        stbThis.Panels(3).Text = "库存:" & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & .TextMatrix(lngRow, COL_门诊单位)
                    End If
                Else
                    stbThis.Panels(3).Text = ""
                End If
            Else
                If rsItem!类别 = "7" And Val(.TextMatrix(lngRow, COL_状态)) = 1 Then
                    Call GetDrugStock(lngRow)
                End If
                stbThis.Panels(3).Text = ""
            End If
            
            '显示医嘱单价和费用类型
            If .TextMatrix(lngRow, COL_单价) = "" Then '用""不是零
                .TextMatrix(lngRow, COL_单价) = GetItemPrice(lngRow)
            End If
            dblPrice = Val(.TextMatrix(lngRow, COL_单价))
            If dblPrice <> 0 Then
                If InStr(",4,5,6,", rsItem!类别) > 0 Then
                    stbThis.Panels(4).Text = "每" & .TextMatrix(lngRow, COL_门诊单位) & ":" & FormatEx(dblPrice, 5) & "元"
                ElseIf rsItem!类别 = "7" Then
                    stbThis.Panels(4).Text = "每付:" & FormatEx(dblPrice, 5) & "元"
                Else
                    stbThis.Panels(4).Text = IIF(IsNull(rsItem!计算单位), "价格:", "每" & Nvl(rsItem!计算单位) & ":") & FormatEx(dblPrice, 5) & "元"
                End If
            Else
                stbThis.Panels(4).Text = ""
            End If
            
            '显示费用类型
            strTmp = Get费用类型(lngRow)
            If strTmp <> "" Then
                stbThis.Panels(4).Text = stbThis.Panels(4).Text & IIF(stbThis.Panels(4).Text <> "", ",", "") & strTmp
            End If
            
            '待审核的用血医嘱，则不允许:输血成分、执行科室、预定输血量
            If .TextMatrix(lngRow, COL_类别) = "K" And Val(.TextMatrix(lngRow, COL_检查方法)) = 1 And gbln血库系统 = True And blnEditable Then
                If InitObjBlood = True Then
                    If gobjPublicBlood.GetPrepareBloodRs(Val(.RowData(lngRow)), rsBlood) = True Then
                        If Val(rsBlood!记录性质 & "") = 2 And Val(rsBlood!记录状态 & "") = 1 Then
                            cmdSel.Enabled = False
                            txt医嘱内容.Enabled = False
                            txt总量.Enabled = False
                            cbo执行科室.Enabled = False
                        End If
                    End If
                End If
            End If
            
        End If
    End With
    
    '清除编辑标志
    Call ClearItemTag
    
    '显示或隐藏药品相关的两行项目
    Call SetMediInfoItem(bln显示单量, bln显示抗菌药物)
    
    '显示计价窗体
    Call ShowPrice(lngRow)
    
    '调用外挂接口
    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
        If OldRow <> NewRow Then
            Call zlPluginAdviceRowChange(NewRow)
            With vsAdvice
                If OldRow <> -1 Then
                    If Val(.RowData(OldRow)) <> 0 Then
                        If Val(.TextMatrix(OldRow, COL_EDIT)) = 1 Or Val(.TextMatrix(OldRow, COL_EDIT)) = 2 Then
                            Call zlPluginAdviceRowChange(OldRow, 1)
                        End If
                    End If
                End If
            End With
        End If
    End If
    
    cbsMain.RecalcLayout '即时刷新,有Lock可不要
    LockWindowUpdate 0
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'根据当前行的情况显示计价窗体
    If mblnModal Then Exit Sub
    
    If vsAdvice.RowData(lngRow) = 0 Or Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_状态))) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",4,5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf RowIn配方行(lngRow) Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf stbThis.Panels("Price").Bevel = sbrNoBevel Then
        stbThis.Panels("Price").Visible = True
        If Val(stbThis.Panels("Price").Tag) <> 0 Then
            stbThis.Panels("Price").Bevel = sbrInset
        Else
            stbThis.Panels("Price").Bevel = sbrRaised
        End If
    End If
    
    If stbThis.Panels("Price").Bevel <> sbrInset Then
        '关闭计价窗体
        mfrmPrice.HideMe
    Else
        Call mfrmPrice.ShowMe(Me, vsAdvice, mlng病人ID, 0, mlng病人科室id, 1, mint险类, _
            COL_序号 & "," & COL_相关ID & "," & COL_状态 & "," & COL_类别 & "," & COL_诊疗项目ID & "," & _
            COL_收费细目ID & "," & COL_标本部位 & "," & COL_检查方法 & "," & COL_执行标记 & "," & COL_计价性质 & "," & COL_执行性质 & "," & COL_执行科室ID)
    End If
End Sub

Private Function Get费用类型(ByVal lngRow As Long) As String
'功能：获取指定行的费用类型
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str类型 As String, str大类 As String, lng收费细目ID As Long
    
    lng收费细目ID = Val(vsAdvice.TextMatrix(lngRow, COL_收费细目ID))
    If lng收费细目ID <> 0 Then
        '取医保的费用类型
        If mint险类 <> 0 Then
            str类型 = gclsInsure.GetItemInsure(mlng病人ID, lng收费细目ID, 0, True, mint险类)
            If str类型 <> "" Then
                If UBound(Split(str类型, ";")) >= 5 Then
                    str类型 = Split(str类型, ";")(5)
                Else
                    str类型 = ""
                End If
            End If
        End If
        '没有则取HIS的费用类型
        strSQL = "Select A.费用类型,N.名称 as 医保大类 From 收费项目目录 A,保险支付项目 M,保险支付大类 N" & _
            " Where A.ID=[1] And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng收费细目ID, mint险类)
        If Not rsTmp.EOF Then
            If str类型 = "" Then str类型 = Nvl(rsTmp!费用类型)
            str大类 = Nvl(rsTmp!医保大类)
        End If
    End If
        
    Get费用类型 = Mid(IIF(str类型 <> "", ",类型:" & str类型, "") & IIF(str大类 <> "", ",大类:" & str大类, ""), 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Should附加执行(ByVal lngRow As Long, lngRow2 As Long, str执行科室 As String) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
'参数：lngRow2=返回附加行的医嘱行号
'      str执行科室=附加执行科室类型
    Dim i As Long
    
    lngRow2 = -1
    str执行科室 = "附加执行"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function
        
        If RowIn配方行(lngRow) Then
            '中药用法
            lngRow2 = lngRow
            str执行科室 = "用法执行"
            Should附加执行 = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            '给药途径
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
            str执行科室 = "给药执行"
            Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
            '手术麻醉
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            str执行科室 = "麻醉执行"
            If lngRow2 <> -1 Then Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "K" Then
            '输血途径
            If Val(.TextMatrix(lngRow, COL_检查方法)) = 0 And gbln血库系统 = True Then
                str执行科室 = "采集执行"
            Else
                str执行科室 = "输血执行"
            End If
            
            lngRow2 = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
            If lngRow2 <> -1 Then Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
            And .TextMatrix(lngRow - 1, COL_类别) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
            '采集方式
            lngRow2 = lngRow
            str执行科室 = "采集执行"
            Should附加执行 = True
        ElseIf .TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" Then
            '原液皮试药房
            lngRow2 = lngRow
            str执行科室 = "原液药房"
            Should附加执行 = True
        End If
        
        '叮嘱或院外执行
        If Should附加执行 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_执行性质))) > 0 Then
                Should附加执行 = False
            End If
        End If
    End With
End Function

Private Function GetItemPrice(ByVal lngRow As Long) As Double
'功能：获取当前医嘱行的价格(药品为一个药房包装的单价,其它根据收费对照)
'说明：药品不包含给药途径及中药用法煎法
    Dim rsTmp As New ADODB.Recordset
    Dim str医嘱IDs As String, str单量s As String, str诊疗收费 As String
    Dim str项目IDs As String, str医嘱 As String, str项目科室 As String, strTmp As String
    Dim strAdviceIDs As String, lng执行科室ID As Long
    Dim dblPrice As Double, dbl数量 As Double
    Dim bln药品 As Boolean, strSQL As String
    Dim str管码项目 As String, i As Long
    
    With vsAdvice
        bln药品 = True
        If InStr(",4,5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
            '西药及中成药按规格下才能计算价格
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                str项目IDs = str项目IDs & "," & Val(.TextMatrix(lngRow, COL_收费细目ID))
            End If
            lng执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        ElseIf RowIn配方行(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" And Val(.TextMatrix(i, COL_收费细目ID)) <> 0 Then
                        If lng执行科室ID = 0 Then
                            lng执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                        End If
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COL_收费细目ID))
                        str单量s = str单量s & ";" & Val(.TextMatrix(i, COL_单量))
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            bln药品 = False
            '其它医嘱,未校对(计价)的按收费对照计算,否则直接取医嘱计价
            '不包含不计价和手工计价的项目,不包含叮嘱和院外执行的项目
            If Val(.TextMatrix(lngRow, COL_计价性质)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                If InStr(",1,2,", .TextMatrix(lngRow, COL_状态)) > 0 Then
                    str管码项目 = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                    If RowIn检验行(lngRow) Then
                        i = .FindRow(CStr(.RowData(lngRow)), .FixedRows, COL_相关ID)
                        If i <> -1 Then
                            str管码项目 = Val(.TextMatrix(i, COL_诊疗项目ID))
                        End If
                    End If
                    
                    str项目科室 = str项目科室 & "," & Val(.TextMatrix(lngRow, COL_诊疗项目ID)) & ":" & Val(.TextMatrix(lngRow, COL_执行科室ID))
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                    str医嘱 = str医嘱 & " Union ALL Select " & _
                        IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, "-NULL", Val(.TextMatrix(lngRow, COL_相关ID))) & " as 相关ID," & _
                        Val(.TextMatrix(lngRow, COL_诊疗项目ID)) & " as 诊疗项目ID," & Val(.TextMatrix(lngRow, COL_执行标记)) & " as 执行标记," & _
                        IIF(.TextMatrix(lngRow, COL_标本部位) = "", "NULL", "'" & .TextMatrix(lngRow, COL_标本部位) & "'") & " as 标本部位," & _
                        IIF(.TextMatrix(lngRow, COL_检查方法) = "", "NULL", "'" & .TextMatrix(lngRow, COL_检查方法) & "'") & " as 检查方法," & _
                        str管码项目 & " as 管码项目ID From Dual"
                Else
                    str医嘱IDs = str医嘱IDs & "," & .RowData(lngRow)
                End If
            End If
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_计价性质)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_状态)) > 0 Then
                            strTmp = Val(.TextMatrix(i, COL_诊疗项目ID)) & ":" & Val(.TextMatrix(i, COL_执行科室ID))
                            If InStr("," & str项目科室 & ",", "," & strTmp & ",") = 0 Then str项目科室 = str项目科室 & "," & strTmp
                            
                            str医嘱 = str医嘱 & " Union ALL Select " & _
                                IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, "-NULL", Val(.TextMatrix(i, COL_相关ID))) & " as 相关ID," & _
                                Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 诊疗项目ID," & Val(.TextMatrix(i, COL_执行标记)) & " as 执行标记," & _
                                IIF(.TextMatrix(i, COL_标本部位) = "", "NULL", "'" & .TextMatrix(i, COL_标本部位) & "'") & " as 标本部位," & _
                                IIF(.TextMatrix(i, COL_检查方法) = "", "NULL", "'" & .TextMatrix(i, COL_检查方法) & "'") & " as 检查方法," & _
                                Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 管码项目ID From Dual"
                        Else
                            str医嘱IDs = str医嘱IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1 '检验组合
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_计价性质)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_状态)) > 0 Then
                            strTmp = Val(.TextMatrix(i, COL_诊疗项目ID)) & ":" & Val(.TextMatrix(i, COL_执行科室ID))
                            If InStr("," & str项目科室 & ",", "," & strTmp & ",") = 0 Then str项目科室 = str项目科室 & "," & strTmp
                            
                            str医嘱 = str医嘱 & " Union ALL Select " & _
                                IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, "-NULL", Val(.TextMatrix(i, COL_相关ID))) & " as 相关ID," & _
                                Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 诊疗项目ID," & Val(.TextMatrix(i, COL_执行标记)) & " as 执行标记," & _
                                IIF(.TextMatrix(i, COL_标本部位) = "", "NULL", "'" & .TextMatrix(i, COL_标本部位) & "'") & " as 标本部位," & _
                                IIF(.TextMatrix(i, COL_检查方法) = "", "NULL", "'" & .TextMatrix(i, COL_检查方法) & "'") & " as 检查方法," & _
                                Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 管码项目ID From Dual"
                        Else
                            str医嘱IDs = str医嘱IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    str医嘱IDs = Mid(str医嘱IDs, 2)
    str单量s = Mid(str单量s, 2)
    str项目IDs = Mid(str项目IDs, 2)
    str项目科室 = Mid(str项目科室, 2)
    str医嘱 = Mid(str医嘱, 12)
    
    On Error GoTo errH
    
    If bln药品 Then
        If str项目IDs = "" Then Exit Function
    
        strSQL = "Select Rownum As 序号,Column_Value As ID From Table(f_Num2list([1]))"
        strSQL = "Select /*+ RULE */ A.ID,A.类别,A.是否变价,D.跟踪在用,Nvl(B.门诊包装,1) as 门诊包装,Nvl(B.剂量系数,1) as 剂量系数,B.门诊可否分零 As 可否分零" & _
            " From 收费项目目录 A,药品规格 B,材料特性 D,(" & strSQL & ") C" & _
            " Where A.ID=B.药品ID(+) And A.ID=D.材料ID(+) And A.ID=C.ID Order By C.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str项目IDs)
        For i = 1 To rsTmp.RecordCount
            '售价数量
            If str单量s <> "" Then '中药配方才管每味剂量
                dbl数量 = Val(Split(str单量s, ";")(i - 1))
                
                '售价数量：中药药房单位按不可分零处理:每付
                If Nvl(rsTmp!可否分零, 0) = 0 Then
                    dbl数量 = Format(dbl数量 / Nvl(rsTmp!剂量系数, 1), "0.00000")
                Else
                    dbl数量 = Format(IntEx(dbl数量 / Nvl(rsTmp!剂量系数, 1) / Nvl(rsTmp!门诊包装, 1)) * Nvl(rsTmp!门诊包装, 1), "0.00000")
                End If
            Else
                dbl数量 = Nvl(rsTmp!门诊包装, 1) '1个药房单位的售价数量
            End If
            If Nvl(rsTmp!是否变价, 0) = 1 And (rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0) Then
                dblPrice = dblPrice + Format(Format(CalcDrugPrice(rsTmp!ID, lng执行科室ID, dbl数量), "0.00000") * dbl数量, "0.00000")
            Else
                dblPrice = dblPrice + Format(Format(CalcPrice(rsTmp!ID), "0.00000") * dbl数量, "0.00000")
            End If
            
            rsTmp.MoveNext
        Next
    Else
        If str医嘱 = "" And str医嘱IDs = "" Then Exit Function
    
        If str医嘱IDs <> "" Then
            strSQL = _
                " Select /*+ RULE */B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价" & _
                " From 病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                " And ((Sysdate Between D.执行日期 And D.终止日期) Or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And B.医嘱ID IN(Select Column_Value From Table(f_Num2list([1])))" & _
                " Group by B.数量,C.是否变价,B.单价"
        End If
        If str医嘱 <> "" Then
            '由于没有加部位等条件，所以要用Distinct
            str诊疗收费 = "Select * From (" & _
                "Select Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,C.适用科室id" & _
                " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                " From 诊疗收费关系 C,Table(f_Num2list2([2])) D Where C.诊疗项目ID=D.c1" & _
                "      And (C.适用科室ID is Null or C.适用科室ID = D.c2 And C.病人来源 = 1)" & _
                " ) Where Nvl(适用科室id, 0) = Top"
                
            strSQL = IIF(strSQL = "", "", strSQL & " Union ALL") & _
                " Select " & IIF(strSQL = "", "/*+ RULE */", "") & "B.收费数量 as 数量,Decode(C.是否变价,1,Sum(D.缺省价格),Sum(D.现价)) as 单价" & _
                " From (" & str医嘱 & ") A,(" & str诊疗收费 & ") B,收费项目目录 C,收费价目 D,诊疗项目目录 E,采血管类型 F" & _
                " Where A.诊疗项目ID=B.诊疗项目ID And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And A.管码项目ID=E.ID And E.试管编码=F.编码(+)" & _
                " And (Nvl(B.收费方式,0)=1 And C.类别='4' And B.收费项目ID=F.材料ID Or Not(Nvl(B.收费方式,0)=1 And C.类别='4' And F.材料ID Is Not NULL))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                " And (A.相关ID is Null And A.执行标记 IN(1,2) And B.费用性质=1" & _
                "       Or A.标本部位=B.检查部位 And A.检查方法=B.检查方法 And Nvl(B.费用性质,0)=0" & _
                "       Or (A.检查方法 is Null Or e.类别 = 'E' And e.操作类型='4') And Nvl(B.费用性质,0)=0 And B.检查部位 is Null And B.检查方法 is Null)" & _
                " Group by B.收费数量,C.是否变价"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str医嘱IDs, str项目科室)
        For i = 1 To rsTmp.RecordCount
            dblPrice = dblPrice + Format(Nvl(rsTmp!数量, 0) * Nvl(rsTmp!单价, 0), "0.00000")
            rsTmp.MoveNext
        Next
    End If
    
    GetItemPrice = Format(dblPrice, gstrDecPrice)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function MakePriceRecord(ByVal lngRow As Long) As Boolean
'功能：根据当前新开医嘱行内容，生成对应的用于医保的费用明细记录集
'参数：lngRow=当前医嘱行
'返回：有计价数据记录集内容才返回True
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strAdvice As String, str项目科室 As String, str诊疗收费 As String
    Dim lngBegin As Long, lngEnd As Long
    Dim lng执行科室ID As Long, blnLoad As Boolean
    Dim dbl单价 As Double, dbl金额 As Double, dbl实收 As Double
    Dim str管码项目 As String, i As Long
    Dim str项目 As String, blnDo As Boolean
    Dim lng组ID As Long, lng诊断ID As Long, lng疾病ID As Long
    
    On Error GoTo errH
        
    With vsAdvice
        If .RowData(lngRow) = 0 Then Exit Function
        
        If RowIn检验行(lngRow) Then
            i = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If i <> -1 Then str管码项目 = Val(.TextMatrix(i, COL_诊疗项目ID))
        End If
        
        '生成病人医嘱记录临时表
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                str项目科室 = str项目科室 & "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & ":" & Val(.TextMatrix(i, COL_执行科室ID))
                strAdvice = strAdvice & " Union ALL " & _
                    "Select " & .RowData(i) & " as ID," & Val(.TextMatrix(i, COL_序号)) & " as 序号," & _
                    ZVal(.TextMatrix(i, COL_相关ID), True) & " as 相关ID,'" & .TextMatrix(i, COL_类别) & "' as 诊疗类别," & _
                    IIF(str管码项目 = "", Val(.TextMatrix(i, COL_诊疗项目ID)), str管码项目) & " as 管码项目ID," & _
                    Val(.TextMatrix(i, COL_诊疗项目ID)) & " as 诊疗项目ID," & ZVal(.TextMatrix(i, COL_收费细目ID), True) & " as 收费细目ID," & _
                    Val(.TextMatrix(i, COL_总量)) & " as 总量," & Val(.TextMatrix(i, COL_单量)) & " as 单量," & _
                    "'" & .TextMatrix(i, COL_标本部位) & "' as 标本部位,'" & .TextMatrix(i, COL_检查方法) & "' as 检查方法," & _
                    Val(.TextMatrix(i, COL_执行标记)) & " as 执行标记," & Val(.TextMatrix(i, COL_计价性质)) & " as 计价特性," & _
                    IIF(.TextMatrix(i, COL_类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0, 1, 0) & " as 附加手术," & _
                    Val(.TextMatrix(i, COL_执行性质)) & " as 执行性质," & ZVal(.TextMatrix(i, COL_执行科室ID), True) & " as 执行科室ID From Dual"
            End If
        Next
        strAdvice = Mid(strAdvice, 12)
        str项目科室 = Mid(str项目科室, 2)
    End With
    
    blnLoad = True
    
    '药品、卫材的计价：按售价数量、单价计算
    If vsAdvice.TextMatrix(lngRow, COL_类别) = "4" Then
        '卫材：固定按规格下达
        strSQL = "Select A.序号,A.诊疗类别,C.类别 as 收费类别,A.收费细目ID,D.收入项目ID," & _
            " Decode(A.总量,0,1,A.总量) as 数量,Decode(Nvl(C.是否变价,0),1,D.缺省价格,D.现价) as 单价," & _
            " C.是否变价,C.屏蔽费别,B.跟踪在用,A.执行科室ID,A.附加手术,D.附术收费率" & _
            " From (" & strAdvice & ") A,材料特性 B,收费项目目录 C,收费价目 D" & _
            " Where A.ID=[1] And Nvl(A.执行性质,0)<>5 And A.收费细目ID=B.材料ID" & _
            " And A.收费细目ID=C.ID And C.服务对象 IN(1,3) And D.收费细目ID=C.ID" & _
            " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        blnLoad = False
    ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
        '中,西成药:可能按规格下医嘱,计算1个药房包装的单价
        strSQL = "Select A.序号,A.诊疗类别,C.类别 as 收费类别,C.ID as 收费细目ID,D.收入项目ID," & _
            " Decode(A.总量,0,1,A.总量)/Decode(1,1,B.门诊包装,B.住院包装) as 数量," & _
            " Decode(Nvl(C.是否变价,0),1,-NULL,D.现价) as 单价,C.是否变价,C.屏蔽费别," & _
            " 0 as 跟踪在用,A.执行科室ID,A.附加手术,D.附术收费率" & _
            " From (" & strAdvice & ") A,药品规格 B,收费项目目录 C,收费价目 D" & _
            " Where A.ID=[1] And Nvl(A.执行性质,0)<>5 And A.收费细目ID=B.药品ID" & _
            " And B.药品ID=C.ID And C.服务对象 IN(1,3) And D.收费细目ID=C.ID" & _
            " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
            
        '仅一并给药(如果是)的第一成药行才显示给药途径的计价
        blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
    ElseIf RowIn配方行(lngRow) Then
        '中草药:一定对应有规格记录且填写了收费细目ID
        strSQL = "Select A.序号,A.诊疗类别,C.类别 as 收费类别,C.ID as 收费细目ID,D.收入项目ID," & _
            " Decode(A.总量,0,1,A.总量)*A.单量/Nvl(B.剂量系数,1) as 数量," & _
            " Decode(Nvl(C.是否变价,0),1,-NULL,D.现价) as 单价,C.是否变价,C.屏蔽费别," & _
            " 0 as 跟踪在用,A.执行科室ID,A.附加手术,D.附术收费率" & _
            " From (" & strAdvice & ") A,药品规格 B,收费项目目录 C,收费价目 D" & _
            " Where A.诊疗类别='7' And A.相关ID=[1] And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID" & _
            " And C.服务对象 IN(1,3) And D.收费细目ID=C.ID And Nvl(A.执行性质,0)<>5" & _
            " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
    End If
    
    '读取计价关系：除药品、卫材医嘱外的计价,包含相关医嘱计价；不计价,手工计价的医嘱不读取
    If blnLoad Then
        '由于没有加部位等条件，所以要用Distinct
        str诊疗收费 = "Select * From (" & _
                "Select Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,C.适用科室id" & _
                " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                " From 诊疗收费关系 C,Table(f_Num2list2([3])) D Where C.诊疗项目ID=D.c1" & _
                "      And (C.适用科室ID is Null or C.适用科室ID = D.c2 And C.病人来源 = 1)" & _
                " ) Where Nvl(适用科室id, 0) = Top"
                
        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
            " Select A.序号,A.诊疗类别,C.类别 as 收费类别,B.收费项目ID as 收费细目ID,D.收入项目ID," & _
            " Decode(A.总量,0,1,A.总量)*B.收费数量 as 数量,Decode(C.是否变价,1,D.缺省价格,D.现价) as 单价," & _
            " C.是否变价,C.屏蔽费别,0 as 跟踪在用,A.执行科室ID,A.附加手术,D.附术收费率" & _
            " From (" & strAdvice & ") A,(" & str诊疗收费 & ") B,收费项目目录 C,收费价目 D,诊疗项目目录 E,采血管类型 F" & _
            " Where A.诊疗类别 Not IN('4','5','6','7') And A.诊疗项目ID=B.诊疗项目ID" & _
            " And (A.相关ID is Null And A.执行标记 IN(1,2) And B.费用性质=1" & _
            "       Or A.标本部位=B.检查部位 And A.检查方法=B.检查方法 And Nvl(B.费用性质,0)=0" & _
            "       Or (A.检查方法 is Null Or e.类别 = 'E' And e.操作类型='4') And Nvl(B.费用性质,0)=0 And B.检查部位 is Null And B.检查方法 is Null)" & _
            " And A.管码项目ID=E.ID And E.试管编码=F.编码(+)" & _
            "   And (Nvl(B.收费方式,0)=1 And C.类别='4' And B.收费项目ID=F.材料ID" & _
            "       Or Not(Nvl(B.收费方式,0)=1 And C.类别='4' And F.材料ID Is Not NULL))" & _
            " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
            " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
            " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And C.服务对象 IN(1,3)" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And (A.ID=[1] Or A.ID=[2] Or A.相关ID=[1])"
    End If
    
    strSQL = "Select /*+ RULE */ A.* From (" & strSQL & ") A Order by 序号,收入项目ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)), str项目科室)
    If Not rsTmp.EOF Then
        '获取疾病ID、诊断ID
        With vsDiag
            lng组ID = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) = 0, vsAdvice.RowData(lngRow), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
            For i = 1 To .Rows - 1
                If InStr("," & .TextMatrix(i, col医嘱ID) & ",", "," & lng组ID & ",") > 0 Then
                    lng疾病ID = Val(.TextMatrix(i, col疾病ID))
                    lng诊断ID = Val(.TextMatrix(i, col诊断ID))
                    Exit For
                End If
            Next
        End With
    
        '初始化记录集
        Set mrsPrice = New ADODB.Recordset
        mrsPrice.Fields.Append "病人ID", adBigInt
        mrsPrice.Fields.Append "主页ID", adBigInt, , adFldIsNullable
        mrsPrice.Fields.Append "收费类别", adVarChar, 1
        mrsPrice.Fields.Append "收费细目ID", adBigInt
        mrsPrice.Fields.Append "数量", adDouble
        mrsPrice.Fields.Append "单价", adDouble
        mrsPrice.Fields.Append "实收金额", adDouble
        mrsPrice.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
        mrsPrice.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
        mrsPrice.Fields.Append "疾病ID", adBigInt, , adFldIsNullable
        mrsPrice.Fields.Append "诊断ID", adBigInt, , adFldIsNullable
        mrsPrice.CursorLocation = adUseClient
        mrsPrice.LockType = adLockOptimistic
        mrsPrice.CursorType = adOpenStatic
        mrsPrice.Open
        
        '加入费用明细
        dbl实收 = 0: blnDo = True
        Do While Not rsTmp.EOF
            '执行科室
            If blnDo Then
                lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                If rsTmp!收费类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!收费类别) > 0 And InStr(",5,6,7,", rsTmp!诊疗类别) = 0 Then
                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!收费类别, rsTmp!收费细目ID, 4, mlng病人科室id, 0, 2, lng执行科室ID)
                End If
            End If
            
            '单价
            If InStr(",5,6,7,", rsTmp!收费类别) > 0 And Nvl(rsTmp!是否变价, 0) = 1 Then
                '药品时价
                dbl单价 = Format(CalcDrugPrice(rsTmp!收费细目ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), gstrDecPrice)
            ElseIf rsTmp!收费类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 And Nvl(rsTmp!是否变价, 0) = 1 Then
                '跟踪在用卫材时价
                dbl单价 = Format(CalcDrugPrice(rsTmp!收费细目ID, lng执行科室ID, Nvl(rsTmp!数量, 0)), gstrDecPrice)
            Else
                dbl单价 = Format(Nvl(rsTmp!单价, 0), gstrDecPrice) '其他变价取的缺省价格
            End If
            
            '金额
            dbl金额 = CCur(Nvl(rsTmp!数量, 0) * dbl单价)
            If Nvl(rsTmp!附加手术, 0) = 1 Then
                dbl金额 = dbl金额 * Nvl(rsTmp!附术收费率, 100) / 100
            End If
            dbl金额 = Format(dbl金额, gstrDec)
            
            If Nvl(rsTmp!屏蔽费别, 0) = 0 And mstr费别 <> "" Then
                dbl金额 = ActualMoney(mstr费别, rsTmp!收入项目ID, dbl金额, rsTmp!收费细目ID, lng执行科室ID, Nvl(rsTmp!数量, 0))
            End If
            
            dbl实收 = dbl实收 + dbl金额
            
            '项目变化时加入
            str项目 = rsTmp!序号 & "," & rsTmp!收费细目ID
            blnDo = False: rsTmp.MoveNext
            If Not rsTmp.EOF Then
                If rsTmp!序号 & "," & rsTmp!收费细目ID <> str项目 Then blnDo = True
            Else
                blnDo = True
            End If
            rsTmp.MovePrevious
            
            If blnDo Then
                With vsAdvice
                    mrsPrice.AddNew
                    mrsPrice!病人ID = mlng病人ID
                    mrsPrice!主页ID = Null
                    mrsPrice!收费类别 = rsTmp!收费类别
                    mrsPrice!收费细目ID = rsTmp!收费细目ID
                    mrsPrice!数量 = Nvl(rsTmp!数量, 0)
                    mrsPrice!单价 = dbl单价
                    mrsPrice!实收金额 = dbl金额
                    If .TextMatrix(lngRow, COL_开嘱医生) <> "" Then
                        mrsPrice!开单人 = .TextMatrix(lngRow, COL_开嘱医生)
                    End If
                    If Val(.TextMatrix(lngRow, COL_开嘱科室ID)) <> 0 Then
                        mrsPrice!开单科室 = CStr(GetItemField("部门表", Val(.TextMatrix(lngRow, COL_开嘱科室ID)), "名称"))
                    End If
                    If lng疾病ID <> 0 Then mrsPrice!疾病id = lng疾病ID
                    If lng诊断ID <> 0 Then mrsPrice!诊断id = lng诊断ID
                    mrsPrice.Update
                End With
                dbl实收 = 0
            End If
            
            rsTmp.MoveNext
        Loop
        
        mrsPrice.MoveFirst
        MakePriceRecord = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDrugStock(ByVal lngRow As Long)
'功能：重新获取指定药品行的药品库存
'参数：lngRow=卫材、成药行或中药用法行
'说明：如果是中药配方行,一次性获取整个配方中的所有中药的库存
    Dim i As Long
    
    With vsAdvice
        If InStr(",4,5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Or Val(.TextMatrix(lngRow, COL_收费细目ID)) = 0 _
                Or .TextMatrix(lngRow, COL_类别) = "4" And Val(.TextMatrix(lngRow, COL_跟踪在用)) = 0 Then
                .TextMatrix(lngRow, COL_库存) = ""
            Else
                .TextMatrix(lngRow, COL_库存) = GetStock(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), 1)
            End If
        ElseIf RowIn配方行(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Or Val(.TextMatrix(i, COL_收费细目ID)) = 0 Then
                            .TextMatrix(i, COL_库存) = ""
                        Else
                            .TextMatrix(i, COL_库存) = GetStock(Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_执行科室ID)), 1)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = col_医嘱内容 Then Call vsAdvice.AutoSize(col_医嘱内容)
    End If
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpDate.Visible Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = True
    End If
    If fraAdvice.Tag <> "" Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        ElseIf Col = COL_诊断 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
    'PASS
    With vsAdvice
        If mblnPass Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
            End If
        End If
    
    End With
End Sub

Private Sub vsAdvice_DblClick()
    Dim lngRow As Long, lngCol As Long
    
    With vsAdvice
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
            If lngCol = COL_诊断 Then
                Call vsAdvice_KeyPress(32)    '切换诊断关联显示
            ElseIf lngCol >= .FixedCols And lngCol <= .Cols - 1 Then
                Call vsAdvice_KeyPress(13)    '定位到对应的编辑控件
                'PASS合理用药检测
                If mblnPass Then
                    If gobjPass.zlPassCheck(mobjPassMap) Then
                        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
                    End If
                End If
            ElseIf .MouseCol = COL_F标志 Then
                '填写申请
                '##
            End If
        End If
    End With
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否最后一可见行
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            lngLeft = COL_诊断: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_天数: lngRight = COL_用法
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowIn一并给药(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Get一并给药范围(Val(.TextMatrix(Row - 1, COL_相关ID)), lngBegin, lngEnd)
            Else
                Call Get一并给药范围(Val(.TextMatrix(Row, COL_相关ID)), lngBegin, lngEnd)
            End If
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cbsMain.FindControl(, conMenu_Delete, True, True).Execute
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    Dim lngDiag As Long
    
    If KeyAscii = 13 Then
        '定位到对应的编辑控件
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_开始时间
                If txt开始时间.TabStop Then
                    Set objEdit = txt开始时间 '缺省不定位到开始时间
                Else
                    Set objEdit = txt医嘱内容
                End If
            Case col_医嘱内容
                Set objEdit = txt医嘱内容
            Case COL_标本部位
                Set objEdit = txt安排时间
            Case COL_单量
                Set objEdit = txt单量
            Case COL_天数
                Set objEdit = txt天数
            Case COL_总量
                Set objEdit = txt总量
            Case COL_用法
                Set objEdit = txt用法
            Case COL_频率
                Set objEdit = txt频率
            Case COL_执行时间
                Set objEdit = cbo执行时间
            Case COL_执行科室ID
                Set objEdit = cbo执行科室
            Case COL_医生嘱托
                Set objEdit = cbo医生嘱托
            Case COL_标志
                Set objEdit = chk紧急
            Case COL_超量说明
                Set objEdit = txt超量说明
            Case COL_用药目的
                Set objEdit = cboDruPur
            Case COL_用药理由
                Set objEdit = txt用药理由
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    ElseIf KeyAscii = 32 Then
        '切换诊断关联
        With vsAdvice
            If .Col = COL_诊断 And .RowData(.Row) <> 0 And vsDiag.TextMatrix(vsDiag.Row, col诊断) <> "" Then
                KeyAscii = 0
                Call SetDiagFlag(.Row, IIF(AdviceHaveDiag(.Row) = vsDiag.Row, 0, 1))
            End If
        End With
    End If
End Sub

Private Sub ClearItemTag()
'功能：清除控件编辑标志
    txt开始时间.Tag = ""
    txt安排时间.Tag = ""
    txt单量.Tag = ""
    txt天数.Tag = ""
    txt总量.Tag = ""
    txt用法.Tag = ""
    txt频率.Tag = ""
    cbo执行时间.Tag = ""
    cbo医生嘱托.Tag = ""
    cbo执行科室.Tag = ""
    cbo执行性质.Tag = ""
    cbo附加执行.Tag = ""
    chk紧急.Tag = ""
    chk免试.Tag = ""
    cbo滴速.Tag = ""
    chkZeroBilling.Tag = ""
    lbl用药目的.Tag = ""
    txt用药理由.Tag = ""
    txt超量说明.Tag = ""
    lbl超量说明.Tag = ""
End Sub

Private Sub SetStartTime(ByVal Editable As Boolean)
'功能：设置开始时间是否允许编辑
    'txt开始时间.TabStop = Editable '缺省不定位到开始时间
    txt开始时间.Locked = Not Editable
    cmd开始时间.Enabled = Editable
    If Editable Then
        txt开始时间.BackColor = vsAdvice.BackColor
    Else
        txt开始时间.BackColor = &HE0E0E0
    End If
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'功能：设置执行天数可用和或见状态
'参数：0-保持不变,-1-禁止,1-允许
    If intEnabled = -1 Then
        txt天数.Enabled = False
        txt天数.BackColor = Me.BackColor
        txt天数.Text = ""
    ElseIf intEnabled = 1 Then
        txt天数.TabStop = True
        txt天数.Enabled = True
        txt天数.BackColor = vsAdvice.BackColor
    End If
    
    If intVisible = -1 Then
        lbl天数.Visible = False
        txt天数.Visible = False
        txt天数.Text = ""
        
        lbl总量.Left = lbl用法.Left + lbl用法.Width - lbl总量.Width
        txt总量.Left = txt用法.Left
        txt总量.Width = txt用法.Width - cmd用法.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        lbl单量.Left = lbl频率.Left + lbl频率.Width - lbl单量.Width
        txt单量.Left = txt频率.Left
        txt单量.Width = txt频率.Width - cmd频率.Width - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        txt总量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt总量.TabIndex + 1
        txt单量.TabIndex = txt天数.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl天数.Visible = True
        txt天数.Visible = True
        
        lbl单量.Left = lbl用法.Left + lbl用法.Width - lbl单量.Width
        txt单量.Left = txt用法.Left
        txt单量.Width = txt用法.Width - txt天数.Width - Me.TextWidth("三个字!") - 15
        lbl单量单位.Left = txt单量.Left + txt单量.Width + 30
        
        lbl总量.Left = lbl频率.Left + lbl频率.Width - lbl总量.Width
        txt总量.Left = txt频率.Left
        txt总量.Width = txt频率.Width - cmd频率.Width - 15
        lbl总量单位.Left = txt总量.Left + txt总量.Width + 30
        
        txt单量.TabIndex = cmd频率.TabIndex + 1
        txt天数.TabIndex = txt单量.TabIndex + 1
        txt总量.TabIndex = txt天数.TabIndex + 1
    End If
End Sub

Private Sub SetItemEditable(Optional int单量 As Integer, Optional int总量 As Integer, _
    Optional int用法 As Integer, Optional int频率 As Integer, _
    Optional int执行时间 As Integer, Optional int执行科室 As Integer, _
    Optional int执行性质 As Integer, Optional int附加执行 As Integer, _
    Optional int安排时间 As Integer, Optional int滴速 As Integer, _
    Optional int用药目的 As Integer, Optional int用药理由 As Integer, _
    Optional int超量说明 As Integer, Optional int输血原因 As Integer)
'功能：设置指定编辑项的可用状态
'参数：0-保持不变,-1-禁止,1-允许,2-锁定
'说明：禁止时,同时清除该项目数据(不是全部)

    '依次设置为禁止时,会引发焦点改变,从而可能引发Validate事件,所以先禁止焦点顺序
    If int单量 = -1 Then txt单量.TabStop = False
    If int总量 = -1 Then txt总量.TabStop = False
    If int用法 = -1 Then txt用法.TabStop = False
    If int频率 = -1 Then txt频率.TabStop = False
    If int执行时间 = -1 Then cbo执行时间.TabStop = False
    If int执行科室 = -1 Then cbo执行科室.TabStop = False
    If int执行性质 = -1 Then cbo执行性质.TabStop = False
    If int附加执行 = -1 Then cbo附加执行.TabStop = False
     
    If int用药目的 = -1 Then cboDruPur.TabStop = False
    If int用药理由 = -1 Then txt用药理由.TabStop = False
    
    If int单量 = -1 Then
        txt单量.Enabled = False
        txt单量.BackColor = Me.BackColor
        txt单量.Text = ""
        lbl单量单位.Caption = "" '"单位"
    ElseIf int单量 = 1 Then
        txt单量.TabStop = True
        txt单量.Enabled = True
        txt单量.BackColor = vsAdvice.BackColor
    End If

    If int总量 = -1 Then
        txt总量.Enabled = False
        txt总量.BackColor = Me.BackColor
        txt总量.Text = ""
        lbl总量单位.Caption = "" '"单位"
    ElseIf int总量 = 1 Then
        txt总量.TabStop = True
        txt总量.Enabled = True
        txt总量.BackColor = vsAdvice.BackColor
    End If
    
    If int用法 = -1 Then
        txt用法.Enabled = False
        txt用法.BackColor = Me.BackColor
        txt用法.Text = ""
        cmd用法.Enabled = False
        lbl用法.Caption = "用法"
    ElseIf int用法 = 1 Then
        txt用法.TabStop = True
        txt用法.Enabled = True
        cmd用法.Enabled = True
        txt用法.BackColor = vsAdvice.BackColor
    End If

    If int频率 = -1 Then
        txt频率.Enabled = False
        cmd频率.Enabled = False
        txt频率.BackColor = Me.BackColor
        txt频率.Text = ""
    ElseIf int频率 = 1 Then
        txt频率.TabStop = True
        txt频率.Enabled = True
        cmd频率.Enabled = True
        txt频率.BackColor = vsAdvice.BackColor
    End If

    If int执行时间 = -1 Then
        cbo执行时间.Text = ""
        cbo执行时间.Enabled = False
        cbo执行时间.BackColor = Me.BackColor
        cbo执行时间.Clear
    ElseIf int执行时间 = 1 Then
        cbo执行时间.TabStop = True
        cbo执行时间.Enabled = True
        cbo执行时间.BackColor = vsAdvice.BackColor
    End If

    If int执行科室 = -1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.Enabled = False
        cbo执行科室.BackColor = Me.BackColor
        cbo执行科室.Clear
    ElseIf int执行科室 = 1 Then
        lbl执行科室.Caption = "执行科室"
        cbo执行科室.TabStop = True
        cbo执行科室.Enabled = True
        cbo执行科室.BackColor = vsAdvice.BackColor
    End If

    If int执行性质 = -1 Then
        cbo执行性质.Enabled = False
        cbo执行性质.BackColor = Me.BackColor
        Call zlControl.CboSetIndex(cbo执行性质.hWnd, -1) '不清除
    ElseIf int执行性质 = 1 Then
        cbo执行性质.TabStop = True
        cbo执行性质.Enabled = True
        cbo执行性质.BackColor = vsAdvice.BackColor
    End If
    
    If int附加执行 = -1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.Enabled = False
        cbo附加执行.BackColor = Me.BackColor
        cbo附加执行.Clear
    ElseIf int附加执行 = 1 Then
        lbl附加执行.Caption = "附加执行"
        cbo附加执行.TabStop = True
        cbo附加执行.Enabled = True
        cbo附加执行.BackColor = vsAdvice.BackColor
    End If
    
    If int安排时间 = -1 Then
        txt安排时间.Text = ""
        lbl安排时间.Visible = False
        txt安排时间.Visible = False
        cmd安排时间.Visible = False
        lbl用法.Visible = True
        txt用法.Visible = True
        cmd用法.Visible = True
    ElseIf int安排时间 = 1 Then
        lbl安排时间.Visible = True
        txt安排时间.Visible = True
        cmd安排时间.Visible = True
        lbl用法.Visible = False
        txt用法.Visible = False
        cmd用法.Visible = False
    End If
    
    If int滴速 = -1 Then
        cbo滴速.Text = ""
        lbl滴速.Visible = False
        cbo滴速.Visible = False
        lbl滴速单位.Visible = False
    ElseIf int滴速 = 1 Then
        lbl滴速.Visible = True
        cbo滴速.Visible = True
        lbl滴速单位.Visible = True
    End If
    
    '缺省不选中
    If int用药目的 = -1 Then
        Call zlControl.CboSetIndex(cboDruPur.hWnd, 0)
        cboDruPur.Enabled = True
    ElseIf int用药目的 = 1 Then
        Call zlControl.CboSetIndex(cboDruPur.hWnd, 0)
        cboDruPur.Enabled = True
        cboDruPur.TabStop = True
    End If
        
    If int用药理由 = -1 Then
        txt用药理由.Text = ""   '没有隐藏，所以需要清空
        txt用药理由.Enabled = False
        txt用药理由.BackColor = Me.BackColor
        cmd收藏用药理由.Enabled = False
        cmdReason.Enabled = False
    ElseIf int用药理由 = 1 Then
        txt用药理由.Enabled = True
        txt用药理由.TabStop = True
        txt用药理由.BackColor = vsAdvice.BackColor
        cmd收藏用药理由.Enabled = True
        cmdReason.Enabled = True
    End If
    
    If int超量说明 = -1 Then
        txt超量说明.Text = ""   '没有隐藏，所以需要清空
        txt超量说明.Enabled = False
        txt超量说明.BackColor = Me.BackColor
        cmdExcReason.Enabled = False
        cmdComExcReason.Enabled = False
    ElseIf int超量说明 = 1 Then
        txt超量说明.Enabled = True
        txt超量说明.TabStop = True
        txt超量说明.BackColor = vsAdvice.BackColor
        cmdExcReason.Enabled = True
        cmdComExcReason.Enabled = True
    End If
    
    '=1输血医嘱，显示输血原因
    If int输血原因 = -1 Then
        lbl用药目的.Visible = True
        cboDruPur.Visible = True
        cmdReason.Visible = True
        cmd收藏用药理由.Visible = True
        lbl用药理由.Caption = "用药理由"
    ElseIf int输血原因 = 1 Then
        lbl用药目的.Visible = False
        cboDruPur.Visible = False
        cmdReason.Visible = False
        cmd收藏用药理由.Visible = False
        lbl用药理由.Caption = "输血原因"
    End If
End Sub

Private Function Get合同单位ID() As Long
'功能：获取当前病人的合同单位ID
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(合同单位ID,0) as 合同单位ID From 病人信息 Where 病人ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    
    If rsTmp.RecordCount > 0 Then Get合同单位ID = rsTmp!合同单位ID
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get诊断描述(ByVal lng诊断ID As Long, ByVal lng疾病ID As Long) As String
'功能：根据诊断ID或疾病ID获取字典表中的名称（病人诊断记录中的名称可以是修改后的,允许加前缀或后缀），以便再次修改时判断
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If lng诊断ID <> 0 Then
        strSQL = "Select 名称 From 疾病诊断目录 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng诊断ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    ElseIf lng疾病ID <> 0 Then
        strSQL = "Select 名称 From 疾病编码目录 Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng疾病ID)
        If rsTmp.RecordCount > 0 Then Get诊断描述 = "" & rsTmp!名称
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiInfo() As Boolean
'功能：读取病人信息
    Dim rsTmp As ADODB.Recordset
    Dim rsSub As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strTmp As String
    Dim blnMsgOk As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.ID,a.急诊,a.复诊,A.姓名,A.性别,A.年龄,B.出生日期,B.门诊号,B.费别,B.医疗付款方式," & _
        " Nvl(D.预交余额,0)-Nvl(D.费用余额,0) as 预交款,B.险类,B.就诊诊室,A.登记时间," & _
        " A.执行部门ID as 病人科室ID,b.身份证号" & _
        " From 病人挂号记录 A,病人信息 B,病人余额 D" & _
        " Where A.NO=[1] And a.记录性质=1 And a.记录状态=1 And A.病人ID+0=[2]" & _
        " And A.病人ID=B.病人ID And B.病人ID=D.病人ID(+) And D.性质(+)=1 And D.类型(+) = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单, mlng病人ID)
    
    lblPati.Caption = _
        "姓名：" & rsTmp!姓名 & "　性别：" & Nvl(rsTmp!性别) & "　年龄：" & Nvl(rsTmp!年龄) & _
        "　费别：" & Nvl(rsTmp!费别) & "　 医疗付款方式：" & Nvl(rsTmp!医疗付款方式) & "　预交款：" & Format(Nvl(rsTmp!预交款, 0), "0.00")
    lblPati.Tag = rsTmp!姓名 '用于医嘱复制提示
    
    '病人的准确年龄:用于判断
    mlng挂号ID = rsTmp!ID
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            mdbl门诊号 = Val(rsTmp!门诊号 & "")
        End If
    End If
    mint年龄 = GetPatiYear(mlng病人ID)
    If IsNull(rsTmp!出生日期) Then
        mDat出生日期 = DateAdd("yyyy", -mint年龄, zlDatabase.Currentdate)
    Else
        mDat出生日期 = rsTmp!出生日期
    End If
    mstr姓名 = rsTmp!姓名
    mstr身份证号 = "" & rsTmp!身份证号
    
    mstr性别 = Nvl(rsTmp!性别)
    mstr费别 = Nvl(rsTmp!费别)
    mdat挂号时间 = rsTmp!登记时间
    mlng病人科室id = rsTmp!病人科室id
    mstr付款码 = Get医疗付款码(Nvl(rsTmp!医疗付款方式))
    mbln中医 = Have部门性质(rsTmp!病人科室id, "中医科")
    mbytPatiType = IIF(Val(Nvl(rsTmp!急诊)) = 0, 1, 2)
    mbln复诊 = Val(rsTmp!复诊 & "") <> 0
    
    'PASS 传人病人信息
    If mblnPass Then
        If gobjPass.zlPassCheck(mobjPassMap) Then
            Call zlPASSPati
        End If
    End If
    '保险病人用红色显示
    mint险类 = 0
    If Not IsNull(rsTmp!险类) Then
        mint险类 = rsTmp!险类
        lblPati.ForeColor = vbRed
    End If

    mbln提醒对码 = True
    
    '诊断对应医嘱
    strSQL = "Select B.诊断ID,B.医嘱ID From 病人诊断记录 A,病人诊断医嘱 B" & _
        " Where A.ID=B.诊断ID And A.记录来源=3 And A.诊断类型 IN(1,11) And A.取消时间 is Null And A.病人ID=[1] And A.主页ID=[2]"
    Set rsSub = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    '病人诊断信息
    Set vsDiag.Cell(flexcpPicture, 0, col标志) = img16.ListImages("诊断_当前").Picture
    vsDiag.Cell(flexcpPictureAlignment, 0, col标志) = 4
    
    strSQL = "Select A.ID,A.记录来源,A.诊断次序,A.诊断类型,A.疾病ID,A.诊断ID,A.证候ID,A.诊断描述,A.是否疑诊, a.记录日期, a.记录人,B.编码 as ICD码,c.编码 as 诊断编码,d.编码 as 证候编码,A.发病时间," & _
        " b.附码 As 疾病附码, b.类别 As 疾病类别,d.名称 As 证候名称 From 病人诊断记录 A,疾病编码目录 B, 疾病诊断目录 C,疾病编码目录 D" & _
        " Where A.疾病ID=B.ID(+)  And a.诊断id = c.Id(+) And  a.证候ID=d.ID(+)  And A.记录来源 IN(1,3) And A.诊断类型 IN(1,11)" & _
        " And A.取消时间 is Null And A.病人ID=[1] And A.主页ID=[2]" & _
        " Order by A.诊断类型,A.诊断次序,a.编码序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    If Not mclsMipModule Is Nothing Then
        blnMsgOk = mclsMipModule.IsConnect
    End If
    If blnMsgOk Then
        Set mrs诊断 = New ADODB.Recordset
        mrs诊断.Fields.Append "id", adBigInt
        mrs诊断.Fields.Append "显示编码", adVarChar, 120
        mrs诊断.Fields.Append "诊断编码", adVarChar, 120
        mrs诊断.Fields.Append "疾病编码", adVarChar, 120
        mrs诊断.Fields.Append "状态", adBigInt '0-原始记录，1-被修改，2-新增的记录
        mrs诊断.CursorLocation = adUseClient
        mrs诊断.LockType = adLockOptimistic
        mrs诊断.CursorType = adOpenStatic
        mrs诊断.Open
    End If
    
    If Not rsTmp.EOF Then
        '西医诊断
        rsTmp.Filter = "记录来源=3 " & IIF(Not mbln中医, " And 诊断类型=1", "") '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" & IIF(Not mbln中医, " And 诊断类型=1", "") '其它来源的作为缺省显示
        With vsDiag
            Set mrsDiag = zlDatabase.CopyNewRec(rsTmp)
            If Not rsTmp.EOF Then
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    Call SetDiagType(i, rsTmp!诊断类型)
                    
                    If IsNull(rsTmp!诊断描述) Then
                        .TextMatrix(i, col编码) = ""
                        .TextMatrix(i, col诊断) = ""
                    Else
                        If Mid(rsTmp!诊断描述, 1, 1) <> "(" Or (Val(rsTmp!诊断id & "") = 0 And Val(rsTmp!疾病id & "") = 0) Then '中医的诊断描述后面加了（候症），所以只判断第一个字符
                            '由于疾病编码和诊断可以对应，如果两个都不为空的时候，先判断疾病编码，先取疾病编码
                            If Val(rsTmp!疾病id & "") <> 0 Then
                                .TextMatrix(i, col编码) = Nvl(rsTmp!ICD码)
                            ElseIf Val(rsTmp!诊断id & "") <> 0 Then
                                .TextMatrix(i, col编码) = Nvl(rsTmp!诊断编码)
                            Else
                                .TextMatrix(i, col编码) = ""
                            End If
                            .TextMatrix(i, col诊断) = rsTmp!诊断描述
                        Else
                            .TextMatrix(i, col编码) = Mid(rsTmp!诊断描述, 2, InStr(rsTmp!诊断描述, ")") - 2)
                            .TextMatrix(i, col诊断) = Mid(rsTmp!诊断描述, InStr(rsTmp!诊断描述, ")") + 1)
                        End If
                    End If
                    
                    '取证候名称
                    If InStr(.TextMatrix(i, col诊断), "(") > 0 And InStr(.TextMatrix(i, col诊断), ")") > 0 And Val(rsTmp!诊断类型 & "") = 11 Then
                        strTmp = Mid(.TextMatrix(i, col诊断), InStrRev(.TextMatrix(i, col诊断), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '先取证候
                        .TextMatrix(i, col中医证候) = strTmp
                        '去掉诊断描述的证候
                        .TextMatrix(i, col诊断) = Mid(.TextMatrix(i, col诊断), 1, InStrRev(.TextMatrix(i, col诊断), "(") - 1)
                    Else
                       .TextMatrix(i, col中医证候) = ""
                    End If
                    If Not IsNull(rsTmp!疾病id) Or Not IsNull(rsTmp!诊断id) Then
                        .Cell(flexcpData, i, col诊断) = Get诊断描述(Val("" & rsTmp!诊断id), Val("" & rsTmp!疾病id))    '获取原始名称以便修改时判断
                    Else
                        .Cell(flexcpData, i, col诊断) = .TextMatrix(i, col诊断)
                    End If
                    
                    .Cell(flexcpData, i, col疑诊) = Val(Nvl(rsTmp!是否疑诊, 0))
                    .Cell(flexcpForeColor, i, col疑诊) = IIF(Nvl(rsTmp!是否疑诊, 0) = 1, vbRed, .GridColor)
                    .TextMatrix(i, col诊断ID) = Nvl(rsTmp!诊断id, 0)
                    .Cell(flexcpData, i, col诊断ID) = Nvl(rsTmp!ID, 0)
                    .TextMatrix(i, col疾病ID) = Nvl(rsTmp!疾病id, 0)
                    .TextMatrix(i, col证候ID) = Nvl(rsTmp!证候id, 0)
                    .TextMatrix(i, colICD码) = Nvl(rsTmp!ICD码)
                    .TextMatrix(i, col发病时间) = Format(rsTmp!发病时间 & "", "YYYY-MM-DD HH:mm")
                  
                    .TextMatrix(i, col诊断编码) = Nvl(rsTmp!诊断编码)
                    .TextMatrix(i, col疾病编码) = Nvl(rsTmp!ICD码)
                    .TextMatrix(i, col疾病类别) = Nvl(rsTmp!疾病类别)
                    .TextMatrix(i, col疾病附码) = Nvl(rsTmp!疾病附码)
                    .TextMatrix(i, col证候编码) = Nvl(rsTmp!证候编码)
                    
                    If blnMsgOk Then
                        mrs诊断.AddNew
                        mrs诊断!ID = rsTmp!ID
                        mrs诊断!显示编码 = .TextMatrix(i, col编码)
                        mrs诊断!诊断编码 = .TextMatrix(i, col诊断编码)
                        mrs诊断!疾病编码 = .TextMatrix(i, col疾病编码)
                        mrs诊断!状态 = 0
                        mrs诊断.Update
                    End If
                    
                    '填写对应的关联医嘱
                    strSQL = ""
                    rsSub.Filter = "诊断ID=" & rsTmp!ID
                    Do While Not rsSub.EOF
                        strSQL = strSQL & "," & rsSub!医嘱ID
                        rsSub.MoveNext
                    Loop
                    .TextMatrix(i, col医嘱ID) = Mid(strSQL, 2)
                    rsTmp.MoveNext
                Next
            End If
        End With
    Else
        Call SetDiagType(1, IIF(mbln中医, 11, 1))
    End If
    
    '西医时只输入西医诊断
    vsDiag.ColHidden(col中医) = Not mbln中医
    vsDiag.ColHidden(COL西医) = Not mbln中医
    vsDiag.ColHidden(col中医证候) = Not mbln中医
    vsDiag.ColWidth(col诊断) = IIF(mbln中医, 2760, 4360)
    vsDiag.ColHidden(COLDEL) = False
    vsDiag.ColWidth(COLDEL) = vsDiag.ColWidth(col增加)
 
    vsDiag.Col = col诊断: vsDiag.Row = vsDiag.Rows - 1
    Call vsDiag_AfterRowColChange(-1, -1, vsDiag.Row, vsDiag.Col)
    Call SetDiagHeight
    'PASS诊断传人
    If mblnPass Then
        zlPassDrags
    End If
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetDiagType(ByVal lngRow As Long, ByVal int诊断类型 As Integer)
'功能：设置某一诊断行的诊断类型
    With vsDiag
        .TextMatrix(lngRow, col中医) = "中"
        .TextMatrix(lngRow, COL西医) = "西"
        
        .Cell(flexcpData, lngRow, col中医) = IIF(int诊断类型 = 11, 1, 0)
        .Cell(flexcpForeColor, lngRow, col中医) = IIF(int诊断类型 = 11, .ForeColor, .GridColor)
        .Cell(flexcpFontBold, lngRow, col中医) = IIF(int诊断类型 = 11, True, False)
        
        .Cell(flexcpData, lngRow, COL西医) = IIF(int诊断类型 = 1, 1, 0)
        .Cell(flexcpForeColor, lngRow, COL西医) = IIF(int诊断类型 = 1, .ForeColor, .GridColor)
        .Cell(flexcpFontBold, lngRow, COL西医) = IIF(int诊断类型 = 1, True, False)
    End With
End Sub

Private Function GetPreRow(ByVal lngRow As Long) As Long
'功能：取上一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'功能：取下一最近有效可见行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Function GetDefaultTime(lngRow As Long) As String
'功能：获取新开医嘱的缺省开始时间
'说明：
'      最近一条有效时间为当天，且间隔现在在补录间隔以内
'      如果没有,则取最近新开一条的时间
'      如果没有,则取当前时间
    Dim curDate As Date, strDate As String, i As Long
    
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '先从当前行向回找
        For i = lngRow - 1 To .FixedRows Step -1
            If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) Then
                If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                    If DateAdd("n", gint门诊新开医嘱间隔, CDate(.Cell(flexcpData, i, COL_开始时间))) >= curDate Then
                        strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                End If
            End If
        Next
            
        '再从最后行向回找
        If strDate = "" Then
            For i = .Rows - 1 To lngRow + 1 Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) Then
                    If Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                        If DateAdd("n", gint门诊新开医嘱间隔, CDate(.Cell(flexcpData, i, COL_开始时间))) >= curDate Then
                            strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        
        If strDate = "" Then
            '先从当前行向回找
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) _
                    And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                    strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                    Exit For
                End If
            Next
            '再从最后行向回找
            If strDate = "" Then
                For i = .Rows - 1 To lngRow + 1 Step -1
                    If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_开始时间)) _
                        And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        strDate = Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    If strDate = "" Then strDate = Format(curDate, "yyyy-MM-dd HH:mm")
    GetDefaultTime = strDate
End Function

Private Function GetCurRow序号(lngRow As Long) As Long
'功能：获取指定行可用的的序号
'参数：lngRow=要取序号的行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng序号 As Long, i As Long
    Dim lng序号1 As Long, lng序号2 As Long
            
    '取之后最近一个有效序号,直接使用
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                lng序号 = Val(vsAdvice.TextMatrix(i, COL_序号))
                Exit For
            End If
        End If
    Next
    If lng序号 = 0 Then
        '后面没有,则取数据库之中的最大序号与之前的最大序号比较
        On Error GoTo errH
        strSQL = "Select Max(序号) as 序号 From 病人医嘱记录 Where 病人ID+0=[1] And 挂号单=[2] And Nvl(婴儿,0)=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, cbo婴儿.ListIndex)
        If Not rsTmp.EOF Then lng序号1 = Nvl(rsTmp!序号, 0)
        On Error GoTo 0
        
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                    And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                    lng序号2 = Val(vsAdvice.TextMatrix(i, COL_序号))
                    Exit For
                End If
            End If
        Next
        
        If lng序号1 > lng序号2 Then
            lng序号 = lng序号1
        Else
            lng序号 = lng序号2
        End If

        If lng序号 <> 0 Then lng序号 = lng序号 + 1 '最大序号+1
    End If
    If lng序号 = 0 Then lng序号 = 1
    GetCurRow序号 = lng序号
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSet医嘱序号(lngRow As Long, intStep As Integer)
'功能：将当前病人医嘱记录中序号前移或后移
'参数：lngRow=起始调整行,intStep=调整步长,如1或-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_序号)) Then
                vsAdvice.TextMatrix(i, COL_序号) = Val(vsAdvice.TextMatrix(i, COL_序号)) + intStep
                If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 3 '标志修改了序号
                End If
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'功能：指定的医嘱删除处理
    Dim lngBegin As Long, lngEnd As Long
    Dim lng相关ID As Long, blnGroup As Boolean
    Dim lng医嘱ID As Long, i As Integer
    Dim lngDiag As Long, lng审核状态 As Long
    
    lngDiag = -1
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        '调用删除前外挂接口
        On Error Resume Next
        If Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(vsAdvice.RowData(lngRow)), mint场合) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
        
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            lng医嘱ID = vsAdvice.RowData(lngRow)
            lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
            blnGroup = RowIn一并给药(lngRow)
            If blnGroup Then
                '先删除一并给药中的空行(一定要删)
                Call Get一并给药范围(lng相关ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '必须反向
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                '删除之后当前行号可能变了
                lngRow = vsAdvice.FindRow(lng医嘱ID, lngBegin)
                    
                '一并给药只删除当前行
                lngDiag = AdviceHaveDiag(lngRow) '记录这一行的关联诊断
                lng审核状态 = Val(vsAdvice.TextMatrix(lngRow, COL_审核状态))
                Call DeleteRow(lngRow)
                
                If lng审核状态 <> 0 Then
                    lngRow = vsAdvice.FindRow(lng相关ID, lngBegin)
                    Call ReSet审核状态图标(lngRow)
                End If
            Else
                '单独的成药：删除给药途径行及当前行
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,K,", vsAdvice.TextMatrix(lngRow, COL_类别)) > 0 Then
            Call Delete检查手术输血(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn配方行(lngRow) Then
            '删除组成味药及煎法行:删除之后重新定位的当前行
            lngRow = Delete中药配方(lngRow)
            '删除当前行(中药用法行)
            Call DeleteRow(lngRow)
        ElseIf RowIn检验行(lngRow) Then
            lngRow = Delete检验组合(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '标记为未保存
    Else
        '空行直接删除
        Call DeleteRow(lngRow)
    End If
    
    '重新定位行
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    '恢复一并给药的诊断关联
    If lngDiag <> -1 Then
        Call SetDiagFlag(vsAdvice.FindRow(lng相关ID), 1, lngDiag) '当前行也应是恢复在一并给药中的
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'功能：删除表格中的一行,但不改变当前行
'参数：blnClear=是否仅清除该行内容,不删除
'      blnDelID=是否记录要删除的医嘱ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '调整序号
            Call AdviceSet医嘱序号(lngRow + 1, -1)
            
            '记录要删除的ID(除了才新增的)
            If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 And blnDelID Then
                If .TextMatrix(lngRow, COL_处方审查状态) = "1" Or .TextMatrix(lngRow, COL_处方审查状态) = "2" Then
                    If InStr("," & mstrAduitDelIDs & ",", "," & .TextMatrix(lngRow, COL_相关ID) & ",") = 0 And .TextMatrix(lngRow, COL_相关ID) <> "" Then
                        mstrAduitDelIDs = mstrAduitDelIDs & "," & .TextMatrix(lngRow, COL_相关ID)
                        mstrDelIDs = Replace(mstrDelIDs, "," & .TextMatrix(lngRow, COL_相关ID), "")
                    End If
                Else
                    mstrDelIDs = mstrDelIDs & "," & .RowData(lngRow)
                End If
                If .TextMatrix(lngRow, COL_类别) = "K" And gbln血库系统 Then
                    mstrDel输血 = mstrDel输血 & "," & .RowData(lngRow)
                End If
            End If
            
            '删除后关联诊断的标记处理：可能是增改一组医嘱时的部分行删除，没关系，增改后立即会再标记上
            Call SetDiagFlag(lngRow, 0)
        End If
            
        '如果为行1且仅剩行1或仅清除,则保留
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '清除该行数据
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '文字
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '数据
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '粗体
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '文字色
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '固定列文字色
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '固定列背景色
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '单元图片
            Set .Cell(flexcpPicture, lngRow, COL_警示) = Nothing 'Pass警示灯
            
            '单元格边框
            .Select lngRow, .FixedCols, lngRow, COL_标志
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '因为有删除行,所以调用程序肯定有行定位,所以不必恢复行
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete检查手术输血(ByVal lngRow As Long, Optional ByVal bln申请序号 As Boolean, Optional ByRef lngTmpRow As Long)
'功能：1.删除检查组合项目的部位行
'      2.删除手术项目的附加手术行及麻醉项目行
'      3.删除输血项目的输血途径行
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lngNo As Long
    Dim strRows As String
    Dim varArr As Variant
    Dim lngTmp As Long
    On Error GoTo errH
    With vsAdvice
        If bln申请序号 Then
            lngNo = Val(.TextMatrix(lngRow, COL_申请序号))
        End If
        If lngNo = 0 Then
            i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID) '不一定有,所以用查找
            If i <> -1 Then
                lngBegin = i
                For i = lngBegin To vsAdvice.Rows - 1
                    If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                        lngEnd = i
                    Else
                        Exit For
                    End If
                Next
                For i = lngEnd To lngBegin Step -1
                    Call DeleteRow(i)
                Next
            End If
        Else
            lngTmp = -1
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_状态)) < 2 Then
                    If lngNo = Val(.TextMatrix(i, COL_申请序号)) Then
                        strRows = i & IIF(strRows = "", "", "," & strRows)
                    End If
                End If
                If i > lngRow Then
                    If lngNo <> Val(.TextMatrix(i, COL_申请序号)) Then
                        If lngTmp = -1 Then
                            lngTmp = vsAdvice.RowData(i)
                        End If
                    End If
                End If
            Next
            varArr = Split(strRows, ",")
            For i = 0 To UBound(varArr)
                Call DeleteRow(Val(varArr(i)))
            Next
            
            If lngTmp = -1 Then
                '先删除中间间隔的空行
                mblnRowChange = False
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                mblnRowChange = True
                .AddItem "", .Rows
                .Row = .Rows - 1
                .Col = .FixedCols
                lngTmpRow = .Row
            Else
                For i = .FixedRows To .Rows - 1
                    If lngTmp = Val(vsAdvice.RowData(i)) Then
                        lngTmp = i
                        Exit For
                    End If
                Next
                i = lngTmp
                If i <> -1 Then
                    .AddItem "", i
                    .Row = i
                    lngTmpRow = i
                Else
                    '先删除中间间隔的空行
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                    lngTmpRow = .Row
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Delete检验申请单(ByVal lngRow As Long, Optional ByVal bln申请序号 As Boolean, Optional ByRef lngTmpRow As Long)
'功能：1.删除检查组合项目的部位行
'      2.删除手术项目的附加手术行及麻醉项目行
'      3.删除输血项目的输血途径行
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim lngNo As Long
    Dim strRows As String
    Dim varArr As Variant
    Dim lngTmp As Long
    On Error GoTo errH
    With vsAdvice
        If bln申请序号 Then
            lngNo = Val(.TextMatrix(lngRow, COL_申请序号))
        End If
        If lngNo = 0 Then
            lngTmpRow = Delete检验组合(lngRow)
        Else
            lngTmp = -1
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COL_状态)) < 2 Then
                    If lngNo = Val(.TextMatrix(i, COL_申请序号)) Then
                        strRows = i & IIF(strRows = "", "", "," & strRows)
                    End If
                End If
                If i > lngRow Then
                    If lngNo <> Val(.TextMatrix(i, COL_申请序号)) Then
                        If lngTmp = -1 Then
                            lngTmp = vsAdvice.RowData(i)
                        End If
                    End If
                End If
            Next
            varArr = Split(strRows, ",")
            For i = 0 To UBound(varArr)
                Call DeleteRow(Val(varArr(i)))
            Next
            
            If lngTmp = -1 Then
                '先删除中间间隔的空行
                mblnRowChange = False
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                mblnRowChange = True
                .AddItem "", .Rows
                .Row = .Rows - 1
                .Col = .FixedCols
                lngTmpRow = .Row
            Else
                For i = .FixedRows To .Rows - 1
                    If lngTmp = Val(vsAdvice.RowData(i)) Then
                        lngTmp = i
                        Exit For
                    End If
                Next
                i = lngTmp
                If i <> -1 Then
                    .AddItem "", i
                    .Row = i
                    lngTmpRow = i
                Else
                    '先删除中间间隔的空行
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                    lngTmpRow = .Row
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Delete中药配方(ByVal lngRow As Long) As Long
'功能：删除中药配方的组成味药及煎法行
'参数：lngRow=中药配方用法行(可见)
'返回：删除之后重新定位的当前行(中药用法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到中药用法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete中药配方 = vsAdvice.Row
End Function

Private Function Delete检验组合(ByVal lngRow As Long) As Long
'功能：删除一并采集的多个检验项目行
'参数：lngRow=采集方法行(可见)
'返回：删除之后重新定位的当前行(采集方法行)
    Dim lngBegin As Long, lngEnd As Long
    Dim lng医嘱ID As Long, i As Long
    
    lng医嘱ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '因为是在前面删除,需要重新定位到采集方法行
    i = vsAdvice.FindRow(lng医嘱ID)
    vsAdvice.Row = i '不可能找不到
    
    mblnRowChange = True
    
    Delete检验组合 = vsAdvice.Row
End Function

Private Function Get检查部位方法(ByVal lngRow As Long) As String
'功能：获取指定行的检查部位方法串
'参数：lngRow=检查医嘱的可见行
'返回："部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'      如果是老的检查组合方式，或者是以前的单部位检查，则返回空以便程序识别
    Dim str部位 As String, str部位Last As String
    Dim str方法 As String, i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_诊疗项目ID)) <> Val(.TextMatrix(lngRow, COL_诊疗项目ID)) Then Exit Function '老的方式
                
                If .TextMatrix(i, COL_标本部位) <> "" Then
                    If .TextMatrix(i, COL_标本部位) <> str部位Last And str部位Last <> "" Then
                        str部位 = str部位 & "|" & str部位Last & IIF(str方法 <> "", ";" & Mid(str方法, 2), "")
                        str方法 = ""
                    End If
                    If .TextMatrix(i, COL_检查方法) <> "" Then
                        str方法 = str方法 & "," & .TextMatrix(i, COL_检查方法)
                    End If
                    
                    str部位Last = .TextMatrix(i, COL_标本部位)
                End If
            Else
                Exit For
            End If
        Next
        If str部位Last <> "" Then
            str部位 = str部位 & "|" & str部位Last & IIF(str方法 <> "", ";" & Mid(str方法, 2), "")
        End If
        Get检查部位方法 = Mid(str部位, 2) & vbTab & Val(.TextMatrix(lngRow, COL_执行标记))
    End With
End Function

Private Function Get手术附加IDs(ByVal lngRow As Long) As String
'功能：获取指定手术行的附加手术及麻醉项目ID串
'返回："手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
    Dim strTmp As String, lng麻醉ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_相关ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_相关ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_类别) = "G" Then
                    lng麻醉ID = Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_诊疗项目ID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get手术附加IDs = Mid(strTmp, 2) & ";" & IIF(lng麻醉ID = 0, "", lng麻醉ID)
End Function

Private Function Get中药配方IDs(ByVal lngRow As Long) As String
'功能：获取中药配方的组成味药及煎法ID串
'返回："中药ID1,单量1,脚注1;中药ID2,单量2,脚注2;...|煎法ID"
    Dim lng煎法ID As Long, str中药IDs As String, i As Long, lng形态 As Long
    Dim lng付数 As Long, lng药房ID As Long
    Dim strTmp As String
    
    With vsAdvice
        lng形态 = Val(.TextMatrix(lngRow, COL_中药形态))    '用法行
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_类别) = "E" Then
                    lng煎法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    strTmp = .TextMatrix(i, COL_标本部位) '代表中药的 煎量
                ElseIf .TextMatrix(i, COL_类别) = "7" Then
                    str中药IDs = Val(.TextMatrix(i, COL_收费细目ID)) & "," & _
                        .TextMatrix(i, COL_单量) & "," & .TextMatrix(i, COL_医生嘱托) & _
                        ";" & str中药IDs
                    If lng药房ID = 0 Then
                        lng药房ID = .TextMatrix(i, COL_执行科室ID)
                        lng付数 = .TextMatrix(i, COL_总量)
                    End If
                End If
            Else
                Exit For
            End If
        Next
        Get中药配方IDs = Mid(str中药IDs, 1, Len(str中药IDs) - 1) & "|" & lng煎法ID & "|" & lng形态 & "|" & lng付数 & "|" & lng药房ID & "|" & strTmp
    End With
End Function

Private Function Get检验组合IDs(ByVal lngRow As Long) As String
'功能：获取一并采集的检验组合项目ID及标本
'返回："'      检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本""
    Dim str项目IDs As String, str标本 As String, i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                If Val(.TextMatrix(i, COL_组合项目ID)) = 0 And mblnNewLIS Then
                    For j = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngRow) Then
                            If Val(.TextMatrix(j, COL_组合项目ID)) = Val(.TextMatrix(i, COL_诊疗项目ID)) And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                                str项目IDs = "|" & Val(.TextMatrix(j, COL_诊疗项目ID)) & str项目IDs
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    str项目IDs = "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & str项目IDs
                Else
                    If Not mblnNewLIS Then
                        str项目IDs = "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & str项目IDs
                    End If
                End If
                str标本 = .TextMatrix(i, COL_标本部位)
            Else
                Exit For
            End If
        Next
    End With
    Get检验组合IDs = Right(str项目IDs, Len(str项目IDs) - 1) & ";" & str标本
End Function

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, COL_类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "C" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_类别) = "E" Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, COL_类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_类别) = "7" And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn一并给药(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中
'参数：lngRow=可见的行,可能是空行
'说明：一并给药的范围中可能存在空行
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng相关ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngNextRow, COL_相关ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 _
            And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            
            lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_类别)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_相关ID)) = lng相关ID Then blnGroup = True
            End If
        End If
    End With
    RowIn一并给药 = blnGroup
End Function

Private Function AdviceSet过敏试验(ByVal lngRow As Long, ByVal lng皮试ID As Long) As Boolean
'功能：自动增加皮试行
'参数：lngRow=当前输入行,已经输入西药或中成药
'      lng皮试ID=要增加的皮试项目ID
'说明：自动增加之后,当前行及光标仍定位在已刚输入的药品行位置
    Dim rsInput As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
        
    '类别ID,名称,诊疗项目ID,收费细目ID,规格,产地
    strSQL = "Select 类别 as 类别ID,名称,ID as 诊疗项目ID,NULL as 收费细目ID,NULL as 规格,NULL as 产地,NULL as 项目特性 From 诊疗项目目录 Where ID=[1]"
    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng皮试ID)
        
    '寻找实际要加入皮试的行:一并给药的情况
    With vsAdvice
        For i = lngRow - 1 To .FixedRows - 1 Step -1 '可能是增加在最前面
            If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lngRow = i + 1: Exit For '新增行的行号
            End If
        Next
    End With
    
    '加入空行
    vsAdvice.AddItem "", lngRow
    
    '增加皮试
    Call AdviceInput(rsInput, lngRow, True)
    
    '重新定位到输入的药品行
    mblnRowChange = False
    vsAdvice.Row = vsAdvice.Row + 1
    mblnRowChange = True
    
    AdviceSet过敏试验 = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long, Optional ByVal blnBy皮试 As Boolean) As Boolean
'功能：根据新输的诊疗项目(新增或更换)设置缺省的医嘱数据
'参数：rsInput=输入或选择返回的记录集,lngRow=当前输入行,blnBy皮试=是否自动增加皮试输入
'返回：本次录入是否有效
    Dim str过敏 As String, blnGroup As Boolean, i As Long
    Dim lng用法ID As Long, lngGroupRow As Long, lngRowID As Long
    Dim lng皮试ID As Long, bln皮试行 As Boolean
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, strAppend As String
    Dim intType As Integer, str摘要 As String, lng审核状态 As Long, str摘要Out As String
    Dim strMsg As String, vMsg As VbMsgBoxResult
    Dim objControl As CommandBarControl
    Dim lngPreRow中药房 As Long, lng药品ID As Long
    Dim blnOK As Boolean
    Dim t_Pati As TYPE_PatiInfoEx
    Dim str手术部位 As String
    Dim strIDs1 As String, strIDs2 As String, str医嘱内容 As String
    Dim lngBegin As Long, lngEnd As Long, sng天数 As Single
    Dim lngAppType As Long '申请单应用
    Dim objAppPages()  As clsApplicationData
    Dim rsCard As ADODB.Recordset
    Dim lngTmpRow As Long
    Dim bln备血 As Boolean '是否为备血医嘱 备血=0，用血=1,存于K类别医嘱行的 检查方法  字段;备血-采集方式 / 用血-输血途径
    Dim strWhere As String
        Dim lngApplyID As Long
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) '取上一有效行,某些内容缺省与上一行相同
    lngNextRow = GetNextRow(lngRow) '取下一有效行
    
    '项目附加数据输入及输入合法性检查
    '---------------------------------------------------------------------------------------------------------------
    txt医嘱内容.Text = rsInput!名称 '暂时显示
    
    '药品处方职务检查(护士站在保存时检查)
    If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
        strMsg = CheckOneDuty(rsInput!名称, Nvl(rsInput!处方职务ID), UserInfo.姓名, InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> "")
        If strMsg <> "" Then
            vsAdvice.Refresh
            MsgBox strMsg, vbInformation, gstrSysName
            vsAdvice.Refresh: Exit Function
        End If
        If mblnPass Then
            If gobjPass.zlPassCheck(mobjPassMap) Then
                Call gobjPass.zlPassAdviceInput(mobjPassMap, Val(rsInput!诊疗项目ID & ""), Val(rsInput!收费细目ID & ""), rsInput!名称 & "")
            End If
        End If
    End If
    With vsAdvice
        '检查是否存在有效的住院医嘱或留观医嘱
        If rsInput!类别ID = "Z" And InStr(",留观,住院,", "," & rsInput!项目特性 & ",") > 0 Then
            If CheckInHosAdvice Then
                Exit Function
            End If
        End If
        
        '医保病人输入内容时的提示：非医保病人也要调(Or And mint险类 <> 0)
        If InStr(",7,8,9,", rsInput!类别ID) = 0 Then '成套，配方，单味中药不在这里提示
            str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, Nvl(rsInput!收费细目ID, 0), "", 0, "", rsInput!诊疗项目ID)
        End If
    
        
        '检验项目：采集方法判断
        If rsInput!类别ID = "C" Then
            '所有数据中取一个缺省的采集方法,同时判断是否有采集方法数据
            lng用法ID = Get缺省用法ID(6, 1)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的标本采集方法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '缺省与上一行相同
            If lngPreRow <> -1 Then
                If RowIn检验行(lngPreRow) Then
                    If Val(.TextMatrix(lngPreRow, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
                End If
            End If
        End If
        
        '中药配方：给成与中药用法判断
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            If rsInput!类别ID = "8" Then
                If GetGroupCount(rsInput!诊疗项目ID, 1, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!名称 & """是一个中药配方，但没有设置有效的组成中药。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
            
                '部份药无效的提示
                strMsg = GetGroupNone(rsInput!诊疗项目ID, 1)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "配方""" & rsInput!名称 & """中以下药品已撤档或服务对象不匹配：" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "这些药品将不会出现在配方中。", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '所有数据中取一个缺省的中药用法,同时判断是否有中药用法数据
            lng用法ID = Get缺省用法ID(4, 1)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的中药用(服)法,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '中药用法缺省与上一行相同
            If RowIn配方行(lngPreRow) Then
                If Val(.TextMatrix(lngPreRow, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(lngPreRow, COL_诊疗项目ID))
            End If
        End If
        
        '输血医嘱：输血途径判断
        If rsInput!类别ID = "K" Then
            If gbln血库系统 Then
                vMsg = frmMsgBox.ShowMsgBox("请选择输血医嘱类型。", Me, , 2)
                If vMsg = vbNo Then
                    bln备血 = True
                ElseIf vMsg = vbCancel Then
                    Exit Function
                End If
            Else
                bln备血 = True
            End If
            strWhere = ""
            If bln备血 = False And gbln血库系统 = True Then
                strWhere = " And NVL(执行分类,0)=1 "
            End If
            '所有数据中取一个缺省的输血途径
            lng用法ID = Get缺省用法ID(IIF(bln备血 And gbln血库系统, 9, 8), 2, strWhere)
            If lng用法ID = 0 Then
                .Refresh
                MsgBox "没有可用的输血" & IIF(bln备血 And gbln血库系统, "采集方法", "途径") & ",请先到诊疗项目管理中设置！", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '缺省与上一行相同
            If lngPreRow <> -1 Then
                If .TextMatrix(lngPreRow, COL_类别) = "K" And Val(.TextMatrix(lngPreRow, COL_检查方法)) = IIF(bln备血, "0", "1") Then
                    i = .FindRow(CStr(.RowData(lngPreRow)), lngPreRow + 1, COL_相关ID)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        End If
        
        '中西成药：给药途径判断
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
'            '所有数据中取一个缺省的给药途径,同时判断是否有给药途径数据
'            lng用法ID = Get缺省用法ID(2, 1)
'            If lng用法ID = 0 Then
'                .Refresh
'                MsgBox "没有可用的给药途径,请先到诊疗项目管理中设置！", vbInformation, gstrSysName
'                .Refresh: Exit Function
'            End If
            '给药途径缺省与上一个行相同剂型的相同
            If lngPreRow <> -1 And Not IsNull(rsInput!药品剂型) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 And .TextMatrix(lngPreRow, COL_药品剂型) = Nvl(rsInput!药品剂型) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_相关ID)), lngPreRow + 1)
                    If i <> -1 Then
                        If Val(.TextMatrix(i, COL_是否停用)) = 0 Then lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        End If
        
        '中西成药：过敏试验检查
        If InStr(",5,6,", rsInput!类别ID) > 0 And gint过敏登记有效天数 <> 0 Then
            str过敏 = Check过敏试验(Me, txt医嘱内容, mlng病人ID, rsInput!诊疗项目ID, rsInput!名称, mbln自动皮试, lng皮试ID)
            If str过敏 <> "" Then
                .Refresh
                If MsgBox(str过敏, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .Refresh: Exit Function
                End If
            End If
            
            '自动添加皮试
            If lng皮试ID <> 0 Then
                '先检查是否已有该皮试(在有效时间内手工或自动添加的,包括免试皮试)
                i = .FixedRows - 1
                Do
                    i = .FindRow(CStr(lng皮试ID), i + 1, COL_诊疗项目ID)
                    If i <> -1 Then
                        If Not .RowHidden(i) Then
                            If Int(CDate(.Cell(flexcpData, i, COL_开始时间))) >= Int(zlDatabase.Currentdate - gint过敏登记有效天数) Then
                                bln皮试行 = True: Exit Do '记录以作标志,当前行输入完成后再增加
                            End If
                        End If
                    End If
                Loop Until i = -1
            End If
        End If
        
        '中西成药：一并给药的判断,新行缺省是按下一并的（如果上一行是一并）
        blnGroup = RowIn一并给药(lngRow) And Not blnBy皮试
        If blnGroup Then
            If rsInput!类别ID = "9" Then
                .Refresh
                MsgBox "不能在一并给药的药品中直接输入成套方案。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        
            If .RowData(lngRow) = 0 Then
                '一并给药中的待输入空行：只有插入在一并给药的中间,才能自动成为一并给药
                lngGroupRow = lngPreRow
            Else
                '一并给药中的药品行：可能是第一行或最后一行'取当前行的下一行，避免在操作已有医嘱时重选诊疗项目操作时，当前行的内容被删除，后续过程无法取到其中的值
                If lngPreRow = -1 Then
                    lngGroupRow = vsAdvice.FindRow(.TextMatrix(lngRow, COL_相关ID), lngRow + 1, COL_相关ID)
                Else
                    If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 _
                        And Val(.TextMatrix(lngPreRow, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                        lngGroupRow = lngPreRow
                    Else
                        lngGroupRow = lngNextRow
                    End If
                End If
            End If
            
            '一并给药的,类别，必须相同
            If Decode(rsInput!类别ID, "5", "Y", "6", "Y", "N") <> Decode(.TextMatrix(lngGroupRow, COL_类别), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "该组一并给药的药品必须都为西成药或中成药。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
            lng用法ID = Val(.TextMatrix(i, COL_诊疗项目ID)) '一并给药的给药途径相同
            
            '检查一并给药的的给药途径是否适合于当前输入药品(非一并给药的缺省用法在输入函数中作了判断处理)
            If Not Check适用用法(lng用法ID, rsInput!诊疗项目ID, 1) Then
                .Refresh
                MsgBox "一并的给药途径为""" & .TextMatrix(i, col_医嘱内容) & """，不适用于当前输入药品。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
        End If
    
        '成套项目
        If rsInput!类别ID = "9" Then
            If GetGroupCount(rsInput!诊疗项目ID, 1) = 0 Then
                .Refresh
                MsgBox """" & rsInput!名称 & """是一个成套方案，但没有设置有效的组成项目。" & vbCrLf & "请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!诊疗项目ID, 1, mlng病人科室id, mstr性别)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '需要输入更多数据的一些项目
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        lngAppType = -1
        If rsInput!类别ID <> "9" Then strExtData = ""
        If rsInput!类别ID = "D" Then
            If gblnOut必用 And CanUseApply("D", Val(rsInput!诊疗项目ID & "")) Then
                lngAppType = 0
            Else
                '检查项目：都要扩展编辑了，不象以前还有单部位项目
                intType = 0
            End If
        ElseIf rsInput!类别ID = "F" Then
            '手术：需要输入麻醉项目，及可选择附加手术
            intType = 1
        ElseIf InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理)
            intType = 2
        ElseIf rsInput!类别ID = "C" Then
            If gblnOut必用 And CanUseApply("C", Val(rsInput!诊疗项目ID & ""), rsInput!编码 & "") Then
                lngAppType = 3
            Else
                '输入一并采集的多个检验项目及检验标本
                intType = 4
                strExtData = rsInput!诊疗项目ID & ";" & Nvl(rsInput!规格) '项目;标本
            End If
        ElseIf rsInput!类别ID = "K" Or rsInput!类别ID = "E" Or rsInput!类别ID = "Z" Then
            If gblnOut必用 And CanUseApply(rsInput!类别ID) Then
                lngAppType = 1
            ElseIf CheckApplication(Val(rsInput!诊疗项目ID & ""), 1) Then
                '治疗和输血类的如果有申请附项则填写
                intType = 5
            End If
        End If
        '判断当前项目是否绑定有自定义申请单，如果有则弹出
        lngApplyID = GetApplyCustom(Val(rsInput!诊疗项目ID & ""))
        If intType <> -1 And lngApplyID = 0 Then
            With t_Pati
                .bln医保 = InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> ""
                .int险类 = mint险类
                .int婴儿 = mint婴儿
                .lng病人ID = mlng病人ID
                .lng病人科室ID = mlng病人科室id
                .str挂号单 = mstr挂号单
                .str性别 = mstr性别
            End With
            If intType = 2 Then
                lngPreRow中药房 = GetPreRow中药房(lngRow)
                lng药品ID = Val("" & rsInput!收费细目ID)   '一组配方时为空
            End If
            On Error Resume Next
            '改造接口：以前int场合传未传，现在传0，bytUseType以前未传，现在传0
            If intType = 2 Then
                blnOK = frmAdviceFormula.ShowMe(Me, gclsInsure, txt医嘱内容.hWnd, t_Pati, 0, 0, 1, 1, 1, rsInput!诊疗项目ID, strExtData, _
                             str摘要Out, lng药品ID, lngPreRow中药房)
            Else
                blnOK = frmAdviceEditEx.ShowMe(Me, txt医嘱内容.hWnd, t_Pati, 0, intType, 0, 1, 1, 1, mblnNewLIS, True, rsInput!诊疗项目ID, strExtData, _
                            strAppend, GetAdviceAppendItem, GetAdviceDiagnosis, str手术部位)
            End If
            
            On Error GoTo errH
            If intType = 2 Then str摘要 = str摘要Out
            If Not blnOK Then Exit Function
        End If
        
        If lngAppType <> -1 Or lngApplyID <> 0 Then
            On Error Resume Next
            If lngAppType = 0 Then
                blnOK = ApplyNew检查申请(0, rsInput!诊疗项目ID & "", objAppPages())
            ElseIf lngAppType = 1 Then
                blnOK = ApplyNew输血申请(0, rsInput!诊疗项目ID & "", rsCard, bln备血)
            ElseIf lngAppType = 3 Then
                blnOK = ApplyNew检验申请(0, rsInput!编码 & "", rsCard)
                        ElseIf lngApplyID <> 0 Then
                FuncApplyCustom 0, lngApplyID, , Val(rsInput!诊疗项目ID & "")
            End If
            On Error GoTo errH
            If Not blnOK Then Exit Function
        End If
        
        '修改已有项目时,先删除当前医嘱的内容
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '西成药、中成药
                If Not blnGroup Then
                    '单个成药删除给药途径行,并清除当前行
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    '一组成药时,只清除当前行
                    lng审核状态 = Val(vsAdvice.TextMatrix(lngRow, COL_审核状态))
                    Call DeleteRow(lngRow, True)
                    If lng审核状态 <> 0 Then Call ReSet审核状态图标(lngGroupRow)
                End If
            ElseIf InStr(",D,F,K,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If .TextMatrix(lngRow, COL_类别) = "D" And 0 <> Val(.TextMatrix(lngRow, COL_申请序号)) Then
                    Call Delete检查手术输血(lngRow, True, lngTmpRow)
                    lngRow = lngTmpRow
                Else
                    '检查组合项目，手术项目，输血医嘱
                    '删除部位行，或手术附加行(附加手术,麻醉项目)，或输血途径
                    Call Delete检查手术输血(lngRow)
                    '清除当前行
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf RowIn配方行(lngRow) Then
                '中药配方：顺序(序号)要求必须严格控制
                '删除组成味药及煎法行:删除之后重新定位的当前行
                lngRow = Delete中药配方(lngRow)
                '清除当前行(中药用法行)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn检验行(lngRow) Then
                If lngAppType = 3 Then
                    Call Delete检验申请单(lngRow, True, lngTmpRow)
                    lngRow = lngTmpRow
                Else
                    '删除检验项目行:删除之后重新定位的当前行
                    lngRow = Delete检验组合(lngRow)
                    '清除当前行(采集方法行)
                    Call DeleteRow(lngRow, True)
                End If
            Else
                '其它项目直接清除当前行内容
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '当前行新增医嘱
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!类别ID) > 0 Then
            '中药配方(单味草药当配方处理):处理之后重新定位的当前行
            lngRow = AdviceSet中药配方(rsInput!诊疗项目ID, lngRow, lng用法ID, strExtData, , str摘要)
    
            '新增后关联诊断的标记处理
            Call SetDiagFlag(vsAdvice.Row, 1)
        ElseIf rsInput!类别ID = "9" Then
            '成套医嘱需要分解为多个项目加入
            Call AdviceSet成套项目(rsInput!诊疗项目ID, lngRow, strExtData)
            
            '新增后关联诊断的标记处理
            '成套和复制时在子过程中批量处理了
        ElseIf rsInput!类别ID = "C" Then
            If lngAppType = 3 Then
                Call AdviceSet检验申请(lngRow, rsCard)
            Else
                '检验组合
                lngRow = AdviceSet检验组合(lngRow, lng用法ID, strExtData, , str摘要)
            
                '新增后关联诊断的标记处理
                Call SetDiagFlag(vsAdvice.Row, 1)
            End If
        ElseIf lngAppType = 0 Then
            '加入检查申请单
            lngRow = AdviceSet检查申请(lngRow, objAppPages())
            Call SetDiagFlag(vsAdvice.Row, 1)
        ElseIf lngAppType = 1 Then
            lngRow = AdviceSet输血申请(lngRow, rsCard)
            Call SetDiagFlag(vsAdvice.Row, 1)
        Else
            '输血医嘱检查，必须中级及以上专业技术职务的医师才允许下达
            If rsInput!类别ID & "" = "K" And gbln输血申请中级以上 Then
                If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                    MsgBox "启用了输血分级管理后，输血医嘱只有中级及以上专业技术职务医师才能下达。", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
            '中、西成药，卫材，检查(组合)，手术(组合)，输血，及其它诊疗项目
            Call AdviceSet诊疗项目(rsInput, lngRow, lng用法ID, lngGroupRow, strExtData, str摘要, str手术部位, bln备血)
            
            '新增后关联诊断的标记处理
            Call SetDiagFlag(vsAdvice.Row, 1)
            
            '自动设置一并给药
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                i = CheckAutoMerge(lngRow)
                If i = 1 Then
                    mblnRowMerge = True
                ElseIf i = 2 Then
                    mblnRowMerge = False
                    Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                    objControl.Checked = False
                End If
                If Not RowIn一并给药(lngRow) Then
                    If mblnRowMerge Then
                        '手工使一并给药
                        lngRowID = .RowData(lngRow)
                        Call MergeRow(lngPreRow, lngRow) '本来就是显示当前行的内容,不用再强行RowChange
                        lngRow = .FindRow(lngRowID, lngPreRow + 1)
                        '抗菌药物新增一行时一并给药的，应该缺省与上一行的“用药目的”和“用药理由”相同。
                        If Val(.TextMatrix(lngRow, COL_抗菌等级)) <> 0 Then
                            If lngRow > 1 Then
                                txt用药理由.Text = .TextMatrix(lngRow - 1, COL_用药理由)
                                If Val(.TextMatrix(lngRow - 1, COL_用药目的)) <> 0 Then
                                    cboDruPur.ListIndex = Val(.TextMatrix(lngRow - 1, COL_用药目的))
                                End If
                            End If
                        End If
                        '新录入的药品一并给药时，如果没有启用天数录入参数，但用法用量指定了单量则自动反算总量，根据第一条医嘱的天数来反算
                        If Val(.TextMatrix(lngRow, COL_单量)) > 0 And mbln天数 = False Then
                            Call GetRowScope(lngRow, lngBegin, lngEnd)
                            If Val(.TextMatrix(lngBegin, COL_总量)) > 0 And Val(.TextMatrix(lngBegin, COL_单量)) > 0 And .TextMatrix(lngBegin, COL_频率) <> "" _
                                And Val(.TextMatrix(lngBegin, COL_频率次数)) <> 0 And Val(.TextMatrix(lngBegin, COL_频率间隔)) <> 0 _
                                And Val(.TextMatrix(lngBegin, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngBegin, COL_门诊包装)) <> 0 Then
                                sng天数 = Calc缺省药品天数(Val(.TextMatrix(lngBegin, COL_总量)), Val(.TextMatrix(lngBegin, COL_单量)), _
                                                Val(.TextMatrix(lngBegin, COL_频率次数)), Val(.TextMatrix(lngBegin, COL_频率间隔)), .TextMatrix(lngBegin, COL_间隔单位), _
                                                Val(.TextMatrix(lngBegin, COL_剂量系数)), Val(.TextMatrix(lngBegin, COL_门诊包装)), _
                                                Val(.TextMatrix(lngBegin, COL_可否分零)))
                                .TextMatrix(lngRow, COL_总量) = ReGet药品总量(Val(.TextMatrix(lngRow, COL_总量)), Val(.TextMatrix(lngRow, COL_单量)), sng天数, lngRow)
                            End If
                        End If
                    ElseIf lngPreRow <> -1 Then
                        '自动使一并给药
                        Set objControl = cbsMain.FindControl(, conMenu_Merge, , True)
                        If objControl.Checked = True Then
                            If .TextMatrix(lngPreRow, COL_类别) = rsInput!类别ID Then
                                If RowIn一并给药(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                    mblnRowMerge = True
                                    cbsMain.RecalcLayout '即时刷新
                                    lngRowID = .RowData(lngRow)
                                    Call MergeRow(lngPreRow, lngRow, False)
                                    lngRow = .FindRow(lngRowID, lngPreRow + 1)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        '更新附件内容:以当前可见行为准
        If strAppend <> "" Then
            .TextMatrix(.Row, COL_附项) = strAppend
            .Cell(flexcpData, .Row, COL_附项) = 1 '表明需要重新写入(新增或修改)
            Call ReplaceAdviceAppend(.Row) '缺省替换其他医嘱的申请附项
        End If
        
        '输入西药可成药时自动增加皮试行:增加之后仍定位在当前药品
        If lng皮试ID <> 0 And Not bln皮试行 Then
            Call AdviceSet过敏试验(.Row, lng皮试ID) '注意用当前行,因为一并之后定位改变
        End If
        
        '重新自动调整行高
        Call .AutoSize(col_医嘱内容)
    End With
    
    mblnNoSave = True '标记为未保存
    
    '对保险对码进行检查
    Call GetInsureStr(strIDs1, strIDs2, str医嘱内容, vsAdvice.Row)
    strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, 1, strIDs1, strIDs2, str医嘱内容)
    If strMsg <> "" Then
        If gint医保对码 = 2 Then strMsg = strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。"
        vMsg = frmMsgBox.ShowMsgBox(strMsg, Me, True)
        If vMsg = vbIgnore Then mbln提醒对码 = False
    End If
    
    '调用外挂接口
    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
        If zlPluginAdviceEnter(vsAdvice.Row) = False Then
            vsAdvice.Refresh: Exit Function
        End If
    End If
    
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlPASSMap()
'功能:设置Pass VsAdvie及列映射
'注意:删除或修改下面列中数据时，请检查合理用药部件中的关联处理。
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "合理用药监测", True)
        mblnPass = Not mobjPassMap Is Nothing And Not gobjPass Is Nothing
    End If
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_门诊编辑
            .int场合 = mint场合
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .objCmdBar = cmdAlley
            
            Set .diags = .GetDiags
            
            Set .VSCOL = .GetVSCOL( _
                , COL_相关ID, COL_类别, COL_诊疗项目ID, COL_收费细目ID, col_医嘱内容, , COL_单量, COL_单量单位, COL_用法, _
                COL_天数, COL_婴儿, COL_开嘱时间, COL_开嘱医生, COL_开始时间, COL_开嘱科室ID, , COL_频率, COL_频率次数, COL_频率间隔, _
                COL_间隔单位, COL_警示, COL_序号, COL_状态, COL_EDIT, , , , COL_执行性质, COL_标本部位, _
                , , , , , COL_总量, COL_总量单位, COL_医生嘱托, COL_用药目的, COL_操作类型, , _
                COL_用药理由, COL_标志, COL_处方序号, COL_执行分类)
            Set .PassPati = .GetPatient()
            Call zlPASSPati
        End With
    End If
End Sub

Private Sub zlPASSPati()
'功能:设置病人信息
    If Not mobjPassMap Is Nothing Then
        With mobjPassMap.PassPati
            .int婴儿 = IIF(cbo婴儿.ListIndex = -1, 0, cbo婴儿.ListIndex)  '缺省新增病人为0
            .dbl标识号 = mdbl门诊号
            .Dat出生日期 = mDat出生日期
            .lng病人ID = mlng病人ID
            .lng挂号ID = mlng挂号ID
            .str挂号单 = mstr挂号单
            .str性别 = mstr性别
            .str姓名 = mstr姓名
        End With
    End If
End Sub

Private Sub zlPassDrags()
'功能:设置病人诊断信息
    Dim i As Long
    
    If Not mobjPassMap Is Nothing Then
        Set mobjPassMap.diags = Nothing '清空重新赋值
        Set mobjPassMap.diags = mobjPassMap.GetDiags
        With vsDiag
            For i = .FixedRows To .Rows - 1
                mobjPassMap.diags.Add .TextMatrix(i, col诊断), .TextMatrix(i, col诊断编码), .TextMatrix(i, col疾病编码), "_" & i
            Next
        End With
    End If
End Sub

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'功能：将两行设置为一并给药
'参数：lngRow1=前面行,可能本来已经属于一并给药
'      lngRow2=当前行
'说明：设置完成后,表格仍定位在原lngRow2的当前行
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    Dim lngDiag As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            
            '记录当前行的关联诊断,一并后以此为准
            lngDiag = AdviceHaveDiag(lngRow2)
            
            lngTmp = .RowData(lngRow2) '记录以再定位到当前行
            '先取消之前的一并给药
            If RowIn一并给药(lngRow1) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow1, COL_相关ID)), lngBegin, lngEnd)
                Call AdviceSet单独给药(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSet一并给药(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            
            '以一并之前的当前行为准恢复关联诊断
            '一并过程中前面的药品诊断关联已被DeleteRow移除
            If lngDiag <> -1 Then
                Call SetDiagFlag(.Row, 1, lngDiag)
            End If
            
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'功能：将指定行从一并给药中独立出来(该组一并给药必须至少包含三行)
'参数：lngRow=当前行,且为一并给药中的最后一药品行
'说明：设置完成后,表格仍定位在原lngRow的当前行
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '记录用于恢复定位当前行
        Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
        
        '先取消整个的一并给药
        Call AdviceSet单独给药(lngBegin, lngEnd)
        
        '再设置除最后行外的行为一并给药
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSet一并给药(lngBegin, lngEnd)
        
        '恢复当前行
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Private Sub AdviceSet复制医嘱(ByVal lng病人ID As Long, ByVal str挂号单 As String, ByVal strIDs As String, Optional ByVal blnHistory As Boolean)
'功能：复制指定病人的指定医嘱产生成为新医嘱
'说明：可供外部调用,调用之前处于新增医嘱行
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln配方 As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim curDate As Date, blnHide As Boolean
    Dim lng开嘱科室ID As Long, dbl相关ID As Double
    Dim lng序号 As Long, intCount As Integer
    Dim lngRow As Long
    Dim i As Long, j As Long
    
    Dim lng西药房ID As Long, lng成药房ID As Long, lng中药房ID As Long, lng发料部门ID As Long
    Dim str药房IDs As String, str附项 As String
    Dim bln药品小数输入 As Boolean
    Dim str高危药品 As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select a.ID+0.1 As ID,Decode(a.相关id,Null,Null,a.相关id+0.1) As 相关id,Nvl(A.婴儿,0) as 婴儿,A.序号,A.医嘱期效," & _
        " A.医嘱状态,A.诊疗类别,A.诊疗项目ID,B.名称,A.标本部位,A.检查方法,A.执行标记,A.收费细目ID," & _
        " A.开始执行时间,Nvl(B.名称,A.医嘱内容) 医嘱内容,A.医生嘱托,A.单次用量,A.天数,A.总给予量,B.计算单位," & _
        " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,B.计算方式,B.执行频率,B.操作类型,B.单独应用,B.执行分类," & _
        " B.计价性质,A.执行时间方案,Decode(nvl(Instr(',5,6,7,',a.诊疗类别),0),0,b.执行科室,a.执行性质) as 执行性质,A.执行科室ID,A.开嘱科室ID,A.开嘱医生,A.开嘱时间," & _
        " A.紧急标志,C.毒理分类,C.抗生素,C.药品剂型,B.录入限量,C.处方限量,C.处方职务," & _
        " D.剂量系数,D.门诊包装,D.门诊单位,F.计算单位 as 散装单位,E.跟踪在用,D.门诊可否分零 As 可否分零,a.配方ID,c.临床自管药,d.高危药品,a.组合项目ID,c.溶媒,d.基本药物" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,药品特性 C,药品规格 D,材料特性 E,收费项目目录 F" & _
        " Where Nvl(A.医嘱期效,0)=1 And A.诊疗项目ID=B.ID" & _
        " And A.诊疗项目ID=C.药名ID(+) And A.收费细目ID=D.药品ID(+)" & _
        " And A.收费细目ID=E.材料ID(+) And E.材料ID=F.ID(+)" & _
        " And A.病人ID+0=[1] And A.挂号单=[2]" & _
        " And Instr([3],','||A.ID||',')>0" & _
        " Order by 婴儿,序号"
    If blnHistory Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, str挂号单, "," & strIDs & ",")
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        If rsTmp!诊疗类别 = "Z" And (rsTmp!操作类型 = "1" Or rsTmp!操作类型 = "2") Then
            If CheckInHosAdvice Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        
    
        lngBegin = vsAdvice.Row '开始新增行
        lng序号 = GetCurRow序号(lngBegin) '起始序号
        intCount = 0 '已经设置的行数
        curDate = zlDatabase.Currentdate
        lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
        '药品总量是否可以输入小数，中药不可以输入小数，因此只检查西药与中成药
        bln药品小数输入 = InStr(GetInsidePrivs(p住院医嘱下达), "药品小数输入") > 0
            
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            For i = lngBegin To rsTmp.RecordCount + lngBegin - 1
                If i > lngBegin Then .AddItem "", i

                bln配方 = False
                
                .RowData(i) = -1 * rsTmp!ID
                If Not IsNull(rsTmp!相关ID) Then
                    .TextMatrix(i, COL_相关ID) = -1 * rsTmp!相关ID
                End If
                .TextMatrix(i, COL_序号) = lng序号 + intCount
                
                .TextMatrix(i, COL_EDIT) = 1 '新增
                .Cell(flexcpData, i, COL_EDIT) = CStr(lng病人ID & "," & str挂号单) '记录相关的复制项目
                .TextMatrix(i, COL_状态) = 1 '新开
                .TextMatrix(i, COL_婴儿) = cbo婴儿.ListIndex
                .TextMatrix(i, COL_类别) = rsTmp!诊疗类别
                .TextMatrix(i, COL_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(i, COL_名称) = rsTmp!名称
                .TextMatrix(i, COL_标本部位) = Nvl(rsTmp!标本部位)
                .TextMatrix(i, COL_检查方法) = Nvl(rsTmp!检查方法)
                .TextMatrix(i, COL_执行标记) = Nvl(rsTmp!执行标记, 0)
                .TextMatrix(i, COL_收费细目ID) = Nvl(rsTmp!收费细目ID)
                .TextMatrix(i, col_医嘱内容) = Nvl(rsTmp!医嘱内容)
                .TextMatrix(i, COL_医生嘱托) = Nvl(rsTmp!医生嘱托)
                .Cell(flexcpData, i, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, Val(.TextMatrix(i, COL_收费细目ID)), "", 0, "", .TextMatrix(i, COL_诊疗项目ID))
                
                .TextMatrix(i, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
                .TextMatrix(i, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                .TextMatrix(i, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                .TextMatrix(i, COL_操作类型) = Nvl(rsTmp!操作类型)
                .TextMatrix(i, COL_单独应用) = Nvl(rsTmp!单独应用)
                .TextMatrix(i, COL_执行分类) = Nvl(rsTmp!执行分类, 0)
                .TextMatrix(i, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(i, COL_抗菌等级) = Val("" & rsTmp!抗生素)
                .TextMatrix(i, COL_配方ID) = Nvl(rsTmp!配方ID)
                .TextMatrix(i, COL_临床自管药) = rsTmp!临床自管药 & ""
                .TextMatrix(i, COL_组合项目ID) = rsTmp!组合项目ID & ""
                .TextMatrix(i, COL_高危药品) = Val(rsTmp!高危药品 & "")
                .TextMatrix(i, COL_是否溶媒) = Val(rsTmp!溶媒 & "")
                .TextMatrix(i, COL_基本药物) = rsTmp!基本药物 & ""
                If Val(.TextMatrix(i, COL_高危药品)) <> 0 Then
                    str高危药品 = str高危药品 & vbCrLf & .TextMatrix(i, col_医嘱内容) & ":" & Decode(Val(.TextMatrix(i, COL_高危药品)), 1, "A", 2, "B", 3, "C", "") & "级；"
                End If
                
                .TextMatrix(i, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!处方限量)
                Else
                    .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!录入限量)
                End If
                .TextMatrix(i, COL_处方职务) = Nvl(rsTmp!处方职务)
                
                If InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 Then
                    .TextMatrix(i, COL_剂量系数) = Nvl(rsTmp!剂量系数)
                    .TextMatrix(i, COL_门诊包装) = Nvl(rsTmp!门诊包装)
                    .TextMatrix(i, COL_门诊单位) = Nvl(rsTmp!门诊单位)
                    If Not IsNull(rsTmp!剂量系数) Then
                        .TextMatrix(i, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
                    End If
                ElseIf .TextMatrix(i, COL_类别) = "4" Then
                    .TextMatrix(i, COL_剂量系数) = 1
                    .TextMatrix(i, COL_门诊包装) = 1
                    .TextMatrix(i, COL_门诊单位) = Nvl(rsTmp!散装单位)
                    .TextMatrix(i, COL_跟踪在用) = Nvl(rsTmp!跟踪在用)
                End If
                
                If IsDate(txt开始时间.Text) Then
                    .TextMatrix(i, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开始时间) = txt开始时间.Text
                    
                    '手术/输血时间：复制时缺省与开始时间相同,在标本部位处理后
                    If rsTmp!诊疗类别 = "K" Or rsTmp!诊疗类别 = "F" Or rsTmp!诊疗类别 = "G" _
                        And Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                        .TextMatrix(i, COL_手术时间) = txt开始时间.Text
                    End If
                End If
                
                .TextMatrix(i, COL_频率) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, COL_频率次数) = Nvl(rsTmp!频率次数)
                .TextMatrix(i, COL_频率间隔) = Nvl(rsTmp!频率间隔)
                .TextMatrix(i, COL_间隔单位) = Nvl(rsTmp!间隔单位)
                .TextMatrix(i, COL_执行时间) = Nvl(rsTmp!执行时间方案)
                
                .TextMatrix(i, COL_执行性质) = Nvl(rsTmp!执行性质, 0)
                
                '处理执行科室
                If rsTmp!诊疗类别 = "Z" Then
                    .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID)
                ElseIf InStr(",0,5,", Nvl(rsTmp!执行性质, 0)) = 0 Then
                    If Nvl(rsTmp!执行科室ID, 0) <> 0 Then
                        If InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                            str药房IDs = Get可用药房IDs(rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), mlng病人科室id, 1)
                            If InStr("," & str药房IDs & ",", "," & rsTmp!执行科室ID & ",") > 0 Then
                                .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                            End If
                        ElseIf Nvl(rsTmp!诊疗类别) = "4" Then
                            str药房IDs = Get可用发料部门IDs(Nvl(rsTmp!收费细目ID, 0), mlng病人科室id, 1)
                            If InStr("," & str药房IDs & ",", "," & rsTmp!执行科室ID & ",") > 0 Then
                                .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_执行性质)) = 4 Then
                            '4-指定科室时才取,其它的固定生成
                            .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID, 0)
                            
                            '检查执行科室的有效性
                            If Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                                If CheckExecDeptValidate(Val(.TextMatrix(i, COL_执行科室ID)), mlng病人科室id, 1, Val(.TextMatrix(i, COL_诊疗项目ID))) = False Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                        '药品、卫材类的整个成套相同
                        If rsTmp!诊疗类别 = "5" Then
                            If lng西药房ID = 0 Then
                                lng西药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng西药房ID
                        ElseIf rsTmp!诊疗类别 = "6" Then
                            If lng成药房ID = 0 Then
                                lng成药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng成药房ID
                        ElseIf rsTmp!诊疗类别 = "7" Then
                            If lng中药房ID = 0 Then
                                lng中药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsTmp!诊疗类别, rsTmp!诊疗项目ID, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng中药房ID
                        ElseIf Nvl(rsTmp!诊疗类别) = "4" Then
                            If lng发料部门ID = 0 Then
                                lng发料部门ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!诊疗类别, Nvl(rsTmp!收费细目ID, 0), 4, mlng病人科室id, lng开嘱科室ID, 1, , 1)
                            End If
                            .TextMatrix(i, COL_执行科室ID) = lng发料部门ID
                        Else
                            .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsTmp!诊疗类别, rsTmp!诊疗项目ID, 0, Nvl(rsTmp!执行性质, 0), mlng病人科室id, lng开嘱科室ID, 1, 1)
                        End If
                    End If
                End If
                
                If rsTmp!诊疗类别 = "E" Then
                    If Nvl(rsTmp!相关ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_相关ID)) = -1 * rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是成药的给药途径,可能是一并给药的
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = -1 * rsTmp!ID Then
                                    '显示给药途径
                                    .TextMatrix(j, COL_用法) = rsTmp!名称 & rsTmp!医生嘱托
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是中药配方的用法,即配方显示行
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                            bln配方 = True
                        ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                        End If
                    ElseIf Not IsNull(rsTmp!相关ID) And .TextMatrix(i - 1, COL_类别) = "K" And -1 * Nvl(rsTmp!相关ID, 0) = .RowData(i - 1) Then
                        '当前记录是输血途径行
                        .TextMatrix(i - 1, COL_用法) = rsTmp!名称
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        '当前记录是中药配方煎法行
                        bln配方 = True
                    End If
                ElseIf rsTmp!诊疗类别 = "7" Then
                    bln配方 = True
                End If
                
                '单量
                .TextMatrix(i, COL_单量) = FormatEx(Nvl(rsTmp!单次用量), 5)
                If Nvl(rsTmp!诊疗类别) = "4" Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!散装单位)
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 _
                    Or (Val(.TextMatrix(i, COL_频率性质)) = 0 And InStr(",1,2,", Nvl(rsTmp!计算方式, 0)) > 0) Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!计算单位)
                End If
                
                '天数
                If mbln天数 Then
                    .TextMatrix(i, COL_天数) = Nvl(rsTmp!天数)
                End If

                '总量
                If InStr(",5,6,", rsTmp!诊疗类别) > 0 Then
                    '成药临嘱有总量,以零售单位存放,门诊单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!门诊包装) Then
                        .TextMatrix(i, COL_总量) = FormatEx(rsTmp!总给予量 / rsTmp!门诊包装, 5)
                    End If
                    .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!门诊单位)
                    
                    If Val(.TextMatrix(i, COL_可否分零)) = 0 And Not bln药品小数输入 And InStr(.TextMatrix(i, COL_总量), ".") > 0 Then
                        .TextMatrix(i, COL_总量) = IntEx(Val(.TextMatrix(i, COL_总量)))
                    End If
                    
                    '超量说明不复制
                    Call Set用药天数是否超期(i)
                ElseIf bln配方 Then
                    If Not IsNull(rsTmp!总给予量) Then .TextMatrix(i, COL_总量) = rsTmp!总给予量
                    
                    .TextMatrix(i, COL_总量单位) = "付" '中药配方总量单位为"付"
                    
                    If rsTmp!诊疗类别 = "E" And rsTmp!操作类型 = "4" Then   '中药用法
                        Call Set用药天数是否超期(i)
                    End If
                Else
                    '其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then .TextMatrix(i, COL_总量) = rsTmp!总给予量
                        
                    If Nvl(rsTmp!诊疗类别) = "4" Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!散装单位)
                    Else
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If
                
                '抗菌药物缺省用药目的
                If Val(.TextMatrix(i, COL_抗菌等级)) > 0 Then .TextMatrix(i, COL_用药目的) = mstrPurMed
                
                .TextMatrix(i, COL_标志) = chk紧急.value '可以在界面先统一设置为紧急
                If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(i, COL_抗菌等级)) And .TextMatrix(i, COL_标志) <> "1" Then

                    .TextMatrix(i, COL_审核状态) = 1
                End If
                '处理输血医嘱审核状态
                If .TextMatrix(i, COL_类别) = "E" And .TextMatrix(i, COL_操作类型) = "8" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    strSQL = ""
                    strSQL = GetBloodState(IIF(.TextMatrix(i, COL_标志) = "1", 1, 0), Val(.TextMatrix(i, COL_执行分类)))
                    .TextMatrix(i - 1, COL_审核状态) = strSQL
                    .TextMatrix(i, COL_审核状态) = strSQL
                End If
                
                .TextMatrix(i, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(i, COL_开嘱科室ID) = lng开嘱科室ID
                .TextMatrix(i, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                Call SetRow标志图标(i, 1)
                
                
                '毒麻精药品标识:中药配方及组成味中药不处理
                If InStr(",5,6,", rsTmp!诊疗类别) > 0 And Not IsNull(rsTmp!毒理分类) Then
                    If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", rsTmp!毒理分类) > 0 Then
                        .Cell(flexcpFontBold, i, col_医嘱内容) = True
                    End If
                End If
                
                lngEnd = i
                intCount = intCount + 1
                
                rsTmp.MoveNext
            Next
            
            '产生假医嘱ID
            For i = lngBegin To lngEnd
                dbl相关ID = .RowData(i)
                .RowData(i) = GetNext医嘱ID
                For j = i - 1 To lngBegin Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = dbl相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngEnd
                    If Val(.TextMatrix(j, COL_相关ID)) = dbl相关ID Then
                        .TextMatrix(j, COL_相关ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
            If gblnOut必用 Then Call MakeAppNo(2, lngBegin, lngEnd)
            Call Set医嘱超量(lngBegin, lngEnd)
            '调整受影响行的序号
            Call AdviceSet医嘱序号(lngEnd + 1, intCount)
            
            '显示/隐藏行
            lngRow = 0
            For i = lngBegin To lngEnd
                blnHide = False
                If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '处理医嘱内容的变化
                If Not .RowHidden(i) Then
                    .TextMatrix(i, col_医嘱内容) = AdviceTextMake(i)
                End If
                
                '预先计算诊疗单价
                If Not .RowHidden(i) And .TextMatrix(i, COL_单价) = "" Then
                    .TextMatrix(i, COL_单价) = GetItemPrice(i)
                End If
                
                '产生新嘱时，在可见行读取诊疗项目附项串，用于检查，不用于保存
                If Not .RowHidden(i) Then
                    If Not RowIn配方行(i) Then
                        If RowIn检验行(i) Then
                            j = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                            If j <> -1 Then
                                .TextMatrix(i, COL_附项) = Get医嘱项目附件(Val(.TextMatrix(j, COL_诊疗项目ID)), 1)
                            End If
                        Else
                            .TextMatrix(i, COL_附项) = Get医嘱项目附件(Val(.TextMatrix(i, COL_诊疗项目ID)), 1)
                        End If
                        If .TextMatrix(i, COL_附项) <> "" Then
                            str附项 = str附项 & vbCrLf & "●" & .TextMatrix(i, col_医嘱内容)
                        End If
                    End If
                End If
                
                '新增后关联诊断的标记处理
                If Not .RowHidden(i) Then
                    Call SetDiagFlag(i, 1)
                End If
            Next
            
            '图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            
            .Row = lngRow: .Col = col_医嘱内容
            
            Call .AutoSize(col_医嘱内容)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
        mblnNoSave = True '标记为未保存
    End If

    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call CalcAdviceMoney '显示新开医嘱金额

    Screen.MousePointer = 0
    
    If str附项 <> "" Then
        MsgBox "以下医嘱需要填写申请附项，请注意填写：" & vbCrLf & str附项, vbInformation, gstrSysName
    End If
    If str高危药品 <> "" Then
        MsgBox "以下医嘱是高危药品：" & str高危药品, vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdviceSet成套项目(ByVal lng成套ID As Long, ByVal lngRow As Long, Optional ByVal str序号 As String)
'功能：输入成套项目(包括一并给药,检查组合,手术附加,中药配方)
'参数：lngRow=空的输入行(可能是插入的新行,但不位于一并给药中间)
    Dim rsItems As New ADODB.Recordset
    Dim rs规格 As New ADODB.Recordset
    Dim rs材料 As New ADODB.Recordset
    Dim rs疗程 As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng序号 As Long
    Dim lngPreRow As Long, vCurDate As Date, lngTmp As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim bln给药途径 As Boolean, bln采集方法 As Boolean, bln输血途径 As Boolean
    Dim bln中药用法 As Boolean, bln中药煎法 As Boolean, bln配方 As Boolean
    Dim lng西药房ID As Long, lng成药房ID As Long, lng中药房ID As Long
    Dim dbl相关ID As Double, int适用范围 As Integer, str频率 As String
    Dim str医生 As String, lng医生ID As Long
    Dim lng倍数 As Long, vBookMark As Variant, str药房IDs As String
    Dim sng天数 As Single, strSQL序号 As String, str附项 As String
    Dim int频率性质 As Integer, lng发料部门ID As Long, blnAdd As Boolean
    Dim str高危药品 As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '产生序号过滤串
    If str序号 <> "" Then
        If Left(str序号, 1) = "+" Then
            strSQL序号 = " And Instr([2],','||A.序号||',')>0"
        ElseIf Left(str序号, 1) = "-" Then
            strSQL序号 = " And Instr([2],','||A.序号||',')=0"
        End If
    End If
    
    '药品规格信息:虽然存了收费细目ID,但以前的数据没存
    strSQL = "Select A.序号,B.药名ID,B.药品ID,B.剂量系数,B.门诊包装,B.门诊单位,b.基本药物," & _
        " B.门诊可否分零 As 可否分零,C.编码,Nvl(D.名称,C.名称) as 名称,C.规格,C.产地,b.高危药品" & _
        " From 诊疗项目组合 A,药品规格 B,收费项目目录 C,收费项目别名 D" & _
        " Where A.诊疗项目ID=B.药名ID And B.药品ID=C.ID" & _
        " And C.ID=D.收费细目ID(+) And D.码类(+)=1 And D.性质(+)=[3]" & _
        " And A.诊疗组合ID=[1]" & strSQL序号 & _
        " Order by A.序号,C.编码"
    Set rs规格 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",", IIF(gbyt药品名称显示 = 0, 1, 3))
    
    '卫材信息
    strSQL = "Select A.序号,B.材料ID,B.跟踪在用,C.名称,C.计算单位" & _
        " From 诊疗项目组合 A,材料特性 B,收费项目目录 C" & _
        " Where A.收费细目ID=B.材料ID And B.材料ID=C.ID" & _
        " And A.诊疗组合ID=[1]" & strSQL序号 & _
        " Order by A.序号,C.编码"
    Set rs材料 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    
    '成药疗程信息(因成套中无直接对应配方,中药取不到疗程)
    strSQL = "Select Distinct A.诊疗项目ID,C.疗程" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,诊疗用法用量 C" & _
        " Where A.诊疗项目ID=B.ID And B.类别 IN('5','6')" & _
        " And A.诊疗项目ID=C.项目ID And A.诊疗组合ID=[1]" & strSQL序号
    Set rs疗程 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    
    '按序号排列后应该与医嘱编辑时的次序一致
    '排开执行频率为持续性的长嘱(这种方法不对,只取临嘱),所有方案转为临嘱处理
    strSQL = "Select 1 as 期效,a.序号+0.1 As 序号,Decode(a.相关序号,Null,Null,a.相关序号+0.1) As 相关序号,A.诊疗项目ID,A.收费细目ID,A.医嘱内容,A.天数,A.总给予量,A.单次用量," & _
        " A.医生嘱托,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.执行科室ID,Nvl(B.类别,'*') 类别,B.名称," & _
        " B.计算单位,Decode(B.类别,'D',A.标本部位,Nvl(A.标本部位,B.标本部位)) as 标本部位,A.检查方法," & _
        " A.时间方案,Nvl(A.执行性质,B.执行科室) as 执行性质,B.计价性质,B.单独应用,B.操作类型,B.执行分类,B.计算方式," & _
        " B.执行频率,B.录入限量,C.处方限量,C.处方职务,C.毒理分类,C.抗生素,C.药品剂型,A.配方ID,C.临床自管药,A.组合项目ID,C.溶媒" & _
        " From 诊疗项目组合 A,诊疗项目目录 B,药品特性 C,收费项目目录 D" & _
        " Where A.诊疗项目ID=B.ID(+) And A.诊疗项目ID=C.药名ID(+)  And d.id(+)=a.收费细目ID " & _
        " And A.期效=1 And A.诊疗组合ID=[1]" & strSQL序号 & _
        " And (NVL(d.撤档时间,b.撤档时间) is null or NVL(d.撤档时间,b.撤档时间) = To_Date('3000/1/1', 'yyyy/mm/dd') Or Not (b.类别 = '7' Or b.类别 = 'E' And b.执行分类 = 0 And b.操作类型 = '3'))" & _
        " Order by A.序号"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng成套ID, "," & Mid(str序号, 2) & ",")
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        
        lngPreRow = GetPreRow(lngRow) '前一参照行
        intCount = 0 '已经设置的行数
        lng序号 = GetCurRow序号(lngRow) '起始序号
        vCurDate = zlDatabase.Currentdate
        
        For i = 1 To rsItems.RecordCount
            blnAdd = True
            
            '检查是否存在有效的住院医嘱或留观医嘱
            If rsItems!类别 = "Z" And InStr(",1,2,", "," & rsItems!操作类型 & ",") > 0 Then
                If CheckInHosAdvice Then
                    blnAdd = False
                End If
            End If
            
            If blnAdd Then
                lngCurRow = lngRow + intCount
                If lngCurRow > lngRow Then .AddItem "", lngCurRow
                 
                '记录相对ID
                .RowData(lngCurRow) = -1 * rsItems!序号
                If Not IsNull(rsItems!相关序号) Then
                    .TextMatrix(lngCurRow, COL_相关ID) = -1 * rsItems!相关序号
                End If
                
                .TextMatrix(lngCurRow, COL_EDIT) = 1 '新增的
                .Cell(flexcpData, lngCurRow, COL_EDIT) = lng成套ID '记录相关的成套项目
                
                .TextMatrix(lngCurRow, COL_婴儿) = cbo婴儿.ListIndex
                .TextMatrix(lngCurRow, COL_序号) = lng序号 + intCount
                .TextMatrix(lngCurRow, COL_状态) = 1 '新开
                .TextMatrix(lngCurRow, COL_类别) = rsItems!类别
                .TextMatrix(lngCurRow, COL_诊疗项目ID) = Nvl(rsItems!诊疗项目ID, 0)
                .TextMatrix(lngCurRow, COL_名称) = Nvl(rsItems!名称)
                .TextMatrix(lngCurRow, COL_标本部位) = Nvl(rsItems!标本部位)
                .TextMatrix(lngCurRow, COL_检查方法) = Nvl(rsItems!检查方法)
                
                If IsDate(txt开始时间.Text) Then
                    .TextMatrix(lngCurRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, lngCurRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    
                    '手术/输血时间：复制时缺省与开始时间相同,在标本部位处理后
                    If rsItems!类别 = "K" Or rsItems!类别 = "F" Or rsItems!类别 = "G" _
                        And Val(.TextMatrix(lngCurRow, COL_相关ID)) = Val(.TextMatrix(lngCurRow - 1, COL_相关ID)) Then
                        .TextMatrix(lngCurRow, COL_手术时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    End If
                End If
                
                '其它
                .TextMatrix(lngCurRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
                .TextMatrix(lngCurRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
                .TextMatrix(lngCurRow, COL_操作类型) = Nvl(rsItems!操作类型)
                .TextMatrix(lngCurRow, COL_单独应用) = Nvl(rsItems!单独应用)
                .TextMatrix(lngCurRow, COL_执行分类) = Nvl(rsItems!执行分类, 0)
                .TextMatrix(lngCurRow, COL_毒理分类) = Nvl(rsItems!毒理分类)
                .TextMatrix(lngCurRow, COL_抗菌等级) = Val("" & rsItems!抗生素)
                .TextMatrix(lngCurRow, COL_配方ID) = Nvl(rsItems!配方ID)
                .TextMatrix(lngCurRow, COL_临床自管药) = rsItems!临床自管药 & ""
                .TextMatrix(lngCurRow, COL_组合项目ID) = rsItems!组合项目ID & ""
                .TextMatrix(lngCurRow, COL_是否溶媒) = Val(rsItems!溶媒 & "")
                
                .TextMatrix(lngCurRow, COL_药品剂型) = Nvl(rsItems!药品剂型)
                If InStr(",5,6,7,", rsItems!类别) > 0 Then
                    .TextMatrix(lngCurRow, COL_处方限量) = Nvl(rsItems!处方限量)
                Else
                    .TextMatrix(lngCurRow, COL_处方限量) = Nvl(rsItems!录入限量)
                End If
                .TextMatrix(lngCurRow, COL_处方职务) = Nvl(rsItems!处方职务)
                
                '药品规格信息:中草药肯定有,成药按单量与剂量单位自动匹配
                lng倍数 = 0: vBookMark = 0
                If rsItems!类别 = "7" Or (InStr(",5,6,", rsItems!类别) > 0) Then
                    If Not IsNull(rsItems!收费细目ID) Then '可能以前未保存收费细目ID
                        rs规格.Filter = "药品ID=" & rsItems!收费细目ID
                    Else
                        rs规格.Filter = "药名ID=" & rsItems!诊疗项目ID
                    End If
                    If Not rs规格.EOF Then
                        If IsNull(rsItems!收费细目ID) Then
                            '取剂量系数为单量的最小整倍数的那一个规格
                            If CInt(Nvl(rsItems!单次用量, 0)) <> 0 Then
                                Do While Not rs规格.EOF
                                    If rs规格!剂量系数 / rsItems!单次用量 = Int(rs规格!剂量系数 / rsItems!单次用量) Then
                                        If rs规格!剂量系数 / rsItems!单次用量 < lng倍数 Or lng倍数 = 0 Then
                                            vBookMark = rs规格.Bookmark
                                            lng倍数 = rs规格!剂量系数 / rsItems!单次用量
                                        End If
                                    End If
                                    rs规格.MoveNext
                                Loop
                                If vBookMark <> 0 Then rs规格.Bookmark = vBookMark
                            End If
                            If rs规格.EOF Then rs规格.MoveFirst
                        End If
                        .TextMatrix(lngCurRow, COL_名称) = Nvl(rs规格!名称)
                        .TextMatrix(lngCurRow, COL_收费细目ID) = rs规格!药品ID
                        .TextMatrix(lngCurRow, COL_剂量系数) = Nvl(rs规格!剂量系数)
                        .TextMatrix(lngCurRow, COL_门诊包装) = Nvl(rs规格!门诊包装)
                        .TextMatrix(lngCurRow, COL_门诊单位) = Nvl(rs规格!门诊单位)
                        .TextMatrix(lngCurRow, COL_可否分零) = Nvl(rs规格!可否分零, 0)
                        .TextMatrix(lngCurRow, COL_高危药品) = Nvl(rs规格!高危药品, 0)
                        .TextMatrix(lngCurRow, COL_基本药物) = rs规格!基本药物 & ""
                        If Val(.TextMatrix(lngCurRow, COL_高危药品)) <> 0 Then
                            str高危药品 = str高危药品 & vbCrLf & rsItems!名称 & ":" & Decode(Val(.TextMatrix(lngCurRow, COL_高危药品)), 1, "A", 2, "B", 3, "C", "") & "级；"
                        End If
                    End If
                ElseIf rsItems!类别 = "4" Then
                    rs材料.Filter = "材料ID=" & Nvl(rsItems!收费细目ID, 0)
                    If Not rs材料.EOF Then
                        .TextMatrix(lngCurRow, COL_名称) = Nvl(rs材料!名称)
                        .TextMatrix(lngCurRow, COL_门诊单位) = Nvl(rs材料!计算单位) '散装单位
                        .TextMatrix(lngCurRow, COL_跟踪在用) = Nvl(rs材料!跟踪在用, 0)
                    End If
                    .TextMatrix(lngCurRow, COL_剂量系数) = 1
                    .TextMatrix(lngCurRow, COL_门诊包装) = 1
                    .TextMatrix(lngCurRow, COL_收费细目ID) = Nvl(rsItems!收费细目ID, 0)
                End If
                                    
                '判断是否特定行
                bln给药途径 = False: bln采集方法 = False: bln输血途径 = False
                bln中药用法 = False: bln中药煎法 = False: bln配方 = False
                If rsItems!类别 = "E" Then
                    If IsNull(rsItems!相关序号) Then
                        If Val(.TextMatrix(lngCurRow - 1, COL_相关ID)) = .RowData(lngCurRow) Then
                            If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_类别)) > 0 Then
                                bln给药途径 = True
                            ElseIf .TextMatrix(lngCurRow - 1, COL_类别) = "C" Then
                                bln采集方法 = True
                            Else
                                bln中药用法 = True
                            End If
                        End If
                    ElseIf .TextMatrix(lngCurRow - 1, COL_类别) = "K" And .RowData(lngCurRow - 1) = Val(.TextMatrix(lngCurRow, COL_相关ID)) Then
                        bln输血途径 = True
                    Else
                        bln中药煎法 = True
                    End If
                End If
                If rsItems!类别 = "7" Or bln中药煎法 Or bln中药用法 Then bln配方 = True
                        
                '获取当前项目的适用范围
                If bln采集方法 Then
                    '采集方法以检验项目的为准
                    lngTmp = .FindRow(CStr(.RowData(lngCurRow)), , COL_相关ID)
                    int频率性质 = .TextMatrix(lngTmp, COL_频率性质)
                Else
                    int频率性质 = Nvl(rsItems!执行频率, 0)
                End If
                If bln配方 Then
                    int适用范围 = 2 '中药配方(包括煎法,用法)用中医
    '            ElseIf bln采集方法 Then
    '                int适用范围 = -1 '设置与检验项目相同:一次性
                ElseIf int频率性质 = 0 Or bln给药途径 _
                    Or InStr(",5,6,", .TextMatrix(lngCurRow, COL_类别)) > 0 Then
                    int适用范围 = 1 '"可选频率"或成药(包括给药途径)用西医
                ElseIf int频率性质 = 1 Then
                    int适用范围 = -1 '一次性
                ElseIf int频率性质 = 2 Then
                    int适用范围 = -2 '持续性
                End If
                        
                '频率,频率次数,频率间隔,间隔单位
                .TextMatrix(lngCurRow, COL_频率性质) = int频率性质
                If Not IsNull(rsItems!执行频次) Then
                    If Check频率可用(Nvl(rsItems!诊疗项目ID, 0), int适用范围, Nvl(rsItems!执行频次)) Then
                        If Get频率信息_名称(rsItems!执行频次, int频率次数, int频率间隔, str间隔单位, CStr(int适用范围)) Then
                            .TextMatrix(lngCurRow, COL_频率) = rsItems!执行频次
                            .TextMatrix(lngCurRow, COL_频率次数) = int频率次数
                            .TextMatrix(lngCurRow, COL_频率间隔) = int频率间隔
                            .TextMatrix(lngCurRow, COL_间隔单位) = str间隔单位
                        End If
                    End If
                End If
                If .TextMatrix(lngCurRow, COL_频率) = "" And Not IsNull(rsItems!诊疗项目ID) Then '取缺省的
                    Call Get缺省频率(Nvl(rsItems!诊疗项目ID, 0), int适用范围, str频率, int频率次数, int频率间隔, str间隔单位)
                    .TextMatrix(lngCurRow, COL_频率) = str频率
                    .TextMatrix(lngCurRow, COL_频率次数) = int频率次数
                    .TextMatrix(lngCurRow, COL_频率间隔) = int频率间隔
                    .TextMatrix(lngCurRow, COL_间隔单位) = str间隔单位
                End If
                
                '单量
                .TextMatrix(lngCurRow, COL_单量) = FormatEx(Nvl(rsItems!单次用量), 5)
                If rsItems!类别 = "4" Then
                    .TextMatrix(lngCurRow, COL_单量单位) = .TextMatrix(lngCurRow, COL_门诊单位) '散装单位
                ElseIf bln中药用法 Then
                    .TextMatrix(lngCurRow, COL_单量单位) = ""
                Else
                    If InStr(",5,6,7,", rsItems!类别) > 0 Or (int频率性质 = 0 And InStr(",1,2,", Nvl(rsItems!计算方式, 0)) > 0) Then
                        .TextMatrix(lngCurRow, COL_单量单位) = Nvl(rsItems!计算单位)
                    End If
                End If
                
                '总量
                If InStr(",5,6,", rsItems!类别) > 0 Then
                    '成药临嘱(有对应规格)
                    .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow, COL_门诊单位)
                                        
                    sng天数 = Nvl(rsItems!天数, msng天数)
                    If mbln天数 Then
                        If .TextMatrix(lngCurRow, COL_间隔单位) = "周" Then
                            If 7 > sng天数 Then sng天数 = 7
                        ElseIf .TextMatrix(lngCurRow, COL_间隔单位) = "天" Then
                            If Val(.TextMatrix(lngCurRow, COL_频率间隔)) > sng天数 Then
                                sng天数 = Val(.TextMatrix(lngCurRow, COL_频率间隔))
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_间隔单位) = "小时" Then
                            If Val(.TextMatrix(lngCurRow, COL_频率间隔)) \ 24 > sng天数 Then
                                sng天数 = Val(.TextMatrix(lngCurRow, COL_频率间隔)) \ 24
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_间隔单位) = "分钟" Then
                            If sng天数 = 0 Then sng天数 = 1
                        End If
                        If sng天数 = 0 Then sng天数 = 1
                    End If
                    
                    If Not IsNull(rsItems!总给予量) Then
                        '转换为门诊单位
                        .TextMatrix(lngCurRow, COL_总量) = FormatEx(rsItems!总给予量 / Val(.TextMatrix(lngCurRow, COL_门诊包装)), 5)
                    ElseIf .TextMatrix(lngCurRow, COL_频率) <> "" Then
                        '计算缺省总量
                        rs疗程.Filter = "诊疗项目ID=" & rsItems!诊疗项目ID
                        If Not rs疗程.EOF Then
                            If Nvl(rs疗程!疗程, 1) > sng天数 Then sng天数 = Nvl(rs疗程!疗程, 1)
                        End If
                    
                        If (Val(.TextMatrix(lngCurRow, COL_单量)) <> 0 _
                            And Val(.TextMatrix(lngCurRow, COL_门诊包装)) <> 0 _
                            And Val(.TextMatrix(lngCurRow, COL_剂量系数)) <> 0) Then
                            
                            .TextMatrix(lngCurRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(lngCurRow, COL_单量)), sng天数, _
                                    Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                    Val(.TextMatrix(lngCurRow, COL_频率间隔)), _
                                    .TextMatrix(lngCurRow, COL_间隔单位), _
                                    .TextMatrix(lngCurRow, COL_执行时间), _
                                    Val(.TextMatrix(lngCurRow, COL_剂量系数)), _
                                    Val(.TextMatrix(lngCurRow, COL_门诊包装)), _
                                    Val(.TextMatrix(lngCurRow, COL_可否分零))), 5)
                            If Val(.TextMatrix(lngCurRow, COL_可否分零)) <> 0 Then
                                .TextMatrix(lngCurRow, COL_总量) = IntEx(Val(.TextMatrix(lngCurRow, COL_总量)))
                            End If
                        End If
                    End If
                    
                    If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                        .TextMatrix(lngCurRow, COL_总量) = IntEx(Val(.TextMatrix(lngCurRow, COL_总量)))
                    End If
                    
                    '处方限量
                    If Val(.TextMatrix(lngCurRow, COL_处方限量)) <> 0 Then
                        If Val(.TextMatrix(lngCurRow, COL_总量)) > FormatEx(Val(.TextMatrix(lngCurRow, COL_处方限量)) / Val(.TextMatrix(lngCurRow, COL_剂量系数)) / Val(.TextMatrix(lngCurRow, COL_门诊包装)), 5) Then
                            .TextMatrix(lngCurRow, COL_是否超量) = "1"
                        End If
                    End If
                     
                    If mbln天数 Then
                        .TextMatrix(lngCurRow, COL_天数) = IIF(sng天数 = 0, "", sng天数)
                    End If
                    Call Set用药天数是否超期(lngCurRow)
                    
                ElseIf bln配方 Then
                    If rsItems!类别 = "7" Then
                        .TextMatrix(lngCurRow, COL_总量单位) = "付"
                                                
                        If Not IsNull(rsItems!总给予量) Then
                            .TextMatrix(lngCurRow, COL_总量) = rsItems!总给予量
                        ElseIf .TextMatrix(lngCurRow, COL_频率) <> "" Then
                             .TextMatrix(lngCurRow, COL_总量) = Calc缺省药品总量(1, 1, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                        Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位))
                        End If
                    Else
                        '中药煎法,用法的总量与组成药相同(为了显示)
                        .TextMatrix(lngCurRow, COL_总量) = .TextMatrix(lngCurRow - 1, COL_总量)
                        .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow - 1, COL_总量单位)
                         
                    End If
                Else
                    '其它临嘱都需要总量
                    '如果为一次性或计次临嘱缺省总量为1
                    If Not IsNull(rsItems!总给予量) Then
                        vsAdvice.TextMatrix(lngCurRow, COL_总量) = rsItems!总给予量
                    ElseIf int频率性质 = 1 Or Nvl(rsItems!计算方式, 0) = 3 Then
                        vsAdvice.TextMatrix(lngCurRow, COL_总量) = 1
                    End If
                    If rsItems!类别 = "4" Then
                        .TextMatrix(lngCurRow, COL_总量单位) = .TextMatrix(lngCurRow, COL_门诊单位) '散装单位
                    Else
                        .TextMatrix(lngCurRow, COL_总量单位) = Nvl(rsItems!计算单位)
                    End If
                End If
                
                '抗菌药物缺省用药目的
                If Val(.TextMatrix(lngCurRow, COL_抗菌等级)) > 0 Then .TextMatrix(lngCurRow, COL_用药目的) = mstrPurMed
                        
                '执行时间(总量,频率,执行时间之后)
                If .TextMatrix(lngCurRow, COL_频率) <> "" Then
                    '可能求出缺省执行时间方案
                    If bln给药途径 Or bln中药用法 Then
                        If Not IsNull(rsItems!时间方案) Then
                            If ExeTimeValid(rsItems!时间方案, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位)) Then
                                .TextMatrix(lngCurRow, COL_执行时间) = rsItems!时间方案
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_执行时间) = "" Then
                            .TextMatrix(lngCurRow, COL_执行时间) = Get缺省时间(int适用范围, .TextMatrix(lngCurRow, COL_频率), rsItems!诊疗项目ID)
                        End If
                    ElseIf int频率性质 = 0 Then
                        If Not IsNull(rsItems!时间方案) Then
                            If ExeTimeValid(rsItems!时间方案, Val(.TextMatrix(lngCurRow, COL_频率次数)), _
                                Val(.TextMatrix(lngCurRow, COL_频率间隔)), .TextMatrix(lngCurRow, COL_间隔单位)) Then
                                .TextMatrix(lngCurRow, COL_执行时间) = rsItems!时间方案
                            End If
                        End If
                        If .TextMatrix(lngCurRow, COL_执行时间) = "" Then
                            .TextMatrix(lngCurRow, COL_执行时间) = Get缺省时间(int适用范围, .TextMatrix(lngCurRow, COL_频率))
                        End If
                    End If
                    If bln采集方法 Then
                        .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
                    ElseIf bln给药途径 Or bln中药用法 Then
                        '成药和中药配方的用法,执行时间
                        If bln中药用法 Then
                            .TextMatrix(lngCurRow, COL_用法) = rsItems!名称
                        End If
                        For j = lngCurRow - 1 To lngRow Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = .RowData(lngCurRow) Then
                                If bln给药途径 Then .TextMatrix(j, COL_用法) = rsItems!名称 & rsItems!医生嘱托
                                .TextMatrix(j, COL_执行时间) = .TextMatrix(lngCurRow, COL_执行时间)
                            Else
                                Exit For
                            End If
                        Next
                    ElseIf bln输血途径 Then
                        .TextMatrix(lngCurRow - 1, COL_用法) = rsItems!名称
                    End If
                End If
                
                '开嘱医生和开嘱科室
                .TextMatrix(lngCurRow, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(lngCurRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
                                    
                '执行性质
                If InStr(",5,6,7,", rsItems!类别) > 0 Then
                    If Nvl(rsItems!执行性质, 0) = 5 Then
                        .TextMatrix(lngCurRow, COL_执行性质) = 5
                    Else
                        .TextMatrix(lngCurRow, COL_执行性质) = 4
                    End If
                ElseIf rsItems!类别 = "4" Then
                    .TextMatrix(lngCurRow, COL_执行性质) = 4
                ElseIf bln给药途径 Or bln中药煎法 Or bln中药用法 Or bln采集方法 Then
                    .TextMatrix(lngCurRow, COL_执行性质) = Nvl(rsItems!执行性质, 0)
                Else
                    .TextMatrix(lngCurRow, COL_执行性质) = Nvl(rsItems!执行性质, 0)
                End If
                
                '执行科室ID:为0-叮嘱,5-院外执行时取出为0
                If rsItems!类别 = "Z" And InStr(",1,2,", Nvl(rsItems!操作类型, 0)) > 0 Then
                    If Nvl(rsItems!执行科室ID, 0) <> 0 Then
                        .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                    Else
                        '留观或住院医嘱取临床科室(不管执行性质)
                        If Nvl(rsItems!操作类型, 0) = 1 Then
                            '留观:包含门诊或住院临床科室
                            Call Get临床科室(3, , lngTmp, , True, False, True)
                        ElseIf Nvl(rsItems!操作类型, 0) = 2 Then
                            '住院:包含住院临床科室
                            Call Get临床科室(2, , lngTmp, , True, False, True)
                        End If
                        .TextMatrix(lngCurRow, COL_执行科室ID) = lngTmp
                    End If
                ElseIf InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_执行性质))) = 0 Then
                    If Nvl(rsItems!执行科室ID, 0) <> 0 Then
                        If InStr(",5,6,7,", rsItems!类别) > 0 Then
                            str药房IDs = Get可用药房IDs(rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), mlng病人科室id, 1)
                            If InStr("," & str药房IDs & ",", "," & rsItems!执行科室ID & ",") > 0 Then
                                .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                            End If
                        ElseIf rsItems!类别 = "4" Then
                            str药房IDs = Get可用发料部门IDs(Val(.TextMatrix(lngCurRow, COL_收费细目ID)), mlng病人科室id, 1)
                            If InStr("," & str药房IDs & ",", "," & rsItems!执行科室ID & ",") > 0 Then
                                .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(lngCurRow, COL_执行性质)) = 4 Then
                            '4-指定科室时才取,其它的固定生成
                            .TextMatrix(lngCurRow, COL_执行科室ID) = Nvl(rsItems!执行科室ID, 0)
                            
                            '检查执行科室的有效性
                            If Val(.TextMatrix(lngCurRow, COL_执行科室ID)) <> 0 Then
                                If CheckExecDeptValidate(Val(.TextMatrix(lngCurRow, COL_执行科室ID)), mlng病人科室id, 1, Val(.TextMatrix(lngCurRow, COL_诊疗项目ID))) = False Then
                                    .TextMatrix(lngCurRow, COL_执行科室ID) = 0
                                End If
                            End If
                        End If
                    End If
                    If Val(.TextMatrix(lngCurRow, COL_执行科室ID)) = 0 Then
                        '药品类的整个成套相同
                        If rsItems!类别 = "5" Then
                            If lng西药房ID = 0 Then
                                lng西药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_执行科室ID) = lng西药房ID
                        ElseIf rsItems!类别 = "6" Then
                            If lng成药房ID = 0 Then
                                lng成药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_执行科室ID) = lng成药房ID
                        ElseIf rsItems!类别 = "7" Then
                            If lng中药房ID = 0 Then
                                lng中药房ID = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!诊疗项目ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, 1, True)
                            End If
                            .TextMatrix(lngCurRow, COL_执行科室ID) = lng中药房ID
                        ElseIf rsItems!类别 = "4" Then
                            If lng发料部门ID = 0 Then
                                lng发料部门ID = Get收费执行科室ID(mlng病人ID, 0, rsItems!类别, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, , 1)
                            End If
                            .TextMatrix(lngCurRow, COL_执行科室ID) = lng发料部门ID
                        Else
                            '之前先求开嘱科室ID
                            .TextMatrix(lngCurRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!诊疗项目ID, 0, _
                                Val(.TextMatrix(lngCurRow, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(lngCurRow, COL_开嘱科室ID)), 1, 1)
                        End If
                    End If
                End If
                
                '医生嘱托
                .TextMatrix(lngCurRow, COL_医生嘱托) = Nvl(rsItems!医生嘱托)
                .Cell(flexcpData, lngCurRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, Val(.TextMatrix(lngCurRow, COL_收费细目ID)), "", 0, "", .TextMatrix(lngCurRow, COL_诊疗项目ID) & "|1")
                
                
                '开嘱时间
                .TextMatrix(lngCurRow, COL_开嘱时间) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngCurRow, COL_开嘱时间) = Format(vCurDate, "yyyy-MM-dd HH:mm")
                
                '紧急标志
                .TextMatrix(lngCurRow, COL_标志) = chk紧急.value '可以在界面先统一设置为紧急
                If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(lngCurRow, COL_抗菌等级)) And .TextMatrix(lngCurRow, COL_标志) <> "1" Then

                    .TextMatrix(lngCurRow, COL_审核状态) = 1
                End If
                
                If .TextMatrix(lngCurRow, COL_类别) = "E" And .TextMatrix(lngCurRow, COL_操作类型) = "8" And Val(.TextMatrix(lngCurRow, COL_相关ID)) <> 0 Then
                    strSQL = ""
                    strSQL = GetBloodState(IIF(.TextMatrix(lngCurRow, COL_标志) = "1", 1, 0), Val(.TextMatrix(lngCurRow, COL_执行分类)))
                    .TextMatrix(lngCurRow - 1, COL_审核状态) = strSQL
                    .TextMatrix(lngCurRow, COL_审核状态) = strSQL
                End If
                Call SetRow标志图标(lngCurRow, 1)
                
                '读取药品库存
                If InStr(",5,6,7,", .TextMatrix(lngCurRow, COL_类别)) > 0 Or .TextMatrix(lngCurRow, COL_类别) = "4" And Val(.TextMatrix(lngCurRow, COL_跟踪在用)) = 1 Then
                    If Val(.TextMatrix(lngCurRow, COL_收费细目ID)) <> 0 And Val(.TextMatrix(lngCurRow, COL_执行科室ID)) <> 0 Then
                        .TextMatrix(lngCurRow, COL_库存) = GetStock(Val(.TextMatrix(lngCurRow, COL_收费细目ID)), Val(.TextMatrix(lngCurRow, COL_执行科室ID)), 1)
                    End If
                End If
                
                '----------------------
                '毒麻精药品标识:中药配方及组成味中药不处理
                If InStr(",5,6,", .TextMatrix(lngCurRow, COL_类别)) > 0 And .TextMatrix(lngCurRow, COL_毒理分类) <> "" Then
                    If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(lngCurRow, COL_毒理分类)) > 0 Then
                        .Cell(flexcpFontBold, lngCurRow, col_医嘱内容) = True
                    End If
                End If
                
                '隐蔽一些附加行
                If (InStr(",F,G,D,7,E,C,", rsItems!类别) > 0 And Not IsNull(rsItems!相关序号)) Or bln给药途径 Then
                    .RowHidden(lngCurRow) = True
                End If
                
                '医嘱内容
                If Not .RowHidden(lngCurRow) Then
                    If IsNull(rsItems!诊疗项目ID) Then
                        .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!医嘱内容 '自由录入医嘱
                    ElseIf InStr(",F,D,", rsItems!类别) > 0 And IsNull(rsItems!相关序号) Then
                        .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!名称 '临时
                    Else
                        .TextMatrix(lngCurRow, col_医嘱内容) = AdviceTextMake(lngCurRow)
                    End If
                Else
                    .TextMatrix(lngCurRow, col_医嘱内容) = rsItems!名称
                End If
                
                '产生新嘱时，在可见行读取诊疗项目附项串，用于检查，不用于保存
                If Not .RowHidden(lngCurRow) Then
                    If Not bln中药用法 Then
                        If bln采集方法 Then
                            j = .FindRow(CStr(.RowData(lngCurRow)), , COL_相关ID)
                            If j <> -1 Then
                                .TextMatrix(lngCurRow, COL_附项) = Get医嘱项目附件(Val(.TextMatrix(j, COL_诊疗项目ID)), 1)
                            End If
                        Else
                            .TextMatrix(lngCurRow, COL_附项) = Get医嘱项目附件(Val(.TextMatrix(lngCurRow, COL_诊疗项目ID)), 1)
                        End If
                        If .TextMatrix(lngCurRow, COL_附项) <> "" Then
                            str附项 = str附项 & vbCrLf & "●" & .TextMatrix(lngCurRow, col_医嘱内容)
                        End If
                    End If
                End If
                
                If lngPreRow = -1 And Not .RowHidden(lngCurRow) Then lngPreRow = lngCurRow
                            
                '----------------------
                intCount = intCount + 1
            End If
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        '其它附加处理
        For i = lngRow To lngCurRow
            '取检查和手术的医嘱内容
            If InStr(",F,D,", .TextMatrix(i, COL_类别)) > 0 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                .TextMatrix(i, col_医嘱内容) = AdviceTextMake(i)
            End If
            
            '计算诊疗单价
            If Not .RowHidden(i) And .TextMatrix(i, COL_单价) = "" Then
                .TextMatrix(i, COL_单价) = GetItemPrice(i)
            End If
        Next
        
        '调整受影响行的序号
        Call AdviceSet医嘱序号(lngCurRow + 1, intCount)
        '产生假医嘱ID
        For i = lngRow To lngCurRow
            dbl相关ID = .RowData(i)
            .RowData(i) = GetNext医嘱ID
            For j = i - 1 To lngRow Step -1
                If Val(.TextMatrix(j, COL_相关ID)) = dbl相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            For j = i + 1 To lngCurRow
                If Val(.TextMatrix(j, COL_相关ID)) = dbl相关ID Then
                    .TextMatrix(j, COL_相关ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            
            '新增后关联诊断的标记处理
            If Val(.TextMatrix(i, COL_相关ID)) = 0 Then '以组行处理后为准
                Call SetDiagFlag(i, 1)
            End If
        Next
        If gblnOut必用 Then Call MakeAppNo(2, lngRow, lngCurRow)
        Call Set医嘱超量(lngRow, lngCurRow)
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then '寻找可见行(如配方和检验之后)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        '定位到成套方案的第一行，是因为医生调用成套后会检查或修改医嘱，第一行之后的医嘱为成套方案调用的医嘱
        .Row = lngRow: .Col = col_医嘱内容
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    
    If str附项 <> "" Then
        MsgBox "以下医嘱需要填写申请附项，请注意填写：" & vbCrLf & str附项, vbInformation, gstrSysName
    End If
    If str高危药品 <> "" Then
        MsgBox "以下医嘱是高危药品：" & str高危药品, vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPreRow中药房(lngRow As Long) As Long
'功能：获取上一或下一中药行的执行科室
    Dim lngDrugRow As Long, lngCopyRow As Long
    
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn配方行(lngCopyRow) Then
            '如果上一有效行是中药配方的,则取它的第一中药行
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_相关ID)
        End If
    End If
    
    If lngDrugRow <> -1 Then '缺省与上一配方行相同
        GetPreRow中药房 = Val("" & vsAdvice.TextMatrix(lngDrugRow, COL_执行科室ID))
    End If
End Function

Private Function AdviceSet中药配方(lng诊疗项目ID As Long, ByVal lngRow As Long, ByVal lng用法ID As Long, _
    ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal str摘要 As String, Optional ByVal lng配方ID As Long) As Long
'功能：(重新)处理中药配方的缺省医嘱数据
'参数：lng诊疗项目ID=输入的中药配方ID或单味中药ID
'      lngRow=当前输入行
'      lng用法ID=缺省中药用法ID
'      strExtData=包含配方组成味药及煎法数据:规格ID1,数量,脚注;规格ID2,数量,脚注...|中药煎法|中药形态|付数|药房ID|煎量"
'      rsCurr=如果是修改了配方内容后调用,则包含要保持的一些当前值
'      str摘要=医保摘要
'返回：处理后的中药配方的当前显示行号
    Dim rsItems As New ADODB.Recordset '中药详细信息
    Dim rsUse As New ADODB.Recordset '中药用法信息
    Dim rs煎法 As New ADODB.Recordset '中药煎法项目信息
    Dim rs用法 As New ADODB.Recordset '中药用法项目信息
    Dim arr中药s As Variant, str中药IDs As String, lng相关ID As Long
    Dim lngCopyRow As Long '缺省参照行
    Dim lngDrugRow As Long '如果缺省参照行是中药配方,则为该配方的第一个中药行
    Dim lngFirstRow As Long '当前配方的第一个中药行
    Dim strSQL As String, i As Long
    Dim blnOutOfRange As Boolean
    
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim lng煎法ID As Long, int疗程 As Integer
    Dim str医生 As String, lng医生ID As Long
    Dim lng形态 As Long
    Dim str高危药品 As String, str煎量 As String
        
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn配方行(lngCopyRow) Then
            '如果上一有效行是中药配方的,则取它的第一中药行
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_相关ID)
        End If
    End If
    
    '获取相关数据库信息
    '------------------
    arr中药s = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr中药s)
        str中药IDs = str中药IDs & "," & CStr(Split(arr中药s(i), ",")(0))
    Next
    str中药IDs = Mid(str中药IDs, 2)
    lng煎法ID = Val(Split(strExtData, "|")(1))
    lng形态 = Val(Split(strExtData, "|")(2))
    str煎量 = Split(strExtData, "|")(5)

    
    '配方用法信息:直接输入配方时才有可能有,输入单味中药无
    strSQL = "Select A.用法ID,A.频次,A.疗程,A.医生嘱托" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And B.服务对象 IN(1,3)" & _
        " And Nvl(A.性质,0)=0 And A.项目ID=[1]" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
        " And (b.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or b.撤档时间 is NULL)"
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng诊疗项目ID)
    If Not rsUse.EOF Then lng用法ID = rsUse!用法ID '缺省设置的中药配方用法优先
    
    '配方组成味中药信息:中药无规格概念,对应的的规格记录一定有且只有一条
    strSQL = "Select /*+ rule*/A.计算规则,A.站点,A.类别,A.分类ID,A.ID,A.编码,A.名称,A.标本部位,A.计算单位,A.计算方式,A.执行频率,A.适用性别," & _
        "A.单独应用,A.组合项目,A.操作类型,A.执行安排,A.执行科室,A.服务对象,A.计价性质,A.参考目录ID,A.人员ID,A.建档时间,A.撤档时间," & _
        "A.录入限量,A.试管编码,A.执行分类,A.执行标记,B.药品ID,B.剂量系数,B.门诊包装,B.门诊单位,B.门诊可否分零 As 可否分零,C.处方限量,C.处方职务,c.临床自管药,b.高危药品" & _
        " From 诊疗项目目录 A,药品规格 B,药品特性 C" & _
        " Where A.ID=B.药名ID And A.ID=C.药名ID And B.药品ID IN(Select Column_Value From Table(f_Num2list([1])))"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药IDs)
    
    '配方煎法项目信息
    Set rs煎法 = Get诊疗项目记录(lng煎法ID)
    
    '配方用法项目信息
    Set rs用法 = Get诊疗项目记录(lng用法ID)
    
    '加入配方组成味中药行:按照用户输入顺序
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '中药用法的医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了配方中的内容,用法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的中药配方
        lng相关ID = GetNext医嘱ID
    End If
    
    For i = 0 To UBound(arr中药s)
        rsItems.Filter = "药品ID=" & CStr(Split(arr中药s(i), ",")(0)) '应该肯定有
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = GetNext医嘱ID
        vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到后面的中药用法行
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '新增
        vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
        vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        vsAdvice.TextMatrix(lngRow, COL_类别) = rsItems!类别
        vsAdvice.TextMatrix(lngRow, col_医嘱内容) = rsItems!名称
        vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
        vsAdvice.TextMatrix(lngRow, COL_频率性质) = Nvl(rsItems!执行频率, 0)
        vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rsItems!操作类型)
        vsAdvice.TextMatrix(lngRow, COL_单独应用) = Nvl(rsItems!单独应用)
        vsAdvice.TextMatrix(lngRow, COL_执行分类) = Nvl(rsItems!执行分类, 0)
        
        vsAdvice.TextMatrix(lngRow, COL_单量) = FormatEx(Val(Split(arr中药s(i), ",")(1)), 5) '单味药的单次用量
        vsAdvice.TextMatrix(lngRow, COL_单量单位) = Nvl(rsItems!计算单位)
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = CStr(Split(arr中药s(i), ",")(2)) '单味药的脚注
        vsAdvice.Cell(flexcpData, lngRow, COL_医生嘱托) = str摘要
        
        '规格信息:中药不存在规格概念,一定有
        vsAdvice.TextMatrix(lngRow, COL_收费细目ID) = rsItems!药品ID
        vsAdvice.TextMatrix(lngRow, COL_处方限量) = Nvl(rsItems!处方限量)
        vsAdvice.TextMatrix(lngRow, COL_剂量系数) = rsItems!剂量系数
        vsAdvice.TextMatrix(lngRow, COL_门诊单位) = rsItems!门诊单位
        vsAdvice.TextMatrix(lngRow, COL_门诊包装) = rsItems!门诊包装
        vsAdvice.TextMatrix(lngRow, COL_可否分零) = Nvl(rsItems!可否分零, 0) '对中药实际上无用
        vsAdvice.TextMatrix(lngRow, COL_处方职务) = Nvl(rsItems!处方职务)
        vsAdvice.TextMatrix(lngRow, COL_临床自管药) = rsItems!临床自管药 & ""
        vsAdvice.TextMatrix(lngRow, COL_高危药品) = rsItems!高危药品 & ""
        If Val(vsAdvice.TextMatrix(lngRow, COL_高危药品)) <> 0 Then
            str高危药品 = str高危药品 & vbCrLf & vsAdvice.TextMatrix(lngRow, col_医嘱内容) & ":" & Decode(Val(vsAdvice.TextMatrix(lngRow, COL_高危药品)), 1, "A", 2, "B", 3, "C", "") & "级；"
        End If
        vsAdvice.Cell(flexcpData, lngRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, Val(vsAdvice.TextMatrix(lngRow, COL_收费细目ID)), "", 0, "", vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))

        '计价性质:各自独立
        vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
        
        If lngFirstRow <> 0 Then
            '与上一行已设置的组成中药相同
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = vsAdvice.TextMatrix(lngFirstRow, COL_执行性质)
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_执行科室ID)
            vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
            vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
            vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
            
            vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
        ElseIf Not rsCurr Is Nothing Then
            '修改了配方内容后重新设置,保持与当前的值
            
            '执行性质:修改时根据当前界面设置决定
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "自备药", 5, 4)
            '执行科室
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_执行性质)) = 5, 0, Val("" & rsCurr!执行科室ID))
            
            vsAdvice.TextMatrix(lngRow, COL_频率) = Nvl(rsCurr!频率)
            vsAdvice.TextMatrix(lngRow, COL_频率次数) = Nvl(rsCurr!频率次数)
            vsAdvice.TextMatrix(lngRow, COL_频率间隔) = Nvl(rsCurr!频率间隔)
            vsAdvice.TextMatrix(lngRow, COL_间隔单位) = Nvl(rsCurr!间隔单位)
            vsAdvice.TextMatrix(lngRow, COL_总量) = Nvl(rsCurr!总量)
            vsAdvice.TextMatrix(lngRow, COL_执行时间) = Nvl(rsCurr!执行时间)
            
            vsAdvice.TextMatrix(lngRow, COL_开始时间) = Format(Nvl(rsCurr!开始时间), "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = CStr(Nvl(rsCurr!开始时间))
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsCurr!开嘱医生)
            vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsCurr!开嘱科室id)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(Nvl(rsCurr!开嘱时间), "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = CStr(Nvl(rsCurr!开嘱时间))
            
            vsAdvice.TextMatrix(lngRow, COL_标志) = Nvl(rsCurr!标志)
        Else
            '执行性质:中药配方组成中药相同,缺省=4-指定科室,5-自备药
            vsAdvice.TextMatrix(lngRow, COL_执行性质) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_临床自管药)) = 1, 5, 4)
        
            '执行科室(先在配方界面选择)
            vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = IIF(Val(vsAdvice.TextMatrix(lngRow, COL_执行性质)) = 5, 0, Val(Split(strExtData, "|")(4)))
            
            '执行频率
            '根据用法里面设置的优先
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!频次) Then
                    Call Get频率信息_编码(rsUse!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                    vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                End If
            End If
            '或缺省与上一行相同
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" And lngDrugRow <> -1 Then
                If Val(vsAdvice.TextMatrix(lngDrugRow, COL_EDIT)) = 1 And vsAdvice.TextMatrix(lngDrugRow, COL_频率) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率)
                    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngDrugRow, COL_频率次数)
                    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngDrugRow, COL_频率间隔)
                    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngDrugRow, COL_间隔单位)
                End If
            End If
            '或取缺省值
            If vsAdvice.TextMatrix(lngRow, COL_频率) = "" Then
                Call Get缺省频率(Nvl(rsItems!ID, 0), 2, str频率, int频率次数, int频率间隔, str间隔单位)
                vsAdvice.TextMatrix(lngRow, COL_频率) = str频率
                vsAdvice.TextMatrix(lngRow, COL_频率次数) = int频率次数
                vsAdvice.TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                vsAdvice.TextMatrix(lngRow, COL_间隔单位) = str间隔单位
            End If
            
            '总量(付数):非散装形态已确定付数
            If Val(Split(strExtData, "|")(3)) > 1 Or lng形态 <> 0 Then
                vsAdvice.TextMatrix(lngRow, COL_总量) = Val(Split(strExtData, "|")(3))
            Else
                If vsAdvice.TextMatrix(lngRow, COL_频率) <> "" Then
                    int疗程 = 1
                    If Not rsUse.EOF Then int疗程 = Nvl(rsUse!疗程, 1)
                    '配方付数
                    vsAdvice.TextMatrix(lngRow, COL_总量) = Calc缺省药品总量(1, int疗程, _
                            Val(vsAdvice.TextMatrix(lngRow, COL_频率次数)), _
                            Val(vsAdvice.TextMatrix(lngRow, COL_频率间隔)), _
                            vsAdvice.TextMatrix(lngRow, COL_间隔单位))
                End If
            End If
            
            '执行时间
            If lngDrugRow <> -1 Then '缺省与上一行相同
                If vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngDrugRow, COL_频率) Then
                    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngDrugRow, COL_执行时间)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_执行时间) = "" Then '缺省时间方案
                vsAdvice.TextMatrix(lngRow, COL_执行时间) = Get缺省时间(2, vsAdvice.TextMatrix(lngRow, COL_频率), lng用法ID)
            End If
            
            '开始时间
            If IsDate(txt开始时间.Text) Then
                vsAdvice.TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
            End If
            
            '开嘱医生
            vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
            vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
            
            vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            vsAdvice.TextMatrix(lngRow, COL_标志) = chk紧急.value
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '该中药配方的第一个组成中药行
        
        lngRow = lngRow + 1 '保持当前输入行位置
    Next
    
    '加入中药配方煎法行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = GetNext医嘱ID
    vsAdvice.TextMatrix(lngRow, COL_相关ID) = lng相关ID
    vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '新增
    vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs煎法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng煎法ID
    vsAdvice.TextMatrix(lngRow, COL_标本部位) = str煎量
    vsAdvice.Cell(flexcpData, lngRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
    
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rs煎法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = Nvl(rs煎法!执行频率, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rs煎法!操作类型)
    vsAdvice.TextMatrix(lngRow, COL_单独应用) = Nvl(rs煎法!单独应用)
    vsAdvice.TextMatrix(lngRow, COL_执行分类) = Nvl(rs煎法!执行分类, 0)
    
    '!中药煎法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    
    vsAdvice.TextMatrix(lngRow, col_医嘱内容) = rs煎法!名称
    
    vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
    
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Nvl(rs煎法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "离院带药", 5, Nvl(rs煎法!执行科室, 0))
    End If
    
    '中药煎法如果未设置执行科室,则缺省为病人所在病区(门诊要改为病人所在科室!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rs煎法!类别, lng煎法ID, 0, _
            Nvl(rs煎法!执行科室, 0), mlng病人科室id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rs煎法!计价性质, 0)
    vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
    
    vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
    
    vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
    
    '保持当前输入行位置
    lngRow = lngRow + 1
    
    '设置中药配方用法行:中药配方的显示行
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng相关ID
    If Get诊疗项目记录(lng诊疗项目ID)!类别 & "" = "8" Then
        vsAdvice.TextMatrix(lngRow, COL_配方ID) = lng诊疗项目ID
    End If
    If lng配方ID <> 0 Then
        vsAdvice.TextMatrix(lngRow, COL_配方ID) = lng配方ID
    End If
    
    If Not rsCurr Is Nothing Then
        '修改了配方内容,标记为修改
        If InStr(",0,3,", rsCurr!Edit) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
        Else
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '本来就是新增或修改
        End If
    Else
        '新输入的中药配方,为新增
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_状态) = 1 '新开
    vsAdvice.TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
    Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
    vsAdvice.TextMatrix(lngRow, COL_类别) = rs用法!类别
    vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID) = lng用法ID
    vsAdvice.Cell(flexcpData, lngRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID))
    
    vsAdvice.TextMatrix(lngRow, COL_计算方式) = Nvl(rs用法!计算方式, 0)
    vsAdvice.TextMatrix(lngRow, COL_频率性质) = Nvl(rs用法!执行频率, 0)
    vsAdvice.TextMatrix(lngRow, COL_操作类型) = Nvl(rs用法!操作类型)
    vsAdvice.TextMatrix(lngRow, COL_单独应用) = Nvl(rs用法!单独应用)
    vsAdvice.TextMatrix(lngRow, COL_执行分类) = Nvl(rs用法!执行分类, 0)
    
    '!中药用法中也存放中药的付数
    vsAdvice.TextMatrix(lngRow, COL_总量) = vsAdvice.TextMatrix(lngFirstRow, COL_总量)
    vsAdvice.TextMatrix(lngRow, COL_总量单位) = "付"
    
    vsAdvice.TextMatrix(lngRow, COL_开始时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开始时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开始时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开始时间)
    
    vsAdvice.TextMatrix(lngRow, COL_名称) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_用法) = rs用法!名称
    vsAdvice.TextMatrix(lngRow, COL_频率) = vsAdvice.TextMatrix(lngFirstRow, COL_频率)
    vsAdvice.TextMatrix(lngRow, COL_频率次数) = vsAdvice.TextMatrix(lngFirstRow, COL_频率次数)
    vsAdvice.TextMatrix(lngRow, COL_频率间隔) = vsAdvice.TextMatrix(lngFirstRow, COL_频率间隔)
    vsAdvice.TextMatrix(lngRow, COL_间隔单位) = vsAdvice.TextMatrix(lngFirstRow, COL_间隔单位)
    vsAdvice.TextMatrix(lngRow, COL_执行时间) = vsAdvice.TextMatrix(lngFirstRow, COL_执行时间)
    
    '执行性质:缺省根据项目设置(不可能为院外执行),修改时根据当前界面设置
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Nvl(rs用法!执行科室, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_执行性质) = Decode(Nvl(rsCurr!执行性质), "离院带药", 5, Nvl(rs用法!执行科室, 0))
    End If
    
    '中药用法如果未设置执行科室,则缺省为病人所在病区(门诊要改为病人所在科室!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_执行性质))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rs用法!类别, lng用法ID, 0, _
            Nvl(rs用法!执行科室, 0), mlng病人科室id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_计价性质) = Nvl(rs用法!计价性质, 0)
    vsAdvice.TextMatrix(lngRow, COL_开嘱科室ID) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱科室ID)
    vsAdvice.TextMatrix(lngRow, COL_开嘱医生) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱医生)
    
    vsAdvice.TextMatrix(lngRow, COL_开嘱时间) = vsAdvice.TextMatrix(lngFirstRow, COL_开嘱时间)
    vsAdvice.Cell(flexcpData, lngRow, COL_开嘱时间) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_开嘱时间)
    
    vsAdvice.TextMatrix(lngRow, COL_标志) = vsAdvice.TextMatrix(lngFirstRow, COL_标志)
    Call SetRow标志图标(lngRow)
        
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsCurr!医生嘱托)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsUse!医生嘱托)
    End If
    '中药形态(用于AdviceTextMake中)
    vsAdvice.TextMatrix(lngRow, COL_中药形态) = lng形态
    '处方限量检查
    Call CheckCHLimited(lngRow, vsAdvice.TextMatrix(lngRow, COL_总量), blnOutOfRange, vsAdvice, COL_相关ID, COL_诊疗项目ID, COL_类别, COL_单量)
    If blnOutOfRange Then vsAdvice.TextMatrix(lngRow, COL_是否超量) = "1"
    '中药配方的中药库存
    Call GetDrugStock(lngRow)
    
    '中药配方医嘱内容
    vsAdvice.TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
    
    If str高危药品 <> "" Then
        MsgBox "以下医嘱是高危药品：" & str高危药品, vbInformation, gstrSysName
    End If
        
    AdviceSet中药配方 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet检验组合(ByVal lngRow As Long, ByVal lng采集方法ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset, Optional ByVal str摘要 As String, Optional ByVal bln申请单 As Boolean) As Long
'功能：处理新增的检验(组合)
'参数：rsItems=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng采集方法ID=缺省的采集方法
'      strExtData=检查:"项目ID1,项目ID2,...;检验标本"如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"
'      rsCurr=修改检验项目时用
'      str摘要=医保摘要
'      bln申请单 申请单保存后调用
'返回：处理之后的当前显示行号
    Dim rsMore As New ADODB.Recordset '采集方法信息
    Dim rsItems As New ADODB.Recordset '检验项目信息
    Dim arrItems As Variant, strItems As String
    Dim strSQL As String, curDate As Date
    Dim str医生 As String, lng医生ID As Long
    Dim str频率 As String, int频率次数 As Integer
    Dim int频率间隔 As Integer, str间隔单位 As String
    Dim lng相关ID As Long, str医嘱内容 As String
    Dim lngCopyRow As Long, lngFirstRow As Long, i As Long
    Dim rsLIS As New ADODB.Recordset
    Dim strTmp As String
    Dim Y As Long
    Dim blnLis As Boolean
    
    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    '当前时间
    curDate = zlDatabase.Currentdate
    
    '检验项目信息
    '----------------------------------------------------------------------------
    '各个检验项目信息:按输入顺序
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        If mblnNewLIS Then
            strTmp = arrItems(i)
            If InStr(strTmp, "|") > 0 Then
                For Y = 0 To UBound(Split(strTmp, "|"))
                    strItems = strItems & "," & Val(Split(strTmp, "|")(Y))
                    If Y > 0 Then
                        strSQL = strSQL & " Union All " & " Select '" & Val(Split(strTmp, "|")(Y)) & "' as 子项,'" & Val(Split(strTmp, "|")(0)) & "' as 父项 From Dual "
                    End If
                Next
            Else
                strItems = strItems & "," & Val(strTmp)
            End If
        Else
            strItems = strItems & "," & Val(arrItems(i))
        End If
    Next
    If strSQL <> "" Then
        Set rsLIS = zlDatabase.OpenSQLRecord(Mid(strSQL, 11), Me.Caption)
        blnLis = True
    End If
    Set rsItems = Get诊疗项目记录(0, Mid(strItems, 2))
    
        If Not bln申请单 Then
    '取某个检验项目的采集方法
    strSQL = "Select /*+ RULE */ A.项目ID,Nvl(A.性质,0) as 序号,A.用法ID" & _
        " From 诊疗用法用量 A,诊疗项目目录 B" & _
        " Where A.用法ID=B.ID And B.服务对象 IN(1,3)" & _
        " And A.项目ID IN(Select Column_Value From Table(f_Num2list([1])))" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
        " And (b.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or b.撤档时间 is NULL)" & _
        " Order by A.项目ID,Nvl(A.性质,0)"
    Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strItems, 2))
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng采集方法ID = 0 Then
            lng采集方法ID = rsMore!用法ID '修改时不变
        End If
    End If
        End If

    Set rsMore = Get诊疗项目记录(lng采集方法ID)
    
    mblnRowChange = False
    
    '设置各行检验项目
    '----------------------------------------------------------------------------
    '采集方法医嘱ID,ID顺序与序号不一定一致
    If Not rsCurr Is Nothing Then
        '修改了检验组合中的内容,采集方法行标记为修改,医嘱ID不变
        lng相关ID = rsCurr!医嘱ID
    Else
        '新输入的
        lng相关ID = GetNext医嘱ID
    End If
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = GetNext医嘱ID
            .TextMatrix(lngRow, COL_相关ID) = lng相关ID '对应到采集方法行
            .TextMatrix(lngRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
            .TextMatrix(lngRow, COL_状态) = 1 '新开
            
            .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
            Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
            
            .TextMatrix(lngRow, COL_类别) = rsItems!类别
            .TextMatrix(lngRow, col_医嘱内容) = rsItems!名称
            .TextMatrix(lngRow, COL_诊疗项目ID) = rsItems!ID
            .TextMatrix(lngRow, COL_计算方式) = Nvl(rsItems!计算方式, 0)
            .TextMatrix(lngRow, COL_频率性质) = Nvl(rsItems!执行频率, 0)
            .TextMatrix(lngRow, COL_操作类型) = Nvl(rsItems!操作类型)
            .TextMatrix(lngRow, COL_单独应用) = Nvl(rsItems!单独应用)
            .TextMatrix(lngRow, COL_执行分类) = Nvl(rsItems!执行分类, 0)
            .TextMatrix(lngRow, COL_处方限量) = Nvl(rsItems!录入限量)
            .TextMatrix(lngRow, COL_计价性质) = Nvl(rsItems!计价性质, 0)
            .TextMatrix(lngRow, COL_执行性质) = Nvl(rsItems!执行科室, 0)
            '检验标本
            .TextMatrix(lngRow, COL_标本部位) = Split(strExtData, ";")(1)
            If mblnNewLIS And rsItems!ID & "" <> "" And blnLis Then
                rsLIS.Filter = "子项=" & rsItems!ID
                If rsLIS.EOF = False Then
                    .TextMatrix(lngRow, COL_组合项目ID) = rsLIS!父项 & ""
                End If
            End If
            
            .Cell(flexcpData, lngRow, COL_医生嘱托) = str摘要
            
            '部份内容一并采集的检验项目相同
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
                
                '一并采集的检验项目应该相同
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngFirstRow, COL_执行科室ID)
                End If
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
                
                .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngFirstRow, COL_开始时间)
                .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngFirstRow, COL_开始时间)
                
                .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngFirstRow, COL_开嘱医生)
                .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngFirstRow, COL_开嘱科室ID)
                
                .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngFirstRow, COL_开嘱时间)
                .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngFirstRow, COL_开嘱时间)
                
                .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngFirstRow, COL_标志)
            ElseIf Not rsCurr Is Nothing Then
                .TextMatrix(lngRow, COL_总量) = Nvl(rsCurr!总量, 1)
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    If Nvl(rsCurr!执行科室ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = rsCurr!执行科室ID
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!ID, 0, _
                            Nvl(rsItems!执行科室, 0), mlng病人科室id, Nvl(rsCurr!开嘱科室id, 0), 1, 1)
                    End If
                End If
                
                '执行频率
                .TextMatrix(lngRow, COL_频率) = Nvl(rsCurr!频率)
                .TextMatrix(lngRow, COL_频率次数) = Nvl(rsCurr!频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = Nvl(rsCurr!频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = Nvl(rsCurr!间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = Nvl(rsCurr!执行时间)
                
                '时间/科室/医生
                .TextMatrix(lngRow, COL_开始时间) = Format(Nvl(rsCurr!开始时间), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = CStr(Nvl(rsCurr!开始时间))
                
                .TextMatrix(lngRow, COL_开嘱时间) = Format(Nvl(rsCurr!开嘱时间), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = CStr(Nvl(rsCurr!开嘱时间))
                
                .TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsCurr!开嘱医生)
                .TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsCurr!开嘱科室id)
                
                .TextMatrix(lngRow, COL_标志) = Nvl(rsCurr!标志)
            Else
                .TextMatrix(lngRow, COL_总量) = 1 '缺省为1(次)
                
                '开嘱医生
                .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
                
                '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                    '之前要求出开嘱科室ID
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsItems!类别, rsItems!ID, 0, _
                        Nvl(rsItems!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
                End If
                
                '执行频率
                Call Get缺省频率(Nvl(rsItems!ID, 0), Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
                .TextMatrix(lngRow, COL_频率) = str频率
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                
                '执行时间:"可选频率"(药品是可选频率,但可能设置为一次性)
                If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Then
                    If lngCopyRow <> -1 Then '与上一行相同
                        If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                            .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_执行时间) = "" Then  '缺省时间方案
                        .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率))
                    End If
                End If
            
                '开始时间
                If IsDate(txt开始时间.Text) Then
                    .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
                End If
                
                '开嘱时间
                .TextMatrix(lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                '紧急标志
                .TextMatrix(lngRow, COL_标志) = chk紧急.value
            End If
            
            str医嘱内容 = str医嘱内容 & "," & rsItems!名称 '医嘱内容
            If lngFirstRow = 0 Then lngFirstRow = lngRow '第一项目行
            lngRow = lngRow + 1 '保持当前输入行位置
            
            rsItems.MoveNext
        Next
        
        '设置标本的采集方法
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng相关ID
        
        If Not rsCurr Is Nothing Then
            '修改了检验组合内容,标记为修改
            If InStr(",0,3,", rsCurr!Edit) > 0 Then
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '标记为被修改
            Else
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '本来就是新增或修改
            End If
        Else
            '新输入的检验组合,为新增
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
        End If
        
        .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngRow, COL_状态) = 1 '新开
        
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1) '调整序号
        
        .TextMatrix(lngRow, COL_类别) = rsMore!类别
        .TextMatrix(lngRow, COL_名称) = rsMore!名称
        .TextMatrix(lngRow, COL_用法) = rsMore!名称
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsMore!ID
        .TextMatrix(lngRow, COL_计算方式) = Nvl(rsMore!计算方式, 0)
        .TextMatrix(lngRow, COL_频率性质) = Nvl(rsMore!执行频率, 0)
        .TextMatrix(lngRow, COL_操作类型) = Nvl(rsMore!操作类型)
        .TextMatrix(lngRow, COL_单独应用) = Nvl(rsMore!单独应用)
        .TextMatrix(lngRow, COL_执行分类) = Nvl(rsMore!执行分类, 0)
        .TextMatrix(lngRow, COL_计价性质) = Nvl(rsMore!计价性质, 0)
        .TextMatrix(lngRow, COL_标本部位) = .TextMatrix(lngFirstRow, COL_标本部位)
        
        '总量为检验项目的,与检验项目相同
        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngFirstRow, COL_总量)
        .TextMatrix(lngRow, COL_总量单位) = Nvl(rsMore!计算单位)
        
        '执行频率
        .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngFirstRow, COL_频率)
        .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngFirstRow, COL_频率次数)
        .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngFirstRow, COL_频率间隔)
        .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngFirstRow, COL_间隔单位)
        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngFirstRow, COL_执行时间)
        .TextMatrix(lngRow, COL_执行性质) = Nvl(rsMore!执行科室, 0)
        
        '执行科室:执行性质为(0-叮嘱,5-院外执行)无执行科室
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
            .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsMore!类别, rsMore!ID, 0, _
                Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngFirstRow, COL_开嘱科室ID)), 1, 1)
        End If
        
        '时间/科室/医生
        .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngFirstRow, COL_开始时间)
        .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngFirstRow, COL_开始时间)
        .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngFirstRow, COL_开嘱时间)
        .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngFirstRow, COL_开嘱时间)
        .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngFirstRow, COL_开嘱科室ID)
        .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngFirstRow, COL_开嘱医生)
        
        '显示紧急标志
        .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngFirstRow, COL_标志)
        Call SetRow标志图标(lngRow)
                
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsCurr!医生嘱托)
        End If
        .Cell(flexcpData, lngRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", .TextMatrix(lngRow, COL_诊疗项目ID))
        
        '医嘱内容:检验1,检验2(标本 采集方法)
        .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet检验组合 = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet诊疗项目(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng给药途径ID As Long, ByVal lngGroupRow As Long, _
        ByVal strExtData As String, ByVal str摘要 As String, Optional ByVal str手术部位 As String, Optional ByVal bln备血 As Boolean = True)
'功能：处理新增(插入)的中、西成药，检查(组合)，手术(组合)，卫材，输血，及其它诊疗项目的缺省医嘱数据
'参数：rsInput=输入或选择返回的记录集
'      lngRow=当前输入行
'      lng给药途径ID=缺省给药途径ID,或一并给药时的给药途径ID
'      lngGroupRow=在一并给药的一组成药中插入新的成药行时,对应一并给药的一行行号
'      strExtData=检查:包含检查部位信息,手术:包含附加手术及麻醉的信息,可能无附加手术
'      str摘要=医保摘要
'      str手术部位=申请附项中的手术部位
'      bln备血 当前的输血医嘱为备血医嘱，仅对类别为K的诊疗项目
    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '诊疗项目详细信息
    Dim strSQL As String, lngCopyRow As Long
    Dim lngTmp As Long, i As Long
    Dim str医生 As String, lng医生ID As Long
    Dim str药房IDs As String, sng天数 As Single
    Dim lng执行科室ID As Long, vCurDate As Date
    
    Dim str频率 As String, int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim bln强制缺省 As Boolean, str默认药房 As String

    On Error GoTo errH
    
    '取上一或下一有效行,某些内容缺省与该行相同
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
                
    With vsAdvice
        '开始设置医嘱缺省内容
        .RowData(lngRow) = GetNext医嘱ID
        .TextMatrix(lngRow, COL_EDIT) = 1 '新增
        .TextMatrix(lngRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngRow, COL_状态) = 1 '新开
        
        '序号:保持连续,当前行占用新序号后,后面的序号向后移
        .TextMatrix(lngRow, COL_序号) = GetCurRow序号(lngRow)
        Call AdviceSet医嘱序号(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_类别) = rsInput!类别ID
        .TextMatrix(lngRow, COL_名称) = rsInput!名称 '该名称可能是别名
        .TextMatrix(lngRow, COL_诊疗项目ID) = rsInput!诊疗项目ID
        .TextMatrix(lngRow, COL_收费细目ID) = Nvl(rsInput!收费细目ID)
        .Cell(flexcpData, lngRow, COL_医生嘱托) = str摘要
        '药品、卫材的规格信息
        If Not IsNull(rsInput!收费细目ID) Then
            If InStr(",5,6,", rsInput!类别ID) > 0 Then
                strSQL = "Select Nvl(C.名称,A.名称) as 名称,b.基本药物," & _
                    " B.剂量系数,B.门诊单位,B.门诊包装,B.门诊可否分零 As 可否分零,b.高危药品" & _
                    " From 收费项目目录 A,药品规格 B,收费项目别名 C" & _
                    " Where A.ID=B.药品ID And A.ID=[1]" & _
                    " And A.ID=C.收费细目ID(+) And C.码类(+)=1 And C.性质(+)=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!收费细目ID), IIF(gbyt药品名称显示 = 0, 1, 3))
                .TextMatrix(lngRow, COL_名称) = rsTmp!名称 '将别名换成正式规格名称
                .TextMatrix(lngRow, COL_剂量系数) = rsTmp!剂量系数
                .TextMatrix(lngRow, COL_门诊单位) = rsTmp!门诊单位
                .TextMatrix(lngRow, COL_门诊包装) = rsTmp!门诊包装
                .TextMatrix(lngRow, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
                .TextMatrix(lngRow, COL_高危药品) = Nvl(rsTmp!高危药品, 0)
                .TextMatrix(lngRow, COL_基本药物) = rsTmp!基本药物 & ""
                If Val(.TextMatrix(lngRow, COL_高危药品)) <> 0 Then
                    MsgBox "当前新开的是" & Decode(Val(.TextMatrix(lngRow, COL_高危药品)), 1, "A", 2, "B", 3, "C", "") & "级高危药品，请谨慎使用。", vbInformation, Me.Caption
                End If
            ElseIf rsInput!类别ID = "4" Then
                strSQL = "Select A.跟踪在用,B.名称,B.计算单位 From 材料特性 A,收费项目目录 B Where A.材料ID=B.ID And A.材料ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!收费细目ID))
                .TextMatrix(lngRow, COL_名称) = rsTmp!名称 '将别名换成正式规格名称
                .TextMatrix(lngRow, COL_剂量系数) = 1
                .TextMatrix(lngRow, COL_门诊包装) = 1
                .TextMatrix(lngRow, COL_门诊单位) = Nvl(rsTmp!计算单位) '散装单位
                .TextMatrix(lngRow, COL_跟踪在用) = Nvl(rsTmp!跟踪在用, 0)
                .TextMatrix(lngRow, COL_检查方法) = rsInput!批次 & ""
            End If
        End If
        
        '药品特性
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            strSQL = "Select 毒理分类,抗生素,药品剂型,处方限量,处方职务,临床自管药,溶媒 From 药品特性 Where 药名ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!诊疗项目ID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(lngRow, COL_抗菌等级) = Val("" & rsTmp!抗生素)
                .TextMatrix(lngRow, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                .TextMatrix(lngRow, COL_处方限量) = Nvl(rsTmp!处方限量)
                .TextMatrix(lngRow, COL_处方职务) = Nvl(rsTmp!处方职务)
                .TextMatrix(lngRow, COL_临床自管药) = rsTmp!临床自管药 & ""
                .TextMatrix(lngRow, COL_是否溶媒) = Val(rsTmp!溶媒 & "")
                
                If gblnKSSStrict And UserInfo.用药级别 < Val("" & rsTmp!抗生素) Then
                    .TextMatrix(lngRow, COL_审核状态) = 1
                End If
            End If
        End If
        
        If rsInput!类别ID & "" <> "K" Then
            If chk紧急.value = 1 Then
                If Val(.TextMatrix(lngRow, COL_审核状态)) = 1 Then .TextMatrix(lngRow, COL_审核状态) = ""
            Else
                If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(lngRow, COL_抗菌等级)) Then .TextMatrix(lngRow, COL_审核状态) = 1
            End If
        End If
        
        '获取更多诊疗项目信息
        '----------------------------------------------------------------------------
        strSQL = "Select A.*" & _
            " From 诊疗用法用量 A,诊疗项目目录 B" & _
            " Where A.用法ID=B.ID And (Nvl(A.性质,0)=0 Or B.服务对象 IN(1,3))" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
            " And A.项目ID=[1]"
        strSQL = "Select A.*,Nvl(B.性质,0) as 性质,B.用法ID," & _
            " B.频次,B.成人剂量,B.小儿剂量,B.医生嘱托,B.疗程" & _
            " From 诊疗项目目录 A,(" & strSQL & ") B" & _
            " Where A.ID=B.项目ID(+) And A.ID=[1]" & _
            " Order by 性质"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!诊疗项目ID))
                
        If IsNull(rsInput!收费细目ID) Then '将别名换成正式诊疗名称
            .TextMatrix(lngRow, COL_名称) = rsMore!名称
        End If
                
        If rsInput!类别ID = "4" Then
            .TextMatrix(lngRow, COL_单量单位) = .TextMatrix(lngRow, COL_门诊单位) '散装单位
        ElseIf InStr(",5,6,", rsInput!类别ID) > 0 Or (Nvl(rsMore!执行频率, 0) = 0 And InStr(",1,2,", Nvl(rsMore!计算方式, 0)) > 0) Then
            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsMore!计算单位) '药品为剂量单位
        End If
        
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '中、西成药临嘱的总量单位就是门诊单位
            .TextMatrix(lngRow, COL_总量单位) = .TextMatrix(lngRow, COL_门诊单位)
        ElseIf rsInput!类别ID = "4" Then
            .TextMatrix(lngRow, COL_总量单位) = .TextMatrix(lngRow, COL_门诊单位) '散装单位
        Else
            '其它临嘱要输入总量(计算单位)
            '如果为一次性或计次临嘱缺省总量为1
            If Nvl(rsMore!执行频率, 0) = 1 Or Nvl(rsMore!计算方式, 0) = 3 Then
                .TextMatrix(lngRow, COL_总量) = 1
            End If
            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsMore!计算单位)
        End If
        
        '抗菌药物缺省用药目的
        If Val(.TextMatrix(lngRow, COL_抗菌等级)) > 0 Then .TextMatrix(lngRow, COL_用药目的) = mstrPurMed
        
        .TextMatrix(lngRow, COL_计算方式) = Nvl(rsMore!计算方式, 0)
        .TextMatrix(lngRow, COL_频率性质) = Nvl(rsMore!执行频率, 0)
        .TextMatrix(lngRow, COL_操作类型) = Nvl(rsMore!操作类型)
        .TextMatrix(lngRow, COL_单独应用) = Nvl(rsMore!单独应用)
        .TextMatrix(lngRow, COL_执行分类) = Nvl(rsMore!执行分类, 0)
        If InStr(",5,6,7,", rsInput!类别ID) = 0 Then
            .TextMatrix(lngRow, COL_处方限量) = Nvl(rsMore!录入限量)
        End If
        
        '标本部位
        If InStr(",4,5,6,", rsInput!类别ID) > 0 Then
            .TextMatrix(lngRow, COL_标本部位) = rsInput!名称 '记录药品、卫材输入时选择的名称
        ElseIf rsInput!类别ID = "F" Or rsInput!类别ID = "K" Then
            .TextMatrix(lngRow, COL_手术时间) = "" '记录手术/输血时间
        ElseIf rsInput!类别ID <> "D" Then
            .TextMatrix(lngRow, COL_标本部位) = Nvl(rsMore!标本部位)
        End If
        
        '计价性质
        .TextMatrix(lngRow, COL_计价性质) = Nvl(rsMore!计价性质, 0)
    
        '执行性质:新增项目时根据项目设置,药品、卫材=4-指定科室,一并给药的相同
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_执行性质) = .TextMatrix(lngGroupRow, COL_执行性质)
            Else
                .TextMatrix(lngRow, COL_执行性质) = IIF(Val(.TextMatrix(lngRow, COL_临床自管药)) = 1, 5, 4) '自备药设为院外执行
            End If
        ElseIf rsInput!类别ID = "4" Then
            .TextMatrix(lngRow, COL_执行性质) = 4
        Else
            .TextMatrix(lngRow, COL_执行性质) = Nvl(rsMore!执行科室, 0)
        End If
            
        '开嘱医生和科室
        If lngGroupRow = 0 Then
            .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
            .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
        Else
            .TextMatrix(lngRow, COL_开嘱医生) = .TextMatrix(lngGroupRow, COL_开嘱医生)
            .TextMatrix(lngRow, COL_开嘱科室ID) = .TextMatrix(lngGroupRow, COL_开嘱科室ID)
        End If
    
        '执行科室:药品缺省与上一行相同,一并给药的相同
        lng执行科室ID = 0
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                .TextMatrix(lngRow, COL_执行科室ID) = 0
            Else

                str药房IDs = Get可用药房IDs(rsInput!类别ID, rsInput!诊疗项目ID, Nvl(rsInput!收费细目ID, 0), mlng病人科室id, 1)
                If lngGroupRow <> 0 Then
                    If InStr("," & str药房IDs & ",", "," & .TextMatrix(lngGroupRow, COL_执行科室ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = .TextMatrix(lngGroupRow, COL_执行科室ID)
                    End If
                ElseIf lngCopyRow <> -1 Then
                    bln强制缺省 = Val(zlDatabase.GetPara("门诊医嘱下达强制缺省药房", glngSys, p门诊医嘱下达, 1)) = 1
                    str默认药房 = zlDatabase.GetPara("门诊缺省" & IIF(Val(rsInput!类别ID) = 5, "西", "成") & "药房", glngSys, p门诊医嘱下达, mlng病人科室id)
                    If bln强制缺省 And InStr(str药房IDs, str默认药房) > 0 And str默认药房 <> "" Then
                        lng执行科室ID = 0
                    Else
                        If rsInput!类别ID = .TextMatrix(lngCopyRow, COL_类别) Then
                            lng执行科室ID = Val(.TextMatrix(lngCopyRow, COL_执行科室ID))
                        End If
                    End If
                End If
            End If
        End If

        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
            If rsInput!类别ID = "Z" And InStr(",1,2,", Nvl(rsMore!操作类型, 0)) > 0 Then
                '留观或住院医嘱不缺省
                If Nvl(rsMore!操作类型, 0) = 1 Then
                    '留观:包含门诊或住院的临床科室
                    Call Get临床科室(3, , lngTmp, , True, False, True)
                ElseIf Nvl(rsMore!操作类型, 0) = 2 Then
                    '住院:住院临床科室
                    Call Get临床科室(2, , lngTmp, , True, False, True)
                End If
                .TextMatrix(lngRow, COL_执行科室ID) = lngTmp
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                '执行性质为(0-叮嘱,5-院外执行)无执行科室
                '之前先求出开嘱科室ID
                If rsInput!类别ID = "4" Then
                    .TextMatrix(lngRow, COL_执行科室ID) = Get收费执行科室ID(mlng病人ID, 0, _
                        rsInput!类别ID, Nvl(rsInput!收费细目ID, 0), 4, mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, , 1)
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsInput!类别ID, rsInput!诊疗项目ID, _
                        Nvl(rsInput!收费细目ID, 0), Nvl(rsMore!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1, InStr(",5,6,", rsInput!类别ID) > 0, lng执行科室ID)
                End If
            End If
        End If
        
        '药品库存
        If (InStr(",5,6,", rsInput!类别ID) > 0 Or rsInput!类别ID = "4" And Val(.TextMatrix(lngRow, COL_跟踪在用)) = 1) And Nvl(rsInput!收费细目ID, 0) <> 0 Then
            Call GetDrugStock(lngRow)
        End If
        
        '执行频率:可选频率,一次性或持续性
        
        '缺省与上一新增行相同
        If lngCopyRow <> -1 Then
            If Get频率范围(lngRow) = Get频率范围(lngCopyRow) Then
                If Val(.TextMatrix(lngCopyRow, COL_EDIT)) = 1 And .TextMatrix(lngCopyRow, COL_频率) <> "" _
                    And Not (.TextMatrix(lngRow, COL_类别) = "7" And Not RowIn配方行(lngCopyRow)) _
                    And Not (.TextMatrix(lngRow, COL_类别) <> "7" And RowIn配方行(lngCopyRow)) _
                    And Check频率可用(Nvl(rsInput!诊疗项目ID, 0), Get频率范围(lngRow), .TextMatrix(lngCopyRow, COL_频率)) Then
                    .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率)
                    .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngCopyRow, COL_频率次数)
                    .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngCopyRow, COL_频率间隔)
                    .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngCopyRow, COL_间隔单位)
                End If
            End If
        End If
        '或取缺省频率
        If .TextMatrix(lngRow, COL_频率) = "" Then
            Call Get缺省频率(Nvl(rsInput!诊疗项目ID, 0), Get频率范围(lngRow), str频率, int频率次数, int频率间隔, str间隔单位)
            .TextMatrix(lngRow, COL_频率) = str频率
            .TextMatrix(lngRow, COL_频率次数) = int频率次数
            .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
            .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
        End If
        
        
        '中，西成药的一些缺省信息
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '执行频率
            If lngGroupRow <> 0 Then
                '一并给药的相同
                .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngGroupRow, COL_频率)
                .TextMatrix(lngRow, COL_频率次数) = .TextMatrix(lngGroupRow, COL_频率次数)
                .TextMatrix(lngRow, COL_频率间隔) = .TextMatrix(lngGroupRow, COL_频率间隔)
                .TextMatrix(lngRow, COL_间隔单位) = .TextMatrix(lngGroupRow, COL_间隔单位)
                .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngGroupRow, COL_执行时间)
                
                If Val(.TextMatrix(lngRow, COL_抗菌等级)) > 0 Then
                    .TextMatrix(lngRow, COL_用药目的) = .TextMatrix(lngGroupRow, COL_用药目的)
                    .TextMatrix(lngRow, COL_用药理由) = .TextMatrix(lngGroupRow, COL_用药理由)
                End If
            End If
            
            '确定临嘱用药天数：
            '1.最少为一个频率周期天数
            '2-有疗程则为疗程天数(应大于一个频率周期天数)
            sng天数 = msng天数
            If mbln天数 Then
                If .TextMatrix(lngRow, COL_间隔单位) = "周" Then
                    If 7 > sng天数 Then sng天数 = 7
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "天" Then
                    If Val(.TextMatrix(lngRow, COL_频率间隔)) > sng天数 Then
                        sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔))
                    End If
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "小时" Then
                    If Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24 > sng天数 Then
                        sng天数 = Val(.TextMatrix(lngRow, COL_频率间隔)) \ 24
                    End If
                ElseIf .TextMatrix(lngRow, COL_间隔单位) = "分钟" Then
                    If sng天数 = 0 Then sng天数 = 1
                End If
                If sng天数 = 0 Then sng天数 = 1
            End If
            
            rsMore.Filter = "性质>0" '取第一种给药途径用为缺省设置
            If Not rsMore.EOF Then
                '不是一并给药时,设置的缺省用法频率优先
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!用法ID) Then lng给药途径ID = rsMore!用法ID
                    If Not IsNull(rsMore!频次) Then
                        Call Get频率信息_编码(rsMore!频次, str频率, int频率次数, int频率间隔, str间隔单位)
                        .TextMatrix(lngRow, COL_频率) = str频率
                        .TextMatrix(lngRow, COL_频率次数) = int频率次数
                        .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                        .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                    End If
                End If
                
                '医生嘱托
                .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsMore!医生嘱托) '一般为给药途径的说明
                
                '药品单量
                If mint年龄 > 12 Then
                    If Nvl(rsMore!成人剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!成人剂量, 5)
                    End If
                Else
                    If Nvl(rsMore!小儿剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!小儿剂量, 5)
                    ElseIf Nvl(rsMore!成人剂量, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_单量) = FormatEx(rsMore!成人剂量 * (mint年龄 + 2) * 5 / 100, 5)
                    End If
                End If
                If Val(.TextMatrix(lngRow, COL_单量)) = 0 Then .TextMatrix(lngRow, COL_单量) = ""
                
                '药品临嘱总量:门诊包装
                If Nvl(rsMore!疗程, 1) > sng天数 Then sng天数 = Nvl(rsMore!疗程, 1)
                If .TextMatrix(lngRow, COL_频率) <> "" And Val(.TextMatrix(lngRow, COL_单量)) <> 0 _
                    And Val(.TextMatrix(lngRow, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngRow, COL_门诊包装)) <> 0 Then
                    '仅按疗程算改为按最少用药天数算
                    .TextMatrix(lngRow, COL_总量) = FormatEx(Calc缺省药品总量( _
                            Val(.TextMatrix(lngRow, COL_单量)), sng天数, _
                            Val(.TextMatrix(lngRow, COL_频率次数)), _
                            Val(.TextMatrix(lngRow, COL_频率间隔)), _
                            .TextMatrix(lngRow, COL_间隔单位), _
                            .TextMatrix(lngRow, COL_执行时间), _
                            Val(.TextMatrix(lngRow, COL_剂量系数)), _
                            Val(.TextMatrix(lngRow, COL_门诊包装)), _
                            Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                    If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                        .TextMatrix(lngRow, COL_总量) = IntEx(Val(.TextMatrix(lngRow, COL_总量)))
                    ElseIf Val(.TextMatrix(lngRow, COL_可否分零)) <> 0 Then
                        .TextMatrix(lngRow, COL_总量) = IntEx(Val(.TextMatrix(lngRow, COL_总量)))
                    End If
                End If
            End If
            
            '记录缺省天数
            If mbln天数 Then .TextMatrix(lngRow, COL_天数) = IIF(sng天数 = 0, "", sng天数)
            '当总量，（天数），单量，可能被程序自动（不是手动录入）设值，如设置了用法用量，则不会执行超量检查的代码（控件Validate事件）
            Call Set用药天数是否超期(lngRow)
            If Val(.TextMatrix(lngRow, COL_处方限量)) <> 0 And Val(.TextMatrix(lngRow, COL_总量)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_门诊包装)) * Val(.TextMatrix(lngRow, COL_剂量系数)) > Val(.TextMatrix(lngRow, COL_处方限量)) Then
                    .TextMatrix(lngRow, COL_是否超量) = "1"
                End If
            End If
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        '执行时间:"可选频率"或药品
        If Nvl(rsMore!执行频率, 0) = 0 Or InStr(",5,6,", rsInput!类别ID) > 0 Then
            If .TextMatrix(lngRow, COL_执行时间) = "" Then
                If lngCopyRow <> -1 Then '与上一行相同
                    If .TextMatrix(lngRow, COL_频率) = .TextMatrix(lngCopyRow, COL_频率) Then
                        .TextMatrix(lngRow, COL_执行时间) = .TextMatrix(lngCopyRow, COL_执行时间)
                    End If
                End If
                If .TextMatrix(lngRow, COL_执行时间) = "" Then '缺省时间方案
                    .TextMatrix(lngRow, COL_执行时间) = Get缺省时间(1, .TextMatrix(lngRow, COL_频率), lng给药途径ID)
                End If
            End If
        End If
        
        '其它(与项目无关)
        '---------------------------------------------------------------------
        If lngGroupRow = 0 Then
            vCurDate = zlDatabase.Currentdate
            If IsDate(txt开始时间.Text) Then
                .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
            
                '手术/输血时间缺省为开始时间
                If rsInput!类别ID = "F" Or rsInput!类别ID = "K" Then
                    .TextMatrix(lngRow, COL_手术时间) = txt开始时间.Text
                End If
            End If
            
            .TextMatrix(lngRow, COL_开嘱时间) = Format(vCurDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(vCurDate, "yyyy-MM-dd HH:mm")
    
            .TextMatrix(lngRow, COL_标志) = chk紧急.value
        Else
            .TextMatrix(lngRow, COL_开始时间) = .TextMatrix(lngGroupRow, COL_开始时间)
            .Cell(flexcpData, lngRow, COL_开始时间) = .Cell(flexcpData, lngGroupRow, COL_开始时间)
            
            .TextMatrix(lngRow, COL_开嘱时间) = .TextMatrix(lngGroupRow, COL_开嘱时间)
            .Cell(flexcpData, lngRow, COL_开嘱时间) = .Cell(flexcpData, lngGroupRow, COL_开嘱时间)
            
            .TextMatrix(lngRow, COL_标志) = .TextMatrix(lngGroupRow, COL_标志)
        End If
                        
        
        '在主行处理完成之后处理附加行,并组合医嘱内容
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!类别ID) > 0 Then
            '新增一个给药途径项目,并设置相关
            If lng给药途径ID <> 0 Then
                .TextMatrix(lngRow, COL_用法) = Get项目名称(lng给药途径ID)
            End If
            If lngGroupRow <> 0 Then
                '一并给药的关联相同的给药途径行
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_相关ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_相关ID) = .TextMatrix(lngGroupRow, COL_相关ID)
                Else
                    '这种情况是仅为了使用一并给药的相同设置
                    .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
                End If
            Else '独立新增的成药关联独立的给药途径行
                .TextMatrix(lngRow, COL_相关ID) = AdviceSet给药途径(lngRow, lng给药途径ID)
            End If
            
            '毒麻精的颜色标识
            If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(lngRow, COL_毒理分类)) > 0 _
                And .TextMatrix(lngRow, COL_毒理分类) <> "" Then
                .Cell(flexcpFontBold, lngRow, col_医嘱内容) = True
            End If
        ElseIf rsInput!类别ID = "D" And strExtData <> "" Then
            '检查的组合部位行
            Call AdviceSet检查组合(lngRow, strExtData)
        ElseIf rsInput!类别ID = "F" And strExtData <> "" Then
            '手术的附加手术及麻醉项目行
            Call AdviceSet手术组合(lngRow, strExtData)
            vsAdvice.Cell(flexcpData, lngRow, COL_标本部位) = str手术部位
        ElseIf rsInput!类别ID = "K" Then
            If bln备血 Then
                .TextMatrix(lngRow, COL_检查方法) = ""
            Else
                .TextMatrix(lngRow, COL_检查方法) = 1
            End If
            '输血的途径行
            If lng给药途径ID <> 0 Then
                If gbln血库系统 = True Then
                    strSQL = "Select a.名称,a.操作类型,a.执行分类 From 诊疗项目目录 A where a.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng给药途径ID)
                    .TextMatrix(lngRow, COL_用法) = rsTmp!名称 & ""
                    If Val(rsTmp!操作类型 & "") = 8 And Val(rsTmp!执行分类 & "") = 1 Then '如果是编辑界面用申请单时需要重设一次
                        .TextMatrix(lngRow, COL_检查方法) = 1
                    Else
                        .TextMatrix(lngRow, COL_检查方法) = ""
                    End If
                Else
                    .TextMatrix(lngRow, COL_用法) = Get项目名称(lng给药途径ID)
                End If
                Call AdviceSet输血途径(lngRow, lng给药途径ID)
            End If
        End If
        
        '紧急标志
        If lngGroupRow <> 0 And .TextMatrix(lngRow, COL_审核状态) <> "" Then
            Call SetRow标志图标(lngRow, 0)
        Else
            Call SetRow标志图标(lngRow, 1)
        End If
        
        .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AdviceSet检查组合(ByVal lngRow As Long, ByVal strExData As String)
'功能：重新设置指定检查组合项目的部位方法行,用于新输入检查组合项目或修改部位方法
'参数：lngRow=当前输入行
'      strExData=包含检查部位方法等信息,格式为:"部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
    Dim arrItems As Variant, arrMethod As Variant
    Dim int执行标记 As Integer, str检查部位 As String
    Dim i As Integer, j As Integer, k As Integer
    
    '删除现有的检查部位方法行
    Call Delete检查手术输血(lngRow)
    
    '重新加入部位方法行
    If strExData <> "" Then
        '执行标记
        If UBound(Split(strExData, vbTab)) >= 1 Then
            int执行标记 = Val(Split(strExData, vbTab)(1))
        End If
        vsAdvice.TextMatrix(lngRow, COL_执行标记) = int执行标记 '在这里统一更新主行的执行标记
        
        arrItems = Split(Split(strExData, vbTab)(0), "|")
        For i = 0 To UBound(arrItems)
            str检查部位 = Split(arrItems(i), ";")(0)
            arrMethod = Split(Split(arrItems(i), ";")(1), ",")
            For j = 0 To UBound(arrMethod)
                k = k + 1
                With vsAdvice
                    .AddItem "", lngRow + k
                    .RowHidden(lngRow + k) = True
                    
                    .RowData(lngRow + k) = GetNext医嘱ID
                    .TextMatrix(lngRow + k, COL_相关ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + k, COL_EDIT) = 1 '新增
                    
                    .TextMatrix(lngRow + k, COL_婴儿) = cbo婴儿.ListIndex
                    .TextMatrix(lngRow + k, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + k
                    .TextMatrix(lngRow + k, COL_状态) = 1 '新开
                    
                    .TextMatrix(lngRow + k, COL_类别) = .TextMatrix(lngRow, COL_类别)
                    .TextMatrix(lngRow + k, COL_诊疗项目ID) = .TextMatrix(lngRow, COL_诊疗项目ID) '为同一个检查项目
                    
                    .TextMatrix(lngRow + k, COL_计算方式) = .TextMatrix(lngRow, COL_计算方式)
                    .TextMatrix(lngRow + k, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质)
                    .TextMatrix(lngRow + k, COL_操作类型) = .TextMatrix(lngRow, COL_操作类型)
                    .TextMatrix(lngRow + k, COL_单独应用) = .TextMatrix(lngRow, COL_单独应用)
                    .TextMatrix(lngRow + k, COL_执行分类) = .TextMatrix(lngRow, COL_执行分类)
                    .TextMatrix(lngRow + k, COL_处方限量) = .TextMatrix(lngRow, COL_处方限量)
                    
                    .TextMatrix(lngRow + k, col_医嘱内容) = .TextMatrix(lngRow, COL_名称) '记录为检查项目名称
                    .TextMatrix(lngRow + k, COL_标本部位) = str检查部位
                    .TextMatrix(lngRow + k, COL_检查方法) = arrMethod(j)
                    .TextMatrix(lngRow + k, COL_执行标记) = int执行标记
                    
                    .TextMatrix(lngRow + k, COL_计价性质) = .TextMatrix(lngRow, COL_计价性质)
                    
                    .TextMatrix(lngRow + k, COL_单量) = .TextMatrix(lngRow, COL_单量)
                    .TextMatrix(lngRow + k, COL_总量) = .TextMatrix(lngRow, COL_总量)
                    
                    .TextMatrix(lngRow + k, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                    .TextMatrix(lngRow + k, COL_频率) = .TextMatrix(lngRow, COL_频率)
                    .TextMatrix(lngRow + k, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                    .TextMatrix(lngRow + k, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                    .TextMatrix(lngRow + k, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                    
                    .TextMatrix(lngRow + k, COL_执行性质) = .TextMatrix(lngRow, COL_执行性质)
                    .TextMatrix(lngRow + k, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                    
                    .TextMatrix(lngRow + k, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                    .Cell(flexcpData, lngRow + k, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                    
                    .TextMatrix(lngRow + k, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                    .TextMatrix(lngRow + k, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                    
                    .TextMatrix(lngRow + k, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                    .Cell(flexcpData, lngRow + k, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                    
                    .TextMatrix(lngRow + k, COL_标志) = .TextMatrix(lngRow, COL_标志)
                End With
            Next
        Next
                
        '调整后面医嘱的序号
        Call AdviceSet医嘱序号(lngRow + k + 1, k)
    End If
End Sub

Private Sub AdviceSet手术组合(ByVal lngRow As Long, ByVal strDataIDs As String)
'功能：重新设置指定手术项目的附加手术及麻醉项目行,用于新输入手术项目或手术项目的附加手术及麻醉项目
'参数：lngRow=当前输入行
'      strDataIDs=包含附加手术及麻醉项目信息,其中可能没有附加手术和麻醉
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    '删除现有的附加手术及麻醉项目行
    Call Delete检查手术输血(lngRow)
    
    '重新加入附加手术行及麻醉项目行
    strDataIDs = Trim(Replace(strDataIDs, ";", ","))
    If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
    If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    
    If strDataIDs <> "" Then
        Set rsTmp = Get诊疗项目记录(0, strDataIDs)
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '按用户输入项目顺序
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '不可能EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = GetNext医嘱ID
                    .TextMatrix(lngRow + i + 1, COL_相关ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_EDIT) = 1 '新增
                    
                    .TextMatrix(lngRow + i + 1, COL_婴儿) = cbo婴儿.ListIndex
                    .TextMatrix(lngRow + i + 1, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_状态) = 1 '新开
                    
                    .TextMatrix(lngRow + i + 1, COL_类别) = rsTmp!类别
                    .TextMatrix(lngRow + i + 1, COL_诊疗项目ID) = rsTmp!ID
                    .Cell(flexcpData, lngRow + i + 1, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", .TextMatrix(lngRow + i + 1, COL_诊疗项目ID))
                    
                    .TextMatrix(lngRow + i + 1, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                    .TextMatrix(lngRow + i + 1, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                    .TextMatrix(lngRow + i + 1, COL_操作类型) = Nvl(rsTmp!操作类型)
                    .TextMatrix(lngRow + i + 1, COL_单独应用) = Nvl(rsTmp!单独应用)
                    .TextMatrix(lngRow + i + 1, COL_执行分类) = Nvl(rsTmp!执行分类, 0)
                    .TextMatrix(lngRow + i + 1, COL_处方限量) = Nvl(rsTmp!录入限量)
                    
                    .TextMatrix(lngRow + i + 1, COL_手术时间) = .TextMatrix(lngRow, COL_手术时间) '手术/输血时间
                    .TextMatrix(lngRow + i + 1, col_医嘱内容) = rsTmp!名称
                    
                    .TextMatrix(lngRow + i + 1, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
                    
                    .TextMatrix(lngRow + i + 1, COL_单量) = .TextMatrix(lngRow, COL_单量)
                    .TextMatrix(lngRow + i + 1, COL_总量) = .TextMatrix(lngRow, COL_总量)
    
                    .TextMatrix(lngRow + i + 1, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                    .TextMatrix(lngRow + i + 1, COL_频率) = .TextMatrix(lngRow, COL_频率)
                    .TextMatrix(lngRow + i + 1, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                    .TextMatrix(lngRow + i + 1, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                    .TextMatrix(lngRow + i + 1, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                    
                    '执行性质:根据项目自身设置
                    .TextMatrix(lngRow + i + 1, COL_执行性质) = Nvl(rsTmp!执行科室, 0)
                    
                    '叮嘱和院外执行无执行科室,手术麻醉单独执行科室
                    '否则不管其执行科室设置,一个手术组合应该相同
                    If InStr(",0,5,", Nvl(rsTmp!执行科室, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_执行科室ID) = 0
                    Else
                        If rsTmp!类别 = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!ID, 0, _
                                Nvl(rsTmp!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
                        Else
                            .TextMatrix(lngRow + i + 1, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                        End If
                    End If
                    
                    .TextMatrix(lngRow + i + 1, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                    .Cell(flexcpData, lngRow + i + 1, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                    
                    .TextMatrix(lngRow + i + 1, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                    .TextMatrix(lngRow + i + 1, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                    
                    .TextMatrix(lngRow + i + 1, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                    .Cell(flexcpData, lngRow + i + 1, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                    
                    .TextMatrix(lngRow + i + 1, COL_标志) = .TextMatrix(lngRow, COL_标志)
                End With
            Next
                
            '调整序号
            Call AdviceSet医嘱序号(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet给药途径(ByVal lngRow As Long, ByVal lng给药途径ID As Long, _
    Optional ByVal str执行性质 As String, Optional ByVal lng给药执行ID As Long, Optional ByVal str滴速 As String) As Long
'功能：为录入的中，西成药设置对应的给药途径行(新增或修改)
'参数：lngRow=要处理给药途径的药品行
'      lng给药途径ID=给药途径ID
'      str执行性质=修改给药途径时,当前界面设置的执行性质
'      lng给药执行ID=修改给药途径时,当前界面设置的执行科室
'      str滴速=修改给药途径时,当前界面设置的滴速
'返回：被设置的给药途径行的医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get诊疗项目记录(lng给药途径ID)
    If rsTmp.EOF Then lng给药途径ID = 0 '没有数据，先设置以保持关系
        
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then '未设置"相关ID"时
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '修改医嘱的内容时重新设置给药途径内容(不是更换诊疗项目)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
        End If
        
        '无效内容：名称,收费细目ID,剂量系数,门诊单位,门诊包装,标本部位,医生嘱托,单量,总量,用法
        If blnNew Then
            .RowData(lngNewRow) = GetNext医嘱ID
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngNewRow, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + 1
        Else
            '医嘱ID(RowData),序号:保持不变
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '标志为内容修改
                .TextMatrix(lngNewRow, COL_状态) = 1 '修改后变为新开
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngNewRow, COL_状态) = 1 '新开
        
        .TextMatrix(lngNewRow, COL_类别) = "E" '给药途径属于治疗
        .TextMatrix(lngNewRow, COL_诊疗项目ID) = lng给药途径ID
        '如果没有确定给药途径，暂时不设置的内容
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
            .TextMatrix(lngNewRow, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
            .TextMatrix(lngNewRow, COL_操作类型) = Nvl(rsTmp!操作类型)
            .TextMatrix(lngNewRow, COL_单独应用) = Nvl(rsTmp!单独应用)
            .TextMatrix(lngNewRow, COL_执行分类) = Nvl(rsTmp!执行分类, 0)
            .TextMatrix(lngNewRow, col_医嘱内容) = rsTmp!名称
            
            .TextMatrix(lngNewRow, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
            
            '滴速
            If str滴速 <> "" Then
                .TextMatrix(lngNewRow, COL_医生嘱托) = str滴速
            End If
            .Cell(flexcpData, lngNewRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", .TextMatrix(lngNewRow, COL_诊疗项目ID))
            
            '执行性质:缺省根据项目设置,修改时根据当前界面设置
            If str执行性质 = "" Then
                .TextMatrix(lngNewRow, COL_执行性质) = Nvl(rsTmp!执行科室, 0)
            Else
                .TextMatrix(lngNewRow, COL_执行性质) = Decode(str执行性质, "离院带药", 5, Nvl(rsTmp!执行科室, 0))
            End If
            
            '给药途径如果未设置执行科室,则缺省为病人所在病区(门诊要改为病人所在科室!!)
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_执行性质))) = 0 Then
                If lng给药执行ID <> 0 Then
                    .TextMatrix(lngNewRow, COL_执行科室ID) = lng给药执行ID
                Else
                    .TextMatrix(lngNewRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, "E", lng给药途径ID, 0, _
                        Nvl(rsTmp!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
                End If
            Else
                .TextMatrix(lngNewRow, COL_执行科室ID) = 0
            End If
        End If
        
        '给药途径天数与药品相同
        .TextMatrix(lngNewRow, COL_天数) = .TextMatrix(lngRow, COL_天数)
        
        .TextMatrix(lngNewRow, COL_频率) = .TextMatrix(lngRow, COL_频率)
        .TextMatrix(lngNewRow, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
        .TextMatrix(lngNewRow, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
        .TextMatrix(lngNewRow, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
        .TextMatrix(lngNewRow, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
        
        .TextMatrix(lngNewRow, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
        .Cell(flexcpData, lngNewRow, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
        
        .TextMatrix(lngNewRow, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
        .TextMatrix(lngNewRow, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
        
        .TextMatrix(lngNewRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
        .Cell(flexcpData, lngNewRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
        
        .TextMatrix(lngNewRow, COL_标志) = .TextMatrix(lngRow, COL_标志)
        .TextMatrix(lngNewRow, COL_审核状态) = .TextMatrix(lngRow, COL_审核状态)
            
        '往后调整序号
        If blnNew Then Call AdviceSet医嘱序号(lngNewRow + 1, 1)
        
        AdviceSet给药途径 = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet输血途径(ByVal lngRow As Long, ByVal lng输血途径ID As Long, Optional ByVal lng输血执行ID As Long) As Long
'功能：为录入的中，西成药设置对应的给药途径行(新增或修改)
'参数：lngRow=要处理输血途径的输血医嘱行
'      lng输血途径ID=输血途径ID
'      lng输血执行ID=修改输血途径时,当前界面设置的执行科室
'返回：被设置的输血途径行的医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    Set rsTmp = Get诊疗项目记录(lng输血途径ID)
    
    With vsAdvice
        lngNewRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
        If lngNewRow = -1 Then '尚未设置输血途径时
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        End If
        
        '无效内容：名称,收费细目ID,剂量系数,门诊单位,门诊包装,标本部位,医生嘱托,单量,总量,用法
        If blnNew Then
            .RowData(lngNewRow) = GetNext医嘱ID
            .TextMatrix(lngNewRow, COL_相关ID) = .RowData(lngRow)
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '新增
            .TextMatrix(lngNewRow, COL_序号) = Val(.TextMatrix(lngRow, COL_序号)) + 1
        Else
            '医嘱ID(RowData),序号:保持不变
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '标志为内容修改
                .TextMatrix(lngNewRow, COL_状态) = 1 '修改后变为新开
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_婴儿) = cbo婴儿.ListIndex
        .TextMatrix(lngNewRow, COL_状态) = 1 '新开
        
        .TextMatrix(lngNewRow, COL_类别) = "E" '输血途径属于治疗
        .TextMatrix(lngNewRow, COL_诊疗项目ID) = lng输血途径ID
        .Cell(flexcpData, lngNewRow, COL_医生嘱托) = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", .TextMatrix(lngNewRow, COL_诊疗项目ID))
        
        .TextMatrix(lngNewRow, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
        .TextMatrix(lngNewRow, COL_操作类型) = Nvl(rsTmp!操作类型)
        .TextMatrix(lngNewRow, COL_单独应用) = Nvl(rsTmp!单独应用)
        .TextMatrix(lngNewRow, COL_执行分类) = Nvl(rsTmp!执行分类, 0)
        .TextMatrix(lngNewRow, col_医嘱内容) = rsTmp!名称
        
        .TextMatrix(lngNewRow, COL_计价性质) = Nvl(rsTmp!计价性质, 0)
        .TextMatrix(lngNewRow, COL_执行性质) = Nvl(rsTmp!执行科室, 0)
        
        If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_执行性质))) = 0 Then
            If lng输血执行ID <> 0 Then
                .TextMatrix(lngNewRow, COL_执行科室ID) = lng输血执行ID
            Else
                .TextMatrix(lngNewRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, "E", lng输血途径ID, 0, _
                    Nvl(rsTmp!执行科室, 0), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
            End If
        Else
            .TextMatrix(lngNewRow, COL_执行科室ID) = 0
        End If
        
        .TextMatrix(lngNewRow, COL_频率性质) = .TextMatrix(lngRow, COL_频率性质) '以药品的为准
        .TextMatrix(lngNewRow, COL_频率) = .TextMatrix(lngRow, COL_频率)
        .TextMatrix(lngNewRow, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
        .TextMatrix(lngNewRow, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
        .TextMatrix(lngNewRow, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
        .TextMatrix(lngNewRow, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
        .TextMatrix(lngNewRow, COL_审核状态) = .TextMatrix(lngRow, COL_审核状态)
        
        .TextMatrix(lngNewRow, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
        .Cell(flexcpData, lngNewRow, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
        
        .TextMatrix(lngNewRow, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
        .TextMatrix(lngNewRow, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
        
        .TextMatrix(lngNewRow, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
        .Cell(flexcpData, lngNewRow, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
        
        .TextMatrix(lngNewRow, COL_标志) = .TextMatrix(lngRow, COL_标志)
            
        '往后调整序号
        If blnNew Then Call AdviceSet医嘱序号(lngNewRow + 1, 1)
        
        AdviceSet输血途径 = .RowData(lngNewRow)
        
        '跟据输血途径来重设输血医嘱的审核状态
        strTmp = GetBloodState(IIF(Val(.TextMatrix(lngNewRow, COL_标志)) = 1, 1, 0), Val(.TextMatrix(lngNewRow, COL_执行分类)))
        
        .TextMatrix(lngNewRow, COL_审核状态) = strTmp
        .TextMatrix(lngRow, COL_审核状态) = strTmp
        Call SetRow标志图标(lngRow, 2)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'功能：根据当前医嘱卡片中的内容，更新当前医嘱内容
'说明：对于ListIndex=-1而对应医嘱项又有内容的，保持原内容不更新
    Dim lngRow As Long, lngBeginRow As Long, lngEndRow As Long
    Dim int频率次数 As Integer, int频率间隔 As Integer, str间隔单位 As String
    Dim blnCurDo As Boolean, blnOtherDo As Boolean
    Dim lngTmp As Long, strTmp As String, blnTmp As Boolean
    Dim strCurDate As String, lng开嘱科室ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long
    Dim lng执行科室ID As Long, lngBegin As Long, lngEnd As Long
    Dim blnReSet超量说明 As Boolean
    Dim dbl总量 As Double
    
    With vsAdvice
        lngRow = .Row
        
        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub '清除编辑标志
        
        If RowIn配方行(lngRow) Then
            '中药配方
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            For i = lngBeginRow To lngRow
                '修改处理配方的所有行内容(包括煎法和用法)
                If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then
                    .TextMatrix(i, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开始时间) = txt开始时间.Text
                    blnCurDo = True
                End If
                If chk紧急.Visible And chk紧急.Tag <> "" Then
                    .TextMatrix(i, COL_标志) = chk紧急.value
                    If i = lngRow Then '用法行显示紧急标志
                         Call SetRow标志图标(i, 0)
                    End If
                    blnCurDo = True
                End If
                If txt总量.Enabled And txt总量.Tag <> "" Then
                    .TextMatrix(i, COL_总量) = FormatEx(IIF(Val(txt总量.Text) = 0, "", Val(txt总量.Text)), 5)
                    
                    '总量变了，需要根据是否超量设置超量说明的可用性
                    blnReSet超量说明 = True
                    blnCurDo = True
                End If
                If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                    .TextMatrix(i, COL_频率) = txt频率.Text
                    Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, 2) '中医范围
                    .TextMatrix(i, COL_频率次数) = int频率次数
                    .TextMatrix(i, COL_频率间隔) = int频率间隔
                    .TextMatrix(i, COL_间隔单位) = str间隔单位
                    blnCurDo = True
                End If
                If cbo执行时间.Tag <> "" Then
                    .TextMatrix(i, COL_执行时间) = cbo执行时间.Text
                    blnCurDo = True
                End If
                
                If .TextMatrix(i, COL_类别) = "7" Then
                    '更改的是组成中药的执行科室(用法煎法的改不到)
                    If cbo执行科室.ListIndex <> -1 And cbo执行科室.Tag <> "" Then
                        .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                        blnCurDo = True
                    End If
                    
                    '执行性质:配方中所有组成的中药相同
                    If cbo执行性质.Tag <> "" Then
                        .TextMatrix(i, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "自备药", 5, 4)
                        If Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                            .TextMatrix(i, COL_执行科室ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            '恢复缺省执行科室,缺省与前面相同
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_类别) = "7" And Val(.TextMatrix(j, COL_执行科室ID)) <> 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(j, COL_执行科室ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, .TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), Val(.TextMatrix(i, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, 1, True)
                                End If
                            Else
                                .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngBeginRow, COL_执行科室ID)
                            End If
                        End If
                        blnReInRow = True '界面执行科室编辑性变化
                        blnCurDo = True
                    End If
                End If
                
                '修改时自动更新部份内容
                blnTmp = False
                If cbo医生嘱托.Tag <> "" Or cbo执行性质.Tag <> "" _
                    Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                    blnTmp = True
                End If
                If blnCurDo Or blnTmp Then
                    '修改了内容则更新开嘱时间
                    If strCurDate = "" Then
                        strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                    End If
                    .TextMatrix(i, COL_开嘱时间) = Format(strCurDate, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_开嘱时间) = strCurDate
                    
                    '检查开嘱医生
                    If .TextMatrix(i, COL_开嘱医生) <> UserInfo.姓名 Then
                        .TextMatrix(i, COL_开嘱医生) = UserInfo.姓名
                        If lng开嘱科室ID = 0 Then
                            lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
                        End If
                        .TextMatrix(i, COL_开嘱科室ID) = lng开嘱科室ID
                    End If
                End If
                                                    
                If .TextMatrix(i, COL_类别) = "E" And i <> lngRow Then lngTmp = i '煎法行号
                                                    
                '---------------
                If blnCurDo Then '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                        If Not .RowHidden(i) Then Call ReSetColor(i) '用法行才设置
                    End If
                    mblnNoSave = True '标记为未保存
                End If
            Next
            
            '涉及中药用法行的内容:直接更改当前行的内容(煎法行在配方编辑中才能改)
            '-----------------------------------------------------------
            blnCurDo = False
                    
            '医生嘱托:是放在中药用法行(显示行)中的
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If
        
            '中药用法
            If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                .TextMatrix(lngRow, COL_用法) = txt用法.Text
                
                '同时更改计价性质和执行性质
                .TextMatrix(lngRow, COL_计价性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "计价性质"), 0)
                i = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, "E", Val(cmd用法.Tag), 0, _
                        Val(.TextMatrix(lngRow, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
                End If
                
                blnReInRow = True '需要刷新中药用法执行科室
                blnCurDo = True
            End If
            
            '用法和煎法的执行性质
            If cbo执行性质.Tag <> "" Then
                '用法
                i = Nvl(GetItemField("诊疗项目目录", Val(.TextMatrix(lngRow, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, .TextMatrix(lngRow, COL_类别), _
                        Val(.TextMatrix(lngRow, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngRow, COL_执行性质)), _
                        mlng病人科室id, Val(Val(.TextMatrix(lngRow, COL_开嘱科室ID))), 1, 1)
                End If
                
                '煎法
                i = Nvl(GetItemField("诊疗项目目录", Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), "执行科室"), 0)
                .TextMatrix(lngTmp, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "离院带药", 5, i)
                If Val(.TextMatrix(lngTmp, COL_执行性质)) = 5 Then
                    .TextMatrix(lngTmp, COL_执行科室ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, .TextMatrix(lngTmp, COL_类别), _
                        Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), 0, Val(.TextMatrix(lngTmp, COL_执行性质)), _
                        mlng病人科室id, Val(.TextMatrix(lngTmp, COL_开嘱科室ID)), 1, 1)
                End If
                
                If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                    .TextMatrix(lngTmp, COL_EDIT) = 2
                    .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                End If
                mblnNoSave = True '标记为未保存
                
                blnCurDo = True
            End If
            
            '中药用法执行科室:即配方当前显示行的执行科室
            If cbo附加执行.ListIndex <> -1 And cbo附加执行.Tag <> "" Then
                .TextMatrix(lngRow, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                blnCurDo = True
            End If
            
            Call Set用药天数是否超期(lngRow)
            
        Else '其它诊疗项目
            If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_开始时间) = Format(txt开始时间.Text, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开始时间) = txt开始时间.Text
                blnCurDo = True
            End If
            If IsDate(txt安排时间.Text) And txt安排时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_手术时间) = txt安排时间.Text
                blnCurDo = True
            End If
            '免试标记
            If chk免试.Visible And chk免试.Tag <> "" Then
                .TextMatrix(lngRow, COL_免试) = chk免试.value
                blnCurDo = True
            End If
            
            If chk紧急.Visible And chk紧急.Tag <> "" Then
                .TextMatrix(lngRow, COL_标志) = chk紧急.value
                
                '后面处理一并给药中的其他行，待审核的改为无需审核
                If Not (.TextMatrix(lngRow, COL_类别) = "K") Then
                    If chk紧急.value = 1 Then
                        If Val(.TextMatrix(lngRow, COL_审核状态)) = 1 Then .TextMatrix(lngRow, COL_审核状态) = ""
                    Else
                        If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(lngRow, COL_抗菌等级)) Then .TextMatrix(lngRow, COL_审核状态) = 1
                    End If
                End If
                If (gbln输血分级管理 Or gbln血库系统) And .TextMatrix(lngRow, COL_类别) = "K" Then
                    blnReInRow = True
                End If
                '显示紧急标志,一并给药显示在第一行
                Call SetRow标志图标(lngRow, 0)
                                
                blnCurDo = True
            End If
            If txt单量.Enabled And (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" Then
                .TextMatrix(lngRow, COL_单量) = FormatEx(txt单量.Text, 5)
                If Not mbln天数 Then Call Set用药天数是否超期(lngRow): blnReSet超量说明 = True
                               
                blnCurDo = True
            End If
            
            If txt天数.Visible And txt天数.Enabled And txt天数.Tag <> "" Then
                .TextMatrix(lngRow, COL_天数) = txt天数.Text
                
                If Val(txt天数.Text) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                    .TextMatrix(lngRow, COL_是否超期) = "1"
                Else
                    .TextMatrix(lngRow, COL_是否超期) = ""
                End If
                blnReSet超量说明 = True
                
                blnCurDo = True
            End If
            
            If txt总量.Enabled And (IsNumeric(txt总量.Text) Or txt总量.Text = "") And txt总量.Tag <> "" Then
                .TextMatrix(lngRow, COL_总量) = FormatEx(txt总量.Text, 5)
                If Not mbln天数 Then Call Set用药天数是否超期(lngRow)
                
                '总量变化需设置超量说明的可用性
                blnReSet超量说明 = True
                               
                blnCurDo = True
            End If
            
            If txt频率.Enabled And cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                .TextMatrix(lngRow, COL_频率) = txt频率.Text
                Call Get频率信息_名称(txt频率.Text, int频率次数, int频率间隔, str间隔单位, Get频率范围(lngRow))
                .TextMatrix(lngRow, COL_频率次数) = int频率次数
                .TextMatrix(lngRow, COL_频率间隔) = int频率间隔
                .TextMatrix(lngRow, COL_间隔单位) = str间隔单位
                
                If Not mbln天数 Then Call Set用药天数是否超期(lngRow): blnReSet超量说明 = True
                
                blnCurDo = True
            End If
            
            If cbo执行时间.Tag <> "" Then
                .TextMatrix(lngRow, COL_执行时间) = cbo执行时间.Text
                blnCurDo = True
            End If
            If cbo医生嘱托.Tag <> "" Then
                .TextMatrix(lngRow, COL_医生嘱托) = cbo医生嘱托.Text
                blnCurDo = True
            End If
            
            If cbo执行科室.ListIndex <> -1 And cbo执行科室.Tag <> "" Then
                If Not RowIn检验行(lngRow) Then '采集方法的执行科室不同
                    .TextMatrix(lngRow, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                End If
                blnCurDo = True
            End If
            
            
            '用药目的和理由
            If lbl用药目的.Tag <> "" Then
                .TextMatrix(lngRow, COL_用药目的) = cboDruPur.ListIndex
                blnCurDo = True
            End If
            If txt用药理由.Tag <> "" Then
                .TextMatrix(lngRow, COL_用药理由) = Trim(txt用药理由.Text)
                blnCurDo = True
            End If
            
            '滴速：输液药品
            If cbo滴速.Tag <> "" Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                If lngTmp <> -1 Then
                    If cbo滴速.Text <> "" Then
                        .TextMatrix(lngTmp, COL_医生嘱托) = cbo滴速.Text & lbl滴速单位.Caption
                    Else
                        .TextMatrix(lngTmp, COL_医生嘱托) = ""
                    End If
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                    End If
                    'mblnNoSave = True '标记为未保存
                    blnCurDo = True
                End If
                '显示给药途径
                If cbo滴速.Text <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text & cbo滴速.Text & lbl滴速单位.Caption
                Else
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                End If
            End If
            
            '附加执行科室：给药途径,手术麻醉,采集方法，原液皮试项目
            If cbo附加执行.ListIndex <> -1 And cbo附加执行.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" Then
                    lngTmp = lngRow '原液皮试项目
                ElseIf .TextMatrix(lngRow, COL_类别) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_类别) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_类别) = "E" _
                    And .TextMatrix(lngRow - 1, COL_类别) = "C" _
                    And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                ElseIf .TextMatrix(lngRow, COL_类别) = "K" _
                    And .TextMatrix(lngRow + 1, COL_类别) = "E" _
                    And Val(.TextMatrix(lngRow + 1, COL_相关ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow + 1
                End If
                
                '只更新对应行,不影响其它行
                If lngTmp <> -1 Then
                    '原液皮试
                    If Not (.TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5") Then
                        .TextMatrix(lngTmp, COL_执行科室ID) = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                    End If
                    
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                    End If
                    'mblnNoSave = True '标记为未保存
                    blnCurDo = True
                End If
            End If
            
            '原液皮试项目可以不指定附加的药房科室
            If cbo附加执行.Tag <> "" Then
                If .TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" Then
                    lngTmp = lngRow
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_状态) = 1 '修改后变为新开
                    End If
                    'mblnNoSave = True '标记为未保存
                    blnCurDo = True
                End If
            End If
            
            '执行性质,给药途径:为更新开嘱时间(包括给药途径的同步更改),先判断是否改变
            If InStr(",5,6,K,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                If cbo执行性质.Tag <> "" Then blnCurDo = True
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then blnCurDo = True
            End If
                                    
            '修改时自动更新部份内容
            blnTmp = False
            If cbo执行性质.Tag <> "" Or (Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "") Then
                blnReInRow = True '需要刷新给药途径,采集方式的执行科室
                blnTmp = True
            End If
            If blnCurDo Or blnTmp Then
                '修改了内容则更新开嘱时间
                strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                .TextMatrix(lngRow, COL_开嘱时间) = Format(strCurDate, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_开嘱时间) = strCurDate
                
                '检查开嘱医生
                If .TextMatrix(lngRow, COL_开嘱医生) <> UserInfo.姓名 Then
                    .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
                    If lng开嘱科室ID = 0 Then
                        lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
                    End If
                    .TextMatrix(lngRow, COL_开嘱科室ID) = lng开嘱科室ID
                End If
            End If
                                    
            '其它需要同步处理的关联行
            '----------------------------------------------------------------
            If RowIn检验行(lngRow) Then
                '采集方法
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_诊疗项目ID) = Val(cmd用法.Tag)
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    .TextMatrix(lngRow, COL_名称) = txt用法.Text
                    
                    '同时更改计价性质和执行性质
                    .TextMatrix(lngRow, COL_计价性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "计价性质"), 0)
                    .TextMatrix(lngRow, COL_执行性质) = Nvl(GetItemField("诊疗项目目录", Val(cmd用法.Tag), "执行科室"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, "E", Val(cmd用法.Tag), 0, _
                            Val(.TextMatrix(lngRow, COL_执行性质)), mlng病人科室id, Val(.TextMatrix(lngRow, COL_开嘱科室ID)), 1, 1)
                    Else
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    End If

                    blnCurDo = True
                End If
                
                '设置一并采集的各个检验项目
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                                blnOtherDo = True
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                blnOtherDo = True
                            End If
                            If cbo执行科室.Tag <> "" And cbo执行科室.ListIndex <> -1 Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                Else
                                    .TextMatrix(i, COL_执行科室ID) = cbo执行科室.ItemData(cbo执行科室.ListIndex)
                                End If
                                blnOtherDo = True
                            End If
                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                blnOtherDo = True
                            End If
                            If txt开始时间.Tag <> "" Then
                                .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                blnOtherDo = True
                            End If
                            If chk紧急.Tag <> "" Then
                                .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                blnOtherDo = True
                            End If
                                            
                            '开嘱时间
                            If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            
                            '开嘱医生
                            If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                blnOtherDo = True
                            End If
                                            
                            '开嘱科室ID
                            If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            
                            '标记为修改
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                '中、西成药处理给药途径及一并给药的情况
                
                '执行性质
                If cbo执行性质.Tag <> "" Then
                    .TextMatrix(lngRow, COL_执行性质) = Decode(NeedName(cbo执行性质.Text), "自备药", 5, 4)
                    If Val(.TextMatrix(lngRow, COL_执行性质)) = 5 Then
                        .TextMatrix(lngRow, COL_执行科室ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                        '恢复缺省药房,缺省与前面的成药相同
                        strTmp = Get可用药房IDs(.TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), mlng病人科室id, 1)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '西成药和中成药的药房可能不同,所以类别要相同
                            If .TextMatrix(i, COL_类别) = .TextMatrix(lngRow, COL_类别) And Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_执行科室ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_执行科室ID) = Val(.TextMatrix(i, COL_执行科室ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                            .TextMatrix(lngRow, COL_执行科室ID) = Get诊疗执行科室ID(mlng病人ID, 0, .TextMatrix(lngRow, COL_类别), Val(.TextMatrix(lngRow, COL_诊疗项目ID)), Val(.TextMatrix(lngRow, COL_收费细目ID)), 4, mlng病人科室id, 0, 1, 1, True)
                        End If
                    End If
                    cbo执行科室.Tag = "1" '标明执行科室一并给药的要同步变
                    blnReInRow = True '界面执行科室编辑性变化
                End If
                
                '给药途径本身及其它相关数据同步更改
                lng执行科室ID = 0
                If cbo附加执行.ListIndex <> -1 Then
                    lng执行科室ID = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                End If
                strTmp = ""
                If Trim(cbo滴速.Text) <> "" Then
                    strTmp = cbo滴速.Text & lbl滴速单位.Caption
                End If
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text & strTmp
                    
                    If CheckExecDeptValidate(lng执行科室ID, mlng病人科室id, 1, Val(cmd用法.Tag)) = False Then
                        lng执行科室ID = 0
                    End If
                    
                    Call AdviceSet给药途径(lngRow, Val(cmd用法.Tag), NeedName(cbo执行性质.Text), lng执行科室ID, strTmp)
                ElseIf blnCurDo Then 'cbo执行性质.Tag <> "" Then
                    '如果执行性质更改了,需要强行修改对应的给药途径的执行性质和执行科室
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    Call AdviceSet给药途径(lngRow, Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), NeedName(cbo执行性质.Text), lng执行科室ID, strTmp)
                End If
                
                '一并给药:不处理给药途径,前面已单独设置
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
                    If cbo执行科室.Tag <> "" Then
                        For i = lngBeginRow To .Rows - 1
                            If .TextMatrix(i, COL_相关ID) = "" Then
                                lngTmp = i: Exit For
                            End If
                        Next
                    End If
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 _
                            And Val(.TextMatrix(i, COL_婴儿)) = Val(.TextMatrix(lngRow, COL_婴儿)) Then '可能现在中间有空行
                            If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                                If txt开始时间.Tag <> "" Then
                                    .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                    .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                    blnOtherDo = True
                                End If
                                If txt用法.Tag <> "" Then
                                    .TextMatrix(i, COL_用法) = .TextMatrix(lngRow, COL_用法)
                                    blnOtherDo = True
                                End If
                                If txt频率.Tag <> "" Then
                                    .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                    .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                    .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                    .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                    blnOtherDo = True
                                End If
                                
                                '将滴速补填到一并给药的其他行中
                                If cbo滴速.Tag <> "" Then
                                    .TextMatrix(i, COL_用法) = txt用法.Text & strTmp
                                    blnOtherDo = True
                                End If
                                
                                '一并给药的,天数相同变化,总量重新计算
                                If txt天数.Tag <> "" Then
                                    .TextMatrix(i, COL_天数) = .TextMatrix(lngRow, COL_天数)
                                    If .TextMatrix(i, COL_频率) <> "" _
                                        And Val(.TextMatrix(i, COL_单量)) <> 0 _
                                        And Val(.TextMatrix(i, COL_剂量系数)) <> 0 _
                                        And Val(.TextMatrix(i, COL_门诊包装)) <> 0 Then
                                        
                                        .TextMatrix(i, COL_总量) = FormatEx(Calc缺省药品总量( _
                                            Val(.TextMatrix(i, COL_单量)), Val(.TextMatrix(i, COL_天数)), _
                                            Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), _
                                            .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), _
                                            Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_门诊包装)), _
                                            Val(.TextMatrix(i, COL_可否分零))), 5)
                                        If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                                            .TextMatrix(i, COL_总量) = IntEx(Val(.TextMatrix(i, COL_总量)))
                                        ElseIf Val(.TextMatrix(i, COL_可否分零)) <> 0 Then
                                            .TextMatrix(i, COL_总量) = IntEx(Val(.TextMatrix(i, COL_总量)))
                                        End If
                                    End If
                                    '重新检查处方限量超天
                                    .TextMatrix(i, COL_是否超期) = "": .TextMatrix(i, COL_是否超量) = ""
                                    If Val(.TextMatrix(i, COL_处方限量)) <> 0 Then
                                        dbl总量 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_门诊包装)) * Val(.TextMatrix(i, COL_剂量系数))
                                        If dbl总量 > Val(.TextMatrix(i, COL_处方限量)) Then .TextMatrix(i, COL_是否超量) = "1"
                                    End If
                                    
                                    If Val(.TextMatrix(i, COL_天数)) > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
                                        .TextMatrix(i, COL_是否超期) = "1"
                                    End If
                                    
                                    If .TextMatrix(i, COL_是否超期) = "" And .TextMatrix(i, COL_是否超量) = "" Then .TextMatrix(i, COL_超量说明) = ""
                                    
                                    blnOtherDo = True
                                End If
                                
                                If cbo执行时间.Tag <> "" Then
                                    .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                    blnOtherDo = True
                                End If
                                
                                '执行性质:离院带药在一并给药中需一致，其它可单独设置
                                If cbo执行性质.Tag <> "" And NeedName(cbo执行性质.Text) = "离院带药" Then
                                    .TextMatrix(i, COL_执行性质) = .TextMatrix(lngRow, COL_执行性质)
                                    '由自备药转过来时需要重新设置执行科室
                                    If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                                        .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                    End If
                                    blnOtherDo = True
                                End If
                                
                                '执行科室:执行科室(药房)可以不同,除非是配制中心
                                If cbo执行科室.Tag <> "" Then
                                    '输入行改为自备药，或某行为自备药的情况不管它
                                    If Not (Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 And Val(.TextMatrix(lngRow, COL_执行性质)) = 5) _
                                        And Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                                        If .TextMatrix(lngTmp, COL_类别) = "E" And .TextMatrix(lngTmp, COL_操作类型) = "2" And .TextMatrix(lngTmp, COL_执行分类) = "1" Then
                                            If Have部门性质(Val(.TextMatrix(lngRow, COL_执行科室ID)), "配制中心") Then
                                                '输入行药品由普通药房或其他配制中心改为新的配制中心,则该组药都改为该配制中心
                                                .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                                blnOtherDo = True
                                            ElseIf Have部门性质(Val(.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                                                '输入行药品由配制中心改成普通药房,则该组药都改为该普通药房
                                                .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                                blnOtherDo = True
                                            End If
                                        End If
                                    End If
                                End If
                                
                                '紧急标志
                                If chk紧急.Tag <> "" Then
                                    .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                    .TextMatrix(i, COL_审核状态) = .TextMatrix(lngRow, COL_审核状态)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱时间
                                If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱医生
                                If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                    .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                    blnOtherDo = True
                                End If
                                
                                '开嘱科室ID
                                If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                    .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                    blnOtherDo = True
                                End If
                                
                                '标记为修改
                                If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                    .TextMatrix(i, COL_EDIT) = 2
                                    .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_类别) = "K" Then
                '输血医嘱的处理(前面已处理输血时间(安排时间)的修改)
                
                lng执行科室ID = 0
                If cbo附加执行.ListIndex <> -1 Then
                    lng执行科室ID = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                End If
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then
                    .TextMatrix(lngRow, COL_用法) = txt用法.Text
                    Call AdviceSet输血途径(lngRow, Val(cmd用法.Tag), lng执行科室ID)
                ElseIf blnCurDo Then
                    lngTmp = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                    If lngTmp <> -1 Then
                        Call AdviceSet输血途径(lngRow, Val(.TextMatrix(lngTmp, COL_诊疗项目ID)), lng执行科室ID)
                    End If
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_类别)) > 0 And blnCurDo Then
                '检查组合项目行或手术附加行
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_相关ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                            If txt单量.Tag <> "" Then
                                .TextMatrix(i, COL_单量) = .TextMatrix(lngRow, COL_单量)
                                blnOtherDo = True
                            End If
                            If txt总量.Tag <> "" Then
                                .TextMatrix(i, COL_总量) = .TextMatrix(lngRow, COL_总量)
                                blnOtherDo = True
                            End If
                            
                            If cbo执行时间.Tag <> "" Then
                                .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow, COL_执行时间)
                                blnOtherDo = True
                            End If
                            If txt频率.Tag <> "" Then
                                .TextMatrix(i, COL_频率) = .TextMatrix(lngRow, COL_频率)
                                .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow, COL_频率次数)
                                .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow, COL_频率间隔)
                                .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow, COL_间隔单位)
                                blnOtherDo = True
                            End If
                            If cbo执行科室.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) > 0 Then
                                    .TextMatrix(i, COL_执行科室ID) = 0
                                ElseIf .TextMatrix(i, COL_类别) <> "G" Then '手术麻醉的执行科室为单独
                                    .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow, COL_执行科室ID)
                                End If
                                blnOtherDo = True
                            End If
                            If txt开始时间.Tag <> "" Then
                                .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow, COL_开始时间)
                                .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow, COL_开始时间)
                                blnOtherDo = True
                            End If
                            If txt安排时间.Tag <> "" Then
                                .TextMatrix(i, COL_手术时间) = .TextMatrix(lngRow, COL_手术时间)
                                blnOtherDo = True
                            End If
                            If chk紧急.Tag <> "" Then
                                .TextMatrix(i, COL_标志) = .TextMatrix(lngRow, COL_标志)
                                blnOtherDo = True
                            End If
                            
                            '开嘱时间
                            If .TextMatrix(i, COL_开嘱时间) <> .TextMatrix(lngRow, COL_开嘱时间) Then
                                .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                                .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                                blnOtherDo = True
                            End If
                            
                            '开嘱医生
                            If .TextMatrix(i, COL_开嘱医生) <> .TextMatrix(lngRow, COL_开嘱医生) Then
                                .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                                blnOtherDo = True
                            End If
                            
                            '开嘱科室ID
                            If .TextMatrix(i, COL_开嘱科室ID) <> .TextMatrix(lngRow, COL_开嘱科室ID) Then
                                .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                                blnOtherDo = True
                            End If
                            
                            '标记为修改
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf .TextMatrix(lngRow, COL_类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "1" And .TextMatrix(lngRow, COL_执行分类) = "5" Then
                '原液皮试
                If cbo附加执行.Tag <> "" Then
                    lng执行科室ID = 0
                    If cbo附加执行.ListIndex <> -1 Then
                        lng执行科室ID = cbo附加执行.ItemData(cbo附加执行.ListIndex)
                    End If
                    .TextMatrix(lngRow, COL_用药理由) = lng执行科室ID
                End If
            End If
        End If
        
        '配方和其他诊疗项目都可能存在超期
        If txt超量说明.Enabled And txt超量说明.Tag <> "" Then
            .TextMatrix(lngRow, COL_超量说明) = txt超量说明.Text
            blnCurDo = True
        End If
        If lbl超量说明.Tag <> "" Then
            blnReSet超量说明 = True
        End If
        '统一处理相关行的属性变化
        If chkZeroBilling.Tag <> "" Then
            Call GetRowScope(lngRow, lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                .TextMatrix(i, COL_零费记帐) = chkZeroBilling.value
            Next
            blnCurDo = True
        End If
                    
        '---------------
        If blnCurDo Then '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
            If InStr(",0,2,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                 '审核未通过的是一组药品,则需要改变组内其他行的审核状态为未审核或无需审核
                If Val(.TextMatrix(lngRow, COL_审核状态)) <> 2 And .TextMatrix(lngRow, COL_类别) <> "K" Then
                    Call GetRowScope(lngRow, lngBeginRow, lngEndRow)
                    For i = lngBeginRow To lngEndRow
                        '如果是紧急医嘱，则无须审核
                        If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(i, COL_抗菌等级)) And .TextMatrix(i, COL_标志) <> 1 Then
                            .TextMatrix(i, COL_审核状态) = 1
                        Else
                            .TextMatrix(i, COL_审核状态) = ""
                        End If
                        Call SetRow标志图标(i, 2)
                    Next
                End If
                
                .TextMatrix(lngRow, COL_EDIT) = 2
                .TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
                Call ReSetColor(lngRow)
            End If
            mblnNoSave = True '标记为未保存
        End If
        
        '更新医嘱内容
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
            txt医嘱内容.Text = .TextMatrix(lngRow, col_医嘱内容)
        End If
    End With
        
    '清除编辑标志
    Call ClearItemTag
    
    '某些情况下需要重新设置卡片的项目编辑性(如修改了执行性质时)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        If vsAdvice.TextMatrix(lngRow, COL_是否超量) = "" And vsAdvice.TextMatrix(lngRow, COL_是否超期) = "" Then vsAdvice.TextMatrix(lngRow, COL_超量说明) = ""
    ElseIf blnReSet超量说明 Then
        SetItemEditable , , , , , , , , , , , , IIF(vsAdvice.TextMatrix(lngRow, COL_是否超量) = "1" Or vsAdvice.TextMatrix(lngRow, COL_是否超期) = "1", 1, -1)
        If vsAdvice.TextMatrix(lngRow, COL_是否超量) = "" And vsAdvice.TextMatrix(lngRow, COL_是否超期) = "" Then vsAdvice.TextMatrix(lngRow, COL_超量说明) = ""
    End If
End Sub

Private Sub ReSetColor(ByVal lngRow As Long)
'功能：重新设置指定行的颜色
'说明：
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        '一并给药范围
        lngBegin = lngRow: lngEnd = lngRow
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
            If RowIn一并给药(lngRow) Then
                Call Get一并给药范围(Val(.TextMatrix(lngRow, COL_相关ID)), lngBegin, lngEnd)
            End If
        End If
        '恢复成正常色
        For i = lngBegin To lngEnd
            .Cell(flexcpForeColor, i, .FixedCols, i, COL_开嘱医生) = .ForeColor
            '毒麻精的颜色标识
            If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(i, COL_毒理分类)) > 0 _
                And .TextMatrix(i, COL_毒理分类) <> "" Then
                .Cell(flexcpFontBold, i, col_医嘱内容) = True
            End If
        Next
        .ForeColorSel = .Cell(flexcpForeColor, lngRow, COL_开始时间)
    End With
End Sub

Private Sub AdviceSet一并给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：将选择范围内的药品设置为一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
'说明：以第一行药品的给药途径为准,但位置放在最后一行药品之后
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng相关ID As Long, i As Long
    Dim strStart As String, curDate As Date
    Dim lng配制中心 As Long
        
    With vsAdvice
        lngRow1 = .FindRow(CLng(.TextMatrix(lngBegin, COL_相关ID)), lngBegin + 1) '第一给药途径行
        lngRow2 = .FindRow(CLng(.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1) '最后给药途径行
        
        
        '删除给药途径行之前记录执行性质,以便后面作判断
        For i = lngRow2 To lngRow1 Step -1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex And .RowHidden(i) Then
                .Cell(flexcpData, i - 1, COL_执行性质) = Val(.TextMatrix(i, COL_执行性质))
            End If
        Next
        
        '复制第一行的给药途径到最后一行的给药途径
        For i = .FixedCols To .Cols - 1
            If i <> COL_EDIT And i <> COL_相关ID And i <> COL_序号 And i <> COL_状态 Then
                .TextMatrix(lngRow2, i) = .TextMatrix(lngRow1, i)
            End If
        Next
        .Cell(flexcpData, lngRow2, COL_开始时间) = .TextMatrix(lngRow2, COL_开始时间)
        
        '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号
        If InStr(",0,3,", .TextMatrix(lngRow2, COL_EDIT)) > 0 Then
            .TextMatrix(lngRow2, COL_EDIT) = 2 '标记为已修改
            .TextMatrix(lngRow2, COL_状态) = 1 '修改后变为新开
        End If
        lng相关ID = .RowData(lngRow2)
        
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        '删除除最后一行给药途径外的其它给药途径
        For i = lngEnd To lngBegin Step -1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                If .RowHidden(i) Then
                    Call DeleteRow(i)
                Else
                    .TextMatrix(i, COL_相关ID) = lng相关ID
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                        .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                    End If
                End If
            End If
        Next
        
        '行号已变更
        lngRow1 = lngBegin '开始一并给药行
        curDate = zlDatabase.Currentdate
        
        '检查医生是否变更
        If .TextMatrix(lngRow1, COL_开嘱医生) <> UserInfo.姓名 Then
            '更新相关信息:前面已标记为修改,且手工操作完成时已有进入界面刷新
            .TextMatrix(lngRow1, COL_开嘱医生) = UserInfo.姓名
            .TextMatrix(lngRow1, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
            
            .TextMatrix(lngRow1, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow1, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        
        '处理一并给药其他行的相同信息
        For i = lngRow1 + 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                    lngRow2 = i '记录新的结束行号
                    
                    '一并给药的部分信息相同
                    .TextMatrix(i, COL_开始时间) = .TextMatrix(lngRow1, COL_开始时间)
                    .Cell(flexcpData, i, COL_开始时间) = .Cell(flexcpData, lngRow1, COL_开始时间)
                    
                    .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow1, COL_开嘱医生)
                    .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow1, COL_开嘱科室ID)
                    
                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow1, COL_开嘱时间) '一并给药的开嘱时间相同
                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow1, COL_开嘱时间)
                    
                    .TextMatrix(i, COL_天数) = .TextMatrix(lngRow1, COL_天数)
                    
                    .TextMatrix(i, COL_用法) = .TextMatrix(lngRow1, COL_用法)
                    
                    .TextMatrix(i, COL_频率) = .TextMatrix(lngRow1, COL_频率)
                    .TextMatrix(i, COL_频率次数) = .TextMatrix(lngRow1, COL_频率次数)
                    .TextMatrix(i, COL_频率间隔) = .TextMatrix(lngRow1, COL_频率间隔)
                    .TextMatrix(i, COL_间隔单位) = .TextMatrix(lngRow1, COL_间隔单位)
                    .TextMatrix(i, COL_执行时间) = .TextMatrix(lngRow1, COL_执行时间)
                    
                    '启用了天数控制才反算（71152）
                    If mbln天数 Then .TextMatrix(i, COL_总量) = ReGet药品总量(Val(.TextMatrix(i, COL_总量)), Val(.TextMatrix(i, COL_单量)), Val(.TextMatrix(i, COL_天数)), i)
                    
                    
                    '处方限量
                    .TextMatrix(i, COL_是否超量) = ""
                    If Val(.TextMatrix(i, COL_处方限量)) <> 0 Then
                        If Val(.TextMatrix(i, COL_总量)) > FormatEx(Val(.TextMatrix(i, COL_处方限量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_门诊包装)), 5) Then
                            .TextMatrix(i, COL_是否超量) = "1"
                        End If
                    End If
                    
                    .TextMatrix(i, COL_是否超期) = ""
                    Call Set用药天数是否超期(i)
                    
                    .TextMatrix(i, COL_标志) = .TextMatrix(lngRow1, COL_标志)
                    Set .Cell(flexcpPicture, i, COL_F标志) = Nothing '在开始行显示
                    
                    '离院带药一组相同
                    If Val(.TextMatrix(lngRow1, COL_执行性质)) <> 5 And Val(.Cell(flexcpData, lngRow1, COL_执行性质)) = 5 Then
                        '第一行是离院带药,全部设置为离院带药
                        .TextMatrix(i, COL_执行性质) = .TextMatrix(lngRow1, COL_执行性质)
                        If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then '执行科室可以不同,没有时才缺省相同
                            .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow1, COL_执行科室ID)
                        End If
                    ElseIf Val(.TextMatrix(i, COL_执行性质)) <> 5 And Val(.Cell(flexcpData, i, COL_执行性质)) = 5 Then
                        '当前行是离院带药,则设置为与第一行相同
                        .TextMatrix(i, COL_执行性质) = .TextMatrix(lngRow1, COL_执行性质)
                        If Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            .TextMatrix(i, COL_执行科室ID) = .TextMatrix(lngRow1, COL_执行科室ID)
                        End If
                    Else
                        '否则保持不变
                    End If
                    
                    '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                    End If
                Else
                    Exit For
                End If
            End If
        Next
    
        '检查这些药品中是否存在配制中心拿药的，以第一个为准
        For i = lngRow1 To .Rows - 1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                    '自备药的情况不管它
                    If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                        If Have部门性质(Val(.TextMatrix(i, COL_执行科室ID)), "配制中心") Then
                            lng配制中心 = Val(.TextMatrix(i, COL_执行科室ID)): Exit For
                        End If
                    End If
                Else
                    Exit For
                End If
            End If
        Next
        '配制中心一组相同
        If lng配制中心 <> 0 Then
            For i = lngRow1 To .Rows - 1
                If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                    If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                        '自备药的情况不管它
                        If Not (Val(.TextMatrix(i, COL_执行科室ID)) = 0 And Val(.TextMatrix(i, COL_执行性质)) = 5) Then
                            .TextMatrix(i, COL_执行科室ID) = lng配制中心
                            
                            '标记为修改:0-原始的,1-新增的,2-修改了内容,3-修改了序号
                            If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                            End If
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
        
        '开始执行时间处理(新开的不能太早)
        strStart = ""
        For i = lngRow1 To lngRow2
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                If Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                    If DateDiff("n", CDate(.Cell(flexcpData, i, COL_开始时间)), curDate) > gint门诊新开医嘱间隔 Then
                        strStart = GetDefaultTime(i): Exit For
                    End If
                End If
            End If
        Next
        If strStart <> "" Then
            For i = lngRow1 To lngRow2 + 1
                If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                    .Cell(flexcpData, i, COL_开始时间) = strStart
                    .TextMatrix(i, COL_开始时间) = Format(strStart, "yyyy-MM-dd HH:mm")
                End If
            Next
        End If
        
        Call ReSet审核状态图标(lngBegin)
    
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '标记为未保存
    End With
End Sub

Private Sub AdviceSet单独给药(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：取消一组药品的一并给药
'参数：起止行号,中间不包含空行,不包含最后一行药品的给药途径行
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng给药途径ID As Long, lng给药执行ID As Long, i As Long
    Dim int执行性质 As Integer, str执行性质 As String, str滴速 As String
    Dim lngRow As Long, curDate As Date, blnUpdate As Boolean
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        '一并给药途径
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_相关ID)), lngEnd + 1)
        lng给药途径ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
        lng给药执行ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        int执行性质 = Val(.TextMatrix(lngRow, COL_执行性质))
        str滴速 = .TextMatrix(lngRow, COL_医生嘱托)
                
        '检查医生变更:以给药途径行为准变化
        If .TextMatrix(lngRow, COL_开嘱医生) <> UserInfo.姓名 Then
            '更新相关信息:手工操作完成时有进入界面刷新
            .TextMatrix(lngRow, COL_开嘱医生) = UserInfo.姓名
            .TextMatrix(lngRow, COL_开嘱科室ID) = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 1)
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_开嘱时间) = Format(curDate, "yyyy-MM-dd HH:mm")
            
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2 '标记为已修改
                .TextMatrix(lngRow, COL_状态) = 1 '修改后变为新开
            End If
            blnUpdate = True
        End If
                
        '显示紧急标志:每一行
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                '药品行相应变化
                If blnUpdate Then
                    .TextMatrix(i, COL_开嘱医生) = .TextMatrix(lngRow, COL_开嘱医生)
                    .TextMatrix(i, COL_开嘱科室ID) = .TextMatrix(lngRow, COL_开嘱科室ID)
                    .TextMatrix(i, COL_开嘱时间) = .TextMatrix(lngRow, COL_开嘱时间)
                    .Cell(flexcpData, i, COL_开嘱时间) = .Cell(flexcpData, lngRow, COL_开嘱时间)
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                        .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                    End If
                End If
                If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(i, COL_抗菌等级)) And .TextMatrix(i, COL_标志) <> "1" Then
                    .TextMatrix(i, COL_审核状态) = 1
                Else
                    .TextMatrix(i, COL_审核状态) = ""
                End If
                Call SetRow标志图标(i, 2)
            End If
        Next
        
        For i = lngEnd - 1 To lngBegin Step -1 '必须反向
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                '设置给药途径行
                If Val(.TextMatrix(i, COL_执行性质)) = 5 And int执行性质 <> 5 Then
                    str执行性质 = "自备药"
                ElseIf Val(.TextMatrix(i, COL_执行性质)) <> 5 And int执行性质 = 5 Then
                    str执行性质 = "离院带药"
                Else
                    str执行性质 = ""
                End If
                .TextMatrix(i, COL_相关ID) = "" '必须清除作为标志
                .TextMatrix(i, COL_相关ID) = AdviceSet给药途径(i, lng给药途径ID, str执行性质, lng给药执行ID, str滴速)
                
                If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    .TextMatrix(i, COL_EDIT) = 2 '标记为已修改
                    .TextMatrix(i, COL_状态) = 1 '修改后变为新开
                End If
            End If
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '标记为未保存
    End With
End Sub

Private Sub ShowAdvice()
'功能：显示当前界面条件下的医嘱记录
'说明：1.根据程序编辑方式,相关的数据行是按序号严格排列在一在的。
'      2.这里不处理一并给药的边框及配方行高，状态颜色等格式内容,它们已在读取或编辑时设置
    Dim lngRow As Long, blnHide As Boolean, i As Long
    
    Screen.MousePointer = 11
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
        
    '先删除无效行
    For i = vsAdvice.Rows - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) = 0 Then vsAdvice.RemoveItem i
    Next
    
    '根据当前期效,婴儿显示
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                blnHide = False
                '隐藏以下数据行：
                '1.成药的给药途径行
                '2.手术的附加手术及麻醉项目行
                '3.检查组合的部位行
                '4.中药配方的组成味中药及中药煎法行
                '5.(一并采集的)检验项目
                '6.输血项目的输血途径
                If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_类别)) > 0 _
                    And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '计算诊疗单价:这里为加快速度,只读取新开的,其它的进入再读
                If Not .RowHidden(i) _
                    And Val(.TextMatrix(i, COL_状态)) = 1 And .TextMatrix(i, COL_单价) = "" Then
                    .TextMatrix(i, COL_单价) = GetItemPrice(i)
                End If
            Else
                .RowHidden(i) = True
            End If
        Next
    End With
    
    '没有数据行,添加一行空
    If lngRow = 0 Then
        vsAdvice.AddItem ""
        lngRow = vsAdvice.Rows - 1
    End If
    
    vsAdvice.Row = lngRow
    If vsAdvice.RowData(lngRow) = 0 Then
        vsAdvice.Col = vsAdvice.FixedCols
    Else
        vsAdvice.Col = col_医嘱内容
    End If
    vsAdvice.Redraw = flexRDDirect
    mblnRowChange = True
    
    '显示当前行:进入时在FormLoad中处理,以加快速度
    If Me.Visible Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Call CalcAdviceMoney '显示新开医嘱金额
    
    Screen.MousePointer = 0
End Sub

Private Function SaveAdvice() As Boolean
'功能：保存当前病人的医嘱记录
    Dim arrSQL As Variant, arrDelID() As String
    Dim strSQL As String, dbl总量 As Double
    Dim arrAppend As Variant, i As Long, j As Long
    Dim blnChecked As Boolean
    Dim curDate As Date
    Dim blnTrans As Boolean
    Dim blnDiagChange As Boolean, blnRecipeNo As Boolean
    Dim strFilter As String, strTmp As String
    Dim rsMsg As ADODB.Recordset
    Dim blnMsgOk As Boolean
    Dim lng诊断记录id As Long
    Dim str记录人 As String
    Dim str给药IDs As String
    Dim lng相关ID As Long
    Dim str撤销给药IDs As String
    Dim rsCard As ADODB.Recordset
    Dim rsBlood As ADODB.Recordset, intBloodState As Integer
    Dim varTmp As Variant
    Dim blnRIS As Boolean 'RIS接口
    Dim strNewAdvice As String '新增和修改后的医嘱 格式 医嘱ID:诊疗项目ID,....
    Dim str本次变动 As String '本次删除和修改的医嘱，用于判断RIS预约的医嘱是否已经被预约，已经被预约的医嘱不能删
    Dim rsTmp As ADODB.Recordset
    Dim rs输血 As ADODB.Recordset '输血医嘱血库接口调用时的参数信息
    Dim bln血库 As Boolean '是否需要进行血库接口调用
    Dim strAdvices输血 As String
    Dim strErr As String
    Dim bln用血审核 As Boolean
    
    If HaveRIS And gbln启用影像信息系统预约 Then
        blnRIS = True
    ElseIf gbln启用影像信息系统接口 = True And gbln启用影像信息系统预约 = True Then
        MsgBox "RIS接口创建失败，不能继续当前操作。可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '替换成真实的医嘱ID，包括诊断医嘱对应
    Call MakeRealID
    
    'Pass自动用药审查
    If mblnPass Then
        If gobjPass.HaveRecipNo() Then
            Call UpdateRecipeNo '杭州逸曜
            blnRecipeNo = True
        End If
        If gobjPass.zlPassCheck(mobjPassMap) Then
            If Not gobjPass.zlPassAdviceSave(mobjPassMap, mblnNoSave) Then Exit Function
        End If
    End If

    '调用外挂接口
    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
        If zlPluginAdviceSave = False Then Exit Function
    End If

    Screen.MousePointer = 11
    
    If gbln血库系统 Then bln血库 = InitObjBlood()
    Set rs输血 = New ADODB.Recordset
    With rs输血
        .Fields.Append "医嘱ID", adBigInt
        .Fields.Append "类型", adInteger '0－新开，1－修改，2－删除
        .Fields.Append "状态", adInteger '0-调用血库接口,1-不掉用血库接口(用血医嘱审核)
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    If mstrDel输血 <> "" Then
        arrDelID = Split(mstrDel输血, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                Call rs输血.AddNew(Array("医嘱ID", "类型"), Array(Val(arrDelID(i)), 2))
            End If
        Next
    End If
    
    If Not (mclsMipModule Is Nothing) Then
        blnMsgOk = mclsMipModule.IsConnect
    Else
        blnMsgOk = False
    End If
    
    '生成SQL
    arrSQL = Array()
    curDate = zlDatabase.Currentdate

        '处方审查系统
    If InitObjRecipeAudit(p门诊医嘱下达) Then
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 2 Then
                If vsAdvice.TextMatrix(i, COL_类别) = "5" Or vsAdvice.TextMatrix(i, COL_类别) = "6" Or vsAdvice.TextMatrix(i, COL_类别) = "7" Then
                    If vsAdvice.TextMatrix(i, COL_处方审查状态) = "1" Or vsAdvice.TextMatrix(i, COL_处方审查状态) = "2" Then
                        If InStr("," & mstrAduitDelIDs & ",", "," & vsAdvice.TextMatrix(i, COL_相关ID) & ",") = 0 And vsAdvice.TextMatrix(i, COL_相关ID) <> "" Then
                            mstrAduitDelIDs = mstrAduitDelIDs & "," & vsAdvice.TextMatrix(i, COL_相关ID)
                        End If
                    End If
                    If vsAdvice.TextMatrix(i, COL_处方审查状态) <> "" Then
                        If InStr("," & str撤销给药IDs & ",", "," & vsAdvice.TextMatrix(i, COL_相关ID) & ",") = 0 And vsAdvice.TextMatrix(i, COL_相关ID) <> "" Then
                            str撤销给药IDs = str撤销给药IDs & "," & vsAdvice.TextMatrix(i, COL_相关ID)
                        End If
                    End If
                End If
            End If
        Next
        mstrAduitDelIDs = Mid(mstrAduitDelIDs, 2)
        For i = 0 To UBound(Split(mstrAduitDelIDs, ","))
            If InStr("," & str撤销给药IDs & ",", "," & Split(mstrAduitDelIDs, ",")(i) & ",") = 0 Then
                str撤销给药IDs = str撤销给药IDs & "," & Split(mstrAduitDelIDs, ",")(i)
            End If
        Next
        str撤销给药IDs = Mid(str撤销给药IDs, 2)
        '如果有一个被假删除，则整组假删除
        If mstrAduitDelIDs <> "" Then
            For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
                If InStr("," & mstrAduitDelIDs & ",", "," & vsAdvice.RowData(i) & ",") > 0 Then
                    lng相关ID = vsAdvice.RowData(i)
                    vsAdvice.RowData(i) = zlDatabase.GetNextID("病人医嘱记录")
                    vsAdvice.TextMatrix(i, COL_EDIT) = 1
                    For j = i - 1 To vsAdvice.FixedRows Step -1
                        If Val(vsAdvice.TextMatrix(j, COL_相关ID)) <> lng相关ID Then Exit For
                        vsAdvice.TextMatrix(j, COL_相关ID) = vsAdvice.RowData(i)
                        vsAdvice.RowData(j) = zlDatabase.GetNextID("病人医嘱记录")
                        vsAdvice.TextMatrix(j, COL_EDIT) = 1
                    Next
                End If
            Next

            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱记录_处方审查删除('" & mstrAduitDelIDs & "')"
        End If
        If str撤销给药IDs <> "" Then
            '产生了处方审查数据的都要撤销
            Call gobjRecipeAudit.CancelData(str撤销给药IDs, "")
        End If
    End If
    
    '删除了的记录
    arrDelID = Split(mstrDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & Val(arrDelID(i)) & ")"
            If blnRIS Then str本次变动 = str本次变动 & "," & Val(arrDelID(i))
        End If
    Next

    '编辑标志：0-原始的,1-新增的,2-修改了内容,3-修改了序号
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then    '所有医嘱记录
                '总量转换
                dbl总量 = 0
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    If Val(.TextMatrix(i, COL_总量)) <> 0 Then
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            '成药转换成零售单位
                            dbl总量 = Format(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_门诊包装)), "0.00000")
                        Else
                            '中药配方付数或非药临嘱总量,不转换
                            dbl总量 = Val(.TextMatrix(i, COL_总量))
                        End If
                    End If
                End If

                '33629
                '麻醉药品和第一类精神药品处方还应当包括患者身份证明编号，代办人姓名、身份证明编号。
                If blnChecked = False And mblnAddAgent Then
                    If .RowHidden(i) = False And AgentInfo.本次就诊已录入 = False Then
                        If Val(.TextMatrix(i, COL_状态)) = 1 And InStr(",麻醉药,毒性药,精神I类,", "," & Trim(.TextMatrix(i, COL_毒理分类)) & ",") > 0 Then
                            blnChecked = frmAgentInfo.ShowMe(Me, mlng病人ID, mlng挂号ID, mstr姓名, mstr身份证号, AgentInfo.代办人姓名, AgentInfo.代办人身份证号)
                            If blnChecked Then
                                Call GetAgentInfo
                            Else
                                Screen.MousePointer = 0
                                Exit Function
                            End If
                        End If
                    End If
                End If
                
                If .TextMatrix(i, COL_类别) = "K" And bln血库 Then
                    If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                        intBloodState = 0
                        '用血医嘱已经发血，则修改医嘱审核状态为2
                        If Val(.TextMatrix(i, COL_检查方法)) = 1 Then
                            If gobjPublicBlood.GetPrepareBloodRs(Val(.RowData(i)), rsBlood) = True Then
                                If Val(rsBlood!记录性质 & "") = 2 And Val(rsBlood!记录状态 & "") = 1 Then
                                    .TextMatrix(i, COL_审核状态) = 2
                                    intBloodState = 1
                                    bln用血审核 = True
                                End If
                            End If
                        End If
                        Call rs输血.AddNew(Array("医嘱ID", "类型", "状态"), Array(Val(.RowData(i)), 1, intBloodState))
                    ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        Call rs输血.AddNew(Array("医嘱ID", "类型"), Array(Val(.RowData(i)), 0))
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_EDIT)) = 3 Then    '修改了序号的记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新序号(" & .RowData(i) & "," & Val(.TextMatrix(i, COL_序号)) & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 2 Then    '修改了内容的记录
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Update(" & _
                                             .RowData(i) & "," & ZVal(.TextMatrix(i, COL_相关ID)) & "," & _
                                             Val(.TextMatrix(i, COL_序号)) & "," & Val(.TextMatrix(i, COL_状态)) & ",1," & _
                                             Val(.TextMatrix(i, COL_诊疗项目ID)) & "," & ZVal(.TextMatrix(i, COL_收费细目ID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_天数)) & "," & ZVal(.TextMatrix(i, COL_单量)) & "," & ZVal(dbl总量) & "," & _
                                             "'" & Replace(.TextMatrix(i, col_医嘱内容), "'", "''") & "','" & Replace(.TextMatrix(i, COL_医生嘱托), "'", "''") & "'," & _
                                             "'" & .TextMatrix(i, COL_标本部位) & "','" & .TextMatrix(i, COL_频率) & "'," & _
                                             ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & "," & _
                                             "'" & .TextMatrix(i, COL_间隔单位) & "','" & .TextMatrix(i, COL_执行时间) & "'," & _
                                             Val(.TextMatrix(i, COL_计价性质)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & _
                                             Val(.TextMatrix(i, COL_执行性质)) & "," & Val(.TextMatrix(i, COL_标志)) & "," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                                             mlng病人科室id & "," & Val(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                             "'" & .TextMatrix(i, COL_检查方法) & "'," & Val(.TextMatrix(i, COL_执行标记)) & "," & _
                                             "NULL,'" & .Cell(flexcpData, i, COL_医生嘱托) & "','" & UserInfo.姓名 & "'," & ZVal(.TextMatrix(i, COL_零费记帐)) & "," & _
                                             ZVal(Val(.TextMatrix(i, COL_用药目的))) & ",'" & .TextMatrix(i, COL_用药理由) & "'," & ZVal(Val(.TextMatrix(i, COL_审核状态))) & ",'" & .TextMatrix(i, COL_超量说明) & "'" & _
                                             ",'',Null," & ZVal(Val(.TextMatrix(i, COL_组合项目ID))) & ",NULL," & IIF(blnRecipeNo, Val(.TextMatrix(i, COL_处方序号)), "NULL") & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then    '新增的记录
                
                    If .TextMatrix(i, COL_类别) = "K" Then
                        '输血类医嘱从新获取医嘱审核状态
                        If TypeName(.Cell(flexcpData, i, COL_申请序号)) = "Recordset" Then
                            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, i, COL_申请序号))
                            If Not rsCard.EOF Then
                                rsCard.MoveFirst
                                strTmp = Nvl(rsCard!申请项目 & "", .TextMatrix(i, COL_诊疗项目ID) & "," & dbl总量)
                                .TextMatrix(i, COL_审核状态) = GetBloodVerifyState(1, mlng病人ID, mlng挂号ID, .TextMatrix(i, COL_标本部位), GetBloodTotalByML(strTmp), Val(.TextMatrix(i, COL_标志)), Val(.TextMatrix(i, COL_检查方法)), Val(.TextMatrix(i, COL_婴儿)), .RowData(i), strTmp)
                                If i < .Rows - 1 Then
                                    If Val(.TextMatrix(i + 1, COL_相关ID)) = .RowData(i) Then
                                        .TextMatrix(i + 1, COL_审核状态) = .TextMatrix(i, COL_审核状态)
                                    End If
                                End If
                                Call SetRow标志图标(i, 2)
                            End If
                        Else
                            '非申请单下达的医嘱
                            If bln血库 Then
                                strTmp = .TextMatrix(i, COL_诊疗项目ID) & "," & dbl总量
                                .TextMatrix(i, COL_审核状态) = GetBloodVerifyState(1, mlng病人ID, mlng挂号ID, .TextMatrix(i, COL_标本部位), GetBloodTotalByML(strTmp), Val(.TextMatrix(i, COL_标志)), Val(.TextMatrix(i, COL_检查方法)), Val(.TextMatrix(i, COL_婴儿)), .RowData(i), strTmp)
                                If i < .Rows - 1 Then
                                    If Val(.TextMatrix(i + 1, COL_相关ID)) = .RowData(i) Then
                                        .TextMatrix(i + 1, COL_审核状态) = .TextMatrix(i, COL_审核状态)
                                    End If
                                End If
                                Call SetRow标志图标(i, 2)
                            End If
                        End If
                    End If
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                                             .RowData(i) & "," & ZVal(.TextMatrix(i, COL_相关ID)) & "," & _
                                             Val(.TextMatrix(i, COL_序号)) & ",1," & mlng病人ID & ",NULL," & _
                                             Val(.TextMatrix(i, COL_婴儿)) & "," & Val(.TextMatrix(i, COL_状态)) & ",1," & _
                                             "'" & IIF(.TextMatrix(i, COL_类别) = "*", "", .TextMatrix(i, COL_类别)) & "'," & Val(.TextMatrix(i, COL_诊疗项目ID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_收费细目ID)) & "," & _
                                             ZVal(.TextMatrix(i, COL_天数)) & "," & ZVal(.TextMatrix(i, COL_单量)) & "," & ZVal(dbl总量) & "," & _
                                             "'" & Replace(.TextMatrix(i, col_医嘱内容), "'", "''") & "','" & Replace(.TextMatrix(i, COL_医生嘱托), "'", "''") & "'," & _
                                             "'" & .TextMatrix(i, COL_标本部位) & "','" & .TextMatrix(i, COL_频率) & "'," & _
                                             ZVal(.TextMatrix(i, COL_频率次数)) & "," & ZVal(.TextMatrix(i, COL_频率间隔)) & "," & _
                                             "'" & .TextMatrix(i, COL_间隔单位) & "','" & .TextMatrix(i, COL_执行时间) & "'," & _
                                             Val(.TextMatrix(i, COL_计价性质)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & _
                                             Val(.TextMatrix(i, COL_执行性质)) & "," & Val(.TextMatrix(i, COL_标志)) & "," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_开始时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                                             mlng病人科室id & "," & Val(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                             "To_Date('" & Format(.Cell(flexcpData, i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                                             "'" & mstr挂号单 & "'," & ZVal(mlng前提ID) & ",'" & .TextMatrix(i, COL_检查方法) & "'," & _
                                             Val(.TextMatrix(i, COL_执行标记)) & ",NULL,'" & .Cell(flexcpData, i, COL_医生嘱托) & "','" & UserInfo.姓名 & "'," & ZVal(.TextMatrix(i, COL_零费记帐)) & "," & _
                                             ZVal(Val(.TextMatrix(i, COL_用药目的))) & ",'" & .TextMatrix(i, COL_用药理由) & "'," & ZVal(Val(.TextMatrix(i, COL_审核状态))) & "," & ZVal(Val(.TextMatrix(i, COL_申请序号))) & ",'" & .TextMatrix(i, COL_超量说明) & _
                                             "',Null," & ZVal(Val(.TextMatrix(i, COL_配方ID))) & ",Null," & ZVal(Val(.TextMatrix(i, COL_组合项目ID))) & ",NULL," & IIF(blnRecipeNo, Val(.TextMatrix(i, COL_处方序号)), "NULL") & ")"
                    
                    
                    '收集本次新增的医嘱用于关联病人危急值记录
                    If mlng危急值ID <> 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_病人危急值医嘱_Update(1," & mlng危急值ID & "," & .RowData(i) & ")"
                        End If
                    End If
                    
                    '输血申请单的额外数据保存
                    If .TextMatrix(i, COL_类别) = "K" Then
                        If TypeName(.Cell(flexcpData, i, COL_申请序号)) = "Recordset" Then
                            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, i, COL_申请序号))
                            If Not rsCard.EOF Then
                                rsCard.MoveFirst
                                strTmp = rsCard!申请其他项目SQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[相关ID]", .RowData(i))
                                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                    arrSQL(UBound(arrSQL)) = strTmp
                                End If
                                strTmp = rsCard!检验项目SQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[相关ID]", .RowData(i))
                                    varTmp = Split(strTmp, "<splitSQL>")
                                    For j = 0 To UBound(varTmp)
                                        If varTmp(j) <> "" Then
                                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                            arrSQL(UBound(arrSQL)) = varTmp(j)
                                        End If
                                    Next
                                End If
                                strTmp = rsCard!诊断关联信息SQL & ""
                                If strTmp <> "" Then
                                    strTmp = Replace(strTmp, "[相关ID]", .RowData(i))
                                    varTmp = Split(strTmp, "<splitSQL>")
                                    For j = 0 To UBound(varTmp)
                                        If j = 1 Then
                                            If varTmp(j) <> "" Then
                                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                                arrSQL(UBound(arrSQL)) = varTmp(j)
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If
                
                If blnRIS Then
                    If InStr(",F,D,", .TextMatrix(i, COL_类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_操作类型))) > 0 And .TextMatrix(i, COL_类别) = "E" Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                                str本次变动 = str本次变动 & "," & .RowData(i)
                            End If
                            If InStr(",1,2,", Val(.TextMatrix(i, COL_EDIT))) > 0 And Val(.TextMatrix(i, COL_标志)) <> 1 Then
                                strNewAdvice = strNewAdvice & "," & .RowData(i) & ":" & Val(.TextMatrix(i, COL_诊疗项目ID))
                            End If
                        End If
                    End If
                End If
                
                '标记免试的
                If .Cell(flexcpData, i, COL_免试) & "" <> "" & .TextMatrix(i, COL_免试) Then
                    If .TextMatrix(i, COL_免试) = "1" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_皮试(" & Val(.RowData(i)) & ",'免试',NULL)"
                    ElseIf .TextMatrix(i, COL_免试) = "0" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_皮试(" & Val(.RowData(i)) & ",'',NULL)"
                    End If
                End If
                '单据申请附项
                If .TextMatrix(i, COL_附项) <> "" And .Cell(flexcpData, i, COL_附项) = 1 Then
                    arrAppend = Split(.TextMatrix(i, COL_附项), "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & .RowData(i) & "," & _
                                                 "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                                                 j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                                                 IIF(j = 0, ",1", "") & ")"
                    Next
                End If

                'Pass:更新审查结果
                If Val(.Cell(flexcpData, i, COL_序号)) = 1 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_更新审查(" & .RowData(i) & "," & _
                                             IIF(CStr(.Cell(flexcpData, i, COL_警示)) = "", "NULL", Val(.Cell(flexcpData, i, COL_警示))) & ")"
                End If
            End If
        Next
    End With

    '诊断部分内容(先要保存医嘱ID数据)
    If lbl诊断.Tag = "1" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_Delete(" & mlng病人ID & "," & mlng挂号ID & ",3,Null,'1,11')"
        If blnMsgOk Then Call InitRsMsg(rsMsg)
        With vsDiag
            j = 0
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, col诊断)) <> "" Then
                    blnDiagChange = True
                    If Val(.Cell(flexcpData, i, col诊断ID)) > 0 And Not mrsDiag Is Nothing Then
                        strFilter = "诊断类型=" & IIF(.Cell(flexcpData, i, col中医) = 1, "11", "1") & " And 记录来源=3 And 疾病id=" & ZVal(.TextMatrix(i, col疾病ID)) & " And 诊断id=" & ZVal(.TextMatrix(i, col诊断ID))
                        strTmp = IIF(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & IIF(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "")
                        strFilter = strFilter & " And 诊断描述= '" & strTmp & "'"
                        If IsDate(.TextMatrix(i, col发病时间)) Then
                            strFilter = strFilter & " And  发病时间= '" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "'"
                        Else
                            strFilter = strFilter & " And  发病时间= Null "
                        End If

                        strFilter = strFilter & " And  证候ID= " & ZVal(.TextMatrix(i, col证候ID))

                        strFilter = strFilter & " And 是否疑诊=" & Val(.Cell(flexcpData, i, col疑诊))
                        mrsDiag.Filter = strFilter
                        blnDiagChange = mrsDiag.EOF
                    End If
                    lng诊断记录id = zlDatabase.GetNextID("病人诊断记录")
                    strTmp = .TextMatrix(i, col医嘱ID)
                    If Len(strTmp) > 4000 Then
                        strTmp = ""
                    End If
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1): j = j + 1
                    If blnDiagChange Then
                        str记录人 = UserInfo.姓名
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                                                 "Null," & IIF(.Cell(flexcpData, i, col中医) = 1, "11", "1") & "," & ZVal(.TextMatrix(i, col疾病ID)) & "," & _
                                                 ZVal(.TextMatrix(i, col诊断ID)) & "," & ZVal(.TextMatrix(i, col证候ID)) & "," & _
                                                 "'" & IIF(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & IIF(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "") & "',Null,Null," & Val(.Cell(flexcpData, i, col疑诊)) & "," & _
                                                 "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                 IIF(strTmp = "", "null", "'" & strTmp & "'") & "," & j & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & UserInfo.姓名 & "'," & lng诊断记录id & ")"
                    Else
                        str记录人 = mrsDiag!记录人
                        arrSQL(UBound(arrSQL)) = "ZL_病人诊断记录_INSERT(" & mlng病人ID & "," & mlng挂号ID & ",3," & _
                                                 "Null," & IIF(.Cell(flexcpData, i, col中医) = 1, "11", "1") & "," & ZVal(.TextMatrix(i, col疾病ID)) & "," & _
                                                 ZVal(.TextMatrix(i, col诊断ID)) & "," & ZVal(.TextMatrix(i, col证候ID)) & "," & _
                                                 "'" & IIF(.TextMatrix(i, col编码) <> "", "(" & .TextMatrix(i, col编码) & ")", "") & .TextMatrix(i, col诊断) & IIF(.TextMatrix(i, col中医证候) <> "", "(" & .TextMatrix(i, col中医证候) & ")", "") & "',Null,Null," & Val(.Cell(flexcpData, i, col疑诊)) & "," & _
                                                 "To_Date('" & Format(CDate(mrsDiag!记录日期), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                 IIF(strTmp = "", "null", "'" & strTmp & "'") & "," & j & ",Null,Null,To_date('" & Format(.TextMatrix(i, col发病时间), "yyyy-MM-dd HH:mm") & "','yyyy-MM-dd HH24:mi'),'" & mrsDiag!记录人 & "'," & lng诊断记录id & ")"
                    End If
                    strTmp = .TextMatrix(i, col医嘱ID)
                    If Len(strTmp) > 4000 Then Call Make诊断医嘱对应(arrSQL, strTmp, lng诊断记录id)
                    If blnMsgOk Then
                        Call SendMsg诊断(i, j, lng诊断记录id, str记录人, rsMsg)
                    End If
                End If
            Next
        End With
    End If
    
    If bln血库 Then
        bln血库 = rs输血.RecordCount > 0
        If bln血库 Then rs输血.MoveFirst
    End If
    
    If blnRIS Then
        If str本次变动 <> "" Then
            str本次变动 = Mid(str本次变动, 2)
            Set rsTmp = GetDataRIS预约(str本次变动)
            If rsTmp.RecordCount > 0 Then
                str本次变动 = "有数据"
            Else
                str本次变动 = ""
            End If
        End If
    Else
        str本次变动 = ""
    End If
    '处理RIS的取消预约
    If str本次变动 <> "" Then
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!预约id & "")) Then
                MsgBox "当前启用了影像信息系统接口，本次操作删除或修改了已经预约医嘱，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    '提交数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    '处理血库的接口
    If bln血库 Then
        For i = 1 To rs输血.RecordCount
            If Val("" & rs输血!状态) <> 1 Then
                If gobjPublicBlood.AdviceOperation(p门诊医嘱下达, Val(rs输血!医嘱ID & ""), Val(rs输血!类型 & ""), , strErr) = False Then
                    gcnOracle.RollbackTrans: blnTrans = False
                    Screen.MousePointer = 0
                    MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            rs输血.MoveNext
        Next
    End If
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    mlngID序列 = 0
    'Pass自动用药审查上传处方
    If mblnPass Then
        Call gobjPass.zlPassUpLoad(mobjPassMap)
    End If
    
    If blnMsgOk Then
        If Not (mrs诊断 Is Nothing) Then
            mrs诊断.Filter = "状态 = 0"
            If Not mrs诊断.EOF Then
                For i = 1 To mrs诊断.RecordCount
                    Call ZLHIS_CIS_011(mclsMipModule, mlng病人ID, mstr姓名, 1, mlng挂号ID, mlng病人科室id, mrs诊断!ID, mrs诊断!诊断编码, mrs诊断!疾病编码)
                    mrs诊断.Delete
                    mrs诊断.MoveNext
                Next
            End If
            mrs诊断.Filter = "状态 <> 0"
            If Not mrs诊断.EOF Then
                For i = 1 To mrs诊断.RecordCount
                    mrs诊断!状态 = 0
                    mrs诊断.MoveNext
                Next
            End If
            mrs诊断.Filter = 0
        End If
        If Not (rsMsg Is Nothing) Then
            rsMsg.Filter = "状态 = 1"
            If Not rsMsg.EOF Then
                For i = 1 To rsMsg.RecordCount
                    Call ZLHIS_CIS_010(mclsMipModule, mlng病人ID, mstr姓名, 1, mlng挂号ID, mlng病人科室id, _
                        rsMsg!诊断id, rsMsg!诊断类型, rsMsg!是否疑诊, rsMsg!诊断次序, rsMsg!诊断编码, rsMsg!疾病编码, rsMsg!疾病附码, rsMsg!疾病类别, rsMsg!证候编码, rsMsg!证候名称, rsMsg!记录日期, rsMsg!记录人员)
                    rsMsg.MoveNext
                Next
            End If
        End If
    End If
    
    If bln用血审核 Then Call ReadMsg
    
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    '调用删除后外挂接口
    On Error Resume Next
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            If Not gobjPlugIn Is Nothing Then
                Call gobjPlugIn.AdviceDeleted(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(arrDelID(i)), mint场合)
                Call zlPlugInErrH(err, "AdviceDeleted")
            End If
        End If
    Next
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0

    '调用处方审查系统检查
    If InitObjRecipeAudit(p门诊医嘱下达) And mblnNoSave Then
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If vsAdvice.TextMatrix(i, COL_类别) = "E" And vsAdvice.TextMatrix(i, COL_操作类型) = "2" Then
                '当前新开状态的医嘱带需要传入
                If Val(vsAdvice.TextMatrix(i, COL_状态)) = 1 Then
                    str给药IDs = str给药IDs & "," & vsAdvice.RowData(i)
                End If
            End If
        Next
        If Mid(str给药IDs, 2) <> "" Then
            Call gobjRecipeAudit.AutoAudit(Me, 1, Mid(str给药IDs, 2), mlng病人科室id, 0, mlng病人ID, mlng挂号ID)
        End If
    End If

    '保存成功后,所有记录变成原始记录
    With vsAdvice
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, COL_EDIT) = 0
                .Cell(flexcpData, i, COL_序号) = Empty    'Pass:保存后清除标志
                .Cell(flexcpData, i, COL_附项) = 0    '附项:保存后清除标志
            End If
        Next
    End With

    '保存后重新进入行(比如开始时间不准改了)
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Screen.MousePointer = 0
    lbl诊断.Tag = ""
    mblnNoSave = False
    mstrDelIDs = ""
    mstrAduitDelIDs = ""
    mstrDel输血 = ""
    SaveAdvice = True
    mblnOK = True
    mlng危急值ID = 0
    
    '医嘱数据保存提交后，再RIS预约，只有普能病人才预约
    If blnRIS And mbytPatiType = 1 Then
        If strNewAdvice <> "" Then
            strTmp = Mid(strNewAdvice, 2)
            varTmp = Split(strTmp, ",")
            On Error Resume Next
            For i = 0 To UBound(varTmp)
                strTmp = varTmp(i)
                Call gobjRis.HISScheduling(1, Val(Split(strTmp, ":")(0)), Val(Split(strTmp, ":")(1)))
            Next
            err.Clear: On Error GoTo 0
        End If
    End If
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SendMsg诊断(ByVal lngRow As Long, ByVal lng次序 As Long, ByVal lngID As Long, ByVal str记录人 As String, ByRef rsMsg As ADODB.Recordset)
    Dim i As Long
    With vsDiag
        rsMsg.AddNew
        rsMsg!诊断id = lngID
        rsMsg!诊断类型 = IIF(.Cell(flexcpData, lngRow, col中医) = 1, "11", "1")
        rsMsg!是否疑诊 = Val(.Cell(flexcpData, lngRow, col疑诊))
        rsMsg!诊断次序 = lng次序
        rsMsg!诊断编码 = .TextMatrix(lngRow, col诊断编码)
        rsMsg!疾病编码 = .TextMatrix(lngRow, col疾病编码)
        rsMsg!疾病附码 = .TextMatrix(lngRow, col疾病附码)
        rsMsg!疾病类别 = .TextMatrix(lngRow, col疾病类别)
        rsMsg!证候编码 = .TextMatrix(lngRow, col证候编码)
        rsMsg!证候名称 = .TextMatrix(lngRow, col中医证候)
        rsMsg!记录日期 = Format(.TextMatrix(lngRow, col发病时间), "yyyy-MM-dd HH:mm:ss")
        rsMsg!记录人员 = str记录人
        rsMsg!状态 = 1
        
        mrs诊断.Filter = "显示编码='" & .TextMatrix(lngRow, col编码) & "'"
        
        If mrs诊断.EOF Then
            mrs诊断.AddNew
            mrs诊断!ID = lngID
            mrs诊断!显示编码 = .TextMatrix(lngRow, col编码)
            mrs诊断!诊断编码 = .TextMatrix(lngRow, col诊断编码)
            mrs诊断!疾病编码 = .TextMatrix(lngRow, col疾病编码)
            mrs诊断!状态 = 2
            mrs诊断.Update
        Else
            rsMsg!状态 = 0
            mrs诊断!状态 = 1
            mrs诊断!ID = lngID
        End If
        rsMsg.Update
    End With
End Sub

Private Sub InitRsMsg(ByRef rsMsg As ADODB.Recordset)
'功能：初始化消息记录集
    Set rsMsg = New ADODB.Recordset
    rsMsg.Fields.Append "诊断id", adBigInt
    rsMsg.Fields.Append "诊断类型", adVarChar, 6
    rsMsg.Fields.Append "是否疑诊", adVarChar, 6
    rsMsg.Fields.Append "诊断次序", adBigInt
    rsMsg.Fields.Append "诊断编码", adVarChar, 60
    rsMsg.Fields.Append "疾病编码", adVarChar, 60
    rsMsg.Fields.Append "疾病附码", adVarChar, 60
    rsMsg.Fields.Append "疾病类别", adVarChar, 60
    rsMsg.Fields.Append "证候编码", adVarChar, 60
    rsMsg.Fields.Append "证候名称", adVarChar, 120
    rsMsg.Fields.Append "记录日期", adVarChar, 60
    rsMsg.Fields.Append "记录人员", adVarChar, 120
    rsMsg.Fields.Append "状态", adBigInt '1 －新增
    rsMsg.CursorLocation = adUseClient
    rsMsg.LockType = adLockOptimistic
    rsMsg.CursorType = adOpenStatic
    rsMsg.Open
End Sub

Private Function LoadAdvice() As Boolean
''功能：读取当前病人的医嘱记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln配方 As Boolean
    Dim i As Long, j As Long, lng相关ID As Long

    Screen.MousePointer = 11

    On Error GoTo errH

    '下医嘱缺省的天数
    If msng天数 = 0 Then msng天数 = 1

    '临床和医技不互相编辑
    strSQL = " And Nvl(A.前提ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X)"
    
    '读取"1-新开,8-已停止(已发送)"的医嘱
    strSQL = _
    " Select A.ID,A.相关ID,Nvl(A.婴儿,0) as 婴儿,A.序号,A.医嘱期效,A.医嘱状态,A.诊疗类别,A.诊疗项目ID,B.名称,A.标本部位,A.检查方法," & _
             " A.执行标记,A.收费细目ID,A.开始执行时间,A.医嘱内容,A.医生嘱托,A.单次用量,A.天数,A.总给予量,B.计算单位,A.执行频次," & _
             " A.频率次数,A.频率间隔,A.间隔单位,B.计算方式,B.执行频率,B.操作类型,B.单独应用,B.执行分类,A.计价特性,A.执行时间方案,A.执行性质," & _
             " A.执行科室ID,A.开嘱科室ID,A.开嘱医生,A.开嘱时间,A.紧急标志,Decode(Nvl(c.处方限量, 0), 0, b.录入限量, c.处方限量) As 处方限量,C.处方职务,C.毒理分类,C.抗生素,C.药品剂型," & _
             " D.剂量系数,D.门诊包装,D.门诊单位,F.计算单位 as 散装单位,E.跟踪在用,D.门诊可否分零 As 可否分零,A.审查结果," & _
             " Decode(A.新开签名ID,NULL,0,1) as 签名否,A.摘要,A.零费记帐,A.用药目的,A.用药理由,A.审核状态,A.皮试结果,A.超量说明," & _
             " a.配方ID,c.临床自管药,d.高危药品,a.组合项目ID,b.撤档时间,C.溶媒,a.申请序号,J.状态 as 处方审查状态,J.审查结果 as 处方审查结果,A.处方序号,d.基本药物,Nvl(Max(g.危急值id), Max(h.危急值id)) As 危急值id" & _
             " From 病人医嘱记录 A,诊疗项目目录 B,药品特性 C,药品规格 D,材料特性 E,收费项目目录 F, 处方审查明细 I, 处方审查记录 J, 病人危急值医嘱 G,病人危急值医嘱 H" & _
             " Where Nvl(A.医嘱期效,0)=1 And A.诊疗项目ID=B.ID And A.诊疗项目ID=C.药名ID(+)  And a.ID = i.医嘱ID(+) And I.审方ID = J.ID(+) and (I.最后提交 =1 Or I.审方ID is NULL) And Nvl(A.执行标记,0)<>-1 And a.ID = h.医嘱ID(+) and a.相关ID=g.医嘱ID(+)  " & _
             " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=E.材料ID(+) And E.材料ID=F.ID(+) And A.医嘱状态 IN(1,8)" & strSQL & _
             " And A.病人ID+0=[1] And A.挂号单=[2] And A.开始执行时间 is Not NULL And A.病人来源<>3"
    strSQL = strSQL & " " & _
            "Group By a.Id, a.相关id, a.婴儿, a.序号, a.医嘱期效, a.医嘱状态, a.诊疗类别, a.诊疗项目id, b.名称, a.标本部位, a.检查方法, a.执行标记, a.收费细目id, a.开始执行时间," & vbNewLine & _
            "         a.医嘱内容, a.医生嘱托, a.单次用量, a.天数, a.总给予量, b.计算单位, a.执行频次, a.频率次数, a.频率间隔, a.间隔单位, b.计算方式, b.执行频率, b.操作类型,b.单独应用, b.执行分类," & vbNewLine & _
            "         a.计价特性, a.执行时间方案, a.执行性质, a.执行科室id, a.开嘱科室id, a.开嘱医生, a.开嘱时间, a.紧急标志, c.处方限量, c.处方职务, c.毒理分类, c.抗生素, c.药品剂型," & vbNewLine & _
            "         d.剂量系数, d.门诊包装, d.门诊单位, f.计算单位, e.跟踪在用, d.门诊可否分零, a.审查结果, a.新开签名id, a.摘要, a.零费记帐, a.用药目的, a.用药理由, a.审核状态," & vbNewLine & _
            "         a.皮试结果, a.超量说明, a.配方id, c.临床自管药, d.高危药品, a.组合项目id, b.撤档时间, c.溶媒, a.申请序号, j.状态, j.审查结果,d.基本药物, a.处方序号,b.录入限量" & vbNewLine & _
            "Order By a.婴儿, a.序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, IIF(mstr前提IDs = "", "0", mstr前提IDs))
    On Error GoTo 0

    If Not rsTmp.EOF Then
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                bln配方 = False

                .RowData(i) = CLng(rsTmp!ID)
                .TextMatrix(i, COL_EDIT) = 0    '原始记录
                .TextMatrix(i, COL_相关ID) = Nvl(rsTmp!相关ID)
                .TextMatrix(i, COL_婴儿) = Nvl(rsTmp!婴儿, 0)
                .TextMatrix(i, COL_序号) = rsTmp!序号
                .TextMatrix(i, COL_状态) = Nvl(rsTmp!医嘱状态, 0)

                .TextMatrix(i, COL_类别) = rsTmp!诊疗类别
                .TextMatrix(i, COL_诊疗项目ID) = rsTmp!诊疗项目ID
                .TextMatrix(i, COL_名称) = rsTmp!名称
                .TextMatrix(i, COL_标本部位) = Nvl(rsTmp!标本部位)
                .TextMatrix(i, COL_检查方法) = Nvl(rsTmp!检查方法)
                .TextMatrix(i, COL_执行标记) = Nvl(rsTmp!执行标记, 0)
                .TextMatrix(i, COL_收费细目ID) = Nvl(rsTmp!收费细目ID)
                .TextMatrix(i, col_医嘱内容) = Nvl(rsTmp!医嘱内容)
                .TextMatrix(i, COL_医生嘱托) = Nvl(rsTmp!医生嘱托)
                .Cell(flexcpData, i, COL_医生嘱托) = CStr(Nvl(rsTmp!摘要))

                .TextMatrix(i, COL_计价性质) = Nvl(rsTmp!计价特性, 0)
                .TextMatrix(i, COL_计算方式) = Nvl(rsTmp!计算方式, 0)
                .TextMatrix(i, COL_频率性质) = Nvl(rsTmp!执行频率, 0)
                .TextMatrix(i, COL_操作类型) = Nvl(rsTmp!操作类型)
                .TextMatrix(i, COL_单独应用) = Nvl(rsTmp!单独应用)
                .TextMatrix(i, COL_执行分类) = Nvl(rsTmp!执行分类, 0)
                .TextMatrix(i, COL_毒理分类) = Nvl(rsTmp!毒理分类)
                .TextMatrix(i, COL_抗菌等级) = Val("" & rsTmp!抗生素)
                .TextMatrix(i, COL_药品剂型) = Nvl(rsTmp!药品剂型)
                .TextMatrix(i, COL_处方限量) = Nvl(rsTmp!处方限量)
                .TextMatrix(i, COL_处方职务) = Nvl(rsTmp!处方职务)
                .TextMatrix(i, COL_基本药物) = rsTmp!基本药物 & ""
                If gblnKSSStrict Or gbln输血分级管理 Or gbln血库系统 Then
                    .TextMatrix(i, COL_审核状态) = Val("" & rsTmp!审核状态)
                End If
                .TextMatrix(i, COL_用药目的) = "" & rsTmp!用药目的
                .TextMatrix(i, COL_用药理由) = "" & rsTmp!用药理由
                .TextMatrix(i, COL_配方ID) = Nvl(rsTmp!配方ID)
                .TextMatrix(i, COL_临床自管药) = rsTmp!临床自管药 & ""
                .TextMatrix(i, COL_高危药品) = Nvl(rsTmp!高危药品, 0)
                .TextMatrix(i, COL_组合项目ID) = "" & rsTmp!组合项目ID
                If Format(Nvl(rsTmp!撤档时间, "3000/1/1"), "yyyy-MM-dd") <> "3000-01-01" Then
                    .TextMatrix(i, COL_是否停用) = 1
                End If
                If InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 Then
                    .TextMatrix(i, COL_剂量系数) = Nvl(rsTmp!剂量系数)
                    .TextMatrix(i, COL_门诊包装) = Nvl(rsTmp!门诊包装)
                    .TextMatrix(i, COL_门诊单位) = Nvl(rsTmp!门诊单位)
                    If Not IsNull(rsTmp!剂量系数) Then
                        .TextMatrix(i, COL_可否分零) = Nvl(rsTmp!可否分零, 0)
                    End If
                ElseIf .TextMatrix(i, COL_类别) = "4" Then
                    .TextMatrix(i, COL_剂量系数) = 1
                    .TextMatrix(i, COL_门诊包装) = 1
                    .TextMatrix(i, COL_门诊单位) = Nvl(rsTmp!散装单位)
                    .TextMatrix(i, COL_跟踪在用) = Nvl(rsTmp!跟踪在用, 0)
                End If

                .TextMatrix(i, COL_开始时间) = Format(rsTmp!开始执行时间, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_开始时间) = Format(rsTmp!开始执行时间, "yyyy-MM-dd HH:mm")

                .TextMatrix(i, COL_频率) = Nvl(rsTmp!执行频次)
                .TextMatrix(i, COL_频率次数) = Nvl(rsTmp!频率次数)
                .TextMatrix(i, COL_频率间隔) = Nvl(rsTmp!频率间隔)
                .TextMatrix(i, COL_间隔单位) = Nvl(rsTmp!间隔单位)
                .TextMatrix(i, COL_执行时间) = Nvl(rsTmp!执行时间方案)

                .TextMatrix(i, COL_执行科室ID) = Nvl(rsTmp!执行科室ID)
                .TextMatrix(i, COL_执行性质) = Nvl(rsTmp!执行性质, 0)
                .TextMatrix(i, COL_免试) = IIF(Nvl(rsTmp!皮试结果, "") = "免试", "1", "0")
                .Cell(flexcpData, i, COL_免试) = .TextMatrix(i, COL_免试)
                .TextMatrix(i, COL_是否溶媒) = Val(rsTmp!溶媒 & "")
                .TextMatrix(i, COL_申请序号) = rsTmp!申请序号 & ""
                .TextMatrix(i, COL_处方审查状态) = rsTmp!处方审查状态 & ""
                .TextMatrix(i, COL_处方审查结果) = rsTmp!处方审查结果 & ""
                .TextMatrix(i, COL_处方序号) = Val("" & rsTmp!处方序号)
                .TextMatrix(i, COL_危急值ID) = Val("" & rsTmp!危急值ID)
                
                If rsTmp!诊疗类别 = "E" Then
                    If Nvl(rsTmp!相关ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_相关ID)) = rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是成药的给药途径,可能是一并给药的
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = rsTmp!ID Then
                                    '显示给药途径
                                    .TextMatrix(j, COL_用法) = rsTmp!名称 & rsTmp!医生嘱托
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                            '当前记录是中药配方的用法,即配方显示行
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                            .TextMatrix(i, COL_中药形态) = Val("" & rsTmp!检查方法)    '中药用法行的检查方法字段存储了中药形态
                            bln配方 = True
                        ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                            .TextMatrix(i, COL_用法) = rsTmp!名称
                        End If
                    ElseIf Not IsNull(rsTmp!相关ID) And .TextMatrix(i - 1, COL_类别) = "K" And Nvl(rsTmp!相关ID, 0) = .RowData(i - 1) Then
                        '当前记录是输血途径行
                        .TextMatrix(i - 1, COL_用法) = rsTmp!名称
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        '当前记录是中药配方煎法行
                        bln配方 = True
                    End If
                ElseIf rsTmp!诊疗类别 = "7" Then
                    bln配方 = True
                End If

                '单量
                .TextMatrix(i, COL_单量) = FormatEx(Nvl(rsTmp!单次用量), 5)
                If .TextMatrix(i, COL_类别) = "4" Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!散装单位)
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 _
                       Or (Val(.TextMatrix(i, COL_频率性质)) = 0 And InStr(",1,2,", Nvl(rsTmp!计算方式, 0)) > 0) Then
                    .TextMatrix(i, COL_单量单位) = Nvl(rsTmp!计算单位)
                End If

                '天数
                .TextMatrix(i, COL_天数) = Nvl(rsTmp!天数)
                '取最近新开医嘱的开数作为缺省天数
                If InStr(",1,2,", Nvl(rsTmp!医嘱状态, 0)) > 0 _
                   And InStr(",5,6,", rsTmp!诊疗类别) > 0 And Nvl(rsTmp!天数, 0) <> 0 Then
                    msng天数 = Nvl(rsTmp!天数, 1)
                End If

                '总量
                If InStr(",5,6,", rsTmp!诊疗类别) > 0 Then
                    '成药临嘱有总量,以零售单位存放,门诊单位显示
                    If Not IsNull(rsTmp!总给予量) And Not IsNull(rsTmp!门诊包装) Then
                        .TextMatrix(i, COL_总量) = FormatEx(rsTmp!总给予量 / rsTmp!门诊包装, 5)
                    End If
                    .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!门诊单位)

                    Call Set用药天数是否超期(i)

                ElseIf bln配方 Then
                    If Not IsNull(rsTmp!总给予量) Then .TextMatrix(i, COL_总量) = rsTmp!总给予量
                    .TextMatrix(i, COL_总量单位) = "付"    '中药配方总量单位为"付"
                    Call Set用药天数是否超期(i)
                Else
                    '其它情况有中药和其它临嘱
                    If Not IsNull(rsTmp!总给予量) Then .TextMatrix(i, COL_总量) = rsTmp!总给予量

                    If .TextMatrix(i, COL_类别) = "4" Then
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!散装单位)
                    Else
                        .TextMatrix(i, COL_总量单位) = Nvl(rsTmp!计算单位)
                    End If
                End If

                .TextMatrix(i, COL_超量说明) = rsTmp!超量说明 & ""

                .TextMatrix(i, COL_开嘱科室ID) = rsTmp!开嘱科室id
                .TextMatrix(i, COL_开嘱医生) = rsTmp!开嘱医生

                .TextMatrix(i, COL_开嘱时间) = Format(rsTmp!开嘱时间, "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COL_开嘱时间) = Format(rsTmp!开嘱时间, "yyyy-MM-dd HH:mm")

                .TextMatrix(i, COL_零费记帐) = Val("" & rsTmp!零费记帐)

                '显示紧急标志:一并给药只显示在第一行
                .TextMatrix(i, COL_标志) = Nvl(rsTmp!紧急标志, 0)
                .TextMatrix(i, COL_签名否) = Nvl(rsTmp!签名否)
                Call SetRow标志图标(i, 1)


                '根据医嘱状态,药品毒理设置颜色
                '-------------------------------------------------------------------
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
                If rsTmp!医嘱状态 = 8 Then
                    '已停止(已发送)
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '深蓝
                End If

                '毒麻精药品标识:中药配方及组成味中药不处理
                If InStr(",5,6,", rsTmp!诊疗类别) > 0 And Not IsNull(rsTmp!毒理分类) Then
                    If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", rsTmp!毒理分类) > 0 Then
                        .Cell(flexcpFontBold, i, col_医嘱内容) = True
                    End If
                End If

                'Pass根据审查结果显示警示灯
                If mblnPass Then
                    If gobjPass.zlPassCheck(mobjPassMap) And Not IsNull(rsTmp!审查结果) Then
                        Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, rsTmp!审查结果)
                    End If
                End If

                Call Set医嘱超量(i, i)

                rsTmp.MoveNext
            Next

            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, col_医嘱内容, .Rows - 1, col_医嘱内容) = 0

            Call .AutoSize(col_医嘱内容)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
    Else
        mblnRowChange = False
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        mblnRowChange = True
    End If

    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check开始时间(ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的开始时间是否合法
'说明：
'1.开始时间不能小于病人的挂号时间
'2.正常录入时,开始时间不能小于当前时间之前30分钟(从而可能造成开嘱时间大于开始时间30分钟)
    If Not IsDate(strStart) Then
        MsgBox "输入的医嘱开始执行时间无效。", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Format(strStart, "yyyy-MM-dd HH:mm") < Format(mdat挂号时间, "yyyy-MM-dd HH:mm") Then
        strMsg = "医嘱的开始执行时间不能小于病人的挂号时间 " & Format(mdat挂号时间, "yyyy-MM-dd HH:mm") & " 。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
    Check开始时间 = True
End Function

Private Function Check安排时间(ByVal strDate As String, ByVal strStart As String, ByVal strType As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查输入的手术/输血时间是否合法
'说明：
'1.手术/输血时间不能小于医嘱的开始时间
    Dim strInDate As String, strDateType As String
    
    If strType = "F" Then
        strDateType = "手术"
    ElseIf strType = "K" Then
        strDateType = "输血"
    End If
    
    If Not IsDate(strDate) Then
        strMsg = "输入的" & strDateType & "时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    ElseIf IsDate(strStart) Then
        If Format(strDate, "yyyy-MM-dd HH:mm") < Format(strStart, "yyyy-MM-dd HH:mm") Then
            strMsg = strDateType & "时间不能小于医嘱开始时间。"
            If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Check安排时间 = True
End Function

Private Function Check开嘱时间(ByVal strDate As String, ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'功能：检查开嘱时间是否有效
'说明：不应小于病人挂号时间
    If Not IsDate(strDate) Then
        strMsg = "输入的开嘱时间无效。"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
        
'    If Format(strDate, "yyyy-MM-dd HH:mm") < Format(mdat挂号时间, "yyyy-MM-dd HH:mm") Then
'        strMsg = "开嘱时间不能小于病人的挂号时间 " & Format(mdat挂号时间, "yyyy-MM-dd HH:mm") & " 。"
'        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'        Exit Function
'    End If
    Check开嘱时间 = True
End Function

Private Function Check配伍禁忌(ByVal str药品IDs As String) As Boolean
'功能：检查西成药,中成药的配伍禁忌;中药配方不在这里检查
'参数：str药品IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr慎用 As Variant, arr禁用 As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng项目id As Long, str名称 As String, bln未编辑 As Boolean
    Dim lng组编号 As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr慎用 = Array(): arr禁用 = Array()
        
    strSQL = "Select /*+ rule*/ 组编号 From 诊疗互斥项目" & _
        " Where 项目ID IN(Select Column_Value From Table(f_Num2list([1]))) Group by 组编号 Having Count(*)>1"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str药品IDs)
    For k = 1 To rsMain.RecordCount
        strSQL = "Select /*+ RULE */ A.组编号,A.类型,A.项目ID,B.名称" & _
            " From 诊疗互斥项目 A,诊疗项目目录 B" & _
            " Where A.项目ID=B.ID And A.组编号=[2]" & _
            " And A.项目ID IN(Select Column_Value From Table(f_Num2list([1])))" & _
            " Order by A.组编号,B.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str药品IDs, Val(rsMain!组编号))
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    ReDim Preserve arr慎用(UBound(arr慎用) + 1)
                Else
                    ReDim Preserve arr禁用(UBound(arr禁用) + 1)
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                arr慎用(UBound(arr慎用)) = arr慎用(UBound(arr慎用)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            Else
                arr禁用(UBound(arr禁用)) = arr禁用(UBound(arr禁用)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    
    '先检查禁用部份(禁止继续)
    If UBound(arr禁用) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr禁用) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(Mid(arr禁用(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) '每项目
                lng项目id = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "，" & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目id), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & "● " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下药品互相禁用：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查慎用部份(提醒是否继续)
    If UBound(arr慎用) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr慎用) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(Mid(arr慎用(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) '每项目
                lng项目id = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "，" & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目id), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & "● " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("在病人医嘱中发现以下药品互相慎用：" & strMsg & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check配伍禁忌 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check诊疗互斥(ByVal str诊疗IDs As String) As Boolean
'功能：检查非药品(成药,中药)的互斥
'参数：str诊疗IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr提醒 As Variant, arr禁止 As Variant, arr停止 As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng项目id As Long, str名称 As String, bln未编辑 As Boolean
    Dim lng组编号 As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr提醒 = Array(): arr禁止 = Array(): arr停止 = Array()
    
    strSQL = "Select /*+ rule*/ 组编号 From 诊疗互斥项目" & _
        " Where 项目ID IN(Select Column_Value From Table(f_Num2list([1]))) Group by 组编号 Having Count(*)>1"
    Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊疗IDs)
    For k = 1 To rsMain.RecordCount
        strSQL = "Select /*+ RULE */ A.组编号,A.组名称,A.类型,A.项目ID,B.名称" & _
            " From 诊疗互斥项目 A,诊疗项目目录 B" & _
            " Where A.项目ID=B.ID And A.组编号=[2]" & _
            " And A.项目ID IN(Select Column_Value From Table(f_Num2list([1])))" & _
            " Order by A.组编号,B.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊疗IDs, Val(rsMain!组编号))
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    ReDim Preserve arr提醒(UBound(arr提醒) + 1)
                    arr提醒(UBound(arr提醒)) = rsTmp!组名称
                ElseIf rsTmp!类型 = 2 Then
                    ReDim Preserve arr禁止(UBound(arr禁止) + 1)
                    arr禁止(UBound(arr禁止)) = rsTmp!组名称
                Else
                    ReDim Preserve arr停止(UBound(arr停止) + 1)
                    arr停止(UBound(arr停止)) = rsTmp!组名称
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                arr提醒(UBound(arr提醒)) = arr提醒(UBound(arr提醒)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            ElseIf rsTmp!类型 = 2 Then
                arr禁止(UBound(arr禁止)) = arr禁止(UBound(arr禁止)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            Else
                arr停止(UBound(arr停止)) = arr停止(UBound(arr停止)) & Chr(234) & rsTmp!项目ID & Chr(8) & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    '先检查禁止继续部份
    If UBound(arr禁止) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr禁止) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(arr禁止(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目id = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目id), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) = 1 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下内容互相排斥：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查自动停止部份,门诊处理为禁止(临嘱)
    If UBound(arr停止) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr停止) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(arr停止(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目id = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目id), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False ': Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "在病人医嘱中发现以下内容互相排斥：" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '再检查提醒是否继续部份
    If UBound(arr提醒) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr提醒) '每组
            strTmp = "": bln未编辑 = True
            arrItems = Split(arr提醒(i), Chr(234))
            For j = 1 To UBound(arrItems) '每项目
                lng项目id = Split(arrItems(j), Chr(8))(0)
                str名称 = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str名称
                
                '为了定位,在医嘱中查找本次新增或修改的该项目(可能有多个)所在行
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng项目id), lngRow + 1, COL_诊疗项目ID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '编辑过的最小行优先定位
                        bln未编辑 = False: Exit Do
                    End If
                Loop
            Next
            If Not bln未编辑 Then '如果一组中的项目在本次都未编辑过,则不管
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "：" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = col_医嘱内容: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("在病人医嘱中发现以下内容互相排斥：" & strMsg & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check诊疗互斥 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lngRow As Long) As String
'功能：检查指定药品行的库存情况
'返回：空=表示通过
    Dim dbl总量 As Double, strMsg As String
    Dim lng执行科室ID As Long, i As Integer
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Or .TextMatrix(lngRow, COL_类别) = "4" And Val(.TextMatrix(lngRow, COL_跟踪在用)) = 1 Then
            If TheStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID)), .TextMatrix(lngRow, COL_类别)) <> 0 Then
                If .TextMatrix(lngRow, COL_库存) <> "" Then
                    '成药临嘱直接检查总量
                    dbl总量 = Val(.TextMatrix(lngRow, COL_总量))
                    If dbl总量 > 0 Then
                        If dbl总量 > Val(.TextMatrix(lngRow, COL_库存)) Then
                            strMsg = """" & .TextMatrix(lngRow, col_医嘱内容) & """库存提醒：" & _
                                vbCrLf & vbCrLf & Get部门名称(Val(.TextMatrix(lngRow, COL_执行科室ID))) & _
                                IIF(InStr(GetInsidePrivs(p门诊医嘱下达), "显示药品库存") = 0, _
                                "当前可用库存不足 " & FormatEx(dbl总量, 5) & .TextMatrix(lngRow, COL_门诊单位) & "。", _
                                "当前可用库存为 " & FormatEx(Val(.TextMatrix(lngRow, COL_库存)), 5) & .TextMatrix(lngRow, COL_门诊单位) & "，不足 " & FormatEx(dbl总量, 5) & .TextMatrix(lngRow, COL_门诊单位) & "。")
                        End If
                    End If
                End If
            End If
        ElseIf RowIn配方行(lngRow) And Val(.TextMatrix(lngRow, COL_总量)) <> 0 Then
            '根据付数计算总量
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" And .TextMatrix(i, COL_库存) <> "" Then
                        '总量=门诊包装(单味剂量*付数)
                        '中药药房单位按不可分零处理:每付
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            dbl总量 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_门诊包装))
                        Else
                            dbl总量 = Val(.TextMatrix(i, COL_总量)) * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_门诊包装)))
                        End If
                        If dbl总量 > Val(.TextMatrix(i, COL_库存)) Then
                            lng执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                            If TheStockCheck(lng执行科室ID, .TextMatrix(i, COL_类别)) = 0 Then Exit For
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, col_医嘱内容) & _
                                "：所需总量 " & FormatEx(dbl总量, 5) & .TextMatrix(i, COL_门诊单位) & _
                                "，可用库存" & IIF(InStr(GetInsidePrivs(p门诊医嘱下达), "显示药品库存") = 0, _
                                    "不足", " " & FormatEx(Val(.TextMatrix(i, COL_库存)), 5) & .TextMatrix(i, COL_门诊单位))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            If strMsg <> "" Then
                strMsg = "中药配方库存提醒，" & Get部门名称(lng执行科室ID) & "中以下味药库存不足：" & vbCrLf & strMsg
            End If
        End If
    End With
    CheckStock = strMsg
End Function

Private Function CheckMoney() As Boolean
'功能：费用报警检查
'说明：病区有累计费用报警方式时,只提醒。
    Dim rsTmp As New ADODB.Recordset
    Dim str适用病人 As String, strSQL As String
    Dim cur预交 As Currency, cur余额 As Currency
    Dim cur担保额 As Currency
    
    On Error GoTo errH
    '费用余额
    strSQL = "Select 预交余额,Nvl(预交余额,0)-Nvl(费用余额,0) as 余额 From 病人余额 Where 性质=1 And 类型 = 1 And 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If Not rsTmp.EOF Then
        cur预交 = Nvl(rsTmp!预交余额, 0)
        cur余额 = Nvl(rsTmp!余额, 0)
    End If
    
    '担保额
    strSQL = "Select 担保额 From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If Not rsTmp.EOF Then cur担保额 = Nvl(rsTmp!担保额, 0)
    
    '有预交款的病人才判断
    If cur预交 <> 0 Then
        '是否医保
        strSQL = "Select zl_PatiWarnScheme([1]) as 适用病人 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        If Not rsTmp.EOF Then str适用病人 = Nvl(rsTmp!适用病人)
            
        '报警值:NULL与0当作不同意义处理
        strSQL = "Select 报警值 From 记帐报警线 Where 报警方法=1 And Nvl(病区ID,0)=0 And 报警值 is Not NULL And 适用病人=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str适用病人)
        If Not rsTmp.EOF Then
            If cur余额 + cur担保额 < Nvl(rsTmp!报警值, 0) Then
                If MsgBox("病人当前可用剩余款 " & FormatEx(cur余额 + cur担保额, 2) & IIF(cur担保额 <> 0, "(含担保额:" & FormatEx(cur担保额, 2) & ")", "") & " 低于报警值 " & FormatEx(Nvl(rsTmp!报警值, 0), 2) & "，要继续吗？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    CheckMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
'功能：获取组ID相同的一组医嘱行号范围(注意考虑一并给药中的空行)
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_相关ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_相关ID)))
            If Not (.RowData(i) = 0 And i >= .FixedRows) Then '跳过空行
                If lngO组ID = lngS组ID Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function CheckAdvice() As Boolean
'功能：检查当前病人(婴儿)的医嘱输入是否合法
'说明：如果有不合法的地方，在本函数中提示及定位
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    Dim dbl总量 As Double, strMsg As String
    Dim str药品IDs As String, str诊疗IDs As String
    Dim lngCount As Long, lngRow As Long
    Dim blnSkipStock As Boolean, blnSkipTotal As Boolean
    Dim vMsg As VbMsgBoxResult, sng天数 As Single
    Dim blnValid As Boolean, lngRxCount As Long
    Dim blnAppend As Boolean, i As Long, j As Long, k As Long
    Dim str疾病IDs As String, str诊断IDs As String
    Dim str中药名 As String, strExtra As String, lng主ID As Long
    Dim lng麻醉诊疗ID As Long, lng给药执行性质 As Long, str部位方法 As String
    Dim lngBegin As Long, lngEnd As Long
    Dim blnExists As Boolean
    Dim dblOneDay As Double
    Dim datCur As Date
    Dim blnNo As Boolean, blnCheck超量 As Boolean, blnOut As Boolean
    Dim strIDs1 As String, strIDs2 As String, str医嘱内容 As String
    Dim lngSame As Long
    
    On Error GoTo errH
    If vsAdvice.Enabled = True Then
        If Me.ActiveControl.Name = "txt总量" Then
            Call txt总量_Validate(False)
        ElseIf Me.ActiveControl.Name = "txt单量" Then
            Call txt单量_Validate(False)
        ElseIf Me.ActiveControl.Name = "txt天数" Then
            Call txt天数_Validate(False)
        End If
        vsAdvice.SetFocus  '处理光标定位，如果光标还没离开输入框，则还未更新数据。
    End If
    If Not CheckApply Then Exit Function
    datCur = zlDatabase.Currentdate
    '诊断的检查
    '-----------------------------------------------------------------------------------------
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                If mint险类 = 920 Then '北京医保无理要求
                    If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 82 Then
                        .Row = i: .Col = col诊断
                        MsgBox "诊断内容太长，只允许82个字符或41个汉字。", vbInformation, gstrSysName
                        vsDiag.SetFocus: Exit Function
                    End If
                End If
                If zlCommFun.ActualLen(.TextMatrix(i, col诊断)) > 200 Then
                    .Row = i: .Col = col诊断
                    MsgBox "诊断内容太长，只允许200个字符或100个汉字。", vbInformation, gstrSysName
                    vsDiag.SetFocus: Exit Function
                End If
                If .TextMatrix(i, col发病时间) <> "" Then
                    If Format(datCur, "YYYY-MM-DD HH:mm") < Format(.TextMatrix(i, col发病时间), "YYYY-MM-DD HH:mm") Then
                         .Row = i: .Col = col发病时间
                        MsgBox "发病时间应该早于当前时间。", vbInformation, gstrSysName
                        vsDiag.SetFocus: Exit Function
                    End If
                End If
                lngSame = 0
                For j = i + 1 To .Rows - 1
                    If Trim(.TextMatrix(j, col诊断)) <> "" And .Cell(flexcpData, i, col中医) = .Cell(flexcpData, j, col中医) Then '同类型诊断
                        If .TextMatrix(j, col诊断) & "|" & .TextMatrix(j, col中医证候) = .TextMatrix(i, col诊断) & "|" & .TextMatrix(i, col中医证候) Then
                            .Row = i: .Col = col诊断
                            MsgBox "发现存在两行相同的诊断信息。", vbInformation, gstrSysName
                            vsDiag.SetFocus: Exit Function
                        ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                            If Val(.TextMatrix(j, col疾病ID)) & "|" & .TextMatrix(j, col中医证候) = Val(.TextMatrix(i, col疾病ID)) & "|" & .TextMatrix(i, col中医证候) Then
                                .Row = i: .Col = col诊断
                                MsgBox "发现存在两行相同的疾病信息。", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        ElseIf Val(.TextMatrix(i, col诊断ID)) <> 0 And .Cell(flexcpData, i, col中医) = 0 Then
                            '因中医诊断带证候,可能无对应证候ID,诊断ID又相同
                            If Val(.TextMatrix(j, col诊断ID)) = Val(.TextMatrix(i, col诊断ID)) Then
                                .Row = i: .Col = col诊断
                                MsgBox "发现存在两行相同的诊断信息。", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        End If
                        If .Cell(flexcpData, i, col中医) <> 0 Then '中医诊断
                             If .TextMatrix(j, col诊断) = .TextMatrix(i, col诊断) Then
                                lngSame = lngSame + 1
                             ElseIf Val(.TextMatrix(i, col疾病ID)) <> 0 Then
                                If Val(.TextMatrix(j, col疾病ID)) = Val(.TextMatrix(i, col疾病ID)) Then
                                    lngSame = lngSame + 1
                                End If
                             End If
                            If lngSame >= 2 Then
                                .Row = i: .Col = col诊断
                                MsgBox "存在两条以上的诊断相同且证候不同的诊断，诊断不明确。", vbInformation, gstrSysName
                                vsDiag.SetFocus: Exit Function
                            End If
                        End If
                    End If
                Next
                If Val(.TextMatrix(i, col疾病ID)) <> 0 Then str疾病IDs = str疾病IDs & "," & Val(.TextMatrix(i, col疾病ID))
                If Val(.TextMatrix(i, col诊断ID)) <> 0 Then str诊断IDs = str诊断IDs & "," & Val(.TextMatrix(i, col诊断ID))
            End If
        Next
    End With
        
    '-----------------------------------------------------
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            bln配方行 = False: bln检验行 = False
            '本次新增或修改药品行的处方职务检查
            If .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                strMsg = CheckOneDuty(.TextMatrix(i, col_医嘱内容), .TextMatrix(i, COL_处方职务), .TextMatrix(i, COL_开嘱医生), InStr(",1,2,", mstr付款码) > 0 And mstr付款码 <> "")
                If strMsg <> "" Then
                    .Col = col_医嘱内容
                    If .TextMatrix(i, COL_类别) = "7" Then
                        lngRow = .FindRow(CLng(.TextMatrix(i, COL_相关ID)), i + 1)
                        If lngRow <> -1 Then .Row = lngRow
                    Else
                        .Row = i
                    End If
                    Call .ShowCell(.Row, .Col)
                    MsgBox strMsg, vbInformation, gstrSysName
                    .Refresh
                    Call vsAdvice_KeyPress(13)
                    Exit Function
                End If
                
                '抗菌用药检查
                If gblnKSSStrict Then
                    If Val(.TextMatrix(i, COL_抗菌等级)) > 0 Then
                        If Val(.TextMatrix(i, COL_用药目的)) = 0 Then
                            strMsg = ",抗菌用药要求登记用药目的。"
                            .Col = COL_用药目的: Exit For
                        End If
                        
                        If Val(.TextMatrix(i, COL_抗菌等级)) = 3 And mbytPatiType <> 2 Then
                            strMsg = ",门诊非急诊挂号不能使用特殊使用级的抗菌药物。"
                            .Col = col_医嘱内容: Exit For
                        End If
                        
                        '如果操作员没有抗菌用药处方权，则禁止保存
                        If UserInfo.用药级别 = 0 Then
                            strMsg = ",您没有抗菌用药权限，请联系管理员。"
                            .Col = col_医嘱内容: Exit For
                        End If
                        
                        If Val(.TextMatrix(i, COL_EDIT)) = 2 Then
                            If UserInfo.用药级别 < Val(.TextMatrix(i, COL_抗菌等级)) And Val(.TextMatrix(i, COL_标志)) <> 1 Then
                                .TextMatrix(i, COL_审核状态) = 1
                            End If
                        End If
                        
                        '一组药品中，只要有一个为待审核或未审核通过，则整组更改(没在输入时处理，是因为可能给药途径是后续才加上的，处理的点较多)
                        If Val(.TextMatrix(i, COL_审核状态)) > 0 Then
                            Call GetRowScope(i, lngBegin, lngEnd)
                            For j = lngBegin To lngEnd
                                If j <> i Then
                                    .TextMatrix(j, COL_审核状态) = .TextMatrix(i, COL_审核状态)
                                End If
                            Next
                        End If
                        
                        
                        '紧急医嘱检查
                        If Val(.TextMatrix(i, COL_标志)) = 1 Then
                            If .TextMatrix(i, COL_用药理由) = "" Then
                                strMsg = ",紧急使用的抗菌用药要求输入用药理由。"
                                .Col = COL_用药理由: Exit For
                            End If
                            If Val(.TextMatrix(i, COL_天数)) > 1 Then
                                strMsg = ",紧急使用的抗菌用药要求仅限一天。"
                                .Col = COL_天数: Exit For
                            ElseIf Val(.TextMatrix(i, COL_天数)) = 0 Then
                                '未输入天数时进行反算天数：总量/单量/频率
                                If .TextMatrix(i, COL_剂量系数) <> "" And .TextMatrix(i, COL_门诊包装) <> "" Then
                                    '计算出一天的总量
                                    dblOneDay = FormatEx(Calc缺省药品总量( _
                                    Val(.TextMatrix(i, COL_单量)), 1, _
                                    Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), _
                                    .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), _
                                    Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_门诊包装)), _
                                    Val(.TextMatrix(i, COL_可否分零))), 5)
                                    If Val(.TextMatrix(i, COL_总量)) > dblOneDay Then
                                        strMsg = ",紧急使用的抗菌用药总量超出了一天的使用量：" & dblOneDay & "。"
                                        .Col = COL_天数: Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            
            '调用自定义的医嘱检查函数
            If .RowData(i) <> 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                
                strExtra = CStr(.Cell(flexcpData, i, COL_医生嘱托))
                lng麻醉诊疗ID = 0
                lng给药执行性质 = 0
                str部位方法 = ""
                If .TextMatrix(i, COL_类别) = "F" Then
                    lng主ID = Val(.TextMatrix(i, COL_相关ID))
                    If lng主ID = 0 Then lng主ID = .RowData(i)
                    For k = i + 1 To .Rows - 1
                        If Val(.TextMatrix(k, COL_相关ID)) <> lng主ID Then Exit For
                        If .TextMatrix(k, COL_类别) = "G" Then
                            lng麻醉诊疗ID = .TextMatrix(k, COL_诊疗项目ID)
                        End If
                    Next
                ElseIf InStr("4,5,6,7", .TextMatrix(i, COL_类别)) > 0 Then
                    k = .FindRow(CLng(Val(.TextMatrix(i, COL_相关ID))), i + 1)
                    If k > 0 Then lng给药执行性质 = Val(.TextMatrix(k, COL_执行性质))
                
                ElseIf .TextMatrix(i, COL_类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    lng主ID = .RowData(i)
                    For k = i + 1 To .Rows - 1
                        If Val(.TextMatrix(k, COL_相关ID)) <> lng主ID Then Exit For
                        str部位方法 = str部位方法 & "," & .TextMatrix(k, COL_标本部位) & ":" & .TextMatrix(k, COL_检查方法)
                    Next
                    str部位方法 = Mid(str部位方法, 2)
                End If
                
                strExtra = strExtra & "||" & lng给药执行性质 & "||" & lng麻醉诊疗ID & "||" & IIF(str部位方法 = "", " ", str部位方法) & "||" & Val(.TextMatrix(i, COL_收费细目ID))
                
                strSQL = "Select zl_AdviceCheck([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14]) as 结果 From Dual"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceCheck", 1, mlng病人ID, mlng挂号ID, mint险类, 1, _
                     .TextMatrix(i, COL_类别), Val(.TextMatrix(i, COL_诊疗项目ID)), _
                    Val(.TextMatrix(i, COL_开嘱科室ID)), CStr(.TextMatrix(i, COL_开嘱医生)), _
                    Val(.TextMatrix(i, COL_执行科室ID)), Val(.TextMatrix(i, COL_执行性质)), Val(.TextMatrix(i, COL_执行标记)), _
                    Val(.TextMatrix(i, COL_单量)), strExtra)
                
                If Not rsTmp.EOF Then
                    strMsg = Nvl(rsTmp!结果)
                    If strMsg <> "" Then
                        Select Case Val(Split(strMsg, "|")(0))
                        Case 1 '提示
                            If MsgBox(Split(strMsg, "|")(1), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                strMsg = "": Exit For
                            End If
                        Case 2 '禁止
                            MsgBox Split(strMsg, "|")(1), vbInformation, gstrSysName
                            strMsg = "": Exit For
                        End Select
                        strMsg = ""
                    End If
                End If
            End If
            
            '其它输入合法性检查
            If .RowData(i) <> 0 Then
            
                '有性别限制项目的检查（只针对：检验、检查、手术）
                If InStr(",C,D,F,5,6,7,", .TextMatrix(i, COL_类别)) > 0 And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 And (mstr性别 = "男" Or mstr性别 = "女") And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                    If .TextMatrix(i, COL_类别) = "F" Or .TextMatrix(i, COL_类别) = "D" And Not .RowHidden(i) Or .TextMatrix(i, COL_类别) = "C" And .RowHidden(i) Or InStr(",5,6,7,", .TextMatrix(i, COL_类别)) > 0 Then
                        strSQL = "Select Decode(a.适用性别, 1, '男', 2, '女', '未知') As 性别 From 诊疗项目目录 A Where a.Id = [1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_诊疗项目ID)))
                        
                        If rsTmp!性别 <> mstr性别 And rsTmp!性别 <> "未知" Then
                            If .TextMatrix(i, COL_类别) = "D" Then
                                strMsg = "手术，对性别""" & mstr性别 & """不适用。"
                            Else
                                strMsg = Decode(.TextMatrix(i, COL_类别), "C", "检验项目", "F", "附加手术", "D", "检查项目", "药品") & "，对性别""" & mstr性别 & """不适用。"
                            End If
                            .Col = col_医嘱内容: Exit For
                        End If
                    End If
                End If
                
                If .RowHidden(i) Then
                    
                    '本次新增或修改的行
                    '---------------------------------------------------
                    If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        If Not Check单独应用(i, strMsg) Then Exit For
                        '检查执行科室
                        If Val(.TextMatrix(i, COL_执行科室ID)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                            strMsg = "没有确定执行科室。"
                            .Col = COL_执行科室ID: Exit For
                        End If
                    
                        If .TextMatrix(i, COL_类别) = "D" And .TextMatrix(i, COL_标本部位) <> "" Then
                            If Check检查部位Enable(.TextMatrix(i, COL_诊疗项目ID), .TextMatrix(i, COL_标本部位), mstr性别, .TextMatrix(i, COL_检查方法), blnExists) = False Then
                                If blnExists = True Then
                                    strMsg = "中的部位：" & .TextMatrix(i, COL_标本部位) & "，对性别""" & mstr性别 & """不适用。"
                                Else
                                    Call cmdExt_Click
                                End If
                                .Col = col_医嘱内容: Exit For
                            End If
                        End If
                    End If
                Else
                    bln配方行 = RowIn配方行(i)
                    bln检验行 = RowIn检验行(i)
                    lngRow = i
                    If bln配方行 Then '得到配方的第一药品行
                        lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                        '自备药不检查药房
                        If Not (Val(.TextMatrix(lngRow, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5) And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                            
                            str中药名 = ""
                            If Check中药存储库房(lngRow, i, str中药名, vsAdvice, 1, mlng病人科室id, COL_类别, col_医嘱内容, COL_收费细目ID, COL_执行科室ID) = False Then
                                strMsg = "中的[" & str中药名 & "]没有存储在当前选择的药房，或该药房不是服务于当前病人科室的，不能使用该药品。"
                                Exit For
                            End If
                        End If
                    ElseIf bln检验行 Then '得到检验医嘱行
                        lngRow = .FindRow(CStr(.RowData(i)), , COL_相关ID)
                    End If
                    
                    '未发送的医嘱行
                    '------------------------------------
                    If Val(.TextMatrix(i, COL_状态)) = 1 Then
                        lngCount = lngCount + 1
                        
                        '必须录入单量:临嘱:成药或可选择频率的计时,计量项目可以录入(也可不录)
                        If (Val(.TextMatrix(lngRow, COL_频率性质)) = 0 And InStr(",1,2,", Val(.TextMatrix(lngRow, COL_计算方式))) > 0) _
                            Or InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                            '药品必须录入单量
                            If .TextMatrix(lngRow, COL_单量) <> "" Or mbln单量 And InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Then
                                If Not IsNumeric(.TextMatrix(lngRow, COL_单量)) Or Val(.TextMatrix(lngRow, COL_单量)) <= 0 Then
                                    strMsg = "没有录入正确的单次用量。"
                                    .Col = COL_单量: Exit For
                                End If
                            End If
                        End If
                        
                        '必须录入天数：对临嘱药品，如果指定了要录入
                        If mbln天数 And InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_天数)) <= 0 Then
                                strMsg = "请录入正确的用药天数。"
                                .Col = COL_天数: Exit For
                            End If
                        End If
                        
                        '必须录入总量:配方,临嘱(药品或其它)
                        If Not IsNumeric(.TextMatrix(i, COL_总量)) Or Val(.TextMatrix(i, COL_总量)) <= 0 Then
                            '输血医嘱允许不输入总量
                            If Not (.TextMatrix(i, COL_类别) = "K" And .TextMatrix(i, COL_总量) = "") Then
                                If bln配方行 Then
                                    strMsg = "没有录入正确的中药配方付数。"
                                ElseIf InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                                    strMsg = "没有录入正确的药品总给予量。"
                                Else
                                    strMsg = "没有录入正确的总量。"
                                End If
                                .Col = COL_总量: Exit For
                            End If
                        End If
                                            
                        '必须录入频率:临嘱也要检查,用于指导使用,可以不录入执行时间
                        If Val(.TextMatrix(lngRow, COL_频率性质)) = 0 Or InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 Or bln配方行 Then
                            If .TextMatrix(lngRow, COL_频率) = "" Then
                                strMsg = "没有确定执行频率。"
                                .Col = COL_频率: Exit For
                            End If
                            
                            '执行时间判断:可选频率的必须输入(对临嘱将来可能允许不录入,要注意发送等地方的处理)
                            If .TextMatrix(lngRow, COL_执行时间) = "" And .TextMatrix(lngRow, COL_间隔单位) <> "分钟" And .TextMatrix(lngRow, COL_频率) <> "必要时" And .TextMatrix(lngRow, COL_频率) <> "需要时" Then
                                If Not bln检验行 Then '检验组合显示行的采集方法为可选频率,但检验项目为一次性
                                    If Val(.TextMatrix(lngRow, COL_频率性质)) <> 1 Then
                                        strMsg = "没有录入执行时间方案。"
                                        .Col = COL_执行时间: Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                        '必须录入执行科室:非叮嘱和院外执行时(配方以药品行进行判断)
                        If Val(.TextMatrix(lngRow, COL_执行科室ID)) = 0 Then
                            If .TextMatrix(lngRow, COL_类别) = "Z" And InStr(",1,2,", Val(.TextMatrix(lngRow, COL_操作类型))) > 0 Then
                                If Val(.TextMatrix(lngRow, COL_操作类型)) = 1 Then
                                    strMsg = "没有确定留观医嘱的留观科室。"
                                ElseIf Val(.TextMatrix(lngRow, COL_操作类型)) = 2 Then
                                    strMsg = "没有确定住院医嘱的住院科室。"
                                End If
                                .Col = COL_执行科室ID: Exit For
                            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质))) = 0 Then
                                strMsg = "没有确定执行科室。"
                                .Col = COL_执行科室ID: Exit For
                            End If
                        End If
                        If lngRow <> i And Val(.TextMatrix(i, COL_执行科室ID)) = 0 Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质))) = 0 Then
                                strMsg = "没有确定执行科室。"
                                .Col = COL_执行科室ID: Exit For
                            End If
                        End If
                        
                        '开嘱时间判断
                        If Not Check开嘱时间(.Cell(flexcpData, i, COL_开嘱时间), .Cell(flexcpData, i, COL_开始时间), False, strMsg) Then
                            .Col = COL_开嘱时间: Exit For
                        End If
                        
                        '处方条数限制检查
                        If gintRXCount > 0 And InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 _
                            And Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            lngRxCount = GetMergeCount(vsAdvice, i, COL_相关ID, COL_收费细目ID)
                            If lngRxCount > gintRXCount Then
                                strMsg = "一并给药的药品种数 " & lngRxCount & " 种已达到或超过药品处方最多允许的种数 " & gintRXCount & " 种。"
                                .Col = col_医嘱内容: Exit For
                            End If
                        End If
                    End If
                    
                    '本次新增或修改的行
                    '---------------------------------------------------
                    If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        If Not Check单独应用(i, strMsg) Then Exit For
                        '开始时间判断:只对新增的医嘱进行判断,因为否则是不准修改开始时间的(不好判断被修改的医嘱开始时间的相对有效性)
                        If .TextMatrix(i, COL_EDIT) = "1" Then
                            If Not Check开始时间(.Cell(flexcpData, i, COL_开始时间), False, strMsg) Then
                                .Col = COL_开始时间: Exit For
                            End If
                        End If
                        '手术/输血医嘱的手术/输血时间判断
                        If .TextMatrix(i, COL_类别) = "F" Or .TextMatrix(i, COL_类别) = "K" Then
                            If Not Check安排时间(.TextMatrix(i, COL_手术时间), .Cell(flexcpData, i, COL_开始时间), .TextMatrix(.Row, COL_类别), False, strMsg) Then
                                .Col = COL_手术时间: Exit For
                            End If
                            If .TextMatrix(i, COL_类别) = "K" Then
                                '只有紧急医嘱保存输血原因
                                If .TextMatrix(i, COL_用药理由) <> "" And .TextMatrix(i, COL_标志) <> "1" Then
                                    .TextMatrix(i, COL_用药理由) = ""
                                Else
                                    If gbln输血分级管理 And .TextMatrix(i, COL_标志) = "1" And .TextMatrix(i, COL_用药理由) = "" Then
                                        strMsg = "启用了输血分级管理后，紧急输血医嘱必须填写输血原因。"
                                        .Col = COL_用药理由: Exit For
                                    End If
                                End If
                            End If
                        End If
                        
                        '给药途径，中药用法，采集方法设置检查
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_诊疗项目ID)) = 0 Then
                                strMsg = "没有设置对应的给药途径。"
                                .Col = COL_用法: Exit For
                            End If
                        End If
                        If .TextMatrix(i, COL_类别) = "E" And Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
                            If .RowData(i) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                If InStr(",7,E,", .TextMatrix(i - 1, COL_类别)) > 0 Then
                                    strMsg = "中药配方没有设置对应的用法。"
                                ElseIf .TextMatrix(i - 1, COL_类别) = "C" Then
                                    strMsg = "没有设置对应的标本采集方法。"
                                End If
                                .Col = COL_用法: Exit For
                            End If
                        End If
                                                
                        '最少总量检查:至少要满足一个频次周期的用量
                        If InStr(",4,5,6,", .TextMatrix(i, COL_类别)) > 0 Or bln配方行 Then
                            If Not blnSkipTotal And .TextMatrix(i, COL_频率) <> "" And Val(.TextMatrix(i, COL_频率次数)) <> 0 And Val(.TextMatrix(i, COL_频率间隔)) <> 0 Then
                                strMsg = ""
                                If bln配方行 Then '判断
                                    dbl总量 = Calc缺省药品总量(1, 1, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                                    If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                        strMsg = .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & _
                                            "在按""" & .TextMatrix(i, COL_频率) & """执行时,至少需要 " & dbl总量 & "付。"
                                    End If
                                ElseIf Val(.TextMatrix(i, COL_剂量系数)) <> 0 And Val(.TextMatrix(i, COL_单量)) <> 0 Then
                                    sng天数 = Val(.TextMatrix(i, COL_天数))
                                    If sng天数 = 0 Then sng天数 = 1
                                    dbl总量 = Calc缺省药品总量(Val(.TextMatrix(i, COL_单量)), sng天数, Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位), .TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_剂量系数)), Val(.TextMatrix(i, COL_门诊包装)), Val(.TextMatrix(i, COL_可否分零)))
                                    If Val(.TextMatrix(i, COL_总量)) < dbl总量 Then
                                        strMsg = .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & _
                                            "在按每次 " & .TextMatrix(i, COL_单量) & .TextMatrix(i, COL_单量单位) & "," & _
                                            .TextMatrix(i, COL_频率) & IIF(mbln天数 And .TextMatrix(i, COL_类别) <> "4", ",用药 " & sng天数 & " 天", "") & _
                                            "执行时,至少需要 " & dbl总量 & .TextMatrix(i, COL_总量单位) & "。"
                                    End If
                                End If
                                If strMsg <> "" And False Then '提示
                                    .Row = i: .Col = COL_总量: Call .ShowCell(.Row, .Col)
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^要继续吗？", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then
                                        If txt总量.Enabled And txt总量.Visible Then txt总量.SetFocus
                                        Exit Function
                                    ElseIf vMsg = vbIgnore Then
                                        blnSkipTotal = True
                                    End If
                                End If
                            End If
                        End If
                            
                            '检查药品是否超量和超期，遍历时应从第一条没有填写超量说明的行开始
                            If (gbyt超量原因 = 1 And InStr(gstr不录超量科室, "," & mlng病人科室id & ",") = 0) And Not blnCheck超量 _
                                And (.TextMatrix(i, COL_是否超量) = "1" Or .TextMatrix(i, COL_是否超期) = "1") And .TextMatrix(i, COL_超量说明) = "" Then
                                blnCheck超量 = SetAll超量说明(i, blnOut) '超量的检查只要这个方法执行了一次就算是所有都已经检查了不用再遍历了
                                If blnOut Then
                                    strMsg = "Not Null" '这里应该赋值确保 strMsg 不为空即可
                                    .Col = COL_超量说明: Exit For
                                End If
                            End If
                        
                        
                        '药品库存检查:只提醒,所以也只对本次编辑的才判断
                        If (InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Or bln配方行 _
                            Or .TextMatrix(i, COL_类别) = "4" And Val(.TextMatrix(i, COL_跟踪在用)) = 1) And Not blnSkipStock Then
                            strMsg = CheckStock(i)
                            If strMsg <> "" Then
                                .Row = i: .Col = col_医嘱内容: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^要继续吗？", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipStock = True
                                End If
                            End If
                        End If
                        
                        '执行时间合法性检查
                        If .TextMatrix(i, COL_执行时间) <> "" And .TextMatrix(i, COL_频率) <> "" Then
                            blnValid = ExeTimeValid(.TextMatrix(i, COL_执行时间), Val(.TextMatrix(i, COL_频率次数)), Val(.TextMatrix(i, COL_频率间隔)), .TextMatrix(i, COL_间隔单位))
                            If Not blnValid Then
                                If .TextMatrix(i, COL_间隔单位) = "周" Then
                                    strMsg = COL_按周执行
                                ElseIf .TextMatrix(i, COL_间隔单位) = "天" Then
                                    strMsg = COL_按天执行
                                ElseIf .TextMatrix(i, COL_间隔单位) = "小时" Then
                                    strMsg = COL_按时执行
                                End If
                                strMsg = "录入的执行时间方案格式不正确，请检查。" & vbCrLf & vbCrLf & "例：" & vbCrLf & strMsg
                                .Col = COL_执行时间: Exit For
                            End If
                        End If
                        
                        '医保对码检查:以一组医保第一可见行为准
                        If InStr(",5,6,", .TextMatrix(i, COL_类别)) = 0 _
                            Or Val(.TextMatrix(i - 1, COL_相关ID)) <> Val(.TextMatrix(i, COL_相关ID)) Then
                            If gint医保对码 = 2 Then mbln提醒对码 = True
                            Call GetInsureStr(strIDs1, strIDs2, str医嘱内容, i)
                            strMsg = CheckAdviceInsure(mint险类, mbln提醒对码, mlng病人ID, 1, strIDs1, strIDs2, str医嘱内容)
                            If strMsg <> "" Then
                                .Row = i: .Col = col_医嘱内容: Call .ShowCell(.Row, .Col)
                                If gint医保对码 = 1 Then
                                    vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                                    If vMsg = vbNo Or vMsg = vbCancel Then Exit Function
                                    If vMsg = vbIgnore Then mbln提醒对码 = False
                                ElseIf gint医保对码 = 2 Then
                                    MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                                    Exit Function
                                End If
                                strMsg = "" '防止后面再作处理
                            End If
                        End If
                        
                        '医保管控实时监测：首次输入(经过)或者更改时检查
                        If mint险类 <> 0 And .Cell(flexcpData, i, COL_状态) = 0 Then
                            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                                If MakePriceRecord(i) Then
                                    If Not gclsInsure.CheckItem(mint险类, 0, 0, mrsPrice) Then
                                        .Row = i: .Col = col_医嘱内容
                                        Call .ShowCell(.Row, .Col)
                                        If txt总量.Enabled Then
                                            txt总量.SetFocus
                                        ElseIf txt医嘱内容.Enabled Then
                                            txt医嘱内容.SetFocus
                                        End If
                                        Exit Function
                                    End If
                                End If
                                '标记为已经作了检查
                                .Cell(flexcpData, .Row, COL_状态) = 1
                            End If
                        End If
                    End If
                                    
                    '医嘱申请附项填写检查：
                    '只针对新录入的医嘱，修改的医嘱修改时已检查
                    '".Cell(flexcpData, i, COL_附项) = 1"的可能还没有在输入过程中检查，只是自动替换了
                    If Val(.TextMatrix(i, COL_EDIT)) = 1 And .TextMatrix(i, COL_附项) <> "" Then
                        strMsg = CheckAdviceAppend(.TextMatrix(i, COL_附项))
                        If strMsg <> "" Then
                            strMsg = "申请附项""" & strMsg & """没有录入，请确认系统默认填写的信息是否正确。"
                            .Col = col_医嘱内容: blnAppend = True: Exit For
                        End If
                    End If
                                    
                    '互斥数据收集:在所有有效医嘱中,因为可能已发送的与未发送的互斥
                    If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                        '用于药品配伍禁忌检查:不分期效
                        str药品IDs = str药品IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                    ElseIf Not bln配方行 Then
                        '不管检查组合与手术附加内部之间及内部与其它项目之间
                        str诊疗IDs = str诊疗IDs & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '中间退出的错误提示
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln配方行 Then
                    strMsg = "该中药配方" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, col_医嘱内容) & """" & strMsg
                End If
                If Not blnOut Then '超量说明提示特殊处理
                    MsgBox strMsg, vbInformation, gstrSysName
                End If
                blnOut = False
                .Refresh
            End If
            Call vsAdvice_KeyPress(13)
            If blnAppend Then '是否弹出申请附项编辑
                If cmdExt.Enabled And cmdExt.Visible Then Call cmdExt_Click
            End If
            Exit Function
        End If
        
        '检查药品配伍禁忌
        If str药品IDs <> "" Then
            If Not Check配伍禁忌(Mid(str药品IDs, 2)) Then Exit Function
        End If
        '检查诊疗项目互斥
        If str诊疗IDs <> "" Then
            If Not Check诊疗互斥(Mid(str诊疗IDs, 2)) Then Exit Function
        End If
    End With
    
    '费用报警:有未发送医嘱时
    If lngCount > 0 Then
        If Not CheckMoney Then Exit Function
    End If
    '复诊不弹出传染病报告卡
    If Not mbln复诊 Then
        '根据诊断判断是否应该书写传染病报告卡
        RaiseEvent CheckInfectDisease(False, Mid(str疾病IDs, 2), Mid(str诊断IDs, 2), blnNo)
    End If
        If blnNo Then Exit Function
    '--检查病人是否退号，是否取消了接诊
    If Not CheckBackNo(mstr挂号单) Then Exit Function
    
    CheckAdvice = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SeekNextControl() As Boolean
'功能：定位到下一个焦点的控件上,并根据情况决定是否自动新增一行医嘱
'返回：如果通过SetFocus强制定位的,则返回True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange '强制更新医嘱内容
                                vsAdvice.Row = i
                                
                                '可能已在其他事件中被定位了，不用重复移动定位表格
                                If Not Me.ActiveControl Is Nothing Then
                                    If Not Me.ActiveControl Is vsAdvice Then
                                        Call zlCommFun.PressKey(vbKeyTab)
                                    End If
                                End If
                                '表格行无内容则再移动定位到医嘱输入框
                                blnDo = vsAdvice.RowData(i) <> 0
                                
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            cbsMain.FindControl(, conMenu_New, True, True).Execute
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        If objNext.Enabled And objNext.Visible Then
                            blnDo = True
                            objNext.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab) '自然定位
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'功能：获取输入医嘱过程中，回车光标应跳过的控件
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        '一并给药中的药品输入时应跳过的内容
        If InStr(",5,6,", .TextMatrix(lngRow, COL_类别)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_相关ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                '给药途径,附加执行
                If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_诊疗项目ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt用法.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_执行科室ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo附加执行.Name
                        End If
                    End If
                End If
                '频率
                If .TextMatrix(lngRow, COL_频率) <> "" Then strSkip = strSkip & ";" & Me.txt频率.Name
                '执行时间
                If .TextMatrix(lngRow, COL_执行时间) <> "" Then strSkip = strSkip & ";" & Me.cbo执行时间.Name
            End If
        ElseIf InStr(",C,D,F,G,Z", .TextMatrix(lngRow, COL_类别)) > 0 And .RowData(lngRow) <> 0 And .TextMatrix(lngRow, COL_频率) = "一次性" Then
            strSkip = strSkip & ";" & Me.txt频率.Name
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Sub SetBabyVisible(ByVal lng科室id As Long)
'功能：根据科室性质设置婴儿医嘱是否可以选择
'说明：产科才有婴儿医嘱
    If DeptIsWoman(lng科室id) Then
        lbl婴儿.Visible = True
        cbo婴儿.Visible = True
    Else
        Call zlControl.CboSetIndex(cbo婴儿.hWnd, 0)
        cbo婴儿.Tag = 0
        lbl婴儿.Visible = False
        cbo婴儿.Visible = False
    End If
End Sub

Private Sub CalcAdviceMoney()
'功能：计算新开医嘱金额
'说明：只管当前显示出的部份新开医嘱
    Dim dblMoney As Double, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) And Val(.TextMatrix(i, COL_状态)) = 1 Then
                dblMoney = dblMoney + Format(CCur(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单价))), gstrDec)
            End If
        Next
        stbThis.Panels(5).Text = "新开:" & FormatEx(dblMoney, 5) & "元"
    End With
End Sub

Private Sub AdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名id As Long, lng证书ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '自动保存
    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mstr挂号单, strIDs, 0, False, strSource, mstr前提IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lng签名id = zlDatabase.GetNextID("医嘱签名记录")
            strSQL = "zl_医嘱签名记录_Insert(" & lng签名id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        '重新读取显示医嘱
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        mblnOK = True
        If txt医嘱内容.Enabled Then
            txt医嘱内容.SetFocus
        Else
            vsAdvice.SetFocus
        End If

        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'功能：当医嘱卡片输入内容变化时，判断医嘱内容文本是否应该重新组织
    Dim str类别 As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 Then '中药配方或一组检验
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If lngRow <> -1 Then str类别 = .TextMatrix(lngRow, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '检查内容变动
        If blnDefine Then '公共字段部份或可以公共处理的部份
            If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" And InStr(strText, "[开始时间]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsDate(txt安排时间.Text) And txt安排时间.Tag <> "" Then
                If InStr(strText, "[手术时间]") > 0 Or InStr(strText, "[输血时间]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cbo医生嘱托.Tag <> "" And InStr(strText, "[医生嘱托]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmd频率.Tag <> "" And txt频率.Tag <> "" Then
                If InStr(strText, "[中文频率]") > 0 Or InStr(strText, "[英文频率]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cbo执行时间.Tag <> "" And InStr(strText, "[执行时间]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt单量.Text) Or txt单量.Text = "") And txt单量.Tag <> "" And InStr(strText, "[单量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[总量]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str类别 '不同的类别检查
        Case "5", "6" '中西成药
            If Not blnDefine Then
                
            Else
                '[输入名][通用名][商品名][英文名][规格][产地]是输入或修改整个药品时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[给药途径]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '中药配方
            If Not blnDefine Then
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmd频率.Tag <> "" And txt频率.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[配方组成][煎法]是输入或修改整个配方时变化
                If IsNumeric(txt总量.Text) And txt总量.Tag <> "" And InStr(strText, "[付数]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[用法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '检验
            If Not blnDefine Then
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[检验项目][检验标本]是输入或修改整个项目时变化
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[采集方法]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '检查
            If Not blnDefine Then
                
            Else
                '[检查项目][检查部位]是输入或修改整个项目时变化
            End If
        Case "F" '手术
            If Not blnDefine Then
                If IsDate(txt安排时间.Text) And txt安排时间.Tag <> "" Then AdviceTextChange = True: Exit Function
                If IsDate(txt开始时间.Text) And txt开始时间.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[主要手术][附加手术][麻醉方法]是输入或修改整个项目时变化
            End If
        Case "K" '输血
            If Not blnDefine Then
                If IsDate(txt安排时间.Text) And txt安排时间.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[输血途径]
                If Val(cmd用法.Tag) <> 0 And txt用法.Tag <> "" And InStr(strText, "[输血途径]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case Else '其他
            If Not blnDefine Then
                
            Else
                '[诊疗项目]是输入或修改整个项目时变化
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'功能：获取医嘱内容文本
'参数：lngRow=已有医嘱数据的可见行
    Dim rsTmp As New ADODB.Recordset
    Dim rsCard As New ADODB.Recordset
    Dim blnDefine As Boolean, str类别 As String
    Dim strText As String, strSQL As String
    Dim strField As String, int频率范围 As Integer
    Dim i As Long, k As Long
    Dim blnDo As Boolean
    Dim str中药名称 As String
    
    Dim str中药 As String, str煎法 As String, str形态 As String
    Dim str麻醉 As String, str附术 As String
    Dim str检验 As String, str标本 As String
    Dim str部位 As String, str部位Last As String, str方法 As String
    Dim dbl数量 As Double
    Dim str中药诊疗项目IDS As String, strSame As String
    
    On Error GoTo errH
    
    With vsAdvice
        '确定医嘱类别
        str类别 = .TextMatrix(lngRow, COL_类别)
        If str类别 = "E" Then '中药配方或一组检验
            k = .FindRow(CStr(.RowData(lngRow)), , COL_相关ID)
            If k <> -1 Then str类别 = .TextMatrix(k, COL_类别)
        End If
        If str类别 = "7" Then str类别 = "8"
                
        '确定是否定义
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!医嘱内容)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '用于按定义公式计算失败，重新按缺省规则进行组织
        strText = ""
        If blnDefine Then strText = mrsDefine!医嘱内容
        
        '产生医嘱内容
        Select Case str类别
        Case "C" '检验-------------------------------------------------------------
            str检验 = "": str标本 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_组合项目ID)) = 0 And mblnNewLIS Or Not mblnNewLIS Then
                        str检验 = .TextMatrix(i, col_医嘱内容) & "," & str检验
                    End If
                    str标本 = .TextMatrix(i, COL_标本部位)
                Else
                    Exit For
                End If
            Next
            If str检验 = "" Then '老的方式
                str检验 = .TextMatrix(lngRow, COL_名称)
            Else
                str检验 = Left(str检验, Len(str检验) - 1)
            End If
            
            If Not blnDefine Then
                strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
            Else
                If InStr(strText, "[检验项目]") > 0 Then
                    strField = str检验
                    strText = Replace(strText, "[检验项目]", """" & strField & """")
                End If
                If InStr(strText, "[检验标本]") > 0 Then
                    strField = str标本
                    strText = Replace(strText, "[检验标本]", """" & strField & """")
                End If
                If InStr(strText, "[采集方法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[采集方法]", """" & strField & """")
                End If
            End If
        Case "D" '检查-------------------------------------------------------------
            str部位 = "": str方法 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_标本部位) <> "" Then
                        If .TextMatrix(i, COL_标本部位) <> str部位Last And str部位Last <> "" Then
                            str部位 = str部位 & "," & str部位Last & IIF(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
                            str方法 = ""
                        End If
                        If .TextMatrix(i, COL_检查方法) <> "" Then
                            str方法 = str方法 & "," & .TextMatrix(i, COL_检查方法)
                        End If
                        
                        str部位Last = .TextMatrix(i, COL_标本部位)
                    End If
                Else
                    Exit For
                End If
            Next
            If str部位Last <> "" Then
                str部位 = str部位 & "," & str部位Last & IIF(str方法 <> "", "(" & Mid(str方法, 2) & ")", "")
            End If
            str部位 = Mid(str部位, 2) '检查组合项目的部位
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称) & _
                    Decode(Val(.TextMatrix(lngRow, COL_执行标记)), 1, ",床旁执行", 2, ",术中执行", "") & IIF(str部位 <> "", ":" & str部位, "")
            Else
                If InStr(strText, "[检查项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称) & _
                        Decode(Val(.TextMatrix(lngRow, COL_执行标记)), 1, ",床旁执行", 2, ",术中执行", "")
                    strText = Replace(strText, "[检查项目]", """" & strField & """")
                End If
                If InStr(strText, "[检查部位]") > 0 Then
                    strField = str部位
                    strText = Replace(strText, "[检查部位]", """" & strField & """")
                End If
            End If
        Case "F" '手术-------------------------------------------------------------
            str麻醉 = "": str附术 = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "G" Then
                        str麻醉 = .TextMatrix(i, col_医嘱内容)
                    Else
                        str附术 = str附术 & "," & .TextMatrix(i, col_医嘱内容)
                    End If
                Else
                    Exit For
                End If
            Next
            str附术 = Mid(str附术, 2)
            
            If Not blnDefine Then
                If IsDate(.TextMatrix(lngRow, COL_标本部位)) Then
                    strText = Format(.TextMatrix(lngRow, COL_标本部位), "MM月dd日HH:mm")
                Else
                    strText = Format(.Cell(flexcpData, lngRow, COL_开始时间), "MM月dd日HH:mm")
                End If
                If str麻醉 <> "" Then
                    strText = strText & IIF(str麻醉 <> "", " 在 " & str麻醉 & " 下行 ", " 行 ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_名称) & IIF(.Cell(flexcpData, lngRow, COL_标本部位) = "", "", "(部位:" & .Cell(flexcpData, lngRow, COL_标本部位) & ")")
                If str附术 <> "" Then
                    strText = strText & " 及 " & str附术
                End If
            Else
                If InStr(strText, "[手术时间]") > 0 Then
                    If IsDate(.TextMatrix(lngRow, COL_手术时间)) Then
                        strField = .TextMatrix(lngRow, COL_手术时间)
                    Else
                        strField = .Cell(flexcpData, lngRow, COL_开始时间)
                    End If
                    strText = Replace(strText, "[手术时间]", """" & strField & """")
                End If
                If InStr(strText, "[主要手术]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称) & IIF(.Cell(flexcpData, lngRow, COL_标本部位) = "", "", "(部位:" & .Cell(flexcpData, lngRow, COL_标本部位) & ")")
                    strText = Replace(strText, "[主要手术]", """" & strField & """")
                End If
                If InStr(strText, "[附加手术]") > 0 Then
                    strField = str附术
                    strText = Replace(strText, "[附加手术]", """" & strField & """")
                End If
                If InStr(strText, "[麻醉方法]") > 0 Then
                    strField = str麻醉
                    strText = Replace(strText, "[麻醉方法]", """" & strField & """")
                End If
            End If
        Case "8" '中药配方---------------------------------------------------------
            str中药 = "": str煎法 = "": str中药诊疗项目IDS = "": strSame = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        If InStr("," & str中药诊疗项目IDS & ",", "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & ",") > 0 Then
                            strSame = strSame & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                        End If
                        str中药诊疗项目IDS = str中药诊疗项目IDS & "," & Val(.TextMatrix(i, COL_诊疗项目ID))
                    End If
                Else
                    Exit For
                End If
            Next
            strSame = Mid(strSame, 2)
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_类别) = "7" Then
                        dbl数量 = dbl数量 + Val(.TextMatrix(i, COL_单量))
                        
                        If Val(.TextMatrix(lngRow, COL_中药形态)) = 0 Then
                            blnDo = .TextMatrix(i, COL_收费细目ID) <> .TextMatrix(i - 1, COL_收费细目ID)
                        Else
                            blnDo = .TextMatrix(i, COL_诊疗项目ID) <> .TextMatrix(i - 1, COL_诊疗项目ID)
                        End If
                        
                        If blnDo Then
                            str中药名称 = .TextMatrix(i, col_医嘱内容)
                            
                            If Val(.TextMatrix(lngRow, COL_中药形态)) = 0 And InStr("," & strSame & ",", "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & ",") > 0 Then
                                strSQL = "Select 规格 as 名称 From 收费项目目录 Where ID=[1] And Exists(Select 1 From 药品规格 Where 药品ID<>[1] And 药名ID=[2])"
                                Set rsTmp = New ADODB.Recordset '清除Filter
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_诊疗项目ID)))
                                If rsTmp.RecordCount > 0 Then
                                    If Not IsNull(rsTmp!名称) Then str中药名称 = str中药名称 & "(" & rsTmp!名称 & ")"
                                End If
                            End If
                        
                            str中药 = RTrim(str中药名称 & _
                                " " & FormatEx(dbl数量, 5) & .TextMatrix(i, COL_单量单位) & _
                                " " & .TextMatrix(i, COL_医生嘱托)) & "," & str中药
                            dbl数量 = 0
                        End If
                    ElseIf .TextMatrix(i, COL_类别) = "E" Then
                        str煎法 = .TextMatrix(i, col_医嘱内容) & .TextMatrix(i, COL_标本部位)
                    End If
                Else
                    Exit For
                End If
            Next
            If str中药 <> "" Then
                str中药 = Mid(str中药, 1, Len(str中药) - 1)
            End If
            If Not blnDefine Then
                If .TextMatrix(lngRow, COL_中药形态) = "1" Then
                    str形态 = "[饮片]"
                ElseIf .TextMatrix(lngRow, COL_中药形态) = "2" Then
                    str形态 = "[免煎剂]"
                End If
                '数字后加了空格在文本框中会自动换行
                strText = "中药" & str形态 & .TextMatrix(lngRow, COL_总量) & "付," & _
                    .TextMatrix(lngRow, COL_频率) & "," & str煎法 & "," & _
                    .TextMatrix(lngRow, COL_用法) & ":" & str中药
            Else
                If InStr(strText, "[付数]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_总量)
                    strText = Replace(strText, "[付数]", """" & strField & """")
                End If
                If InStr(strText, "[配方组成]") > 0 Then
                    strField = str中药
                    strText = Replace(strText, "[配方组成]", """" & strField & """")
                End If
                If InStr(strText, "[用法]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[用法]", """" & strField & """")
                End If
                If InStr(strText, "[煎法]") > 0 Then
                    strField = str煎法
                    strText = Replace(strText, "[煎法]", """" & strField & """")
                End If
            End If
        Case "4" '卫材------------------------------------------------------------
                strSQL = "Select 名称,规格,产地 From 收费项目目录 Where ID=[1]"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
                
                If Not blnDefine Then
                    strText = .TextMatrix(lngRow, COL_名称)
                    If Not IsNull(rsTmp!规格) Then
                        strText = strText & " " & rsTmp!规格
                    End If
                Else
                    If InStr(strText, "[卫生材料]") > 0 Then
                        strField = rsTmp!名称
                        strText = Replace(strText, "[卫生材料]", """" & strField & """")
                    End If
                    If InStr(strText, "[规格]") > 0 Then
                        strField = Nvl(rsTmp!规格)
                        strText = Replace(strText, "[规格]", """" & strField & """")
                    End If
                    If InStr(strText, "[产地]") > 0 Then
                        strField = Nvl(rsTmp!产地)
                        strText = Replace(strText, "[产地]", """" & strField & """")
                    End If
                End If
        Case "5", "6" '西成药，中成药---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                '性质:0-正名,1-英文名,3-商品名
                strSQL = "Select Nvl(B.名称,A.名称) as 名称,A.规格,A.产地,B.性质" & _
                    " From 收费项目目录 A,收费项目别名 B Where A.ID=B.收费细目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_收费细目ID)))
            ElseIf blnDefine Then
                '性质:0-正名,1-英文名
                strSQL = "Select Nvl(B.名称,A.名称) as 名称,Null as 规格,Null as 产地,B.性质" & _
                    " From 诊疗项目目录 A,诊疗项目别名 B Where A.ID=B.诊疗项目ID(+) And A.ID=[1] Order by B.性质,B.码类"
                Set rsTmp = New ADODB.Recordset '清除Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_诊疗项目ID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_标本部位)
                If Val(.TextMatrix(lngRow, COL_收费细目ID)) <> 0 Then
                    If strText = "" Then
                        If gbyt药品名称显示 <> 0 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!名称
                    End If
                    If Not IsNull(rsTmp!产地) Then
                        strText = strText & "(" & rsTmp!产地 & ")"
                    End If
                    If Not IsNull(rsTmp!规格) Then
                        strText = strText & " " & rsTmp!规格
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_名称)
                    End If
                End If
            Else
                If InStr(strText, "[输入名]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_标本部位)
                    If strField = "" Then
                        If gbyt药品名称显示 <> 0 Then rsTmp.Filter = "性质=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[输入名]", """" & strField & """")
                End If
                If InStr(strText, "[通用名]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!名称
                    strText = Replace(strText, "[通用名]", """" & strField & """")
                End If
                If InStr(strText, "[商品名]") > 0 Then
                    rsTmp.Filter = "性质=3"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[商品名]", """" & strField & """")
                End If
                If InStr(strText, "[英文名]") > 0 Then
                    rsTmp.Filter = "性质=2"
                    If rsTmp.EOF Then
                        strField = ""
                    Else
                        strField = rsTmp!名称
                    End If
                    strText = Replace(strText, "[英文名]", """" & strField & """")
                End If
                If InStr(strText, "[规格]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!规格)
                    strText = Replace(strText, "[规格]", """" & strField & """")
                End If
                If InStr(strText, "[产地]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!产地)
                    strText = Replace(strText, "[产地]", """" & strField & """")
                End If
                If InStr(strText, "[给药途径]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[给药途径]", """" & strField & """")
                End If
            End If
        Case "K" '输血医嘱
            If Not blnDefine Then
                If IsDate(.TextMatrix(lngRow, COL_输血时间)) Then
                    strText = Format(.TextMatrix(lngRow, COL_输血时间), "MM月dd日HH:mm")
                Else
                    strText = Format(.Cell(flexcpData, lngRow, COL_开始时间), "MM月dd日HH:mm")
                End If
            
                strText = "于" & strText & "输" & .TextMatrix(lngRow, COL_名称)
                If .TextMatrix(lngRow, COL_用法) <> "" Then
                    strText = strText & "(" & .TextMatrix(lngRow, COL_用法) & ")"
                End If
            Else
                If InStr(strText, "[输血时间]") > 0 Then
                    If IsDate(.TextMatrix(lngRow, COL_输血时间)) Then
                        strField = .TextMatrix(lngRow, COL_输血时间)
                    Else
                        strField = .Cell(flexcpData, lngRow, COL_开始时间)
                    End If
                    strText = Replace(strText, "[输血时间]", """" & strField & """")
                End If
                If InStr(strText, "[诊疗项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[诊疗项目]", """" & strField & """")
                End If
                If InStr(strText, "[输血项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[输血项目]", """" & strField & """")
                End If
                If InStr(strText, "[输血途径]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_用法)
                    strText = Replace(strText, "[输血途径]", """" & strField & """")
                End If
                If TypeName(.Cell(flexcpData, lngRow, COL_申请序号)) = "Recordset" Then
                    Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_申请序号))
                    If InStr(strText, "[血型]") > 0 Then
                        If rsCard.EOF Then
                            strField = ""
                        Else
                            strField = Decode(Val("" & rsCard!血型), 1, "A", 2, "B", 3, "O", 4, "AB", "")
                        End If
                        strText = Replace(strText, "[血型]", """" & strField & """")
                    End If
                    If InStr(strText, "[RH]") > 0 Then
                        If rsCard.EOF Then
                            strField = ""
                        Else
                            strField = Decode(Val("" & rsCard!RHD), 1, "-", 2, "+", "")
                        End If
                        strText = Replace(strText, "[RH]", """" & strField & """")
                    End If
                End If
                If InStr(strText, "[执行分类]") > 0 Then
                    strField = Val(.TextMatrix(lngRow + 1, COL_执行分类))
                    strText = Replace(strText, "[执行分类]", """" & strField & """")
                End If
            End If
        Case Else '其它所有类别-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_名称)
            Else
                If InStr(strText, "[诊疗项目]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_名称)
                    strText = Replace(strText, "[诊疗项目]", """" & strField & """")
                End If
            End If
            '术后医嘱特殊显示
            If .TextMatrix(lngRow, COL_类别) = "Z" And (Val(.TextMatrix(lngRow, COL_操作类型)) = 4 Or Val(.TextMatrix(lngRow, COL_操作类型)) = 14) Then
                strText = "━━━" & strText & "━━━"
            End If
        End Select
        
        '公共字段或可以公共处理的字段-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[开始时间]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_开始时间)
                strText = Replace(strText, "[开始时间]", """" & strField & """")
            End If
            If InStr(strText, "[医生嘱托]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_医生嘱托)
                If .TextMatrix(lngRow, COL_医生嘱托) <> "" Then
                    If strField <> "" Then
                        strField = strField & "," & .TextMatrix(lngRow, COL_医生嘱托)
                    Else
                        strField = .TextMatrix(lngRow, COL_医生嘱托)
                    End If
                End If
                strText = Replace(strText, "[医生嘱托]", """" & strField & """")
            End If
            If InStr(strText, "[中文频率]") > 0 Then
                strField = .TextMatrix(lngRow, COL_频率)
                strText = Replace(strText, "[中文频率]", """" & strField & """")
            End If
            If InStr(strText, "[英文频率]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_频率) <> "" Then
                    int频率范围 = Get频率范围(lngRow)
                    strSQL = "Select 英文名称 From 诊疗频率项目 Where 名称=[1] And 适用范围=[2]"
                    Set rsTmp = New ADODB.Recordset '清除Filter
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, COL_频率), int频率范围)
                    If Not rsTmp.EOF Then strField = Nvl(rsTmp!英文名称)
                End If
                strText = Replace(strText, "[英文频率]", """" & strField & """")
            End If
            If InStr(strText, "[单量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_单量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_单量) & .TextMatrix(lngRow, COL_单量单位)
                End If
                strText = Replace(strText, "[单量]", """" & strField & """")
            End If
            If InStr(strText, "[总量]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_总量) <> "" Then
                    strField = .TextMatrix(lngRow, COL_总量) & .TextMatrix(lngRow, COL_总量单位)
                End If
                strText = Replace(strText, "[总量]", """" & strField & """")
            End If
            If InStr(strText, "[执行时间]") > 0 Then
                strField = .TextMatrix(lngRow, COL_执行时间)
                strText = Replace(strText, "[执行时间]", """" & strField & """")
            End If
        End If
                
        '计算医嘱内容
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetAdviceAppendItem() As String
'功能：获取未保存医嘱(新开或修改的)的最新单据附项
'返回：项目名1<Split2>内容1<Split1>项目名2<Split2>内容2<Split1>...
    Dim arrAppend As Variant, i As Long, j As Long
    Dim strName As String, strText As String
    Dim strResult As String
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) <> 0 And .TextMatrix(i, COL_附项) <> "" And .Cell(flexcpData, i, COL_附项) = 1 Then
                arrAppend = Split(.TextMatrix(i, COL_附项), "<Split1>")
                For j = 0 To UBound(arrAppend)
                    strName = Split(arrAppend(j), "<Split2>")(0)
                    strText = Split(arrAppend(j), "<Split2>")(3)
                    
                    If InStr(strResult, "<Split1>" & strName & "<Split2>") = 0 Then
                        strResult = strResult & "<Split1>" & strName & "<Split2>" & strText
                    End If
                Next
            End If
        Next
    End With
    
    GetAdviceAppendItem = Mid(strResult, Len("<Split1>") + 1)
End Function

Private Function GetAdviceDiagnosis() As String
'功能：获取当前已录入的未保存的诊断
'返回："1、诊断名一 2、诊断名二"
    Dim strText As String, i As Long, j As Long
    
    With vsDiag
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                j = j + 1
                strText = strText & "  " & j & "、" & .TextMatrix(i, col诊断) & IIF(Val(.Cell(flexcpData, i, col疑诊)) = 1, "(？)", "")
            End If
        Next
    End With
    
    strText = Mid(strText, 3)
    If j = 1 Then strText = Mid(strText, 3)
    GetAdviceDiagnosis = strText
End Function

Private Sub vsDiag_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiag
        If Col = col诊断 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '在调用vsDiagXY_KeyDown(vbKeyDelete, 0)点是可以删除当前行，点否则恢复原始数据
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiag_KeyDown(vbKeyDelete, 0)
            End If
        End If
        If .Col = Col Then Call vsDiag_AfterRowColChange(-1, -1, Row, Col)
  
    End With
End Sub

Private Sub vsDiag_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiag
        '清除图片
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col增加) Is Nothing Then
                Set .Cell(flexcpPicture, i, col增加) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, COLDEL) Is Nothing Then
               Set .Cell(flexcpPicture, i, COLDEL) = Nothing
            End If
        Next
        '设置编辑可见特性
        If Not DiagCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            If NewCol = col诊断 Then
                .ComboList = "..."
            ElseIf NewCol = col增加 Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = COLDEL Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            ElseIf NewCol = col中医证候 Then
                If .TextMatrix(NewRow, col诊断) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '显示图片
            If NewCol <> col增加 And .TextMatrix(NewRow, col诊断) <> "" Then
                Set .Cell(flexcpPicture, NewRow, col增加) = imgButtonNew.Picture
            End If
            '显示图片
            If NewCol <> COLDEL Then
                Set .Cell(flexcpPicture, NewRow, COLDEL) = imgButtonDel.Picture
            End If
        End If
        
        If NewRow <> OldRow Then
            '当前行标志显示
            Set .Cell(flexcpPicture, 0, col标志, .Rows - 1, col标志) = Nothing
            Set .Cell(flexcpPicture, NewRow, col标志) = img16.ListImages("诊断_当前").Picture
            .Cell(flexcpPictureAlignment, NewRow, col标志) = 4
                    
            '显示关联医嘱标识
            Call ShowDiagFlag(NewRow)
        End If
    End With
End Sub

Private Sub vsDiag_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '在StartEdit事件中处理会显示出前一列的按钮
    If Col = col疑诊 Then Cancel = True
End Sub

Private Sub vsDiag_Click()
    With vsDiag
        If (.MouseCol = col增加 Or .MouseCol = COLDEL) And .MouseRow >= .FixedRows Then
            If .MouseCol = col增加 Then
                If .TextMatrix(.MouseRow, col诊断) = "" Then Exit Sub
            End If
            
            .Select .MouseRow, .MouseCol
            Call vsDiag_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiag_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = col标志 Then
        Cancel = True
    ElseIf (NewCol = col中医 Or NewCol = COL西医) And vsDiag.TextMatrix(NewRow, col诊断) <> "" Then
        Cancel = True
    End If
End Sub

Private Sub vsDiag_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str性别 As String
    Dim lngRow As Long, int诊断类型 As Integer
    Dim blnCancle As Boolean
    Dim str类别 As String
    
    With vsDiag
        If Col = col诊断 Then
            If .Cell(flexcpData, Row, col中医) = 1 Then
                If opt诊断(0).value Then
                    '按诊断输入:中医部份，一个诊断可能属于多个分类
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng病人科室id, , True, False, , , 1)
                    str类别 = "2"
                Else
                    'B-中医疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng病人科室id, mstr性别, True, , , , 1)
                    str类别 = "B"
                End If
            Else
                If opt诊断(0).value Then
                    '按诊断输入:西医部份，一个诊断可能属于多个分类
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng病人科室id, , True, False, , , 1)
                    str类别 = "1"
                Else
                    'D-ICD-10疾病编码
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", mlng病人科室id, mstr性别, True, , , , 1)
                    str类别 = "D"
                End If
            End If
            If rsTmp Is Nothing Then
                If opt诊断(0).value Then
                    MsgBox "没有疾病诊断数据可以选择。", vbInformation, gstrSysName
                End If
            Else
                Call SetDiagInput(Row, rsTmp, str类别)
                Call DiagEnterNextCell
            End If
        ElseIf Col = col中医证候 Then
            If opt诊断(0).value Then
                '按诊断输入:先查是否有对应
                If Not Set中医证候(Row, Val(.TextMatrix(Row, col诊断ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng病人科室id, mstr性别, True, , , , 1)
                Else
                    Exit Sub
                End If
            Else
                'Z-中医疾病编码
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng病人科室id, mstr性别, True, , , , 1)
            End If
            If Not rsTmp Is Nothing Then
                Call Set中医证候(Row, 0, rsTmp)
                Call DiagEnterNextCell
            End If
        ElseIf Col = col增加 Then
            If .Rows < M_LNG_DIAGCOUNT Then
                lngRow = Row + 1: .AddItem "", lngRow
                .Row = lngRow: .Col = col诊断
                
                int诊断类型 = IIF(mbln中医, 11, 1)
                If lngRow - 1 >= .FixedRows Then
                    int诊断类型 = IIF(.Cell(flexcpData, lngRow - 1, col中医) = 1, 11, 1)
                End If
                Call SetDiagType(lngRow, int诊断类型)
                
                Call SetDiagHeight
            End If
        ElseIf Col = COLDEL Then
            Call vsDiag_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiag_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsDiag
        If Col = col诊断 Then
            .TextMatrix(Row, col疑诊) = IIF(.TextMatrix(Row, Col) <> "", "？", "")
        End If
    End With
End Sub

Private Sub vsDiag_DblClick()
    Call vsDiag_KeyPress(32)
End Sub

Private Sub vsDiag_GotFocus()
    If Me.Visible Then vsDiag_AfterRowColChange -1, -1, vsDiag.Row, vsDiag.Col
End Sub

Private Sub vsDiag_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnDo As Boolean, i As Long
    Dim int诊断类型 As Integer
    
    With vsDiag
        If KeyCode = vbKeyF4 Then
            If .Col = col诊断 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            '检查是否允许删除
            For i = 1 To vsAdvice.Rows - 1
                If InStr("," & .TextMatrix(.Row, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_状态) = "8" Then
                    MsgBox "该诊断对应的处方已发送，不能删除。", vbInformation, Me.Caption
                    Exit Sub
                End If
                '医技工作站调用,若诊断关联医嘱,存在医嘱的开嘱医生非当前操作员,则不允许删除诊断
                If mint场合 = 2 And InStr("," & .TextMatrix(.Row, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_开嘱医生) <> UserInfo.姓名 Then
                    MsgBox "该诊断存在关联医嘱,且该医嘱非您下达，不能删除。", vbInformation, Me.Caption
                    Exit Sub
                End If
            Next
            blnDo = True
            If .TextMatrix(.Row, col诊断) <> "" Then
                If MsgBox("确实要删除该行诊断信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnDo = False
            End If
            If blnDo Then
                If .TextMatrix(.Row, col诊断) <> "" Then
                    lbl诊断.Tag = "1"
                    mblnNoSave = True
                    '删除主/次要诊断后调用外挂接口
                    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
                        On Error Resume Next
                        Call gobjPlugIn.DiagnosisDeleted(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(.TextMatrix(.Row, col诊断ID)), .TextMatrix(.Row, col诊断), mint场合)
                        Call zlPlugInErrH(err, "DiagnosisDeleted")
                        err.Clear: On Error GoTo 0
                    End If
                    If mblnPass Then
                        zlPassDrags
                    End If
                End If
                
                If .Rows = 2 And .Row = 1 Then
                    int诊断类型 = IIF(.Cell(flexcpData, .Row, col中医) = 1, 11, 1)
                    .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, 0, .Row, .Cols - 1) = Empty
                    Set .Cell(flexcpPicture, .Row, 0, .Row, .Cols - 1) = Nothing
                    Call SetDiagType(.Row, int诊断类型)
                Else
                    .RemoveItem .Row
                    Call SetDiagHeight
                End If
            End If
            .SetFocus
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsDiag_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiag_KeyPress(KeyAscii As Integer)
    With vsDiag
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiagEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col疑诊 Or .Col = col中医 Or .Col = COL西医) Then
            KeyAscii = 0
            If .Col = col中医 Then
                If .Cell(flexcpData, .Row, col中医) = 0 Then
                    Call SetDiagType(.Row, 11): .Col = col诊断
                End If
            ElseIf .Col = COL西医 Then
                If .Cell(flexcpData, .Row, COL西医) = 0 Then
                    Call SetDiagType(.Row, 1): .Col = col诊断
                End If
            ElseIf .Col = col疑诊 Then
                If DiagCellEditable(.Row, .Col) Then
                    KeyAscii = 0
                    .Cell(flexcpData, .Row, .Col) = IIF(.Cell(flexcpData, .Row, .Col) = 1, 0, 1)
                    .Cell(flexcpForeColor, .Row, .Col) = IIF(.Cell(flexcpData, .Row, .Col) = 1, vbRed, .GridColor)
                    
                    lbl诊断.Tag = "1"
                    mblnNoSave = True
                End If
            End If
        Else
            If .Col = col诊断 Or .Col = col中医证候 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiag_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiag_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiag_LostFocus()
    If vsDiag.Col <> col发病时间 Then vsDiag.Col = IIF(vsDiag.Col = col中医证候, col中医证候, col诊断)
End Sub

Private Sub vsDiag_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiag.EditSelStart = 0
    vsDiag.EditSelLength = zlCommFun.ActualLen(vsDiag.EditText)
End Sub

Private Sub vsDiag_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not DiagCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col疑诊 Or Col = col中医 Or Col = COL西医 Then
        Cancel = True '不直接编辑
    End If
End Sub

Private Function GetDiagSQL(ByVal Row As Long, ByRef strInput As String, ByRef strSQL As String, ByRef str性别 As String, Optional ByVal strType As String) As String
'功能：获得查询诊断的SQL
'参数：strInput-查询条件,strsql--返回的SQL，str性别--病人的性别  ,strType疾病编码种类。
'返回：strsql--查询中医诊断的SQL
    If vsDiag.Cell(flexcpData, Row, col中医) = 1 Then
        If opt诊断(0).value And strType <> "Z" Then
            '按诊断输入:中医部份，一个诊断可能属于多个分类
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "B.名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
           strSQL = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,Null as 类别,A.名称,A.说明,A.编者," & vbNewLine & _
                " Decode(b.名称, [5], 1, Decode(b.简码,[5],1,decode(a.编码,[5],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                " Decode(Substr(b.名称, 1, Length([5])), [5], 1, Decode(Substr(b.简码, 1, Length([5])),[5],1,decode(Substr(a.编码, 1, Length([5])),[5],1,NULL))) As 排序3ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B, 疾病诊断科室 C, 疾病诊断科室 D" & _
                " Where A.ID=B.诊断ID And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And A.类别=2" & _
                " And B.码类=[4] And d.人员id(+) = [6] And c.科室id(+)=[7] And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                " And ( Nvl(A.适用范围,0) = 0 or  A.适用范围 = 1) " & _
                " And (" & strSQL & ")" & _
                " Order by 排序1ID, 排序2ID, 排序3ID,A.编码"
                '排序顺序：先是完全匹配(名称、简码、编码）、个人收藏、其次是科室收藏、然后是左匹配(名称、简码、编码）、最后是双向匹配
        Else
            'B-中医疾病编码
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "A.名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSQL = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIF(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]"
            End If
            strSQL = _
                "Select Distinct a.Id, a.Id As 项目id, a.编码, a.类别, a.附码, a.名称," & IIF(mint简码 = 0, "A.简码", "A.五笔码 as 简码") & ", a.说明," & _
                " Decode(a.名称, [5], 1, Decode(" & IIF(mint简码 = 0, "A.简码", "A.五笔码") & ",[5],1,decode(a.编码,[5],1,NULL))) As 排序1ID," & vbNewLine & _
                "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                "                Decode(Substr(a.名称, 1, Length([5])), [5], 1, Decode(Substr(" & IIF(mint简码 = 0, "A.简码", "A.五笔码") & ", 1, Length([5])),[5],1,decode(Substr(a.编码, 1, Length([5])),[5],1,NULL))) As 排序3ID" & vbNewLine & _
                "From 疾病编码目录 A, 疾病编码科室 C, 疾病编码科室 D" & vbNewLine & _
                "Where a.类别 = '" & IIF(strType = "", "B", strType) & "' And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.疾病id(+) = a.Id And" & vbNewLine & _
                "      d.疾病id(+) = a.Id And c.科室id(+)=[7] And d.人员id(+) = [6]" & vbNewLine & _
                IIF(str性别 <> "", " And (A.性别限制=[3] Or A.性别限制 is NULL)", "") & _
                " And ( Nvl(A.适用范围,0) = 0 or  A.适用范围 = 1) " & _
                " And (" & strSQL & ")" & _
                "Order By 排序1ID, 排序2ID, 排序3ID, 编码"
        End If
    Else
        If opt诊断(0).value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "B.名称 Like [2]" '输入汉字时,只匹配名称
            Else
                strSQL = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
            strSQL = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,Null as 类别,A.名称,A.说明,A.编者," & vbNewLine & _
                " Decode(b.名称, [5], 1, Decode(b.简码,[5],1,decode(a.编码,[5],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                " Decode(Substr(b.名称, 1, Length([5])), [5], 1, Decode(Substr(b.简码, 1, Length([5])),[5],1,decode(Substr(a.编码, 1, Length([5])),[5],1,NULL))) As 排序3ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B, 疾病诊断科室 C, 疾病诊断科室 D" & _
                " Where A.ID=B.诊断ID And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And A.类别=1" & _
                " And B.码类=[4] And d.人员id(+) = [6] And c.科室id(+)=[7] And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                " And ( Nvl(A.适用范围,0) = 0 or  A.适用范围 = 1) " & _
                " And (" & strSQL & ")" & _
                " Order by 排序1ID, 排序2ID, 排序3ID,A.编码"
                '排序顺序：先是完全匹配(名称、简码、编码）、个人收藏、其次是科室收藏、然后是左匹配(名称、简码、编码）、最后是双向匹配
        Else
            'D-ICD-10疾病编码
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "A.名称 Like [2]" '输入汉字时,只匹配名称
            Else
                strSQL = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIF(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]"
            End If
            strSQL = _
                "Select Distinct a.Id, a.Id As 项目id, a.编码, a.类别, a.附码, a.名称," & IIF(mint简码 = 0, "A.简码", "A.五笔码 as 简码") & ", a.说明," & _
                " Decode(a.名称, [5], 1, Decode(" & IIF(mint简码 = 0, "A.简码", "A.五笔码") & ",[5],1,decode(a.编码,[5],1,NULL))) As 排序1ID," & vbNewLine & _
                "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                "                Decode(Substr(a.名称, 1, Length([5])), [5], 1, Decode(Substr(" & IIF(mint简码 = 0, "A.简码", "A.五笔码") & ", 1, Length([5])),[5],1,decode(Substr(a.编码, 1, Length([5])),[5],1,NULL))) As 排序3ID" & vbNewLine & _
                "From 疾病编码目录 A, 疾病编码科室 C, 疾病编码科室 D" & vbNewLine & _
                "Where a.类别 = 'D' And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.疾病id(+) = a.Id And" & vbNewLine & _
                "      d.疾病id(+) = a.Id And c.科室id(+)=[7] And d.人员id(+) = [6]" & vbNewLine & _
                IIF(str性别 <> "", " And (A.性别限制=[3] Or A.性别限制 is NULL)", "") & _
                " And ( Nvl(A.适用范围,0) = 0 or  A.适用范围 = 1) " & _
                " And (" & strSQL & ")" & _
                "Order By 排序1ID, 排序2ID, 排序3ID, 编码"

        End If
    End If
    GetDiagSQL = strSQL
End Function

Private Sub vsDiag_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As PointAPI
    Dim str性别 As String, int诊断输入 As Integer
    Dim str类别 As String
    
    On Error GoTo errH
    
    With vsDiag
        If Col = col诊断 Or Col = col中医证候 Then
            If .EditText = "" Then
                If .TextMatrix(Row, col编码) <> "" And Col = col诊断 Then
                    .EditText = .Cell(flexcpData, Row, Col)
                Else
                    '中医证候则清除备份数据
                    If Col = col中医证候 Then
                        .Cell(flexcpData, Row, Col) = ""
                    End If
                End If
                If mblnReturn Then Call DiagEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiagEnterNextCell
            ElseIf Col = col诊断 And .TextMatrix(Row, col编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                strSQL = GetDiagSQL(Row, strInput, strSQL, str性别)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, strInput, str性别, mint简码 + 1, strInput, UserInfo.ID, mlng病人科室id)
                If rsTmp.RecordCount = 1 Then
                    Call SetDiagInput(Row, rsTmp, rsTmp!类别 & ""): .EditText = .Text
                Else
                    '允许在标准的名称前后输入附加信息
                    .TextMatrix(Row, col诊断) = .EditText
                    lbl诊断.Tag = "1"
                    mblnNoSave = True
                End If
            ElseIf Col = col诊断 And .TextMatrix(Row, col编码) <> "" And .Cell(flexcpData, Row, Col) <> "" And mblnFreeInput Then
                strInput = UCase(.EditText)
                strSQL = GetDiagSQL(Row, strInput, strSQL, str性别)
                On Error GoTo errH
                vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(opt诊断(0).value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1, strInput, UserInfo.ID, mlng病人科室id, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                    Cancel = True
                Else
                    If rsTmp Is Nothing Then
                        .TextMatrix(Row, col诊断) = .EditText
                        lbl诊断.Tag = "1"
                        mblnNoSave = True
                    Else
                         Call SetDiagInput(Row, rsTmp, rsTmp!类别 & ""): .EditText = .Text
                    End If
                End If
            Else
                int诊断输入 = Val(Mid(gstr诊断输入, 1, 1))
                If int诊断输入 = 0 Then int诊断输入 = 1
                
                If mstr性别 Like "*男*" Then
                    str性别 = "男"
                ElseIf mstr性别 Like "*女*" Then
                    str性别 = "女"
                End If
                                
                strInput = UCase(.EditText)
                
                strSQL = GetDiagSQL(Row, strInput, strSQL, str性别, IIF(Col = col诊断, "B", "Z"))
                If Col = col诊断 Then
                    If int诊断输入 = 1 And zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1, strInput, UserInfo.ID, mlng病人科室id)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = Nothing '自由录入时有多个匹配不进行选择
                        End If
                        If Not rsTmp Is Nothing Then str类别 = rsTmp!类别 & ""
                        Call SetDiagInput(Row, rsTmp, str类别): .EditText = .Text
                        If mblnReturn Then Call DiagEnterNextCell
                    Else
                        vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(opt诊断(0).value, "疾病诊断", "疾病编码"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str性别, mint简码 + 1, strInput, UserInfo.ID, mlng病人科室id, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                        If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                            Cancel = True
                        Else
                            '检查诊断输入方式
                            If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                                MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                                Cancel = True
                            ElseIf Not (rsTmp Is Nothing) Then
                                Call SetDiagInput(Row, rsTmp, rsTmp!类别 & ""): .EditText = .Text
                                If mblnReturn Then Call DiagEnterNextCell
                            Else
                                '没有匹配成功再次当成自由录入
                                If int诊断输入 = 1 Or (int诊断输入 = 3 And (rsTmp Is Nothing) And mint险类 = 0) Then
                                    Call SetDiagInput(Row, Nothing, str类别): .EditText = .Text
                                    If mblnReturn Then Call DiagEnterNextCell
                                Else
                                    Cancel = True
                                End If
                            End If
                        End If
                    End If
                ElseIf Col = col中医证候 Then
                    If opt诊断(0).value Then
                        '按诊断输入:先查是否有对应
                        If Set中医证候(Row, Val(.TextMatrix(Row, col诊断ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%", str性别, mint简码 + 1, strInput, UserInfo.ID, mlng病人科室id, "ColSet:列宽设置|说明,2400|悬浮提示|说明")
                    If blnCancel Then '无匹配输入时,按任意输入处理,取消不同
                        Cancel = True
                    Else
                        '检查诊断输入方式
                        If rsTmp Is Nothing And (int诊断输入 = 2 Or int诊断输入 = 3 And mint险类 <> 0) Then
                            MsgBox "没有找到与输入匹配的内容。", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set中医证候(Row, 0, rsTmp, rsTmp Is Nothing)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = col发病时间 Then
            If .EditText <> "" Then
                strInput = GetFullDate(.EditText)
                If IsDate(strInput) Then
                    .EditText = Format(strInput, "yyyy-MM-dd HH:mm")
                    mblnCancle = False
                Else
                    MsgBox "请输入正确的发病时间，例如：""2012-12-21 00:00""。"
                    Cancel = True
                End If
            End If
            If .EditText <> .TextMatrix(Row, Col) Then mblnNoSave = True: lbl诊断.Tag = "1"
        End If
        mblnCancle = Cancel
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowDiagFlag(ByVal lngDiag As Long)
'功能：显示诊断行所对应关联的医嘱标记
'参数：lngDiag=诊断表格当前行
    Dim strALL As String, strCurr As String
    Dim lng组ID As Long, i As Long
        
    With vsDiag
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col医嘱ID) <> "" Then
                strALL = strALL & "," & .TextMatrix(i, col医嘱ID)
                If i = lngDiag Then
                    strCurr = .TextMatrix(i, col医嘱ID)
                End If
            End If
        Next
        strALL = Mid(strALL, 2)
    End With
    
    With vsAdvice
        .Redraw = flexRDNone
        Set .Cell(flexcpPicture, .FixedRows, COL_诊断, .Rows - 1, COL_诊断) = Nothing
        
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then
                '每行都处理：隐藏行不会显示，一并给药非首行被OwnerDraw
                lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), .RowData(i))
                If InStr("," & strCurr & ",", "," & lng组ID & ",") > 0 Then
                    Set .Cell(flexcpPicture, i, COL_诊断) = img16.ListImages("诊断_当前").Picture
                ElseIf InStr("," & strALL & ",", "," & lng组ID & ",") > 0 Then
                    Set .Cell(flexcpPicture, i, COL_诊断) = img16.ListImages("诊断_关联").Picture
                End If
            End If
        Next
        
        .Cell(flexcpPictureAlignment, .FixedRows, COL_诊断, .Rows - 1, COL_诊断) = 4
        .Redraw = flexRDDirect
    End With
End Sub

Private Function GetDiagRow(ByVal lng医嘱ID As Long) As Long
'功能：返回指定的医嘱所关联到的诊断行
'参数：lng医嘱ID=医嘱的组ID
'返回：如果没有关联，则返回-1
    Dim i As Long
    
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, col诊断) <> "" Then
                If InStr("," & .TextMatrix(i, col医嘱ID) & ",", "," & lng医嘱ID & ",") > 0 Then
                    GetDiagRow = i: Exit Function
                End If
            End If
        Next
    End With
    
    GetDiagRow = -1
End Function

Private Sub SetDiagFlag(ByVal lngAdvice As Long, ByVal intFlag As Integer, Optional ByVal lngDiag As Long = -1)
'功能：设置指定医嘱行与当前诊断行关联，或者取消与诊断行的关联
'参数：lngAdvice=医嘱表格当前行
'      blnFlag=0-清除,1-标记
'      lngDiag=是否设置为与指定诊断行关联，没有指定时为-1表示与当前诊断行关联
    Dim lngBegin As Long, lngEnd As Long
    Dim str医嘱ID As String, lng组ID As Long
    Dim blnDo As Boolean, i As Long
    
    If vsAdvice.RowData(lngAdvice) = 0 Then Exit Sub
    '如果是已经发送的医嘱，只准关联，不准取消关联
    If (intFlag = 0 Or intFlag = 1 And Not vsAdvice.Cell(flexcpPicture, lngAdvice, COL_诊断) Is Nothing) And vsAdvice.TextMatrix(lngAdvice, COL_状态) = "8" Then Exit Sub
    
    With vsAdvice
        '关联数据设置
        lng组ID = IIF(Val(.TextMatrix(lngAdvice, COL_相关ID)) <> 0, Val(.TextMatrix(lngAdvice, COL_相关ID)), .RowData(lngAdvice))
        With vsDiag
            '首先取消当前医嘱行与任何诊断行的关联
            For i = 1 To .Rows - 1
                str医嘱ID = .TextMatrix(i, col医嘱ID)
                If InStr("," & str医嘱ID & ",", "," & lng组ID & ",") > 0 Then
                    str医嘱ID = Replace("," & str医嘱ID & ",", "," & lng组ID & ",", ",")
                    If Left(str医嘱ID, 1) = "," Then str医嘱ID = Mid(str医嘱ID, 2)
                    If Right(str医嘱ID, 1) = "," Then str医嘱ID = Mid(str医嘱ID, 1, Len(str医嘱ID) - 1)
                    .TextMatrix(i, col医嘱ID) = str医嘱ID
                    
                    If intFlag = 0 Then blnDo = True
                End If
            Next
            
            '再设置为与当前诊断行关联
            If intFlag = 1 Then
                If lngDiag = -1 Then lngDiag = .Row
                
                If .TextMatrix(lngDiag, col诊断) <> "" Then
                    str医嘱ID = .TextMatrix(lngDiag, col医嘱ID)
                    If InStr("," & str医嘱ID & ",", "," & lng组ID & ",") = 0 Then
                        str医嘱ID = str医嘱ID & "," & lng组ID
                        blnDo = True
                    End If
                    If Left(str医嘱ID, 1) = "," Then str医嘱ID = Mid(str医嘱ID, 2)
                    .TextMatrix(lngDiag, col医嘱ID) = str医嘱ID
                Else
                    '指定要关联的诊断行无诊断时，处理为无关联的效果
                    intFlag = 0
                End If
            End If
        End With
        
        '界面显示切换
        Call GetRowScope(lngAdvice, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If .RowData(i) <> 0 And Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                If intFlag = 1 Then
                    If lngDiag = vsDiag.Row Then
                        Set .Cell(flexcpPicture, i, COL_诊断) = img16.ListImages("诊断_当前").Picture
                    Else
                        Set .Cell(flexcpPicture, i, COL_诊断) = img16.ListImages("诊断_关联").Picture
                    End If
                    .Cell(flexcpPictureAlignment, i, COL_诊断) = 4
                ElseIf intFlag = 0 Then
                    Set .Cell(flexcpPicture, i, COL_诊断) = Nothing
                End If
            End If
        Next
        
        If blnDo Then
            lbl诊断.Tag = "1"
            mblnNoSave = True
        End If
    End With
End Sub

Private Function AdviceHaveDiag(ByVal lngAdvice As Long) As Long
'功能：判断指定行的医嘱是否已关联诊断行
'参数：lngAdvice=医嘱表格当前行
'返回：指定医嘱行所具体关联的诊断行，返回-1表示无关联
    Dim lng组ID As Long, i As Long
    
    AdviceHaveDiag = -1
    If vsAdvice.RowData(lngAdvice) = 0 Then Exit Function
    
    With vsAdvice
        lng组ID = IIF(Val(.TextMatrix(lngAdvice, COL_相关ID)) <> 0, Val(.TextMatrix(lngAdvice, COL_相关ID)), .RowData(lngAdvice))
    End With
    
    With vsDiag
        For i = 1 To .Rows - 1
            If InStr("," & .TextMatrix(i, col医嘱ID) & ",", "," & lng组ID & ",") > 0 Then
                AdviceHaveDiag = i: Exit Function
            End If
        Next
    End With
End Function

Private Sub SetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset, ByVal str类别 As String)
'功能：处理诊断项目的输入
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI, i As Long
    Dim blnDo As Boolean
    Dim strTmp As String
    Dim lng原诊断id As Long '0 表示新添加的诊断， 不为0表示修改诊断，lng原诊断id 的值就是修改前的 诊断ID或疾病ID
    Dim int诊断类型 As Integer
    
    On Error GoTo errH
    With vsDiag
        '检查是否允许修改
        For i = 1 To vsAdvice.Rows - 1
            If InStr("," & .TextMatrix(.Row, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_状态) = "8" Then
                MsgBox "该诊断对应的处方已发送，不能修改。", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '医技工作站调用,若诊断关联医嘱,存在医嘱的开嘱医生非当前操作员,则不允许修改诊断
            If mint场合 = 2 And InStr("," & .TextMatrix(.Row, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_开嘱医生) <> UserInfo.姓名 Then
                MsgBox "该诊断存在关联医嘱,且该医嘱非您下达，不能修改。", vbInformation, Me.Caption
                Exit Sub
            End If
        Next
        If Not rsInput Is Nothing Then
            int诊断类型 = IIF(.Cell(flexcpFontBold, lngRow, COL西医), 1, 11)
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    If .Rows > M_LNG_DIAGCOUNT Then Exit For
                    .AddItem "", lngRow + 1: lngRow = lngRow + 1
                    lng原诊断id = 0
                Else
                    lng原诊断id = Val(.TextMatrix(lngRow, col诊断ID))
                End If
                
                If InStr(.TextMatrix(lngRow, col诊断), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col诊断), InStrRev(.TextMatrix(lngRow, col诊断), "("))
                End If
                .TextMatrix(lngRow, col诊断) = Nvl(rsInput!名称) & strTmp
                
                '根据诊断确定疾病,或根据疾病确定诊断
                If opt诊断(0).value Then
                    .TextMatrix(lngRow, col诊断ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col诊断编码) = rsInput!编码 & ""
                    .TextMatrix(lngRow, col疾病ID) = ""
                    strSQL = "Select 疾病ID as ID From 疾病诊断对照 Where 诊断ID=[1]"
                Else
                    .TextMatrix(lngRow, col疾病ID) = rsInput!项目ID
                    .TextMatrix(lngRow, col疾病编码) = rsInput!编码 & ""
                    .TextMatrix(lngRow, col疾病类别) = str类别
                    .TextMatrix(lngRow, col疾病附码) = rsInput!附码 & ""
                    .TextMatrix(lngRow, col诊断ID) = ""
                    strSQL = "Select 诊断ID as ID From 疾病诊断对照 Where 疾病ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!项目ID))
                If Not rsTmp.EOF Then
                    If opt诊断(0).value Then
                        .TextMatrix(lngRow, col疾病ID) = Nvl(rsTmp!ID)
                        strSQL = "Select 编码,附码,类别 From 疾病编码目录 where id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!ID)))
                        If Not rsTmp.EOF Then
                            .TextMatrix(lngRow, col疾病编码) = rsTmp!编码 & ""
                            .TextMatrix(lngRow, col疾病类别) = rsTmp!类别 & ""
                            .TextMatrix(lngRow, col疾病附码) = rsTmp!附码 & ""
                        End If
                    Else
                        .TextMatrix(lngRow, col诊断ID) = Nvl(rsTmp!ID)
                        strSQL = "Select 编码 From 疾病诊断目录 where id=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp!ID)))
                        If Not rsTmp.EOF Then .TextMatrix(lngRow, col诊断编码) = rsTmp!编码 & ""
                    End If
                End If
                
                '中医根据疾病诊断参考取证候
                If .Cell(flexcpData, lngRow, col中医) = 1 Then
                    Call Set中医证候(lngRow, Val(.TextMatrix(lngRow, col诊断ID)))
                End If
                
                .TextMatrix(lngRow, col编码) = IIF(Not IsNull(rsInput!编码), rsInput!编码, "")
                .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
                
                .Cell(flexcpData, lngRow, col疑诊) = 0
                .Cell(flexcpForeColor, lngRow, col疑诊) = .GridColor
                
                .TextMatrix(lngRow, colICD码) = Nvl(rsInput!编码)
                
                '输入主/次要诊断后调用外挂接口
                If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
                    On Error Resume Next
                    If lngRow = .FixedRows Then
                        Call gobjPlugIn.DiagnosisEnter(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(.TextMatrix(lngRow, col诊断ID)), .TextMatrix(lngRow, col诊断), lng原诊断id, mint场合)
                        Call zlPlugInErrH(err, "DiagnosisEnter")
                    Else
                        Call gobjPlugIn.DiagnosisOtherEnter(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, Val(.TextMatrix(lngRow, col诊断ID)), .TextMatrix(lngRow, col诊断), lng原诊断id, mint场合)
                        Call zlPlugInErrH(err, "DiagnosisOtherEnter")
                    End If
                    err.Clear: On Error GoTo 0
                End If
                Call SetDiagType(lngRow, int诊断类型)
                rsInput.MoveNext
            Next
            
            Call SetDiagHeight
        Else
            .TextMatrix(lngRow, col诊断) = .EditText
            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
            
            .Cell(flexcpData, lngRow, col疑诊) = 0
            .Cell(flexcpForeColor, lngRow, col疑诊) = .GridColor
            
            .TextMatrix(lngRow, col编码) = ""
            .TextMatrix(lngRow, col诊断ID) = ""
            .TextMatrix(lngRow, col疾病ID) = ""
            .TextMatrix(lngRow, colICD码) = ""
        End If
        
        lbl诊断.Tag = "1"
        mblnNoSave = True
    End With
    'PASS 诊断录入
    If mblnPass Then
        zlPassDrags
    End If
    '将本次录入尚未关联诊断的医嘱与当前诊断关联
    If vsDiag.TextMatrix(lngRow, col诊断) <> "" Then
        blnDo = False
        
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .RowData(i) <> 0 And Val(.TextMatrix(i, COL_婴儿)) = cbo婴儿.ListIndex Then
                    If Val(.TextMatrix(i, COL_状态)) = 1 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                        If GetDiagRow(.RowData(i)) = -1 Then
                            Call SetDiagFlag(i, 1, lngRow)
                            blnDo = True
                        End If
                    End If
                End If
            Next
        End With
        
        If blnDo Then
            Call ShowDiagFlag(lngRow)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Set中医证候(ByVal lngRow As Long, ByVal lng诊断ID As Long, Optional ByVal rsInput As Recordset, Optional ByVal blnFreeInput As Boolean) As Boolean
'功能：中医根据疾病诊断参考取证候
'参数：rsInput-如果不为空，则输出指定的中药证候记录集
'      blnFreeInput  true - 自由录入
'返回：是否有对应关系
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strTmp As String
    
    With vsDiag
        '去掉已有的证候
        If InStr(.TextMatrix(lngRow, col诊断), "(") > 0 And InStr(.TextMatrix(lngRow, col诊断), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col诊断), 1, InStrRev(.TextMatrix(lngRow, col诊断), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col诊断)
        End If
        If blnFreeInput Then
            .TextMatrix(lngRow, col证候ID) = ""
            .TextMatrix(lngRow, col证候编码) = ""
            .TextMatrix(lngRow, col中医证候) = .EditText
            .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
            mblnNoSave = True
            Exit Function
        Else
            If rsInput Is Nothing Then
                If lng诊断ID <> 0 Then
                    strSQL = "Select Distinct a.证候序号 as ID,a.证候ID,a.证候名称,b.编码 as 证候编码" & _
                        " From 疾病诊断参考 A,疾病编码目录 B" & _
                        " Where a.证候ID=b.ID(+) And a.诊断ID=[1] And a.证候名称 is Not NULL" & _
                        " Order by a.证候序号"
                    vPoint = GetCoordPos(.hWnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = Nothing
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中医证候", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng诊断ID)
                    If Not rsTmp Is Nothing Then
                        .TextMatrix(lngRow, col证候ID) = Nvl(rsTmp!证候id)
                        .TextMatrix(lngRow, col证候编码) = Nvl(rsTmp!证候编码)
                        If Not IsNull(rsTmp!证候名称) Then
                            .TextMatrix(lngRow, col诊断) = strTmp
                            .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
                            .TextMatrix(lngRow, col中医证候) = Nvl(rsTmp!证候名称)
                            .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
                            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
                            mblnNoSave = True
                        End If
                        Set中医证候 = True
                    Else
                        If blnCancel Then
                            Set中医证候 = True
                            If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col中医证候)
                        Else
                            Set中医证候 = False
                        End If
                    End If
                Else
                    Set中医证候 = False
                End If
            Else
                .TextMatrix(lngRow, col证候ID) = Nvl(rsInput!项目ID)
                .TextMatrix(lngRow, col证候编码) = Nvl(rsInput!编码)
                .TextMatrix(lngRow, col诊断) = strTmp
                .Cell(flexcpData, lngRow, col诊断) = .TextMatrix(lngRow, col诊断)
                .TextMatrix(lngRow, col中医证候) = Nvl(rsInput!名称)
                .Cell(flexcpData, lngRow, col中医证候) = .TextMatrix(lngRow, col中医证候)
                If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col中医证候)
                mblnNoSave = True
            End If
        End If
    End With
End Function

Private Sub DiagEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiag
        '从下一单元开始循环搜索
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col诊断) To col增加
                If DiagCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col增加 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            If txt医嘱内容.Enabled And txt医嘱内容.Visible Then
                txt医嘱内容.SetFocus
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Function DiagCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    Dim i As Long
    
    With vsDiag
        If .ColHidden(lngCol) Then Exit Function
        If .TextMatrix(lngRow, col医嘱ID) <> "" Then
            If lngCol = col诊断 Then
                For i = 1 To vsAdvice.Rows - 1
                    If InStr("," & .TextMatrix(lngRow, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_状态) = "8" Then
                        Exit Function
                    End If
                    '医技工作站调用,若诊断关联医嘱,存在医嘱的开嘱医生非当前操作员,则不允许删除诊断
                    If mint场合 = 2 And InStr("," & .TextMatrix(lngRow, col医嘱ID) & ",", "," & vsAdvice.RowData(i) & ",") > 0 And vsAdvice.TextMatrix(i, COL_开嘱医生) <> UserInfo.姓名 Then
                        Exit Function
                    End If
                Next
            End If
        End If
        '必须先输入诊断
        If .TextMatrix(lngRow, col诊断) = "" Then
            If lngCol = col疑诊 Or lngCol = col增加 Or lngCol = col发病时间 Then
                Exit Function
            End If
        End If
        If lngCol = col编码 Then Exit Function
        '必须先输诊断再输证候
        If lngCol = col中医证候 Then
            If .TextMatrix(lngRow, col诊断) = "" Then Exit Function
            If .Cell(flexcpData, lngRow, col中医) <> 1 Then Exit Function
        End If
    End With
    DiagCellEditable = True
End Function

Private Sub GetAgentInfo()
'功能：读取代办人信息
    Dim rsTmp As ADODB.Recordset
    
    gstrSQL = "Select c.信息名, c.信息值" & vbNewLine & _
                "  From 病人信息从表 C" & vbNewLine & _
                "  Where c.就诊id = [2] And c.病人id = [1] And Instr(',代办人姓名,代办人身份证号,病人身份证号,',','||c.信息名||',')>0"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    AgentInfo.本次就诊已录入 = False
    AgentInfo.代办人姓名 = ""
    AgentInfo.代办人身份证号 = ""
    If rsTmp.EOF Then Exit Sub
    
    AgentInfo.本次就诊已录入 = True
    While Not rsTmp.EOF
        Select Case Nvl(rsTmp!信息名)
            Case "代办人姓名"
                AgentInfo.代办人姓名 = Nvl(rsTmp!信息值)
            Case "代办人身份证号"
                AgentInfo.代办人身份证号 = Nvl(rsTmp!信息值)
        End Select
        rsTmp.MoveNext
    Wend
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDiagHeight()
    vsDiag.Height = vsDiag.Rows * vsDiag.RowHeightMin + IIF(mbytSize = 0, 2, 12) * Screen.TwipsPerPixelY
    fra诊断.Height = vsDiag.Height + 4 * Screen.TwipsPerPixelY + IIF(mbytSize = 0, 50, 0)
    Call Form_Resize
End Sub



Private Sub vsAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    'Pass
    If Button = 2 Then
        With vsAdvice
            lngRow = .MouseRow
            If lngRow >= .FixedRows And lngRow <= .Rows - 1 Then
                If Not .RowHidden(lngRow) Then .Row = lngRow
            End If
        End With
    End If
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim blnDo As Boolean
    
    If Button = 2 Then
        If cbsMain Is Nothing Then Exit Sub
        Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If mblnPass = False Then
            blnDo = True
        Else
            blnDo = gobjPass.PassType = 2 Or gobjPass.PassType = 4 Or (gobjPass.PassType = 1 And gobjPass.PassVersion = "4.0")
            '门诊编辑界面菜单级数比医生站少一级
            If Not blnDo Then
                If gobjPass.zlPassCheck(mobjPassMap) Then
                    Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, objPopup.CommandBar, conMenu_Edit_MediAudit)
                End If
            End If
        End If
        If gobjPlugIn Is Nothing And blnDo And gobjDrugExplain Is Nothing Then Exit Sub '当弹出没有菜单项目时会显示一个空白小方块
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
    
End Sub

Private Sub ExePlugIn(ByVal strName As String)
'功能：执行外挂功能
    Dim lngID As String
    
    If CreatePlugInOK(p门诊医嘱下达, mint场合) Then
        With vsAdvice
            lngID = .RowData(.Row)
            If InStr(",1,2,", Val(.TextMatrix(.Row, COL_EDIT))) > 0 Then
                If MsgBox("当前选中的医嘱未保存，是否先保存？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                cbsMain.FindControl(, conMenu_Save, True, True).Execute: Exit Sub
            End If
        End With
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, p门诊医嘱下达, strName, mlng病人ID, mlng挂号ID, lngID, mlng前提ID, mint场合)
        Call zlPlugInErrH(err, "ExecuteFunc")
        err.Clear: On Error GoTo 0
    End If
End Sub

Private Function CheckInHosAdvice() As Boolean
'功能：检查当前病人是否存在有效的留观或住院医嘱
'参数：
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, intDays As Integer
    
    On Error GoTo errH
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                If .TextMatrix(i, COL_EDIT) = "1" And i <> .Row Then
                    If .TextMatrix(i, COL_操作类型) = "1" Or .TextMatrix(i, COL_操作类型) = "2" Then
                        CheckInHosAdvice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
     
    strSQL = "Select Trunc(Sysdate - 入院日期) 天数 From 病案主页 Where 病人id = [1] And 主页id = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    If rsTmp.RecordCount > 0 Then
        intDays = IIF(gint普通挂号天数 > gint急诊挂号天数, gint普通挂号天数, gint急诊挂号天数)
        If intDays = 0 Then intDays = 1
        If rsTmp!天数 <= intDays Then   '医嘱发送时再删除预约登记（71009）
            If MsgBox("存在旧的预约申请，产生新的预约申请时必须要删除旧的，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                CheckInHosAdvice = True
                Exit Function
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceRSAddRow(ByRef rsAdvice As ADODB.Recordset, ByVal i As Long)
'功能：根据指定行的信息复制产生医嘱记录集的一行数据
    Dim lng组ID As Long, lng疾病ID As Long, j As Long
    
    '获取疾病ID
    With vsDiag
        lng组ID = IIF(Val(vsAdvice.TextMatrix(i, COL_相关ID)) = 0, vsAdvice.RowData(i), Val(vsAdvice.TextMatrix(i, COL_相关ID)))
        For j = .FixedRows To .Rows - 1
            If InStr("," & .TextMatrix(j, col医嘱ID) & ",", "," & lng组ID & ",") > 0 Then
                lng疾病ID = Val(.TextMatrix(j, col疾病ID))
                Exit For
            End If
        Next
    End With
    
    With vsAdvice
        rsAdvice.AddNew
        rsAdvice!ID = .RowData(i)
        rsAdvice!相关ID = Val("" & .TextMatrix(i, COL_相关ID))
        rsAdvice!前提ID = mlng前提ID
        rsAdvice!病人来源 = 1
        rsAdvice!病人ID = mlng病人ID
        rsAdvice!挂号单 = mstr挂号单
        
        rsAdvice!婴儿 = Val("" & .TextMatrix(i, COL_婴儿))
        rsAdvice!姓名 = mstr姓名
        rsAdvice!性别 = mstr性别
        rsAdvice!年龄 = mint年龄
        rsAdvice!病人科室id = mlng病人科室id
        
        rsAdvice!序号 = .TextMatrix(i, COL_序号)
        rsAdvice!医嘱状态 = .TextMatrix(i, COL_状态)
        rsAdvice!医嘱期效 = 1
        rsAdvice!诊疗类别 = .TextMatrix(i, COL_类别)
        rsAdvice!诊疗项目ID = Val("" & .TextMatrix(i, COL_诊疗项目ID))
        rsAdvice!标本部位 = .TextMatrix(i, COL_标本部位)
        rsAdvice!检查方法 = .TextMatrix(i, COL_检查方法)
        
        rsAdvice!收费细目ID = Val("" & .TextMatrix(i, COL_收费细目ID))
        rsAdvice!天数 = Val("" & .TextMatrix(i, COL_天数))
        rsAdvice!单次用量 = Val("" & .TextMatrix(i, COL_单量))
        rsAdvice!总给予量 = Val("" & .TextMatrix(i, COL_总量))
        rsAdvice!医嘱内容 = .TextMatrix(i, col_医嘱内容)
        rsAdvice!医生嘱托 = .TextMatrix(i, COL_医生嘱托)
        rsAdvice!执行科室ID = Val("" & .TextMatrix(i, COL_执行科室ID))
        rsAdvice!执行频次 = .TextMatrix(i, COL_频率)
        rsAdvice!频率次数 = Val("" & .TextMatrix(i, COL_频率次数))
        rsAdvice!频率间隔 = Val("" & .TextMatrix(i, COL_频率间隔))
        rsAdvice!间隔单位 = .TextMatrix(i, COL_间隔单位)
        rsAdvice!执行时间方案 = .TextMatrix(i, COL_执行时间)
        rsAdvice!计价特性 = Val("" & .TextMatrix(i, COL_计价性质))
        rsAdvice!执行性质 = Val("" & .TextMatrix(i, COL_执行性质))
        rsAdvice!执行标记 = Val("" & .TextMatrix(i, COL_执行标记))
                    
        rsAdvice!可否分零 = Val("" & .TextMatrix(i, COL_可否分零))
        rsAdvice!紧急标志 = Val("" & .TextMatrix(i, COL_标志))
        rsAdvice!开始执行时间 = .TextMatrix(i, COL_开始时间)
        rsAdvice!开嘱科室id = Val("" & .TextMatrix(i, COL_开嘱科室ID))
        rsAdvice!开嘱医生 = .TextMatrix(i, COL_开嘱医生)
        rsAdvice!开嘱时间 = CDate(.TextMatrix(i, COL_开嘱时间))
        rsAdvice!摘要 = .Cell(flexcpData, i, COL_医生嘱托)
        rsAdvice!疾病id = lng疾病ID
        rsAdvice!EditState = Val(.TextMatrix(i, COL_EDIT)) '1-新增，2－修改
        rsAdvice!用药目的 = Val("" & .TextMatrix(i, COL_用药目的))     '1-预防,2-治疗
        rsAdvice!用药理由 = .TextMatrix(i, COL_用药理由)
        rsAdvice.Update
    End With
End Sub

Private Function zlPluginAdviceEnter(ByVal lngRow As Long) As Boolean
'功能：输入医嘱完成后调用外挂接口
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim rsAdvice As ADODB.Recordset
    
    Set rsAdvice = GetAdviceRs
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    For i = lngBegin To lngEnd
        Call AdviceRSAddRow(rsAdvice, i)
    Next
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    On Error Resume Next
    zlPluginAdviceEnter = gobjPlugIn.AdviceEnter(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, rsAdvice, mint场合)
    If err.Number <> 0 And zlPluginAdviceEnter = False Then zlPluginAdviceEnter = True
    Call zlPlugInErrH(err, "AdviceEnter")
    err.Clear: On Error GoTo 0
End Function

Private Function zlPluginAdviceSave() As Boolean
'功能：医嘱保存前调用外挂接口
    Dim i As Long
    Dim rsAdvice As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As ADODB.Recordset
    Set rsAdvice = GetAdviceRs
    With vsAdvice
        '医嘱录入完成之后直接保存可能导到当前这行医嘱没有调用到外挂 AdviceEditAfter此处再调用一次。
        i = .Row
        If .RowData(i) <> 0 And (Val(.TextMatrix(i, COL_EDIT)) = 1 Or Val(.TextMatrix(i, COL_EDIT)) = 2) Then
            Set rsTmp = zlDatabase.zlCopyDataStructure(rsAdvice)
            Call GetRowScope(i, lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                Call AdviceRSAddRow(rsTmp, i)
            Next
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        End If
        For i = .FixedRows To .Rows - 1
            '新增或修改的有效行(非空行)
            If .RowData(i) <> 0 And (.TextMatrix(i, COL_EDIT) = "2" Or .TextMatrix(i, COL_EDIT) = "1") Then
                Call AdviceRSAddRow(rsAdvice, i)
            End If
        Next
    End With
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    On Error Resume Next
    If Not (rsTmp Is Nothing) Then
        If rsTmp.RecordCount > 0 Then
            Call gobjPlugIn.AdviceEditAfter(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, rsTmp, mint场合)
            Call zlPlugInErrH(err, "AdviceEditAfter")
        End If
    End If
    zlPluginAdviceSave = gobjPlugIn.AdviceSave(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, rsAdvice, mint场合)
    If err.Number <> 0 And zlPluginAdviceSave = False Then zlPluginAdviceSave = True
    Call zlPlugInErrH(err, "AdviceSave")
    err.Clear: On Error GoTo 0
End Function

Private Sub zlPluginAdviceRowChange(ByVal lngRow As Long, Optional ByVal intType As Integer)
'功能：医嘱切换行后调用外挂接口
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim rsAdvice As ADODB.Recordset
    
    If Val(vsAdvice.RowData(lngRow)) = 0 Then Exit Sub
    Set rsAdvice = GetAdviceRs
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    For i = lngBegin To lngEnd
        Call AdviceRSAddRow(rsAdvice, i)
    Next
    If rsAdvice.RecordCount > 0 Then rsAdvice.MoveFirst
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    On Error Resume Next
    If intType = 0 Then
        Call gobjPlugIn.AdviceRowChange(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, rsAdvice, mint场合)
        Call zlPlugInErrH(err, "AdviceRowChange")
    Else
        Call gobjPlugIn.AdviceEditAfter(glngSys, p门诊医嘱下达, mlng病人ID, mlng挂号ID, rsAdvice, mint场合)
        Call zlPlugInErrH(err, "AdviceEditAfter")
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Function CheckBackNo(ByVal str挂号单 As String) As Boolean
'功能：检查是否已经退号，或者是否已经取消接诊
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select ID From 病人挂号记录 Where 执行状态 In (0, -1) And NO = [1] And 记录性质=1 And 记录状态=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", str挂号单)
    If rsTmp.RecordCount > 0 Then
        MsgBox "不能操作没有就诊的病人。", vbInformation, "门诊医嘱编辑"
        Exit Function
    End If
    
    strSQL = "Select ID From 门诊费用记录 Where 记录性质=4 And 记录状态=2 And NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl_AdviceSendCheck", str挂号单)
    If rsTmp.RecordCount > 0 Then
        MsgBox "不能操作已经退号的病人。", vbInformation, "门诊医嘱编辑"
        Exit Function
    End If
    CheckBackNo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Set安排时间(ByVal strType As String)
'功能：设置手术或输血医嘱的安排时间位置
'参数：strType=F-手术,K-输血

    If strType = "F" Then
        lbl安排时间.Caption = "手术时间"
        lbl安排时间.Top = lbl用法.Top
        txt安排时间.Top = txt用法.Top
    
        lbl安排时间.Left = lbl医嘱内容.Left
        txt安排时间.Left = txt用法.Left
    Else
        lbl安排时间.Caption = "输血时间"
        lbl安排时间.Top = lbl开始时间.Top
        txt安排时间.Top = txt开始时间.Top
    
        lbl安排时间.Left = lbl医生嘱托.Left
        txt安排时间.Left = cbo医生嘱托.Left
        
        'SetItemEditable中设置了显示“安排时间”就不显示“用法”
        lbl用法.Visible = True
        txt用法.Visible = True
        cmd用法.Visible = True
    End If
    
    cmd安排时间.Top = txt安排时间.Top + 30
    cmd安排时间.Left = txt安排时间.Left + txt安排时间.Width - cmd安排时间.Width - 30
End Sub


Private Sub SetCbo执行性质(ByVal bln含自备药 As Boolean, ByVal bln临床自管药 As Boolean)
    cbo执行性质.Clear
    
    If bln临床自管药 Then
        cbo执行性质.AddItem "1-自备药"
    Else
        cbo执行性质.AddItem "0-正常"
        If bln含自备药 Then cbo执行性质.AddItem "1-自备药"
        cbo执行性质.AddItem "2-离院带药"
    End If
End Sub


Private Sub txt用药理由_GotFocus()
    zlControl.TxtSelAll txt用药理由
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt用药理由_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt用药理由_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt用药理由.Text <> "" And vsAdvice.TextMatrix(vsAdvice.Row, COL_类别) <> "K" Then
            If ReasonSelect(txt用药理由.Text, 1) Then Exit Sub
        End If
        If SeekNextControl Then Call txt用药理由_Validate(False)
    End If
End Sub

Private Sub txt用药理由_Change()
    txt用药理由.Tag = "1"
End Sub

Private Sub txt用药理由_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(txt用药理由.Text) > 1000 Then
        MsgBox "输入内容不过超过 500 个汉字或 1000 个字符。", vbInformation, gstrSysName
        txt用药理由_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '更新数据
    Call AdviceChange
End Sub

Private Sub cboDruPur_Click()
    lbl用药目的.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboDruPur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SeekNextControl
    End If
End Sub

Private Sub SetRow标志图标(ByVal i As Long, Optional ByVal bytMode As Byte)
'功能：根据当前行的状态，设置标志列的图标显示
'参数：i=当前行
'      bytMode=一并给药时传入，0-根据传入行的状态处理组内首行，1-传入行是组内首行时才处理,2-只处理传入行
    Dim blnFirst As Boolean, lngRow As Long
    Dim int图标数 As Integer '医嘱内容上面的图标个数
    
    With vsAdvice
        '自由录入
        If Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 Then
             Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("自由").Picture
             .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
        Else
            blnFirst = True
            lngRow = i
            '一并给药的图标只显示在第一行(审核状态除外)
            If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 Then
                If bytMode > 0 Then
                    If bytMode = 1 Then
                        '判断传入行是组内首行时才处理(可能还没有设置给药途径)
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then blnFirst = False
                    Else
                        '只处理传入的行(取消一并时)
                    End If
                Else
                    '根据传入行的状态处理组内首行
                    lngRow = .FindRow(.TextMatrix(i, COL_相关ID), , COL_相关ID)
                End If
            End If
        
            If blnFirst Then
                If .TextMatrix(i, COL_标志) = "2" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("补录").Picture
                ElseIf .TextMatrix(i, COL_标志) = "1" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("紧急").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = Nothing
                End If
            
                '一并给药显示在第一行
                If Val(.TextMatrix(i, COL_状态)) < 2 Then   '新开或暂存的医嘱
                    Select Case Val(.TextMatrix(i, COL_审核状态))
                    '0-无需审核，1-待审核，2-审核通过，3-审核未通过
                        Case 1
                            If .TextMatrix(i, COL_类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1 Then
                                '用血医嘱审核图标单独显示(表明是有医生核对)
                                Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("核对").Picture
                            Else
                                Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                            End If
                        Case 2
                            If Not (.TextMatrix(i, COL_类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1) Then
                                Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                            End If
                        Case 3
                            Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                        Case 4, 5
                            If gbln血库系统 = False Then Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                        Case 7
                            Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("待签发").Picture
                    End Select
                End If
                                '处方审查系统
                If .TextMatrix(i, COL_处方审查状态) = "0" Then
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                ElseIf .TextMatrix(i, COL_处方审查状态) = "2" Or .TextMatrix(i, COL_处方审查结果) = "1" Then
                    '超时免审当作合格处理
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                ElseIf .TextMatrix(i, COL_处方审查结果) = "2" Then
                    ' 不合格
                    Set .Cell(flexcpPicture, lngRow, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                End If
                .Cell(flexcpPictureAlignment, lngRow, COL_F标志) = 4
            End If
            If Val(.TextMatrix(i, COL_签名否)) = 1 Then
                Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgSign.ListImages("签名").Picture
                int图标数 = 1
            End If
            
            If Val(.TextMatrix(i, COL_高危药品)) > 0 Then
                If .Cell(flexcpPicture, i, col_医嘱内容) Is Nothing Then
                    Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
                    int图标数 = 1
                Else
                    If .Cell(flexcpPicture, i, col_医嘱内容) <> frmIcons.imgQuestion.ListImages("高危药品").Picture Then
                        pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, i, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                        pictmp.PaintPicture frmIcons.imgQuestion.ListImages("高危药品").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                        Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                        int图标数 = 2
                    End If
                End If
            End If
            '危急值图标
            If mlng危急值ID > 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Or Val(.TextMatrix(i, COL_危急值ID)) > 0 Then
                If int图标数 = 0 Then
                    Set .Cell(flexcpPicture, i, col_医嘱内容) = frmIcons.imgQuestion.ListImages("危急值").Picture
                ElseIf int图标数 = 1 Then
                    pictmp.Cls
                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                    Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                    int图标数 = 2
                ElseIf int图标数 = 2 Then
                    pictmp.Cls
                    pictmp.Width = 720
                    pictmp.PaintPicture .Cell(flexcpPicture, i, col_医嘱内容), 0, 0, 480, pictmp.Height
                    pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, 480, 0, 240, pictmp.Height
                    Set .Cell(flexcpPicture, i, col_医嘱内容) = pictmp.Image
                    pictmp.Width = 480
                    int图标数 = 3
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowOrHideQuestion()
'功能：显示或隐藏抗菌用药审核未通过的说明
    Dim strMsg As String
    
    If lbl疑问.Caption <> "" Then lbl疑问.Caption = ""
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 3 Then
        strMsg = GetKSSAuditQuestion(Val(vsAdvice.RowData(vsAdvice.Row)))
        If strMsg <> "" Then lbl疑问.Caption = "审核反馈：" & strMsg
        
    End If
    pic疑问.Visible = lbl疑问.Caption <> ""
    
    Call Form_Resize
End Sub

Private Sub ReSet审核状态图标(ByVal lngRow As Long)
'功能：删除或修改一并给药中的一行医嘱后，重设组中首行的图标和抗菌药嘱行的审核状态
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim blnDo As Boolean
    
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    
    '如果有审核未通过的，删除后变为待审核
    With vsAdvice
        For i = lngBegin To lngEnd
            If gblnKSSStrict And UserInfo.用药级别 < Val(.TextMatrix(i, COL_抗菌等级)) And .TextMatrix(i, COL_标志) <> "1" Then
                .TextMatrix(i, COL_审核状态) = 1
            Else
                .TextMatrix(i, COL_审核状态) = ""   'CheckAdvice中会将一组的设置为相同
            End If
            If .TextMatrix(i, COL_审核状态) <> "" Then
                Set .Cell(flexcpPicture, lngBegin, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                .Cell(flexcpPictureAlignment, lngBegin, COL_F标志) = 4
                blnDo = True
            End If
        Next
        If blnDo = False Then Set .Cell(flexcpPicture, lngBegin, COL_F标志) = Nothing
    End With
End Sub


Private Sub SetMediInfoItem(ByVal bln显示单量 As Boolean, ByVal bln显示抗菌药物 As Boolean)
'功能：隐藏或显示药品相关的信息项目：首次用量，超量说明，用药目的，用药理由
    Dim lngHeight As Long, lngHeightOld As Long
    Dim bytHideType As Byte
        
    lngHeightOld = fraAdvice.Height
    lngHeight = cbo附加执行.Top + cbo附加执行.Height + 60
    
    If bln显示单量 Or bln显示抗菌药物 Then
        If bln显示抗菌药物 = False Then
            lngHeight = lngHeight + txt超量说明.Height + 90
        Else
            lngHeight = lngHeight + txt超量说明.Height + 90 + txt用药理由.Height + 90
        End If
    End If
    
    If lngHeightOld <> lngHeight Then
        fraAdvice.Height = lngHeight
        '处理滚动条bug
        fraAdvice.Tag = "不滚动"
        Call cbsMain_Resize
        '需要调用两次才能生效
        fraAdvice.Tag = ""
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
        Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub SetFontSize(ByVal bytSize As Byte)
'功能：进行界面字体的统一设置
'参数：bytSize  0-9号字体，1-12号字体
    Call SetPublicFontSize(Me, bytSize)
    Call zlControl.VSFSetFontSize(vsAdvice, IIF(bytSize = 0, 9, 12))
    Call zlControl.VSFSetFontSize(vsDiag, IIF(bytSize = 0, 9, 12))
    Call SetCtrlPos
End Sub

Private Sub SetCtrlPos()
'功能：设置控件位置,请注意所有控件位置的设置尽可能在此函数中设置
    Dim lngDistance1 As Long, lngDistance2 As Long
    Dim lngHeight As Long
    
    lngDistance1 = 30: lngDistance2 = 180
    cmdLastDiag.Top = lbl诊断.Top + lbl诊断.Height + 60
    Call SetCtrlPosOnLine(True, -1, lbl诊断, 60, cmdLastDiag)
    vsDiag.Left = cmdLastDiag.Left + cmdLastDiag.Width + lngDistance1
    opt诊断(0).Left = vsDiag.Left + vsDiag.Width + lngDistance1 * 2
    opt诊断(1).Left = opt诊断(0).Left
    fra诊断.Height = vsDiag.Top + vsDiag.Height + 120
    
    '垂直设置控件位置
    lbl开始时间.Left = 120
    '先设置左边第一行控件，其次设置右边控件位置，最后设置左边控件位置，目的为了设置医嘱内容高度
    '左边第一排，为了地位第一排右边控件位置
    Call SetCtrlPosOnLine(False, 0, lbl开始时间, lngDistance1, txt开始时间, -1 * cmd开始时间.Width, cmd开始时间, lngDistance2, chk紧急, lngDistance2, chk免试)
    lbl滴速.Left = chk免试.Left + chk免试.Width + 5 * lngDistance2
    
    '设置右边控件位置
    Call SetCtrlPosOnLine(True, 1, lbl滴速, lngDistance2, lbl医生嘱托, lngDistance2, lbl执行时间, lngDistance2, lbl执行科室, lngDistance2, lbl附加执行, lngDistance2, lbl超量说明)
    Call SetCtrlPosOnLine(False, 0, lbl滴速, lngDistance1, cbo滴速, lngDistance1, lbl滴速单位, lngDistance2, chkZeroBilling)
    Call SetCtrlPosOnLine(False, 0, lbl医生嘱托, lngDistance1, cbo医生嘱托, -1 * cmd医生嘱托.Width, cmd医生嘱托, lngDistance1, cmd常用嘱托)
    Call SetCtrlPosOnLine(False, 0, lbl执行时间, lngDistance1, cbo执行时间)
    Call SetCtrlPosOnLine(False, 0, lbl执行科室, lngDistance1, cbo执行科室, lngDistance2, lbl执行性质, lngDistance1, cbo执行性质)
    Call SetCtrlPosOnLine(False, 0, lbl附加执行, lngDistance1, cbo附加执行)
    Call SetCtrlPosOnLine(False, 0, lbl超量说明, lngDistance1, txt超量说明, -1 * cmdExcReason.Width, cmdExcReason, lngDistance1, cmdComExcReason)
    cbo执行时间.Width = cmd常用嘱托.Left + cmd常用嘱托.Width - cbo执行时间.Left
    cbo执行性质.Width = cmd常用嘱托.Left + cmd常用嘱托.Width - cbo执行性质.Left
    txt超量说明.Width = cmd医生嘱托.Left + cmd医生嘱托.Width - txt超量说明.Left - 80
    
    '设置左边控件位置
    txt医嘱内容.Height = txt开始时间.Height '为了中部对齐
    Call SetCtrlPosOnLine(True, 1, lbl开始时间, lngDistance2, lbl医嘱内容, lbl执行科室.Top - lbl医生嘱托.Top - lbl医生嘱托.Height, lbl分零, lngDistance2, lbl用法, -1 * lbl安排时间.Height, lbl安排时间, lngDistance2, lbl总量, lngDistance2 + 10, lbl用药目的)
    Call SetCtrlPosOnLine(False, 0, lbl医嘱内容, lngDistance1, txt医嘱内容, lngDistance1, cmdExt)
    cmdSel.Top = cmdExt.Top + cmdExt.Height + 120
    cmdSel.Left = cmdExt.Left
    Call SetCtrlPosOnLine(False, 0, cmdSel, lngDistance1, picHelp)
    txt医嘱内容.Height = cbo执行科室.Top + cbo执行科室.Height - txt医嘱内容.Top
    Call SetCtrlPosOnLine(False, 0, lbl分零, lngDistance1, cbo分零)
    Call SetCtrlPosOnLine(False, 0, lbl安排时间, -1 * lbl用法.Width, lbl用法, lngDistance1, txt安排时间, -1 * cmd安排时间.Width, cmd安排时间, -1 * txt用法.Width, txt用法, -1 * cmd用法.Width, cmd用法, lngDistance2, lbl频率, lngDistance1, txt频率, -1 * cmd频率.Width, cmd频率)
    txt频率.Width = txt医嘱内容.Width + txt医嘱内容.Left - txt频率.Left
    cmd频率.Left = txt医嘱内容.Width + txt医嘱内容.Left - cmd频率.Width
    Call SetCtrlPosOnLine(False, 0, lbl总量, lngDistance1, txt总量, lngDistance1, lbl总量单位, -1 * lbl天数.Width, lbl天数, -1 * (txt天数.Width + Me.TextWidth("字")), txt天数, lngDistance2, lbl单量, lngDistance1, txt单量, lngDistance1, lbl单量单位)
    lbl用药目的.Top = IIF(mbytSize = 0, lbl用药目的.Top - 50, lbl用药目的.Top)
    Call SetCtrlPosOnLine(False, 0, lbl用药目的, lngDistance1, cboDruPur, lngDistance2, lbl用药理由, lngDistance1, txt用药理由, -1 * cmdReason.Width, cmdReason, 0.5 * lngDistance1, cmd收藏用药理由)
    lbl单量单位.Left = txt医嘱内容.Width + txt医嘱内容.Left - lbl单量单位.Width
    txt单量.Width = lbl单量单位.Left - lngDistance1 - txt单量.Left
    
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        fraAdvice.Width = cmd常用嘱托.Left + cmd常用嘱托.Width + 500
        Me.Width = fraAdvice.Width + fraAdvice.Left
    End If
End Sub

Private Sub CheckDrugOutOfRange(ByVal lngRow As Long, ByVal sngDays As Single)
'功能：检查药品用量或天数是否超过允许的范围,并且弹出选择提示
'返回：选择是否继续
    Dim blnReturn As Boolean
    Dim strOld As String
    
    strOld = vsAdvice.TextMatrix(lngRow, COL_是否超期)

    If sngDays > IIF(mbytPatiType = 1, conOrdinary, conEmergency) And vsAdvice.TextMatrix(lngRow, COL_超量说明) = "" Then
        vsAdvice.TextMatrix(lngRow, COL_是否超期) = "1"
    Else
        vsAdvice.TextMatrix(lngRow, COL_是否超期) = ""
    End If
    If strOld <> vsAdvice.TextMatrix(lngRow, COL_是否超期) Then
        lbl超量说明.Tag = "1"
    End If
End Sub


Private Sub Set用药天数是否超期(ByVal lngRow As Long)
'功能：根据给定医嘱行的总量、单量和频率信息计算天数，并且设置“是否超期”列的值
    Dim sng天数 As Single
    
    With vsAdvice
        If RowIn配方行(lngRow) And Val(.TextMatrix(lngRow, COL_总量)) > 0 Then
            sng天数 = Val(.TextMatrix(lngRow, COL_总量))
        Else
            If mbln天数 Then
                sng天数 = Val(.TextMatrix(lngRow, COL_天数))
            ElseIf Val(.TextMatrix(lngRow, COL_总量)) <> 0 And Val(.TextMatrix(lngRow, COL_单量)) <> 0 _
                And .TextMatrix(lngRow, COL_频率) <> "" And Val(.TextMatrix(lngRow, COL_频率次数)) <> 0 And Val(.TextMatrix(lngRow, COL_频率间隔)) <> 0 _
                And Val(.TextMatrix(lngRow, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngRow, COL_门诊包装)) <> 0 Then
                
                sng天数 = Calc缺省药品天数(Val(.TextMatrix(lngRow, COL_总量)), Val(.TextMatrix(lngRow, COL_单量)), _
                    Val(.TextMatrix(lngRow, COL_频率次数)), Val(.TextMatrix(lngRow, COL_频率间隔)), .TextMatrix(lngRow, COL_间隔单位), _
                    Val(.TextMatrix(lngRow, COL_剂量系数)), Val(.TextMatrix(lngRow, COL_门诊包装)), _
                    Val(.TextMatrix(lngRow, COL_可否分零)))
            End If
        End If
        
        If sng天数 > IIF(mbytPatiType = 1, conOrdinary, conEmergency) Then
            .TextMatrix(lngRow, COL_是否超期) = "1"
        Else
            .TextMatrix(lngRow, COL_是否超期) = ""
        End If
        
    End With
End Sub

Private Function ReGet药品总量(ByVal dbl原总量 As Double, ByVal dbl单量 As Double, ByVal sng天数 As Long, ByVal lngRow As Long) As Double
'功能：重新根据单量、天数及当前行的频率信息等计算总量
'返回：计算的总量
    Dim dbl总量 As Double
    
    ReGet药品总量 = dbl原总量
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_频率) <> "" And Val(.TextMatrix(lngRow, COL_频率次数)) <> 0 And Val(.TextMatrix(lngRow, COL_频率间隔)) <> 0 _
            And dbl单量 <> 0 And Val(.TextMatrix(lngRow, COL_剂量系数)) <> 0 And Val(.TextMatrix(lngRow, COL_门诊包装)) <> 0 Then
            
            dbl总量 = FormatEx(Calc缺省药品总量(dbl单量, sng天数, _
                Val(.TextMatrix(lngRow, COL_频率次数)), Val(.TextMatrix(lngRow, COL_频率间隔)), _
                .TextMatrix(lngRow, COL_间隔单位), .TextMatrix(lngRow, COL_执行时间), _
                Val(.TextMatrix(lngRow, COL_剂量系数)), Val(.TextMatrix(lngRow, COL_门诊包装)), _
                Val(.TextMatrix(lngRow, COL_可否分零))), 5)
                
            
            If InStr(GetInsidePrivs(p门诊医嘱下达), "药品小数输入") = 0 Then
                dbl总量 = IntEx(dbl总量)
            ElseIf Val(.TextMatrix(lngRow, COL_可否分零)) <> 0 Then
                dbl总量 = IntEx(dbl总量)
            End If
        End If
    End With
    ReGet药品总量 = dbl总量
End Function

Private Sub Set医嘱超量(ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：为指定范围内的医嘱设置超标记 .TextMatrix(i, COL_是否超量) .TextMatrix(i, COL_超量说明)
    Dim dbl总量 As Double
    Dim lng相关ID As Long
    Dim i As Long
    Dim j As Long
    
    With vsAdvice
        For i = lngBegin To lngEnd
            If .RowData(i) = 0 Then Exit Sub
            .TextMatrix(i, COL_是否超量) = ""
            If InStr(",5,6,", .TextMatrix(i, COL_类别)) > 0 And Val(.TextMatrix(i, COL_处方限量)) > 0 Then
                dbl总量 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_门诊包装)) * Val(.TextMatrix(i, COL_剂量系数))
                If dbl总量 > Val(.TextMatrix(i, COL_处方限量)) Then .TextMatrix(i, COL_是否超量) = "1"
                If .TextMatrix(i, COL_是否超量) = "" And .TextMatrix(i, COL_是否超期) = "" Then .TextMatrix(i, COL_超量说明) = ""
            ElseIf Val(.TextMatrix(i, COL_类别)) = 7 And Val(.TextMatrix(i, COL_处方限量)) > 0 Then  '中药
                dbl总量 = Val(.TextMatrix(i, COL_单量)) * Val(.TextMatrix(i, COL_总量))
                If dbl总量 > Val(.TextMatrix(i, COL_处方限量)) Then .TextMatrix(i, COL_是否超量) = "1"
            ElseIf .TextMatrix(i, COL_类别) = "E" And .TextMatrix(i, COL_操作类型) = "4" Then
                lng相关ID = .RowData(i)
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        If .TextMatrix(j, COL_是否超量) = "1" Then
                            .TextMatrix(i, COL_是否超量) = "1"
                            .TextMatrix(j, COL_是否超量) = ""
                        End If
                    Else
                        Exit For
                    End If
                Next
            ElseIf Val(.TextMatrix(i, COL_处方限量)) > 0 Then '其他诊疗项目
                If Val(.TextMatrix(i, COL_总量)) > Val(.TextMatrix(i, COL_处方限量)) Then .TextMatrix(i, COL_是否超量) = "1"
            End If
            If .TextMatrix(i, COL_是否超量) = "" And .TextMatrix(i, COL_是否超期) = "" Then .TextMatrix(i, COL_超量说明) = ""
        Next
    End With
End Sub

Private Function GetInsureStr(ByRef strIDs1 As String, ByRef strIDs2 As String, ByRef str医嘱内容 As String, ByVal lngRow As Long) As Boolean
'功能：获取医保对码的字符串
'   strIDs1:药品卫材的收费细目ID字符串，strIDs2 ：其他诊疗项目的诊疗项目ID:执行科室字符串
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Call GetRowScope(lngRow, lngBegin, lngEnd)
    With vsAdvice
        '为利用索引,用Union方式
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 Then
                If InStr(",4,5,6,7,", .TextMatrix(i, COL_类别)) > 0 Then
                    '药品、卫材无对应关系,药品只处理按规格下达时
                    If Val(.TextMatrix(i, COL_收费细目ID)) <> 0 And InStr("," & strIDs1 & ",", "," & Val(.TextMatrix(i, COL_收费细目ID)) & ",") = 0 Then
                        strIDs1 = strIDs1 & "," & .TextMatrix(i, COL_收费细目ID)
                    End If
                ElseIf InStr("," & strIDs2 & ",", "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & ",") = 0 Then
                    '包含了收费数量为0的
                    strIDs2 = strIDs2 & "," & Val(.TextMatrix(i, COL_诊疗项目ID)) & ":" & Val(.TextMatrix(i, COL_执行科室ID))
                End If
            End If
        Next
        str医嘱内容 = Left(vsAdvice.TextMatrix(lngRow, col_医嘱内容), 50)
    End With
End Function

Private Function SetAll超量说明(ByVal lngRow As Long, ByRef blnOut As Boolean) As Boolean
'功能：检查没有写超量说明的医嘱然后为其添加超量说明
'参数：lngRow开始行号
'      blnOut 表示弹出了输入界面但没填写任何内容
    Dim i As Long
    Dim strTmp As String
    Dim str配方 As String
    Dim str行IDs As String '需要填写超量说明的医嘱行号
    Dim str超量说明 As String
    Dim strMsg As String
    Dim varArr As Variant
    
    With vsAdvice
        For i = lngRow To .Rows - 1
            If (.TextMatrix(i, COL_是否超量) = "1" Or .TextMatrix(i, COL_是否超期) = "1") And .TextMatrix(i, COL_超量说明) = "" Then
                Select Case (Val(.TextMatrix(i, COL_是否超量)) - Val(.TextMatrix(i, COL_是否超期)))
                    Case 1
                        strTmp = "超出了处方限量，请填写超量说明。"
                    Case -1
                        strTmp = "超出了用药疗程(" & IIF(mbytPatiType = 1, conOrdinary, conEmergency) & "天)，请填写超量说明。"
                    Case 0
                        strTmp = "超出了处方限量和用药疗程(" & IIF(mbytPatiType = 1, conOrdinary, conEmergency) & "天)，请填写超量说明。"
                End Select
                If .TextMatrix(i, COL_类别) = "E" And .TextMatrix(i, COL_操作类型) = "4" Then '中草药的超量标志放到了显示行，只能大概提示
                    str配方 = .TextMatrix(i, col_医嘱内容)
                    str配方 = "中草药配方：" & Mid(str配方, InStr(str配方, ":") + 1)
                    strMsg = strMsg & """" & str配方 & """" & strTmp
                Else
                    strMsg = strMsg & """" & .TextMatrix(i, col_医嘱内容) & """" & strTmp
                End If
                str行IDs = str行IDs & "," & i
            End If
        Next
    End With
    str行IDs = Mid(str行IDs, 2)
    
    If strMsg <> "" And str行IDs <> "" Then
        Call frmMsgDruExcess.ShowMe(Me, strMsg, str超量说明)
    Else
        SetAll超量说明 = True
        Exit Function
    End If
    
    If str超量说明 = "*NULL*" Then 'str超量说明 = "*NULL*" 表示没有填写任何东西
        blnOut = True
    ElseIf str超量说明 <> "" Then  'str超量说明 = "" 表示没有执行 frmMsgDruExcess.ShowMe(Me, strMsg, str超量说明)
        varArr = Split(str行IDs, ",")
        For i = 0 To UBound(varArr)
            vsAdvice.TextMatrix(Val(varArr(i)), COL_超量说明) = str超量说明
        Next
    End If
    SetAll超量说明 = True
End Function

Private Function CheckAutoMerge(ByVal lngRow As Long) As Integer
'功能：自动判断溶媒进行自动一并给药和取消一并
        '判断溶媒，自动一并/取消一并
        '如果上一行医嘱是输液药品，当前输入的也是输液药品时：
        '1、如果当前录入的是药品，如果前面一组液体中，溶媒是第一个，那么自动为“一并”状态。
        '2、如果当前录入的是药品，如果前面一组液体中，溶媒是最后一个（上一条医嘱），那么自动取消“一并”状态（后输入溶媒的情况）。
        '3、如果当前行录入的溶媒（药品特性.溶媒），则判断如果一组液体中已经存在一个溶媒时，则自动取消“一并”状态（先输入溶媒的情况）。
'参数：lngRow=当前行
'返回：0-不做处理，1-一并给药，2-取消一并给药
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long
    
    With vsAdvice
        lngPreRow = GetPreRow(lngRow) '取上一有效行,某些内容缺省与上一行相同
        If lngPreRow <> -1 Then
            If InStr(",5,6,", .TextMatrix(lngPreRow, COL_类别)) > 0 Then
                i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_相关ID)), lngPreRow + 1)
                If .TextMatrix(i, COL_类别) = "E" And .TextMatrix(i, COL_操作类型) = "2" And .TextMatrix(i, COL_执行分类) = "1" Then
                    j = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1)
                    If .TextMatrix(j, COL_类别) = "E" And .TextMatrix(j, COL_操作类型) = "2" And .TextMatrix(j, COL_执行分类) = "1" Then
                        If RowCanMerge(lngPreRow, lngRow) Then
                            '获取上一条医嘱的开始行号
                            Call GetRowScope(i, lngBegin, lngEnd)
                            If mblnRowMerge = False Then
                                '第一行是溶媒(1)
                                If Val(.TextMatrix(lngBegin, COL_是否溶媒)) = 1 And Val(.TextMatrix(lngRow, COL_是否溶媒)) <> 1 Then
                                    CheckAutoMerge = 1
                                End If
                            Else
                                '当前行是溶媒，前面已经存在溶媒了(3)
                                If Val(.TextMatrix(lngRow, COL_是否溶媒)) = 1 Then
                                    For i = lngBegin To lngEnd
                                        If Val(.TextMatrix(i, COL_是否溶媒)) = 1 Then
                                            CheckAutoMerge = 2
                                            Exit For
                                        End If
                                    Next
                                Else
                                    '当前是药品，上一行是溶媒，但又不是第一行时(2)
                                    If Val(.TextMatrix(lngPreRow, COL_是否溶媒)) = 1 And lngPreRow <> lngBegin Then
                                        CheckAutoMerge = 2
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
    If mblnNoRefresh Then Exit Sub
    If mblnNoSave Then
        mblnNoRefresh = True
        tbcSub.Item(0).Selected = True
        MsgBox "当前医嘱内容编辑后尚未保存，请先保存。", vbInformation, gstrSysName
        mblnNoRefresh = False
        Exit Sub
    End If
    Call SubWinRefreshData(Item)
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新数据
    If objItem.Tag = "医嘱编辑" Then
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
    Else
        If objItem.Tag <> "" Then
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.RefreshForm(glngSys, p住院医嘱下达, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng病人ID, mstr挂号单, 0, False, _
                    0, 0, 0, mlng病人科室id)
                Call zlPlugInErrH(err, "RefreshForm")
                err.Clear: On Error GoTo 0
            End If
        End If
    End If
End Sub

Private Function CanUseApply(ByVal str类别 As String, Optional ByVal lng项目id As Long, Optional ByVal str项目编码 As String) As Boolean
'功能：是否可以使用相应的申请单
    '只用在医生工作站
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
        Dim str编码 As String
    Dim strEx As String
 
    If mint场合 <> 0 Then Exit Function
    
    If str类别 = "D" And Val(Mid(gstrOutUseApp, 1, 1)) = 1 Or _
        str类别 = "C" And Val(Mid(gstrOutUseApp, 2, 1)) = 1 Or _
        str类别 = "K" And Val(Mid(gstrOutUseApp, 3, 1)) = 1 Then
        blnTmp = True
    End If
    
    If lng项目id <> 0 And blnTmp Then
        If str类别 = "D" Then
            strSQL = "select 1 from 病历单据应用 where 诊疗项目ID=[1] and 应用场合 = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, 1)
            If rsTmp.EOF Then blnTmp = False
        ElseIf str类别 = "C" Then
            blnTmp = False
            If Not gobjLIS Is Nothing Then
                If str项目编码 <> "" Then
                    str编码 = str项目编码
                Else
                    strSQL = "select b.编码 from 诊疗项目目录 b where b.id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id)
                    str编码 = rsTmp!编码 & ""
                End If
                On Error Resume Next
                blnTmp = gobjLIS.CanUseLISApp(str编码, strEx)
                err.Clear: On Error GoTo 0
            End If
        End If
    End If
    CanUseApply = blnTmp
End Function

Private Function CheckApply() As Boolean
'功能：判断申请单的填写情况
    Dim i As Long
    If Val(gstrOutUseApp) > 0 Then
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    If Val(.TextMatrix(i, COL_申请序号)) < 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        .Row = i
                        Call cmdExt_Click
                    End If
                End If
            Next
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    If Val(.TextMatrix(i, COL_申请序号)) < 0 And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    CheckApply = True
End Function

Private Function ApplySelect() As Integer
'功能：弹出申请单选择器
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer, blnCancel As Boolean
    Dim vRect As RECT
    
    If mint场合 <> 0 Then Exit Function
 
    intTmp = Val(Mid(gstrOutUseApp, 1, 1))
    If intTmp = 1 Then strSQL = "select 1 as id,'检查申请单' as 申请单 from dual"
    intTmp = Val(Mid(gstrOutUseApp, 2, 1))
    If intTmp = 1 And Not gobjLIS Is Nothing Then strSQL = IIF(strSQL = "", "", strSQL & " union all ") & "select 2 as id,'检验申请单' as 申请单 from dual"
    intTmp = Val(Mid(gstrOutUseApp, 3, 1))
    If intTmp = 1 Then strSQL = IIF(strSQL = "", "", strSQL & " union all ") & "select 3 as id,'输血申请单' as 申请单 from dual"
    
    If strSQL = "" Then Exit Function
    
    strSQL = strSQL & " order by id"

    On Error GoTo errH
    vRect = GetControlRect(txt医嘱内容.hWnd)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, Me.Caption, , , , , , True, vRect.Left, vRect.Top, txt医嘱内容.Height, blnCancel, , True)
    If Not rsTmp Is Nothing Then ApplySelect = Val(rsTmp!ID & "")
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceInput申请单(ByVal intType As Integer)
'功能：输入诊疗项目后弹出申请单
    Dim str类别 As String
    Dim lngRow As Long
    Dim blnOK As Boolean
    Dim objAppPages() As clsApplicationData
    Dim rsCard As ADODB.Recordset
 
    On Error Resume Next
'    '新增
    cbsMain.FindControl(, conMenu_New, True, True).Execute
    lngRow = vsAdvice.Row
    Select Case intType
    Case 1 '检查申请
        blnOK = ApplyNew检查申请(0, "", objAppPages())
    Case 2 '检验申请
        blnOK = ApplyNew检验申请(0, "", rsCard)
    Case 3 '输血申请
        blnOK = ApplyNew输血申请(0, "", rsCard)
    End Select
    err.Clear

    On Error GoTo errH
    '根据选择项目设置缺省医嘱信息
    If blnOK Then
        Select Case intType
        Case 1 '检查申请
            Call AdviceSet检查申请(lngRow, objAppPages())
        Case 2 '检验申请
            Call AdviceSet检验申请(lngRow, rsCard)
        Case 3 '输血申请
            Call AdviceSet输血申请(lngRow, rsCard)
        End Select

        '显示已缺省设置的值
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        Call CalcAdviceMoney '显示新开医嘱金额
        '医保管控实时监测
        If mint险类 <> 0 And Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_EDIT)) = 0 Then
            '总量不可输入：缺省并固定总量的医嘱，以及长嘱
            '成套医嘱不在这里检查
            If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) And Not txt总量.Enabled Then
                If MakePriceRecord(vsAdvice.Row) Then
                    If Not gclsInsure.CheckItem(mint险类, 0, 0, mrsPrice) Then
                        Call AdviceCurRowClear: Exit Sub
                    End If
                End If
                '标记为已经作了检查
                vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_状态) = 1
            End If
        End If
        txt医嘱内容.SetFocus: Call SeekNextControl '必须先定位
        mblnNoSave = True
    Else
        '恢复原值(AdviceInput函数中可能处理了一下)
        txt医嘱内容.Text = vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容)
        txt医嘱内容.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetData检查申请(ByVal lngRow As Long, ByRef objAppPages() As clsApplicationData)
'功能：从医嘱表格中获取数据生成对象只传新开状态的医嘱和未保存的医嘱数据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAppend As String, strExtData As String
    Dim lng申请序号 As Long
    Dim lng相关ID As Long
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim objTmp As clsApplicationData
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngObjIndex As Long
    Dim str诊断 As String, str诊断IDs As String
 
    On Error GoTo errH
    With vsAdvice
        lngEnd = -1
        lngObjIndex = -1
        lng申请序号 = Val(.TextMatrix(lngRow, COL_申请序号))
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 And i > lngEnd Then
                lng相关ID = Val(.RowData(i))
                lngBegin = i
                lngRow = i
                Set objTmp = New clsApplicationData
                objTmp.blnIsModify = True
                objTmp.blnAllowUpdate = True
                objTmp.blnIsAdditionalRec = False
                objTmp.lngProjectId = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
                objTmp.blnIsPriority = Val(.TextMatrix(lngRow, COL_标志)) = 1
                objTmp.strStartExeTime = .TextMatrix(lngRow, COL_开始时间)
                objTmp.lngExeRoomId = Val(.TextMatrix(lngRow, COL_执行科室ID))
                objTmp.strExeRoomName = Get部门名称(objTmp.lngExeRoomId)
                objTmp.lngExeRoomType = Val(.TextMatrix(lngRow, COL_执行性质))
                objTmp.strRequestTime = .TextMatrix(lngRow, COL_开嘱时间)
                objTmp.lngRequestRoomId = Val(.TextMatrix(lngRow, COL_开嘱科室ID))
                
                If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
                    str诊断IDs = GetAdviceDiag(.RowData(lngRow), str诊断)
                    objTmp.strDiagnoseId = str诊断IDs
                End If
                
                strTmp = objTmp.Get申请单信息(objTmp.lngProjectId, 1)
                If strTmp <> "" Then
                    objTmp.lngApplicationPageId = Val(Split(strTmp, "<Split>")(0))
                    objTmp.strApplicationPageName = Split(strTmp, "<Split>")(1)
                    objTmp.strRequestAffixCfg = objTmp.Get申请附项目配置(objTmp.lngApplicationPageId)
                End If
                
                strExtData = Get检查部位方法(lngRow)
                If InStr(strExtData, vbTab) > 0 Then
                    objTmp.strPartMethod = Split(strExtData, vbTab)(0)
                    objTmp.lngExeType = Val(Split(strExtData, vbTab)(1))
                End If
                
                strAppend = .TextMatrix(lngRow, COL_附项)
                If strAppend <> "" Then
                    varArr = Split(strAppend, "<Split1>")
                    strAppend = ""
                    For j = 0 To UBound(varArr)
                        strTmp = varArr(j)
                        If strTmp <> "" Then
                            strAppend = strAppend & "|" & Split(strTmp, "<Split2>")(0) & ":" & Split(strTmp, "<Split2>")(3)
                        End If
                    Next
                    objTmp.strRequestAffix = Mid(strAppend, 2)
                End If
                
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng相关ID Then
                        lngEnd = j
                    Else
                        Exit For
                    End If
                Next
                lngObjIndex = lngObjIndex + 1
                ReDim Preserve objAppPages(lngObjIndex)
                Set objAppPages(lngObjIndex) = objTmp
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew检查申请(ByVal intType As Integer, ByVal str项目IDs As String, ByRef objAppPages() As clsApplicationData) As Boolean
'功能：调用申请单
'参数：intType 0-新增,1-修改
    Dim objPacspplication As New clsPacsApplication
    Dim lng项目id As Long
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    On Error GoTo errH
    lng项目id = Val(str项目IDs)
    '初始化检查申请单对象
    Call objPacspplication.InitComponents(mlng病人科室id, Me)
    ApplyNew检查申请 = objPacspplication.ShowApplicationForm(mlng病人ID, 1, mlng挂号ID, 0, intType, objAppPages(), , , lng项目id)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet检查申请(ByVal lngRow As Long, ByRef objAppPages() As clsApplicationData) As Long
'功能：把检查申请单数据加载到表格中
    Dim objTmp As clsApplicationData
    Dim i As Long, j As Long, k As Long
    Dim lngCnt As Long, lng申请序号 As Long, lngCurRow As Long
    Dim strSQL As String, strAppend As String
    Dim varTmp As Variant, strTmp As String
    Dim rsInput As ADODB.Recordset
    Dim rsAppend As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim str诊断IDs As String
    Dim str附项内容 As String
    
    On Error GoTo errH
    lng申请序号 = Get申请序号
    mblnRowChange = False
    For i = 0 To UBound(objAppPages)
        Set objTmp = objAppPages(i)
        strSQL = "Select a.名称,a.Id As 诊疗项目id, Null As 收费细目id,a.类别 As 类别ID From 诊疗项目目录 A where a.id=[1]"
        Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objTmp.lngProjectId)
        strAppend = ""
        If objTmp.strRequestAffix <> "" Then
            strSQL = "Select C.项目,C.内容,C.要素ID,C.必填,d.中文名,decode(D.表示法,4,D.数值域,NULL) as 数值域" & _
                " From 病历单据应用 A,病历文件列表 B,病历单据附项 C,诊治所见项目 D" & _
                " Where A.诊疗项目ID=[1] And A.应用场合=[2]" & _
                " And A.病历文件ID=B.ID And B.种类=7 And B.ID=C.文件ID And c.要素id=d.id(+)" & _
                " Order by C.排列"
            Set rsAppend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objTmp.lngProjectId, 1)
            varTmp = Split(objTmp.strRequestAffix, "|")
            For j = 0 To UBound(varTmp)
                If InStr(varTmp(j), ":") > 0 Then
                    rsAppend.Filter = "项目='" & Split(varTmp(j), ":")(0) & "'"
                    If Not rsAppend.EOF Then
                        strSQL = varTmp(j)
                        str附项内容 = Replace(strSQL, Mid(strSQL, 1, InStr(strSQL, ":")), "")
                        strAppend = IIF(strAppend = "", "", strAppend & "<Split1>") & rsAppend!项目 & "<Split2>" & Val(rsAppend!必填 & "") & _
                            "<Split2>" & rsAppend!要素ID & "<Split2>" & str附项内容
                    End If
                End If
            Next
        End If
        If i <> 0 Then
            vsAdvice.AddItem "", lngEnd + 1
            lngRow = lngEnd + 1
        End If
        str诊断IDs = objTmp.strDiagnoseId
        strTmp = objTmp.strPartMethod & vbTab & objTmp.lngExeType
        Call AdviceSet诊疗项目(rsInput, lngRow, 0, 0, strTmp, "")
        lngBegin = lngRow
        With vsAdvice
            For j = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(j, COL_相关ID)) = Val(.RowData(lngRow)) Then
                    lngEnd = j
                Else
                    Exit For
                End If
            Next
            If lngEnd < lngBegin Then lngEnd = lngBegin
            For j = lngBegin To lngEnd
                '更新一些项目
                .TextMatrix(j, COL_标志) = IIF(objTmp.blnIsPriority, 1, 0) '门诊无补录
                .TextMatrix(j, COL_开始时间) = Format(objTmp.strStartExeTime, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, j, COL_开始时间) = .TextMatrix(j, COL_开始时间)
                .TextMatrix(j, COL_执行科室ID) = IIF(objTmp.lngExeRoomId <= 0, 0, objTmp.lngExeRoomId)
                .TextMatrix(j, COL_执行性质) = objTmp.lngExeRoomType
                .TextMatrix(j, COL_开嘱时间) = Format(objTmp.strRequestTime, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, j, COL_开嘱时间) = .TextMatrix(j, COL_开嘱时间)
                .TextMatrix(j, COL_开嘱科室ID) = objTmp.lngRequestRoomId
                .TextMatrix(j, COL_申请序号) = lng申请序号
            Next
            .Cell(flexcpData, lngRow, COL_申请序号) = str诊断IDs
            '新增后关联诊断的标记处理
            Call SetDiagFlag(lngRow, 1)
            '更新附件内容:以当前可见行为准
            If strAppend <> "" Then
                .TextMatrix(lngRow, COL_附项) = strAppend
                .Cell(flexcpData, lngRow, COL_附项) = 1 '表明需要重新写入(新增或修改)
                Call ReplaceAdviceAppend(lngRow) '缺省替换其他医嘱的申请附项
            End If
            '重新自动调整行高
            Call .AutoSize(col_医嘱内容)
            Call SetRow标志图标(lngRow, 0)
        End With
    Next
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetData输血申请(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset)
'功能：从医嘱表格中获取数据生成对象只传新开状态的医嘱和未保存的医嘱数据
    Dim rsCardBak As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As New ADODB.Recordset
    Dim blnTmp As Boolean
    Dim lng医嘱ID As Long
    Dim strSQL As String
    Dim str诊断 As String
    Dim str诊断IDs As String
    Dim strTmp As String
    Dim var1 As Variant
    Dim var2 As Variant
    Dim i As Long, j As Long
    Dim str申请项目 As String, str输血目的 As String
    
    On Error GoTo errH
    With vsAdvice
        If TypeName(.Cell(flexcpData, lngRow, COL_申请序号)) = "Recordset" Then
            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_申请序号))
        Else
            Call InitCardRsBlood(rsCard)
            rsCard.AddNew
        End If
        If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            lng医嘱ID = Val(.RowData(lngRow))
            
            strSQL = "Select 诊疗项目id, 申请量, 申请血型, 申请rh,血液信息 From 输血申请项目 Where 医嘱id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "输血申请项目", lng医嘱ID)
            Do While Not rsTmp.EOF
                str申请项目 = str申请项目 & ";" & rsTmp!诊疗项目ID & "," & rsTmp!申请量 & "," & rsTmp!申请血型 & "," & rsTmp!申请rh & IIF(rsTmp!血液信息 & "" <> "", "," & rsTmp!血液信息, "")
            rsTmp.MoveNext
            Loop
            If Left(str申请项目, 1) = ";" Then str申请项目 = Mid(str申请项目, 2)
            
            strSQL = "Select 是否待诊 as 待诊,输血类型,输血目的, 输血性质, 即往输血史, 既往输血反应史, 输血禁忌及过敏史, 孕产情况, 受血者属地, 输血血型 as 血型, RHD" & vbNewLine & _
                " From 输血申请记录" & vbNewLine & _
                " Where 医嘱id = [1]"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            If Not rsTmp.EOF Then
                If Val(rsTmp!待诊 & "") = 0 Then
                   str诊断IDs = GetAdviceDiag(lng医嘱ID, str诊断)
                    '从附项中获取诊断如果附项中有以附项为准
                    strSQL = "select 内容 from 病人医嘱附件 where 医嘱ID=[1] and 项目='申请单诊断'"
                    Set rsTmpOther = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                    If Not rsTmpOther.EOF Then str诊断 = rsTmpOther!内容 & ""
                End If
                rsCard!临床诊断IDs = str诊断IDs
                rsCard!临床诊断描述 = str诊断
                rsCard!血型 = Val(rsTmp!血型 & "")
                rsCard!RHD = Val(rsTmp!RHD & "")
                rsCard!待诊 = Val(rsTmp!待诊 & "")
                rsCard!输血性质 = Val(rsTmp!输血性质 & "")
                rsCard!即往输血史 = Val(rsTmp!即往输血史 & "")
                rsCard!既往输血反应史 = Val(rsTmp!既往输血反应史 & "")
                rsCard!输血禁忌及过敏史 = Val(rsTmp!输血禁忌及过敏史 & "")
                rsCard!输血类型 = rsTmp!输血类型 & ""
                rsCard!输血目的 = rsTmp!输血目的 & ""
                str输血目的 = rsTmp!输血目的 & ""
                rsCard!孕产情况 = Val(rsTmp!孕产情况 & "")
                rsCard!受血者属地 = Val(rsTmp!受血者属地 & "")
            End If
            var2 = Array()
            strSQL = "select 序号,检验项目ID,指标代码,指标中文名,指标英文名,指标结果,结果单位,结果标志,结果参考,取值序列,是否人工填写 from 输血检验结果 Where 医嘱ID=[1] order by 序号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
            For i = 1 To rsTmp.RecordCount
                var1 = Array()
                For j = 0 To rsTmp.Fields.Count - 1
                    ReDim Preserve var1(j)
                    var1(j) = rsTmp.Fields(j).value & ""
                Next
                strTmp = Join(var1, "<SplitCol>")
                ReDim Preserve var2(UBound(var2) + 1)
                var2(UBound(var2)) = strTmp
                rsTmp.MoveNext
            Next
            rsCard!检查结果 = Join(var2, "<SplitRow>")
        End If
        
        rsCard!申请项目 = str申请项目
        If str输血目的 = "" Then rsCard!输血目的 = .TextMatrix(lngRow, COL_用药理由)
        rsCard!用血安排 = Val(.TextMatrix(lngRow, COL_标志))
        rsCard!预定输血日期 = .TextMatrix(lngRow, COL_输血时间)
        rsCard!输血项目ID = Val(.TextMatrix(lngRow, COL_诊疗项目ID))
        rsCard!输血执行科室ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        rsCard!预定输血量 = Val(.TextMatrix(lngRow, COL_总量))
        rsCard!输血途径项目ID = Val(.TextMatrix(lngRow + 1, COL_诊疗项目ID))
        rsCard!输血途径执行科室ID = Val(.TextMatrix(lngRow + 1, COL_执行科室ID))
        rsCard!备注 = .TextMatrix(lngRow, COL_医生嘱托)
        rsCard!输血申请日期 = .TextMatrix(lngRow, COL_开始时间)
        rsCard!申请科室ID = .TextMatrix(lngRow, COL_开嘱科室ID)
        rsCard!滴速 = .TextMatrix(lngRow + 1, COL_医生嘱托)
        rsCard.Update
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew输血申请(ByVal intType As Integer, ByVal str项目IDs As String, ByRef rsCard As ADODB.Recordset, Optional ByVal bln备血 As Boolean = True) As Boolean
'功能：调用申请单
'参数：intType 0-新增,1-修改
    Dim lng项目id As Long
    Dim lng开嘱科室ID As Long
    
    On Error GoTo errH
    lng项目id = Val(str项目IDs)
 
    '输血医嘱检查，必须中级及以上专业技术职务的医师才允许下达
    If gbln输血申请中级以上 Then
        If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
            MsgBox "启用了输血分级管理后，输血医嘱只有中级及以上专业技术职务医师才能下达。", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
    lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng医技科室ID, mlng病人科室id, 2)
    
    If Not rsCard Is Nothing Then
        If Not rsCard.EOF Then
            If Val(rsCard!申请科室ID & "") <> 0 Then
                lng开嘱科室ID = Val(rsCard!申请科室ID & "")
            End If
        End If
    End If
    
    If gbln血库系统 = True Then
        ApplyNew输血申请 = frmBloodApplyNew.ShowMe(Me, mlng病人ID, 0, 1, intType, 0, mlng病人科室id, _
             , lng开嘱科室ID, , , mrsDefine, , 1, mstr挂号单, , lng项目id, rsCard, mint婴儿, 1, IIF(bln备血 = True, 0, 1))
    Else
        ApplyNew输血申请 = frmBloodApply.ShowMe(Me, mlng病人ID, 0, 1, intType, 0, mlng病人科室id, _
             , lng开嘱科室ID, , , mrsDefine, , 1, mstr挂号单, , lng项目id, rsCard, mint婴儿, 1)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet输血申请(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset) As Long
'功能：把检查申请单数据加载到表格中
    Dim objTmp As clsApplicationData
    Dim i As Long, j As Long, k As Long
    Dim lngCnt As Long, lng申请序号 As Long, lngCurRow As Long
    Dim strSQL As String, strAppend As String
    Dim varTmp As Variant, strTmp As String
    Dim rsInput As ADODB.Recordset
    Dim rsAppend As ADODB.Recordset
    Dim lngBegin As Long, lngEnd As Long
    Dim arrItem, strIDs As String
    
    On Error GoTo errH
    lng申请序号 = Get申请序号
    mblnRowChange = False
    
    strSQL = "Select a.名称,a.Id As 诊疗项目id, Null As 收费细目id,a.类别 As 类别ID From 诊疗项目目录 A where a.id=[1]"
    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsCard!输血项目ID & ""))
    Call AdviceSet诊疗项目(rsInput, lngRow, Val(rsCard!输血途径项目ID & ""), 0, "", "")
    
    '备血申请可能申请多个品种，此处重新设置名称列
    arrItem = Split(rsCard!申请项目 & "", ";")
    strIDs = ""
    If UBound(arrItem) > 0 Then
        For i = 0 To UBound(arrItem)
            strIDs = strIDs & "," & Val(arrItem(i))
        Next
        strIDs = Mid(strIDs, 2)
        strSQL = "Select /*+ CARDINALITY(C 10) */" & vbNewLine & _
            "  f_List2str(Cast(Collect(a.名称) As t_Strlist)) 名称" & vbNewLine & _
            " From 诊疗项目目录 a, Table(f_Num2list([1])) b" & vbNewLine & _
            " Where a.Id = b.Column_Value"
        Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
        If rsInput.EOF = False Then
            vsAdvice.TextMatrix(lngRow, COL_名称) = rsInput!名称
        End If
    End If
    
    With vsAdvice
        For j = lngRow To lngRow + 1
            '更新一些项目
            .TextMatrix(j, COL_标志) = rsCard!用血安排
            .TextMatrix(j, COL_开嘱时间) = rsCard!输血申请日期 '", adVarChar, 500
            .TextMatrix(j, COL_开始时间) = rsCard!输血申请日期
            .Cell(flexcpData, j, COL_开始时间) = .TextMatrix(j, COL_开始时间)
            .Cell(flexcpData, j, COL_开嘱时间) = .TextMatrix(j, COL_开嘱时间)
            .TextMatrix(j, COL_开嘱科室ID) = Val(rsCard!申请科室ID & "")
            .TextMatrix(j, COL_申请序号) = lng申请序号
        Next
        .TextMatrix(lngRow, COL_输血时间) = rsCard!预定输血日期
        .TextMatrix(lngRow, COL_总量) = rsCard!预定输血量
        .TextMatrix(lngRow, COL_医生嘱托) = rsCard!备注
        .TextMatrix(lngRow, COL_用药理由) = rsCard!输血目的
        .TextMatrix(lngRow, COL_执行科室ID) = Val(rsCard!输血执行科室ID & "")
        .TextMatrix(lngRow + 1, COL_执行科室ID) = Val(rsCard!输血途径执行科室ID & "")
        .TextMatrix(lngRow + 1, COL_医生嘱托) = rsCard!滴速 & ""
        '新增后关联诊断的标记处理
        Call SetDiagFlag(lngRow, 1)
        Set .Cell(flexcpData, lngRow, COL_申请序号) = zlDatabase.CopyNewRec(rsCard)
        .TextMatrix(lngRow, col_医嘱内容) = AdviceTextMake(lngRow)
        '重新自动调整行高
        Call .AutoSize(col_医嘱内容)
    End With
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetData检验申请(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset)
'功能：从医嘱表格中获取数据生成对象只传新开状态的医嘱和未保存的医嘱数据
    Dim rsCardBak As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    Dim strSQL As String
    Dim str诊断 As String
    Dim str诊断IDs As String
    Dim strTmp As String
    Dim var1 As Variant
    Dim var2 As Variant
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strResult As String
    Dim lng申请序号 As Long
    Dim strAppend As String
    
    On Error GoTo errH
    With vsAdvice
    
        If TypeName(.Cell(flexcpData, lngRow, COL_申请序号)) = "Recordset" Then
            Set rsCard = zlDatabase.CopyNewRec(.Cell(flexcpData, lngRow, COL_申请序号))
        Else
            Call InitCardRsLIS(rsCard)
            rsCard.AddNew
        End If
        If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
            str诊断IDs = GetAdviceDiag(Val(.RowData(lngRow)), str诊断)
            rsCard!临床诊断IDs = str诊断IDs
        End If
     
        lng申请序号 = Val(.TextMatrix(lngRow, COL_申请序号))
        For i = .FixedRows To .Rows - 1
            '检验医嘱显示行是采集方式
            If Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 And Not .RowHidden(i) Then
         
                lngRow = i
                '采诊科室id , 执行科室id, 申请时间1, 标本1, 附项, 嘱托, 是否急症, 采集id, 诊疗项目id1
                strLIS = .TextMatrix(lngRow, COL_执行科室ID) & "<Split A>" & .TextMatrix(lngRow - 1, COL_执行科室ID) & "<Split A>" & .TextMatrix(lngRow, COL_开始时间) & _
                      "<Split A>" & .TextMatrix(lngRow, COL_标本部位)
                strAppend = .TextMatrix(lngRow, COL_附项)
                If strAppend = "" And Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 Then
                     strAppend = Get病人医嘱附件(.RowData(lngRow))
                End If
                strLIS = strLIS & "<Split A>" & strAppend & "<Split A>" & .TextMatrix(lngRow, COL_医生嘱托) & "<Split A>" & Val(.TextMatrix(lngRow, COL_标志)) & "<Split A>" & .TextMatrix(lngRow, COL_诊疗项目ID) & "<Split A>" & .TextMatrix(lngRow - 1, COL_诊疗项目ID)
                If strLIS <> "" Then
                    strResult = strLIS & IIF(strResult = "", "", "<Split B>" & strResult)
                End If
            End If
        Next
        rsCard!申请信息 = strResult
        rsCard.Update
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ApplyNew检验申请(ByVal intType As Integer, ByVal str项目编码 As String, ByRef rsCard As ADODB.Recordset) As Boolean
'功能：调用申请单
'参数：intType 0-新增,1-修改
    Dim lng开嘱科室ID As Long
    Dim strSQL As String
    Dim rsPati As ADODB.Recordset
    Dim strResult As String
    Dim strDept As String
    Dim lng申请序号 As String
    Dim strDiag As String
    Dim blnCancel As Boolean
    Dim strErr As String
    Dim strLIS As String

    On Error GoTo errH
    
    '执行部门(号别科室)即病人科室
    strSQL = "Select A.姓名,A.性别,A.年龄,B.门诊号,B.住院号,B.健康号,a.ID as 挂号ID," & _
        " B.险类,B.就诊诊室,C.名称 as 执行部门,A.登记时间" & _
        " From 病人挂号记录 A,病人信息 B,部门表 C" & _
        " Where A.NO(+)=[2] And a.记录性质(+)=1 And a.记录状态(+)=1 And B.病人ID=[1]" & _
        " And A.病人ID(+)=B.病人ID And A.执行部门ID=C.ID(+)"
 
    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
    
    If rsPati.RecordCount = 0 Then
        MsgBox "未能正确读取病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strDept = Get部门名称(mlng病人科室id)
    Call InitObjLis(p门诊医嘱下达)
    If gobjLIS Is Nothing Then Exit Function
    Call CreatePlugInOK(p门诊医嘱下达, mint场合)
    On Error GoTo errH
    If Not rsCard Is Nothing Then
        If Not rsCard.EOF Then
            strLIS = rsCard!申请信息 & ""
            strDiag = rsCard!临床诊断IDs & ""
        End If
    End If
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, 0, mlng病人ID, 0, Val("" & rsPati!挂号ID), rsPati!姓名, "" & rsPati!性别, "" & rsPati!年龄, 1, _
        Val("" & rsPati!门诊号), Val("" & rsPati!住院号), Val("" & rsPati!健康号), strDiag, UserInfo.姓名, UserInfo.部门ID, UserInfo.部门名, mlng病人科室id, strDept, blnCancel, strErr, strLIS, str项目编码)
    
    If strErr <> "" Then
        MsgBox "检验接口内部错误：" & strErr, vbInformation, gstrSysName
        Exit Function
    ElseIf blnCancel Then
        Exit Function
    End If
    
    Call InitCardRsLIS(rsCard)
    rsCard.AddNew Array("临床诊断IDs", "申请信息"), Array(strDiag, strResult)
    ApplyNew检验申请 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AdviceSet检验申请(ByVal lngRow As Long, ByRef rsCard As ADODB.Recordset) As Long
'功能：把检查申请单数据加载到表格中
    Dim i As Long, j As Long
    Dim varTmp As Variant, strTmp As String
    Dim varArr As Variant
    Dim lng申请序号 As Long
    Dim strLIS As String
    Dim str诊断 As String
    Dim str附项 As String
    
    On Error GoTo errH
    mblnRowChange = False
    strLIS = rsCard!申请信息 & ""
    str诊断 = rsCard!临床诊断IDs & ""
    If str诊断 <> "" Then
        str诊断 = GetDiag诊断描述(str诊断)
        If str诊断 <> "" Then
            str诊断 = "申请单诊断<Split2>0<Split2><Split2>" & str诊断
        End If
    End If
    varTmp = Split(strLIS, "<Split B>")
    For i = 0 To UBound(varTmp)
        If i <> 0 Then
            vsAdvice.AddItem "", lngRow + 1
            lngRow = lngRow + 1
        End If
        lng申请序号 = Get申请序号
        varArr = Split(varTmp(i), "<Split A>")
        Call AdviceSet检验组合(lngRow, Val(varArr(7)), varArr(8) & ";" & varArr(3), , , True)
        lngRow = lngRow + 1
        With vsAdvice
            For j = lngRow - 1 To lngRow
                .TextMatrix(j, COL_标志) = Val(varArr(6))
                .TextMatrix(j, COL_开嘱时间) = varArr(2)
                    .Cell(flexcpData, j, COL_开嘱时间) = .TextMatrix(j, COL_开嘱时间)
                .TextMatrix(j, COL_申请序号) = lng申请序号
            Next
            .TextMatrix(lngRow, COL_执行科室ID) = Val(varArr(0))
            .TextMatrix(lngRow, COL_医生嘱托) = varArr(5)
            .TextMatrix(lngRow, COL_诊疗项目ID) = Val(varArr(7))
            .TextMatrix(lngRow - 1, COL_执行科室ID) = Val(varArr(1))
            
            str附项 = varArr(4)
            If str附项 <> "" And str诊断 <> "" Then
                str附项 = str附项 & "<Split1>" & str诊断
            ElseIf str附项 = "" And str诊断 <> "" Then
                str附项 = str诊断
            End If
            If str附项 <> "" Then
                .TextMatrix(lngRow, COL_附项) = str附项
                .Cell(flexcpData, lngRow, COL_附项) = 1
            End If
            
            Set .Cell(flexcpData, lngRow, COL_申请序号) = zlDatabase.CopyNewRec(rsCard)
            Call .AutoSize(col_医嘱内容)
        End With
        Call SetDiagFlag(lngRow, 1)
        Call SetRow标志图标(lngRow, 0)
    Next
    vsAdvice.Row = lngRow
    mblnRowChange = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check单独应用(ByVal lngRow As Long, ByRef strMsg As String) As Boolean
    '检查是否有医嘱的诊疗项目未勾选“可以单独应用”。
    With vsAdvice
        If Not (InStr(",4,5,6,7,G,", .TextMatrix(lngRow, COL_类别)) > 0) Then
            If Not (.TextMatrix(lngRow, COL_类别) = "E" And InStr(",2,3,4,6,7,8,9,", .TextMatrix(lngRow, COL_操作类型)) > 0) Then
                If Val(.TextMatrix(lngRow, COL_单独应用)) = 0 And Val(.TextMatrix(lngRow, COL_诊疗项目ID)) <> 0 Then
                        strMsg = "”对应的诊疗项目不能单独应用！"
                        Check单独应用 = False
                        Exit Function
                End If
            End If
        End If
    End With
    Check单独应用 = True
End Function

Private Sub MakeAppNo(ByVal intType As Integer, ByVal lngBegin As Long, ByVal lngEnd As Long)
'功能：给表格上的医嘱数据产生申请序号
'参数：lngBegin - lngEnd 表格行的范围，intType －1 单组医嘱，2－多组医嘱
    Dim str类别 As String
    Dim lngNo As Long
    Dim lng项目id As Long
    Dim i As Long, j As Long
    Dim lng组ID As Long
    Dim lngPre组ID As Long
    Dim lngStart As Long
    Dim lngStop As Long
    
    On Error GoTo errH
    With vsAdvice
        If intType = 1 Then
            str类别 = .TextMatrix(lngBegin, COL_类别)
            lng项目id = Val(.TextMatrix(lngBegin, COL_诊疗项目ID))
            Select Case str类别
            Case "D", "C"
                If CanUseApply(str类别, lng项目id) Then
                    lngNo = -1 * Get申请序号
                End If
            Case "K"
                If CanUseApply("K", lng项目id) Then
                    lngNo = -1 * Get申请序号
                End If
            End Select
            If lngNo <> 0 Then
                For i = lngBegin To lngEnd
                    .TextMatrix(i, COL_申请序号) = lngNo
                Next
            End If
        Else
            For i = lngBegin To lngEnd
                If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    lng组ID = .RowData(i)
                Else
                    lng组ID = Val(.TextMatrix(i, COL_相关ID))
                End If
                
                lngStop = -1
                lngStart = i
                For j = i + 1 To lngEnd
                    If lng组ID = IIF(Val(.TextMatrix(j, COL_相关ID)) = 0, Val(.RowData(j)), Val(.TextMatrix(j, COL_相关ID))) Then
                        lngStop = j
                    Else
                        Exit For
                    End If
                Next
                If lngStop = -1 Then lngStop = lngStart
                
                Call MakeAppNo(1, lngStart, lngStop)
                
                If lngStop = lngEnd Then
                    Exit Sub
                Else
                    i = lngStop
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Make诊断医嘱对应(ByRef arrSQL As Variant, ByVal str医嘱IDs As String, ByVal lng诊断ID As Long)
'功能：生成诊断医嘱对应的可执行SQL问题号102666
    Dim varTmp As Variant
    Dim i As Long
    Dim strTmp As String
    Dim strIDs As String
    varTmp = Array()
    strIDs = str医嘱IDs
    
    Do While Len(strIDs) > 4000
        strTmp = Mid(strIDs, 1, 3980)
        strTmp = Mid(strTmp, 1, InStrRev(strTmp, ",") - 1)
        ReDim Preserve varTmp(UBound(varTmp) + 1)
        varTmp(UBound(varTmp)) = strTmp
        strIDs = Replace(strIDs, strTmp & ",", "")
    Loop
    If strIDs <> "" Then
        ReDim Preserve varTmp(UBound(varTmp) + 1)
        varTmp(UBound(varTmp)) = strIDs
    End If
    For i = 0 To UBound(varTmp)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(NULL,NULL," & lng诊断ID & ",'" & varTmp(i) & "')"
    Next
End Sub

Private Function GetNext医嘱ID() As Long
'功能：生成临时的医嘱ID
    mlngID序列 = mlngID序列 - 1
    GetNext医嘱ID = mlngID序列
End Function

Private Function GetID指定值(ByRef colIn As Collection, ByVal strKey As String) As Long
'功能：获取指定的真实医嘱ID，用集合方式生成键值对
    Dim strID As String
    
    On Error Resume Next
    
    strID = colIn(strKey)
    If err.Number <> 0 Then
        strID = zlDatabase.GetNextID("病人医嘱记录")
        colIn.Add strID, strKey
    End If
    err.Clear
    
    GetID指定值 = Val(strID)
End Function

Private Sub MakeRealID()
'功能：将医嘱表格中的医嘱ID生成为真实的医嘱ID，包含诊断医嘱对应中的医嘱ID
'说明：提交数据时可能失败，可能会被重复执行，可以用当前值是否大于0进行判断，只有本次新的医嘱才需要重新产生
    Dim i As Long, j As Long
    Dim colID As New Collection
    Dim varTmp As Variant
    Dim str医嘱IDs As String
    Screen.MousePointer = 11
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If InStr("1,2", Val(.TextMatrix(i, COL_EDIT))) > 0 Then   '所有医嘱记录
                If Val(.RowData(i)) < 0 Then
                    .RowData(i) = GetID指定值(colID, CStr(.RowData(i)))
                End If
                If Val(.TextMatrix(i, COL_相关ID)) < 0 Then
                    .TextMatrix(i, COL_相关ID) = GetID指定值(colID, CStr(Val(.TextMatrix(i, COL_相关ID))))
                End If
            End If
        Next
    End With
    
    With vsDiag
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, col诊断)) <> "" Then
                If "" <> .TextMatrix(i, col医嘱ID) Then
                    varTmp = Split(.TextMatrix(i, col医嘱ID), ",")
                    str医嘱IDs = ""
                    For j = 0 To UBound(varTmp)
                        If Val(varTmp(j)) < 0 Then
                            str医嘱IDs = str医嘱IDs & "," & GetID指定值(colID, CStr(Val(varTmp(j))))
                        Else
                            str医嘱IDs = str医嘱IDs & "," & Val(varTmp(j))
                        End If
                    Next
                    .TextMatrix(i, col医嘱ID) = Mid(str医嘱IDs, 2)
                End If
            End If
        Next
    End With
    Screen.MousePointer = 0
End Sub


Private Sub UpdateRecipeNo()
'功能:获取处方序号
'说明:处方序号生成规则
'1-西药和中成药产生一个处方序号，中药饮片应当单独产生处方序号。
'2-西药和中成药每张处方不得超过5种药品（一并给药算一种药品）。
'
    Dim i               As Long
    Dim j               As Long
    Dim rsRecipe        As ADODB.Recordset
    Dim lng医嘱ID       As Long
    Dim lngSumCount     As Long
    Dim lngCount        As Long
    Dim lngNo           As Long
    Dim lngTemp         As Long
    Dim lngRecipeCount  As Long    '处方条数 =5
    
    lngRecipeCount = 5
    '构造缓存医嘱ID与处方序号的记录集缓存
    Set rsRecipe = New ADODB.Recordset
    With rsRecipe
        .Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
        .Fields.Append "RecipeNo", adBigInt
        .Fields.Append "Type", adInteger    '1-西药和中成药;2-中药饮片
        .Fields.Append "Tag", adInteger    '1-已经产生处方序号;2-需要新增处方序号
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_状态)) = 1 And InStr(",5,6,7,", "," & .TextMatrix(i, COL_类别) & ",") > 0 And lng医嘱ID <> Val(.TextMatrix(i, COL_相关ID)) Then
                lng医嘱ID = Val(.TextMatrix(i, COL_相关ID))
                rsRecipe.AddNew
                rsRecipe!医嘱ID = Val(.TextMatrix(i, COL_相关ID))
                rsRecipe!RecipeNo = Val(.TextMatrix(i, COL_处方序号))
                rsRecipe!Tag = IIF(Val(.TextMatrix(i, COL_处方序号)) = 0, 2, 1)
                rsRecipe!Type = IIF(InStr(",5,6,", "," & .TextMatrix(i, COL_类别) & ",") > 0, 1, 2)
            End If
        Next
        If rsRecipe.RecordCount > 0 Then rsRecipe.UpdateBatch
        '西药中成药处方序号
        rsRecipe.Filter = "Type = 1"
        lngSumCount = rsRecipe.RecordCount
        If lngSumCount > 0 Then
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            lngCount = rsRecipe.RecordCount
        End If
        If lngSumCount = lngCount And lngCount > 0 Then
            For i = 1 To rsRecipe.RecordCount
                If i Mod lngRecipeCount = 1 Then
                    lngNo = GetRecipeNo()
                End If
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        ElseIf lngCount > 0 And lngSumCount - lngCount > 0 Then
            rsRecipe.Filter = "Type = 1 And RecipeNo > 0"
            rsRecipe.Sort = "RecipeNo"
            '遍历获取处方序号及需要新添加到该处方序号的医嘱条数
            lngTemp = 0: lngNo = 0
            For i = 1 To rsRecipe.RecordCount
                If lngNo <> rsRecipe!RecipeNo Then
                    If lngTemp > 0 Then Exit For
                    lngNo = rsRecipe!RecipeNo
                    lngTemp = lngRecipeCount
                End If
                If lngNo = rsRecipe!RecipeNo Then lngTemp = lngTemp - 1
                rsRecipe.MoveNext
            Next
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            rsRecipe.Sort = ""
            For i = 1 To rsRecipe.RecordCount
                If i <= lngTemp Then
                    rsRecipe!RecipeNo = lngNo
                Else
                    Exit For
                End If
                rsRecipe.MoveNext
            Next
            '需要新产生处方序号
            rsRecipe.Filter = "Type = 1 And RecipeNo = 0"
            For i = 1 To rsRecipe.RecordCount
                If i Mod lngRecipeCount = 1 Then lngNo = GetRecipeNo()
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        End If
        
        '中药饮片生成处方序号
        rsRecipe.Filter = "Type = 2"
        lngSumCount = rsRecipe.RecordCount
        If lngSumCount > 0 Then
            rsRecipe.Filter = "Type = 2 And RecipeNo = 0"
            lngCount = rsRecipe.RecordCount
        End If
        If lngSumCount = lngCount And lngCount > 0 Then
            lngNo = GetRecipeNo()
            For i = 1 To rsRecipe.RecordCount
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        ElseIf lngCount > 0 And lngSumCount - lngCount > 0 Then
            rsRecipe.Filter = "Type = 2 And RecipeNo > 0"
            If Not rsRecipe.EOF Then lngNo = rsRecipe!RecipeNo
            rsRecipe.Filter = "Type = 2 And RecipeNo = 0"
            For i = 1 To rsRecipe.RecordCount
                rsRecipe!RecipeNo = lngNo
                rsRecipe.MoveNext
            Next
        End If
        rsRecipe.Filter = "Tag=2"
        rsRecipe.Sort = ""
        '将处方序号追加到表格
        For i = 1 To rsRecipe.RecordCount
            lng医嘱ID = rsRecipe!医嘱ID
            lngCount = .FindRow(lng医嘱ID, , COL_相关ID)
            If lngCount <> -1 Then
                For j = lngCount To .Rows - 1
                    If Val(.TextMatrix(j, COL_相关ID)) = lng医嘱ID Or CLng(.RowData(j)) = lng医嘱ID Then
                        .TextMatrix(j, COL_处方序号) = rsRecipe!RecipeNo & ""
                    Else
                        Exit For
                    End If
                Next
            End If
            rsRecipe.MoveNext
        Next
    End With
End Sub

Private Function GetRecipeNo() As Long
'功能:获取处方序号
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select 病人医嘱记录_处方序号.Nextval as 处方序号 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetRecipeNo = Val(rsTmp!处方序号)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReadMsg()
'功能：消息阅读 目前暂时处理ZLHIS_BLOOD_004消息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
        
    strSQL = "select 1 from 病人医嘱记录 a where a.挂号单=[1] and a.医嘱状态=1 and a.诊疗类别='K' and a.检查方法='1' and a.审核状态=1 and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
    
    If rsTmp.EOF Then '没有这类数据了，将消息设为已阅
        strSQL = "select 1 from 业务消息清单 a where a.病人id=[1] and a.就诊id=[2] and a.类型编码='ZLHIS_BLOOD_004' and nvl(a.是否已阅,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
        If Not rsTmp.EOF Then
            strSQL = "Zl_业务消息清单_Read(" & mlng病人ID & "," & mlng挂号ID & ",'ZLHIS_BLOOD_004',1,'" & UserInfo.姓名 & "'," & mlng病人科室id & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
