VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBalance 
   AutoRedraw      =   -1  'True
   Caption         =   "病人结帐处理"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManageBalance.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1035
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":08CA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0AE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0CFE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0F18
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1132
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":18AC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1AC6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1CE0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1EFA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2114
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":232E
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2548
            Key             =   "mzBalance"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2C42
            Key             =   "zyBalance"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":333C
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   270
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":CCD3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":CEED
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D107
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D321
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D53B
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":DCB5
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":DECF
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E0E9
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E303
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E51D
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E737
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E951
            Key             =   "mzBalance"
            Object.Tag             =   "mzBalance"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":F04B
            Key             =   "zyBalance"
            Object.Tag             =   "zyBalance"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":F745
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7410
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1695
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4170
      Width           =   45
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   15
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9675
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4140
      Width           =   9675
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "门诊"
               Key             =   "MzBalance"
               Object.ToolTipText     =   "进入门诊结帐窗口"
               Object.Tag             =   "门诊"
               ImageKey        =   "mzBalance"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "住院"
               Key             =   "ZyBalance"
               Description     =   "结帐"
               Object.ToolTipText     =   "进入住院结帐窗口"
               Object.Tag             =   "住院"
               ImageKey        =   "zyBalance"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "作废"
               Key             =   "Del"
               Description     =   "作废"
               Object.ToolTipText     =   "将当前选中单据作废"
               Object.Tag             =   "作废"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "轧帐"
               Key             =   "轧帐"
               Object.ToolTipText     =   "收费轧帐"
               Object.Tag             =   "轧帐"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5844
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBalance.frx":FE3F
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11986
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   1665
      Left            =   7470
      TabIndex        =   2
      Top             =   4185
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageBalance.frx":106D3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   4185
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageBalance.frx":109ED
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3315
      Left            =   90
      TabIndex        =   0
      Top             =   825
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   5847
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageBalance.frx":10D07
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "现金点钞(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "收费轧帐(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInsure 
         Caption         =   "保险类别(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_MzBalance 
         Caption         =   "门诊病人结帐(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_ZyBalance 
         Caption         =   "住院病人结帐(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_BalanceBat 
         Caption         =   "批量中途结帐(&T)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEdit_BalanceUnit 
         Caption         =   "合约单位结帐(&U)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_Due 
         Caption         =   "应收款管理(&Y)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_RefundDeposit 
         Caption         =   "余额退款(&R)"
      End
      Begin VB.Menu mnuEditSplitMZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzToZy 
         Caption         =   "门诊费用转住院(&Z)"
      End
      Begin VB.Menu mnuEditmzXZ 
         Caption         =   "转住院费用销帐(&X)"
      End
      Begin VB.Menu mnuEditYbVerfy 
         Caption         =   "医保校对(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "单据作废(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "重打结帐票据(&R)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "补打结帐票据(&B)"
      End
      Begin VB.Menu mnuEdit_PrintDetail 
         Caption         =   "打印结帐明细(&L)"
      End
      Begin VB.Menu mnuEditPatiPrint 
         Caption         =   "按病人补打结帐票据(&P)"
      End
      Begin VB.Menu mnuEditSplitWriteCard 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWriteCard 
         Caption         =   "结帐信息写卡(&W)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmManageBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String
Private mlngModul As Long
Private mrsList As ADODB.Recordset  '单据列表
Private mrsDetail As ADODB.Recordset
Private mrsMoney As ADODB.Recordset
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    InPatientID As String
    OutExseID As String
    Patient As String
    Operator As String
    Flag As Byte
    str来源 As String   '0000:门诊;住院;体检;其他
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mblnNOMoved As Boolean '记录当前选择的单据是否是在后备数据表中
Private mbytType As Byte    '记录性质1-结帐记录,2-作废新记录,3-作废原记录
Private mobjInPati As Object
Private mbln立即销帐  As Boolean
Private mstrWriteCardTypeIDs As String   '当前包含的所有卡类别ID
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限
Private mbln医保校对 As Boolean '医保需要校对

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEdit_BalanceBat_Click()
    gblnOK = False
    frmBalanceBat.Show GetModuleType, Me
        
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_BalanceUnit_Click()
    If frmBalanceUnit.ShowMe(Me, 0, 0, False) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Due_Click()
    frmManageDue.mstrPrivs = mstrPrivs
    frmManageDue.mlngModul = mlngModul
    frmManageDue.Show 0, Me
End Sub



Private Sub mnuEdit_PrintDetail_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以打印证明！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 7, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    Call PrintDetail
End Sub

Private Sub mnuEdit_RefundDeposit_Click()
'---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:余额退款
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun          As Object
    
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    
    If objFun.RefundDeposit(glngSys, gcnOracle, Me, gstrDBUser) = False Then
        Set objFun = Nothing
        Exit Sub
    End If
    Set objFun = Nothing
End Sub

Private Sub mnuEditMzToZy_Click()
    '门诊费用转住院费用:33635
    If InStr(1, mstrPrivs, ";门诊费用转住院;") = 0 Then Exit Sub
    mnuEditMzToZy.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0
    If mobjInPati Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
        
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   住院病人部件(zl9InPatient)创建失败,请与系统管理员联系!"
            Exit Sub
        End If
    End If
    'zlOutFeeToInFee(
    '   ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    '   ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, strDBUser As String, _
    '   ByVal lng病人ID As Long, intPatientRange As Integer)
    Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, gstrDBUser, 0, 0)
End Sub

Private Sub mnuEditmzXZ_Click()
    If InStr(mstrPrivs, ";转住院费用销帐;") = 0 Or mbln立即销帐 Then Exit Sub
    If frmFeeRefundment.zlShowEdit(Me, 2, mlngModul, mstrPrivs) = False Then
        Exit Sub
    End If
End Sub

Private Sub mnuEditPatiPrint_Click()
    '按病人补打票据:56283
    If frmMakeupPrintBill.zlRePrintBill(Me, mlngModul, mstrPrivs) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditWriteCard_Click()
    Dim lngCardTypeID As Long, strExpend As String, lng病人ID As Long
    Dim lng结帐ID As Long, strNo As String
    Dim bytFunc As Byte
    
    With mshList
        strNo = .TextMatrix(.Row, GetColNum("单据号"))
        lng结帐ID = Val(.TextMatrix(.Row, GetColNum("结帐ID")))
        lng病人ID = Val(.TextMatrix(.Row, GetColNum("病人ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, GetColNum("标志"))) = 1, 0, 1)
    End With
    '功能:将住院信息写入卡中
    '问题:56615
    If mstrWriteCardTypeIDs = "" Then Exit Sub
    If bytFunc = 0 Then '门诊记帐费用
        If InStr(mstrPrivs, ";门诊信息写卡;") = 0 Then Exit Sub
    Else
        If InStr(mstrPrivs, ";住院信息写卡;") = 0 Then Exit Sub
    End If
    
    If strNo = "" Then
        MsgBox "当前没有单据可以重新写卡！", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, lngCardTypeID, bytFunc, lng结帐ID, lng病人ID)
End Sub
Private Function IsYbBalanceCheck(ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否医保存在校对
    '编制:刘兴洪
    '日期:2015-05-07 16:18:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "" & _
    "   Select 结算方式,金额 From 保险结算明细 " & _
    "   Where 结帐id = [1] And 结算方式<>'现金' " & _
    "           And 标志=1 and Rownum <2"  '医保管控的过程固定写入了一条"现金"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "保险结算管理", lng结帐ID)
    IsYbBalanceCheck = Not rsTmp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub mnuEditYbVerfy_Click()
    Dim strNo As String, lng结帐ID As Long
    Dim int记录状态 As Long, str操作员姓名 As String
    Dim blnYb As Boolean, blnThreeDeposit As Boolean
    Dim lng病人ID As Long, int预交类别 As Integer
    
    With mshList
        lng结帐ID = Val(.TextMatrix(.Row, GetColNum("结帐ID")))
        strNo = .TextMatrix(.Row, GetColNum("单据号"))
        int记录状态 = Val(.TextMatrix(.Row, GetColNum("记录状态")))
        blnYb = .TextMatrix(.Row, GetColNum("医保")) <> ""
        str操作员姓名 = .TextMatrix(.Row, GetColNum("操作员"))
    End With
    If lng结帐ID = 0 Then
        MsgBox "不存在校对的结帐单!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If Not blnYb Then
        MsgBox "当前结帐单不是医保结算单据，不存在医保校对的情况!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If int记录状态 <> 1 Then
        MsgBox "作废的单据，不存在医保校对!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If str操作员姓名 <> UserInfo.姓名 Then
        MsgBox "当前结帐单是操作员“" & str操作员姓名 & "”操作的单据，只能对自己的单据进行医保校对!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If IsYbBalanceCheck(lng结帐ID) = False Then
        MsgBox "当前结帐单不用进行校对操作!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If frmMedicareReckoning.ShowMeFromOut(Me, mstrPrivs, lng结帐ID, blnThreeDeposit, lng病人ID, int预交类别) = False Then Exit Sub
    
    If blnThreeDeposit Then
        frmBalanceDeposit.ShowMe Me, mlngModul, lng结帐ID, lng病人ID, True, False, "", "", int预交类别
    End If
    
    If mnuViewRefeshOptionItem(1).Checked Then
      If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
          mnuViewReFlash_Click
       End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
       mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 1
    frmSetExpence.Show 1, Me
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub
 

Private Sub mnuFileRollingCurtain_Click()
   Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "病人ID=" & .TextMatrix(.Row, GetColNum("病人ID")), _
                "住院号=" & .TextMatrix(.Row, GetColNum("住院号")), _
                "结帐ID=" & .TextMatrix(.Row, GetColNum("结帐ID")), _
                "NO=" & strNo, _
                "记录状态=" & mbytType)
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmBalanceFilter.Show 1, Me
    If gblnOK Then
        With frmBalanceFilter
            mstrFilter = .mstrFilter
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.FactB = .txtFactBegin.Text
            SQLCondition.FactE = .txtFactEnd.Text
            SQLCondition.InPatientID = Trim(.txt住院号.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt姓名.Text) & "%"
            SQLCondition.Operator = NeedName(.cbo操作员.Text)
            SQLCondition.OutExseID = Trim(.txtClinic.Text)
            SQLCondition.str来源 = .mstr来源
            
            If .chkType(0).Value = 1 And .chkType(1).Value = 1 Then
                SQLCondition.Flag = 0
            ElseIf .chkType(0).Value = 1 Then
                SQLCondition.Flag = 1
            ElseIf .chkType(1).Value = 1 Then
                SQLCondition.Flag = 2
            End If
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim lng结帐ID As Long
    Dim strNo As String, int来源 As Integer
    Dim blnYb As Boolean, blnMzBalance As Boolean, blnZyBalance As Boolean
    Dim bln普通病人 As Boolean, bln医保病人 As Boolean
    Dim bln校对 As Boolean, bytFunc As Byte
    
    lng结帐ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    mbytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("记录状态")))
    blnYb = mshList.TextMatrix(mshList.Row, GetColNum("医保")) <> ""
    bln校对 = Trim(mshList.TextMatrix(mshList.Row, GetColNum("校对"))) <> ""
    bytFunc = IIf(Val(mshList.TextMatrix(mshList.Row, GetColNum("标志"))) = 1, 0, 1)
    
    bln普通病人 = InStr(mstrPrivs, ";普通病人结算;") > 0
    bln医保病人 = InStr(mstrPrivs, ";保险结算;") > 0
    blnMzBalance = InStr(1, mstrPrivs, ";门诊费用结帐;") > 0 And bln医保病人
    blnZyBalance = InStr(1, mstrPrivs, ";住院费用结帐;") > 0 And bln医保病人
    
    mnuEditYbVerfy.Visible = False
    If blnYb And lng结帐ID <> 0 And mbytType = 1 Then
        blnYb = bln校对
        mnuEditYbVerfy.Visible = blnYb And (blnMzBalance Or blnZyBalance)
    End If
    
    If mshList.Row = 0 Or lng结帐ID = 0 Then
        mnuEdit_PrintDetail.Enabled = False
        mnuEdit_Print_Supplemental.Enabled = False
        mnuEdit_Print.Enabled = False
        mnuEdit_Del.Enabled = False
        tbr.Buttons("Del").Enabled = False
        mnuEditWriteCard.Enabled = False
        Exit Sub
    End If
    
    mnuEditWriteCard.Enabled = (bytFunc = 0 And InStr(mstrPrivs, ";门诊信息写卡;") > 0) _
                            Or (bytFunc = 1 And InStr(mstrPrivs, ";住院信息写卡;") > 0)
    int来源 = Val(mshList.TextMatrix(mshList.Row, GetColNum("标志")))
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(lng结帐ID, , strNo, int来源)
    Call ShowMoney(lng结帐ID, , strNo)
            
    mnuEdit_PrintDetail.Enabled = mbytType = 1
    mnuEdit_Print_Supplemental.Enabled = mbytType = 1 And Trim(mshList.TextMatrix(mshList.Row, GetColNum("票据号"))) = ""
    mnuEdit_Print.Enabled = mbytType = 1
    mnuEdit_Del.Enabled = mbytType = 1 And Not bln校对
    tbr.Buttons("Del").Enabled = mbytType = 1 And Not bln校对
    
    mshList.ForeColorSel = mshList.CellForeColor
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
        mshMoney.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
        mshMoney.BackColorSel = &HE0E0E0
    ElseIf obj Is mshMoney Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshMoney.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    
    If strNo = "" Then
        MsgBox "当前没有单据可以作废！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mstrInNO = strNo
    frmBalance.Show GetModuleType, Me
        
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_zyBalance_Click()
    On Error Resume Next
    Err.Clear
    
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mbytFunc = 1
    frmBalance.Show GetModuleType, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub
Private Sub mnuEdit_MzBalance_Click()
    On Error Resume Next
    Err.Clear
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mbytFunc = 0
    frmBalance.Show GetModuleType, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前内容已变化,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, lngPatientID As Long, lng结帐ID As Long
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    lngPatientID = Val(mshList.TextMatrix(mshList.Row, GetColNum("病人ID")))
    lng结帐ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")))
    If lngPatientID = 0 Then
      '显示单据内容
        Call frmBalanceUnit.ShowMe(Me, 1, lng结帐ID, IIf(mbytType = 2, True, False), mblnNOMoved)
    Else
        '显示单据内容
        frmBalance.mlngModul = mlngModul
        frmBalance.mstrPrivs = mstrPrivs
        frmBalance.mbytInState = 1
        frmBalance.mblnViewCancel = IIf(mbytType = 2, True, False)
        frmBalance.mstrInNO = strNo
        frmBalance.mblnNOMoved = mblnNOMoved
        frmBalance.mlngBillID = Val(mshList.TextMatrix(mshList.Row, 0))
        frmBalance.Show GetModuleType, Me
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshMoney_EnterCell()
    mshMoney.ForeColorSel = mshMoney.CellForeColor
End Sub

Private Sub mshMoney_GotFocus()
    Call SetActiveList(mshMoney)
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        picVsc.Top = picVsc.Top + Y
        picVsc.Height = picVsc.Height - Y
        mshMoney.Top = mshMoney.Top + Y
        mshMoney.Height = mshMoney.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDetail.Width + X < 1000 Or mshMoney.Width - X < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        mshDetail.Width = mshDetail.Width + X
        mshMoney.Left = mshMoney.Left + X
        mshMoney.Width = mshMoney.Width - X
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "ZyBalance"
            mnuEdit_zyBalance_Click
        Case "MzBalance"
            mnuEdit_MzBalance_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "轧帐"
            mnuFileRollingCurtain_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '表头
    objOut.Title.Text = "病人结帐单据清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmBalanceFilter
        objRow.Add "时间：" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " 至 " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objRow.Add "性质：" & IIf(mbytType = 2, "作废单据", "结帐单据")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    mnuEdit_Print.Enabled = blnUsed
    mnuEdit_PrintDetail.Enabled = blnUsed
    mnuEdit_Print_Supplemental.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed

    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub
Private Sub 权限控制()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:权限控制
    '编制:刘兴洪
    '日期:2011-09-20 23:27:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean, blnZyBalance As Boolean '门诊结帐,住院结帐
    Dim bln普通病人 As Boolean, bln医保病人 As Boolean, bln作废 As Boolean
    Dim blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    mnuEditMzToZy.Visible = InStr(1, mstrPrivs, ";门诊费用转住院;") > 0 '33635
    mnuEditmzXZ.Visible = InStr(mstrPrivs, ";转住院费用销帐;") > 0 And Not mbln立即销帐
    mnuEditSplitMZ.Visible = InStr(mstrPrivs, ";转住院费用销帐;") > 0 And Not mbln立即销帐 Or InStr(1, mstrPrivs, ";门诊费用转住院;") > 0
    
    bln普通病人 = InStr(mstrPrivs, ";普通病人结算;") > 0
    bln医保病人 = InStr(mstrPrivs, ";保险结算;") > 0
    blnMzBalance = InStr(1, mstrPrivs, ";门诊费用结帐;") > 0 And (bln普通病人 Or bln医保病人)
    mnuEdit_MzBalance.Visible = blnMzBalance
    mnuEdit_BalanceUnit.Visible = blnMzBalance '合约单位结帐
    tbr.Buttons("MzBalance").Visible = blnMzBalance
        
    blnZyBalance = InStr(1, mstrPrivs, ";住院费用结帐;") > 0 And (bln普通病人 Or bln医保病人)
    mnuEdit_ZyBalance.Visible = blnZyBalance
    tbr.Buttons("ZyBalance").Visible = blnZyBalance
    
    mnuEdit_Print.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";重打票据;") > 0 '重打票据
    '52328
    mnuEdit_Print_Supplemental.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";补打票据;") > 0        '补打票据
    '问题:56283
    mnuEditPatiPrint.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";补打票据;") > 0        '补打票据
    
    mnuEdit_Due.Visible = InStr(mstrPrivs, ";应收款管理;") > 0
    
    mnuEdit_BalanceBat.Visible = InStr(mstrPrivs, ";批量中途结帐;") > 0
    
    '结帐分隔
    mnuEdit_0.Visible = blnMzBalance Or blnZyBalance Or InStr(mstrPrivs, ";应收款管理;") > 0 Or InStr(mstrPrivs, ";批量中途结帐;") > 0
    
    bln作废 = InStr(mstrPrivs, ";结帐作废;") > 0 And (bln普通病人 Or bln医保病人)
    mnuEdit_Del.Visible = bln作废
    tbr.Buttons("Del").Visible = bln作废
    
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("轧帐").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    mnuEditSplitWriteCard.Visible = (InStr(mstrPrivs, ";住院信息写卡;") > 0 Or InStr(mstrPrivs, ";门诊信息写卡;") > 0) _
                                    And mstrWriteCardTypeIDs <> ""
    mnuEditWriteCard.Visible = (InStr(mstrPrivs, ";住院信息写卡;") > 0 Or InStr(mstrPrivs, ";门诊信息写卡;") > 0) _
                                And mstrWriteCardTypeIDs <> ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    
    Call RestoreWinState(Me, App.ProductName)
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    mstrWriteCardTypeIDs = ""
    If Not gobjSquare Is Nothing Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            mstrWriteCardTypeIDs = gobjSquare.objSquareCard.zlGetAvailabilityWriteCardType
        End If
    End If
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    '创建并检测税控打印对象
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(2)
        End If
        On Error GoTo 0
    End If
    
    '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.编号, UserInfo.姓名)
    End If
    On Error GoTo 0
    
    
    '权限设置
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1137_1")
    Call 权限控制
    
    Call ClearErrInvoice
    

    '刷新时缺省过滤条件(当天内),进入时不显示任何记录
    mstrFilter = " And A.收费时间 Between Trunc(sysdate) And Trunc(sysdate+1) " & _
                 " And A.操作员姓名||''=[1] And A.记录状态=1"
    SQLCondition.Default = True
    SQLCondition.Operator = UserInfo.姓名
            
    Call SetHeader
    Call SetDetail
    Call SetMoney
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "请刷新清单或设置过滤条件"
End Sub


Private Sub ClearErrInvoice()
'功能：清除操作员上次异常退出时只填了实际票号而没有实际打印的单据的结帐记录中的票据号
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select A.NO" & vbNewLine & _
            "From 病人结帐记录 A," & vbNewLine & _
            "     (Select Max(NO) NO From 病人结帐记录 Where 收费时间 > Sysdate - 1 And 操作员姓名 = [1]) B" & vbNewLine & _
            "Where A.NO = B.NO And A.实际票号 Is Not Null And Not Exists (Select 1 From 票据打印内容 C Where C.NO = B.NO)"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        strSql = "Zl_票据起始号_Update('" & rsTmp!NO & "','',3)"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim sngVsc As Single, sngHsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    sngHsc = mshMoney.Width / (mshMoney.Width + mshDetail.Width)
    
    If mblnMax Then
        sngVsc = 0.3: sngHsc = 0.2
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    mshList.Left = Me.ScaleLeft
    mshList.Top = Me.ScaleTop + cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = 0
    picHsc.Width = mshList.Width
    
    mshDetail.Left = 0
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - picHsc.Height - mshList.Height
    mshDetail.Width = (Me.ScaleWidth - picVsc.Width) * (1 - sngHsc)
    
    picVsc.Top = mshDetail.Top
    picVsc.Left = mshDetail.Left + mshDetail.Width
    picVsc.Height = mshDetail.Height
    
    mshMoney.Top = mshDetail.Top
    mshMoney.Left = picVsc.Left + picVsc.Width
    mshMoney.Height = mshDetail.Height
    mshMoney.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
        
    mbytType = 0
    mstrFilter = ""
    Unload frmBalanceFilter
    Unload frmBalanceGo
    Call SaveWinState(Me, App.ProductName)
    '33635
    Set mobjInPati = Nothing
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmBalanceGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBalanceGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, int来源 As Integer
    Dim blnFill As Boolean
    Dim strNo As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmBalanceGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("单据号")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("票据号")) = .txtFact.Text
            End If
            If .txt住院号.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("住院号")) = .txt住院号.Text
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mlngGo = i + 1
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
            int来源 = Val(mshList.TextMatrix(mshList.Row, GetColNum("标志")))
            Call ShowDetail(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")), , strNo, int来源)
            Call ShowMoney(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")), , strNo)
            
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(bytMode As Byte)
'功能：当前收款记录重新打印一张票据
'bytMode=0-重打,1-补打
    Dim strNo As String, lng结帐ID As Long, blnMediCare As Boolean, bytFlag As Byte '门诊还是住院
    Dim intInsure As Integer
    Dim lng病人ID As Long, bytFunc As Byte
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
        Exit Sub
    End If
    lng结帐ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")))
    lng病人ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("病人ID")))
    bytFunc = IIf(Val(mshList.TextMatrix(mshList.Row, GetColNum("标志"))) = 1, 0, 1)
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 7, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
     '单据权限
    If bytMode = 0 Then
        If Not BillOperCheck(7, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("收费时间"))), "重打") Then Exit Sub
    Else
        If Trim(mshList.TextMatrix(mshList.Row, GetColNum("票据号"))) <> "" Then
            MsgBox "当前单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    intInsure = BalanceExistInsure(strNo, bytFlag)
    If RePrintBalance(strNo, Me, lng结帐ID, intInsure) Then
    
        '银医一卡通写卡，85950
        Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng结帐ID, lng病人ID)

        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("单据号")) = "" Then Exit Sub
        If mshList.ColWidth(lngCol) = 0 Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    '问题65105:增加门诊号列信息
    strHead = "结帐ID,1,0|标志,1,0|医保,4,500|单据号,4,850|票据号,4,850|病人ID,1,750|门诊号,1,750|住院号,1,750|姓名,4,800|性别,4,500|年龄,4,500|费别,4,750|开始日期,4,1000|结束日期,4,1000|结帐金额,7,850|操作员,4,800|收费时间,4,1850|校对,4,500|中途结帐,4,800|记录状态,1,0"
    
    With mshList
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        
        Call mshList_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional strIF As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strIF=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, k As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim strSql As String, str来源 As String, strWhere As String
    
    
    On Error GoTo errH
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        DoEvents
        Me.Refresh
        If SQLCondition.str来源 = "" Then SQLCondition.str来源 = "1111" '0000:门诊;住院;体检;其他
            
        str来源 = ""
        For i = 1 To Len(SQLCondition.str来源)
            If Mid(SQLCondition.str来源, i, 1) = 1 Then
                str来源 = str来源 & "," & Choose(i, 1, 2, 4, 3)  '1-门诊;2-住院;3-其他(就诊卡等额外的收费);4-体检
            End If
        Next
        If str来源 <> "" Then str来源 = Mid(str来源, 2)
        If str来源 = "" Then str来源 = "-1"
        
        
        strTable = "" & _
        "   Select A.ID ,1 as 住院标志,0 as 门诊标志,A.NO,A.实际票号,A.病人ID,B.病人ID as 费用病人ID,A.开始日期,A.结束日期,A.记录状态,B.结帐金额,A.操作员姓名,A.收费时间,A.中途结帐,A.原因 as 合约单位,A.结帐类型,B.主页ID " & _
        "   From 病人结帐记录 A,住院费用记录 B,病人信息 C " & _
        "   Where A.ID=B.结帐ID and  B.病人ID=C.病人ID" & _
                IIf(SQLCondition.str来源 = "1111", "", " And Instr(',' || [11] || ',', ',' || Nvl(B.门诊标志,0) || ',') > 0 ") & strIF
        
        
        Select Case SQLCondition.str来源
        Case "1010", "1000", "0010"  '门诊
            strTable = Replace(strTable, "住院费用记录", "门诊费用记录")
            strTable = Replace(strTable, "B.主页ID", "Null As 主页ID")
            strTable = Replace(strTable, "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志")
        Case "0101", "0001", "0100" '住院
            '已经存在
        Case Else '门诊和住院
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "住院费用记录", "门诊费用记录"), "1 as 住院标志,0 as 门诊标志", "0 as 住院标志,1 as 门诊标志"), "B.主页ID", "Null As 主页ID")
        End Select
        
        If frmBalanceFilter.mblnDateMoved Then
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "病人结帐记录", "H病人结帐记录"), "住院费用记录", "H住院费用记录"), "门诊费用记录", "H门诊费用记录")
        End If
        
        '问题65105,刘尔旋:门诊结帐时才显示门诊号
        strSql = _
        " Select A.ID 结帐ID,decode(Max(住院标志),1,decode(max(门诊标志),1,3,2),1) as 标志 ,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√') as 医保,A.NO as 单据号,A.实际票号 as 票据号," & _
        "        Decode(A.病人ID,Null,' ',A.病人ID) 病人ID,Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)) 门诊号,Decode(A.病人ID,Null,' ',Nvl(P.住院号,C.住院号)) 住院号," & _
        "        Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名) 姓名,Decode(A.病人ID,Null,' ',C.性别) 性别," & _
        "        Decode(A.病人ID,Null,' ',C.年龄) 年龄,Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别)) as 费别," & _
        "        To_Char(A.开始日期,'YYYY-MM-DD') as 开始日期,To_Char(A.结束日期,'YYYY-MM-DD') as 结束日期," & _
        "        To_Char(Sum(Decode(A.记录状态,2,-1,1) *A.结帐金额),'999999999" & gstrDec & "') as 结帐金额," & _
        "        A.操作员姓名 as 操作员,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间," & _
        "        ' '  as 校对," & _
        "        Decode(Nvl(A.中途结帐,0),1,'√',' ') 中途结帐,Max(A.记录状态) as 记录状态" & _
        " From ( " & strTable & ") A,病人信息 C,病案主页 P,合约单位 Q,人员表 N" & _
        " Where  A.费用病人ID=C.病人ID And A.操作员姓名=N.姓名  " & _
        "        And A.费用病人ID=P.病人ID(+) And Nvl(A.主页ID,0)=P.主页ID(+) And C.合同单位ID=Q.ID(+)" & _
        "       And (N.站点='" & gstrNodeNo & "' Or N.站点 is Null)" & vbNewLine & _
        " Group by A.ID,Decode(P.险类,NULL,Decode(C.险类,NULL,NULL,'√'),'√'),A.NO,A.实际票号,Decode(A.病人ID,Null,' ',A.病人ID),Decode(Nvl(A.结帐类型,0),2,' ',Decode(A.病人ID,Null,' ',C.门诊号)),Decode(A.病人ID,Null,' ',Nvl(P.住院号,C.住院号))," & _
        "           Decode(A.病人ID,Null,nvl(A.合约单位,Q.名称),C.姓名),Decode(A.病人ID,Null,' ',C.性别),Decode(A.病人ID,Null,' ',C.年龄),Decode(A.病人ID,Null,' ',Nvl(P.费别,C.费别))," & _
        "           To_Char(A.开始日期,'YYYY-MM-DD'),To_Char(A.结束日期,'YYYY-MM-DD')," & _
        "           A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),Decode(Nvl(A.中途结帐,0),1,'√',' ')"
        
        strSql = strSql & " Order by 收费时间 Desc,单据号 Desc"
        
        With SQLCondition
            If .Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .Operator, "", "", "", "", "", 0, "", "", "", str来源)
            Else
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .FactB, .FactE, Val(.InPatientID), .Patient, .Operator, .OutExseID, str来源)
            End If
        End With
    End If
    
    mshList.Redraw = False
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshMoney.Clear
    mshMoney.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        stbThis.Panels(2) = "共 " & mrsList.RecordCount & " 张单据"
        Call SetMenu(True)
    End If
    
    '设置颜色
    If SQLCondition.Flag = 2 Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("记录状态")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '退费记录用红色
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '包含退过费的用蓝色
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
            strSql = "Select 1 From 保险结算明细 Where 结帐ID=[1] And 标志=1 And 结算方式 <> '现金'"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mshList.TextMatrix(i, GetColNum("结帐ID"))))
            If Not rsTemp.EOF Then
                mshList.TextMatrix(i, GetColNum("校对")) = "√"
            Else
                mshList.TextMatrix(i, GetColNum("校对")) = ""
            End If
        Next
    End If
    
    Call SetHeader '此过程已包括Call SetDetail,Call SetMoney
    
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")) = "" Then
        Call SetDetail
        Call SetMoney
    End If
    
    If Not blnSort Then Call zlCommFun.StopFlash
    mshList.Redraw = True
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    mshList.Redraw = True
End Sub

Private Sub ShowDetail(Optional lng结帐ID As Long, Optional blnSort As Boolean, Optional strNo As String, Optional int来源 As Integer = 2)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：显示明细费用记录
    '入参：int来源-1-门诊;2-住院;3-门诊和住院
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-08 19:34:16
    '说明：
    '------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, strSql As String, strDec As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If Not blnSort Then
        
        If frmBalanceFilter.mblnDateMoved Then
            '结帐单通过结帐ID与费用记录相连,如果指定NO号的结帐单在后备表中,则该次结帐的单据号一定在后备表中
            '一张结帐单关联的记帐单不可能同时在在线表和后备表中
            mblnNOMoved = zlDatabase.NOMoved("病人结帐记录", strNo, , , Me.Caption)
        Else
            mblnNOMoved = False   '必须要有这一句
        End If
        
        strDec = gstrDec
        If lng结帐ID <> 0 Then
            Select Case int来源
            Case 1 '门诊
                strSql = "Select Max(Length(Abs(结帐金额) - Trunc(Abs(结帐金额))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1]"
            Case 2 '住院
                strSql = "Select Max(Length(Abs(结帐金额) - Trunc(Abs(结帐金额))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
            Case Else
                
                strSql = "Select Length(Abs(结帐金额) - Trunc(Abs(结帐金额)))  as  declen From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 Where 结帐ID=[1] Union ALL " & _
                         "Select Length(Abs(结帐金额) - Trunc(Abs(结帐金额)))   as  declen  From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 Where 结帐ID=[1]"
                strSql = "Select Max(declen)-1 as declen  From ( " & strSql & ")"
            End Select
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
            If rsTmp.RecordCount > 0 Then
                If Len(strDec) < Len("0." & String(rsTmp!declen, "0")) Then
                    strDec = "0." & String(rsTmp!declen, "0")
                End If
            End If
        End If
        
        Select Case int来源
        Case 1  '门诊
            strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] ) A "
            'strSQL = IIf(mblnNOMoved, "H", "") & "门诊费用记录 A "
        Case 2  '住院
            strSql = IIf(mblnNOMoved, "H", "") & "住院费用记录 A"
        Case Else '门诊和住院
            strSql = " (Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,0 as 主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A where A.结帐ID=[1] Union ALL " & _
                       " Select 结帐ID,NO,序号,开单部门ID,收费细目ID,门诊标志,主页ID,收据费目,婴儿费,结帐金额,发生时间 From " & IIf(mblnNOMoved, "H", "") & "住院费用记录 A where A.结帐ID=[1] )  A"
        End Select
        
        strSql = _
        "   Select Decode(门诊标志,1,'门诊',4,'门诊','第'||Nvl(A.主页ID,0)||'次') as 住院," & _
        "         A.NO as 单据号,Nvl(B.名称,'未知') as 开单科室,Nvl(E.名称,D.名称) as 项目," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & _
        "       A.收据费目 as 费目,Decode(Nvl(A.婴儿费,0),0,'','√') as 婴儿费," & _
        "       To_Char(" & IIf(mbytType = 2, "-1*", "") & "A.结帐金额,'999999999" & strDec & "') as 结帐金额," & _
        "       To_Char(A.发生时间,'YYYY-MM-DD HH24:MI:SS') as 费用时间" & _
        " From " & strSql & ",部门表 B,收费项目目录 D,收费项目别名 E" & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
        " Where A.开单部门ID=B.ID(+) And A.收费细目ID=D.ID" & _
        "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
                IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
        "       And A.结帐ID=[1]" & _
        " Order by 住院 Desc,费用时间 Desc,单据号 Desc,A.序号"
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    End If
    
    mshDetail.Redraw = False
    mshDetail.Clear
    mshDetail.Rows = 2
    mshDetail.ForeColor = IIf(mbytType = 2, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '设置颜色
    If mbytType = 2 Then
        '退费直接为红色
        mshDetail.ForeColor = &HC0
    Else
        '原始单据退过的为蓝色
        mshDetail.ForeColor = ForeColor
        If mbytType = 3 Then
            For i = 1 To mshDetail.Rows - 1
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            Next
        End If
    End If
    
    Call SetDetail
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mshDetail.Redraw = True
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "住院,4,750|单据号,4,850|开单科室,1,850|项目,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,1600", "") & "|费目,1,850|婴儿费,4,650|结帐金额,7,850|费用时间,1,1850"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '刘兴洪:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
                
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowMoney(Optional lng结帐ID As Long, Optional blnSort As Boolean, _
    Optional strNo As String)
    Dim i As Long, strSql As String
    On Error GoTo errH
    
    '如果当前结帐单在后备表中,则它相关的结帐方式一定在后备表中
    If Not blnSort Then
        strSql = "" & _
        " Select Decode(Substr(记录性质,Length(记录性质),1),1,'冲预交',2,'补款') as 类型," & _
        "       NO as 单据号,To_Char(" & IIf(mbytType = 2, "-1*", "") & "冲预交,'FM9999999990.00999') as 金额," & _
        "       结算方式,结算号码" & _
        " From 病人预交记录 " & _
        " Where 结帐ID=[1] And 冲预交 <> 0 " & _
        " Order by 类型 Desc,NO Desc,结算方式"
        If mblnNOMoved Then strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
        Set mrsMoney = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    End If
    
    mshMoney.Clear
    mshMoney.Rows = 2
    mshMoney.ForeColor = mshList.ForeColor
    If Not mrsMoney.EOF Then Set mshMoney.DataSource = mrsMoney
    Call SetMoney
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMoney()
    Dim strHead As String
    Dim i As Long
    
    strHead = "类型,4,650|单据号,4,850|金额,7,850|结算方式,1,850|结算号码,1,1000"
    With mshMoney
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshMoney, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        Call mshMoney_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, True)
    End If
End Sub

Private Sub mshMoney_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshMoney.MouseRow = 0 Then
        mshMoney.MousePointer = 99
    Else
        mshMoney.MousePointer = 0
    End If
End Sub

Private Sub mshMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshMoney.MouseCol
    
    If Button = 1 And mshMoney.MousePointer = 99 Then
        If mshMoney.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshMoney.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsMoney Is Nothing Then Exit Sub
        
        Set mshMoney.DataSource = Nothing

        mrsMoney.Sort = mshMoney.TextMatrix(0, lngCol) & IIf(mshMoney.ColData(lngCol) = 0, "", " DESC")
        mshMoney.ColData(lngCol) = (mshMoney.ColData(lngCol) + 1) Mod 2
        
        Call ShowMoney(, True)
    End If
End Sub

Private Sub PrintDetail()
'功能：输入出列表
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshDetail.Row
    
    '表头
    objOut.Title.Text = "病人结帐单据明细"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    With frmBalanceFilter
        objRow.Add "单据号：" & mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        objRow.Add "结帐范围：" & mshList.TextMatrix(mshList.Row, GetColNum("开始日期")) & " 至 " & mshList.TextMatrix(mshList.Row, GetColNum("结束日期"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "住院号：" & mshList.TextMatrix(mshList.Row, GetColNum("住院号"))
        objRow.Add "姓名：" & mshList.TextMatrix(mshList.Row, GetColNum("姓名"))
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshDetail.Redraw = False
    Set objOut.Body = mshDetail
    
    bytR = zlPrintAsk(objOut)
    Me.Refresh
    If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    
    mshDetail.Row = intRow
    mshDetail.Col = 0: mshDetail.ColSel = mshDetail.Cols - 1
    mshDetail.Redraw = True
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

