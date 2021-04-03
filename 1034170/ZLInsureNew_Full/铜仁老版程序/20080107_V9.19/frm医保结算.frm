VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保结算 
   Caption         =   "医保结算管理"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   Icon            =   "frm医保结算.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2580
      TabIndex        =   10
      Top             =   1785
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   6150
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2850
      ScaleWidth      =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2190
      Width           =   90
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5445
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保结算.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
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
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   -60
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   9195
      TabIndex        =   6
      Top             =   3960
      Width           =   9195
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5220
      Top             =   420
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
            Picture         =   "frm医保结算.frx":115C
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":1376
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":1590
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":17AA
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":19C4
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":1BDE
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":22D8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":29D2
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":30CC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":32E6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":3500
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":371A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":3934
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":3B4E
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5865
      Top             =   450
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
            Picture         =   "frm医保结算.frx":3D68
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":3F82
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":419C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":43B6
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":45D0
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":47EA
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":4EE4
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":55DE
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":5CD8
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":5EF2
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":610C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":6326
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":6540
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保结算.frx":675A
            Key             =   "Cancel"
            Object.Tag             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   1244
      BandCount       =   2
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   9525
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   810
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "保险类别"
      Child2          =   "cmb险类"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1935
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
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
               Caption         =   "编辑"
               Key             =   "编辑"
               Object.ToolTipText     =   "编辑最高限额"
               Object.Tag             =   "编辑"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存最高限额"
               Object.Tag             =   "保存"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "放弃"
               Key             =   "放弃"
               Object.ToolTipText     =   "放弃所编辑的最高限额"
               Object.Tag             =   "放弃"
               ImageKey        =   "Cancel"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_1"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Find"
               Description     =   "过滤"
               Object.ToolTipText     =   "查找医保帐户"
               Object.Tag             =   "过滤"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmb险类 
         Height          =   300
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   195
         Width           =   1995
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh记录_S 
      Height          =   2805
      Left            =   60
      TabIndex        =   3
      Top             =   1110
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   4948
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
      MouseIcon       =   "frm医保结算.frx":6974
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh明细 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   4050
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2355
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
      MouseIcon       =   "frm医保结算.frx":6C8E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh分档 
      Height          =   1335
      Left            =   4710
      TabIndex        =   8
      Top             =   4020
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   2355
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
      MouseIcon       =   "frm医保结算.frx":6FA8
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip tab性质 
      Height          =   345
      Left            =   30
      TabIndex        =   7
      Top             =   750
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   609
      TabWidthStyle   =   2
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "收费"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "结算"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "预交"
            Key             =   "K3"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFileSplitSet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSplitPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintBalance 
         Caption         =   "打印结算单(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplitExcel 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBalance 
         Caption         =   "重提发票信息(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBill 
         Caption         =   "打印报销单据(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileDetail 
         Caption         =   "费用明细(&D)"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "明细批量打印(&B)"
      End
      Begin VB.Menu mnuFileSplitReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditClinic 
         Caption         =   "门诊报销(&C)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditIn_Patient 
         Caption         =   "住院报销(&P)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "作废(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnusplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditXE 
         Caption         =   "编辑限额(&X)"
      End
      Begin VB.Menu mnuEditSave 
         Caption         =   "保存限额(&S)"
      End
      Begin VB.Menu mnuEditCacel 
         Caption         =   "放弃限额(&F)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "查阅(&V)"
      End
   End
   Begin VB.Menu mnuBalance 
      Caption         =   "结算(&B)"
      Visible         =   0   'False
      Begin VB.Menu mnuBalanceBill 
         Caption         =   "提取结算单(&B)"
      End
      Begin VB.Menu mnuBalanceCollect 
         Caption         =   "提取结算表(&C)"
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
Attribute VB_Name = "frm医保结算"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private Enum 记录Enum
    col记录ID = 0
    col单据号 = 1
    col中心 = 2
    col卡号 = 3
    col病人ID = 4
    col住院号 = 5
    col姓名 = 6
    col身份 = 7
    col性别 = 8
    col年龄 = 9
    col科室 = 10
    col操作员姓名 = 11
    col登记时间 = 12
    col收退标志 = 13
    col个人帐户 = 14
    col发生费用 = 15
    col实际起付线 = 16
    col进入统筹 = 17
    col统筹报销 = 18
End Enum

Private Enum 明细Enum
    det收费类别 = 0
    det收费细目 = 1
    det规格 = 2
    det单位 = 3
    det数次 = 4
    det单价 = 5
    det实收金额 = 6
    det统筹金额 = 7
    det医保大类 = 8
    det费用类型 = 9
    det收退 = 10
    det状态 = 11
End Enum

Private mblnLoad As Boolean                     '第一次启动

Private mint险类 As Integer
Private mint性质 As Integer
Private mdatBegin As Date, mdatEnd As Date
Private mstrCardCond As String

Dim msngStartX As Single, msngStartY As Single    '移动前鼠标的位置
Private mrs结算记录 As New ADODB.Recordset
Private mcol中心 As New Collection              '保存医保类别具有中心特性
Private mblnChange As Boolean           '编辑改变
Private mblnEdit As Boolean             '当前是否处于编辑状态
Private mblnNOScroll As Boolean         '不滚动
Private Const mintCol最高限额 = 14      '大连医保用,编辑最高限额的例

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmb险类_Click()
    Dim blnYes As Boolean
    Dim i As Long
    
    With cmb险类
        If mint险类 = .ItemData(.ListIndex) Then Exit Sub
        If mint险类 = type_大连市 Or mint险类 = type_大连开发区 Then
            If mblnEdit And mblnChange = True Then
                ShowMsgbox "当前正在于编辑状态且已经被修改，是否放弃本次修改？", True, blnYes
                If Not blnYes Then
                    For i = 0 To .ListCount - 1
                        If mint险类 = .ItemData(i) Then
                            .ListIndex = i
                            Exit For
                        End If
                    Next
                    Exit Sub
                End If
            End If
        End If
        mint险类 = .ItemData(.ListIndex)
        mnuFileBalance.Visible = False
        If mint险类 = TYPE_沈阳市 Then
            mnuFileBalance.Visible = True
            mnuBalance.Visible = True
        End If
        mnuPrintBalance.Visible = (mint险类 = TYPE_重庆银海版)
    End With
    Call 权限控制
    Call FillList
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        mdatBegin = CDate(Format(zlDataBase.Currentdate, "yyyy-MM-dd"))
        mdatEnd = CDate(Format(mdatBegin, "yyyy-MM-dd") & " 23:59:59")
        mstrCardCond = ""
        
        
        '强制显示
        msh明细.Visible = False
        '显示记录
        Call tab性质_Click
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mint险类 = -1
    mint性质 = -1
    
    mstrPrivs = gstrPrivs
    zlControl.CboSetHeight cmb险类, 3600
    Call InitTable
    
    RestoreWinState Me, App.ProductName
    Call 权限控制
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbr.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbr.Visible, cbr.Top + cbrHeight, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    If Me.WindowState = 1 Then Exit Sub
    
    tab性质.Left = ScaleLeft
    tab性质.Width = ScaleWidth
    tab性质.Top = sngTop
    
    msh记录_S.Left = ScaleLeft
    msh记录_S.Width = ScaleWidth
    msh记录_S.Top = tab性质.Top + tab性质.Height
    
    If picSplitH.Visible = False Then
        '当显示预交记录时没有明细
        msh记录_S.Height = IIf(sngBottom - msh记录_S.Top > 0, sngBottom - msh记录_S.Top, 0)
        Exit Sub
    Else
        If msh记录_S.Height > ScaleHeight - msh记录_S.Top - IIf(stbThis.Visible, stbThis.Height, 0) Then
            msh记录_S.Height = ScaleHeight - msh记录_S.Top - IIf(stbThis.Visible, stbThis.Height, 0)
        End If
    End If
    picSplitH.Left = ScaleLeft
    picSplitH.Width = ScaleWidth
    picSplitH.Top = msh记录_S.Top + msh记录_S.Height
    
    msh明细.Left = ScaleLeft
    msh明细.Top = picSplitH.Top + picSplitH.Height
    msh明细.Height = IIf(sngBottom - msh明细.Top > 0, sngBottom - msh明细.Top, 0)
    
    msh分档.Left = IIf(ScaleWidth - msh分档.Width > 0, ScaleWidth - msh分档.Width, 0)
    picSplitV.Left = msh分档.Left - picSplitV.Width
    If msh分档.Visible = False Then
        '当显示收费记录时，没有分档统筹数据
        msh明细.Width = IIf(ScaleWidth - msh明细.Left > 0, ScaleWidth - msh明细.Left, 0)
        Exit Sub
    Else
        msh明细.Width = IIf(picSplitV.Left - msh明细.Left > 0, picSplitV.Left - msh明细.Left, 0)
    End If
    
    msh分档.Top = msh明细.Top
    msh分档.Height = msh明细.Height
    picSplitV.Top = msh明细.Top
    picSplitV.Height = msh明细.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
        If mint性质 = 2 Then
            saveFlexState msh记录_S, "费用结算_大连"
        End If
    End If
End Sub

Private Sub mnuBalanceBill_Click()
    '格式：1-门诊;2-门诊规定病;3-住院
    '提取单个病人的结算单
    Const strBill As String = "ZL1_INSIDE_1605_10"
    Dim lng病人ID As Long, lng结帐ID As Long
    Dim int业务类型 As Integer
    Dim str业务序列号 As String
    On Error GoTo ErrHand
    
    lng病人ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col病人ID))
    If lng病人ID = 0 Then Exit Sub
    lng结帐ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng结帐ID = 0 Then Exit Sub
    
    Select Case mint险类
    Case TYPE_沈阳市
        If Not 提取结算单_沈阳市(lng病人ID, lng结帐ID, int业务类型, str业务序列号) Then Exit Sub
        '调报表预览
        Call ReportOpen(gcnOracle, glngSys, strBill, Me, "业务序列号=" & str业务序列号, "ReportFormat=" & int业务类型, 1)
    Case Else
        Exit Sub
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuBalanceCollect_Click()
    '格式：1-门诊;2-门诊规定病;3-住院
    '提取结算汇总表，用于医院打印与中心的对帐单用
    Const strBill As String = "ZL1_INSIDE_1605_11"
    On Error GoTo ErrHand
    
    Select Case mint险类
    Case TYPE_沈阳市
        If Not 提取结算表_沈阳市() Then Exit Sub
        
        '调报表预览
        Call ReportOpen(gcnOracle, glngSys, strBill, Me)
    Case Else
        Exit Sub
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditCacel_Click()
        
    If MsgBox("你是否真的要放弃所编辑的最高限额吗?", vbQuestion + vbDefaultButton1 + vbYesNo) <> vbYes Then Exit Sub
    mblnEdit = False
    mblnChange = False
    MoveEditCtl
    SetMenu
    mblnChange = False
    Call tab性质_Click
End Sub

Private Sub mnuEditClinic_Click()
    If frm费用报销.ShowME(1) And mint性质 = 1 Then Call FillList
End Sub
    
Private Sub mnuEditDelete_Click()
    Dim lng记录id As Long, lng结帐ID As Long, lng病人ID As Long
    Dim int住院次数 As Integer, lng年度 As Long, lng当前年度 As Long
    Dim cur帐户余额_年度 As Currency, cur统筹累计_年度 As Currency
    Dim cur帐户支付 As Currency, cur进入统筹 As Currency '门诊仅下帐户；住院仅下统筹
    Dim int本院 As Integer, int外院 As Integer, int当前住院次数_本院 As Integer, int当前住院次数_外院 As Integer
    Dim bln本院 As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If mint性质 = 3 Then Exit Sub
    If Trim(msh记录_S.TextMatrix(msh记录_S.Row, col收退标志)) = "退" Then
        MsgBox "不能对作废单据执行作废操作（只能对收费类型的单据进行作废）！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrSQL = "Select Count(*) Records From 保险结算记录 Where 险类=25 And 支付顺序号=" & Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp!Records = 1 Then
        MsgBox "该单据已被作废，不能再次作废！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("你确定要做废该结算记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '作废
    lng当前年度 = Format(zlDataBase.Currentdate, "yyyy")
    lng记录id = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng记录id = 0 Then Exit Sub
    lng结帐ID = zlDataBase.GetNextId("病人结帐记录")
    
    gcnOracle.BeginTrans
    '保存负的保险结算记录
    With rsTemp
        gstrSQL = "Select * From 保险结算记录 Where 险类=" & gintInsure & " And 记录ID=" & lng记录id
        Call OpenRecordset(rsTemp, Me.Caption)
        int住院次数 = Nvl(!住院次数, 1) - 1
        lng病人ID = !病人ID
        lng年度 = !年度
        cur帐户支付 = Nvl(!个人帐户支付, 0)
        cur进入统筹 = Nvl(!进入统筹金额, 0)
        
        '只能作废最后一次的结算记录
        gstrSQL = "Select Nvl(住院次数累计,0) 本院,Nvl(外院住院次数,0) 外院 From 帐户年度信息 Where 年度=" & lng当前年度 & " ANd 病人ID =" & lng病人ID
        Call OpenRecordset(rsTest, Me.Caption)
        If rsTest.EOF Then
            int当前住院次数_本院 = 0
            int当前住院次数_外院 = 0
        Else
            int当前住院次数_本院 = rsTest!本院 - IIf(mint性质 = 2, 1, 0)
            int当前住院次数_外院 = rsTest!外院 - IIf(mint性质 = 2, 1, 0)
        End If
        If lng年度 <> Format(zlDataBase.Currentdate, "yyyy") Then
            MsgBox "不能冲销以往年度的单据！", vbInformation, gstrSysName
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        gstrSQL = "Select * From 保险结算记录 Where 险类=" & gintInsure & " And 记录ID=" & lng记录id
        Call OpenRecordset(rsTemp, Me.Caption)
        gstrSQL = "zl_保险结算记录_insert(" & !性质 & "," & lng结帐ID & ",25," & !病人ID & "," & _
            lng年度 & ",0,0,0,0," & Nvl(!住院次数, 0) & "," & -1 * Nvl(!起付线, 0) & ",0," & -1 * Nvl(!实际起付线, 0) & "," & _
            -1 * Nvl(!发生费用金额, 0) & "," & -1 * Nvl(!全自付金额, 0) & "," & -1 * Nvl(!首先自付金额, 0) & "," & -1 * Nvl(!进入统筹金额, 0) & "," & -1 * Nvl(!统筹报销金额, 0) & ",0," & _
            0 & "," & -1 * cur帐户支付 & ",'" & lng记录id & "',null,null,null,null,'" & gstrUserName & "')" '支付顺序号用来保存被冲销的记录ID
        Call ExecuteProcedure("保存住院结算数据")
        '保存负的保险报销记录(大类)
        gstrSQL = "Select * From 保险报销记录 Where 记录ID=" & lng记录id
        Call OpenRecordset(rsTemp, Me.Caption)
        Do While Not .EOF
            gstrSQL = "ZL_保险报销记录_INSERT(" & !性质 & "," & lng结帐ID & "," & _
            "'" & !大类编码 & "','" & !大类名称 & "'," & !统筹比额 & "," & _
            "" & !特准定额 & "," & !特准天数 & "," & -1 * !费用总额 & "," & -1 * !报销总额 & ")"
            Call ExecuteProcedure("保存大类报销数据")
            bln本院 = (!性质 = 1)
            .MoveNext
        Loop
        
        '如果不是冲销最后一次结算记录，禁止!
        If mint性质 = 2 Then
            If bln本院 Then
                If int当前住院次数_本院 <> int住院次数 Then
                    MsgBox "只能作废最后一次结算的单据！", vbInformation, gstrSysName
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            Else
                If int当前住院次数_外院 <> int住院次数 Then
                    MsgBox "只能作废最后一次结算的单据！", vbInformation, gstrSysName
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            End If
        End If
        
        gstrSQL = "Select * From 保险结算计算 Where 结帐ID=" & lng记录id
        Call OpenRecordset(rsTemp, Me.Caption)
        '保存负的分档报销记录
        Do While Not .EOF
            gstrSQL = "ZL_保险结算计算_INSERT(" & lng结帐ID & "," & !档次 & "," & -1 * !进入统筹金额 & "," & -1 * !统筹报销金额 & "," & !比例 & ")"
            Call ExecuteProcedure("保存分档报销明细")
            .MoveNext
        Loop
    End With
        
    '取本年度帐户余额与进入统筹
    gstrSQL = " Select Nvl(帐户增加累计,0) 帐户余额,Nvl(进入统筹累计,0) 进入统筹" & _
              " ,Nvl(住院次数累计,0) 本院,Nvl(外院住院次数,0) 外院" & _
              " From 帐户年度信息 " & _
              " Where 年度=" & lng当前年度 & " And 病人ID=" & lng病人ID ' gComInfo_眉山.病人ID
    Call OpenRecordset(rsTemp, Me.Caption)
    
    cur帐户余额_年度 = 0
    cur统筹累计_年度 = 0
    int本院 = 0: int外院 = 0
    If Not rsTemp.EOF Then
        cur帐户余额_年度 = Nvl(rsTemp!帐户余额, 0)
        cur统筹累计_年度 = Nvl(rsTemp!进入统筹, 0)
        int本院 = rsTemp!本院
        int外院 = rsTemp!外院
    End If
    If bln本院 Then
        int本院 = int当前住院次数_本院
    Else
        int外院 = int当前住院次数_外院
    End If
    '由于下个人帐户函数会自动更新帐户余额，因此本处不更新
'    cur帐户余额_年度 = cur帐户余额_年度 + cur帐户支付
    '门诊下帐户的算法和住院一样，只是由帐户支付，因此也存在进入统筹，所以此处要判断一下
    cur统筹累计_年度 = cur统筹累计_年度 - IIf(mint性质 = 1, 0, cur进入统筹)
    
    gstrSQL = "zl_帐户年度信息_Insert(" & lng病人ID & ",25," & lng年度 & _
              "," & cur帐户余额_年度 & ",0," & cur统筹累计_年度 & ",0," & int本院 & "," & int外院 & ")"
    Call ExecuteProcedure("更新住院次数")
    
    '给帐户上帐
    If cur帐户支付 <> 0 Then
        If Not 下个人帐户(lng病人ID, cur帐户支付) Then Exit Sub 'gComInfo_眉山.病人ID
    End If
    
    gcnOracle.CommitTrans
    Call FillList
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditIn_Patient_Click()
    If frm费用报销.ShowME(1, True) And mint性质 = 2 Then Call FillList
End Sub

Private Sub mnuEditSave_Click()
    
    If Save保存最高限额(mint险类) = False Then Exit Sub
    
    mblnChange = False
    mblnEdit = False
    MoveEditCtl
    SetMenu
    mblnChange = False
    
    Call cmb险类_Click
End Sub

Private Sub mnuEditView_Click()
    Dim lng结帐ID  As Long
    
    If mint性质 = 1 Or mint性质 = 2 Then
        lng结帐ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
        If lng结帐ID = 0 Then Exit Sub
        Call frm费用报销.ShowME(2, (mint性质 = 2), lng结帐ID)
    End If
End Sub

Private Sub mnuEditXE_Click()

    '  '主要录入大连的最高限额
    '    Dim lng病人id As Long
    '    Dim strIdentify As String
    '    Dim bytType As Byte
    '    Dim cls医保 As New clsInsure
    '    Dim lng性质 As Long
    '    Dim lng记录id As Long
    '    Dim int性质 As Long
    '    Dim frmMain As New frmIdentify大连
    '
    '    lng记录id = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    '    If lng记录id = 0 Then Exit Sub
    '
    '    lng病人id = Val(msh记录_S.TextMatrix(msh记录_S.Row, col病人ID))
    '    If lng病人id = 0 Then Exit Sub
    '
    '    int性质 = Val(Mid(tab性质.SelectedItem.Key, 2))
    '
    '    strIdentify = frmMain.GetPatient(9, lng病人id, int性质, lng记录id)
    '
    '    mint性质 = 0
    '
    '    If strIdentify <> "" Then
    '        tab性质_Click
    '    End If
    
    If mrs结算记录.RecordCount = 0 Then Exit Sub
    
    mblnEdit = True
    'msh记录_S.SelectionMode = flexSelectionFree
    MoveEditCtl
    mblnChange = False
    SetMenu
    
End Sub

Private Sub mnuFileBalance_Click()
    Dim lng病人ID As Long, lng记录id As Long
    Dim str医院编码 As String, str业务序列号 As String
    Dim rsTemp As New ADODB.Recordset
    '只有沈阳铁路医保存在该功能，用以调用接口获取某次结算的信息，存入临时表中以供打印之需
    lng病人ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col病人ID))
    lng记录id = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng病人ID = 0 Then Exit Sub
    If lng记录id = 0 Then Exit Sub
    
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_沈阳市
    Call OpenRecordset(rsTemp, "读取医院编码")
    str医院编码 = Nvl(rsTemp!医院编码)
    If Trim(str医院编码) = "" Then
        MsgBox "请设置了医院编码后再使用该功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "Select * From 保险结算记录 Where 险类=" & TYPE_沈阳市 & " And 记录ID=" & lng记录id
    Call OpenRecordset(rsTemp, "获取业务序列号")
    If rsTemp.EOF Then
        MsgBox "没有找到任何结算记录！", vbInformation, gstrSysName
        Exit Sub
    End If
    str业务序列号 = Nvl(rsTemp!支付顺序号)
    If Trim(str业务序列号) = "" Then
        MsgBox "保险结算数据错误，无法继续！（业务序列号不能为空）", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '20031228:周韬:加入结帐ID
    Call GetBalance(lng病人ID, lng记录id, str业务序列号, str医院编码)
End Sub

Private Sub mnuFileBill_Click()
    Dim lng结算ID  As Long
    lng结算ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng结算ID = 0 Then Exit Sub
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605_" & mint性质, Me, "险类=" & mint险类, "记录ID=" & lng结算ID, 2)
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuPrintBalance_Click()
    Dim str结算流水号 As String
    Dim lng结帐ID As Long
    Dim rsTemp As New ADODB.Recordset
    '打印票据
    On Error GoTo ErrHand
    
    lng结帐ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng结帐ID = 0 Then Exit Sub
    
    '先获取指定结算记录的结算交易流水号（备注第三个字段）
    gstrSQL = "Select 备注 From 保险结算记录 Where 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "取结算交易流水号")
    If rsTemp.RecordCount = 0 Then
        MsgBox "未找到与保险结算相关的记录！", vbInformation, gstrSysName
        Exit Sub
    End If
    str结算流水号 = Split(rsTemp!备注, "|")(2)
    If str结算流水号 = "" Then
        MsgBox "结算交易流水号为空！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Not 医保初始化_重庆银海版 Then Exit Sub
    Call 调用接口_准备_重庆银海版("21", str结算流水号)
    Call 调用接口_重庆银海版
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuViewFind_Click()
    If frm医保结算查找.GetTimeScope(mdatBegin, mdatEnd, mstrCardCond, Me) = True Then
        Call FillList
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    
    Call FillList
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
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbrThis.ButtonHeight
    Form_Resize
End Sub

Private Sub msh记录_S_EnterCell()
    '选择某个帐户,则提取相关信息
    Select Case mint险类
    Case type_大连开发区, type_大连市
        MoveEditCtl
        If mblnEdit Then Exit Sub
    End Select
    Call FillDetail
End Sub

Private Sub msh记录_S_KeyDown(KeyCode As Integer, Shift As Integer)
    If mint险类 = TYPE_四川眉山 Then
        If KeyCode = vbKeyReturn Then Call mnuEditView_Click
    End If
End Sub

Private Sub msh记录_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strSort As String
    Err = 0
    On Error Resume Next
    If Button = 1 Then
        '按列头排序
        If msh记录_S.MouseRow = 0 Then
            If mblnEdit And (mint险类 = 82 Or mint险类 = 83) Then Exit Sub
            If mint性质 = 2 And (mint险类 = 82 Or mint险类 = 83) Then
                strSort = "科室," & msh记录_S.TextMatrix(0, msh记录_S.MouseCol)
            Else
                strSort = msh记录_S.TextMatrix(0, msh记录_S.MouseCol)
            End If
            If strSort = "住院号" And mint性质 = 1 Then strSort = "门诊号"
            
            If strSort = "" Then Exit Sub
            If mrs结算记录.Sort = strSort Then
                mrs结算记录.Sort = strSort & " DESC"
            Else
                mrs结算记录.Sort = strSort
            End If
            Call 绑定数据(msh记录_S, mrs结算记录)
        End If
    End If
End Sub

Private Sub msh记录_S_Scroll()
    If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
        If mblnNOScroll Then Exit Sub
        MoveEditCtl
    End If
End Sub

Private Sub tab性质_Click()
    Dim int性质 As Integer
    Dim sngHeight As Single
    Call 权限控制
    int性质 = Val(Mid(tab性质.SelectedItem.Key, 2))
    If mint性质 = int性质 Then Exit Sub
    
    
    mint性质 = int性质
    
    Select Case mint性质
        Case 1 '收费
            msh分档.Visible = False
            picSplitV.Visible = False
            
            If msh明细.Visible = False Then
                '前一个状态是显示预交
                msh明细.Visible = True
                picSplitH.Visible = True
                
                sngHeight = ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - msh记录_S.Top
                
                If sngHeight - msh记录_S.Height < 1000 Then
                    msh记录_S.Height = msh记录_S.Height / 2
                End If
            End If
        Case 2 '结算
            msh分档.Visible = True
            picSplitV.Visible = True
            If msh明细.Visible = False Then
                '前一个状态是显示预交
                msh明细.Visible = True
                picSplitH.Visible = True
                msh记录_S.Height = msh记录_S.Height / 2
            End If
            
        
        Case 3 '
            picSplitH.Visible = False
            msh明细.Visible = False
            msh分档.Visible = False
            picSplitV.Visible = False
    End Select
    '重新调整
    Call Form_Resize
    '显示数据
    Call FillList
   ' SetMenu
      
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Find"
            mnuViewFind_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "编辑"
            mnuEditXE_Click
        Case "保存"
            mnuEditSave_Click
        Case "放弃"
            mnuEditCacel_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFileDetail_Click()
    Dim lng结算ID As Long
    
    lng结算ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng结算ID <> 0 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605", Me, "险类=" & mint险类, "ID=" & lng结算ID, 1)
    End If
End Sub

Private Sub mnuFileBatch_Click()
    Dim lngRow As Integer, int险类 As Integer
    Dim lng结算ID As Long
    
    '批理处理结算记录
    For lngRow = 1 To msh记录_S.Rows - 1
        lng结算ID = Val(msh记录_S.TextMatrix(lngRow, col记录ID))
        If lng结算ID <> 0 Then
            Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605", Me, "险类=" & mint险类, "ID=" & lng结算ID, 1)
        End If
    Next
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
    '功能：输入出列表
    '参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh记录_S.Row
    
    '表头
    objOut.Title.Text = "医保个人费用结算清单（" & tab性质.SelectedItem.Caption & "）"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDataBase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = msh记录_S
    
    '输出
    msh记录_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh记录_S.Redraw = True
    
    msh记录_S.Row = intRow
    msh记录_S.Col = 0: msh记录_S.ColSel = msh记录_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp > msh记录_S.Top + 1000 And (msh明细.Top + msh明细.Height) - (sngTemp + picSplitH.Height) > 1000 Then
            picSplitH.Top = sngTemp
            msh记录_S.Height = picSplitH.Top - msh记录_S.Top
            Form_Resize
        End If
    End If
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > msh明细.Left + 1000 And ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            msh分档.Width = ScaleWidth - (sngTemp + picSplitV.Width)
            Form_Resize
        End If
    End If
End Sub

Private Function FillList() As Boolean
    '提取所有帐户(如果权限允许,则提出密码字段)的数据
    Dim strBegin As String
    Dim strEnd As String
    
    strBegin = "to_date('" & Format(mdatBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') "
    strEnd = "to_date('" & Format(mdatEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') "
        
        
    
    If mrs结算记录.State = adStateOpen Then mrs结算记录.Close
    
    MousePointer = vbHourglass
    On Error GoTo errHandle
    
    Call GetSpecialSQL(mint性质, strBegin, strEnd)
    
    mrs结算记录.Sort = ""
    Call OpenRecordset(mrs结算记录, Me.Caption)
    Call 绑定数据(msh记录_S, mrs结算记录)
    
    Call FillDetail
    FillList = True
    MousePointer = vbDefault
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Function

Private Sub 绑定数据(mshBind As MSHFlexGrid, rsBind As ADODB.Recordset)
    Dim lngCol As Long
    
    '将帐户数据装入FLEXGRID，绑定数据
    If (mint险类 = type_大连开发区 Or mint险类 = type_大连市) And mint性质 = 2 Then
        saveFlexState msh记录_S, "费用结算_大连"
        '
    End If
    
    If mshBind Is msh记录_S Then
        Call Init记录Table '由于不同的数据其表格内容很大程度上不同，所以每次都初始化
    End If
    
    With mshBind
        If rsBind.RecordCount <> 0 Then
            Set .DataSource = rsBind
            DoEvents
            .Col = 0
            .Row = .FixedRows - 1
            .ColSel = .Cols - 1
            .RowSel = .Row
            .FillStyle = flexFillRepeat
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
            .Row = .FixedRows: .Col = 0
            .ColSel = .Cols - 1: .RowSel = .Row
            If mint性质 = 2 Then
                If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
                    Call SetCOLAlign_大连
                End If
            End If
        Else
            Set .DataSource = Nothing
            .Rows = 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(1, lngCol) = ""
            Next
            If mint性质 = 2 Then
                If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
                    Call SetCOLAlign_大连
                End If
            End If
            
        End If
        
        If mshBind Is msh记录_S Then
            '隐藏多余的列
            If mcol中心("K" & mint险类) = "0" Then
                .ColWidth(col中心) = 0
            Else
                If .ColWidth(col中心) = 0 Then
                    .ColWidth(col中心) = 1000
                End If
            End If
        End If
    End With
End Sub

Private Sub Init记录Table()
    Dim lngCol As Integer
    
    '设置格式
    With msh记录_S
        .Rows = 2
        .Cols = 19 '为了设置一些公共的列
        For lngCol = 0 To .Cols - 1
            .ColPosition(lngCol) = 0
        Next
        
        .TextMatrix(0, col记录ID) = "中心"
        .TextMatrix(0, col单据号) = "单据号"
        .TextMatrix(0, col中心) = "中心"
        .TextMatrix(0, col卡号) = "卡号"
        .TextMatrix(0, col病人ID) = "病人ID"
        .TextMatrix(0, col姓名) = "姓名"
        .TextMatrix(0, col身份) = "身份"
        .TextMatrix(0, col性别) = "性别"
        .TextMatrix(0, col年龄) = "年龄"
        .TextMatrix(0, col收退标志) = "收退标志"
        .TextMatrix(0, col个人帐户) = "个人帐户"
        .ColWidth(col记录ID) = 0
        .ColWidth(col单据号) = 900
        .ColWidth(col中心) = 0
        .ColWidth(col卡号) = 900
        .ColWidth(col病人ID) = 800
        .ColWidth(col姓名) = 800
        .ColWidth(col身份) = 600
        .ColWidth(col性别) = 400
        .ColWidth(col年龄) = 400
        .ColWidth(col收退标志) = 855
        .ColWidth(col个人帐户) = 930
        
        .ColWidth(col住院号) = 800
        .ColWidth(col操作员姓名) = 1200
        .ColWidth(col登记时间) = 1200
        Select Case mint性质
            Case 1 '-收费
                .Cols = 16
                .TextMatrix(0, col住院号) = "住院号"
                .TextMatrix(0, col科室) = "开单科室"
                .TextMatrix(0, col操作员姓名) = "收费员"
                .TextMatrix(0, col登记时间) = "收费时间"
                .TextMatrix(0, col发生费用) = "发生费用"
                
                .ColWidth(col科室) = 1200
                .ColWidth(col发生费用) = 930
                
                '改变某些列的显示顺序
                .ColPosition(col个人帐户) = col操作员姓名
                .ColPosition(col发生费用) = col操作员姓名
            Case 2 '-结算（包括住院结算、特殊门诊结算）
                Select Case mint险类
                Case type_大连开发区, type_大连市
                    Call ReSetTableCOl_大连
                Case Else
                    .TextMatrix(0, col住院号) = "门诊号"
                    .TextMatrix(0, col科室) = "开单科室"
                    .TextMatrix(0, col操作员姓名) = "结帐人"
                    .TextMatrix(0, col登记时间) = "结帐时间"
                    .TextMatrix(0, col发生费用) = "发生费用"
                    .TextMatrix(0, col实际起付线) = "实际起付线"
                    .TextMatrix(0, col进入统筹) = "进入统筹"
                    .TextMatrix(0, col统筹报销) = "统筹报销"
                        
                    .ColWidth(col科室) = 0
                    .ColWidth(col发生费用) = 930
                    .ColWidth(col实际起付线) = 1120
                    .ColWidth(col进入统筹) = 930
                    .ColWidth(col统筹报销) = 930
                    '改变某些列的显示顺序
                    .ColPosition(col个人帐户) = col操作员姓名
                    .ColPosition(col发生费用) = col操作员姓名
                    .ColPosition(col实际起付线) = col登记时间
                    .ColPosition(col进入统筹) = col登记时间 + 1
                    .ColPosition(col统筹报销) = col登记时间 + 1
                End Select
            Case 3 '-预交
                .Cols = 15
                .TextMatrix(0, col住院号) = "住院号"
                .TextMatrix(0, col科室) = "科室"
                .TextMatrix(0, col操作员姓名) = "收款人"
                .TextMatrix(0, col登记时间) = "收款时间"
                
                .ColWidth(col科室) = 1200
                
                '改变某些列的显示顺序
                .ColPosition(col个人帐户) = col操作员姓名
        End Select
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub
Private Sub ReSetTableCOl_大连(Optional ByVal blnOlnyColAlignment As Boolean = False)
    '重新对列进行排列,主要针对大连且为结算
    '        结帐ID,单据号,中心,卡号,病人ID,医保号,住院号,住院次数,姓名,身份,性别,年龄,科室,统筹总额,最高限额,结帐人,
    '        结帐时间,收退标志,发生费用,起付线,个人帐户,基本统筹支付,基本统筹自付,补充统筹支付,补充统筹自付,补助保险支付,非补助保险支付,
    '        补助帐户支付 , 保险范围外自付
    Dim i As Long
   With msh记录_S
        
        .Rows = 2
        .Clear
        .Cols = 29
        For i = 0 To .Cols - 1
            .ColPosition(i) = 0
        Next
        
        .TextMatrix(0, col记录ID) = "结帐ID": .ColWidth(col记录ID) = 0
        .TextMatrix(0, col单据号) = "单据号"
        .TextMatrix(0, col中心) = "中心"
        .TextMatrix(0, col卡号) = "卡号"
        .TextMatrix(0, col病人ID) = "病人ID": .ColWidth(col病人ID) = 0
        
        i = 5:     .TextMatrix(0, i) = "医保号": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "住院号": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "住院次数": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "姓名": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "身份": .ColAlignment(i) = 4: .ColWidth(i) = 800
        i = i + 1: .TextMatrix(0, i) = "性别": .ColAlignment(i) = 4: .ColWidth(i) = 600
        i = i + 1: .TextMatrix(0, i) = "年龄": .ColAlignment(i) = 4: .ColWidth(i) = 600
        i = i + 1: .TextMatrix(0, i) = "科室": .ColAlignment(i) = 1: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "统筹总额": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        
        i = i + 1: .TextMatrix(0, i) = "最高限额": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "门诊号": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "结帐人": .ColAlignment(i) = 4: .ColWidth(i) = 1000
        i = i + 1: .TextMatrix(0, i) = "结帐时间": .ColAlignment(i) = 4: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "收退标志": .ColAlignment(i) = 4: .ColWidth(i) = 800
        i = i + 1: .TextMatrix(0, i) = "发生费用": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "实际起付线": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "个人帐户": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "基本统筹支付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "基本统筹自付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "补充统筹支付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "补充统筹自付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "非补助保险支付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "补助帐户支付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
        i = i + 1: .TextMatrix(0, i) = "保险范围外自付": .ColAlignment(i) = 7: .ColWidth(i) = 1200
'        For i = 0 To .Cols - 1
'            .ColAlignmentFixed(i) = 4
'        Next
        
        '恢复例设置
        RestoreFlexState msh记录_S, "费用结算_大连"
        .ColWidth(col记录ID) = 0
        .ColWidth(col病人ID) = 0
    End With
End Sub
Private Function FillDetail()
    Dim lngCount As Long, lng结帐ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    Call SetMenu
    
    If mint性质 = 3 Then
        '预交不处理
        Exit Function
    End If
    
    '清除相关信息
    msh明细.Rows = 2
    msh分档.Rows = 2
    For lngCount = 0 To msh明细.Cols - 1
        msh明细.TextMatrix(1, lngCount) = ""
    Next
    For lngCount = 0 To msh分档.Cols - 1
        msh分档.TextMatrix(1, lngCount) = ""
    Next
    
    lng结帐ID = Val(msh记录_S.TextMatrix(msh记录_S.Row, col记录ID))
    If lng结帐ID = 0 Then
        Exit Function
    End If
    
    '提取结算记录的明细数据
    gstrSQL = _
        " Select A.NO,C.类别,B.名称,B.规格,A.计算单位 as 单位," & _
        " Ltrim(To_Char(Avg(Nvl(A.付数,1)*decode(A.记录状态,2,-1,1)*A.数次),'999990.000')) as 数次, " & _
        " Ltrim(To_Char(Sum(A.标准单价),'999990.000')) as 单价, " & _
        " Ltrim(To_Char(Sum(decode(A.记录状态,2,-1,1)*A.实收金额),'999990.00')) as 实收金额, " & _
        " Ltrim(To_Char(Sum(decode(A.记录状态,2,-1,1)*A.统筹金额),'999990.00')) as 统筹金额, " & _
        IIf(mint性质 = 2, " Ltrim(To_Char(Sum(A.结帐金额),'999990.00')) as 结帐金额, ", "") & _
        " E.名称 as 医保大类,B.费用类型 as 类型," & _
        " Decode(A.记录状态,2,'退','收') as 收退" & _
        " From 病人费用记录 A,收费细目 B,收费类别 C,保险支付大类 E " & _
        " Where A.收费细目ID=B.ID and A.收费类别=C.编码 " & _
        "       And A.保险大类ID=E.ID(+) And A.结帐ID=" & lng结帐ID & _
        " Group by mod(A.记录性质,10),A.NO,Decode(A.价格父号,NULL,A.序号,A.价格父号),A.记录状态,A.收费类别,C.类别,B.名称,B.规格,A.计算单位,B.费用类型,E.名称" & _
        " Order by A.NO,Decode(A.价格父号,NULL,A.序号,A.价格父号)"
    
    Call OpenRecordset(rsTemp, Me.Caption)
    Call 绑定数据(msh明细, rsTemp)
    
    If mint性质 = 1 Then
        '收费不处理
        Exit Function
    End If
    
    '提取结算记录的分档数据
    If mint险类 = TYPE_四川眉山 Then
        gstrSQL = _
            " Select D.名称," & _
            "   Ltrim(To_Char(A.进入统筹金额,'999990.00')) as 进入统筹, " & _
            "   Ltrim(To_Char(A.统筹报销金额,'999990.00')) as 统筹报销, " & _
            "   Ltrim(To_Char(A.比例,'999990.00')) as 比例 " & _
            " From 保险结算计算 A,保险结算记录 B,保险帐户 C,保险费用档 D " & _
            " Where B.记录ID=" & lng结帐ID & " and B.性质=2 And B.险类=" & mint险类 & _
            "   And B.病人ID=C.病人ID and C.险类=B.险类 and D.险类=C.险类 and D.中心=C.中心 " & _
            "   And A.结帐ID=" & lng结帐ID & "and A.档次=D.档次(+) "
    ElseIf mint险类 = TYPE_泸州市 Then
        gstrSQL = _
            " Select D.名称," & _
            "   Ltrim(To_Char(decode(E.记录状态,2,-1,1)*A.进入统筹金额,'999990.00')) as 进入统筹, " & _
            "   Ltrim(To_Char(decode(E.记录状态,2,-1,1)*A.统筹报销金额,'999990.00')) as 统筹报销, " & _
            "   Ltrim(To_Char(A.比例,'999990.00')) as 比例 " & _
            " From 保险结算计算 A,保险结算记录 B,保险帐户 C,保险费用档 D,病人结帐记录 E " & _
            " Where E.ID=B.记录ID And B.记录ID=" & lng结帐ID & " and B.性质=2 And B.险类=" & mint险类 & _
            "   And B.病人ID=C.病人ID and C.险类=B.险类 and D.险类=C.险类 and D.中心=C.中心 " & _
            "   And A.结帐ID=" & lng结帐ID & "and A.档次=D.档次(+) and c.在职=d.在职 "
    Else
        gstrSQL = _
            " Select D.名称," & _
            "   Ltrim(To_Char(decode(E.记录状态,2,-1,1)*A.进入统筹金额,'999990.00')) as 进入统筹, " & _
            "   Ltrim(To_Char(decode(E.记录状态,2,-1,1)*A.统筹报销金额,'999990.00')) as 统筹报销, " & _
            "   Ltrim(To_Char(A.比例,'999990.00')) as 比例 " & _
            " From 保险结算计算 A,保险结算记录 B,保险帐户 C,保险费用档 D,病人结帐记录 E " & _
            " Where E.ID=B.记录ID And B.记录ID=" & lng结帐ID & " and B.性质=2 And B.险类=" & mint险类 & _
            "   And B.病人ID=C.病人ID and C.险类=B.险类 and D.险类=C.险类 and D.中心=C.中心 " & _
            "   And A.结帐ID=" & lng结帐ID & "and A.档次=D.档次(+) "
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    Call OpenRecordset(rsTemp, Me.Caption)
    Call 绑定数据(msh分档, rsTemp)
End Function

Private Sub InitTable()
    Dim lngCol As Integer
    
    '设置格式
    With msh明细
        .Rows = 2
        .Cols = 12 '为了设置一些公共的列
        .TextMatrix(0, det收费类别) = "收费类别"
        .TextMatrix(0, det收费细目) = "收费细目"
        .TextMatrix(0, det规格) = "规格"
        .TextMatrix(0, det单位) = "单位"
        .TextMatrix(0, det数次) = "数次"
        .TextMatrix(0, det单价) = "单价"
        .TextMatrix(0, det实收金额) = "实收金额"
        .TextMatrix(0, det统筹金额) = "统筹金额"
        .TextMatrix(0, det医保大类) = "医保大类"
        .TextMatrix(0, det费用类型) = "费用类型"
        .TextMatrix(0, det收退) = "收退"
        .TextMatrix(0, det状态) = "状态"
        
        .ColWidth(det收费类别) = 600
        .ColWidth(det收费细目) = 1000
        .ColWidth(det规格) = 900
        .ColWidth(det单位) = 600
        .ColWidth(det数次) = 800
        .ColWidth(det单价) = 800
        .ColWidth(det实收金额) = 930
        .ColWidth(det统筹金额) = 930
        .ColWidth(det医保大类) = 800
        .ColWidth(det费用类型) = 800
        .ColWidth(det收退) = 600
        .ColWidth(det状态) = 600
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
    
    With msh分档
        .Rows = 2
        .Cols = 4 '为了设置一些公共的列
        .TextMatrix(0, 0) = "费用档"
        .TextMatrix(0, 1) = "进入统筹"
        .TextMatrix(0, 2) = "统筹报销"
        .TextMatrix(0, 3) = "比例"
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub 权限控制()
    If InStr(mstrPrivs, "结算表") = 0 Then
        mnuFileBatch.Visible = False
        mnuFileDetail.Visible = False
        mnuFileSplitReport.Visible = False
    End If
    If InStr(mstrPrivs, "报销") <> 0 Then
        mnuEdit.Visible = True
        mnuFileBill.Visible = True
    Else
        mnuEdit.Visible = False
        mnuFileBill.Visible = False
    End If
    
    If mint险类 = type_大连开发区 Or mint险类 = type_大连市 Then
        mnuEditXE.Visible = True
        mnuEditClinic.Visible = mnuEdit.Visible
        mnuEditIn_Patient.Visible = mnuEdit.Visible
        mnuEditDelete.Visible = mnuEdit.Visible
        mnuEditView.Visible = mnuEdit.Visible
        mnuEdit.Visible = True
        mnusplit1.Visible = False
        mnuEditXE.Visible = True
        mnuEditSave.Visible = True
        mnuEditCacel.Visible = True
        mnuEditSplit.Visible = False
        tbrThis.Buttons("编辑").Visible = True
        tbrThis.Buttons("Split").Visible = True
        tbrThis.Buttons("保存").Visible = True
        tbrThis.Buttons("放弃").Visible = True
        tbrThis.Buttons("Edit_1").Visible = True
        
    Else
        If mnuEdit.Visible = False Then
        Else
            mnusplit1.Visible = False
            mnuEditXE.Visible = False
            mnuEditSave.Visible = False
            mnuEditCacel.Visible = False
            mnuEditSplit.Visible = False
        End If
        tbrThis.Buttons("编辑").Visible = False
        tbrThis.Buttons("Split").Visible = False
        tbrThis.Buttons("保存").Visible = False
        tbrThis.Buttons("放弃").Visible = False
        tbrThis.Buttons("Edit_1").Visible = False
        txtEdit.Visible = False
    End If
End Sub

Private Sub SetMenu()
    Dim blnData As Boolean
    Dim lng性质 As Long
    blnData = (mrs结算记录.RecordCount > 0)
    stbThis.Panels(2).Text = "当前共有" & mrs结算记录.RecordCount & "个医保帐户"

    tbrThis.Buttons("Print").Enabled = blnData
    tbrThis.Buttons("Preview").Enabled = blnData
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuFileBatch.Enabled = blnData And (mint性质 = 2)
    mnuFileDetail.Enabled = mnuFileBatch.Enabled
    
    '主要应用于大连医保
    Select Case mint险类
    Case type_大连开发区, type_大连市
        lng性质 = Val(Mid(tab性质.SelectedItem.Key, 2))
        mnuEditXE.Enabled = Not mblnEdit And lng性质 = 2 And blnData
        mnuEditSave.Enabled = mblnEdit And mblnChange And lng性质 = 2 And blnData
        mnuEditCacel.Enabled = mblnEdit And lng性质 = 2 And blnData
        tbrThis.Buttons("编辑").Enabled = mnuEditXE.Enabled
        tbrThis.Buttons("保存").Enabled = mnuEditSave.Enabled
        tbrThis.Buttons("放弃").Enabled = mnuEditCacel.Enabled
        
        tbrThis.Buttons("Find").Enabled = Not mblnEdit
        mnuViewFind.Enabled = Not mblnEdit
        mnuViewRefresh.Enabled = Not mblnEdit
        txtEdit.Visible = mblnEdit And lng性质 = 2
        tab性质.Enabled = Not mblnEdit
    Case Else
        txtEdit.Visible = False
    End Select
    
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select 序号,名称,nvl(具有中心,0) as 具有中心 from 保险类别 where nvl(是否禁止,0)<>1 order by 序号"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm医保结算.Visible = True Then
        frm医保结算.Show
        Exit Sub
    End If
    
    Set mcol中心 = New Collection
    
    With cmb险类
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("名称")
            .ItemData(.NewIndex) = rsTemp("序号")
            mcol中心.Add Val(rsTemp("具有中心")), "K" & rsTemp("序号")
            If rsTemp("序号") = gintInsure Then
                '当前医保。
                '使用API，可以不激活Click事件
                zlControl.CboSetIndex .hwnd, .NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            '使用API，可以不激活Click事件
            zlControl.CboSetIndex .hwnd, 0
        End If
        
        mint险类 = .ItemData(.ListIndex)
        If mint险类 = TYPE_沈阳市 Then
            mnuFileBalance.Visible = True
            mnuBalance.Visible = True
        End If
        mnuPrintBalance.Visible = (mint险类 = TYPE_重庆银海版)
    End With
    
    
    frm医保结算.Show , frmParent
End Sub

Private Sub GetSpecialSQL(ByVal intTYPE As Integer, ByVal strBegin As String, ByVal strEnd As String)
    Select Case intTYPE
        Case 1 '-收费
            Select Case mint险类
            Case TYPE_重庆市
                gstrSQL = _
                    "Select A.结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,A.标识号 as 门诊号,Ltrim(A.姓名) as 姓名,F.名称 as 身份,A.性别,A.年龄,B.名称 as 开单科室," & _
                    "   A.操作员姓名 as 收费员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'999990.00')) as 进入统筹,Decode(C.超限自付金额,14,'特殊门诊',11,'急诊抢救','普通门诊') 门诊类别,C.备注 病种" & _
                    " From 病人费用记录 A,部门表 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F" & _
                    " Where A.记录性质 = 1 And A.操作员姓名 IS NOT NULL AND A.开单部门ID = B.ID(+) And A.登记时间>=" & strBegin & " and A.登记时间<=" & strEnd & _
                    "       and A.序号=1 and A.结帐ID=C.记录ID and C.性质=1 and C.险类=" & mint险类 & _
                    "       and A.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " ANd D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.结帐ID,A.NO,E.名称,D.卡号,A.病人ID,A.标识号,A.姓名,A.性别,A.年龄,B.名称,A.操作员姓名,A.登记时间,A.记录状态,F.名称,Decode(C.超限自付金额,14,'特殊门诊',11,'急诊抢救','普通门诊'),C.备注" & _
                    " Order by A.登记时间 Desc,A.NO Desc"
            Case TYPE_四川眉山
                gstrSQL = _
                    "Select C.记录ID,H.NO as 单据号,E.名称 as 中心,D.卡号,D.病人ID,'' as 门诊号,Ltrim(A.姓名) as 姓名,F.名称 as 身份,A.性别,A.年龄,H.开单科室 as 开单科室," & _
                    "   H.开单人 as 操作员,To_Char(C.结算时间,'YYYY-MM-DD HH24:MI:SS') as 结算时间,Decode(sign(C.发生费用金额),-1,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(Sum(C.进入统筹金额),'999990.00')) as 进入统筹,Ltrim(To_Char(Sum(C.超限自付金额),'999990.00')) as 超限自付金额,G.名称 病种,D.退休证号 职工属性" & _
                    " From 病人信息 A,(Select H.*,B.名称 开单科室 From 病人费用记录 H,部门表 B Where H.记录性质 = 1 And H.操作员姓名 IS NOT NULL And H.序号=1 AND H.开单部门ID = B.ID(+) ) H," & _
                    "      保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F,(Select * From 保险病种 Where 险类=" & mint险类 & ") G" & _
                    " Where H.结帐ID(+)=C.记录ID " & _
                    "       And A.病人ID=D.病人ID And C.病种ID=G.ID(+) And C.病人ID=D.病人ID And C.结算时间>=" & strBegin & " and C.结算时间<=" & strEnd & _
                    "       And C.性质=1 and C.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " and D.险类=C.险类 And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by C.记录ID,H.NO,E.名称,D.卡号,D.病人ID,A.姓名,A.性别,A.年龄,H.开单科室,H.开单人,C.结算时间,Decode(sign(C.发生费用金额),-1,'退','收'),F.名称,G.名称,D.退休证号" & _
                    " Order by C.结算时间 Desc"
            Case type_大连开发区, type_大连市
                '原过程参数:
                 '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
                 "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
                 '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
                 '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
                 '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
                 '过程新值代表为:
                 '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
                 '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
                 '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
                 '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,超限自付金额_IN,dbl个人帐户支付
                 '       支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
            
                gstrSQL = _
                    "Select A.结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,A.标识号 as 门诊号,Ltrim(A.姓名) as 姓名,F.名称 as 身份,A.性别,A.年龄,B.名称 as 开单科室," & _
                    "   A.操作员姓名 as 收费员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.起付线),'999990.00')) as 起付线, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 基本统筹支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 基本统筹自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'999990.00')) as 补充统筹支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.统筹报销金额),'999990.00')) as 补充统筹自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计进入统筹),'999990.00')) as 补助保险支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.大病自付金额),'999990.00')) as 非补助保险支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计统筹报销),'999990.00')) as 补助帐户支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.封顶线),'999990.00')) as 保险范围外自付," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.最高限额),'999990.00')) as 最高限额" & _
                    " From 病人费用记录 A,部门表 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F" & _
                    " Where A.记录性质 = 1 And A.操作员姓名 IS NOT NULL AND A.开单部门ID = B.ID(+) And A.登记时间>=" & strBegin & " and A.登记时间<=" & strEnd & _
                    "       and A.序号=1 and A.结帐ID=C.记录ID and C.性质=1 and C.险类=" & mint险类 & _
                    "       and A.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.结帐ID,A.NO,E.名称,D.卡号,A.病人ID,A.标识号,A.姓名,A.性别,A.年龄,B.名称,A.操作员姓名,A.登记时间,A.记录状态,F.名称" & _
                    " Order by A.登记时间 Desc,A.NO Desc"
                    
            Case Else
                gstrSQL = _
                    "Select A.结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,A.标识号 as 门诊号,Ltrim(A.姓名) as 姓名,F.名称 as 身份,A.性别,A.年龄,B.名称 as 开单科室," & _
                    "   A.操作员姓名 as 收费员,To_Char(A.登记时间,'YYYY-MM-DD HH24:MI:SS') as 收费时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'99999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'99999990.00')) as 进入统筹 " & _
                    " From 病人费用记录 A,部门表 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F" & _
                    " Where A.记录性质 = 1 And A.操作员姓名 IS NOT NULL AND A.开单部门ID = B.ID(+) And A.登记时间>=" & strBegin & " and A.登记时间<=" & strEnd & _
                    "       and A.序号=1 and A.结帐ID=C.记录ID and C.性质=1 and C.险类=" & mint险类 & _
                    "       and A.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.结帐ID,A.NO,E.名称,D.卡号,A.病人ID,A.标识号,A.姓名,A.性别,A.年龄,B.名称,A.操作员姓名,A.登记时间,A.记录状态,F.名称" & _
                    " Order by A.登记时间 Desc,A.NO Desc"
            End Select
        Case 2 '-结算（包括住院结算、特殊门诊结算）
            Select Case mint险类
            Case TYPE_重庆市
                gstrSQL = _
                    "Select A.ID as 结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,B.住院号,B.姓名,F.名称 as 身份,B.性别,B.年龄,'' as 科室," & _
                    "   A.操作员姓名 as 结帐人,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 结帐时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'9999999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'999990.00')) as 进入统筹," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.起付线),'999990.00')) as 起付线," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.实际起付线),'999990.00')) as 实际起付线," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.统筹报销金额),'999990.00')) as 统筹报销," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.超限自付金额),'999990.00')) as 超限金额,C.备注 病种" & _
                    " From 病人结帐记录 A,病人信息 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F " & _
                    " Where A.病人ID=C.病人ID And A.ID=C.记录ID  And A.收费时间>=" & strBegin & " and A.收费时间<=" & strEnd & _
                    "       and C.性质=2  and C.病人ID=B.病人ID and B.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.ID,A.NO,E.名称,D.卡号,A.病人ID,B.住院号,B.姓名,B.性别,B.年龄,A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),A.记录状态,F.名称,C.备注" & _
                    " Order by To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc"
            Case TYPE_四川眉山
                gstrSQL = _
                    "Select C.记录ID,'' as 单据号,E.名称 as 中心,D.卡号,D.病人ID,'' as 门诊号,Ltrim(A.姓名) as 姓名,F.名称 as 身份,A.性别,A.年龄,'' as 开单科室," & _
                    "   C.经办人 as 操作员,To_Char(C.结算时间,'YYYY-MM-DD HH24:MI:SS') as 结算时间,Decode(sign(C.发生费用金额),-1,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(Sum(C.进入统筹金额),'999990.00')) as 进入统筹,Ltrim(To_Char(Sum(C.超限自付金额),'999990.00')) as 超限自付金额,Ltrim(To_Char(Sum(C.统筹报销金额),'999990.00')) 统筹报销金额,D.退休证号 职工属性" & _
                    " From 病人信息 A,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F,(Select * From 保险病种 Where 险类=" & mint险类 & ") G" & _
                    " Where A.病人ID=D.病人ID And C.病种ID=G.ID(+) And C.病人ID=D.病人ID And C.结算时间>=" & strBegin & " and C.结算时间<=" & strEnd & _
                    "       and C.性质=2 and C.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " and D.险类=C.险类 And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by C.记录ID,E.名称,D.卡号,D.病人ID,A.姓名,A.性别,A.年龄,C.经办人,C.结算时间,Decode(sign(C.发生费用金额),-1,'退','收'),F.名称,D.退休证号" & _
                    " Order by C.结算时间 Desc"
            Case type_大连开发区, type_大连市
                '原过程参数:
                 '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
                 "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
                 '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
                 '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
                 '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
                 '过程新值代表为:
                 '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
                 '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
                 '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
                 '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,超限自付金额_IN,dbl个人帐户支付
                 '       支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
                 '说明:需要再增加一列为“统筹总额”，最好放在“最高限额”前面，其公式为“统筹总额=基本统筹支付+补充统筹支付”，因为最高限额需要对照此项金额输入；
                 
                gstrSQL = _
                            "Select A.ID as 结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,D.医保号,B.住院号,C.主页id as 住院次数,B.姓名,F.名称 as 身份,B.性别,B.年龄,L.名称 as 科室," & _
                            "    Ltrim(to_Char(sum(nvl(c.全自付金额,0)+nvl(c.进入统筹金额,0)),'999999999999999990.99')) as 统筹总额,to_char(max(C.最高限额),'999999999999999990.99') 最高限额," & _
                            "   A.操作员姓名 as 结帐人,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 结帐时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.起付线),'999990.00')) as 起付线, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 基本统筹支付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 基本统筹自付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'999990.00')) as 补充统筹支付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.统筹报销金额),'999990.00')) as 补充统筹自付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计进入统筹),'999990.00')) as 补助保险支付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.大病自付金额),'999990.00')) as 非补助保险支付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计统筹报销),'999990.00')) as 补助帐户支付, " & _
                            "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.封顶线),'999990.00')) as 保险范围外自付" & _
                            " From 病人结帐记录 A,病人信息 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F,病案主页 Q,部门表 L" & _
                            " Where A.病人ID=C.病人ID And A.ID=C.记录ID    And A.收费时间>=" & strBegin & " and A.收费时间<=" & strEnd & _
                            "       and b.病人id=Q.病人id and nvl(C.主页id,0)=nvl(Q.主页id,0)  and Q.出院科室id =L.ID(+)  " & _
                            "       and C.性质=2  and C.病人ID=B.病人ID and B.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                            " Group by A.ID,A.NO,E.名称,D.卡号,A.病人ID,D.医保号,B.住院号,c.主页id,L.名称,B.姓名,B.性别,B.年龄,A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),A.记录状态,F.名称" & _
                            " Order by 科室,医保号"
                            
'                gstrSQL = _
                    "Select A.ID as 结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,D.医保号,B.住院号,B.住院次数,B.姓名,F.名称 as 身份,B.性别,B.年龄,max(L.名称) as 科室," & _
                    "    Ltrim(to_Char(sum(nvl(c.全自付金额,0)+nvl(c.进入统筹金额,0)),'999999999999999990.99')) as 统筹总额,to_char(max(C.最高限额),'999999999999999990.99') 最高限额," & _
                    "   A.操作员姓名 as 结帐人,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 结帐时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.起付线),'999990.00')) as 起付线, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 基本统筹支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 基本统筹自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'999990.00')) as 补充统筹支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.统筹报销金额),'999990.00')) as 补充统筹自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计进入统筹),'999990.00')) as 补助保险支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.大病自付金额),'999990.00')) as 非补助保险支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.累计统筹报销),'999990.00')) as 补助帐户支付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.封顶线),'999990.00')) as 保险范围外自付" & _
                    " From 病人结帐记录 A,病人信息 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F,病案主页 Q,部门表 L" & _
                    " Where A.病人ID=C.病人ID And A.ID=C.记录ID    And A.收费时间>=" & strBegin & " and A.收费时间<=" & strEnd & _
                    "       and b.病人id=Q.病人id(+) and b.住院次数=Q.主页id(+)  and Q.出院科室id =L.ID(+)  " & _
                    "       and C.性质=2  and C.病人ID=B.病人ID and B.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.ID,A.NO,E.名称,D.卡号,A.病人ID,D.医保号,B.住院号,B.住院次数,B.姓名,B.性别,B.年龄,A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),A.记录状态,F.名称" & _
                    " Order by 科室,医保号"
                    
                    saveFlexState msh记录_S, "费用结算_大连"
                    ',To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc
            Case Else
                gstrSQL = _
                    "Select A.ID as 结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,B.住院号,B.姓名,F.名称 as 身份,B.性别,B.年龄,'' as 科室," & _
                    "   A.操作员姓名 as 结帐人,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') as 结帐时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.个人帐户支付),'9999999990.00')) as 个人帐户," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.发生费用金额),'99999990.00')) as 发生费用, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.全自付金额),'999990.00')) as 全自费, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.首先自付金额),'999990.00')) as 首先自付, " & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.进入统筹金额),'99999990.00')) as 进入统筹," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.起付线),'999990.00')) as 起付线," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.实际起付线),'999990.00')) as 实际起付线," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.统筹报销金额),'999990.00')) as 统筹报销," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(C.超限自付金额),'999990.00')) as 超限金额" & _
                    " From 病人结帐记录 A,病人信息 B,保险结算记录 C,保险帐户 D,保险中心目录 E,保险人群 F " & _
                    " Where A.病人ID=C.病人ID And A.ID=C.记录ID  And A.收费时间>=" & strBegin & " and A.收费时间<=" & strEnd & _
                    "       and C.性质=2  and C.病人ID=B.病人ID and B.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.ID,A.NO,E.名称,D.卡号,A.病人ID,B.住院号,B.姓名,B.性别,B.年龄,A.操作员姓名,To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS'),A.记录状态,F.名称" & _
                    " Order by To_Char(A.收费时间,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc"
            End Select
        Case 3 '-预交
            Select Case mint险类
            Case Else
                gstrSQL = _
                    "Select A.ID as 结帐ID,A.NO as 单据号,E.名称 as 中心,D.卡号,A.病人ID,B.住院号,B.姓名,F.名称 as 身份,B.性别,B.年龄,C.名称 as 科室," & _
                    "   A.操作员姓名 as 收款人,To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS') as 收款时间,Decode(A.记录状态,2,'退','收') as 收退标志," & _
                    "   Ltrim(To_Char(decode(A.记录状态,2,-1,1)*Sum(A.金额),'9999999990.00')) as 个人帐户" & _
                    "   From 病人预交记录 A,病人信息 B,部门表 C,保险帐户 D,保险中心目录 E,保险人群 F" & _
                    " Where A.记录性质=1 And A.病人ID=B.病人ID And A.科室ID=C.ID(+) " & _
                    "       and A.结算方式='个人帐户' and A.收款时间>=" & strBegin & " and A.收款时间<=" & strEnd & _
                    "       and B.病人ID=D.病人ID and D.险类=" & mint险类 & IIf(mstrCardCond = "", "", " And D.医保号='" & mstrCardCond & "'") & " And D.险类=E.险类 and D.中心=E.序号 and D.险类=F.险类 and D.在职=F.序号 " & _
                    " Group by A.ID,A.NO,E.名称,D.卡号,A.病人ID,B.住院号,B.姓名,B.性别,B.年龄,C.名称," & _
                    "     A.操作员姓名,To_Char(A.收款时间,'YYYY-MM-DD HH24:MI:SS'),A.记录状态,F.名称" & _
                    " Order by 收款时间 Desc,单据号 Desc"
            End Select
    End Select
End Sub
Private Function Save保存最高限额(ByVal int险类 As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对大连有用,保存最高限额
    '--入参数:
    '--出参数:
    '--返  回:
    '--修改人:刘兴宏;20040630
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lng记录id As Long
    Dim lng病人ID As Long
    Dim dbl限额 As Double
    Dim int性质 As Integer
    Dim strSQL As String
    Dim lngPross  As Long
    Dim lngprossCount     As Long
    On Error GoTo ErrHand:
    
    Save保存最高限额 = False

    gcnOracle.BeginTrans
    With msh记录_S
        lngPross = 1
        lngprossCount = .Rows - 1
        For lngRow = 1 To .Rows - 1
            lng记录id = Val(.TextMatrix(lngRow, col记录ID))
            lng病人ID = Val(.TextMatrix(lngRow, col病人ID))
            If lng记录id <> 0 And lng病人ID <> 0 And .RowData(lngRow) = 1 Then
                int性质 = Val(Mid(tab性质.SelectedItem.Key, 2))
                dbl限额 = Val(.TextMatrix(lngRow, mintCol最高限额))
                strSQL = "zl_保险结算记录限额_Update(" & _
                             int性质 & "," & _
                            lng记录id & "," & _
                             dbl限额 & ")"
                gcnOracle.Execute strSQL
            End If
            Call ShowPercent(lngPross / lngprossCount, "正在保存限额")
            lngPross = lngPross + 1
        Next
    End With
    gcnOracle.CommitTrans
    Save保存最高限额 = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
End Function
Private Sub MoveEditCtl()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:移动编辑控件
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Err = 0
    On Error Resume Next
    If Not mblnEdit Then Exit Sub
    mblnNOScroll = True
    With msh记录_S
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        .LeftCol = col单据号
        If Not .ColIsVisible(mintCol最高限额) Then
            .LeftCol = mintCol最高限额
        End If
        .Col = mintCol最高限额
        txtEdit.Left = .Left + .CellLeft - 15
        txtEdit.Top = .Top + .CellTop - 15
        txtEdit.Height = .RowHeight(.Row) - 15
        txtEdit.Width = .CellWidth - 20
        txtEdit.Text = Format(Val(.TextMatrix(.Row, mintCol最高限额)), "####0.00;####0.00; ;")
        .Col = 0
        .ColSel = .Cols - 1
    End With
    txtEdit.Visible = mblnEdit
    If txtEdit.Visible Then
        txtEdit.SetFocus
    End If
    mblnNOScroll = False
End Sub

Private Sub txtEdit_Change()
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCol As Integer
    Dim intNextCol As Integer
    Dim intRow As Integer
    
    Select Case KeyCode
    Case vbKeyReturn         '按下回车
        With msh记录_S
            If Val(.TextMatrix(.Row, mintCol最高限额)) <> Val(txtEdit.Text) Then
                .RowData(.Row) = 1
                .TextMatrix(.Row, mintCol最高限额) = Format(Val(txtEdit.Text), "####0.00;-###0.00; ;")
            End If
            If .Rows - 1 = .Row Then    '是尾行,则返回第一行
                .Row = 1
            Else
                .Row = .Row + 1
            End If
            '设置文本
            MoveEditCtl
            KeyCode = 0
            zlControl.TxtSelAll txtEdit
        End With
    Case vbKeyDown      '下箭头
        With msh记录_S
            If Val(.TextMatrix(.Row, mintCol最高限额)) <> Val(txtEdit.Text) Then
                .RowData(.Row) = 1
                .TextMatrix(.Row, mintCol最高限额) = Format(Val(txtEdit.Text), "####0.00;-###0.00; ;")
            End If
            If .Rows - 1 = .Row Then    '是尾行,则返回第一行
                .Row = 1
            Else
                .Row = .Row + 1
            End If
        End With
        '设置文本
        MoveEditCtl
        KeyCode = 0
        zlControl.TxtSelAll txtEdit
    Case vbKeyUp                '上箭头
        With msh记录_S
            If Val(.TextMatrix(.Row, mintCol最高限额)) <> Val(txtEdit.Text) Then
                .RowData(.Row) = 1
                .TextMatrix(.Row, mintCol最高限额) = Format(Val(txtEdit.Text), "####0.00;-###0.00; ;")
            End If
            If .Row <= 1 Then    '是尾行,则返回第一行
                .Row = .Rows - 1
            Else
                .Row = .Row - 1
            End If
        End With
        '设置文本
        MoveEditCtl
        KeyCode = 0
        zlControl.TxtSelAll txtEdit
    Case vbKeyLeft              '左箭头
    Case vbKeyRight             '右简箭头
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m金额式
    mblnChange = True
    SetMenu
End Sub
Private Sub ShowPercent(sngPercent As Single, Optional strCaption As String = "")
    '功能:在状态条上根据百分比显示当前处理进度(█)
    Dim intAll As Integer
    If strCaption = "" Then
        intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
        stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
    Else
        intAll = stbThis.Panels(2).Width / TextWidth("█") - zlCommFun.ActualLen(strCaption) - 2
        stbThis.Panels(2).Text = strCaption & "  " & Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
    End If
End Sub
Private Sub SetCOLAlign_大连()
    '只设置列对剂
    Dim i As Long
    With msh记录_S
        .ColWidth(col记录ID) = 0
        .ColWidth(col病人ID) = 0
        
        i = 5: .ColAlignment(i) = 4
        i = i + 1:  .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 1
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 4
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
        i = i + 1: .ColAlignment(i) = 7
    End With
    RestoreFlexState msh记录_S, "费用结算_大连"
End Sub
