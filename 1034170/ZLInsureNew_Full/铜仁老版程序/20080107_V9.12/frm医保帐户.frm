VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保帐户 
   Caption         =   "医保帐户管理"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "frm医保帐户.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   7050
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2850
      ScaleWidth      =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1890
      Width           =   45
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   8685
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":06EA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":0904
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":0B1E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":0D38
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":0F52
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":164C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":1866
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":1A80
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   9285
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":1C9A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":1EB4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":20CE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":22E8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":2502
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":2BFC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":2E16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保帐户.frx":3030
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6390
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保帐户.frx":324A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12541
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   1270
      BandCount       =   2
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   9975
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      BandForeColor1  =   -2147483635
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   660
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
      Begin VB.ComboBox cmb险类 
         Height          =   300
         Left            =   7890
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   1995
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
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
               Caption         =   "验证"
               Key             =   "Modify"
               Object.ToolTipText     =   "身份验证"
               Object.Tag             =   "验证"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitModify"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "身份"
               Key             =   "Custom"
               Description     =   "Custom"
               Object.ToolTipText     =   "自定义病人身份"
               Object.Tag             =   "身份"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "查找医保帐户"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh帐户_S 
      Height          =   5655
      Left            =   15
      TabIndex        =   3
      Top             =   735
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9975
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
      Appearance      =   0
      MouseIcon       =   "frm医保帐户.frx":3ADC
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   5625
      Left            =   7200
      ScaleHeight     =   5595
      ScaleWidth      =   2715
      TabIndex        =   5
      Top             =   750
      Width           =   2745
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh身份信息 
         Height          =   1365
         Left            =   -30
         TabIndex        =   9
         Top             =   4260
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2408
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
         Appearance      =   0
         MouseIcon       =   "frm医保帐户.frx":3DF6
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh年度 
         Height          =   3405
         Left            =   -30
         TabIndex        =   10
         Top             =   450
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6006
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         MouseIcon       =   "frm医保帐户.frx":4110
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Lbl其它身份信息 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H009B6737&
         Caption         =   "其它身份信息："
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label Lbl年度情况 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H009B6737&
         Caption         =   "年度情况："
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   135
         Width           =   900
      End
      Begin VB.Label lbl年度 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2002"
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCard 
         Caption         =   "卡片打印(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSplit2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditModify 
         Caption         =   "身份验证(&I)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除帐户(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPassword 
         Caption         =   "修改密码(&M)"
      End
      Begin VB.Menu mnuEditSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisease 
         Caption         =   "病种选择(&D)"
      End
      Begin VB.Menu mnuEditQuery 
         Caption         =   "查询单位欠费(&Q)"
      End
      Begin VB.Menu mnuEditSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "补充登记(&S)"
      End
      Begin VB.Menu mnuEditRollIncome 
         Caption         =   "撤消入院登记(&R)"
      End
      Begin VB.Menu mnuEditRollAdmit 
         Caption         =   "撤消急诊登记(&R)"
      End
      Begin VB.Menu mnuEditOut 
         Caption         =   "补充出院登记(&O)"
      End
      Begin VB.Menu mnuEditSplit3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_UpDetail 
         Caption         =   "补充上传门诊明细(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Account 
         Caption         =   "核对帐户支付信息(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Hospital 
         Caption         =   "核对入出院信息(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_ZYPrice 
         Caption         =   "核对住院结算信息(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Price 
         Caption         =   "核对费用结算信息(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditVerify_Detail 
         Caption         =   "核对费用明细信息(&D)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSp 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditXE 
         Caption         =   "限额编辑(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSplit4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditMend 
         Caption         =   "补卡(&E)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditLoss 
         Caption         =   "禁止或开启(&L)"
         Visible         =   0   'False
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCustom 
         Caption         =   "自定义身份信息(&A)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
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
Attribute VB_Name = "frm医保帐户"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 行Enum
    row住院次数 = 1
    row帐户余额 = 2
    row帐户增加 = 3
    row帐户支出 = 4
    row本次起付线 = 5
    row起付线累计 = 6
    row统筹限额 = 7
    row进入统筹 = 8
    row统筹报销 = 9
    row大额限额 = 10
    row大额累计 = 11
    row封锁信息 = 12
End Enum

Private Enum 列Enum
    col中心 = 0
    col卡号 = 1
    col医保号 = 2
    col病人ID = 3
    col姓名 = 4
    col性别 = 5
    col出生日期 = 6
    col身份证号 = 7
    col人员身份 = 8
    col身份编码 = 9
    col单位编码 = 10
    col退休证号 = 11
    col病种 = 12
    col状态 = 13
    col帐户余额 = 14
    col就诊时间 = 15
End Enum

Private mblnLoad As Boolean                     '第一次启动
Private mstr补充字段 As String                  '用户设置的字段
Private mstrFind As String                      '查找条件

Private mrs帐户 As New ADODB.Recordset
Private mint险类 As Integer
Private mcol中心 As New Collection              '保存医保类别具有中心特性
Private mcol可用 As New Collection              '保存该医保是否可以初始化
Private msngStartX As Single
Private strServer As String, strUser As String, strPass As String
Private mcnYB As New ADODB.Connection   '医保前置服务器连接
Private mrs病种 As New ADODB.Recordset

Private Sub cmb险类_Click()
    Dim blnCanUse As Boolean
    
    With cmb险类
        If mint险类 = .ItemData(.ListIndex) Then Exit Sub
        mint险类 = .ItemData(.ListIndex)
    End With
    mnuEditPassword.Enabled = (mint险类 = gintInsure)
    mnuEditModify.Enabled = (mint险类 = gintInsure)
    
    mnuEditXE.Visible = False ' mint险类 = TYPE_大连市 Or mint险类 = TYPE_大连开发区
    mnuEditSp.Visible = mnuEditXE.Visible
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    
    blnCanUse = GetInsureInit(mint险类)
    mnuEditSub.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditDisease.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditRollIncome.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditRollAdmit.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditQuery.Enabled = (mint险类 = gintInsure) And blnCanUse
    
    Call FillList
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        lbl年度.Caption = Format(zlDatabase.Currentdate, "yyyy")
        mstrFind = " and A.就诊时间>=sysdate-30"
        If mint险类 = TYPE_四川眉山 Then mstrFind = " And Nvl(A.灰度级,0)<>9"
        
        '显示帐户
        Call FillList
        Call GetAccountInfo
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    mblnLoad = True
    '取注册表
    mstr补充字段 = Replace(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "补充字段", ""), "'", "")
    
    zlControl.CboSetHeight cmb险类, 3600
    Call InitTable
    RestoreWinState Me, App.ProductName
    Call 权限控制
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbr.Height)
    Call GetAccountInfo
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    picSplitV.Left = Me.ScaleWidth - 3000
    With msh帐户_S
        .Top = IIf(cbr.Visible, cbrHeight, 0)
        .Width = picSplitV.Left - 25
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    With picSplitV
        .Top = msh帐户_S.Top
        .Height = msh帐户_S.Height
    End With
    With picOther
        .Top = msh帐户_S.Top
        .Left = picSplitV.Left + picSplitV.Width
        .Height = msh帐户_S.Height
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub mnuEditDelete_Click()
    Dim lng病人id As Long
    Dim blnDelete As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '删除帐户信息(灰度级:0-正常;1-挂失;2-禁止个帐;3-禁止统筹;9-帐户已撤销)
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    If lng病人id <= 0 Then
        MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("你真的要删除该保险帐户吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    On Error Resume Next
    Err = 0
    gcnOracle.BeginTrans
    
'    gstrSQL = "Select count(*) Records From 保险结算记录 Where 病人ID=" & lng病人ID
'    Call OpenRecordset(rsTemp, Me.Caption)
'    blnDelete = (rsTemp.RecordCount = 0)
'
'    If blnDelete Then
'        gstrSQL = "Delete 保险帐户 Where 险类=" & gintInsure & " And 病人ID=" & lng病人ID
'        gcnOracle.Execute gstrSQL
'    Else
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_四川眉山 & ",'灰度级','9')"
        gcnOracle.Execute gstrSQL
'    End If
    gcnOracle.CommitTrans
    
    Call FillList
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditDisease_Click()
    Dim lng病人id As Long, lng主页ID As Long, rs病种 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng病种 As Long, str病种 As String
    
    On Error GoTo errHandle
    
    With msh帐户_S
        '病人直接从列表中取得
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '判断该病人是否有住院记录
        gstrSQL = "select A.姓名,B.主页ID,B.入院日期,B.出院日期 " & _
                  "  from 病人信息 A,病案主页 B " & _
                  "  Where A.病人ID = " & lng病人id & " And A.病人ID = B.病人ID And B.险类 = " & mint险类 & _
                  "  Order by B.入院日期 Desc"
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '无法从记录集中取得病人姓名
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 无法找到有效的住院记录。可能未住院或未以医保身份入院。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If IsNull(rsTemp("出院日期")) = False Then
            If MsgBox("病人 " & rsTemp("姓名") & " 已于" & Format(rsTemp("出院日期"), "yyyy-MM-dd") & "出院，是否还需要更新疾病信息？", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
        
        lng主页ID = rsTemp!主页ID
        If mint险类 = TYPE_重庆市 Then
            Call 更新出院疾病_重庆(lng病人id, lng主页ID)
        ElseIf mint险类 = TYPE_重庆银海版 Then
            Call 更新疾病_重庆银海版(Me, lng病人id, lng主页ID)
        ElseIf mint险类 = TYPE_自贡市 Then
            '提取已选择的病种
            Dim rsSelected As New ADODB.Recordset
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A,zlyb.病种信息 B where A.险类=" & TYPE_自贡市 & " And B.病人ID=" & lng病人id & " And A.ID=B.病种ID And A.险类=B.险类"
            Call OpenRecordset(rsSelected, "提取上次已选择的病种")
            
            '住院要选择病种，以确认一些特殊收费项目
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.险类=" & TYPE_自贡市
            Call OpenRecordset(rs病种, "身份验证")
            If rs病种.RecordCount > 0 Then
                If frm多病种选择.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：", rsSelected, False) = True Then
                    lng病种 = 0
                    str病种 = ""
                    With rs病种
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            lng病种 = rs病种("ID")
                        End If
                        Do While Not .EOF
                            str病种 = str病种 & "|" & rs病种!ID
                            .MoveNext
                        Loop
                        If str病种 <> "" Then str病种 = Mid(str病种, 2)
                        
                        gstrSQL = "zlyb.zl_病种信息_INSERT(" & TYPE_自贡市 & "," & lng病人id & ",'" & str病种 & "')"
                        Call ExecuteProcedure("更新病种")
                    End With
                    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_自贡市 & ",'病种ID','" & lng病种 & "')"
                    Call ExecuteProcedure("更新病种")
                End If
            End If
        ElseIf mint险类 = TYPE_泸州市 Then
            '住院要选择病种，以确认一些特殊收费项目
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.险类=" & TYPE_泸州市
            Call OpenRecordset(rs病种, "身份验证")
            If rs病种.RecordCount > 0 Then
                If frmListSel.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
                    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_泸州市 & ",'病种ID','" & rs病种("ID") & "')"
                    Call ExecuteProcedure("更新病种")
                End If
            End If
        ElseIf mint险类 = TYPE_沈阳市 Then
            '用于修改病种，可能入院登记时病种选择错误或其它情况
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码 " & _
                    " From 保险病种 A where A.险类=" & TYPE_沈阳市
            Call OpenRecordset(rs病种, "身份验证")
            If rs病种.RecordCount > 0 Then
                If frmListSel.ShowSelect(rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
                    str病种 = rs病种!编码
                    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_沈阳市 & ",'病种ID','" & rs病种("ID") & "')"
                    Call ExecuteProcedure("更新病种")
                Else
                    Exit Sub
                End If
            End If
            
            If Not 医保初始化_沈阳市 Then Exit Sub
            Call 更新病种_沈阳市(lng病人id, lng主页ID, str病种)
        End If
        Call FillList
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditLoss_Click()
    Dim int状态 As Integer
    Dim lng病人id As Long
    Dim strMsg As String
    Dim rsAccount As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '封锁该帐户(灰度级:0-正常;1-禁止个人帐户;9-帐户已撤销)
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    If lng病人id <= 0 Then
        MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查卡的状态，如果已经封锁，则表示要解锁；否则将封锁该卡
    gstrSQL = "Select Nvl(灰度级,0) 状态 From 保险帐户 Where 病人ID=" & lng病人id & " And 险类=" & TYPE_四川眉山
    Call OpenRecordset(rsAccount, Me.Caption)
    If rsAccount!状态 = 0 Then
        int状态 = 1
        strMsg = "封锁该卡吗？（封锁后将不能使用）"
    Else
        int状态 = 0
        strMsg = "恢复该卡的状态为正常吗？"
    End If
    If MsgBox("你确定要" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '请输入封锁说明！
    strMsg = ""
    If int状态 = 1 Then
        Do While True
            strMsg = InputBox("请输入封锁卡的信息：", "封锁医保卡的使用")
            If Trim(strMsg) <> "" Then Exit Do
        Loop
    End If
    
    '更新帐户信息
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_四川眉山 & ",'灰度级','" & int状态 & " ')"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    If int状态 = 1 Then
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_四川眉山 & ",'备注','''" & strMsg & " ''')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Else
        gstrSQL = "ZL_保险帐户_更新信息(" & lng病人id & "," & TYPE_四川眉山 & ",'备注','NULL')"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
    End If
    
    Call FillList
    Call msh帐户_S_EnterCell
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditMend_Click()
    '补卡
    Dim strIdentify As String
    Dim bytType As Byte
    Dim lng病人id As Long
    Dim cls医保 As New clsInsure
    
    On Error GoTo ErrHand
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    If lng病人id = 0 Then
        MsgBox "请选择一位医保病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    bytType = 4
    strIdentify = cls医保.Identify(bytType, lng病人id)
    If strIdentify <> "" Then
        Call FillList
    End If
    
    Call msh帐户_S_EnterCell
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditOut_Click()
    Dim lng病人id As Long, lng主页ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    
    'Modified by 朱玉宝 20031218 地区：福州
    If Not (mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or _
    mint险类 = TYPE_南平市 Or mint险类 = TYPE_昆明市 Or mint险类 = TYPE_云南省 Or _
    mint险类 = TYPE_沈阳市 Or mint险类 = TYPE_重庆银海版) Then Exit Sub
    If lng病人id = 0 Then
        MsgBox "请选择一位医保病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    If MsgBox("你确定要为该病人补办出院手续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '取得主页ID
    gstrSQL = "Select Nvl(住院次数,0) 主页ID From 病人信息 Where 病人ID=" & lng病人id
    Call OpenRecordset(rsTemp, "取主页ID")
    lng主页ID = rsTemp!主页ID
    '存在未结费用时，不允许办理出院手续
    If 存在未结费用(lng病人id, lng主页ID) Then
        MsgBox "该医保病人还存在未结费用，不允许办理出院手续！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '调用医保出院接口
    Select Case mint险类
    Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
        If Not frm等待响应.ShowMe(操作方式.出院, 请求目的.刷卡) Then Exit Sub
        If lng病人id <> 获取病人ID Then
            MsgBox "病人信息不符！", vbInformation, gstrSysName
            Exit Sub
        End If
        If Not frm等待响应.ShowMe(操作方式.出院, 请求目的.申请, lng病人id) Then Exit Sub
    
        '出院登记
        gstrSQL = "zl_保险帐户_出院(" & lng病人id & "," & mint险类 & ")"
        Call ExecuteProcedure("福建巨龙")
        MsgBox "该医保病人成功补办出院手续！", vbInformation, gstrSysName
    Case TYPE_昆明市, TYPE_云南省
        gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
                " C.卡号,C.医保号,C.密码,C.顺序号 " & _
                " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
                " Where A.病人ID=D.病人ID And A.病人ID=" & lng病人id & " And A.主页ID=" & lng主页ID & _
                " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=" & mint险类
        Call OpenRecordset(rsTemp, "取顺序号")
    
        If rsTemp.EOF Then
            MsgBox "没有此病人或此病人不是医保病人，无法办理出院手续！", vbExclamation, gstrSysName
            Exit Sub
        End If
        If IsNull(rsTemp!顺序号) Then
            MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not 出院登记_云南(lng病人id, lng主页ID, rsTemp!顺序号) Then Exit Sub
        MsgBox "该医保病人成功补办出院手续！", vbInformation, gstrSysName
    Case TYPE_沈阳市
        Call 出院登记_沈阳市(lng病人id, lng主页ID)
    Case TYPE_重庆银海版
        Call 出院登记_重庆银海版(lng病人id, lng主页ID)
    End Select
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuEditPassword_Click()
    Dim str卡号 As String, str医保号 As String, str分中心 As String, str密码 As String
    Dim lng病种ID As Long
    
    Select Case mint险类
        Case TYPE_自贡市
            Call frmIdentify中软.GetPatient(False, True)
        Case TYPE_泸州市
            Call frmIdentify泸州.GetPatient(2, True, lng病种ID)
        Case TYPE_成都郊县
            Call frmIdentify贵阳.GetIdentify(TYPE_成都郊县, str卡号, str医保号, str分中心, str密码, True, True)
    End Select
End Sub

Private Sub mnuEditQuery_Click()
    Dim lng病人id  As Long, rsTemp As New ADODB.Recordset
    
    If gintInsure <> TYPE_贵阳市 Then Exit Sub
    
    With msh帐户_S
        '病人直接从列表中取得
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '判断该病人是否有住院记录
        gstrSQL = "select A.单位编码 " & _
                  "  from 保险帐户 A " & _
                  "  Where A.病人ID = " & lng病人id & " And A.险类 = " & mint险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '无法从记录集中取得病人姓名
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 无法找到有效的登记信息。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call 查询欠费单位_贵阳(NVL(rsTemp("单位编码"), ""))
    End With
    
End Sub

Private Sub mnuEditRollAdmit_Click()
    Dim lng病人id As Long
    Dim rsTemp As New ADODB.Recordset
    
    With msh帐户_S
        '病人直接从列表中取得
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你是否要将病人“" & .TextMatrix(.Row, col姓名) & "”的急诊登记撤消？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '判断该病人是否有住院记录
        gstrSQL = "select A.顺序号,B.编码 " & _
                  "  from 保险帐户 A,保险病种 B " & _
                  "  Where A.病人ID = " & lng病人id & " And A.病种ID = B.ID And B.险类 = " & mint险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '无法从记录集中取得病人姓名
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 无法找到有效的登记信息，可能未以急诊病人登记。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If IsNull(rsTemp("顺序号")) = True Then
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 登记信息不完整，可能己做了其它登记。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If rsTemp("编码") = "0090" Then
            If 撤消急诊登记_云南(rsTemp("顺序号")) = True Then
                MsgBox "撤消成功，可以再次进行急诊登记。", vbInformation, gstrSysName
            End If
        End If
    End With

End Sub

Private Sub mnuEditRollIncome_Click()
    Dim lng病人id As Long, lng主页ID  As Long
    Dim rsTemp As New ADODB.Recordset
    
    With msh帐户_S
        '病人直接从列表中取得
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你是否要将病人“" & .TextMatrix(.Row, col姓名) & "”从医保病人转为普通病人？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        '获得主页ID
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & lng病人id
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以从医保病人转为普通病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        lng主页ID = rsTemp("主页ID")
        
        '判断该病人是否有住院记录
        gstrSQL = "select A.顺序号 " & _
                  "  from 保险帐户 A " & _
                  "  Where A.病人ID = " & lng病人id & " And A.险类 = " & mint险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            '无法从记录集中取得病人姓名
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 无法找到有效的登记信息，可能未以急诊病人登记。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If IsNull(rsTemp("顺序号")) = True Then
            MsgBox "病人 " & msh帐户_S.TextMatrix(.Row, col姓名) & " 登记信息不完整，可能己做了其它登记。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint险类 = TYPE_重庆市 Then
            If 撤消医保入院_重庆(lng病人id, lng主页ID, rsTemp("顺序号")) = True Then
                MsgBox "撤消成功，该病人已经由医保病人转为普通病人。", vbInformation, gstrSysName
            End If
        ElseIf mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
            If Not frm等待响应.ShowMe(操作方式.入院, 请求目的.冲销, lng病人id) Then Exit Sub
            gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人id & "," & lng主页ID & ")"
            Call ExecuteProcedure("撤消医保入院")
        End If
    End With

End Sub

Private Sub mnuEditVerify_Account_Click()
    Dim lng病人id As Long
    With msh帐户_S
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint险类 = TYPE_成都郊县 Then
            Call 核对帐户支付_成都效县(lng病人id)
        ElseIf mint险类 = type_米易 Then
            Call 核对帐户支付_米易
        End If
    End With
End Sub

Private Sub mnuEditVerify_Hospital_Click()
    Dim lng病人id As Long
    With msh帐户_S
        lng病人id = Val(.TextMatrix(.Row, col病人ID))
        If lng病人id <= 0 Then
            MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mint险类 = TYPE_成都郊县 Then
            Call 核对入出院_成都效县(lng病人id)
        End If
    End With
End Sub

Private Sub mnuEditVerify_Price_Click()
    Dim lng病人id As Long
    With msh帐户_S
        If mint险类 <> TYPE_重庆银海版 Then
            lng病人id = Val(.TextMatrix(.Row, col病人ID))
            If lng病人id <= 0 Then
                MsgBox "请选择一位医保病人。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '调试重庆医保银海版 204-03-31
        If mint险类 = TYPE_成都郊县 Then
            Call 核对费用结算_成都效县(lng病人id)
        ElseIf mint险类 = TYPE_重庆银海版 Then
            Call 核对费用结算_重庆银海版
        End If
    End With
End Sub

Private Sub mnuEditVerify_UpDetail_Click()
    If mint险类 = type_米易 Then
        Call 补充上传门诊明细
    End If
End Sub

Private Sub mnuEditVerify_ZYPrice_Click()
    If mint险类 = type_米易 Then
        Call 核对住院结算_米易
    End If
End Sub

Private Sub mnuEditXE_Click()
    '主要录入大连的最高限额
    Dim lng病人id As Long
    Dim strIdentify As String
    Dim bytType As Byte
    Dim cls医保 As New clsInsure
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    If lng病人id = 0 Then Exit Sub
    bytType = 9
    strIdentify = cls医保.Identify(bytType, lng病人id)
    If strIdentify <> "" Then
        Call FillList
    End If
    
End Sub

Private Sub mnuFileCard_Click()
    Dim str医保号 As String
    '打印卡片
    str医保号 = Trim(msh帐户_S.TextMatrix(msh帐户_S.Row, col医保号))
    If str医保号 = "" Then Exit Sub
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1604", Me, "险类=" & mint险类, "医保号=" & str医保号, 2)
End Sub

Private Sub msh年度_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    msh年度.ToolTipText = msh年度.TextMatrix(msh年度.MouseRow, msh年度.MouseCol)
End Sub

Private Sub msh帐户_S_Scroll()
    Call GetAccountInfo
End Sub

Private Sub picOther_Resize()
    msh年度.Left = 0
    msh年度.Width = picOther.ScaleWidth
    
    msh身份信息.Left = 0
    msh身份信息.Width = picOther.ScaleWidth
    If picOther.ScaleHeight - msh身份信息.Top > 0 Then
        msh身份信息.Height = picOther.ScaleHeight - msh身份信息.Top
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
        If sngTemp > msh帐户_S.Left + 2000 And ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            msh帐户_S.Width = picSplitV.Left - msh帐户_S.Left
            picOther.Left = sngTemp + picSplitV.Width
            picOther.Width = ScaleWidth - (sngTemp + picSplitV.Width)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lbl数值_Click(Index As Integer)
End Sub

Private Sub mnuEditModify_Click()
'身份验证
    Dim strIdentify As String
    Dim bytType As Byte
    Dim cls医保 As New clsInsure
    
    bytType = 2
    strIdentify = cls医保.Identify(bytType)
    If strIdentify <> "" Then
        Call FillList
    End If
End Sub

Private Sub mnuEditSub_Click()
    frm医保帐户补入院.Show vbModal, Me
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

Private Sub mnuViewFind_Click()
    If frm医保帐户查找.GetFind(mstrFind) = False Then
        Exit Sub
    End If
    
    Call FillList
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

Private Sub mnuViewCustom_Click()
    If frm医保帐户信息定义.SelectFields() = True Then
        '取注册表
        mstr补充字段 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "补充字段", "")
        mstr补充字段 = Replace(mstr补充字段, "'", "")
        
        Call Fill帐户相关信息
    End If
End Sub

Private Sub msh帐户_S_EnterCell()
    Dim lng病人id As Long
    Dim rsAccount As New ADODB.Recordset
    '选择某个帐户,则提取相关信息
    Call Fill帐户相关信息
    If mint险类 = TYPE_四川眉山 Then
        '删除帐户信息(灰度级:0-正常;1-禁止个人帐户;9-帐户已撤销)
        lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
        If lng病人id = 0 Then Exit Sub
        
        '检查卡的状态，如果已经封锁，则表示要解锁；否则将封锁该卡
        gstrSQL = "Select Nvl(灰度级,0) 状态 From 保险帐户 Where 病人ID=" & lng病人id & " And 险类=" & TYPE_四川眉山
        Call OpenRecordset(rsAccount, Me.Caption)
        If rsAccount!状态 = 0 Then
            mnuEditLoss.Caption = "封锁医保卡(&L)"
        Else
            mnuEditLoss.Caption = "解除封锁(&L)"
        End If
    End If
End Sub

Private Sub msh帐户_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strSort As String
    
    If Button = 1 Then
        '按列头排序
        If msh帐户_S.MouseRow = 0 Then
            strSort = msh帐户_S.TextMatrix(0, msh帐户_S.MouseCol)
            
            If strSort = "" Then Exit Sub
            If mrs帐户.Sort = strSort Then
                mrs帐户.Sort = strSort & " DESC"
            Else
                mrs帐户.Sort = strSort
            End If
            Call 绑定数据
        End If
    
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Find"
            mnuViewFind_Click
        Case "Custom"
            mnuViewCustom_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
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
    intRow = msh帐户_S.Row
    
    '表头
    objOut.Title.Text = "医保帐户清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表项
    objRow.Add "医保类别：" & cmb险类.Text
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate, "yyyy年MM月DD日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = msh帐户_S
    
    '输出
    msh帐户_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh帐户_S.Redraw = True
    
    msh帐户_S.Row = intRow
    msh帐户_S.Col = 0: msh帐户_S.ColSel = msh帐户_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Function FillList()
    '提取所有帐户(如果权限允许,则提出密码字段)的数据
    If mrs帐户.State = adStateOpen Then mrs帐户.Close
    Dim str退休证号 As String
    Dim str单位编码 As String
    
    str单位编码 = "A.单位编码"
    Select Case mint险类
        Case TYPE_重庆市
            '重庆医保退休证号作病种编码保存
            str退休证号 = "A.退休证号 AS 病种编码"
        Case TYPE_大连市, TYPE_大连开发区
            str退休证号 = "A.退休证号 as 补助个人帐户余额"
            str单位编码 = "decode(A.单位编码,'0','企保','1','事保','其他') as 参保类别"
        Case Else
            str退休证号 = "A.退休证号"
    End Select
    
    If mcol中心("K" & mint险类) = "1" Then
        '具有医保中心
        gstrSQL = " Select C.名称 as 中心,A.卡号,A.医保号,P.病人ID,P.姓名,P.性别,To_Char(P.出生日期,'yyyy-MM-dd') as  出生日期,P.身份证号 " & _
                  "        ,E.名称 人员身份,A.人员身份 as 身份编码," & str单位编码 & "," & str退休证号 & ",D.名称 as 病种,Decode(A.当前状态,0,'普通','在院') as 状态,A.帐户余额,to_char(A.就诊时间,'yyyy-MM-dd') as 就诊时间  " & _
                  " " & IIf(mint险类 = TYPE_四川眉山, ",备注 封锁信息", "") & IIf(mint险类 = TYPE_重庆市, ",A.并发症", "") & _
                  " From 保险帐户 A,病人信息 P,保险中心目录 C,保险病种 D,保险人群 E " & _
                  " Where A.病人ID = P.病人ID and A.险类=C.险类 and A.中心=C.序号 " & IIf(mint险类 = TYPE_四川眉山, " And Nvl(A.灰度级,0)<>9", "") & _
                  "       And A.险类=E.险类 and A.在职=E.序号 And A.病种ID=D.ID(+) And A.险类=" & mint险类 & _
                  mstrFind & " Order by C.名称,A.卡号"
    Else
        Select Case mint险类
        Case TYPE_大连市, TYPE_大连开发区
            gstrSQL = " Select '' as 中心,A.卡号,A.医保号,P.病人ID,P.姓名,P.性别,To_Char(P.出生日期,'yyyy-MM-dd') as  出生日期,P.身份证号 " & _
                      "        ,E.名称 人员身份,A.人员身份 as 身份编码," & str单位编码 & "," & str退休证号 & ",D.名称 as 病种," & _
                      "        decode(参保类别1, 0,'不享受高额',1,'享受高额','医疗保险不可用') as 参保类别1," & _
                      "        decode(参保类别2, 0,'不享受',1,'商业','公务员') as 参保类别2," & _
                      "        decode(参保类别3, 0,'企保','事保') as 参保类别3," & _
                      "        decode(参保类别4, 0,'生育不可用',1,'生育可用','生育不可用') as 参保类别4," & _
                      "        decode(参保类别5, 0,'工伤不可用',1,'工伤可用','工伤不可用') as 参保类别5," & _
                      "        to_char(最高限额,'99999999999999.99') as 最高限额," & _
                      "         Decode(A.当前状态,0,'普通','在院') as 状态,A.帐户余额,to_char(A.就诊时间,'yyyy-MM-dd') as 就诊时间  " & _
                      " " & IIf(mint险类 = TYPE_四川眉山, ",备注 封锁信息", "") & IIf(mint险类 = TYPE_重庆市, ",A.并发症", "") & _
                      " From 保险帐户 A,病人信息 P,保险病种 D,保险人群 E " & _
                      " Where A.病人ID = P.病人ID And A.险类=E.险类 and A.在职=E.序号 " & IIf(mint险类 = TYPE_四川眉山, " And Nvl(A.灰度级,0)<>9", "") & _
                      "       And A.病种ID=D.ID(+) And A.险类=" & mint险类 & mstrFind & " Order by A.卡号"
        Case Else
            gstrSQL = " Select '' as 中心,A.卡号,A.医保号,P.病人ID,P.姓名,P.性别,To_Char(P.出生日期,'yyyy-MM-dd') as  出生日期,P.身份证号 " & _
                      "        ,E.名称 人员身份,A.人员身份 as 身份编码," & str单位编码 & "," & str退休证号 & ",D.名称 as 病种,Decode(A.当前状态,0,'普通','在院') as 状态,A.帐户余额,to_char(A.就诊时间,'yyyy-MM-dd') as 就诊时间  " & _
                      " " & IIf(mint险类 = TYPE_四川眉山, ",备注 封锁信息", "") & IIf(mint险类 = TYPE_重庆市, ",A.并发症", "") & _
                      " From 保险帐户 A,病人信息 P,保险病种 D,保险人群 E " & _
                      " Where A.病人ID = P.病人ID And A.险类=E.险类 and A.在职=E.序号 " & IIf(mint险类 = TYPE_四川眉山, " And Nvl(A.灰度级,0)<>9", "") & _
                      "       And A.病种ID=D.ID(+) And A.险类=" & mint险类 & mstrFind & " Order by A.卡号"
        End Select
    End If
    Call OpenRecordset(mrs帐户, Me.Caption)
    
    Call 绑定数据
    Call Fill帐户相关信息
End Function

Private Sub 绑定数据()
    Dim lngCol As Long
    
    '将帐户数据装入FLEXGRID，绑定数据
    With msh帐户_S
        If mrs帐户.RecordCount <> 0 Then
            Set .DataSource = mrs帐户
            
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
            
        Else
            Set .DataSource = Nothing
            .Rows = 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(1, lngCol) = ""
            Next
        End If
        
        '隐藏多余的列
        If mcol中心("K" & mint险类) = "0" Then
            .ColWidth(col中心) = 0
        Else
            If .ColWidth(col中心) = 0 Then
                .ColWidth(col中心) = 1000
            End If
        End If
        
        .ColWidth(col病人ID) = 0
        .ColWidth(col帐户余额) = 0
        If mint险类 = TYPE_四川眉山 Then .ColWidth(.Cols - 1) = 1200
    End With
    Call SetMenu
End Sub

Private Sub InitTable()
    Dim lngCol As Integer
    '设置格式
    With msh帐户_S
        .Rows = 2
        .Cols = 16
        If mint险类 = TYPE_大连开发区 Or mint险类 = TYPE_大连市 Then
            .Cols = .Cols + 6
        End If
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
        Next
        
        If mblnLoad Then
            .TextMatrix(0, col中心) = "中心"
            .TextMatrix(0, col卡号) = "卡号"
            .TextMatrix(0, col医保号) = "医保号"
            .TextMatrix(0, col病人ID) = "病人ID"
            .TextMatrix(0, col姓名) = "姓名"
            .TextMatrix(0, col性别) = "性别"
            .TextMatrix(0, col出生日期) = "出生日期"
            .TextMatrix(0, col身份证号) = "身份证号"
            .TextMatrix(0, col人员身份) = "人员身份"
            .TextMatrix(0, col身份编码) = "身份编码"
            .TextMatrix(0, col单位编码) = "单位编码"
            .TextMatrix(0, col退休证号) = "退休证号"
            .TextMatrix(0, col病种) = "病种"
            lngCol = 0
            If mint险类 = TYPE_大连开发区 Or mint险类 = TYPE_大连市 Then
                .TextMatrix(0, col病种 + 1) = "参保类别1"
                lngCol = lngCol + 1
                .TextMatrix(0, col病种 + 2) = "参保类别2"
                lngCol = lngCol + 1
                .TextMatrix(0, col病种 + 3) = "参保类别3"
                lngCol = lngCol + 1
                .TextMatrix(0, col病种 + 4) = "参保类别4"
                lngCol = lngCol + 1
                .TextMatrix(0, col病种 + 5) = "参保类别5"
                lngCol = lngCol + 1
                .TextMatrix(0, col病种 + 6) = "最高限额"
                lngCol = lngCol + 1
            End If
            .TextMatrix(0, col状态 + lngCol) = "状态"
            .TextMatrix(0, col帐户余额 + lngCol) = "帐户余额"
            .TextMatrix(0, col就诊时间 + lngCol) = "就诊时间"
            .ColWidth(col中心) = 0
            .ColWidth(col卡号) = 900
            .ColWidth(col医保号) = 900
            .ColWidth(col病人ID) = 0
            .ColWidth(col姓名) = 800
            .ColWidth(col性别) = 400
            .ColWidth(col出生日期) = 1200
            .ColWidth(col身份证号) = 1400
            .ColWidth(col人员身份) = 800
            .ColWidth(col身份编码) = 600
            .ColWidth(col单位编码) = 600
            .ColWidth(col退休证号) = 900
            .ColWidth(col病种) = 800
            If mint险类 = TYPE_大连开发区 Or mint险类 = TYPE_大连市 Then
                .ColWidth(col病种 + 1) = 800
                .ColWidth(col病种 + 2) = 800
                .ColWidth(col病种 + 3) = 800
                .ColWidth(col病种 + 4) = 800
                .ColWidth(col病种 + 5) = 800
                .ColWidth(col病种 + 6) = 800
                .ColWidth(col状态 + 6) = 800
                .ColAlignment(col状态 + 6) = 7
                .ColWidth(col帐户余额 + 6) = 0
                .ColWidth(col就诊时间 + 6) = 1400
            Else
                .ColWidth(col状态) = 800
                .ColWidth(col帐户余额) = 0
                .ColWidth(col就诊时间) = 1400
                
            End If
            .ColWidth(col状态) = 800
            .ColWidth(col帐户余额) = 0
            .ColWidth(col就诊时间) = 1400
        End If
        
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
    
    With msh年度
        .Rows = 13: .Cols = 2
        .ColWidth(0) = 1600: .ColAlignment(0) = 1
        .ColWidth(1) = 1000: .ColAlignment(1) = 7
        
        .TextMatrix(0, 0) = "年度信息": .TextMatrix(0, 1) = "值"
        
        .TextMatrix(row住院次数, 0) = "住院次数"
        .TextMatrix(row帐户余额, 0) = "帐户余额"
        .TextMatrix(row帐户增加, 0) = "帐户增加累计"
        .TextMatrix(row帐户支出, 0) = "帐户支出累计"
        .TextMatrix(row本次起付线, 0) = "本次起付线"
        .TextMatrix(row起付线累计, 0) = "支付起付线累计"
        .TextMatrix(row统筹限额, 0) = "基本统筹支付限额"
        .TextMatrix(row进入统筹, 0) = "进入基本统筹累计"
        .TextMatrix(row统筹报销, 0) = "支付基本统筹累计"
        .TextMatrix(row大额限额, 0) = "大额统筹支付限额"
        .TextMatrix(row大额累计, 0) = "大额统筹支付累计"
        .TextMatrix(row封锁信息, 0) = "封锁信息"
    End With
End Sub

Private Function Fill帐户相关信息()
    Dim lngCount As Long, lng病人id As Long
    Dim arrayCol, strColumn As String, intColumn As Integer
    Dim rsTemp As New ADODB.Recordset
    
    '清除相关信息
    Call ClearOther
    
    lng病人id = Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col病人ID))
    If lng病人id = 0 Then
        Exit Function
    End If
    
    '提取指定帐户的相关信息
    strColumn = ""
    arrayCol = Split(mstr补充字段, ",")
    For intColumn = 0 To UBound(arrayCol)
        strColumn = strColumn & ",P." & arrayCol(intColumn)
    Next
    
    'If InStr(1, strColumn, "P.年龄") <> 0 Then strColumn = Replace(strColumn, "P.年龄", "trunc(Months_between(to_Date(to_Char(sysdate,'yyyy')||'-01'||'-01','yyyy-MM-dd'),Decode(P.出生日期,NULL,P.登记时间,P.出生日期))/12) 年龄")
    gstrSQL = " Select P.出生日期,P.工作单位,P.婚姻状况" & strColumn & _
              " From 保险帐户 A,病人信息 P " & _
              " Where A.病人ID = P.病人ID " & _
              "       And A.险类=" & mint险类 & " And A.病人ID=" & lng病人id
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp.RecordCount > 0 Then
        With msh身份信息
            For lngCount = 1 To .Rows - 1
                .TextMatrix(lngCount, 1) = IIf(IsNull(rsTemp.Fields(.TextMatrix(lngCount, 0)).Value), "", rsTemp.Fields(.TextMatrix(lngCount, 0)).Value)
            Next
        End With
    End If
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    gstrSQL = " Select * " & _
              " From 帐户年度信息 Y" & _
              " Where Y.险类=" & mint险类 & " And Y.年度=" & lbl年度.Caption & " And Y.病人ID=" & lng病人id
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '装入指定帐户的相关数据
        With msh年度
            .TextMatrix(row住院次数, 1) = Format(rsTemp("住院次数累计"), "#####;-#####; ;")
            .TextMatrix(row帐户余额, 1) = Format(Val(msh帐户_S.TextMatrix(msh帐户_S.Row, col帐户余额)), "#####0.00;-#####0.00; ;")
            .TextMatrix(row帐户增加, 1) = Format(rsTemp("帐户增加累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row帐户支出, 1) = Format(rsTemp("帐户支出累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row本次起付线, 1) = Format(rsTemp("本次起付线"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row起付线累计, 1) = Format(rsTemp("起付线累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row统筹限额, 1) = Format(rsTemp("基本统筹限额"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row进入统筹, 1) = Format(rsTemp("进入统筹累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row统筹报销, 1) = Format(rsTemp("统筹报销累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row大额限额, 1) = Format(rsTemp("大额统筹限额"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row大额累计, 1) = Format(rsTemp("大额统筹累计"), "#####0.00;-#####0.00; ;")
            .TextMatrix(row封锁信息, 1) = IIf(IsNull(rsTemp("封销信息")), "", rsTemp("封销信息"))
        End With
    End If
End Function

Private Sub ClearOther()
    Dim lngCount As Long
    
    '清空相关信息
    With msh年度
        For lngCount = 1 To .Rows - 1
            .TextMatrix(lngCount, 1) = ""
        Next
    End With
        
    With msh身份信息
        .ColWidth(0) = 1170
        .ColWidth(1) = 1380
        .TextMatrix(0, 0) = "名称"
        .TextMatrix(0, 1) = "内容"
        .ColAlignment(1) = 1
        
        .Rows = 5 + UBound(Split(mstr补充字段, ",")) '行标题,初始三行,及用户需要
        .TextMatrix(1, 0) = "出生日期"
        .TextMatrix(2, 0) = "工作单位"
        .TextMatrix(3, 0) = "婚姻状况"
        .TextMatrix(1, 1) = ""
        .TextMatrix(2, 1) = ""
        .TextMatrix(3, 1) = ""
        
        For lngCount = 4 To .Rows - 1
            .TextMatrix(lngCount, 0) = Split(mstr补充字段, ",")(lngCount - 4)
            .TextMatrix(lngCount, 1) = ""
        Next
    End With
End Sub

Private Sub 权限控制()
    If InStr(gstrPrivs, "增删改") = 0 Then
        mnuEdit.Visible = False
        mnuEditModify.Enabled = False
        mnuEditPassword.Enabled = False
        mnuEditSub.Enabled = False
        
        mnuEditPassword.Visible = False
        mnuEditSub.Visible = False
        tbrThis.Buttons("Modify").Visible = False
        tbrThis.Buttons("SplitModify").Visible = False
    End If
    
    mnuEditXE.Visible = False ' mint险类 = TYPE_大连市 Or mint险类 = TYPE_大连开发区
    mnuEditSp.Visible = mnuEditXE.Visible
    
End Sub

Private Sub SetMenu()
    Dim blnData As Boolean
        
    blnData = (mrs帐户.RecordCount > 0)
    stbThis.Panels(2).Text = "当前共有" & mrs帐户.RecordCount & "个医保帐户"
    
    tbrThis.Buttons("Print").Enabled = blnData
    tbrThis.Buttons("Preview").Enabled = blnData
    
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuEditDelete.Enabled = blnData
    mnuEditMend.Enabled = blnData
    mnuEditLoss.Enabled = blnData
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    Dim blnCanUse As Boolean
    
    gstrSQL = "select 序号,名称,nvl(具有中心,0) as 具有中心 from 保险类别 where nvl(是否禁止,0)<>1 order by 序号"
    Call OpenRecordset(rsTemp, "保险帐户")
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm医保帐户.Visible = True Then
        frm医保帐户.Show
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
    End With
    
    mnuEditPassword.Visible = (gintInsure = TYPE_自贡市 Or gintInsure = TYPE_泸州市 Or gintInsure = TYPE_成都郊县)
    mnuEditQuery.Visible = (gintInsure = TYPE_贵阳市)
    mnuEditDisease.Visible = (gintInsure = TYPE_重庆市 Or gintInsure = TYPE_重庆银海版 Or gintInsure = TYPE_自贡市 Or gintInsure = TYPE_泸州市 Or gintInsure = TYPE_沈阳市) '重庆医保和自贡医保才可见
    'Modified by 朱玉宝 20031218 地区：福州
    mnuEditRollIncome.Visible = (gintInsure = TYPE_重庆市 Or gintInsure = TYPE_福建巨龙 Or gintInsure = TYPE_南平市 Or gintInsure = TYPE_福建省 Or gintInsure = TYPE_福州市) '重庆医保才可见
    'Modified by 朱玉宝 20031218 地区：福州
    mnuEditOut.Visible = (gintInsure = TYPE_福建巨龙 Or gintInsure = TYPE_南平市 Or gintInsure = TYPE_福建省 Or _
    gintInsure = TYPE_福州市 Or gintInsure = TYPE_昆明市 Or gintInsure = TYPE_云南省 Or _
    gintInsure = TYPE_沈阳市 Or gintInsure = TYPE_重庆银海版)
    mnuEditSplit0.Visible = mnuEditPassword.Visible
    mnuEditSplit1.Visible = mnuEditDisease.Visible
    mnuEditRollAdmit.Visible = (gintInsure = TYPE_云南省 Or gintInsure = TYPE_昆明市) '由于不能使删除未结费用，所以不支持 TYPE_云南建水
    
    mnuEditXE.Visible = False ' mint险类 = TYPE_大连市 Or mint险类 = TYPE_大连开发区
    mnuEditSp.Visible = mnuEditXE.Visible
    
    If TYPE_四川眉山 = gintInsure Then
        mnuEditSplit4.Visible = True
        mnuEditDelete.Visible = True
        mnuEditMend.Visible = True
        mnuEditLoss.Visible = True
        mnuFileCard.Visible = True
        mnuFileSplit2.Visible = True
    End If
    
    mnuEditModify.Enabled = (mint险类 = gintInsure)
    mnuEditPassword.Enabled = (mint险类 = gintInsure)
    
    blnCanUse = GetInsureInit(mint险类)
    mnuEditSub.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditDisease.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditRollIncome.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditRollAdmit.Enabled = (mint险类 = gintInsure) And blnCanUse
    mnuEditQuery.Enabled = (mint险类 = gintInsure) And blnCanUse
    
    '下列医保支持核对
    If gintInsure = TYPE_成都郊县 Then
        mnuEditSplit3.Visible = True
        mnuEditVerify_Account.Visible = True
'        mnuEditVerify_Detail.Visible = True
        mnuEditVerify_Hospital.Visible = True
        mnuEditVerify_Price.Visible = True
    ElseIf gintInsure = type_米易 Then
        mnuEditSplit3.Visible = True
        mnuEditVerify_Account.Visible = True
        mnuEditVerify_ZYPrice.Visible = True
        mnuEditVerify_UpDetail.Visible = True
    '调试重庆医保银海版 204-03-31
    ElseIf gintInsure = TYPE_重庆银海版 Then
        mnuEditVerify_Price.Visible = True
    End If
    
    tbrThis.Buttons("Modify").Enabled = mnuEditModify.Enabled
    
    frm医保帐户.Show , frmParent
End Sub

Private Function GetInsureInit(ByVal intInsure As Integer) As Boolean
'功能：读取该险类是否完成医保初始化
    Dim classInsure As New clsInsure
    Dim blnCanUse As Boolean
    Dim varCanUse As Variant
    
    On Error Resume Next
    varCanUse = mcol可用("K" & intInsure)
    
    If Err <> 0 Then
        '尚未读出该医保是否可用
        blnCanUse = classInsure.InitInsure(gcnOracle, intInsure)
        '将其加入集合中
        mcol可用.Add blnCanUse, "K" & intInsure
        GetInsureInit = blnCanUse
        Exit Function
    End If
    
    GetInsureInit = varCanUse
End Function

Private Sub GetAccountInfo()
    Dim lngRow As Long
    Dim strTemp As String
    '对保险帐户进行额外的处理
    '如果重庆医保由于种种原因，在退休证号中保存的是疾病的编码，而用户需要看到疾病的名称，而疾病的信息是在前置服务器上（数据随时可能发生变化）
    
    Select Case mint险类
    Case TYPE_重庆市
        '首先读出参数，打开连接
        If mcnYB.State <> 1 Then
            gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & TYPE_重庆市
            Call OpenRecordset(mrs病种, Me.Caption)
            Do Until mrs病种.EOF
                strTemp = IIf(IsNull(mrs病种("参数值")), "", mrs病种("参数值"))
                Select Case mrs病种("参数名")
                    Case "医保服务器"
                        strServer = strTemp
                    Case "医保用户名"
                        strUser = strTemp
                    Case "医保用户密码"
                        strPass = strTemp
                End Select
                mrs病种.MoveNext
            Loop
            If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
                MousePointer = vbDefault
                Exit Sub
            End If
            
            If mrs病种.State = adStateOpen Then mrs病种.Close
            mrs病种.Open "select BZBM 编码,BZMC 名称,ZJM 简码  from BZML Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
            If mrs病种.EOF = True Then
                MousePointer = vbDefault
                MsgBox "未从医保前置服务器中读到相关病种。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '修改当前显示的所有保险帐户的病种显示
        With msh帐户_S
            lngRow = .TopRow
            Do While .RowIsVisible(lngRow)
                If Trim(.TextMatrix(lngRow, col退休证号)) <> "" And Trim(.TextMatrix(lngRow, col病种)) = "" Then
                    mrs病种.MoveFirst
                    mrs病种.Find "编码='" & .TextMatrix(lngRow, col退休证号) & "'"
                    If Not mrs病种.EOF Then
                        .TextMatrix(lngRow, col病种) = mrs病种!名称
                    End If
                End If
                lngRow = lngRow + 1
                If lngRow > .Rows - 1 Then Exit Do
            Loop
        End With
    Case Else
    End Select
End Sub
