VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmTransferCard 
   Caption         =   "药品移库单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "frmTransferCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10770
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd无库存数据筛选 
      Caption         =   "无库存数据筛选"
      Height          =   350
      Left            =   3240
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh批次信息 
      Height          =   2175
      Left            =   5880
      TabIndex        =   33
      Top             =   1095
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   2310
      TabIndex        =   32
      Top             =   1485
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExpend 
      Caption         =   "自动分解(&A)"
      Height          =   350
      Left            =   4950
      TabIndex        =   7
      Top             =   5490
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6180
      TabIndex        =   31
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7500
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6180
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
      Begin VB.CheckBox chkIn 
         Caption         =   "导"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "导入记帐单:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   4
         Top             =   960
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   23
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   20
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   19
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品移库单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "移出库房(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2160
         TabIndex        =   16
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   15
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "移入库房(&I)"
         Height          =   180
         Left            =   8040
         TabIndex        =   2
         Top             =   660
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
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
            Picture         =   "frmTransferCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1000
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
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
            Picture         =   "frmTransferCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   6615
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransferCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin VB.Label lblCode 
      Caption         =   "编码"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "列名"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(编码和名称)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅编码)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅名称)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "无库存数据筛选"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterDrug 
         Caption         =   "无库存排在最后"
         Index           =   0
      End
      Begin VB.Menu mnuFilterDrug 
         Caption         =   "删除无库存数据"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmTransferCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '是否可选库房
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5；6-冲销；10-发送,11-从入库单读取数据
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mbln申领单 As Boolean               '是否是申领单，如果是则允许执行自动分解的功能
Private mintApplyType As Integer            '申领方式：0-手工申领;1-根据消耗量;2-根据上限;3-根据下限;4-根据上下限;5-根据申领单未发数;6-根据日销售量;7-根据销售总量
Private mstrEndTime As String               '当自动申领方式为7时，返回时间范围中的结束时间
Private mbln明确批次 As Boolean             '是否明确批次，仅对申领单有效
Private mblnEnterCell As Boolean            '是否允许激活ENTERCELL（）事件，缺省为真
Private mlng出库库房 As Long
Private mlng移入库房 As Long                '用于利用入库单移库
Private mstr入库单号 As String              '用于利用入库单移库
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mint库存检查入库库房 As Integer     '仅用于冲销时对原入库库房是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                 '权限
Private mblnRS As Boolean                   '用来记录数据集的状态
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private rsDepend As New ADODB.Recordset
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集
Private mrsMyAppend As New ADODB.Recordset  '创建动态记录集

Private Const MStrCaption As String = "药品移库管理"

Private mobjPlugIn As Object '外挂部件

Private Const mlng紫色 As Long = &HC000C0

Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容

Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

Private mint移库处理流程 As Integer                   '1-需要备药、发送、接收这一过程  0-不需要这一过程
Private mint处理方式 As Integer                       '冲销时：0－正常冲销；1－产生冲销申请单据；2－审核已产生的冲销申请单据
Private mbln自动分解未完成 As Boolean                 '需要自动分解并且自动分解未完成
Private mbln下可用数量 As Boolean                     '是否下发药药房的可用数量

'=========================================================================================
Private Const mconIntCol序号 As Integer = 1
Private Const mconIntCol行号 As Integer = 2
Private Const mconIntCol药名 As Integer = 3
Private Const mconIntCol商品名 As Integer = 4
Private Const mconIntCol来源 As Integer = 5
Private Const mconIntCol基本药物 As Integer = 6
Private Const mconIntCol规格 As Integer = 7
Private Const mconIntCol分批核算 As Integer = 8
Private Const mconIntCol最大效期 As Integer = 9
Private Const mconIntCol可用数量 As Integer = 10
Private Const mconIntCol指导差价率 As Integer = 11
Private Const mconIntCol实际金额 As Integer = 12
Private Const mconIntCol实际差价 As Integer = 13
Private Const mconIntCol比例系数 As Integer = 14
Private Const mconIntCol批次 As Integer = 15
Private Const mconIntCol产地 As Integer = 16
Private Const mconIntCol单位 As Integer = 17
Private Const mconIntCol送货单位 As Integer = 18
Private Const mconIntCol批号 As Integer = 19
Private Const mconIntCol效期 As Integer = 20
Private Const mconIntCol批准文号 As Integer = 21
Private Const mconIntCol库房库存 As Integer = 22
Private Const mconIntCol对方库存 As Integer = 23
Private Const mconIntCol填写数量 As Integer = 24
Private Const mconIntCol实际数量 As Integer = 25
Private Const mconIntCol采购价 As Integer = 26
Private Const mconIntCol采购金额 As Integer = 27
Private Const mconIntCol售价 As Integer = 28
Private Const mconIntCol售价金额 As Integer = 29
Private Const mconintCol差价 As Integer = 30
Private Const mconIntCol上次供应商ID As Integer = 31
Private Const mconintCol真实数量 As Integer = 32
Private Const mconIntCol药品编码和名称 = 33
Private Const mconIntCol药品编码 = 34
Private Const mconIntCol药品名称 = 35
Private Const mconIntCol分批属性 = 36
Private Const mconIntColS  As Integer = 37             '总列数
'=========================================================================================

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng药品ID As Long
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) B, 收费项目目录 C" & _
            " Where a.单据 = 6 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPriceDigit & ") <> Round(b.现价, " & intPriceDigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0  and b.执行日期>a.填制日期" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 6 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & intPriceDigit & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & intPriceDigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 6 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(b.平均成本价," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1  and b.性质=1" & _
            " Order By 类型, 药品id, 序号"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取当前价格]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量))
        dbl成本价 = Val(mshBill.TextMatrix(lngRow, mconIntCol采购价))
        dbl零售价 = Val(mshBill.TextMatrix(lngRow, mconIntCol售价))
        dbl成本金额 = dbl成本价 * Dbl数量
        dbl零售金额 = dbl零售价 * Dbl数量
        dbl差价 = dbl零售金额 - dbl成本金额
                
        If lng药品ID <> 0 Then
            rsPrice.Filter = "类型='售价' And 药品ID=" & lng药品ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl零售价 = Val(GetFormat(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intPriceDigit))
                dbl零售金额 = Val(GetFormat(Val(FormatEx(dbl零售价, intPriceDigit)) * Dbl数量, mintMoneyDigit))
                dbl差价 = Val(GetFormat(dbl零售金额 - dbl成本金额, mintMoneyDigit))
            End If
            
            rsPrice.Filter = "类型='成本价' And 药品ID=" & lng药品ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(GetFormat(Val(FormatEx(dbl零售价, intPriceDigit)) * Dbl数量, mintMoneyDigit))
                dbl成本价 = Val(GetFormat(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intCostDigit))
                dbl成本金额 = Val(GetFormat(dbl成本价 * Dbl数量, mintMoneyDigit))
                dbl差价 = Val(GetFormat(dbl零售金额 - dbl成本金额, mintMoneyDigit))
            End If
            
            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mconIntCol售价) = GetFormat(dbl零售价, intPriceDigit)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = GetFormat(dbl零售金额, mintMoneyDigit)
                mshBill.TextMatrix(lngRow, mconIntCol采购价) = GetFormat(dbl成本价, intCostDigit)
                mshBill.TextMatrix(lngRow, mconIntCol采购金额) = GetFormat(dbl成本金额, mintMoneyDigit)
                mshBill.TextMatrix(lngRow, mconintCol差价) = GetFormat(dbl差价, mintMoneyDigit)
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "行号", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !行号 = n
                !序号 = IIf(Val(mshBill.TextMatrix(n, mconIntCol序号)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol序号)))
                !药品id = Val(mshBill.TextMatrix(n, 0))
                !批次 = Val(mshBill.TextMatrix(n, mconIntCol批次))
                
                .Update
            End If
        Next
        
    End With
End Sub
Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = "Select 审核日期,配药日期,配药人 From 药品收发记录 " & _
              "Where 单据=6 And NO=[1] And 记录状态=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查单据]", strNo)
    
    With rs
        '返回空，表示已经删除
        If .EOF Then
            CheckBill = "该单据已经被其他操作员删除！"
        End If
        If mint编辑状态 = 3 Then
            If Not IsNull(!审核日期) Then
                CheckBill = "该单据已经被其他操作员审核！"
            End If
            Exit Function
        End If
        
        If mint编辑状态 = 10 Then
            If Not IsNull(!配药日期) Then
                CheckBill = "该单据已经被其他操作员发送！"
            End If
            Exit Function
        End If
                    
        If Not IsNull(!配药人) Then
            CheckBill = "该单据已经被其他操作员备药！"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function
Private Function Auto处理移库流程(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim blnTrans As Boolean
        
    '自动处理移库流程 1－备药 2－发送 3－接收
    
    On Error GoTo errHandle
    
    If Not 检查单价(6, txtNo, False) And Not mblnUpdate Then
        '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
        MsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
        Call RefreshBill
        mblnUpdate = True
        mblnChange = True
        Exit Function
    End If
    
    If Not 药品单据审核(Txt填制人.Caption) Then Exit Function
    
    If Not bln强制保存 Then
        blnTrans = True
        gcnOracle.BeginTrans
    End If
    
    '1-
    gstrSQL = "zl_药品移库_PREPARE('" & txtNo.Tag & "','" & UserInfo.用户姓名 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "备药")
    
    '2-
    If Not ValidData Then
        If blnTrans Then
            gcnOracle.RollbackTrans
        End If
        Exit Function
    End If
    
    '如果参数下可用数量为假，则要检查可用库存是否足够
    If mbln下可用数量 = False Then
        If Not CheckStock Then
            If blnTrans Then
                gcnOracle.RollbackTrans
            End If
            Exit Function
        End If
    End If
    
    '先删除申领单，再依据当前数据产生移库单；如果是从入库转入移库的单据，则不执行
    If mint编辑状态 <> 11 And mblnChange = True Then
        If Not SaveCard(True) Then
            If blnTrans Then
                gcnOracle.RollbackTrans
            End If
            Exit Function
        End If
    End If
    
    '备药
    gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "','" & UserInfo.用户姓名 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "备药")
    '发送（下出库库房的药品可用库存）
    gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "发送")
       
   
    '3-
    If SaveCheck(True) = True Then
        If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品移库)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        Unload Me
    Else
        GoTo errHandle
    End If
    
    If Not bln强制保存 Then
        blnTrans = True
        gcnOracle.CommitTrans
    End If
    
    Auto处理移库流程 = True
    
    Exit Function
    
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Auto处理移库流程 = False
End Function

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strMsg As String
    GetDepend = False
    On Error GoTo errHandle
    
    '检查药品入出类别是否完整
    strMsg = "没有设置药品移库的入库及出库类别，请检查药品入出分类！"
    gstrSQL = "SELECT B.Id,B.系数 " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 6 "
    
    If rsDepend.State = 1 Then rsDepend.Close
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "药品移库管理")
    
    With rsDepend
        If .RecordCount = 0 Then Exit Function
        .Filter = "系数=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置药品移库的入库类别，请检查药品入出分类！"
            Exit Function
        End If
        .Filter = "系数=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置药品移库的出库类别，请检查药品入出分类！"
            Exit Function
        End If
        .Filter = 0
        .Close
    End With
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
    Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False, Optional int处理方式 As Integer = 0)
    mblnSave = False
    mblnSuccess = False
    If int编辑状态 = 11 Then
        mstr入库单号 = str单据号
        mstr单据号 = ""
    Else
        mstr单据号 = str单据号
    End If
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mint处理方式 = int处理方式
    mstrPrivs = GetPrivFunc(glngSys, 1304)
    
    mint移库处理流程 = Val(zlDataBase.GetPara("移库流程", glngSys, 模块号.药品移库))
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    mblnEdit = False
    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 6 Then
        mblnEdit = False
        
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint处理方式 = 1 Then
            CmdSave.Caption = "申请冲销(&O)"
            CmdSave.Width = CmdSave.Width + 200
        ElseIf mint处理方式 = 2 Then
            CmdSave.Caption = "审核冲销(&V)"
            CmdSave.Width = CmdSave.Width + 200
            
            cmdAllSel.Visible = False
            cmdAllCls.Visible = False
        Else
            CmdSave.Caption = "冲销(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    ElseIf mint编辑状态 = 11 Then
        mblnEdit = True
        
        '仅当用户具有审核权限并且不需要备药发送过程时，可以直接审核
        If IsHavePrivs(mstrPrivs, "审核") And mint移库处理流程 = 0 Then
            CmdSave.Caption = "审核(&V)"
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
End Sub

Private Sub cboEnterStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        If mblnRS Then
            Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房)
        End If
        mblnRS = True
    End If
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboEnterStock
        If .ListCount = 0 Then Exit Sub
        If .ListIndex <> Val(.Tag) Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("如果改变移入库房，有可能要改变相应药品的单位和数量，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    cboEnterStock.Tag = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = Val(.Tag)
                End If
            Else
                .Tag = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsStock As New ADODB.Recordset
    Dim lngEnterStockIndex As Long
    Dim blnHaveIndex As Boolean
    
    '检查并装入移入库房
    On Error Resume Next
    
    lngEnterStockIndex = 0
    blnHaveIndex = False
    
    Set rsStock = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), MStrCaption, True, 1304)
    
    With rsStock
         cboEnterStock.Clear
         Do While Not .EOF
             cboEnterStock.AddItem !名称
             cboEnterStock.ItemData(cboEnterStock.NewIndex) = !id
             If Not blnHaveIndex And mint编辑状态 = 11 Then
                 If .Fields(0) = mlng移入库房 Then
                     lngEnterStockIndex = .AbsolutePosition - 1
                     blnHaveIndex = True
                 End If
             End If
             .MoveNext
         Loop
         cboEnterStock.ListIndex = 0
         
         If cboEnterStock.ListCount > 0 Then
            If cboEnterStock.ListCount > Val(cboEnterStock.Tag) Or (lngEnterStockIndex <> 0 And cboEnterStock.ListCount > lngEnterStockIndex) Then
                cboEnterStock.ListIndex = IIf(lngEnterStockIndex = 0, Val(cboEnterStock.Tag), lngEnterStockIndex)
                cboEnterStock.Tag = cboEnterStock.ListIndex
            End If
         End If
             
    End With
    
    mlng出库库房 = cboStock.ItemData(cboStock.ListIndex)
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint库存检查入库库房 = MediWork_GetCheckStockRule(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        Dim blnReturn As Boolean
        
        blnReturn = False
        
        cboStock_Validate blnReturn
        If blnReturn = True Then Exit Sub
        
        zlCommFun.PressKey (vbKeyTab)
    End If
    
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim blnReturn As Boolean

    If KeyAscii <> 13 Then Exit Sub
    blnReturn = False
    cboEnterStock_Validate blnReturn
    If blnReturn = True Then Exit Sub

    With mshBill
        .SetFocus
        .Row = 1
        .Col = mconIntCol药名
    End With
        
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("如果改变移出库房，有可能要改变相应药品的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    
                    mlng出库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
                
                mlng出库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            End If
        End If
        
        
    End With
End Sub

Private Sub chkIn_Click()
    txtIn.Enabled = chkIn.Value
    If chkIn.Value Then
        txtIn.SetFocus
    Else
        txtIn.Text = ""
    End If
End Sub


Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol实际数量) = GetFormat(0, mintNumberDigit)
                .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconintCol差价) = GetFormat(0, mintMoneyDigit)
            End If
        Next
    End With
    Call 显示合计金额
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol实际数量) = .TextMatrix(intRow, mconIntCol填写数量)
                .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(.TextMatrix(intRow, mconIntCol填写数量) * .TextMatrix(intRow, mconIntCol采购价), mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(.TextMatrix(intRow, mconIntCol填写数量) * .TextMatrix(intRow, mconIntCol售价), mintMoneyDigit)
                .TextMatrix(intRow, mconintCol差价) = GetFormat(.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconIntCol采购金额), mintMoneyDigit)
            End If
        Next
    End With
    Call 显示合计金额
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExpend_Click()
    Call AutoExpend
End Sub

'查找
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntCol药品编码和名称, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码和名称)
                End If
            End If
        Next
    End With
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln库房 As Boolean
    Dim bln分批 As Boolean
    Dim intRow As Integer
    Dim lng药品ID As Long
    
    On Error GoTo ErrHand
    '发送的程序处理流程：自动分解、检查库存、删除原单据、按现有数据产生新的移库单、重新备药、发送
    '审核的程序处理流程：审核单据（如果实际数量与填写数量不符，需修正出库库房的可用数量），下出库库房的实际数量、上入库库房的可用与实际数量级
    
    '设置排序数据集
    Call SetSortRecord
   
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 10 Then        '发送
        '考虑如果不分解，则库存检查过不了，因此此处不检查，强制用户手工点击分解功能
        'If Not AutoExpend(True) Then Exit Sub
        
        If mbln自动分解未完成 = True Then
            MsgBox "有药品未进行自动分解，请先执行自动分解！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ValidData Then Exit Sub
        
        '如果参数下可用数量为假，则要检查可用库存是否足够
        If mbln下可用数量 = False Then
            If Not CheckStock Then Exit Sub
        End If
 
        '检查是否已备药
        gstrSQL = "Select 1 From 药品收发记录 Where 单据=6 And NO=[1] And 配药人 Is Not NULL"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查是否备药]", txtNo.Tag)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "该单据已被其它操作员取消备药，当前操作中止！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查是否已发送
        gstrSQL = "Select 1 From 药品收发记录 Where 单据=6 And NO=[1] And 配药日期 Is Not NULL"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查是否发送]", txtNo.Tag)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "该单据已被其它操作员发送，当前操作中止！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '先删除申领单，再依据当前数据产生移库单
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        '备药
        gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "','" & Txt审核人.Caption & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, "备药")
        '发送（下出库库房的药品可用库存）
        gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, "发送")
        
        gcnOracle.CommitTrans
        blnTrans = True
        
        If Val(zlDataBase.GetPara("发送打印", glngSys, 模块号.药品移库)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then       '审核
        '移库审核时需要先判断是分批但又没有批次的记录
        If cmdExpend.Visible = True And mbln自动分解未完成 = True Then '自动分解不可见表示发药时已经自动分解；mbln自动分解未完成记录审核时已经自动分解
            If cmdExpend.Enabled = True Then
                bln库房 = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))
                With mshBill
                    For intRow = 1 To .rows - 1
                        lng药品ID = Val(.TextMatrix(intRow, 0))
                        If lng药品ID <> 0 Then
                            gstrSQL = " Select Nvl(A.药库分批,0) 药库分批,Nvl(A.药房分批,0) 药房分批" & _
                                              " From 药品规格 A" & _
                                              " Where A.药品ID =[1] "
                            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取分批属性]", lng药品ID)
                            bln分批 = IIf(bln库房, (rsTemp!药库分批 = 1), (rsTemp!药房分批 = 1))
                            If bln分批 = True And Val(.TextMatrix(intRow, mconIntCol批次)) = 0 Then
                                MsgBox .TextMatrix(intRow, mconIntCol药品名称) & "是不按批次移库药品，请先自动分解后再审核！", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Next
                End With
            End If
        End If
        
        If Not 检查单价(6, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '检查分批属性和批号关系，分批药品移库必须录入批号和产地
        If CheckBatchNum = False Then
            Exit Sub
        End If
        
        
        '判断是否自动执行移库流程，如果是就自动完成备药、发送、接收过程
        If mint移库处理流程 = 0 Then
            BlnSuccess = Auto处理移库流程
            Exit Sub
        End If

        
        '执行常规审核操作
        If Not SendPhysic Then Exit Sub
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '如果参数下可用数量为假，则要检查可用库存是否足够
        If mbln下可用数量 = False Then
            If Not CheckStock Then
                If blnTrans Then
                    gcnOracle.RollbackTrans
                End If
                Exit Sub
            End If
        End If
        
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            '备药
            gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "','" & UserInfo.用户姓名 & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "备药")
            '发送（下出库库房的药品可用库存）
            gstrSQL = "zl_药品移库_Prepare('" & txtNo.Tag & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "发送")
        End If
        
        If Not SaveCheck(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If

        gcnOracle.CommitTrans
        
        If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品移库)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
        If mblnChange = False And mint处理方式 <> 2 Then
            MsgBox "请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确实要冲销单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品移库)) = 1 And mint处理方式 = 2 Then
                    '打印
                    If IsHavePrivs(mstrPrivs, "单据打印") Then
                        printbill
                    End If
                End If
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    '修改状态要检查下单价
    If mint编辑状态 = 2 Then
        If Not 检查单价(6, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    If mint编辑状态 = 1 Then '新增保存时，判断价格是否已经更新
        If 检查价格 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    
    '从入库转入移库操作的单据，如果具有审核权限，则保存单据后自动审核
    If mint编辑状态 = 11 And CmdSave.Caption = "审核(&V)" Then
        blnTrans = True
        gcnOracle.BeginTrans
        
        '保存单据
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        mstr单据号 = txtNo.Tag
        
        '执行执行自动审核操作
        If Not Auto处理移库流程(True) Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        gcnOracle.CommitTrans
        blnTrans = True
        Unload Me
        Exit Sub
    End If
    
    BlnSuccess = SaveCard
    
    If BlnSuccess = True Then
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.药品移库)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    
    txt摘要.Text = ""
    cboEnterStock.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
    
    If mint编辑状态 = 11 Then
        Unload Me
    End If
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd无库存数据筛选_Click()
    PopupMenu mnuFilter, 2
End Sub

Private Sub Form_Activate()
    Debug.Print "结束装载：" & Now
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 6 Then
                MsgBox "该单据已没有可以冲销的药品，请检查！", vbOKOnly, gstrSysName
            Else
                '单据已被删除
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 4
            '请设置流向控制
            MsgBox "该库房未设置药品流向控制！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsPara As New ADODB.Recordset
    
    mblnUpdate = False
    mblnEnterCell = False
    mintBatchNoLen = GetBatchNoLen()
    mintSelectStock = Val(zlDataBase.GetPara("是否选择库房", glngSys, 模块号.药品移库))
    mintApplyType = -1
    mstrEndTime = ""
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品移库管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    
    txtNo = mstr单据号
    txtNo.Tag = mstr单据号

    If mint编辑状态 = 11 Then
        mlng移入库房 = mfrmMain.cboEnterStock.ItemData(mfrmMain.cboEnterStock.ListIndex)
    End If
    
    '取系统参数“明确申领药品批次”
    mbln明确批次 = (gtype_UserSysParms.P174_药品移库明确批次 = 1)
    
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
    
    '出库库房缺省为主界面当前选择的库房，对于新增有效
    On Error Resume Next
    mlng出库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        
    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
        
    mstrTime_Start = GetBillInfo(6, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '根据系统参数决定药房人员查看单据时，是否显示成本价
    mshBill.ColWidth(mconIntCol采购价) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol采购金额) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconIntCol上次供应商ID) = 0
    mshBill.ColWidth(mconintCol真实数量) = 0
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint库存检查入库库房 = MediWork_GetCheckStockRule(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    mshBill.MsfObj.FixedCols = 4
    mshBill.CmdVisible = False
    mblnEnterCell = True
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim str批次 As String
    Dim strArray As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim lng入库库房 As Long, lng出库库房 As Long
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPriceDigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    Dim rsPrice As ADODB.Recordset
    
    '库房
    mbln申领单 = False
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品移库)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        strSqlOrder = "药品编码"
    ElseIf strCompare = "2" Then
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            strSqlOrder = "通用名"
        Else
            strSqlOrder = "Nvl(商品名, 通用名)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    On Error GoTo ErrHand
   
    '取指定单据的出库库房与入库库房
    gstrSQL = " Select 库房ID,对方部门ID From 药品收发记录" & _
              " Where NO=[1] And 单据=6 And 入出系数=-1 And Rownum<2"
    Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取指定单据的出库库房与入库库房]", mstr单据号)
              
    If rsInitCard.RecordCount <> 0 Then
        lng出库库房 = rsInitCard!库房id
        lng入库库房 = rsInitCard!对方部门id
        
        If lng出库库房 > 0 Then
            mlng出库库房 = lng出库库房
                
            Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        End If
    Else
        lng出库库房 = mlng出库库房
    End If
    
    intCostDigit = mintCostDigit
    intPriceDigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    
    If mint编辑状态 <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                If .ItemData(i) = lng出库库房 Then cboStock.ListIndex = cboStock.ListCount - 1
            Next
            mintcboIndex = cboStock.ListIndex
            '如果没有指定的药房，将其加入
            If mintcboIndex = -1 Then
                gstrSQL = "Select ID,名称 From 部门表 Where ID=[1] "
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[如果没有指定的药房，将其加入]", lng出库库房)
                
                cboStock.AddItem rsInitCard!名称
                cboStock.ItemData(cboStock.NewIndex) = rsInitCard!id
                cboStock.ListIndex = cboStock.ListCount - 1
            End If
            mintcboIndex = cboStock.ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            If cboEnterStock.ListCount <> 0 Then
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    If cboEnterStock.ListCount > 1 Then
                        cboEnterStock.ListIndex = cboEnterStock.ListIndex + 1
                    End If
                End If
            Else
                mintParallelRecord = 4
                Exit Sub
            End If
        Case 2, 3, 4, 6, 10, 11
            initGrid
            '检查该单据是否是申领单据
            gstrSQL = " Select Nvl(发药方式,0) 申领 From 药品收发记录 " & _
                      " Where 单据=6 And NO=[1] And 入出系数 = -1 and rownum = 1"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查该单据是否是申领单据]", mstr单据号)
                      
            If Not rsTemp.EOF Then
                mbln申领单 = (rsTemp!申领 = 1)
                If mbln申领单 Then LblTitle.Caption = "药品申领单"
            End If
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select distinct b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 6 and a.no=[1] and a.入出系数=-1"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!名称
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "B.送货单位,B.送货包装,C.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case mconint门诊单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,(A.实际数量 / B.门诊包装) AS 实际数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,B.门诊包装 as 比例系数,"
                Case mconint住院单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,(A.实际数量 / B.住院包装) AS 实际数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,B.住院包装 as 比例系数,"
                Case mconint药库单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,(A.实际数量 / B.药库包装) AS 实际数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,B.药库包装 as 比例系数,"
            End Select
            
            Select Case mint编辑状态
            Case 6
                '正常冲销
                If mint处理方式 <> 2 Then
                    gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.药品ID,A.序号,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, E.名称 As 商品名," & _
                        "     B.药品来源,B.基本药物,C.规格,C.产地 AS 原产地,A.产地, A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算," & _
                        "     B.最大效期,A.效期," & strUnitQuantity & _
                        "     A.成本金额,0 零售金额, 0 差价,D.摘要,A.库房ID,A.对方部门ID,C.是否变价,B.药房分批 AS 药房分批核算,A.上次供应商ID,A.批准文号,A.填写数量 真实数量 " & _
                        "     FROM " & _
                        "         (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,0 实际数量,SUM(成本金额) AS 成本金额,药品ID,序号,产地, 批号,效期,NVL(批次,0) 批次,扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0) 上次供应商ID,批准文号" & _
                        "          FROM 药品收发记录 X " & _
                        "          WHERE NO=[1] AND 单据=6 AND 入出系数=-1 " & _
                        "          GROUP BY 药品ID,序号,产地,批号,效期,NVL(批次,0),扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0),批准文号" & _
                        "          HAVING SUM(实际数量)<>0 ) A," & _
                        "     药品规格 B,收费项目目录 C,收费项目别名 E, " & _
                        " (Select 序号, 摘要 From 药品收发记录 " & _
                        "  Where 单据 = 6 And NO = [1] And 入出系数 = -1 And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) D " & _
                        "     WHERE A.药品ID = B.药品ID AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND B.药品ID=C.ID And A.序号 = D.序号) W," & _
                        "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                        "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                        " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                        " ORDER BY " & strSqlOrder
                Else
                    '用于审核冲销时，显示未审核的申请冲销单据
                    gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.药品ID,A.序号,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, E.名称 As 商品名," & _
                        "     B.药品来源,B.基本药物,C.规格,C.产地 AS 原产地,A.产地, A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算," & _
                        "     B.最大效期,A.效期," & strUnitQuantity & "A.成本金额,A.零售金额, A.差价,A.配药人, " & _
                        "     A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.对方部门ID,C.是否变价,B.药房分批 AS 药房分批核算,NVL(A.供药单位ID,0) 上次供应商ID,A.批准文号,A.填写数量 真实数量 " & _
                        "     FROM 药品收发记录 A, 药品规格 B,收费项目目录 C,收费项目别名 E " & _
                        "     WHERE A.药品ID = B.药品ID AND B.药品ID=C.ID AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                        "     AND A.记录状态 =[3] " & _
                        "     AND A.单据 = 6 AND A.入出系数=-1 AND A.NO =[1] ) W," & _
                        "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                        "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                        " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                        " ORDER BY " & strSqlOrder
                End If
            Case 11
                gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价,z.平均成本价 * w.比例系数 As 成本价," & _
                    " z.零售价*w.比例系数 as 零售价, w.实际数量 * z.平均成本价 * w.比例系数 As 成本金额,z.实际数量/w.比例系数 as 库存数量 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, E.名称 As 商品名," & _
                    "     B.药品来源,B.基本药物,C.规格,C.产地 AS 原产地,A.产地, A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算," & _
                    "     B.最大效期,A.效期," & strUnitQuantity & "A.成本金额,A.零售金额, A.差价,A.配药人,A.填写数量 真实数量, " & _
                    "     A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID," & cboEnterStock.ItemData(cboEnterStock.ListIndex) & " 对方部门ID,C.是否变价,B.药房分批 AS 药房分批核算,NVL(A.供药单位ID,0) 上次供应商ID,A.批准文号 " & _
                    "     FROM 药品收发记录 A, 药品规格 B,收费项目目录 C,收费项目别名 E " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=C.ID AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    "     AND A.记录状态 =[3] " & _
                    "     AND A.单据 = 1 AND A.NO = [1] And A.审核人 Is Not Null) W," & _
                    "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际数量,实际金额,实际差价,平均成本价,nvl(零售价,0) as 零售价  " & _
                    "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z, " & _
                    "     (Select Distinct 收费细目id From 收费执行科室 f Where 执行科室ID=[4] ) Y " & _
                    " WHERE W.药品ID=Z.药品ID(+) AND W.药品id=Y.收费细目id AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " ORDER BY " & strSqlOrder
            Case Else
                gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, E.名称 As 商品名," & _
                    "     B.药品来源,B.基本药物,C.规格,C.产地 AS 原产地,A.产地, A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算," & _
                    "     B.最大效期,A.效期," & strUnitQuantity & "A.成本金额,A.零售金额, A.差价,A.配药人,Nvl(A.单量,-1) As 申领方式,A.频次 As 结束时间, " & _
                    "     A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.对方部门ID,C.是否变价,B.药房分批 AS 药房分批核算,NVL(A.供药单位ID,0) 上次供应商ID,A.批准文号,A.填写数量 真实数量 " & _
                    "     FROM 药品收发记录 A, 药品规格 B,收费项目目录 C,收费项目别名 E " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=C.ID AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    "     AND A.记录状态 =[3] " & _
                    "     AND A.单据 = 6 AND A.入出系数=-1 AND A.NO =[1] ) W," & _
                    "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " ORDER BY " & strSqlOrder
            End Select

            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(mint编辑状态 = 11, mstr入库单号, mstr单据号), cboStock.ItemData(cboStock.ListIndex), mint记录状态, cboEnterStock.ItemData(cboEnterStock.ListIndex))
                        
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
            Case 2, 6, 10, 11
                Txt填制人 = UserInfo.用户姓名
                Txt填制日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint编辑状态 = 6 Or mint编辑状态 = 10 Then
                    Txt审核人 = UserInfo.用户姓名
                    Txt审核日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
                If mint编辑状态 = 10 Then
                    Txt审核人 = Nvl(rsInitCard!配药人)
                    Txt填制人 = rsInitCard!填制人
                    Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Lbl审核人.Caption = "备药人"
                    Lbl审核日期.Caption = "发送日期"
                End If
            Case Else
                Txt填制人 = rsInitCard!填制人
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                mintApplyType = rsInitCard!申领方式
                mstrEndTime = Nvl(rsInitCard!结束时间)
            End If
            
            Dim intCount As Integer
            With cboEnterStock
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = lng入库库房 Then
                        .ListIndex = intCount
                        .Tag = intCount
                        Exit For
                    End If
                Next
            End With
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = rsInitCard!通用名
                    Else
                        str药名 = IIf(IsNull(rsInitCard!商品名), rsInitCard!通用名, rsInitCard!商品名)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol药品编码和名称) = rsInitCard!药品编码 & str药名
                    .TextMatrix(intRow, mconIntCol药品编码) = rsInitCard!药品编码
                    .TextMatrix(intRow, mconIntCol药品名称) = str药名
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
                    Else
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsInitCard!商品名), "", rsInitCard!商品名)

                    .TextMatrix(intRow, mconIntCol来源) = Nvl(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = Nvl(rsInitCard!基本药物)
                    If mint编辑状态 <> 11 Then .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    If IIf(IsNull(rsInitCard!比例系数), 0, rsInitCard!比例系数) = 0 Or Nvl(rsInitCard!送货包装) = "" Or Nvl(rsInitCard!送货单位) = "" Then
                        .TextMatrix(intRow, mconIntCol送货单位) = ""
                    Else
                        .TextMatrix(intRow, mconIntCol送货单位) = rsInitCard!送货单位 & "(1" & rsInitCard!送货单位 & "=" & zlStr.FormatEx(rsInitCard!送货包装 / rsInitCard!比例系数, 1, , True) & rsInitCard!单位 & ")"
                    End If
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol填写数量) = GetFormat(IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!填写数量, intNumberDigit)
                    .TextMatrix(intRow, mconIntCol实际数量) = GetFormat(IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!实际数量, intNumberDigit)
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol采购价) = GetFormat(rsInitCard!成本价, intCostDigit)
                        If Val(rsInitCard!填写数量) <> 0 And Val(.TextMatrix(intRow, mconIntCol采购价)) = 0 Then
                            .TextMatrix(intRow, mconIntCol采购价) = GetFormat((rsInitCard!零售金额 - rsInitCard!差价) / Val(rsInitCard!填写数量), intCostDigit)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol采购价) = GetFormat(rsInitCard!成本价, intCostDigit)
                        If Val(rsInitCard!实际数量) <> 0 And Val(.TextMatrix(intRow, mconIntCol采购价)) = 0 Then
                            .TextMatrix(intRow, mconIntCol采购价) = GetFormat((rsInitCard!零售金额 - rsInitCard!差价) / Val(rsInitCard!实际数量), intCostDigit)
                        End If
                    End If
                    .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(IIf(mint编辑状态 = 6 And mint处理方式 <> 2, 0, IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!成本金额), intMoneyDigit)
                    
                    If mint编辑状态 = 11 Then
                        If rsInitCard!是否变价 = 0 Then
                            gstrSQL = "Select 现价 From 收费价目 Where 收费细目id = [1] And Sysdate Between 执行日期 And 终止日期"
                            
                            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "查询价格", rsInitCard!药品id)
                            .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsPrice!现价 * rsInitCard!比例系数, intPriceDigit, , True)
                            .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) * rsInitCard!实际数量, intMoneyDigit, , True)
                        Else
                            '时价
                            If rsInitCard!零售价 = 0 Then
                                If rsInitCard!库存数量 <> 0 Then
                                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!实际金额 / rsInitCard!库存数量, intPriceDigit, , True)
                                Else
                                    gstrSQL = "Select 现价 From 收费价目 Where 收费细目id = [1] And Sysdate Between 执行日期 And 终止日期"
                                    
                                    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "查询价格", rsInitCard!药品id)
                                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsPrice!现价 * rsInitCard!比例系数, intPriceDigit, , True)
                                End If
                            Else
                                .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!零售价, intPriceDigit, , True)
                            End If
                            .TextMatrix(intRow, mconIntCol售价金额) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价)) * rsInitCard!实际数量, intMoneyDigit, , True)
                        End If
                        .TextMatrix(intRow, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol售价金额)) - Val(.TextMatrix(intRow, mconIntCol采购金额)), intMoneyDigit, , True)
                    Else
                        .TextMatrix(intRow, mconIntCol售价) = GetFormat(rsInitCard!零售价, intPriceDigit)
                        .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!零售金额, intMoneyDigit)
                        .TextMatrix(intRow, mconintCol差价) = GetFormat(IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!差价, intMoneyDigit)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol最大效期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol指导差价率) = rsInitCard!指导差价率
                    .TextMatrix(intRow, mconIntCol分批核算) = IIf(IsNull(rsInitCard!分批核算), "0", rsInitCard!分批核算)
                    .TextMatrix(intRow, mconIntCol可用数量) = GetFormat(IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量), intNumberDigit)
                    .TextMatrix(intRow, mconIntCol实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mconIntCol实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    .TextMatrix(intRow, mconIntCol上次供应商ID) = rsInitCard!上次供应商ID
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconIntCol分批属性) = IIf(Get药品分批属性(Val(.TextMatrix(intRow, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex)) = True, "1", 0)
                
'                    If (mint编辑状态 = 3 Or mint编辑状态 = 10) And Nvl(rsInitCard!分批核算, 0) = 1 And Nvl(rsInitCard!批次, 0) = 0 And mbln自动分解未完成 = False Then
'                        mbln自动分解未完成 = True
'                    End If

                    If (mint编辑状态 = 3 Or mint编辑状态 = 10) And mbln自动分解未完成 = False Then
                        If Get分批属性(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0))) = 1 And Nvl(rsInitCard!批次, 0) = 0 Then
                            mbln自动分解未完成 = True
                        End If
                    End If
                    
                    Call 提示库存数(intRow)
                                        
                    If mint编辑状态 = 2 Or mint编辑状态 = 6 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Or mint编辑状态 = 11 Then
                        .TextMatrix(intRow, mconintCol真实数量) = IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * rsInitCard!真实数量
                    End If
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!药品id & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str批次 = rsInitCard!药品id & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                        If mint编辑状态 = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!填写数量), "0", rsInitCard!填写数量)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!实际数量), "0", rsInitCard!实际数量)
                        End If
                        mcolUsedCount.Add Array(str批次, strArray), str批次
                    End If
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    
    SetEdit         '设置编辑属性
    '修改或审核时可以对无库存数筛选
    If (mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10) Then
        If mint移库处理流程 = 1 And mint编辑状态 = 3 Then
            cmd无库存数据筛选.Visible = False
        Else
            cmd无库存数据筛选.Visible = True
        End If
        
    End If
    '查阅、修改或审核时，根据库存与申领数量显示单据
    If (mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 4 Or mint编辑状态 = 10) Then
'        If mbln申领单 Then
        Call ShowColor
        Select Case mint编辑状态
        Case 10 '2, 10 修改时自动分解不可见
            cmdExpend.Visible = True
        End Select
    End If
    If mint移库处理流程 = 0 And mint编辑状态 = 3 Then
        cmdExpend.Visible = True
    End If
    Call 显示合计金额
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = (mint编辑状态 = 6)
            
            If mint编辑状态 = 10 Or (mint编辑状态 = 6 And mint处理方式 <> 2) Then
                .ColData(mconIntCol实际数量) = 4
            End If
        Else
            .ColData(0) = 5
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol序号) = 5
            .ColData(mconIntCol规格) = 5
            .ColData(mconIntCol产地) = 5
            .ColData(mconIntCol单位) = 5
            .ColData(mconIntCol送货单位) = 5
            .ColData(mconIntCol批号) = 5
            .ColData(mconIntCol效期) = 5
            .ColData(mconIntCol批准文号) = 5
            If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                .ColData(mconIntCol填写数量) = 4
                .ColData(mconIntCol实际数量) = 5
            ElseIf mint编辑状态 = 3 Then
                .ColData(mconIntCol填写数量) = 5
                .ColData(mconIntCol实际数量) = 4
            ElseIf mint编辑状态 = 11 Then
                If mint移库处理流程 = 1 Then
                    .ColData(mconIntCol填写数量) = 4
                    .ColData(mconIntCol实际数量) = 5
                Else
                    .ColData(mconIntCol填写数量) = 5
                    .ColData(mconIntCol实际数量) = 4
                End If
            End If
            .ColData(mconIntCol采购价) = 5
            .ColData(mconIntCol采购金额) = 5
            .ColData(mconIntCol售价) = 5
            .ColData(mconIntCol售价金额) = 5
            .ColData(mconintCol差价) = 5
            
            .ColData(mconIntCol分批核算) = 5
            .ColData(mconIntCol可用数量) = 5
            .ColData(mconIntCol最大效期) = 5
            
            .ColData(mconIntCol指导差价率) = 5
            .ColData(mconIntCol实际金额) = 5
            .ColData(mconIntCol实际差价) = 5
            .ColData(mconIntCol比例系数) = 5
            .ColData(mconIntCol批次) = 5
                     
            .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
            .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
            .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
            .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
            .ColAlignment(mconIntCol送货单位) = flexAlignCenterCenter
            .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
            .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol填写数量) = flexAlignRightCenter
            .ColAlignment(mconIntCol实际数量) = flexAlignRightCenter
            
            .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
            .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
            .ColAlignment(mconintCol差价) = flexAlignRightCenter
            
            If mintSelectStock = 0 Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            cboEnterStock.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 4
        
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol送货单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol库房库存) = "库房库存"
        .TextMatrix(0, mconIntCol对方库存) = "对方库存"
        .TextMatrix(0, mconIntCol填写数量) = IIf(mint编辑状态 = 6, "数量", "填写数量")
        .TextMatrix(0, mconIntCol实际数量) = IIf(mint编辑状态 = 6, "冲销数量", "实际数量")
        
        .TextMatrix(0, mconIntCol采购价) = "成本价"
        .TextMatrix(0, mconIntCol采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol分批核算) = "分批核算"
        .TextMatrix(0, mconIntCol最大效期) = "最大效期"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconIntCol指导差价率) = "指导差价率"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol上次供应商ID) = "上次供应商ID"
        .TextMatrix(0, mconintCol真实数量) = "真实数量"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconIntCol分批属性) = "分批属性"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntCol送货单位) = 2000
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol填写数量) = 1100
        .ColWidth(mconIntCol实际数量) = 1100
        .ColWidth(mconIntCol采购价) = 1000
        .ColWidth(mconIntCol采购金额) = 900
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = 800
        
        .ColWidth(mconIntCol分批核算) = 0
        .ColWidth(mconIntCol可用数量) = 0
        .ColWidth(mconIntCol最大效期) = 0
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        .ColWidth(mconIntCol指导差价率) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol上次供应商ID) = 0
        .ColWidth(mconintCol真实数量) = 0
        
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        .ColWidth(mconIntCol分批属性) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol送货单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconIntCol库房库存) = 5
        .ColData(mconIntCol对方库存) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        .ColData(mconIntCol分批属性) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            cboEnterStock.Enabled = True
            txt摘要.Enabled = True
            
            If mintSelectStock = 0 Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol填写数量) = 4
            .ColData(mconIntCol实际数量) = 5
        ElseIf mint编辑状态 = 3 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol药名) = 5
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 4
        ElseIf mint编辑状态 = 6 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = True
            
            .ColData(mconIntCol药名) = 5
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 5
                
            If mint处理方式 <> 2 Then
                .ColData(mconIntCol实际数量) = 4
            End If
        ElseIf mint编辑状态 = 4 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 5
            .ColData(mconIntCol药名) = 5
        ElseIf mint编辑状态 = 11 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = True
            txt摘要.Enabled = True
            
            If mint移库处理流程 = 1 Then
                .ColData(mconIntCol填写数量) = 4
                .ColData(mconIntCol实际数量) = 5
            Else
                .ColData(mconIntCol填写数量) = 5
                .ColData(mconIntCol实际数量) = 4
            End If
            .ColData(mconIntCol药名) = 5
        End If
        
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        
        .ColData(mconIntCol分批核算) = 5
        .ColData(mconIntCol可用数量) = 5
        .ColData(mconIntCol最大效期) = 5
        .ColData(mconIntCol实际差价) = 5
        .ColData(mconIntCol实际金额) = 5
        .ColData(mconIntCol指导差价率) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol批次) = 5
        .ColData(mconIntCol上次供应商ID) = 5
        .ColData(mconintCol真实数量) = 5
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol送货单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol库房库存) = flexAlignRightCenter
        .ColAlignment(mconIntCol对方库存) = flexAlignRightCenter
        .ColAlignment(mconIntCol填写数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol实际数量) = flexAlignRightCenter
        .ColAlignment(mconintCol真实数量) = flexAlignRightCenter
        
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        If InStr(1, "346", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
    End With
    txt摘要.MaxLength = GetLength("药品收发记录", "摘要")
    chkIn.Visible = (mint编辑状态 = 1)
    txtIn.Visible = (mint编辑状态 = 1)
End Sub

Private Function CheckBatchNum() As Boolean
    '功能：用来检查分批药品批号是否为空
    '返回值：true-分批药品都有批次，false-分批药品存在批次为空情况
    Dim intRow As Integer
    
    With mshBill
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If .TextMatrix(intRow, mconIntCol分批属性) = "1" And .TextMatrix(intRow, mconIntCol批号) = "" And .TextMatrix(intRow, 0) <> "" Then
                    CheckBatchNum = False
                    MsgBox "第" & intRow & "行，入库库房是分批管理，必须录入批号！", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    Exit Function
                End If
            Next
            CheckBatchNum = True
        Else
            CheckBatchNum = True
        End If
    End With
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
   
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboEnterStock.Left = mshBill.Left + mshBill.Width - cboEnterStock.Width
    
    LblEnterStock.Left = cboEnterStock.Left - LblEnterStock.Width - 100
    
    
    With Lbl填制人
        .Top = Pic单据.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl填制人.Left + Lbl填制人.Width + 100
    End With
    
    With Lbl填制日期
        .Top = Lbl填制人.Top
        .Left = Txt填制人.Left + Txt填制人.Width + 250
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 80
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With Txt审核日期
        .Top = Lbl填制人.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl审核日期
        .Top = Lbl填制人.Top
        .Left = Txt审核日期.Left - 100 - .Width
    End With
    
    With Txt审核人
        .Top = Lbl填制人.Top - 80
        .Left = Lbl审核日期.Left - 200 - .Width
    End With
    
    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Txt审核人.Left - 100 - .Width
    End With
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 100
        '.Width = .Left - .Left
        Debug.Print .Width
    End With
    
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With cmdExpend
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width
    End With

    With cmd无库存数据筛选
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width - cmdExpend.Width - 100
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品移库管理", "药品名称显示方式", mintDrugNameShow)
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS  '卸载数据集
        mblnRS = False
        zlPlugIn_Unload mobjPlugIn
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS  '卸载数据集
    mblnRS = False
    zlPlugIn_Unload mobjPlugIn
End Sub

Private Function SaveCheck(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim rs类别 As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng库房ID As Long
    Dim lng对方部门id As Long
    Dim str审核人 As String
    
    Dim lng药品ID As Long
    Dim str产地 As String
    Dim lng出批次 As Long
    Dim num填写数量 As Double
    Dim num实际数量 As Double
    Dim num成本价 As Double
    Dim num成本金额 As Double
    Dim dbl售价 As Double
    Dim num零售金额 As Double
    Dim num差价 As Double
    Dim lng出类别id As Long
    Dim lng入类别id As Long
    Dim str批号 As String
    Dim dat效期 As String
    Dim dat审核日期 As String
    Dim int序列号 As Integer
    Dim lng上次供应商ID As Long
    Dim str批准文号 As String
        
    Dim arrSql As Variant
    Dim n As Integer
    
    arrSql = Array()
    mblnSave = False
    SaveCheck = False
    
    '检查该单据是否在进入编辑界面后，被其他操作员修改；如果是入库转入移库单据，则不检查
    If mint编辑状态 <> 11 Then
        mstrTime_End = GetBillInfo(6, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Function
        End If
   
        If mint移库处理流程 <> 0 Then
            If mstrTime_End > mstrTime_Start Then
                MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    str审核人 = UserInfo.用户姓名
    strNo = txtNo.Tag
    
    gstrSQL = "SELECT b.系数,b.id AS 类别id " _
            & "FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID AND a.单据 = 6 "
    Set rs类别 = zlDataBase.OpenSQLRecord(gstrSQL, "药品移库管理")
    
    If rs类别.EOF Then
        MsgBox "对不起，药品入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rs类别.RecordCount < 2 Then
        MsgBox "对不起，药品入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    rs类别.MoveFirst
    Do While Not rs类别.EOF
        If rs类别!系数 = 1 Then
            lng入类别id = rs类别!类别id
        Else
            lng出类别id = rs类别!类别id
        End If
        rs类别.MoveNext
    Loop
    rs类别.Close
    
    dat审核日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo errHandle
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lng药品ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                lng出批次 = .TextMatrix(intRow, mconIntCol批次)
                
                If .TextMatrix(intRow, mconIntCol填写数量) = .TextMatrix(intRow, mconIntCol实际数量) Then
                    num填写数量 = .TextMatrix(intRow, mconintCol真实数量)
                    num实际数量 = .TextMatrix(intRow, mconintCol真实数量)
                Else
                    num填写数量 = GetFormat(Val(.TextMatrix(intRow, mconIntCol填写数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                    num实际数量 = GetFormat(Val(.TextMatrix(intRow, mconIntCol实际数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
'                num成本价 = GetFormat(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                num成本价 = Get成本价(lng药品ID, lng库房ID, lng出批次)
                num成本金额 = Val(.TextMatrix(intRow, mconIntCol采购金额))
'                dbl售价 = GetFormat(Val(.TextMatrix(intRow, mconIntCol售价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dbl售价 = Get售价(Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1, lng药品ID, lng库房ID, lng出批次)
                num零售金额 = Val(.TextMatrix(intRow, mconIntCol售价金额))
                num差价 = Val(.TextMatrix(intRow, mconintCol差价))
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                dat效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                
                If mint编辑状态 = 11 And CmdSave.Caption = "审核(&V)" Then
                    '由于是直接填单后审核，所以审核时传的序号实际应该是单数（即出库那笔的序号）
                    int序列号 = 2 * intRow - 1 '2 * Val(.TextMatrix(intRow, mconIntCol序号)) - 1
                Else
                    int序列号 = Val(.TextMatrix(intRow, mconIntCol序号))
                End If
                
                lng上次供应商ID = .TextMatrix(intRow, mconIntCol上次供应商ID)
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                
                gstrSQL = "zl_药品移库_Verify("
                '序号
                gstrSQL = gstrSQL & int序列号
                '库房ID
                gstrSQL = gstrSQL & "," & lng库房ID
                '对方部门ID
                gstrSQL = gstrSQL & "," & lng对方部门id
                '药品ID
                gstrSQL = gstrSQL & "," & lng药品ID
                '产地
                gstrSQL = gstrSQL & ",'" & str产地 & "'"
                '出批次
                gstrSQL = gstrSQL & "," & lng出批次
                '实际数量
                gstrSQL = gstrSQL & "," & num实际数量
                '成本价
                gstrSQL = gstrSQL & "," & num成本价
                '成本金额
                gstrSQL = gstrSQL & "," & num成本金额
                '零售金额
                gstrSQL = gstrSQL & "," & num零售金额
                '差价
                gstrSQL = gstrSQL & "," & num差价
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '审核人
                gstrSQL = gstrSQL & ",'" & str审核人 & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & str批号 & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '审核日期
                gstrSQL = gstrSQL & ",to_date('" & dat审核日期 & "','yyyy-mm-dd HH24:MI:SS')"
                '供药单位ID
                gstrSQL = gstrSQL & "," & IIf(lng上次供应商ID = 0, "NULL", lng上次供应商ID)
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '售价
                gstrSQL = gstrSQL & "," & dbl售价
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lng药品ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSql, MStrCaption, False, Not bln强制保存) Then Exit Function

    If Not bln强制保存 Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '外挂功能
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, lng库房ID, strNo, 单据号.药品移库)
    
    Exit Function
errHandle:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mnuFilterDrug_Click(Index As Integer)
    
    If Index = 1 Then
        If MsgBox("你确实要删除实际数量为0的药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call MyAppend
    Call AddAppend(Index)
    With mrsMyAppend
        mshBill.ClearBill
        mshBill.rows = 2
        
        If Not .EOF Then .MoveFirst
        Do While Not .EOF
            mshBill.TextMatrix(mshBill.rows - 1, 0) = .Fields!药品id
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol行号) = mshBill.rows - 1
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol序号) = (mshBill.rows - 2) * 2 + 1
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol药名) = .Fields!药名
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol商品名) = .Fields!商品名
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol来源) = .Fields!来源
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol基本药物) = .Fields!基本药物
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol规格) = .Fields!规格
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol分批核算) = .Fields!分批核算
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol最大效期) = .Fields!最大效期
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol可用数量) = zlStr.FormatEx(.Fields!可用数量, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol指导差价率) = .Fields!指导差价率
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol实际金额) = .Fields!实际金额
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol实际差价) = .Fields!实际差价
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol比例系数) = .Fields!比例系数
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol批次) = .Fields!批次
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol产地) = .Fields!产地
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol单位) = .Fields!单位
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol送货单位) = .Fields!送货单位
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol批号) = .Fields!批号
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol效期) = .Fields!效期
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol批准文号) = .Fields!批准文号
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol库房库存) = zlStr.FormatEx(.Fields!库房库存, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol对方库存) = zlStr.FormatEx(.Fields!对方库存, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol填写数量) = zlStr.FormatEx(.Fields!填写数量, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol实际数量) = zlStr.FormatEx(.Fields!实际数量, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol采购价) = zlStr.FormatEx(.Fields!采购价, mintCostDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol采购金额) = zlStr.FormatEx(.Fields!采购金额, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol售价) = zlStr.FormatEx(.Fields!售价, mintPriceDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol售价金额) = zlStr.FormatEx(.Fields!售价金额, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconintCol差价) = zlStr.FormatEx(.Fields!差价, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol上次供应商ID) = .Fields!上次供应商ID
            mshBill.TextMatrix(mshBill.rows - 1, mconintCol真实数量) = zlStr.FormatEx(.Fields!真实数量, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol药品编码和名称) = .Fields!药品编码和名称
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol药品编码) = .Fields!药品编码
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol药品名称) = .Fields!药品名称
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol分批属性) = .Fields!分批属性
            
            mshBill.rows = mshBill.rows + 1
            .MoveNext
        Loop
        
        mshBill.Row = mshBill.rows - 1
    End With
    
    Call ShowColor
End Sub

Private Sub MyAppend()
    '创建动态纪录集
    Set mrsMyAppend = New ADODB.Recordset
    With mrsMyAppend
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "来源", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "基本药物", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "分批核算", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "最大效期", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "可用数量", adDouble, 18, adFldIsNullable
        .Fields.Append "指导差价率", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "实际金额", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "实际差价", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "比例系数", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "送货单位", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "效期", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "批准文号", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "库房库存", adDouble, 18, adFldIsNullable
        .Fields.Append "对方库存", adDouble, 18, adFldIsNullable
        .Fields.Append "填写数量", adDouble, 18, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
        .Fields.Append "采购价", adDouble, 18, adFldIsNullable
        .Fields.Append "采购金额", adDouble, 18, adFldIsNullable
        .Fields.Append "售价", adDouble, 18, adFldIsNullable
        .Fields.Append "售价金额", adDouble, 18, adFldIsNullable
        .Fields.Append "差价", adDouble, 18, adFldIsNullable
        .Fields.Append "上次供应商ID", adDouble, 18, adFldIsNullable
        .Fields.Append "真实数量", adDouble, 18, adFldIsNullable
        .Fields.Append "药品编码和名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "分批属性", adLongVarChar, 40, adFldIsNullable
    
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub AddAppend(ByVal Index As Integer)
    '往动态纪录集增加值
    Dim i As Integer
    On Error GoTo ErrHand

    With mrsMyAppend
        For i = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Val(mshBill.TextMatrix(i, mconIntCol实际数量)) <> 0 Then
                .AddNew
                .Fields!药品id = mshBill.TextMatrix(i, 0)
                .Fields!药名 = mshBill.TextMatrix(i, mconIntCol药名)
                .Fields!商品名 = mshBill.TextMatrix(i, mconIntCol商品名)
                .Fields!来源 = mshBill.TextMatrix(i, mconIntCol来源)
                .Fields!基本药物 = mshBill.TextMatrix(i, mconIntCol基本药物)
                .Fields!规格 = mshBill.TextMatrix(i, mconIntCol规格)
                .Fields!分批核算 = mshBill.TextMatrix(i, mconIntCol分批核算)
                .Fields!最大效期 = mshBill.TextMatrix(i, mconIntCol最大效期)
                .Fields!可用数量 = mshBill.TextMatrix(i, mconIntCol可用数量)
                .Fields!指导差价率 = mshBill.TextMatrix(i, mconIntCol指导差价率)
                .Fields!实际金额 = mshBill.TextMatrix(i, mconIntCol实际金额)
                .Fields!实际差价 = mshBill.TextMatrix(i, mconIntCol实际差价)
                .Fields!比例系数 = mshBill.TextMatrix(i, mconIntCol比例系数)
                .Fields!批次 = mshBill.TextMatrix(i, mconIntCol批次)
                .Fields!产地 = mshBill.TextMatrix(i, mconIntCol产地)
                .Fields!单位 = mshBill.TextMatrix(i, mconIntCol单位)
                .Fields!送货单位 = mshBill.TextMatrix(i, mconIntCol送货单位)
                .Fields!批号 = mshBill.TextMatrix(i, mconIntCol批号)
                .Fields!效期 = mshBill.TextMatrix(i, mconIntCol效期)
                .Fields!批准文号 = mshBill.TextMatrix(i, mconIntCol批准文号)
                .Fields!库房库存 = mshBill.TextMatrix(i, mconIntCol库房库存)
                .Fields!对方库存 = mshBill.TextMatrix(i, mconIntCol对方库存)
                .Fields!填写数量 = IIf(mshBill.TextMatrix(i, mconIntCol填写数量) = "", 0, mshBill.TextMatrix(i, mconIntCol填写数量))
                .Fields!实际数量 = IIf(mshBill.TextMatrix(i, mconIntCol实际数量) = "", 0, mshBill.TextMatrix(i, mconIntCol实际数量))
                .Fields!采购价 = mshBill.TextMatrix(i, mconIntCol采购价)
                .Fields!采购金额 = IIf(mshBill.TextMatrix(i, mconIntCol采购金额) = "", 0, mshBill.TextMatrix(i, mconIntCol采购金额))
                .Fields!售价 = mshBill.TextMatrix(i, mconIntCol售价)
                .Fields!售价金额 = IIf(mshBill.TextMatrix(i, mconIntCol售价金额) = "", 0, mshBill.TextMatrix(i, mconIntCol售价金额))
                .Fields!差价 = IIf(mshBill.TextMatrix(i, mconintCol差价) = "", 0, mshBill.TextMatrix(i, mconintCol差价))
                .Fields!上次供应商ID = mshBill.TextMatrix(i, mconIntCol上次供应商ID)
                .Fields!真实数量 = IIf(mshBill.TextMatrix(i, mconintCol真实数量) = "", 0, mshBill.TextMatrix(i, mconintCol真实数量))
                .Fields!药品编码和名称 = mshBill.TextMatrix(i, mconIntCol药品编码和名称)
                .Fields!药品编码 = mshBill.TextMatrix(i, mconIntCol药品编码)
                .Fields!药品名称 = mshBill.TextMatrix(i, mconIntCol药品名称)
                .Fields!分批属性 = mshBill.TextMatrix(i, mconIntCol分批属性)
                .Update
            End If
        Next
    
        For i = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Index = 0 And Val(mshBill.TextMatrix(i, mconIntCol实际数量)) = 0 Then
                .AddNew
                .Fields!药品id = mshBill.TextMatrix(i, 0)
                .Fields!药名 = mshBill.TextMatrix(i, mconIntCol药名)
                .Fields!商品名 = mshBill.TextMatrix(i, mconIntCol商品名)
                .Fields!来源 = mshBill.TextMatrix(i, mconIntCol来源)
                .Fields!基本药物 = mshBill.TextMatrix(i, mconIntCol基本药物)
                .Fields!规格 = mshBill.TextMatrix(i, mconIntCol规格)
                .Fields!分批核算 = mshBill.TextMatrix(i, mconIntCol分批核算)
                .Fields!最大效期 = mshBill.TextMatrix(i, mconIntCol最大效期)
                .Fields!可用数量 = mshBill.TextMatrix(i, mconIntCol可用数量)
                .Fields!指导差价率 = mshBill.TextMatrix(i, mconIntCol指导差价率)
                .Fields!实际金额 = mshBill.TextMatrix(i, mconIntCol实际金额)
                .Fields!实际差价 = mshBill.TextMatrix(i, mconIntCol实际差价)
                .Fields!比例系数 = mshBill.TextMatrix(i, mconIntCol比例系数)
                .Fields!批次 = mshBill.TextMatrix(i, mconIntCol批次)
                .Fields!产地 = mshBill.TextMatrix(i, mconIntCol产地)
                .Fields!单位 = mshBill.TextMatrix(i, mconIntCol单位)
                .Fields!送货单位 = mshBill.TextMatrix(i, mconIntCol送货单位)
                .Fields!批号 = mshBill.TextMatrix(i, mconIntCol批号)
                .Fields!效期 = mshBill.TextMatrix(i, mconIntCol效期)
                .Fields!批准文号 = mshBill.TextMatrix(i, mconIntCol批准文号)
                .Fields!库房库存 = mshBill.TextMatrix(i, mconIntCol库房库存)
                .Fields!对方库存 = mshBill.TextMatrix(i, mconIntCol对方库存)
                .Fields!填写数量 = IIf(mshBill.TextMatrix(i, mconIntCol填写数量) = "", 0, mshBill.TextMatrix(i, mconIntCol填写数量))
                .Fields!实际数量 = IIf(mshBill.TextMatrix(i, mconIntCol实际数量) = "", 0, mshBill.TextMatrix(i, mconIntCol实际数量))
                .Fields!采购价 = mshBill.TextMatrix(i, mconIntCol采购价)
                .Fields!采购金额 = IIf(mshBill.TextMatrix(i, mconIntCol采购金额) = "", 0, mshBill.TextMatrix(i, mconIntCol采购金额))
                .Fields!售价 = mshBill.TextMatrix(i, mconIntCol售价)
                .Fields!售价金额 = IIf(mshBill.TextMatrix(i, mconIntCol售价金额) = "", 0, mshBill.TextMatrix(i, mconIntCol售价金额))
                .Fields!差价 = IIf(mshBill.TextMatrix(i, mconintCol差价) = "", 0, mshBill.TextMatrix(i, mconintCol差价))
                .Fields!上次供应商ID = mshBill.TextMatrix(i, mconIntCol上次供应商ID)
                .Fields!真实数量 = IIf(mshBill.TextMatrix(i, mconintCol真实数量) = "", 0, mshBill.TextMatrix(i, mconintCol真实数量))
                .Fields!药品编码和名称 = mshBill.TextMatrix(i, mconIntCol药品编码和名称)
                .Fields!药品编码 = mshBill.TextMatrix(i, mconIntCol药品编码)
                .Fields!药品名称 = mshBill.TextMatrix(i, mconIntCol药品名称)
                .Fields!分批属性 = mshBill.TextMatrix(i, mconIntCol分批属性)
                .Update
            End If
        Next
    End With
       
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If mint编辑状态 = 10 Then
        Cancel = True
        Exit Sub
    End If
    If InStr(1, "34", mint编辑状态) <> 0 Then
        If mint编辑状态 = 3 And mbln申领单 Then Exit Sub
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim i As Integer
    Dim intRowvalid As Integer  '记录有效的行数
    Dim RecReturn As Recordset
    Dim rsMaterial As New ADODB.Recordset
    Dim intCheckAll As Integer
    Dim blnReturn As Boolean    '用来判断结果集中是否是多选数据
    Dim intRow As Integer       '当前行
    Dim str药品ID As String     '有哪些是重复的药品id
    Dim rsTemp As ADODB.Recordset '临时记录过滤重复值后的数据集
    Dim lng药品ID As Long
    Dim strTemp As String
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    If cboEnterStock.ListCount = 0 Then Exit Sub
    intOldRow = mshBill.Row
    intRow = mshBill.Row
    Select Case mshBill.Col
    Case mconIntCol药名
        mshBill.CmdEnable = False
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房)
        End If
        If Not mbln申领单 Then
'            Set RecReturn = Frm药品选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, _
'                True, True, False, False, True)
            Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, , True, True, True, , , mstrPrivs)
        Else    '申领单
'            Set RecReturn = Frm药品选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, _
'                mbln明确批次, Not mbln明确批次, False, False, True)
            Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, , True, True, True, , , mstrPrivs)
        End If
        mshBill.CmdEnable = True
        
        If RecReturn.RecordCount > 0 Then
            Set RecReturn = CheckData(RecReturn) '检查重复记录和时价无库存的记录，并将满足条件的记录过滤掉
        End If
        
        If RecReturn.RecordCount > 0 Then
            With mshBill
                Dim lngCurRow As Long
                
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intRow = .Row
                    If IsSelf_Command(RecReturn!药品id) Then
                        '提取该自制药品的组成药品，并依次产生数据
                        Set rsMaterial = GetMaterial(RecReturn!药品id)
                        
                        If rsMaterial.RecordCount > 0 Then
                            Set rsMaterial = CheckData(rsMaterial) '检查重复记录和时价无库存的记录，并将满足条件的记录过滤掉
                        End If
                        
                        If rsMaterial.RecordCount <> 0 Then '如果有数据，将数据移动到第一条记录
                            rsMaterial.MoveFirst
                        End If
                        lngCurRow = mshBill.Row
                        mshBill.rows = mshBill.rows + rsMaterial.RecordCount
                        mshBill.Row = lngCurRow
                        With rsMaterial
                            Do While Not .EOF
                                mshBill.TextMatrix(mshBill.Row, mconIntCol行号) = mshBill.Row
                                SetColValue mshBill.Row, !药品id, "[" & !药品编码 & "]", !通用名, IIf(IsNull(!商品名), "", !商品名), _
                                    Nvl(!药品来源), "" & !基本药物, _
                                    IIf(IsNull(!规格), "", !规格), IIf(IsNull(!产地), "", !产地), _
                                    Choose(mintUnit, !售价单位, !门诊单位, !住院单位, !药库单位), _
                                    !售价, IIf(IsNull(!批号), "", !批号), _
                                    IIf(IsNull(!效期), "", Format(!效期, "yyyy-MM-dd")), _
                                    IIf(IsNull(!最大效期), "0", !最大效期), _
                                    !药库分批, _
                                    IIf(IsNull(!可用数量), "0", !可用数量), _
                                    IIf(IsNull(!实际金额), "0", !实际金额), _
                                    IIf(IsNull(!实际差价), "0", !实际差价), _
                                    IIf(IsNull(!指导差价率), "0", !指导差价率), _
                                    Choose(mintUnit, 1, !门诊包装, !住院包装, !药库包装), _
                                    IIf(IsNull(!批次), 0, !批次), !时价, !药房分批, !上次供应商ID, _
                                    IIf(IsNull(!批准文号), "", !批准文号)
                                .MoveNext
                                mshBill.Row = mshBill.Row + 1
                            Loop
                        End With
'                        mshBill.Row = lngCurRow
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol行号) = .Row
                        SetColValue .Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                            Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                            IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                            Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                            RecReturn!售价, IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                            IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                            RecReturn!药库分批, _
                            IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                            IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                            IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                            IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                            Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                            IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                            RecReturn!上次供应商ID, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
                    End If
                    
                    .Col = mconIntCol填写数量

                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End With
            RecReturn.Close
        Else
            mshBill.Row = intOldRow
        End If
    Case mconIntCol批号
        gstrSQL = "Select Distinct 上次批号,上次产地,批准文号,上次供应商ID From 药品库存 Where 性质=1 And 库房id=[1] And 药品id=[2] "
        Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "取批号信息", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0))
        If RecReturn.RecordCount = 0 Then
            MsgBox "没有找到该药品的批号信息，请手工输入批号。"
            Exit Sub
        End If
        
        Set msh批次信息.Recordset = RecReturn
        With msh批次信息
            .Redraw = False
            .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
            .Visible = True
            .SetFocus
            .ColWidth(0) = 800
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 0
            .Row = 1
            .Col = 0
            .TopRow = 1
            .ColSel = .Cols - 1
            .Redraw = True
            Exit Sub
        End With
    Case mconIntCol产地
        Dim rsProvider As New Recordset
        
        gstrSQL = "Select 编码,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "药品生产商", gstrNodeNo)
        
        If rsProvider.EOF Then
            rsProvider.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rsProvider
            .StrNode = "所有药品生产商"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = .CurrentName
                mshBill.Col = mconIntCol批号
            End If
        End With
        Unload FrmSelect
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
'        If strkey = "" Then
'            strkey = .TextMatrix(.Row, .Col)
'        End If
        Select Case .Col
            Case mconIntCol填写数量, mconIntCol实际数量
                intDigit = mintNumberDigit
        End Select
        
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If .SelLength = Len(strkey) Then Exit Sub
            If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnterCell Then Exit Sub
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        If .Row <> .LastRow Then
            SetInputFormat .Row
        End If
        
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                
            Case mconIntCol批号
                .txtCheck = False
'                .TextMask = "1234567890"
                .MaxLength = mintBatchNoLen
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                    .ColData(mconIntCol批号) = 4
                End If
            Case mconIntCol效期
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol效期) = "" And .TextMatrix(.Row, mconIntCol批号) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) And .TextMatrix(.Row, mconIntCol最大效期) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol最大效期), "||")(0) <> 0 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol最大效期), "||")(0), strxq), "yyyy-mm-dd")
                                If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                                    '换算为有效期
                                    .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                                End If
                            End If
                        End If
                    End If
                End If
            Case mconIntCol填写数量, mconIntCol实际数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
            Case mconIntCol产地
                If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                    .ColData(mconIntCol产地) = 1
                End If
                
                OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
                .txtCheck = False
                .MaxLength = 30
                .TxtSetFocus
            Case mconIntCol批准文号
                .txtCheck = False
                .MaxLength = 40
        End Select
        
    End With
End Sub

Private Sub mshBill_GotFocus()
    If mintParallelRecord <> 1 Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
        MsgBox "对不起，移入库房和移出库房相同了，请检查后重新选择！", vbOKOnly + vbExclamation, gstrSysName
        cboEnterStock.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim str药品ID As String
    Dim i As Integer
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        
        intOldRow = .Row
        intRow = .Row
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntCol药名
                If strkey <> "" Then
                    Dim lngCurRow As Long
                    Dim rsMaterial As New ADODB.Recordset

                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房)
                    End If
'                    If Not mbln申领单 Then
'                        Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, _
'                            strkey, sngLeft, sngTop, True, True, False, False, True)
'                    Else
'                        Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, _
'                            strkey, sngLeft, sngTop, mbln明确批次, Not mbln明确批次, False, False, True)
'                    End If
                    
                    If mbln申领单 Then
                        Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, , True, True, True, , , mstrPrivs)
                    Else
                        Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng出库库房, , True, True, True, , , mstrPrivs)
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn) '检查重复记录和时价无库存的记录，并将满足条件的记录过滤掉
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            intRow = .Row
                            If IsSelf_Command(RecReturn!药品id) Then
                                '提取该自制药品的组成药品，并依次产生数据
                                Set rsMaterial = GetMaterial(RecReturn!药品id)
                                
                                If rsMaterial.RecordCount > 0 Then
                                    Set rsMaterial = CheckData(rsMaterial) '检查重复记录和时价无库存的记录，并将满足条件的记录过滤掉
                                End If
                                
                                If rsMaterial.RecordCount <> 0 Then '如果有数据，将数据移动到第一条记录
                                    rsMaterial.MoveFirst
                                End If
                                
                                lngCurRow = mshBill.Row
                                mshBill.rows = mshBill.rows + rsMaterial.RecordCount
                                mshBill.Row = lngCurRow
                                With rsMaterial
                                    Do While Not .EOF
                                        mshBill.TextMatrix(mshBill.Row, mconIntCol行号) = mshBill.Row
                                        SetColValue mshBill.Row, !药品id, "[" & !药品编码 & "]", !通用名, IIf(IsNull(!商品名), "", !商品名), _
                                            Nvl(!药品来源), "" & !基本药物, _
                                            IIf(IsNull(!规格), "", !规格), IIf(IsNull(!产地), "", !产地), _
                                            Choose(mintUnit, !售价单位, !门诊单位, !住院单位, !药库单位), _
                                            !售价, IIf(IsNull(!批号), "", !批号), _
                                            IIf(IsNull(!效期), "", Format(!效期, "yyyy-MM-dd")), _
                                            IIf(IsNull(!最大效期), "0", !最大效期), _
                                            !药库分批, _
                                            IIf(IsNull(!可用数量), "0", !可用数量), _
                                            IIf(IsNull(!实际金额), "0", !实际金额), _
                                            IIf(IsNull(!实际差价), "0", !实际差价), _
                                            IIf(IsNull(!指导差价率), "0", !指导差价率), _
                                            Choose(mintUnit, 1, !门诊包装, !住院包装, !药库包装), _
                                            IIf(IsNull(!批次), 0, !批次), !时价, !药房分批, !上次供应商ID, _
                                            IIf(IsNull(!批准文号), "", !批准文号)
                                        .MoveNext
                                        mshBill.Row = mshBill.Row + 1
                                    Loop
                                End With
'                                mshBill.Row = lngCurRow
                            Else
                                mshBill.TextMatrix(mshBill.Row, mconIntCol行号) = .Row
                                If SetColValue(.Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                        Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                                        IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                        Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                        IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                        IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                        IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                                        RecReturn!药库分批, _
                                        IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                        IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                        IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                        IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                                        Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                                        IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                                        RecReturn!上次供应商ID, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) = False Then
                                    Cancel = True
                                    Exit Sub
                                End If
                                
                                .Text = .TextMatrix(.Row, .Col)
                            End If
                            
                            If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                        Cancel = True
                    End If
                End If
            Case mconIntCol批号
                '无处理
                If strkey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol批号) = ""
                    End If
                    If .ColData(mconIntCol效期) = 2 Then
                        .Col = mconIntCol效期
                    Else
                        .Col = mconIntCol填写数量
                    End If
                    Cancel = True
                    Exit Sub
                Else
                    gstrSQL = "Select Distinct 上次批号,上次产地,批准文号,上次供应商ID From 药品库存 Where 性质=1 And 库房id=[1] And 药品id=[2] And 上次批号 like [3] "
                    Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "取批号信息", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0), IIf(gstrMatchMethod = "0", "%", "") & strkey & "%")
                    If RecReturn.RecordCount = 0 Then
                        If .ColData(mconIntCol效期) = 2 Then
                            .Col = mconIntCol效期
                        Else
                            .Col = mconIntCol填写数量
                        End If
                        .TextMatrix(.Row, mconIntCol批号) = strkey
                        Cancel = True
                        Exit Sub
                    ElseIf RecReturn.RecordCount = 1 Then
                        .TextMatrix(.Row, mconIntCol批号) = Nvl(RecReturn.Fields("上次批号"), "")
                        .Text = Nvl(RecReturn.Fields("上次批号"), "")
                        .TextMatrix(.Row, mconIntCol产地) = Nvl(RecReturn.Fields("上次产地"), "")
                        .TextMatrix(.Row, mconIntCol批准文号) = Nvl(RecReturn.Fields("批准文号"), "")
                        .TextMatrix(.Row, mconIntCol上次供应商ID) = Nvl(RecReturn.Fields("上次供应商ID"), 0)
                    Else
                        Set msh批次信息.Recordset = RecReturn
                        With msh批次信息
                            .Redraw = False
                            .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                            .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                            .Visible = True
                            .SetFocus
                            .ColWidth(0) = 800
                            .ColWidth(1) = 1000
                            .ColWidth(2) = 1000
                            .ColWidth(3) = 0
                            .Row = 1
                            .Col = 0
                            .TopRow = 1
                            .ColSel = .Cols - 1
                            .Redraw = True
                            Cancel = True
                            Exit Sub
                        End With
                    End If
                End If
            Case mconIntCol效期
                '有处理
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            MsgBox "对不起，失效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strkey
                        Exit Sub
                    End If
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，失效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strkey = "" And strkey <> .TextMatrix(.Row, mconIntCol效期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            
            Case mconIntCol填写数量, mconIntCol实际数量
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, .Col) = "" And strkey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
'                If .TextMatrix(.Row, .Col) <> "" And strkey = "" Then
'                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
'                    Cancel = True
'                    .TxtSetFocus
'                    Exit Sub
'                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) <= 0 And Not (mint编辑状态 = 3 Or mint编辑状态 = 6 Or mint编辑状态 = 10) Then
                        MsgBox "对不起，数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint编辑状态 = 6 Then
                        If Not 相同符号(Val(strkey), Val(.TextMatrix(.Row, mconIntCol填写数量))) Then
                            MsgBox "对不起，冲销数量的符号应该与原有数量一致！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strkey) >= 0 Then
                            If Val(strkey) > Val(.TextMatrix(.Row, mconIntCol填写数量)) Then
                                MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If Val(strkey) < Val(.TextMatrix(.Row, mconIntCol填写数量)) Then
                                MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If gtype_UserSysParms.P174_药品移库明确批次 = 1 Then
                        If Not CompareUsableQuantity(.Row, strkey) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    .TextMatrix(.Row, .Col) = .Text
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol售价) * strkey, mintMoneyDigit)
                    End If
                    
'                    .TextMatrix(.Row, mconintCol差价) = FormatEx(Get出库差价(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(.TextMatrix(.Row, mconIntCol实际金额)), Val(.TextMatrix(.Row, mconIntCol实际差价)), Val(.TextMatrix(.Row, mconIntCol售价金额)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol比例系数))), mintMoneyDigit)
                        
                    If strkey <> 0 And (mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 3) Then
                        .TextMatrix(.Row, mconIntCol采购价) = GetFormat(Get成本价(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol批次))) * Val(mshBill.TextMatrix(.Row, mconIntCol比例系数)), mintCostDigit)
                    End If
                    .TextMatrix(.Row, mconIntCol采购金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol采购价)) * strkey, mintMoneyDigit)
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价金额)) - .TextMatrix(.Row, mconIntCol采购金额), mintMoneyDigit)
                    
                    If .Col = mconIntCol填写数量 Then
                        .TextMatrix(.Row, mconIntCol实际数量) = strkey
                    End If
                End If
                显示合计金额
                If mbln申领单 Then Call ShowColor(mshBill.Row)
            Case mconIntCol产地
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol产地) = ""
                    End If
                    If .ColData(mconIntCol批号) = 5 Then
                        .Col = mconIntCol填写数量
                    Else
                        .Col = mconIntCol批号
                    End If
                    Cancel = True
                            
                    Exit Sub
                Else
                    Dim rs产地 As New Recordset
                    
                    gstrSQL = "Select 编码,简码,名称 From 药品生产商 " _
                            & "Where (站点 = [4] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [2] or Upper(简码) like [3]) Order By 编码"
                    Set rs产地 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        gstrNodeNo)
                    
                    If rs产地.RecordCount = 0 Then
                        Exit Sub
                    ElseIf rs产地.RecordCount = 1 Then
                        .TextMatrix(.Row, mconIntCol产地) = rs产地.Fields("名称")
                        .Text = rs产地.Fields("名称")
                    Else
                        Set msh产地.Recordset = rs产地
                        With msh产地
                            .Redraw = False
                            .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                            .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                            .Visible = True
                            .SetFocus
                            .ColWidth(0) = 800
                            .ColWidth(1) = 800
                            .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                            .Row = 1
                            .Col = 0
                            .TopRow = 1
                            .ColSel = .Cols - 1
                            .Redraw = True
                            Cancel = True
                            Exit Sub
                        End With
                    End If
                End If
                OpenIme
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, ByVal str基本药物 As String, _
    ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal int分批核算 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal lng上次供应商ID As Long, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim dblCost As Double
    Dim str药名 As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double, dbl成本价 As Double
    
    SetColValue = False
    
    On Error GoTo errHandle
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, 0) = lng药品ID
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(intRow, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(intRow, mconIntCol药品编码) = str药品编码
        .TextMatrix(intRow, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
        Else
            .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(intRow, mconIntCol商品名) = str商品名
        
        .TextMatrix(intRow, mconIntCol来源) = str药品来源
        .TextMatrix(intRow, mconIntCol基本药物) = str基本药物
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol单位) = str单位
        
        '加载大包装入库信息
        gstrSQL = "select a.送货单位,a.送货包装 from 药品规格 a where a.药品id=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "入库信息", lng药品ID)
        If num比例系数 = 0 Or Nvl(rsTemp!送货包装) = "" Or Nvl(rsTemp!送货单位) = "" Then
            .TextMatrix(intRow, mconIntCol送货单位) = ""
        Else
            .TextMatrix(intRow, mconIntCol送货单位) = rsTemp!送货单位 & "(1" & rsTemp!送货单位 & "=" & zlStr.FormatEx(rsTemp!送货包装 / num比例系数, 1, , True) & str单位 & ")"
        End If
        
        If gtype_UserSysParms.P174_药品移库明确批次 = 0 Then
            .TextMatrix(intRow, mconIntCol批次) = 0
        Else
            .TextMatrix(intRow, mconIntCol批次) = lng批次
        End If
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, mconIntCol售价) = GetFormat(num售价 * num比例系数, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol分批核算) = int分批核算
        .TextMatrix(intRow, mconIntCol可用数量) = GetFormat(num可用数量, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol最大效期) = int最大效期 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntCol指导差价率) = num指导差价率
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        
        .TextMatrix(intRow, mconIntCol上次供应商ID) = lng上次供应商ID
        .TextMatrix(intRow, mconIntCol分批属性) = IIf(Get药品分批属性(lng药品ID, cboEnterStock.ItemData(cboEnterStock.ListIndex)) = True, "1", 0)
                
        If int是否变价 = 1 Then
            .TextMatrix(intRow, mconIntCol售价) = GetFormat(Get售价(True, lng药品ID, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol批次))) * num比例系数, mintPriceDigit)
'            If mbln明确批次 = True Then
'                gstrSQL = "Select Decode(Nvl(批次, 0), 0, 实际金额 / 实际数量, Nvl(零售价, 实际金额 / 实际数量))*" & num比例系数 & " as  售价 " _
'                    & "  from 药品库存 " _
'                    & " where 库房id=[1] " _
'                    & " and 药品id=[2] " _
'                    & " and 性质=1 and 实际数量>0 and " _
'                    & " nvl(批次,0)=[3] "
'                Set rsRecord = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)
'
'                If Not rsRecord.EOF Then
'                    dblPrice = rsRecord.Fields(0).Value
'                Else
'                    dblPrice = 0
'                End If
'                .TextMatrix(intRow, mconIntCol售价) = GetFormat(dblPrice, mintPriceDigit)
'            End If
        End If
        
        If IsLowerLimit(mlng出库库房, lng药品ID) Then Call SetForeColor_ROW(mlng紫色)
        Call CheckLapse(str效期)
        SetInputFormat intRow
        
        Call 提示库存数(intRow)
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim bln库房 As Boolean
    Dim bln药库分批 As Boolean, bln药房分批 As Boolean
    Dim rsTemp As ADODB.Recordset
    
    '说明：只根据入库库房进行判断
    '   1、入库库房且药库分批，则允许输入批次信息
    '   2、入库药房且药房分批，则允许输入批次信息
    
    On Error GoTo errHandle
    If mblnEdit = False Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If Val(mshBill.TextMatrix(mshBill.Row, 0)) = 0 Then Exit Sub
    bln药库分批 = (mshBill.TextMatrix(mshBill.Row, mconIntCol分批核算) = 1)
    bln药房分批 = (Split(mshBill.TextMatrix(mshBill.Row, mconIntCol最大效期), "||")(2) = 1)
    bln库房 = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    With mshBill
        If .TextMatrix(intRow, mconIntCol批号) = "" And _
            ((bln库房 And bln药库分批) Or (Not bln库房 And bln药房分批)) Then
            .ColData(mconIntCol批号) = 4              '纯文本输入
            .ColData(mconIntCol批号) = 1               '按钮
            If .TextMatrix(intRow, mconIntCol最大效期) <> "" Then
                If Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(0) <> 0 Then
                    .ColData(mconIntCol效期) = 2          '日期输入框
                Else
                    .ColData(mconIntCol效期) = 5
                End If
            Else
                .ColData(mconIntCol效期) = 5
            End If
        ElseIf bln库房 And bln药库分批 And Not bln药房分批 Then '药房向药库移库，药房不分批且药库分批
            gstrSQL = "Select 库房id From 药品库存 Where 库房id=[1] And 药品id=[2] And Rownum=1 "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房有无库存", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0))
            
            If rsTemp.RecordCount = 0 Then
                .ColData(mconIntCol批号) = 4
                .ColData(mconIntCol产地) = 1
                .ColData(mconIntCol批准文号) = 4
            Else
                .ColData(mconIntCol批号) = 1
                If .TextMatrix(intRow, mconIntCol最大效期) <> "" Then
                    If Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(0) <> 0 Then
                        .ColData(mconIntCol效期) = 2          '日期输入框
                    Else
                        .ColData(mconIntCol效期) = 5
                    End If
                Else
                    .ColData(mconIntCol效期) = 5
                End If
            End If
        Else
            .ColData(mconIntCol批号) = 5              '禁止
            .ColData(mconIntCol效期) = 5
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub


Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh产地.TextMatrix(msh产地.Row, 2)
            msh产地.Visible = False
            .Col = mconIntCol批号
            .SetFocus
        End If
    
    End With
End Sub


Private Sub msh产地_LostFocus()
    If msh产地.Visible Then
        msh产地.Visible = False
    End If
End Sub


Private Sub msh批次信息_DblClick()
    msh批次信息_KeyDown vbKeyReturn, 0
End Sub


Private Sub msh批次信息_KeyDown(KeyCode As Integer, Shift As Integer)
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh批次信息.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, mconIntCol批号) = msh批次信息.TextMatrix(msh批次信息.Row, 0)
            .TextMatrix(.Row, mconIntCol产地) = msh批次信息.TextMatrix(msh批次信息.Row, 1)
            .TextMatrix(.Row, mconIntCol批准文号) = msh批次信息.TextMatrix(msh批次信息.Row, 2)
            .TextMatrix(.Row, mconIntCol上次供应商ID) = Val(msh批次信息.TextMatrix(msh批次信息.Row, 3))
            msh批次信息.Visible = False
            .Col = mconIntCol填写数量
            .SetFocus
        End If
    
    End With
End Sub


Private Sub msh批次信息_LostFocus()
    If msh批次信息.Visible Then
        msh批次信息.Visible = False
    End If
End Sub


Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    Dim bln入库库房 As Boolean, bln出库库房 As Boolean
    Dim bln药库分批 As Boolean, bln药房分批 As Boolean
    ValidData = False
    bln入库库房 = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    bln出库库房 = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If cboEnterStock.ListCount = 0 Then
                MsgBox "请设置允许调拨的部门，[基础参数设置]中的药品流向！", vbInformation, gstrSysName
                Exit Function
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "对不起，移入库房和移出库房相同了，请重新选择！", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Function
            End If
            
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol填写数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol填写数量
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol批号))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "第" & intLop & "行药品的批号超长,最多能输入" & Int(mintBatchNoLen / 2) & "个汉字或" & mintBatchNoLen & "个字符!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol批号
                        Exit Function
                    End If
                    
                    
'                    '说明：只根据入库库房进行判断
'                    '   1、入库库房且药库分批，则允许输入批次信息
'                    '   2、入库药房且药房分批，则允许输入批次信息
'                    bln药库分批 = (mshBill.TextMatrix(intLop, mconIntCol分批核算) = 1)
'                    bln药房分批 = (Split(mshBill.TextMatrix(intLop, mconIntCol最大效期), "||")(2) = 1)
'                    If ((bln入库库房 And bln药库分批) Or (Not bln入库库房 And bln药房分批)) And Val(mshBill.TextMatrix(intLop, mconIntCol实际数量)) <> 0 Then
'                        If Split(.TextMatrix(intLop, mconIntCol最大效期), "||")(0) <> 0 Then
'                            If .TextMatrix(intLop, mconIntCol批号) = "" Or .TextMatrix(intLop, mconIntCol效期) = "" Then
'                                MsgBox "第" & intLop & "行的药品是效期药品,请把它的批号及失效期完整输入单据中！", vbInformation, gstrSysName
'                                mshBill.SetFocus
'                                .Row = intLop
'                                .MsfObj.TopRow = intLop
'                                If .TextMatrix(intLop, mconIntCol批号) = "" Then
'                                    .Col = mconIntCol批号
'                                Else
'                                    .Col = mconIntCol效期
'                                End If
'                                Exit Function
'                            End If
'                        End If
'                    End If
                    '只有申领才可能产生如此记录
                    '   3、出库库房且药库分批或出库药房且药房分批，如果批次小于等于零，说明该批次药品无库存，不允许发送（允许保存）
                    If mint编辑状态 <> 2 Then
                        If ((bln出库库房 And bln药库分批) Or (Not bln出库库房 And bln药房分批)) Then
                            If Val(.TextMatrix(intLop, mconIntCol批次)) = 0 And Val(.TextMatrix(intLop, mconIntCol实际数量)) <> 0 Then
                                MsgBox "第" & intLop & "行的药品是批次药品且无库存，不允许发送！", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .Col = mconIntCol实际数量
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol填写数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的填写数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol填写数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol实际数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的实际数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol实际数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol采购金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol填写数量) = 4, mconIntCol填写数量, mconIntCol实际数量)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol填写数量) = 4, mconIntCol填写数量, mconIntCol实际数量)
                        Exit Function
                    End If
                    
                    If .TextMatrix(intLop, mconIntCol分批属性) = "1" And (.TextMatrix(intLop, mconIntCol批号) = "" Or .TextMatrix(intLop, mconIntCol产地) = "") And .TextMatrix(intLop, 0) <> "" Then
                        MsgBox "第" & intLop & "行，入库库房是分批管理，必须录入批号和产地！", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol批号) = "" Then
                            .Col = mconIntCol批号
                        Else
                            .Col = mconIntCol产地
                        End If
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function



Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockID As Long
    Dim lngEnterStockID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblRealNum As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim blnTran As Boolean
    Dim lng上次供应商ID As Long
    Dim strCheckString As String
    Dim str批准文号 As String
    Dim n As Integer
    
    arrSql = Array()
    SaveCard = False
    
    '检查该单据是否在进入编辑界面后，被其他操作员修改；从入库转入移库的单据不处理
    If mint编辑状态 = 2 Or (bln强制保存 And mint编辑状态 <> 11) Then        '修改
        mstrTime_End = GetBillInfo(6, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Function
        End If
        strCheckString = CheckBill(mstr单据号)
        If strCheckString <> "" Then
            MsgBox strCheckString, vbInformation, gstrSysName
            Exit Function
        End If
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error GoTo errHandle
     
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = zlDataBase.GetNextNo(26, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lngEnterStockID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        If Txt填制日期.Caption = "" Or Not IsDate(Txt填制日期.Caption) Then
            datBookDate = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Else
            datBookDate = Format(Txt填制日期.Caption, "yyyy-mm-dd hh:mm:ss")
        End If
        strAssessor = Txt审核人
        
        'Modified by ZYB 2004-05-16 昆明处理：不是强制保存则开始事务
        If bln强制保存 Then blnTran = True
        
        '从入库转入移库的单据不处理
        If mint编辑状态 = 2 Or (bln强制保存 And mint编辑状态 <> 11) Then        '修改
            If Not mbln申领单 Then
                gstrSQL = "zl_药品移库_Delete('" & mstr单据号 & "')"
            Else
                gstrSQL = "zl_药品申领_Delete('" & mstr单据号 & "')"
            End If
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
        End If
        
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                lngBatchID = .TextMatrix(intRow, mconIntCol批次)
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", .TextMatrix(intRow, mconIntCol效期))
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = GetFormat(Val(.TextMatrix(intRow, mconIntCol填写数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                dblRealNum = GetFormat(Val(.TextMatrix(intRow, mconIntCol实际数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                
                If Val(.TextMatrix(intRow, mconintCol真实数量)) <> 0 Then
                    If Val(GetFormat(Val(.TextMatrix(intRow, mconintCol真实数量)) / Val(.TextMatrix(intRow, mconIntCol比例系数)), mintNumberDigit)) = Val(.TextMatrix(intRow, mconIntCol填写数量)) Then
                        If dblQuantity = dblRealNum Then
                            dblQuantity = Val(.TextMatrix(intRow, mconintCol真实数量))
                            dblRealNum = Val(.TextMatrix(intRow, mconintCol真实数量))
                        End If
                    End If
                End If
                
'                dblPurchasePrice = FormatEx(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchasePrice = Get成本价(lngDrugID, lngStockID, lngBatchID)
                dblPurchaseMoney = Val(.TextMatrix(intRow, mconIntCol采购金额))
'                dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSalePrice = Get售价(Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1, lngDrugID, lngStockID, lngBatchID)
                dblSaleMoney = Val(.TextMatrix(intRow, mconIntCol售价金额))
                dblMistakePrice = Val(.TextMatrix(intRow, mconintCol差价))
                lng上次供应商ID = .TextMatrix(intRow, mconIntCol上次供应商ID)
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lngSerial = 2 * intRow - 1
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                lngSerial = 2 * intRow - 1
                .TextMatrix(intRow, mconIntCol序号) = lngSerial
                
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))

                If Not mbln申领单 Or bln强制保存 Then
                    gstrSQL = "zl_药品移库_INSERT("
                Else
                    gstrSQL = "zl_药品申领_INSERT("
                End If
                
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockID
                '对方部门ID
                gstrSQL = gstrSQL & "," & lngEnterStockID
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '批次
                gstrSQL = gstrSQL & "," & lngBatchID
                '填写数量
                gstrSQL = gstrSQL & "," & dblQuantity
                '实际数量
                gstrSQL = gstrSQL & "," & dblRealNum
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '零售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '零售金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '供药单位ID
                gstrSQL = gstrSQL & "," & IIf(lng上次供应商ID = 0, "NULL", lng上次供应商ID)
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '申领方式
                gstrSQL = gstrSQL & "," & IIf(mintApplyType = -1, "Null", mintApplyType)
                '结束时间
                gstrSQL = gstrSQL & ",'" & mstrEndTime & "'"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSql, MStrCaption, False, Not bln强制保存) Then Exit Function
        
        If Not bln强制保存 Then gcnOracle.CommitTrans: blnTran = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get药品分批属性(ByVal lng药品ID As Long, ByVal lng库房ID As Long) As Boolean
'功能：根据传过来的药品id和库房id判断该药品所在库房是否分批
'参数 lng药品id：需要检查的药品
'参数 lng库房id：需要检查的库房
'返回值： true-被检测药品所在库房分批；false-被检测药品所在库房不分批
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int分批属性 As Integer      '0-不分批;1-分批
    Dim int药库分批 As Integer      '0-不分批;1-分批
    Dim int药房分批 As Integer      '0-不分批;1-分批
    Dim bln是否具有药房性质 As Boolean  'True-具有药房性质;False-不具有药房性质
    
    If lng药品ID = 0 Then Exit Function
    On Error GoTo errHandle
    strSQL = "SELECT NVL(药库分批, 0) 药库分批,NVL(药房分批, 0) 药房分批 " & _
            " From 药品规格 WHERE 药品ID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取药品库房分批属性", lng药品ID)
    
    If rsTemp.RecordCount > 0 Then
        int药库分批 = rsTemp!药库分批
        int药房分批 = rsTemp!药房分批
    End If
    
    If int药房分批 = 1 Then     '如果药房分批，则分批属性为1
        Get药品分批属性 = True
    Else
        If int药库分批 = 1 Then
            strSQL = "SELECT 部门ID From 部门性质说明 " & _
                    " WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门ID = [1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "取部门性质", lng库房ID)
            
            bln是否具有药房性质 = (rsTemp.RecordCount > 0)
                    
            If bln是否具有药房性质 Then
                Get药品分批属性 = False
            Else
                Get药品分批属性 = True
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '单笔冲销 Write by zyb, ##20021016##
    Dim 行次_IN As Integer
    Dim 原记录状态_IN As Integer
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 药品ID_IN As Long
    Dim 冲销数量_IN As Double
    Dim 填制人_IN As String
    Dim 填制日期_IN  As String
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim n As Integer
    Dim 摘要_IN As String
    Dim str药品ID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim j As Integer
    Dim str药品 As String
    
    SaveStrike = False
    arrSql = Array()
    
    With mshBill
        For intRow = 1 To .rows - 1
            '检查冲销数量，不能小于零
            If Val(.TextMatrix(intRow, mconIntCol实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol填写数量)), Val(.TextMatrix(intRow, mconIntCol实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '检查可用数量是否足够，参数设置为不检查库存时不进行
            '分批药品按批次检查，不分批药品汇总列表中所有数量检查；冲销只判断冲销库房分批属性
            If mint库存检查入库库房 <> 0 And .TextMatrix(intRow, 0) <> "" Then
                If mbln下可用数量 = True And mint处理方式 = 2 Then
                    '不检查
                Else
                    If .TextMatrix(intRow, mconIntCol实际数量) = .TextMatrix(intRow, mconIntCol填写数量) Then
                        冲销数量_IN = .TextMatrix(intRow, mconintCol真实数量)
                    Else
                        冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                    End If
                    
                    '分批药品按批次检查，不分批药品汇总列表中所有数量检查；冲销只判断冲销库房分批属性
                    If Get药品分批属性(.TextMatrix(intRow, 0), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex))) = False Then
                        For j = 1 To .rows - 1
                            If intRow <> j Then
                                If .TextMatrix(intRow, 0) = .TextMatrix(j, 0) And .TextMatrix(intRow, 0) <> "" And .TextMatrix(j, 0) <> "" Then
                                    If .TextMatrix(j, mconIntCol实际数量) = .TextMatrix(j, mconIntCol填写数量) Then
                                        冲销数量_IN = 冲销数量_IN + .TextMatrix(j, mconintCol真实数量)
                                    Else
                                        冲销数量_IN = 冲销数量_IN + GetFormat(.TextMatrix(j, mconIntCol实际数量) * .TextMatrix(j, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If CheckStrickUsable(6, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(.TextMatrix(intRow, 0)), .TextMatrix(intRow, mconIntCol药名), _
                        Val(.TextMatrix(intRow, mconIntCol批次)), Val(冲销数量_IN), mint库存检查入库库房, Trim(txtNo.Tag), Val(.TextMatrix(intRow, mconIntCol序号)) + 1) = False Then
                        .Row = intRow
                        .MsfObj.TopRow = intRow
                        Exit Function
                    End If
                    冲销数量_IN = 0
                End If
            End If
        Next
        
        '普通和审核冲销检查实际数量
        If mint编辑状态 = 6 And mint处理方式 <> 1 Then
            str药品 = CheckNumStock(mshBill, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), 0, mconIntCol批次, mconIntCol实际数量, mconIntCol比例系数, 2, 0, mintNumberDigit)
            If str药品 <> "" Then
                If mint库存检查入库库房 = 1 Then '不足提醒
                    If MsgBox(str药品 & " 药品“实际库存”不足，是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf mint库存检查入库库房 = 2 Then '不足禁止
                    MsgBox str药品 & " 药品“实际库存”不足，不能冲销！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        填制人_IN = UserInfo.用户姓名
        填制日期_IN = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        摘要_IN = Trim(txt摘要.Text)
        
        On Error GoTo errHandle
        
        行次_IN = 0
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol实际数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                药品ID_IN = .TextMatrix(intRow, 0)
                str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & 药品ID_IN
                If .TextMatrix(intRow, mconIntCol实际数量) = .TextMatrix(intRow, mconIntCol填写数量) Then
                    冲销数量_IN = .TextMatrix(intRow, mconintCol真实数量)
                Else
                    冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
                冲销数量_IN = IIf(mint编辑状态 = 6 And mint处理方式 = 2, -1, 1) * 冲销数量_IN
                
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                
                gstrSQL = "ZL_药品移库_STRIKE("
                '行次
                gstrSQL = gstrSQL & 行次_IN
                '原记录状态
                gstrSQL = gstrSQL & "," & 原记录状态_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '序号
                gstrSQL = gstrSQL & "," & 序号_IN
                '药品ID
                gstrSQL = gstrSQL & "," & 药品ID_IN
                '冲销数量
                gstrSQL = gstrSQL & "," & 冲销数量_IN
                '填制人
                gstrSQL = gstrSQL & ",'" & 填制人_IN & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                '摘要
                gstrSQL = gstrSQL & "," & IIf(摘要_IN = "", "Null", "'" & 摘要_IN & "'")
                '冲销方式
                gstrSQL = gstrSQL & "," & mint处理方式
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行药品来冲销，请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '提示停用药品
        If str药品ID <> "" And mint处理方式 <> 1 Then
            Call CheckStopMedi(str药品ID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol采购金额))
            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    lblPurchasePrice.Caption = "成本金额合计：" & GetFormat(curTotal, mintMoneyDigit)
    lblSalePrice.Caption = "售价金额合计：" & GetFormat(Cur记帐金额, mintMoneyDigit)
    lblDifference.Caption = "差价合计：" & GetFormat(Cur记帐差价, mintMoneyDigit)
End Sub

Private Sub 提示库存数(ByVal intRow As Integer)
    Dim strNote As String, strUnit As String
    Dim rsUseCount As New Recordset
    Dim int对方库房分批属性 As Integer
    Dim int工作性质 As Integer
    Dim str当前库存 As String
    Dim str对方库存 As String
    Dim dbl对方库存 As Double
    Dim dbl当前库存 As Double
    Dim strTemp As String
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Sub
        
        int工作性质 = 0
        '对方库房分批属性
        gstrSQL = "Select a.药库分批,a.药房分批,b.工作性质 " & _
            " From 药品规格 a,部门性质说明 b " & _
            " Where a.药品id =[1] And b.部门id =[2] "
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[接收库房分批属性]", Val(.TextMatrix(intRow, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
        Do While Not rsUseCount.EOF
            If int工作性质 <> 1 Then
                If InStr(rsUseCount!工作性质, "药房") > 0 Or rsUseCount!工作性质 = "制剂室" Then
                    int工作性质 = 1
                    int对方库房分批属性 = rsUseCount!药房分批
                ElseIf InStr(rsUseCount!工作性质, "药库") > 0 Then
                    int工作性质 = 2
                    int对方库房分批属性 = rsUseCount!药库分批
                End If
            End If
            rsUseCount.MoveNext
        Loop
        
        strUnit = .TextMatrix(intRow, mconIntCol单位)
        
        If mint编辑状态 <> 10 Then
            '对方库房库存
            gstrSQL = "select sum(可用数量/" & .TextMatrix(intRow, mconIntCol比例系数) & ") as  可用数量 from 药品库存 where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1 "
            If int对方库房分批属性 = 1 And Val(.TextMatrix(intRow, mconIntCol批次)) <> 0 And gtype_UserSysParms.P174_药品移库明确批次 = 1 Then
                gstrSQL = gstrSQL & " and nvl(批次,0)=[3] "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboEnterStock.ItemData(cboEnterStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol批次)))
            
            If Not rsUseCount.EOF Then
                dbl对方库存 = GetFormat(Nvl(rsUseCount!可用数量, 0), mintNumberDigit)
            End If
            rsUseCount.Close
            
            '当前库房库存
            gstrSQL = "select Sum(Nvl(可用数量,0))/" & .TextMatrix(intRow, mconIntCol比例系数) & " as  可用数量 from 药品库存 where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1  "
            
            If Not (mbln申领单 = True And mbln明确批次 = False) Or Val(.TextMatrix(intRow, mconIntCol批次)) > 0 Then
                If gtype_UserSysParms.P174_药品移库明确批次 = 1 Then    '对于库存的查询必须是在没有明确药品移库时明确批次的情况下
                    gstrSQL = gstrSQL & " and nvl(批次,0)=[3] "
                End If
            End If
            
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol批次)))
                
            If rsUseCount.EOF Then
                .TextMatrix(intRow, mconIntCol可用数量) = 0
            Else
                .TextMatrix(intRow, mconIntCol可用数量) = GetFormat(IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0)), mintNumberDigit)
            End If
            rsUseCount.Close
            
            dbl当前库存 = GetFormat(.TextMatrix(intRow, mconIntCol可用数量), mintNumberDigit)
        Else
            '仅在发送时，显示该药品在所有库房的库存，以便于库房人员决定实际的发送数量
            If gtype_UserSysParms.P174_药品移库明确批次 = 1 And Val(.TextMatrix(intRow, mconIntCol批次)) > 0 Then  '对于库存的查询必须是在没有明确药品移库时明确批次的情况下
                strTemp = " and nvl(a.批次,0)=[2] "
            End If
            
            gstrSQL = "Select B.名称 AS 库房,Sum(Nvl(A.可用数量,0))/" & .TextMatrix(intRow, mconIntCol比例系数) & " as 可用数量 " _
                & " From 药品库存 A,部门表 B" _
                & " Where A.库房ID=B.ID And A.性质=1 And A.药品id=[1] " & strTemp _
                & " Group By B.名称, 药品id "
                
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol批次)))
            With rsUseCount
                Do While Not .EOF
                    strNote = strNote & "," & !库房 & ":" & GetFormat(Nvl(!可用数量, 0), mintNumberDigit) & strUnit
                    If cboStock.Text = !库房 Then
                        dbl当前库存 = GetFormat(Nvl(!可用数量, 0), mintNumberDigit)
                    End If
                    
                    If cboEnterStock.Text = !库房 Then
                        dbl对方库存 = GetFormat(Nvl(!可用数量, 0), mintNumberDigit)
                    End If
                    .MoveNext
                Loop
            End With
            str当前库存 = Mid(strNote, 2)
        End If
        .TextMatrix(intRow, mconIntCol库房库存) = GetFormat(dbl当前库存, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol对方库存) = GetFormat(dbl对方库存, mintNumberDigit)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtIn_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim IntCheck As Integer
    Dim intRow As Integer
    Dim blnEXIST As Boolean
    Dim intIndex As Integer, intCount As Integer
    Dim rsBill As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    Dim int包装系数 As Integer
    Dim lng药品ID As Long
    Dim blnInput As Boolean
    
    On Error GoTo ErrHand
    '初始准备
    intNO = 28
    lng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng库房ID)
    End If
    
    '需要要清除现有单据内容
    For IntCheck = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.rows Then
        If MsgBox("需要要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '处理药品单位改变
        mshBill.ClearBill
    End If
    
    '取出库检查性质
    IntCheck = 0
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取出库检查性质]", Me.cboStock.ItemData(Me.cboStock.ListIndex))

    If rsTemp.RecordCount <> 0 Then IntCheck = rsTemp!库存检查
    
    gstrSQL = "select 收费细目id,执行科室id from 收费执行科室"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "查询存储库房")
    
    '提取该单据并清空表格（只允许提取正常单据，且非退货单）
    gstrSQL = "SELECT A.药品ID,'['||C.编码||']' As 编码,'['||C.编码||']'|| Nvl(F.名称,C.名称) As 药品名称, C.名称 As 通用名,F.名称 As 商品名,C.规格,a.产地," & _
             "        C.计算单位 AS 零售单位,1 AS 零售系数,B.门诊单位,B.门诊包装,B.住院单位,B.住院包装,B.药库单位,B.药库包装, " & _
             "        NVL(A.批次,0) AS 批次,Nvl(C.是否变价,0) AS 时价,Nvl(B.药房分批,0) AS 药房分批,Nvl(B.药库分批,0) AS 药库分批,b.最大效期,A.批号,A.效期," & _
             "        B.管理费比例,B.指导差价率,A.实际数量,D.可用数量,D.实际金额,D.实际差价,E.现价,A.批准文号,B.药品来源,B.基本药物,nvl(d.平均成本价,0) as 平均成本价,a.供药单位id " & _
             " FROM 药品收发记录 A,药品规格 B,收费项目目录 C,药品库存 D,收费价目 E,收费项目别名 F " & _
             " WHERE A.药品ID=B.药品ID AND B.药品ID=C.ID AND B.药品ID=D.药品ID(+) " & _
             " AND B.药品ID=F.收费细目ID(+) AND F.性质(+)=3 AND F.码类(+)=1" & _
             " AND B.药品ID=E.收费细目ID(+) AND SYSDATE >=E.执行日期(+)  AND sysdate<=NVL(E.终止日期(+),SYSDATE)" & _
             " AND D.库房ID(+)=[2] AND D.性质(+)=1 AND Nvl(A.批次,0)=Nvl(D.批次,0)" & _
             " AND A.单据=1 AND A.记录状态=1 AND NVL(A.发药方式,0)=0 AND A.审核日期 Is Not NULL" & _
             " AND A.NO=[1] And A.库房ID+0=[2] " & _
             " ORDER BY A.序号"
    Set rsBill = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取外购入库单]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "没有找到该外购入库单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            lng药品ID = !药品id
            rsTemp.Filter = " 收费细目id=" & lng药品ID & " and 执行科室id=" & lng库房ID
            If rsTemp.RecordCount = 0 Then
                MsgBox "药品[" & !药品名称 & "]未在" & cboStock.Text & "中设置存储属性，将不能移库！"
                blnInput = True
            End If
            rsTemp.Filter = ""
            rsTemp.Filter = " 收费细目id=" & lng药品ID & " and 执行科室id=" & cboEnterStock.ItemData(cboEnterStock.ListIndex)
            If rsTemp.RecordCount = 0 Then
                MsgBox "药品[" & !药品名称 & "]未在" & cboEnterStock.Text & "中设置存储属性，将不能移库！"
                blnInput = True
            End If
            
            If blnInput = False Then
                '导入计划单相当于都是按批次移库，需要在装入数据前，先检查库存
                If !实际数量 > !可用数量 Then
                    '批次或时价药品不允许零出库
                    If !批次 <> 0 Or !时价 <> 0 Then
                        MsgBox !药品名称 & "库存不足，不允许出库！（时价或分批药品）", vbInformation, gstrSysName
                        blnInput = True
                    End If
                    '只提示一次
                    If blnInput = False Then
                        Select Case IntCheck
                        Case 1
                            If MsgBox(!药品名称 & "库存不足，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                blnInput = True
                            End If
                        Case 2
                            MsgBox !药品名称 & "库存不足，将不能移库！", vbInformation, gstrSysName
                            blnInput = True
                        End Select
                    End If
                End If
            End If
            
            '装入数据(SetColValue)
            If blnInput = False Then
                int包装系数 = Choose(mintUnit, 1, !门诊包装, !住院包装, !药库包装)
                If Not SetColValue(intRow, !药品id, !编码, !通用名, IIf(IsNull(!商品名), "", !商品名), _
                    Nvl(!药品来源), Nvl(!基本药物), Nvl(!规格), Nvl(!产地), _
                    Choose(mintUnit, !零售单位, !门诊单位, !住院单位, !药库单位), Nvl(!现价, 0), _
                    Nvl(!批号), Nvl(!效期), Nvl(!最大效期, 24), !药库分批, Nvl(!可用数量, 0), Nvl(!实际金额, 0), Nvl(!实际差价, 0), _
                    Nvl(!指导差价率, 0), int包装系数, Nvl(!批次, 0), !时价, _
                    !药房分批, !供药单位ID, IIf(IsNull(!批准文号), "", !批准文号)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If
    
                '填写数量、采购价、售价等列
                mshBill.TextMatrix(intRow, mconIntCol行号) = intRow
                mshBill.TextMatrix(intRow, mconIntCol实际数量) = GetFormat(!实际数量 / int包装系数, mintNumberDigit)
                mshBill.TextMatrix(intRow, mconIntCol填写数量) = GetFormat(!实际数量 / int包装系数, mintNumberDigit)
                mshBill.TextMatrix(intRow, mconIntCol采购价) = GetFormat(!平均成本价 * int包装系数, mintCostDigit)
                mshBill.TextMatrix(intRow, mconIntCol采购金额) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol采购价)) * Val(mshBill.TextMatrix(intRow, mconIntCol实际数量)), mintMoneyDigit)
                mshBill.TextMatrix(intRow, mconIntCol售价金额) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol售价)) * Val(mshBill.TextMatrix(intRow, mconIntCol实际数量)), mintMoneyDigit)
                mshBill.TextMatrix(intRow, mconintCol差价) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol售价金额)) - mshBill.TextMatrix(intRow, mconIntCol采购金额), mintMoneyDigit)
    
                intRow = intRow + 1
                mshBill.rows = mshBill.rows + 1
            End If
            blnInput = False
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.ClearBill
End Sub

Private Sub txt摘要_Change()
   
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    
    OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    OpenIme
End Sub

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim dbltotal As Double              '某种药品输入的所有数量
    Dim intLop As Integer
    Dim dbl原填写数量 As Double
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False
'    If (mbln申领单 And Not mbln明确批次) Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        dblUsableQuantity = GetFormat(.TextMatrix(intRow, mconIntCol可用数量), mintNumberDigit)
        If .TextMatrix(intRow, mconIntCol批次) > 0 Or Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1 Then     '对分批或者时价药品检查库存
            If mint编辑状态 = 1 Then
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And .TextMatrix(intRow, mconIntCol批次) = .TextMatrix(intLop, mconIntCol批次) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol填写数量)
                        End If
                    End If
                Next
                
                If dbl填写数量 + dbltotal > dblUsableQuantity Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity - dbltotal & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And .TextMatrix(intRow, mconIntCol批次) = .TextMatrix(intLop, mconIntCol批次) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol实际数量)
                        End If
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                If dbl填写数量 + dbltotal > dblUsableQuantity + dbl原填写数量 Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 - dbltotal & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
            CompareUsableQuantity = True
            Exit Function
        End If
        
        ' 对移出库房是库房且药品是分批核算的药品以外的判断
        
        If mint库存检查 = 0 Then
            '0-不检查
        ElseIf mint库存检查 = 1 Then
            '1-检查，不足提醒
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    If MsgBox("对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                If dbl填写数量 > dblUsableQuantity + dbl原填写数量 Then
                    If MsgBox("对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Then
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mconIntCol批次) = "", "0", .TextMatrix(intLop, mconIntCol批次)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol填写数量)
                        End If
                    End If
                Next
                
                If dbl填写数量 + dbltotal > dblUsableQuantity Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity - dbltotal & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Or mint编辑状态 = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mconIntCol批次) = "", "0", .TextMatrix(intLop, mconIntCol批次)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol实际数量)
                        End If
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                If dbl填写数量 + dbltotal > dblUsableQuantity + dbl原填写数量 Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 - dbltotal & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'打印单据
Private Sub printbill()
    Dim int单位系数 As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint售价单位
            int单位系数 = 4
        Case mconint门诊单位
            int单位系数 = 2
        Case mconint住院单位
            int单位系数 = 1
        Case mconint药库单位
            int单位系数 = 3
    End Select
    strNo = txtNo.Tag
    DoEvents
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), mint记录状态, int单位系数, 1304, "药品调拨冲销单", strNo
End Sub

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    On Error GoTo errHandle
    
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Call zlDataBase.OpenRecordset(rsBatchNolen, gstrSQL, "取字段长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    Dim lng库房ID As Long, lng药品ID As Long, lng药品ID_Last As Long, lng批次 As Long
    Dim bln库房 As Boolean, bln分批 As Boolean, bln时价 As Boolean, blnAddRow As Boolean
    Dim dbl填写数量 As Double, dbl申领数量 As Double, Dbl数量 As Double, dbl比例系数 As Double
    Dim dbl现价 As Currency, dbl现价_时价 As Double, dbl成本价 As Double
    Dim lngCol As Long, lngCols As Long, lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim dbl实际数量 As Double
    Dim intCount As Integer
            
    '对药品记录进行自动分解，仅处理批次药品
    On Error GoTo ErrHand
    Debug.Print "开始分解：" & Now
    Screen.MousePointer = 11
    lngRow = 1: lngCols = mshBill.Cols - 1
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln库房 = CheckStockProperty(lng库房ID)
    
    Do While True
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl申领数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol填写数量))
'        dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量))
        dbl填写数量 = dbl申领数量
        dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数))
        lng批次 = Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
        
        If lng药品ID = 0 Then Exit Do
        
        '提取该药品对于出库库房是否分批、时价的属性
        If lng药品ID <> lng药品ID_Last Then
            lng药品ID_Last = lng药品ID
            gstrSQL = " Select Nvl(A.药库分批,0) 药库分批,Nvl(A.药房分批,0) 药房分批," & _
                      " Nvl(B.是否变价,0) 时价,Nvl(P.现价,0) 现价,Nvl(A.成本价,0) 成本价" & _
                      " From 药品规格 A,收费项目目录 B,收费价目 P" & _
                      " Where A.药品ID = B.ID And B.ID=P.收费细目ID And A.药品ID =[1] " & _
                      " And Sysdate between P.执行日期 And Nvl(P.终止日期,Sysdate)"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品对于出库库房是否分批、时价的属性]", lng药品ID)
            
            bln时价 = (rsTemp!时价 = 1)
            dbl现价 = rsTemp!现价 * dbl比例系数
            bln分批 = IIf(bln库房, (rsTemp!药库分批 = 1), (rsTemp!药房分批 = 1))
        End If
        
        '如果该药品是分批药品，但批次为零，则说明需要自动分解
        blnAddRow = False
        If bln分批 = True And lng批次 = 0 Then
'            If blnCheck Then
'                If dbl填写数量 > Val(mshBill.TextMatrix(lngRow, mconIntCol可用数量)) Then
'                    MsgBox "第" & lngRow & "行的药品是批次或时价药品，而该药品当前库存不足，不能继续！", vbInformation, gstrSysName
'                    Screen.MousePointer = 0: Exit Function
'                End If
'            End If
            gstrSQL = " Select Nvl(可用数量,0)/" & dbl比例系数 & " As 可用数量,Nvl(实际数量,0)/" & dbl比例系数 & " As 实际数量," & _
                      " Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价, nvl(平均成本价,0) 平均成本价," & _
                      " Nvl(批次,0) 批次,上次批号 批号,to_char(效期,'yyyy-MM-dd') 效期,上次产地 产地,NVL(上次供应商ID,0) 上次供应商ID,批准文号,nvl(零售价,0)*" & dbl比例系数 & " As 零售价 " & _
                      " From 药品库存 Where 库房ID=[1] And 药品ID=[2] And 性质=1 And Nvl(可用数量,0)>0 "
        
            If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
                gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
            ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 1 Then
                gstrSQL = gstrSQL & " Order by 效期,Nvl(批次,0)"
            ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 2 Then
                gstrSQL = gstrSQL & " Order by 上次批号,Nvl(批次,0)"
            Else
                gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
            End If

            Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品在指定库存的所有库存记录]", lng库房ID, lng药品ID)
            With rsCheck
                intCount = 0
                Do While Not .EOF
                    intCount = intCount + 1
                    mshBill.Redraw = False
                    '重新写记录
                    blnAddRow = False
                    If .AbsolutePosition <> 1 Then
                        mshBill.MsfObj.AddItem "", lngRow
                        For lngCol = 0 To lngCols
                            mshBill.TextMatrix(lngRow, lngCol) = mshBill.TextMatrix(lngRow - 1, lngCol)
                        Next
                        mshBill.TextMatrix(lngRow, mconIntCol填写数量) = "0"
                        mshBill.RowData(lngRow) = mshBill.RowData(lngRow - 1)
                    End If
                    
                    If intCount = 1 Then
                        dbl实际数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量))
                    End If
                    
                    '填写批次相关信息
                    mshBill.TextMatrix(lngRow, mconIntCol行号) = lngRow
                    mshBill.TextMatrix(lngRow, mconIntCol序号) = (lngRow - 1) * 2 + 1
                    mshBill.TextMatrix(lngRow, mconIntCol批次) = rsCheck!批次
                    mshBill.TextMatrix(lngRow, mconIntCol批号) = IIf(IsNull(rsCheck!批号), "", rsCheck!批号)
                    mshBill.TextMatrix(lngRow, mconIntCol产地) = IIf(IsNull(rsCheck!产地), "", rsCheck!产地)
                    mshBill.TextMatrix(lngRow, mconIntCol效期) = IIf(IsNull(rsCheck!效期), "", rsCheck!效期)
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And mshBill.TextMatrix(lngRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        mshBill.TextMatrix(lngRow, mconIntCol效期) = Format(DateAdd("D", -1, mshBill.TextMatrix(lngRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    mshBill.TextMatrix(lngRow, mconIntCol上次供应商ID) = rsCheck!上次供应商ID
                    mshBill.TextMatrix(lngRow, mconIntCol批准文号) = IIf(IsNull(rsCheck!批准文号), "", rsCheck!批准文号)
                    
                    '重新计算价格相关信息
                    If rsCheck!实际数量 > 0 Then
                        If Val(mshBill.TextMatrix(lngRow, mconIntCol批次)) > 0 Then
                            dbl现价_时价 = IIf(rsCheck!零售价 > 0, rsCheck!零售价, rsCheck!实际金额 / rsCheck!实际数量)
                        Else
                            dbl现价_时价 = rsCheck!实际金额 / rsCheck!实际数量
                        End If
                    End If
                    
                    If dbl填写数量 <= rsCheck!可用数量 Then
                        Dbl数量 = dbl填写数量
                    Else
                        Dbl数量 = rsCheck!可用数量
                    End If
                    If Dbl数量 > dbl填写数量 Then Dbl数量 = dbl填写数量
                    
                    mshBill.TextMatrix(lngRow, mconIntCol填写数量) = GetFormat(Dbl数量, mintNumberDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol实际数量) = GetFormat(Dbl数量, mintNumberDigit)
                    
                    '特殊处理，当分批且没有库存时需要将批号和上次产地自动填上（无库存填上信息后不影响），方便管理员操作
                    If Val(mshBill.TextMatrix(lngRow, mconIntCol分批属性)) = 1 And Dbl数量 = 0 Then
                        gstrSQL = "select 上次产地,上次批号 from 药品规格 where 药品id=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "产地批号信息", lng药品ID)
                        mshBill.TextMatrix(lngRow, mconIntCol产地) = IIf(IsNull(rsTemp!上次产地), "", rsTemp!上次产地)
                        mshBill.TextMatrix(lngRow, mconIntCol批号) = IIf(IsNull(rsTemp!上次批号), "", rsTemp!上次批号)
                    End If
                    
                    If dbl实际数量 <> mshBill.TextMatrix(lngRow, mconIntCol实际数量) Then
                        mshBill.TextMatrix(lngRow, mconintCol真实数量) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量)) * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mintNumberDigit)
                    End If
                    
                    If Trim(mshBill.TextMatrix(lngRow, mconIntCol实际数量)) = "" Then mshBill.TextMatrix(lngRow, mconIntCol实际数量) = GetFormat(0, mintNumberDigit)
                    
                    mshBill.TextMatrix(lngRow, mconIntCol实际差价) = GetFormat(rsCheck!实际差价, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol实际金额) = GetFormat(rsCheck!实际金额, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol可用数量) = GetFormat(rsCheck!可用数量, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol售价) = GetFormat(IIf(bln时价, dbl现价_时价, dbl现价), mintPriceDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol售价金额) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol售价)) * Dbl数量, mintMoneyDigit)
                    
                    If Dbl数量 <> 0 Then
                        mshBill.TextMatrix(lngRow, mconIntCol采购价) = GetFormat(rsCheck!平均成本价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mintCostDigit)
                    End If
                    mshBill.TextMatrix(lngRow, mconIntCol采购金额) = GetFormat(mshBill.TextMatrix(lngRow, mconIntCol采购价) * Dbl数量, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconintCol差价) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol售价金额)) - Val(mshBill.TextMatrix(lngRow, mconIntCol采购金额)), mintMoneyDigit)
                    
                    dbl填写数量 = dbl填写数量 - Dbl数量
                    dbl申领数量 = dbl申领数量 - Dbl数量
                    If dbl填写数量 = 0 Then Exit Do
                    lngRow = lngRow + 1
                    blnAddRow = True
                    .MoveNext
                Loop
                If dbl申领数量 <> 0 And rsCheck.RecordCount <> 0 Then
                    If blnAddRow Then
                        mshBill.TextMatrix(lngRow - 1, mconIntCol填写数量) = GetFormat(dbl申领数量 + Dbl数量, mintNumberDigit)
                    Else
                        mshBill.TextMatrix(lngRow, mconIntCol填写数量) = GetFormat(dbl申领数量 + Dbl数量, mintNumberDigit)
                    End If
                End If
            End With
            
            '如果库存记录为零，则说明未进行分解，需要将申领数量与实际数量清为零
            If dbl填写数量 <> 0 And rsCheck.RecordCount = 0 Then
                mshBill.TextMatrix(lngRow, mconIntCol行号) = lngRow
                mshBill.TextMatrix(lngRow, mconIntCol序号) = (lngRow - 1) * 2 + 1
                mshBill.TextMatrix(lngRow, mconIntCol实际数量) = GetFormat(0, mintNumberDigit)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = ""
                mshBill.TextMatrix(lngRow, mconIntCol采购金额) = ""
                mshBill.TextMatrix(lngRow, mconintCol差价) = ""
                
                '特殊处理，当分批且没有库存时需要将批号和上次产地自动填上（无库存填上信息后不影响），方便管理员操作
                If Val(mshBill.TextMatrix(lngRow, mconIntCol分批属性)) = 1 Then
                    gstrSQL = "select 上次产地,上次批号 from 药品规格 where 药品id=[1]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "产地批号信息", lng药品ID)
                    mshBill.TextMatrix(lngRow, mconIntCol产地) = IIf(IsNull(rsTemp!上次产地), "", rsTemp!上次产地)
                    mshBill.TextMatrix(lngRow, mconIntCol批号) = IIf(IsNull(rsTemp!上次批号), "", rsTemp!上次批号)
                End If
            End If
        Else
            mshBill.TextMatrix(lngRow, mconIntCol行号) = lngRow
            mshBill.TextMatrix(lngRow, mconIntCol序号) = (lngRow - 1) * 2 + 1
        End If
        If blnAddRow = False Then lngRow = lngRow + 1
    Loop
    
    mblnChange = True
    AutoExpend = True
    mshBill.Redraw = True
    Call ShowColor
    Screen.MousePointer = 0
    Debug.Print "结束分解：" & Now
    
    If mbln自动分解未完成 = True Then mbln自动分解未完成 = False
    
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckStockProperty(ByVal lng库房ID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHandle
    
    '检查指定库房是药库、药房还是制剂室(传入的库房肯定是药库、药房或制剂室中的一个)
    gstrSQL = " Select 部门ID From 部门性质说明 " & _
              " Where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1] "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[判断是不是药房或制剂室]", lng库房ID)
              
    If rsCheck.EOF Then
        CheckStockProperty = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsertRow(ByVal lngRow As Long)
    Dim lngReserve As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    Debug.Print Now
    lngReserve = lngRow
    lngRows = mshBill.rows - 1
    lngCols = mshBill.Cols - 1
    mshBill.rows = mshBill.rows + 1
    
    '将当前行及以下行全部下移
    For lngRow = lngRows To lngReserve Step -1
        For lngCol = 0 To lngCols
            mshBill.TextMatrix(lngRow + 1, lngCol) = mshBill.TextMatrix(lngRow, lngCol)
        Next
        mshBill.RowData(lngRow + 1) = mshBill.RowData(lngRow)
        '校正行号
        mshBill.TextMatrix(lngRow + 1, mconIntCol行号) = lngRow + 1
    Next
    Debug.Print Now
End Sub

Private Sub ShowColor(Optional ByVal lngCurRow As Long = 0)
    '在查阅或审核时，将库存不足的记录以暗红色显示出来
    Dim lngSelect_Row  As Long, lngSelect_Col As Long
    Dim lng药品ID As Long
    Dim lngColor As Long, lngNewColor As Long '如果现在的颜色与要上的颜色一样，则不处理
    Dim dbl填写数量 As Double, dbl可用数量 As Double
    Dim lngRow As Long, BlnDO As Boolean
    Dim i As Long, j As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHand
    mshBill.Redraw = False
    mblnEnterCell = False
    lngSelect_Row = mshBill.Row: lngSelect_Col = mshBill.Col
    lngRow = IIf(lngCurRow > 0, lngCurRow, 1)
    
    Do While True
        If lngRow > mshBill.rows - 1 Then Exit Do
        mshBill.Row = lngRow: mshBill.Col = mconIntCol药名
        lngColor = mshBill.MsfObj.CellForeColor
        
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol填写数量))
        dbl可用数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol可用数量))
        If lng药品ID = 0 Then Exit Do
        
        gstrSQL = "select decode(药库分批,Null,0,药库分批) 药库分批,decode(药房分批,Null,0,药房分批) 药房分批 from 药品规格 where 药品id=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "查询分批", lng药品ID)
        
        If rsTemp Is Nothing Then
            Exit Sub
        Else
            If rsTemp!药库分批 = 1 Or rsTemp!药房分批 = 1 Then
                '库存不足的药品设置颜色
                BlnDO = False
                If dbl可用数量 < dbl填写数量 Then BlnDO = True
                lngNewColor = IIf(BlnDO, &HC0, &H0)
                If lngColor <> lngNewColor Then
                    '只对药名列进行上色处理
                    j = mshBill.ColData(mconIntCol药名)
                    If j = 5 Then mshBill.ColData(mconIntCol药名) = 0
                    mshBill.Col = mconIntCol药名
                    mshBill.MsfObj.CellForeColor = lngNewColor
                    mshBill.ColData(mconIntCol药名) = j
                End If
            End If
            If lngCurRow > 0 Then Exit Do
            lngRow = lngRow + 1
        End If
    Loop
    mshBill.Row = lngSelect_Row: mshBill.Col = lngSelect_Col
    mshBill.Redraw = True
    mblnEnterCell = True
    Exit Sub
ErrHand:
    mshBill.Redraw = True
    mblnEnterCell = True
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckStock() As Boolean
    Dim dbl比例系数 As Double, dbl实际数量 As Double, dbl填写数量 As Double
    Dim lngRow As Long, lngRows As Long, int库存检查 As Integer
    Dim lng药品ID As Long, lng库房ID As Long, lng批次 As Long
    Dim bln库房 As Boolean, bln特药 As Boolean
    Dim str药品ID As String, strMsg As String
    Dim rsTemp As ADODB.Recordset
    Dim rsProperty As ADODB.Recordset           '药品规格
    Dim rsCheck As ADODB.Recordset              '药品库存
    Dim arrDrugID As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    Set rsProperty = New ADODB.Recordset
    With rsProperty
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "药库分批", adDouble, 18, adFldIsNullable
        .Fields.Append "药房分批", adDouble, 18, adFldIsNullable
        .Fields.Append "是否变价", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsCheck = New ADODB.Recordset
    With rsCheck
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "可用数量", adDouble, 18, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrDrugID = Array()
    
    '检查单据中各药品的库存
    'mint库存检查:0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '分批或时价药品不受此限
    Debug.Print "开始检查库存:" & Now
    lngRows = mshBill.rows - 1
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln库房 = CheckStockProperty(lng库房ID)
    For lngRow = 1 To lngRows
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng药品ID <> 0 Then
            If InStr(1, "," & str药品ID & ",", "," & lng药品ID & ",") = 0 Then
                If Len(IIf(str药品ID = "", "", str药品ID & ",") & lng药品ID) > 4000 Then
                    ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
                    arrDrugID(UBound(arrDrugID)) = str药品ID
                    str药品ID = lng药品ID
                Else
                    str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & lng药品ID
                End If
            End If
        End If
    Next
    
    If str药品ID = "" And UBound(arrDrugID) < 0 Then
        CheckStock = True
        Exit Function
    ElseIf str药品ID <> "" Then
        ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
        arrDrugID(UBound(arrDrugID)) = str药品ID
    End If
    
    '提取本单据内所有药品的属性
    gstrSQL = " Select  A.药品ID,'['||B.编码||']'||B.名称 通用名,A.药库分批,A.药房分批,B.是否变价" & _
              " From 药品规格 A,收费项目目录 B" & _
              " Where A.药品ID=B.ID And A.药品ID in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) "
    
    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "提取本单据内所有药品的属性", CStr(arrDrugID(i)))
        
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsProperty
                    .AddNew
                    !药品id = rsTemp!药品id
                    !通用名 = rsTemp!通用名
                    !药库分批 = rsTemp!药库分批
                    !药房分批 = rsTemp!药房分批
                    !是否变价 = rsTemp!是否变价
                    
                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next
    
    gstrSQL = "Select a.药品id, Nvl(a.批次, 0) As 批次, Sum(Nvl(可用数量, 0)) As 可用数量,Sum(Nvl(实际数量, 0)) As 实际数量 " & _
        " From 药品库存 A, 药品规格 C" & _
        " Where a.库房id = [1] And a.药品id = c.药品id And a.性质 = 1 And c.药品id in (select * from Table(Cast(f_Num2list([2]) As Zltools.t_Numlist))) " & _
        " Group By a.药品id, Nvl(a.批次, 0) "
    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取本单据内所有药品的当前库存]", lng库房ID, CStr(arrDrugID(i)))
    
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsCheck
                    .AddNew
                    !药品id = rsTemp!药品id
                    !批次 = rsTemp!批次
                    !可用数量 = rsTemp!可用数量
                    !实际数量 = rsTemp!实际数量
                    
                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next
    
    '检查每个药品
    For lngRow = 1 To lngRows
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng药品ID <> 0 Then
            lng批次 = Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            dbl比例系数 = Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数))
            dbl填写数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量))
            
            dbl实际数量 = 0
            '查找该药品的库存记录
            rsCheck.Filter = "药品ID=" & lng药品ID & " And 批次=" & lng批次
            
            If rsCheck.RecordCount <> 0 Then
                If mint编辑状态 = 10 Then   '发送用可用数量判断
                    dbl实际数量 = Val(GetFormat(Nvl(rsCheck!可用数量, 0) / dbl比例系数, mintNumberDigit))
                Else    '审核用实际数量判断
                    dbl实际数量 = Val(GetFormat(Nvl(rsCheck!实际数量, 0) / dbl比例系数, mintNumberDigit))
                End If
            End If
            
            '如果库存的实际数量不够
            If Not (dbl实际数量 >= dbl填写数量) Then
                int库存检查 = mint库存检查
                '如果该药品是时价或分批，库存不足不允许出库，相当于禁止出库
                rsProperty.Filter = "药品ID=" & lng药品ID
                bln特药 = (IIf(bln库房, (rsProperty!药库分批 = 1), (rsProperty!药房分批 = 1)) Or (rsProperty!是否变价 = 1))
                strMsg = ""
                If bln特药 Then
                    int库存检查 = 2
                    '如果是批次药品，但批次小于等于零，说明未执行分解功能
                    If lng批次 <= 0 And IIf(bln库房, (rsProperty!药库分批 = 1), (rsProperty!药房分批 = 1)) Then
                        strMsg = "（请先执行分解功能明确批次药品的出库批次）"
                    End If
                End If
                
                '定位到具体行
                mshBill.Row = lngRow
                mshBill.MsfObj.TopRow = lngRow
                '按正常流程进行提示或禁止
                Select Case int库存检查
                Case 1  '仅提示
                    Debug.Print "无库存退出:" & Now
                    If MsgBox(rsProperty!通用名 & "的库存不足，是否继续？" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    Debug.Print "无库存退出:" & Now
                    MsgBox rsProperty!通用名 & "的库存不足！" & strMsg, vbInformation, gstrSysName
                    Exit Function
                End Select
            End If
        End If
    Next
    
    rsCheck.Filter = 0
    rsProperty.Filter = 0
    CheckStock = True
    Debug.Print "完成检查库存:" & Now
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendPhysic() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '检查当前单据是否已发送
    On Error GoTo ErrHand

    gstrSQL = "Select 配药日期 From 药品收发记录 " & _
              "Where 单据=6 And NO=[1] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查当前单据是否已发送]", Me.txtNo.Tag)
              
    If (Nvl(rsTemp!配药日期) = "") Then
        MsgBox "该单据已被其他操作员取消发送，不允许接收！", vbInformation, gstrSysName
        Exit Function
    End If
    SendPhysic = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetForeColor_ROW(ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer
    '设置某行的颜色
    With mshBill
        intCol = .Col
        mblnEnterCell = False
        For i = mconIntCol药名 To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        mblnEnterCell = True
    End With
End Sub

Private Function IsSelf_Command(ByVal lng药品ID As Long) As Boolean
    '判断是否为自制药品，且移入库房是制剂室（含有制剂室的属性）
    Dim bln自制药品 As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '检查移入库房
    gstrSQL = "Select 1 From 部门性质说明 Where 部门ID=[1] And 工作性质='制剂室'"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查移入库房]", cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '检查是否是自制药品
    gstrSQL = "Select Nvl(自制药品,0) As 自制药品 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查是否自制药品]", lng药品ID)
    
    bln自制药品 = (rsTemp!自制药品 = 1)
    '提取自制组成药品
    If bln自制药品 Then
        gstrSQL = "Select 原料药品ID,分子,分母 From 自制药品构成 Where 自制药品ID=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取自制组成药品]", lng药品ID)
        bln自制药品 = (rsTemp.RecordCount <> 0)
    End If
    
    IsSelf_Command = bln自制药品
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetMaterial(ByVal lng药品ID As Long) As ADODB.Recordset
    '获取自制药品的原料药品信息
    Dim rsMaterial As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        " Select B.药品ID,Nvl(B.药库分批,0) As 药库分批,Nvl(B.药房分批,0) As 药房分批,C.编码 AS 药品编码,D.名称 As 商品名,C.名称 As 通用名," & _
        "        B.药品来源,B.基本药物,C.规格,C.产地,C.计算单位 AS 售价单位,B.门诊单位,B.门诊包装,B.住院单位,B.住院包装,B.药库单位,B.药库包装,Nvl(C.是否变价,0) As 时价," & _
        "        E.现价 AS 售价,Nvl(F.批次,0) As 批次,F.上次批号 As 批号,F.效期 As 效期,Nvl(B.最大效期,0) As 最大效期,Nvl(F.可用数量,0) As 可用数量," & _
        "        Nvl(F.实际金额,0) As 实际金额,Nvl(F.实际差价,0) As 实际差价,Nvl(B.指导差价率,0) As 指导差价率,Nvl(F.上次供应商ID,0) 上次供应商ID,F.批准文号 " & _
        " From 自制药品构成 A,药品规格 B,收费项目目录 C,收费项目别名 D,收费价目 E," & _
        "      (Select 库房ID,药品ID,批次,上次批号,效期,可用数量,实际金额,实际差价,上次供应商ID,批准文号 From 药品库存" & _
        "      Where (库房ID,药品ID,Nvl(批次,0)) In" & _
        "           (Select A.库房ID,A.药品ID,Min(Nvl(A.批次,0)) From 药品库存 A,自制药品构成 B" & _
        "            Where A.库房ID = [1] And A.药品ID = B.原料药品ID And A.性质 = 1 And B.自制药品ID =[2] " & _
        "            Group By A.库房ID,A.药品ID)) F" & _
        " Where A.自制药品ID = [2] And A.原料药品ID = B.药品ID And B.药品ID = C.Id" & _
        " And B.药品ID=D.收费细目Id(+) And D.性质(+)=3 And D.码类(+)=1" & _
        " And B.药品ID=E.收费细目ID And ((Sysdate Between 执行日期 And 终止日期) Or 终止日期 Is Null )" & _
        " And B.药品ID=F.药品ID"
    Set rsMaterial = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[获取自制药品的原料药品信息]", cboStock.ItemData(cboStock.ListIndex), lng药品ID)
    Set GetMaterial = rsMaterial
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As String
    '功能：用来检查列表中已有药品与新选择的药品是否重复，以此来判断需要新增多少行
    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    
    rsTemp.MoveFirst
    str批次 = ""
    Do While Not rsTemp.EOF
        If gtype_UserSysParms.P174_药品移库明确批次 = 0 Then
            str批次 = "0"
        Else
            str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        End If
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & str批次 & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 Then
                CheckRedo = CheckRedo & .TextMatrix(i, 0) & ","
            End If
        Next
    End With
End Function

'Private Function GetRs(ByVal str药品id As String, ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
'    '功能：检验是否有重复记录，并将重复的记录过滤掉
'    '当同时选择了多个记录 并且有多个记录与之相同时，则只提示一次
'
'    Dim strTemp As String
'    Dim i As Integer
'
'    If str药品id <> "" Then
'        strTemp = ""
'        For i = 0 To UBound(Split(str药品id, ",")) - 1
'            strTemp = strTemp & "药品id<>" & Split(str药品id, ",")(i) & " and "
'        Next
'
'        If strTemp <> "" Then
'            strTemp = Mid(strTemp, 1, Len(strTemp) - 4)
'        End If
'        rsTemp.Filter = strTemp
'    End If
'    If str药品id <> "" And mbln提示 = False Then
'        MsgBox "对不起，已有该药品或该药品的相同批次，重复记录将不添加！", vbInformation, gstrSysName
'        mbln提示 = True
'    End If
'    Set GetRs = rsTemp
'End Function

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：用来检查列表中已有药品与新选择的药品是否重复和时价药品是否有库存

    Dim i As Integer
    Dim strTemp As String
    Dim str批次 As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str库存 As String
    Dim strSQL As String
    Dim strDub As String    '重复药品
    Dim strNotNum As String  '无库存药品
    Dim str重复药名 As String   '用来记录重复选择了的药品名称
    Dim strNot药名 As String    '用来记录哪些药品是时价但无库存
    Dim rsRe As ADODB.Recordset
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str批次 = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        If gtype_UserSysParms.P174_药品移库明确批次 = 0 Then
            str批次 = "0"
        Else
            str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        End If
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
        
        If rsTemp!时价 = 1 Then '时价无库存的药品
            If Not mbln申领单 Or (mbln申领单 And mbln明确批次) Then
                gstrSQL = "Select Decode(Nvl(批次, 0), 0, 实际金额 / 实际数量, Nvl(零售价, 实际金额 / 实际数量))*" & Choose(mintUnit, 1, rsTemp!门诊包装, rsTemp!住院包装, rsTemp!药库包装) & " as  售价 " _
                    & "  from 药品库存 " _
                    & " where 库房id=[1] " _
                    & " and 药品id=[2] " _
                    & " and 性质=1 and 实际数量>0 and " _
                    & " nvl(批次,0)=[3] "
                Set rsRe = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!药品id, IIf(IsNull(rsTemp!批次), 0, rsTemp!批次))
                
                If rsRe.EOF Then
                    str库存 = str库存 & rsTemp!药品id & "," & rsTemp!通用名 & "|"
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
        
    With mshBill    '把重复的查询出来
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol批次)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
            End If
        Next
        
        If strInfo <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        If str库存 <> "" Then
            strNotNum = ""
            For i = 0 To UBound(Split(str库存, "|")) - 1
                strNotNum = strNotNum & "药品id<>" & Split(Split(str库存, "|")(i), ",")(0) & " and "
                If UBound(Split(strNot药名, ",")) <= 2 Then
                    strNot药名 = strNot药名 & Split(Split(str库存, "|")(i), ",")(1) & ","
                End If
            Next
            If strNotNum <> "" Then
                strNotNum = Mid(strNotNum, 1, Len(strNotNum) - 4)
            End If
        End If
        '判断以什么方式拼接sql
        
        If str重复药名 <> "" And strNot药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & strNot药名 & "是时价药品，没有库存不允许出库！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub & " and " & strNotNum
        End If
        If str重复药名 <> "" And strNot药名 = "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If str重复药名 = "" And strNot药名 <> "" Then
            MsgBox strNot药名 & "是时价药品，没有库存不允许出库！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSQL = strNotNum
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 检查价格() As Boolean
    '功能：新增时，判断药品是否是最新价格，不是则修改后提示
    Dim strMsg As String '保存提示信息
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim bln是否时价 As Boolean
    
    On Error GoTo errHandle
    
    检查价格 = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
            
                bln是否时价 = Val(Split(.TextMatrix(i, mconIntCol最大效期), "||")(1)) = 1
                Dbl数量 = Val(.TextMatrix(i, mconIntCol实际数量))
                
                '检查成本价
                dbl成本价 = zlStr.FormatEx(Get成本价(Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol批次))) * Val(.TextMatrix(i, mconIntCol比例系数)), mintCostDigit)
                If .TextMatrix(i, mconIntCol采购价) <> dbl成本价 Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol采购价) = zlStr.FormatEx(dbl成本价, mintCostDigit, , True)
                    .TextMatrix(i, mconIntCol采购金额) = zlStr.FormatEx(.TextMatrix(i, mconIntCol采购价) * Dbl数量, mintMoneyDigit, , True)
                End If
                
                '检查售价
                dbl零售价 = zlStr.FormatEx(Get售价(bln是否时价, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol批次))) * Val(.TextMatrix(i, mconIntCol比例系数)), mintPriceDigit)
                If .TextMatrix(i, mconIntCol售价) <> dbl零售价 Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol售价) = zlStr.FormatEx(dbl零售价, mintPriceDigit, , True)
                    .TextMatrix(i, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(i, mconIntCol售价) * Dbl数量, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mconintCol差价) = zlStr.FormatEx(Val(.TextMatrix(i, mconIntCol售价金额)) - Val(.TextMatrix(i, mconIntCol采购金额)), mintMoneyDigit, , True)
                
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "有记录未使用最新价格，程序已自动完成更新（成本价、成本金额、售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            检查价格 = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
