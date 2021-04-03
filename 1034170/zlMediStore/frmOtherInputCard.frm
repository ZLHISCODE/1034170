VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherInputCard 
   Caption         =   "药品其他入库单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmOtherInputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh产地 
      Height          =   2175
      Left            =   2760
      TabIndex        =   37
      Top             =   1380
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
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   2775
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   2805
      Begin VB.CommandButton CmdYes 
         Caption         =   "确定"
         Height          =   345
         Left            =   810
         TabIndex        =   35
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "取消"
         Height          =   345
         Left            =   1800
         TabIndex        =   36
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox Txt加价率 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         MaxLength       =   8
         TabIndex        =   34
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.Label Lbl加价率 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "加成率(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   33
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "    请输入加成率，零售价的计算公式：零售价=成本价*(1+加成率%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   32
         Top             =   150
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   10
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   3
         Top             =   950
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
         TabIndex        =   5
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
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
         TabIndex        =   27
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6510
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9210
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   4
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品其他入库单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   17
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   5925
         TabIndex        =   14
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   8400
         TabIndex        =   13
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入出类别(&T)"
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
            Picture         =   "frmOtherInputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1000
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
            Picture         =   "frmOtherInputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOtherInputCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":3080
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
      TabIndex        =   24
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
End
Attribute VB_Name = "frmOtherInputCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '是否可选库房
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintBatchNoLen As Integer           '数据库中批号定义长度
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mbln下可用数量 As Boolean           '填单是否下可用数量
Private mrs分段加成 As ADODB.Recordset      '分段加成集合
Private mint时价分段加成方式 As Integer     ' 0-不按分段加成（默认） 1-按分段加成
Private mint时价入库时取上次售价 As Integer '0-时价药品入库时不取上次售价,1-时价药品入库时取上次售价
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private mint取上次成本价方式 As Integer     '0-优先从药品库存取;1-优先从药品规格取
Private mrsInOutType As Recordset           '入出类别
Private mbln加价率 As Boolean               '时价药品是否必须输入加价率
Private mdbl加价率 As Double
Private mstrPrivs As String                 '权限
Private mbln日期提示 As Boolean
Private mbln效期提示 As Boolean             '是否提示失效期的药品,主要用于在加载单据时繁琐的过期药品提示。true-提示;false-不提示

Private marrFrom As Variant                   '纪录用户恢复窗体表列格宽度
Private marrInitGrid As Variant                '纪录初始化窗体表列格宽度

'Private mint时价售价位数 As Integer         '记录时价药品用户自定的小数位数
Private mintLastCol As Integer              '用户的列设置中的最后可见列的列号

Private mcolUsedCount As Collection         '已使用的数量集合
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容

Private mlng入库库房 As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称
Private Const MStrCaption As String = "药品其他入库管理"

Private mobjPlugIn As Object '外挂部件

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrTime_Start As String                      '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

Private mstr选择列 As String
Private mstr屏蔽列 As String

'=========================================================================================
Private mconIntCol行号 As Integer
Private mconIntCol药名 As Integer
Private mconIntCol商品名 As Integer
Private mconIntCol来源 As Integer
Private mconIntCol基本药物 As Integer
Private mconIntCol序号 As Integer
Private mconIntCol规格 As Integer
Private mconIntCol原产地 As Integer
Private mconIntCol原销期 As Integer
Private mconIntCol比例系数 As Integer
Private mconIntCol产地 As Integer
Private mconIntCol单位 As Integer
Private mconIntCol批号 As Integer
Private mconIntCol生产日期 As Integer
Private mconIntCol效期 As Integer
Private mconIntCol批准文号 As Integer
Private mconIntCol外观 As Integer
Private mconIntCol数量 As Integer
Private mconIntCol冲销数量 As Integer
Private mconintCol成本价 As Integer
Private mconintCol成本金额 As Integer
Private mconIntCol售价 As Integer
Private mconIntCol售价金额 As Integer
Private mconintCol差价 As Integer

Private mconintCol零售价 As Integer
Private mconintCol零售单位 As Integer
Private mconintCol零售金额 As Integer
Private mconintCol零售差价 As Integer

Private mconintCol真实数量 As Integer
Private mconIntCol分批属性 As Integer
Private mconIntCol是否新行 As Integer
Private mconIntCol药品编码和名称 As Integer
Private mconIntCol药品编码 As Integer
Private mconIntCol药品名称 As Integer
Private mconIntCol批次 As Integer
Private Const mconIntColS = 36
'=========================================================================================

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
Private Function CheckStock(ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal Dbl数量 As Double, Optional ByVal intRow As Integer) As Boolean
    Dim lng库房ID As Long
    Dim blnMsg As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim dbl可用数量 As Double
    Dim intCurRow As Integer
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim dbl原填写数量 As Double
    
    On Error GoTo errHandle
    If intRow > 0 Then
        intCurRow = intRow
    Else
        intCurRow = mshBill.Row
    End If
    
    If mint编辑状态 = 2 Then
        '取单据的原始数量
        numUsedCount = 0
        For Each vardrug In mcolUsedCount
            If vardrug(0) = CStr(lng药品ID) & CStr(lng批次) Then
                numUsedCount = vardrug(1)
                Exit For
            End If
        Next
    End If
    
    dbl原填写数量 = IIf(mbln下可用数量, numUsedCount * Val(mshBill.TextMatrix(intCurRow, mconIntCol比例系数)), 0)
    
    '负数入库时检查库存
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    gstrSQL = "Select Nvl(Sum(可用数量),0) 可用数量,Nvl(Sum(实际数量),0) 实际数量 From 药品库存 Where 库房ID=[1] And Nvl(批次,0)=[3] And 性质=1 And 药品ID=[2] "
    Set rsCheck = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查库存是否足够——负数入库货]", lng库房ID, lng药品ID, lng批次)

    If mint库存检查 = 0 Then
        CheckStock = True
    Else
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            dbl可用数量 = rsCheck!可用数量 + Abs(dbl原填写数量)
        Else
            dbl可用数量 = rsCheck!实际数量
        End If
        If Abs(Dbl数量) > Abs(dbl可用数量) Then
            blnMsg = True
        End If
        
        If blnMsg Then
            If mint库存检查 = 1 Then
                If MsgBox("负数入库数量绝对值大于现有的库存数量（当前库存数量为：" & dbl可用数量 / Val(mshBill.TextMatrix(intCurRow, mconIntCol比例系数)) & "），是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                MsgBox "负数入库数量绝对值不能大于现有的库存数量（当前库存数量为：" & dbl可用数量 / Val(mshBill.TextMatrix(intCurRow, mconIntCol比例系数)) & "）！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        CheckStock = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.名称 " _
        & " FROM 药品单据性质 A, 药品入出类别 B " _
        & "Where A.类别id = B.ID " _
      & "AND A.单据 = 4 "
    Call SQLTest(App.Title, "药品其他入库管理", gstrSQL)
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
    Call SQLTest
    If rsDepend.EOF Then
        MsgBox "没有设置药品其他入库的入出类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    Set mrsInOutType = rsDepend
       
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mbln日期提示 = False
    mbln效期提示 = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1302)
    mint时价分段加成方式 = Val(zldatabase.GetPara("时价药品入库采用分段加成", glngSys, 模块号.其他入库, 0))
    mint时价入库时取上次售价 = Val(zldatabase.GetPara("时价药品入库时取上次售价", glngSys, 模块号.其他入库, 0))
        
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
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
        CmdSave.Caption = "冲销(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    mbln效期提示 = True
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub
Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        zlCommFun.PressKey (vbKeyTab)
    End If
    
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
                If MsgBox("如果改变库房，有可能要改变相应药品的单位，且要清除现有单据内容，你是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '处理药品单位改变
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    
                    mlng入库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng入库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
                
                mlng入库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                Call GetDrugDigit(mlng入库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            End If
        End If
        
    End With
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    With mshBill
        .SetFocus
        .Row = 1
        .Col = mconIntCol药名
    End With
        
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol冲销数量) = GetFormat(0, mintNumberDigit)
                .TextMatrix(intRow, mconintCol成本金额) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconintCol差价) = GetFormat(0, mintMoneyDigit)
                
                Call Set时价分批药品零售价(intRow, Val(.TextMatrix(intRow, mconintCol零售价)))
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
                .TextMatrix(intRow, mconIntCol冲销数量) = .TextMatrix(intRow, mconIntCol数量)
                .TextMatrix(intRow, mconintCol成本金额) = GetFormat(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconintCol成本价), mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(.TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol售价), mintMoneyDigit)
                .TextMatrix(intRow, mconintCol差价) = GetFormat(.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconintCol成本金额), mintMoneyDigit)
                
                Call Set时价分批药品零售价(intRow, Val(.TextMatrix(intRow, mconintCol零售价)))
            End If
        Next
    End With
    Call 显示合计金额
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
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
    End Select
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zldatabase.GetPara("简码方式", , , 0))    '默认拼音简码
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

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim intLop As Integer
    
    On Error GoTo ErrHand
    
    '设置排序数据集
    Call SetSortRecord
        
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        mstrTime_End = GetBillInfo(4, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not 检查单价(4, txtNo.Tag, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        For intLop = 1 To mshBill.rows - 1
            If Trim(mshBill.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                '判断已失效药品是否禁止入库
                If CheckLapse(Trim(mshBill.TextMatrix(intLop, mconIntCol效期)), True) = False Then
                    mbln日期提示 = True
                    MsgBox "第" & intLop & "行药品已经失效了，失效药品不能保存！", vbInformation, gstrSysName
                    mshBill.SetFocus
                    mshBill.Row = intLop
                    mshBill.MsfObj.TopRow = intLop
                    mshBill.Col = mconIntCol效期
                    mbln日期提示 = False
                    Exit Sub
                End If
            End If
        Next
        
        blnTrans = True
        gcnOracle.BeginTrans
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If Not SaveCheck Then
            gcnOracle.RollbackTrans: Exit Sub
            Exit Sub
        End If
        
        gcnOracle.CommitTrans
        
        If Val(zldatabase.GetPara("审核打印", glngSys, 模块号.其他入库)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
        
        If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
            MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
            txt摘要.SetFocus
            Exit Sub
        End If
        
        If mblnChange = False Then
            MsgBox "请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("你确实要冲销单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 2 Then
        If Not 检查单价(4, txtNo.Tag, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新售价，程序将自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint编辑状态 = 1 Then '新增保存时，判断售价是否已经更新
        If 检查售价 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
            
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zldatabase.GetPara("存盘打印", glngSys, 模块号.其他入库)) = 1 Then
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
    SetEdit
    
    txt摘要.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshBill()
    '以最新价格最新单据相关数据，用于单据审核时
    Dim lngRow As Long, lngRows As Long, lng药品ID As Long
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    Dim rsPrice As New ADODB.Recordset
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
        
    gstrSQL = " Select 收费细目ID,nvl(现价,0) 现价 From 收费价目 " & _
            " Where (终止日期 Is NULL Or sysdate Between 执行日期 And nvl(终止日期,to_date('3000-01-01','yyyy-MM-dd')))"
    gstrSQL = "Select A.序号,A.药品ID,B.现价 From 药品收发记录 A,(" & gstrSQL & ") B,收费项目目录 C" & _
            " Where A.单据=4 And A.NO=[1] And A.药品ID=B.收费细目ID And C.ID=B.收费细目ID And Round(A.零售价," & intPriceDigit & ")<>Round(B.现价," & intPriceDigit & ") And Nvl(C.是否变价,0)=0" & _
            " Union All " & _
            " Select A.序号, A.药品id, B.零售价 现价 " & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C, 药品规格 D " & _
            " Where A.单据 = 4 And A.NO = [1] And C.ID = A.药品id And Round(A.零售价, " & intPriceDigit & ") <> Round(B.零售价, " & intPriceDigit & ") And " & _
            " Nvl(C.是否变价, 0) = 1 And D.药品id = A.药品id And B.性质 = 1 And B.库房id = A.库房id And B.药品id = A.药品id And " & _
            " Nvl(B.批次, 0) = Nvl(A.批次, 0) And Nvl(A.批次, 0) > 0 And Nvl(B.零售价, 0) > 0 " & _
            " Order by 药品id,序号"
    Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[取当前价格]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        lng药品ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng药品ID <> 0 Then
            rsPrice.Filter = "药品ID=" & lng药品ID
            If rsPrice.RecordCount <> 0 Then
                '以当前最新价格最新单据相关数据（单价、零售金额、差价）
                dbl零售价 = rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数))
                dbl成本价 = Val(mshBill.TextMatrix(lngRow, mconintCol成本价))
                Dbl数量 = Val(mshBill.TextMatrix(lngRow, mconIntCol数量))
                dbl成本金额 = dbl成本价 * Dbl数量
                dbl零售金额 = dbl零售价 * Dbl数量
                dbl差价 = dbl零售金额 - dbl成本金额
                
                mshBill.TextMatrix(lngRow, mconIntCol售价) = GetFormat(dbl零售价, intPriceDigit)
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = GetFormat(dbl零售金额, mintMoneyDigit)
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


Private Sub Form_Load()
    Dim i As Integer, j As Integer
    
    marrFrom = Array()
    marrInitGrid = Array()
    mintBatchNoLen = GetBatchNoLen()
    mbln加价率 = Get加价率
    mintSelectStock = Val(zldatabase.GetPara("是否选择库房", glngSys, 模块号.其他入库))
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    mint取上次成本价方式 = Val(zldatabase.GetPara("取上次成本价方式", glngSys, 模块号.外购入库))
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    mblnUpdate = False
    
    On Error GoTo errHandle
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品其他入库管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    Set mrs分段加成 = Nothing
    If mint时价分段加成方式 = 1 Then
        gstrSQL = "select 序号, 最低价, 最高价, 加成率, 差价额, 说明, 类型 from 药品加成方案 order by 序号"
        Set mrs分段加成 = zldatabase.OpenSQLRecord(gstrSQL, "查询分段加成")
    End If
    mshBill.Value = Format(zldatabase.Currentdate, "YYYY-MM-DD")
    
    With cboType
        .Clear
        Do While Not mrsInOutType.EOF
            .AddItem mrsInOutType.Fields(1)
            .ItemData(.NewIndex) = mrsInOutType.Fields(0)
            mrsInOutType.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    mlng入库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng入库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(4, mstr单据号)
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    RestoreWinState Me, App.ProductName, MStrCaption
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next
    
    mshBill.ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 1100, 0)
    
    If mintUnit = mconint售价单位 Then
        mshBill.ColWidth(mconintCol零售价) = 0
        mshBill.ColWidth(mconintCol零售单位) = 0
        mshBill.ColWidth(mconintCol零售金额) = 0
        mshBill.ColWidth(mconintCol零售差价) = 0
    Else
        mshBill.ColWidth(mconintCol零售价) = 0
        mshBill.ColWidth(mconintCol零售单位) = 0
        mshBill.ColWidth(mconintCol零售金额) = 0
        mshBill.ColWidth(mconintCol零售差价) = 0
        
        If InStr(1, "|" & mstr屏蔽列 & "|", "|零售价|") = 0 Then mshBill.ColWidth(mconintCol零售价) = 1000
        If InStr(1, "|" & mstr屏蔽列 & "|", "|零售单位|") = 0 Then mshBill.ColWidth(mconintCol零售单位) = 1000
        If InStr(1, "|" & mstr屏蔽列 & "|", "|零售金额|") = 0 Then mshBill.ColWidth(mconintCol零售金额) = 1000
        If InStr(1, "|" & mstr屏蔽列 & "|", "|零售差价|") = 0 Then mshBill.ColWidth(mconintCol零售差价) = 1000
    End If
    
    '根据人员权限判断，是否显示成本价
    If InStr(1, "|" & mstr屏蔽列 & "|", "|成本价|") = 0 Then mshBill.ColWidth(mconintCol成本价) = IIf(mblnViewCost, 1000, 0)
    If InStr(1, "|" & mstr屏蔽列 & "|", "|成本金额|") = 0 Then mshBill.ColWidth(mconintCol成本金额) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstr屏蔽列 & "|", "|差价|") = 0 Then mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstr屏蔽列 & "|", "|零售差价|") = 0 Then mshBill.ColWidth(mconintCol零售差价) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconintCol真实数量) = 0
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    mblnChange = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim str批次 As String
    Dim strArray As String
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPriceDigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    
    '库房
    strOrder = zldatabase.GetPara("排序", glngSys, 模块号.其他入库)
    strCompare = Mid(strOrder, 1, 1)
    
    On Error GoTo errHandle
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
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 6
                
            Call initGrid
            If mint编辑状态 = 4 Then
                gstrSQL = "select b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 4 and a.no=[1]"
                Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
                
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
                    strUnitQuantity = "F.计算单位 AS 售价单位,F.计算单位 AS 单位, A.填写数量 AS 数量,b.指导批发价 as 指导批发价, a.成本价,A.零售价,1 as 比例系数,"
                Case mconint门诊单位
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 数量,b.指导批发价*B.门诊包装 as 指导批发价 , a.成本价*B.门诊包装 as 成本价,A.零售价*B.门诊包装 as 零售价 ,B.门诊包装 as 比例系数,"
                Case mconint住院单位
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 数量,b.指导批发价*B.住院包装 as 指导批发价 , a.成本价*B.住院包装 as 成本价,A.零售价*B.住院包装 as 零售价 ,  B.住院包装 as 比例系数,"
                Case mconint药库单位
                    strUnitQuantity = "F.计算单位 AS 售价单位,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 数量,b.指导批发价*B.药库包装 as 指导批发价 , a.成本价*B.药库包装 as 成本价,A.零售价*B.药库包装 as 零售价 ,B.药库包装 as 比例系数,"
            End Select
            
            If mint编辑状态 <> 6 Then
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' ||F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    " B.药品来源,B.基本药物,F.规格,F.产地 AS 原产地,A.产地, A.批号,A.批次," & _
                    " B.最大效期,A.效期," & strUnitQuantity & " A.成本金额, " & _
                    " A.零售金额, A.差价,nvl(B.加成率,0)/100 AS 加成率,F.是否变价,B.药房分批 AS 药房分批核算, " & _
                    " A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,G.名称 AS 部门,A.入出类别ID,A.生产日期,A.批准文号,A.外观, Nvl(A.用法, 0) As 金额差 " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目别名 E,收费项目目录 F,部门表 G " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID AND A.库房ID=G.ID" & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1 " & _
                    " AND A.记录状态 =[2] " & _
                    " AND A.单据 = 4 AND A.NO = [1])" & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.药品ID,A.序号,'[' ||F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名, " & _
                    " B.药品来源,B.基本药物,F.规格,F.产地 AS 原产地,A.产地, A.批号,A.批次," & _
                    " B.最大效期,A.效期," & strUnitQuantity & " A.成本金额, " & _
                    " 0 零售金额,0 差价,nvl(B.加成率,0)/100 AS 加成率,F.是否变价,B.药房分批 AS 药房分批核算, " & _
                    " A.库房ID,G.名称 AS 部门,A.入出类别ID, A.生产日期,A.批准文号,A.外观,A.填写数量 真实数量,A.金额差 " & _
                    " FROM " & _
                    "     (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,SUM(成本金额) AS 成本金额,Sum(To_Number(Nvl(用法, 0))) As 金额差," & _
                    "     药品ID,序号,产地, 批号,nvl(批次,0) as 批次,效期,扣率,成本价,零售价,库房ID,入出类别ID,X.生产日期,X.批准文号,X.外观" & _
                    "     FROM 药品收发记录 X " & _
                    "     WHERE NO=[1] AND 单据=4  " & _
                    "     GROUP BY 药品ID,序号,产地, 批号,nvl(批次,0),效期,扣率,成本价,零售价,库房ID,入出类别ID,X.生产日期,X.批准文号,X.外观" & _
                    "     HAVING SUM(实际数量)<>0 ) A," & _
                    "     药品规格 B,收费项目别名 E ,收费项目目录 F,部门表 G " & _
                    " WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID AND A.库房ID=G.ID" & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1 )" & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, mint记录状态)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            Select Case mint编辑状态
                Case 2, 6
                    Txt填制人 = UserInfo.用户姓名
                    Txt填制日期 = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    If mint编辑状态 = 2 Then
                        Txt审核人 = ""
                        Txt审核日期 = ""
                    Else
                        Txt审核人 = UserInfo.用户姓名
                        Txt审核日期 = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    End If
                Case Else
                    Txt填制人 = rsInitCard!填制人
                    Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                    Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                    Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            If mint编辑状态 <> 6 Then
                txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            Else
                txt摘要.Text = Get摘要(mstr单据号)
            End If
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!入出类别ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            
            If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = intRow + 1
                    'intRow = rsInitCard!序号
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
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    .TextMatrix(intRow, mconIntCol批号) = IIf(IsNull(rsInitCard!批号), "", rsInitCard!批号)
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol数量) = GetFormat(rsInitCard!数量, intNumberDigit)
                    .TextMatrix(intRow, mconIntCol生产日期) = IIf(IsNull(rsInitCard!生产日期), "", rsInitCard!生产日期)
                    If rsInitCard!数量 <> 0 Then
                        .TextMatrix(intRow, mconintCol成本价) = GetFormat(rsInitCard!成本价, intCostDigit)
                    Else
                        .TextMatrix(intRow, mconintCol成本价) = IIf(mintUnit = mconint药库单位, "0.00000", "0.0000000")
                    End If
                    .TextMatrix(intRow, mconintCol成本金额) = GetFormat(IIf(mint编辑状态 = 6, 0, rsInitCard!成本金额), intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol售价) = GetFormat(rsInitCard!零售价, intPriceDigit)
                    .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(rsInitCard!零售金额, intMoneyDigit)
                    .TextMatrix(intRow, mconintCol差价) = GetFormat(rsInitCard!差价, intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "!", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconIntCol外观) = IIf(IsNull(rsInitCard!外观), "", rsInitCard!外观)
                    .TextMatrix(intRow, mconIntCol是否新行) = "否"
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol冲销数量) = GetFormat(0, intNumberDigit)
                        .TextMatrix(intRow, mconintCol真实数量) = rsInitCard!真实数量
                    End If
                        
                    .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!加成率 & "||" & IIf(IsNull(rsInitCard!是否变价), 0, rsInitCard!是否变价) & "||" & IIf(IsNull(rsInitCard!药房分批核算), 0, rsInitCard!药房分批核算)
                        
                    '分批属性
                    Call Get药品分批属性(intRow)
                    
                    '时价分批药品处理，需要重算界面的售价、售价金额、差价
                    If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
                        If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                            .TextMatrix(intRow, mconintCol零售单位) = rsInitCard!售价单位
                            .TextMatrix(intRow, mconintCol零售价) = GetFormat(rsInitCard!零售价 / Val(rsInitCard!比例系数), gtype_UserDrugDigits.Digit_零售价)
                            .TextMatrix(intRow, mconintCol零售金额) = GetFormat(rsInitCard!零售金额, intMoneyDigit)
                            .TextMatrix(intRow, mconintCol零售差价) = GetFormat(rsInitCard!差价, intMoneyDigit)
                            
                            If mint编辑状态 <> 6 Then
                                '不是冲销时
                                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(rsInitCard!金额差), intMoneyDigit)
                                .TextMatrix(intRow, mconintCol差价) = GetFormat(Val(.TextMatrix(intRow, mconintCol零售差价)) - Val(rsInitCard!金额差), intMoneyDigit)
                                .TextMatrix(intRow, mconIntCol售价) = GetFormat(Val(.TextMatrix(intRow, mconIntCol售价金额)) / Val(rsInitCard!数量), intPriceDigit)
                            Else
                                '冲销时
                                .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(0, intMoneyDigit)
                                .TextMatrix(intRow, mconintCol差价) = GetFormat(0, intMoneyDigit)
                                .TextMatrix(intRow, mconIntCol售价) = GetFormat((Val(.TextMatrix(intRow, mconintCol零售价)) * Val(rsInitCard!比例系数) * Val(rsInitCard!数量) - Val(rsInitCard!金额差)) / Val(rsInitCard!数量), intPriceDigit)
                            End If
                        End If
                    End If
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!药品id & "0") Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str批次 = rsInitCard!药品id & "0"
                        strArray = numUseAbleCount + IIf(IsNull(rsInitCard!数量), "0", rsInitCard!数量)
                        mcolUsedCount.Add Array(str批次, strArray), str批次
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    
    SetEdit         '设置编辑属性
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get摘要(ByVal strNo As String) As String
    '获取新的摘要
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
         '冲销(取最后一次冲销的摘要)
    gstrSQL = "Select 摘要 From 药品收发记录 Where 单据=4 And No=[1] and (记录状态 =1 or mod(记录状态,3)=0) Order By 审核日期 Desc "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取摘要信息", strNo)
    
    If Not rsTemp.EOF Then
        Get摘要 = Nvl(rsTemp!摘要)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = IIf(mint编辑状态 = 6, 5, 0)
            Next
            If mint编辑状态 = 6 Then
                .ColData(mconIntCol药名) = 0
                .ColData(mconIntCol冲销数量) = 4
                txt摘要.Enabled = True
            End If
            
            cboStock.Enabled = False
            cboType.Enabled = False
            
            If mint编辑状态 <> 6 Then
                txt摘要.Enabled = False
            End If
        Else
            .ColData(0) = 5
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol序号) = 5
            .ColData(mconIntCol规格) = 5
            If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
                .ColData(mconIntCol产地) = 1
            Else
                .ColData(mconIntCol产地) = 5
            End If
            .ColData(mconIntCol单位) = 5
            .ColData(mconIntCol批号) = 4
            .ColData(mconIntCol批次) = 5
            .ColData(mconIntCol生产日期) = 2
            .ColData(mconIntCol效期) = 5
            .ColData(mconIntCol数量) = 4
            .ColData(mconintCol成本价) = 4
            .ColData(mconintCol成本金额) = 4
            .ColData(mconIntCol售价) = 5
            .ColData(mconIntCol售价金额) = 5
            .ColData(mconintCol差价) = 5
            
            .ColData(mconIntCol原产地) = 5
            .ColData(mconIntCol原销期) = 5
            .ColData(mconIntCol比例系数) = 5
            .ColData(mconIntCol批准文号) = 4
            .ColData(mconIntCol外观) = 1
            
            .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
            .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
            .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
            .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
            .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
            .ColAlignment(mconIntCol批次) = flexAlignLeftCenter
            .ColAlignment(mconIntCol生产日期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
            .ColAlignment(mconIntCol数量) = flexAlignRightCenter
            .ColAlignment(mconintCol成本价) = flexAlignRightCenter
            .ColAlignment(mconintCol成本金额) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价) = flexAlignRightCenter
            .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
            .ColAlignment(mconintCol差价) = flexAlignRightCenter
            
            If mintSelectStock = 0 Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            cboType.Enabled = True
            txt摘要.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        Call SetColumnByUserDefine
        .TextMatrix(0, mconIntCol行号) = ""
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconIntCol生产日期) = "生产日期"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol外观) = "外观"
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol冲销数量) = "冲销数量"
        .TextMatrix(0, mconintCol成本价) = "成本价"
        .TextMatrix(0, mconintCol成本金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconintCol零售价) = "零售价"
        .TextMatrix(0, mconintCol零售单位) = "零售单位"
        .TextMatrix(0, mconintCol零售金额) = "零售金额"
        .TextMatrix(0, mconintCol零售差价) = "零售差价"
        .TextMatrix(0, mconIntCol原产地) = "原产地"
        .TextMatrix(0, mconIntCol原销期) = "原效期"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconintCol真实数量) = "真实数量"
        .TextMatrix(0, mconIntCol分批属性) = "分批属性"
        .TextMatrix(0, mconIntCol是否新行) = "是否新行"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        
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
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconIntCol生产日期) = 1000
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconIntCol外观) = 1000
        .ColWidth(mconIntCol数量) = 1100
        .ColWidth(mconIntCol冲销数量) = IIf(mint编辑状态 = 6, 1100, 0)
        .ColWidth(mconintCol成本价) = 1000
        .ColWidth(mconintCol成本金额) = 900
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = 800
        .ColWidth(mconintCol零售价) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售单位) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售金额) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconintCol零售差价) = IIf(mintUnit = mconint售价单位, 0, 1000)
        .ColWidth(mconIntCol原产地) = 0
        .ColWidth(mconIntCol原销期) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconintCol真实数量) = 0
        .ColWidth(mconIntCol分批属性) = 0
        .ColWidth(mconIntCol是否新行) = 0
        
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
                
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol药名) = 1
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol规格) = 5
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            .ColData(mconIntCol产地) = 1
        Else
            .ColData(mconIntCol产地) = 5
        End If
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 4
        .ColData(mconIntCol批次) = 4
        .ColData(mconIntCol生产日期) = 2
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconIntCol外观) = 5
        .ColData(mconIntCol数量) = 4
        .ColData(mconIntCol冲销数量) = 4
        .ColData(mconintCol成本价) = 4
        .ColData(mconintCol成本金额) = 4
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        .ColData(mconintCol零售价) = 5
        .ColData(mconintCol零售单位) = 5
        .ColData(mconintCol零售金额) = 5
        .ColData(mconintCol零售差价) = 5
        .ColData(mconIntCol原产地) = 5
        .ColData(mconIntCol原销期) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconintCol真实数量) = 5
        .ColData(mconIntCol是否新行) = 5
        
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批次) = flexAlignLeftCenter
        .ColAlignment(mconIntCol生产日期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol外观) = flexAlignLeftCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconintCol成本价) = flexAlignRightCenter
        .ColAlignment(mconintCol成本金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        .ColAlignment(mconintCol零售价) = flexAlignRightCenter
        .ColAlignment(mconintCol零售单位) = flexAlignRightCenter
        .ColAlignment(mconintCol零售金额) = flexAlignRightCenter
        .ColAlignment(mconintCol零售差价) = flexAlignRightCenter
        .ColAlignment(mconintCol真实数量) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
    End With
    txt摘要.MaxLength = GetLength("药品收发记录", "摘要")
    Call SetColumnByUserDefine
End Sub

Private Sub SetColumnByUserDefine()
    Dim intCol As Integer
    Dim arr总列, arr可设置列
    Dim str总列 As String, str可设置列 As String
    Dim intColumns As Integer
    Dim intCols As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected
    
    On Error GoTo ErrHand
    mstr选择列 = zldatabase.GetPara("选择列", glngSys, 模块号.其他入库)
    mstr屏蔽列 = zldatabase.GetPara("屏蔽列", glngSys, 模块号.其他入库)
    
    str总列 = "药名|药品来源|基本药物|规格|产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|" & _
                        "售价|售价金额|差价|零售价|零售单位|零售金额|零售差价|批准文号|外观"
                        
    '主要是零售价这几列默认为不显示，所以做一些特殊处理
    If mstr选择列 <> "" Then
        If mstr屏蔽列 <> "" Then
            str可设置列 = mstr选择列 & "|" & mstr屏蔽列
        Else
            str可设置列 = mstr选择列
        End If
        arr总列 = Split(str总列, "|")
        arr可设置列 = Split(str可设置列, "|")
        If UBound(arr总列) <> UBound(arr可设置列) Or InStr(1, "|" & mstr屏蔽列 & "|", "|产地|") <> 0 Or InStr(1, "|" & mstr选择列 & "|", "|产地|") = 0 Or InStr(1, "|" & mstr屏蔽列 & "|", "|采购价|") <> 0 Or InStr(1, "|" & mstr选择列 & "|", "|采购价|") <> 0 Then
            mstr选择列 = "药名|药品来源|基本药物|规格|产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|" & _
                                                    "售价|售价金额|差价|批准文号|外观"
            mstr屏蔽列 = "零售价|零售单位|零售金额|零售差价"
            zldatabase.SetPara "选择列", mstr选择列, glngSys, 模块号.其他入库
            zldatabase.SetPara "屏蔽列", mstr屏蔽列, glngSys, 模块号.其他入库
        End If
    Else
        mstr选择列 = "药名|药品来源|基本药物|规格|产地|批号|生产日期|效期|单位|数量|冲销数量|成本价|成本金额|" & _
                                                    "售价|售价金额|差价|批准文号|外观"
        mstr屏蔽列 = "零售价|零售单位|零售金额|零售差价"
        zldatabase.SetPara "选择列", mstr选择列, glngSys, 模块号.其他入库
        zldatabase.SetPara "屏蔽列", mstr屏蔽列, glngSys, 模块号.其他入库
    End If

'    mstr屏蔽列 = "|" & mstr屏蔽列 & "|"
    With mshBill
        For intCol = 1 To .Cols - 1
            If InStr("|" & mstr屏蔽列 & "|", "|" & .TextMatrix(0, intCol) & "|") > 0 Then
                .ColWidth(intCol) = 0
                .ColData(intCol) = 5
            End If
        Next
    End With
    
    strColumn_All = "药名,2|药品来源,4|基本药物,5|规格,7|产地,11|单位,12|批号,13|生产日期,14|效期,15|批准文号,16|外观,17|数量,18|冲销数量,19|成本价,20|成本金额,21|" & _
                    "售价,22|售价金额,23|差价,24|零售价,25|零售单位,26|零售金额,27|零售差价,28"

    '先装入缺省设置
    mconIntCol行号 = 1
    mconIntCol药名 = 2
    mconIntCol商品名 = 3
    mconIntCol来源 = 4
    mconIntCol基本药物 = 5
    mconIntCol序号 = 6
    mconIntCol规格 = 7
    mconIntCol原产地 = 8
    mconIntCol原销期 = 9
    mconIntCol比例系数 = 10
    mconIntCol产地 = 11
    mconIntCol单位 = 12
    mconIntCol批号 = 13
    mconIntCol生产日期 = 14
    mconIntCol效期 = 15
    mconIntCol批准文号 = 16
    mconIntCol外观 = 17
    mconIntCol数量 = 18
    mconIntCol冲销数量 = 19
    mconintCol成本价 = 20
    mconintCol成本金额 = 21
    mconIntCol售价 = 22
    mconIntCol售价金额 = 23
    mconintCol差价 = 24
    mconintCol零售价 = 25
    mconintCol零售单位 = 26
    mconintCol零售金额 = 27
    mconintCol零售差价 = 28
    mconintCol真实数量 = 29
    mconIntCol分批属性 = 30
    mconIntCol是否新行 = 31
    mconIntCol药品编码和名称 = 32
    mconIntCol药品编码 = 33
    mconIntCol药品名称 = 34
    mconIntCol批次 = 35
    
    mintLastCol = 35
    '根据用户设置调整列顺序
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(mstr选择列, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstr屏蔽列, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    
    Exit Sub
ErrHand:
    MsgBox "恢复列设置时发生错误，请重新进行列设置！", vbInformation, gstrSysName
End Sub
Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str列名
    Case "药名"
        mconIntCol药名 = intValue
    Case "药品来源"
        mconIntCol来源 = intValue
    Case "基本药物"
        mconIntCol基本药物 = intValue
    Case "规格"
        mconIntCol规格 = intValue
    Case "产地"
        mconIntCol产地 = intValue
    Case "单位"
        mconIntCol单位 = intValue
    Case "批号"
        mconIntCol批号 = intValue
    Case "生产日期"
        mconIntCol生产日期 = intValue
    Case "效期"
        mconIntCol效期 = intValue
    Case "批准文号"
        mconIntCol批准文号 = intValue
    Case "外观"
        mconIntCol外观 = intValue
    Case "数量"
        mconIntCol数量 = intValue
    Case "冲销数量"
        mconIntCol冲销数量 = intValue
    Case "成本价"
        mconintCol成本价 = intValue
    Case "成本金额"
        mconintCol成本金额 = intValue
    Case "售价"
        mconIntCol售价 = intValue
    Case "售价金额"
        mconIntCol售价金额 = intValue
    Case "差价"
        mconintCol差价 = intValue
    Case "零售价"
        mconintCol零售价 = intValue
    Case "零售单位"
        mconintCol零售单位 = intValue
    Case "零售金额"
        mconintCol零售金额 = intValue
    Case "零售差价"
        mconintCol零售差价 = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Sub Set时价分批药品零售价(ByVal intRow As Integer, ByVal dblPrice As Double)
    Dim Dbl数量 As Double

    With mshBill
        If .TextMatrix(intRow, mconIntCol原销期) = "" Then Exit Sub
        If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) <> 1 Or Val(.TextMatrix(intRow, mconIntCol分批属性)) <> 1 Then Exit Sub
        
       .TextMatrix(intRow, mconintCol零售价) = GetFormat(dblPrice, gtype_UserDrugDigits.Digit_零售价) '零售价字段本来就是最小单位，因此不能用药品卫材精度设置进行控制，直接按照7位进行显示
        
        If mint编辑状态 = 6 Then
            Dbl数量 = Val(.TextMatrix(intRow, mconIntCol冲销数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
        Else
            Dbl数量 = Val(.TextMatrix(intRow, mconIntCol数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
        End If
        If Val(.TextMatrix(intRow, mconintCol成本价)) = Val(.TextMatrix(intRow, mconIntCol售价)) Then
            '通过技术处理零差价销售管理外购入库，防止出现除不尽尽快
            .TextMatrix(intRow, mconintCol零售金额) = .TextMatrix(intRow, mconIntCol售价金额)
        Else
            .TextMatrix(intRow, mconintCol零售金额) = GetFormat(Dbl数量 * Val(.TextMatrix(intRow, mconintCol零售价)), mintMoneyDigit)
        End If
        .TextMatrix(intRow, mconintCol零售差价) = GetFormat(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconintCol成本金额)), mintMoneyDigit)
    End With
End Sub

Private Sub Get药品分批属性(intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strsql As String
    Dim int分批属性 As Integer      '0-不分批;1-分批
    Dim int药库分批 As Integer      '0-不分批;1-分批
    Dim int药房分批 As Integer      '0-不分批;1-分批
    Dim bln是否具有药房性质 As Boolean  'True-具有药房性质;False-不具有药房性质
    
    If Val(mshBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo errHandle
    strsql = "SELECT NVL(药库分批, 0) 药库分批,NVL(药房分批, 0) 药房分批 " & _
            " From 药品规格 WHERE 药品ID = [1] "
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "取药品库房分批属性", Val(mshBill.TextMatrix(intBillRow, 0)))
    
    If rsTemp.RecordCount > 0 Then
        int药库分批 = rsTemp!药库分批
        int药房分批 = rsTemp!药房分批
    End If
    
    If int药房分批 = 1 Then     '如果药房分批，则分批属性为1
        int分批属性 = 1
    Else
        If int药库分批 = 1 Then
            strsql = "SELECT 部门ID From 部门性质说明 " & _
                    " WHERE ((工作性质 LIKE '%药房') OR (工作性质 LIKE '制剂室')) AND 部门ID = [1] "
            Set rsTemp = zldatabase.OpenSQLRecord(strsql, "取部门性质", cboStock.ItemData(Me.cboStock.ListIndex))
            
            bln是否具有药房性质 = (rsTemp.RecordCount > 0)
                    
            If bln是否具有药房性质 Then
                int分批属性 = 0
            Else
                int分批属性 = 1
            End If
        End If
    End If
    
    mshBill.TextMatrix(intBillRow, mconIntCol分批属性) = int分批属性
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
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
    
    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
    
    LblType.Left = cboType.Left - LblType.Width - 100
    
    
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
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
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
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品其他入库管理", "药品名称显示方式", mintDrugNameShow)
    
    If mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        zlPlugIn_Unload mobjPlugIn
        Exit Sub
    End If
    If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
    zlPlugIn_Unload mobjPlugIn
End Sub



Private Function SaveCheck() As Boolean
    mblnSave = False
    SaveCheck = False
    
    Dim n As Integer
    Dim m As Integer
    Dim dbl合计数量 As Double
    Dim lng药品ID As Long
    
    '如果入库药品合计数量小于0，则要做库存检查(主要对不分批药品负数入库做校验)
    With mshBill
        For n = 1 To .rows - 1
            If Val(.TextMatrix(n, 0)) <> 0 Then
                dbl合计数量 = 0
                lng药品ID = Val(.TextMatrix(n, 0))
                For m = 1 To .rows - 1
                    If Val(.TextMatrix(m, 0)) <> 0 And lng药品ID = Val(.TextMatrix(m, 0)) Then
                        dbl合计数量 = dbl合计数量 + Val(.TextMatrix(m, mconIntCol数量)) * Val(.TextMatrix(m, mconIntCol比例系数))
                    End If
                Next
                
                '合计数量为负数时才校验库存
                If dbl合计数量 < 0 Then
                    If Not CheckStock(lng药品ID, 0, dbl合计数量, n) Then
                        MsgBox "药品[" & .TextMatrix(n, mconIntCol药名) & "]库存不足，不能负数入库。"
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    gstrSQL = "zl_药品其他入库_Verify('" & txtNo.Tag & "','" & UserInfo.用户姓名 & "')"
    
    On Error GoTo errHandle
    Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '外挂功能
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, Val(cboStock.ItemData(cboStock.ListIndex)), txtNo.Tag, 单据号.其他入库)
    
    Exit Function
errHandle:
    'MsgBox "审核失败！", vbInformation, gstrSysName
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

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
    mshBill.Value = Format(zldatabase.Currentdate, "YYYY-MM-DD")
    mshBill.TextMatrix(Row, mconIntCol是否新行) = "是"
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
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
    Dim intRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    
    On Error GoTo errHandle
    Select Case mshBill.Col
    Case mconIntCol药名
        Dim RecReturn As Recordset
        
        mshBill.CmdEnable = False
'        Set RecReturn = Frm药品选择器.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex))
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        End If
        
        Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
        
        mshBill.CmdEnable = True
        If RecReturn.RecordCount > 0 Then
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intRow = .Row
                    .TextMatrix(intRow, mconIntCol行号) = .Row
                    SetColValue .Row, RecReturn!药品id, _
                        "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                        Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, IIf(IsNull(RecReturn!规格), "", RecReturn!规格), _
                        IIf(IsNull(RecReturn!产地), "", RecReturn!产地), Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                        IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), RecReturn!指导批发价 * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                        IIf(IsNull(RecReturn!产地), "!", RecReturn!产地), RecReturn!最大效期, Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                        RecReturn!时价, RecReturn!药房分批, RecReturn!指导差价率 / 100, _
                        IIf(IsNull(RecReturn!生产日期), "", Format(RecReturn!生产日期, "yyyy-mm-dd")), RecReturn!售价单位
'                    If .TextMatrix(.Row, mconIntCol原产地) = "!" Then
'                        .Col = mconIntCol产地
'                    Else
'                        .Col = mconIntCol批号
'                    End If
                    
                    .Col = GetNextEnableCol(mconIntCol药名)
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End With
            RecReturn.Close
        End If
    Case mconIntCol产地
        Dim rsProvider As Recordset
        Dim vRect As RECT, blnCancel As Boolean
        vRect = GetControlRect(mshBill.MsfObj.hWnd)
        
        gstrSQL = "Select 编码 as id,名称,简码 From 药品生产商 Where 站点 = [1] Or 站点 is Null Order By 编码"
        Set rsProvider = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
        True, vRect.Right / 2, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then

            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = rsProvider!名称
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                        Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
            End If
        End If
    Case mconIntCol外观
        Dim rs外观 As New Recordset
                    
        gstrSQL = "Select 编码,名称,简码 From 药品外观 Order By 编码"
        Set rs外观 = zldatabase.OpenSQLRecord(gstrSQL, "药品外观")
                
        If rs外观.EOF Then
            rs外观.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs外观
            .StrNode = "所有药品外观"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol外观) = .CurrentName
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
        If strkey = "" Then
            strkey = .TextMatrix(.Row, .Col)
        End If
        
        If .Col = mconIntCol数量 Or .Col = mconIntCol冲销数量 Or .Col = mconintCol成本价 Or .Col = mconIntCol售价 Or .Col = mconintCol零售价 Or .Col = mconintCol成本金额 Then
            Select Case .Col
                Case mconIntCol数量, mconIntCol冲销数量
                    intDigit = mintNumberDigit
                Case mconintCol成本价
                   intDigit = mintCostDigit
                Case mconIntCol售价
                    intDigit = mintPriceDigit
                Case mconintCol零售价
                    intDigit = gtype_UserDrugDigits.Digit_零售价
                Case mconintCol成本金额
                    intDigit = mintMoneyDigit
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
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim lngRow As Long
    Dim strxq As String
    Dim dblTemp售价 As Double
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '重新计算零售价、差价
                dblTemp售价 = Val(.TextMatrix(lngRow, mconintCol成本价)) * (1 + (Val(Txt加价率) / 100))
                .TextMatrix(lngRow, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mconintCol成本价)), Val(Txt加价率) / 100, dblTemp售价, lngRow), mintPriceDigit)
                .TextMatrix(lngRow, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(lngRow, mconIntCol售价)) * Val(.TextMatrix(lngRow, mconIntCol数量)), mintMoneyDigit)
                .TextMatrix(lngRow, mconintCol差价) = GetFormat(IIf(.TextMatrix(lngRow, mconIntCol售价金额) = "", 0, .TextMatrix(lngRow, mconIntCol售价金额)) - IIf(.TextMatrix(lngRow, mconintCol成本金额) = "", 0, .TextMatrix(lngRow, mconintCol成本金额)), mintMoneyDigit)
                PicInput.Visible = False
            End If
        End If
        SetInputFormat .Row
        
        'Modified by zyb 2002-10-30
        If Not (.Col = mconintCol成本价 Or .Col = mconintCol成本金额) Then PicInput.Visible = False
        If .Col = mconintCol成本金额 And PicInput.Visible Then Txt加价率.SetFocus: Exit Sub
        
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mconIntCol产地
                OpenIme GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
                .txtCheck = False
                .MaxLength = 60
                .TxtSetFocus
                
            Case mconIntCol批号
                .txtCheck = False
                '.TextMask = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                .MaxLength = mintBatchNoLen
            Case mconIntCol生产日期
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol生产日期)) = "" Then
                                strxq = TranNumToDate(strxq)
                                If Trim(strxq) = "" Then Exit Sub
                                .TextMatrix(.Row, mconIntCol生产日期) = Format(strxq, "yyyy-mm-dd")
                            End If
                         End If
                    End If
                End If
            Case mconIntCol效期
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If Trim(.TextMatrix(.Row, mconIntCol原销期)) = "" Then
                    Exit Sub
                End If
                If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0) = "0" Then
                    Exit Sub
                End If
                If .TextMatrix(.Row, mconIntCol生产日期) <> "" Then
                    If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol生产日期))
                    End If
                ElseIf .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                                strxq = TranNumToDate(strxq)
                            Else
                                Exit Sub
                            End If
                        Else
                            strxq = ""
                        End If
                    Else
                        strxq = ""
                    End If
                End If
                If Trim(strxq) = "" Then Exit Sub
                .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                    '换算为有效期
                    .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                End If
                
'                Call CheckLapse(.TextMatrix(.Row, mconIntCol效期))
            Case mconintCol成本价
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconintCol成本金额
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
                
            Case mconIntCol数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
            Case mconIntCol批准文号
                .txtCheck = False
                .MaxLength = 40
            Case mconIntCol售价
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconIntCol外观
                .txtCheck = False
                .MaxLength = 100
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim dbl加成率 As Double
    Dim strUnitQuantity As String
    Dim dbl指导零售价 As Double
    Dim rsTemp As ADODB.Recordset
    Dim strxq As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim dblTemp售价 As Double
    
    intOldRow = mshBill.Row
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        
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
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex), , strkey, sngLeft, sngTop)
                    
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, _
                               IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, IIf(IsNull(RecReturn!规格), "", RecReturn!规格), _
                               IIf(IsNull(RecReturn!产地), "", RecReturn!产地), Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                               IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), RecReturn!指导批发价 * Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                               IIf(IsNull(RecReturn!产地), "!", RecReturn!产地), RecReturn!最大效期, Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), RecReturn!时价, _
                               RecReturn!药房分批, RecReturn!指导差价率 / 100, IIf(IsNull(RecReturn!生产日期), "", Format(RecReturn!生产日期, "yyyy-mm-dd")), RecReturn!售价单位) = False Then
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        Cancel = True
                    End If
                End If
                Call 提示库存数
                'End If
            Case mconIntCol产地
                '如果找不到对应的产地，则以输入做为产地
                Dim rsProvider As ADODB.Recordset
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, .Col) = ""
                        .Text = " "
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    
                    Exit Sub
                Else
                    Dim rs产地 As New ADODB.Recordset
                    
                    gstrSQL = "Select 编码,简码,名称 From 药品生产商 " _
                            & "Where (站点 = [3] Or 站点 is Null) And (upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [2]) Order By 编码"
                    Set rs产地 = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[获取药品生产商]", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", strkey & "%", gstrNodeNo)
                    
                    If rs产地.EOF Then
                        If MsgBox("药品生产商没有找到你输入的产地，你要把它加入药品生产商中吗？", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol产地) = ""
                            mshBill.Text = ""
                            Exit Sub
                        Else
                            If LenB(strkey) > 60 Then
                                MsgBox "生产厂商名称过长(最多60个字符或30个汉字)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            
                            Dim rsMax As New Recordset
                            Dim int编码 As Integer, strCode As String, strSpecify As String
                            
                            If rsMax.State = 1 Then rsMax.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)),2) As Length FROM 药品生产商"
                            Set rsMax = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption)
                            int编码 = rsMax!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM 药品生产商"
                            rsMax.Close
                            Set rsMax = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption)
                            strCode = rsMax!Code
                            
                            int编码 = Len(strCode)
                            strCode = strCode + 1
                            If int编码 >= Len(strCode) Then
                                strCode = String(int编码 - Len(strCode), "0") & strCode
                            End If
                            strSpecify = zlCommFun.SpellCode(strkey)
                            
                            gstrSQL = "ZL_药品生产商_INSERT('" & strCode & "','" & strkey & "',zlSpellCode('" & strkey & "',10))"
                            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
                            
                            Exit Sub
                        End If
                    End If
                    If rs产地.RecordCount > 1 Then
                        Set msh产地.Recordset = rs产地
                        Dim intCol As Integer
                        Dim intRow As Integer
                        
                        With msh产地
                            If .Visible = False Then .Visible = True
                            .Redraw = False
                            .SetFocus
                            .Tag = "产地"
                            
                            For intRow = 0 To .rows - 1
                                .Row = intRow
                                For intCol = 0 To .Cols - 1
                                    .Col = intCol
                                    If .Row = 0 Then
                                        .CellFontBold = True
                                    Else
                                        .CellFontBold = False
                                    End If
                                Next
                            Next
                            .Font.Bold = False
                            .FontFixed.Bold = True
                            .ColWidth(0) = 1000
                            .ColWidth(1) = 2700
                            .ColWidth(2) = 1200
                            .Row = 1
                            .TopRow = 1
                            .Col = 0
                            .ColSel = .Cols - 1
                            
                            .Left = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                            .Top = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                            .Redraw = True
                            Cancel = True
                            Exit Sub
                        End With
                    Else
                        .Text = rs产地!名称
                    End If
                    gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.Text, mshBill.TextMatrix(mshBill.Row, 0))
                    If Not rsProvider.EOF Then
                        mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
                    End If
                End If
            Case mconIntCol批号
                '无处理
                If strkey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol批号) = ""
                        .Text = " "
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub
                End If
            Case mconIntCol生产日期
                '有处理
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            MsgBox "对不起，生产日期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strkey
                        .TextMatrix(.Row, mconIntCol生产日期) = .Text
                        
                        '设置效期
                        If Trim(.TextMatrix(.Row, mconIntCol原销期)) = "" Then
                            Exit Sub
                        End If
                        If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0) = "0" Then
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol生产日期) <> "" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol生产日期))
                        ElseIf .TextMatrix(.Row, mconIntCol批号) <> "" And Len(.TextMatrix(.Row, mconIntCol批号)) = 8 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol批号))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                If IsNumeric(strxq) Then
                                    If Trim(.TextMatrix(.Row, mconIntCol效期)) = "" Then
                                        strxq = TranNumToDate(strxq)
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    strxq = ""
                                End If
                            Else
                                strxq = ""
                            End If
                        End If
                        If Trim(strxq) = "" Then Exit Sub
                        
                        .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol原销期), "||")(0), strxq), "yyyy-mm-dd")
                        
                        If gtype_UserSysParms.P149_效期显示方式 = 1 Then
                            '换算为有效期
                            .TextMatrix(.Row, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntCol效期)), "yyyy-mm-dd")
                        End If
                        
'                        Call CheckLapse(.TextMatrix(.Row, mconIntCol效期))
                        Exit Sub
                    End If
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，生产日期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strkey = "" And strkey <> .TextMatrix(.Row, mconIntCol生产日期) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    
'                    Cancel = True
                    Exit Sub
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
'                        Call CheckLapse(strkey)
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
'                Call CheckLapse(strkey)
            Case mconintCol成本价
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，成本价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    strkey = GetFormat(strkey, mintCostDigit)
                    .Text = strkey
                End If
                
                '对时价药品的处理
                If strkey <> "" And strkey <> .TextMatrix(.Row, mconintCol成本价) And .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                    If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                        'Modified by zyb 2002-10-30
                        .Text = GetFormat(strkey, mintCostDigit)
                        If mbln加价率 Then
                            If mint时价入库时取上次售价 <> 1 Then
                                sngLeft = Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                sngTop = Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                If sngTop + 1700 > Screen.Height Then
                                    sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                End If
                                
                                With PicInput
                                    .Top = sngTop
                                    .Left = sngLeft
                                    .Visible = True
                                End With
                                
                                Txt加价率.Text = Val(Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1)) * 100 '默认规格的加成率
                                .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(strkey), Val(Txt加价率) / 100, Val(strkey) * (1 + (Val(Txt加价率) / 100))), mintPriceDigit)
                                If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And Val(strkey) <> 0 Then
                                    Txt加价率 = GetFormat(计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), Val(strkey)), 5)
                                End If
                                Txt加价率.Tag = Txt加价率
                                Txt加价率.SetFocus
                            End If
                        Else
                            If mint时价分段加成方式 = 1 Then
                                If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), strkey, dbl加成率, dblTemp售价) = False Then
                                    .TxtSetFocus
                                    Cancel = True
                                    Exit Sub
                                End If
                            Else
                                dbl加成率 = Val(Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1))
                                dblTemp售价 = strkey * (1 + dbl加成率)
                            End If
                                                        
                            If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), strkey, dbl加成率, dblTemp售价), mintPriceDigit)
                            If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                                .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit)
                            End If
                        End If
                    End If
                End If
                
                '设置金额
                If strkey <> "" And strkey <> .TextMatrix(.Row, mconintCol成本价) And .TextMatrix(.Row, mconIntCol数量) <> "" Then
                    .TextMatrix(.Row, mconintCol成本金额) = GetFormat(.TextMatrix(.Row, mconIntCol数量) * strkey, mintMoneyDigit)
                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol数量) * Val(.TextMatrix(.Row, mconIntCol售价)), mintMoneyDigit)
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                End If
                显示合计金额
                If .TextMatrix(.Row, mconIntCol售价) <> "" And .TextMatrix(.Row, mconIntCol比例系数) <> "" Then
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
                End If
            Case mconintCol成本金额
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，成本金额必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) * Val(.TextMatrix(.Row, mconIntCol数量)) < 0 Then
                        MsgBox "成本金额符号应与数量符号一致！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                '格式化金额
                If strkey <> "" Then
                    strkey = GetFormat(strkey, mintMoneyDigit)
                    .Text = strkey
                End If
                
                If strkey <> "" And strkey <> .TextMatrix(.Row, mconintCol成本金额) Then
                    If .TextMatrix(.Row, mconIntCol数量) <> "" Then
                        If mbln加价率 Then
                            '取得改变成本金额前的加价率
                            mdbl加价率 = 15
                            If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And Val(.TextMatrix(.Row, mconintCol成本价)) <> 0 Then
                                mdbl加价率 = 计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), Val(.TextMatrix(.Row, mconintCol成本价)))
                            End If
                        End If
                        
                        .Text = GetFormat(strkey, mintMoneyDigit)
                        .TextMatrix(.Row, mconintCol成本价) = GetFormat(strkey / .TextMatrix(.Row, mconIntCol数量), mintCostDigit)
                        '对时价药品的处理
                        If .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                            If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                                '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
                                If mbln加价率 Then
                                    If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), (mdbl加价率 / 100), Val(.TextMatrix(.Row, mconintCol成本价)) * (1 + (mdbl加价率 / 100))), mintPriceDigit)
                                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
                                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                                Else
                                    If mint时价分段加成方式 = 1 Then
                                        If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, dblTemp售价) = False Then
                                            .TxtSetFocus
                                            Cancel = True
                                            Exit Sub
                                        End If
                                    Else
                                        dbl加成率 = Val(Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1))
                                        dblTemp售价 = .TextMatrix(.Row, mconintCol成本价) * (1 + dbl加成率)
                                    End If
                                    
                                    If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, dblTemp售价), mintPriceDigit)
                                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol数量) * .TextMatrix(.Row, mconIntCol售价), mintMoneyDigit)
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mconIntCol数量)) <> 0 Then
                        .TextMatrix(.Row, mconintCol成本价) = GetFormat(strkey / Val(.TextMatrix(.Row, mconIntCol数量)), mintCostDigit)
                    End If
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - strkey, mintMoneyDigit)
                    .TextMatrix(.Row, mconintCol成本金额) = GetFormat(strkey, mintMoneyDigit)
                    
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
                End If
                显示合计金额
            Case mconIntCol数量
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
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Abs(Val(strkey)) = 0 Then
                        MsgBox "对不起，数量的绝对值必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mint编辑状态 = 2 And Val(.TextMatrix(.Row, mconIntCol数量)) <> 0 And .TextMatrix(.Row, mconIntCol是否新行) = "否" Then
                        If Not 相同符号(Val(strkey), Val(.TextMatrix(.Row, mconIntCol数量))) Then
                            MsgBox "对不起，数量的符号应该与原单据数量的符号一致！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strkey) < 0 Then
                        If Not IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol分批属性) = 1 Then
                            MsgBox "分批药品不允许负数入库，请重输", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                                        
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    
                    '检查负数退库时库存是否足够
                    If Val(strkey) < 0 Then
                        If Not CheckStock(Val(.TextMatrix(.Row, 0)), 0, Val(.Text) * Val(.TextMatrix(.Row, mconIntCol比例系数))) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Trim(.TextMatrix(.Row, mconintCol成本价)) <> "" Then
                        .TextMatrix(.Row, mconintCol成本金额) = GetFormat(.TextMatrix(.Row, mconintCol成本价) * strkey, mintMoneyDigit)
                        
                        '时价药品的处理
                        If .TextMatrix(.Row, mconIntCol原销期) <> "" Then
                            If Split(.TextMatrix(.Row, mconIntCol原销期), "||")(2) = 1 Then
                                'Modified by ZYB 2002-10-30
                                If mbln加价率 Then
                                    mdbl加价率 = 15
                                    If Val(.TextMatrix(.Row, mconIntCol售价)) <> 0 And Val(.TextMatrix(.Row, mconintCol成本价)) <> 0 Then
                                        mdbl加价率 = 计算加成率(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol售价)), Val(.TextMatrix(.Row, mconintCol成本价)))
                                    End If
                                    If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), (mdbl加价率 / 100), Val(.TextMatrix(.Row, mconintCol成本价)) * (1 + (mdbl加价率 / 100))), mintPriceDigit)
                                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * strkey, mintMoneyDigit)
                                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                                Else
                                    If mint时价分段加成方式 = 1 Then
                                        If get分段加成售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol比例系数)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, dblTemp售价) = False Then
                                            .TxtSetFocus
                                            Cancel = True
                                            Exit Sub
                                        End If
                                    Else
                                        dbl加成率 = Val(Split(.TextMatrix(.Row, mconIntCol原销期), "||")(1))
                                        dblTemp售价 = .TextMatrix(.Row, mconintCol成本价) * (1 + dbl加成率)
                                    End If
                                    If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, dblTemp售价), mintPriceDigit)
                                End If
                            End If
                        End If
                    End If
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol售价) * strkey, mintMoneyDigit)
                    End If
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                    
                    .TextMatrix(.Row, mconIntCol数量) = strkey
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
                End If
                显示合计金额
            Case mconIntCol冲销数量
                If .TextMatrix(.Row, .Col) = "" And strkey = "" Then
                    MsgBox "对不起，数量必须输入！", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Not 相同符号(Val(strkey), Val(.TextMatrix(.Row, mconIntCol数量))) Then
                        MsgBox "对不起，冲销数量的符号应该与原有数量一致！", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 0 Then
                        If Val(strkey) > Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strkey) < Val(.TextMatrix(.Row, mconIntCol数量)) Then
                            MsgBox "对不起，冲销数量不能大于原有数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "冲销数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    
                    If .TextMatrix(.Row, mconintCol成本价) <> "" Then
                        .TextMatrix(.Row, mconintCol成本金额) = GetFormat(.TextMatrix(.Row, mconintCol成本价) * strkey, mintMoneyDigit)
                    End If
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol售价) * strkey, mintMoneyDigit)
                    End If
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                    
                    .TextMatrix(.Row, mconIntCol冲销数量) = strkey
                    Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconintCol零售价)))
                End If
                显示合计金额
            Case mconIntCol批准文号
                If strkey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                        .TextMatrix(.Row, mconIntCol批准文号) = ""
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub
                End If
            Case mconIntCol售价
                '输入的售价不能大于指导零售价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "售价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strkey = GetFormat(.TextMatrix(.Row, mconIntCol售价), mintPriceDigit)
                
                '判断输入的零售价与指导零售价
                gstrSQL = "Select 指导零售价 From 药品目录 Where 药品ID=[1] "
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                
                dbl指导零售价 = Round(rsTemp!指导零售价 * Val(.TextMatrix(.Row, mconIntCol比例系数)), 5)
                strkey = Round(Val(strkey), 5)
                If Val(strkey) > dbl指导零售价 Then
                    MsgBox "输入的零售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                .Text = GetFormat(strkey, mintPriceDigit)
                .TextMatrix(.Row, .Col) = .Text
                
                '重算差价
                .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
                .TextMatrix(.Row, mconintCol差价) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                
                Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
            Case mconintCol零售价
                '输入的零售价不能大于指导零售价
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "零售价必须为数字型，请重输！", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strkey = GetFormat(.TextMatrix(.Row, mconintCol零售价), gtype_UserDrugDigits.Digit_零售价)
                
                '判断输入的零售价与指导零售价
                gstrSQL = "Select 指导零售价 From 药品目录 Where 药品ID=[1] "
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(.TextMatrix(.Row, 0)))
                
                dbl指导零售价 = Round(rsTemp!指导零售价, gtype_UserDrugDigits.Digit_零售价)
                strkey = Round(strkey, gtype_UserDrugDigits.Digit_零售价)
                If Val(strkey) > dbl指导零售价 Then
                    MsgBox "输入的零售价不能大于指导零售价（指导零售价：￥" & dbl指导零售价 & "）", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                .Text = GetFormat(strkey, gtype_UserDrugDigits.Digit_零售价)
                .TextMatrix(.Row, .Col) = .Text
                
                .TextMatrix(.Row, mconIntCol售价) = GetFormat(Val(.TextMatrix(.Row, .Col)) * Val(.TextMatrix(.Row, mconIntCol比例系数)), mintPriceDigit)
                If Val(.TextMatrix(.Row, mconintCol成本价)) = Val(.TextMatrix(.Row, mconIntCol售价)) Then
                '通过技术手段单独处理零差价销售情况下零售价和售价不等的情况
                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
                Else
                    .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, .Col)) * Val(.TextMatrix(.Row, mconIntCol比例系数)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
                End If
                .TextMatrix(.Row, mconintCol差价) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
                
                Call Set时价分批药品零售价(.Row, Val(.Text))
                Call 提示库存数
            Case mconIntCol外观
                '无处理
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol外观) = ""
                        .Text = " "
                    End If
                    
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub

                Else
                    Dim rs外观 As New Recordset
                    
                    gstrSQL = "Select 编码,简码,名称 From 药品外观 " _
                            & "Where upper(名称) like [1] or Upper(编码) like [2] or Upper(简码) like [3] "
                    Set rs外观 = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", strkey & "%")
                    
                    If rs外观.EOF Then
                        .TextMatrix(.Row, mconIntCol外观) = .Text
'                        .Col = mconIntCol数量
'                        Cancel = True
                        Exit Sub
                    Else
                        If rs外观.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol外观) = rs外观.Fields("名称")
                            .Text = rs外观.Fields("名称")
                        Else
                            Set msh产地.Recordset = rs外观
                            With msh产地
                                .Tag = "外观"
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
                End If
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
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, ByVal str药品编码 As String, _
    ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, ByVal str基本药物, _
    ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, _
    ByVal num指导批发价 As Double, ByVal str原产地 As String, _
    ByVal int原效期 As Integer, dbl比例系数 As Double, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal dbl指导差价率 As Double, ByVal str生产日期 As String, ByVal str售价单位 As String) As Boolean
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbl成本价 As Double, dbl加成率 As Double
    Dim rsPrice As New Recordset
    Dim lngDepartid As Long
    Dim str药名 As String
    Dim rsProvider As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim str规格产地 As String
    Dim rsTemp As ADODB.Recordset
    Dim rs售价 As ADODB.Recordset
    
    SetColValue = False
    lngDepartid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    On Error GoTo errHandle
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol行号 Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT a.加成率 from 药品规格 a where a.药品id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "加成率", lng药品ID)
        dbl加成率 = Nvl(rsTemp!加成率, 0) / 100
        
        gstrSQL = "SELECT Nvl(a.差价让利比,0) 差价让利比,nvl(a.扣率,0) 扣率,Nvl(a.招标药品,0) 招标药品,nvl(a.成本价,0) 成本价,a.上次批准文号, a.批准文号,a.上次产地 ,b.产地,a.上次生产日期" & _
                " from 药品规格 a,收费项目目录 b where a.药品id=b.id and 药品id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取扣率]", lng药品ID)
        
        dbl成本价 = rsTemp!成本价
        
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
        
        '产地、批准文号、生产日期规则，根据参数设置取
        '参数：优先从上次入库取
        '产地：直接从规格表中取上次产地，如果没有则从收费项目中取产地，没有则不填产地
        '批准文号：优先从规格表中取上次批准文号，如果没有则从规格表中取批准文号，还没有则不填批准文号
        '生产日期：优先从规格表中取上次生产日期，如果没有则不填
        '成本价：从规格表中取成本价
        
        '参数：优先从最近库存批次取
        '产地：优先从库存表最近批次中取产地，如果没有则从收费项目中取产地，没有则不填产地
        '批准文号：优先从库存表最近批次中取批准文号，如果没有则从规格表中取批准文号，还没有则不填批准文号
        '生产日期：优先从库存表最近批次中取生产日期，如果没有则不填
        '成本价：优先从药品库存表最近批次中取上次成本价，没有则从规格表中取成本价
        If IIf(IsNull(rsTemp!上次产地), "", rsTemp!上次产地) <> "" Then
            .TextMatrix(intRow, mconIntCol产地) = rsTemp!上次产地
        Else
            .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
        End If
        .TextMatrix(intRow, mconIntCol生产日期) = IIf(IsNull(rsTemp!上次生产日期), "", rsTemp!上次生产日期)
        .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsTemp!上次批准文号), IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号), rsTemp!上次批准文号)
        
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol售价) = GetFormat(num售价 * dbl比例系数, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(str原产地), "", str原产地)
        .TextMatrix(intRow, mconIntCol原销期) = IIf(IsNull(int原效期), "0", int原效期) & "||" & dbl加成率 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol比例系数) = dbl比例系数
        
        SetInputFormat intRow
        '分批属性
        Call Get药品分批属性(intRow)
        
        '说明：这里区分分批核算和不分批核算的目的是提高运行速度。
        '本来可以不分这些，直接用第一条SQL语句实现，但不分批的药品就多在数据库中扫描一次
        '0-优先从药品库存取;1-优先从药品规格取。
        If mint取上次成本价方式 = 0 Then
            If Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                gstrSQL = "select 上次采购价 as 上次成本价 ,上次产地,批准文号,上次生产日期 from 药品库存 where 性质=1 and 库房id=[1] and 药品id=[2] " & _
                        " and nvl(批次,0) =(select max(nvl(批次,0)) from 药品库存 where 性质=1 and 库房id=[1] and 药品id=[2] )"
            Else
                gstrSQL = "select 上次采购价 as 上次成本价,上次产地,批准文号,上次生产日期 from 药品库存 where 性质=1 and 库房id=[1] and 药品id=[2]"
            End If
            Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次成本价]", lngDepartid, lng药品ID)
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsPrice!上次产地), IIf(IsNull(rsTemp!产地), "", rsTemp!产地), rsPrice!上次产地)
                'mint时价入库售价加成方式
                If Nvl(rsPrice!上次成本价) = 0 Then
                    If dbl成本价 >= 0 Then
                        .TextMatrix(intRow, mconintCol成本价) = GetFormat(dbl成本价 * dbl比例系数, mintCostDigit)
                    End If
                Else
                    .TextMatrix(intRow, mconintCol成本价) = GetFormat(IIf(IsNull(rsPrice!上次成本价), 0, rsPrice!上次成本价) * dbl比例系数, mintCostDigit)
                End If
                .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsPrice!批准文号), IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号), rsPrice!批准文号)
                .TextMatrix(intRow, mconIntCol生产日期) = IIf(IsNull(rsPrice!上次生产日期), "", Format(rsPrice!上次生产日期, "yyyy-mm-dd"))
            Else
                .TextMatrix(intRow, mconIntCol生产日期) = ""
                If dbl成本价 >= 0 Then
                    .TextMatrix(intRow, mconintCol成本价) = GetFormat(dbl成本价 * dbl比例系数, mintCostDigit)
                End If
            End If
        Else
            If dbl成本价 >= 0 Then
                .TextMatrix(intRow, mconintCol成本价) = GetFormat(dbl成本价 * dbl比例系数, mintCostDigit)
            End If
        End If
        '时价药品处理
        If int是否变价 = 1 Then
            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
            '由于存在差价让利比的存在,需要按加成率计算,因此将指导差价率转换成加成率计算 公式：加成率=1/(1-差价率)-1
            If mint时价入库时取上次售价 = 1 Then
                gstrSQL = "select nvl(上次售价,0) 上次售价 from 药品规格 where 药品id=[1]"
                                 
                Set rs售价 = zldatabase.OpenSQLRecord(gstrSQL, "查询售价", lng药品ID)
                If rs售价!上次售价 > 0 Then
                    .TextMatrix(intRow, mconIntCol售价) = GetFormat(rs售价!上次售价 * dbl比例系数, mintPriceDigit)
                Else
                    .TextMatrix(intRow, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, .TextMatrix(intRow, mconintCol成本价) * (1 + dbl加成率)), mintPriceDigit)
                End If
            Else
                .TextMatrix(intRow, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), dbl加成率, .TextMatrix(intRow, mconintCol成本价) * (1 + dbl加成率)), mintPriceDigit)
            End If
            
        End If
        
        If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
            If mintUnit <> mconint售价单位 And Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                .TextMatrix(intRow, mconintCol零售单位) = str售价单位
            End If
        End If
        
        If .TextMatrix(intRow, mconIntCol产地) <> "" Then
            gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
            Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
               .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
            End If
        End If
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
    If mblnEdit = False Then Exit Sub
    With mshBill
    
        If mint编辑状态 = 1 Then
            .ColData(mconIntCol产地) = 1
        End If
        
        If .TextMatrix(intRow, mconIntCol原销期) <> "" Then
            .ColData(mconIntCol效期) = 2                '日期输入框
            '如果是时价药品，则允许输入售价
            If Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2) = 1 Then
                .ColData(mconIntCol售价) = IIf(Get时价药品直接确定售价, 4, 5)
            Else
                .ColData(mconIntCol售价) = 5
            End If
        Else
            .ColData(mconIntCol效期) = 5
        End If
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            If mshBill.TextMatrix(intRow, mconIntCol原销期) <> "" Then
                mshBill.ColData(mconintCol零售价) = 5
                If Val(Split(mshBill.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(mshBill.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                    mshBill.ColData(mconintCol零售价) = 4
                End If
            End If
        End If
    End With
End Sub


Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    OpenIme
    If mbln效期提示 Then
        If mshBill.Col = mconIntCol效期 Then
            If mbln日期提示 = False Then CheckLapse (mshBill.TextMatrix(mshBill.Row, mconIntCol效期))
        End If
    End If
End Sub

Private Sub mshBill_LostFocus()
    OpenIme
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

Private Sub msh产地_DblClick()
    msh产地_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh产地_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    On Error GoTo errHandle
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh产地.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            If msh产地.Tag = "产地" Then
                .TextMatrix(.Row, mconIntCol产地) = msh产地.TextMatrix(msh产地.Row, 2)
                .Text = msh产地.TextMatrix(msh产地.Row, 2)
                
                gstrSQL = "select 批准文号 from 药品生产商对照 where 厂家名称=[1] and 药品id=[2]"
                Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol产地), mshBill.TextMatrix(mshBill.Row, 0))
                If Not rsProvider.EOF Then
                    mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = IIf(IsNull(rsProvider!批准文号), "", rsProvider!批准文号)
                Else
                    mshBill.TextMatrix(mshBill.Row, mconIntCol批准文号) = ""
                End If
                .Col = mconIntCol批号
            Else
                .TextMatrix(.Row, mconIntCol外观) = msh产地.TextMatrix(msh产地.Row, 2)
                .Text = msh产地.TextMatrix(msh产地.Row, 2)
            End If
            
            msh产地.Visible = False
            .SetFocus
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh产地_LostFocus()
    If msh产地.Visible Then
        msh产地.Visible = False
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "SELECT count(*)" _
              & " From 部门性质说明 " _
             & " WHERE ((工作性质 LIKE '%药房') " _
                  & "   OR (工作性质 LIKE '制剂室')) " _
               & " AND 部门id =[1]"
    Set rsStock = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查]", cboStock.ItemData(cboStock.ListIndex))
               
               
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > txt摘要.MaxLength Then
                MsgBox "摘要超长,最多能输入" & CInt(txt摘要.MaxLength / 2) & "个汉字或" & txt摘要.MaxLength & "个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    '判断已失效药品是否禁止入库
                    If CheckLapse(Trim(.TextMatrix(intLop, mconIntCol效期)), True) = False Then
                        mbln日期提示 = True
                        MsgBox "第" & intLop & "行药品已经失效了，失效药品不能保存！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol效期
                        mbln日期提示 = False
                        Exit Function
                    End If
                
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconintCol成本价))) = "" Then
                        MsgBox "第" & intLop & "行药品的成本价为空了，请检查！", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol成本价
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconintCol成本金额))) = "" Then
                        MsgBox "第" & intLop & "行药品的成本金额为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol成本金额
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
                    
                    If Split(.TextMatrix(intLop, mconIntCol原销期), "||")(0) <> "0" And Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 Then
                        If blnStock = True And (Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Or Trim(.TextMatrix(intLop, mconIntCol效期)) = "") Then
                            MsgBox "第" & intLop & "行的药品是效期药品,请把它的批号及效期信息完整输入单据中！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If Trim(.TextMatrix(intLop, mconIntCol批号)) = "" Then
                                .Col = mconIntCol批号
                            Else
                                .Col = mconIntCol效期
                            End If
                            Exit Function
                        End If
                    End If
                    
                    '分批药品必须录入产地和批号
                    If Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 And (Trim(.TextMatrix(intLop, mconIntCol产地)) = "" Or Trim(.TextMatrix(intLop, mconIntCol批号)) = "") Then
                        MsgBox "第" & intLop & "行的药品是分批药品,请把它的产地和批号" & vbCrLf & "信息输入单据中！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If Trim(.TextMatrix(intLop, mconIntCol产地)) = "" Then
                            .Col = mconIntCol产地
                        Else
                            .Col = mconIntCol批号
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol成本价)) > 9999999999# Then
                        MsgBox "  第" & intLop & "行药品的成本价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol成本价
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol成本金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的成本金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol成本金额
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol数量
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCard(Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockID As Long
    Dim lngInOutTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
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
    Dim intRow As Integer
    Dim datTimeProduct As String
    Dim str批准文号 As String
    Dim n As Integer
    Dim m As Integer
    Dim dbl合计数量 As Double
    Dim lng药品ID As Long
    Dim str外观 As String
    Dim dbl金额差 As Double
    
    SaveCard = False
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = zldatabase.GetNextNo(24, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lngInOutTypeID = cboType.ItemData(cboType.ListIndex)
        strBrief = Trim(txt摘要.Text)
        strBooker = Trim(Txt填制人)
        datBookDate = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Trim(Txt审核人)
        
        '如果入库药品合计数量小于0，则要做库存检查(主要对不分批药品负数入库做校验)
        For n = 1 To .rows - 1
            If Val(.TextMatrix(n, 0)) <> 0 Then
                dbl合计数量 = 0
                lng药品ID = Val(.TextMatrix(n, 0))
                For m = 1 To .rows - 1
                    If Val(.TextMatrix(m, 0)) <> 0 And lng药品ID = Val(.TextMatrix(m, 0)) Then
                        dbl合计数量 = dbl合计数量 + Val(.TextMatrix(m, mconIntCol数量)) * Val(.TextMatrix(m, mconIntCol比例系数))
                    End If
                Next
                
                '合计数量为负数时才校验库存
                If dbl合计数量 < 0 Then
                    If Not CheckStock(lng药品ID, 0, dbl合计数量, n) Then
                        MsgBox "药品[" & .TextMatrix(n, mconIntCol药名) & "]库存不足，不能负数入库。"
                        Exit Function
                    End If
                End If
            End If
        Next
        
'        gcnOracle.BeginTrans
        If mint编辑状态 = 2 Or bln强制保存 Then        '修改
            gstrSQL = "zl_药品其他入库_Delete('" & mstr单据号 & "')"
            Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
        End If
            
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = Trim(.TextMatrix(intRow, mconIntCol产地))
                strBatchNo = Trim(.TextMatrix(intRow, mconIntCol批号))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol生产日期)) = "", "", Trim(.TextMatrix(intRow, mconIntCol生产日期)))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntCol效期)) = "", "", Trim(.TextMatrix(intRow, mconIntCol效期)))
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = .TextMatrix(intRow, mconIntCol数量) * .TextMatrix(intRow, mconIntCol比例系数)
                dblPurchasePrice = Round(.TextMatrix(intRow, mconintCol成本价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchaseMoney = .TextMatrix(intRow, mconintCol成本金额)
                dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol售价金额)
                dblMistakePrice = .TextMatrix(intRow, mconintCol差价)
                
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                lngSerial = intRow
                
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", Trim(.TextMatrix(intRow, mconIntCol批准文号)))
                str外观 = Trim(.TextMatrix(intRow, mconIntCol外观))
                
                '时价分批药品处理
                If Val(Split(.TextMatrix(intRow, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol分批属性)) = 1 Then
                    dblSalePrice = FormatEx(.TextMatrix(intRow, mconintCol零售价), gtype_UserDrugDigits.Digit_零售价)
                    dblSaleMoney = .TextMatrix(intRow, mconintCol零售金额)
                    dblMistakePrice = .TextMatrix(intRow, mconintCol零售差价)
                    dbl金额差 = GetFormat(Val(.TextMatrix(intRow, mconintCol零售金额)) - Val(.TextMatrix(intRow, mconIntCol售价金额)), mintMoneyDigit)
                End If
                
                gstrSQL = "zl_药品其他入库_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '序号
                gstrSQL = gstrSQL & "," & lngSerial
                '库房ID
                gstrSQL = gstrSQL & "," & lngStockID
                '入出类别ID
                gstrSQL = gstrSQL & "," & lngInOutTypeID
                '药品ID
                gstrSQL = gstrSQL & "," & lngDrugID
                '填写数量
                gstrSQL = gstrSQL & "," & dblQuantity
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
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                '生产日期
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '外观
                gstrSQL = gstrSQL & ",'" & str外观 & "'"
                '金额差
                gstrSQL = gstrSQL & "," & IIf(dbl金额差 <> 0, dbl金额差, "NULL")
                gstrSQL = gstrSQL & ")"

                Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
            recSort.MoveNext
        Next
        
'        gcnOracle.CommitTrans
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
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
    Dim n As Integer
    Dim str药品ID As String
    Dim str摘要 As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim str药品 As String
    
    arrSql = Array()
    SaveStrike = False
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol数量)), Val(.TextMatrix(intRow, mconIntCol冲销数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '检查可用数量是否足够，参数设置为不检查库存时不进行
            If mint库存检查 <> 0 And .TextMatrix(intRow, 0) <> "" Then
                If .TextMatrix(intRow, mconIntCol冲销数量) = .TextMatrix(intRow, mconIntCol数量) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mconintCol真实数量))
                Else
                    冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol冲销数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
                If CheckStrickUsable(4, Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), .TextMatrix(intRow, mconIntCol药名), _
                    0, Val(冲销数量_IN), mint库存检查, Trim(txtNo.Tag), Val(.TextMatrix(intRow, mconIntCol序号))) = False Then
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
            End If
        Next
        
        '检查库存 防止可以数量大于库存数量，冲销后库存数量为负
        str药品 = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol批次, mconIntCol冲销数量, mconIntCol比例系数, 2, , mintNumberDigit)
        If str药品 <> "" Then
            If mint库存检查 = 1 Then '不足提醒
                If MsgBox(str药品 & " 药品“实际库存”不足，是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf mint库存检查 = 2 Then '不足禁止
                MsgBox str药品 & " 药品“实际库存”不足，不能冲销！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        填制人_IN = UserInfo.用户姓名
        填制日期_IN = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        On Error GoTo errHandle
        行次_IN = 0
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol冲销数量)) <> 0 Then
                行次_IN = 行次_IN + 1
                
                药品ID_IN = .TextMatrix(intRow, 0)
                str药品ID = IIf(str药品ID = "", "", str药品ID & ",") & 药品ID_IN
                If .TextMatrix(intRow, mconIntCol冲销数量) = .TextMatrix(intRow, mconIntCol数量) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mconintCol真实数量))
                Else
                    冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol冲销数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
                str摘要 = txt摘要.Text
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                
                gstrSQL = "ZL_药品其他入库_STRIKE("
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
                '摘要
                gstrSQL = gstrSQL & ",'" & str摘要 & "'"
                '冲销数量
                gstrSQL = gstrSQL & "," & 冲销数量_IN
                '填制人
                gstrSQL = gstrSQL & ",'" & 填制人_IN & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & Format(填制日期_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If 行次_IN = 0 Then
            MsgBox "没有选择一行药品来冲销，请录入冲销数量！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '提示停用药品
        If str药品ID <> "" Then
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
    Dim dbl时价分批 As Boolean
    
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconintCol成本金额))
'            Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            If .TextMatrix(intLop, mconIntCol原销期) <> "" Then
                If Val(Split(.TextMatrix(intLop, mconIntCol原销期), "||")(2)) = 1 And Val(.TextMatrix(intLop, mconIntCol分批属性)) = 1 Then
                    dbl时价分批 = True
                    Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconintCol零售金额))
                Else
                    Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
                End If
            Else
                Cur记帐金额 = Cur记帐金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            End If
        Next
    End With
    
    Cur记帐差价 = Cur记帐金额 - curTotal
    
    lblPurchasePrice.Caption = "成本金额合计：" & GetFormat(curTotal, mintMoneyDigit)
    
    If dbl时价分批 = True Then
        lblSalePrice.Caption = "售价金额(时价分批按零售金额)合计：" & GetFormat(Cur记帐金额, mintMoneyDigit)
        lblDifference.Caption = "差价(时价分批按零售差价)合计：" & GetFormat(Cur记帐差价, mintMoneyDigit)
    Else
        lblDifference.Caption = "差价合计：" & GetFormat(Cur记帐差价, mintMoneyDigit)
        lblSalePrice.Caption = "售价金额合计：" & GetFormat(Cur记帐金额, mintMoneyDigit)
    End If
End Sub

Private Sub 提示库存数()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl数量 As Double
    Dim str单位 As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntCol药名) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)

    Select Case mintUnit
        Case mconint售价单位
            strUnit = "C.计算单位"
            strQuantity = "可用数量 "
        Case mconint门诊单位
            strUnit = "B.门诊单位"
            strQuantity = "可用数量/门诊包装 "
        Case mconint住院单位
            strUnit = "B.住院单位"
            strQuantity = "可用数量/住院包装 "
        Case mconint药库单位
            strUnit = "B.药库单位"
            strQuantity = "可用数量/药库包装 "
    End Select
    
    gstrSQL = "Select b.药品ID," & strUnit & " as 单位, Sum(" & strQuantity & ") as 数量 " & _
        " From 药品库存 a,药品规格 b,收费项目目录 C " & _
        " Where a.性质=1 and a.药品id=b.药品id And B.药品ID=C.ID " & _
        " And 可用数量<>0 And 库房ID=[1] and b.药品ID=[2] " & _
        " Group by b.药品ID," & strUnit
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboStock.ItemData(cboStock.ListIndex), intID)
    
    With RecTmp
        If .EOF Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        Dbl数量 = IIf(IsNull(!数量), 0, !数量)
        
        staThis.Panels(2).Text = "该药品当前库存数为[" & GetFormat(Dbl数量, mintNumberDigit) & "]" & !单位
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


'打印单据
Private Sub printbill()
    Dim strUnit As String
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1302", "zl8_bill_1302"), mint记录状态, int单位系数, 1302, "药品其它入库单", strNo
End Sub


'取指导批发价定价单位的设置值，缺省为0-按售价单位定价，可选为1-按药库单位定价；
Private Function GetUnit() As Integer
    GetUnit = gtype_UserSysParms.P29_指导批发价定价单位
End Function

'取数据库中批号的长度，这样，程序中的批号长度与数据库中保持一致了
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 批号 from 药品收发记录 where rownum<1 "
    Call zldatabase.OpenRecordset(rsBatchNolen, gstrSQL, "取字段长度")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PicInput_LostFocus()
    Dim strActive As String
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDYES,CMDNO,TXT加价率", strActive) <> 0 Then
        Exit Sub
    Else
        If strActive = "MSHBILL" Then
            If mshBill.Col = mconintCol成本价 Or mshBill.Col = mconintCol成本金额 Then Exit Sub
        End If
    End If
    PicInput.Visible = False
End Sub

Private Sub Txt加价率_GotFocus()
    Txt加价率.SelStart = 0
    Txt加价率.SelLength = Len(Txt加价率)
End Sub

Private Sub Txt加价率_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdYes_Click
End Sub

Private Sub Txt加价率_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub Txt加价率_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdYes_Click()
    If Val(Txt加价率) > 100 Or Val(Txt加价率) < 0 Then
        MsgBox "请输入合法的加成率！", vbInformation, gstrSysName
        Txt加价率.SetFocus
        Exit Sub
    End If
    
    With mshBill
        '重新计算零售价、差价
        If mint时价入库时取上次售价 <> 1 Then .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), Val(Txt加价率) / 100, Val(.TextMatrix(.Row, mconintCol成本价)) * (1 + (Val(Txt加价率) / 100))), mintPriceDigit)
        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
        .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
        
        Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub

Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdNo_Click()
    With mshBill
        mdbl加价率 = Val(Txt加价率.Tag)
        
        '重新计算零售价、差价
        If mint时价入库时取上次售价 <> 1 Then
            .TextMatrix(.Row, mconIntCol售价) = GetFormat(时价药品零售价(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconintCol成本价)), mdbl加价率 / 100, Val(.TextMatrix(.Row, mconintCol成本价)) * (1 + (mdbl加价率 / 100))), mintPriceDigit)
        End If
        
        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价)) * Val(.TextMatrix(.Row, mconIntCol数量)), mintMoneyDigit)
        .TextMatrix(.Row, mconintCol差价) = GetFormat(IIf(.TextMatrix(.Row, mconIntCol售价金额) = "", 0, .TextMatrix(.Row, mconIntCol售价金额)) - IIf(.TextMatrix(.Row, mconintCol成本金额) = "", 0, .TextMatrix(.Row, mconintCol成本金额)), mintMoneyDigit)
    
        Call Set时价分批药品零售价(.Row, Val(.TextMatrix(.Row, mconIntCol售价)) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

'取时价药品入库时，是否必须输入加价率
Private Function Get加价率() As Boolean
    Get加价率 = (Val(zldatabase.GetPara("时价药品以加价率入库", glngSys, 模块号.其他入库, 0)) = 1)
End Function

Private Function Get时价药品直接确定售价() As Boolean
    Get时价药品直接确定售价 = (gtype_UserSysParms.P76_时价药品直接确定售价 = 1)
End Function
Private Sub GetSysParm()
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
End Sub
Private Function 时价药品零售价(ByVal lng药品ID As Long, ByVal sin成本价 As Double, ByVal sin加成率 As Double, ByVal sin售价 As Double, Optional ByVal lngLastRow As Long = -1) As Double
    Dim sin零售价 As Double, sin指导零售价 As Double, sin差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    Dim sin差价让利 As Double
    '时价药品零售价计算公式:成本价*(1+加成率)
    '改为:成本价*(1+加成率)+(指导零售价-成本价*(1+加成率))*(1-差价让利比)
    '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-成本价*(1+加成率))*(1-差价让利比)
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    On Error GoTo errHandle
    gstrSQL = "Select 指导零售价,Nvl(差价让利比,100) 差价让利比 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比

    时价药品零售价 = 0
    If sin差价让利比 = 100 Then
        时价药品零售价 = sin售价
        Exit Function
    End If
    
    sin零售价 = sin成本价 * (1 + sin加成率)
    If sin零售价 / Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数)) >= sin指导零售价 Then
        时价药品零售价 = sin售价
        Exit Function
    End If
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
    sin差价让利 = (sin指导零售价 - sin零售价) * (1 - sin差价让利比 / 100)
    
    时价药品零售价 = IIf(sin售价 + sin差价让利 > sin指导零售价, sin指导零售价, sin售价 + sin差价让利)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 计算加成率(ByVal lng药品ID As Long, ByVal sin零售价 As Double, ByVal sin成本价 As Double) As Double
    Dim sin指导零售价 As Double, sin差价让利比 As Double
    Dim rsTemp As New ADODB.Recordset
    '根据零售价反算成本价,由于时价药品公式的变化,导致原来计算加成率的公式无效,需重新计算
    '原公式:(零售价/成本价-1)*100
    '现公式的理论:由于零售价是按加成率算出来后,再加上了让利外那部分金额,因此实际按加成率算出的零售价=指导零售价-(指导零售价-零售价)/差价让利比
    '再套用原公式算出实际的加成率
    计算加成率 = 0.15
    On Error GoTo errHandle
    gstrSQL = " Select 指导零售价,Nvl(差价让利比,100) 差价让利比,Nvl(是否变价,0) 时价 " & _
              " From 药品规格 A,收费项目目录 B Where A.药品ID=B.ID AND A.药品ID=[1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    
    sin指导零售价 = rsTemp!指导零售价
    sin差价让利比 = rsTemp!差价让利比
    If rsTemp!时价 = 0 Then Exit Function
    
    '指导零售价-(指导零售价-零售价)/差价让利比
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(mshBill.Row, mconIntCol比例系数))
    If sin差价让利比 <> 100 And sin差价让利比 > 0 Then
        sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价) / sin差价让利比 * 100
    Else
        sin零售价 = sin指导零售价 - (sin指导零售价 - sin零售价)
    End If
    计算加成率 = (Val(sin零售价) / Val(sin成本价) - 1) * 100
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 校正零售价(ByVal sin零售价 As Double, Optional ByVal lngLastRow As Long = -1) As Double
    '得到按当前单位系数计算出来的指导零售价，如果时价药品计算出来的零售价大于指导零售价，以指导零售价为准
    Dim sin指导零售价 As Double
    Dim rsTemp As New ADODB.Recordset
    
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    On Error GoTo errHandle
    gstrSQL = " Select 指导零售价,Nvl(差价让利比,100) 差价让利比 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", Val(mshBill.TextMatrix(lngLastRow, 0)))
    
    sin指导零售价 = rsTemp!指导零售价
    sin指导零售价 = sin指导零售价 * Val(mshBill.TextMatrix(lngLastRow, mconIntCol比例系数))
    
    校正零售价 = IIf(sin零售价 > sin指导零售价, sin指导零售价, sin零售价)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function get分段加成售价(ByVal lng药品ID As Long, ByVal lng比例系数 As Long, ByVal dbl成本价 As Double, ByRef dblR加成率 As Double, ByRef dbl售价 As Double) As Boolean
    '功能:在启用时价药品分段加成入库后，根据成本价计算出相应的售价
    '售价计算公式：购进价格在2000元/支、瓶或盒（含2000元）以下的药品，最高零售价格=实际购进价×（1+差价率）+差价额；
    '               购进价格在2000元/支、瓶或盒（不含2000元）以上的药品：最高零售价格 = 实际购进价 + 差价额（此段已经调整，不再适用）

    '参数：成本价
    Dim dbl加成率 As Double
    Dim dbl差价额 As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    dbl加成率 = 0
    dbl差价额 = 0
    
    gstrSQL = "select 类别 from  收费项目目录 a where a.id=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取得药品材质分类", lng药品ID)
    If rsTemp!类别 = 7 Then
        mrs分段加成.Filter = "类型=1"
    Else
        mrs分段加成.Filter = "类型=0"
    End If
      
    If mrs分段加成.RecordCount <> 0 Then
        mrs分段加成.MoveFirst
        Do While Not mrs分段加成.EOF
            With mrs分段加成
                If dbl成本价 > !最低价 And dbl成本价 <= !最高价 Then
                    dbl加成率 = IIf(IsNull(!加成率), 0, !加成率) / 100
                    dblR加成率 = dbl加成率
                    dbl差价额 = IIf(IsNull(!差价额), 0, !差价额)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs分段加成.MoveNext
        Loop
    End If
    
    If blnData = False Then
        If rsTemp!类别 = 7 Then
            MsgBox "【草药】未设置金额段为：" & dbl成本价 & " " & "的分段加成数据，请到药品目录管理中分段加成率设置！", vbInformation, gstrSysName
        Else
            MsgBox "【西药/成药】未设置金额段为：" & dbl成本价 & " " & "的分段加成数据，请到药品目录管理中分段加成率设置！", vbInformation, gstrSysName
        End If
        get分段加成售价 = False
    End If
    
    dbl售价 = dbl成本价 * (1 + dbl加成率) + dbl差价额
    
    Set rsTemp = Nothing
    gstrSQL = "Select 指导零售价 From 药品规格 Where 药品ID=[1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取指导零售价]", lng药品ID)
    If rsTemp!指导零售价 * lng比例系数 < dbl售价 Then
        dbl售价 = rsTemp!指导零售价 * lng比例系数
    End If
    
    get分段加成售价 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function 检查售价() As Boolean
    '功能：外购新增时，判断定价药品是否是最新售价，是则修改后提示
    Dim strMsg As String '保存提示信息
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl数量 As Double, dbl成本价 As Double, dbl成本金额 As Double, dbl零售价 As Double, dbl零售金额 As Double, dbl差价 As Double
    
    On Error GoTo errHandle
    
    检查售价 = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
                
                If Val(Split(.TextMatrix(i, mconIntCol原销期), "||")(2)) = 0 Then '判断定价

                    dbl零售价 = zlStr.FormatEx(Get售价(False, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), 0) * Val(.TextMatrix(i, mconIntCol比例系数)), mintPriceDigit)
                    
                    If .TextMatrix(i, mconIntCol售价) <> dbl零售价 Then
                        intSum = intSum + 1 '记录更新了几条数据
                        
                        dbl成本价 = Val(.TextMatrix(i, mconintCol成本价))
                        Dbl数量 = Val(.TextMatrix(i, mconIntCol数量))
                        dbl成本金额 = dbl成本价 * Dbl数量
                        dbl零售金额 = dbl零售价 * Dbl数量
                        dbl差价 = dbl零售金额 - dbl成本金额
                        
                        '更新售价相关数据
                        .TextMatrix(i, mconIntCol售价) = zlStr.FormatEx(dbl零售价, mintPriceDigit, , True)
                        .TextMatrix(i, mconIntCol售价金额) = zlStr.FormatEx(dbl零售金额, mintMoneyDigit, , True)
                        .TextMatrix(i, mconintCol差价) = zlStr.FormatEx(dbl差价, mintMoneyDigit, , True)
                        
                    End If
                End If
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "有记录未使用最新售价，程序已自动完成更新（售价、售价金额、差价），更新后请检查！", vbInformation, gstrSysName
            检查售价 = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '返回下一个可见并可用的列号
    Dim n As Integer
    Dim intNextCol As Integer

    If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
        GetNextEnableCol = mintLastCol
        Exit Function
    End If
    
    With mshBill
        For n = intCurrCol + 1 To .Cols - 1
            If .ColWidth(n) > 0 And .ColData(n) <> 5 Then
                intNextCol = n
                Exit For
            End If
        Next
    End With
    
    GetNextEnableCol = IIf(intNextCol = 0, mintLastCol, intNextCol)
End Function


