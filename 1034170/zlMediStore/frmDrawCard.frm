VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrawCard 
   Caption         =   "药品领用单"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDrawCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExpend 
      Caption         =   "自动分解(&A)"
      Height          =   350
      Left            =   4800
      TabIndex        =   35
      Top             =   5480
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   33
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   32
      Top             =   5460
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   14
      Top             =   0
      Width           =   11715
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
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
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "导入记帐单:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cboDrawPerson 
         Height          =   300
         Left            =   9645
         TabIndex        =   4
         Top             =   615
         Width           =   1515
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1815
         Left            =   7320
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3201
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
      Begin VB.TextBox txtDraw 
         Height          =   300
         Left            =   5355
         TabIndex        =   3
         Top             =   615
         Width           =   2415
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "…"
         Height          =   300
         Left            =   7755
         TabIndex        =   5
         Top             =   615
         Width           =   300
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
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
         TabIndex        =   8
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
      Begin VB.Label lblDrawPerson 
         AutoSize        =   -1  'True
         Caption         =   "领用人(&P)"
         Height          =   180
         Left            =   8730
         TabIndex        =   34
         Top             =   675
         Width           =   810
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   29
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   25
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   24
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   23
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   7
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品领用单"
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
         TabIndex        =   19
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发药库房(&S)"
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领药部门(&D)"
         Height          =   180
         Left            =   4275
         TabIndex        =   2
         Top             =   675
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
            Picture         =   "frmDrawCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1000
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
            Picture         =   "frmDrawCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
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
            Picture         =   "frmDrawCard.frx":22EA
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
            Picture         =   "frmDrawCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrawCard.frx":3080
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
      TabIndex        =   26
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
Attribute VB_Name = "frmDrawCard"
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
Private mblnStock As Boolean                '当前操作员是否是库房人员
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEnterCell As Boolean            '是否允许激法ENTERCELL()事件
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mblnAutoExp As Boolean                           '单据发生了自动分解
Private mbln提示 As Boolean                 '在药品选择器中选择的药品与界面中已有数据的比较看是否重复，对于重复的数据只提示一次，true 已经提示了，false还没有提示

Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                 '权限
Private mbln下可用数量 As Boolean           '填单是否下可用数量
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价

Private Const mlng紫色 As Long = &HC000C0

Private mint领用方式 As Integer              '0-向库房领药;1-向科室留存领药
Private str留存不足提示 As String
Private mint留存方式 As Integer             '0-按年留存 1-按月留存
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mblnUpdate As Boolean               '表示是否已根据最新价格更新单据内容

Private mlng出库库房 As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称
Private Const MStrCaption As String = "药品领用管理"

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

'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol来源 As Integer = 4
Private Const mconIntCol基本药物 As Integer = 5
Private Const mconIntCol序号 As Integer = 6
Private Const mconIntCol规格 As Integer = 7
Private Const mconIntCol可用数量 As Integer = 8
Private Const mconIntCol指导差价率 As Integer = 9
Private Const mconIntCol实际金额 As Integer = 10
Private Const mconIntCol实际差价 As Integer = 11
Private Const mconIntCol比例系数 As Integer = 12
Private Const mconIntCol批次 As Integer = 13
Private Const mconIntCol产地 As Integer = 14
Private Const mconIntCol单位 As Integer = 15
Private Const mconIntCol批号 As Integer = 16
Private Const mconIntCol效期 As Integer = 17
Private Const mconIntCol批准文号 As Integer = 18
Private Const mconIntCol填写数量 As Integer = 19
Private Const mconIntCol实际数量 As Integer = 20
Private Const mconIntCol采购价 As Integer = 21
Private Const mconIntCol采购金额 As Integer = 22
Private Const mconIntCol售价 As Integer = 23
Private Const mconIntCol售价金额 As Integer = 24
Private Const mconintCol差价 As Integer = 25
Private Const mconintCol真实数量 As Integer = 26
Private Const mconIntCol药品编码和名称 As Integer = 27
Private Const mconIntCol药品编码 As Integer = 28
Private Const mconIntCol药品名称 As Integer = 29
Private Const mconintCol原始数量 As Integer = 30
Private Const mconIntColS  As Integer = 31             '总列数
'=========================================================================================

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
Private Function Check留存() As Boolean
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim n As Integer
    Dim strSQL As String
    
    '检查科室留存可用数量是否足够
    On Error GoTo errHandle
    
    If mshBill.TextMatrix(1, 0) = "" Then Check留存 = True: Exit Function
    
    With rs
        .Fields.Append "药品ID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
        .Fields.Append "界面实际数量", adDouble, 18, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 40, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) > 0 Then
                If .RecordCount = 0 Then
                    .AddNew
                    !药品id = Val(mshBill.TextMatrix(n, 0))
                    !实际数量 = Val(mshBill.TextMatrix(n, mconIntCol实际数量)) * Val(mshBill.TextMatrix(n, mconIntCol比例系数))
                    !界面实际数量 = Val(mshBill.TextMatrix(n, mconIntCol实际数量))
                    !药品名称 = mshBill.TextMatrix(n, 2)
                    !单位 = Nvl(mshBill.TextMatrix(n, mconIntCol单位))
                    .Update
                Else
                    .MoveFirst
                    .Find "药品ID=" & Val(mshBill.TextMatrix(n, 0)) & " "
                    If .EOF Then
                        .AddNew
                        !药品id = Val(mshBill.TextMatrix(n, 0))
                        !实际数量 = Val(mshBill.TextMatrix(n, mconIntCol实际数量)) * Val(mshBill.TextMatrix(n, mconIntCol比例系数))
                        !界面实际数量 = Val(mshBill.TextMatrix(n, mconIntCol实际数量))
                        !药品名称 = mshBill.TextMatrix(n, 2)
                        !单位 = Nvl(mshBill.TextMatrix(n, mconIntCol单位))
                        .Update
                    Else
                        !实际数量 = !实际数量 + Val(mshBill.TextMatrix(n, mconIntCol实际数量)) * Val(mshBill.TextMatrix(n, mconIntCol比例系数))
                        !界面实际数量 = !界面实际数量 + Val(mshBill.TextMatrix(n, mconIntCol实际数量))
                        !单位 = Nvl(mshBill.TextMatrix(n, mconIntCol单位))
                        .Update
                    End If
                End If
            End If
        Next
    End With
    
    rs.MoveFirst
    For n = 1 To rs.RecordCount
        strSQL = "select a.实际数量,a.可用数量,b.计算单位 as 单位 from 药品留存 a,收费项目目录 b where a.药品ID=b.id and a.科室id=[2] and a.库房id=[1] " & _
        " and a.药品id=[3] and a.期间 = [4]"
        Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), txtDraw.Tag, CLng(rs!药品id), Format(zlDataBase.Currentdate(), IIf(mint留存方式 = 0, "yyyy", "yyyymm")))
        
        If gtype_UserSysParms.P175_药品领用明确批次 <> 1 Then '不明确批次检查可用数量
            If rsTmp.RecordCount = 0 Then
                Check留存 = False
                str留存不足提示 = "该期间【" & rs!药品名称 & "】没有留存数量不能领用，请修改单据！"
                Exit Function
            ElseIf rsTmp!可用数量 < rs!实际数量 Then
                Check留存 = False
                If rs!单位 <> rsTmp!单位 Then
                    str留存不足提示 = rs!药品名称 & "领用数量[" & rs!界面实际数量 & rs!单位 & "=" & rs!实际数量 & rsTmp!单位 & "]大于了留存可用数量[" & rsTmp!可用数量 & rsTmp!单位 & "]不能领用，请修改填写数量！"
                Else
                   str留存不足提示 = rs!药品名称 & "领用数量[" & rs!实际数量 & rs!单位 & "]大于了留存可用数量[" & rsTmp!可用数量 & rsTmp!单位 & "]不能领用，请修改填写数量！"
                End If
                mshBill.SetFocus
                mshBill.Row = n
                mshBill.MsfObj.TopRow = n
                mshBill.Col = mconIntCol填写数量
                
                Exit Function
            End If
        End If
        
        If rsTmp.RecordCount = 0 Then
            Check留存 = False
            str留存不足提示 = "该期间【" & rs!药品名称 & "】没有留存数量不能领用，请修改单据！"
            Exit Function
        ElseIf rsTmp!实际数量 < rs!实际数量 Then
            Check留存 = False
            If rs!单位 <> rsTmp!单位 Then
                str留存不足提示 = rs!药品名称 & "领用数量[" & rs!界面实际数量 & rs!单位 & "=" & rs!实际数量 & rsTmp!单位 & "]大于了留存数量[" & rsTmp!实际数量 & rsTmp!单位 & "]不能领用，请修改填写数量！"
            Else
               str留存不足提示 = rs!药品名称 & "领用数量[" & rs!实际数量 & rs!单位 & "]大于了留存数量[" & rsTmp!实际数量 & rsTmp!单位 & "]不能领用，请修改填写数量！"
            End If
            mshBill.SetFocus
            mshBill.Row = n
            mshBill.MsfObj.TopRow = n
            mshBill.Col = mconIntCol填写数量
            
            Exit Function
        End If
        rs.MoveNext
    Next
    
    Check留存 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetSysParm()
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
End Sub
Private Sub GetDrawPerson(ByVal strDeptId As String)
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    cboDrawPerson.Clear
    
    If strDeptId = "" Then Exit Sub

    gstrSQL = "Select 编号,姓名,简码 From 人员表 Where (站点 = [2] Or 站点 is Null) And Id In (Select 人员id From 部门人员 Where 部门id=[1]) " & _
              " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, strDeptId, gstrNodeNo)
    
    If rs.RecordCount = 0 Then Exit Sub
    
    For n = 1 To rs.RecordCount
        cboDrawPerson.AddItem (rs!姓名)
        rs.MoveNext
    Next
    rs.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 7 "
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "没有设置药品领用的出库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close

    If mblnStock Then
        strSQL = "SELECT DISTINCT a.id, a.名称 " _
               & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
               & "Where (a.站点 = [2] Or a.站点 is Null) And c.工作性质 = b.名称 " _
               & "  AND b.编码 ='O' AND a.id = c.部门id " _
               & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"
    Else
        strSQL = " Select C.ID " & _
                 " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                 " Where (c.站点 = [2] Or c.站点 is Null) And A.工作性质=B.名称 And A.部门ID=C.ID " & _
                 "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                 "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
    End If
    Set rsDepend = zlDataBase.OpenSQLRecord(strSQL, "药品领用单", UserInfo.用户ID, gstrNodeNo)
    
    If rsDepend.EOF Then
        If mblnStock Then
            MsgBox "领药部门性质信息不全,请查看部门管理！", vbInformation, gstrSysName
        Else
            MsgBox "你不属于任何领用部门,不能填写药品领用单,请查看部门管理！", vbInformation, gstrSysName
        End If
        rsDepend.Close
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, ByVal bln药库人员 As Boolean, _
    Optional int记录状态 As Integer = 1, Optional int领用方式 As Integer = 0, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mblnStock = bln药库人员
    mintParallelRecord = 1
    mint领用方式 = int领用方式
    mstrPrivs = GetPrivFunc(glngSys, 1305)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
        cmdExpend.Visible = True
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 6 Then
        CmdSave.Caption = "冲销(&O)"
        cmdAllCls.Visible = True
        cmdAllSel.Visible = True
    End If
    LblTitle.Caption = GetUnitName & "药品领用单" & IIf(mint领用方式 = 0, "(库房领用)", "(留存领用)")
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
End Sub

Private Sub cboDrawPerson_Click()
    'mshBill.SetFocus
    mshBill.Col = 1
    mshBill.Row = 1
End Sub

Private Sub cboDrawPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    With cboDrawPerson
        If Trim(.Text) = "" Then Exit Sub
        strText = UCase(.Text)
        
        mshProvider.Tag = 1
        
        gstrSQL = "Select 编号,姓名,简码 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And Id In(Select 人员id From 部门人员 Where 部门id=[1]) " & _
                  "  And (简码 like [2] Or 编号 like [2] or 姓名 like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
            Val(Me.txtDraw.Tag), _
            IIf(gstrMatchMethod = "0", "%", "") & strText & "%", _
            gstrNodeNo)
        
        If rs.EOF Then
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
        
        If rs.RecordCount > 1 Then
            Set mshProvider.Recordset = rs
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
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
                .ColWidth(0) = 800
                .ColWidth(1) = 800
                .ColWidth(2) = 800
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Width = lblDrawPerson.Width + cboDrawPerson.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = lblDrawPerson.Left
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = rs!姓名
            mshBill.SetFocus
            mshBill.Col = 1
            mshBill.Row = 1
        End If
        rs.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDrawPerson_KeyPress(KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
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
                    
                    mlng出库库房 = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                    
                    If Not mblnStock Then
                    MsgBox "请重新设置领药部门和领药人！", vbOKOnly, gstrSysName
                        txtDraw.Text = ""
                        txtDraw.Tag = ""
                        cboDrawPerson.Clear
                    End If
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
        '2010-5-7 标记已修改
        mblnChange = True
    End With
    Call 显示合计金额
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDraw_Click()
    Dim rsProvider As New Recordset
    Dim str站点限制 As String
    
    On Error GoTo errHandle
    str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
    If mblnStock Then
        gstrSQL = "SELECT DISTINCT a.id,null 上级id,1 as 末级, a.编码,a.简码,a.名称 " _
                & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                & "Where " & IIf(str站点限制 <> "", " (a.站点 = [3] or a.站点 is null) AND ", "") & " c.工作性质 = b.名称 " _
                & "  AND b.编码 ='O' AND a.id = c.部门id " _
                & "  AND (TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' Or a.撤档时间 Is NULL) " _
                & "Order By a.编码 "
    Else
        gstrSQL = " Select C.ID " & _
                  " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                  " Where " & IIf(str站点限制 <> "", " (C.站点 = [3] or C.站点 is null) And ", "") & " A.工作性质=B.名称 And A.部门ID=C.ID " & _
                  "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                  "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])"
        '只提取设置了领用流向的科室
        gstrSQL = " SELECT DISTINCT C.id,null 上级id,1 as 末级, C.编码,C.简码,C.名称" & _
                  " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                  " Where " & IIf(str站点限制 <> "", " (C.站点 = [3] or C.站点 is null) And ", "") & " A.工作性质=B.名称 And A.部门ID=C.ID " & _
                  "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                  "   And C.ID IN (Select Distinct 领用部门ID From 药品领用控制 Where 对方库房id=[2] And 领用部门ID IN (" & gstrSQL & ")) " & _
                  " Order By C.编码 "
    End If
    Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取所有领药部门]", _
        UserInfo.用户ID, _
        cboStock.ItemData(cboStock.ListIndex), _
        str站点限制)
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "所有领药部门"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = False Then
            Unload FrmSelect
            Exit Sub
        End If
        
        Me.txtDraw.Tag = .CurrentID
        Me.txtDraw = .CurrentName
    End With
    Unload FrmSelect
    
    Call GetDrawPerson(Me.txtDraw.Tag)
    cboDrawPerson.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExpend_Click()
    mblnAutoExp = AutoExpend
End Sub

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    '功能：自动分解
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
            gstrSQL = " Select Nvl(可用数量,0)/" & dbl比例系数 & " As 可用数量,Nvl(实际数量,0)/" & dbl比例系数 & " As 实际数量,平均成本价," & _
                      " Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价," & _
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
                    
                    If Dbl数量 <> mshBill.TextMatrix(lngRow, mconIntCol实际数量) Then
                        mshBill.TextMatrix(lngRow, mconintCol真实数量) = zlStr.FormatEx(Dbl数量 * dbl比例系数, mintNumberDigit, , True)
                    End If
                    
                    mshBill.TextMatrix(lngRow, mconIntCol填写数量) = GetFormat(Dbl数量, mintNumberDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol实际数量) = GetFormat(Dbl数量, mintNumberDigit)
                     
                    If Trim(mshBill.TextMatrix(lngRow, mconIntCol实际数量)) = "" Then mshBill.TextMatrix(lngRow, mconIntCol实际数量) = 0
                    
                    mshBill.TextMatrix(lngRow, mconIntCol实际差价) = GetFormat(rsCheck!实际差价, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol实际金额) = GetFormat(rsCheck!实际金额, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol可用数量) = GetFormat(rsCheck!可用数量, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol售价) = GetFormat(IIf(bln时价, dbl现价_时价, dbl现价), mintPriceDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol售价金额) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol售价)) * Dbl数量, mintMoneyDigit)
'                    mshBill.TextMatrix(lngRow, mconintCol差价) = FormatEx(Get出库差价(Val(cboStock.ItemData(cboStock.ListIndex)), lng药品ID, rsCheck!批次, rsCheck!实际金额, rsCheck!实际差价, Val(mshBill.TextMatrix(lngRow, mconIntCol售价金额)), Val(mshBill.TextMatrix(lngRow, mconIntCol实际数量)) * dbl比例系数), mintMoneyDigit)
'                    mshBill.TextMatrix(lngRow, mconIntCol采购金额) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol售价金额)) - Val(mshBill.TextMatrix(lngRow, mconintCol差价)), mintMoneyDigit)
                    
                    If Dbl数量 <> 0 Then
                        mshBill.TextMatrix(lngRow, mconIntCol采购价) = GetFormat(rsCheck!平均成本价 * dbl比例系数, mintCostDigit)
                    End If
                    mshBill.TextMatrix(lngRow, mconIntCol采购金额) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol采购价)) * Dbl数量, mintMoneyDigit)
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
                mshBill.TextMatrix(lngRow, mconIntCol实际数量) = 0
                mshBill.TextMatrix(lngRow, mconIntCol售价金额) = ""
                mshBill.TextMatrix(lngRow, mconIntCol采购金额) = ""
                mshBill.TextMatrix(lngRow, mconintCol差价) = ""
            End If
        Else
            gstrSQL = " Select Nvl(可用数量,0)/" & dbl比例系数 & " As 可用数量,Nvl(实际数量,0)/" & dbl比例系数 & " As 实际数量," & _
                      " Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价," & _
                      " Nvl(批次,0) 批次,上次批号 批号,to_char(效期,'yyyy-MM-dd') 效期,上次产地 产地,NVL(上次供应商ID,0) 上次供应商ID,批准文号 " & _
                      " From 药品库存 Where 性质=1 And 库房ID=[1] And 药品ID=[2] And Nvl(批次,0)=[3] And Nvl(可用数量,0)>0 "
                      
            If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
                gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
            ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 1 Then
                gstrSQL = gstrSQL & " Order by 效期,Nvl(批次,0)"
            ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 2 Then
                gstrSQL = gstrSQL & " Order by 上次批号,Nvl(批次,0)"
            Else
                gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
            End If

            Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提取该药品在指定库存的所有库存记录]", lng库房ID, lng药品ID, lng批次)
            
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
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function

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
                
                If lngCurRow > 0 Then Exit Do
                lngRow = lngRow + 1
            Else
                Exit Do
            End If
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
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

'
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
        
        If CheckStock = False Then Exit Sub '检查是否分解和实际数量
        
        mstrTime_End = GetBillInfo(7, mstr单据号)
        If mstrTime_End = "" Then
            MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not 检查单价(7, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        
        '留存检查
        If mint领用方式 = 1 Then
            If Check留存 = False Then
                MsgBox str留存不足提示
                Exit Sub
            End If
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        '如果审核时修改了单据，则重新生成单据保存
        If mblnChange Or mblnAutoExp = True Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If Not SaveCheck Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        
        If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品领用)) = 1 Then
            '打印
            If IsHavePrivs(mstrPrivs, "单据打印") Then
                printbill
            End If
        End If
        
        gcnOracle.CommitTrans
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then '冲销
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
        If Not 检查单价(7, txtNo, False) And Not mblnUpdate Then
            '以最新的价格更新单据体，退出的目的是让用户看一下最终的单据
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '留存检查
        If mint领用方式 = 1 Then
            If Check留存 = False Then
                MsgBox str留存不足提示, vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If mint编辑状态 = 1 Then '新增保存时，判断价格是否已经更新
        If 检查价格 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '留存检查
        If mint领用方式 = 1 Then
            If Check留存 = False Then
                MsgBox str留存不足提示, vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.药品领用)) = 1 Then
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

    txtDraw.Text = ""
    txtDraw.Tag = "0"
    txt摘要.Text = ""
    txtDraw.SetFocus
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
    
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
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
    gstrSQL = " Select A.药品ID,'['||B.编码||']'||B.名称 通用名,A.药库分批,A.药房分批,B.是否变价" & _
              " From 药品规格 A,收费项目目录 B " & _
              " Where A.药品ID=B.ID And A.药品ID in(select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) "

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

    gstrSQL = "Select a.药品id, Nvl(a.批次, 0) As 批次, Sum(Nvl(实际数量, 0)) As 实际数量 " & _
        " From 药品库存 A, 药品规格 C " & _
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
                dbl实际数量 = Val(GetFormat(Nvl(rsCheck!实际数量, 0) / dbl比例系数, mintNumberDigit))
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsEnterStock As New Recordset
    
    mblnUpdate = False
    mblnEnterCell = False
    mint留存方式 = Val(zlDataBase.GetPara("按月留存领用", glngSys, 模块号.药品领用))
    mintSelectStock = Val(zlDataBase.GetPara("是否选择库房", glngSys, 模块号.药品领用))
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品领用管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    mlng出库库房 = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng出库库房, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    initCard
    
    mstrTime_Start = GetBillInfo(7, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '根据系统参数决定药房人员查看单据时，是否显示成本价
    mshBill.ColWidth(mconIntCol采购价) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol采购金额) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    mblnEnterCell = True
    mblnChange = False
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim str批次 As String, strArray As String
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPriceDigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    
    '库房
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品领用)
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
            Txt填制日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            If Not mblnStock Then
                Me.txtDraw.Tag = UserInfo.部门ID
                Me.txtDraw.Text = UserInfo.部门名称
                Call GetDrawPerson(UserInfo.部门ID)
            End If
            
            initGrid
        Case 2, 3, 4, 6
            initGrid
            
            If mint编辑状态 = 4 Then
                gstrSQL = "select b.id,b.名称 from 药品收发记录 a,部门表 b where a.库房id=b.id and A.单据 = 7 and a.no=[1]"
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
                    strUnitQuantity = "F.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                Case mconint门诊单位
                    strUnitQuantity = "B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,(A.实际数量 / B.门诊包装) AS 实际数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,B.门诊包装 as 比例系数,"
                Case mconint住院单位
                    strUnitQuantity = "B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,(A.实际数量 / B.住院包装) AS 实际数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,B.住院包装 as 比例系数,"
                Case mconint药库单位
                    strUnitQuantity = "B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,(A.实际数量 / B.药库包装) AS 实际数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,B.药库包装 as 比例系数,"
            End Select
            
            If mint编辑状态 <> 6 Then
                gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名," & _
                    "     NVL(E.名称,F.名称) 名称,B.药品来源,B.基本药物,F.规格,F.产地 AS 原产地,A.产地, A.批号,Nvl(A.批次,0) As 批次,B.指导差价率,A.效期," & _
                    strUnitQuantity & _
                    "     A.成本金额,A.零售金额, A.差价,A.摘要,填制人,填制日期,审核人,审核日期," & _
                    "     A.库房ID,A.对方部门ID,C.名称 AS 领用部门,F.是否变价,B.药房分批 As 药房分批核算,NVL(A.领用人,'') As 领用人,A.批准文号,A.发药方式,A.实际数量 原始数量 " & _
                    "     FROM 药品收发记录 A, 药品规格 B,收费项目别名 E ,收费项目目录 F,部门表 C " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=F.ID " & _
                    "     AND B.药品ID=E.收费细目ID(+) And E.性质(+)=3 " & _
                    "     AND A.对方部门ID=C.ID AND A.记录状态 =[3] " & _
                    "     AND A.单据 = 7 AND A.NO = [1]) W,药品库存 Z" & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " And Z.库房ID(+)=[2] And Z.性质(+)=1" & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, E.名称 As 商品名," & _
                    "     NVL(E.名称,F.名称) 名称,B.药品来源,B.基本药物,F.规格,F.产地 AS 原产地,A.产地, A.批号,NVL(A.批次,0) 批次,B.指导差价率,A.效期," & _
                    strUnitQuantity & _
                    "     A.成本金额,0 零售金额,0 差价,A.摘要,A.库房ID,A.对方部门ID,C.名称 AS 领用部门,F.是否变价,B.药房分批 AS 药房分批核算,A.领用人,A.批准文号,A.发药方式,A.填写数量 原始数量 " & _
                    "     FROM " & _
                    "         (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,0 实际数量,SUM(成本金额) AS 成本金额,药品ID,序号,产地, 批号,效期,NVL(批次,0) 批次,扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID,NVL(X.领用人,'') As 领用人,X.批准文号,X.发药方式 " & _
                    "         FROM 药品收发记录 X " & _
                    "         WHERE NO=[1] AND 单据=7  " & _
                    "         GROUP BY 药品ID,序号,产地,批号,效期,NVL(批次,0),扣率,成本价,零售价,摘要,库房ID,对方部门ID,入出类别ID,领用人,批准文号,发药方式" & _
                    "         HAVING SUM(实际数量)<>0 ) A," & _
                    "         药品规格 B,收费项目别名 E ,收费项目目录 F,部门表 C " & _
                    "     WHERE A.药品ID = B.药品ID AND b.药品ID=F.ID AND A.对方部门ID=C.ID " & _
                    "     AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 ) W,药品库存 Z" & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Nvl(Z.批次(+),0) " & _
                    " And Z.库房ID(+)=[2] And Z.性质(+)=1" & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, cboStock.ItemData(cboStock.ListIndex), mint记录状态)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint编辑状态
            Case 2, 6
                Txt填制人 = UserInfo.用户姓名
                Txt填制日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint编辑状态 = 6 Then
                    Txt审核人 = UserInfo.用户姓名
                    Txt审核日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
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
            
            txtDraw.Tag = rsInitCard!对方部门id
            txtDraw.Text = rsInitCard!领用部门
            
            mint领用方式 = IIf(IsNull(rsInitCard!发药方式), 0, rsInitCard!发药方式)
            LblTitle.Caption = GetUnitName & "药品领用单" & IIf(mint领用方式 = 0, "(库房领用)", "(留存领用)")
            
            Call GetDrawPerson(txtDraw.Tag)
            cboDrawPerson.Text = IIf(IsNull(rsInitCard!领用人), "", rsInitCard!领用人)
            
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
                    .TextMatrix(intRow, mconIntCol效期) = IIf(IsNull(rsInitCard!效期), "", Format(rsInitCard!效期, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And .TextMatrix(intRow, mconIntCol效期) <> "" Then
                        '换算为有效期
                        .TextMatrix(intRow, mconIntCol效期) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntCol效期)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol填写数量) = GetFormat(rsInitCard!填写数量, intNumberDigit)
                    .TextMatrix(intRow, mconIntCol实际数量) = GetFormat(rsInitCard!实际数量, intNumberDigit)
                    .TextMatrix(intRow, mconIntCol采购价) = GetFormat(rsInitCard!成本价, intCostDigit)
                    .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(IIf(mint编辑状态 = 6, 0, rsInitCard!成本金额), intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol售价) = GetFormat(rsInitCard!零售价, intPriceDigit)
                    .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(rsInitCard!零售金额, intMoneyDigit)
                    .TextMatrix(intRow, mconintCol差价) = GetFormat(rsInitCard!差价, intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol指导差价率) = rsInitCard!指导差价率 & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol可用数量) = IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量)
                    .TextMatrix(intRow, mconIntCol实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mconIntCol实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    .TextMatrix(intRow, mconintCol真实数量) = IIf(IsNull(rsInitCard!原始数量), "0", rsInitCard!原始数量)
                    .TextMatrix(intRow, mconintCol原始数量) = .TextMatrix(intRow, mconIntCol实际数量)
                    
                    If mint编辑状态 = 2 Or mint编辑状态 = 3 Then
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
            
            If mint编辑状态 = 3 Then    '审核的情况下
                Call ShowColor
            End If
    End Select
    Call RefreshRowNO(mshBill, mconIntCol行号, 1)
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
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
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconIntCol填写数量) = IIf(mint编辑状态 = 6, "数量", "填写数量")
        .TextMatrix(0, mconIntCol实际数量) = IIf(mint编辑状态 = 6, "冲销数量", "实际数量")
        .TextMatrix(0, mconIntCol采购价) = "成本价"
        .TextMatrix(0, mconIntCol采购金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconintCol差价) = "差价"
        .TextMatrix(0, mconIntCol可用数量) = "可用数量"
        .TextMatrix(0, mconIntCol实际差价) = "实际差价"
        .TextMatrix(0, mconIntCol实际金额) = "实际金额"
        .TextMatrix(0, mconIntCol指导差价率) = "指导差价率"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntCol批次) = "批次"
        .TextMatrix(0, mconintCol真实数量) = "真实数量"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconintCol原始数量) = "原始数量"
        
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
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconIntCol填写数量) = 1000
        .ColWidth(mconIntCol实际数量) = 1000
        .ColWidth(mconIntCol采购价) = 900
        .ColWidth(mconIntCol采购金额) = 900
        .ColWidth(mconIntCol售价) = 900
        .ColWidth(mconIntCol售价金额) = 900
        .ColWidth(mconintCol差价) = 800
        .ColWidth(mconIntCol可用数量) = 0
        .ColWidth(mconIntCol实际差价) = 0
        .ColWidth(mconIntCol实际金额) = 0
        .ColWidth(mconIntCol指导差价率) = 0
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconIntCol批次) = 0
        .ColWidth(mconintCol真实数量) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        .ColWidth(mconintCol原始数量) = 0
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconintCol真实数量) = 5
        .ColData(mconintCol原始数量) = 5
        
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txtDraw.Enabled = True
            cmdDraw.Enabled = True
            txt摘要.Enabled = True
            If mintSelectStock = 0 And mblnStock Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol填写数量) = 4
            .ColData(mconIntCol实际数量) = 5
        ElseIf mint编辑状态 = 3 Or mint编辑状态 = 6 Then
            cboDrawPerson.Enabled = False
            
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 4
        ElseIf mint编辑状态 = 4 Then
            cboDrawPerson.Enabled = False
        
            cboStock.Enabled = False
            
            txtDraw.Enabled = False
            cmdDraw.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 5
            
        End If
        
        .ColData(mconIntCol采购价) = 5
        .ColData(mconIntCol采购金额) = 5
        .ColData(mconIntCol售价) = 5
        .ColData(mconIntCol售价金额) = 5
        .ColData(mconintCol差价) = 5
        .ColData(mconIntCol可用数量) = 5
        .ColData(mconIntCol实际差价) = 5
        .ColData(mconIntCol实际金额) = 5
        .ColData(mconIntCol指导差价率) = 5
        .ColData(mconIntCol比例系数) = 5
        .ColData(mconIntCol批次) = 5
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol填写数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol实际数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol差价) = flexAlignRightCenter
        .ColAlignment(mconintCol真实数量) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
    End With
    txt摘要.MaxLength = GetLength("药品收发记录", "摘要")
    chkIn.Visible = (mint编辑状态 = 1)
    txtIn.Visible = (mint编辑状态 = 1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Width < 11500 Then
        Me.Width = 11500
        Exit Sub
    End If
    
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
    cboDrawPerson.Left = mshBill.Left + mshBill.Width - cboDrawPerson.Width
    lblDrawPerson.Left = cboDrawPerson.Left - lblDrawPerson.Width - 100
    
    LblEnterStock.Left = cboStock.Left + cboStock.Width + (lblDrawPerson.Left - cboStock.Left - cboStock.Width - LblEnterStock.Width - txtDraw.Width - cmdDraw.Width - 100) / 2
    txtDraw.Left = LblEnterStock.Left + LblEnterStock.Width + 100
    cmdDraw.Left = txtDraw.Left + txtDraw.Width
    
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
        .Left = CmdSave.Left - CmdSave.Width - 500
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    With mshProvider
        If .Visible = True Then
            If .Tag = 0 Then
                .Width = LblEnterStock.Width + txtDraw.Width + cmdDraw.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
            Else
                .Width = lblDrawPerson.Width + cboDrawPerson.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = lblDrawPerson.Left
            End If
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品领用管理", "药品名称显示方式", mintDrugNameShow)
    
    mblnAutoExp = False
    If mshProvider.Visible = True Then
        mshProvider.Visible = False
        txtDraw.SetFocus
        txtDraw.SelLength = Len(txtDraw.Text)
        txtDraw.SelStart = 0
        Cancel = True
        Exit Sub
    End If
    
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
    Dim rs类别 As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng库房ID As Long
    Dim lng对方部门id As Long
    Dim str审核人 As String
    Dim dat审核日期 As String
    Dim int序号 As Integer
    Dim lng药品ID As Long
    Dim str产地 As String
    Dim lng批次 As Long
    Dim num填写数量 As Double
    Dim num实际数量 As Double
    Dim num成本价 As Double
    Dim num成本金额 As Double
    Dim num零售金额 As Double
    Dim num差价 As Double
    Dim lng入出类别id As Long
    Dim str批号 As String
    Dim dat效期 As String
    Dim arrSql As Variant
    Dim str批准文号 As String
    Dim n As Integer
    
    mblnSave = False
    SaveCheck = False
    arrSql = Array()
    
    On Error GoTo errHandle
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = txtDraw.Tag
    str审核人 = UserInfo.用户姓名
    strNo = txtNo.Tag
    gstrSQL = "SELECT b.id " _
            & " FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID " _
            & "  AND a.单据 = 7 "
    Call SQLTest(App.Title, "药品领用单", gstrSQL)
    Set rs类别 = zlDataBase.OpenSQLRecord(gstrSQL, "SaveCheck")
    Call SQLTest
    
    If rs类别.EOF Then
        MsgBox "对不起，没有设置药品领用的入出类别，请检查药品入出分类!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng入出类别id = rs类别!id
    rs类别.Close
    
    dat审核日期 = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If .TextMatrix(intRow, 0) <> "" Then
                
                lng药品ID = .TextMatrix(intRow, 0)
                str产地 = .TextMatrix(intRow, mconIntCol产地)
                lng批次 = .TextMatrix(intRow, mconIntCol批次)
                
                If Val(.TextMatrix(intRow, mconIntCol填写数量)) = Val(.TextMatrix(intRow, mconIntCol实际数量)) Then
                    num填写数量 = Val(.TextMatrix(intRow, mconintCol真实数量))
                    num实际数量 = Val(.TextMatrix(intRow, mconintCol真实数量))
                Else
                    num填写数量 = .TextMatrix(intRow, mconIntCol填写数量) * .TextMatrix(intRow, mconIntCol比例系数)
                    num实际数量 = .TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数)
                End If
                
'                num成本价 = FormatEx(.TextMatrix(intRow, mconIntCol采购价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                num成本价 = Get成本价(lng药品ID, lng库房ID, lng批次)
                
                num成本金额 = .TextMatrix(intRow, mconIntCol采购金额)
                num零售金额 = .TextMatrix(intRow, mconIntCol售价金额)
                num差价 = .TextMatrix(intRow, mconintCol差价)
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                dat效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                
                int序号 = Val(.TextMatrix(intRow, mconIntCol序号))
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                
                gstrSQL = "zl_药品领用_Verify("
                '序号
                gstrSQL = gstrSQL & int序号
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '库房ID
                gstrSQL = gstrSQL & "," & lng库房ID
                '对方部门ID
                gstrSQL = gstrSQL & "," & lng对方部门id
                '药品ID
                gstrSQL = gstrSQL & "," & lng药品ID
                '产地
                gstrSQL = gstrSQL & ",'" & str产地 & "'"
                '批次
                gstrSQL = gstrSQL & "," & lng批次
                '填写数量
                gstrSQL = gstrSQL & "," & num填写数量
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
                '审核人
                gstrSQL = gstrSQL & ",'" & str审核人 & "'"
                '审核日期
                gstrSQL = gstrSQL & ",to_date('" & dat审核日期 & "','yyyy-mm-dd HH24:MI:SS')"
                '批号
                gstrSQL = gstrSQL & ",'" & str批号 & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(dat效期 = "", "Null", "to_date('" & Format(dat效期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '领用方式
                gstrSQL = gstrSQL & "," & mint领用方式
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lng药品ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    If Not ExecuteSql(arrSql, MStrCaption, False, False) Then Exit Function
   
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '外挂功能
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, lng库房ID, strNo, 单据号.药品领用)
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
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
    Dim i As Integer
    Dim arrSql As Variant
    
    SaveStrike = False
    arrSql = Array()
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol填写数量)), Val(.TextMatrix(intRow, mconIntCol实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    
        NO_IN = Trim(txtNo.Tag)
        填制人_IN = UserInfo.用户姓名
        填制日期_IN = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        原记录状态_IN = mint记录状态
        
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
                If Val(.TextMatrix(intRow, mconIntCol填写数量)) = Val(.TextMatrix(intRow, mconIntCol实际数量)) Then
                    冲销数量_IN = Val(.TextMatrix(intRow, mconintCol真实数量))
                Else
                    冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                序号_IN = .TextMatrix(intRow, mconIntCol序号)
                
                gstrSQL = "ZL_药品领用_STRIKE("
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
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol药名) = 0 Then
        'Cancel = True    '等待加CANCEL参数
        Exit Sub
    End If
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
    Dim RecReturn As Recordset
    Dim str药品ID As String
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow  As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
'    Set RecReturn = Frm药品选择器.ShowME(Me, 2,cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), True, True, False, False, True, 0, False, mint领用方式)
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
    End If
    
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), 0, True, True, True, False, , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
    End If
    
    mshBill.CmdEnable = True
    
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        With mshBill
            For i = 1 To RecReturn.RecordCount
                intRow = .Row
'            If RecReturn.RecordCount = 1 Then
                .TextMatrix(intRow, mconIntCol行号) = .Row
                SetColValue .Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                    Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                    IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
                .Col = mconIntCol填写数量
'            End If
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
        Select Case .Col
            Case mconIntCol填写数量, mconIntCol实际数量
                intDigit = mintNumberDigit
            Case mconIntCol采购价
               intDigit = mintCostDigit
            Case mconIntCol售价
                intDigit = mintPriceDigit
            Case mconIntCol采购金额, mconIntCol售价金额
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
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnterCell Then Exit Sub
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
            Case mconIntCol填写数量, mconIntCol实际数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
                Call 提示库存数
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow  As Integer
    
    intOldRow = mshBill.Row
    
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
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), strkey, sngLeft, sngTop, True, True, False, False, True, 0, False, mint领用方式)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), Val(txtDraw.Tag))
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), , Val(txtDraw.Tag), 0, True, True, True, False, , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn) '将重复的记录和时价无库存的药品过滤掉
                    End If
                                        
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                    Nvl(RecReturn!药品来源), "" & RecReturn!基本药物, _
                                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                    IIf(IsNull(RecReturn!效期), "", Format(RecReturn!效期, "yyyy-MM-dd")), _
                                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)) = False Then
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
                        .Text = .TextMatrix(.Row, .Col)
                        Cancel = True
                    End If
                    Call 提示库存数
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
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) = 0 And mint编辑状态 <> 3 And mint编辑状态 <> 6 Then  '冲销数量可以为0
                        MsgBox "对不起，数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strkey) < 0 Then
                        If Not IsHavePrivs(mstrPrivs, "负数开单") Then
                            MsgBox "对不起，你没有负数开单的权限，请重输！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
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
                    
                    If .TextMatrix(.Row, 0) = "" Then Exit Sub
                    
                    If gtype_UserSysParms.P175_药品领用明确批次 = 1 Then
                        If Not CompareUsableQuantity(.Row, strkey) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If

                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol售价) * strkey, mintMoneyDigit)
                    End If
                    
'                    .TextMatrix(.Row, mconintCol差价) = FormatEx(Get出库差价(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(.TextMatrix(.Row, mconIntCol实际金额)), Val(.TextMatrix(.Row, mconIntCol实际差价)), Val(.TextMatrix(.Row, mconIntCol售价金额)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol比例系数))), mintMoneyDigit)
                    
                    If strkey <> 0 And (mint编辑状态 = 1 Or mint编辑状态 = 2) Then
'                        .TextMatrix(.Row, mconIntCol采购价) = FormatEx((.TextMatrix(.Row, mconIntCol售价金额) - .TextMatrix(.Row, mconintCol差价)) / strkey, mintCostDigit)
                        .TextMatrix(.Row, mconIntCol采购价) = GetFormat(Get成本价(Val(.TextMatrix(.Row, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, mconIntCol批次))) * Val(Val(mshBill.TextMatrix(.Row, mconIntCol比例系数))), mintCostDigit)
                    End If
                    .TextMatrix(.Row, mconIntCol采购金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol采购价)) * strkey, mintMoneyDigit)
                    .TextMatrix(.Row, mconintCol差价) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol采购金额)), mintMoneyDigit)
                    
                    If .Col = mconIntCol填写数量 Then
                        .TextMatrix(.Row, mconIntCol实际数量) = strkey
                    End If
                End If
                显示合计金额
            
        End Select
    End With
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, _
    ByVal str药品来源 As String, ByVal str基本药物 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal num可用数量 As Double, ByVal num实际金额 As Double, _
    ByVal num实际差价 As Double, ByVal num指导差价率 As Double, _
    ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal str批准文号 As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    Dim str药名 As String
    
    SetColValue = False
    
    '检查是否重复
'    If Not CheckRepeatMedicine(mshBill, lng药品ID & "," & "0" & "|" & lng批次 & "," & mconIntCol批次, intRow) Then
'        Exit Function
'    End If
    
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
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, mconIntCol售价) = GetFormat(num售价 * num比例系数, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol可用数量) = FormatEx(num可用数量, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntCol指导差价率) = num指导差价率 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        
        If gtype_UserSysParms.P175_药品领用明确批次 = 0 Then
            .TextMatrix(intRow, mconIntCol批次) = 0
        Else
            .TextMatrix(intRow, mconIntCol批次) = lng批次
        End If
        
        If int是否变价 = 1 Then
            dblPrice = Get售价(True, lng药品ID, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol批次)))  'GetPrice(lng药品ID, lng批次, num比例系数)
            .TextMatrix(intRow, mconIntCol售价) = GetFormat(dblPrice * num比例系数, mintPriceDigit)
        End If
        If IsLowerLimit(cboStock.ItemData(cboStock.ListIndex), lng药品ID) Then Call SetForeColor_ROW(mlng紫色)
        Call CheckLapse(str效期)
        
    End With
    SetColValue = True
End Function

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub mshProvider_DblClick()
    mshProvider_KeyDown vbKeyReturn, 0
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mshProvider.Visible = False
        If mshProvider.Tag = 0 Then
            txtDraw.SetFocus
            txtDraw.SelStart = 0
            txtDraw.SelLength = Len(txtDraw.Text)
        Else
            cboDrawPerson.SetFocus
            cboDrawPerson.SelStart = 0
            cboDrawPerson.SelLength = Len(cboDrawPerson.Text)
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If mshProvider.Tag = 0 Then
            txtDraw.Text = mshProvider.TextMatrix(mshProvider.Row, 3)
            txtDraw.Tag = mshProvider.TextMatrix(mshProvider.Row, 0)
            mshProvider.Visible = False
            Call GetDrawPerson(txtDraw.Tag)
            cboDrawPerson.SetFocus
        Else
            cboDrawPerson.Text = mshProvider.TextMatrix(mshProvider.Row, 1)
            mshBill.SetFocus
            mshBill.Col = 1
            mshBill.Row = 1
        End If
    End If
    
End Sub

Private Sub mshProvider_LostFocus()
    If mshProvider.Visible Then
        mshProvider.Visible = False
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
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            If Val(txtDraw.Tag) = 0 Then
                If Trim(txtDraw.Text) = "" Then
                    MsgBox "对不起，领药部门不能为空！", vbOKOnly + vbInformation, gstrSysName
                    txtDraw.SetFocus
                    Exit Function
                Else
                    MsgBox "对不起，没有你输入的领药部门！", vbOKOnly + vbInformation, gstrSysName
                    txtDraw.SetFocus
                    Exit Function
                End If
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
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol实际数量))) = "" Then
                        MsgBox "第" & intLop & "行药品的数量为空了，请检查！", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol实际数量
                        Exit Function
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
                    
                    If gtype_UserSysParms.P175_药品领用明确批次 = 1 Then
                        If Not CompareUsableQuantity(intLop, Val(.TextMatrix(intLop, mconIntCol实际数量))) Then
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = IIf(.ColData(mconIntCol填写数量) = 4, mconIntCol填写数量, mconIntCol实际数量)
                            Exit Function
                        End If
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
    Dim lng入出类别id As Long
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
    Dim rs入出类别 As New Recordset
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim n As Integer
    Dim blnTran As Boolean
    
    SaveCard = False
    arrSql = Array()
    
    '在外面设置入出类别ID，主要是所有药品都要用他
    gstrSQL = "SELECT b.id " _
             & "FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID " _
              & "AND a.单据 = 7 " _
              & "AND b.系数 = -1 " _
              & "AND ROWNUM < 2"
    Call zlDataBase.OpenRecordset(rs入出类别, gstrSQL, "取入出类别")
    If rs入出类别.EOF Then
        MsgBox "对不起，没有设置药品领用的出库类别，请检查药品入出分类！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lng入出类别id = rs入出类别.Fields(0)
    rs入出类别.Close
    
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = zlDataBase.GetNextNo(27, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lngEnterStockID = txtDraw.Tag
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        datBookDate = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Txt审核人
        On Error GoTo errHandle
        
        If bln强制保存 Then blnTran = True
        
        If mint编辑状态 = 2 Or bln强制保存 Then        '修改
            gstrSQL = "zl_药品领用_Delete('" & mstr单据号 & "')"
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
                datTimeLimit = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                    '换算为失效期来保存
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                If Val(.TextMatrix(intRow, mconintCol原始数量)) = Val(.TextMatrix(intRow, mconIntCol填写数量)) Then
                    dblQuantity = Val(.TextMatrix(intRow, mconintCol真实数量))
                Else
                    dblQuantity = FormatEx(.TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                                
'                dblPurchasePrice = FormatEx(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                dblPurchasePrice = Get成本价(lngDrugID, lngStockID, lngBatchID)
                
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol采购金额)
                
                dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dblSalePrice = Get售价(Split(.TextMatrix(intRow, mconIntCol指导差价率), "||")(1) = 1, lngDrugID, lngStockID, lngBatchID)
                
                dblSaleMoney = .TextMatrix(intRow, mconIntCol售价金额)
                dblMistakePrice = .TextMatrix(intRow, mconintCol差价)
                
'                If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                End If
                If mint编辑状态 = 3 Then
                    lngSerial = .TextMatrix(intRow, mconIntCol序号)
                Else
                    lngSerial = intRow
                End If
                
                str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                
                gstrSQL = "zl_药品领用_INSERT("
                '入出类别ID
                gstrSQL = gstrSQL & lng入出类别id
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
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
                '成本价
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '成本金额
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '售价
                gstrSQL = gstrSQL & "," & dblSalePrice
                '售价金额
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '差价
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '填制人
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '产地
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '摘要
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '领用人
                gstrSQL = gstrSQL & ",'" & cboDrawPerson.Text & "'"
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '领用方式
                gstrSQL = gstrSQL & "," & mint领用方式
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
            " Where a.单据 = 7 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPriceDigit & ") <> Round(b.现价, " & intPriceDigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0  and b.执行日期>a.填制日期" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 7 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & intPriceDigit & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & intPriceDigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次,  0 原价,b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 7 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(b.平均成本价," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1  and b.性质=1" & _
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
                dbl零售价 = Val(FormatEx(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intPriceDigit))
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

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(.Row, mconIntCol药名) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        If mint领用方式 = 0 Then
            gstrSQL = "select 可用数量 from 药品库存 where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1 and " _
                & " nvl(批次,0)=[3]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
        Else
            gstrSQL = "select 可用数量 from 药品留存 where 期间=[1] and 库房id=[2] " _
                & " and 药品id=[3] And 科室ID=[4] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[提示库存数]", Format(zlDataBase.Currentdate(), IIf(mint留存方式 = 0, "yyyy", "yyyymm")), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(txtDraw.Tag))
        End If
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol可用数量) = 0
        Else
            .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0) / Val(.TextMatrix(.Row, mconIntCol比例系数)))
        End If
        rsUseCount.Close
        
        staThis.Panels(2).Text = "该药品当前库存数为[" & FormatEx(.TextMatrix(.Row, mconIntCol可用数量), mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol单位)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDraw_LostFocus()
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub txtDraw_Validate(Cancel As Boolean)
    If txtDraw.Text = "" Then
        txtDraw.Tag = "0"
        Exit Sub
    End If
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
    On Error GoTo ErrHand
    Dim int包装系数 As Integer
    Dim lng药品ID As Long
    Dim blnInput As Boolean
    
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
                MsgBox "药品[" & !药品名称 & "]未在" & cboStock.Text & "中设置存储属性，将不能领用！"
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
                        Select Case mint库存检查
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
                    Nvl(!批号), Nvl(!效期), Nvl(!可用数量, 0), Nvl(!实际金额, 0), Nvl(!实际差价, 0), _
                    Nvl(!指导差价率, 0), int包装系数, Nvl(!批次, 0), !时价, _
                    !药房分批, IIf(IsNull(!批准文号), "", !批准文号)) Then
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
    Dim numUsedCount As Double, dbltotal As Double
    Dim vardrug As Variant, intLop As Integer
    Dim dbl原填写数量 As Double
    Dim rsUseCount As New Recordset
    
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    
    CompareUsableQuantity = False

    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        '获取库存可用数量
        If mint领用方式 = 0 Then
            gstrSQL = "select 可用数量 from 药品库存 where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1 and " _
                & " nvl(批次,0)=[3]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[CompareUsableQuantity]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol批次)))
        Else
            gstrSQL = "select 可用数量 from 药品留存 where 期间=[1] and 库房id=[2] " _
                & " and 药品id=[3] And 科室ID=[4] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[CompareUsableQuantity]", Format(zlDataBase.Currentdate(), IIf(mint留存方式 = 0, "yyyy", "yyyymm")), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(txtDraw.Tag))
        End If
        If rsUseCount.EOF Then
            dblUsableQuantity = GetFormat(0, mintNumberDigit)
        Else
            dblUsableQuantity = GetFormat(IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0) / Val(.TextMatrix(intRow, mconIntCol比例系数))), mintNumberDigit)
        End If
        rsUseCount.Close
        
        '数量比较
        If .TextMatrix(intRow, mconIntCol批次) > 0 Or Split(.TextMatrix(intRow, mconIntCol指导差价率), "||")(1) = 1 Then '对分批或者时价药品检查库存
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
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And .TextMatrix(intRow, mconIntCol批次) = .TextMatrix(intLop, mconIntCol批次) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol实际数量)
                        End If
                    End If
                Next
                
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
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
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
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Or mint编辑状态 = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                If dbl填写数量 > dblUsableQuantity + dbl原填写数量 Then
                    MsgBox "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1305", "zl8_bill_1305"), mint记录状态, int单位系数, 1305, "药品领用单", strNo
End Sub


Private Sub txtDraw_Change()
    With txtDraw
        .Text = UCase(.Text)
        .SelStart = Len(.Text)
    End With
    mblnChange = True
End Sub

Private Sub txtDraw_GotFocus()
    txtDraw.SelStart = 0
    txtDraw.SelLength = Len(txtDraw.Text)
End Sub

Private Sub txtDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String, str站点限制 As String
    Dim adoProvider As New Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑状态 = 3 Or mint编辑状态 = 4 Then Exit Sub
    
    On Error GoTo errHandle
    With txtDraw
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        str站点限制 = GetDeptStationNode(cboStock.ItemData(cboStock.ListIndex))
        
        If mblnStock Then
            gstrSQL = "SELECT DISTINCT a.id,a.编码,a.简码,a.名称 " _
                    & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                    & "Where " & IIf(str站点限制 <> "", "(a.站点 = [4] or a.站点 is null) And ", "") & "c.工作性质 = b.名称 " _
                    & "  AND b.编码 = 'O' AND a.id = c.部门id " _
                    & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                    & "  And (a.简码 like [1] Or a.编码 like [1] or a.名称 like [1]) " _
                    & "Order By a.编码"
        Else
            gstrSQL = " Select C.ID " & _
                " From 部门性质说明 A,部门性质分类 B,部门表 C " & _
                " Where " & IIf(str站点限制 <> "", "(C.站点 = [4] or C.站点 is null) And ", "") & "A.工作性质=B.名称 And A.部门ID=C.ID " & _
                "   AND TO_CHAR(C.撤档时间, 'yyyy-MM-dd')='3000-01-01' And B.编码='O'" & _
                "   And C.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])"
                
            '只提取设置了领用流向的科室
            gstrSQL = "SELECT DISTINCT a.id,a.编码,a.简码,a.名称 " _
                 & " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                 & " Where " & IIf(str站点限制 <> "", "(a.站点 = [4] or a.站点 is null) And ", "") & " c.工作性质 = b.名称 " _
                 & "   AND b.编码 ='O' AND a.id = c.部门id " _
                 & "   AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
                 & "   And (a.简码 like [1] Or a.编码 like [1] or a.名称 like [1])" _
                 & "   And a.ID IN (Select Distinct 领用部门ID From 药品领用控制 Where 对方库房id=[3] And 领用部门ID IN (" & gstrSQL & "))" _
                 & " Order By a.编码 "
        End If
            
        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
            IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", _
            UserInfo.用户ID, _
            cboStock.ItemData(cboStock.ListIndex), _
            str站点限制)
        
        mshProvider.Tag = 0
        
        If adoProvider.EOF Then
            MsgBox "没有你输入的领药部门，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        If adoProvider.RecordCount > 1 Then
            Set mshProvider.Recordset = adoProvider
            Dim intCol As Integer
            Dim intRow As Integer
            
            With mshProvider
                If .Visible = False Then .Visible = True
                .Redraw = False
                .SetFocus
                
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
                .ColWidth(0) = 0
                .ColWidth(1) = 800
                .ColWidth(2) = 1000
                .ColWidth(3) = 2500
                .Row = 1
                .TopRow = 1
                .Col = 0
                .ColSel = .Cols - 1
                
                .Width = LblEnterStock.Width + txtDraw.Width + cmdDraw.Width + 80
                .Top = txtDraw.Top + txtDraw.Height
                .Left = cmdDraw.Left + cmdDraw.Width - .Width
                .Redraw = True
                Exit Sub
            End With
        Else
            .Text = adoProvider!名称
            .Tag = adoProvider!id
        End If
        adoProvider.Close
        Call GetDrawPerson(.Tag)
        cboDrawPerson.SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

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
    
    rsTemp.MoveFirst
    str批次 = ""
    strTemp = ""
    On Error GoTo errHandle
    Do While Not rsTemp.EOF
        If gtype_UserSysParms.P175_药品领用明确批次 = 0 Then
            str批次 = "0"
        Else
            str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        End If
        If InStr(1, strTemp, rsTemp!药品id & "," & str批次) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & str批次 & "," & rsTemp!通用名 & "|"
        End If
        
        If rsTemp!时价 = 1 Then '将时价无库存的记录找出来
            gstrSQL = "select Decode(Nvl(批次,0),0,实际金额/实际数量,Nvl(零售价,实际金额/实际数量))*" & Choose(mintUnit, 1, rsTemp!门诊包装, rsTemp!住院包装, rsTemp!药库包装) & " as  售价 " _
                & "  from 药品库存 " _
                & " where 库房id=[1] " _
                & " and 药品id=[2] " _
                & " and 性质=1 and 实际数量>0 and " _
                & " nvl(批次,0)=[3]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!药品id, IIf(IsNull(rsTemp!批次), 0, rsTemp!批次))
            If rsPrice.EOF Then
                str库存 = str库存 & rsTemp!药品id & "," & rsTemp!通用名 & "|"
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

Private Function GetPrice(ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double) As Double
    Dim rsPrice As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(批次,0),0,实际金额/实际数量,Nvl(零售价,实际金额/实际数量))*" & dbl比例系数 & " as  售价 " _
        & "  from 药品库存 " _
        & " where 库房id=[1] " _
        & " and 药品id=[2] " _
        & " and 性质=1 and 实际数量>0 and " _
        & " nvl(批次,0)=[3]"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)

    If rsPrice.EOF Then
        GetPrice = 0
        Exit Function
    End If
    GetPrice = rsPrice.Fields(0).Value
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
            If mshBill.TextMatrix(i, 0) <> "" And Trim(.TextMatrix(i, mconIntCol填写数量)) <> "" Then
            
                bln是否时价 = Val(Split(.TextMatrix(i, mconIntCol指导差价率), "||")(1)) = 1
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
        
        If intSum > 0 Then '大于0表示有价格更新
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

