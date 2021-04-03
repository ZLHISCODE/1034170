VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmRequestDrugCard 
   Caption         =   "药品申领单"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "frmRequestDrugCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   10785
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd全部复制 
      Caption         =   "全部复制"
      Height          =   350
      Left            =   9360
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmd全清 
      Caption         =   "全部清除"
      Height          =   350
      Left            =   8040
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox chkExportPlan 
      Caption         =   "保存时只同步产生非常备药品的计划单据"
      Height          =   380
      Left            =   5160
      TabIndex        =   29
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   6480
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全冲(&A)"
      Height          =   350
      Left            =   5160
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8040
      TabIndex        =   5
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   6
      Top             =   5520
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   557
         Width           =   1515
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
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
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "差价合计:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "售价金额合计:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "成本金额合计:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Top             =   550
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
         TabIndex        =   16
         Top             =   587
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品申领单"
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
         TabIndex        =   15
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
         Top             =   617
         Width           =   990
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   300
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   4500
         Width           =   720
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
            Picture         =   "frmRequestDrugCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1000
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
            Picture         =   "frmRequestDrugCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   5985
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestDrugCard.frx":22EA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":3080
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
      Left            =   2760
      TabIndex        =   22
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
Attribute VB_Name = "frmRequestDrugCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、通过向导新增；6、接受（接收后记录接收登记人，可以取消错误的接收）；7、拒收
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mbln申领状态 As Boolean
Private mstr库房 As String                  '记录已经添加了的库房

Private mint明确批次 As Integer             '表示在填写申领单时，是否明确药品的批次
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mint库存检查入库库房 As Integer     '仅用于冲销时对原入库库房是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mcolUsedCount As Collection         '已使用的数量集合
Private mstrPrivs As String                     '权限
Private mlngStockID As Long                 '当前用户所选的药房ID
Private mintApplyType As Integer            '申领方式：0-手工申领;1-根据消耗量;2-根据上限;3-根据下限;4-根据上下限;5-根据申领单未发数;6-根据日销售量;7-根据销售总量
Private mstrEndTime As String               '当自动申领方式为7时，返回时间范围中的结束时间
Private rsDepend As New ADODB.Recordset

Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mstrTime_Start As String                        '进入单据编辑界面时，待编辑单据的最大修改时间
Private mstrTime_End As String                        '此刻该编辑单据的最大修改时间

Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称
Private mblnUpdate As Boolean               '用来记录调价审核后是否更新了新价格
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private Const MStrCaption As String = "药品申领管理"

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


Private mbln下可用数量 As Boolean           '是否下发药药房的可用数量
Private recSort As ADODB.Recordset          '按药品ID排序的专用记录集

Private mint处理方式 As Integer             '冲销时：0－正常冲销；1－产生冲销申请单据

Private mbln检查库存 As Boolean

'=========================================================================================
Private Const mconIntCol行号 As Integer = 1
Private Const mconIntCol药名 As Integer = 2
Private Const mconIntCol商品名 As Integer = 3
Private Const mconIntCol来源 As Integer = 4
Private Const mconIntCol基本药物 As Integer = 5
Private Const mconIntCol序号 As Integer = 6
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
Private Const mconintcol当前库存 As Integer = 22
Private Const mconintcol对方库存 As Integer = 23
Private Const mconIntCol申领数量 As Integer = 24
Private Const mconIntCol填写数量 As Integer = 25
Private Const mconIntCol实际数量 As Integer = 26
Private Const mconIntCol采购价 As Integer = 27
Private Const mconIntCol采购金额 As Integer = 28
Private Const mconIntCol售价 As Integer = 29
Private Const mconIntCol售价金额 As Integer = 30
Private Const mconintCol差价 As Integer = 31
Private Const mconIntCol上次供应商ID As Integer = 32
Private Const mconintCol真实数量 As Integer = 33
Private Const mconIntCol药品编码和名称 As Integer = 34
Private Const mconIntCol药品编码 As Integer = 35
Private Const mconIntCol药品名称 As Integer = 36
Private Const mconIntCol常备药品 As Integer = 37
Private Const mconIntCol原始数量 As Integer = 38
Private Const mconIntColS  As Integer = 39             '总列数
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
Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = " Select 审核日期,配药日期,配药人 From 药品收发记录 " & _
            " Where 单据=6 And NO=[1] And 记录状态=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查单据]", strNo)
    
    With rs
        '返回空，表示已经删除
        If .EOF Then
            CheckBill = "该单据已经被其他操作员删除！"
        ElseIf Not IsNull(!审核日期) Then
            CheckBill = "该单据已经被其他操作员审核！"
        ElseIf Not IsNull(!配药日期) Then
            CheckBill = "该单据已经被其他操作员发送！"
        ElseIf Not IsNull(!配药人) Then
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
'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strMsg As String

    GetDepend = False
    On Error GoTo ErrHand

    '检查药品入出类别是否完整
    strMsg = "没有设置药品移库的入库及出库类别，请检查药品入出分类！"
    gstrSQL = "SELECT B.Id,B.系数 " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 6"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "药品移库管理")

    With rsDepend
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "系数=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置药品移库的入库类别，请检查药品入出分类！"
            GoTo ErrHand
        End If
        .Filter = "系数=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置药品移库的出库类别，请检查药品入出分类！"
            GoTo ErrHand
        End If
        .Filter = 0
        
        'gstrSQL = ReturnSQL(mlngStockID, False)
    End With
    'Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "药品申领管理", mlngStockID)
    Set rsDepend = ReturnSQL(mlngStockID, "药品申领管理", False, 1343)

    strMsg = "没有任何库房允许申领，请在[基础参数设置]的药品流向中设置！"
    If rsDepend.RecordCount = 0 Then
        MsgBox strMsg, vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsDepend.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False, Optional lngStockid As Long = 0, Optional int处理方式 As Integer = 0, Optional intApplyType As Integer = 0)
    Dim strSQL As String
    Dim rsPara As New ADODB.Recordset
    
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mint处理方式 = int处理方式
    mintApplyType = intApplyType
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1343)
    mlngStockID = IIf(lngStockid = 0, glngDeptId, lngStockid)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint库存检查入库库房 = MediWork_GetCheckStockRule(mlngStockID)
    
    mint明确批次 = gtype_UserSysParms.P73_明确申领药品批次
    
    If mint编辑状态 <> 5 Then
        Me.cmd全部复制.Visible = False
        Me.cmd全清.Visible = False
    End If
    
    If mint明确批次 = 0 Then
        mint库存检查 = 0
    Else
        mbln检查库存 = True
    End If
    
    mbln下可用数量 = (gtype_UserSysParms.P96_药品填单下可用库存 = 1)
    
    mblnEdit = False
         
    If mint编辑状态 = 5 Then
        Me.Height = Me.Height + Me.cmd全清.Height
    End If
         
    If mint编辑状态 = 1 Or mint编辑状态 = 5 Then
        mblnEdit = True
        mblnFirst = True
        
        chkExportPlan.Visible = True
    
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint编辑状态 = 4 Then
        mblnFirst = True
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not IsHavePrivs(mstrPrivs, "单据打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint编辑状态 = 7 Then
        mblnEdit = False
        mblnFirst = True
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint处理方式 = 1 Then
            CmdSave.Caption = "申请冲销(&O)"
            CmdSave.Width = CmdSave.Width + 200
        Else
            CmdSave.Caption = "冲销(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub

Private Sub cboStock_Click()
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
        Call SetSelectorRS(2, "药品申领管理", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln检查库存)
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

Private Sub cboStock_Change()
    mblnChange = True
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
                    
                    If Me.mshBill.ColWidth(mconIntCol申领数量) > 0 Then
                        Me.mshBill.ColWidth(mconIntCol申领数量) = 0
                        Me.cmd全部复制.Visible = False
                        Me.cmd全清.Visible = False
                        Call Form_Resize
                    End If
                    
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        mint库存检查 = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
        If mint明确批次 = 0 Then
            mint库存检查 = 0
        Else
            mbln检查库存 = True
        End If
        
    End With
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
        FindRow mshBill, mconIntCol药名, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmd全部复制_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("您确定要将申领数量的值作为填写数量和实际数量吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol填写数量) = Me.mshBill.TextMatrix(Row, mconIntCol申领数量)
            Me.mshBill.TextMatrix(Row, mconIntCol实际数量) = Me.mshBill.TextMatrix(Row, mconIntCol申领数量)
            If Val(Me.mshBill.TextMatrix(Row, mconIntCol填写数量)) <> 0 Then
                Call GetPrice(Row)
            Else
                With Me.mshBill
                    .TextMatrix(Row, mconIntCol售价金额) = 0
                    .TextMatrix(Row, mconintCol差价) = 0
                    .TextMatrix(Row, mconIntCol采购价) = 0
                    .TextMatrix(Row, mconIntCol采购金额) = 0
                End With
            End If
        Next
        Call 显示合计金额
    End If
End Sub

Private Sub cmd全清_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("您确定要将填写数量和实际数量清为0吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol填写数量) = 0
            Me.mshBill.TextMatrix(Row, mconIntCol实际数量) = 0
            With Me.mshBill
                .TextMatrix(Row, mconIntCol售价金额) = 0
                .TextMatrix(Row, mconintCol差价) = 0
                .TextMatrix(Row, mconIntCol采购价) = 0
                .TextMatrix(Row, mconIntCol采购金额) = 0
            End With
        Next
        Call 显示合计金额
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then
        If mshBill.rows > 50 Then
            Call AviShow(Me) '提示用户正在查询数据
        End If
        Call get库存数量    '为当前库存数量和对方库存数量列赋值
        If mshBill.rows > 50 Then
            Call AviShow(Me, False)
        End If
        Exit Sub
    End If
    
    mblnFirst = False
    If mint编辑状态 = 5 Then
        If Not frmRequestNavigation.ShowNavigation(Me, mlngStockID, mintApplyType, mstrEndTime, mbln申领状态) = True Then
            Unload Me
            Exit Sub
        End If
        mshBill.SetFocus
        If mintApplyType = 7 And Not IsHavePrivs(mstrPrivs, "自动申领时修改药品内容") Then
            mshBill.Active = False
        End If
    End If
    If mbln申领状态 = True Then
        Call Form_Resize
    End If
    
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            If mint编辑状态 = 7 Then
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
    Dim Row As Integer
    Dim count As Integer
    Dim intRows As Integer
    
    '设置排序数据集
    Call SetSortRecord
        
    If mint编辑状态 = 4 Then    '查看
        '打印
        printbill
        '退出
        Unload Me
        Exit Sub
    End If
    
    If Me.mshBill.TextMatrix(Me.mshBill.rows - 1, 0) <> "" Then
        intRows = Me.mshBill.rows - 1
    Else
        intRows = Me.mshBill.rows - 2
    End If
    
    For Row = 1 To intRows
        If Val(Me.mshBill.TextMatrix(Row, mconIntCol填写数量)) = 0 Then
            count = count + 1
            If count = intRows Then
                MsgBox "该申领单上的所有药品的填写数量均为0，不能继续操作！", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
    Next

    For Row = 1 To Me.mshBill.rows - 2
        If NVL(Me.mshBill.TextMatrix(Row, mconIntCol填写数量), 0) = 0 Then
            If MsgBox("该申领单中有填写数量为0的药品，" & vbCrLf & "填写数量为0的药品将不能保存为申领单，" & vbCrLf & "是否继续操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    If mint编辑状态 = 6 Then       '审核
        If Not 药品单据审核(Txt填制人.Caption) Then Exit Sub
        
        '如果参数下可用数量为假，则要检查可用库存是否足够
        If mbln下可用数量 = False Then
            If Not CheckStock Then Exit Sub
        End If
        
        If SaveCheck() = True Then
            If Val(zlDataBase.GetPara("审核打印", glngSys, 1343)) = 1 Then
                '打印
                If IsHavePrivs(mstrPrivs, "单据打印") Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 7 Then '冲销
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
    
    If mint编辑状态 = 2 And mblnUpdate = False Then
        If Not 检查单价(6, txtNo.Tag, True, True) Then
            MsgBox "有记录未使用最新价格，程序将自动完成更新（售价、成本价、售价金额、成本金额、差价），更新后请检查！", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            Exit Sub
        End If
    End If
        
    If ValidData = False Then Exit Sub
    
    If mint编辑状态 = 1 Or mint编辑状态 = 5 Then '新增保存时，判断价格是否已经更新
        If 检查价格 Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If

    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("存盘打印", glngSys, 1343)) = 1 Then
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
    cboStock.SetFocus
    mblnChange = False

    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsStock As New Recordset
    Dim intStock As Integer
    
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    mblnViewCost = IsHavePrivs(mstrPrivs, "查看成本价")
    chkExportPlan.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "同步生成计划单", 0))
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品申领管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    intStock = -1
    With cboStock
        .Clear
        mstr库房 = ""
        Do While Not rsDepend.EOF
            If InStr(1, mstr库房, "|" & rsDepend!Id & "|") = 0 Then
                .AddItem rsDepend!名称
                .ItemData(.NewIndex) = rsDepend!Id
                mstr库房 = mstr库房 & "|" & rsDepend!Id & "|"
                
                If rsDepend!药库性质 = 1 And intStock = -1 Then
                    intStock = .NewIndex
                End If
            End If
            
            rsDepend.MoveNext
        Loop
        .ListIndex = IIf(intStock = -1, 0, intStock)
    End With
    
    If mlngStockID = 0 Then
        mlngStockID = mfrmMain.cboStock.ItemData(Me.cboStock.ListIndex)
    End If
    Call GetDrugDigit(mlngStockID, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(6, mstr单据号)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    
    '根据系统参数决定药房人员查看单据时，是否显示成本价
    mshBill.ColWidth(mconIntCol采购价) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol采购金额) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol差价) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconIntCol上次供应商ID) = 0
    mshBill.ColWidth(mconintCol真实数量) = 0
    mshBill.ColWidth(mconIntCol申领数量) = 0
    
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
        mshBill.ColWidth(mconintcol当前库存) = 1100
        mshBill.ColWidth(mconintcol对方库存) = 1100
    Else
        mshBill.ColWidth(mconintcol当前库存) = 0
        mshBill.ColWidth(mconintcol对方库存) = 0
    End If
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
    mbln检查库存 = (Val(zlDataBase.GetPara("显示无库存药品", glngSys, 1343, 0)) = 0)
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim lngStockid As Long
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim strUnitQuantity_Stock As String
    Dim intRow As Integer
    Dim vardrug As Variant
    Dim numUseAbleCount As Double
    Dim dateCurDate As Date
    Dim strOrder As String, strCompare As String
    Dim IntCount As Integer
    Dim intCostDigit As Integer        '成本价小数位数
    Dim intPricedigit As Integer       '售价小数位数
    Dim intNumberDigit As Integer      '数量小数位数
    Dim intMoneyDigit As Integer       '金额小数位数
    Dim str药名 As String
    Dim strSqlOrder As String
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 1343)
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
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
    If mint编辑状态 = 4 Then
        With cboStock
            '取指定单据的出库库房与入库库房
            gstrSQL = " Select 库房ID,对方部门ID From 药品收发记录" & _
                      " Where NO=[1] And 单据=6 And 入出系数=-1 And Rownum<2"
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取指定单据的出库库房与入库库房]", mstr单据号)
            
            If rsInitCard.RecordCount <> 0 Then
                lngStockid = rsInitCard!库房id
            End If
            
            For IntCount = 0 To .ListCount - 1
                If .ItemData(IntCount) = lngStockid Then
                    .ListIndex = IntCount: Exit For
                End If
            Next
        End With
    Else
        With cboStock
            If Not (mint编辑状态 = 1 Or mint编辑状态 = 5) Then
                '取指定单据的出库库房与入库库房
                gstrSQL = " Select 库房ID,对方部门ID From 药品收发记录" & _
                          " Where NO=[1] And 单据=6 And 入出系数=-1 And Rownum<2"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取指定单据的出库库房与入库库房]", mstr单据号)
                
                If rsInitCard.RecordCount <> 0 Then
                    lngStockid = rsInitCard!库房id
                End If
            End If
            For IntCount = 0 To .ListCount - 1
                If .ItemData(IntCount) = lngStockid Then
                    .ListIndex = IntCount: Exit For
                End If
            Next
            mintcboIndex = .ListIndex
        End With
    End If
    
    If mint编辑状态 = 7 Then
       lngStockid = mlngStockID
    End If
    
    dateCurDate = zlDataBase.Currentdate()
    
    Select Case mint编辑状态
        Case 1, 5
            Txt填制人 = gstrUserName
            Txt填制日期 = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 4, 6, 7
            initGrid
            
            Select Case mintUnit
                Case mconint售价单位
                    strUnitQuantity = "B.送货单位,B.送货包装,D.计算单位 AS 单位, A.填写数量,a.实际数量,a.成本价,a.零售价,'1' as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量,Z.实际金额,Z.实际差价"
                Case mconint门诊单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.门诊单位 AS 单位,(A.填写数量 / B.门诊包装) AS 填写数量,(A.实际数量 / B.门诊包装) AS 实际数量,a.成本价*B.门诊包装 as 成本价,a.零售价*B.门诊包装 as 零售价,B.门诊包装 as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量/B.门诊包装 As 可用数量,Z.实际金额,Z.实际差价"
                Case mconint住院单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.住院单位 AS 单位,(A.填写数量 / B.住院包装) AS 填写数量,(A.实际数量 / B.住院包装) AS 实际数量,a.成本价*B.住院包装 as 成本价,a.零售价*B.住院包装 as 零售价,B.住院包装 as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量/B.住院包装 As 可用数量,Z.实际金额,Z.实际差价"
                Case mconint药库单位
                    strUnitQuantity = "B.送货单位,B.送货包装,B.药库单位 AS 单位,(A.填写数量 / B.药库包装) AS 填写数量,(A.实际数量 / B.药库包装) AS 实际数量,a.成本价*B.药库包装 as 成本价,a.零售价*B.药库包装 as 零售价,B.药库包装 as 比例系数,"
                    strUnitQuantity_Stock = "Z.可用数量/B.药库包装 As 可用数量,Z.实际金额,Z.实际差价"
            End Select
            
            If mint编辑状态 = 7 Then
                gstrSQL = "SELECT W.*,Z.可用数量/W.比例系数 AS  可用数量,Z.实际金额,Z.实际差价 " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                    "     B.药品来源,B.基本药物,D.规格,D.产地 AS 原产地,A.产地, A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算," & _
                    "     B.最大效期,A.效期," & strUnitQuantity & _
                    "     A.成本金额,0 零售金额, 0 差价,D.摘要,A.库房ID,A.对方部门ID,D.是否变价,B.药房分批 AS 药房分批核算,A.上次供应商ID,A.批准文号,A.填写数量 真实数量 " & _
                    "     FROM " & _
                    "         (SELECT MIN(ID) AS ID, SUM(实际数量) AS 填写数量,0 实际数量,SUM(成本金额) AS 成本金额," & _
                    "          药品ID,序号,产地, 批号,效期,NVL(批次,0) 批次,扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0) 上次供应商ID,批准文号 " & _
                    "          FROM 药品收发记录 X " & _
                    "          WHERE NO=[1] AND 单据=6 AND 入出系数=-1 " & _
                    "          GROUP BY 药品ID,序号,产地,批号,效期,NVL(批次,0),扣率,成本价,零售价,库房ID,对方部门ID,入出类别ID,NVL(供药单位ID,0),批准文号" & _
                    "          HAVING SUM(实际数量)<>0 ) A," & _
                    "     药品规格 B,收费项目目录 D,收费项目别名 E, " & _
                    " (Select 序号, 摘要 From 药品收发记录 " & _
                    "  Where 单据 = 6 And NO = [1] And 入出系数 = -1 And (记录状态 = 1 Or Mod(记录状态, 3) = 0)) D " & _
                    "     WHERE A.药品ID = B.药品ID AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 AND B.药品ID=D.ID And A.序号 = D.序号) W," & _
                    "     (SELECT  药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                    " WHERE W.药品ID=Z.药品ID(+) AND NVL(W.批次,0)=Z.批次(+) " & _
                     " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT DISTINCT A.药品ID,A.序号,'[' || D.编码 || ']' As 药品编码, D.名称 As 通用名, E.名称 As 商品名," & _
                    " B.药品来源,B.基本药物,D.规格,D.产地 AS 原产地,A.产地,A.批号,A.批次,B.指导差价率,B.药库分批 AS 分批核算,A.填写数量 as 原始数量, " & _
                    " B.最大效期,A.效期," & strUnitQuantity & _
                    " A.成本金额,A.零售金额, A.差价, " & strUnitQuantity_Stock & _
                    " ,A.摘要,填制人,填制日期,审核人,审核日期,A.库房ID,A.对方部门ID,D.是否变价,B.药房分批 AS 药房分批核算,NVL(A.供药单位ID,0) 上次供应商ID,A.批准文号,nvl(A.单量,0) As 申领方式  " & _
                    " FROM 药品收发记录 A, 药品规格 B,收费项目别名 E ,收费项目目录 D, " & _
                    "     (SELECT 药品ID,NVL(批次,0) 批次,可用数量,实际金额,实际差价 " & _
                    "     FROM 药品库存 WHERE 库房ID=[2] AND 性质=1) Z " & _
                    " WHERE A.药品ID = B.药品ID AND b.药品ID=D.ID " & _
                    " AND B.药品ID=E.收费细目ID(+) AND E.性质(+)=3 " & _
                    " AND A.单据 = 6 AND A.入出系数=-1 AND A.NO = [1] AND A.记录状态 =[3] " & _
                    " AND A.药品ID=Z.药品ID(+) AND NVL(A.批次,0)=Z.批次(+) " & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号, lngStockid, mint记录状态)
        
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            If mint编辑状态 = 2 Or mint编辑状态 = 4 Or mint编辑状态 = 6 Then
                mintApplyType = rsInitCard!申领方式
            End If
            mshBill.Active = IIf(mintApplyType = 0, True, IsHavePrivs(mstrPrivs, "自动申领时修改药品内容"))
            
            If mint编辑状态 = 7 Then
                Txt填制人 = gstrUserName
                Txt填制日期 = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = gstrUserName
                Txt审核日期 = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            Else
                Txt填制人 = rsInitCard!填制人
                If mint编辑状态 = 2 Then
                    Txt填制人 = gstrUserName
                End If
                Txt填制日期 = Format(rsInitCard!填制日期, "yyyy-mm-dd hh:mm:ss")
                Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
                Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            End If
            txt摘要.Text = IIf(IsNull(rsInitCard!摘要), "", rsInitCard!摘要)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint编辑状态 = 2 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    'IntRow = rsInitCard!序号
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
                    .TextMatrix(intRow, mconIntCol来源) = NVL(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol基本药物) = NVL(rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol序号) = rsInitCard!序号
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconIntCol产地) = IIf(IsNull(rsInitCard!产地), "", rsInitCard!产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    If IIf(IsNull(rsInitCard!比例系数), 0, rsInitCard!比例系数) = 0 Or NVL(rsInitCard!送货包装) = "" Or NVL(rsInitCard!送货单位) = "" Then
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
                    
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                                
                    .TextMatrix(intRow, mconIntCol填写数量) = GetFormat(rsInitCard!填写数量, intNumberDigit)
                    .TextMatrix(intRow, mconIntCol实际数量) = GetFormat(rsInitCard!实际数量, intNumberDigit)
                    
                    If mint编辑状态 = 6 Then
                        .TextMatrix(intRow, mconIntCol原始数量) = GetFormat(rsInitCard!原始数量, intNumberDigit)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol采购价) = GetFormat(rsInitCard!成本价, intCostDigit)
                    
                    .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(IIf(mint编辑状态 = 7, 0, rsInitCard!成本金额), intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol售价) = GetFormat(rsInitCard!零售价, intPricedigit)
                    .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(rsInitCard!零售金额, intMoneyDigit)
                    .TextMatrix(intRow, mconintCol差价) = GetFormat(rsInitCard!差价, intMoneyDigit)
                    
                    .TextMatrix(intRow, mconIntCol最大效期) = IIf(IsNull(rsInitCard!最大效期), "0", rsInitCard!最大效期) & "||" & rsInitCard!是否变价 & "||" & rsInitCard!药房分批核算
                    .TextMatrix(intRow, mconIntCol批次) = IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol指导差价率) = rsInitCard!指导差价率
                    .TextMatrix(intRow, mconIntCol分批核算) = IIf(IsNull(rsInitCard!分批核算), "0", rsInitCard!分批核算)
                    .TextMatrix(intRow, mconIntCol可用数量) = IIf(IsNull(rsInitCard!可用数量), "0", rsInitCard!可用数量)
                    .TextMatrix(intRow, mconIntCol实际差价) = IIf(IsNull(rsInitCard!实际差价), "0", rsInitCard!实际差价)
                    .TextMatrix(intRow, mconIntCol实际金额) = IIf(IsNull(rsInitCard!实际金额), "0", rsInitCard!实际金额)
                    .TextMatrix(intRow, mconIntCol上次供应商ID) = rsInitCard!上次供应商ID
                                        
                    If mint编辑状态 = 7 Then
                        .TextMatrix(intRow, mconintCol真实数量) = rsInitCard!真实数量
                    End If
                        
                    
                    If mint编辑状态 = 2 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!药品id & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!药品id & IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!填写数量), "0", rsInitCard!填写数量))), CStr(rsInitCard!药品id) & CStr(IIf(IsNull(rsInitCard!批次), "0", rsInitCard!批次))
                        
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    
    Call get库存数量
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
        .TextMatrix(0, mconIntCol送货单位) = "送货单位"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconIntCol效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        .TextMatrix(0, mconintcol当前库存) = "当前库存"
        .TextMatrix(0, mconintcol对方库存) = "对方库存"
        .TextMatrix(0, mconIntCol申领数量) = "申领数量"
        .TextMatrix(0, mconIntCol填写数量) = IIf(mint编辑状态 = 7, "数量", "填写数量")
        .TextMatrix(0, mconIntCol实际数量) = IIf(mint编辑状态 = 7, "冲销数量", "实际数量")
        
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
        .TextMatrix(0, mconIntCol常备药品) = "常备药品"
        .TextMatrix(0, mconIntCol原始数量) = "原始数量"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol行号) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol序号) = 0
        .ColWidth(mconIntCol行号) = 300
        .ColWidth(mconIntCol药名) = 2200
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 900
        .ColWidth(mconIntCol基本药物) = 900
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol产地) = 800
        .ColWidth(mconIntCol单位) = 400
        .ColWidth(mconIntCol送货单位) = 2000
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconIntCol效期) = 1000
        .ColWidth(mconIntCol批准文号) = 1000
        .ColWidth(mconIntCol申领数量) = 0
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
        .ColWidth(mconIntCol常备药品) = 0
        .ColWidth(mconIntCol原始数量) = 0
        
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(0) = 5
        .ColData(mconIntCol申领数量) = 0
        .ColData(mconIntCol商品名) = 5
        .ColData(mconIntCol行号) = 5
        .ColData(mconIntCol来源) = 5
        .ColData(mconIntCol基本药物) = 5
        .ColData(mconIntCol序号) = 5
        .ColData(mconIntCol规格) = 5
        .ColData(mconIntCol产地) = 5
        .ColData(mconIntCol单位) = 5
        .ColData(mconIntCol送货单位) = 5
        .ColData(mconIntCol批号) = 5
        .ColData(mconIntCol效期) = 5
        .ColData(mconIntCol批准文号) = 5
        .ColData(mconIntCol药品编码和名称) = 5
        .ColData(mconIntCol药品编码) = 5
        .ColData(mconIntCol药品名称) = 5
        .ColData(mconIntCol常备药品) = 5
        .ColData(mconIntCol原始数量) = 5
        
        '列状态为不能编辑
        .ColData(mconintcol当前库存) = 5
        .ColData(mconintcol对方库存) = 5
        
        '仅当单据为编辑状态（新增或修改）时可见
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
            .ColWidth(mconintcol当前库存) = 1100
            '当没有显示对方库存权限的时候不能显示对方库存
            If IsHavePrivs(mstrPrivs, "显示对方库存") Then
                .ColWidth(mconintcol对方库存) = 1100
            Else
                .ColWidth(mconintcol对方库存) = 0
            End If
        Else
            .ColWidth(mconintcol当前库存) = 0
            .ColWidth(mconintcol对方库存) = 0
        End If
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
            
            cboStock.Enabled = True
            txt摘要.Enabled = True
            
            .ColData(mconIntCol药名) = 1
            .ColData(mconIntCol填写数量) = 4
            .ColData(mconIntCol实际数量) = 5
        ElseIf mint编辑状态 = 4 Or mint编辑状态 = 6 Then
            cboStock.Enabled = False
            txt摘要.Enabled = False
            
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = IIf(mint编辑状态 <> 6, 4, 5)
            .ColData(mconIntCol药名) = 0
        ElseIf mint编辑状态 = 7 Then
            cboStock.Enabled = False
            txt摘要.Enabled = True
            
            .ColData(mconIntCol填写数量) = 5
            .ColData(mconIntCol实际数量) = 4
            .ColData(mconIntCol药名) = 0
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
        .ColData(mconintCol真实数量) = 5
        .ColData(mconIntCol上次供应商ID) = 5
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol送货单位) = flexAlignCenterCenter
        .ColAlignment(mconIntCol批号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol效期) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        .ColAlignment(mconintcol当前库存) = flexAlignRightCenter
        .ColAlignment(mconintcol对方库存) = flexAlignRightCenter
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
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
    End With
    txt摘要.MaxLength = GetLength("药品收发记录", "摘要")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - IIf(Me.cmd全部复制.Visible, 350, 0) - .Top - 100 - CmdCancel.Height - 200
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
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
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
    

    
    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
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
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    With chkExportPlan
        .Top = lblCode.Top
    End With
    
    With cmd全清
        If .Visible = True Then
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Top
        End If
    End With
    
    With cmd全部复制
        If .Visible = True Then
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Top
        End If
    End With
    
    If mint编辑状态 = 5 And Me.cmd全清.Visible = True Then
        With Me.CmdSave
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Height + Me.CmdSave.Top + 100
        End With
    
        With Me.CmdCancel
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Height + Me.CmdCancel.Top + 100
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintApplyType = 0
    mstrEndTime = ""
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品申领管理", "同步生成计划单", Me.chkExportPlan.Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品申领管理", "药品名称显示方式", mintDrugNameShow)
    
    If mblnChange = False Or mint编辑状态 = 4 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        mblnUpdate = False
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
    mblnUpdate = False
    zlPlugIn_Unload mobjPlugIn
End Sub

Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
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
Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol行号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call 显示合计金额
    Call RefreshRowNO(mshBill, mconIntCol行号, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "3467", mint编辑状态) <> 0 Then
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
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
'    Set RecReturn = Frm药品选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
'        mlngStockID, mlngStockID, mbln检查库存, IIf(mint明确批次 = 0, False, True), _
'        False, False, IsHavePrivs(mstrPrivs, "显示对方库存"))
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "药品申领管理", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln检查库存)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln检查库存, IIf(mint明确批次 = 0, False, True), IsHavePrivs(mstrPrivs, "显示对方库存"), False, , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
    End If
    mshBill.CmdEnable = True
    
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        For i = 1 To RecReturn.RecordCount
            intCurRow = mshBill.Row
            With mshBill
                .TextMatrix(intCurRow, mconIntCol行号) = .Row
                SetColValue .Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                    NVL(RecReturn!药品来源), NVL(RecReturn!基本药物), _
                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                    IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
                    IIf(IsNull(RecReturn!最大效期), "0", RecReturn!最大效期), _
                    RecReturn!药库分批, _
                    IIf(IsNull(RecReturn!可用数量), "0", RecReturn!可用数量), _
                    IIf(IsNull(RecReturn!实际金额), "0", RecReturn!实际金额), _
                    IIf(IsNull(RecReturn!实际差价), "0", RecReturn!实际差价), _
                    IIf(IsNull(RecReturn!指导差价率), "0", RecReturn!指导差价率), _
                    Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装), _
                    IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), RecReturn!时价, RecReturn!药房分批, _
                    RecReturn!上次供应商ID, IIf(IsNull(RecReturn!批准文号), "", RecReturn!批准文号)
                .Col = mconIntCol填写数量
'                .TextMatrix(.Row, mconIntCol常备药品) = True
                
                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                    .rows = .rows + 1
                End If
                .Row = .rows - 1
                RecReturn.MoveNext
            End With
        Next
        mshBill.Row = intOldRow
        RecReturn.Close
    End If
End Sub

Private Sub mshBill_DblClick(Cancel As Boolean)
    If Me.mshBill.Row <> Me.mshBill.rows - 1 Then
        If Me.mshBill.Col = mconIntCol申领数量 And Me.mshBill.Row <> 0 Then
            If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol填写数量)) = 0 Then
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol填写数量) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol申领数量)
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol实际数量) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol申领数量)
            Else
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol填写数量) = 0
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol实际数量) = 0
            End If
        End If
        
        If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol填写数量)) <> 0 Then
            Call GetPrice(Me.mshBill.Row)
        Else
             With Me.mshBill
                .TextMatrix(Me.mshBill.Row, mconIntCol售价金额) = 0
                .TextMatrix(Me.mshBill.Row, mconintCol差价) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol采购价) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol采购金额) = 0
            End With
        End If
        
        Call 显示合计金额
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    With mshBill
        If .Col <> mconIntCol产地 Then
            mshBill.Text = UCase(curText)
            mshBill.SelStart = Len(mshBill.Text)
        End If
    End With
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol填写数量 Or .Col = mconIntCol实际数量 Then
            strkey = .Text
            If strkey = "" Then
                strkey = .TextMatrix(.Row, .Col)
            End If
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
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntCol药名
                .TxtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
                Call 提示库存数
                
            Case mconIntCol批号
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 8
            
            Case mconIntCol效期
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol批号) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol批号)) And .TextMatrix(.Row, mconIntCol最大效期) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol最大效期), "||")(0) <> 0 Then
                            strxq = .TextMatrix(.Row, mconIntCol批号)
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
            Case mconIntCol填写数量, mconIntCol实际数量
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                Call 提示库存数
                
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
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
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
'                    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
'                        mlngStockID, mlngStockID, strkey, sngLeft, sngTop, mbln检查库存, _
'                        IIf(mint明确批次 = 0, False, True), False, False, IsHavePrivs(mstrPrivs, "显示对方库存"))

                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "药品申领管理", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln检查库存)
                    End If
                    Set RecReturn = frmSelector.ShowMe(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln检查库存, IIf(mint明确批次 = 0, False, True), IsHavePrivs(mstrPrivs, "显示对方库存"), False, , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '检查重复记录 并将重复记录的药品id返回回来
                    End If
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol行号) = .Row
                            If SetColValue(.Row, RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), _
                                    NVL(RecReturn!药品来源), NVL(RecReturn!基本药物), _
                                    IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                    Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位), _
                                    IIf(IsNull(RecReturn!售价), 0, RecReturn!售价), IIf(IsNull(RecReturn!批号), "", RecReturn!批号), _
                                    IIf(IsNull(RecReturn!效期), "", RecReturn!效期), _
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
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                        .Col = mconIntCol填写数量
                    Else
                        .TextMatrix(.Row, mconIntCol常备药品) = True
                        Cancel = True
                    End If
                    Call 提示库存数
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
                End If
                
                If Len(strkey) < 8 Then
                    MsgBox "对不起，批号长度不够，必须为8位,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mconIntCol效期
                '有处理
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            MsgBox "对不起，效期必须为日期型！", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strkey
                        Exit Sub
                    End If
                    If Not IsDate(strkey) Then
                        MsgBox "对不起，效期必须为日期型如(2000-10-10) 或（20001010）,请重输！", vbInformation + vbOKOnly, gstrSysName
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
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
'                    If Val(strkey) = 0 Then
'                        MsgBox "对不起，数量不能为零,请重输！", vbInformation + vbOKOnly, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
'
'                    If Val(strkey) < 0.00001 Then
'                        MsgBox "对不起，数量必须为正数并且不能小于0.00001,请重输！", vbInformation + vbOKOnly, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "数量必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Not CompareUsableQuantity(.Row, strkey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '成本价的公式：     出库金额=数量*售价
                    '                  出库差价=出库金额*（实际差价/实际金额）
                    '                  if 实际金额=0 then  出库差价=出库金额*指导差价率
                    '                  购价（成本价）=直接从库存表中取平均成本价
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '实际金额=0的情况下，依次考虑从“库存记录的上次采购价”、“药品规格的成本价”、“指导差价率”取值
                    
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = GetFormat(.TextMatrix(.Row, mconIntCol售价) * strkey, mintMoneyDigit)
                    End If
                    
                    If strkey <> 0 Then
'                        .TextMatrix(.Row, mconIntCol采购价) = FormatEx((Val(.TextMatrix(.Row, mconIntCol售价金额)) - .TextMatrix(.Row, mconintCol差价)) / strkey, mintCostDigit)
                        If mint编辑状态 <> 7 Then .TextMatrix(.Row, mconIntCol采购价) = GetFormat(Get成本价(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol批次))) * Val(.TextMatrix(.Row, mconIntCol比例系数)), mintCostDigit)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol采购金额) = GetFormat(Val(.TextMatrix(.Row, mconIntCol采购价)) * strkey, mintMoneyDigit)
                    
'                    If mint编辑状态 = 7 Then
                        .TextMatrix(.Row, mconintCol差价) = GetFormat(Val(.TextMatrix(.Row, mconIntCol售价金额)) - Val(.TextMatrix(.Row, mconIntCol采购金额)), mintMoneyDigit)
'                    Else
'                        .TextMatrix(.Row, mconintCol差价) = GetFormat(Get出库差价(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)), Val(.TextMatrix(.Row, mconIntCol实际金额)), Val(.TextMatrix(.Row, mconIntCol实际差价)), Val(.TextMatrix(.Row, mconIntCol售价金额)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol比例系数))), mintMoneyDigit)
'                    End If
                    
                    If .Col = mconIntCol填写数量 Then
                        .TextMatrix(.Row, mconIntCol实际数量) = strkey
                    End If
                    
                    
                End If
                
                显示合计金额
            
        End Select
    End With
End Sub

Private Sub GetPrice(ByVal intRow As Integer)
    With Me.mshBill
        .TextMatrix(intRow, mconIntCol售价金额) = GetFormat(.TextMatrix(intRow, mconIntCol售价) * Me.mshBill.TextMatrix(intRow, mconIntCol填写数量), mintMoneyDigit)
        .TextMatrix(intRow, mconIntCol采购价) = GetFormat(Get成本价(Val(.TextMatrix(intRow, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol批次))) * Val(.TextMatrix(intRow, mconIntCol比例系数)), mintCostDigit)
        .TextMatrix(intRow, mconIntCol采购金额) = GetFormat(Val(.TextMatrix(intRow, mconIntCol采购价)) * Val(Me.mshBill.TextMatrix(intRow, mconIntCol填写数量)), mintMoneyDigit)
        .TextMatrix(intRow, mconintCol差价) = GetFormat(Val(.TextMatrix(intRow, mconIntCol售价金额)) - .TextMatrix(intRow, mconIntCol采购金额), mintMoneyDigit)
    End With
End Sub

'从药品目录中取值并附给相应的列
Private Function SetColValue(ByVal intRow As Integer, ByVal lng药品ID As Long, _
    ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, _
    ByVal str基本药物 As String, ByVal str规格 As String, ByVal str产地 As String, _
    ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal int分批核算 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal int是否变价 As Integer, ByVal int药房分批 As Integer, ByVal lng上次供应商ID As Long, ByVal str批准文号 As String) As Boolean
    
    Dim IntCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    Dim str药名 As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    SetColValue = False
    
    '检查是否重复
'    If Not CheckRepeatMedicine(mshBill, lng药品ID & "," & "0" & "|" & IIf(mint明确批次 = 1, lng批次, 0) & "," & mconIntCol批次, intRow) Then
'        Exit Function
'    End If
    
    With mshBill
'        If int是否变价 = 1 Then
'            gstrSQL = "select Decode(Nvl(批次, 0), 0, 实际金额 / 实际数量, Nvl(零售价, 实际金额 / 实际数量))*" & num比例系数 & " as  售价 " _
'                & "  from 药品库存 " _
'                & " where 库房id=[1] " _
'                & " and 药品id=[2] " _
'                & " and 性质=1 and 实际数量>0 and " _
'                & " nvl(批次,0)=[3]"
'            Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lng药品ID, lng批次)
'
'            If rsPrice.EOF Then
'                If mint明确批次 = 1 Then
'                    MsgBox "时价药品没有库存，不能出库，请检查！", vbOKOnly, gstrSysName
'                    Exit Function
'                End If
'            Else
'                dblPrice = rsPrice.Fields(0).Value
'            End If
'        End If
        
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
        If num比例系数 = 0 Or NVL(rsTemp!送货包装) = "" Or NVL(rsTemp!送货单位) = "" Then
            .TextMatrix(intRow, mconIntCol送货单位) = ""
        Else
            .TextMatrix(intRow, mconIntCol送货单位) = rsTemp!送货单位 & "(1" & rsTemp!送货单位 & "=" & zlStr.FormatEx(rsTemp!送货包装 / num比例系数, 1, , True) & str单位 & ")"
        End If
        
        .TextMatrix(intRow, mconIntCol售价) = GetFormat(num售价 * num比例系数, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol分批核算) = int分批核算
        .TextMatrix(intRow, mconIntCol可用数量) = GetFormat(num可用数量, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol最大效期) = int最大效期 & "||" & int是否变价 & "||" & int药房分批
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntCol指导差价率) = num指导差价率
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol上次供应商ID) = lng上次供应商ID

        If mint明确批次 = 1 Then
            .TextMatrix(intRow, mconIntCol批次) = lng批次
            .TextMatrix(intRow, mconIntCol批号) = str批号
            .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
            .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        Else
            .TextMatrix(intRow, mconIntCol批次) = 0
            .TextMatrix(intRow, mconIntCol批号) = ""
            .TextMatrix(intRow, mconIntCol效期) = ""
            .TextMatrix(intRow, mconIntCol批准文号) = ""
        End If
        If int是否变价 = 1 Then
            dblPrice = Get售价(True, lng药品ID, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol批次))) 'Get价格(lng药品id, Val(.TextMatrix(intRow, mconIntCol批次)), num比例系数)
            .TextMatrix(intRow, mconIntCol售价) = GetFormat(dblPrice * num比例系数, mintPriceDigit)
        End If
        Call CheckLapse(str效期)
        
        '是否常备药品
        Dim rsTmp As ADODB.Recordset
        gstrSQL = "select nvl(是否常备,0) 是否常备 from 药品规格 where 药品id=[1]"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品ID)
        .TextMatrix(intRow, mconIntCol常备药品) = IIf(rsTmp!是否常备 = 1, False, True)
    End With
    
    Call get库存数量
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据
            
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
                                       
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngEnterStockID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblRealQuantity As Double
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
    Dim lng上次供应商ID As Long
    Dim str批准文号 As String
    Dim int序号 As Integer
    
    Dim intRow As Integer
    Dim arrSql As Variant
    '药品采购计划
    Dim strSQLDrugPlan As String
    Dim arrSQLDrugPlanDetail As Variant
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim arrSum As Variant
    
    '自动分解申领记录时使用
    Dim blnAuto As Boolean              '是否需要自动分解
    Dim rsStock As New ADODB.Recordset
    
    Dim strCheckString As String
    Dim n As Integer, intPlanSN As Integer
    Dim rsSpec As ADODB.Recordset   '规格数据集
    Dim dbl送货数量 As Double
    
    SaveCard = False
    arrSql = Array()
    arrSQLDrugPlanDetail = Array()
    arrSum = Array()
    
    On Error GoTo errHandle
    
    '当下可用库存数量参数为真时，保存时要检查库存
    If mbln下可用数量 Then
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
                If Not CompareUsableQuantity(n, mshBill.TextMatrix(n, mconIntCol实际数量)) Then
                    Exit Function
                End If
            End If
        Next
    End If
    
    With mshBill
        chrNo = Trim(txtNo)
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = zlDataBase.GetNextNo(26, lngStockid)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        lngEnterStockID = mlngStockID
        strBrief = Trim(txt摘要.Text)
        strBooker = Txt填制人
        datBookDate = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Txt审核人
        
        ID_IN = zlDataBase.GetNextId("药品采购计划")
        NO_IN = zlDataBase.GetNextNo(32, mlngStockID)
        
        If mint编辑状态 = 2 Then        '修改
            strCheckString = CheckBill(chrNo)
            If strCheckString <> "" Then
                MsgBox strCheckString, vbInformation, gstrSysName
                Exit Function
            End If
        
            gstrSQL = "zl_药品移库_Delete('" & mstr单据号 & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
        End If
        
        '按药品ID顺序更新数据
        recSort.Sort = "药品id,批次,序号"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!行号
            If NVL(.TextMatrix(intRow, mconIntCol填写数量), 0) <> 0 Then
                int序号 = intRow 'int序号 + 1
                If .TextMatrix(intRow, 0) <> "" Then
                    '如果当前批次药品不够，自动取其它批次的药品，产生多笔申领记录
                    lngDrugID = .TextMatrix(intRow, 0)
                    strProducingArea = .TextMatrix(intRow, mconIntCol产地)
                    strBatchNo = .TextMatrix(intRow, mconIntCol批号)
                    lngBatchID = .TextMatrix(intRow, mconIntCol批次)
                    datTimeLimit = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And datTimeLimit <> "" Then
                        '换算为失效期来保存
                        datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                    End If
                    
                    dblQuantity = .TextMatrix(intRow, mconIntCol填写数量) * .TextMatrix(intRow, mconIntCol比例系数)
                    dblRealQuantity = .TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数)
'                    dblPurchasePrice = FormatEx(.TextMatrix(intRow, mconIntCol采购价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                    dblPurchasePrice = Get成本价(lngDrugID, lngStockid, lngBatchID)
                                        
                    dblPurchaseMoney = .TextMatrix(intRow, mconIntCol采购金额)
'                    dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol售价) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                    dblSalePrice = Get零售价(Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                    
                    dblSaleMoney = .TextMatrix(intRow, mconIntCol售价金额)
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol差价))
                    lng上次供应商ID = .TextMatrix(intRow, mconIntCol上次供应商ID)
                    str批准文号 = IIf(Trim(.TextMatrix(intRow, mconIntCol批准文号)) = "", "", .TextMatrix(intRow, mconIntCol批准文号))
                    
'                    If Val(.TextMatrix(intRow, mconIntCol序号)) = 0 Then
'                        lngSerial = 2 * int序号 - 1  '求奇数：公式为：2n-1;出库序号为偶数
'                    Else
'                        lngSerial = Val(.TextMatrix(intRow, mconIntCol序号))
'                    End If
                    lngSerial = 2 * int序号 - 1
                    
                    gstrSQL = "zl_药品申领_INSERT("
                    'NO
                    gstrSQL = gstrSQL & "'" & chrNo & "'"
                    '序号
                    gstrSQL = gstrSQL & "," & lngSerial
                    '库房ID
                    gstrSQL = gstrSQL & "," & lngStockid
                    '对方部门ID
                    gstrSQL = gstrSQL & "," & lngEnterStockID
                    '药品ID
                    gstrSQL = gstrSQL & "," & lngDrugID
                    '批次
                    gstrSQL = gstrSQL & "," & lngBatchID
                    '填写数量
                    gstrSQL = gstrSQL & "," & dblQuantity
                    '实际数量
                    gstrSQL = gstrSQL & "," & dblRealQuantity
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
                    gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & datTimeLimit & "','yyyy-mm-dd')")
                    '摘要
                    gstrSQL = gstrSQL & ",'" & strBrief & "'"
                    '填制日期
                    gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                    '供应商ID
                    gstrSQL = gstrSQL & "," & IIf(lng上次供应商ID = 0, "NULL", lng上次供应商ID)
                    '批准文号
                    gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                    '申领方式
                    gstrSQL = gstrSQL & "," & mintApplyType
                    '结束时间
                    gstrSQL = gstrSQL & ",'" & mstrEndTime & "'"
                    gstrSQL = gstrSQL & ")"
    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
                    
                    '药品采购计划内容
                    If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
                        If .TextMatrix(intRow, mconIntCol常备药品) = "" Then .TextMatrix(intRow, mconIntCol常备药品) = True
                        If .TextMatrix(intRow, mconIntCol常备药品) = False Then
                            gstrSQL = "Select 送货单位,送货包装 From 药品规格 Where 药品id = [1]"
                            Set rsSpec = zlDataBase.OpenSQLRecord(gstrSQL, "查询送货单位", lngDrugID)
                            If IsNull(rsSpec!送货单位) = False Then
                                dbl送货数量 = GetFormat(dblRealQuantity / rsSpec!送货包装, 1)
                            End If
                            '向下找相同药品ID，合并数量
                            If CheckRepeatDrugID(recSort, n, lngDrugID) Then
                                '合并数量
                                SumQuantity arrSum, lngDrugID, dblQuantity
                            Else
                                intPlanSN = intPlanSN + 1
                                gstrSQL = "zl_药品计划管理次表_INSERT(" & _
                                          ID_IN & "," & _
                                          lngDrugID & "," & _
                                          intPlanSN & "," & _
                                          GetQuantity(arrSum, lngDrugID, dblQuantity) & "," & _
                                          dblPurchasePrice & "," & _
                                          dblPurchaseMoney & "," & _
                                          "null,null,0," & _
                                          IIf(lng上次供应商ID <= 0, "null", "'" & GetProvider(lng上次供应商ID) & "'") & "," & _
                                          IIf(strProducingArea = "", "null", "'" & strProducingArea & "'") & "," & _
                                          "null," & _
                                          dblSalePrice & "," & _
                                          dblSaleMoney & "," & _
                                          "null,null," & _
                                          dbl送货数量 & ")"
                                
                                ReDim Preserve arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail) + 1)
                                arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail)) = gstrSQL & ";"
                            End If
                        End If
                    End If
                End If
            End If
            recSort.MoveNext
        Next
        
        '药品采购计划
        If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
            strSQLDrugPlan = "zl_药品计划管理主表_INSERT(" & _
                             ID_IN & ",'" & _
                             NO_IN & "'," & _
                             "0," & _
                             "null," & _
                             lngStockid & "," & _
                             lngEnterStockID & "," & _
                             "0,'" & _
                             strBooker & "'," & _
                             "to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                             "由药品申购管理自动生成。')"
        End If
         
        If Not ExecuteSql(arrSql, strSQLDrugPlan, arrSQLDrugPlanDetail, MStrCaption) Then Exit Function
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
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
    On Error GoTo errHandle
    '检查该单据是否在进入编辑界面后，被其他操作员修改
    mstrTime_End = GetBillInfo(6, mstr单据号)
    If mstrTime_End = "" Then
        MsgBox "该单据已经被其他操作员删除！", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrTime_End > mstrTime_Start Then
        MsgBox "该单据已经被其他操作员编辑，请退出后重试！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查该单据是否被正常发送
    gstrSQL = " Select 配药日期 From 药品收发记录 " & _
            " Where 单据=6 And NO=[1] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查该单据是否被正常发送]", Me.txtNo.Tag)
    
    If IsNull(rsTemp!配药日期) Then
        MsgBox "该单据被其他操作员取消发送，不允许接收！", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    lng对方部门id = mlngStockID
    str审核人 = gstrUserName
    strNo = txtNo.Tag
    
    gstrSQL = "SELECT b.系数,b.id AS 类别id " _
            & " FROM 药品单据性质 a, 药品入出类别 b " _
            & "Where a.类别id = b.ID " _
            & "  AND a.单据 = 6 "
    
    Call SQLTest(App.Title, "药品移库管理", gstrSQL)
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "SaveCheck")
    Call SQLTest
    
    If rsTemp.EOF Then
        MsgBox "对不起，药品入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.RecordCount < 2 Then
        MsgBox "对不起，药品入出分类不全，请检查!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!系数 = 1 Then
            lng入类别id = rsTemp!类别id
        Else
            lng出类别id = rsTemp!类别id
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If mblnUpdate = False Then
        If Not 检查单价(6, txtNo.Tag, True, True) Then
            Call RefreshBill
            mblnUpdate = True
            Exit Function
        End If
    End If
    
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
                
                If Val(.TextMatrix(intRow, mconIntCol填写数量)) = Val(.TextMatrix(intRow, mconIntCol实际数量)) Then
                    num填写数量 = Val(.TextMatrix(intRow, mconIntCol原始数量))
                    num实际数量 = Val(.TextMatrix(intRow, mconIntCol原始数量))
                Else
                    num填写数量 = GetFormat(Val(.TextMatrix(intRow, mconIntCol填写数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                    num实际数量 = GetFormat(Val(.TextMatrix(intRow, mconIntCol实际数量)) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
'                num成本价 = GetFormat(Val(.TextMatrix(intRow, mconIntCol采购价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_成本价)
                num成本价 = Get成本价(lng药品ID, lng库房ID, lng出批次)
                num成本金额 = Val(.TextMatrix(intRow, mconIntCol采购金额))
'                dbl售价 = GetFormat(Val(.TextMatrix(intRow, mconIntCol售价)) / .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_零售价)
                dbl售价 = Get零售价(Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1, lng药品ID, lng库房ID, lng出批次)
                num零售金额 = Val(.TextMatrix(intRow, mconIntCol售价金额))
                num差价 = Val(.TextMatrix(intRow, mconintCol差价))
                str批号 = .TextMatrix(intRow, mconIntCol批号)
                dat效期 = IIf(.TextMatrix(intRow, mconIntCol效期) = "", "", .TextMatrix(intRow, mconIntCol效期))
                If gtype_UserSysParms.P149_效期显示方式 = 1 And dat效期 <> "" Then
                    '换算为失效期来保存
                    dat效期 = Format(DateAdd("D", 1, dat效期), "yyyy-mm-dd")
                End If
                                
                int序列号 = Val(.TextMatrix(intRow, mconIntCol序号))
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
                '供应商ID
                gstrSQL = gstrSQL & "," & IIf(lng上次供应商ID = 0, "NULL", lng上次供应商ID)
                '批准文号
                gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
                '零售价
                gstrSQL = gstrSQL & "," & dbl售价
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lng药品ID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
'    gcnOracle.BeginTrans
    If Not ExecuteSql(arrSql, "", "", MStrCaption) Then
'        gcnOracle.RollbackTrans
        Exit Function
    End If
'    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '外挂功能
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, lng库房ID, strNo, 单据号.药品移库)
    
    Exit Function
errHandle:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
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
    Dim intPricedigit As Integer
            
    On Error GoTo errHandle
    intPricedigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) B, 收费项目目录 C" & _
            " Where a.单据 = 6 And a.No = [1] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPricedigit & ") <> Round(b.现价, " & intPricedigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0  and b.执行日期>a.填制日期" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = 6 And a.No = [1] And c.Id = a.药品id And Round(a.零售价," & intPricedigit & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & intPricedigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id ,nvl(a.批次,0) as 批次, 0 原价, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = 6 And a.No = [1] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(b.平均成本价," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
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
                dbl零售价 = Val(GetFormat(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intPricedigit))
                dbl零售金额 = Val(GetFormat(dbl零售价 * Dbl数量, mintMoneyDigit))
                dbl差价 = Val(GetFormat(dbl零售金额 - dbl成本金额, mintMoneyDigit))
            End If
            
            rsPrice.Filter = "类型='成本价' And 药品ID=" & lng药品ID & " And 批次=" & Val(mshBill.TextMatrix(lngRow, mconIntCol批次))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl零售金额 = Val(GetFormat(dbl零售价 * Dbl数量, mintMoneyDigit))
                dbl成本价 = Val(GetFormat(rsPrice!现价 * Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), intCostDigit))
                dbl成本金额 = Val(GetFormat(dbl成本价 * Dbl数量, mintMoneyDigit))
                dbl差价 = Val(GetFormat(dbl零售金额 - dbl成本金额, mintMoneyDigit))
            End If
            
            If blnAdj = True Then
                '以当前最新价格最新单据相关数据（售价、成本价、零售金额、成本金额、差价）
                mshBill.TextMatrix(lngRow, mconIntCol售价) = GetFormat(dbl零售价, intPricedigit)
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
    Dim str药品 As String
    
    arrSql = Array()
    SaveStrike = False
    
    With mshBill
        '检查冲销数量，不能小于零
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol实际数量)) <> 0 Then
                If Not 相同符号(Val(.TextMatrix(intRow, mconIntCol填写数量)), Val(.TextMatrix(intRow, mconIntCol实际数量))) Then
                    MsgBox "请输入合法的冲销数量（第" & intRow & "行）！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '检查可用数量是否足够，参数设置为不检查库存时不进行
            If mint库存检查入库库房 <> 0 And .TextMatrix(intRow, 0) <> "" Then
                If .TextMatrix(intRow, mconIntCol实际数量) = .TextMatrix(intRow, mconIntCol填写数量) Then
                    冲销数量_IN = .TextMatrix(intRow, mconintCol真实数量)
                Else
                    冲销数量_IN = GetFormat(.TextMatrix(intRow, mconIntCol实际数量) * .TextMatrix(intRow, mconIntCol比例系数), gtype_UserDrugDigits.Digit_数量)
                End If
                
                If CheckStrickUsable(6, mlngStockID, Val(.TextMatrix(intRow, 0)), .TextMatrix(intRow, mconIntCol药名), _
                    Val(.TextMatrix(intRow, mconIntCol批次)), Val(冲销数量_IN), mint库存检查入库库房, Trim(txtNo.Tag), Val(.TextMatrix(intRow, mconIntCol序号)) + 1) = False Then
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
            End If
        Next
        
        '普通冲销检查实际数量
        If mint编辑状态 = 7 And mint处理方式 = 0 Then
            str药品 = CheckNumStock(mshBill, mlngStockID, 0, mconIntCol批次, mconIntCol实际数量, mconIntCol比例系数, 2, 0, mconintCol真实数量)
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
        填制人_IN = gstrUserName
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
        If str药品ID <> "" And mint处理方式 = 0 Then
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

Private Sub 提示库存数()
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim int发出库房分批属性 As Integer
    Dim int接收库房分批属性 As Integer
    Dim int工作性质 As Integer
    
    On Error GoTo errHandle
    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 5 Then
        Exit Sub
    Else
        With mshBill
            int工作性质 = 0
        
            '取发出和接收库房的分批属性
            '发出库房
            gstrSQL = "Select a.药库分批,a.药房分批,b.工作性质 " & _
                " From 药品规格 a,部门性质说明 b " & _
                " Where a.药品id = [2] And b.部门id = [1] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[发出库房分批属性]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)))
            
            Do While Not rsUseCount.EOF
                If int工作性质 <> 1 Then
                    If InStr(rsUseCount!工作性质, "药房") > 0 Or rsUseCount!工作性质 = "制剂室" Then
                        int工作性质 = 1
                        int发出库房分批属性 = rsUseCount!药房分批
                    ElseIf InStr(rsUseCount!工作性质, "药库") > 0 Then
                        int工作性质 = 2
                        int发出库房分批属性 = rsUseCount!药库分批
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            int工作性质 = 0
            '接收库房
            gstrSQL = "Select a.药库分批,a.药房分批,b.工作性质 " & _
                " From 药品规格 a,部门性质说明 b " & _
                " Where a.药品id = [2] And b.部门id = [1]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[接收库房分批属性]", mlngStockID, Val(.TextMatrix(.Row, 0)))
            
            Do While Not rsUseCount.EOF
                If int工作性质 <> 1 Then
                    If InStr(rsUseCount!工作性质, "药房") > 0 Or rsUseCount!工作性质 = "制剂室" Then
                        int工作性质 = 1
                        int接收库房分批属性 = rsUseCount!药房分批
                    ElseIf InStr(rsUseCount!工作性质, "药库") > 0 Then
                        int工作性质 = 2
                        int接收库房分批属性 = rsUseCount!药库分批
                    End If
                End If
                rsUseCount.MoveNext
            Loop
        
            
            If .TextMatrix(.Row, mconIntCol药名) = "" Then
                staThis.Panels(2).Text = ""
                Exit Sub
            End If
            If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
            '发出库存的当前药品的可用数量，统计批次数量，或者总的数量
            If mint明确批次 = 1 And int发出库房分批属性 = 1 Then
                gstrSQL = " Select 可用数量/" & .TextMatrix(.Row, mconIntCol比例系数) & " as 可用数量 from 药品库存 " & _
                          " Where 库房id=[1] " & _
                          " And 药品id=[2] And 性质=1 " & _
                          " And Nvl(批次,0)=[3]"
            Else
                gstrSQL = " Select Sum(可用数量)/" & .TextMatrix(.Row, mconIntCol比例系数) & " as 可用数量 from 药品库存 " & _
                          " Where 库房id=[1] " & _
                          " And 药品id=[2] And 性质=1 "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[发出库房可用数量]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol批次)))
            
            If rsUseCount.EOF Then
                .TextMatrix(.Row, mconIntCol可用数量) = 0
            Else
                .TextMatrix(.Row, mconIntCol可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            rsUseCount.Close
            
            '当前药房的可用数量，仅统计所有数量
            gstrSQL = " Select Sum(可用数量/" & .TextMatrix(.Row, mconIntCol比例系数) & ") as 可用数量 from 药品库存 where 库房id=[1] " & _
                      " And 药品id=[2] And 性质=1 "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[当前药房可用数量]", mlngStockID, Val(.TextMatrix(.Row, 0)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            
            Dim blnIs显示对方库存 As Boolean
            Dim str对方库存数 As String
            
            blnIs显示对方库存 = IsHavePrivs(mstrPrivs, "显示对方库存")
            str对方库存数 = "；" & Me.cboStock.Text & "库存数为[" & GetFormat(.TextMatrix(.Row, mconIntCol可用数量), mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol单位)
            
            staThis.Panels(2).Text = "该药品" & frmRequestDrugList.cboStock.Text & "库存数为[" & GetFormat(dblStock, mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol单位) _
                & IIf(blnIs显示对方库存, str对方库存数, "")
        End With
    End If
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

'转换数值为日期
Private Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 2000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    
End Function

'与可用数量进行比较
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl填写数量 As Double) As Boolean
    Dim dblUsableQuantity As Double      '实际数量对应的组成数量
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim dbl原填写数量 As Double
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        
    'mint库存检查: 0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '只要是分批药品，允许输入比当前批次大的数量，程序自动分解，而仅仅是时价药品属性的不允许
    CompareUsableQuantity = False
    If mint明确批次 = 0 Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        lng药品ID = .TextMatrix(intRow, 0)
        lng批次 = .TextMatrix(intRow, mconIntCol批次)
        
        gstrSQL = "Select Nvl(Sum(可用数量),0) 可用数量 From 药品库存 Where 库房ID=[1] And Nvl(批次,0)=[3] And 性质=1 And 药品ID=[2] "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[检查库存是否足够]", lng库房ID, lng药品ID, lng批次)
                
        dblUsableQuantity = rsCheck!可用数量 / .TextMatrix(intRow, mconIntCol比例系数)
        
        If .TextMatrix(intRow, mconIntCol批次) > 0 Or Split(.TextMatrix(intRow, mconIntCol最大效期), "||")(1) = 1 Then     '对移出库房是库房且药品是分批核算的药品和时价药品的判断
            If mint编辑状态 = 1 Or mint编辑状态 = 5 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & "：" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                If dbl填写数量 > dblUsableQuantity + dbl原填写数量 Then
                    MsgBox "第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & "：" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
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
                    If MsgBox("第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & "：" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbl原填写数量 = IIf(mbln下可用数量, numUsedCount, 0)
                
                If dbl填写数量 > dblUsableQuantity + dbl原填写数量 Then
                    If MsgBox("第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & "：" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + dbl原填写数量 & "”，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint库存检查 = 2 Then
            '2-检查，不足禁止
            If mint编辑状态 = 1 Then
                If dbl填写数量 > dblUsableQuantity Then
                    MsgBox "第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & ":" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint编辑状态 = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol批次) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If dbl填写数量 > dblUsableQuantity + numUsedCount Then
                    MsgBox "第" & intRow & "行药品" & .TextMatrix(intRow, mconIntCol药名) & "：" & vbCrLf & "对不起，你输入的数量“" & dbl填写数量 & "”大于了该药品的可用库存数量“" & dblUsableQuantity + numUsedCount & "”，请重输！", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
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

Private Function ExecuteSql(ByRef arrSql As Variant, ByVal strSQLDrugPlan As String _
    , ByRef arrSQLDrugPlanDetail As Variant, strTitle As String, Optional ByVal bln强制保存 As Boolean = False) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer

    ExecuteSql = False
    If UBound(arrSql) >= 0 Then
        '对SQL序列按药品ID升序排序
        For i = 0 To UBound(arrSql) - 1
            For j = i + 1 To UBound(arrSql)
                If CLng(Split(arrSql(j), ";")(0)) < CLng(Split(arrSql(i), ";")(0)) Then
                    strTmp = CStr(arrSql(j))
                    arrSql(j) = arrSql(i)
                    arrSql(i) = strTmp
                End If
            Next
        Next
        
        '执行SQL语句
        On Error GoTo errH
        If Not bln强制保存 Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(Mid(arrSql(i), InStr(arrSql(i), ";") + 1)), strTitle)
        Next
        '药品采购计划
        If Trim(strSQLDrugPlan) <> "" Then
            If UBound(arrSQLDrugPlanDetail) >= 0 Then
                Call zlDataBase.ExecuteProcedure(strSQLDrugPlan, strTitle & "-采购计划")
                For i = 0 To UBound(arrSQLDrugPlanDetail)
                    Call zlDataBase.ExecuteProcedure(CStr(Split(arrSQLDrugPlanDetail(i), ";")(0)), strTitle & "-采购计划内容")
                Next
            End If
        End If
        
        If Not bln强制保存 Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
errH:
    If Not bln强制保存 Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'打印单据
Private Sub printbill()
    Dim int单位系数 As Integer
    
    With mshBill
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
        FrmBillPrint.ShowMe Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), mint记录状态, int单位系数, 1304, "药品申领单", txtNo.Tag
    End With
End Sub


Private Sub get库存数量()
'''''''''''''''''''''''''''''''''''''
'获取库存数量的方法
'''''''''''''''''''''''''''''''''''''
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim int发出库房分批属性 As Integer
    Dim int接收库房分批属性 As Integer
    Dim int工作性质 As Integer '获取库房的工作性质，是药库还是药房
    Dim blnIs显示对方库存 As Boolean
    Dim str对方库存数 As String
    Dim i As Integer
    
    On Error GoTo errHandle
    With mshBill
        For i = 1 To .rows - 1
            If .TextMatrix(i, 0) = "" Then Exit Sub
            int工作性质 = 0
        
            '取发出和接收库房的分批属性
            '发出库房
            gstrSQL = "Select a.药库分批,a.药房分批,b.工作性质 " & _
                " From 药品规格 a,部门性质说明 b " & _
                " Where a.药品id = [2] And b.部门id = [1] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[发出库房分批属性]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)))
            
            Do While Not rsUseCount.EOF
                If int工作性质 <> 1 Then
                    If InStr(rsUseCount!工作性质, "药房") > 0 Or rsUseCount!工作性质 = "制剂室" Then
                        int工作性质 = 1
                        int发出库房分批属性 = rsUseCount!药房分批
                    ElseIf InStr(rsUseCount!工作性质, "药库") > 0 Then
                        int工作性质 = 2
                        int发出库房分批属性 = rsUseCount!药库分批
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            int工作性质 = 0
            '接收库房
            gstrSQL = "Select a.药库分批,a.药房分批,b.工作性质 " & _
                " From 药品规格 a,部门性质说明 b " & _
                " Where a.药品id = [2] And b.部门id = [1]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[接收库房分批属性]", mlngStockID, Val(.TextMatrix(i, 0)))
            
            Do While Not rsUseCount.EOF
                If int工作性质 <> 1 Then
                    If InStr(rsUseCount!工作性质, "药房") > 0 Or rsUseCount!工作性质 = "制剂室" Then
                        int工作性质 = 1
                        int接收库房分批属性 = rsUseCount!药房分批
                    ElseIf InStr(rsUseCount!工作性质, "药库") > 0 Then
                        int工作性质 = 2
                        int接收库房分批属性 = rsUseCount!药库分批
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
            blnIs显示对方库存 = IsHavePrivs(mstrPrivs, "显示对方库存")
            
            If blnIs显示对方库存 Then
                '发出库存的当前药品的可用数量，统计批次数量，或者总的数量
                If mint明确批次 = 1 And int发出库房分批属性 = 1 Then
                    gstrSQL = " Select 可用数量/" & .TextMatrix(i, mconIntCol比例系数) & " as 可用数量 from 药品库存 " & _
                              " Where 库房id=[1] " & _
                              " And 药品id=[2] And 性质=1 " & _
                              " And Nvl(批次,0)=[3]"
                Else
                    gstrSQL = " Select Sum(可用数量)/" & .TextMatrix(i, mconIntCol比例系数) & " as 可用数量 from 药品库存 " & _
                              " Where 库房id=[1] " & _
                              " And 药品id=[2] And 性质=1 "
                End If
                Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[发出库房可用数量]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, mconIntCol批次)))
                
                If rsUseCount.EOF Then
                    .TextMatrix(i, mconIntCol可用数量) = 0
                Else
                    .TextMatrix(i, mconIntCol可用数量) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
                End If
                .TextMatrix(i, mconintcol对方库存) = GetFormat(.TextMatrix(i, mconIntCol可用数量), mintNumberDigit)
                rsUseCount.Close
            End If
                
            '当前药房的可用数量，仅统计所有数量
            gstrSQL = " Select Sum(可用数量/" & .TextMatrix(i, mconIntCol比例系数) & ") as 可用数量 from 药品库存 where 库房id=[1] " & _
                      " And 药品id=[2] And 性质=1 "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[当前药房可用数量]", mlngStockID, Val(.TextMatrix(i, 0)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            .TextMatrix(i, mconintcol当前库存) = GetFormat(dblStock, mintNumberDigit)
       Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetProvider(ByVal lngProviderID As Long) As String
    Dim rsTemp As ADODB.Recordset
    
    If lngProviderID <= 0 Then Exit Function
    On Error GoTo errHandle
    gstrSQL = "select 名称 from 供应商 where id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "供应商名称", lngProviderID)
    If Not rsTemp.EOF Then
        GetProvider = NVL(rsTemp!名称)
    End If
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str批次 = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        If mint明确批次 = 1 Then
            str批次 = IIf(IsNull(rsTemp!批次), "0", rsTemp!批次)
        Else
            str批次 = "0"
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
                If mint明确批次 = 1 Then
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

Private Function Get价格(ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double) As Double
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
        If mint明确批次 = 1 Then
            Get价格 = 0
            Exit Function
        End If
    Else
        Get价格 = rsPrice.Fields(0).Value
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRepeatDrugID(ByVal rsTemp As ADODB.Recordset, ByVal intRecEnd As Integer, ByVal lngDrugID As Long) As Boolean
'----------------------
'功能：向下检查重复药品
'----------------------
    Dim i As Integer
    Dim rsClone As ADODB.Recordset
    
    CheckRepeatDrugID = False
    Set rsClone = rsTemp.Clone
    With rsClone
        .Sort = "药品id,批次,序号"
        .MoveFirst
        For i = 1 To .RecordCount
            If i > intRecEnd Then
                If lngDrugID = !药品id Then
                    CheckRepeatDrugID = True
                    Exit Function
                End If
            End If
            .MoveNext
        Next
    End With

End Function

Private Sub SumQuantity(ByRef arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double)
'------------------------
'功能：汇总同药品ID的数量
'------------------------
    Dim i As Integer
    Dim blnFind As Boolean
    
    If UBound(arrVal) > 0 Then
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                arrVal(1, i) = arrVal(1, i) + dblQTY
                blnFind = True
                Exit For
            End If
        Next
    Else
        ReDim arrVal(2, 1)
        arrVal(0, 0) = lngDrugID
        arrVal(1, 0) = dblQTY
        blnFind = True
    End If
    If blnFind = False Then
        ReDim Preserve arrVal(2, UBound(arrVal) + 1)
        arrVal(0, UBound(arrVal)) = lngDrugID
        arrVal(1, UBound(arrVal)) = dblQTY
    End If
End Sub

Private Function GetQuantity(ByVal arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double) As Double
'----------------------------
'功能：提取数组中药品ID的数量
'----------------------------
    If UBound(arrVal) > 0 Then
        Dim i As Integer
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                GetQuantity = arrVal(1, i) + dblQTY
                Exit Function
            End If
        Next
    End If
    GetQuantity = dblQTY
End Function


Private Function 检查价格() As Boolean
    '功能：新增时，判断药品是否是最新价格，不是则修改后提示
    Dim strMsg As String '保存提示信息
    Dim i As Integer, intSum As Integer, intPricedigit As Integer
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
                If mint编辑状态 = 6 Then   '审核用实际数量判断
                    dbl实际数量 = Val(GetFormat(NVL(rsCheck!实际数量, 0) / dbl比例系数, mintNumberDigit))
                Else    '审核用实际数量判断
                    dbl实际数量 = Val(GetFormat(NVL(rsCheck!可用数量, 0) / dbl比例系数, mintNumberDigit))
                End If
            End If
            
            '如果库存的实际数量不够
            If Not (dbl实际数量 >= dbl填写数量) Then
                int库存检查 = mint库存检查
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
                    If MsgBox(rsProperty!通用名 & "的库存不足，是否继续？" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
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

