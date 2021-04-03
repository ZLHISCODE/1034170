VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchUpdate 
   Caption         =   "批量修改规格"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmBatchUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6885
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   3000
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2520
      ScaleHeight     =   2415
      ScaleWidth      =   3855
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
      Begin VB.Frame fraSplit 
         Height          =   50
         Left            =   -120
         TabIndex        =   7
         Top             =   1440
         Width           =   3855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfOtherName 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
         _cx             =   6376
         _cy             =   873
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3375
         _cx             =   5953
         _cy             =   1720
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   Begin VB.PictureBox picDetails 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   2520
      ScaleHeight     =   1815
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin XtremeSuiteControls.TabControl tbcDetails 
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picClass 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      Begin VB.CheckBox chkAllDetails 
         Caption         =   "显示所有下级药品"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.TreeView tvwDetails 
         Height          =   4800
         Left            =   0
         TabIndex        =   9
         Tag             =   "1000"
         Top             =   600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgTvw"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5190
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBatchUpdate.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
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
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   1680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":70E6
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":7680
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":DEE2
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Bindings        =   "frmBatchUpdate.frx":E47C
      Left            =   1320
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBatchUpdate.frx":E490
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBatchUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint状态 As Integer         '记录是品种修改还是规格修改 1-品种 2-规格
Private mint次数 As Integer         '记录是不是首次加载 1-首次 2-不是
Private mblnData As Boolean  '用来判断是否在窗体加载时在树中有值
Private mstr上次节点 As String  '用来保存上次所选中的节点
Private mintRow As Integer        '用来记录上次所选中的行号
Private mintRow上次 As Integer
Private mintCol上次 As Integer
Private mbln库存 As Boolean        '用来记录是否有库存 true-有库存 flase-无库存
Private mbln药库分批 As Boolean    '药库分批 true-分批 false-不分批
Private mbln药房分批 As Boolean    '药房分批 true-分批 false-不分批
Private mint是否变价 As Integer     '定价还是时价 0-定价 1-时价
Private mstr类别 As String         '用来记录是什么分类 中草药，西成药、中成药
Private mstrNode As String         '记录被点击的节点的值
Private mstrPrivs As String        '记录用户有哪些权限
Private mrsRecord As ADODB.Recordset '用来记录选中节点查询出来的数据，为以后恢复数据做准备
Private mstrOtherName As String    '记录别名
Private mintOtherRow As Integer
Private mintExit As Integer         '用来记录退出时是否点击了保存按钮 1
Private mintLen As Integer          '记录住院单位的长度

'从参数表中取药品价格小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer

Private mstrFind As String           '用来记录要要查询的值
Private mlngFind As Long
Private mlngFindFirst As Long
Private mrsFindName As ADODB.Recordset
Private mstrValue As String         '用来记录查找框中的值

Private mstrMatch As String         '匹配方式
Private mstrOldValue As String      '记录原来的单元格中的值
Private mblnClick As Boolean
Private mblnSetKey As Boolean       '判断是否设置了
Private mint当前单位 As Integer      '用来系统参数中设置的显示单位

Private Const mcon应用于本列 As Integer = 101
Private Const mcon默认值 As Integer = 102
Private Const mcon保存 As Integer = 103
Private Const mcon帮助 As Integer = 104
Private Const mcon退出 As Integer = 105
Private Const mcon查找 As Integer = 106
Private Const mconFind As Integer = 107

Private Const cstcolor_backcolor = &H80000005   '白色
Private Const CSTCOLOR_UNMODIFY = &HC0C0FF       '粉红 选项页颜色
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '
Private Const mlngColor As Long = &H8000000F        '不能修改的列将背景颜色改成灰色
Private Const mlngApplyColor As Long = &HB18383          '淡蓝色

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl
Private mcbrToolBar As CommandBar


'品种类别
Private Enum mVariList
    基本信息 = 0
    品种属性 = 1
    临床应用 = 2
End Enum
'品种列
Private Enum mVaricolumn
    品种_序号 = 0
    品种_id = 1
    品种_分类id = 2
    品种_药品分类
    品种_药品编码
    品种_通用名称
    品种_英文名称
    品种_拼音码
    品种_五笔码
    '品种属性
    品种_毒理分类
    品种_价值分类
    品种_货源情况
    品种_用药梯次
    品种_药品类型
    品种_剂型
    品种_原研药
    品种_专利药
    品种_单独定价
    品种_急救药
    品种_新药
    品种_肿瘤药
    品种_溶媒
    品种_辅助用药
    品种_原料药
    品种_单味使用
    '临床应用
    品种_参考项目
    品种_处方职务
    品种_医保职务
    品种_处方限量
    品种_适用性别
    品种_剂量单位
    品种_皮试
    品种_抗生素
    品种_ATCCODE
    品种_品种下长期医嘱
    品种_参考项目ID
    品种_Count
End Enum

'规格类别
Private Enum mSpecList
    基本信息 = 0
    商品信息 = 1
    包装单位 = 2
    价格信息 = 3
    药价属性 = 4
    分批管理 = 5
    临床应用 = 6
    配药属性 = 7
End Enum

'规格列
Private Enum mSpecColumn
    规格_序号 = 0
    规格_id = 1
    规格_药名id = 2
'    规格_药品编码 = 3
    规格_通用名称
    规格_规格编码
    规格_药品规格
    规格_本位码
    规格_数字码
    规格_标识码
    规格_备选码
    规格_容量
    规格_商品名称
    规格_生产厂牌
    规格_来源分类
    规格_拼音码
    规格_五笔码
    规格_合同单位
    规格_批准文号
    规格_注册商标
    规格_GMP认证
    规格_非常备药
    规格_售价单位
    规格_剂量系数
    规格_剂量单位
    规格_住院单位
    规格_住院系数
    规格_门诊单位
    规格_门诊系数
    规格_药库单位
    规格_药库系数
    规格_送货单位
    规格_送货包装
    规格_申领单位
    规格_申领阀值
    规格_中药形态
    规格_药价属性
    规格_采购限价
    规格_采购扣率
    规格_结算价
    规格_指导售价
    规格_指导差率
    规格_加成率
    规格_差价让利
    规格_成本价格
    规格_当前售价
    规格_收入项目
    规格_病案费目
    规格_管理费比例
    规格_药价级别
    规格_屏蔽费别
    规格_增值税率
    规格_医保类型
    规格_药库分批
    规格_药房分批
    规格_保质期
    规格_标识说明
    规格_发药类型
    规格_站点编号
    规格_DDD值
    规格_服务对象
    规格_住院分零使用
    规格_住院动态分零
    规格_门诊分零使用
    规格_高危药品
    规格_基本药物
    规格_存储温度
    规格_存储条件
    规格_配药类型
    规格_输液注意事项
    规格_不予调配
    规格_招标药品
    规格_合同单位id
    规格_收入项目id
    规格_原药库分批
    规格_原药房分批
    规格_count = 75
End Enum

Private Sub CheckValue(ByVal intRow As Integer, ByVal lng药品ID As Long)
    Dim rsTemp As ADODB.Recordset
    Dim dblTemp As Double
    
    gstrSql = ""
    On Error GoTo ErrHandle
    With vsfDetails
        If .TextMatrix(intRow, mSpecColumn.规格_药库分批) = "0" Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药房分批, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.规格_药房分批) = 0
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.规格_保质期) = 0
        Else
            If Val(.TextMatrix(intRow, mSpecColumn.规格_保质期)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor
            End If
        End If
        
        '提取显示当前售价
        If Mid(.TextMatrix(intRow, mSpecColumn.规格_药价属性), 1, 1) <> 0 Then
            '时价药品，取库存金额/库存数量做为其价格，无库存时取价表定价 非时价药品调价，取其价格记录中的价格
            gstrSql = "select Decode(K.库存数量,0,P.现价,K.库存金额/Nvl(K.库存数量,1)) as 现价,P.收入项目id" & _
                    " from 收费价目 P," & _
                    "     (Select nvl(Sum(实际金额),0) as 库存金额,nvl(Sum(实际数量),0) as 库存数量" & _
                    "      From 药品库存 Where 药品ID=[1]) K" & _
                    " where P.收费细目id=[1] and (P.终止日期 is null or Sysdate Between P.执行日期 And P.终止日期)"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        End If
    
        If gstrSql <> "" Then
            If rsTemp.RecordCount > 0 Then
                If Val(mint当前单位) <> 0 Then
                    .TextMatrix(intRow, mSpecColumn.规格_当前售价) = FormatEx(rsTemp!现价 * Val(.TextMatrix(intRow, mSpecColumn.规格_药库系数)), mintPriceDigit)
                Else
                    .TextMatrix(intRow, mSpecColumn.规格_当前售价) = FormatEx(rsTemp!现价, mintPriceDigit)
                End If
                .TextMatrix(intRow, mSpecColumn.规格_收入项目id) = rsTemp!收入项目id
            End If
        End If

        '根据是否有发生，确定：药价属性、成本价格、零售价格可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品收发记录 Where 药品ID=[1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        If rsTemp.Fields(0).Value > 0 Then
            If Mid(.TextMatrix(intRow, mSpecColumn.规格_药价属性), 1, 1) <> 0 Then .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药价属性, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_成本价格, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_当前售价, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_收入项目, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院系数, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_门诊系数, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库系数, intRow) = mlngColor
        End If
        
        '根据品种是否是抗生素确定规格DDD值是否能够修改
        gstrSql = " Select Nvl(b.抗生素, 0) As 抗生素 From 药品规格 A, 药品特性 B Where a.药名id = b.药名id And a.药品id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        If rsTemp!抗生素 = 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_DDD值, intRow) = mlngColor
        End If
        
        '根据是否存在医嘱记录，确定剂量系数是否能够修改
        gstrSql = "Select 1 From 病人医嘱记录 Where 收费细目ID=[1] And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        If rsTemp.RecordCount > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_剂量系数, intRow) = mlngColor
        End If
        
        '根据是否有库存，确定：分批特性可修改否
        gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                 " Where A.药品ID=[1] And A.库房ID=B.部门ID And B.工作性质 Like '%药库'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
        
        If rsTemp.Fields(0).Value > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_保质期, intRow) = mlngColor
        End If
        If .TextMatrix(intRow, mSpecColumn.规格_药库分批) <> "0" Then
            gstrSql = " Select nvl(Count(*),0) From 药品库存 A,部门性质说明 B" & _
                     " Where A.药品ID=[1] And A.库房ID=B.部门ID And (B.工作性质 Like '%药房' Or B.工作性质 Like '%制剂室')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng药品ID)
            
            If rsTemp.Fields(0).Value > 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药房分批, intRow) = mlngColor
                If .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批) <> mlngColor Then
                    .Cell(flexcpBackColor, intRow, mSpecColumn.规格_药库分批, intRow) = IIf(.TextMatrix(intRow, mSpecColumn.规格_药房分批) = "0", cstcolor_backcolor, mlngColor)
                End If
            End If
        End If
            .Cell(flexcpBackColor, intRow, mSpecColumn.规格_结算价, intRow) = mlngColor
            If Val(Mid(.TextMatrix(intRow, mSpecColumn.规格_住院分零使用), 1, 1)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院动态分零, intRow) = mlngColor
            End If
            If .TextMatrix(intRow, mSpecColumn.规格_中药形态) = "散装" And mstrNode Like "中草药*" Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_住院分零使用, intRow) = mlngColor
                .Cell(flexcpBackColor, intRow, mSpecColumn.规格_门诊分零使用, intRow) = mlngColor
            End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal int状态 As Integer, ByVal strPrivs As String)
    '提供其他窗体访问本窗体的公用方法
    mint状态 = int状态
    mstrPrivs = strPrivs
    
    Me.Show vbModal, frmMediLists
End Sub

Private Sub InitTreeView()
    With tvwDetails
        .LabelEdit = 1  '设置treeview为不可编辑状态
    End With
End Sub

Private Sub InitComandBars()
    '初始化工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

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
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons
    
    '工具栏定义
    Set mcbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny
    
    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mcon应用于本列, "应用于本列")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, mcon默认值, "恢复默认值")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        Set cbrControlMain = .Add(xtpControlButton, mcon保存, "保存")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
'        Set cbrControlMain = .Add(xtpControlButton, mcon帮助, "帮助")
'        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mcon退出, "退出")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
        Set cbrControlMain = .Add(xtpControlLabel, mcon查找, "查找")
        cbrControlMain.Flags = xtpFlagRightAlign    '靠右对齐

        Set ctrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, mconFind, "查询")
        ctrCustom.Handle = txtFind.hWnd
        ctrCustom.Flags = xtpFlagRightAlign
    End With
    
    cbsMain.Item(1).Delete
    
    '右键菜单
    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopup.Controls
        Set mobjControl = .Add(xtpControlButton, mcon应用于本列, "应用于本列")
        Set mobjControl = .Add(xtpControlButton, mcon默认值, "恢复默认值")
    End With
    
    '快键绑定
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconFind
    End With
End Sub

Private Sub initPanel()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneCon As Pane
    Dim objPaneDetail As Pane
    
    Me.dkpPanel.SetCommandBars Me.cbsMain
    Me.dkpPanel.Options.UseSplitterTracker = False '实时拖动
    Me.dkpPanel.Options.ThemedFloatingFrames = True
    Me.dkpPanel.Options.AlphaDockingContext = True
    
    Set objPaneCon = Me.dkpPanel.CreatePane(1, 200, 0, DockLeftOf, Nothing)
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    objPaneCon.Title = "分类"
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTemp As String
    
    Select Case Control.ID
        Case mcon应用于本列
            Call SetBatch
        Case mcon默认值
'            mrsRecord.MoveFirst
'            Call showColumn(mrsRecord, mstrNode)
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        Case mcon保存
            Call Save
        Case mconFind
'            If TypeName(Control) = "ICommandBarButton" Then
'                Call FindGridRow(mstrValue)
'            Else
'                strTemp = Trim(UCase(Control.Text))
'                mstrValue = strTemp
'            End If
'            Call FindGridRow(strTemp)
        Case mcon退出
            Call ExitFrom
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Me.picDetails.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - stbThis.Height
    
    Call initControl
End Sub

Private Sub chkAllDetails_Click()
    If mint状态 = 1 Then
        With vsfDetails
            If chkAllDetails.Value = 1 Then
                .ColWidth(mVaricolumn.品种_药品分类) = 2000
                .ColHidden(mVaricolumn.品种_药品分类) = False
            Else
                .ColHidden(mVaricolumn.品种_药品分类) = True
            End If
        End With
    End If
    Call tvwDetails_NodeClick(tvwDetails.SelectedItem)
End Sub

Private Sub dkpPanel_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picClass.hWnd '将控件加入到dockingpanel控件中
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    Me.Width = 14000    '第一次加载时，窗体大小
    Me.Height = 9000
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        chkAllDetails = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "是否显示下级", 0)
    End If
    
    mint次数 = 1
    Call InitTreeView   '初始化树
    Call InitComandBars '初始化菜单和工具栏
    Call initPanel  '初始化面板
    Call InitTabControl '向TabControl控件中加入窗体
    Call initControl    '初始化控件
    
    If mint状态 = 1 Then
        Call initColumn_品种信息    '初始化品种列
        mint次数 = 2
    ElseIf mint状态 = 2 Then
        Call initColumn_规格信息
        mint次数 = 2
    End If
    
'    mstrNode = "西成药"
    mblnData = ReadAndSendDataToTvw(mint状态)     '往树中填充值
    Call setColumn(0)    '初始化vsflexgrid控件列
    Call Set权限判断 '权限判断
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")  '匹配方式
    mint当前单位 = Val(GetSysPara(29))  '记录当前设置的显示单位
        
    mintCostDigit = GetDigit(1, 1, IIf(mint当前单位 = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(mint当前单位 = 0, 1, 4))
    
    mintSaleCostDigit = GetDigit(1, 1, 1)
    mintSalePriceDigit = GetDigit(1, 2, 1)
    
    If tvwDetails.Nodes.Count > 0 Then
        If chkAllDetails = 1 And Not tvwDetails.Nodes(tvwDetails.SelectedItem.Index) Is Nothing Then
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        End If
    End If
End Sub

Private Sub initControl()
    '重新布局控件位置
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    chkAllDetails.Move 0, 0, picClass.Width
    tvwDetails.Move 0, chkAllDetails.Height + chkAllDetails.Top, picClass.ScaleWidth, lngBottom - lngTop - stbThis.Height - chkAllDetails.Height - 300
    tbcDetails.Move 0, 0, picDetails.ScaleWidth, picDetails.ScaleHeight
     
    If mint状态 = 1 Then    '品种才有别名
        frmBatchUpdate.Caption = "品种批量修改"
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    Else    '规格无别名
        frmBatchUpdate.Caption = "规格批量修改"
        
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call initControl
End Sub

Private Sub InitTabControl()
    '初始化Tabcontrol控件
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        If mint状态 = 1 Then    '品种
            .InsertItem(mVariList.基本信息, "基本信息", picList.hWnd, 0).Tag = "基本信息_"
            .InsertItem(mVariList.品种属性, "品种属性", picList.hWnd, 0).Tag = "品种属性_"
            .InsertItem(mVariList.临床应用, "临床应用", picList.hWnd, 0).Tag = "临床应用_"
            
            .Item(mVariList.品种属性).Selected = True
            .Item(mVariList.基本信息).Selected = True
            
        Else    '规格
            .InsertItem(mSpecList.基本信息, "基本信息", picList.hWnd, 0).Tag = "基本信息_"
            .InsertItem(mSpecList.商品信息, "商品信息", picList.hWnd, 0).Tag = "商品信息_"
            .InsertItem(mSpecList.包装单位, "包装单位", picList.hWnd, 0).Tag = "包装单位_"
            .InsertItem(mSpecList.价格信息, "价格信息", picList.hWnd, 0).Tag = "价格信息_"
            .InsertItem(mSpecList.药价属性, "药价属性", picList.hWnd, 0).Tag = "药价属性_"
            .InsertItem(mSpecList.分批管理, "分批管理", picList.hWnd, 0).Tag = "分批管理_"
            .InsertItem(mSpecList.临床应用, "临床应用", picList.hWnd, 0).Tag = "临床应用_"
            .InsertItem(mSpecList.配药属性, "配药属性", picList.hWnd, 0).Tag = "配药属性_"
            
            .Item(mSpecList.商品信息).Selected = True
            .Item(mSpecList.基本信息).Selected = True
        End If
    End With
    Call setTabControlColor(tbcDetails)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Recover
    mblnSetKey = False
    mintExit = 0
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    If GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "1") = "1" Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "是否显示下级", chkAllDetails.Value)
    End If
    Unload Me
End Sub
Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And vsfDetails.Height + y > 100 And fraSplit.Height + fraSplit.Top + y < stbThis.Top - 1000 Then
        vsfDetails.Move 0, 0, picList.ScaleWidth, vsfDetails.Height + y
        fraSplit.Move 0, fraSplit.Top + y, picList.ScaleWidth, 50
        vsfOtherName.Move 0, fraSplit.Top + fraSplit.Height, picList.ScaleWidth, vsfOtherName.Height - y
    End If
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'只有在药品品种基本页面才会有别名供用户修改

    If mint状态 = 1 Then    '品种
        fraSplit.Visible = False
        vsfOtherName.Visible = False
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
    Else
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
    End If
    
    Call setTabControlColor(tbcDetails)

    If mint次数 = 2 Then    '只有在列初始化后才能进行列设置
        Call setColumn(Item.Index)  '列隐藏显示设置
    End If
End Sub

Private Sub setTabControlColor(ByVal objtbc As TabControl)
    '对Tabcontrol控件进行颜色判断
    Dim i As Integer
    
    With objtbc
        For i = 0 To .ItemCount - 1
            If .Item(i).Selected = True Then
                .Item(i).Color = CSTCOLOR_UNMODIFY
            Else
                .Item(i).Color = CSTCOLOR_NORECORDS
            End If
        Next
    End With
End Sub

Private Sub setColumn(ByVal intPageItem As Integer)
    '列显示与隐藏设置
    With vsfDetails
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeRestrictColumns
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '不能多选单元格
    End With
    
    With vsfDetails
        If mint状态 = 1 Then '品种
            vsfDetails.MergeCol(mVaricolumn.品种_药品分类) = True   '与上面的.MergeCells属性结合使用不同行同列内容相同的合并
            '基本信息
            .ColWidth(mVaricolumn.品种_序号) = 600
            .ColHidden(mVaricolumn.品种_id) = True
            .ColHidden(mVaricolumn.品种_分类id) = True
            .ColHidden(mVaricolumn.品种_参考项目ID) = True
            
            .ColWidth(mVaricolumn.品种_通用名称) = 2000 '不隐藏该列
            .ColHidden(mVaricolumn.品种_药品编码) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_英文名称) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_拼音码) = IIf(intPageItem = mVariList.基本信息, False, True)
            .ColHidden(mVaricolumn.品种_五笔码) = IIf(intPageItem = mVariList.基本信息, False, True)
            
            '品种属性
            .ColHidden(mVaricolumn.品种_毒理分类) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_价值分类) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_货源情况) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_用药梯次) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_药品类型) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_剂型) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_原研药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_专利药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_单独定价) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_急救药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_新药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_原料药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_单味使用) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_辅助用药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_肿瘤药) = IIf(intPageItem = mVariList.品种属性, False, True)
            .ColHidden(mVaricolumn.品种_溶媒) = IIf(intPageItem = mVariList.品种属性, False, True)
            
            '临床应用
            .ColHidden(mVaricolumn.品种_参考项目) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_处方职务) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_医保职务) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_处方限量) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_适用性别) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_剂量单位) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_皮试) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_抗生素) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_ATCCODE) = IIf(intPageItem = mVariList.临床应用, False, True)
            .ColHidden(mVaricolumn.品种_品种下长期医嘱) = IIf(intPageItem = mVariList.临床应用, False, True)
            
            If mstrNode Like "中草药*" And intPageItem = mVariList.临床应用 Then
                .ColHidden(mVaricolumn.品种_皮试) = True
                .ColHidden(mVaricolumn.品种_抗生素) = True
                .ColHidden(mVaricolumn.品种_ATCCODE) = True
                .ColHidden(mVaricolumn.品种_品种下长期医嘱) = True
            Else
                If intPageItem = mVariList.临床应用 Then
                    .ColHidden(mVaricolumn.品种_皮试) = False
                    .ColHidden(mVaricolumn.品种_抗生素) = False
                    .ColHidden(mVaricolumn.品种_ATCCODE) = False
                    .ColHidden(mVaricolumn.品种_品种下长期医嘱) = False
                End If
            End If
            
            If mstrNode Like "中草药*" Then
                If intPageItem = mVariList.品种属性 Then
                    .ColHidden(mVaricolumn.品种_单味使用) = False
                    .ColHidden(mVaricolumn.品种_原料药) = False
                End If
                .ColHidden(mVaricolumn.品种_剂型) = True
                .ColHidden(mVaricolumn.品种_原研药) = True
                .ColHidden(mVaricolumn.品种_专利药) = True
                .ColHidden(mVaricolumn.品种_单独定价) = True
                .ColHidden(mVaricolumn.品种_急救药) = True
                .ColHidden(mVaricolumn.品种_新药) = True
                .ColHidden(mVaricolumn.品种_肿瘤药) = True
                .ColHidden(mVaricolumn.品种_溶媒) = True
            Else
                .ColHidden(mVaricolumn.品种_单味使用) = True
                If intPageItem = mVariList.品种属性 Then
                    .ColHidden(mVaricolumn.品种_剂型) = False
                    .ColHidden(mVaricolumn.品种_原研药) = False
                    .ColHidden(mVaricolumn.品种_专利药) = False
                    .ColHidden(mVaricolumn.品种_单独定价) = False
                    .ColHidden(mVaricolumn.品种_急救药) = False
                    .ColHidden(mVaricolumn.品种_新药) = False
                    .ColHidden(mVaricolumn.品种_原料药) = False
                    .ColHidden(mVaricolumn.品种_肿瘤药) = False
                    .ColHidden(mVaricolumn.品种_溶媒) = False
                End If
            End If
            
            If chkAllDetails.Value = 1 Then
                .ColHidden(mVaricolumn.品种_药品分类) = False
            Else
                .ColHidden(mVaricolumn.品种_药品分类) = True
            End If
        Else    '规格
            vsfDetails.MergeCol(mSpecColumn.规格_通用名称) = True    '设置合并
            
            .ColWidth(mSpecColumn.规格_序号) = 600
'            .ColWidth(mSpecColumn.规格_药品编码) = 1000
            .ColWidth(mSpecColumn.规格_通用名称) = 1800
            .ColWidth(mSpecColumn.规格_药品规格) = 1500
            .ColHidden(mSpecColumn.规格_id) = True
            .ColHidden(mSpecColumn.规格_药名id) = True
            .ColHidden(mSpecColumn.规格_招标药品) = True
            .ColHidden(mSpecColumn.规格_合同单位id) = True
            .ColHidden(mSpecColumn.规格_收入项目id) = True
            .ColHidden(mSpecColumn.规格_原药库分批) = True
            .ColHidden(mSpecColumn.规格_原药房分批) = True
            '基本信息
            .ColHidden(mSpecColumn.规格_规格编码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_本位码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_数字码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_标识码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            .ColHidden(mSpecColumn.规格_备选码) = IIf(intPageItem = mSpecList.基本信息, False, True)
            
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_容量) = True
            Else
                .ColHidden(mSpecColumn.规格_容量) = IIf(intPageItem = mSpecList.基本信息, False, True)
            End If
            '商品信息
            .ColHidden(mSpecColumn.规格_商品名称) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_生产厂牌) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_来源分类) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_合同单位) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_批准文号) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_注册商标) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_拼音码) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_五笔码) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_GMP认证) = IIf(intPageItem = mSpecList.商品信息, False, True)
            .ColHidden(mSpecColumn.规格_非常备药) = IIf(intPageItem = mSpecList.商品信息, False, True)
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_拼音码) = True
                .ColHidden(mSpecColumn.规格_五笔码) = True
                .ColHidden(mSpecColumn.规格_GMP认证) = True
                .ColHidden(mSpecColumn.规格_商品名称) = True
            End If
            
            '包装单位
            .ColHidden(mSpecColumn.规格_售价单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_剂量系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_剂量单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_住院单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_住院系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_门诊单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_门诊系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_药库单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_药库系数) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_送货单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_送货包装) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_申领单位) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_申领阀值) = IIf(intPageItem = mSpecList.包装单位, False, True)
            .ColHidden(mSpecColumn.规格_中药形态) = IIf(intPageItem = mSpecList.包装单位, False, True)
            
            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.包装单位 Then
                    .ColHidden(mSpecColumn.规格_中药形态) = False
                    .ColHidden(mSpecColumn.规格_门诊单位) = True
                    .ColHidden(mSpecColumn.规格_门诊系数) = True
                    VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "药房单位", 1000, flexAlignLeftCenter, "药房单位"
                    VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "药房系数", 1000, flexAlignRightCenter, "药房系数"
                End If
            Else
                VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "住院单位", 1000, flexAlignLeftCenter, "住院单位"
                VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "住院系数", 1000, flexAlignRightCenter, "住院系数"
                .ColHidden(mSpecColumn.规格_中药形态) = True
            End If
            '价格信息
            .ColHidden(mSpecColumn.规格_药价属性) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_采购限价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_采购扣率) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_结算价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_指导售价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_指导差率) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_加成率) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_差价让利) = IIf(intPageItem = mSpecList.价格信息, False, True)
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_差价让利) = True
            Else
                If intPageItem = mSpecList.价格信息 Then
                    .ColHidden(mSpecColumn.规格_差价让利) = False
                Else
                    .ColHidden(mSpecColumn.规格_差价让利) = True
                End If
            End If
            .ColHidden(mSpecColumn.规格_成本价格) = IIf(intPageItem = mSpecList.价格信息, False, True)
            .ColHidden(mSpecColumn.规格_当前售价) = IIf(intPageItem = mSpecList.价格信息, False, True)
            '药价属性
            .ColHidden(mSpecColumn.规格_收入项目) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_病案费目) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_管理费比例) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_药价级别) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_屏蔽费别) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_增值税率) = IIf(intPageItem = mSpecList.药价属性, False, True)
            .ColHidden(mSpecColumn.规格_医保类型) = IIf(intPageItem = mSpecList.药价属性, False, True)
            '分批管理
            .ColHidden(mSpecColumn.规格_药库分批) = IIf(intPageItem = mSpecList.分批管理, False, True)
            .ColHidden(mSpecColumn.规格_药房分批) = IIf(intPageItem = mSpecList.分批管理, False, True)
            .ColHidden(mSpecColumn.规格_保质期) = IIf(intPageItem = mSpecList.分批管理, False, True)
            
            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.分批管理 Then
                    .ColHidden(mSpecColumn.规格_保质期) = True
                End If
            Else
                If intPageItem = mSpecList.分批管理 Then
                    .ColHidden(mSpecColumn.规格_保质期) = False
                End If
            End If
            
            '临床应用
            .ColHidden(mSpecColumn.规格_标识说明) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_发药类型) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_站点编号) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_DDD值) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_服务对象) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_住院分零使用) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_门诊分零使用) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_基本药物) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_住院动态分零) = IIf(intPageItem = mSpecList.临床应用, False, True)
            .ColHidden(mSpecColumn.规格_高危药品) = IIf(intPageItem = mSpecList.临床应用, False, True)
            If mstrNode Like "中草药*" Then
                .ColHidden(mSpecColumn.规格_基本药物) = True
                .ColHidden(mSpecColumn.规格_住院动态分零) = True
                .ColHidden(mSpecColumn.规格_高危药品) = True
            Else
                If intPageItem = mSpecList.临床应用 Then
                    .ColHidden(mSpecColumn.规格_基本药物) = False
                    .ColHidden(mSpecColumn.规格_住院动态分零) = False
                    .ColHidden(mSpecColumn.规格_高危药品) = False
                End If
            End If
            
            '配药属性
            .ColHidden(mSpecColumn.规格_存储温度) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_存储条件) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_配药类型) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_不予调配) = IIf(intPageItem = mSpecList.配药属性, False, True)
            .ColHidden(mSpecColumn.规格_输液注意事项) = IIf(intPageItem = mSpecList.配药属性, False, True)
            
            If mstrNode Like "中草药*" Then
                If intPageItem = mSpecList.配药属性 Then
                    tbcDetails.Item(mSpecList.基本信息).Selected = True
                End If
                tbcDetails.Item(mSpecList.配药属性).Visible = False
                .ColHidden(mSpecColumn.规格_存储温度) = True
                .ColHidden(mSpecColumn.规格_存储条件) = True
                .ColHidden(mSpecColumn.规格_配药类型) = True
                .ColHidden(mSpecColumn.规格_不予调配) = True
                .ColHidden(mSpecColumn.规格_输液注意事项) = True
            Else
                tbcDetails.Item(mSpecList.配药属性).Visible = True
                If intPageItem = mSpecList.配药属性 Then
                    .ColHidden(mSpecColumn.规格_存储温度) = False
                    .ColHidden(mSpecColumn.规格_存储条件) = False
                    .ColHidden(mSpecColumn.规格_配药类型) = False
                    .ColHidden(mSpecColumn.规格_不予调配) = False
                    .ColHidden(mSpecColumn.规格_输液注意事项) = False
                Else
                    .ColHidden(mSpecColumn.规格_存储温度) = True
                    .ColHidden(mSpecColumn.规格_存储条件) = True
                    .ColHidden(mSpecColumn.规格_配药类型) = True
                    .ColHidden(mSpecColumn.规格_不予调配) = True
                    .ColHidden(mSpecColumn.规格_输液注意事项) = True
                End If
            End If
        End If
    End With
End Sub

Private Sub initColumn_品种信息()
    '初始化基本信息页面
    Dim rsRecord As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    With vsfDetails
        .Cols = mVaricolumn.品种_Count
        .Rows = 1
        '基本信息
        VsfGridColFormat vsfDetails, mVaricolumn.品种_序号, "序号", 600, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_id, "id", 300, flexAlignCenterCenter, "id"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_分类id, "分类id", 300, flexAlignCenterCenter, "分类id"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品分类, "药品分类", 2000, flexAlignLeftCenter, "药品分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品编码, "药品编码", 1000, flexAlignLeftCenter, "药品编码"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_通用名称, "通用名称", 1000, flexAlignLeftCenter, "通用名称"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_英文名称, "英文名称", 1000, flexAlignLeftCenter, "英文名称"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_拼音码, "拼音码", 1000, flexAlignLeftCenter, "拼音码"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_五笔码, "五笔码", 1000, flexAlignLeftCenter, "五笔码"
        '品种属性
        VsfGridColFormat vsfDetails, mVaricolumn.品种_毒理分类, "毒理分类", 1000, flexAlignLeftCenter, "毒理分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_价值分类, "价值分类", 1000, flexAlignLeftCenter, "价值分类"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_货源情况, "货源情况", 1000, flexAlignLeftCenter, "货源情况"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_用药梯次, "用药梯次", 1000, flexAlignLeftCenter, "用药梯次"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_药品类型, "药品类型", 1000, flexAlignLeftCenter, "药品类型"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_剂型, "剂型", 2000, flexAlignLeftCenter, "剂型"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_原研药, "原研药", 800, flexAlignCenterCenter, "原研药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_专利药, "专利药", 800, flexAlignCenterCenter, "专利药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_单独定价, "单独定价", 1000, flexAlignCenterCenter, "单独定价"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_急救药, "急救药", 800, flexAlignCenterCenter, "急救药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_新药, "新药", 800, flexAlignRightCenter, "新药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_原料药, "原料药", 1000, flexAlignLeftCenter, "原料药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_单味使用, "单味使用", 1000, flexAlignLeftCenter, "单味使用"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_辅助用药, "辅助用药", 1000, flexAlignCenterCenter, "辅助用药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_肿瘤药, "肿瘤药", 1000, flexAlignLeftCenter, "肿瘤药"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_溶媒, "溶媒", 1000, flexAlignCenterCenter, "溶媒"
        '临床应用
        VsfGridColFormat vsfDetails, mVaricolumn.品种_参考项目, "参考项目", 1000, flexAlignLeftCenter, "参考项目"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_处方职务, "处方职务", 1000, flexAlignLeftCenter, "处方职务"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_医保职务, "医保职务", 1000, flexAlignLeftCenter, "医保职务"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_处方限量, "处方限量", 1000, flexAlignRightCenter, "处方限量"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_适用性别, "使用性别", 1500, flexAlignLeftCenter, "使用性别"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_剂量单位, "剂量单位", 1000, flexAlignLeftCenter, "剂量单位"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_皮试, "皮试", 800, flexAlignCenterCenter, "皮试"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_抗生素, "抗生素", 1500, flexAlignLeftCenter, "抗生素"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_ATCCODE, "ATCCODE", 1500, flexAlignLeftCenter, "ATCCODE"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_品种下长期医嘱, "品种下长期医嘱", 1500, flexAlignLeftCenter, "品种下长期医嘱"
        VsfGridColFormat vsfDetails, mVaricolumn.品种_参考项目ID, "参考项目id", 10, flexAlignLeftCenter, "参考项目id"
        
        If chkAllDetails.Value = 1 Then
            .ColWidth(mVaricolumn.品种_药品分类) = 2000
        Else
            .ColHidden(mVaricolumn.品种_药品分类) = True
        End If
    End With
    
    With vsfDetails
        '原研药
        .ColDataType(mVaricolumn.品种_原研药) = flexDTBoolean
        '专利药
        .ColDataType(mVaricolumn.品种_专利药) = flexDTBoolean
        '单独定价
        .ColDataType(mVaricolumn.品种_单独定价) = flexDTBoolean
        '急救药
        .ColDataType(mVaricolumn.品种_急救药) = flexDTBoolean
        '新药
        .ColDataType(mVaricolumn.品种_新药) = flexDTBoolean
        '辅助用药
        .ColDataType(mVaricolumn.品种_辅助用药) = flexDTBoolean
        '原料药
        .ColDataType(mVaricolumn.品种_原料药) = flexDTBoolean
        '肿瘤药
        .ColDataType(mVaricolumn.品种_肿瘤药) = flexDTBoolean
        '溶媒
        .ColDataType(mVaricolumn.品种_溶媒) = flexDTBoolean
        '皮试
        .ColDataType(mVaricolumn.品种_皮试) = flexDTBoolean
        '药品按品种下长期医嘱
        .ColDataType(mVaricolumn.品种_品种下长期医嘱) = flexDTBoolean
        '抗生素
        .ColComboList(mVaricolumn.品种_抗生素) = "0-非抗生素|1-非限制使用|2-限制使用|3-特殊使用"
        '单味使用
        .ColDataType(mVaricolumn.品种_单味使用) = flexDTBoolean
    
        '剂量单位
        gstrSql = "select distinct 计算单位 from 诊疗项目目录 where 类别  in ('5','6','7')"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        If Not rsRecord.EOF Then
            For i = 1 To rsRecord.RecordCount
                strTemp = strTemp & "|" & rsRecord!计算单位
                rsRecord.MoveNext
            Next
        End If
        .ColComboList(mVaricolumn.品种_剂量单位) = strTemp
        
        '剂型
        gstrSql = "select 编码||'-'|| 名称 as 剂型 from 药品剂型"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_剂型) = vsfDetails.BuildComboList(rsRecord, "剂型")
        '参考项目
        .ColComboList(mVaricolumn.品种_参考项目) = "|..."
        '毒理分类
        gstrSql = "select 编码||'-'|| 名称 as 毒理分类 from 药品毒理分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_毒理分类) = vsfDetails.BuildComboList(rsRecord, "毒理分类")
        '价值分类
        gstrSql = "select 编码||'-'|| 名称 as 价值分类 from 药品价值分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_价值分类) = vsfDetails.BuildComboList(rsRecord, "价值分类")
        '货源情况
        gstrSql = "select 编码||'-'|| 名称 as 货源情况 from 药品货源情况"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_货源情况) = vsfDetails.BuildComboList(rsRecord, "货源情况")
        '用药梯次
        gstrSql = "select 编码||'-'|| 名称 as 用药梯次 from 药品用药梯次"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mVaricolumn.品种_用药梯次) = vsfDetails.BuildComboList(rsRecord, "用药梯次")
        '药品类型
        .ColComboList(mVaricolumn.品种_药品类型) = "0-未设定|1-处方药|2-甲类非处方药|3-乙类非处方药|4-非处方药|5-其它用药"
        '处方职务
        .ColComboList(mVaricolumn.品种_处方职务) = "0-不限|1-正高|2-副高|3-中级|4-助理/师级|5-员/士|9-待聘"
        '医保职务
        .ColComboList(mVaricolumn.品种_医保职务) = "0-不限|1-正高|2-副高|3-中级|4-助理/师级|5-员/士|9-待聘"
        '适用性别
        .ColComboList(mVaricolumn.品种_适用性别) = "0-无性别区分|1-男性|2-女性"
        
    End With
End Sub

Private Sub initColumn_规格信息()
    Dim rsRecord As ADODB.Recordset
    
    '初始化规格列
    On Error GoTo ErrHandle
    With vsfDetails
        .Cols = mSpecColumn.规格_count
        .Rows = 1
        '基本信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_序号, "序号", 600, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_id, "id", 300, flexAlignLeftCenter, "id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药名id, "药名id", 600, flexAlignCenterCenter, "药名id"
'        VsfGridColFormat vsfDetails, mSpecColumn.规格_药品编码, "药品编码", 300, flexAlignLeftCenter, "药品编码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_通用名称, "通用名称", 1000, flexAlignLeftCenter, "通用名称"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_规格编码, "规格编码", 1000, flexAlignLeftCenter, "规格编码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药品规格, "药品规格", 1500, flexAlignLeftCenter, "药品规格"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_本位码, "本位码", 2500, flexAlignLeftCenter, "本位码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_数字码, "数字码", 1000, flexAlignLeftCenter, "数字码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_标识码, "标识码", 1000, flexAlignLeftCenter, "标识码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_备选码, "备选码", 1000, flexAlignLeftCenter, "备选码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_容量, "容量", 800, flexAlignRightCenter, "容量"
        '商品信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_商品名称, "商品名称", 1500, flexAlignLeftCenter, "商品名称"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_生产厂牌, "生产厂牌", 1500, flexAlignLeftCenter, "生产厂牌"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_来源分类, "来源分类", 1000, flexAlignLeftCenter, "来源分类"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_拼音码, "拼音码", 1000, flexAlignLeftCenter, "拼音码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_五笔码, "五笔码", 1000, flexAlignLeftCenter, "五笔码"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_合同单位, "合同单位", 1000, flexAlignLeftCenter, "合同单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_批准文号, "批准文号", 1000, flexAlignLeftCenter, "批准文号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_注册商标, "注册商标", 1000, flexAlignLeftCenter, "注册商标"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_GMP认证, "GMP认证", 800, flexAlignCenterCenter, "GMP认证"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_非常备药, "非常备药", 800, flexAlignCenterCenter, "非常备药"
        '包装单位
        VsfGridColFormat vsfDetails, mSpecColumn.规格_售价单位, "售价单位", 1000, flexAlignLeftCenter, "售价单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_剂量系数, "剂量系数", 1000, flexAlignRightCenter, "剂量系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_剂量单位, "剂量单位", 1000, flexAlignRightCenter, "剂量单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院单位, "住院单位", 1000, flexAlignLeftCenter, "住院单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院系数, "住院系数", 1000, flexAlignRightCenter, "住院系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊单位, "门诊单位", 1000, flexAlignLeftCenter, "门诊单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊系数, "门诊系数", 1000, flexAlignRightCenter, "门诊系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库单位, "药库单位", 1000, flexAlignLeftCenter, "药库单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库系数, "药库系数", 1000, flexAlignRightCenter, "药库系数"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_送货单位, "送货单位", 1000, flexAlignRightCenter, "送货单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_送货包装, "送货包装", 1000, flexAlignRightCenter, "送货包装"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_申领单位, "申领单位", 1000, flexAlignLeftCenter, "申领单位"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_申领阀值, "申领阀值", 1000, flexAlignRightCenter, "申领阀值"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_中药形态, "中药形态", 1000, flexAlignRightCenter, "中药形态"
        '价格信息
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药价属性, "药价属性", 900, flexAlignLeftCenter, "药价属性"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_采购限价, "采购限价", 1000, flexAlignRightCenter, "采购限价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_采购扣率, "采购扣率", 1000, flexAlignRightCenter, "采购扣率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_结算价, "结算价", 1000, flexAlignRightCenter, "结算价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_指导售价, "指导售价", 1000, flexAlignRightCenter, "指导售价"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_指导差率, "指导差率", 0, flexAlignRightCenter, "指导差率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_加成率, "加成率", 1000, flexAlignRightCenter, "加成率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_差价让利, "差价让利", 1000, flexAlignRightCenter, "差价让利"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_成本价格, "成本价格", 1000, flexAlignRightCenter, "成本价格"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_当前售价, "当前售价", 1000, flexAlignRightCenter, "当前售价"
        '药价属性
        VsfGridColFormat vsfDetails, mSpecColumn.规格_收入项目, "收入项目", 1500, flexAlignLeftCenter, "收入项目"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_病案费目, "病案费目", 1000, flexAlignLeftCenter, "病案费目"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_管理费比例, "管理费比例", 1000, flexAlignRightCenter, "管理费比例"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药价级别, "药价级别", 1000, flexAlignLeftCenter, "药价级别"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_屏蔽费别, "屏蔽费别", 900, flexAlignLeftCenter, "屏蔽费别"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_增值税率, "增值税率", 1000, flexAlignRightCenter, "增值税率"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_医保类型, "医保类型", 1000, flexAlignLeftCenter, "医保类型"
        '分批管理
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药库分批, "药库分批", 800, flexAlignCenterCenter, "药库分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_药房分批, "药房分批", 800, flexAlignCenterCenter, "药房分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_原药库分批, "原药库分批", 800, flexAlignCenterCenter, "原药库分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_原药房分批, "原药房分批", 800, flexAlignCenterCenter, "原药房分批"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_保质期, "保质期(月)", 1000, flexAlignRightCenter, "保质期(月)"
        '临床应用
        VsfGridColFormat vsfDetails, mSpecColumn.规格_标识说明, "标识说明", 1000, flexAlignLeftCenter, "标识说明"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_发药类型, "发药类型", 900, flexAlignLeftCenter, "发药类型"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_站点编号, "站点编号", 900, flexAlignLeftCenter, "站点编号"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_DDD值, "DDD值", 900, flexAlignLeftCenter, "DDD值"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_服务对象, "服务对象", 1500, flexAlignLeftCenter, "服务对象"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院分零使用, "住院分零使用", 1300, flexAlignLeftCenter, "住院分零使用"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_门诊分零使用, "门诊分零使用", 1300, flexAlignLeftCenter, "门诊分零使用"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_住院动态分零, "住院动态分零", 1300, flexAlignLeftCenter, "住院动态分零"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_基本药物, "基本药物", 1000, flexAlignLeftCenter, "基本药物"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_高危药品, "高危药品", 1000, flexAlignLeftCenter, "高危药品"
        '配药属性
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储温度, "存储温度", 1500, flexAlignLeftCenter, "存储温度"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_存储条件, "存储条件", 1000, flexAlignLeftCenter, "存储条件"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_配药类型, "配药类型", 1500, flexAlignLeftCenter, "配药类型"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_不予调配, "不予调配", 1000, flexAlignLeftCenter, "不予调配"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_招标药品, "招标药品", 1000, flexAlignLeftCenter, "招标药品"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_合同单位id, "合同单位id", 1000, flexAlignLeftCenter, "合同单位id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_收入项目id, "收入项目id", 1000, flexAlignLeftCenter, "收入项目id"
        VsfGridColFormat vsfDetails, mSpecColumn.规格_输液注意事项, "输液注意事项", 5000, flexAlignLeftCenter, "输液注意事项"
    End With
    
    With vsfDetails
        '屏蔽费别
        .ColDataType(mSpecColumn.规格_屏蔽费别) = flexDTBoolean
        '住院/门诊动态分零
        .ColDataType(mSpecColumn.规格_住院动态分零) = flexDTBoolean
        'GMP认证
        .ColDataType(mSpecColumn.规格_GMP认证) = flexDTBoolean
        '非常备药
        .ColDataType(mSpecColumn.规格_非常备药) = flexDTBoolean
        '药库分批
        .ColDataType(mSpecColumn.规格_药库分批) = flexDTBoolean
        '药房分批
        .ColDataType(mSpecColumn.规格_药房分批) = flexDTBoolean
        '存储条件
        .ColDataType(mSpecColumn.规格_存储条件) = flexDTBoolean
        '不予调配
        .ColDataType(mSpecColumn.规格_不予调配) = flexDTBoolean
        
        '生产厂牌
        .ColComboList(mSpecColumn.规格_生产厂牌) = "|..."
        '来源分类
        gstrSql = "select 编码||'-'|| 名称 as 来源分类 from 药品来源分类"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_来源分类) = vsfDetails.BuildComboList(rsRecord, "来源分类")
        '合同单位
        .ColComboList(mSpecColumn.规格_合同单位) = "|..."
        '发药类型
        gstrSql = "select 名称 as 发药类型 from 发药类型"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_发药类型) = vsfDetails.BuildComboList(rsRecord, "发药类型")
        '站点编号
        gstrSql = "select 编号||'-'||名称 as 站点编号 from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_站点编号) = vsfDetails.BuildComboList(rsRecord, "站点编号")
        '申领单位
        .ColComboList(mSpecColumn.规格_申领单位) = "售价单位|住院单位|门诊单位|药库单位"
        '药价属性
        .ColComboList(mSpecColumn.规格_药价属性) = "0-定价|1-时价"
        '基本药物
        gstrSql = "Select 名称 as 基本药物  From 基本药物说明  Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_基本药物) = vsfDetails.BuildComboList(rsRecord, "基本药物")
        '收入项目
        gstrSql = "Select ID, '[' || 编码 || ']' || 名称 As 收入项目" & _
                  "  From 收入项目" & _
                  "  Where 末级 = 1 And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                  "  Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_收入项目) = vsfDetails.BuildComboList(rsRecord, "收入项目")
        '病案费目
        .ColComboList(mSpecColumn.规格_病案费目) = "..."
        '药价管理级别
        gstrSql = "select 编码||'-'||名称 as 管理级别 from 药价管理级别"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_药价级别) = vsfDetails.BuildComboList(rsRecord, "管理级别")
        '医保类型
        gstrSql = "Select 编码||'-'||名称 as 医保类型 From 费用类型 where 性质=1 Order By 编码"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_品种信息")
        .ColComboList(mSpecColumn.规格_医保类型) = vsfDetails.BuildComboList(rsRecord, "医保类型")
        '服务对象
        .ColComboList(mSpecColumn.规格_服务对象) = "0-不应用于病人|1-门诊|2-住院|3-门诊和住院"
        '住院/门诊分零使用
        .ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        .ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        '存储温度
        .ColComboList(mSpecColumn.规格_存储温度) = " |1-常温(0-30℃)|2-阴凉(20℃以下)|3-冷藏(2-8℃)"
        '配药类型
        gstrSql = "Select 编码||'-'|| 名称 as 配药类型 From 输液配药类型 "
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_规格信息")
        .ColComboList(mSpecColumn.规格_配药类型) = " |" & vsfDetails.BuildComboList(rsRecord, "配药类型")
        '中药形态
        .ColComboList(mSpecColumn.规格_中药形态) = "散装|中药饮片|免煎剂"
        '高危药品
        .ColComboList(mSpecColumn.规格_高危药品) = " |1-A级|2-B级|3-C级"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadAndSendDataToTvw(ByVal int状态 As Integer) As Boolean
'功能：用来向树中填充节点
'参数 int状态 用来判断界面加载时是品种修改还是规格修改

    Dim NodeThis As Node
    Dim Int末级 As Integer
    Dim lng库房ID As Long
    Dim rs材质分类 As ADODB.Recordset
    Dim recdata As ADODB.Recordset
    
    '药品用途分类是否有数据
    ReadAndSendDataToTvw = False
    On Error GoTo ErrHandle
    gstrSql = " Select 编码,名称 From 诊疗项目类别 " & _
              " Where Instr([1],编码,1) > 0 " & _
              " Order by 编码"
    Set rs材质分类 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "567")
    
    If rs材质分类 Is Nothing Then
        Exit Function
    End If
    
'    Set rs材质分类 = GetFilter分类(rs材质分类)
    With tvwDetails
        .Nodes.Clear
        Do While Not rs材质分类.EOF
            .Nodes.Add , , "Root" & rs材质分类!名称, rs材质分类!名称, 1, 1
            .Nodes("Root" & rs材质分类!名称).Tag = rs材质分类!编码
            rs材质分类.MoveNext
        Loop
    End With
    
'    gstrSql = "Select Rownum As ID, ID As 项目id, 上级id, 编码, 名称, 分类, 类别" & _
'               " From (Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
'               " From 诊疗分类目录" & _
'               " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'               " Start With 上级id Is Null" & _
'               "  Connect By Prior ID = 上级id " & _
'               " Union All " & _
'               " Select ID, 分类id, 编码, 名称, Decode(类别, 5, '西成药', 6, '中成药', 7, '中草药') 分类, '品种' As 类别" & _
'               " From 诊疗项目目录" & _
'               " Where 分类id In (Select ID " & _
'               "               From 诊疗分类目录" & _
'               "               Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'               "               Start With 上级id Is Null " & _
'               "               Connect By Prior ID = 上级id)) " & _
'        " Start  With 上级id Is Null" & _
'        " Connect By Prior ID = 上级id order by id,项目id"

    gstrSql = "Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
            " From 诊疗分类目录" & _
            " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' " & _
            " Start With 上级id Is Null" & _
            " Connect By Prior ID = 上级id"

    Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "ReadAndSendDataToTvw")
    
    If recdata.EOF Then
        MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
        Exit Function
    End If
    
    With recdata
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set NodeThis = tvwDetails.Nodes.Add("Root" & !分类, 4, "K_" & !ID, !名称, 1, 1)
            Else
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !上级ID, 4, "K_" & !ID, !名称, 1, 1)
            End If
            NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With
    
    If int状态 <> 1 Then '品种修改
        gstrSql = "Select ID, 分类id, 编码, 名称, Decode(类别, 5, '西成药', 6, '中成药', 7, '中草药') 分类, '品种' As 类别" & _
                  "  From 诊疗项目目录" & _
                  "  Where 分类id In (Select ID" & _
                                   " From 诊疗分类目录" & _
                                   " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
                                   " Start With 上级id Is Null" & _
                                   " Connect By Prior ID = 上级id)"
        Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "品种")
        
        With recdata
            Do While Not .EOF
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !分类id, 4, !类别 & "K_" & !ID, !名称, 1, 1)
                NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
                .MoveNext
            Loop
        End With
    End If
        
    Call GetFilter权限  '根据用户所具有的权限来过滤数据
    
    With tvwDetails
        If .Nodes.Count <> 0 Then
            .Nodes(1).Selected = True
            If .Nodes(1).Children <> 0 Then
                Int末级 = 1
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(2).Children <> 0 Then
                Int末级 = 2
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(3).Children <> 0 Then
                Int末级 = 3
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            Else
                Int末级 = 0
                .Nodes(1).Selected = True
                .SelectedItem.Selected = True
            End If
            If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
        End If
    End With
    
    ReadAndSendDataToTvw = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetFilter权限()
    Dim strTemp As String
    
    With tvwDetails
        If mint状态 = 1 Then
            If InStr(1, mstrPrivs, "管理西成药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root西成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中成药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root中成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中草药品种") = 0 Then
                .Nodes.Remove (.Nodes("Root中草药").Index)
            End If
        Else
            If InStr(1, mstrPrivs, "管理西成药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root西成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中成药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root中成药").Index)
            End If
            If InStr(1, mstrPrivs, "管理中草药规格") = 0 Then
                .Nodes.Remove (.Nodes("Root中草药").Index)
            End If
        End If
    End With
End Sub

Private Sub tvwDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    '节点点击事件
    Dim rsRecord As ADODB.Recordset
    Dim lngkey As Long  '用来保存所选中的key值
    Dim str分类 As String   '药品规格修改中用来判断选中的节点是品种还是分类
    Dim intupdate As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bln修改 As Boolean  '用来记录是否有值被修改了
    
    If Node Is Nothing Then
        Exit Sub
    End If
    mstrNode = Node.Tag '记录节点中的值
    mblnClick = False
    
    On Error GoTo ErrHandle
    If Node.Tag Like "中草药*" And mint状态 = 2 Then
        vsfDetails.ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零"
        vsfDetails.ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零"
    ElseIf mint状态 = 2 Then
        vsfDetails.ColComboList(mSpecColumn.规格_住院分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
        vsfDetails.ColComboList(mSpecColumn.规格_门诊分零使用) = "0-可以分零|1-不可分零|2-一次性使用|3-分零后一天内有效|4-分零后两天内有效|5-分零后三天内有效"
    End If
    If Node.Key Like "Root*" Then Exit Sub  '如果选择的节点时最顶级节点则退出
'    If Node.Key = mstr上次节点 Then
'        Exit Sub
'    Else
'        mstr上次节点 = Node.Key
'    End If
    
    '判断界面中是否有值刚被修改了
    bln修改 = Check修改
    
    If bln修改 = True Then
        intupdate = MsgBox("刚有内容被修改了，是否继续？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbNo Then Exit Sub
    End If
    
    If mint状态 = 1 Then    '品种
        gstrSql = "Select Distinct a.id,c.id as 类别id,a.参考目录ID, '['||c.编码||']'||c.名称 as 名称, a.编码, a.名称 As 通用名称, d.英文名, d.拼音码, d.五笔码, e.毒理分类, e.价值分类, e.货源情况, e.用药梯次, nvl(e.药品类型,0) as 药品类型, e.药品剂型, nvl(e.急救药否,0)  as 急救药否," & _
                        "  e.是否肿瘤药 as 肿瘤药, e.溶媒, e.ATCCODE, e.是否原研药, e.是否专利药, e.是否单独定价, nvl(e.是否新药,0) as 是否新药, nvl(e.是否原料,0) as 是否原料, f.名称 As 参考项目, nvl(e.处方职务,'00') as 处方职务, nvl(e.处方限量,0) as 处方限量, Nvl(a.适用性别,0) AS 适用性别, a.计算单位 As 剂量单位, nvl(e.是否皮试,0) as 是否皮试, nvl(e.抗生素,0) as 抗生素, nvl(e.品种医嘱,0) as 品种医嘱,a.单独应用 as 单味使用,e.是否辅助用药 as 辅助用药" & _
                        "  From 诊疗项目目录 A, 诊疗项目别名 B, 诊疗分类目录 C," & _
                    " (Select n.诊疗项目id, n.名称, n.拼音码, m.五笔码, p.英文名" & _
                    "  From (Select 诊疗项目id, 名称, 简码 As 拼音码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 1) N," & _
                    "       (Select 诊疗项目id, 名称, 简码 As 五笔码 From 诊疗项目别名 Where 性质 = 1 And 码类 = 2) M," & _
                    "       (Select 诊疗项目id, 名称 As 英文名 From 诊疗项目别名 Where 性质 = 2) P" & _
                    "  Where n.诊疗项目id = m.诊疗项目id And n.诊疗项目id = p.诊疗项目id) D, 药品特性 E, 诊疗参考目录 F " & _
                    "   Where a.Id = b.诊疗项目id(+) And a.分类id = c.Id And a.Id = d.诊疗项目id(+) And a.Id = e.药名id And a.参考目录id = f.Id(+) And " & _
                    " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') And " & _
                    " a.分类id "
        
        If chkAllDetails.Value = 1 Then '当选择了显示所有节点中的数据时
            gstrSql = gstrSql & " in (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id) order by id"
        Else
            gstrSql = gstrSql & " =[1] order by id"
        End If
    Else    '规格
        str分类 = Node.Tag
        If str分类 Like "*品种" Then '选中的是品种节点
            gstrSql = "Select a.Id, c.药名id, a.编码 As 规格编码, a.规格, j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名, a.产地 As 生产厂牌, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标," & _
                              " c.Gmp认证, c.是否常备, a.计算单位 As 售价单位, c.剂量系数 As 售价系数,j.计算单位, c.住院单位, c.住院包装, c.门诊单位, c.门诊包装, c.药库单位, c.药库包装, c.申领单位, c.申领阀值," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.指导差价率 As 指导差率, c.差价让利比 As 差价让利, c.成本价," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目, c.管理费比例,c.容量, c.药价级别, a.屏蔽费别, c.增值税率, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, c.合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零, c.动态分零 as 住院动态分零,c.门诊可否分零, c.基本药物, a.站点 As 站点编号,C.ddd值, i.存储温度, i.存储条件," & _
                              " i.配药类型, i.是否不予配置 As 不予调配,C.本位码,c.高危药品, c.送货单位, c.送货包装,i.输液注意事项 " & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 诊疗项目目录 J, 供应商 D, 收费价目 E," & _
                            " 收入项目 F, 输液药品属性 I" & _
                       " Where c.药名id = j.Id And j.Id = [1] And a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD') And a.Id = c.药品id And" & _
                             " c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And a.Id = m.收费细目id(+) And" & _
                             " a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+)  and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                       " Order By a.Id"

        Else    '选中的是分类节点
            gstrSql = " Select a.Id, c.药名id, a.编码 As 规格编码, a.规格, j.编码 As 品种编码, j.名称 As 通用名称, m.数字码, c.标识码, a.备选码," & _
                              " Decode(n.商品名, Null, p.商品名, n.商品名) 商品名, a.产地 As 生产厂牌, n.拼音码, p.五笔码, c.药品来源 As 来源分类, d.名称 As 合同单位, c.批准文号, c.注册商标, " & _
                              " c.Gmp认证, c.是否常备, a.计算单位 As 售价单位, c.剂量系数 As 售价系数,j.计算单位, c.住院单位, c.住院包装, c.门诊单位, c.门诊包装, c.药库单位, c.药库包装, c.申领单位, c.申领阀值," & _
                              " c.中药形态, a.是否变价 As 药价属性, c.指导批发价 As 采购限价, c.扣率 As 采购扣率, c.指导零售价 As 指导售价, c.指导差价率 As 指导差率, c.差价让利比 As 差价让利, c.成本价," & _
                              " e.现价 As 当前售价, f.名称 As 收入项目,a.病案费目, c.管理费比例, c.容量,c.药价级别, a.屏蔽费别, c.增值税率, a.费用类型 As 医保类型, c.药库分批, c.药房分批, c.招标药品, 合同单位id," & _
                              " e.收入项目id, c.最大效期 As 保质期, a.说明 As 标识说明, c.发药类型, a.服务对象, c.住院可否分零, c.动态分零 as 住院动态分零,c.门诊可否分零,c.基本药物, a.站点 As 站点编号,c.DDD值, i.存储温度, i.存储条件," & _
                              " i.配药类型, i.是否不予配置 As 不予调配,C.本位码,c.高危药品, c.送货单位, c.送货包装,i.输液注意事项" & _
                       " From 收费项目目录 A, (Select 收费细目id, 简码 As 数字码 From 收费项目别名 Where 码类 = 3 And 性质 = 1) M," & _
                            " (Select 收费细目id, 简码 As 拼音码, 名称 As 商品名 From 收费项目别名 Where 码类 = 1 And 性质 = 3) N," & _
                            " (Select 收费细目id, 简码 As 五笔码, 名称 As 商品名 From 收费项目别名 Where 码类 = 2 And 性质 = 3) P, 药品规格 C, 供应商 D, 收费价目 E, 收入项目 F," & _
                            " 输液药品属性 I, 诊疗项目目录 J" & _
                       " Where a.Id In" & _
                            "  (Select 药品id" & _
                              " From 药品规格" & _
                              " Where 药名id In " & _
                                    " (Select ID " & _
                                    "  From 诊疗项目目录 " & _
                                     " Where 分类id In " & _
                                          "  (Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id))) And" & _
                             " a.撤档时间 = To_Date('3000-1-1', 'yyyy-MM-DD')" & _
                             " And a.Id = c.药品id And c.合同单位id = d.Id(+) And e.收费细目id = a.Id And e.收入项目id = f.Id And a.Id = i.药品id(+) And" & _
                             " c.药名id = j.Id And a.Id = m.收费细目id(+) And a.Id = n.收费细目id(+) And a.Id = p.收费细目id(+) and (e.终止日期 is null or Sysdate Between e.执行日期 And e.终止日期)" & _
                       " Order By j.名称,a.Id"

        End If
        Call setColumn(tbcDetails.Selected.Index)
        If chkAllDetails.Value = 0 Then '不能获取到下级节点
            If Node.Tag Like "*分类" Then
                vsfDetails.Rows = 1
                Exit Sub
            End If
        End If
    End If

    If mint状态 = 2 Then '规格
        If Node.Tag Like "中草药*" Then  '是否显示配药属性
            tbcDetails.Item(mSpecList.配药属性).Visible = False
            
            With vsfDetails
                .ColHidden(mSpecColumn.规格_存储温度) = True
                .ColHidden(mSpecColumn.规格_存储条件) = True
                .ColHidden(mSpecColumn.规格_配药类型) = True
                .ColHidden(mSpecColumn.规格_不予调配) = True
                .ColHidden(mSpecColumn.规格_输液注意事项) = True
                If tbcDetails.Selected.Index = tbcDetails.ItemCount - 1 Then
                    tbcDetails.Item(mSpecList.基本信息).Selected = True
                End If
            End With
        Else
            tbcDetails.Item(mSpecList.配药属性).Visible = True
            With vsfDetails
                If tbcDetails.Item(mSpecList.配药属性).Selected = True Then
                    .ColHidden(mSpecColumn.规格_存储温度) = False
                    .ColHidden(mSpecColumn.规格_存储条件) = False
                    .ColHidden(mSpecColumn.规格_配药类型) = False
                    .ColHidden(mSpecColumn.规格_不予调配) = False
                    .ColHidden(mSpecColumn.规格_输液注意事项) = False
                Else
                    .ColHidden(mSpecColumn.规格_存储温度) = True
                    .ColHidden(mSpecColumn.规格_存储条件) = True
                    .ColHidden(mSpecColumn.规格_配药类型) = True
                    .ColHidden(mSpecColumn.规格_不予调配) = True
                    .ColHidden(mSpecColumn.规格_输液注意事项) = True
                End If
            End With
        End If
    End If
    '获取key值
    lngkey = Mid(Node.Key, InStr(1, Node.Key, "_") + 1, Len(Node.Key) - InStr(1, Node.Key, "_"))
    Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "节点点击", lngkey)
    
    vsfDetails.Rows = 1
    If rsRecord.EOF Then
        Call setColumn(tbcDetails.Selected.Index)
        Exit Sub
    End If
    Set mrsRecord = rsRecord.Clone  '克隆
    
    Call showColumn(rsRecord, Node.Tag)   '将值绑定到vsflexgrid控件中
    Call setColumn(tbcDetails.Selected.Index)
    Call GetDefineSize(rsRecord)
    With vsfDetails
        If .Rows > 1 Then
            .Row = 1
            .Col = mVaricolumn.品种_通用名称
        End If
    End With
    Call Set权限判断
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub showColumn(ByVal rsRecord As ADODB.Recordset, ByVal str分类 As String)
    '当点击树节点时，将值绑定到vsflexgrid控件中
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim bln剂量系数 As Boolean

    vsfDetails.Rows = rsRecord.RecordCount + 1 '根据查询出来的值的数量来确定列表行数
    
    vsfDetails.Select 1, 1
    If mint状态 = 1 Then    '品种
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mVaricolumn.品种_序号) = i
                .TextMatrix(i, mVaricolumn.品种_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mVaricolumn.品种_分类id) = IIf(IsNull(rsRecord!类别id), "", rsRecord!类别id)
                .TextMatrix(i, mVaricolumn.品种_药品分类) = IIf(IsNull(rsRecord!名称), "", rsRecord!名称)
                .TextMatrix(i, mVaricolumn.品种_药品编码) = IIf(IsNull(rsRecord!编码), "", rsRecord!编码)
                .TextMatrix(i, mVaricolumn.品种_通用名称) = IIf(IsNull(rsRecord!通用名称), "", rsRecord!通用名称)
                .TextMatrix(i, mVaricolumn.品种_英文名称) = IIf(IsNull(rsRecord!英文名), "", rsRecord!英文名)
                .TextMatrix(i, mVaricolumn.品种_拼音码) = IIf(IsNull(rsRecord!拼音码), "", rsRecord!拼音码)
                .TextMatrix(i, mVaricolumn.品种_五笔码) = IIf(IsNull(rsRecord!五笔码), "", rsRecord!五笔码)
                
                If .TextMatrix(i, mVaricolumn.品种_拼音码) = "" Then
                    .TextMatrix(i, mVaricolumn.品种_拼音码) = zlGetSymbol(.TextMatrix(i, mVaricolumn.品种_通用名称), 0, 30)
                End If
                
                If .TextMatrix(i, mVaricolumn.品种_五笔码) = "" Then
                    .TextMatrix(i, mVaricolumn.品种_五笔码) = zlGetSymbol(.TextMatrix(i, mVaricolumn.品种_通用名称), 1, 30)
                End If
                
                .TextMatrix(i, mVaricolumn.品种_毒理分类) = ShowValue(.ColComboList(mVaricolumn.品种_毒理分类), IIf(IsNull(rsRecord!毒理分类), "", rsRecord!毒理分类), "-")
                .TextMatrix(i, mVaricolumn.品种_价值分类) = ShowValue(.ColComboList(mVaricolumn.品种_价值分类), IIf(IsNull(rsRecord!价值分类), "", rsRecord!价值分类), "-")
                .TextMatrix(i, mVaricolumn.品种_货源情况) = ShowValue(.ColComboList(mVaricolumn.品种_货源情况), IIf(IsNull(rsRecord!货源情况), "", rsRecord!货源情况), "-")
                .TextMatrix(i, mVaricolumn.品种_用药梯次) = ShowValue(.ColComboList(mVaricolumn.品种_用药梯次), IIf(IsNull(rsRecord!用药梯次), "", rsRecord!用药梯次), "-")
                .TextMatrix(i, mVaricolumn.品种_药品类型) = ShowValue(.ColComboList(mVaricolumn.品种_药品类型), IIf(IsNull(rsRecord!药品类型), "", rsRecord!药品类型), "-", True)
                .TextMatrix(i, mVaricolumn.品种_剂型) = ShowValue(.ColComboList(mVaricolumn.品种_剂型), IIf(IsNull(rsRecord!药品剂型), "", rsRecord!药品剂型), "-")
                .TextMatrix(i, mVaricolumn.品种_原研药) = IIf(IsNull(rsRecord!是否原研药), "", rsRecord!是否原研药)
                .TextMatrix(i, mVaricolumn.品种_专利药) = IIf(IsNull(rsRecord!是否专利药), "", rsRecord!是否专利药)
                .TextMatrix(i, mVaricolumn.品种_单独定价) = IIf(IsNull(rsRecord!是否单独定价), "", rsRecord!是否单独定价)
                .TextMatrix(i, mVaricolumn.品种_急救药) = IIf(IsNull(rsRecord!急救药否), "", rsRecord!急救药否)
                .TextMatrix(i, mVaricolumn.品种_新药) = IIf(IsNull(rsRecord!是否新药), "", rsRecord!是否新药)
                .TextMatrix(i, mVaricolumn.品种_原料药) = IIf(IsNull(rsRecord!是否原料), "", rsRecord!是否原料)
                .TextMatrix(i, mVaricolumn.品种_单味使用) = IIf(IsNull(rsRecord!单味使用), "", rsRecord!单味使用)
                .TextMatrix(i, mVaricolumn.品种_辅助用药) = IIf(IsNull(rsRecord!辅助用药), "", rsRecord!辅助用药)
                .TextMatrix(i, mVaricolumn.品种_肿瘤药) = IIf(IsNull(rsRecord!肿瘤药), "", rsRecord!肿瘤药)
                .TextMatrix(i, mVaricolumn.品种_溶媒) = IIf(IsNull(rsRecord!溶媒), "", rsRecord!溶媒)
                .TextMatrix(i, mVaricolumn.品种_ATCCODE) = IIf(IsNull(rsRecord!ATCCODE), "", rsRecord!ATCCODE)
                .TextMatrix(i, mVaricolumn.品种_参考项目) = IIf(IsNull(rsRecord!参考项目), "", rsRecord!参考项目)
                .TextMatrix(i, mVaricolumn.品种_处方职务) = ShowValue(.ColComboList(mVaricolumn.品种_处方职务), IIf(IsNull(Mid(rsRecord!处方职务, 1, 1)), "", Mid(rsRecord!处方职务, 1, 1)), "-", True)
                .TextMatrix(i, mVaricolumn.品种_医保职务) = ShowValue(.ColComboList(mVaricolumn.品种_医保职务), IIf(IsNull(Mid(rsRecord!处方职务, 2, 1)), "", Mid(rsRecord!处方职务, 2, 1)), "-", True)
                .TextMatrix(i, mVaricolumn.品种_处方限量) = IIf(IsNull(rsRecord!处方限量), "", rsRecord!处方限量)
                .TextMatrix(i, mVaricolumn.品种_适用性别) = ShowValue(.ColComboList(mVaricolumn.品种_适用性别), IIf(IsNull(rsRecord!适用性别), "0", rsRecord!适用性别), "-", True)
                .TextMatrix(i, mVaricolumn.品种_剂量单位) = IIf(IsNull(rsRecord!剂量单位), "", rsRecord!剂量单位)
                .TextMatrix(i, mVaricolumn.品种_皮试) = IIf(IsNull(rsRecord!是否皮试), "", rsRecord!是否皮试)
                .TextMatrix(i, mVaricolumn.品种_抗生素) = ShowValue(.ColComboList(mVaricolumn.品种_抗生素), IIf(IsNull(rsRecord!抗生素), "", rsRecord!抗生素), "-", True)
                .TextMatrix(i, mVaricolumn.品种_品种下长期医嘱) = IIf(IsNull(rsRecord!品种医嘱), "0", rsRecord!品种医嘱)
                .TextMatrix(i, mVaricolumn.品种_参考项目ID) = IIf(IsNull(rsRecord!参考目录ID), "", rsRecord!参考目录ID)
                
                If rsRecord!抗生素 = 0 Then
                    .Cell(flexcpBackColor, i, mVaricolumn.品种_ATCCODE, i) = mlngColor
                End If
            End With
'            Call ShowPercent(rsRecord.AbsolutePosition / rsRecord.RecordCount) '进度条
            
            rsRecord.MoveNext
        Next
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.品种_药品编码, vsfDetails.Rows - 1) = mlngColor    '设置不可编辑列的背景颜色为灰色
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.品种_药品分类, vsfDetails.Rows - 1) = mlngColor     '设置不可编辑列的背景颜色为灰色
        
'        With vsfDetails
'            If str分类 Like "中草药*" Then
'                .ColHidden(mVaricolumn.品种_单味使用) = False
'                .ColHidden(mVaricolumn.品种_剂型) = True
'                .ColHidden(mVaricolumn.品种_急救药) = True
'                .ColHidden(mVaricolumn.品种_新药) = True
'                .ColHidden(mVaricolumn.品种_皮试) = True
'                .ColHidden(mVaricolumn.品种_抗生素) = True
'                .ColHidden(mVaricolumn.品种_品种下长期医嘱) = True
'            Else
'                .ColHidden(mVaricolumn.品种_单味使用) = True
'                .ColHidden(mVaricolumn.品种_剂型) = False
'                .ColHidden(mVaricolumn.品种_急救药) = False
'                .ColHidden(mVaricolumn.品种_新药) = False
'                .ColHidden(mVaricolumn.品种_皮试) = False
'                .ColHidden(mVaricolumn.品种_抗生素) = False
'                .ColHidden(mVaricolumn.品种_品种下长期医嘱) = False
'            End If
'        End With
        
        vsfDetails.MergeCol(mVaricolumn.品种_药品分类) = True  '相同列中 药品分类相同合并
    Else    '规格
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mSpecColumn.规格_序号) = i
                .TextMatrix(i, mSpecColumn.规格_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mSpecColumn.规格_药名id) = IIf(IsNull(rsRecord!药名ID), "", rsRecord!药名ID)
'                .TextMatrix(i, mSpecColumn.规格_药品编码) = IIf(IsNull(rsRecord!品种编码), "", rsRecord!品种编码)
                .TextMatrix(i, mSpecColumn.规格_通用名称) = IIf(IsNull(rsRecord!通用名称), "", rsRecord!通用名称)
                .TextMatrix(i, mSpecColumn.规格_规格编码) = IIf(IsNull(rsRecord!规格编码), "", rsRecord!规格编码)
                .TextMatrix(i, mSpecColumn.规格_药品规格) = IIf(IsNull(rsRecord!规格), "", rsRecord!规格)
                .TextMatrix(i, mSpecColumn.规格_本位码) = IIf(IsNull(rsRecord!本位码), "", rsRecord!本位码)
                .TextMatrix(i, mSpecColumn.规格_数字码) = IIf(IsNull(rsRecord!数字码), "", rsRecord!数字码)
                
                If .TextMatrix(i, mSpecColumn.规格_数字码) = "" And .TextMatrix(i, mSpecColumn.规格_药品规格) <> "" Then
                    .TextMatrix(i, mSpecColumn.规格_数字码) = zlGetDigitSign(rsRecord!药名ID, rsRecord!规格)
                End If
                
                .TextMatrix(i, mSpecColumn.规格_标识码) = IIf(IsNull(rsRecord!标识码), "", rsRecord!标识码)
                .TextMatrix(i, mSpecColumn.规格_备选码) = IIf(IsNull(rsRecord!备选码), "", rsRecord!备选码)
                .TextMatrix(i, mSpecColumn.规格_容量) = IIf(IsNull(rsRecord!容量), "", rsRecord!容量)
                .TextMatrix(i, mSpecColumn.规格_商品名称) = IIf(IsNull(rsRecord!商品名), "", rsRecord!商品名)
                .TextMatrix(i, mSpecColumn.规格_生产厂牌) = IIf(IsNull(rsRecord!生产厂牌), "", rsRecord!生产厂牌)
                .TextMatrix(i, mSpecColumn.规格_来源分类) = ShowValue(.ColComboList(mSpecColumn.规格_来源分类), IIf(IsNull(rsRecord!来源分类), "", rsRecord!来源分类), "-")
                .TextMatrix(i, mSpecColumn.规格_拼音码) = IIf(IsNull(rsRecord!拼音码), "", rsRecord!拼音码)
                .TextMatrix(i, mSpecColumn.规格_五笔码) = IIf(IsNull(rsRecord!五笔码), "", rsRecord!五笔码)
                
                If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" And .TextMatrix(i, mSpecColumn.规格_拼音码) = "" Then
                    .TextMatrix(i, mSpecColumn.规格_拼音码) = zlGetSymbol(.TextMatrix(i, mSpecColumn.规格_通用名称), 0, 30)
                End If
                
                If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" And .TextMatrix(i, mSpecColumn.规格_拼音码) = "" Then
                    .TextMatrix(i, mSpecColumn.规格_五笔码) = zlGetSymbol(.TextMatrix(i, mSpecColumn.规格_通用名称), 1, 30)
                End If
                
                .TextMatrix(i, mSpecColumn.规格_合同单位) = IIf(IsNull(rsRecord!合同单位), "", rsRecord!合同单位)
                .TextMatrix(i, mSpecColumn.规格_批准文号) = IIf(IsNull(rsRecord!批准文号), "", rsRecord!批准文号)
                
                .TextMatrix(i, mSpecColumn.规格_注册商标) = IIf(IsNull(rsRecord!注册商标), "", rsRecord!注册商标)
                .TextMatrix(i, mSpecColumn.规格_GMP认证) = IIf(IsNull(rsRecord!GMP认证), "", rsRecord!GMP认证)
                .TextMatrix(i, mSpecColumn.规格_非常备药) = IIf(IsNull(rsRecord!是否常备), "", rsRecord!是否常备)
                .TextMatrix(i, mSpecColumn.规格_售价单位) = IIf(IsNull(rsRecord!售价单位), "", rsRecord!售价单位)
                .TextMatrix(i, mSpecColumn.规格_剂量系数) = IIf(IsNull(rsRecord!售价系数), "", rsRecord!售价系数)
                .TextMatrix(i, mSpecColumn.规格_剂量单位) = IIf(IsNull(rsRecord!计算单位), "", rsRecord!计算单位)
                .TextMatrix(i, mSpecColumn.规格_住院单位) = IIf(IsNull(rsRecord!住院单位), "", rsRecord!住院单位)
                .TextMatrix(i, mSpecColumn.规格_住院系数) = IIf(IsNull(rsRecord!住院包装), "", rsRecord!住院包装)
                .TextMatrix(i, mSpecColumn.规格_门诊单位) = IIf(IsNull(rsRecord!门诊单位), "", rsRecord!门诊单位)
                .TextMatrix(i, mSpecColumn.规格_门诊系数) = IIf(IsNull(rsRecord!门诊包装), "", rsRecord!门诊包装)
                .TextMatrix(i, mSpecColumn.规格_药库单位) = IIf(IsNull(rsRecord!药库单位), "", rsRecord!药库单位)
                
                .TextMatrix(i, mSpecColumn.规格_药价属性) = ShowValue(.ColComboList(mSpecColumn.规格_药价属性), IIf(IsNull(rsRecord!药价属性), "", rsRecord!药价属性), "-", True)
                .TextMatrix(i, mSpecColumn.规格_药库系数) = IIf(IsNull(rsRecord!药库包装), "", rsRecord!药库包装)
                .TextMatrix(i, mSpecColumn.规格_送货单位) = IIf(IsNull(rsRecord!送货单位), "", rsRecord!送货单位)
                .TextMatrix(i, mSpecColumn.规格_送货包装) = IIf(IsNull(rsRecord!送货包装), "", rsRecord!送货包装)
                Select Case rsRecord!中药形态
                    Case "0"
                        strTemp = "散装"
                    Case "1"
                        strTemp = "中药饮片"
                    Case Else
                        strTemp = "免煎剂"
                End Select
                
                .TextMatrix(i, mSpecColumn.规格_中药形态) = strTemp
                
                Select Case rsRecord!申领单位
                    Case "1"
                        strTemp = "售价单位"
                    Case "2"
                        strTemp = "住院单位"
                    Case "3"
                        strTemp = "门诊单位"
                    Case "4"
                        strTemp = "药库单位"
                    Case Else
                        strTemp = "售价单位"
                End Select
                .TextMatrix(i, mSpecColumn.规格_申领单位) = strTemp
                
                Select Case Nvl(rsRecord!申领单位, 1)
                    Case 1 '零售
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0), "#0.00;-#0.00; ;")
                    Case 2 '住院
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!住院包装, 1), "#0.00;-#0.00; ;")
                    Case 3 '门诊
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!门诊包装, 1), "#0.00;-#0.00; ;")
                    Case 4 '药库
                        .TextMatrix(i, mSpecColumn.规格_申领阀值) = Format(Nvl(rsRecord!申领阀值, 0) / Nvl(rsRecord!药库包装, 1), "#0.00;-#0.00; ;")
                End Select
                
                If mint当前单位 <> 0 Then
                    .TextMatrix(i, mSpecColumn.规格_采购限价) = FormatEx(IIf(IsNull(rsRecord!采购限价), 0, rsRecord!采购限价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.规格_指导售价) = FormatEx(IIf(IsNull(rsRecord!指导售价), 0, rsRecord!指导售价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.规格_成本价格) = FormatEx(IIf(IsNull(rsRecord!成本价), "", rsRecord!成本价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintCostDigit)
                Else
                    .TextMatrix(i, mSpecColumn.规格_采购限价) = FormatEx(IIf(IsNull(rsRecord!采购限价), 0, rsRecord!采购限价), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.规格_指导售价) = FormatEx(IIf(IsNull(rsRecord!指导售价), 0, rsRecord!指导售价), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.规格_成本价格) = FormatEx(IIf(IsNull(rsRecord!成本价), "", rsRecord!成本价), mintCostDigit)
                End If
                
                .TextMatrix(i, mSpecColumn.规格_采购扣率) = IIf(IsNull(rsRecord!采购扣率), "", rsRecord!采购扣率)
                .TextMatrix(i, mSpecColumn.规格_结算价) = FormatEx(.TextMatrix(i, mSpecColumn.规格_采购限价) * (.TextMatrix(i, mSpecColumn.规格_采购扣率) / 100), mintCostDigit)
                .TextMatrix(i, mSpecColumn.规格_指导差率) = Format(IIf(IsNull(rsRecord!指导差率), "", rsRecord!指导差率), "0.00000")
                .TextMatrix(i, mSpecColumn.规格_加成率) = Format((1 / (1 - .TextMatrix(i, mSpecColumn.规格_指导差率) / 100) - 1) * 100, "0.00")
                .TextMatrix(i, mSpecColumn.规格_差价让利) = Format(IIf(IsNull(rsRecord!差价让利), "", rsRecord!差价让利), "0.00")
                
                If mint当前单位 <> 0 Then
                    .TextMatrix(i, mSpecColumn.规格_当前售价) = FormatEx(IIf(IsNull(rsRecord!当前售价), 0, rsRecord!当前售价) * .TextMatrix(i, mSpecColumn.规格_药库系数), mintPriceDigit)
                Else
                    .TextMatrix(i, mSpecColumn.规格_当前售价) = FormatEx(IIf(IsNull(rsRecord!当前售价), 0, rsRecord!当前售价), mintPriceDigit)
                End If
                .TextMatrix(i, mSpecColumn.规格_收入项目) = ShowValue(.ColComboList(mSpecColumn.规格_收入项目), rsRecord!收入项目, "]")
                .TextMatrix(i, mSpecColumn.规格_病案费目) = IIf(IsNull(rsRecord!病案费目), "", rsRecord!病案费目)
                .TextMatrix(i, mSpecColumn.规格_管理费比例) = Format(IIf(IsNull(rsRecord!管理费比例), "", rsRecord!管理费比例), "0.00")
                .TextMatrix(i, mSpecColumn.规格_药价级别) = ShowValue(.ColComboList(mSpecColumn.规格_药价级别), IIf(IsNull(rsRecord!药价级别), "", rsRecord!药价级别), "-")
                .TextMatrix(i, mSpecColumn.规格_屏蔽费别) = IIf(IsNull(rsRecord!屏蔽费别), "", rsRecord!屏蔽费别)
                .TextMatrix(i, mSpecColumn.规格_增值税率) = Format(IIf(IsNull(rsRecord!增值税率), "", rsRecord!增值税率), "0.00")
                .TextMatrix(i, mSpecColumn.规格_医保类型) = ShowValue(.ColComboList(mSpecColumn.规格_医保类型), IIf(IsNull(rsRecord!医保类型), "", rsRecord!医保类型), "-")
                .TextMatrix(i, mSpecColumn.规格_药库分批) = IIf(IsNull(rsRecord!药库分批), "", rsRecord!药库分批)
                .TextMatrix(i, mSpecColumn.规格_药房分批) = IIf(IsNull(rsRecord!药房分批), "", rsRecord!药房分批)
                .TextMatrix(i, mSpecColumn.规格_原药库分批) = IIf(IsNull(rsRecord!药库分批), "", rsRecord!药库分批)
                .TextMatrix(i, mSpecColumn.规格_原药房分批) = IIf(IsNull(rsRecord!药房分批), "", rsRecord!药房分批)
                
                .TextMatrix(i, mSpecColumn.规格_保质期) = FormatEx(IIf(Nvl(rsRecord!保质期, 0) = 0, 0, rsRecord!保质期), 5)
                .TextMatrix(i, mSpecColumn.规格_标识说明) = IIf(IsNull(rsRecord!标识说明), "", rsRecord!标识说明)
                .TextMatrix(i, mSpecColumn.规格_发药类型) = ShowValue(.ColComboList(mSpecColumn.规格_发药类型), IIf(IsNull(rsRecord!发药类型), "", rsRecord!发药类型))
                .TextMatrix(i, mSpecColumn.规格_站点编号) = ShowValue(.ColComboList(mSpecColumn.规格_站点编号), IIf(IsNull(rsRecord!站点编号), "", rsRecord!站点编号), "-", True)
                .TextMatrix(i, mSpecColumn.规格_DDD值) = IIf(IsNull(rsRecord!ddd值), "", rsRecord!ddd值)
                .TextMatrix(i, mSpecColumn.规格_服务对象) = ShowValue(.ColComboList(mSpecColumn.规格_服务对象), IIf(IsNull(rsRecord!服务对象), "", rsRecord!服务对象), "-", True)
                .TextMatrix(i, mSpecColumn.规格_高危药品) = ShowValue(.ColComboList(mSpecColumn.规格_高危药品), IIf(IsNull(rsRecord!高危药品), "", rsRecord!高危药品), "-", True)

                If str分类 Like "中草药*" Then
                    If IsNull(rsRecord!住院可否分零) Or rsRecord!住院可否分零 = 0 Then
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "0-可以分零"
                    Else
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "1-不可分零"
                    End If
                    If IsNull(rsRecord!门诊可否分零) Or rsRecord!门诊可否分零 = 0 Then
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "0-可以分零"
                    Else
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "1-不可分零"
                    End If
                    
                    If .TextMatrix(i, mSpecColumn.规格_中药形态) = "散装" Then
                        .TextMatrix(i, mSpecColumn.规格_住院分零使用) = "0-可以分零"
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_住院分零使用) = mlngColor
                        .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = "0-可以分零"
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_门诊分零使用) = mlngColor
                    Else
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_住院分零使用) = mlngApplyColor
                        .Cell(flexcpBackColor, i, mSpecColumn.规格_门诊分零使用) = mlngApplyColor
                    End If
                Else
                    If IsNull(rsRecord!住院可否分零) Or rsRecord!住院可否分零 = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!住院可否分零 = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!住院可否分零 = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!住院可否分零 = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!住院可否分零 = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!住院可否分零 = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.规格_住院分零使用) = ShowValue(.ColComboList(mSpecColumn.规格_住院分零使用), IIf(IsNull(rsRecord!门诊可否分零), "", intTemp), "-", True)
                    
                    If IsNull(rsRecord!门诊可否分零) Or rsRecord!门诊可否分零 = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!门诊可否分零 = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!门诊可否分零 = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!门诊可否分零 = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!门诊可否分零 = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!门诊可否分零 = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.规格_门诊分零使用) = ShowValue(.ColComboList(mSpecColumn.规格_门诊分零使用), IIf(IsNull(rsRecord!门诊可否分零), "", intTemp), "-", True)
                End If
                .TextMatrix(i, mSpecColumn.规格_基本药物) = ShowValue(.ColComboList(mSpecColumn.规格_基本药物), IIf(IsNull(rsRecord!基本药物), "", rsRecord!基本药物))
                .TextMatrix(i, mSpecColumn.规格_住院动态分零) = IIf(IsNull(rsRecord!住院动态分零), "", rsRecord!住院动态分零)
                .TextMatrix(i, mSpecColumn.规格_存储温度) = ShowValue(.ColComboList(mSpecColumn.规格_存储温度), IIf(IsNull(rsRecord!存储温度), "", rsRecord!存储温度), "-", True)
                .TextMatrix(i, mSpecColumn.规格_存储条件) = IIf(IsNull(rsRecord!存储条件), "", rsRecord!存储条件)
                .TextMatrix(i, mSpecColumn.规格_输液注意事项) = IIf(IsNull(rsRecord!输液注意事项), "", rsRecord!输液注意事项)
                .TextMatrix(i, mSpecColumn.规格_配药类型) = ShowValue(.ColComboList(mSpecColumn.规格_配药类型), IIf(IsNull(rsRecord!配药类型), "", rsRecord!配药类型))
                .TextMatrix(i, mSpecColumn.规格_不予调配) = IIf(IsNull(rsRecord!不予调配), "", rsRecord!不予调配)
                .TextMatrix(i, mSpecColumn.规格_招标药品) = IIf(IsNull(rsRecord!招标药品), 0, rsRecord!招标药品)
                .TextMatrix(i, mSpecColumn.规格_合同单位id) = IIf(IsNull(rsRecord!合同单位id), "", rsRecord!合同单位id)
                .TextMatrix(i, mSpecColumn.规格_收入项目id) = IIf(IsNull(rsRecord!收入项目id), "", rsRecord!收入项目id)
                
                Call CheckValue(i, rsRecord!ID)
            End With
            rsRecord.MoveNext
        Next
        vsfDetails.MergeCol(mSpecColumn.规格_通用名称) = True   '合并通用名称
        With vsfDetails
'            .Cell(flexcpBackColor, 1, mSpecColumn.规格_药品编码, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_规格编码, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_通用名称, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.规格_剂量单位, .Rows - 1) = mlngColor
            
'            If str分类 Like "中草药*" Then
'                .ColHidden(mSpecColumn.规格_商品名称) = True
'                .ColHidden(mSpecColumn.规格_拼音码) = True
'                .ColHidden(mSpecColumn.规格_五笔码) = True
'                .ColHidden(mSpecColumn.规格_GMP认证) = True
'                .ColHidden(mSpecColumn.规格_住院单位) = True
'                .ColHidden(mSpecColumn.规格_住院系数) = True
'                .ColHidden(mSpecColumn.规格_中药形态) = False
'                .ColHidden(mSpecColumn.规格_差价让利) = True
'                .ColHidden(mSpecColumn.规格_保质期) = True
'                .ColHidden(mSpecColumn.规格_基本药物) = True
'                .ColHidden(mSpecColumn.规格_住院动态分零) = True
'            Else
'                .ColHidden(mSpecColumn.规格_商品名称) = False
'                .ColHidden(mSpecColumn.规格_拼音码) = False
'                .ColHidden(mSpecColumn.规格_五笔码) = False
'                .ColHidden(mSpecColumn.规格_GMP认证) = False
'                .ColHidden(mSpecColumn.规格_住院单位) = False
'                .ColHidden(mSpecColumn.规格_住院系数) = False
'                .ColHidden(mSpecColumn.规格_中药形态) = True
'                .ColHidden(mSpecColumn.规格_差价让利) = False
'                .ColHidden(mSpecColumn.规格_保质期) = False
'                .ColHidden(mSpecColumn.规格_基本药物) = False
'                .ColHidden(mSpecColumn.规格_住院动态分零) = False
'            End If
        End With
    End If
    
    Call Recover    '将修改了的颜色改变回来
    
    '调整行高
    With vsfDetails
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub

Private Function ShowValue(ByVal strValue As String, ByVal strBiJiao As String, Optional str分解符 As String, Optional bln左匹配 As Boolean) As String
    '功能 ：通过传入的值比较返回所获取的值
    '参数 strvalue 原字符串
    'strBiJiao 需要比较的字符串
    'str分解符 单个字符串中的分隔符号，用于进行对比分解符右边或左边的字符串 "0-可分零"
    'bln左匹配=true 对比分解符左边的字符串/=false 对比分解符右边的字符串
    Dim arr As Variant
    Dim i As Integer

    If strValue = "" Then Exit Function
    ReDim arr(UBound(Split(strValue, "|"))) As String   '重新定义数组长度

    '将值分解开来保存到数组中
    For i = 0 To UBound(Split(strValue, "|"))
        arr(i) = Split(strValue, "|")(i)
    Next
    If strBiJiao = "" Then
        ShowValue = ""
        Exit Function
    End If

    '循环比较
    For i = 0 To UBound(Split(strValue, "|"))
        If Trim(arr(i)) <> "" Then
            If str分解符 = "" Then '分解符为空直接进行对比
                If arr(i) = strBiJiao Then
                    ShowValue = arr(i)
                    Exit Function
                End If
            Else
                If bln左匹配 Then '对比分解符左边的字符串
                    If Mid(arr(i), 1, InStr(1, arr(i), str分解符) - 1) = strBiJiao Then
                        ShowValue = arr(i)
                        Exit Function
                    End If
                Else                    '对比分解符右边的字符串
                    If Mid(arr(i), InStr(1, arr(i), str分解符) + 1) = strBiJiao Then
                        ShowValue = arr(i)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindGridRow(UCase(txtFind))
        txtFind.SetFocus
    End If
End Sub

Private Sub vsfDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If mint次数 = 2 Then
        With vsfDetails
            If mint状态 <> 1 Then   '规格
                Select Case Col
                    Case mSpecColumn.规格_药库分批
                        If .TextMatrix(Row, Col) <> "0" Then
                            .Cell(flexcpBackColor, Row, mSpecColumn.规格_药房分批) = cstcolor_backcolor
                            .Cell(flexcpBackColor, Row, mSpecColumn.规格_保质期) = cstcolor_backcolor
                            If .TextMatrix(Row, mSpecColumn.规格_保质期) = 0 Then
                                .TextMatrix(Row, mSpecColumn.规格_保质期) = 24
                                .Cell(flexcpForeColor, Row, mSpecColumn.规格_保质期) = mlngApplyColor
                                .Cell(flexcpFontBold, Row, mSpecColumn.规格_保质期) = 10
                            End If
                        Else
                            .Cell(flexcpBackColor, Row, mSpecColumn.规格_药房分批) = mlngColor
                            .Cell(flexcpBackColor, Row, mSpecColumn.规格_保质期) = mlngColor
                            .TextMatrix(Row, mSpecColumn.规格_药房分批) = 0
                            .TextMatrix(Row, mSpecColumn.规格_保质期) = 0
                            .Cell(flexcpForeColor, Row, mSpecColumn.规格_保质期) = mlngApplyColor
                            .Cell(flexcpFontBold, Row, mSpecColumn.规格_保质期) = 10
                        End If
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                    Case mSpecColumn.规格_屏蔽费别, mSpecColumn.规格_住院动态分零, mSpecColumn.规格_GMP认证, mSpecColumn.规格_非常备药, mSpecColumn.规格_药房分批, mSpecColumn.规格_存储条件, mSpecColumn.规格_不予调配
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                End Select
            Else    '品种
                Select Case Col
                    Case mVaricolumn.品种_原研药, mVaricolumn.品种_专利药, mVaricolumn.品种_单独定价, mVaricolumn.品种_急救药, mVaricolumn.品种_新药, mVaricolumn.品种_辅助用药, mVaricolumn.品种_原料药, mVaricolumn.品种_肿瘤药, mVaricolumn.品种_溶媒, mVaricolumn.品种_品种下长期医嘱, mVaricolumn.品种_皮试, mVaricolumn.品种_单味使用
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                End Select
            End If
        End With
    End If
End Sub

Private Sub vsfDetails_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    Dim j As Integer
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim intupdate As Integer
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strTemp As String

    On Error GoTo ErrHandle
    With vsfDetails
        If .Cell(flexcpBackColor, NewRow, NewCol) = mlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        
        If .Row < OldRow Then
            OldRow = 1
        End If
        
        If .Rows = 1 Then
           OldRow = 0
        End If
    End With
    
    '控制菜单中应用于所有列显示与否
    With vsfDetails
        If mint状态 = 1 Then '品种
            Select Case NewCol
                Case mVaricolumn.品种_通用名称
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.品种_英文名称
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.品种_拼音码
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.品种_五笔码
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Visible = False
                        mobjPopup.Controls(1).Visible = False
                    Else
                        mcbrToolBar.Controls(1).Visible = True
                        mobjPopup.Controls(1).Visible = True
                    End If
            End Select
            
            Select Case OldCol
                Case mVaricolumn.品种_通用名称, mVaricolumn.品种_剂量单位
                    If vsfDetails.TextMatrix(OldRow, OldCol) = "" Then
                        MsgBox "该单元格内容不能为空，请输入！", vbInformation, gstrSysName
                        vsfDetails.Select OldRow, OldCol
                    End If
            End Select
        Else    '规格
            Select Case NewCol
                Case mSpecColumn.规格_商品名称
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mSpecColumn.规格_拼音码
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mSpecColumn.规格_五笔码
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Visible = False
                        mobjPopup.Controls(1).Visible = False
                    Else
                        mcbrToolBar.Controls(1).Visible = True
                        mobjPopup.Controls(1).Visible = True
                    End If
            End Select
            
            Select Case OldCol
                Case mSpecColumn.规格_指导差率
                    If Val(.TextMatrix(OldRow, OldCol)) < 100 Then
                       .TextMatrix(OldRow, mSpecColumn.规格_加成率) = Format((1 / (1 - Val(.TextMatrix(OldRow, OldCol)) / 100) - 1) * 100, "0.00")
                       
                       If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_加成率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_加成率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_加成率) = True
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_指导差率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_指导差率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_指导差率) = True
                       End If
                    Else
                        '不允许出现指导差价率大于等于100的情况，因此需要从加成率反算回来
                        .TextMatrix(OldRow, OldCol) = Format((1 - (1 / (1 + Val(.TextMatrix(OldRow, mSpecColumn.规格_加成率)) / 100))) * 100, "0.00000")
                        If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_指导差率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_指导差率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_指导差率) = True
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_加成率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_加成率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_加成率) = True
                        End If
                    End If
                Case mSpecColumn.规格_加成率
                    If .TextMatrix(OldRow, OldCol) <> "" Then
                       .TextMatrix(OldRow, mSpecColumn.规格_指导差率) = Format((1 - (1 / (1 + Val(.TextMatrix(OldRow, mSpecColumn.规格_加成率)) / 100))) * 100, "0.00000")
                    End If
                    If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                        .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_指导差率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_指导差率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_指导差率) = True
                        .Cell(flexcpForeColor, OldRow, mSpecColumn.规格_加成率) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.规格_加成率) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.规格_加成率) = True
                    End If
                Case mSpecColumn.规格_剂量系数, mSpecColumn.规格_住院系数, mSpecColumn.规格_门诊系数, mSpecColumn.规格_药库系数, mSpecColumn.规格_送货包装, mSpecColumn.规格_申领阀值, mSpecColumn.规格_采购限价, mSpecColumn.规格_采购扣率, mSpecColumn.规格_结算价, mSpecColumn.规格_指导售价, mSpecColumn.规格_指导差率, mSpecColumn.规格_加成率, mSpecColumn.规格_剂量系数, mSpecColumn.规格_差价让利, mSpecColumn.规格_成本价格, mSpecColumn.规格_当前售价, mSpecColumn.规格_管理费比例, mSpecColumn.规格_增值税率
                    If .TextMatrix(OldRow, OldCol) <> "" Then
                        If Mid(.TextMatrix(OldRow, OldCol), 1, 1) = "." Then
                            .TextMatrix(OldRow, OldCol) = "0" & .TextMatrix(OldRow, OldCol)
                        End If
                        
                        If Mid(.TextMatrix(OldRow, OldCol), Len(.TextMatrix(OldRow, OldCol)), 1) = "." Then
                            .TextMatrix(OldRow, OldCol) = Mid(.TextMatrix(OldRow, OldCol), 1, Len(.TextMatrix(OldRow, OldCol)) - 1)
                        End If
                    End If
                Case mSpecColumn.规格_收入项目
                    If OldRow <> 0 Then
                        If .TextMatrix(OldRow, OldCol) <> "" Then
                            strTemp = Mid(.TextMatrix(OldRow, OldCol), 2, InStr(1, .TextMatrix(OldRow, OldCol), "]") - 2)
                        End If
                        gstrSql = "Select ID" & _
                                  "  From 收入项目" & _
                                  "  Where 编码=[1] and 末级 = 1 And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "收入项目查询", strTemp)
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(OldRow, mSpecColumn.规格_收入项目id) = rsTmp!ID
                        End If
                    End If
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '控制哪些列可以编辑，那些列不可以编辑,当背景颜色为灰色的列都不允许修改
    With vsfDetails
        mstrOldValue = vsfDetails.TextMatrix(Row, Col)
        If .Cell(flexcpBackColor, Row, Col) = mlngColor Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfDetails_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = GetControlRect(vsfDetails.hWnd) '获取位置
    dblLeft = vRect.Left + vsfDetails.CellLeft
    dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
    On Error GoTo ErrHandle
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If Col = mVaricolumn.品种_参考项目 Then
                strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.品种_分类id))
                
                If rsTmp.EOF Then
                    intAttr = -1
                Else
                    intAttr = rsTmp!类型
                End If
                
                strSql = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=[1] Order By 编码"
                
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, intAttr)

                If rsRecord Is Nothing Then
                    Exit Sub
                End If
                .TextMatrix(.Row, mVaricolumn.品种_参考项目) = rsRecord!名称
                .TextMatrix(.Row, mVaricolumn.品种_参考项目ID) = rsRecord!ID
            End If
        Else    '规格
            Select Case Col
                Case mSpecColumn.规格_生产厂牌
                    strSql = "Select 编码 as id,名称,简码 From 药品生产商 Order By 编码 "
                    
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)
                    
                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.规格_生产厂牌) = rsRecord!名称
                    End If
                Case mSpecColumn.规格_合同单位
                    strSql = "Select id,编码,名称,简码" & _
                                " From 供应商" & _
                                " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                " Order By 编码 "
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)
                    
                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.规格_合同单位) = rsRecord!名称
                        .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = rsRecord!ID
                    End If
                Case mSpecColumn.规格_病案费目
                    Dim blnRe As Boolean
                    Dim str名称 As String
                    Dim strID As String
                    
                    gstrSql = "Select 编码 as id,上级 as 上级id, 名称, 简码, 末级 From 病案费目 Start With 上级 Is Null Connect By Prior 编码 = 上级"
                    blnRe = frmTreeLeafSel.ShowTree(gstrSql, strID, str名称, "病案费目")
                    '成功返回
                    If blnRe Then
                        .TextMatrix(.Row, mSpecColumn.规格_病案费目) = str名称
                    End If
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfDetails
        If mstrOldValue <> .TextMatrix(Row, Col) And .CellBackColor <> mlngColor And mblnClick = True And Row = mintRow And .Rows <> 1 Then
            .Cell(flexcpForeColor, Row, Col) = mlngApplyColor: .Cell(flexcpFontSize, Row, Col) = 10: .Cell(flexcpFontBold, Row, Col) = True
        End If
    End With
End Sub

Private Sub vsfDetails_ChangeEdit()
    Dim lngId As Long
    Dim strTemp As String
    
    With vsfDetails
        If mint状态 = 1 Then '品种
            Select Case .Col
                Case mVaricolumn.品种_通用名称
                    .TextMatrix(.Row, mVaricolumn.品种_拼音码) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mVaricolumn.品种_五笔码) = zlGetSymbol(.EditText, 1, 30)
                    .Cell(flexcpForeColor, .Row, mVaricolumn.品种_拼音码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mVaricolumn.品种_拼音码) = 10: .Cell(flexcpFontBold, .Row, mVaricolumn.品种_拼音码) = True
                    .Cell(flexcpForeColor, .Row, mVaricolumn.品种_五笔码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mVaricolumn.品种_五笔码) = 10: .Cell(flexcpFontBold, .Row, mVaricolumn.品种_五笔码) = True
                 Case mVaricolumn.品种_抗生素
                    If Mid(.EditText, 1, InStr(1, .EditText, "-") - 1) = 0 Then
                        .Cell(flexcpBackColor, .Row, mVaricolumn.品种_ATCCODE, .Row) = mlngColor
                        .TextMatrix(.Row, mVaricolumn.品种_ATCCODE) = ""
                    Else
                        .Cell(flexcpBackColor, .Row, mVaricolumn.品种_ATCCODE, .Row) = cstcolor_backcolor
                    End If
            End Select
        Else    '规格
            Select Case .Col
                Case mSpecColumn.规格_商品名称
                    .TextMatrix(.Row, mSpecColumn.规格_拼音码) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mSpecColumn.规格_五笔码) = zlGetSymbol(.EditText, 1, 30)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_拼音码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_拼音码) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_拼音码) = True
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_五笔码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_五笔码) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_五笔码) = True
                Case mSpecColumn.规格_药品规格
                    lngId = .TextMatrix(.Row, mSpecColumn.规格_id)
                    .TextMatrix(.Row, mSpecColumn.规格_数字码) = zlGetDigitSign(lngId, .EditText)
                    If mstrOldValue <> .EditText Then
                        .Cell(flexcpForeColor, .Row, mSpecColumn.规格_数字码) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_数字码) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_数字码) = True
                    End If
                Case mSpecColumn.规格_成本价格
                    .TextMatrix(.Row, mSpecColumn.规格_结算价) = FormatEx(Val(.EditText) * (Val(.TextMatrix(.Row, mSpecColumn.规格_采购扣率)) / 100), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_结算价) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_结算价) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_结算价) = True
                    .TextMatrix(.Row, mSpecColumn.规格_采购限价) = FormatEx(Val(.EditText), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_采购限价) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_采购限价) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_采购限价) = True
                Case mSpecColumn.规格_采购限价
                    .TextMatrix(.Row, mSpecColumn.规格_结算价) = FormatEx(Val(.EditText) * (Val(.TextMatrix(.Row, mSpecColumn.规格_采购扣率)) / 100), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_结算价) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_结算价) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_结算价) = True
                Case mSpecColumn.规格_采购扣率
                    .TextMatrix(.Row, mSpecColumn.规格_结算价) = FormatEx(Val(.EditText) / 100 * Val(.TextMatrix(.Row, mSpecColumn.规格_采购限价)), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.规格_结算价) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.规格_结算价) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.规格_结算价) = True
                Case mSpecColumn.规格_住院分零使用
                    If Val(Mid(.EditText, 1, 1)) = 0 Then
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院动态分零) = mlngColor
                    Else
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院动态分零) = cstcolor_backcolor
                    End If
                Case mSpecColumn.规格_中药形态
                    If .EditText <> "散装" Then
                        .Cell(flexcpForeColor, .Row, mSpecColumn.规格_中药形态) = mlngApplyColor
                        .Cell(flexcpFontBold, .Row, mSpecColumn.规格_中药形态) = True
                        .Cell(flexcpFontSize, .Row, mSpecColumn.规格_中药形态) = 10
                        MsgBox "你修改了“中药形态”，系统将强制设定“临床应用”页中分零使用为“不可分零”！", vbInformation, gstrSysName
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院分零使用) = cstcolor_backcolor
                        .TextMatrix(.Row, mSpecColumn.规格_住院分零使用) = "1-不可分零"
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_门诊分零使用) = cstcolor_backcolor
                        .TextMatrix(.Row, mSpecColumn.规格_门诊分零使用) = "1-不可分零"
                    Else
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_住院分零使用) = mlngColor
                        .Cell(flexcpBackColor, .Row, mSpecColumn.规格_门诊分零使用) = mlngColor
                    End If
            End Select
        End If
    End With
End Sub

Private Sub vsfDetails_Click()
    mblnClick = True
End Sub

Private Sub vsfDetails_EnterCell()
    Dim rsRecord As ADODB.Recordset
    Dim strKey As String
    Dim i As Integer
    Dim j As Integer
    
    If mintRow上次 > vsfDetails.Rows - 1 Then
        mintRow上次 = 1
    End If
    
    If vsfDetails.Rows <> 1 Then
        vsfDetails.Cell(flexcpPicture, mintRow上次, 0, mintRow上次, 0) = Nothing    '设置图片
        For i = 1 To vsfDetails.Rows - 1    '当在切换选项页+排序时会出现多个图片 在这种情况下先将多余的一个清除掉
            If Not vsfDetails.Cell(flexcpPicture, i, 0, i, 0) Is Nothing Then
                vsfDetails.Cell(flexcpPicture, i, 0, i, 0) = Nothing
                Exit For
            End If
        Next
        vsfDetails.Cell(flexcpPicture, vsfDetails.Row, 0, vsfDetails.Row, 0) = Me.ImgTvw.ListImages(2).Picture
    End If
    
    With vsfDetails
        If .Row = mintRow Then Exit Sub
        mintRow = .Row '记录当前行
        strKey = .TextMatrix(.Row, mVaricolumn.品种_id)
    End With
End Sub

Private Sub vsfDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call MoveRowCol
    End If
End Sub

Private Sub vsfDetails_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strSql As String, strSQLItem As String
    Dim rsRecord As ADODB.Recordset
    Dim iAttr As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intAllCol As Integer
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If vsfDetails.EditText = "" Then
            Call MoveRowCol
            Exit Sub
        End If
        
        If mint状态 = 1 Then '品种
            vRect = GetControlRect(vsfDetails.hWnd) '获取位置
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .Col = mVaricolumn.品种_参考项目 Then
                    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.品种_分类id))
                    
                    If rsRecord.EOF Then
                        iAttr = -1
                    Else
                        iAttr = rsRecord(0)
                    End If
                    If .EditText = "" Then
                        strSql = " Select ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 编码"
                    Else
                        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
                            " Where A.ID=B.参考目录ID And A.类型=[1]" & _
                            " And (Upper(A.编码) Like [2] " & _
                            " Or Upper(A.名称) Like [3] " & _
                            " Or Upper(B.名称) Like [3] " & _
                            " Or Upper(B.简码) Like [3] " & ")"
                
                        strSql = " Select DISTINCT A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, iAttr, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")
        
                        If rsRecord Is Nothing Then
                            .EditText = ""
                            .TextMatrix(.Row, mVaricolumn.品种_参考项目) = ""
                            .TextMatrix(.Row, mVaricolumn.品种_参考项目ID) = ""
                            Exit Sub
                        End If
                        .EditText = rsRecord!名称
                        .TextMatrix(.Row, mVaricolumn.品种_参考项目) = rsRecord!名称
                        .TextMatrix(.Row, mVaricolumn.品种_参考项目ID) = rsRecord!ID
                        End If
                End If
            End With
        Else    '规格
            Dim str As String
            vRect = GetControlRect(vsfDetails.hWnd) '获取位置
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .EditText = "" Then Exit Sub
                Select Case Col
                    Case mSpecColumn.规格_生产厂牌
                        str = UCase(.EditText)
                        If .Col = mSpecColumn.规格_生产厂牌 Then
                            strSql = "Select 编码 as id,名称,简码" & _
                                        " From 药品生产商" & _
                                        " where 编码 Like [1] " & _
                                        "       Or 名称 Like [2] " & _
                                        "       Or 简码 Like [2] Order By 编码 "
                            Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, str & "%", mstrMatch & str & "%")
                            If rsRecord Is Nothing Then
                                .EditText = ""
                                Exit Sub
                            Else
                                .EditText = rsRecord!名称
                                .TextMatrix(.Row, mSpecColumn.规格_生产厂牌) = rsRecord!名称
                            End If
                        End If
                    Case mSpecColumn.规格_合同单位
                        strSql = "Select 编码,名称,简码,id" & _
                                    " From 供应商" & _
                                    " where (编码 Like [1] " & _
                                    "       Or 名称 Like [2] " & _
                                    "       Or 简码 Like [2])" & _
                                    " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                    " Order By 编码 "
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "诊疗项目", False, "", "", False, False, _
                            True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")
                        
                        If rsRecord Is Nothing Then
                            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位) = ""
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = ""
                            Exit Sub
                        Else
                            .EditText = rsRecord!名称
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位) = rsRecord!名称
                            .TextMatrix(.Row, mSpecColumn.规格_合同单位id) = rsRecord!ID
                        End If
                End Select
            End With
        End If
        
        Call MoveRowCol
    End If
    
    If KeyAscii <> vbKeyBack Then
        With vsfDetails
            If mint状态 = 1 Then    '品种
                Select Case Col
                    Case mVaricolumn.品种_通用名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_通用名称)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_英文名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_英文名称)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_拼音码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_拼音码)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_五笔码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_五笔码)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_处方限量
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_处方限量)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mVaricolumn.品种_剂量单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_处方限量)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.品种_ATCCODE
                        If KeyAscii <> vbKeyDelete Then
                            If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.品种_ATCCODE)) Then
                                KeyAscii = 0
                            Else
                                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                            End If
                        End If
                End Select
            Else    '规格
                Select Case Col
                    Case mSpecColumn.规格_药品规格
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药品规格)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_本位码
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 20 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_数字码
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 7 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_标识码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_标识码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_备选码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_备选码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_容量
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_商品名称
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_商品名称)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_生产厂牌
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_生产厂牌)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_拼音码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_拼音码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_五笔码
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_五笔码)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_合同单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_合同单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_批准文号
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_批准文号)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_注册商标
                    
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_注册商标)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_售价单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_售价单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_剂量系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_剂量系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_住院单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= mintLen Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_住院系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_住院系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_门诊单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_门诊单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_门诊系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_门诊系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_药库单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药库单位)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_药库系数
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_药库系数)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_送货单位
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_送货单位)) Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_送货包装
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_送货包装)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_申领阀值
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_申领阀值)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_采购限价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_采购限价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_采购扣率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_采购扣率)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_指导售价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_指导售价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_指导差率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_指导差率)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_加成率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 19 Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_差价让利
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_差价让利)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_成本价格
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_成本价格)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_当前售价
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_当前售价)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_管理费比例
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_管理费比例)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_增值税率
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_增值税率)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_保质期
                        If KeyAscii = vbKeyDelete Then
                            KeyAscii = 0
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_保质期)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.规格_标识说明
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_标识说明)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.规格_输液注意事项
                      If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.规格_输液注意事项)) Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                          KeyAscii = 0
                      End If
                End Select
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_LeaveCell()
    mintRow上次 = vsfDetails.Row
    mintCol上次 = vsfDetails.Col
End Sub


Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mobjPopup.ShowPopup
    End If
End Sub

Private Sub Set权限判断()
'权限判断过程
    With vsfDetails
        If .Rows > 1 Then
            If mint状态 = 1 Then    '品种
                If InStr(1, mstrPrivs, "医保用药目录") = 0 Then
                    .Cell(flexcpBackColor, 1, mVaricolumn.品种_医保职务, .Rows - 1, mVaricolumn.品种_医保职务) = mlngColor
                End If
            Else    '规格
                If InStr(1, mstrPrivs, "医保用药目录") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_医保类型, .Rows - 1, mSpecColumn.规格_医保类型) = mlngColor
                End If
                If InStr(1, mstrPrivs, "管理扣率") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_采购扣率, .Rows - 1, mSpecColumn.规格_采购扣率) = mlngColor
                End If
                If InStr(1, mstrPrivs, "指导价格管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_指导差率, .Rows - 1, mSpecColumn.规格_指导差率) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_加成率, .Rows - 1, mSpecColumn.规格_加成率) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_采购限价, .Rows - 1, mSpecColumn.规格_采购限价) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_指导售价, .Rows - 1, mSpecColumn.规格_指导售价) = mlngColor
                End If
                If InStr(1, mstrPrivs, "售价管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药价属性, .Rows - 1, mSpecColumn.规格_药价属性) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_收入项目, .Rows - 1, mSpecColumn.规格_收入项目) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_差价让利, .Rows - 1, mSpecColumn.规格_差价让利) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_当前售价, .Rows - 1, mSpecColumn.规格_当前售价) = mlngColor
                End If
                If InStr(1, mstrPrivs, "药价级别") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药价级别, .Rows - 1, mSpecColumn.规格_药价级别) = mlngColor
                End If
                If InStr(1, mstrPrivs, "成本价管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_成本价格, .Rows - 1, mSpecColumn.规格_成本价格) = mlngColor
                End If
                If InStr(1, mstrPrivs, "调整服务对象") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_服务对象, .Rows - 1, mSpecColumn.规格_服务对象) = mlngColor
                End If
                If InStr(1, mstrPrivs, "药品单位管理") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_售价单位, .Rows - 1, mSpecColumn.规格_售价单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_住院单位, .Rows - 1, mSpecColumn.规格_住院单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_门诊单位, .Rows - 1, mSpecColumn.规格_门诊单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药库单位, .Rows - 1, mSpecColumn.规格_药库单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_剂量系数, .Rows - 1, mSpecColumn.规格_剂量系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_住院系数, .Rows - 1, mSpecColumn.规格_住院系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_门诊系数, .Rows - 1, mSpecColumn.规格_门诊系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_药库系数, .Rows - 1, mSpecColumn.规格_药库系数) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_送货单位, .Rows - 1, mSpecColumn.规格_送货单位) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.规格_送货包装, .Rows - 1, mSpecColumn.规格_送货包装) = mlngColor
                End If
                
                If mstrNode Like "中草药*" Then
                    If InStr(1, mstrPrivs, "草药分包管理") = 0 Then
                        .Cell(flexcpBackColor, 1, mSpecColumn.规格_中药形态, .Rows - 1, mSpecColumn.规格_中药形态) = mlngColor
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub Save()
    '数据保存方法
    Dim i As Integer
    Dim strTemp As String
    Dim j As Integer
    Dim m As Integer
    Dim n As Integer
    Dim intupdate As Integer
    Dim rsRecord As ADODB.Recordset
    Dim str别名 As String
    Dim intCount As Integer
    Dim bln修改 As Boolean
    Dim lng保存 As Long
    Dim lngSave As Long
    Dim intTemp As Integer
    
    bln修改 = Check修改
    On Error GoTo ErrHandle
    If bln修改 = False Then '没有修改的话直接退出不进行保存
        Exit Sub
    End If
    
    If mintExit <> 2 Then
        lngSave = MsgBox("确定保存？", vbInformation + vbYesNo, gstrSysName)
        If lngSave = vbNo Then
            Exit Sub
        End If
        mintExit = 0
    End If
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If .TextMatrix(1, mVaricolumn.品种_id) = "" Then Exit Sub
            '检查数据的合法性
            If CheckData = False Then Exit Sub
            
            If mstrNode Like "中草药*" Then '中草药
                For i = 1 To .Rows - 1
                    gstrSql = ""
                    strTemp = ""
                    gstrSql = "Zl_草药品种_Update (" & .TextMatrix(i, mVaricolumn.品种_分类id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_药品编码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_通用名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_拼音码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_五笔码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_英文名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_剂量单位) + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_毒理分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_毒理分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_价值分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_价值分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_货源情况), InStr(1, .TextMatrix(i, mVaricolumn.品种_货源情况), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_用药梯次), InStr(1, .TextMatrix(i, mVaricolumn.品种_用药梯次), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_药品类型), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_处方职务), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.品种_医保职务), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_处方限量) & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_单味使用) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原料药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_适用性别), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_参考项目) = "", "Null", IIf(.TextMatrix(i, mVaricolumn.品种_参考项目ID) = "", "Null", .TextMatrix(i, mVaricolumn.品种_参考项目ID)))
                    gstrSql = gstrSql + strTemp & ","
                    
                    str别名 = "Select distinct n.名称 as 药品名称, p.简码 As 拼音, w.简码 As 五笔" & _
                              "  From (Select Distinct 诊疗项目id,名称 From 诊疗项目别名 Where  性质 = 9) N," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 1) P," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 2) W" & _
                               " Where n.名称 = p.名称(+) And n.名称 = w.名称(+) and n.诊疗项目id = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str别名, "品种保存", .TextMatrix(i, mVaricolumn.品种_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!药品名称 & "^" & rsRecord!拼音 & "^" & rsRecord!五笔
                            rsRecord.MoveNext
                        Loop
                    End If

                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" + strTemp + "'" & ")"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp
                    End If

                    strTemp = ",NULL," & IIf(.TextMatrix(i, mVaricolumn.品种_辅助用药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "保存"
                Next
                '其他别名
            Else    '西成药、中成药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    strTemp = ""
                    
                    gstrSql = "Zl_成药品种_Update (" & .TextMatrix(i, mVaricolumn.品种_分类id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_药品编码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_通用名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_拼音码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_五笔码) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_英文名称) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.品种_剂量单位) + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_剂型), InStr(1, .TextMatrix(i, mVaricolumn.品种_剂型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_毒理分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_毒理分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_价值分类), InStr(1, .TextMatrix(i, mVaricolumn.品种_价值分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_货源情况), InStr(1, .TextMatrix(i, mVaricolumn.品种_货源情况), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_用药梯次), InStr(1, .TextMatrix(i, mVaricolumn.品种_用药梯次), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_药品类型), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_处方职务), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.品种_医保职务), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.品种_处方限量) & ","
                    
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_急救药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_新药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原料药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_皮试) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_抗生素), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    '参考目录id
                    '''''''''''''''''''''
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_参考项目) = "", "Null", IIf(.TextMatrix(i, mVaricolumn.品种_参考项目ID) = "", "Null", .TextMatrix(i, mVaricolumn.品种_参考项目ID)))
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_品种下长期医嘱) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.品种_适用性别), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    '别名
                    str别名 = "Select distinct n.名称 as 药品名称, p.简码 As 拼音, w.简码 As 五笔" & _
                              "  From (Select Distinct 诊疗项目id,名称 From 诊疗项目别名 Where  性质 = 9) N," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 1) P," & _
                                    " (Select 名称, 简码 From 诊疗项目别名 Where  性质 = 9 And 码类 = 2) W" & _
                               " Where n.名称 = p.名称(+) And n.名称 = w.名称(+) and n.诊疗项目id = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str别名, "品种保存", .TextMatrix(i, mVaricolumn.品种_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!药品名称 & "^" & rsRecord!拼音 & "^" & rsRecord!五笔
                            rsRecord.MoveNext
                        Loop
                    End If
                    
                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" & strTemp & "',"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp & ","
                    End If
                    gstrSql = gstrSql + "Null,"
                    gstrSql = gstrSql + "'" & .TextMatrix(i, mVaricolumn.品种_ATCCODE) & "',"
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_肿瘤药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_溶媒) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_原研药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_专利药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_单独定价) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.品种_辅助用药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "保存"
                Next
            End If
        Else    '规格
            If .TextMatrix(1, mSpecColumn.规格_id) = "" Then Exit Sub
            '检查数据的合法性
            If CheckData = False Then Exit Sub
            
            For i = 1 To vsfDetails.Rows - 1
                If .TextMatrix(i, mSpecColumn.规格_药品规格) = "" Then
                    MsgBox "第" & i & "行药品规格为空，请收入药品规格！", vbExclamation, gstrSysName
                    Exit Sub
                End If
            Next
            
            If mstrNode Like "中草药*" Then '中草药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = "zl_草药规格_Update(" & .TextMatrix(i, mSpecColumn.规格_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药品规格) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_生产厂牌) & "',"
                    
                    If .TextMatrix(i, mSpecColumn.规格_商品名称) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_商品名称) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = "Null"    '拼音码
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = "Null"    '五笔码
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_数字码) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_数字码) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_标识码) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_标识码) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_来源分类) <> "" Then
                        strTemp = "'" & Mid(.TextMatrix(i, mSpecColumn.规格_来源分类), InStr(1, .TextMatrix(i, mSpecColumn.规格_来源分类), "-") + 1) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp + ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_批准文号) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_批准文号) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_注册商标) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_注册商标) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_售价单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_剂量系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_门诊单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_门诊系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药库单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_药库系数) & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.规格_申领单位)
                        Case "售价单位"
                            strTemp = 1
                        Case "住院单位"
                            strTemp = 2
                        Case "门诊单位"
                            strTemp = 3
                        Case "药库单位"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_申领阀值)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_申领阀值)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价属性), 1, 1)
                    gstrSql = gstrSql & strTemp & ","

                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) & ","
                    End If
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购扣率) & ","
                    
                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) & ","
                    End If
                    
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.规格_加成率)) & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_管理费比例) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_管理费比例)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                        
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价级别), InStr(1, .TextMatrix(i, mSpecColumn.规格_药价级别), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_医保类型), InStr(1, .TextMatrix(i, mSpecColumn.规格_医保类型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_服务对象), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_GMP认证) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_招标药品) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_屏蔽费别) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_住院分零使用), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药库分批) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药房分批) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_保质期) & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_差价让利) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_差价让利)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) & ","
                    End If
                    
                    
                    strTemp = .TextMatrix(i, mSpecColumn.规格_收入项目id)
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_合同单位id) = "", "Null", .TextMatrix(i, mSpecColumn.规格_合同单位id)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_标识说明) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_住院动态分零) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_发药类型) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_备选码) & "',"
                    
                    If .TextMatrix(i, mSpecColumn.规格_增值税率) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_增值税率), 1, 1)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_基本药物) <> "" Then
                        strTemp = "'" & Mid(.TextMatrix(i, mSpecColumn.规格_基本药物), 1, 1) & "'"
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.规格_中药形态)
                        Case "散装"
                            strTemp = 0
                        Case "中药饮片"
                            strTemp = 1
                        Case "免煎剂"
                            strTemp = 2
                    End Select
                    gstrSql = gstrSql + strTemp & ","
                        
                    If .TextMatrix(i, mSpecColumn.规格_站点编号) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_站点编号), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_非常备药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_病案费目) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_病案费目) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_门诊分零使用), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)))
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_本位码) & "'"
                    gstrSql = gstrSql + strTemp & ")"
                    
                                                            
                    zlDatabase.ExecuteProcedure gstrSql, "草药规格保存"
                Next
            Else    '西成药、中成药
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    gstrSql = "zl_成药规格_Update(" & .TextMatrix(i, mSpecColumn.规格_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药品规格) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_生产厂牌) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_商品名称) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_拼音码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_五笔码) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_数字码) & "',"
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_标识码)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_标识码) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_来源分类), InStr(1, .TextMatrix(i, mSpecColumn.规格_来源分类), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_批准文号) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_注册商标) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_售价单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_剂量系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_门诊单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_门诊系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_住院单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_住院系数) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_药库单位) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_药库系数) & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.规格_申领单位)
                        Case "售价单位"
                            strTemp = 1
                        Case "住院单位"
                            strTemp = 2
                        Case "门诊单位"
                            strTemp = 3
                        Case "药库单位"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_申领阀值)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.规格_申领阀值)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价属性), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购限价) & ","
                    End If
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_采购扣率) & ","
                    
                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_指导售价) & ","
                    End If
                    
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.规格_加成率)) & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_管理费比例) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.规格_管理费比例)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_药价级别), InStr(1, .TextMatrix(i, mSpecColumn.规格_药价级别), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_医保类型), InStr(1, .TextMatrix(i, mSpecColumn.规格_医保类型), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_服务对象), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_GMP认证) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_招标药品) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_屏蔽费别) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_住院分零使用) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_住院分零使用), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药库分批) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_药房分批) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.规格_保质期)) & ","
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_差价让利) = "", "null", .TextMatrix(i, mSpecColumn.规格_差价让利)) & ","
                    
                    If mint当前单位 <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) / Nvl(.TextMatrix(i, mSpecColumn.规格_药库系数), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_成本价格) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.规格_当前售价) & ","
                    End If
                    
                    strTemp = .TextMatrix(i, mSpecColumn.规格_收入项目id)
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.规格_合同单位id) = "", "Null", .TextMatrix(i, mSpecColumn.规格_合同单位id)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.规格_标识说明) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_住院动态分零) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_发药类型) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_发药类型) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_备选码) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_备选码) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_增值税率) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.规格_增值税率)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_基本药物) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_基本药物) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_站点编号) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_站点编号), 1, 1)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_非常备药) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_存储温度)) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_存储温度), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_存储条件) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_配药类型)) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_配药类型) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.规格_不予调配) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_容量)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_容量)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_病案费目) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_病案费目) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.规格_门诊分零使用) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.规格_门诊分零使用), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.规格_DDD值)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.规格_DDD值)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_高危药品)) = "", 0, Mid(Trim(.TextMatrix(i, mSpecColumn.规格_高危药品)), 1, 1))
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.规格_送货包装)))
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.规格_输液注意事项)) & "',"
                    
                    strTemp = "'" & .TextMatrix(i, mSpecColumn.规格_本位码) & "'"
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "规格保存"
                Next
            End If
        End If
    End With
    Call Recover    '保存后刷新界面
    Call tvwDetails_NodeClick(tvwDetails.SelectedItem)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Recover()
    '使窗体中改变的颜色或者字体还原
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
               If .Cell(flexcpBackColor, i, j) <> mlngColor Then
                    .Cell(flexcpBackColor, i, j) = cstcolor_backcolor
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If j = mSpecColumn.规格_保质期 And mint状态 = 2 Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
            Next
        Next
    End With
End Sub

Private Sub SetBatch()
    '批量设置每一列的值
    Dim i As Integer
    
    With vsfDetails
        For i = 1 To .Rows - 1
            If .Cell(flexcpBackColor, i) <> mlngColor Then '只有在背景颜色不是灰色的情况下才能进行设置
                .TextMatrix(i, .Col) = .TextMatrix(.Row, .Col)
                If mint状态 = 1 Then   '品种
                    If .Col = mVaricolumn.品种_参考项目 Then
                        .TextMatrix(i, mVaricolumn.品种_参考项目ID) = .TextMatrix(.Row, mVaricolumn.品种_参考项目ID)
                    End If
                End If
                If mint状态 <> 1 Then   '规格
                    If .Col = mSpecColumn.规格_收入项目 Then
                        .TextMatrix(i, mSpecColumn.规格_收入项目id) = .TextMatrix(.Row, mSpecColumn.规格_收入项目id)
                    End If
                End If
                .Cell(flexcpForeColor, i, .Col) = mlngApplyColor
                .Cell(flexcpFontSize, i, .Col) = 10
                .Cell(flexcpFontBold, i, .Col) = True
                
                
                If .ColDataType(.Col) = 11 Then '复选框
                    .Cell(flexcpBackColor, i, .Col) = mlngApplyColor
                    If mint状态 <> 1 Then   '规格
                        If .Col = mSpecColumn.规格_药库分批 Then
                            If .TextMatrix(.Row, .Col) = "0" Then
                                .Cell(flexcpBackColor, i, mSpecColumn.规格_药房分批) = mlngColor
                                .Cell(flexcpBackColor, i, mSpecColumn.规格_保质期) = mlngColor
                                .TextMatrix(i, mSpecColumn.规格_保质期) = 0
                                .Cell(flexcpForeColor, i, mSpecColumn.规格_保质期) = mlngApplyColor
                                .Cell(flexcpFontBold, i, mSpecColumn.规格_保质期) = 10
                            Else
                                .Cell(flexcpBackColor, i, mSpecColumn.规格_药房分批) = cstcolor_backcolor
                                .Cell(flexcpBackColor, i, mSpecColumn.规格_保质期) = cstcolor_backcolor
                                .TextMatrix(i, mSpecColumn.规格_保质期) = 24
                                .Cell(flexcpForeColor, i, mSpecColumn.规格_保质期) = mlngApplyColor
                                .Cell(flexcpFontBold, i, mSpecColumn.规格_保质期) = 10
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub GetDefineSize(ByVal rsRecord As ADODB.Recordset)
    '功能：得到数据库的表字段的长度
    If mblnSetKey = False Then
        mblnSetKey = True
        With vsfDetails
            If mint状态 = 1 Then
                .ColKey(mVaricolumn.品种_通用名称) = rsRecord.Fields("通用名称").DefinedSize
                .ColKey(mVaricolumn.品种_英文名称) = rsRecord.Fields("英文名").DefinedSize
                .ColKey(mVaricolumn.品种_拼音码) = rsRecord.Fields("拼音码").DefinedSize
                .ColKey(mVaricolumn.品种_五笔码) = rsRecord.Fields("五笔码").DefinedSize
                .ColKey(mVaricolumn.品种_处方限量) = rsRecord.Fields("处方限量").DefinedSize
                .ColKey(mVaricolumn.品种_剂量单位) = rsRecord.Fields("剂量单位").DefinedSize
                .ColKey(mVaricolumn.品种_ATCCODE) = rsRecord.Fields("ATCCODE").DefinedSize
            Else
                .ColKey(mSpecColumn.规格_药品规格) = rsRecord.Fields("规格").DefinedSize
                .ColKey(mSpecColumn.规格_本位码) = rsRecord.Fields("本位码").DefinedSize
                .ColKey(mSpecColumn.规格_数字码) = rsRecord.Fields("数字码").DefinedSize
                .ColKey(mSpecColumn.规格_标识码) = rsRecord.Fields("标识码").DefinedSize
                .ColKey(mSpecColumn.规格_备选码) = rsRecord.Fields("备选码").DefinedSize
                .ColKey(mSpecColumn.规格_容量) = rsRecord.Fields("容量").DefinedSize
                .ColKey(mSpecColumn.规格_商品名称) = rsRecord.Fields("商品名").DefinedSize
                .ColKey(mSpecColumn.规格_生产厂牌) = rsRecord.Fields("生产厂牌").DefinedSize
                .ColKey(mSpecColumn.规格_拼音码) = rsRecord.Fields("拼音码").DefinedSize
                .ColKey(mSpecColumn.规格_五笔码) = rsRecord.Fields("五笔码").DefinedSize
                .ColKey(mSpecColumn.规格_合同单位) = rsRecord.Fields("合同单位").DefinedSize
                .ColKey(mSpecColumn.规格_批准文号) = rsRecord.Fields("批准文号").DefinedSize
                .ColKey(mSpecColumn.规格_注册商标) = rsRecord.Fields("注册商标").DefinedSize
                .ColKey(mSpecColumn.规格_售价单位) = rsRecord.Fields("售价单位").DefinedSize
                .ColKey(mSpecColumn.规格_剂量系数) = rsRecord.Fields("售价系数").DefinedSize
                .ColKey(mSpecColumn.规格_住院单位) = rsRecord.Fields("住院单位").DefinedSize
                mintLen = Val(rsRecord.Fields("住院单位").DefinedSize)
                .ColKey(mSpecColumn.规格_住院系数) = rsRecord.Fields("住院包装").DefinedSize
                .ColKey(mSpecColumn.规格_门诊单位) = rsRecord.Fields("门诊单位").DefinedSize
                .ColKey(mSpecColumn.规格_门诊系数) = rsRecord.Fields("门诊包装").DefinedSize
                .ColKey(mSpecColumn.规格_药库单位) = rsRecord.Fields("药库单位").DefinedSize
                .ColKey(mSpecColumn.规格_药库系数) = rsRecord.Fields("药库包装").DefinedSize
                .ColKey(mSpecColumn.规格_送货单位) = rsRecord.Fields("送货单位").DefinedSize
                .ColKey(mSpecColumn.规格_送货包装) = rsRecord.Fields("送货包装").DefinedSize
                .ColKey(mSpecColumn.规格_申领阀值) = rsRecord.Fields("申领阀值").DefinedSize
                .ColKey(mSpecColumn.规格_采购限价) = rsRecord.Fields("采购限价").DefinedSize
                .ColKey(mSpecColumn.规格_采购扣率) = rsRecord.Fields("采购扣率").DefinedSize
                .ColKey(mSpecColumn.规格_指导售价) = rsRecord.Fields("指导售价").DefinedSize
                .ColKey(mSpecColumn.规格_指导差率) = rsRecord.Fields("指导差率").DefinedSize
                .ColKey(mSpecColumn.规格_差价让利) = rsRecord.Fields("差价让利").DefinedSize
                .ColKey(mSpecColumn.规格_成本价格) = rsRecord.Fields("成本价").DefinedSize
                .ColKey(mSpecColumn.规格_当前售价) = rsRecord.Fields("当前售价").DefinedSize
                .ColKey(mSpecColumn.规格_管理费比例) = rsRecord.Fields("管理费比例").DefinedSize
                .ColKey(mSpecColumn.规格_增值税率) = rsRecord.Fields("增值税率").DefinedSize
                .ColKey(mSpecColumn.规格_保质期) = rsRecord.Fields("保质期").DefinedSize
                .ColKey(mSpecColumn.规格_标识说明) = rsRecord.Fields("标识说明").DefinedSize
                .ColKey(mSpecColumn.规格_输液注意事项) = rsRecord.Fields("输液注意事项").DefinedSize
            End If
        End With
   End If
End Sub

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    If bytIsWB Then
        strSql = "select zlWBcode('" & strInput & "'," & intOutNum & ") from dual"
    Else
        strSql = "select zlSpellcode('" & strInput & "'," & intOutNum & ") from dual"
    End If
    On Error GoTo ErrHand
    With rsTmp
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, "mdlCISBase", strSql)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "zlGetSymbol")
'        Call SQLTest
        zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Private Sub FindGridRow(ByVal strInput As String)
    '在控件中查询指定的品种和规格
    
    Dim lngStart As Long, lngRows As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim str其他名称 As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strFindStyle As String
    Dim strTmp As String
    
    If strInput = "" Then Exit Sub
    '查找药品
    If strInput = mstrFind Then
        '表示查找下一条记录
        If mlngFind >= vsfDetails.Rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '表示新的查找
        lngStart = 0
        mlngFindFirst = 0
        mstrFind = strInput
        
        strFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
        
        Set mrsFindName = New ADODB.Recordset
        
        If mint状态 = 1 Then    '品种
            gstrSql = "Select Distinct a.Id, a.编码" & _
                      "  From 诊疗项目目录 A, 诊疗项目别名 B " & _
                      " Where a.Id = b.诊疗项目id And a.类别 = [1] "
        Else    '规格
            gstrSql = "Select Distinct A.Id,A.编码 From 收费项目目录 A,收费项目别名 B" & _
                 " Where A.Id =B.收费细目id And A.类别=[1] "
        End If

        If IsNumeric(Replace(strInput, "-", "")) Then       '输入全是数字（或包含一个"-"）时只匹配编码
            gstrSql = gstrSql & " And A.编码 Like [2] Or B.简码 Like [2] And B.码类=3 "
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '输入全是字母时只匹配简码
            gstrSql = gstrSql & " And B.简码 Like [3] "
        ElseIf zlCommFun.IsCharChinese(strInput) Then       '输入全是汉字时只匹配名称
            gstrSql = gstrSql & " And B.名称 Like [3] "
        Else
            gstrSql = gstrSql & " And (A.编码 Like [2] Or B.名称 Like [3] Or B.简码 Like [3] )"
        End If
        
        gstrSql = gstrSql & " Order By A.编码 "
        
        If mstrNode Like "西成药*" Then
            strTmp = "5"
        ElseIf mstrNode Like "中成药*" Then
            strTmp = "6"
        Else
            strTmp = "7"
        End If
                 
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSql, "取匹配的药品ID", strTmp, strInput & "%", strFindStyle & strInput & "%")
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    lngStart = lngStart + 1
    lngRows = vsfDetails.Rows - 1
    
    With mrsFindName
        If .EOF Then .MoveFirst
        
        Do While Not .EOF
            If mint状态 = 1 Then    '品种
                lngFindRow = vsfDetails.FindRow(!编码, lngStart, mVaricolumn.品种_药品编码, True, True)
            Else    '规格
                lngFindRow = vsfDetails.FindRow(!编码, lngStart, mSpecColumn.规格_规格编码, True, True)
            End If
        
            If lngFindRow > 0 Then
                vsfDetails.SetFocus
                vsfDetails.TopRow = lngFindRow
                vsfDetails.Row = lngFindRow
                
                mlngFind = lngFindRow
                
                '记录找到的第1条记录
                If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
                
                mrsFindName.MoveNext
                Exit Do
            End If
            mrsFindName.MoveNext
    
            '如果到底了，则返回第1条记录
            If .EOF And lngFindRow = -1 Then
                mlngFind = mlngFindFirst
                If vsfDetails.Rows > 1 Then
                    vsfDetails.Row = 1
                End If
            End If
        Loop
    End With
End Sub

Public Function zlGetDigitSign(ByVal lngMediId As Long, ByVal strSpec As String) As String
    '-------------------------------------------------------------
    '功能：根据药品通用名称、剂型的数字标记码和规格前三位数值，产生返回药品七位码
    '入参：strSpellcode-通用名称的拼音码；strDoseCode:剂型的数字标记码, strSpec：规格数值
    '返回：药品简码
    '-------------------------------------------------------------
    Dim rsThis As New ADODB.Recordset
    Dim strSpellcode As String, strDoseCode As String
    Dim strChange As String
    Dim intLocate As Integer
    Dim strTemp As String
    Dim intCount As Integer
    
    gstrSql = "Select 简码 From 诊疗项目别名 where 诊疗项目id=[1] and 性质=1 and 码类=1"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strSpellcode = IIf(IsNull(rsThis!简码), "", rsThis!简码)
    Else
        strSpellcode = ""
    End If
    
    gstrSql = "select P.标记码 from 药品特性 T,药品剂型 P where T.药品剂型=P.名称(+) and 药名id=[1]"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strDoseCode = IIf(IsNull(rsThis!标记码), "", rsThis!标记码)
    Else
        strDoseCode = ""
    End If

    strChange = "AOEYUVBP MF DT NL GKHJQXZCSRW "
    
    strTemp = ""
    strSpellcode = Mid(strSpellcode, 1, 3)
    For intCount = 1 To Len(strSpellcode)
        intLocate = InStr(1, strChange, Mid(strSpellcode, intCount, 1))
        If intLocate Mod 3 = 0 Then
            intLocate = (intLocate \ 3) - 1
        Else
            intLocate = intLocate \ 3
        End If
        If intLocate <> -1 Then strTemp = strTemp & CStr(intLocate)
    Next
    strTemp = strTemp & strDoseCode & Format(Val(Mid(strSpec, 1, 3)), "000")
    zlGetDigitSign = strTemp
End Function

Private Sub ExitFrom()
    '退出时过程
    '判断界面中是否有值刚被修改了
    Dim i As Integer
    Dim j As Integer
    Dim intupdate As Integer
    Dim bln修改 As Boolean
    
    bln修改 = Check修改
    mintExit = 0
    
    If bln修改 = True Then
        intupdate = MsgBox("刚有内容被修改了，退出之前是否保存？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbYes Then
            mintExit = 2
            Call Save
            Unload Me
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Function CheckData() As Boolean
    '检查数据的合法性和完整性
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim intupdate As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo ErrHandle
    With vsfDetails
        If mint状态 = 1 Then '品种
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mVaricolumn.品种_通用名称) = "" Then
                    MsgBox "基本信息页第" & i & "行通用名称不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.基本信息).Selected = True
                    .Select i, mVaricolumn.品种_通用名称
                    Exit Function
                End If
                If .TextMatrix(i, mVaricolumn.品种_剂量单位) = "" Then
                    MsgBox "临床应用页第" & i & "行剂量单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.临床应用).Selected = True
                    .Select i, mVaricolumn.品种_剂量单位
                    Exit Function
                End If
                For j = 2 To .Rows - 1
                    If .TextMatrix(i, mVaricolumn.品种_通用名称) = .TextMatrix(j, mVaricolumn.品种_通用名称) And i <> j Then
                        MsgBox "基本信息页第" & i & "行通用名称与第" & j & "行通用名称相同了！", vbExclamation, gstrSysName
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .Select i, mVaricolumn.品种_通用名称
                        Exit Function
                    End If
                Next
            Next
        Else    '规格
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mSpecColumn.规格_药品规格) = "" Then
                    MsgBox "基本信息页第" & i & "行药品规格不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(0).Selected = True
                    .Select i, mSpecColumn.规格_药品规格
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_容量)) >= 100000000000# Then
                    MsgBox "基本信息页第" & i & "行容量过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(0).Selected = True
                    .Select i, mSpecColumn.规格_容量
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_剂量系数)) >= 100000 Then
                    MsgBox "包装单位页第" & i & "行售价系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_剂量系数
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_剂量系数) = "" Then
                    MsgBox "包装单位页第" & i & "行剂量系数不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_剂量系数
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.规格_售价单位)) = "" Then
                    MsgBox "包装单位页第" & i & "行售价单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_售价单位
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_门诊系数)) >= 100000 Then
                    MsgBox "包装单位页第" & i & "行门诊系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_门诊系数
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_门诊系数) = "" Then
                    MsgBox "包装单位页第" & i & "行门诊系数不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_门诊系数
                    Exit Function
                End If
                 If Trim(.TextMatrix(i, mSpecColumn.规格_门诊单位)) = "" Then
                    MsgBox "包装单位页第" & i & "行门诊单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_门诊单位
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_住院系数)) >= 100000 Then
                    MsgBox "包装单位页第" & i & "行" & IIf(mstrNode Like "中草药*", "药房系数", "住院系数") & "过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_住院系数
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_住院系数) = "" Then
                    MsgBox "包装单位页第" & i & "行" & IIf(mstrNode Like "中草药*", "药房系数", "住院系数") & "不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_住院系数
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.规格_住院单位)) = "" Then
                    MsgBox "包装单位页第" & i & "行" & IIf(mstrNode Like "中草药*", "药房单位", "住院单位") & "不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_住院单位
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_药库系数)) >= 100000 Then
                    MsgBox "包装单位页第" & i & "行药库系数过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_药库系数
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_药库系数) = "" Then
                    MsgBox "包装单位页第" & i & "行药库系数不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_药库系数
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.规格_药库单位)) = "" Then
                    MsgBox "包装单位页第" & i & "行药库单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_药库单位
                Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_申领阀值)) >= 100000 Then
                    MsgBox "包装单位页第" & i & "行申领阀值过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_申领阀值
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_送货包装) = "" And Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) <> "" Then
                    MsgBox "包装单位页第" & i & "行有送货单位情况下，送货包装不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_送货包装
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_送货包装) <> "" And Trim(.TextMatrix(i, mSpecColumn.规格_送货单位)) = "" Then
                    MsgBox "包装单位页第" & i & "行有送货包装情况下，送货单位不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.规格_送货单位
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_采购限价)) > 1000000 Then
                    MsgBox "价格信息页第" & i & "行采购限价过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购限价
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_指导售价)) > 1000000 Then
                    MsgBox "价格信息页第" & i & "行指导售价过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_指导售价
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_成本价格)) > 1000000 Then
                    MsgBox "价格信息页第" & i & "行成本价格过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_成本价格
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_当前售价)) > 1000000 Then
                    MsgBox "价格信息页第" & i & "行当前售价过大，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_当前售价
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_加成率)) > 1000000 Then
                    MsgBox "价格信息页第" & i & "行加成率超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_加成率
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_采购扣率)) > 100 Then
                    MsgBox "价格信息页第" & i & "行采购扣率超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购扣率
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_采购限价) = "" Then
                    MsgBox "价格信息页第" & i & "行采购限价不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购限价
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_成本价格) = "" Then
                    MsgBox "价格信息页第" & i & "行成本价格不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_成本价格
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_当前售价) = "" Then
                    MsgBox "价格信息页第" & i & "行当前售价不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_当前售价
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_指导售价) = "" Then
                    MsgBox "价格信息页第" & i & "行指导售价不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_指导售价
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.规格_采购扣率) = "" Then
                    MsgBox "价格信息页第" & i & "行采购扣率不能为空，请输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_采购扣率
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_指导差率)) > 100 Then
                    MsgBox "价格信息页第" & i & "行指导差率超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.规格_指导差率
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_管理费比例)) > 100 Then
                    MsgBox "药价属性页第" & i & "行管理费比例超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(4).Selected = True
                    .Select i, mSpecColumn.规格_管理费比例
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.规格_增值税率)) > 100 Then
                    MsgBox "药价属性页第" & i & "行增值税率超过了最大值，请重新输入！", vbInformation, gstrSysName
                    tbcDetails.Item(4).Selected = True
                    .Select i, mSpecColumn.规格_增值税率
                    Exit Function
                End If
                If CheckUnit(i) = False Then
                    Exit Function
                End If
                If CheckBatches(.TextMatrix(i, mSpecColumn.规格_药库分批) Like "*1", .TextMatrix(i, mSpecColumn.规格_药房分批) Like "*1") = True Then
                    If Not (.TextMatrix(i, mSpecColumn.规格_原药库分批) Like "*1" And .TextMatrix(i, mSpecColumn.规格_原药房分批) = "0") Then
                        MsgBox "分批管理页第" & i & "行当前有部门的工作性质同时设置了药库药房，请同时设置药库药房分批或不分批！", vbInformation, gstrSysName
                        tbcDetails.Item(5).Selected = True
                        .Select i, mSpecColumn.规格_药库分批
                        Exit Function
                    Else
                        n = n + 1
                        If n < 4 Then
                            strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & "[" & .TextMatrix(i, mSpecColumn.规格_规格编码) & "]" & _
                                                    .TextMatrix(i, mSpecColumn.规格_通用名称) & "(" & .TextMatrix(i, mSpecColumn.规格_药品规格) & ")" & "；"
                        End If
                    End If
                End If
                
            Next
        End If
    End With
    
    If strMsg <> "" Then
        MsgBox strMsg & vbCrLf & IIf(n > 3, "等以上超过3种", "以上" & n & "种") & "药品设置了药库分批药房不分批，且有部门的工作性质同时设置了药库和药房，请注意查看！", vbExclamation, gstrSysName
    End If
    CheckData = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckUnit(ByVal intRow As Integer) As Boolean
    Dim intOut As Integer, intIN As Integer
    Dim arr单位, arr系数
    Dim str单位 As String, str系数 As String
    Dim str单位_Tmp As String, str系数_Tmp As String
    Dim int位置 As Integer
    Dim strTemp As String
    
    With vsfDetails
        '检查是否存在单位名称一样，但系数不一致的情况
        '检查是否存在系数一样，但单位名称不一样的情况
        If mstrNode Like "中草药*" Then
            str单位 = .TextMatrix(intRow, mSpecColumn.规格_售价单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库单位)
            str系数 = .TextMatrix(intRow, mSpecColumn.规格_剂量系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库系数)
        Else
            str单位 = .TextMatrix(intRow, mSpecColumn.规格_售价单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_门诊单位) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库单位)
            str系数 = .TextMatrix(intRow, mSpecColumn.规格_剂量系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_住院系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_门诊系数) & "|" & .TextMatrix(intRow, mSpecColumn.规格_药库系数)
        End If
                
        '考虑到其他单位可能与售价单位一致，但系数肯定不一致，所以必须分开判断
        '除售价单位外的检查
        For intOut = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
            If mstrNode Like "中草药*" Then
                str单位_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))
                str系数_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))
            Else
                str单位_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_售价单位), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))))
                str系数_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_剂量系数), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))))
            End If
            arr单位 = Split(str单位, "|")
            arr系数 = Split(str系数, "|")
            For intIN = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
                If intIN <> intOut Then
                    '单位相同系数不同
                    If str单位_Tmp = arr单位(intIN - 1) And (Val(str系数_Tmp) <> Val(arr系数(intIN - 1))) Then
                        If mstrNode Like "中草药*" Then
                            strTemp = IIf(intOut = 2, "药房", "药库") & "单位与" & IIf(intIN = 2, "药房", "药库") & "单位一致，但其系数却不相同，请检查！"
                        Else
                            strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "单位与" & IIf(intIN = 2, "住院", IIf(intIN = 3, "门诊", "药库")) & "单位一致，但其系数却不相同，请检查！"
                        End If
                        
                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True
                        If InStr(1, strTemp, "单位与住院") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "单位与门诊") > 0 Then
                            int位置 = mSpecColumn.规格_门诊单位
                        ElseIf InStr(1, strTemp, "单位与药库") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        ElseIf InStr(1, strTemp, "药房单位一致") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "药库单位一致") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        End If
                        
                        .Select intRow, int位置
                        Exit Function
                    End If
                    If str单位_Tmp <> arr单位(intIN - 1) And (Val(str系数_Tmp) = Val(arr系数(intIN - 1))) Then
                        If mstrNode Like "中草药*" Then
                            strTemp = IIf(intOut = 2, "药房", "药库") & "包装与" & IIf(intIN = 2, "药房", "药库") & "包装一致，但其单位却不相同，请检查！"
                        Else
                            strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "包装与" & IIf(intIN = 2, "住院", IIf(intIN = 3, "门诊", "药库")) & "包装一致，但其单位却不相同，请检查！"
                        End If
                        
                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True
                        
                        If InStr(1, strTemp, "包装与住院") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "包装与门诊") > 0 Then
                            int位置 = mSpecColumn.规格_门诊单位
                        ElseIf InStr(1, strTemp, "包装与药库") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        ElseIf InStr(1, strTemp, "药房包装一致") > 0 Then
                            int位置 = mSpecColumn.规格_住院单位
                        ElseIf InStr(1, strTemp, "药库包装一致") > 0 Then
                            int位置 = mSpecColumn.规格_药库单位
                        End If
                        .Select intRow, int位置
                        Exit Function
                    End If
                End If
            Next
        Next
        
        '避免其它单位与售价单位相同，但系数不为1的情况
        '各单位与售价单位进行检查
        For intOut = 2 To IIf(mstrNode Like "中草药*" = True, 3, 4)
            If mstrNode Like "中草药*" Then
                str单位_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))
                str系数_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))
            Else
                str单位_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_售价单位), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院单位), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊单位), .TextMatrix(intRow, mSpecColumn.规格_药库单位))))
                str系数_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.规格_剂量系数), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.规格_住院系数), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.规格_门诊系数), .TextMatrix(intRow, mSpecColumn.规格_药库系数)))))
            End If
            
            If str单位_Tmp = .TextMatrix(intRow, mSpecColumn.规格_售价单位) And Val(str系数_Tmp) <> 1 Then
                If mstrNode Like "中草药*" Then
                    strTemp = IIf(intOut = 2, "药房", "药库") & "单位与售价单位一致，" & IIf(intOut = 2, "药房", "药库") & "系数应该为1"
                Else
                    strTemp = IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "单位与售价单位一致，" & IIf(intOut = 2, "住院", IIf(intOut = 3, "门诊", "药库")) & "系数应该为1"
                End If
                MsgBox strTemp, vbInformation, gstrSysName
                tbcDetails.Item(2).Selected = True
                
                If InStr(1, strTemp, "住院系数") > 0 Then
                    int位置 = mSpecColumn.规格_住院单位
                ElseIf InStr(1, strTemp, "门诊系数") > 0 Then
                    int位置 = mSpecColumn.规格_门诊单位
                ElseIf InStr(1, strTemp, "药库系数") > 0 Then
                    int位置 = mSpecColumn.规格_药库单位
                ElseIf InStr(1, strTemp, "药房系数") > 0 Then
                    int位置 = mSpecColumn.规格_住院单位
                ElseIf InStr(1, strTemp, "药库系数") > 0 Then
                    int位置 = mSpecColumn.规格_药库单位
                End If
                .Select intRow, int位置
                Exit Function
            End If
        Next
        
    End With
    CheckUnit = True
End Function

'Private Sub ShowPercent(sngPercent As Single)
''功能:在状态条上根据百分比显示当前处理进度(█)
'    Dim intAll As Integer
'    intAll = stbThis.Panels(2).Width / TextWidth("█") - 4
'    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "█")
'End Sub

Private Function Check修改() As Boolean
    '判断界面中是否有值刚被修改了
    '返回值为true 已经修改了 否者未修改
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        Check修改 = False
        For i = 1 To .Rows - 1
            For j = 1 To vsfDetails.Cols - 1
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                    Check修改 = True
                    Exit Function
                End If
            Next
        Next
    End With
End Function

Private Sub MoveRowCol()
    '行列移动方法
    With vsfDetails
        If mint状态 = 1 Then    '品种
            If mstrNode Like "中草药*" Then
                If tbcDetails.Selected.Index = mVariList.基本信息 Then    '基本页面
                    If .Col = mVaricolumn.品种_五笔码 Then
                        tbcDetails.Item(mVariList.品种属性).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_毒理分类
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.品种属性 Then    '品种属性
                    If .Col = mVaricolumn.品种_单味使用 Then
                        tbcDetails.Item(mVariList.临床应用).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_参考项目
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_毒理分类
                        ElseIf .Col = mVaricolumn.品种_药品类型 Then
                            .Col = mVaricolumn.品种_辅助用药
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.临床应用 Then    '临床应用
                    If .Col = mVaricolumn.品种_剂量单位 And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = mVaricolumn.品种_通用名称
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_参考项目
                        Else
                            If .Col <> mVaricolumn.品种_剂量单位 Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            Else    '西成药、中成药
                If tbcDetails.Selected.Index = mVariList.基本信息 Then    '基本页面
                    If .Col = mVaricolumn.品种_五笔码 Then
                        tbcDetails.Item(mVariList.品种属性).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_毒理分类
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.品种属性 Then    '品种属性
                    If .Col = mVaricolumn.品种_原料药 Then
                        tbcDetails.Item(mVariList.临床应用).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.品种_参考项目
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_毒理分类
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.临床应用 Then    '临床应用
                    If .Col = mVaricolumn.品种_品种下长期医嘱 And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.基本信息).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = 2
                    Else
                        If .Col = mVaricolumn.品种_通用名称 Then
                            .Col = mVaricolumn.品种_参考项目
                        Else
                            If .Col <> mVaricolumn.品种_品种下长期医嘱 Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            End If
        Else    '规格
            If mstrNode Like "中草药*" Then '中草药
                If tbcDetails.Selected.Index = mSpecList.基本信息 Then
                    If .Col = mSpecColumn.规格_备选码 Then
                        tbcDetails.Item(mSpecList.商品信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_商品名称
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.商品信息 Then
                    If .Col = mSpecColumn.规格_注册商标 Then
                        tbcDetails.Item(mSpecList.包装单位).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_售价单位
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_商品名称
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.规格_来源分类 Then
                            .Col = .Col + 3
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.包装单位 Then
                    If .Col = mSpecColumn.规格_中药形态 Then
                        tbcDetails.Item(mSpecList.价格信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药价属性
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_售价单位
                            Exit Sub
                        ElseIf .Col = mSpecColumn.规格_住院系数 Then
                            .Col = mSpecColumn.规格_药库单位
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.价格信息 Then
                    If .Col = mSpecColumn.规格_当前售价 Then
                        tbcDetails.Item(mSpecList.药价属性).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_收入项目
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药价属性
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.规格_指导售价 Then
                            .Col = mSpecColumn.规格_加成率
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.药价属性 Then
                    If .Col = mSpecColumn.规格_医保类型 Then
                        tbcDetails.Item(mSpecList.分批管理).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药库分批
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_收入项目
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.分批管理 Then
                    If .Col = mSpecColumn.规格_保质期 Then
                        tbcDetails.Item(mSpecList.临床应用).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_标识说明
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药库分批
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.临床应用 Then
                    If .Col <> mSpecColumn.规格_基本药物 Then
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_标识说明
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.基本信息).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.规格_药品规格
                        End If
                    End If
                End If
            Else    '西成药，中成药
                If tbcDetails.Selected.Index = mSpecList.基本信息 Then
                    If .Col = mSpecColumn.规格_容量 Then
                        tbcDetails.Item(mSpecList.商品信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_商品名称
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.商品信息 Then
                    If .Col = mSpecColumn.规格_非常备药 Then
                        tbcDetails.Item(mSpecList.包装单位).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_售价单位
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_商品名称
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.包装单位 Then
                    If .Col = mSpecColumn.规格_申领阀值 Then
                        tbcDetails.Item(mSpecList.价格信息).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药价属性
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_售价单位
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.价格信息 Then
                    If .Col = mSpecColumn.规格_当前售价 Then
                        tbcDetails.Item(mSpecList.药价属性).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_收入项目
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药价属性
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.规格_指导售价 Then
                            .Col = mSpecColumn.规格_加成率
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.药价属性 Then
                    If .Col = mSpecColumn.规格_医保类型 Then
                        tbcDetails.Item(mSpecList.分批管理).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_药库分批
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_收入项目
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.分批管理 Then
                    If .Col = mSpecColumn.规格_保质期 Then
                        tbcDetails.Item(mSpecList.临床应用).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_标识说明
                    Else
                        If .Col = mSpecColumn.规格_药品规格 Then
                            .Col = mSpecColumn.规格_药库分批
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.临床应用 Then
                    If .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_标识说明
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.规格_基本药物 Then
                        .Col = .Col + 1
                    Else
                        tbcDetails.Item(mSpecList.配药属性).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.规格_存储温度
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.配药属性 Then
                    If .Col = mSpecColumn.规格_药品规格 Then
                        .Col = mSpecColumn.规格_存储温度
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.规格_不予调配 Then
                        .Col = .Col + 1
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.基本信息).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.规格_药品规格
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfDetails_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDetails
        If mint状态 = 1 Then '品种
            If .Col = mVaricolumn.品种_参考项目 Then
                If .TextMatrix(Row, mVaricolumn.品种_参考项目ID) = "" Then
                    .TextMatrix(Row, mVaricolumn.品种_参考项目) = ""
                    .EditText = ""
                End If
            End If
        End If
    End With
End Sub
