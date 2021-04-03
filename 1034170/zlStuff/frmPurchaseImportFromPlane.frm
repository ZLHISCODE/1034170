VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseImportFromPlane 
   Caption         =   "导入计划单"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12210
   Icon            =   "frmPurchaseImportFromPlane.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12210
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   18
      Top             =   7080
      Width           =   3855
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor3 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   22
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor3 
         AutoSize        =   -1  'True
         Caption         =   "已停用"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.CheckBox chk允许导入停用卫材 
      Caption         =   "允许导入停用卫材"
      Height          =   180
      Left            =   7440
      TabIndex        =   15
      Top             =   7045
      Width           =   1815
   End
   Begin VB.Frame frmCondition 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   12012
      Begin VB.ComboBox cboStock 
         Height          =   276
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   205
         Visible         =   0   'False
         Width           =   1872
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   127467523
         CurrentDate     =   36263
      End
      Begin VB.CheckBox chkNoTime 
         Caption         =   "忽略"
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Tag             =   "1|0"
         Top             =   262
         Width           =   735
      End
      Begin VB.TextBox txtNo 
         Height          =   300
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   6
         Top             =   193
         Width           =   1605
      End
      Begin VB.CommandButton cmd提取 
         Caption         =   "提取(&G)"
         Height          =   350
         Left            =   10800
         TabIndex        =   5
         Top             =   168
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   195
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   127467523
         CurrentDate     =   36263
      End
      Begin VB.Label lbl移入库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "移入库房"
         Height          =   180
         Left            =   7920
         TabIndex        =   17
         Top             =   253
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "计划单审核日期"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   260
         Width           =   1260
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   255
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No"
         Height          =   180
         Left            =   5880
         TabIndex        =   9
         Top             =   252
         Width           =   180
      End
   End
   Begin VB.PictureBox picLine 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "产生入库单(&O)"
      Height          =   350
      Left            =   9480
      TabIndex        =   1
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   10920
      TabIndex        =   0
      Top             =   6960
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7428
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPurchaseImportFromPlane.frx":030A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16457
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2208
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "双击单据，选择要导入的计划单！"
      Top             =   840
      Width           =   12012
      _cx             =   21188
      _cy             =   3895
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
      BackColorSel    =   16764622
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseImportFromPlane.frx":0B9E
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   2772
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   12012
      _cx             =   21188
      _cy             =   4890
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
      BackColorSel    =   16764622
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseImportFromPlane.frx":0D0B
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "注意：未设置供应商的卫材将不会导入！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   3240
   End
End
Attribute VB_Name = "frmPurchaseImportFromPlane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSum As Long '记录导入卫材计划单中未导入停用卫材个数
Private mstrMsg As String '导入卫材计划单有停用卫材未导入时的提示信息

'主界面传入参数
Dim mfrmMain As Form
Dim mStr库房 As String
Dim mlng库房id As Long
Dim mintUnit As Integer                 '显示单位:0-散装单位,1-包装单位
Dim mbln所有库房 As Boolean
Dim mblnSuccess As Boolean
Private mint查询方式 As Integer     '用于区别是查询计划单还是申购单:0-计划单;1-申购单
Private mlngMode As Long
Private mint库存检查 As Integer             '表示卫材出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
Private mint明确批次 As Integer             '卫材是否按批次出库
Private mint不足提示 As Integer             '针对定价不分批卫材负数出库的情况

'其他参数
Dim mOraFMT As g_FmtString

Private Sub 提取卫材成本价和售价(ByRef rsData As ADODB.Recordset, ByVal lng材料ID As Long, ByVal lng批次 As Long, _
                                ByVal int是否分批 As Integer, ByVal bln是否变价 As Boolean, ByVal dbl比例系数 As Double)
    Dim rsprice As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If mint明确批次 = 1 Then
        gstrSQL = "Select a.实际数量, a.实际金额, a.实际差价, a.零售价, a.平均成本价, b.现价, c.成本价" & vbNewLine & _
                "From 药品库存 A, 收费价目 B, 材料特性 C" & vbNewLine & _
                "Where a.药品id = b.收费细目id And a.药品id = c.材料id And a.药品id = [1] And Nvl(a.批次, 0) = [2] And b.执行日期 <= Sysdate And" & vbNewLine & _
                "      b.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2"
    Else
        gstrSQL = "Select a.实际数量, a.实际金额, a.实际差价, a.零售价, a.平均成本价, b.现价, c.成本价" & vbNewLine & _
                "From 药品库存 A, 收费价目 B, 材料特性 C" & vbNewLine & _
                "Where a.药品id = b.收费细目id And a.药品id = c.材料id And a.药品id = [1] And b.执行日期 <= Sysdate And" & vbNewLine & _
                "      b.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd') And Rownum < 2"
    End If
         
    Set rsprice = zldatabase.OpenSQLRecord(gstrSQL, "提取卫材成本价和售价", lng材料ID, lng批次)
    
    If mint明确批次 = 1 Then
        If int是否分批 = 1 Then
            If bln是否变价 = False Then     '定价分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!现价), 0, rsprice!现价) * dbl比例系数
            Else                            '时价分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!零售价), 0, rsprice!零售价) * dbl比例系数
            End If
            
            rsData!成本金额 = rsData!实际数量 * rsData!成本价
            rsData!售价金额 = rsData!实际数量 * rsData!售价
        Else
            If bln是否变价 = False Then     '定价不分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!现价), 0, rsprice!现价) * dbl比例系数
            Else                            '时价不分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = rsprice!实际金额 / rsprice!实际数量 * dbl比例系数
            End If
            
            rsData!成本金额 = rsData!实际数量 * rsData!成本价
            rsData!售价金额 = rsData!实际数量 * rsData!售价
        End If
    Else
        If int是否分批 = 1 Then
            If bln是否变价 = False Then     '定价分批
                rsData!成本价 = IIf(IsNull(rsprice!成本价), 0, rsprice!成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!现价), 0, rsprice!现价) * dbl比例系数
            Else                            '时价分批
                rsData!成本价 = IIf(IsNull(rsprice!成本价), 0, rsprice!成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!现价), 0, rsprice!现价) * dbl比例系数
            End If
            
            rsData!成本金额 = rsData!实际数量 * rsData!成本价
            rsData!售价金额 = rsData!实际数量 * rsData!售价
        Else
            If bln是否变价 = False Then     '定价不分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = IIf(IsNull(rsprice!现价), 0, rsprice!现价) * dbl比例系数
            Else                            '时价不分批
                rsData!成本价 = IIf(IsNull(rsprice!平均成本价), 0, rsprice!平均成本价) * dbl比例系数
                rsData!售价 = rsprice!实际金额 / rsprice!实际数量 * dbl比例系数
            End If
            
            rsData!成本金额 = rsData!实际数量 * rsData!成本价
            rsData!售价金额 = rsData!实际数量 * rsData!售价
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function 检查库存(ByVal lng库房ID As Long, ByVal lng材料ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln出库分批 As Boolean
    Dim bln库房分批 As Boolean
    Dim bln在用分批 As Boolean
    Dim bln是否变价 As Boolean
    
    检查库存 = False
    On Error GoTo ErrHandle
    
    '如果是不分批卫材，且不检查库存，则直接退出此过程
    '---------------------------------------
    '提取当前卫材分批情况
    gstrSQL = "Select Nvl(a.库房分批, 0) 库房分批, Nvl(a.在用分批, 0) 在用分批, b.是否变价" & vbNewLine & _
            "From 材料特性 A, 收费项目目录 B" & vbNewLine & _
            "Where a.材料id = b.Id And a.材料id = [1]"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取当前卫材分批情况", lng材料ID)
    bln库房分批 = rsTemp!库房分批
    bln在用分批 = rsTemp!在用分批
    bln是否变价 = rsTemp!是否变价
    
    '提取出库库房部门性质
    gstrSQL = "Select 1 From 部门性质说明 Where 工作性质 In '发料部门' And 部门id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取当前卫材分批情况", lng库房ID)
    If rsTemp.EOF Then
        bln出库分批 = bln库房分批
    Else
        bln出库分批 = bln在用分批
    End If

    If bln出库分批 = False And mint库存检查 <> 2 And bln是否变价 = False Then
        检查库存 = True
        If mint库存检查 = 1 Then
            mint不足提示 = 1
        End If
        Exit Function
    End If
    '---------------------------------------
    
    '如果没有库存记录，则直接退出
    gstrSQL = "" & _
        "   Select Count(*) 记录数 From 药品库存 " & _
        "   Where 库房ID=[1] And 性质=1 And 药品ID=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查库存数据是否存在", lng库房ID, lng材料ID)
    If rsTemp!记录数 <> 0 Then
        检查库存 = True
        Exit Function
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 卫材分解(ByRef rsData As ADODB.Recordset, ByVal lng库房ID As Long, ByVal lng材料ID As Long, _
                    ByVal dbl填写数量 As Double, ByVal dbl比例系数 As Double)

    Dim rsTemp As New ADODB.Recordset
    Dim dbl可用数量 As Double
    Dim dbl剩余数量 As Double
    Dim bln出库分批 As Boolean
    Dim bln库房分批 As Boolean
    Dim bln在用分批 As Boolean
    Dim dbl成本价 As Double
    Dim dbl成本金额 As Double
    Dim dbl售价 As Double
    Dim dbl售价金额 As Double
    Dim str产地 As String
    Dim lng部门ID As Long
    Dim bln是否变价 As Boolean
          
    On Error GoTo ErrHandle
    
    '缓存数据
    str产地 = rsData!产地
    lng库房ID = rsData!库房id
    lng部门ID = rsData!部门ID
    
    '提取当前卫材分批情况
    gstrSQL = "Select Nvl(a.库房分批, 0) 库房分批, Nvl(a.在用分批, 0) 在用分批, b.是否变价" & vbNewLine & _
            "From 材料特性 A, 收费项目目录 B" & vbNewLine & _
            "Where a.材料id = b.Id And a.材料id = [1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取当前卫材分批情况", lng材料ID)
    bln库房分批 = rsTemp!库房分批
    bln在用分批 = rsTemp!在用分批
    bln是否变价 = rsTemp!是否变价
    
    '提取出库库房部门性质
    gstrSQL = "Select 1 From 部门性质说明 Where 工作性质 In '发料部门' And 部门id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取当前卫材分批情况", lng库房ID)
    If rsTemp.EOF Then
        bln出库分批 = bln库房分批
    Else
        bln出库分批 = bln在用分批
    End If
    
    '不按批次出库的卫材单独处理
    If mint明确批次 = 0 Then
        Call 提取卫材成本价和售价(rsData, rsData!材料ID, 0, IIf(bln出库分批, 1, 0), bln是否变价, dbl比例系数)
        Exit Sub
    End If
    
    '出库分批，则需分解;出库不分批,则无需分解
    gstrSQL = " Select Nvl(可用数量,0)/" & dbl比例系数 & " 可用数量,Nvl(批次,0) 批次, " & _
        " 上次产地 as 产地,上次批号 as 批号,效期,灭菌效期,上次生产日期 as 生产日期 " & _
        " From 药品库存 Where 性质=1 and 库房id = [1] And 药品id = [2] And nvl(可用数量,0)<>0 " & _
        " Order by Nvl(批次,0) "
        
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取可用库存", lng库房ID, lng材料ID)
        
    If bln出库分批 Then
        dbl剩余数量 = dbl填写数量
        If dbl剩余数量 > rsTemp!可用数量 Then
            rsData.Delete
            Do While Not rsTemp.EOF
                If rsTemp!可用数量 > 0 Then
                    If dbl剩余数量 > rsTemp!可用数量 Then
                        rsData.AddNew
                            
                        rsData!实际数量 = rsTemp!可用数量
                        rsData!批次 = rsTemp!批次
                        rsData!产地 = rsTemp!产地
                        rsData!批号 = rsTemp!批号
                        rsData!效期 = rsTemp!效期
                        rsData!灭菌效期 = rsTemp!灭菌效期
                        rsData!材料ID = lng材料ID
                        rsData!换算系数 = dbl比例系数
                        rsData!库房id = lng库房ID
                        rsData!部门ID = lng部门ID
                        
                        Call 提取卫材成本价和售价(rsData, rsData!材料ID, rsData!批次, 1, bln是否变价, dbl比例系数)
                        
                        dbl剩余数量 = dbl剩余数量 - rsTemp!可用数量
                    Else
                        rsData.AddNew
                        
                        rsData!实际数量 = dbl剩余数量
                        rsData!批次 = rsTemp!批次
                        rsData!产地 = rsTemp!产地
                        rsData!批号 = rsTemp!批号
                        rsData!效期 = rsTemp!效期
                        rsData!灭菌效期 = rsTemp!灭菌效期
                        rsData!材料ID = lng材料ID
                        rsData!换算系数 = dbl比例系数
                        rsData!库房id = lng库房ID
                        rsData!部门ID = lng部门ID
                        
                        Call 提取卫材成本价和售价(rsData, rsData!材料ID, rsData!批次, 1, bln是否变价, dbl比例系数)
                        
                        Exit Do
                    End If
                End If
                rsTemp.MoveNext
            Loop
        Else
            rsData!实际数量 = dbl填写数量
            rsData!批次 = rsTemp!批次
            rsData!产地 = rsTemp!产地
            rsData!批号 = rsTemp!批号
            rsData!效期 = rsTemp!效期
            rsData!灭菌效期 = rsTemp!灭菌效期
            
            Call 提取卫材成本价和售价(rsData, rsData!材料ID, rsData!批次, 1, bln是否变价, dbl比例系数)
        End If
    Else
        '根据库存检查判断填写数量是否大于可用数量。
        '1)若 可用数量 < 填写数量，则 填写数量 = 可用数量
        '2)若 可用数量 >= 填写数量，则 填写数量 = 填写数量
'        gstrSQL = " Select sum(Nvl(可用数量,0)/" & dbl比例系数 & ") 可用数量 From 药品库存 Where 库房id = [1] And 药品id = [2]"
'
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取可用库存", lng库房ID, lng材料ID)
    
        If mint库存检查 = 2 Then
            If rsTemp!可用数量 < dbl填写数量 Then
                rsData!实际数量 = rsTemp!可用数量
                If rsTemp!可用数量 = 0 Then
                    rsData.Delete
                    Exit Sub
                End If
            Else
                rsData!实际数量 = dbl填写数量
            End If
        Else
            rsData!实际数量 = dbl填写数量
        End If
        
        rsData!批次 = 0
        
        Call 提取卫材成本价和售价(rsData, rsData!材料ID, rsData!批次, 0, bln是否变价, dbl比例系数)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'获取当前库房的流通部门
Private Sub getDept()
    
    Dim rsTemp As New ADODB.Recordset
    
    '检查并装入移入库房
    err = 0: On Error Resume Next
    Set rsTemp = ReturnSQL(mlng库房id, Me.Caption, True, , 1716)
    With rsTemp
        cboStock.Clear
        Do While Not .EOF
            cboStock.AddItem !名称
            cboStock.ItemData(cboStock.NewIndex) = !Id
            .MoveNext
        Loop
        If cboStock.ListIndex < 0 Then cboStock.ListIndex = 0
    End With
End Sub

'检查数据依赖性
Private Function GetDepend() As Boolean
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    GetDepend = False
    With rsTemp
        '检查卫材入出类别是否完整
        strMsg = "没有设置卫材移库的入库及出库类别，请在入出分类中设置！"
        
        gstrSQL = "" & _
            "   SELECT B.Id,B.系数 " & _
            "   FROM 药品单据性质 A, 药品入出类别 B " & _
            "   Where A.类别id = B.ID  AND A.单据 = 34"
            
        zldatabase.OpenRecordset rsTemp, gstrSQL, "卫材移库管理"
        
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "系数=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的入库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = "系数=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "没有设置卫材移库的出库类别，请在入出分类中设置！"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    
    If mlngMode = 1716 Then
        Set rsTemp = ReturnSQL(mlng库房id, "卫材移库管理", True, , 1716)
        strMsg = "没有任何可移入库房，请在[卫材参数设置]的卫材流向中设置！"
    ElseIf mlngMode = 1722 Then
        Set rsTemp = ReturnSQL(mlng库房id, "卫材申领管理", True, , 1722)
        strMsg = "没有任何库房允许申领，请在[卫材参数设置]的卫材流向中设置！"
    End If
    rsTemp.Filter = "ID<>" & mlng库房id
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsTemp.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDetail()
    Dim rsTemp As New Recordset
    Dim IntBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim strUnit As String                       '单位名称:如门诊单位，住院单位等
    Dim str包装系数 As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        Select Case mintUnit
            Case 0
                str包装系数 = "1"
            Case Else
                str包装系数 = "D.换算系数"
        End Select
        
        
        gstrSQL = "" & _
            "   SELECT b.序号,'['||M.编码||']'||M.名称 as 名称, M.规格," & IIf(mintUnit = 0, "M.计算单位", "D.包装单位") & " as  单位," & _
            "           trim(to_char(b.前期数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 前期数量," & _
            "           trim(to_char(b.上期数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 上期数量," & _
            "           trim(to_char(b.库存数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 库存数量," & _
            "           trim(to_char(b.计划数量 /" & str包装系数 & "," & mOraFMT.FM_数量 & ")) 计划数量," & _
            "           trim(to_char(b.单价 *" & str包装系数 & "," & mOraFMT.FM_成本价 & ")) 单价," & _
            "           trim(to_char(b.金额," & mOraFMT.FM_金额 & ")) 金额, " & _
            " Trim(To_Char(Decode(M.是否变价, 0, P.现价 * " & str包装系数 & ", B.单价 * " & str包装系数 & " * (1+(1 / (1 - D.指导差价率 / 100) - 1))), " & mOraFMT.FM_零售价 & ")) 售价, " & _
            " Trim(To_Char(Decode(M.是否变价, 0, P.现价 , B.单价 * (1+(1 / (1 - D.指导差价率 / 100) - 1))) * B.计划数量," & mOraFMT.FM_金额 & ")) 售价金额, " & _
            " b.上次供应商 as 供应商,b.上次生产商 as 生产商,b.材料ID " & _
            "   FROM 材料采购计划 a, 材料计划内容 b,部门表 c,材料特性 D,收费项目目录 M, 收费价目 P " & _
            "   Where a.id = b.计划id " & _
            "           and nvl(a.库房id,0)=c.id(+) " & _
            "           and b.材料id=d.材料id and b.材料id=M.id  And M.ID = P.收费细目id " & _
            "   And (P.终止日期 Is Null Or Sysdate Between P.执行日期 And Nvl(P.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
            "           AND b.计划ID =[1] " & _
            "   Order by 序号"
        
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取计划内容", Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("ID"))))
        
        With vsfDetail
            .Redraw = flexRDNone
            Set .DataSource = rsTemp.DataSource
            .Redraw = flexRDDirect
            If mint查询方式 = 1 Then
                .TextMatrix(0, 7) = "申购数量"
                .ColHidden(.ColIndex("前期数量")) = True
                .ColHidden(.ColIndex("上期数量")) = True
                .ColHidden(.ColIndex("库存数量")) = True
            End If
        End With
        
        With vsfDetail
        
            '隐藏材料id列
            .ColHidden(.ColIndex("材料ID")) = True
            For i = 1 To .Rows - 1
                '判断是否停用，停用显示未
                If 是否停用(Val(.TextMatrix(i, .ColIndex("材料ID")))) Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF00FF
                End If
            Next
            
        End With
        
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetImportData(ByVal strNo As String) As ADODB.Recordset
    On Error GoTo ErrHandle
    If mint查询方式 = 0 Then
        gstrSQL = "Select 材料id, 编码, Sum(实际数量) As 实际数量, 成本价, Sum(成本金额) As 成本金额, 售价, Sum(售价金额) As 售价金额, 供应商id, 产地 " & _
            " From (Select B.材料id, M.编码, B.计划数量 As 实际数量, B.单价 As 成本价, B.金额 As 成本金额," & _
            " Decode(M.是否变价, 0, P.现价, B.单价 * (1+(1 / (1 - D.指导差价率 / 100) - 1))) As 售价, " & _
            " Decode(M.是否变价, 0, P.现价, B.单价 * (1+(1 / (1 - D.指导差价率 / 100) - 1))) * B.计划数量 As 售价金额, G.ID As 供应商id, B.上次生产商 As 产地 " & _
            " From 材料采购计划 A, 材料计划内容 B, 部门表 C, 材料特性 D, 收费项目目录 M, 收费价目 P, 供应商 G, " & _
            " Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) L " & _
            " Where A.ID = B.计划id And Nvl(A.库房id, 0) = C.ID(+) And B.材料id = D.材料id And B.材料id = M.ID And M.ID = P.收费细目id And " & _
            " (P.终止日期 Is Null Or Sysdate Between P.执行日期 And Nvl(P.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) And a.单据 = 0 And " & _
            " (G.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd') Or G.撤档时间 Is Null) And Substr(G.类型, 5, 1) = 1 And " & _
            " Nvl(G.末级, 0) = 1 And B.上次供应商 = G.名称 And A.NO = L.Column_Value) " & _
            " Group By 材料id, 编码, 成本价, 售价, 供应商id, 产地 " & _
            " Order By 供应商id, 编码 "
    ElseIf mint查询方式 = 1 And (mlngMode = 1716 Or mlngMode = 1722) Then
        gstrSQL = "Select 材料id, 编码, Sum(实际数量) As 实际数量, 产地, 库房id, 部门id, 换算系数" & vbNewLine & _
                "From (Select b.材料id, m.编码, b.计划数量 / d.换算系数 As 实际数量, d.换算系数, b.上次生产商 As 产地, 库房id, 部门id" & vbNewLine & _
                "       From 材料采购计划 A, 材料计划内容 B, 部门表 C, 材料特性 D, 收费项目目录 M, 收费价目 P," & vbNewLine & _
                "            Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)) L" & vbNewLine & _
                "       Where a.Id = b.计划id And Nvl(a.库房id, 0) = c.Id(+) And b.材料id = d.材料id And b.材料id = m.Id And m.Id = p.收费细目id And" & vbNewLine & _
                "             (p.终止日期 Is Null Or Sysdate Between p.执行日期 And Nvl(p.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) And a.单据 = 1 And" & vbNewLine & _
                "             a.No = l.Column_Value)" & vbNewLine & _
                "Group By 材料id, 编码, 产地, 库房id, 部门id, 换算系数" & vbNewLine & _
                "Order By 编码"
    End If
            
    Set GetImportData = zldatabase.OpenSQLRecord(gstrSQL, "取计划明细", strNo)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetList()
    Dim rsTemp As New Recordset
    Dim lng部门ID As Long
    
    On Error GoTo ErrHandle
    If mint查询方式 = 1 And (mlngMode = 1716 Or mlngMode = 1722) Then
        lng部门ID = cboStock.ItemData(cboStock.ListIndex)
    Else
        lng部门ID = 0
    End If
    
    If mint查询方式 = 0 Then
        gstrSQL = "" & _
            "   SELECT id,'' As 选择,期间,no, decode(计划类型,1,'月度计划',2,'季度计划',3,'年度计划','周度计划') as 计划类型 ," & _
            "           decode(编制方法,1,'往年同期线形参照法',2,'临近期间平均参照法',3,'材料储备定额参照法',4, '卫材日销售量参照法', '部门申购参照法') as 编制方法 ," & _
            "           编制人,to_char(编制日期,'yyyy-mm-dd HH24:MI:SS') as 编制日期, 审核人," & _
            "           to_char(审核日期,'yyyy-mm-dd HH24:MI:SS') as 审核日期,编制说明 " & _
            "   From 材料采购计划 a " & _
            "  Where 单据=0 And 审核日期 Is Not Null "
    Else
        gstrSQL = "" & _
                "   SELECT id,'' As 选择,期间,no, decode(计划类型,1,'月度计划',2,'季度计划',3,'年度计划','周度计划') as 计划类型 ," & _
                "           decode(编制方法,1,'往年同期线形参照法',2,'临近期间平均参照法','材料储备定额参照法') as 编制方法 ," & _
                "           编制人,to_char(编制日期,'yyyy-mm-dd HH24:MI:SS') as 编制日期, 审核人," & _
                "           to_char(审核日期,'yyyy-mm-dd HH24:MI:SS') as 审核日期,编制说明 " & _
                "   From 材料采购计划 a " & _
                "  Where 单据=1 And 审核日期 Is Not Null "
    End If
    
    
    If mint查询方式 = 0 Then
        If mbln所有库房 = True Then
            gstrSQL = gstrSQL & " And (nvl(库房id,0) =[1] Or 库房id Is Null) "
        Else
            gstrSQL = gstrSQL & " And nvl(库房id,0) =[1]"
        End If
    ElseIf mint查询方式 = 1 Then
        If mlngMode = 1716 Then
            gstrSQL = gstrSQL & " And nvl(库房id,0) =[1] and  部门id = [5] "
        ElseIf mlngMode = 1722 Then
            gstrSQL = gstrSQL & " And nvl(部门id,0) =[1] and  库房id = [5] "
        End If
    End If
    
    If chkNoTime.Value = 0 Then
        gstrSQL = gstrSQL & " and 审核日期 Between [2] And [3] "
    End If
    
    If Trim(txtNo.Text) <> "" Then
        gstrSQL = gstrSQL & " And No=[4] "
    End If
         
    gstrSQL = gstrSQL & " ORDER BY 期间,no "

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取采购计划", _
        mlng库房id, _
        CDate(Format(dtp开始时间.Value, "yyyy-mm-dd") & " 00:00:00"), _
        CDate(Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"), _
        txtNo.Text, _
        lng部门ID)
    
    With vsfList
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        .Redraw = flexRDDirect
        If mint查询方式 = 1 Then
            .ColHidden(.ColIndex("编制方法")) = True
        End If
        If rsTemp.EOF = False Then .Row = 1
        vsfDetail.Rows = 1
    End With
    
    staThis.Panels(2).Text = "当前共有" & rsTemp.RecordCount & "张单据；没有选择单据"
    
    Call vsfList_EnterCell
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveCard() As Boolean
    Dim intRow As Integer
    Dim strNo串 As String
    Dim rsData As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lngCur供应商ID As Long
    Dim int序号 As Integer
    Dim strNo As String
    Dim strDate As String
    Dim blnBeginTrans As Boolean
    
    On Error GoTo ErrHandle
    
    With vsfList
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("选择")) = "√" Then
                strNo串 = IIf(strNo串 = "", "", strNo串 & ",") & .TextMatrix(intRow, .ColIndex("NO"))
            End If
        Next
    End With
    
    If mint查询方式 = 0 And strNo串 = "" Then
        MsgBox "请选择要导入的采购计划单据！", vbOKOnly, gstrSysName
        Exit Function
    ElseIf mint查询方式 = 1 And strNo串 = "" Then
        MsgBox "请选择要导入的申购单据！", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set rsData = GetImportData(strNo串)
    
    If rsData Is Nothing Then Exit Function
    If rsData.EOF Then
        If mint查询方式 = 1 Then
            MsgBox "无法产生单据，请检查已选中的卫材是否有库存。"
        End If
        Exit Function
    End If
    
'    If mint查询方式 = 1 Then
'        If MsgBox("程序将自动对移出库房进行分解，请问是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'            Exit Function
'        End If
'    End If
    
    If mint查询方式 = 1 Then
        '建立本地记录集
        With rsTmp
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
            .Fields.Append "成本价", adDouble, , adFldIsNullable
            .Fields.Append "成本金额", adDouble, , adFldIsNullable
            .Fields.Append "售价", adDouble, , adFldIsNullable
            .Fields.Append "售价金额", adDouble, , adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
            .Fields.Append "效期", adDate, , adFldIsNullable
            .Fields.Append "灭菌效期", adDate, , adFldIsNullable
            .Fields.Append "材料ID", adDouble, 18, adFldIsNullable
            .Fields.Append "换算系数", adDouble, 18, adFldIsNullable
            .Fields.Append "库房ID", adDouble, 18, adFldIsNullable
            .Fields.Append "部门ID", adDouble, 18, adFldIsNullable
            .Fields.Append "批次", adDouble, 18, adFldIsNullable
            .Open
            
            rsData.MoveFirst
            Do While Not rsData.EOF
                If 是否导入(rsData!材料ID) Then
                    .AddNew
                    !实际数量 = IIf(IsNull(rsData!实际数量), 0, rsData!实际数量)
'                    !成本价 = IIf(IsNull(rsData!成本价), 0, rsData!成本价)
'                    !成本金额 = IIf(IsNull(rsData!成本金额), 0, rsData!成本金额)
'                    !售价 = IIf(IsNull(rsData!售价), 0, rsData!售价)
'                    !售价金额 = IIf(IsNull(rsData!售价金额), 0, rsData!售价金额)
                    !产地 = IIf(IsNull(rsData!产地), "", rsData!产地)
                    !材料ID = IIf(IsNull(rsData!材料ID), 0, rsData!材料ID)
                    !换算系数 = IIf(IsNull(rsData!换算系数), 1, rsData!换算系数)
                    !库房id = IIf(IsNull(rsData!库房id), 0, rsData!库房id)
                    !部门ID = IIf(IsNull(rsData!部门ID), 0, rsData!部门ID)
                    .Update
                End If
                rsData.MoveNext
            Loop
            
            rsTmp.Sort = "材料ID"
        End With
        
        If mlngSum > 0 Then
            If mlngMode = 1716 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个卫材已停用，这部分卫材将不导入移库单中！", "等" & mlngSum & "个卫材已停用，这部分卫材将不导入移库单中！"), vbInformation, gstrSysName
            ElseIf mlngMode = 1722 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个卫材已停用，这部分卫材将不导入申领单中！", "等" & mlngSum & "个卫材已停用，这部分卫材将不导入申领单中！"), vbInformation, gstrSysName
            End If
            mlngSum = 0
            mstrMsg = ""
        End If
        
        '库存检查
        If mlngMode = 1716 Then
            mint库存检查 = Get出库检查(mlng库房id)
        ElseIf mlngMode = 1722 Then
            mint库存检查 = Get出库检查(cboStock.ItemData(cboStock.ListIndex))
        End If
        
        '[按批次出库]或[库存检查为"不足禁止"]，则没有库存的卫材不能出库。
        If mint库存检查 = 2 Or mint明确批次 = 1 Then
            If Not rsTmp.EOF Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                '检查出库的卫材是否有库存
                If mlngMode = 1716 Then
                    If 检查库存(mlng库房id, rsTmp!材料ID) = False Then
                        rsTmp.Delete
                    End If
                ElseIf mlngMode = 1722 Then
                    If 检查库存(rsTmp!库房id, rsTmp!材料ID) = False Then
                        rsTmp.Delete
                    End If
                End If
                
                rsTmp.MoveNext
            Loop
            
            If mint不足提示 = 1 Then
                If MsgBox("部分卫材库存不足，请问是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                   Exit Function
               End If
            End If
            
            rsTmp.UpdateBatch
            
            If rsTmp.EOF And rsTmp.RecordCount = 0 Then
                MsgBox "无法产生单据，请检查已选中的卫材是否有可用库存。"
                Exit Function
            End If
            
            rsTmp.MoveFirst
        End If
        
        '按批次出库。则应该分解到对应的批次上。
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If mlngMode = 1716 Then
                '对卫材进行分解
                Call 卫材分解(rsTmp, mlng库房id, rsTmp!材料ID, rsTmp!实际数量, rsTmp!换算系数)
            End If
            If mlngMode = 1722 Then
                '对卫材进行分解
                Call 卫材分解(rsTmp, cboStock.ItemData(cboStock.ListIndex), rsTmp!材料ID, rsTmp!实际数量, rsTmp!换算系数)
            End If
            rsTmp.MoveNext
        Loop
        
        rsTmp.UpdateBatch
        
        If rsTmp.EOF And rsTmp.RecordCount = 0 Then
            MsgBox "无法产生单据，请检查已选中的卫材是否有可用库存。"
            Exit Function
        End If
            
        rsTmp.MoveFirst
    End If
    
    strDate = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    If mint查询方式 = 0 Then
        With rsData
            Do While Not .EOF
                If 是否导入(Val(!材料ID)) Then
                    int序号 = int序号 + 1
                    If lngCur供应商ID <> !供应商ID Then
                        lngCur供应商ID = !供应商ID
                        int序号 = 0
                        strNo = zldatabase.GetNextNo(68, mlng库房id)
                    End If
                    
                    gstrSQL = "zl_材料外购_INSERT("
                    '  No_In         In 药品收发记录.NO%Type,
                    gstrSQL = gstrSQL & "'" & strNo & "',"
                    '  序号_In       In 药品收发记录.序号%Type,
                    gstrSQL = gstrSQL & "" & int序号 & ","
                    '  库房id_In     In 药品收发记录.库房id%Type,
                    gstrSQL = gstrSQL & "" & mlng库房id & ","
                    '  供药单位id_In In 药品收发记录.供药单位id%Type,
                    gstrSQL = gstrSQL & "" & !供应商ID & ","
                    '  材料id_In     In 药品收发记录.药品id%Type,
                    gstrSQL = gstrSQL & "" & !材料ID & ","
                    '  产地_In       In 药品收发记录.产地%Type := Null,
                    gstrSQL = gstrSQL & "'" & !产地 & "',"
                    '  批号_In       In 药品收发记录.批号%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  生产日期_In   In 药品收发记录.生产日期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  效期_In       In 药品收发记录.效期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  灭菌日期_In   In 药品收发记录.灭菌日期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  实际数量_In   In 药品收发记录.实际数量%Type := Null,
                    gstrSQL = gstrSQL & "" & !实际数量 & ","
                    '  成本价_In     In 药品收发记录.成本价%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!成本价, g_小数位数.obj_最大小数.成本价小数) & ","
                    '  成本金额_In   In 药品收发记录.成本金额%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                    '  扣率_In       In 药品收发记录.扣率%Type := Null,
                    gstrSQL = gstrSQL & "100,"
                    '  零售价_In     In 药品收发记录.零售价%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!售价, g_小数位数.obj_最大小数.零售价小数) & ","
                    '  零售金额_In   In 药品收发记录.零售金额%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) & ","
                    '  差价_In       In 药品收发记录.差价%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) - Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                    '  零售差价_In   In 药品收发记录.差价%Type := Null,目前存放在用法字段
                    gstrSQL = gstrSQL & "Null,"
                    '  摘要_In       In 药品收发记录.摘要%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '   注册证号_In   In 药品收发记录.注册证号%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  填制人_In     In 药品收发记录.填制人%Type := Null,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '  随货单号_In   In 应付记录.随货单号%Type := Null
                    gstrSQL = gstrSQL & "Null,"
                    '  发票号_In     In 应付记录.发票号%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  发票日期_In   In 应付记录.发票日期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  发票金额_In   In 应付记录.发票金额%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                    gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'),"
                    '  核查人_In     In 药品收发记录.配药人%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  核查日期_In   In 药品收发记录.配药日期%Type := Null,
                    gstrSQL = gstrSQL & "Null,"
                    '  批次_In       In 药品收发记录.批次%Type := 0,
                    gstrSQL = gstrSQL & "0,"
                    '  退货_In       In Number := 1
                    gstrSQL = gstrSQL & "1)"
                        
                    If blnBeginTrans = False Then gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    blnBeginTrans = True
                
                End If
                        
                .MoveNext
            Loop
        End With
    Else
        With rsTmp
            Do While Not .EOF
'                If 是否导入(Val(!材料ID)) Then
                    int序号 = int序号 + 1
                    If mlngMode = 1716 Then    '卫材移库
                        If int序号 = 1 Then
                            strNo = sys.GetNextNo(72, mlng库房id)
                        Else
                            '因为移库是2个库房，所以序号以"2"递增
                            int序号 = int序号 + 1
                        End If
                            
                        gstrSQL = "Zl_材料移库_Insert("
                        '  No_In         In 药品收发记录.No%Type,
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '  序号_In       In 药品收发记录.序号%Type,
                        gstrSQL = gstrSQL & "" & int序号 & ","
                        '  库房id_In     In 药品收发记录.库房id%Type,
                        gstrSQL = gstrSQL & "" & mlng库房id & ","
                        '  对方部门id_In In 药品收发记录.对方部门id%Type,
                        gstrSQL = gstrSQL & "" & !部门ID & ","
                        '  材料id_In     In 药品收发记录.药品id%Type,
                        gstrSQL = gstrSQL & "" & !材料ID & ","
                        '  批次_In       In 药品收发记录.批次%Type,
                        gstrSQL = gstrSQL & IIf(mint明确批次 = 1, "" & !批次 & ",", "0,")
                        '  填写数量_In   In 药品收发记录.填写数量%Type,
                        gstrSQL = gstrSQL & "" & !实际数量 * !换算系数 & ","
                        '  实际数量_In   In 药品收发记录.实际数量%Type,
                        gstrSQL = gstrSQL & "" & !实际数量 * !换算系数 & ","
                        '  成本价_In     In 药品收发记录.成本价%Type,
                        gstrSQL = gstrSQL & "" & Round(!成本价 / !换算系数, g_小数位数.obj_最大小数.成本价小数) & ","
                        '  成本金额_In   In 药品收发记录.成本金额%Type,
                        gstrSQL = gstrSQL & "" & Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  零售价_In     In 药品收发记录.零售价%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价 / !换算系数, g_小数位数.obj_最大小数.零售价小数) & ","
                        '  零售金额_In   In 药品收发记录.零售金额%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  差价_In       In 药品收发记录.差价%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) - Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  填制人_In     In 药品收发记录.填制人%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '  产地_In       In 药品收发记录.产地%Type := Null,
                        gstrSQL = gstrSQL & "'" & !产地 & "',"
                        '  批号_In       In 药品收发记录.批号%Type := Null,
                        gstrSQL = gstrSQL & "'" & !批号 & "',"
                        '  效期_In       In 药品收发记录.效期%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!效期) = "", "Null", "to_date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!灭菌效期) = "", "Null", "to_date('" & Format(!灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  摘要_In       In 药品收发记录.摘要%Type := Null,
                        gstrSQL = gstrSQL & "Null,"
                        '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                        gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'))"
                            
                    ElseIf mint查询方式 = 1 And mlngMode = 1722 Then    '卫材申领
                        If int序号 = 1 Then
                            strNo = sys.GetNextNo(72, mlng库房id)
                        Else
                            '因为移库是2个库房，所以序号以"2"递增
                            int序号 = int序号 + 1
                        End If
                            
                        gstrSQL = "Zl_材料申领_Insert("
                        '  No_In         In 药品收发记录.No%Type,
                        gstrSQL = gstrSQL & "'" & strNo & "',"
                        '  序号_In       In 药品收发记录.序号%Type,
                        gstrSQL = gstrSQL & "" & int序号 & ","
                        '  库房id_In     In 药品收发记录.库房id%Type,
                        gstrSQL = gstrSQL & "" & !库房id & ","
                        '  对方部门id_In In 药品收发记录.对方部门id%Type,
                        gstrSQL = gstrSQL & "" & mlng库房id & ","
                        '  材料id_In     In 药品收发记录.药品id%Type,
                        gstrSQL = gstrSQL & "" & !材料ID & ","
                        '  批次_In       In 药品收发记录.批次%Type,
                        gstrSQL = gstrSQL & IIf(mint明确批次 = 1, "" & !批次 & ",", "0,")
                        '  填写数量_In   In 药品收发记录.填写数量%Type,
                        gstrSQL = gstrSQL & "" & !实际数量 * !换算系数 & ","
                        '  实际数量_In   In 药品收发记录.实际数量%Type,
                        gstrSQL = gstrSQL & "" & !实际数量 * !换算系数 & ","
                        '  成本价_In     In 药品收发记录.成本价%Type,
                        gstrSQL = gstrSQL & "" & Round(!成本价 / !换算系数, g_小数位数.obj_最大小数.成本价小数) & ","
                        '  成本金额_In   In 药品收发记录.成本金额%Type,
                        gstrSQL = gstrSQL & "" & Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  零售价_In     In 药品收发记录.零售价%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价 / !换算系数, g_小数位数.obj_最大小数.零售价小数) & ","
                        '  零售金额_In   In 药品收发记录.零售金额%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  差价_In       In 药品收发记录.差价%Type,
                        gstrSQL = gstrSQL & "" & Round(!售价金额, g_小数位数.obj_散装小数.金额小数) - Round(!成本金额, g_小数位数.obj_散装小数.金额小数) & ","
                        '  填制人_In     In 药品收发记录.填制人%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '  产地_In       In 药品收发记录.产地%Type := Null,
                        gstrSQL = gstrSQL & "'" & !产地 & "',"
                        '  批号_In       In 药品收发记录.批号%Type := Null,
                        gstrSQL = gstrSQL & "'" & !批号 & "',"
                        '  效期_In       In 药品收发记录.效期%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!效期) = "", "Null", "to_date('" & Format(!效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  灭菌效期_In   In 药品收发记录.灭菌效期%Type := Null,
                        gstrSQL = gstrSQL & IIf(NVL(!灭菌效期) = "", "Null", "to_date('" & Format(!灭菌效期, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ","
                        '  摘要_In       In 药品收发记录.摘要%Type := Null,
                        gstrSQL = gstrSQL & "Null,"
                        '  填制日期_In   In 药品收发记录.填制日期%Type := Null,
                        gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS'))"
                    
                    End If
                    
                    If blnBeginTrans = False Then gcnOracle.BeginTrans
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    blnBeginTrans = True
'                End If
            
                .MoveNext
            Loop
        End With
    End If
        
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '提示信息
    If mlngSum > 0 Then
        If mint查询方式 = 0 Then
            MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个卫材已停用，这部分卫材将不导入外购入库单中！", "等" & mlngSum & "个卫材已停用，这部分卫材将不导入外购入库单中！"), vbInformation, gstrSysName
        Else
            If mlngMode = 1716 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个卫材已停用，这部分卫材将不导入移库单中！", "等" & mlngSum & "个卫材已停用，这部分卫材将不导入移库单中！"), vbInformation, gstrSysName
            ElseIf mlngMode = 1722 Then
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "个卫材已停用，这部分卫材将不导入申领单中！", "等" & mlngSum & "个卫材已停用，这部分卫材将不导入申领单中！"), vbInformation, gstrSysName
            End If
        End If
        mlngSum = 0
        mstrMsg = ""
    End If
    
    SaveCard = True
    Exit Function
ErrHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'功能：判断卫材是否停用，再根据复选框“允许导入停用卫材”返回值
'当勾选时（允许导入停用卫材），不用判断卫材是否停用直接返回TRUE
'当不勾选时（不允许导入停用卫材），判断卫材是否停用，停用返回false
Private Function 是否导入(ByVal lng材料ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lng材料ID = 0 Then Exit Function
    
    If chk允许导入停用卫材.Value = 1 Then '允许导入停用卫材
        是否导入 = True
        Exit Function
    Else '不允许导入停用卫材
    
        '判断卫材是否停用
        gstrSQL = "select 名称,规格 from 收费项目目录 where ID = [1] and nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD')"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查卫材是否停用", lng材料ID)
        
        If rsTemp.RecordCount = 0 Then 'rsTemp.RecordCount = 0说明该卫材未停用
            是否导入 = True
        Else
            是否导入 = False
            
            mlngSum = mlngSum + 1
            If mlngSum <= 3 Then '拼提示信息串
                mstrMsg = mstrMsg & "【" & rsTemp!名称 & "(" & rsTemp!规格 & ")】" & Chr(10)
            End If
            
        End If
    End If

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(frmMain As Form, ByVal str库房 As String, ByVal lng库房ID As Long, ByVal intUnit As Integer, _
                    ByVal bln所有库房 As Boolean, Optional blnSuccess As Boolean = False, _
                    Optional int查询方式 As Integer, Optional lngMode As Integer, Optional int明确批次 As Integer)
    
    Set mfrmMain = frmMain
    
    mStr库房 = str库房
    mlng库房id = lng库房ID
    mintUnit = intUnit
    mbln所有库房 = bln所有库房
    mint查询方式 = int查询方式
    mlngMode = lngMode
    mint明确批次 = int明确批次
    
    If int查询方式 = 1 Then
        If Not GetDepend Then Exit Sub
    End If
    
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
End Sub





Private Sub chkNoTime_Click()
    If chkNoTime.Value = 0 Then
        dtp开始时间.Enabled = True
        dtp结束时间.Enabled = True
    Else
        dtp开始时间.Enabled = False
        dtp结束时间.Enabled = False
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mblnSuccess = SaveCard
    If mblnSuccess = True Then
        Unload Me
    End If
End Sub

Private Sub cmd提取_Click()
    GetList
End Sub


Private Sub Form_Activate()
    Me.Caption = Me.Caption & "(" & mStr库房 & ")"
End Sub

Private Sub Form_Load()

    chk允许导入停用卫材.Value = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Stuff", "允许导入停用卫材", 0)
    
    staThis.Panels(2).Picture = picColor
    
    dtp结束时间.Value = zldatabase.Currentdate
    dtp开始时间.Value = DateAdd("m", -1, Me.dtp结束时间.Value)
    
    If mint查询方式 = 1 Then
        chk允许导入停用卫材.Value = 0
        chk允许导入停用卫材.Visible = False
        lbl时间.Caption = "申购单审核日期"
        vsfList.ColHidden(5) = True     '隐藏[编制方法]
        vsfDetail.ColHidden(4) = True   '隐藏[前期数量]
        vsfDetail.ColHidden(5) = True   '隐藏[上期数量]
        vsfDetail.ColHidden(6) = True   '隐藏[库存数量]
        Me.Caption = "导入申购单"
        If mlngMode = 1716 Then
            CmdSave.Caption = "产生移库单(&O)"
        ElseIf mlngMode = 1722 Then
            CmdSave.Caption = "产生申领单(&O)"
        End If
        vsfDetail.TextMatrix(0, 7) = "申购数量"
        
        If mlngMode = 1716 Or mlngMode = 1722 Then
            If mlngMode = 1722 Then
                lbl移入库房.Caption = "发料库房"
            End If
            lbl移入库房.Visible = True
            cboStock.Visible = True
            Call getDept
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    Dim dblStateHeight As Double
    
    On Error Resume Next
    
    If Me.Height < 8325 Then Me.Height = 8325
    If Me.Width < 12420 Then Me.Width = 12420
    
    dblStateHeight = IIf(staThis.Visible, staThis.Height, 0)
    
    With CmdCancel
        .Top = Me.ScaleHeight - dblStateHeight - .Height - 200
        .Left = Me.ScaleWidth - .Width - 200
    End With
    
    With CmdSave
        .Top = CmdCancel.Top
        .Left = CmdCancel.Left - .Width - 200
    End With
    
    With chk允许导入停用卫材
        .Top = CmdSave.Top + (CmdSave.Height - .Height) / 2
        .Left = CmdSave.Left - .Width - 200
    End With
    
    With lblMsg
        .Top = chk允许导入停用卫材.Top
    End With
    
    With frmCondition
        .Width = Me.ScaleWidth - 200
    End With
    
    With vsfList
        .Width = frmCondition.Width
    End With
    
    With picLine
        .Top = vsfList.Top + vsfList.Height
        .Width = frmCondition.Width
    End With
    
    With vsfDetail
        .Top = picLine.Top + picLine.Height
        .Width = frmCondition.Width
        .Height = CmdCancel.Top - .Top - 200
    End With
        
    With cmd提取
        .Left = frmCondition.Width - .Width - 200
    End With
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - staThis.Panels(3).Width - staThis.Panels(4).Width - .Width - 300
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    '保存注册表信息(是否显示停用卫材)
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\zl9Stuff", "允许导入停用卫材", chk允许导入停用卫材.Value
End Sub

Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfList.Height + y <= 500 Or vsfDetail.Height - y <= 500 Then Exit Sub
        
        picLine.Top = picLine.Top + y
        vsfList.Height = vsfList.Height + y
        vsfDetail.Height = vsfDetail.Height - y
        vsfDetail.Top = vsfDetail.Top + y
        
        Me.Refresh
    End If
End Sub


Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If mint查询方式 = 1 Then
        If KeyCode = vbKeyReturn Then
            If Len(txtNo) < 8 And Len(txtNo) > 0 Then
                txtNo.Text = GetFullNO(txtNo.Text, 72, mlng库房id)
            End If
            zlCommFun.PressKey (vbKeyTab)
        End If
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        If Len(txtNo) < 8 And Len(txtNo) > 0 Then
            txtNo.Text = GetFullNO(txtNo.Text, 77, mlng库房id)
            GetList
        End If
    End If
End Sub


Private Sub vsfList_DblClick()
    Dim intRow As Integer
    Dim intSelectCount As Integer
    
    With vsfList
        If .Row = 0 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        .Redraw = flexRDNone
        
        If .TextMatrix(.Row, .ColIndex("选择")) = "√" Then
            .TextMatrix(.Row, .ColIndex("选择")) = ""
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000008
        Else
            .TextMatrix(.Row, .ColIndex("选择")) = "√"
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlue
        End If
        
        .Redraw = flexRDDirect
        
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("选择")) = "√" Then
                intSelectCount = intSelectCount + 1
            End If
        Next
        
        If intSelectCount = 0 Then
            staThis.Panels(2).Text = "当前共有" & .Rows - 1 & "张单据；没有选择单据"
        Else
            staThis.Panels(2).Text = "当前共有" & .Rows - 1 & "张单据；选择了" & intSelectCount & "张单据"
        End If
    End With
End Sub
Private Sub vsfList_EnterCell()
    GetDetail
End Sub


'功能：判断是否停用,true - 停用
Private Function 是否停用(ByVal lng药品ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If lng药品ID = 0 Then Exit Function

    
    '判断药品是否停用
    gstrSQL = "select 名称,规格 from 收费项目目录 where ID = [1] and nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD')) <> to_date('3000-01-01','YYYY-MM-DD') "
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查药品是否停用", lng药品ID)
    
    是否停用 = rsTemp.RecordCount <> 0  '说明该药品未停用

    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

