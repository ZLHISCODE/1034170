Attribute VB_Name = "mdlMediStore"
Option Explicit

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrprivs As String                  '当前用户具有的当前模块的功能
Public gstrStockSearchPrivs As String       '专门针对库存查询的权限

Public glngModul As Long
Public glngSys As Long                      '系统编号参数
Public gstrSysName As String                '系统名称
Public gstrVersion As String                '系统版本
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrSQL As String                    '用着作为所有临时SQL语句
Public gstrDbUser As String                 '当前登录ORACLE用户名
Public mblnCostPrice As Boolean             '出库单据是否显示成本价

Public Const StrFormat As String = "'999999999990.99999'"
Public gstrMatchMethod As String            '匹配方式:0表示双向匹配
Public gstrUserName As String               '传递用户姓名
Public gobjDrugPurchase As Object           '采购平台部件
Public gbytSimpleCodeTrans As Byte          '卡片界面是否允许简码切换控制

'用户信息------------------------
Public Type TYPE_USER_INFO
    用户ID As Long
    用户编码 As String
    用户姓名 As String
    用户简码 As String
    部门ID As Long
    部门编码 As String
    部门名称 As String
    strMaterial As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Enum 编辑
    '1.新增；2、修改；3、验收；4、查看；5、修改发票；6、冲销；
    '7、财务审核（冲销、产生新单据并审核；已付款的单据不允许财务审核，同样，财务审核后的单据不允许冲销）;
    '8-药库退货
    新增 = 1
    修改 = 2
    审核 = 3
    查阅 = 4
    修改发票 = 5            '允许对已审核的单据进行供药单位、发票信息进行修改
    冲销 = 6
    财务审核 = 7            '用于对已审核的单据进行成本价、供药单位及发票信息的审核（冲销原始单据，产生新单据）
    药库退货 = 8            '用于药库向供货单位退货
    核查 = 9                '用于核查成本价
    发送 = 10               '用于下出库库房的可用数量
End Enum

'药品库存查询中，各批次报警的字体颜色常数
Public Const glng报警 As Long = &HC00000
Public Const glng正常 As Long = &H80000008
Public Const glng停用 As Long = &HC0

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012
 
Public gint简码方式 As Integer              '0-拼音，1-五笔
Public gint药品名称显示 As Integer          '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
Public gint输入药品显示 As Integer          '0-按输入匹配显示，1-固定显示通用名和商品名

Public grsMaster As New ADODB.Recordset        '药品选择器：药品规格缓存数据集
Public grsMasterInput As New ADODB.Recordset   '药品选择器：药品规格录入简码时的缓存数据集
Public grsSlave As New ADODB.Recordset         '药品选择器：批次缓存数据集
Public grsUsingDrugs As New ADODB.Recordset    '在用药品缓存数据集

'模块号
Public Enum 模块号
    外购入库 = 1300
    自制入库 = 1301
    其他入库 = 1302
    差价调整 = 1303
    药品移库 = 1304
    药品领用 = 1305
    其他出库 = 1306
    药品盘点 = 1307
    药品计划 = 1330
    质量管理 = 1331
End Enum

'业务单据号
Public Enum 单据号
    外购入库 = 1
    自制入库 = 2
    协药入库 = 3
    其他入库 = 4
    差价调整 = 5
    药品移库 = 6
    药品领用 = 7
    收费处方发药 = 8
    记帐单处方发药 = 9
    记帐表处方发药 = 10
    其他出库 = 11
    盘点表 = 12
    调价变动 = 13
    盘点单 = 14
    留存记录 = 27
End Enum


'药品流通模块要使用到的系统参数
Public Type Type_SysParms
    P9_费用金额保留位数 As Integer
    P29_指导批发价定价单位 As Integer
    P44_输入匹配 As String
'    P54_时价药品以加价率入库 As Integer
    P64_审核限制 As Integer
    P73_明确申领药品批次 As Integer
    P75_外购入库需要核查 As Integer
    P76_时价药品直接确定售价 As Integer
    P85_药房查看单据成本价 As Integer
    P96_药品填单下可用库存 As Integer
    P126_时价药品售价加成方式 As Integer
    P149_效期显示方式 As Integer
    P150_药品出库优先算法 As Integer
    P173_经过标记付款后才能进行付款管理 As Integer
    P174_药品移库明确批次 As Integer
    P175_药品领用明确批次 As Integer
'    P181_药品入库按分段加成 As Integer
'    P183_时价取上次售价 As Integer
    P221_药品结存时点 As Integer
    P306_已失效药品禁止入库 As Integer
End Type
Public gtype_UserSysParms As Type_SysParms     '系统参数

'药品金额、价格、数量最大精度
Public Type Type_Digits
    Digit_金额 As Integer
    Digit_成本价 As Integer
    Digit_零售价 As Integer
    Digit_数量 As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

Public Type Type_SaleDigits
    Digit_成本价 As Integer
    Digit_零售价 As Integer
    Digit_数量 As Integer
End Type
Public gtype_UserSaleDigits As Type_SaleDigits

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type POINTAPI
     x As Long
     y As Long
End Type

'API申明
Public Const CB_FINDSTRING = &H14C
Public Const GWL_HWNDPARENT = (-8)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function Get分批属性(ByVal lng库房ID As Long, ByVal lng药品ID As Long) As Integer
    '返回指定库房，指定药品的分批属性
    '返回：0-不分批，1-分批
    Dim rsCheck As New ADODB.Recordset
    Dim int分批 As Integer
    Dim bln药房 As Boolean
    Dim strSQL As String
        
    On Error GoTo errHandle
    
    '判断是否是药房或制剂室
    strSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(strSQL, "Get分批属性", lng库房ID)

    bln药房 = (Not rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    strSQL = " Select Nvl(药库分批,0) As 药库分批,nvl(药房分批,0) As 药房分批 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(strSQL, "Get分批属性", lng药品ID)
              
    If bln药房 Then
        int分批 = rsCheck!药房分批
    Else
        int分批 = rsCheck!药库分批
    End If
    
    Get分批属性 = int分批
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function CheckIsAccount(ByVal lng库房ID As Long) As Boolean
'    '判断是否本期已经结存或结存已经审核
'    Dim rsData As ADODB.Recordset
'    Dim lng结存id As Long
'
'    gstrSQL = "Select Nvl(Max(ID), 0) as 结存id From 药品结存记录 Where 库房id = [1] "
'    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsAccount", lng库房ID)
'
'    lng结存id = rsData!结存ID
'
'    '如果之前进行过结存
'    If lng结存id > 0 Then
'        gstrSQL = "Select 期初日期, 期末日期, 填制人, 填制日期, 审核人, 审核日期, 上次结存id, 期间, 性质 From 药品结存记录 Where id=[1]"
'        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsAccount", lng结存id)
'
'        '检查是否有未审核的结存数据
'        If Not rsData.EOF Then
'            If Nvl(rsData!审核日期) = "" Then
'                MsgBox "提示：结存数据还未审核。" & vbCrLf & "为确保数据准确性，请先审核结存，再进行其他业务操作！", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'    End If
'
'    CheckIsAccount = True
'End Function

Public Sub AutoAdjustPrice_ByID(ByVal lngDrugID As Long)
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '按指定药品ID检查
    '在药品选择器中调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    
    On Error GoTo errHandle
    gstrSQL = "Select ID From 收费价目 Where 收费细目ID=[1] And 变动原因=0 And 执行日期 <= Sysdate "
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice_ByID", lngDrugID)

    With rsData
        Do While Not .EOF
            lngAdjustID = !id
            gstrSQL = "zl_药品收发记录_Adjust(" & lngAdjustID & ")"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice_ByID")
            
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_ByNO(ByVal int单据 As Integer, ByVal strNo As String)
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '按指定单据,NO中的药品号进行检查
    '个流通业务模块的审核时调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct a.Id " & _
        " From 收费价目 A, 药品收发记录 B " & _
        " Where a.收费细目id = b.药品id And a.变动原因 = 0 And a.执行日期 <= Sysdate And b.审核日期 Is Null " & _
        " And b.单据 = [1] And b.No = [2]"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice", int单据, strNo)

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call zlCommFun.ShowFlash("正在批量执行调价，请稍后......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !id
            gstrSQL = "zl_药品收发记录_Adjust(" & lngAdjustID & ")"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call zlCommFun.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub AutoAdjustPrice_Batch()
    '检查所有已到执行日期而价格未执行的药品，执行调价过程
    '检查所有药品
    '在药品选择器数据集初始化时调用
    Dim rsData As ADODB.Recordset
    Dim lngAdjustID As Long
    Dim blnMore As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select a.Id From 收费价目 A, 收费项目目录 B " & _
        " Where a.收费细目id = b.Id And a.变动原因 = 0 And a.执行日期 <= Sysdate And b.类别 In ('5', '6', '7') "
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "AutoAdjustPrice")

    With rsData
        If .RecordCount > 20 Then
            blnMore = True
            Call zlCommFun.ShowFlash("正在批量执行调价，请稍后......")
        End If
        
        Do While Not .EOF
            lngAdjustID = !id
            gstrSQL = "zl_药品收发记录_Adjust(" & lngAdjustID & ")"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "AutoAdjustPrice")
            
            .MoveNext
        Loop
        
        If blnMore = True Then
            Call zlCommFun.StopFlash
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNotVerifyClosingAccount() As ADODB.Recordset
    '查询当前操作员所属的部门是否存在未审核的结存记录
    Dim rsData As ADODB.Recordset
    Dim strDept As String
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct b.Id, b.名称, '未审核误差' As 类型" & vbNewLine & _
            "From 部门人员 A, 部门表 B, 部门性质说明 C, 药品结存记录 D, 药品结存误差 E" & vbNewLine & _
            "Where a.部门id = b.Id And b.Id = c.部门id And b.Id = d.库房id And d.Id = e.结存id And a.人员id = [1] And" & vbNewLine & _
            "      c.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室') And d.审核日期 Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Distinct b.Id, b.名称, '未审核结存' As 类型" & vbNewLine & _
            "From 部门人员 A, 部门表 B, 部门性质说明 C" & vbNewLine & _
            "Where a.部门id = b.Id And b.Id = c.部门id And a.人员id = [1] And c.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房', '制剂室') And" & vbNewLine & _
            "      Exists (Select 1 From 药品结存记录 D Where b.Id = d.库房id And d.审核日期 Is Null)"

    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "结存查询", UserInfo.用户ID)
    
    Set CheckNotVerifyClosingAccount = rsData
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'取药品金额、价格和数量的小数位数
Public Function GetDigitTiaoJia(ByVal int类别 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int类别：1-药品;2-卫材
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         药品单位:1-售价;2-门诊;3-住院;4-药库;
    '         卫材单位:1-散装;2-包装
    '性质: 0-计算金额;1-显示精度
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax金额 As Integer
    Dim intMax成本价 As Integer
    Dim intMax零售价 As Integer
    Dim intMax数量 As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum = 1"
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "取药品精度")
    
    intMax金额 = rs.Fields(0).NumericScale
    intMax成本价 = rs.Fields(1).NumericScale
    intMax零售价 = rs.Fields(2).NumericScale
    intMax数量 = rs.Fields(3).NumericScale
    
    If int内容 = 4 Then
        int单位 = 5
    End If
    gstrSQL = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 类别 = [1] And 内容 = [2] And 单位 = [3] "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int类别, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        GetDigitTiaoJia = rsTmp!精度
    End If
    
    If GetDigitTiaoJia = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        GetDigitTiaoJia = Choose(int内容, intMax成本价, intMax零售价, intMax数量, intMax金额)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigitTiaoJia = Choose(int内容, intMax成本价, intMax零售价, intMax数量, intMax金额)
End Function

Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub


Public Function Get现价(ByVal lng药品ID As Long) As Double
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) "
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "[提取该药品的零售单位价格]", lng药品ID)
    
    If Not rstemp.EOF Then
        Get现价 = rstemp!现价
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get售价(ByVal bln是否时价 As Boolean, lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long) As Double
    '功能：获取原始的售价单位售价，主要用于出库
    '参数: bln是否时价:false-定价,true-时价
    '返回值：最小单位的价格
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo errHandle

    '取定价药品售价
    If bln是否时价 = False Then
        gstrSQL = "Select 现价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) "
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "Get售价-取定价药品售价", lng药品ID)
        
        If Not rsData.EOF Then
            Get售价 = rsData!现价
        End If
    Else
        '取时价药品售价
        gstrSQL = "select Decode(Nvl(零售价, 0), 0, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 " & _
            " from 药品库存 where 性质=1 and  药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "GetOriPrice-零售价", lng药品ID, lng库房ID, lng批次)
        
        If rsData.EOF Then
            '时价药品零售价计算公式:采购价*(1+加成率)
            '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
            gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比  From 药品规格 Where 药品ID=[1]"
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get售价 = 0
            dbl成本价 = Get成本价(lng药品ID, lng库房ID, lng批次)
            dbl加成率 = Nvl(rsData!加成率, 0) / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
        Else
            If rsData!零售价 = 0 Then
                gstrSQL = "Select 指导零售价,nvl(加成率,15) as 加成率,Nvl(差价让利比,100) 差价让利比 ,加成率 From 药品规格 Where 药品ID=[1]"
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "零售价", lng药品ID)
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get售价 = 0
                dbl成本价 = Get成本价(lng药品ID, lng库房ID, lng批次)
                dbl加成率 = Nvl(rsData!加成率, 0) / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
            Else
                Get售价 = rsData!零售价
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

Public Function CheckNoStock(ByVal lng库房ID As Long, ByVal lng药品ID As Long, Optional ByVal lng批次 As Long = -1) As Boolean
    '检查是否无库存，用于判断时价不分批药品无库存盘点时新增
    '检查时不管批次，只管有无数据
    '返回：true-无库存;false-有库存
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 1 From 药品库存 " & _
        " Where 性质 = 1 And 库房id = [1] And 药品id = [2] And (Nvl(实际数量, 0) <> 0 Or Nvl(实际金额, 0) <> 0 Or Nvl(实际差价, 0) <> 0) "
    
    If lng批次 <> -1 Then
        gstrSQL = gstrSQL & " And Nvl(批次,0) = [3] "
    End If
    
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "CheckNoStock", lng库房ID, lng药品ID, lng批次)
    
    CheckNoStock = rsData.EOF
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetUserInfo() As Boolean
    Dim strSQL As String
    Dim rsUser As New ADODB.Recordset
    
    Set rsUser = zlDataBase.GetUserInfo
    
    With rsUser
        If Not .EOF Then
            UserInfo.用户ID = !id '当前用户id
            UserInfo.用户编码 = !编号 '当前用户编码
            UserInfo.用户姓名 = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            UserInfo.用户简码 = IIf(IsNull(!简码), "", !简码)  '当前用户简码
            UserInfo.部门ID = !部门ID '当前用户部门id
            UserInfo.部门编码 = !部门码 '当前用户
            UserInfo.部门名称 = !部门名  '当前用户
            UserInfo.strMaterial = GetMaterial(UserInfo.部门ID)
            GetUserInfo = True
        Else
            UserInfo.用户ID = 0 '当前用户id
            UserInfo.用户编码 = "" '当前用户编码
            UserInfo.用户姓名 = "" '当前用户姓名
            UserInfo.用户简码 = "" '当前用户简码
            UserInfo.部门ID = 0    '当前用户部门id
            UserInfo.部门编码 = ""  '当前用户
            UserInfo.部门名称 = ""  '当前用户
        End If
    End With
End Function
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer, Optional ByVal lng科室ID As Long) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    Dim intYear As Integer
    Dim PreFixNO As String  '年份前缀
    Dim strPre As String    '最大号码表中前2位
    Dim str编号 As String
    Dim dateCurDate As Date
    Dim intMonth As Integer
    Dim strMonth As String
    
    On Error GoTo errh
    
    dateCurDate = zlDataBase.Currentdate
    intYear = Format(dateCurDate, "YYYY") - 1990
    PreFixNO = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(dateCurDate)
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = PreFixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strSQL = "Select 编号规则,最大号码,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(strSQL, "GetFullNO", intNum)
        
    If Not rsTmp.EOF Then
        intType = Nvl(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
        strPre = Left(Nvl(rsTmp!最大号码, PreFixNO & "0"), 2)
    End If
    
    If intType = 0 Then
        '按年编号
        GetFullNO = strPre & Format(Right(strNo, 6), "000000")
    ElseIf intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNo, 4), "0000")
    ElseIf intType = 2 Then
        '按科室分月编码
        gstrSQL = "Select 编号 From 科室号码表 Where 项目序号=[1] And Nvl(科室ID,0)=[2]"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "GetFullNO", intNum, lng科室ID)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        If Nvl(rsTmp!编号) = "" Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        str编号 = Nvl(rsTmp!编号)
        
        '小于四位，按本月产生号码
        '五位或六位，则认为是指定月份的号码
        '七位，则认为是产生本年指定科室、月份的号码
        '大于等于八位，不处理
        If Len(strNo) <= 4 Then
            GetFullNO = PreFixNO & str编号 & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            GetFullNO = String(6 - Len(strNo), "0") & GetFullNO
            GetFullNO = PreFixNO & str编号 & GetFullNO
        ElseIf Len(strNo) = 7 Then
            GetFullNO = PreFixNO & GetFullNO
        End If
    End If
    Exit Function
errh:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub
Public Function CheckRepeatMedicine(ByVal MyBill As Object, ByVal strDrugInfo As String, ByVal intExceptRow As Integer) As Boolean
    '药品流通编辑界面检查录入的药品是否重复
    'MyBill：表单控件（药品列表）
    'strDrugInfo：药品ID，批次及对应列号（格式：药品ID,药品ID列号|批次,批次列号）
    'intExceptRow：排除指定的行（不检查这一行）
    Dim n As Integer
    Dim lng药品ID As Long
    Dim int药品ID列号 As Integer
    Dim lng批次 As Long
    Dim int批次列号 As Integer
    
    lng药品ID = Val(Split(Split(strDrugInfo, "|")(0), ",")(0))
    int药品ID列号 = Val(Split(Split(strDrugInfo, "|")(0), ",")(1))
    lng批次 = Val(Split(Split(strDrugInfo, "|")(1), ",")(0))
    int批次列号 = Val(Split(Split(strDrugInfo, "|")(1), ",")(1))
    
    With MyBill
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If n <> intExceptRow And Val(.TextMatrix(n, int药品ID列号)) = lng药品ID And Val(.TextMatrix(n, int批次列号)) = lng批次 Then
                    If MsgBox("对不起，已有该药品或该药品的相同批次，不能重复输入！ 需要移动到那行吗？" _
                        , vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                        .Row = n
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    CheckRepeatMedicine = True
End Function

Public Function GetCheck库房(ByVal lng库房ID As Long) As Integer
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "获取是否库存检查设置", lng库房ID)
    If Not rstemp.EOF Then GetCheck库房 = Nvl(rstemp!库存检查, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub CheckStopMedi(ByVal varInput As Variant)
    '检查药品是否停用
    'varInput两种格式：传入单据信息（单据|No）;传入药品ID串（格式：药品ID1，药品ID2.....）
    Dim rstemp As ADODB.Recordset
    Dim strMsg As String
    Dim int单据 As Integer
    Dim strNo As String
    Dim n As Integer
    Dim str药名 As String
    
    On Error GoTo errHandle
    If InStr(varInput, "|") > 0 Then
        int单据 = Mid(varInput, 1, InStr(varInput, "|") - 1)
        strNo = Mid(varInput, InStr(varInput, "|") + 1)
        
        gstrSQL = "Select Distinct '[' || C.编码 || ']' AS 药品编码,C.名称 As 通用名,B.名称 As 商品名 " & _
                " From 药品收发记录 A, 收费项目别名 B, 收费项目目录 C " & _
                " Where A.药品id = C.ID And A.药品id = B.收费细目id(+) And B.性质(+) = 3 " & _
                " And Nvl(C.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') " & _
                " And A.单据 = [1] And A.NO = [2]"
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查停用药品", int单据, strNo)
    Else
        gstrSQL = "Select Distinct '[' || C.编码 || ']' AS 药品编码,C.名称 As 通用名,B.名称 As 商品名 " & _
                " From Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) A, 收费项目别名 B, 收费项目目录 C " & _
                " Where A.Column_Value = C.ID  And A.Column_Value = B.收费细目id(+) And B.性质(+) = 3 " & _
                " And Nvl(C.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')) <> To_Date('3000-01-01', 'yyyy-MM-dd') "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查停用药品", varInput)
    End If
    
    With rstemp
        If Not .EOF Then
            For n = 1 To .RecordCount
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = !药品编码 & !通用名
                Else
                    str药名 = !药品编码 & IIf(IsNull(!商品名), !通用名, !商品名)
                End If
                
                If n > 5 Then
                    strMsg = strMsg & vbCrLf & "还有其他" & .RecordCount - 5 & "个药品......"
                    Exit For
                End If
                strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & str药名
                .MoveNext
            Next
            
            strMsg = "注意，以下药品已被停用：" & vbCrLf & strMsg
        End If
    End With
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckNumStock(ByVal objVSF As Object, ByVal lng库房ID As Long, ByVal lntCol药品id As Integer, ByVal intCol批次 As Integer, ByVal intCol数量 As Integer, ByVal intCol比例系数 As Integer, ByVal intMethod As Integer, Optional int入出业务 As Integer, Optional ByVal int精度 As Integer) As String
    '功能：审核出库类单据时，检查库存表实际数量是否足够
    '参数：objVSF-需要检查的表格;lng库房id；intcol批次-批次所在列；intCol数量-数量所在列；intCol比例系数-比例系数所在列
    '参数：intMethod，1-正常审核，2-冲销，3-退库审核
    '参数：int入出业务，0-入库；1-出库
    '返回值：哪行具体的药品名称，为空-检查通过，数量充足；不为空-检查未通过，数量不充足
    Dim objCol As Collection         '已使用的数量集合
    Dim i, j As Integer
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim rsData As ADODB.Recordset
    Dim strkey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    Dim dbl比例系数 As Double
    Dim intSum  As Integer
    
    On Error GoTo errHandle
    
    '先组合表格中数量，组合数量主要是考虑不分批的情况
    Set objCol = New Collection
    With objVSF
        If .rows < 2 Then Exit Function
        For lngRow = 1 To .rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntCol药品id) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntCol药品id) & "," & Val(.TextMatrix(lngRow, intCol批次)) & "," & Val(.TextMatrix(lngRow, intCol比例系数)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strkey = .TextMatrix(lngRow, lntCol药品id) & "," & Val(.TextMatrix(lngRow, intCol批次)) & "," & Val(.TextMatrix(lngRow, intCol比例系数))
                '以最小单位保存数量，方便审核时数量与库存数据比较
                strArray = dblNum + (Val(.TextMatrix(lngRow, intCol数量)))
                objCol.Add Array(strkey, strArray), strkey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '格式是药品id,批次,比例系数
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int入出业务 = 0 Then '入库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数入库，需要减库存，所以需要判断库存是否充足
                    dblNum = Abs(dblNum)
                Else
                    '正数入库，不见库存，所以不检查
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '退库审核，退库必须录入正数
                dblNum = dblNum
            End If
        Else    '出库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数入库，需要减库存，所以需要判断库存是否充足
                    dblNum = 0
                Else
                    '正数入库，不见库存，所以不检查
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        '只有有数量才判断
        If dblNum > 0 Then
'            For i = 0 To UBound(varTemp)
            lng药品ID = varTemp(0)
            lng批次 = varTemp(1)
            dbl比例系数 = varTemp(2)
            
            gstrSQL = "Select Nvl(a.实际数量, 0) As 实际数量, '[' || b.编码 || ']' || b.名称 名称" & vbNewLine & _
                        "From 药品库存 A, 收费项目目录 B" & vbNewLine & _
                        "Where a.药品id = b.Id And a.药品id = [2] And a.库房id = [3] And Nvl(a.批次, 0) = [4] And b.类别 In ('5', '6', '7') And a.性质 = 1"
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", dblNum, lng药品ID, lng库房ID, lng批次)
            If rsData.RecordCount = 0 Then
                gstrSQL = "select '[' || 编码 || ']' || 名称 名称 from 收费项目目录 where id=[1]"
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", lng药品ID)
                
                intSum = intSum + 1
                If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "、" & vbCrLf) & rsData!名称
            Else
                If zlStr.FormatEx(rsData!实际数量 / dbl比例系数, int精度, , False) >= dblNum Then
                Else
                    intSum = intSum + 1
                    If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "、" & vbCrLf) & rsData!名称
                End If
            End If
'            Next
        End If
    Next
    
    CheckNumStock = CheckNumStock & IIf(intSum > 3, "等" & intSum & "个", "")
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 库存实际数量检查(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, ByVal dbl数量 As Double, ByVal int比例系数 As Integer, ByVal lng小数位数 As Long) As Boolean
'功能：在出库时检查库存实际数量是否足够，足够则返回true，反之为false
    Dim rsData As ADODB.Recordset
    Dim str条件 As String
    
    On Error GoTo errHandle
    
    '检查有无库存
    If lng药品ID <= 0 Then Exit Function
    If lng库房ID <= 0 Then Exit Function
    
    gstrSQL = "Select a.实际数量, '[' || b.编码 || ']' || b.名称 名称" & vbNewLine & _
                            "From 药品库存 A, 收费项目目录 B" & vbNewLine & _
                            "Where a.药品id = b.Id And a.药品id = [1] And a.库房id = [2] And Nvl(a.批次, 0) = [3] And b.类别 In ('5', '6', '7') And a.性质 = 1"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", lng药品ID, lng库房ID, lng批次)
    
    If rsData.RecordCount = 0 Then '无库存记录
        gstrSQL = "select '[' || 编码 || ']' || 名称 名称 from 收费项目目录 where id=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "库存检查", lng药品ID)
        
        库存实际数量检查 = False
        Exit Function
    Else '有库存记录
        库存实际数量检查 = zlStr.FormatEx(rsData!实际数量 / int比例系数, lng小数位数, , True) >= dbl数量 '实际数量大于出库数量
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckStrickUsable(ByVal int单据 As Integer, ByVal lng库房ID As Long, _
        ByVal lng药品ID As Long, ByVal str药品名称 As String, _
        ByVal lng批次 As Long, ByVal dbl冲销数量 As Double, ByVal int库存检查 As Integer, _
        Optional ByVal strNo As String = "", Optional ByVal int序号 As Integer = 0) As Boolean
    '冲销单据时检查：原单据入库库房是否可用数量足够（可用数量等于或小于实际数量），实际冲销数量不能大于可用数量
    '对于移库单据、他入库单，需要取原单据入库那笔的批次，再根据批次来取可用数量；
    '对于自制入库、协定入库单据，由于是全部冲销，可以根据单据号，序号来取冲销数量，再来和库存可用数量比较
    '其他单据可直接根据批次取库存可用数量
    'int库存检查：表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止
    '只有冲销时是出库类型（原单据是入库类型）的要做此检查：外购入库、自制入库（原单据入的那笔）、协定入库（原单据入的那笔）、其他入库、移库（原单据入的那笔）
    
    Dim rstemp As ADODB.Recordset
    Dim lng入库批次 As Long
    Dim dbl可用数量 As Double
    
    On Error GoTo errHandle
    '冲销数量为0时可以不需要校验库存数量（排除了因为其他问题造成库存可用数量小于0，进而无法冲销的情况）
    If dbl冲销数量 = 0 Then
        CheckStrickUsable = True
        Exit Function
    End If
    
    If int单据 = 2 Or int单据 = 3 Then  '自制入库、协定入库单据
        If strNo = "" Or int序号 = 0 Then Exit Function
        gstrSQL = "Select 1 From 药品收发记录 A, 药品库存 B " & _
            " Where A.单据 = [1] And A.NO = [2] And A.序号 = [3] And A.记录状态 = 1 And A.入出系数 = 1 And B.性质 = 1 And A.库房id = B.库房id And A.药品id = B.药品id And " & _
            " Nvl(A.批次, 0) = Nvl(B.批次, 0) And A.实际数量 > B.可用数量 And Rownum = 1"
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查可用数量", int单据, strNo, int序号)
        
        '按正常流程进行提示或禁止
        If rstemp.RecordCount > 0 Then
            Select Case int库存检查
            Case 1  '提示
                If MsgBox(str药品名称 & "的可用库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '禁止
                MsgBox str药品名称 & "的可用库存不足！", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    Else
        If int单据 = 6 Or int单据 = 4 Then   '移库单，其他入库单
            If strNo = "" Or int序号 = 0 Then Exit Function
            
            gstrSQL = "Select Nvl(批次, 0) 批次 From 药品收发记录 Where 单据 = [1] And NO = [2] And 序号 = [3] And 药品id = [4] And 入出系数 = 1"
            Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "取入库批次", int单据, strNo, int序号, lng药品ID)
            
            If rstemp.RecordCount = 0 Then Exit Function
            
            lng入库批次 = rstemp!批次
        Else
            '其他单据根据传入的批次来取库存可用数量
            lng入库批次 = lng批次
        End If
        
        gstrSQL = "Select Nvl(可用数量, 0) 可用数量 From 药品库存 Where 性质 = 1 And 库房id = [1] And 药品id = [2] And Nvl(批次, 0) = [3] "
        Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "取可用数量", lng库房ID, lng药品ID, lng入库批次)
        
        If rstemp.RecordCount > 0 Then
            dbl可用数量 = rstemp!可用数量
        End If
        
        '按正常流程进行提示或禁止
        If dbl可用数量 < Abs(dbl冲销数量) Then
            Select Case int库存检查
            Case 1  '提示
                If MsgBox(str药品名称 & "的可用库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Case 2  '禁止
                MsgBox str药品名称 & "的可用库存不足！", vbInformation, gstrSysName
                Exit Function
            End Select
        End If
    End If
    
    CheckStrickUsable = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetControlItem(ByVal int单据 As Integer, ByVal int环节 As Integer) As String
    '单据环节控制，允许修改的项目，暂时只有外购入库
    'int环节：1-核查;2-审核;3-财务审核（药品外购）
    '所有项目：采购价,扣率,结算价,结算金额,售价,外观,发票号,发票日期,发票金额
    Dim rsTmp As ADODB.Recordset
    Dim strControlItem As String
    Const cst单据_外购 As Integer = 1
    Const cst环节_核查 As Integer = 1
    Const cst环节_审核 As Integer = 2
    Const cst环节_财务审核 As Integer = 3
    Const cst项目_核查 As String = "结算价,采购价,售价,外观"
    Const cst项目_审核 As String = "外观,发票号,发票代码,发票日期,发票金额"
    Const cst项目_财务审核 As String = "采购价,扣率,结算价,结算金额,外观,发票号,发票代码,发票日期,发票金额"
    
    On Error GoTo errHandle
    gstrSQL = "Select 内容 From 单据环节控制 Where 单据 = [1] And 环节 = [2]"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "单据环节控制", int单据, int环节)
    
    If Not rsTmp.EOF Then
        strControlItem = IIf(IsNull(rsTmp!内容), "", rsTmp!内容)
    End If
    
    If strControlItem = "" Then
        Select Case int单据
            Case cst单据_外购
                Select Case int环节
                    Case cst环节_核查
                        strControlItem = cst项目_核查
                    Case cst环节_审核
                        strControlItem = cst项目_审核
                    Case cst环节_财务审核
                        strControlItem = cst项目_财务审核
                End Select
        End Select
    End If
    
    GetControlItem = strControlItem
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get精度(ByVal int内容 As Integer, ByVal int单位 As Integer) As Integer
    '功能：用来返回成本价和售价、数量允许输入的长度
    '参数1：int内容=1 成本价;int内容=2 零售价;int内容=3 数量
    '参数2：int单位=1 售价;int单位=2 门诊;int单位=3 住院;int单位=4 药库
    '返回值：根据参数判断精度大小
    Dim rstemp As ADODB.Recordset
    Dim strFilter As String
    On Error GoTo errHandle
    
    gstrSQL = "Select 内容,单位,Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = 0 And 类别 = 1"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "查询精度")
    
    strFilter = " 内容=" & int内容 & " And 单位=" & int单位
    rstemp.Filter = strFilter
    
    If rstemp.RecordCount > 0 Then
        Get精度 = rstemp!精度
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'取系统参数值
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errh
    
    gstrSQL = "Select 参数号, 参数值, 缺省值 From Zlparameters Where 系统 = [1] And Nvl(私有, 0) = 0 And 模块 Is Null Order By 参数号 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "取系统参数", glngSys)
    
    If rs.RecordCount > 0 Then
        rs.Filter = "参数号=9"
        If Not rs.EOF Then gtype_UserSysParms.P9_费用金额保留位数 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=29"
        If Not rs.EOF Then gtype_UserSysParms.P29_指导批发价定价单位 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=44"
        If Not rs.EOF Then gtype_UserSysParms.P44_输入匹配 = Val(Nvl(rs!参数值, rs!缺省值))
        
'        rs.Filter = "参数号=54"
'        If Not rs.EOF Then gtype_UserSysParms.P54_时价药品以加价率入库 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=64"
        If Not rs.EOF Then gtype_UserSysParms.P64_审核限制 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=73"
        If Not rs.EOF Then gtype_UserSysParms.P73_明确申领药品批次 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=75"
        If Not rs.EOF Then gtype_UserSysParms.P75_外购入库需要核查 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=76"
        If Not rs.EOF Then gtype_UserSysParms.P76_时价药品直接确定售价 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=85"
        If Not rs.EOF Then gtype_UserSysParms.P85_药房查看单据成本价 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=96"
        If Not rs.EOF Then gtype_UserSysParms.P96_药品填单下可用库存 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=126"
        If Not rs.EOF Then gtype_UserSysParms.P126_时价药品售价加成方式 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=149"
        If Not rs.EOF Then gtype_UserSysParms.P149_效期显示方式 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=150"
        If Not rs.EOF Then gtype_UserSysParms.P150_药品出库优先算法 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=173"
        If Not rs.EOF Then gtype_UserSysParms.P173_经过标记付款后才能进行付款管理 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=174"
        If Not rs.EOF Then gtype_UserSysParms.P174_药品移库明确批次 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=175"
        If Not rs.EOF Then gtype_UserSysParms.P175_药品领用明确批次 = Val(Nvl(rs!参数值, rs!缺省值))
        
'        rs.Filter = "参数号=181"
'        If Not rs.EOF Then gtype_UserSysParms.P181_药品入库按分段加成 = Val(Nvl(rs!参数值, rs!缺省值))
'
'        rs.Filter = "参数号=183"
'        If Not rs.EOF Then gtype_UserSysParms.P183_时价取上次售价 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=221"
        If Not rs.EOF Then gtype_UserSysParms.P221_药品结存时点 = Val(Nvl(rs!参数值, rs!缺省值))
        
        rs.Filter = "参数号=306"
        If Not rs.EOF Then gtype_UserSysParms.P306_已失效药品禁止入库 = Val(Nvl(rs!参数值, rs!缺省值))
        
    ElseIf rs.RecordCount = 0 Then
        strMsg = "请先设置系统参数。"
    End If
    
    '取药品最大允许精度
    gstrSQL = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "取药品精度")
    gtype_UserDrugDigits.Digit_金额 = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_成本价 = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_零售价 = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_数量 = rs.Fields(3).NumericScale
    
    '取药品售价单位小数位数
    gstrSQL = "Select 内容, Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = 0 And 类别 = 1 And 单位 = 1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, "取药品售价单位小数位数")
    
    If rs.RecordCount > 0 Then
        rs.Filter = "内容=1"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_成本价 = rs!精度
        
        rs.Filter = "内容=2"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_零售价 = rs!精度
        
        rs.Filter = "内容=3"
        If Not rs.EOF Then gtype_UserSaleDigits.Digit_数量 = rs!精度
        
        If gtype_UserSaleDigits.Digit_成本价 < 2 Or gtype_UserSaleDigits.Digit_成本价 > gtype_UserDrugDigits.Digit_成本价 Then
            gtype_UserSaleDigits.Digit_成本价 = gtype_UserDrugDigits.Digit_成本价
        End If
        
        If gtype_UserSaleDigits.Digit_零售价 < 2 Or gtype_UserSaleDigits.Digit_零售价 > gtype_UserDrugDigits.Digit_零售价 Then
            gtype_UserSaleDigits.Digit_零售价 = gtype_UserDrugDigits.Digit_零售价
        End If
        
        If gtype_UserSaleDigits.Digit_数量 < 2 Or gtype_UserSaleDigits.Digit_数量 > gtype_UserDrugDigits.Digit_数量 Then
            gtype_UserSaleDigits.Digit_数量 = gtype_UserDrugDigits.Digit_数量
        End If
    End If
    
    '药品名称显示方式
    gint药品名称显示 = Val(zlDataBase.GetPara("药品名称显示", , , 2))
    gint输入药品显示 = Val(zlDataBase.GetPara("输入药品显示"))
    
    If gint药品名称显示 < 0 Or gint药品名称显示 > 2 Then gint药品名称显示 = 2
    If gint输入药品显示 < 0 Or gint输入药品显示 > 1 Then gint输入药品显示 = 0
    
    '简码方式
    gint简码方式 = Val(zlDataBase.GetPara("简码方式"))
    If gint简码方式 < 0 Or gint简码方式 > 1 Then gint简码方式 = 0
    
    Exit Sub
errh:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。
Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
    Static lngPreTime As Single, lngPreHwnd As Long
    Static StrFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: StrFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            StrFind = ""
        End If
        StrFind = StrFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal StrFind)
        If MatchIndex = 0 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
    If MatchIndex = -1 Then MatchIndex = 1
End Function

'返回指定库房指定适用范围的单位
Public Function GetSpecUnit(ByVal lng库房ID As Long, ByVal int范围 As Integer) As String
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(性质,1) AS 单位 From 药品库房单位 Where 库房ID=[1] And 适用范围=[2]"
    Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "提取单位", lng库房ID, int范围)

    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!单位
    Else
'        MsgBox "该库房未设置库房单位，根据部门性质以及服务对象取缺省单位！" & _
'            vbCrLf & "缺省单位的规则：" & _
'            vbCrLf & "  服务对象是住院或门诊和住院的，取住院单位" & _
'            vbCrLf & "  仅服务于门诊的，取门诊单位" & _
'            vbCrLf & "  具有药库属性的，取药库单位" & _
'            vbCrLf & "  其他取售价单位", vbInformation, gstrSysName
        
        gstrSQL = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "读取药品单位", lng库房ID)

        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '住院单位
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '门诊单位
            strUnit = 2
        ElseIf InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            strUnit = 4
        Else
            '售价单位：主要是制剂室
            strUnit = 1
        End If
    End If
    
    '转换为真实的单位返回给调用者
    GetSpecUnit = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    If glngSys / 100 = 8 Then
        '药店只有售价单位与药库单位
        GetSpecUnit = IIf(strUnit = 1, "售价单位", "药库单位")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'取药品单位名称
Public Function GetDrugUnit(ByVal lng库房ID As Long, ByVal frmCaption As String) As String
    Dim rsProperty As New Recordset
    Dim strobjTemp As String                    '保存服务对象字符串
    Dim strWorkTemp As String                   '保存工作性质字符串
    Dim intUnit As Integer, strUnit As String
    Dim bln缺省 As Boolean
    Dim lngModul As Long
    On Error GoTo ErrHand
    
    If frmCaption Like "药品外购入库管理*" Then
        lngModul = 1300
    ElseIf frmCaption Like "药品自制入库管理*" Then
        lngModul = 1301
    ElseIf frmCaption Like "药品其他入库管理*" Then
        lngModul = 1302
    ElseIf frmCaption Like "库存差价调整管理*" Then
        lngModul = 1303
    ElseIf frmCaption Like "药品移库管理*" Then
        lngModul = 1304
    ElseIf frmCaption Like "药品领用管理*" Then
        lngModul = 1305
    ElseIf frmCaption Like "药品其他出库管理*" Then
        lngModul = 1306
    ElseIf frmCaption Like "药品盘点管理*" Then
        lngModul = 1307
    ElseIf frmCaption Like "药品差价计算*" Then
        lngModul = 1308
    ElseIf frmCaption Like "药品计划管理*" Or frmCaption Like "药品采购计划*" Then
        lngModul = 1330
    ElseIf frmCaption Like "药品质量管理*" Then
        lngModul = 1331
    ElseIf frmCaption Like "药品申领管理*" Then
        lngModul = 1343
    End If
    
    intUnit = 0
    '如果是申领单，则直接返回注册表中的单位
    If lngModul > 0 And lngModul <> 1331 And lngModul <> 1307 And lngModul <> 1308 Then
        intUnit = Val(zlDataBase.GetPara("药品单位", glngSys, lngModul))
        '本地参数设置的单位顺序如下：0-缺省;1-药库;2-门诊;3-住院;4-售价，需要转换为与系统参数的一致
        If intUnit = 1 Then
            intUnit = 4
        ElseIf intUnit = 4 Then
            intUnit = 1
        End If
        strUnit = intUnit
    End If
    
    If intUnit = 0 Then
        gstrSQL = "SELECT distinct 服务对象,工作性质 From 部门性质说明 Where 部门ID =[1]"
        Set rsProperty = zlDataBase.OpenSQLRecord(gstrSQL, "读取药品单位", lng库房ID)

        '取服务对象及部门性质
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strWorkTemp, "药库") <> 0 Then
            '药库单位
            intUnit = 1
            strUnit = 4
        ElseIf InStr(strobjTemp, "1") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            '门诊单位
            intUnit = 2
            strUnit = 2
        ElseIf InStr(strobjTemp, "2") <> 0 Then
            '住院单位
            intUnit = 3
            strUnit = 3
        Else
            '售价单位：主要是制剂室
            intUnit = 4
            strUnit = 1
        End If
        
        '取该药房缺省该使用的单位
        GetDrugUnit = GetSpecUnit(lng库房ID, intUnit)
    Else
        GetDrugUnit = Switch(strUnit = 1, "售价单位", strUnit = 2, "门诊单位", strUnit = 3, "住院单位", strUnit = 4, "药库单位")
    End If
    
    '转换为真实的单位返回给调用者
    
    If glngSys / 100 = 8 Then
        '药店只有售价单位与药库单位
        GetDrugUnit = IIf(strUnit = 1, "售价单位", "药库单位")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDrugUnit = "售价单位"
End Function
'取数值的小数位数
Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function


Public Function MediWork_GetCheckStockRule(ByVal lng库房ID As Long) As Integer
    '取出库检查规则
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "取出库检查规则", lng库房ID)

    If Not rsData.EOF Then
        MediWork_GetCheckStockRule = rsData!库存检查
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get成本价(ByVal lng药品ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long) As Double
'功能：获取当前药品的成本价格
'参数：药品id,库房id,批次
'返回值： 成本价格
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo errHandle
    
    gstrSQL = "select 平均成本价 from 药品库存 where 性质=1 and 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "成本价", lng药品ID, lng库房ID, lng批次)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!平均成本价) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!平均成本价) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get成本价 = rsData!平均成本价
    Else
        '如果无法从库存中取成本价，则从药品规格中取
        gstrSQL = "select 成本价 from 药品规格 where 药品id=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "成本价", lng药品ID)
        If Not rsData.EOF Then
            If Val(Nvl(rsData!成本价, 0)) > 0 Then
                Get成本价 = rsData!成本价
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

'按编码，名称，别名查找某一列
Public Function FindRow(ByVal mshBill As BillEdit, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo errHandle
    FindRow = True
    With mshBill
        If .rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(strCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .MsfObj.TopRow = .Row
                    .SetRowColor CLng(intRow), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = " SELECT DISTINCT b.编码 " & _
                  " FROM " & _
                  "    (SELECT DISTINCT A.收费细目id " & _
                  "    FROM 收费项目别名 A" & _
                  "    Where A.简码 LIKE [1]) a," & _
                  " 收费项目目录 B " & _
                  " Where a.收费细目id = b.ID"
        Set rsCode = zlDataBase.OpenSQLRecord(gstrSQL, "查找指定药品", IIf(gstrMatchMethod = "0", "%", "") & str比较值 & "%")
        
        If rsCode.EOF Then
            FindRow = False
            Exit Function
        End If
        
        For intRow = intStartRow To .rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .MsfObj.TopRow = .Row
                        .SetRowColor CLng(intRow), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            
            End If
        Next
        rsCode.Close
    End With
    FindRow = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


'功能：根据所属部门返回所能访问的材质
'返回：如'西成药'；'中成药',空表示所有
Public Function GetMaterial(lngUnitID As Long) As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(gstrprivs, "所有药房") > 0 Then Exit Function
    
    On Error GoTo errh
    
    rsTmp.CursorLocation = adUseClient

    gstrSQL = "Select A.工作性质,B.名称 From 部门性质说明 A,部门表 B Where A.部门ID=B.ID And B.ID=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "取指定部门的工作性质", lngUnitID)
    
    If rsTmp.EOF Then Exit Function
    
    rsTmp.Filter = "工作性质='西药房' or 工作性质='西药库' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'西成药'"
    
    rsTmp.Filter = "工作性质='成药房' or 工作性质='成药库' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'中成药'"
    
    rsTmp.Filter = "工作性质='中药房' or 工作性质='中药库' "
    If Not rsTmp.EOF Then GetMaterial = GetMaterial & ",'中草药'"
    
    GetMaterial = Mid(GetMaterial, 2)
    Exit Function
errh:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Public Function ExecuteSql(ByRef arrSql As Variant, strTitle As String, _
Optional ByVal blnCommit As Boolean = True, Optional ByVal blnBeginTrans As Boolean = True) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer
    Dim intouter As Integer
    Dim intInner As Integer
    
    ExecuteSql = False
    If UBound(arrSql) >= 0 Then
        '对SQL序列按药品ID升序排序
        intouter = UBound(arrSql) - 1
        If Split(arrSql(UBound(arrSql)), ":")(0) = "出库" Then
            intouter = UBound(arrSql) - 2
        Else
            intouter = UBound(arrSql) - 1
        End If
        
        For i = 0 To intouter
            For j = i + 1 To intouter + 1
                If CLng(Split(arrSql(j), ";")(0)) < CLng(Split(arrSql(i), ";")(0)) Then
                    strTmp = CStr(arrSql(j))
                    arrSql(j) = arrSql(i)
                    arrSql(i) = strTmp
                End If
            Next
        Next
        
        '执行SQL语句
        On Error GoTo errh
        If blnBeginTrans Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(Mid(arrSql(i), InStr(arrSql(i), ";") + 1)), strTitle)
        Next
        If blnCommit Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
       
errh:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


'取指定列头的列位置
Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    Dim i As Integer
    
    GetCol = -1
    
    If TypeName(mshFlex) = "MSHFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    ElseIf TypeName(mshFlex) = "VSFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    End If
End Function

'获取指定表特定字段的长度
Public Function GetLength(ByVal strTable As String, ByVal strColumn As String) As Integer
    Dim rsPar As New ADODB.Recordset
        
    On Error Resume Next
    
    Err = 0
    GetLength = 0
    
    With rsPar
        gstrSQL = "Select " & strColumn & " From " & strTable & " Where Rownum<1"
        Call zlDataBase.OpenRecordset(rsPar, gstrSQL, "取字段长度")
        
        If Err <> 0 Then
            MsgBox "数据表[" & strTable & "]不存在，请与开发商联系！", vbInformation, gstrSysName
        End If
        GetLength = .Fields(0).DefinedSize
        .Close
    End With
End Function

'根据药品流向控制表的数据，提取对方库房
'Writed by zyb
'-----------------调拨-----------------
'所在库房是当前库房的，提取流向 In (1"可流向对方库房",3"可双向流通")
'对方库房是当前库房的，提取流向 IN (2"可流向所在库房",3"可双向流通")
'-----------------申领-----------------
'所在库房是当前库房的，提取流向 In (2"可流向所在库房",3"可双向流通")
'对方库房是当前库房的，提取流向 IN (1"可流向对方库房",3"可双向流通")
Public Function ReturnSQL(ByVal lng库房ID As Long, ByVal strCaption As String, _
    Optional ByVal bln调拨 As Boolean = True, _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    
    Dim str库房性质 As String, str药品流向 As String, str站点限制 As String, strSQL As String
    
    On Error GoTo errHandle
    str站点限制 = GetDeptStationNode(lng库房ID)
    str库房性质 = "('H','I','J','K','L','M','N')"
    
    str药品流向 = ",(Select 对方库房ID ID From 药品流向控制" & _
            " Where 所在库房ID=[1] And 流向 In (" & IIf(bln调拨, 1, 2) & ",3)" & _
            " Union" & _
            " Select 所在库房ID ID From 药品流向控制" & _
            " Where 对方库房ID=[1] And 流向 In (" & IIf(bln调拨, 2, 1) & ",3)) D"
    Select Case lngModuleNO
        Case 1304   '药品移库管理
            strSQL = " SELECT DISTINCT a.id,a.编码,a.名称" & _
                    " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a" & str药品流向 & _
                    " Where c.工作性质 = b.名称" & _
                    "   AND b.编码||'' in " & str库房性质 & _
                    "   AND a.id = c.部门id And A.ID=D.ID " & _
                    "   AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.编码"
        Case Else
            strSQL = " SELECT DISTINCT a.id,a.编码,a.名称" & _
                    " FROM 部门性质说明 c, 部门性质分类 b, 部门表 a" & str药品流向 & _
                    " Where c.工作性质 = b.名称" & _
                    "   AND b.编码||'' in " & str库房性质 & _
                    "   AND a.id = c.部门id And A.ID=D.ID" & IIf(str站点限制 <> "", " AND (a.站点=[2] or a.站点 is null) ", "") & _
                    "   AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
                    " Order by a.编码"
    End Select
    
    Set ReturnSQL = zlDataBase.OpenSQLRecord(strSQL, strCaption, lng库房ID, str站点限制)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 相同符号(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_负数 As Boolean, blnSecond_负数 As Boolean
    
    相同符号 = False
    
    If sinFirst = 0 Or sinSecond = 0 Then '0无正负号之分
        相同符号 = True
        Exit Function
    End If
    
    blnFirst_负数 = (sinFirst <= 0)
    blnSecond_负数 = (sinSecond <= 0)
    
    相同符号 = (blnFirst_负数 = blnSecond_负数)
End Function

'从指定行开始更新序号
Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng序号列 As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    
    With mshBill
        lngRows = .rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng序号列) = lngRow
        Next
    End With
End Sub



'转换数值为日期
Public Function TranNumToDate(ByVal strNum As String) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
End Function

'获取指定窗体的父窗体
Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function

'获取指定窗体的标题
Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    
    On Error Resume Next
   
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function

Public Function CheckLapse(ByVal str效期 As String, Optional bln已失效药品禁止入库 As Boolean = False) As Boolean
    '失效药品检查
    CheckLapse = True
    
    If Not IsDate(str效期) Then Exit Function
    
    If gtype_UserSysParms.P149_效期显示方式 = 1 Then
        '换算为失效期
        str效期 = Format(DateAdd("D", 1, CDate(str效期)), "yyyy-mm-dd")
    End If
    
    If Format(str效期, "yyyy-MM-dd") <= Format(zlDataBase.Currentdate, "yyyy-MM-dd") Then
        If bln已失效药品禁止入库 = False Then MsgBox "该药品已经失效了！", vbInformation, gstrSysName
        
        If gtype_UserSysParms.P306_已失效药品禁止入库 = 1 And bln已失效药品禁止入库 = True Then
            CheckLapse = False
        End If
    End If
    
End Function

'药品单据审核时，是否判断审核人与填制人，其返回审核结果
Public Function 药品单据审核(ByVal str填制人 As String) As Boolean
    Dim blnBillVerify As Boolean
    
    药品单据审核 = True
    
    blnBillVerify = IIf(gtype_UserSysParms.P64_审核限制 = 0, False, True)
    If Not blnBillVerify Then Exit Function
    
    药品单据审核 = (Trim(str填制人) <> Trim(UserInfo.用户姓名))
    If Not 药品单据审核 Then MsgBox "填制人与审核人不能是同一人，请检查！", vbInformation, gstrSysName
End Function
'通过药品选择器输入药品时，如果药品库存中的数据与从部门性质、药品目录中的分批属性判断出的不一致，则报错
Public Function 检查库存数据(ByVal lng库房ID As Long, ByVal lng药品ID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim bln库存是否分批 As Boolean, bln分批 As Boolean, bln库房 As Boolean
    
    检查库存数据 = False
    On Error GoTo errHandle
    '如果没有库存记录，则直接退出
    gstrSQL = " Select Count(*) 记录数 From 药品库存 " & _
              " Where 库房ID=[1] And 性质=1 And 药品ID=[2]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "是否存在库存数据", lng库房ID, lng药品ID)
    
    If rsCheck!记录数 = 0 Then
        检查库存数据 = True
        Exit Function
    End If
    
    '存在分批记录则表明分批
    gstrSQL = " Select Count(*) 分批 From 药品库存 " & _
              " Where 库房ID=[1] And 性质=1 And Nvl(批次,0)<>0 And 药品ID=[2]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "检查库存数据", lng库房ID, lng药品ID)
              
    bln库存是否分批 = (rsCheck!分批 <> 0)
    
    '先判断是否是库房
    gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "取部门性质", lng库房ID)

    bln库房 = (rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    gstrSQL = " Select Nvl(药库分批,0) 分批核算,nvl(药房分批,0) 药房分批核算 " & _
              " From 药品规格 Where 药品ID=[1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "取药品目录中的分批属性", lng药品ID)
              
    If bln库房 Then
        bln分批 = (rsCheck!分批核算 = 1)
    Else
        bln分批 = (rsCheck!药房分批核算 = 1)
    End If
    
    检查库存数据 = (bln库存是否分批 = bln分批)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'检查药品的价格是否为最新的价格（按药库单位进行比较，时价不分批药品不检查），允许继续操作
'由于在保存前判断很麻烦，且各种单据的表格中保存的数据不一样，因此，待保存完成之后且提交前对已保存的数据进行检查
'药品相同的记录略过
Public Function 检查单价(ByVal lng单据 As Long, ByVal strNo As String, Optional ByVal blnMsg As Boolean = True, Optional ByVal bln移库单 As Boolean = False) As Boolean
    Dim rsPrice As New ADODB.Recordset
    Dim lng药品_Last As Long, lng药品_Cur As Long
    Dim intPriceDigit As Integer
    Dim intCostDigit As Integer
             
    On Error GoTo errHandle
    '自动批量检查并执行调价
    Call AutoAdjustPrice_ByNO(lng单据, strNo)
    
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
    
    '定价药品从收费价目取最新价格；时价分批药品从库存表取最新价格（时价药品调价是按库存调整的，如果无库存则表示无调价，而且严格控制库存的条件下无库存也不能允许出库）
        
    gstrSQL = " Select '售价' As 类型, a.序号, a.药品id , 0 原价, b.现价" & _
            " From 药品收发记录 A," & _
                 " (Select 收费细目id, Nvl(现价, 0) 现价, 执行日期" & _
                   " From 收费价目" & _
                   " Where (终止日期 Is Null Or Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'yyyy-MM-dd')))) B, 收费项目目录 C" & _
            " Where a.单据 = [1] And a.No = [2] And a.药品id = b.收费细目id And c.Id = b.收费细目id And Round(a.零售价," & intPriceDigit & ") <> Round(b.现价, " & intPriceDigit & ") And" & _
              "    NVL(c.是否变价, 0) = 0 and b.执行日期>a.填制日期" & _
            " Union All" & _
            " Select '售价' As 类型, a.序号, a.药品id , 0 原价, decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价) As 现价" & _
            " From 药品收发记录 A, 药品库存 B, 收费项目目录 C" & _
            " Where a.单据 = [1] And a.No = [2] And c.Id = a.药品id And Round(a.零售价," & intPriceDigit & ") <> Round(decode(nvl(b.批次,0),0,b.实际金额 / b.实际数量,b.零售价), " & intPriceDigit & ") And Nvl(c.是否变价, 0) = 1 And" & _
                  " b.性质 = 1 And b.库房id = a.库房id And b.药品id = a.药品id And NVL(b.批次, 0) = NVL(a.批次, 0) And NVL(b.实际数量, 0) <> 0 And a.入出系数 = -1" & _
            " Union All" & _
            " Select '成本价' As 类型, a.序号, a.药品id , 0 原价, b.平均成本价 As 现价" & _
            " From 药品收发记录 A, 药品库存 B" & _
            " Where a.单据 = [1] And a.No = [2] And a.药品id = b.药品id And Nvl(a.批次, 0) = Nvl(b.批次, 0) and round(a.成本价," & intCostDigit & ")<>round(b.平均成本价," & intCostDigit & ") And a.库房id = b.库房id and a.入出系数=-1 and b.性质=1" & _
            " Order By 类型, 药品id, 序号"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "取当前价格", lng单据, strNo)
    
    If rsPrice.EOF Then
        检查单价 = True
        Exit Function
    End If
    
    lng药品_Last = 0
    With rsPrice
        Do While Not .EOF
            lng药品_Cur = !药品id
            If lng药品_Cur <> lng药品_Last Then
                If blnMsg Then
                    If MsgBox("第" & IIf(bln移库单, Round(!序号 / 2 + 0.49), !序号) & "行药品的" & !类型 & "不是最新价格，是否继续保存单据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lng药品_Last = lng药品_Cur
            .MoveNext
        Loop
        检查单价 = True
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'------------------------------------------------
'功能： 密码转换函数
'参数：
'   strOld：原密码
'返回： 加密生成的密码
'------------------------------------------------
Public Function TranPasswd(strOld As String) As String
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function GetBillInfo(ByVal lng单据 As Long, ByVal strNo As String, Optional ByVal bln填制日期 As Boolean = True) As String
    Dim rsBillInfo As New ADODB.Recordset
    
    On Error GoTo errHandle
    '获取单据的最大修改时间
    gstrSQL = " Select to_char(Max(" & IIf(bln填制日期, "填制日期", "审核日期") & "),'yyyyMMddhh24miss') 日期 From 药品收发记录 " & _
            " Where 单据=[1] And NO=[2]"
    Set rsBillInfo = zlDataBase.OpenSQLRecord(gstrSQL, "获取单据的最大修改时间", lng单据, strNo)
    
    With rsBillInfo
        '返回空，表示已经删除
        If .EOF Then Exit Function
        If IsNull(!日期) Then Exit Function
        GetBillInfo = !日期
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'取药品金额、价格和数量的小数位数
Public Function GetDigit(ByVal int性质 As Integer, ByVal int类别 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int性质：0-计算精度;
    'int类别：1-药品;2-卫材
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         药品单位:1-售价;2-门诊;3-住院;4-药库;
    '         卫材单位:1-散装;2-包装
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHand
    If int内容 = 4 Then '取金额 单位=5的才是金额
        int单位 = 5
    End If
    
    gstrSQL = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 性质 = [1] And 类别 = [2] And 内容 = [3] And 单位 = [4] "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int性质, int类别, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!精度
    End If
    
    If GetDigit = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        GetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int内容, gtype_UserDrugDigits.Digit_成本价, gtype_UserDrugDigits.Digit_零售价, gtype_UserDrugDigits.Digit_数量, gtype_UserDrugDigits.Digit_金额)
End Function

'根据库房的包装单位来取药品的价格、数量、金额小数位数（计算精度）
Public Sub GetDrugDigit(ByRef lng库房ID As Long, ByVal frmCaption As String, ByRef intUnit As Integer, ByRef intCostDigit As Integer, ByRef intPriceDigit As Integer, ByRef intNumberDigit As Integer, ByRef intMoneyDigit As Integer)
    Dim strUnit As String
    Dim intTemp As Integer
    
    Const conInt精度 As Integer = 0
    
    Const conInt药品 As Integer = 1
    
    Const conint售价单位 As Integer = 1
    Const conint门诊单位 As Integer = 2
    Const conint住院单位 As Integer = 3
    Const conint药库单位 As Integer = 4
        
    Const conInt成本价 As Integer = 1
    Const conInt售价 As Integer = 2
    Const conInt数量 As Integer = 3
    Const conInt金额 As Integer = 4
    
    If lng库房ID > 0 Then
        If frmCaption Like "药品验收管理*" Then
            strUnit = conint药库单位
        Else
            strUnit = GetDrugUnit(lng库房ID, frmCaption)
        
            Select Case strUnit
                Case "售价单位"             '售价单位：主要是制剂室
                    intUnit = conint售价单位
                Case "门诊单位"
                    intUnit = conint门诊单位
                Case "住院单位"
                    intUnit = conint住院单位
                Case "药库单位"
                    intUnit = conint药库单位
            End Select
        End If
    Else
        
        If frmCaption Like "药品计划管理*" Or frmCaption Like "药品采购计划*" Then
            intTemp = Val(zlDataBase.GetPara("药品单位", glngSys, 1330))
            Select Case intTemp
            Case 1 '药库
                intUnit = conint药库单位
            Case 2  '门诊
                intUnit = conint门诊单位
            Case 3  '住院
                intUnit = conint住院单位
            Case 4  '售价
                intUnit = conint售价单位
            Case Else
                intUnit = conint药库单位
            End Select
        Else
            intUnit = conint药库单位
        End If
    End If

    '分别取药品成本价、售价、数量、金额的小数位数
    intCostDigit = GetDigit(conInt精度, conInt药品, conInt成本价, intUnit)
    intPriceDigit = GetDigit(conInt精度, conInt药品, conInt售价, intUnit)
    intNumberDigit = GetDigit(conInt精度, conInt药品, conInt数量, intUnit)
    intMoneyDigit = GetDigit(conInt精度, conInt药品, conInt金额)

End Sub
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant

    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可
    '返回:初如成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rstemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Select部门选择器(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln操作员 As Boolean = False, _
    Optional strSQL As String = "") As Boolean
    '------------------------------------------------------------------------------
    '功能:部门选择器
    '参数:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln操作员-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strkey As String, strTittle As String, lngH As Long, StrFind As String
    Dim vRect As RECT
    Dim rstemp  As ADODB.Recordset
    Dim strPa As String
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "部门选择器"
    vRect = GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strkey = GetMatchingSting(strSearch, False)
    
    strPa = zlDataBase.GetPara(44, glngSys, 0): strPa = IIf(strPa = "", "11", strPa)
    
    If strSQL <> "" Then
    
        gstrSQL = strSQL
    Else
        gstrSQL = "" & _
        "   Select distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
        "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
    
        If str工作性质 = "" And bln操作员 = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c" & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.编码 in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) ") & _
            "         AND a.id = c.部门id " & _
            IIf(bln操作员 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) And (a.站点=[4] or a.站点 is null) "
    End If
    
    StrFind = ""
    If strSearch <> "" Then
        StrFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            If Mid(strPa, 1, 1) = "1" Then StrFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            If Mid(strPa, 2, 1) = "1" Then StrFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            StrFind = " And a.名称 Like [3] "
        End If
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSQL = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSQL = gstrSQL & vbCrLf & StrFind & vbCrLf & " Order by A.编码"
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSQL = "" Then
        '分上下级
        Set rstemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strkey)
    Else
        Set rstemp = zlDataBase.ShowSQLSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.用户ID, str工作性质, strkey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rstemp Is Nothing Then
        MsgBox "没有满足条件的部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlCtlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rstemp!id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            MsgBox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = Nvl(rstemp!编码) & "-" & Nvl(rstemp!名称)
        objCtl.Tag = Val(rstemp!id)
    End If
    zlCommFun.PressKey vbKeyTab
    Select部门选择器 = True
End Function

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub CostPrice()
    '是否允许药房人员查看单据的成本价
    mblnCostPrice = IIf(gtype_UserSysParms.P85_药房查看单据成本价 = 1, True, False)
End Sub

Public Function DepotProperty(ByVal lng人员id As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHandle
    '返回指定人员是否具有药库性质
    gstrSQL = "Select Distinct 工作性质 From 部门人员 B,部门性质说明 A " & _
             " Where A.工作性质 like '%药库' And " & _
             " A.部门id = B.部门id And B.人员id = [1]"
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, "取部门性质", lng人员id)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ShowCostPrice() As Boolean
    '药库人员不管，只管药房人员，以参数控制为准
    Call CostPrice
    If DepotProperty(UserInfo.用户ID) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = mblnCostPrice
    End If
End Function

'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function CheckNOExists(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 药品收发记录 Where NO=[1] And 单据=[2] And Rownum<2"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "判断是否存在该单据", strNo, int单据)
    
    If rstemp.RecordCount = 0 Then Exit Function
    CheckNOExists = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'判断该药品在当前库存的库存是否低于库存下限，是则返回真
Public Function IsLowerLimit(ByVal lng库房ID As Long, ByVal lng药品ID As Long) As Boolean
    Dim dbl库存数量 As Double, dbl下限 As Double
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '提取库存数量
    gstrSQL = " Select Sum(Nvl(实际数量,0)) AS 库存数量 From 药品库存" & _
              " Where 性质=1 And 库房ID=[1] And 药品ID=[2]"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "提取指定库房的实际库存", lng库房ID, lng药品ID)
    
    If rstemp.RecordCount = 1 Then dbl库存数量 = Nvl(rstemp!库存数量, 0)
    
    '提取储备限额中的下限
    gstrSQL = " Select Nvl(下限,0) AS 下限 From 药品储备限额" & _
              " Where 库房ID=[1] And 药品ID=[2]"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "提取储备限额中的下限", lng库房ID, lng药品ID)
    
    If rstemp.RecordCount = 1 Then dbl下限 = rstemp!下限
    
    IsLowerLimit = (dbl库存数量 < dbl下限)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'简码方式
'staVal: StartusBar控件
'bytType: 0=拼音; 1=五笔;  当前简码状态
    Dim i As Integer
    For i = 1 To staVal.Panels.count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDataBase.SetPara "简码方式", 0
                gint简码方式 = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDataBase.SetPara "简码方式", 1
                gint简码方式 = 1
            End If
        End If
    Next
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'获取部门所属站点信息
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    On Error GoTo errHandle
    strTmp = "select 站点 from 部门表 where id=[1]"
    Set rsSQL = zlDataBase.OpenSQLRecord(strTmp, "获取部门所属站点信息", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = Nvl(rsSQL!站点)
    End If
    rsSQL.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetVSFlexRows(ByVal vsfVal As VSFlexGrid, Optional ByVal blnHidden = False) As Long
'--------------------------------------------------------------
'功能：求VSFlexGrid的行数量，含列头行
'参数：
'  blnHidden：True计算非隐藏的行数；False计算隐藏的行数。
'返回：行数量
'--------------------------------------------------------------
    Dim i As Long, lngRows As Long
    For i = 0 To vsfVal.rows - 1
        If blnHidden Then
            If vsfVal.RowHidden(i) Then lngRows = lngRows + 1
        Else
            If vsfVal.RowHidden(i) = False Then lngRows = lngRows + 1
        End If
    Next
    GetVSFlexRows = lngRows
End Function

Public Sub SetSelectorRS( _
    ByVal byt编辑模式 As Byte, _
    ByVal strModeName As String, _
    Optional ByVal lng来源库房 As Long = 0, _
    Optional ByVal lng目标库房 As Long = 0, _
    Optional ByVal lng使用部门 As Long = 0, _
    Optional ByVal lng供应商 As Long = 0, _
    Optional ByVal byt领用方式 As Byte = 0, _
    Optional ByVal bln包含停用药品 As Boolean = False, _
    Optional ByVal bln盘无存储库房药品 As Boolean = False, _
    Optional ByVal byt盘点单据 As Byte = 0, _
    Optional ByVal bln检测库存 As Boolean = True, _
    Optional ByVal bln调价 As Boolean = False, _
    Optional ByVal bln忽略服务对象 As Boolean = True _
    )
'----------------------------------------------------------------------------------------
'功能：初始化grsMaster、grsMasterInput、grsSlave对象，
'      为调用药品选择器(frmSelector)做数据准备。
'参数：
'  byt编辑模式： 1：入库； 2：出库
'  lng来源库房：
'----------------------------------------------------------------------------------------
    Const CON_FMT = "'999999999990.99999'"
    
    Dim strSQL As String, strTmp As String
    Dim strUnit As String, strConversionUnit As String
    Dim rstemp As ADODB.Recordset
    Dim IntStockCheck As Integer
    Dim intUnit As Integer, intCostDigit As Integer, intPriceDigit As Integer, intNumberDigit As Integer, intMoneyDigit As Integer
    
    On Error GoTo errHandle
    With grsMaster
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsMasterInput
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    With grsSlave
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic     'adOpenStatic
        .LockType = adLockReadOnly
    End With
    
    '数量单位
    If strModeName = "药品申领管理" Or strModeName = "药品移库管理" Then
        Call GetDrugDigit(lng使用部门, strModeName, intUnit, intCostDigit, intPriceDigit, intNumberDigit, intMoneyDigit)
    Else
        Call GetDrugDigit(IIf(lng来源库房 = 0, lng目标库房, lng来源库房), strModeName, intUnit, intCostDigit, intPriceDigit, intNumberDigit, intMoneyDigit)
    End If
    Select Case intUnit
        Case 1: strConversionUnit = "1"
        Case 2: strConversionUnit = "d.门诊包装"
        Case 3: strConversionUnit = "d.住院包装"
        Case Else
            strConversionUnit = "d.药库包装"
    End Select
    
    '检查并执行调价
    Call AutoAdjustPrice_Batch
    
    '提取库存检查参数，确定库存不足的不提取数据
    strSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1] "
    Set rstemp = zlDataBase.OpenSQLRecord(strSQL, "获取是否库存检查设置", lng来源库房)
    If Not rstemp.EOF Then IntStockCheck = Nvl(rstemp!库存检查, 0)
    rstemp.Close
    
    '*选择模式的数据集*'
    strSQL = _
        "Select " & _
        " d.剂型,d.中药形态, d.药名编码, d.通用名称, d.药品来源 As 来源, d.基本药物, d.药典id, d.用途分类id, d.剂量单位, d.药品编码, d.药品名称, " & _
        " d.商品名, d.规格, d.产地 As 产地, d.药名id, d.药品id, " & _
        " trim(to_char(d.初始成本价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) 上次采购价, " & _
        " trim(to_char(Decode(d.时价, '是', Decode(s.平均售价, Null, p.售价, s.平均售价), p.售价) * " & strConversionUnit & ", '99999999999990." & String(intPriceDigit, "0") & "')) 售价, " & _
        " d.售价单位, d.剂量系数 As 售价包装," & _
        " d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, " & _
        " trim(to_char(s.可用数量 / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')) 可用数量, " & _
        " s.库存数量, s.库存金额, s.库存差价,  d.最大效期 有效期, d.药库分批, d.药房分批, d.时价," & _
        " trim(to_char(d.指导批发价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导批发价, " & _
        " trim(to_char(d.指导零售价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导零售价, " & _
        " d.指导差价率, e.库房货位, d.批准文号, s.库存数量 实际数量, " & _
        " s.留存数量, d.合同单位, d.药价级别,e.领用标志,d.上次供应商 " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.名称 剂型,Decode(c.类别, '7', Decode(d.中药形态, 1, '饮片', 2, '免煎剂', '散装'), '') As 中药形态,A.名称 商品名, C.编码 药名编码,C.名称 通用名称, 0 AS 药典ID,C.编码 药品编码,C.名称 药品名称," & vbNewLine & _
        "     C.规格,C.产地,C.类别,C.计算单位 AS 售价单位,DECODE(C.是否变价,1,'是','否') 时价,D.药品来源,D.基本药物,D.批准文号, D.药名ID," & vbNewLine & _
        "     D.药品ID, nvl(to_char(D.最大效期,'9999990'),0) 最大效期," & vbNewLine & _
        "     DECODE(D.药库分批,1,'是','否') 药库分批,DECODE(D.药房分批,1,'是','否') 药房分批," & vbNewLine & _
        "     to_char(D.剂量系数, " & CON_FMT & ") 剂量系数," & vbLf & _
        "     D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & vbNewLine & _
        "     D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & vbNewLine & _
        "     D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & vbNewLine & _
        "     D.指导批发价,d.指导零售价, nvl(D.成本价,0) 初始成本价,D.指导差价率,D.药价级别," & vbNewLine & _
        "     M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位,Q.名称 As 合同单位,f.名称 上次供应商 " & vbNewLine
    strSQL = strSQL & _
        "   FROM 收费项目目录 C,药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M,供应商 Q, 诊疗分类目录 E,供应商 F" & vbNewLine
        
    If bln调价 = False Then
        strSQL = strSQL & IIf(lng来源库房 <> 0, " ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K", "") & vbNewLine & _
        IIf(lng目标库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[3] Group By 执行科室ID,收费细目ID) I ", "") & vbNewLine
    End If
    strSQL = strSQL & "   WHERE C.ID=D.药品ID AND D.药名ID=T.药名ID AND T.药名ID=M.ID and m.分类id=e.id AND M.类别 IN ('5','6','7') and t.临床自管药 is null And D.上次供应商ID=f.ID(+) "
    
    If bln调价 = False Then
        strSQL = strSQL & IIf(lng来源库房 <> 0, "     And D.药品ID=K.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        IIf(lng目标库房 <> 0, "     And D.药品ID=I.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "")
    End If
    
    If bln忽略服务对象 = False Then
        strSQL = strSQL & " And" & _
             " (Decode(c.服务对象, 1, 1, 3, 1, 0) = " & _
             " (Select Distinct '1' From 部门性质说明 Where 工作性质 Like '%药房' And 部门id = [2] And 服务对象 In (1, 3)) Or " & _
             " Decode(c.服务对象, 2, 1, 3, 1, 0) =" & _
             " (Select Distinct '1' From 部门性质说明 Where 工作性质 Like '%药房' And 部门id = [2] And 服务对象 In (2, 3)) Or Exists" & _
             " (Select 1 From 部门性质说明 Where 工作性质 Like '%药库' And 部门id = [2])) "
    End If
    
    strSQL = strSQL & _
        "     AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
        "     And (C.站点 = [1] or c.站点 is null) AND T.药品剂型=J.名称(+) And D.合同单位ID=Q.ID(+) " & _
        IIf(bln包含停用药品 = False, " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "(Select 收费细目id, 现价 售价 " & _
        " From 收费价目 Where (Sysdate Between 执行日期 And 终止日期 or Sysdate>=执行日期 And 终止日期 Is Null)) P," & vbNewLine
    If byt领用方式 = 1 Then
       '向留存领药
       strSQL = strSQL & _
           "(Select a.药品id,Max(上次产地) AS 产地,Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(nvl(实际数量,0)), 0, null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价," & _
           " To_Char(Sum(b.实际数量), '99999999999990.99') 留存数量 " & vbNewLine & _
           "From 药品库存 A, 药品留存 B " & vbNewLine & _
           "Where a.性质=1 and a.药品id=b.药品id And a.库房id=b.库房id and b.科室id=[3] and b.期间=to_date(sysdate,'yyyy') "
    Else
       '向药房领药
       strSQL = strSQL & _
           "(Select a.药品id, Max(a.上次产地) AS 产地, Sum(a.可用数量) 可用数量," & _
           " Sum(a.实际数量) 库存数量," & _
           " Sum(a.实际金额) 库存金额," & _
           " Sum(a.实际差价) 库存差价," & _
           " Decode(Sum(nvl(实际数量,0)), 0, null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价," & _
           " '' 留存数量 " & vbNewLine & _
           "From 药品库存 A " & vbNewLine & _
           "Where 性质=1 "
    End If
    If lng来源库房 <> 0 Or lng目标库房 <> 0 Then
       strSQL = strSQL & " And a.库房ID=" & IIf(lng来源库房 = 0, "[3]", "[2]")
    End If
    strSQL = strSQL & vbNewLine & _
       "Group By a.药品id) S," & vbNewLine & _
       "(Select 药品ID,库房ID,库房货位,领用标志 From 药品储备限额 Where 库房ID=[2]) E " & vbNewLine & _
       "Where D.药品ID=P.收费细目ID And D.药品ID=S.药品ID" & IIf(Not (IntStockCheck = 2 And byt编辑模式 = 2) Or byt盘点单据 = 1 Or Not bln检测库存, "(+)", "") & _
       "  And D.药品ID=E.药品ID(+) " & vbNewLine & _
       "Order By D.药名编码,D.药品编码 "
    Set grsMaster = zlDataBase.OpenSQLRecord(strSQL, "药品规格", gstrNodeNo, lng来源库房, lng目标库房)
    
    
    '*录入模式的数据集*'
    strSQL = _
        "Select " & _
        " d.剂型, d.药名编码, d.通用名称, d.药品来源 来源, d.基本药物, d.药典id, d.用途分类id, d.剂量单位, d.药品编码, d.药品名称, " & _
        " d.商品名, d.规格, d.产地 As 产地, d.药名id, d.药品id, " & _
        " trim(to_char(d.初始成本价 * " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) 上次采购价, " & _
        " trim(to_char(Decode(d.时价, '是', Decode(s.平均售价, Null, p.售价, s.平均售价), p.售价) * " & strConversionUnit & ", '99999999999990." & String(intPriceDigit, "0") & "')) 售价, " & _
        " d.售价单位, d.剂量系数 售价包装, " & _
        " d.门诊单位, d.门诊包装, d.住院单位, d.住院包装, d.药库单位, d.药库包装, " & _
        " trim(to_char(s.可用数量 / " & strConversionUnit & ", '99999999999990." & String(intNumberDigit, "0") & "')) 可用数量, " & _
        " s.库存数量,s.库存金额, s.库存差价, d.最大效期 有效期, d.药库分批, d.药房分批, d.时价, " & _
        " trim(to_char(d.指导批发价* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导批发价, " & _
        " trim(to_char(d.指导零售价* " & strConversionUnit & ",'99999999999990." & String(intCostDigit, "0") & "')) as 指导零售价, " & _
        " d.指导差价率, e.库房货位, d.批准文号, s.库存数量 实际数量," & _
        " s.留存数量, d.合同单位, d.药价级别,e.领用标志, d.简码, d.数字简码, d.五笔码,d.上次供应商 " & vbNewLine & _
        "From " & vbNewLine & _
        "  (SELECT DISTINCT J.名称 剂型,C.编码 药名编码,C.名称 AS 通用名称,0 AS 药典ID,M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位, " & _
        "   C.编码 AS 药品编码, c.输入名称 As 药品名称, a.名称 As 商品名, c.规格, c.产地, d.药品来源, d.基本药物, d.批准文号, d.药名id, " & _
        "   d.药品id, c.计算单位 As 售价单位, nvl(to_char(d.最大效期, '9999990'),0) 最大效期, " & _
        "   DECODE(D.药库分批,1,'是','否') 药库分批, DECODE(D.药房分批,1,'是','否') 药房分批, " & _
        "   to_char(D.剂量系数, " & CON_FMT & ") 剂量系数," & vbLf & _
        "   D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & vbNewLine & _
        "   D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & vbNewLine & _
        "   D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & vbNewLine & _
        "   D.指导批发价,d.指导零售价,nvl(D.成本价,0) 初始成本价, D.指导差价率, q.名称 合同单位, D.药价级别, c.简码, c.数字简码, c.五笔码, " & _
        "   DECODE(C.是否变价,1,'是','否') 时价,f.名称 上次供应商 " & vbNewLine
    strSQL = strSQL & _
        "   From (Select A.*, B.名称 As 输入名称, B.简码, b.数字简码, b.五笔码 From 收费项目目录 A," & vbNewLine & _
        "           (Select 收费细目id, max(Decode(码类, '3', 简码, Null)) 数字简码, max(Decode(码类, '1', 简码, Null)) 简码, max(Decode(码类, '2', 简码, Null)) 五笔码, 名称 " & vbNewLine & _
        "            From 收费项目别名 Where 码类 in (1,2,3) And " & IIf(gint输入药品显示 = 0, "性质 In (1, 3,9)", "性质=1") & vbNewLine & _
        "            Group By 收费细目id,名称 ) B " & vbNewLine & _
        "         Where A.ID=b.收费细目id And A.类别 IN ('5','6','7') And (A.站点 = [1] Or A.站点 is Null) "
    
    If bln忽略服务对象 = False Then
        strSQL = strSQL & " And" & _
             " (Decode(a.服务对象, 1, 1, 3, 1, 0) = " & _
             " (Select Distinct '1' From 部门性质说明 Where 工作性质 Like '%药房' And 部门id = [2] And 服务对象 In (1, 3)) Or " & _
             " Decode(a.服务对象, 2, 1, 3, 1, 0) =" & _
             " (Select Distinct '1' From 部门性质说明 Where 工作性质 Like '%药房' And 部门id = [2] And 服务对象 In (2, 3)) Or Exists" & _
             " (Select 1 From 部门性质说明 Where 工作性质 Like '%药库' And 部门id = [2])) "
    End If
    
    strSQL = strSQL & " ) C," & vbNewLine & _
        "   药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M,供应商 Q,供应商 F" & vbNewLine
        
    If bln调价 = False Then
        strSQL = strSQL & IIf(lng来源库房 <> 0, " ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[2] Group By 执行科室ID,收费细目ID) K", "") & vbNewLine & _
        IIf(lng目标库房 <> 0, "     ,(Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID=[3] Group By 执行科室ID,收费细目ID) I ", "") & vbNewLine
    End If
    
    strSQL = strSQL & _
        "   Where c.Id = d.药品id And d.药名id = t.药名id And t.药名id = m.Id And m.类别 In ('5', '6', '7') and t.临床自管药 is null And d.药品id = a.收费细目id(+) " & _
        "     And a.性质(+) = 3 And t.药品剂型 = j.名称(+) And d.合同单位id = q.Id(+) And D.上次供应商ID=f.ID(+) "
    If bln调价 = False Then
        strSQL = strSQL & IIf(lng来源库房 <> 0, "     And D.药品ID=K.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "") & _
        IIf(lng目标库房 <> 0, "     And D.药品ID=I.收费细目ID" & IIf(bln盘无存储库房药品, "(+)", ""), "")
    End If
    strSQL = strSQL & _
        IIf(bln包含停用药品 = False, " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01') ) D,", ") D,") & vbNewLine & _
        "  (Select 收费细目id, Trim(To_Char(现价, '999999999990." & String(7, "0") & "')) 售价 " & _
        "   From 收费价目 Where (Sysdate Between 执行日期 And 终止日期 or Sysdate>=执行日期 And 终止日期 Is Null)) P," & vbNewLine

    If byt领用方式 = 1 Then
       '向留存领药
       strSQL = strSQL & _
           "(Select a.药品id,Max(上次产地) AS 产地, Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价, " & _
           " To_Char(Sum(b.实际数量), '99999999999990.99') 留存数量 " & vbNewLine & _
           "From 药品库存 A, 药品留存 B " & vbNewLine & _
           "Where a.性质=1 and a.药品id=b.药品id And a.库房id=b.库房id and b.科室id=[3] and b.期间=to_date(sysdate,'yyyy') "
    Else
       '向药房领药
       strSQL = strSQL & _
           "(Select a.药品id, Max(a.上次产地) AS 产地, Sum(a.可用数量) 可用数量," & _
           " To_Char(Sum(a.实际数量), " & CON_FMT & ") 库存数量," & _
           " To_Char(Sum(a.实际金额), " & CON_FMT & ") 库存金额," & _
           " To_Char(Sum(a.实际差价), " & CON_FMT & ") 库存差价," & _
           " Decode(Sum(Nvl(实际数量, 0)), 0, Null, Sum(a.实际金额) / Sum(a.实际数量)) As 平均售价, " & _
           " '' 留存数量 " & vbNewLine & _
           "From 药品库存 A " & vbNewLine & _
           "Where 性质=1 "
    End If
    If lng来源库房 <> 0 Or lng目标库房 <> 0 Then
       strSQL = strSQL & " And a.库房ID=" & IIf(lng来源库房 = 0, "[3]", "[2]")
    End If
    strSQL = strSQL & vbNewLine & _
       "Group By a.药品id) S," & vbNewLine & _
       "(Select 药品ID,库房ID,库房货位,领用标志 From 药品储备限额 Where 库房ID=" & IIf(byt编辑模式 = 2, "[2]", "[3]") & ") E " & vbNewLine & _
       "Where D.药品ID=P.收费细目ID And D.药品ID=S.药品ID" & IIf(Not (IntStockCheck = 2 And byt编辑模式 = 2) Or byt盘点单据 = 1 Or Not bln检测库存, "(+)", "") & _
       "  And D.药品ID=E.药品ID(+) " & vbNewLine & _
       "Order By D.药名编码,D.药品编码 "
    Set grsMasterInput = zlDataBase.OpenSQLRecord(strSQL, "药品规格", gstrNodeNo, lng来源库房, lng目标库房, IIf(gint简码方式 = 0, 1, 2))
    
    '*药品分批*'
    If byt编辑模式 = 2 Then
        strSQL = _
            "Select Rid,库房,药品ID,批次,入库日期,批号,生产日期,产地,成本价,售价,时价,门诊单位,门诊包装,住院单位,住院包装,药库单位,药库包装," & _
            "  有效期,实际数量,可用数量,库存数量,库存金额,库存差价,上次供应商ID,批准文号,供应商 " & vbLf & _
            "From (Select Distinct 2 Rid, p.名称 库房, k.药品id, k.批次, To_Char(b.入库日期, 'yyyy-mm-dd') As 入库日期, k.上次批号 批号," & _
            "  To_Char(k.上次生产日期, 'YYYY-MM-DD') 生产日期, k.上次产地 产地," & _
            " trim(to_char(k.平均成本价 " & ",'99999999999990." & String(intCostDigit, "0") & "')) 成本价, " & _
            "  Decode(Nvl(k.批次, 0), 0, Decode(Sign(k.实际数量), 1, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量), A.现价) " & _
            "        ,Nvl(k.零售价, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量) ) ) 售价," & _
            "  Nvl(k.零售价, k.实际金额 / decode(nvl(k.实际数量,0), 0, 1, k.实际数量) ) 时价," & _
            "  D.门诊单位, to_char(D.门诊包装, " & CON_FMT & ") 门诊包装," & _
            "  D.住院单位, to_char(D.住院包装, " & CON_FMT & ") 住院包装," & _
            "  D.药库单位, to_char(D.药库包装, " & CON_FMT & ") 药库包装," & _
            "  k.效期" & IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "-1", "") & " 有效期," & _
            "  k.实际数量, k.可用数量, k.实际数量 库存数量, k.实际金额 库存金额, k.实际差价 库存差价, k.上次供应商id, k.批准文号,f.名称 供应商 " & vbNewLine & _
            "From 部门表 P, 药品规格 D, 药品库存 K, 收费价目 A, 药品入库信息 B, 供应商 F " & vbNewLine & _
            "Where k.库房id = p.Id And d.药品id = k.药品id And d.药品id=a.收费细目id " & _
            "  And k.性质 = 1 And k.药品id = b.药品id(+) And k.库房id = b.库房id(+) And k.批次 = b.批次(+) And k.库房id = [1] And K.上次供应商ID=f.ID(+) "
        If byt盘点单据 = 1 Then
            strSQL = strSQL & " And (K.实际数量<>0 Or K.实际金额<>0 Or K.实际差价<>0) ) " & vbNewLine
'        ElseIf byt盘点单据 = 2 Then
'            '1303 如果是库存差价调整模块，则允许过滤库存数量为0的药品记录
'            gstrSQL = strSQL & " ) " & vbNewLine
        Else
            strSQL = strSQL & " And K.实际数量<>0 ) " & vbNewLine
        End If
        If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
            strSQL = strSQL & "Order By 药品id, 批次 "
        ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 1 Then
            strSQL = strSQL & "Order By 药品id, 有效期, 批次 "
        ElseIf gtype_UserSysParms.P150_药品出库优先算法 = 2 Then
            strSQL = strSQL & "Order By 药品id, 批号, 批次 "
        Else
            strSQL = strSQL & "Order By 药品id, 批次 "
        End If

        Set grsSlave = zlDataBase.OpenSQLRecord(strSQL, "药品分批", IIf(lng来源库房 = 0, lng目标库房, lng来源库房))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReleaseSelectorRS()
    If Not grsMaster Is Nothing Then
        If grsMaster.State = adStateOpen Then grsMaster.Close
        Set grsMaster = Nothing
    End If
    
    If Not grsMasterInput Is Nothing Then
        If grsMasterInput.State = adStateOpen Then grsMasterInput.Close
        Set grsMasterInput = Nothing
    End If
    
    If Not grsSlave Is Nothing Then
        If grsSlave.State = adStateOpen Then grsSlave.Close
        Set grsSlave = Nothing
    End If
End Sub

Public Sub SetUsingDrugsRS()
    '----------------------------------------------------------------------------------------
    '功能：初始化grsUsingDrugs对象，
    '      加载在用药品列表
    '----------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    With grsUsingDrugs
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockReadOnly
    End With
    
    strSQL = "Select b.药品id" & vbNewLine & _
            "From 收费项目目录 A, 药品规格 B" & vbNewLine & _
            "Where a.Id = b.药品id And Nvl(a.撤档时间, To_Date('3000-01-01', 'yyyy-MM-dd')) = To_Date('3000-01-01', 'yyyy-MM-dd')"
            
    Set grsUsingDrugs = zlDataBase.OpenSQLRecord(strSQL, "在用药品")
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ReleaseUsingDrugsRS()
    If Not grsUsingDrugs Is Nothing Then
        If grsUsingDrugs.State = adStateOpen Then grsUsingDrugs.Close
        Set grsUsingDrugs = Nothing
    End If
End Sub

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '外挂扩展接口初始化
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
            End If
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

Public Sub zlPlugIn_Unload(objPlugIn As Object)
    '卸载外挂接口
    Set objPlugIn = Nothing
End Sub

Public Function CallPlugInDrugStuffWork(objPlugIn As Object, ByVal intType As Integer, ByVal lng库房ID As Long, ByVal strNo As String, _
    ByVal int单据 As Integer, Optional ByVal strDate As String, _
    Optional ByVal strReserve As String) As Boolean
'    功能：调用外挂部件的药品卫材流通通用业务接口
'    参数：intType：业务类型，0-新增，1-修改，2-删除，3-审核，4-冲销
'      lng库房ID ：库房ID
'      strNO：单据号
'      int单据：单据类型
'      药品：0-计划；1-外购入库；2-自制入库；3-协药入库；4-其他入库；6-库房移出；7-部门领用；11-其他出库；12-盘点；
'      卫材：0-计划；15-材料外购入库,16-材料自制入库,17-材料其他入库,19-材料移库,20-部门材料领用,21-材料其他出库,22-材料盘点
'      strDate：单据日期，只有冲销时传入（因为允许部分冲销，用时间来区分本次冲销的是哪些数据，如果不传就是不允许部分冲销）
'      strReserve=扩展参数
    
    If Not objPlugIn Is Nothing Then '确保 objPlugIn 对象不为 Nothing
        On Error Resume Next
        objPlugIn.DrugStuffWork intType, lng库房ID, strNo, int单据, strDate, strReserve
        If InStr(",438,0,", "," & Err.Number & ",") = 0 Then
            MsgBox "zlPlugIn 外挂部件执行 DrugStuffWork 时出错：" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbInformation, gstrSysName
        End If
        Err.Clear: On Error GoTo 0
    End If
    
End Function
