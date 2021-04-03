Attribute VB_Name = "mdl兴安"
Option Explicit
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
End Type
Public InitInfor_兴安 As InitbaseInfor
Public g病人身份_兴安 As 病人身份
'显示当前运行的窗体的API声明
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Const WS_BORDER = &H800000
Const OFS_MAXPATHNAME = 128
Const OF_EXIST = &H4000

 
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
'关闭当前运行的窗体的API声明
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
Public Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long


Private Type 病人身份
    个人编号            As String
    卡号                As String
    姓名                As String
    性别                As String
    年龄                As Integer
    单位名称            As String
    人员类别            As String
    交易类型            As String
    帐户状态            As String
    帐户余额            As Double
    当日已用药品金额    As Double
    病种代码            As String
    病种名称            As String
    当年已用自负段      As Double
    当年已用慢病统筹    As Double
    费用总额            As Double       '表示当前费用总额
    虚拟结算            As Variant      '门诊用.
    byt类型             As Byte         ''0-门诊收费，1-住院
    住院登记号          As String       '住院登记号
    本年已用基本统筹    As Double       '住院用
    本年已用大病统筹    As Double       '住院用
    当年第几次住院      As Integer      '住院用
    本次住院起付标准    As Double       '住院用
    医院类型            As String       '住院用
    医院名称            As String       '住院用
    门诊流水号          As String
End Type
Public Enum 业务类型_兴安
        建立数据库连接 = 0
        关闭数据库连接
        操作员注册
        读取个人信息
        读取医保项目信息
        门诊预处理
        门诊明细写入
        门诊结算提交
        门诊结算冲销
        住院登记
        取消住院登记
        住院明细写入
        住院明细取消
        住院结算
        住院结算取消
        住院事务开始
        住院事务提交
        住院事务回滚
End Enum
Private gobj兴安 As Object             '引用兴安对象Dll
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'常用函数过程声明
Public Function 医保初始化_兴安() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始医保的相关变量
    '--入参数:
    '--出参数:
    '--返  回:初始化成功，返回true；否则，返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim strUser As String
    Dim strServer As String
    Dim strPass As String
    Dim rsTemp As New ADODB.Recordset
    
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_兴安.模拟数据 = True
    Else
        InitInfor_兴安.模拟数据 = False
    End If
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_兴安
    
    Call OpenRecordset(rsTemp, "读取医院编码")
    InitInfor_兴安.医院编码 = Nvl(rsTemp!医院编码)
        
        
    '调用中转站程序
    If ExcuteExeFile() = False Then Exit Function
    
    Err = 0
    On Error GoTo ErrHand:
    '打开数据库联接.
    Dim intReturn As Integer
    
    '打开医保数据库
    If 业务请求_兴安(建立数据库连接, "", "") = False Then
        Exit Function
    End If
    医保初始化_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ExcuteExeFile() As Boolean
    '执行监听器
    Dim mError As String
    Dim strFile As String
    
    ExcuteExeFile = False
    
    '先关掉监听
    Err = 0
    On Error Resume Next
    Call 关闭监听
    
    strFile = App.Path & "\中联监听器.exe"
    If FindFile(strFile) = False Then
        ShowMsgbox "文件(" & App.Path & "\中联监听器.exe)不存在!请与中联公司联系"
        Exit Function
    End If
    Err = 0
    On Error Resume Next
    mError = Shell(strFile, vbNormalFocus)
    ExcuteExeFile = True
End Function

Public Function FindFile(ByVal strFileName As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '--功能:查找指定的文件是否存在
    '--返回: 如果存在此文件为True,否则为Flase
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim typOfStruct As OFSTRUCT
    
    On Error Resume Next
    FindFile = False
    If Len(strFileName) > 0 Then
        apiOpenFile strFileName, typOfStruct, OF_EXIST
        FindFile = typOfStruct.nErrCode <> 2
    End If
End Function

Public Function 医保终止_兴安() As Boolean
    
    Err = 0
    On Error Resume Next
    Set gobj兴安 = Nothing
    '打开医保数据库
    If 业务请求_兴安(关闭数据库连接, "", "") = False Then
        Exit Function
    End If
    
    Call 关闭监听
    医保终止_兴安 = True
End Function


Public Sub 关闭监听()
    Dim app_hwnd As Long
    app_hwnd = FindWindow(vbNullString, "中联监听器")
    SendMessage app_hwnd, WM_CLOSE, 0, 0
End Sub

Public Function 身份标识_兴安(Optional bytType As Byte, Optional lng病人ID As Long) As String
    Dim str备注 As String, rsPatient As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    
    身份标识_兴安 = frmIdentify兴安.GetPatient(bytType, lng病人ID)
    
End Function
Public Function 身份标识_兴安2(ByVal strCard As String, ByVal strPass As String, Optional lng病人ID As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    身份标识_兴安2 = frmIdentify兴安.GetPatient(3, lng病人ID)
End Function

Private Function Get病人信息(ByVal lng病人ID As Long)
    Dim rsTemp As New ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '--保险帐户保存对比说明:
    '--病人id, 险类, 中心, 卡号（医保卡号), 医保号(个人编号), 密码, 人员身份(人员类别), 单位编码(单位名称), 顺序号(无),
    '--退休证号(病种编码-病种名称), 帐户余额(帐户余额), 当前状态, 病种id（无), 在职(1), 年龄段(年龄), 灰度级(无),
    '--就诊时间(无)
    
    Dim strTemp As String
    Dim strArr
    
    Err = 0
    On Error GoTo ErrHand:
    
    DebugTool "进入Get病人信息函数"
    
   '保险帐户:增加字段:当日已用药品,当年已用自负段,当年已用慢病,本年已用基本统筹,本年已用大病统筹,当年住院次数,本次住院起付标准
    
    gstrSQL = "select a.卡号,a.医保号,a.密码,a.人员身份,a.单位编码,b.工作单位,a.顺序号,a.退休证号,a.帐户余额,a.当前状态,a.病种id,a.在职,a.年龄段,a.灰度级,a.就诊时间," & _
             "        b.姓名,b.性别, b.年龄, b.出生日期, b.身份证号,A.当日已用药品,A.当年已用自负段,A.当年已用慢病,A.本年已用基本统筹,A.本年已用大病统筹,A.当年住院次数,A.本次住院起付标准" & _
             " from 保险帐户 a,病人信息 b " & _
             " WHERE a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_兴安
 
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人信息"
    
    With g病人身份_兴安
        .卡号 = Nvl(rsTemp!卡号)
        .个人编号 = Nvl(rsTemp!医保号)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .单位名称 = Nvl(rsTemp!单位编码)
        .年龄 = Nvl(rsTemp!年龄段, 0)
        .人员类别 = Nvl(rsTemp!人员身份)
        .住院登记号 = Nvl(rsTemp!顺序号)
        strTemp = Nvl(rsTemp!退休证号, "")
        If strTemp <> "" And InStr(1, strTemp, "-") <> 0 Then
            .病种代码 = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
            .病种名称 = Mid(strTemp, InStr(1, strTemp, "-") + 1)
        Else
            .病种代码 = ""
            .病种名称 = ""
        End If
        .帐户余额 = Nvl(rsTemp!帐户余额, 0)
        
        .当年第几次住院 = Nvl(rsTemp!当年住院次数, 1)
        .本年已用大病统筹 = Nvl(rsTemp!本年已用大病统筹, 0)
        .本年已用基本统筹 = Nvl(rsTemp!本年已用基本统筹, 0)
        .本次住院起付标准 = Nvl(rsTemp!本次住院起付标准, 0)
        .当年已用慢病统筹 = Nvl(rsTemp!当年已用慢病, 0)
        .当年已用自负段 = Nvl(rsTemp!当年已用自负段, 0)
        .当日已用药品金额 = Nvl(rsTemp!当日已用药品, 0)
    End With
    DebugTool "退出Get病人信息函数"
Exit Function
ErrHand:
    DebugTool "获取病人信息失败" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误信息:" & Err.Description
End Function

Public Function 身份鉴别_兴安() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:远程身份鉴别
    '--入参数:
    '--出参数:
    '--返  回:成功true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim strOutput As String
    Dim strInput As String
    Dim blnReturn As Boolean
    Dim strArr
    
        
    Err = 0
    On Error GoTo ErrHand:
    身份鉴别_兴安 = False
    
    If g病人身份_兴安.byt类型 = 0 Then
        strInput = g病人身份_兴安.卡号
    Else
        strInput = g病人身份_兴安.个人编号 & vbTab
        strInput = strInput & InitInfor_兴安.医院编码
    End If
    
    DebugTool "进入身份鉴别函数"
    
    '业务请求
    blnReturn = 业务请求_兴安(读取个人信息, strInput, strOutput)
    
    If blnReturn = False Then
        Exit Function
    End If
    If strOutput = "" Then
        '刘兴宏 /*200408*/
        DebugTool "读取个人信息时出现了传出串为空了!"
        Exit Function
    End If
    strArr = Split(strOutput, vbTab)
    
    '给公用变量赋值
    With g病人身份_兴安
        'byt类型 0-门诊,1-住院
        If .byt类型 = 0 Then
            .个人编号 = strArr(0)
            .姓名 = strArr(1)
            .性别 = strArr(2)
            .年龄 = Val(strArr(3))
            .单位名称 = strArr(4)
            .人员类别 = strArr(5)
            .帐户状态 = strArr(6)
            .帐户余额 = Val(strArr(7))
            .当日已用药品金额 = Val(strArr(8))
            .病种代码 = strArr(9)
            .病种名称 = strArr(10)
            .当年已用自负段 = Val(strArr(11))
            .当年已用慢病统筹 = Val(strArr(12))
            身份鉴别_兴安 = True
            DebugTool "身份鉴别成功"
            Exit Function
        End If
        .住院登记号 = strArr(0)
        .姓名 = strArr(1)
        .性别 = strArr(2)
        .年龄 = Val(strArr(3))
        .单位名称 = strArr(4)
        .人员类别 = strArr(5)
        .本年已用基本统筹 = Val(strArr(6))
        .本年已用大病统筹 = Val(strArr(7))
        .当年第几次住院 = Val(strArr(8))
        .本次住院起付标准 = Val(strArr(9))
        .医院类型 = strArr(10)
        .医院名称 = strArr(11)
        .病种代码 = ""
        .病种名称 = ""
    End With
    身份鉴别_兴安 = True
    DebugTool "身份鉴别成功"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    身份鉴别_兴安 = False
End Function

Public Function 门诊虚拟结算_兴安(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim strArr
    Dim strInput  As String
    Dim strOutput  As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    DebugTool "进入门诊虚拟结算接口"
    
    With rs明细
        Do While Not .EOF
            gstrSQL = "Select 编码,名称 From 收费细目 where id=" & Nvl(!收费细目ID, 0)
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目编码"
            If 业务请求_兴安(读取医保项目信息, Nvl(rsTemp!编码), strOutput) = False Then
                Exit Function
            End If
            If strOutput = "" Then
                DebugTool "获取医保项目信息时,输出串为空了"
                Exit Function
            End If
            strArr = Split(strOutput, vbTab)
            
            '入参:序号,金额,费用类别,医保标志,医保代码,处方天数
            strInput = .AbsolutePosition & vbTab
            strInput = strInput & Nvl(!实收金额, 0) & vbTab
            strInput = strInput & strArr(0) & vbTab
            strInput = strInput & strArr(1) & vbTab
            strInput = strInput & strArr(2) & vbTab
            If g病人身份_兴安.交易类型 = "普通医保门诊" Then
                strInput = strInput & "1" & vbTab
            Else
                '摘要内容:日次数;次用量;次单位;日用量;处方天数;处方总量
                strTemp = Nvl(!摘要)
                strTemp = strTemp & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":" & vbTab & ":"
                strTemp = Split(strTemp, vbTab)(4)
                strTemp = Split(strTemp, ":")(1)
                strInput = strInput & IIf(Val(strTemp) = 0, 1, Val(strTemp)) & vbTab
            End If
            
            If 业务请求_兴安(门诊预处理, strInput, strOutput) = False Then
                Exit Function
            End If
            If strOutput = "" Then
                DebugTool "门诊预处理时,输出串为空了"
                Exit Function
            End If
            '出参:本次个人帐户金额,本次个人帐户余额,本次自负段金额,本次统筹帐金额,本次自负金额,返因值
            strArr = Split(strOutput, vbTab)
            .MoveNext
        Loop
    End With
    
    g病人身份_兴安.虚拟结算 = strArr
    
    str结算方式 = "个人帐户;" & Format(Val(strArr(0)), "###0.00;-###0.00;0;0") & ";0" '本次基本个人帐户支付,不充许修改
    str结算方式 = str结算方式 & "|" & "医保基金;" & Format(Val(strArr(3)), "###0.00;-###0.00;0;0") & ";0"
    
    DebugTool "门诊虚拟结算成功"
    门诊虚拟结算_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get明细记录(ByRef lng结帐ID As Long, Optional strNO As String, Optional lng记录性质 As Long, Optional lng记录状态 As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取本次结帐的明细记录
    '--入参数:lng结帐ID-本次结帐的ID记录
    '         strno-本次处方的单据号,lng记录性质=记录性质,lng记录状态
    '--出参数:
    '--返  回:SQL语句
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim strFields As String
    
    If lng结帐ID = 0 Then
            strSql = " " & _
                "  Select Rownum 标识号,A.ID,A.病人ID,a.主页id,A.收费细目id,收入项目id,A.NO,A.序号 ,A.记录性质,A.记录状态,A.发生时间 as 经办时间 ,c.名称 as 开单部门,a.开单人 as 开单医生,nvl(a.是否上传,0) 是否上传, " & _
                "      A.数次*A.付数 as 数量,A.计算单位,A.实收金额 as 实际金额,Round(A.实收金额/(A.数次*A.付数),4) as 实际价格,A.实收金额 as 实收金额, " & _
                "      A.收费类别,A.摘要,A.操作员姓名 as 经办人," & _
                "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.顺序号,L.病种ID,L.就诊时间 ,J.编码,J.名称 as 商品名,J.规格" & _
                "  From (Select * From 病人费用记录 Where 记录状态<>0 and NO='" & strNO & "' and 记录性质=" & lng记录性质 & " and 记录状态=" & lng记录状态 & " and  Nvl(附加标志,0)<>9 ) A,部门表 C," & _
                "       保险帐户 L,收费细目 J " & _
                "  Where A.开单部门id=C.id(+)  and  A.病人id=L.病人id  and a.收费细目id=J.id and L.险类=" & gintInsure & "  " & _
                "  Order by a.NO,A.记录性质,A.记录状态,a.序号"
                
    Else
        strSql = " " & _
            "  Select Rownum 标识号,A.ID,A.病人ID,a.主页id,A.收费细目id,收入项目id,A.NO,A.序号 ,A.记录性质,A.记录状态,A.发生时间 as 经办时间 ,c.名称 as 开单部门,a.开单人 as 开单医生,nvl(a.是否上传,0) 是否上传, " & _
            "      A.数次*A.付数 as 数量,A.计算单位,A.实收金额 as 实际金额,Round(A.结帐金额/(A.数次*A.付数),4) as 实际价格,A.结帐金额 as 实收金额, " & _
            "      A.收费类别,A.摘要,A.操作员姓名 as 经办人," & _
            "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.顺序号,L.病种ID,L.就诊时间,J.编码 ,J.名称 as 商品名,J.规格" & _
            "  From (Select * From 病人费用记录 Where 记录状态<>0 and 结帐ID=" & lng结帐ID & " and  Nvl(附加标志,0)<>9 ) A,部门表 C," & _
            "       保险帐户 L,收费细目 J " & _
            "  Where A.开单部门id=C.id(+) and  A.病人id=L.病人id and a.收费细目id=J.id and L.险类=" & gintInsure & _
            "   Order by a.NO,A.记录性质,A.记录状态,a.序号"
    End If
    Get明细记录 = strSql
End Function
Public Function 门诊结算_兴安(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    
    Dim lng病人ID As Long, strOutput As String, strInput As String
    Dim strArr
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim str医保代码 As String
    门诊结算_兴安 = False
    
    DebugTool "进入门诊结算"
        
    Err = 0
    On Error GoTo ErrHand:
    
    '读取明细记录
    gstrSQL = Get明细记录(lng结帐ID)
    zlDatabase.OpenRecordset rs明细, gstrSQL, "获取明细记录"
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "没有一条明细记录,不能进行结算!"
        Exit Function
    End If
    DebugTool "开始传明细"
    g病人身份_兴安.费用总额 = 0
    With rs明细
            lng病人ID = Nvl(!病人ID, 0)
            Do While Not .EOF
                '获取医保的相关信息
                If 业务请求_兴安(读取医保项目信息, Nvl(!编码), strOutput) = False Then Exit Function
                If strOutput = "" Then
                    DebugTool "在读取区保项目信息时,没有传出串!"
                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                
                '普通门诊
                '入参:kbname(医生名称),ysname(科室名称),xh(序号),fycode(费用代码),fyname(费用名称),gg(规格),dw(单位),dj(单价),sl(数量),je(金额),fylb(费用类别),ypbz(医保类别),ybdm(医保代码)
                '特殊门诊:
                '入参:kbname(医生名称),ysname(科室名称),xh(序号),fycode(费用代码),fyname(费用名称),gg(规格),dw(单位),dj(单价),sl(数量),je(金额),fylb(费用类别),ypbz(医保类别),ybdm(医保代码),yf(日次数),yl(次用量),yfdw(次单位),mryl(日用量),cfts(处方天数),cfzl(处方总量)
                
                '写明细记录
                strInput = Nvl(!开单部门) & vbTab
                strInput = strInput & Nvl(!开单医生) & vbTab
                strInput = strInput & Nvl(!序号, 0) & vbTab
                strInput = strInput & Nvl(!编码) & vbTab
                strInput = strInput & Nvl(!商品名) & vbTab
                strInput = strInput & Nvl(!规格) & vbTab
                strInput = strInput & Nvl(!计算单位) & vbTab
                strInput = strInput & Nvl(!实际价格, 0) & vbTab
                strInput = strInput & Nvl(!数量, 0) & vbTab
                strInput = strInput & Nvl(!实收金额, 0) & vbTab
                strInput = strInput & strArr(0) & vbTab
                strInput = strInput & strArr(1) & vbTab
                str医保代码 = strArr(2)
                
                If g病人身份_兴安.交易类型 = "普通医保门诊" Then
                    strInput = strInput & str医保代码
                Else
                    '摘要内容:日次数;次用量;次单位;日用量;处方天数;处方总量
                    '日次数:2    次用量:2    次单位:片   日用量:4    处方天数:5  处方总量:20 帐户余额:0
                    strInput = strInput & str医保代码 & vbTab
                    strTemp = Nvl(!摘要, ":0" & vbTab & ":0" & vbTab & ":" & vbTab & ":0" & vbTab & ":1" & vbTab & ":0")
                    strTemp = strTemp & ":0" & vbTab & ":0" & vbTab & ":" & vbTab & ":0" & vbTab & ":1" & vbTab & ":0"
                    
                    strArr = Split(strTemp, vbTab)
                    
                    strInput = strInput & Val(Split(strArr(0), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(1), ":")(1)) & vbTab
                    strInput = strInput & Split(strArr(2), ":")(1) & vbTab
                    strInput = strInput & Val(Split(strArr(3), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(4), ":")(1)) & vbTab
                    strInput = strInput & Val(Split(strArr(5), ":")(1)) & vbTab
                End If
                strOutput = ""
                If 业务请求_兴安(门诊明细写入, strInput, strOutput) = False Then Exit Function
                If strOutput = "" Then
                    DebugTool "在门诊明细写入时,没有传出串!"
                    Exit Function
                End If
                strArr = Split(strOutput, vbTab)
                '为病人费用记录打上标记，以便随时上传
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                '摘要值:日次数;次用量;次单位;日用量;处方天数;处方总量;当帐余额
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(!摘要) & vbTab & "帐户余额:" & Val(strArr(0)) & "')"
                zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                g病人身份_兴安.费用总额 = g病人身份_兴安.费用总额 + Nvl(!实收金额, 0)
                .MoveNext
            Loop
    End With
    
    DebugTool "明细上传成功，并开始结算交易提交"

    '正试结算
    strInput = ""
    If 业务请求_兴安(门诊结算提交, strInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN(本次个人帐户余额),累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN(费用总额),全自付金额_IN(本次自负段金额),首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN(本次统筹帐金额),大病自付金额_IN,超限自付金额_IN(本次自负金额),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(门诊流水号),主页ID_IN,中途结帐_IN,备注_IN
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    With g病人身份_兴安
        gstrSQL = "zl_保险结算记录_insert( 1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          "NUll," & Val(.虚拟结算(4)) & ",Null,NULL,NULL,null,Null,NULL," & _
         .费用总额 & "," & Val(.虚拟结算(0)) & ",Null," & _
         "Null," & Val(.虚拟结算(1)) & ",Null," & Val(.虚拟结算(2)) & "," & Val(.虚拟结算(3)) & ",'" & _
         strArr(0) & "',Null,Null,NULl)"
    End With
    Call ExecuteProcedure("保存保险结算记录")
    DebugTool "门诊结算成功"
    门诊结算_兴安 = True
    Exit Function
ErrHand:
    DebugTool "门诊结算(门诊结算_兴安)" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & "错误信息:" & Err.Description
End Function

Private Function Get冲销ID(ByVal lng结帐ID As Long) As Long
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取当前冲销记录的ID值
    '--入参数:
    '--出参数:
    '--返  回:冲销ID
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '取冲销记录的结帐ID
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "读新产生的结帐ID")
    If rsTemp.EOF Then
        Get冲销ID = 0
    Else
        Get冲销ID = Nvl(rsTemp!结帐ID, 0)
    End If
End Function

Public Function 门诊结算冲销_兴安(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    

    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    Dim str门诊流水号 As String
    Dim lng冲销ID As Long
    Dim strOutput As String
    Dim strArr
    
    门诊结算冲销_兴安 = False
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "进入门诊结算冲销"
    
    gstrSQL = "Select * From 保险结算记录  where 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取门诊流水号"
    
    lng冲销ID = Get冲销ID(lng结帐ID)
    str门诊流水号 = Nvl(rsTemp!支付顺序号)
    
    '请求取消门诊结算
    If 业务请求_兴安(门诊结算冲销, str门诊流水号, strOutput) = False Then Exit Function
    If strOutput = "" Then
        strOutput = "0"
    End If
    strArr = Split(strOutput, vbTab)
    
    DebugTool "进入保存保险结算记录"
   
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN(本次个人帐户余额),累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN(费用总额),全自付金额_IN(本次自负段金额),首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN(本次统筹帐金额),大病自付金额_IN,超限自付金额_IN(本次自负金额),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(门诊流水号),主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert( 1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & -1 * Nvl(rsTemp!帐户累计支出, 0) & ",Null,NULL,NULL,null,Null,NULL," & _
     -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",Null," & _
     "Null," & -1 * Nvl(rsTemp!统筹报销金额, 0) & ",Null," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & _
     strArr(0) & "',Null,Null,NULl)"
    Call ExecuteProcedure("保存保险结算记录")
    DebugTool "门诊结算冲销成功"
    门诊结算冲销_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保设置_兴安() As Boolean
    医保设置_兴安 = True
End Function

Public Function 入院登记_兴安(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    DebugTool "进入入院登记接口"
    
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用,请先进行结帐!"
        Exit Function
    End If
    
   ' Call Get病人信息(病人ID)
    
    If 住院信息提交(lng病人ID, lng主页ID) = False Then Exit Function


    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_兴安 & ")"
    Call ExecuteProcedure("办理入院登记")
    
    DebugTool "办理入院成功"
    入院登记_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_兴安(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    
    Err = 0
    On Error GoTo ErrHand
    
    DebugTool "进入扩院登撤消接口"
    
    入院登记撤销_兴安 = False
    If 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "存在未结费用，不能撤消入院登记"
        Exit Function
    End If
    
    Get病人信息 lng病人ID
    '补充住院信息
    strInput = g病人身份_兴安.住院登记号 & vbTab
    strInput = strInput & InitInfor_兴安.医院编码
    
    
    If 业务请求_兴安(取消住院登记, strInput, strOutput) = False Then Exit Function
        
    DebugTool "调用医保的取消业务成功,并开始更新保险帐户的相关状态！"
    
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_兴安 & ")"
    Call ExecuteProcedure("办理撤销入院登记")
    
    DebugTool "取消成功"
    入院登记撤销_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 出院登记_兴安(lng病人ID As Long, lng主页ID As Long) As Boolean
    
    Err = 0
    On Error GoTo ErrHand:
    DebugTool "进入出院登记"
    
    出院登记_兴安 = False
    Get病人信息 lng病人ID
    
    If 在院病人信息_兴安(lng病人ID, lng主页ID) = False Then Exit Function

    
    
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & TYPE_兴安 & ")"
    Call ExecuteProcedure("出院登记")
    
    DebugTool "出院登记成功"
    出院登记_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 出院登记撤销_兴安(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "进入出院登记撤销!"
    出院登记撤销_兴安 = False
    
    If Not 存在未结费用(lng病人ID, lng主页ID) Then
        ShowMsgbox "该病人已经结帐,不能再进行出院撤消."
        Exit Function
    End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_兴安 & ")"
    Call ExecuteProcedure("办理撤销出院登记")
    
    DebugTool "出院登记撤销成功!"
    出院登记撤销_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_兴安(ByVal lng病人ID As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    
    '读卡失败则退出
    Err = 0: On Error GoTo ErrHand:
    DebugTool "进入获取个人帐户余额(个人余额_兴安)"
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额 From 保险帐户 Where 险类=" & gintInsure
    gstrSQL = gstrSQL & " And 病人id=" & lng病人ID
    
    Call OpenRecordset(rsTemp, "读取帐户余额")
    个人余额_兴安 = Nvl(rsTemp!帐户余额, 0)
    
    DebugTool "获取成功,余额为:" & Nvl(rsTemp!帐户余额, 0)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_兴安(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand:
    住院虚拟结算_兴安 = ""
    If bln结帐处 = False Then Exit Function
    
    gstrSQL = "Select 当前状态 From 保险帐户 where 病人id=" & lng病人ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否出院没有"
    If Nvl(rsTemp!当前状态, 0) = 1 Then
        ShowMsgbox "该病人还办理出院,所以不能结算!"
        Exit Function
    End If
    With rsExse
        g病人身份_兴安.费用总额 = 0
        Do While Not .EOF
            g病人身份_兴安.费用总额 = g病人身份_兴安.费用总额 + Nvl(!金额, 0)
            .MoveNext
        Loop
    End With
'    If bln结帐处 Then
'        '结帐处需验证身份
'         If frmIdentify兴安.GetPatient(4, lng病人id) = "" Then Exit Function
'    End If
    Call Get病人信息(lng病人ID)
    住院虚拟结算_兴安 = "医保基金;" & g病人身份_兴安.费用总额 & ";0"
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 上传处方明细(ByVal rs明细 As ADODB.Recordset) As Boolean

    '上传处方明细
    Dim strInput As String, strOutput As String, strTemp As String
    Dim strArr
    Dim rsTemp As New ADODB.Recordset
     
    上传处方明细 = False
    DebugTool "进入上传处方明细函数    "
    Err = 0
    On Error GoTo ErrHand:
    g病人身份_兴安.费用总额 = 0
    With rs明细
        If .RecordCount = 0 Then
            ShowMsgbox "本次结算没有相关的明细记录,不能进行月结"
            Exit Function
        End If
        
        '写未上传的明细记录
        Do While Not .EOF
            
            If Nvl(!是否上传, 0) <> 1 And Nvl(!实际金额, 0) <> 0 Then
                    
                    If !记录状态 <> 1 Then
                        '表示被冲销的记录
                        '需确定原始记录中的项目序号
                        gstrSQL = "Select 摘要 From 病人费用记录 where mod(记录状态,3)=0 and NO='" & Nvl(!NO) & "' and 记录性质=" & Nvl(!记录性质, 0) & " and 序号=" & Nvl(!序号)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取被冲销的项目序号"
                        If rsTemp.RecordCount = 0 Then
                            ShowMsgbox "冲销的原始单据未找到!" & Nvl(!NO)
                            Exit Function
                        End If
                        strTemp = Nvl(rsTemp!摘要) & vbTab & vbTab
                        strArr = Split(strTemp, vbTab)
                        If Trim(strArr(0)) = "" Then
                            ShowMsgbox "原始单据记录未找到相应的项目序号!" & vbCrLf & "单据号:" & Nvl(!NO) & vbCrLf & " 单据:" & Nvl(!单据, 0) & vbCrLf & " 行号为:" & Nvl(!序号)
                            Exit Function
                        End If
                        strTemp = strArr(0)
                        strInput = Nvl(!顺序号) & vbTab
                        strInput = strInput & InitInfor_兴安.医院编码 & vbTab
                        strInput = strInput & Val(strArr(0)) & vbTab
                        strInput = strInput & Nvl(!经办人)
                        
                        If 业务请求_兴安(住院明细取消, strInput, strOutput) = False Then Exit Function
                        strArr = Split(strOutput, vbTab)
                        
                        '为病人费用记录打上标记，以便随时上传
                        'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                        gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & ";;;;;" & vbTab & Val(strTemp) & vbTab & Val(strArr(0)) & "')"
                        zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                    Else
                    
                            '读取医保项目信息
                            If 业务请求_兴安(读取医保项目信息, Nvl(!编码), strOutput) = False Then Exit Function
                            
                            
                            If strOutput = "" Then
                                DebugTool "在进行处方明细上传函数中的读取医保信息时，返回串为空了"
                                Exit Function
                            End If
                            strArr = Split(strOutput, vbTab)
                            '入参:lsh(住院登记号),yycode(医院代码),rq(记帐日期),kbname(医生名称),ysname(科室名称),fycode(费用代码),fyname(费用名称),gg(规格),dw(单位),
                            '       dj(单价),sl(数量),je(金额),fylb(费用类别),ypbz(医保类别),ybdm(医保代码),czyname(记帐人)
                            '出参:bl(病人自付比例),xh(项目序号)
                            
                            strInput = Nvl(!顺序号) & vbTab
                            strInput = strInput & InitInfor_兴安.医院编码 & vbTab
                            strInput = strInput & Format(!经办时间, "yyyyMMDD") & vbTab
                            strInput = strInput & Nvl(!开单部门) & vbTab
                            strInput = strInput & Nvl(!开单医生) & vbTab
                            strInput = strInput & Nvl(!编码) & vbTab
                            strInput = strInput & Nvl(!商品名) & vbTab
                            strInput = strInput & Nvl(!规格) & vbTab
                            strInput = strInput & Nvl(!计算单位) & vbTab
                            strInput = strInput & Nvl(!实际价格, 0) & vbTab
                            strInput = strInput & Nvl(!数量, 0) & vbTab
                            strInput = strInput & Nvl(!实收金额, 0) & vbTab
                            strInput = strInput & strArr(0) & vbTab
                            strInput = strInput & strArr(1) & vbTab
                            strInput = strInput & strArr(2) & vbTab
                            strInput = strInput & Nvl(!经办人)
                                                        
                            If 业务请求_兴安(住院明细写入, strInput, strOutput) = False Then Exit Function
                            strArr = Split(strOutput, vbTab)
                            
                            '为病人费用记录打上标记，以便随时上传
                            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Val(strArr(0)) & vbTab & Val(strArr(1)) & "')"
                            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
                        End If
                    End If
                g病人身份_兴安.费用总额 = g病人身份_兴安.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    上传处方明细 = True
    Exit Function
ErrHand:
    
    DebugTool "上传处方明细失败!" & vbCrLf & "错误号:" & Err.Number & vbCrLf & "错误描述:" & Err.Description
 Resume
End Function
Private Function Get入出SQL(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As String
    Dim strSql As String
    
    strSql = "Select C.住院号,C.当前病区id,A.入院病床 ,c.住院号,to_char(A.确诊日期,'yyyyMMdd') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyyMMdd') 入院经办时间," & _
        " to_char(A.入院日期,'yyyyMMdd') 入院日期, A.出院方式,a.出院日期 ,a.出院病床,H.名称 as 出院科室,G.出院诊断 " & _
        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码||'-'||b.名称,'')) AS 入院诊断 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1  and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   D," & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码||'-'||b.名称,'')) AS 出院诊断 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 = 3 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   G" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        "       and A.主页id=G.主页id(+) and a.病人id=G.病人id(+) " & _
        ""
    Get入出SQL = strSql
End Function
Public Function 住院结算_兴安(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean

    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim lng主页ID As Long
    Dim strInput As String
    Dim strOutput As String
    Dim strArr
    Dim str结算方式 As String
    
    住院结算_兴安 = False
        
    DebugTool "进入住院结算接口"
    
    
    gstrSQL = Get明细记录(lng结帐ID)
    zlDatabase.OpenRecordset rs明细, gstrSQL, "获取结帐明细记录"
    
    '获取出院病人相关信息
    lng主页ID = Nvl(rs明细!主页ID, 0)
    gstrSQL = Get入出SQL(lng病人ID, Nvl(rs明细!主页ID, 0))
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取入出信息"
    
    Err = 0
    On Error GoTo ErrHand
    If 业务请求_兴安(住院事务开始, InitInfor_兴安.医院编码, "") = False Then Exit Function
    If 上传处方明细(rs明细) = False Then
        If 业务请求_兴安(住院事务回滚, "", "") = False Then Exit Function
        Exit Function
    End If
    
    If 业务请求_兴安(住院事务提交, "", "") = False Then
        Exit Function
    End If
     
    
    '住院结算
    strInput = g病人身份_兴安.住院登记号 & vbTab
    strInput = strInput & InitInfor_兴安.医院编码 & vbTab
    strInput = strInput & Format(rsTemp!出院日期, "yyyymmdd") & vbTab
    strInput = strInput & Nvl(rsTemp!出院诊断) & vbTab
    strInput = strInput & Nvl(rsTemp!出院科室) & vbTab
    
    gstrSQL = "Select 操作员姓名 From 病人结帐记录 where  ID=" & lng结帐ID
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取结算人"
    strInput = strInput & Nvl(rsTemp!操作员姓名) & vbTab
        
    If 业务请求_兴安(住院结算, strInput, strOutput) = False Then Exit Function
    strArr = Split(strOutput, vbTab)
    
    str结算方式 = str结算方式 & "||医保基金|" & Val(strArr(1))
    str结算方式 = str结算方式 & "||大病统筹|" & Val(strArr(2))
    'str结算方式 = str结算方式 & "||个人帐户|" & Val(strArr(0))
    
    
    str结算方式 = Mid(str结算方式, 3)
    
    '保存相关的结算信息
    gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',1)"
    Call ExecuteProcedure("更新预交记录")
    Dim intMouse As Integer
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    '显示结帐情况
    If frm结算信息.ShowMe(lng结帐ID) = False Then
        Exit Function
    End If
    Screen.MousePointer = intMouse
  
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN(本次个人帐户余额),累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN(费用总额),全自付金额_IN(本次自负段金额),首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN(本次统筹帐金额),大病自付金额_IN(大病记帐金额),超限自付金额_IN(本次自负金额),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(门诊流水号,住院：住院登记号),主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert( 2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & g病人身份_兴安.帐户余额 & ",Null,NULL," & g病人身份_兴安.当年第几次住院 & ",null,Null,NULL," & _
     g病人身份_兴安.费用总额 & "," & 0 & ",Null," & _
     "Null," & Val(strArr(1)) & "," & Val(strArr(2)) & "," & 0 & "," & Val(strArr(0)) & ",'" & _
     g病人身份_兴安.住院登记号 & "'," & lng主页ID & ",Null,NULl)"
    Call ExecuteProcedure("保存保险结算记录")
    
    DebugTool "住院结算成功"
    住院结算_兴安 = True
    Exit Function
ErrHand:
    
    DebugTool "住院结算(住院结算_兴安)" & vbCrLf & " 错误号:" & Err & vbCrLf & "错误信息:" & Err.Description
End Function

Public Function 住院结算冲销_兴安(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    
   

    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    
    Dim strOutput As String
    Dim strInput As String
    
    住院结算冲销_兴安 = False
    
    Err = 0
    On Error GoTo ErrHand
    DebugTool "进入住院结算冲销"
    
    gstrSQL = "Select * From 保险结算记录  where 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取住院登记号"
    
    lng病人ID = Nvl(rsTemp!病人ID, 0)
    lng主页ID = Nvl(rsTemp!主页ID, 0)
    
    lng冲销ID = Get冲销ID(lng结帐ID)
       
    strInput = Nvl(rsTemp!支付顺序号) & vbTab
    strInput = strInput & InitInfor_兴安.医院编码 & vbTab
    
    DebugTool "调用住院结算取消请求"
    '请求取消门诊结算
    If 业务请求_兴安(住院结算取消, strInput, strOutput) = False Then Exit Function
    
    DebugTool "进入保存保险结算记录"
   
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN(本次个人帐户余额),累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN(费用总额),全自付金额_IN(本次自负段金额),首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN(本次统筹帐金额),大病自付金额_IN(大病记帐金额),超限自付金额_IN(本次自负金额),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(门诊流水号),主页ID_IN,中途结帐_IN,备注_IN
    
    gstrSQL = "zl_保险结算记录_insert( 1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
      "NUll," & -1 * Nvl(rsTemp!帐户累计支出, 0) & ",Null,NULL," & Nvl(rsTemp!住院次数, 1) & ",null,Null,NULL," & _
     -1 * Nvl(rsTemp!发生费用金额, 0) & "," & -1 * Nvl(rsTemp!全自付金额, 0) & ",Null," & _
     "Null," & -1 * Nvl(rsTemp!统筹报销金额, 0) & "," & -1 * Nvl(rsTemp!大病自付金额, 0) & "," & -1 * Nvl(rsTemp!超限自付金额, 0) & "," & -1 * Nvl(rsTemp!个人帐户支付, 0) & ",'" & _
      Nvl(rsTemp!支付顺序号) & "'," & lng主页ID & ",Null,NULl)"
    Call ExecuteProcedure("保存保险结算记录")
    DebugTool "住院结算冲销成功"
    住院结算冲销_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    住院结算冲销_兴安 = False
End Function

Public Function 处方登记_兴安(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim rs明细 As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand:
    
    处方登记_兴安 = False
    
    
    '第一步: 读取费用明细记录
    gstrSQL = Get明细记录(0, str单据号, lng记录性质, lng记录状态)
    Call OpenRecordset(rs明细, "获取处方明细")
    
    If rs明细.RecordCount = 0 Then
        ShowMsgbox "没有明细记录!"
        Exit Function
    End If
    '入参:医院代码
    If 业务请求_兴安(住院事务开始, InitInfor_兴安.医院编码, "") = False Then Exit Function
    
    If 上传处方明细(rs明细) = False Then
        If 业务请求_兴安(住院事务回滚, "", "") = False Then Exit Function
        Exit Function
    End If
    
    If 业务请求_兴安(住院事务提交, "", "") = False Then
        Exit Function
    End If
    处方登记_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If 业务请求_兴安(住院事务回滚, "", "") = False Then Exit Function
End Function
Public Sub WriteParaInfor_兴安(ByVal strInfor As String)
        '将参数信息写入文件中
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim strFile As String
        
        strFile = "C:\接口交换串.log"
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForWriting)
        objText.WriteLine strInfor
        objText.Close
End Sub

Public Function 业务请求_兴安(ByVal int业务类型 As 业务类型_兴安, ByVal strInputString As String, ByRef strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInput As String, intReturn As Integer, strReturn As String
    Dim strOutput(0 To 10) As String, dblOutPut(0 To 10) As Double, intOutPut(0 To 5) As Integer
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim strReg As String
    
    Dim i As Integer
    
    DebugTool "进入业务请求函数(业务类型为:" & int业务类型 & "),输入参数为" & vbCrLf & strInputString
    
    
    业务请求_兴安 = False
    
    strInput = strInputString
    
    If InitInfor_兴安.模拟数据 Then
        '读取模拟数据
        Read模拟数据 int业务类型, strInputString, strOutPutstring
         业务请求_兴安 = True
        Exit Function
    End If
    
    
    Err = 0
    On Error GoTo ErrHand:
    
    strReg = "Z|" & int业务类型
    WriteParaInfor_兴安 strInputString
     SaveSetting "ZLSOFT", "医保", "数据交换", strReg
    If 请求等待() = False Then Exit Function
    strReg = GetSetting("ZLSOFT", "医保", "数据交换", "|")
    If strReg = "" Then strReg = "|"
    strArr1 = Split(strReg, "|")
    
    intReturn = Val(strArr1(1))
    Select Case int业务类型
        Case 建立数据库连接
            If intReturn < -99999 Then
                ShowMsgbox "医保部件加载失败!请与医保接口商联系!"
                Exit Function
            ElseIf intReturn < 0 Then:
                ShowMsgbox "打开医保中心的数据库连接失败!": Exit Function
            End If
        Case 关闭数据库连接
            If intReturn < 0 Then: ShowMsgbox "并闭医保中心的数据库连接失败!": Exit Function
        Case 操作员注册
            '入参:yycode(医院代码),czycode(操作员代码),czyname(操作员姓名),jylx(交易类型 1-普通医保门诊（在职、退休、镇南关）,2-特殊医保门诊（离休、二乙、特慢病）
            '出参:lsh(交易流水号)
            'intReturn = gobj兴安.Operator_Login(strArr(0), strArr(1), strArr(2), strArr(3), strOutput(0))
            '返串构建
            If intReturn = -4 Then
                ShowMsgbox "无效的操作员!": Exit Function
            ElseIf intReturn = -5 Then
                ShowMsgbox "不是本医院的操作员!": Exit Function
            ElseIf intReturn = -99 Then
                ShowMsgbox "医保部件内部错误,请与医保提供商联系!": Exit Function
            End If
            Read数据 strReturn
        Case 读取个人信息
            Select Case intReturn
            Case 0  '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据加连接！": Exit Function
            Case -3: ShowMsgbox "无效的卡号,请检查卡号是否正确！": Exit Function
            Case -4: ShowMsgbox "未进行年审,不能使用！": Exit Function
            Case -5: ShowMsgbox "该卡已经停止使用！": Exit Function
            Case -41: ShowMsgbox "未办理住院登记": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建
            Read数据 strReturn
        Case 读取医保项目信息
            '入参:yyfycode(医院费用代码)
            '出参:yplb(医保项目类别),ybbz(医保项目标志),ybdm(医保项目代码)
            Select Case intReturn
            Case Is >= 0 '操作成功
            Case -6: ShowMsgbox "无效的项目代码！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建
            Read数据 strReturn
        Case 门诊预处理
            '入参:xh(序号),je(金额),fylb(费用类别),ypbz(医保类别),ypcode(药品编码),cfts(处方天数)
            '出参:bczfdje(本次自负段金额),bctcje(本次统筹帐金额),bczfje(本次自负金额),bczhje(本次个人帐户金额),grzhye(本次个人帐户余额额),bz(备注)
            Select Case intReturn
            Case 0  '使用个人帐户
            Case 1  '使用自负段
            Case 2  '使用统筹
            Case 3  '超统筹限制
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -97: ShowMsgbox "上次购药未用完！": Exit Function
            Case -98: ShowMsgbox "个人帐户余额不足或使用了自费药品！"
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
            strArr1 = Split(strReturn, vbTab)
            If Trim(strArr1(6)) <> "" Then
                ShowMsgbox strArr1(6)
            End If
            
        Case 门诊明细写入
            Select Case intReturn
            Case Is >= 0 '操作成功,返加帐户余额
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -21: ShowMsgbox "存普通明细记录失败！": Exit Function
            Case -31: ShowMsgbox "存特殊医保明细记录失败！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
        Case 门诊结算提交
            
            '入参:无
            '出参:mzcode(门诊流水号),grzhye(帐户余额),bczhje(本次交易金额),xjzf(本次现金自负额)
            
            Select Case intReturn
            Case Is >= 0  '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -22: ShowMsgbox "未写入门诊明细记录！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
            
        Case 门诊结算冲销
            '入参:mzcode(门诊流水号)
            '出参:无
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -23: ShowMsgbox "没有此笔交易！": Exit Function
            Case -24: ShowMsgbox "此笔交易已取消！": Exit Function
            Case -25: ShowMsgbox "此笔交易已结算！": Exit Function
            Case -26: ShowMsgbox "数据有误！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
        Case 住院登记
            '入参:lsh(住院登记号),yycode(医院代码),ryrq(入院日期),zyh(住院号),kbname(科室名称),ysname(医生姓名),cwcode(床位号),ryzd(入院诊断),zt(操作状态(登记 修改)
            '出参:无
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -41: ShowMsgbox "未办理住院登记审批！": Exit Function
            Case -44: ShowMsgbox "未办理住院登记！": Exit Function
            Case -42: ShowMsgbox "无效的操作状态！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            strReturn = ""
        Case 取消住院登记
            '入参:    lsh(住院登记号),yycode(医院代码)
            '出参:
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -41: ShowMsgbox "未办理住院登记审批！": Exit Function
            Case -43: ShowMsgbox "有记帐费用，不能取消！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            strReturn = ""
        Case 住院明细写入
            '入参:lsh(住院登记号),yycode(医院代码),rq(记帐日期),kbname(医生名称),ysname(科室名称),fycode(费用代码),fyname(费用名称),gg(规格),dw(单位),dj(单价),sl(数量),je(金额),fylb(费用类别),ypbz(医保类别),ybdm(医保代码),czyname(记帐人)
            '出参:bl(病人自付比例),xh(项目序号)
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -44: ShowMsgbox "未办理住院登记！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
        Case 住院结算
            '入参:lsh(住院登记号),yycode(医院代码),cyrq(出院日期),Cyzd(出院诊断),kbname(出院科室),czyname(结算人)
            '出参:xjzf(病人应付现金),tcjzje(统筹记帐金额),dbjzje(大病记帐金额)
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -45: ShowMsgbox "未办理住院结算！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
        Case 住院结算取消
            '入参:lsh(住院登记号),yycode(医院代码)
            '出参:无
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -45: ShowMsgbox "未办理住院结算！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            strReturn = ""
        Case 住院明细取消
            '入参:lsh(住院登记号),yycode(医院代码),czyname(记帐人),xh(项目序号)

            '出参:
            Select Case intReturn
            Case 0   '操作成功
            Case -2: ShowMsgbox "没有数据库连接,请检查医保数据库是否已经连接好！": Exit Function
            Case -44: ShowMsgbox "未办理住院登记！": Exit Function
            Case -46: ShowMsgbox "没有相应的费用！": Exit Function
            Case Else: ShowMsgbox "在调取兴安医保接口时出现内部错误,请与接口供应商联系！": Exit Function '-99
            End Select
            '返串构建,在传出参数串后面多了个返回值
            Read数据 strReturn
        Case 住院事务开始
            Select Case intReturn
            Case 0   '操作成功
            Case -1: ShowMsgbox "住院事务开始失败！": Exit Function
            End Select
            '返串构建,在传出参数串后面多了个返回值
        Case 住院事务提交
            Select Case intReturn
            Case 0   '操作成功
            Case -1: ShowMsgbox "住院事务提交失败！": Exit Function
            End Select
        Case 住院事务回滚
            Select Case intReturn
            Case 0   '操作成功
            Case -1: ShowMsgbox "住院事务回滚失败！": Exit Function
            End Select
    End Select
    strOutPutstring = strReturn
    业务请求_兴安 = True
    DebugTool "业务请求成功(业务类型为:" & int业务类型 & ")." & vbCrLf & "输入参数为" & vbCrLf & strInputString & vbCrLf & "输出参数为:" & vbCrLf & strReturn
    Exit Function
ErrHand:
    DebugTool "业务请求失败(业务类型为:" & int业务类型 & ")." & vbCrLf & "输入参数为" & vbCrLf & strInputString & vbCrLf & "输出参数为:" & vbCrLf & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Read模拟数据(ByVal int业务类型 As 业务类型_兴安, ByVal strInputString As String, ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    strFile = App.Path & "\模拟提交串.txt"
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Select Case int业务类型
    Case 建立数据库连接
        Exit Function
    Case 关闭数据库连接
        Exit Function
    Case 操作员注册
        strName = "操作员注册"
    Case 读取个人信息
        strName = "读取个人信息"
    Case 读取医保项目信息
        strName = "读取医保项目信息"
    Case 门诊预处理
        strName = "门诊预处理"
    Case 门诊明细写入
        strName = "门诊明细写入"
    Case 门诊结算提交
        strName = "门诊结算提交"
    Case 门诊结算冲销
        strName = "门诊结算冲销"
    Case 住院登记
        strName = "住院登记"
    Case 取消住院登记
        Exit Function
    Case 住院明细写入
        Exit Function
    Case 住院明细取消
        Exit Function
    Case 住院结算
        strName = "住院结算"
    Case 住院结算取消
        strName = "住院结算取消"
    End Select
   
    Dim blnStart As Boolean
    Dim strArr
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            blnStart = False
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                    
                If blnStart Then
                    If strText = "" Then
                        strText = "" & vbTab
                    End If
                    strArr = Split(strText, "|")
                    
                    If Val(strArr(0)) = 1 Then
                        str = strArr(1)
                        Exit Do
                    End If
                Else
                     If "<" & strName & ">" = strText Then
                         blnStart = True
                     End If
                End If
                If "</" & strName & ">" = strText Then
                    Exit Do
                End If
            Loop
            objText.Close
            strOutPutstring = str
    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Public Function 挂号结算_兴安(ByVal lng结帐ID As Long) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    
    挂号结算_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 挂号冲销_兴安(ByVal lng结帐ID As Long) As Boolean

    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Err = 0
    On Error GoTo ErrHand
    
    挂号冲销_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub WriteDebugInfor_兴安(ByVal strInfor As String)
        '将调试信息写入文件中
        Dim objFile As New FileSystemObject
        Dim objText As TextStream
        Dim intDebug As Integer
        
        intDebug = GetSetting("ZLSOFT", "医保", "将串写入文本文件", 0)
        If intDebug <> 1 Then Exit Sub

        Dim strFile As String
        Dim rsTemp As New ADODB.Recordset
        strFile = App.Path & "\Test.log"
        
        If Not Dir(strFile) <> "" Then
            objFile.CreateTextFile strFile
        End If
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfor
        objText.Close
End Sub
Public Function 在院病人信息_兴安(lng病人ID As Long, lng主页ID As Long) As Boolean
 
    
    在院病人信息_兴安 = False
    On Error GoTo ErrHand
    DebugTool "进入在院病人信息接口"
    
    If 住院信息提交(lng病人ID, lng主页ID, True) = False Then Exit Function
    
    DebugTool "在院病人信息修改成功"
    在院病人信息_兴安 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 住院信息提交(lng病人ID As Long, lng主页ID As Long, Optional bln修改 As Boolean = False) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutput As String
    
    '写住院信息
    DebugTool "读取入院的相关信息"
    Err = 0
    On Error GoTo ErrHand:
    住院信息提交 = False
    
    '获取相关病人信息
    gstrSQL = "Select C.住院号,C.当前床号,to_char(A.确诊日期,'yyyyMMdd') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyyMMdd') 入院经办时间," & _
        " to_char(A.登记时间,'yyyyMMdd') 入院日期  ,to_char(A.登记时间,'yyyyMMdd') 入院时间,D.入院诊断 " & _
        " From 病案主页 A,部门表 B,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码||'-'||b.名称,'')) AS 入院诊断 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页ID & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        ""
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取入院信息"
    
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    '入参:住院登记号,医院代码,入院日期,住院号,科室名称,医生姓名,床位号,入院诊断,操作状态
    strInput = g病人身份_兴安.住院登记号 & vbTab
    strInput = strInput & InitInfor_兴安.医院编码 & vbTab
    strInput = strInput & Nvl(rsTemp!入院日期) & vbTab
    strInput = strInput & Nvl(rsTemp!住院号) & vbTab
    strInput = strInput & Nvl(rsTemp!入院科室) & vbTab
    strInput = strInput & Nvl(rsTemp!住院医师) & vbTab
    strInput = strInput & Nvl(rsTemp!当前床号) & vbTab
    strInput = strInput & Nvl(rsTemp!入院诊断) & vbTab
    strInput = strInput & IIf(bln修改, "修改", "登记")
    
    DebugTool "调用住院修改请求"
    
    If 业务请求_兴安(住院登记, strInput, strOutput) = False Then
        Exit Function
    End If
    DebugTool "在院病人信息写入成功"
    
    住院信息提交 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetItemInfo_兴安(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal str摘要 As String, Optional intType As Integer = 0) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取大连病人的相关提示信息
    '--入参数:
    '--出参数:
    '--返  回:提示串
    '-----------------------------------------------------------------------------------------------------------
    Dim strMsgInfor As String
    Dim str原摘要 As String
    str原摘要 = str摘要
    If g病人身份_兴安.交易类型 = "普通医保门诊" Then
        GetItemInfo_兴安 = str原摘要
        Exit Function
    End If
    strMsgInfor = str摘要
    If frm处方信息输入_兴安.EditCard(strMsgInfor) = False Then
        GetItemInfo_兴安 = str原摘要
        Exit Function
    End If
    GetItemInfo_兴安 = strMsgInfor
End Function
Private Function Read数据(ByRef strOutPutstring As String)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以便测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strTemp As String
    Dim strFile As String
    Dim str As String
    Dim strName As String
    
    strFile = "C:\接口交换串.log"
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    Err = 0
    On Error GoTo ErrHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            str = ""
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                Exit Do
            Loop
            objText.Close
            strOutPutstring = strText
    End If
    Exit Function
ErrHand:
    Exit Function
End Function

Private Function 请求等待() As Boolean
    '等待数据处理,true处理成功,fale处理失败
    Dim strReg As String
    Dim strArr1
    请求等待 = False
    
    Dim strDate As String
    strDate = Format(DateAdd("s", 400, Now), "yyyymmdd HH:MM:SS")
    Do While True
        '等待交易完成
        strReg = GetSetting("ZLSOFT", "医保", "数据交换", "|")
        strArr1 = Split(strReg, "|")
        If strArr1(0) = "H" Then
            Exit Do
        End If
        
        If strArr1(0) = "" Then
            Exit Function
        End If
        If Format(Now, "yyyymmdd HH:MM:SS") >= strDate Then
            '交易等待过长，将取消本次交易,
            ShowMsgbox "交易等待过长，将取消本次交易"
            Exit Function
        End If
    Loop
    请求等待 = True
End Function


