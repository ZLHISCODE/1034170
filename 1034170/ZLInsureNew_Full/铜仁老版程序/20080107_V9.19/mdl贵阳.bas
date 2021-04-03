Attribute VB_Name = "mdl贵阳"
Option Explicit

Public mdomInput As MSXML2.DOMDocument
Public mdomOutput As MSXML2.DOMDocument

Private mstr卡号 As String
Private mstr密码 As String

Private mstr医保号 As String
Private mdbl余额 As Double

Private mlng病人ID As Long

Public Function 医保初始化_贵阳() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    If Err <> 0 Then
        MsgBox "不能创建XML分析器，请注册msxml3.dll部件。", vbInformation, gstrSysName
    Else
        医保初始化_贵阳 = True
    End If
End Function

Public Function 身份标识_贵阳(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str保险类别 As String, str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, cur帐户余额 As Currency
    Dim str姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str人员类别 As String, str单位编码 As String, str单位名称 As String
    Dim strIdentify As String, str附加 As String, lng病种ID As Long
    Dim rsTemp As New ADODB.Recordset, rs病种 As ADODB.Recordset
    
    '初始化一些变量，在程序中途退出时值却已经赋了
    mstr卡号 = "": mstr密码 = ""
    If frmIdentify贵阳.GetIdentify(TYPE_贵阳市, str卡号, str医保号, str分中心编号, str密码, True, True) = False Then
        Exit Function
    End If
    '还原数据
    str保险类别 = Split(str卡号, "^")(1)
    str卡号 = Split(str卡号, "^")(0)
    
    If bytType = id门诊确认 Then
        '该返回值暂时没有作用，只要不为空就表示成功了
        身份标识_贵阳 = str卡号 & ";" & str医保号 & ";" & str密码
        Exit Function
    End If
    
    '取得返回值
    str姓名 = GetElemnetValue("PERSONNAME")
    str性别 = GetElemnetValue("SEX")
    str性别 = Switch(str性别 = "1", "男", str性别 = "2", "女", str性别 = "9", "其它", True, str性别)
    str身份证号码 = GetElemnetValue("PID")
    
    str出生日期 = AddDate(GetElemnetValue("BIRTHDAY"))
    If IsDate(str出生日期) = True Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), zlDatabase.Currentdate)
    Else
        str出生日期 = ""
    End If
    
    str人员类别 = GetElemnetValue("PERSONTYPE")
    str人员类别 = Switch(str人员类别 = "11", "在职", str人员类别 = "21", "退休" _
                      , str人员类别 = "32", "省属离休", str人员类别 = "34", "市属离休", True, "其他")
    str单位编码 = ToVarchar(GetElemnetValue("DEPTCODE"), 12)
    str单位名称 = ToVarchar(GetElemnetValue("DEPTNAME"), 36) '字段长度本是50，但由于还要保存编码及括号
    cur帐户余额 = Val(GetElemnetValue("ACCTBALANCE"))
    
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    strIdentify = str卡号 & ";" & str医保号 & ";" & str密码 & ";" & str姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";" & str单位名称 & "(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    '特殊门诊
    'Modified By 朱玉宝 2003-12-03 地区： 原因：入院时取消病种选择，改为在虚拟结算时，如果没有病种，必需选择
    If bytType = id门诊收费 And Get保险参数_贵阳("支持特殊门诊") = "1" Then
        gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                " From 保险病种 A where A.险类=" & gintInsure
        
        Set rs病种 = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
        If Not rs病种 Is Nothing Then
            lng病种ID = rs病种("ID")
        End If
    End If
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";" & str分中心编号             '9.顺序号  但本医保用于保存医保分中心编码（避免建立医保中心）
    str附加 = str附加 & ";" & str人员类别               '10人员身份
    str附加 = str附加 & ";" & cur帐户余额               '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";" & IIf(lng病种ID <> 0, lng病种ID, "")   '13病种ID
    str附加 = str附加 & ";" & IIf(str人员类别 = "在职", 1, 2)      '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & cur帐户余额               '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";"                             '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID)
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_贵阳 = strIdentify & ";" & lng病人ID & str附加
        
        mstr卡号 = str卡号
        mstr密码 = str密码
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 更改密码_贵阳市(ByVal str磁卡数据 As String, ByVal str密码 As String, ByVal str新密码 As String) As Boolean
    If InitXML = False Then Exit Function
    
    Call InsertChild(mdomInput.documentElement, "CARDDATA", str磁卡数据)            ' 磁卡数据
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)                ' 密码
    Call InsertChild(mdomInput.documentElement, "NEWPASSWORD", str新密码)           ' 密码
    
    '调用接口
    If CommServer("MODIFYCARD") = False Then Exit Function
    更改密码_贵阳市 = True
End Function

Public Function 个人余额_贵阳(strSelfNo As String) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    '从数据库中读取（因为刚才才保存了的，应该是准确的）
    If mstr医保号 = "" Or strSelfNo <> mstr医保号 Then
        gstrSQL = "Select 帐户余额 From 保险帐户 where 险类=" & gintInsure & " and 中心=0 and 医保号='" & strSelfNo & "'"
        Call OpenRecordset(rsTemp, "贵阳医保")
        
        If rsTemp.EOF = False Then
            个人余额_贵阳 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
        End If
    Else
        个人余额_贵阳 = mdbl余额
    End If
    '只能用一次
    mstr医保号 = ""
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_贵阳(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim dbl个人帐户 As Double
    Dim lng病人ID As Long, str疾病编码 As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_贵阳 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    
    '判断该病人是否是特殊门诊
    gstrSQL = "select A.人员身份,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=" & lng病人ID & " and A.险类=" & gintInsure & "  and A.病种ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "门诊预算")
    If rsTemp.EOF = False Then
        str疾病编码 = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21" _
                      , str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", True, "11")
    End If
    datCurr = zlDatabase.Currentdate
    
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str人员类别)         ' 密码
    If str疾病编码 <> "" Then '特殊门诊
        '补足8位长度
        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)         '特种病编码
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '待遇开始享受时间
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' 是否结算
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' 账户支付额
    Call InsertChild(mdomInput.documentElement, "INVOICENO", " ") ' 发票号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
    
    Do Until rs明细.EOF
        gstrSQL = "SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM 药品目录 C,药品信息 D,药品剂型 E WHERE C.药品ID=" & rs明细("收费细目ID") & " AND C.药名ID=D.药名ID AND D.剂型=E.编码"
        gstrSQL = "select A.类别,A.名称,B.项目编码,nvl(A.规格,F.规格) AS 规格,F.剂型,A.计算单位 from 收费细目 A,保险支付项目 B,(" & gstrSQL & _
                ") F where A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID  AND A.Id=F.药品ID(+) and B.险类=" & gintInsure
        Call OpenRecordset(rsTemp, "门诊预算")
        If rsTemp.EOF = True Then
            MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
            Exit Function
        End If

        Set nodRow = InsertChild(nodRowset, "ROW", "")
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(rsTemp("项目编码"), 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsTemp("名称"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("类别")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("规格"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("剂型"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("计算单位"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs明细("单价"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs明细("数量"), "0.00"))
        Call nodRow.setAttribute("FROMOFFICE", ToVarchar(UserInfo.部门, 56)) '开单科室
        Call nodRow.setAttribute("FROMDOCT", Format(UserInfo.姓名, 20))      '开单医生
        Call nodRow.setAttribute("TOOFFICE", ToVarchar(UserInfo.部门, 56))  '受单科室
        Call nodRow.setAttribute("TODOCT", Format(UserInfo.姓名, 20))       '受单医生
        Call nodRow.setAttribute("DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))        '办理日期
        Call nodRow.setAttribute("NOTE", ToVarchar(rs明细("摘要"), 512))        '备注
        
        rs明细.MoveNext
    Loop
    
    '调用接口
    If CommServer(IIf(str疾病编码 <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    '离休人员不存在普通门诊与特殊门诊，统一由ALLOWFUND支付；
    '基本医疗人员特殊门诊由FUND1PAY及FUND2PAY支付，普通门诊由个人帐户支付
    If str人员类别 = "32" Or str人员类别 = "34" Then
        str结算方式 = "医保基金;" & Val(GetElemnetValue("ALLOWFUND")) & ";0"
    Else
        str结算方式 = "个人帐户;" & Val(GetElemnetValue("ACCTPAY")) & ";1"  '允许修改个人帐户
        If str疾病编码 <> "" Then
            str结算方式 = str结算方式 & "|医保基金;" & Val(GetElemnetValue("FUND1PAY")) & ";0" & _
                         "|大病统筹;" & Val(GetElemnetValue("FUND2PAY")) & ";0"
        End If
    End If
    门诊虚拟结算_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_贵阳(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
'    Dim str卡号Re As String, str医保号Re As String, str分中心编号Re As String, str密码Re As String
    Dim str医生 As String, str科室 As String, cur发生费用 As Double, datCurr As Date
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng病人ID  As Long, str疾病编码   As String, lng项目数 As Long
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    lng项目数 = Val(Get保险参数_贵阳("门诊最大项目数"))

    gstrSQL = "SELECT Nvl(从属父号,Nvl(价格父号,序号)) AS 主序号 FROM 病人费用记录  " & _
             " WHERE 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0" & _
             " GROUP BY Nvl(从属父号,Nvl(价格父号,序号))"
    Call OpenRecordset(rs明细, "贵阳医保")
    If rs明细.RecordCount > lng项目数 Then
        MsgBox "门诊收费的项目数不能超过" & lng项目数 & "。", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select A.ID,A.序号,A.病人ID,A.NO,A.登记时间,A.开单人 as 医生,A.登记时间," & _
            "   A.数次*A.付数 as 数量,A.标准单价 as 实际价格,A.结帐金额," & _
            "   A.收费类别,D.项目编码,B.名称 as 项目名称,C.名称 as 科室名称,nvl(B.规格,F.规格) AS 规格,F.剂型,B.计算单位,A.摘要 " & _
            " From (Select * From 病人费用记录 Where 结帐ID=" & lng结帐ID & ") A,收费细目 B,部门表 C,保险支付项目 D " & _
            "     ,(SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM 病人费用记录 A,药品目录 C,药品信息 D,药品剂型 E WHERE A.结帐ID=" & lng结帐ID & " AND A.收费细目ID=C.药品ID AND C.药名ID=D.药名ID AND D.剂型=E.编码) F " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID  AND B.ID=F.药品ID(+) And D.险类=" & gintInsure & " And Nvl(A.附加标志,0)<>9 And Nvl(A.记录状态,0)<>0" & _
            " Order by A.ID"
    Call OpenRecordset(rs明细, "贵阳医保")
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str医生 = ToVarchar(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 20)
    str科室 = ToVarchar(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 56)
    datCurr = zlDatabase.Currentdate
    
    
    '一、费用明细传递
    
    '判断该病人是否是特殊门诊
    gstrSQL = "select A.人员身份,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=" & lng病人ID & " and A.险类=" & gintInsure & "  and A.病种ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "门诊预算")
    If rsTemp.EOF = False Then
        str疾病编码 = IIf(IsNull(rsTemp("编码")), "", rsTemp("编码"))
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21" _
                      , str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", True, "11")
    End If
    
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    '门诊收费时需要再刷一次卡
'    If frmIdentify贵阳.GetIdentify(TYPE_贵阳市, str卡号Re, str医保号Re, str分中心编号Re, str密码Re, False) = False Then
'        Exit Function
'    Else
'        If str卡号 <> str卡号Re Or str医保号 <> str医保号Re Then
'            MsgBox "请使用当前病人的卡再刷一次。", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
        
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str人员类别)         ' 密码
    If str疾病编码 <> "" Then '特殊门诊
        '补足8位长度
        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)         '特种病编码
        Call InsertChild(mdomInput.documentElement, "STARTDATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) '待遇开始享受时间
    End If
    Call InsertChild(mdomInput.documentElement, "ISCAL", 1)         ' 是否结算
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", Format(cur个人帐户, "0.00")) ' 账户支付额
    Call InsertChild(mdomInput.documentElement, "INVOICENO", rs明细("NO")) ' 发票号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
    
    Do Until rs明细.EOF
        cur发生费用 = cur发生费用 + rs明细("结帐金额")
        
        Set nodRow = InsertChild(nodRowset, "ROW", "")
        Call nodRow.setAttribute("ITEMCODE", ToVarchar(rs明细("项目编码"), 12))
        Call nodRow.setAttribute("ITEMNAME", ToVarchar(rs明细("项目名称"), 72))
        Call nodRow.setAttribute("SUBJECT", Subject(rs明细("收费类别")))
        Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rs明细("规格"), 40))
        Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rs明细("剂型"), 20))
        Call nodRow.setAttribute("UNIT", ToVarchar(rs明细("计算单位"), 20))
        Call nodRow.setAttribute("PRICE", Format(rs明细("实际价格"), "0.0000"))
        Call nodRow.setAttribute("QUANTITY", Format(rs明细("数量"), "0.00"))
        Call nodRow.setAttribute("FROMOFFICE", str科室)    '开单科室
        Call nodRow.setAttribute("FROMDOCT", str医生)      '开单医生
        Call nodRow.setAttribute("TOOFFICE", str科室)     '受单科室
        Call nodRow.setAttribute("TODOCT", str医生)       '受单医生
        
        '处理时间时，为了保证同一保险项目的的收费时间不同，因此在登记时间上按序号加上秒数
        Call nodRow.setAttribute("DODATE", Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss"))    '办理日期
        Call nodRow.setAttribute("NOTE", ToVarchar(rs明细("摘要"), 512))         '备注
        
        rs明细.MoveNext
    Loop
    
    '调用接口
    If CommServer(IIf(str疾病编码 <> "", "CALSPECCLIN", "CALCLIN")) = False Then Exit Function
    
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
            
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, cur起付线 As Double
    Dim str就诊顺序号 As String, str结算编号 As String
            
    cur全自付 = Val(GetElemnetValue("FEEOUT"))
    cur挂钩自付 = Val(GetElemnetValue("FEESELF"))
    cur起付线 = Val(GetElemnetValue("STARTFEE"))
    cur基数自付 = Val(GetElemnetValue("ENTERSTARTFEE"))
    If str人员类别 = "32" Or str人员类别 = "34" Then
        cur统筹支付 = Val(GetElemnetValue("ALLOWFUND"))
        cur大病统筹 = 0
    Else
        cur统筹支付 = Val(GetElemnetValue("FUND1PAY"))
        cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
    End If
    cur统筹自付 = Val(GetElemnetValue("FUND1SELF"))
    cur大病自付 = Val(GetElemnetValue("FUND2SELF"))
    cur超限自付 = Val(GetElemnetValue("FEEOVER"))
    
    str结算编号 = GetElemnetValue("BALANCEID")
    str就诊顺序号 = GetElemnetValue("BILLNO")
    If str疾病编码 <> "" Then
        str就诊顺序号 = "特殊" & str疾病编码 & str就诊顺序号 '把疾病编码与就诊顺序号连在一起
    Else
        str就诊顺序号 = "普通" & str就诊顺序号         '表示普通门诊
    End If
    
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 + cur统筹支付 + cur大病统筹 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",0," & cur基数自付 & "," & cur发生费用 & "," & _
        cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & str结算编号 & "',null,null,'" & str就诊顺序号 & "')"
    Call ExecuteProcedure("贵阳医保")
    '---------------------------------------------------------------------------------------------
    
    门诊结算_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_贵阳(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim str结算编号 As String, str就诊顺序号 As String, curDate As Date, rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, int住院次数累计 As Integer
    Dim lng冲销ID As Long
    
    On Error GoTo errHandle
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "门诊退费")
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "门诊退费")
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID
    cur个人帐户 = Nvl(rsTemp!个人帐户支付, 0)
    str结算编号 = Nvl(rsTemp("支付顺序号"), "")
    str就诊顺序号 = Nvl(rsTemp("备注"), "")
    If str就诊顺序号 = "" Then
        MsgBox "该单据没有保存就诊顺序号，不能做废。", vbInformation, gstrSysName
        Exit Function
    End If
    If Left(str就诊顺序号, 2) = "特殊" Then
        MsgBox "目前不支持特殊门诊的作废。", vbInformation, gstrSysName
        Exit Function
    End If
    str就诊顺序号 = Mid(str就诊顺序号, 3)
    curDate = zlDatabase.Currentdate
    
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID, True) = False Then Exit Function
        
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "BILLNO", str就诊顺序号)     ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str结算编号)    ' 结算编号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)    ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(curDate, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    
    '调用接口
    If CommServer("RETCLIN") = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str结算编号 & "',null,null,'" & Nvl(rsTemp("备注"), "") & "')"
    Call ExecuteProcedure("贵阳医保")
    
    门诊结算冲销_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人帐户转预交_贵阳(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
    个人帐户转预交_贵阳 = False
End Function

Public Function 入院登记_贵阳(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim str卡号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str疾病编码 As String
    Dim strtemp As String, str提示 As String, str诊断 As String, lng参保前在院 As Long
    
    On Error GoTo errHandle
    
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    '判断该病人是否参保前在院
    lng参保前在院 = 0
    If Get保险参数_贵阳("入院时选择参保前在院") = "1" Then
        If MsgBox("该病人参保前是否已经在院？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            lng参保前在院 = 1
        End If
    End If
    
    '判断该病人是否是特殊病
    gstrSQL = "select A.人员身份,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=" & lng病人ID & " and A.险类=" & gintInsure & "  and A.病种ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "入院登记")
    If rsTemp.EOF = False Then
        str疾病编码 = Nvl(rsTemp("编码"), "")
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21" _
                      , str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", True, "11")
    End If
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型=1 and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "出院登记")
    If rsTemp.EOF = False Then
        str诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
    Else
        str诊断 = "疾病"   '诊断不论如何不能为空
    End If
    
    '获得其它入院信息
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.入院方式,nvl(A.二级院转入,0) as 二级院转入,A.门诊医师,A.入院日期,A.入院病床,B.名称 as 入院科室,C.住院号 from 病案主页 A,部门表 B,病人信息 C " & _
              " Where A.病人ID=C.病人ID and A.入院科室ID = B.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "入院登记")
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str人员类别)   ' 人员类别
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", IIf(rsTemp("入院方式") = "转入", "37", "31"))     ' 支付类别 31：住院；37：转院
    
    If str疾病编码 <> "" Then
        '补足8位长度
        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' 转大额标志
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)   ' 特种病编码
    Else
        '没有特殊病
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "")            ' 转大额标志
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", "")      ' 特种病编码
    End If
    
    Call InsertChild(mdomInput.documentElement, "HOSPNO", ToVarchar(rsTemp("住院号"), 20))     ' 住院号
    Call InsertChild(mdomInput.documentElement, "ISINHOSP", lng参保前在院)     ' 参保前已在院 1：是；0：否
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", str诊断) ' 诊断
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("门诊医师"), 20)) ' 诊断医生
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("入院科室"), 20)) ' 科室
    Call InsertChild(mdomInput.documentElement, "POSITION", ToVarchar(rsTemp("入院病床"), 10)) ' 床位
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("入院日期"), "yyyy-MM-dd HH:mm:ss")) ' 入院时间
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    
    '调用接口
    If CommServer("HOSPREG") = False Then Exit Function
    
    Dim int住院次数累计 As Integer
    Dim cur本次起付线 As Currency
    Dim cur起付线累计 As Currency
    Dim cur基本统筹限额 As Currency
    Dim cur统筹报销累计 As Currency
    Dim cur大额统筹限额 As Currency
    Dim cur大额统筹累计 As Currency
    
    Dim str封锁信息 As String
    
    int住院次数累计 = Val(GetElemnetValue("HOSPTIMES"))
    
    cur本次起付线 = Val(GetElemnetValue("STARTFEE"))
    cur起付线累计 = Val(GetElemnetValue("STARTFEEPAID"))
    cur基本统筹限额 = Val(GetElemnetValue("FUND1LMT"))
    cur统筹报销累计 = Val(GetElemnetValue("FUND1PAID"))
    cur大额统筹限额 = Val(GetElemnetValue("FUND2LMT"))
    cur大额统筹累计 = Val(GetElemnetValue("FUND2PAID"))
    
    str封锁信息 = GetElemnetValue("LOCKINFO")
    Do Until str封锁信息 = ""
        strtemp = Left(str封锁信息, 2)
        str封锁信息 = Mid(str封锁信息, 41)
        
        str提示 = str提示 & Switch(strtemp = "11", "、待遇审核期", strtemp = "21", "、卡封锁", strtemp = "31", "、基本统筹欠费", _
                                   strtemp = "32", "、大额统筹未缴费", strtemp = "41", "、停保", strtemp = "51", "、退保")
        
    Loop
    If str提示 <> "" Then
        MsgBox "请注意该医保病人情况：" & Mid(str提示, 2) & "。", vbInformation, gstrSysName
    End If
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        "0,0,0," & cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & _
         "," & cur基本统筹限额 & "," & cur大额统筹限额 & "," & cur大额统筹累计 & ",'" & ToVarchar(str提示, 100) & "')"
    Call ExecuteProcedure("贵阳医保")
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("贵阳医保")
    
    入院登记_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_贵阳(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim str医保号 As String, str分中心编号 As String
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str诊断 As String, str其它诊断 As String
    Dim str病案号 As String, str出院转归 As String, lngPos As Long
    
    On Error GoTo errHandle
    
    '从数据库中读出已存储的值
    gstrSQL = "select 卡号,医保号,顺序号 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "出院登记")
    
    str医保号 = IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号"))
    str分中心编号 = IIf(IsNull(rsTemp("顺序号")), "", rsTemp("顺序号"))
    
    '获得病人出院信息
    gstrSQL = "SELECT A.出院方式,nvl(C.病案号,B.住院号) AS 病案号  " & _
             " FROM 病案主页 A,病人信息 B,住院病案记录 C " & _
             " WHERE A.病人ID=" & lng病人ID & " AND A.主页id=" & lng主页ID & " AND A.病人id=B.病人id AND A.病人id=C.病人id(+)"
    Call OpenRecordset(rsTemp, "出院登记")
    str病案号 = rsTemp("病案号")
    Select Case rsTemp("出院方式")
        Case "正常", "治愈"
            str出院转归 = "1"
        Case "好转"
            str出院转归 = "2"
        Case "死亡"
            str出院转归 = "3"
        Case Else
            str出院转归 = "9"
    End Select
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型=3 and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "出院登记")
    If rsTemp.EOF = False Then
        str诊断 = Nvl(rsTemp("描述信息"), "疾病")
        '将不同形式的分隔符统一
        str诊断 = Replace(str诊断, "，", ",")
        str诊断 = Replace(str诊断, "；", ",")
        str诊断 = Replace(str诊断, "、", ",")
        str诊断 = Replace(str诊断, ";", ",")
        lngPos = InStr(str诊断, ",")
        If lngPos > 0 Then
            str其它诊断 = Mid(str诊断, lngPos + 1)
            str诊断 = Mid(str诊断, 1, lngPos - 1)
        End If
    Else
        str诊断 = "疾病"   '诊断不论如何不能为空
    End If
        
    '获得其它出院信息
    datCurr = zlDatabase.Currentdate
    gstrSQL = "select A.住院医师,A.出院日期,A.出院病床,B.名称 as 出院科室 from 病案主页 A,部门表 B " & _
             " Where A.出院科室ID = B.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "出院登记")
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "DOCNO", str病案号)          ' 病案号
    Call InsertChild(mdomInput.documentElement, "DIAGNOSES", ToVarchar(str诊断, 128))          ' 诊断
    Call InsertChild(mdomInput.documentElement, "OTHERDIAGNOSES", ToVarchar(str其它诊断, 128)) ' 其他诊断
    Call InsertChild(mdomInput.documentElement, "OUTTYPE", str出院转归)                        ' 转归类别
    Call InsertChild(mdomInput.documentElement, "DOCTOR", ToVarchar(rsTemp("住院医师"), 20))   ' 诊断医生
    Call InsertChild(mdomInput.documentElement, "OFFICE", ToVarchar(rsTemp("出院科室"), 20))   ' 科室
    'Call InsertChild(mdomInput.documentElement, "POSITION", ToVarchar(rsTemp("出院病床"), 10)) ' 床位
    Call InsertChild(mdomInput.documentElement, "REGDATE", Format(rsTemp("出院日期"), "yyyy-MM-dd HH:mm:ss")) ' 出院日期
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    
    '调用接口
    If CommServer("HOSPOUT") = False Then Exit Function
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("贵阳医保")
    
    出院登记_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function 住院虚拟结算_贵阳(rsExse As Recordset, ByVal lng病人ID As Long) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs病种 As ADODB.Recordset
    Dim lng病种ID As Long, str疾病编码 As String
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim cur个人帐户 As Double, cur统筹支付 As Double, cur大病统筹 As Double, cur发生费用 As Double
    Dim str医生 As String, str科室 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    On Error GoTo errHandle
    mlng病人ID = 0         '初始化。只要一选择病人，就会调用本过程，也就会设成0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    rsExse.MoveFirst
    '打开另外一个连接串，以达到不受当前连接事务的控制
    cn上传.ConnectionString = gcnOracle.ConnectionString
    cn上传.Open
    
    '此处密码确定是得不到的，所以要强制刷卡
    Screen.MousePointer = vbDefault
    
    '取该病人的基本信息
    gstrSQL = "select A.人员身份,B.编码 from 保险帐户 A,保险病种 B where A.病人ID=" & lng病人ID & " and A.险类=" & gintInsure & "  and A.病种ID=B.ID(+)"
    Call OpenRecordset(rsTemp, "住院预算")
    If rsTemp.EOF = False Then
        str人员类别 = Nvl(rsTemp("人员身份"), "")
        '转换人员身份
        str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21" _
                      , str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", True, "11")
    End If
    
    mstr卡号 = ""
    mstr密码 = ""
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    Screen.MousePointer = vbHourglass
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)           ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)         ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str人员类别)         ' 人员类别
    Call InsertChild(mdomInput.documentElement, "ISCAL", 0)         ' 是否结算
    Call InsertChild(mdomInput.documentElement, "ACCTWANTTOPAY", "0")     ' 账户支付额
    Call InsertChild(mdomInput.documentElement, "INVOICENO", " ") ' 发票号
'    'Modified By 朱玉宝 2003-12-03 地区： 原因：增加上传病种编码
'    '补足8位长度
'    If str疾病编码 <> "" Then
'        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
'        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' 转大额标志
'        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)   ' 特种病编码
'    End If
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    Set nodRowset = InsertChild(mdomInput.documentElement, "ROWSET", "") ' 消费明细
    
    rsExse.Sort = "NO,序号,登记时间 asc"
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" Then
            gstrSQL = "SELECT C.药品ID,C.规格,E.名称 AS 剂型  FROM 药品目录 C,药品信息 D,药品剂型 E WHERE C.药品ID=" & rsExse("收费细目ID") & " AND C.药名ID=D.药名ID AND D.剂型=E.编码"
            gstrSQL = "select A.类别,A.名称,B.项目编码,nvl(A.规格,F.规格) AS 规格,F.剂型,A.计算单位 from 收费细目 A,保险支付项目 B,(" & gstrSQL & _
                    ") F where A.ID=" & rsExse("收费细目ID") & " and A.ID=B.收费细目ID  AND A.Id=F.药品ID(+) and B.险类=" & gintInsure
            Call OpenRecordset(rsTemp, "住院预算")
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '只上传只传递过的数据
            str医生 = ToVarchar(IIf(IsNull(rsExse("医生")), UserInfo.姓名, rsExse("医生")), 20)
            str科室 = ToVarchar(IIf(IsNull(rsExse("开单部门")), UserInfo.部门, rsExse("开单部门")), 56)
            
            Set nodRow = InsertChild(nodRowset, "ROW", "")
            Call nodRow.setAttribute("ITEMSERIAL", ToVarchar(rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") & "_" & rsExse("记录状态"), 20)) '数据批号，用于唯一代表数据
            Call nodRow.setAttribute("ITEMCODE", ToVarchar(rsExse("医保项目编码"), 12))
            Call nodRow.setAttribute("ITEMNAME", ToVarchar(rsExse("收费名称"), 72))
            Call nodRow.setAttribute("SUBJECT", Subject(rsTemp("类别")))
            Call nodRow.setAttribute("SPECIFICATION", ToVarchar(rsTemp("规格"), 40))
            Call nodRow.setAttribute("AGENTTYPE", ToVarchar(rsTemp("剂型"), 20))
            Call nodRow.setAttribute("UNIT", ToVarchar(rsTemp("计算单位"), 20))
            Call nodRow.setAttribute("PRICE", Format(rsExse("价格"), "0.0000"))
            Call nodRow.setAttribute("QUANTITY", Format(rsExse("数量"), "0.00"))
            Call nodRow.setAttribute("FROMOFFICE", str科室)   '开单科室
            Call nodRow.setAttribute("FROMDOCT", str医生)     '开单医生
            Call nodRow.setAttribute("TOOFFICE", str科室)    '受单科室
            Call nodRow.setAttribute("TODOCT", str医生)      '受单医生
            '处理时间时，为了保证同一保险项目的的收费时间不同，因此在登记时间上按序号加上秒数
            Call nodRow.setAttribute("DODATE", Format(rsExse("登记时间"), "yyyy-MM-dd HH:mm:ss"))      '办理日期
            Call nodRow.setAttribute("NOTE", ToVarchar(rsExse("摘要"), 512))     '备注
        End If
        cur发生费用 = cur发生费用 + rsExse("金额")
        rsExse.MoveNext
    Loop
    
    '调用接口
    If CommServer("CALHOSP") = False Then Exit Function
    '首先强调不能少传，所以等医保服务器正确返回后再打上标记
    If rsExse.RecordCount > 0 Then rsExse.MoveFirst
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" Then
            '为该条费用记录打上上传标志。上传一条处理一条
            gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
            cn上传.Execute gstrSQL, , adCmdStoredProc
        End If
        rsExse.MoveNext
    Loop
    
    cur个人帐户 = Val(GetElemnetValue("ACCTPAY"))
    If str人员类别 = "32" Or str人员类别 = "34" Then
        cur统筹支付 = Val(GetElemnetValue("ALLOWFUND"))
    Else
        cur统筹支付 = Val(GetElemnetValue("FUND1PAY"))
    End If
    cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
    
    '保存病人个人帐户余额
    mstr医保号 = str医保号
    mdbl余额 = cur个人帐户
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    
    住院虚拟结算_贵阳 = "医保基金;" & cur统筹支付 & ";0"
    If cur个人帐户 <> 0 Then
        住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|个人帐户;" & cur个人帐户 & ";1" '允许修改个人帐户
    End If
'    If cur大病统筹 <> 0 Then
        '这样做的目的是避免前端程序修改该结算方式的金额
        住院虚拟结算_贵阳 = 住院虚拟结算_贵阳 & "|大病统筹;" & cur大病统筹 & ";0"
'    End If
    
    mlng病人ID = lng病人ID  '表示该病人已经进行了虚拟结算
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_贵阳(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset
    
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, cur个人帐户 As Double, cur起付线 As Currency
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, datCurr As Date, strNO As String
    Dim str就诊顺序号 As String, str结算编号 As String
    Dim lng病种ID As Long
    Dim str疾病编码 As String
    Dim rs病种 As ADODB.Recordset
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    'Modified By 朱玉宝 2003-12-03 地区： 原因：入院登记身份验证后取消病种的选择，改在结算时必需确定病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & gintInsure & " Order by A.编码"
    
    Set rs病种 = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If Not rs病种 Is Nothing Then
        lng病种ID = rs病种("ID")
        str疾病编码 = rs病种("编码")
    Else
        lng病种ID = 0
        str疾病编码 = ""
    End If
    
    '更新病种信息
    gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & TYPE_贵阳市 & ",'病种ID','" & lng病种ID & "')"
    Call ExecuteProcedure("更新病种信息")
    
    'Modified By 朱玉宝 2003-12-03 地区： 原因：增加上传病种编码
    '补足8位长度
    If str疾病编码 <> "" Then
        str疾病编码 = String(8 - Len(str疾病编码), "0") & str疾病编码
        Call InsertChild(mdomInput.documentElement, "FUND2FLAG", "1")                 ' 转大额标志
        Call InsertChild(mdomInput.documentElement, "SPECILLNESSCODE", str疾病编码)   ' 特种病编码
    End If
    
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "住院结算")
    If Not rsTemp.EOF Then cur个人帐户 = rsTemp("金额")
    '求单据号
    gstrSQL = "Select NO,收费时间 From 病人结帐记录 Where ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "住院结算")
    
    'XML文档已经完成初始化，此时只需要更新部分值
    Call SetElemnetValue("ISCAL", "1")
    Call SetElemnetValue("ACCTWANTTOPAY", Format(cur个人帐户, "0.00"))
    Call SetElemnetValue("INVOICENO", rsTemp("NO"))
    Call SetElemnetValue("DODATE", Format(rsTemp("收费时间"), "yyyy-MM-dd HH:mm:ss"))
    '预算时已经传递，结帐不需要再传递明细数据
    Call SetElemnetValue("ROWSET", "")
    '调用接口
    If CommServer("CALHOSP") = False Then Exit Function
    
    cur全自付 = Val(GetElemnetValue("FEEOUT"))
    cur挂钩自付 = Val(GetElemnetValue("FEESELF"))
    cur起付线 = Val(GetElemnetValue("STARTFEE"))
    cur基数自付 = Val(GetElemnetValue("ENTERSTARTFEE"))
    cur统筹支付 = Val(GetElemnetValue("FUND1PAY")) + Val(GetElemnetValue("ALLOWFUND"))
    cur统筹自付 = Val(GetElemnetValue("FUND1SELF"))
    cur大病统筹 = Val(GetElemnetValue("FUND2PAY"))
    cur大病自付 = Val(GetElemnetValue("FUND2SELF"))
    cur超限自付 = Val(GetElemnetValue("FEEOVER"))
    
    str结算编号 = GetElemnetValue("BALANCEID")
    str就诊顺序号 = GetElemnetValue("BILLNO")
    
    '填写结算表
    datCurr = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 + cur统筹支付 + cur大病统筹 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("贵阳医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",NULL," & cur基数自付 & "," & _
        g结算数据.发生费用金额 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & str结算编号 & "',null,null,'" & str就诊顺序号 & "')"
    Call ExecuteProcedure("贵阳医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & ",NULL)"
    Call ExecuteProcedure("贵阳医保")
    
    '如果清算方式不是按日清单且人员类别不是离休人员，则提示操作员为该病人办理出院手续
    gstrSQL = "Select 单病种,人员身份 From 保险帐户 Where 险类=" & TYPE_贵阳市 & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "提取清算方式")
    If Right(rsTemp!单病种, 1) <> 4 And Not (rsTemp!人员身份 = "市属离休" Or rsTemp!人员身份 = "省属离休") Then
        MsgBox "请为该参保人员办理出院手续！", vbInformation, gstrSysName
    End If
    
    住院结算_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_贵阳(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    Dim lng冲销ID As Long, lng病人ID As Long
    Dim str结帐日期 As String, str当前日期 As String
    Dim rsTemp  As New ADODB.Recordset, rsCheck As New ADODB.Recordset
    
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str人员类别 As String
    Dim str就诊顺序号 As String, str结算编号 As String
    Dim cur个人帐户 As Currency
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim curDate As Date    '退费
    
    On Error GoTo ErrHand
    curDate = zlDatabase.Currentdate
    
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
        " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    lng冲销ID = rsTemp("ID") '冲销单据的ID
    
    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "结算冲销")
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    lng病人ID = rsTemp!病人ID
    str结算编号 = IIf(IsNull(rsTemp!支付顺序号), "", rsTemp!支付顺序号)
    str就诊顺序号 = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
    
    '判断是否为离体人员
    gstrSQL = "Select 人员身份 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & gintInsure
    Call OpenRecordset(rsCheck, "判断是否为离休人员")
    If Not (rsCheck!人员身份 = "省属离休" Or rsCheck!人员身份 = "市属离休") Then
        MsgBox "基本医疗产生的结帐记录不允许冲销！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '非本月结帐的单据，不允许冲销
    gstrSQL = "select to_char(收费时间,'yyyy-MM-dd') 结帐时间 From 病人结帐记录 Where ID=" & lng结帐ID
    Call OpenRecordset(rsCheck, "取结帐日期")
    str结帐日期 = Format(rsCheck!结帐时间, "yyyyMM")
    str当前日期 = Format(zlDatabase.Currentdate, "yyyyMM")
    If str当前日期 <> str结帐日期 Then
        MsgBox "只能冲销本月的结帐单据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '----准备冲销结帐----
    '读取医保病人的基本信息
    gstrSQL = "Select 卡号,医保号,顺序号 中心,人员身份,密码 From 保险帐户 Where 险类=" & gintInsure & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsCheck, "获取医保病人的基本信息")
    str卡号 = rsCheck!卡号
    str医保号 = rsCheck!医保号
    str分中心编号 = rsCheck!中心
    str人员类别 = rsCheck!人员身份
    str人员类别 = Switch(str人员类别 = "在职", "11", str人员类别 = "退休", "21" _
                  , str人员类别 = "省属离休", "32", str人员类别 = "市属离休", "34", True, "11")
    str密码 = IIf(IsNull(rsCheck!密码), "", rsCheck!密码)
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "CARDID", str卡号)                  ' 磁卡编码
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)            ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号)        ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "PASSWORD", str密码)                ' 密码
    Call InsertChild(mdomInput.documentElement, "PERSONTYPE", str人员类别)          ' 人员类别
    Call InsertChild(mdomInput.documentElement, "BILLNO", str就诊顺序号)            ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str结算编号)           ' 结算编号
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)           ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    
    '调用接口
    If CommServer("RETHOSP") = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("贵阳医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & gintInsure & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str结算编号 & "',null,null,'" & str就诊顺序号 & "')"
    Call ExecuteProcedure("贵阳医保")
    
    住院结算冲销_贵阳 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub 查询欠费单位_贵阳(ByVal str单位编码 As String)
'功能：调用接口查询欠费单位
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str提示 As String
    
    If str单位编码 = "" Then Exit Sub
'    str单位编码 = String(12 - Len(str单位编码), "0") & str单位编码
    
    On Error GoTo errHandle
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101")   ' 分中心编码(贵阳医保)
    Call InsertChild(mdomInput.documentElement, "DEPTCODE", str单位编码)         ' 单位编码
    
    '调用接口
    If CommServer("QUERYARREARDEPT") = False Then Exit Sub
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then
        MsgBox "病人单位无欠费情况。", vbInformation, gstrSysName
        Exit Sub
    End If
    '根据编码得到险种名称
    For Each nodRow In nodRowset.childNodes
        Select Case GetAttributeValue(nodRow, "INSUREKIND")
            Case "3"
                str提示 = str提示 & "、基本医疗"
            Case "8"
                str提示 = str提示 & "、大额医疗"
        End Select
    Next
    
    If str提示 <> "" Then
        MsgBox "病人单位以下险种有欠费情况：" & Mid(str提示, 2) & "。", vbInformation, gstrSysName
    Else
        MsgBox "病人单位无欠费情况。", vbInformation, gstrSysName
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function 错误信息_贵阳(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 医保项目_贵阳(rsTemp As ADODB.Recordset) As Boolean
'功能：医保诊疗药品目录查询
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim str编码 As String, str名称 As String, str简码, str失效 As String
        
    On Error GoTo errHandle
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "ITEMCODE", "")         ' 医保编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", "0101") ' 分中心编码(贵阳医保)
    
    '调用接口
    If CommServer("QUERYSERVICE") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    For Each nodRow In nodRowset.childNodes
        str编码 = GetAttributeValue(nodRow, "ITEMCODE")
        str名称 = ToVarchar(Replace(GetAttributeValue(nodRow, "ITEMNAME"), "'", ""), 40)
        str简码 = ToVarchar(zlCommFun.SpellCode(str名称), 10)
        str失效 = GetAttributeValue(nodRow, "ISVALID")
        If str编码 <> "" And str失效 <> "1" Then
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY"), Array("1", str编码, str名称, str简码)
            rsTemp.Update
        End If
    Next
    
    
    医保项目_贵阳 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function InitXML() As Boolean
'功能：初始化XML，增加声明和根节点
    Dim pi As MSXML2.IXMLDOMProcessingInstruction
    Dim nodData As MSXML2.IXMLDOMElement
    
    On Error Resume Next
    
    Set mdomInput = New MSXML2.DOMDocument
    Set mdomOutput = New MSXML2.DOMDocument
    If Err <> 0 Then
        Err.Clear
        Exit Function
    End If
    
'    'XML声明
'    Set pi = mdomInput.createProcessingInstruction("xml", "version=""1.0"" encoding=""GB2312"" standalone=""yes""")
'    mdomInput.appendChild pi
    
    '根节点
    Set nodData = mdomInput.createElement("DATA")
    Set mdomInput.documentElement = nodData
    
    InitXML = True
End Function

Public Function InsertChild(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String) As MSXML2.IXMLDOMElement
'功能：在指定XML元素下增加子元素
    Set InsertChild = mdomInput.createElement(Name)
    InsertChild.Text = Value
    
    nodParent.appendChild InsertChild
End Function

Public Sub InsertAttrib(nodParent As MSXML2.IXMLDOMElement, ByVal Name As String, ByVal Value As String)
'功能：在指定XML元素下增加属性
    Dim attTemp As MSXML2.IXMLDOMAttribute
    
    Set attTemp = mdomInput.createAttribute(Name)
    attTemp.Text = Value
    
    nodParent.setAttributeNode attTemp
End Sub
'
'Private Function CommServer(ByVal strFunction As String) As Boolean
''功能：与医保服务器进行通讯，得到返回值
'    Dim cnComm As New ADODB.Connection
'    Dim rsTemp As New ADODB.Recordset
'    Dim lngID As Long
'
'    Dim lng行数 As Long, lng序号 As Long, strTemp As String, strInput As String
'    Dim timStart As Date, bln已处理 As Boolean
'
'
'    '为了实现事务控制，需要另建一个连接
'    '使用事务的目的是为了保证读出来的传入参数和传出参数都是完整的
'    cnComm.ConnectionString = gcnOracle.ConnectionString
'    cnComm.Open
'
'    '参数的传入
'    strInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
'    lngID = zlDatabase.GetNextId("保险接口表")
'    lng行数 = Abs(Int(Len(strInput) / -2000))  '有可能传入参数大长，要分成多行才能保存
'    On Error Resume Next
'    cnComm.BeginTrans
'    For lng序号 = 1 To lng行数
'        '分成若干行
'        strTemp = Replace(Mid(strInput, (lng序号 - 1) * 2000 + 1, 2000), "'", "''")
'        gstrSQL = "insert into 保险接口表(ID,序号,行数,类型,传入参数,传出参数,状态) values (" & _
'            lngID & "," & lng序号 & "," & lng行数 & ",'" & strFunction & "','" & strTemp & "','未处理',0)"
'        cnComm.Execute gstrSQL
'    Next
'    If Err <> 0 Then
'        '出错
'        Err.Clear
'        cnComm.RollbackTrans
'        Exit Function
'    End If
'    cnComm.CommitTrans
'
'    On Error GoTo errHandle
'    '等待答复
'    timStart = Now
'    Do While True '为了保证医保服务器的处理结果一定能得到接收，因此不使用超时退出     DateDiff("s", timStart, Now) < 600 '小于600秒钟
'        DoEvents
'        If rsTemp.State = adStateOpen Then rsTemp.Close
'        gstrSQL = "select 传出参数 from 保险接口表 where ID=" & lngID & " and 传出参数<>'未处理' order by 序号"
'        rsTemp.Open gstrSQL, cnComm, adOpenStatic, adLockReadOnly
'
'        If rsTemp.EOF = False Then
'            '取得返回值了
'            strTemp = ""
'            Do Until rsTemp.EOF
'                strTemp = strTemp & IIf(IsNull(rsTemp("传出参数")), "", rsTemp("传出参数"))
'                rsTemp.MoveNext
'            Loop
'
'            If mdomOutput.loadXML(strTemp) = False Then
'                MsgBox "医保服务器返回值格式不正确。", vbInformation, gstrSysName
'            Else
'                '再对整个调用是否成功进行分析
'                If Val(GetElemnetValue("RETCODE")) = 0 Then
'                    '调用成功
'                    CommServer = True
'                Else
'                    '调用失败
'                    strTemp = GetElemnetValue("INFO")
'                    If strTemp = "" Then strTemp = "服务器调用失败。"
'                    MsgBox "医保服务器返回错误：" & vbCrLf & vbCrLf & strTemp, vbInformation, gstrSysName
'                End If
'            End If
'            bln已处理 = True
'            Exit Do
'        End If
'    Loop
'
'    If bln已处理 = False Then
'        MsgBox "与医保服务器连接超时。", vbInformation, gstrSysName
'    End If
'errHandle:
'    cnComm.Execute "Delete from 保险接口表 where id=" & lngID '不论成功与否，都把数据删除
'End Function

Public Function CommServer(ByVal strFunction As String) As Boolean
'功能：调用医保部件
    Dim obj医保 As Object
    Dim InvokeServer As String '调用前置服务器的返回值
    Dim strInput As String, strServer As String
    
    On Error Resume Next
    '如果用全局变量，有时调用时会等很久，可能资源分配的原故
    strServer = Get保险参数_贵阳("医保服务器")
    If strServer = "" Then
        Set obj医保 = CreateObject("HospCOMSvr.HospCOMServer")
    Else
        Set obj医保 = CreateObject("HospCOMSvr.HospCOMServer", strServer)
    End If
    If Err <> 0 Then
        MsgBox "无法创建医保接口部件（HospCOMSvr.HospCOMServer）。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '参数的传入
    strInput = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?>" & vbCrLf & mdomInput.xml
    
    Select Case strFunction
        Case "READCARD"         '身份识别/信息获取
            InvokeServer = obj医保.ReadCard("ZFRJ", strInput)
        Case "READCARD_M"       '身份识别/信息获取（手工方式）
            InvokeServer = obj医保.ReadCard_M("ZFRJ", strInput)
        Case "MODIFYCARD"
            InvokeServer = obj医保.MODIFYCARD("ZFRJ", strInput)
        'Modified By 朱玉宝 2004-05-25 原因：医保接口变动
        '------------------------------------------------
        Case "GETCLINNO"        '门诊挂号
            InvokeServer = obj医保.GETCLINNO("ZFRJ", strInput)
        '------------------------------------------------
        Case "CALCLIN"          '普通门诊支付
            InvokeServer = obj医保.CALCLIN("ZFRJ", strInput)
        Case "CALSPECCLIN"      '特殊门诊支付
            InvokeServer = obj医保.CALSPECCLIN("ZFRJ", strInput)
        Case "RETCLIN"          '收费冲销
            InvokeServer = obj医保.RETCLIN("ZFRJ", strInput)
        Case "HOSPREG"          '住院登记
            InvokeServer = obj医保.HOSPREG("ZFRJ", strInput)
        Case "HOSPOUT"          '出院登记
            InvokeServer = obj医保.HOSPOUT("ZFRJ", strInput)
        Case "CALHOSP"          '住院支付
            InvokeServer = obj医保.CALHOSP("ZFRJ", strInput)
        Case "RETHOSP"          '结帐冲销
            InvokeServer = obj医保.RETHOSP("ZFRJ", strInput)
        Case "SETRECKONINGTYPE"
            InvokeServer = obj医保.SETRECKONINGTYPE("ZFRJ", strInput)
        Case "QUERYHOSPSINGLEILLNESS"   '单病种清算数据
            InvokeServer = obj医保.QUERYHOSPSINGLEILLNESS("ZFRJ", strInput)
        Case "QUERYSERVICE"     '医保诊疗药品目录查询
            InvokeServer = obj医保.QUERYSERVICE("ZFRJ", strInput)
        Case "QUERYARREARDEPT"
            InvokeServer = obj医保.QUERYARREARDEPT("ZFRJ", strInput)
        Case "GETHOSPSINGLEILLNESS"
            InvokeServer = obj医保.GETHOSPSINGLEILLNESS("ZFRJ", strInput)
        Case Else
            MsgBox "可能医保接口发生变化，无法继续执行交易，请与软件提供商联系！", vbInformation, gstrSysName
            Exit Function
    End Select
    
    '断点设置处
    If InvokeServer = "" Then
        '调用失败，返回固定的错误信息
        InvokeServer = "<?xml version=""1.0"" encoding=""GB2312"" standalone=""yes""?><DATA><RETCODE>-1</RETCODE><INFO>医保服务器调用失败</INFO></DATA>"
    End If
            
    If mdomOutput.loadXML(InvokeServer) = False Then
        MsgBox "医保服务器返回值格式不正确。", vbInformation, gstrSysName
    Else
        '再对整个调用是否成功进行分析
        If Val(GetElemnetValue("RETCODE")) = 0 Then
            '调用成功
            CommServer = True
        Else
            '调用失败
            InvokeServer = GetElemnetValue("INFO")
            If InvokeServer = "" Then InvokeServer = "服务器调用失败。"
            MsgBox "医保服务器返回错误：" & vbCrLf & vbCrLf & InvokeServer, vbInformation, gstrSysName
        End If
    End If
End Function

Private Function Get保险参数_贵阳(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名='" & str参数名 & "' and A.险类=" & TYPE_贵阳市 & " and A.中心 is null "
    Call OpenRecordset(rsTemp, "贵阳医保")
    
    If rsTemp.EOF = False Then
        Get保险参数_贵阳 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function SetElemnetValue(ByVal Name As String, ByVal Value As String) As Boolean
'功能：得到指定元素的值
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomInput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        xmlElement.nodeTypedValue = Value
        SetElemnetValue = True
    End If
End Function

Public Function GetElemnetValue(ByVal Name As String) As String
'功能：得到指定元素的值
    Dim xmlElement As MSXML2.IXMLDOMElement
    
    Set xmlElement = mdomOutput.documentElement.selectSingleNode(Name)
    If Not xmlElement Is Nothing Then
        '找到指定子元素
        GetElemnetValue = xmlElement.Text
'    Else
'        '取消
'        Debug.Assert False
    End If
End Function

Public Function GetAttributeValue(xmlElement As MSXML2.IXMLDOMElement, ByVal Name As String) As String
'功能：得到指定属性的值
    Dim varAttribute As Variant
    
    varAttribute = xmlElement.getAttribute(Name)
    If IsNull(varAttribute) = False Then
        GetAttributeValue = varAttribute
    End If
End Function

Public Function Get验证_贵阳(str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, _
                ByVal lng病人ID As Long, Optional bln强制刷卡 As Boolean = False) As Boolean
'功能：得到医保病人的基本功的身份验证信息
    Dim rsTemp As New ADODB.Recordset
    Dim strtemp As String
    
    If bln强制刷卡 = False And lng病人ID > 0 Then
        '从数据库中读出已存储的值
        gstrSQL = "select 卡号,医保号,顺序号,密码 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
        Call OpenRecordset(rsTemp, "贵阳医保")
        
        If rsTemp.EOF = False Then
            strtemp = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
            If strtemp = mstr卡号 And mstr卡号 <> "" Then
                '是同一病人
                str卡号 = mstr卡号
                str密码 = mstr密码
            Else
                str卡号 = strtemp
                str密码 = IIf(IsNull(rsTemp("密码")), "", rsTemp("密码"))
            End If
            
            str医保号 = IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号"))
            str分中心编号 = IIf(IsNull(rsTemp("顺序号")), "", rsTemp("顺序号"))
            
            Get验证_贵阳 = True
            Exit Function
        End If
    End If
    
    If frmIdentify贵阳.GetIdentify(TYPE_贵阳市, str卡号, str医保号, str分中心编号, str密码, True, True) = False Then
        Exit Function
    Else
        '刷卡虽然正确，但要检查是否就是当前病人的
            str卡号 = Split(str卡号, "^")(0)
            If lng病人ID > 0 Then
            '从数据库中读出已存储的值
            gstrSQL = "select 卡号,医保号,顺序号 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
            Call OpenRecordset(rsTemp, "贵阳医保")
            
            If str卡号 <> IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号")) Or str医保号 <> IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号")) Then
                MsgBox "当前使用的卡与病人不符。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
    End If
    
    Get验证_贵阳 = True
End Function

Public Function OwnerUser(ByVal strUserName As String) As Boolean
    Dim RecUser As New ADODB.Recordset
    '判断当前用户是不是所有者
    OwnerUser = True
    With RecUser
        If .State = 1 Then .Close
        .Open "Select Count(*) 所有者 From ZlSystems Where 所有者='" & strUserName & "'", gcnOracle
        
        If Not .EOF Then
            If Not IsNull(!所有者) Then
                If !所有者 = 0 Then OwnerUser = False
            End If
        End If
    End With
End Function

Public Function Subject(ByVal strData As String) As String
    Dim rsSubject As New ADODB.Recordset
    '返回对应的归属科目编码
    gstrSQL = "" & _
             " Select B.编码,B.类别,A.参数值 归属科目编码   " & _
             " From 保险参数 A,收费类别 B " & _
             " Where A.序号>=6 And A.险类=" & gintInsure & " And A.参数名=B.编码 And B.编码='" & strData & "'"
    Call OpenRecordset(rsSubject, "获取对应的归属科目编码")
    
    If rsSubject.EOF Then
        Subject = "11"  '无对应项目返回对应的归属科目编码'11',表示其他
    Else
        Subject = rsSubject!归属科目编码
    End If
End Function

Public Function 门诊挂号_贵阳(ByVal lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim datCurr As Date
    Dim str结算方式 As String, arr结算方式
    Dim intTotal  As Integer, intStart As Integer
    Dim cur帐户余额 As Double, cur个人帐户 As Currency
    Dim cur医保基金 As Currency, cur大额统筹 As Currency
    Dim str卡号 As String, str医保号 As String, str分中心编号 As String, str密码 As String, str就诊顺序号 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
        " From 病人预交记录 A,保险帐户 B " & _
        " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
        "       And A.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "取病人ID")
    If rsTemp.EOF Then Exit Function
    lng病人ID = rsTemp!病人ID
    If Get验证_贵阳(str卡号, str医保号, str分中心编号, str密码, lng病人ID) = False Then Exit Function
    
    datCurr = zlDatabase.Currentdate()
    
    '取帐户余额
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额 From 保险帐户 Where 险类=" & gintInsure & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "取帐户余额")
    cur帐户余额 = rsTemp!帐户余额
    
    '对XML对象赋值
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", str医保号)     ' 个人编码
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", str分中心编号) ' 分中心编码
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(datCurr, "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    
    '调用接口
    If CommServer("GETCLINNO") = False Then Exit Function
    str就诊顺序号 = GetElemnetValue("BILLNO")
    
    gstrSQL = "Select 病人ID,收费细目ID,数次*NVL(付数,1) AS 数量,标准单价 AS 单价,'  ' AS 摘要" & _
        " From 病人费用记录 " & _
        " Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call OpenRecordset(rsTemp, "重庆医保")
    If Not 门诊虚拟结算_贵阳(rsTemp, str结算方式) Then Exit Function
    
    '分解各种结算方式
    arr结算方式 = Split(str结算方式, "|")
    intTotal = UBound(arr结算方式)
    For intStart = 0 To intTotal
        Select Case Split(arr结算方式(intStart), ";")(0)
        Case "个人帐户"
            cur个人帐户 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "医保基金"
            cur医保基金 = Val(Split(arr结算方式(intStart), ";")(1))
        Case "大额统筹"
            cur大额统筹 = Val(Split(arr结算方式(intStart), ";")(1))
        End Select
    Next
    
    If Not 门诊结算_贵阳(lng结帐ID, cur个人帐户, "") Then Exit Function
    
   '需要修正结算结果
    str结算方式 = ""
    If cur个人帐户 <> 0 Then str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    If cur医保基金 <> 0 Then str结算方式 = str结算方式 & "||医保基金|" & cur医保基金
    If cur大额统筹 <> 0 Then str结算方式 = str结算方式 & "||大额统筹|" & cur大额统筹
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        Call ExecuteProcedure("更新预交记录")
    End If
    
    门诊挂号_贵阳 = True
    
    Call frm结算信息.ShowMe(lng结帐ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 设置清算方式_贵阳(ByVal lng病人ID As Long, ByVal frmParent As Object) As Boolean
    设置清算方式_贵阳 = frm设置清算方式.ShowSelect(lng病人ID, TYPE_贵阳市, frmParent)
End Function

'调试用
'txtEdit(0).Text = "GY0001"
'txtEdit(1).Text = "01"
'txtEdit(2).Text = "01"
'str类别 = "32"
'str姓名 = "贵阳"
'str性别 = "男"
'str身份证号码 = "510224770909071"
'str人员类别 = "省属离休"
'cur帐户余额 = 500
'str就诊顺序号 = "00000001"
'str结算编号 = "JS000001"
'str结算方式 = "个人帐户;1;0|医保基金;2;0"
'cur统筹支付 = 2

