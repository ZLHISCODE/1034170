Attribute VB_Name = "mdl重庆"
Option Explicit
'API函数声明

'1、接口初始化：检查整个网络环境是否畅通，包括医院客户端与前置机、前置机与中心服务器间。
Private Declare Function dy_Init Lib "SiInterface" Alias "INIT" () As Long

'2 业务处理：调用执行医保业务所需要的处理
Private Declare Function dy_Business_Handle Lib "SiInterface" Alias "BUSINESS_HANDLE" _
    (ByVal InputData As String, ByVal OutputData As String) As Long
    
Private mstr医保号 As String
Private mdbl余额 As Double
Private mlng病人ID As Long
Private mstr门诊号 As String
Private mblnIint As Boolean

Public Function 医保初始化_重庆() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim lngReturn As Long
    
    If mblnIint = True Then
        '只需要调用一次
        医保初始化_重庆 = True
        Exit Function
    End If
    
    On Error Resume Next
    
    lngReturn = dy_Init
    If Err <> 0 Then
        MsgBox "不能正确调用医保接口程序。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngReturn = -1 Then
        MsgBox "不能完成医保接口初始化工作，请检查整个网络环境是否畅通。包括：" & vbCrLf & vbCrLf & _
          "1、医院客户端与医院前置机应用服务器之间；" & vbCrLf & _
          "2、医院前置机应用服务器与医保中心应用服务器之间。", vbInformation, gstrSysName
    Else
        医保初始化_重庆 = True
        mblnIint = True
    End If
End Function

Public Function 身份标识_重庆(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str医保号 As String, strInput As String, arrOutput  As Variant, int类别 As Integer
    Dim str姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str人员类别 As String, str单位编码 As String, str单位名称 As String
    Dim strIdentify As String, str附加 As String, str中心编号 As String, str门诊号 As String
    Dim datCurr As Date
    
    '初始化一些变量
    mlng病人ID = 0
    mstr门诊号 = ""
    mstr医保号 = ""
    mdbl余额 = 0
    
    int类别 = bytType
    If frmIdentify重庆.GetIdentify(str医保号, int类别) = False Then
        Exit Function
    End If
    
    '调用接口
    
    strInput = "01|" & str医保号
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '取得返回值
    str姓名 = arrOutput(1)
    str性别 = arrOutput(2)
    lng年龄 = Val(arrOutput(3))
    str身份证号码 = arrOutput(4)
    str出生日期 = Get出生日期(str身份证号码, lng年龄)
    
    str人员类别 = ToVarchar(arrOutput(7), 8) 'VARCHAR2 (20) 在职，在职驻外，临时用工，自由职业军转干，退休，退休异地居住，退职，退职异地居住等
    'arrOutput(8)   公务员标志               'VARCHAR2 (3)
    str单位编码 = ""
    str单位名称 = ToVarchar(arrOutput(9), 48) '50的长度，还要扣除2个符号
    str中心编号 = arrOutput(10)
    
    If arrOutput(11) = "2" Then
        MsgBox "该病人医保卡不能继续使用。" & arrOutput(12)
        Exit Function
    End If
    
    If arrOutput(11) = "1" And bytType = 1 Then
        '住院时要提醒
        MsgBox "该医保病人统筹金额不能使用。" & arrOutput(12)
    End If
    
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    strIdentify = str医保号 & ";" & str医保号 & ";;" & str姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";" & str单位名称 & "(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";"                             '9.顺序号
    str附加 = str附加 & ";" & str人员类别               '10人员身份
    str附加 = str附加 & ";0"                            '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";"                             '13病种ID
    str附加 = str附加 & ";" & IIf(Left(str人员类别, 1) = "退", 2, 1)     '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";0"                            '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";" & IIf(int类别 = 14, 1, "")  '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID)
    
    If bytType = 0 Then        '如果是门诊，同时进行就诊登记
        '如果是特殊病或急诊抢救，需要选择病人疾病
        Dim rs病种 As ADODB.Recordset
        Dim str分类 As String, str疾病编码 As String, str并发症 As String
        
        If int类别 = 13 Or int类别 = 14 Then
            If int类别 = 13 Then
                '获得审批信息
                strInput = "07|" & str医保号
                If HandleBusiness(strInput, arrOutput) = False Then Exit Function
                
                str分类 = "特殊病"
                If frm疾病选择重庆.GetCode(arrOutput, str分类, str疾病编码, str并发症) = False Then Exit Function
            Else
                str分类 = "急诊"
                If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str并发症) = False Then Exit Function
            End If
        End If
                
        datCurr = zlDataBase.Currentdate
        str门诊号 = ToVarchar(lng病人ID & Format(datCurr, "yyMMddHHmmss"), 18)
        strInput = "02|" & str门诊号 & "|" & int类别 & "|" & str医保号 & _
                   "|门诊|" & ToVarchar(UserInfo.姓名, 20) & "|" & _
                   Format(datCurr, "yyyy-MM-dd") & "|" & str疾病编码 & "|" & ToVarchar(UserInfo.姓名, 20) & "|" & str并发症
        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
        
        mlng病人ID = lng病人ID
        mstr门诊号 = str门诊号
        mstr医保号 = str医保号
        mdbl余额 = Val(arrOutput(2))
    End If
    g结算数据.超限自付金额 = int类别 '用于暂时保存，门诊类别
    
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_重庆 = strIdentify & ";" & lng病人ID & str附加
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_重庆(strSelfNo As String) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHandle
    
    '从数据库中读取（因为刚才才保存了的，应该是准确的）
    If mstr医保号 = "" Or strSelfNo <> mstr医保号 Then
        gstrSQL = "Select 帐户余额 From 保险帐户 where 险类=" & gintInsure & " and 中心=0 and 医保号='" & strSelfNo & "'"
        Call OpenRecordset(rsTemp, "重庆医保")
        
        If rsTemp.EOF = False Then
            个人余额_重庆 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
        End If
    Else
        个人余额_重庆 = mdbl余额
    End If
    '只能用一次
    mstr医保号 = ""
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_重庆(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Static str门诊号Pre As String
    Dim str医保号 As String, strInput As String, arrOutput  As Variant
    Dim dbl个人帐户 As Double, strMessage As String
    Dim lng病人ID As Long, str规格 As String, datCurr As Date
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_重庆 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    datCurr = zlDataBase.Currentdate
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '首先退掉以前发生的所有未结的费用，包括多次执行预结算
    If str门诊号Pre = mstr门诊号 Then
        '已经赋值，说明该病人进行过预算
        strInput = "10|" & mstr门诊号 & "|" & mstr门诊号
        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    End If
    '保存该值
    str门诊号Pre = mstr门诊号
    
    '然后插入处方明细
    Do Until rs明细.EOF
        gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,A.计算单位,E.规格,G.名称 剂型 " & _
                  "from 收费细目 A,保险支付项目 B,药品目录 E ,药品信息 F,药品剂型 G " & _
                  "where A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID and B.险类=" & gintInsure & _
                 "        AND A.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) "
        Call OpenRecordset(rsTemp, "门诊预算")
        If rsTemp.EOF = True Then
            MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
            Exit Function
        End If
        
        strInput = "04|" & mstr门诊号 & "|" & mstr门诊号 & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss")
        strInput = strInput & "|" & ToVarchar(rsTemp("项目编码"), 10)  '医保流水号
        strInput = strInput & "|" & ToVarchar(rsTemp("编码"), 20)      '医院内码
        strInput = strInput & "|" & ToVarchar(rsTemp("名称"), 50)      '项目名称
        strInput = strInput & "|" & Format(rs明细("单价"), "0.0000")   '单价
        strInput = strInput & "|" & Format(rs明细("数量"), "0.00")     '数量
        strInput = strInput & "|" & IIf(rs明细("是否急诊") = 1, 1, 0)  '急诊标志
        strInput = strInput & "|" & Format(UserInfo.姓名, 20)          '处方医生
        strInput = strInput & "|" & Format(UserInfo.姓名, 20)          '经办人
        strInput = strInput & "|" & ToVarchar(rsTemp("计算单位"), 20)     '单位
        strInput = strInput & "|" & ToVarchar(rsTemp("规格"), 14)         '规格
        strInput = strInput & "|" & ToVarchar(rsTemp("剂型"), 20)         '剂理
        strInput = strInput & "|"                                         '冲销明细流水号
        strInput = strInput & "|" & Format(rs明细("实收金额"), "#####0.0000")         '金额
        
        If HandleBusiness(strInput, arrOutput) = False Then Exit Function
        Call AddMessage(strMessage, arrOutput, ToVarchar(rsTemp("名称"), 50), rs明细("单价"))
        
        rs明细.MoveNext
    Loop
    
    If strMessage <> "" Then
        strMessage = "病人费用明细传输过程中得到医保中心如下反馈信息，是否继续？" & vbCrLf & vbCrLf & strMessage
        If MsgBox(strMessage, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            '用户选择取消，先退掉明细
            strInput = "10|" & mstr门诊号 & "|" & mstr门诊号
            If HandleBusiness(strInput, arrOutput) = False Then Exit Function
                        
            Exit Function
        End If
    End If
    '调用预结算
    
    strInput = "06|" & mstr门诊号
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    
    str结算方式 = "个人帐户;" & Val(arrOutput(2)) & ";0"  '不能修改个人帐户，因为结算时已经不再传金额到前置机了
    If Val(arrOutput(1)) > 0 Then
        str结算方式 = str结算方式 & "|医保基金;" & Val(arrOutput(1)) & ";0"
    End If
    If Val(arrOutput(3)) > 0 Then
        str结算方式 = str结算方式 & "|公务员补助;" & Val(arrOutput(3)) & ";0"
    End If
    If Val(arrOutput(5)) > 0 Then
        str结算方式 = str结算方式 & "|大额统筹;" & Val(arrOutput(5)) & ";0"
    End If
    If Val(arrOutput(6)) > 0 Then
        str结算方式 = str结算方式 & "|公务员返还;" & Val(arrOutput(6)) & ";0"
    End If
    
    门诊虚拟结算_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_重庆(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str医保号 As String, strInput As String, arrOutput  As Variant
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset
    Dim str操作员 As String, str病种 As String, cur发生费用, datCurr As Date
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9"
    Call OpenRecordset(rs明细, "重庆医保")
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    Do Until rs明细.EOF
        cur发生费用 = cur发生费用 + rs明细("结帐金额")
        rs明细.MoveNext
    Loop
    
    '调用结算
    strInput = "05|" & mstr门诊号 & "|1||" & str操作员 & "|0" '用帐户余额支付
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
            
    Dim cur统筹支付 As Double
    Dim cur公务员补助 As Double
    Dim cur大额统筹 As Double
    
    cur发生费用 = Val(Format(cur发生费用, "#####0.00"))
    cur统筹支付 = Val(arrOutput(2))
    cur公务员补助 = Val(arrOutput(4))
    cur大额统筹 = Val(arrOutput(6))
    
    '帐户年度信息
    datCurr = zlDataBase.Currentdate
    Call 读取病种(lng病人ID, str病种)
    
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 + cur统筹支付 & "," & _
        cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("重庆医保")
    
    'g结算数据.超限自付金额中保存的是门诊病人就诊类型（急诊、特殊病门诊或普通门诊），结算记录的备注保存的是病种的名称
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)'超限自付金额用于暂时保存，门诊类别
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur发生费用 & ",0,0," & _
        cur统筹支付 & "," & cur统筹支付 & ",0," & g结算数据.超限自付金额 & "," & cur个人帐户 & ",'" & arrOutput(1) & "',NULL,NULL," & IIf(str病种 = "", "NULL", "'" & str病种 & "'") & ")"
    Call ExecuteProcedure("重庆医保")
    '---------------------------------------------------------------------------------------------

    门诊结算_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_重庆(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDataBase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额  From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9"
    Call OpenRecordset(rsTemp, "重庆医保")
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    cur票据总金额 = Val(Format(cur票据总金额, "#####0.00"))
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "重庆医保")
    
    lng冲销ID = rsTemp("结帐ID")
    
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "重庆医保")
    
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    str流水号 = rsTemp("支付顺序号")
    
    strInput = "99|" & str流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("重庆医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str流水号 & "')"
    Call ExecuteProcedure("重庆医保")

    门诊结算冲销_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人帐户转预交_重庆(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    
    个人帐户转预交_重庆 = False
End Function

Public Function 入院登记_重庆(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strInput As String, arrOutput  As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim str卡号 As String, str顺序号 As String
    Dim strTemp As String, str提示 As String, str诊断 As String
    
    On Error GoTo errHandle
    
    
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型=1 and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "入院登记")
    If rsTemp.EOF = False Then
        str诊断 = ToVarchar(rsTemp("描述信息"), 40)
    End If
    
    '获得医保号
    gstrSQL = "select 卡号,医保号 from 保险帐户 where 险类=" & TYPE_重庆市 & " and 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "入院登记")
    str卡号 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
    str医保号 = rsTemp("医保号")
    
    '获得其它入院信息
    datCurr = zlDataBase.Currentdate
    gstrSQL = "select A.入院方式,nvl(A.二级院转入,0) as 二级院转入,A.门诊医师,A.入院日期,A.入院病床,B.名称 as 入院科室 from 病案主页 A,部门表 B " & _
             " Where A.入院科室ID = B.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "入院登记")
    
    '调用入院接口
    strInput = "02|" & lng病人ID & "_" & lng主页ID & "|" & IIf(rsTemp("入院方式") = "转入", "22", "21") & "|" & str医保号 & "|" & _
               ToVarchar(rsTemp("入院科室"), 30) & "|" & ToVarchar(rsTemp("门诊医师"), 20) & "|" & _
               Format(rsTemp("入院日期"), "yyyy-MM-dd") & "|" & ToVarchar(str诊断, 40) & "|" & ToVarchar(UserInfo.姓名, 20) & "|0"
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    str顺序号 = arrOutput(1)
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        "0,0,0,0,0,0,0,0,0,0,'')"
    Call ExecuteProcedure("重庆医保")
    
    '强制把登记顺序号、及新的医保号填入
    gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & gintInsure & _
                ",'" & str卡号 & "','" & str医保号 & "','" & str顺序号 & "')"
    Call ExecuteProcedure("重庆医保")
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("重庆医保")
    
    入院登记_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_重庆(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim datCurr As Date, rsTemp As New ADODB.Recordset
    Dim strInput As String, arrOutput  As Variant, bln零费用出院 As Boolean
    Dim str诊断 As String
    
    On Error GoTo errHandle
    
    '获得病人出院诊断
    gstrSQL = "select A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型=3 and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "出院登记")
    If rsTemp.EOF = False Then
        str诊断 = NVL(rsTemp("描述信息"), "无")
    Else
        str诊断 = "无"   '诊断不论如何不能为空
    End If
    str诊断 = ToVarchar(str诊断, 40)
    
    '获得其它出院信息
    datCurr = zlDataBase.Currentdate
    gstrSQL = "select A.住院医师,A.入院日期,A.出院日期,A.出院病床,B.名称 as 出院科室 from 病案主页 A,部门表 B " & _
             " Where A.出院科室ID = B.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "出院登记")
    '调用接口，更新其住院信息
    strInput = "03|" & lng病人ID & "_" & lng主页ID & "|0001010010|21|||" & Format(rsTemp("入院日期"), "yyyy-MM-dd") & "||" & _
                Format(rsTemp("出院日期"), "yyyy-MM-dd") & "|||" & ToVarchar(UserInfo.姓名, 20) & "|0"
    
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额  from 病人费用记录 where 病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID
    Call OpenRecordset(rsTemp, "病人出院")
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
    If bln零费用出院 = True Then
        '对于零费用出院，就将其处理为退入院。而不用更新其住院信息
        gstrSQL = "Select 顺序号 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
        Call OpenRecordset(rsTemp, "病人出院")
        strInput = "99|" & rsTemp("顺序号") & "|" & ToVarchar(UserInfo.姓名, 20)
    End If
    
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("重庆医保")
    
    出院登记_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 更新出院疾病_重庆(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str分类 As String, str并发症 As String, str疾病编码 As String
    Dim strInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    '获得病人出院病种及并发症
    gstrSQL = "Select 退休证号 病种编码,并发症 From 保险帐户 Where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "获取病人出院病种及并发症")
    str疾病编码 = NVL(rsTemp!病种编码)
    str并发症 = NVL(rsTemp!并发症)
    
    str分类 = "出院"
    If frm疾病选择重庆.GetCode("", str分类, str疾病编码, str并发症) = False Then
        Exit Function
    End If
    str疾病编码 = ToVarchar(str疾病编码, 20)
    str并发症 = ToVarchar(str并发症, 200)
    
    '调用接口
    strInput = "03|" & lng病人ID & "_" & lng主页ID & "|0000001001|21||||||" & str疾病编码 & "|||" & str并发症
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'退休证号','''" & str疾病编码 & "''')"
    Call ExecuteProcedure("更新病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_重庆市 & ",'并发症','''" & str并发症 & "''')"
    Call ExecuteProcedure("更新并发症")
    
    更新出院疾病_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 撤消医保入院_重庆(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str顺序号 As String) As Boolean
'功能：更新病人的出院疾病。如果是肿瘤，则结算时起付线会减半
    Dim strInput As String, arrOutput  As Variant
    
    On Error GoTo errHandle
    
    '调用接口
    strInput = "99|" & str顺序号 & "|" & ToVarchar(UserInfo.姓名, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    gstrSQL = "ZL_病案主页_撤消医保入院(" & lng病人ID & "," & lng主页ID & ")"
    Call ExecuteProcedure("撤消医保入院")
    
    撤消医保入院_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_重庆(rsExse As Recordset, ByVal lng病人ID As Long, ByVal str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset

    Dim strInput As String, arrOutput   As Variant
    Dim cur个人帐户 As Double, cur统筹支付 As Double, cur大额统筹 As Double, cur公务员补助 As Double, cur发生费用 As Double
    Dim str总金额医院 As String, str总金额医保 As String
    Dim str医生 As String, datCurr As Date, intMsg As Integer
    
    On Error GoTo errHandle
    mlng病人ID = 0         '初始化。只要一选择病人，就会调用本过程，也就会设成0
    
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    rsExse.MoveFirst
    
    datCurr = zlDataBase.Currentdate
    With g结算数据
        .病人ID = rsExse("病人ID")
        
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & rsExse("病人ID")
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        .主页ID = rsTemp("主页ID")
        .年度 = Int(Format(datCurr, "yyyy"))
    End With
    
    Screen.MousePointer = vbHourglass
    '1.2 读出病人的入院时间
    gstrSQL = "select 入院日期,nvl(出院日期,to_date('3000-01-01','yyyy-MM-dd')) as 出院日期 " & _
              "from 病案主页 where 病人ID=" & g结算数据.病人ID & " and 主页ID=" & g结算数据.主页ID
    Call OpenRecordset(rsTemp, "虚拟结算")
    If rsTemp("出院日期") = CDate("3000-01-01") Then
        g结算数据.中途结帐 = 1
        g结算数据.住院床日 = DateDiff("d", rsTemp("入院日期"), datCurr)
    Else
        '表示该病人已经出院
        g结算数据.中途结帐 = 0
        g结算数据.住院床日 = DateDiff("d", rsTemp("入院日期"), rsTemp("出院日期"))
    End If
    If g结算数据.住院床日 < 1 Then g结算数据.住院床日 = 1 '至少有一天
    
    
    Do Until rsExse.EOF
        cur发生费用 = cur发生费用 + rsExse("金额")
        rsExse.MoveNext
    Loop
    cur发生费用 = Val(Format(cur发生费用, "#####0.00"))
    
    '只有出院结算才上传所有未上传明细，中途结算只对已上传数据进行结算
    If g结算数据.中途结帐 = 0 Then
        '读出未上传明细
        gstrSQL = "Select A.ID,A.NO,A.记录性质,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额" & _
                  "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
                  "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
                  "  From 病人费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
                  "  where A.病人ID=" & lng病人ID & " and A.主页ID=" & g结算数据.主页ID & " and A.记帐费用=1 and A.实收金额<>0 and nvl(A.是否上传,0)=0 " & _
                  "        and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & gintInsure & _
                  "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 " & _
                  "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
                  "  Order by A.发生时间"
        Call OpenRecordset(rs明细, "虚拟结算")
        
        '打开另外一个连接串，以达到不受当前连接事务的控制
        cn上传.ConnectionString = gcnOracle.ConnectionString
        cn上传.Open
        
        intMsg = 0
        Do Until rs明细.EOF
            '只上传只传递过的数据
            str医生 = ToVarchar(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 20)
            
            strInput = "04|" & lng病人ID & "_" & g结算数据.主页ID
            strInput = strInput & "|" & rs明细("NO") & "_" & rs明细("记录性质")
            strInput = strInput & "|" & Format(rs明细("登记时间"), "yyyy-MM-dd HH:mm:ss")
            strInput = strInput & "|" & ToVarchar(rs明细("项目编码"), 10) '中心编码
            strInput = strInput & "|" & ToVarchar(rs明细("编码"), 20) '医院内码
            strInput = strInput & "|" & ToVarchar(rs明细("名称"), 50)     '项目名称
            strInput = strInput & "|" & Format(rs明细("价格"), "0.0000")      '单价
            strInput = strInput & "|" & Format(rs明细("数量"), "0.00")        '数量
            strInput = strInput & "|" & IIf(rs明细("是否急诊") = 1, 1, 0)     '急诊标志
            strInput = strInput & "|" & str医生                               '医生
            strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)          '经办人
            strInput = strInput & "|" & ToVarchar(rs明细("计算单位"), 20)     '单位
            strInput = strInput & "|" & ToVarchar(rs明细("规格"), 14)         '规格
            strInput = strInput & "|" & ToVarchar(rs明细("剂型"), 20)         '剂理
            strInput = strInput & "|"                                         '冲销明细流水号
            strInput = strInput & "|" & Format(rs明细("实收金额"), "#####0.0000")         '金额
            
            If HandleBusiness(strInput, arrOutput) = False Then
                '费用上传失败
                If MsgBox("单据“" & rs明细("NO") & "”中" & rs明细("名称") & "费用上传失败，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
                If intMsg = 0 Then
                    If MsgBox("上传数据失败，是否停止数据上传并直接进行结帐？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                        intMsg = 1
                        Exit Do
                    Else
                        intMsg = -1
                    End If
                End If
            Else
                '费用上传成功才做上标记
                gstrSQL = "ZL_病人记帐记录_上传(" & rs明细("ID") & "," & Val(arrOutput(2)) * rs明细("数量") & ",'" & arrOutput(1) & "')"
                '与其它地方的上传不同，没有采了另一个连接串执行。因为如果出错，可以与该单据一起回滚。
                cn上传.Execute gstrSQL, , adCmdStoredProc
            End If
            
            rs明细.MoveNext
        Loop
    End If
    
    '调用预结算
    strInput = "06|" & lng病人ID & "_" & g结算数据.主页ID
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    cur个人帐户 = Val(arrOutput(2))
    cur统筹支付 = Val(arrOutput(1))
    cur大额统筹 = Val(arrOutput(5))
    cur公务员补助 = Val(arrOutput(3))
    
    '保存病人个人帐户余额
    mstr医保号 = str医保号
    mdbl余额 = cur个人帐户
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    
    str总金额医院 = Format(cur发生费用, "#####0.00")
    str总金额医保 = Format(cur统筹支付 + cur个人帐户 + cur公务员补助 + cur大额统筹 + Val(arrOutput(4)), "#####0.00")
    If str总金额医院 <> str总金额医保 Then
        If MsgBox("医院的费用总金额(" & str总金额医院 & ")与医保中心的费用总额(" & str总金额医保 & ")不等，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    住院虚拟结算_重庆 = "医保基金;" & cur统筹支付 & ";0"
    If cur个人帐户 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|个人帐户;" & cur个人帐户 & ";0" '不允许修改个人帐户
    End If
    If cur大额统筹 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|大额统筹;" & cur大额统筹 & ";0"
    End If
    If cur公务员补助 <> 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|公务员补助;" & cur公务员补助 & ";0"
    End If
    If Val(arrOutput(6)) > 0 Then
        住院虚拟结算_重庆 = 住院虚拟结算_重庆 & "|公务员返还;" & Val(arrOutput(6)) & ";0"
    End If
    
    '保存预结算金额，在结算时再比较一次，避免出现差错
    With g结算数据
        .统筹报销金额 = cur统筹支付       '1
        .个人帐户支付 = cur个人帐户       '2
        .累计进入统筹 = cur公务员补助     '3
        .全自费金额 = Val(arrOutput(4))   '4
        .进入统筹金额 = cur大额统筹       '5
        .累计统筹报销 = Val(arrOutput(6)) '6
    End With
    
    mlng病人ID = lng病人ID  '表示该病人已经进行了虚拟结算
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_重庆(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim str病种 As String
    Dim str操作员 As String, lng结算标志 As Long
    Dim cur统筹支付 As Double, cur个人帐户 As Double
    Dim cur大额统筹 As Double, cur公务员补助 As Double
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, datCurr As Date, strNO As String
    Dim strFormat As String
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    '调用预结算
    strInput = "06|" & lng病人ID & "_" & g结算数据.主页ID
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    '将返回结算与预结算的再一次比较
    strFormat = "#####0.00;-#####0.00;0;"
    With g结算数据
        If CDbl(Format(.统筹报销金额, strFormat)) <> CDbl(Format(arrOutput(1), strFormat)) Or _
           CDbl(Format(.个人帐户支付, strFormat)) <> CDbl(Format(arrOutput(2), strFormat)) Or _
           CDbl(Format(.累计进入统筹, strFormat)) <> CDbl(Format(arrOutput(3), strFormat)) Or _
           CDbl(Format(.全自费金额, strFormat)) <> CDbl(Format(arrOutput(4), strFormat)) Or _
           CDbl(Format(.进入统筹金额, strFormat)) <> CDbl(Format(arrOutput(5), strFormat)) Or _
           CDbl(Format(.累计统筹报销, strFormat)) <> CDbl(Format(arrOutput(6), strFormat)) Then
            
           If MsgBox("结算数据与预结算的结果不一致，可能是由于病人又有新的费用上传。继续结算吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End With
    
    '求个人帐户支付金额
    gstrSQL = "Select Nvl(冲预交,0) as 金额 From 病人预交记录 Where 结算方式='个人帐户' And 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "住院结算")
    If Not rsTemp.EOF Then cur个人帐户 = rsTemp("金额")
    
    '调用结算
    With g结算数据
        If .中途结帐 = 1 Then
'            If MsgBox("该病人是否进行转家庭病床结算？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
'                lng结算标志 = 20 '出院转家庭病床
'            Else
                lng结算标志 = 10 '中途结算
'            End If
        Else
            lng结算标志 = 0 '正常结算
        End If
        
        strInput = "05|" & lng病人ID & "_" & .主页ID & "|" & lng结算标志 & "|" & g结算数据.住院床日 & "|" & UserInfo.姓名 & "|0" '用个人帐户余额支付
    End With
    
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '填写结算表
    datCurr = zlDataBase.Currentdate
    cur统筹支付 = Val(arrOutput(2))
    cur公务员补助 = Val(arrOutput(4))
    cur大额统筹 = Val(arrOutput(6))
    
    Call 读取病种(lng病人ID, str病种)
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
        cur进入统筹累计 + cur统筹支付 & "," & _
        cur统筹报销累计 + cur统筹支付 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("重庆医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,NULL,0," & g结算数据.发生费用金额 & ",0,0," & _
        cur统筹支付 & "," & cur统筹支付 & ",0,0,0,'" & arrOutput(1) & "'," & g结算数据.主页ID & "," & g结算数据.中途结帐 & "," & IIf(str病种 = "", "NULL", "'" & str病种 & "'") & ")"
    Call ExecuteProcedure("重庆医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 & "," & cur统筹支付 & ",NULL)"
    Call ExecuteProcedure("重庆医保")
    
    住院结算_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_重庆(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim curDate As Date
        
    On Error GoTo errHandle
    curDate = zlDataBase.Currentdate
    
    '退费
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
    If Can住院结算冲销(rsTemp("病人ID"), rsTemp("主页ID")) = False Then Exit Function
    
    str流水号 = rsTemp("支付顺序号")
    
    strInput = "99|" & str流水号 & "|" & ToVarchar(UserInfo.姓名, 20)
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    
    '帐户年度信息
    Call Get帐户信息(rsTemp("病人ID"), Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & rsTemp("病人ID") & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("重庆医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & gintInsure & "," & rsTemp("病人ID") & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - rsTemp("个人帐户支付") & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & rsTemp("发生费用金额") * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & _
        rsTemp("个人帐户支付") * -1 & ",'" & str流水号 & "'," & rsTemp("主页ID") & "," & rsTemp("中途结帐") & ")"
    Call ExecuteProcedure("重庆医保")

    住院结算冲销_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 错误信息_重庆(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Public Function 医院编码_重庆() As String
'功能：得到医院的医保编码
    Dim strInput As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    strInput = "11"
    If HandleBusiness(strInput, arrOutput) = False Then Exit Function
    医院编码_重庆 = arrOutput(1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function HandleBusiness(ByVal strInput As String, varOut As Variant) As Boolean
'功能：调用医保部件，进行业务处理
    Dim strReturn As String '调用前置服务器的返回值
    Dim lngReturn As Long
    Dim varArray As Variant, lngCount As Long
    
    On Error Resume Next
    varOut = ""
    Screen.MousePointer = vbHourglass
    strReturn = Space(1024)
    lngReturn = dy_Business_Handle(strInput, strReturn)
    If Err <> 0 Or lngReturn = -1 Then
        varArray = Split(strReturn, "|")
        
        If UBound(varArray) > 0 Then
            strReturn = "医保接口调用失败。" & vbCrLf & varArray(1)
        Else
            strReturn = "医保接口调用失败。" & vbCrLf & strReturn
        End If
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    strReturn = TruncZero(strReturn)
    
    varArray = Split(strReturn, "|")
    If varArray(0) = "-1" Then
        '业务调用失败
        If UBound(varArray) > 0 Then
            strReturn = "医保接口出现警告。" & vbCrLf & varArray(1)
        Else
            strReturn = "医保业务处理失败。"
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '交易成功
    varOut = Split(strReturn, "|")
    
    HandleBusiness = True
    Screen.MousePointer = vbDefault
End Function

Private Function Get保险参数_重庆(ByVal str参数名 As String) As String
'功能：获得保险参数
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select A.参数名,A.参数值 from 保险参数 A " & _
              " where A.参数名='" & str参数名 & "' and A.险类=" & TYPE_重庆市 & " and A.中心 is null "
    Call OpenRecordset(rsTemp, "重庆医保")
    
    If rsTemp.EOF = False Then
        Get保险参数_重庆 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
    End If
End Function

Public Function 价格判断_重庆(ByVal dbl医院 As Double, ByVal dbl医保 As Double, ByVal str限价方式 As String, _
                              ByVal bln特价 As Boolean, ByVal dbl特价 As Double) As Boolean
'功能：判断医院的价格是否超过医保规定的单价
    Dim str医院类别 As String
    
    On Error GoTo errHandle
    
    If InStr(str限价方式, "二级") > 0 Then
        str医院类别 = Get保险参数_重庆("医院等级")
        '给出的标准价格为二级医院的最高限价，三级医院的最高限价在此基础上可以上浮10%，一级医院的最高限价在此基础上下调5%
        
        Select Case str医院类别
            Case "三级"
                dbl医保 = dbl医保 * 1.1
            Case "一级"
                dbl医保 = dbl医保 * 0.95
        End Select
    End If
    
    If bln特价 = True And dbl特价 > dbl医保 Then
        '允许使用特价
        dbl医保 = dbl特价
    End If
    
    If dbl医院 > dbl医保 Then
        If MsgBox("医院单价" & Format(dbl医院, "0.000") & " 高于医保中心核准的价格" & Format(dbl医保, "0.000") & "，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    价格判断_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐传输_重庆(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
'功能:上传新产生的记帐明细到医保中心
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
'返回:
    Dim rsTemp As New ADODB.Recordset, cn上传 As New ADODB.Connection
    Dim strInput As String, arrOutput   As Variant, cur统筹金额 As Currency
    Dim str医生 As String, str经办人 As String
    Dim col病人 As New Collection, lngPre病人ID As Long, var病人 As Variant, bln成功 As Boolean
    
    '请注意：重庆医保是在记帐单保存后再调用传输过程的。
    
    On Error GoTo errHandle
    
    cn上传.ConnectionString = gcnOracle.ConnectionString
    cn上传.Open
    
    '读出该张单据的费用明细
    
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.主页ID,A.发生时间 as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         ,C.项目编码,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型 " & _
              "  From 病人费用记录 A,收费细目 B,保险支付项目 C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where A.NO='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态=1 " & _
              "        and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & gintInsure & IIf(lng病人ID = 0, "", " and A.病人ID=" & lng病人ID) & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID and C.险类=D.险类 " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.病人ID,A.发生时间"
    Call OpenRecordset(rsTemp, "记帐传输")
    
    '进行费用明细的传输
    Do Until rsTemp.EOF
        str医生 = ToVarchar(IIf(IsNull(rsTemp("医生")), UserInfo.姓名, rsTemp("医生")), 20)
        str经办人 = ToVarchar(IIf(IsNull(rsTemp("操作员姓名")), UserInfo.姓名, rsTemp("操作员姓名")), 20)
        
        strInput = "04|" & rsTemp("病人ID") & "_" & rsTemp("主页ID")
        strInput = strInput & "|" & rsTemp("NO") & "_" & int性质
        strInput = strInput & "|" & Format(rsTemp("登记时间"), "yyyy-MM-dd HH:mm:ss")
        strInput = strInput & "|" & ToVarchar(rsTemp("项目编码"), 10)     '中心编码
        strInput = strInput & "|" & ToVarchar(rsTemp("编码"), 20)         '医院内码
        strInput = strInput & "|" & ToVarchar(rsTemp("名称"), 50)         '项目名称
        strInput = strInput & "|" & Format(rsTemp("价格"), "0.0000")      '单价
        strInput = strInput & "|" & Format(rsTemp("数量"), "0.00")        '数量
        strInput = strInput & "|" & IIf(rsTemp("是否急诊") = 1, 1, 0)     '急诊标志
        strInput = strInput & "|" & str医生                               '医生
        strInput = strInput & "|" & str经办人                             '经办人
        strInput = strInput & "|" & ToVarchar(rsTemp("计算单位"), 20)     '单位
        strInput = strInput & "|" & ToVarchar(rsTemp("规格"), 14)         '规格
        strInput = strInput & "|" & ToVarchar(rsTemp("剂型"), 20)         '剂理
        strInput = strInput & "|"                                         '冲销明细流水号
        strInput = strInput & "|" & Format(rsTemp("实收金额"), "#####0.0000")         '金额
        
        If HandleBusiness(strInput, arrOutput) = False Then
            '如果费用上传失败，则冲正已经上传的交易
            '冲销采用按单据进行而不是每笔明细，主要考虑减少网络传输
'            For Each var病人 In col病人
'                strInput = "10|" & var病人 & "|" & rsTemp("NO") & "_" & int性质
'                Call HandleBusiness(strInput, arrOutput)
'            Next
'
            If bln成功 = True Then
                MsgBox "数据上传中途发生错误，并且已经部分已经上传，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
            Else
                MsgBox "数据上传发生错误，没有成功上传的记录，请在预结算处完成剩余数据的上传。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        Call AddMessage(str消息, arrOutput, rsTemp("名称"), rsTemp("价格")) '可以产生的提醒信息
        
        If lngPre病人ID <> rsTemp("病人ID") Then '判断时没有考虑主页ID，是因为同一病人不可能同时有两次住院的明细
            '将已经上传的病人信息记录下来（因为记帐表是多病人的）
            col病人.Add rsTemp("病人ID") & "_" & rsTemp("主页ID")
            lngPre病人ID = rsTemp("病人ID")
        End If
        
        '在费用记录上打上标记，说明已经上传，并保存返回的金额
        If arrOutput(3) = 2 Then
            '未通过审批
            cur统筹金额 = 0
        Else
            '特准单价 * 数量
            cur统筹金额 = Val(arrOutput(2)) * rsTemp("数量")
        End If
        gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & "," & cur统筹金额 & ",'" & arrOutput(1) & "')"
        '与其它地方的上传不同，没有采了另一个连接串执行。因为如果出错，可以与该单据一起回滚。
        cn上传.Execute gstrSQL, , adCmdStoredProc
        bln成功 = True
        
        rsTemp.MoveNext
    Loop
    
    If str消息 <> "" Then
        str消息 = "病人费用明细传输过程中得到医保中心如下反馈信息，但目前数据已经保存。" & vbCrLf & "如果有何不妥，你可以选择作废该单据。" & vbCrLf & vbCrLf & str消息
    End If
        
    记帐传输_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐作废_重庆(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String) As Boolean
'功能:作废已经上传到医保中心的记帐明细
'参数:  str单据号   NO
'       int性质     记录性质
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'返回:
    
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, arrOutput As Variant
    Dim lngPre病人ID As Long
    
    On Error GoTo errHandle
    
    '读出该张单据的费用明细中有未上传的记录（取原始单据）
    gstrSQL = "Select nvl(count(A.ID),0) as 总数,nvl(sum(A.是否上传),0) 上传数 " & _
              "  From 病人费用记录 A,病案主页 B,保险支付项目 C" & _
              "  where A.NO='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态<>2 and nvl(A.实收金额,0)<>0  " & _
              "        and A.病人ID=B.病人ID and A.主页ID=B.主页ID And B.险类=" & gintInsure & " and A.收费细目ID=C.收费细目ID and B.险类=C.险类"
    Call OpenRecordset(rsTemp, "记帐作废")
    
    If rsTemp.EOF = True Then
        MsgBox "该单据里没有可上传的作废费用明细。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp("上传数") = 0 Then
        '明细根本就没有上传，所以也就不需要处理作废
        记帐作废_重庆 = True
        Exit Function
    End If
    
    If rsTemp("上传数") < rsTemp("总数") Then
        MsgBox "该单据里还有未上传的费用明细，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '读出该单据内病人情况（因为记帐表是多病人的）
    gstrSQL = "Select A.ID,A.病人ID,A.主页ID,A.摘要 流水号" & _
              "  From 病人费用记录 A,病案主页 B,保险支付项目 C " & _
              "  where A.NO='" & str单据号 & "' and A.记录性质=" & int性质 & " and A.记录状态<>1 " & _
              "        and A.病人ID=B.病人ID and A.主页ID=B.主页ID And B.险类=" & gintInsure & " and A.收费细目ID=C.收费细目ID and B.险类=C.险类 " & _
              " Order by A.病人ID,A.主页ID"
    Call OpenRecordset(rsTemp, "记帐作废")
    
    '进行费用明细的传输
    Do Until rsTemp.EOF
        '整张单据冲销
        If lngPre病人ID <> rsTemp("病人ID") Then '判断时没有考虑主页ID，是因为同一病人不可能同时有两次住院的明细
            '将已经上传的病人信息记录
            strInput = "10|" & rsTemp("病人ID") & "_" & rsTemp("主页ID") & "|" & str单据号 & "_" & int性质
            If HandleBusiness(strInput, arrOutput) = False Then
                '如果出错，作废仍然继续，但没有打上作废标志
                记帐作废_重庆 = True
                Exit Function
            End If
            lngPre病人ID = rsTemp("病人ID")
        End If
        
        '在产生的作废费用记录上打上标记，说明已经上传
        gstrSQL = "ZL_病人记帐记录_上传(" & rsTemp("ID") & ")"
        '与其它地方的上传不同，没有采了另一个连接串执行。因为如果出错，可以与该单据一起回滚。
        Call ExecuteProcedure("重庆医保")
        
        rsTemp.MoveNext
    Loop
    
    记帐作废_重庆 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AddMessage(strMessage As String, arrOutput As Variant, ByVal str项目 As String, ByVal dbl单价 As Currency)
'功能：在病人费用明细传输过程中可能产生一些需要提醒操作人员的信息
    Dim strTemp As String
    
    If dbl单价 > Val(arrOutput(2)) And Val(arrOutput(2)) > 0 Then
        strTemp = "●    " & str项目 & "的医院价格 " & Format(dbl单价, "0.0000") & " 高于中心返回价格 " & Format(Val(arrOutput(2)), "0.0000") & vbCrLf
    End If
    If arrOutput(3) = 2 Then
        strTemp = "●    " & str项目 & "需要审批，但没有审批记录，只能作为自费项目" & vbCrLf
    End If
    
    If InStr(strMessage, strTemp) = 0 Then
        strMessage = strMessage & strTemp
    End If
    
End Sub

Private Sub 读取病种(ByVal lng病人ID As Long, str病种 As String)
    Dim strServer As String, strUser As String, strPass As String
    Dim strTemp As String
    Dim rs病种 As New ADODB.Recordset
    Dim cnYB As New ADODB.Connection
    
    '读取该医保病人的病种信息
    gstrSQL = "Select 退休证号 As 编码 From 保险帐户 Where 险类=" & gintInsure & " And 病人ID=" & lng病人ID
    Call OpenRecordset(rs病种, "获取医保病人的病种编码")
    str病种 = NVL(rs病种!编码, "")
    
    '如果病种编码不为空，则取病种名称保存于结算记录的备注字段中，以备以后查看
    If str病种 <> "" Then
        '打开前置机连接，读取病种信息
        gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & gintInsure
        Call OpenRecordset(rs病种, "读取保险参数")
        Do Until rs病种.EOF
            strTemp = IIf(IsNull(rs病种("参数值")), "", rs病种("参数值"))
            Select Case rs病种("参数名")
                Case "医保服务器"
                    strServer = strTemp
                Case "医保用户名"
                    strUser = strTemp
                Case "医保用户密码"
                    strPass = strTemp
            End Select
            rs病种.MoveNext
        Loop
        If OraDataOpen(cnYB, strServer, strUser, strPass) Then
            If rs病种.State = adStateOpen Then rs病种.Close
            rs病种.Open "select BZBM 编码,BZMC 名称,ZJM 简码  from BZML Where BZBM='" & str病种 & "'", cnYB
            If rs病种.RecordCount <> 0 Then str病种 = NVL(rs病种!名称, "")
        Else
            str病种 = ""
        End If
        
        '关闭连接
        If cnYB.State = 1 Then cnYB.Close
        Set cnYB = Nothing
    End If
End Sub
