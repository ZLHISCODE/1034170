Attribute VB_Name = "mdl广元"
Option Explicit
'全局变量均在mdl涪陵.bas中进行定义，包括函数原型定义

Private mblnReturn As Boolean

Public Function 医保初始化_广元() As Boolean
'    If gstr医保机构编码 = "" Then
'        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
'checkCard:
'        initType
'        mblnReturn = getybjgbm(gstrOutPara)
'        TrimType
'        If mblnReturn = False Then
'            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
'                GoTo checkCard
'            Else
'                Exit Function
'            End If
'        End If
'        gstr医保机构编码 = gstrOutPara.out1
'        gstr医院编码 = gstrOutPara.out2
'    End If
    医保初始化_广元 = True
End Function

Public Function 身份标识_广元(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim frmIDentified As New frmIdentify广元
    Dim strPatiInfo As String, cur余额 As Currency, str就诊编号 As String
    Dim arr, datCurr As Date, str门诊号 As String
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Dim strTemp As String
    If lng病人ID = 0 Then
        strTemp = "0"
    Else
        gstrSQL = "Select * From 保险帐户 where 病人id=" & lng病人ID
        OpenRecordset rsTemp, gstrSysName
        If rsTemp.EOF Then
            strTemp = "0"
        Else
            strTemp = rsTemp!退休证号
        End If
    End If
    
    strPatiInfo = frmIDentified.GetPatient(bytType, strTemp)
    
    On Error GoTo errHandle
    If strPatiInfo <> "" Then
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
        If lng病人ID = 0 Then
            lng病人ID = BuildPatiInfo(bytType, strPatiInfo, lng病人ID)
        End If
        '返回格式:中间插入病人ID
        strPatiInfo = frmIDentified.mstrPatient & lng病人ID & ";" & frmIDentified.mstrOther
        str就诊编号 = frmIDentified.mstr就诊编号
        '写入就诊编号
        If bytType = 0 Or bytType = 5 Then
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_广元 & ",'顺序号','''" & str就诊编号 & "''')"
            Call ExecuteProcedure("身份标识_广元")
            gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_广元 & ",'退休证号','''" & CLng(strTemp) + 1 & "''')"
            Call ExecuteProcedure("身份标识_广元")
        End If
        Unload frmIDentified
    Else
        身份标识_广元 = ""
        MsgBox "医保病人信息提取失败", vbInformation, gstrSysName
        Unload frmIDentified
        Exit Function
    End If
    身份标识_广元 = strPatiInfo
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_广元 = ""
End Function

Public Function 个人余额_广元(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_广元
    Call OpenRecordset(rsTemp, "读取个人帐户余额")
    
    If rsTemp.EOF Then
        个人余额_广元 = 0
    Else
        个人余额_广元 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
    End If
End Function

Public Function 门诊虚拟结算_广元(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim cur自付 As Currency, cur报销 As Currency, cur余额 As Currency, lngErr As Long
    Dim lng病人ID As Long, rsTemp As New ADODB.Recordset, str报销明细 As String
    Dim strTemp As String, curTemp As Currency, str自付比例 As String, str可报销额 As String
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有费用，不能进行预结算。", vbInformation, gstrSysName
        门诊虚拟结算_广元 = False
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID"): lngErr = 1
    cur自付 = 0: cur报销 = 0: lngErr = 2
    gstrSQL = "Select * from 保险帐户 where 病人id=" & lng病人ID: lngErr = 3
    OpenRecordset rsTemp, "医保预结算": lngErr = 4
    cur余额 = rsTemp!帐户余额: lngErr = 5
    strTemp = rsTemp!在职: lngErr = 4
    str报销明细 = ""
    While Not rs明细.EOF
        gstrSQL = "select * from 收费细目 where id=" & rs明细!收费细目ID: lngErr = 6
        OpenRecordset rsTemp, "医保预结算": lngErr = 7
        
        '获取收费细目的自付比例
        initType
        mblnReturn = readzfbl(gstr医保机构编码, gstr医院编码, rsTemp!类别 & "_" & rsTemp!ID, _
            IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), _
            strTemp, gstrOutPara): lngErr = 8
        TrimType
        
        If mblnReturn = False Then
            MsgBox "在获取项目[" & rsTemp!ID & "]的自付比例时，医保接口返回以下错误：" & Chr(13) & Chr(10) & gstrOutPara.errtext
            门诊虚拟结算_广元 = False
            Exit Function
        End If
        Select Case gstrOutPara.out2
            Case "1"            '返回为自付比例
                curTemp = rs明细!实收金额 * (1 - CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0))): lngErr = 9
            Case "2"            '返回为报销限额
                curTemp = IIf(rs明细!实收金额 > CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), CCur(IIf(IsNumeric(gstrOutPara.out1), gstrOutPara.out1, 0)), rs明细!实收金额): lngErr = 10
            Case "3"            '按自付比例计算报销金额，若大于可报销额，则取可报销额
                str自付比例 = Left(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") - 1): lngErr = 11
                str可报销额 = Mid(gstrOutPara.out1, InStr(gstrOutPara.out1, "|") + 1): lngErr = 12
                str自付比例 = IIf(IsNumeric(str自付比例), str自付比例, 0): lngErr = 13
                str可报销额 = IIf(IsNumeric(str可报销额), str可报销额, 0): lngErr = 14
                curTemp = rs明细!实收金额 * (1 - CCur(str自付比例)): lngErr = 15
                curTemp = IIf(curTemp > CCur(str可报销额), CCur(str可报销额), curTemp): lngErr = 16
            Case "4", "5"       '自付比例为100%
                curTemp = 0
        End Select
        str报销明细 = str报销明细 & "项目名称:" & rsTemp!名称 & "[" & rsTemp!类别 & "_" & rsTemp!ID & "]　　自付比例:[" & _
            gstrOutPara.out2 & "]" & gstrOutPara.out1 & "　　报销金额:" & curTemp & Chr(13) & Chr(10)
        
        cur报销 = cur报销 + curTemp: lngErr = 17
        cur自付 = rs明细!实收金额 - curTemp: lngErr = 18
        rs明细.MoveNext: lngErr = 19
    Wend
    
    '如果报销额大于帐户余额，则允许从帐户中支付的最大额为帐户余额，多余部分计入现金支付
    If cur报销 > cur余额 Then
        curTemp = cur报销 - cur余额: lngErr = 20
        cur报销 = cur余额: lngErr = 21
        cur自付 = cur自付 + curTemp: lngErr = 22
    End If
    
'    MsgBox str报销明细, vbInformation, "报销明细"
    
    str结算方式 = "个人帐户;" & cur报销 & ";0": lngErr = 23
    门诊虚拟结算_广元 = True
    Exit Function
errHandle:
    MsgBox "错误出现在[门诊预结算]模块，第" & lngErr & "行，错误信息：" & Chr(13) & Chr(10) & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Public Function 门诊结算_广元(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur起付线 As Currency, cur基本统筹限额 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur先自付 As Currency, lng病种ID As Long
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs明细, gstrSysName)
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(IIf(IsNull(rs明细("操作员姓名")), UserInfo.姓名, rs明细("操作员姓名")), 20)
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & gintInsure

    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
    Else
        门诊结算_广元 = False
        Exit Function
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_广元 & ",'病种ID'," & lng病种ID & ")"
    Call ExecuteProcedure("身份标识_广元")

    '需要先上传费用明细
    费用明细传递_广元 lng结帐ID
    
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,病种id From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
    Set rsTemp = gcnOracle.Execute(gstrSQL)
    lng病种ID = rsTemp!病种ID
    str就诊编号 = rsTemp!顺序号
    
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
'    mblnReturn = pcalc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "1", "0", gstrOutPara)
    mblnReturn = calc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        门诊结算_广元 = False
        Exit Function
    End If
'间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
'             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
'             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额
    
    '获取个人帐户支付和个人现金支付
    cur个人帐户 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur余额 = CCur(gstrOutPara.out21)
    cur全自付 = CCur(gstrOutPara.out20) + CCur(cur个人帐户)
    cur发生费用 = CCur(gstrOutPara.out1)
    cur先自付 = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & Get病人ID(CStr(str医保号), CStr(gintInsure)) & _
            "," & gintInsure & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call ExecuteProcedure(gstrSysName)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & _
            Get病人ID(CStr(str医保号), CStr(gintInsure)) & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL,NULL,NULL,NULL," & _
            cur个人帐户 & ",NULL)"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    门诊结算_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 费用明细传递_广元(lng结帐ID As Long, Optional rs明细IN As ADODB.Recordset = Nothing) As Boolean
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, cur发生费用, str就诊编号 As String, strBillNO As String
    Dim lng病种ID As Long, str病种名称 As String, str病种编码 As String, int特病标志 As Integer
    Dim str科室编号 As String, str科室名称 As String, lng科室ID As Long
    Dim str明细编码 As String, str明细名称 As String, str处方号 As String
    Dim strTemp As String, iLoop As Long
'    病人ID         adBigInt, 19, adFldIsNullable
'    收费类别       adVarChar, 2, adFldIsNullable
'    收据费目       adVarChar, 20, adFldIsNullable
'    计算单位       adVarChar, 6, adFldIsNullable
'    开单人         adVarChar, 20, adFldIsNullable
'    收费细目ID     adBigInt, 19, adFldIsNullable
'    数量           adSingle, 15, adFldIsNullable
'    单价           adSingle, 15, adFldIsNullable
'    实收金额       adSingle, 15, adFldIsNullable
'    统筹金额       adSingle, 15, adFldIsNullable
'    保险支付大类ID adBigInt, 19, adFldIsNullable
'    是否医保       adBigInt, 19, adFldIsNullable
'    摘要           adVarChar, 200, adFldIsNullable
'    是否急诊       adBigInt, 19, adFldIsNullable
'    str结算方式  "报销方式;金额;是否允许修改|...."
    
    On Error GoTo errHandle
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    If rs明细IN Is Nothing Then
        gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
        Call OpenRecordset(rs明细, gstrSysName)
    Else
        Set rs明细 = rs明细IN.Clone
    End If
    If rs明细.EOF = True Then
'        MsgBox "没有需要上传的收费记录", vbExclamation, gstrSysName
        费用明细传递_广元 = False
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = ToVarchar(UserInfo.姓名, 20)
    
'    gstrSQL = "select max(主页ID) as 主页ID from 病案主页 where 病人ID =" & lng病人ID
'    Call OpenRecordset(rsTemp, gstrsysname)
'    strBillNo = CStr(lng病人ID) & "_" & CStr(rsTemp("主页ID"))
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,病种ID,中心,退休证号 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
    Call OpenRecordset(rsTemp, gstrSysName)
    str处方号 = rsTemp!退休证号
    str就诊编号 = rsTemp!顺序号
    lng病种ID = NVL(rsTemp!病种ID, 0)
'    gstr医保机构编码 = rsTemp!中心
    gstrSQL = "Select * From 保险病种 Where ID=" & lng病种ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        str病种名称 = "未知"
        str病种编码 = "0"
        int特病标志 = 0
    Else
        str病种名称 = rsTemp!名称
        str病种编码 = rsTemp!ID
        int特病标志 = IIf(rsTemp!类别 = 2, 1, 0)
    End If
    lng科室ID = rs明细!开单部门ID
    gstrSQL = "Select * From 部门表 where id=" & lng科室ID
    Call OpenRecordset(rsTemp, gstrSysName)
    str科室编号 = rsTemp!编码
    str科室名称 = rsTemp!名称
    
'    str处方号 = NVL(rs明细!主页ID, 0) & Right(rs明细!NO, 2)
    '写处方信息
    initType
    mblnReturn = wrecipe(gstr医保机构编码, gstr医院编码, str就诊编号, str处方号, str病种编码, str病种名称, _
                         int特病标志, NVL(rs明细!开单人, rs明细!划价人), NVL(rs明细!操作员姓名, UserInfo.姓名), str科室编号, _
                         str科室名称, Format(rs明细!登记时间, "yyyy-MM-dd"), gstrOutPara)
    TrimType
    If mblnReturn = False Then
        If InStr(gstrOutPara.errtext, "(YBYY.PRI_QTYL42_T)") > 0 Then
            费用明细传递_广元 = True
        Else
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            费用明细传递_广元 = False
            Exit Function
        End If
    End If
    gcnOracle.Execute "Update 保险帐户 set 退休证号=" & CLng(str处方号) + 1 & " where 病人id=" & lng病人ID
    iLoop = 1
    '写处方明细
    Do Until rs明细.EOF
        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        str明细编码 = rsTemp!ID
        str明细名称 = rsTemp!名称
        initType
        If InStr(NVL(rsTemp!规格, " "), "┆") > 0 Then
            strTemp = Left(rsTemp!规格, InStr(rsTemp!规格, "┆") - 1)
        Else
            strTemp = NVL(rsTemp!规格, " ")
        End If
'入口参数:医保机构编码,医院编号,医保就诊编号,处方编号,明细序号,医院明细编码,医院明细名称,产地,规格,类别,
'         单位,单价,数量,时间,录入人,标志
        If IsNull(rs明细!是否上传) Or rs明细!是否上传 = 0 Then
            mblnReturn = wdetails(gstr医保机构编码, gstr医院编码, str就诊编号, str处方号, iLoop, _
                rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, " ", strTemp, NVL(rsTemp!费用类型, " "), NVL(rsTemp!计算单位, " "), rs明细!标准单价, _
                rs明细!付数 * rs明细!数次, Format(rs明细!登记时间, "yyyy-MM-dd"), NVL(rs明细!操作员姓名, UserInfo.姓名), _
                IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), gstrOutPara)
'        Else
'            mblnReturn = udetails(gstr医保机构编码, gstr医院编码, str就诊编号, str处方号, rs明细!序号, _
'                rsTemp!类别 & "_" & rsTemp!ID, rsTemp!名称, " ", strTemp, NVL(rsTemp!费用类型, " "), NVL(rsTemp!计算单位, " "), rs明细!标准单价, _
'                rs明细!付数 * rs明细!数次, Format(rs明细!登记时间, "yyyy-MM-dd"), NVL(rs明细!操作员姓名, UserInfo.姓名), _
'                IIf(rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7", "1", IIf(rsTemp!类别 = "J", "3", "2")), gstrOutPara)
        End If
        TrimType
        If mblnReturn = False Then
            MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
            费用明细传递_广元 = False
            Exit Function
        End If
        gstrSQL = "Update 病人费用记录 Set 是否上传=1 Where ID='" & rs明细!ID & "'"
        gcnOracle.Execute gstrSQL
        rs明细.MoveNext
        iLoop = iLoop + 1
    Loop
'    rs明细.MoveFirst
'    If lng结帐ID = 0 Then
'
'    Else
'        gstrSQL = "Update 病人费用记录 Set 是否上传=1 Where 结帐ID=" & lng结帐ID
'    End If
'    gcnOracle.Execute gstrSQL
    费用明细传递_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_广元(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID: lngErr = 1
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=" & lng病人ID: lngErr = 2
    Call OpenRecordset(rsTemp, gstrSysName)
    str就诊编号 = NVL(rsTemp!顺序号, "0")
'    gstr医保机构编码 = rsTemp!中心
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID: lngErr = 3
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID: lngErr = 4
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    '调用接口数冲销
    initType
    mblnReturn = canrollback(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox "判断是否可以冲销时，医保端返回以下信息，退费不能继续。" & Chr(13) & Chr(10) & gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    initType
    mblnReturn = rollbackcalc(gstr医保机构编码, gstr医院编码, str就诊编号, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计): lngErr = 5
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - NVL(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - NVL(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")": lngErr = 6
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        NVL(rsTemp("进入统筹金额"), 0) * -1 & "," & NVL(rsTemp("统筹报销金额"), 0) * -1 & ",0," & NVL(rsTemp("超限自付金额"), 0) & "," & _
        cur个人帐户 * -1 & ",'" & str流水号 & "')": lngErr = 7
    Call ExecuteProcedure(gstrSysName)

    门诊结算冲销_广元 = True
    Exit Function
errHandle:
    MsgBox "错误发生在[门诊结算冲销]模块，第" & lngErr & "行，错误信息：" & Chr(13) & Chr(10) & Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_广元(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strSql As String, strInNote As String, rsTemp As New ADODB.Recordset, str病种 As String, str病种编码 As String
    Dim rsTmp As New ADODB.Recordset, str就诊编号 As String, datCurr As Date
    Dim lng病种ID As Long
    
    '求出病人的相关信息
    datCurr = zlDatabase.Currentdate
    On Error GoTo errHandle
    gstrSQL = "select A.入院日期,B.住院号,D.名称 as 住院科室,A.入院病床,A.住院医师,C.卡号," & _
            "C.密码 from 病案主页 A,病人信息 B,保险帐户 C,部门表 D " & _
            "Where A.病人ID = B.病人ID And A.病人ID = C.病人ID And " & _
            "A.入院科室ID = D.ID And A.主页ID = " & lng主页ID & " And A.病人ID = " & lng病人ID
    Call OpenRecordset(rsTmp, gstrSysName)
    
    strInNote = 获取入出院诊断(lng病人ID, lng主页ID)   '入院诊断
    If rsTmp.BOF Then 入院登记_广元 = False: Exit Function
    '强制选择病种
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & gintInsure
    
    Set rsTemp = frmPubSel.ShowSelect(Nothing, gstrSQL, 0, "医保病种")
    If rsTemp.State = 1 Then
        lng病种ID = rsTemp("ID")
        str病种 = rsTemp!名称
        str病种编码 = rsTemp!ID
    Else
        入院登记_广元 = False
        Exit Function
    End If
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If

    initType
    mblnReturn = reg(gstr医保机构编码, gstr医院编码, 1, UserInfo.姓名, Format(zlDatabase.Currentdate, "yyyy-MM-dd"), "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        入院登记_广元 = False
        Exit Function
    End If
    str就诊编号 = gstrOutPara.out1
    
    initType
'入口参数:医保机构编码,医院编号,医保就诊编号,医院疾病编码,医院疾病名称,申请日期,原因
'         急诊标志, 医生姓名,特病标志
    '进行入院请求
    mblnReturn = request(gstr医保机构编码, gstr医院编码, str就诊编号, str病种编码, str病种, Format(datCurr, "yyyy-MM-dd"), _
            strInNote, "0", UserInfo.姓名, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        入院登记_广元 = False
        Exit Function
    End If
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_广元 & ",'顺序号'," & str就诊编号 & ")"
    Call ExecuteProcedure("身份标识_广元")
    
    gstrSQL = "ZL_保险帐户_更新信息(" & lng病人ID & "," & TYPE_广元 & ",'病种ID'," & lng病种ID & ")"
    Call ExecuteProcedure("身份标识_广元")
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    入院登记_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_广元 = False
End Function

Public Function 住院结算冲销_广元(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String, lng病人ID As Long
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur票据总金额 As Currency
    Dim datCurr As Date, cur个人帐户 As Currency
        
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    MsgBox "已结算的数据不允许冲销", vbInformation, gstrSysName
    住院结算冲销_广元 = False
    Exit Function
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "Select * from 保险帐户 where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
    Call OpenRecordset(rsTemp, gstrSysName)
    str就诊编号 = NVL(rsTemp!顺序号, "0")
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B" & _
              " where b.nvl(附加标志,0)<>9 and a.nvl(附加标志,0)<>9 and A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    cur个人帐户 = rsTemp!个人帐户支付
    '调用接口数冲销
    initType
    mblnReturn = canrollback(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    initType
    mblnReturn = rollbackcalc(gstr医保机构编码, gstr医院编码, str就诊编号, "0", gstrOutPara)
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 - rsTemp("进入统筹金额") & "," & _
        cur统筹报销累计 - rsTemp("统筹报销金额") & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0," & rsTemp("超限自付金额") & "," & _
        cur个人帐户 * -1 & ",'" & str流水号 & "')"
    Call ExecuteProcedure(gstrSysName)

    住院结算冲销_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_广元(lng结帐ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；
'        当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结
'        果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String
    Dim int住院次数累计 As Integer, cur帐户增加累计 As Currency
    Dim cur帐户支出累计 As Currency, cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim cur个人帐户 As Currency, cur起付线 As Currency, cur基本统筹限额 As Currency
    Dim cur大额统筹限额 As Currency, cur基数自付 As Currency, cur余额 As Currency
    Dim cur发生费用 As Currency, cur全自付 As Currency, cur先自付 As Currency
    
    On Error GoTo errHandle
    '需要先上传费用明细
'    费用明细传递_广元 lng结帐ID
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs明细, gstrSysName)
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    
    lng病人ID = rs明细("病人ID")
    str操作员 = UserInfo.姓名
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
    Call OpenRecordset(rsTemp, gstrSysName)
    str就诊编号 = rsTemp!顺序号
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = calc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        住院结算_广元 = False
        Exit Function
    End If
'间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
'             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
'             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额

    '获取个人帐户支付和个人现金支付
    cur个人帐户 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur余额 = CCur(gstrOutPara.out21)
    cur全自付 = CCur(gstrOutPara.out20) - cur个人帐户
    cur发生费用 = CCur(gstrOutPara.out1)
    cur先自付 = CCur(gstrOutPara.out10) + CCur(gstrOutPara.out11)
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & _
            "," & gintInsure & "," & Year(datCurr) & "," & cur帐户增加累计 & _
            "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & _
            cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call ExecuteProcedure(gstrSysName)
    '保险结算记录
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & _
            lng病人ID & "," & Year(datCurr) & "," & _
            cur余额 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & ",NULL,NULL,NULL," & _
            cur发生费用 & "," & cur全自付 & "," & cur先自付 & ",NULL,NULL,NULL,NULL," & _
            cur个人帐户 & ",NULL)"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    住院结算_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_广元(rs费用明细 As Recordset, lng病人ID As Long, str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rs费用明细-需要结算的费用明细记录集合
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    
    Dim cur个人帐户支付 As Currency, cur个人现金支付 As Currency
    Dim cur统筹支付 As Currency, cur医保支付 As Currency, cur补充医保 As Currency
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim str操作员 As String, datCurr As Date, str就诊编号 As String
    Dim curCount As Currency
    
    On Error GoTo errHandle
    '需要先上传费用明细
'    费用明细传递_广元 0, rs费用明细
'
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    Set rs明细 = rs费用明细.Clone

    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs明细.EOF
        curCount = curCount + rs明细!金额
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    
    lng病人ID = rs明细("病人ID")
    str操作员 = UserInfo.姓名
    
    记帐传输_广元 "", 0, "", lng病人ID
    
    gstrSQL = "Select nvl(顺序号,0) as 顺序号,中心 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
    Call OpenRecordset(rsTemp, gstrSysName)
    str就诊编号 = rsTemp!顺序号
'    gstr医保机构编码 = rsTemp!中心
    '医保机构编码, 医院编号, 医保就诊编号， 出院日期，操作员，显示标志
    datCurr = zlDatabase.Currentdate
    initType
    mblnReturn = pcalc(gstr医保机构编码, gstr医院编码, str就诊编号, Format(datCurr, "yyyy-MM-dd"), str操作员, "1", "0", gstrOutPara)
    TrimType
    If mblnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        住院虚拟结算_广元 = ""
        Exit Function
    End If
'间接出口参数:1费用合计,2特殊病种费用,3本次本年帐户支付,4本次历年帐户支付,5累计分段自付,6统筹金支付,7起付段支付,
'             8单位支付,9自费费用,10特检先自付,11特治先自付,12特检费用,13特治费用,14补充医疗保险支付,15本次统筹记入累计,
'             16补充医疗记入累计,17门诊统筹记入累计,18未报销费用,19医保支付,20个人现金支付,21个人帐户余额
    

    '获取个人帐户支付和个人现金支付
    cur个人帐户支付 = CCur(gstrOutPara.out3) + CCur(gstrOutPara.out4)
    cur个人现金支付 = CCur(gstrOutPara.out20)
    cur统筹支付 = CCur(gstrOutPara.out6)
    cur医保支付 = CCur(gstrOutPara.out19)
    cur补充医保 = CCur(gstrOutPara.out14)
    If curCount <> CCur(gstrOutPara.out1) Then
        MsgBox "请注意：医保返回结算金额与当前单据金额不符", vbInformation, gstrSysName
    End If
    住院虚拟结算_广元 = "个人帐户;" & cur个人帐户支付 & ";0" '不允许修改个人帐户
'    If cur个人现金支付 <> 0 Then
'        住院虚拟结算_广元 = 住院虚拟结算_广元 & "|现金;" & cur个人现金支付 & ";0" '不允许修改现金支付
'    End If
    If cur统筹支付 <> 0 Then
        住院虚拟结算_广元 = 住院虚拟结算_广元 & "|医保基金;" & cur统筹支付 & ";0" '不允许修改统筹支付
    End If
    If cur补充医保 <> 0 Then
        住院虚拟结算_广元 = 住院虚拟结算_广元 & "|补充医疗保险;" & cur补充医保 & ";0"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    住院虚拟结算_广元 = ""
End Function

Public Function 出院登记_广元(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    '个人状态的修改
    Dim str就诊编号 As String, rsTemp As New ADODB.Recordset
    Dim bln零费用出院 As Boolean
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
checkCard:
        initType
        mblnReturn = getybjgbm(gstrOutPara)
        TrimType
        If mblnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo checkCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
    
    On Error GoTo errHandle
    '检查该次住院是否没有费用发生
    gstrSQL = "Select nvl(sum(实收金额),0) as 金额  from 病人费用记录 where nvl(附加标志,0)<>9 and 病人ID=" & lng病人ID & " and 主页ID=" & lng主页ID
    Call OpenRecordset(rsTemp, "病人出院")
    If rsTemp.EOF = True Then
        bln零费用出院 = True
    Else
        bln零费用出院 = (rsTemp("金额") = 0)
    End If
    
    If bln零费用出院 = True Then
        gstrSQL = "Select nvl(顺序号,0) as 顺序号 From 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_广元
        Call OpenRecordset(rsTemp, gstrSysName)
        str就诊编号 = rsTemp!顺序号
        initType
        mblnReturn = dall(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
        If mblnReturn = False Then
            出院登记_广元 = False
            Exit Function
        End If
    End If
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    出院登记_广元 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    出院登记_广元 = False
End Function

Public Function 医保设置_广元() As Boolean
    医保设置_广元 = frmSet广元.ShowME(TYPE_广元)
End Function

Private Function Get病人ID(str医保号 As String, str医保中心编码 As String) As String
'功能：通过医保中心号码和医保号求出病人ID
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 病人ID from 保险帐户 where 险类 = '" & TYPE_广元 & _
            "' and 医保号 = '" & str医保号 & "'"
    Call OpenRecordset(rsTmp, gstrSysName)
    If Not rsTmp.BOF Then
        Get病人ID = CStr(rsTmp("病人ID"))
    Else
        Get病人ID = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Get病人ID = ""
End Function

Public Function 记帐传输_广元(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    If str单据号 <> "" Then
        gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and NO='" & str单据号 & "'"
        Call OpenRecordset(rsTemp, gstrSQL)
        If lng病人ID = 0 Then lng病人ID = rsTemp!病人ID
        gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and NO='" & str单据号 & "' order by 主页ID,序号"
    Else
        gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and 病人id=" & lng病人ID & " order by 主页ID,序号"
    End If
    Call OpenRecordset(rsTemp, gstrSQL)
'    While Not rsTemp.EOF
'        gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and id=" & rsTemp!ID
'        Call OpenRecordset(rsTmp, gstrSQL)
    
        记帐传输_广元 = 费用明细传递_广元(0, rsTemp)
        If 记帐传输_广元 = False Then Exit Function
'        rsTemp.MoveNext
'    Wend
End Function
