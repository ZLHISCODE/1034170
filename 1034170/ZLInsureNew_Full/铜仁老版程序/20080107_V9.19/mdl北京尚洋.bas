Attribute VB_Name = "mdl北京尚洋"
Option Explicit

Public gcn尚洋 As New ADODB.Connection, gint适用地区_尚洋 As Integer
Private mcur统筹金额 As Currency, mcur个帐支付 As Currency

Public Function 医保初始化_北京尚洋() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs北京尚洋 As New ADODB.Recordset, str参数值 As String
    '如果连接已经打开，那就不用再测试
    If gcn尚洋.State = adStateOpen Then
        医保初始化_北京尚洋 = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "用户名"
                strUser = str参数值
            Case "服务器"
                strServer = str参数值
            Case "用户密码"
                strPass = str参数值
            Case "适用地区"
                gint适用地区_尚洋 = Val(str参数值)
            Case "统筹区号"
                gstr医保机构编码 = str参数值
        End Select
        rsTemp.MoveNext
    Loop
    If strUser = "" Or strServer = "" Or strPass = "" Then
        MsgBox "参数设置不完整,请到医保参数设置中重新设置", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
    If gint适用地区_尚洋 = 1 Then
'        gcn尚洋.ConnectionString = "Provider=Sybase.ASEOLEDBProvider.2;口令=" & strPass & ";持续安全性信息=True;用户 ID=" & strUser & ";数据源=" & strServer
        gcn尚洋.ConnectionString = "Provider=MSDASQL.1;Password=" & strPass & ";Persist Security Info=True;User ID=" & strUser & ";Data Source=" & strServer
    Else
        gcn尚洋.ConnectionString = "Provider=MSDAORA.1;Password=" & strPass & ";User ID=" & strUser & ";Data Source=" & strServer & ";Persist Security Info=True"
    End If
    gcn尚洋.CursorLocation = adUseClient
    gcn尚洋.Open
    
    If Err <> 0 Then
        MsgBox "连接前置服务器发生错误。", vbInformation, gstrSysName
        医保初始化_北京尚洋 = False
        Exit Function
    End If
    医保初始化_北京尚洋 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    医保初始化_北京尚洋 = False
End Function

Public Function 医保设置_北京尚洋() As Boolean
    医保设置_北京尚洋 = frmSet北京尚洋.参数设置()
End Function

Public Function 个人余额_北京尚洋(lng病人id As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    '因为能提取个帐余额,因此赋予足够大的数,具体支付金额由医保决定
    个人余额_北京尚洋 = Nvl(rsTemp!帐户余额, 1000000)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function 身份标识_北京尚洋(Optional bytType As Byte = 0, Optional lng病人id As Long = 0) As String
    '北京尚洋医保没提供专门的身分验证接口
    Dim strTemp As String
    strTemp = frmIdentify北京尚洋.Identify(bytType, lng病人id)
    Unload frmIdentify北京尚洋
    If strTemp = "" Then
        MsgBox "未提取病人信息", vbInformation, gstrSysName
    Else
        身份标识_北京尚洋 = strTemp
    End If
End Function
'
'Public Function 门诊虚拟结算_北京尚洋(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
''因为北京尚洋未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
'    Dim str流水号 As String, lng病人ID As Long, datCurr As Date, strSql As String, strTemp As String
'    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long, str科室 As String
'    Dim strCardNO As String, str收据项目 As String, str科室项目 As String, str会计项目 As String
''    病人ID         adBigInt, 19, adFldIsNullable
''    收费类别       adVarChar, 2, adFldIsNullable
''    收据费目       adVarChar, 20, adFldIsNullable
''    计算单位       adVarChar, 6, adFldIsNullable
''    开单人         adVarChar, 20, adFldIsNullable
''    收费细目ID     adBigInt, 19, adFldIsNullable
''    数量           adSingle, 15, adFldIsNullable
''    单价           adSingle, 15, adFldIsNullable
''    实收金额       adSingle, 15, adFldIsNullable
''    统筹金额       adSingle, 15, adFldIsNullable
''    保险支付大类ID adBigInt, 19, adFldIsNullable
''    是否医保       adBigInt, 19, adFldIsNullable
''    摘要           adVarChar, 200, adFldIsNullable
''    是否急诊       adBigInt, 19, adFldIsNullable
''    str结算方式  "报销方式;金额;是否允许修改|...."
'    On Error GoTo errHandle
'    If rs明细.RecordCount = 0 Then
'        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    datCurr = zlDatabase.Currentdate
'    lng病人ID = rs明细(0)
'    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人ID & " And 险类=" & gintInsure
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If rsTemp.EOF Then
'        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
'        Exit Function
'    End If
'    strCardNO = rsTemp!卡号
'    '生成流水号
'    str流水号 = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng病人ID, "0######"), 35)
'
'    '判断是否有医保编码未对应
'    Do Until rs明细.EOF
'        gstrSQL = "select A.项目编码,B.名称,B.说明 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        If IsNull(rsTemp!项目编码) Then
'            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
'            Exit Function
'        End If
'        If IsNull(rsTemp!说明) Then
'            MsgBox "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别", vbInformation, gstrSysName
'            Exit Function
'        ElseIf Len(rsTemp!说明) < 2 Then
'            MsgBox "不能确定项目<" & rsTemp!名称 & ">的科室核算类别", vbInformation, gstrSysName
'            Exit Function
'        End If
'        strTemp = rsTemp!项目编码
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstr医保机构编码 & "' And Item_Code='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "在中间数据中未找到编码为[" & UCase(strTemp) & "]的项目，请核查", vbInformation, gstrSysName
'            Exit Function
'        End If
'        rs明细.MoveNext
'    Loop
'
'    '生成DBF文件
'    lng序号 = 1
'    rs明细.MoveFirst
'    While Not rs明细.EOF
'        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)
'        strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
'        str收据项目 = Left(strTemp, 1)
'        str科室类别 = Mid(strTemp, 2, 1)
'        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
'            str会计类别 = "A"       '药品
'        Else
'            str会计类别 = "B"       '医疗
'        End If
'
'        gstrSQL = "Select 项目编码 From 保险支付项目 Where 险类=" & gintInsure & " And 收费细目id=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
'
'        strSql = "Select * From PARA_CAPTURE_ITEM Where Areaid='" & gstr医保机构编码 & "' And Item_Code='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
''    VISIT_NUMBER                char(18)        not null,   //处方号码
''    ITEM_NO                     numeric(6, 0)   not null,   //同一处方中项目序号
''    ITEM_CLASS                  char(1)         not null,   //收费项目类别:如A西药
''    ITEM_CODE                   char(12)        not null,   //项目编码
''    ITEM_NAME                   char(40)        not null,   //项目名称
''    SPEC                        varchar(50)     not null,   //规格
''    PRICE_UNIT                  char(8)         not null,   //计价单位
''    PRICE                       numeric(9, 4)   not null,   //单价
''    QUANTITY                    numeric(6, 2)   not null,   //数量
''    COST                        numeric(8, 2)   not null,   //金额
''    RECEIPT_CLASS               char(1)         not null,   //收据项目分类
''    COLLATE_RELATION            char(12)        null,       //与医保中心对应关系
''    OPERATOR                    char(15)        null,       //经办人
''    OPERATE_TIME                datetime        null,       //经办日期
''    CLINIC_FLAG                 numeric(1, 0)   not null,   //门诊/住院标志
''    EXE_DEPT                    char(20)        null,       //执行科室
''    APP_DOCTOR                  char(30)        null,       //开方医生
''    APP_DEPT                    char(20)        null,       //开单科室
''    TAKE_MEDICINE_FLAG          char(8)         not null,   //出院带药标志
''    ITEM_NO_DEPT_STAT           char(2)         null,       //科室核算项目类别
''    ITEM_NO_ACCOUNTANT_ITEM char(2)         null,       //会计核算项目类别
''    constraint PK_SICK_PRICE_ITEM PRIMARY KEY CLUSTERED (VISIT_NUMBER, ITEM_NO)
''A 西药，B 成药，C 草药，D 治疗，E 检查，F 放射，G 化验，H 手术费，I 输血费，J 输氧费，K CT。ECT，L 其它，M B超，N 心电图，O 脑电图，P 胃镜，Q 喉镜
'
'        gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM values ('" & str流水号 & "'," & lng序号 & ",'" & _
'            Trim(rsTemp!ITEM_TYPE) & "','" & Trim(rsTemp!ITEM_CODE) & "','" & ToVarchar(Trim(rsTemp!ITEM_NAME), 40) & "','" & _
'            ToVarchar(Trim(rsTemp!ITEM_SPEC), 50) & "','" & ToVarchar(Trim(rsTemp!PRICE_UNIT), 8) & "','" & _
'            Trim(rsTemp!CUnit) & "'," & rs明细!单价 & "," & rs明细!数量 & "," & rs明细!实收金额 & ",'" & _
'            str收据项目 & "','" & trim(rstemp!ITEM_CODE) & "','" & userinfo.姓名 & "',to_Date('" & _
'            format(zldatabase.Currentdate,"yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS'),0,'" & _
'
'        lng序号 = lng序号 + 1
'        rs明细.MoveNext
'    Wend
'    On Error GoTo errHandle
'
'    '等待返回结算数据
'    If frm等待返回北京尚洋.waitReturn(mstrSavePath & "\SM" & str流水号) = False Then
'        MsgBox "预结算被中止", vbInformation, gstrSysName
'        Unload frm等待返回北京尚洋
'        Exit Function
'    End If
'    Unload frm等待返回北京尚洋
'
'    '返回结算结果
'    strSql = "Select * From " & mstrSavePath & "\SM" & str流水号
'    Set rsTemp = gcn尚洋.Execute(strSql)
'    mcur个帐支付 = Val(rsTemp!JkAccR)
'    mcur统筹金额 = Val(rsTemp!JkSocialR)
'    str结算方式 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
'    str结算方式 = str结算方式 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
'    门诊虚拟结算_北京尚洋 = True
'    Exit Function
'
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

Public Function 门诊结算_北京尚洋(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim str流水号 As String, lng病人id As Long, datCurr As Date, strSql As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng序号 As Long, str执行部门 As String, str开单部门 As String
    Dim strCardNO As String, str收据项目 As String, str科室类别 As String, str会计类别 As String
    Dim str出院带药 As String, cur基本统筹 As Currency, cur大病统筹 As Currency, rs明细 As New ADODB.Recordset
    Dim cur公务员补助 As Currency, cur补充医疗 As Currency, str结算方式 As String
    Dim strTempID As String
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select * From 病人费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 And 结帐id=" & lng结帐ID
    Call OpenRecordset(rs明细, gstrSysName)
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng病人id = rs明细!病人ID
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
'    str流水号 = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng病人ID, "0######"), 35)
    str流水号 = rs明细!NO
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称,B.说明,B.类别 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
'        If IsNull(rsTemp!项目编码) Then
'            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
'            Exit Function
'        End If
        If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
        
        Else
            If IsNull(rsTemp!说明) Then
                MsgBox "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别", vbInformation, gstrSysName
                Exit Function
            ElseIf Len(rsTemp!说明) < 2 Then
                MsgBox "不能确定项目<" & rsTemp!名称 & ">的科室核算类别", vbInformation, gstrSysName
                Exit Function
            End If
        End If
'        strTemp = rsTemp!项目编码
'        strSql = "Select * From PARA_CAPTURE_ITEM Where AREAID='" & gstr医保机构编码 & "' And ITEM_CODE='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "在中间数据中未找到编码为[" & UCase(strTemp) & "]的项目，请核查", vbInformation, gstrSysName
'            Exit Function
'        End If
        rs明细.MoveNext
    Loop
    
    '传费用明细
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select * From 部门表 Where ID=" & rs明细!执行部门id
        Call OpenRecordset(rsTemp, gstrSysName)
        str执行部门 = rsTemp!名称
        strSql = "Select * From 部门表 Where ID=" & rs明细!开单部门ID
        str开单部门 = rsTemp!名称
        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull(rsTemp!说明) Then
            If rsTemp!类别 = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!类别 = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!类别 = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
        End If
        str收据项目 = Left(strTemp, 1)
        str科室类别 = Mid(strTemp, 2, 1)
        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
            str会计类别 = "A"       '药品
        Else
            str会计类别 = "B"       '医疗
        End If
        Select Case rsTemp!类别
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "I", "Z"
                strItemType = "Z"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
        End Select
        str出院带药 = "在院用药"           '出院带药标志（取值还不清楚，有待询问）
        gstrSQL = "Select 项目编码,收费细目ID From 保险支付项目 Where 险类=" & gintInsure & " And 收费细目id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
        If rsTemp.EOF Then
            strTempID = ""
        ElseIf IsNull(rsTemp!项目编码) Then
            strTempID = ""
        Else
            strTempID = rsTemp!项目编码
        End If
        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        strItemName = rsTemp!名称
        strItemCode = rsTemp!编码
        strPriceUnit = Nvl(rsTemp!计算单位)
        strItemSpec = Nvl(rsTemp!规格)
        
        '向中间表写入数据,住院/门诊标志有待询问
        gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM values ('" & str流水号 & "'," & lng序号 & ",'" & _
            strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
            ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
            rs明细!实收金额 / (rs明细!付数 * rs明细!数次) & "," & rs明细!付数 * rs明细!数次 & "," & rs明细!实收金额 & ",'" & _
            str收据项目 & "','" & strTempID & "','" & rs明细!操作员姓名 & "','" & _
            Format(rs明细!登记时间, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
            str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str出院带药 & "','" & _
            str科室类别 & "','" & str会计类别 & "')"
        
        '向中间表写入数据
'        gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM (VISIT_NUMBER,ITEM_NO,ITEM_CLASS,ITEM_CODE,ITEM_NAME,SPEC,PRICE_UNIT,PRICE,QUANTITY,COST,RECEIPT_CLASS,COLLATE_RELATION,OPERATOR,OPERATE_TIME,CLINIC_FLAG,EXE_DEPT,APP_DOCTOR,APP_DEPT,TAKE_MEDICINE_FLAG,ITEM_NO_DEPT_STAT,ITEM_NO_ACCOUNTANT_ITEM) values ('" & str流水号 & "'," & rs明细!序号 & ",'" & _
            strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
            ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
            rs明细!实收金额 / (rs明细!付数 * rs明细!数次) & "," & rs明细!付数 * rs明细!数次 & "," & rs明细!实收金额 & ",'" & _
            str收据项目 & "','" & strTempID & "','" & rs明细!操作员姓名 & "','" & _
            Format(rs明细!登记时间, "yyyy-MM-dd HH:mm:ss") & "',0,'" & _
            str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str出院带药 & "','" & _
            str科室类别 & "','" & str会计类别 & "')"
        lng序号 = lng序号 + 1
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    
    '等待返回结算数据
    strTemp = frm等待返回北京尚洋.waitReturn(str流水号)
    If strTemp = "" Then
        MsgBox "结算过程被中止", vbInformation, gstrSysName
        gcn尚洋.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str流水号 & "'"
        Unload frm等待返回北京尚洋
        Exit Function
    End If
    Unload frm等待返回北京尚洋
    
    '返回结算结果
    strSql = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn尚洋.Execute(strSql)
    
    cur个人帐户 = rsTemp!PAY_SIDE2
    cur基本统筹 = rsTemp!PAY_SIDE3
    cur大病统筹 = rsTemp!PAY_SIDE4
    cur补充医疗 = rsTemp!PAY_SIDE5
    cur公务员补助 = rsTemp!PAY_SIDE6
    '写结算结果
    If cur个人帐户 <> 0 Then
        str结算方式 = str结算方式 & "||个人帐户|" & cur个人帐户
    End If
    If cur基本统筹 <> 0 Then
        str结算方式 = str结算方式 & "||基本基金|" & cur基本统筹
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "||大病基金|" & cur大病统筹
    End If
    If cur补充医疗 <> 0 Then
        str结算方式 = str结算方式 & "||补充基金|" & cur补充医疗
    End If
    If cur公务员补助 <> 0 Then
        str结算方式 = str结算方式 & "||公务员津贴|" & cur公务员补助
    End If
    
    '如果存在
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        gstrSQL = "zl_病人结算记录_Update(" & lng结帐ID & ",'" & str结算方式 & "',0)"
        Call ExecuteProcedure("更新预交记录")
    End If
    frm结算信息.ShowME (lng结帐ID)
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng病人id = 0 Then lng病人id = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '帐户年度信息
    Call Get帐户信息(lng病人id, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人id & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    门诊结算_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_北京尚洋(lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng病人id = 0 Then lng病人id = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        门诊结算冲销_北京尚洋 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(lng病人id, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人id & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    门诊结算冲销_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_北京尚洋(rsDetail As ADODB.Recordset, lng病人id As Long, str医保号 As String) As String
'因为北京尚洋未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, datCurr As Date, strSql As String, strTemp As String
    Dim rsTemp As New ADODB.Recordset, lng序号 As Long, str执行部门 As String, str开单部门 As String
    Dim strCardNO As String, str收据项目 As String, str科室类别 As String, str会计类别 As String
    Dim str出院带药 As String, cur基本统筹 As Currency, cur大病统筹 As Currency, rs明细 As New ADODB.Recordset
    Dim cur公务员补助 As Currency, cur补充医疗 As Currency, str结算方式 As String, cur个人帐户 As Currency
    Dim strTempID As String
    Dim strItemType As String, strItemCode As String, strItemName As String, strItemSpec As String, strPriceUnit As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select Max(主页ID) From 病人费用记录 Where 病人id=" & lng病人id
    Call OpenRecordset(rsTemp, gstrSysName)
    gstrSQL = "Select * From 病人费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 And 病人id=" & lng病人id & " And 主页id=" & rsTemp(0)
    Call OpenRecordset(rs明细, gstrSysName)
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng病人id = rs明细!病人ID
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    str流水号 = toHex(Format(datCurr, "YYMMDDHHMMSS") & Format(lng病人id, "0######"), 35)
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称,B.说明,B.编码,B.类别 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
'        If IsNull(rsTemp!项目编码) Then
'            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
'            Exit Function
'        End If
        If rsTemp!类别 = "5" Or rsTemp!类别 = "6" Or rsTemp!类别 = "7" Then
        
        Else
            If IsNull(rsTemp!说明) Then
                MsgBox "不能确定项目<" & rsTemp!名称 & ">的收据项目类别和科室核算类别", vbInformation, gstrSysName
                Exit Function
            ElseIf Len(rsTemp!说明) < 2 Then
                MsgBox "不能确定项目<" & rsTemp!名称 & ">的科室核算类别", vbInformation, gstrSysName
                Exit Function
            End If
        End If
'        strTemp = rsTemp!项目编码
'        strSql = "Select * From PARA_CAPTURE_ITEM Where AREAID='" & gstr医保机构编码 & "' And ITEM_CODE='" & UCase(rsTemp!项目编码) & "'"
'        Set rsTemp = gcn尚洋.Execute(strSql)
'        If rsTemp.EOF Then
'            MsgBox "在中间数据中未找到编码为[" & UCase(strTemp) & "]的项目，请核查", vbInformation, gstrSysName
'            Exit Function
'        End If
        rs明细.MoveNext
    Loop
    
    '传费用明细
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select * From 部门表 Where ID=" & rs明细!执行部门id
        Call OpenRecordset(rsTemp, gstrSysName)
        str执行部门 = rsTemp!名称
        strSql = "Select * From 部门表 Where ID=" & rs明细!开单部门ID
        str开单部门 = rsTemp!名称
        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull(rsTemp!说明) Then
            If rsTemp!类别 = "5" Then
                strTemp = "AA"
            ElseIf rsTemp!类别 = "6" Then
                strTemp = "BB"
            ElseIf rsTemp!类别 = "7" Then
                strTemp = "CC"
            End If
        Else
            strTemp = rsTemp!说明      '说明不能为空，其中第一位存放收据项目类别，第二位存放科室核算类别
        End If
        str收据项目 = Left(strTemp, 1)
        str科室类别 = Mid(strTemp, 2, 1)
        
        If rsTemp!类别 = 5 Or rsTemp!类别 = 6 Or rsTemp!类别 = 7 Then
            str会计类别 = "A"       '药品
        Else
            str会计类别 = "B"       '医疗
        End If
        Select Case rsTemp!类别
            Case "5"
                strItemType = "A"
            Case "6"
                strItemType = "C"
            Case "7"
                strItemType = "C"
            Case "C"
                strItemType = "D"
            Case "D"
                strItemType = "E"
            Case "E", "L"
                strItemType = "F"
            Case "F"
                strItemType = "G"
            Case "G"
                strItemType = "H"
            Case "H"
                strItemType = "I"
            Case "I", "Z"
                strItemType = "Z"
            Case "J"
                strItemType = "J"
            Case "K"
                strItemType = "L"
            Case "M"
                strItemType = "K"
        End Select
        gstrSQL = "Select 扣率 From 药品收发记录 Where 费用ID=" & rs明细!ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            str出院带药 = "在院用药"
        Else
            If Mid(CStr(Nvl(rsTemp(0), 0)), 2, 1) = "3" Then
                str出院带药 = "出院带药"
            Else
                str出院带药 = "在院用药"
            End If
        End If
        
        gstrSQL = "Select 项目编码,收费细目ID From 保险支付项目 Where 险类=" & gintInsure & " And 收费细目id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
        If rsTemp.EOF Then
            strTempID = ""
        ElseIf IsNull(rsTemp!项目编码) Then
            strTempID = ""
        Else
            strTempID = rsTemp!项目编码
        End If
        gstrSQL = "Select * From 收费细目 Where ID=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        strItemName = rsTemp!名称
        strItemCode = rsTemp!编码
        strPriceUnit = Nvl(rsTemp!计算单位)
        strItemSpec = Nvl(rsTemp!规格)
        
        '向中间表写入数据,住院/门诊标志有待询问
        gcn尚洋.Execute "Insert Into SICK_PRICE_ITEM values ('" & str流水号 & "'," & lng序号 & ",'" & _
            strItemType & "','" & strItemCode & "','" & ToVarchar(strItemName, 40) & "','" & _
            ToVarchar(strItemSpec, 50) & "','" & ToVarchar(strPriceUnit, 8) & "'," & _
            rs明细!实收金额 / (rs明细!付数 * rs明细!数次) & "," & rs明细!付数 * rs明细!数次 & "," & rs明细!实收金额 & ",'" & _
            str收据项目 & "','" & strTempID & "','" & rs明细!操作员姓名 & "','" & _
            Format(rs明细!发生时间, "yyyy-MM-dd HH:mm:ss") & "',1,'" & _
            str执行部门 & "','" & rs明细!开单人 & "','" & str开单部门 & "','" & str出院带药 & "','" & _
            str科室类别 & "','" & str会计类别 & "')"
        lng序号 = lng序号 + 1
        
'        gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
'        Call ExecuteProcedure(gstrSysName)
        
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    
    '等待返回结算数据
    Screen.MousePointer = 0
    strTemp = frm等待返回北京尚洋.waitReturn(str流水号)
    If strTemp = "" Then
        MsgBox "结算过程被中止", vbInformation, gstrSysName
        gcn尚洋.Execute "Delete From SICK_PRICE_ITEM Where VISIT_NUMBER='" & str流水号 & "'"
        Unload frm等待返回北京尚洋
        Exit Function
    End If
    Unload frm等待返回北京尚洋
    
    '返回结算结果
    strSql = "Select * From MED_RECEIPT_RECORD_MASTER Where CHARGE_NUMBER='" & strTemp & "'"
    Set rsTemp = gcn尚洋.Execute(strSql)
    
    mcur个帐支付 = rsTemp!PAY_SIDE2
    mcur统筹金额 = rsTemp!PAY_SIDE3 + rsTemp!PAY_SIDE4 + rsTemp!PAY_SIDE5 + rsTemp!PAY_SIDE6

    cur个人帐户 = rsTemp!PAY_SIDE2
    cur基本统筹 = rsTemp!PAY_SIDE3
    cur大病统筹 = rsTemp!PAY_SIDE4
    cur补充医疗 = rsTemp!PAY_SIDE5
    cur公务员补助 = rsTemp!PAY_SIDE6
    '写结算结果
    If cur个人帐户 <> 0 Then
        str结算方式 = str结算方式 & "|个人帐户;" & cur个人帐户 & ";0"
    End If
    If cur基本统筹 <> 0 Then
        str结算方式 = str结算方式 & "|基本基金;" & cur基本统筹 & ";0"
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "|大病基金;" & cur大病统筹 & ";0"
    End If
    If cur补充医疗 <> 0 Then
        str结算方式 = str结算方式 & "|补充基金;" & cur补充医疗 & ";0"
    End If
    If cur公务员补助 <> 0 Then
        str结算方式 = str结算方式 & "|公务员津贴;" & cur公务员补助 & ";0"
    End If
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    住院虚拟结算_北京尚洋 = str结算方式
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    住院虚拟结算_北京尚洋 = ""
End Function

Public Function 住院结算_北京尚洋(lng结帐ID As Long, ByVal lng病人id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim datCurr As Date
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '帐户年度信息
    Call Get帐户信息(lng病人id, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人id & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 + mcur个帐支付 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 + mcur统筹金额 & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 & ",0,0," & _
        "0," & mcur统筹金额 & ",0,0," & mcur个帐支付 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    住院结算_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_北京尚洋(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng冲销ID As Long, str流水号 As String, str就诊编号 As String
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, sngArrInfo(20) As Single
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, lng病人id As Long
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency, lngErr As Long
    Dim datCurr As Date, strRecCode As String, strBillCode As String
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
    gstrSQL = "Select 病人ID,结帐金额 From 病人费用记录 Where nvl(附加标志,0)<>9 and 结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    Do Until rsTemp.EOF
        If lng病人id = 0 Then lng病人id = rsTemp("病人ID")
        
        cur票据总金额 = cur票据总金额 + rsTemp("结帐金额")
        rsTemp.MoveNext
    Loop
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B" & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    
    lng冲销ID = rsTemp("结帐ID")
    
    gstrSQL = "select * from 保险结算记录 where 性质=2 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        住院结算冲销_北京尚洋 = False
        Exit Function
    End If
    
    '帐户年度信息
    Call Get帐户信息(lng病人id, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人id & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 - Nvl(rsTemp("进入统筹金额"), 0) & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(2," & lng冲销ID & "," & gintInsure & "," & lng病人id & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 - Nvl(rsTemp("个人帐户支付"), 0) & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 - Nvl(rsTemp("统筹报销金额"), 0) & "," & int住院次数累计 & ",0,0,0," & cur票据总金额 * -1 & ",0,0," & _
        Nvl(rsTemp("进入统筹金额"), 0) * -1 & "," & Nvl(rsTemp("统筹报销金额"), 0) * -1 & ",0," & Nvl(rsTemp("超限自付金额"), 0) & "," & _
        Nvl(rsTemp("个人帐户支付"), 0) * -1 & ",Null,Null,Null,Null)"
    Call ExecuteProcedure(gstrSysName)

    住院结算冲销_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_北京尚洋(lng病人id As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    出院登记_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_北京尚洋 = False
End Function

Public Function 入院登记_北京尚洋(lng病人id As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHandle
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    入院登记_北京尚洋 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_北京尚洋 = False
End Function

Private Function toHex(ByVal dblNum As Double, Optional ByVal dblKey As Double = 16) As String
    Dim dblTemp As Double, dblMod As Double, strTemp As String
    dblTemp = dblNum
    Do
        dblMod = dblTemp - Int(dblTemp / dblKey) * dblKey
        dblTemp = Int(dblTemp / dblKey)
        If dblMod >= 10 Then
            strTemp = Chr(dblMod + 55) & strTemp
        Else
            strTemp = dblMod & strTemp
        End If
    Loop While dblTemp >= dblKey
    dblMod = dblTemp
    If dblMod >= 10 Then
        strTemp = Chr(dblMod + 55) & strTemp
    Else
        strTemp = dblMod & strTemp
    End If
    toHex = strTemp
End Function

Public Sub WriteInfo(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    strFileName = "C:\信息" & Format(Date, "MMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub


