Attribute VB_Name = "mdl华东"
Option Explicit
Private mcur统筹金额 As Currency, mcur个帐支付 As Currency
Public gcn华东 As New ADODB.Connection, mstrSavePath As String

Public Const MAX_PATH = 260

Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '设置浏览窗口
        .hwndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function


Public Function 医保初始化_华东() As Boolean
'功能：测试是否可以连接到前置服务器上
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim strSql As String, rs华东 As New ADODB.Recordset
    '如果连接已经打开，那就不用再测试
    If gcn华东.State = adStateOpen Then
        医保初始化_华东 = True
        Exit Function
    End If
     
    On Error GoTo errH
    
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        If rsTemp!参数名 = "文件存放位置" Then mstrSavePath = rsTemp!参数值
        rsTemp.MoveNext
    Loop
    If Trim(mstrSavePath) = "" Then
        MsgBox "请到医保参数设置中设置文件存放位置", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error Resume Next
    gcn华东.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DSN=Visual FoxPro Tables;UID=;SourceDB=" & mstrSavePath & ";SourceType=DBF;Exclusive=No;BackgroundFetch=Yes;Collate=Machine;Null=Yes;Deleted=Yes;"""
    gcn华东.CursorLocation = adUseClient
    gcn华东.Open
    
    If Err <> 0 Then
        MsgBox "文件存放位置指定错误。", vbInformation, gstrSysName
        医保初始化_华东 = False
        Exit Function
    End If
    医保初始化_华东 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    医保初始化_华东 = False
End Function

Public Function 医保设置_华东() As Boolean
    医保设置_华东 = frmSet华东.ShowMe(gintInsure)
End Function

Public Function 个人余额_华东(lng病人id As Long) As Currency
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select * From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    个人余额_华东 = Nvl(rsTemp!帐户余额, 0)
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Function 身份标识_华东(Optional bytType As Byte = 0, Optional lng病人id As Long = 0) As String
    '华东医保没提供专门的身分验证接口，通过调取挂号单号来实现验证
    Dim strTemp As String
    strTemp = frmIdentify华东.Identify(bytType, lng病人id)
    Unload frmIdentify华东
    If strTemp = "" Then
        MsgBox "未提取病人信息", vbInformation, gstrSysName
    Else
        身份标识_华东 = strTemp
    End If
End Function

Public Function 门诊虚拟结算_华东(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'因为华东未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, lng病人id As Long, datCurr As Date, strSql As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long
    Dim strCardNO As String
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
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    datCurr = zlDatabase.Currentdate
    lng病人id = rs明细(0)
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    str流水号 = Mid(Format(datCurr, "YYMMDDHHMMSS"), 2, 10) & Format(lng病人id, "0####")
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    
    '生成DBF文件
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
    
    On Error GoTo errHandle
    gcn华东.Execute "Create Table " & mstrSavePath & "\YM" & str流水号 & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    lng序号 = 1
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select A.项目编码,B.ID,B.名称,B.规格,B.计算单位 From 保险支付项目 A,收费细目 B Where B.ID=A.收费细目ID And A.收费细目id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
        
        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
        gcn华东.Execute "Insert Into " & mstrSavePath & "\YM" & str流水号 & " values ('','" & str流水号 & "'," & _
            lng序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
            "''," & rs明细!数量 & "," & rs明细!单价 & "," & rs明细!实收金额 & "," & _
            "0)"
        lng序号 = lng序号 + 1
        rs明细.MoveNext
    Wend
    On Error GoTo errHandle
    '等待返回结算数据
    If frm等待返回华东.waitReturn(mstrSavePath & "\SM" & str流水号) = False Then
        MsgBox "预结算被中止", vbInformation, gstrSysName
        On Error Resume Next
        gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
        Unload frm等待返回华东
        Exit Function
    End If
    Unload frm等待返回华东
    
    '返回结算结果
    strSql = "Select * From " & mstrSavePath & "\SM" & str流水号
    Set rsTemp = gcn华东.Execute(strSql)
    mcur个帐支付 = Val(rsTemp!JkAccR)
    mcur统筹金额 = Val(rsTemp!JkSocialR)
    str结算方式 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
    str结算方式 = str结算方式 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YM" & str流水号
    gcn华东.Execute "Drop Table " & mstrSavePath & "\SM" & str流水号
    门诊虚拟结算_华东 = True
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_华东(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, cur票据总金额 As Currency
    Dim datCurr As Date, lng病人id As Long
    
    On Error GoTo errHandle
    datCurr = zlDatabase.Currentdate
    
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

    门诊结算_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_华东(lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long) As Boolean
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
        门诊结算冲销_华东 = False
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

    门诊结算冲销_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_华东(rs明细 As ADODB.Recordset, lng病人id As Long, str医保号 As String) As String
'因为华东未提供预结算接口，因此所得到的结算数据为医保结算的正式数据，即得到数据时医保已正式结算
    Dim str流水号 As String, datCurr As Date, strSql As String
    Dim rsTemp As New ADODB.Recordset, rsDBF As New ADODB.Recordset, lng序号 As Long
    Dim strCardNO As String
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
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If

    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=" & lng病人id
    Call OpenRecordset(rsTemp, gstrSysName)
    str流水号 = Format(lng病人id, "0######") & "_" & rsTemp(0)

    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    记帐传输_华东 "", 0, "", lng病人id
    '生成DBF文件
'    gcn华东.Execute "Create Table " & mstrSavePath & "\YZ" & str流水号 & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
'        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
'        "SelfJe N(18,4))"
'    strSql = "Select * From " & mstrSavePath & "\YZ" & str流水号
'    lng序号 = 1
'    rs明细.MoveFirst
'    While Not rs明细.EOF
'        gstrSQL = "Select A.项目编码,B.ID,B.名称,B.规格,B.计算单位 From 保险支付项目 A,收费细目 B Where B.ID=A.收费细目ID And A.收费细目id=" & rs明细!收费细目ID
'        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
'
'        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
'        gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('" & strCardNO & "','" & str流水号 & "'," & _
'            lng序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
'            "''," & rs明细!数量 & "," & rs明细!单价 & "," & rs明细!实收金额 & "," & _
'            "0)"
'        lng序号 = lng序号 + 1
'        rs明细.MoveNext
'    Wend
    On Error GoTo errHandle
    
    '等待返回结算数据
    If frm等待返回华东.waitReturn(mstrSavePath & "\SZ" & str流水号) = False Then
        MsgBox "预结算被中止", vbInformation, gstrSysName
'        gcn华东.Execute "Drop Table " & mstrSavePath & "\YZ" & str流水号
        Unload frm等待返回华东
        Exit Function
    End If
    Unload frm等待返回华东
    
    '返回结算结果
    strSql = "Select Sum(JkaccR) As JkaccR,Sum(JkSocialR) As JkSocialR From " & mstrSavePath & "\SZ" & str流水号
    Set rsTemp = gcn华东.Execute(strSql)
    mcur个帐支付 = Val(rsTemp!JkAccR)
    mcur统筹金额 = Val(rsTemp!JkSocialR)
    住院虚拟结算_华东 = "个人帐户;" & Val(rsTemp!JkAccR) & ";0"
    住院虚拟结算_华东 = 住院虚拟结算_华东 & "|统筹记帐;" & Val(rsTemp!JkSocialR) & ";0"
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\SZ" & str流水号
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_华东(lng结帐ID As Long, ByVal lng病人id As Long) As Boolean
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

    住院结算_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_华东(lng结帐ID As Long) As Boolean
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
        住院结算冲销_华东 = False
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

    住院结算冲销_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_华东(lng病人id As Long, lng主页ID As Long) As Boolean
    On Error GoTo errHandle
    '对HIS之中的基础数据进行修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    出院登记_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_华东 = False
End Function

Public Function 入院登记_华东(lng病人id As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    
    On Error GoTo errHandle
     '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    入院登记_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_华东 = False
End Function

Public Function 记帐传输_华东(ByVal str单据号 As String, ByVal int性质 As Integer, str消息 As String, Optional ByVal lng病人id As Long = 0) As Boolean
    Dim rs明细 As New ADODB.Recordset, lng主页ID As Long, rsTemp As New ADODB.Recordset
    Dim str流水号 As String, datCurr As Date, strSql As String, lng序号 As Long
    Dim strCardNO As String
    If str单据号 <> "" Then
        gstrSQL = "Select 病人id From 病人费用记录 Where NO='" & str单据号 & "'"
        Call OpenRecordset(rsTemp, gstrSysName)
        lng病人id = rsTemp(0)
    End If
    gstrSQL = "Select Max(主页ID) From 病案主页 Where 病人id=" & lng病人id
    Call OpenRecordset(rsTemp, gstrSysName)
    lng主页ID = Nvl(rsTemp(0), 1)
    If str单据号 <> "" Then
        gstrSQL = "Select * From 病人费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 记录性质=" & int性质 & " and NO='" & str单据号 & "' order by 主页ID,序号"
    Else
        gstrSQL = "Select * From 病人费用记录 Where 记录状态<>0 And Nvl(是否上传,0)=0 And nvl(附加标志,0)<>9 and 病人id=" & lng病人id & " And 主页id=" & lng主页ID & " And (是否上传 Is Null Or 是否上传=0) order by 主页ID,序号"
    End If
    Call OpenRecordset(rs明细, gstrSQL)
    
    On Error GoTo errHandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用，不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select * From 病人费用记录 Where nvl(附加标志,0)<>9 and 病人id=" & lng病人id & " And 主页id=" & lng主页ID & " And 是否上传=1 order by 主页ID,序号"
    Call OpenRecordset(rsTemp, gstrSysName)
    lng序号 = rsTemp.RecordCount + 1
    
    datCurr = zlDatabase.Currentdate
    gstrSQL = "Select 卡号 From 保险帐户 Where 病人id=" & lng病人id & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, gstrSysName)
    If rsTemp.EOF Then
        MsgBox "没有找到病人信息或医保选择错误", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNO = rsTemp!卡号
    '生成流水号
    str流水号 = Format(lng病人id, "0######") & "_" & lng主页ID
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=" & gintInsure & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    
    '生成DBF文件
    On Error Resume Next
    gcn华东.Execute "Drop Table " & mstrSavePath & "\YZ" & str流水号
    
    On Error GoTo errHandle
    gcn华东.Execute "Create Table " & mstrSavePath & "\YZ" & str流水号 & " (IDNo C(18),CaseNo C(15),OrderNo N(18,4)," & _
        "IntelCode C(14),CName C(70),SubCode C(8),Standard C(20),CUnit C(4),Num N(18,4),Price N(18,4),SumJe N(18,4)," & _
        "SelfJe N(18,4))"
    strSql = "Select * From " & mstrSavePath & "\YZ" & str流水号
    rs明细.MoveFirst
    While Not rs明细.EOF
        gstrSQL = "Select A.项目编码,B.ID,B.名称,B.规格,B.计算单位 From 保险支付项目 A,收费细目 B Where B.ID=A.收费细目ID And A.收费细目id=" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, gstrSysName)             '因为之前检查了是否进行对码，所以读出的记录一定不会空
        
        '卡号、流水号、序号、内码、名称、科目编码、规格、计量单位、数量、单价、金额、自费金额
        gcn华东.Execute "Insert Into " & mstrSavePath & "\YZ" & str流水号 & " values ('','" & str流水号 & "'," & _
            lng序号 & ",'" & Trim(rsTemp!项目编码) & "','" & Trim(rsTemp!名称) & "','" & Trim(rsTemp!规格) & "','" & Trim(rsTemp!计算单位) & "'," & _
            "''," & rs明细!付数 * rs明细!数次 & "," & rs明细!实收金额 / (rs明细!付数 * rs明细!数次) & "," & rs明细!实收金额 & "," & _
            "0)"
        lng序号 = lng序号 + 1
        gstrSQL = "zl_病人记帐记录_上传 ('" & rs明细!ID & "')"
        Call ExecuteProcedure(gstrSysName)
        rs明细.MoveNext
    Wend
    记帐传输_华东 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

