Attribute VB_Name = "mdl南京市"
Option Explicit
Private mstrPatID As String
Private mobjSystem As New FileSystemObject
Private mobjStream As TextStream
Private mcur个帐余额 As Currency
Public gstr正确姓名 As String

Private Type patInfo_南京市
    就诊时间 As String
    病人姓名 As String
    医生编码 As String
    医生姓名 As String
    病种编码 As String
    病种名称 As String
    医保就诊科室码 As String
    医保就诊科室名 As String
    操作人编码 As String
End Type
Public gPatInfo_南京市 As patInfo_南京市

Private Type detailFee_南京市
    住院序号 As String
    病人姓名 As String
    标志 As String
    费用发生时间 As String
    医院编码 As String
    医院自编码  As String
    医保编码 As String
    名称 As String
    剂量单位 As String
    单价 As Double
    数量 As Double
    操作人编码 As String
    产地 As String
    产地特征 As String
    规格 As String
End Type
Private mDetailFee_南京市 As detailFee_南京市

Private Type feeBalance_南京市
    住院序号 As String
    医保卡号 As String
    费用发生时间 As String
    门诊费用合计 As Double
    药费合计 As Double
    治疗项目合计 As Double
    自理费用 As Double
    医保范围费用 As Double
    个人帐户支付 As Double
    统筹支付 As Double
    大病支付 As Double
    个人自付 As Double
    期初个人帐户 As Double
    期末个人帐户 As Double
    操作员编码 As String
    单据号 As String
End Type
Public mFeeBalance As feeBalance_南京市

Public Function 医保初始化_南京市() As Boolean
     医保初始化_南京市 = True
End Function

Public Function 身份标识_南京市(Optional bytType As Byte, Optional lng病人id As Long) As String
    
    On Error GoTo errorhandle
    If bytType = 0 Then
        身份标识_南京市 = frmIdentify南京市.Identify(bytType)
    Else
        身份标识_南京市 = frm数据交换.getFeeBalance(bytType)
        Unload frm数据交换
    End If
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_南京市(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '字段：开单人,病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保

    Dim rsTemp As New ADODB.Recordset, curCount As Currency
    Dim strFile As String, strWrite As String
    Dim strTemp As String
    
    '删除可能存在的前次结算信息文件
    On Error Resume Next
    Call Kill("C:\NJYB\MZJSHZ.TXT")
    
    On Error GoTo errorhandle
    If rs明细.RecordCount = 0 Then
        MsgBox "没有病人费用记录，不能进行结算", vbInformation, gstrSysName
        Exit Function
    End If
    curCount = 0
    While Not rs明细.EOF
        curCount = curCount + rs明细!实收金额
        rs明细.MoveNext
    Wend
    rs明细.MoveFirst
    
    '取出病人信息所需内容
    mstrPatID = rs明细!病人ID
    With gPatInfo_南京市
        .就诊时间 = Format(zlDatabase.Currentdate, "yyyyMMddHHmmss")        '得到就诊时间
        .医生姓名 = Nvl(rs明细!开单人)                                               '得到医生姓名
    End With
    
    If Trim(gPatInfo_南京市.医生姓名) = "" Then
        MsgBox "医保病人收费必须输入医生", vbInformation, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "select A.电子邮件 as 医生编码,C.简码 as 医生科室编码,C.名称 as 医生科室名称 from 人员表 A,部门人员 B,部门表 C,临床部门 D " & _
              "where A.id=B.人员id and B.部门id = C.id and C.id=D.部门id and B.缺省=1 and  A.姓名='" & rs明细!开单人 & "'"
    Call OpenRecordset(rsTemp, "医生编码")
    If rsTemp.EOF Then
        MsgBox "未对应科室的诊疗科目编码,请先正确对应", vbInformation, gstrSysName
    End If
    
    With gPatInfo_南京市
        .医生编码 = rsTemp!医生编码                                               '取得医生编码
        .医保就诊科室码 = rsTemp!医生科室编码
        .医保就诊科室名 = rsTemp!医生科室名称
        .操作人编码 = UserInfo.编号
    End With
    '写入医保病人信息文件
    strFile = "C:\NJYB\MZJZXX.TXT"
    strWrite = gPatInfo_南京市.就诊时间 & fillSpa(gPatInfo_南京市.病人姓名, 12) & _
               fillSpa(gPatInfo_南京市.医生编码, 10) & fillSpa(gPatInfo_南京市.医生姓名, 8) & _
               fillSpa(gPatInfo_南京市.病种编码, 4) & fillSpa(gPatInfo_南京市.病种名称, 40) & _
               fillSpa(gPatInfo_南京市.医保就诊科室码, 4) & fillSpa(gPatInfo_南京市.医保就诊科室名, 30) & _
               fillSpa(gPatInfo_南京市.操作人编码, 10)
    Call writeTxtFile(strFile, strWrite)
    
    '取出明细费用所需内容
    gstrSQL = "select 医院编码 from 保险类别 where 序号=" & TYPE_南京市
    Call OpenRecordset(rsTemp, "医院编码")
    If rsTemp.EOF Then
        MsgBox "医院编码未设置,请先设置医院编码", vbInformation, gstrSysName
        Exit Function
    End If
    With mDetailFee_南京市
        .病人姓名 = gPatInfo_南京市.病人姓名
        .费用发生时间 = gPatInfo_南京市.就诊时间
        .医院编码 = rsTemp!医院编码
        .操作人编码 = gPatInfo_南京市.操作人编码
    End With
    
    '判断是否有医保编码未对应
    Do Until rs明细.EOF
        gstrSQL = "select A.项目编码,B.名称 from (select * from 保险支付项目 where 险类=" & TYPE_南京市 & ") A, 收费细目 B where A.收费细目id(+)=B.id and B.id = " & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, "医保项目")
        If IsNull((rsTemp!项目编码)) Then
            MsgBox "<" & rsTemp!名称 & ">未对应医保编码,请先进行对码", vbInformation, gstrSysName
            Exit Function
        End If
        rs明细.MoveNext
    Loop
    
    strFile = "C:\NJYB\MZCFSJ.TXT"
    Call writeTxtFile(strFile, "")
    rs明细.MoveFirst
    Do Until rs明细.EOF
        gstrSQL = "select decode(A.类别,'5',0,'6',0,'7',0,1) 标志,A.名称,C.项目编码,A.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格" & _
                  " from 收费细目 A,药品目录 B,保险支付项目 C where A.id = C.收费细目id and A.id=B.药品id(+) and A.id =" & rs明细!收费细目ID
        Call OpenRecordset(rsTemp, "收细明细")
        With mDetailFee_南京市
            .标志 = rsTemp!标志
            .名称 = rsTemp!名称
            .医保编码 = rsTemp!项目编码
            .剂量单位 = zlCommFun.Nvl(rsTemp!计算单位)
            .单价 = rs明细!单价
            .数量 = rs明细!数量
            .产地 = Nvl(rsTemp!产地)
            .产地特征 = Nvl(rsTemp!产地特征)
            .规格 = Nvl(rsTemp!规格)
        End With
        strWrite = fillSpa(mDetailFee_南京市.病人姓名, 12) & mDetailFee_南京市.标志 & _
                 mDetailFee_南京市.费用发生时间 & _
                 fillSpa(mDetailFee_南京市.医保编码, 40) & fillSpa(mDetailFee_南京市.名称, 40) & _
                 fillSpa(mDetailFee_南京市.剂量单位, 10) & Lpad(mDetailFee_南京市.单价, 10) & _
                 Lpad(Format(mDetailFee_南京市.数量, "#0.00"), 10) & fillSpa(mDetailFee_南京市.操作人编码, 10) & _
                 fillSpa(mDetailFee_南京市.产地, 20) & fillSpa(mDetailFee_南京市.产地特征, 1) & fillSpa(mDetailFee_南京市.规格, 40)
        Call writeTxtFile(strFile, strWrite, False)
        rs明细.MoveNext
    Loop
    Call writeTxtFile(strFile, "", False)
    
    '读出医保结算结果
    strTemp = frm数据交换.getFeeBalance
    On Error Resume Next
    Unload frm数据交换
    On Error GoTo errorhandle
    If strTemp = "" Then
        MsgBox "读取医保结算文件过程被中止,无法完成预结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    '取出信息为门诊结算做准备
    With mFeeBalance
        .医保卡号 = Val(analyseStr(strTemp, 1, 20))
        .门诊费用合计 = Val(analyseStr(strTemp, 35, 12))
        .自理费用 = Val(analyseStr(strTemp, 67, 10))
        .医保范围费用 = Val(analyseStr(strTemp, 77, 10))
        .个人帐户支付 = Val(analyseStr(strTemp, 87, 10))
        .统筹支付 = Val(analyseStr(strTemp, 97, 10))
        .大病支付 = Val(analyseStr(strTemp, 107, 10))
        .个人自付 = Val(analyseStr(strTemp, 117, 10))
        .单据号 = Val(analyseStr(strTemp, 147, 20))
    End With
    If curCount <> CCur(mFeeBalance.门诊费用合计) Then
        MsgBox "请注意：医保返回费用合计与医院结算费用合计不等" & vbCrLf & _
            "医院：" & curCount & Space(10) & "医保：" & mFeeBalance.门诊费用合计
    End If
    mcur个帐余额 = Val(analyseStr(strTemp, 127, 10))
    
    gstrSQL = "zl_保险帐户_更新信息(" & mstrPatID & "," & TYPE_南京市 & ",'帐户余额','" & mcur个帐余额 & "')"
    Call ExecuteProcedure(gstrSysName)
    
    str结算方式 = "个人帐户;" & mFeeBalance.个人帐户支付 & ";0"
    If mFeeBalance.统筹支付 <> 0 Then
        str结算方式 = str结算方式 & "|统筹基金;" & mFeeBalance.统筹支付 & ";0"
    End If
    If mFeeBalance.大病支付 <> 0 Then
        str结算方式 = str结算方式 & "|大病统筹;" & mFeeBalance.大病支付 & ";0"
    End If
'    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    门诊虚拟结算_南京市 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_南京市(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & TYPE_南京市 & "," & mstrPatID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.门诊费用合计 & "," & mFeeBalance.自理费用 + mFeeBalance.个人自付 & ",0," & _
              mFeeBalance.医保范围费用 & "," & mFeeBalance.统筹支付 & "," & mFeeBalance.大病支付 & "," & _
              "0," & mFeeBalance.个人帐户支付 & ",null,null,null," & mFeeBalance.单据号 & ")"
    Call ExecuteProcedure("南京市医保")
    
    门诊结算_南京市 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_南京市(lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng销帐ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.结帐id  from 病人费用记录 A,病人费用记录 B where A.记录状态=2 and A.NO=B.NO and B.结帐id=" & lng结帐ID
    Call OpenRecordset(rsTemp, "销帐id")
    lng销帐ID = rsTemp!结帐ID
    
    gstrSQL = "select * from 保险结算记录 where 记录id=" & lng结帐ID
    Call OpenRecordset(rsTemp, "原始记录")
    If rsTemp.EOF Then
        MsgBox "保险结算记录中原始结帐单据不存在,不允许退费", vbInformation, gstrSysName
        Exit Function
    Else
        gstrSQL = "zl_保险结算记录_insert(1," & lng销帐ID & "," & TYPE_南京市 & "," & rsTemp!病人ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!发生费用金额 & "," & -rsTemp!全自付金额 & "," & -rsTemp!首先自付金额 & "," & -rsTemp!进入统筹金额 & "," & -rsTemp!统筹报销金额 & "," & -rsTemp!大病自付金额 & "," & _
              "0," & -rsTemp!个人帐户支付 & ",null,null,null,null)"
        Call ExecuteProcedure("销帐记录")
    End If
    
    门诊结算冲销_南京市 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_南京市(rsExse As Recordset, ByVal lng病人id As Long) As String
    Dim bytType As Byte
    Dim strFile As String, strWrite As String
    Dim strStream As String
    Dim dblSettleSum As Double
    Dim rsTemp As New ADODB.Recordset
    '删除可能存在的前次结算信息文件
    On Error Resume Next
    Call Kill("C:\NJYB\CYJSD.TXT")
    On Error GoTo errorhandle
    '上传还未上传的明细费用
    gstrSQL = "select 顺序号 from 保险帐户 where 病人id=" & lng病人id
    Call OpenRecordset(rsTemp, "顺序号")
    mDetailFee_南京市.住院序号 = rsTemp!顺序号
    
    '打开文件
    strFile = "C:\NJYB\ZYFYMX.TXT"
    Call writeTxtFile(strFile, "")
    Do Until rsExse.EOF
        If rsExse!是否上传 = 1 Then GoTo haddeliver             '找出已上传记录
        gstrSQL = "select decode(A.类别,'5',0,'6',0,'7',0,1) 标志,A.名称,A.编码,C.项目编码,A.计算单位,B.产地,decode(B.药品来源,'国产',1,'合资',2,'进口',3,null) 产地特征,B.规格" & _
                  " from 收费细目 A,药品目录 B,保险支付项目 C where A.id = C.收费细目id and A.id=B.药品id(+) and A.id =" & rsExse!收费细目ID
        Call OpenRecordset(rsTemp, "收细明细")
        
        With mDetailFee_南京市
            .标志 = rsTemp!标志
            .费用发生时间 = Format(rsExse!发生时间, "yyyyMMddHHmmss")
            .医院自编码 = rsTemp!编码
            .医保编码 = rsTemp!项目编码
            .名称 = rsTemp!名称
            .剂量单位 = zlCommFun.Nvl(rsTemp!计算单位)
            .单价 = rsExse!价格
            .数量 = rsExse!数量
            .产地 = zlCommFun.Nvl(rsTemp!产地)
            .产地特征 = zlCommFun.Nvl(rsTemp!产地特征)
            .规格 = zlCommFun.Nvl(rsTemp!规格)
        End With
        
        gstrSQL = "select 操作员编号 from 病人费用记录 where NO='" & rsExse!NO & "' and 序号=" & rsExse!序号 & _
                " and 记录性质=" & rsExse!记录性质 & " and 记录状态=" & rsExse!记录状态
        Call OpenRecordset(rsTemp, "操作员编号")
        mDetailFee_南京市.操作人编码 = rsTemp!操作员编号
        
        strWrite = mDetailFee_南京市.标志 & fillSpa(mDetailFee_南京市.住院序号, 20) & _
                   mDetailFee_南京市.费用发生时间 & _
                   fillSpa(mDetailFee_南京市.医保编码, 40) & fillSpa(mDetailFee_南京市.名称, 40) & _
                   fillSpa(mDetailFee_南京市.剂量单位, 10) & Lpad(mDetailFee_南京市.单价, 10) & _
                   Lpad(Format(mDetailFee_南京市.数量, "#0.00"), 10) & fillSpa(mDetailFee_南京市.操作人编码, 10) & _
                   fillSpa(mDetailFee_南京市.产地, 20) & fillSpa(mDetailFee_南京市.产地特征, 1) & fillSpa(mDetailFee_南京市.规格, 40)
        Call writeTxtFile(strFile, strWrite, False)
haddeliver:
        dblSettleSum = dblSettleSum + rsExse!金额           '得出结帐总金额
        rsExse.MoveNext
    Loop
    '关闭文件
    Call writeTxtFile(strFile, "", False)
    
    bytType = 9                          '表示住院预结算状态
    
    strStream = frm数据交换.getFeeBalance(bytType)
    On Error Resume Next
    Unload frm数据交换
    On Error GoTo errorhandle
    If strStream = "" Then
        MsgBox "读取医保结算文件过程被中止,无法完成预结算", vbInformation, gstrSysName
        Exit Function
    End If
    
    With mFeeBalance
        .住院序号 = analyseStr(strStream, 1, 20)
        .门诊费用合计 = Val(analyseStr(strStream, 35, 10))
        .医保范围费用 = Val(analyseStr(strStream, 65, 10))
        .自理费用 = Val(analyseStr(strStream, 75, 10))
        .个人自付 = Val(analyseStr(strStream, 85, 10))
        .统筹支付 = Val(analyseStr(strStream, 95, 10))
        .大病支付 = Val(analyseStr(strStream, 105, 10))
        .个人帐户支付 = Val(analyseStr(strStream, 115, 10))
    End With
    
    If mFeeBalance.住院序号 <> mDetailFee_南京市.住院序号 Then
        MsgBox "此结帐病人与医保结算文件中病人不一致,不能结算", vbInformation, gstrSysName
        Exit Function
    End If
    If Format(dblSettleSum, "#0.00") <> Format(mFeeBalance.门诊费用合计, "#0.00") Then
        MsgBox "请注意:医院总费用与医保中心返回的总费用不一致" & vbCrLf & _
        "总费用:(医院)￥" & Format(dblSettleSum, "#0.00") & Space(10) & "(医保)￥" & Format(mFeeBalance.门诊费用合计, "#0.00"), vbInformation, gstrSysName
    End If

    strStream = "统筹基金;" & mFeeBalance.统筹支付 & ";0"
    If mFeeBalance.个人帐户支付 <> 0 Then
        strStream = strStream & "|个人帐户;" & mFeeBalance.个人帐户支付 & ";0"
    End If
    If mFeeBalance.大病支付 <> 0 Then
        strStream = strStream & "|大病统筹;" & mFeeBalance.大病支付 & ";0"
    End If
    
    住院虚拟结算_南京市 = strStream
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_南京市(lng结帐ID As Long, lng病人id) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errorhandle
    gstrSQL = "select NO,序号,记录状态,记录性质 from 病人费用记录 where nvl(是否上传,0)=0 and 结帐id=" & lng结帐ID
    Call OpenRecordset(rsTemp, "查找记录")
    Do Until rsTemp.EOF
        gstrSQL = "ZL_病人费用记录_上传('" & rsTemp!NO & "'," & rsTemp!序号 & "," & rsTemp!记录性质 & "," & rsTemp!记录状态 & ")"
        Call ExecuteProcedure("更新上传标志")
        rsTemp.MoveNext
    Loop
    
    gstrSQL = "select 住院次数 from 病人信息 where 病人id=" & lng病人id
    Call OpenRecordset(rsTemp, "主页id")
    
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & TYPE_南京市 & "," & lng病人id & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              mFeeBalance.门诊费用合计 & "," & mFeeBalance.自理费用 + mFeeBalance.个人自付 & ",0," & _
              mFeeBalance.医保范围费用 & "," & mFeeBalance.统筹支付 & "," & mFeeBalance.大病支付 & "," & _
              "0," & mFeeBalance.个人帐户支付 & ",'" & mFeeBalance.住院序号 & "'," & rsTemp!住院次数 & ",null,null)"
    ExecuteProcedure ("插入保险帐户")
    
    住院结算_南京市 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_南京市(lng结帐ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng销帐ID As Long
    
    On Error GoTo errorhandle
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "销帐id")
    lng销帐ID = rsTemp!ID
    
    gstrSQL = "select * from 保险结算记录 where 记录id=" & lng结帐ID
    Call OpenRecordset(rsTemp, "原始记录")
    If rsTemp.EOF Then
        MsgBox "保险结算记录中原始结帐单据不存在,不允许退费", vbInformation, gstrSysName
        Exit Function
    Else
        gstrSQL = "zl_保险结算记录_insert(2," & lng销帐ID & "," & TYPE_南京市 & "," & rsTemp!病人ID & "," & Year(zlDatabase.Currentdate) & ",null,null,null,null,null,null,null,null," & _
              -rsTemp!发生费用金额 & "," & -rsTemp!全自付金额 & "," & -rsTemp!首先自付金额 & "," & -rsTemp!进入统筹金额 & "," & -rsTemp!统筹报销金额 & "," & -rsTemp!大病自付金额 & "," & _
              "0," & -rsTemp!个人帐户支付 & ",'" & rsTemp!支付顺序号 & "'," & rsTemp!主页ID & ",null,null)"
        ExecuteProcedure ("销帐记录")
    End If
    
    住院结算冲销_南京市 = True
    Exit Function
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub writeTxtFile(strFile As String, strWrite As String, Optional ByVal openFile As Boolean = True)
    Dim intSymbol As Long
    Dim strFolder As String
    
    On Error GoTo errorhandle
    Do Until InStr(intSymbol + 1, strFile, "\") = 0
        intSymbol = InStr(intSymbol + 1, strFile, "\")
        strFolder = Mid(strFile, 1, intSymbol)
        If Not mobjSystem.FolderExists(strFolder) Then mobjSystem.CreateFolder (strFolder)
    Loop

    If openFile Then                    '打开文件
        If Not mobjSystem.FileExists(strFile) Then mobjSystem.CreateTextFile (strFile)
        Set mobjStream = mobjSystem.OpenTextFile(strFile, ForWriting)
        If strWrite <> "" Then          '如果有内容进行写入
            mobjStream.WriteLine (UCase(strWrite))
            mobjStream.Close
        End If
    Else
        If strWrite = "" Then
            mobjStream.Close
        Else
            mobjStream.WriteLine (UCase(strWrite))   '如果有写入内容但打开标志为false,只进行写入
        End If
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mobjStream.Close
End Sub

Public Function readTxtFile(strFile As String) As String
    On Error GoTo errHandle
    
    If mobjSystem.FileExists(strFile) Then
        Set mobjStream = mobjSystem.OpenTextFile(strFile)
        readTxtFile = mobjStream.ReadLine
        mobjStream.Close
    End If
    Exit Function
    
errHandle:
    Err.Clear
    On Error Resume Next
    mobjStream.Close
End Function

Private Function fillSpa(strTemp As Variant, lngLen As Long, Optional fromRigth As Boolean = True) As String
    Dim lngStrLeng As Long
    Dim strStream As String
    Dim strUnion As String
    
    strTemp = IIf(IsNull(strTemp), "", Trim(strTemp))
    
    strUnion = StrConv(Trim(strTemp), vbFromUnicode)
    lngStrLeng = IIf(LenB(strUnion) > lngLen, lngLen, LenB(strUnion))
    strStream = IIf(LenB(strUnion) > lngLen, StrConv(LeftB(strUnion, 20), vbUnicode), strTemp)
    
    If fromRigth Then
        fillSpa = strStream & String(lngLen - lngStrLeng, " ")
    Else
        fillSpa = String(lngLen - lngStrLeng, " ") & strStream
    End If
End Function

Public Function analyseStr(strTemp As String, lngStart As Long, lngLen As Long) As String
    Dim strStream As String
    
    strStream = StrConv(UCase(strTemp), vbFromUnicode)
    
    analyseStr = Trim(StrConv(MidB(strStream, lngStart, lngLen), vbUnicode))
End Function

Public Function 个人余额_南京市(ByVal lng病人id As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
'    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_南京市
'    Call OpenRecordset(rsTemp, gstrSysName)
'
'    If rsTemp.EOF Then
'        个人余额_南京市 = 100000
'    Else
'        个人余额_南京市 = IIf(rsTemp("帐户余额") = 0, 100000, rsTemp("帐户余额"))
'    End If
    个人余额_南京市 = 100000
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(strNO As String) As String
'功能：由用户输入的部份单号，返回当年的单号。
    If Len(strNO) >= 8 Then GetFullNO = Right(strNO, 8): Exit Function
    GetFullNO = PreFixNO & Format(strNO, "0000000")
End Function

Public Function FileExists(ByVal FileName As String, Optional ErrFlag As Boolean = True) As Boolean
    Dim Temp
    FileExists = True
    On Error Resume Next
proshow:
    Temp = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                If ErrFlag Then
                    If MsgBox("磁盘没有准备好。", vbInformation + vbRetryCancel, "错误") = vbRetry Then
                        GoTo proshow:
                    End If
                End If
                FileExists = False
            End If
    End Select
End Function
