Attribute VB_Name = "mdl银海新都"
Option Explicit
'API函数声明

'1、数据上传
Private Declare Function DataUnloading Lib "yhybReckoning.dll" Alias "_DataUnloading@12" _
        (ByVal str_UploadData As String, ByVal str_UploadLsh As String, ByVal str_Fzxbm As String) As String

'2、帐户支付
Private Declare Function reckoning Lib "yhybReckoning.dll" Alias "_reckoning@64" (ByVal str卡号 As String, _
        ByVal str医保号 As String, ByVal str分中心 As String, ByVal str密码 As String, _
        ByVal str就诊顺序号 As String, ByVal str支付类别 As String, ByVal str医院编码 As String, _
        ByVal str分院编码 As String, ByVal dbl帐户支付 As String, ByVal dat支付时间 As String, _
        ByVal dbl总额 As String, ByVal dbl全自费 As String, ByVal dbl挂钩自付 As String, _
        ByVal dbl允许报销 As String, ByVal str经办人 As String, ByVal str结算编号 As String) As String

'3、获取当前医院基本信息
Private Declare Function GetHospitalInfo Lib "yhybReckoning.dll" Alias "_GetHospitalInfo@0" () As String

'4、费用明细分割
'Private Declare Function DivideUp Lib "yhybDivideUp.dll" Alias "_DivideUp@24" _
        (ByVal str分中心编号 As String, ByVal str医保项目编码 As String, ByVal str支付类别 As String, _
        ByVal str医疗人员类别 As String, ByVal dbl分割金额 As Double) As String
Private Declare Function DivideUp Lib "yhybReckoning.dll" Alias "_DivideUp@24" _
        (ByVal str分中心编号 As String, ByVal str医保项目编码 As String, ByVal str支付类别 As String, _
        ByVal str医疗人员类别 As String, ByVal dbl分割金额 As Double) As String

'5、计算可支付金额
Private Declare Function GetPayCount Lib "yhybReckoning.dll" Alias "_GetPayCount@48" _
        (ByVal str分中心编号 As String, ByVal str支付类别 As String, _
        ByVal dbl个人自付 As Double, ByVal dbl全自费 As Double, ByVal dbl挂钩自费 As Double, _
        ByVal dbl起付线 As Double, ByVal dbl帐户余额 As Double) As String

'6、费用结算
Private Declare Function CalculateFeeCD Lib "yhybBill.dll" Alias "_CalculateFeeCD@84" _
        (ByVal dbl费用总额 As Double, ByVal dbl起付线 As Double, ByVal dbl统筹限额 As Double, _
        ByVal dbl统筹支付累计 As Double, ByVal int实足年龄 As Integer, ByVal dbl已结算起付线 As Double, _
        ByVal dbl已结算挂钩自付 As Double, ByVal dbl允许报销部分 As Double, ByVal dbl全自费 As Double, _
        ByVal dbl挂钩自费 As Double, ByVal dbl统筹报销比例 As Double) As String
'7、医保服务目录文件
Private Declare Function MakeTxt Lib "yhybReckoning.dll" Alias "_MakeTxt@8" (ByVal str服务目录文件 As String, _
        ByVal str病种目录文件 As String) As String

'8、卡解析服务
Private Declare Function GetKard Lib "yhybReckoning.dll" Alias "_GetKard@4" (ByVal str_UploadData As String) As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public mint适用地区_新都 As Integer
Private mstr医保号 As String
Private mstr密码 As String
Private mlng病人ID As Long
Private mstr门诊号 As String
Private mstrInfo As String                      '调试信息，用于产生日志文件
Private mstr门诊流水号 As String                '由于住院允许进行门诊业务，因此门诊顺序号未更新到保险帐户中，避免冲了住院的顺序号
Private mcol诊明细 As New Collection

Public Function 医保初始化_新都() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    '提取当前接口适用地区
    mint适用地区_新都 = 0
    '保留以下代码,可能会有多个地区使用本接口
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=" & TYPE_新都 & " And 序号=1"
    Call OpenRecordset(rsTemp, "提取当前接口适用地区")
    If Not rsTemp.EOF Then mint适用地区_新都 = Nvl(rsTemp!参数值, 0)
    
    医保初始化_新都 = True
End Function

Public Function 医保设置_新都() As Boolean
'功能： 该方法用于供相关应用部件调用配置连接医保数据服务器的连接串
'返回：接口配置成功，返回true；否则，返回false
    Dim strConn As String
    
    医保设置_新都 = frmSet新都银海.ShowSet
End Function

Public Function 身份标识_新都(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str卡号 As String, str医保号 As String, str密码 As String
    Dim str姓名 As String, str性别 As String, str身份证号码 As String, lng年龄 As Long
    Dim str出生日期 As String, str人员类别 As String, str单位编码 As String, str单位名称 As String
    Dim strIdentify As String, str附加 As String, str门诊号 As String
    Dim datCurr As Date, str医院编码 As String
    Dim strReturn As String, str流水号 As String, str住院顺序号 As String, str中心编号 As String, strInput As String, arrOutput As Variant
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency, cur个人帐户 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency, cur本次起付线 As Currency, cur起付线累计 As Currency
    Dim int住院次数累计 As Integer, bln读取帐户年度信息 As Boolean, cur统筹限额 As Currency
    bln读取帐户年度信息 = False
    
    '初始化一些变量
    mlng病人ID = 0
    mstr门诊号 = ""
    mstr医保号 = ""
    mstr密码 = ""
    
    '获得病人医保号、分中心编号等信息
    If frmIdentify贵阳.GetIdentify(TYPE_新都, str卡号, str医保号, str中心编号, str密码) = False Then Exit Function
    
    '检查该病人是否以医保身份正在住院
    Dim rsTemp As New ADODB.Recordset
    '检查该病人是否在院
    gstrSQL = "select nvl(当前状态,0) as 当前状态,顺序号 from 保险帐户 where 医保号='" & str医保号 & "' and 险类=" & TYPE_新都
    Call OpenRecordset(rsTemp, gstrSysName)
    
    If rsTemp.EOF = False Then
        If rsTemp("当前状态") = 1 Then
            '双流地区允许在院期间发生门诊业务
            str住院顺序号 = Nvl(rsTemp!顺序号)
'            If mint适用地区_新都 = 1 Then
'                MsgBox "该病人正以医保身份在院，不能再进行身份验证。", vbInformation, gstrSysName
'                Exit Function
'            End If
        End If
    End If
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '调用身份验证
    If Get流水号("A", str医院编码, str流水号) = False Then Exit Function
    '卡号|个人编码|分中心编号|密码|获取何种流水号#
    strInput = str卡号 & "|" & str医保号 & "|" & str中心编号 & "|" & str密码 & "|" & IIf(bytType = 1, "31", "11") & "#"
    Call WriteLog("DataUnloading(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '双流加的判断，如果密码为111111，说明是初始密码，必须要求用户修改，并退出本次交易
'    If mint适用地区_新都 = 1 Then
'        If str密码 = "111111" Then
'            MsgBox "此密码为社保局初始密码，请重新输入密码！", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    '取得返回值
    str卡号 = arrOutput(1)
    str医保号 = arrOutput(3)
    
    str姓名 = arrOutput(4)
    str性别 = IIf(arrOutput(5) = "2", "女", "男")
    str身份证号码 = arrOutput(6)
    str出生日期 = arrOutput(7)
    If IsDate(str出生日期) = False Then
        str出生日期 = Get出生日期(str身份证号码, 0)
    End If
    If IsDate(str出生日期) Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), zlDataBase.Currentdate)
        str出生日期 = Format(CDate(str出生日期), "yyyy-MM-dd")
    Else
        str出生日期 = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    End If
    
    str人员类别 = arrOutput(8)
    str单位编码 = arrOutput(9)
    str单位名称 = arrOutput(10)
    '双流地区允许在院期间发生门诊业务，因此，在进行门诊业务时，如果住院顺序号不为空，说明在院，不更新顺序号
    str流水号 = arrOutput(12)
    mstr门诊流水号 = arrOutput(12)
    If str住院顺序号 <> "" Then str流水号 = str住院顺序号
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    strIdentify = str卡号 & ";" & str医保号 & ";;" & str姓名 & ";" & str性别 & ";" & str出生日期 & ";" & str身份证号码 & ";" & str单位名称 & "(" & str单位编码 & ")"
    strIdentify = Replace(strIdentify, " ", "")
    cur个人帐户 = arrOutput(11)
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";" & str流水号                 '9.顺序号
    str附加 = str附加 & ";" & str人员类别               '10人员身份
    str附加 = str附加 & ";" & arrOutput(11)             '11帐户余额
    str附加 = str附加 & ";" & IIf(str住院顺序号 <> "", "1", "0")                       '12当前状态
    str附加 = str附加 & ";"                             '13病种ID
    str附加 = str附加 & ";" & IIf(Left(str人员类别, 1) = "退", 2, 1)     '14在职(1,2)
    str附加 = str附加 & ";" & str中心编号               '15退休证号 但本医保用于保存医保分中心编码（避免建立医保中心）
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & cur个人帐户             '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";"                             '23就诊类型 (1、急诊门诊)
        '建立病人档案信息，传入格式：
        '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8中心;9.顺序号;
        '10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
        '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
    
    gstrSQL = "Select * From 保险帐户 Where 医保号='" & str医保号 & "' And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "取病人ID")
    If Not rsTemp.EOF Then
        lng病人ID = rsTemp!病人ID
    End If
    datCurr = zlDataBase.Currentdate
    If lng病人ID <> 0 Then          '如果病人已存在，则读取帐户年度信息
        '帐户年度信息
        Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur统筹限额)
        bln读取帐户年度信息 = True
    End If
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID)
    
    If bln读取帐户年度信息 = True Then          '如果读取出帐户年度信息，则重新写入
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
            cur个人帐户 & ",0," & _
            cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur统筹限额 & ")"
        Call ExecuteProcedure(gstrSysName)
    End If
    
    '返回格式:中间插入病人ID
    If lng病人ID <> 0 Then
        身份标识_新都 = strIdentify & ";" & lng病人ID & str附加
        
        mstr医保号 = str医保号
        mstr密码 = str密码
    Else
        mstr门诊流水号 = ""
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_新都(strSelfNo As String, ByVal bytPlace As Byte) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'返回: 返回个人帐户余额的金额
    Dim rsTemp As New ADODB.Recordset, str卡号 As String, str医保号 As String, str密码 As String
    Dim strReturn As String, str流水号 As String, str中心编号 As String, strInput As String, arrOutput  As Variant
    Dim str医院编码 As String
    
    On Error GoTo errHandle
    
    
    If bytPlace = balan预交 Then
        '在病人入院与缴预交之间可变化，可以导致病人余额已经不准确了
        '获得病人医保号、分中心编号等信息
        If frmIdentify贵阳.GetIdentify(TYPE_新都, str卡号, str医保号, str中心编号, str密码) = False Then Exit Function
        
        If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
        
        '调用身份验证
        If Get流水号("A", str医院编码, str流水号) = False Then Exit Function
        strInput = str卡号 & "|" & str医保号 & "|" & str中心编号 & "|" & str密码 & "|11#"
        Call WriteLog("DataUnloading(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
        strReturn = DataUnloading(strInput, str流水号, str中心编号)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        mstr医保号 = str医保号
        mstr密码 = str密码
        个人余额_新都 = Val(arrOutput(11))
    Else
        '从数据库中读取（因为刚才才保存了的，应该是准确的）
        gstrSQL = "Select 帐户余额 From 保险帐户 where 险类=" & gintInsure & " and 中心=0 and 医保号='" & strSelfNo & "'"
        Call OpenRecordset(rsTemp, gstrSysName)
        
        If rsTemp.EOF = False Then
            个人余额_新都 = IIf(IsNull(rsTemp("帐户余额")), 0, rsTemp("帐户余额"))
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_新都(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str医保号 As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim dbl个人帐户 As Double
    Dim lng病人ID As Long, datCurr As Date, lng序号 As Long
    Dim str中心编号 As String, str就诊顺序号 As String, str人员类别 As String, str医院编码 As String
    Dim dbl总金额 As Double, dbl全自费 As Double, dbl挂钩自付 As Double, dbl起付线 As Double, dbl余额 As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If rs明细.RecordCount = 0 Then
        str结算方式 = "个人帐户;0;0"
        门诊虚拟结算_新都 = True
        Exit Function
    End If
    rs明细.MoveFirst
    lng病人ID = rs明细("病人ID")
    datCurr = zlDataBase.Currentdate
    
    '从保险帐户获得登记信息
    gstrSQL = "select 医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别  " & _
              "from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "门诊预算")
    'str就诊顺序号 = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))
    str就诊顺序号 = mstr门诊流水号
    str中心编号 = IIf(IsNull(rsTemp("中心编号")), "", rsTemp("中心编号"))
    str人员类别 = IIf(IsNull(rsTemp("人员类别")), "", rsTemp("人员类别"))
    str医保号 = rsTemp("医保号")
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '首先清除已经保存的费用明细
    Set mcol诊明细 = Nothing
    
    '然后插入处方明细
    Do Until rs明细.EOF
        '得到费用明细
        gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,C.类别 as 大类 " & _
                 " from 收费细目 A,保险支付项目 B,收费类别 C " & _
                 " where A.类别=C.编码 and  A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID and B.险类=" & gintInsure
        Call OpenRecordset(rsTemp, "门诊预算")
        
        '进行费用分割
        strReturn = DivideUp(str中心编号, ToVarchar(rsTemp("项目编码"), 12), "11", str人员类别, Val(rs明细("单价")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '第二个就诊顺序号是作为结算单号
        strInput = str就诊顺序号 & "|" & str就诊顺序号
        strInput = strInput & "|" & str就诊顺序号 & "_" & lng序号      '序号
        strInput = strInput & "|" & str医保号 & "|" & str中心编号 & "|" & str医院编码 & "|000"
        strInput = strInput & "|" & ToVarchar(rsTemp("项目编码"), 12)  '医保流水号
        strInput = strInput & "|" & ToVarchar(rsTemp("大类"), 10)      '收费大类名称
        strInput = strInput & "|" & Format(rs明细("数量"), "0.00")
        strInput = strInput & "|" & Format(rs明细("单价"), "0.00")
        strInput = strInput & "|" & Format(rs明细("实收金额"), "0.00")
        strInput = strInput & "|" & arrOutput(4)                       '自付比例
        strInput = strInput & "|" & Format(Val(arrOutput(1)) * rs明细("数量"), "#0.00") '全自费部分
        strInput = strInput & "|" & Format(Val(arrOutput(2)) * rs明细("数量"), "#0.00") '挂钩自费部分
        strInput = strInput & "|" & Format(Val(arrOutput(3)) * rs明细("数量"), "#0.00") '允许报销部分
        strInput = strInput & "||11"                                   '特项标志、支付类别
        strInput = strInput & "|" & ToVarchar(UserInfo.部门, 56)       '开单科室名称
        strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)       '开单处方医生
        strInput = strInput & "|" & ToVarchar(UserInfo.部门, 56)       '受单科室名称
        strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)       '受单处方医生
        strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)        '经办人
        strInput = strInput & "|" & Format(datCurr + lng序号 / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '经办时间
        strInput = strInput & "|" & ToVarchar(rsTemp("名称"), 200)       '收费项目
        strInput = strInput & "|" & ToVarchar(rsTemp("规格"), 200)       '规格
        strInput = strInput & "|"                                        '产地
        strInput = strInput & "|" & ToVarchar(rsTemp("计算单位"), 30)    '单位
        strInput = strInput & "|||"                                      '英文名、化学名
        strInput = strInput & lng序号 & "#"                             '序号
        Call WriteLog(strInput)
        mcol诊明细.Add strInput  '首先将明细保存，待结算时再上传
        
        lng序号 = lng序号 + 1
        dbl总金额 = dbl总金额 + Val(rs明细("实收金额"))
        dbl全自费 = dbl全自费 + Val(arrOutput(1)) * rs明细("数量")
        dbl挂钩自付 = dbl挂钩自付 + Val(arrOutput(2)) * rs明细("数量")
        dbl起付线 = dbl起付线 + Val(arrOutput(3)) * rs明细("数量")    '目前使用允许报销部分。待定
        
        rs明细.MoveNext
    Loop
    
    '得到病人余额
    dbl余额 = 个人余额_新都(str医保号, balan门诊)
    With g结算数据
        .发生费用金额 = dbl总金额
        .全自费金额 = dbl全自费
        .首先自付金额 = dbl挂钩自付
        .进入统筹金额 = dbl起付线
        .支付顺序号 = str就诊顺序号
    End With
    '调用预结算
    strReturn = GetPayCount(str中心编号, "11", dbl总金额, dbl全自费, dbl挂钩自付, dbl起付线, dbl余额)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    dbl个人帐户 = Val(arrOutput(1))                 '取接口允许帐户支付的金额
    If mint适用地区_新都 = 0 Then
        '双流允许全部由帐户支付，接口返回的帐户支付不再有效
        dbl个人帐户 = IIf(dbl余额 < dbl总金额, dbl余额, dbl总金额)
    End If
    str结算方式 = "个人帐户;" & dbl个人帐户 & ";1"   '允许修改个人帐户
    门诊虚拟结算_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_新都(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '此时所有收费细目必然有对应的医保编码
    Dim str医保号 As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim lng病人ID  As Long, rs明细 As New ADODB.Recordset
    Dim datCurr As Date, var明细 As Variant, rsTemp As New ADODB.Recordset
    Dim str中心编号 As String, str就诊顺序号 As String, str医院编码 As String, str卡号 As String, str流水号 As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select * From 病人费用记录 Where 结帐ID=" & lng结帐ID
    Call OpenRecordset(rs明细, "门诊预算")
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    lng病人ID = rs明细("病人ID")
    datCurr = rs明细("登记时间")
    
    If mstr医保号 <> strSelfNo Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '获得帐户相关信息
    gstrSQL = "select 卡号,医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别  " & _
              "from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rs明细, "门诊预算")
    str就诊顺序号 = mstr门诊流水号
    str中心编号 = IIf(IsNull(rs明细("中心编号")), "", rs明细("中心编号"))
    str卡号 = IIf(IsNull(rs明细("卡号")), "", rs明细("卡号")) '条码卡就没有卡号
    str医保号 = rs明细("医保号")
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '上传费用明细，统一用一个流水号（待定）
    If Get流水号("G", str医院编码, str流水号) = False Then Exit Function
    For Each var明细 In mcol诊明细
        Call WriteLog("上传:" & var明细)
        strReturn = DataUnloading(var明细, str流水号, str中心编号)
        Call WriteLog("返回:" & strReturn)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    Next
    
    '调用结算
    With g结算数据
    Call WriteLog("下帐(" & str卡号 & "," & str医保号 & "," & str中心编号 & "," & mstr密码 & "," & str就诊顺序号 & "," & "11" & "," & str医院编码 & "," & "000" & "," & CStr(cur个人帐户) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(.发生费用金额) & "," & CStr(.全自费金额) & "," & CStr(.首先自付金额) & "," & CStr(.进入统筹金额) & "," & ToVarchar(UserInfo.姓名, 20) & "," & ToVarchar(.支付顺序号, 20) & ")")
    strReturn = reckoning(str卡号, str医保号, str中心编号, mstr密码, str就诊顺序号, "11", str医院编码, "000", Format(cur个人帐户, "0.##"), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               Format(.发生费用金额, "0.##"), Format(.全自费金额, "0.##"), Format(.首先自付金额, "0.##"), Format(.进入统筹金额, "0.##"), ToVarchar(UserInfo.姓名, 20), ToVarchar(.支付顺序号, 20))
    Call WriteLog("返回:" & strReturn)
    End With
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
    Dim cur起付线累计 As Currency, cur本次起付线 As Currency, cur统筹限额 As Currency
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计, cur本次起付线, cur起付线累计, cur统筹限额)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur本次起付线 & "," & cur起付线累计 & "," & cur统筹限额 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & g结算数据.发生费用金额 & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur个人帐户 & ",'')"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    门诊结算_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_新都(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额

    门诊结算冲销_新都 = True
End Function

Public Function 个人帐户转预交_新都(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim str医保号 As String, strInput As String, arrOutput  As Variant, strReturn As String
    Dim datCurr As Date, var明细 As Variant, rs明细 As New ADODB.Recordset
    Dim str中心编号 As String, str就诊顺序号 As String, str医院编码 As String, str卡号 As String, str流水号 As String
    
    On Error GoTo errHandle
    
    datCurr = zlDataBase.Currentdate
    
    If mstr医保号 <> strSelfNo Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '获得帐户相关信息
    gstrSQL = "select 卡号,医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别  " & _
              "from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rs明细, "缴预交金")
    str就诊顺序号 = IIf(IsNull(rs明细("就诊序号")), "", rs明细("就诊序号"))
    str中心编号 = IIf(IsNull(rs明细("中心编号")), "", rs明细("中心编号"))
    str卡号 = IIf(IsNull(rs明细("卡号")), "", rs明细("卡号")) '条码卡没有卡号
    str医保号 = rs明细("医保号")
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '首先判断金额是否可以使用
    strReturn = GetPayCount(str中心编号, "31", cur个人帐户, 0, 0, cur个人帐户, cur个人帐户)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    If Val(arrOutput(1)) < cur个人帐户 Then
        MsgBox "个人帐户不能用于支付预交金。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '调用结算
    Call WriteLog("reckoning(" & str卡号 & "," & str医保号 & "," & str中心编号 & "," & mstr密码 & "," & str就诊顺序号 & "," & "31" & "," & str医院编码 & "," & "000" & "," & CStr(cur个人帐户) & "," & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "," & _
               CStr(cur个人帐户) & "," & CStr(0) & "," & CStr(0) & "," & CStr(cur个人帐户) & "," & ToVarchar(UserInfo.姓名, 20) & "," & ToVarchar(str就诊顺序号, 20) & ")")
    strReturn = reckoning(str卡号, str医保号, str中心编号, mstr密码, str就诊顺序号, "31", str医院编码, "000", CDbl(cur个人帐户), Format(datCurr, "yyyy-MM-dd HH:mm:ss"), _
               CDbl(cur个人帐户), CDbl(0), CDbl(0), CDbl(cur个人帐户), ToVarchar(UserInfo.姓名, 20), ToVarchar(str就诊顺序号, 20))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer
            
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
                
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & _
        cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure(gstrSysName)
    
    '
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur个人帐户 & ",0,0," & _
        0 & "," & 0 & ",0,0," & cur个人帐户 & ",'')"
    Call ExecuteProcedure(gstrSysName)
    '---------------------------------------------------------------------------------------------

    个人帐户转预交_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_新都(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String, Optional ByVal blnFirst As Boolean = True) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
    Dim strInput As String, arrOutput  As Variant, arrTmp As Variant
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str流水号 As String, strReturn As String
    Dim str中心编号 As String, str就诊顺序号 As String, str医院编码 As String, str卡号 As String
    Dim str入院诊断 As String, str出院诊断 As String, str医院类别 As String, str入院年份 As String
    Dim intValue As Integer
    Dim dbl统筹限额 As Double, dbl统筹累计 As Double

    On Error GoTo errHandle
    
    '获取保险参数值，以决定医保病人入院时，是否同时上传入院信息
    intValue = 1
'    gstrSQL = "Select Nvl(参数值,0) Value From 保险参数 Where 险类=" & TYPE_新都 & " And 参数名='上传入院信息'"
'    Call OpenRecordset(rsTemp, "获取上传入院信息参数值")
'
'    If Not rsTemp.EOF Then
'        intValue = rsTemp!Value
'    End If
    
    '获得医保号
    gstrSQL = "select 医保号,卡号,顺序号 as 就诊序号,退休证号 as 中心编号 from 保险帐户 where 险类=" & TYPE_新都 & " and 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "入院登记")
    
    str卡号 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号")) '如果是条码卡,卡号就为空
    str医保号 = rsTemp("医保号")
    str就诊顺序号 = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))
    str中心编号 = IIf(IsNull(rsTemp("中心编号")), "", rsTemp("中心编号"))
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '获得病人出院诊断
    gstrSQL = "select A.诊断类型,A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型 in (1,3) and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "出院登记")
    Do Until rsTemp.EOF
        If rsTemp("诊断类型") = 1 Then
            str入院诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
        Else
            str出院诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str入院诊断 = "" Then str入院诊断 = "疾病" '诊断不论如何不能为空
    If str出院诊断 = "" Then str出院诊断 = "疾病" '诊断不论如何不能为空
    
    '获得其它入院信息
    datCurr = zlDataBase.Currentdate
    gstrSQL = " select A.入院日期,A.登记时间,B.名称 入院科室 " & _
              " from 病案主页 A,部门表 B " & _
              " Where A.入院科室ID=B.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "入院登记")
    str入院年份 = Year(rsTemp!入院日期)
    
    '获得医院类别
    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Function
    
    '调用入院接口
    If blnFirst Then
        '待遇审核
        If Get流水号("C", str医院编码, str流水号) = False Then Exit Function
        strInput = str卡号 & "|" & str医保号 & "|" & str中心编号 & "|" & mstr密码 & _
                    "|" & str就诊顺序号 & "|" & str医院编码 & _
                    "|000|0|000|31|0" & _
                    "|" & Format(rsTemp("入院日期"), "yyyy-MM-dd HH:mm:ss") & _
                    "|" & ToVarchar(UserInfo.姓名, 20) & _
                    "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"
        Call WriteLog("DataUnloadint(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
        strReturn = DataUnloading(strInput, str流水号, str中心编号)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        '如果是双流县医保，需要对基本统筹限额进行判断，如果为零，则禁止办理入院，提示以普通病人办理，同时，
        '需要将基本统筹限额与统筹支付累计提示出来给操作员
        If mint适用地区_新都 = 0 Then
            If Val(arrOutput(6)) = 0 Then
                MsgBox "基本统筹限额为零，不允许再以医保身份入院，请按普通病人办理入院！", vbInformation, gstrSysName
                Exit Function
            End If
            dbl统筹限额 = Val(arrOutput(6))
            dbl统筹累计 = Val(arrOutput(8))
        End If
        
        '上传入院登记
        If intValue = 1 Then
            If Get流水号("E", str医院编码, str流水号) = False Then Exit Function
            strInput = str就诊顺序号 & "|" & str医保号 & "|" & str医院编码 & "|000|" & str医院类别 & "|31|0"
            strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)    '入院经办人
            strInput = strInput & "|" & ToVarchar(rsTemp("入院科室"), 20)  '入院科室
            strInput = strInput & "|" & str入院诊断
            strInput = strInput & "|" & Format(rsTemp("入院日期"), "yyyy-MM-dd HH:mm:ss")
            strInput = strInput & "|" & Format(rsTemp("登记时间"), "yyyy-MM-dd HH:mm:ss") & "|||入院登记|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|" & Format("2000-01-01", "yyyy-MM-dd HH:mm:ss") & "|9#"
            Call WriteLog("上传入院登记(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
            strReturn = DataUnloading(strInput, str流水号, str中心编号)
            Call WriteLog("返回:" & strReturn)
            If JudgeReturn(strReturn, arrTmp) = False Then Exit Function
        End If
        
        Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
        Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
        Dim int住院次数累计 As Integer
                
        '帐户年度信息   注意各字段与实际用处之间的对应关系
        '本次起付线    ----   住院起付线
        '起付线累计    ----   基本统筹支付累计
        '基本统筹限额  ----   住院统筹限额
        '大额统筹限额  ----   实足年龄
        '大额统筹累计  ----   统筹报销比例
        Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & str入院年份 & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & cur统筹报销累计 & "," & int住院次数累计 & "," & _
            arrOutput(5) & "," & arrOutput(8) & "," & arrOutput(6) & "," & arrOutput(3) & "," & arrOutput(11) & ")"
        Call ExecuteProcedure(gstrSysName)
        
        '个人状态的修改
        gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
        Call ExecuteProcedure(gstrSysName)
        
        If mint适用地区_新都 = 0 Then
            MsgBox "该参保病人的住院相关信息：" & vbCrLf & _
                   "    基本统筹限额：￥" & Format(dbl统筹限额, "#0.00") & _
                   "    统筹支付累计：￥" & Format(dbl统筹累计, "#0.00"), vbInformation, gstrSysName
        End If
    End If
    
    入院登记_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_新都(lng病人ID As Long, lng主页ID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim datCurr As Date, rsTemp As New ADODB.Recordset, str入院诊断 As String, str出院诊断 As String
    Dim strInput As String, arrOutput  As Variant, str流水号 As String, strReturn As String
    Dim str中心编号 As String, str就诊顺序号 As String, str医院编码 As String, str医保号 As String
    Dim str医院类别 As String
    
    On Error GoTo errHandle
    
    '获得医保号
    gstrSQL = "select 医保号,卡号,顺序号 as 就诊序号,退休证号 as 中心编号 from 保险帐户 where 险类=" & TYPE_新都 & " and 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "入院登记")
    str医保号 = rsTemp("医保号")
    str就诊顺序号 = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))
    str中心编号 = IIf(IsNull(rsTemp("中心编号")), "", rsTemp("中心编号"))
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    '获得病人出院诊断
    gstrSQL = "select A.诊断类型,A.描述信息 from 诊断情况 A where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页ID & _
              " and A.诊断类型 in (1,3) and A.诊断次序=1"
    Call OpenRecordset(rsTemp, "出院登记")
    Do Until rsTemp.EOF
        If rsTemp("诊断类型") = 1 Then
            str入院诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
        Else
            str出院诊断 = ToVarchar(IIf(IsNull(rsTemp("描述信息")), "疾病", rsTemp("描述信息")), 128)
        End If
        rsTemp.MoveNext
    Loop
    If str入院诊断 = "" Then str入院诊断 = "疾病" '诊断不论如何不能为空
    If str出院诊断 = "" Then str出院诊断 = "疾病" '诊断不论如何不能为空
        
    '获得其它出院信息
    datCurr = zlDataBase.Currentdate
    gstrSQL = "select A.门诊医师,A.住院医师,A.登记时间,A.入院日期,A.出院日期,A.出院方式,B.名称 as 入院科室,C.名称 as 出院科室 " & _
             " from 病案主页 A,部门表 B,部门表 C " & _
             " Where A.入院科室ID = B.ID And A.出院科室ID = C.ID And A.病人ID =" & lng病人ID & " And A.主页ID = " & lng主页ID
    Call OpenRecordset(rsTemp, "出院登记")
    
    '获得医院类别
    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Function

    '调用出院接口
    If Get流水号("E", str医院编码, str流水号) = False Then Exit Function
    strInput = str就诊顺序号 & "|" & str医保号 & "|" & str医院编码 & "|000|" & str医院类别 & "|31|" & _
                IIf(Format(rsTemp("入院日期"), "yyyy") = Format(rsTemp("出院日期"), "yyyy"), "0", "1")
    strInput = strInput & "|" & ToVarchar(rsTemp("门诊医师"), 20)  '入院经办人
    strInput = strInput & "|" & ToVarchar(rsTemp("入院科室"), 20)  '入院科室
    strInput = strInput & "|" & str入院诊断
    strInput = strInput & "|" & Format(rsTemp("入院日期"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & Format(rsTemp("登记时间"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)       '出院经办人
    strInput = strInput & "|" & ToVarchar(rsTemp("出院科室"), 20)  '出院科室
    strInput = strInput & "|" & str出院诊断
    strInput = strInput & "|" & Format(rsTemp("出院日期"), "yyyy-MM-dd HH:mm:ss")
    strInput = strInput & "|" & Format(zlDataBase.Currentdate, "yyyy-MM-dd HH:mm:ss") '出院经办时间
    strInput = strInput & "|" & Switch(rsTemp("出院方式") = "正常", 0, rsTemp("出院方式") = "死亡", 1, rsTemp("出院方式") = "转院", 2, True, 9) & "#"
    
    Call WriteLog("DataUnloadint(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    
    出院登记_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 记帐传输_新都(strNO As String, int性质 As Integer, int状态 As Integer, Optional lng病人ID As Long) As Boolean
'功能：将住院病人的记帐单据上传到医保前置服务器
'参数：lng病人ID=是否只上传单据中指定病人的费用
    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim rsBill As New ADODB.Recordset, rsTemp As New ADODB.Recordset, rs收费类别 As New ADODB.Recordset
    Dim lng当前病人 As Long
    '费用传输使用的变量
    Dim str中心编号 As String, str就诊顺序号 As String, str人员类别 As String, str医院编码 As String
    Dim str流水号 As String, str收费类别 As String, str医保号 As String
    
    记帐传输_新都 = True '首先保证单据能得到保存。即使本次上传败，也可以在以后继续上传。
    On Error GoTo errHandle
    
    '列出所有收费类别
    gstrSQL = "Select 编码,类别 as 名称 From 收费类别"
    Call OpenRecordset(rs收费类别, gstrSysName)
    
    '读取单据明细(医保号,顺序号,登记时间,项目编码,项目名称,产地,规格,数量,单价,金额,医生,开单科室)
    '单据中非该医保的费用不传,未设置医保编码的不传,无顺序号的不传,婴儿费不上传。按病人排序
    gstrSQL = _
        "Select Nvl(A.价格父号,序号) as 序号," & _
        " A.病人ID,A.主页ID,F.医保号,F.顺序号,A.登记时间,D.项目编码,B.名称 as 项目名称,A.收费类别, " & _
        " Decode(Instr(B.规格,'┆'),0,B.规格,Substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
        " Decode(Instr(B.规格,'┆'),0,'',Substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
        " Avg(Nvl(A.付数,1)*A.数次) as 数量,Sum(A.标准单价) as 单价,Sum(A.实收金额) as 金额," & _
        " A.开单人 as 医生,C.名称 as 开单科室" & _
        " From 病人费用记录 A,收费细目 B,部门表 C,保险支付项目 D,病案主页 E,保险帐户 F" & _
        " Where A.记录状态<>0 And Nvl(A.是否上传,0)=0 And A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID" & _
        " And A.病人ID=E.病人ID And A.主页ID=E.主页ID And A.病人ID=F.病人ID" & _
        " And F.顺序号 is Not NULL And Nvl(A.婴儿费,0)=0" & _
        " And D.险类=" & gintInsure & " And E.险类=" & gintInsure & " And F.险类=" & gintInsure & _
        " And A.NO='" & strNO & "' And A.记录性质=" & int性质 & " And A.记录状态=" & int状态 & _
        IIf(lng病人ID = 0, "", " And A.病人ID=" & lng病人ID) & _
        " Group by Nvl(A.价格父号,序号),A.病人ID,A.主页ID,F.医保号,F.顺序号," & _
        " A.登记时间,D.项目编码,B.名称,A.收费类别,B.规格,A.开单人,C.名称" & _
        " Order by 病人ID,序号"
    rsBill.CursorLocation = adUseClient
    Call OpenRecordset(rsBill, "记帐传输")
    
    Do Until rsBill.EOF
        '记帐单中有多个病人,要分别处理
        If lng当前病人 <> rsBill("病人ID") Then
            '对该病人作相应的初始化工作-------------------------------------------------
            lng当前病人 = rsBill("病人ID")
            
            '得到入院审批信息（已经重新验证的）
            gstrSQL = "Select 医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别  " & _
                      "       ,NVL(A.本次起付线,0) as 住院起付线,NVL(A.起付线累计,0) as 基本统筹支付累计" & _
                      "       ,NVL(A.基本统筹限额,0) as 住院统筹限额,NVL(A.大额统筹限额,0) as 实足年龄,NVL(A.大额统筹累计,0) as 统筹报销比例" & _
                      "  From 帐户年度信息 A,病案主页 B,保险帐户 C " & _
                      "  where B.病人ID=" & lng当前病人 & " and B.主页ID=" & rsBill("主页ID") & " and A.病人ID=B.病人ID and A.险类=" & gintInsure & " and A.年度=to_char(B.入院日期,'yyyy')" & _
                      "     and C.病人ID=A.病人ID and C.险类=A.险类"
            Call OpenRecordset(rsTemp, "记帐传输")
            str就诊顺序号 = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))
            str中心编号 = IIf(IsNull(rsTemp("中心编号")), "", rsTemp("中心编号"))
            str医保号 = IIf(IsNull(rsTemp("医保号")), "", rsTemp("医保号"))
            str人员类别 = IIf(IsNull(rsTemp("人员类别")), "", rsTemp("人员类别"))
            
            If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
            If Get流水号("G", str医院编码, str流水号) = False Then Exit Function
        End If
            
        '进行费用分割
        Call WriteLog("DivideUp(" & str中心编号 & "," & ToVarchar(rsBill!项目编码, 12) & "," & "31" & "," & str人员类别 & "," & Val(rsBill!单价) & ")")
        strReturn = DivideUp(str中心编号, ToVarchar(rsBill("项目编码"), 12), "31", str人员类别, Val(rsBill("单价")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        rs收费类别.Filter = "编码 = '" & rsBill("收费类别") & "'"
        If rs收费类别.EOF = False Then str收费类别 = rs收费类别("名称")
        
        '第二个就诊顺序号是作为结算单号
        strInput = str就诊顺序号 & "|" & str就诊顺序号
        strInput = strInput & "|" & strNO & "_" & rsBill("序号") & "_" & int性质 & "_" & int状态  '序号
        strInput = strInput & "|" & str医保号 & "|" & str中心编号 & "|" & str医院编码 & "|000"
        strInput = strInput & "|" & ToVarchar(rsBill("项目编码"), 12)  '医保流水号
        strInput = strInput & "|" & ToVarchar(str收费类别, 10)      '收费大类名称
        strInput = strInput & "|" & Format(rsBill("数量"), "0.00")
        strInput = strInput & "|" & Format(rsBill("单价"), "0.00")
        strInput = strInput & "|" & Format(rsBill("金额"), "0.00")
        strInput = strInput & "|" & arrOutput(4)                       '自付比例
        strInput = strInput & "|" & Format(Val(arrOutput(1)) * rsBill("数量"), "#0.00") '全自费部分
        strInput = strInput & "|" & Format(Val(arrOutput(2)) * rsBill("数量"), "#0.00") '挂钩自费部分
        strInput = strInput & "|" & Format(Val(arrOutput(3)) * rsBill("数量"), "#0.00") '允许报销部分
        strInput = strInput & "||31"                                   '特项标志、支付类别
        strInput = strInput & "|" & ToVarchar(rsBill("开单科室"), 56)  '开单科室名称
        strInput = strInput & "|" & ToVarchar(rsBill("医生"), 20)      '开单处方医生
        strInput = strInput & "|" & ToVarchar(rsBill("开单科室"), 56)  '受单科室名称
        strInput = strInput & "|" & ToVarchar(rsBill("医生"), 20)      '受单处方医生
        strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)        '经办人
        strInput = strInput & "|" & Format(rsBill("登记时间") + rsBill("序号") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '经办时间
        strInput = strInput & "|" & ToVarchar(rsBill("项目名称"), 200)       '收费项目
        strInput = strInput & "|" & ToVarchar(rsBill("规格"), 200)       '规格
        strInput = strInput & "|" & ToVarchar(rsBill("产地"), 200)       '产地
        strInput = strInput & "|"                                        '单位
        strInput = strInput & "||"                                      '英文名、化学名
        'modify by ccy ,唯一
        strInput = strInput & Format(rsBill("登记时间"), "yyyyMMddHHmmss") & rsBill("序号") & "#"
        
        Call WriteLog("DataUnloading(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
        strReturn = DataUnloading(strInput, str流水号, str中心编号)
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        
        gstrSQL = "zl_病人费用记录_上传('" & strNO & "'," & rsBill("序号") & "," & int性质 & "," & int状态 & ")"
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        
        rsBill.MoveNext
    Loop
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院虚拟结算_新都(rsExse As Recordset, ByVal lng病人ID As Long, ByVal str医保号 As String) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset, rs收费类别 As New ADODB.Recordset

    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim str中心编号 As String, str就诊顺序号 As String, str人员类别 As String, str医院编码 As String
    Dim cur个人帐户 As Double, cur统筹支付 As Double
    Dim dbl总金额 As Double, dbl全自费 As Double, dbl挂钩自付 As Double, dbl允许报销 As Double
    Dim dbl住院起付线 As Double, dbl基本统筹支付累计 As Double, dbl住院统筹限额 As Double, lng实足年龄 As Long, dbl统筹报销比例 As Double
    Dim str医生 As String, datCurr As Date, str流水号 As String, str收费类别 As String
    
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
    End With
    
    '重新进行待遇审核
    Dim str卡号_New As String, str医保号_New As String, str中心编号_New As String, str密码_New As String
    If frmIdentify贵阳.GetIdentify(TYPE_新都, str卡号_New, str医保号_New, str中心编号_New, str密码_New) = False Then
        '身份验证未通过
        Exit Function
    End If
'    If str医保号 <> str医保号_New Then
'        MsgBox "该卡不是当前病人的，请检查一下。", vbInformation, gstrSysName
'        Exit Function
'    End If
    If 入院登记_新都(g结算数据.病人ID, g结算数据.主页ID, str医保号, False) = False Then
        Exit Function
    End If
    
    '得到入院审批信息（已经重新验证的）
    gstrSQL = "Select 医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别  " & _
              "       ,NVL(A.本次起付线,0) as 住院起付线,NVL(A.起付线累计,0) as 基本统筹支付累计" & _
              "       ,NVL(A.基本统筹限额,0) as 住院统筹限额,NVL(A.大额统筹限额,0) as 实足年龄,NVL(A.大额统筹累计,0) as 统筹报销比例" & _
              "  From 帐户年度信息 A,病案主页 B,保险帐户 C " & _
              "  where B.病人ID=" & lng病人ID & " and B.主页ID=" & g结算数据.主页ID & " and A.病人ID=B.病人ID and A.险类=" & gintInsure & " and A.年度=to_char(B.入院日期,'yyyy')" & _
              "     and C.病人ID=A.病人ID and C.险类=A.险类"
    Call OpenRecordset(rsTemp, "住院预算")
    str就诊顺序号 = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))
    str中心编号 = IIf(IsNull(rsTemp("中心编号")), "", rsTemp("中心编号"))
    str人员类别 = IIf(IsNull(rsTemp("人员类别")), "", rsTemp("人员类别"))
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    
    dbl住院起付线 = rsTemp("住院起付线")
    dbl基本统筹支付累计 = rsTemp("基本统筹支付累计")
    dbl住院统筹限额 = rsTemp("住院统筹限额")
    lng实足年龄 = rsTemp("实足年龄")
    dbl统筹报销比例 = rsTemp("统筹报销比例")
    
    '列出所有收费类别
    gstrSQL = "Select 编码,类别 as 名称 From 收费类别"
    Call OpenRecordset(rs收费类别, gstrSysName)
    '打开另外一个连接串，以达到不受当前连接事务的控制
    cn上传.ConnectionString = gcnOracle.ConnectionString
    cn上传.Open
    
    Screen.MousePointer = vbHourglass
    
    
    If Get流水号("G", str医院编码, str流水号) = False Then Exit Function
    Do Until rsExse.EOF
        '进行费用分割
        Call WriteLog("费用分割(" & str中心编号 & "," & ToVarchar(rsExse!医保项目编码, 12) & ",31," & str人员类别 & "," & Val(rsExse!价格) & ")")
        strReturn = DivideUp(str中心编号, ToVarchar(rsExse("医保项目编码"), 12), "31", str人员类别, Val(rsExse("价格")))
        If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
        
        dbl总金额 = dbl总金额 + rsExse("金额")
        dbl全自费 = dbl全自费 + Val(arrOutput(1)) * rsExse("数量")
        dbl挂钩自付 = dbl挂钩自付 + Val(arrOutput(2)) * rsExse("数量")
        dbl允许报销 = dbl允许报销 + Val(arrOutput(3)) * rsExse("数量")
        
'        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" Then
'            '只上传只传递过的数据
'            rs收费类别.Filter = "编码 = '" & rsExse("收费类别") & "'"
'            If rs收费类别.EOF = False Then str收费类别 = rs收费类别("名称")
'
'            '第二个就诊顺序号是作为结算单号
'            strInput = str就诊顺序号 & "|" & str就诊顺序号
'            strInput = strInput & "|" & rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") & "_" & rsExse("记录状态")  '序号
'            strInput = strInput & "|" & str医保号 & "|" & str中心编号 & "|" & str医院编码 & "|000"
'            strInput = strInput & "|" & ToVarchar(rsExse("医保项目编码"), 12)  '医保流水号
'            strInput = strInput & "|" & ToVarchar(str收费类别, 10)      '收费大类名称
'            strInput = strInput & "|" & Format(rsExse("数量"), "0.00")
'            strInput = strInput & "|" & Format(rsExse("价格"), "0.00")
'            strInput = strInput & "|" & Format(rsExse("金额"), "0.00")
'            strInput = strInput & "|" & arrOutput(4)                       '自付比例
'            strInput = strInput & "|" & Val(arrOutput(1)) * rsExse("数量") '全自费部分
'            strInput = strInput & "|" & Val(arrOutput(2)) * rsExse("数量") '挂钩自费部分
'            strInput = strInput & "|" & Val(arrOutput(3)) * rsExse("数量") '允许报销部分
'            strInput = strInput & "||31"                                   '特项标志、支付类别
'            strInput = strInput & "|" & ToVarchar(rsExse("开单部门"), 56)  '开单科室名称
'            strInput = strInput & "|" & ToVarchar(rsExse("医生"), 20)      '开单处方医生
'            strInput = strInput & "|" & ToVarchar(rsExse("开单部门"), 56)  '受单科室名称
'            strInput = strInput & "|" & ToVarchar(rsExse("医生"), 20)      '受单处方医生
'            strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)        '经办人
'            strInput = strInput & "|" & Format(rsExse("登记时间") + rsExse("序号") / 24 / 3600, "yyyy-MM-dd HH:mm:ss") '经办时间
'            strInput = strInput & "|" & ToVarchar(rsExse("收费名称"), 200)       '收费项目
'            strInput = strInput & "|" & ToVarchar(rsExse("规格"), 200)       '规格
'            strInput = strInput & "|" & ToVarchar(rsExse("产地"), 200)       '产地
'            strInput = strInput & "|"                                        '单位
'            strInput = strInput & "||"                                      '英文名、化学名
'            'modify by ccy ,唯一
'            strInput = strInput & Format(rsExse("登记时间"), "yyyyMMddHHmmss") & rsExse("序号") & "#"
'
'            Call WriteLog("DataUnloading(" & strInput & "," & str流水号 & "," & str中心编号 & ")")
'            strReturn = DataUnloading(strInput, str流水号, str中心编号)
'            If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
'
'
'            gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
'            cn上传.Execute gstrSQL, , adCmdStoredProc
'        End If
        
        rsExse.MoveNext
    Loop
    
    '调用预结算
    '2107,404.2,44020,0,37,0,0,1604,103,400,.824
    Call WriteLog("预结算:" & dbl总金额 & "," & dbl住院起付线 & "," & dbl住院统筹限额 & "," & dbl基本统筹支付累计 & "," & lng实足年龄 & "," & 0 & "," & 0 & "," & _
                dbl允许报销 & "," & dbl全自费 & "," & dbl挂钩自付 & "," & dbl统筹报销比例)
    strReturn = CalculateFeeCD(dbl总金额, dbl住院起付线, dbl住院统筹限额, dbl基本统筹支付累计, lng实足年龄, 0, 0, _
                dbl允许报销, dbl全自费, dbl挂钩自付, dbl统筹报销比例)
    Call WriteLog("返回:" & strReturn)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    cur统筹支付 = Val(arrOutput(2))
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = dbl总金额
        .实际起付线 = Val(arrOutput(1))
        .统筹报销金额 = cur统筹支付
        .超限自付金额 = Val(arrOutput(4))
    
        .进入统筹金额 = dbl允许报销
        .全自费金额 = dbl全自费
        .首先自付金额 = dbl挂钩自付
        .个人帐户支付 = Val(arrOutput(3)) '基本统筹自付部分
    End With
    
    住院虚拟结算_新都 = "医保基金;" & cur统筹支付 & ";0"
    
    mlng病人ID = lng病人ID  '表示该病人已经进行了虚拟结算
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_新都(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    
    On Error GoTo errHandle
    Dim rsTemp As New ADODB.Recordset, strInput As String, arrOutput  As Variant, str流水号 As String, strReturn As String
    Dim str医院类别 As String, str医院编码 As String, str医保号 As String
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, datCurr As Date, cur起付线 As Currency
    
    If mlng病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    
    datCurr = zlDataBase.Currentdate
    
    '得到入院审批信息
    gstrSQL = "Select 医保号,顺序号 as 就诊序号,退休证号 as 中心编号,人员身份 as 人员类别,C.单位编码  " & _
              " ,D.姓名,D.性别,D.出生日期 " & _
              " ,nvl(A.本次起付线,0) as 住院起付线,nvl(A.起付线累计,0) as 基本统筹支付累计,nvl(A.基本统筹限额,0) as 住院统筹限额,nvl(A.大额统筹限额,0) as 实足年龄,nvl(A.大额统筹累计,0) as 统筹报销比例" & _
              "  From 帐户年度信息 A,病案主页 B,保险帐户 C,病人信息 D " & _
              "  where B.病人ID=" & lng病人ID & " and B.主页ID=" & g结算数据.主页ID & " and A.病人ID=B.病人ID and A.险类=" & gintInsure & " and A.年度=to_char(B.入院日期,'yyyy')" & _
              "     and C.病人ID=A.病人ID and C.险类=A.险类   and B.病人ID=D.病人ID"
    Call OpenRecordset(rsTemp, "住院预算")
    If Get医院编码(str医院编码, rsTemp("中心编号")) = False Then Exit Function
    If Get医院编码(str医院类别, rsTemp("中心编号"), True) = False Then Exit Function
    
    cur起付线 = rsTemp("住院起付线")
    '调用结算
    If Get流水号("F", str医院编码, str流水号) = False Then Exit Function
    strInput = IIf(IsNull(rsTemp("就诊序号")), "", rsTemp("就诊序号"))  '就诊顺序号
    strInput = strInput & "|" & ToVarchar(rsTemp("中心编号"), 4)        '分中心编号
    strInput = strInput & "|" & ToVarchar(rsTemp("医保号"), 20)         '个人编码
    strInput = strInput & "|" & ToVarchar(rsTemp("单位编码"), 12)        '单位编码
    strInput = strInput & "|" & ToVarchar(rsTemp("姓名"), 20)            '姓名
    strInput = strInput & "|" & ToVarchar(IIf(rsTemp("性别") = "女", "2", "1"), 4)         '性别
    strInput = strInput & "|" & Format(rsTemp("出生日期"), "yyyy-MM-dd") '出生日期
    strInput = strInput & "|" & Format(rsTemp("实足年龄"), "0")         '实足年龄
    strInput = strInput & "|"                                           '缴费年限
    strInput = strInput & "|" & str医院编码
    strInput = strInput & "|000"                                        '分院编码
    strInput = strInput & "|" & str医院类别                             '医院类别
    strInput = strInput & "|31"                                         '支付类别
    strInput = strInput & "|0"                                          '特种病标志
    strInput = strInput & "|000"                                        '特种病编码
    strInput = strInput & "|" & ToVarchar(rsTemp("就诊序号"), 20)       '结算编号
    strInput = strInput & "|"                                           '退单编号
    strInput = strInput & "|" & ToVarchar(rsTemp("人员类别"), 20)       '医疗人员类别
    With g结算数据
        strInput = strInput & "|" & Format(cur起付线, "0.00")        '起付线
        strInput = strInput & "|" & Format(.发生费用金额, "0.00")    '费用总额
        strInput = strInput & "|" & Format(.全自费金额, "0.00")      '全自费部分
        strInput = strInput & "|" & Format(.首先自付金额, "0.00")    '挂钩自付部分
        strInput = strInput & "|" & Format(.进入统筹金额, "0.00")    '允许报销部分
        strInput = strInput & "|" & Format(.实际起付线, "0.00")      '进入起付线部分
        strInput = strInput & "|" & Format(.统筹报销金额, "0.00")    '基本医疗统筹支付部分
        strInput = strInput & "|" & Format(.个人帐户支付, "0.00")    '基本医疗统筹自付部分
        strInput = strInput & "|" & Format(0, "0.00")                '大额补助统筹支付部分
        strInput = strInput & "|" & Format(0, "0.00")                '大额补助统筹自付部分
        strInput = strInput & "|" & Format(.超限自付金额, "0.00")    '超限自付金额
        strInput = strInput & "|" & Format(0, "0.00")                '个人账户支付金额
    End With
    strInput = strInput & "|"                                              '发票号
    strInput = strInput & "|" & ToVarchar(UserInfo.姓名, 20)               '经办人
    strInput = strInput & "|" & Format(datCurr, "yyyy-MM-dd HH:mm:ss") & "#"    '经办时间

    Call WriteLog("DataUnloading(" & strInput & "," & str流水号 & "," & rsTemp!中心编号 & ")")
    strReturn = DataUnloading(strInput, str流水号, rsTemp("中心编号"))
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    
    '填写结算表
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(datCurr), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    With g结算数据
        gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(datCurr) & "," & _
            cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
            cur进入统筹累计 + .进入统筹金额 & "," & _
            cur统筹报销累计 + .统筹报销金额 & "," & int住院次数累计 & ")"
        Call ExecuteProcedure(gstrSysName)
        
        gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
            Year(datCurr) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
            cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & ",NULL," & .实际起付线 & "," & g结算数据.发生费用金额 & _
            "," & .全自费金额 & "," & .首先自付金额 & "," & .进入统筹金额 & "," & .统筹报销金额 & ",0," & .超限自付金额 & ",0,''," & .主页ID & ")"
        Call ExecuteProcedure(gstrSysName)
        
        '保险结算计算
        gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & .进入统筹金额 & "," & .统筹报销金额 & ",NULL)"
        Call ExecuteProcedure(gstrSysName)
    End With
        
    住院结算_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_新都(lng结帐ID As Long) As Boolean
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    
    住院结算冲销_新都 = False
End Function

Private Function Get医院编码(ByRef str医院编码 As String, ByVal str分中心编码 As String, Optional ByVal bln医院类别 As Boolean) As Boolean
'功能：得到医院的医保编码
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    
    On Error GoTo errHandle
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '首先将字串还原
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '支掉第一个增加的|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '支掉最后的#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(3) = str分中心编码 Then
                If bln医院类别 = True Then
                    strHospital = arrOutput(2) '医院类别
                Else
                    strHospital = arrOutput(0) '医院编码
                End If
            End If
        End If
    Next
    
    If strHospital <> "" Then
        str医院编码 = strHospital
        Get医院编码 = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get中心编码() As String
'功能：得到医院的医保编码
    Dim strReturn As String, arrOutput As Variant
    Dim strTemp As String, varList As Variant, lngIndex As Long, strHospital As String
    Dim str医院编码 As String, rsTmp As New ADODB.Recordset
        
    On Error GoTo errHandle
    '获取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_新都
    Call OpenRecordset(rsTmp, gstrSysName)
    
    If IsNull(rsTmp("医院编码")) = True Then
        MsgBox "由于未设置医院编号，无法执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    str医院编码 = rsTmp!医院编码
    
    strReturn = GetHospitalInfo()
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    '首先将字串还原
    strTemp = ""
    For lngIndex = 1 To UBound(arrOutput)
        strTemp = strTemp & "|" & arrOutput(lngIndex)
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2) '支掉第一个增加的|
    If Right(strTemp, 1) = "#" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1) '支掉最后的#
    
    varList = Split(strTemp, "$")
    
    For lngIndex = 0 To UBound(varList)
        arrOutput = Split(varList(lngIndex), "|")
        
        If UBound(arrOutput) > 3 Then
            If arrOutput(0) = str医院编码 Then
                Get中心编码 = arrOutput(3) '中心编码
                Exit For
            End If
        End If
    Next
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function JudgeReturn(ByRef strReturn As String, ByRef varOut As Variant) As Boolean
'功能：判断返回值是否合法。
    Dim varArray As Variant, lngReturn As Long, lngPos As Long
    Dim strSuggest
    
    strReturn = TruncZero(strReturn)
    lngPos = InStr(strReturn, "#")
    If lngPos > 0 Then
        strReturn = Mid(strReturn, 1, lngPos - 1)
    End If
    
    varArray = Split(strReturn, "|")
    
    lngReturn = Val(varArray(0))
    If lngReturn < 0 Then
        '业务调用失败
        If UBound(varArray) > 0 Then
            strReturn = "医保业务处理失败。" & vbCrLf & varArray(1)
        Else
            strReturn = "医保业务处理失败。"
        End If
        
        Select Case lngReturn
            Case -1101
                strSuggest = "可以重新身份识别并获取新的流水号。"
            Case -1102, -1210, -1216, -1404, -1405, -1502
                strSuggest = "需要银海公司检查。"
            Case -1201, -1203, -1204, -1205, -1207, -1213, -1215, -1217, -1220
                strSuggest = "需要到社保局确认。"
            Case -1208
                strSuggest = "可能在用条码卡，换成病人的磁卡重新刷。"
        End Select
        
        If strSuggest <> "" Then
            strReturn = strSuggest & vbCrLf & vbCrLf & "建议处理方法：" & strSuggest
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strReturn, vbExclamation, gstrSysName
        Exit Function
    End If
    
    varOut = varArray
    JudgeReturn = True
End Function

Private Function Get流水号(ByVal str标志 As String, ByVal str医院编码 As String, ByRef str流水号 As String) As Boolean
    Dim datCurr As Date
    
    datCurr = zlDataBase.Currentdate
    '[信息标志+医院编码+YYMMDD+6位流水号]
    str流水号 = str标志 & str医院编码 & Format(datCurr, "yyMMddHHmmss")
    Get流水号 = True
End Function

Public Function 医保项目_新都(rsTemp As ADODB.Recordset) As Boolean
'功能：医保诊疗药品目录查询
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim strPath As String, strFile As String, strReturn As String, arrOutput As Variant
    Dim lngFile  As Long, str中心编号 As String
    
    
    str中心编号 = Get中心编码
    If str中心编号 = "" Then Exit Function
    
    '调用接口，生成文件
    strFile = Space(255)
    GetTempPath 255, strFile
    strPath = TrimStr(strFile)
    strFile = strPath & "MakeTxt.txt"
    
    strReturn = MakeTxt(strFile, strPath & "Temp.txt") '病种目录虽然不要,但也必须传
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    
    On Error GoTo errHandle
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        
        arrOutput = Split(strReturn, vbTab)
        If UBound(arrOutput) >= 11 Then
            str编码 = arrOutput(0)
            str名称 = ToVarchar(arrOutput(1), 40)
            str简码 = ToVarchar(zlCommFun.SpellCode(arrOutput(1)), 10)
        End If
        If str编码 <> "" And arrOutput(11) = str中心编号 Then
            '只取当前中心的医保编码,其它中心的编码可能不同
            rsTemp.AddNew Array("CLASSCODE", "CODE", "NAME", "PY"), Array("1", str编码, str名称, str简码)
            rsTemp.Update
        End If
    Loop
    Close #lngFile
    Kill strFile
    Kill strPath & "Temp.txt"
    
    医保项目_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Close #lngFile
    
End Function

Public Function 卡解析_新都(ByVal str卡内容 As String, str医保号 As String, str卡号 As String, str中心编号 As String) As Boolean
'功能：进磁卡内容进行解析
    Dim strReturn  As String, arrOutput As Variant
    
    On Error GoTo errHandle
    
    If str卡内容 = "" Then
        MsgBox "请先进行刷卡操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    strReturn = GetKard(str卡内容)  '依次为医保号、卡号、医院编码、分中心编号
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    str医保号 = arrOutput(1)
    str卡号 = arrOutput(2)
    str中心编号 = arrOutput(3)
    卡解析_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 更改密码_新都(ByVal str卡号 As String, ByVal str医保号 As String, ByVal str中心编号 As String, _
            ByVal str原密码 As String, ByVal str新密码 As String) As Boolean

'功能：修改用户密码
    Dim strInput As String, arrOutput   As Variant, strReturn As String
    Dim str医院编码 As String, str流水号 As String
    
    On Error GoTo errHandle
    
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Function
    If Get流水号("B", str医院编码, str流水号) = False Then Exit Function
    
    strInput = str卡号 & "|" & str医保号 & "|" & str中心编号 & "|" & str原密码 & "|" & str新密码 & "#"
    
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Function
    
    MsgBox "新密码保存成功。", vbInformation, gstrSysName
    更改密码_新都 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub 核对帐户支付_新都(ByVal lng病人ID As Long)
    Dim int记录数_OUT As Integer, cur金额_OUT As Currency
    Dim int记录数_Client As Integer, cur金额_Client As Currency
    Dim lng主页ID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str中心编号 As String, str就诊序号 As String, str医院编码 As String, str医院类别 As String, str流水号 As String
    Dim rsTemp As New ADODB.Recordset
    '仅对出院病人进行检查
    On Error GoTo ErrHand
    
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "该病人还未出院！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '取上次住院的主页ID，因为该功能主要用于出院后使用，因此假定该病人未再次入院
    gstrSQL = "Select nvl(住院次数,1) 主页ID From 病人信息 Where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "最上次住院时的主页ID")
    lng主页ID = rsTemp!主页ID
    
    '取帐户支付记录数及支付金额
    gstrSQL = "Select Sum(A.冲预交) 帐户支付,Count(*) 记录数  " & _
             " From 病人预交记录 A, " & _
             "      (Select 病人ID,结帐ID  " & _
             "      From 病人费用记录 " & _
             "      Where 病人ID=1 And 主页ID=1) B " & _
             " Where A.结帐ID=B.结帐ID And A.结算方式='个人帐户'"
    Call OpenRecordset(rsTemp, "取帐户支付额及记录数")
    int记录数_Client = Nvl(rsTemp!记录数, 0)
    cur金额_Client = Nvl(rsTemp!帐户支付, 0)
    
    '获取基本信息
    gstrSQL = " Select 退休证号 中心编号,顺序号 就诊序号 From 保险帐户 " & _
            " Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_新都
    Call OpenRecordset(rsTemp, "获取基本信息")
    str就诊序号 = rsTemp!就诊序号
    str中心编号 = rsTemp!中心编号
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Sub
    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Sub
    
    '调用核对接口
    If Get流水号("H", str医院编码, str流水号) = False Then Exit Sub
    strInput = ToVarchar(str中心编号, 4)
    strInput = strInput & "|" & ToVarchar(str医院编码, 8)
    strInput = strInput & "|" & str就诊序号
    strInput = strInput & "|" & str就诊序号 & "|%#"
    
'    MsgBox "核对帐户支付：DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '如果与医保中心接收到的不符，给出提示（1-记录数;2-支付额）
    int记录数_OUT = arrOutput(1)
    cur金额_OUT = arrOutput(2)
    
    If Format(cur金额_OUT, "#####0.00;-#####0.00;0;") <> Format(cur金额_Client, "#####0.00;-#####0.00;0;") Then
        MsgBox "个人帐户支付额与医保中心返回的不一致，请检查！" & vbCrLf & _
               "本地实际帐户支付：" & cur金额_Client & String(4, " ") & "医保中心统计出的帐户支付：" & cur金额_OUT & vbCrLf & _
               "本地帐户支付次数：" & int记录数_Client & String(4, " ") & "医保中心统计出的支付次数：" & int记录数_OUT
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub 核对入出院_新都(ByVal lng病人ID As Long)
    '仅对出院病人进行检查
    Dim int记录数_OUT As Integer, cur金额_OUT As Currency
    Dim int记录数_Client As Integer, cur金额_Client As Currency
    Dim lng主页ID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str中心编号 As String, str就诊序号 As String, str医院编码 As String, str医院类别 As String, str流水号 As String
    Dim rsTemp As New ADODB.Recordset
    '仅对出院病人进行检查
    On Error GoTo ErrHand
    
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "该病人还未出院！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    int记录数_Client = 1
    
    '获取基本信息
    gstrSQL = " Select 退休证号 中心编号,顺序号 就诊序号 From 保险帐户 " & _
            " Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_新都
    Call OpenRecordset(rsTemp, "获取基本信息")
    str就诊序号 = rsTemp!就诊序号
    str中心编号 = rsTemp!中心编号
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Sub
    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Sub
    
    '调用核对接口
    If Get流水号("I", str医院编码, str流水号) = False Then Exit Sub
    strInput = ToVarchar(str中心编号, 4)
    strInput = strInput & "|" & ToVarchar(str医院编码, 8)
    strInput = strInput & "|" & str就诊序号
    strInput = strInput & "|" & str就诊序号 & "|#"
    
'    MsgBox "核对入出院记录：DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '如果与医保中心接收到的不符，给出提示（1-记录数）
    int记录数_OUT = arrOutput(1)
    
    If int记录数_OUT <> int记录数_Client Then
        MsgBox "病人入出院记录与医保中心返回的不一致，请检查！" & vbCrLf & _
               "病人入出院记录数：" & int记录数_Client & String(4, " ") & "医保中心统计出的入出院记录数：" & int记录数_OUT
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub 核对费用结算_新都(ByVal lng病人ID As Long)
    Dim int记录数_OUT As Integer, cur金额_OUT As Currency
    Dim cur起付线_OUT As Currency, cur全自费_OUT As Currency
    Dim cur首先自付_OUT As Currency, cur实际起付线_OUT As Currency
    Dim cur统筹支付_OUT As Currency, cur统筹自付_OUT As Currency
    Dim cur超限自付_OUT As Currency, cur帐户支付_OUT As Currency
    Dim int记录数_Client As Integer, cur金额_Client As Currency
    Dim cur起付线_Client As Currency, cur全自费_Client As Currency
    Dim cur首先自付_Client As Currency, cur实际起付线_Client As Currency
    Dim cur统筹支付_Client As Currency, cur统筹自付_Client As Currency
    Dim cur超限自付_Client As Currency, cur帐户支付_Client As Currency
    Dim lng主页ID As Long
    Dim strInput As String, strReturn As String, arrOutput
    Dim str中心编号 As String, str就诊序号 As String, str医院编码 As String, str医院类别 As String, str流水号 As String
    Dim rsTemp As New ADODB.Recordset
    '仅对出院病人进行检查
    On Error GoTo ErrHand
    
    If Not 医保病人已经出院(lng病人ID) Then
        MsgBox "该病人还未出院！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '取上次住院的主页ID，因为该功能主要用于出院后使用，因此假定该病人未再次入院
    gstrSQL = "Select nvl(住院次数,1) 主页ID From 病人信息 Where 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "最上次住院时的主页ID")
    lng主页ID = rsTemp!主页ID
    
    '取帐户支付记录数及支付金额
    gstrSQL = "SELECT SUM(发生费用金额) 发生费用,SUM(进入统筹金额) 进入统筹,SUM(统筹报销金额) 统筹报销, " & _
             " SUM(首先自付金额) 首先自付,SUM(起付线) 起付线,SUM(实际起付线) 实际起付线," & _
             " SUM(超限自付金额) 超限自付,SUM(个人帐户支付) 个人帐户支付,Count(*) 记录数 " & _
             " FROM  " & _
             "      (SELECT 病人ID,结帐ID FROM 病人费用记录 " & _
             "      WHERE 病人ID=" & lng病人ID & " AND 主页ID= " & lng主页ID & _
             "      ) A,保险结算记录 B " & _
             " WHERE A.病人ID=B.病人ID AND B.记录ID=A.结帐ID AND B.险类=" & TYPE_新都 & " AND B.性质=2 " & _
             " GROUP BY A.病人ID"
    Call OpenRecordset(rsTemp, "取帐户支付额及记录数")
    int记录数_Client = Nvl(rsTemp!记录数, 0)
    cur金额_Client = Nvl(rsTemp!发生费用, 0)
    cur统筹支付_Client = Nvl(rsTemp!统筹报销, 0)
    cur帐户支付_Client = Nvl(rsTemp!个人帐户支付, 0)
    
    '获取基本信息
    gstrSQL = " Select 退休证号 中心编号,顺序号 就诊序号 From 保险帐户 " & _
            " Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_新都
    Call OpenRecordset(rsTemp, "获取基本信息")
    str就诊序号 = rsTemp!就诊序号
    str中心编号 = rsTemp!中心编号
    If Get医院编码(str医院编码, str中心编号) = False Then Exit Sub
    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Sub
    
    '调用核对接口
    If Get流水号("J", str医院编码, str流水号) = False Then Exit Sub
    strInput = ToVarchar(str中心编号, 4)
    strInput = strInput & "|" & ToVarchar(str医院编码, 8)
    strInput = strInput & "|" & str就诊序号
    strInput = strInput & "|" & str就诊序号 & "|%|%|%#"
    
'    MsgBox "核对费用结算：DataUnloading" & strInput
    strReturn = DataUnloading(strInput, str流水号, str中心编号)
    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
    
    '如果与医保中心接收到的不符，给出提示（1-记录数;2-支付额）
    int记录数_OUT = arrOutput(1)
    cur起付线_OUT = arrOutput(2)
    cur金额_OUT = arrOutput(3)
    cur全自费_OUT = arrOutput(4)
    cur首先自付_OUT = arrOutput(5)
    'cur进入统筹_OUT = arrOutput(6)
    cur实际起付线_OUT = arrOutput(7)
    cur统筹支付_OUT = arrOutput(8)
    cur统筹自付_OUT = arrOutput(9)
    cur超限自付_OUT = arrOutput(10)
    cur帐户支付_OUT = arrOutput(11)
    
    '只要统筹支付、帐户支付及费用总额一致即可
    If Not (Format(cur金额_OUT, "#####0.00;-#####0.00;0;") = Format(cur金额_Client, "#####0.00;-#####0.00;0;") _
    And Format(cur统筹支付_OUT, "#####0.00;-#####0.00;0;") = Format(cur统筹支付_Client, "#####0.00;-#####0.00;0;") _
    And Format(cur帐户支付_OUT, "#####0.00;-#####0.00;0;") = Format(cur帐户支付_Client, "#####0.00;-#####0.00;0;")) Then
        MsgBox "本地结算数据与医保中心返回的不一致，请检查！" & vbCrLf & _
               "（医保）费用总额：" & cur金额_OUT & String(4, " ") & "统筹支付：" & cur统筹支付_OUT & String(4, " ") & "帐户支付：" & cur帐户支付_OUT & vbCrLf & _
               "（本地）费用总额：" & cur金额_Client & String(4, " ") & "统筹支付：" & cur统筹支付_Client & String(4, " ") & "帐户支付：" & cur帐户支付_Client
    Else
        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub 核对费用明细_新都(ByVal lng病人ID As Long)
'    Dim int记录数_OUT As Integer, cur金额_OUT As Currency
'    Dim int记录数_Client As Integer, cur金额_Client As Currency
'    Dim lng主页ID As Long
'    Dim strInput As String, strReturn As String, arrOutput
'    Dim str中心编号 As String, str就诊序号 As String, str医院编码 As String, str医院类别 As String, str流水号 As String
'    Dim rsTemp As New ADODB.Recordset
'    '仅对出院病人进行检查
'    On Error GoTo ErrHand
'
'    If Not 医保病人已经出院(lng病人ID) Then
'        MsgBox "该病人还未出院！", vbInformation, gstrSysName
'        Exit Sub
'    End If
'
'    '取上次住院的主页ID，因为该功能主要用于出院后使用，因此假定该病人未再次入院
'    gstrSQL = "Select nvl(住院次数,1) 主页ID From 病人信息 Where 病人ID=" & lng病人ID
'    Call OpenRecordset(rsTemp, "最上次住院时的主页ID")
'    lng主页ID = rsTemp!主页ID
'
'    '取帐户支付记录数及支付金额
'    gstrSQL = "Select Sum(A.冲预交) 帐户支付,Count(*) 记录数  " & _
'             " From 病人预交记录 A, " & _
'             "      (Select 病人ID,结帐ID  " & _
'             "      From 病人费用记录 " & _
'             "      Where 病人ID=1 And 主页ID=1) B " & _
'             " Where A.结帐ID=B.结帐ID And A.结算方式='个人帐户'"
'    Call OpenRecordset(rsTemp, "取帐户支付额及记录数")
'    int记录数_Client = NVL(rsTemp!记录数, 0)
'    cur金额_Client = NVL(rsTemp!帐户支付, 0)
'
'    '获取基本信息
'    gstrSQL = " Select 退休证号 中心编号,顺序号 就诊序号 From 保险帐户 " & _
'            " Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_新都
'    Call OpenRecordset(rsTemp, "获取基本信息")
'    str就诊序号 = rsTemp!就诊序号
'    str中心编号 = rsTemp!中心编号
'    If Get医院编码(str医院编码, str中心编号) = False Then Exit Sub
'    If Get医院编码(str医院类别, str中心编号, True) = False Then Exit Sub
'
'    '调用核对接口
'    If Get流水号("H", str医院编码, str流水号) = False Then Exit Sub
'    strInput = ToVarchar(str中心编号, 4)
'    strInput = strInput & "|" & ToVarchar(str医院编码, 8)
'    strInput = strInput & "|" & str就诊序号
'    strInput = strInput & "|" & str就诊序号 & "|%#"
'
'    strReturn = DataUnloading(strInput, str流水号, str中心编号)
'    If JudgeReturn(strReturn, arrOutput) = False Then Exit Sub
'
'    '如果与医保中心接收到的不符，给出提示（1-记录数;2-支付额）
'    int记录数_OUT = arrOutput(1)
'    cur金额_OUT = arrOutput(2)
'
'    If Format(cur金额_OUT, "#####0.00;-#####0.00;0;") <> Format(cur金额_Client, "#####0.00;-#####0.00;0;") Then
'        MsgBox "个人帐户支付额与医保中心返回的不一致，请检查！" & vbCrLf & _
'               "本地实际帐户支付：" & cur金额_Client & String(4, " ") & "医保中心统计出的帐户支付：" & cur金额_OUT & vbCrLf & _
'               "本地帐户支付次数：" & int记录数_Client & String(4, " ") & "医保中心统计出的支付次数：" & int记录数_OUT
'    Else
'        MsgBox "数据正确无误，核对成功！", vbInformation, gstrSysName
'    End If
'    Exit Sub
'ErrHand:
'    If ErrCenter = 1 Then Resume
End Sub

Private Sub WriteLog(ByVal strInfo As String)
    Dim strFileName As String
    Dim objSystem As FileSystemObject
    Dim objStream As TextStream
    
    If Val(GetSetting("ZLSOFT", "医保", "调试", 0)) = 0 Then Exit Sub
    strFileName = "C:\" & Format(Date, "yyyyMMdd") & ".txt"
    Set objSystem = New FileSystemObject
    If Not objSystem.FileExists(strFileName) Then Call objSystem.CreateTextFile(strFileName, False)
    Set objStream = objSystem.OpenTextFile(strFileName, ForAppending, False, TristateMixed)
    objStream.WriteLine (strInfo)
    objStream.Close
End Sub
