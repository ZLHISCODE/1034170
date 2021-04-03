Attribute VB_Name = "mdl大连"
Option Explicit
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--开发区接口
    '参数说明:
    '   msgType-业务请求类型,见以下的参数表
    '   packageType-数据解析格式类型，系统重组数据时使用,见以下的参数表
    '   packageLength-数据串的长度,见以下的参数表
    '   str-数据串,调用时，通过数据串传入参数；函数返回时，数据串中包含返回的数据
    '   strCom:数据请求串口（根据读卡器插口位置，本参数可以取值：'com1','com2')
    '返回:
    '   I.  当函数返回值等于0时，表示成功，字符串中包含了业务处理后返回的数据
    '   II. 当函数返回值不等于0时，参见错误代码一览表，应用需要分析错误代码然后进行适当的处理

'读卡器驱动程序
Private Declare Function IC_Read_Base Lib "ICCNII32.DLL" (ByVal szData As String) As Long
Private Declare Function IC_Read_Plus Lib "ICCNII32.DLL" (nSequence As Long, ByVal szData As String) As Long
    
Private Declare Function KfqTransData Lib "OltpTransKfq03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
    ByVal str As String, ByVal strCom As String) As Long
    
'--普通接口
Private Declare Function OltpTransData Lib "OltpTransIc03.dll" ( _
    ByVal msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
     ByVal str As String, ByVal strCom As String) As Long
'以下为开发区的参数表
'业务请求类型    数据解析格式类型    数据串最小长度         说明
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 95                     实时验卡（读卡、验卡）
'1002            12                  420                    实时结算
'1003            7                   297                    实时医疗明细数据提交
'1004            9                   136                    实时住院登记数据提交
'1006            12                  420                    实时结算预算
'1008            101                 95                     实时查询（直接查询中心数据）

'以下为大连市的参数表
'业务请求类型    数据解析格式类型    数据串最小长度         说明
'------------    ----------------    --------------         -----------------------------------------------
'1001            101                 94                     实时验卡（读卡、验卡）
'1002            12                  424                    实时结算
'1003            7                   230                    实时医疗明细数据提交
'1004            9                   206                    实时住院登记数据提交
'1006            12                  424                    实时结算预算
'1008            101                 94                     实时查询（直接查询中心数据）
'1005            8                   274                    实时医嘱传输
'1007            2                   55                     慢病帐户查询
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public gblnKFQCom_大连  As Boolean   'true-开发区接口,False-普通接口

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public g病人身份_大连 As 病人身份
Private Type 病人身份
    个人编号            As String
    姓名                As String
    性别                As String
    出生日期            As String
    年龄                As Integer
    身份证号            As String
    IC卡号              As Long
    治疗序号            As Long
    职工就医类别        As String
    基本个人帐户余额    As Double
    补助个人帐户余额    As Double
    统筹累计            As Double
    月缴费基数          As Double
    帐户状态            As String
    参保类别1           As String
    参保类别2           As String
    参保类别3           As String
    参保类别4           As String
    参保类别5           As String
    
    转诊单号            As String           '身份验证时输入
    医保中心            As Long             '身份验证时选择,保存的序号
    就诊分类            As Long             '身份验证时选择,保存的是结算方式代码
    支付金额            As Double           '
    诊断编码            As String           '诊断编码时输入,门诊有效
    诊断名称            As String           '诊断名称时输入,门诊有效
    
    补助帐户原始值      As Double          '慢病查询获取
    补助帐户当前值      As Double          '慢病查询获取
    慢病帐户状态        As Double          '慢病查询获取
    起付线              As Double
End Type

Public Const gbln模拟接口 = False     '模拟接口数据

Public gstr医院编码_大连 As String        '医院编码,只能为4位
Public gintComPort_大连 As Integer
Public gbln门诊明细时实上传 As Boolean
Public gbln住院明细时实上传 As Boolean

Private Function Read模拟数据(ByVal lng中心代码 As Long, _
        msgType As Long, ByVal packageType As Long, ByVal packageLength As Long, _
        str As String)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:通过该功能读取模拟数据,以例测试
    '--入参数:
    '--出参数:
    '--返  回:字串
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim strArr
    Dim strArr1
    Dim strText As String
    Dim strTemp As String
    
    If Dir(App.Path & "\大连医保\大连医保模拟数据" & lng中心代码 & ".txt") <> "" Then
            Set objText = objFile.OpenTextFile(App.Path & "\大连医保\大连医保模拟数据" & lng中心代码 & ".txt")
            Do While Not objText.AtEndOfStream
                strTemp = Trim(objText.ReadLine)
                strArr = Split(strTemp, "||")
                strArr1 = Split(strArr(0), "|")
                If Val(strArr1(0)) = msgType Then
                     str = strArr(1)
                     Exit Do
                End If
            Loop
            objText.Close
    End If
    
End Function
Public Function 读取病人身份_大连(ByVal lng中心代码 As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取病人的相关身份,并将信息赋给g病人身份_大连
    '--入参数:lng中心代码(2代表开发区)
    '--出参数:
    '--返  回:读取成功,返回True,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    
    Dim strInfor As String
    Dim lngReturn As Long
    Dim int性别 As Integer
    读取病人身份_大连 = False
    Err = 0
    On Error GoTo ErrHand:
    '周海全调试 2003-12-17
    '如果此处不传入空格值时，程序运行至此处会直接退出
    strInfor = Space(100)
    If gbln模拟接口 Then
        Read模拟数据 lng中心代码, 1001, 101, 94, strInfor
        If strInfor = "" Then Exit Function
    Else
        If lng中心代码 = 2 Then
            '1001    101 95  实时验卡（读卡、验卡）
            lngReturn = KfqTransData(1001, 101, 95, strInfor, "com" & gintComPort_大连)
        Else
            '1001    101 94  实时验卡（读卡、验卡）
            lngReturn = OltpTransData(1001, 101, 94, strInfor, "com" & gintComPort_大连)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn))
            Exit Function
        End If
    End If
    '取掉控格
    strInfor = Mid(strInfor, 2)
    With g病人身份_大连
        .医保中心 = lng中心代码
        If lng中心代码 = 2 Then
            .个人编号 = Substr(strInfor, 1, 10) '个人保号    1   10      中心返回
            .姓名 = Substr(strInfor, 11, 8)     '姓名    11  8       中心返回
            .身份证号 = Substr(strInfor, 19, 18)    '身份证号    19  18      中心返回
            .IC卡号 = Substr(strInfor, 37, 7)       'IC卡号  37  7       中心返回
            .治疗序号 = Val(Substr(strInfor, 44, 4))    '治疗序号    44  4       中心返回
            .职工就医类别 = Substr(strInfor, 48, 1)     '职工就医类别    48  1   A在职、B退休    中心返回
            .基本个人帐户余额 = Val(Substr(strInfor, 49, 10)) '基本个人帐户余额    49  10      中心返回
            .补助个人帐户余额 = Val(Substr(strInfor, 59, 10)) '补助个人帐户余额    59  10      中心返回
            .统筹累计 = Val(Substr(strInfor, 69, 10)) '统筹累计    69  10      中心返回
            .月缴费基数 = Val(Substr(strInfor, 79, 10)) '月缴费基数  79  10  月缴费工资  中心返回
            .帐户状态 = Substr(strInfor, 89, 1) '帐户状态    89  1   A正常、B半止付、C全止付、D销户  中心返回
            .参保类别1 = Substr(strInfor, 90, 1) '参保类别1   90  1   是否享受高额 1 享受 0 不享受    中心返回
            .参保类别2 = Substr(strInfor, 91, 1) '参保类别2   91  1   是否享受补助（商业补助、公务员补助）'0 不享受 1 商业 2 公务员    中心返回
            .参保类别3 = Substr(strInfor, 92, 1) '参保类别3   92  1   0 企保、1 事保  中心返回
            .参保类别4 = Substr(strInfor, 93, 1) '参保类别4   93  1   备用    中心返回
            .参保类别5 = Substr(strInfor, 94, 1) '参保类别5   94  1   备用    中心返回
        Else
            .个人编号 = Substr(strInfor, 1, 8)  '个人编号    CHAR    1   8   医保编号    中心
            .姓名 = Substr(strInfor, 9, 8)      '姓名    CHAR    9   8       中心
            .身份证号 = Substr(strInfor, 17, 18)    '身份证号    CHAR    17  18  18位或15位  中心
            .IC卡号 = Substr(strInfor, 35, 7)       'IC卡号  NUM 35  7       中心
            .治疗序号 = Val(Substr(strInfor, 42, 4))    '治疗序号    NUM 42  4       中心
            
            '周海全调试 2003-12-17
            '加入：Q企业公费
            .职工就医类别 = Substr(strInfor, 46, 1)     '职工就医类别    CHAR    46  1   A在职、B退休、L离休、T特诊、Q企业公费  中心
            .基本个人帐户余额 = Val(Substr(strInfor, 47, 10))   '基本个人帐户余额    NUM 47  10      中心
            .补助个人帐户余额 = Val(Substr(strInfor, 57, 10))   '补助个人帐户余额    NUM 57  10  现用于公务员单独列帐    中心
            .统筹累计 = Val(Substr(strInfor, 67, 10))   '统筹累计    NUM 67  10      中心
            .月缴费基数 = Val(Substr(strInfor, 77, 10)) '月缴费基数  NUM 77  10  月缴费工资  中心
            .帐户状态 = Substr(strInfor, 87, 1)         '帐户状态    CHAR    87  1   A正常、B半止付、C全止付、D销户  中心
            .参保类别1 = Substr(strInfor, 88, 1)        '参保类别1   CHAR    88  1   是否享受高额: 0 不享受高额、1 享受高额、2 医疗保险不可用    中心
            .参保类别2 = Substr(strInfor, 89, 1)        '参保类别2   CHAR    89  1   是否享受补助（商业补助、公务员补助）0 不享受 1 商业 2 公务员    中心
            .参保类别3 = Substr(strInfor, 90, 1)        '参保类别3   CHAR    90  1   0 企保、1 事保  中心
            .参保类别4 = Substr(strInfor, 91, 1)        '参保类别4   CHAR    91  1   0生育不可用、1生育可用  中心
            .参保类别5 = Substr(strInfor, 92, 1)        '参保类别5   CHAR    92  1   0工伤不可用、1工伤可用  中心
        End If
        int性别 = Val(IIf(Len(.身份证号) = 18, Mid(.身份证号, 17, 1), Right(.身份证号, 1))) Mod 2
        '根据身份证取出相应的性别
        .性别 = IIf(int性别 = 0, "女", "男")
        .出生日期 = zlCommFun.GetIDCardDate(Trim(.身份证号))
        '计算年龄
        If IsDate(.出生日期) And .出生日期 <> "" Then
            .年龄 = Abs(Int((zlDatabase.Currentdate - CDate(.出生日期)) / 365))
        Else
            .年龄 = 0
        End If
        
    End With
    读取病人身份_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    读取病人身份_大连 = False
End Function

Public Function 业务请求_大连( _
            ByVal lng中心代码 As Long, _
            ByVal lngMsgType As Long, _
            strTans As String _
    ) As Boolean
    
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对相关的业务请求,并返回相应的结果
    '--入参数:lng中心代码(2代表开发区)
    '   lngMsgType-业务请求类型
    '   lngPackageType-数据解析格式类型
    '   lngPackageLength-数据串的长度
    '   strTans-数据串,调用时，通过数据串传入参数；函数返回时，数据串中包含返回的数据
    '返回:
    '   成功-true,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngPackageType As Long
    Dim lngPackageLength As Long
    Dim i As Long
    Dim strTmp As String
    
    i = lngMsgType
    
    '以下为开发区的参数表
    '业务请求类型    数据解析格式类型    数据串最小长度         说明
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 95                     实时验卡（读卡、验卡）
    '1002            12                  420                    实时结算
    '1003            7                   297                    实时医疗明细数据提交
    '1004            9                   136                    实时住院登记数据提交
    '1006            12                  420                    实时结算预算
    '1008            101                 95                     实时查询（直接查询中心数据）
    
    '以下为大连市的参数表
    '业务请求类型    数据解析格式类型    数据串最小长度         说明
    '------------    ----------------    --------------         -----------------------------------------------
    '1001            101                 94                     实时验卡（读卡、验卡）
    '1002            12                  424                    实时结算
    '1003            7                   230                    实时医疗明细数据提交
    '1004            9                   206                    实时住院登记数据提交
    '1006            12                  424                    实时结算预算
    '1008            101                 94                     实时查询（直接查询中心数据）
    '1005            8                   274                    实时医嘱传输
    '1007            2                   55                     慢病帐户查询
    
    Dim strInfor As String
    Dim lngReturn As Long
    业务请求_大连 = False
    Err = 0
    On Error Resume Next
    If lng中心代码 = 2 Then
        strTmp = Switch(i = 1001, "101|95", i = 1002, "12|420", i = 1003, "7|297", i = 1004, "9|136", i = 1006, "12|420", _
            i = 1008, "101|95")
        If Err <> 0 Then
            strTmp = "|"
        End If
    Else
            strTmp = Switch(i = 1001, "101|94", i = 1002, "12|424", i = 1003, "7|230", i = 1004, "9|206", i = 1006, "12|424", _
                i = 1008, "101|94", i = 1005, "8|274", i = 1007, "2|55")
        If Err <> 0 Then
            strTmp = "|"
        End If
    End If
    lngPackageType = Val(Split(strTmp, "|")(0))
    lngPackageLength = Val(Split(strTmp, "|")(1))
    
    Err = 0
    On Error GoTo ErrHand:
    strInfor = strTans
    If gbln模拟接口 Then
        Read模拟数据 lng中心代码, lngMsgType, lngPackageType, lngPackageLength, strInfor
        If strInfor = "" Then
            strTans = strInfor
            Exit Function
        End If
    Else
        '因刘洋说,在所有业务类型的请求中都需前加空格.所以特加入:" " &
        strInfor = " " & strInfor
        If lng中心代码 = 2 Then
            lngReturn = KfqTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_大连)
        Else
            lngReturn = OltpTransData(lngMsgType, lngPackageType, lngPackageLength, strInfor, "com" & gintComPort_大连)
        End If
        If lngReturn <> 0 Or strInfor = "" Then
            ShowMsgbox GetErrInfo(CStr(lngReturn))
            strTans = ""
            Exit Function
        End If
    End If
    '取掉控格
    strInfor = Mid(strInfor, 2)
    
    strTans = strInfor
    业务请求_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    strTans = ""
    业务请求_大连 = False
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    
    Substr = Trim(StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode))
    Exit Function
ErrHand:
    Substr = ""
End Function

Public Function 医保初始化_大连() As Boolean

    Dim rsTemp  As New ADODB.Recordset
    Dim strReg As String
    
    '功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
    '返回：初始化成功，返回true；否则，返回false
    
    On Error Resume Next
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & gintInsure
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "获取医院编码")
    gstr医院编码_大连 = NVL(rsTemp!医院编码, "")
    
    '设置端口号
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)

    If Val(strReg) = 0 Then
        gintComPort_大连 = 1
    Else
        gintComPort_大连 = IIf(Val(strReg) > 99, 1, Val(strReg))
    End If
    
    Call GetRegInFor(g公共模块, "操作", "开发区", strReg)
    
    If gintInsure = TYPE_大连开发区 Then
        gblnKFQCom_大连 = True
    Else
        gblnKFQCom_大连 = False
    End If
    '设置上传明细参数
    gstrSQL = "Select * From 保险参数 where 参数名 in ('门诊明细时实上传','住院明细时实上传') and 险类=" & gintInsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险参数"
    gbln门诊明细时实上传 = True
    gbln住院明细时实上传 = True
    Do While Not rsTemp.EOF
        Select Case NVL(rsTemp!参数名)
        Case "门诊明细时实上传"
            gbln门诊明细时实上传 = IIf(Val(NVL(rsTemp!参数值)) = 1, True, False)
        Case "住院明细时实上传"
            gbln住院明细时实上传 = IIf(Val(NVL(rsTemp!参数值)) = 1, True, False)
        End Select
        rsTemp.MoveNext
    Loop
    医保初始化_大连 = True
End Function

Public Function 个人余额_大连(ByVal lng病人id As Long) As Currency
    '功能: 根据病人id取出余额
    '参数: 病人id
    '返回: 返回个人帐户余额
    Dim rsAcc As New ADODB.Recordset
    
    
    '读卡失败则退出
    gstrSQL = "Select Nvl(帐户余额,0) 帐户余额,退休证号 From 保险帐户 Where 险类=" & gintInsure
    gstrSQL = gstrSQL & " And 病人id=" & lng病人id
    
    Call OpenRecordset(rsAcc, "读取帐户余额")
    
    With g病人身份_大连
        .基本个人帐户余额 = NVL(rsAcc!帐户余额, 0)
        .补助个人帐户余额 = Val(NVL(rsAcc!退休证号))
        个人余额_大连 = .基本个人帐户余额
    End With
End Function

Public Function 医保设置_大连(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    医保设置_大连 = frmSet大连.ShowME(lng险类, lng医保中心)
End Function

Public Function 身份标识_大连(Optional bytType As Byte, Optional lng病人id As Long) As String
    Dim str备注 As String, rsPatient As New ADODB.Recordset
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊，1-住院
    '返回：空或信息串
    '注意：1)主要利用接口的身份识别交易；
    '      2)如果识别错误，在此函数内直接提示错误信息；
    '      3)识别正确，而个人信息缺少某项，必须以空格填充；
    
    身份标识_大连 = frmIdentify大连.GetPatient(bytType, lng病人id)
End Function
Public Function 身份标识_大连2(ByVal strCard As String, ByVal strPass As String, Optional lng病人id As Long) As String
    Dim lngReturn As Long
    Dim strNewPass As String
    '/**?
    身份标识_大连2 = frmIdentify大连.GetPatient(3, lng病人id)
End Function

Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = strTmp
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = Len(strCode)
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    End If
    Rpad = strTmp
End Function
Private Function Get就诊分类(ByVal byt业务 As Byte, ByVal int分类 As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取就诊分类标识
    '--入参数:byt业务-(0-结算,1-冲帐)
    '         int分类 门诊:(1-普通门诊,2-急诊门诊,3-门诊大病,4-门诊慢病补助)
    '                 住院:(5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院)
    '--出参数:
    '--返  回:医保中心的分类标识
    '-----------------------------------------------------------------------------------------------------------
    '医保中心的就诊分类的对应系统
    
    '1 门诊结算
    'A 门诊结算冲账
    '3 急诊结算
    '7 急诊结算冲账
    '5 门诊大病结算
    'B 门诊大病冲账
    'S 慢病补助结算
    'T 慢病补助冲帐
    
    '2 住院结算
    'D 住院冲冲账
    '9 住院冲补账  此功能暂不做
    '4 家庭病床结算
    'C 家庭病房冲冲账
    '8 家庭病房补账     '此功能暂不做
    'O 生育保险住院结算
    'P 生育保险住院冲帐
    'Q 工伤保险结算
    'R 工伤保险冲帐


    Dim i As Integer
    Dim strTmp As String
    i = int分类
    
    '刘兴宏标注:200404
    '     门诊:1-1,2-3,3-5,4-"S"
    '     住院:5-2,6-4,7-"O",8-"Q"
            
            
    Select Case int分类
        Case 1  '1-普通门诊
            strTmp = Decode(byt业务, 0, "1", "A")
        Case 2  '2-急诊门诊
            strTmp = Decode(byt业务, 0, "3", "7")
        Case 3  '3-门诊大病
            strTmp = Decode(byt业务, 0, "5", "B")
        Case 4  '4-门诊慢病补助
            strTmp = Decode(byt业务, 0, "S", "T")
        Case 5  '5-普通住院,
            strTmp = Decode(byt业务, 0, "2", "D")
        Case 6  '6-家庭病床住院,
            strTmp = Decode(byt业务, 0, "4", "C")
        Case 7  '7-生育保险住院
            strTmp = Decode(byt业务, 0, "O", "P")
        Case 8  '8-工伤保险住院
            strTmp = Decode(byt业务, 0, "Q", "R")
        Case Else
            strTmp = ""
    End Select
    Get就诊分类 = strTmp
End Function
Public Function 门诊虚拟结算_大连(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    Dim curTotal As Currency, cur个人帐户 As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs大类 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double
    Dim dbl草药费 As Double
    Dim dbl成药费 As Double
    Dim dbl西药费 As Double
    Dim dbl检查费 As Double
    Dim dbl治疗费 As Double
    Dim dbl大检费 As Double
    Dim dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double
    Dim dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double
    Dim dbl统筹比例 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl起付标准 As Double
    
    Dim lng病人id As Long
    
    Dim str诊断编码 As String  '疾病编码
    Dim str医师代码 As String
    Dim str操作员代码 As String
    Dim str诊断名称 As String
    Dim str治愈情况标识 As String
    Dim strTmp As String
    Dim str医生 As String
    Dim str明细 As String       '明细串
    Dim str国家编码 As String
    Dim dbl比例 As Double
    Dim str项目统计分类 As String
    Dim str项目编码 As String
    Dim dbl项目名称 As Double
    
    
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '明细字段
    '   病人ID,收费类别,收据费目,计算单位,开单人,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,摘要,是否急诊
    
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传
    
    '将保险支付大类读在本地,以便计算保费及自费
    gstrSQL = "Select * From 保险支付大类"
    zlDatabase.OpenRecordset rs大类, gstrSQL, "保险支付大类"
    
    Dim rs特准项目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim lng病种ID As Long
    With rs明细
        '确定病种
        If Not .EOF Then
            lng病人id = NVL(!病人ID, 0)
            gstrSQL = "  select 病种id from 保险帐户 where 病人id=" & lng病人id & "  and 险类=" & gintInsure & "  and 医保号='" & g病人身份_大连.个人编号 & "'"
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病种信息"
            If Not rsTemp.EOF Then
                lng病种ID = NVL(rsTemp!病种ID, 0)
            Else
                lng病种ID = 0
            End If
          '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
            
        End If
        
        '取出本次发生费用的金额合计
        Do While Not .EOF
            '---周顺利,对金额是否为负数进行判断,如果为负数不准执行医保收费
            If !实收金额 < 0 Then
                ShowMsgbox "该单据中包含有金额为负数的项目,不能执行医保收费!请检查后重新收费"
                门诊虚拟结算_大连 = False
                Exit Function
            End If
            
            If lng病种ID <> 0 Then
                    '第一步,确定允许的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & NVL(!收费细目ID, 0)
                    If rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & NVL(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        ShowMsgbox "收费细目为“" & NVL(rsTemp!名称) & "”的项目不是病种中所设定的项目."
                        Exit Function
                    End If
                    
                    '第二步,确定允许的保险大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & NVL(!保险支付大类ID, 0)
                    If rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了结算以外的保险支付大类,不能继续。"
                        Exit Function
                    End If
                    '第三步,'确定禁止的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & NVL(!收费细目ID, 0)
                    If Not rs特准项目.EOF Then
                        gstrSQL = "Select 编码,名称 from 收费细目 where id=" & NVL(!收费细目ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取收费细目"
                        ShowMsgbox "收费细目为“" & NVL(rsTemp!名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                        Exit Function
                    End If
                    '第四步,'确定禁止的大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & NVL(!保险支付大类ID, 0)
                    If Not rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了禁止使用的保险支付大类,不能继续。"
                    End If
            End If
        
            '先判断是否都设置了医保对应项目编码
            gstrSQL = " Select 项目编码,项目名称 From 保险支付项目" & _
                      " Where 险类=" & gintInsure & " And 收费细目ID=" & !收费细目ID
                      
            Call OpenRecordset(rsTemp, "判断是否设置了对应的医保项目")
            If rsTemp.EOF = True Then
                MsgBox "有项目未设置医保项目，不能结算。", vbInformation, gstrSysName
                Exit Function
            End If
            If str医生 = "" Then
                str医生 = NVL(!开单人)
            End If
            
            str项目编码 = NVL(rsTemp!项目编码)
            dbl项目名称 = Val(NVL(rsTemp!项目名称))
            lng病人id = NVL(!病人ID, 0)
            gstrSQL = "" & _
                " Select b.参数名,b.参数值 from 收费类别 a,保险参数 b " & _
                " Where a.类别=b.参数名 and b.险类=" & gintInsure & _
                "        and a.编码='" & NVL(!收费类别) & "'"
            
            zlDatabase.OpenRecordset rsTemp, gstrSQL, "保费计算"
            
            If rsTemp.EOF Then
                strTmp = ""
            Else
                strTmp = NVL(rsTemp!参数值)
            End If
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '计算保费
                rs大类.Find "id=" & NVL(!保险支付大类ID, 0), , adSearchForward, 1
                If Not rs大类.EOF Then
                    dbl统筹比例 = NVL(rs大类!统筹比额, 0) / 100
                Else
                    dbl统筹比例 = 1
                End If
                '中心为:A在职、B退休、L离休、T特诊,Q企业公费,我们默认为1在职、2退休、3离休、4特诊
                If gintInsure <> TYPE_大连开发区 And g病人身份_大连.职工就医类别 = "L" _
                    And g病人身份_大连.参保类别3 = "0" And NVL(!是否医保, 0) = 1 Then  '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '  大连市  企业单位离休医保：不完全执行医保政策，如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl统筹比例 = 1
                End If
                
                If gintInsure = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl统筹比例 = dbl项目名称
                End If
                
                If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl统筹比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                                
                '周海全调试 2003-12-17
                '对于特治项目，只要是标识为“特治”的，不应该再区分类别
                'If NVL(!收费类别) = "治疗" And str项目编码 = "特治" Then
                If str项目编码 = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If str项目编码 = "大检" Then
                    strTmp = "大检费"
                End If
                '计算扣除自费部分的费用
                If Not rsTemp.EOF Then
                    Select Case strTmp
                        Case "诊察费"
                            dbl诊察费 = dbl诊察费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        Case "草药费"
                            dbl草药费 = dbl草药费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        Case "成药费"
                            dbl成药费 = dbl成药费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        Case "西药费"
                            dbl西药费 = dbl西药费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        Case "检查费"
                            dbl检查费 = dbl检查费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        Case "治疗费"
                            dbl治疗费 = dbl治疗费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                        '周海全调试 2003-12-17
                        '大检费在医保参数设置中无法对应项目，这里是如何取得的？
                        Case "大检费"
                                 If gintInsure = TYPE_大连市 Then
                                       '---周顺利
                                       '大连市和开发区对大检费用处理不同,
                                       '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                       dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                                       
                                       If g病人身份_大连.职工就医类别 = "Q" Then
                                           '自费部分放入保险内自费费用中
                                       Else
                                           dbl大检自费 = dbl大检自费 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                                       End If
                                 Else
                                       dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0), 2)
                                       dbl大检自费 = dbl大检自费 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                                End If
                        Case "特殊治疗费"
                            '大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                            If gintInsure = TYPE_大连市 Then
                                dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!实收金额, 0), 2)
                            Else
                                dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!实收金额, 0) * dbl统筹比例, 2)
                            End If
                        
                            If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                                '自费部分放入 保险内自费费用中
                            Else
                                dbl特殊治疗自费 = dbl特殊治疗自费 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                            End If
                    End Select
                    If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                        '自费部分放入 保险内自费费用中
                        If dbl统筹比例 <> 0 Then
                            If !是否医保 = 1 Then
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                            End If
                        Else
                            '100自费部分放入非保险费用中
                            dbl非保险费用 = dbl非保险费用 + Round(NVL(!实收金额, 0), 2)
                        End If
                    Else
'                            If InStr(1, "567", NVL(!收费类别, 0)) <> 0 And Len(NVL(!收费类别)) = 1 Then
                                If gintInsure = TYPE_大连开发区 Then
                                    If !是否医保 = 1 And dbl统筹比例 <> 0 Then
                                        '险内药品自费  NUM 155 10  医保用药自费部分    院端填写
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                                    Else
                                        '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
                                        dbl其它费 = dbl其它费 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                                    End If
                                Else
                                    If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !是否医保 = 1 And dbl统筹比例 <> 0 Then
                                        '医保用药以及除了大检、特治外检查治疗项目的自费部分
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl统筹比例), 2)
                                    End If
                                    
                                    If !是否医保 <> 1 Or dbl统筹比例 = 0 Then
                                        '非医保用药以及诊疗项目
                                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!实收金额, 0), 2)
                                    End If
                                End If
 '                           End If
                        End If
                    End If
            End If
            curTotal = curTotal + Round(NVL(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
    
    '计算起付线
'    gstrSQL = "" & _
'        "   Select 本次起付线 " & _
'        "   From 帐户年度信息 " & _
'        "   where 险类=" & gintInsure & " and 病人ID=" & lng病人id & " and  年度=to_char(sysdate,'yyyy')"
'    zlDatabase.OpenRecordset rsTemp, gstrSQL, "计算起付线"
    
'    If rsTemp.EOF Then
'        dbl起付标准 = 0
'    Else
'        dbl起付标准 = NVL(rsTemp!本次起付线, 0)
'    End If
    If str医生 <> "" Then
        gstrSQL = "Select 编号 From 人员表  where 姓名='" & str医生 & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医生编号"
        If Not rsTemp.EOF Then
            str医生 = NVL(rsTemp!编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
        Else
            str医生 = ""
        End If
    End If
    '找出疾病编码
    str诊断编码 = g病人身份_大连.诊断编码
    str诊断名称 = g病人身份_大连.诊断名称
    With g病人身份_大连
        dbl起付标准 = .起付线
        If .医保中心 = 2 Then   '开发区
            strInfor = Lpad(gstr医院编码_大连, 6)       '医院代码
        Else
            strInfor = Lpad(gstr医院编码_大连, 4)       '医院代码
        End If
        strInfor = strInfor & " "      '子门诊标识
        If gintInsure = TYPE_大连开发区 Then     '开发区
            strInfor = strInfor & Lpad(.个人编号, 10)       '个人编号
        Else
            strInfor = strInfor & Lpad(.个人编号, 8)      '个人编号
        End If
        strInfor = strInfor & Lpad(.IC卡号, 7)       'IC卡号
        strInfor = strInfor & Lpad(.治疗序号 + 1, 4)      '治疗序号
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '结算时间
        strInfor = strInfor & String(10, " ") '病志号
        
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl诊察费, 2))), 10) '诊察费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl草药费, 2))), 10) '草药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl成药费, 2))), 10) '成药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl西药费, 2))), 10)  '西药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl检查费, 2))), 10)  '检查费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl治疗费, 2))), 10)   '治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl大检费, 2))), 10)   '大检费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl特殊治疗费, 2))), 10)   '特殊治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl大检自费, 2))), 10)   '大检自费
        If gintInsure = TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl特殊治疗自费, 2))), 10)    '特治自费    NUM 145 10      院端填写
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(dbl保险内自费费用, 2))), 10)    '保险内自费费用
        
        If gintInsure = TYPE_大连开发区 Then       '开发区
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl其它费, 2))), 10)    '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl非保险费用, 2))), 10)    '非保险费用
        End If
        
        strInfor = strInfor & String(10, " ")    '中心返回:结算后个人帐户余额;开发区:结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
        strInfor = strInfor & String(10, " ")    '中心返回:结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
            
        If gintInsure = TYPE_大连开发区 Then
                strInfor = strInfor & Lpad(.基本个人帐户余额, 10)  '结算前基本帐户余额  NUM 195 10  根据验卡返回结果    院端填写
                strInfor = strInfor & Lpad(Trim(CStr(.补助个人帐户余额)), 10)   '结算前补助账户余额  NUM 205 10  根据验卡返回结果    院端填写
                strInfor = strInfor & Lpad(Trim(CStr(.统筹累计)), 10)    '结算前统筹支付累计  NUM 215 10  基本统筹累计＋补充统筹累计根据验卡返回结果 院端填写
        Else
            '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
            If Get就诊分类(0, .就诊分类) = "S" Then
                strInfor = strInfor & Lpad(.补助帐户当前值, 10)   '结算前基本帐户余额
                strInfor = strInfor & Lpad("0", 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInfor = strInfor & Lpad("0", 10)   '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            Else
                strInfor = strInfor & Lpad(.基本个人帐户余额, 10)  '结算前基本帐户余额
                strInfor = strInfor & Lpad(Trim(CStr(.补助个人帐户余额)), 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInfor = strInfor & Lpad(Trim(CStr(.统筹累计)), 10)    '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            End If
        End If
        
        strInfor = strInfor & String(10, " ")    '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
        strInfor = strInfor & String(10, " ")    '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
        strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹支付
        strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹自付
        strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹支付
        strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹自付
        strInfor = strInfor & String(10, " ")    '中心返回:本次基本补助保险支付 ；开发区:公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
        strInfor = strInfor & String(10, " ")    '中心返回:本次非基本补助保险支付；开发区:公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
        strInfor = strInfor & String(10, " ")    '中心返回:本次保险范围外自付；开发区:限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
        
        If gintInsure <> TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(Trim(CStr(dbl特殊治疗自费)), 10)    '本次特殊治疗自付
        End If
        
        strInfor = strInfor & Lpad(Trim(CStr(dbl起付标准)), 10)    '起付标准；开发区:本次住院门槛费  NUM 315 10      院端填写
        strInfor = strInfor & Lpad(.转诊单号, 6)     '转诊单号
        strInfor = strInfor & Lpad(Get就诊分类(0, .就诊分类), 1)     '就诊分类
        If gintInsure <> TYPE_大连开发区 Then
            
            strInfor = strInfor & Lpad(.参保类别3, 1)    '参保类别3:0 企保、1 事保，根据验卡结果
        End If
        strInfor = strInfor & Lpad(.职工就医类别, 1)       '职工就医类别
        
        strInfor = strInfor & Lpad(.诊断编码, 16)    '诊断编码
        
        strInfor = strInfor & Lpad(str医生, 6)    '医师代码
        strInfor = strInfor & Lpad(UserInfo.编号, 6)    '操作员代码
        strInfor = strInfor & Lpad(.诊断名称, 30)    '诊断名称
        'A-治愈、B-好转、C-未愈、D-死亡、E-其他
        strInfor = strInfor & "A"    '治愈情况标识
        strInfor = strInfor & String(8, " ")      '出院日期
        
        If gintInsure = TYPE_大连开发区 Then       '开发区
        Else
            strInfor = strInfor & String(16, " ")      '传输时间
        End If
        strInfor = strInfor & String(10, " ")      '错误代码
    End With
    
    '调用虚拟接口(1006    12  423   实时结算预算
    门诊虚拟结算_大连 = 业务请求_大连(IIf(gintInsure = TYPE_大连开发区, 2, 1), 1006, strInfor)
    If 门诊虚拟结算_大连 = False Then
        Exit Function
    End If
    
    '开发区:
    '    本次基本个人帐户支付    NUM 225 10      中心返回
    '    本次补助个人帐户支付    NUM 235 10      中心返回
    '    本次基本统筹支付    NUM 245 10      中心返回
    '    本次基本统筹自付    NUM 255 10      中心返回
    '    本次补充统筹支付    NUM 265 10      中心返回
    '    本次补充统筹自付    NUM 275 10      中心返回
    '    本次基本补助保险支付    NUM 285 10  公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
    '    本次非基本补助保险支付  NUM 295 10  公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
    '    本次保险范围外自付  NUM 305 10  限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
    '大连市:
    '    本次基本个人帐户支付    NUM 211 10  如果是慢病结算，表示慢病帐户支付    中心
    '    本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0 中心
    '    本次基本统筹支付    NUM 231 10      中心
    '    本次基本统筹自付    NUM 241 10      中心
    '    本次补充统筹支付    NUM 251 10  如果是生育结算，本字段用于存放生育保险支付  中心
    '    本次补充统筹自付    NUM 261 10      中心
    '    本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付 2． 如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    '    本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付   2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分    中心
    '    本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    Dim i As Long
    If gintInsure = TYPE_大连开发区 Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    '确定本次结算方式
    str结算方式 = "个人帐户;" & Format(Val(Substr(strInfor, i + 10, 10)), "###0.00;-###0.00;0;0") & ";0" '本次基本个人帐户支付,不充许修改
    str结算方式 = str结算方式 & "|" & "补助帐户;" & Format(Val(Substr(strInfor, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "基本统筹;" & Format(Val(Substr(strInfor, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "补充统筹;" & Format(Val(Substr(strInfor, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "补助保险;" & Format(Val(Substr(strInfor, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "非补助保险;" & Format(Val(Substr(strInfor, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    
    门诊虚拟结算_大连 = True
End Function

Public Function 门诊结算_大连(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
    Dim lng病人id As Long
    门诊结算_大连 = Set门诊结算或冲销(False, lng结帐ID, cur个人帐户, lng病人id, strSelfNo)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    门诊结算_大连 = False
End Function
Private Function Set门诊结算或冲销(ByVal bln冲销 As Boolean, lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long, strSelfNo As String) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim curTotal As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double
    Dim dbl草药费 As Double
    Dim dbl成药费 As Double
    Dim dbl西药费 As Double
    Dim dbl检查费 As Double
    Dim dbl治疗费 As Double
    Dim dbl大检费 As Double
    Dim dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double
    Dim dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double
    Dim dbl统筹比例 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl起付标准 As Double
    Dim dbl比例 As Double
    Dim str医生 As String
    Dim str明细 As String       '明细串
    Dim str国家编码 As String
    Dim str项目编码 As String
    Dim str项目统计分类 As String
    Dim strTmp As String
    Dim int业务 As Integer
    Dim lng冲销ID As Long
    Dim strNO As String
    Dim lng记录性质 As Long
    
    Dim dbl个人帐户余额 As Double
    Dim dbl统筹支付累计 As Double
    Dim dbl个人帐户支付 As Double
    Dim dbl补助帐户支付 As Double
    Dim dbl基本统筹支付 As Double
    Dim dbl基本统筹自付 As Double
    Dim dbl补充统筹支付 As Double
    Dim dbl补充统筹自付 As Double
    Dim dbl补助保险支付 As Double
    Dim dbl非补助保险支付 As Double
    Dim dbl保险范围外自付 As Double
    
    Dim dbl结算前基本帐户余额  As Double
    Dim dbl结算前补助账户余额  As Double
    Dim dbl结算前统筹累计  As Double
    Dim lngTmp As Long
    Dim rs特准项目 As New ADODB.Recordset
    Dim lng病种ID As Long
    
    int业务 = IIf(bln冲销, 1, 0)
     Set门诊结算或冲销 = False
   
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '个人帐户可以支付全自费、首先自付部分，因此，只要卡上有足够的金额，可以全部使用个人帐户支付
    '注意：接口规定，门诊明细需结算后上传；住院明细需预结算时上传，如果卡内金额不足，可以使用圈存接口，即将卡外的钱，调到卡内，以增加卡内金额
    '卡内余额需要通过卡操作函数读取，可圈存金额是接口返回，需要修改
    
    On Error GoTo ErrHand
    '重新读卡
    If 读取病人身份_大连(IIf(gintInsure = TYPE_大连开发区, 2, 1)) = False Then
        Exit Function
    End If
    If bln冲销 Then
        '验证是否为该病人的IC卡
        gstrSQL = "Select * From  保险帐户 where 病人id=" & lng病人id
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
        If rsTemp.EOF Then
            ShowMsgbox "该病人在保险帐户中无记录!"
            Exit Function
        End If
        
        If g病人身份_大连.IC卡号 <> NVL(rsTemp!卡号) Then
            ShowMsgbox "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
            Exit Function
        End If
        '确定就诊分类,转诊单号,诊断编码,诊断名称
        ' 支付顺序号_IN(就诊分类;转诊单号;诊断编码),备注(诊断名称_IN)
        gstrSQL = "Select 支付顺序号,备注 from 保险结算记录  where 记录ID=" & lng结帐ID
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取就诊分类"
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "在结算记录中无结算记录!"
            Exit Function
        End If
        Dim strArr
        strArr = Split(NVL(rsTemp!支付顺序号), ";")
        
        '就诊分类;转诊单号;诊断编码
        '1-普通门诊("1", "A"),2-急诊门诊("3", "7")
        '3-门诊大病("5", "B"),4-门诊慢病补助("S", "T")
        If UBound(strArr) >= 2 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
            g病人身份_大连.诊断编码 = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
            g病人身份_大连.转诊单号 = strArr(1)
        Else
            g病人身份_大连.就诊分类 = Decode(strArr(0), "1", 1, "A", 1, "3", 2, "7", 2, "5", 3, "B", 3, 4)
        End If
        g病人身份_大连.诊断名称 = NVL(rsTemp!备注)
        
        
        '确定退费记录
        '退费
          gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
                    " where A.NO=B.NO and A.记录性质=B.记录性质  and A.记录状态=2 and B.结帐ID=" & lng结帐ID
          Call OpenRecordset(rsTemp, "门诊退费")
          If rsTemp.EOF Then
            ShowMsgbox "不存在病人费用冲销记录!"
            Exit Function
          Else
            lng冲销ID = rsTemp("结帐ID")
          End If
          
    End If
    '打开本次结算明细记录
    gstrSQL = " " & _
        "  Select Rownum 标识号,A.ID,A.病人ID,A.收费细目id,A.NO,A.序号,A.记录性质,A.记录状态,A.登记时间,A.开单人 as 医生,H.编号 as 医生编号, " & _
        "      A.数次*A.付数 as 数量,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 as 实收金额,F.参数值,G.id as 大类id,G.统筹比额, " & _
        "      A.收费类别,B.编码 as 项目编码,B.名称 as 项目名称,B.标识子码||B.标识主码 as 国家编码, " & _
        "      D.项目编码 医保编码,D.项目名称 as 医保名称,J.名称 as 剂型,D.是否医保,C.名称 开单部门,E.名称 受单部门, " & _
        "      L.险类,L.中心,L.卡号,L.医保号,L.人员身份,L.单位编码,L.顺序号,L.退休证号,L.帐户余额,L.当前状态,L.病种ID,L.在职,L.年龄段,L.灰度级,L.就诊时间 " & _
        "  From (Select * From 病人费用记录 Where 结帐ID=" & IIf(bln冲销, lng冲销ID, lng结帐ID) & " and  Nvl(附加标志,0)<>9 ) A,收费细目 B,部门表 C,保险支付项目 D,部门表 E,  " & _
        "       (Select U.*,K.参数值 From 收费类别 U,保险参数 K where U.类别=K.参数名 and K.险类=" & gintInsure & "  ) F, " & _
        "       (Select distinct Q.药品id,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J, " & _
        "       保险支付大类 G,人员表 H,保险帐户 L" & _
        "  Where A.收费细目ID=B.ID And A.开单部门ID=C.ID(+) and A.病人id=L.病人id and L.险类=" & gintInsure & " and A.收费类别=F.编码(+)  and d.大类id=G.id and a.收费细目id=J.药品id(+) " & _
        "        And A.执行部门ID=E.ID(+) And A.收费细目ID=D.收费细目ID And D.险类= " & gintInsure & " and a.开单人=H.姓名(+) " & _
        "  Order by A.ID"
        
    '上传费用明细记录
    zlDatabase.OpenRecordset rs明细, gstrSQL, "读取本次结帐费用明细"
    
    With rs明细
        If Not .EOF Then
            lng病人id = NVL(!病人ID, 0)
            str医生 = NVL(!医生编号)
            If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                str医生 = Substr(str医生, 1, 6)
            End If
            lng病种ID = NVL(!病种ID, 0)
            '打开特准项目
            gstrSQL = "Select * from 保险特准项目  where 病种ID=  " & lng病种ID
            zlDatabase.OpenRecordset rs特准项目, gstrSQL, "获取病种项目数据"
        End If
        Do While Not .EOF
            If lng病种ID <> 0 Then
                    '第一步,确定允许的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=1 and 收费细目id=" & NVL(!收费细目ID, 0)
                    If rs特准项目.EOF Then
                        ShowMsgbox "收费细目为“" & NVL(!项目名称) & "”的项目不是病种中所设定的项目."
                        Exit Function
                    End If
                    '第二步,确定允许的保险大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=1 and  收费细目id=" & NVL(!大类ID, 0)
                    If rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了结算以外的保险支付大类,不能继续。"
                        Exit Function
                    End If
                    '第三步,'确定禁止的收费细目
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=0 And 性质=2 and 收费细目id=" & NVL(!收费细目ID, 0)
                    If Not rs特准项目.EOF Then
                        ShowMsgbox "收费细目为“" & NVL(!项目名称) & "”的项目是被禁止使用的项目." & vbCrLf & "不能继续!"
                        Exit Function
                    End If
                    '第四步,'确定禁止的大类
                    rs特准项目.Filter = 0
                    rs特准项目.Filter = "大类=1 And 性质=2 and 收费细目id=" & NVL(!大类ID, 0)
                    If Not rs特准项目.EOF Then
                        ShowMsgbox "在结算中存在了禁止使用的保险支付大类,不能继续。"
                    End If
            End If
            strTmp = NVL(!参数值)
            lng病人id = NVL(!病人ID, 0)
            '确定相关数据
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str项目统计分类 = ""
                Else
                    str项目统计分类 = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '比例
                '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
                    
                If NVL(!险类, 0) <> TYPE_大连开发区 And Val(NVL(!单位编码, "99")) = 0 And NVL(!在职, 0) = 3 And NVL(!是否医保, 0) = 1 Then   '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '大连市    企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                Else
                    dbl比例 = NVL(!统筹比额, 0) / 100
                End If
                
                If NVL(!险类, 0) = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(NVL(!医保名称))
                End If
                
                If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl统筹比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                
                If NVL(!医保编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If NVL(!医保编码) = "大检" Then
                    strTmp = "大检费"
                End If
                
                Select Case strTmp
                    Case "诊察费"
                            dbl诊察费 = dbl诊察费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "草药费"
                           dbl草药费 = dbl草药费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "成药费"
                            dbl成药费 = dbl成药费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "西药费"
                        dbl西药费 = dbl西药费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "检查费"
                        dbl检查费 = dbl检查费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "治疗费"
                        dbl治疗费 = dbl治疗费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                    Case "大检费"
                          If gintInsure = TYPE_大连市 Then
                                '---周顺利
                                '大连市和开发区对大检费用处理不同,
                                '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                                
                                If g病人身份_大连.职工就医类别 = "Q" Then
                                    '自费部分放入保险内自费费用中
                                Else
                                    dbl大检自费 = dbl大检自费 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                                End If
                          Else
                                dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0), 2)
                                dbl大检自费 = dbl大检自费 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                         End If
'
'                        If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
'                            '自费部分放入 保险内自费费用中
'                        Else
'                            dbl大检自费 = dbl大检自费 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
'                        End If
                    Case "特殊治疗费"
                        '大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                        If gintInsure = TYPE_大连市 Then
                            dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!实收金额, 0), 2)
                        Else
                            dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                        End If
                        If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                            '自费部分放入 保险内自费费用中
                        Else
                            dbl特殊治疗自费 = dbl特殊治疗自费 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                        End If
                End Select
                If gintInsure = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '自费部分放入 保险内自费费用中
                    If dbl统筹比例 <> 0 Then
                        If !是否医保 = 1 Then
                            dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                        End If
                    Else
                        '100自费部分放入非保险费用中
                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!实收金额, 0), 2)
                    End If
                Else
'                        If InStr(1, "567", NVL(!收费类别, 0)) <> 0 And Len(NVL(!收费类别)) = 1 Then
                        If gintInsure = TYPE_大连开发区 Then
                            If !是否医保 = 1 And dbl比例 <> 0 Then
                                '险内药品自费  NUM 155 10  医保用药自费部分    院端填写
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                            Else
                                '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
                                dbl其它费 = dbl其它费 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                            End If
                        Else
                            If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !是否医保 = 1 And dbl比例 <> 0 Then
                                '医保用药以及除了大检、特治外检查治疗项目的自费部分
                                dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!实收金额, 0) * (1 - dbl比例), 2)
                            End If
                                            
                            '要求是100%的自费部将放入非保险费用中
                            If !是否医保 <> 1 Or dbl比例 = 0 Then
                                '非医保用药以及诊疗项目
                                dbl非保险费用 = dbl非保险费用 + Round(NVL(!实收金额, 0), 2)
                            End If
                          End If
    '                    End If
               End If
            Else
                dbl比例 = 1
                str项目统计分类 = ""
            End If

            '上传明细记录,实时医疗明细数据
            '参数控制明细上传
            If gbln门诊明细时实上传 Then
                
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                        str明细 = str明细 & Lpad(NVL(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
                    Else
                        str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                        str明细 = str明细 & Lpad(NVL(!医保号), 8)   '个人编号    CHAR    5   8       院端
                    End If
                
                    str明细 = str明细 & Space(10)   '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
                    str明细 = str明细 & Lpad(NVL(!顺序号, 0), 4)   '治疗序号    NUM 23  4   住院明细：必须等于入院登记时治疗序号门诊明细:                         必须等于本次结算治疗序号 院端
                    str明细 = str明细 & Lpad(NVL(!NO, 0), 10)       '处方号  NUM 27  10      院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                    Else
                        str明细 = str明细 & Lpad(CStr(.AbsolutePosition), 10)       '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
                    End If
                    '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
                    str明细 = str明细 & Space(10)       '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
                    
                    str明细 = str明细 & Get就诊分类(int业务, NVL(!灰度级, 0))         '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        '开发区为就诊时间    DATETIME    52  16  精确到秒（开处方时间）格式为：yyyymmddhhmiss后面以空格补位  院端填写
                        str明细 = str明细 & Rpad(Format(!就诊时间, "yyyymmddHHmmss"), 16)
                    Else
                        str明细 = str明细 & Rpad(Format(!登记时间, "yyyymmddHHmmss"), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
                    End If
                    
                    str明细 = str明细 & Lpad(NVL(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
                    str明细 = str明细 & Lpad(NVL(!项目名称), 20)      '项目名称    CHAR    94  20      院端
        
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                    Else
        
                        If !是否医保 = 1 Then
                            str明细 = str明细 & Lpad(1 - dbl比例, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                        Else
                            str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                        End If
                        str明细 = str明细 & Lpad(str项目统计分类, 1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    End If
                    str明细 = str明细 & Lpad(NVL(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
                    str明细 = str明细 & Lpad(NVL(!实际价格), 8) '单价    NUM 127 8   不允许出现负值  院端
                    str明细 = str明细 & Lpad(NVL(!计算单位), 4) '单位    CHAR    135 4       院端
                    str明细 = str明细 & Lpad(NVL(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        '获取病人单量等.
                        gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & NVL(!ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                        If rsTemp.EOF Then
                            str明细 = str明细 & Space(5)       '每次用量    NUM 146 5       院端填写
                            str明细 = str明细 & Space(20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                            str明细 = str明细 & Space(50)      '用法    CHAR    171 50  如：口服    院端填写
                        Else
                            str明细 = str明细 & Lpad(NVL(rsTemp!单量), 5)      '每次用量    NUM 146 5       院端填写
                            str明细 = str明细 & Lpad(NVL(rsTemp!频次), 20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                            str明细 = str明细 & Lpad(NVL(rsTemp!用法), 50)      '用法    CHAR    171 50  如：口服    院端填写
                        End If
                        str明细 = str明细 & Space(4)      '执行天数    NUM 221 4       院端填写
                        str明细 = str明细 & Lpad(NVL(!医生编号), 6)      '医师编码    CHAR    225 6       院端填写
                    Else
                        str明细 = str明细 & Lpad(NVL(!医生), 8)      '医师姓名    CHAR    159 8       院端
                    End If
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断编码, 16)      '诊断编码    CHAR    167 16      院端
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断名称, 30)     '诊断名称    CHAR    183 30      院端
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = str明细 & Lpad(NVL(!开单部门), 20)    '科别名称    CHAR    277 20      院端填写
                    Else
                        str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
                    End If
                    
                
                '上传明细
                '1003    7   230 实时医疗明细数据提交
                Set门诊结算或冲销 = 业务请求_大连(IIf(NVL(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细)
                If Set门诊结算或冲销 = False Then
                    ShowMsgbox "门诊结算时医疗明细数据提交失败,不能继续!"
                    Exit Function
                End If
                '为病人费用记录打上标记，以便随时上传
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                gstrSQL = "ZL_病人费用记录_更新医保(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            '计算总额,待用
            curTotal = curTotal + Round(NVL(!实收金额, 0), 2)
            .MoveNext
        Loop
    End With
    Set门诊结算或冲销 = False
    '冲销时,重新获取病人的相关信息.
'    If bln冲销 Then
'        Call 获取病人信息_大连(lng病人id)
'    End If
    '计算起付线
        dbl起付标准 = g病人身份_大连.起付线

    If bln冲销 Then
       '需确定上次中心返回的数据

        gstrSQL = "" & _
            "   Select *  " & _
            "   From 保险结算记录 " & _
            "   Where 记录id=" & lng结帐ID
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "提取中心收费时返回的数据"
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "不存在上次收费的结算记录!"
            Exit Function
        End If
        '/???
        '原过程参数:
        '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
        "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
        '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
        '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
        '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
        '过程新值代表为:
        '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
        '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
        '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
        '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,超限自付金额_IN,dbl个人帐户支付
        '       支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
           dbl个人帐户余额 = Round(NVL(rsTemp!帐户累计增加, 0), 2)
           dbl统筹支付累计 = Round(NVL(rsTemp!帐户累计支出, 0), 2)
           dbl补助保险支付 = Round(NVL(rsTemp!累计进入统筹, 0), 2)
           dbl补助帐户支付 = Round(NVL(rsTemp!累计统筹报销, 0), 2)
           dbl起付标准 = Round(NVL(rsTemp!起付线, 0), 2)
           dbl保险范围外自付 = Round(NVL(rsTemp!封顶线, 0), 2)
           dbl基本统筹支付 = Round(NVL(rsTemp!全自付金额, 0), 2)
           dbl基本统筹自付 = Round(NVL(rsTemp!首先自付金额, 0), 2)
           dbl补充统筹支付 = Round(NVL(rsTemp!进入统筹金额, 0), 2)
           dbl补充统筹自付 = Round(NVL(rsTemp!统筹报销金额, 0), 2)
           dbl非补助保险支付 = Round(NVL(rsTemp!大病自付金额, 0), 2)
           dbl个人帐户支付 = Round(NVL(rsTemp!个人帐户支付, 0), 2)
           dbl结算前基本帐户余额 = Round(NVL(rsTemp!结算前基本帐户余额, 0), 2)
           dbl结算前补助账户余额 = Round(NVL(rsTemp!结算前补助账户余额, 0), 2)
           dbl结算前统筹累计 = Round(NVL(rsTemp!结算前统筹累计, 0), 2)
    End If
    '求出医生编码
    
    '找出疾病编码
    With g病人身份_大连
        If gintInsure = TYPE_大连开发区 Then    '开发区
            strInfor = Lpad(gstr医院编码_大连, 6)       '医院代码
        Else
            strInfor = Lpad(gstr医院编码_大连, 4)       '医院代码
        End If
        strInfor = strInfor & " "      '子门诊标识
        If gintInsure = TYPE_大连开发区 Then   '开发区
            strInfor = strInfor & Lpad(.个人编号, 10)         '个人编号
        Else
            strInfor = strInfor & Lpad(.个人编号, 8)      '个人编号
        End If
        strInfor = strInfor & Lpad(.IC卡号, 7)       'IC卡号
        .治疗序号 = .治疗序号 + 1
        strInfor = strInfor & Lpad(.治疗序号, 4)       '治疗序号
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '结算时间
        strInfor = strInfor & String(10, " ") '病志号
        
        '周海全调试 2003-12-17
        '由于不管是结算还是冲销都不允许为负数，所以此处只能取绝对值
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl诊察费), 2))), 10) '诊察费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl草药费), 2))), 10) '草药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl成药费), 2))), 10) '成药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl西药费), 2))), 10)  '西药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl检查费), 2))), 10)  '检查费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl治疗费), 2))), 10)   '治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl大检费), 2))), 10)    '大检费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl特殊治疗费), 2))), 10)   '特殊治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl大检自费), 2))), 10)   '大检自费
        If gintInsure = TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl特殊治疗自费), 2))), 10)    '特治自费    NUM 145 10      院端填写
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl保险内自费费用), 2))), 10)    '保险内自费费用
        
        If gintInsure = TYPE_大连开发区 Then       '开发区
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl其它费), 2))), 10)    '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl非保险费用), 2))), 10)    '非保险费用
        End If
      '周海全调试 2003-12-22
        '此处如果是冲销应该同时提取上次结算情况填写
        
'        strInfor = strInfor & String(10, " ")    '中心返回:结算后个人帐户余额;开发区:结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
'        strInfor = strInfor & String(10, " ")    '中心返回:结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
        strInfor = strInfor & Lpad(dbl个人帐户余额, 10)
        strInfor = strInfor & Lpad(dbl统筹支付累计, 10)
        
        '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
        Dim dbl结算前余额(1 To 3) As Double '1-结算前基本帐户余额,2-结算前补助账户余额,3-结算前统筹支付累计
        
        dbl结算前余额(1) = .基本个人帐户余额
        dbl结算前余额(2) = .补助个人帐户余额
        dbl结算前余额(3) = .统筹累计
        
        '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
        If bln冲销 Then
                strInfor = strInfor & Lpad(dbl结算前基本帐户余额, 10)   '结算前基本帐户余额
                strInfor = strInfor & Lpad(dbl结算前补助账户余额, 10)    '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInfor = strInfor & Lpad(dbl结算前统筹累计, 10)     '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
        Else
            If gintInsure <> TYPE_大连开发区 And Get就诊分类(0, .就诊分类) = "S" Then
                strInfor = strInfor & Lpad(.补助帐户当前值, 10)   '结算前基本帐户余额
                strInfor = strInfor & Lpad("0", 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInfor = strInfor & Lpad("0", 10)   '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
                dbl结算前余额(1) = .补助帐户当前值
                dbl结算前余额(2) = 0
                dbl结算前余额(3) = 0
            Else
                strInfor = strInfor & Lpad(.基本个人帐户余额, 10)  '结算前基本帐户余额
                strInfor = strInfor & Lpad(Trim(CStr(.补助个人帐户余额)), 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
                strInfor = strInfor & Lpad(Trim(CStr(.统筹累计)), 10)    '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            End If
        End If
        
        If bln冲销 Then
            'dbl个人帐户余额 = Round(NVL(rsTemp!帐户累计增加, 0), 2)
            'dbl统筹支付累计 = Round(NVL(rsTemp!帐户累计支出, 0), 2)
            'dbl起付标准 = Round(NVL(rsTemp!起付线, 0), 2)
            
            strInfor = strInfor & Lpad(dbl个人帐户支付, 10) ' = Round(NVL(rsTemp!个人帐户支付, 0), 2)
            strInfor = strInfor & Lpad(dbl补助帐户支付, 10) ' = Round(NVL(rsTemp!累计统筹报销, 0), 2)
            strInfor = strInfor & Lpad(dbl基本统筹支付, 10) ' = Round(NVL(rsTemp!全自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl基本统筹自付, 10) ' = Round(NVL(rsTemp!首先自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补充统筹支付, 10) ' = Round(NVL(rsTemp!进入统筹金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补充统筹自付, 10) ' = Round(NVL(rsTemp!统筹报销金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补助保险支付, 10) ' = Round(NVL(rsTemp!累计进入统筹, 0), 2)
            strInfor = strInfor & Lpad(dbl非补助保险支付, 10) ' = Round(NVL(rsTemp!大病自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl保险范围外自付, 10) ' = Round(NVL(rsTemp!封顶线, 0), 2)
        Else
        
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
            strInfor = strInfor & String(10, " ")    '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本补助保险支付 ；开发区:公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
            strInfor = strInfor & String(10, " ")    '中心返回:本次非基本补助保险支付；开发区:公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
            strInfor = strInfor & String(10, " ")    '中心返回:本次保险范围外自付；开发区:限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
        End If
        If gintInsure <> TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(Trim(CStr(dbl特殊治疗自费)), 10)    '本次特殊治疗自付
        End If
        
       
        '周海全调试 2003-12-22
        '门诊无需要起付标准，应该传0
        '        strInfor = strInfor & Lpad(Trim(CStr(dbl起付标准)), 10)    '起付标准
        strInfor = strInfor & Lpad(0, 10)
        strInfor = strInfor & Lpad(.转诊单号, 6)     '转诊单号
        strInfor = strInfor & Lpad(Get就诊分类(int业务, .就诊分类), 1)     '就诊分类
        
        If gintInsure <> TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(.参保类别3, 1)    '参保类别3:0 企保、1 事保，根据验卡结果
        End If
        strInfor = strInfor & Lpad(.职工就医类别, 1)       '职工就医类别
        
        strInfor = strInfor & Lpad(.诊断编码, 16)    '诊断编码
        strInfor = strInfor & Lpad(str医生, 6)    '医师代码
        strInfor = strInfor & Lpad(UserInfo.编号, 6)    '操作员代码
        strInfor = strInfor & Lpad(.诊断名称, 30)    '诊断名称
        'A-治愈、B-好转、C-未愈、D-死亡、E-其他
        strInfor = strInfor & "A"    '治愈情况标识
        strInfor = strInfor & String(8, " ")      '出院日期
        
        If gintInsure = TYPE_大连开发区 Then       '开发区
        Else
            strInfor = strInfor & String(16, " ")      '传输时间
        End If
        strInfor = strInfor & String(10, " ")      '错误代码
    End With
    
    '调用1002    12  423 实时结算
    Set门诊结算或冲销 = 业务请求_大连(IIf(gintInsure = TYPE_大连开发区, 2, 1), 1002, strInfor)
    If Set门诊结算或冲销 = False Then
        Exit Function
    End If
    
    '
   
    
   
    
    '开发区:
    '   结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
    '   结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
    
    '    本次基本个人帐户支付    NUM 225 10      中心返回
    '    本次补助个人帐户支付    NUM 235 10      中心返回
    '    本次基本统筹支付    NUM 245 10      中心返回
    '    本次基本统筹自付    NUM 255 10      中心返回
    '    本次补充统筹支付    NUM 265 10      中心返回
    '    本次补充统筹自付    NUM 275 10      中心返回
    '    本次基本补助保险支付    NUM 285 10  公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
    '    本次非基本补助保险支付  NUM 295 10  公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
    '    本次保险范围外自付  NUM 305 10  限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
    '大连市:
    '   结算后个人帐户余额  NUM 161 10  ①  如果是基本医疗结算表示：基本个人帐户＋补助个人帐户② 如果是慢病结算表示: 慢病帐户结算后余额 中心
    '   结算后统筹支付累计  NUM 171 10  基本统筹累计＋补充统筹累计  中心
    
    '    本次基本个人帐户支付    NUM 211 10  如果是慢病结算，表示慢病帐户支付    中心
    '    本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0 中心
    '    本次基本统筹支付    NUM 231 10      中心
    '    本次基本统筹自付    NUM 241 10      中心
    '    本次补充统筹支付    NUM 251 10  如果是生育结算，本字段用于存放生育保险支付  中心
    '    本次补充统筹自付    NUM 261 10      中心
    '    本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付 2． 如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    '    本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付   2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分    中心
    '    本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    Dim i As Long
    If gintInsure = TYPE_大连开发区 Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    
    dbl个人帐户余额 = Val(Substr(strInfor, i - 40, 10))
    dbl统筹支付累计 = Val(Substr(strInfor, i - 30, 10))  '结算后统筹支付累计=基本统筹累计＋补充统筹累计
    
    dbl个人帐户支付 = Val(Substr(strInfor, i + 10, 10)) '本次基本个人帐户支付=如果是慢病结算，表示慢病帐户支付
    dbl补助帐户支付 = Val(Substr(strInfor, i + 20, 10))    '本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0
    dbl基本统筹支付 = Val(Substr(strInfor, i + 30, 10))   '本次基本统筹支付    NUM 231 10      中心
    dbl基本统筹自付 = Val(Substr(strInfor, i + 40, 10))     '本次基本统筹自付    NUM 241 10      中心
    dbl补充统筹支付 = Val(Substr(strInfor, i + 50, 10))     '本次补充统筹支付    NUM 251 10      中心
    dbl补充统筹自付 = Val(Substr(strInfor, i + 60, 10))     '本次补充统筹自付    NUM 261 10      中心
    dbl补助保险支付 = Val(Substr(strInfor, i + 70, 10))     '本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付2．   如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    dbl非补助保险支付 = Val(Substr(strInfor, i + 80, 10))     '本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分
    dbl保险范围外自付 = Val(Substr(strInfor, i + 90, 10))     '本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    '/???
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    '    诊察费_IN,草药费_IN,成药费_IN,西药费_IN,检查费_IN,治疗费_IN,大检费_IN,大检自费_IN,特殊治疗费_IN,特殊治疗自费_IN,保险内自费费用_IN,非保险费用_IN,统筹比例_IN,其它费
    '   结算前基本帐户余额,结算前补助账户余额,结算前统筹累计
    '过程新值代表为:
    '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
    '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
    '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
    '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,结算前基本帐户余额(参见:说明) ,dbl个人帐户支付
    '       支付顺序号_IN(就诊分类;转诊单号;诊断编码),主页ID_IN,中途结帐_IN,诊断名称_IN
    '    诊察费_IN,草药费_IN,成药费_IN,西药费_IN,检查费_IN,治疗费_IN,大检费_IN,大检自费_IN,特殊治疗费_IN,特殊治疗自费_IN,保险内自费费用_IN,非保险费用_IN,统筹比例_IN,其它费,
    '   结算前基本帐户余额,结算前补助账户余额,结算前统筹累计
    '说明:
    '
    gstrSQL = "zl_保险结算记录_insert(1," & IIf(bln冲销, lng冲销ID, lng结帐ID) & "," & gintInsure & "," & lng病人id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
       dbl个人帐户余额 & "," & dbl统筹支付累计 & "," & dbl补助保险支付 & "," & dbl补助帐户支付 & "," & "Null" & "," & dbl起付标准 & "," & dbl保险范围外自付 & "," & dbl起付标准 & "," & _
       curTotal & "," & dbl基本统筹支付 & "," & dbl基本统筹自付 & "," & _
       dbl补充统筹支付 & "," & dbl补充统筹自付 & "," & dbl非补助保险支付 & ",Null," & dbl个人帐户支付 & ",'" & _
       Get就诊分类(int业务, g病人身份_大连.就诊分类) & ";" & g病人身份_大连.转诊单号 & ";" & g病人身份_大连.诊断编码 & "',null,null,'" & g病人身份_大连.诊断名称 & "'," & _
        dbl诊察费 & "," & dbl草药费 & "," & dbl成药费 & "," & dbl西药费 & "," & dbl检查费 & "," & dbl治疗费 & "," & dbl大检费 & "," & dbl大检自费 & "," & dbl特殊治疗费 & "," & dbl特殊治疗自费 & "," & dbl保险内自费费用 & "," & dbl非保险费用 & "," & dbl统筹比例 & "," & dbl其它费 & "," & _
         dbl结算前余额(1) & "," & dbl结算前余额(2) & "," & dbl结算前余额(3) & _
         " )"
             
    zlDatabase.ExecuteProcedure gstrSQL, "保存门诊收费数据"
    
    
    Set门诊结算或冲销 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 门诊结算冲销_大连(lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    Err = 0
    On Error GoTo ErrHand:
    门诊结算冲销_大连 = Set门诊结算或冲销(True, lng结帐ID, cur个人帐户, lng病人id, "")
    Exit Function
ErrHand:
    门诊结算冲销_大连 = False
End Function

Public Function 入院登记_大连(lng病人id As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo ErrHand
    
    '读取病人的相关保险信息

    gstrSQL = "select * From 保险帐户 where  险类=" & gintInsure & "  and 病人id=" & lng病人id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    str转诊单号 = NVL(rsTemp!人员身份)
    lng中心 = IIf(gintInsure = 83, 2, 1)
    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!顺序号, 1), 4)      '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(NVL(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人id & _
            "       and A.病人ID=" & lng病人id & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID"
            
    Call OpenRecordset(rsTemp, "读取入院信息")
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = NVL(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(NVL(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(NVL(rsTemp!入院经办时间), 16)     '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    If lng中心 = 2 Then
        '开发区为:住院 2、家床 4取消住院登记 C
        strInfor = strInfor & IIf(str就诊分类 = "4", "4", "2")
    Else
        strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    End If

    gstrSQL = "Select * From 床位状况记录 D where 病区ID=" & NVL(rsTemp!当前病区ID, 0) & " And 床号=" & NVL(rsTemp!当前床号, 0)
    Call OpenRecordset(rsTemp, "读取床位信息")
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(NVL(rsTemp!房间号)) & "室" & Trim(NVL(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = NVL(rsTemp!入院诊断)
        str确诊诊断名称 = NVL(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    If lng中心 = 2 Then
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
    Else
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
        strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
        strInfor = strInfor & Lpad(str确诊诊断名称, 30)  '确诊诊断名称    CHAR    114 30      N   院端
    End If
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    End If
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & "A"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
        strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    End If
    '1004    9   206 实时住院登记数据提交
    入院登记_大连 = 业务请求_大连(lng中心, 1004, strInfor)
    If 入院登记_大连 = False Then
        ShowMsgbox "实时住院登记数据提交失败!"
        Exit Function
    End If
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure("办理入院登记")
    入院登记_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 入院登记撤销_大连(lng病人id As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
                '取入院登记验证所返回的顺序号
                
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    gstrSQL = " Select Count(*) Records From 病人费用记录 " & _
              " Where 病人ID=" & lng病人id & " And 主页ID=" & lng主页ID
    Call OpenRecordset(rsTemp, "撤销入院检查")
    
    If rsTemp!Records <> 0 Then
        MsgBox "已经存在费用记录，不允许办理撤销入院登记！", vbInformation, gstrSysName
        Exit Function
    End If

    
    On Error GoTo ErrHand
    
    '读取病人的相关保险信息

    gstrSQL = "select * From 保险帐户 where  险类=" & gintInsure & "  and 病人id=" & lng病人id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "撤消入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    str转诊单号 = NVL(rsTemp!人员身份)
    lng中心 = IIf(gintInsure = 83, 2, 1)

    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!顺序号, 1), 4)      '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(NVL(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人id & _
            "       and A.病人ID=" & lng病人id & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID"
            
    Call OpenRecordset(rsTemp, "读取入院信息")
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = NVL(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(NVL(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(NVL(rsTemp!入院经办时间), 16)      '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    
    If lng中心 = 2 Then
        '开发区为:住院 2、家床 4取消住院登记 C
        strInfor = strInfor & "C"                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    Else
        strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    End If
    gstrSQL = "Select * From 床位状况记录 D where 病区ID=" & NVL(rsTemp!当前病区ID, 0) & " And 床号=" & NVL(rsTemp!当前床号, 0)
    Call OpenRecordset(rsTemp, "读取床位信息")
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(NVL(rsTemp!房间号)) & "室" & Trim(NVL(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人id & " and a.主页id=" & lng主页ID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = NVL(rsTemp!入院诊断)
        str确诊诊断名称 = NVL(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    If lng中心 = 2 Then
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
    Else
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
        strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
        strInfor = strInfor & Lpad(str确诊诊断名称, 30)  '确诊诊断名称    CHAR    114 30      N   院端
    End If
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    End If
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & "C"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
        strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    End If
    
    '1004    9   206 实时住院登记数据提交
    入院登记撤销_大连 = 业务请求_大连(lng中心, 1004, strInfor)
    If 入院登记撤销_大连 = False Then
        ShowMsgbox "实时住院登记撤消数据提交失败!"
        Exit Function
    End If
    gstrSQL = "zl_保险帐户_出院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure("办理撤销入院登记")
    入院登记撤销_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_大连(lng病人id As Long, lng主页ID As Long) As Boolean
    '
    '办理HIS出院
    gstrSQL = "zl_保险帐户_出院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure("出院登记")
    出院登记_大连 = True
End Function
Public Function 出院登记撤销_大连(lng病人id As Long, lng主页ID As Long) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    '需生新验卡
     lng中心 = IIf(gintInsure = 83, 2, 1)
    
     If 读取病人身份_大连(lng中心) = False Then Exit Function
    
    '存在未结费用的病人才允许撤销HIS出院；否则认为已办理医保出院，不允许再办理HIS出院
    If Not 存在未结费用(lng病人id, lng主页ID) Then
        MsgBox "医保已出院的病人不允许撤销出院！", vbInformation, gstrSysName
        Exit Function
    End If
               
    On Error GoTo ErrHand
    
    '读取病人的相关保险信息
    gstrSQL = "select * From 保险帐户 where  险类=" & gintInsure & "  and 病人id=" & lng病人id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "撤消入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    
    str转诊单号 = NVL(rsTemp!人员身份)
    

    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(g病人身份_大连.治疗序号, 4)       '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(NVL(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人id & _
            "       and A.病人ID=" & lng病人id & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID"
            
    Call OpenRecordset(rsTemp, "读取入院信息")
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = NVL(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(NVL(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(NVL(rsTemp!入院经办时间), 16)      '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    
    If lng中心 = 2 Then
        '开发区为:住院 2、家床 4取消住院登记 C
        strInfor = strInfor & "C"                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    Else
        strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    End If
    gstrSQL = "Select * From 床位状况记录 D where 病区ID=" & NVL(rsTemp!当前病区ID, 0) & " And 床号=" & NVL(rsTemp!当前床号, 0)
    Call OpenRecordset(rsTemp, "读取床位信息")
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(NVL(rsTemp!房间号)) & "室" & Trim(NVL(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1 and a.病人id=" & lng病人id & " and a.主页id=" & lng主页ID
         
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = NVL(rsTemp!入院诊断)
        str确诊诊断名称 = NVL(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    If lng中心 = 2 Then
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
    Else
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
        strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
        strInfor = strInfor & Lpad(str确诊诊断名称, 30)  '确诊诊断名称    CHAR    114 30      N   院端
    End If
    
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    End If
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & "A"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
        strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    End If
    
    '1004    9   206 实时住院登记数据提交
    出院登记撤销_大连 = 业务请求_大连(lng中心, 1004, strInfor)
    If 出院登记撤销_大连 = False Then
        ShowMsgbox "实时住院登记撤消数据提交失败!"
        Exit Function
    End If
    
    gstrSQL = "zl_保险帐户_入院(" & lng病人id & "," & gintInsure & ")"
    Call ExecuteProcedure("办理撤销出院登记")
    出院登记撤销_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub 获取病人信息_大连(ByVal lng病人id As Long)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的相关信息,将其值赋给G病人身份
    '--入参数:lng病人id
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    '读取医保病人相关信息，并更新公用结构体
        
    gstrSQL = "" & _
        "   Select *" & _
        "   From 保险帐户" & _
        "   Where 险类=" & gintInsure & " And 病人ID=" & lng病人id
    Call OpenRecordset(rsTemp, "读取医保病人的相关信息")
    
    If Not rsTemp.EOF Then
        With g病人身份_大连
            .IC卡号 = NVL(rsTemp!卡号, 0)
            .个人编号 = NVL(rsTemp!医保号)
            .医保中心 = IIf(gintInsure = 83, 2, 1) ' NVL(rsTemp!中心, 1)
            .治疗序号 = NVL(rsTemp!顺序号, 0)
            .转诊单号 = NVL(rsTemp!人员身份)
            .基本个人帐户余额 = NVL(rsTemp!帐户余额, 0)
            .补助个人帐户余额 = Val(NVL(rsTemp!退休证号))
            
            .职工就医类别 = Decode(NVL(rsTemp!在职, 1), 1, "A", 2, "B", 3, "L", 4, "T", 5, "Q", "")
            .就诊分类 = NVL(rsTemp!灰度级, 0)
            .参保类别3 = NVL(rsTemp!单位编码, 0)
            '.起付线 = NVL(rsTemp!统筹报销累计, 0)
        End With
    End If
End Sub
Private Function Get治渝情况_大连(lng病人id As Long, lng主页ID As Long) As String
    '功能:获取治渝情况标识
    '     A-治愈、B-好转、C-未愈、D-死亡、E-其他
    
    Dim rsInNote As New ADODB.Recordset
    Dim strTmp As String
    
    strTmp = " Select A.出院情况" & _
             " From 诊断情况 A,疾病编码目录 B " & _
             " Where A.病人ID=" & lng病人id & " And A.疾病ID=B.ID(+) And A.主页ID=" & lng主页ID & _
             "       And A.诊断类型 in (2,3)" & _
             " Order by A.诊断类型 Desc"
    
    rsInNote.CursorLocation = adUseClient
    Call OpenRecordset(rsInNote, "医保接口", strTmp)
    strTmp = ""
    If Not rsInNote.EOF Then
        strTmp = NVL(rsInNote!出院情况)
    End If
    strTmp = Decode(strTmp, "治愈", "A", "好转", "B", "未愈", "C", "死亡", "D", "其他", "E")
    
End Function
Public Function 住院虚拟结算_大连(rsExse As Recordset, ByVal lng病人id As Long) As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '      字段:记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID, _
    '           收费类别,收费细目ID,收费名称,开单部门,规格,产地,数量,价格,金额,医生,登记时间, _
    '           是否上传,是否急诊,保险项目否,摘要
    
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    '接口返回的报销额减去本次住院期间以往报销额的汇总金额后，才是本次的实际报销额
    'rsExse记录集中的字段清单
    '记录性质,记录状态,NO,序号,病人ID,主页ID,婴儿费,医保项目编码,保险大类ID,
    '收费类别,收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门
    '规格,产地,数量,价格,金额,医生,登记时间,是否上传,是否急诊,保险项目否,摘要
    Dim rsTemp As New ADODB.Recordset
    Dim rs费用 As New ADODB.Recordset
'    Dim rs大类 As New ADODB.Recordset
    Dim curTotal As Currency
    
    Dim lng主页ID As Long
    Dim cur个人自付 As Currency, cur个人帐户 As Currency
    Dim str入院年份 As String, str结算年份 As String
    Dim str结算时间 As String, str经办时间 As String
    Dim strInfor As String  '定义中心返回串
    Dim dbl诊察费 As Double
    Dim dbl草药费 As Double
    Dim dbl成药费 As Double
    Dim dbl西药费 As Double
    Dim dbl检查费 As Double
    Dim dbl治疗费 As Double
    Dim dbl大检费 As Double
    Dim dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double
    Dim dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double
    Dim dbl比例 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl起付标准 As Double
    
    Dim str诊断编码 As String  '疾病编码
    Dim str医师代码 As String
    Dim str操作员代码 As String
    Dim str诊断名称 As String
    Dim str治愈情况标识 As String
    Dim strTmp As String
    Dim str医生 As String
    Dim str明细 As String       '明细串
    Dim str国家编码 As String
    Dim str项目统计分类 As String
    Dim str出院日期 As String
    Dim dbl项目名称 As Double
    Dim str住院号 As String
    
    Dim intMouse As Integer
    On Error GoTo ErrHand
    intMouse = Screen.MousePointer
    
    '在虚拟结算前需验证身分
    Screen.MousePointer = 1
    If 身份标识_大连(4, lng病人id) = "" Then
        Screen.MousePointer = intMouse
        住院虚拟结算_大连 = ""
        Exit Function
    End If
    Screen.MousePointer = intMouse
    
'    '获取病人信息
'    Call 获取病人信息_大连(lng病人id)
    
    cur个人帐户 = g病人身份_大连.基本个人帐户余额

    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份,A.出院日期,B.住院号" & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=" & lng病人id & " And A.主页ID=B.住院次数 And A.病人ID=B.病人ID"

    Call OpenRecordset(rsTemp, "获取病人入院时间")
    str入院年份 = rsTemp!入院年份
    lng主页ID = rsTemp!主页ID
    str出院日期 = Format(rsTemp!出院日期, "yyyymmdd")
    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str结算时间 = str经办时间
    str结算年份 = Mid(str经办时间, 1, 4)
    str住院号 = NVL(rsTemp!住院号)
    
    '重新获取记录
    Set rs费用 = Get住院虚拟记录(lng病人id)
    If rs费用.RecordCount <= 0 Then
        ShowMsgbox "有项目未设置医保项目，不能结算!"
        Exit Function
    End If
    dbl起付标准 = g病人身份_大连.起付线

    With rs费用
        '上传费用明细
        curTotal = 0
        Do While Not .EOF
        
            If !金额 < 0 Or !价格 < 0 Or !记录状态 <> 1 Then
                MsgBox "该病人的待结费用中有项目的金额或价格为负数,或者状态不正确,请检查后重新结算!", vbOKOnly
                Exit Function
            End If
        
            If str医生 = "" Then
                str医生 = NVL(!医生编号)
                If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                    str医生 = Substr(str医生, 1, 6)
                End If
            End If
            curTotal = curTotal + NVL(!金额, 0)
            
            lng病人id = NVL(!病人ID, 0)
            strTmp = NVL(!参数值)
            
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                strTmp = Split(strTmp, ";")(0)
                
                '计算保费
                dbl比例 = NVL(!住院比额, 0) / 100
                
                '中心为:A在职、B退休、L离休、T特诊,我们默认为1在职、2退休、3离休、4特诊
                If g病人身份_大连.医保中心 <> 2 And g病人身份_大连.职工就医类别 = "L" And g病人身份_大连.参保类别3 = "0" And NVL(!保险项目否, 0) = 1 Then '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '  大连市  企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                End If
                If NVL(!医保项目编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If NVL(!医保项目编码) = "大检" Then
                    strTmp = "大检费"
                End If
                
                If g病人身份_大连.医保中心 <> 2 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(NVL(rsTemp!医保项目名称))
                End If
                
                If g病人身份_大连.医保中心 <> 2 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                
                '如果是床位,则需按如下方式处理,主包床按统筹比例计算,被包床为100的自费,不分开发区和大连市
                If NVL(!收费类别) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select 附加床位 From 病人变动记录 " & _
                        "   Where 床号=" & NVL(!床号, 0) & _
                        "         And ( (to_date('" & Format(!登记时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss') between 开始时间 and 终止时间) or" & _
                        "               ( 终止时间 is null  and 开始时间<=to_date('" & Format(!登记时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))) " & _
                        "         And 床号 is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定是否为包床!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!附加床位 = 1 Then
                            '表示被包床位,为全自费
                            dbl比例 = 0
                       End If
                    End If
                End If
                If dbl比例 <> 0 Then
                    Select Case strTmp
                        Case "诊察费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl诊察费 = dbl诊察费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl诊察费 = dbl诊察费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl诊察费 = dbl诊察费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "草药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl草药费 = dbl草药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl草药费 = dbl草药费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl草药费 = dbl草药费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "成药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl成药费 = dbl成药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl成药费 = dbl成药费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl成药费 = dbl成药费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "西药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl西药费 = dbl西药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl西药费 = dbl西药费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl西药费 = dbl西药费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "检查费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl检查费 = dbl检查费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按金额
                                        dbl检查费 = dbl检查费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl检查费 = dbl检查费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "治疗费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl治疗费 = dbl治疗费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按金额
                                        dbl治疗费 = dbl治疗费 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    dbl治疗费 = dbl治疗费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                End If
                        Case "大检费"
                            If NVL(!算法, 0) = 2 Then
                                        '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl大检费 = dbl大检费 + Round(NVL(!特准定额, 0), 2)
                                        If g病人身份_大连.医保中心 = 1 And g病人身份_大连.职工就医类别 = "Q" Then
                                            '自费部分放入 保险内自费费用中
                                        Else
                                             dbl大检自费 = dbl大检自费 + NVL(!金额, 0) - NVL(!特准定额, 0)
                                        End If
                                    Else
                                        '则按金额
                                        dbl大检费 = dbl大检费 + Round(NVL(!金额, 0), 2)
                                    End If
                            Else
                                    
                                If g病人身份_大连.医保中心 = 1 Then
                                    '---周顺利
                                    '大连市和开发区对大检费用处理不同,
                                    '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                    dbl大检费 = dbl大检费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                    
                                    If g病人身份_大连.职工就医类别 = "Q" Then
                                        '自费部分放入保险内自费费用中
                                    Else
                                        dbl大检自费 = dbl大检自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    End If
                                
                                Else
                                    
                                    dbl大检费 = dbl大检费 + Round(NVL(!金额, 0), 2)
                                    
                                    dbl大检自费 = dbl大检自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    
                                End If
                                    
                                    
                                    
    '                                If g病人身份_大连.医保中心 = 1 Then
    '                                    dbl大检费 = dbl大检费 + Round(NVL(!金额, 0), 2)
    '                                Else
    '                                    dbl大检费 = dbl大检费 + Round(NVL(!金额, 0) * dbl比例, 2)
    '                                End If
    '
    '                                If g病人身份_大连.医保中心 = 1 And g病人身份_大连.职工就医类别 = "Q" Then
    '                                    '自费部分放入 保险内自费费用中
    '                                Else
    '                                    dbl大检自费 = dbl大检自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
    '                                End If
                            End If
                        Case "特殊治疗费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        '则按定额计算
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按金额
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!金额, 0), 2)
                                        If g病人身份_大连.医保中心 = 1 And g病人身份_大连.职工就医类别 = "Q" Then
                                            '自费部分放入 保险内自费费用中
                                        Else
                                            dbl特殊治疗自费 = dbl特殊治疗自费 + NVL(!金额, 0) - NVL(!特准定额, 0)
                                        End If
                                    End If
                                Else
                                    '大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                    If g病人身份_大连.医保中心 = 1 Then
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!金额, 0), 2)
                                    Else
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                    End If
                                    If g病人身份_大连.医保中心 = 1 And g病人身份_大连.职工就医类别 = "Q" Then
                                        '自费部分放入 保险内自费费用中
                                    Else
                                        dbl特殊治疗自费 = dbl特殊治疗自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    End If
                                End If
                    End Select
                End If
                If g病人身份_大连.医保中心 <> 2 And g病人身份_大连.职工就医类别 = "Q" Then
                        '自费部分放入 保险内自费费用中
                         If NVL(!算法, 0) = 2 Then
                                '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                If NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) - NVL(!特准定额, 0), 2)
                                End If
                          Else
                                If dbl比例 <> 0 Then
                                    If !保险项目否 = 1 Then
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    End If
                                Else
                                    '100自费部分放入非保险费用中
                                    dbl非保险费用 = dbl非保险费用 + Round(NVL(!金额, 0), 2)
                                End If
                          End If
                Else
                         If gintInsure = TYPE_大连开发区 Then
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) And !保险项目否 = 1 Then
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) - NVL(!特准定额, 0), 2)
                                    End If
                                    If NVL(!特准定额, 0) < NVL(!金额, 0) And !保险项目否 <> 1 Then
                                        dbl其它费 = dbl其它费 + Round(NVL(!金额, 0) - NVL(!特准定额, 0), 2)
                                    End If
                                Else
                         
                                    If !保险项目否 = 1 And dbl比例 <> 0 Then
                                        '险内药品自费  NUM 155 10  医保用药自费部分    院端填写
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    Else
                                        '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
                                        dbl其它费 = dbl其它费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    End If
                                End If
                         Else
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                     If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !保险项目否 = 1 And NVL(!特准定额, 0) < NVL(!金额, 0) Then
                                        ''医保用药以及除了大检、特治外检查治疗项目的自费部分
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) - NVL(!特准定额, 0), 2)
                                     End If
                                    If !保险项目否 <> 1 Or dbl比例 = 0 Then
                                        '非医保用药以及诊疗项目
                                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!金额, 0), 2)
                                    End If
                                Else
                                    If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !保险项目否 = 1 And dbl比例 <> 0 Then
                                        '医保用药以及除了大检、特治外检查治疗项目的自费部分
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                    End If
                                    If !保险项目否 <> 1 Or dbl比例 = 0 Then
                                        '非医保用药以及诊疗项目
                                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!金额, 0), 2)
                                    End If
                                End If
                         End If
                 End If
            End If
            .MoveNext
        Loop
        
        With g病人身份_大连
            If .医保中心 = 2 Then   '开发区
                strInfor = Lpad(gstr医院编码_大连, 6)       '医院代码
            Else
                strInfor = Lpad(gstr医院编码_大连, 4)       '医院代码
            End If
            strInfor = strInfor & " "      '子门诊标识
            If .医保中心 = 2 Then   '开发区
                strInfor = strInfor & Lpad(.个人编号, 10)       '个人编号
            Else
                strInfor = strInfor & Lpad(.个人编号, 8)      '个人编号
            End If
            strInfor = strInfor & Lpad(.IC卡号, 7)       'IC卡号
            strInfor = strInfor & Lpad(.治疗序号 + 1, 4)      '治疗序号, 验卡返回结果值＋1
            strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '结算时间
            strInfor = strInfor & Lpad(str住院号, 10) '病志号:住院就为住院号
            
            
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl诊察费, 2))), 10) '诊察费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl草药费, 2))), 10) '草药费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl成药费, 2))), 10) '成药费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl西药费, 2))), 10) '西药费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl检查费, 2))), 10) '检查费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl治疗费, 2))), 10)  '治疗费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl大检费, 2))), 10)  '大检费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl特殊治疗费, 2))), 10)  '特殊治疗费
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl大检自费, 2))), 10)  '大检自费
            If .医保中心 = 2 Then
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl特殊治疗自费, 2))), 10)   '特治自费    NUM 145 10      院端填写
            End If
            strInfor = strInfor & Lpad(Trim(CStr(Round(dbl保险内自费费用, 2))), 10)   '保险内自费费用
            
            If .医保中心 = 2 Then        '开发区
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl其它费, 2))), 10)   '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
            Else
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl非保险费用, 2))), 10)    '非保险费用
            End If
            
            strInfor = strInfor & String(10, " ")    '中心返回:结算后个人帐户余额;开发区:结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
            strInfor = strInfor & String(10, " ")    '中心返回:结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
            
            '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
            strInfor = strInfor & Lpad(.基本个人帐户余额, 10)  '结算前基本帐户余额
            strInfor = strInfor & Lpad(Trim(CStr(.补助个人帐户余额)), 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
            strInfor = strInfor & Lpad(Trim(CStr(.统筹累计)), 10)    '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
            strInfor = strInfor & String(10, " ")    '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本补助保险支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次非基本补助保险支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次保险范围外自付
              
              
            If .医保中心 <> 2 Then
                strInfor = strInfor & Lpad(Trim(CStr(Round(dbl特殊治疗自费, 2))), 10)   '本次特殊治疗自付
            End If
            
            strInfor = strInfor & Lpad(Trim(CStr(dbl起付标准)), 10)    '起付标准
              
            strInfor = strInfor & Lpad(.转诊单号, 6)     '转诊单号
            strInfor = strInfor & Lpad(Get就诊分类(0, .就诊分类), 1)     '就诊分类
            If .医保中心 <> 2 Then
                strInfor = strInfor & Lpad(.参保类别3, 1)    '参保类别3:0 企保、1 事保，根据验卡结果
            End If
            strInfor = strInfor & Lpad(.职工就医类别, 1)       '职工就医类别
              
            strInfor = strInfor & Lpad(.诊断编码, 16)    '诊断编码
            strInfor = strInfor & Lpad(str医生, 6)    '医师代码
            strInfor = strInfor & Lpad(UserInfo.编号, 6)    '操作员代码
            strInfor = strInfor & Lpad(.诊断名称, 30)    '诊断名称
            
            'A-治愈、B-好转、C-未愈、D-死亡、E-其他
            strInfor = strInfor & Lpad(Get治渝情况_大连(lng病人id, lng主页ID), 1)   '治愈情况标识
            strInfor = strInfor & Lpad(IIf(str出院日期 = "", Format(zlDatabase.Currentdate, "yyyyMMDD"), str出院日期), 8) '出院日期
            
            If .医保中心 = 2 Then       '开发区
            Else
                strInfor = strInfor & String(16, " ")      '传输时间
            End If
            strInfor = strInfor & String(10, " ")      '错误代码
          End With
    
        '调用虚拟接口(1006    12  423   实时结算预算
        If 业务请求_大连(g病人身份_大连.医保中心, 1006, strInfor) = False Then
            ShowMsgbox "住院虚拟结算失败!"
            Exit Function
        End If
        
        
        g病人身份_大连.支付金额 = curTotal
    End With
    
    Dim str结算方式  As String
    
  
    '开发区:
    '    本次基本个人帐户支付    NUM 225 10      中心返回
    '    本次补助个人帐户支付    NUM 235 10      中心返回
    '    本次基本统筹支付    NUM 245 10      中心返回
    '    本次基本统筹自付    NUM 255 10      中心返回
    '    本次补充统筹支付    NUM 265 10      中心返回
    '    本次补充统筹自付    NUM 275 10      中心返回
    '    本次基本补助保险支付    NUM 285 10  公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
    '    本次非基本补助保险支付  NUM 295 10  公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
    '    本次保险范围外自付  NUM 305 10  限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
    '大连市:
    '    本次基本个人帐户支付    NUM 211 10  如果是慢病结算，表示慢病帐户支付    中心
    '    本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0 中心
    '    本次基本统筹支付    NUM 231 10      中心
    '    本次基本统筹自付    NUM 241 10      中心
    '    本次补充统筹支付    NUM 251 10  如果是生育结算，本字段用于存放生育保险支付  中心
    '    本次补充统筹自付    NUM 261 10      中心
    '    本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付 2． 如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    '    本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付   2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分    中心
    '    本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    Dim i As Long
    
    If g病人身份_大连.医保中心 = 2 Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    '确定本次结算方式
    str结算方式 = "个人帐户;" & Format(Val(Substr(strInfor, i + 10, 10)), "###0.00;-###0.00;0;0") & ";0" '本次基本个人帐户支付,不充许修改
    str结算方式 = str结算方式 & "|" & "补助帐户;" & Format(Val(Substr(strInfor, i + 20, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "基本统筹;" & Format(Val(Substr(strInfor, i + 30, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "补充统筹;" & Format(Val(Substr(strInfor, i + 50, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "补助保险;" & Format(Val(Substr(strInfor, i + 70, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    str结算方式 = str结算方式 & "|" & "非补助保险;" & Format(Val(Substr(strInfor, i + 80, 10)), "###0.00;-###0.00;0;0") & ";0" '不充许修改
    
    住院虚拟结算_大连 = str结算方式
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function Get慢病帐户余额_大连() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取慢病帐户余额
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    '医院代码    CHAR    1   4       院端
    '个人编号    CHAR    5   8       院端
    '补助病种    CHAR    13  16  目前为: WZMB    院端
    '治疗序号    NUM 29  4       中心
    '补助帐户原始值  NUM 33  10  每次补助的累计值    中心
    '补助帐户当前值  NUM 43  10      中心
    '帐户状态    CHAR    53  1   A正常、C止付    中心

    
    Dim strTmp As String
    Err = 0
    On Error GoTo ErrHand:
    With g病人身份_大连
        strTmp = Lpad(gstr医院编码_大连, 4)      '医院代码    CHAR    1   4       院端
        strTmp = strTmp & Lpad(.个人编号, 8) '个人编号    CHAR    5   8       院端
        strTmp = strTmp & Lpad("WZMB", 16)  '补助病种    CHAR    13  16  目前为: WZMB    院端
        strTmp = strTmp & Space(4)  '治疗序号    NUM 29  4       中心
        strTmp = strTmp & Space(10)  '补助帐户原始值  NUM 33  10  每次补助的累计值    中心
        strTmp = strTmp & Space(10)  '补助帐户当前值  NUM 43  10      中心
        strTmp = strTmp & Space(1)   '帐户状态    CHAR    53  1   A正常、C止付    中心
        '向医保中心查询慢病
        '   1007    2   55  慢病帐户查询
        Get慢病帐户余额_大连 = 业务请求_大连(.医保中心, 1007, strTmp)
        If Get慢病帐户余额_大连 = False Then
            .补助帐户原始值 = 0
            .补助帐户当前值 = 0
            Exit Function
        End If
        .补助帐户原始值 = Val(Substr(strTmp, 33, 10))
        .补助帐户当前值 = Val(Substr(strTmp, 43, 10))
    End With
    Exit Function
ErrHand:
End Function
Private Function 住院结算及冲帐_大连(ByVal bln冲销 As Boolean, ByVal lng病人id As Long, ByVal lng结帐ID As Long, ByVal 原结帐id As Long, ByVal lng主页ID As Long) As Boolean

    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim str住院号 As String
    Dim str项目统计分类  As String
    Dim strInfor As String  '定义中心返回串
    Dim curTotal As Double
    Dim dbl诊察费 As Double
    Dim dbl草药费 As Double
    Dim dbl成药费 As Double
    Dim dbl西药费 As Double
    Dim dbl检查费 As Double
    Dim dbl治疗费 As Double
    Dim dbl大检费 As Double
    Dim dbl大检自费 As Double
    Dim dbl特殊治疗费 As Double
    Dim dbl特殊治疗自费 As Double
    Dim dbl保险内自费费用 As Double
    Dim dbl非保险费用 As Double
    Dim dbl比例 As Double
    Dim dbl其它费 As Double     '针对大连开发区的
    Dim dbl起付标准 As Double
   
    Dim dbl个人帐户余额 As Double
    Dim dbl统筹支付累计 As Double
    Dim dbl个人帐户支付 As Double
    Dim dbl补助帐户支付 As Double
    Dim dbl基本统筹支付 As Double
    Dim dbl基本统筹自付 As Double
    Dim dbl补充统筹支付 As Double
    Dim dbl补充统筹自付 As Double
    Dim dbl补助保险支付 As Double
    Dim dbl非补助保险支付 As Double
    Dim dbl保险范围外自付 As Double
    
    Dim dbl结算前基本帐户余额  As Double
    Dim dbl结算前补助账户余额  As Double
    Dim dbl结算前统筹累计  As Double
    
    Dim str医生 As String
    Dim str明细 As String       '明细串
    Dim str国家编码 As String
    Dim int业务 As Integer
    Dim str出院日期 As String
    
    int业务 = IIf(bln冲销, 1, 0)
    
    Err = 0
    On Error GoTo ErrHand:
    
    '住院应用保险支付大类中的住院比额
    gstrSQL = " " & _
        "        select a.id,a.记录性质,a.主页id,a.记录状态,a.登记时间,a.no,a.病人病区id,a.床号,a.序号,a.标识号 as 住院号,a.病人科室id,a.病人id,a.收费类别,b.类别,a.计算单位, " & _
        "               A.计算单位,A.数次*Nvl(A.付数,1) 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额 ,a.开单人 as 医生,c.编号 as 医生编号, " & _
        "               a.医嘱序号, A.实收金额,A.是否上传, " & _
        "               F.参数值,D.编码 as 项目编码,D.名称 as 项目名称,D.标识子码||D.标识主码 as 国家编码, " & _
        "               E.项目编码 as 医保编码,E.项目名称 as 医保名称,e.是否医保,e.大类id,G.住院比额 as 统筹比额,G.特准定额,G.算法,H.名称 as 开单部门,J.名称 as 剂型, " & _
        "               L.险类,l.中心 , l.卡号, l.医保号, l.人员身份, l.单位编码, l.顺序号, l.退休证号, l.帐户余额, l.当前状态, l.病种ID, l.在职, l.年龄段, l.灰度级, l.就诊时间 " & _
        "        from 病人费用记录 a,收费类别 b,人员表 c,收费细目 D,保险支付项目 E,保险支付大类 G,保险帐户 L,部门表 H, " & _
        "             (Select U.*,K.参数值 From 收费类别 U,保险参数 K where U.类别=K.参数名 and K.险类=" & gintInsure & "  ) F ," & _
        "             (Select distinct Q.药品id,T.名称 From 药品目录 Q,药品信息 R,药品剂型 T  Where  Q.药名id=R.药名id and R.剂型=T.编码 ) J " & _
        "        where a.收费类别=b.编码 and a.收费细目id=J.药品id(+)   and  Nvl(a.附加标志,0)<>9 and a.收费细目id=D.id and a.开单人=c.姓名(+) and a.收费类别=F.编码(+) and " & _
        "              a.收费细目id=E.收费细目ID  and E.大类id=G.id and a.病人id=L.病人ID and a.开单部门id=h.id  and " & _
        "              a.病人ID = " & lng病人id & " And a.结帐ID = " & lng结帐ID & " And E.险类 = " & gintInsure
        
    zlDatabase.OpenRecordset rs明细, gstrSQL, "提取住院结帐明细"
    
    '确定该病人是否已经出院
    gstrSQL = "Select * From 病案主页 where 病人id=" & lng病人id & " and 主页id=" & lng主页ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人是否出院"
    str出院日期 = ""
    If rsTemp.EOF Then
        strTmp = Get入院诊断(lng病人id, lng主页ID, , True)
    Else
        If IsNull(rsTemp!出院日期) Then
            strTmp = Get入院诊断(lng病人id, lng主页ID, , True)
        Else
            strTmp = 获取入出院诊断(lng病人id, lng主页ID, False, , True)
            str出院日期 = Format(rsTemp!出院日期, "yyyymmdd")
        End If
    End If
    If InStr(1, strTmp, "|") <> 0 Then
        g病人身份_大连.诊断编码 = Split(strTmp, "|")(1)
        g病人身份_大连.诊断名称 = Split(strTmp, "|")(0)
    End If
    
    With rs明细
        If Not .EOF Then
            str住院号 = NVL(!住院号)
        End If
        Do While Not .EOF
            strTmp = NVL(!参数值)
            lng病人id = NVL(!病人ID, 0)
            If str医生 = "" Then
                str医生 = NVL(!医生编号)
                If LenB(StrConv(str医生, vbFromUnicode)) > 6 Then
                    str医生 = Substr(str医生, 1, 6)
                End If
            End If
            '确定相关数据
            If strTmp <> "" And InStr(1, strTmp, ";") <> 0 Then
                If Split(strTmp, ";")(1) = "" Then
                    str项目统计分类 = ""
                Else
                    str项目统计分类 = Mid(Split(strTmp, ";")(1), 1, 1)
                End If
                
                strTmp = Split(strTmp, ";")(0)
                '比例
                dbl比例 = NVL(!统筹比额, 0) / 100
                If NVL(!险类, 0) <> TYPE_大连开发区 And Val(NVL(!单位编码, "99")) = 0 And NVL(!在职, 0) = 3 And NVL(!是否医保, 0) = 1 Then '是企保和离休人员且是医保项目
                    '单位编码存储的是参保类别3   CHAR    90  1   0 企保、1 事保
                    '    企业单位离休医保：不完全执行医保政策，（如普通医保20%、10%自费部分不计入医保，现金支付，但此类病人这种自费部分计入医保，打印医保收据，只有100%自费需自付现金，开现金发票（可手写实现）注明: 此种病人属于拨款到医院单位
                    dbl比例 = 1
                End If
                If NVL(!医保编码) = "特治" Then
                    strTmp = "特殊治疗费"
                End If
                If NVL(!医保编码) = "大检" Then
                    strTmp = "大检费"
                End If
                If NVL(!险类, 0) = TYPE_大连市 And (g病人身份_大连.职工就医类别 = "L" Or _
                     g病人身份_大连.职工就医类别 = "T") Then
                    '如果是L离休和T特诊的就按事业比例计算
                    dbl比例 = Val(NVL(!医保名称))
                End If
                
                If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                    '如果是Q企业公费,如果比例为100自费,则需放入非保险费用中
                    If dbl比例 = 0 Then
                        '自费100
                        strTmp = ""
                    Else
                        '自费部分放入 保险内自费费用中
                    End If
                End If
                '如果是床位,则需按如下方式处理,主包床按统筹比例计算,被包床为100的自费,不分开发区和大连市
                If NVL(!收费类别) = "J" Then
                    
                    gstrSQL = "" & _
                        "   Select 附加床位 From 病人变动记录 " & _
                        "   Where 床号=" & NVL(!床号, 0) & _
                        "         And ( (to_date('" & Format(!登记时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss') between 开始时间 and 终止时间) or" & _
                        "               ( 终止时间 is null  and 开始时间<=to_date('" & Format(!登记时间, "YYYY-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))) " & _
                        "         And 床号 is not null"
                    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定是否为包床!"
                    If rsTemp.RecordCount >= 1 Then
                       If rsTemp!附加床位 = 1 Then
                            '表示被包床位,为全自费
                            dbl比例 = 0
                       End If
                    End If
                End If
                If dbl比例 <> 0 Then
                    '---周顺利
                    '---因为对报销比例为0的项目本身,是对任何人都是作为保险外自费
                    '---因为包床的特殊性,也是作为全自费费用处理,因此发现dbl比例=0则直接跳往保险内外自费费用处理
                    '---原来的处理上面对dbl比例=0的包床费用在治疗费和保险外费用都进行了处理，费用发生了重复
                               
                    Select Case strTmp
                        Case "诊察费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl诊察费 = dbl诊察费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl诊察费 = dbl诊察费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl诊察费 = dbl诊察费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "草药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl草药费 = dbl草药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl草药费 = dbl草药费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl草药费 = dbl草药费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "成药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl成药费 = dbl成药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl成药费 = dbl成药费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl成药费 = dbl成药费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "西药费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl西药费 = dbl西药费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按定额计算
                                        dbl西药费 = dbl西药费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl西药费 = dbl西药费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "检查费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl检查费 = dbl检查费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按结帐金额
                                        dbl检查费 = dbl检查费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl检查费 = dbl检查费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "治疗费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl治疗费 = dbl治疗费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按结帐金额
                                        dbl治疗费 = dbl治疗费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    dbl治疗费 = dbl治疗费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                End If
                        Case "大检费"
                            If NVL(!算法, 0) = 2 Then
                                        '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl大检费 = dbl大检费 + Round(NVL(!特准定额, 0), 2)
                                        If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                                            '自费部分放入 保险内自费费用中
                                        Else
                                             dbl大检自费 = dbl大检自费 + NVL(!结帐金额, 0) - NVL(!特准定额, 0)
                                        End If
                                    Else
                                        '则按结帐金额
                                        dbl大检费 = dbl大检费 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                            Else
                                    If NVL(!险类, 0) = TYPE_大连市 Then
                                        '---周顺利
                                        '大连市和开发区对大检费用处理不同,
                                        '大连市为扣除大检项目金额扣除大检自费的金额,其中的离休病人的大检自费全部记入险内自费
                                        
                                          dbl大检费 = dbl大检费 + Round(NVL(!金额, 0) * dbl比例, 2)
                                          
                                          If g病人身份_大连.职工就医类别 = "Q" Then
                                              '自费部分放入保险内自费费用中
                                          Else
                                              dbl大检自费 = dbl大检自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                          End If
                                      
                                      Else
                                          
                                          dbl大检费 = dbl大检费 + Round(NVL(!金额, 0), 2)
                                          
                                          dbl大检自费 = dbl大检自费 + Round(NVL(!金额, 0) * (1 - dbl比例), 2)
                                          
                                      End If
                                    
                                    'If NVL(!险类, 0) = TYPE_大连市 Then
                                    '    dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0), 2)
                                    'Else
                                    '    dbl大检费 = dbl大检费 + Round(NVL(!实收金额, 0) * dbl比例, 2)
                                    'End If
                                    '
                                    'If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                                    '    '自费部分放入 保险内自费费用中
                                    'Else
                                    '    dbl大检自费 = dbl大检自费 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    'End If
                            End If
                        Case "特殊治疗费"
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        '则按定额计算
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!特准定额, 0), 2)
                                    Else
                                        '则按结帐金额
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!结帐金额, 0), 2)
                                        If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                                            '自费部分放入 保险内自费费用中
                                        Else
                                            dbl特殊治疗自费 = dbl特殊治疗自费 + NVL(!结帐金额, 0) - NVL(!特准定额, 0)
                                        End If
                                    End If
                                Else
                                    '大连市与开发区计算方式不一致，大连市是总额，开发区仅是统筹部分
                                    If NVL(!险类, 0) = TYPE_大连市 Then
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!结帐金额, 0), 2)
                                    Else
                                        dbl特殊治疗费 = dbl特殊治疗费 + Round(NVL(!结帐金额, 0) * dbl比例, 2)
                                    End If
                                    If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                                        '自费部分放入 保险内自费费用中
                                    Else
                                        dbl特殊治疗自费 = dbl特殊治疗自费 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    End If
                                End If
                    End Select
                End If
                If NVL(!险类, 0) = TYPE_大连市 And g病人身份_大连.职工就医类别 = "Q" Then
                        '自费部分放入 保险内自费费用中
                         If NVL(!算法, 0) = 2 Then
                                '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                If NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                    dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) - NVL(!特准定额, 0), 2)
                                End If
                          Else
                                If dbl比例 <> 0 Then
                                    If !是否医保 = 1 Then
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    End If
                                Else
                                    '100自费部分放入非保险费用中
                                    dbl非保险费用 = dbl非保险费用 + Round(NVL(!结帐金额, 0), 2)
                                End If
                          End If
                Else
                         If gintInsure = TYPE_大连开发区 Then
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) And !是否医保 = 1 Then
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) - NVL(!特准定额, 0), 2)
                                    End If
                                    If NVL(!特准定额, 0) < NVL(!结帐金额, 0) And !是否医保 <> 1 Then
                                        dbl其它费 = dbl其它费 + Round(NVL(!结帐金额, 0) - NVL(!特准定额, 0), 2)
                                    End If
                                Else
                         
                                    If !是否医保 = 1 And dbl比例 <> 0 Then
                                        '险内药品自费  NUM 155 10  医保用药自费部分    院端填写
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    Else
                                        '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
                                        dbl其它费 = dbl其它费 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    End If
                                End If
                         Else
                                If NVL(!算法, 0) = 2 Then
                                    '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算.
                                     If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !是否医保 = 1 And NVL(!特准定额, 0) < NVL(!结帐金额, 0) Then
                                        ''医保用药以及除了大检、特治外检查治疗项目的自费部分
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) - NVL(!特准定额, 0), 2)
                                     End If
                                    If !是否医保 <> 1 Or dbl比例 = 0 Then
                                        '非医保用药以及诊疗项目
                                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                Else
                                    If strTmp <> "特殊治疗费" And strTmp <> "大检费" And !是否医保 = 1 And dbl比例 <> 0 Then
                                        '医保用药以及除了大检、特治外检查治疗项目的自费部分
                                        dbl保险内自费费用 = dbl保险内自费费用 + Round(NVL(!结帐金额, 0) * (1 - dbl比例), 2)
                                    End If
                                    If !是否医保 <> 1 Or dbl比例 = 0 Then
                                        '非医保用药以及诊疗项目
                                        dbl非保险费用 = dbl非保险费用 + Round(NVL(!结帐金额, 0), 2)
                                    End If
                                End If
                         End If
                 End If
            Else
                dbl比例 = 1
                str项目统计分类 = ""
            End If

 
            '上传明细记录,实时医疗明细数据
            If gbln住院明细时实上传 And bln冲销 = False And NVL(!是否上传, 0) = 0 Then
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                        str明细 = str明细 & Lpad(NVL(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
                    Else
                        str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                        str明细 = str明细 & Lpad(NVL(!医保号), 8)   '个人编号    CHAR    5   8       院端
                    End If
                    
                    str明细 = str明细 & Lpad(NVL(!住院号, 0), 10) '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
                    str明细 = str明细 & Lpad(NVL(!顺序号, 0), 4)   '治疗序号    NUM 23  4   住院明细：必须等于入院登记时治疗序号门诊明细:                         必须等于本次结算治疗序号 院端
                    str明细 = str明细 & Lpad(NVL(!NO, 0), 10)       '处方号  NUM 27  10      院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                    Else
                        str明细 = str明细 & Lpad(NVL(!序号, 0), 10)      '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
                    End If
                    
                    '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
                    str明细 = str明细 & Lpad(NVL(!医嘱序号, 0), 10)     '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
                    g病人身份_大连.就诊分类 = NVL(!灰度级, 0)
                    
                    str明细 = str明细 & Get就诊分类(int业务, NVL(!灰度级, 0))         '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then  '开发区
                        '开发区为就诊时间    DATETIME    52  16  精确到秒（开处方时间）格式为：yyyymmddhhmiss后面以空格补位  院端填写
                        str明细 = str明细 & Rpad(Format(!就诊时间, "yyyymmddHHmmss"), 16)
                    Else
                        str明细 = str明细 & Rpad(Format(!登记时间, "yyyymmddHHmmss"), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
                    End If
                    
                    str明细 = str明细 & Lpad(NVL(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
                    str明细 = str明细 & Lpad(NVL(!项目名称), 20)      '项目名称    CHAR    94  20      院端
        
                    If NVL(!险类, 0) = TYPE_大连开发区 Then  '开发区
                    Else
        
                        If !是否医保 = 1 Then
                            str明细 = str明细 & Lpad(1 - dbl比例, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                        Else
                            str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                        End If
                        str明细 = str明细 & Lpad(str项目统计分类, 1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    End If
                    str明细 = str明细 & Lpad(NVL(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
                    str明细 = str明细 & Lpad(NVL(!实际价格), 8) '单价    NUM 127 8   不允许出现负值  院端
                    str明细 = str明细 & Lpad(NVL(!计算单位), 4) '单位    CHAR    135 4       院端
                    str明细 = str明细 & Lpad(NVL(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
                    
                    If NVL(!险类, 0) = TYPE_大连开发区 Then  '开发区
                        '获取病人单量等.
                        gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & NVL(!ID, 0)
                        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                        If rsTemp.EOF Then
                            str明细 = str明细 & Space(5)       '每次用量    NUM 146 5       院端填写
                            str明细 = str明细 & Space(20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                            str明细 = str明细 & Space(50)      '用法    CHAR    171 50  如：口服    院端填写
                        Else
                            str明细 = str明细 & Lpad(NVL(rsTemp!单量), 5)      '每次用量    NUM 146 5       院端填写
                            str明细 = str明细 & Lpad(NVL(rsTemp!频次), 20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                            str明细 = str明细 & Lpad(NVL(rsTemp!用法), 50)      '用法    CHAR    171 50  如：口服    院端填写
                        End If
                        str明细 = str明细 & Space(4)      '执行天数    NUM 221 4       院端填写
                        str明细 = str明细 & Lpad(NVL(!医生编号), 6)      '医师编码    CHAR    225 6       院端填写
                    Else
                        str明细 = str明细 & Lpad(NVL(!医生), 8)      '医师姓名    CHAR    159 8       院端
                    End If
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断编码, 16)      '诊断编码    CHAR    167 16      院端
                    str明细 = str明细 & Lpad(g病人身份_大连.诊断名称, 30)     '诊断名称    CHAR    183 30      院端
                    If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                        str明细 = str明细 & Lpad(NVL(!开单部门), 20)    '科别名称    CHAR    277 20      院端填写
                    Else
                        str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
                    End If
                    
                    '上传明细
                    '1003    7   230 实时医疗明细数据提交
                    住院结算及冲帐_大连 = 业务请求_大连(IIf(NVL(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细)
                    If 住院结算及冲帐_大连 = False Then
                        ShowMsgbox "住院结算或冲帐明细数据提交失败,不能继续!"
                        Exit Function
                    End If
                    '上传医嘱明细
                    If NVL(!医嘱序号, 0) <> 0 Then
                    
                        If 医嘱明细数据提交(!医嘱序号, NVL(!住院号), str项目统计分类) = False Then
                            ShowMsgbox "医嘱明细数据提交失败,不能继续!"
                            Exit Function
                        End If
                    End If
                    '为病人费用记录打上标记，以便随时上传
                    'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                    gstrSQL = "ZL_病人费用记录_更新医保(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
                    zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            End If
            '计算总额,待用
            curTotal = curTotal + Round(NVL(!结帐金额, 0), 2)
            .MoveNext
        Loop
    End With
    
 
    '填写结算记录
    '计算起付线
    dbl起付标准 = g病人身份_大连.起付线
    
    If bln冲销 Then
          '需确定上次中心返回的数据
    
           gstrSQL = "" & _
               "   Select *  " & _
               "   From 保险结算记录 " & _
               "   Where 记录id=" & 原结帐id
           zlDatabase.OpenRecordset rsTemp, gstrSQL, "提取中心收费时返回的数据"
           If rsTemp.RecordCount = 0 Then
               ShowMsgbox "不存在上次收费的结算记录!"
               Exit Function
           End If
           '/???
           '原过程参数:
           '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
           "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
           '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
           '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
           '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
           '过程新值代表为:
           '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
           '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
           '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
           '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,超限自付金额_IN,dbl个人帐户支付
           '       支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
            dbl个人帐户余额 = Round(NVL(rsTemp!帐户累计增加, 0), 2)
            dbl统筹支付累计 = Round(NVL(rsTemp!帐户累计支出, 0), 2)
            dbl补助保险支付 = Round(NVL(rsTemp!累计进入统筹, 0), 2)
            dbl补助帐户支付 = Round(NVL(rsTemp!累计统筹报销, 0), 2)
            dbl起付标准 = Round(NVL(rsTemp!起付线, 0), 2)
            dbl保险范围外自付 = Round(NVL(rsTemp!封顶线, 0), 2)
            dbl基本统筹支付 = Round(NVL(rsTemp!全自付金额, 0), 2)
            dbl基本统筹自付 = Round(NVL(rsTemp!首先自付金额, 0), 2)
            dbl补充统筹支付 = Round(NVL(rsTemp!进入统筹金额, 0), 2)
            dbl补充统筹自付 = Round(NVL(rsTemp!统筹报销金额, 0), 2)
            dbl非补助保险支付 = Round(NVL(rsTemp!大病自付金额, 0), 2)
            dbl个人帐户支付 = Round(NVL(rsTemp!个人帐户支付, 0), 2)
            
            dbl结算前基本帐户余额 = Round(NVL(rsTemp!结算前基本帐户余额, 0), 2)
            dbl结算前补助账户余额 = Round(NVL(rsTemp!结算前补助账户余额, 0), 2)
            dbl结算前统筹累计 = Round(NVL(rsTemp!结算前统筹累计, 0), 2)
              
       End If
    '找出疾病编码
    With g病人身份_大连
        If gintInsure = TYPE_大连开发区 Then    '开发区
            strInfor = Lpad(gstr医院编码_大连, 6)       '医院代码
        Else
            strInfor = Lpad(gstr医院编码_大连, 4)       '医院代码
        End If
        strInfor = strInfor & " "      '子门诊标识
        If gintInsure = TYPE_大连开发区 Then    '开发区
            strInfor = strInfor & Lpad(.个人编号, 10)       '个人编号
        Else
            strInfor = strInfor & Lpad(.个人编号, 8)      '个人编号
        End If
        strInfor = strInfor & Lpad(.IC卡号, 7)       'IC卡号
        strInfor = strInfor & Lpad(.治疗序号 + 1, 4)      '治疗序号
        strInfor = strInfor & Rpad(Format(zlDatabase.Currentdate, "yyyymmddHHmmss"), 16)      '结算时间
        strInfor = strInfor & Lpad(str住院号, 10) '病志号
        
        
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl诊察费), 2))), 10) '诊察费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl草药费), 2))), 10) '草药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl成药费), 2))), 10) '成药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl西药费), 2))), 10) '西药费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl检查费), 2))), 10) '检查费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl治疗费), 2))), 10)  '治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl大检费), 2))), 10)  '大检费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl特殊治疗费), 2))), 10)  '特殊治疗费
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl大检自费), 2))), 10)  '大检自费
        If gintInsure = TYPE_大连开发区 Then        '开发区
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl特殊治疗自费), 2))), 10)   '特治自费    NUM 145 10      院端填写
        End If
        strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl保险内自费费用), 2))), 10)   '保险内自费费用
        
        If gintInsure = TYPE_大连开发区 Then        '开发区
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl其它费), 2))), 10)   '保险外自费  NUM 165 10  非医保用药自费部分  院端填写
        Else
            strInfor = strInfor & Lpad(Trim(CStr(Round(Abs(dbl非保险费用), 2))), 10)    '非保险费用
        End If
    
        Dim dbl结算前余额(1 To 3) As Double '1-结算前基本帐户余额,2-结算前补助账户余额,3-结算前统筹支付累计
        dbl结算前余额(1) = .基本个人帐户余额
        dbl结算前余额(2) = .补助个人帐户余额
        dbl结算前余额(3) = .统筹累计
        
        If bln冲销 Then
            strInfor = strInfor & Lpad(dbl个人帐户余额, 10)
            strInfor = strInfor & Lpad(dbl统筹支付累计, 10)
            strInfor = strInfor & Lpad(dbl结算前基本帐户余额, 10)   '结算前基本帐户余额
            strInfor = strInfor & Lpad(dbl结算前补助账户余额, 10)    '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
            strInfor = strInfor & Lpad(dbl结算前统筹累计, 10)     '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            strInfor = strInfor & Lpad(dbl个人帐户支付, 10) ' = Round(NVL(rsTemp!个人帐户支付, 0), 2)
            strInfor = strInfor & Lpad(dbl补助帐户支付, 10) ' = Round(NVL(rsTemp!累计统筹报销, 0), 2)
            strInfor = strInfor & Lpad(dbl基本统筹支付, 10) ' = Round(NVL(rsTemp!全自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl基本统筹自付, 10) ' = Round(NVL(rsTemp!首先自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补充统筹支付, 10) ' = Round(NVL(rsTemp!进入统筹金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补充统筹自付, 10) ' = Round(NVL(rsTemp!统筹报销金额, 0), 2)
            strInfor = strInfor & Lpad(dbl补助保险支付, 10) ' = Round(NVL(rsTemp!累计进入统筹, 0), 2)
            strInfor = strInfor & Lpad(dbl非补助保险支付, 10) ' = Round(NVL(rsTemp!大病自付金额, 0), 2)
            strInfor = strInfor & Lpad(dbl保险范围外自付, 10) ' = Round(NVL(rsTemp!封顶线, 0), 2)
        Else
            strInfor = strInfor & String(10, " ")    '中心返回:结算后个人帐户余额;开发区:结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
            strInfor = strInfor & String(10, " ")    '中心返回:结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
            '结算前基本帐户余额（根据验卡返回结果，如果是慢病结算应根据慢病查询的慢病帐户余额填写）
            strInfor = strInfor & Lpad(.基本个人帐户余额, 10)  '结算前基本帐户余额
            strInfor = strInfor & Lpad(Trim(CStr(.补助个人帐户余额)), 10)   '结算前补助账户余额(根据验卡返回结果，如果是慢病结算添0)
            strInfor = strInfor & Lpad(Trim(CStr(.统筹累计)), 10)    '结算前统筹支付累计:根据验卡返回结果，如果是慢病结算添0
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本个人帐户支付(如果是慢病结算，表示慢病帐户支付)
            strInfor = strInfor & String(10, " ")    '中心返回:本次补助个人帐户支付(如果是慢病结算返回0)
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次补充统筹自付
            strInfor = strInfor & String(10, " ")    '中心返回:本次基本补助保险支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次非基本补助保险支付
            strInfor = strInfor & String(10, " ")    '中心返回:本次保险范围外自付
        End If
        
        If gintInsure <> TYPE_大连开发区 Then        '开发区
            strInfor = strInfor & Lpad(Trim(CStr(dbl特殊治疗自费)), 10)    '本次特殊治疗自付
        End If
        
        strInfor = strInfor & Lpad(Trim(CStr(dbl起付标准)), 10)    '起付标准
        
        strInfor = strInfor & Lpad(.转诊单号, 6)     '转诊单号
        strInfor = strInfor & Lpad(Get就诊分类(int业务, .就诊分类), 1)     '就诊分类
        If gintInsure <> TYPE_大连开发区 Then
            strInfor = strInfor & Lpad(.参保类别3, 1)    '参保类别3:0 企保、1 事保，根据验卡结果
        End If
        
        strInfor = strInfor & Lpad(.职工就医类别, 1)       '职工就医类别
        
        strInfor = strInfor & Lpad(.诊断编码, 16)    '诊断编码
        strInfor = strInfor & Lpad(str医生, 6)    '医师代码
        strInfor = strInfor & Lpad(UserInfo.编号, 6)    '操作员代码
        strInfor = strInfor & Lpad(.诊断名称, 30)    '诊断名称
        'A-治愈、B-好转、C-未愈、D-死亡、E-其他
        strInfor = strInfor & Lpad(Get治渝情况_大连(lng病人id, lng主页ID), 1)    '治愈情况标识
        strInfor = strInfor & Lpad(str出院日期, 8)      '出院日期
        
        If gintInsure = TYPE_大连开发区 Then        '开发区
        Else
            strInfor = strInfor & String(16, " ")      '传输时间
        End If
        strInfor = strInfor & String(10, " ")      '错误代码
    End With
    '调用1002    12  423 实时结算
    住院结算及冲帐_大连 = 业务请求_大连(IIf(gintInsure = TYPE_大连开发区, 2, 1), 1002, strInfor)
    
    
    '保存结算记录
 
   
    '开发区:
    '   结算后个人帐户余额  NUM 175 10  基本个人帐户＋补助个人帐户  中心返回
    '   结算后统筹支付累计  NUM 185 10  基本统筹累计＋补充统筹累计  中心返回
    
    '    本次基本个人帐户支付    NUM 225 10      中心返回
    '    本次补助个人帐户支付    NUM 235 10      中心返回
    '    本次基本统筹支付    NUM 245 10      中心返回
    '    本次基本统筹自付    NUM 255 10      中心返回
    '    本次补充统筹支付    NUM 265 10      中心返回
    '    本次补充统筹自付    NUM 275 10      中心返回
    '    本次基本补助保险支付    NUM 285 10  公务员补助该字段包括门槛费补助部分和基本统筹自付部分的公务员补助支付 中心返回
    '    本次非基本补助保险支付  NUM 295 10  公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付，该部分（超过基本统筹最高限额部分）除去公务员补助支付后，全部打入"本次保险范围外自付"部分  中心返回
    '    本次保险范围外自付  NUM 305 10  限额以外＋门槛费自付部分（个人帐户充抵后）＋各种自费去掉补助部分    中心返回
    '大连市:
    '   结算后个人帐户余额  NUM 161 10  ①  如果是基本医疗结算表示：基本个人帐户＋补助个人帐户② 如果是慢病结算表示: 慢病帐户结算后余额 中心
    '   结算后统筹支付累计  NUM 171 10  基本统筹累计＋补充统筹累计  中心
    
    '    本次基本个人帐户支付    NUM 211 10  如果是慢病结算，表示慢病帐户支付    中心
    '    本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0 中心
    '    本次基本统筹支付    NUM 231 10      中心
    '    本次基本统筹自付    NUM 241 10      中心
    '    本次补充统筹支付    NUM 251 10  如果是生育结算，本字段用于存放生育保险支付  中心
    '    本次补充统筹自付    NUM 261 10      中心
    '    本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付 2． 如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    '    本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付   2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分    中心
    '    本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    Dim i As Long
    If gintInsure = TYPE_大连开发区 Then
        i = 225 - 10
    Else
        i = 211 - 10
    End If
    
    
    dbl个人帐户余额 = Val(Substr(strInfor, i - 40, 10))
    dbl统筹支付累计 = Val(Substr(strInfor, i - 30, 10))  '结算后统筹支付累计=基本统筹累计＋补充统筹累计
    
    dbl个人帐户支付 = Val(Substr(strInfor, i + 10, 10)) '本次基本个人帐户支付=如果是慢病结算，表示慢病帐户支付
    dbl补助帐户支付 = Val(Substr(strInfor, i + 20, 10))    '本次补助个人帐户支付    NUM 221 10  如果是慢病结算返回0
    dbl基本统筹支付 = Val(Substr(strInfor, i + 30, 10))   '本次基本统筹支付    NUM 231 10      中心
    dbl基本统筹自付 = Val(Substr(strInfor, i + 40, 10))     '本次基本统筹自付    NUM 241 10      中心
    dbl补充统筹支付 = Val(Substr(strInfor, i + 50, 10))     '本次补充统筹支付    NUM 251 10      中心
    dbl补充统筹自付 = Val(Substr(strInfor, i + 60, 10))     '本次补充统筹自付    NUM 261 10      中心
    dbl补助保险支付 = Val(Substr(strInfor, i + 70, 10))     '本次基本补助保险支付    NUM 271 10  1． 如果是商业保险该字段包括基本统筹自付部分的商业保险支付2．   如果是公务员补助该字段包括门槛费补助部分、基本统筹自付部分的公务员补助支付；基本统筹最高限额内公务员补助支付后剩余款项计入"本次保险范围外自付"部分  中心
    dbl非补助保险支付 = Val(Substr(strInfor, i + 80, 10))     '本次非基本补助保险支付  NUM 281 10  1． 如果是商业保险该字段是补充统筹自付部分的商业保险支付2． 如果是公务员补助该字段是超过基本统筹最高限额部分的公务员补助支付；超过基本统筹最高限额公务员补助支付后，剩余款项计入"本次保险范围外自付"部分
    dbl保险范围外自付 = Val(Substr(strInfor, i + 90, 10))     '本次保险范围外自付  NUM 291 10  限额以外（去掉补助后）＋门槛费自付部分（个人帐户充抵后）＋保险内自费费用＋非保险费用+大检自费   中心
    
    '/???
       '原过程参数:
       '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
       "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
       '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
       '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
       '    支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN,
       '    诊察费_IN,草药费_IN,成药费_IN,西药费_IN,检查费_IN,治疗费_IN,大检费_IN,大检自费_IN,特殊治疗费_IN,特殊治疗自费_IN,保险内自费费用_IN,非保险费用_IN,统筹比例_IN,其它费
        '   结算前基本帐户余额,结算前补助账户余额,结算前统筹累计
       '过程新值代表为:
       '       性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN, _
       '       dbl个人帐户余额,dbl统筹支付累计,dbl补助保险支付,dbl补助帐户支付,住院次数_IN,起付线_IN,dbl保险范围外自付,实际起付线_IN
       '       发生费用金额_IN,dbl基本统筹支付,dbl基本统筹自付,
       '       dbl补充统筹支付,dbl补充统筹自付,dbl非补助保险支付,dbl个人帐户支付
       '       支付顺序号_IN(就诊分类;转诊单号;诊断编码),主页ID_IN,中途结帐_IN,诊断名称_IN
       '    诊察费_IN,草药费_IN,成药费_IN,西药费_IN,检查费_IN,治疗费_IN,大检费_IN,大检自费_IN,特殊治疗费_IN,特殊治疗自费_IN,保险内自费费用_IN,非保险费用_IN,统筹比例_IN,其它费
        '   结算前基本帐户余额,结算前补助账户余额,结算前统筹累计
              
       gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人id & "," & Format(zlDatabase.Currentdate, "YYYY") & "," & _
          dbl个人帐户余额 & "," & dbl统筹支付累计 & "," & dbl补助保险支付 & "," & dbl补助帐户支付 & "," & "Null" & "," & dbl起付标准 & "," & dbl保险范围外自付 & "," & dbl起付标准 & "," & _
          curTotal & "," & dbl基本统筹支付 & "," & dbl基本统筹自付 & "," & _
          dbl补充统筹支付 & "," & dbl补充统筹自付 & "," & dbl非补助保险支付 & ",Null," & dbl个人帐户支付 & ",'" & _
          Get就诊分类(int业务, g病人身份_大连.就诊分类) & ";" & g病人身份_大连.转诊单号 & ";" & g病人身份_大连.诊断编码 & "'," & lng主页ID & ",null,'" & g病人身份_大连.诊断名称 & "'," & _
           dbl诊察费 & "," & dbl草药费 & "," & dbl成药费 & "," & dbl西药费 & "," & dbl检查费 & "," & dbl治疗费 & "," & dbl大检费 & "," & dbl大检自费 & "," & dbl特殊治疗费 & "," & dbl特殊治疗自费 & "," & dbl保险内自费费用 & "," & dbl非保险费用 & "," & dbl比例 & "," & dbl其它费 & "," & _
            dbl结算前余额(1) & "," & dbl结算前余额(2) & "," & dbl结算前余额(3) & _
            " )"
            
        zlDatabase.ExecuteProcedure gstrSQL, "保存住院结帐收费数据"
        Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 住院结算_大连(lng结帐ID As Long, ByVal lng病人id As Long) As Boolean

    Dim cur个人帐户 As Currency
    Dim lng主页ID As Long
    Dim blnError As Boolean
    Dim str入院年份 As String, str结算年份 As String
    Dim str经办时间 As String, str结算时间 As String
    Dim str就诊编号 As String
    Dim rsTemp As New ADODB.Recordset
    '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    '虚拟结算（返回的数据减去历次结算数据，就等于本次的真实结算数据）
    On Error GoTo ErrHand
    
    Call 个人余额_大连(lng病人id)

    cur个人帐户 = g病人身份_大连.基本个人帐户余额

    gstrSQL = " Select B.住院次数 主页ID,to_char(A.入院日期,'yyyy') 入院年份 " & _
              " From 病案主页 A,病人信息 B" & _
              " Where B.病人ID=" & lng病人id & " And A.主页ID=B.住院次数 And A.病人ID=B.病人ID"
    Call OpenRecordset(rsTemp, "获取病人入院时间")
    str入院年份 = rsTemp!入院年份
    lng主页ID = rsTemp!主页ID

    str经办时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str结算时间 = str经办时间
    str结算年份 = Mid(str经办时间, 1, 4)

    住院结算_大连 = 住院结算及冲帐_大连(False, lng病人id, lng结帐ID, lng结帐ID, lng主页ID)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function Get起付线(ByVal str职工就医类别 As String, ByVal lng年龄 As Long) As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取起付线
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------

       Dim strCaption As String
       Dim rsTmp As New ADODB.Recordset
       strCaption = Decode(str职工就医类别, "A", "在职", "B", "退休", "L", "离休", "T", "特诊", "Q", "企业公费", "在职")
    
        gstrSQL = "" & _
            "   Select d.金额*a.比例/100 as 起付线" & _
            "   From 保险支付比例 a,保险人群 b, " & _
            "      (Select * From 保险年龄段  " & _
            "       where ((" & lng年龄 & ">=下限 and " & lng年龄 & "<=上限) or (" & lng年龄 & ">下限 and 上限=0) ) and 险类=" & gintInsure & _
            "       ) c,保险支付限额 d " & _
            " where a.险类=" & gintInsure & " and b.险类 =a.险类 and a.在职=b.序号 and b.名称='" & strCaption & "' and  " & _
            "       a.年龄段=c.年龄段 and a.在职=c.在职 and a.险类=d.险类 and d.年度='" & Format(zlDatabase.Currentdate, "yyyy") & "' and d.性质='1'"
    
       Err = 0
       On Error GoTo ErrHand:
       zlDatabase.OpenRecordset rsTmp, gstrSQL, "计算起付线"
       If Not rsTmp.EOF Then
            Get起付线 = NVL(rsTmp!起付线, 0)
       Else
            Get起付线 = 0
       End If
       Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
       Get起付线 = 0
   
End Function
Public Function 住院结算冲销_大连(lng结帐ID As Long) As Boolean
    Dim lng冲销ID As Long
    Dim str退单编号 As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人id As Long
    Dim lng主页ID As Long
    
    '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '      4)只能作废当月离退体人员的结帐单据
    '----------------------------------------------------------------
    On Error GoTo ErrHand
    gstrSQL = "select distinct A.ID from 病人结帐记录 A,病人结帐记录 B " & _
              " where A.NO=B.NO and  A.记录状态=2 and B.ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "大连医保")
    lng冲销ID = rsTemp("ID") '冲销单据的ID

    '为了将当时写卡的金额读出，故再次访问记录
    gstrSQL = "Select * " & _
              "  From 保险结算记录 Where 性质=2 and 记录ID='" & lng结帐ID & "'"
    Call OpenRecordset(rsTemp, "大连医保")
    If rsTemp.EOF Then
        ShowMsgbox "在保险结算记录中无该结算记录!"
        Exit Function
    End If
    lng病人id = NVL(rsTemp!病人ID, 0)
    lng主页ID = NVL(rsTemp!主页ID, 0)
        
        
    '重新读卡
    If 读取病人身份_大连(IIf(gintInsure = TYPE_大连开发区, 2, 1)) = False Then
        Exit Function
    End If
    
    Dim strArr
    strArr = Split(NVL(rsTemp!支付顺序号), ";")
    
    '就诊分类;转诊单号;诊断编码
    '5-普通住院("2", "D"),6-家庭病床住院("4", "C")
    '7-生育保险住院("O", "P"),8-工伤保险住院("Q", "R")
    With rsTemp
        If UBound(strArr) >= 2 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g病人身份_大连.转诊单号 = strArr(1)
            g病人身份_大连.诊断编码 = strArr(2)
        ElseIf UBound(strArr) = 1 Then
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
            g病人身份_大连.转诊单号 = strArr(1)
        Else
            g病人身份_大连.就诊分类 = Decode(strArr(0), "2", 5, "D", 5, "4", 6, "C", 6, "0", 7, "P", 7, 8)
        End If
        g病人身份_大连.诊断名称 = NVL(rsTemp!备注)
    End With
    
    '验证是否为该病人的IC卡
    gstrSQL = "Select * From  保险帐户 where 病人id=" & lng病人id
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "读取病人的医保号"
    If rsTemp.EOF Then
        ShowMsgbox "该病人在保险帐户中无记录!"
        Exit Function
    End If
    
    If g病人身份_大连.IC卡号 <> NVL(rsTemp!卡号) Then
        ShowMsgbox "该病人的IC卡插入错误,可能是插入了其他人的IC卡!"
        Exit Function
    End If
    '调用撤销结算接口
    住院结算冲销_大连 = 住院结算及冲帐_大连(True, lng病人id, lng冲销ID, lng结帐ID, lng主页ID)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保终止_大连() As Boolean
    医保终止_大连 = True
End Function

Public Function 处方登记_大连(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '先写入单据头，再写入单据体
    '记录状态（1-新增;否则为删除），费用处单据只能整张单据删除后，再产生新单据
    On Error GoTo ErrHand
    处方登记_大连 = False
    If gbln住院明细时实上传 = False Then
        处方登记_大连 = True
        Exit Function
    End If
    With rsTemp
        If .State = 1 Then .Close
        .CursorLocation = adUseClient

        gstrSQL = " " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,C.统筹比额,F.住院次数 AS 主页id, " & _
            "        G.标识主码||G.标识子码 AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 病人费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连市 & " and M.大类ID=Q.id) C " & _
            " Where     a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=82 AND F.住院次数= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连市 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=" & lng记录性质 & " and  A.记录状态=" & lng记录状态 & " And A.NO='" & str单据号 & "'" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 "
        
        gstrSQL = gstrSQL & " Union all " & _
            " Select A.id,A.病人ID,F.住院号,A.NO,A.序号,A.医嘱序号,A.记录性质,A.记录状态,A.收费类别,D.类别,to_char(A.登记时间,'yyyyMMddhh24miss') 登记时间, " & _
            "        A.开单人 医生,V.编号 AS 医生编号,B.名称 开单部门,A.收费细目ID,A.计算单位,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.实收金额 金额,A.数次*Nvl(A.付数,1) 数量,Nvl(A.是否上传,0) 是否上传, " & _
            "        C.项目编码 医保项目编码 ,C.是否医保,C.统筹比额,F.住院次数 AS 主页id, " & _
            "        G.标识主码||G.标识子码 AS 国家编码,G.名称 AS 项目名称,K.名称 AS 剂型, " & _
            "        E.险类,E.中心,E.卡号,E.医保号,E.密码,E.人员身份,E.单位编码,E.顺序号,E.退休证号,E.帐户余额,E.当前状态, " & _
            "        E.病种ID,E.在职,E.年龄段,E.灰度级,to_char(E.就诊时间,'yyyyMMddhh24miss') 就诊时间 " & _
            " From 病人费用记录 A,部门表 B,收费类别 D,保险帐户 E,病人信息 F,病案主页 F1,收费细目 G,人员表 V," & _
            "       (Select J.名称,O.药品id From 药品目录 O, 药品信息 H,药品剂型 J WHERE O.药名id=H.药名id and H.剂型=J.编码) K, " & _
            "       (Select M.项目编码,M.项目名称,M.是否医保,M.收费细目id,Q.统筹比额  From 保险支付项目 M,保险支付大类 Q Where M.险类=" & TYPE_大连开发区 & " and M.大类ID=Q.id) C " & _
            " Where     a.病人id=E.病人ID AND a.病人id=F.病人ID AND A.病人id=F1.病人id and F1.险类=83 AND F.住院次数= F1.主页id  AND  a.开单人=V.姓名(+) AND a.收费细目id=k.药品id(+) AND a.收费细目id=G.id AND E.险类=" & TYPE_大连开发区 & "   AND A.收费类别=D.编码 AND  " & _
            "           A.记录性质=" & lng记录性质 & " and  A.记录状态=" & lng记录状态 & " And A.NO='" & str单据号 & "'" & _
            "           And A.开单部门ID+0=B.ID And A.收费细目ID+0=C.收费细目ID(+) And Nvl(A.是否上传,0)=0 " & _
            " Order by 病人ID"
            
        Call OpenRecordset(rsTemp, "处方登记")
        
        If .RecordCount = 0 Then
            MsgBox "未找到处方记录，向医保服务器传输数据失败！[处方登记]", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    处方登记_大连 = 上传处方_大连(rsTemp)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 上传处方_大连(ByVal rsExse As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:rsExse-明细数据
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------


    Dim lng病人id As Long
    Dim curTotal As Currency
    Dim blnUpload As Boolean
    Dim rsPara As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim str明细 As String
    Dim str项目统计分类 As String
    Dim strTmp As String
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "select * from 保险参数 where 险类 in (82,83)"
    zlDatabase.OpenRecordset rsPara, gstrSQL, "上传处方读取参数"
    With rsExse
        Do While Not .EOF
            lng病人id = NVL(!病人ID, 0)
            '确定相关数据
            '上传明细记录,实时医疗明细数据
                
            If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                str明细 = Lpad(gstr医院编码_大连, 6)     '医院代号    CHAR    1   6       院端填写
                str明细 = str明细 & Lpad(NVL(!医保号), 10)  '保险编号    CHAR    7   10      院端填写
            Else
                str明细 = Lpad(gstr医院编码_大连, 4)     '医院代码    CHAR    1   4       院端
                str明细 = str明细 & Lpad(NVL(!医保号), 8)   '个人编号    CHAR    5   8       院端
            End If
            
            str明细 = str明细 & Lpad(NVL(!住院号, 0), 10) '病志号  CHAR    13  10  门诊明细以空格补位,住院是住院号  院端
            str明细 = str明细 & Lpad(NVL(!顺序号, 0), 4)   '治疗序号    NUM 23  4   住院明细：必须等于入院登记时治疗序号门诊明细:                         必须等于本次结算治疗序号 院端
            str明细 = str明细 & Lpad(NVL(!NO, 0), 10)       '处方号  NUM 27  10      院端
            
            If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
            Else
                str明细 = str明细 & Lpad(CStr(NVL(!序号, 0)), 10)      '处方项目序号    NUM 37  10  对应处方号的记价项目序号    院端
            End If
            
            '开发区为单据号  CHAR    41  10  医嘱号，    院端填写
            str明细 = str明细 & Lpad(NVL(!医嘱序号, 0), 10)     '医嘱号  CHAR    47  10  处方对应医嘱的医嘱记录号，门诊明细或没有医嘱的医院以空格补位    院端
            
            str明细 = str明细 & Get就诊分类(0, NVL(!灰度级))      '就诊分类    CHAR    57  1   取值详见"就诊分类"说明  院端
            
            If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
                '开发区为就诊时间    DATETIME    52  16  精确到秒（开处方时间）格式为：yyyymmddhhmiss后面以空格补位  院端填写
                str明细 = str明细 & Rpad(NVL(!就诊时间), 16)
            Else
                str明细 = str明细 & Rpad(NVL(!登记时间), 16)      '处方生成时间（投药时间）    DATETIME    58  16  精确到秒格式为：yyyymmddhhmiss后面以空格补位    院端
            End If
            
            str明细 = str明细 & Lpad(NVL(!国家编码), 20)      '项目代码    CHAR    74  20  计价项目代码    院端
            str明细 = str明细 & Lpad(NVL(!项目名称), 20)      '项目名称    CHAR    94  20      院端

            If NVL(!险类, 0) = TYPE_大连开发区 Then '开发区
            Else

                If !是否医保 = 1 Then
                    str明细 = str明细 & Lpad(1 - NVL(!统筹比额, 0), 6)   '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                Else
                    str明细 = str明细 & Lpad(1, 6)    '自费比例 Char 114 6   如果是保险范围内费用，自费比例可能为：0或者0.1（0％或10％）等 如果是保险范围外用药自费比例为：1（100％）  院端
                End If
                rsPara.Filter = 0
                rsPara.Filter = " 参数名='" & NVL(!类别) & "' and 险类=" & NVL(!险类, 0)
                str项目统计分类 = ""
                If Not rsPara.EOF Then
                    strTmp = NVL(rsPara!参数值)
                    If InStr(1, strTmp, ";") <> 0 And strTmp <> ";" Then
                        strTmp = Split(strTmp, ";")(1)
                        If strTmp <> "" Then
                            str项目统计分类 = Substr(strTmp, 1, 1)
                            str明细 = str明细 & Substr(strTmp, 1, 1)   '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                        Else
                            str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                        End If
                    Else
                        str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                    End If
                Else
                        str明细 = str明细 & Space(1)    '项目统计分类    CHAR    120 1   详见注①,具体实现方式?  院端
                End If
            End If
            
            str明细 = str明细 & Lpad(NVL(!数量), 6)  '数量    NUM 121 6   冲方划价为负值  院端
            str明细 = str明细 & Lpad(NVL(!实际价格), 8) '单价    NUM 127 8   不允许出现负值  院端
            str明细 = str明细 & Lpad(NVL(!计算单位), 4) '单位    CHAR    135 4       院端
            str明细 = str明细 & Lpad(NVL(!剂型), 20)      '剂型    CHAR    139 20  针剂、片剂…    院端
            
            If NVL(!险类, 0) = TYPE_大连开发区 Then  '开发区
                '获取病人单量等.
                gstrSQL = "Select 单量,频次,用法 From 药品收发记录 where 费用id=" & NVL(!ID, 0)
                zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人单理及频次"
                If rsTemp.EOF Then
                    str明细 = str明细 & Space(5)       '每次用量    NUM 146 5       院端填写
                    str明细 = str明细 & Space(20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                    str明细 = str明细 & Space(50)      '用法    CHAR    171 50  如：口服    院端填写
                Else
                    str明细 = str明细 & Lpad(NVL(rsTemp!单量), 5)      '每次用量    NUM 146 5       院端填写
                    str明细 = str明细 & Lpad(NVL(rsTemp!频次), 20)      '使用频次    CHAR    151 20  如：1天2次  院端填写
                    str明细 = str明细 & Lpad(NVL(rsTemp!用法), 50)      '用法    CHAR    171 50  如：口服    院端填写
                End If
                str明细 = str明细 & Space(4)      '执行天数    NUM 221 4       院端填写
                str明细 = str明细 & Lpad(NVL(!医生编号), 6)      '医师编码    CHAR    225 6       院端填写
            Else
                str明细 = str明细 & Lpad(NVL(!医生), 8)      '医师姓名    CHAR    159 8       院端
            End If
            '确定诊断情况
            
            strTmp = Get入院诊断(NVL(!病人ID), NVL(!主页ID, 0), False, True)
            If InStr(1, strTmp, "|") <> 0 Then
                
                str明细 = str明细 & Lpad(Split(strTmp, "|")(1), 16)     '诊断编码    CHAR    167 16      院端
                strTmp = Split(strTmp, "|")(0)
                strTmp = Lpad(strTmp, 30)
                strTmp = Substr(strTmp, 1, 30)
                str明细 = str明细 & strTmp     '诊断名称    CHAR    183 30      院端
            Else
                str明细 = str明细 & Space(16)      '诊断编码    CHAR    167 16      院端
                str明细 = str明细 & Space(30)     '诊断名称    CHAR    183 30      院端
            End If
            
            If NVL(!险类, 0) = TYPE_大连开发区 Then  '开发区
                str明细 = str明细 & Lpad(NVL(!开单部门), 20)    '科别名称    CHAR    277 20      院端填写
            Else
                str明细 = str明细 & Space(16)     '传输时间    DATETIME    213 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，院端空格补位  中心
            End If
            
            '上传明细
            '1003    7   230 实时医疗明细数据提交
            上传处方_大连 = 业务请求_大连(IIf(NVL(!险类, 0) = TYPE_大连开发区, 2, 1), 1003, str明细)
            If 上传处方_大连 = False Then
                ShowMsgbox "门诊结算时医疗明细数据提交失败,不能继续!"
                Exit Function
            End If
            '上传医嘱明细
            If NVL(!医嘱序号, 0) <> 0 Then
                上传处方_大连 = False
                If 医嘱明细数据提交(NVL(!医嘱序号, 0), NVL(!住院号), str项目统计分类) = False Then
                    ShowMsgbox "医嘱明细数据提交失败,不能继续!"
                    Exit Function
                End If
            End If

            '为病人费用记录打上标记，以便随时上传
            'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
            gstrSQL = "ZL_病人费用记录_更新医保(" & NVL(!ID, 0) & ",NULL,NULL,NULL,NULL,1,Null)"
            zlDatabase.ExecuteProcedure gstrSQL, "打上上传标志"
            .MoveNext
        Loop
    End With
    上传处方_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function Decode(ParamArray arrPar() As Variant) As Variant

'功能：模拟Oracle的Decode函数

    Dim varValue As Variant, i As Integer

    

    i = 1

    varValue = arrPar(0)

    Do While i <= UBound(arrPar)

        If i = UBound(arrPar) Then

            Decode = arrPar(i): Exit Function

        ElseIf varValue = arrPar(i) Then

            Decode = arrPar(i + 1): Exit Function

        Else

            i = i + 2

        End If

    Loop

End Function

Private Function 医嘱明细数据提交(ByVal lng医嘱ID As Long, ByVal str住院号 As String, ByVal str项目统计分类 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:提取医嘱明细
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    
    '5.  实时医嘱数据提交接口

    '开发区无医嘱接口
    If gintInsure = TYPE_大连开发区 Then
        医嘱明细数据提交 = True
        Exit Function
    End If
    gstrSQL = " " & _
         " select ID,父序号 as 分组号,decode(期效,1,1,0) as  医嘱类型,药品单量,剂量单位,执行频次,频率次数,医嘱内容, " & _
         "        下医嘱医生,下医嘱时间,to_char(开始执行时间,'yyyymmddhh24miss') as 开始执行时间,校对护士 as 执行医嘱护士姓名," & _
         "        停医嘱医生,to_char(停医嘱时间,'yyyymmddhh24miss') as 停医嘱时间,附加说明" & _
         " from 医嘱记录  " & _
         " Where id=" & lng医嘱ID
    Err = 0
    On Error GoTo ErrHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医嘱明细记录"
    If rsTemp.EOF Then
        ShowMsgbox "无对应的医嘱记录!"
        Exit Function
    End If
    With g病人身份_大连
        strInfor = Lpad(gstr医院编码_大连, 4)             '医院代码    CHAR    1   4       院端
        strInfor = strInfor & Lpad(.个人编号, 8)    '个人编号    CHAR    5   8       院端
        strInfor = strInfor & Lpad(.治疗序号, 4)    '治疗序号    NUM 13  4   必须等于入院时治疗序号  院端
        strInfor = strInfor & Lpad(str住院号, 10)    '病志号  CHAR    17  10      院端
        strInfor = strInfor & Lpad(lng医嘱ID, 10)     '医嘱号  CHAR    27  10      院端
        
        '注②：医嘱分组号用来表示医嘱中同时使用的项目，例如：有两条医嘱记录分别为青霉素和氯化钠注射液，医生要求两种药物同时给患者使用，此时就可以将两条记录的分组号设为相同的值，值的内容不同医院可以根据自身医嘱系统的具体情况自定，只要能用2位字符标识出对该患者同时使用的记价项目即可。
        strInfor = strInfor & Lpad(lng医嘱ID, 10)     '医嘱分组号  CHAR    37  3   详见注②
        strInfor = strInfor & Lpad(NVL(rsTemp!医嘱类型, 0), 1)   '医嘱类型    CHAR    40  1   1 长嘱，0 临时医嘱
        strInfor = strInfor & Space(20)   '项目代码    CHAR    41  20  计价项目代码，对于描述性医嘱例如：明日出院等 项目代码统一用'000000' 院端
        strInfor = strInfor & Space(20)   '项目名称    CHAR    61  20      院端
        strInfor = strInfor & Lpad(str项目统计分类, 1)  '项目统计分类    CHAR    81  1   详见注①    院端
        strInfor = strInfor & Lpad(NVL(rsTemp!药品单量, 0), 15) '每次用量    CHAR    82  15  例如：10    院端
        strInfor = strInfor & Lpad(NVL(rsTemp!剂量单位), 4) '剂量单位    CHAR    97  4   例如：ml
        strInfor = strInfor & Lpad(NVL(rsTemp!执行频次), 20) '使用频次    CHAR    101 20  如：1天2次  院端
        strInfor = strInfor & Substr(Lpad(NVL(rsTemp!医嘱内容), 50), 1, 50) '用法    CHAR    121 50  如：口服；静脉滴注（20滴/分钟）。。。。 院端
        strInfor = strInfor & Lpad(NVL(rsTemp!下医嘱医生), 8) '下医嘱医师姓名  CHAR    171 8       院端
        strInfor = strInfor & Rpad(NVL(rsTemp!开始执行时间), 16) '开始执行时间    DATATIME    179 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，此项必添  院端
        strInfor = strInfor & Lpad(NVL(rsTemp!执行医嘱护士姓名), 8) '执行医嘱护士姓名    CHAR    195 8       院端
        strInfor = strInfor & Lpad(NVL(rsTemp!停医嘱医生), 8) '终止医嘱医师姓名    CHAR    203 8       院端
        strInfor = strInfor & Rpad(NVL(rsTemp!停医嘱时间), 16) '终止医嘱时间    DATATIME    211 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，对于长期医嘱此项必添对于临时医嘱  院端
        strInfor = strInfor & Substr(Lpad(NVL(rsTemp!附加说明), 30), 1, 30) '备注    CHAR    227 30  用于临时医嘱执行反馈或者其他描述    院端
        strInfor = strInfor & Space(16)                  '传输时间    DATATIME    257 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  中心
    End With
    '1005    8   274 实时医嘱传输
    医嘱明细数据提交 = 业务请求_大连(g病人身份_大连.医保中心, 1005, strInfor)
    Exit Function
ErrHand:
    '如果没装医嘱就不执行
    医嘱明细数据提交 = True
End Function

Private Function Get病人变动记录(ByVal lng病人id As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人的变动情况
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "" & _
        "   Select  床号,附加床位,开始时间,终止时间,床位等级id " & _
        "   From 病人变动记录  " & _
        "   Where  病人id=" & lng病人id & " and 主页id=" & lng主页ID & " and 床号 is not null"
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人变动情况"
    Set Get病人变动记录 = rsTemp
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Set Get病人变动记录 = Nothing
    Exit Function
End Function
Private Function Get住院虚拟记录(ByVal lng病人id As Long) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取本次虚拟未结记录
    '--入参数:
    '--出参数:
    '--返  回:未结费用
    '-----------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset

    '
    strSql = _
        "   Select  A.记录性质,A.记录状态,A.NO,A.序号,A.床号," & _
        "           A.病人ID,A.主页ID,A.婴儿费," & _
        "           A.保险大类ID,A.收费类别,A.收费细目ID,B.名称 as 收费名称,X.名称 as 开单部门," & _
        "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
        "           Decode(Sign(Instr(B.规格,'┆')),0,B.规格,Substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
        "           A.数量,Decode(A.数量,0,0,Round(A.金额/A.数量,4)) as 价格,A.金额,A.医生,w.编号 as 医生编号,A.登记时间," & _
        "           A.是否上传,A.是否急诊,A.保险项目否,A.摘要,C.项目编码 as 医保项目编码," & _
        "           C.项目名称 as 医保项目名称,Q.参数值,Q.参数名,J.统筹比额,J.住院比额,J.特准定额,J.算法" & _
        "   From (" & _
        "           Select  Mod(A.记录性质,10) as 记录性质,A.记录状态,A.床号,A.NO,Nvl(A.价格父号,序号) as 序号,A.病人ID,A.主页ID,Nvl(A.婴儿费,0) as 婴儿费," & _
        "                   A.开单人 as 医生,A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0) as 保险大类ID,Avg(Nvl(A.付数,1)*A.数次) as 数量," & _
        "                   Sum(A.标准单价) as 标准单价,Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0)) as 金额,A.登记时间,Nvl(A.是否上传,0) as 是否上传,Nvl(A.是否急诊,0) as 是否急诊,Nvl(A.保险项目否,0) as 保险项目否,A.摘要" & _
        "           From 病人费用记录 A,收入项目 B" & _
        "           Where A.记帐费用=1 And A.收入项目ID=B.ID And A.病人ID=" & lng病人id & _
        "           Group by    Mod(A.记录性质,10),A.记录状态,A.NO,Nvl(A.价格父号,序号),A.病人ID,A.主页ID,A.床号,Nvl(A.婴儿费,0),A.开单人," & _
        "                       A.开单部门ID,A.收费类别,A.收费细目ID,Nvl(A.保险大类ID,0),A.登记时间,Nvl(A.是否上传,0),Nvl(A.是否急诊,0),Nvl(A.保险项目否,0),A.摘要" & _
        "           Having Sum(Nvl(A.实收金额,0))-Sum(Nvl(A.结帐金额,0))<>0) A,收费细目 B,部门表 X," & _
        "           (Select * From 保险支付项目 Where 险类=" & gintInsure & ") C," & _
        "           (Select M.编码, L.参数名,L.参数值 from 收费类别 M,保险参数 L  Where M.类别=L.参数名 and L.险类=" & gintInsure & ")  Q," & _
        "           (Select * from 保险支付大类  Where 险类=" & gintInsure & ")  J,人员表 W" & _
        "   Where     A.收费细目ID=B.ID and a.医生=w.姓名(+) and C.大类id=J.ID and a.收费类别=Q.编码(+) And A.收费细目ID=C.收费细目ID And A.开单部门ID=X.ID"
    Err = 0
    On Error GoTo ErrHand:
    zlDatabase.OpenRecordset rsTmp, strSql, "获取本次医保未结费用"
    Set Get住院虚拟记录 = rsTmp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Set Get住院虚拟记录 = Nothing
    Exit Function
End Function


Private Function Set门诊挂号结算或冲销(ByVal bln冲销 As Boolean, lng结帐ID As Long, cur个人帐户 As Currency, lng病人id As Long, strSelfNo As String) As Boolean
  '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；
    '      cur个人帐户   从个人帐户中支出的金额
    
    Set门诊挂号结算或冲销 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 挂号结算_大连(ByVal lng结帐ID As Long) As Boolean
     挂号结算_大连 = Set门诊挂号结算或冲销(False, lng结帐ID, 0, 0, 0)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function 挂号冲销_大连(ByVal lng结帐ID As Long) As Boolean
    挂号冲销_大连 = Set门诊挂号结算或冲销(False, lng结帐ID, 0, 0, 0)
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function 在院病人信息_大连(lng病人id As Long, lng主页ID As Long) As Boolean
    Dim str入院经办时间 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strInfor As String
    Dim str就诊分类 As String
    Dim str入院科室 As String
    Dim str床位号 As String
    Dim str转诊单号 As String
    Dim lng中心 As Long
    
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    
    On Error GoTo ErrHand
    
    '读取病人的相关保险信息

    gstrSQL = "select * From 保险帐户 where  险类=" & gintInsure & "  and 病人id=" & lng病人id
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "入院读取保险帐户信息"
    If rsTemp.EOF Then
        ShowMsgbox "在保险帐户中无该病人的保险信息!"
        Exit Function
    End If
    str转诊单号 = NVL(rsTemp!人员身份)
    lng中心 = IIf(gintInsure = 83, 2, 1)
    If lng中心 = 2 Then
        strInfor = Lpad(gstr医院编码_大连, 6) '医院代码    CHAR    1   6      Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 10)     '保险编号    CHAR    7   10      院端填写
    Else
        strInfor = Lpad(gstr医院编码_大连, 4) '医院代码    CHAR    1   4       Y   院端
        strInfor = strInfor & Lpad(NVL(rsTemp!医保号), 8)     '保险编号    CHAR    5   8       Y   院端
    End If
    
    strInfor = strInfor & Lpad(NVL(rsTemp!顺序号, 1), 4)      '治疗序号    NUM 13  4   必须等于入院时治疗序号  Y   院端
    
    
    '内部标识:5-普通住院,6-家庭病床住院,7-生育保险住院,8-工伤保险住院
    '医保标识:2-住院结算,4-家庭病床结算,O-生育保险住院结算,Q-工伤保险结算
    
    str就诊分类 = Decode(NVL(rsTemp!灰度级, 0), 5, "2", 6, "4", 7, "O", 8, "Q", "2")
    '读取病人信息
    gstrSQL = "Select C.住院号,C.当前病区id,C.当前床号,A.登记人 经办人,B.名称 入院科室,to_char(A.登记时间,'yyyyMMddhh24miss') 入院经办时间," & _
            " to_char(A.登记时间,'yyyyMMdd') 入院日期" & _
            " From 病案主页 A,部门表 B,病人信息 C" & _
            " Where A.病人id=C.病人id and C.病人id=" & lng病人id & _
            "       and A.病人ID=" & lng病人id & " And A.主页ID=" & lng主页ID & " And A.入院科室ID=B.ID"
            
    Call OpenRecordset(rsTemp, "读取入院信息")
    If rsTemp.EOF Then
        ShowMsgbox "在病案主页中无此病人!"
        Exit Function
    End If
    
    str入院科室 = NVL(rsTemp!入院科室)
    
    strInfor = strInfor & Lpad(NVL(rsTemp!住院号, 0), 10)      '病志号  CHAR    17  10      Y   院端门诊暂定为空，住院就为住院号
    strInfor = strInfor & Lpad(NVL(rsTemp!入院日期), 8)      '入院日期 Date 27  8   患者实际入院日期，格式为yyyymmdd    Y   院端
    strInfor = strInfor & Rpad(NVL(rsTemp!入院经办时间), 16)     '登记时间    DATETIME    35  16  精确到秒，数据返回后格式为yyyymmddhhmiss后面以空格补位  Y   院端
    If lng中心 = 2 Then
        '开发区为:住院 2、家床 4取消住院登记 C
        strInfor = strInfor & IIf(str就诊分类 = "4", "4", "2")
    Else
        strInfor = strInfor & Lpad(str就诊分类, 1)                  '就诊分类    CHAR    51  1   2住院、4家床、O生育、   Y   院端
    End If

    gstrSQL = "Select * From 床位状况记录 D where 病区ID=" & NVL(rsTemp!当前病区ID, 0) & " And 床号=" & NVL(rsTemp!当前床号, 0)
    Call OpenRecordset(rsTemp, "读取床位信息")
    If rsTemp.EOF Then
        str床位号 = Space(10)
    Else
        str床位号 = Trim(NVL(rsTemp!房间号)) & "室" & Trim(NVL(rsTemp!床号)) & "床"
        str床位号 = Lpad(str床位号, 10)
        str床位号 = Substr(str床位号, 1, 10)
    End If
    
    gstrSQL = "" & _
         " select max(decode(A.诊断类型,1,b.编码||'~^||'||b.名称,null)) as 入院诊断,  " & _
         "        max(decode(A.诊断类型,1,null,b.编码||'~^||'||b.名称)) as 确诊诊断 " & _
         " from 诊断情况 A,疾病编码目录 b " & _
         " where a.疾病id=b.id and  a.诊断类型 in(1,2) and a.诊断次序=1"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定诊断编码和名称"
    Dim str入院诊断编码 As String
    Dim str入院诊断名称  As String
    Dim str确诊诊断编码 As String
    Dim str确诊诊断名称  As String
    
    If rsTemp.EOF Then
        str入院诊断编码 = ""
        str入院诊断名称 = ""
        str确诊诊断编码 = ""
        str确诊诊断名称 = ""
    Else
        str入院诊断名称 = NVL(rsTemp!入院诊断)
        str确诊诊断名称 = NVL(rsTemp!确诊诊断)
        If InStr(1, str入院诊断名称, "~^||") <> 0 Then
            str入院诊断编码 = Split(str入院诊断名称, "~^||")(0)
            str入院诊断名称 = Split(str入院诊断名称, "~^||")(1)
        Else
            str入院诊断编码 = ""
            str入院诊断名称 = ""
        End If
        If InStr(1, str确诊诊断名称, "~^||") <> 0 Then
            str确诊诊断编码 = Split(str确诊诊断名称, "~^||")(0)
            str确诊诊断名称 = Split(str确诊诊断名称, "~^||")(1)
        Else
            str确诊诊断编码 = ""
            str确诊诊断名称 = ""
        End If
    End If
    If lng中心 = 2 Then
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
    Else
        strInfor = strInfor & Lpad(str入院诊断编码, 16)  '入院诊断编码    CHAR    52  16      Y   院端
        strInfor = strInfor & Lpad(str入院诊断名称, 30)  '入院诊断名称    CHAR    68  30      y 院端
        strInfor = strInfor & Lpad(str确诊诊断编码, 16)  '确诊诊断编码    CHAR    98  16      N   院端
        strInfor = strInfor & Lpad(str确诊诊断名称, 30)  '确诊诊断名称    CHAR    114 30      N   院端
    End If
    strInfor = strInfor & Lpad(str入院科室, 20)  '科别名称    CHAR    144 20  如：内科    Y   院端
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & str床位号              '床位号  CHAR    164 10  如：2003室12床  N   院端
    End If
    strInfor = strInfor & Lpad(str转诊单号, 6)   '转诊单号    CHAR    174 6       N   院端
    strInfor = strInfor & Space(8)   '出院时间    DATE    180 8   系统利用患者结算数据的出院时间自动生成，医院端用空格补位即可。  N   无
    If lng中心 = 2 Then
    Else
        strInfor = strInfor & "M"   '传输标志    CHAR    188 1   A 入院登记，M 修改在院状态，C取消入院登记   Y   院端
        strInfor = strInfor & Space(16)   '传输时间    DATATIME    189 16  精确到秒格式为：yyyymmddhhmiss后面以空格补位，用于记录数据到达医保中心的时间，院端空格补位  N   中心
    End If
    '1004    9   206 实时住院登记数据提交
    在院病人信息_大连 = 业务请求_大连(lng中心, 1004, strInfor)
    If 在院病人信息_大连 = False Then
        ShowMsgbox "实时住院登记数据提交失败!"
        Exit Function
    End If
    在院病人信息_大连 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function



Public Function GetItemInfo_大连(ByVal lngPatiID As Long, ByVal lngItemID As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取大连病人的相关提示信息
    '--入参数:
    '--出参数:
    '--返  回:提示串
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim str医疗付款方式 As String
    Dim int险类 As Integer
    Dim bln在院 As Boolean
    Dim dbl统筹比例 As Double
    Dim strMsgInfor As String
    
    '第一步:确定是否医保病人
    gstrSQL = "Select 病人id,险类,nvl(当前状态,0) as 状态 from 保险帐户  where 病人id=" & lngPatiID & " and 险类=" & gintInsure
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "判断是否为医保病人!"
    If rsTemp.EOF Then
        rsTemp.Close
        GetItemInfo_大连 = ""
        Exit Function
    End If
    
    int险类 = NVL(rsTemp!险类, 0)
    bln在院 = NVL(rsTemp!状态, 0) > 0
    '第二步:确定医疗付款方式
    gstrSQL = "Select 医疗付款方式 from 病人信息 where 病人id=" & lngPatiID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取医疗付款式"
    str医疗付款方式 = NVL(rsTemp!医疗付款方式)
        
    '第三步：确定收费细目的相关数据
    gstrSQL = "" & _
        "   Select b.编码,b.名称,b.性质,b.算法,a.项目名称,b.统筹比额,b.特准定额,b.住院比额,a.是否医保 " & _
        "   From 保险支付项目 a,保险支付大类 b " & _
        "   where a.大类id=b.id and a.险类=b.险类 and a.收费细目id=" & lngItemID & " and a.险类=" & int险类
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取保险支付比例"
    strMsgInfor = ""
    If InStr(1, "社会基本医疗保险;企业离休;工伤保险;生育保险;商业保险;", IIf(str医疗付款方式 = "", "D", str医疗付款方式) & ";") <> 0 Then
        '   医疗付款方式为大连市医保、外地医保、企业离休、工伤保险、生育保险、商业保险的，报销比例按照大连市医保接口的医保项目管理中的医保大类定义中的报销比例进行提示
        '   医疗付款方式为开发区医保的，报销比例按照开发区医保接口的医保大类定义中的报销比例进行提示
        If bln在院 Then
            If NVL(rsTemp!算法, 0) = 2 Then
                 '刘兴宏:200404,采用的算法2(采用定额处理),不需进行比例计算
                 strMsgInfor = "该项目固定报销:" & Format(rsTemp!特准定额, "#####0.00;-####0.00; ;") & "元"
            Else
                 strMsgInfor = "该项目报销比例:" & Format(rsTemp!住院比额, "#####0.00;-####0.00; ;") & "%"
            End If
        Else
                 strMsgInfor = "该项目报销比例:" & Format(rsTemp!统筹比额, "#####0.00;-####0.00; ;") & "%"
        End If
    ElseIf InStr(1, "公费医疗;合同单位", "") <> 0 Then
        '   医疗付款方式为公费医疗、合同单位的，报销比例按照大连市医保接口的医保项目管理中的事业公费比例定义进行提示。
        strMsgInfor = "该项目公费比例:" & Format(Val(NVL(rsTemp!项目名称)), "#####0.00;-####0.00; ;") & "%"
    End If
    If strMsgInfor <> "" Then
        ShowMsgbox strMsgInfor
    End If
    GetItemInfo_大连 = strMsgInfor
End Function
