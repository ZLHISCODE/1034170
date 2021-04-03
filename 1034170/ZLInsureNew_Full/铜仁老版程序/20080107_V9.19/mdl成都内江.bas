Attribute VB_Name = "mdl成都内江"
Option Explicit

Public Enum 业务类型_成都内江
    读病人信息_内江 = 0
    更改密码_内江
    获取帐户余额_内江
    门诊明细写入_内江
    门诊消费确认_内江
    门诊消费取消_内江
    住院登记_内江
    住院交易上传_内江
    住院产易上传取消_内江
    出院登记上传_内江
    获取单位欠缴情况_内江
    初始化函数_内江
End Enum

Private gInitCard As Boolean                '初始化了卡的
Private Type InitbaseInfor
    医院编码 As String                      '初始医院编码
    串号号_内江 As Integer
    读卡器_内江 As Integer                  '0-明化,1-德森公司
    
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    解析卡内数据 As Boolean
End Type
Public InitInfor_成都内江 As InitbaseInfor
Private mblnStartTran   As Boolean '启动了事务的
Private Type 病人身份
        卡号       As String
        个人编号   As String
        身份证号   As String
        姓名       As String
        性别       As String
        工况类别   As String
        出生日期   As String
        单位号码   As String
        统筹编号   As String
        制卡日期   As String
        卡有效期   As String
        补卡次数   As String
        制卡单位   As String
        年龄        As Integer
        帐户余额    As Double
        在职情况    As String
        
        住院流水号 As String
        交易类别 As String
        lng病人ID   As Long
        
        费用总额  As Double
        结算方式    As String   '结算方式串
End Type

Private Type 结算数据
    待遇标志 As String
    医保交易流水号 As String
    医保内费用   As Double
    医保外费用   As Double
    基本医保支付    As Double
    高额医保支付    As Double
    公务员医疗补助  As Double
    帐户可用余额  As Double
    帐户支付        As Double
    比例支付        As Double
    起付标准        As Double
    结算标志        As Byte '0-门诊,1-住院
    结帐ID          As Long
    
End Type

Public g病人身份_成都内江 As 病人身份
Public gcnOracle_成都内江 As ADODB.Connection     '中间库连接
Private g结算数据   As 结算数据


'****************************************************************************************************************************************************************************************************************************************
'1 相关读卡组件函数
'****************************************************************************************************************************************************************************************************************************************
'   0-读病人信息函数(明华的)
Private Declare Function GetCardInfo_MW Lib "Mwic_32.dll" Alias "GetCardInfo" (ByVal lngPort As Long, ByVal strPassWord As String, ByVal str卡号 As String, _
        ByVal str个人编号 As String, ByVal str身份证号 As String, ByVal str姓名 As String, ByVal str性别 As String, _
        ByVal str工况类别 As String, ByVal str出生日期 As String, ByVal str单位号码 As String, ByVal str统筹编号 As String, _
        ByVal str制卡日期 As String, ByVal str卡有效期 As String, ByVal str补卡次数 As String, ByVal str制卡单位 As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'明华:
'函数原型:function GetCardInfo(port: integer;UserPassword:PChar; var CardNum,PersonNum,
'                   IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  UserPassword：输入参数，为用户密码，要求长度为6，字符串中只能包含0到9的数字；
'       c)  CardNum：输出参数，为卡号，长度为10；
'       d)  PersonNum：输出参数，为个人编号（医保编号），长度为8；
'       e)  IDNum：输出参数，为身份证号码，长度为18；
'       f)  Name：输出参数，为姓名，长度为20；
'       g)  Sex：输出参数，为性别编码，长度为1，其中'1'为男，'2'为女；
'       h)  PersonKind：输出参数，为工况类别，长度为1；
'       i)  Birthday：输出参数，为出生日期，长度为8，例如1982年6月23日表示为'19820623'；
'       j)  DeptNum：输出参数，为单位号码，长度为6；
'       k)  Zone：输出参数，为统筹地区编码，长度为1；
'       l)  MAKEDATE：输出参数，为制卡日期，长度为8，表示方式同出生日期；
'       m)  EXPIREDATE：输出参数，为卡有效日期（卡的有效期为99年，如制卡日期为20021101，则卡有效日期为21011101），长度为8，表示方式同出生日期；
'       n)  REISSUE：输出参数，为补卡次数，长度为2，例如：首次制卡，补卡次数为'00'，第一次补卡，补卡次数为'01'，以次类推；
'       o)  MAKEDEPT：输出参数，为制卡单位，长度为1，例如：'0'表示制卡方为温州德森公司。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   1-读病人信息函数(科瑞奇)
Private Declare Function GetCardInfo_KRQ Lib "Mwic_32.dll" Alias "GetCardInfo" (ByVal lngPort As Long, ByVal str卡号 As String, _
        ByVal str个人编号 As String, ByVal str身份证号 As String, ByVal str姓名 As String, ByVal str性别 As String, _
        ByVal str工况类别 As String, ByVal str出生日期 As String, ByVal str单位号码 As String, ByVal str统筹编号 As String, _
        ByVal str制卡日期 As String, ByVal str卡有效期 As String, ByVal str补卡次数 As String, ByVal str制卡单位 As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'科瑞奇
'说明:  与明细相比,没有密码的输入
'函数原型:function GetCardInfoForKRQ(port: integer; var CardNum,PersonNum,IDNum,Name,Sex,PersonKind,Birthday,DeptNum,Zone,
'           MAKEDATE,EXPIREDATE,REISSUE,MAKEDEPT: PChar):integer;
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  CardNum：输出参数，为卡号，长度为10；
'       c)  PersonNum：输出参数，为个人编号（医保编号），长度为8；
'       d)  IDNum：输出参数，为身份证号码，长度为18；
'       e)  Name：输出参数，为姓名，长度为20；
'       f)  Sex：输出参数，为性别编码，长度为1，其中'1'为男，'2'为女；
'       g)  PersonKind：输出参数，为工况类别，长度为1；
'       h)  Birthday：输出参数，为出生日期，长度为8，例如1982年6月23日表示为'19820623'；
'       i)  DeptNum：输出参数，为单位号码，长度为6；
'       j)  Zone：输出参数，为统筹地区编码，长度为1；
'       k)  MAKEDATE：输出参数，为制卡日期，长度为8，表示方式同出生日期；
'       l)  EXPIREDATE：输出参数，为卡有效日期（卡的有效期为99年，如制卡日期为20021101，则卡有效日期为21011101），长度为8，表示方式同出生日期；
'       m)  REISSUE：输出参数，为补卡次数，长度为2，例如：首次制卡，补卡次数为'00'，第一次补卡，补卡次数为'01'，以次类推；
'       n)  MAKEDEPT：输出参数，为制卡单位，长度为1，例如：'0'表示制卡方为温州德森公司。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'   2-修改密码
Private Declare Function ChangePassword Lib "Mwic_32.dll" (ByVal lngPort As Long, ByVal strOldPassWord As String, ByVal strNewPassWord As String) As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'说明:  与明细相比,没有密码的输入
'函数原型:function ChangePassword(port:integer;OldPassword,NewPassword:PChar):integer;
'参数:  a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
'       b)  OldPassword：输入参数，为原密码，要求长度为6，字符串中只能包含0到9的数字；
'       c)  NewPassword：输入参数，为新密码，要求长度为6，字符串中只能包含0到9的数字。
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'****************************************************************************************************************************************************************************************************************************************
'2 业务对象
'****************************************************************************************************************************************************************************************************************************************
Public gobj成都内江 As Object


Public Function 医保初始化_成都内江() As Boolean
    
    
    Dim strReg As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    
    GetRegInFor g公共全局, "医保", "读卡器", strReg
    InitInfor_成都内江.读卡器_内江 = Val(strReg)
    
    
    GetRegInFor g公共全局, "医保", "串口号", strReg
    InitInfor_成都内江.串号号_内江 = IIf(strReg = "", 1, Val(strReg))
        
        
        
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_成都内江.模拟数据 = True
    Else
        InitInfor_成都内江.模拟数据 = False
    End If
    
    Call GetRegInFor(g公共模块, "操作", "解析卡内数据", strReg)
    If Val(strReg) = 1 Then
        InitInfor_成都内江.解析卡内数据 = True
    Else
        InitInfor_成都内江.解析卡内数据 = False
    End If
    InitInfor_成都内江.解析卡内数据 = InitInfor_成都内江.解析卡内数据 Or InitInfor_成都内江.模拟数据
    
    
    '创建医保对象
    If gobj成都内江 Is Nothing Then
        Err = 0
        On Error Resume Next
        Set gobj成都内江 = CreateObject("SocketOcxForNC.SocketOcxForNC")
        If Err <> 0 Then
                ShowMsgbox "不能创建医保接口,请检查SocketOcxForNC.ocx是否正常注册!"
                Exit Function
        End If
    End If
    
    
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_成都内江
    Call OpenRecordset(rsTemp, "读取医院编码")
    InitInfor_成都内江.医院编码 = Nvl(rsTemp!医院编码)
    If Open中间库 = False Then Exit Function
    
    医保初始化_成都内江 = True
End Function
Private Function Open中间库() As Boolean
    '连接中间库
    '中间库连接
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strServer As String, strPass As String, strReg As String
    Dim strInPut As String, strOutPut As String
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=" & TYPE_成都内江
    Call OpenRecordset(rsTemp, "获取相关参数值")
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    Set gcnOracle_成都内江 = New ADODB.Connection

    If OraDataOpen(gcnOracle_成都内江, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查网络是否畅通无阻
          
    GetRegInFor g公共全局, "医保", "ConfigFileName", strReg
    strInPut = strReg
    GetRegInFor g公共全局, "医保", "HostPort", strReg
    strInPut = strInPut & vbTab & strReg
    GetRegInFor g公共全局, "医保", "IPAddress", strReg
    strInPut = strInPut & vbTab & strReg
    
    If 业务请求_成都内江(初始化函数_内江, strInPut, strOutPut) = False Then Exit Function
    
    Open中间库 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 医保终止_成都内江() As Boolean
    '结束读写卡组件
    Dim strReg As String
    
    Err = 0
    On Error Resume Next
    
    Set gobj成都内江 = Nothing
    If gcnOracle_成都内江.State = 1 Then
        gcnOracle_成都内江.Close
    End If
    医保终止_成都内江 = True
End Function

Public Function 身份标识_成都内江(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo ErrHand:
    身份标识_成都内江 = frmIdentify成都内江.GetPatient(bytType, lng病人ID)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_成都内江 = ""
End Function

Public Function 个人余额_成都内江(ByVal lng病人ID As Long) As Currency
    '功能: 提取参保病人个人帐户余额
    '返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_成都内江
    Call OpenRecordset(rsTemp, "读取个人帐户余额")
    
    If rsTemp.EOF Then
        个人余额_成都内江 = 0
    Else
        个人余额_成都内江 = rsTemp("帐户余额")
    End If
End Function

Public Function 门诊虚拟结算_成都内江(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
  
    门诊虚拟结算_成都内江 = False
    Exit Function
End Function
Private Function 获取个人帐户支付() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取个人帐户值(从预交记录中获取)
    '--入参数:
    '--出参数:
    '--返  回:成功,返回本次个人帐户支付,否则返回0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 金额 From 病人预交记录 where 结帐ID=" & g结算数据.结帐ID & " and  结算方式='个人帐户'"
    
    OpenRecordset rsTemp, "获取个人帐户支付"
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!金额, 0)
    End If
    
End Function


Public Function 门诊结算_成都内江(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim strInPut As String, strOutPut As String
    Dim strArr
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim lng病人ID  As Long
    Dim str操作员编码 As String
    Err = 0: On Error GoTo ErrHand:
    
    
    Call DebugTool("进入门诊结算")
    gstrSQL = "Select 收费细目ID From 病人费用记录  where 结帐id=" & lng结帐ID & " group by 收费细目iD   having Count(收费细目id)>=2 "
    Call OpenRecordset(rsTemp, "判断明细是否重复")
    
    If Not rsTemp.EOF Then
        MsgBox "存在重复的收费细目,请合并后再结算!"
        Exit Function
    End If
    
    gstrSQL = "Select a.*,a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 From 病人费用记录 a Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call OpenRecordset(rs明细, "获取明细记录")
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If

    lng病人ID = rs明细("病人ID")
    str操作员编码 = Nvl(rs明细!操作员姓名)
    If g病人身份_成都内江.lng病人ID <> lng病人ID Then
        MsgBox "该病人还没有经过身份 验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    g结算数据.结算标志 = 0
    g结算数据.结帐ID = lng结帐ID
    
    
    '写入明细
    If 门诊明细写入(rs明细, False) = False Then Exit Function
    If 结算方式更正 = False Then Exit Function
    
    '获取帐户支付
    g结算数据.帐户支付 = 获取个人帐户支付()
    
    
    '消费进行确认
    '输入参数: 个人编号    String(8)   In
    '          社保卡号码  String(10)  In
    '          医院代码    String(5)   In
    '          操作员卡号码    String(10)  In
    '          统筹地区编码    String(1)   In
    '          医保交易流水号  String(20)  In
    '          交易类别    String(1)   In
    '          个人帐户支付    String(10)  In
    With g病人身份_成都内江
        strInPut = Rpad(.个人编号, 8)
        strInPut = strInPut & vbTab & Rpad(.卡号, 10)
        strInPut = strInPut & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        strInPut = strInPut & vbTab & Substr(Rpad(str操作员编码, 10), 1, 10)
        strInPut = strInPut & vbTab & Rpad(.统筹编号, 1)
        strInPut = strInPut & vbTab & Substr(Rpad(g结算数据.医保交易流水号, 20), 1, 20)
        strInPut = strInPut & vbTab & Rpad(.交易类别, 1)
        strInPut = strInPut & vbTab & Rpad(g结算数据.帐户支付, 1)
    End With
    
    '调用结算
    Call DebugTool("准备调用门诊消费确认")
    If 业务请求_成都内江(门诊消费确认_内江, strInPut, strOutPut) Then Exit Function
    Call DebugTool("调用门诊消费确认结束")
    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(无),住院次数_IN(无),起付线(无),封顶线_IN(帐户可用余额),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0,0,0,0," & g结算数据.帐户可用余额 & ",0," & _
            g病人身份_成都内江.费用总额 & "," & g结算数据.医保内费用 & "," & g结算数据.医保外费用 & "," & _
           "0,0,0,0," & g结算数据.帐户支付 & ",'" & _
            g结算数据.医保交易流水号 & "',NULL,NULL,NULL)"
            
    Call ExecuteProcedure("保存结算记录")
    '---------------------------------------------------------------------------------------------
    门诊结算_成都内江 = True
    Exit Function
ErrHand::
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Get交易流水号() As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取交易流水号
    '--入参数:
    '--出参数:
    '--返  回:交易流水号
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 医院交易流水号_ID.nextval as 序列 From dual"
    OpenRecordset_成都内江 rsTemp, "获取交易流水号"
    Get交易流水号 = InitInfor_成都内江.医院编码 & Format(zlDatabase.Currentdate, "yyyyMMDD") & Lpad(Nvl(rsTemp!序列), 7, "0")
End Function

Private Function 门诊明细写入(ByVal rs明细 As ADODB.Recordset, Optional ByVal bln虚拟 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传门诊明细数据
    '--入参数:rs明细-明细记录
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    
    Dim strInPut As String, strOutPut As String, str明细 As String
    Dim strInsert As String
    Dim lngSumLen As Long
    
    Dim str交易流水号 As String
    Dim lng处方条数 As Long
    
    Dim strArr
    
    门诊明细写入 = False
    
    DebugTool "进入门诊明细上传接口"
    
    g病人身份_成都内江.费用总额 = 0
        

    
    Err = 0
    On Error GoTo ErrHand:
    str明细 = ""
    '然后插入处方明细
    str交易流水号 = Get交易流水号
    With rs明细
        lng处方条数 = 0
        Do While Not .EOF
            
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                '处方明细
                '1   处方项目种类 Varchar2(1)    '1'：药品编码   '2'：服务项目
                '2   处方项目代码    Varchar2(20)    "药品编码"或者"诊疗项目编码"
                '3   数量    Varchar2(10)    实际数量*100上传
                '4   规格    Varchar2(10)    汉字算2字节
                '5   单项费用    Varchar2(10)    由医院上传(主要传什么?)
                
                
                
                gstrSQL = "select A.名称,A.编码,A.类别,A.规格,A.计算单位,B.项目编码,B.附注,B.是否医保,A.计算单位,E.规格,G.名称 剂型,B.大类编码 " & _
                          "from 收费细目 A," & _
                          "         (   Select a.*,b.大类编码 " & _
                          "             From 保险支付项目 a,保险项目 b" & _
                          "             where a.险类=b.险类 and a.项目编码=b.编码 and A.收费细目ID =" & rs明细!收费细目ID & " and a.险类=" & gintInsure & ") B,药品目录 E ,药品信息 F,药品剂型 G " & _
                          "where A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID(+) " & _
                         "        AND A.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) "
                OpenRecordset rsTemp, "获取医保项目"
                
                If rsTemp.EOF Then
                    ShowMsgbox "存在未对码的项目,请在保险项目管理中进行对码!"
                    Exit Function
                End If
                
                
                str明细 = str明细 & Substr(Rpad(Nvl(rsTemp!大类编码), 1), 1, 1)
                str明细 = str明细 & Rpad(Nvl(rsTemp!项目编码), 20)
                str明细 = str明细 & Lpad(Nvl(!数量) * 100, 10, "0")
                str明细 = str明细 & Rpad(Nvl(rsTemp!规格), 10)
                str明细 = str明细 & Lpad(Nvl(!实收金额) * 100, 10, "0")
                
                '为病人费用记录打上标记，以便随时上传
                'ID_IN,统筹金额_IN,保险大类ID_IN,保险项目否_IN,保险编码_IN,是否上传_IN,摘要_IN
                '摘要值:医院交易流水号
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str交易流水号 & "')"
                ExecuteProcedure "打上上传标志"
                lng处方条数 = lng处方条数 + 1
            End If
            g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rs明细!实收金额, 0)
            rs明细.MoveNext
        Loop
        
        If lng处方条数 > 99 Then
            ShowMsgbox "门诊处方明细不能大于99种项目,请分成两张处方进行录入!"
            Exit Function
        End If
        
        If .RecordCount <> 0 Then
            .MoveFirst
            '输入参数：个人编号    String(8)   In
                          '       社保卡号码  String(10)  In
                          '       医院代码    String(5)   In
                          '       操作员卡号码    String(10)  In
                          '       统筹地区编码    String(1)   In
                          '       医院交易流水号  String(20)  In
                          '       交易类别    String(1)   In
                          '       处方条数    String(2)   In
                          '       处方明细    String处方条数×51  In
            strInPut = Rpad(g病人身份_成都内江.个人编号, 8)
            strInPut = strInPut & vbTab & Rpad(g病人身份_成都内江.卡号, 10)
            strInPut = strInPut & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
            strInPut = strInPut & vbTab & Rpad(Nvl(!操作员编号), 10)
            strInPut = strInPut & vbTab & Rpad(g病人身份_成都内江.统筹编号, 1)
            strInPut = strInPut & vbTab & Rpad(str交易流水号, 20)
            strInPut = strInPut & vbTab & Rpad(g病人身份_成都内江.交易类别, 1)
            strInPut = strInPut & vbTab & Rpad(lng处方条数, 2)
            strInPut = strInPut & vbTab & str明细
            If 业务请求_成都内江(门诊明细写入_内江, strInPut, strOutPut) = False Then Exit Function
            
            '保存相关数据
            '    医院流水号_IN IN 医保消费信息.医院流水号%TYPE,
            '    病人ID_IN IN 医保消费信息.病人ID%TYPE,
            '    医保流水号_IN IN 医保消费信息.医保流水号%TYPE,
            '    医保内费用_IN IN 医保消费信息.医保内费用%TYPE,
            '    医保外费用_IN IN 医保消费信息.医保外费用%TYPE,
            '    帐户可用余额_IN IN 医保消费信息.帐户可用余额%TYPE,
            '    在职情况_IN IN 医保消费信息.在职情况%TYPE,
            '    医保项目种类_IN IN 医保消费信息.医保项目种类%TYPE,
            '    医保项目编码_IN IN 医保消费信息.医保项目编码%TYPE,
            '    医保内费用1_IN IN 医保消费信息.医保内费用1%TYPE,
            '    费用类别_IN IN 医保消费信息.费用类别%TYPE,
            '    项目费用_IN IN 医保消费信息.项目费用%TYPE
            strArr = Split(strOutPut, vbTab)
            
            With g结算数据
                .医保交易流水号 = strArr(0)
                .医保内费用 = Val(strArr(1))
                .医保外费用 = Val(strArr(2))
                .帐户可用余额 = Val(strArr(3))
            End With
            strInsert = "ZL_医保消费信息_INSERT("
            strInsert = strInsert & "'" & str交易流水号 & "',"
            strInsert = strInsert & "" & g病人身份_成都内江.lng病人ID & ","
            strInsert = strInsert & "'" & strArr(0) & "',"
            strInsert = strInsert & "" & Val(strArr(1)) & ","
            strInsert = strInsert & "" & Val(strArr(2)) & ","
            strInsert = strInsert & "" & Val(strArr(3)) & ","
            strInsert = strInsert & "'" & strArr(5) & "',"
            
            
            '分解明细记录记录
                        
            '1   处方项目种类 Varchar2(1)    '1'：药品编码   '2'：服务项目
            '2   处方项目代码    Varchar2(20)    "药品编码"或者"诊疗项目编码"
            '3   医保内费用  Varchar2(10)    实际数量*100
            '4   费用类别 Varchar2(10)(中药?西药?化验等)
            '5   项目费用    Varchar2(10)    实际数量*100
            str明细 = strArr(4)
            lngSumLen = zlCommFun.ActualLen(str明细)
            strInPut = ""
            Dim r As Long, i As Integer
            
            For i = 1 To lngSumLen Step 51
                r = 1
                strInPut = strInPut & "'" & Substr(str明细, r, 1) & "',"
                r = r + 1
                strInPut = strInPut & "'" & Substr(str明细, r, 20) & "',"
                r = r + 20
                strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ","
                r = r + 10
                strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ","
                r = r + 10
                strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ")"
                
                '组合SQL误句
                gstrSQL = strInsert & strInPut
                ExecuteProcedure_ZLNJ "插入明细数据到中间库"
            Next
        End If
    End With
    门诊明细写入 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_成都内江(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut  As String, str流水号 As String
    Dim lng冲销ID As Long, lng病人id1 As Long
    Dim strArr
    Dim rs明细 As New ADODB.Recordset
    Dim i As Long
    Dim intMouse  As Integer
 
    On Error GoTo ErrHand:

    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    '身份验证
    If 身份标识_成都内江(0, lng病人id1) = "" Then
        Screen.MousePointer = intMouse
        门诊结算冲销_成都内江 = False
        Exit Function
    End If
    Screen.MousePointer = intMouse
    
    
    gstrSQL = "Select * From 病人费用记录  " & _
        " Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call OpenRecordset(rs明细, "获取冲销记录")
    
    
    
    g病人身份_成都内江.费用总额 = 0
    Do Until rs明细.EOF
        If lng病人ID = 0 Then lng病人ID = rsTemp("病人ID")
        str流水号 = Nvl(rs明细!摘要)
        g病人身份_成都内江.费用总额 = g病人身份_成都内江.费用总额 + Nvl(rs明细("结帐金额"), 0)
        rs明细.MoveNext
    Loop
    g病人身份_成都内江.费用总额 = Round(g病人身份_成都内江.费用总额, 2)
    
    If lng病人ID <> lng病人id1 Then
        ShowMsgbox " 验卡病人不是当前要冲销的病人,不能冲销结算"
        Exit Function
    End If
    
    
    '退费
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "重庆医保")
    lng冲销ID = rsTemp("结帐ID")

    

    gstrSQL = "Select * From 病人费用记录 " & _
        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
    Call OpenRecordset(rsTemp, "获取冲销记录")
    
    Do While Not rsTemp.EOF
        '更新上传标志
        gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(rsTemp!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str流水号 & "')"
        ExecuteProcedure "打上上传标志"
        rsTemp.MoveNext
    Loop

    gstrSQL = "select * from 保险结算记录 where 性质=1 and 险类=" & gintInsure & " and 记录ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "获取原来的结算记录")

    If rsTemp.EOF = True Then
        MsgBox "原单据的医保记录不存在，不能作废。", vbInformation, gstrSysName
        Exit Function
    End If
    
    g结算数据.医保交易流水号 = rsTemp("支付顺序号")
    
    '    个人编号    String(8)   In
    '    社保卡号码  String(10)  In
    '    医院代码    String(5)   In
    '    操作员卡号码    String(10)  In
    '    统筹地区编码    String(1)   In
    '    医保交易流水号  String(20)  In
    '    交易类别    String(1)   In
    With g病人身份_成都内江
        strInPut = Rpad(.个人编号, 8)
        strInPut = strInPut & vbTab & Rpad(.卡号, 10)
        strInPut = strInPut & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        strInPut = strInPut & vbTab & Rpad(gstrUserName, 10)
        strInPut = strInPut & vbTab & Rpad(.统筹编号, 1)
        strInPut = strInPut & vbTab & Rpad(g结算数据.医保交易流水号, 20)
        strInPut = strInPut & vbTab & Rpad(.交易类别, 1)
    End With
    If 业务请求_成都内江(门诊消费取消_内江, strInPut, strOutPut) = False Then Exit Function
    
 '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN(帐户可用余额),实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(无),住院次数_IN(无),起付线(无),封顶线_IN(帐户可用余额),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(无),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
        "0,0,0,0,0,0,0," & -1 * Nvl(rsTemp!封顶线, 0) & "," & _
        Nvl(rsTemp!发生费用金额, 0) * -1 & "," & Nvl(rsTemp!全自付金额, 0) * -1 & "," & Nvl(rsTemp!首先自付金额, 0) * -1 & "," & _
        rsTemp("进入统筹金额") * -1 & "," & rsTemp("统筹报销金额") * -1 & ",0,0," & rsTemp("个人帐户支付") * -1 & ",'" & _
       g结算数据.医保交易流水号 & "',NULL,0,null)"
    Call ExecuteProcedure("更新保险结算信息")
    门诊结算冲销_成都内江 = True
    Exit Function
ErrHand::
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function 病人入院登记处理(lng病人ID As Long, lng主页id As Long) As Boolean
    '进行门诊登记
    Dim strInPut As String, strOutPut As String
    Dim str交易流水号 As String
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "Select C.住院号,C.当前床号,to_char(A.确诊日期,'yyyy-MM-dd') as 确诊日期,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.入院日期,'yyyymmdd') 入院日期  ,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') 入院时间,D.入院诊断编码,D.入院诊断名称,G.确诊诊断编码,g.确诊诊断名称 " & _
        " From 病案主页 A,部门表 B,病人信息 C, " & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,1,b.编码,'')) AS 入院诊断编码,max(DECODE(a.诊断次序,1,b.名称,'')) AS 入院诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页id & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   D," & _
        "       (Select 病人id,主页id,max(DECODE(a.诊断次序,2,b.编码,'')) AS 确诊诊断编码,max(DECODE(a.诊断次序,2,b.名称,'')) AS 确诊诊断名称 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =1 and a.主页id=" & lng主页id & " and a.病人id=" & lng病人ID & " Group by  病人id,主页id)   g" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页id & " And A.入院科室ID=B.ID " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        "       and A.主页id=g.主页id(+) and a.病人id=g.病人id(+) " & _
        ""

    OpenRecordset rsTemp, gstrSQL, "读取入院信息"

    With g病人身份_成都内江
        '输入参数
        '    个人编号    String(8)   In
        '    社保卡号码  String(10)  In
        '    医院代码    String(5)   In
        '    操作员卡号码    String(10)  In
        '    统筹地区编码    String(1)   In
        '    入院日期    String(8)   In
        '    入院科别    String(10)  In
        '    入院诊治医生    String(10)  In
        '    诊断编码    String(20)  In
        strInPut = Rpad(.个人编号, 8)
        strInPut = strInPut & vbTab & Rpad(.卡号, 10)
        strInPut = strInPut & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        strInPut = strInPut & vbTab & Rpad(gstrUserName, 10)
        strInPut = strInPut & vbTab & Rpad(.统筹编号, 1)
        strInPut = strInPut & vbTab & Rpad(rsTemp!入院日期, 8)
        strInPut = strInPut & vbTab & Rpad(Substr(rsTemp!入院科室, 1, 10), 10)
        strInPut = strInPut & vbTab & Rpad(Substr(rsTemp!住院医师, 1, 10), 10)
        strInPut = strInPut & vbTab & Rpad(Substr(rsTemp!入院诊断编码, 1, 20), 20)
                
        If 业务请求_成都内江(住院登记_内江, strInPut, strOutPut) = False Then
            Exit Function
        End If
        
        '输出参数
        '    住院流水号  String(20)  Out
        '    享受待遇标志    Small int   Out
        '    起付标准    Long    Out
        '    在职情况    String(1)   Out
        
        strArr = Split(strOutPut, vbTab)

        '保存将交易流水号
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & gintInsure & ",'顺序号','''" & strArr(0) & "''')"
        Call ExecuteProcedure("保存交易流水号")
        '保存享受待遇标志
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & gintInsure & ",'享受待遇标志','''" & Val(strArr(1)) & "''')"
        Call ExecuteProcedure("保存享受待遇标志")
        '保存起付标准
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & gintInsure & ",'起付标准','''" & Val(strArr(2)) & "''')"
        Call ExecuteProcedure("保存起付标准")
    End With

    病人入院登记处理 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_成都内江(lng病人ID As Long, lng主页id As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset, rsData As New ADODB.Recordset
    Dim strOutPut As String, strInPut As String
    
    '获取住院号
    Err = 0
    On Error GoTo ErrHand:
 
    '先进行登记处理
    If 病人入院登记处理(lng病人ID, lng主页id) = False Then
        Exit Function
    End If

    '将病人的状态进行修改
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    入院登记_成都内江 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_成都内江 = False
End Function

Public Function 入院登记撤销_成都内江(lng病人ID As Long, lng主页id As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false

    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    Dim str医保号  As String
    Dim str出院日期 As String

    Err = 0
    On Error GoTo ErrHand
    ShowMsgbox "该医保接口不支持入院登记撤消,只能"
    Exit Function
   入院登记撤销_成都内江 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 出院登记_成都内江(lng病人ID As Long, lng主页id As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '个人状态的修改
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure(gstrSysName)
    出院登记_成都内江 = True
    Exit Function
ErrHand::
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_成都内江 = False
End Function
Public Function 出院登记撤销_成都内江(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
    '出院登记撤消
     '改变病人状态
     If Not 存在未结费用(lng病人ID, lng主页id) Then
            ShowMsgbox "该病人已经出院结算了,不能出院登记撤消!"
            Exit Function
     End If
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & TYPE_成都内江 & ")"
    Call ExecuteProcedure("办理入院登记")
    出院登记撤销_成都内江 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function 住院结算_成都内江(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)

    Dim rsTemp As New ADODB.Recordset, strInPut As String, strOutPut As String

    Dim str操作员 As String
    Dim lng主页id As Long
    Dim strArr
    
    Dim i As Integer

    If g病人身份_成都内江.lng病人ID <> lng病人ID Then
        MsgBox "该病人没有完成医保的预结算操作，不能进行结算。", vbInformation, gstrSysName
        Exit Function
    End If


    Err = 0: On Error GoTo ErrHand:
    Call DebugTool("进入住院结算")


    With g结算数据
        gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & lng病人ID
        Call OpenRecordset(rsTemp, "虚拟结算")
        If IsNull(rsTemp("主页ID")) = True Then
            MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
            Exit Function
        End If
        lng主页id = rsTemp("主页ID")
    End With
    
   gstrSQL = "Select A.ID From 病人费用记录 a,药品收发记录 B where A.no=b.No and B.单据 in (9,10) and a.id=b.费用ID and a.结帐ID=" & lng结帐ID & " and b.扣率 like '_3%' and rownum<=2"

    
    Dim bln出院带药 As Boolean
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定出院带药"
    If rsTemp.EOF Then
        bln出院带药 = False
    Else
        bln出院带药 = True
    End If
    
  gstrSQL = "Select c.住院号,A.登记人 经办人,B.名称 入院科室,A.住院医师,to_char(A.登记时间,'yyyy-MM-dd hh24:mi:ss') 入院经办时间," & _
        " to_char(A.入院日期,'yyyyMMdd') 入院日期,J.终止时间,J.操作员,D.诊断编码,A.出院方式,to_Char(a.出院日期,'yyyyMMDD') as 出院日期,a.出院病床,H.名称 as 出院科室" & _
        " From 病案主页 A,部门表 B,病人信息 C,部门表 H, " & _
        "       (Select 病人id,主页id max(DECODE(a.诊断次序,2,b.编码,'')) AS 诊断编码 From 诊断情况 A ,疾病编码目录 B Where a.疾病ID = b.ID And a.诊断类型 =3  and a.主页id=" & lng主页id & " and a.病人id=" & lng病人ID & " Group by 病人id,主页id)   D" & _
        " Where A.病人id=C.病人id and C.病人id=" & lng病人ID & _
        "       and A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页id & " And A.入院科室ID=B.ID and A.出院科室ID=H.id(+) " & _
        "       and A.主页id=D.主页id(+) and a.病人id=D.病人id(+) " & _
        ""
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "确定入出类别"
    
    
    
    '入参:
    '    个人编号    String(8)   In
    '    社保卡号码  String(10)  In
    '    医院代码    String(5)   In
    '    操作员卡号码    String(10)  In
    '    统筹地区编码    String(1)   In
    '    出院日期    String(8)   In
    '    出院科别    String(10)  In
    '    出院诊治医生    String(10)  In
    '    诊断编码    String(20)  In
    '    出院带药    String(1)   In
    '    出院类别    String(1)   In
    '    住院流水号  String(20)  In
      
    With g病人身份_成都内江
        strInPut = Rpad(.个人编号, 8)
        strInPut = strInPut & vbTab & Rpad(.卡号, 10)
        strInPut = strInPut & vbTab & Rpad(InitInfor_成都内江.医院编码, 5)
        strInPut = strInPut & vbTab & Rpad(Nvl(rsTemp!操作员), 10)
        strInPut = strInPut & vbTab & Rpad(.统筹编号, 1)
        strInPut = strInPut & vbTab & Rpad(Nvl(rsTemp!出院日期), 8)
        strInPut = strInPut & vbTab & Substr(Rpad(Nvl(rsTemp!出院科室), 10), 1, 10)
        strInPut = strInPut & vbTab & Substr(Rpad(Nvl(rsTemp!住院医师), 10), 1, 10)
        strInPut = strInPut & vbTab & Substr(Rpad(Nvl(rsTemp!诊断编码), 20), 1, 20)
        strInPut = strInPut & vbTab & IIf(bln出院带药, "1", "0")
        strInPut = strInPut & vbTab & Substr(Nvl(rsTemp!出院方式), 1, 1)
        strInPut = strInPut & vbTab & Substr(Rpad(.住院流水号, 20), 1, 20)
        If 业务请求_成都内江(出院登记上传_内江, strInPut, strOutPut) = False Then Exit Function
    End With
    If strOutPut = "" Then Exit Function
    strArr = Split(strOutPut, vbTab)
    
   '出参
    '    TRANSDETIAL输出 (计算费用明细)
    '    享受待遇标志    String(1)   Out
    '    医保内费用  String(10)  Out
    '    医保外费用  String(10)  Out
    '    基本医保支付
    '    如果参加大病医保，则为大病医保支付  String(10)  Out
    '    高额医保支付    String(10)  Out
    '    公务员医疗补助  String(10)  Out
    '    个人按比例支付  String(10)  Out
    '    TRANSDETIAL结束
    '    起付标准    String(10)  Out
    '    个人帐户可用余额    String(10)  Out
    strOutPut = strArr(0)
    With g结算数据
        .结算标志 = 1
        .待遇标志 = Substr(strOutPut, 1, 1)
        .医保内费用 = Val(Substr(strOutPut, 2, 10))
        .医保外费用 = Val(Substr(strOutPut, 12, 10))
        .基本医保支付 = Val(Substr(strOutPut, 22, 10))
        .高额医保支付 = Val(Substr(strOutPut, 32, 10))
        .公务员医疗补助 = Val(Substr(strOutPut, 42, 10))
        .比例支付 = Val(Substr(strOutPut, 52, 10))
        .起付标准 = Val(strArr(1))
        .帐户可用余额 = Val(strArr(2))
        
    End With
    
    If 结算方式更正() = False Then Exit Function
        
     '获取帐户支付
    g结算数据.帐户支付 = 获取个人帐户支付()
    

  '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(无),帐户累计支出_IN(无),累计进入统筹_IN(无),累计统筹报销_IN(无),住院次数_IN(无),起付线(比例支付),封顶线_IN(帐户可用余额),实际起付线_IN(起付标准),
    '   发生费用金额_IN(费用总额),全自付金额_IN(医保内费用),首先自付金额_IN(医保外费用),
    '   进入统筹金额_IN(统筹支付),统筹报销金额_IN(统筹支付),大病自付金额_IN(高额医保支付),超限自付金额_IN(公务员医疗补助),个人帐户支付_IN(个人帐户支付),"
    '   支付顺序号_IN(结算时产生流水号),主页ID_IN,中途结帐_IN,备注_IN(享受待遇标志)
    
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
            "0,0,0,0,0," & g结算数据.比例支付 & "," & g结算数据.帐户可用余额 & "," & g结算数据.起付标准 & "," & _
            g病人身份_成都内江.费用总额 & "," & g结算数据.医保内费用 & "," & g结算数据.医保外费用 & "," & _
           g结算数据.基本医保支付 & "," & g结算数据.基本医保支付 & "," & g结算数据.高额医保支付 & "," & g结算数据.公务员医疗补助 & "," & g结算数据.帐户支付 & ",'" & _
            g病人身份_成都内江.住院流水号 & "',NULL,NULL,'" & g结算数据.待遇标志 & "')"
            

    Call ExecuteProcedure("保存结算记录")
    '---------------------------------------------------------------------------------------------

    住院结算_成都内江 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function 住院结算冲销_成都内江(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------

    Err = 0: On Error GoTo ErrHand:
    ShowMsgbox "本医保接不支持反结帐"
    住院结算冲销_成都内江 = False
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 保存明细结果到中间库(ByVal str交易流水号 As String, ByVal strOutPut As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存明细结果到中间库
    '--入参数:以vbtab分离
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim strInPut As String, str明细 As String
    Dim strArr
    Dim lngSumLen As Long
    Dim r As Long, i As Integer
    'strOutPut: 医保交易流水号  String(20)  Out
    '           处方明细    String处方条数×51  Out
    '           TRANSDETIAL输出 (计算费用明细) Out
    
    'TRANSDETIAL输出
    '        享受待遇标志    String(1)   Out
    '        医保内费用  String(10)  Out
    '        医保外费用  String(10)  Out
    '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
    '        高额医保支付    String(10)  Out
    '        公务员医疗补助  String(10)  Out
    '        个人按比例支付  String(10)  Out


    保存明细结果到中间库 = False
    Err = 0: On Error GoTo ErrHand:
    strArr = Split(strOutPut, vbTab)
    
    '过程参数
    '    医院流水号_IN IN 医保消费信息.医院流水号%TYPE,
    '    病人ID_IN IN 医保消费信息.病人ID%TYPE,
    '    医保流水号_IN IN 医保消费信息.医保流水号%TYPE,
    '    医保内费用_IN IN 医保消费信息.医保内费用%TYPE,
    '    医保外费用_IN IN 医保消费信息.医保外费用%TYPE,
    '    帐户可用余额_IN IN 医保消费信息.帐户可用余额%TYPE:=NULL,
    '    在职情况_IN IN 医保消费信息.在职情况%TYPE:=NULL,
    '    医保项目种类_IN IN 医保消费信息.医保项目种类%TYPE:=NULL,
    '    医保项目编码_IN IN 医保消费信息.医保项目编码%TYPE,
    '    医保内费用1_IN IN 医保消费信息.医保内费用1%TYPE,
    '    费用类别_IN IN 医保消费信息.费用类别%TYPE:=NULL,
    '    项目费用_IN IN 医保消费信息.项目费用%TYPE:=NULL,
    '    享受待遇标志_IN IN 医保消费信息.享受待遇标志%TYPE:=NULL,
    '    基本医保支付_IN IN 医保消费信息.基本医保支付%TYPE:=NULL,
    '    高额医保支付_IN IN 医保消费信息.高额医保支付%TYPE:=NULL,
    '    公务员补助_IN IN 医保消费信息.公务员补助%TYPE:=NULL,
    '    个人比例支付_IN IN 医保消费信息.个人比例支付%TYPE:=NULL
    strHead = "ZL_医保消费信息_INSERT("
    strHead = strHead & "'" & str交易流水号 & "',"
    strHead = strHead & "" & g病人身份_成都内江.lng病人ID & ","
    strHead = strHead & "'" & g结算数据.医保交易流水号 & "',"
    
    strHead = strHead & "" & Val(Substr(strArr(2), 2, 10)) & ","
    strHead = strHead & "" & Val(Substr(strArr(2), 12, 10)) & ","
    
    str明细 = strArr(1)
    lngSumLen = zlCommFun.ActualLen(strArr(2))
    
    For i = 1 To lngSumLen Step 51
        r = 1
        strInPut = strInPut & "'" & Substr(str明细, r, 1) & "',"
        r = r + 1
        strInPut = strInPut & "'" & Substr(str明细, r, 20) & "',"
        r = r + 20
        strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ","
        r = r + 10
        strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ","
        r = r + 10
        strInPut = strInPut & "" & Val(Substr(str明细, r, 10)) & ")"
        '加上
        'TRANSDETIAL输出
         '        享受待遇标志    String(1)   Out
         '        医保内费用  String(10)  Out
         '        医保外费用  String(10)  Out
         '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
         '        高额医保支付    String(10)  Out
         '        公务员医疗补助  String(10)  Out
         '        个人按比例支付  String(10)  Out


        strInPut = strInPut & "'" & Substr(strArr(2), 1, 1) & "',"
        strInPut = strInPut & "" & Val(Substr(strArr(2), 22, 1)) & ","
        strInPut = strInPut & "" & Val(Substr(strArr(2), 32, 1)) & ","
        strInPut = strInPut & "" & Val(Substr(strArr(2), 42, 1)) & ","
        strInPut = strInPut & "" & Val(Substr(strArr(2), 52, 1)) & ")"
        '组合SQL误句
        gstrSQL = strHead & strInPut
        ExecuteProcedure_ZLNJ "插入明细数据到中间库"
    Next
    保存明细结果到中间库 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function StartOrCommitorRollbackTransaction(ByVal bytType As Byte, Optional blnGcnoracle As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:启动、提交、回滚事务
    '--入参数:byttype-0启动,1提交,2回滚
    '         blnGcnoracle-是否存在事务(gcnoracle)
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Select Case bytType
        Case 0
            gcnOracle_成都内江.BeginTrans
            If Not blnGcnoracle Then
                gcnOracle.BeginTrans
            End If
            mblnStartTran = True
        Case 1
            gcnOracle_成都内江.CommitTrans
            If Not blnGcnoracle Then
                gcnOracle.CommitTrans
            End If
            mblnStartTran = False
        Case Else
            gcnOracle_成都内江.RollbackTrans
            If Not blnGcnoracle Then
                gcnOracle.RollbackTrans
            End If
            mblnStartTran = False
        End Select
End Function

Private Function 处方上传(ByVal lng记录性质 As Long, lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '处方明细上传
    '功能:上传新产生的记帐明细到医保中心
    '参数:  str单据号   NO
    '       int性质     记录性质
    '       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
    '返回:
    Dim rsTemp As New ADODB.Recordset, rs明细 As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String, strArr As Variant
    Dim lng病人ID As Long, str明细 As String
    Dim i As Long
    
    
    处方上传 = False
    
    Err = 0
    On Error GoTo ErrHand:


   '读出该张单据的费用明细
    gstrSQL = "Select A.ID,A.NO,A.病人ID,A.主页ID,to_char(A.登记时间,'yyyy-mm-dd hh24:mi:ss') as 登记时间,Round(A.实收金额,4) 实收金额 " & _
              "         ,A.收费细目ID,A.数次*nvl(A.付数,1) as 数量,Decode(A.数次*nvl(A.付数,1),0,0,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4)) as 价格 " & _
              "         Z.名称 as 开单部门,C.项目编码,C.大类编码,J.类别 as 收费类别,C.是否医保,B.编码,B.名称,A.是否急诊,nvl(A.开单人,A.操作员姓名) as 医生,A.操作员姓名,B.计算单位,E.规格,G.名称 剂型,M.医保号,M.就诊次数 " & _
              "  From 病人费用记录 A,部门表 Z,收费类别 J,收费细目 B,保险帐户 M,(Select O.*,Z.大类编码 From 保险支付项目 O,保险项目 Z where O.险类=Z.险类 and O.项目编码=Z.编码 and O.险类=" & gintInsure & ") C,病案主页 D,药品目录 E ,药品信息 F,药品剂型 G " & _
              "  where a.病人id=M.病人id and a.开单部门ID=Z.iD(+)   and M.险类=" & gintInsure & " and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=1 And Nvl(A.是否上传,0)=0 " & _
              "        and A.收费类别=J.编码(+)  and A.病人ID=D.病人ID and A.主页ID=D.主页ID And D.险类=" & gintInsure & _
              "        and A.收费细目ID=B.ID and A.收费细目ID=C.收费细目ID(+) " & _
              "        AND B.ID=E.药品ID(+) AND E.药名ID=F.药名ID(+) AND F.剂型=G.编码(+) " & _
              "  Order by A.病人ID,A.登记时间"

    Call OpenRecordset(rs明细, "处方明细上传")
    Dim lng冲销ID As Long

    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
        '上传明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '单价检查
        
            If Val(!数量) < 0 Or Val(!价格) < 0 Then
                ShowMsgbox "在单据中不能输入负单据!"
                Exit Function
            End If
            If Nvl(!项目编码) = "" Then
                 MsgBox "有项目未设置医保编码，不能上传明细!", vbInformation, gstrSysName
                 Exit Function
            End If
            .MoveNext
        Loop
    End With
    
    If rs明细.RecordCount <> 0 Then rs明细.MoveFirst
    
    Dim str交易流水号 As String
    Dim blnStarTran As Boolean '启动事务
    
    strInPut = ""
    mblnStartTran = False
    '进行费用传输
    With rs明细
        Do Until .EOF
                If mblnStartTran = False Then
                    '启动事务
                    Call StartOrCommitorRollbackTransaction(0)
                End If
               If lng病人ID <> Nvl(!病人ID, 0) And i >= 38 Then
                    i = 1
                    str交易流水号 = Get交易流水号
                    lng病人ID = Nvl(!病人ID, 0)
                    If strInPut <> "" Then
                        strInPut = strInPut & vbTab & i & vbTab & str明细
                        '请求相关的业务数据
                        If 业务请求_成都内江(住院交易上传_内江, strInPut, strOutPut) = False Then
                            '回滚事务
                            Call StartOrCommitorRollbackTransaction(2)
                            Exit Function
                        End If
                        If 保存明细结果到中间库(str交易流水号, strOutPut) = False Then
                            '提交事务,中间库数据未保存起
                            Call StartOrCommitorRollbackTransaction(1)
                            Exit Function
                        End If
                    End If
                    Call Get病人信息(lng病人ID)
                    
                    strInPut = Rpad(g病人身份_成都内江.个人编号, 8)
                    
                    strInPut = strInPut & vbTab & Rpad(g病人身份_成都内江.卡号, 10)
                    strInPut = strInPut & vbTab & Rpad(InitInfor_成都德阳.医院编码, 5)
                    strInPut = strInPut & vbTab & Rpad(g病人身份_成都内江.统筹编号, 1)
                    strInPut = strInPut & vbTab & Rpad(str交易流水号, 20)
                    If Nvl(!大类编码) = "1" Then
                        strInPut = strInPut & vbTab & IIf(IS出院带药(Nvl(!NO), Nvl(!ID, 0)), "1", "0")
                    Else
                        strInPut = strInPut & vbTab & "0"
                    End If
                    strInPut = strInPut & vbTab & Rpad(Substr(!开单部门, 1, 10), 10)
                    strInPut = strInPut & vbTab & Rpad(Substr(!医生, 1, 10), 10)
                    strInPut = strInPut & vbTab & Rpad(Substr(g病人身份_成都内江.住院流水号, 1, 20), 10)
                    str明细 = ""
                    '个人编号    String(8)   In
                    '社保卡号码  String(10)  In
                    '医院代码    String(5)   In
                    '统筹地区编码    String(1)   In
                    '医院交易流水号  String(20)  In
                    '出院带药类别    String(1)   In
                    
                    '科别    String(10)  In
                    '医生    String(10)  In
                    '住院流水号  String(20)  In
                    '处方条数    String(2)   In
                    '处方明细    String处方条数×51  In
               End If
                str明细 = str明细 & Substr(Rpad(Nvl(!大类编码), 1), 1, 1)
                str明细 = str明细 & Rpad(Nvl(!项目编码), 20)
                str明细 = str明细 & Lpad(Nvl(!数量) * 100, 10, "0")
                str明细 = str明细 & Rpad(Nvl(!规格), 10)
                str明细 = str明细 & Lpad(Nvl(!实收金额) * 100, 10, "0")
               i = i + 1
                gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & str交易流水号 & "')"
                ExecuteProcedure "打上上传标志"
            .MoveNext
        Loop
    End With
    
    If strInPut <> "" Then
        strInPut = strInPut & vbTab & i & vbTab & str明细
        '请求相关的业务数据
        If 业务请求_成都内江(住院交易上传_内江, strInPut, strOutPut) = False Then
            '提交事务,中间库数据未保存起
            Call StartOrCommitorRollbackTransaction(2)
            Exit Function
        End If
        If 保存明细结果到中间库(str交易流水号, strOutPut) = False Then
            '提交事务,中间库数据未保存起
            Call StartOrCommitorRollbackTransaction(1)
            Exit Function
        End If
    End If
    处方上传 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mblnStartTran Then
        '提交事务,中间库数据未保存起
        Call StartOrCommitorRollbackTransaction(2)
    End If
End Function
Private Function IS出院带药(ByVal strNO As String, lng费用ID As Long) As Boolean
    '检查是否出院带药
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select ID From 药品收发记录 where NO='" & strNO & "' and 单据 IN(9,10) and 费用id=" & lng费用ID & " and 扣率 like '_3%'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取是否出院带药"
    If rsTemp.EOF Then
        IS出院带药 = False
        Exit Function
    End If
    IS出院带药 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function 处方登记_成都内江(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------

    Dim lng病人ID As Long
    Dim lng主页id As Long
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String
    
    Err = 0
    On Error GoTo ErrHand:


    处方登记_成都内江 = False
    
    If lng记录状态 = 1 Then
        '正常单据
        If 处方上传(lng记录性质, lng记录状态, str单据号) = False Then
            Exit Function
        End If
    Else
        '开始事务
        Call StartOrCommitorRollbackTransaction(0)
        '冲销单据
        If 处方作废(lng记录性质, lng记录状态, str单据号) = False Then
            '提交事务,中间库数据未保存起,所以回滚
            Call StartOrCommitorRollbackTransaction(2)
            Exit Function
        End If
        '提交事务
        Call StartOrCommitorRollbackTransaction(1)
    End If
    处方登记_成都内江 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 处方作废(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:记帐处方作废,即记录状态=2的记录
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------
    Dim rs明细 As New ADODB.Recordset
    Dim rs原明细 As New ADODB.Recordset
    
    Dim rsTemp As New ADODB.Recordset
    Dim strInPut As String, strOutPut As String, str交易流水号 As String
    Dim strArr
    Dim lng病人ID As Long
    
    处方作废 = False

    Err = 0: On Error GoTo ErrHand:

    
    gstrSQL = " Select 摘要,A.ID,a.收费细目id,A.序号,A.数次*nvl(A.付数,1) as 数量,Round(A.实收金额/(A.数次*nvl(A.付数,1)),4) as 单价 " & _
              " From 病人费用记录 A,保险帐户 B " & _
              " where a.病人id=b.病人id and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=3 and   Nvl(附加标志,0)<>9  " & _
              " order by A.病人id,A.摘要"
              
    Call OpenRecordset(rs原明细, "处方明细上传")
    
    If rs原明细.EOF Then
        ShowMsgbox "该单据没有相应的明细记录,不能作废!"
        Exit Function
    End If

    gstrSQL = " Select * " & _
              " From 病人费用记录 A,保险帐户 b" & _
              " where a.病人id=b.病人id and A.NO='" & str单据号 & "' and A.记录性质=" & lng记录性质 & " and A.记录状态=2 and  Nvl(附加标志,0)<>9 AND nvl(a.是否上传,0)=0 " & _
              " order by A.病人ID"
              
    Call OpenRecordset(rs明细, "处方明细上传")

    lng病人ID = 0
    '更新原单据的值
    With rs明细
        Do While Not .EOF
            rs原明细.Filter = "序号=" & Nvl(!序号, 0) & "  and 收费细目id=" & Nvl(!收费细目ID, 0)
            If rs原明细.EOF Then
                ShowMsgbox "冲销时未找到相应的记录,冲销失败!"
                Exit Function
            End If
            str交易流水号 = Nvl(rs原明细!摘要)
            If str交易流水号 = "" Then
                ShowMsgbox "在原单中不存在交易流水号,不能继续！"
                Exit Function
            End If
            '检查消费明细中有交易流水号没有
            gstrSQL = "Select 医保流水号 From 医保消费信息 where 医院流水号='" & str交易流水号 & "' and 病人id=" & Nvl(!病人ID, 0)
            OpenRecordset_成都内江 rsTemp, gstrSQL, "获取医保消费信息"
            If rsTemp.EOF Then
                ShowMsgbox "不存在医保交易数据,请与系统管理员联系!"
                Exit Function
            End If
            If Nvl(rsTemp!医保流水号) = "" Then
                ShowMsgbox "不存在医保交易数据,请与系统管理员联系!"
                Exit Function
            End If
            '更新上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,'" & Nvl(rs原明细!摘要) & "')"
            ExecuteProcedure "打上上传标志"
            .MoveNext
        Loop
    End With
    Dim str摘要 As String
    
    lng病人ID = 0
    str摘要 = ""
    With rs原明细
        .MoveFirst
        Do While Not .EOF
                If lng病人ID <> Nvl(!病人ID, 0) And str摘要 <> Nvl(!摘要) Then
                    If lng病人ID <> Nvl(!病人ID, 0) Then
                        lng病人ID = Nvl(!病人ID, 0)
                        '需重新获取相关的病人信息
                        If Get病人信息(lng病人ID) = False Then
                            ShowMsgbox "在获取病人信息时间出了错误,请与系统员立即联系!"
                            Exit Function
                        End If
                    End If
                    str摘要 = Nvl(!摘要)
                    gstrSQL = "Select 医保流水号 From 医保消费信息 where 医院流水号='" & str摘要 & "' and 病人id=" & lng病人ID
                    OpenRecordset_成都内江 rsTemp, gstrSQL, "获取医保消费信息"
                    
                    strInPut = Rpad(g病人身份_成都内江.个人编号, 8)
                    strInPut = strInPut & Rpad(g病人身份_成都内江.卡号, 10)
                    strInPut = strInPut & Rpad(Substr(gstrUserName, 1, 10), 10)
                    strInPut = strInPut & Rpad(g病人身份_成都内江.统筹编号, 1)
                    strInPut = strInPut & Rpad(g病人身份_成都内江.住院流水号, 20)
                    strInPut = strInPut & Rpad(Nvl(rsTemp!医保流水号), 20)
                    
                    '取消
                    '    个人编号    String(8)   In
                    '    社保卡号码  String(10)  In
                    '    操作员卡号码    String(10)  In
                    '    统筹地区编码    String(1)   In
                    '    住院流水号  String(20)  In
                    '    医保交易流水号  String(20)  In
                    
                    If 业务请求_成都内江(住院产易上传取消_内江, strInPut, strOutPut) = False Then Exit Function
                End If
            .MoveNext
        Loop
    End With
    处方作废 = True
    Exit Function
ErrHand:
   If ErrCenter = 1 Then
        Resume
   End If
End Function
Private Function Read模拟数据(ByVal int业务类型 As 业务类型_成都内江, ByVal strInputString As String, ByRef strOutPutstring As String)
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
    
    If int业务类型 = 读取卡内数据 Then
        strFile = App.Path & "\解析卡.txt"
    Else
        strFile = App.Path & "\模拟提交串.txt"
    End If
    
    
    If Not Dir(strFile) <> "" Then
        objFile.CreateTextFile strFile
    End If
    strName = Get交易代码(int业务类型, True)
    
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
                If int业务类型 = 读取卡内数据 Then
                    strArr = Split(strText, vbTab & "|")
                    If Val(strArr(0)) = 1 Then
                            str = strArr(1)
                            Exit Do
                    End If
                Else
                        If blnStart Then
                            If strText = "" Then
                                strText = "" & vbTab & "|"
                            End If
                            strArr = Split(strText, vbTab & "|")
                            
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

Private Function Get病人信息(ByVal lng病人ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim strArr

    Get病人信息 = False
    
    Err = 0
    On Error GoTo ErrHand:
    'COMMENT ON COLUMN 保险帐户.病人ID   is '病人ID';
    'COMMENT ON COLUMN 保险帐户.险类     is '固定值:106';
    'COMMENT ON COLUMN 保险帐户.中心     is '0';
    'COMMENT ON COLUMN 保险帐户.卡号     is '卡号';
    'COMMENT ON COLUMN 保险帐户.医保号   is '个人编号';
    'COMMENT ON COLUMN 保险帐户.密码     is '无';
    'COMMENT ON COLUMN 保险帐户.人员身份 is '交易类别';
    'COMMENT ON COLUMN 保险帐户.单位编码 is '单位号码';
    '
    'COMMENT ON COLUMN 保险帐户.顺序号   is '只针对住院:住院流水号';
    'COMMENT ON COLUMN 保险帐户.退休证号 is '统筹地区编码|制卡日期|卡有效日期|制卡单位|在职情况';
    'COMMENT ON COLUMN 保险帐户.帐户余额 is '帐户余额';
    'COMMENT ON COLUMN 保险帐户.当前状态 is '0-门诊,1-在院';
    'COMMENT ON COLUMN 保险帐户.病种ID   is '无';
    'COMMENT ON COLUMN 保险帐户.在职     is '目前保存的值是1，无用处';
    'COMMENT ON COLUMN 保险帐户.年龄段   is '补卡次数';
    'COMMENT ON COLUMN 保险帐户.灰度级   is '工况类别';
    'COMMENT ON COLUMN 保险帐户.就诊时间 is '当前就诊的时间';
    '
    'COMMENT ON COLUMN 保险帐户.享受待遇标志 is '只针对住院:享受待遇标志';
    'COMMENT ON COLUMN 保险帐户.起付标准 is '只针对住院:起付标准';
    
    gstrSQL = "select a.*,b.姓名,b.性别, b.年龄, b.出生日期, b.身份证号,b.工作单位 " & _
             " from 保险帐户 a,病人信息 b " & _
             " WHERE a.病人id=" & lng病人ID & " AND a.病人id=b.病人id and a.险类=" & TYPE_成都内江

    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取病人信息"

    With g病人身份_成都内江
        .卡号 = Nvl(rsTemp!卡号)
        .个人编号 = Nvl(rsTemp!个人编号)
        .身份证号 = Nvl(rsTemp!身份证号)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Decode(Nvl(rsTemp!性别), "男", 1, "女", 2, 1)
        .工况类别 = Nvl(rsTemp!灰度级)
        .出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
        .单位号码 = Nvl(rsTemp!单位编码)
        strArr = Split(Nvl(rsTemp!退休证号) & "|||||", "|")
        .统筹编号 = strArr(0)
        .制卡日期 = strArr(1)
        .卡有效期 = strArr(2)
        .补卡次数 = Nvl(rsTemp!年龄段)
        .制卡单位 = strArr(3)
        .年龄 = Nvl(rsTemp!年龄, 0)
        .帐户余额 = Nvl(rsTemp!帐户余额, 0)
        .在职情况 = strArr(4)
        .交易类别 = Nvl(rsTemp!人员身份)
        .住院流水号 = Nvl(rsTemp!顺序号)
    End With
    Get病人信息 = True
Exit Function
ErrHand:
        DebugTool "获取病人信息失败" & vbCrLf & " 错误号:" & Err.Number & vbCrLf & " 错误信息:" & Err.Description
End Function

Private Sub OpenRecordset_成都内江(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "", Optional cnOracle As ADODB.Connection)
    '功能：打开记录集
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    If cnOracle Is Nothing Then
        rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_成都内江, adOpenStatic, adLockReadOnly
    Else
        If cnOracle.State <> 1 Then
            rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_成都内江, adOpenStatic, adLockReadOnly
        Else
            rsTemp.Open IIf(strSql = "", gstrSQL, strSql), cnOracle, adOpenStatic, adLockReadOnly
        End If
    End If
    Call SQLTest
End Sub


Public Function 住院虚拟结算_成都内江(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；

    Dim cn上传 As New ADODB.Connection, rsTemp As New ADODB.Recordset
    Dim rs明细 As New ADODB.Recordset
    Dim lng主页id As Long
    Dim strInPut As String, strOutPut   As String
    Dim strArr As Variant
    Dim intMouse As Integer

    Err = 0: On Error GoTo ErrHand:

    g病人身份_成都内江.lng病人ID = lng病人ID
    If rsExse.RecordCount = 0 Then
        MsgBox "该病人没有有发生费用，无法进行结算操作。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Get病人信息(lng病人ID) = False Then Exit Function

    If bln结帐处 Then
        Screen.MousePointer = 1
        If 身份标识_成都内江(4, lng病人ID) = "" Then
            Screen.MousePointer = intMouse
            住院虚拟结算_成都内江 = ""
            Exit Function
        End If
        If lng病人ID <> g病人身份_成都内江.lng病人ID Then
            ShowMsgbox "你的卡可能有误,不能进行结算!"
            Exit Function
        End If
        Screen.MousePointer = intMouse
    End If
    
    gstrSQL = "select MAX(主页ID) AS 主页ID from 病案主页 where 病人ID=" & rsExse("病人ID")
    Call OpenRecordset(rsTemp, "虚拟结算")
    If IsNull(rsTemp("主页ID")) = True Then
        MsgBox "只有住院病人才可以使用医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    lng主页id = rsTemp("主页ID")
    Screen.MousePointer = vbHourglass
    
    '补传明细
    If 补传住院明细记录(lng病人ID, lng主页id) = False Then Exit Function
    住院虚拟结算_成都内江 = "个人帐户|0"
    g病人身份_成都内江.lng病人ID = lng病人ID   '表示该病人已经进行了虚拟结算
    
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function 补传住院明细记录(ByVal lng病人ID As Long, ByVal lng主页id As Long) As Boolean
    '补传相关明细记录
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim strInPut  As String, strOutPut As String
    Dim strArr, strArr摘要
    Dim lng冲销ID As Long
    Err = 0
    On Error GoTo ErrHand:


    补传住院明细记录 = False

    '读出未上传明细（排序，以便先上传正明细，再上传负明细）
    gstrSQL = "" & _
        "   Select distinct A.NO,A.记录性质,A.记录状态 " & _
        "   From 病人费用记录 A " & _
        "   Where A.病人ID=" & lng病人ID & " and A.主页ID=" & lng主页id & " and A.记帐费用=1  and A.实收金额<>0 and nvl(A.是否上传,0)=0 And Nvl(A.记录状态,0)<>0 " & _
        "   Order by A.记录性质,A.NO,Decode(A.记录状态,2,2,1)"
        
    
 
    '先检查是否存在退单情况，如果存在，看有否对应的记录单.
    With rs明细
        '上传明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!项目编码) = "" Then
                ShowMsgbox "项目:[" & Nvl(!编码) & "] 未设置对应的医保项目,请设置对应关系!"
                Exit Function
            End If
            If (Val(!数量) < 0 Or Val(!价格) < 0) And rs明细!记录状态 = 1 Then
                ShowMsgbox "项目:[" & Nvl(!编码) & "] 不能输入负单据!"
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    '先传正单据
    With rs明细
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!记录状态, 1) = 1 Then
                '上传指定处方
                If 处方上传(Nvl(!记录性质, 0), Nvl(!记录状态, 0), Nvl(!NO)) = False Then
                    Exit Function
                End If
            Else
                '上传指定处方
                gcnOracle_成都内江.BeginTrans
                gcnOracle.BeginTrans
                If 处方作废(Nvl(!记录性质, 0), Nvl(!记录状态, 0), Nvl(!NO)) = False Then
                    gcnOracle.RollbackTrans
                    gcnOracle_成都内江.RollbackTrans
                    Exit Function
                    
                End If
                gcnOracle.CommitTrans
                gcnOracle_成都内江.CommitTrans
            End If
           .MoveNext
        Loop
    End With
    补传住院明细记录 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function Get原单据摘要(ByVal strNO As String, ByVal int序号 As Integer, ByVal int性质 As Integer) As Variant
    '根据指定的值，获取摘要的相关信息
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
        
    
    gstrSQL = " Select 摘要 From 病人费用记录" & _
              " Where NO='" & strNO & "' And 序号=" & int序号 & _
              " And 记录性质=" & int性质 & " And 记录状态=3"
    
    Call OpenRecordset(rsTemp, "取原始处方明细的流水号")
    
    If Not rsTemp.EOF Then
        strTemp = Nvl(rsTemp!摘要) & "|||||||"
    Else
        strTemp = "|||||||"
    End If
    Get原单据摘要 = Split(strTemp, "|")
End Function

'----200410刘兴宏加入
Public Function 医保设置_成都内江() As Boolean
    医保设置_成都内江 = frmSet成都内江.参数设置
    
End Function
'
'Public Function 下载服务项目目录_成都内江(ByVal bytType As Byte, ByVal objProgss As Object) As Boolean
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:下载服务项目目录
'    '参数:bytType-1-药品,2-诊疗,3-服务,4-费用类别,5-病种目录
'    '返回:下载成功,返回true,否则返回False
'    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSql As String
'    Dim rsTemp As New ADODB.Recordset
'    Dim strDate As String, strInput As String, strOutput As String
'    Dim lngCount As Long
'    Dim i As Long
'    Dim strArr
'    Dim strTitle As String
'
'    下载服务项目目录_成都内江 = False
'    strTitle = Switch(bytType = 1, "药品", bytType = 2, "诊疗项目", bytType = 3, "服务设施", bytType = 4, "费用类别", True, "病种目录取")
'
'    Err = 0
'    On Error GoTo ErrHand:
'    strSql = "" & _
'        "   Select to_char(Max(变更时间),'yyyy-mm-dd hh24:mi:ss')  as 变更时间 " & _
'        "   From 医保收费目录 " & _
'        "   where 类别=" & bytType
'    zlDatabase.OpenRecordset rsTemp, strSql, "获取最大变更时间"
'
'    strDate = Nvl(rsTemp!变更时间)
'    strDate = IIf(strDate = "", "1977-01-01 00:00:00", strDate)
'
'    If Not objProgss Is Nothing Then
'    Else
'        zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,请等待..."
'    End If
'    '预处理
'    strInput = bytType & "|" & strDate
'    If 业务请求_成都内江(收费目录下载预处理, strInput, strOutput) = False Then Exit Function
'    strArr = Split(strOutput, "|")
'    lngCount = Val(strArr(1))
'
'    If Not objProgss Is Nothing Then
'        objProgss.Max = IIf(lngCount = 0, 1, lngCount)
'        objProgss.Min = 1
'        objProgss.Value = 1
'    End If
'
'   For i = 1 To lngCount
'        '正试下载
'        If 业务请求_成都内江(收费目录下载处理, strInput, strOutput) = False Then Exit Function
'        strArr = Split(strOutput, "|")
'        '更新收费目录
'
'        '过程:类别,编码,名称,英文名称,收费类别,收费等级,费用等级,拼音码,单位,单价,规格,备注,变更时间,可维护标志,支付标准
'        strSql = "ZL_医保收费目录_UPDATE("
'        strSql = strSql & bytType & ",'"
'        strSql = strSql & strArr(1) & "','" '编码
'        strSql = strSql & strArr(2) & "','" '名称
'        Select Case bytType
'        Case 1
'            strSql = strSql & strArr(3) & "','" '英文名称
'            strSql = strSql & strArr(4) & "','" '收费类别
'            strSql = strSql & strArr(5) & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & strArr(7) & "','" '单位
'            strSql = strSql & strArr(8) & "','" '单价
'            strSql = strSql & strArr(9) & "','" '剂型
'            strSql = strSql & strArr(10) & "','" '规格
'            strSql = strSql & strArr(11) & "',to_date('" '备注
'            strSql = strSql & strArr(12) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(13) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 2
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(4) & "','" '拼音码
'            strSql = strSql & strArr(5) & "','" '单位
'            strSql = strSql & strArr(6) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & strArr(7) & "',to_date('" '备注
'            strSql = strSql & strArr(8) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & strArr(9) & "','"     '可维护标志
'            strSql = strSql & "" & "')" '支付标准
'        Case 3
'            strSql = strSql & "" & "','" '英文名称
'            strSql = strSql & strArr(3) & "','" '收费类别
'            strSql = strSql & "" & "','" '费用等级
'            strSql = strSql & strArr(6) & "','" '拼音码
'            strSql = strSql & "" & "','" '单位
'            strSql = strSql & strArr(4) & "','" '单价
'            strSql = strSql & "" & "','" '剂型
'            strSql = strSql & "" & "','" '规格
'            strSql = strSql & "" & "',to_date('" '备注
'            strSql = strSql & strArr(7) & "','yyyy-mm-dd hh24:mi:ss'),'"  '变更时间
'            strSql = strSql & "" & "',"     '可维护标志
'            strSql = strSql & strArr(5) & "')" '支付标准
'        Case 4
'            ' 费用类别编码|费用类别名称
'            strSql = "ZL_医保收费类别_UPDATE("
'
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "')" '名称
'        Case Else
'            '病种编码|病种名称|拼音码|变更日期
'            strSql = "ZL_医保病种目录_UPDATE("
'            strSql = strSql & strArr(1) & "','" '编码
'            strSql = strSql & strArr(2) & "','" '名称
'            strSql = strSql & strArr(3) & "',to_date('" '助记码
'            strSql = strSql & strArr(4) & "','yyyy-mm-dd hh24:mi:ss')" '变更时间
'        End Select
'        gcnOracle_成都内江.Execute strSql, , adCmdStoredProc
'        If Not objProgss Is Nothing Then
'            objProgss.Value = i
'        Else
'            zlCommFun.ShowFlash "正在下载《" & strTitle & "》数据,已下载" & i & "/" & lngCount & ""
'        End If
'   Next
'   下载服务项目目录_成都内江 = True
'   Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'End Function

Public Function 获取参保人员信息_成都内江(ByVal strInPut As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------

    '获取参保人员信息
    Dim strOutPut As String
    Dim strArr
    
    获取参保人员信息_成都内江 = False
    
    Err = 0
    On Error GoTo ErrHand:
    
    If 业务请求_成都内江(读病人信息_内江, strInPut, strOutPut) = False Then Exit Function
    '返回串是:卡号vbtab个人编号vbtab身份证号vbtab姓名vbtab性别vbtab工况类别vbtab出生日期vbtab单位号码vbtab统筹编号vbtab制卡日期vbtab卡有效期vbtab补卡次数vbtab制卡单位
    If strOutPut = "" Then Exit Function
    strArr = Split(strOutPut, vbTab)
    
    With g病人身份_成都内江
        .卡号 = strArr(0)
        .个人编号 = strArr(1)
        .身份证号 = strArr(2)
        .姓名 = strArr(3)
        .性别 = strArr(4)
        .工况类别 = strArr(5)
        .出生日期 = zlCommFun.AddDate(strArr(6))
        .单位号码 = strArr(7)
        .统筹编号 = strArr(8)
        .制卡日期 = strArr(9)
        .卡有效期 = strArr(10)
        .补卡次数 = strArr(11)
        .制卡单位 = strArr(12)
        .年龄 = Get年龄(.出生日期)
    End With
    
    '--获取个人帐户余额
    If 获取帐户余额_成都内江() = False Then Exit Function
    
    获取参保人员信息_成都内江 = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Public Function 获取帐户余额_成都内江() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取病人当前帐户余额
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String, strOutPut As String
    Dim strArr
    Err = 0
    On Error GoTo ErrHand:
    获取帐户余额_成都内江 = False
    With g病人身份_成都内江
        '    个人编号    String (8)  IN
        '    社保卡号码  String (10) IN
        '    统筹地区编码    String (1)  IN
        strInPut = .个人编号
        strInPut = strInPut & vbTab & .卡号
        strInPut = strInPut & vbTab & .统筹编号
    End With
    
    If 业务请求_成都内江(获取帐户余额_内江, strInPut, strOutPut) = False Then Exit Function
    If strOutPut = "" Then Exit Function
    strArr = Split(strOutPut, vbTab)
    With g病人身份_成都内江
        .帐户余额 = Val(strArr(0))
        .在职情况 = strArr(1)
    End With
    获取帐户余额_成都内江 = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then Resume
End Function
Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    OpenRecordset rsTemp, "获取年龄"
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
ErrHand:
End Function

Private Function GetErrInfor(ByVal strErrCode As String) As String
        Dim strErrMsg As String
        
        Select Case strErrCode
                '---读卡的相关错误
                Case "1 ": strErrMsg = " 错误号：1 " & vbCrLf & " 错误描述：检测通讯方式错误(chk_baud异常错误)"
                Case "2 ": strErrMsg = " 错误号：2 " & vbCrLf & " 错误描述：始化端口错误(auto_init)"
                Case "3 ": strErrMsg = " 错误号：3 " & vbCrLf & " 错误描述：关闭通讯口错误(ic_exit)"
                Case "4 ": strErrMsg = " 错误号：4 " & vbCrLf & " 错误描述：读写器错误"
                Case "5 ": strErrMsg = " 错误号：5 " & vbCrLf & " 错误描述：无法初始化卡密码"
                Case "10": strErrMsg = " 错误号：10" & vbCrLf & " 错误描述： 检测读写器中是否有卡错误(get_status)"
                Case "11": strErrMsg = " 错误号：11" & vbCrLf & " 错误描述： 卡型错误（非4428卡）(chk_4428)"
                Case "12": strErrMsg = " 错误号：12" & vbCrLf & " 错误描述： 卡密码错误(csc_4428)"
                Case "13": strErrMsg = " 错误号：13" & vbCrLf & " 错误描述： 修改卡密码错误"
                Case "20": strErrMsg = " 错误号：20" & vbCrLf & " 错误描述： 读卡芯数据错误(srd_4428)"
                Case "21": strErrMsg = " 错误号：21" & vbCrLf & " 错误描述： 写卡芯数据(用户数据)错误(swr_4428)"
                Case "23": strErrMsg = " 错误号：23" & vbCrLf & " 错误描述： 写卡芯数据(用户密码)错误(swr_4428)"
                Case "30": strErrMsg = " 错误号：30" & vbCrLf & " 错误描述： 用户密码错误"
                Case "31": strErrMsg = " 错误号：31" & vbCrLf & " 错误描述： 用户数据加密错误(ic_decrypt)"
                Case "32": strErrMsg = " 错误号：32" & vbCrLf & " 错误描述： 用户数据解密错误(ic_decrypt)"
                Case "33": strErrMsg = " 错误号：33" & vbCrLf & " 错误描述： 用户密码加密错误(ic_encrypt)"
                Case "34": strErrMsg = " 错误号：34" & vbCrLf & " 错误描述： 用户密码解密错误(ic_decrypt)"
                Case "35": strErrMsg = " 错误号：35" & vbCrLf & " 错误描述： 用户原密码长度为零或者大于6"
                Case "36": strErrMsg = " 错误号：36" & vbCrLf & " 错误描述： 用户新密码长度为零或者大于6"
                Case "40": strErrMsg = " 错误号：40" & vbCrLf & " 错误描述：   不能打开数据库"
                Case "41": strErrMsg = " 错误号：41" & vbCrLf & " 错误描述：   没有制卡数据"
                Case "42": strErrMsg = " 错误号：42" & vbCrLf & " 错误描述：   个人信息不完整（姓名、性别、民族等文本信息）"
                '---医保接口返回的相关错误
                Case "000": strErrMsg = "执行成功"
                Case "001": strErrMsg = " 错误号： 001" & vbCrLf & " 错误描述：读卡器无响应"
                Case "002": strErrMsg = " 错误号： 002" & vbCrLf & " 错误描述：没有社保卡"
                Case "003": strErrMsg = " 错误号： 003" & vbCrLf & " 错误描述：社保卡无响应"
                Case "004": strErrMsg = " 错误号： 004" & vbCrLf & " 错误描述：主机无响应"
                Case "051": strErrMsg = " 错误号： 051" & vbCrLf & " 错误描述：输入参数不足"
                Case "052": strErrMsg = " 错误号： 052" & vbCrLf & " 错误描述：卡号与社保号码不符"
                Case "053": strErrMsg = " 错误号： 053" & vbCrLf & " 错误描述：处方明细与实际纪录数量不符"
                Case "054": strErrMsg = " 错误号： 054" & vbCrLf & " 错误描述：没有此交易流水号"
                Case "055": strErrMsg = " 错误号： 055" & vbCrLf & " 错误描述：处方项目不符"
                Case "056": strErrMsg = " 错误号： 056" & vbCrLf & " 错误描述：没有此住院流水号（输入医保交易号+输入的住院流水号与库表中医保流水号对应的住院流水号不一致）等等"
                Case "058": strErrMsg = " 错误号： 058" & vbCrLf & " 错误描述：重复业务操作例如：已住院还进行住院操作，已出院还进行出院操作"
                Case "059": strErrMsg = " 错误号： 059" & vbCrLf & " 错误描述：输入社保号码和对应交易流水号不一致(住院时用住院流水号对应)"
                Case "060": strErrMsg = " 错误号： 060" & vbCrLf & " 错误描述：流水号不为最大撤消时"
                Case "061": strErrMsg = " 错误号： 061" & vbCrLf & " 错误描述：交易未确认上传住院交易前"
                Case "062": strErrMsg = " 错误号： 062" & vbCrLf & " 错误描述：未进行住院登记"
                Case "071": strErrMsg = " 错误号： 071" & vbCrLf & " 错误描述：医院交易流水号异常由HIS系统生成（空，长度不正常）"
                Case "072": strErrMsg = " 错误号： 072" & vbCrLf & " 错误描述：重复数据包传送"
                Case "073": strErrMsg = " 错误号： 073" & vbCrLf & " 错误描述：交叉数据包传送"
                Case "074": strErrMsg = " 错误号： 074" & vbCrLf & " 错误描述：应该上传明细而没有上传明细"
                Case "075": strErrMsg = " 错误号： 075" & vbCrLf & " 错误描述：检查项目种类异常项目种类不在[1]，[2]之内"
                Case "077": strErrMsg = " 错误号： 077" & vbCrLf & " 错误描述：上传医疗机构编码异常在KB01之中不存在"
                Case "078": strErrMsg = " 错误号： 078" & vbCrLf & " 错误描述：非定点医疗机构"
                Case "079": strErrMsg = " 错误号： 079" & vbCrLf & " 错误描述：社保卡号与医保交易号不对应门诊消费撤销时，只允许用消费的卡来撤销交易"
                Case "080": strErrMsg = " 错误号： 080" & vbCrLf & " 错误描述：医疗机构编号与交易流水号不对应(住院时用住院流水号对应)只允许交易医疗机构撤销自己的交易"
                Case "081": strErrMsg = " 错误号： 081" & vbCrLf & " 错误描述：对应流水号明细不存在Kc07,KC08K1,Kc08k2"
                Case "082": strErrMsg = " 错误号： 082" & vbCrLf & " 错误描述：没有对应的住院流水号Kc08"
                Case "083": strErrMsg = " 错误号： 083" & vbCrLf & " 错误描述：不允许撤销已出院的交易已经出院的交易补允许撤销"
                Case "084": strErrMsg = " 错误号： 084" & vbCrLf & " 错误描述：出院医院不是入院医院"
                Case "085": strErrMsg = " 错误号： 085" & vbCrLf & " 错误描述：输入输出参数长度错误输入长度与约定长度不符"
                Case "086": strErrMsg = " 错误号： 086" & vbCrLf & " 错误描述：出院人员不是原来那个住院人员防止一卡号对应多个人编号的情况"
                Case "101": strErrMsg = " 错误号： 101" & vbCrLf & " 错误描述：个人状态异常"
                Case "102": strErrMsg = " 错误号： 102" & vbCrLf & " 错误描述：社保卡为黑名单卡"
                Case "103": strErrMsg = " 错误号： 103" & vbCrLf & " 错误描述：帐户被冻结"
                Case "104": strErrMsg = " 错误号： 104" & vbCrLf & " 错误描述：不能享受统筹待遇"
                Case "106": strErrMsg = " 错误号： 106" & vbCrLf & " 错误描述：不存在此人"
                Case "107": strErrMsg = " 错误号： 107" & vbCrLf & " 错误描述：没有参加医疗保险"
                Case "108": strErrMsg = " 错误号： 108" & vbCrLf & " 错误描述：帐户被注销只有死亡，出国定居才存在注销"
                Case "109": strErrMsg = " 错误号： 109" & vbCrLf & " 错误描述：出院已上传，但未确认出院上传时"
                Case "110": strErrMsg = " 错误号： 110" & vbCrLf & " 错误描述：出院已确认出院确认传时"
                Case "111": strErrMsg = " 错误号： 111" & vbCrLf & " 错误描述：没有此人医保卡数据T_cardinfo中"
                Case "112": strErrMsg = " 错误号： 112" & vbCrLf & " 错误描述：挂失卡已经挂失"
                Case "113": strErrMsg = " 错误号： 113" & vbCrLf & " 错误描述：卡状态异常"
                Case "114": strErrMsg = " 错误号： 114" & vbCrLf & " 错误描述：不存在个人账户Kc04无数据"
                Case "115": strErrMsg = " 错误号： 115" & vbCrLf & " 错误描述：系统没有当年起付线参数"
                Case "116": strErrMsg = " 错误号： 116" & vbCrLf & " 错误描述：没有当年基本医疗保险报销比例"
                Case "117": strErrMsg = " 错误号： 117" & vbCrLf & " 错误描述：没有当年大病医疗保险报销比例"
                Case "118": strErrMsg = " 错误号： 118" & vbCrLf & " 错误描述：没有当年高额医疗保险报销比例"
                Case "119": strErrMsg = " 错误号： 119" & vbCrLf & " 错误描述：没有当年公务员医疗保险报销比例"
                Case "120": strErrMsg = " 错误号： 120" & vbCrLf & " 错误描述：入院时间无效"
                Case "121": strErrMsg = " 错误号： 121" & vbCrLf & " 错误描述：没有单位封锁信息Kb02"
                Case "150": strErrMsg = " 错误号： 150" & vbCrLf & " 错误描述：支出金额为负数输入的支出的金额为负数"
                Case "151": strErrMsg = " 错误号： 151" & vbCrLf & " 错误描述：个人帐户超支（扣减为负）账户扣出负数来"
                Case "152": strErrMsg = " 错误号： 152" & vbCrLf & " 错误描述：基本统筹超支"
                Case "153": strErrMsg = " 错误号： 153" & vbCrLf & " 错误描述：大病统筹超支"
                Case "154": strErrMsg = " 错误号： 154" & vbCrLf & " 错误描述：公务员医疗补助超支"
                Case "155": strErrMsg = " 错误号： 155" & vbCrLf & " 错误描述：账户支付费用超过个人应当支付费用"
                Case "255": strErrMsg = " 错误号： 255" & vbCrLf & " 错误描述：服务程序出错"
                Case "998": strErrMsg = " 错误号： 998" & vbCrLf & " 错误描述：获取各类流水号失败"
                Case "999": strErrMsg = " 错误号： 999" & vbCrLf & " 错误描述：数据库sql错误，或者未找到数据"
                Case "800": strErrMsg = " 错误号： 800" & vbCrLf & " 错误描述：未办理出院手续，却试图进行出院确认操作"
                Case "801": strErrMsg = " 错误号： 801" & vbCrLf & " 错误描述：已经出院确认，再次试图办理出院确认"
            Case Else
                strErrMsg = "不能确定的错误代码,代码号为" & strErrCode
    End Select
    GetErrInfor = strErrMsg
End Function
Private Sub ExecuteProcedure_ZLNJ(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_成都内江.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function 结算方式更正() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:更正及显示结算结果
    '--入参数:
    '--出参数:str结算方式
    '--返  回:成功返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    Dim dbl费用总额 As Double
        
    '费用总额=病人自费金额+基本统筹支付金额+大病统筹金额      此解释是由刘兴宏根据以面公式转换而来的
    
    '病人自费金额 = 总费用额 - 基本统筹支付金额 - 大病 / 高额统筹支付金额
    '自费金额＝现金支付额＋帐户支付额 (即:可选择由现金或用帐户支付)
    '大病统筹与高额统筹意义相同
    '统筹支付金额等于医保内费用根据不同的起付标准和报销比例由医保中心算
    '此说明依据北京科瑞奇技术开发股份有限公司蒋红彬负责的解释
    结算方式更正 = False
    
    Err = 0
    On Error GoTo ErrHand:
    DebugTool "进入(" & "Get结算方式" & ")"
    
'    If g结算数据.结算标志 = 0 Then
        dbl费用总额 = g结算数据.医保内费用 + g结算数据.医保外费用
'    End If
    
    
    If g结算数据.帐户可用余额 >= g结算数据.医保内费用 Then
        g结算数据.帐户支付 = g结算数据.医保内费用
    Else
        g结算数据.帐户支付 = g结算数据.帐户可用余额
    End If
    
    str结算方式 = "||个人帐户|" & g结算数据.帐户支付
    If g结算数据.基本医保支付 <> 0 Then
        str结算方式 = str结算方式 & "||基本统筹|" & g结算数据.基本医保支付
    End If
    If g结算数据.高额医保支付 <> 0 Then
        str结算方式 = str结算方式 & "||大病支付|" & g结算数据.高额医保支付
    End If
    If g结算数据.公务员医疗补助 <> 0 Then
        str结算方式 = str结算方式 & "||公务员补助|" & g结算数据.公务员医疗补助
    End If
    
    If Format(g病人身份_成都内江.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '费用总额与医保中心返回总额不致,不能进行结算
        ShowMsgbox "本次结算总额(" & g病人身份_成都内江.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的总额(" & dbl费用总额 & ")不致产能结算?"
        Exit Function
    End If
    
   '如果存在,则保存冲预交记录中
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        g病人身份_成都内江.结算方式 = str结算方式
        
        If g结算数据.结算标志 = 0 Then
            gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "', 0)"
            Call ExecuteProcedure("更新预交记录")
        Else
                gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐ID & ",'" & str结算方式 & "',1)"
                Call ExecuteProcedure("更新预交记录")
        End If
    End If
    
    '显示结算信息
    If frm结算信息.ShowME(g结算数据.结帐ID, False, "个人帐户:" & g结算数据.帐户支付, IIf(g结算数据.结算标志 = 0, 0, 1)) = False Then
        结算方式更正 = False
        Exit Function
    End If
    结算方式更正 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function Get交易代码(ByVal intType As 业务类型_成都内江, Optional bln读名称 As Boolean = False) As String
    Select Case intType
        Case 读病人信息_内江
            Get交易代码 = IIf(bln读名称, "读病人信息", "01")
        Case 更改密码_内江
            Get交易代码 = IIf(bln读名称, "更改密码", "02")
        Case 获取帐户余额_内江
            Get交易代码 = IIf(bln读名称, "获取帐户余额", "03")
        Case 门诊明细写入_内江
            Get交易代码 = IIf(bln读名称, "门诊明细写入", "04")
        Case 门诊消费确认_内江
            Get交易代码 = IIf(bln读名称, "门诊消费确认", "05")
        Case 门诊消费取消_内江
            Get交易代码 = IIf(bln读名称, "门诊消费取消", "06")
        Case 住院登记_内江
            Get交易代码 = IIf(bln读名称, "住院登记", "07")
        Case 出院登记上传_内江
            Get交易代码 = IIf(bln读名称, "出院登记上传", "08")
        Case 住院交易上传_内江
            Get交易代码 = IIf(bln读名称, "住院交易上传", "09")
        Case 住院产易上传取消_内江
            Get交易代码 = IIf(bln读名称, "住院产易上传取消", "10")
        Case 获取单位欠缴情况_内江
            Get交易代码 = IIf(bln读名称, "获取单位欠缴情况", "11")
        Case 初始化函数_内江
            Get交易代码 = IIf(bln读名称, "初始化函数", "12")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function

Public Function 业务请求_成都内江(ByVal intType As 业务类型_成都内江, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInPut As String, lngReturn As Long, strReturn As String
    Dim strOutPut(0 To 20) As String, dblOutPut(0 To 25) As Double, intOutPut(0 To 5) As Integer, lngOutPut(0 To 5) As Long
    Dim strArr1
    Dim strArr(0 To 20) As String
    Dim str业务 As String
    Dim strReg As String

    Dim i As Integer
    
    str业务 = Get交易代码(intType, True)
    
    DebugTool "进入业务请求函数(业务类型代码为:" & intType & " 业务名称：" & str业务 & ")" & vbCrLf & "        输入参数为:" & strInputString
    
    
    业务请求_成都内江 = False
    
    strInPut = strInputString
    
    If InitInfor_成都内江.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, strInputString, strOutPutstring
         业务请求_成都内江 = True
        Exit Function
    End If
   
    
    strArr1 = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr1)
        strArr(i) = strArr1(i)
    Next
        
    
    Err = 0
    On Error GoTo ErrHand:
    
    Select Case intType
        Case 读病人信息_内江
            If InitInfor_成都内江.读卡器_内江 = 0 Then
                '输入参数:
                lngReturn = GetCardInfo_MW(Val(strArr(0)), strArr(1), strOutPut(0), strOutPut(1), strOutPut(2), strOutPut(3), strOutPut(4), strOutPut(5), strOutPut(6), strOutPut(7), strOutPut(8), strOutPut(9), strOutPut(10), strOutPut(11), strOutPut(12))
            Else
                lngReturn = GetCardInfo_KRQ(Val(strArr(0)), strOutPut(0), strOutPut(1), strOutPut(2), strOutPut(3), strOutPut(4), strOutPut(5), strOutPut(6), strOutPut(7), strOutPut(8), strOutPut(9), strOutPut(10), strOutPut(11), strOutPut(12))
            End If
           '构建返回串
           strReturn = strOutPut(0) & vbTab & strOutPut(1) & vbTab & strOutPut(2) & vbTab & strOutPut(3) & vbTab & strOutPut(4) & vbTab & strOutPut(5) & vbTab & strOutPut(6) & vbTab & strOutPut(7) & vbTab & strOutPut(8) & vbTab & strOutPut(9) & vbTab & strOutPut(10) & vbTab & strOutPut(11) & vbTab & strOutPut(12)
        Case 更改密码_内江
            lngReturn = ChangePassword(Val(strArr(0)), strArr(1), strArr(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(CStr(lngReturn))
                 Exit Function
            End If
        Case 获取帐户余额_内江
            '输入参数:  个人编号    String (8)  IN
            '           社保卡号码  String (10) IN
            '           统筹地区编码    String (1)  IN
            '输出参数:  帐户余额    Long    OUT
            '           在职情况    String(1)   OUT
            lngReturn = gobj成都内江.GetAccountAmountFunc(strArr(0), strArr(1), strArr(2), dblOutPut(0), strOutPut(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = dblOutPut(0) & vbTab & strOutPut(0)
        Case 门诊明细写入_内江
            '输入参数:  个人编号    String(8)   In
            '           社保卡号码  String(10)  In
            '           医院代码    String(5)   In
            '           操作员卡号码    String(10)  In
            '           统筹地区编码    String(1)   In
            '           医院交易流水号  String(20)  In
            '           交易类别    String(1)   In
            '           处方条数    String(2)   In
            '           处方明细    String处方条数×51  In

            '输出参数:  医保流水号  String(20)  Out
            '           医保内费用  String(10)  Out
            '           医保外费用  String(10)  Out
            '           个人帐户可用余额    String(10)  Out
            '           处方明细    String处方条数×51  Out
            '           在职情况    String(1)   Out
            '
            lngReturn = gobj成都内江.DoConsumeTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strOutPut(0), strOutPut(1), strOutPut(2), strOutPut(3), strOutPut(4), strOutPut(5))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0) & vbTab & strOutPut(1) & vbTab & strOutPut(2) & vbTab & strOutPut(3) & vbTab & strOutPut(4) & vbTab & strOutPut(5)
        Case 门诊消费确认_内江
            '输入参数: 个人编号    String(8)   In
            '          社保卡号码  String(10)  In
            '          医院代码    String(5)   In
            '          操作员卡号码    String(10)  In
            '          统筹地区编码    String(1)   In
            '          医保交易流水号  String(20)  In
            '          交易类别    String(1)   In
            '          个人帐户支付    String(10)  In
            lngReturn = gobj成都内江.DoConsumeAffirmFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
        Case 门诊消费取消_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        医保交易流水号  String(20)  In
            '        交易类别    String(1)   In

            lngReturn = gobj成都内江.DoConsumeCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
        Case 住院登记_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        入院日期    String(8)   In
            '        入院科别    String(10)  In
            '        入院诊治医生    String(10)  In
            '        诊断编码    String(20)  In
            
            '输出参数:住院流水号  String(20)  Out
            '        享受待遇标志    Small int   Out
            '        起付标准    Long    Out
            '        在职情况    String(1)   Out


            lngReturn = gobj成都内江.DoHospInFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(7), strArr(8), strOutPut(0), lngOutPut(0), lngOutPut(1), strOutPut(1))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0) & vbTab & lngOutPut(0) & vbTab & lngOutPut(1) & vbTab & strOutPut(1)
        Case 住院交易上传_内江
            '输入参数: 个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        统筹地区编码    String(1)   In
            '        医院交易流水号  String(20)  In
            '        出院带药类别    String(1)   In
            '        科别    String(10)  In
            '        医生    String(10)  In
            '        住院流水号  String(20)  In
            '        处方条数    String(2)   In
            '        处方明细    String处方条数×51  In
            '输出参数:  医保交易流水号  String(20)  Out
            '           处方明细    String处方条数×51  Out
            '           TRANSDETIAL输出 (计算费用明细) Out

            lngReturn = gobj成都内江.DoHospTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(7), strArr(8), strArr(9), strArr(10), strOutPut(0), strOutPut(1), strOutPut(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0) & vbTab & strOutPut(1) & vbTab & strOutPut(2)
        Case 住院产易上传取消_内江
            '输入参数:个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        住院流水号  String(20)  In
            '        医保交易流水号  String(20)  In
            '输出参数:

            lngReturn = gobj成都内江.DoHospCancelFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0) & vbTab & strOutPut(1) & vbTab & strOutPut(2)
        Case 出院登记上传_内江
            '输入参数:个人编号    String(8)   In
            '        社保卡号码  String(10)  In
            '        医院代码    String(5)   In
            '        操作员卡号码    String(10)  In
            '        统筹地区编码    String(1)   In
            '        出院日期    String(8)   In
            '        出院科别    String(10)  In
            '        出院诊治医生    String(10)  In
            '        诊断编码    String(20)  In
            '        出院带药    String(1)   In
            '        出院类别    String(1)   In
            '        住院流水号  String(20)  In
            '输出参数
            '        TRANSDETIAL输出 (计算费用明细)
            '        享受待遇标志    String(1)   Out
            '        医保内费用  String(10)  Out
            '        医保外费用  String(10)  Out
            '        基本医保支付 如果参加大病医保，则为大病医保支付  String(10)  Out
            '        高额医保支付    String(10)  Out
            '        公务员医疗补助  String(10)  Out
            '        个人按比例支付  String(10)  Out
            '        TRANSDETIAL结束
            '        起付标准    String(10)  Out
            '        个人帐户可用余额    String(10)  Out
            lngReturn = gobj成都内江.DoHospOutTransFunc(strArr(0), strArr(1), strArr(2), strArr(3), strArr(4), strArr(5), strArr(6), strArr(7), strArr(8), strArr(9), strArr(10), strArr(11), strOutPut(0), strOutPut(1), strOutPut(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0) & vbTab & strOutPut(1) & vbTab & strOutPut(2)
        Case 获取单位欠缴情况_内江
            '输入参数:个人编号    String (8)  IN
            '        社保卡号码  String (10) IN
            '        统筹地区编码    String (1)  IN
            '输出参数
            '        单位欠缴情况    String(1)   OUT
            lngReturn = gobj成都内江.GetArrearInfo(strArr(0), strArr(1), strArr(2), strOutPut(0))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = strOutPut(0)
        Case 初始化函数_内江
            '输入参数:ConfigFileName
            '        HostPort
            '        IPAddress
            lngReturn = gobj成都内江.SetCommPara(strArr(0), strArr(1), strArr(2))
            If lngReturn <> 0 Then
                 ShowMsgbox GetErrInfor(Rpad(lngReturn, 3, "0"))
                 Exit Function
            End If
            strReturn = ""
    End Select
    strOutPutstring = strReturn
    业务请求_成都内江 = True
    DebugTool "  输出参数为:" & strReturn
     Exit Function
ErrHand:
    DebugTool "业务请求失败  输出参数为:" & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

