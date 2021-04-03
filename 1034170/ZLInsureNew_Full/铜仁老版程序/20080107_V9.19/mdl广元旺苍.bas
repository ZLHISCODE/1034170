Attribute VB_Name = "mdl广元旺苍"
Option Explicit
Public Enum 业务类型_广元旺苍
    获得社保机构_旺苍 = 0
    获得参保人员资料_旺苍
    获取帐户余额_旺苍
    检查拔号连接_旺苍
    建立拔号连接_旺苍
    断开拔号连接_旺苍
    个人帐户消费_旺苍
    个人帐户消费_金额_旺苍
    消费冲正_旺苍
'    打印出院结算报表函数
'    打印住院人员财务结算单
'    获取住院记录号
    获取药品信息
End Enum
Private Type InitbaseInfor
    模拟数据 As Boolean                     '当前是否处于模拟读取医保接口数据
    医院编码 As String                      '初始医院编码
    机构编码 As String                      '默认的社保机构编码
    
End Type
Public InitInfor_广元旺苍 As InitbaseInfor

Private Type 病人身份
    医保卡号    As String
    医保证号    As String
    身份证号码  As String
    记录号      As String
    姓名        As String
    性别        As String
    出生日期    As String
    年龄        As Integer
    单位名称    As String
    机构编码    As String
    
    帐户余额    As String
    费用总额    As Double
    密码        As String
    社保中心    As Long
    病人id      As Long
End Type

Private Type 结算数据
    卡号 As String
    姓名    As String
    消费前帐户余额 As Double
    个人帐户支付金额 As Double
    自费金额 As Double
    消费后帐户余额 As Double
    交易时间  As String
    前端单据号  As String
    中心单据号  As String
    处方号  As String
    操作员姓名  As String
    前端名称  As String
    
    结帐id As Long
    结算标志 As Byte    '0-门诊,1-住院
End Type
Private g结算数据 As 结算数据
Public g病人身份_广元旺苍 As 病人身份
Public gcnOracle_广元旺苍 As ADODB.Connection     '中间库连接

Private gbln检查连接 As Boolean
Private gbln已经初始 As Boolean             '已经被初始化了.

'1.获得社保机构_旺苍编号和名称列表
Private Declare Function GetSBJGLB Lib "cdgk_Yb.dll" Alias "GETSBJGLB" () As String
'===============================================================================================================
'原型:FUNCTION GETSBJGLB:PCHAR
'功能: 获得社保机构_旺苍编号和名称列表
'入口参数: 无
'出口参数: 无
'返回: A社保机构编号+列间隔符+A社保机构名称+行间隔符+B社保机构编号+列间隔符+B社保机构名称+……
'===============================================================================================================

'2．获得参保人员的基本资料
Private Declare Function GETKZL Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'原型:FUNCTION GETKZL:PCHAR
'功能: 获得参保人员的基本资料
'入口参数:
'出口参数: 无
'返回: OK(或错误信息)@$医保卡号||医保证号||个人记录号||姓名||身份证号码||单位名称||性别||出生日期（格式：YYYY-MM-DD）
'===============================================================================================================

'3.个人帐户余额查询
Private Declare Function GETZHYE Lib "cdgk_Yb.dll" (ByVal str机构编码 As String, ByVal strPassWord As String) As String
'===============================================================================================================
'原型:FUNCTION GETZHYE(YBJGBH,CPASSWORD:PCHAR):PCHAR
'功能: 获得持卡人员个人帐户余额
'入口参数:YBJGBH  PCHAR   保险机构编号
'         CPASSWORD   PCHAR   持卡人卡密码
'出口参数: 无
'返回:  OK(或错误信息)@$个人帐户余额
'===============================================================================================================

'4.检测拔号连接是否连接成功
Private Declare Function CheckCon Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'原型:FUNCTION CHECKCON:PCHAR;
'功能:检测拔号连接是否连接成功
'入口参数:
'返回:OK或错误信息
'===============================================================================================================

'5.建立拔号连接
Private Declare Function RasDial Lib "cdgk_Yb.dll" (ByVal str机构代码 As String) As String
'===============================================================================================================
'原型:FUNCTION RASDIAL(SBXJGBH:PCHAR):PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:SBXJGBH PCHAR   保险机构编号
'返回:  成功    川大金键HIS拔号器状态栏显示"连接"
'       失败 错误信息
'===============================================================================================================



'6.断开与社保局的连接
Private Declare Function DisDial Lib "cdgk_Yb.dll" () As String
'===============================================================================================================
'原型:FUNCTION DISDIAL:PCHAR
'功能:拔号至选择的社保局，与其建立连接
'入口参数:
'返回:
'===============================================================================================================

'7.个人帐户销费
Private Declare Function GRZHXF_CF Lib "cdgk_Yb.dll" (ByVal str机构编号 As String, str处方号 As String, _
            ByVal str明细数据 As String, ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:Function GRZHXF_CF()(YBJGBH,CFH:PCHAR;CFMXDATA:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:进行个人帐户消费
'入口参数:YBJGBH  PCHAR   保险机构编号
'        CFH PCHAR   处方号
'        CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+        ……        处方N(医保药品编号+列间隔符+单价+列间隔符+数量
'        CPASSWORD   PCHAR   持卡人卡密码
'        CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================


'8.个人帐户消费（直接输入消费金额）

Private Declare Function GRZHXF_JE Lib "cdgk_Yb.dll" (ByVal str机构编号 As String, str处方号 As String, _
             ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:FUNCTION GRZHXF_JE(YBJGBH,XFJE:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:进行个人帐户消费
'入口参数:YBJGBH  PCHAR   保险机构编号
'    XFJE    PCHAR   消费金额(保证为小数，并且保留二位小数)
'    CPASSWORD   PCHAR   持卡人卡密码
'    CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================

'9.消费冲正

Private Declare Function XFCZ Lib "cdgk_Yb.dll" (ByVal str机构编号 As String, str中心单据号 As String, _
             ByVal strPassWord As String, ByVal str操作员 As String) As String
'===============================================================================================================
'原型:FUNCTION XFCZ(YBJGBH ，CZXDJH:PCHAR; CPASSWORD:PCHAR;CCZYXM:PCHAR):PCHAR
'功能:对已经消费的记录进行冲正。
'入口参数:YBJGBH  PCHAR   保险机构编号
'        cZXDJH  PCHAR   中心单据号(消费时返回)
'        CPASSWORD   PCHAR   持卡人卡密码
'        CCZYXM  PCHAR   操作员姓名
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
'===============================================================================================================












'12.打印出院结算报表函数
Private Declare Function JSReport Lib "cdgk_Yb.dll" (ByVal str开始住院号 As String, ByVal str结束住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION JSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL
'功能:打印社保机构提供的动态报表，目前德阳地区所用动态报表："住院结算统计表（补充）"、"个人住院结算单"和"住院结算统计表"三张报表。使用"21．提取基础资料"函数，自动更新本地报表。
'入口参数:
'    ASTARTZYH   PCHAR   打印开始住院号
'    AENDZYH PCHAR      打印结束住院号
'   注意:
'    1 ?二个住院号之间所有的住院记录必须为同一个社保局?
'    2、当只打印一个住院号的报表时，二个参数值一样。
'出口参数: 无
'返回:无须注意返回值
'===============================================================================================================

'13.打印住院人员财务结算单
Private Declare Function CWJSReport Lib "cdgk_Yb.dll" (ByVal str开始住院号 As String, ByVal str结束住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION CWJSREPORT(ASTARTZYH,AENDZYH:PCHAR):PCHAR;STDCALL;
'功能:打印住院人员财务结算单。
'入口参数:
'    ASTARTZYH   PCHAR   打印开始住院号
'    AENDZYH PCHAR      打印结束住院号
'   注意:
'    1 ?二个住院号之间所有的住院记录必须为同一个社保局?
'    2、当只打印一个住院号的报表时，二个参数值一样。
'出口参数: 无
'返回:无须注意返回值
'===============================================================================================================

'14.提取基础资料
Private Declare Function GetJCXX Lib "cdgk_Yb.dll" (ByVal str机构编号 As String, ByVal str下载标志 As String) As String
'===============================================================================================================
'原型:GETJCXX(SBXJGBH:PCHAR;DOWNALL:INTEGER):PCHAR
'功能:向指定的社保机构提取基础资料
'入口参数:
'    SBXJGBH PCHAR   保险机构编号
'    DOWNALL PCHAR   当值为0时表示下载本地医保数据库中没有的基础资料，为其他时表示全部重新下载
'出口参数: 无
'返回:'OK'或错误信息
'===============================================================================================================

'15 根据住院号得到住院记录号
Private Declare Function GetZYIDByZyBH Lib "cdgk_Yb.dll" (ByVal str住院号 As String) As String
'===============================================================================================================
'原型:FUNCTION GETZYIDBYZYBH(AZYH:PCHAR):PCHAR
'功能:根据住院号得到住院记录号
'入口参数:
'   AZYH    PCHAR   住院号'出口参数: 无
'返回:'OK'@$住院记录号或错误信息
'===============================================================================================================


'19 根据药品编号得到药品信息
Private Declare Function GetSINYPXX Lib "cdgk_Yb.dll" (ByVal str机构编码 As String, ByVal str药品编码 As String) As String
'===============================================================================================================
'原型:FUNCTION GETSINYPXX(SBXJGBH,CYPBH:PCHAR):PCHAR
'功能:根据药品编号得到药品信息
'入口参数:
'    SBXJGBH PCHAR   保险机构编号
'    CYPBH   PCHAR   药品编号
'返回:OK@$类别:药品||中文名称:阿莫西林钠（克拉维酸钾）||计量单位:支||单价上线:0||自费比例:20
'===============================================================================================================




Public Function 医保初始化_广元旺苍() As Boolean
    Dim strReg As String, strOutPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '初始模拟接口
    Call GetRegInFor(g公共模块, "操作", "模拟接口", strReg)
    If Val(strReg) = 1 Then
        InitInfor_广元旺苍.模拟数据 = True
    Else
        InitInfor_广元旺苍.模拟数据 = False
    End If
    
   Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
   
   InitInfor_广元旺苍.机构编码 = strReg
   g病人身份_广元旺苍.机构编码 = strReg
   
   If strReg = "" Then
        MsgBox "你未设置默认的社保机构编码，请检查参数设置!"
        Exit Function
   End If
   
    '取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "读取医院编码")
    InitInfor_广元旺苍.医院编码 = Nvl(rsTemp!医院编码)
    
    
    '中间库连接
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "渝北医保")
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "检查拨号连接"
                 gbln检查连接 = Nvl(rsTemp("参数值"), 0) = 1
        End Select
        rsTemp.MoveNext
    Loop
    
    Set gcnOracle_广元旺苍 = New ADODB.Connection
    If OraDataOpen(gcnOracle_广元旺苍, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到医保中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    
   '建立拔号连接
   If gbln已经初始 = False And gbln检查连接 Then
       If 建立拨号连接() = False Then Exit Function
   End If
   
   If gbln检查连接 Then
        '检查拔号连接
        If 业务请求_广元旺苍(检查拔号连接, "", strOutPut) = False Then
             Exit Function
        End If
    End If
    gbln已经初始 = True
    医保初始化_广元旺苍 = True
End Function

Public Function 医保终止_广元旺苍() As Boolean
    Dim strOutPut As String
    
    If gcnOracle_广元旺苍.State = 1 Then
        gcnOracle_广元旺苍.Close
    End If
    '建立拔号连接
   Call 业务请求_广元旺苍(断开拔号连接, "", strOutPut)
    Err = 0
    On Error Resume Next
    医保终止_广元旺苍 = True
End Function

Public Function 身份标识_广元旺苍(Optional bytType As Byte, Optional lng病人ID As Long) As String
    '功能：识别指定人员是否为参保病人，返回病人的信息
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    '返回：空或信息串
    Err = 0
    On Error GoTo ErrHand:
    If bytType = 0 Or bytType = 3 Then Exit Function
    
    身份标识_广元旺苍 = frmIdentify广元旺苍.GetPatient(bytType, lng病人ID)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    身份标识_广元旺苍 = ""
End Function


Public Function 个人余额_广元旺苍(ByVal lng病人ID As Long) As Currency
'功能: 提取参保病人个人帐户余额
'返回: 返回个人帐户余额
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select nvl(帐户余额,0) as 帐户余额 from 保险帐户 where 病人ID='" & lng病人ID & "' and 险类=" & TYPE_广元旺苍
    Call OpenRecordset(rsTemp, "读取个人帐户余额")
    
    If rsTemp.EOF Then
        个人余额_广元旺苍 = 0
    Else
        个人余额_广元旺苍 = rsTemp("帐户余额")
    End If
End Function
Public Function 门诊虚拟结算_广元旺苍(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
    '参数：rsDetail     费用明细(传入)
    '      cur结算方式  "报销方式;金额;是否允许修改|...."
    '字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    门诊虚拟结算_广元旺苍 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function 建立拨号连接() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Static str机构编号 As String
    Dim strInput As String, strOutPut As String
    建立拨号连接 = False
    
    Err = 0: On Error GoTo ErrHand:
    If str机构编号 <> g病人身份_广元旺苍.机构编码 Then
        '检查网络是否正常连接
        If str机构编号 = "" Then
            '表求第一次远行,需断开
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutPut) = False Then
                Exit Function
            End If
        Else
            '表示至少有两次以上的操作,则需断开连接
            Call 业务请求_广元旺苍(断开拔号连接_旺苍, "", strOutPut)
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutPut) = False Then Exit Function
        End If
        If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutPut) = False Then Exit Function
    Else
        If 业务请求_广元旺苍(检查拔号连接_旺苍, "", strOutPut) = False Then
            '需重新建立拨号连接
            If 业务请求_广元旺苍(建立拔号连接_旺苍, g病人身份_广元旺苍.机构编码, strOutPut) = False Then
                Exit Function
            End If
        End If
    End If
    str机构编号 = g病人身份_广元旺苍.机构编码
    建立拨号连接 = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then Resume
End Function
Public Function 门诊结算_广元旺苍(lng结帐ID As Long, cur个人帐户 As Currency, str医保号 As String, cur全自付 As Currency) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur支付金额   从个人帐户中支出的金额
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
        '此时所有收费细目必然有对应的医保编码
    Dim strInput As String, strOutPut As String
    Dim lng病人ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strArr As Variant
    If 建立拨号连接() = False Then Exit Function
    
    On Error GoTo errHandle
    
    Call DebugTool("进入门诊结算")
    
    gstrSQL = "" & _
        "   Select a.*,a.*,a.付数*a.数次 as 数量,a.实收金额/(nvl(a.付数,1)*nvl(a.数次,1)) as 单价 " & _
        "   From 病人费用记录 a " & _
        "   Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
        
    Call OpenRecordset(rs明细, "获取明细记录")
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If

    lng病人ID = rs明细("病人ID")
    
    If g病人身份_广元旺苍.病人id <> lng病人ID Then
        MsgBox "该病人还没有经过身份验证，不能进行医保结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    g结算数据.结帐id = lng结帐ID
    g结算数据.结算标志 = 0
    '写入明细
    If 门诊明细写入(rs明细, False) = False Then Exit Function
    
    '显示其结处方式
    If 结算方式更正() = False Then
        Exit Function
    End If
    
    
    
    Dim dbl个人帐户 As Double
    dbl个人帐户 = 获取个人帐户支付()
    If dbl个人帐户 <> g结算数据.个人帐户支付金额 Then
        '更新个人帐户支付
        '入:YBJGBH  PCHAR   保险机构编号
        '    XFJE    PCHAR   消费金额(保证为小数，并且保留二位小数)
        '    CPASSWORD   PCHAR   持卡人卡密码
        '    CCZYXM  PCHAR   操作员姓名
        '返回:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
        strInput = g病人身份_广元旺苍.机构编码
        strInput = strInput & vbTab & Format(dbl个人帐户, "###0.00;-###0.00;0.00;0.00")
        strInput = strInput & vbTab & g病人身份_广元旺苍.密码
        strInput = strInput & vbTab & gstrUserName
        If 业务请求_广元旺苍(个人帐户消费_金额_旺苍, strInput, strOutPut) = False Then Exit Function
        If strOutPut = "" Then Exit Function
        strArr = Split(strOutPut, "||")
        
        With g结算数据
            .卡号 = strArr(0)
            .姓名 = strArr(1)
            .消费前帐户余额 = Val(strArr(2))
            .个人帐户支付金额 = Val(strArr(3))
            .自费金额 = Val(strArr(4))
            .消费后帐户余额 = Val(strArr(5))
            .交易时间 = strArr(6)
            .前端单据号 = strArr(7)
            .中心单据号 = strArr(8)
            .处方号 = strArr(9)
            .操作员姓名 = strArr(10)
            .前端名称 = strArr(11)
        End With
    End If
       
    '填写结算表
    Call DebugTool("填写结算记录")
    

    
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(消费前帐户余额),累计统筹报销_IN(消费后帐户余额),住院次数_IN(无),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(自费金额),
    '   进入统筹金额_IN(无),统筹报销金额_IN(无),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(中心单据号),主页ID_IN(无),中途结帐_IN,备注_IN(前端单据号|处方号|操作员姓名|前端名称)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & g结算数据.消费前帐户余额 & "," & g结算数据.消费后帐户余额 & ",null,0,0,0," & _
            g病人身份_广元旺苍.费用总额 & ",0," & g结算数据.自费金额 & "," & _
          "0,0,0,0," & g结算数据.个人帐户支付金额 & ",'" & _
            g结算数据.中心单据号 & " ',NULL,NULL,'" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
            
    Call ExecuteProcedure("保存结算记录")
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    门诊结算_广元旺苍 = True
    Exit Function

Err反结算:

'入口参数:YBJGBH  PCHAR   保险机构编号
'        cZXDJH  PCHAR   中心单据号(消费时返回)
'        CPASSWORD   PCHAR   持卡人卡密码
'        CCZYXM  PCHAR   操作员姓名
    strInput = g病人身份_广元旺苍.机构编码
    strInput = strInput & vbTab & g结算数据.中心单据号
    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
    strInput = strInput & vbTab & gstrUserName
    
    If 业务请求_广元旺苍(消费冲正_旺苍, strInput, strOutPut) = False Then Exit Function
'返回:个人帐户消费信息(OK@$个人帐户消费信息)
'   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
    If strOutPut = "" Then Exit Function
     strArr = Split(strOutPut, "||")
    
    With g结算数据
        .卡号 = strArr(0)
        .姓名 = strArr(1)
        .消费前帐户余额 = Val(strArr(2))
        .个人帐户支付金额 = Val(strArr(3))
        .自费金额 = Val(strArr(4))
        .消费后帐户余额 = Val(strArr(5))
        .交易时间 = strArr(6)
        .前端单据号 = strArr(7)
        .中心单据号 = strArr(8)
        .处方号 = strArr(9)
        .操作员姓名 = strArr(10)
        .前端名称 = strArr(11)
    End With

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function 门诊结算冲销_广元旺苍(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
    '功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
    '参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
    '      cur个人帐户   从个人帐户中支出的金额
    
    Dim intMouse As Integer
    Dim lng冲销ID  As Long
    Dim rs明细 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    Dim strArr As Variant
    
    门诊结算冲销_广元旺苍 = False
    
    '身份验证
    intMouse = Screen.MousePointer
    Screen.MousePointer = 1
    If 身份标识_大连(0, lng病人ID) = "" Then
        Screen.MousePointer = intMouse
        Exit Function
    End If
    Screen.MousePointer = intMouse
    
    gstrSQL = "select distinct A.结帐ID from 病人费用记录 A,病人费用记录 B " & _
              " where A.NO=B.NO and A.记录性质=B.记录性质 and A.记录状态=2 and B.结帐ID=" & lng结帐ID
    Call OpenRecordset(rsTemp, "重庆医保")
    lng冲销ID = rsTemp("结帐ID")
    
    
    
    gstrSQL = "Select * From 病人费用记录 " & _
        " Where 结帐ID=" & lng冲销ID & " And Nvl(附加标志,0)<>9 And Nvl(记录状态,0)<>0"
        
    Call OpenRecordset(rs明细, "获取冲销记录")
    g病人身份_广元旺苍.费用总额 = 0
    With rs明细
        Do While Not .EOF
                '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            ExecuteProcedure "打上上传标志"
            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!实收金额, 0)
            .MoveNext
        Loop
    End With
    
    '冲正:
    gstrSQL = "Select 支付顺序号 from 保险结算记录 where 性质=1 and 记录id=" & lng结帐ID
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取中心单据号"
    If rsTemp.EOF Then
        ShowMsgbox "不存在结算记录,不能冲销!"
        Exit Function
    End If
    
    '入口参数:YBJGBH  PCHAR   保险机构编号
    '        cZXDJH  PCHAR   中心单据号(消费时返回)
    '        CPASSWORD   PCHAR   持卡人卡密码
    '        CCZYXM  PCHAR   操作员姓名
    strInput = g病人身份_广元旺苍.机构编码
    strInput = strInput & vbTab & Nvl(rsTemp!支付顺序号)
    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
    strInput = strInput & vbTab & gstrUserName
    
    If 业务请求_广元旺苍(消费冲正_旺苍, strInput, strOutPut) = False Then Exit Function
    '返回:个人帐户消费信息(OK@$个人帐户消费信息)
    '   格式:卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
    If strOutPut = "" Then Exit Function
     strArr = Split(strOutPut, "||")
    
    With g结算数据
        .卡号 = strArr(0)
        .姓名 = strArr(1)
        .消费前帐户余额 = Val(strArr(2))
        .个人帐户支付金额 = Val(strArr(3))
        .自费金额 = Val(strArr(4))
        .消费后帐户余额 = Val(strArr(5))
        .交易时间 = strArr(6)
        .前端单据号 = strArr(7)
        .中心单据号 = strArr(8)
        .处方号 = strArr(9)
        .操作员姓名 = strArr(10)
        .前端名称 = strArr(11)
    End With
    门诊结算冲销_广元旺苍 = True
        
   '插入保险结算记录
    '原过程参数:
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN,帐户累计支出_IN,累计进入统筹_IN,累计统筹报销_IN,住院次数_IN,起付线_IN,封顶线_IN,实际起付线_IN,
    '   发生费用金额_IN,全自付金额_IN,首先自付金额_IN,
    '   进入统筹金额_IN,统筹报销金额_IN,大病自付金额_IN,超限自付金额_IN,个人帐户支付_IN,"
    '   支付顺序号_IN,主页ID_IN,中途结帐_IN,备注_IN
    
    '新值代表
    '   性质_IN  ,记录ID_IN,险类_IN,病人ID_IN,年度_IN," & _
    "   帐户累计增加_IN(),帐户累计支出_IN(),累计进入统筹_IN(消费前帐户余额),累计统筹报销_IN(消费后帐户余额),住院次数_IN(无),起付线(无),封顶线_IN(无),实际起付线_IN(无),
    '   发生费用金额_IN(费用总额),全自付金额_IN(无),首先自付金额_IN(自费金额),
    '   进入统筹金额_IN(无),统筹报销金额_IN(无),大病自付金额_IN(无),超限自付金额_IN(无),个人帐户支付_IN(个人帐户支付金额),"
    '   支付顺序号_IN(中心单据号),主页ID_IN(无),中途结帐_IN,备注_IN(前端单据号|处方号|操作员姓名|前端名称)
    DebugTool "结算交易提交成功,并开始保存保险结算记录"
   
    gstrSQL = "zl_保险结算记录_insert(1," & lng冲销ID & "," & gintInsure & "," & lng病人ID & "," & Year(zlDatabase.Currentdate) & "," & _
             "NULL,NULL," & -1 * g结算数据.消费前帐户余额 & "," & -1 * g结算数据.消费后帐户余额 & ",null,0,0,0," & _
           -1 * g病人身份_广元旺苍.费用总额 & ",0," & -1 * g结算数据.自费金额 & "," & _
          "0,0,0,0," & -1 * g结算数据.个人帐户支付金额 & ",'" & _
            g结算数据.中心单据号 & " ',NULL,NULL,'" & g结算数据.前端单据号 & "|" & g结算数据.处方号 & "|" & g结算数据.操作员姓名 & "|" & g结算数据.前端单据号 & "')"
            
    Call ExecuteProcedure("保存结算记录")
    '---------------------------------------------------------------------------------------------
    门诊结算冲销_广元旺苍 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_广元旺苍(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
    '功能：将入院登记信息发送医保前置服务器确认；
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    '将病人的状态进行修改
    ShowMsgbox "本医保接口不支付住院部分"
    
    入院登记_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    入院登记_广元旺苍 = False
End Function
Private Function Get交易代码(ByVal intType As 业务类型_广元旺苍, Optional bln读名称 As Boolean = False) As String
    '代码暂没用
    Select Case intType
        Case 获得社保机构_旺苍
            Get交易代码 = IIf(bln读名称, "获得社保机构", "01")
        Case 获得参保人员资料_旺苍
            Get交易代码 = IIf(bln读名称, "获得参保人员资料", "02")
        Case 获取帐户余额_旺苍
                Get交易代码 = IIf(bln读名称, "获取帐户余额", "03")
        Case 检查拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "检查拔号连接", "04")
        Case 建立拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "建立拔号连接", "05")
        Case 断开拔号连接_旺苍
            Get交易代码 = IIf(bln读名称, "断开拔号连接", "06")
        Case 个人帐户消费_旺苍
            Get交易代码 = IIf(bln读名称, "个人帐户消费", "07")
        Case 个人帐户消费_金额_旺苍
            Get交易代码 = IIf(bln读名称, "个人帐户消费_金额", "08")
        Case 消费冲正_旺苍
            Get交易代码 = IIf(bln读名称, "消费冲正", "09")
        Case Else
            Get交易代码 = IIf(bln读名称, "错误的交易代码", "-1")
    End Select
End Function
Public Function 业务请求_广元旺苍(ByVal intType As 业务类型_广元旺苍, strInputString As String, strOutPutstring As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:对所有业务进行业务请求
    '--入参数:strinPutString-输入串,按参数顺序,以tab键分隔的传入串
    '--出参数:strOutPutString-输出串,按参数顺序,以tab键分隔的返回串
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim strInput As String, lngReturn As Long, strOutPut As String, strReturn As String
    Dim strInValue(0 To 20) As String
    
    Dim str交易代码 As String
    Dim i As Integer
    Dim strArr
    
    str交易代码 = Get交易代码(intType, True)
    strInput = strInputString
    DebugTool "进入业务请求函数(业务类型代码为:" & intType & " 业务名称：" & str交易代码 & ")" & vbCrLf & "        输入参数为:" & strInputString
    
    业务请求_广元旺苍 = False
    If InitInfor_广元旺苍.模拟数据 Then
        '读取模拟数据
        Read模拟数据 intType, strInput, strOutPutstring
         业务请求_广元旺苍 = True
        Exit Function
    End If
    strArr = Split(strInputString, vbTab)
    For i = 0 To UBound(strArr)
        strInValue(i) = strArr(i)
    Next
    
    Err = 0
    On Error GoTo ErrHand:
    
    Select Case intType
        Case 获得社保机构_旺苍
            strOutPut = GetSBJGLB()
            
            If strOutPut = "" Then
                MsgBox "获取社保机构时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case 获得参保人员资料_旺苍
            strOutPut = GETKZL()
            If strOutPut = "" Then
                MsgBox "获得参保人员资料_旺苍时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case 获取帐户余额_旺苍
            strOutPut = GETZHYE(strInValue(0), strInValue(1))
            ''OK'+行间隔符+个人帐户余额
            If strOutPut = "" Then
                MsgBox "获取帐户余额_时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        
        Case 检查拔号连接_旺苍
            strOutPut = CheckCon()
            If strOutPut = "" Then
                MsgBox "检查拔号连接时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case 建立拔号连接_旺苍
            strOutPut = RasDial(strInValue(0))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case 断开拔号连接_旺苍
            strOutPut = DisDial()
            strOutPut = ""
        Case 个人帐户消费_旺苍
            strOutPut = GRZHXF_CF(strInValue(0), strInValue(1), strInValue(2), strInValue(3), strInValue(4))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutPut = strArr(1)
        Case 个人帐户消费_金额_旺苍
            strOutPut = GRZHXF_JE(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutPut = strArr(1)
        Case 消费冲正_旺苍
            strOutPut = XFCZ(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
            strOutPut = strArr(1)
        
'        Case 入院登记
'            '
'            strOutput = RYDJ(strInValue(0), Replace(strInValue(1), vbTab & "|", "||"), strInValue(2))
'            If strOutput = "" Then
'                MsgBox "入院登记时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case 取消入院登记
'            strOutput = ZYQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "取消入院登记时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 出院登记
'            strOutput = CYCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "出院登记时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 取消出院登记
'            strOutput = CYCSQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "取消出院登记时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 增加处方单据
'            strOutput = AddCFJL(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
'            If strOutput = "" Then
'                MsgBox "增加处方单据时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case 增加处方明细
'            strOutput = AddCFMX(strInValue(0), strInValue(1), strInValue(2), strInValue(3))
'            If strOutput = "" Then
'                MsgBox "增加处方明细时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'            For i = 1 To UBound(strArr)
'                strOutput = "||" & strArr(i)
'            Next
'            If strOutput <> "" Then
'                strOutput = Mid(strOutput, 3)
'            End If
'        Case 删除处方单据及其明细
'            strOutput = DELCFJL(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "删除处方单据及其明细时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 单条处方传输
'            strOutput = CFCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "单条处方传输时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 出院结算
'            strOutput = CFCS(strInValue(0), strInValue(1))
'            If strOutput = "" Then
'                MsgBox "出院结算时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = strArr(1)
'        Case 取消出院结算
'            strOutput = CYJSQX(strInValue(0))
'            If strOutput = "" Then
'                MsgBox "出院结算时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutput, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutput = ""
'        Case 打印出院结算报表函数
'            strOutput = JSReport(strInValue(0), strInValue(1))
'            strOutput = ""
'        Case 打印住院人员财务结算单
'            strOutput = CWJSReport(strInValue(0), strInValue(1))
'            strOutput = ""
        
        Case 重提人员基本资料
            '与上面重复了"打印住院人员财务结算单"
'            strOutPut = CWJSREPORT(strInValue(0))
'              If strOutPut = "" Then
'                MsgBox "重提人员基本资料时,返回了空值。", vbInformation, gstrSysName
'                Exit Function
'            End If
'            strArr = Split(strOutPut, "@$")
'            If strArr(0) <> "OK" Then
'                MsgBox strArr(0), vbInformation, gstrSysName
'                Exit Function
'            End If
'            strOutPut = ""
        Case 提取基础资料
        
            strOutPut = GetJCXX(strInValue(0), strInValue(1))
              If strOutPut = "" Then
                MsgBox "提取基础资料时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = ""
        Case 获取住院记录号
            strOutPut = GetZYIDByZyBH(strInValue(0))
            If strOutPut = "" Then
                MsgBox "获取住院记录号时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
        Case 获取药品信息
             strOutPut = GetSINYPXX(strInValue(0), strInValue(1))
            If strOutPut = "" Then
                MsgBox "获取药品信息时,返回了空值。", vbInformation, gstrSysName
                Exit Function
            End If
            strArr = Split(strOutPut, "@$")
            If strArr(0) <> "OK" Then
                MsgBox strArr(0), vbInformation, gstrSysName
                Exit Function
            End If
            strOutPut = strArr(1)
    End Select
    strOutPutstring = strOutPut
    业务请求_广元旺苍 = True
    DebugTool "业务请求成功(业务类型为:" & intType & ")." & vbCrLf & "输入参数为" & vbCrLf & strInputString & vbCrLf & "输出参数为:" & vbCrLf & strReturn
     Exit Function
    
ErrHand:
    DebugTool "业务请求失败(业务类型为:" & intType & ")." & vbCrLf & "输入参数为" & vbCrLf & strInputString & vbCrLf & "输出参数为:" & vbCrLf & strReturn
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 入院登记撤销_广元旺苍(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认（如果没发生费用，则调入院登记撤销接口）
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
            
    '刘兴宏:20040923增加的
    入院登记撤销_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function 出院登记_广元旺苍(lng病人ID As Long, lng主页ID As Long) As Boolean
    '功能：将出院信息发送医保前置服务器确认；由于只针对撤消出院的病人，因此这个流程相对简单
    '参数：lng病人ID-病人ID；lng主页ID-主页ID
    '返回：交易成功返回true；否则，返回false
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    
    Err = 0
    On Error GoTo ErrHand:
    
    出院登记_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    出院登记_广元旺苍 = False
End Function
Public Function 出院登记撤销_广元旺苍(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '出院登记撤消
    出院登记撤销_广元旺苍 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_广元旺苍(lng结帐ID As Long, ByVal lng病人ID As Long) As Boolean
  '功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
    '参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
    '      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    '      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    住院结算_广元旺苍 = True
    Exit Function
End Function
Public Function 住院结算冲销_广元旺苍(lng结帐ID As Long) As Boolean
     '----------------------------------------------------------------
    '功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
    '参数：lng结帐ID-需要作废的结帐单ID号；
    '返回：交易成功返回true；否则，返回false
    '注意：1)主要使用结帐恢复交易和费用删除交易；
    '      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，在病人费用记录中根据结帐ID查找；
    '      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
    '----------------------------------------------------------------
    住院结算冲销_广元旺苍 = True
End Function
Public Function 处方登记_广元旺苍(ByVal lng记录性质 As Long, ByVal lng记录状态 As Long, ByVal str单据号 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传处理明细数据
    '--入参数:
    '--出参数:
    '--返  回:上传成功返回True,否则False
    '-----------------------------------------------------------------------------------------------------------


    处方登记_广元旺苍 = True
End Function

Private Function Read模拟数据(ByVal int业务类型 As 业务类型_广元旺苍, ByVal strInputString As String, ByRef strOutPutstring As String)
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
'    If InStr(1, strOutPutstring, "@$") <> 0 Then
'        strOutPutstring = Split(strOutPutstring, "@$")(1)
'    End If
    Exit Function
ErrHand:
    DebugTool Err.Description
    Exit Function
End Function
Private Sub OpenRecordset_广元旺苍(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSQL, strSql))
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle_广元旺苍, adOpenStatic, adLockReadOnly
    Call SQLTest
End Sub

Public Function 住院虚拟结算_广元旺苍(rsExse As Recordset, ByVal lng病人ID As Long, Optional bln结帐处 As Boolean = True) As String
    
    'rsExse:字符集
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    ShowMsgbox "本医保接口不支付住院部分"
    Exit Function
End Function
Public Function 医保设置_广元旺苍(ByVal lng险类 As Long, ByVal lng医保中心 As Integer) As Boolean
    医保设置_广元旺苍 = frmSet广元旺苍.参数设置
End Function
Public Sub ExecuteProcedure_广元旺苍(ByVal strCaption As String)
'功能：执行SQL语句
    Call SQLTest(App.ProductName, strCaption, gstrSQL)
    gcnOracle_广元旺苍.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
End Sub

Private Function 门诊明细写入(ByVal rs明细 As ADODB.Recordset, Optional ByVal bln虚拟 As Boolean = True) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:上传明细记录
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strInput As String, strOutPut As String
    Dim str明细 As String
    
    Dim strArr
    
    门诊明细写入 = False
    g病人身份_黔南.费用总额 = 0
    
    Err = 0:    On Error GoTo ErrHand:
    '然后插入处方明细
    With rs明细
        If .RecordCount = 0 Then
            ShowMsgbox "不存在相关的明细费用记录!"
            Exit Function
        End If
        'YBJGBH  PCHAR   保险机构编号
        'CFH PCHAR   处方号
        'CFMXDATA    PCHAR   处方明细数据    格式说明：处方1(医保药品编号+列间隔符+单价+列间隔符数量+)+行间隔符+
        'CPASSWORD   PCHAR   持卡人卡密码
        'CCZYXM  PCHAR   操作员姓名
        strInput = g病人身份_广元旺苍.机构编码
        strInput = strInput & vbTab & Nvl(!no)
        
        Do While Not rs明细.EOF
            gstrSQL = "Select * From 医保支付项目 where 险类=" & gintInsure & " and 中心=" & g病人身份_广元旺苍.社保中心 & " and 收费细目id=" & Nvl(!收费细目ID, 0)
            Call OpenRecordset(rsTemp, "确定医保支付项目")
            If rsTemp.EOF Then
                gstrSQL = "Select * From 收费细目 where id=" & Nvl(!收费细目ID, 0)
                If rsTemp.EOF Then
                    ShowMsgbox "不存在相关的收费项目!"
                Else
                    ShowMsgbox "在收费项目中，项目为:" & rsTemp!名称 & "未进行相关对码!"
                End If
                Exit Function
            End If
            If Val(Nvl(rs明细("实收金额"), 0)) <> 0 Then
                str明细 = str明细 & "@$" & Nvl(rsTemp!项目编码)
                str明细 = str明细 & "||" & Nvl(rsTemp!项目编码)
                str明细 = str明细 & "||" & Nvl(rsTemp!单价, 0)
                str明细 = str明细 & "||" & Nvl(rsTemp!数量, 0)
            End If
            g病人身份_广元旺苍.费用总额 = g病人身份_广元旺苍.费用总额 + Nvl(!实收金额, 0)
            
            '写上传标志
            gstrSQL = "ZL_病人费用记录_更新医保(" & Nvl(!ID, 0) & ",NULL,NULL,NULL,NULL,1,null)"
            ExecuteProcedure "打上上传标志"
            
            rs明细.MoveNext
        Loop
    End With
    str明细 = Mid(str明细, 3)
    strInput = strInput & vbTab & str明细
    strInput = strInput & vbTab & g病人身份_广元旺苍.密码
    strInput = strInput & vbTab & gstrUserName
    
    If 业务请求_广元旺苍(个人帐户消费_旺苍, strInput, strOutPut) = False Then Exit Function
    If strOutPut = "" Then Exit Function
    strArr = Split(strOutPut, "||")
    
    With g结算数据
        .卡号 = strArr(0)
        .姓名 = strArr(1)
        .消费前帐户余额 = Val(strArr(2))
        .个人帐户支付金额 = Val(strArr(3))
        .自费金额 = Val(strArr(4))
        .消费后帐户余额 = Val(strArr(5))
        .交易时间 = strArr(6)
        .前端单据号 = strArr(7)
        .中心单据号 = strArr(8)
        .处方号 = strArr(9)
        .操作员姓名 = strArr(10)
        .前端名称 = strArr(11)
    End With
    门诊明细写入 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
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
    
    Err = 0:    On Error GoTo ErrHand:
    DebugTool "进入(" & "Get结算方式" & ")"
    
    '卡号||姓名||消费前帐户余额||个人帐户支付金额||自费金额||消费后帐户余额||交易时间||前端单据号||中心单据号||处方号||操作员姓名||前端名称
    dbl费用总额 = g结算数据.个人帐户支付金额 + g结算数据.自费金额
    str结算方式 = "||个人帐户|" & g结算数据.个人帐户支付金额
    
    If Format(g病人身份_广元旺苍.费用总额, "###0.00;-###0.00;0;0") <> Format(dbl费用总额, "###0.00;-###0.00;0;0") Then
        Dim blnYes As Boolean
        '费用总额与医保中心返回总额不致,不能进行结算
        ShowMsgbox "本次结算总额(" & g病人身份_广元旺苍.费用总额 & ") 与" & vbCrLf & _
                    "   中心返回的总额(" & dbl费用总额 & ")不致产能结算?"
        Exit Function
    End If
    
   '如果存在,则保存冲预交记录中
    If str结算方式 <> "" Then
        str结算方式 = Mid(str结算方式, 3)
        g病人身份_成都内江.结算方式 = str结算方式
        
        If g结算数据.结算标志 = 0 Then
            gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐id & ",'" & str结算方式 & "', 0)"
            Call ExecuteProcedure("更新预交记录")
        Else
                gstrSQL = "zl_病人结算记录_Update(" & g结算数据.结帐id & ",'" & str结算方式 & "',1)"
                Call ExecuteProcedure("更新预交记录")
        End If
    End If
    
    '显示结算信息
    If frm结算信息.ShowME(g结算数据.结帐id, False, "个人帐户:" & g结算数据.个人帐户支付金额, IIf(g结算数据.结算标志 = 0, 0, 1)) = False Then
        结算方式更正 = False
        Exit Function
    End If
    结算方式更正 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function 获取个人帐户支付() As Double
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取个人帐户值(从预交记录中获取)
    '--入参数:
    '--出参数:
    '--返  回:成功,返回本次个人帐户支付,否则返回0
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 金额 From 病人预交记录 where 结帐ID=" & g结算数据.结帐id & " and  结算方式='个人帐户'"
    
    OpenRecordset rsTemp, "获取个人帐户支付"
    If Not rsTemp.EOF Then
        获取个人帐户支付 = Nvl(rsTemp!金额, 0)
    End If
    
End Function
