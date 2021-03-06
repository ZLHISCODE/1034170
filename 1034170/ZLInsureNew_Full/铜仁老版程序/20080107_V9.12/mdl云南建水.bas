Attribute VB_Name = "mdl云南建水"
Option Explicit
'仅用于云南医保的内部门诊变量
Private mstr顺序号 As String        '存放顺序号,仅用于门诊,住院存放于保险帐户中
Private mstr医保号 As String        '存放医保号,仅用于门诊
Private mcur帐户余额 As Double      '存放个人帐户余额,如果要用,仅用于门诊(身份验证返回)

Private mlng病人ID As Long          '存放病人ID，仅用于特殊门诊
Private mstr明细事务号 As String    '存放事务控制号，仅用于处理门诊费用明细撤消

Private mstrErr As String * 4

'###医保接口函数原型，需要改写为API方式
'以下几点需注意：
'（1）字符串参数不论传入还是传出，都加上ByVal关键字；
'（2）传出的字符串参数在调用前必须初始化；
'（3）数值参数对于传入的情况是要加上ByVal关键字的，但传出的一定不能加
'（4）对于浮点参数，对应类型是Double
'（5）千万别入结构的域

'====================================================================================
'1 费用明细传递
'输入：顺序号（就诊登记号）、数据批号、收费大类编码、收费项目编码、项目名称、数量、价格（单价）、产地、规格、用法用量、经办人、科室名称、事务控制号、医生姓名；
'输出：自付比例、自付金额、允许报销金额，错误代码；

Private Declare Sub yh_feedetailtrans Lib "Hisint" Alias "int_feedetailtrans" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Charge_Category As String, _
    ByVal Charge_Item As String, ByVal Charge_Name As String, ByVal Count As Double, ByVal Price As Double, ByVal Pr_Area As String, _
    ByVal Standard As String, ByVal Usage_Dosage As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal Transaction_No As String, _
    ByVal Doctor_Name As String, Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'2 费用结算
'见下面

'3、费用明细更改（备注，可用来完成退费操作）
'见下面

'4 入院登记
'输入：卡介质类型、医院编码、经办人、科室名称、病历号、住院号、是否特种病、特种病编码、入院时间、入院诊断、事务控制号；
'输出：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、初始化机构名称、错误编码；
'注：特种病编码可以为空
Private Declare Sub yh_admit Lib "Hisint" Alias "int_admit" _
    (ByVal card_mode As String, ByVal Hospial_No As String, ByVal Arranger As String, ByVal Section_Name As String, _
    ByVal anamnesis_No As String, ByVal Admit_No As String, ByVal Ifspecialsick As String, ByVal specialsick_no As String, _
    ByVal admit_time As String, ByVal admit_diagnose As String, ByVal Transaction_No As String, ByVal Serial_No As String, ByVal card_no As String, _
    ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal initinstitution As String, ByVal error_code As String)
    
'5 IC卡支付
'输入：卡介质类型、顺序号（就诊登记号）、经办人、支付原因,支付金额；
'输出：初始化机构名称、错误代码；
Private Declare Sub yh_cardpay Lib "Hisint" Alias "int_cardpay" _
    (ByVal card_mode As String, ByVal Serial_No As String, ByVal Arranger As String, ByVal Pay_reason As String, ByVal Pay_amount As Double, _
     ByVal initinstitution As String, ByVal error_code As String)


'6 虚拟结算
'见下面

'7 门诊身份识别
'输入：卡介质类型、医院编码、经办人、科室名称、病历号、门诊号、就医时间；
'输出：顺序号、卡号、个人编码、姓名、性别、出生日期、身份证号、初始化机构名称、卡余额、错误编码；
Private Declare Sub yh_outpatientidentify Lib "Hisint" Alias "int_outpatientidentify" _
    (ByVal card_mode As String, ByVal Hospital_No As String, ByVal Arranger As String, ByVal Section_No As String, _
    ByVal anamnesis_No As String, ByVal outpatient_No As String, ByVal hospitalize_time As String, ByVal Serial_No As String, _
    ByVal card_no As String, ByVal Personal_No As String, ByVal Name As String, ByVal Sex As String, ByVal birthdate As String, _
    ByVal Identify As String, ByVal initinstitution As String, accountremain As Double, ByVal error_code As String)

'8 IC卡基本信息查询
'输入：卡介质类型；
'输出: 余额、卡号、姓名、性别、身份证号、年龄、错误代码
Private Declare Sub yh_cardinfo Lib "Hisint" Alias "int_cardinfo" _
    (ByVal Code_Mode As String, Amount As Double, ByVal card_no As String, ByVal Name As String, _
    ByVal Sex As String, ByVal Identify As String, age As Double, ByVal error_code As String)

'9 密码更改
'输入: 卡介质类型
'输出: 错误代码
Private Declare Sub yh_changepassword Lib "Hisint" Alias "int_changepassword" _
    (ByVal Code_Mode As String, ByVal error_code As String)

'10    个人帐户支出查询
'输入：顺序号；
'输出：已支付总额，错误代码
Private Declare Sub yh_accountpay Lib "Hisint" Alias "int_accountpay" _
    (ByVal Serial_No As String, Amount As Double, ByVal error_code As String)

'11    门诊帐户支付
'输入：卡介质类型、医院编码、科室名称、经办人、支付原因、费用总额、帐户支付额；
'输出：初始化机构名称、顺序号、错误代码；
Private Declare Sub yh_outpay Lib "Hisint" Alias "outpay" _
    (ByVal card_mode As String, ByVal Hospital_No As String, ByVal Section_No As String, ByVal Arranger As String, ByVal payreason As String, _
    ByVal Amount As Double, ByVal accountpay As Double, ByVal initinstitution As String, ByVal Serial_No As String, ByVal error_code As String)

'12    初始化
'输入: 无
'输出: 错误代码
Private Declare Sub yh_init Lib "Hisint" Alias "init" _
    (ByVal Errcode As String)

'13    断开连接
'输入：无
'输出: 无
'Public Declare Sub yh_quit Lib "Hisint" Alias "quit" ()    '在云南医保中已经声明

'14 IC卡圈存
'输入：无
'输出: 错误代码
Private Declare Sub yh_loadcard Lib "Hisint" Alias "int_loadcard" (ByVal error_code As String)
    
'15 数据传输
'输入：无
'输出: 错误代码
Private Declare Sub yh_datatrans Lib "Hisint" Alias "int_datatrans" (ByVal error_code As String)


'16 事务控制
'输入：交易类别，就诊顺序号，事务控制号，事务控制类型；
'输出: 错误代码
Private Declare Sub yh_transaction Lib "Hisint" Alias "int_transaction" _
    (ByVal Trade_Sort As String, ByVal Serial_No As String, ByVal Transaction_No As String, ByVal Affirm_Mode As String, ByVal error_code As String)

'17 获取事务控制号
'输入：无；
'输出: 事务控制号
Private Declare Sub yh_gettranssequence Lib "Hisint" Alias "int_gettranssequence" (ByVal Transaction_No As String)

'18    待遇变更分段费用查询
'输入参数：顺序号；
'输出参数：分段标准、分段序号、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、专项补助款支付额、错误代码；
Private Declare Sub yh_SubsecFee Lib "Hisint" Alias "int_SubsecFee" _
    (ByVal Serial_No As String, ByVal Standard_Subsec As String, ByVal Subsec_No As String, _
      Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, _
      Basepay As Double, outpay As Double, Preqpay As Double, Preqselfpay As Double, _
      SubsidyPay As Double, ByVal error_code As String)

'19 退费处理
'输入参数：顺序号，回退标志，结算编号，事务控制号；
'输出参数: 错误码
Private Declare Sub yh_recedefeebalance Lib "Hisint" Alias "int_recedefeebalance" _
    (ByVal Serial_No As String, ByVal return_flag As String, ByVal balance_no As String, ByVal Transaction_No As String, _
        ByVal error_code As String)

'删除所有未执行结算或预结算前的费用明细。如果数据只是做了虚拟结算，仍会被删除
Private Declare Sub yh_rollbackdetail Lib "Hisint" Alias "int_rollbackdetail" _
    (ByVal Serial_No As String, ByVal error_code As String)

'查询某次结算后病人统筹累计,基本统筹支付限额，大病统筹支付限额等信息
'输入参数：顺序号；
'输出参数: 起付线，统筹累计，基本统筹支付限额，大病统筹支付限额，错误代码
Private Declare Sub yh_RyspInfo Lib "Hisint" Alias "int_RyspInfo" _
   (ByVal series_no As String, qfx As Double, tclj As Double, dczfxe As Double, _
    dbxe As Double, ByVal error_code As String)


'======================================nt==============================================
'银海医保（2.0版本）本处只声明与昆明医保不同的函数
'2 费用结算
'输入：顺序号（就诊登记号）、经办人、科室名称、事务控制号；
'输出：全自付金额、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、
'       医疗照顾人员的自费部分、医疗照顾人员的统筹部分、初始化机构名称、错误代码；
Private Declare Sub yh2_feebalance Lib "Hisint" Alias "int_feebalance" _
    (ByVal Serial_No As String, ByVal Arranger As String, ByVal Section_Name As String, ByVal Transaction_No As String, _
    Selfpay As Double, ByRef Hookpay As Double, ByRef Tcpay As Double, ByRef Tcselfpay As Double, ByRef Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    ByVal initinstitution As String, ByVal error_code As String)
    
'3、费用明细更改（备注，可用来完成退费操作）
'输入：顺序号（就诊登记号）、数据批号、大类编码、新的数量、新的价格；
'输出：自付比例、自付金额、允许报销金额、错误代码；
Private Declare Sub yh2_recedefeedetail Lib "Hisint" Alias "int_recedefeedetail" _
    (ByVal Serial_No As String, ByVal data_number As String, ByVal Charge_Category As String, ByVal Count As Double, ByVal Price As Double, _
     Pay_Proportion As Double, Pay_amount As Double, Wipe_Amount As Double, ByVal error_code As String)

'6 虚拟结算
'输入、输出参数、使用场合和时间与费用结算相同。
'输入：顺序号（就诊登记号）
'输出：全自付金额、挂钩自付金额、统筹支付金额、统筹自付金额、基数自付额、超限自付额、大病统筹支付额、大病自付金额、
'       医疗照顾人员的自费部分、医疗照顾人员的统筹部分、初始化机构名称、错误代码；

Private Declare Sub yh2_virtualbalance Lib "Hisint" Alias "int_virtualbalance" _
    (ByVal Serial_No As String, _
    Selfpay As Double, Hookpay As Double, Tcpay As Double, Tcselfpay As Double, Basepay As Double, _
    outpay As Double, Preqpay As Double, Preqselfpay As Double, ActualselfPay As Double, SubsidyPay As Double, _
    ByVal initinstitution As String, ByVal error_code As String)


'本交易作用是出院办理时，修改出院诊断、出院时间时调用。
'输入：顺序号、出院原因、出院时间、出院诊断、出院经办人、出院科室、出院床位；
'输出：错误编码；
Private Declare Sub yh_ReLeaveHosInfo Lib "Hisint" Alias "int_ReLeaveHosInfo" _
   (ByVal series_no As String, ByVal Cyyy As String, ByVal Cysj As String, ByVal Cyzd As String, _
   ByVal Cyjbr As String, ByVal Cyks As String, ByVal Cycw As String, ByVal error_code As String)

'====================================================================================


Public Function 医保初始化_云南建水() As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'返回：初始化成功，返回true；否则，返回false
    On Error GoTo errHandle

    mstrErr = "0000"
    Call yh_init(mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbExclamation, gstrSysName
    Else
        医保初始化_云南建水 = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function 身份标识_云南建水(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strSelfNO-个人编号，刷卡得到；strSelfPwd-病人密码；bytType-识别类型，0-门诊，1-住院
'返回：空或信息串
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空格填充；
    Dim str卡号 As String, str姓名 As String, str性别 As String
    Dim str身份证号 As String, str出生日期 As String, lng年龄 As Long
    Dim str初始化机构 As String, str事务号 As String
    
    Dim strArranger As String
    Dim strSection As String
    Dim strPatiNo As String
    
    Dim str卡类型 As String, lng病种ID As Long, str疾病编码 As String
    Dim rsTemp As New ADODB.Recordset
    Dim dat当前 As Date
    Dim strIdentify As String, str附加 As String
    
    
    On Error GoTo errHandle
    '初始化几个全局的变量
    mstr医保号 = Space(20)
    mstr顺序号 = Space(19)
    mcur帐户余额 = 0
    
    str卡号 = Space(18)
    str姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    str出生日期 = Space(10)
    str初始化机构 = Space(4)
    dat当前 = zlDatabase.Currentdate
    
    
    If frmIdentify云南.GetIdentifyMode(bytType, str卡类型, lng病种ID, str疾病编码) = False Then
        Exit Function
    End If
    DoEvents
        
    '门诊身份证验
    '返回的本次交易的顺序号放在:mstr顺序号,在交易时使用
    '返回的余额存放在mcur帐户余额中，在取余额时使用
    
    '读取IC卡信息
    strArranger = LeftDB(UserInfo.姓名, 8)
    strSection = LeftDB(UserInfo.部门, 24)
    strPatiNo = LeftDB(UserInfo.编号, 12)
    
    Screen.MousePointer = vbHourglass
    mstrErr = "0000"
    Call yh_outpatientidentify(str卡类型, gstr医院编码, strArranger, strSection, strPatiNo, _
        strPatiNo, Format(dat当前, "yyyy-MM-dd"), mstr顺序号, str卡号, _
        mstr医保号, str姓名, str性别, str出生日期, str身份证号, str初始化机构, mcur帐户余额, mstrErr)
    If mstrErr <> "0000" Then
        Screen.MousePointer = vbDefault
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    mstr顺序号 = TrimStr(mstr顺序号)
    mstr医保号 = TrimStr(mstr医保号)
    str卡号 = TrimStr(str卡号)
    
    If mstr顺序号 = "" Then
        MsgBox "未能从前置服务器获得顺序号,请重试或检查卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If str卡号 = "" Then
        MsgBox "未能从卡中读取卡号,请重试或检查卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mstr医保号 = "" Then
        MsgBox "未能从卡中读取医保号,请重试或检查卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '卡号;医保号;密码;姓名;性别;出生日期;身份证;工作单位
    '医保号第一位为卡类型
    mstr医保号 = str卡类型 & Left(mstr医保号, 19)
    strIdentify = str卡号 & ";" & mstr医保号 & ";;" & TrimStr(str姓名) & ";" & TrimStr(str性别) & ";" & TrimStr(str出生日期) & ";" & TrimStr(str身份证号) & ";"
    strIdentify = Replace(strIdentify, " ", "")
    
    If bytType = 1 Then '住院
        '住院身份识别也调用与门诊相同函数，目的是为了多得一些信息（医保号、出生日期）
        '但会在中心数据库中留下一条垃圾数据
        
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    ';8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计;23就诊类型 (1、急诊门诊)
    If bytType = 0 Then
        '如果是门诊,且当前住院,则不允许更新顺序号并退出
        gstrSQL = "Select Count(病人ID) Records From 保险帐户 Where nvl(当前状态,0)=1 And 医保号='" & mstr医保号 & "' And 险类=" & gintInsure
        Call OpenRecordset(rsTemp, "判断是否入院")
        If rsTemp!Records <> 0 Then
            MsgBox "当前医保病人已经在院,不允许在门诊登记!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If bytType = 2 Then
        '不区分门诊与住院，那就不能使用新的顺序号。而不能用以前的
        gstrSQL = "select 顺序号 from 保险帐户 where 险类=" & gintInsure & " and 卡号='" & str卡号 & "'"
        Call OpenRecordset(rsTemp, "建水医保")
        
        If rsTemp.RecordCount > 0 Then
            mstr顺序号 = IIf(IsNull(rsTemp("顺序号")), mstr顺序号, rsTemp("顺序号"))
        End If
    End If
    
    If IsDate(str出生日期) = True Then
        lng年龄 = DateDiff("yyyy", CDate(str出生日期), dat当前)
    End If
    
    str附加 = ";"                                       '8.中心代码
    str附加 = str附加 & ";" & mstr顺序号                '9.顺序号
    str附加 = str附加 & ";"                             '10人员身份
    str附加 = str附加 & ";" & mcur帐户余额              '11帐户余额
    str附加 = str附加 & ";0"                            '12当前状态
    str附加 = str附加 & ";" & IIf(lng病种ID <> 0, lng病种ID, "") '13病种ID
    str附加 = str附加 & ";1"                            '14在职(1,2)
    str附加 = str附加 & ";"                             '15退休证号
    str附加 = str附加 & ";" & lng年龄                   '16年龄段
    str附加 = str附加 & ";"                             '17灰度级
    str附加 = str附加 & ";" & mcur帐户余额              '18帐户增加累计
    str附加 = str附加 & ";0"                            '19帐户支出累计
    str附加 = str附加 & ";"                             '20进入统筹累计
    str附加 = str附加 & ";"                             '21统筹报销累计
    str附加 = str附加 & ";"                             '22住院次数累计
    str附加 = str附加 & ";"                             '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(bytType, strIdentify & str附加, lng病人ID)
    If lng病人ID = 0 Then Exit Function '未建立正确的保险帐户
    
    If bytType = 0 And lng病种ID > 0 Then
        '如果是特殊病、慢性病门诊，同时进行就诊登记
        
        '再次初始化变量
        mstr医保号 = Space(20)
        str卡号 = Space(18)
        str姓名 = Space(60)
        str性别 = Space(3)
        str身份证号 = Space(20)
        str出生日期 = Space(10)
        str初始化机构 = Space(4)
        
        
        str事务号 = Get事务号
        If str事务号 = "" Then
            Exit Function
        End If
        
        mstrErr = "0000"
        Call yh_admit(str卡类型, gstr医院编码, LeftDB(UserInfo.姓名, 8), "门诊", _
            LeftDB(lng病人ID, 12), LeftDB(lng病人ID, 12), "1", LeftDB(str疾病编码, 8), _
            Format(dat当前, "yyyy-MM-dd"), "无", str事务号, mstr顺序号, str卡号, _
            mstr医保号, str姓名, str性别, str出生日期, str身份证号, str初始化机构, mstrErr)
        
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
            '医保数据库回滚
            Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
            
            Exit Function
        End If
        mstr顺序号 = TrimStr(mstr顺序号) '1、用于门诊预算
        If mstr顺序号 = "" Then
            MsgBox "不能得到正确的入院登记顺序号。", vbInformation, gstrSysName
            Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
            Exit Function
        End If
        mstr医保号 = str卡类型 & Left(TrimStr(mstr医保号), 19) '2、用于门诊预算
        str卡号 = TrimStr(str卡号)
    
        '强制把登记顺序号、及新的医保号填入
        gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & gintInsure & _
                    ",'" & str卡号 & "','" & mstr医保号 & "','" & mstr顺序号 & "')"
        Call ExecuteProcedure("建水医保")
        
    End If
    '得到费用明细传递的事务控制号，以便于多次重试
    If bytType = 0 Then
        mstr明细事务号 = Get事务号 '3、用于门诊结算
        If mstr明细事务号 = "" Then
            Exit Function
        End If
    End If
    
    mlng病人ID = lng病人ID '4、用于门诊预算
    
    '返回格式:中间插入病人ID
    身份标识_云南建水 = strIdentify & ";" & lng病人ID & str附加
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 个人余额_云南建水(strSelfNo As String, ByVal bytPlace As Byte) As Currency
'功能: 提取参保病人个人帐户余额
'参数: strSelfNO-病人个人编号
'      表示调用位置：10-门诊,20-入院,30-预交,40-结算
'返回: 返回个人帐户余额的金额
    
    On Error GoTo errHandle
    
    If strSelfNo = mstr医保号 And (bytPlace = 10 Or bytPlace = 20) Then
        '直接利用上次身份识别时得到的数据返回
        个人余额_云南建水 = mcur帐户余额
    Else
        '读IC卡上的余额
        Call Get卡余额(strSelfNo, 个人余额_云南建水)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊虚拟结算_云南建水(rs明细 As ADODB.Recordset, str结算方式 As String) As Boolean
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'字段：病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保
    Dim str数据批号 As String, strTemp As String
    
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String, str规格 As String, str产地 As String
    Dim cur发生费用 As Currency, dbl金额 As Double, dbl数量 As Double
    
    Dim rsTemp As New ADODB.Recordset
    
    If rs明细.EOF = True Then
        MsgBox "请输入费用明细再进行医保预算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rs明细("病人ID") <> mlng病人ID Then
        MsgBox "该病人未通过身份验证，不能进行预结算。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '只有特殊门诊才使用本函数
    On Error GoTo errHandle
    
    '删除前置服务器的所有未结明细
    mstrErr = "0000"
    Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
    '不论是否成功暂时不管，
'    If mstrErr <> "0000" Then
'        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
'        Exit Function
'    End If
            
    '费用明细传递
    strTemp = rs明细("病人ID") & "_" & Format(zlDatabase.Currentdate, "ddHHmmss")
    Do Until rs明细.EOF
        gstrSQL = "select A.名称,A.编码,A.类别,A.计算单位,B.项目编码,B.附注" & _
                    " ,Decode(Sign(Instr(A.规格,'┆')),0,A.规格,Substr(A.规格,1,Instr(A.规格,'┆')-1)) as 规格" & _
                    " ,Decode(Sign(Instr(A.规格,'┆')),0,A.规格,Substr(A.规格,Instr(A.规格,'┆')+1)) as 产地" & _
                    " from 收费细目 A,保险支付项目 B where A.ID=" & rs明细("收费细目ID") & " and A.ID=B.收费细目ID and B.险类=" & gintInsure
        Call OpenRecordset(rsTemp, "门诊预算")
        If rsTemp.EOF = True Then
            MsgBox "有项目未设置医保编码，不能结算。", vbInformation, gstrSysName
            Exit Function
        End If
        
        str医生 = LeftDB(UserInfo.姓名, 8)
        str规格 = LeftDB(IIf(IsNull(rsTemp("规格")), "无规格", rsTemp("规格")), 30)
        str产地 = LeftDB(IIf(IsNull(rsTemp("产地")), "", rsTemp("产地")), 30)
        str科室 = LeftDB(UserInfo.部门, 24)
        '不能传递负数，传0的目的是为了删除已经上传但被冲销的费用记录
        dbl数量 = Val(IIf(rs明细("数量") > 0, rs明细("数量"), 0))
        dbl金额 = Val(IIf(rs明细("单价") > 0, rs明细("单价"), 0))
        
        str数据批号 = ToVarchar(strTemp & "_" & rs明细.AbsolutePosition, 18)
        
        mstrErr = "0000"
        Call yh_feedetailtrans(mstr顺序号, str数据批号, ToVarchar(rsTemp("项目编码"), 2), rsTemp("项目编码"), _
            rsTemp("名称"), dbl数量, dbl金额, str产地, str规格, "", str医生, str科室, mstr明细事务号, str医生, _
            cur自付比例, cur自付金额, cur报销金额, mstrErr)
        If mstrErr <> "0000" Then
            MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
            '医保数据库回滚
            Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
            Exit Function
        End If
        
        cur发生费用 = cur发生费用 + rs明细("实收金额")
        rs明细.MoveNext
    Loop
        
    '虚拟结算
    Dim str结算标志 As String, cur病人自费 As Double, cur余额 As Currency
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    Dim str结算事务号 As String
    
    '用于门诊预算
    str结算事务号 = Get事务号
    If str结算事务号 = "" Then
        Exit Function
    End If
    
    str初始化机构 = Space(4)
    mstrErr = "0000"
    Call yh2_virtualbalance(mstr顺序号, cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, _
        cur超限自付, cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .发生费用金额 = cur发生费用
    End With
    
    cur余额 = 个人余额_云南建水(mstr医保号, 10)
    If cur特殊人员统筹 > 0 Then
        cur病人自费 = cur特殊人员自付
    Else
        cur病人自费 = cur全自付 + cur挂钩自付 + cur基数自付 + cur统筹自付 + cur大病自付 + cur超限自付 - cur公务员统筹
    End If
    cur余额 = IIf(cur余额 > cur病人自费, cur病人自费, cur余额) '取两者的小值
        
    str结算方式 = "个人帐户;" & cur余额 & ";1" '允许修改
    
    If cur统筹支付 <> 0 Then
        str结算方式 = str结算方式 & "|医保基金;" & cur统筹支付 & ";0"
    End If
    If cur大病统筹 <> 0 Then
        str结算方式 = str结算方式 & "|大病统筹;" & cur大病统筹 & ";0"
    End If
    If cur公务员统筹 <> 0 Then
        str结算方式 = str结算方式 & "|公务员补助;" & cur公务员统筹 & ";0"
    End If
    If cur特殊人员统筹 > 0 Then
        str结算方式 = str结算方式 & "|特殊补助;" & cur特殊人员统筹 & ";0"
    End If
    
    门诊虚拟结算_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算_云南建水(lng结帐ID As Long, cur个人帐户 As Currency, strSelfNo As String) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur支付金额   从个人帐户中支出的金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
    Dim rs明细 As New ADODB.Recordset, rsTemp As New ADODB.Recordset, lng病种ID As Long
    Dim i As Long, curDate As Date, cur发生费用 As Currency, lng病人ID As Long
    Dim str卡类型 As String
    Dim str结算事务号 As String   '事务控制号
    Dim str初始化机构 As String
    
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String
    Dim str规格 As String, str产地 As String
    
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double
    Dim cur特殊人员统筹 As Double, cur特殊人员自付 As Double, cur公务员统筹 As Double
    
    
    On Error GoTo errHandle
    '此时所有收费细目必然有对应的医保编码
    gstrSQL = "Select A.ID,A.病人ID,A.NO,A.登记时间,A.开单人 as 医生," & _
            "   A.数次*A.付数 as 数量,Round(A.结帐金额/(A.数次*A.付数),2) as 实际价格,A.结帐金额," & _
            "   A.收费类别,D.项目编码 as 收费项目,B.名称 as 项目名称," & _
            "   decode(Instr(B.规格,'┆'),0,B.规格,substr(B.规格,1,Instr(B.规格,'┆')-1)) as 规格," & _
            "   decode(Instr(B.规格,'┆'),0,'',substr(B.规格,Instr(B.规格,'┆')+1)) as 产地," & _
            "   C.名称 as 科室名称" & _
            " From (Select * From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9) A,收费细目 B,部门表 C,保险支付项目 D " & _
            " Where A.收费细目ID=B.ID And A.开单部门ID=C.ID And A.收费细目ID=D.收费细目ID And D.险类=" & gintInsure & _
            " Order by A.ID"
    Call OpenRecordset(rs明细, "建水医保")
    
    If rs明细.EOF = True Then
        MsgBox "没有填写收费记录", vbExclamation, gstrSysName
        Exit Function
    End If
    lng病人ID = rs明细("病人ID")
    
    '判断该病人是否属于特殊门诊
    gstrSQL = "select nvl(病种ID,0) 病种ID from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "医保接口")
    If rsTemp.EOF = False Then
        '有特殊病的病人需要预算
        lng病种ID = rsTemp("病种ID")
    End If
    
    '一、费用明细传递
    '顺序号采用身份验证时返回的值:mstr顺序号
    str医生 = LeftDB(IIf(IsNull(rs明细("医生")), UserInfo.姓名, rs明细("医生")), 8)
    str科室 = LeftDB(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 24)
    If lng病种ID = 0 Then
        '普通门诊由于没有预算，所以还需要传输费用明细
        
        '删除前置服务器的所有未结明细（由于前一次确定时明细传输成功，但结算失败时）
        mstrErr = "0000"
        Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
        
        Do Until rs明细.EOF
            str规格 = LeftDB(IIf(IsNull(rs明细("规格")), "无规格", rs明细("规格")), 30)
            str产地 = LeftDB(IIf(IsNull(rs明细("产地")), "", rs明细("产地")), 30)
            str科室 = LeftDB(IIf(IsNull(rs明细("科室名称")), UserInfo.部门, rs明细("科室名称")), 24)
            cur发生费用 = cur发生费用 + rs明细("结帐金额")
            
            mstrErr = "0000"
            Call yh_feedetailtrans(mstr顺序号, rs明细("ID"), LeftDB(rs明细("收费项目"), 2), rs明细("收费项目"), LeftDB(rs明细("项目名称"), 24), _
                rs明细("数量"), rs明细("实际价格"), str产地, str规格, "", str医生, str科室, mstr明细事务号, str医生, _
                cur自付比例, cur自付金额, cur报销金额, mstrErr)
            
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
                '医保数据库回滚
                Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
                Exit Function
            End If
            rs明细.MoveNext
        Loop
    Else
        '经过预算的，只处理总金额
'        Do Until rs明细.EOF
'            cur发生费用 = cur发生费用 + rs明细("结帐金额")
'            rs明细.MoveNext
'        Loop
        cur发生费用 = g结算数据.发生费用金额 '该处是应收金额，与预算保持一致
    End If
        
    '二、写IC卡
    str卡类型 = Left(strSelfNo, 1)
    str初始化机构 = Space(4)
    mstrErr = "0000"
    Call yh_cardpay(str卡类型, mstr顺序号, str医生, "门诊收费", CDbl(cur个人帐户), str初始化机构, mstrErr)
    
    If mstrErr <> "0000" Then
        '医保数据库回滚
        Call yh_transaction("1", mstr顺序号, mstr明细事务号, "0", mstrErr)
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    '三、费用结算
    str结算事务号 = Get事务号
    If str结算事务号 = "" Then
        Exit Function
    End If
    
    str初始化机构 = Space(4)
    mstrErr = "0000"
    Call yh2_feebalance(mstr顺序号, str医生, str科室, str结算事务号, _
        cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, cur超限自付, cur大病统筹, _
        cur大病自付, cur特殊人员自付, cur特殊人员统筹, str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str结算事务号, "0", mstrErr)
        Exit Function
    End If
    Call yh_transaction("2", mstr顺序号, str结算事务号, "1", mstrErr)
    
    '四、保存结算记录
    '---------------------------------------------------------------------------------------------
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    '定义 cur统筹累计 变量的目的是为了调用API，类型兼容
    Dim cur起付线 As Double, cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double
    Dim int住院次数累计 As Integer
    curDate = zlDatabase.Currentdate
            
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call ExecuteProcedure("建水医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    
    gstrSQL = "zl_保险结算记录_insert(1," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur基数自付 & "," & Get病种编码(lng病种ID) & "," & cur特殊人员自付 & "," & _
        cur发生费用 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & "," & cur个人帐户 & ",'" & mstr顺序号 & "')"
    Call ExecuteProcedure("建水医保")
    '---------------------------------------------------------------------------------------------
    
    门诊结算_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 门诊结算冲销_云南建水(lng结帐ID As Long, cur个人帐户 As Currency, lng病人ID As Long) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID     收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      cur个人帐户   从个人帐户中支出的金额

    门诊结算冲销_云南建水 = False
End Function

Public Function 个人帐户转预交_云南建水(lng预交ID As Long, cur个人帐户 As Currency, strSelfNo As String, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
    Dim str卡类型 As String
    Dim str初始化机构 As String
    Dim str医生 As String
    
    On Error GoTo errHandle
    
    If Is卡正确(lng病人ID) = False Then Exit Function
    
    str初始化机构 = Space(4)
    str卡类型 = Left(strSelfNo, 1)
    
    mstrErr = "0000"
    str医生 = LeftDB(UserInfo.姓名, 8)
    Call yh_cardpay(str卡类型, str顺序号, CStr(UserInfo.姓名), "预交款", cur个人帐户, str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim rsTemp As New ADODB.Recordset
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date
    
    '---------------------------------------------------------------------------------------------
    '填写结算表
    curDate = zlDatabase.Currentdate
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
            
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 + cur个人帐户 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ")"
    Call ExecuteProcedure("建水医保")
    
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_保险结算记录_insert(3," & lng预交ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & ",0,0,0," & cur个人帐户 & ",0,0,0,0,0,0," & _
        cur个人帐户 & ",'" & str顺序号 & "')"
    Call ExecuteProcedure("建水医保")
    
    个人帐户转预交_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 入院登记_云南建水(lng病人ID As Long, lng主页ID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false

    Dim rsTemp As New ADODB.Recordset
    Dim str卡类型 As String
    Dim str卡号 As String
    Dim str姓名 As String
    Dim str性别 As String
    Dim str出生日期 As String
    Dim str身份证号 As String
    Dim str初始化机构 As String
    Dim str事务号 As String   '事务控制号
    
    On Error GoTo errHandle
    mstr顺序号 = Space(19)
    str医保号 = Space(20)
    str事务号 = Space(18)
    str卡号 = Space(18)
    str姓名 = Space(60)
    str性别 = Space(3)
    str出生日期 = Space(10)
    str身份证号 = Space(20)
    str初始化机构 = Space(4)
    
    '注意：此时不能读保险帐户，因为尚未取到医保号，而是需要返回医保号
    gstrSQL = "Select A.入院日期,A.入院病床,B.名称 as 入院科室,C.住院号,A.登记时间,D.医保号,E.编码 as 病种编码,E.类别 as 病种类别 " & _
            " From 病案主页 A,部门表 B,病人信息 C,保险帐户 D,保险病种 E " & _
            " Where A.入院科室ID=B.ID And A.病人ID=" & lng病人ID & " And A.主页ID=" & lng主页ID & _
            " And A.病人ID=C.病人ID And A.病人ID=D.病人ID and D.险类=" & gintInsure & " and D.病种ID=E.ID(+)"
    Call OpenRecordset(rsTemp, "建水医保")
    
    If rsTemp.EOF = True Then
        MsgBox "没有发现此病人的信息！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If IsNull(rsTemp("医保号")) = False Then
        str卡类型 = Left(rsTemp("医保号"), 1)
    Else
        Dim lng疾病ID As Long, str疾病编码 As String
        If frmIdentify云南.GetIdentifyMode(1, str卡类型, lng疾病ID, str疾病编码) = False Then Exit Function
    End If
    
    '入院登记
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    
    mstrErr = "0000"
    Call yh_admit(str卡类型, gstr医院编码, LeftDB(UserInfo.姓名, 8), LeftDB(rsTemp("入院科室"), 24), _
        LeftDB(lng病人ID, 12), LeftDB(rsTemp("住院号"), 12), IIf(rsTemp("病种类别") <> "0", "1", "0"), LeftDB(IIf(IsNull(rsTemp("病种编码")), "", rsTemp("病种编码")), 8), _
        Format(rsTemp!入院日期, "yyyy-MM-dd"), LeftDB(获取入出院诊断(lng病人ID, lng主页ID, True, False), 50), str事务号, mstr顺序号, str卡号, _
        str医保号, str姓名, str性别, str出生日期, str身份证号, str初始化机构, mstrErr)
    
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
        
        Exit Function
    End If
    mstr顺序号 = TrimStr(mstr顺序号)
    If mstr顺序号 = "" Then
        MsgBox "不能得到正确的入院登记顺序号。", vbInformation, gstrSysName
        Call yh_transaction("0", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    str医保号 = str卡类型 & Left(TrimStr(str医保号), 19)
    str卡号 = TrimStr(str卡号)
    
    '强制把登记顺序号、及新的医保号填入
    gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & gintInsure & _
                ",'" & str卡号 & "','" & str医保号 & "','" & mstr顺序号 & "')"
    Call ExecuteProcedure("建水医保")
    '改变病人状态
    gstrSQL = "zl_保险帐户_入院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("建水医保")
    
    Call yh_transaction("0", mstr顺序号, str事务号, "1", mstrErr)
    
    入院登记_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 出院登记_云南建水(lng病人ID As Long, lng主页ID As Long, str顺序号 As String) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lng病人ID-病人ID；lng主页ID-主页ID
'返回：交易成功返回true；否则，返回false
            '取入院登记验证所返回的顺序号
    Dim str事务号 As String   '事务控制号
    
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    
    Dim rsInfo As New ADODB.Recordset
    Dim str出院原因 As String, str出院时间 As String, str出院诊断 As String
    Dim str出院经办人 As String, str出院科室 As String, str出院床号 As String
    '出院方式:1-正常;2-转院;3-死亡；对应医保的出院原因：0、正常出院；1、死亡；2、转院；3、审批未住院（中途取消）；9、其他

    On Error GoTo errHandle
    str初始化机构 = Space(4)
    
    str事务号 = Get事务号
    If str事务号 = "" Then
        
    End If
    mstr顺序号 = str顺序号
        
    '出院登记是通过调用结算交易完成。此时假设病人的费用已经全部结清
    mstrErr = "0000"
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If

    '更新出院诊断
    mstrErr = "0000"
    gstrSQL = "select decode(出院方式,'正常',0,'转院',2,'死亡',1,9) 出院方式 From 病案主页 " & _
            " Where 病人ID = " & lng病人ID & " And 主页ID = " & lng主页ID
    Call OpenRecordset(rsInfo, "出院方式")
    str出院原因 = rsInfo!出院方式

    gstrSQL = "select b.名称 出院科室,床号,终止时间,操作员姓名  " & _
             " from 病人变动记录 A,部门表 B  " & _
             " where 病人ID=" & lng病人ID & " and 终止原因=1 " & _
             " and A.科室ID=B.ID"
    Call OpenRecordset(rsInfo, "出院情况")
    str出院时间 = Format(rsInfo!终止时间, "yyyy-MM-dd HH:mm:ss")
    str出院科室 = LeftDB(rsInfo!出院科室, 20)
    str出院床号 = LeftDB(rsInfo!床号, 10)
    str出院经办人 = LeftDB(rsInfo!操作员姓名, 20)
    str出院诊断 = LeftDB(获取入出院诊断(lng病人ID, lng主页ID, False, False), 100)
    Call yh_ReLeaveHosInfo(mstr顺序号, str出院原因, str出院时间, str出院诊断, str出院经办人, str出院科室, str出院床号, mstrErr)
    
    '改变病人状态
    gstrSQL = "zl_保险帐户_出院(" & lng病人ID & "," & gintInsure & ")"
    Call ExecuteProcedure("建水医保")
    
    '结算过程不用调用本函数
    Call yh_transaction("2", mstr顺序号, str事务号, "1", mstrErr)
    
    出院登记_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Function 住院虚拟结算_云南建水(rsExse As Recordset, ByVal lng病人ID As Long) As String
'功能：获取该病人指定结帐内容的可报销金额；
'参数：rsExse-需要结算的费用明细记录集合；strSelfNO-医保号；strSelfPwd-病人密码；
'返回：可报销金额串:"报销方式;金额;是否允许修改|...."
'注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    Dim str事务号 As String   '事务控制号
    Dim cn上传 As New ADODB.Connection, str数据批号 As String
    
    Dim cur自付比例 As Double, cur自付金额 As Double, cur报销金额 As Double
    Dim str医生 As String, str科室 As String, str规格 As String, str产地 As String
    Dim cur发生费用 As Currency, dbl金额 As Double, dbl数量 As Double
    
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
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
    '打开另外一个连接串，以达到不受当前连接事务的控制
    cn上传.ConnectionString = gcnOracle.ConnectionString
    cn上传.Open
    
    '顺序号取入院登记验证返回的
    gstrSQL = "Select 医保号,顺序号 From 保险帐户 " & _
              "Where 顺序号 is Not NULL And 病人ID=" & lng病人ID & " And 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "虚拟结算")
    
    If rsTemp.EOF Then
        MsgBox "未发现该病人的住院交易顺序号,不能执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mstr顺序号 = rsTemp("顺序号")
    
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    
    '费用明细传递
    Do Until rsExse.EOF
        If IIf(IsNull(rsExse("是否上传")), "0", rsExse("是否上传")) = "0" Then
            '建水医保只处理尚未上传的费用记录
            
            str医生 = LeftDB(IIf(IsNull(rsExse("医生")), UserInfo.姓名, rsExse("医生")), 8)
            str规格 = LeftDB(IIf(IsNull(rsExse("规格")), "无规格", rsExse("规格")), 30)
            str产地 = LeftDB(IIf(IsNull(rsExse("产地")), "", rsExse("产地")), 30)
            str科室 = LeftDB(IIf(IsNull(rsExse("开单部门")), UserInfo.部门, rsExse("开单部门")), 24)
            '不能传递负数：已经被注释（XQ 2003-04-24）
'            If rsExse("记录状态") = 1 And rsExse("数量") < 0 Then
'                MsgBox "医保不支持直接录入负数，只能选择原有单据进行冲销。", vbInformation, gstrSysName
'                Exit Function
'            End If
            '传0的目的是为了删除已经上传但被冲销的费用记录
'            dbl数量 = Val(IIf(rsExse("数量") > 0, rsExse("数量"), 0))
'            dbl金额 = Val(IIf(rsExse("价格") > 0, rsExse("价格"), 0))
            dbl数量 = CDbl(NVL(rsExse("数量"), 0))
            dbl金额 = CDbl(NVL(rsExse("价格"), 0))
            
            mstrErr = "0000"
            
            '为了让负记录能正确找到正记录，所以数据批号中不包含记录状态
'            str数据批号 = rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") '& "_" & rsExse("记录状态")
            str数据批号 = rsExse("NO") & "_" & rsExse("序号") & "_" & rsExse("记录性质") & "_" & rsExse("记录状态")
            Call yh_feedetailtrans(mstr顺序号, str数据批号, Left(rsExse("医保项目编码"), 2), rsExse("医保项目编码"), _
                rsExse("收费名称"), dbl数量, dbl金额, str产地, str规格, "", str医生, str科室, str事务号, str医生, _
                cur自付比例, cur自付金额, cur报销金额, mstrErr)
            If mstrErr <> "0000" Then
                MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
                '医保数据库回滚
                Call yh_transaction("1", mstr顺序号, str事务号, "0", mstrErr)
                Exit Function
            End If
            
            '为该条费用记录打上上传标志。上传一条处理一条
            gstrSQL = "zl_病人费用记录_上传('" & rsExse("NO") & "'," & rsExse("序号") & "," & rsExse("记录性质") & "," & rsExse("记录状态") & ")"
            cn上传.Execute gstrSQL, , adCmdStoredProc
        End If
        
        cur发生费用 = cur发生费用 + rsExse("金额")
        rsExse.MoveNext
    Loop
        
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    '虚拟结算
    Dim str结算标志 As String
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    
    str初始化机构 = Space(4)
    mstrErr = "0000"
    str结算标志 = "0" '虚拟结算
    Call yh2_virtualbalance(mstr顺序号, cur全自付, cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, _
        cur超限自付, cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    '保存临时数据，为结算操作做准备
    With g结算数据
        .病人ID = lng病人ID
        .发生费用金额 = cur发生费用
    End With
    
    住院虚拟结算_云南建水 = "医保基金;" & cur统筹支付 & ";0"
    If cur大病统筹 <> 0 Then
        住院虚拟结算_云南建水 = 住院虚拟结算_云南建水 & "|大病统筹;" & cur大病统筹 & ";0"
    End If
    If cur公务员统筹 <> 0 Then
        住院虚拟结算_云南建水 = 住院虚拟结算_云南建水 & "|公务员补助;" & cur公务员统筹 & ";0"
    End If
    If cur特殊人员统筹 > 0 Then
        住院虚拟结算_云南建水 = 住院虚拟结算_云南建水 & "|特殊补助;" & cur特殊人员统筹 & ";0"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算_云南建水(lng结帐ID As Long, str顺序号 As String, ByVal lng病人ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID -病人结帐记录ID, 从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。(由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
    Dim str事务号 As String   '事务控制号
    Dim str结算标志 As String
    Dim cur全自付 As Double, cur挂钩自付 As Double, cur统筹支付 As Double
    Dim cur统筹自付 As Double, cur基数自付 As Double, cur超限自付 As Double
    Dim cur大病统筹 As Double, cur大病自付 As Double, str初始化机构 As String
    Dim cur特殊人员自付 As Double, cur特殊人员统筹 As Double, cur公务员统筹 As Double
    
    Dim cur帐户增加累计 As Currency, cur帐户支出累计 As Currency
    Dim cur进入统筹累计 As Currency, cur统筹报销累计 As Currency
    Dim int住院次数累计 As Integer, curDate As Date, lng病种ID As Long, rsTemp As New ADODB.Recordset
    
    str初始化机构 = Space(4)
    
    On Error GoTo errHandle
    '取入院登记验证所返回的顺序号
    mstr顺序号 = str顺序号
    str事务号 = Get事务号
    If str事务号 = "" Then
        Exit Function
    End If
    
    
    '费用结算:结帐。为了达到中途结帐的目的，没有使用结算函数
    mstrErr = "0000"
    str结算标志 = "1"   '预结算
    Call yh2_feebalance(mstr顺序号, LeftDB(UserInfo.姓名, 8), LeftDB(UserInfo.部门, 24), str事务号, cur全自付, _
        cur挂钩自付, cur统筹支付, cur统筹自付, cur基数自付, cur超限自付, _
        cur大病统筹, cur大病自付, cur特殊人员自付, cur特殊人员统筹, _
        str初始化机构, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        '医保数据库回滚
        Call yh_transaction("2", mstr顺序号, str事务号, "0", mstrErr)
        Exit Function
    End If
    
    
    '填写结算表
    curDate = zlDatabase.Currentdate
    '读出该病人本次结算的病种信息
    gstrSQL = "Select nvl(病种ID,0) 病种ID From 保险帐户 A Where A.险类=" & gintInsure & " and A.病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "保险结算")
    If rsTemp.EOF = False Then
        lng病种ID = rsTemp("病种ID")
    End If
    
    
    '帐户年度信息
    Call Get帐户信息(lng病人ID, Year(curDate), int住院次数累计, cur帐户增加累计, cur帐户支出累计, cur进入统筹累计, cur统筹报销累计)
    If int住院次数累计 = 0 Then int住院次数累计 = Get住院次数(lng病人ID)
            
    '定义 cur统筹累计 变量的目的是为了调用API，类型兼容
    Dim cur起付线 As Double, cur统筹累计 As Double, cur基本统筹限额 As Double, cur大额统筹限额 As Double
    cur统筹报销累计 = cur统筹报销累计 + cur统筹支付 + cur大病统筹
            
    '保险结算记录(因为"性质,记录ID"唯一,所以本次新结帐ID肯定为插入)
    gstrSQL = "zl_帐户年度信息_insert(" & lng病人ID & "," & gintInsure & "," & Year(curDate) & "," & _
        cur帐户增加累计 & "," & cur帐户支出累计 & "," & _
        cur进入统筹累计 + cur统筹支付 + cur统筹自付 + cur基数自付 + cur超限自付 + cur大病统筹 + cur大病自付 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur起付线 & "," & cur起付线 & "," & cur基本统筹限额 & "," & cur大额统筹限额 & ")"
    Call ExecuteProcedure("建水医保")
    
    gstrSQL = "zl_保险结算记录_insert(2," & lng结帐ID & "," & gintInsure & "," & lng病人ID & "," & _
        Year(curDate) & "," & cur帐户增加累计 & "," & cur帐户支出累计 & "," & cur进入统筹累计 & "," & _
        cur统筹报销累计 & "," & int住院次数累计 & "," & cur基数自付 & "," & Get病种编码(lng病种ID) & "," & cur特殊人员自付 & "," & _
        g结算数据.发生费用金额 & "," & cur全自付 & "," & cur挂钩自付 & "," & _
        cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & "," & cur大病自付 & "," & cur超限自付 & ",0,'" & mstr顺序号 & "'," & g结算数据.主页ID & ")"
    Call ExecuteProcedure("建水医保")
    
    '保险结算计算
    gstrSQL = "zl_保险结算计算_insert(" & lng结帐ID & ",0," & cur统筹支付 + cur统筹自付 & "," & cur统筹支付 & ",NULL)"
    Call ExecuteProcedure("建水医保")
    
    住院结算_云南建水 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function 住院结算冲销_云南建水(lng结帐ID As Long) As Boolean
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
    
    If gintInsure = TYPE_云南建水 Then Exit Function '建水医保不支持
    
End Function

Public Function 错误信息_云南建水(ByVal lngErrCode As Long) As String
'功能：根据错误号返回错误信息

End Function

Private Function LeftDB(ByVal strText As String, ByVal lngLength As Long)
'功能：按数据库的长度计算方式得到字符串的实际可用子串
    LeftDB = StrConv(MidB(StrConv(strText, vbFromUnicode), 1, lngLength), vbUnicode)
End Function

Private Function Get事务号() As String
    Dim str事务号 As String
    
    On Error GoTo errHandle
    
    str事务号 = Space(18)
    Call yh_gettranssequence(str事务号) '这里费用传递和结算是两个事务号
    str事务号 = TrimStr(str事务号)
    If str事务号 = "" Then
        MsgBox "获取事务控制号失败。", vbInformation, gstrSysName
    End If
    
    Get事务号 = str事务号
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Is卡正确(ByVal lng病人ID As Long) As Boolean
'功能：判断读卡器的卡是否就是要操作的病人的
    Dim rsTemp As New ADODB.Recordset
    Dim str卡号_库 As String, str卡号 As String, str卡类型 As String
    
    Dim cur余额 As Double, str姓名 As String, str性别 As String
    Dim str身份证号 As String, lng年龄 As Double
    
    On Error GoTo errHandle
    
    gstrSQL = "select 卡号,医保号 from 保险帐户 where 险类=" & gintInsure & " and 病人ID=" & lng病人ID
    Call OpenRecordset(rsTemp, "建水医保")
    
    str卡号_库 = IIf(IsNull(rsTemp("卡号")), "", rsTemp("卡号"))
    str卡类型 = Left(rsTemp("医保号"), 1)
    
    str卡号 = Space(20)
    str姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    
    mstrErr = "0000"
    Call yh_cardinfo(str卡类型, cur余额, str卡号, str姓名, str性别, str身份证号, lng年龄, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    str卡号 = TrimStr(str卡号)
    
    If str卡号 <> str卡号_库 Then
        MsgBox "刷卡器中的卡不是当前病人的，请插入正确的IC卡。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Is卡正确 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get卡余额(ByVal str医保号 As String, 卡余额 As Currency) As Boolean
'功能：得到卡余额
    Dim cur余额 As Double, str姓名 As String, str性别 As String, str卡号 As String
    Dim str身份证号 As String, lng年龄 As Double, str卡类型 As String
    
    str卡类型 = Left(str医保号, 1)
    
    str卡号 = Space(20)
    str姓名 = Space(60)
    str性别 = Space(3)
    str身份证号 = Space(20)
    
    mstrErr = "0000"
    Call yh_cardinfo(str卡类型, cur余额, str卡号, str姓名, str性别, str身份证号, lng年龄, mstrErr)
    If mstrErr <> "0000" Then
        MsgBox GetErrInfo(mstrErr), vbInformation, gstrSysName
        Exit Function
    End If
    
    卡余额 = cur余额
    Get卡余额 = True
End Function

Private Function Get病种编码(ByVal lng病种ID As Long) As String
'功能：判断读卡器的卡是否就是要操作的病人的
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "select 编码 from 保险病种 where ID=" & lng病种ID
    Call OpenRecordset(rsTemp, "建水医保")
    
    If rsTemp.EOF = False Then
        Get病种编码 = Val(rsTemp("编码")) '为了保存在封顶线字段，所以必须是数字
        If Val(Get病种编码) = 0 Then Get病种编码 = "9999" '特批特种病也为0000，所以强制改为9999
    Else
        Get病种编码 = 0
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



