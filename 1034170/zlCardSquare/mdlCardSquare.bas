Attribute VB_Name = "mdlCardSquare"
Option Explicit
Public Enum g小数类型
    g_数量 = 0
    g_成本价
    g_售价
    g_金额
    g_折扣率
End Enum
Private Type m_小数位
    数量小数 As Integer
    成本价小数 As Integer
    零售价小数 As Integer
    金额小数 As Integer
    折扣率 As Integer
End Type
Public g_小数位数 As m_小数位

'小数格式化串
Public Type g_FmtString
    FM_数量 As String
    FM_成本价 As String
    FM_零售价 As String
    FM_金额 As String
    FM_折扣率 As String
End Type
Public Enum gCardEditType   '卡编辑类型
    gEd_发卡 = 0
    gEd_批量发卡 = 1
    gEd_修改 = 2
    gEd_删除 = 3
    gEd_查询 = 4
    gEd_充值 = 5
    gEd_回退 = 6
    gEd_回收 = 7
    gEd_取消回收 = 8
    gEd_退卡 = 9
    gEd_取消退卡 = 10
End Enum
Public Type zlTyCustumRecordset
    rs收费类别 As ADODB.Recordset
    rs消费卡接口 As ADODB.Recordset
    rs收费类别汇总 As ADODB.Recordset
    rs分单类别汇总 As ADODB.Recordset
    dbl费用总额 As Double
    dblHIS最大消费额 As Double
    dbl已刷累计额 As Double
End Type
Public gblnShowCard As Boolean  '就诊卡号显示(true,显示卡号,false,加密显示)
Public gObjXFCards As clsCards  '专门针对消费卡的(要管理卡号)
Public gobjSquare As SquareCard

Public grsStatic As zlTyCustumRecordset
Public gVbFmtString As g_FmtString
Public gOraFmtString As g_FmtString
Public gbln自动读取 As Boolean '当前是否为射频卡
Public gblnCardNoSHowPW As Boolean  '卡号显示密文
Public gDebug As Boolean '调试开关
Public gobjComLib As Object
Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjControl As Object
Public gstrLike As String  '项目匹配方法,%或空
Private Type Ty_TestDebug
    blndebug As Boolean
    objSquareCard As clsCard
    BytType  As Byte  '1-随机产生卡号,2-读取卡号
    strStartNo As String    '开始卡号
    bln补调交易 As Boolean
End Type
Public gTy_TestBug As Ty_TestDebug
Public gobjStartCards As Collection  '启动的刷卡对象集
 
Public gbytDec As Byte '费用金额的小数点位数
Public gstrDec As String '按小数位数计算的格式化串,如"0.0000"
Public gintFeePrecision As Integer    '费用小数精度
Public gstrFeePrecisionFmt As String '费用小数格式:0.00000
Public gblnOK As Boolean
'LED语音报价控制
Public gblnLED As Boolean '是否使用Led显示

Public Enum 医院业务
    support门诊预算 = 0
    
    
    support预交退个人帐户 = 2
    support结帐退个人帐户 = 3
    
    support收费帐户全自费 = 4       '门诊收费和挂号是否用个人帐户支付全自费部分。全自费：指统筹比例为0的金额或超出限价的床位费部分
    support收费帐户首先自付 = 5     '门诊收费和挂号是否用个人帐户支付首先自付部分。首先自付：（1-统筹比例）* 金额
    
    support结算帐户全自费 = 6       '住院结算与特殊门诊是否用个人帐户支付全自费部分。
    support结算帐户首先自付 = 7     '住院结算与特殊门诊是否用个人帐户支付首先自付部分。
    support结算帐户超限 = 8         '住院结算与特殊门诊是否用个人帐户支付超限部分。
    
    support结算使用个人帐户 = 9     '结算时可使用个人帐户支付
    support未结清出院 = 10          '允许病人还有未结费用时出院
    
    'support门诊部分退现金 = 11      '只有在门诊医保不支持退费才使用本参数。也就是说在退现金时才考虑部分退与否，而退回到个人帐户的医保都必须整张退费。
    support允许不设置医保项目 = 12  '在结算时，不对各收费细目是否设置医保项目进行检查
    
    support门诊必须传递明细 = 13    '门诊收费和挂号是否必须传递明细
    
    support记帐上传 = 14            '住院记帐费用明细实时传输
    support记帐作废上传 = 15        '住院费用退费实时传输

    support出院病人结算作废 = 16    '允许出院病人结帐作废
    support撤销出院 = 17            '允许撤消病人出院
    support必须录入入出诊断 = 18    '病人入院与出院时，必须录入诊断名
    support记帐完成后上传 = 19      '要求上传在记帐数据提交后再进行
    support出院结算必须出院 = 20    '病人结帐时如果选择出院结帐，就检查必须出院才可以进行
    
    support挂号使用个人帐户 = 21    '使用医保挂号时是否使用个人帐户进行支付

    support门诊连续收费 = 22        '门诊在身份验证后，可进行多次收费操作
    support门诊收费完成后验证 = 23  '在门诊收费完成，是否再次调用身份验证
    
    support医嘱上传 = 24            '医嘱产生费用时是否实时传输
    support分币处理 = 25            '医保病人是否处理分币
    support中途结算仅处理已上传部分 = 26 '提供对已上传部分数据的结算功能
    support允许冲销已结帐的记帐单据 = 27 '是否允许冲销记帐单据，如果该单据已经结帐
    
    support允许部份冲销单据 = 28
    support出院无实际交易 = 29      '出院接口中是否要与接口商进行交易
    support多单据收费 = 30          '是否支持多单据收费
    
    support门诊收费存为划价单 = 31  '将门诊收费单转为划价单保存，修改以前固定判断某个医保的方式
    
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support多单据收费必须全退 = 39  '多单据收费必须全退
    
    support医保接口打印票据 = 46    'HIS中只走票据号但不调打印，医保接口(北京)中打印
    support多单据一次结算 = 47      '多单据预结算时，医保接口仅在最后一次调用时返回结算结果，HIS中再分摊到每张单据上
    
    support住院病人不受特准项目限制 = 50            '同一种病,在住院时允许录入所有的项目
    support门诊病人不受特准项目限制 = 51            '允许门诊在某种情况下可以录入所有项目
    support医生确定处方类型 = 48
    support实时监控 = 60             '是否启用费用实时监控
    
    '刘兴洪:27536 20100119
    support不提醒缴款金额不足 = 64            '在收费时,如果收费参数的"不进行缴款输入和累计控制"为true时,同时是医保病人时没有输入缴款金额时不提醒用户
    support退费后打印回单 = 65   '医保病人是否退费后打印回单:问题
End Enum

Public Sub zlinitSystemPara(Optional cnOracle As ADODB.Connection)
    '------------------------------------------------------------------------------
    '------------------------------------------------------------------------------
    '功能:初始化相关的系统参数
    '返回:填充成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objDataBase As Object, objTemp As clsDataBase
    
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDataBase = objTemp
    Else
         Set objDataBase = zlDatabase
    End If
    
    '问题:52913
    strSQL = "Select 卡号密文 From 医疗卡类别 Where 名称='就诊卡' and nvl(是否固定,0)=1"
    Set rsTemp = objDataBase.OpenSQLRecord(strSQL, "读取原就诊卡卡号密文显示规则")
    gblnShowCard = False
    If Not rsTemp.EOF Then
        gblnShowCard = Nvl(rsTemp!卡号密文) = ""
    End If
    '78773:李南春,2014-10-29,LED显示一卡通支付信息
    gblnLED = Val(GetSetting("ZLSOFT", "公共全局", "使用", 0)) <> 0
    gstrLike = IIf(Val(objDataBase.GetPara("输入匹配")) = 0, "%", "")
    With gSystemPara
        '0-拼音码,1-五笔码,2-两者
        .int简码方式 = Val(objDataBase.GetPara("简码方式"))
        .bln个性化风格 = objDataBase.GetPara("使用个性化风格") = "1"
        
        '第1位1-全数字只查编码,第2位1-全字母只查简码,在HIS基础参数中设置
        strTemp = objDataBase.GetPara(44, glngSys)
        If strTemp = "" Then strTemp = "00"
        If Len(strTemp) = 1 Then strTemp = strTemp & "0"
        .bln全数字按编码查 = Val(Left(strTemp, 1)) = 1
        .bln全字母按简码查 = Val(Mid(strTemp, 2, 1)) = 1
        '费用金额小数点位数
        gbytDec = Val(objDataBase.GetPara(9, glngSys, , 2))
        gstrDec = "0." & String(gbytDec, "0")
        '刘兴洪 问题:????    日期:2010-12-06 23:38:53
        '费用单价保留位数
        gintFeePrecision = Val(objDataBase.GetPara(157, glngSys, , "5"))
        gstrFeePrecisionFmt = "0." & String(gintFeePrecision, "0")
     End With
     gintDebug = -1
     '初如化站点信息
     Call Init站点信息: Call 初始小数位数
     Call zlInitColorSet
     Set objDataBase = Nothing
     Set objTemp = Nothing
End Sub
Public Sub 初始小数位数()
    '------------------------------------------------------------------------------------------------------
    '功能:初始小数位数
    '入参:
    '出参:
    '返回:7
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    With g_小数位数
        .成本价小数 = 7
        .零售价小数 = 7
        .金额小数 = 2
        .数量小数 = 3
        .折扣率 = 2
    End With
    With gVbFmtString
        .FM_成本价 = GetFmtString(g_成本价, False)
        .FM_金额 = GetFmtString(g_金额, False)
        .FM_零售价 = GetFmtString(g_售价, False)
        .FM_数量 = GetFmtString(g_数量, False)
        .FM_折扣率 = GetFmtString(g_折扣率, False)
    End With
    With gOraFmtString
        .FM_成本价 = GetFmtString(g_成本价, True)
        .FM_金额 = GetFmtString(g_金额, True)
        .FM_零售价 = GetFmtString(g_售价, True)
        .FM_数量 = GetFmtString(g_数量, True)
        .FM_折扣率 = GetFmtString(g_折扣率, True)
    End With
End Sub

Public Function GetFmtString(ByVal 小数类型 As g小数类型, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参: lng小数位数-小数位数
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim int位数 As Integer
    Select Case 小数类型
    Case g_数量
         int位数 = g_小数位数.数量小数
    Case g_金额
         int位数 = g_小数位数.金额小数
    Case g_成本价
         int位数 = g_小数位数.成本价小数
    Case g_售价
         int位数 = g_小数位数.零售价小数
    Case Else
        int位数 = 0
    End Select
    If blnOracle Then
       GetFmtString = "'999999999990." & String(int位数, "9") & "'"
    Else
       GetFmtString = "#0." & String(int位数, "0") & ";-#0." & String(int位数, "0") & "; ;"
    End If
End Function

Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的权限是否存在
    '参数:strPrivs-权限串
    '     strMyPriv-具体权限
    '返回,存在权限,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 14:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function
Public Function zlGet收费类别() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取收费类别
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    
    On Error GoTo errHandle
    
    gstrSQL = "Select  编码,名称 From 收费项目类别"
    If grsStatic.rs收费类别 Is Nothing Then
        Set grsStatic.rs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    ElseIf grsStatic.rs收费类别.State <> 1 Then
        Set grsStatic.rs收费类别 = zlDatabase.OpenSQLRecord(gstrSQL, "获取收费类别")
    End If
    Set zlGet收费类别 = grsStatic.rs收费类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlGet消费卡接口(Optional cnOracle As ADODB.Connection) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取消费卡接口
    '编制:刘兴洪
    '日期:2009-12-09 14:37:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '先缓存到本地
    Dim objDataBase  As Object, objTemp As clsDataBase
    On Error GoTo errHandle
    Set objDataBase = zlDatabase
    If Not cnOracle Is Nothing Then
        Set objTemp = New clsDataBase
        Call objTemp.InitCommon(cnOracle)
        Set objDataBase = objTemp
    End If
    '56615
    gstrSQL = "" & _
    "   Select 编号,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度, " & _
    "           nvl(是否退现,0) as 是否退现,nvl(是否全退,0) as 是否全退,nvl(是否刷卡,0) as 是否刷卡," & _
    "           nvl(密码长度,10) as 密码长度,nvl(密码长度限制,0) as 密码长度限制,nvl(密码规则,0) as 密码规则," & _
    "           部件,系统,是否密文,0 as 密码输入限制,0 as 是否缺省密码," & _
    "           0 as 是否模糊查找,0 as 是否制卡, 1 as 是否发卡, 0 as 是否写卡" & _
    "   From 卡消费接口目录 where nvl(启用,0)=1 "
    If grsStatic.rs消费卡接口 Is Nothing Then
        Set grsStatic.rs消费卡接口 = objDataBase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    ElseIf grsStatic.rs消费卡接口.State <> 1 Then
        Set grsStatic.rs消费卡接口 = objDataBase.OpenSQLRecord(gstrSQL, "获取消费卡接口 ")
    End If
     

    grsStatic.rs消费卡接口.Filter = 0
    Set zlGet消费卡接口 = grsStatic.rs消费卡接口
    Exit Function
errHandle:
    If Not cnOracle Is Nothing And Not objTemp Is Nothing Then
        If objTemp.ErrCenter = 1 Then Resume
        Set objTemp = Nothing: Set objDataBase = Nothing
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Set objTemp = Nothing: Set objDataBase = Nothing
End Function

Public Function zlIsCardNoShowPW(ByRef lng序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示卡号是否密文显示
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-10-25 10:31:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    If rsTemp.EOF Then Exit Function
    rsTemp.Filter = "编号=" & lng序号
    If rsTemp.EOF Then
        zlIsCardNoShowPW = False
    Else
         zlIsCardNoShowPW = Val(Nvl(rsTemp!是否密文)) = 1
    End If
    rsTemp.Filter = 0
End Function
Public Function zlCreateBrushObjects(ByVal objCard As clsCard, ByRef objBrhushCardObject As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建刷卡对象
    '入参:clsCard-卡对象
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-31 14:46:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCommpentName As String
    If objCard.启用 Then
        '检查设备是否启用
        If objCard.接口程序名 = "" Then
            '消费卡
            Set objBrhushCardObject = New clsSimulateSquareCard: zlCreateBrushObjects = True
        Else
            strCommpentName = objCard.接口程序名 & "." & "cls" & Replace(Replace(UCase(objCard.接口程序名), "ZL9", ""), "ZL", "")
            Err = 0: On Error Resume Next
            Set objBrhushCardObject = CreateObject(strCommpentName)
            If Err <> 0 Then
                ShowMsgbox "部件:" & objCard.接口编码 & "-" & objCard.名称 & "( " & strCommpentName & ")创建失败!" & vbCrLf & "详细的信息为:" & Err.Description
                Call WritLog("mdlCardSquare.zlCreateBrushObjects", "", "部件:" & objCard.接口编码 & "-" & objCard.名称 & "创建失败!详细的信息为:" & Err.Description)
                Exit Function
            End If
            zlCreateBrushObjects = True
        End If
    Else
        Set objBrhushCardObject = Nothing
    End If
End Function
Public Function zlGetCardObject(ByVal lng接口编号 As Long, ByRef objBrushCard As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据指定结算卡序号获取结算卡对象
    '入参：lng接口编号-结算卡对序号
    '出参：objCard-返回结算卡对象
    '返回：获取成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-18 11:58:54
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCardTemp As Object
    If gobjStartCards Is Nothing Then Exit Function
    
    If gobjStartCards.count = 0 Then Exit Function
    For i = 1 To gobjStartCards.count
         Err = 0: On Error Resume Next
         Set objCardTemp = gobjStartCards(i)(0)
         If Err = 0 Then
            If gobjStartCards(i)(2) = lng接口编号 Then
                Set objBrushCard = objCardTemp
                zlGetCardObject = True: Exit Function
            End If
        End If
        On Error GoTo 0
    Next
    zlGetCardObject = False
End Function

Public Function zlInitCards() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡集
    '返回:成功!返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-12-15 14:31:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, int自动读取 As Integer, bln启用 As Boolean, str部件 As String, objCard As clsCard
    Dim objBrushCards As Object, int自动间隔 As Integer
    
    Err = 0: On Error GoTo Errhand:
    Set gObjXFCards = New clsCards
    Set gobjStartCards = New Collection '格式;array(部件对象,自制卡,接口编号)
    Set rsTemp = zlGet消费卡接口
    With rsTemp
        '自制卡(即消费卡)
        .Filter = "自制卡=1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            ' "公共全局\SquareCard\" & mlngCardNo, "自动读取"
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取", "0"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "启用", "1")) = 1
            int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "自动读取间隔", "1"))
                
            str部件 = Nvl(rsTemp!部件)
            Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!编号)), Nvl(!编号), Nvl(rsTemp!名称), Left(Nvl(rsTemp!名称), 1), bln启用, True, str部件, True, 1, int自动读取, int自动间隔, Val(Nvl(rsTemp!系统)) = 1, Nvl(rsTemp!结算方式), Nvl(rsTemp!前缀文本), Val(Nvl(rsTemp!卡号长度)), True, Val(Nvl(rsTemp!是否刷卡)) = 1, False, Val(Nvl(rsTemp!是否全退)) = 1, "", "", True, Val(Nvl(rsTemp!是否密文)), Val(Nvl(rsTemp!是否退现)) = 1, Val(Nvl(rsTemp!密码长度)), Val(Nvl(rsTemp!密码长度限制)), Val(Nvl(rsTemp!密码规则)), "K" & Nvl(rsTemp!编号))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, "1", CStr(Nvl(!编号))), "K" & Nvl(!编号)
            End If
            .MoveNext
        Loop
        '银联卡
        .Filter = "自制卡<>1"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            int自动读取 = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard\" & Nvl(!编号), "自动读取", 0))
            int自动间隔 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "自动读取间隔", "1"))
            bln启用 = Val(GetSetting("ZLSOFT", "公共模块\zlSquareCard\" & Nvl(!编号), "启用", "1")) = 1
            str部件 = Nvl(rsTemp!部件)
             Set objCard = gObjXFCards.AddItem(EM_CardType_Consume, Val(Nvl(!编号)), Nvl(!编号), Nvl(rsTemp!名称), Left(Nvl(rsTemp!名称), 1), bln启用, True, str部件, False, 1, int自动读取, int自动间隔, Val(Nvl(rsTemp!系统)) = 1, Nvl(rsTemp!结算方式), Nvl(rsTemp!前缀文本), Val(Nvl(rsTemp!卡号长度)), True, Val(Nvl(rsTemp!是否刷卡)) = 1, True, Val(Nvl(rsTemp!是否全退)) = 1, "", "", True, Val(Nvl(rsTemp!是否密文)), Val(Nvl(rsTemp!是否退现)) = 1, Val(Nvl(rsTemp!密码长度)), Val(Nvl(rsTemp!密码长度限制)), Val(Nvl(rsTemp!密码规则)), "K" & Nvl(rsTemp!编号))
            If zlCreateBrushObjects(objCard, objBrushCards) Then
                gobjStartCards.Add Array(objBrushCards, 0, CStr(Nvl(!编号))), "K" & Nvl(!编号)
            End If
            .MoveNext
        Loop
    End With
    zlInitCards = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub WritLog(ByVal strDev As String, strInput As String, strOutPut As String)
    On Error GoTo errHandle
    If gDebug Then
        Open App.Path & "\SquareCard" & Format(Now(), "yyyyMMdd") & ".log" For Append As #1
        Write #1, Now
        Write #1, strDev; strInput; strOutPut
        Write #1, "======================================================================="
        Close #1
    End If
    Exit Sub
errHandle:
    MsgBox "写日志出现错误！" & vbNewLine & Err.Description, vbExclamation, "IC卡接口"
End Sub

Public Function Read模拟卡号(ByVal strFile As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从已经产生的卡号中读取一个带标志的卡号(如果有多个,以最后一个为准)
    '编制:刘兴洪
    '日期:2009-12-17 10:35:51
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim objFile As New FileSystemObject, objText As TextStream, varData As Variant
    Dim strText As String, strCardNo As String
    strCardNo = ""
    Set objText = objFile.OpenTextFile(strFile)
    Do While Not objText.AtEndOfStream
        strText = Trim(objText.ReadLine)
        varData = Split(strText, vbTab)
        If Val(varData(0)) = 1 Then
            strCardNo = varData(1)
        End If
    Loop
    objText.Close
    Read模拟卡号 = strCardNo
    Exit Function
Errhand:
End Function
Public Sub zlInitBrushCardRec(ByRef rsTemp As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化本地记录集
    '出参:返回本地结算的初化记录休
    '编制:刘兴洪
    '日期:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = adStateOpen Then .Close
        .Fields.Append "接口编号", adDouble, 18, adFldIsNullable
        .Fields.Append "消费卡ID", adDouble, 18, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "卡名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "余额", adDouble, 16, adFldIsNullable
        .Fields.Append "结算金额", adDouble, 16, adFldIsNullable
        .Fields.Append "交易时间", adDate, 50, adFldIsNullable
        .Fields.Append "交易流水号", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "备注", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "结算标志", adNumeric, 2, adFldIsNullable
        .Fields.Append "分摊页码", adLongVarChar, 600, adFldIsNullable  '多单据有效,在HIS结算后自动分配:用逗号分离,如,2,3表示,此条刷卡消费分配在第二张单据和第三张单据
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Sub zlInit收费类别Struc()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化本地记录集
    '出参:返回本地结算的初化记录休
    '编制:刘兴洪
    '日期:2009-12-23 11:22:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set grsStatic.rs收费类别汇总 = New ADODB.Recordset
    Set grsStatic.rs分单类别汇总 = New ADODB.Recordset
    
    grsStatic.dbl费用总额 = 0: grsStatic.dbl已刷累计额 = 0
    With grsStatic.rs收费类别汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
    
    With grsStatic.rs分单类别汇总
        If .State = adStateOpen Then .Close
        .Fields.Append "分类", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "单据序号", adDouble, 18, adFldIsNullable
        .Fields.Append "收费类别", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 16, adFldIsNullable
        .Fields.Append "分摊金额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With
End Sub
Public Function zlInit收费类别数据(ByVal rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据费用记录集，获取当前卡可以消费的最大额度
    '入参:rsFeeList-明细费用:
    '    字段: 费别,NO,实际票号、结算时间、病人ID、收费类别、收据费目、计算单位、开单人、收费细目ID、数量、单价、实收金额、是否急诊、开单部门ID、执行部门ID
    '出参:
    '编制:刘兴洪
    '日期:2009-12-23 16:11:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl最大消费额 As Double, str收费类别 As String, lng序号 As Long
    Err = 0: On Error GoTo Errhand:
    Call zlInit收费类别Struc
    lng序号 = 0
    With rsFeeList
        .Sort = "收费类别"
        Do While Not rsFeeList.EOF
            If str收费类别 <> Nvl(!收费类别) Then
                grsStatic.rs收费类别汇总.AddNew
                grsStatic.rs收费类别汇总!收费类别 = Nvl(!收费类别)
                str收费类别 = Nvl(!收费类别)
            End If
            grsStatic.rs收费类别汇总!实收金额 = Val(Nvl(grsStatic.rs收费类别汇总!实收金额)) + Val(Nvl(!实收金额))
            grsStatic.rs收费类别汇总.Update
            grsStatic.dbl费用总额 = grsStatic.dbl费用总额 + Val(Nvl(!实收金额))
            
            grsStatic.rs分单类别汇总.Find "分类='" & Nvl(rsFeeList!单据序号) & "_" & Nvl(!收费类别) & "'", , , 1
            If grsStatic.rs分单类别汇总.EOF Then
                grsStatic.rs分单类别汇总.AddNew
                grsStatic.rs分单类别汇总!收费类别 = Nvl(!收费类别)
                
            End If
            grsStatic.rs分单类别汇总!单据序号 = Val(Nvl(!单据序号))
            grsStatic.rs分单类别汇总!实收金额 = Val(Nvl(grsStatic.rs分单类别汇总!实收金额)) + Val(Nvl(!实收金额))
            grsStatic.rs分单类别汇总.Update
            rsFeeList.MoveNext
        Loop
    End With
    zlInit收费类别数据 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zl获取最大消费额(ByVal str限制类别 As String, ByVal dbl最大消费额 As Double, ByVal dbl已刷累计 As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取最大消费额
    '    dbl最大消费额=-1表示未传入最大消费额
    '编制:刘兴洪
    '日期:2009-12-24 10:24:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl限定金额 As Double, dbl可消费 As Double
    Err = 0: On Error GoTo Errhand:
    
    If str限制类别 <> "" Then
        str限制类别 = zlGet获取限制类别FromNameToCode(str限制类别)
    End If
    dbl限定金额 = 0
    If str限制类别 <> "" Then
        With grsStatic.rs收费类别汇总
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                If InStr(1, str限制类别, "," & Nvl(!收费类别) & ",") > 0 Then
                    dbl限定金额 = dbl限定金额 + Val(Nvl(!实收金额))
                End If
                .MoveNext
            Loop
        End With
    End If
    '计算公式:
    '最大可消费额= 总费用-冲预交-已消费额-限定金额
    dbl可消费 = dbl最大消费额 - dbl限定金额 - dbl已刷累计
    zl获取最大消费额 = IIf(dbl可消费 < 0, 0, dbl可消费)
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlGet失效面额(ByVal lng消费卡ID As Long, ByVal lng接口编号 As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取失效面额
    '返回:失效面额
    '编制:刘兴洪
    '日期:2009-12-23 15:08:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, dblTemp As Double
    Err = 0: On Error GoTo Errhand:
    gstrSQL = " " & _
    "  Select Sum(Nvl(失效金额, 0)) As 失效金额 " & _
    "  From (Select 卡面金额 As 失效金额 " & _
    "         From 消费卡目录 A " & _
    "         Where ID =  [1]  " & _
    "         Union All " & _
    "         Select -1 * Sum(Nvl(A.结算金额, 0)) As 失效金额 " & _
    "         From 病人卡结算记录 A, 消费卡目录 B " & _
    "         Where A.消费卡id = B.ID And A.消费卡id =  [1]  And A.接口编号 =  [2]  And " & _
    "               A.交易时间 <= Nvl(B.有效期, To_Date('3000-01-01', 'yyyy-mm-dd')))"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取失效额", lng消费卡ID, lng接口编号)
    dblTemp = Val(Nvl(rsTemp!失效金额))
    If dblTemp < 0 Then dblTemp = 0
    zlGet失效面额 = dblTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGet获取限制类别FromNameToCode(ByVal str限制类别 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据限制类别名称取相关的编码
    '返回:
    '编制:刘兴洪
    '日期:2009-12-23 16:31:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = zlGet收费类别
    rsTemp.Filter = 0
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    If str限制类别 = "" Then zlGet获取限制类别FromNameToCode = "": Exit Function
    str限制类别 = "," & str限制类别 & ","
    With rsTemp
        Do While Not .EOF
            str限制类别 = Replace(str限制类别, "," & Nvl(rsTemp!名称) & ",", "," & Nvl(rsTemp!编码) & ",")
            .MoveNext
        Loop
    End With
    zlGet获取限制类别FromNameToCode = str限制类别
 End Function
Public Function zl分摊结算数据(ByRef rsRquare As ADODB.Recordset, ByRef rs分摊 As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:将刷卡结果分摊结算数据给每张单据明细
    '入出参 rsRquare-(接口编号 消费卡ID,卡号,结算方式,卡名称,余额,结算金额 交易时间,备注,结算标志)
    '       rs分摊-显示每张单据分摊情况
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-06 10:13:43
    '规则说明:
    '   1.先分摊限制类别的
    '   2.再分摊不限制类别的
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strTemp As String, str限制类别 As String, dbl金额 As Double
    Dim dbl总额 As Double
    Set rs分摊 = New ADODB.Recordset
    With rs分摊
        If .State = adStateOpen Then .Close
        .Fields.Append "单据序号", adDouble, 18, adFldIsNullable
        .Fields.Append "消费卡ID", adDouble, 18, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "结算方式", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "分摊额", adDouble, 16, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic
    End With

    Set rsTemp = zlDatabase.CopyNewRec(rsRquare)
    Err = 0: On Error GoTo Errhand:
    
    '先确定，存在哪些限制类别
    rsTemp.Filter = "消费卡ID >0"
    str限制类别 = ""
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), False)
        If InStr(1, str限制类别, strTemp) <= 0 Then
            str限制类别 = str限制类别 & "," & strTemp
        End If
        rsTemp.MoveNext
    Loop
    
    rsTemp.Filter = 0
    If str限制类别 <> "" Then
        str限制类别 = zlGet获取限制类别FromNameToCode(str限制类别) & ","
    End If
    
    rsTemp.Filter = 0
    With grsStatic.rs分单类别汇总
        '先将限制类别的进行分摊
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '需要计算
            If InStr(1, str限制类别, "," & Nvl(!收费类别) & ",") > 0 Then
                '存在限制类别,先将这部分分摊掉
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!收费类别) & ",") <= 0 And Val(Nvl(rsTemp!结算金额)) > 0 Then
                      '只有用不限定的类别的分摊
                       dbl金额 = Val(Nvl(!实收金额))
                      If dbl金额 >= Val(Nvl(rsTemp!结算金额)) Then
                        dbl金额 = Val(Nvl(rsTemp!结算金额))
                        rsTemp!结算金额 = 0
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                        .Update
                      Else
                        '小的话
                        rsTemp!结算金额 = Val(Nvl(rsTemp!结算金额)) - dbl金额
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                      End If
                      rs分摊.Filter = "单据序号=" & Val(Nvl(rsTemp!单据序号)) & " And 消费卡ID=" & Val(Nvl(rsTemp!消费卡ID)) & " And 卡号='" & Nvl(rsTemp!卡号) & "'"
                      If rs分摊.EOF Then
                          rs分摊.AddNew
                      End If
                      rs分摊!单据序号 = Val(Nvl(rsTemp!单据序号))
                      rs分摊!消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                      rs分摊!卡号 = Nvl(rsTemp!卡号)
                      rs分摊!结算方式 = Trim(Nvl(rsTemp!结算方式))
                      rs分摊!分摊额 = Val(Nvl(rs分摊!分摊额)) + dbl金额
                      rs分摊.Update
                   End If
                   If !分摊金额 = !实收金额 Then Exit Do
                   rsTemp.MoveNext
                Loop
            End If
            .MoveNext
        Loop
        '再分摊不限定的
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
             If Val(Nvl(!分摊金额)) <= Val(Nvl(!实收金额)) Then
                
                rsTemp.Filter = 0
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                   strTemp = zlFromCardGet限制类别(Val(Nvl(rsTemp!消费卡ID)), True)
                   If InStr(1, strTemp, "," & Nvl(!收费类别) & ",") <= 0 And Val(Nvl(rsTemp!结算金额)) > 0 Then
                      dbl金额 = Val(Nvl(!实收金额))
                      If dbl金额 >= Val(Nvl(rsTemp!结算金额)) Then
                        dbl金额 = Val(Nvl(rsTemp!结算金额))
                        rsTemp!结算金额 = 0
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                        .Update
                      Else
                        '小的话
                        rsTemp!结算金额 = Val(Nvl(rsTemp!结算金额)) - dbl金额
                        rsTemp.Update
                        !分摊金额 = Val(Nvl(!分摊金额)) + dbl金额
                      End If
                      rs分摊.Filter = "单据序号=" & Val(Nvl(!单据序号)) & " And 消费卡ID=" & Val(Nvl(rsTemp!消费卡ID)) & " And 卡号='" & Nvl(rsTemp!卡号) & "'"
                      If rs分摊.EOF Then
                          rs分摊.AddNew
                      End If
                      rs分摊!单据序号 = Val(Nvl(!单据序号))
                      rs分摊!消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                      rs分摊!卡号 = Nvl(rsTemp!卡号)
                      rs分摊!结算方式 = Trim(Nvl(rsTemp!结算方式))
                      rs分摊!分摊额 = Val(Nvl(rs分摊!分摊额)) + dbl金额
                      rs分摊.Update
                   End If
                   If !分摊金额 = !实收金额 Then Exit Do
                   rsTemp.MoveNext
                Loop
             End If
             .MoveNext
        Loop
    End With
    
    With rs分摊
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        dbl金额 = 0
        Do While Not .EOF
            dbl金额 = dbl金额 + Val(Nvl(!分摊额))
            .MoveNext
        Loop
    End With
    dbl总额 = 0
    With rsRquare
        .Filter = 0
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            dbl总额 = dbl总额 + Val(Nvl(!结算金额))
            .MoveNext
        Loop
        If .RecordCount > 0 Then .MoveFirst
    End With
    
    If Round(dbl总额, 4) <> Round(dbl金额, 4) Then
        ShowMsgbox "多单据分摊时，出现了不等情况,请重新刷卡!"
        Exit Function
    End If
    '检查计算后的明细分摊额与总的是否一致
    zl分摊结算数据 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlFromCardGet限制类别(ByVal lng消费卡ID As Long, ByVal blnCode As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据消费卡,获取相关的限定类另
    '入参:lng消费卡ID-消费卡ID
    '     blnCode-编码
    '出参:
    '返回:返回限制类别串
    '编制:刘兴洪
    '日期:2010-01-06 11:18:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, str限制类别 As String
    Err = 0: On Error GoTo Errhand:
    gstrSQL = "Select 限制类别 From 消费卡目录 Where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取消费卡目录的限制类别", lng消费卡ID)
    If rsTemp.EOF Then Exit Function
    str限制类别 = Nvl(rsTemp!限制类别)
    If blnCode Then
        zlFromCardGet限制类别 = zlGet获取限制类别FromNameToCode(str限制类别)
    Else
        zlFromCardGet限制类别 = str限制类别
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlGetRquare(ByVal str结帐ID_IN As String, ByRef rsSquare As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡结算交易时的相关预结数据
    '入参:str结帐ID_IN-指定的结算ID
    '出参:rsSquare-结帐数据
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-15 11:08:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, lngID As Long
    
    On Error GoTo errHandle
    
    Call zlInitBrushCardRec(rsSquare)
    If str结帐ID_IN = "" Then str结帐ID_IN = "0"
    
    strSQL = "Select  /*+ rule */ Distinct A.ID, 接口编号, A.消费卡id, A.序号, A.记录状态, A.结算方式, A.结算金额, A.卡号, A.交易流水号, " & _
             "                   A.交易时间, A.备注, A.结算标志, C.结帐id " & _
             "   From 病人预交记录 C, 病人卡结算记录 A, 病人卡结算对照 B, " & _
             "        (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) J " & _
             "   Where B.预交id = C.ID And B.卡结算id = A.ID And C.结帐id = J.Column_Value And A.结算标志 = 0 And C.记录状态 = 1" & _
             " Order by ID,结帐ID"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取结帐ID的相关刷卡信息", str结帐ID_IN)
    gTy_TestBug.bln补调交易 = True
    With rsSquare
        Do While Not rsTemp.EOF
            If lngID <> Val(Nvl(rsTemp!id)) Then
                .AddNew
                !接口编号 = Val(Nvl(rsTemp!接口编号))
                !消费卡ID = Val(Nvl(rsTemp!消费卡ID))
                !卡号 = Nvl(rsTemp!卡号)
                !结算方式 = Nvl(rsTemp!结算方式)
                !卡名称 = zlGet接口名称(Val(Nvl(rsTemp!接口编号)))
                !余额 = 0
                !结算金额 = Val(Nvl(rsTemp!结算金额))
                !交易时间 = rsTemp!交易时间
                !交易流水号 = IIf(Val(Nvl(rsTemp!消费卡ID)) = 0, Nvl(rsTemp!交易流水号), Nvl(rsTemp!id))     '对于，消费卡的处理，没有特别的处理，在补传交易时，只是模拟作用。简单的更新相关的标识
                !备注 = Nvl(rsTemp!备注)
                !结算标志 = 0
            End If
            !分摊页码 = Nvl(!分摊页码) & "," & Val(Nvl(rsTemp!结帐ID))
            .Update
            rsTemp.MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    zlGetRquare = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function zlGet接口名称(ByVal lng接口编号 As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取接口名称
    '返回:接口名称
    '编制:刘兴洪
    '日期:2010-01-15 11:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As ADODB.Recordset
    Set rsTemp = zlGet消费卡接口
    rsTemp.Filter = "编号=" & lng接口编号
    If rsTemp.EOF Then
        zlGet接口名称 = ""
    Else
        zlGet接口名称 = Nvl(rsTemp!名称)
    End If
End Function
Public Function zlGet接口编号(ByVal lng预交ID As Long) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '功能：根据预交ID,获取相应的接口编号
    '返回:结算卡的接口ID
    '编制：刘兴洪
    '日期：2010-06-18 14:05:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select  distinct A.接口编号 " & _
    "   From  病人卡结算记录 A,病人卡结算对照 C" & _
    "   Where  C.卡结算ID=A.ID  and  C.预交ID=[1]    "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取退单的接口编号", lng预交ID)
    If rsTemp.RecordCount = 0 Then zlGet接口编号 = 0: Exit Function
    zlGet接口编号 = Val(Nvl(rsTemp!接口编号))
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function zlSave卡结算记录(ByVal lng预交ID As Long, ByVal strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：保存相关的结算数据
    '           用||分隔: 接口编号||消费卡ID(可传'')||结算方式||结算金额||卡号||交易流水号||交易时间(yyyy-mm-dd hh24:mi:ss)||备注
    '编制：刘兴洪
    '日期：2010-06-18 16:07:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, strSQL As String, strTemp As String
    If strBlanceInfor = "" Then Exit Function
    varData = Split(strBlanceInfor, "||")
    If UBound(varData) < 7 Then Exit Function
    
    'Zl_病人卡结算记录_Insert
    strSQL = "Zl_病人卡结算记录_Insert("
    '  接口编号_In   In 病人卡结算记录.接口编号%Type,
    strSQL = strSQL & "" & Val(varData(0)) & ","
    '  消费卡id_In   In 病人卡结算记录.消费卡id%Type,
    strSQL = strSQL & "" & IIf(Val(varData(1)) = 0, "NULL", Val(varData(1))) & ","
    '  结算方式_In   In 病人卡结算记录.结算方式%Type,
    strSQL = strSQL & "'" & Trim(varData(2)) & "',"
    '  结算金额_In   In 病人卡结算记录.结算金额%Type,
    strSQL = strSQL & "" & Val(varData(3)) & ","
    '  卡号_In       In 病人卡结算记录.卡号%Type,
    strSQL = strSQL & "'" & Trim(varData(4)) & "',"
    '  交易流水号_In In 病人卡结算记录.交易流水号%Type,
    strSQL = strSQL & "'" & Trim(varData(5)) & "',"
    '  交易时间_In   In 病人卡结算记录.交易时间%Type,
    If Trim(varData(6)) = "" Or IsDate(varData(6)) = False Then
        strTemp = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        strTemp = Trim(varData(6))
    End If
    If strTemp = "" Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "to_date('" & strTemp & "','yyyy-mm-dd hh24:mi:ss'),"
    End If
    '  备注_In       In 病人卡结算记录.备注%Type,
    strSQL = strSQL & "'" & Trim(varData(7)) & "',"
    '  结帐id_In     In Varchar2
    strSQL = strSQL & "NULL,"
    '   预交id_In     In 病人预交记录.ID%Type := -1
    strSQL = strSQL & "" & lng预交ID & ")"
    zlDatabase.ExecuteProcedure strSQL, "保存卡结算记录"
    zlSave卡结算记录 = True
End Function

Public Function zlInputIsCard(ByRef txtInput As Object, ByVal KeyAscii As Integer, ByVal lngSys As Long, Optional ByVal blnPassWd As Boolean = False) As Boolean
'功能：判断指定文本框中当前输入是否在刷卡(是否达到卡号长度，在调用程序中判断),并根据系统参数处理是否密文显示
'参数：KeyAscii=在KeyPress事件中调用的参数
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
    
     '刷卡时含有特殊符号的由调用方取消输入
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '判断是否在刷卡
    If KeyAscii > 32 Then
        sngNow = timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True   '用一台笔记本测试，一般在0.014左右
        End If
    End If
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(Not blnPassWd, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    zlInputIsCard = blnCard
End Function

Public Function zl_Get预约方式ByNo(strNO As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:根据挂号单据号获取病人预约方式
    '入参:strNo-挂号单据号
    '返回:预约方式
    '编制:王吉
    '日期:2012-07-03
    '问题号:48350
    '-----------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim str预约方式 As String
    Dim rsTemp As Recordset
    strSQL = "" & _
        "Select 预约方式 From 病人挂号记录 Where 记录状态=1 And No=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取预约方式", strNO)
    If rsTemp Is Nothing Then zl_Get预约方式ByNo = "": Exit Function
    If rsTemp.RecordCount = 0 Then zl_Get预约方式ByNo = "": Exit Function
    While rsTemp.EOF = False
        str预约方式 = rsTemp!预约方式
        rsTemp.MoveNext
    Wend
    zl_Get预约方式ByNo = str预约方式
End Function
Public Sub CreateSquareCardObject(ByRef frmMain As Object, _
    ByVal lngModule As Long, Optional cnOracle As ADODB.Connection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If gobjSquare Is Nothing Then Set gobjSquare = New SquareCard
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    If gobjSquare.objSquareCard Is Nothing Then
        Set gobjSquare.objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
        If Err <> 0 Then
            Err = 0: On Error GoTo 0:      Exit Sub
        End If
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If gobjSquare.objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDBUser, IIf(cnOracle Is Nothing, gcnOracle, cnOracle), False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub
Public Sub CloseSquareCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If gobjSquare Is Nothing Then Exit Sub
    If Not gobjSquare.objSquareCard Is Nothing Then
         'Call gobjSquare.objSquareCard.CloseWindows
         Set gobjSquare.objSquareCard = Nothing
     End If
     If Err <> 0 Then Err.Clear: Err = 0
     Set gobjSquare = Nothing
End Sub

Public Function ZVal(ByVal varValue As Variant, Optional ByVal varDefault As Variant = 0) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue) = 0, varDefault, varValue)
    ZVal = IIf(Val(varTmp) = 0, "NULL", varTmp)
End Function

Public Function zlGet支付方式(ByVal lng卡类别ID As Long, ByVal str结算方式 As String) As String
    '根据结算方式查找支付方式
    Dim strSQL As String, rsTemp As Recordset
    '名称|结算方式|是否退现|是否全退|结算性质
    zlGet支付方式 = str结算方式 & "|" & str结算方式 & "|1|0"
    On Error GoTo Errhand
    strSQL = "" & _
            " Select A.名称,A.是否退现,A.是否全退,B.性质 from 医疗卡类别 A,结算方式 B where A.结算方式 = B.名称 And A.ID = [1] And A.结算方式=[2]" & _
            " Union All " & _
            " Select A.名称,A.是否退现,A.是否全退,B.性质 from 卡消费接口目录 A,结算方式 B where A.结算方式 = B.名称 And A.编号=[1] And A.结算方式=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取支付卡结算方式", lng卡类别ID, str结算方式)
    If Not rsTemp.EOF Then
        zlGet支付方式 = Nvl(rsTemp!名称, str结算方式) & "|" & str结算方式 & "|" & Nvl(rsTemp!是否退现, 1) & "|" & Nvl(rsTemp!是否全退, 0) & "|" & Nvl(rsTemp!性质, 0)
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatiIDFromProcedure(ByVal lngModel As Long, ByVal frmParent As Object, _
                                                                ByVal strIDnumber As String, ByVal objControl As Object, _
                                                                Optional ByVal strPatiName As String = "", _
                                                                Optional ByVal strPatiSex As String = "", _
                                                                Optional ByRef blnCancel As Boolean = False) As Long
    '功能:根据病人身份证号(姓名,性别)获取病人id,病人id有可能是多个
    '入参:lngModel-模块号
    '       frmParent-显示的父窗体
    '       vRect-控件在屏幕中的位置
    '       objControl-输入身份证或刷身份证的控件
    '       strIDnumber-身份证号
    '       strPatiName-病人姓名
    '       strPatiSex-病人性别
    Dim strSQL As String, strPatiIDs As String
    Dim rsTmp  As ADODB.Recordset
    Dim vRect As RECT
    On Error GoTo Errhand
    strSQL = "Select zl_Custom_PatiIDs_Get([1],[2],[3],[4]) As 病人IDs From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, frmParent.Caption, lngModel, strIDnumber, strPatiName, strPatiSex)
    If rsTmp.EOF Then
        GetPatiIDFromProcedure = 0: Exit Function
    End If
    strPatiIDs = Nvl(rsTmp!病人IDs)
    If InStr(strPatiIDs, ",") > 0 Then
        strSQL = _
                    " Select /*+cardinality(B,10)*/ distinct A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位 " & _
                    " From 病人信息 A, Table(f_Str2List([1])) B " & _
                    " Where a.病人ID=b.Column_Value" & _
                    " Order by 姓名,性别,年龄"
        strSQL = "Select  *  From (" & strSQL & ") Where Rownum < 101"
        
        vRect = zlControl.GetControlRect(objControl.hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, objControl.Height, blnCancel, False, True, strPatiIDs)
        If Not rsTmp Is Nothing Then
            If Val(rsTmp!id) <> 0 Then GetPatiIDFromProcedure = Val(rsTmp!id)
        End If
    Else
        GetPatiIDFromProcedure = strPatiIDs
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


