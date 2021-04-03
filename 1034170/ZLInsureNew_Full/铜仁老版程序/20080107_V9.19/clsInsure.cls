VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
'------------------------------------------------
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngSys : 系统编号
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'返回：
'------------------------------------------------
    Set gcnOracle = cnMain
    gstrDbUser = strDbUser
    glngSys = lngSys
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    
    gintInsure = GetSetting("ZLSOFT", "公共全局", "医保类别", 0)
    gstrMatchMethod = GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0)
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr单位名称 = GetUnitName()
    Call GetUserInfo '获取登陆用户信息
'    If Not CheckValid Then Exit Sub

    '-------------------------------------------------
    '1600  保险类别设置
    '1601  年度结算规则
    '1602  医保大类管理
    '1603  医保项目管理
    '1604  医保帐户管理
    '1605  医保结算管理
    '1606  帐户支出管理
    '1607  医保统计报表
    
    On Error Resume Next
 
    Select Case lngModul
        Case 1600
            frm医保类别.Show , frmMain
        Case 1601
            frm结算规则.ShowForm frmMain
'            frm保险报销政策.ShowForm frmMain
        Case 1602
            If gintInsure = type_大连市 Or gintInsure = type_大连开发区 Then
                frm保险大类_大连.ShowForm frmMain
            Else
                frm保险大类.ShowForm frmMain
            End If
        Case 1603
            If gintInsure = TYPE_广元旺苍 Then
                frm保险项目_广元旺苍.ShowForm frmMain
            Else
                frm保险项目.ShowForm frmMain
            End If
            
        Case 1604
            frm医保帐户.ShowForm frmMain
        Case 1605
            frm医保结算.ShowForm frmMain
        Case 1606
            frm保险病种.ShowForm frmMain
        Case 1607
            frm医保帐户收支管理.Show , frmMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Public Sub InitOracle(cnOracle As ADODB.Connection)
'功能：初始化医保部件的本地Oracle连接
'说明：用于某些不需要连接医保前置机,但又要调用相关接口的场合,如住院记帐
'      如果调用于InitInsure,则不必再调用本函数
    Set gcnOracle = cnOracle
End Sub

Public Function InitInsure(cnOracle As ADODB.Connection, Optional intType As Integer) As Boolean
'功能：传递应用部件已经建立的ORacle连接，同时根据配置信息建立与医保服务器的连接。
'参数: cnOracle -应用部件中已经建立的到ORacle的连接
'      intType=保险类别,在系统参数中设置时需要传递保险类别
'返回：初始化成功，返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-基本,1121-基本,1131-基本,1132-基本,1133-基本,1134-基本,1135-基本,1137-基本
    Dim intTmp As Integer
    Dim rsTmp As New ADODB.Recordset
    
    Set gcnOracle = cnOracle
    
    gintInsure = GetSetting("ZLSOFT", "公共全局", "医保类别", 0)
        
    If gintInsure = 0 Then
        MsgBox "不能确定保险类别,无法执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '系统名
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    Call GetUserInfo
    
    '获取医院编码
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号=" & gintInsure
    Call OpenRecordset(rsTmp, "医保接口")
    
    If rsTmp.EOF = True Then
        MsgBox "本地选择的医保类型已经不存在，请重新设置！", vbExclamation, gstrSysName
        Exit Function
    End If
    If IsNull(rsTmp("医院编码")) = True Then
        MsgBox "由于未设置医院编号，无法执行医保交易！", vbExclamation, gstrSysName
        Exit Function
    End If
    gstr医院编码 = rsTmp!医院编码
    
    intTmp = IIf(intType = 0, gintInsure, intType)
    
    Select Case intTmp
        Case TYPE_成都市农医
            InitInsure = True
        Case TYPE_宁海
            InitInsure = 医保初始化_宁海
        Case TYPE_余姚
            InitInsure = 医保初始化_余姚()
        Case TYPE_浙江
            InitInsure = 医保初始化_浙江()
        Case TYPE_吉林
            '刘兴宏:20040923加入
            InitInsure = 医保初始化_吉林()
        Case TYPE_北京尚洋
            InitInsure = 医保初始化_北京尚洋
        Case TYPE_华东
            InitInsure = 医保初始化_华东
        Case TYPE_南京市
            InitInsure = 医保初始化_南京市
        Case TYPE_江苏
            InitInsure = 医保初始化_江苏
        Case TYPE_重庆市
            InitInsure = 医保初始化_重庆
        Case TYPE_涪陵
            InitInsure = 医保初始化_涪陵
        Case TYPE_广元
            InitInsure = 医保初始化_广元
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            InitInsure = 医保初始化_松藻(intTmp)
        Case TYPE_重庆壁山
            InitInsure = 医保初始化_壁山
        Case TYPE_重庆中梁山
            InitInsure = 医保初始化_中梁山
        Case TYPE_昆明市, TYPE_云南省
            InitInsure = 医保初始化_云南
        Case TYPE_云南建水
            InitInsure = 医保初始化_云南建水
        Case TYPE_自贡市
            InitInsure = 医保初始化_中软
        Case TYPE_泸州市
            InitInsure = 医保初始化_泸州
        Case TYPE_铜仁
            InitInsure = 医保初始化_铜仁
        Case TYPE_贵阳市
            InitInsure = 医保初始化_贵阳
        Case TYPE_成都市
            InitInsure = 医保初始化_成都
        Case TYPE_成都莲合
            InitInsure = 医保初始化_莲合
        Case TYPE_新都
            InitInsure = 医保初始化_新都
        Case TYPE_成都郊县
            InitInsure = 医保初始化_成都郊县
        Case TYPE_成都南充
            InitInsure = 医保初始化_成都南充
        Case TYPE_咸阳市
            InitInsure = 医保初始化_咸阳
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            InitInsure = 医保初始化_福建巨龙
        'Modified by zyb #2003-10-08
        Case type_米易
            InitInsure = 医保初始化_米易
        Case type_大连市, type_大连开发区
            '刘兴宏200311
            InitInsure = 医保初始化_大连
        Case TYPE_四川眉山
            InitInsure = 医保初始化_眉山
        Case TYPE_沈阳市
            InitInsure = 医保初始化_沈阳市
        Case TYPE_乐山
            InitInsure = 医保初始化_乐山
        Case TYPE_重大校园卡
            '刘兴宏:200403
            InitInsure = 医保初始化_重大校园卡
        Case TYPE_重庆银海版
            InitInsure = 医保初始化_重庆银海版
        Case TYPE_北京
            InitInsure = 医保初始化_北京
        Case TYPE_重庆渝北
            '刘兴宏:20040705
            InitInsure = 医保初始化_重庆渝北
        Case TYPE_兴安
            '刘兴宏:20040827
            InitInsure = 医保初始化_兴安
        Case TYPE_毕节
            InitInsure = 医保初始化_毕节
        Case TYPE_黔南
            '刘兴宏:20041022
            InitInsure = 医保初始化_黔南
        Case TYPE_成都德阳
            '刘兴宏:20041103
            InitInsure = 医保初始化_成都德阳
        Case TYPE_成都内江
            '刘兴宏:20041207
            InitInsure = 医保初始化_成都内江
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            InitInsure = 医保初始化_广元旺苍
        Case TYPE_渝北农医
            InitInsure = 医保初始化_渝北农医
        Case Is > 900
            '中联医保
            InitInsure = 医保初始化_中联(intTmp)
    End Select
End Function

Public Function EndInsure() As Boolean
'功能：断开与医保系统的连接
    Select Case gintInsure
        Case TYPE_宁海
            Call 医保终止_宁海
        Case TYPE_余姚
            EndInsure = 医保终止_余姚()
        Case TYPE_云南省, TYPE_昆明市
            Call yh_quit
        Case TYPE_云南建水
            Call yh_quit
        Case TYPE_重庆壁山
'            gcn壁山.Close
            Set gcn壁山 = Nothing
        Case TYPE_贵阳市
            '不需要完成特别的操作
        Case TYPE_自贡市, TYPE_泸州市, TYPE_铜仁
            '不需要完成特别的操作
        Case TYPE_涪陵, TYPE_南京市, TYPE_江苏
            '不需要完成特别的操作
        Case TYPE_华东
            Set gcn华东 = Nothing
        Case TYPE_北京尚洋
            Set gcn尚洋 = Nothing
        Case TYPE_广元
            '不需要完成特别的操作
        Case TYPE_成都市
            Set gcnSybase = Nothing
        Case TYPE_成都莲合
            Set gcnSybase = Nothing
        Case TYPE_成都南充
            Set gcnInterbase = Nothing
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            EndInsure = 医保终止_福建巨龙
        Case type_米易
            EndInsure = 医保终止_米易
        Case TYPE_四川眉山
            EndInsure = 医保终止_眉山
        Case TYPE_沈阳市
            EndInsure = 医保终止_沈阳市
        Case type_大连市, type_大连开发区
            EndInsure = 医保终止_大连
        Case TYPE_乐山
            EndInsure = 医保终止_乐山
        Case TYPE_重大校园卡
            '刘兴宏:200403
            EndInsure = 医保终止_重大校园卡
        Case TYPE_重庆银海版
            EndInsure = 医保终止_重庆银海版
        Case TYPE_重庆渝北
            '刘兴宏:20040705
            EndInsure = 医保终止_重庆渝北
        Case TYPE_兴安
            '刘兴宏:20040827
            EndInsure = 医保终止_兴安
        Case TYPE_毕节
            EndInsure = 医保终止_毕节
        Case TYPE_黔南
            '刘兴宏:20041022
            EndInsure = 医保终止_黔南
        Case TYPE_成都德阳
            '刘兴宏:20041103
            EndInsure = 医保终止_成都德阳
        Case TYPE_成都内江
            '刘兴宏:20041207
            EndInsure = 医保终止_成都内江
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            EndInsure = 医保终止_广元旺苍
        Case TYPE_渝北农医
            EndInsure = 医保终止_渝北农医
    End Select
    EndInsure = True
End Function

Public Function Identify2(strCard As String, strPass As String, Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：strCard=要验证的卡号,在调用程序中刷卡产生
'      strPass=要验证的密码,在调用程序中输入
'      bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'      lng病人ID=对病人进行补充验证时使用
'========================================================================================================
'返回：空或：
'       0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
'       9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
'       19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
'========================================================================================================
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空填充；
'------------------------------------------------------------------------------------------------------------------
'调用模块：1115-自助挂号

    Dim intMouse As Integer
    
    If strCard = "" Then Exit Function
    
    intMouse = Screen.MousePointer
    
    Select Case gintInsure
        Case TYPE_吉林
            '20040923刘兴宏加入
            Identify2 = ""
        Case TYPE_江苏, TYPE_南京, TYPE_余姚, TYPE_华东
            Identify2 = ""
        Case TYPE_重庆市
            Identify2 = ""
        Case TYPE_重庆松藻
            Identify2 = ""
        Case TYPE_自贡市
            Identify2 = ""
        Case TYPE_泸州市
            Identify2 = 身份标识_泸州2(strCard, strPass, lng病人ID)
        Case TYPE_铜仁
            Identify2 = ""
        Case TYPE_贵阳市
            Identify2 = ""
        Case TYPE_云南省, TYPE_昆明市
            Identify2 = ""
        Case TYPE_云南建水
            Identify2 = ""
        Case TYPE_成都市
            Identify2 = 身份标识_成都2(strCard, strPass, IIf(bytType = 3, 0, IIf(bytType = 4, 1, bytType)), lng病人ID)
        Case TYPE_成都莲合
            Identify2 = ""
        Case TYPE_成都郊县, TYPE_新都
            Identify2 = ""
        Case TYPE_成都南充
            Identify2 = ""
        Case TYPE_咸阳市
            Identify2 = ""
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            Identify2 = 身份标识_福建巨龙(操作方式.挂号, lng病人ID)
        Case type_米易
            Identify2 = ""
        Case TYPE_四川眉山
            Identify2 = ""
        Case TYPE_沈阳市
            Identify2 = ""
        Case type_大连市, type_大连开发区
            '刘兴宏(200403)
            Identify2 = 身份标识_大连2(strCard, strPass, lng病人ID)
        Case TYPE_涪陵
            Identify2 = ""
        Case TYPE_广元
            Identify2 = ""
        Case TYPE_乐山
            Identify2 = ""
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            Identify2 = 身份标识_重大校园卡2(strCard, strPass, lng病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            Identify2 = ""
        Case TYPE_重大校园卡
            '刘兴宏(20040705)
            Identify2 = ""
        Case TYPE_黔南
            '刘兴宏:20041022
            Identify2 = ""
        Case TYPE_成都德阳
            '刘兴宏:20041103
            Identify2 = ""
        Case TYPE_成都内江
            '刘兴宏:20041207
            Identify2 = ""
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            Identify2 = ""
        Case TYPE_渝北农医
            Identify2 = ""
        Case Is > 900
            Identify2 = ""
    End Select
    
    '还原鼠标形状，因为接口很容易中间退出，此时尚未恢复
    Screen.MousePointer = intMouse
End Function

Public Function Identify(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
'功能：识别指定人员是否为参保病人，返回病人的信息
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'========================================================================================================
'返回：空或：
'       0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8《病人ID》
'       9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
'       19帐户增加累计,20帐户支出累计,21进入统筹累计,22统筹报销累计,23住院次数累计;24就诊类型 (1、急诊门诊);25开单科室
'========================================================================================================
'注意：1)主要利用接口的身份识别交易；
'      2)如果识别错误，在此函数内直接提示错误信息；
'      3)识别正确，而个人信息缺少某项，必须以空填充；
'------------------------------------------------------------------------------------------------------------------
'调用模块：1121-门诊收费,1131-办理登记,1137-结帐办理
    Dim intMouse As Integer
    
    If bytType = 3 Then
        '挂号不能进行身份验证
        Select Case gintInsure
            'Modified By 朱玉宝 2004-05-25 原因：医保接口变动
            '------------------------------------------------
            Case TYPE_成都市, TYPE_福建巨龙, TYPE_南平市, TYPE_福建省, TYPE_江苏, _
            TYPE_福州市, TYPE_四川眉山, TYPE_泸州市, TYPE_重庆壁山, TYPE_沈阳市, _
            TYPE_重大校园卡, type_大连开发区, type_大连市, TYPE_贵阳市, TYPE_重庆市, _
            TYPE_重庆渝北, TYPE_宁海, TYPE_渝北农医
            '------------------------------------------------
            Case Else
                Exit Function
        End Select
    End If
    
    intMouse = Screen.MousePointer
    
    Select Case gintInsure
        Case TYPE_宁海
            Identify = 身份标识_宁海(bytType, lng病人ID)
        Case TYPE_余姚
            Identify = 身份标识_余姚(bytType, lng病人ID)
        Case TYPE_浙江
            Identify = 身份标识_浙江(bytType, lng病人ID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            Identify = 身份标识_吉林(bytType, lng病人ID)
        Case TYPE_北京尚洋
            Identify = 身份标识_北京尚洋(bytType, lng病人ID)
        Case TYPE_华东
            Identify = 身份标识_华东(bytType, lng病人ID)
        Case TYPE_南京市
            Identify = 身份标识_南京市(bytType, lng病人ID)
        Case TYPE_江苏
            Identify = 身份标识_江苏(bytType, lng病人ID)
        Case TYPE_涪陵
            Identify = 身份标识_涪陵(bytType, lng病人ID)
        Case TYPE_广元
            Identify = 身份标识_广元(bytType, lng病人ID)
        Case TYPE_重庆市
            Identify = 身份标识_重庆(bytType, lng病人ID)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            Identify = 身份标识_松藻(bytType, lng病人ID)
        Case TYPE_重庆壁山
            Identify = 身份标识_壁山(bytType, lng病人ID)
        Case TYPE_重庆中梁山
            Identify = 身份标识_中梁山(bytType, lng病人ID)
        Case TYPE_自贡市
            '直接调用子函数
            Identify = 身份标识_中软(bytType, lng病人ID)
        Case TYPE_泸州市
            '直接调用子函数
            Identify = 身份标识_泸州(bytType, lng病人ID)
        Case TYPE_铜仁
            '直接调用子函数
            Identify = 身份标识_铜仁(bytType, lng病人ID)
        Case TYPE_贵阳市
            Identify = 身份标识_贵阳(bytType, lng病人ID)
        Case TYPE_云南省, TYPE_昆明市
            Identify = 身份标识_云南(bytType, lng病人ID)
        Case TYPE_云南建水
            Identify = 身份标识_云南建水(bytType, lng病人ID)
        Case TYPE_成都市
            Identify = 身份标识_成都(IIf(bytType = 3, 0, IIf(bytType = 4, 1, bytType)), lng病人ID)
        Case TYPE_成都莲合
            Identify = 身份标识_莲合(bytType, lng病人ID)
        Case TYPE_新都
            Identify = 身份标识_新都(bytType, lng病人ID)
        Case TYPE_成都郊县
            Identify = 身份标识_成都郊县(bytType, lng病人ID)
        Case TYPE_成都南充
            Identify = 身份标识_成都南充(bytType, lng病人ID)
        Case TYPE_咸阳市
            Identify = 身份标识_咸阳(bytType, lng病人ID)
        Case type_米易
            Identify = 身份标识_米易(bytType, lng病人ID)
        Case TYPE_四川眉山
            Identify = 身份标识_眉山(bytType, lng病人ID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            Identify = ""
            If InStr(1, "0123", bytType) = 0 Then Exit Function
            Identify = 身份标识_福建巨龙(IIf(bytType = 0, 操作方式.收费, IIf(bytType = 1, 操作方式.入院, IIf(bytType = 3, 操作方式.挂号, 操作方式.验证))), lng病人ID)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            Identify = 身份标识_大连(bytType, lng病人ID)
        Case TYPE_沈阳市
            Identify = 身份标识_沈阳市(bytType, lng病人ID)
        Case TYPE_乐山
            Identify = 身份标识_乐山(bytType, lng病人ID)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            Identify = 身份标识_重大校园卡(bytType, lng病人ID)
        Case TYPE_重庆银海版
            Identify = 身份标识_重庆银海版(bytType, lng病人ID)
        Case TYPE_北京
            Identify = 身份标识_北京(bytType, lng病人ID)
        Case TYPE_重庆渝北
            '刘兴宏(20040705)
            Identify = 身份标识_重庆渝北(bytType, lng病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            Identify = 身份标识_兴安(bytType, lng病人ID)
        Case TYPE_毕节
            Identify = 身份标识_毕节(bytType, lng病人ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            Identify = 身份标识_黔南(bytType, lng病人ID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            Identify = 身份标识_成都德阳(bytType, lng病人ID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            Identify = 身份标识_成都内江(bytType, lng病人ID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            Identify = 身份标识_广元旺苍(bytType, lng病人ID)
        Case TYPE_渝北农医
            Identify = 身份标识_渝北农医(bytType, lng病人ID)
        Case Is > 900 '中联医保
            Identify = 身份标识_中联(bytType, lng病人ID)
    End Select
    
    '还原鼠标形状，因为接口很容易中间退出，此时尚未恢复
    Screen.MousePointer = intMouse
End Function

Public Function SelfBalance(病人ID As Long, str医保号 As String, Optional bytFlag As Byte, Optional cur透支额 As Currency = 0) As Currency
'功能: 提取参保病人个人帐户余额
'参数: bytFlag
'          个位表示余额类型：0-所有余额,1-本年余额,2-往年余额
'          十位表示调用位置：10-门诊,20-入院,30-预交,40-结算
'返回: 返回个人帐户余额的金额
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交收款,1121-门诊收费,1131-办理登记,1137-结帐办理
    Dim bytYear As Byte, bytPlace As Byte
    
    Screen.MousePointer = vbHourglass
    
    bytYear = bytFlag Mod 10        '得到个位
    bytPlace = (bytFlag \ 10) * 10  '得到十位
    
    Select Case gintInsure
        Case TYPE_宁海
            SelfBalance = 个人余额_宁海(病人ID)
        Case TYPE_余姚
            SelfBalance = 个人余额_余姚(病人ID)
        Case TYPE_浙江
            SelfBalance = 个人余额_浙江(病人ID)
            cur透支额 = 50000
        Case TYPE_吉林
            '刘兴宏:20040925
            SelfBalance = 个人余额_吉林(病人ID)
        Case TYPE_北京尚洋
            SelfBalance = 个人余额_北京尚洋(病人ID)
        Case TYPE_华东
            SelfBalance = 个人余额_华东(病人ID)
        Case TYPE_江苏
            SelfBalance = 个人余额_江苏(病人ID)
        Case TYPE_南京市
            SelfBalance = 个人余额_南京市(病人ID)
        Case TYPE_涪陵
            SelfBalance = 个人余额_涪陵(病人ID)
        Case TYPE_广元
            SelfBalance = 个人余额_广元(病人ID)
        Case TYPE_重庆市
            SelfBalance = 个人余额_重庆(str医保号)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            SelfBalance = 个人余额_松藻(病人ID)
        Case TYPE_重庆壁山
            SelfBalance = 个人余额_壁山(CStr(病人ID))
        Case TYPE_重庆中梁山
            SelfBalance = 个人余额_中梁山(CStr(病人ID))
        Case TYPE_自贡市
            '读IC卡上的余额，可能需要装钱
            SelfBalance = 个人余额_中软(病人ID)
        Case TYPE_泸州市
            '读IC卡上的余额，可能需要装钱
            SelfBalance = 个人余额_泸州(病人ID)
        Case TYPE_铜仁
            '读IC卡上的余额，可能需要装钱
            SelfBalance = 个人余额_铜仁(病人ID)
        Case TYPE_贵阳市
            SelfBalance = 个人余额_贵阳(str医保号)
        Case TYPE_云南省, TYPE_昆明市
            SelfBalance = 个人余额_云南(str医保号, bytPlace)
        Case TYPE_云南建水
            SelfBalance = 个人余额_云南建水(str医保号, bytPlace)
        Case TYPE_成都市
            SelfBalance = 个人余额_成都(str医保号, bytYear)
        Case TYPE_成都莲合
            SelfBalance = 个人余额_莲合()
        Case TYPE_新都
            SelfBalance = 个人余额_新都(str医保号, bytPlace)
        Case TYPE_成都郊县
            SelfBalance = 个人余额_成都郊县(str医保号, bytPlace)
        Case TYPE_咸阳市
            SelfBalance = 个人余额_咸阳(病人ID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            SelfBalance = 个人余额_福建巨龙(病人ID)
        Case type_米易
            SelfBalance = 个人余额_米易(True, "", IIf(bytFlag = 10, False, True))
        Case TYPE_四川眉山
            SelfBalance = 个人余额_眉山(str医保号)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            SelfBalance = 个人余额_大连(病人ID)
        Case TYPE_沈阳市
            SelfBalance = 个人余额_沈阳市(str医保号)
        Case TYPE_乐山
            SelfBalance = 个人余额_乐山(str医保号)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            SelfBalance = 个人余额_重大校园卡(病人ID, bytPlace)
        Case TYPE_重庆银海版
            SelfBalance = 个人余额_重庆银海版(str医保号)
        Case TYPE_重庆渝北
            SelfBalance = 个人余额_重庆渝北(病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            SelfBalance = 个人余额_兴安(病人ID)
        Case TYPE_北京
            selfbalace = 个人余额_北京(病人ID)
        Case TYPE_毕节
            SelfBalance = 个人余额_毕节(病人ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            SelfBalance = 个人余额_黔南(病人ID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            SelfBalance = 个人余额_成都德阳(病人ID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            SelfBalance = 个人余额_成都内江(病人ID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            SelfBalance = 个人余额_广元旺苍(病人ID)
        Case TYPE_渝北农医
            SelfBalance = 个人余额_渝北农医(str医保号)
        Case Is > 900
            '中联医保
            SelfBalance = 个人余额_中联(病人ID)
    End Select
    Screen.MousePointer = vbDefault
End Function

Public Function ClinicPreSwap(rsDetail As ADODB.Recordset, str结算方式 As String) As Boolean
'功能：将门诊收费的明细进行计算,以得出病人的报销金额
'参数：rsDetail     费用明细(传入)
'      cur结算方式  "报销方式;金额;是否允许修改|...."
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'------------------------------------------------------------------------------------------------------------------
'调用模块：1121-门诊收费
    Select Case gintInsure
        Case TYPE_宁海
            ClinicPreSwap = 门诊虚拟结算_宁海(rsDetail, str结算方式)
        Case TYPE_余姚
            ClinicPreSwap = 门诊虚拟结算_余姚(rsDetail, str结算方式)
        Case TYPE_浙江
            ClinicPreSwap = 门诊虚拟结算_浙江(rsDetail, str结算方式)
        Case TYPE_华东
            ClinicPreSwap = 门诊虚拟结算_华东(rsDetail, str结算方式)
        Case TYPE_南京市
            ClinicPreSwap = 门诊虚拟结算_南京市(rsDetail, str结算方式)
        Case TYPE_江苏
            ClinicPreSwap = 门诊虚拟结算_江苏(rsDetail, str结算方式)
        Case TYPE_涪陵
            ClinicPreSwap = 门诊虚拟结算_涪陵(rsDetail, str结算方式)
        Case TYPE_广元
            ClinicPreSwap = 门诊虚拟结算_广元(rsDetail, str结算方式)
        Case TYPE_重庆市
            ClinicPreSwap = 门诊虚拟结算_重庆(rsDetail, str结算方式)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ClinicPreSwap = 门诊虚拟结算_松藻(rsDetail, str结算方式)
        Case TYPE_重庆壁山
            ClinicPreSwap = 门诊虚拟结算_壁山(rsDetail, str结算方式)
        Case TYPE_重庆中梁山
            ClinicPreSwap = False
        Case TYPE_自贡市, Is > 900
            ClinicPreSwap = 门诊虚拟结算(rsDetail, str结算方式)
        Case TYPE_泸州市
            ClinicPreSwap = 门诊虚拟结算_泸州(rsDetail, str结算方式)
        Case TYPE_铜仁
            ClinicPreSwap = 门诊虚拟结算_铜仁(rsDetail, str结算方式)
        Case TYPE_贵阳市
            ClinicPreSwap = 门诊虚拟结算_贵阳(rsDetail, str结算方式)
        Case TYPE_云南省, TYPE_昆明市
            ClinicPreSwap = 门诊虚拟结算_云南(rsDetail, str结算方式)
        Case TYPE_云南建水
            ClinicPreSwap = 门诊虚拟结算_云南建水(rsDetail, str结算方式)
        Case TYPE_成都市
            ClinicPreSwap = False
        Case TYPE_成都莲合
            ClinicPreSwap = False
        Case TYPE_新都
            ClinicPreSwap = 门诊虚拟结算_新都(rsDetail, str结算方式)
        Case TYPE_成都郊县
            ClinicPreSwap = 门诊虚拟结算_成都郊县(rsDetail, str结算方式)
        Case TYPE_咸阳市
            ClinicPreSwap = 门诊虚拟结算_咸阳(rsDetail, str结算方式)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ClinicPreSwap = 门诊虚拟结算_福建巨龙(rsDetail, str结算方式)
        Case type_米易
            ClinicPreSwap = 门诊虚拟结算_米易(rsDetail, str结算方式)
        Case TYPE_四川眉山
            ClinicPreSwap = 门诊虚拟结算_眉山(rsDetail, str结算方式)
        Case TYPE_沈阳市
            ClinicPreSwap = 门诊虚拟结算_沈阳市(rsDetail, str结算方式)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            ClinicPreSwap = 门诊虚拟结算_大连(rsDetail, str结算方式)
        Case TYPE_乐山
            ClinicPreSwap = 门诊虚拟结算_乐山(rsDetail, str结算方式)
        Case TYPE_重大校园卡
              '刘兴宏(20040705)
            ClinicPreSwap = 门诊虚拟结算_重大校园卡(rsDetail, str结算方式)
        Case TYPE_重庆银海版
            ClinicPreSwap = 门诊虚拟结算_重庆银海版(rsDetail, str结算方式)
        Case TYPE_重庆渝北
            ClinicPreSwap = 门诊虚拟结算_重庆渝北(rsDetail, str结算方式)
        Case TYPE_兴安
            '刘兴宏:20040827
            ClinicPreSwap = 门诊虚拟结算_兴安(rsDetail, str结算方式)
        Case TYPE_北京
            ClinicPreSwap = 门诊虚拟结算_北京(rsDetail, str结算方式)
        Case TYPE_毕节
            ClinicPreSwap = 门诊虚拟结算_毕节(rsDetail, str结算方式)
        Case TYPE_黔南
            '刘兴宏:20041022
            ClinicPreSwap = 门诊虚拟结算_黔南(rsDetail, str结算方式)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ClinicPreSwap = 门诊虚拟结算_成都德阳(rsDetail, str结算方式)
        Case TYPE_成都内江
            '刘兴宏:20041207
            ClinicPreSwap = 门诊虚拟结算_成都内江(rsDetail, str结算方式)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ClinicPreSwap = 门诊虚拟结算_广元旺苍(rsDetail, str结算方式)
        Case TYPE_渝北农医
            ClinicPreSwap = 门诊虚拟结算_渝北农医(rsDetail, str结算方式)
    End Select
    '还原鼠标形状，因为接口很容易中间退出，此时尚未恢复
    Screen.MousePointer = intMouse
End Function

Public Function ClinicSwap(lng结帐ID As Long, cur个帐支付 As Currency, cur医保基金 As Currency, cur全自付 As Currency, cur先自付 As Currency) As Boolean
'功能：将门诊收费的明细和结算数据转发送医保前置服务器确认；
'参数：lng结帐ID-收费记录的结帐ID；，从预交记录中可以检索医保号和密码
'
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。
'但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，
'但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。'
'这样才能保证数据的完全统一
'------------------------------------------------------------------------------------------------------------------
'调用模块：1121-门诊收费
    Dim str医保号 As String, str密码 As String
    Dim lng病人ID As Long, str卡号 As String
    Dim rsTmp As New ADODB.Recordset
    
    With rsTmp
        'Modified by ZYB 2002-10-28
        '松藻医保不需要对每个项目设置对应的医保项目
        If Not GetCapability(support允许不设置医保项目) Then
            gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=" & gintInsure & ") A," & _
                " (Select 收费细目ID,'Y' as TEST From 病人费用记录 Where 结帐ID=" & lng结帐ID & " And Nvl(附加标志,0)<>9) B,收费细目 C" & _
                " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
            Call OpenRecordset(rsTmp, "医保接口")
            
            If Not .EOF Then
                MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                ClinicSwap = False: Exit Function
            End If
        End If
        
        'Modified by 朱玉宝（考虑到医保某次结算，不一定存在个人帐户和医保基金的情况，为了保证仍然调用医保接口）
'        If GetCapability(support门诊必须传递明细) = False Then
'            gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
'                " From 病人预交记录 A,保险帐户 B " & _
'                " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
'                " And A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lng结帐ID
'        Else
            '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
            gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                " From 病人预交记录 A,保险帐户 B " & _
                " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
                "       And A.结帐ID=" & lng结帐ID
'        End If
        Call OpenRecordset(rsTmp, "医保接口")
        
        If .RecordCount = 0 Then
            MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
            ClinicSwap = True: Exit Function
        End If
    End With
    
    lng病人ID = rsTmp!病人ID
    str卡号 = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
    str医保号 = TrimStr(IIf(IsNull(rsTmp!医保号), str卡号, rsTmp!医保号))
    str密码 = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
    
    Select Case gintInsure
        Case TYPE_宁海
            ClinicSwap = 门诊结算_宁海(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_余姚
            ClinicSwap = 门诊结算_余姚(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_浙江
            ClinicSwap = 门诊结算_浙江(lng结帐ID, cur个帐支付, str医保号, cur全自付, cur先自付, cur医保基金)
        Case TYPE_吉林
            '20040923刘兴宏加入
            ClinicSwap = 门诊结算_吉林(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_北京尚洋
            ClinicSwap = 门诊结算_北京尚洋(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_华东
            ClinicSwap = 门诊结算_华东(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_南京市
            ClinicSwap = 门诊结算_南京市(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_江苏
            ClinicSwap = 门诊结算_江苏(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_涪陵
            ClinicSwap = 门诊结算_涪陵(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_广元
            ClinicSwap = 门诊结算_广元(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_重庆市
            ClinicSwap = 门诊结算_重庆(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ClinicSwap = 门诊结算_松藻(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付)
        Case TYPE_重庆壁山
            ClinicSwap = 门诊结算_壁山(lng结帐ID, cur个帐支付, str医保号, cur全自付, cur先自付, cur医保基金)
        Case TYPE_重庆中梁山
            ClinicSwap = 门诊结算_中梁山(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付)
        Case TYPE_自贡市
            ClinicSwap = 门诊结算_中软(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
        Case TYPE_泸州市
            ClinicSwap = 门诊结算_泸州(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
        Case TYPE_铜仁
            ClinicSwap = 门诊结算_铜仁(lng结帐ID, cur个帐支付, cur全自付, cur先自付)
        Case TYPE_贵阳市
            ClinicSwap = 门诊结算_贵阳(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_云南省, TYPE_昆明市
            ClinicSwap = 门诊结算_云南(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_云南建水
            ClinicSwap = 门诊结算_云南建水(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_成都市
            ClinicSwap = 门诊结算_成都(lng结帐ID, lng病人ID, str医保号, str密码, str卡号, cur全自付)
        Case TYPE_成都莲合
            ClinicSwap = 门诊结算_莲合(lng结帐ID)
        Case TYPE_成都郊县
            ClinicSwap = 门诊结算_成都郊县(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_新都
            ClinicSwap = 门诊结算_新都(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_咸阳市
            ClinicSwap = 门诊结算_咸阳(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ClinicSwap = 门诊结算_福建巨龙(lng病人ID, lng结帐ID)
        Case type_米易
            ClinicSwap = 门诊结算_米易(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_四川眉山
            ClinicSwap = 门诊结算_眉山(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_沈阳市
            ClinicSwap = 门诊结算_沈阳市(lng结帐ID, cur个帐支付, str医保号)
        Case type_大连市, type_大连开发区
            '200311
            ClinicSwap = 门诊结算_大连(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_乐山
            ClinicSwap = 门诊结算_乐山(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_重庆银海版
            ClinicSwap = 门诊结算_重庆银海版(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_重大校园卡
            '刘兴宏200403
            ClinicSwap = 门诊结算_重大校园卡(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_重庆渝北
            '刘兴宏20040705
            ClinicSwap = 门诊结算_重庆渝北(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_北京
            ClinicSwap = 门诊结算_北京(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_兴安
            '刘兴宏:20040827
            ClinicSwap = 门诊结算_兴安(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_毕节
            ClinicSwap = 门诊结算_毕节(lng结帐ID, cur个帐支付, str医保号)
        Case TYPE_黔南
            '刘兴宏:20041022
            ClinicSwap = 门诊结算_黔南(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ClinicSwap = 门诊结算_成都德阳(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_成都内江
            '刘兴宏:20041207
            ClinicSwap = 门诊结算_成都内江(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ClinicSwap = 门诊结算_广元旺苍(lng结帐ID, cur个帐支付, str医保号, cur全自付)
        Case TYPE_渝北农医
            ClinicSwap = 门诊结算_渝北农医(lng结帐ID, cur个帐支付, str医保号)
        Case Is > 900
            ClinicSwap = 门诊结算_中联(lng结帐ID, cur个帐支付, lng病人ID, cur全自付, cur先自付)
    End Select
End Function

Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln退费 As Boolean = True) As Boolean
'功能：将门诊退费的明细和结算数据转发送医保前置服务器确认；
'参数：lngStlID-将要退的费记录的结帐ID；，从预交记录中可以检索医保号和密码
'      bln退费 -表明是退费交易还是改费交易在调用本接口
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们保证了个人帐户结算金额不大于个人帐户余额，因此交易必然成功。但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'------------------------------------------------------------------------------------------------------------------
'调用模块：1121-门诊退费
    Dim strSelfNo As String, strSelfPwd As String, str卡号 As String
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim lng病人ID As Long
    
    If gintInsure <> TYPE_沈阳市 And bln退费 = False Then
        MsgBox "该医保不支持改费功能！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Modified by ZYB 2002-10-28
    '松藻医保不需要对每个项目设置对应的医保项目
    If Not GetCapability(support允许不设置医保项目) Then
        gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
            " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=" & gintInsure & ") A," & _
            " (Select 收费细目ID,'Y' as TEST From 病人费用记录 Where 结帐ID=" & lngStlID & " And Nvl(附加标志,0)<>9) B,收费细目 C" & _
            " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
        Call OpenRecordset(rsTmp, "医保接口")
        
        If Not rsTmp.EOF Then
            MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
            ClinicDelSwap = False: Exit Function
        End If
    End If
    
    'Modified by 朱玉宝
'    '读出个人帐户使用金额
'    If GetCapability(support门诊必须传递明细) = False Then
'        gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
'            " Where A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lngStlID
'    Else
        '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
        gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
            " Where A.结帐ID=" & lngStlID
    
'    End If
    
    Call OpenRecordset(rsTmp, "门诊退费")
    If rsTmp.RecordCount = 0 And gintInsure <> TYPE_余姚 Then
        MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
        ClinicDelSwap = True
        Exit Function
    End If
    
    lng病人ID = rsTmp("病人ID")
    
    Do Until rsTmp.EOF
        If rsTmp("结算方式") = "个人帐户" Then
            curMoney = curMoney + rsTmp("冲预交")
        End If
        rsTmp.MoveNext
    Loop
    
    Select Case gintInsure
        Case TYPE_宁海
            ClinicDelSwap = 门诊结算冲销_宁海(lngStlID, curMoney, lng病人ID)
        Case TYPE_余姚
            ClinicDelSwap = 门诊结算冲销_余姚(lngStlID, curMoney, lng病人ID)
        Case TYPE_浙江
            ClinicDelSwap = 门诊结算冲销_浙江(lngStlID, curMoney, lng病人ID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            ClinicDelSwap = 门诊结算冲销_吉林(lngStlID, curMoney, lng病人ID)
        Case TYPE_北京尚洋
            ClinicDelSwap = 门诊结算冲销_北京尚洋(lngStlID, curMoney, lng病人ID)
        Case TYPE_华东
            ClinicDelSwap = 门诊结算冲销_华东(lngStlID, curMoney, lng病人ID)
        Case TYPE_南京市
            ClinicDelSwap = 门诊结算冲销_南京市(lngStlID, curMoney, lng病人ID)
        Case TYPE_江苏
            ClinicDelSwap = 门诊结算冲销_江苏(lngStlID, curMoney, lng病人ID)
        Case TYPE_涪陵
            ClinicDelSwap = 门诊结算冲销_涪陵(lngStlID, curMoney, lng病人ID)
        Case TYPE_广元
            ClinicDelSwap = 门诊结算冲销_广元(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆市
            ClinicDelSwap = 门诊结算冲销_重庆(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ClinicDelSwap = 门诊结算冲销_松藻(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆壁山
            ClinicDelSwap = 门诊结算冲销_壁山(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆中梁山
            ClinicDelSwap = 门诊结算冲销_中梁山(lngStlID, curMoney, lng病人ID)
        Case TYPE_自贡市
            ClinicDelSwap = 门诊结算冲销_中软(lngStlID, curMoney)
        Case TYPE_泸州市
            ClinicDelSwap = 门诊结算冲销_泸州(lngStlID, curMoney)
        Case TYPE_铜仁
            ClinicDelSwap = 门诊结算冲销_铜仁(lngStlID, curMoney)
        Case TYPE_云南省, TYPE_昆明市
            ClinicDelSwap = 门诊结算冲销_云南(lngStlID, curMoney, lng病人ID)
        Case TYPE_云南建水
            '由于没有yh_recedefeebalance，所以不支持
            ClinicDelSwap = False
        Case TYPE_贵阳市
            ClinicDelSwap = 门诊结算冲销_贵阳(lngStlID, curMoney, lng病人ID)
        Case TYPE_成都市
            ClinicDelSwap = False
        Case TYPE_成都莲合
            ClinicDelSwap = False
        Case TYPE_成都郊县, TYPE_新都
            ClinicDelSwap = False
        Case TYPE_咸阳市
            ClinicDelSwap = False
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ClinicDelSwap = 门诊结算冲销_福建巨龙(lngStlID, lng病人ID)
        Case type_米易
            ClinicDelSwap = 门诊结算冲销_米易(lngStlID, curMoney, lng病人ID)
        Case TYPE_四川眉山
            ClinicDelSwap = 门诊结算冲销_眉山(lngStlID, curMoney, lng病人ID)
        Case TYPE_沈阳市
            ClinicDelSwap = 门诊结算冲销_沈阳市(lngStlID, curMoney, lng病人ID, bln退费)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            ClinicDelSwap = 门诊结算冲销_大连(lngStlID, curMoney, lng病人ID)
        Case TYPE_乐山
            ClinicDelSwap = 门诊结算冲销_乐山(lngStlID, curMoney, lng病人ID)
        Case TYPE_重大校园卡
            '刘兴宏200403
            ClinicDelSwap = 门诊结算冲销_重大校园卡(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆银海版
            ClinicDelSwap = 门诊结算冲销_重庆银海版(lngStlID, curMoney, lng病人ID)
        Case TYPE_重庆渝北
            '刘兴宏20040705
            ClinicDelSwap = 门诊结算冲销_重庆渝北(lngStlID, curMoney, lng病人ID)
        Case TYPE_北京
            ClinicDelSwap = 门诊结算冲销_北京(lngStlID, curMoney, lng病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            ClinicDelSwap = 门诊结算冲销_兴安(lngStlID, curMoney, lng病人ID)
        Case TYPE_毕节
            ClinicDelSwap = 门诊结算冲销_毕节(lngStlID, curMoney, lng病人ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            ClinicDelSwap = 门诊结算冲销_黔南(lngStlID, curMoney, lng病人ID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ClinicDelSwap = 门诊结算冲销_成都德阳(lngStlID, curMoney, lng病人ID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            ClinicDelSwap = 门诊结算冲销_成都内江(lngStlID, curMoney, lng病人ID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ClinicDelSwap = 门诊结算冲销_广元旺苍(lngStlID, curMoney, lng病人ID)
        Case TYPE_渝北农医
            ClinicDelSwap = 门诊结算冲销_渝北农医(lngStlID, curMoney, lng病人ID)
        Case Is > 900
            ClinicDelSwap = 门诊结算冲销_中联(lngStlID, curMoney, lng病人ID)
    End Select
End Function

Public Function TransferSwap(lng预交ID As Long, curMoney As Currency) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lng预交ID-当前预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交收款,1131-办理登记
    Dim str医保号 As String
    Dim rsTmp As New ADODB.Recordset

    On Error Resume Next
    With rsTmp
        gstrSQL = "Select B.病人ID,A.主页ID,A.金额,B.卡号,B.密码,B.医保号,B.顺序号 " & _
            " From 病人预交记录 A,保险帐户 B " & _
            " Where A.病人ID=B.病人ID And A.结算方式='个人帐户'" & _
            " And A.ID=" & lng预交ID & " And B.险类=" & gintInsure
        Call OpenRecordset(rsTmp, "医保接口")
        
        If .EOF Then
            MsgBox "没有使用个人帐户或非参保病人，不能交易！", vbExclamation, gstrSysName
            TransferSwap = True: Exit Function
        End If
    End With
    
    Select Case gintInsure
        Case TYPE_涪陵, TYPE_华东, TYPE_北京尚洋, TYPE_余姚
            TransferSwap = False
        Case TYPE_广元
            TransferSwap = False
        Case TYPE_自贡市
            TransferSwap = 个人帐户转预交_中软(lng预交ID, curMoney)
        Case TYPE_泸州市
            TransferSwap = 个人帐户转预交_泸州(lng预交ID, curMoney)
        Case TYPE_铜仁
            TransferSwap = 个人帐户转预交_铜仁(lng预交ID, curMoney)
        Case TYPE_云南省, TYPE_昆明市
            '顺序号采用入院验证时返回的
'            If IsNull(rsTmp!顺序号) Then
'                MsgBox "未能发现该病人的住院交易顺序号，不能交易！", vbExclamation, gstrSysName
'                Exit Function
'            End If
'            str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
'            TransferSwap = 个人帐户转预交_云南(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
            MsgBox "不支持个人帐户转预交！", vbInformation, gstrSysName
            TransferSwap = False
        Case TYPE_云南建水
            '顺序号采用入院验证时返回的
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未能发现该病人的住院交易顺序号，不能交易！", vbExclamation, gstrSysName
                Exit Function
            End If
            str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
            TransferSwap = 个人帐户转预交_云南建水(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
        Case TYPE_重庆市
            TransferSwap = False
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28 重庆松藻不支持个人帐户转预交
            TransferSwap = False
        Case TYPE_重庆壁山
            TransferSwap = False
        Case TYPE_重庆中梁山
            TransferSwap = False
        Case TYPE_贵阳市
            TransferSwap = False
        Case TYPE_成都市
            TransferSwap = 个人帐户转预交_成都(lng预交ID, curMoney, rsTmp)
        Case TYPE_成都莲合
            TransferSwap = 个人帐户转预交_莲合(lng预交ID, curMoney)
        Case TYPE_成都郊县
            str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
            TransferSwap = 个人帐户转预交_成都郊县(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
        Case TYPE_新都
            str医保号 = TrimStr(IIf(IsNull(rsTmp("医保号")), "", rsTmp("医保号")))
            TransferSwap = 个人帐户转预交_新都(lng预交ID, curMoney, str医保号, rsTmp("顺序号"), rsTmp("病人ID"))
        Case TYPE_咸阳市
            TransferSwap = False
        Case type_米易
            TransferSwap = False
        Case TYPE_四川眉山
            TransferSwap = False
        Case type_大连市, type_大连开发区
              '刘兴宏(200311),无预交
            TransferSwap = False
        Case TYPE_乐山
            TransferSwap = False
        Case TYPE_重大校园卡
              '刘兴宏(200403)
             TransferSwap = 个人帐户转预交_重大校园卡(lng预交ID, curMoney, rsTmp)
        Case TYPE_重庆渝北
              '刘兴宏(20040705)
            TransferSwap = False
        Case TYPE_兴安
            '刘兴宏:20040827
            TransferSwap = False
        Case TYPE_黔南
            '刘兴宏:20041022
            TransferSwap = False
        Case TYPE_成都德阳
            '刘兴宏:20041103
            TransferSwap = False
        Case TYPE_成都内江
            '刘兴宏:20041207
            TransferSwap = False
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            TransferSwap = False
        Case TYPE_渝北农医
            TransferSwap = False
        Case Is > 900
            TransferSwap = 个人帐户转预交_中联(lng预交ID, curMoney, rsTmp!病人ID)
    End Select
End Function

Public Function TransferDelSwap(lng预交ID As Long) As Boolean
'功能：将需要从个人帐户余额转入预交款的数据记录发送医保前置服务器确认；
'参数：lngPrepID-将要退的预交记录的ID，从预交记录中可以检索医保号和密码
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1103-预交退款
    
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency, lng病人ID As Long
    
    On Error Resume Next
    With rsTmp
        gstrSQL = "Select B.病人ID,A.金额,B.卡号,B.密码,B.医保号,B.顺序号 " & _
            " From 病人预交记录 A,保险帐户 B " & _
            " Where A.病人ID=B.病人ID And A.结算方式='个人帐户'" & _
            " And A.ID=" & lng预交ID & " And B.险类=" & gintInsure
        Call OpenRecordset(rsTmp, "医保接口")

        If .EOF Then
            MsgBox "没有使用个人帐户或非参保病人，不能交易！", vbExclamation, gstrSysName
            TransferDelSwap = True: Exit Function
        End If
        
        lng病人ID = rsTmp("病人ID")
        Do Until rsTmp.EOF
            curMoney = curMoney + rsTmp("金额")
            rsTmp.MoveNext
        Loop
    End With
    
    Select Case gintInsure
        Case TYPE_江苏, TYPE_华东, TYPE_北京尚洋, TYPE_余姚
            TransferDelSwap = False
        Case TYPE_涪陵, TYPE_广元
            TransferDelSwap = False
        Case TYPE_自贡市
            TransferDelSwap = False
        Case TYPE_泸州市, TYPE_铜仁
            TransferDelSwap = False
        Case TYPE_云南省, TYPE_昆明市, TYPE_云南建水
            TransferDelSwap = False
        Case TYPE_重庆市
            TransferDelSwap = False
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            TransferDelSwap = False
        Case TYPE_重庆壁山
            TransferDelSwap = False
        Case TYPE_重庆中梁山
            TransferDelSwap = False
        Case TYPE_贵阳市
            TransferDelSwap = False
        Case TYPE_成都市
            TransferDelSwap = False
        Case TYPE_成都莲合
            TransferDelSwap = False
        Case TYPE_成都郊县, TYPE_新都
            TransferDelSwap = False
        Case TYPE_咸阳市
            TransferDelSwap = False
        Case type_米易
            TransferDelSwap = False
        Case TYPE_四川眉山
            TransferDelSwap = False
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)，无退款
            TransferDelSwap = False
        Case TYPE_乐山
            TransferDelSwap = False
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            TransferDelSwap = 个人帐户转预交冲销_重大校园卡(lng预交ID, curMoney, lng病人ID)
        Case TYPE_重庆渝北
              '刘兴宏(20040705)
            TransferDelSwap = False
        Case TYPE_兴安
            '刘兴宏:20040827
            TransferDelSwap = False
        Case TYPE_黔南
            '刘兴宏:20041022
            TransferDelSwap = False
        Case TYPE_成都德阳
            '刘兴宏:20041103
            TransferDelSwap = False
        Case TYPE_成都内江
            '刘兴宏:20041207
            TransferDelSwap = False
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            TransferDelSwap = False
        Case TYPE_渝北农医
            TransferDelSwap = False
        Case Is > 900
            TransferDelSwap = 个人帐户转预交冲销_中联(lng预交ID, curMoney, lng病人ID)
    End Select
End Function

Public Function WipeoffMoney(rsExse As Recordset, lng病人ID As Long, str医保号 As String, str密码 As String) As String
    '功能：获取该病人指定结帐内容的可报销金额；
    '参数：rsExse-需要结算的费用明细记录集合；str医保号-医保号；str密码-病人密码；
    '返回：可报销金额串:"报销方式;金额;是否允许修改|...."
    '注意：1)该函数主要使用模拟结算交易，查询结果返回获取基金报销额；
    'str密码：由于各医保一直未使用，现改为：1-结帐处调用;0-其他地方调用（病人费用查询）
    '------------------------------------------------------------------------------------------------------------------
    '调用模块：1137-结帐办理
    Dim strSort As String
    If rsExse.State = adStateClosed Then Exit Function
    If rsExse.RecordCount = 0 Then
        MsgBox "没有费用明细，或者收费项目没有设置对应的医保编码。", vbInformation, gstrSysName
        Exit Function
    End If
    
    Select Case gintInsure
        Case TYPE_宁海
            WipeoffMoney = 住院虚拟结算_宁海(rsExse, lng病人ID)
        Case TYPE_余姚
            WipeoffMoney = 住院虚拟结算_余姚(rsExse, lng病人ID, str医保号)
        Case TYPE_浙江
            WipeoffMoney = 住院虚拟结算_浙江(rsExse, lng病人ID, str医保号)
        Case TYPE_吉林
            '20040923刘兴宏加入
            WipeoffMoney = 住院虚拟结算_吉林(rsExse, lng病人ID, IIf(str密码 = "1", True, False))
        Case TYPE_北京尚洋
            WipeoffMoney = 住院虚拟结算_北京尚洋(rsExse, lng病人ID, str医保号)
        Case TYPE_华东
            WipeoffMoney = 住院虚拟结算_华东(rsExse, lng病人ID, str医保号)
        Case TYPE_南京市
            WipeoffMoney = 住院虚拟结算_南京市(rsExse, lng病人ID)
        Case TYPE_江苏
            WipeoffMoney = 住院虚拟结算_江苏(rsExse, lng病人ID, str医保号)
        Case TYPE_涪陵
            WipeoffMoney = 住院虚拟结算_涪陵(rsExse, lng病人ID, str医保号)
        Case TYPE_广元
            WipeoffMoney = 住院虚拟结算_广元(rsExse, lng病人ID, str医保号)
        Case TYPE_重庆市
            WipeoffMoney = 住院虚拟结算_重庆(rsExse, lng病人ID, str医保号, IIf(str密码 = "1", False, True))
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            WipeoffMoney = 住院虚拟结算_松藻(rsExse)
        Case TYPE_重庆壁山
            WipeoffMoney = 住院虚拟结算_壁山(rsExse, lng病人ID, str医保号, str密码)
        Case TYPE_重庆中梁山
            WipeoffMoney = 住院虚拟结算_中梁山(rsExse, lng病人ID)
        Case TYPE_贵阳市
            WipeoffMoney = 住院虚拟结算_贵阳(rsExse, lng病人ID)
        Case TYPE_自贡市
            WipeoffMoney = 住院虚拟结算_中软(rsExse)
        Case TYPE_泸州市
            WipeoffMoney = 住院虚拟结算_泸州(rsExse)
        Case TYPE_铜仁
            WipeoffMoney = 住院虚拟结算_铜仁(rsExse)
        Case TYPE_云南省, TYPE_昆明市
            WipeoffMoney = 住院虚拟结算_云南(rsExse, lng病人ID)
        Case TYPE_云南建水
            WipeoffMoney = 住院虚拟结算_云南建水(rsExse, lng病人ID)
        Case TYPE_成都市
            WipeoffMoney = 住院虚拟结算_成都(rsExse, str医保号, str密码)
        Case TYPE_成都莲合
            WipeoffMoney = 住院虚拟结算_莲合(rsExse, str医保号, str密码)
        Case TYPE_成都郊县
            WipeoffMoney = 住院虚拟结算_成都郊县(rsExse, lng病人ID, str医保号)
        Case TYPE_新都
            WipeoffMoney = 住院虚拟结算_新都(rsExse, lng病人ID, str医保号)
        Case TYPE_成都南充
            WipeoffMoney = 住院虚拟结算_成都南充(lng病人ID)
        Case TYPE_咸阳市
            WipeoffMoney = 住院虚拟结算_咸阳(rsExse, lng病人ID, str医保号)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            WipeoffMoney = 住院虚拟结算_福建巨龙(lng病人ID)
        Case type_米易
            WipeoffMoney = 住院虚拟结算_米易(rsExse, lng病人ID)
        Case TYPE_四川眉山
            WipeoffMoney = 住院虚拟结算_眉山(rsExse, lng病人ID)
        Case TYPE_沈阳市
            WipeoffMoney = 住院虚拟结算_沈阳市(rsExse, lng病人ID)
        Case type_大连市, type_大连开发区
            '刘兴宏(200311)
            WipeoffMoney = 住院虚拟结算_大连(rsExse, lng病人ID)
        Case TYPE_乐山
            WipeoffMoney = 住院虚拟结算_乐山(rsExse, lng病人ID)
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            WipeoffMoney = 住院虚拟结算_重大校园卡(rsExse, lng病人ID)
        Case TYPE_重庆银海版
            WipeoffMoney = 住院虚拟结算_重庆银海版(rsExse, lng病人ID)
        Case TYPE_重庆渝北
            '刘兴宏(20040705)
            WipeoffMoney = 住院虚拟结算_重庆渝北(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_北京
            WipeoffMoney = 住院虚拟结算_北京(rsExse, lng病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            WipeoffMoney = 住院虚拟结算_兴安(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_毕节
            WipeoffMoney = 住院虚拟结算_毕节(rsExse, lng病人ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            WipeoffMoney = 住院虚拟结算_黔南(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_成都德阳
            '刘兴宏:20041103
            WipeoffMoney = 住院虚拟结算_成都德阳(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_成都内江
            '刘兴宏:20041207
            WipeoffMoney = 住院虚拟结算_成都内江(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            WipeoffMoney = 住院虚拟结算_广元旺苍(rsExse, lng病人ID, IIf(Val(str密码) = 0, False, True))
        Case TYPE_渝北农医
            WipeoffMoney = 住院虚拟结算_渝北农医(rsExse, lng病人ID)
        Case Is > 900
            WipeoffMoney = 住院虚拟结算_中联(rsExse)
    End Select
    
    '20040107:周韬:有些接口未填写主页ID
    rsExse.Filter = 0
    strSort = rsExse.Sort
    rsExse.Sort = "主页ID"
    rsExse.MoveLast '取最大的
    g结算数据.主页ID = Nvl(rsExse!主页ID, 0)
    rsExse.Sort = strSort
    DebugTool ("当前主页ID：" & rsExse!主页ID)
    Call Insert虚拟结算数据(lng病人ID, g结算数据.主页ID, WipeoffMoney)
End Function

Public Function SettleSwap(lng结帐ID As Long) As Boolean
'功能：将需要本次结帐的费用明细和结帐数据发送医保前置服务器确认；
'参数: lng结帐ID-病人结帐记录ID, 从预交记录中可以检索出各种结算方式及金额
'返回：交易成功返回true；否则，返回false
'注意：1)主要利用接口的费用明细传输交易和辅助结算交易；
'      2)理论上，由于我们通过模拟结算提取了基金报销额，保证了医保基金结算金额的正确性，因此交易必然成功。
'        但从安全角度考虑；当辅助结算交易失败时，需要使用费用删除交易处理；如果辅助结算交易成功，但费用
'        分割结果与我们处理结果不一致，需要执行恢复结算交易和费用删除交易。这样才能保证数据的完全统一。
'      3)由于结帐之后，可能使用结帐作废交易，这时需要结帐时执行结算交易的交易号，因此我们需要同时结帐交易号。
'        (由于门诊收费作废时，已经不再和医保有关系，所以不需要保存结帐的交易号)
'------------------------------------------------------------------------------------------------------------------
'调用模块：1137-结帐办理
    
    Dim rsTmp As New ADODB.Recordset, lng病人ID As Long
    
    With rsTmp
        .CursorLocation = adUseClient
        
        'Modified by ZYB 2002-10-28
        '松藻医保不需要对每个项目设置对应的医保项目
        'Modified by ZYB 2003-12-14
        '其它只判断未上传的费用明细
        If Not GetCapability(support允许不设置医保项目) Then
            gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=" & gintInsure & ") A," & _
                " (Select 收费细目ID,'Y' as TEST From 病人费用记录 Where Nvl(是否上传,0)=0 And Nvl(附加标志,0)<>9 And 结帐ID=" & lng结帐ID & ") B,收费细目 C" & _
                " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
            Call OpenRecordset(rsTmp, "医保接口")
            
            If Not .EOF Then
                MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行交易！", vbExclamation, gstrSysName
                Exit Function
            End If
        End If
        
        gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码,B.顺序号" & _
            " From 病人结帐记录 A,保险帐户 B " & _
            " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
            " And A.ID=" & lng结帐ID
        Call OpenRecordset(rsTmp, "医保接口")
        
        If .EOF Then
            MsgBox "没有发现该病人的结算信息,不能执行交易！", vbExclamation, gstrSysName
            Exit Function
        End If
        lng病人ID = rsTmp("病人ID")
    End With
    
    Select Case gintInsure
        Case TYPE_宁海
            SettleSwap = 住院结算_宁海(lng结帐ID, lng病人ID)
        Case TYPE_余姚
            SettleSwap = 住院结算_余姚(lng结帐ID)
        Case TYPE_浙江
            SettleSwap = 住院结算_浙江(lng结帐ID, lng病人ID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            SettleSwap = 住院结算_吉林(lng结帐ID)
        Case TYPE_北京尚洋
            SettleSwap = 住院结算_北京尚洋(lng结帐ID, lng病人ID)
        Case TYPE_华东
            SettleSwap = 住院结算_华东(lng结帐ID, lng病人ID)
        Case TYPE_南京市
            SettleSwap = 住院结算_南京市(lng结帐ID, lng病人ID)
        Case TYPE_江苏
            SettleSwap = 住院结算_江苏(lng结帐ID)
        Case TYPE_涪陵
            SettleSwap = 住院结算_涪陵(lng结帐ID)
        Case TYPE_广元
            SettleSwap = 住院结算_广元(lng结帐ID)
        Case TYPE_重庆市
            SettleSwap = 住院结算_重庆(lng结帐ID, lng病人ID)
        Case TYPE_重庆松藻
            SettleSwap = 住院结算_松藻(lng结帐ID)
        Case TYPE_重庆壁山
            SettleSwap = 住院结算_壁山(lng结帐ID, lng病人ID)
        Case TYPE_重庆中梁山
            SettleSwap = 住院结算_中梁山(lng结帐ID)
        Case TYPE_自贡市
            SettleSwap = 住院结算_中软(lng结帐ID)
        Case TYPE_泸州市
            SettleSwap = 住院结算_泸州(lng结帐ID)
        Case TYPE_铜仁
            SettleSwap = 住院结算_铜仁(lng结帐ID)
        Case TYPE_贵阳市
            SettleSwap = 住院结算_贵阳(lng结帐ID, lng病人ID)
        Case TYPE_云南省, TYPE_昆明市
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbExclamation, gstrSysName
                Exit Function
            End If
            SettleSwap = 住院结算_云南(lng结帐ID, rsTmp!顺序号, lng病人ID)
        Case TYPE_云南建水
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbExclamation, gstrSysName
                Exit Function
            End If
            SettleSwap = 住院结算_云南建水(lng结帐ID, rsTmp!顺序号, lng病人ID)
        Case TYPE_成都市
            SettleSwap = 住院结算_成都(lng结帐ID, rsTmp)
        Case TYPE_成都莲合
            '因为在试结算时，已经保存了，如果不成功，到不到此！所以在此不作保存
            SettleSwap = 住院结算_莲合(lng结帐ID, rsTmp)
        Case TYPE_成都南充
            SettleSwap = 住院结算_成都南充(lng结帐ID, rsTmp)
        Case TYPE_成都郊县
            SettleSwap = 住院结算_成都郊县(lng结帐ID, lng病人ID)
        Case TYPE_新都
            SettleSwap = 住院结算_新都(lng结帐ID, lng病人ID)
        Case TYPE_咸阳市
            SettleSwap = 住院结算_咸阳(lng结帐ID, lng病人ID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            SettleSwap = 住院结算_福建巨龙(lng病人ID, lng结帐ID)
        Case type_米易
            SettleSwap = 住院结算_米易(lng结帐ID, lng病人ID)
        Case TYPE_四川眉山
            SettleSwap = 住院结算_眉山(lng结帐ID, lng病人ID)
        Case TYPE_沈阳市
            SettleSwap = 住院结算_沈阳(lng结帐ID, lng病人ID)
        Case type_大连市, type_大连开发区
            '刘兴宏(200311)
            SettleSwap = 住院结算_大连(lng结帐ID, lng病人ID)
        Case TYPE_乐山
            SettleSwap = 住院结算_乐山(lng结帐ID, lng病人ID)
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            SettleSwap = 住院结算_重大校园卡(lng结帐ID, lng病人ID)
        Case TYPE_重庆银海版
            SettleSwap = 住院结算_重庆银海版(lng结帐ID, lng病人ID)
        Case TYPE_重庆渝北
            '刘兴宏(20040705)
            SettleSwap = 住院结算_重庆渝北(lng结帐ID, lng病人ID)
        Case TYPE_北京
            SettleSwap = 住院结算_北京(lng结帐ID, lng病人ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            SettleSwap = 住院结算_兴安(lng结帐ID, lng病人ID)
        Case TYPE_毕节
            SettleSwap = 住院结算_毕节(lng结帐ID, lng病人ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            SettleSwap = 住院结算_黔南(lng结帐ID, lng病人ID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            SettleSwap = 住院结算_成都德阳(lng结帐ID, lng病人ID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            SettleSwap = 住院结算_成都内江(lng结帐ID, lng病人ID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            SettleSwap = 住院结算_广元旺苍(lng结帐ID, lng病人ID)
        Case TYPE_渝北农医
            SettleSwap = 住院结算_渝北农医(lng结帐ID, lng病人ID)
        Case Is > 900
            '中联医保
            SettleSwap = 住院结算_中联(lng结帐ID)
    End Select
    
    Call Clear虚拟结算数据(lng病人ID, g结算数据.主页ID)
End Function

Public Function SettleDelSwap(lng结帐ID As Long) As Boolean
'功能：将指定结帐涉及的结帐交易和费用明细从医保数据中删除；
'参数：lng结帐ID=需要作废的结帐单ID号（注意是没作废的那条记录）；
'返回：交易成功返回true；否则，返回false
'注意：1)主要使用结帐恢复交易和费用删除交易；
'      2)有关原结算交易号，在病人结帐记录中根据结帐单ID查找；原费用明细传输交易的交易号，
'        在病人费用记录中根据结帐ID查找；
'      3)作废的结帐记录(记录性质=2)其交易号，填写本次结帐恢复交易的交易号；
'        因结帐作废而产生的费用记录的交易号号，填写为本次费用删除交易的交易号。
'------------------------------------------------------------------------------------------------------------------
'调用模块：1137-结帐作废

    Dim rsTmp As New ADODB.Recordset
    
    'On Error Resume Next
    
    With rsTmp
        .CursorLocation = adUseClient
        gstrSQL = "Select A.病人ID,B.医保号,B.卡号,B.密码,B.顺序号" & _
            " From 保险结算记录 A,保险帐户 B " & _
            " Where A.性质=2 And A.记录ID=" & lng结帐ID & " And A.险类=" & gintInsure & _
            " And A.险类=B.险类 And A.病人ID=B.病人ID"
        Call OpenRecordset(rsTmp, "医保接口")
        
        If .RecordCount = 0 Then
            MsgBox "没有使用医保结算，可以不向医保交易", vbExclamation, gstrSysName
            SettleDelSwap = False: Exit Function
        End If
    End With
    
    Select Case gintInsure
        Case TYPE_宁海
            SettleDelSwap = 住院结算冲销_宁海(lng结帐ID)
        Case TYPE_余姚
            SettleDelSwap = 住院结算冲销_余姚(lng结帐ID)
        Case TYPE_浙江
            SettleDelSwap = 住院结算冲销_浙江(lng结帐ID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            SettleDelSwap = 住院结算冲销_吉林(lng结帐ID)
        Case TYPE_北京尚洋
            SettleDelSwap = 住院结算冲销_北京尚洋(lng结帐ID)
        Case TYPE_华东
            SettleDelSwap = 住院结算冲销_华东(lng结帐ID)
        Case TYPE_南京市
            SettleDelSwap = 住院结算冲销_南京市(lng结帐ID)
        Case TYPE_江苏
            SettleDelSwap = 住院结算冲销_江苏(lng结帐ID)
        Case TYPE_涪陵
            SettleDelSwap = 住院结算冲销_涪陵(lng结帐ID)
        Case TYPE_广元
            SettleDelSwap = 住院结算冲销_广元(lng结帐ID)
        Case TYPE_重庆市
            SettleDelSwap = 住院结算冲销_重庆(lng结帐ID)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            SettleDelSwap = 住院结算冲销_松藻(lng结帐ID)
        Case TYPE_重庆壁山
            SettleDelSwap = 住院结算冲销_壁山(lng结帐ID)
        Case TYPE_重庆中梁山
            SettleDelSwap = 住院结算冲销_中梁山(lng结帐ID)
        Case TYPE_自贡市
            SettleDelSwap = 住院结算冲销_中软(lng结帐ID)
        Case TYPE_泸州市
            SettleDelSwap = 住院结算冲销_泸州(lng结帐ID)
        Case TYPE_铜仁
            SettleDelSwap = 住院结算冲销_铜仁(lng结帐ID)
        Case TYPE_云南省, TYPE_昆明市
            SettleDelSwap = 住院结算冲销_云南(lng结帐ID)
        Case TYPE_云南建水
            '暂时没处理取消结帐
            SettleDelSwap = False
        Case TYPE_贵阳市
            SettleDelSwap = 住院结算冲销_贵阳(lng结帐ID)
        Case TYPE_成都市
            SettleDelSwap = 住院结算冲销_成都(lng结帐ID, rsTmp)
        Case TYPE_成都莲合
            SettleDelSwap = 住院结算冲销_莲合(lng结帐ID, rsTmp)
        Case TYPE_成都郊县, TYPE_新都
            SettleDelSwap = False
        Case TYPE_咸阳市
            SettleDelSwap = False
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            SettleDelSwap = 住院结算冲销_福建巨龙(lng结帐ID)
        Case type_米易
            SettleDelSwap = 住院结算冲销_米易(lng结帐ID)
        Case TYPE_四川眉山
            SettleDelSwap = 住院结算冲销_眉山(lng结帐ID)
        Case TYPE_沈阳市
            SettleDelSwap = 住院结算冲销_沈阳市(lng结帐ID)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            SettleDelSwap = 住院结算冲销_大连(lng结帐ID)
        Case TYPE_乐山
            SettleDelSwap = 住院结算冲销_乐山(lng结帐ID)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            SettleDelSwap = 住院结算冲销_重大校园卡(lng结帐ID)
        Case TYPE_重庆银海版
            SettleDelSwap = 住院结算冲销_重庆银海版(lng结帐ID)
        Case TYPE_重庆渝北
              '刘兴宏(20040705)
            SettleDelSwap = 住院结算冲销_重庆渝北(lng结帐ID)
        Case TYPE_北京
            SettleDelSwap = 住院结算冲销_北京(lng结帐ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            SettleDelSwap = 住院结算冲销_兴安(lng结帐ID)
        Case TYPE_毕节
            SettleDelSwap = 住院结算冲销_毕节(lng结帐ID)
        Case TYPE_黔南
            '刘兴宏:20041022
            SettleDelSwap = 住院结算冲销_黔南(lng结帐ID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            SettleDelSwap = 住院结算冲销_成都德阳(lng结帐ID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            SettleDelSwap = 住院结算冲销_成都内江(lng结帐ID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            SettleDelSwap = 住院结算冲销_广元旺苍(lng结帐ID)
        Case TYPE_渝北农医
            SettleDelSwap = 住院结算冲销_渝北农医(lng结帐ID)
        Case Is > 900
            '中联医保
            SettleDelSwap = 住院结算冲销_中联(lng结帐ID)
    End Select
End Function

Public Function ComeInSwap(lngPatiID As Long, lngPageID As Long, ByRef str医保号 As String) As Boolean
'功能：将入院登记信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-办理登记,1132-撤消出院
    
    Select Case gintInsure
        Case TYPE_宁海
            ComeInSwap = 入院登记_宁海(lngPatiID, lngPageID, str医保号)
        Case TYPE_余姚
            ComeInSwap = 入院登记_余姚(lngPatiID, lngPageID, str医保号)
        Case TYPE_浙江
            ComeInSwap = 入院登记_浙江(lngPatiID, lngPageID, str医保号)
        Case TYPE_吉林
            '20040923刘兴宏加入
            ComeInSwap = 入院登记_吉林(lngPatiID, lngPageID, str医保号)
        Case TYPE_北京尚洋
            ComeInSwap = 入院登记_北京尚洋(lngPatiID, lngPageID, str医保号)
        Case TYPE_华东
            ComeInSwap = 入院登记_华东(lngPatiID, lngPageID, str医保号)
        Case TYPE_南京市
            ComeInSwap = True
        Case TYPE_江苏
            ComeInSwap = 入院登记_江苏(lngPatiID, lngPageID, str医保号)
        Case TYPE_涪陵
            ComeInSwap = 入院登记_涪陵(lngPatiID, lngPageID, str医保号)
        Case TYPE_广元
            ComeInSwap = 入院登记_广元(lngPatiID, lngPageID, str医保号)
        Case TYPE_重庆市
            ComeInSwap = 入院登记_重庆(lngPatiID, lngPageID, str医保号)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ComeInSwap = 入院登记_松藻(lngPatiID)
        Case TYPE_重庆壁山
            ComeInSwap = 入院登记_壁山(lngPatiID, lngPageID, str医保号)
        Case TYPE_重庆中梁山
            ComeInSwap = 入院登记_中梁山(lngPatiID, lngPageID)
        Case TYPE_自贡市
            ComeInSwap = 入院登记_中软(lngPatiID, lngPageID, str医保号)
        Case TYPE_泸州市
            ComeInSwap = 入院登记_泸州(lngPatiID, lngPageID, str医保号)
        Case TYPE_铜仁
            ComeInSwap = 入院登记_铜仁(lngPatiID, lngPageID, str医保号)
        Case TYPE_贵阳市
            ComeInSwap = 入院登记_贵阳(lngPatiID, lngPageID, str医保号)
        Case TYPE_云南省, TYPE_昆明市
            ComeInSwap = 入院登记_云南(lngPatiID, lngPageID, str医保号)
        Case TYPE_云南建水
            ComeInSwap = 入院登记_云南建水(lngPatiID, lngPageID, str医保号)
        Case TYPE_成都市
            ComeInSwap = 入院登记_成都(lngPatiID, lngPageID, str医保号)
        Case TYPE_成都莲合
            ComeInSwap = 入院登记_莲合(lngPatiID, lngPageID)
        Case TYPE_成都郊县
            ComeInSwap = 入院登记_成都郊县(lngPatiID, lngPageID, str医保号)
        Case TYPE_新都
            ComeInSwap = 入院登记_新都(lngPatiID, lngPageID, str医保号)
        Case TYPE_成都南充
            ComeInSwap = 入院登记_成都南充(lngPatiID)
        Case TYPE_咸阳市
            ComeInSwap = 入院登记_咸阳(lngPatiID, lngPageID, str医保号)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ComeInSwap = 入院登记_福建巨龙(lngPatiID)
        Case type_米易
            ComeInSwap = 入院登记_米易(lngPatiID, lngPageID, str医保号)
        Case TYPE_四川眉山
            ComeInSwap = 入院登记_眉山(lngPatiID, lngPageID, str医保号)
        Case TYPE_沈阳市
            ComeInSwap = 入院登记_沈阳市(lngPatiID, lngPageID, str医保号)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            ComeInSwap = 入院登记_大连(lngPatiID, lngPageID, str医保号)
        Case TYPE_乐山
            ComeInSwap = 入院登记_乐山(lngPatiID, lngPageID, str医保号)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            ComeInSwap = 入院登记_大连(lngPatiID, lngPageID, str医保号)
        Case TYPE_重庆银海版
            ComeInSwap = 入院登记_重庆银海版(lngPatiID, lngPageID, str医保号)
        Case TYPE_重庆渝北
              '刘兴宏(20040705)
            ComeInSwap = 入院登记_重庆渝北(lngPatiID, lngPageID, str医保号)
        Case TYPE_北京
            ComeInSwap = 入院登记_北京(lngPatiID, lngPageID)
        Case TYPE_兴安
            '刘兴宏:20040827
            ComeInSwap = 入院登记_兴安(lngPatiID, lngPageID, str医保号)
        Case TYPE_毕节
            ComeInSwap = 入院登记_毕节(lngPatiID, lngPageID)
        Case TYPE_黔南
            '刘兴宏:20041022
            ComeInSwap = 入院登记_黔南(lngPatiID, lngPageID, str医保号)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ComeInSwap = 入院登记_成都德阳(lngPatiID, lngPageID, str医保号)
        Case TYPE_成都内江
            '刘兴宏:20041207
            ComeInSwap = 入院登记_成都内江(lngPatiID, lngPageID, str医保号)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ComeInSwap = 入院登记_广元旺苍(lngPatiID, lngPageID, str医保号)
        Case TYPE_渝北农医
            ComeInSwap = 入院登记_渝北农医(lngPatiID, lngPageID, str医保号)
        Case Is > 900
            '中联医保
            ComeInSwap = 入院登记_中联(lngPatiID)
    End Select
End Function


Public Function LeaveSwap(lngPatiID As Long, lngPageID As Long) As Boolean
'功能：将出院信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-取消入院,1132-病人出院
    Dim rsTmp As New ADODB.Recordset
    
    'On Error Resume Next
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=" & lngPatiID & " And A.主页ID=" & lngPageID & _
            " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=" & gintInsure
    Call OpenRecordset(rsTmp, "医保接口")
    
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Select Case gintInsure
        Case TYPE_宁海
            LeaveSwap = 出院登记_宁海(lngPatiID, lngPageID)
        Case TYPE_余姚
            LeaveSwap = 出院登记_余姚(lngPatiID, lngPageID)
        Case TYPE_浙江
            LeaveSwap = 出院登记_浙江(lngPatiID, lngPageID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            LeaveSwap = 出院登记_吉林(lngPatiID, lngPageID)
        Case TYPE_北京尚洋
            LeaveSwap = 出院登记_北京尚洋(lngPatiID, lngPageID)
        Case TYPE_华东
            LeaveSwap = 出院登记_华东(lngPatiID, lngPageID)
        Case TYPE_南京市
            LeaveSwap = True
        Case TYPE_江苏
            LeaveSwap = 出院登记_江苏(lngPatiID, lngPageID)
        Case TYPE_涪陵
            LeaveSwap = 出院登记_涪陵(lngPatiID, lngPageID)
        Case TYPE_广元
            LeaveSwap = 出院登记_广元(lngPatiID, lngPageID)
        Case TYPE_重庆市
            LeaveSwap = 出院登记_重庆(lngPatiID, lngPageID)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            LeaveSwap = 出院登记_松藻(lngPatiID, lngPageID)
        Case TYPE_重庆壁山
            LeaveSwap = 出院登记_壁山(lngPatiID, lngPageID)
        Case TYPE_重庆中梁山
            LeaveSwap = 出院登记_中梁山(lngPatiID, lngPageID)
        Case TYPE_自贡市
            LeaveSwap = 出院登记_中软(lngPatiID)
        Case TYPE_泸州市
            LeaveSwap = 出院登记_泸州(lngPatiID)
        Case TYPE_铜仁
            LeaveSwap = 出院登记_铜仁(lngPatiID)
        Case TYPE_贵阳市
            LeaveSwap = 出院登记_贵阳(lngPatiID, lngPageID)
        Case TYPE_云南省, TYPE_昆明市
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                Exit Function
            End If
            LeaveSwap = 出院登记_云南(lngPatiID, lngPageID, rsTmp!顺序号)
        Case TYPE_云南建水
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                Exit Function
            End If
            LeaveSwap = 出院登记_云南建水(lngPatiID, lngPageID, rsTmp!顺序号)
        Case TYPE_成都市
            LeaveSwap = 出院登记_成都(lngPatiID, lngPageID, rsTmp)
        Case TYPE_成都莲合
            LeaveSwap = 出院登记_莲合(lngPatiID, lngPageID)
        Case TYPE_成都郊县
            LeaveSwap = 出院登记_成都郊县(lngPatiID, lngPageID)
        Case TYPE_新都
            LeaveSwap = 出院登记_新都(lngPatiID, lngPageID)
        Case TYPE_成都南充
            LeaveSwap = 出院登记_成都南充(lngPatiID)
        Case TYPE_咸阳市
            LeaveSwap = 出院登记_咸阳(lngPatiID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            LeaveSwap = 出院登记_福建巨龙(lngPatiID, lngPageID)
        Case type_米易
            LeaveSwap = 出院登记_米易(lngPatiID, lngPageID)
        Case TYPE_四川眉山
            LeaveSwap = 出院登记_眉山(lngPatiID, lngPageID)
        Case TYPE_沈阳市
            LeaveSwap = 出院登记_沈阳市(lngPatiID, lngPageID)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            LeaveSwap = 出院登记_大连(lngPatiID, lngPageID)
        Case TYPE_乐山
            LeaveSwap = 出院登记_乐山(lngPatiID, lngPageID)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            LeaveSwap = 出院登记_重大校园卡(lngPatiID, lngPageID)
        Case TYPE_重庆银海版
            LeaveSwap = 出院登记_重庆银海版(lngPatiID, lngPageID)
        Case TYPE_重庆渝北
              '刘兴宏(20040705)
            LeaveSwap = 出院登记_重庆渝北(lngPatiID, lngPageID)
        Case TYPE_北京
            LeaveSwap = 出院登记_北京(lngPatiID, lngPageID)
        Case TYPE_兴安
            '刘兴宏:20040827
            LeaveSwap = 出院登记_兴安(lngPatiID, lngPageID)
        Case TYPE_毕节
            LeaveSwap = 出院登记_毕节(lngPatiID, lngPageID)
        Case TYPE_黔南
            '刘兴宏:20041022
            LeaveSwap = 出院登记_黔南(lngPatiID, lngPageID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            LeaveSwap = 出院登记_成都德阳(lngPatiID, lngPageID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            LeaveSwap = 出院登记_成都内江(lngPatiID, lngPageID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            LeaveSwap = 出院登记_广元旺苍(lngPatiID, lngPageID)
        Case TYPE_渝北农医
            LeaveSwap = 出院登记_渝北农医(lngPatiID, lngPageID)
        Case Is > 900
            '中联医保
            LeaveSwap = 出院登记_中联(lngPatiID, lngPageID)
    End Select
End Function

Public Function GetCapability(ByVal cap参数 As Integer, Optional ByVal lng病人ID As Long = 0, Optional ByVal int险类 As Integer) As Boolean
'功能：判断后来补充的一些业务在不同的医保软件是否得到支持
'参数：cap业务     医院的业务代号
'      int险类     指定险类,缺省为当前险类(如果已连接)
'返回：True        支持该业务
'      False       不支持该业务
'说明：不需要进行医保初始化就可以完成的查询

    Dim rsTemp As New ADODB.Recordset
    Dim cap业务 As 医院业务
    
    '由于本函数只对新增业务进行判断，缺省医保是不支持这些新业务
    GetCapability = False
        
    cap业务 = cap参数
    
    If int险类 = 0 Then int险类 = gintInsure
    
    Select Case int险类
        'support门诊预算在下面有单独的处理
        Case TYPE_渝北农医
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, support撤消出院, support出院病人结算作废, _
                    support门诊预算, support门诊退费, support未结清出院, support允许不设置医保项目, _
                    support医嘱上传, support记帐上传, support记帐作废上传, support结帐退个人帐户, support允许部份冲销单据
                GetCapability = True
            End Select
        Case TYPE_宁海
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, _
                    support撤销出院, support允许不设置医保项目, support门诊退费, _
                    support门诊预算, support允许冲销已结帐的记帐单据, support撤消出院
                GetCapability = True
            End Select
        Case TYPE_余姚
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊预算, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support撤消出院, _
                     support允许不设置医保项目
'                     support出院病人结算作废, _
'                     support记帐上传, _
'                     support记帐作废上传, _
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_浙江
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support记帐完成后上传, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support撤消出院, _
                     support允许不设置医保项目
'                     support结帐退个人帐户, _
'                     support收费帐户全自费, _
'                     support收费帐户首先自付, _
'                     support结算使用个人帐户, _
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_毕节
            Select Case cap业务
                Case support门诊必须传递明细, support结算使用个人帐户, support记帐上传, _
                    support记帐作废上传, support医嘱上传, support出院病人结算作废, _
                    support撤销出院, support允许不设置医保项目, support门诊退费, _
                    support门诊预算, support未结清出院, support允许冲销已结帐的记帐单据
                GetCapability = True
            End Select
        Case TYPE_吉林
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤消出院
            '                     support结算使用个人帐户, _
                     '刘兴宏:20040923取消此参数:support允许不设置医保项目
                     '支持的业务
                     GetCapability = True
        End Select
        Case TYPE_北京尚洋
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤消出院, _
                     support允许不设置医保项目
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_华东
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support结帐退个人帐户, _
                     support出院病人结算作废, _
                     support撤消出院
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_南京市
            Select Case cap业务
                Case support门诊预算, _
                     support结算使用个人帐户, _
                     support门诊必须传递明细, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊退费, _
                     support撤消出院
                     GetCapability = True
            End Select
        Case TYPE_江苏
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support记帐完成后上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support门诊预算, _
                     support挂号使用个人帐户, _
                     support撤消出院, _
                     support出院病人结算作废
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_重庆市
            Select Case cap业务
                Case support门诊退费, _
                     support结算使用个人帐户, _
                     support结帐退个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support出院结算必须出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support记帐完成后上传, _
                     support出院病人结算作废, _
                     support必须录入入出诊断, _
                     support中途结算仅处理已上传部分, _
                     support允许部份冲销单据, _
                     support挂号使用个人帐户, _
                     support撤消出院
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
                Case support门诊必须传递明细
                    GetCapability = True
            End Select
        Case TYPE_重庆松藻
            Select Case cap业务
                Case support门诊退费, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support未结清出院, support门诊预算
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
        Case TYPE_涪陵
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support允许不设置医保项目, _
                     support撤消出院
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_广元
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐完成后上传, _
                     support未结清出院, _
                     support医嘱上传, _
                     support出院结算必须出院, _
                     support门诊退费, _
                     support允许不设置医保项目, _
                     support门诊预算, _
                     support撤消出院
                     '支持的业务
                     GetCapability = True
            End Select
        Case TYPE_重庆壁山                 'add by zyc 2003-4-9,add by ccy 2003-11-28
            Select Case cap业务
                Case support门诊必须传递明细, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support必须录入入出诊断, _
                     support记帐完成后上传, _
                     support出院结算必须出院, _
                     support未结清出院, _
                     support门诊部分退现金, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support结算使用个人帐户, _
                     support记帐完成后上传, _
                     support撤消出院
                     '支持的业务
                     GetCapability = True
            End Select
            If Val(Get保险参数_壁山("适用地区")) = 2 Then
                If cap业务 = support门诊部分退现金 Then GetCapability = False
                If cap业务 = support门诊退费 Then GetCapability = True
                If cap业务 = support结帐退个人帐户 Then GetCapability = True
                If cap业务 = support记帐上传 Then GetCapability = False                                 '新增部分
                If cap业务 = support记帐作废上传 Then GetCapability = False                             '新增部分
'                If cap业务 = support必须录入入出诊断 Then GetCapability = False
            End If
            If cap业务 = support门诊预算 Then
                If Val(Get保险参数_壁山("适用地区")) = 2 Or Val(Get保险参数_壁山("适用地区")) = 1 Then
                    GetCapability = True
                Else
                    GetCapability = gbln特殊门诊
                End If
            End If
        Case TYPE_重庆中梁山
            Select Case cap业务
                Case support撤消出院, support未结清出院
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = True
            End Select
        Case TYPE_重庆银海版
            Select Case cap业务
            Case support撤消出院, support结算使用个人帐户, support门诊必须传递明细, _
            support结帐退个人帐户, support门诊退费, support未结清出院, support门诊预算, support允许冲销已结帐的记帐单据
                GetCapability = True
            End Select
        Case TYPE_自贡市
            '自贡医保改成固定参数，参数设置时已经无效。
            Select Case cap业务
                Case support门诊退费, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support未结清出院, _
                     support必须录入入出诊断, _
                     support结算帐户首先自付, _
                     support收费帐户首先自付              '收费帐户全自费、结算帐户全自费、结算帐户超限全不支持
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
        Case TYPE_泸州市, TYPE_铜仁
            '自贡医保改成固定参数，参数设置时已经无效。
            Select Case cap业务
                Case support门诊退费, _
                     support未结清出院, _
                     support必须录入入出诊断, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support医嘱上传, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support结算帐户首先自付, _
                     support收费帐户首先自付              '收费帐户全自费、结算帐户全自费、结算帐户超限全不支持
                    '支持的业务
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
        Case TYPE_贵阳市
            Select Case cap业务
                Case support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support门诊退费
                    '支持的业务
                    GetCapability = True
                Case support门诊必须传递明细
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
        Case TYPE_云南省, TYPE_昆明市
            Select Case cap业务
                Case support门诊退费, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support门诊必须传递明细, _
                     support必须录入入出诊断, _
                     support未结清出院, _
                     support结算使用个人帐户
                    '支持的业务
                    GetCapability = True
                Case support撤消出院
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
        Case TYPE_云南建水
            '由于不支持门诊退费，所以单列
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付, support必须录入入出诊断
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support结算使用个人帐户
                    GetCapability = True
            End Select
        Case TYPE_成都市
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support记帐完成后上传
                    GetCapability = True
                Case support记帐上传, support记帐作废上传, support医嘱上传
                    '200308z012:不实时上传
                    GetCapability = False
            End Select
        Case TYPE_成都莲合
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
                Case support未结清出院
                    GetCapability = True
            End Select
        Case TYPE_成都郊县
            Select Case cap业务
                Case support收费帐户全自费, support收费帐户首先自付, support记帐上传, support记帐作废上传, support医嘱上传, support记帐完成后上传
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
            End Select
        Case TYPE_新都
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support医嘱上传, _
                     support记帐完成后上传
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
            End Select
        Case TYPE_成都南充
            Select Case cap业务
                Case support记帐上传, support记帐作废上传, support医嘱上传, _
                    support允许不设置医保项目, support未结清出院, support撤消出院
                    GetCapability = True
            End Select
        Case TYPE_咸阳市
            Select Case cap业务
                Case support结算使用个人帐户, support收费帐户全自费, support收费帐户首先自付
                    GetCapability = True
                Case support门诊部分退现金  '因为门诊可以退现金
                    GetCapability = True
            End Select
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            Select Case cap业务
                Case support门诊退费, support结算使用个人帐户, _
                support门诊必须传递明细, support撤消出院, support未结清出院, support结帐退个人帐户
                    GetCapability = True
                'Modified by 朱玉宝 20031218 地区：福州
                Case support门诊预算
                    If int险类 = TYPE_福建巨龙 Then
                        GetCapability = True
                    Else
                        GetCapability = False
                    End If
                Case support必须录入入出诊断
                    If int险类 = TYPE_南平市 Then
                        GetCapability = True
                    Else
                        GetCapability = False
                    End If
            End Select
        Case type_米易
            Select Case cap业务
                Case support门诊预算, support记帐上传, support记帐作废上传, _
                support结算使用个人帐户, support撤消出院, support未结清出院, support必须录入入出诊断
                    GetCapability = True
            End Select
        Case TYPE_四川眉山
            Select Case cap业务
                Case support门诊预算, support结算使用个人帐户, support门诊退费, _
                support撤消出院, support未结清出院, support必须录入入出诊断
                    GetCapability = True
            End Select
        Case TYPE_沈阳市
            Select Case cap业务
                Case support门诊预算, support结算使用个人帐户, support门诊退费, support撤消出院, _
                support记帐上传, support记帐作废上传, support医嘱上传, support允许部份冲销单据, _
                support未结清出院, support结帐退个人帐户, support门诊必须传递明细       'support必须录入入出诊断
                    GetCapability = True
            End Select
        Case type_大连市, type_大连开发区
            Select Case cap业务
                Case support门诊预算, _
                     support门诊退费, _
                     support未结清出院, _
                     support结算使用个人帐户, _
                     support门诊必须传递明细, _
                     support出院病人结算作废, _
                     support必须录入入出诊断, _
                     support撤消出院, _
                     support出院无实际交易, _
                     support结帐退个人帐户
                     
                    GetCapability = True
            End Select
        Case TYPE_北京
            Select Case cap业务
                Case support门诊预算, _
                     support门诊退费, _
                     support未结清出院, _
                     support结算使用个人帐户, _
                     support必须录入入出诊断, _
                     support撤消出院, _
                     support结帐退个人帐户, _
                     support允许不设置医保项目
                    GetCapability = True
            End Select
        Case TYPE_乐山
            Select Case cap业务
                'Modified by ZYB 2004-08-17
                Case support未结清出院, support挂号使用个人帐户, _
                support结算使用个人帐户, support撤消出院, support门诊预算, _
                support门诊退费, support记帐上传, support记帐完成后上传, support记帐作废上传, support医嘱上传
                    GetCapability = True
            End Select
        Case TYPE_重大校园卡
            Select Case cap业务
                Case support门诊退费, _
                     support预交退个人帐户, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, support挂号使用个人帐户, support允许不设置医保项目, support门诊预算
                    GetCapability = True
            End Select
        Case TYPE_重庆渝北
            Select Case cap业务
               Case support门诊退费, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support必须录入入出诊断, _
                     support门诊必须传递明细, _
                     support撤消出院
                    '支持的业务
                    GetCapability = True
            End Select
        Case TYPE_兴安
            '刘兴宏:20040827
            Select Case cap业务
                Case support门诊退费, support门诊预算, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support必须录入入出诊断, _
                     support门诊必须传递明细, _
                     support撤消出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                    '支持的业务
                    GetCapability = True
            End Select
        Case TYPE_黔南
             '刘兴宏:20041022
            Select Case cap业务
                Case support门诊退费, support门诊预算, _
                     support结算使用个人帐户, _
                     support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support必须录入入出诊断, _
                     support门诊必须传递明细, _
                     support撤消出院, _
                     support允许不设置医保项目, _
                     support出院病人结算作废
                     'support出院结算必须出院
                    '支持的业务
                    GetCapability = True
            End Select
        Case TYPE_成都德阳
            '刘兴宏:20041103
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support撤消出院, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                    '支持的业务
                    GetCapability = True
            End Select
        Case TYPE_成都内江
            '刘兴宏:20041207
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                     support未结清出院, _
                     support记帐上传, _
                     support记帐作废上传, _
                     support记帐完成后上传, _
                     support撤消出院, _
                     support出院病人结算作废, _
                     support出院结算必须出院
                    '支持的业务
                    GetCapability = True
            End Select
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            Select Case cap业务
                Case support收费帐户全自费, _
                     support收费帐户首先自付, _
                    support门诊必须传递明细, _
                    support允许不设置医保项目
                    '支持的业务
                    GetCapability = True
            End Select
            
        Case Is > 900
            Select Case cap业务
                Case support门诊退费, _
                     support预交退个人帐户, _
                     support结帐退个人帐户, _
                     support结算使用个人帐户, _
                     support必须录入入出诊断, _
                     support未结清出院
                    
                    GetCapability = True
                Case support撤消出院
                    GetCapability = True
                Case support门诊部分退现金
                    GetCapability = False  '因为支持门诊退费
            End Select
    End Select
    If cap业务 = support分币处理 Then
        Select Case int险类
        Case TYPE_浙江, TYPE_余姚
            GetCapability = True
        Case Else
            GetCapability = False
        End Select
    End If
    
    If cap业务 = support挂号使用个人帐户 Then
        Select Case int险类
            Case TYPE_成都市
                GetCapability = True
            Case TYPE_重庆壁山
                GetCapability = True      '2003-4-9modify zyc
        End Select
    End If
    
    If cap业务 = support门诊预算 Then
        Select Case int险类
            Case TYPE_重庆市
                '重庆市所有的门诊病人都使用预算
                GetCapability = True
            Case TYPE_成都郊县, TYPE_新都
                GetCapability = True
            Case TYPE_贵阳市
                '贵阳的所有的门诊病人都使用预算
                GetCapability = True
            Case TYPE_云南省, TYPE_昆明市
                gstrSQL = "select nvl(病种ID,0) 病种ID from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & int险类
                Call OpenRecordset(rsTemp, "医保接口")
                If rsTemp.EOF = False Then
                    '有特殊病的病人需要预算
                    GetCapability = IIf(rsTemp("病种ID") > 0, True, False)
                End If
                '公务员预结算 gzh
                If GetCapability = False Then GetCapability = mbln公务员
            Case TYPE_云南建水
                GetCapability = True
            Case TYPE_自贡市
                gstrSQL = "select 在职 from 保险帐户 where 病人ID=" & lng病人ID & " and 险类=" & int险类
                Call OpenRecordset(rsTemp, "医保接口")
                If rsTemp.EOF = False Then
                    GetCapability = IIf(rsTemp("在职") = 3, True, False)
                End If
            Case TYPE_泸州市, TYPE_铜仁
                '所有病人，因为有可能有慢病或公务员支持预结算
                GetCapability = True
            Case Is > 900
                gstrSQL = "select 参数值  from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='收费使用医保基金'"
                Call OpenRecordset(rsTemp, "医保接口")
                If rsTemp.EOF = False Then
                    GetCapability = IIf(rsTemp("参数值") = 1, True, False)
                End If
        End Select
    End If
    
    If int险类 = TYPE_贵阳市 Then
        '针对贵阳医保的特殊处理 XQ 2003-05-12
        If cap业务 = support门诊连续收费 Then
            gstrSQL = "select 参数值 值 from 保险参数 where 险类=" & int险类 & " and Nvl(中心,0)=0 and 参数名='门诊连续收费'"
            Call OpenRecordset(rsTemp, "医保接口")
            If rsTemp.EOF = False Then
                GetCapability = IIf(rsTemp("值") = 1, True, False)
            End If
            Exit Function
        End If
        If cap业务 = support门诊收费完成后验证 Then
            GetCapability = True
            Exit Function
        End If
    End If
        
    If int险类 = TYPE_重庆松藻 Then
        gstrSQL = ""
        If cap业务 = support允许不设置医保项目 Then
            gstrSQL = "select 参数值 值 from 保险参数 where 险类=" & int险类 & " and Nvl(中心,0)=0 and 参数名='允许不设置医保项目'"
        End If
        If gstrSQL <> "" Then
            Call OpenRecordset(rsTemp, "医保接口")
            If rsTemp.EOF = False Then
                GetCapability = IIf(rsTemp("值") = 1, True, False)
            End If
        End If
    End If
    
    If int险类 > 900 Then
        gstrSQL = ""
        '读保险参数表
        If cap业务 = support收费帐户全自费 Then
            gstrSQL = "select substr(参数值,1,1) 值 from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='收费个人帐户使用范围'"
        ElseIf cap业务 = support收费帐户首先自付 Then
            gstrSQL = "select substr(参数值,2,1) 值 from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='收费个人帐户使用范围'"
        ElseIf cap业务 = support结算帐户全自费 Then
            gstrSQL = "select substr(参数值,1,1) 值 from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='结算个人帐户使用范围'"
        ElseIf cap业务 = support结算帐户首先自付 Then
            gstrSQL = "select substr(参数值,2,1) 值 from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='结算个人帐户使用范围'"
        ElseIf cap业务 = support结算帐户超限 Then
            gstrSQL = "select substr(参数值,3,1) 值 from 保险参数 where 险类=" & int险类 & " and 中心 is null and 参数名='结算个人帐户使用范围'"
        End If
        
        If gstrSQL <> "" Then
            Call OpenRecordset(rsTemp, "医保接口")
            If rsTemp.EOF = False Then
                GetCapability = IIf(rsTemp("值") = 1, True, False)
            End If
        End If
    End If
End Function

Public Function GetItemInsure(病人ID As Long, 收费细目ID As Long, 金额 As Currency, _
    ByVal bln门诊 As Boolean, Optional ByVal int险类 As Integer) As String
'功能：获取收费项目对应的医保信息
'参数：金额=收入项目级的实收金额。
'      int险类=指定险类,缺省为当前险类(如果已连接)
'返回："保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要"
'调用模块：
    Dim rsTmp As ADODB.Recordset ', bln全额统筹 As Boolean
    Dim rs特准 As New ADODB.Recordset
    Dim bln保险项目 As Boolean, lng大类ID As Long, cur统筹金额 As Currency
    Dim lng服务对象 As Long, str项目编码 As String
    
    If int险类 = 0 Then int险类 = gintInsure
    
    bln保险项目 = False
    lng大类ID = 0
    cur统筹金额 = 0
    
    'bln全额统筹 = Is全额统筹(病人ID)
    
    'Modified by ZYB 2002-10-30
    '如果参数“允许不设置医保项目”为真，则返回实收金额（意思是全部进入统筹）
    If GetCapability(support允许不设置医保项目) Then
        GetItemInsure = "1;0;" & 金额 & ";"
        Exit Function
    End If
    If int险类 <> TYPE_自贡市 Then
        If int险类 <> TYPE_四川眉山 Then
            gstrSQL = " Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*" & _
                    " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=" & int险类 & ") B" & _
                    " Where A.大类ID=B.ID(+) And A.险类=" & int险类 & " And A.收费细目ID=" & 收费细目ID
        Else
            gstrSQL = "Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*,C.服务对象 项目服务对象 " & _
                     " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=" & int险类 & ") B, " & _
                     " (Select * From 保险项目 Where 险类=" & int险类 & ") C " & _
                     " Where A.大类ID=B.ID(+) And A.项目编码=C.编码(+) And A.险类=" & int险类 & " And A.收费细目ID=" & 收费细目ID
        End If
    Else
        '如果病种中有特准项目，以另外计算
        gstrSQL = "SELECT C.收费细目ID " & _
                 " FROM 保险帐户 A,保险特准项目 C,(Select * From ZLYB.病种信息 Where 险类=" & int险类 & " And 病人ID=" & 病人ID & ") B " & _
                 " WHERE A.病人ID=" & 病人ID & " AND A.险类=" & int险类 & " AND C.收费细目ID=" & 收费细目ID & _
                 " And C.病种ID=B.病种ID"
        Call OpenRecordset(rs特准, "医保接口")
        If rs特准.RecordCount = 0 Then
            '按照普通病人，正常计算
            gstrSQL = "Select A.大类ID,A.项目编码,A.项目名称,A.是否医保 as 项目是否医保,B.*" & _
                " From 保险支付项目 A,(Select * From 保险支付大类 Where 险类=" & int险类 & ") B" & _
                " Where A.大类ID=B.ID(+) And A.险类=" & int险类 & " And A.收费细目ID=" & 收费细目ID
        Else
            '使用原有大类编码
            gstrSQL = "SELECT B.ID 大类ID,A.项目编码,A.项目名称,A.是否医保 AS 项目是否医保,B.* " & _
                " FROM 保险支付项目 A,(SELECT * FROM 保险支付大类 WHERE 险类 =" & int险类 & ") B,保险项目 C " & _
                " WHERE A.险类=" & int险类 & " AND A.收费细目ID=" & 收费细目ID & " AND A.项目编码=C.编码 AND A.险类=C.险类 AND C.大类编码=B.编码(+)"
            
        End If
    End If
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call OpenRecordset(rsTmp, "医保接口")
    
    If Not rsTmp.EOF Then
        bln保险项目 = True
        lng大类ID = IIf(IsNull(rsTmp!大类ID), 0, rsTmp!大类ID)
        str项目编码 = IIf(IsNull(rsTmp!项目编码), "", rsTmp!项目编码)
        If lng大类ID <> 0 Then
            'Modified by 朱玉宝 20031218 地区：福州
'            If int险类 <> TYPE_福建巨龙 Then
            If Not (int险类 = TYPE_福建巨龙 Or int险类 = TYPE_福建省 Or int险类 = TYPE_福州市 Or int险类 = TYPE_南平市) Then
                bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!是否医保), 0, rsTmp!是否医保) = 1)
            End If
            If IIf(IsNull(rsTmp!算法), 0, rsTmp!算法) = 1 Then
                cur统筹金额 = 金额 * (IIf(IsNull(rsTmp!统筹比额), 0, rsTmp!统筹比额) / 100)
                'Modified by 朱玉宝 20031218 地区：福州
'                If int险类 <> TYPE_福建巨龙 Then
                If Not (int险类 = TYPE_福建巨龙 Or int险类 = TYPE_福建省 Or int险类 = TYPE_福州市 Or int险类 = TYPE_南平市) Then
                    bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!统筹比额), 0, rsTmp!统筹比额) > 0)
                End If
            End If
            
            '如果服务对象不对，也认为它不是医保项目
            lng服务对象 = IIf(IsNull(rsTmp("服务对象")), 3, rsTmp("服务对象")) '缺省认为该大类可同时服务于门诊与住院病人
            If bln门诊 = True Then
                If lng服务对象 <> 1 And lng服务对象 <> 3 Then bln保险项目 = False
            Else
                If lng服务对象 <> 2 And lng服务对象 <> 3 Then bln保险项目 = False
            End If
            
            '再判断各个项目的服务对象
            If int险类 = TYPE_四川眉山 Then
                lng服务对象 = IIf(IsNull(rsTmp("项目服务对象")), 3, rsTmp("项目服务对象")) '缺省认为该项目可同时服务于门诊与住院病人
                If bln门诊 = True Then
                    If lng服务对象 <> 1 And lng服务对象 <> 3 Then bln保险项目 = False
                Else
                    If lng服务对象 <> 2 And lng服务对象 <> 3 Then bln保险项目 = False
                End If
            End If
        End If
        bln保险项目 = bln保险项目 And (IIf(IsNull(rsTmp!项目是否医保), 0, rsTmp!项目是否医保) = 1)
        
        '全额统筹金额
        'If bln全额统筹 And bln保险项目 Then cur统筹金额 = 金额
    Else
'        If int险类 = TYPE_自贡市 Then
'            MsgBox "未找到对应的保险大类，请检查保险项目是否正确设置。", vbInformation, gstrSysName
'            Exit Function
'        End If
    End If
    GetItemInsure = IIf(bln保险项目, 1, 0) & ";" & lng大类ID & ";" & cur统筹金额 & ";" & str项目编码
End Function

Public Function TranChargeDetail(ByVal int类别 As Integer, ByVal str单据号 As String, ByVal int性质 As Integer, ByVal int状态 As Integer, _
        str消息 As String, Optional ByVal lng病人ID As Long = 0) As Boolean
'功能: 把发生的费用明细传输到医保中间数据库
'参数:  int类别    1:表示门诊收费；２：住院计帐
'       str单据号  NO
'       int性质    记录性质
'       int状态    记录状态
'       str消息    如果传输过程中有提醒，传回前台程序完成（避免长时间的死锁）
'       lng病人ID  默认为0，表示传输整张单据，否则为单据中指定病人的。（主要是因为医嘱在保存记帐单时，是分病人在提交数据而不是一起提交）
'返回: 传输成功与否
'------------------------------------------------------------------------------------------------------------------
'调用模块：
'   1121-门诊收费:记录修改;门诊收费;门诊退费
'   1133-住院记帐:记录修改;住院划价;住院记帐;住院销帐;专项记帐
'   1134-分散记帐:记录修改;住院记帐;住院销帐;专项记帐
'   1135-医技记帐:记录修改;住院记帐;住院销帐;专项记帐
'医嘱系统 (5100)
'   1261 - 医嘱摆药生成
'   1263 - 执行非药医嘱
'
'医护系统
'   402 - 医嘱摆药处理
'   403 - 执行非药医嘱


    Dim intMouse As Integer
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    Select Case gintInsure
        Case TYPE_余姚
            If int类别 = 1 Then
                Screen.MousePointer = intMouse
                str消息 = "门诊记帐数据不能上传到医保前置机，请通过门诊费用管理收取门诊费用"
                TranChargeDetail = False
                Exit Function
            End If
            TranChargeDetail = 记帐传输_余姚(str单据号, int性质, str消息, lng病人ID)
        Case TYPE_吉林
            '刘兴宏:20040927
            TranChargeDetail = 处方登记_吉林(int性质, int状态, str单据号)
        Case TYPE_北京尚洋
            TranChargeDetail = False
        Case TYPE_华东
            TranChargeDetail = 记帐传输_华东(str单据号, int性质, str消息, lng病人ID)
        Case TYPE_江苏
            TranChargeDetail = True
        Case TYPE_涪陵
            TranChargeDetail = 记帐传输_涪陵(str单据号, int性质, str消息, lng病人ID)
        Case TYPE_广元
            TranChargeDetail = 记帐传输_广元(str单据号, int性质, str消息, lng病人ID)
        Case TYPE_重庆市
            If int类别 = 2 Then
                '只处理住院记帐
                If int状态 = 1 Then
                    TranChargeDetail = 记帐传输_重庆(str单据号, int性质, str消息, lng病人ID)
                Else
                    '作废单据
                    TranChargeDetail = 记帐作废_重庆(str单据号, int性质, str消息)
                End If
            End If
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            TranChargeDetail = False
        Case TYPE_重庆壁山
            If int类别 = 2 Then '住院计帐
                TranChargeDetail = 记帐传输_壁山(str单据号, int性质, int状态, lng病人ID)
            End If
        Case TYPE_重庆中梁山
            If int类别 = 2 Then '住院计帐
                TranChargeDetail = 记帐传输_中梁山(str单据号, int性质, int状态, lng病人ID)
            End If
        Case TYPE_自贡市
            TranChargeDetail = False
        Case TYPE_泸州市
            If int类别 = 2 Then
                '只处理住院记帐
                If int状态 = 1 Then
                    TranChargeDetail = 记帐传输_泸州(str单据号, int性质, lng病人ID)
                Else
                    TranChargeDetail = 记帐作废_泸州(str单据号, int性质, lng病人ID)
                End If
            End If
        Case TYPE_铜仁
            If int类别 = 2 Then
                '只处理住院记帐
                If int状态 = 1 Then
                    TranChargeDetail = 记帐传输_铜仁(str单据号, int性质, lng病人ID)
                Else
                    TranChargeDetail = 记帐作废_铜仁(str单据号, int性质, lng病人ID)
                End If
            End If
        Case TYPE_贵阳市
            TranChargeDetail = False
        Case TYPE_云南省, TYPE_昆明市
            TranChargeDetail = False
        Case TYPE_云南建水
            TranChargeDetail = False
        Case TYPE_成都市
            '200308z012:不实时上传
            TranChargeDetail = False
'            If int类别 = 2 Then '住院计帐
'                TranChargeDetail = 记帐传输_成都(str单据号, int性质, int状态, lng病人ID)
'            End If
        Case TYPE_成都郊县
            If int类别 = 2 Then '住院计帐
                TranChargeDetail = 记帐传输_成都郊县(str单据号, int性质, int状态, lng病人ID)
            End If
        Case TYPE_新都
            If int类别 = 2 Then '住院计帐
                TranChargeDetail = 记帐传输_新都(str单据号, int性质, int状态, lng病人ID)
            End If
        Case TYPE_成都莲合
            TranChargeDetail = 传输明细_莲合(str单据号, int性质, int状态, int类别)
        Case TYPE_成都南充
            If int类别 = 2 Then
                '只处理住院记帐
                If int状态 = 1 Then
                    TranChargeDetail = 处方登记_成都南充(int性质, int状态, str单据号)
                Else
                    '作废单据
                    TranChargeDetail = 处方删除_成都南充(int性质, int状态, str单据号)
                End If
            End If
        Case type_米易
            If int类别 = 2 Then
                '只处理住院记帐
                TranChargeDetail = 处方登记_米易(int性质, int状态, str单据号)
            End If
        Case TYPE_沈阳市
            If int类别 = 2 Then
                If int状态 = 1 Then
                    TranChargeDetail = 处方登记_沈阳市(int性质, int状态, str单据号)
                Else
                    TranChargeDetail = 处方作废_沈阳市(int性质, int状态, str单据号)
                End If
            End If
        Case type_大连市, type_大连开发区
            If int类别 = 2 Then
              '刘兴宏(200311)
                '只处理住院记帐
                TranChargeDetail = 处方登记_大连(int性质, int状态, str单据号)
            End If
        Case TYPE_乐山
            TranChargeDetail = 上传处方_乐山(int性质, int状态, str单据号)
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            TranChargeDetail = False
        Case TYPE_重庆渝北
             '刘兴宏200407
             TranChargeDetail = 处方登记_重庆渝北(int性质, int状态, str单据号)
        Case TYPE_兴安
            '刘兴宏:20040827
            TranChargeDetail = 处方登记_兴安(int性质, int状态, str单据号)
        Case TYPE_毕节
            TranChargeDetail = 处方上传_毕节(int性质, int状态, str单据号)
        Case TYPE_黔南
            '刘兴宏:20041022
            TranChargeDetail = 处方登记_黔南(int性质, int状态, str单据号)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            TranChargeDetail = 处方登记_成都德阳(int性质, int状态, str单据号)
        Case TYPE_成都内江
            '刘兴宏:20041207
            TranChargeDetail = 处方登记_成都内江(int性质, int状态, str单据号)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            TranChargeDetail = 处方登记_广元旺苍(int性质, int状态, str单据号)
        Case TYPE_渝北农医
            TranChargeDetail = 处方上传_渝北农医(int性质, int状态, str单据号)
        Case Is > 900
            TranChargeDetail = False
    End Select
    Screen.MousePointer = intMouse
End Function

Public Function LeaveDelSwap(lngPatiID As Long, lngPageID As Long) As Boolean
'功能：将撤消病人的出院信息
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1132-撤消出院
    Dim rsTmp As New ADODB.Recordset
    
    'On Error Resume Next
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=" & lngPatiID & " And A.主页ID=" & lngPageID & _
            " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=" & gintInsure
    Call OpenRecordset(rsTmp, "医保接口")
    
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Select Case gintInsure
        Case TYPE_宁海
            LeaveDelSwap = 出院登记撤销_宁海(lngPatiID, lngPageID)
        Case TYPE_余姚
            LeaveDelSwap = True
        Case TYPE_浙江
            LeaveDelSwap = 出院登记撤消_浙江(lngPatiID, lngPageID)
        Case TYPE_吉林
            '20040923刘兴宏加入
            LeaveDelSwap = 出院登记撤销_吉林(lngPatiID, lngPageID)
        Case TYPE_江苏, TYPE_南京市, TYPE_华东, TYPE_北京尚洋
            LeaveDelSwap = True
        Case TYPE_涪陵, TYPE_广元
            LeaveDelSwap = False
        Case TYPE_重庆市
            LeaveDelSwap = 出院登记撤销_重庆(lngPatiID, lngPageID)
        Case TYPE_重庆壁山
            LeaveDelSwap = True
        Case TYPE_重庆中梁山
            LeaveDelSwap = 出院登记撤消_中梁山(lngPatiID, lngPageID)
        Case TYPE_云南省, TYPE_昆明市
            LeaveDelSwap = 出院登记撤消_云南(lngPatiID, lngPageID)
        Case TYPE_成都南充
            LeaveDelSwap = 出院登记撤销_成都南充(lngPatiID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            LeaveDelSwap = 出院登记撤销_福建巨龙(lngPatiID)
        Case type_米易
            LeaveDelSwap = 出院登记撤销_米易(lngPatiID, lngPageID)
        Case TYPE_四川眉山
            LeaveDelSwap = 出院登记撤销_眉山(lngPatiID, lngPageID)
        Case TYPE_沈阳市
            LeaveDelSwap = 出院登记撤销_沈阳市(lngPatiID, lngPageID)
        Case type_大连市, type_大连开发区
            '刘兴宏(200311)
            LeaveDelSwap = 出院登记撤销_大连(lngPatiID, lngPageID)
        Case TYPE_乐山
            LeaveDelSwap = 出院登记撤销_乐山(lngPatiID, lngPageID)
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            LeaveDelSwap = 出院登记撤销_重大校园卡(lngPatiID, lngPageID)
        Case TYPE_重庆银海版
            LeaveDelSwap = 出院登记撤销_重庆银海版(lngPatiID, lngPageID)
        Case TYPE_重庆渝北
            LeaveDelSwap = 出院登记撤销_重庆渝北(lngPatiID, lngPageID)
        Case TYPE_北京
            LeaveDelSwap = 出院登记撤销_北京(lngPatiID, lngPageID)
        Case TYPE_兴安
            '刘兴宏:20040827
            LeaveDelSwap = 出院登记撤销_兴安(lngPatiID, lngPageID)
        Case TYPE_毕节
            LeaveDelSwap = 出院登记撤销_毕节(lngPatiID, lngPageID)
        Case TYPE_黔南
            '刘兴宏:20041022
            LeaveDelSwap = 出院登记撤销_黔南(lngPatiID, lngPageID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            LeaveDelSwap = 出院登记撤销_成都德阳(lngPatiID, lngPageID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            LeaveDelSwap = 出院登记撤销_成都内江(lngPatiID, lngPageID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            LeaveDelSwap = 出院登记撤销_广元旺苍(lngPatiID, lngPageID)
        Case TYPE_渝北农医
            LeaveDelSwap = 出院登记撤销_渝北农医(lngPatiID, lngPageID)
        Case Is > 900
            LeaveDelSwap = 出院登记撤消_中联(lngPatiID, lngPageID)
    End Select
End Function

Public Function RegistSwap(lng结帐ID As Long, cur金额 As Currency) As Boolean
'功能：门诊挂号结算
'参数：lng结帐ID=挂号单保存后的结帐ID,cur金额=个人帐户支付金额
'说明：目前挂号交易只支持全部使用个人帐户结算,不支持其它结算方式及预结算
'使用模块：1115-自助挂号
    Dim str医保号 As String, str密码 As String
    Dim lng病人ID As Long, str卡号 As String
    Dim rsTmp As New ADODB.Recordset
    
    If GetCapability(support挂号使用个人帐户) And gintInsure <> TYPE_江苏 Then
        With rsTmp
            '松藻医保不需要对每个项目设置对应的医保项目
            If Not GetCapability(support允许不设置医保项目) Then
                gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                    " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=" & gintInsure & ") A," & _
                    " (Select 收费细目ID,'Y' as TEST From 病人费用记录 Where 结帐ID=" & lng结帐ID & ") B,收费细目 C" & _
                    " Where B.收费细目ID=C.ID And B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
                Call OpenRecordset(rsTmp, "医保接口")
                
                If Not .EOF Then
                    MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                    RegistSwap = False: Exit Function
                End If
            End If
            
            If GetCapability(support门诊必须传递明细) = False Then
                gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                    " From 病人预交记录 A,保险帐户 B " & _
                    " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
                    " And A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lng结帐ID
            Else
                '病人该次挂号可能没有发生个人帐户费用，便也要传递费用明细
                gstrSQL = "Select B.病人ID,B.卡号,B.医保号,B.密码 " & _
                    " From 病人预交记录 A,保险帐户 B " & _
                    " Where A.病人ID=B.病人ID And B.险类=" & gintInsure & _
                    "       And A.结帐ID=" & lng结帐ID
            End If
            Call OpenRecordset(rsTmp, "医保接口")
            
            If .RecordCount = 0 Then
                MsgBox "没有使用保险结算，可以不向医保交易", vbExclamation, gstrSysName
                RegistSwap = True: Exit Function
            End If
        End With
        
        lng病人ID = rsTmp!病人ID
        str卡号 = TrimStr(IIf(IsNull(rsTmp!卡号), "", rsTmp!卡号))
        str医保号 = TrimStr(IIf(IsNull(rsTmp!医保号), str卡号, rsTmp!医保号))
        str密码 = TrimStr(IIf(IsNull(rsTmp!密码), "", rsTmp!密码))
    End If
    
    Select Case gintInsure
        Case TYPE_江苏
            RegistSwap = 挂号结算_江苏(lng结帐ID, cur金额)
        Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_尚洋
            RegistSwap = False
        Case TYPE_重庆市
            RegistSwap = 挂号结算_重庆(lng结帐ID)
        Case TYPE_重庆松藻
            RegistSwap = False
        Case TYPE_重庆壁山
            RegistSwap = True
        Case TYPE_自贡市
            RegistSwap = False
        Case TYPE_泸州市
            'Modified by 朱玉宝 2004-01-07
            RegistSwap = 挂号结算_泸州(lng结帐ID)
        Case TYPE_铜仁
            RegistSwap = False
        Case TYPE_贵阳市
            'Modified By 朱玉宝 2004-05-25 原因：医保接口变动
            '------------------------------------------------
            RegistSwap = 门诊挂号_贵阳(lng结帐ID, lng病人ID)
            '------------------------------------------------
        Case TYPE_云南省, TYPE_昆明市
            RegistSwap = False
        Case TYPE_云南建水
            RegistSwap = False
        Case TYPE_成都市
            RegistSwap = 挂号结算_成都(lng结帐ID, lng病人ID, str医保号, str密码, str卡号)
        Case TYPE_成都莲合
            RegistSwap = False
        Case TYPE_成都郊县, TYPE_新都
            RegistSwap = False
        Case TYPE_咸阳市
            RegistSwap = False
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            RegistSwap = 挂号结算_福建巨龙(lng结帐ID)
        Case type_米易
            RegistSwap = False
        Case TYPE_沈阳市
            RegistSwap = 挂号结算_沈阳市(lng结帐ID)
        Case TYPE_四川眉山
            RegistSwap = 挂号结算_眉山(lng结帐ID, cur金额)
        Case type_大连市, type_大连开发区
            '刘兴宏(200403)
            RegistSwap = 挂号结算_大连(lng结帐ID)
        Case TYPE_乐山
            RegistSwap = False
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            RegistSwap = 挂号结算_重大校园卡(lng结帐ID)
        Case TYPE_重庆渝北
            '刘兴宏20040713
            RegistSwap = 挂号结算_重庆渝北(lng结帐ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            RegistSwap = False
        Case TYPE_黔南
            '刘兴宏:20041022
            RegistSwap = False
        Case TYPE_成都德阳
            '刘兴宏:20041103
            RegistSwap = False
        Case TYPE_成都内江
            '刘兴宏:20041207
            RegistSwap = False
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            RegistSwap = False
        Case TYPE_渝北农医
            RegistSwap = True
        Case Is > 900
            RegistSwap = False
    End Select
End Function

Public Function RegistDelSwap(lng结帐ID As Long) As Boolean
'功能：门诊挂号作废交易
'参数：lng结帐ID=作废前的挂号记录的结帐ID
'返回：交易成功返回true；否则，返回false
'调用模块：
    Dim strSelfNo As String, strSelfPwd As String, str卡号 As String
    Dim rsTmp As New ADODB.Recordset
    Dim curMoney As Currency
    Dim lng病人ID As Long
    
    If GetCapability(support挂号使用个人帐户) Then
        '松藻医保不需要对每个项目设置对应的医保项目
        If Not GetCapability(support允许不设置医保项目) Then
            gstrSQL = "Select 名称 From (Select C.名称,A.TEST From" & _
                " (Select 收费细目ID,'X' as TEST From 保险支付项目 Where 险类=" & gintInsure & ") A," & _
                " (Select 收费细目ID,'Y' as TEST From 病人费用记录 Where 结帐ID=" & lng结帐ID & ") B,收费细目 C" & _
                " Where B.收费细目ID=C.ID and B.收费细目ID=A.收费细目ID(+)) Where Test is NULL"
            Call OpenRecordset(rsTmp, "医保接口")
            
            If Not rsTmp.EOF Then
                MsgBox "费用中包含未设置保险支付项目的收费项目（" & rsTmp("名称") & "），不能执行医保交易。", vbExclamation, gstrSysName
                RegistDelSwap = False: Exit Function
            End If
        End If
        
        '读出个人帐户使用金额
        If GetCapability(support门诊必须传递明细) = False Then
            gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
                " Where A.结算方式 in ('个人帐户','医保基金') And A.结帐ID=" & lng结帐ID
        Else
            '病人该次收费可能没有发生个人帐户费用，便也要传递费用明细
            gstrSQL = "Select A.病人ID,A.冲预交,A.结算方式 From 病人预交记录 A " & _
                " Where A.结帐ID=" & lng结帐ID
        
        End If
        
        Call OpenRecordset(rsTmp, "门诊退费")
        If rsTmp.RecordCount = 0 Then
            MsgBox "没有使用个人帐户或医保基金结算，可以不向医保交易", vbExclamation, gstrSysName
            RegistDelSwap = True
            Exit Function
        End If
        
        lng病人ID = rsTmp("病人ID")
        
        Do Until rsTmp.EOF
            If rsTmp("结算方式") = "个人帐户" Then
                curMoney = curMoney + rsTmp("冲预交")
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    Select Case gintInsure
        Case TYPE_江苏
            RegistDelSwap = 挂号结算冲销_江苏(lng结帐ID)
        Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_尚洋
            RegistDelSwap = False
        Case TYPE_重庆市
            RegistDelSwap = 门诊结算冲销_重庆(lng结帐ID, 0, 0)
        Case TYPE_重庆松藻
            RegistDelSwap = False
        Case TYPE_重庆壁山
            RegistDelSwap = False
        Case TYPE_自贡市
            RegistDelSwap = False
        Case TYPE_泸州市
            'Modified by 朱玉宝 2004-01-07
            RegistDelSwap = 挂号结算冲销_泸州(lng结帐ID)
        Case TYPE_铜仁
            RegistDelSwap = False
        Case TYPE_云南省, TYPE_昆明市
            RegistDelSwap = False
        Case TYPE_云南建水
            RegistDelSwap = False
        Case TYPE_贵阳市
            RegistDelSwap = 门诊结算冲销_贵阳(lng结帐ID, 0, 0)
        Case TYPE_成都市
            RegistDelSwap = False
        Case TYPE_成都莲合
            RegistDelSwap = False
        Case TYPE_成都郊县, TYPE_新都
            RegistDelSwap = False
        Case TYPE_咸阳市
            RegistDelSwap = False
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            RegistDelSwap = 挂号结算冲销_福建巨龙(lng结帐ID)
        Case type_米易
            RegistDelSwap = False
        Case TYPE_四川眉山
            RegistDelSwap = 挂号结算冲销_眉山(lng结帐ID)
        Case type_大连市, type_大连开发区
            '刘兴宏(200311)
            RegistDelSwap = 挂号冲销_大连(lng结帐ID)
        Case TYPE_沈阳市
            RegistDelSwap = 挂号冲销_沈阳市(lng结帐ID)
        Case TYPE_乐山
            RegistDelSwap = False
        Case TYPE_重大校园卡
            '刘兴宏(200403)
            RegistDelSwap = 挂号冲销_重大校园卡(lng结帐ID)
        Case TYPE_重庆渝北
            '刘兴宏(20040714)
            RegistDelSwap = 挂号冲销_重庆渝北(lng结帐ID)
        Case TYPE_兴安
            '刘兴宏:20040827
            RegistDelSwap = False
        Case TYPE_黔南
            '刘兴宏:20041022
            RegistDelSwap = False
        Case TYPE_成都德阳
            '刘兴宏:20041103
            RegistDelSwap = False
        Case TYPE_成都内江
            '刘兴宏:20041207
            RegistDelSwap = False
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            RegistDelSwap = False
        Case TYPE_渝北农医
            RegistDelSwap = True
        Case Is > 900
            RegistDelSwap = False
    End Select
End Function

Public Function ComeInDelSwap(lngPatiID As Long, lngPageID As Long) As Boolean
'功能：将撤销入院信息发送医保前置服务器确认；
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1131-取消入院
    Dim rsTmp As New ADODB.Recordset
    
    'On Error Resume Next
    
    gstrSQL = "Select A.出院日期,A.出院病床,Decode(A.出院方式,'正常',0,'死亡',1,'转院',2,9) as 出院方式,B.名称,D.住院号,Sysdate as 经办时间," & _
            " C.卡号,C.医保号,C.密码,C.顺序号 " & _
            " From 病案主页 A,部门表 B,保险帐户 C,病人信息 D " & _
            " Where A.病人ID=D.病人ID And A.病人ID=" & lngPatiID & " And A.主页ID=" & lngPageID & _
            " And A.入院科室ID=B.ID And A.病人ID=C.病人ID And C.险类=" & gintInsure
    Call OpenRecordset(rsTmp, "医保接口")
    
    If rsTmp.EOF Then
        MsgBox "没有此病人或此病人不是医保病人！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Select Case gintInsure
        Case TYPE_宁海
            ComeInDelSwap = 入院登记撤销_宁海(lngPatiID, lngPageID)
        Case TYPE_余姚
            ComeInDelSwap = 出院登记_余姚(lngPatiID, lngPageID)
        Case TYPE_浙江
            ComeInDelSwap = 入院登记撤消_浙江(lngPatiID, lngPageID)
        Case TYPE_吉林
          '刘兴宏(20040924)
            ComeInDelSwap = 入院登记撤销_吉林(lngPatiID, lngPageID)
        Case TYPE_北京尚洋
            ComeInDelSwap = 出院登记_北京尚洋(lngPatiID, lngPageID)
        Case TYPE_华东
            ComeInDelSwap = 出院登记_华东(lngPatiID, lngPageID)
        Case TYPE_江苏
            ComeInDelSwap = 入院登记撤消_江苏(lngPatiID, lngPageID)
        Case TYPE_重庆市
            ComeInDelSwap = 出院登记_重庆(lngPatiID, lngPageID)
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ComeInDelSwap = 出院登记_松藻(lngPatiID, lngPageID)
        Case TYPE_重庆壁山
            ComeInDelSwap = 出院登记_壁山(lngPatiID, lngPageID)
        Case TYPE_重庆中梁山
            ComeInDelSwap = 出院登记_中梁山(lngPatiID, lngPageID)
        Case TYPE_自贡市
            ComeInDelSwap = 出院登记_中软(lngPatiID)
        Case TYPE_泸州市
            ComeInDelSwap = 出院登记_泸州(lngPatiID)
        Case TYPE_铜仁
            ComeInDelSwap = 出院登记_铜仁(lngPatiID)
        Case TYPE_贵阳市
            ComeInDelSwap = 出院登记_贵阳(lngPatiID, lngPageID)
        Case TYPE_云南省, TYPE_昆明市
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                Exit Function
            End If
            ComeInDelSwap = 出院登记_云南(lngPatiID, lngPageID, rsTmp!顺序号)
        Case TYPE_云南建水
            If IsNull(rsTmp!顺序号) Then
                MsgBox "未发现该病人的住院交易顺序号,不能执行交易！", vbInformation, gstrSysName
                Exit Function
            End If
            ComeInDelSwap = 入院登记撤销_云南建水(lngPatiID, lngPageID, rsTmp!顺序号)
        Case TYPE_成都市
            ComeInDelSwap = 出院登记_成都(lngPatiID, lngPageID, rsTmp)
        Case TYPE_成都莲合
            ComeInDelSwap = 出院登记_莲合(lngPatiID, lngPageID)
        Case TYPE_成都郊县
            ComeInDelSwap = 出院登记_成都郊县(lngPatiID, lngPageID)
        Case TYPE_新都
            ComeInDelSwap = 出院登记_新都(lngPatiID, lngPageID)
        Case TYPE_成都南充
            ComeInDelSwap = 出院登记_成都南充(lngPatiID)
        Case TYPE_咸阳市
            ComeInDelSwap = 出院登记_咸阳(lngPatiID)
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ComeInDelSwap = 出院登记_福建巨龙(lngPatiID, lngPageID, True)
        Case type_米易
            ComeInDelSwap = 入院登记撤销_米易(lngPatiID, lngPageID)
        Case TYPE_四川眉山
            ComeInDelSwap = 入院登记撤销_眉山(lngPatiID, lngPageID)
        Case TYPE_沈阳市
            ComeInDelSwap = 入院登记撤销_沈阳市(lngPatiID, lngPageID)
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            ComeInDelSwap = 入院登记撤销_大连(lngPatiID, lngPageID)
        Case TYPE_乐山
            ComeInDelSwap = 入院登记撤销_乐山(lngPatiID, lngPageID)
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            ComeInDelSwap = 入院登记撤销_重大校园卡(lngPatiID, lngPageID)
        Case TYPE_重庆银海版
            ComeInDelSwap = 入院登记撤销_重庆银海版(lngPatiID, lngPageID)
        Case TYPE_重庆渝北
            ComeInDelSwap = 入院登记撤销_重庆渝北(lngPatiID, lngPageID)
        Case TYPE_北京
            ComeInDelSwap = 入院登记撤销_北京(lngPatiID, lngPageID)
        Case TYPE_兴安
            '刘兴宏:20040827
            ComeInDelSwap = 入院登记撤销_兴安(lngPatiID, lngPageID)
        Case TYPE_毕节
            ComeInDelSwap = 入院登记撤销_毕节(lngPatiID, lngPageID)
        Case TYPE_黔南
            '刘兴宏:20041022
            ComeInDelSwap = 入院登记撤销_黔南(lngPatiID, lngPageID)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ComeInDelSwap = 入院登记撤销_成都德阳(lngPatiID, lngPageID)
        Case TYPE_成都内江
            '刘兴宏:20041207
            ComeInDelSwap = 入院登记撤销_成都内江(lngPatiID, lngPageID)
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ComeInDelSwap = 入院登记撤销_广元旺苍(lngPatiID, lngPageID)
        Case TYPE_渝北农医
            ComeInDelSwap = 入院登记撤销_渝北农医(lngPatiID, lngPageID)
        Case Is > 900
            '中联医保
            ComeInDelSwap = 出院登记_中联(lngPatiID, lngPageID)
    End Select
End Function

Public Function ModiPatiSwap(lngPatiID As Long, lngPageID As Long) As Boolean
'功能：在院病人的床位变动,转科,医生变化及诊断情况等相关信息变化时调用此接口
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回true；否则，返回false
'------------------------------------------------------------------------------------------------------------------
'调用模块：1132-病人入出管理(当病人状态发生变化时调用,如：床位变动,转科,医生变化等.)
    
    Select Case gintInsure
        Case TYPE_宁海
            ModiPatiSwap = 住院信息变动_宁海(lngPatiID, lngPageID)
        Case TYPE_余姚
            ModiPatiSwap = 转科转床_余姚(lngPatiID, lngPageID)
        Case TYPE_江苏
            ModiPatiSwap = 转科转床_江苏(lngPatiID, lngPageID)
        Case TYPE_涪陵, TYPE_广元
            ModiPatiSwap = True
        Case TYPE_重庆市
            ModiPatiSwap = True
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            ModiPatiSwap = True
        Case TYPE_重庆壁山
            ModiPatiSwap = True
        Case TYPE_重庆中梁山
            ModiPatiSwap = True
        Case TYPE_自贡市
            ModiPatiSwap = True
        Case TYPE_泸州市
            ModiPatiSwap = True
        Case TYPE_铜仁
            ModiPatiSwap = True
        Case TYPE_贵阳市
            ModiPatiSwap = True
        Case TYPE_云南省, TYPE_昆明市
            ModiPatiSwap = True
        Case TYPE_云南建水
            ModiPatiSwap = True
        Case TYPE_成都市
            ModiPatiSwap = True
        Case TYPE_成都莲合
            ModiPatiSwap = True
        Case TYPE_成都郊县, TYPE_新都
            ModiPatiSwap = True
        Case TYPE_成都南充
            ModiPatiSwap = True
        Case TYPE_咸阳市
            ModiPatiSwap = True
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            ModiPatiSwap = True
        Case type_米易
            ModiPatiSwap = True
        Case TYPE_四川眉山
            ModiPatiSwap = True
        Case TYPE_沈阳市
            ModiPatiSwap = True
        Case type_大连市, type_大连开发区
              '刘兴宏(200311)
            ModiPatiSwap = 在院病人信息_大连(lngPatiID, lngPageID)
        Case TYPE_乐山
            ModiPatiSwap = True
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            ModiPatiSwap = True
        Case TYPE_重庆银海版
            ModiPatiSwap = True
        Case TYPE_重庆渝北
            ModiPatiSwap = True
        Case TYPE_兴安
            '刘兴宏:20040827
            ModiPatiSwap = 在院病人信息_兴安(lngPatiID, lngPageID)
        Case TYPE_毕节
            ModiPatiSwap = 病人变动记录上传_毕节(lngPatiID, lngPageID)
        Case TYPE_黔南
            '刘兴宏:20041022
            ModiPatiSwap = True
        Case TYPE_成都德阳
            '刘兴宏:20041103
            ModiPatiSwap = True
        Case TYPE_成都内江
            '刘兴宏:20041207
            ModiPatiSwap = True
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            ModiPatiSwap = True
        Case Is > 900
            '中联医保
            ModiPatiSwap = True
    End Select
End Function

'刘兴宏:2004/08/30
'主要是区分在输入单据时提示信息，同时可以返回摘要的相关信息.

Public Function GetItemInfo(ByVal lngPatiID As Long, ByVal lngItemID As Long, Optional ByVal str摘要 As String, Optional intType As Integer = 0) As String
    '功能：获取医保项目的相关信息，目前仅大连使用，返回指定项目的报销比例
    '参数：lngPatiID-病人ID；lngItemID-收费细目ID
    '      str摘要=用于修改
    '      intType-调用类型(0-医嘱,1-门诊收费,2-住院记帐)
    '返回：不为空,返回摘要中；否则，表示无摘要
    '------------------------------------------------------------------------------------------------------------------
    '调用模块：所有记费模块使用，如住院记帐，门诊收费等
    
    Select Case gintInsure
        Case TYPE_吉林
            '刘兴宏20040923加入
            GetItemInfo = ""
        Case TYPE_涪陵, TYPE_广元, TYPE_华东, TYPE_余姚, TYPE_北京尚洋
            GetItemInfo = ""
        Case TYPE_重庆市
            GetItemInfo = ""
        Case TYPE_重庆松藻  'Modified by ZYB 2002-10-28
            GetItemInfo = ""
        Case TYPE_重庆壁山
            GetItemInfo = ""
        Case TYPE_重庆中梁山
            GetItemInfo = ""
        Case TYPE_自贡市
            GetItemInfo = ""
        Case TYPE_泸州市
            GetItemInfo = ""
        Case TYPE_铜仁
            GetItemInfo = ""
        Case TYPE_贵阳市
            GetItemInfo = ""
        Case TYPE_云南省, TYPE_昆明市
            GetItemInfo = ""
        Case TYPE_云南建水
            GetItemInfo = ""
        Case TYPE_成都市
            GetItemInfo = ""
        Case TYPE_成都莲合
            GetItemInfo = ""
        Case TYPE_成都郊县, TYPE_新都
            GetItemInfo = ""
        Case TYPE_成都南充
            GetItemInfo = ""
        Case TYPE_咸阳市
            GetItemInfo = ""
        Case TYPE_福建巨龙, TYPE_福建省, TYPE_福州市, TYPE_南平市
            GetItemInfo = ""
        Case type_米易
            GetItemInfo = ""
        Case TYPE_四川眉山
            GetItemInfo = ""
        Case TYPE_沈阳市
            GetItemInfo = ""
        Case type_大连市, type_大连开发区
            If intType = 0 Then
                '只有医嘱才有
                GetItemInfo = GetItemInfo_大连(lngPatiID, lngItemID)
            End If
        Case TYPE_乐山
            '如果该项目可能做为特殊项目才提示操作员，本明细是否做为特殊项目上传
            Dim int可做为特殊项目 As Integer
            Dim rsTemp As New ADODB.Recordset
            
            Call DebugTool("判断可否做为特殊项目")
            int可做为特殊项目 = 0
            gstrSQL = "Select 附注 From 保险支付项目 Where 险类=" & TYPE_乐山 & " And 收费细目ID=" & lngItemID
            Call OpenRecordset(rsTemp, "判断医保项目是否允许做为特殊项目")
            If Not rsTemp.EOF Then
                If Not IsNull(rsTemp!附注) Then
                    If UBound(Split(rsTemp!附注, "|")) >= 2 Then
                        int可做为特殊项目 = Val(Split(rsTemp!附注, "|")(2))
                    End If
                End If
            End If
            
            GetItemInfo = ""
            If int可做为特殊项目 = 1 Then
                If MsgBox("该项目做为特殊项目上传至医保中心吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    GetItemInfo = "1"
                End If
            End If
        Case TYPE_重大校园卡
              '刘兴宏(200403)
            GetItemInfo = ""
        Case TYPE_重庆银海版
            GetItemInfo = ""
        Case TYPE_重庆渝北
            GetItemInfo = ""
        Case TYPE_兴安
            'GetItemInfo = ""
            If intType = 1 Then
                GetItemInfo = GetItemInfo_兴安(lngPatiID, lngItemID, str摘要, intType)
            End If
        Case TYPE_黔南
          '刘兴宏:20041022
            GetItemInfo = GetItemInfo_黔南(lngPatiID, lngItemID, str摘要, intType)
        Case TYPE_成都德阳
            '刘兴宏:20041103
            GetItemInfo = ""
        Case TYPE_成都内江
            '刘兴宏:20041207
            GetItemInfo = ""
        Case TYPE_广元旺苍
            '刘兴宏:20041210
            GetItemInfo = ""
        Case Is > 900
            '中联医保
            GetItemInfo = ""
    End Select
End Function

Public Sub ChooseDisease(ByVal lngPatiID As Long, ByVal lngPageID As Long)
'功能：在病人入院管理模块中，供选择病人的出院病种
'参数：lngPatiID-病人ID；lngPageID-主页ID
'返回：交易成功返回真
'------------------------------------------------------------------------------------------------------------------
'调用模块：病人入出管理
    
    Select Case gintInsure
        Case TYPE_重庆市
            Call 更新出院疾病_重庆(lngPatiID, lngPageID)
        Case TYPE_毕节
            Call 更新病种_毕节(lngPatiID, lngPageID)
        Case TYPE_宁海
            Call 住院信息变动_宁海(lngPatiID, lngPageID)
    End Select
End Sub

Public Sub IdentifyCancel(ByVal bytType As Byte, ByVal lng病人ID As Long)
    '说明：用于医保病人身份验证成功后，取消就诊的情况
    '例：就诊登记成功，但发现病人选择错误；已录入明细并完成虚拟结算，病人因某种原因不想继续进行结算操作
    Select Case gintInsure
    Case TYPE_渝北农医
        Call 取消就诊登记_渝北农医(bytType, lng病人ID)
    End Select
End Sub
