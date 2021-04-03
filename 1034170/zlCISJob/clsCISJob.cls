VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Public Type PatiInfo
    排序 As String
    状态 As Integer '病案主页.状态
    性质 As Integer '0-普通住院病人,1-门诊留观病人,2-住院留观病人
    住院号 As String
    床号 As String
    主页ID As Long
    病区ID As Long
    科室ID As Long
    产科 As Boolean
    入院日期 As Date
    出院日期 As Date
    路径状态 As Integer
    病案状态 As Integer
    数据转出 As Boolean
    险类 As Integer
    结清 As Boolean
End Type

Private mobjArchiveView As frmArchiveView
Private mobjOneArchiveView As frmArchiveView '显示某一个病人的电子病案

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
    Dim objPlugIn As Object
    On Error GoTo ErrHand

    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub
    On Error Resume Next
    Set gobjEmr = gfrmMain.mobjEMR  '新病历需要部件
    If err.Number <> 0 Then
        Set gobjEmr = Nothing
        err.Clear
    End If
    
    On Error GoTo ErrHand
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    If frmMain Is Nothing Then  '医保接口调试调用
        gblnShowInTaskBar = blnShowInTaskBar
    Else
        gblnShowInTaskBar = blnShowInTaskBar Or frmMain.frmHide
    End If
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, blnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs, gobjEmr)
    
    '判断临床工作站必须使用插件
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara("临床工作站必须使用zlPlugIn部件", 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "当前系统设置为必须使用医嘱插件(zlPlugIn)，本机缺少该插件，请联系管理员。", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If
    
    '调用菜单应用模块
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '门诊医生工作站
            Call InitCommunity
            frmOutDoctorStation.Show , frmMain
        Case 1261 '住院医生工作站
            frmInDoctorStation.ShowMe frmMain
        Case 1262 '住院护士工作站
            frmInNurseStation.Show , frmMain
        Case 1263 '医技工作站
            frmTechnicStation.Show , frmMain
        Case 1265 '新版住院护士工作站
            frmSublimeInNurseStation.Show , frmMain
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1078, 1275 '临床路径管理,临床路径跟踪
            Call gobjPath.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1266
            frmKSSExamine.Show , frmMain
        Case 1267
            frmOPSExamine.Show , frmMain
        Case 1268
            frmTransfuseExamine.ShowMe frmMain, 2
        Case 1269
            frmKssStatistics.Show , frmMain
        Case 1284 '危急值查询
            Call gobjKernel.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1910  '安装新血库系统后才有该模块
            If gbln血库系统 = False Then
                MsgBox "血库系统未正常安装或未勾选系统参数启用血库系统，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood = True Then
                Call gobjPublicBlood.zlShowBloodReceives(frmMain, 100, lngModul, 0)
            End If
            
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    Dim objPlugIn As Object
    
    Set gfrmMain = Nothing
    gblnShowInTaskBar = True
    'If Not CheckValid Then Exit Sub
    If Not objLogin Is Nothing Then
        Set gobjEmr = objLogin.mobjEMR  '新病历需要部件
        If err.Number <> 0 Then
            Set gobjEmr = Nothing
            err.Clear
        End If
    End If
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, gblnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    
    '判断临床工作站必须使用插件
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara("临床工作站必须使用zlPlugIn部件", 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "当前系统设置为必须使用医嘱插件(zlPlugIn)，本机缺少该插件，请联系管理员。", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If

    '调用菜单应用模块
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '门诊医生工作站
            Call InitCommunity
            zlCommFun.ShowChildWindow frmOutDoctorStation.hWnd, lngMain, True
        Case 1261 '住院医生工作站
            zlCommFun.ShowChildWindow frmInDoctorStation.hWnd, lngMain, True
        Case 1262 '住院护士工作站
            zlCommFun.ShowChildWindow frmInNurseStation.hWnd, lngMain, True
        Case 1263 '医技工作站
            zlCommFun.ShowChildWindow frmTechnicStation.hWnd, lngMain, True
        Case 1265 '新版住院护士工作站
            zlCommFun.ShowChildWindow frmSublimeInNurseStation.hWnd, lngMain, True
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1078, 1275 '临床路径管理,临床路径跟踪
            Call gobjPath.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1266
            zlCommFun.ShowChildWindow frmKSSExamine.hWnd, lngMain, True
        Case 1267
            zlCommFun.ShowChildWindow frmOPSExamine.hWnd, lngMain, True
        Case 1268
            zlCommFun.ShowChildWindow frmTransfuseExamine.hWnd, lngMain, True
        Case 1269
            zlCommFun.ShowChildWindow frmKssStatistics.hWnd, lngMain, True
        Case 1910 '安装血库系统后才会有该模块
            If gbln血库系统 = False Then
                MsgBox "血库系统未正常安装或未勾选系统参数启用血库系统，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood = True Then
                Call gobjPublicBlood.zlShowBloodReceives(lngMain, 100, lngModul, 0)
            End If
    End Select
End Sub

Private Sub InitCommunity()
'功能：入口初始化社区接口对象
    If gobjCommunity Is Nothing Then
        On Error Resume Next
        Set gobjCommunity = CreateObject("zlCommunity.clsCommunity")
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjCommunity Is Nothing Then
        If Not gobjCommunity.Initialize(gcnOracle) Then
            Set gobjCommunity = Nothing
        End If
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    Call gobjEmr.CloseForms
    blnChildren = gobjRichEPR.CloseWindows()
    blnChildren = blnChildren And gobjKernel.CloseWindows()
    blnChildren = blnChildren And gobjPath.CloseWindows()
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    If Not gobjCommunity Is Nothing Then
        Call gobjCommunity.Terminate
        Set gobjCommunity = Nothing
    End If
    If Not gobjPass Is Nothing Then
        Set gobjPass = Nothing
    End If
    If Not mobjOneArchiveView Is Nothing Then
        Unload mobjOneArchiveView
        Set mobjOneArchiveView = Nothing
    End If
    If Not mobjArchiveView Is Nothing Then
        Unload mobjArchiveView
        Set mobjArchiveView = Nothing
    End If
End Sub

Public Sub ShowArchive(ByRef objForm As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnModal As Boolean = False)
'功能：调用电子病案查阅功能，单个病人查阅
    If mobjOneArchiveView Is Nothing Then Set mobjOneArchiveView = New frmArchiveView
    Call mobjOneArchiveView.ShowArchive(objForm, lng病人ID, lng主页ID, blnModal)
End Sub

Public Function GetArchiveOneFrom() As Object
'功能：调用电子病案查阅功能，单个病人查阅的窗体
    Set GetArchiveOneFrom = mobjOneArchiveView
End Function

Public Function GetArchiveFrom() As Object
'功能：调用电子病案查阅功能，嵌入式获取窗体对象
    Set mobjArchiveView = New frmArchiveView
    mobjArchiveView.BorderStyle = FormBorderStyleConstants.vbBSNone '设置为无边框
    mobjArchiveView.Caption = mobjArchiveView.Caption       '重点是这一句
    Set GetArchiveFrom = mobjArchiveView
End Function

Public Sub zlArchiveRefresh(ByVal lng病人ID As Long, ByVal lng就诊ID As Long)
'功能：刷新电子病案查阅界面
    Call mobjArchiveView.zlRefresh(lng病人ID, lng就诊ID)
End Sub

Public Sub ViewEPRReport(frmParent As Object, ByVal lng报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(frmParent, lng报告ID, CanPrint)
End Sub

Public Sub ViewPACSImage(frmParent As Object, ByVal lng医嘱ID As Long, ByVal blnMoved As Boolean)
'功能：PACS观片处理
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(lng医嘱ID, frmParent, blnMoved)
    End If
End Sub

Public Sub InitCISJob(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal strMainPrivs As String, Optional ByRef objEmr As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    
    glngSys = lngSys

    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    gstrPrivs = strMainPrivs
    If Not objEmr Is Nothing Then
        Set gobjEmr = objEmr
    End If
    '初始化操作员信息
    Call GetUserInfo
    
    '初始化系统参数:可分散到具体的GetForm中去初始化
    Call InitSysPar
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs, gobjEmr)
End Sub
