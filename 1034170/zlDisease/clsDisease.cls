VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Public Event PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)  '转科
Public Event Closed(ByVal lngFunID As Long, ByVal strTag As String)  'lngFunID，功能ID，0－表示 阳性结果反馈单填写了诊断后关闭窗体；strTag 扩展参数。
Private WithEvents frmDisRegist As frmDiseaseRegist
Attribute frmDisRegist.VB_VarHelpID = -1
Private frmFeedBackList As frmFeedBackList

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    If frmDisRegist Is Nothing Then Set frmDisRegist = New frmDiseaseRegist
    If frmFeedBackList Is Nothing Then Set frmFeedBackList = New frmFeedBackList
    If gobjComlib Is Nothing Then Set gobjComlib = CreateObject("zl9ComLib.clsComLib")
End Sub

Private Sub Class_Terminate()
     If Not frmDisRegist Is Nothing Then Set frmDisRegist = Nothing
     If Not frmFeedBackList Is Nothing Then Set frmFeedBackList = Nothing
End Sub

Public Sub InitDisease(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal lngModul As Long, ByVal strMainPrivs As String, Optional ByRef objMip As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    Call gobjComlib.InitCommon(cnMain)
    glngSys = lngSys
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    gstrPrivs = strMainPrivs
    If strMainPrivs = "" Then
        gstrPrivs = gobjComlib.GetPrivFunc(lngSys, lngModul)
    End If
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    End If
    '初始化操作员信息
    Call GetUserInfo
    '初始化系统参数:可分散到具体的GetForm中去初始化
    Call InitSysPar
    gstrDBOwer = GetDbOwner(lngSys)
    If Not (objMip Is Nothing) Then Set gclsMipModule = objMip
End Sub


Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
    On Error GoTo errHand

    Set gfrmMain = frmMain
    
    Set gcnOracle = cnMain
    Call gobjComlib.InitCommon(cnMain)
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    gstrDBOwer = GetDbOwner(lngSys)
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub

    gstrUnitName = GetUnitName
    gblnShowInTaskBar = blnShowInTaskBar
    
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    End If
    
    Call InitSysPar
    On Error Resume Next
    Select Case lngModul
        Case 1278 '传染病工作站
            frmDiseaseStation.Show , frmMain
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    
    Set gfrmMain = Nothing
    
    
    Set gcnOracle = cnMain
    Call gobjComlib.InitCommon(cnMain)
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    gstrDBOwer = GetDbOwner(lngSys)
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub

    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub

    gstrUnitName = GetUnitName
    gblnShowInTaskBar = True
    
    If gObjRichEPR Is Nothing Then
        Set gObjRichEPR = New zlRichEPR.cRichEPR
        Call gObjRichEPR.InitRichEPR(gcnOracle, Nothing, glngSys, False)
    End If
    
    Call InitSysPar
    On Error Resume Next
    Select Case lngModul
        Case 1278 '传染病工作站
            gobjComlib.ZLCommFun.ShowChildWindow frmDiseaseStation.hwnd, lngMain, True
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    blnChildren = gObjRichEPR.CloseWindows()
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function

Private Sub frmDisRegist_Closed(ByVal lngFunID As Long, ByVal strTag As String)
    RaiseEvent Closed(lngFunID, strTag)
End Sub

Private Sub frmDisRegist_PatiTransfer(ByVal lng病人ID As Long, ByVal str挂号No As String)
     RaiseEvent PatiTransfer(lng病人ID, str挂号No)
End Sub

Public Sub HideFrm(ByVal lngFrmID As Long)
'功能：隐藏窗体
'参数：lngFrmID,窗体标识，0－表示 frmDisRegist
    On Error Resume Next
    frmDisRegist.Hide
    Err.Clear
End Sub

Public Function ShowDisRegist(ByRef frmParent As Object, ByVal intType As Integer, Optional ByVal lngID As Long, _
                Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal str挂号No As String, _
                Optional ByVal lng医嘱id As Long, Optional ByVal var登记科室 As Variant, Optional ByVal dat送检时间 As Date, Optional ByVal var送检科室 As Variant, _
                Optional ByVal str送检医生 As String, Optional ByVal str标本名称 As String, Optional ByVal str反馈结果 As String, _
                Optional ByVal dat检查时间 As Date, Optional ByVal str疑似传染病 As String, Optional ByRef objMip As Object) As Integer
'功能：调用传染病阳性结果反馈单
'参数： intType 0表示填写（只显示上半部分），1-表示医生处理（只有下半部分可编辑），2-查看（所有不可编辑，可查看下半部分）
'       lngID  = 疾病阳性记录 ID
'       lng病人ID = 病人ID
'       lng主页ID=住院:主页ID
'       str挂号No =门诊：挂号单NO
'       lng医嘱ID =要填写反馈单的医嘱的医嘱ID
'       var登记科室 = 当前填写反馈单的科室
'       dat送检时间 = 标本到达的时间
'       var送检科室 = 让病人前来检查的科室
'       str送检医生 = 让病人前来检查的医生
'       str标本名称 = 病人检测所用的标本
'       str反馈结果 = 当前填写的反馈结果
'       dat检查时间 = 病人的检查时间
'       str疑似传染病  病人的疑似传染病
'       objMip  消息对象
'      如果intType为0的话，lng病人ID,lng主页ID,str挂号No必填
'      如果intType为1或者2的话，lngID必填
'返回值：1：只在intType 为0表示填写时才会出现，表示已经填写好反馈单发送成功；
'        2：只在intType 为1表示医生处理时才会出现，表示已经填写处理说明，反馈单已经处理完成；
'        3：只在intType 为1表示医生处理时才会出现，表示不处理反馈单，进行了转科处理；
'        其他情况返回值为0

    ShowDisRegist = frmDisRegist.ShowDiseaseRegist(frmParent, intType, lngID, lng病人ID, lng主页ID, str挂号No, lng医嘱id, var登记科室, dat送检时间, var送检科室, str送检医生, str标本名称, str反馈结果, dat检查时间, str疑似传染病, objMip)
End Function

Public Function ShowDisQuery(ByVal lng登记科室ID As Variant) As Long
'功能：调用传染病阳性结果查询界面
'参数：lng登记科室ID=疾病阳性记录 的 登记科室ID
'返回值：返回查询到的记录数量
    ShowDisQuery = frmDiseaseQuery.ShowDiseaseQuery(lng登记科室ID)
End Function


Public Function ShowPatiDis(ByVal rsTemp As ADODB.Recordset, ByRef frmParent As Object) As Long
'功能：调用传染病阳性结果查询界面
'参数：rsTemp-病人的阳性反馈单记录集
'返回值：返回选中的反馈单ID
    ShowPatiDis = frmDiseaseQuery.ShowPatiDis(rsTemp, frmParent)
End Function

Public Function GetFeedbackList() As Object
'功能：获取窗体
    Set GetFeedbackList = frmFeedBackList
End Function

Public Function EditNotFillReason(ByVal frmPar As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngPatFrom As Long) As Boolean
'功能：编辑保存不填写报告卡的原因
    Dim rsDisease As ADODB.Recordset
    Dim strTmp As String
    Dim strID As String
    If CheckDisNum(lngPatiID, lngPageId, lngPatFrom, rsDisease) Then
        If rsDisease.RecordCount = 1 Then
            strTmp = frmReason.ShowMe(frmPar, NVL(rsDisease!处理情况说明))
            Call SaveReason(strTmp, rsDisease!ID, rsDisease!记录状态)
        ElseIf rsDisease.RecordCount > 1 Then
            If frmFeedBackList.ShowMe(frmPar, rsDisease, strID, 2) Then
            rsDisease.Filter = "ID = " & strID
            If rsDisease.RecordCount > 0 Then
                strTmp = frmReason.ShowMe(frmPar, NVL(rsDisease!处理情况说明))
                Call SaveReason(strTmp, rsDisease!ID, rsDisease!记录状态)
            End If
            End If
        End If
    End If
End Function

