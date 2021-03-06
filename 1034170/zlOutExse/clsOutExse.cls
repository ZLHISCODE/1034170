VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = frmMain: glngMain = 0
    If Not CheckValid Then Exit Sub

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "零售药店管理"
        gstrCustomerAppellation = "客户"
    Else
        gstrCustomerAppellation = "病人"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '本地参数在模块的管理窗体激活时会再次读取(这里先读是因为Form_Load时可能会用到),因为模块之间窗口切换会改变全局变量的值
    Call InitLocPar(lngModul)
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(frmMain, lngModul)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 '药房划价管理
            frmManagePrice.Show , frmMain
        Case 1121 '病人收费管理
            If gfrmMain Is Nothing Then
                frmManageCharge.Show
            Else
                frmManageCharge.Show , frmMain
            End If
        Case 1122 '门诊记帐管理
            frmManageBilling.Show , frmMain
        Case 1124 '保险补充结算
            frmReplenishTheBalanceManage.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = Nothing: glngMain = lngMain
    If glngMain = 0 Then glngMain = -1
    
    If Not CheckValid Then Exit Sub

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    If glngSys Like "8??" Then
        App.Title = "零售药店管理"
        gstrCustomerAppellation = "客户"
    Else
        gstrCustomerAppellation = "病人"
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '本地参数在模块的管理窗体激活时会再次读取(这里先读是因为Form_Load时可能会用到),因为模块之间窗口切换会改变全局变量的值
    Call InitLocPar(lngModul)
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreateSquareCardObject(gfrmMain, lngModul)
    Call CreatePlugIn(lngModul)
    Call InitSysPar
    
    Select Case glngModul
        Case 1120 '药房划价管理
            zlCommFun.ShowChildWindow frmManagePrice.hWnd, lngMain
        Case 1121 '病人收费管理
            zlCommFun.ShowChildWindow frmManageCharge.hWnd, lngMain
        Case 1122 '门诊记帐管理
            zlCommFun.ShowChildWindow frmManageBilling.hWnd, lngMain
        Case 1124 '保险补充结算
            zlCommFun.ShowChildWindow frmReplenishTheBalanceManage.hWnd, lngMain
    End Select
End Sub

Public Function Charge(frmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String, ByVal bytInFun As Byte, ByVal bytInState As Byte) As Boolean
'功能：直接调用收费/划价功能
'返回：是否已划价
    '公共变量
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
    
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    
    If Not blnNoFirst Then
        blnNoFirst = True
        
        Set gfrmMain = frmMain: glngMain = 0
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        glngSys = lngSys
        gstrDBUser = strDBUser
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrUnitName = GetUnitName
        gstrPrivs = ";" & GetPrivFunc(lngSys, Choose(bytInFun + 1, 1121, 1120, 1122)) & ";"
        
        If InStr(gstrPrivs, ";划价;") = 0 Then
            MsgBox "你没有被授予使用划价功能的权限,如需使用,请与系统管理员联系!", vbInformation, gstrSysName
            Exit Function
        End If
                    
        '用户信息
        Call GetUserInfo
            
    End If
    '初始化参数
    Call InitLocPar(Choose(bytInFun + 1, 1121, 1120, 1122))
    Call InitSysPar
    Call CreateSquareCardObject(gfrmMain, Choose(bytInFun + 1, 1121, 1120, 1122))
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(Choose(bytInFun + 1, 1121, 1120, 1122))
    
    '调用划价
    Err.Clear
    On Error Resume Next
    frmCharge.mlngModul = Choose(bytInFun + 1, 1121, 1120, 1122)
    frmCharge.mstrPrivs = gstrPrivs
    frmCharge.mbytInFun = bytInFun
    frmCharge.mbytInState = bytInState
    If gfrmMain Is Nothing And glngMain = 0 Then
        frmCharge.Show
    Else
        frmCharge.Show 1, frmMain
    End If
    
    Charge = gblnOK
End Function
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Load frmPrint
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    CloseSquareCardObject
    Set grs医疗付款方式 = Nothing
    Unload frmPrint
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

Public Function zlCheckExcuteItemValied(frmMain As Object, cnOracle As ADODB.Connection, ByVal strDBUser As String, _
                                         ByVal lngSys As Long, ByVal lngModule As Long, ByVal lng病人ID As Long, _
                                         ByVal str收费类别 As String, ByVal strNos As String, ByVal str收费细目IDs As String) _
                                         As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:先诊疗后结算方式，检查执行项目的合法性
    '入参:frmMain-调用的主窗体
    '     cnOracle-传入Oracle连接
    '     strDBUser-数据库用户名
    '     lngSys-系统号
    '     lngModule-调用的模块号(主要是考虑后续可能存在根据模块来进行判断)
    '     lng病人ID-当前检查的病人ID
    '     str收费类别-本次执行的收费类别,可以传入多个,多个时用逗号分离:比如:4,5,6,7等,为空时，不根据收费类别进行检查
    '     strNOs-本次执行的单据号(记帐单号),可以传入多个，多个时用逗号分离,比如:4,5等，为空时，不根据单据号进行检查
    '     str收费细目IDs-表示本次执行的收费项目,可以为多个，用逗号分离,比如：123,456等，为空时，不按收费细目IDs进行检查
    '返回:True表示数据合法,允许执行；False-表示数据不合法，不允许执行该项目
    '编制:刘尔旋
    '日期:2013-12-24
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim rsCheck As New ADODB.Recordset, strSQL As String
    Dim strReturn As String, strArray() As String
    Set gcnOracle = cnOracle
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrDBUser = strDBUser: glngSys = lngSys: glngModul = lngModule
    strSQL = "Select zl_Get_ExcuteItem_Infor([1],[2],[3],[4],[5]) As 返回信息 From Dual"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "检查执行项目", lngModule, lng病人ID, str收费类别, strNos, str收费细目IDs)
    strReturn = Nvl(rsCheck!返回信息)
    If strReturn <> "" And InStr(strReturn, "|") > 0 Then
        strArray = Split(strReturn, "|")
        If Val(strArray(0)) = 2 Then
            MsgBox "数据不合法，禁止执行！" & vbCrLf & "提示信息:" & strArray(1), vbInformation, gstrSysName
            zlCheckExcuteItemValied = False
            Exit Function
        End If
        If Val(strArray(0)) = 1 Then
            If MsgBox("确定要执行项目？" & vbCrLf & "提示信息:" & strArray(1), vbYesNo + vbQuestion, gstrSysName) = vbYes Then
                zlCheckExcuteItemValied = True
                Exit Function
            Else
                zlCheckExcuteItemValied = False
                Exit Function
            End If
        End If
    End If
    If strReturn = "" Then
        zlCheckExcuteItemValied = True
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
