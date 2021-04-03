VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmRegist As frmRegist
Attribute mfrmRegist.VB_VarHelpID = -1

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    Call gclsInsure.InitOracle(gcnOracle)
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    Set mfrmRegist = New frmRegist
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlInitData(ByVal bytState As Byte, _
     Optional ByVal strDocs As String, Optional ByVal strDeptIDs As String, Optional ByVal strRooms As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的限制条件
    '入参:bytState-应用场合：目前支持 0-分诊台调用，1-门诊医生站调用，2-住院医生站调用
    '     strDocs-医生ID或者医生姓名,用逗号分隔，格式为 147,235,313…或者张三，李四，王五…，传入时显示传入医生的挂号安排
    '     strDeptIDs-科室ID,用逗号分隔，格式为 147,235,313…，传入时显示该科室的挂号安排
    '     strRooms-医生站调用时，可传入诊室，挂号时自动确定诊室。
    '返回:初始化成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    gbytState = bytState
    Select Case bytState
        Case 0
            glngModul = 1113
        Case 1
            glngModul = 1260
        Case 2
            glngModul = 1261
        Case 3
            glngModul = 1115
    End Select
    gstrDocs = strDocs
    gstrDeptIDs = strDeptIDs
    gstrRooms = strRooms
    gstrPrivs = ";" & gobjComlib.GetPrivFunc(glngSys, 9000) & ";"
    gbytRegistMode = Val(Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(0))
    If Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1) <> "" Then
        gdatRegistTime = CDate(Format(Split(gobjDatabase.GetPara("挂号排班模式", glngSys) & "|", "|")(1), "yyyy-mm-dd hh:mm:ss"))
    End If
    
    zlInitData = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlDefCommandBars(ByVal frmMain As Object, ByVal CommandBars As Object, _
                            Optional ByVal blnAddInTool As Boolean, Optional ByVal MenuControlBefore As Object, _
                            Optional ByVal ToolControlBefore As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建窗体菜单和工具条
    '入参:frmMain-调用的主窗体
    '     CommandBars-调用窗体的CommandBars对象
    '     blnAddInTool-是否将功能按钮中添加进工具栏
    '     MenuControlBefore-菜单栏中要添加功能对象的后一个Control对象，用以控制添加的功能对象的位置，不传时，根据默认规则指定按钮位置。
    '     ToolControlBefore-工具栏中要添加功能对象的后一个Control对象，用以控制添加的功能对象的位置，不传时，根据默认规则指定按钮位置。
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlDefCommandBars(frmMain, CommandBars, blnAddInTool, MenuControlBefore, ToolControlBefore)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置窗体的菜单和工具条的各项功能的Visible属性或Eanbeld属性
    '入参:Control-菜单控件
    '编制:刘尔旋
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal frmMain As Object, ByVal Control As Object, ByRef strOutNO As String, Optional ByVal lngPatiID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行挂号的相关功能
    '入参:frmMain-调用的主窗体
    '     Control-菜单控件
    '出参:strOutNo-成功保存的挂号单据号
    '编制:刘兴洪
    '日期:2015-04-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmRegist.zlExecuteCommandBars(frmMain, Control, strOutNO, lngPatiID)
End Sub

Public Function zlOpenStopedPlanBySN(ByVal frmMain As Object, ByVal lngModule As Long, _
    Optional ByVal lng记录ID As Long, _
    Optional ByVal lngDeptID As Long, Optional ByVal lngDoctorID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能：对启用了序号控制分时段的已停诊安排按序号开放挂号
    '入参：
    '   frmMain 调用的主窗体
    '   lngModule 调用模块号
    '   lng记录ID 记录ID,1114模块调用时传入
    '   lngDeptID 科室ID
    '   lngDoctorID 医生ID
    '返回：成功返回True，失败返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmOpenStopedPlanBySN As New frmOpenStopedPlanBySN
    zlOpenStopedPlanBySN = frmOpenStopedPlanBySN.ShowMe(frmMain, lngModule, lng记录ID, lngDeptID, lngDoctorID)
End Function

