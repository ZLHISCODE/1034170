Attribute VB_Name = "mdlQueueManage"
Option Explicit

Private mobjQueueManage As Object    '呼叫接口部件 zlQueueManage.clsQueueManage
Private mobjLCDShow As Object        '显示接口部件  zl9LCDShow.clsLCDShow

Public Sub QueueInit()
    '初始化呼叫对象
    Dim strName(1) As String
    Dim strPrivs As String
    
    '排队叫号权限
    strPrivs = GetPrivFunc(glngSys, 1160)
    If Trim(strPrivs) = "" Then
        Exit Sub
    End If
    
    strName(1) = "ABC"
    On Error GoTo hErr
    Set mobjQueueManage = CreateObject("zlQueueManage.clsQueueManage")
    
    If mobjQueueManage Is Nothing Then
        Exit Sub
    Else
        Call mobjQueueManage.zlInitVar(gcnOracle, glngSys, 3, 0)
    End If
    
    If zlDatabase.GetPara("显示排队队列", glngSys, 1160, "1") = "1" Then
        Set mobjLCDShow = CreateObject("zl9LCDShow.clsLCDShow")
        If Not mobjLCDShow Is Nothing Then
            Call mobjLCDShow.zlShow(gcnOracle, strName)     'LCDSHOW
        End If
    End If
    Exit Sub

hErr:
    Call QueueUnload
End Sub

Public Sub QueueUnload()
    If Not mobjQueueManage Is Nothing Then
        mobjQueueManage.CloseWindows
        Set mobjQueueManage = Nothing
    End If
    
    If Not mobjLCDShow Is Nothing Then
        mobjLCDShow.zlClose
        Set mobjLCDShow = Nothing
    End If
End Sub

Public Function QueueTimeCall() As Object
    On Error GoTo hErr
    If Not mobjQueueManage Is Nothing Then
        Set QueueTimeCall = mobjQueueManage.zlGetForm
    Else
        MsgBox "实例化排队叫号部件失败！", vbInformation, gstrSysName
    End If
    Exit Function
    
hErr:
    Set QueueTimeCall = Nothing
    If ErrCenter = 1 Then Resume
End Function

Public Sub QueueOnePlay(ByVal strNo As String, ByVal strPlayInfo As String)
'功能：呼叫指定病人

    If mobjQueueManage Is Nothing Then Exit Sub
    Call mobjQueueManage.zlQueueBroadcastCall(strPlayInfo)

End Sub

Public Sub QueueCall(ByVal strQueueName As String, ByVal lngDept As Long, _
                     ByVal objPati As cPatient)
'功能：顺呼
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lng挂号ID As Long
    
    '--- 顺呼
    On Error GoTo hErr

    If mobjQueueManage Is Nothing Then Exit Sub     '呼叫部件不对
    If objPati Is Nothing Then Exit Sub
    
'    'strSQL = "Select 挂号单1,挂号单2,序号 From 门诊穿刺台 Where 科室ID=[1] And (挂号单1=[2] Or 挂号单2=[2])"
'
'    '调整用挂号ID，统一用Long类型
'    strSQL = "Select a.序号, Decode(b.Id, Null, c.Id, b.Id) ID " & _
'             "From 门诊穿刺台 A, 病人挂号记录 B, 病人挂号记录 C " & _
'             "Where a.挂号单1 = b.No(+) And a.挂号单2 = c.No(+) " & _
'             "    And b.记录性质(+) = 1 And b.记录状态(+) = 1 " & _
'             "    And c.记录性质(+) = 1 And c.记录状态(+) = 1 " & _
'             "    And 科室ID = [1] And (a.挂号单1 = [2] Or a.挂号单2 = [2]) "
'
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "是否可呼叫", lngDept, strNo)
'    If Not rsTmp.EOF Then
'        lng挂号ID = zlcommfun.NVL(rsTmp!ID, 0)
'
'        '清除上一个
'        Call mobjQueueManage.zlDelQueue(strQueueName, lng挂号ID)
'        '插入排队
'        If mobjQueueManage.zlInQueue(strQueueName, 3, lng挂号ID, lngDept, strPatiName, lngPatiId, rsTmp!序号, strDocName) Then
'            '执行排除
'            Call mobjQueueManage.zlQueueExec(strQueueName, 3, lng挂号ID, 1)
'            SaveOperLog lngDept, strNo, CALLS, "显示并呼叫"
'        End If
'    End If

    '清除所有门诊输液排队叫号队列
    Call mobjQueueManage.zlDelQueue("输液类")
    '插入排队
    If mobjQueueManage.zlInQueue("输液类", 3, objPati.单据ID, lngDept, objPati.姓名, objPati.病人ID, "", "") Then
        '执行排队
        Call mobjQueueManage.zlQueueExec("输液类", 3, objPati.单据ID, 1)
        SaveOperLog lngDept, objPati, CALLS, "显示并呼叫"
    End If
    '刷新排队叫号LCD显示
    Call mobjQueueManage.zlRefresh(Split("|输液类", "|"), "输液类", objPati.单据ID)

    Exit Sub
hErr:
    SaveErrLog
End Sub

Public Sub QueueSetup(ByVal frmMe As Form)
    If Not mobjQueueManage Is Nothing Then
        Call mobjQueueManage.zlQueueParameterSetup(frmMe, glngSys)
    Else
        MsgBox "缺少10.30.40以上版本的呼叫接口部件（zlQueueManage）,请检查", vbQuestion, "门诊输液"
    End If
End Sub
